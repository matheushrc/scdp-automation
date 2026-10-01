"""Validação e exportação dos campos visíveis de Relatórios > Viagem."""

from __future__ import annotations

import json
import os
import re
import tempfile
import unicodedata
from pathlib import Path
from typing import Annotated, TypedDict

from pydantic import BaseModel, BeforeValidator, ConfigDict, Field, TypeAdapter
from pydantic_core import PydanticCustomError


def snake_case(label: str) -> str:
    """Preserva o rótulo, removendo acentos e convertendo separadores em _."""
    ascii_label = (
        unicodedata.normalize("NFKD", label).encode("ascii", "ignore").decode()
    )
    return re.sub(r"[^a-z0-9]+", "_", ascii_label.lower()).strip("_")


def brazilian_number(value: object) -> object:
    """Converte números do relatório, sem transformar ausências em zero."""
    if isinstance(value, bool):
        raise PydanticCustomError("number_type", "Booleano não é um valor monetário.")
    if isinstance(value, str):
        value = value.strip().replace("\xa0", "")
        if not re.fullmatch(r"-?(?:\d+|\d{1,3}(?:\.\d{3})+),\d+", value):
            raise ValueError(f"Número brasileiro inválido: {value!r}")
        return float(value.replace(".", "").replace(",", "."))
    return value


Number = Annotated[float, BeforeValidator(brazilian_number), Field(allow_inf_nan=False)]


class Model(BaseModel):
    model_config = ConfigDict(
        extra="forbid", str_strip_whitespace=True, validate_assignment=True
    )


class CustoBilhetes(Model):
    passagens_e_taxas_iniciais_r: Number
    total_r: Number


class SubTotal(CustoBilhetes):
    quantidade_diarias: Number
    diarias_r: Number


class Trecho(SubTotal):
    inicio: str = Field(pattern=r"^\d{2}/\d{2}/\d{4}$")
    termino: str = Field(pattern=r"^\d{2}/\d{2}/\d{4}$")
    origem: str = Field(min_length=1)
    destino: str = Field(min_length=1)
    meio_de_transporte: str = Field(min_length=1)


PCDP_PATTERN = r"^\d{6}/\d{2}(?:-\d+[A-Z]+)?$"


class Viagem(Model):
    numero_da_solicitacao: str = Field(pattern=PCDP_PATTERN)
    nome_do_proposto: str = Field(min_length=1)
    orgao_solicitante: str = Field(min_length=1)
    orgao_superior: str = Field(min_length=1)
    tipo_da_viagem: str = Field(min_length=1)
    situacao_da_viagem: str = Field(min_length=1)
    motivo_viagem: str = Field(min_length=1)
    trechos: list[Trecho] = Field(min_length=1)
    custo_com_bilhetes_remarcados_nao_utilizados_cancelados_r: CustoBilhetes
    sub_total: SubTotal
    total_adicional_r: Number
    descontos_r: Number
    restituicao_r: Number
    reembolso_r: Number
    total_da_viagem_r: Number
    descricao_do_motivo_da_viagem: str | None = None


class Cell(TypedDict):
    text: str
    rowspan: int
    colspan: int


def header_names(rows: list[list[Cell]]) -> list[str]:
    grid: dict[tuple[int, int], str] = {}
    for r, cells in enumerate(rows):
        col = 0
        for cell in cells:
            while (r, col) in grid:
                col += 1
            for dy in range(cell["rowspan"]):
                for dx in range(cell["colspan"]):
                    grid[r + dy, col + dx] = snake_case(cell["text"])
            col += cell["colspan"]
    return [grid[len(rows) - 1, c] for c in range(16)]


def parse_report_rows(rows: list[list[Cell]]) -> list[Viagem]:
    """Lê cabeçalhos agrupados, trechos com rowspan e os rodapés de cada PCDP."""
    if len(rows) < 2:
        raise ValueError("Relatório sem cabeçalhos.")
    headers = header_names(rows[:2])
    result: list[Viagem] = []
    current: dict[str, object] = {}
    segments: list[dict[str, str]] = []
    for cells in rows[2:]:
        values = [c["text"].strip() for c in cells]
        if not any(values):
            continue
        first = values[0]
        if first in {"Sub-Total Geral", "Total (R$)"}:
            # Rodapé geral do relatório, não pertence à última solicitação.
            continue
        if re.fullmatch(PCDP_PATTERN, first):
            if current:
                result.append(Viagem.model_validate(current))
            if len(values) != 16:
                raise ValueError(
                    f"PCDP {first}: esperadas 16 colunas, recebidas {len(values)}."
                )
            segments = [dict(zip(headers[7:], values[7:], strict=True))]
            current = dict(zip(headers[:7], values[:7], strict=True))
            current["trechos"] = segments
        elif re.fullmatch(r"\d{2}/\d{2}/\d{4}", first) and current:
            segments.append(dict(zip(headers[7:], values, strict=True)))
        elif first.startswith("Custo com Bilhetes") and current:
            current[snake_case(first)] = dict(
                zip(headers[-2:], values[1:], strict=True)
            )
        elif first == "Sub-Total" and current:
            current["sub_total"] = dict(zip(headers[-4:], values[1:], strict=True))
        elif first == "Total Adicional (R$)" and current:
            if len(values) % 2:
                raise ValueError("Rodapé de totais incompleto.")
            current.update(
                {snake_case(values[i]): values[i + 1] for i in range(0, len(values), 2)}
            )
        else:
            raise ValueError(f"Linha inesperada no relatório: {first!r}")
    if current:
        result.append(Viagem.model_validate(current))
    return result


TRIPS = TypeAdapter(list[Viagem])


def load_trips(output: Path) -> list[Viagem]:
    """Valida o checkpoint inteiro antes de retomar ou sobrescrever arquivos."""
    if not output.exists():
        return []
    return TRIPS.validate_json(output.read_text(encoding="utf-8"))


def load_completed(output: Path) -> set[str]:
    return {
        v.numero_da_solicitacao
        for v in load_trips(output)
        if v.descricao_do_motivo_da_viagem is not None
    }


def save_json(output: Path, trips: list[Viagem]) -> None:
    """Grava o checkpoint JSON atomicamente após validar todos os registros."""
    validated = TRIPS.validate_python([trip.model_dump() for trip in trips])
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary: str | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            dir=output.parent,
            delete=False,
        ) as stream:
            temporary = stream.name
            stream.write(
                json.dumps(
                    [trip.model_dump(mode="json") for trip in validated],
                    ensure_ascii=False,
                    indent=2,
                    allow_nan=False,
                )
                + "\n"
            )
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, output)
    finally:
        if temporary and os.path.exists(temporary):
            os.unlink(temporary)
