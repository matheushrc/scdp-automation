"""Read the local extraction year and workbook output name."""

from __future__ import annotations

import os
import re
import tempfile
import tomllib
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from zoneinfo import ZoneInfo

REPO_ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = REPO_ROOT / ".scdp-config.toml"


@dataclass(frozen=True, slots=True)
class ExtractionConfig:
    ano: int

    @property
    def nome_planilha(self) -> str:
        return f"gastos_scdp_{self.ano}.xlsx"

    @property
    def checkpoint_name(self) -> str:
        return f"viagens_scdp_{self.ano}.json"


def current_date() -> date:
    return datetime.now(ZoneInfo("America/Sao_Paulo")).date()


def current_year() -> int:
    return current_date().year


def load_extraction_config(path: Path = CONFIG_PATH) -> ExtractionConfig:
    data = {}
    if path.exists():
        try:
            with path.open("rb") as stream:
                data = tomllib.load(stream)
        except (OSError, tomllib.TOMLDecodeError) as error:
            raise ValueError(
                f"Não foi possível ler a configuração em {path}."
            ) from error
    year = data.get("ano", current_year())
    if isinstance(year, bool) or not isinstance(year, int) or not 1000 <= year <= 9999:
        raise ValueError("ano deve ser um número inteiro de quatro dígitos.")
    return ExtractionConfig(year)


def sync_extraction_config(
    path: Path = CONFIG_PATH, *, current_year: int | None = None
) -> ExtractionConfig:
    """Record the current year while preserving the selected Chrome profile."""
    year = current_year or current_date().year
    load_extraction_config(path)
    content = path.read_text(encoding="utf-8") if path.exists() else "version = 1\n"
    config = ExtractionConfig(year)
    table = re.search(r"(?m)^\s*\[", content)
    boundary = table.start() if table else len(content)
    header, tables = content[:boundary], content[boundary:]
    header = re.sub(r"(?m)^nome_planilha\s*=.*\n?", "", header)
    pattern = r"(?m)^ano\s*=.*$"
    replacement = f"ano = {year}"
    if re.search(pattern, header):
        header = re.sub(pattern, lambda _: replacement, header)
    else:
        header = header.rstrip() + "\n" + replacement + "\n\n"
    updated = header + tables
    if updated != content or not path.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
        temporary_path: Path | None = None
        try:
            with tempfile.NamedTemporaryFile(
                mode="w", encoding="utf-8", dir=path.parent, delete=False
            ) as stream:
                temporary_path = Path(stream.name)
                stream.write(updated)
                stream.flush()
                os.fsync(stream.fileno())
            if os.name != "nt":
                temporary_path.chmod(0o600)
            os.replace(temporary_path, path)
        finally:
            if temporary_path is not None:
                temporary_path.unlink(missing_ok=True)
    return config
