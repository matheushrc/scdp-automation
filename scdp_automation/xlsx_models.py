"""Validated aggregate rows and stable debit categories."""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from datetime import date

from scdp_automation.relatorio import Viagem


@dataclass(frozen=True, slots=True)
class DebitCategory:
    """A stable debit key with its display name and budget segment."""

    code: str
    name: str
    segment: str
    review_required: bool = False


@dataclass(frozen=True, slots=True)
class TripSummary:
    """One row of aggregated values from a complete SCDP request."""

    pcdp: str
    proposed: str
    status: str
    daily_count: float
    daily_amount: float
    ticket_amount: float
    additional_amount: float
    discount_amount: float
    restitution_amount: float
    reimbursement_amount: float
    trip_total: float
    start_date: date | None = None
    end_date: date | None = None
    verified_date: date | None = None


DEBIT_CATEGORIES: tuple[DebitCategory, ...] = (
    DebitCategory("ADMINISTRAÇÃO", "Administração", "SEG 1 GRADUAÇÃO"),
    DebitCategory("AGRONOMIA", "Agronomia", "SEG 1 GRADUAÇÃO"),
    DebitCategory("C COMPUTAÇÃO", "Ciência da Computação", "SEG 1 GRADUAÇÃO"),
    DebitCategory("C ECONÔMICAS", "Ciências Econômicas", "SEG 1 GRADUAÇÃO"),
    DebitCategory("CIÊNCIAS SOCIAIS", "Ciências Sociais", "SEG 1 GRADUAÇÃO"),
    DebitCategory("ENFERMAGEM", "Enfermagem", "SEG 1 GRADUAÇÃO"),
    DebitCategory(
        "ENG AMBIENTAL", "Engenharia Ambiental e Sanitária", "SEG 1 GRADUAÇÃO"
    ),
    DebitCategory("ENGENHARIA CIVIL", "Engenharia Civil", "SEG 1 GRADUAÇÃO"),
    DebitCategory("FILOSOFIA", "Filosofia", "SEG 1 GRADUAÇÃO"),
    DebitCategory("GEOGRAFIA", "Geografia", "SEG 1 GRADUAÇÃO"),
    DebitCategory("HISTÓRIA", "História", "SEG 1 GRADUAÇÃO"),
    DebitCategory("LETRAS", "Letras – Português e Espanhol", "SEG 1 GRADUAÇÃO"),
    DebitCategory("MATEMÁTICA", "Matemática", "SEG 1 GRADUAÇÃO"),
    DebitCategory("MEDICINA", "Medicina", "SEG 1 GRADUAÇÃO"),
    DebitCategory("PEDAGOGIA", "Pedagogia", "SEG 1 GRADUAÇÃO"),
    DebitCategory(
        "Lato Oncologia", "Especialização em Enfermagem em Oncologia", "SEG 2 MESTRADO"
    ),
    DebitCategory("PPGCB", "Mestrado em Ciências Biomédicas", "SEG 2 MESTRADO"),
    DebitCategory("PPGE", "Mestrado em Educação", "SEG 2 MESTRADO"),
    DebitCategory(
        "PPGEL", "Estudos Linguísticos – Mestrado e Doutorado", "SEG 2 MESTRADO"
    ),
    DebitCategory("PPGEL +", "Estudos Linguísticos – Rateio", "SEG 2 MESTRADO"),
    DebitCategory("PPGEnf", "Mestrado em Enfermagem", "SEG 2 MESTRADO"),
    DebitCategory("PPGFil", "Mestrado em Filosofia", "SEG 2 MESTRADO"),
    DebitCategory("PPGGeo", "Mestrado em Geografia", "SEG 2 MESTRADO"),
    DebitCategory("PPGH/PPGDH", "Mestrado e Doutorado em História", "SEG 2 MESTRADO"),
    DebitCategory(
        "PROFIAP",
        "Mestrado Profissional em Administração Pública em Rede Nacional",
        "SEG 2 MESTRADO",
    ),
    DebitCategory(
        "PROFMAT",
        "Mestrado Profissional em Matemática em Rede Nacional",
        "SEG 2 MESTRADO",
    ),
    DebitCategory("DIREÇÃO", "Geral (Direção/Coordenações)", "SEG 3 OUTROS"),
    DebitCategory("DIREÇÃO - AGAS", "DIREÇÃO - AGAS", "SEG 3 OUTROS"),
    DebitCategory("DIREÇÃO - Banca Libras", "DIREÇÃO - Banca Libras", "SEG 3 OUTROS"),
    DebitCategory("DIREÇÃO - CAAEX", "DIREÇÃO - CAAEX", "SEG 3 OUTROS"),
    DebitCategory("DIREÇÃO - Empr Junior", "DIREÇÃO - Empr Junior", "SEG 3 OUTROS"),
    DebitCategory(
        "DIREÇÃO - StartUp Summit", "DIREÇÃO - StartUp Summit", "SEG 3 OUTROS"
    ),
    DebitCategory(
        "DIREÇÃO - StartUp Weekend", "DIREÇÃO - StartUp Weekend", "SEG 3 OUTROS"
    ),
    DebitCategory("DIREÇÃO - Sunset", "DIREÇÃO - Sunset", "SEG 3 OUTROS"),
    DebitCategory("CAPPG - Res 49", "CAPPG - Res 49", "SEG 4 AUX EVENTOS"),
    DebitCategory("AFAST PAÍS", "Afastamento no país", "SEG 5 AFAST PAÍS"),
)


def summarize_trips(trips: Sequence[Viagem]) -> list[TripSummary]:
    """Map one summary row per unique full PCDP, using trip-level subtotals."""
    summaries: list[TripSummary] = []
    seen_pcdps: set[str] = set()

    for trip in trips:
        pcdp = trip.numero_da_solicitacao
        if pcdp in seen_pcdps:
            raise ValueError("PCDP duplicada na listagem.")
        seen_pcdps.add(pcdp)

        summaries.append(
            TripSummary(
                pcdp=pcdp,
                proposed=trip.nome_do_proposto,
                status=trip.situacao_da_viagem,
                daily_count=trip.sub_total.quantidade_diarias,
                daily_amount=trip.sub_total.diarias_r,
                ticket_amount=trip.sub_total.passagens_e_taxas_iniciais_r,
                additional_amount=trip.total_adicional_r,
                discount_amount=trip.descontos_r,
                restitution_amount=trip.restituicao_r,
                reimbursement_amount=trip.reembolso_r,
                trip_total=trip.total_da_viagem_r,
                start_date=trip.data_inicio,
                end_date=trip.data_termino,
                verified_date=trip.data_da_ultima_verificacao,
            )
        )

    return summaries
