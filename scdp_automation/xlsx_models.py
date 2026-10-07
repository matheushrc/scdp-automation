"""Validated aggregate rows from complete SCDP trips."""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from datetime import date

from scdp_automation.relatorio import Viagem


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
    description: str | None = None


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
                description=trip.descricao_do_motivo_da_viagem,
            )
        )

    return summaries
