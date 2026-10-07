import json
import tempfile
import unittest
from pathlib import Path

from pydantic import ValidationError

from scdp_automation import relatorio


def sample_rows() -> list[list[relatorio.Cell]]:
    headers = [
        "Número da Solicitação",
        "Nome do Proposto",
        "Órgão Solicitante",
        "Órgão Superior",
        "Tipo da Viagem",
        "Situação da Viagem",
        "Motivo Viagem",
    ]
    money = [
        "Quantidade Diárias",
        "Diárias (R$)",
        "Passagens e Taxas Iniciais (R$)",
        "Total (R$)",
    ]

    def cells(values: list[str], rowspan: int = 1) -> list[relatorio.Cell]:
        return [{"text": v, "rowspan": rowspan, "colspan": 1} for v in values]

    return [
        cells(headers, 2)
        + [
            {"text": "Período", "rowspan": 1, "colspan": 2},
            {"text": "Trecho", "rowspan": 1, "colspan": 3},
        ]
        + cells(money, 2),
        cells(["Início", "Término", "Origem", "Destino", "Meio de Transporte"]),
        cells(
            [
                "999999/26",
                "Pessoa Exemplo",
                "ORG-TESTE",
                "ORG-SUPERIOR-TESTE",
                "NACIONAL",
                "Cancelada",
                "Nacional - A Serviço",
            ],
            2,
        )
        + cells(
            [
                "29/01/2026",
                "01/02/2026",
                "Cidade Alfa (AA)",
                "Cidade Beta (BB)",
                "Terrestre",
                "2,0",
                "1.234,56",
                "100,00",
                "1.334,56",
            ]
        ),
        cells(
            [
                "01/02/2026",
                "01/02/2026",
                "Cidade Beta (BB)",
                "Retorno para Cidade Alfa (AA)",
                "Terrestre",
                "0,0",
                "0,00",
                "0,00",
                "0,00",
            ]
        ),
        [
            {
                "text": "Custo com Bilhetes Remarcados/Não Utilizados/Cancelados (R$)",
                "rowspan": 1,
                "colspan": 14,
            }
        ]
        + cells(["20,00", "20,00"]),
        [{"text": "Sub-Total", "rowspan": 1, "colspan": 12}]
        + cells(["2,0", "1.234,56", "120,00", "1.354,56"]),
        cells(
            [
                "Total Adicional (R$)",
                "95,00",
                "Descontos (R$)",
                "10,00",
                "Restituição (R$)",
                "0,00",
                "Reembolso (R$)",
                "0,00",
                "Total da Viagem (R$)",
                "1.439,56",
            ]
        ),
    ]


class ReportTests(unittest.TestCase):
    def test_complementary_pcdp_keeps_its_suffix(self) -> None:
        rows = sample_rows()
        rows[2][0]["text"] = "999932/26-1C"
        self.assertEqual(
            relatorio.parse_report_rows(rows)[0].numero_da_solicitacao,
            "999932/26-1C",
        )

    def test_general_footer_does_not_overwrite_last_trip(self) -> None:
        rows = sample_rows()
        rows.extend(
            [
                [{"text": v, "rowspan": 1, "colspan": 1} for v in values]
                for values in [
                    ["Sub-Total Geral", "2,0", "1.234,56", "120,00", "1.354,56"],
                    ["Total (R$)", "1.439,56"],
                ]
            ]
        )
        trips = relatorio.parse_report_rows(rows)
        self.assertEqual(len(trips), 1)
        self.assertEqual(trips[0].total_da_viagem_r, 1439.56)

    def test_merged_headers_and_rows_preserve_segments_and_totals(self) -> None:
        trip = relatorio.parse_report_rows(sample_rows())[0]
        self.assertEqual(trip.numero_da_solicitacao, "999999/26")
        self.assertEqual(trip.nome_do_proposto, "Pessoa Exemplo")
        self.assertEqual(len(trip.trechos), 2)
        self.assertEqual(trip.trechos[1].destino, "Retorno para Cidade Alfa (AA)")
        self.assertEqual(trip.trechos[0].diarias_r, 1234.56)
        self.assertEqual(trip.sub_total.total_r, 1354.56)
        self.assertEqual(
            trip.custo_com_bilhetes_remarcados_nao_utilizados_cancelados_r.total_r, 20.0
        )
        self.assertEqual(trip.total_da_viagem_r, 1439.56)

    def test_invalid_or_missing_money_is_not_silently_zero(self) -> None:
        for value in ["", "---", "NaN", "1.23,45"]:
            rows = sample_rows()
            rows[2][-3]["text"] = value
            with self.subTest(value=value), self.assertRaises(ValidationError):
                relatorio.parse_report_rows(rows)

    def test_missing_subtotal_is_rejected(self) -> None:
        rows = sample_rows()
        del rows[-2]
        with self.assertRaises(ValidationError):
            relatorio.parse_report_rows(rows)

    def test_export_and_resume_require_complete_description(self) -> None:
        trip = relatorio.parse_report_rows(sample_rows())[0]
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "viagens.json"
            relatorio.save_json(path, [trip])
            trip.descricao_do_motivo_da_viagem = "Primeira linha\nSegunda linha"
            relatorio.save_json(path, [trip])
            data = json.loads(path.read_text())
            self.assertIsInstance(data[0]["trechos"][0]["diarias_r"], float)
            self.assertEqual(
                data[0]["descricao_do_motivo_da_viagem"],
                "Primeira linha\nSegunda linha",
            )
            self.assertFalse(path.with_suffix(".csv").exists())
            self.assertEqual(list(Path(tmp).iterdir()), [path])

    def test_checkpoint_rejects_boolean_and_nonfinite_amounts(self) -> None:
        data = relatorio.parse_report_rows(sample_rows())[0].model_dump()
        for invalid in (True, float("inf"), float("nan")):
            with self.subTest(value=invalid), self.assertRaises(ValidationError):
                relatorio.Viagem.model_validate({**data, "total_da_viagem_r": invalid})

    def test_json_checkpoint_remains_intact_when_atomic_replace_fails(self) -> None:
        from unittest.mock import patch

        trip = relatorio.parse_report_rows(sample_rows())[0]
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "viagens.json"
            path.write_text("previous checkpoint", encoding="utf-8")
            with (
                patch("scdp_automation.relatorio.os.replace", side_effect=OSError),
                self.assertRaises(OSError),
            ):
                relatorio.save_json(path, [trip])
            self.assertEqual(path.read_text(encoding="utf-8"), "previous checkpoint")
            self.assertEqual(list(Path(tmp).iterdir()), [path])

    def test_corrupt_checkpoint_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "viagens.json"
            path.write_text('[{"numero_da_solicitacao":"999999/26"}]')
            with self.assertRaises(ValidationError):
                relatorio.load_trips(path)
