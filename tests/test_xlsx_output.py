import unittest
from dataclasses import FrozenInstanceError

from scdp_automation.relatorio import Viagem
from scdp_automation.xlsx_output import (
    DEBIT_CATEGORIES,
    DebitCategory,
    TripSummary,
    summarize_trips,
)


def make_trip(pcdp: str = "123456/26-2B") -> Viagem:
    """Create a validated trip whose itinerary differs from its totals."""
    return Viagem.model_validate(
        {
            "numero_da_solicitacao": pcdp,
            "nome_do_proposto": "Pessoa Exemplo",
            "orgao_solicitante": "ORG-TESTE",
            "orgao_superior": "ORG-SUPERIOR-TESTE",
            "tipo_da_viagem": "NACIONAL",
            "situacao_da_viagem": "Autorizada",
            "motivo_viagem": "Nacional - A Serviço",
            "trechos": [
                {
                    "inicio": "01/03/2026",
                    "termino": "03/03/2026",
                    "origem": "Cidade Alfa (AA)",
                    "destino": "Cidade Beta (BB)",
                    "meio_de_transporte": "Aéreo",
                    "quantidade_diarias": 2.0,
                    "diarias_r": 700.25,
                    "passagens_e_taxas_iniciais_r": 800.50,
                    "total_r": 1500.75,
                },
                {
                    "inicio": "03/03/2026",
                    "termino": "05/03/2026",
                    "origem": "Cidade Beta (BB)",
                    "destino": "Cidade Alfa (AA)",
                    "meio_de_transporte": "Aéreo",
                    "quantidade_diarias": 3.0,
                    "diarias_r": 900.75,
                    "passagens_e_taxas_iniciais_r": 1000.25,
                    "total_r": 1901.00,
                },
            ],
            "custo_com_bilhetes_remarcados_nao_utilizados_cancelados_r": {
                "passagens_e_taxas_iniciais_r": 2.0,
                "total_r": 2.0,
            },
            "sub_total": {
                "quantidade_diarias": 8.5,
                "diarias_r": 1234.56,
                "passagens_e_taxas_iniciais_r": 789.01,
                "total_r": 2023.57,
            },
            "total_adicional_r": 33.05,
            "descontos_r": 4.01,
            "restituicao_r": 5.02,
            "reembolso_r": 6.03,
            "total_da_viagem_r": 2063.66,
        }
    )


class TripSummaryTests(unittest.TestCase):
    def test_summarize_trip_uses_trip_subtotal_once(self) -> None:
        trip = make_trip()

        summary = summarize_trips([trip])[0]

        self.assertIsInstance(summary, TripSummary)
        self.assertEqual(
            summary,
            TripSummary(
                pcdp="123456/26-2B",
                proposed="Pessoa Exemplo",
                status="Autorizada",
                daily_count=8.5,
                daily_amount=1234.56,
                ticket_amount=789.01,
                additional_amount=33.05,
                discount_amount=4.01,
                restitution_amount=5.02,
                reimbursement_amount=6.03,
                trip_total=2063.66,
            ),
        )

    def test_summarize_trips_rejects_duplicate_full_pcdp(self) -> None:
        with self.assertRaisesRegex(ValueError, "123456/26-2B"):
            summarize_trips([make_trip(), make_trip()])

        summaries = summarize_trips([make_trip("123456/26"), make_trip("123456/26-2B")])

        self.assertEqual(
            [summary.pcdp for summary in summaries], ["123456/26", "123456/26-2B"]
        )
        complementary_summaries = summarize_trips(
            [make_trip("123456/26-1C"), make_trip("123456/26-2C")]
        )
        self.assertEqual(
            [summary.pcdp for summary in complementary_summaries],
            ["123456/26-1C", "123456/26-2C"],
        )


class DebitCategoryTests(unittest.TestCase):
    def test_debit_catalog_covers_active_and_future_categories(self) -> None:
        self.assertTrue(DEBIT_CATEGORIES)
        self.assertTrue(
            all(isinstance(item, DebitCategory) for item in DEBIT_CATEGORIES)
        )
        self.assertTrue(
            all(item.__dataclass_params__.frozen for item in DEBIT_CATEGORIES)
        )
        self.assertEqual(
            {item.segment for item in DEBIT_CATEGORIES},
            {
                "SEG 1 GRADUAÇÃO",
                "SEG 2 MESTRADO",
                "SEG 3 OUTROS",
                "SEG 4 AUX EVENTOS",
                "SEG 5 AFAST PAÍS",
            },
        )

        codes = [item.code for item in DEBIT_CATEGORIES]
        self.assertEqual(len(codes), len(set(codes)))
        self.assertEqual(
            set(codes),
            {
                "ADMINISTRAÇÃO",
                "AGRONOMIA",
                "C COMPUTAÇÃO",
                "C ECONÔMICAS",
                "CIÊNCIAS SOCIAIS",
                "ENFERMAGEM",
                "ENG AMBIENTAL",
                "ENGENHARIA CIVIL",
                "FILOSOFIA",
                "GEOGRAFIA",
                "HISTÓRIA",
                "LETRAS",
                "MATEMÁTICA",
                "MEDICINA",
                "PEDAGOGIA",
                "Lato Oncologia",
                "PPGCB",
                "PPGE",
                "PPGEL",
                "PPGEL +",
                "PPGEnf",
                "PPGFil",
                "PPGGeo",
                "PPGH",
                "PPGDH",
                "PROFIAP",
                "PROFMAT",
                "DIREÇÃO",
                "DIREÇÃO - AGAS",
                "DIREÇÃO - Banca Libras",
                "DIREÇÃO - CAAEX",
                "DIREÇÃO - Empr Junior",
                "DIREÇÃO - StartUp Summit",
                "DIREÇÃO - StartUp Weekend",
                "DIREÇÃO - Sunset",
                "CAPPG - Res 49",
                "AFAST PAÍS",
            },
        )
        with self.assertRaises(FrozenInstanceError):
            DEBIT_CATEGORIES[0].__setattr__("name", "alterado")

        category_by_code = {item.code: item for item in DEBIT_CATEGORIES}
        self.assertEqual(category_by_code["AGRONOMIA"].name, "Agronomia")
        self.assertEqual(category_by_code["Lato Oncologia"].name, "LS Enf em Oncologia")
        self.assertEqual(category_by_code["PPGEL +"].name, "PPGEL +")
        self.assertEqual(
            category_by_code["DIREÇÃO"].name, "Geral (Direção/Coordenações)"
        )
        for label in (
            "DIREÇÃO - AGAS",
            "DIREÇÃO - Banca Libras",
            "DIREÇÃO - CAAEX",
            "DIREÇÃO - Empr Junior",
            "DIREÇÃO - StartUp Summit",
            "DIREÇÃO - StartUp Weekend",
            "DIREÇÃO - Sunset",
        ):
            self.assertEqual(category_by_code[label].name, label)

        for code in (
            "ADMINISTRAÇÃO",
            "C COMPUTAÇÃO",
            "C ECONÔMICAS",
            "LETRAS",
            "PPGFil",
            "PPGDH",
            "PROFMAT",
        ):
            self.assertIn(code, category_by_code)

        ppghd = category_by_code["PPGDH"]
        self.assertTrue(ppghd.review_required)
        self.assertEqual(ppghd.name, "PPGDH")


if __name__ == "__main__":
    unittest.main()
