import tempfile
import unittest
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock, patch

from scdp_automation.chrome_profile_setup import ProfileConfig, write_profile_config
from scdp_automation.config import load_extraction_config, sync_extraction_config
from scdp_automation.extrator import ensure_annual_period, parse_args


class ExtractionConfigTests(unittest.TestCase):
    def test_rollover_updates_year_and_name_without_changing_profile(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / ".scdp-config.toml"
            profile = (
                '\n[chrome_profile]\nsource_path = "/tmp/chrome/Default"\nemail = ""\n'
            )
            path.write_text(
                'version = 1\nano = 2025\nnome_planilha = "gastos_scdp_2025.xlsx"\n'
                + profile
            )
            config = sync_extraction_config(path, current_year=2026)
            self.assertEqual(config.ano, 2026)
            self.assertEqual(config.nome_planilha, "gastos_scdp_2026.xlsx")
            self.assertEqual(config.checkpoint_name, "viagens_scdp_2026.json")
            self.assertIn(profile, path.read_text())
            before = path.read_bytes()
            sync_extraction_config(path, current_year=2026)
            self.assertEqual(path.read_bytes(), before)

    def test_name_is_fixed_and_obsolete_name_setting_is_removed(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / ".scdp-config.toml"
            path.write_text('ano = 2026\nnome_planilha = "gastos_cch.xlsx"\n')
            config = sync_extraction_config(path, current_year=2026)
            self.assertEqual(config.nome_planilha, "gastos_scdp_2026.xlsx")
            self.assertNotIn("nome_planilha", path.read_text())

    def test_cli_uses_synchronized_year_and_workbook_name(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / ".scdp-config.toml"
            path.write_text('ano = 2025\nnome_planilha = "gastos_scdp_2025.xlsx"\n')
            config = sync_extraction_config(path, current_year=2026)
            with patch(
                "scdp_automation.extrator.sync_extraction_config", return_value=config
            ):
                args = parse_args([])
            self.assertEqual(args.ano, 2026)
            self.assertEqual(args.output.name, "viagens_scdp_2026.json")
            self.assertEqual(args.workbook.name, "gastos_scdp_2026.xlsx")

    def test_year_controls_checkpoint_and_default_workbook_name(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / ".scdp-config.toml").write_text("ano = 2025\n")
            config = load_extraction_config(root / ".scdp-config.toml")
            self.assertEqual(config.ano, 2025)
            self.assertEqual(config.checkpoint_name, "viagens_scdp_2025.json")
            self.assertEqual(config.nome_planilha, "gastos_scdp_2025.xlsx")

    def test_profile_write_preserves_extraction_settings(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / ".scdp-config.toml"
            path.write_text('ano = 2025\nnome_planilha = "gastos_cch.xlsx"\n')
            write_profile_config(path, ProfileConfig("/tmp/chrome/Default", ""))
            config = load_extraction_config(path)
            self.assertEqual(config.ano, 2025)
            self.assertEqual(config.nome_planilha, "gastos_scdp_2025.xlsx")
            self.assertNotIn("nome_planilha", path.read_text())

    def test_invalid_settings_are_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / ".scdp-config.toml"
            for content in (
                "ano = true",
                'ano = "2025"',
                "ano = 25",
            ):
                with self.subTest(content=content):
                    path.write_text(content)
                    with self.assertRaises(ValueError):
                        load_extraction_config(path)


class AnnualPeriodTests(unittest.IsolatedAsyncioTestCase):
    async def test_old_period_is_changed_to_the_current_year(self):
        page = MagicMock()
        annual, start = AsyncMock(), AsyncMock()
        annual.is_checked.return_value = True
        page.locator.side_effect = [annual, start]
        page.evaluate = AsyncMock(return_value=False)
        page.wait_for_function = AsyncMock()

        await ensure_annual_period(page, 2026)

        annual.uncheck.assert_awaited_once_with(force=True)
        annual.check.assert_awaited_once_with(force=True)
        self.assertEqual(page.wait_for_function.await_count, 2)
        self.assertEqual(page.wait_for_function.await_args_list[-1].kwargs["arg"], 2026)
