import unittest
from pathlib import Path

from scdp_automation.extrator import (
    DEFAULT_OUTPUT,
    extract_description,
    parse_args,
)


class ExtratorTests(unittest.TestCase):
    def test_extract_description_from_label_on_separate_line(self) -> None:
        text = "Solicitação\nDescrição do Motivo da Viagem\nParticipação em reunião"
        self.assertEqual(extract_description(text), "Participação em reunião")

    def test_output_path_is_fixed_to_json_checkpoint(self) -> None:
        from unittest.mock import patch

        with patch("sys.argv", ["scdp-extrair"]):
            args = parse_args()

        self.assertEqual(DEFAULT_OUTPUT, Path("output/viagens_scdp_2026.json"))
        self.assertEqual(args.output, DEFAULT_OUTPUT)
        self.assertEqual(args.limite, 0)


class LoginRoutingTests(unittest.IsolatedAsyncioTestCase):
    async def test_existing_govbr_login_does_not_search_for_scdp_button(self) -> None:
        from unittest.mock import AsyncMock, MagicMock, patch

        from playwright.async_api import Page

        from scdp_automation.extrator import start_login_if_needed

        page = MagicMock(spec=Page)
        page.url = "https://sso.acesso.gov.br/login"
        page.locator.side_effect = AssertionError("Botão SCDP não existe no gov.br")
        with (
            patch("scdp_automation.extrator.load_credentials"),
            patch(
                "scdp_automation.extrator.authenticate_gov_br", new=AsyncMock()
            ) as authenticate,
        ):
            await start_login_if_needed(page)
        authenticate.assert_awaited_once()

    async def test_authenticated_menu_can_finish_loading_before_login_check(
        self,
    ) -> None:
        from unittest.mock import AsyncMock, MagicMock

        from playwright.async_api import Page

        from scdp_automation.extrator import start_login_if_needed

        menu = MagicMock()
        menu.wait_for = AsyncMock()
        page = MagicMock(spec=Page)
        page.url = "https://www2.scdp.gov.br/novoscdp/home.xhtml"
        page.get_by_role.return_value = menu
        await start_login_if_needed(page)
        menu.wait_for.assert_awaited_once_with(state="visible", timeout=2_000)
        page.locator.assert_not_called()


class QueryRetryTests(unittest.IsolatedAsyncioTestCase):
    async def test_timeout_retries_read_only_query(self) -> None:
        from unittest.mock import MagicMock, patch

        from playwright.async_api import Page, TimeoutError

        from scdp_automation.extrator import consult_trip_reason

        with patch(
            "scdp_automation.extrator._consult_trip_reason",
            side_effect=[TimeoutError("timeout"), "Motivo completo"],
        ):
            reason = await consult_trip_reason(MagicMock(spec=Page), "000548/26")
        self.assertEqual(reason, "Motivo completo")

    async def test_repeated_timeout_is_not_saved_as_empty_description(self) -> None:
        from unittest.mock import MagicMock, patch

        from playwright.async_api import Page, TimeoutError

        from scdp_automation.extrator import consult_trip_reason

        with (
            patch(
                "scdp_automation.extrator._consult_trip_reason",
                side_effect=TimeoutError("timeout"),
            ),
            self.assertRaises(TimeoutError),
        ):
            await consult_trip_reason(MagicMock(spec=Page), "000548/26")


class BrowserModuleTests(unittest.TestCase):
    def test_entry_points_use_shared_visible_chrome_connector(self) -> None:
        from scdp_automation import cli, extrator, navegador_chrome

        self.assertIs(
            cli.connect_visible_chrome, navegador_chrome.connect_visible_chrome
        )
        self.assertIs(
            extrator.connect_visible_chrome, navegador_chrome.connect_visible_chrome
        )


class CliLoginTests(unittest.TestCase):
    def test_login_option_runs_authentication_without_extractor(self) -> None:
        from unittest.mock import MagicMock, patch

        from scdp_automation import cli

        login_coroutine = object()
        with (
            patch("sys.argv", ["scdp-extrair", "--login"]),
            patch(
                "scdp_automation.cli.login_only",
                new=MagicMock(return_value=login_coroutine),
            ) as login,
            patch(
                "scdp_automation.cli.run", return_value="extract-coroutine"
            ) as extract,
            patch("scdp_automation.cli.asyncio.run") as run_async,
        ):
            cli.main()

        login.assert_called_once_with()
        run_async.assert_called_once_with(login_coroutine)
        extract.assert_not_called()


class SafeFailureTests(unittest.IsolatedAsyncioTestCase):
    async def test_query_failure_does_not_expose_original_exception_text(self) -> None:
        from unittest.mock import AsyncMock, MagicMock, patch

        from playwright.async_api import Error as PlaywrightError
        from playwright.async_api import Page

        from scdp_automation.extrator import _consult_with_safe_failure

        private_detail = "https://example.invalid/path?token=private-value"
        with (
            patch(
                "scdp_automation.extrator.consult_trip_reason",
                new=AsyncMock(side_effect=PlaywrightError(private_detail)),
            ),
            self.assertRaises(RuntimeError) as raised,
        ):
            await _consult_with_safe_failure(MagicMock(spec=Page), "999999/26", 1, 3)

        self.assertEqual(
            str(raised.exception),
            "Falha ao consultar descrição. O progresso foi preservado; execute novamente para retomar.",
        )
        self.assertTrue(raised.exception.__suppress_context__)
        self.assertNotIn(private_detail, str(raised.exception))
