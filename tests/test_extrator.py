import asyncio
import unittest
from pathlib import Path
from unittest.mock import patch

from scdp_automation.extrator import (
    DEFAULT_OUTPUT,
    collect_pending_descriptions,
    extract_description,
    parse_args,
    resolve_report_url,
)
from scdp_automation.relatorio import Viagem


class ExtratorTests(unittest.TestCase):
    def test_extract_description_from_label_on_separate_line(self) -> None:
        text = "Solicitação\nDescrição do Motivo da Viagem\nParticipação em reunião"
        self.assertEqual(extract_description(text), "Participação em reunião")

    def test_output_path_is_fixed_to_json_checkpoint(self) -> None:
        with patch("sys.argv", ["scdp-extrair"]):
            args = parse_args()

        self.assertEqual(DEFAULT_OUTPUT, Path("output/viagens_scdp_2026.json"))
        self.assertEqual(args.output, DEFAULT_OUTPUT)

    def test_report_link_is_resolved_against_authenticated_scdp_page(self) -> None:
        report_url = resolve_report_url(
            "https://www2.scdp.gov.br/novoscdp/home.xhtml",
            "pages/relatorio/relatorio_viagem.xhtml?faces-redirect=true",
        )

        self.assertEqual(
            report_url,
            "https://www2.scdp.gov.br/novoscdp/pages/relatorio/"
            "relatorio_viagem.xhtml?faces-redirect=true",
        )


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

    async def test_authenticated_menu_loading_while_login_link_appears_skips_login(
        self,
    ) -> None:
        from unittest.mock import AsyncMock, MagicMock

        from playwright.async_api import Page, TimeoutError

        from scdp_automation.extrator import start_login_if_needed

        menu = MagicMock()
        menu.wait_for = AsyncMock(side_effect=TimeoutError("slow home"))
        menu.is_visible = AsyncMock(return_value=True)
        login_link = MagicMock()
        login_link.wait_for = AsyncMock()
        login_link.click = AsyncMock()
        page = MagicMock(spec=Page)
        page.url = "https://www2.scdp.gov.br/novoscdp/home.xhtml"
        page.get_by_role.return_value = menu
        page.locator.return_value = login_link

        with (
            patch("scdp_automation.extrator.authenticate_gov_br", new=AsyncMock()),
            patch("scdp_automation.extrator.load_credentials"),
        ):
            await start_login_if_needed(page)

        menu.is_visible.assert_awaited_once()
        login_link.click.assert_not_awaited()

    async def test_click_returning_to_scdp_home_uses_existing_session(self) -> None:
        from unittest.mock import AsyncMock, MagicMock

        from playwright.async_api import Page, TimeoutError

        from scdp_automation.extrator import start_login_if_needed

        menu = MagicMock()
        menu.wait_for = AsyncMock(side_effect=[TimeoutError("home slow"), None])
        menu.is_visible = AsyncMock(return_value=False)
        login_link = MagicMock()
        login_link.wait_for = AsyncMock()
        login_link.click = AsyncMock()
        page = MagicMock(spec=Page)
        page.url = "https://www2.scdp.gov.br/novoscdp/home.xhtml"
        page.get_by_role.return_value = menu
        page.locator.return_value = login_link

        async def return_to_scdp(url_pattern: object, **_: object) -> None:
            page.url = "https://www2.scdp.gov.br/novoscdp/pages/main.xhtml"

        page.wait_for_url = AsyncMock(side_effect=return_to_scdp)
        with patch(
            "scdp_automation.extrator.authenticate_gov_br", new=AsyncMock()
        ) as authenticate:
            await start_login_if_needed(page)

        wait_call = page.wait_for_url.await_args
        assert wait_call is not None
        self.assertEqual(
            wait_call.args[0].pattern,
            r"^https://(?:sso\.acesso\.gov\.br/login|www2\.scdp\.gov\.br/novoscdp/pages/main\.xhtml)",
        )
        menu.wait_for.assert_awaited_with(state="visible", timeout=30_000)
        authenticate.assert_not_awaited()


class DescriptionCollectionTests(unittest.IsolatedAsyncioTestCase):
    async def test_single_tab_serializes_and_checkpoints_each_description(self) -> None:
        from tempfile import TemporaryDirectory
        from unittest.mock import AsyncMock, MagicMock

        from playwright.async_api import Page

        page = MagicMock(spec=Page)
        page.bring_to_front = AsyncMock()
        trips = [
            Viagem.model_construct(
                numero_da_solicitacao=f"00000{i}/26",
                descricao_do_motivo_da_viagem=None,
            )
            for i in range(1, 3)
        ]
        active_queries = 0

        async def consult(*_: object) -> str:
            nonlocal active_queries
            self.assertEqual(active_queries, 0)
            active_queries += 1
            await asyncio.sleep(0)
            active_queries -= 1
            return "Descrição consultada"

        with TemporaryDirectory() as directory:
            output = Path(directory) / "viagens.json"
            with (
                patch(
                    "scdp_automation.extrator._consult_with_safe_failure",
                    new=AsyncMock(side_effect=consult),
                ),
                patch("scdp_automation.extrator.save_json") as save_json,
            ):
                await collect_pending_descriptions(page, trips, trips, output)

        self.assertEqual(save_json.call_count, 2)
        self.assertEqual(page.bring_to_front.await_count, 2)
        self.assertEqual(
            [trip.descricao_do_motivo_da_viagem for trip in trips],
            ["Descrição consultada", "Descrição consultada"],
        )


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
    def test_help_does_not_trigger_profile_selection(self) -> None:
        from unittest.mock import patch

        from scdp_automation import cli

        with (
            patch("sys.argv", ["scdp-extrair", "--help"]),
            patch("scdp_automation.cli.prepare_chrome_profile") as prepare_profile,
            self.assertRaises(SystemExit) as raised,
        ):
            cli.main()

        self.assertEqual(raised.exception.code, 0)
        prepare_profile.assert_not_called()

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
            patch("scdp_automation.cli.prepare_chrome_profile") as prepare_profile,
            patch(
                "scdp_automation.cli.run", return_value="extract-coroutine"
            ) as extract,
            patch("scdp_automation.cli.asyncio.run") as run_async,
        ):
            cli.main()

        login.assert_called_once_with()
        prepare_profile.assert_called_once()
        run_async.assert_called_once_with(login_coroutine)
        extract.assert_not_called()

    def test_profile_selection_runs_before_login_action(self) -> None:
        from unittest.mock import MagicMock, patch

        from scdp_automation import cli

        login_coroutine = object()
        with (
            patch(
                "sys.argv",
                [
                    "scdp-extrair",
                    "--selecionar-perfil-chrome",
                    "--login",
                ],
            ),
            patch("scdp_automation.cli.prepare_chrome_profile") as prepare_profile,
            patch(
                "scdp_automation.cli.login_only",
                new=MagicMock(return_value=login_coroutine),
            ) as login,
            patch("scdp_automation.cli.asyncio.run") as run_async,
        ):
            cli.main()

        prepare_profile.assert_called_once()
        self.assertTrue(prepare_profile.call_args.kwargs["force_reselect"])
        login.assert_called_once_with()
        run_async.assert_called_once_with(login_coroutine)

    def test_profile_selector_flag_is_removed_before_extractor_arguments(self) -> None:
        from unittest.mock import MagicMock, patch

        from scdp_automation import cli

        extraction_coroutine = object()
        with (
            patch(
                "sys.argv",
                [
                    "scdp-extrair",
                    "--selecionar-perfil-chrome",
                    "--limite",
                    "2",
                ],
            ),
            patch("scdp_automation.cli.prepare_chrome_profile"),
            patch(
                "scdp_automation.cli.run",
                new=MagicMock(return_value=extraction_coroutine),
            ) as extract,
            patch("scdp_automation.cli.asyncio.run") as run_async,
        ):
            cli.main()

        extract.assert_called_once_with(["--limite", "2"])
        run_async.assert_called_once_with(extraction_coroutine)


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
