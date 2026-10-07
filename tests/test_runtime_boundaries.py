import os
import subprocess
import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock, patch

from playwright.async_api import Browser

from scdp_automation import navegador_chrome, xlsx_output
from scdp_automation.config import REPO_ROOT, ExtractionConfig
from scdp_automation.extrator import parse_args
from scdp_automation.logging_config import configure_logging


class RuntimePathsTests(unittest.TestCase):
    def test_paths_are_checkout_relative_from_other_cwd(self):
        with TemporaryDirectory() as directory:
            previous = Path.cwd()
            try:
                os.chdir(directory)
                with patch(
                    "scdp_automation.extrator.sync_extraction_config",
                    return_value=ExtractionConfig(2026),
                ):
                    args = parse_args([])
                self.assertEqual(
                    args.output, REPO_ROOT / "output" / "viagens_scdp_2026.json"
                )
                self.assertEqual(
                    args.workbook, REPO_ROOT / "output" / "gastos_scdp_2026.xlsx"
                )
                self.assertEqual(
                    xlsx_output.DEFAULT_TEMPLATE,
                    REPO_ROOT / "input" / "gastos_scdp_template.xlsx",
                )
                with patch("scdp_automation.logging_config.logger") as logger:
                    configure_logging()
                self.assertEqual(
                    Path(logger.add.call_args_list[-1].args[0]).parent,
                    REPO_ROOT / "logs" / "scdp",
                )
            finally:
                os.chdir(previous)

    def test_package_import_does_not_load_cli(self):
        result = subprocess.run(
            [
                sys.executable,
                "-c",
                "import sys; import scdp_automation; assert 'scdp_automation.cli' not in sys.modules",
            ],
            check=False,
            capture_output=True,
            text=True,
        )
        self.assertEqual(result.returncode, 0, result.stderr)


class PageSelectionTests(unittest.IsolatedAsyncioTestCase):
    async def test_selects_existing_scdp_or_gov_page_without_creating_or_navigating(
        self,
    ):
        for url in (
            "https://www2.scdp.gov.br/novoscdp/",
            "https://sso.acesso.gov.br/login",
        ):
            with self.subTest(url=url):
                unrelated = SimpleNamespace(url="https://example.com")
                selected = SimpleNamespace(url=url)
                context = SimpleNamespace(
                    pages=[unrelated, selected],
                    new_page=AsyncMock(side_effect=AssertionError("unnecessary page")),
                )
                self.assertIs(
                    await navegador_chrome.select_scdp_page(
                        MagicMock(spec=Browser, contexts=[context])
                    ),
                    selected,
                )

    async def test_creates_page_only_for_empty_context(self):
        selected = SimpleNamespace(url="about:blank")
        context = SimpleNamespace(pages=[], new_page=AsyncMock(return_value=selected))
        self.assertIs(
            await navegador_chrome.select_scdp_page(
                MagicMock(spec=Browser, contexts=[context])
            ),
            selected,
        )

    async def test_reuses_unrelated_page_when_no_scdp_page_exists(self):
        selected = SimpleNamespace(url="https://example.com")
        context = SimpleNamespace(
            pages=[selected],
            new_page=AsyncMock(side_effect=AssertionError("unnecessary page")),
        )
        self.assertIs(
            await navegador_chrome.select_scdp_page(
                MagicMock(spec=Browser, contexts=[context])
            ),
            selected,
        )

    async def test_rejects_browser_without_context(self):
        with self.assertRaises(RuntimeError):
            await navegador_chrome.select_scdp_page(
                MagicMock(spec=Browser, contexts=[])
            )


class BrowserActionTests(unittest.IsolatedAsyncioTestCase):
    async def test_open_browser_navigates_gov_tab_to_scdp(self):
        from scdp_automation import cli

        destinations = []

        async def navigate(url, *, wait_until):
            destinations.append((url, wait_until))

        page = SimpleNamespace(
            url="https://sso.acesso.gov.br/login",
            goto=navigate,
            bring_to_front=AsyncMock(),
        )
        browser = MagicMock(
            spec=Browser,
            contexts=[SimpleNamespace(pages=[page])],
        )
        manager = MagicMock()
        manager.__aenter__ = AsyncMock(return_value=object())
        manager.__aexit__ = AsyncMock(return_value=False)
        event = SimpleNamespace(wait=AsyncMock())
        with (
            patch.object(cli, "configure_logging"),
            patch.object(cli, "async_playwright", return_value=manager),
            patch.object(
                cli, "connect_visible_chrome", AsyncMock(return_value=browser)
            ),
            patch.object(cli.asyncio, "Event", return_value=event),
        ):
            await cli.open_browser()
        self.assertEqual(
            destinations, [("https://www2.scdp.gov.br/", "domcontentloaded")]
        )

    async def test_login_keeps_existing_gov_tab_for_authentication(self):
        from scdp_automation import cli

        page = SimpleNamespace(
            url="https://sso.acesso.gov.br/login",
            goto=AsyncMock(side_effect=AssertionError("must preserve gov login")),
            bring_to_front=AsyncMock(),
        )
        browser = MagicMock(spec=Browser, contexts=[SimpleNamespace(pages=[page])])
        manager = MagicMock()
        manager.__aenter__ = AsyncMock(return_value=object())
        manager.__aexit__ = AsyncMock(return_value=False)
        authenticated = []

        async def authenticate(selected):
            authenticated.append(selected.url)

        with (
            patch.object(cli, "configure_logging"),
            patch.object(cli, "async_playwright", return_value=manager),
            patch.object(
                cli, "connect_visible_chrome", AsyncMock(return_value=browser)
            ),
            patch.object(cli, "start_login_if_needed", authenticate),
            patch.object(cli, "wait_for_login", AsyncMock()),
        ):
            await cli.login_only()
        self.assertEqual(authenticated, ["https://sso.acesso.gov.br/login"])
