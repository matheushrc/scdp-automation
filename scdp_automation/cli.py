"""Ponto de entrada da linha de comando."""

from __future__ import annotations

import argparse
import asyncio
from urllib.parse import urlsplit

from playwright.async_api import async_playwright

from scdp_automation.chrome_profile_setup import prepare_chrome_profile
from scdp_automation.chrome_profiles import (
    ChromeProfileError,
    ProfileSelectionCancelled,
)
from scdp_automation.config import REPO_ROOT
from scdp_automation.extrator import (
    parse_args,
    run,
    start_login_if_needed,
    wait_for_login,
)
from scdp_automation.logging_config import configure_logging, logger
from scdp_automation.navegador_chrome import connect_visible_chrome, select_scdp_page

SCDP_URL = "https://www2.scdp.gov.br/"


async def open_browser() -> None:
    """Abre o Chrome headed usando somente o perfil local do projeto."""
    configure_logging()
    profile = REPO_ROOT / ".scdp-browser"
    async with async_playwright() as playwright:
        browser = await connect_visible_chrome(playwright, profile)
        page = await select_scdp_page(browser)
        if urlsplit(page.url).hostname != "www2.scdp.gov.br":
            await page.goto(SCDP_URL, wait_until="domcontentloaded")
        await page.bring_to_front()
        logger.info("Chrome aberto com perfil local do projeto.")
        logger.info("Mantenha este comando em execução enquanto usar a janela.")
        await asyncio.Event().wait()


async def login_only() -> None:
    """Autentica no gov.br e retorna ao SCDP sem iniciar a extração."""
    configure_logging()
    profile = REPO_ROOT / ".scdp-browser"
    async with async_playwright() as playwright:
        browser = await connect_visible_chrome(playwright, profile)
        page = await select_scdp_page(browser)
        await page.bring_to_front()
        if urlsplit(page.url).hostname not in {
            "www2.scdp.gov.br",
            "sso.acesso.gov.br",
            "acesso.gov.br",
        }:
            await page.goto(SCDP_URL, wait_until="domcontentloaded")
        await start_login_if_needed(page)
        await wait_for_login(page)
        logger.info("Login concluído; extração não iniciada.")
        # Encerrar o Playwright desconecta a automação; mantém o Chrome visível.


def main(argv: list[str] | None = None) -> None:
    """Executa o extrator com os argumentos recebidos do terminal."""
    parser = argparse.ArgumentParser(add_help=False)
    actions = parser.add_mutually_exclusive_group()
    actions.add_argument("--abrir-navegador", action="store_true")
    actions.add_argument("--login", action="store_true")
    parser.add_argument("--selecionar-perfil-chrome", action="store_true")
    known, extraction_args = parser.parse_known_args(argv)
    if "-h" in extraction_args or "--help" in extraction_args:
        parse_args(extraction_args)
    repo_root = REPO_ROOT
    try:
        prepare_chrome_profile(repo_root, force_reselect=known.selecionar_perfil_chrome)
    except ProfileSelectionCancelled as exc:
        parser.exit(0, f"{exc}\n")
    except (ChromeProfileError, OSError) as exc:
        parser.error(str(exc))
    if known.login:
        asyncio.run(login_only())
    elif known.abrir_navegador:
        asyncio.run(open_browser())
    else:
        asyncio.run(run(extraction_args))


if __name__ == "__main__":
    main()
