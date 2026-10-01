"""Ponto de entrada da linha de comando."""

from __future__ import annotations

import asyncio
from pathlib import Path
from urllib.parse import urlsplit

from playwright.async_api import async_playwright

from scdp_automation.extrator import run
from scdp_automation.logging_config import configure_logging, logger
from scdp_automation.navegador_chrome import connect_visible_chrome

SCDP_URL = "https://www2.scdp.gov.br/"


async def open_browser() -> None:
    """Abre o Chrome headed usando somente o perfil local do projeto."""
    configure_logging()
    profile = Path(__file__).resolve().parents[1] / ".scdp-browser"
    async with async_playwright() as playwright:
        browser = await connect_visible_chrome(playwright, profile)
        context = browser.contexts[0]
        page = context.pages[0] if context.pages else await context.new_page()
        if urlsplit(page.url).hostname != "www2.scdp.gov.br":
            await page.goto(SCDP_URL, wait_until="domcontentloaded")
        await page.bring_to_front()
        logger.info("Chrome aberto com perfil local do projeto.")
        logger.info("Mantenha este comando em execução enquanto usar a janela.")
        await asyncio.Event().wait()


def main() -> None:
    """Executa o extrator com os argumentos recebidos do terminal."""
    import argparse

    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--abrir-navegador", action="store_true")
    known, _ = parser.parse_known_args()
    if known.abrir_navegador:
        asyncio.run(open_browser())
    else:
        asyncio.run(run())


if __name__ == "__main__":
    main()
