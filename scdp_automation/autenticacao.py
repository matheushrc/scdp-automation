"""Login gov.br com credenciais locais e CAPTCHA resolvido pelo operador."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import dotenv_values
from loguru import logger
from playwright.async_api import Page
from playwright.async_api import TimeoutError as PlaywrightTimeoutError


@dataclass(frozen=True, repr=False)
class GovBrCredentials:
    username: str
    password: str


def load_credentials(env_path: Path = Path(".env")) -> GovBrCredentials:
    """Read gov.br credentials, preferring the process environment."""
    file_values = dotenv_values(env_path)
    username = os.environ.get("USERNAME") or file_values.get("USERNAME")
    password = os.environ.get("PASSWORD") or file_values.get("PASSWORD")
    missing = [
        name
        for name, value in (("USERNAME", username), ("PASSWORD", password))
        if not value
    ]
    if missing:
        names = " e ".join(missing)
        raise RuntimeError(
            f"Configure {names} no arquivo .env ou nas variáveis de ambiente."
        )
    if username is None or password is None:
        raise RuntimeError("As credenciais configuradas estão vazias.")
    return GovBrCredentials(username=username, password=password)


async def authenticate_gov_br(page: Page, credentials: GovBrCredentials) -> None:
    """Fill gov.br credentials after the operator completes any CAPTCHA."""
    await page.locator("#accountId").fill(credentials.username)
    captcha_frames = page.locator('iframe[src*="hcaptcha.com"]')
    if await captcha_frames.count():
        logger.info("Resolva o CAPTCHA na janela visível para continuar o login.")
        await page.wait_for_function(
            """() => Boolean(document.querySelector(
              'textarea[name="h-captcha-response"]')?.value.trim())""",
            polling=500,
            timeout=0,
        )

    password_field = page.locator('input[type="password"]')
    continue_button = page.get_by_role("button", name="Continuar", exact=True)
    await continue_button.click()
    try:
        await password_field.wait_for(state="visible", timeout=30_000)
    except PlaywrightTimeoutError:
        # Quando o gov.br exige CAPTCHA depois de enviar o CPF, o widget pode
        # aparecer apenas nessa etapa. Aguarda a resolução do operador e reenvia.
        late_captcha = page.locator('iframe[src*="hcaptcha.com"]')
        if not await late_captcha.count():
            raise
        logger.info("Resolva o CAPTCHA na janela visível para continuar o login.")
        await page.wait_for_function(
            """() => Boolean(document.querySelector(
              'textarea[name="h-captcha-response"]')?.value.trim())""",
            polling=500,
            timeout=0,
        )
        await continue_button.click()
        await password_field.wait_for(state="visible", timeout=120_000)
    await password_field.fill(credentials.password)
    await page.get_by_role("button", name="Entrar", exact=True).click()
    await page.wait_for_url("https://www2.scdp.gov.br/novoscdp/**", timeout=120_000)
