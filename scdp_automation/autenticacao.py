"""Login gov.br com credenciais locais e CAPTCHA resolvido pelo operador."""

from __future__ import annotations

from pathlib import Path

from loguru import logger
from playwright.async_api import Page
from playwright.async_api import TimeoutError as PlaywrightTimeoutError
from pydantic import Field, SecretStr, ValidationError
from pydantic_settings import (
    BaseSettings,
    PydanticBaseSettingsSource,
    SettingsConfigDict,
)


class GovBrCredentials(BaseSettings):
    """Validated credentials; the .env file outranks generic shell variables."""

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
        populate_by_name=True,
    )

    username: SecretStr = Field(validation_alias="USERNAME", min_length=1)
    password: SecretStr = Field(validation_alias="PASSWORD", min_length=1)

    @classmethod
    def settings_customise_sources(
        cls,
        settings_cls: type[BaseSettings],
        init_settings: PydanticBaseSettingsSource,
        env_settings: PydanticBaseSettingsSource,
        dotenv_settings: PydanticBaseSettingsSource,
        file_secret_settings: PydanticBaseSettingsSource,
    ) -> tuple[PydanticBaseSettingsSource, ...]:
        return init_settings, dotenv_settings, env_settings, file_secret_settings


def load_credentials(env_path: Path = Path(".env")) -> GovBrCredentials:
    """Load and validate gov.br credentials from the configured dotenv file."""
    try:
        return GovBrCredentials(_env_file=env_path)
    except ValidationError:
        raise RuntimeError(
            "Configure USERNAME e PASSWORD no arquivo .env ou no ambiente."
        ) from None


async def authenticate_gov_br(page: Page, credentials: GovBrCredentials) -> None:
    """Fill gov.br credentials after the operator completes any CAPTCHA."""
    await page.locator("#accountId").fill(credentials.username.get_secret_value())
    password_field = page.locator('input[type="password"]')
    continue_button = page.get_by_role("button", name="Continuar", exact=True)
    await continue_button.click()
    try:
        await password_field.wait_for(state="visible", timeout=30_000)
    except PlaywrightTimeoutError:
        # O gov.br pode carregar iframes hCaptcha ocultos antes do desafio real.
        # Só espera quando o iframe está visível após enviar o CPF.
        captcha_frames = page.locator('iframe[src*="hcaptcha.com"]:visible')
        if not await captcha_frames.count():
            raise
        logger.info("Resolva o CAPTCHA na janela visível e aguarde a etapa da senha.")
        await page.wait_for_function(
            """() => Boolean(document.querySelector(
              'textarea[name="h-captcha-response"]')?.value.trim())""",
            polling=500,
            timeout=0,
        )
        await continue_button.click()
        await password_field.wait_for(state="visible", timeout=120_000)
    await password_field.fill(credentials.password.get_secret_value())
    password_form = page.locator('form:has(input[type="password"])')
    await password_form.evaluate(
        """form => {
          const submitButton = form.querySelector('button[type="submit"]');
          if (submitButton) {
            form.requestSubmit(submitButton);
          } else {
            form.requestSubmit();
          }
        }"""
    )
    await page.wait_for_url("https://www2.scdp.gov.br/novoscdp/**", timeout=120_000)
