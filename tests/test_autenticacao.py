import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock, patch

from playwright.async_api import Page
from playwright.async_api import TimeoutError as PlaywrightTimeoutError

from scdp_automation.autenticacao import (
    GovBrCredentials,
    authenticate_gov_br,
    load_credentials,
)


class CredentialsTests(unittest.TestCase):
    def test_reads_username_and_password_from_env_file(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            env_file = Path(directory) / ".env"
            env_file.write_text(
                "USERNAME=12345678901\nPASSWORD=example-secret\n", encoding="utf-8"
            )
            with patch.dict(os.environ, {}, clear=True):
                credentials = load_credentials(env_file)

        self.assertEqual(credentials.username, "12345678901")
        self.assertEqual(credentials.password, "example-secret")
        self.assertNotIn("example-secret", repr(credentials))

    def test_environment_values_override_env_file(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            env_file = Path(directory) / ".env"
            env_file.write_text(
                "USERNAME=file-user\nPASSWORD=file-password\n", encoding="utf-8"
            )
            with patch.dict(
                os.environ,
                {"USERNAME": "environment-user", "PASSWORD": "environment-password"},
                clear=True,
            ):
                credentials = load_credentials(env_file)

        self.assertEqual(credentials.username, "environment-user")
        self.assertEqual(credentials.password, "environment-password")

    def test_missing_credentials_error_does_not_include_values(self) -> None:
        with (
            tempfile.TemporaryDirectory() as directory,
            patch.dict(os.environ, {}, clear=True),
            self.assertRaisesRegex(RuntimeError, "USERNAME.*PASSWORD"),
        ):
            load_credentials(Path(directory) / "missing.env")


class GovBrLoginTests(unittest.IsolatedAsyncioTestCase):
    async def test_password_and_submit_wait_for_manual_captcha_completion(self) -> None:
        events: list[str] = []
        page = MagicMock(spec=Page)
        username_field = MagicMock()
        username_field.fill = AsyncMock(side_effect=lambda _: events.append("username"))
        captcha_frames = MagicMock()
        captcha_frames.count = AsyncMock(return_value=1)
        password_field = MagicMock()
        password_field.wait_for = AsyncMock(
            side_effect=lambda **_: events.append("password-visible")
        )
        password_field.fill = AsyncMock(side_effect=lambda _: events.append("password"))
        page.locator.side_effect = {
            "#accountId": username_field,
            'iframe[src*="hcaptcha.com"]': captcha_frames,
            'input[type="password"]': password_field,
        }.__getitem__
        page.wait_for_function = AsyncMock(
            side_effect=lambda *_args, **_kwargs: events.append("captcha-passed")
        )
        continue_button = MagicMock()
        continue_button.click = AsyncMock(side_effect=lambda: events.append("continue"))
        login_button = MagicMock()
        login_button.click = AsyncMock(side_effect=lambda: events.append("submit"))
        page.get_by_role.side_effect = [continue_button, login_button]
        page.wait_for_url = AsyncMock(
            side_effect=lambda *_args, **_kwargs: events.append("redirect")
        )
        credentials = GovBrCredentials(
            username="12345678901", password="dummy-password"
        )

        await authenticate_gov_br(page, credentials)

        self.assertLess(events.index("username"), events.index("captcha-passed"))
        self.assertLess(events.index("captcha-passed"), events.index("continue"))
        self.assertLess(events.index("continue"), events.index("password-visible"))
        self.assertLess(events.index("password-visible"), events.index("password"))
        self.assertLess(events.index("password"), events.index("submit"))
        self.assertLess(events.index("submit"), events.index("redirect"))
        username_field.fill.assert_awaited_once_with("12345678901")
        password_field.fill.assert_awaited_once_with("dummy-password")

    async def test_does_not_wait_for_captcha_when_no_challenge_is_present(self) -> None:
        page = MagicMock(spec=Page)
        username_field = MagicMock()
        username_field.fill = AsyncMock()
        captcha_frames = MagicMock()
        captcha_frames.count = AsyncMock(return_value=0)
        password_field = MagicMock()
        password_field.wait_for = AsyncMock()
        password_field.fill = AsyncMock()
        page.locator.side_effect = {
            "#accountId": username_field,
            'iframe[src*="hcaptcha.com"]': captcha_frames,
            'input[type="password"]': password_field,
        }.__getitem__
        page.wait_for_function = AsyncMock()
        page.get_by_role.side_effect = [
            MagicMock(click=AsyncMock()),
            MagicMock(click=AsyncMock()),
        ]
        page.wait_for_url = AsyncMock()

        await authenticate_gov_br(
            page, GovBrCredentials(username="12345678901", password="dummy-password")
        )

        page.wait_for_function.assert_not_awaited()

    async def test_waits_for_captcha_shown_after_cpf_submission(self) -> None:
        page = MagicMock(spec=Page)
        username_field = MagicMock()
        username_field.fill = AsyncMock()
        captcha_frames = MagicMock()
        captcha_frames.count = AsyncMock(side_effect=[0, 1])
        password_field = MagicMock()
        password_field.wait_for = AsyncMock(
            side_effect=[PlaywrightTimeoutError("password not yet visible"), None]
        )
        password_field.fill = AsyncMock()
        page.locator.side_effect = {
            "#accountId": username_field,
            'iframe[src*="hcaptcha.com"]': captcha_frames,
            'input[type="password"]': password_field,
        }.__getitem__
        page.wait_for_function = AsyncMock()
        continue_button = MagicMock()
        continue_button.click = AsyncMock()
        login_button = MagicMock()
        login_button.click = AsyncMock()
        page.get_by_role.side_effect = [continue_button, login_button]
        page.wait_for_url = AsyncMock()

        await authenticate_gov_br(
            page, GovBrCredentials(username="12345678901", password="dummy-password")
        )

        page.wait_for_function.assert_awaited_once()
        self.assertEqual(continue_button.click.await_count, 2)


if __name__ == "__main__":
    unittest.main()
