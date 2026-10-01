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

        self.assertEqual(credentials.username.get_secret_value(), "12345678901")
        self.assertEqual(credentials.password.get_secret_value(), "example-secret")
        self.assertNotIn("example-secret", repr(credentials))

    def test_env_file_values_override_generic_environment_variables(self) -> None:
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

        self.assertEqual(credentials.username.get_secret_value(), "file-user")
        self.assertEqual(credentials.password.get_secret_value(), "file-password")

    def test_environment_variables_are_fallback_when_env_file_is_empty(self) -> None:
        with (
            tempfile.TemporaryDirectory() as directory,
            patch.dict(
                os.environ,
                {"USERNAME": "fallback-user", "PASSWORD": "fallback-password"},
                clear=True,
            ),
        ):
            env_file = Path(directory) / ".env"
            env_file.write_text("", encoding="utf-8")
            credentials = load_credentials(env_file)

        self.assertEqual(credentials.username.get_secret_value(), "fallback-user")
        self.assertEqual(credentials.password.get_secret_value(), "fallback-password")

    def test_missing_credentials_error_does_not_include_values(self) -> None:
        with (
            tempfile.TemporaryDirectory() as directory,
            patch.dict(os.environ, {}, clear=True),
            self.assertRaisesRegex(RuntimeError, "USERNAME.*PASSWORD"),
        ):
            load_credentials(Path(directory) / "missing.env")


class GovBrLoginTests(unittest.IsolatedAsyncioTestCase):
    async def test_clicks_continue_before_waiting_for_visible_captcha(self) -> None:
        events: list[str] = []
        page = MagicMock(spec=Page)
        username_field = MagicMock()
        username_field.fill = AsyncMock(side_effect=lambda _: events.append("username"))
        captcha_frames = MagicMock()
        captcha_frames.count = AsyncMock(return_value=2)
        visible_captcha_frames = MagicMock()
        visible_captcha_frames.count = AsyncMock(return_value=0)
        password_field = MagicMock()
        password_field.wait_for = AsyncMock(
            side_effect=lambda **_: events.append("password-visible")
        )
        password_field.fill = AsyncMock(side_effect=lambda _: events.append("password"))
        password_form = MagicMock()
        password_form.evaluate = AsyncMock(
            side_effect=lambda _: events.append("submit")
        )
        page.locator.side_effect = {
            "#accountId": username_field,
            'iframe[src*="hcaptcha.com"]': captcha_frames,
            'iframe[src*="hcaptcha.com"]:visible': visible_captcha_frames,
            'input[type="password"]': password_field,
            'form:has(input[type="password"])': password_form,
        }.__getitem__
        page.wait_for_function = AsyncMock()
        continue_button = MagicMock()
        continue_button.click = AsyncMock(side_effect=lambda: events.append("continue"))
        page.get_by_role.side_effect = [continue_button]
        page.wait_for_url = AsyncMock(
            side_effect=lambda *_args, **_kwargs: events.append("redirect")
        )
        credentials = GovBrCredentials(
            username="12345678901", password="dummy-password"
        )

        await authenticate_gov_br(page, credentials)

        self.assertLess(events.index("username"), events.index("continue"))
        self.assertLess(events.index("continue"), events.index("password-visible"))
        self.assertLess(events.index("password-visible"), events.index("password"))
        self.assertLess(events.index("password"), events.index("submit"))
        self.assertLess(events.index("submit"), events.index("redirect"))
        page.wait_for_function.assert_not_awaited()
        username_field.fill.assert_awaited_once_with("12345678901")
        password_field.fill.assert_awaited_once_with("dummy-password")
        page.locator.assert_any_call('form:has(input[type="password"])')
        password_form.evaluate.assert_awaited_once()
        self.assertIn(
            "form.requestSubmit(submitButton)",
            password_form.evaluate.await_args_list[0].args[0],
        )

    async def test_does_not_wait_for_captcha_when_no_challenge_is_present(self) -> None:
        page = MagicMock(spec=Page)
        username_field = MagicMock()
        username_field.fill = AsyncMock()
        captcha_frames = MagicMock()
        captcha_frames.count = AsyncMock(return_value=2)
        visible_captcha_frames = MagicMock()
        visible_captcha_frames.count = AsyncMock(return_value=0)
        password_field = MagicMock()
        password_field.wait_for = AsyncMock()
        password_field.fill = AsyncMock()
        password_form = MagicMock()
        password_form.evaluate = AsyncMock()
        page.locator.side_effect = {
            "#accountId": username_field,
            'iframe[src*="hcaptcha.com"]': captcha_frames,
            'iframe[src*="hcaptcha.com"]:visible': visible_captcha_frames,
            'input[type="password"]': password_field,
            'form:has(input[type="password"])': password_form,
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
        captcha_frames.count = AsyncMock(return_value=2)
        visible_captcha_frames = MagicMock()
        visible_captcha_frames.count = AsyncMock(return_value=1)
        password_field = MagicMock()
        password_field.wait_for = AsyncMock(
            side_effect=[PlaywrightTimeoutError("password not yet visible"), None]
        )
        password_field.fill = AsyncMock()
        password_form = MagicMock()
        password_form.evaluate = AsyncMock()
        page.locator.side_effect = {
            "#accountId": username_field,
            'iframe[src*="hcaptcha.com"]': captcha_frames,
            'iframe[src*="hcaptcha.com"]:visible': visible_captcha_frames,
            'input[type="password"]': password_field,
            'form:has(input[type="password"])': password_form,
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
        visible_captcha_frames.count.assert_awaited_once()


if __name__ == "__main__":
    unittest.main()
