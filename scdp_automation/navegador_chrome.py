"""Abre Chrome normal e conecta Playwright via CDP local."""

from __future__ import annotations

import asyncio
import json
import os
import platform
import shutil
import socket
import urllib.error
import urllib.request
from collections.abc import Callable, Mapping
from pathlib import Path
from urllib.parse import urlsplit

from playwright.async_api import Browser, Playwright

from scdp_automation.chrome_profile_setup import selected_profile_directory

SCDP_URL = "https://www2.scdp.gov.br/"


def _free_local_port() -> int:
    with socket.socket() as listener:
        listener.bind(("127.0.0.1", 0))
        return int(listener.getsockname()[1])


def _debugger_ready(port: int) -> bool:
    try:
        with urllib.request.urlopen(
            f"http://127.0.0.1:{port}/json/version", timeout=0.5
        ) as response:
            data = json.load(response)
        endpoint = urlsplit(data.get("webSocketDebuggerUrl", ""))
        return (
            endpoint.scheme == "ws"
            and endpoint.hostname in {"127.0.0.1", "localhost"}
            and endpoint.port == port
        )
    except OSError, ValueError, json.JSONDecodeError, urllib.error.URLError:
        return False


def _profile_lock_is_live(profile: Path) -> bool:
    lock = profile / "SingletonLock"
    if not lock.is_symlink():
        return lock.exists()
    try:
        pid = int(lock.resolve().name.rsplit("-", 1)[-1])
        os.kill(pid, 0)
    except ProcessLookupError:
        for marker in profile.glob("Singleton*"):
            if marker.is_symlink():
                marker.unlink()
        return False
    except ValueError, OSError, PermissionError:
        return True
    return True


def resolve_chrome_executable(
    platform_name: str | None = None,
    environ: Mapping[str, str] | None = None,
    path_lookup: Callable[[str], str | None] | None = None,
) -> str:
    """Find Chrome Stable without shelling through a platform-specific command."""
    system = platform_name or platform.system()
    environment = os.environ if environ is None else environ
    lookup = shutil.which if path_lookup is None else path_lookup

    if system in {"Windows", "nt"}:
        candidates: list[Path] = []
        for variable in (
            "LOCALAPPDATA",
            "PROGRAMFILES",
            "PROGRAMFILES(X86)",
            "PROGRAMW6432",
        ):
            base = environment.get(variable)
            if base:
                candidates.append(
                    Path(base) / "Google" / "Chrome" / "Application" / "chrome.exe"
                )
        for candidate in candidates:
            if candidate.is_file():
                return str(candidate)
        located = lookup("chrome.exe") or lookup("chrome")
        if located:
            return str(located)
        raise RuntimeError("Não encontrei o executável do Google Chrome.")

    if system in {"Linux", "posix"}:
        located = lookup("google-chrome") or lookup("google-chrome-stable")
        if located:
            return str(located)
        raise RuntimeError("Não encontrei o executável google-chrome no PATH.")

    raise RuntimeError("A inicialização do Chrome não é suportada neste sistema.")


async def connect_visible_chrome(playwright: Playwright, profile: Path) -> Browser:
    """Attach to the clone's live Chrome, starting ordinary headed Chrome if needed."""
    profile.mkdir(parents=True, exist_ok=True)
    profile_directory = selected_profile_directory(profile)
    if not (profile / profile_directory).is_dir():
        raise RuntimeError("A cópia selecionada do perfil Chrome está incompleta.")
    port_file = profile / "cdp-port"
    if port_file.exists():
        try:
            port = int(port_file.read_text(encoding="ascii").strip())
        except OSError, ValueError:
            port = 0
        if 0 < port <= 65535 and _debugger_ready(port):
            return await playwright.chromium.connect_over_cdp(
                f"http://127.0.0.1:{port}", is_local=True, no_defaults=True
            )
        port_file.unlink(missing_ok=True)

    if _profile_lock_is_live(profile):
        raise RuntimeError(
            "O perfil .scdp-browser já está aberto no Chrome. Use a janela "
            "existente para concluir a autenticação; a conexão CDP não está ativa."
        )

    chrome = resolve_chrome_executable()
    port = _free_local_port()
    chrome_arguments = (
        chrome,
        f"--user-data-dir={profile.resolve()}",
        f"--profile-directory={profile_directory}",
        "--remote-debugging-address=127.0.0.1",
        f"--remote-debugging-port={port}",
        "--disable-extensions",
        "--disable-component-extensions-with-background-pages",
        "--no-first-run",
        SCDP_URL,
    )
    if platform.system() != "Windows":
        process = await asyncio.create_subprocess_exec(
            *chrome_arguments,
            stdin=asyncio.subprocess.DEVNULL,
            stdout=asyncio.subprocess.DEVNULL,
            stderr=asyncio.subprocess.DEVNULL,
            start_new_session=True,
        )
    else:
        process = await asyncio.create_subprocess_exec(
            *chrome_arguments,
            stdin=asyncio.subprocess.DEVNULL,
            stdout=asyncio.subprocess.DEVNULL,
            stderr=asyncio.subprocess.DEVNULL,
        )
    for _ in range(120):
        if _debugger_ready(port):
            port_file.write_text(f"{port}\n", encoding="ascii")
            port_file.chmod(0o600)
            return await playwright.chromium.connect_over_cdp(
                f"http://127.0.0.1:{port}", is_local=True, no_defaults=True
            )
        if process.returncode is not None:
            break
        await asyncio.sleep(0.25)
    raise RuntimeError(
        "O Chrome abriu, mas não iniciou a conexão CDP local. "
        "Feche a janela do clone e tente novamente."
    )
