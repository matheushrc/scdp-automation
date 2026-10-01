"""Abre Chrome normal e conecta Playwright via CDP local."""

from __future__ import annotations

import asyncio
import json
import os
import shutil
import socket
import urllib.error
import urllib.request
from pathlib import Path
from urllib.parse import urlsplit

from playwright.async_api import Browser, Playwright


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
        return False
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


async def connect_visible_chrome(playwright: Playwright, profile: Path) -> Browser:
    """Attach to the clone's live Chrome, starting ordinary headed Chrome if needed."""
    profile.mkdir(parents=True, exist_ok=True)
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

    chrome = shutil.which("google-chrome")
    if not chrome:
        raise RuntimeError("Não encontrei o executável google-chrome no PATH.")
    port = _free_local_port()
    process = await asyncio.create_subprocess_exec(
        chrome,
        f"--user-data-dir={profile.resolve()}",
        "--profile-directory=Default",
        "--remote-debugging-address=127.0.0.1",
        f"--remote-debugging-port={port}",
        "--disable-extensions",
        "--disable-component-extensions-with-background-pages",
        "--no-first-run",
        "about:blank",
        stdin=asyncio.subprocess.DEVNULL,
        stdout=asyncio.subprocess.DEVNULL,
        stderr=asyncio.subprocess.DEVNULL,
        start_new_session=True,
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
