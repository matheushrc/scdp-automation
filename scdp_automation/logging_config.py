"""Configuração de logs operacionais sem dados de autenticação ou de viagens."""

from __future__ import annotations

import sys
from pathlib import Path

from loguru import logger

from scdp_automation.config import REPO_ROOT


def configure_logging(log_directory: Path = REPO_ROOT / "logs" / "scdp") -> None:
    """Configure logs diários com retenção de 30 dias no diretório do projeto."""
    log_directory.mkdir(parents=True, exist_ok=True)
    logger.remove()
    logger.add(
        sys.stderr,
        level="INFO",
        format="<green>{time:HH:mm:ss}</green> | <level>{level}</level> | {message}",
    )
    logger.add(
        str(log_directory / "scdp_{time:YYYY-MM-DD}.log"),
        rotation="00:00",
        retention="30 days",
        encoding="utf-8",
        format="{time:YYYY-MM-DD HH:mm:ss} | {level} | {message}",
        enqueue=True,
        backtrace=False,
        diagnose=False,
    )
