"""Serialize output sessions and keep four previous execution pairs across years."""

from __future__ import annotations

import hashlib
import re
import shutil
import tempfile
from datetime import UTC, datetime
from pathlib import Path
from types import TracebackType
from typing import Self
from uuid import uuid4

from filelock import FileLock

_ANNUAL_WORKBOOK = re.compile(r"gastos_scdp_(\d{4})\.xlsx\Z")
_HISTORY_WORKBOOK = re.compile(
    r"gastos_scdp_(\d{4})\.history\.(\d{8}T\d{12}Z[0-9a-f]{8})\.xlsx\Z"
)
_RECOVERY = re.compile(r"recovery\.[0-9a-f]{32}\.xlsx\Z")


def _execution_id() -> str:
    return datetime.now(UTC).strftime("%Y%m%dT%H%M%S%fZ") + uuid4().hex[:8]


class OutputHistory:
    def __init__(self, checkpoint: Path, workbook: Path) -> None:
        if checkpoint.parent.resolve() != workbook.parent.resolve():
            raise ValueError("A planilha e o JSON devem estar na mesma pasta.")
        self.checkpoint = checkpoint
        self.workbook = workbook
        self.directory = workbook.parent / "backup"
        self.execution_id = _execution_id()
        identifier = hashlib.sha256(str(workbook.parent.resolve()).encode()).hexdigest()
        self.lock = FileLock(
            str(Path(tempfile.gettempdir()) / f"scdp-output-{identifier}.lock"),
            timeout=30,
        )
        self.archived = False
        self.backup: Path | None = None
        self.previous_pairs: list[tuple[Path, Path]] = []

    def __enter__(self) -> Self:
        self.lock.acquire()
        self.execution_id = _execution_id()
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_value: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        self.lock.release()

    def _current_pairs(self) -> list[tuple[Path, Path]]:
        pairs = []
        for workbook in sorted(self.workbook.parent.glob("gastos_scdp_*.xlsx")):
            match = _ANNUAL_WORKBOOK.fullmatch(workbook.name)
            if match:
                checkpoint = workbook.with_name(f"viagens_scdp_{match[1]}.json")
                if checkpoint.is_file():
                    pairs.append((checkpoint, workbook))
        current = (self.checkpoint, self.workbook)
        if all(path.is_file() for path in current) and current not in pairs:
            pairs.append(current)
        return pairs

    def archive_previous(self) -> Path | None:
        if self.archived:
            return self.backup
        pairs = self._current_pairs()
        copies: list[Path] = []
        try:
            for checkpoint, workbook in pairs:
                self.directory.mkdir(parents=True, exist_ok=True)
                checkpoint_copy = self.directory / (
                    f"{checkpoint.stem}.history.{self.execution_id}.json"
                )
                workbook_copy = self.directory / (
                    f"{workbook.stem}.history.{self.execution_id}.xlsx"
                )
                copies.extend((checkpoint_copy, workbook_copy))
                shutil.copy2(checkpoint, checkpoint_copy)
                shutil.copy2(workbook, workbook_copy)
                self.backup = workbook_copy
        except OSError:
            for path in copies:
                path.unlink(missing_ok=True)
            self.backup = None
            raise
        self.previous_pairs = pairs
        self.archived = True
        return self.backup

    def prune(self) -> None:
        # Only the publisher's successful caller invokes this after both files exist.
        if not self.checkpoint.is_file() or not self.workbook.is_file():
            return
        for checkpoint, workbook in self.previous_pairs:
            if workbook != self.workbook:
                checkpoint.unlink(missing_ok=True)
                workbook.unlink(missing_ok=True)
        pairs = []
        prefix = f"{self.workbook.stem}.history."
        for workbook in self.directory.glob("*.xlsx"):
            match = _HISTORY_WORKBOOK.fullmatch(workbook.name)
            if match:
                execution_id = match[2]
                checkpoint = self.directory / (
                    f"viagens_scdp_{match[1]}.history.{execution_id}.json"
                )
            elif workbook.name.startswith(prefix):
                execution_id = workbook.name[len(prefix) : -len(".xlsx")]
                if not re.fullmatch(r"\d{8}T\d{12}Z[0-9a-f]{8}", execution_id):
                    continue
                checkpoint = self.directory / (
                    f"{self.checkpoint.stem}.history.{execution_id}.json"
                )
            else:
                continue
            if checkpoint.is_file():
                pairs.append((execution_id, checkpoint, workbook))
        for _, checkpoint, workbook in sorted(pairs, reverse=True)[4:]:
            checkpoint.unlink()
            workbook.unlink()
        for path in self.directory.glob("recovery.*.xlsx"):
            if _RECOVERY.fullmatch(path.name):
                path.unlink()
