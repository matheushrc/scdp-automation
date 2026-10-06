"""Keep the active output pair and four previous execution pairs in output/backup."""

import shutil
from datetime import UTC, datetime
from pathlib import Path
from uuid import uuid4


class OutputHistory:
    def __init__(self, checkpoint: Path, workbook: Path) -> None:
        if checkpoint.parent.resolve() != workbook.parent.resolve():
            raise ValueError("A planilha e o JSON devem estar na mesma pasta.")
        self.checkpoint = checkpoint
        self.workbook = workbook
        self.directory = workbook.parent / "backup"
        self.execution_id = (
            datetime.now(UTC).strftime("%Y%m%dT%H%M%S%fZ") + uuid4().hex[:8]
        )
        self.archived = False
        self.backup: Path | None = None

    def archive_previous(self) -> Path | None:
        if self.archived:
            return self.backup
        if self.checkpoint.is_file() and self.workbook.is_file():
            self.directory.mkdir(parents=True, exist_ok=True)
            checkpoint_copy = self.directory / (
                f"{self.checkpoint.stem}.history.{self.execution_id}.json"
            )
            workbook_copy = self.directory / (
                f"{self.workbook.stem}.history.{self.execution_id}.xlsx"
            )
            try:
                shutil.copy2(self.checkpoint, checkpoint_copy)
                shutil.copy2(self.workbook, workbook_copy)
            except OSError:
                checkpoint_copy.unlink(missing_ok=True)
                workbook_copy.unlink(missing_ok=True)
                raise
            self.backup = workbook_copy
        self.archived = True
        return self.backup

    def prune(self) -> None:
        pairs = []
        prefix = f"{self.workbook.stem}.history."
        for workbook in self.directory.glob(f"{prefix}*.xlsx"):
            execution_id = workbook.name[len(prefix) : -len(".xlsx")]
            checkpoint = self.directory / (
                f"{self.checkpoint.stem}.history.{execution_id}.json"
            )
            if checkpoint.is_file():
                pairs.append((execution_id, checkpoint, workbook))
        for _, checkpoint, workbook in sorted(pairs, reverse=True)[4:]:
            checkpoint.unlink()
            workbook.unlink()
        # Legacy publication backups were created several times per execution.
        if self.workbook.is_file() and self.checkpoint.is_file():
            for backup in self.workbook.parent.glob(
                f"{self.workbook.stem}.backup.*.xlsx"
            ):
                backup.unlink()
