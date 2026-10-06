import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from scdp_automation.output_history import OutputHistory


class OutputHistoryTests(unittest.TestCase):
    def test_keeps_current_and_four_previous_execution_pairs_in_backup_directory(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            checkpoint = root / "viagens_scdp_2026.json"
            workbook = root / "gastos_scdp_2026.xlsx"
            for index in range(8):
                history = OutputHistory(checkpoint, workbook)
                history.archive_previous()
                checkpoint.write_text(f"json-{index}")
                workbook.write_text(f"xlsx-{index}")
                history.archive_previous()
                history.prune()
            self.assertEqual(
                set(root.iterdir()), {checkpoint, workbook, root / "backup"}
            )
            self.assertEqual(len(list((root / "backup").iterdir())), 8)
            self.assertEqual(
                {p.read_text() for p in root.rglob("*.json")},
                {f"json-{index}" for index in range(3, 8)},
            )
            self.assertEqual(
                {p.read_text() for p in root.rglob("*.xlsx")},
                {f"xlsx-{index}" for index in range(3, 8)},
            )

    def test_does_not_archive_an_incomplete_pair_or_delete_unrelated_files(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            checkpoint = root / "viagens_scdp_2026.json"
            workbook = root / "gastos_scdp_2026.xlsx"
            checkpoint.write_text("checkpoint")
            unrelated = root / "manual.xlsx"
            unrelated.write_text("manual")
            history = OutputHistory(checkpoint, workbook)
            self.assertIsNone(history.archive_previous())
            history.prune()
            self.assertEqual(set(root.iterdir()), {checkpoint, unrelated})

    def test_failed_copy_preserves_original_pair_and_removes_partial_archive(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            checkpoint = root / "viagens.json"
            workbook = root / "gastos.xlsx"
            checkpoint.write_text("json")
            workbook.write_text("xlsx")
            history = OutputHistory(checkpoint, workbook)
            copy = shutil.copy2

            def fail_workbook(source, destination):
                if source == workbook:
                    raise OSError("copy failed")
                return copy(source, destination)

            with (
                patch(
                    "scdp_automation.output_history.shutil.copy2",
                    side_effect=fail_workbook,
                ),
                self.assertRaises(OSError),
            ):
                history.archive_previous()
            self.assertEqual(list(history.directory.iterdir()), [])
            self.assertEqual(checkpoint.read_text(), "json")
            self.assertEqual(workbook.read_text(), "xlsx")
            self.assertFalse(history.archived)
