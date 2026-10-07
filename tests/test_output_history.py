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

    def test_history_is_global_across_years(self):
        # Per-year pruning would leave both annual pairs and eight histories.
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory)
            for year in (2026, 2027):
                for index in range(6):
                    checkpoint = output / f"viagens_scdp_{year}.json"
                    workbook = output / f"gastos_scdp_{year}.xlsx"
                    history = OutputHistory(checkpoint, workbook)
                    history.archive_previous()
                    checkpoint.write_text(f"{year}-{index}")
                    workbook.write_text(f"{year}-{index}")
                    history.prune()
            self.assertEqual(
                {p.name for p in output.iterdir()},
                {"gastos_scdp_2027.xlsx", "viagens_scdp_2027.json", "backup"},
            )
            self.assertEqual(len(list((output / "backup").iterdir())), 8)
            for extension in ("json", "xlsx"):
                self.assertEqual(
                    {p.read_text() for p in (output / "backup").glob(f"*.{extension}")},
                    {"2027-1", "2027-2", "2027-3", "2027-4"},
                )

    def test_year_rollover_failure_preserves_previous_pair(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            old_json = root / "viagens_scdp_2026.json"
            old_book = root / "gastos_scdp_2026.xlsx"
            old_json.write_text("old-json")
            old_book.write_text("old-book")
            history = OutputHistory(
                root / "viagens_scdp_2027.json", root / "gastos_scdp_2027.xlsx"
            )
            backup = history.archive_previous()
            self.assertIsNotNone(backup)
            assert backup is not None
            history.checkpoint.write_text("new-json")
            # No publication: even an accidental prune must preserve the old pair.
            history.prune()
            self.assertEqual(old_json.read_text(), "old-json")
            self.assertEqual(old_book.read_text(), "old-book")
            self.assertEqual(backup.read_text(), "old-book")

    def test_orphan_and_manual_files_are_preserved(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            orphan = root / "gastos_scdp_2025.xlsx"
            manual = root / "manual.xlsx"
            orphan.write_text("orphan")
            manual.write_text("manual")
            history = OutputHistory(
                root / "viagens_scdp_2027.json", root / "gastos_scdp_2027.xlsx"
            )
            self.assertIsNone(history.archive_previous())
            history.checkpoint.write_text("json")
            history.workbook.write_text("xlsx")
            history.prune()
            self.assertEqual(orphan.read_text(), "orphan")
            self.assertEqual(manual.read_text(), "manual")
            self.assertFalse((root / "viagens_scdp_2025.json").exists())

    def test_two_publications_archive_once(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            history = OutputHistory(root / "viagens.json", root / "gastos.xlsx")
            history.checkpoint.write_text("old-json")
            history.workbook.write_text("old-book")
            first = history.archive_previous()
            assert first is not None
            history.checkpoint.write_text("new-json")
            history.workbook.write_text("new-book")
            self.assertEqual(history.archive_previous(), first)
            history.prune()
            self.assertEqual(len(list(history.directory.iterdir())), 2)
            self.assertEqual(first.read_text(), "old-book")

    def test_history_orders_sessions_by_acquisition(self):
        # Creating a waiting session early must not make its later snapshot old.
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            checkpoint = root / "viagens_scdp_2026.json"
            workbook = root / "gastos_scdp_2026.xlsx"
            sessions = [OutputHistory(checkpoint, workbook) for _ in range(6)]
            for index, session in enumerate(reversed(sessions)):
                with session as history:
                    history.archive_previous()
                    checkpoint.write_text(str(index))
                    workbook.write_text(str(index))
                    history.prune()
            self.assertEqual(
                {p.read_text() for p in (root / "backup").glob("*.json")},
                {"1", "2", "3", "4"},
            )

    def test_output_session_serializes_processes(self):
        # Actual independent processes signal readiness, acquisition and release.
        import multiprocessing

        context = multiprocessing.get_context("spawn")
        with tempfile.TemporaryDirectory() as directory:
            first_entered = context.Event()
            first_release = context.Event()
            second_ready = context.Event()
            second_entered = context.Event()
            second_release = context.Event()
            processes = [
                context.Process(
                    target=hold_output_session,
                    args=(directory, first_entered, first_release, None, True),
                ),
                context.Process(
                    target=hold_output_session,
                    args=(
                        directory,
                        second_entered,
                        second_release,
                        second_ready,
                        False,
                    ),
                ),
            ]
            try:
                processes[0].start()
                self.assertTrue(first_entered.wait(5), "first session did not acquire")
                processes[1].start()
                self.assertTrue(second_ready.wait(5), "second process did not attempt")
                self.assertFalse(
                    second_entered.wait(0.3),
                    "second session entered while first held lock",
                )
                first_release.set()
                self.assertTrue(
                    second_entered.wait(5), "exception did not release lock"
                )
                second_release.set()
                for process in processes:
                    process.join(5)
                    self.assertEqual(process.exitcode, 0)
                self.assertEqual(list(Path(directory).iterdir()), [])
            finally:
                first_release.set()
                second_release.set()
                for process in processes:
                    if process.pid is not None:
                        process.join(2)
                        if process.is_alive():
                            process.terminate()
                            process.join()


def hold_output_session(directory, entered, release, ready, fail):
    root = Path(directory)
    if ready is not None:
        ready.set()
    try:
        with OutputHistory(root / "viagens.json", root / "gastos.xlsx"):
            entered.set()
            if not release.wait(10):
                raise RuntimeError("release signal missing")
            if fail:
                raise ValueError("exercise exceptional release")
    except ValueError:
        pass
