import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from scdp_automation.logging_config import configure_logging


class LoggingConfigTests(unittest.TestCase):
    def test_configures_daily_loguru_sink_with_thirty_day_retention(self) -> None:
        with TemporaryDirectory() as directory:
            log_directory = Path(directory) / "logs" / "scdp"
            with (
                patch("scdp_automation.logging_config.logger.remove") as remove,
                patch("scdp_automation.logging_config.logger.add") as add,
            ):
                configure_logging(log_directory)

        remove.assert_called_once_with()
        self.assertEqual(add.call_count, 2)
        sink = add.call_args_list[1].args[0]
        options = add.call_args_list[1].kwargs
        self.assertEqual(sink, str(log_directory / "scdp_{time:YYYY-MM-DD}.log"))
        self.assertEqual(options["rotation"], "00:00")
        self.assertEqual(options["retention"], "30 days")
        self.assertEqual(options["encoding"], "utf-8")


if __name__ == "__main__":
    unittest.main()
