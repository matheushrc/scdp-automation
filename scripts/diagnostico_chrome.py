"""Record Chrome profile copy failures without exposing profile paths or data."""

import json
import re
import runpy
import shutil
import sys
from pathlib import Path
from types import FrameType
from typing import Any


def main() -> None:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    errors: list[dict[str, Any]] = []

    def trace(frame: FrameType, event: str, arg: Any) -> Any:
        if (
            event == "exception"
            and frame.f_code.co_name == "copy_chrome_profile"
            and frame.f_globals.get("__name__")
            == "scdp_automation.chrome_profile_setup"
            and isinstance(arg[1], OSError)
        ):
            error = arg[1]
            record = {
                "tipo": type(error).__name__,
                "errno": error.errno,
                "winerror": getattr(error, "winerror", None),
                "linha": frame.f_lineno,
            }
            if isinstance(error, shutil.Error):
                record["falhas"] = [
                    {
                        "tamanho_caminho_origem": len(str(source)),
                        "tamanho_caminho_destino": len(str(destination)),
                        "codigos": re.findall(
                            r"\[(?:WinError|Errno) \d+\]", str(message)
                        ),
                    }
                    for source, destination, message in error.args[0]
                ]
            errors.append(record)
        return trace

    sys.argv = ["scdp_automation", "--abrir-navegador"]
    sys.settrace(trace)
    try:
        runpy.run_module("scdp_automation", run_name="__main__")
    finally:
        sys.settrace(None)
        Path("diagnostico-chrome.txt").write_text(
            json.dumps(errors, indent=2, ensure_ascii=True), encoding="utf-8"
        )


if __name__ == "__main__":
    main()
