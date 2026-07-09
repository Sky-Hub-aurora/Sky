# -*- coding: utf-8 -*-
"""ASCII launcher for the local report generator web service."""

from __future__ import annotations

import runpy
import sys
from pathlib import Path


def main() -> int:
    app_dir = Path(__file__).resolve().parent
    target = app_dir / "实训报告生成器网页.py"
    if not target.exists():
        matches = sorted(app_dir.glob("*网页.py"))
        if matches:
            target = matches[0]
    if not target.exists():
        print("Cannot find the web service Python file.")
        print("Please keep start_report_web.py in the project folder.")
        return 1

    sys.argv = [str(target), *sys.argv[1:]]
    runpy.run_path(str(target), run_name="__main__")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
