# -*- coding: utf-8 -*-
"""ASCII launcher for the local report generator web service."""

from __future__ import annotations

import runpy
import sys
from pathlib import Path


REQUIRED_MODULES = [
    ("docx", "python-docx"),
    ("pptx", "python-pptx"),
    ("PIL", "Pillow"),
]


def configure_console() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def find_missing_modules() -> list[str]:
    missing: list[str] = []
    for module_name, package_name in REQUIRED_MODULES:
        try:
            __import__(module_name)
        except ModuleNotFoundError:
            missing.append(package_name)
    return missing


def print_dependency_help(app_dir: Path, missing: list[str]) -> None:
    requirements = app_dir / "requirements.txt"
    print("实训报告生成器启动失败：当前 Python 缺少依赖。")
    print(f"当前 Python：{sys.executable}")
    print("缺少模块：" + "、".join(missing))
    print()
    if requirements.exists():
        print("请在本文件夹打开命令行运行：")
        print(f'"{sys.executable}" -m pip install -r "{requirements}"')
    else:
        print("没有找到 requirements.txt，请手动安装：")
        print(f'"{sys.executable}" -m pip install ' + " ".join(missing))
    print()
    print("也可以直接双击：启动实训报告网页.bat")


def main() -> int:
    configure_console()
    app_dir = Path(__file__).resolve().parent
    missing = find_missing_modules()
    if missing:
        print_dependency_help(app_dir, missing)
        return 2

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
