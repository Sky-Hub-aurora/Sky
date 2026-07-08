# -*- coding: utf-8 -*-
"""实训报告生成器网页端本地服务。

运行后访问 http://127.0.0.1:8765 ，在网页里拖入 TXT 即可生成 DOCX。
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import mimetypes
import re
import sys
import threading
import uuid
import warnings
import webbrowser
from http import HTTPStatus
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from pathlib import Path
from types import SimpleNamespace
from urllib.parse import parse_qs, urlparse

warnings.filterwarnings("ignore", category=DeprecationWarning)
import cgi


APP_DIR = Path(__file__).resolve().parent
HTML_PATH = APP_DIR / "实训报告生成器.html"
CORE_PATH = APP_DIR / "实训报告生成器.py"
CACHE_DIR = APP_DIR / ".report_web_cache"
UPLOAD_DIR = CACHE_DIR / "uploads"
OUTPUT_DIR = CACHE_DIR / "outputs"

JOBS: dict[str, tuple[Path, str]] = {}
CORE = None


def main() -> int:
    parser = argparse.ArgumentParser(description="启动实训报告生成器本地网页服务。")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-browser", action="store_true", help="只启动服务，不自动打开浏览器。")
    args = parser.parse_args()

    ensure_runtime()
    server = ThreadingHTTPServer((args.host, args.port), ReportWebHandler)
    url = f"http://{args.host}:{args.port}"

    safe_log("实训报告生成器网页端已启动")
    safe_log(f"访问地址：{url}")
    safe_log("按 Ctrl+C 结束服务")

    if not args.no_browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        safe_log("\n服务已结束")
    finally:
        server.server_close()
    return 0


def ensure_runtime() -> None:
    for folder in (UPLOAD_DIR, OUTPUT_DIR):
        folder.mkdir(parents=True, exist_ok=True)
    load_core()


def load_core():
    global CORE
    if CORE is not None:
        return CORE
    if not CORE_PATH.exists():
        raise FileNotFoundError(f"没有找到核心脚本：{CORE_PATH}")
    spec = importlib.util.spec_from_file_location("report_generator_core", CORE_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError("无法加载核心脚本。")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    CORE = module
    return module


class ReportWebHandler(BaseHTTPRequestHandler):
    server_version = "ReportGeneratorWeb/1.0"

    def do_OPTIONS(self) -> None:
        self.send_response(HTTPStatus.NO_CONTENT)
        self.send_cors_headers()
        self.end_headers()

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path in {"/", "/index.html", "/实训报告生成器.html"}:
            self.send_file(HTML_PATH, "text/html; charset=utf-8")
            return
        if parsed.path == "/download":
            self.handle_download(parsed.query)
            return
        if parsed.path == "/health":
            self.send_json({"ok": True})
            return
        self.send_error(HTTPStatus.NOT_FOUND, "找不到页面")

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/generate":
            self.handle_generate()
            return
        self.send_error(HTTPStatus.NOT_FOUND, "找不到接口")

    def handle_generate(self) -> None:
        try:
            content_type = self.headers.get("Content-Type", "")
            if "multipart/form-data" not in content_type:
                raise ValueError("请使用网页表单上传 TXT 文件。")

            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={
                    "REQUEST_METHOD": "POST",
                    "CONTENT_TYPE": content_type,
                },
            )
            uploaded_paths = save_uploaded_txt_files(form)
            output_name = build_output_name(form)
            output_path = unique_path(OUTPUT_DIR / output_name)
            day_count = generate_report(form, uploaded_paths, output_path)

            job_id = uuid.uuid4().hex
            JOBS[job_id] = (output_path, output_path.name)
            self.send_json(
                {
                    "ok": True,
                    "filename": output_path.name,
                    "downloadUrl": f"/download?id={job_id}",
                    "message": f"已根据 {len(uploaded_paths)} 个 TXT 整合为 {day_count} 天并生成报告。",
                }
            )
        except Exception as exc:  # noqa: BLE001 - 本地网页需要把错误反馈给前端
            self.send_json({"ok": False, "error": str(exc)}, status=HTTPStatus.BAD_REQUEST)

    def handle_download(self, query: str) -> None:
        job_id = parse_qs(query).get("id", [""])[0]
        if not job_id or job_id not in JOBS:
            self.send_error(HTTPStatus.NOT_FOUND, "下载文件不存在或服务已重启")
            return
        file_path, filename = JOBS[job_id]
        if not file_path.exists():
            self.send_error(HTTPStatus.NOT_FOUND, "下载文件已不存在")
            return

        data = file_path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_cors_headers()
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{url_quote(filename)}")
        self.end_headers()
        self.wfile.write(data)

    def send_file(self, path: Path, content_type: str | None = None) -> None:
        if not path.exists():
            self.send_error(HTTPStatus.NOT_FOUND, f"缺少文件：{path.name}")
            return
        data = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_cors_headers()
        self.send_header("Content-Type", content_type or mimetypes.guess_type(path.name)[0] or "application/octet-stream")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def send_json(self, payload: dict, status: HTTPStatus = HTTPStatus.OK) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_cors_headers()
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def send_cors_headers(self) -> None:
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def log_message(self, format: str, *args) -> None:  # noqa: A002 - 父类接口名
        safe_log("%s - %s" % (self.address_string(), format % args))


def save_uploaded_txt_files(form: cgi.FieldStorage) -> list[Path]:
    file_fields = form["notes"] if "notes" in form else []
    if not isinstance(file_fields, list):
        file_fields = [file_fields]

    job_folder = UPLOAD_DIR / uuid.uuid4().hex
    job_folder.mkdir(parents=True, exist_ok=True)
    saved: list[Path] = []
    for index, field in enumerate(file_fields, start=1):
        if not getattr(field, "filename", ""):
            continue
        original_name = safe_filename(field.filename)
        if Path(original_name).suffix.lower() != ".txt":
            raise ValueError(f"只能上传 TXT 文件：{original_name}")
        target = unique_path(job_folder / original_name)
        data = field.file.read()
        if not data:
            raise ValueError(f"TXT 文件内容为空：{original_name}")
        target.write_bytes(data)
        saved.append(target)

    saved.sort(key=natural_sort_key)
    if not saved:
        raise ValueError("请至少拖入一个 TXT 文本文件。")
    return saved


def generate_report(form: cgi.FieldStorage, sources: list[Path], output_path: Path) -> int:
    core = load_core()
    template_path = core.resolve_template(get_form_value(form, "templatePath") or None)
    doc = core.Document(str(template_path))

    if get_bool(form, "fillCover"):
        cover_args = SimpleNamespace(
            college=get_form_value(form, "college"),
            major_class=get_form_value(form, "majorClass"),
            class_name=get_form_value(form, "className"),
            group=get_form_value(form, "group"),
            student_id=get_form_value(form, "studentId"),
            name=get_form_value(form, "studentName"),
        )
        core.fill_cover(doc, cover_args)

    core.clear_template_body_after_cover(doc)
    core.ensure_report_styles(doc)

    max_tasks = int(get_form_value(form, "maxTasks") or "8")
    max_tasks = max(1, min(max_tasks, 20))
    ai_config = core.AIConfig(
        enabled=get_bool(form, "aiEnabled"),
        base_url=get_form_value(form, "aiBaseUrl"),
        api_key=get_form_value(form, "aiApiKey"),
        model=get_form_value(form, "aiModel"),
        temperature=float(get_form_value(form, "aiTemperature") or "0.2"),
    )
    source_groups = core.group_sources_by_day(sources)
    for index, group in enumerate(source_groups, start=1):
        raw_text = core.read_source_group_text(group)
        report = core.build_day_report(
            source=group.paths[0],
            day_index=index,
            raw_text=raw_text,
            title_override=None,
            max_tasks=max_tasks,
            ai_config=ai_config,
        )
        if index > 1 and get_bool(form, "dayPageBreak"):
            doc.add_page_break()
        core.append_day_report(doc, report, screenshots_dir=None)

    doc.save(str(output_path))
    return len(source_groups)


def build_output_name(form: cgi.FieldStorage) -> str:
    explicit = get_form_value(form, "outputName")
    if explicit:
        name = safe_filename(explicit)
        return name if name.lower().endswith(".docx") else f"{name}.docx"

    class_name = get_form_value(form, "className")
    group = get_form_value(form, "group")
    student_id = get_form_value(form, "studentId")
    student_name = get_form_value(form, "studentName")
    if class_name and group and student_id and student_name:
        group_text = group if group.endswith("组") else f"{group}组"
        return safe_filename(f"{class_name}-{group_text}-{student_id}-{student_name}.docx")
    return f"实训报告_{uuid.uuid4().hex[:8]}.docx"


def get_form_value(form: cgi.FieldStorage, key: str) -> str:
    if key not in form:
        return ""
    value = form.getfirst(key, "")
    return str(value).strip()


def get_bool(form: cgi.FieldStorage, key: str) -> bool:
    return get_form_value(form, key).lower() in {"1", "true", "yes", "on"}


def safe_filename(name: str) -> str:
    name = Path(name).name.strip()
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    name = re.sub(r"\s+", " ", name)
    return name or "实训报告.docx"


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    for index in range(1, 1000):
        candidate = path.with_name(f"{path.stem}_{index}{path.suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"无法生成唯一输出文件名：{path.name}")


def natural_sort_key(path: Path) -> list[object]:
    parts = re.split(r"(\d+)", path.name)
    return [int(part) if part.isdigit() else part.lower() for part in parts]


def url_quote(text: str) -> str:
    safe = []
    for byte in text.encode("utf-8"):
        char = chr(byte)
        if char.isalnum() or char in ".-_":
            safe.append(char)
        else:
            safe.append(f"%{byte:02X}")
    return "".join(safe)


def safe_log(message: str) -> None:
    try:
        print(message)
    except Exception:
        pass


if __name__ == "__main__":
    raise SystemExit(main())
