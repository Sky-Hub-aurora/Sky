# -*- coding: utf-8 -*-
"""实训报告生成器网页端本地服务。

运行后访问 http://127.0.0.1:8765 ，在网页里拖入 TXT 即可生成 DOCX。
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import mimetypes
import os
import re
import sys
import threading
import uuid
import webbrowser
from dataclasses import dataclass, field
from email import policy
from email.parser import BytesParser
from http import HTTPStatus
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from pathlib import Path
from types import SimpleNamespace
from urllib import error as url_error
from urllib import request as url_request
from urllib.parse import parse_qs, urlparse


APP_DIR = Path(__file__).resolve().parent
HTML_PATH = APP_DIR / "实训报告生成器.html"
CORE_PATH = APP_DIR / "实训报告生成器.py"
CACHE_DIR = APP_DIR / ".report_web_cache"
UPLOAD_DIR = CACHE_DIR / "uploads"
OUTPUT_DIR = CACHE_DIR / "outputs"
ENV_PATH = APP_DIR / ".env"
AI_ENV_KEYS = [
    "AI_PROVIDER",
    "AI_BASE_URL",
    "AI_API_KEY",
    "AI_MODEL",
    "AI_TEMPERATURE",
    "AI_TIMEOUT",
    "AI_PROXY_URL",
]
SKY_AGENT_ENV_KEYS = [
    "SKY_AGENT_BASE_URL",
    "SKY_AGENT_API_KEY",
    "SKY_AGENT_MODEL",
    "SKY_AGENT_TEMPERATURE",
    "SKY_AGENT_TIMEOUT",
    "SKY_AGENT_PROXY_URL",
]
AI_PROVIDER_PRESETS = [
    {"id": "deepseek", "name": "DeepSeek", "baseUrl": "https://api.deepseek.com", "model": "deepseek-chat"},
    {"id": "openrouter", "name": "OpenRouter", "baseUrl": "https://openrouter.ai/api/v1", "model": "openai/gpt-4o-mini"},
    {"id": "kimi", "name": "Kimi", "baseUrl": "https://api.moonshot.cn/v1", "model": "moonshot-v1-8k"},
    {"id": "qwen", "name": "通义千问", "baseUrl": "https://dashscope.aliyuncs.com/compatible-mode/v1", "model": "qwen-plus"},
    {"id": "zhipu", "name": "智谱清言", "baseUrl": "https://open.bigmodel.cn/api/paas/v4", "model": "glm-4-flash"},
    {"id": "siliconflow", "name": "硅基流动", "baseUrl": "https://api.siliconflow.cn/v1", "model": "Qwen/Qwen2.5-7B-Instruct"},
    {"id": "doubao", "name": "豆包/火山方舟", "baseUrl": "https://ark.cn-beijing.volces.com/api/v3", "model": "doubao-seed-1-6-flash-250615"},
    {"id": "minimax", "name": "MiniMax", "baseUrl": "https://api.minimax.chat/v1", "model": "MiniMax-Text-01"},
    {"id": "yi", "name": "零一万物", "baseUrl": "https://api.lingyiwanwu.com/v1", "model": "yi-lightning"},
]

JOBS: dict[str, tuple[Path, str]] = {}
CORE = None


@dataclass
class UploadedFile:
    filename: str
    data: bytes


@dataclass
class MultipartForm:
    fields: dict[str, list[str]] = field(default_factory=dict)
    files: dict[str, list[UploadedFile]] = field(default_factory=dict)

    def add_field(self, name: str, value: str) -> None:
        self.fields.setdefault(name, []).append(value)

    def add_file(self, name: str, uploaded_file: UploadedFile) -> None:
        self.files.setdefault(name, []).append(uploaded_file)

    def getfirst(self, name: str, default: str = "") -> str:
        values = self.fields.get(name)
        return values[0] if values else default

    def get_files(self, name: str) -> list[UploadedFile]:
        return self.files.get(name, [])


def main() -> int:
    parser = argparse.ArgumentParser(description="启动实训报告生成器本地网页服务。")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-browser", action="store_true", help="只启动服务，不自动打开浏览器。")
    args = parser.parse_args()

    url = f"http://{args.host}:{args.port}"
    try:
        ensure_runtime()
        server = ThreadingHTTPServer((args.host, args.port), ReportWebHandler)
    except ModuleNotFoundError as exc:
        missing_name = exc.name or str(exc)
        safe_log("启动失败：当前 Python 缺少依赖模块。")
        safe_log(f"当前 Python：{sys.executable}")
        safe_log(f"缺少模块：{missing_name}")
        safe_log("请运行：python -m pip install -r requirements.txt")
        safe_log("或者直接双击“启动实训报告网页.bat”，让它自动安装依赖。")
        return 2
    except FileNotFoundError as exc:
        safe_log(f"启动失败：{exc}")
        safe_log("请确认项目文件没有缺失，尤其是 实训报告生成器.html 和 实训报告生成器.py。")
        return 2
    except OSError as exc:
        if is_port_in_use_error(exc):
            safe_log(f"端口 {args.port} 已被占用。")
            if is_existing_report_service(url):
                safe_log("检测到已有实训报告生成器服务正在运行，将直接打开现有页面。")
                if not args.no_browser:
                    webbrowser.open(url)
                return 0
            safe_log("解决方法：")
            safe_log("1. 关闭之前打开的网页服务黑色窗口后重新双击 BAT。")
            safe_log(f"2. 或者换一个端口启动：python start_report_web.py --port {args.port + 1}")
            return 2
        safe_log(f"启动失败：{exc}")
        return 2

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


def is_port_in_use_error(exc: OSError) -> bool:
    return getattr(exc, "winerror", None) == 10048 or exc.errno in {48, 98}


def is_existing_report_service(url: str) -> bool:
    try:
        with url_request.urlopen(f"{url}/health", timeout=1.5) as response:
            payload = json.loads(response.read().decode("utf-8", errors="replace") or "{}")
        return bool(payload.get("ok"))
    except (OSError, ValueError, url_error.URLError, json.JSONDecodeError):
        return False


def ensure_runtime() -> None:
    for folder in (UPLOAD_DIR, OUTPUT_DIR):
        folder.mkdir(parents=True, exist_ok=True)
    load_env_file()
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


def load_env_file() -> None:
    if not ENV_PATH.exists():
        return
    for raw_line in ENV_PATH.read_text(encoding="utf-8", errors="replace").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key:
            os.environ[key] = value


def read_env_values() -> dict[str, str]:
    values: dict[str, str] = {}
    if not ENV_PATH.exists():
        return values
    for raw_line in ENV_PATH.read_text(encoding="utf-8", errors="replace").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        values[key.strip()] = value.strip().strip('"').strip("'")
    return values


def write_ai_env_values(updates: dict[str, str], clear_api_key: bool = False) -> None:
    current = read_env_values()
    for key, value in updates.items():
        if key in AI_ENV_KEYS and value is not None:
            current[key] = str(value).strip()
    if clear_api_key:
        current["AI_API_KEY"] = ""

    lines: list[str] = [
        "# 实训报告生成器 AI 配置",
        "# 这些值可由网页自动写入；AI_API_KEY 已加入 .gitignore，不要上传真实密钥。",
    ]
    for key in AI_ENV_KEYS:
        if key in current:
            lines.append(f"{key}={current.get(key, '')}")
    for key, value in current.items():
        if key not in AI_ENV_KEYS:
            lines.append(f"{key}={value}")
    ENV_PATH.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")

    for key in AI_ENV_KEYS:
        value = current.get(key, "")
        if key == "AI_API_KEY" and clear_api_key:
            os.environ.pop(key, None)
        else:
            os.environ[key] = value


def mask_secret(value: str) -> str:
    value = (value or "").strip()
    if not value:
        return ""
    if len(value) <= 8:
        return "*" * len(value)
    return f"{value[:4]}{'*' * max(4, len(value) - 8)}{value[-4:]}"


def validate_ai_payload(values: dict | None, require_api_key: bool = False) -> list[str]:
    values = values or {}
    errors: list[str] = []
    base_url = pick_config_value(values, "baseUrl", "aiBaseUrl")
    api_key = pick_config_value(values, "apiKey", "aiApiKey")
    model = pick_config_value(values, "model", "aiModel")
    temperature = pick_config_value(values, "temperature", "aiTemperature")
    proxy_url = pick_config_value(values, "proxyUrl", "aiProxyUrl")

    if base_url:
        parsed = urlparse(base_url)
        if parsed.scheme not in {"http", "https"} or not parsed.netloc:
            errors.append("AI API 地址格式不正确，应类似：https://api.deepseek.com")
    if model and len(model) < 2:
        errors.append("模型名称太短，请填写接口文档里的真实模型名。")
    if temperature:
        temp = parse_float(temperature, -1.0)
        if temp < 0 or temp > 2:
            errors.append("温度建议填写 0 到 2 之间的数字。")
    if proxy_url:
        parsed_proxy = urlparse(proxy_url)
        if parsed_proxy.scheme not in {"http", "https"} or not parsed_proxy.netloc:
            errors.append("代理地址格式不正确，应类似：http://127.0.0.1:7890。当前工具使用 Python 标准库，只支持 HTTP/HTTPS 代理。")

    if require_api_key or api_key:
        if not api_key:
            errors.append("请填写 API Key，或者先保存一份可用的 .env 配置。")
        elif re.search(r"\s", api_key):
            errors.append("API Key 里不能包含空格、换行或制表符。")
        elif api_key.lower().startswith(("http://", "https://")):
            errors.append("API Key 输入框里不能填写网址，网址请填到 AI API 地址。")
        elif "bearer " in api_key.lower():
            errors.append("API Key 不要带 Bearer 前缀，只粘贴密钥本身。")
        elif len(api_key) < 12:
            errors.append("API Key 看起来太短，请检查是否复制完整。")
        elif any(token in api_key.lower() for token in ("你的", "api密钥", "apikey", "api_key", "your-key", "your_api_key")):
            errors.append("API Key 还是示例占位内容，请换成真实密钥。")
    return errors


def save_ai_config_from_values(values: dict | None, require_api_key: bool = False) -> str:
    values = values or {}
    errors = validate_ai_payload(values, require_api_key=require_api_key)
    if errors:
        raise ValueError("；".join(errors))
    provider_value = pick_config_value(values, "provider", "aiProvider")
    base_url_value = pick_config_value(values, "baseUrl", "aiBaseUrl")
    model_value = normalize_provider_model(provider_value, base_url_value, pick_config_value(values, "model", "aiModel"))
    updates = {
        "AI_PROVIDER": provider_value,
        "AI_BASE_URL": base_url_value,
        "AI_API_KEY": pick_config_value(values, "apiKey", "aiApiKey"),
        "AI_MODEL": model_value,
        "AI_TEMPERATURE": pick_config_value(values, "temperature", "aiTemperature") or "0.6",
        "AI_TIMEOUT": pick_config_value(values, "timeout", "aiTimeout") or "60",
        "AI_PROXY_URL": pick_config_value(values, "proxyUrl", "aiProxyUrl"),
    }
    updates = {key: value for key, value in updates.items() if value != ""}
    if updates:
        write_ai_env_values(updates)
    saved_values = read_env_values()
    return mask_secret(updates.get("AI_API_KEY") or saved_values.get("AI_API_KEY") or os.getenv("AI_API_KEY", ""))


def default_model_for_provider(provider: str, base_url: str = "") -> str:
    provider = (provider or "").strip().lower()
    base_url = (base_url or "").strip().lower()
    for preset in AI_PROVIDER_PRESETS:
        if provider == preset["id"] or preset["baseUrl"].lower() in base_url:
            return preset["model"]
    return ""


def normalize_provider_model(provider: str, base_url: str, model: str) -> str:
    model = (model or "").strip()
    provider = (provider or "").strip().lower()
    base_url = (base_url or "").strip().lower()
    if provider == "deepseek" or "api.deepseek.com" in base_url:
        if not model or model in {"deepseek-v4-flash", "deepseek-v3", "deepseek-v3.1"}:
            return "deepseek-chat"
    return model or default_model_for_provider(provider, base_url)


def pick_config_value(values: dict | None, *names: str) -> str:
    if not values:
        return ""
    for name in names:
        value = values.get(name)
        if value is not None and str(value).strip():
            return str(value).strip()
    return ""


def parse_float(value: str, default: float) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def parse_int(value: str, default: int) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def build_request_ai_config(core, enabled: bool, values: dict | None = None):
    load_env_file()
    provider = pick_config_value(values, "provider", "aiProvider") or os.getenv("AI_PROVIDER", "").strip()
    base_url = pick_config_value(values, "baseUrl", "aiBaseUrl") or os.getenv("AI_BASE_URL", "").strip()
    api_key = pick_config_value(values, "apiKey", "aiApiKey") or os.getenv("AI_API_KEY", "").strip()
    model = pick_config_value(values, "model", "aiModel") or os.getenv("AI_MODEL", "").strip()
    model = normalize_provider_model(provider, base_url, model)
    temperature = pick_config_value(values, "temperature", "aiTemperature") or os.getenv("AI_TEMPERATURE", "0.4").strip()
    timeout = pick_config_value(values, "timeout", "aiTimeout") or os.getenv("AI_TIMEOUT", "60").strip()
    proxy_url = pick_config_value(values, "proxyUrl", "aiProxyUrl") or os.getenv("AI_PROXY_URL", "").strip()
    return core.AIConfig(
        enabled=enabled,
        base_url=base_url,
        api_key=api_key,
        model=model,
        temperature=parse_float(temperature, 0.4),
        timeout=parse_int(timeout, 60),
        proxy_url=proxy_url,
    )


def build_env_ai_config(core, enabled: bool):
    return build_request_ai_config(core, enabled, None)


def build_sky_agent_config(core):
    load_env_file()
    base_url = os.getenv("SKY_AGENT_BASE_URL", "").strip() or "https://api.deepseek.com"
    api_key = os.getenv("SKY_AGENT_API_KEY", "").strip()
    model = os.getenv("SKY_AGENT_MODEL", "").strip() or "deepseek-chat"
    temperature = os.getenv("SKY_AGENT_TEMPERATURE", "0.5").strip()
    timeout = os.getenv("SKY_AGENT_TIMEOUT", "60").strip()
    proxy_url = os.getenv("SKY_AGENT_PROXY_URL", "").strip() or os.getenv("AI_PROXY_URL", "").strip()
    return core.AIConfig(
        enabled=True,
        base_url=base_url,
        api_key=api_key,
        model=normalize_provider_model("deepseek", base_url, model),
        temperature=parse_float(temperature, 0.5),
        timeout=parse_int(timeout, 60),
        proxy_url=proxy_url,
    )


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
        if parsed.path == "/ai-env-status":
            self.handle_ai_env_status()
            return
        if parsed.path == "/detect-ai-proxy":
            self.handle_detect_ai_proxy()
            return
        self.send_error(HTTPStatus.NOT_FOUND, "Not found")

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/generate":
            self.handle_generate()
            return
        if parsed.path == "/test-ai":
            self.handle_test_ai()
            return
        if parsed.path == "/api/chat":
            self.handle_api_chat()
            return
        if parsed.path == "/api/sky-agent":
            self.handle_sky_agent()
            return
        if parsed.path == "/save-ai-config":
            self.handle_save_ai_config()
            return
        if parsed.path == "/clear-ai-key":
            self.handle_clear_ai_key()
            return
        if parsed.path == "/probe-ai-routes":
            self.handle_probe_ai_routes()
            return
        self.send_error(HTTPStatus.NOT_FOUND, "Not found")

    def handle_generate(self) -> None:
        try:
            content_type = self.headers.get("Content-Type", "")
            if "multipart/form-data" not in content_type:
                raise ValueError("请使用网页表单上传 TXT 文件。")

            content_length = int(self.headers.get("Content-Length", "0") or "0")
            form = parse_multipart_form(content_type, self.rfile.read(content_length))
            uploaded_paths = save_uploaded_txt_files(form)
            output_name = build_output_name(form)
            output_path = unique_path(OUTPUT_DIR / output_name)
            xlsx_path = unique_path(OUTPUT_DIR / "对照表.xlsx")
            day_count, warnings, ai_status = generate_report(form, uploaded_paths, output_path, xlsx_path)
            message = f"已根据 {len(uploaded_paths)} 个 TXT 整合为 {day_count} 天并生成 Word 和 xlsx 对照表。"
            if warnings:
                message += " AI 增强未成功，已自动改用本地规则生成。"
            elif ai_status["requested"] and ai_status["used"]:
                message += f" AI 增强已真实调用模型：{ai_status['model']}。"
            elif not ai_status["requested"]:
                message += " 本次未开启 AI 增强，已使用本地反重复规则生成。"

            job_id = uuid.uuid4().hex
            xlsx_job_id = uuid.uuid4().hex
            JOBS[job_id] = (output_path, output_path.name)
            JOBS[xlsx_job_id] = (xlsx_path, xlsx_path.name)
            self.send_json(
                {
                    "ok": True,
                    "filename": output_path.name,
                    "downloadUrl": f"/download?id={job_id}",
                    "xlsxFilename": xlsx_path.name,
                    "xlsxDownloadUrl": f"/download?id={xlsx_job_id}",
                    "message": message,
                    "warnings": warnings,
                    "aiStatus": ai_status,
                }
            )
        except Exception as exc:  # noqa: BLE001 - 本地网页需要把错误反馈给前端
            self.send_json({"ok": False, "error": str(exc)}, status=HTTPStatus.BAD_REQUEST)

    def handle_test_ai(self) -> None:
        try:
            content_length = int(self.headers.get("Content-Length", "0") or "0")
            payload = json.loads(self.rfile.read(content_length).decode("utf-8") or "{}")
            saved_key = save_ai_config_from_values(payload, require_api_key=False)
            core = load_core()
            ai_config = build_request_ai_config(core, enabled=True, values=payload)
            if not ai_config.api_key:
                raise ValueError("请先填写 API Key，或先保存一份可用的 .env 密钥。")
            message = core.test_ai_connection(ai_config)
            self.send_json(
                {
                    "ok": True,
                    "message": message,
                    "model": ai_config.model,
                    "baseUrl": ai_config.base_url,
                    "savedKey": saved_key,
                    "pythonPath": sys.executable,
                }
            )
        except Exception as exc:  # noqa: BLE001 - 测试按钮需要把具体连接错误回显给前端
            self.send_json(
                {
                    "ok": False,
                    "error": str(exc),
                    "pythonPath": sys.executable,
                },
                status=HTTPStatus.BAD_REQUEST,
            )

    def handle_api_chat(self) -> None:
        try:
            content_length = int(self.headers.get("Content-Length", "0") or "0")
            payload = json.loads(self.rfile.read(content_length).decode("utf-8") or "{}")
            message = str(payload.get("message") or "").strip()
            if not message:
                raise ValueError("请输入要发送给 AI 的内容。")
            saved_key = save_ai_config_from_values(payload, require_api_key=False)
            core = load_core()
            ai_config = build_request_ai_config(core, enabled=True, values=payload)
            if not ai_config.api_key:
                raise ValueError("请先填写 API Key，或先保存一份可用的 .env 密钥。")
            answer = core.call_openai_compatible_text(
                ai_config,
                message,
                system_prompt="你是实训报告生成器的 AI 助手。请根据用户输入给出可用于实训报告优化的中文建议，回答要具体、自然、可操作。",
            )
            self.send_json(
                {
                    "ok": True,
                    "answer": answer,
                    "model": ai_config.model,
                    "baseUrl": ai_config.base_url,
                    "savedKey": saved_key,
                    "pythonPath": sys.executable,
                }
            )
        except Exception as exc:  # noqa: BLE001 - 前端需要展示 API 中转错误
            self.send_json(
                {
                    "ok": False,
                    "error": str(exc),
                    "pythonPath": sys.executable,
                },
                status=HTTPStatus.BAD_REQUEST,
            )

    def handle_sky_agent(self) -> None:
        try:
            content_length = int(self.headers.get("Content-Length", "0") or "0")
            payload = json.loads(self.rfile.read(content_length).decode("utf-8") or "{}")
            message = str(payload.get("message") or "").strip()
            if not message:
                raise ValueError("请输入要发送给 Sky Agent 的内容。")

            core = load_core()
            ai_config = build_sky_agent_config(core)
            if not ai_config.api_key:
                raise ValueError("Sky Agent API Key 未配置。请把 SKY_AGENT_API_KEY 写入本地 .env。")

            history = payload.get("history") if isinstance(payload.get("history"), list) else []
            context = payload.get("context") if isinstance(payload.get("context"), dict) else {}
            agent_prompt = json.dumps(
                {
                    "user_message": message,
                    "recent_history": history[-8:],
                    "current_report_context": context,
                },
                ensure_ascii=False,
                indent=2,
            )
            raw_answer = core.call_openai_compatible_text(
                ai_config,
                agent_prompt,
                system_prompt=(
                    "你是右下角的 Sky Agent，是实训报告生成器里的交互式助手。"
                    "你需要用简洁自然的中文和用户聊天，并把用户关于报告生成方式的想法整理成可执行的生成控制指令。"
                    "只输出 JSON，不要 Markdown。JSON 格式必须是："
                    '{"reply":"给用户看的回复","directive":"写入报告生成流程的控制指令","apply":true}。'
                    "directive 要能直接影响 Word 报告生成，例如写作风格、内容侧重点、代码说明深度、避免事项、和原有 AI 生成器对话时要遵守的要求。"
                    "如果用户只是普通聊天，也要给出简短 reply，directive 可以为空字符串，apply 为 false。"
                ),
            )
            try:
                parsed = core.extract_json_object(raw_answer)
                answer = str(parsed.get("reply") or raw_answer).strip()
                directive = str(parsed.get("directive") or "").strip()
                apply_directive = bool(parsed.get("apply", bool(directive)))
            except Exception:
                answer = raw_answer.strip()
                directive = message
                apply_directive = True

            self.send_json(
                {
                    "ok": True,
                    "answer": answer,
                    "directive": directive,
                    "apply": apply_directive,
                    "model": ai_config.model,
                    "baseUrl": ai_config.base_url,
                    "pythonPath": sys.executable,
                }
            )
        except Exception as exc:  # noqa: BLE001 - Sky Agent 需要把错误反馈给前端
            self.send_json(
                {
                    "ok": False,
                    "error": str(exc),
                    "pythonPath": sys.executable,
                },
                status=HTTPStatus.BAD_REQUEST,
            )

    def handle_save_ai_config(self) -> None:
        try:
            content_length = int(self.headers.get("Content-Length", "0") or "0")
            payload = json.loads(self.rfile.read(content_length).decode("utf-8") or "{}")
            existing_key = read_env_values().get("AI_API_KEY") or os.getenv("AI_API_KEY", "")
            typed_key = pick_config_value(payload, "apiKey", "aiApiKey")
            saved_key = save_ai_config_from_values(payload, require_api_key=not bool(existing_key or typed_key))
            self.send_json(
                {
                    "ok": True,
                    "message": "AI 配置已保存到 .env。",
                    "savedKey": saved_key,
                    "envPath": str(ENV_PATH),
                }
            )
        except Exception as exc:  # noqa: BLE001 - 前端要展示校验错误
            self.send_json({"ok": False, "error": str(exc)}, status=HTTPStatus.BAD_REQUEST)

    def handle_clear_ai_key(self) -> None:
        try:
            write_ai_env_values({}, clear_api_key=True)
            self.send_json({"ok": True, "message": "已清除网页输入和 .env 中的 AI_API_KEY。"})
        except Exception as exc:  # noqa: BLE001 - 前端要展示清除错误
            self.send_json({"ok": False, "error": str(exc)}, status=HTTPStatus.BAD_REQUEST)

    def handle_ai_env_status(self) -> None:
        load_env_file()
        values = read_env_values()
        api_key = values.get("AI_API_KEY") or os.getenv("AI_API_KEY", "")
        provider = values.get("AI_PROVIDER", os.getenv("AI_PROVIDER", ""))
        base_url = values.get("AI_BASE_URL", os.getenv("AI_BASE_URL", ""))
        model = normalize_provider_model(provider, base_url, values.get("AI_MODEL", os.getenv("AI_MODEL", "")))
        self.send_json(
            {
                "ok": True,
                "hasKey": bool(api_key),
                "maskedKey": mask_secret(api_key),
                "provider": provider,
                "baseUrl": base_url,
                "model": model,
                "temperature": values.get("AI_TEMPERATURE", os.getenv("AI_TEMPERATURE", "")),
                "proxyUrl": values.get("AI_PROXY_URL", os.getenv("AI_PROXY_URL", "")),
            }
        )

    def handle_detect_ai_proxy(self) -> None:
        try:
            core = load_core()
            routes = [{"label": label, "proxyUrl": proxy} for label, proxy in core.detect_proxy_routes()]
            first_proxy = routes[0]["proxyUrl"] if routes else ""
            self.send_json(
                {
                    "ok": True,
                    "routes": routes,
                    "firstProxy": first_proxy,
                    "message": "已检测到可用代理候选。" if routes else "未检测到系统代理或常见本地代理端口。",
                }
            )
        except Exception as exc:  # noqa: BLE001 - 前端需要展示检测错误
            self.send_json({"ok": False, "error": str(exc)}, status=HTTPStatus.BAD_REQUEST)

    def handle_probe_ai_routes(self) -> None:
        try:
            content_length = int(self.headers.get("Content-Length", "0") or "0")
            payload = json.loads(self.rfile.read(content_length).decode("utf-8") or "{}")
            proxy_url = str(payload.get("proxyUrl") or "").strip()
            presets = payload.get("presets") if isinstance(payload.get("presets"), list) else AI_PROVIDER_PRESETS
            core = load_core()
            results = []
            for preset in presets:
                if not isinstance(preset, dict):
                    continue
                name = str(preset.get("name") or preset.get("id") or "未知通道")
                base_url = str(preset.get("baseUrl") or "").strip()
                if not base_url:
                    results.append({"name": name, "ok": False, "message": "缺少 API 地址。"})
                    continue
                probe = core.probe_url_connectivity(base_url, proxy_url=proxy_url, timeout=8)
                results.append({"name": name, "baseUrl": base_url, **probe})
            self.send_json({"ok": True, "results": results})
        except Exception as exc:  # noqa: BLE001 - 前端需要展示检测错误
            self.send_json({"ok": False, "error": str(exc)}, status=HTTPStatus.BAD_REQUEST)

    def handle_download(self, query: str) -> None:
        job_id = parse_qs(query).get("id", [""])[0]
        if not job_id or job_id not in JOBS:
            self.send_error(HTTPStatus.NOT_FOUND, "Download not found")
            return
        file_path, filename = JOBS[job_id]
        if not file_path.exists():
            self.send_error(HTTPStatus.NOT_FOUND, "Download file missing")
            return

        data = file_path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_cors_headers()
        self.send_header("Content-Type", guess_download_content_type(file_path))
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{url_quote(filename)}")
        self.end_headers()
        self.wfile.write(data)

    def send_file(self, path: Path, content_type: str | None = None) -> None:
        if not path.exists():
            self.send_error(HTTPStatus.NOT_FOUND, "File missing")
            return
        data = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_cors_headers()
        self.send_header("Content-Type", content_type or mimetypes.guess_type(path.name)[0] or "application/octet-stream")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store, max-age=0")
        self.send_header("Pragma", "no-cache")
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


def parse_multipart_form(content_type: str, body: bytes) -> MultipartForm:
    if not body:
        raise ValueError("上传表单内容为空，请重新选择 TXT 文件。")

    message_bytes = (
        f"Content-Type: {content_type}\r\n"
        "MIME-Version: 1.0\r\n"
        "\r\n"
    ).encode("utf-8") + body
    message = BytesParser(policy=policy.default).parsebytes(message_bytes)
    if not message.is_multipart():
        raise ValueError("上传表单格式不正确，请刷新网页后重试。")

    form = MultipartForm()
    for part in message.iter_parts():
        if part.get_content_maintype() == "multipart":
            continue
        disposition = part.get("Content-Disposition", "")
        if "form-data" not in disposition:
            continue
        name = part.get_param("name", header="content-disposition")
        if not name:
            continue

        payload = part.get_payload(decode=True) or b""
        filename = part.get_filename()
        if filename:
            form.add_file(name, UploadedFile(filename=filename, data=payload))
            continue

        charset = part.get_content_charset() or "utf-8"
        form.add_field(name, payload.decode(charset, errors="replace").strip())
    return form


def save_uploaded_txt_files(form: MultipartForm) -> list[Path]:
    file_fields = form.get_files("notes")

    job_folder = UPLOAD_DIR / uuid.uuid4().hex
    job_folder.mkdir(parents=True, exist_ok=True)
    saved: list[Path] = []
    for index, field in enumerate(file_fields, start=1):
        if not field.filename:
            continue
        original_name = safe_filename(field.filename)
        if Path(original_name).suffix.lower() != ".txt":
            raise ValueError(f"只能上传 TXT 文件：{original_name}")
        target = unique_path(job_folder / original_name)
        data = field.data
        if not data:
            raise ValueError(f"TXT 文件内容为空：{original_name}")
        target.write_bytes(data)
        saved.append(target)

    saved.sort(key=natural_sort_key)
    if not saved:
        raise ValueError("请至少拖入一个 TXT 文本文件。")
    return saved


def generate_report(form: MultipartForm, sources: list[Path], output_path: Path, xlsx_path: Path) -> tuple[int, list[str], dict]:
    core = load_core()
    core.clear_generation_warnings()
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
    request_ai_values = {
        "aiProvider": get_form_value(form, "aiProvider"),
        "aiBaseUrl": get_form_value(form, "aiBaseUrl"),
        "aiApiKey": get_form_value(form, "aiApiKey"),
        "aiModel": get_form_value(form, "aiModel"),
        "aiTemperature": get_form_value(form, "aiTemperature"),
        "aiProxyUrl": get_form_value(form, "aiProxyUrl"),
    }
    if get_bool(form, "aiEnabled"):
        save_ai_config_from_values(request_ai_values, require_api_key=False)
    ai_config = build_request_ai_config(core, enabled=get_bool(form, "aiEnabled"), values=request_ai_values)
    personal_note = get_form_value(form, "personalNote")
    sky_agent_directive = get_form_value(form, "skyAgentDirective")
    if sky_agent_directive:
        personal_note = "\n".join(
            part for part in [
                personal_note.strip(),
                f"Sky Agent 生成控制指令：{sky_agent_directive.strip()}",
            ] if part
        )
    source_groups = core.group_sources_by_day(sources)
    variation_seed = f"{output_path.name}-{uuid.uuid4().hex}"
    comparison_rows = []
    for index, group in enumerate(source_groups, start=1):
        raw_text = core.read_source_group_text(group)
        report = core.build_day_report(
            source=group.paths[0],
            day_index=index,
            raw_text=raw_text,
            title_override=None,
            max_tasks=max_tasks,
            ai_config=ai_config,
            variation_seed=f"{variation_seed}-{group.label}-{index}",
            personal_note=personal_note,
        )
        comparison_rows.extend(core.build_comparison_rows(report, raw_text, group))
        if index > 1 and get_bool(form, "dayPageBreak"):
            doc.add_page_break()
        core.append_day_report(doc, report, screenshots_dir=None, auto_result_images=True)

    doc.save(str(output_path))
    core.write_comparison_xlsx(xlsx_path, comparison_rows)
    ai_events = core.get_ai_usage_events()
    ai_status = {
        "requested": ai_config.enabled,
        "used": bool(ai_events),
        "model": ai_config.model,
        "baseUrl": ai_config.base_url,
        "events": ai_events,
    }
    return len(source_groups), core.get_generation_warnings(), ai_status


def build_output_name(form: MultipartForm) -> str:
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


def get_form_value(form: MultipartForm, key: str) -> str:
    value = form.getfirst(key, "")
    return str(value).strip()


def get_bool(form: MultipartForm, key: str) -> bool:
    return get_form_value(form, key).lower() in {"1", "true", "yes", "on"}


def safe_filename(name: str) -> str:
    name = Path(name).name.strip()
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    name = re.sub(r"\s+", " ", name)
    return name or "实训报告.docx"


def guess_download_content_type(path: Path) -> str:
    if path.suffix.lower() == ".docx":
        return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    if path.suffix.lower() == ".xlsx":
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return mimetypes.guess_type(path.name)[0] or "application/octet-stream"


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
