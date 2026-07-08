
# -*- coding: utf-8 -*-
"""实训报告生成器

基于实训报告模板和按天整理的 TXT/PPTX/DOCX 笔记，自动生成后续课程报告正文。

常用示例：p
    python 实训报告生成器.py
    python 实训报告生成器.py 笔记1.txt 笔记2.txt 笔记3.txt --output 计科1班-2组-2023032175-折柯阳.docx
    python 实训报告生成器.py --screenshots 截图文件夹 --fill-cover --college 计算机科学与技术 --major-class 23060101 --student-id 2023032175 --name 折柯阳
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from textwrap import dedent
from typing import Iterable
from urllib import error as url_error
from urllib import request as url_request

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt


DEFAULT_TEMPLATE_CANDIDATES = [
    Path("实训报告（过程）模板.docx"),
    Path("实训报告（过程）模板(2).docx"),
    Path(r"C:\Users\SKY\Documents\xwechat_files\wxid_v6es2tk5aij412_a2db\msg\file\2026-07\实训报告（过程）模板(2).docx"),
    Path(r"C:\Users\SKY\Documents\xwechat_files\wxid_v6es2tk5aij412_a2db\msg\file\2026-07\实训报告（过程）模板.docx"),
]

DEFAULT_SOURCE_CANDIDATES = [
    Path(r"C:\Users\SKY\Desktop\学习笔记7-5-01.txt"),
    Path(r"C:\Users\SKY\Desktop\学习笔记7-6-01.txt"),
]

TEXT_ENCODINGS = ("utf-8-sig", "utf-8", "gb18030", "gbk", "cp936")
BODY_SIZE = 10.5
BODY_FIRST_LINE_INDENT = Pt(BODY_SIZE * 2)
CN_NUMS = "零一二三四五六七八九十"


@dataclass
class TaskItem:
    title: str
    requirement: str
    code: str
    caption: str


@dataclass
class DayReport:
    source: Path
    day_index: int
    title: str
    goals: list[str]
    topics: list[str]
    details: list[str]
    tasks: list[TaskItem]


@dataclass
class SourceGroup:
    label: str
    paths: list[Path]
    sort_key: tuple


@dataclass
class AIConfig:
    enabled: bool = False
    base_url: str = ""
    api_key: str = ""
    model: str = ""
    temperature: float = 0.2
    timeout: int = 60


TOPIC_RULES = [
    ("实训规划与注意事项", ("实训的规划", "实训规划", "注意事项")),
    ("Python编程语言与业务应用", ("python编程语言工具", "可以完成的业务", "业务之间的关系")),
    ("数据类型书写与元素访问", ("数据类型的书写", "元素的访问", "数据类型")),
    ("输入输出命令", ("输入/输出", "输入输出", "print", "input")),
    ("字符串大小写转换", ("upper()", "lower()", "大写", "小写")),
    ("字符串查找与统计", ("find(", "count(", "查找", "统计")),
    ("字符串替换", ("replace(", "替换")),
    ("字符串去空", ("strip(", "去掉字符串前后", "前后空格")),
    ("字符串拆分与组合", ("split(", "join(", "拆分", "连接成字符串")),
    ("字符串类型测试", ("isalnum", "isalpha", "isdigit", "isspace", "islower", "isupper", "istitle")),
    ("字符编码转换", ("ord(", "chr(", "编码转换")),
    ("运算符和表达式", ("运算符", "表达式")),
    ("选择结构", ("选择结构", "if", "elif", "else")),
    ("循环结构", ("循环结构", "for", "while", "九九乘法表", "猜数字")),
    ("break与continue", ("break", "continue")),
    ("模块导入", ("import", "from", "random", "导入python功能库")),
    ("函数定义与调用", ("函数定义", "def", "函数的调用")),
    ("函数返回值", ("return", "返回值")),
    ("函数参数传递", ("形参", "实参", "参数的传递")),
    ("变量作用域", ("全局变量", "局部变量", "变量的作用域")),
    ("文件读写命令", ("with", "open", "read()", "write()", "文件读写")),
    ("文件路径", ("绝对路径", "相对路径")),
]

DETAIL_BANK = {
    "实训规划与注意事项": "实训开始阶段需要明确课程安排、学习纪律、阶段任务、报告提交与答辩要求，为后续项目实践建立清晰的学习路线。",
    "Python编程语言与业务应用": "Python是本次实训的主要编程工具，可用于数据处理、自动化办公、数据可视化、机器学习、深度学习及项目原型开发。",
    "数据类型书写与元素访问": "Python常见数据类型包括字符串、数值、布尔值、列表、字典、元组和集合，序列类型可通过下标和切片访问元素，字典可通过键访问值。",
    "输入输出命令": "print()用于在运行结果栏输出信息，input()用于获取键盘输入，二者可结合f-string完成交互式程序设计。",
    "字符串大小写转换": "upper()和lower()用于字符串大小写转换，适合处理用户名、编码、英文文本统一格式等场景。",
    "字符串查找与统计": "find()可返回指定子串首次出现的位置，count()可统计子串出现次数，能够完成文本检索与简单统计分析。",
    "字符串替换": "replace()可把字符串中的指定内容替换为新内容，可用于文本纠错、统一术语和批量清洗字符。",
    "字符串去空": "strip()用于去除字符串首尾空白字符或指定字符，常用于清理用户输入、文件读取内容和网络文本。",
    "字符串拆分与组合": "split()可按分隔符把字符串拆分为列表，join()可把序列元素重新组合为字符串，是文本结构化处理的基础方法。",
    "字符串类型测试": "isalnum()、isalpha()、isdigit()等方法返回布尔值，可用于判断字符串是否满足用户名、编号或纯数字等格式要求。",
    "字符编码转换": "ord()可以把字符转换为编码值，chr()可以把编码值转换为字符，帮助理解字符在计算机中的存储表示。",
    "运算符和表达式": "表达式由数据、变量和运算符组成，常见运算符包括算术运算符、关系运算符和逻辑运算符，是程序判断与计算的基础。",
    "选择结构": "选择结构通过if、elif、else按照条件执行不同语句，可实现费用判断、成绩分类、权限判断等分支逻辑。",
    "循环结构": "循环结构通过for或while重复执行代码，双层循环可控制行列输出，适合九九乘法表、累加、遍历列表等任务。",
    "break与continue": "break用于提前结束循环，continue用于跳过本轮循环后继续下一轮，可实现更加灵活的循环控制。",
    "模块导入": "import和from...import用于导入标准库或第三方库，导入后即可调用库中的函数扩展程序功能。",
    "函数定义与调用": "函数通过def定义，将重复使用的功能封装为独立代码块，调用函数时传入实参即可执行对应功能。",
    "函数返回值": "return用于把函数处理结果返回给调用位置，若函数没有显式return，则默认返回None。",
    "函数参数传递": "参数传递的本质是把实参提供给形参，使函数能够处理不同输入，提高代码复用性。",
    "变量作用域": "全局变量定义在函数外，局部变量定义在函数内，理解作用域有助于避免变量覆盖和数据混乱。",
    "文件读写命令": "with open()可安全打开文件，read()、readline()、readlines()用于读取，write()用于写入文本内容。",
    "文件路径": "绝对路径从磁盘根目录开始定位文件，相对路径以当前程序运行目录为参照，二者都会影响文件能否正确打开。",
}

GOAL_BANK = {
    "实训规划与注意事项": "了解暑期实训的整体安排、学习纪律、报告提交和项目答辩要求。",
    "Python编程语言与业务应用": "认识Python编程语言在数据分析、人工智能和项目实践中的工具作用。",
    "数据类型书写与元素访问": "掌握Python常见数据类型的书写方式，并能够使用下标、切片或键访问元素。",
    "输入输出命令": "能够使用print()、input()和格式化输出完成基础交互程序。",
    "字符串大小写转换": "掌握字符串大小写转换方法，并能完成用户名等文本格式处理。",
    "字符串查找与统计": "能够使用find()和count()完成文本查找与次数统计。",
    "字符串替换": "能够使用replace()对文本中的指定字符或词语进行批量替换。",
    "字符串去空": "能够使用strip()清理字符串首尾空白字符。",
    "字符串拆分与组合": "能够使用split()和join()完成字符串与列表之间的转换。",
    "字符串类型测试": "能够使用字符串测试函数判断文本是否满足指定格式。",
    "字符编码转换": "理解ord()与chr()在字符和编码之间转换的作用。",
    "运算符和表达式": "理解运算符和表达式的基本构成，并能完成简单计算与判断。",
    "选择结构": "掌握if、elif、else选择结构的基本写法。",
    "循环结构": "掌握for、while循环结构，并能够编写九九乘法表和猜数字等程序。",
    "break与continue": "理解break和continue对循环执行流程的控制作用。",
    "模块导入": "掌握import和from...import导入模块的基本方法。",
    "函数定义与调用": "能够使用def定义函数，并通过实参调用函数解决重复性问题。",
    "函数返回值": "理解return返回值的作用，并能够获取函数处理结果。",
    "函数参数传递": "理解形参与实参的关系，能够编写带参数的函数。",
    "变量作用域": "区分全局变量和局部变量，避免函数内外变量使用错误。",
    "文件读写命令": "掌握with open()文件读写结构，并能够完成文本文件的创建、写入和读取。",
    "文件路径": "理解绝对路径和相对路径的区别，能够正确定位程序运行所需文件。",
}


def main() -> int:
    args = parse_args()
    try:
        template_path = resolve_template(args.template)
        source_paths = resolve_sources(args.sources)
        output_path = resolve_output_path(args)

        doc = Document(str(template_path))
        if args.fill_cover:
            fill_cover(doc, args)

        clear_template_body_after_cover(doc)
        ensure_report_styles(doc)

        source_groups = group_sources_by_day(source_paths)
        ai_config = build_ai_config_from_args(args)
        for idx, group in enumerate(source_groups, start=1):
            raw_text = read_source_group_text(group)
            title_override = args.day_title[idx - 1] if idx <= len(args.day_title) else None
            report = build_day_report(
                source=group.paths[0],
                day_index=idx,
                raw_text=raw_text,
                title_override=title_override,
                max_tasks=args.max_tasks_per_day,
                ai_config=ai_config,
            )
            if idx > 1 and args.day_page_break:
                doc.add_page_break()
            append_day_report(doc, report, screenshots_dir=Path(args.screenshots) if args.screenshots else None)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(output_path))

        print(f"已生成实训报告：{output_path}")
        print(f"使用模板：{template_path}")
        print("按天使用的资料：")
        for day_index, group in enumerate(source_groups, start=1):
            print(f"  第{day_index}天（{group.label}）：")
            for source in group.paths:
                print(f"    - {source}")
        return 0
    except Exception as exc:  # noqa: BLE001 - 命令行工具需要给出清楚错误
        print(f"生成失败：{exc}", file=sys.stderr)
        return 1


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="按实训报告模板生成课程目标、课程内容、内容详情、代码及运行结果截图占位。",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("sources", nargs="*", help="按天顺序传入 TXT/PPTX/DOCX 笔记文件。为空时使用脚本内置的两个样例 TXT。")
    parser.add_argument("--template", help="实训报告模板 DOCX 路径。为空时自动寻找同目录模板或本机默认模板。")
    parser.add_argument("-o", "--output", help="输出 DOCX 路径。为空时自动生成文件名。")
    parser.add_argument("--screenshots", help="运行结果截图文件夹；自动匹配 1-1*.png、2-3*.jpg 等文件并插入。")
    parser.add_argument("--day-title", action="append", default=[], help="覆盖某一天标题，可重复传入，顺序对应资料文件。")
    parser.add_argument("--day-page-break", action="store_true", help="每天之间强制分页。")
    parser.add_argument("--max-tasks-per-day", type=int, default=8, help="每天最多写入的任务/练习数量。")
    parser.add_argument("--fill-cover", action="store_true", help="默认不动第一页；开启后只填封面学院、专业班级、学号、姓名。")
    parser.add_argument("--college", default="", help="封面学院。")
    parser.add_argument("--major-class", default="", help="封面专业班级。")
    parser.add_argument("--class-name", default="", help="报告命名用班级，例如：计科1班。")
    parser.add_argument("--group", default="", help="报告命名用组别，例如：2组。")
    parser.add_argument("--student-id", default="", help="封面学号和报告命名用学号。")
    parser.add_argument("--name", default="", help="封面姓名和报告命名用姓名。")
    parser.add_argument("--ai", action="store_true", help="启用 OpenAI 兼容接口增强课程总结与代码生成。")
    parser.add_argument("--ai-base-url", default="", help="OpenAI 兼容 API 地址，例如：https://example.com/v1。")
    parser.add_argument("--ai-api-key", default="", help="API Key；没有鉴权要求的本地/免费接口可以留空。")
    parser.add_argument("--ai-model", default="", help="模型名称。")
    parser.add_argument("--ai-temperature", type=float, default=0.2, help="AI 生成温度。")
    return parser.parse_args()


def build_ai_config_from_args(args: argparse.Namespace) -> AIConfig:
    if not getattr(args, "ai", False):
        return AIConfig()
    return AIConfig(
        enabled=True,
        base_url=args.ai_base_url.strip(),
        api_key=args.ai_api_key.strip(),
        model=args.ai_model.strip(),
        temperature=args.ai_temperature,
    )


def resolve_template(template_arg: str | None) -> Path:
    candidates = [Path(template_arg)] if template_arg else DEFAULT_TEMPLATE_CANDIDATES
    for candidate in candidates:
        if candidate and candidate.exists():
            return candidate.resolve()
    raise FileNotFoundError("没有找到实训报告模板，请使用 --template 指定模板 DOCX 路径。")


def resolve_sources(source_args: list[str]) -> list[Path]:
    if source_args:
        paths = [Path(item).resolve() for item in source_args]
    else:
        paths = discover_uploaded_txt_sources()
        if not paths:
            paths = [item.resolve() for item in DEFAULT_SOURCE_CANDIDATES if item.exists()]
    if not paths:
        raise FileNotFoundError("没有找到课堂笔记资料，请把 TXT/PPTX/DOCX 路径作为参数传入。")
    missing = [str(path) for path in paths if not path.exists()]
    if missing:
        raise FileNotFoundError("以下资料文件不存在：" + "；".join(missing))
    return paths


def discover_uploaded_txt_sources() -> list[Path]:
    txt_files = [path.resolve() for path in Path.cwd().glob("*.txt") if path.is_file()]
    return sorted(txt_files, key=natural_sort_key)


def natural_sort_key(path: Path) -> list[object]:
    parts = re.split(r"(\d+)", path.name)
    return [int(part) if part.isdigit() else part.lower() for part in parts]


def group_sources_by_day(paths: list[Path]) -> list[SourceGroup]:
    grouped: dict[tuple, SourceGroup] = {}
    for path in sorted(paths, key=natural_sort_key):
        month_day = extract_month_day_key(path)
        if month_day:
            month, day = month_day
            key = ("date", month, day)
            label = f"{month}-{day}"
        else:
            key = ("file", path.name.lower())
            label = path.stem
        if key not in grouped:
            grouped[key] = SourceGroup(label=label, paths=[], sort_key=key)
        grouped[key].paths.append(path)
    return sorted(grouped.values(), key=lambda group: group.sort_key)


def extract_month_day_key(path: Path) -> tuple[int, int] | None:
    name = path.stem

    year_match = re.search(r"(?:^|[^\d])(?:19|20)\d{2}[-_.年](\d{1,2})[-_.月](\d{1,2})(?=[-_.日]|[^\d]|$)", name)
    if year_match:
        return int(year_match.group(1)), int(year_match.group(2))

    month_day_match = re.search(r"(?:^|[^\d])(\d{1,2})[-_.月](\d{1,2})(?=[-_.日]|[^\d]|$)", name)
    if month_day_match:
        return int(month_day_match.group(1)), int(month_day_match.group(2))
    return None


def read_source_group_text(group: SourceGroup) -> str:
    blocks: list[str] = []
    for path in group.paths:
        text = read_source_text(path).strip()
        blocks.append(f"【资料：{path.name}】\n{text}")
    return "\n\n".join(blocks)


def resolve_output_path(args: argparse.Namespace) -> Path:
    if args.output:
        return unique_path(Path(args.output).resolve())

    if args.class_name and args.group and args.student_id and args.name:
        group = args.group if args.group.endswith("组") else f"{args.group}组"
        filename = f"{args.class_name}-{group}-{args.student_id}-{args.name}.docx"
    else:
        filename = f"实训报告_{datetime.now():%Y%m%d_%H%M%S}.docx"
    return unique_path((Path.cwd() / filename).resolve())


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    for index in range(1, 1000):
        candidate = path.with_name(f"{stem}_{index}{suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"无法为输出文件生成唯一名称：{path}")


def fill_cover(doc: Document, args: argparse.Namespace) -> None:
    values = {
        "学    院": args.college,
        "专业班级": args.major_class or args.class_name,
        "学    号": args.student_id,
        "姓    名": args.name,
    }
    for paragraph in doc.paragraphs[:30]:
        text = paragraph.text
        for label, value in values.items():
            if value and label in text:
                clear_paragraph(paragraph)
                paragraph.add_run(f"{label}： {value}")
                for run in paragraph.runs:
                    set_run_font(run, size=BODY_SIZE)


def clear_template_body_after_cover(doc: Document) -> None:
    body = doc.element.body
    remove_from = None
    paragraph_elements = {paragraph._element: paragraph for paragraph in doc.paragraphs}

    for element in list(body):
        paragraph = paragraph_elements.get(element)
        if paragraph is None:
            continue
        text = paragraph.text.strip()
        if text.startswith("一、第1天") or text.startswith("第一天"):
            remove_from = element
            break

    if remove_from is None:
        saw_cover_section = False
        for element in list(body):
            paragraph = paragraph_elements.get(element)
            if paragraph is None:
                continue
            if saw_cover_section:
                remove_from = element
                break
            if paragraph._p.pPr is not None and paragraph._p.pPr.sectPr is not None:
                saw_cover_section = True

    if remove_from is None:
        doc.add_page_break()
        return

    removing = False
    for element in list(body):
        if element is remove_from:
            removing = True
        if removing and element.tag != qn("w:sectPr"):
            body.remove(element)


def ensure_report_styles(doc: Document) -> None:
    ensure_style(doc, "实训报告_正文", size=BODY_SIZE, first_line=True)
    ensure_style(doc, "实训报告_小节标题", size=BODY_SIZE, first_line=False)
    ensure_style(doc, "实训报告_代码", size=BODY_SIZE, first_line=False, left_indent=BODY_FIRST_LINE_INDENT)
    ensure_style(doc, "实训报告_图题", size=BODY_SIZE, first_line=False, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    ensure_style(doc, "实训报告_日标题", size=14, first_line=False, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER)


def ensure_style(
    doc: Document,
    name: str,
    size: float,
    first_line: bool,
    bold: bool = False,
    alignment: WD_ALIGN_PARAGRAPH | None = None,
    left_indent=Pt(0),
) -> None:
    try:
        style = doc.styles[name]
    except KeyError:
        style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    style.base_style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(size)
    style.font.bold = bold
    set_style_east_asia_font(style, "宋体")

    fmt = style.paragraph_format
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    fmt.line_spacing = 1.5
    fmt.first_line_indent = BODY_FIRST_LINE_INDENT if first_line else Pt(0)
    fmt.left_indent = left_indent
    if alignment is not None:
        fmt.alignment = alignment


def set_style_east_asia_font(style, east_asia_font: str) -> None:
    r_pr = style.element.rPr
    if r_pr is None:
        r_pr = OxmlElement("w:rPr")
        style.element.append(r_pr)
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    for attr in ("w:eastAsia", "w:cs"):
        r_fonts.set(qn(attr), east_asia_font)
    for attr in ("w:ascii", "w:hAnsi"):
        r_fonts.set(qn(attr), "Times New Roman")


def read_source_text(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".txt":
        return read_text_file(path)
    if suffix == ".pptx":
        return read_pptx_text(path)
    if suffix == ".docx":
        return read_docx_text(path)
    raise ValueError(f"暂不支持的资料格式：{path.suffix}，请使用 TXT/PPTX/DOCX。")


def read_text_file(path: Path) -> str:
    for encoding in TEXT_ENCODINGS:
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="gb18030", errors="replace")


def read_docx_text(path: Path) -> str:
    document = Document(str(path))
    lines = [paragraph.text for paragraph in document.paragraphs if paragraph.text.strip()]
    return "\n".join(lines)


def read_pptx_text(path: Path) -> str:
    try:
        from pptx import Presentation
    except ImportError as exc:
        raise RuntimeError("读取 PPTX 需要安装 python-pptx；可以先把 PPT 内容整理成 TXT 再运行。") from exc

    presentation = Presentation(str(path))
    lines: list[str] = []
    for slide_index, slide in enumerate(presentation.slides, start=1):
        lines.append(f"第{slide_index}页PPT")
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                lines.append(shape.text.strip())
    return "\n".join(lines)


def build_day_report(
    source: Path,
    day_index: int,
    raw_text: str,
    title_override: str | None,
    max_tasks: int,
    ai_config: AIConfig | None = None,
) -> DayReport:
    topics = extract_topics(raw_text)
    goals = generate_goals(topics)
    details = generate_details(topics, raw_text)
    tasks = extract_tasks(raw_text)
    tasks.extend(suggest_tasks_from_keywords(raw_text, tasks))
    tasks = merge_tasks(tasks)[:max_tasks]
    if not tasks:
        tasks = [
            build_task(
                title="课堂知识点综合练习",
                requirement="根据当天课堂笔记，编写一个综合示例程序，对主要知识点进行练习。",
            )
        ]

    title = title_override or infer_day_title(source, raw_text, topics)
    report = DayReport(source, day_index, title, goals, topics, details, tasks)
    if ai_config and ai_config.enabled:
        return enhance_day_report_with_ai(report, raw_text, max_tasks, ai_config)
    return report


def enhance_day_report_with_ai(
    fallback: DayReport,
    raw_text: str,
    max_tasks: int,
    ai_config: AIConfig,
) -> DayReport:
    if not ai_config.base_url or not ai_config.model:
        raise ValueError("启用 AI 增强时必须填写 API 地址和模型名。")

    prompt = build_ai_prompt(fallback, raw_text, max_tasks)
    content = call_openai_compatible_chat(ai_config, prompt)
    payload = extract_json_object(content)
    return merge_ai_payload_into_report(fallback, payload, max_tasks)


def build_ai_prompt(fallback: DayReport, raw_text: str, max_tasks: int) -> str:
    text = raw_text.strip()
    if len(text) > 9000:
        text = text[:9000] + "\n……（资料过长，后文已截断）"
    fallback_tasks = [
        {"title": task.title, "requirement": task.requirement, "code": task.code}
        for task in fallback.tasks[:max_tasks]
    ]
    return f"""
你是暑期实训报告助手。请根据课堂 TXT 笔记，为第{fallback.day_index}天生成实训报告内容。

必须遵守：
1. 只输出一个 JSON 对象，不要 Markdown，不要解释。
2. 中文内容用于 Word 报告，语言自然、像学生实训报告。
3. 课程目标反向概括当天学习目标。
4. 课程内容只写知识点名称，用数组列出。
5. 课程内容详情用数组列出，每条是完整句子。
6. 课程代码及执行过程要把任务/练习拆开，每个任务单独写 requirement 和 Python code。
7. code 必须是纯文本 Python 代码，不要图片，不要 Markdown 代码块。
8. 最多生成 {max_tasks} 个任务；如果 TXT 任务少，可以根据知识点补充合理练习。

JSON 格式：
{{
  "title": "当天标题",
  "goals": ["课程目标1", "课程目标2"],
  "topics": ["知识点1", "知识点2"],
  "details": ["内容详情1", "内容详情2"],
  "tasks": [
    {{
      "title": "任务标题",
      "requirement": "任务要求",
      "code": "Python代码",
      "caption": "图题短名称"
    }}
  ]
}}

本地规则底稿：
{json.dumps({
    "title": fallback.title,
    "goals": fallback.goals,
    "topics": fallback.topics,
    "details": fallback.details,
    "tasks": fallback_tasks,
}, ensure_ascii=False, indent=2)}

课堂 TXT 笔记：
{text}
""".strip()


def call_openai_compatible_chat(ai_config: AIConfig, prompt: str) -> str:
    endpoint = normalize_chat_endpoint(ai_config.base_url)
    body = {
        "model": ai_config.model,
        "temperature": ai_config.temperature,
        "stream": False,
        "max_tokens": 6000,
        "response_format": {"type": "json_object"},
        "messages": [
            {
                "role": "system",
                "content": "你只输出严格 JSON。不要输出 Markdown、解释、前后缀。",
            },
            {"role": "user", "content": prompt},
        ],
    }
    data = json.dumps(body, ensure_ascii=False).encode("utf-8")
    headers = {"Content-Type": "application/json"}
    if ai_config.api_key:
        headers["Authorization"] = f"Bearer {ai_config.api_key}"
    req = url_request.Request(endpoint, data=data, headers=headers, method="POST")
    try:
        with url_request.urlopen(req, timeout=ai_config.timeout) as response:
            response_text = response.read().decode("utf-8")
    except url_error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"AI 接口请求失败：HTTP {exc.code}，{detail[:300]}") from exc
    except url_error.URLError as exc:
        raise RuntimeError(f"AI 接口连接失败：{format_ai_connection_error(exc.reason)}") from exc

    payload = json.loads(response_text)
    choices = payload.get("choices") or []
    if not choices:
        raise RuntimeError("AI 接口没有返回 choices。")
    message = choices[0].get("message") or {}
    content = message.get("content") or choices[0].get("text") or ""
    if not content.strip():
        raise RuntimeError("AI 接口返回内容为空。")
    return content


def normalize_chat_endpoint(base_url: str) -> str:
    base_url = base_url.strip().rstrip("/")
    if not base_url:
        raise ValueError("AI API 地址不能为空。")
    if base_url.endswith("/chat/completions"):
        return base_url
    return f"{base_url}/chat/completions"


def format_ai_connection_error(reason) -> str:
    reason_text = str(reason)
    if "10013" in reason_text:
        return (
            f"{reason_text}。这是 Windows 拒绝外网 socket 连接的错误，通常不是 API Key 写错。"
            "请优先关闭当前网页服务窗口后，双击“启动实训报告网页.bat”重新启动；"
            "如果仍失败，请在 Windows 防火墙/杀毒软件里允许 python.exe 访问网络，并检查 VPN/代理是否拦截了 https://api.deepseek.com。"
        )
    if "getaddrinfo failed" in reason_text or "Name or service not known" in reason_text:
        return f"{reason_text}。请检查 AI API 地址是否填写正确，以及网络/DNS 是否可用。"
    if "timed out" in reason_text or "timeout" in reason_text.lower():
        return f"{reason_text}。连接超时，请检查网络、代理/VPN 或 DeepSeek 服务状态。"
    return reason_text


def extract_json_object(content: str) -> dict:
    text = content.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\s*```$", "", text)
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        raise RuntimeError("AI 返回内容不是 JSON 对象。")
    return json.loads(text[start : end + 1])


def merge_ai_payload_into_report(fallback: DayReport, payload: dict, max_tasks: int) -> DayReport:
    title = cleanup_heading(str(payload.get("title") or fallback.title)) or fallback.title
    goals = sanitize_string_list(payload.get("goals"), fallback.goals, 7)
    topics = sanitize_string_list(payload.get("topics"), fallback.topics, 24)
    details = sanitize_string_list(payload.get("details"), fallback.details, 10)

    tasks: list[TaskItem] = []
    for item in payload.get("tasks") or []:
        if not isinstance(item, dict):
            continue
        task_title = cleanup_heading(str(item.get("title") or "课堂练习"))
        requirement = cleanup_requirement(str(item.get("requirement") or task_title))
        code = str(item.get("code") or "").strip()
        caption = cleanup_caption(str(item.get("caption") or task_title))
        if not code:
            code = generate_code(task_title, requirement)
        tasks.append(TaskItem(task_title, requirement, code, caption))
        if len(tasks) >= max_tasks:
            break
    if not tasks:
        tasks = fallback.tasks[:max_tasks]

    return DayReport(fallback.source, fallback.day_index, title, goals, topics, details, tasks)


def sanitize_string_list(value, fallback: list[str], limit: int) -> list[str]:
    if not isinstance(value, list):
        return fallback[:limit]
    items: list[str] = []
    for item in value:
        text = str(item).strip()
        if text:
            items.append(text)
    return unique_keep_order(items)[:limit] or fallback[:limit]


def extract_topics(text: str) -> list[str]:
    lowered = text.lower()
    topics: list[str] = []
    for topic, keywords in TOPIC_RULES:
        if any(keyword.lower() in lowered for keyword in keywords):
            topics.append(topic)

    for line in clean_lines(text):
        match = re.match(r"^[一二三四五六七八九十]+[、.．]\s*(.+)$", line)
        if not match:
            continue
        heading = cleanup_heading(match.group(1))
        if heading and heading not in {"复习", "任务", "提交报告要求"}:
            topics.append(heading)
    return unique_keep_order(topics)


def generate_goals(topics: list[str]) -> list[str]:
    goals = [GOAL_BANK[topic] for topic in topics if topic in GOAL_BANK]
    if not goals:
        goals = ["根据课堂笔记梳理当天学习内容，掌握相关知识点的基本概念、常用语法和实践方法。"]
    return unique_keep_order(goals)[:7]


def generate_details(topics: list[str], text: str) -> list[str]:
    details = [DETAIL_BANK[topic] for topic in topics if topic in DETAIL_BANK]
    if len(details) < 4:
        for line in clean_lines(text):
            normalized = cleanup_heading(line)
            if 8 <= len(normalized) <= 80 and not is_assignment_line(normalized):
                details.append(f"课堂笔记中还涉及“{normalized}”相关内容，需要结合示例进一步理解其用法。")
            if len(details) >= 6:
                break
    return unique_keep_order(details)[:10]


def extract_tasks(text: str) -> list[TaskItem]:
    lines = clean_lines(text, keep_indent=True)
    tasks: list[TaskItem] = []
    current_title = ""
    current_lines: list[str] = []
    in_exercise = False

    def flush_current() -> None:
        nonlocal current_title, current_lines
        if current_title:
            requirement = " ".join(item.strip() for item in current_lines if item.strip())
            tasks.append(build_task(current_title, requirement or current_title))
        current_title = ""
        current_lines = []

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue
        if re.match(r"^练习[:：]?$", line):
            in_exercise = True
            flush_current()
            continue
        if in_exercise and re.match(r"^[一二三四五六七八九十]+[、.．]\s*", line):
            break
        if in_exercise:
            match = re.match(r"^\s*(\d+)[.、]\s*(.+)$", raw_line)
            if match:
                flush_current()
                current_title = infer_task_title(match.group(2))
                current_lines = [match.group(2)]
            else:
                current_lines.append(line)

    flush_current()
    return tasks


def suggest_tasks_from_keywords(text: str, existing: list[TaskItem]) -> list[TaskItem]:
    additions: list[TaskItem] = []
    existing_text = " ".join(task.title + task.requirement for task in existing)

    def add_once(title: str, requirement: str, *needles: str) -> None:
        haystack = text + existing_text
        if all(needle not in haystack for needle in needles):
            return
        if any(title in task.title for task in existing + additions):
            return
        additions.append(build_task(title, requirement))

    add_once("九九乘法表输出", "使用双层for循环打印九九乘法表，外层控制行，内层控制列。", "九九乘法表")
    add_once("猜数字游戏", "随机产生1到100之间的数字，提供6次猜测机会，记录每次猜测并提示大小关系。", "猜数字")
    add_once("break与continue循环控制", "编写循环程序，使用continue跳过部分数据，使用break提前结束循环。", "break", "continue")
    add_once("函数定义、调用与返回值", "定义带参数的函数，调用函数并使用return返回计算结果。", "函数定义", "函数的调用", "return")
    add_once("文件写入与读取", "使用with open()创建文本文件，写入实训记录后再次读取并输出。", "文件读写", "with", "open")
    return additions


def build_task(title: str, requirement: str) -> TaskItem:
    title = cleanup_heading(title)
    requirement = cleanup_requirement(requirement)
    code = generate_code(title, requirement)
    caption = cleanup_caption(title)
    return TaskItem(title=title, requirement=requirement, code=code, caption=caption)


def generate_code(title: str, requirement: str) -> str:
    text = f"{title} {requirement}"
    if contains_all(text, ("用户名", "大写")):
        return dedent(
            """
            username = input("请输入用户名（包含数字、字母、下划线，长度为8）：")

            if len(username) == 8 and all(char.isalnum() or char == "_" for char in username):
                print("用户名：", username.upper())
            else:
                print("用户名格式不符合要求")
            """
        ).strip()
    if "足球" in text:
        return dedent(
            """
            news = "在昨晚进行的一场足球比赛中，主场作战的球队以3比1战胜了客队。足球运动吸引了众多球迷观看，年轻球员的配合提升了足球进攻的效率。"

            print("原始新闻：", news)
            print("第1个'足球'出现的位置：", news.find("足球"))
            print("'足球'出现的次数：", news.count("足球"))
            print("替换后的新闻：", news.replace("的", "地"))
            """
        ).strip()
    if "strip" in text.lower() or "去掉字符串前后" in text or "前后的空格" in text:
        return dedent(
            """
            mystr5 = "\\n\\n\\n我们正在进行暑假实训\\n\\n\\n"

            print("原始字符串：", repr(mystr5))
            print("去除前后空白后：", repr(mystr5.strip()))
            """
        ).strip()
    if "split" in text.lower() or "join" in text.lower() or "拆分" in text or "连接成字符串" in text:
        return dedent(
            """
            mystr5 = "姓名:小花，年龄:20，性别:女，电话:13444444444"
            mylist2 = mystr5.split("，")
            result = "+".join(mylist2)

            print("拆分后的列表：", mylist2)
            print("组合后的字符串：", result)
            """
        ).strip()
    if "九九乘法表" in text:
        return dedent(
            """
            for i in range(1, 10):
                for j in range(1, i + 1):
                    print(f"{j}*{i}={i * j}", end="\\t")
                print()
            """
        ).strip()
    if "猜数字" in text:
        return dedent(
            """
            import random

            answer = random.randint(1, 100)
            records = []

            for chance in range(1, 7):
                guess = int(input(f"第{chance}次请输入1到100之间的数字："))
                records.append(guess)

                if guess == answer:
                    print("恭喜你猜对了")
                    break
                if guess > answer:
                    print("您猜大了")
                else:
                    print("您猜小了")

                print("剩余次数：", 6 - chance)
            else:
                print("机会用完，正确数字是：", answer)

            print("猜测记录：", records)
            """
        ).strip()
    if "break" in text or "continue" in text:
        return dedent(
            """
            for number in range(1, 21):
                if number % 2 == 0:
                    continue
                if number > 15:
                    break
                print(number, end=" ")
            """
        ).strip()
    if "函数" in text or "return" in text:
        return dedent(
            """
            def calc_total(price, count):
                total = price * count
                return total


            goods_name = input("请输入商品名称：")
            price = float(input("请输入商品单价："))
            count = int(input("请输入购买数量："))
            total_money = calc_total(price, count)

            print(f"{goods_name}的总价为：{total_money:.2f}元")
            """
        ).strip()
    if "文件" in text or "open" in text.lower() or "write" in text.lower() or "read" in text.lower():
        return dedent(
            """
            file_name = "实训记录.txt"

            with open(file_name, "w", encoding="utf-8") as file:
                file.write("今天学习了Python文件读写命令。\\n")
                file.write("with open()可以自动关闭文件。\\n")

            with open(file_name, "r", encoding="utf-8") as file:
                content = file.read()

            print("文件内容如下：")
            print(content)
            """
        ).strip()
    if "数据类型" in text or "type" in text.lower():
        return dedent(
            """
            samples = [
                ("字符类型", "Python编程语言"),
                ("数值类型", 2026),
                ("布尔类型", True),
                ("列表", ["西安", "北京", "上海"]),
                ("字典", {"姓名": "小明", "年龄": 20}),
                ("元组", ("暑期实训",)),
                ("集合", {1, 2, 2, 3}),
            ]

            for type_name, value in samples:
                print(type_name, "：", value, "，类型为：", type(value))
            """
        ).strip()
    if "输入" in text and "输出" in text:
        return dedent(
            """
            class_name = input("请输入班级：")
            group = input("请输入组别：")
            name = input("请输入姓名：")

            print(f"我是{class_name}班的学生，我是第{group}组，我的姓名是：{name}")
            """
        ).strip()
    return dedent(
        f"""
        # {title}
        print("任务名称：{title}")
        print("任务要求：{requirement}")
        print("请根据课堂运行结果补充截图。")
        """
    ).strip()


def append_day_report(doc: Document, report: DayReport, screenshots_dir: Path | None) -> None:
    add_day_heading(doc, f"第{chinese_number(report.day_index)}天：{report.title}")

    add_section_heading(doc, "1、课程目标：")
    for index, goal in enumerate(report.goals, start=1):
        add_body_paragraph(doc, f"（{index}）{goal}")

    add_section_heading(doc, "2、课程内容：")
    add_body_paragraph(doc, "；".join(report.topics) + "。")

    add_section_heading(doc, "3、课程内容详情：")
    for index, detail in enumerate(report.details, start=1):
        add_body_paragraph(doc, f"（{index}）{detail}")

    add_section_heading(doc, "4、课程代码及执行过程：")
    for task_index, task in enumerate(report.tasks, start=1):
        add_task_block(doc, report.day_index, task_index, task, screenshots_dir)


def add_task_block(
    doc: Document,
    day_index: int,
    task_index: int,
    task: TaskItem,
    screenshots_dir: Path | None,
) -> None:
    add_section_heading(doc, f"任务{task_index}：{task.title}")
    add_body_paragraph(doc, f"【任务要求】{task.requirement}")
    add_body_paragraph(doc, "【代码】")
    add_code_block(doc, task.code)
    add_body_paragraph(doc, "运行结果：")

    image_path = find_screenshot(screenshots_dir, day_index, task_index) if screenshots_dir else None
    if image_path:
        paragraph = doc.add_paragraph(style="实训报告_图题")
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run()
        run.add_picture(str(image_path), width=Inches(5.7))
    else:
        add_caption_paragraph(doc, f"（此处插入 {day_index}-{task_index} {task.caption}运行结果截图）")
    add_caption_paragraph(doc, f"图{day_index}-{task_index} {task.caption}运行结果")


def add_day_heading(doc: Document, text: str) -> None:
    paragraph = doc.add_paragraph(style="实训报告_日标题")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(text)
    set_paragraph_format(paragraph, "day")
    set_run_font(run, size=14, bold=True)


def add_section_heading(doc: Document, text: str) -> None:
    paragraph = doc.add_paragraph(style="实训报告_小节标题")
    run = paragraph.add_run(text)
    set_paragraph_format(paragraph, "section")
    set_run_font(run, size=BODY_SIZE)


def add_body_paragraph(doc: Document, text: str) -> None:
    paragraph = doc.add_paragraph(style="实训报告_正文")
    run = paragraph.add_run(text)
    set_paragraph_format(paragraph, "body")
    set_run_font(run, size=BODY_SIZE)


def add_caption_paragraph(doc: Document, text: str) -> None:
    paragraph = doc.add_paragraph(style="实训报告_图题")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(text)
    set_paragraph_format(paragraph, "caption")
    set_run_font(run, size=BODY_SIZE)


def add_code_block(doc: Document, code: str) -> None:
    for line in code.splitlines():
        paragraph = doc.add_paragraph(style="实训报告_代码")
        paragraph.add_run(line if line else " ")
        set_paragraph_format(paragraph, "code")
        for run in paragraph.runs:
            set_run_font(run, size=BODY_SIZE)


def set_paragraph_format(paragraph, role: str) -> None:
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    fmt.line_spacing = 1.5
    if role == "body":
        fmt.first_line_indent = BODY_FIRST_LINE_INDENT
        fmt.left_indent = Pt(0)
    elif role == "code":
        fmt.first_line_indent = Pt(0)
        fmt.left_indent = BODY_FIRST_LINE_INDENT
    else:
        fmt.first_line_indent = Pt(0)
        fmt.left_indent = Pt(0)
    if role in {"caption", "day"}:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER


def set_run_font(run, size: float = BODY_SIZE, bold: bool | None = None) -> None:
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold

    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    r_fonts.set(qn("w:ascii"), "Times New Roman")
    r_fonts.set(qn("w:hAnsi"), "Times New Roman")
    r_fonts.set(qn("w:eastAsia"), "宋体")
    r_fonts.set(qn("w:cs"), "Times New Roman")


def clear_paragraph(paragraph) -> None:
    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)


def find_screenshot(screenshots_dir: Path | None, day_index: int, task_index: int) -> Path | None:
    if screenshots_dir is None or not screenshots_dir.exists():
        return None
    prefixes = [f"{day_index}-{task_index}", f"图{day_index}-{task_index}"]
    extensions = (".png", ".jpg", ".jpeg", ".bmp")
    matches: list[Path] = []
    for path in screenshots_dir.iterdir():
        if path.suffix.lower() not in extensions:
            continue
        if any(path.stem.startswith(prefix) for prefix in prefixes):
            matches.append(path)
    return sorted(matches)[0] if matches else None


def infer_day_title(source: Path, text: str, topics: list[str]) -> str:
    if "九九乘法表" in text and "文件读写" in text:
        return "循环结构、函数与文件读写"
    if "数据类型函数" in text or "upper()" in text or "split" in text:
        return "字符串函数与文本处理"
    if "数据类型的书写" in text and "输入/输出" in text:
        return "Python编程语言基础与数据类型元素访问"
    for topic in topics:
        if topic not in {"实训规划与注意事项", "Python编程语言与业务应用"}:
            return topic
    return cleanup_heading(source.stem)


def infer_task_title(requirement: str) -> str:
    text = cleanup_requirement(requirement)
    if "用户名" in text and "大写" in text:
        return "用户名大写输出"
    if "足球" in text:
        return "足球新闻字符串查找、统计与替换"
    if "去掉" in text and "空格" in text or "strip" in text.lower():
        return "字符串strip()去空"
    if "拆分" in text or "连接成字符串" in text or "join" in text.lower():
        return "字符串split()拆分与join()组合"
    return cleanup_caption(text[:28]) or "课堂练习"


def cleanup_heading(text: str) -> str:
    text = text.strip()
    text = re.sub(r"^[（(]?\d+[）).、]\s*", "", text)
    text = text.strip("：:；;，,。 ")
    return text


def cleanup_requirement(text: str) -> str:
    text = " ".join(text.split())
    text = re.sub(r"^[（(]?\d+[）).、]\s*", "", text)
    return text.strip()


def cleanup_caption(text: str) -> str:
    text = cleanup_requirement(text)
    text = re.sub(r"[。；;：:，,]+$", "", text)
    text = re.sub(r"\s+", "", text)
    return text[:32]


def clean_lines(text: str, keep_indent: bool = False) -> list[str]:
    lines: list[str] = []
    for raw_line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = raw_line.rstrip()
        if not line.strip():
            continue
        lines.append(line if keep_indent else line.strip())
    return lines


def is_assignment_line(line: str) -> bool:
    return bool(re.search(r"提交|邮箱|附件|命名|最迟|报告要求", line))


def contains_all(text: str, needles: Iterable[str]) -> bool:
    return all(needle in text for needle in needles)


def unique_keep_order(items: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for item in items:
        item = item.strip()
        if not item or item in seen:
            continue
        seen.add(item)
        result.append(item)
    return result


def merge_tasks(tasks: list[TaskItem]) -> list[TaskItem]:
    result: list[TaskItem] = []
    seen: set[str] = set()
    for task in tasks:
        key = task.caption
        if key in seen:
            continue
        seen.add(key)
        result.append(task)
    return result


def chinese_number(number: int) -> str:
    if 0 <= number <= 10:
        return CN_NUMS[number]
    if number < 20:
        return "十" + CN_NUMS[number - 10]
    tens, ones = divmod(number, 10)
    return CN_NUMS[tens] + "十" + (CN_NUMS[ones] if ones else "")


if __name__ == "__main__":
    raise SystemExit(main())
