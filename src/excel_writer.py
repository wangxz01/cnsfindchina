"""Excel 输出（方案 A 布局）。

布局：
  每个 issue 一个 sheet（按 vol/issue 命名）
  Row 1, Col A                 = issue URL
  之后按 section 出现顺序：
    分类名独占一行（Col A）
    每篇文章一行：按 columns schema 渲染
"""
import re
import time
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter


# 列 schema：(表头, 取值键或 callable)
# - 字符串键：直接从 fields dict 取；特殊键 "authors_joined" 自动 join
# - callable：fields dict -> 单元格值
DEFAULT_COLUMNS = [
    ("URL", "url"),
    ("标题", "title"),
    ("DOI", "doi"),
    ("作者", "authors_joined"),
    ("首条单位", "first_aff"),
]

NATURE_COLUMNS = [
    ("相关网址", "url"),
    ("原始数据-标题", "title"),
    ("原始数据-DOI", "doi"),
    ("原始数据-类型", "type"),
    ("原始数据-第一作者", "first_author"),
    ("原始数据-第一完成单位", "first_aff"),
    ("提取-国家", "first_author_country"),
    ("判断-是否中国", "is_china_label"),
    ("原始数据-所有作者", "authors_joined"),
]


def _sheet_name_from_url(url: str) -> str:
    # Cell 格式：/vol/X/issue/Y
    m = re.search(r"/vol(?:umes)?/(\d+)/issues?/(\d+)", url)
    if m:
        return f"v{m.group(1)}-i{m.group(2)}"
    parts = url.rstrip("/").split("/")
    name = "-".join(parts[-2:]) or "issue"
    return name[:31]  # Excel sheet 名最长 31 字符


def _cell_value(key, f: dict):
    if callable(key):
        return key(f)
    if key == "authors_joined":
        return "; ".join(f.get("authors", []))
    if key == "is_china_label":
        return "是" if f.get("is_china") else "否"
    return f.get(key, "")


def write_one_sheet(wb, issue_url: str, results: list,
                    columns=DEFAULT_COLUMNS, col_widths=None) -> None:
    ws = wb.create_sheet(_sheet_name_from_url(issue_url))
    ws.append([issue_url])
    ws.append([name for name, _ in columns])  # 表头行（每个 sheet 仅顶部一次）
    last_section = None
    for section, url, f in results:
        if section != last_section:
            ws.append([section])
            last_section = section
        ws.append([_cell_value(k, f) for _, k in columns])

    widths = col_widths or [60] * len(columns)
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_excel(out_path: str, all_issues: list,
                columns=DEFAULT_COLUMNS, col_widths=None) -> str:
    """all_issues: [(issue_url, results), ...]。每个 issue 一个 sheet。

    返回实际写入的路径（若 out_path 被占用，会改写到带时间戳的备用文件）。
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # 删除默认 sheet
    for issue_url, results in all_issues:
        write_one_sheet(wb, issue_url, results, columns=columns, col_widths=col_widths)

    try:
        wb.save(out_path)
        return out_path
    except PermissionError:
        # 文件被占用（如 Excel 打开中）→ 写到带时间戳的备用文件
        ts = time.strftime("%Y%m%d_%H%M%S")
        alt = str(Path(out_path).with_suffix(f".{ts}.xlsx"))
        print(f"[warn] {out_path} 被占用，改写到: {alt}")
        wb.save(alt)
        return alt


