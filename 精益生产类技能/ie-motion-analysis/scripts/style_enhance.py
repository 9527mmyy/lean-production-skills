#!/usr/bin/env python3
"""
xlsx 外观美化（IE 动作分析 Step 6 美化）

对 build_report.py 输出的素表做二次样式增强：
  - 统一字体微软雅黑
  - 标题 16pt 深蓝加粗居中
  - 表头钢蓝底（#2E75B6）白字 9pt
  - 严重度/价值/ECRS 条件着色
  - 交替行底色
  - A4 横向、fitToPage、水平居中

用法:
  python style_enhance.py <xlsx 文件>
"""

import sys
from pathlib import Path

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.worksheet.properties import PageSetupProperties
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.stderr.write("[错误] 需要 openpyxl：pip install openpyxl\n")
    sys.exit(2)

# ---- 颜色常量 ----
STEEL_BLUE = "2E75B6"
LIGHT_GRAY = "F2F2F2"
WHITE = "FFFFFF"

SEV_FILLS = {
    "高": ("FFC7CE", "9C0006"),
    "中": ("FFEB9C", "9C6500"),
    "低": ("C6EFCE", "006100"),
}
VALUE_FILLS = {
    "V": ("E2EFDA", "1b5e20"),
    "N": ("FCE4D6", "b71c1c"),
    "M": ("FFF2CC", "8d6e00"),
}
ECRS_FILL = ("BDD7EE", "1f4e79")

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def enhance_problem_sheet(ws):
    """问题项检查表美化"""
    # 标题
    ws["A1"].font = Font(name="微软雅黑", size=16, bold=True, color=STEEL_BLUE)
    ws.row_dimensions[1].height = 30

    # 信息栏 R2-R3
    for r in (2, 3):
        for c in range(1, 13):
            cell = ws.cell(row=r, column=c)
            cell.font = Font(name="微软雅黑", size=9)
            cell.border = BORDER
            if cell.value and str(cell.value).endswith("："):
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.font = Font(name="微软雅黑", size=9, bold=True)
                cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=Side(style="medium", color="BFBFBF"))
            else:
                cell.alignment = LEFT
        ws.row_dimensions[r].height = 22

    # 表头 R4
    for c in range(1, 13):
        cell = ws.cell(row=4, column=c)
        cell.font = Font(name="微软雅黑", size=9, bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=STEEL_BLUE)
        cell.alignment = CENTER
        cell.border = BORDER
    ws.row_dimensions[4].height = 28

    # 数据行 R5+
    last_data = ws.max_row
    # Find actual data end (before signature row)
    for r in range(last_data, 4, -1):
        v = ws.cell(row=r, column=1).value
        if v and str(v).startswith(("分析人", "班组长", "改善负责", "日期")):
            last_data = r - 1
            break

    for r in range(5, last_data + 1):
        # Alternating row background
        bg = LIGHT_GRAY if (r % 2 == 0) else None
        for c in range(1, 13):
            cell = ws.cell(row=r, column=c)
            cell.font = Font(name="微软雅黑", size=9)
            cell.border = BORDER
            if c in (2, 3, 4, 5, 8, 12):
                cell.alignment = LEFT
            else:
                cell.alignment = CENTER
            if bg:
                cell.fill = PatternFill("solid", fgColor=bg)
        # Severity coloring (column I)
        sev = str(ws.cell(row=r, column=9).value or "")
        if sev in SEV_FILLS:
            fg, fc = SEV_FILLS[sev]
            cell = ws.cell(row=r, column=9)
            cell.fill = PatternFill("solid", fgColor=fg)
            cell.font = Font(name="微软雅黑", size=9, bold=True, color=fc)
        ws.row_dimensions[r].height = 34

    # Signature row
    if last_data < ws.max_row:
        r = last_data + 1
        for c in range(1, 13):
            cell = ws.cell(row=r, column=c)
            cell.font = Font(name="微软雅黑", size=9, bold=True)
            cell.fill = PatternFill("solid", fgColor="D6E4F0")
            cell.border = Border(left=THIN, right=THIN, top=Side(style="medium", color=STEEL_BLUE), bottom=THIN)
            cell.alignment = CENTER
        ws.row_dimensions[r].height = 28


def enhance_flow_sheet(ws):
    """流程程序图美化 (20 cols, B-T)"""
    # 标题
    for c in range(2, 21):
        cell = ws.cell(row=2, column=c)
        cell.font = Font(name="微软雅黑", size=16, bold=True, color=STEEL_BLUE)
    ws.row_dimensions[2].height = 30

    # 信息栏 R3-R5 (left side)
    for r in (3, 4, 5):
        for c in (2, 3):
            cell = ws.cell(row=r, column=c)
            cell.font = Font(name="微软雅黑", size=9)
            cell.border = BORDER
            if cell.value and str(cell.value).endswith(("号", "员", "力")):
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.font = Font(name="微软雅黑", size=9, bold=True)

    # 表头 R3-R4（统计计数在合计行，由 build_report.py 写入）
    for r in (3, 4):
        for c in range(2, 21):
            cell = ws.cell(row=r, column=c)
            cell.font = Font(name="微软雅黑", size=9, bold=True, color=WHITE)
            cell.fill = PatternFill("solid", fgColor=STEEL_BLUE)
            cell.alignment = CENTER
            cell.border = BORDER
    ws.row_dimensions[3].height = 30
    ws.row_dimensions[4].height = 24

    # 数据行 R5+
    last_data = ws.max_row
    for r in range(last_data, 4, -1):
        v = ws.cell(row=r, column=3).value
        if v and str(v) == "合计":
            last_data = r - 1
            break

    for r in range(5, last_data + 1):
        bg = LIGHT_GRAY if (r % 2 == 0) else None
        for c in range(2, 21):
            cell = ws.cell(row=r, column=c)
            cell.font = Font(name="微软雅黑", size=9)
            cell.border = BORDER
            if c in (3, 14):
                cell.alignment = LEFT
            else:
                cell.alignment = CENTER
            if bg:
                cell.fill = PatternFill("solid", fgColor=bg)

        # Symbol columns F-J (6-10)
        for c in range(6, 11):
            cell = ws.cell(row=r, column=c)
            if cell.value:
                cell.font = Font(name="微软雅黑", size=12, bold=True, color=STEEL_BLUE)
                cell.alignment = CENTER

        # Value columns K-M (11-13)
        for c in range(11, 14):
            cell = ws.cell(row=r, column=c)
            if cell.value:
                v_char = str(cell.value).strip()
                if v_char == "◎":
                    fg, fc = VALUE_FILLS["V"]
                elif v_char == "●":
                    fg, fc = VALUE_FILLS["N"]
                elif v_char == "○":
                    fg, fc = VALUE_FILLS["M"]
                else:
                    continue
                cell.fill = PatternFill("solid", fgColor=fg)
                cell.font = Font(name="微软雅黑", size=11, bold=True, color=fc)
                cell.alignment = CENTER

        # ECRS columns O-R (15-18)
        for c in range(15, 19):
            cell = ws.cell(row=r, column=c)
            if cell.value:
                cell.fill = PatternFill("solid", fgColor=ECRS_FILL[0])
                cell.font = Font(name="微软雅黑", size=9, bold=True, color=ECRS_FILL[1])
                cell.alignment = CENTER

        ws.row_dimensions[r].height = 34

    # 合计行
    if last_data < ws.max_row:
        r = last_data + 1
        for c in range(2, 21):
            cell = ws.cell(row=r, column=c)
            cell.font = Font(name="微软雅黑", size=9, bold=True)
            cell.fill = PatternFill("solid", fgColor="D6E4F0")
            cell.border = Border(left=THIN, right=THIN, top=Side(style="medium", color=STEEL_BLUE), bottom=THIN)
            cell.alignment = CENTER
        ws.row_dimensions[r].height = 28

    # 图例列 U (21)
    for r in range(3, 11):
        cell = ws.cell(row=r, column=21)
        if cell.value:
            cell.font = Font(name="微软雅黑", size=9)
            cell.alignment = CENTER


def set_page_setup(ws):
    """A4 横向、fitToPage、水平居中、窄边距"""
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins.left = 0.3
    ws.page_margins.right = 0.3
    ws.page_margins.top = 0.4
    ws.page_margins.bottom = 0.4
    ws.print_options.horizontalCentered = True


def main():
    if len(sys.argv) < 2:
        print("用法: python style_enhance.py <xlsx 文件>")
        sys.exit(1)

    path = Path(sys.argv[1])
    wb = load_workbook(str(path))

    for sn in wb.sheetnames:
        ws = wb[sn]
        if "问题" in sn:
            enhance_problem_sheet(ws)
        elif "流程" in sn:
            enhance_flow_sheet(ws)
        set_page_setup(ws)

    wb.save(str(path))
    print(f"美化完成: {path}")


if __name__ == "__main__":
    main()
