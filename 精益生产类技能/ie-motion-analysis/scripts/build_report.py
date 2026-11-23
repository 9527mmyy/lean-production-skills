#!/usr/bin/env python3
"""
动作分析交付物生成器（IE 动作分析 Step 6）

从结构化分析 JSON 一次性生成：
  1. <name>_动作分析表单.xlsx —— 问题项检查表 + 流程程序图（数据区 19 列 B-T + U 列图例，合计行 12 类计数与 SUM）
  2. <name>_流程程序图.html  —— 完整 IE 分析报告（7 章节，浏览器查看）

用法:
  python build_report.py --data analysis.json --out-dir OUT --name "安装板焊接"

字段定义见 references/output-schema.md；符号标准见 references/flow-chart-standard.md。
"""

import argparse
import html
import json
import sys
from collections import Counter
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.stderr.write("[错误] 需要 openpyxl：pip install openpyxl\n")
    sys.exit(2)

# ---------- 代码映射 ----------
SYMBOL_CODE = {
    "O": ("operation", "○", "F"), "T": ("transport", "→", "G"),
    "D": ("delay", "Ｄ", "H"), "I": ("inspection", "□", "I"),
    "S": ("storage", "▽", "J"),
}
SYMBOL_ALT = {
    "O": "O", "T": "T", "D": "D", "I": "I", "S": "S",
    "operation": "O", "操作": "O", "○": "O",
    "transport": "T", "搬运": "T", "→": "T",
    "delay": "D", "暂存": "D", "Ｄ": "D", "D_FULL": "D",
    "inspection": "I", "检查": "I", "□": "I",
    "storage": "S", "储存": "S", "▽": "S",
}
VALUE_CODE = {
    "V": ("value_added", "◎", "K"), "N": ("non_value", "●", "L"),
    "M": ("necessary", "○", "M"),
}
VALUE_ALT = {
    "V": "V", "N": "N", "M": "M",
    "value_added": "V", "有价值": "V", "◎": "V",
    "non_value": "N", "无价值": "N", "●": "N",
    "necessary": "M", "无价值但必须": "M", "必要": "M", "○": "M",
}
ECRS_NAME = {"E": "取消 Eliminate", "C": "合并 Combine",
             "R": "重排 Rearrange", "S": "简化 Simplify"}
ECRS_ORDER = ["E", "C", "R", "S"]
ECRS_COL = {"E": "O", "C": "P", "R": "Q", "S": "R"}

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
HDR_FILL = PatternFill("solid", fgColor="D9E1F2")
TITLE_FILL = PatternFill("solid", fgColor="F2F2F2")
VALUE_FILLS = {"V": PatternFill("solid", fgColor="E2EFDA"),
               "N": PatternFill("solid", fgColor="FCE4D6"),
               "M": PatternFill("solid", fgColor="FFF2CC")}
HTML_COLORS = {"V": ("#e8f5e9", "#1b5e20", "有价值 ◎"),
               "N": ("#fdecea", "#b71c1c", "无价值 ●"),
               "M": ("#fff8e1", "#8d6e00", "无价值但必须 ○")}


def norm(code, alt_map, default):
    if code is None:
        return default
    key = str(code).strip()
    if key in alt_map:
        return alt_map[key]
    return default


def load_data(path):
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    for key in ("meta", "actions", "problems"):
        if key not in data:
            raise ValueError(f"分析数据缺少顶层键：{key}")
    return data


def style_range(ws, cell_range, border=True, fill=None, alignment=None, font=None):
    for row in ws[cell_range]:
        for c in row:
            if border:
                c.border = BORDER
            if fill:
                c.fill = fill
            if alignment:
                c.alignment = alignment
            if font:
                c.font = font


# ============================================================
# 1. Excel (20 columns, B-T)
# ============================================================

def build_xlsx(data, out_path, name):
    meta = data["meta"]
    actions = data["actions"]
    problems = data["problems"]
    wb = Workbook()

    # ---------- Sheet 1: 问题项检查表 ----------
    ws = wb.active
    ws.title = "问题项检查表"
    ws.merge_cells("A1:L1")
    ws["A1"] = "动作分析问题项检查表"
    ws["A1"].font = Font(size=15, bold=True)
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 28

    info = [
        ("A2", "分析对象/工位：", "B2", meta.get("station", ""),
         "E2", "作业名称：", "F2", meta.get("operation", "") or name,
         "I2", "分析日期：", "J2", meta.get("date", "")),
        ("A3", "分析人：", "B3", meta.get("researcher", ""),
         "E3", "所属部门：", "F3", meta.get("department", ""),
         "I3", "观测周期：", "J3", meta.get("cycle", "")),
    ]
    for l1, t1, v1, val1, l2, t2, v2, val2, l3, t3, v3, val3 in info:
        for (lc, txt), (vc, val) in [((l1, t1), (v1, val1)), ((l2, t2), (v2, val2)), ((l3, t3), (v3, val3))]:
            ws[lc] = txt
            ws[lc].font = Font(bold=True)
            ws.merge_cells(f"{vc}:{get_column_letter(ws[vc].column + 2)}{ws[vc].row}")
            ws[vc] = val
            ws[vc].alignment = LEFT

    headers = ["序号", "作业/动作描述", "动素", "问题类别", "问题描述",
               "频次\n(次/班)", "单次耗时\n(秒)", "浪费类型", "严重度",
               "责任人", "改善状态", "备注"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=4, column=i, value=h)
        c.font = Font(bold=True)
        c.fill = HDR_FILL
        c.alignment = CENTER
        c.border = BORDER
    ws.row_dimensions[4].height = 34

    for r, p in enumerate(problems, 5):
        vals = [p.get("seq", r - 4), p.get("desc", ""), p.get("therblig", ""),
                p.get("category", ""), p.get("problem", ""), p.get("freq", ""),
                p.get("unit_time", ""), p.get("waste_type", ""), p.get("severity", ""),
                p.get("owner", "待补充"), p.get("status", "未改善"), p.get("remark", "")]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
            cell.alignment = LEFT if c in (2, 3, 4, 5, 8, 12) else CENTER
        ws.cell(row=r, column=6).number_format = "0"
        ws.cell(row=r, column=7).number_format = "0.00"
        ws.row_dimensions[r].height = 30

    last = max(5, 5 + len(problems))
    for start, end, text in [(0, 2, "分析人签字："), (3, 5, "班组长确认："),
                             (6, 8, "改善负责人："), (9, 11, "日期：")]:
        ws.merge_cells(start_row=last, start_column=start + 1,
                       end_row=last, end_column=end + 1)
        c = ws.cell(row=last, column=start + 1, value=text)
        c.alignment = CENTER
        c.font = Font(bold=True)
        c.fill = TITLE_FILL
        c.border = BORDER

    widths = [6, 20, 16, 14, 34, 10, 12, 18, 9, 10, 11, 16]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A5"

    # ---------- Sheet 2: 流程程序图 (20 cols, B-T) ----------
    ps = wb.create_sheet("流程程序图")
    ps.merge_cells("B2:T2")
    ps["B2"] = f"{meta.get('product', '') or name}产品流程程序图{' ' * 20}时间：{meta.get('date', '')}"
    ps["B2"].font = Font(size=14, bold=True)
    ps["B2"].alignment = CENTER
    ps.row_dimensions[2].height = 26

    def kv(label_cell, label, value_cell, value):
        ps[label_cell] = label
        ps[label_cell].font = Font(bold=True)
        ps[value_cell] = value
        ps[value_cell].alignment = LEFT

    kv("B3", "图号", "C3", meta.get("drawing_no", ""))
    kv("B4", "研究人员", "C4", meta.get("researcher", ""))
    kv("B5", "人力", "C5", meta.get("manpower", ""))

    # 统计计数：统一在合计行输出（符号 F-J / 价值 K-M / ECRS O-R）。
    # 不要在 row3 单独写统计带——row3 是两级表头的上半部分，会被下面的合并覆盖。
    sym_counts = Counter()
    val_counts = Counter()
    ecrs_counts = Counter()
    for a in actions:
        sym_counts[norm(a.get("symbol"), SYMBOL_ALT, "O")] += 1
        val_counts[norm(a.get("value"), VALUE_ALT, "M")] += 1
        e = str(a.get("ecrs") or "").strip().upper()
        if e in ECRS_ORDER:
            ecrs_counts[e] += 1

    # 两级表头 R3-R4
    ps.merge_cells("F3:J3"); ps["F3"] = "符号"
    ps.merge_cells("O3:R3"); ps["O3"] = "改善方法"
    for col, title in [(2, "序号"), (3, "作业内容"), (4, "搬运距离（m）"), (5, "时间（S）"),
                       (11, "有价值作业（◎）"), (12, "无价值作业（●）"),
                       (13, "无价值但必须进行作业（○）"), (19, "改进建议"), (20, "备注")]:
        ps.merge_cells(start_row=3, start_column=col, end_row=4, end_column=col)
        ps.cell(row=3, column=col, value=title)
    for col, title in [(6, "操作（○）"), (7, "搬运（→）"), (8, "暂存（Ｄ）"),
                       (9, "检查（□）"), (10, "储存（▽）"),
                       (15, "取消（E）"), (16, "合并（C）"), (17, "重排（R）"), (18, "简化（S）")]:
        ps.cell(row=4, column=col, value=title)
    ps.merge_cells("N3:N4"); ps["N3"] = "问题描述"
    style_range(ps, "B3:T4", fill=HDR_FILL, alignment=CENTER, font=Font(bold=True, size=9))
    ps.row_dimensions[3].height = 30
    ps.row_dimensions[4].height = 24

    # 数据行 (start R5)
    start_row = 5
    for r, a in enumerate(actions, start_row):
        sc = norm(a.get("symbol"), SYMBOL_ALT, "O")
        vc = norm(a.get("value"), VALUE_ALT, "M")
        ec = str(a.get("ecrs") or "").strip().upper()
        if ec not in ECRS_ORDER:
            ec = ""
        scol = SYMBOL_CODE[sc][2]
        vcol = VALUE_CODE[vc][2]
        row_vals = {
            2: a.get("seq", r - start_row + 1), 3: a.get("content", ""),
            4: a.get("distance", 0) or 0, 5: a.get("time", 0) or 0,
            14: a.get("problem", ""),
            19: a.get("suggestion", ""),
            20: a.get("remark", ""),
        }
        for col, val in row_vals.items():
            cell = ps.cell(row=r, column=col, value=val)
            cell.alignment = LEFT if col in (3, 14, 19, 20) else CENTER
        ps.cell(row=r, column=ord(scol) - 64, value=SYMBOL_CODE[sc][1]).alignment = CENTER
        ps.cell(row=r, column=ord(vcol) - 64, value=VALUE_CODE[vc][1]).alignment = CENTER
        ps.cell(row=r, column=ord(vcol) - 64).fill = VALUE_FILLS[vc]
        if ec:
            ps.cell(row=r, column=ord(ECRS_COL[ec]) - 64, value=ec).alignment = CENTER
        style_range(ps, f"B{r}:T{r}", alignment=CENTER)
        for col in (3, 14, 19, 20):
            ps.cell(row=r, column=col).alignment = LEFT
        ps.row_dimensions[r].height = 34

    # 合计行
    total_row = start_row + len(actions)
    ps.cell(row=total_row, column=3, value="合计").font = Font(bold=True)
    ps.cell(row=total_row, column=4, value=f"=SUM(D{start_row}:D{total_row - 1})")
    ps.cell(row=total_row, column=5, value=f"=SUM(E{start_row}:E{total_row - 1})")
    # 12 类统计计数：符号 F-J / 价值 K-M / ECRS O-R，与 flow-chart-standard.md 一致
    for i, code in enumerate(["O", "T", "D", "I", "S"]):
        ps.cell(row=total_row, column=6 + i, value=sym_counts.get(code, 0))
    for i, code in enumerate(["V", "N", "M"]):
        ps.cell(row=total_row, column=11 + i, value=val_counts.get(code, 0))
    for i, code in enumerate(ECRS_ORDER):
        ps.cell(row=total_row, column=15 + i, value=ecrs_counts.get(code, 0))
    style_range(ps, f"B{total_row}:T{total_row}", fill=TITLE_FILL, alignment=CENTER, font=Font(bold=True))

    # 右侧图例 (column U)
    legend = ["○操作", "→搬运", "Ｄ暂存", "□检查", "▽储存", "◎有价值", "●无价值", "○无价值但必须"]
    for i, t in enumerate(legend, 3):
        ps.cell(row=i, column=21, value=t).alignment = CENTER

    widths2 = {"A": 2, "B": 6, "C": 24, "D": 11, "E": 9}
    for col in range(6, 11):
        widths2[get_column_letter(col)] = 7
    for col in range(11, 14):
        widths2[get_column_letter(col)] = 8
    widths2["N"] = 28
    for col in range(15, 19):
        widths2[get_column_letter(col)] = 7
    for col in (19, 20):
        widths2[get_column_letter(col)] = 14
    widths2["U"] = 16
    for col, w in widths2.items():
        ps.column_dimensions[col].width = w
    ps.freeze_panes = "C5"

    wb.save(out_path)
    return {
        "total_distance": sum(float(a.get("distance", 0) or 0) for a in actions),
        "total_time": sum(float(a.get("time", 0) or 0) for a in actions),
        "sym_counts": sym_counts, "val_counts": val_counts, "ecrs_counts": ecrs_counts,
    }


# ============================================================
# 2. HTML 完整报告 (7 chapters)
# ============================================================

def build_html(data, stats, out_path, name):
    meta = data["meta"]
    actions = data["actions"]
    problems = data["problems"]
    sc = stats["sym_counts"]
    vc = stats["val_counts"]
    ec = stats["ecrs_counts"]
    total_time = stats["total_time"]
    total_distance = stats["total_distance"]
    n = len(actions)
    pct = lambda k: (vc[k] / n * 100) if n else 0.0

    # ---- 一、现状分析 ----
    waste_ratio = pct("N")
    waste_kinds = []
    for p in problems:
        wt = str(p.get("waste_type", "")).strip()
        if wt and wt not in waste_kinds:
            waste_kinds.append(wt)
    focus = "、".join(waste_kinds[:4]) if waste_kinds else "无价值动作"

    # ---- 二、问题项清单 ----
    problem_rows = ""
    sev_colors = {"高": ("#fdecea", "#b71c1c"), "中": ("#fff8e1", "#8d6e00"), "低": ("#e8f5e9", "#1b5e20")}
    for p in problems:
        sev = p.get("severity", "")
        bg, fg = sev_colors.get(sev, ("#f5f5f5", "#333"))
        problem_rows += f"""
        <tr>
          <td class="c">{p.get('seq','')}</td>
          <td>{html.escape(str(p.get('desc','')))}</td>
          <td>{html.escape(str(p.get('therblig','')))}</td>
          <td>{html.escape(str(p.get('category','')))}</td>
          <td>{html.escape(str(p.get('problem','')))}</td>
          <td class="c">{p.get('freq','')}</td>
          <td class="c">{float(p.get('unit_time',0) or 0):g}</td>
          <td>{html.escape(str(p.get('waste_type','')))}</td>
          <td class="c" style="background:{bg};color:{fg};font-weight:700">{sev}</td>
          <td>{html.escape(str(p.get('status','未改善')))}</td>
        </tr>"""

    # ---- 三、浪费诊断 ----
    waste_actions = [a for a in actions if a.get("problem")]
    waste_rows = ""
    for a in waste_actions:
        vcode = norm(a.get("value"), VALUE_ALT, "M")
        bg, fg, vlabel = HTML_COLORS[vcode]
        principle = a.get("principle", "")
        if not principle:
            sym = a.get("symbol")
            if sym in ("T",):
                principle = "物料搬运：搬运距离应最短，避免物流交叉"
            elif sym in ("D",):
                principle = "身体使用：避免可避免的等待迟延"
            else:
                principle = "工作地布置/工具设备：应用工装夹具与机械化降低手工作业"
        waste_rows += f"""
        <tr style="background:{bg}">
          <td class="c">{a.get('seq','')}</td>
          <td>{html.escape(str(a.get('content','')))}</td>
          <td class="c" style="color:{fg};font-weight:700">{vlabel}</td>
          <td>{html.escape(str(a.get('problem','')))}</td>
          <td>{html.escape(principle)}</td>
        </tr>"""

    # ---- 四、ECRS 改善建议 ----
    improved = [a for a in actions if a.get("ecrs") and a.get("suggestion")]
    ecrs_html = ""
    ecrs_labels = {"E": ("取消 Eliminate", "#fdecea", "#b71c1c"),
                   "C": ("合并 Combine", "#fff8e1", "#8d6e00"),
                   "R": ("重排 Rearrange", "#e3f2fd", "#0d47a1"),
                   "S": ("简化 Simplify", "#e8f5e9", "#1b5e20")}
    for code in ECRS_ORDER:
        group = [a for a in improved if str(a.get("ecrs", "")).strip().upper() == code]
        if not group:
            continue
        lab, bg, fg = ecrs_labels[code]
        ecrs_html += f"""
        <div class="ecrs-group" style="border-left:4px solid {fg}">
          <h3 style="color:{fg}">{lab}（{len(group)} 项）</h3>
          <ul>{''.join(f'<li><b>序号 {a.get("seq","")} · {html.escape(str(a.get("content","")))}</b>：{html.escape(str(a.get("suggestion","")))}</li>' for a in group)}</ul>
        </div>"""

    # ---- 五、预期效果 ----
    saved = sum(float(a.get("estimated_saving", 0) or 0) for a in improved)
    effect_items = []
    if saved:
        effect_items.append(f"按已量化改善项估算，可节省作业时间约 <b>{saved:.2f} 秒/循环</b>"
                           f"（占当前总时间 {saved / total_time * 100:.1f}%）" if total_time else "")
    effect_items.append(f"无价值作业 {vc['N']} 项、占比 {pct('N'):.1f}%，"
                        f"经取消/合并/重排/简化后可显著压缩。")
    if total_distance:
        effect_items.append(f"搬运距离合计 {total_distance:.2f} m，可通过就近定置与布局重排进一步缩短取放往返。")
    erg_types = [p.get("waste_type", "") for p in problems if "人机" in str(p.get("waste_type", ""))
                 or "体力" in str(p.get("waste_type", "")) or "疲劳" in str(p.get("category", ""))]
    if erg_types:
        effect_items.append("人机工程改善（双手对称作业、专用夹持/定位工装、工位合理布置）可降低单手负荷与疲劳，提升作业稳定性。")
    effect_html = "\n".join(f"<li>{e}</li>" for e in effect_items if e)

    # ---- 六、结论 ----
    applied = [c for c in ECRS_ORDER if ec[c] > 0]
    applied_text = "、".join(f"「{ECRS_NAME[c].split()[0]}」" for c in applied) if applied else "ECRS"
    conclusion = (f"本报告基于视频逐帧拆分形成 {n} 项动作记录，识别问题项 {len(problems)} 项，"
                  f"提出 ECRS 改善建议 {len(improved)} 项（涉及 {applied_text}）。"
                  f"建议按 ECRS 优先级，结合动作经济原则落实具体工装与定置措施，"
                  f"并按《问题项检查表》跟踪责任人与改善状态，形成闭环。")

    # ---- 七、流程程序图明细 ----
    rows = []
    for a in actions:
        scode = norm(a.get("symbol"), SYMBOL_ALT, "O")
        vcode = norm(a.get("value"), VALUE_ALT, "M")
        ecode = str(a.get("ecrs") or "").strip().upper()
        if ecode not in ECRS_ORDER:
            ecode = ""
        bg, fg, vlabel = HTML_COLORS[vcode]
        sym = SYMBOL_CODE[scode][1]
        symbol_name = SYMBOL_CODE[scode][0]
        rows.append(f"""
        <tr style="background:{bg}">
          <td class="c">{a.get('seq','')}</td>
          <td>{html.escape(str(a.get('content','')))}</td>
          <td class="c">{float(a.get('distance',0) or 0):g}</td>
          <td class="c">{float(a.get('time',0) or 0):g}</td>
          <td class="c symbol" title="{symbol_name}">{sym}</td>
          <td class="c" style="color:{fg};font-weight:700">{vlabel}</td>
          <td>{html.escape(str(a.get('problem','')))}</td>
          <td class="c">{html.escape(ecode)}</td>
          <td>{html.escape(str(a.get('suggestion','')))}</td>
        </tr>""")

    stat_cards = "".join(
        f'<div class="card"><div class="num">{v}</div><div class="lab">{lab}</div></div>'
        for lab, v in [
            ("操作 ○", sc["O"]), ("搬运 →", sc["T"]), ("暂存 Ｄ", sc["D"]),
            ("检查 □", sc["I"]), ("储存 ▽", sc["S"]), ("有价值 ◎", vc["V"]),
            ("无价值 ●", vc["N"]), ("无价值但必须 ○", vc["M"]),
            ("取消 E", ec["E"]), ("合并 C", ec["C"]), ("重排 R", ec["R"]), ("简化 S", ec["S"]),
        ])

    doc = f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{html.escape(name)} — IE 动作分析与改善报告</title>
<style>
  :root {{ --ink:#1f2933; --muted:#6b7280; --line:#d9dee5; --accent:#1f4e79;
           --panel:#ffffff; --bg:#f4f6f9; }}
  * {{ box-sizing:border-box; }}
  body {{ margin:0; background:var(--bg); color:var(--ink);
          font-family:"Microsoft YaHei","PingFang SC",system-ui,sans-serif; font-size:14px; }}
  .wrap {{ max-width:1280px; margin:0 auto; padding:28px 20px 60px; }}
  h1 {{ font-size:24px; margin:0 0 6px; color:var(--accent); }}
  h2 {{ font-size:18px; margin:28px 0 10px; color:var(--accent); border-bottom:2px solid var(--line); padding-bottom:6px; }}
  h3 {{ font-size:15px; margin:16px 0 8px; }}
  .sub {{ color:var(--muted); margin-bottom:20px; }}
  .meta {{ display:grid; grid-template-columns:repeat(4,1fr); gap:8px 18px;
           background:var(--panel); border:1px solid var(--line); border-radius:8px;
           padding:14px 18px; margin-bottom:18px; }}
  .meta div {{ display:flex; gap:8px; }}
  .meta b {{ color:var(--muted); font-weight:600; min-width:70px; }}
  .cards {{ display:grid; grid-template-columns:repeat(12,1fr); gap:8px; margin-bottom:18px; }}
  .card {{ background:var(--panel); border:1px solid var(--line); border-radius:8px;
           text-align:center; padding:10px 4px; }}
  .card .num {{ font-size:20px; font-weight:800; color:var(--accent); }}
  .card .lab {{ font-size:11px; color:var(--muted); margin-top:3px; }}
  table {{ width:100%; border-collapse:collapse; background:var(--panel);
           border:1px solid var(--line); border-radius:8px; overflow:hidden; margin-bottom:18px; }}
  th,td {{ border:1px solid var(--line); padding:7px 8px; vertical-align:top; }}
  th {{ background:#e8eef5; color:#1f3a5f; font-size:12px; position:sticky; top:0; }}
  td.c {{ text-align:center; vertical-align:middle; }}
  td.symbol {{ font-size:18px; font-weight:700; color:var(--accent); }}
  tr.total td {{ background:#eef2f7; font-weight:700; }}
  .legend {{ margin-top:14px; color:var(--muted); font-size:12px; }}
  .legend span {{ display:inline-block; margin-right:16px; padding:2px 8px; border-radius:4px; }}
  .ecrs-group {{ background:var(--panel); border-radius:8px; padding:12px 16px; margin-bottom:12px; }}
  .ecrs-group ul {{ margin:8px 0 0 20px; }}
  .ecrs-group li {{ margin-bottom:4px; }}
  @media(max-width:900px){{ .meta{{grid-template-columns:repeat(2,1fr)}} .cards{{grid-template-columns:repeat(4,1fr)}} }}
  @media print {{ body{{background:#fff}} .card,table{{break-inside:avoid}} }}
</style></head><body><div class="wrap">
  <h1>{html.escape(meta.get('product','') or name)} · IE 动作分析与改善报告</h1>
  <div class="sub">图号 {html.escape(str(meta.get('drawing_no','待补充')))} ｜
    分析人 {html.escape(str(meta.get('researcher','待补充')))} ｜
    日期 {html.escape(str(meta.get('date','待补充')))} ｜
    人力 {html.escape(str(meta.get('manpower','—')))}</div>
  <div class="meta">
    <div><b>工位</b>{html.escape(str(meta.get('station','待补充')))}</div>
    <div><b>作业</b>{html.escape(str(meta.get('operation','待补充')))}</div>
    <div><b>总距离</b>{total_distance:g} m</div>
    <div><b>总时间</b>{total_time:g} s（{total_time/60:.2f} min）</div>
    <div><b>动作数</b>{n}</div>
    <div><b>问题项</b>{len(problems)}</div>
    <div><b>观测周期</b>{html.escape(str(meta.get('cycle','—')))}</div>
    <div><b>部门</b>{html.escape(str(meta.get('department','—')))}</div>
  </div>

  <h2>一、现状分析（As-Is）</h2>
  <p>本次分析共拆分动作 <b>{n}</b> 项，累计作业时间 <b>{total_time:.2f} 秒</b>
     （约 {total_time/60:.2f} 分钟），累计搬运距离 <b>{total_distance:.2f} 米</b>。</p>
  <table>
    <thead><tr><th>价值分类</th><th>数量</th><th>占比</th><th>说明</th></tr></thead>
    <tbody>
      <tr><td>有价值作业 ◎</td><td class="c">{vc['V']}</td><td class="c">{pct('V'):.1f}%</td><td>直接创造产品价值</td></tr>
      <tr><td>无价值作业 ●</td><td class="c">{vc['N']}</td><td class="c">{pct('N'):.1f}%</td><td>浪费，优先改善</td></tr>
      <tr><td>无价值但必须 ○</td><td class="c">{vc['M']}</td><td class="c">{pct('M'):.1f}%</td><td>当前条件下不可避免</td></tr>
    </tbody>
  </table>
  <p>按流程符号统计：操作 ○ {sc['O']} 项、搬运 → {sc['T']} 项、暂存 Ｄ {sc['D']} 项、
     检查 □ {sc['I']} 项、储存 ▽ {sc['S']} 项。</p>
  <p><b>现状结论</b>：无价值作业占比 <b>{waste_ratio:.1f}%</b>，
     主要浪费集中在{focus}，具备通过 ECRS 与动作经济原则改善的空间。</p>

  <h2>二、问题项清单</h2>
  <table>
    <thead><tr><th>序号</th><th>作业/动作</th><th>动素</th><th>问题类别</th><th>问题描述</th>
      <th>频次(次/班)</th><th>单次耗时(s)</th><th>浪费类型</th><th>严重度</th><th>状态</th></tr></thead>
    <tbody>{problem_rows}</tbody>
  </table>

  <h2>三、浪费诊断与动作经济原则对照</h2>
  <table>
    <thead><tr><th>序号</th><th>作业内容</th><th>价值</th><th>浪费/问题</th><th>违反的动作经济原则</th></tr></thead>
    <tbody>{waste_rows}</tbody>
  </table>

  <h2>四、改善建议（按 ECRS 分类）</h2>
  {ecrs_html or '<p>（未标注 ECRS 改善项）</p>'}

  <h2>五、预期效果</h2>
  <ul>{effect_html}</ul>

  <h2>六、结论</h2>
  <p>{conclusion}</p>

  <h2>七、流程程序图明细</h2>
  <div class="cards">{stat_cards}</div>
  <table>
    <thead><tr><th>序号</th><th>作业内容</th><th>距离(m)</th><th>时间(S)</th>
      <th>符号</th><th>价值分类</th><th>问题描述</th><th>ECRS</th><th>改进建议</th></tr></thead>
    <tbody>{''.join(rows)}
      <tr class="total"><td colspan="2">合计</td><td class="c">{total_distance:g}</td>
        <td class="c">{total_time:g}</td><td colspan="5"></td></tr>
    </tbody>
  </table>
  <div class="legend">
    <span style="background:#e8f5e9;color:#1b5e20">◎ 有价值作业</span>
    <span style="background:#fdecea;color:#b71c1c">● 无价值作业（浪费）</span>
    <span style="background:#fff8e1;color:#8d6e00">○ 无价值但必须</span>
    <span>○操作 →搬运 Ｄ暂存 □检查 ▽储存</span>
    <span>E取消 C合并 R重排 S简化</span>
  </div>
</div></body></html>"""
    out_path.write_text(doc, encoding="utf-8")


def main():
    ap = argparse.ArgumentParser(description="动作分析交付物生成器")
    ap.add_argument("--data", required=True, help="分析数据 JSON")
    ap.add_argument("--out-dir", required=True, help="输出目录")
    ap.add_argument("--name", default=None, help="作业/产品名称，默认取 meta.product")
    ap.add_argument("--template", default=None, help="标准表单模板路径（仅作参考校验）")
    args = ap.parse_args()

    data = load_data(args.data)
    name = args.name or data["meta"].get("product") or "作业"
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    xlsx = out_dir / f"{name}_动作分析表单.xlsx"
    htm = out_dir / f"{name}_流程程序图.html"

    stats = build_xlsx(data, xlsx, name)
    build_html(data, stats, htm, name)

    print(json.dumps({
        "xlsx": str(xlsx), "html": str(htm),
        "actions": len(data["actions"]), "problems": len(data["problems"]),
        "total_time_sec": round(stats["total_time"], 2),
        "total_distance_m": round(stats["total_distance"], 2),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
