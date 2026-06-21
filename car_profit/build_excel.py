#!/usr/bin/env python3
"""KIA 销售提成计算表 (Excel, 带公式).

提成 = 提成比例 × ( LUX Care份数×单价 + Greenway Advantage + 车架利润 + Trade差价 )

Run:  python3 build_excel.py
Out:  KIA提成表.xlsx
"""
from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

OUT = "KIA提成表.xlsx"
DATA_ROWS = 60

# ---- styling ---------------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")    # 黄 = 自己填
CALC_FILL = PatternFill("solid", fgColor="E2EFDA")     # 绿 = 自动算
SETTING_FILL = PatternFill("solid", fgColor="FCE4D6")  # 橙 = 参数(可改)
TITLE_FONT = Font(bold=True, size=16, color="1F4E78")
LABEL_FONT = Font(bold=True, size=11)
MONEY = '#,##0.00'
PCT = '0%'
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

# table columns: (header, key, kind, width)
COLUMNS = [
    ("编号", "id", "input", 8),
    ("客户 / 车型", "cust", "input", 20),
    ("LUX Care 份数", "lux", "input", 12),
    ("Greenway\n(是/否)", "gw", "input", 11),
    ("车架利润", "front", "input", 12),
    ("Trade 实际价值", "tval", "input", 13),
    ("Trade 折抵价", "tgive", "input", 12),
    ("Trade 差价", "tspread", "calc", 11),
    ("总利润(店)", "gross", "calc", 12),
    ("我的提成", "comm", "calc", 13),
]
COL = {c[1]: i + 1 for i, c in enumerate(COLUMNS)}


def cl(key: str) -> str:
    return get_column_letter(COL[key])


def build() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "提成"

    # title
    ws.merge_cells("A1:J1")
    ws["A1"] = "🚗 KIA 销售提成计算表"
    ws["A1"].font = TITLE_FONT
    ws.row_dimensions[1].height = 26
    ws.merge_cells("A2:J2")
    ws["A2"] = ("黄色=自己填  |  橙色=参数(可改)  |  绿色=自动算  |  "
                "没有 Trade 就把 Trade 两列留空")
    ws["A2"].font = Font(italic=True, size=10, color="808080")

    # ---- settings box (rows 4-6): label merged A:B, value in C ----------
    settings = [
        ("LUX Care 单份利润 ($)", 50, MONEY),
        ("Greenway Advantage ($)", 2995, MONEY),
        ("我的提成比例", 0.25, PCT),
    ]
    LUX_PRICE = "$C$4"
    GW_PRICE = "$C$5"
    RATE = "$C$6"
    for i, (label, val, fmt) in enumerate(settings):
        r = 4 + i
        ws.merge_cells(f"A{r}:B{r}")
        lc = ws.cell(row=r, column=1, value=label)
        lc.font = LABEL_FONT
        lc.alignment = Alignment(horizontal="right", vertical="center")
        vc = ws.cell(row=r, column=3, value=val)
        vc.number_format = fmt
        vc.fill = SETTING_FILL
        vc.border = BORDER
        vc.font = Font(bold=True)
        vc.alignment = CENTER

    header_row = 8
    first = header_row + 1
    last = first + DATA_ROWS - 1

    # header
    for c in COLUMNS:
        ci = COL[c[1]]
        cell = ws.cell(row=header_row, column=ci, value=c[0])
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(ci)].width = c[3]
    ws.row_dimensions[header_row].height = 32

    # data rows
    for r in range(first, last + 1):
        for c in COLUMNS:
            key = c[1]
            cell = ws.cell(row=r, column=COL[key])
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if key == "tspread":
                cell.value = (f'=IF(OR({cl("tval")}{r}="",{cl("tgive")}{r}=""),0,'
                              f'{cl("tval")}{r}-{cl("tgive")}{r})')
                cell.number_format = MONEY
                cell.fill = CALC_FILL
            elif key == "gross":
                cell.value = (
                    f'=IF(COUNTA(A{r}:{cl("tgive")}{r})=0,"",'
                    f'N({cl("lux")}{r})*{LUX_PRICE}'
                    f'+IF({cl("gw")}{r}="是",{GW_PRICE},0)'
                    f'+N({cl("front")}{r})+{cl("tspread")}{r})')
                cell.number_format = MONEY
                cell.fill = CALC_FILL
            elif key == "comm":
                cell.value = (f'=IF({cl("gross")}{r}="","",'
                              f'{cl("gross")}{r}*{RATE})')
                cell.number_format = MONEY
                cell.fill = CALC_FILL
                cell.font = Font(bold=True)
            else:
                cell.fill = INPUT_FILL
                if key in ("front", "tval", "tgive"):
                    cell.number_format = MONEY

    # dropdown 是/否 for Greenway
    dv = DataValidation(type="list", formula1='"是,否"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f'{cl("gw")}{first}:{cl("gw")}{last}')

    # ---- summary --------------------------------------------------------
    s = last + 2
    ws.cell(row=s, column=1, value="📊 汇总").font = Font(bold=True, size=14,
                                                         color="1F4E78")

    def summ(label, formula, fmt=MONEY, off=1):
        row = s + off
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        lc = ws.cell(row=row, column=1, value=label)
        lc.font = LABEL_FONT
        lc.alignment = Alignment(horizontal="right")
        vc = ws.cell(row=row, column=3, value=formula)
        vc.number_format = fmt
        vc.font = Font(bold=True)
        vc.fill = CALC_FILL
        vc.border = BORDER

    cust = f'{cl("cust")}{first}:{cl("cust")}{last}'
    lux = f'{cl("lux")}{first}:{cl("lux")}{last}'
    gw = f'{cl("gw")}{first}:{cl("gw")}{last}'
    gross = f'{cl("gross")}{first}:{cl("gross")}{last}'
    comm = f'{cl("comm")}{first}:{cl("comm")}{last}'

    summ("成交单数", f'=COUNTA({cust})', fmt='0', off=1)
    summ("LUX Care 总份数", f'=SUM({lux})', fmt='0', off=2)
    summ("Greenway 卖出数", f'=COUNTIF({gw},"是")', fmt='0', off=3)
    summ("店总利润", f'=SUM({gross})', off=4)
    summ("我的总提成", f'=SUM({comm})', off=5)
    summ("平均每单提成",
         f'=IF(COUNTA({cust})=0,0,SUM({comm})/COUNTA({cust}))', off=6)

    ws.freeze_panes = f"A{first}"

    # example row (matches user's example -> commission 1761.25)
    ex = first
    ws.cell(row=ex, column=COL["id"], value="示例")
    ws.cell(row=ex, column=COL["cust"], value="张三 / K5")
    ws.cell(row=ex, column=COL["lux"], value=1)
    ws.cell(row=ex, column=COL["gw"], value="是")
    ws.cell(row=ex, column=COL["front"], value=2000)
    ws.cell(row=ex, column=COL["tval"], value=18000)
    ws.cell(row=ex, column=COL["tgive"], value=16000)

    wb.save(OUT)
    print(f"已生成 {OUT}")


if __name__ == "__main__":
    build()
