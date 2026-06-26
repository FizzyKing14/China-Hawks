# -*- coding: utf-8 -*-
"""
圣安地列斯州 · 案件罪名分析器  —  Excel 生成脚本 (问答 / 勾选清单版)
=====================================================================
根据《圣安地列斯州刑法典》(FiveM RP) 生成离线、纯公式的 .xlsx。
用法是「问答 + 逐条勾选」:
  ① 先回答几个问题: 要判罪几个人？受伤几人？死亡几人？
  ② 再在下方罪名清单里, 一条一条勾选每个人「做了什么」(在甲/乙/丙列填 1)
  ③ 顶部自动累加: 每人罪名清单、合计刑期(封顶5年)、罚款、建议保释金, 并指出主要责任方

用法:  python build_workbook.py
依赖:  openpyxl   (pip install openpyxl)
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation

# =====================================================================
#  罪名数据库  —  直接抄录自《圣安地列斯州刑法典》
#  每条:  (章节, 罪名, 定性, 最高刑期_月, 罚款_美元, [关键词...占位], 备注)
#  时间换算: 1 年 = 12 个月 ;  范围/分级取上限, 细节写进备注
# =====================================================================
CHARGES = [
    # ---------- 第二章 涉及人身安全的犯罪 ----------
    ("二·人身", "威胁罪", "轻罪", 12, 8000, [], "口头/书面/网络威胁他人或其亲友"),
    ("二·人身", "持有危险武器人身攻击罪", "重罪", 48, 20000, [], "用武器/工具威胁或伤害他人"),
    ("二·人身", "一级谋杀罪", "重罪", 60, 30000, [], "预谋故意杀人 / 重罪过程中杀人"),
    ("二·人身", "二级谋杀罪", "重罪", 60, 20000, [], "非预谋的故意杀人 (3至5年)"),
    ("二·人身", "三级谋杀罪", "重罪", 60, 15000, [], "无意致死/鲁莽过失致死 (3至5年, 不含正当防卫)"),
    ("二·人身", "非法拘禁罪", "重罪", 36, 15000, [], "违背意愿约束/扣押/拘禁他人"),
    ("二·人身", "故意袭击罪", "重罪", 36, 15000, [], "蓄意暴力致他人身体受伤"),
    ("二·人身", "绑架罪", "重罪", 60, 25000, [], "侵占/羁押/控制他人意图恐吓"),
    ("二·人身", "亵渎尸体罪", "重罪", 24, 20000, [], "毁损/处置尸体掩盖犯罪"),
    ("二·人身", "诈骗罪", "重罪", 60, 5000, [], "以欺诈欺瞒获取他人财物 (1-5年)"),

    # ---------- 第三章 涉及财产安全的犯罪 ----------
    ("三·财产", "非法入侵罪", "轻罪", 24, 8000, [], "被要求离开后仍拒绝离开"),
    ("三·财产", "非法入侵禁区罪", "轻罪", 36, 10000, [], "未经授权进入警局/机场/军事基地等限制区"),
    ("三·财产", "入室盗窃罪", "重罪", 36, 25000, [], "擅入住宅/办公室/仓库盗窃, 另需赔偿"),
    ("三·财产", "抢劫罪", "重罪", 60, 20000, [], "以武力/威胁强夺他人财产, 另需赔偿"),
    ("三·财产", "盗窃罪", "重罪", 48, 20000, [], "盗窃/占用他人财物, 另需赔偿"),
    ("三·财产", "盗窃车辆罪", "轻罪", 12, 5000, [], "默认未造成严重后果(轻罪1年$5000); 严重后果=重罪4年$15000, 均赔偿"),
    ("三·财产", "收受赃物罪", "轻罪", 24, 10000, [], "默认赃值<$5000(轻罪2年$10000); ≥$5000=重罪4年$20000"),
    ("三·财产", "敲诈勒索罪", "重罪", 24, 20000, [], "以威胁/恐吓/滥权强迫他人交付财物"),
    ("三·财产", "蓄意破坏罪", "轻罪", 24, 25000, [], "蓄意破坏他人财产/车辆, 另需赔偿"),
    ("三·财产", "挪用公款公物罪", "轻罪", 48, 80000, [], "擅自挪用受托管理的财物, 另需赔偿"),
    ("三·财产", "持有失窃物品罪", "轻罪", 24, 2000, [], "非法占有他人遗失/失窃财物 (1-2年), 另需赔偿"),

    # ---------- 第四章 违反公共道德的犯罪 ----------
    ("四·公德", "在公共场合从事淫秽行为罪", "轻罪", 24, 10000, [], "公开场所性行为/猥亵/招嫖 (1-2年)"),
    ("四·公德", "组织卖淫罪", "重罪", 36, 15000, [], "支持/介绍/强迫他人卖淫牟利"),
    ("四·公德", "卖淫罪", "轻罪", 24, 15000, [], "以性服务换取财物"),
    ("四·公德", "强奸罪", "重罪", 60, 30000, [], "违背意愿强迫性行为 / 与未满18岁者发生关系"),
    ("四·公德", "传播淫秽色情罪", "轻罪", 12, 8000, [], "制作/传播/贩卖淫秽色情内容"),
    ("四·公德", "歧视罪", "轻罪", 36, 30000, [], "因受保护特征贬低/辱骂/歧视他人"),
    ("四·公德", "诽谤罪", "轻罪", 24, 15000, [], "捏造/传播虚假信息损害名誉"),

    # ---------- 第五章 司法公正的犯罪 ----------
    ("五·司法", "行贿罪", "重罪", 24, 20000, [], "向政府雇员提供财物/利益以影响其职责"),
    ("五·司法", "妨碍司法罪", "重罪", 36, 15000, [], "阻碍法律文书签署/分发/执行"),
    ("五·司法", "拒绝听从执法人员命令与检查罪", "轻罪", 36, 15000, [], "拒绝执法人员合法命令/检查"),
    ("五·司法", "拒绝提供身份罪", "轻罪", 12, 10000, [], "被扣留/逮捕后拒绝提供身份信息"),
    ("五·司法", "破坏证据罪", "轻罪", 36, 20000, [], "隐瞒/销毁/篡改/伪造证据"),
    ("五·司法", "冒充政府雇员罪", "重罪", 24, 15000, [], "假冒政府雇员身份(穿制服/佩戴徽章等)"),
    ("五·司法", "虚假逮捕罪", "重罪", 48, 20000, [], "无法律依据以执法名义扣留/逮捕/没收(仅主管可起诉)"),
    ("五·司法", "伪证罪", "重罪", 48, 25000, [], "司法程序中提供虚假证言/信息"),
    ("五·司法", "冒充律师罪", "轻罪", 36, 15000, [], "未经认证从事律师业务/假扮律师"),
    ("五·司法", "妨碍政府雇员执行公务罪", "重罪", 36, 15000, [], "抵抗/延误/阻碍执法/消防/急救/司法人员执行职责"),
    ("五·司法", "拒捕罪", "重罪", 36, 15000, [], "逮捕时以任何方式拒绝/抗拒/逃避逮捕"),
    ("五·司法", "逃避羁押罪", "重罪", 36, 10000, [], "被合法扣留/逮捕后逃脱监护"),
    ("五·司法", "藐视法庭罪", "重罪", 48, 25000, [], "扰乱法庭秩序/无视法庭命令(须法官当庭宣布)"),
    ("五·司法", "恶意占用公共资源罪", "轻罪", 12, 5000, [], "非紧急/恶意占用政府紧急热线"),
    ("五·司法", "未经许可进入封闭紧急区域罪", "轻罪", 12, 15000, [], "未经授权进入被警戒线封闭的紧急区域"),
    ("五·司法", "袭警罪", "重罪", 48, 25000, [], "袭击/伤害警务人员/警车/警犬"),

    # ---------- 第六章 涉及公共秩序的犯罪 ----------
    ("六·秩序", "扰乱治安罪", "轻罪", 12, 5000, [], "公共场所挑衅/扰乱/冒犯性言语"),
    ("六·秩序", "非法集会罪", "轻罪", 12, 20000, [], "两人以上聚集引发骚乱/非法行动"),
    ("六·秩序", "煽动暴乱罪", "重罪", 48, 0, [], "意图引起暴乱/鼓励他人暴力破坏(仅监禁)"),
    ("六·秩序", "遮挡面部罪", "轻罪", 24, 5000, [], "遮面拒揭 / 犯罪过程中遮面(节庆等除外)"),

    # ---------- 第七章 热武器和设备的管制 ----------
    ("七·热武器", "非法持有热武器罪", "重罪", 24, 15000, [], "无证/非法持有枪支/爆炸物(执法公务除外)"),
    ("七·热武器", "非法展示热武器罪", "轻罪", 12, 8000, [], "公共场所公开携带/挥舞展示武器"),
    ("七·热武器", "非法使用热武器罪", "重罪", 24, 15000, [], "鲁莽/蓄意/威胁公共安全方式开枪"),
    ("七·热武器", "非法贩卖武器罪", "重罪", 48, 25000, [], "贩卖/转让/出售非法或未认证武器"),
    ("七·热武器", "走私罪", "重罪", 48, 30000, [], "非法运输/转让武器爆炸物 / 参与军火走私"),

    # ---------- 第十章 公共安全犯罪 ----------
    ("十·公共安全", "非法持有管制物品罪", "重罪", 48, 30000, [], "默认毒品/非法枪支(重罪4年$30000); 工具/8cm刀具=轻罪1年$15000"),
    ("十·公共安全", "非法持有大量管制物品罪", "重罪", 36, 6000, [], "按量: 10-50个=3年$6000; 51-200=2-4年$10000; 201-500=4-6年$15000; 501-999=6-10年$30000(可超5年)"),
    ("十·公共安全", "非法经营和分发管制物品罪", "重罪", 12, 5000, [], "以出售/分发/储存为目的经营管制物品"),
    ("十·公共安全", "非法制造管制物品罪", "重罪", 24, 8000, [], "生产/制造处方药/大麻等管制物品"),
    ("十·公共安全", "非法贩卖管制物品罪", "重罪", 36, 10000, [], "非法出售/赠送/运输管制物品给他人"),

    # ---------- 第十一章 其他违法行为 ----------
    ("十一·其他", "有组织犯罪", "重罪", 24, 10000, [], "团体以牟利为目的实施违法犯罪(逮捕须先取逮捕令)"),
    ("十一·其他", "洗钱罪", "重罪", 48, 30000, [], "持有/隐藏/转移违法所得 / 协助洗钱"),
    ("十一·其他", "身份盗用罪", "轻罪", 12, 8000, [], "未经许可盗取/使用他人个人信息, 另需赔偿"),
]

# 章节分隔标题
CH_TITLES = {
    "二·人身": "第二章 · 涉及人身安全的犯罪",
    "三·财产": "第三章 · 涉及财产安全的犯罪",
    "四·公德": "第四章 · 违反公共道德的犯罪",
    "五·司法": "第五章 · 司法公正的犯罪",
    "六·秩序": "第六章 · 涉及公共秩序的犯罪",
    "七·热武器": "第七章 · 热武器和设备的管制",
    "十·公共安全": "第十章 · 公共安全犯罪",
    "十一·其他": "第十一章 · 其他违法行为",
}

# 打开即见效果的示例: 当事人甲勾选这几条
EXAMPLE_JIA = {"抢劫罪", "拒捕罪", "逃避羁押罪", "非法使用热武器罪"}

MAX_TOTAL_MONTHS = 60   # 数罪叠加, 总刑期封顶 5 年

# =====================================================================
#  样式  —  执法主题: 深海军蓝 / 金 / 警务红
# =====================================================================
C_NAVY  = "12243B"; C_BLUE = "1F3A5F"; C_STEEL = "2C4A7C"
C_TEAL  = "1E6B5C"; C_PLUM = "5E4B8B"; C_GOLD  = "C9A227"; C_GOLDL = "F3E6BE"
C_PAGE  = "F4F6FA"; C_INK  = "1B2A41"; C_MUTE  = "6B7280"
C_FELONY = "F6D4CE"; C_MISD = "FBE3CC"; C_INFRACT = "FBF1C7"; C_HIT = "CDE8D5"

thin = Side(style="thin", color="D5DCE6")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
gold_b = Side(style="medium", color=C_GOLD)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFTV  = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHTV = Alignment(horizontal="right", vertical="center", wrap_text=True)
WRAP_TOP = Alignment(wrap_text=True, vertical="top")


def fill(c):
    return PatternFill("solid", fgColor=c)


def paint(ws, rng, fillc=None, font=None, align=None, border=None):
    from openpyxl.utils import range_boundaries
    minc, minr, maxc, maxr = range_boundaries(rng)
    for r in range(minr, maxr + 1):
        for c in range(minc, maxc + 1):
            cell = ws.cell(row=r, column=c)
            if fillc:
                cell.fill = fillc
            if border:
                cell.border = border
    tl = ws.cell(row=minr, column=minc)
    if font:
        tl.font = font
    if align:
        tl.alignment = align
    return tl


# =====================================================================
def build():
    wb = Workbook()
    wb.properties.creator = "袁尘"
    wb.properties.lastModifiedBy = "袁尘"
    wb.properties.title = "圣安地列斯州 · 案件罪名分析器"
    wb.properties.subject = "FiveM RP 案件罪名分析"
    wb.properties.description = "原创制作：袁尘"
    wb.properties.keywords = "袁尘 原创"

    ws_help = wb.active
    ws_help.title = "使用说明"
    ws_main = wb.create_sheet("算罪台")
    ws_ref  = wb.create_sheet("罪名表")

    build_main(ws_main)
    build_ref(ws_ref)
    build_help(ws_help)

    out = "圣安地列斯州_案件罪名分析器.xlsx"
    wb.save(out)
    print(f"已生成: {out}")
    print(f"罪名条数: {len(CHARGES)}")


# ---------------------------------------------------------------------
def build_main(m):
    m.sheet_view.showGridLines = False
    party_color = {"甲": C_STEEL, "乙": C_TEAL, "丙": C_PLUM}
    # 列: A章节 B罪名 C定性 D刑期 E罚款 F备注 G甲 H乙 I丙 (J/K/L隐藏累积)
    widths = {"A": 11, "B": 22, "C": 7, "D": 9, "E": 11, "F": 30,
              "G": 7, "H": 7, "I": 7, "J": 2, "K": 2, "L": 2}
    for col, w in widths.items():
        m.column_dimensions[col].width = w
    for col in ("J", "K", "L"):
        m.column_dimensions[col].hidden = True

    paint(m, "A1:I20", fillc=fill(C_PAGE))

    # ---- 标题 ----
    m.merge_cells("A1:I1")
    t = paint(m, "A1:I1", fillc=fill(C_NAVY),
              font=Font(color="FFFFFF", bold=True, size=18), align=CENTER)
    t.value = "⚖   圣 安 地 列 斯 州 · 案 件 罪 名 分 析 器   🦅"
    for c in range(1, 10):
        m.cell(row=1, column=c).border = Border(bottom=gold_b)
    m.row_dimensions[1].height = 38
    m.merge_cells("A2:I2")
    s = paint(m, "A2:I2", fillc=fill(C_NAVY),
              font=Font(color=C_GOLDL, size=10, italic=True), align=CENTER)
    s.value = "问答 + 逐条勾选　|　依据《圣安地列斯州刑法典》自动算罪"
    m.row_dimensions[2].height = 18

    # ---- ① 立案问答 ----
    m.merge_cells("A3:I3")
    sec1 = paint(m, "A3:I3", fillc=fill(C_GOLDL),
                 font=Font(bold=True, size=11, color=C_NAVY), align=LEFTV)
    sec1.value = "　① 立案问答　—　先回答下面三个问题"
    m.cell(row=3, column=1).border = Border(left=Side(style="thick", color=C_GOLD))
    m.row_dimensions[3].height = 22

    # 问答行: [需判罪人数] [受伤人数] [死亡人数]
    def qa(lab_range, val_cell, label, default, color):
        m.merge_cells(lab_range)
        paint(m, lab_range, fillc=fill(color),
              font=Font(color="FFFFFF", bold=True, size=10), align=CENTER).value = label
        vc = m[val_cell]
        vc.value = default
        vc.fill = fill("FFFDF5")
        vc.font = Font(bold=True, size=12, color=C_INK)
        vc.alignment = CENTER
        vc.border = BORDER
    qa("A4:B4", "C4", "❓ 需判罪人数", 1, C_STEEL)
    qa("D4:E4", "F4", "❓ 受伤人数", 0, C_TEAL)
    qa("G4:H4", "I4", "❓ 死亡人数", 0, C_PLUM)
    m.row_dimensions[4].height = 26

    # 动态提示
    m.merge_cells("A5:I5")
    hint = paint(m, "A5:I5", fillc=fill("FFF8E1"),
                 font=Font(size=10, color="8A6D00"), align=LEFTV)
    hint.value = ('="　本案需判罪 "&IF($C$4="","?",$C$4)&" 人。"'
                  '&IF(N($I$4)>0,"  ⚠ 有 "&$I$4&" 人死亡 → 请为相应嫌疑人勾选 一级/二级/三级谋杀罪 之一；","")'
                  '&IF(N($F$4)>0,"  有 "&$F$4&" 人受伤 → 可勾选 故意袭击罪 / 持有危险武器人身攻击罪；","")'
                  '&"  请在下方清单里逐条勾选每个人「做了什么」(在甲/乙/丙列填 1)，结果自动累加。"')
    m.row_dimensions[5].height = 30

    # ---- ② 结果 ----
    m.merge_cells("A6:I6")
    sec2 = paint(m, "A6:I6", fillc=fill(C_GOLDL),
                 font=Font(bold=True, size=11, color=C_NAVY), align=LEFTV)
    sec2.value = "　② 实时结果　—　随勾选自动更新（刑期已按法典封顶 5 年）"
    m.cell(row=6, column=1).border = Border(left=Side(style="thick", color=C_GOLD))
    m.row_dimensions[6].height = 22

    # 结果表头(row7): A当事人 B罪名数 C刑期 D罚款 E保释金 F:I已选罪名
    res_hdr = ["当事人", "罪名数", "合计刑期(月)", "合计罚款($)", "建议保释金($)"]
    for ci, htxt in enumerate(res_hdr, start=1):
        c = m.cell(row=7, column=ci, value=htxt)
        c.font = Font(color="FFFFFF", bold=True)
        c.fill = fill(C_BLUE)
        c.alignment = CENTER
        c.border = BORDER
    m.merge_cells("F7:I7")
    fh = paint(m, "F7:I7", fillc=fill(C_BLUE),
               font=Font(color="FFFFFF", bold=True), align=CENTER, border=BORDER)
    fh.value = "已选罪名"
    m.row_dimensions[7].height = 20

    # 结果数据(row8-10) 甲/乙/丙  → 列 G/H/I, 累积 J/K/L
    sus = [("甲", "G", "J", 8), ("乙", "H", "K", 9), ("丙", "I", "L", 10)]
    # 占位, last 在清单生成后回填
    for pname, col, cum, r in sus:
        pc = m.cell(row=r, column=1, value=f"当事人{pname}")
        pc.font = Font(color="FFFFFF", bold=True)
        pc.fill = fill(party_color[pname])
        pc.alignment = CENTER
        pc.border = BORDER
        for ci in (2, 3, 4, 5):
            cc = m.cell(row=r, column=ci)
            cc.border = BORDER
            cc.alignment = CENTER
            cc.font = Font(bold=True, size=12,
                           color="C0392B" if ci in (3, 5) else C_INK)
            cc.fill = fill("FFFFFF")
        m.cell(row=r, column=4).number_format = '#,##0'
        m.cell(row=r, column=5).number_format = '#,##0'
        m.merge_cells(f"F{r}:I{r}")
        lc = paint(m, f"F{r}:I{r}", fillc=fill("F7FAFC"),
                   font=Font(size=10, color=C_INK), align=LEFTV, border=BORDER)
        m.row_dimensions[r].height = 38

    # ---- 结论(row11) ----
    m.merge_cells("A11:I11")
    v = paint(m, "A11:I11", fillc=fill(C_NAVY),
              font=Font(bold=True, size=12, color="FFFFFF"), align=CENTER)
    for c in range(1, 10):
        m.cell(row=11, column=c).border = Border(top=gold_b, bottom=gold_b)
    m.row_dimensions[11].height = 34

    # ---- ③ 罪名勾选清单 ----
    m.merge_cells("A13:I13")
    sec3 = paint(m, "A13:I13", fillc=fill(C_GOLDL),
                 font=Font(bold=True, size=11, color=C_NAVY), align=LEFTV)
    sec3.value = ("　③ 罪名勾选清单　—　在「甲/乙/丙」列填 1 即给该人定该罪（多名受害人可填件数，"
                  "如打伤3人填3）；同一人同一罪只需填一次")
    m.cell(row=13, column=1).border = Border(left=Side(style="thick", color=C_GOLD))
    m.row_dimensions[13].height = 30
    m.row_dimensions[12].height = 6

    # 清单表头(row14)
    cl_hdr = ["章节", "罪名", "定性", "刑期(月)", "罚款($)", "说明 / 量刑要点", "甲", "乙", "丙"]
    for ci, htxt in enumerate(cl_hdr, start=1):
        c = m.cell(row=14, column=ci, value=htxt)
        c.font = Font(color="FFFFFF", bold=True)
        c.fill = fill(C_BLUE)
        c.alignment = CENTER
        c.border = BORDER
    m.row_dimensions[14].height = 22

    # 清单数据(row15+), 按章节插分隔行
    r = 15
    cur_chap = None
    for (chap, name, kind, months, fine, _kw, note) in CHARGES:
        if chap != cur_chap:
            cur_chap = chap
            m.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
            dv = paint(m, f"A{r}:I{r}", fillc=fill(C_NAVY),
                       font=Font(color="FFFFFF", bold=True, size=10), align=LEFTV)
            dv.value = "▍ " + CH_TITLES.get(chap, chap)
            # 分隔行也要带累积公式(空G→沿用上一行)
            for cum_col, scol in (("J", "G"), ("K", "H"), ("L", "I")):
                m[f"{cum_col}{r}"] = (f'=IF(${scol}{r}>0,IF(${cum_col}{r-1}="",$B{r},'
                                      f'${cum_col}{r-1}&"、"&$B{r}),${cum_col}{r-1})')
            m.row_dimensions[r].height = 18
            r += 1
        # 罪名行
        m.cell(row=r, column=1, value=chap).alignment = CENTER
        m.cell(row=r, column=2, value=name).font = Font(bold=True, color=C_INK)
        kc = m.cell(row=r, column=3, value=kind)
        kc.alignment = CENTER
        kc.fill = fill({"重罪": C_FELONY, "轻罪": C_MISD, "违法": C_INFRACT}.get(kind, "FFFFFF"))
        m.cell(row=r, column=4, value=months).alignment = CENTER
        fcell = m.cell(row=r, column=5, value=fine)
        fcell.alignment = CENTER
        fcell.number_format = '#,##0'
        m.cell(row=r, column=6, value=note).alignment = WRAP_TOP
        # 甲乙丙输入
        for ci, sname in ((7, "甲"), (8, "乙"), (9, "丙")):
            inp = m.cell(row=r, column=ci)
            inp.alignment = CENTER
            inp.fill = fill("FFFDF5")
            inp.font = Font(bold=True, color=C_STEEL)
            if sname == "甲" and name in EXAMPLE_JIA:
                inp.value = 1
        # 累积公式 J/K/L
        for cum_col, scol in (("J", "G"), ("K", "H"), ("L", "I")):
            m[f"{cum_col}{r}"] = (f'=IF(${scol}{r}>0,IF(${cum_col}{r-1}="",$B{r},'
                                  f'${cum_col}{r-1}&"、"&$B{r}),${cum_col}{r-1})')
        for ci in range(1, 10):
            m.cell(row=r, column=ci).border = BORDER
        m.row_dimensions[r].height = 30
        r += 1
    last = r - 1

    # 回填结果区公式
    for pname, col, cum, rr in sus:
        m.cell(row=rr, column=2,
               value=f'=COUNTIF(${col}$15:${col}${last},">0")')
        m.cell(row=rr, column=3,
               value=f'=MIN({MAX_TOTAL_MONTHS},SUMPRODUCT(${col}$15:${col}${last},$D$15:$D${last}))')
        m.cell(row=rr, column=4,
               value=f'=SUMPRODUCT(${col}$15:${col}${last},$E$15:$E${last})')
        m.cell(row=rr, column=5, value=f'=ROUND(C{rr}*D{rr}/6,0)')
        m.cell(row=rr, column=6,
               value=f'=IF(${cum}${last}="","（未勾选 / 无罪名）",${cum}${last})')

    # 结论公式
    m["A11"] = ('=IF(MAX(C8,C9,C10)=0,'
                '"⚠ 还没勾选任何罪名：请在上方问答后，到下面清单里给每个人勾选「做了什么」。",'
                '"🔨 主要责任方：当事人"&IF(C8=MAX(C8,C9,C10),"甲",IF(C9=MAX(C8,C9,C10),"乙","丙"))'
                '&"　|　合计刑期 "&MAX(C8,C9,C10)&" 个月（"&ROUND(MAX(C8,C9,C10)/12,1)&" 年）"'
                '&"　|　仅供参考，正当防卫 / 堡垒原则 / 同类不并罚等请人工复核")')

    # 数据有效性: 甲乙丙列下拉 1-5
    dv_n = DataValidation(type="list", formula1='"1,2,3,4,5"', allow_blank=True)
    m.add_data_validation(dv_n)
    dv_n.add(f"G15:I{last}")
    dv_p = DataValidation(type="list", formula1='"1,2,3"', allow_blank=True)
    m.add_data_validation(dv_p)
    dv_p.add("C4")
    dv_int = DataValidation(type="whole", operator="greaterThanOrEqual",
                            formula1="0", allow_blank=True)
    m.add_data_validation(dv_int)
    dv_int.add("F4"); dv_int.add("I4")

    # 命中行高亮(任意嫌疑人勾选 → 罪名行变绿)
    m.conditional_formatting.add(
        f"A15:I{last}",
        FormulaRule(formula=[f"AND($A15<>\"\",OR($G15>0,$H15>0,$I15>0))"], fill=fill(C_HIT)))

    # 提示 + 署名
    m.merge_cells(f"A{last+2}:I{last+2}")
    tip = paint(m, f"A{last+2}:I{last+2}", fillc=fill(C_PAGE),
                font=Font(size=9, italic=True, color=C_MUTE), align=WRAP_TOP)
    tip.value = ("提示 · 只算一人 → 只填「甲」列；保释金 = 刑期(月)×罚款÷6；"
                 "想加/改罪名 → 直接在本清单插行或改值；红=重罪 橙=轻罪 黄=违法。")
    m.merge_cells(f"A{last+3}:I{last+3}")
    paint(m, f"A{last+3}:I{last+3}", fillc=fill(C_PAGE),
          font=Font(size=10, bold=True, color=C_NAVY), align=RIGHTV).value = "原创制作：袁尘　"

    m.freeze_panes = "A15"


# ---------------------------------------------------------------------
def build_ref(w):
    """罪名表 —— 只读参考(完整列出, 方便查阅)。"""
    w.sheet_view.showGridLines = False
    headers = ["章节", "罪名", "定性", "最高刑期(月)", "罚款($)", "说明 / 量刑要点"]
    widths = {"A": 12, "B": 24, "C": 8, "D": 13, "E": 11, "F": 56}
    for col, ww in widths.items():
        w.column_dimensions[col].width = ww
    w.merge_cells("A1:F1")
    paint(w, "A1:F1", fillc=fill(C_NAVY),
          font=Font(color="FFFFFF", bold=True, size=14), align=CENTER).value = \
        "《圣安地列斯州刑法典》罪名速查表（原创整理：袁尘）"
    w.row_dimensions[1].height = 26
    for ci, h in enumerate(headers, start=1):
        c = w.cell(row=2, column=ci, value=h)
        c.font = Font(color="FFFFFF", bold=True)
        c.fill = fill(C_BLUE)
        c.alignment = CENTER
        c.border = BORDER
    r = 3
    for (chap, name, kind, months, fine, _kw, note) in CHARGES:
        w.cell(row=r, column=1, value=chap).alignment = CENTER
        w.cell(row=r, column=2, value=name).font = Font(bold=True)
        kc = w.cell(row=r, column=3, value=kind)
        kc.alignment = CENTER
        kc.fill = fill({"重罪": C_FELONY, "轻罪": C_MISD, "违法": C_INFRACT}.get(kind, "FFFFFF"))
        w.cell(row=r, column=4, value=months).alignment = CENTER
        fc = w.cell(row=r, column=5, value=fine)
        fc.alignment = CENTER
        fc.number_format = '#,##0'
        w.cell(row=r, column=6, value=note).alignment = WRAP_TOP
        zebra = "FFFFFF" if r % 2 else "EEF4FA"
        for ci in range(1, 7):
            cell = w.cell(row=r, column=ci)
            cell.border = BORDER
            if ci != 3:
                cell.fill = fill(zebra)
        w.row_dimensions[r].height = 28
        r += 1
    w.freeze_panes = "A3"


# ---------------------------------------------------------------------
def build_help(h):
    h.sheet_view.showGridLines = False
    h.merge_cells("A1:B1")
    hc = paint(h, "A1:B1", fillc=fill(C_NAVY),
               font=Font(color="FFFFFF", bold=True, size=15), align=CENTER)
    hc.value = "案件罪名分析器 — 使用说明（原创制作：袁尘）"
    h.row_dimensions[1].height = 28
    lines = [
        ("① 三步走", ""),
        ("", "第1步 立案问答：在「算罪台」顶部填 需判罪人数 / 受伤人数 / 死亡人数。下面会给出对应提示。"),
        ("", "第2步 勾选清单：在罪名清单里，给每个人「做了什么」逐条打勾——在 甲/乙/丙 列填 1。"),
        ("", "第3步 看结果：顶部「实时结果」自动累加每人罪名、刑期、罚款、保释金，并指出主要责任方。"),
        ("② 怎么填", ""),
        ("", "只判一个人 → 只填「甲」列。两三个人 → 分别填 甲 / 乙 / 丙 三列。"),
        ("", "同一个人同一条罪只填一次（填 1）。若同一条罪有多名受害人(如打伤3人)，可在该格填件数 3。"),
        ("", "勾选后整条罪名会变绿，方便核对。"),
        ("③ 计算规则（依法典）", ""),
        ("", "时间换算：1 年 = 12 个月 = 12 分钟（游戏内监禁时间）。"),
        ("", "数罪叠加、罚款累计；但总刑期最高不超过 5 年(60 个月)，同一罪名不重复并罚。"),
        ("", "建议保释金 = 刑期(月) × 罚款 ÷ 6（法典第九章·保释）。"),
        ("", "抢劫/盗窃/破坏等另需按价值赔偿(财产≥$8000；人身/精神$5000~$80000)，本表未自动计算。"),
        ("④ 需人工复核", ""),
        ("", "正当防卫：针对生命威胁的合理防卫，符合条件可无罪释放。"),
        ("", "堡垒原则：他人非法入侵/暴力袭击私人领地，可使用致命武力(引诱进入者除外)。"),
        ("", "未遂 / 同谋 / 教唆 / 从犯：按所对应主罪名「同罪论处」。包庇：按主犯论处但≤其50%。"),
        ("", "分级罪名(收受赃物、盗窃车辆、持有大量管制物品等)：本表取常见档位，请按实际情节核对。"),
        ("⑤ 想自己改", ""),
        ("", "改罪名/刑期/罚款 → 直接在「算罪台」清单或「罪名表」里改值。"),
        ("", "加罪名 → 在清单对应章节下插入一行，填好 罪名/刑期/罚款 即可（结果公式会自动覆盖到末行附近）。"),
        ("⑥ 免责声明", ""),
        ("", "本工具仅为辅助参考，最终定罪与量刑以服务器法官 / 司法部裁量为准。"),
        ("", ""),
        ("原创制作", "袁尘"),
    ]
    r = 3
    for left, right in lines:
        if left:
            h.cell(row=r, column=1, value=left).font = Font(bold=True, color=C_BLUE, size=11)
        if right:
            h.cell(row=r, column=2, value=right).alignment = WRAP_TOP
        r += 1
    h.column_dimensions["A"].width = 22
    h.column_dimensions["B"].width = 96


if __name__ == "__main__":
    build()
