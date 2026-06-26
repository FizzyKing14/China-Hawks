# -*- coding: utf-8 -*-
"""
圣安地列斯州 · 案件罪名分析器  —  Excel 生成脚本 (一句话自动算罪版)
=====================================================================
根据《圣安地列斯州刑法典》(FiveM RP) 生成离线、纯公式的 .xlsx。
用法: 把某人做了什么用「一句话」写进框里, 自动识别出他犯的所有罪名,
并合计刑期(封顶5年)、罚款、保释金, 多个嫌疑人可对比谁责任最大。

原理: 离线关键词识别 —— 句子里出现某罪的关键词(含口语/黑话)即判定该罪。
      想让它认识新说法, 在「关键词库」页加一行即可, 即时生效、无需联网。

用法:  python build_workbook.py
依赖:  openpyxl   (pip install openpyxl)
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule

# =====================================================================
#  罪名数据库  —  抄录自《圣安地列斯州刑法典》
#  (章节, 罪名, 定性, 最高刑期_月, 罚款_美元, [识别关键词...], 备注)
# =====================================================================
CHARGES = [
    # ---------- 第二章 涉及人身安全的犯罪 ----------
    ("二·人身", "威胁罪", "轻罪", 12, 8000,
     ["威胁", "恐吓", "口头威胁", "书面威胁", "扬言", "制造恐慌", "吓唬", "放狠话", "威胁要"],
     "口头/书面/网络威胁他人或其亲友"),
    ("二·人身", "持有危险武器人身攻击罪", "重罪", 48, 20000,
     ["持械攻击", "持刀攻击", "持械伤人", "用刀砍", "用刀捅", "用刀割", "拿刀", "持刀", "致命武器",
      "用武器打", "拿棍子打", "持棍", "用球棒", "抄家伙", "拿管子打",
      "割伤", "砍伤", "捅伤", "刺伤", "划伤", "刀伤"],
     "用武器/工具威胁或伤害他人"),
    ("二·人身", "一级谋杀罪", "重罪", 60, 30000,
     ["一级谋杀", "预谋杀人", "蓄意杀人", "有预谋", "谋划杀", "策划杀人"],
     "预谋故意杀人 / 重罪过程中杀人"),
    ("二·人身", "二级谋杀罪", "重罪", 60, 20000,
     ["二级谋杀", "故意杀人", "杀人", "打死", "枪杀", "致人死亡", "杀死", "弄死", "捅死", "砍死", "杀了"],
     "非预谋的故意杀人 (3至5年)"),
    ("二·人身", "三级谋杀罪", "重罪", 60, 15000,
     ["三级谋杀", "过失致死", "意外致死", "误杀", "失手打死", "冲突致死"],
     "无意致死/鲁莽过失致死 (3至5年, 不含正当防卫)"),
    ("二·人身", "非法拘禁罪", "重罪", 36, 15000,
     ["非法拘禁", "拘禁", "扣押他人", "限制人身自由", "强行关押", "私自关押", "关起来", "囚禁", "绑在"],
     "违背意愿约束/扣押/拘禁他人"),
    ("二·人身", "故意袭击罪", "重罪", 36, 15000,
     ["故意袭击", "袭击", "殴打", "打人", "动手打", "拳打脚踢", "暴力伤人", "揍", "打伤",
      "打了一顿", "围殴", "动手", "打架"],
     "蓄意暴力致他人身体受伤"),
    ("二·人身", "绑架罪", "重罪", 60, 25000,
     ["绑架", "劫持", "挟持", "绑票", "控制人质", "掳走", "抓走人质", "抓人质"],
     "侵占/羁押/控制他人意图恐吓"),
    ("二·人身", "亵渎尸体罪", "重罪", 24, 20000,
     ["亵渎尸体", "毁尸", "处置尸体", "销毁尸体", "藏尸", "焚尸", "碎尸"],
     "毁损/处置尸体掩盖犯罪"),
    ("二·人身", "诈骗罪", "重罪", 60, 5000,
     ["诈骗", "欺诈", "骗钱", "行骗", "诈骗钱财", "骗取", "坑钱", "设局骗", "骗了"],
     "以欺诈欺瞒获取他人财物 (1-5年)"),

    # ---------- 第三章 涉及财产安全的犯罪 ----------
    ("三·财产", "非法入侵罪", "轻罪", 24, 8000,
     ["非法入侵", "拒绝离开", "私闯", "擅自闯入", "赖着不走", "不肯走", "闯进别人", "闯入民宅"],
     "被要求离开后仍拒绝离开"),
    ("三·财产", "非法入侵禁区罪", "轻罪", 36, 10000,
     ["非法入侵禁区", "闯入禁区", "擅闯警局", "擅闯机场", "擅闯军事基地", "限制区域", "闯入监狱", "翻进警局"],
     "未经授权进入警局/机场/军事基地等限制区"),
    ("三·财产", "入室盗窃罪", "重罪", 36, 25000,
     ["入室盗窃", "入室", "撬门盗窃", "闯入盗窃", "破门盗窃", "翻窗偷", "撬锁进", "入室行窃"],
     "擅入住宅/办公室/仓库盗窃, 另需赔偿"),
    ("三·财产", "抢劫罪", "重罪", 60, 20000,
     ["抢劫", "持械抢劫", "持枪抢劫", "抢夺财物", "武力抢", "威胁抢", "打劫", "抢东西", "抢钱", "洗劫", "抢了"],
     "以武力/威胁强夺他人财产, 另需赔偿"),
    ("三·财产", "盗窃罪", "重罪", 48, 20000,
     ["盗窃", "偷窃", "偷东西", "顺走", "擅自使用他人财物", "偷走", "小偷", "扒窃", "顺手牵羊", "偷拿", "偷了"],
     "盗窃/占用他人财物, 另需赔偿"),
    ("三·财产", "盗窃车辆罪", "轻罪", 12, 5000,
     ["盗窃车辆", "偷车", "偷公务车", "开走他人车", "擅自开走", "顺走车", "撬车", "把车开走", "偷了车"],
     "默认未造成严重后果(轻罪1年$5000); 严重后果=重罪4年$15000, 均赔偿"),
    ("三·财产", "收受赃物罪", "轻罪", 24, 10000,
     ["收受赃物", "销赃", "买赃物", "藏赃", "收赃", "收黑货", "买黑车"],
     "默认赃值<$5000(轻罪2年$10000); ≥$5000=重罪4年$20000"),
    ("三·财产", "敲诈勒索罪", "重罪", 24, 20000,
     ["敲诈", "勒索", "敲诈勒索", "威胁索财", "强迫交钱", "收保护费", "逼着给钱", "逼钱"],
     "以威胁/恐吓/滥权强迫他人交付财物"),
    ("三·财产", "蓄意破坏罪", "轻罪", 24, 25000,
     ["蓄意破坏", "破坏财物", "毁损", "砸车", "打砸", "损毁车辆", "故意损坏", "砸东西", "砸玻璃",
      "破坏设施", "划车", "撞坏", "砸店"],
     "蓄意破坏他人财产/车辆, 另需赔偿"),
    ("三·财产", "挪用公款公物罪", "轻罪", 48, 80000,
     ["挪用公款", "挪用公物", "侵占公款", "职务侵占", "挪用公家"],
     "擅自挪用受托管理的财物, 另需赔偿"),
    ("三·财产", "持有失窃物品罪", "轻罪", 24, 2000,
     ["持有失窃物品", "持有失窃", "捡到不还", "非法占有遗失物", "藏着别人丢的"],
     "非法占有他人遗失/失窃财物 (1-2年), 另需赔偿"),

    # ---------- 第四章 违反公共道德的犯罪 ----------
    ("四·公德", "在公共场合从事淫秽行为罪", "轻罪", 24, 10000,
     ["公共场合淫秽", "公然猥亵", "淫秽行为", "招嫖", "公开招募性交易", "当街猥亵", "猥亵"],
     "公开场所性行为/猥亵/招嫖 (1-2年)"),
    ("四·公德", "组织卖淫罪", "重罪", 36, 15000,
     ["组织卖淫", "拉皮条", "强迫卖淫", "介绍卖淫", "开妓院"],
     "支持/介绍/强迫他人卖淫牟利"),
    ("四·公德", "卖淫罪", "轻罪", 24, 15000,
     ["卖淫", "性交易", "提供性服务", "接客"],
     "以性服务换取财物"),
    ("四·公德", "强奸罪", "重罪", 60, 30000,
     ["强奸", "性侵", "强迫性行为", "强迫发生关系", "迷奸"],
     "违背意愿强迫性行为 / 与未满18岁者发生关系"),
    ("四·公德", "传播淫秽色情罪", "轻罪", 12, 8000,
     ["传播淫秽", "传播色情", "淫秽内容", "色情视频", "贩卖色情", "发黄图", "传黄片"],
     "制作/传播/贩卖淫秽色情内容"),
    ("四·公德", "歧视罪", "轻罪", 36, 30000,
     ["歧视", "种族歧视", "辱骂", "侮辱", "贬低", "地图炮"],
     "因受保护特征贬低/辱骂/歧视他人"),
    ("四·公德", "诽谤罪", "轻罪", 24, 15000,
     ["诽谤", "造谣", "污蔑", "捏造", "损害名誉", "传播虚假信息", "抹黑", "造黄谣"],
     "捏造/传播虚假信息损害名誉"),

    # ---------- 第五章 司法公正的犯罪 ----------
    ("五·司法", "行贿罪", "重罪", 24, 20000,
     ["行贿", "贿赂", "送钱给", "塞钱", "贿赂政府", "给警察塞钱", "买通", "塞红包"],
     "向政府雇员提供财物/利益以影响其职责"),
    ("五·司法", "妨碍司法罪", "重罪", 36, 15000,
     ["妨碍司法", "阻碍司法", "阻碍法律文书"],
     "阻碍法律文书签署/分发/执行"),
    ("五·司法", "拒绝听从执法人员命令与检查罪", "轻罪", 36, 15000,
     ["拒绝命令", "不服从命令", "拒绝检查", "不配合检查", "拒绝指示", "不听指挥", "不配合", "无视警察命令", "不靠边"],
     "拒绝执法人员合法命令/检查"),
    ("五·司法", "拒绝提供身份罪", "轻罪", 12, 10000,
     ["拒绝提供身份", "拒报身份", "不提供身份", "拒绝出示证件", "不报姓名", "不给看证件", "拒绝报名字"],
     "被扣留/逮捕后拒绝提供身份信息"),
    ("五·司法", "破坏证据罪", "轻罪", 36, 20000,
     ["破坏证据", "销毁证据", "毁灭证据", "篡改证据", "伪造证据", "藏匿证据", "丢掉证据", "扔掉毒品"],
     "隐瞒/销毁/篡改/伪造证据"),
    ("五·司法", "冒充政府雇员罪", "重罪", 24, 15000,
     ["冒充政府", "冒充警察", "假冒警察", "冒充公职", "假扮警察", "假冒执法", "冒充公务员", "假警察"],
     "假冒政府雇员身份(穿制服/佩戴徽章等)"),
    ("五·司法", "虚假逮捕罪", "重罪", 48, 20000,
     ["虚假逮捕", "非法逮捕", "非法扣押", "无依据逮捕", "乱抓人"],
     "无法律依据以执法名义扣留/逮捕/没收(仅主管可起诉)"),
    ("五·司法", "伪证罪", "重罪", 48, 25000,
     ["伪证", "作伪证", "虚假证言", "假证供", "提供假证", "撒谎作证"],
     "司法程序中提供虚假证言/信息"),
    ("五·司法", "冒充律师罪", "轻罪", 36, 15000,
     ["冒充律师", "假冒律师", "假扮律师"],
     "未经认证从事律师业务/假扮律师"),
    ("五·司法", "妨碍政府雇员执行公务罪", "重罪", 36, 15000,
     ["妨碍公务", "阻碍执法", "妨碍执法", "阻挠执法", "妨碍政府雇员", "阻碍救援", "阻碍消防",
      "挡着警察", "干扰执法", "阻拦警察", "阻碍医护"],
     "抵抗/延误/阻碍执法/消防/急救/司法人员执行职责"),
    ("五·司法", "拒捕罪", "重罪", 36, 15000,
     ["拒捕", "拒绝逮捕", "抗拒逮捕", "拒绝被捕", "反抗逮捕", "不让铐", "挣脱", "反抗抓捕",
      "逃避逮捕", "逃避追捕", "逃避抓捕", "拒不就范"],
     "逮捕/追捕时以任何方式拒绝、抗拒或逃避逮捕（人尚未被捕）"),
    ("五·司法", "逃避羁押罪", "重罪", 36, 10000,
     ["逃避羁押", "越狱", "脱逃", "逃离监管", "逃脱羁押", "从拘留中逃", "从警车上逃",
      "逃出监管", "保释期间潜逃", "押解途中逃"],
     "★须先被合法扣留/逮捕后，再从监护中逃脱才算（仅在现场逃跑不算，那是拒捕罪）"),
    ("五·司法", "藐视法庭罪", "重罪", 48, 25000,
     ["藐视法庭", "扰乱法庭", "无视法庭命令", "不尊重法庭", "在法庭上闹"],
     "扰乱法庭秩序/无视法庭命令(须法官当庭宣布)"),
    ("五·司法", "恶意占用公共资源罪", "轻罪", 12, 5000,
     ["恶意占用公共资源", "骚扰报警电话", "恶意报警", "虚假报警", "乱打报警", "谎报警情", "谎报"],
     "非紧急/恶意占用政府紧急热线"),
    ("五·司法", "未经许可进入封闭紧急区域罪", "轻罪", 12, 15000,
     ["进入封闭区域", "闯入警戒线", "擅闯封锁区", "越过警戒线", "闯入事故现场", "跨过警戒带"],
     "未经授权进入被警戒线封闭的紧急区域"),
    ("五·司法", "袭警罪", "重罪", 48, 25000,
     ["袭警", "袭击警察", "打警察", "攻击警察", "伤害警察", "攻击警车", "攻击警犬", "殴打警察",
      "动手打警察", "撞警车", "枪击警察", "打了警察", "朝警察开枪", "朝警察射击", "向警察开枪",
      "向警察射击", "射击警察", "打伤警察", "击伤警察", "警察重伤", "警察受伤", "枪伤警察"],
     "袭击/伤害警务人员/警车/警犬"),

    # ---------- 第六章 涉及公共秩序的犯罪 ----------
    ("六·秩序", "扰乱治安罪", "轻罪", 12, 5000,
     ["扰乱治安", "挑衅", "寻衅", "扰民", "冒犯性语言", "引起恐慌", "公共场所吵闹", "闹事", "大吵大闹", "挑事", "起哄"],
     "公共场所挑衅/扰乱/冒犯性言语"),
    ("六·秩序", "非法集会罪", "轻罪", 12, 20000,
     ["非法集会", "聚众", "聚集骚乱", "非法聚集", "聚众闹事", "聚集斗殴", "聚众斗殴"],
     "两人以上聚集引发骚乱/非法行动"),
    ("六·秩序", "煽动暴乱罪", "重罪", 48, 0,
     ["煽动暴乱", "煽动暴力", "鼓动暴乱", "鼓动暴力", "带头闹事", "煽动"],
     "意图引起暴乱/鼓励他人暴力破坏(仅监禁)"),
    ("六·秩序", "遮挡面部罪", "轻罪", 24, 5000,
     ["遮挡面部", "蒙面", "戴面具", "兜帽遮脸", "拒绝揭开面罩", "蒙着脸", "戴头套"],
     "遮面拒揭 / 犯罪过程中遮面(节庆等除外)"),

    # ---------- 第七章 热武器和设备的管制 ----------
    ("七·热武器", "非法持有热武器罪", "重罪", 24, 15000,
     ["非法持有热武器", "无证持枪", "非法持枪", "持有非法武器", "无序列号枪支", "全自动枪支",
      "私藏枪支", "持有爆炸物", "身上有枪", "携带枪支", "藏枪"],
     "无证/非法持有枪支/爆炸物(执法公务除外)"),
    ("七·热武器", "非法展示热武器罪", "轻罪", 12, 8000,
     ["非法展示热武器", "公开持枪", "挥舞武器", "亮枪", "展示武器", "公开携带枪支", "掏枪威胁", "拔枪", "亮出枪", "举枪"],
     "公共场所公开携带/挥舞展示武器"),
    ("七·热武器", "非法使用热武器罪", "重罪", 24, 15000,
     ["非法使用热武器", "鲁莽开枪", "乱开枪", "违规射击", "开枪", "朝天开枪", "随意射击",
      "开了一枪", "射击", "扫射", "鸣枪", "对空鸣枪"],
     "鲁莽/蓄意/威胁公共安全方式开枪"),
    ("七·热武器", "非法贩卖武器罪", "重罪", 48, 25000,
     ["非法贩卖武器", "卖枪", "贩卖武器", "出售武器", "转让武器", "倒卖枪支", "军火买卖"],
     "贩卖/转让/出售非法或未认证武器"),
    ("七·热武器", "走私罪", "重罪", 48, 30000,
     ["走私", "军火走私", "走私武器", "走私爆炸物", "走私军火", "偷运军火"],
     "非法运输/转让武器爆炸物 / 参与军火走私"),

    # ---------- 第十章 公共安全犯罪 ----------
    ("十·公共安全", "非法持有管制物品罪", "重罪", 48, 30000,
     ["非法持有管制物品", "持有毒品", "持有大麻", "持有可卡因", "持有海洛因", "持有冰毒",
      "持有管制刀具", "管制器具", "指虎", "弹簧刀", "开锁器", "黑客电脑", "身上有毒品", "带着毒品", "藏毒"],
     "默认毒品/非法枪支(重罪4年$30000); 工具/8cm刀具=轻罪1年$15000"),
    ("十·公共安全", "非法持有大量管制物品罪", "重罪", 36, 6000,
     ["大量管制物品", "大量毒品", "分装毒品", "成批毒品", "整箱毒品", "一堆毒品", "大批毒品"],
     "按量: 10-50个=3年$6000; 51-200=2-4年$10000; 201-500=4-6年$15000; 501-999=6-10年$30000(可超5年)"),
    ("十·公共安全", "非法经营和分发管制物品罪", "重罪", 12, 5000,
     ["经营管制物品", "分发毒品", "贩毒据点", "毒品窝点", "开设毒品区域", "卖毒品的店"],
     "以出售/分发/储存为目的经营管制物品"),
    ("十·公共安全", "非法制造管制物品罪", "重罪", 24, 8000,
     ["制造管制物品", "制毒", "生产毒品", "种植大麻", "制造毒品", "种大麻", "煮毒"],
     "生产/制造处方药/大麻等管制物品"),
    ("十·公共安全", "非法贩卖管制物品罪", "重罪", 36, 10000,
     ["贩卖管制物品", "贩毒", "卖毒品", "运输毒品", "贩卖毒品", "运毒", "送毒品", "倒卖毒品"],
     "非法出售/赠送/运输管制物品给他人"),

    # ---------- 第十一章 其他违法行为 ----------
    ("十一·其他", "有组织犯罪", "重罪", 24, 10000,
     ["有组织犯罪", "团伙犯罪", "黑帮犯罪", "犯罪集团", "帮派作案", "团伙作案"],
     "团体以牟利为目的实施违法犯罪(逮捕须先取逮捕令)"),
    ("十一·其他", "洗钱罪", "重罪", 48, 30000,
     ["洗钱", "转移赃款", "清洗黑钱", "标记钞票", "漂白黑钱", "洗黑钱"],
     "持有/隐藏/转移违法所得 / 协助洗钱"),
    ("十一·其他", "身份盗用罪", "轻罪", 12, 8000,
     ["身份盗用", "盗用身份", "冒用身份", "盗用个人信息", "盗用银行账号", "冒用他人信息"],
     "未经许可盗取/使用他人个人信息, 另需赔偿"),
]

MAX_TOTAL_MONTHS = 60   # 数罪叠加, 总刑期封顶 5 年

# =====================================================================
#  样式  —  执法主题
# =====================================================================
C_NAVY="12243B"; C_BLUE="1F3A5F"; C_STEEL="2C4A7C"; C_TEAL="1E6B5C"; C_PLUM="5E4B8B"
C_GOLD="C9A227"; C_GOLDL="F3E6BE"; C_PAGE="F4F6FA"; C_INK="1B2A41"; C_MUTE="6B7280"
C_INPUT="FFFDF5"; C_FELONY="F6D4CE"; C_MISD="FBE3CC"; C_INFRACT="FBF1C7"; C_HIT="CDE8D5"

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
    ws_law  = wb.create_sheet("罪名表")
    ws_kw   = wb.create_sheet("关键词库")

    kw_rows = build_law_and_kw(ws_law, ws_kw)
    build_main(ws_main)
    build_help(ws_help)

    out = "圣安地列斯州_案件罪名分析器.xlsx"
    wb.save(out)
    print(f"已生成: {out}")
    print(f"罪名条数: {len(CHARGES)} ；关键词条数: {kw_rows}")


# ---------------------------------------------------------------------
def build_law_and_kw(ws_law, ws_kw):
    n = len(CHARGES)
    law_first, law_last = 2, n + 1
    kw_rows = sum(len(c[5]) for c in CHARGES)

    # ---- 罪名表(引擎+参考): 命中列 G/H/I, 累积 J/K/L ----
    headers = ["章节", "罪名", "定性", "刑期(月)", "罚款($)", "说明 / 量刑要点",
               "甲命中", "乙命中", "丙命中", "甲累积", "乙累积", "丙累积"]
    for ci, h in enumerate(headers, start=1):
        c = ws_law.cell(row=1, column=ci, value=h)
        c.font = Font(color="FFFFFF", bold=True)
        c.fill = fill(C_BLUE)
        c.alignment = CENTER
        c.border = BORDER
    for i, (chap, name, kind, months, fine, kws, note) in enumerate(CHARGES):
        r = law_first + i
        ws_law.cell(row=r, column=1, value=chap).alignment = CENTER
        ws_law.cell(row=r, column=2, value=name).font = Font(bold=True)
        kc = ws_law.cell(row=r, column=3, value=kind)
        kc.alignment = CENTER
        kc.fill = fill({"重罪": C_FELONY, "轻罪": C_MISD, "违法": C_INFRACT}.get(kind, "FFFFFF"))
        ws_law.cell(row=r, column=4, value=months).alignment = CENTER
        fc = ws_law.cell(row=r, column=5, value=fine)
        fc.alignment = CENTER
        fc.number_format = '#,##0'
        ws_law.cell(row=r, column=6, value=note).alignment = WRAP_TOP
        # 命中: 该罪在关键词库中, 对应嫌疑人句子里命中的关键词数 > 0
        ws_law.cell(row=r, column=7,
            value=f"=IF(SUMIFS('关键词库'!$C$2:$C${kw_rows+1},'关键词库'!$B$2:$B${kw_rows+1},$B{r})>0,1,0)")
        ws_law.cell(row=r, column=8,
            value=f"=IF(SUMIFS('关键词库'!$D$2:$D${kw_rows+1},'关键词库'!$B$2:$B${kw_rows+1},$B{r})>0,1,0)")
        ws_law.cell(row=r, column=9,
            value=f"=IF(SUMIFS('关键词库'!$E$2:$E${kw_rows+1},'关键词库'!$B$2:$B${kw_rows+1},$B{r})>0,1,0)")
        # 累积命中罪名
        if r == law_first:
            ws_law.cell(row=r, column=10, value=f'=IF($G{r}=1,$B{r},"")')
            ws_law.cell(row=r, column=11, value=f'=IF($H{r}=1,$B{r},"")')
            ws_law.cell(row=r, column=12, value=f'=IF($I{r}=1,$B{r},"")')
        else:
            ws_law.cell(row=r, column=10, value=f'=IF($G{r}=1,IF($J{r-1}="",$B{r},$J{r-1}&"、"&$B{r}),$J{r-1})')
            ws_law.cell(row=r, column=11, value=f'=IF($H{r}=1,IF($K{r-1}="",$B{r},$K{r-1}&"、"&$B{r}),$K{r-1})')
            ws_law.cell(row=r, column=12, value=f'=IF($I{r}=1,IF($L{r-1}="",$B{r},$L{r-1}&"、"&$B{r}),$L{r-1})')
        for ci in range(1, 13):
            ws_law.cell(row=r, column=ci).border = BORDER
        ws_law.row_dimensions[r].height = 26
    for col, w in {"A": 12, "B": 24, "C": 8, "D": 10, "E": 11, "F": 50,
                   "G": 7, "H": 7, "I": 7, "J": 26, "K": 26, "L": 26}.items():
        ws_law.column_dimensions[col].width = w
    for col in ("G", "H", "I", "J", "K", "L"):
        ws_law.column_dimensions[col].hidden = True
    ws_law.freeze_panes = "A2"
    ws_law.sheet_view.showGridLines = False
    ws_law.conditional_formatting.add(
        f"A{law_first}:F{law_last}",
        FormulaRule(formula=[f"OR($G{law_first}=1,$H{law_first}=1,$I{law_first}=1)"], fill=fill(C_HIT)))

    # ---- 关键词库 ----
    for ci, h in enumerate(["关键词", "对应罪名", "甲", "乙", "丙"], start=1):
        c = ws_kw.cell(row=1, column=ci, value=h)
        c.font = Font(color="FFFFFF", bold=True)
        c.fill = fill(C_BLUE)
        c.alignment = CENTER
        c.border = BORDER
    rr = 2
    for (chap, name, kind, months, fine, kws, note) in CHARGES:
        for k in kws:
            ws_kw.cell(row=rr, column=1, value=k)
            ws_kw.cell(row=rr, column=2, value=name)
            for ci, inp in ((3, "B4"), (4, "B5"), (5, "B6")):
                ws_kw.cell(row=rr, column=ci,
                    value=f'=IF(AND($A{rr}<>"",\'算罪台\'!${inp[0]}${inp[1:]}<>"",'
                          f'ISNUMBER(SEARCH($A{rr},\'算罪台\'!${inp[0]}${inp[1:]}))),1,0)')
            for ci in range(1, 6):
                ws_kw.cell(row=rr, column=ci).border = BORDER
            rr += 1
    ws_kw.column_dimensions["A"].width = 20
    ws_kw.column_dimensions["B"].width = 24
    for col in ("C", "D", "E"):
        ws_kw.column_dimensions[col].width = 6
    ws_kw.freeze_panes = "A2"
    ws_kw.sheet_view.showGridLines = False
    # 顶部说明
    ws_kw.insert_rows(1)
    ws_kw.merge_cells("A1:E1")
    paint(ws_kw, "A1:E1", fillc=fill(C_GOLDL),
          font=Font(bold=True, size=10, color=C_NAVY), align=LEFTV).value = \
        "　想让它认识新说法/黑话：在 A 列加关键词、B 列写对应罪名（与「罪名表」一字不差），即时生效。"
    return kw_rows


# ---------------------------------------------------------------------
def build_main(m):
    m.sheet_view.showGridLines = False
    pc = {"甲": C_STEEL, "乙": C_TEAL, "丙": C_PLUM}
    for col, w in {"A": 13, "B": 40, "C": 8, "D": 12, "E": 13, "F": 15}.items():
        m.column_dimensions[col].width = w
    paint(m, "A1:F20", fillc=fill(C_PAGE))

    # 标题
    m.merge_cells("A1:F1")
    t = paint(m, "A1:F1", fillc=fill(C_NAVY),
              font=Font(color="FFFFFF", bold=True, size=18), align=CENTER)
    t.value = "⚖   圣 安 地 列 斯 州 · 案 件 罪 名 分 析 器   🦅"
    for c in range(1, 7):
        m.cell(row=1, column=c).border = Border(bottom=gold_b)
    m.row_dimensions[1].height = 38
    m.merge_cells("A2:F2")
    paint(m, "A2:F2", fillc=fill(C_NAVY),
          font=Font(color=C_GOLDL, size=10, italic=True), align=CENTER).value = \
        "把每个人做了什么，用一句话写进框里 → 自动算出他犯了哪些罪"
    m.row_dimensions[2].height = 18

    # ① 输入
    m.merge_cells("A3:F3")
    s1 = paint(m, "A3:F3", fillc=fill(C_GOLDL),
               font=Font(bold=True, size=11, color=C_NAVY), align=LEFTV)
    s1.value = "　① 一句话案情　—　每个人一句（只算一人就只填甲；可写一长段，命中越多越准）"
    m.cell(row=3, column=1).border = Border(left=Side(style="thick", color=C_GOLD))
    m.row_dimensions[3].height = 24

    examples = {
        4: ("甲", "嫌疑人持枪抢劫便利店，被警察拦下后拒捕并开枪，然后驾车逃跑还撞坏了路边的车。"),
        5: ("乙", ""),
        6: ("丙", ""),
    }
    for row, (pname, ex) in examples.items():
        m.merge_cells(start_row=row, start_column=1, end_row=row, end_column=1)
        lab = paint(m, f"A{row}:A{row}", fillc=fill(pc[pname]),
                    font=Font(color="FFFFFF", bold=True, size=11), align=CENTER)
        lab.value = f"👤\n当事人{pname}"
        m.merge_cells(start_row=row, start_column=2, end_row=row, end_column=6)
        inp = paint(m, f"B{row}:F{row}", fillc=fill(C_INPUT),
                    font=Font(size=10, color=C_INK), align=WRAP_TOP, border=BORDER)
        inp.value = ex
        m.row_dimensions[row].height = 50

    # ② 结果
    m.row_dimensions[7].height = 6
    m.merge_cells("A8:F8")
    s2 = paint(m, "A8:F8", fillc=fill(C_GOLDL),
               font=Font(bold=True, size=11, color=C_NAVY), align=LEFTV)
    s2.value = "　② 自动分析结果　—　识别到的罪名会即时列出（刑期已封顶 5 年）"
    m.cell(row=8, column=1).border = Border(left=Side(style="thick", color=C_GOLD))
    m.row_dimensions[8].height = 24

    res_hdr = ["当事人", "犯下的罪名", "罪名数", "合计刑期(月)", "合计罚款($)", "建议保释金($)"]
    for ci, h in enumerate(res_hdr, start=1):
        c = m.cell(row=9, column=ci, value=h)
        c.font = Font(color="FFFFFF", bold=True)
        c.fill = fill(C_BLUE)
        c.alignment = CENTER
        c.border = BORDER
    m.row_dimensions[9].height = 22

    law_last = len(CHARGES) + 1
    sus = [("甲", "G", "J", 10), ("乙", "H", "K", 11), ("丙", "I", "L", 12)]
    for pname, hit, cum, r in sus:
        nm = m.cell(row=r, column=1, value=f"当事人{pname}")
        nm.font = Font(color="FFFFFF", bold=True)
        nm.fill = fill(pc[pname])
        nm.alignment = CENTER
        nm.border = BORDER
        m.cell(row=r, column=2,
               value=f'=IF(\'罪名表\'!${cum}${law_last}="","（这句话没识别到罪名，请补充细节或去「关键词库」加词）",\'罪名表\'!${cum}${law_last})')
        m.cell(row=r, column=3, value=f"=SUM('罪名表'!${hit}$2:${hit}${law_last})")
        m.cell(row=r, column=4,
               value=f"=MIN({MAX_TOTAL_MONTHS},SUMIF('罪名表'!${hit}$2:${hit}${law_last},1,'罪名表'!$D$2:$D${law_last}))")
        m.cell(row=r, column=5,
               value=f"=SUMIF('罪名表'!${hit}$2:${hit}${law_last},1,'罪名表'!$E$2:$E${law_last})")
        m.cell(row=r, column=6, value=f"=ROUND(D{r}*E{r}/6,0)")
        for ci in range(2, 7):
            cc = m.cell(row=r, column=ci)
            cc.border = BORDER
            if ci == 2:
                cc.alignment = WRAP_TOP
                cc.font = Font(size=10, color=C_INK)
                cc.fill = fill("F7FAFC")
            else:
                cc.alignment = CENTER
                cc.fill = fill("FFFFFF")
                cc.font = Font(bold=True, size=12, color="C0392B" if ci in (4, 6) else C_INK)
        m.cell(row=r, column=5).number_format = '#,##0'
        m.cell(row=r, column=6).number_format = '#,##0'
        m.row_dimensions[r].height = 46

    # 结论
    m.merge_cells("A14:F14")
    v = paint(m, "A14:F14", fillc=fill(C_NAVY),
              font=Font(bold=True, size=12, color="FFFFFF"), align=CENTER)
    for c in range(1, 7):
        m.cell(row=14, column=c).border = Border(top=gold_b, bottom=gold_b)
    v.value = ('=IF(MAX(D10,D11,D12)=0,'
               '"⚠ 还没识别到罪名：在上面把案情写得更具体些，或去「关键词库」补充说法。",'
               '"🔨 主要责任方：当事人"&IF(D10=MAX(D10,D11,D12),"甲",IF(D11=MAX(D10,D11,D12),"乙","丙"))'
               '&"　|　合计刑期 "&MAX(D10,D11,D12)&" 个月（"&ROUND(MAX(D10,D11,D12)/12,1)&" 年）"'
               '&"　|　仅供参考，正当防卫 / 堡垒原则 / 同类不并罚等请人工复核")')
    m.row_dimensions[14].height = 36

    m.row_dimensions[15].height = 6
    m.merge_cells("A16:F16")
    paint(m, "A16:F16", fillc=fill(C_PAGE),
          font=Font(size=9, italic=True, color=C_MUTE), align=WRAP_TOP).value = \
        ("提示 · 写得越具体识别越准（如「持枪、拒捕、开枪、逃跑、贩毒」这些词）；"
         "保释金 = 刑期(月)×罚款÷6；识别不到的说法，去「关键词库」加一行即可。")
    m.row_dimensions[16].height = 32
    m.merge_cells("A18:F18")
    paint(m, "A18:F18", fillc=fill(C_PAGE),
          font=Font(size=10, bold=True, color=C_NAVY), align=RIGHTV).value = "原创制作：袁尘　"


# ---------------------------------------------------------------------
def build_help(h):
    h.sheet_view.showGridLines = False
    h.merge_cells("A1:B1")
    paint(h, "A1:B1", fillc=fill(C_NAVY),
          font=Font(color="FFFFFF", bold=True, size=15), align=CENTER).value = \
        "案件罪名分析器 — 使用说明（原创制作：袁尘）"
    h.row_dimensions[1].height = 28
    lines = [
        ("① 怎么用（全自动）", ""),
        ("", "打开「算罪台」，把某人做了什么用一句话（或一段话）写进黄色框里。"),
        ("", "下面「自动分析结果」会立刻列出他犯的所有罪名，并合计刑期、罚款、保释金。"),
        ("", "只算一个人 → 只填「当事人甲」；要对比多人 → 分别填甲/乙/丙，底部会指出谁责任最大。"),
        ("② 识别原理", ""),
        ("", "离线关键词识别：句子里出现某条罪的关键词（含常见口语/黑话），就判定该罪，不需联网、免费。"),
        ("", "写得越具体越准，比如「持枪抢劫」「拒捕」「开枪」「驾车逃跑」「贩毒」「袭警」这些词都能认。"),
        ("", "如果某种说法没被识别 → 打开「关键词库」页，A 列加上你的说法、B 列写对应罪名即可，立即生效。"),
        ("③ 计算规则（依法典）", ""),
        ("", "时间换算：1 年 = 12 个月 = 12 分钟（游戏内监禁时间）。"),
        ("", "数罪叠加、罚款累计；总刑期最高 5 年(60 个月)，同一罪名不重复并罚。"),
        ("", "建议保释金 = 刑期(月) × 罚款 ÷ 6（法典第九章·保释）。"),
        ("", "抢劫/盗窃/破坏等另需按价值赔偿，本表未自动计算。"),
        ("④ 需人工复核", ""),
        ("", "正当防卫 / 堡垒原则可减免；未遂/同谋/教唆/从犯按主罪同罪；包庇按主犯≤50%。"),
        ("", "分级罪名(收受赃物、盗窃车辆、大量管制物品等)取常见档位，请按实际情节核对。"),
        ("⑤ 三张表", ""),
        ("", "算罪台 = 你用的主界面；罪名表 = 全部罪名速查；关键词库 = 教它认新词的地方。"),
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
