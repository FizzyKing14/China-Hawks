# -*- coding: utf-8 -*-
"""按 build_workbook 的实际配色/布局渲染「算罪台」预览图(仅用于给用户看效果)。"""
from PIL import Image, ImageDraw, ImageFont

FZH = "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"
def F(sz): return ImageFont.truetype(FZH, sz)

# 与 build_workbook 一致的配色
NAVY="#12243B"; BLUE="#1F3A5F"; STEEL="#2C4A7C"; TEAL="#1E6B5C"; PLUM="#5E4B8B"
GOLD="#C9A227"; GOLDL="#F3E6BE"; PAGE="#F4F6FA"; INK="#1B2A41"; MUTE="#6B7280"
RED="#C0392B"; WHITE="#FFFFFF"; ZEB="#EEF4FA"

W=940; H=760
img=Image.new("RGB",(W,H),PAGE); d=ImageDraw.Draw(img)
M=24; w=W-2*M

def box(x,y,bw,bh,c,outline=None,ow=1):
    d.rectangle([x,y,x+bw,y+bh],fill=c,outline=outline,width=ow)
def ctext(s,x,y,bw,bh,font,fill,anchor="mm",pad=0,left=False):
    if left:
        d.text((x+pad,y+bh/2),s,font=font,fill=fill,anchor="lm")
    else:
        d.text((x+bw/2,y+bh/2),s,font=font,fill=fill,anchor="mm")

y=M
# 标题横幅
box(M,y,w,46,NAVY); d.line([M,y+46,M+w,y+46],fill=GOLD,width=3)
ctext("⚖   圣安地列斯州 · 案件罪名分析器   🦅",M,y,w,46,F(24),WHITE)
y+=46
# 副标题
box(M,y,w,26,NAVY)
ctext("STATE OF SAN ANDREAS · PENAL CODE CASE ANALYZER  |  依据《圣安地列斯州刑法典》自动算罪",
      M,y,w,26,F(12),GOLDL)
y+=26+10
# 区块① 标签
box(M,y,w,28,GOLDL); box(M,y,5,28,GOLD)
ctext("①  案情输入  —  把每位当事人「做了什么」分别填进右侧白框（只算一人就只填甲）",
      M,y,w,28,F(13),NAVY,left=True,pad=16)
y+=28+6
# 输入卡片
chipw=92
cards=[("甲",STEEL,"嫌疑人持枪抢劫便利店，被警察拦截后拒捕并开枪，随后驾车逃跑撞坏路边车辆。"),
       ("乙",TEAL,"（在此填写当事人乙的行为…）"),
       ("丙",PLUM,"（在此填写当事人丙的行为…）")]
for name,col,txt in cards:
    box(M,y,chipw,52,col)
    d.text((M+chipw/2,y+18),"👤",font=F(16),fill=WHITE,anchor="mm")
    d.text((M+chipw/2,y+34),f"当事人{name}",font=F(13),fill=WHITE,anchor="mm")
    box(M+chipw,y,w-chipw,52,"#FFFDF5",outline="#D5DCE6",ow=1)
    placeholder = txt.startswith("（")
    d.text((M+chipw+12,y+26),txt,font=F(13),
           fill=MUTE if placeholder else INK,anchor="lm")
    y+=52+4
y+=8
# 区块② 标签
box(M,y,w,28,GOLDL); box(M,y,5,28,GOLD)
ctext("②  分析结果  —  自动匹配罪名 / 合计刑期 / 罚款 / 保释金",
      M,y,w,28,F(13),NAVY,left=True,pad=16)
y+=28+4
# 结果表头
cols=[("当事人",92),("命中罪名",372),("罪名数",90),("合计刑期(月)",118),("合计罚款($)",118),("保释金($)",w-92-372-90-118-118)]
x=M
for title,cw in cols:
    box(x,y,cw,26,BLUE,outline="#D5DCE6")
    ctext(title,x,y,cw,26,F(12),WHITE); x+=cw
y+=26
rows=[("甲",STEEL,WHITE,"抢劫罪、拒捕罪、逃避羁押罪、非法使用热武器罪","4","60","60,000","600,000"),
      ("乙",TEAL,ZEB,"（无 / 请补充案情）","0","0","0","0"),
      ("丙",PLUM,WHITE,"（无 / 请补充案情）","0","0","0","0")]
for name,col,zc,crimes,cnt,mon,fine,bail in rows:
    rh=44; x=M
    box(x,y,92,rh,col); ctext(f"当事人{name}",x,y,92,rh,F(13),WHITE); x+=92
    box(x,y,372,rh,zc,outline="#D5DCE6")
    d.text((x+10,y+rh/2),crimes,font=F(12),fill=INK,anchor="lm"); x+=372
    for val,cw,red in [(cnt,90,False),(mon,118,True),(fine,118,False),(bail,cols[5][1],True)]:
        box(x,y,cw,rh,zc,outline="#D5DCE6")
        ctext(val,x,y,cw,rh,F(13),RED if red else INK); x+=cw
    y+=rh
y+=8
# 判定结论
box(M,y,w,40,NAVY); d.line([M,y,M+w,y],fill=GOLD,width=2); d.line([M,y+40,M+w,y+40],fill=GOLD,width=2)
ctext("🔨 主要责任方：当事人甲  |  合计刑期约 60 个月（5 年）  |  仅供参考，正当防卫/堡垒原则等请人工复核",
      M,y,w,40,F(13),WHITE)
y+=40+8
# 提示
ctext("提示 · ① 只算一人→只填甲  ② 刑期已封顶5年(60月)  ③ 保释金=刑期×罚款÷6  ④ 改罪名/关键词→去「罪名表」「关键词表」",
      M,y,w,22,F(11),MUTE,left=True,pad=2)
y+=26
# 署名
d.text((M+w,y),"原创制作：袁尘",font=F(13),fill=NAVY,anchor="rm")

img.save("preview_算罪台.png")
print("saved preview_算罪台.png", img.size)
