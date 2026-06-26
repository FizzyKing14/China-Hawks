# -*- coding: utf-8 -*-
"""渲染「一句话自动算罪」版预览图。"""
from PIL import Image, ImageDraw, ImageFont
FZH = "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"
def F(sz): return ImageFont.truetype(FZH, sz)
NAVY="#12243B"; BLUE="#1F3A5F"; STEEL="#2C4A7C"; TEAL="#1E6B5C"; PLUM="#5E4B8B"
GOLD="#C9A227"; GOLDL="#F3E6BE"; PAGE="#F4F6FA"; INK="#1B2A41"; MUTE="#6B7280"
RED="#C0392B"; WHITE="#FFFFFF"

W=960; H=560
img=Image.new("RGB",(W,H),PAGE); d=ImageDraw.Draw(img)
M=22; w=W-2*M
def box(x,y,bw,bh,c,ol=None,ow=1): d.rectangle([x,y,x+bw,y+bh],fill=c,outline=ol,width=ow)
def tc(s,x,y,bw,bh,f,fl): d.text((x+bw/2,y+bh/2),s,font=f,fill=fl,anchor="mm")
def tlft(s,x,y,bh,f,fl,pad=10): d.text((x+pad,y+bh/2),s,font=f,fill=fl,anchor="lm")

y=M
box(M,y,w,40,NAVY); d.line([M,y+40,M+w,y+40],fill=GOLD,width=3)
tc("⚖   圣安地列斯州 · 案件罪名分析器   🦅",M,y,w,40,F(22),WHITE); y+=40
box(M,y,w,22,NAVY); tc("把每个人做了什么，用一句话写进框里 → 自动算出他犯了哪些罪",M,y,w,22,F(11),GOLDL); y+=22+8

box(M,y,w,24,GOLDL); box(M,y,5,24,GOLD); tlft("①  一句话案情  —  每个人一句（只算一人就只填甲）",M+8,y,24,F(13),NAVY); y+=24+4
cards=[("甲",STEEL,"嫌疑人持枪抢劫便利店，被警察拦下后拒捕并开枪，然后驾车逃跑还撞坏了路边的车。"),
       ("乙",TEAL,"（在此写当事人乙做了什么…）"),
       ("丙",PLUM,"（在此写当事人丙做了什么…）")]
chipw=92
for name,col,txt in cards:
    box(M,y,chipw,46,col)
    d.text((M+chipw/2,y+16),"👤",font=F(15),fill=WHITE,anchor="mm")
    d.text((M+chipw/2,y+30),f"当事人{name}",font=F(12),fill=WHITE,anchor="mm")
    box(M+chipw,y,w-chipw,46,"#FFFDF5",ol="#D5DCE6")
    ph=txt.startswith("（")
    d.text((M+chipw+12,y+23),txt,font=F(12),fill=MUTE if ph else INK,anchor="lm")
    y+=46+4
y+=8

box(M,y,w,24,GOLDL); box(M,y,5,24,GOLD); tlft("②  自动分析结果  —  识别到的罪名即时列出（刑期已封顶 5 年）",M+8,y,24,F(13),NAVY); y+=24+2
cols=[("当事人",90),("犯下的罪名",330),("罪名数",70),("合计刑期(月)",98),("合计罚款($)",98),("保释金($)",w-90-330-70-98-98)]
x=M
for t,cw in cols: box(x,y,cw,26,BLUE,ol="#D5DCE6"); tc(t,x,y,cw,26,F(11),WHITE); x+=cw
y+=26
rows=[("甲",STEEL,WHITE,"抢劫罪、蓄意破坏罪、拒捕罪、非法使用热武器罪","4","60","75,000","750,000"),
      ("乙",TEAL,"#F7FAFC","（这句话没识别到罪名…）","0","0","0","0"),
      ("丙",PLUM,WHITE,"（这句话没识别到罪名…）","0","0","0","0")]
for name,col,zc,crimes,cnt,mon,fine,bail in rows:
    rh=44; x=M
    box(x,y,90,rh,col); tc(f"当事人{name}",x,y,90,rh,F(12),WHITE); x+=90
    box(x,y,330,rh,zc,ol="#D5DCE6");
    # wrap crimes
    import textwrap
    d.text((x+8,y+rh/2),crimes,font=F(11),fill=INK if not crimes.startswith("（") else MUTE,anchor="lm"); x+=330
    for val,cw,red in [(cnt,70,False),(mon,98,True),(fine,98,False),(bail,cols[5][1],True)]:
        box(x,y,cw,rh,WHITE,ol="#D5DCE6"); tc(val,x,y,cw,rh,F(12),RED if red else INK); x+=cw
    y+=rh
y+=8
box(M,y,w,34,NAVY); d.line([M,y,M+w,y],fill=GOLD,width=2); d.line([M,y+34,M+w,y+34],fill=GOLD,width=2)
tc("🔨 主要责任方：当事人甲  |  合计刑期 60 个月（5 年）  |  仅供参考，正当防卫/堡垒原则等请人工复核",M,y,w,34,F(12),WHITE); y+=34+6
d.text((M+w,y+8),"原创制作：袁尘",font=F(12),fill=NAVY,anchor="rm")
img.save("preview_算罪台.png"); print("saved",img.size)
