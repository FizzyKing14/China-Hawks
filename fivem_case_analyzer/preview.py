# -*- coding: utf-8 -*-
"""渲染问答/勾选版「算罪台」预览图(仅供展示效果)。"""
from PIL import Image, ImageDraw, ImageFont
FZH = "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"
def F(sz): return ImageFont.truetype(FZH, sz)

NAVY="#12243B"; BLUE="#1F3A5F"; STEEL="#2C4A7C"; TEAL="#1E6B5C"; PLUM="#5E4B8B"
GOLD="#C9A227"; GOLDL="#F3E6BE"; PAGE="#F4F6FA"; INK="#1B2A41"; MUTE="#6B7280"
RED="#C0392B"; WHITE="#FFFFFF"; HIT="#CDE8D5"; FEL="#F6D4CE"; MISD="#FBE3CC"; HINT="#FFF8E1"

W=960; H=900
img=Image.new("RGB",(W,H),PAGE); d=ImageDraw.Draw(img)
M=20; w=W-2*M
def box(x,y,bw,bh,c,ol=None,ow=1): d.rectangle([x,y,x+bw,y+bh],fill=c,outline=ol,width=ow)
def tc(s,x,y,bw,bh,f,fl): d.text((x+bw/2,y+bh/2),s,font=f,fill=fl,anchor="mm")
def tl(s,x,y,bh,f,fl,pad=10): d.text((x+pad,y+bh/2),s,font=f,fill=fl,anchor="lm")

y=M
box(M,y,w,40,NAVY); d.line([M,y+40,M+w,y+40],fill=GOLD,width=3)
tc("⚖   圣安地列斯州 · 案件罪名分析器   🦅",M,y,w,40,F(22),WHITE); y+=40
box(M,y,w,22,NAVY); tc("问答 + 逐条勾选  |  依据《圣安地列斯州刑法典》自动算罪",M,y,w,22,F(11),GOLDL); y+=22+6

# ① 立案问答
box(M,y,w,24,GOLDL); box(M,y,5,24,GOLD); tl("①  立案问答  —  先回答下面三个问题",M+8,y,24,F(13),NAVY); y+=24+4
qw=(w-16)//3
for i,(lab,val,col) in enumerate([("❓ 需判罪人数","1",STEEL),("❓ 受伤人数","0",TEAL),("❓ 死亡人数","0",PLUM)]):
    x=M+i*(qw+8)
    box(x,y,qw*0.62,30,col); tc(lab,x,y,qw*0.62,30,F(12),WHITE)
    box(x+qw*0.62,y,qw*0.38,30,"#FFFDF5",ol="#D5DCE6"); tc(val,x+qw*0.62,y,qw*0.38,30,F(14),INK)
y+=30+4
box(M,y,w,28,HINT); tl("本案需判罪 1 人。 请在下方清单里逐条勾选每个人「做了什么」(在甲/乙/丙列填 1)，结果自动累加。",M,y,28,F(11),"#8A6D00"); y+=28+8

# ② 结果
box(M,y,w,24,GOLDL); box(M,y,5,24,GOLD); tl("②  实时结果  —  随勾选自动更新（刑期已封顶 5 年）",M+8,y,24,F(13),NAVY); y+=24+2
cols=[("当事人",90),("罪名数",70),("合计刑期(月)",96),("合计罚款($)",96),("保释金($)",92),("已选罪名",w-90-70-96-96-92)]
x=M
for t,cw in cols: box(x,y,cw,24,BLUE,ol="#D5DCE6"); tc(t,x,y,cw,24,F(11),WHITE); x+=cw
y+=24
rows=[("甲",STEEL,WHITE,"4","60","60,000","600,000","抢劫罪、拒捕罪、逃避羁押罪、非法使用热武器罪"),
      ("乙",TEAL,"#F7FAFC","0","0","0","0","（未勾选 / 无罪名）"),
      ("丙",PLUM,WHITE,"0","0","0","0","（未勾选 / 无罪名）")]
for name,col,zc,cnt,mon,fine,bail,crimes in rows:
    rh=38; x=M
    box(x,y,90,rh,col); tc(f"当事人{name}",x,y,90,rh,F(12),WHITE); x+=90
    for val,cw,red in [(cnt,70,False),(mon,96,True),(fine,96,False),(bail,92,True)]:
        box(x,y,cw,rh,WHITE,ol="#D5DCE6"); tc(val,x,y,cw,rh,F(13),RED if red else INK); x+=cw
    box(x,y,cols[5][1],rh,zc,ol="#D5DCE6"); tl(crimes,x,y,rh,F(11),INK,pad=8)
    y+=rh
y+=6
box(M,y,w,34,NAVY); d.line([M,y,M+w,y],fill=GOLD,width=2); d.line([M,y+34,M+w,y+34],fill=GOLD,width=2)
tc("🔨 主要责任方：当事人甲  |  合计刑期 60 个月（5 年）  |  仅供参考，正当防卫/堡垒原则等请人工复核",M,y,w,34,F(12),WHITE); y+=34+10

# ③ 清单
box(M,y,w,26,GOLDL); box(M,y,5,26,GOLD)
tl("③  罪名勾选清单  —  在「甲/乙/丙」列填 1 即定该罪（多名受害人可填件数）",M+8,y,26,F(12),NAVY); y+=26+2
clc=[("章节",78),("罪名",150),("定性",52),("刑期(月)",62),("罚款($)",70),("说明 / 量刑要点",w-78-150-52-62-70-3*44),("甲",44),("乙",44),("丙",44)]
x=M
for t,cw in clc: box(x,y,cw,22,BLUE,ol="#D5DCE6"); tc(t,x,y,cw,22,F(10),WHITE); x+=cw
y+=22
# 分隔行
box(M,y,w,18,NAVY); tl("▍ 第三章 · 涉及财产安全的犯罪",M,y,18,F(10),WHITE); y+=18
# 几条罪名(抢劫罪 甲已勾)
sample=[("三·财产","入室盗窃罪","重罪",FEL,"36","25,000","擅入住宅/仓库盗窃, 另需赔偿","","",""),
        ("三·财产","抢劫罪","重罪",FEL,"60","20,000","以武力/威胁强夺他人财产, 另需赔偿","1","",""),
        ("三·财产","盗窃罪","重罪",FEL,"48","20,000","盗窃/占用他人财物, 另需赔偿","","",""),
        ("三·财产","盗窃车辆罪","轻罪",MISD,"12","5,000","默认轻罪1年; 严重后果重罪4年","","","")]
for chap,name,kind,kc,mon,fine,note,a,bb,cc in sample:
    rh=26; x=M; ticked=(a=="1")
    rowbg=HIT if ticked else WHITE
    vals=[(chap,clc[0][1],"c",rowbg),(name,clc[1][1],"lb",rowbg),(kind,clc[2][1],"c",kc),
          (mon,clc[3][1],"c",rowbg),(fine,clc[4][1],"c",rowbg),(note,clc[5][1],"l",rowbg)]
    for val,cw,al,bg in vals:
        box(x,y,cw,rh,bg,ol="#D5DCE6")
        if al=="c": tc(val,x,y,cw,rh,F(10),INK)
        elif al=="lb": d.text((x+8,y+rh/2),val,font=F(11),fill=INK,anchor="lm")
        else: d.text((x+6,y+rh/2),val,font=F(9),fill=MUTE,anchor="lm")
        x+=cw
    for val,cw in [(a,44),(bb,44),(cc,44)]:
        box(x,y,cw,rh,"#FFFDF5" if not ticked else HIT,ol="#D5DCE6")
        if val: tc(val,x,y,cw,rh,F(12),STEEL)
        x+=cw
    y+=rh
y+=8
d.text((M+w,y),"原创制作：袁尘",font=F(12),fill=NAVY,anchor="rm")

img.save("preview_算罪台.png"); print("saved", img.size)
