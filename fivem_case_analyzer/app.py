# -*- coding: utf-8 -*-
"""
织梦星STAR · 案件罪名分析器  —  Windows 桌面版 (Tkinter, 简约少女粉)
========================================================================
算罪台一人一页，可增删嫌疑人：名字在上，案件经过(左) / 刑事罪名(右)。
刑事罪名以图表条实时涨跌显示「相似度」：绿80-100 / 黄50-79 / 红<50。
另含 罪名速查（双击看完整说明）、规章速查。离线运行。原创制作：袁尘。
"""
import os
import sys
import math
import tkinter as tk
from tkinter import ttk

import build_workbook as D


def resource_path(rel):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


# 🎀 简约少女粉主题（白底）
NAVY = "#C2185B"; HOT = "#FF7EB3"; HOT2 = "#FF9ECE"; PRESS = "#E0518A"
PINK = "#FFD1DC"; BG = "#FFF0F5"; WHITE = "#FFFFFF"; INK = "#7A3B5D"
MUTE = "#C99BB3"; LINE = "#FAD4E2"; RED = "#D81B60"; GOLD = "#FFC83D"
PLUM = "#B07CD6"; SRV = "#E8800C"
# 三色相似度条
GREEN = "#3FBF87"; YELLOW = "#F2C14E"; REDC = "#FF6B81"
TRACK = "#F4E3EC"
CAP = D.MAX_TOTAL_MONTHS
NONCRIM = ("内务违规", "服务器违规")
CUTE = "幼圆"          # 可爱圆体（Windows 自带；缺失自动回退）


def _rgb(h): return tuple(int(h[i:i + 2], 16) for i in (1, 3, 5))
def _hex(t): return "#%02X%02X%02X" % t
def lerp(a, b, t):
    ca, cb = _rgb(a), _rgb(b)
    return _hex(tuple(int(ca[i] + (cb[i] - ca[i]) * t) for i in range(3)))


def _rr(x0, y0, x1, y1, r):
    return [x0 + r, y0, x1 - r, y0, x1, y0, x1, y0 + r, x1, y1 - r, x1, y1,
            x1 - r, y1, x0 + r, y1, x0, y1, x0, y1 - r, x0, y0 + r, x0, y0]


def confidence(matched):
    """相似度: 直接命中罪名核心词就高，命中越多、词越具体越高。范围 60~99。"""
    m = len(matched)
    lmax = max(len(k) for k in matched)
    return max(60, min(99, 82 + (lmax - 2) * 6 + (m - 1) * 6))


def tier(pct):
    return GREEN if pct >= 80 else (YELLOW if pct >= 50 else REDC)


# 警察对嫌犯依法用武力 → 不算平民袭击（伤到公民才追责，见规章速查“两层并罚”）
_ASSAULT_SKIP = ("故意袭击罪", "持有危险武器人身攻击罪", "威胁罪")


def _lawful_police_force(text):
    police = any(w in text for w in ("警察", "警员", "警官", "巡警", "执法", "条子", "cop"))
    suspect = any(w in text for w in ("嫌疑人", "嫌犯", "疑犯", "罪犯", "逃犯", "通缉犯", "劫匪", "歹徒"))
    civ = any(w in text for w in ("公民", "平民", "市民", "路人", "无辜", "群众", "店员", "行人", "百姓"))
    return police and suspect and not civ


def analyze(text):
    text = text or ""
    crim, vio = [], []
    lawful = _lawful_police_force(text)
    if text.strip():
        for chap, name, kind, months, fine, kws, note in D.CHARGES:
            matched = [k for k in kws if k in text]
            if not matched:
                continue
            if lawful and name in _ASSAULT_SKIP:
                continue
            if kind in NONCRIM:
                vio.append((name, note))
            else:
                crim.append((name, kind, months, fine, note, confidence(matched)))
    crim.sort(key=lambda c: -c[5])
    months = min(CAP, sum(c[2] for c in crim))
    years = months / 12
    years = int(years) if years == int(years) else round(years, 2)
    fine = sum(c[3] for c in crim)
    bail = round(years * fine / 6)
    return crim, vio, years, fine, bail, len(crim)


# ===================== 圆角按钮 =====================
class RoundButton(tk.Canvas):
    def __init__(self, master, text, command, w=130, h=36,
                 base=HOT, hover=HOT2, press=PRESS, fg="white", bg=BG):
        super().__init__(master, width=w, height=h + 6, bg=bg, highlightthickness=0)
        self.command, self.base, self.hover, self.press = command, base, hover, press
        self.w, self.h = w, h
        self.shadow = self.create_polygon(_rr(0, 5, w, h + 5, 13), smooth=True, fill=PINK)
        self.shape = self.create_polygon(_rr(0, 0, w, h, 13), smooth=True, fill=base)
        self.txt = self.create_text(w / 2, h / 2, text=text, fill=fg,
                                    font=("Microsoft YaHei UI", 11, "bold"))
        self.bind("<Enter>", lambda e: self._set(self.hover, 0))
        self.bind("<Leave>", lambda e: self._set(self.base, 0))
        self.bind("<ButtonPress-1>", lambda e: self._set(self.press, 3))
        self.bind("<ButtonRelease-1>", self._release)

    def _set(self, color, dy):
        self.coords(self.shape, *_rr(0, dy, self.w, dy + self.h, 13))
        self.itemconfig(self.shape, fill=color)
        self.coords(self.txt, self.w / 2, dy + self.h / 2)

    def _release(self, e):
        self._set(self.hover, 0)
        if 0 <= e.x <= self.w and 0 <= e.y <= self.h + 6 and self.command:
            self.command()


# ===================== 霓虹粉光边框（围着框转的线） =====================
class NeonBorder(tk.Canvas):
    def __init__(self, master, app, pad=9, inner_bg="#FFF8FB"):
        super().__init__(master, bg=WHITE, highlightthickness=0)
        self.app, self.pad = app, pad
        self.inner = tk.Frame(self, bg=inner_bg)
        self.win = self.create_window(pad, pad, anchor="nw", window=self.inner)
        self.peri = []
        self.off = 0
        self.bind("<Configure>", self._cfg)
        app.fx.append(self)

    def cleanup(self):
        if self in self.app.fx:
            self.app.fx.remove(self)

    def _cfg(self, e):
        w, h = e.width, e.height
        if w <= 1:
            return
        self.itemconfig(self.win, width=w - 2 * self.pad, height=h - 2 * self.pad)
        self.delete("base")
        self.create_polygon(_rr(4, 4, w - 4, h - 4, 14), smooth=True, fill="",
                            outline=LINE, width=2, tags="base")
        self.tag_lower("base")
        self.peri = self._peri(w, h)

    def _peri(self, w, h, step=7):
        x0, y0, x1, y1 = 5, 5, w - 5, h - 5
        pts = []
        x = x0
        while x < x1: pts.append((x, y0)); x += step
        y = y0
        while y < y1: pts.append((x1, y)); y += step
        x = x1
        while x > x0: pts.append((x, y1)); x -= step
        y = y1
        while y > y0: pts.append((x0, y)); y -= step
        return pts

    def tick(self, t):
        if not self.peri:
            return
        try:
            self.delete("comet")
        except tk.TclError:
            return
        n = len(self.peri)
        self.off = (self.off + 2) % n
        K = 16
        for i in range(K):
            x, y = self.peri[(self.off - i) % n]
            f = 1 - i / K
            r = 1 + 3.2 * f
            self.create_oval(x - r, y - r, x + r, y + r,
                             fill=lerp("#FFD1DC", "#FF2D95", f), outline="", tags="comet")


# ===================== 刑事罪名 · 实时涨跌图表 =====================
class ChargeChart(tk.Canvas):
    def __init__(self, master, app):
        super().__init__(master, bg=WHITE, highlightthickness=0)
        self.app = app
        self.order = []
        self.bars = {}
        self.summary = (0, 0, 0, 0, [], False)
        app.fx.append(self)

    def cleanup(self):
        if self in self.app.fx:
            self.app.fx.remove(self)

    def set_data(self, crim, vio, years, fine, bail, n, typed):
        names = []
        for name, kind, mo, fn, note, pct in crim:
            names.append(name)
            if name not in self.bars:
                self.bars[name] = {"cur": 0.0, "target": float(pct)}
                self.order.append(name)
            else:
                self.bars[name]["target"] = float(pct)
        for name, b in self.bars.items():
            if name not in names:
                b["target"] = 0.0
        self.summary = (years, fine, bail, n, vio, typed)

    def tick(self, t):
        dead = []
        for name, b in self.bars.items():
            b["cur"] += (b["target"] - b["cur"]) * 0.18
            if b["target"] == 0.0 and b["cur"] < 0.6:
                dead.append(name)
        for name in dead:
            del self.bars[name]
            if name in self.order:
                self.order.remove(name)
        self._draw()

    def _draw(self):
        try:
            self.delete("all")
        except tk.TclError:
            return
        w = self.winfo_width() or 460
        years, fine, bail, n, vio, typed = self.summary
        x = 12
        if not typed:
            self.create_text(x, 20, anchor="w",
                             text="在左边「案件经过」里写下事情，罪名会实时浮现…",
                             fill=MUTE, font=("Microsoft YaHei UI", 10))
            return
        self.create_text(x, 18, anchor="w", text="共", fill=INK,
                         font=("Microsoft YaHei UI", 10))
        self.create_text(x + 22, 18, anchor="w", text=f"{n}", fill=NAVY,
                         font=("Microsoft YaHei UI", 13, "bold"))
        self.create_text(x + 38, 18, anchor="w", text="条刑事罪名",
                         fill=INK, font=("Microsoft YaHei UI", 10))
        live = [nm for nm in self.order if nm in self.bars]
        bx0, bx1 = 12, max(180, w - 56)
        y = 40
        rowh = 40
        if not live:
            self.create_text(x, y + 4, anchor="w",
                             text="（未识别到刑事罪名，请写得更具体或用法典里的词）",
                             fill=MUTE, font=("Microsoft YaHei UI", 10))
        for name in live[:8]:
            b = self.bars[name]
            pct = b["cur"]
            col = tier(pct)
            self.create_text(bx0, y, anchor="w", text=name, fill=INK,
                             font=("Microsoft YaHei UI", 10, "bold"))
            self.create_text(bx1 + 8, y, anchor="w", text=f"{round(pct)}%", fill=col,
                             font=("Consolas", 11, "bold"))
            ty = y + 12
            self.create_polygon(_rr(bx0, ty, bx1, ty + 14, 7), smooth=True,
                                fill=TRACK, outline="")
            frac = max(0.0, min(1.0, pct / 100.0))
            xe = bx0 + (bx1 - bx0) * frac
            if xe - bx0 > 4:
                self.create_polygon(_rr(bx0, ty, xe, ty + 14, 7), smooth=True,
                                    fill=col, outline="")
                self.create_polygon(_rr(bx0 + 3, ty + 2, max(bx0 + 4, xe - 3), ty + 6, 3),
                                    smooth=True, fill=lerp(col, "#FFFFFF", 0.5), outline="")
            y += rowh
        # 合计
        y += 2
        self.create_line(x, y, w - 12, y, fill=LINE)
        y += 16
        self.create_text(x, y, anchor="w", text="合计刑期", fill=MUTE,
                         font=("Microsoft YaHei UI", 9))
        self.create_text(x + 60, y, anchor="w", text=f"{years} 年", fill=RED,
                         font=("Microsoft YaHei UI", 12, "bold"))
        self.create_text(x + 140, y, anchor="w", text="罚款", fill=MUTE,
                         font=("Microsoft YaHei UI", 9))
        self.create_text(x + 178, y, anchor="w", text=f"${fine:,}", fill=RED,
                         font=("Microsoft YaHei UI", 12, "bold"))
        y += 22
        self.create_text(x, y, anchor="w", text="建议保释金", fill=MUTE,
                         font=("Microsoft YaHei UI", 9))
        self.create_text(x + 70, y, anchor="w", text=f"${bail:,}", fill=RED,
                         font=("Microsoft YaHei UI", 12, "bold"))
        if vio:
            y += 26
            self.create_text(x, y, anchor="w", text="违规 / 处分：", fill=SRV,
                             font=("Microsoft YaHei UI", 10, "bold"))
            for name, note in vio[:5]:
                y += 19
                self.create_text(x + 12, y, anchor="w", text=f"• {name}　{note}",
                                 fill=SRV, font=("Microsoft YaHei UI", 9))


# ===================== 单个嫌疑人页 =====================
class SuspectPage(tk.Frame):
    def __init__(self, master, app, example=False):
        super().__init__(master, bg=WHITE)
        self.app = app
        # 名字（上方）
        top = tk.Frame(self, bg=WHITE); top.pack(fill="x", padx=12, pady=(12, 4))
        tk.Label(top, text="🌸 名字：", bg=WHITE, fg=NAVY,
                 font=(CUTE, 12, "bold")).pack(side="left")
        self.name = tk.Entry(top, relief="flat", bg="#FFF8FB", fg=INK,
                            font=("Microsoft YaHei UI", 11), highlightthickness=2,
                            highlightbackground=LINE, highlightcolor="#FF8FBF")
        self.name.pack(side="left", fill="x", expand=True, padx=(4, 0))
        self.name.insert(0, "（填写嫌疑人名字）")

        # 左右分栏
        body = tk.Frame(self, bg=WHITE); body.pack(fill="both", expand=True, padx=12, pady=(4, 12))
        body.columnconfigure(0, weight=1, uniform="col")
        body.columnconfigure(1, weight=1, uniform="col")
        body.rowconfigure(1, weight=1)
        tk.Label(body, text="📝 案件经过 ✍️", bg=PINK, fg=NAVY, anchor="w", padx=8, pady=3,
                 font=(CUTE, 11, "bold")).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        tk.Label(body, text="⚖️ 刑事罪名 🎀", bg=PINK, fg=NAVY, anchor="w",
                 padx=8, pady=3, font=(CUTE, 11, "bold")).grid(
                 row=0, column=1, sticky="ew", padx=(6, 0))
        cb = NeonBorder(body, app, inner_bg="#FFF8FB")
        cb.grid(row=1, column=0, sticky="nsew", padx=(0, 6), pady=(2, 0))
        self.case = tk.Text(cb.inner, wrap="word", bg="#FFF8FB", fg=INK, relief="flat",
                           font=("Microsoft YaHei UI", 11), highlightthickness=0,
                           padx=8, pady=6)
        self.case.pack(fill="both", expand=True)
        self.case.bind("<KeyRelease>", self._on_change)
        vb = NeonBorder(body, app, inner_bg=WHITE)
        vb.grid(row=1, column=1, sticky="nsew", padx=(6, 0), pady=(2, 0))
        self.chart = ChargeChart(vb.inner, app)
        self.chart.pack(fill="both", expand=True)
        self._borders = [cb, vb]

        if example:
            self.case.insert("1.0", "嫌疑人持枪抢劫便利店，被警察拦下后拒捕并开枪，然后驾车逃跑还撞坏了路边的车。")
        self.after(120, self.refresh)

    def _on_change(self, _e=None):
        if getattr(self, "_job", None):
            self.after_cancel(self._job)
        self._job = self.after(160, self.refresh)

    def refresh(self):
        txt = self.case.get("1.0", "end-1c")
        crim, vio, years, fine, bail, n = analyze(txt)
        self.chart.set_data(crim, vio, years, fine, bail, n, bool(txt.strip()))

    def cleanup(self):
        self.chart.cleanup()
        for b in self._borders:
            b.cleanup()


# ===================== 主程序 =====================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("✨ 织梦星STAR · 案件罪名分析器 ✨")
        self.geometry("1000x780")
        self.configure(bg=BG)
        try:
            self.iconbitmap(resource_path("icon.ico"))
        except Exception:
            pass
        self.option_add("*Font", ("Microsoft YaHei UI", 10))
        self.fx = []
        self._init_style()
        self._build_banner()

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=(8, 4))
        self._build_analyzer(nb)
        self._build_law(nb)
        self._build_rules(nb)

        # 底部小 banner
        foot = tk.Frame(self, bg=PINK); foot.pack(fill="x")
        tk.Label(foot, text="🎀✨ 织梦星 STAR　·　原创：袁尘 ✨🐻", bg=PINK, fg=NAVY,
                 font=(CUTE, 10, "bold")).pack(pady=3)

        self._alive = True
        self._t = 0
        self.protocol("WM_DELETE_WINDOW", self._close)
        self._animate()

    def _close(self):
        self._alive = False
        self.destroy()

    def _init_style(self):
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure("TNotebook", background=BG, borderwidth=0)
        st.configure("TNotebook.Tab", background=PINK, foreground=NAVY,
                     padding=(22, 8), font=("Microsoft YaHei UI", 10, "bold"))
        st.map("TNotebook.Tab", background=[("selected", NAVY)],
               foreground=[("selected", "white")])
        st.configure("Sus.TNotebook", background=WHITE, borderwidth=0)
        st.configure("Sus.TNotebook.Tab", background="#FFE3EE", foreground=NAVY,
                     padding=(16, 5), font=("Microsoft YaHei UI", 10, "bold"))
        st.map("Sus.TNotebook.Tab", background=[("selected", HOT)],
               foreground=[("selected", "white")])
        st.configure("Treeview", background="white", fieldbackground="white",
                     foreground=INK, rowheight=26, borderwidth=0)
        st.configure("Treeview.Heading", background=HOT, foreground="white",
                     font=("Microsoft YaHei UI", 10, "bold"), relief="flat")
        st.configure("TScrollbar", background=PINK, troughcolor=BG, borderwidth=0)

    # ---------- 横幅 ----------
    def _build_banner(self):
        c = tk.Canvas(self, height=92, bg=NAVY, highlightthickness=0)
        c.pack(fill="x")
        self.banner = c
        self.update_idletasks()
        W = self.winfo_width() or 1000
        for i in range(0, W, 6):
            col = lerp(NAVY, "#E2568E", i / max(1, W))
            c.create_rectangle(i, 0, i + 6, 92, fill=col, outline=col)
        try:
            self._logo = tk.PhotoImage(file=resource_path("logo.png"))
            c.create_image(48, 46, image=self._logo)
        except Exception:
            self._logo = None
        c.create_text(98, 48, anchor="w", text="织梦星 STAR · 案件罪名分析器",
                      fill="white", font=("Microsoft YaHei UI", 20, "bold"))
        self._stars = []
        for k, (sx, sy) in enumerate([(W - 56, 58), (W - 96, 30),
                                      (W - 140, 60), (W - 188, 32)]):
            sid = c.create_polygon(self._star_pts(sx, sy, 7), fill=GOLD, outline="")
            self._stars.append([sid, sx, sy, k * 1.7])
        # 顶上跳动的小熊
        self._bears = []
        for k, bx in enumerate((250, 320, 390, 460)):
            ids = self._bear(c, bx, 16, 9, HOT2 if k % 2 else "#FFB6D5")
            self._bears.append([ids, bx, 16, k * 1.3, 16.0])
        tk.Frame(self, bg=HOT, height=4).pack(fill="x")

    def _bear(self, c, cx, cy, s, col):
        ids = []
        ids.append(c.create_oval(cx - s, cy - s * 0.95, cx - s * 0.25, cy - s * 0.2, fill=col, outline=""))
        ids.append(c.create_oval(cx + s * 0.25, cy - s * 0.95, cx + s, cy - s * 0.2, fill=col, outline=""))
        ids.append(c.create_oval(cx - s * 0.85, cy - s * 0.6, cx + s * 0.85, cy + s * 0.85, fill=col, outline=""))
        ids.append(c.create_oval(cx - s * 0.38, cy + s * 0.12, cx + s * 0.38, cy + s * 0.62, fill="white", outline=""))
        ids.append(c.create_oval(cx - s * 0.45, cy - s * 0.18, cx - s * 0.22, cy + s * 0.05, fill="#5A2342", outline=""))
        ids.append(c.create_oval(cx + s * 0.22, cy - s * 0.18, cx + s * 0.45, cy + s * 0.05, fill="#5A2342", outline=""))
        ids.append(c.create_oval(cx - s * 0.1, cy + s * 0.22, cx + s * 0.1, cy + s * 0.42, fill="#5A2342", outline=""))
        return ids

    def _star_pts(self, cx, cy, ro):
        ri = ro * 0.45
        pts = []
        for i in range(10):
            a = math.radians(-90 + i * 36)
            r = ro if i % 2 == 0 else ri
            pts += [cx + r * math.cos(a), cy + r * math.sin(a)]
        return pts

    def _animate(self):
        if not self._alive:
            return
        try:
            self._t += 1
            W = self.banner.winfo_width() or 1000
            offs = [56, 96, 140, 188]
            for i, s in enumerate(self._stars):
                sid, _, sy, ph = s
                sx = W - offs[i]
                tw = 0.5 + 0.5 * math.sin(self._t * 0.07 + ph)
                self.banner.coords(sid, *self._star_pts(sx, sy, 5 + 3 * tw))
                self.banner.itemconfig(sid, fill=lerp("#FFE9A0", GOLD, tw))
            for b in self._bears:
                ids, bx, by, ph, cur = b
                ny = by + math.sin(self._t * 0.16 + ph) * 5
                dy = ny - cur
                for iid in ids:
                    self.banner.move(iid, 0, dy)
                b[4] = ny
            for w in list(self.fx):
                w.tick(self._t)
        except tk.TclError:
            return
        self.after(60, self._animate)

    # ---------- ① 算罪台（一人一页，可增删） ----------
    def _build_analyzer(self, nb):
        f = tk.Frame(nb, bg=WHITE); nb.add(f, text="  算罪台  ")
        bar = tk.Frame(f, bg=WHITE); bar.pack(fill="x", padx=8, pady=(8, 0))
        RoundButton(bar, "🎀 添加嫌疑人", self._add_suspect, w=140, bg=WHITE).pack(side="left")
        RoundButton(bar, "🗑 删除当前", self._del_suspect, w=126, base=PLUM,
                    hover="#C79AE0", press="#8E5BB0", bg=WHITE).pack(side="left", padx=8)
        tk.Label(bar, text="🐻 可以添加好多个嫌疑人哦～ 每人一页 💕", bg=WHITE, fg=HOT,
                 font=(CUTE, 10, "bold")).pack(side="left", padx=6)

        self.inner = ttk.Notebook(f, style="Sus.TNotebook")
        self.inner.pack(fill="both", expand=True, padx=8, pady=6)
        self.pages = []
        self._add_suspect(example=True)

    def _add_suspect(self, example=False):
        pg = SuspectPage(self.inner, self, example=example)
        self.pages.append(pg)
        self.inner.add(pg, text="  嫌疑人  ")
        self._renumber()
        self.inner.select(pg)

    def _del_suspect(self):
        if len(self.pages) <= 1:
            return
        cur = self.inner.select()
        for pg in self.pages:
            if str(pg) == cur:
                self.pages.remove(pg)
                self.inner.forget(pg)
                pg.cleanup()
                pg.destroy()
                break
        self._renumber()

    def _renumber(self):
        for i, pg in enumerate(self.pages, 1):
            self.inner.tab(pg, text=f"  嫌疑人 {i}  ")

    # ---------- ② 罪名速查（双击看完整说明） ----------
    def _build_law(self, nb):
        f = tk.Frame(nb, bg=BG); nb.add(f, text="  罪名速查  ")
        bar = tk.Frame(f, bg=BG); bar.pack(fill="x", padx=8, pady=8)
        tk.Label(bar, text="🔍 搜索：", bg=BG, fg=NAVY).pack(side="left")
        self.law_q = tk.Entry(bar, relief="flat", highlightthickness=2,
                              highlightbackground=LINE, highlightcolor="#FF8FBF")
        self.law_q.pack(side="left", fill="x", expand=True)
        self.law_q.bind("<KeyRelease>", lambda e: self._fill_law())
        tk.Label(bar, text="💡 双击一行看完整说明", bg=BG, fg=MUTE,
                 font=("Microsoft YaHei UI", 9)).pack(side="left", padx=6)

        cols = ("章节", "罪名", "定性", "刑期(年)", "罚款($)", "说明")
        self.tree = ttk.Treeview(f, columns=cols, show="headings")
        for c, w in zip(cols, (110, 200, 84, 70, 90, 380)):
            self.tree.heading(c, text=c); self.tree.column(c, width=w, anchor="w")
        self.tree.tag_configure("odd", background="#FFF5F9")
        self.tree.tag_configure("hover", background="#FFE3EE")
        vsb = ttk.Scrollbar(f, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=(0, 8))
        vsb.pack(side="right", fill="y", pady=(0, 8))
        self._hover_row = None
        self.tree.bind("<Motion>", self._tree_hover)
        self.tree.bind("<Leave>", lambda e: self._tree_unhover())
        self.tree.bind("<Double-1>", self._law_detail)
        self.tree.bind("<Return>", self._law_detail)
        self._fill_law()

    def _tree_hover(self, e):
        row = self.tree.identify_row(e.y)
        if row == self._hover_row:
            return
        self._tree_unhover()
        if row:
            self._hover_base = self.tree.item(row, "tags")
            self.tree.item(row, tags=("hover",))
            self._hover_row = row

    def _tree_unhover(self):
        if self._hover_row:
            try:
                self.tree.item(self._hover_row, tags=getattr(self, "_hover_base", ()))
            except tk.TclError:
                pass
            self._hover_row = None

    def _fill_law(self):
        q = self.law_q.get().strip()
        self.tree.delete(*self.tree.get_children())
        i = 0
        for chap, name, kind, months, fine, kws, note in D.CHARGES:
            if q and q not in (chap + name + kind + note + "".join(kws)):
                continue
            fee = f"{fine:,}" if fine else "—"
            yr = months / 12
            yr = (int(yr) if yr == int(yr) else round(yr, 2)) if months else "—"
            tag = "odd" if i % 2 else ""
            self.tree.insert("", "end", values=(chap, name, kind, yr, fee, note),
                             tags=(tag,) if tag else ())
            i += 1

    def _law_detail(self, e=None):
        item = self.tree.focus()
        if not item:
            return
        vals = self.tree.item(item, "values")
        if not vals:
            return
        chap, name, kind, yr, fee, note = vals
        win = tk.Toplevel(self)
        win.title(name)
        win.configure(bg=BG)
        win.geometry("560x380")
        win.transient(self)
        try:
            win.iconbitmap(resource_path("icon.ico"))
        except Exception:
            pass
        head = tk.Frame(win, bg=NAVY); head.pack(fill="x")
        tk.Label(head, text="⚖ " + name, bg=NAVY, fg="white",
                 font=("Microsoft YaHei UI", 14, "bold")).pack(anchor="w", padx=14, pady=10)
        chips = tk.Frame(win, bg=BG); chips.pack(fill="x", padx=14, pady=(10, 4))
        for k, v in (("章节", chap), ("定性", kind), ("刑期", f"{yr} 年"), ("罚款", f"${fee}")):
            cell = tk.Frame(chips, bg=PINK); cell.pack(side="left", padx=(0, 8))
            tk.Label(cell, text=f"{k}　", bg=PINK, fg=MUTE,
                     font=("Microsoft YaHei UI", 9)).pack(side="left", padx=(8, 0), pady=4)
            tk.Label(cell, text=v, bg=PINK, fg=NAVY,
                     font=("Microsoft YaHei UI", 10, "bold")).pack(side="left", padx=(0, 8), pady=4)
        tk.Label(win, text="📖 完整说明", bg=BG, fg=NAVY, anchor="w",
                 font=("Microsoft YaHei UI", 10, "bold")).pack(fill="x", padx=14, pady=(8, 2))
        body = tk.Text(win, wrap="word", bg="white", fg=INK, relief="flat",
                       font=("Microsoft YaHei UI", 11), padx=12, pady=10,
                       highlightthickness=2, highlightbackground=LINE)
        body.pack(fill="both", expand=True, padx=14, pady=(0, 10))
        body.insert("1.0", note or "（无）")
        body.config(state="disabled")
        RoundButton(win, "关闭", win.destroy, w=100, bg=BG).pack(pady=(0, 12))

    # ---------- ③ 规章速查 ----------
    def _build_rules(self, nb):
        f = tk.Frame(nb, bg=BG); nb.add(f, text="  规章速查  ")
        txt = tk.Text(f, wrap="word", bg="white", relief="flat", font=("Microsoft YaHei UI", 10),
                      padx=12, pady=10, highlightthickness=2, highlightbackground=LINE)
        vsb = ttk.Scrollbar(f, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=vsb.set)
        txt.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        vsb.pack(side="right", fill="y", pady=8)
        txt.tag_config("h", foreground="white", background=NAVY,
                       font=("Microsoft YaHei UI", 11, "bold"), spacing1=7, spacing3=4)
        txt.tag_config("k", foreground=HOT, font=("Microsoft YaHei UI", 10, "bold"))
        for typ, a, b in D.RULES:
            if typ == "H":
                txt.insert("end", " " + a + " \n", "h")
            else:
                txt.insert("end", "  " + a + "：", "k")
                txt.insert("end", b + "\n")
        txt.config(state="disabled")


if __name__ == "__main__":
    App().mainloop()
