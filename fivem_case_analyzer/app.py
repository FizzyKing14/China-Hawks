# -*- coding: utf-8 -*-
"""
织梦星STAR · 案件罪名分析器  —  Windows 桌面版 (Tkinter, 粉色霓虹)
========================================================================
算罪台一人一页（嫌疑人 ①②③）：名字 / 案件经过 / 判决如下。
案件经过边打边算，刑事罪名以霓虹动态图表实时涨跌显示「可判置信度」。
另含 罪名速查、规章速查。离线运行。原创制作：袁尘。
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


# 🎀 粉色霓虹主题
NAVY = "#C2185B"; HOT = "#FF7EB3"; HOT2 = "#FF9ECE"; PRESS = "#E0518A"
PINK = "#FFD1DC"; BG = "#FFF0F5"; INK = "#7A3B5D"; MUTE = "#C99BB3"
RED = "#D81B60"; GOLD = "#FFC83D"; WHITE = "#FFFFFF"; LINE = "#FAD4E2"
# 霓虹（深底亮边）
DARK = "#1E0A16"; PANEL = "#2A1020"; PANEL2 = "#36152A"
NEON = "#FF2D95"; NEON2 = "#FF8FD0"; NEON_DK = "#6E1140"; NTXT = "#FFE3F3"
NMUTE = "#C98BB0"
T_HIGH = "#FF2D95"; T_MID = "#FF6FB5"; T_LOW = "#B97AA0"
CAP = D.MAX_TOTAL_MONTHS
NONCRIM = ("内务违规", "服务器违规")
SUSPECTS = ("①", "②", "③")


def _rgb(h): return tuple(int(h[i:i + 2], 16) for i in (1, 3, 5))
def _hex(t): return "#%02X%02X%02X" % t
def lerp(a, b, t):
    ca, cb = _rgb(a), _rgb(b)
    return _hex(tuple(int(ca[i] + (cb[i] - ca[i]) * t) for i in range(3)))


def _rr(x0, y0, x1, y1, r):
    """圆角矩形多边形坐标（配合 smooth=True）。"""
    return [x0 + r, y0, x1 - r, y0, x1, y0, x1, y0 + r, x1, y1 - r, x1, y1,
            x1 - r, y1, x0 + r, y1, x0, y1, x0, y1 - r, x0, y0 + r, x0, y0]


def confidence(matched):
    """可判置信度: 命中关键词越多、越具体 → 越高。范围 60~98%。"""
    m = len(matched)
    lmax = max(len(k) for k in matched)
    return max(60, min(98, 72 + (m - 1) * 9 + (lmax - 2) * 4))


def tier(pct):
    return T_HIGH if pct >= 85 else (T_MID if pct >= 70 else T_LOW)


def analyze(text):
    text = text or ""
    crim, vio = [], []
    if text.strip():
        for chap, name, kind, months, fine, kws, note in D.CHARGES:
            matched = [k for k in kws if k in text]
            if not matched:
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


# ===================== 霓虹光边容器 =====================
class NeonFrame(tk.Canvas):
    """内部放任意控件，四周环绕脉动的粉色霓虹光边。"""
    def __init__(self, master, app, height=90, pad=11, inner_bg=PANEL2):
        super().__init__(master, height=height, bg=PANEL, highlightthickness=0)
        self.app, self.pad = app, pad
        self.inner = tk.Frame(self, bg=inner_bg)
        self.win = self.create_window(pad, pad, anchor="nw", window=self.inner)
        self.rings = []
        self.bind("<Configure>", self._redraw)
        app.fx.append(self)

    def _redraw(self, e=None):
        w = self.winfo_width(); h = self.winfo_height()
        if w <= 1:
            return
        self.itemconfig(self.win, width=w - 2 * self.pad, height=h - 2 * self.pad)
        self.delete("glow")
        self.rings = []
        specs = [(2, 6, NEON_DK, NEON), (5, 3, NEON, NEON2), (8, 1, NEON2, "#FFFFFF")]
        for ins, wd, lo, hi in specs:
            rid = self.create_polygon(_rr(ins, ins, w - ins, h - ins, 16),
                                      smooth=True, fill="", outline=lo, width=wd,
                                      tags="glow")
            self.rings.append((rid, lo, hi))
        self.tag_lower("glow")

    def tick(self, t):
        if not self.rings:
            return
        b = 0.5 + 0.5 * math.sin(t * 0.09)
        for rid, lo, hi in self.rings:
            try:
                self.itemconfig(rid, outline=lerp(lo, hi, b))
            except tk.TclError:
                return


# ===================== 霓虹动态图表（罪名置信度涨跌） =====================
class ChargeChart(tk.Canvas):
    def __init__(self, master, app, height=300):
        super().__init__(master, height=height, bg=PANEL, highlightthickness=0)
        self.app = app
        self.order = []          # 罪名出现顺序
        self.bars = {}           # name -> {cur,target,tier}
        self.summary = (0, 0, 0, 0, [], False)
        app.fx.append(self)
        self.bind("<Configure>", lambda e: None)

    def set_data(self, crim, vio, years, fine, bail, n, typed):
        names = []
        for name, kind, mo, fn, note, pct in crim:
            names.append(name)
            if name not in self.bars:
                self.bars[name] = {"cur": 0.0, "target": float(pct), "tier": tier(pct)}
                self.order.append(name)
            else:
                self.bars[name]["target"] = float(pct)
                self.bars[name]["tier"] = tier(pct)
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
        self._draw(t)

    def _draw(self, t):
        try:
            self.delete("all")
        except tk.TclError:
            return
        w = self.winfo_width() or 760
        years, fine, bail, n, vio, typed = self.summary
        pulse = 0.5 + 0.5 * math.sin(t * 0.12)
        x = 14
        # 标题
        if not typed:
            self.create_text(x, 22, anchor="w", text="判决如下：在「案件经过」里写下事情，罪名会实时浮现…",
                             fill=NMUTE, font=("Microsoft YaHei UI", 11))
            return
        live = [nm for nm in self.order if nm in self.bars]
        self.create_text(x, 20, anchor="w", text="判决如下 · 刑事罪名", fill=NTXT,
                         font=("Microsoft YaHei UI", 12, "bold"))
        self.create_text(x + 150, 20, anchor="w", text=f"共 {n} 条", fill=NEON2,
                         font=("Microsoft YaHei UI", 11, "bold"))
        bx0, bx1 = 150, max(260, w - 80)
        y = 48
        rowh = 34
        if not live:
            self.create_text(x, y + 6, anchor="w",
                             text="（未识别到刑事罪名，请写得更具体或用法典里的词）",
                             fill=NMUTE, font=("Microsoft YaHei UI", 10))
        for name in live[:8]:
            b = self.bars[name]
            col = b["tier"]
            self.create_text(x, y + 11, anchor="w", text=name, fill=NTXT,
                             font=("Microsoft YaHei UI", 10, "bold"))
            # 轨道
            self.create_polygon(_rr(bx0, y, bx1, y + 22, 11), smooth=True,
                                fill="#3D1830", outline="")
            frac = max(0.0, min(1.0, b["cur"] / 100.0))
            xe = bx0 + (bx1 - bx0) * frac
            if xe - bx0 > 6:
                # 外发光
                glow = lerp(NEON_DK, col, pulse)
                self.create_polygon(_rr(bx0 - 1, y - 2, xe + 1, y + 24, 12),
                                    smooth=True, fill="", outline=glow, width=4)
                # 主条
                self.create_polygon(_rr(bx0, y, xe, y + 22, 11), smooth=True,
                                    fill=col, outline="")
                # 高光
                self.create_polygon(_rr(bx0 + 3, y + 3, xe - 3, y + 9, 5),
                                    smooth=True, fill=lerp(col, "#FFFFFF", 0.55),
                                    outline="")
                # 亮头
                self.create_oval(xe - 5, y + 4, xe + 3, y + 18,
                                 fill=lerp(col, "#FFFFFF", pulse), outline="")
            self.create_text(bx1 + 8, y + 11, anchor="w", text=f"{round(b['cur'])}%",
                             fill=col, font=("Consolas", 12, "bold"))
            y += rowh
        # 合计
        y += 6
        self.create_line(x, y, w - 14, y, fill="#5A2342")
        y += 16
        self.create_text(x, y, anchor="w", text="合计刑期", fill=NMUTE,
                         font=("Microsoft YaHei UI", 10))
        self.create_text(x + 70, y, anchor="w", text=f"{years} 年", fill=NEON,
                         font=("Microsoft YaHei UI", 13, "bold"))
        self.create_text(x + 160, y, anchor="w", text="罚款", fill=NMUTE,
                         font=("Microsoft YaHei UI", 10))
        self.create_text(x + 210, y, anchor="w", text=f"${fine:,}", fill=NEON,
                         font=("Microsoft YaHei UI", 13, "bold"))
        self.create_text(x + 360, y, anchor="w", text="保释金", fill=NMUTE,
                         font=("Microsoft YaHei UI", 10))
        self.create_text(x + 420, y, anchor="w", text=f"${bail:,}", fill=NEON,
                         font=("Microsoft YaHei UI", 13, "bold"))
        # 违规 / 处分
        if vio:
            y += 26
            self.create_text(x, y, anchor="w", text="违规 / 处分：", fill=NEON2,
                             font=("Microsoft YaHei UI", 10, "bold"))
            for name, note in vio[:4]:
                y += 20
                self.create_text(x + 16, y, anchor="w", text=f"• {name}　{note}",
                                 fill="#FFC07A", font=("Microsoft YaHei UI", 9))


# ===================== 圆角霓虹按钮 =====================
class RoundButton(tk.Canvas):
    def __init__(self, master, text, command, w=130, h=40,
                 base=NEON, hover=NEON2, press=PRESS, fg="white", bg=DARK):
        super().__init__(master, width=w, height=h + 8, bg=bg, highlightthickness=0)
        self.command, self.base, self.hover, self.press = command, base, hover, press
        self.w, self.h = w, h
        self.glow = self.create_polygon(self._pts(0, 4), smooth=True, fill="", outline=NEON_DK, width=4)
        self.shape = self.create_polygon(self._pts(0, 0), smooth=True, fill=base)
        self.txt = self.create_text(w / 2, h / 2, text=text, fill=fg,
                                    font=("Microsoft YaHei UI", 11, "bold"))
        self.bind("<Enter>", lambda e: self._set(self.hover, 0))
        self.bind("<Leave>", lambda e: self._set(self.base, 0))
        self.bind("<ButtonPress-1>", lambda e: self._set(self.press, 3))
        self.bind("<ButtonRelease-1>", self._release)

    def _pts(self, dy, exp):
        w, h, r = self.w, self.h, 14
        return _rr(0 - exp, dy - exp, w + exp, dy + h + exp, r)

    def _set(self, color, dy):
        self.coords(self.shape, *self._pts(dy, 0))
        self.itemconfig(self.shape, fill=color)
        self.coords(self.txt, self.w / 2, dy + self.h / 2)

    def _release(self, e):
        self._set(self.hover, 0)
        if 0 <= e.x <= self.w and 0 <= e.y <= self.h + 8 and self.command:
            self.command()


# ===================== 单个嫌疑人页 =====================
class SuspectPage(tk.Frame):
    def __init__(self, master, app, idx):
        super().__init__(master, bg=DARK)
        self.app = app
        pad = {"padx": 16, "pady": (8, 0)}
        tk.Label(self, text=f"🔮 嫌疑人 {idx}", bg=DARK, fg=NEON,
                 font=("Microsoft YaHei UI", 13, "bold")).pack(anchor="w", **pad)

        # 名字
        tk.Label(self, text="名字：", bg=DARK, fg=NTXT,
                 font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w", padx=16, pady=(8, 0))
        nf = NeonFrame(self, app, height=42)
        nf.pack(fill="x", padx=16, pady=(2, 0))
        self.name = tk.Entry(nf.inner, bg=PANEL2, fg=NTXT, insertbackground=NEON,
                             relief="flat", font=("Microsoft YaHei UI", 11),
                             highlightthickness=0)
        self.name.pack(fill="both", expand=True, padx=8, pady=4)
        self.name.insert(0, "（填写嫌疑人名字）")

        # 案件经过
        tk.Label(self, text="案件经过：", bg=DARK, fg=NTXT,
                 font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w", padx=16, pady=(10, 0))
        cf = NeonFrame(self, app, height=92)
        cf.pack(fill="x", padx=16, pady=(2, 0))
        self.case = tk.Text(cf.inner, height=3, wrap="word", bg=PANEL2, fg=NTXT,
                            insertbackground=NEON, relief="flat",
                            font=("Microsoft YaHei UI", 11), highlightthickness=0,
                            padx=6, pady=4)
        self.case.pack(fill="both", expand=True)
        self.case.bind("<KeyRelease>", self._on_change)

        # 判决如下
        tk.Label(self, text="判决如下：", bg=DARK, fg=NTXT,
                 font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w", padx=16, pady=(10, 0))
        vf = NeonFrame(self, app, height=312, pad=11)
        vf.pack(fill="both", expand=True, padx=16, pady=(2, 12))
        self.chart = ChargeChart(vf.inner, app, height=300)
        self.chart.pack(fill="both", expand=True)

        if idx == "①":
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


# ===================== 主程序 =====================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("✨ 织梦星STAR · 案件罪名分析器 ✨")
        self.geometry("980x780")
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
        nb.pack(fill="both", expand=True, padx=10, pady=8)
        self._build_analyzer(nb)
        self._build_law(nb)
        self._build_rules(nb)

        tk.Label(self, text="✨ 织梦星 STAR Roleplay　|　原创制作：袁尘 ✨", bg=BG, fg=NAVY,
                 font=("Microsoft YaHei UI", 9, "bold")).pack(fill="x", pady=4)

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
                     padding=(20, 8), font=("Microsoft YaHei UI", 10, "bold"))
        st.map("TNotebook.Tab", background=[("selected", NAVY)],
               foreground=[("selected", "white")])
        # 内层（嫌疑人）霓虹深色标签
        st.configure("Neon.TNotebook", background=DARK, borderwidth=0)
        st.configure("Neon.TNotebook.Tab", background=PANEL, foreground=NEON2,
                     padding=(22, 8), font=("Microsoft YaHei UI", 10, "bold"))
        st.map("Neon.TNotebook.Tab", background=[("selected", NEON)],
               foreground=[("selected", "white")])
        st.configure("Treeview", background="white", fieldbackground="white",
                     foreground=INK, rowheight=26, borderwidth=0)
        st.configure("Treeview.Heading", background=HOT, foreground="white",
                     font=("Microsoft YaHei UI", 10, "bold"), relief="flat")
        st.configure("TScrollbar", background=PINK, troughcolor=BG, borderwidth=0)

    # ---------- 横幅 ----------
    def _build_banner(self):
        c = tk.Canvas(self, height=96, bg=NAVY, highlightthickness=0)
        c.pack(fill="x")
        self.banner = c
        self.update_idletasks()
        W = self.winfo_width() or 980
        for i in range(0, W, 6):
            col = lerp(NAVY, "#E2568E", i / max(1, W))
            c.create_rectangle(i, 0, i + 6, 96, fill=col, outline=col)
        try:
            self._logo = tk.PhotoImage(file=resource_path("logo.png"))
            c.create_image(50, 48, image=self._logo)
        except Exception:
            self._logo = None
        c.create_text(100, 36, anchor="w", text="织梦星 STAR · 案件罪名分析器",
                      fill="white", font=("Microsoft YaHei UI", 19, "bold"))
        c.create_text(102, 67, anchor="w",
                      text="圣安地列斯州 · STAR Roleplay　💗　一人一页 · 一句话自动算罪",
                      fill="#FFE3EE", font=("Microsoft YaHei UI", 10))
        self._stars = []
        for k, (sx, sy) in enumerate([(W - 56, 30), (W - 96, 60),
                                      (W - 140, 26), (W - 188, 56)]):
            sid = c.create_polygon(self._star_pts(sx, sy, 7), fill=GOLD, outline="")
            self._stars.append([sid, sx, sy, k * 1.7])
        tk.Frame(self, bg=HOT, height=4).pack(fill="x")

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
            W = self.banner.winfo_width() or 980
            offs = [56, 96, 140, 188]
            for i, s in enumerate(self._stars):
                sid, _, sy, ph = s
                sx = W - offs[i]
                tw = 0.5 + 0.5 * math.sin(self._t * 0.07 + ph)
                self.banner.coords(sid, *self._star_pts(sx, sy, 5 + 3 * tw))
                self.banner.itemconfig(sid, fill=lerp("#FFE9A0", GOLD, tw))
            for w in self.fx:
                w.tick(self._t)
        except tk.TclError:
            return
        self.after(55, self._animate)

    # ---------- ① 算罪台（一人一页） ----------
    def _build_analyzer(self, nb):
        f = tk.Frame(nb, bg=DARK); nb.add(f, text="  算罪台  ")
        inner = ttk.Notebook(f, style="Neon.TNotebook")
        inner.pack(fill="both", expand=True, padx=6, pady=6)
        self.pages = []
        for s in SUSPECTS:
            pg = SuspectPage(inner, self, s)
            inner.add(pg, text=f"  嫌疑人 {s}  ")
            self.pages.append(pg)

    # ---------- ② 罪名速查 ----------
    def _build_law(self, nb):
        f = tk.Frame(nb, bg=BG); nb.add(f, text="  罪名速查  ")
        bar = tk.Frame(f, bg=BG); bar.pack(fill="x", padx=8, pady=8)
        tk.Label(bar, text="🔍 搜索：", bg=BG, fg=NAVY).pack(side="left")
        self.law_q = tk.Entry(bar, relief="flat", highlightthickness=2,
                              highlightbackground=LINE, highlightcolor="#FF8FBF")
        self.law_q.pack(side="left", fill="x", expand=True)
        self.law_q.bind("<KeyRelease>", lambda e: self._fill_law())

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
