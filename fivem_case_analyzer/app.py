# -*- coding: utf-8 -*-
"""
织梦星STAR · 案件罪名分析器  —  Windows 桌面版 (Tkinter, 少女粉)
========================================================================
一句话自动算罪：刑事罪名(刑期/罚款/保释金 + 可判置信度%) + 内务违规 +
服务器RP违规 三层；另含 医疗费用计算器、罪名速查、规章速查。离线运行。
原创制作：袁尘。
UI：柔和粉色卡片风，克制的轻特效（按钮轻浮起、输入柔光、星点缓慢呼吸）。
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


# 🎀 少女芭比粉主题
NAVY = "#C2185B"; HOT = "#FF7EB3"; HOT2 = "#FF9ECE"; PRESS = "#E0518A"
PINK = "#FFD1DC"; BG = "#FFF0F5"; CARD = "#FFFFFF"; INK = "#7A3B5D"
MUTE = "#C99BB3"; LINE = "#FAD4E2"
RED = "#D81B60"; GOLD = "#FFC83D"; WHITE = "#FFFFFF"
STEEL = "#FF7EB3"; TEAL = "#5FB0D9"; PLUM = "#B07CD6"; SRV = "#E8800C"
CHIGH = "#D81B60"; CMID = "#FF7EB3"; CLOW = "#C99BB3"
CAP = D.MAX_TOTAL_MONTHS
NONCRIM = ("内务违规", "服务器违规")
PARTY_COLOR = {"甲": STEEL, "乙": TEAL, "丙": PLUM}


def _rgb(h): return tuple(int(h[i:i + 2], 16) for i in (1, 3, 5))
def _hex(t): return "#%02X%02X%02X" % t
def lerp(a, b, t):
    ca, cb = _rgb(a), _rgb(b)
    return _hex(tuple(int(ca[i] + (cb[i] - ca[i]) * t) for i in range(3)))


def confidence(matched):
    """可判置信度: 命中关键词越多、越具体 → 越高。范围 60~98%。"""
    m = len(matched)
    lmax = max(len(k) for k in matched)
    return max(60, min(98, 72 + (m - 1) * 9 + (lmax - 2) * 4))


def conf_bar(pct):
    """5 段进度条字符，按 20% 一格。"""
    filled = max(1, min(5, round(pct / 20)))
    return "▰" * filled + "▱" * (5 - filled)


def conf_tag(pct):
    return "chigh" if pct >= 85 else ("cmid" if pct >= 70 else "clow")


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
    crim.sort(key=lambda c: -c[5])          # 置信度从高到低
    months = min(CAP, sum(c[2] for c in crim))
    years = months / 12
    years = int(years) if years == int(years) else round(years, 2)
    fine = sum(c[3] for c in crim)
    bail = round(years * fine / 6)
    return crim, vio, years, fine, bail, len(crim)


# ===================== 圆角按钮(柔和轻浮起) =====================
class RoundButton(tk.Canvas):
    def __init__(self, master, text, command, w=140, h=44,
                 base=HOT, hover=HOT2, press=PRESS, fg="white", bg=BG):
        super().__init__(master, width=w, height=h + 8, bg=bg, highlightthickness=0)
        self.command, self.base, self.hover, self.press = command, base, hover, press
        self.w, self.h = w, h
        # 柔和投影
        self.shadow = self.create_polygon(self._pts(6), smooth=True, fill=PINK)
        self.shape = self.create_polygon(self._pts(0), smooth=True, fill=base)
        self.txt = self.create_text(w / 2, h / 2, text=text, fill=fg,
                                    font=("Microsoft YaHei UI", 12, "bold"))
        self.bind("<Enter>", lambda e: self._set(self.hover, 0))
        self.bind("<Leave>", lambda e: self._set(self.base, 0))
        self.bind("<ButtonPress-1>", lambda e: self._set(self.press, 4))
        self.bind("<ButtonRelease-1>", self._release)

    def _pts(self, dy):
        w, h, r = self.w, self.h, 16
        return [r, dy, w - r, dy, w, dy, w, dy + r, w, dy + h - r, w, dy + h,
                w - r, dy + h, r, dy + h, 0, dy + h, 0, dy + h - r, 0, dy + r, 0, dy]

    def _set(self, color, dy):
        self.coords(self.shape, *self._pts(dy))
        self.itemconfig(self.shape, fill=color)
        self.coords(self.txt, self.w / 2, dy + self.h / 2)

    def _release(self, e):
        self._set(self.hover, 0)
        if 0 <= e.x <= self.w and 0 <= e.y <= self.h + 8 and self.command:
            self.command()


# ===================== 主程序 =====================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("✨ 织梦星STAR · 案件罪名分析器 ✨")
        self.geometry("980x760")
        self.configure(bg=BG)
        try:
            self.iconbitmap(resource_path("icon.ico"))
        except Exception:
            pass
        self.option_add("*Font", ("Microsoft YaHei UI", 10))
        self._init_style()
        self._build_banner()

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=8)
        self._build_analyzer(nb)
        self._build_medical(nb)
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
        st.configure("Treeview", background="white", fieldbackground="white",
                     foreground=INK, rowheight=26, borderwidth=0)
        st.configure("Treeview.Heading", background=HOT, foreground="white",
                     font=("Microsoft YaHei UI", 10, "bold"), relief="flat")
        st.configure("TScrollbar", background=PINK, troughcolor=BG, borderwidth=0)
        st.configure("TSpinbox", fieldbackground="#FFF5F9", arrowcolor=NAVY)

    # ---------- 横幅(柔光渐变 + 缓慢呼吸星点) ----------
    def _build_banner(self):
        c = tk.Canvas(self, height=96, bg=NAVY, highlightthickness=0)
        c.pack(fill="x")
        self.banner = c
        self.update_idletasks()
        W = self.winfo_width() or 980
        # 柔和横向渐变背景
        for i in range(0, W, 6):
            col = lerp(NAVY, "#E2568E", i / max(1, W))
            c.create_rectangle(i, 0, i + 6, 96, fill=col, outline=col, tags="grad")
        try:
            self._logo = tk.PhotoImage(file=resource_path("logo.png"))
            c.create_image(50, 48, image=self._logo)
        except Exception:
            self._logo = None
        c.create_text(100, 36, anchor="w", text="织梦星 STAR · 案件罪名分析器",
                      fill="white", font=("Microsoft YaHei UI", 19, "bold"))
        c.create_text(102, 67, anchor="w",
                      text="圣安地列斯州 · STAR Roleplay　💗　一句话自动算罪",
                      fill="#FFE3EE", font=("Microsoft YaHei UI", 10))
        # 缓慢呼吸的小星点(克制)
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
                tw = 0.5 + 0.5 * math.sin(self._t * 0.07 + ph)   # 慢速呼吸
                self.banner.coords(sid, *self._star_pts(sx, sy, 5 + 3 * tw))
                self.banner.itemconfig(sid, fill=lerp("#FFE9A0", GOLD, tw))
        except tk.TclError:
            return
        self.after(90, self._animate)

    # ---------- ① 算罪台 ----------
    def _build_analyzer(self, nb):
        f = tk.Frame(nb, bg=BG); nb.add(f, text="  算罪台  ")
        tk.Label(f, text="🎀 ① 把每个人做了什么，用一句话写进框 → 自动算罪（边打边出）",
                 bg=PINK, fg=NAVY, font=("Microsoft YaHei UI", 10, "bold"),
                 anchor="w", padx=10, pady=5).pack(fill="x", padx=8, pady=(8, 4))

        self.inputs = {}
        wrap = tk.Frame(f, bg=BG); wrap.pack(fill="x", padx=8)
        for p in ("甲", "乙", "丙"):
            row = tk.Frame(wrap, bg=BG); row.pack(fill="x", pady=3)
            tk.Label(row, text=f"当事人{p}", bg=PARTY_COLOR[p], fg="white",
                     width=9, font=("Microsoft YaHei UI", 10, "bold")).pack(side="left", fill="y")
            t = tk.Text(row, height=2, wrap="word", bg="#FFFDF8", relief="flat",
                        font=("Microsoft YaHei UI", 10), highlightthickness=2,
                        highlightbackground=LINE, highlightcolor="#FF8FBF")
            t.pack(side="left", fill="x", expand=True, padx=(6, 0))
            t.bind("<KeyRelease>", self._on_change)
            self.inputs[p] = t
        self.inputs["甲"].insert("1.0", "嫌疑人持枪抢劫便利店，被警察拦下后拒捕并开枪，然后驾车逃跑还撞坏了路边的车。")

        br = tk.Frame(f, bg=BG); br.pack(fill="x", padx=14, pady=(6, 0))
        RoundButton(br, "🔍 分析", lambda: self._render(flash=True), w=130).pack(side="left")
        RoundButton(br, "🗑 清空", self._clear, w=110, base=PLUM, hover="#C79AE0",
                    press="#8E5BB0").pack(side="left", padx=10)

        tk.Label(f, text="💗 ② 分析结果（按可判置信度排序）", bg=PINK, fg=NAVY,
                 font=("Microsoft YaHei UI", 10, "bold"),
                 anchor="w", padx=10, pady=5).pack(fill="x", padx=8, pady=(8, 4))
        self.out = tk.Text(f, wrap="word", bg="white", relief="flat", state="disabled",
                           font=("Microsoft YaHei UI", 10), padx=12, pady=10,
                           highlightthickness=2, highlightbackground=LINE)
        self.out.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.out.tag_config("h", foreground="white", background=NAVY,
                            font=("Microsoft YaHei UI", 11, "bold"), spacing1=5, spacing3=5)
        self.out.tag_config("cnt", foreground=NAVY, font=("Microsoft YaHei UI", 10, "bold"))
        self.out.tag_config("name", foreground=INK, font=("Microsoft YaHei UI", 10, "bold"))
        self.out.tag_config("chigh", foreground=CHIGH, font=("Consolas", 11, "bold"))
        self.out.tag_config("cmid", foreground=CMID, font=("Consolas", 11, "bold"))
        self.out.tag_config("clow", foreground=CLOW, font=("Consolas", 11, "bold"))
        self.out.tag_config("num", foreground=RED, font=("Microsoft YaHei UI", 12, "bold"))
        self.out.tag_config("lbl", foreground=INK)
        self.out.tag_config("vio", foreground=SRV)
        self.out.tag_config("disc", foreground=PLUM)
        self.out.tag_config("mute", foreground=MUTE)
        self.out.tag_config("verdict", foreground="white", background=NAVY,
                            font=("Microsoft YaHei UI", 11, "bold"), spacing1=7, spacing3=7)
        self._render()

    def _clear(self):
        for t in self.inputs.values():
            t.delete("1.0", "end")
        self._render()

    def _on_change(self, _evt=None):
        if getattr(self, "_job", None):
            self.after_cancel(self._job)
        self._job = self.after(200, self._render)

    def _render(self, flash=False):
        self.out.config(state="normal")
        self.out.delete("1.0", "end")
        best = (None, -1)
        any_in = False
        for p in ("甲", "乙", "丙"):
            txt = self.inputs[p].get("1.0", "end-1c").strip()
            if not txt:
                continue
            any_in = True
            crim, vio, years, fine, bail, n = analyze(txt)
            self.out.insert("end", f" 当事人{p} \n", "h")
            if crim:
                self.out.insert("end", "  刑事罪名 ")
                self.out.insert("end", f"共 {n} 条", "cnt")
                self.out.insert("end", "：\n", "lbl")
                for name, kind, mo, fn, note, pct in crim:
                    tg = conf_tag(pct)
                    self.out.insert("end", f"     • {name}", "name")
                    self.out.insert("end", f"　{conf_bar(pct)} {pct}%", tg)
                    self.out.insert("end", "  可判\n", "mute")
                self.out.insert("end", "  合计刑期 ", "lbl")
                self.out.insert("end", f"{years}", "num")
                self.out.insert("end", " 年   罚款 ", "lbl")
                self.out.insert("end", f"${fine:,}", "num")
                self.out.insert("end", "   建议保释金 ", "lbl")
                self.out.insert("end", f"${bail:,}\n", "num")
            else:
                self.out.insert("end", "  刑事罪名：（未识别到，请写更具体或用法典里的词）\n", "mute")
            if vio:
                self.out.insert("end", "  违规 / 处分：\n", "lbl")
                for name, note in vio:
                    tag = "disc" if ("内务" in note or "FBI" in note or "开除" in note) else "vio"
                    self.out.insert("end", f"     • {name}　{note}\n", tag)
            self.out.insert("end", "\n")
            if years > best[1]:
                best = (p, years)
        if best[0] and best[1] > 0:
            self.out.insert("end",
                            f" 🔨 主要责任方：当事人{best[0]}  |  合计刑期 {best[1]} 年"
                            f"  |  仅供参考，正当防卫/堡垒原则等请人工复核 \n", "verdict")
            if flash:
                self._flash(0)
        elif not any_in:
            self.out.insert("end", "请在上方输入案情……", "mute")
        self.out.config(state="disabled")

    def _flash(self, step):
        cols = ["#FF9ECE", "#FF7EB3", "#E0518A", NAVY]
        if step < len(cols) and self._alive:
            try:
                self.out.tag_config("verdict", background=cols[step])
            except tk.TclError:
                return
            self.after(90, lambda: self._flash(step + 1))

    # ---------- ② 医疗费用计算器 ----------
    def _build_medical(self, nb):
        f = tk.Frame(nb, bg=BG); nb.add(f, text="  医疗费用计算器  ")
        top = tk.Frame(f, bg=HOT); top.pack(fill="x", padx=8, pady=8)
        tk.Label(top, text="💗 本次合计：", bg=HOT, fg="white",
                 font=("Microsoft YaHei UI", 12, "bold")).pack(side="left", padx=10, pady=7)
        self.med_total = tk.Label(top, text="$0", bg=HOT, fg="white",
                                  font=("Microsoft YaHei UI", 14, "bold"))
        self.med_total.pack(side="left")

        canvas = tk.Canvas(f, bg=BG, highlightthickness=0)
        sb = ttk.Scrollbar(f, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=BG)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True, padx=(8, 0))
        sb.pack(side="right", fill="y")

        hdr = tk.Frame(inner, bg=HOT); hdr.pack(fill="x")
        for txt, w in (("项目", 22), ("单价", 10), ("数量", 6), ("小计", 12), ("说明", 30)):
            tk.Label(hdr, text=txt, bg=HOT, fg="white", width=w,
                     font=("Microsoft YaHei UI", 10, "bold")).pack(side="left")

        self.med_rows = []
        cur = None
        for cat, name, price, note in D.MED_ITEMS:
            if cat != cur:
                cur = cat
                tk.Label(inner, text="▍ " + cat, bg=NAVY, fg="white", anchor="w",
                         font=("Microsoft YaHei UI", 9, "bold")).pack(fill="x")
            row = tk.Frame(inner, bg="white"); row.pack(fill="x")
            tk.Label(row, text=name, bg="white", width=22, anchor="w", fg=INK).pack(side="left")
            tk.Label(row, text=f"${price:,}", bg="white", width=10, anchor="e", fg=INK).pack(side="left")
            var = tk.IntVar(value=0)
            ttk.Spinbox(row, from_=0, to=999, width=5, textvariable=var,
                        command=self._update_med).pack(side="left", padx=2)
            sub = tk.Label(row, text="$0", bg="white", width=12, anchor="e", fg=RED)
            sub.pack(side="left")
            tk.Label(row, text=note, bg="white", anchor="w", fg=MUTE).pack(side="left", padx=4)
            var.trace_add("write", lambda *a: self._update_med())
            self.med_rows.append((price, var, sub))

        for n in D.MED_NOTES:
            tk.Label(inner, text="  " + n, bg="#FFF6E9", fg="#A8741B", anchor="w",
                     wraplength=880, justify="left").pack(fill="x", pady=1)

    def _update_med(self):
        total = 0
        for price, var, sub in self.med_rows:
            try:
                q = int(var.get())
            except (tk.TclError, ValueError):
                q = 0
            s = price * q
            total += s
            sub.config(text=f"${s:,}")
        self.med_total.config(text=f"${total:,}")

    # ---------- ③ 罪名速查 ----------
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

    # ---------- ④ 规章速查 ----------
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
