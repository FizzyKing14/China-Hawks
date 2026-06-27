# -*- coding: utf-8 -*-
"""
圣安地列斯州 · 案件罪名分析器  —  Windows 桌面版 (Tkinter)
=========================================================
一句话自动算罪：刑事罪名(刑期/罚款/保释金) + 内务违规 + 服务器RP违规 三层识别，
另含 医疗费用计算器、罪名速查、规章速查。离线运行，原创制作：袁尘。

直接运行:  python app.py
打包 exe :  见 build_exe.bat 或 .github/workflows/build-windows.yml
"""
import tkinter as tk
from tkinter import ttk

# 复用 build_workbook 里的全部法条/关键词/价格/规章数据（已合并去重）
import build_workbook as D

NAVY = "#12243B"; BLUE = "#1F3A5F"; GOLD = "#C9A227"; GOLDL = "#F3E6BE"
PAGE = "#F4F6FA"; INK = "#1B2A41"; MUTE = "#6B7280"; RED = "#C0392B"
STEEL = "#2C4A7C"; TEAL = "#1E6B5C"; PLUM = "#5E4B8B"; SRV = "#B5731B"
CAP = D.MAX_TOTAL_MONTHS
NONCRIM = ("内务违规", "服务器违规")
PARTY_COLOR = {"甲": STEEL, "乙": TEAL, "丙": PLUM}


def analyze(text):
    """返回 (刑事[(name,kind,months,fine,note)], 违规[(name,note)], 月, 罚款, 保释金, 罪名数)"""
    text = text or ""
    crim, vio = [], []
    if text.strip():
        for chap, name, kind, months, fine, kws, note in D.CHARGES:
            if any(k in text for k in kws):
                if kind in NONCRIM:
                    vio.append((name, note))
                else:
                    crim.append((name, kind, months, fine, note))
    months = min(CAP, sum(c[2] for c in crim))
    fine = sum(c[3] for c in crim)
    bail = round(months * fine / 6)
    return crim, vio, months, fine, bail, len(crim)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("圣安地列斯州 · 案件罪名分析器")
        self.geometry("960x700")
        self.configure(bg=PAGE)
        self.option_add("*Font", ("Microsoft YaHei UI", 10))

        # 顶部横幅
        banner = tk.Frame(self, bg=NAVY)
        banner.pack(fill="x")
        tk.Label(banner, text="⚖   圣安地列斯州 · 案件罪名分析器   🦅",
                 bg=NAVY, fg="white", font=("Microsoft YaHei UI", 17, "bold"),
                 pady=10).pack()
        tk.Frame(self, bg=GOLD, height=3).pack(fill="x")

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=6)
        self._build_analyzer(nb)
        self._build_medical(nb)
        self._build_law(nb)
        self._build_rules(nb)

        tk.Label(self, text="原创制作：袁尘", bg=PAGE, fg=NAVY,
                 font=("Microsoft YaHei UI", 9, "bold"), anchor="e").pack(fill="x", padx=12, pady=2)

    # ---------------- ① 算罪台 ----------------
    def _build_analyzer(self, nb):
        f = tk.Frame(nb, bg=PAGE); nb.add(f, text="  算罪台  ")
        tk.Label(f, text="① 把每个人做了什么，用一句话写进框里 → 自动算罪（边打边出结果）",
                 bg=GOLDL, fg=NAVY, font=("Microsoft YaHei UI", 10, "bold"),
                 anchor="w", padx=8, pady=4).pack(fill="x", padx=6, pady=(6, 2))

        self.inputs = {}
        inp_wrap = tk.Frame(f, bg=PAGE); inp_wrap.pack(fill="x", padx=6)
        for p in ("甲", "乙", "丙"):
            row = tk.Frame(inp_wrap, bg=PAGE); row.pack(fill="x", pady=2)
            tk.Label(row, text=f"当事人{p}", bg=PARTY_COLOR[p], fg="white",
                     width=9, font=("Microsoft YaHei UI", 10, "bold")).pack(side="left", fill="y")
            t = tk.Text(row, height=2, wrap="word", bg="#FFFDF5", relief="solid", bd=1,
                        font=("Microsoft YaHei UI", 10))
            t.pack(side="left", fill="x", expand=True, padx=(4, 0))
            t.bind("<KeyRelease>", self._on_change)
            self.inputs[p] = t
        self.inputs["甲"].insert("1.0", "嫌疑人持枪抢劫便利店，被警察拦下后拒捕并开枪，然后驾车逃跑还撞坏了路边的车。")

        tk.Label(f, text="② 分析结果", bg=GOLDL, fg=NAVY, font=("Microsoft YaHei UI", 10, "bold"),
                 anchor="w", padx=8, pady=4).pack(fill="x", padx=6, pady=(8, 2))

        self.out = tk.Text(f, wrap="word", bg="white", relief="solid", bd=1, state="disabled",
                           font=("Microsoft YaHei UI", 10), padx=10, pady=8)
        self.out.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        self.out.tag_config("h", foreground="white", background=NAVY,
                            font=("Microsoft YaHei UI", 11, "bold"), spacing1=4, spacing3=4)
        self.out.tag_config("crim", foreground=RED, font=("Microsoft YaHei UI", 10, "bold"))
        self.out.tag_config("num", foreground=RED, font=("Microsoft YaHei UI", 12, "bold"))
        self.out.tag_config("vio", foreground=SRV)
        self.out.tag_config("disc", foreground=PLUM)
        self.out.tag_config("mute", foreground=MUTE)
        self.out.tag_config("verdict", foreground="white", background=NAVY,
                            font=("Microsoft YaHei UI", 11, "bold"), spacing1=6, spacing3=6)
        self._render()

    def _on_change(self, _evt=None):
        if getattr(self, "_job", None):
            self.after_cancel(self._job)
        self._job = self.after(200, self._render)

    def _render(self):
        self.out.config(state="normal")
        self.out.delete("1.0", "end")
        best = (None, -1)
        for p in ("甲", "乙", "丙"):
            txt = self.inputs[p].get("1.0", "end-1c").strip()
            if not txt:
                continue
            crim, vio, months, fine, bail, n = analyze(txt)
            self.out.insert("end", f" 当事人{p} \n", "h")
            if crim:
                self.out.insert("end", "  刑事罪名（%d）：" % n)
                self.out.insert("end", "、".join(c[0] for c in crim) + "\n", "crim")
                self.out.insert("end", "  合计刑期 ")
                self.out.insert("end", f"{months}", "num")
                self.out.insert("end", " 个月  |  罚款 ")
                self.out.insert("end", f"${fine:,}", "num")
                self.out.insert("end", "  |  建议保释金 ")
                self.out.insert("end", f"${bail:,}\n", "num")
            else:
                self.out.insert("end", "  刑事罪名：（未识别到，请写更具体或用法典里的词）\n", "mute")
            if vio:
                self.out.insert("end", "  违规 / 处分：\n")
                for name, note in vio:
                    tag = "disc" if "内务" in note or "FBI" in note or "开除" in note else "vio"
                    self.out.insert("end", f"     • {name}　{note}\n", tag)
            self.out.insert("end", "\n")
            if months > best[1]:
                best = (p, months)
        if best[0] and best[1] > 0:
            self.out.insert("end",
                            f" 🔨 主要责任方：当事人{best[0]}  |  合计刑期 {best[1]} 个月"
                            f"（{round(best[1]/12,1)} 年）  |  仅供参考，正当防卫/堡垒原则等请人工复核 \n",
                            "verdict")
        elif not any(self.inputs[p].get("1.0", "end-1c").strip() for p in ("甲", "乙", "丙")):
            self.out.insert("end", "请在上方输入案情……", "mute")
        self.out.config(state="disabled")

    # ---------------- ② 医疗费用计算器 ----------------
    def _build_medical(self, nb):
        f = tk.Frame(nb, bg=PAGE); nb.add(f, text="  医疗费用计算器  ")
        top = tk.Frame(f, bg=GOLD); top.pack(fill="x", padx=6, pady=6)
        tk.Label(top, text="💴 本次合计：", bg=GOLD, fg="white",
                 font=("Microsoft YaHei UI", 12, "bold")).pack(side="left", padx=8, pady=6)
        self.med_total = tk.Label(top, text="$0", bg=GOLD, fg="white",
                                  font=("Microsoft YaHei UI", 14, "bold"))
        self.med_total.pack(side="left")

        canvas = tk.Canvas(f, bg=PAGE, highlightthickness=0)
        sb = ttk.Scrollbar(f, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=PAGE)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True, padx=(6, 0))
        sb.pack(side="right", fill="y")

        hdr = tk.Frame(inner, bg=BLUE); hdr.pack(fill="x")
        for txt, w in (("项目", 22), ("单价", 10), ("数量", 6), ("小计", 12), ("说明", 30)):
            tk.Label(hdr, text=txt, bg=BLUE, fg="white", width=w,
                     font=("Microsoft YaHei UI", 10, "bold")).pack(side="left")

        self.med_rows = []
        cur = None
        for cat, name, price, note in D.MED_ITEMS:
            if cat != cur:
                cur = cat
                tk.Label(inner, text="▍ " + cat, bg=NAVY, fg="white", anchor="w",
                         font=("Microsoft YaHei UI", 9, "bold")).pack(fill="x")
            row = tk.Frame(inner, bg="white"); row.pack(fill="x")
            tk.Label(row, text=name, bg="white", width=22, anchor="w").pack(side="left")
            tk.Label(row, text=f"${price:,}", bg="white", width=10, anchor="e").pack(side="left")
            var = tk.IntVar(value=0)
            sp = ttk.Spinbox(row, from_=0, to=999, width=5, textvariable=var,
                             command=self._update_med)
            sp.pack(side="left", padx=2)
            sub = tk.Label(row, text="$0", bg="white", width=12, anchor="e")
            sub.pack(side="left")
            tk.Label(row, text=note, bg="white", anchor="w", fg=MUTE).pack(side="left", padx=4)
            var.trace_add("write", lambda *a: self._update_med())
            self.med_rows.append((price, var, sub))

        for n in D.MED_NOTES:
            tk.Label(inner, text="  " + n, bg="#FFF8E1", fg="#8A6D00", anchor="w",
                     wraplength=860, justify="left").pack(fill="x", pady=1)

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

    # ---------------- ③ 罪名速查 ----------------
    def _build_law(self, nb):
        f = tk.Frame(nb, bg=PAGE); nb.add(f, text="  罪名速查  ")
        bar = tk.Frame(f, bg=PAGE); bar.pack(fill="x", padx=6, pady=6)
        tk.Label(bar, text="搜索：", bg=PAGE).pack(side="left")
        self.law_q = tk.Entry(bar)
        self.law_q.pack(side="left", fill="x", expand=True)
        self.law_q.bind("<KeyRelease>", lambda e: self._fill_law())

        cols = ("章节", "罪名", "定性", "刑期(月)", "罚款($)", "说明")
        self.tree = ttk.Treeview(f, columns=cols, show="headings")
        widths = (110, 200, 80, 70, 90, 380)
        for c, w in zip(cols, widths):
            self.tree.heading(c, text=c); self.tree.column(c, width=w, anchor="w")
        vsb = ttk.Scrollbar(f, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=(0, 6))
        vsb.pack(side="right", fill="y", pady=(0, 6))
        self._fill_law()

    def _fill_law(self):
        q = self.law_q.get().strip()
        self.tree.delete(*self.tree.get_children())
        for chap, name, kind, months, fine, kws, note in D.CHARGES:
            blob = chap + name + kind + note + "".join(kws)
            if q and q not in blob:
                continue
            fee = f"{fine:,}" if fine else "—"
            mon = months if months else "—"
            self.tree.insert("", "end", values=(chap, name, kind, mon, fee, note))

    # ---------------- ④ 规章速查 ----------------
    def _build_rules(self, nb):
        f = tk.Frame(nb, bg=PAGE); nb.add(f, text="  规章速查  ")
        txt = tk.Text(f, wrap="word", bg="white", relief="solid", bd=1,
                      font=("Microsoft YaHei UI", 10), padx=10, pady=8)
        vsb = ttk.Scrollbar(f, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=vsb.set)
        txt.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
        vsb.pack(side="right", fill="y", pady=6)
        txt.tag_config("h", foreground="white", background=NAVY,
                       font=("Microsoft YaHei UI", 11, "bold"), spacing1=6, spacing3=4)
        txt.tag_config("k", foreground=BLUE, font=("Microsoft YaHei UI", 10, "bold"))
        for typ, a, b in D.RULES:
            if typ == "H":
                txt.insert("end", " " + a + " \n", "h")
            else:
                txt.insert("end", "  " + a + "：", "k")
                txt.insert("end", b + "\n")
        txt.config(state="disabled")


if __name__ == "__main__":
    App().mainloop()
