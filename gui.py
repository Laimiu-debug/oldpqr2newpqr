# -*- coding: utf-8 -*-
"""
gui.py
======
PQR 文档批量转化工具 —— Tkinter 可视化前端。

启动：双击 run.bat 或  python gui.py
"""
import os
import sys
import csv
import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import win32com.client as win32

import convert_core as core


def _launch_office_app():
    """
    自动检测并启动可用的 Office 组件。
    优先顺序：Microsoft Word → WPS Office。
    返回 (app对象, 显示名称)。都不可用则抛异常。
    """
    # 候选 ProgID 列表（优先级从高到低）
    candidates = [
        ("Word.Application", "Microsoft Word"),
        ("KWPS.Application", "WPS Office"),
    ]
    errors = []
    for progid, name in candidates:
        try:
            app = win32.gencache.EnsureDispatch(progid)
            return app, name
        except Exception as e:
            errors.append(f"{name}({progid}): {e}")
    raise RuntimeError("未找到可用的 Office 程序 | " + "; ".join(errors))

# ---------------------------------------------------------------------------
# 路径
# ---------------------------------------------------------------------------
# PyInstaller 打包后，模板等外部文件应放在 exe 同级目录
if getattr(sys, 'frozen', False):
    # 打包模式：exe 所在目录
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    # 开发模式：脚本所在目录
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(BASE_DIR, "PQR新格式空白.docx")
DEFAULT_OUTPUT = os.path.join(BASE_DIR, "输出")
DEFAULT_LOGS = os.path.join(BASE_DIR, "logs")

# 默认命名格式
DEFAULT_NAMING = "PQR{编号}_新格式"


class App:
    def __init__(self, root):
        self.root = root
        root.title("PQR 文档批量转化工具")
        root.geometry("980x680")
        root.minsize(900, 620)

        # 状态
        self.current_dir = tk.StringVar(value=BASE_DIR)
        self.output_dir = tk.StringVar(value=DEFAULT_OUTPUT)
        self.naming = tk.StringVar(value=DEFAULT_NAMING)
        self.search = tk.StringVar()
        self.file_items = []          # [(full_path, display_name), ...]
        self.checked = {}             # {full_path: BooleanVar}
        self.convert_queue = queue.Queue()  # 后台线程 → 主线程 的消息队列
        self.converting = False
        self.latest_report = tk.StringVar(value="")
        self.copy_sketch = tk.BooleanVar(value=True)  # 是否复制简图，默认开

        self._build_ui()
        self._refresh_list()

        # 轮询消息队列
        self.root.after(100, self._drain_queue)

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        pad = {"padx": 6, "pady": 4}

        # ---- 顶部：文件夹选择 ----
        top = ttk.Frame(self.root)
        top.pack(fill="x", padx=10, pady=(10, 4))
        ttk.Button(top, text="选择文件夹", command=self._choose_dir).pack(side="left")
        ttk.Entry(top, textvariable=self.current_dir).pack(
            side="left", fill="x", expand=True, padx=8
        )
        ttk.Button(top, text="刷新列表", command=self._refresh_list).pack(side="left")

        # ---- 中部：列表 + 已选 ----
        mid = ttk.Frame(self.root)
        mid.pack(fill="both", expand=True, padx=10, pady=4)

        # 左：文档列表
        left = ttk.LabelFrame(mid, text="文档列表（勾选要转化的文档）")
        left.pack(side="left", fill="both", expand=True)
        # 搜索框
        sf = ttk.Frame(left)
        sf.pack(fill="x", padx=4, pady=2)
        ttk.Label(sf, text="搜索:").pack(side="left")
        ttk.Entry(sf, textvariable=self.search).pack(
            side="left", fill="x", expand=True, padx=4
        )
        self.search.trace_add("write", lambda *_: self._apply_filter())

        # 列表滚动
        list_frame = ttk.Frame(left)
        list_frame.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(list_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.canvas.yview)
        self.scroll_frame = ttk.Frame(self.canvas)
        self.scroll_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self.canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        # 鼠标滚轮
        self.canvas.bind_all(
            "<MouseWheel>",
            lambda e: self.canvas.yview_scroll(int(-e.delta / 120), "units"),
        )

        # 列表操作按钮
        bf = ttk.Frame(left)
        bf.pack(fill="x", padx=4, pady=4)
        ttk.Button(bf, text="全选", command=self._select_all).pack(side="left")
        ttk.Button(bf, text="全不选", command=self._select_none).pack(side="left")
        ttk.Button(bf, text="仅 .doc", command=self._select_doc_only).pack(side="left")
        self.count_lbl = ttk.Label(bf, text="共 0 份，已选 0 份")
        self.count_lbl.pack(side="right")

        # 右：已选清单
        right = ttk.LabelFrame(mid, text="已选文档", width=240)
        right.pack(side="right", fill="y")
        self.selected_text = tk.Text(right, width=30, height=20, state="disabled")
        self.selected_text.pack(fill="both", expand=True, padx=4, pady=4)

        # ---- 下部：命名 + 输出 ----
        bot = ttk.LabelFrame(self.root, text="输出设置")
        bot.pack(fill="x", padx=10, pady=4)
        r1 = ttk.Frame(bot)
        r1.pack(fill="x", padx=8, pady=4)
        ttk.Label(r1, text="命名格式:").pack(side="left")
        ttk.Entry(r1, textvariable=self.naming, width=28).pack(side="left", padx=4)
        ttk.Label(
            r1,
            text="可用: {编号} {焊接方法} {原文件名}   示例: "
            + self.naming.get()
            + ".docx",
            foreground="gray",
        ).pack(side="left")

        r2 = ttk.Frame(bot)
        r2.pack(fill="x", padx=8, pady=(0, 4))
        ttk.Label(r2, text="输出目录:").pack(side="left")
        ttk.Entry(r2, textvariable=self.output_dir).pack(
            side="left", fill="x", expand=True, padx=4
        )
        ttk.Button(r2, text="更改", command=self._choose_output).pack(side="left")

        # 选项行
        r3 = ttk.Frame(bot)
        r3.pack(fill="x", padx=8, pady=(0, 4))
        ttk.Checkbutton(
            r3,
            text="复制简图（把原图纸简图复制到新文档）",
            variable=self.copy_sketch,
        ).pack(side="left")

        # ---- 转化按钮 + 进度 ----
        act = ttk.Frame(self.root)
        act.pack(fill="x", padx=10, pady=4)
        self.start_btn = ttk.Button(act, text="▶  开始转化", command=self._start_convert)
        self.start_btn.pack(side="left")
        self.progress = ttk.Progressbar(act, mode="determinate", length=400)
        self.progress.pack(side="left", fill="x", expand=True, padx=10)
        self.prog_lbl = ttk.Label(act, text="就绪")
        self.prog_lbl.pack(side="left")

        # ---- 日志 ----
        lf = ttk.LabelFrame(self.root, text="实时日志")
        lf.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        self.log_text = tk.Text(lf, height=10, state="disabled")
        self.log_text.pack(side="left", fill="both", expand=True)
        log_sb = ttk.Scrollbar(lf, orient="vertical", command=self.log_text.yview)
        log_sb.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=log_sb.set)

        # 日志颜色标签
        self.log_text.tag_configure("ok", foreground="#107c10")
        self.log_text.tag_configure("warn", foreground="#c47b00")
        self.log_text.tag_configure("err", foreground="#c50f1f")
        self.log_text.tag_configure("info", foreground="#444444")

        # 底部报告按钮
        bottom = ttk.Frame(self.root)
        bottom.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(bottom, text="查看转化报告", command=self._open_report).pack(side="left")
        ttk.Button(bottom, text="打开输出目录", command=self._open_output).pack(side="left")
        ttk.Button(
            bottom, text="打开使用说明", command=self._open_readme
        ).pack(side="left")

    # ------------------------------------------------------------ 文件列表
    def _choose_dir(self):
        d = filedialog.askdirectory(initialdir=self.current_dir.get())
        if d:
            self.current_dir.set(d)
            self._refresh_list()

    def _refresh_list(self):
        d = self.current_dir.get()
        # 清空旧控件
        for w in self.scroll_frame.winfo_children():
            w.destroy()
        self.file_items = []
        self.checked = {}

        if not os.path.isdir(d):
            self.count_lbl.config(text="目录不存在")
            return

        exts = (".doc", ".docx")
        # 排除模板自身
        try:
            files = [
                f
                for f in sorted(os.listdir(d))
                if f.lower().endswith(exts)
                and os.path.isfile(os.path.join(d, f))
                and os.path.abspath(os.path.join(d, f)) != os.path.abspath(TEMPLATE)
            ]
        except Exception as e:
            self.count_lbl.config(text=f"读取失败: {e}")
            return

        for f in files:
            full = os.path.join(d, f)
            disp = f
            self.file_items.append((full, disp))
            var = tk.BooleanVar(value=False)
            self.checked[full] = var
            cb = ttk.Checkbutton(self.scroll_frame, text=disp, variable=var)
            cb.pack(anchor="w", padx=4, pady=1)
            var.trace_add("write", lambda *_: self._update_selected())

        self._apply_filter()
        self._update_selected()

    def _apply_filter(self):
        kw = self.search.get().strip().lower()
        for w in self.scroll_frame.winfo_children():
            # w 是 Checkbutton
            txt = w.cget("text").lower()
            if kw and kw not in txt:
                w.pack_forget()
            else:
                w.pack(anchor="w", padx=4, pady=1)

    def _select_all(self):
        for v in self.checked.values():
            v.set(True)

    def _select_none(self):
        for v in self.checked.values():
            v.set(False)

    def _select_doc_only(self):
        for full, v in self.checked.items():
            v.set(full.lower().endswith(".doc"))

    def _update_selected(self):
        sel = [disp for full, disp in self.file_items if self.checked.get(full, tk.BooleanVar()).get()]
        self.selected_text.configure(state="normal")
        self.selected_text.delete("1.0", "end")
        self.selected_text.insert("1.0", "\n".join(sel))
        self.selected_text.configure(state="disabled")
        total = len(self.file_items)
        self.count_lbl.config(text=f"共 {total} 份，已选 {len(sel)} 份")

    def _choose_output(self):
        d = filedialog.askdirectory(initialdir=self.output_dir.get())
        if d:
            self.output_dir.set(d)

    # -------------------------------------------------------------- 转化
    def _start_convert(self):
        if self.converting:
            return
        selected = [
            full for full, disp in self.file_items if self.checked.get(full, tk.BooleanVar()).get()
        ]
        if not selected:
            messagebox.showwarning("提示", "请先勾选要转化的文档。")
            return
        if not os.path.exists(TEMPLATE):
            messagebox.showerror("错误", f"未找到模板文件：\n{TEMPLATE}")
            return

        out_dir = self.output_dir.get()
        os.makedirs(out_dir, exist_ok=True)
        os.makedirs(DEFAULT_LOGS, exist_ok=True)

        self.converting = True
        self.start_btn.config(state="disabled", text="转化中…")
        self.progress["value"] = 0
        self.progress["maximum"] = len(selected)
        self._log_clear()
        self._log(f"开始转化 {len(selected)} 份文档…", "info")
        self._log(f"模板: {TEMPLATE}", "info")
        self._log(f"输出: {out_dir}", "info")
        do_sketch = self.copy_sketch.get()
        self._log(f"复制简图: {'是' if do_sketch else '否'}", "info")
        self._log("-" * 50, "info")

        t = threading.Thread(
            target=self._convert_worker,
            args=(selected, out_dir, self.naming.get(), do_sketch),
            daemon=True,
        )
        t.start()

    def _convert_worker(self, selected, out_dir, naming, do_sketch=True):
        """后台线程：复用一个 Word/WPS 实例转化全部文档。"""
        word_app = None
        results = []
        office_name = ""
        try:
            # 自动检测：优先 Word，其次 WPS
            word_app, office_name = _launch_office_app()
            word_app.Visible = False
            try:
                word_app.DisplayAlerts = False
            except Exception:
                pass
            try:
                word_app.ScreenUpdating = False
            except Exception:
                pass
            self.convert_queue.put(("log", f"已启动 {office_name}", "info"))
        except Exception as e:
            self.convert_queue.put(("log", f"✗ 无法启动 Word 或 WPS: {e}", "err"))
            self.convert_queue.put(("log", "请确认电脑已安装 Microsoft Word 或 WPS Office", "err"))
            self.convert_queue.put(("done", None))
            return

        try:
            for i, in_path in enumerate(selected):
                fname = os.path.basename(in_path)
                self.convert_queue.put(
                    ("progress", (i, len(selected), fname))
                )

                def cb(m):
                    self.convert_queue.put(("log", m, "info"))

                r = core.convert_one(in_path, out_dir, naming, word_app, TEMPLATE, cb, copy_sketch=do_sketch)
                results.append(r)
                if r["status"] == "ok":
                    self.convert_queue.put(("log", f"✓ {r['filename']} 转化完成", "ok"))
                elif r["status"] == "warn":
                    self.convert_queue.put(
                        ("log", f"⚠ {r['filename']} 完成（{len(r['issues'])}项未匹配）", "warn")
                    )
                else:
                    self.convert_queue.put(
                        ("log", f"✗ {r['filename']} 失败: {r['message']}", "err")
                    )

            # 写报告
            rep = self._write_report(results)
            self.convert_queue.put(("log", "-" * 50, "info"))
            ok = sum(1 for r in results if r["status"] == "ok")
            warn = sum(1 for r in results if r["status"] == "warn")
            fail = sum(1 for r in results if r["status"] == "fail")
            self.convert_queue.put(
                ("log", f"完成：成功 {ok}，有警告 {warn}，失败 {fail}", "info")
            )
            self.convert_queue.put(("log", f"报告已保存: {rep}", "info"))
            self.latest_report.set(rep)
        finally:
            try:
                if word_app:
                    word_app.Quit()
            except Exception:
                pass
            self.convert_queue.put(("done", None))

    def _unique_path(self, path):
        """若文件已存在，加 (2) (3) 后缀。"""
        if not os.path.exists(path):
            return path
        base, ext = os.path.splitext(path)
        n = 2
        while os.path.exists(f"{base}({n}){ext}"):
            n += 1
        return f"{base}({n}){ext}"

    def _write_report(self, results):
        import datetime

        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = os.path.join(DEFAULT_LOGS, f"转化报告_{ts}.csv")
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["文件名", "状态", "未匹配字段数", "未匹配字段", "失败原因"])
            for r in results:
                status_cn = {"ok": "成功", "warn": "部分成功", "fail": "失败"}.get(
                    r["status"], r["status"]
                )
                fields = ", ".join(i["field"] for i in r["issues"])
                w.writerow(
                    [
                        r["filename"],
                        status_cn,
                        len(r["issues"]),
                        fields,
                        r["message"] if r["status"] == "fail" else "",
                    ]
                )
        return path

    # -------------------------------------------------------------- 队列
    def _drain_queue(self):
        try:
            while True:
                item = self.convert_queue.get_nowait()
                kind = item[0]
                if kind == "log":
                    self._log(item[1], item[2])
                elif kind == "progress":
                    i, total, fname = item[1]
                    self.progress["value"] = i
                    self.prog_lbl.config(text=f"{i}/{total}  {fname[:30]}")
                elif kind == "done":
                    self.progress["value"] = self.progress["maximum"]
                    self.prog_lbl.config(text="完成")
                    self.converting = False
                    self.start_btn.config(state="normal", text="▶  开始转化")
                    messagebox.showinfo("完成", "转化已完成，详见日志和报告。")
        except queue.Empty:
            pass
        self.root.after(100, self._drain_queue)

    # -------------------------------------------------------------- 日志
    def _log(self, msg, tag="info"):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n", tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _log_clear(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    # -------------------------------------------------------------- 打开
    def _open_report(self):
        if not self.latest_report.get():
            # 找最新的报告
            if os.path.isdir(DEFAULT_LOGS):
                reps = sorted(
                    [f for f in os.listdir(DEFAULT_LOGS) if f.endswith(".csv")]
                )
                if reps:
                    self.latest_report.set(os.path.join(DEFAULT_LOGS, reps[-1]))
        p = self.latest_report.get()
        if p and os.path.exists(p):
            os.startfile(p)
        else:
            messagebox.showinfo("提示", "暂无转化报告。")

    def _open_output(self):
        d = self.output_dir.get()
        os.makedirs(d, exist_ok=True)
        os.startfile(d)

    def _open_readme(self):
        p = os.path.join(BASE_DIR, "README.md")
        if os.path.exists(p):
            os.startfile(p)
        else:
            messagebox.showinfo("提示", "未找到 README.md")


def main():
    root = tk.Tk()
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
