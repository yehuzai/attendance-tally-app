#!/usr/bin/env python3
"""作业辅导考勤次数统计 - desktop app.

Aggregates how many times each student appears across one or more attendance
spreadsheets and writes a single summary workbook. The counting logic lives in
attendance_tally.py (same code as the `attendance-tally` skill).

Headless self-test (no window):
    python app.py --selftest INPUT1.xlsx [INPUT2.xlsx ...] -o OUT.xlsx [--merge A=B]
"""

import collections
import os
import queue
import subprocess
import sys
import threading
import traceback

import attendance_tally as core

APP_TITLE = "作业辅导考勤次数统计"
DEFAULT_OUT_NAME = "考勤次数统计.xlsx"

if sys.platform == "darwin":
    BASE_FONT = ("PingFang SC", 13)
    MONO = ("Menlo", 12)
elif os.name == "nt":
    BASE_FONT = ("Microsoft YaHei UI", 10)
    MONO = ("Consolas", 10)
else:
    BASE_FONT = ("Sans", 10)
    MONO = ("Monospace", 10)


# ------------------------------------------------------------ small helpers ---

def open_path(path):
    if sys.platform == "darwin":
        subprocess.Popen(["open", path])
    elif os.name == "nt":
        os.startfile(path)  # noqa: S606
    else:
        subprocess.Popen(["xdg-open", path])


def reveal_path(path):
    if sys.platform == "darwin":
        subprocess.Popen(["open", "-R", path])
    elif os.name == "nt":
        subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
    else:
        open_path(os.path.dirname(path) or ".")


def parse_merges(text):
    merges = {}
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        if "=" not in line:
            raise ValueError("合并规则要写成 旧名=正确名，这一行有问题：" + line)
        a, b = line.split("=", 1)
        a, b = a.strip(), b.strip()
        if a and b:
            merges[a] = b
    return merges


def default_output(paths):
    if not paths:
        return ""
    return os.path.join(os.path.dirname(os.path.abspath(paths[0])), DEFAULT_OUT_NAME)


def run_tally(paths, out_path, merges, verify, log):
    sheets = []
    for p in paths:
        try:
            loaded = core.sheets_openpyxl(p)
        except Exception as exc:  # noqa: BLE001
            raise RuntimeError("打不开文件 %s：%s" % (os.path.basename(p), exc)) from exc
        for title, rows in loaded:
            sheets.append((os.path.basename(p), title, rows))
    counts, order, grades, notes, warnings = core.tally(sheets, merges)
    rows = core.build_rows(counts, order, grades)
    for n in notes:
        log(n)
    if not rows:
        raise RuntimeError("没有找到可统计的考勤行：需要包含“学生所属班级”列"
                           "以及表头带“学生姓名”的名单列。")
    core.write_book(rows, out_path)
    total = sum(c for _, _, c in rows)
    log("合计：%d 人，%d 次" % (len(rows), total))
    for w in warnings:
        log("警告：" + w)
    if verify:
        xml_sheets = []
        for p in paths:
            for title, rows2 in core.sheets_xml(p):
                xml_sheets.append((os.path.basename(p), title, rows2))
        counts2, _, _, _, _ = core.tally(xml_sheets, merges)
        if counts2 != counts:
            raise RuntimeError("校验失败：openpyxl 与原始 XML 两种解析结果不一致")
        got, header = core.read_back(out_path)
        want = collections.Counter({(n, l): c for n, l, c in rows})
        if got != want or tuple(header) != core.OUT_HEADERS:
            raise RuntimeError("校验失败：写出的文件与统计结果不一致")
        log("校验通过：两种解析结果与输出文件一致")
    log("完成，文件已保存：" + out_path)
    return rows


# ------------------------------------------------------------------- gui ------

def build_app():
    from tkinter import ttk
    import tkinter as tk
    from tkinter import filedialog, messagebox

    try:
        from tkinterdnd2 import TkinterDnD, DND_FILES
        root = TkinterDnD.Tk()
        dnd_ok = True
    except Exception:
        root = tk.Tk()
        dnd_ok = False

    root.title(APP_TITLE)
    root.minsize(640, 560)

    style = ttk.Style(root)
    style.configure(".", font=BASE_FONT)
    style.configure("Title.TLabel", font=(BASE_FONT[0], BASE_FONT[1] + 5, "bold"))
    style.configure("Hint.TLabel", foreground="#667085")

    app = {"paths": [], "busy": False, "q": queue.Queue()}

    outer = ttk.Frame(root, padding=16)
    outer.pack(fill="both", expand=True)

    ttk.Label(outer, text=APP_TITLE, style="Title.TLabel").pack(anchor="w")
    ttk.Label(outer, text="把一份或多份考勤表加进来，统计每人在所有日期里出现的次数，"
                          "汇总成一张表。", style="Hint.TLabel").pack(anchor="w", pady=(2, 12))

    # ---- files ----
    files_box = ttk.LabelFrame(outer, text="考勤表文件", padding=10)
    files_box.pack(fill="both", expand=True)

    tree = ttk.Treeview(files_box, columns=("n", "f"), show="headings", height=5,
                        selectmode="extended")
    tree.heading("n", text="#")
    tree.heading("f", text="文件")
    tree.column("n", width=40, anchor="center", stretch=False)
    tree.column("f", width=430, anchor="w")
    vs = ttk.Scrollbar(files_box, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vs.set)
    tree.grid(row=0, column=0, sticky="nsew")
    vs.grid(row=0, column=1, sticky="ns")
    files_box.rowconfigure(0, weight=1)
    files_box.columnconfigure(0, weight=1)

    btns = ttk.Frame(files_box)
    btns.grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))

    def refresh_tree():
        tree.delete(*tree.get_children())
        for i, p in enumerate(app["paths"], 1):
            tree.insert("", "end", iid=str(i), values=(i, p))
        if not out_var.get().strip():
            out_var.set(default_output(app["paths"]))

    def add_files(paths):
        added = 0
        skipped = 0
        for p in paths:
            p = os.path.abspath(p)
            if p in app["paths"]:
                continue
            if not os.path.isfile(p):
                continue
            if not p.lower().endswith((".xlsx", ".xlsm")):
                skipped += 1
                continue
            app["paths"].append(p)
            added += 1
        refresh_tree()
        msg = []
        if added:
            msg.append("已加入 %d 个文件" % added)
        if skipped:
            msg.append("跳过 %d 个非 Excel 文件" % skipped)
        if msg:
            set_status("，".join(msg))

    def pick_files():
        chosen = filedialog.askopenfilenames(
            title="选择考勤表", filetypes=[("Excel 工作簿", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
        if chosen:
            add_files(chosen)

    def remove_selected():
        for iid in sorted((int(x) for x in tree.selection()), reverse=True):
            if 1 <= iid <= len(app["paths"]):
                app["paths"].pop(iid - 1)
        refresh_tree()

    def clear_all():
        app["paths"].clear()
        refresh_tree()

    ttk.Button(btns, text="添加文件...", command=pick_files).pack(side="left")
    ttk.Button(btns, text="移除选中", command=remove_selected).pack(side="left", padx=6)
    ttk.Button(btns, text="清空", command=clear_all).pack(side="left")
    ttk.Label(btns, text="（可多选；支持拖入文件）" if dnd_ok else "（可多选）",
              style="Hint.TLabel").pack(side="left", padx=10)

    if dnd_ok:
        def on_drop(event):
            add_files(root.tk.splitlist(event.data))
        tree.drop_target_register(DND_FILES)
        tree.dnd_bind("<<Drop>>", on_drop)

    # ---- options ----
    opt = ttk.LabelFrame(outer, text="选项", padding=10)
    opt.pack(fill="x", pady=10)
    opt.columnconfigure(1, weight=1)

    ttk.Label(opt, text="输出文件").grid(row=0, column=0, sticky="w", pady=3)
    out_var = tk.StringVar(value="")
    ttk.Entry(opt, textvariable=out_var).grid(row=0, column=1, sticky="ew", padx=8)

    def pick_out():
        p = filedialog.asksaveasfilename(
            title="保存统计结果", defaultextension=".xlsx",
            initialfile=os.path.basename(out_var.get() or DEFAULT_OUT_NAME),
            initialdir=os.path.dirname(out_var.get() or "") or os.getcwd(),
            filetypes=[("Excel 工作簿", "*.xlsx")])
        if p:
            out_var.set(p)

    ttk.Button(opt, text="浏览...", command=pick_out).grid(row=0, column=2)

    ttk.Label(opt, text="合并同名").grid(row=1, column=0, sticky="nw", pady=3)
    merge_txt = tk.Text(opt, height=2, width=1, font=BASE_FONT, relief="solid", borderwidth=1)
    merge_txt.grid(row=1, column=1, columnspan=2, sticky="ew", padx=8, pady=3)
    merge_txt.insert("1.0", "")
    ttk.Label(opt, text="一个人写法不一致时填这里，每行一条，例如：尹湉黙=尹湉默",
              style="Hint.TLabel").grid(row=2, column=1, columnspan=2, sticky="w", padx=8)

    verify_var = tk.BooleanVar(value=True)
    ttk.Checkbutton(opt, text="统计后自动校验（推荐）", variable=verify_var).grid(
        row=3, column=1, sticky="w", padx=8, pady=(6, 0))

    # ---- action ----
    act = ttk.Frame(outer)
    act.pack(fill="x")
    run_btn = ttk.Button(act, text="开始统计")
    run_btn.pack(side="left")
    prog = ttk.Progressbar(act, mode="indeterminate", length=220)
    prog.pack(side="left", padx=10)
    status = ttk.Label(act, text="就绪", style="Hint.TLabel")
    status.pack(side="left")

    # ---- results ----
    res = ttk.LabelFrame(outer, text="结果", padding=10)
    res.pack(fill="both", expand=True, pady=(10, 0))
    logbox = tk.Text(res, height=7, width=1, font=MONO, state="disabled", relief="solid",
                     borderwidth=1, background="#ffffff")
    ls = ttk.Scrollbar(res, orient="vertical", command=logbox.yview)
    logbox.configure(yscrollcommand=ls.set)
    logbox.pack(side="left", fill="both", expand=True)
    ls.pack(side="right", fill="y")

    open_frame = ttk.Frame(outer)
    open_frame.pack(fill="x", pady=(8, 0))
    open_btn = ttk.Button(open_frame, text="打开结果文件", state="disabled",
                          command=lambda: open_path(out_var.get()))
    open_btn.pack(side="left")
    reveal_btn = ttk.Button(open_frame, text="打开所在文件夹", state="disabled",
                            command=lambda: reveal_path(out_var.get()))
    reveal_btn.pack(side="left", padx=6)

    def log(line):
        app["q"].put(("log", line))

    def set_status(text):
        app["q"].put(("status", text))

    def drain():
        try:
            while True:
                kind, payload = app["q"].get_nowait()
                if kind == "log":
                    logbox.configure(state="normal")
                    logbox.insert("end", payload + "\n")
                    logbox.see("end")
                    logbox.configure(state="disabled")
                elif kind == "status":
                    status.configure(text=payload)
                elif kind == "done":
                    prog.stop()
                    run_btn.configure(state="normal")
                    app["busy"] = False
                    if payload:
                        open_btn.configure(state="normal")
                        reveal_btn.configure(state="normal")
                elif kind == "error":
                    messagebox.showerror(APP_TITLE, payload)
        except queue.Empty:
            pass
        root.after(100, drain)

    def worker(paths, out_path, merges, verify):
        try:
            log("开始统计，共 %d 个文件" % len(paths))
            run_tally(paths, out_path, merges, verify, log)
            app["q"].put(("status", "完成"))
            app["q"].put(("done", True))
        except Exception as exc:  # noqa: BLE001
            log("出错：" + str(exc))
            app["q"].put(("status", "失败"))
            app["q"].put(("done", False))
            app["q"].put(("error", str(exc)))

    def start():
        if app["busy"]:
            return
        paths = list(app["paths"])
        if not paths:
            messagebox.showwarning(APP_TITLE, "请先添加至少一个考勤表文件。")
            return
        out_path = out_var.get().strip() or default_output(paths)
        if not out_path.lower().endswith((".xlsx", ".xlsm")):
            out_path += ".xlsx"
        try:
            merges = parse_merges(merge_txt.get("1.0", "end"))
        except ValueError as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return
        if os.path.abspath(out_path) in paths:
            messagebox.showerror(APP_TITLE, "输出文件不能和输入文件是同一个。")
            return
        if os.path.exists(out_path) and not messagebox.askyesno(
                APP_TITLE, "文件已存在，覆盖吗？\n" + out_path):
            return
        app["busy"] = True
        run_btn.configure(state="disabled")
        open_btn.configure(state="disabled")
        reveal_btn.configure(state="disabled")
        set_status("统计中...")
        prog.start(12)
        threading.Thread(target=worker, args=(paths, out_path, merges, verify_var.get()),
                         daemon=True).start()

    run_btn.configure(command=start)

    # Size the window to what the content needs, without running off the screen.
    root.update_idletasks()
    w = min(root.winfo_reqwidth() + 20, root.winfo_screenwidth() - 60)
    h = min(root.winfo_reqheight() + 10, root.winfo_screenheight() - 140)
    root.geometry("%dx%d" % (max(700, w), max(560, h)))

    drain()
    return root, app


def selftest(argv):
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("inputs", nargs="+")
    ap.add_argument("-o", "--output", required=True)
    ap.add_argument("--merge", action="append")
    args = ap.parse_args(argv)
    merges = {}
    for it in args.merge or []:
        a, b = it.split("=", 1)
        merges[a.strip()] = b.strip()
    def safe_log(s):
        try:
            print("  " + s)
        except Exception:   # stdout can be detached in a windowed bundle
            pass
    run_tally(args.inputs, args.output, merges, True, safe_log)
    return 0


def main():
    if "--selftest" in sys.argv:
        return selftest(sys.argv[2:])
    try:
        build_app()[0].mainloop()
    except Exception:
        traceback.print_exc()
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
