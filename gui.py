#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件转换助手 · 向导式桌面界面
步骤：① 选功能 → ② 加文件 → ③ 排序 → ④ 输出设置 → ⑤ 开始
"""

from __future__ import annotations

import os
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Callable

from convert import (
    ALL_INPUT_EXTS,
    IMAGE_EXTS,
    OFFICE_TO_PDF,
    PDF_EXT,
    SVG_EXTS,
    collect_files_from_paths,
    find_soffice,
    run_pipeline,
    sort_paths,
)

# ---------------------------------------------------------------------------
# 功能定义：选一个功能后，自动收窄文件过滤器与默认合并行为
# ---------------------------------------------------------------------------

FEATURES: dict[str, dict] = {
    "img2pdf": {
        "title": "图片合成 PDF",
        "desc": "多张图片按顺序合成一份 PDF",
        "tip": "支持 png / jpg / webp / gif / bmp / tiff / ico / heic 等",
        "target": "pdf",
        "merge_default": True,
        "force_merge": True,
        "need_dpi": False,
        "out_name_default": "图片相册",
        "accept": IMAGE_EXTS | SVG_EXTS,
        "filetypes": [
            ("图片文件", " ".join(f"*{e}" for e in sorted(IMAGE_EXTS | SVG_EXTS))),
            ("所有文件", "*.*"),
        ],
        "need_soffice": False,
    },
    "pdf2img": {
        "title": "PDF 拆成图片",
        "desc": "把 PDF 每一页导出成图片",
        "tip": "可先合并多个 PDF，再按页导出为 PNG / JPG / WEBP",
        "target": "png",  # 可在步骤4改
        "merge_default": True,
        "force_merge": False,
        "need_dpi": True,
        "out_name_default": "页面",
        "accept": PDF_EXT,
        "filetypes": [("PDF 文件", "*.pdf"), ("所有文件", "*.*")],
        "need_soffice": False,
        "image_targets": True,
    },
    "office2pdf": {
        "title": "Word/PPT 转 PDF",
        "desc": "把 Word、PPT 文档转成 PDF",
        "tip": "需要本机安装 LibreOffice；可批量并合并为一个 PDF",
        "target": "pdf",
        "merge_default": False,
        "force_merge": False,
        "need_dpi": False,
        "out_name_default": "合并文档",
        "accept": OFFICE_TO_PDF,
        "filetypes": [
            ("Word / PPT", "*.doc *.docx *.ppt *.pptx *.odt *.odp"),
            ("所有文件", "*.*"),
        ],
        "need_soffice": True,
    },
    "pdf2word": {
        "title": "PDF 转 Word",
        "desc": "把 PDF 转成可编辑的 Word（.docx）",
        "tip": "尽量保留文字与表格；复杂版式可能有偏差",
        "target": "docx",
        "merge_default": False,
        "force_merge": False,
        "need_dpi": False,
        "out_name_default": "合并文档",
        "accept": PDF_EXT,
        "filetypes": [("PDF 文件", "*.pdf"), ("所有文件", "*.*")],
        "need_soffice": False,
    },
    "pdf2ppt": {
        "title": "PDF 转 PPT",
        "desc": "把 PDF 转成 PPT（每页一张图）",
        "tip": "适合放映存档；幻灯片内文字不可直接编辑",
        "target": "pptx",
        "merge_default": False,
        "force_merge": False,
        "need_dpi": True,
        "out_name_default": "合并演示",
        "accept": PDF_EXT,
        "filetypes": [("PDF 文件", "*.pdf"), ("所有文件", "*.*")],
        "need_soffice": False,
    },
    "merge_pdf": {
        "title": "合并多个 PDF",
        "desc": "按顺序把多个 PDF 合成一份",
        "tip": "仅合并，不改变内容；请在下一步调整顺序",
        "target": "pdf",
        "merge_default": True,
        "force_merge": True,
        "need_dpi": False,
        "out_name_default": "合并结果",
        "accept": PDF_EXT,
        "filetypes": [("PDF 文件", "*.pdf"), ("所有文件", "*.*")],
        "need_soffice": False,
    },
}

FEATURE_ORDER = [
    "img2pdf",
    "pdf2img",
    "office2pdf",
    "pdf2word",
    "pdf2ppt",
    "merge_pdf",
]

STEP_LABELS = [
    "① 选功能",
    "② 加文件",
    "③ 排序",
    "④ 输出设置",
    "⑤ 开始",
]


def open_folder(path: Path) -> None:
    path = Path(path)
    if not path.exists():
        return
    folder = path if path.is_dir() else path.parent
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(folder))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(folder)])
        else:
            subprocess.Popen(["xdg-open", str(folder)])
    except Exception:
        pass


class WizardApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("文件转换助手")
        self.geometry("820x640")
        self.minsize(720, 560)
        self.configure(bg="#f4f6f8")

        self.step = 0  # 0..4
        self.feature_id: str | None = None
        self.files: list[Path] = []

        self.merge_var = tk.BooleanVar(value=True)
        self.out_dir = tk.StringVar()
        self.out_name = tk.StringVar(value="合并结果")
        self.dpi_var = tk.StringVar(value="150")
        self.img_fmt = tk.StringVar(value="png")
        self.status = tk.StringVar(value="")
        self.progress_val = tk.DoubleVar(value=0.0)
        self.progress_msg = tk.StringVar(value="")

        self._feature_btns: dict[str, tk.Button] = {}
        self._step_labels: list[tk.Label] = []
        self._page_frames: list[ttk.Frame] = []

        self._setup_style()
        self._build_shell()
        self._build_pages()
        self._show_step(0)
        self._update_soffice_hint()

    # ------------------------------------------------------------------ style
    def _setup_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background="#f4f6f8")
        style.configure("Card.TFrame", background="#ffffff")
        style.configure("TLabel", background="#f4f6f8", font=("Microsoft YaHei UI", 10))
        style.configure("Title.TLabel", background="#f4f6f8", font=("Microsoft YaHei UI", 16, "bold"))
        style.configure("Hint.TLabel", background="#f4f6f8", foreground="#5a6570", font=("Microsoft YaHei UI", 9))
        style.configure("StepOn.TLabel", background="#1a73e8", foreground="#ffffff",
                        font=("Microsoft YaHei UI", 9, "bold"), padding=(10, 6))
        style.configure("StepOff.TLabel", background="#dfe3e8", foreground="#445",
                        font=("Microsoft YaHei UI", 9), padding=(10, 6))
        style.configure("StepDone.TLabel", background="#c8e6c9", foreground="#1b5e20",
                        font=("Microsoft YaHei UI", 9), padding=(10, 6))
        style.configure("CardTitle.TLabel", background="#ffffff", font=("Microsoft YaHei UI", 12, "bold"))
        style.configure("CardDesc.TLabel", background="#ffffff", foreground="#555", font=("Microsoft YaHei UI", 9))
        style.configure("Primary.TButton", font=("Microsoft YaHei UI", 11, "bold"), padding=(16, 8))
        style.configure("Nav.TButton", font=("Microsoft YaHei UI", 10), padding=(12, 6))
        style.configure("TLabelframe", background="#f4f6f8")
        style.configure("TLabelframe.Label", background="#f4f6f8", font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("TCheckbutton", background="#f4f6f8", font=("Microsoft YaHei UI", 10))
        style.configure("TRadiobutton", background="#f4f6f8", font=("Microsoft YaHei UI", 10))
        style.configure("Horizontal.TProgressbar", thickness=16)

    # ------------------------------------------------------------------ shell
    def _build_shell(self) -> None:
        # 顶栏
        header = ttk.Frame(self, padding=(16, 12, 16, 4))
        header.pack(fill=tk.X)
        ttk.Label(header, text="文件转换助手", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="按步骤操作即可：选功能 → 加文件 → 排序 → 设输出 → 开始转换。全程在本机完成，不会上传。",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(2, 0))

        # 步骤条
        step_bar = ttk.Frame(self, padding=(16, 8))
        step_bar.pack(fill=tk.X)
        for i, text in enumerate(STEP_LABELS):
            lbl = ttk.Label(step_bar, text=text, style="StepOff.TLabel", anchor="center")
            lbl.pack(side=tk.LEFT, padx=3, fill=tk.X, expand=True)
            self._step_labels.append(lbl)

        # 内容区
        self.body = ttk.Frame(self, padding=(16, 4, 16, 4))
        self.body.pack(fill=tk.BOTH, expand=True)

        # 底栏导航
        nav = ttk.Frame(self, padding=(16, 8, 16, 14))
        nav.pack(fill=tk.X, side=tk.BOTTOM)
        self.btn_back = ttk.Button(nav, text="← 上一步", style="Nav.TButton", command=self._go_back)
        self.btn_back.pack(side=tk.LEFT)
        self.btn_next = ttk.Button(nav, text="下一步 →", style="Primary.TButton", command=self._go_next)
        self.btn_next.pack(side=tk.RIGHT)
        ttk.Label(nav, textvariable=self.status, style="Hint.TLabel").pack(side=tk.LEFT, padx=16)

    def _build_pages(self) -> None:
        self._page_frames = [
            self._page_feature(),
            self._page_files(),
            self._page_sort(),
            self._page_output(),
            self._page_run(),
        ]
        for fr in self._page_frames:
            fr.place(in_=self.body, relx=0, rely=0, relwidth=1, relheight=1)

    # -------------------------------------------------------------- pages
    def _page_feature(self) -> ttk.Frame:
        fr = ttk.Frame(self.body)
        ttk.Label(fr, text="请选择你要做的事", font=("Microsoft YaHei UI", 13, "bold")).pack(
            anchor="w", pady=(4, 2)
        )
        ttk.Label(
            fr,
            text="点一下下方卡片即可。选好后点右下角「下一步」。",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(0, 10))

        grid = ttk.Frame(fr)
        grid.pack(fill=tk.BOTH, expand=True)

        for i, fid in enumerate(FEATURE_ORDER):
            meta = FEATURES[fid]
            row, col = divmod(i, 2)
            card = tk.Frame(
                grid,
                bg="#ffffff",
                highlightbackground="#cfd6dd",
                highlightthickness=1,
                padx=14,
                pady=12,
                cursor="hand2",
            )
            card.grid(row=row, column=col, sticky="nsew", padx=6, pady=6)
            grid.rowconfigure(row, weight=1)
            grid.columnconfigure(col, weight=1)

            title = tk.Label(
                card, text=meta["title"], bg="#ffffff",
                font=("Microsoft YaHei UI", 12, "bold"), anchor="w", cursor="hand2",
            )
            title.pack(fill=tk.X)
            desc = tk.Label(
                card, text=meta["desc"], bg="#ffffff", fg="#555555",
                font=("Microsoft YaHei UI", 9), anchor="w", justify=tk.LEFT, wraplength=300,
                cursor="hand2",
            )
            desc.pack(fill=tk.X, pady=(4, 0))
            tip = tk.Label(
                card, text=meta["tip"], bg="#ffffff", fg="#888888",
                font=("Microsoft YaHei UI", 8), anchor="w", justify=tk.LEFT, wraplength=300,
                cursor="hand2",
            )
            tip.pack(fill=tk.X, pady=(6, 0))

            def make_handler(feature_id: str) -> Callable[[], None]:
                return lambda: self._select_feature(feature_id)

            handler = make_handler(fid)
            for w in (card, title, desc, tip):
                w.bind("<Button-1>", lambda _e, h=handler: h())

            # 用一个隐藏按钮存引用，方便高亮
            self._feature_btns[fid] = card  # type: ignore[assignment]

        self.soffice_hint = ttk.Label(fr, text="", style="Hint.TLabel")
        self.soffice_hint.pack(anchor="w", pady=(8, 0))
        return fr

    def _page_files(self) -> ttk.Frame:
        fr = ttk.Frame(self.body)
        self.files_title = ttk.Label(fr, text="添加要转换的文件", font=("Microsoft YaHei UI", 13, "bold"))
        self.files_title.pack(anchor="w", pady=(4, 2))
        self.files_tip = ttk.Label(fr, text="", style="Hint.TLabel")
        self.files_tip.pack(anchor="w", pady=(0, 8))

        btn_row = ttk.Frame(fr)
        btn_row.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(btn_row, text="＋ 添加文件…", style="Nav.TButton", command=self._add_files).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(btn_row, text="＋ 添加文件夹…", style="Nav.TButton", command=self._add_folder).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(btn_row, text="删除选中", command=self._remove_selected).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="清空列表", command=self._clear_all).pack(side=tk.LEFT)

        list_frame = ttk.LabelFrame(fr, text="已选文件", padding=8)
        list_frame.pack(fill=tk.BOTH, expand=True)

        cols = ("idx", "name", "type")
        self.tree_files = ttk.Treeview(
            list_frame, columns=cols, show="headings", selectmode="extended", height=12
        )
        self.tree_files.heading("idx", text="#")
        self.tree_files.heading("name", text="文件名")
        self.tree_files.heading("type", text="类型")
        self.tree_files.column("idx", width=48, anchor="center", stretch=False)
        self.tree_files.column("name", width=520, anchor="w")
        self.tree_files.column("type", width=80, anchor="center", stretch=False)
        scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree_files.yview)
        self.tree_files.configure(yscrollcommand=scroll.set)
        self.tree_files.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.files_count = ttk.Label(fr, text="当前 0 个文件", style="Hint.TLabel")
        self.files_count.pack(anchor="w", pady=(6, 0))
        return fr

    def _page_sort(self) -> ttk.Frame:
        fr = ttk.Frame(self.body)
        ttk.Label(fr, text="调整文件顺序", font=("Microsoft YaHei UI", 13, "bold")).pack(
            anchor="w", pady=(4, 2)
        )
        ttk.Label(
            fr,
            text="合并或合成时，列表从上到下就是最终顺序。可一键按文件名排序，或手动上移/下移。",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(0, 8))

        btn_row = ttk.Frame(fr)
        btn_row.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(btn_row, text="按文件名排序（推荐）", style="Nav.TButton", command=self._sort_by_name).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(btn_row, text="上移 ↑", command=self._move_up).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="下移 ↓", command=self._move_down).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="移到最上", command=self._move_top).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="移到最下", command=self._move_bottom).pack(side=tk.LEFT)

        list_frame = ttk.LabelFrame(fr, text="当前顺序", padding=8)
        list_frame.pack(fill=tk.BOTH, expand=True)

        cols = ("idx", "name", "type")
        self.tree_sort = ttk.Treeview(
            list_frame, columns=cols, show="headings", selectmode="browse", height=12
        )
        self.tree_sort.heading("idx", text="#")
        self.tree_sort.heading("name", text="文件名")
        self.tree_sort.heading("type", text="类型")
        self.tree_sort.column("idx", width=48, anchor="center", stretch=False)
        self.tree_sort.column("name", width=520, anchor="w")
        self.tree_sort.column("type", width=80, anchor="center", stretch=False)
        scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree_sort.yview)
        self.tree_sort.configure(yscrollcommand=scroll.set)
        self.tree_sort.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        return fr

    def _page_output(self) -> ttk.Frame:
        fr = ttk.Frame(self.body)
        ttk.Label(fr, text="设置输出结果", font=("Microsoft YaHei UI", 13, "bold")).pack(
            anchor="w", pady=(4, 2)
        )
        self.output_tip = ttk.Label(fr, text="", style="Hint.TLabel")
        self.output_tip.pack(anchor="w", pady=(0, 10))

        # 用普通 Frame 做白底卡片
        card_bg = tk.Frame(fr, bg="#ffffff", highlightbackground="#cfd6dd", highlightthickness=1, padx=18, pady=16)
        card_bg.pack(fill=tk.X)

        # 图片格式（仅 pdf2img）
        self.img_fmt_row = tk.Frame(card_bg, bg="#ffffff")
        tk.Label(
            self.img_fmt_row, text="导出图片格式：", bg="#ffffff",
            font=("Microsoft YaHei UI", 10),
        ).pack(side=tk.LEFT)
        for label, val in (("PNG", "png"), ("JPG", "jpg"), ("WEBP", "webp")):
            ttk.Radiobutton(
                self.img_fmt_row, text=label, value=val, variable=self.img_fmt,
            ).pack(side=tk.LEFT, padx=6)

        # 合并
        self.merge_row = tk.Frame(card_bg, bg="#ffffff")
        self.merge_row.pack(fill=tk.X, pady=4)
        self.merge_chk = ttk.Checkbutton(
            self.merge_row,
            text="合并后再转换（多文件合成一个结果）",
            variable=self.merge_var,
        )
        self.merge_chk.pack(side=tk.LEFT)

        # DPI
        self.dpi_row = tk.Frame(card_bg, bg="#ffffff")
        tk.Label(
            self.dpi_row, text="清晰度 DPI：", bg="#ffffff", font=("Microsoft YaHei UI", 10),
        ).pack(side=tk.LEFT)
        ttk.Entry(self.dpi_row, textvariable=self.dpi_var, width=8).pack(side=tk.LEFT, padx=4)
        tk.Label(
            self.dpi_row, text="（越大越清晰，文件也越大；常用 150～200）",
            bg="#ffffff", fg="#888", font=("Microsoft YaHei UI", 8),
        ).pack(side=tk.LEFT)

        # 输出目录
        row_dir = tk.Frame(card_bg, bg="#ffffff")
        row_dir.pack(fill=tk.X, pady=8)
        tk.Label(row_dir, text="保存到文件夹：", bg="#ffffff", font=("Microsoft YaHei UI", 10)).pack(
            side=tk.LEFT
        )
        ttk.Entry(row_dir, textvariable=self.out_dir).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        ttk.Button(row_dir, text="浏览…", command=self._browse_out_dir).pack(side=tk.LEFT)

        # 合并文件名
        self.name_row = tk.Frame(card_bg, bg="#ffffff")
        self.name_row.pack(fill=tk.X, pady=4)
        tk.Label(self.name_row, text="结果文件名：", bg="#ffffff", font=("Microsoft YaHei UI", 10)).pack(
            side=tk.LEFT
        )
        ttk.Entry(self.name_row, textvariable=self.out_name, width=28).pack(side=tk.LEFT, padx=6)
        tk.Label(
            self.name_row, text="（扩展名会自动补上）", bg="#ffffff", fg="#888",
            font=("Microsoft YaHei UI", 8),
        ).pack(side=tk.LEFT)

        summary = ttk.LabelFrame(fr, text="即将执行", padding=10)
        summary.pack(fill=tk.X, pady=(14, 0))
        self.summary_text = ttk.Label(summary, text="", wraplength=720, justify=tk.LEFT)
        self.summary_text.pack(anchor="w")
        return fr

    def _page_run(self) -> ttk.Frame:
        fr = ttk.Frame(self.body)

        # 先 pack 底部操作区，保证「开始转换」始终可见，不被长文件列表挤出屏幕
        action = ttk.Frame(fr)
        action.pack(side=tk.BOTTOM, fill=tk.X, pady=(8, 0))

        self.start_btn = tk.Button(
            action,
            text="▶  开始转换",
            font=("Microsoft YaHei UI", 14, "bold"),
            bg="#1a73e8",
            fg="#ffffff",
            activebackground="#1557b0",
            activeforeground="#ffffff",
            relief=tk.FLAT,
            padx=28,
            pady=12,
            cursor="hand2",
            command=self._start_convert,
        )
        self.start_btn.pack(pady=(0, 6))

        self.progress = ttk.Progressbar(
            action, mode="indeterminate", variable=self.progress_val,
            style="Horizontal.TProgressbar",
        )
        self.progress.pack(fill=tk.X, pady=(4, 4))
        ttk.Label(action, textvariable=self.progress_msg, style="Hint.TLabel").pack(anchor="w")

        ttk.Label(fr, text="开始转换", font=("Microsoft YaHei UI", 13, "bold")).pack(
            anchor="w", pady=(4, 2)
        )
        ttk.Label(
            fr,
            text="请核对下方信息。确认无误后，点击本页底部的蓝色「开始转换」按钮。",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(0, 8))

        sum_frame = ttk.LabelFrame(fr, text="确认信息与文件顺序（可滚动）", padding=6)
        sum_frame.pack(fill=tk.BOTH, expand=True)

        self.run_summary = tk.Text(
            sum_frame,
            wrap=tk.WORD,
            font=("Microsoft YaHei UI", 10),
            bg="#ffffff",
            fg="#333333",
            relief=tk.FLAT,
            padx=10,
            pady=8,
            height=12,
            cursor="arrow",
        )
        scroll = ttk.Scrollbar(sum_frame, orient=tk.VERTICAL, command=self.run_summary.yview)
        self.run_summary.configure(yscrollcommand=scroll.set)
        self.run_summary.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.run_summary.bind("<Key>", lambda e: "break")  # 只读
        return fr

    # -------------------------------------------------------------- navigation
    def _show_step(self, step: int) -> None:
        self.step = step
        for i, fr in enumerate(self._page_frames):
            if i == step:
                fr.lift()
            # always placed; lift brings to front

        for i, lbl in enumerate(self._step_labels):
            if i < step:
                lbl.configure(style="StepDone.TLabel")
            elif i == step:
                lbl.configure(style="StepOn.TLabel")
            else:
                lbl.configure(style="StepOff.TLabel")

        self.btn_back.configure(state=tk.NORMAL if step > 0 else tk.DISABLED)

        if step == 4:
            self.btn_next.configure(text="完成并关闭", command=self.destroy)
        else:
            self.btn_next.configure(text="下一步 →", command=self._go_next)

        if step == 1:
            self._refresh_files_page()
        elif step == 2:
            self._refresh_tree(self.tree_sort)
        elif step == 3:
            self._refresh_output_page()
        elif step == 4:
            self._refresh_run_page()

    def _go_back(self) -> None:
        if self.step > 0:
            self._show_step(self.step - 1)

    def _go_next(self) -> None:
        if self.step == 0:
            if not self.feature_id:
                messagebox.showinfo("提示", "请先点选一个功能卡片")
                return
            meta = FEATURES[self.feature_id]
            if meta.get("need_soffice") and not find_soffice():
                if not messagebox.askyesno(
                    "未检测到 LibreOffice",
                    "Word/PPT 转 PDF 需要安装 LibreOffice。\n\n"
                    "仍要继续吗？（可先添加文件，稍后再装）\n\n"
                    "下载：https://www.libreoffice.org/download/",
                ):
                    return
            # 进入加文件时清空不匹配的旧文件
            accept = meta["accept"]
            self.files = [f for f in self.files if f.suffix.lower() in accept]
            self.merge_var.set(meta["merge_default"])
            self.out_name.set(meta["out_name_default"])
            self._show_step(1)
            return

        if self.step == 1:
            if not self.files:
                messagebox.showinfo("提示", "请先添加至少一个文件")
                return
            self._show_step(2)
            return

        if self.step == 2:
            self._show_step(3)
            return

        if self.step == 3:
            out = self.out_dir.get().strip()
            if not out:
                # 默认用第一个文件目录
                self.out_dir.set(str(self.files[0].parent))
            try:
                dpi = int(self.dpi_var.get().strip() or "150")
                if dpi < 36 or dpi > 600:
                    raise ValueError
            except ValueError:
                messagebox.showerror("提示", "DPI 请填写 36～600 之间的整数")
                return
            self._show_step(4)
            return

    # -------------------------------------------------------------- feature
    def _select_feature(self, fid: str) -> None:
        self.feature_id = fid
        for key, card in self._feature_btns.items():
            if key == fid:
                card.configure(highlightbackground="#1a73e8", highlightthickness=3, bg="#e8f0fe")
                for child in card.winfo_children():
                    try:
                        child.configure(bg="#e8f0fe")
                    except tk.TclError:
                        pass
            else:
                card.configure(highlightbackground="#cfd6dd", highlightthickness=1, bg="#ffffff")
                for child in card.winfo_children():
                    try:
                        child.configure(bg="#ffffff")
                    except tk.TclError:
                        pass
        self.status.set(f"已选择：{FEATURES[fid]['title']}")

    def _update_soffice_hint(self) -> None:
        if find_soffice():
            self.soffice_hint.configure(text="✓ 已检测到 LibreOffice（Word/PPT 转 PDF 可用）")
        else:
            self.soffice_hint.configure(
                text="○ 未检测到 LibreOffice：若要用「Word/PPT 转 PDF」，请先安装 LibreOffice"
            )

    # -------------------------------------------------------------- files
    def _refresh_files_page(self) -> None:
        meta = FEATURES[self.feature_id or "img2pdf"]
        self.files_title.configure(text=f"添加文件 · {meta['title']}")
        self.files_tip.configure(
            text=f"{meta['tip']}　可多选文件，也可直接添加整个文件夹。"
        )
        self._refresh_tree(self.tree_files)
        self.files_count.configure(text=f"当前 {len(self.files)} 个文件")

    def _refresh_tree(self, tree: ttk.Treeview) -> None:
        tree.delete(*tree.get_children())
        for i, p in enumerate(self.files, 1):
            tree.insert("", tk.END, iid=str(i - 1), values=(i, p.name, p.suffix.lower() or "—"))

    def _current_filetypes(self):
        meta = FEATURES[self.feature_id or "img2pdf"]
        return meta["filetypes"]

    def _filter_accepted(self, paths: list[Path]) -> list[Path]:
        meta = FEATURES[self.feature_id or "img2pdf"]
        accept = meta["accept"]
        ok = [p for p in paths if p.suffix.lower() in accept]
        skipped = len(paths) - len(ok)
        if skipped:
            messagebox.showinfo(
                "已跳过不匹配的文件",
                f"当前功能只接受特定格式，已跳过 {skipped} 个不匹配的文件。",
            )
        return ok

    def _add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="选择文件（可多选）",
            filetypes=self._current_filetypes(),
        )
        if not paths:
            return
        added = collect_files_from_paths(Path(p) for p in paths)
        added = self._filter_accepted(added)
        existing = set(self.files)
        n = 0
        for f in added:
            if f not in existing:
                self.files.append(f)
                existing.add(f)
                n += 1
        self._refresh_tree(self.tree_files)
        self.files_count.configure(text=f"当前 {len(self.files)} 个文件")
        self.status.set(f"新增 {n} 个，共 {len(self.files)} 个")

    def _add_folder(self) -> None:
        folder = filedialog.askdirectory(title="选择文件夹（自动收集匹配的文件）")
        if not folder:
            return
        added = collect_files_from_paths([Path(folder)])
        added = self._filter_accepted(added)
        existing = set(self.files)
        n = 0
        for f in added:
            if f not in existing:
                self.files.append(f)
                existing.add(f)
                n += 1
        self._refresh_tree(self.tree_files)
        self.files_count.configure(text=f"当前 {len(self.files)} 个文件")
        self.status.set(f"从文件夹新增 {n} 个，共 {len(self.files)} 个")

    def _selected_indices(self, tree: ttk.Treeview) -> list[int]:
        return sorted(int(i) for i in tree.selection())

    def _remove_selected(self) -> None:
        tree = self.tree_files if self.step == 1 else self.tree_sort
        idxs = self._selected_indices(tree)
        if not idxs:
            return
        for i in reversed(idxs):
            if 0 <= i < len(self.files):
                self.files.pop(i)
        self._refresh_tree(tree)
        if hasattr(self, "files_count"):
            self.files_count.configure(text=f"当前 {len(self.files)} 个文件")
        self.status.set(f"已删除，剩余 {len(self.files)} 个")

    def _clear_all(self) -> None:
        self.files.clear()
        self._refresh_tree(self.tree_files)
        self.files_count.configure(text="当前 0 个文件")
        self.status.set("列表已清空")

    # -------------------------------------------------------------- sort
    def _selected_index(self) -> int | None:
        sel = self.tree_sort.selection()
        if not sel:
            return None
        return int(sel[0])

    def _sort_by_name(self) -> None:
        if not self.files:
            return
        self.files = sort_paths(self.files, natural=True)
        self._refresh_tree(self.tree_sort)
        self.status.set("已按文件名自然排序（如 img2 在 img10 前）")

    def _move_up(self) -> None:
        idx = self._selected_index()
        if idx is None or idx <= 0:
            return
        self.files[idx - 1], self.files[idx] = self.files[idx], self.files[idx - 1]
        self._refresh_tree(self.tree_sort)
        self.tree_sort.selection_set(str(idx - 1))
        self.tree_sort.see(str(idx - 1))

    def _move_down(self) -> None:
        idx = self._selected_index()
        if idx is None or idx >= len(self.files) - 1:
            return
        self.files[idx + 1], self.files[idx] = self.files[idx], self.files[idx + 1]
        self._refresh_tree(self.tree_sort)
        self.tree_sort.selection_set(str(idx + 1))
        self.tree_sort.see(str(idx + 1))

    def _move_top(self) -> None:
        idx = self._selected_index()
        if idx is None or idx == 0:
            return
        item = self.files.pop(idx)
        self.files.insert(0, item)
        self._refresh_tree(self.tree_sort)
        self.tree_sort.selection_set("0")
        self.tree_sort.see("0")

    def _move_bottom(self) -> None:
        idx = self._selected_index()
        if idx is None or idx >= len(self.files) - 1:
            return
        item = self.files.pop(idx)
        self.files.append(item)
        self._refresh_tree(self.tree_sort)
        last = str(len(self.files) - 1)
        self.tree_sort.selection_set(last)
        self.tree_sort.see(last)

    # -------------------------------------------------------------- output
    def _browse_out_dir(self) -> None:
        path = filedialog.askdirectory(title="选择保存结果的文件夹")
        if path:
            self.out_dir.set(path)
            self._update_summary()

    def _refresh_output_page(self) -> None:
        meta = FEATURES[self.feature_id or "img2pdf"]
        self.output_tip.configure(text=f"功能：{meta['title']}　·　{meta['desc']}")

        # 图片格式行
        if meta.get("image_targets"):
            if not self.img_fmt_row.winfo_ismapped():
                self.img_fmt_row.pack(fill=tk.X, pady=4, before=self.merge_row)
        else:
            self.img_fmt_row.pack_forget()

        # 合并
        if meta.get("force_merge"):
            self.merge_var.set(True)
            self.merge_chk.configure(state=tk.DISABLED)
            self.merge_chk.configure(text="合并后再转换（本功能固定合并）")
        else:
            self.merge_chk.configure(state=tk.NORMAL)
            self.merge_chk.configure(text="合并后再转换（多文件合成一个结果）")

        # DPI
        if meta.get("need_dpi"):
            if not self.dpi_row.winfo_ismapped():
                self.dpi_row.pack(fill=tk.X, pady=4)
        else:
            self.dpi_row.pack_forget()

        if not self.out_dir.get().strip() and self.files:
            self.out_dir.set(str(self.files[0].parent))

        if not getattr(self, "_output_traces_bound", False):
            self.merge_var.trace_add("write", lambda *_: self._update_summary())
            self.img_fmt.trace_add("write", lambda *_: self._update_summary())
            self.out_name.trace_add("write", lambda *_: self._update_summary())
            self.out_dir.trace_add("write", lambda *_: self._update_summary())
            self._output_traces_bound = True

        self._update_summary()

    def _resolve_target(self) -> str:
        meta = FEATURES[self.feature_id or "img2pdf"]
        if meta.get("image_targets"):
            return self.img_fmt.get()
        return meta["target"]

    def _update_summary(self) -> None:
        if not self.feature_id:
            return
        meta = FEATURES[self.feature_id]
        target = self._resolve_target()
        merge = self.merge_var.get()
        name = self.out_name.get().strip() or meta["out_name_default"]
        out = self.out_dir.get().strip() or "（未选，将使用文件所在目录）"
        lines = [
            f"功能：{meta['title']}",
            f"文件数：{len(self.files)} 个",
            f"目标格式：.{target}",
            f"是否合并：{'是' if merge else '否（逐个转换）'}",
            f"结果文件名：{name}",
            f"保存位置：{out}",
        ]
        if hasattr(self, "summary_text"):
            self.summary_text.configure(text="\n".join(lines))

    def _refresh_run_page(self) -> None:
        meta = FEATURES[self.feature_id or "img2pdf"]
        target = self._resolve_target()
        merge = self.merge_var.get()
        name = self.out_name.get().strip() or meta["out_name_default"]
        out = self.out_dir.get().strip()
        lines = [
            f"功能：{meta['title']}",
            f"目标：.{target}　　合并：{'是' if merge else '否'}　　文件名：{name}",
            f"保存到：{out}",
            "",
            f"文件顺序（共 {len(self.files)} 个，可滚动查看）：",
        ]
        for i, p in enumerate(self.files, 1):
            lines.append(f"  {i}. {p.name}")
        body = "\n".join(lines)
        self.run_summary.configure(state=tk.NORMAL)
        self.run_summary.delete("1.0", tk.END)
        self.run_summary.insert("1.0", body)
        self.run_summary.configure(state=tk.DISABLED)
        self.progress_msg.set("准备就绪 — 请点击蓝色「开始转换」")
        self.progress.stop()
        self.start_btn.configure(state=tk.NORMAL, bg="#1a73e8")

    # -------------------------------------------------------------- convert
    def _start_convert(self) -> None:
        if not self.files or not self.feature_id:
            messagebox.showwarning("提示", "没有可转换的文件")
            return
        out = self.out_dir.get().strip()
        if not out:
            out = str(self.files[0].parent)
            self.out_dir.set(out)
        out_path = Path(out)
        try:
            out_path.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            messagebox.showerror("错误", f"无法创建输出目录：\n{e}")
            return

        try:
            dpi = int(self.dpi_var.get().strip() or "150")
        except ValueError:
            messagebox.showerror("错误", "DPI 必须是整数")
            return

        target = self._resolve_target()
        merge = bool(self.merge_var.get())
        name = self.out_name.get().strip() or FEATURES[self.feature_id]["out_name_default"]
        files = list(self.files)

        self.start_btn.configure(state=tk.DISABLED, bg="#9aa0a6")
        self.btn_back.configure(state=tk.DISABLED)
        self.btn_next.configure(state=tk.DISABLED)
        self.progress_msg.set("正在转换，请稍候…")
        self.progress.start(12)

        def worker() -> None:
            try:
                def prog(msg: str) -> None:
                    self.after(0, lambda m=msg: self.progress_msg.set(m))

                results = run_pipeline(
                    files,
                    target,
                    out_path,
                    merge=merge,
                    output_name=name,
                    dpi=dpi,
                    progress=prog,
                )
                self.after(0, lambda: self._on_ok(results, out_path))
            except Exception as e:
                self.after(0, lambda: self._on_err(e))

        threading.Thread(target=worker, daemon=True).start()

    def _on_ok(self, results: list[Path], out_path: Path) -> None:
        self.progress.stop()
        self.start_btn.configure(state=tk.NORMAL, bg="#1a73e8")
        self.btn_back.configure(state=tk.NORMAL)
        self.btn_next.configure(state=tk.NORMAL)
        if not results:
            self.progress_msg.set("完成，但没有输出文件")
            messagebox.showwarning("完成", "没有生成输出文件")
            return
        self.progress_msg.set(f"✓ 完成，共 {len(results)} 个输出")
        lines = "\n".join(str(r) for r in results[:15])
        extra = "" if len(results) <= 15 else f"\n… 另有 {len(results) - 15} 个"
        # 自定义成功对话框：带「打开输出文件夹」
        dlg = tk.Toplevel(self)
        dlg.title("转换成功")
        dlg.transient(self)
        dlg.grab_set()
        dlg.resizable(False, False)
        dlg.configure(bg="#ffffff")
        tk.Label(
            dlg, text="✓ 转换成功", bg="#ffffff", fg="#1b5e20",
            font=("Microsoft YaHei UI", 14, "bold"),
        ).pack(anchor="w", padx=20, pady=(16, 6))
        tk.Label(
            dlg,
            text=f"已保存 {len(results)} 个文件：\n{lines}{extra}",
            bg="#ffffff", fg="#333", justify=tk.LEFT,
            font=("Microsoft YaHei UI", 9),
        ).pack(anchor="w", padx=20, pady=4)
        btn_row = tk.Frame(dlg, bg="#ffffff")
        btn_row.pack(fill=tk.X, padx=20, pady=(12, 16))
        tk.Button(
            btn_row, text="打开输出文件夹",
            font=("Microsoft YaHei UI", 10, "bold"),
            bg="#1a73e8", fg="#fff", relief=tk.FLAT, padx=12, pady=6,
            command=lambda: (open_folder(out_path), dlg.destroy()),
        ).pack(side=tk.LEFT)
        tk.Button(
            btn_row, text="关闭",
            font=("Microsoft YaHei UI", 10),
            relief=tk.GROOVE, padx=12, pady=6,
            command=dlg.destroy,
        ).pack(side=tk.RIGHT)
        dlg.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - dlg.winfo_width()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - dlg.winfo_height()) // 2
        dlg.geometry(f"+{x}+{y}")

    def _on_err(self, err: Exception) -> None:
        self.progress.stop()
        self.start_btn.configure(state=tk.NORMAL, bg="#1a73e8")
        self.btn_back.configure(state=tk.NORMAL)
        self.btn_next.configure(state=tk.NORMAL)
        self.progress_msg.set("转换失败")
        messagebox.showerror("转换失败", str(err))


def main() -> None:
    app = WizardApp()
    app.mainloop()


if __name__ == "__main__":
    main()
