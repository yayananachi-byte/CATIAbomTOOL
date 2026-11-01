# -*- coding: utf-8 -*-
"""CATIA BOM Manager - Windows GUI entry point.

功能：
* 导出当前 CATIA 总成，或选择多个 CATProduct 串行导出；
* 将已导出的 BOM 转写到客户基础数据管制表模板；
* 显示实时日志、进度与安全取消状态；
* 所有 CATIA 操作在后台 STA 线程中执行，界面不会冻结。
"""

from __future__ import annotations

import importlib.util
import os
import queue
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


IS_FROZEN = bool(getattr(sys, "frozen", False))
if IS_FROZEN:
    APP_DIR = Path(sys.executable).resolve().parent
    PROJECT_DIR = Path(getattr(sys, "_MEIPASS", APP_DIR)).resolve()
    USER_DATA_DIR = Path.home() / "Documents" / "YR CATIA BOM Manager"
else:
    APP_DIR = Path(__file__).resolve().parent
    PROJECT_DIR = APP_DIR.parent
    USER_DATA_DIR = PROJECT_DIR
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

APP_NAME = "YR CATIA BOM Manager"
DEFAULT_EXPORT_DIR = USER_DATA_DIR / "CATIA_BOM导出" / "GUI导出"
CONVERTER_PATH = PROJECT_DIR / "模板转化" / "bom_to_control_table.py"
ICON_PATH = PROJECT_DIR / "assets" / "YR.ico"
DEFAULT_TEMPLATE_PATH = "" if IS_FROZEN else str(PROJECT_DIR / "客户XXX-项目名XXX-基础数据管制表-日期.xlsx")
bom: Any = None


def load_bom_core() -> Any:
    """按需加载 CATIA 内核，使 GUI 在依赖未安装时仍能打开并提示修复方式。"""
    global bom
    if bom is not None:
        return bom
    try:
        bom = importlib.import_module("catia_bom_export_v2")
    except SystemExit as exc:
        if IS_FROZEN:
            raise RuntimeError(
                "软件运行组件不完整，请重新运行 YR 安装程序执行安装或修复。\n\n"
                f"原始提示：{exc}"
            ) from exc
        raise RuntimeError(
            "导出内核依赖未安装。请在本目录终端执行：\n"
            "python -m pip install -r requirements.txt\n\n"
            f"原始提示：{exc}"
        ) from exc
    except Exception as exc:
        raise RuntimeError(f"无法加载 CATIA BOM 导出内核：{exc}") from exc
    return bom


class GuiRunLog:
    """将既有导出内核日志同步给 GUI，而不改动其业务逻辑。"""

    def __init__(self, sink: Callable[[str, str, str], None]) -> None:
        self.entries: list[tuple[str, str, str]] = []
        self._sink = sink

    def add(self, level: str, message: str, item: str = "") -> None:
        stamp = datetime.now().strftime("%H:%M:%S")
        self.entries.append((level, item, message))
        self._sink(level, item, message)

    def info(self, message: str, item: str = "") -> None:
        self.add("INFO", message, item)

    def warning(self, message: str, item: str = "") -> None:
        self.add("WARNING", message, item)

    def error(self, message: str, item: str = "") -> None:
        self.add("ERROR", message, item)

    def save(self, path: Path) -> None:
        with path.open("w", encoding="utf-8") as handle:
            for level, item, message in self.entries:
                handle.write(f"{level}\t{item}\t{message}\n")


class GuiWorker(threading.Thread):
    """后台任务：线程内初始化 COM，所有 UI 更新经队列返回主线程。"""

    def __init__(self, action: str, payload: dict[str, Any], events: queue.Queue) -> None:
        super().__init__(daemon=True, name=f"catia-bom-{action}")
        self.action = action
        self.payload = payload
        self.events = events
        self.cancel_event = threading.Event()

    def request_cancel(self) -> None:
        self.cancel_event.set()
        self.events.put(("log", "WARNING", "", "已收到取消请求；将在当前安全步骤完成后停止。"))

    def _log(self, level: str, item: str, message: str) -> None:
        self.events.put(("log", level, item, message))

    def _progress(self, percent: int, message: str) -> None:
        self.events.put(("progress", max(0, min(100, int(percent))), message))

    def _cancel_requested(self) -> bool:
        return self.cancel_event.is_set()

    def run(self) -> None:
        success = False
        summary: dict[str, Any] = {}
        try:
            if self.action == "export":
                summary = self._run_export()
            elif self.action == "convert":
                summary = self._run_convert()
            else:
                raise RuntimeError(f"未知任务类型：{self.action}")
            success = not summary.get("cancelled", False) and not summary.get("failed", False)
        except Exception as exc:
            self._log("ERROR", "", f"任务异常：{exc}")
            self.events.put(("traceback", traceback.format_exc()))
            summary = {"failed": True, "message": str(exc)}
        finally:
            self.events.put(("finished", success, summary))

    def _create_log(self) -> GuiRunLog:
        return GuiRunLog(self._log)

    @staticmethod
    def _timestamp() -> str:
        return datetime.now().strftime("%Y%m%d_%H%M%S")

    def _save_log(self, log: GuiRunLog, output_dir: Path, name: str) -> Path:
        log_dir = output_dir / "日志"
        log_dir.mkdir(parents=True, exist_ok=True)
        path = log_dir / f"{name}_{self._timestamp()}.txt"
        log.save(path)
        return path

    def _run_export(self) -> dict[str, Any]:
        core = load_bom_core()
        output_dir = Path(self.payload["output_dir"]).resolve()
        output_dir.mkdir(parents=True, exist_ok=True)
        mode = self.payload["mode"]
        product_paths = [Path(value).resolve() for value in self.payload.get("products", [])]
        timestamp = self._timestamp()
        log = self._create_log()
        results: list[str] = []
        failures: list[str] = []
        cancelled = False

        import pythoncom

        pythoncom.CoInitialize()
        try:
            if mode == "active":
                self._progress(1, "正在连接 CATIA")
                catia = core.connect_catia(log, start_if_needed=False)
                product_doc, root_product = core.get_active_product_document(catia)
                root_name = core.com_text(root_product.PartNumber) or core.com_text(product_doc.Name)
                output = output_dir / f"CATIA_BOM_{core.clean_filename(root_name)}_{timestamp}.xlsx"
                image_dir = output_dir / "零件图片" / core.clean_filename(root_name)
                self._progress(2, f"开始导出：{root_name}")
                try:
                    core.export_product_document(
                        catia,
                        product_doc,
                        output,
                        log,
                        image_dir,
                        self._progress,
                        self._cancel_requested,
                    )
                    results.append(str(output))
                except core.ExportCancelled as exc:
                    cancelled = True
                    log.warning(str(exc), root_name)
            else:
                if not product_paths:
                    raise ValueError("请至少添加一个 CATProduct 文件。")
                missing = [str(path) for path in product_paths if not path.is_file()]
                if missing:
                    raise FileNotFoundError("找不到 CATProduct：\n" + "\n".join(missing))
                self._progress(1, "正在连接或启动 CATIA")
                catia = core.connect_catia(log, start_if_needed=True)
                total = len(product_paths)
                for index, product_path in enumerate(product_paths, 1):
                    if self._cancel_requested():
                        cancelled = True
                        break
                    document = None
                    item_log = self._create_log()
                    try:
                        self._progress(round((index - 1) * 100 / total), f"打开 {index}/{total}：{product_path.name}")
                        item_log.info("打开 CATProduct", str(product_path))
                        document = catia.Documents.Open(str(product_path))
                        if not core.com_text(document.Name).lower().endswith(".catproduct"):
                            raise core.BomExportError(f"打开的文件不是 CATProduct：{product_path}")
                        stem = core.clean_filename(product_path.stem)
                        unique_stem = f"{index:03d}_{stem}"
                        output = output_dir / f"CATIA_BOM_{unique_stem}_{timestamp}.xlsx"
                        image_dir = output_dir / "零件图片" / unique_stem

                        def item_progress(percent: int, message: str, current: int = index) -> None:
                            overall = round(((current - 1) + percent / 100) * 100 / total)
                            self._progress(overall, f"[{current}/{total}] {message}")

                        core.export_product_document(
                            catia,
                            document,
                            output,
                            item_log,
                            image_dir,
                            item_progress,
                            self._cancel_requested,
                        )
                        results.append(str(output))
                        item_log.info("任务完成", str(output))
                    except core.ExportCancelled as exc:
                        cancelled = True
                        item_log.warning(str(exc), str(product_path))
                    except Exception as exc:
                        failures.append(f"{product_path.name}：{exc}")
                        item_log.error(f"任务失败：{exc}", str(product_path))
                    finally:
                        if document is not None:
                            try:
                                core._close_document_without_saving(catia, document, item_log, product_path.name)
                            except Exception as exc:
                                failures.append(f"{product_path.name}：关闭失败：{exc}")
                                item_log.error(f"关闭 CATProduct 失败：{exc}", str(product_path))
                        self._save_log(item_log, output_dir, f"导出_{core.clean_filename(product_path.stem)}")
                    if cancelled:
                        break
                if results and not cancelled:
                    self._progress(100, "批量导出完成")
        finally:
            log_path = self._save_log(log, output_dir, "GUI导出")
            pythoncom.CoUninitialize()

        return {
            "cancelled": cancelled,
            "failed": bool(failures),
            "results": results,
            "failures": failures,
            "log": str(log_path),
            "output_dir": str(output_dir),
        }

    def _run_convert(self) -> dict[str, Any]:
        bom_paths = [Path(value).resolve() for value in self.payload["bom_paths"]]
        template_path = Path(self.payload["template_path"]).resolve()
        output_path = Path(self.payload["output_path"]).resolve()
        if self._cancel_requested():
            return {"cancelled": True}
        if not CONVERTER_PATH.is_file():
            raise FileNotFoundError(f"未找到模板转换模块：{CONVERTER_PATH}")
        if not bom_paths:
            raise ValueError("请至少添加一份 CATIA BOM 图片清单。")
        missing = [str(path) for path in bom_paths if not path.is_file()]
        if missing or not template_path.is_file():
            details = "\n".join(missing)
            raise FileNotFoundError(f"请检查 BOM 文件和模板文件是否存在。\n{details}")
        if output_path in [*bom_paths, template_path]:
            raise ValueError("输出文件不能覆盖 BOM 文件或模板文件。")

        self._progress(1, f"准备合并 {len(bom_paths)} 份 BOM")
        self._log("INFO", "", "待合并 BOM：\n" + "\n".join(str(path) for path in bom_paths))
        spec = importlib.util.spec_from_file_location("bom_to_control_table_gui", CONVERTER_PATH)
        if spec is None or spec.loader is None:
            raise RuntimeError("无法加载模板转换模块。")
        converter = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = converter
        spec.loader.exec_module(converter)
        try:
            converter.convert_many(
                bom_paths,
                template_path,
                output_path,
                progress_callback=self._progress,
                cancel_requested=self._cancel_requested,
            )
        except converter.ConversionCancelled as exc:
            self._log("WARNING", "", str(exc))
            return {"cancelled": True, "output_dir": str(output_path.parent)}
        self._log("INFO", "", f"已将 {len(bom_paths)} 份 BOM 合并写入同一模板。")
        return {"results": [str(output_path)], "output_dir": str(output_path.parent)}


class BomManagerApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_NAME)
        if ICON_PATH.is_file():
            try:
                self.iconbitmap(default=str(ICON_PATH))
            except tk.TclError:
                pass
        self.geometry("1120x740")
        self.minsize(940, 620)
        self.events: queue.Queue = queue.Queue()
        self.worker: Optional[GuiWorker] = None
        self.product_paths: list[str] = []
        self.last_output_dir: Optional[Path] = None

        self.export_mode = tk.StringVar(value="active")
        self.export_directory = tk.StringVar(value=str(DEFAULT_EXPORT_DIR))
        self.template_bom_paths: list[str] = []
        self.template_file = tk.StringVar(value=DEFAULT_TEMPLATE_PATH)
        self.template_output = tk.StringVar()
        self.status = tk.StringVar(value="就绪")
        self.progress_value = tk.IntVar(value=0)

        self._configure_style()
        self._build_ui()
        self.after(100, self._process_events)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 16, "bold"), foreground="#17365D")
        style.configure("Hint.TLabel", foreground="#5B6573")
        style.configure("Accent.TButton", padding=(12, 7))

    def _build_ui(self) -> None:
        container = ttk.Frame(self, padding=16)
        container.pack(fill="both", expand=True)
        ttk.Label(container, text="CATIA BOM Manager", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            container,
            text="CATIA BOM 导出 · 图片清单 · 基础数据管制表转换",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(3, 12))

        notebook = ttk.Notebook(container)
        notebook.pack(fill="both", expand=True)
        self.export_tab = ttk.Frame(notebook, padding=14)
        self.convert_tab = ttk.Frame(notebook, padding=14)
        self.log_tab = ttk.Frame(notebook, padding=10)
        notebook.add(self.export_tab, text="导出 BOM")
        notebook.add(self.convert_tab, text="转换模板")
        notebook.add(self.log_tab, text="运行日志")
        self._build_export_tab()
        self._build_convert_tab()
        self._build_log_tab()

        footer = ttk.Frame(container, padding=(0, 12, 0, 0))
        footer.pack(fill="x")
        self.progress_bar = ttk.Progressbar(footer, maximum=100, variable=self.progress_value)
        self.progress_bar.pack(side="left", fill="x", expand=True)
        ttk.Label(footer, textvariable=self.status, width=42, anchor="w").pack(side="left", padx=(12, 8))
        self.cancel_button = ttk.Button(footer, text="取消当前任务", command=self._cancel_current, state="disabled")
        self.cancel_button.pack(side="right")
        ttk.Button(footer, text="打开输出目录", command=self._open_output_directory).pack(side="right", padx=(0, 8))

    @staticmethod
    def _add_path_row(parent: ttk.Frame, row: int, label: str, variable: tk.StringVar, command: Callable[[], None]) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=(10, 8), pady=6)
        ttk.Button(parent, text="浏览…", command=command).grid(row=row, column=2, sticky="e", pady=6)

    def _build_export_tab(self) -> None:
        tab = self.export_tab
        tab.columnconfigure(1, weight=1)
        ttk.Label(tab, text="导出方式").grid(row=0, column=0, sticky="nw", pady=6)
        modes = ttk.Frame(tab)
        modes.grid(row=0, column=1, columnspan=2, sticky="w", pady=6)
        ttk.Radiobutton(modes, text="当前 CATIA 活动总成", variable=self.export_mode, value="active", command=self._refresh_mode).pack(side="left")
        ttk.Radiobutton(modes, text="多个 CATProduct 批量导出", variable=self.export_mode, value="batch", command=self._refresh_mode).pack(side="left", padx=(20, 0))
        self._add_path_row(tab, 1, "导出目录", self.export_directory, self._choose_export_directory)

        ttk.Label(tab, text="批量任务").grid(row=2, column=0, sticky="nw", pady=(12, 6))
        task_frame = ttk.Frame(tab)
        task_frame.grid(row=2, column=1, columnspan=2, sticky="nsew", pady=(12, 6))
        tab.rowconfigure(2, weight=1)
        task_frame.columnconfigure(0, weight=1)
        self.product_list = tk.Listbox(task_frame, height=11, selectmode="extended", activestyle="none")
        self.product_list.grid(row=0, column=0, rowspan=4, sticky="nsew")
        scrollbar = ttk.Scrollbar(task_frame, orient="vertical", command=self.product_list.yview)
        scrollbar.grid(row=0, column=1, rowspan=4, sticky="ns")
        self.product_list.configure(yscrollcommand=scrollbar.set)
        self.add_products_button = ttk.Button(task_frame, text="添加 CATProduct…", command=self._add_products)
        self.add_products_button.grid(row=0, column=2, sticky="ew", padx=(10, 0))
        self.remove_products_button = ttk.Button(task_frame, text="移除选中", command=self._remove_selected_products)
        self.remove_products_button.grid(row=1, column=2, sticky="ew", padx=(10, 0), pady=6)
        self.clear_products_button = ttk.Button(task_frame, text="清空列表", command=self._clear_products)
        self.clear_products_button.grid(row=2, column=2, sticky="ew", padx=(10, 0))
        ttk.Label(task_frame, text="批量模式将依次打开、导出并关闭文件，不保存客户源文件。", style="Hint.TLabel", wraplength=220).grid(row=3, column=2, sticky="sw", padx=(10, 0))

        button_frame = ttk.Frame(tab)
        button_frame.grid(row=3, column=1, columnspan=2, sticky="e", pady=(16, 0))
        self.export_button = ttk.Button(button_frame, text="开始导出 BOM", style="Accent.TButton", command=self._start_export)
        self.export_button.pack(side="right")
        ttk.Button(button_frame, text="检测 CATIA", command=self._check_catia_ready).pack(side="right", padx=(0, 8))
        self._refresh_mode()

    def _build_convert_tab(self) -> None:
        tab = self.convert_tab
        tab.columnconfigure(1, weight=1)
        tab.rowconfigure(1, weight=1)
        ttk.Label(
            tab,
            text="可将 1 份或多份 CATIA 图片清单合并写入同一份基础数据管制表模板；每个总成会保留独立层级、图片与重量汇总。",
            style="Hint.TLabel",
            wraplength=760,
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))

        ttk.Label(tab, text="待合并 BOM").grid(row=1, column=0, sticky="nw", pady=6)
        bom_frame = ttk.Frame(tab)
        bom_frame.grid(row=1, column=1, columnspan=2, sticky="nsew", pady=6)
        bom_frame.columnconfigure(0, weight=1)
        bom_frame.rowconfigure(0, weight=1)
        self.convert_bom_list = tk.Listbox(bom_frame, height=8, selectmode="extended", activestyle="none")
        self.convert_bom_list.grid(row=0, column=0, rowspan=3, sticky="nsew")
        scrollbar = ttk.Scrollbar(bom_frame, orient="vertical", command=self.convert_bom_list.yview)
        scrollbar.grid(row=0, column=1, rowspan=3, sticky="ns")
        self.convert_bom_list.configure(yscrollcommand=scrollbar.set)
        self.add_convert_boms_button = ttk.Button(bom_frame, text="添加 BOM…", command=self._add_template_boms)
        self.add_convert_boms_button.grid(row=0, column=2, sticky="ew", padx=(10, 0))
        self.remove_convert_boms_button = ttk.Button(bom_frame, text="移除选中", command=self._remove_selected_template_boms)
        self.remove_convert_boms_button.grid(row=1, column=2, sticky="ew", padx=(10, 0), pady=6)
        self.clear_convert_boms_button = ttk.Button(bom_frame, text="清空列表", command=self._clear_template_boms)
        self.clear_convert_boms_button.grid(row=2, column=2, sticky="ew", padx=(10, 0))

        self._add_path_row(tab, 2, "管制表模板", self.template_file, self._choose_template_file)
        self._add_path_row(tab, 3, "合并输出文件", self.template_output, self._choose_template_output)
        button_frame = ttk.Frame(tab)
        button_frame.grid(row=4, column=1, columnspan=2, sticky="e", pady=(18, 0))
        self.convert_button = ttk.Button(button_frame, text="合并转换为一份模板", style="Accent.TButton", command=self._start_conversion)
        self.convert_button.pack(side="right")

    def _build_log_tab(self) -> None:
        tab = self.log_tab
        tab.rowconfigure(0, weight=1)
        tab.columnconfigure(0, weight=1)
        self.log_text = tk.Text(tab, wrap="word", state="disabled", font=("Consolas", 10), background="#10151C", foreground="#E8EEF7", insertbackground="white")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.tag_configure("ERROR", foreground="#FF7B72")
        self.log_text.tag_configure("WARNING", foreground="#E3B341")
        self.log_text.tag_configure("INFO", foreground="#79C0FF")
        toolbar = ttk.Frame(tab)
        toolbar.grid(row=1, column=0, sticky="e", pady=(8, 0))
        ttk.Button(toolbar, text="清空日志", command=self._clear_log).pack()

    def _refresh_mode(self) -> None:
        state = "normal" if self.export_mode.get() == "batch" else "disabled"
        for widget in (self.product_list, self.add_products_button, self.remove_products_button, self.clear_products_button):
            try:
                widget.configure(state=state)
            except tk.TclError:
                pass

    def _choose_export_directory(self) -> None:
        path = filedialog.askdirectory(title="选择 BOM 输出目录", initialdir=self.export_directory.get() or str(PROJECT_DIR))
        if path:
            self.export_directory.set(path)

    def _add_products(self) -> None:
        names = filedialog.askopenfilenames(title="选择根总成 CATProduct", filetypes=[("CATProduct", "*.CATProduct"), ("所有文件", "*.*")])
        for name in names:
            if name not in self.product_paths:
                self.product_paths.append(name)
                self.product_list.insert("end", name)

    def _remove_selected_products(self) -> None:
        selected = list(self.product_list.curselection())
        for index in reversed(selected):
            del self.product_paths[index]
            self.product_list.delete(index)

    def _clear_products(self) -> None:
        self.product_paths.clear()
        self.product_list.delete(0, "end")

    def _add_template_boms(self) -> None:
        values = filedialog.askopenfilenames(
            title="选择一个或多个 CATIA BOM 图片清单",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        for value in values:
            if value not in self.template_bom_paths:
                self.template_bom_paths.append(value)
                self.convert_bom_list.insert("end", value)
        if values and not self.template_output.get():
            first_bom = Path(values[0])
            self.template_output.set(str(first_bom.with_name("合并_基础数据管制表.xlsx")))

    def _remove_selected_template_boms(self) -> None:
        selected = list(self.convert_bom_list.curselection())
        for index in reversed(selected):
            del self.template_bom_paths[index]
            self.convert_bom_list.delete(index)

    def _clear_template_boms(self) -> None:
        self.template_bom_paths.clear()
        self.convert_bom_list.delete(0, "end")

    def _choose_template_file(self) -> None:
        value = filedialog.askopenfilename(title="选择基础数据管制表模板", filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
        if value:
            self.template_file.set(value)

    def _choose_template_output(self) -> None:
        initial = self.template_output.get() or str(PROJECT_DIR / "模板转化" / "基础数据管制表_结果.xlsx")
        value = filedialog.asksaveasfilename(title="保存转换结果", initialfile=Path(initial).name, initialdir=str(Path(initial).parent), defaultextension=".xlsx", filetypes=[("Excel 工作簿", "*.xlsx")])
        if value:
            self.template_output.set(value)

    def _check_catia_ready(self) -> None:
        try:
            import pythoncom
            pythoncom.CoInitialize()
            try:
                core = load_bom_core()
                log = GuiRunLog(lambda _level, _item, _message: None)
                catia = core.connect_catia(log, start_if_needed=False)
                document, _ = core.get_active_product_document(catia)
                messagebox.showinfo("CATIA 检测", f"连接正常。\n当前活动文档：{document.Name}")
            finally:
                pythoncom.CoUninitialize()
        except Exception as exc:
            messagebox.showerror("CATIA 检测失败", str(exc))

    def _start_export(self) -> None:
        directory = self.export_directory.get().strip()
        if not directory:
            messagebox.showwarning("缺少输出目录", "请选择 BOM 输出目录。")
            return
        if self.export_mode.get() == "batch" and not self.product_paths:
            messagebox.showwarning("缺少任务", "批量模式请至少添加一个 CATProduct。")
            return
        self._start_worker("export", {"mode": self.export_mode.get(), "output_dir": directory, "products": list(self.product_paths)})

    def _start_conversion(self) -> None:
        values = {
            "bom_paths": list(self.template_bom_paths),
            "template_path": self.template_file.get().strip(),
            "output_path": self.template_output.get().strip(),
        }
        if not values["bom_paths"] or not values["template_path"] or not values["output_path"]:
            messagebox.showwarning("文件未选择", "请至少添加一份 BOM，并选择模板和输出文件。")
            return
        self._start_worker("convert", values)

    def _start_worker(self, action: str, payload: dict[str, Any]) -> None:
        if self.worker is not None and self.worker.is_alive():
            messagebox.showwarning("任务进行中", "请等待当前任务结束，或先取消任务。")
            return
        self.progress_value.set(0)
        self.status.set("任务启动中")
        self.export_button.configure(state="disabled")
        self.convert_button.configure(state="disabled")
        self.cancel_button.configure(state="normal")
        self._append_log("INFO", "", f"启动任务：{action}")
        self.worker = GuiWorker(action, payload, self.events)
        self.worker.start()

    def _cancel_current(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            self.worker.request_cancel()
            self.cancel_button.configure(state="disabled")
            self.status.set("正在安全取消…")

    def _process_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                kind = event[0]
                if kind == "log":
                    _, level, item, message = event
                    self._append_log(level, item, message)
                elif kind == "progress":
                    _, value, message = event
                    self.progress_value.set(value)
                    self.status.set(message)
                elif kind == "traceback":
                    self._append_log("ERROR", "", event[1])
                elif kind == "finished":
                    _, success, summary = event
                    self._finish_task(success, summary)
        except queue.Empty:
            pass
        self.after(100, self._process_events)

    def _finish_task(self, success: bool, summary: dict[str, Any]) -> None:
        self.export_button.configure(state="normal")
        self.convert_button.configure(state="normal")
        self.cancel_button.configure(state="disabled")
        self.worker = None
        output_dir = summary.get("output_dir")
        if output_dir:
            self.last_output_dir = Path(output_dir)
        if summary.get("cancelled"):
            self.status.set("任务已取消")
            self._append_log("WARNING", "", "任务已在安全步骤结束后取消。")
            return
        results = summary.get("results", [])
        failures = summary.get("failures", [])
        if success:
            self.status.set("任务完成")
            self.progress_value.set(100)
            self._append_log("INFO", "", "任务完成：\n" + "\n".join(results))
            messagebox.showinfo("任务完成", f"已生成 {len(results)} 个文件。")
        else:
            self.status.set("任务完成，但存在错误")
            for failure in failures:
                self._append_log("ERROR", "", failure)
            if summary.get("message"):
                self._append_log("ERROR", "", summary["message"])
            messagebox.showwarning("任务结束", "任务存在失败项，请查看“运行日志”。")

    def _append_log(self, level: str, item: str, message: str) -> None:
        stamp = datetime.now().strftime("%H:%M:%S")
        prefix = f"[{stamp}] {level:<7}"
        if item:
            prefix += f" [{item}]"
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"{prefix} {message}\n", level)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _open_output_directory(self) -> None:
        path = self.last_output_dir or Path(self.export_directory.get() or DEFAULT_EXPORT_DIR)
        path.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(str(path))
        except OSError as exc:
            messagebox.showerror("打开目录失败", str(exc))

    def _on_close(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            if not messagebox.askyesno("任务进行中", "当前任务仍在运行。是否请求安全取消？"):
                return
            self.worker.request_cancel()
            messagebox.showinfo("正在取消", "取消请求已发送。请等待当前安全步骤结束后再关闭窗口。")
            return
        self.destroy()


def run_packaged_self_check() -> int:
    """Validate bundled runtime components without connecting to or changing CATIA."""
    try:
        import openpyxl  # noqa: F401
        import pythoncom  # noqa: F401
        from PIL import Image  # noqa: F401

        load_bom_core()
        if not CONVERTER_PATH.is_file():
            raise FileNotFoundError(f"模板转换组件不存在：{CONVERTER_PATH}")
        spec = importlib.util.spec_from_file_location("bom_to_control_table_self_check", CONVERTER_PATH)
        if spec is None or spec.loader is None:
            raise RuntimeError("无法加载模板转换组件。")
        converter = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = converter
        spec.loader.exec_module(converter)
        return 0
    except Exception:
        traceback.print_exc()
        return 1


def main() -> None:
    if "--self-check" in sys.argv:
        raise SystemExit(run_packaged_self_check())
    app = BomManagerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
