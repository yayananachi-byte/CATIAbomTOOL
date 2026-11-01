# -*- coding: utf-8 -*-
"""YR CATIA BOM Manager offline per-user installer.

This program is packaged as a single EXE.  It installs the already-frozen
application payload into the current user's LocalAppData folder, so neither
installation nor normal use requires administrator privileges.
"""

from __future__ import annotations

import ctypes
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import winreg
from datetime import datetime
from pathlib import Path
from typing import Callable

import tkinter as tk
from tkinter import messagebox, ttk


APP_NAME = "YR CATIA BOM Manager"
APP_VERSION = "1.0.0"
PUBLISHER = "ZYR"
APP_EXE_NAME = "YR_CATIA_BOM_Manager.exe"
SETUP_EXE_NAME = "YR_CATIA_BOM_Setup.exe"
UNINSTALL_KEY = rf"Software\Microsoft\Windows\CurrentVersion\Uninstall\{APP_NAME}"

RESOURCE_ROOT = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
PAYLOAD_DIR = RESOURCE_ROOT / "payload" / "YR_CATIA_BOM_Manager"
ICON_PATH = RESOURCE_ROOT / "assets" / "YR.ico"

LOCAL_APPDATA = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
INSTALL_DIR = LOCAL_APPDATA / "Programs" / APP_NAME
APP_EXE_PATH = INSTALL_DIR / APP_EXE_NAME
INSTALLED_SETUP_PATH = INSTALL_DIR / SETUP_EXE_NAME
USER_DATA_DIR = Path.home() / "Documents" / APP_NAME


def is_installed() -> bool:
    return APP_EXE_PATH.is_file()


def validate_install_target() -> None:
    expected = (LOCAL_APPDATA / "Programs" / APP_NAME).resolve()
    actual = INSTALL_DIR.resolve()
    if actual != expected or actual.name != APP_NAME:
        raise RuntimeError(f"安装目录安全校验失败：{actual}")


def validate_user_data_target() -> None:
    """Only allow removal of this product's exact Documents subdirectory."""
    documents_dir = (Path.home() / "Documents").resolve()
    configured_parent = USER_DATA_DIR.parent.resolve()
    if (
        USER_DATA_DIR.name != APP_NAME
        or configured_parent != documents_dir
        or USER_DATA_DIR.is_symlink()
    ):
        raise RuntimeError(f"用户数据目录安全校验失败，拒绝删除：{USER_DATA_DIR}")


def get_shell_paths() -> tuple[Path, Path]:
    import win32com.client

    shell = win32com.client.Dispatch("WScript.Shell")
    desktop = Path(shell.SpecialFolders("Desktop"))
    programs = Path(shell.SpecialFolders("Programs"))
    return desktop, programs


def create_shortcut(
    path: Path,
    target: Path,
    working_directory: Path,
    arguments: str = "",
) -> None:
    import win32com.client

    path.parent.mkdir(parents=True, exist_ok=True)
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(str(path))
    shortcut.TargetPath = str(target)
    shortcut.Arguments = arguments
    shortcut.WorkingDirectory = str(working_directory)
    shortcut.IconLocation = f"{target},0"
    shortcut.Description = APP_NAME
    shortcut.Save()


def shortcut_paths() -> tuple[Path, Path, Path]:
    desktop, programs = get_shell_paths()
    start_menu_dir = programs / APP_NAME
    return (
        desktop / f"{APP_NAME}.lnk",
        start_menu_dir / f"{APP_NAME}.lnk",
        start_menu_dir / f"卸载 {APP_NAME}.lnk",
    )


def create_shortcuts() -> None:
    desktop_link, start_link, uninstall_link = shortcut_paths()
    create_shortcut(desktop_link, APP_EXE_PATH, INSTALL_DIR)
    create_shortcut(start_link, APP_EXE_PATH, INSTALL_DIR)
    create_shortcut(uninstall_link, INSTALLED_SETUP_PATH, INSTALL_DIR, "--uninstall")


def remove_shortcuts() -> None:
    try:
        desktop_link, start_link, uninstall_link = shortcut_paths()
    except Exception:
        return
    for path in (desktop_link, start_link, uninstall_link):
        try:
            path.unlink(missing_ok=True)
        except OSError:
            pass
    start_menu_dir = start_link.parent
    try:
        start_menu_dir.rmdir()
    except OSError:
        pass


def register_uninstaller() -> None:
    uninstall_command = f'"{INSTALLED_SETUP_PATH}" --uninstall'
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, UNINSTALL_KEY) as key:
        winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, APP_NAME)
        winreg.SetValueEx(key, "DisplayVersion", 0, winreg.REG_SZ, APP_VERSION)
        winreg.SetValueEx(key, "Publisher", 0, winreg.REG_SZ, PUBLISHER)
        winreg.SetValueEx(key, "InstallLocation", 0, winreg.REG_SZ, str(INSTALL_DIR))
        winreg.SetValueEx(key, "DisplayIcon", 0, winreg.REG_SZ, f"{APP_EXE_PATH},0")
        winreg.SetValueEx(key, "UninstallString", 0, winreg.REG_SZ, uninstall_command)
        winreg.SetValueEx(key, "NoModify", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(key, "NoRepair", 0, winreg.REG_DWORD, 1)


def unregister_uninstaller() -> None:
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, UNINSTALL_KEY)
    except FileNotFoundError:
        pass


def write_install_info() -> None:
    info = {
        "name": APP_NAME,
        "version": APP_VERSION,
        "publisher": PUBLISHER,
        "installed_at": datetime.now().isoformat(timespec="seconds"),
        "install_dir": str(INSTALL_DIR),
        "user_data_dir": str(USER_DATA_DIR),
        "python_runtime": "bundled",
        "administrator_required": False,
    }
    (INSTALL_DIR / "install_info.json").write_text(
        json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def copy_application_payload(progress: Callable[[str], None]) -> None:
    if not PAYLOAD_DIR.is_dir():
        raise FileNotFoundError(f"安装包缺少应用文件：{PAYLOAD_DIR}")
    validate_install_target()
    INSTALL_DIR.mkdir(parents=True, exist_ok=True)

    items = list(PAYLOAD_DIR.iterdir())
    for index, source in enumerate(items, 1):
        progress(f"正在复制应用组件 {index}/{len(items)}：{source.name}")
        destination = INSTALL_DIR / source.name
        if source.is_dir():
            shutil.copytree(source, destination, dirs_exist_ok=True)
        else:
            shutil.copy2(source, destination)

    if not APP_EXE_PATH.is_file():
        raise RuntimeError(f"主程序安装失败，未找到：{APP_EXE_PATH}")


def install_application(progress: Callable[[str], None]) -> str:
    progress("正在检查当前用户安装目录")
    validate_install_target()
    progress("正在安装内置 Python 运行环境和软件组件")
    copy_application_payload(progress)

    progress("正在验证内置 Python、Excel、图片和 CATIA COM 组件")
    check_result = subprocess.run(
        [str(APP_EXE_PATH), "--self-check"],
        cwd=str(INSTALL_DIR),
        timeout=60,
        check=False,
    )
    if check_result.returncode != 0:
        raise RuntimeError(
            f"软件运行组件自检失败（错误码 {check_result.returncode}），请重新下载安装包。"
        )

    progress("正在写入卸载程序")
    shutil.copy2(Path(sys.executable), INSTALLED_SETUP_PATH)
    USER_DATA_DIR.mkdir(parents=True, exist_ok=True)
    write_install_info()

    progress("正在创建桌面和开始菜单快捷方式")
    create_shortcuts()
    register_uninstaller()
    return (
        f"{APP_NAME} {APP_VERSION} 安装完成。\n\n"
        f"安装位置：{INSTALL_DIR}\n"
        f"输出文件默认保存在：{USER_DATA_DIR}\n\n"
        "Python 和所需组件已经内置，不需要另外安装。"
    )


def uninstall_application(
    progress: Callable[[str], None],
    preserve_user_data: bool = True,
) -> str:
    progress("正在移除快捷方式和卸载信息")
    remove_shortcuts()
    unregister_uninstaller()
    validate_install_target()

    progress("正在卸载应用程序文件")
    if INSTALL_DIR.exists():
        shutil.rmtree(INSTALL_DIR)

    if preserve_user_data:
        data_message = f"已保留本地输出文件：\n{USER_DATA_DIR}"
    else:
        progress("正在删除本地 BOM、图片、日志和模板结果")
        validate_user_data_target()
        if USER_DATA_DIR.exists():
            shutil.rmtree(USER_DATA_DIR)
        data_message = f"本地输出文件已删除：\n{USER_DATA_DIR}"
    return f"{APP_NAME} 已卸载。\n\n{data_message}"


def executable_is_inside_install_dir() -> bool:
    try:
        Path(sys.executable).resolve().relative_to(INSTALL_DIR.resolve())
        return True
    except ValueError:
        return False


def launch_temp_uninstaller(delete_user_data: bool = False) -> None:
    temp_exe = Path(tempfile.gettempdir()) / f"YR_CATIA_BOM_Uninstall_{os.getpid()}.exe"
    shutil.copy2(Path(sys.executable), temp_exe)
    arguments = [str(temp_exe), "--uninstall-now"]
    if delete_user_data:
        arguments.append("--delete-user-data")
    subprocess.Popen(
        arguments,
        close_fds=True,
        creationflags=getattr(subprocess, "DETACHED_PROCESS", 0),
    )


def schedule_temp_self_delete() -> None:
    if not Path(sys.executable).name.startswith("YR_CATIA_BOM_Uninstall_"):
        return
    fixed_script = "Start-Sleep -Seconds 3; Remove-Item -LiteralPath $args[0] -Force"
    try:
        subprocess.Popen(
            [
                "powershell.exe",
                "-NoProfile",
                "-WindowStyle",
                "Hidden",
                "-Command",
                fixed_script,
                str(Path(sys.executable)),
            ],
            close_fds=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except OSError:
        pass


def show_native_message(title: str, message: str, error: bool = False) -> None:
    flags = 0x10 if error else 0x40
    ctypes.windll.user32.MessageBoxW(None, message, title, flags)


def run_uninstall_now(delete_user_data: bool = False) -> int:
    time.sleep(1.5)
    try:
        message = uninstall_application(
            lambda _message: None,
            preserve_user_data=not delete_user_data,
        )
        show_native_message("卸载完成", message)
        return 0
    except Exception as exc:
        show_native_message("卸载失败", f"{exc}\n\n请先关闭正在运行的软件后重试。", error=True)
        return 1
    finally:
        schedule_temp_self_delete()


class InstallerApp(tk.Tk):
    def __init__(self, initial_action: str = "install") -> None:
        super().__init__()
        self.title(f"{APP_NAME} 安装程序")
        self.geometry("780x650")
        self.minsize(720, 600)
        self.resizable(True, True)
        self.configure(background="#07152f")
        if ICON_PATH.is_file():
            try:
                self.iconbitmap(default=str(ICON_PATH))
            except tk.TclError:
                pass

        self.action = tk.StringVar(value=initial_action)
        self.uninstall_data_mode = tk.StringVar(value="keep")
        self.status = tk.StringVar(value="请选择要执行的操作")
        self._configure_styles()
        self._build_ui()
        self._refresh_state()
        self._center_window()

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Installer.TFrame", background="#07152f")
        style.configure("Card.TFrame", background="#10264d")
        style.configure("Title.TLabel", background="#07152f", foreground="#ffffff", font=("Microsoft YaHei UI", 20, "bold"))
        style.configure("Sub.TLabel", background="#07152f", foreground="#a8c7ff", font=("Microsoft YaHei UI", 10))
        style.configure("CardTitle.TLabel", background="#10264d", foreground="#ffffff", font=("Microsoft YaHei UI", 13, "bold"))
        style.configure("CardText.TLabel", background="#10264d", foreground="#c7d8f5", font=("Microsoft YaHei UI", 9))
        style.configure("Action.TRadiobutton", background="#10264d", foreground="#ffffff", font=("Microsoft YaHei UI", 11, "bold"))
        style.map("Action.TRadiobutton", background=[("active", "#10264d")])
        style.configure("Primary.TButton", font=("Microsoft YaHei UI", 11, "bold"), padding=(18, 10))
        style.configure("Status.TLabel", background="#07152f", foreground="#a8c7ff", font=("Microsoft YaHei UI", 9))

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, style="Installer.TFrame", padding=(28, 24))
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        header = ttk.Frame(outer, style="Installer.TFrame")
        header.grid(row=0, column=0, sticky="ew")
        if (RESOURCE_ROOT / "assets" / "YR_icon.png").is_file():
            try:
                image = tk.PhotoImage(file=str(RESOURCE_ROOT / "assets" / "YR_icon.png"))
                self._logo_image = image.subsample(max(1, image.width() // 76))
                ttk.Label(header, image=self._logo_image, background="#07152f").pack(side="left", padx=(0, 18))
            except tk.TclError:
                pass
        title_box = ttk.Frame(header, style="Installer.TFrame")
        title_box.pack(side="left", fill="x", expand=True)
        ttk.Label(title_box, text=APP_NAME, style="Title.TLabel").pack(anchor="w")
        ttk.Label(title_box, text=f"离线安装程序 · 版本 {APP_VERSION} · 发布者 {PUBLISHER}", style="Sub.TLabel").pack(anchor="w", pady=(5, 0))

        card = ttk.Frame(outer, style="Card.TFrame", padding=22)
        card.grid(row=1, column=0, sticky="nsew", pady=(24, 16))
        ttk.Label(card, text="选择操作", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Radiobutton(
            card,
            text="安装 / 更新",
            variable=self.action,
            value="install",
            style="Action.TRadiobutton",
            command=self._refresh_state,
        ).pack(anchor="w", pady=(16, 2))
        ttk.Label(card, text="安装完整软件、内置 Python 和依赖组件，并创建快捷方式。", style="CardText.TLabel").pack(anchor="w", padx=(25, 0))
        ttk.Radiobutton(
            card,
            text="卸载",
            variable=self.action,
            value="uninstall",
            style="Action.TRadiobutton",
            command=self._refresh_state,
        ).pack(anchor="w", pady=(18, 2))
        ttk.Label(card, text="移除软件和快捷方式，并选择是否保留本地输出文件。", style="CardText.TLabel").pack(anchor="w", padx=(25, 0))

        data_options = ttk.Frame(card, style="Card.TFrame", padding=(25, 10, 0, 0))
        data_options.pack(fill="x")
        self.keep_data_radio = ttk.Radiobutton(
            data_options,
            text="保留本地文件（推荐）",
            variable=self.uninstall_data_mode,
            value="keep",
            style="Action.TRadiobutton",
        )
        self.keep_data_radio.pack(anchor="w")
        self.delete_data_radio = ttk.Radiobutton(
            data_options,
            text="不保留本地文件（删除 BOM、图片、日志和模板结果）",
            variable=self.uninstall_data_mode,
            value="delete",
            style="Action.TRadiobutton",
        )
        self.delete_data_radio.pack(anchor="w", pady=(6, 0))

        install_info = ttk.Label(
            outer,
            text=f"安装位置：{INSTALL_DIR}\n权限模式：当前用户安装，无需管理员权限",
            style="Sub.TLabel",
        )
        install_info.grid(row=2, column=0, sticky="w", pady=(0, 12))

        self.progress = ttk.Progressbar(outer, mode="indeterminate")
        self.progress.grid(row=3, column=0, sticky="ew")
        ttk.Label(outer, textvariable=self.status, style="Status.TLabel").grid(
            row=4, column=0, sticky="w", pady=(8, 12)
        )

        buttons = ttk.Frame(outer, style="Installer.TFrame")
        buttons.grid(row=5, column=0, sticky="ew")
        ttk.Button(buttons, text="退出", command=self.destroy).pack(side="right")
        self.execute_button = ttk.Button(buttons, text="开始安装", style="Primary.TButton", command=self._execute)
        self.execute_button.pack(side="right", padx=(0, 10))

    def _center_window(self) -> None:
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        required_width = max(780, self.winfo_reqwidth() + 20)
        required_height = max(650, self.winfo_reqheight() + 20)
        width = min(required_width, max(720, screen_width - 80))
        height = min(required_height, max(600, screen_height - 80))
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.geometry(f"{width}x{height}+{max(0, x)}+{max(0, y)}")

    def _refresh_state(self) -> None:
        installed = is_installed()
        action = self.action.get()
        self.execute_button.configure(text="开始安装" if action == "install" else "开始卸载")
        if action == "install":
            self.execute_button.configure(state="normal")
            self.status.set("检测到已安装版本，可执行覆盖更新。" if installed else "尚未安装，可以开始安装。")
        else:
            self.status.set("已检测到安装，可开始卸载。" if installed else "当前用户尚未安装本软件。")
            self.execute_button.configure(state="normal" if installed else "disabled")
        data_option_state = "normal" if action == "uninstall" and installed else "disabled"
        self.keep_data_radio.configure(state=data_option_state)
        self.delete_data_radio.configure(state=data_option_state)

    def _execute(self) -> None:
        action = self.action.get()
        delete_user_data = action == "uninstall" and self.uninstall_data_mode.get() == "delete"
        if action == "uninstall":
            if delete_user_data:
                confirmation = (
                    f"确定卸载 {APP_NAME} 并永久删除本地输出文件？\n\n"
                    "将删除 BOM、零件图片、日志和模板转换结果，且无法恢复。\n\n"
                    f"删除目录：{USER_DATA_DIR}"
                )
            else:
                confirmation = f"确定卸载 {APP_NAME}？\n\n用户输出文件将被保留。"
            if not messagebox.askyesno("确认卸载", confirmation):
                return
            if executable_is_inside_install_dir():
                try:
                    launch_temp_uninstaller(delete_user_data=delete_user_data)
                    self.destroy()
                except Exception as exc:
                    messagebox.showerror("无法启动卸载", str(exc))
                return

        self.execute_button.configure(state="disabled")
        self.progress.start(10)
        worker = threading.Thread(
            target=self._run_action,
            args=(action, delete_user_data),
            daemon=True,
        )
        worker.start()

    def _set_status(self, message: str) -> None:
        self.after(0, self.status.set, message)

    def _run_action(self, action: str, delete_user_data: bool = False) -> None:
        try:
            if action == "install":
                message = install_application(self._set_status)
            else:
                message = uninstall_application(
                    self._set_status,
                    preserve_user_data=not delete_user_data,
                )
            self.after(0, self._finish_success, action, message)
        except Exception as exc:
            details = traceback.format_exc()
            self.after(0, self._finish_error, exc, details)

    def _finish_success(self, action: str, message: str) -> None:
        self.progress.stop()
        self.execute_button.configure(state="normal")
        self.status.set("操作完成")
        messagebox.showinfo("操作完成", message)
        if action == "install" and APP_EXE_PATH.is_file():
            if messagebox.askyesno("启动软件", "是否现在启动 YR CATIA BOM Manager？"):
                subprocess.Popen([str(APP_EXE_PATH)], cwd=str(INSTALL_DIR), close_fds=True)
        self._refresh_state()

    def _finish_error(self, exc: Exception, details: str) -> None:
        self.progress.stop()
        self.execute_button.configure(state="normal")
        self.status.set("操作失败")
        log_path = Path(tempfile.gettempdir()) / "YR_CATIA_BOM_Setup_Error.log"
        try:
            log_path.write_text(details, encoding="utf-8")
        except OSError:
            pass
        messagebox.showerror(
            "操作失败",
            f"{exc}\n\n请先关闭正在运行的软件后重试。\n错误日志：{log_path}",
        )


def main() -> int:
    if "--uninstall-now" in sys.argv:
        return run_uninstall_now(delete_user_data="--delete-user-data" in sys.argv)
    initial_action = "uninstall" if "--uninstall" in sys.argv else "install"
    app = InstallerApp(initial_action=initial_action)
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
