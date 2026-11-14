# -*- coding: utf-8 -*-
"""只读检查：诊断 Python 能否连接到当前 Windows 会话中的 CATIA。"""

from __future__ import annotations

import platform
import struct
import sys

import pythoncom
import pywintypes
import win32com.client


def main() -> int:
    print("=" * 60)
    print("ZYR · CATIA 连接诊断（不会修改 CATIA 或文件）")
    print("=" * 60)
    print(f"Python：{sys.executable}")
    print(f"Python 版本：{platform.python_version()}")
    print(f"Python 位数：{struct.calcsize('P') * 8} 位")
    print()
    pythoncom.CoInitialize()
    try:
        catia = win32com.client.GetActiveObject("CATIA.Application")
        print("结果：成功连接到正在运行的 CATIA。")
        try:
            print(f"CATIA 可见：{bool(catia.Visible)}")
        except Exception:
            pass
        try:
            document = catia.ActiveDocument
            name = str(document.Name)
            print(f"当前活动文档：{name}")
            if name.lower().endswith(".catproduct"):
                print("活动文档检查：正确（CATProduct）。")
            else:
                print("活动文档检查：请在 CATIA 中激活要导出的 .CATProduct。")
        except Exception as exc:
            print(f"当前活动文档：无法读取（{exc}）")
        return 0
    except pywintypes.com_error as exc:
        hresult = getattr(exc, "hresult", None)
        print("结果：无法从 Python 连接到 CATIA。")
        if isinstance(hresult, int):
            print(f"COM HRESULT：0x{hresult & 0xFFFFFFFF:08X}")
        print(f"原始错误：{exc}")
        print()
        print("请依次检查：")
        print("1. CATIA 已完全启动，且不是停在登录、恢复或模态对话框。")
        print("2. CATIA 与本工具使用相同权限运行：两者都普通运行，或两者都管理员运行。")
        print("3. Python 与 CATIA 的位数一致（32 位 CATIA 用 32 位 Python；64 位同理）。")
        print("4. 工具与 CATIA 在同一 Windows 用户、同一远程桌面会话中运行。")
        return 1
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
