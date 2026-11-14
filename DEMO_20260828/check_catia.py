# -*- coding: utf-8 -*-
"""只读诊断脚本：检查此终端能否连接同一 Windows 会话中的 CATIA。"""

from __future__ import annotations

import platform
import struct
import sys


def main() -> int:
    print("=" * 64)
    print("ZYR · CATIA Connection Check")
    print("This check will not change CATIA or any files.")
    print("=" * 64)
    print(f"Python executable : {sys.executable}")
    print(f"Python version    : {platform.python_version()}")
    print(f"Python bitness    : {struct.calcsize('P') * 8}-bit")
    print()

    try:
        import pythoncom
        import pywintypes
        import win32com.client
    except ImportError as exc:
        print(f"RESULT: Python dependency missing: {exc}")
        print("Run: python -m pip install pywin32 openpyxl Pillow")
        return 2

    pythoncom.CoInitialize()
    try:
        try:
            catia = win32com.client.GetActiveObject("CATIA.Application")
        except pywintypes.com_error as exc:
            hresult = getattr(exc, "hresult", None)
            print("RESULT: Cannot connect to a running CATIA instance.")
            if isinstance(hresult, int):
                print(f"COM HRESULT       : 0x{hresult & 0xFFFFFFFF:08X}")
            print(f"COM detail        : {exc}")
            print()
            print("Check these items:")
            print("1. CATIA is fully open (no modal dialog, login, or recovery window).")
            print("2. CATIA and this terminal use the same Windows user and session.")
            print("3. Both run at the same privilege level (normally, both non-admin).")
            print("4. Prefer Python with the same 32/64-bit architecture as CATIA.")
            return 1

        print("RESULT: Connected to CATIA successfully.")
        try:
            print(f"CATIA visible     : {bool(catia.Visible)}")
        except Exception:
            pass
        try:
            document = catia.ActiveDocument
            document_name = str(document.Name)
            print(f"Active document   : {document_name}")
            if document_name.lower().endswith(".catproduct"):
                print("Document check    : PASS (.CATProduct)")
                return 0
            print("Document check    : NOT a .CATProduct; activate the target assembly first.")
            return 1
        except Exception as exc:
            print(f"Active document   : Cannot read ({exc})")
            return 1
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
