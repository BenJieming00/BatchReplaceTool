# -*- coding: utf-8 -*-
"""
Microsoft Word / Excel 无损格式转换（Windows COM 版）
对应 macOS 版的 AppleScript 引擎：
  - doc  -> docx: Word.Application，SaveAs2 FileFormat=12 (wdFormatXMLDocument)
  - xls  -> xlsx: Excel.Application，SaveAs FileFormat=51 (xlOpenXMLWorkbook)
仅 Windows + 已安装 Microsoft Office 环境可用；在 macOS/Linux 上导入不会报错，
但调用转换会抛出明确的"需要 Windows + Office"提示。
"""
import os
import sys

# ---- 平台与 Office 检测（不依赖 pywin32，可在任意平台安全导入） ----

def _win_registry_clsid(progid: str) -> bool:
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid):
            return True
    except Exception:
        return False


def word_available() -> bool:
    if sys.platform != "win32":
        return False
    return _win_registry_clsid("Word.Application")


def excel_available() -> bool:
    if sys.platform != "win32":
        return False
    return _win_registry_clsid("Excel.Application")


class OfficeError(Exception):
    pass


def _friendly_message(exc: Exception) -> str:
    msg = str(exc)
    low = msg.lower()
    if "拒绝访问" in msg or "access is denied" in low:
        return "文件被占用或无权限访问（请先关闭正在打开的该 Office 文件，或检查文件属性）"
    if "正在使用" in msg or "in use" in low or "share violation" in low:
        return "文件正被其他程序占用（请先关闭正在打开它的 Word/Excel/WPS 窗口）"
    if "rpc" in low or "远程过程调用" in msg or "服务器" in msg and "未运行" in msg:
        return "Office 未能正常启动（可能是首次启动较慢或已崩溃），请稍后重试"
    if "没有注册" in msg or "not registered" in low or "无效的类字符串" in msg:
        return "未检测到 Microsoft Office，doc/xls 转换需要安装 Microsoft Word / Excel"
    if "超出内存" in msg or "out of memory" in low:
        return "内存不足，请关闭部分程序后重试"
    return msg


def _norm(path: str) -> str:
    """转绝对路径（COM 需要完整路径）。"""
    return os.path.abspath(path)


def convert_doc_to_docx(src: str, dst: str) -> None:
    """doc -> docx（无损，SaveAs2 wdFormatXMLDocument=12）"""
    if sys.platform != "win32":
        raise OfficeError("doc→docx 转换依赖 Windows + Microsoft Word（当前非 Windows 环境）")
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        raise OfficeError("缺少 pywin32 依赖，请先运行: pip install pywin32")

    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        try:
            word.DisplayAlerts = 0  # wdAlertsNone：抑制弹窗
        except Exception:
            pass
        doc = word.Documents.Open(
            _norm(src), ReadOnly=False, AddToRecentFiles=False, Visible=False
        )
        try:
            doc.SaveAs2(_norm(dst), FileFormat=12)  # wdFormatXMLDocument
        except Exception:
            # 极老版本无 SaveAs2，退回 SaveAs
            doc.SaveAs(_norm(dst), FileFormat=12)
        doc.Close(False)
    except Exception as e:
        raise OfficeError(_friendly_message(e))
    finally:
        if word is not None:
            try:
                word.Quit(False)
            except Exception:
                pass
        pythoncom.CoUninitialize()


def convert_xls_to_xlsx(src: str, dst: str) -> None:
    """xls -> xlsx（无损，SaveAs xlOpenXMLWorkbook=51）"""
    if sys.platform != "win32":
        raise OfficeError("xls→xlsx 转换依赖 Windows + Microsoft Excel（当前非 Windows 环境）")
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        raise OfficeError("缺少 pywin32 依赖，请先运行: pip install pywin32")

    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        try:
            excel.DisplayAlerts = False
        except Exception:
            pass
        wb = excel.Workbooks.Open(
            _norm(src), ReadOnly=False, UpdateLinks=0, AddToMru=False
        )
        try:
            wb.SaveAs(_norm(dst), FileFormat=51)  # xlOpenXMLWorkbook
        except Exception:
            wb.SaveAs2(_norm(dst), FileFormat=51)
        wb.Close(False)
    except Exception as e:
        raise OfficeError(_friendly_message(e))
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
