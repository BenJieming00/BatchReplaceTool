# -*- coding: utf-8 -*-
"""
OOXML 批量文本替换（docx / xlsx）
与 macOS 版（Swift）逻辑一致：
  - xlsx: 处理 xl/sharedStrings.xml 的 <si> 块，合并其中所有 <t> 文本后替换，
          替换成功后把整个 si 重写为 <si><t xml:space="preserve">…</t></si>
  - docx: 处理 word/document.xml 中每个 <w:t> 文本节点
  - 替换文本写入前做 XML 转义（& < > " '）
  - 使用 zipfile 重打包，避免依赖外部 unzip/zip 命令
"""
import os
import re
import zipfile

DOCX_TARGET = "word/document.xml"
XLSX_TARGET = "xl/sharedStrings.xml"


class OoxmlError(Exception):
    pass


def _escape_xml(s: str) -> str:
    """XML 文本转义（与 Swift 版 escapeXML 一致）。"""
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def _apply_rules(text: str, rules, ignore_case: bool) -> tuple:
    """按规则依次替换，返回 (新文本, 替换处数)。find 不转义、replace 转义，与 macOS 版行为一致。"""
    new_text = text
    total = 0
    flags = re.IGNORECASE if ignore_case else 0
    for find, replace in rules:
        if not find:
            continue
        escaped_repl = _escape_xml(replace)
        pattern = re.compile(re.escape(find), flags)
        new_text, n = pattern.subn(escaped_repl, new_text)
        total += n
    return new_text, total


def _process_document_xml(xml: str, rules, ignore_case: bool) -> tuple:
    """docx: 替换每个 <w:t...>...</w:t> 内的文本。"""
    t_pattern = re.compile(r"<w:t[^>]*>(.*?)</w:t>", re.DOTALL)
    matches = list(t_pattern.finditer(xml))
    total = 0
    # 从后往前替换，避免位置偏移
    for m in reversed(matches):
        original = m.group(1)
        new_text, n = _apply_rules(original, rules, ignore_case)
        if n > 0:
            xml = xml[: m.start(1)] + _escape_xml(new_text) + xml[m.end(1):]
            total += n
    return xml, total


def _process_shared_strings(xml: str, rules, ignore_case: bool) -> tuple:
    """xlsx: 合并每个 <si> 内所有 <t> 文本后替换，重写 si 块。"""
    si_pattern = re.compile(r"<si>(.*?)</si>", re.DOTALL)
    t_pattern = re.compile(r"<t[^>]*>(.*?)</t>", re.DOTALL)
    matches = list(si_pattern.finditer(xml))
    total = 0
    for m in reversed(matches):
        si_block = m.group(0)
        texts = [t.group(1) for t in t_pattern.finditer(si_block)]
        combined = "".join(texts)
        new_text, n = _apply_rules(combined, rules, ignore_case)
        if n > 0:
            new_si = '<si><t xml:space="preserve">%s</t></si>' % _escape_xml(new_text)
            xml = xml[: m.start()] + new_si + xml[m.end():]
            total += n
    return xml, total


def replace_text(filepath: str, rules, ignore_case: bool = False) -> int:
    """对 docx/xlsx 执行批量文本替换，返回替换处数。"""
    ext = os.path.splitext(filepath)[1].lower()
    if ext not in (".docx", ".xlsx"):
        raise OoxmlError("不是有效的 docx/xlsx 文件")

    target = XLSX_TARGET if ext == ".xlsx" else DOCX_TARGET

    try:
        with zipfile.ZipFile(filepath, "r") as zin:
            names = zin.namelist()
            if target not in names:
                return 0
            infos = {n: zin.getinfo(n) for n in names}
            others = {n: zin.read(n) for n in names if n != target}
            data = zin.read(target)
    except zipfile.BadZipFile:
        raise OoxmlError("文件不是有效的 OOXML 压缩包（可能已损坏）")

    try:
        xml = data.decode("utf-8")
    except UnicodeDecodeError:
        raise OoxmlError("XML 编码错误")

    if ext == ".xlsx":
        xml, count = _process_shared_strings(xml, rules, ignore_case)
    else:
        xml, count = _process_document_xml(xml, rules, ignore_case)

    if count == 0:
        return 0

    # 重打包：临时名用 "xxx.docx.tmp"（避免双点文件名），成功后替换原文件
    tmp_path = filepath + ".tmp"
    if os.path.exists(tmp_path):
        os.remove(tmp_path)

    try:
        with zipfile.ZipFile(tmp_path, "w") as zout:
            # 目标 XML（已修改），保持原压缩方式与时间戳
            ti = zipfile.ZipInfo(target, date_time=infos[target].date_time)
            ti.compress_type = infos[target].compress_type
            zout.writestr(ti, xml.encode("utf-8"))
            for n in names:
                if n == target:
                    continue
                zi = zipfile.ZipInfo(n, date_time=infos[n].date_time)
                zi.compress_type = infos[n].compress_type
                zout.writestr(zi, others[n])
    except Exception as e:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise OoxmlError("打包失败: %s" % e)

    os.remove(filepath)
    os.rename(tmp_path, filepath)
    return count
