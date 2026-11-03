# -*- coding: utf-8 -*-
"""
批量执行：收集文件 → 备份 → 转换 → 替换 → 汇总
与 macOS 版（Swift BatchRunner）逻辑完全一致：
  - 收集 .doc/.docx/.xls/.xlsx（文件夹可选递归）
  - 备份：单独文件 → 同级「文件名_bak.原扩展名」；用户添加的文件夹 → 同级「文件夹名_bak」
          （已存在则跳过；备份内容为处理前原文）
  - 转换：doc→docx、xls→xlsx（Microsoft 引擎）；docx/xlsx 不转换
  - 替换：有规则则对转换产物/原 docx/xlsx 做 OOXML 替换；无规则仅转换
  - 删源：doc/xls 转换替换成功且对应备份存在时，移除源文件，只留「备份 + 产物」两份
"""
import os
import shutil

from ooxml_replacer import replace_text
from office_converter import (
    convert_doc_to_docx,
    convert_xls_to_xlsx,
    OfficeError,
)

SUPPORTED_EXTS = {".doc", ".docx", ".xls", ".xlsx"}


class JobConfig:
    def __init__(
        self,
        sources,
        recursive=False,
        backup_before_convert=True,
        ignore_case=False,
        overwrite_existing=True,
        output_dir=None,
        rules=None,
    ):
        self.sources = [os.path.abspath(s) for s in sources]
        self.recursive = recursive
        self.backup_before_convert = backup_before_convert
        self.ignore_case = ignore_case
        self.overwrite_existing = overwrite_existing
        self.output_dir = os.path.abspath(output_dir) if output_dir else None
        self.rules = [(r["find"], r["replace"]) for r in (rules or [])]


class JobSummary:
    def __init__(self, success=0, failed=0, skipped=0, replaced_count=0,
                 status_line="完成", detail=""):
        self.success = success
        self.failed = failed
        self.skipped = skipped
        self.replaced_count = replaced_count
        self.status_line = status_line
        self.detail = detail


def _is_dir(p):
    return os.path.isdir(p)


def _norm(p):
    return os.path.normcase(os.path.abspath(p))


def _is_under_folder(file_path, folder_path):
    """判断文件是否位于某文件夹内（Windows 路径大小写不敏感）。"""
    f = _norm(file_path)
    d = _norm(folder_path)
    return f.startswith(d + os.sep)


class BatchRunner:
    def __init__(self, config: JobConfig):
        self.config = config
        self.cancelled = False

    def cancel(self):
        self.cancelled = True

    # MARK: - 入口

    def run(self, on_log=None, on_progress=None, on_finish=None):
        on_log = on_log or (lambda text, level="info": None)
        on_progress = on_progress or (lambda p, s: None)
        on_finish = on_finish or (lambda s: None)

        files = self._collect_files()
        if not files:
            on_log("未找到可处理的文件。", "warning")
            on_finish(JobSummary(detail="未找到文件"))
            return

        outdir = self.config.output_dir or os.path.dirname(files[0])
        try:
            os.makedirs(outdir, exist_ok=True)
        except OSError as e:
            on_log("无法创建输出目录：%s" % e, "error")
            on_finish(JobSummary(failed=len(files), status_line="完成(有错误)",
                                 detail="输出目录创建失败"))
            return

        on_log("开始处理 %d 个文件…" % len(files), "info")
        on_log("输出到：%s" % outdir, "info")

        active_rules = [r for r in self.config.rules if r[0]]
        if active_rules:
            mode = "忽略大小写" if self.config.ignore_case else "区分大小写"
            on_log("已启用 %d 条替换规则，%s" % (len(active_rules), mode), "info")
        else:
            on_log("未添加替换规则，仅执行格式转换", "info")
        if self.config.backup_before_convert:
            on_log("已开启「处理前备份源文件」", "info")

        # 文件夹备份（用户直接添加的文件夹 → 同级「文件夹名_bak」，只备份一次）
        if self.config.backup_before_convert:
            self._backup_folders(on_log)

        success = failed = skipped = replaced_total = 0
        total = len(files)

        for index, file in enumerate(files):
            if self.cancelled:
                on_log("已停止（用户取消）。", "warning")
                break

            on_progress(index / total, "处理中 %d/%d：%s" % (index + 1, total, os.path.basename(file)))

            try:
                result, replaced = self._process_file(file, outdir, index + 1, on_log)
                if result == "success":
                    success += 1
                    replaced_total += replaced
                    msg = "（替换 %d 处）" % replaced if replaced > 0 else ""
                    on_log("✓ %s%s" % (os.path.basename(file), msg), "success")
                else:  # skipped
                    skipped += 1
                    on_log("  [跳过] 目标已存在：%s" % os.path.basename(file), "warning")
            except Exception as e:
                failed += 1
                on_log("× %s：%s" % (os.path.basename(file), e), "error")

        on_progress(1.0, "完成")

        detail = "成功 %d，失败 %d" % (success, failed)
        if skipped:
            detail += "，跳过 %d" % skipped
        if active_rules:
            detail += "，共替换 %d 处" % replaced_total

        summary = JobSummary(
            success=success, failed=failed, skipped=skipped,
            replaced_count=replaced_total,
            status_line="完成(有错误)" if failed else "完成",
            detail=detail,
        )
        on_log(detail, "warning" if failed else "success")
        on_finish(summary)

    # MARK: - 文件收集

    def _collect_files(self):
        result = []
        for source in self.config.sources:
            if not os.path.exists(source):
                continue
            if _is_dir(source):
                if self.config.recursive:
                    for root, dirs, files in os.walk(source):
                        dirs.sort()
                        for name in sorted(files):
                            if os.path.splitext(name)[1].lower() in SUPPORTED_EXTS:
                                result.append(os.path.join(root, name))
                else:
                    for name in sorted(os.listdir(source)):
                        p = os.path.join(source, name)
                        if os.path.isfile(p) and os.path.splitext(name)[1].lower() in SUPPORTED_EXTS:
                            result.append(p)
            else:
                if os.path.splitext(source)[1].lower() in SUPPORTED_EXTS:
                    result.append(source)
        # 去重 + 排序（与 macOS 版一致）
        return sorted(set(result), key=lambda p: _norm(p))

    # MARK: - 文件夹备份

    def _backup_folders(self, on_log):
        for source in self.config.sources:
            if not _is_dir(source):
                continue
            bak = os.path.join(
                os.path.dirname(source), os.path.basename(source) + "_bak"
            )
            if os.path.exists(bak):
                on_log("备份已存在，跳过：%s" % os.path.basename(bak), "info")
                continue
            try:
                shutil.copytree(source, bak)
                on_log("已备份文件夹：%s → %s" % (os.path.basename(source), os.path.basename(bak)), "success")
            except Exception as e:
                on_log("文件夹备份失败：%s（%s）" % (os.path.basename(source), e), "warning")

    def _is_under_folder_source(self, file_path):
        for source in self.config.sources:
            if _is_dir(source) and _is_under_folder(file_path, source):
                return True
        return False

    def _backup_path(self, file_path):
        """文件备份路径：同级 文件名_bak.原扩展名"""
        base, ext = os.path.splitext(file_path)
        return base + "_bak" + ext

    def _has_backup(self, file_path):
        """是否有对应备份：单独文件→文件名_bak.扩展名；文件夹内文件→所属文件夹的 文件夹名_bak"""
        for source in self.config.sources:
            if _is_dir(source) and _is_under_folder(file_path, source):
                bak = os.path.join(
                    os.path.dirname(source), os.path.basename(source) + "_bak"
                )
                return os.path.exists(bak)
        return os.path.exists(self._backup_path(file_path))

    # MARK: - 单文件处理

    def _process_file(self, file_path, outdir, seq, on_log):
        ext = os.path.splitext(file_path)[1].lower()
        active_rules = [r for r in self.config.rules if r[0]]

        # 备份（处理前）：单独添加的文件 → 文件名_bak.原扩展名
        if self.config.backup_before_convert and not self._is_under_folder_source(file_path):
            bpath = self._backup_path(file_path)
            if os.path.exists(bpath):
                on_log("备份已存在，跳过：%s" % os.path.basename(bpath))
            else:
                shutil.copy2(file_path, bpath)
                on_log("已备份文件：%s → %s" % (os.path.basename(file_path), os.path.basename(bpath)))

        # 确定目标文件名
        base, _ = os.path.splitext(os.path.basename(file_path))
        if ext == ".doc":
            target = os.path.join(outdir, base + ".docx")
        elif ext == ".xls":
            target = os.path.join(outdir, base + ".xlsx")
        else:
            target = os.path.join(outdir, os.path.basename(file_path))

        # 检查是否已存在
        if os.path.exists(target) and not self.config.overwrite_existing:
            return "skipped", 0

        # 老格式先转换（Microsoft 引擎）
        if ext == ".doc":
            convert_doc_to_docx(file_path, target)
        elif ext == ".xls":
            convert_xls_to_xlsx(file_path, target)
        elif os.path.abspath(target) != os.path.abspath(file_path):
            # docx/xlsx 复制到输出目录（输出目录与源目录相同时不复制）
            shutil.copy2(file_path, target)

        # 执行替换（无规则时仅格式转换）
        replaced = 0
        if active_rules:
            replaced = replace_text(target, active_rules, self.config.ignore_case)

        # doc/xls 转换+替换成功后移除源文件（内容已由备份保留），只留「备份+产物」两份
        if (ext == ".doc" or ext == ".xls") and self.config.backup_before_convert and self._has_backup(file_path):
            try:
                os.remove(file_path)
                on_log("已移除源文件（备份已保留）：%s" % os.path.basename(file_path))
            except OSError:
                on_log("源文件移除失败（可能被占用）：%s" % os.path.basename(file_path), "warning")

        return "success", replaced
