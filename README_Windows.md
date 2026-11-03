# 批量替换工具（Windows 版）使用与打包说明

与 macOS 版功能完全一致：**doc→docx / xls→xlsx 无损格式转换（Microsoft Word / Excel 引擎）+ OOXML 批量文本替换**。

---

## 一、环境要求

| 项目 | 要求 |
|---|---|
| 操作系统 | Windows 10 / 11（64 位） |
| Microsoft Office | 已安装 Word、Excel（2016 或更新版本；WPS 兼容性未验证） |
| Python | 仅打包/开发模式需要：Python 3.9+（64 位，安装时勾选 "Add to PATH"） |
| 内存 | 2 GB 以上（打包后运行占用很低） |

> 如果只想使用，**不需要 Python**：直接运行打包产出的 `dist\批量替换工具.exe`。
> 但 doc/xls 格式转换**必须安装 Microsoft Office**（Word/Excel）。

---

## 二、功能与行为（与 macOS 版一致）

- **格式转换**：`.doc → .docx`、`.xls → .xlsx`，由 Microsoft Word / Excel 引擎执行，**无损转换**（内容、格式、公式、样式均保留）。
- **批量替换**：对「转换后的 docx/xlsx」以及「本身已是 docx/xlsx 的文件」执行 OOXML 文本替换（docx 处理正文 `<w:t>`，xlsx 处理单元格字符串 sharedStrings.xml）。
- **未添加替换规则**：只做格式转换，不进入替换流程。
- **备份策略**：
  - 单独添加的文件 → 处理前在同目录生成 `文件名_bak.原扩展名`（如 `合同.docx → 合同_bak.docx`）
  - 直接添加的文件夹 → 处理前在同级生成 `文件夹名_bak`（整体备份，只备份一次）
  - 备份已存在 → 自动跳过，不覆盖
- **只保留两份**：`.doc / .xls` 转换替换成功后，确认备份存在即移除源文件，目录里只留 **备份文件 + 转换替换后的文件** 两份；关闭「处理前备份源文件」开关时源文件原样保留。
- **界面选项**：
  - 处理前备份源文件（默认开）
  - 忽略大小写（默认关）
  - 覆盖已存在目标文件（默认开）
  - 递归包含子文件夹（默认关）
  - 输出目录（默认：原文件所在目录）

---

## 三、打包（生成可分发 exe）

在 Windows 上双击 **`build.bat`**，自动完成：检查 Python → 安装依赖 → PyInstaller 打包。

产物：**`dist\批量替换工具.exe`**（单文件，约 60-80 MB，可拷到任何 Windows 电脑双击运行，对方无需装 Python）。

如需手动打包：

```bat
pip install -r requirements.txt
pyinstaller --noconfirm --clean --windowed --onefile --name "批量替换工具" --icon app_icon.ico main.py
```

> 打包时若杀毒软件拦截，请选择"允许"（PyInstaller 为常见打包工具）。

---

## 三-B、GitHub Actions 云端打包（无需 Windows 电脑）

没有 Windows 电脑时，用 GitHub 免费服务器自动打包（微软 Windows 环境）：

1. **注册 GitHub 账号**（免费）：https://github.com/signup
2. **新建仓库**：右上角 `+` → New repository → 起名（如 `batch-replace-tool`）→ Create repository
3. **上传本项目**：进入仓库 → `Add file` → `Upload files` → 把本文件夹内**所有内容**（包括 `.github` 文件夹、`main.py`、`app_icon.ico` 等）拖入 → Commit changes
   > 注意：文件要放在仓库根目录，`.github` 与 `main.py` 同级。
4. **等待自动打包**：仓库顶部 `Actions` 标签页可看到 `Build Windows exe` 运行，约 3–6 分钟完成
5. **下载 exe**：Actions 页 → 点击最新一次运行 → 底部 `Artifacts` → 下载 `BatchReplaceTool-windows.zip` → 解压即得 `批量替换工具.exe`

> 也可以安装 git 后用命令行推送：
> ```bat
> git init && git add . && git commit -m "init" && git branch -M main
> git remote add origin https://github.com/你的用户名/仓库名.git
> git push -u origin main
> ```
> 每次推送到 main 分支都会自动重新打包。

---

## 四、开发模式运行（不打包）

```bat
pip install -r requirements.txt
python main.py
```

---

## 五、文件清单

| 文件 | 说明 |
|---|---|
| `main.py` | 图形界面（PySide6），布局与 macOS 版一致 |
| `batch_runner.py` | 收集 / 备份 / 转换 / 替换 / 汇总主流程 |
| `office_converter.py` | Word / Excel COM 引擎无损转换（doc→docx、xls→xlsx） |
| `ooxml_replacer.py` | OOXML 批量文本替换（zip 重打包） |
| `app_icon.ico` | 应用图标（由 `转换.svg` 生成，7 个尺寸 16–256） |
| `requirements.txt` | 依赖清单 |
| `build.bat` | 本地 Windows 一键打包脚本 |
| `.github/workflows/build-windows.yml` | GitHub Actions 云端自动打包（无需 Windows 电脑） |
| `README_Windows.md` | 本文档 |

---

## 六、常见问题

**Q1：运行提示"未检测到 Microsoft Office"**
请确认已安装 Microsoft Word / Excel（不是 WPS）。安装后重启工具。

**Q2：转换报"文件被占用或无权限访问"**
目标文件正被 Word / Excel / WPS 打开，或文件为只读。关闭占用程序后重试。

**Q3：首次转换较慢**
Office 首次启动需要几秒到十几秒，属正常现象。

**Q4：杀毒软件提示 exe 有风险**
PyInstaller 打包的单文件程序常被误报。可添加信任，或用 `--onedir` 方式打包（build.bat 默认 onefile）。

**Q5：如何替换特殊字符（如 & < > 等）**
直接填写即可，程序会按 XML 规则转义写入，与 macOS 版行为一致。

---

## 七、与 macOS 版的差异说明

- 引擎调用方式不同：macOS 用 AppleScript 驱动 Office，Windows 用 COM（pywin32）驱动 Office，**均为 Microsoft 官方引擎、无损转换**。
- OOXML 替换与备份/删源逻辑逐行对齐 macOS 版，行为一致。
- Windows 版打包为单文件 exe，macOS 版为 .app bundle。
