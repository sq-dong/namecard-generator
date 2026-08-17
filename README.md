# Namecard Generator / 名签生成器

[English](#english) | [中文](#中文)

---

# English

Desktop app for **Windows 10 / Windows 11**. Given a Word namecard template that uses **WordArt**, it duplicates pages for each name in your list and replaces the WordArt text while preserving layout and styling as much as possible.

The app ships with **two UI languages** (Chinese and English) in one codebase, selected by launch argument or an in-window switch.

## Features

| Feature | Description |
|---------|-------------|
| Template-driven | Use your own Word namecard template (`.docx` / `.doc`) |
| Batch generation | Duplicate pages by name list (typically one person per page) |
| WordArt replacement | Find the placeholder WordArt and keep font, color, position, rotation, etc. |
| Import names | Load names from a `.txt` file (UTF-8 / GBK) |
| Progress & cancel | Runs in a background thread; shows progress and supports stop |
| Bilingual UI | Chinese / English; switch language inside the window |
| One-click package | Build a single `.exe` with PyInstaller |

## Requirements (important)

This tool uses **desktop Microsoft Word** COM automation (`pywin32`), so:

- **Windows only** (designed for Win10 / Win11)
- The PC must have **desktop Word** installed (Office Word; **Word Online is not supported**)
- Close the template in Word before generating to avoid file locks

For source runs you also need:

- Python **3.9+**
- Dependencies in `requirements.txt`: `PyQt5`, `pywin32`

## Quick start

### Option 1: Double-click (recommended)

| File | Description |
|------|-------------|
| `run.bat` | Menu to choose Chinese / English UI |
| `run_zh.bat` | Launch **Chinese UI** |
| `run_en.bat` | Launch **English UI** |

The launcher finds a working Python (project `.venv`, then `D:\myenv`, then PATH) and installs missing dependencies when needed.

### Option 2: Command line

```bat
python -m pip install -r requirements.txt

rem Chinese UI (default)
python namecard_generator.py
python namecard_generator.py --lang zh

rem English UI
python namecard_generator.py --lang en
```

### Option 3: Build a standalone APP (.exe)

```bat
build.bat
```

Output:

```text
dist\NamecardGenerator.exe
```

Copy the exe to other Win10/Win11 PCs (**desktop Word is still required** on the target machine).

```bat
NamecardGenerator.exe --lang en
```

Without arguments, the UI defaults to Chinese.

## How to use

1. **Prepare a template**  
   Design one namecard page in Word. Put the name in **WordArt** (e.g. `Zhang San`).  
   Plain text boxes / body paragraphs are **not** replaced by this tool.

2. **Start the app**  
   Use `run_en.bat` / `run_zh.bat`, or `run.bat` to pick a language.

3. **Choose template and output**  
   - Template: your namecard Word file  
   - Output: where to save the generated `.docx`  
   - If output is empty after picking a template, the app suggests `originalname_output.docx`

4. **Enter the placeholder**  
   Type the **exact** WordArt text currently shown in the template (including spaces).

5. **Enter or import names**  
   One name per line, or use **Import .txt**.

6. **Generate**  
   Click **Generate**, wait for progress. When done, open the file or the output folder.

## How it works

1. A worker thread starts a dedicated Word process (`DispatchEx`)
2. Opens the template, selects all, and copies
3. For names 2…N: insert a page break at the end and paste
4. For each name, replace the first WordArt still equal to the placeholder (`msoTextEffect`), restoring style and geometry
5. Save as `.docx` and quit Word

## Chinese vs English UI

There is one main program: `namecard_generator.py`.

UI strings live in the `UI` dict (`zh` / `en`):

```text
--lang zh   → Chinese UI
--lang en   → English UI
```

The top-right **English / 中文** button switches language at runtime (keeps paths and name list).

| Launch | UI language |
|--------|-------------|
| `run_zh.bat` | Chinese |
| `run_en.bat` | English |
| `run.bat` → 1 / 2 | Chinese / English |
| `python namecard_generator.py --lang en` | English |

## Project layout

```text
new_ncg/
├── namecard_generator.py   # Main app (UI strings + Word automation)
├── requirements.txt        # Python dependencies
├── run.bat                 # Language menu launcher
├── run_zh.bat              # Chinese UI shortcut
├── run_en.bat              # English UI shortcut
├── _env.bat                # Shared: find Python + repair/ensure pip
├── _launch.bat             # Shared: call _env, install deps, start app with --lang
├── build.bat               # Build exe (also uses _env.bat for pip)
├── assets/
│   ├── app.ico             # Window / exe icon
│   └── app.png             # Extra artwork
├── .gitignore
└── README.md               # This file
```

`_launch.bat` is an internal helper (not meant to be double-clicked daily). `run.bat` / `run_zh.bat` / `run_en.bat` call it so Python discovery, pip repair, dependency install, and `--lang` stay in one place.

Sample test `.docx` files from the old folder are **intentionally not included**. Use your own templates.

## Dependencies

```text
PyQt5>=5.15.0,<6    # GUI
pywin32>=306        # Word COM automation
```

```bat
python -m pip install -r requirements.txt
```

If you see `No module named pip`:

```bat
python -m ensurepip --default-pip --upgrade
```

Or use a working venv / Anaconda Python.

## FAQ

**Q: WordArt placeholder not found?**  
A: Ensure the name is **WordArt**, and the placeholder string matches the WordArt text **exactly**.

**Q: Generation fails / Word errors?**  
A: Confirm desktop Word is installed; close the template; end leftover `WINWORD.EXE` in Task Manager and retry.

**Q: macOS / Linux?**  
A: No. This tool needs Word COM on Windows.

**Q: Replace normal text instead of WordArt?**  
A: Not in this version (`Shape.Type == 15` only). That can be added later if needed.

**Q: Chinese UI looks wrong?**  
A: Prefers Microsoft YaHei UI, falls back to Segoe UI. Install Chinese fonts if missing.

## Development notes

- Word work runs in `QThread` so the UI stays responsive; COM uses `pythoncom.CoInitialize()` / `CoUninitialize()`
- Cleanup paths try to close the document and quit Word to avoid orphan processes
- Packaging adds `assets\app.ico`; runtime resolves files via `resource_path()`

## License / intended use

For office namecards / place cards. Follow your organization’s rules for event materials and personal data.

---

# 中文

面向 **Windows 10 / Windows 11** 的桌面小工具：根据一份带 **WordArt（艺术字）** 的 Word 名签模板，按姓名列表自动复制多页，并逐页替换姓名，同时尽量保持原有版式与样式。

提供 **中文 UI** 与 **英文 UI** 两个界面版本（同一套程序，启动参数不同）。

## 功能概览

| 功能 | 说明 |
|------|------|
| 模板驱动 | 使用你自己的 Word 名签模板（`.docx` / `.doc`） |
| 批量生成 | 按名单复制页面，一人一页（或按模板页结构扩展） |
| WordArt 替换 | 定位艺术字占位符，替换文字并保留字体、颜色、位置、旋转等 |
| 名单导入 | 支持从 `.txt` 导入（UTF-8 / GBK） |
| 进度与取消 | 后台线程处理，界面可显示进度并可中途停止 |
| 双语界面 | 中文版 / 英文版；窗口内也可一键切换语言 |
| 一键打包 | 可用 PyInstaller 打成单个 `.exe` 分发 |

## 运行环境（重要）

本工具依赖 **Microsoft Word 桌面版** 的 COM 自动化（`pywin32`），因此：

- 仅支持 **Windows**（已在 Win10 / Win11 场景设计）
- 目标电脑必须安装 **桌面版 Word**（Office 套件中的 Word；**不支持** 仅有 Word Online）
- 生成前请关闭正在打开该模板的 Word，避免文件被占用

开发 / 源码运行额外需要：

- Python **3.9+**
- 依赖见 `requirements.txt`：`PyQt5`、`pywin32`

## 快速开始

### 方式一：双击启动（推荐）

| 文件 | 说明 |
|------|------|
| `run.bat` | 启动时选择中文 / 英文界面 |
| `run_zh.bat` | 直接打开 **中文 UI** |
| `run_en.bat` | 直接打开 **英文 UI** |

脚本会自动寻找可用的 Python（优先项目 `.venv`、本机 `D:\myenv`，再回退到 PATH），缺少依赖时会尝试自动安装。

### 方式二：命令行

```bat
python -m pip install -r requirements.txt

rem 中文界面（默认）
python namecard_generator.py
python namecard_generator.py --lang zh

rem 英文界面
python namecard_generator.py --lang en
```

### 方式三：打包成独立 APP（.exe）

```bat
build.bat
```

成功后得到：

```text
dist\NamecardGenerator.exe
```

把该 exe 拷贝到其他 Win10/Win11 电脑即可使用（**目标机仍需安装桌面版 Word**）。

```bat
NamecardGenerator.exe --lang en
```

不传参数时默认中文界面。

## 使用步骤

1. **准备模板**  
   在 Word 中做好一页名签，姓名部分使用 **艺术字（WordArt）**，例如文字为 `张三`。  
   注意：普通文本框 / 正文段落 **不会** 被本工具识别为可替换对象。

2. **打开程序**  
   使用 `run_zh.bat` 或 `run_en.bat`（或 `run.bat` 选语言）。

3. **选择模板与输出路径**  
   - 模板文档：你的名签模板  
   - 输出文档：生成的 `.docx` 保存位置  
   - 选择模板后，若输出为空，会自动建议 `原文件名_output.docx`

4. **填写占位符**  
   输入模板 WordArt 里 **当前显示的完整文字**（须完全一致，含空格）。

5. **填写或导入姓名列表**  
   每行一个姓名；也可点击「导入 .txt」。

6. **开始生成**  
   点击「开始生成」，等待进度完成。成功后可选择直接打开结果文件，或「打开输出目录」。

## 工作原理（简要）

1. 后台线程启动独立的 Word 进程（`DispatchEx`）  
2. 打开模板，全选并复制当前内容  
3. 对名单中第 2～N 个人：文末插入分页符并粘贴，得到多页相同版式  
4. 依次查找第一个仍等于「占位符」的 WordArt（类型 `msoTextEffect`），替换为对应姓名，并写回字体、填充、位置等属性  
5. 保存为 `.docx` 后关闭 Word  

这样可以在批量替换时尽量不破坏原设计。

## 中英文两个版本如何区分

核心代码只有一份：`namecard_generator.py`。

界面文案集中在文件内的 `UI` 字典（`zh` / `en`），通过参数选择：

```text
--lang zh   → 中文 UI
--lang en   → 英文 UI
```

窗口右上角还有 **English / 中文** 按钮，可在运行中切换语言（会保留已填写的路径与名单）。

| 启动方式 | 界面语言 |
|----------|----------|
| `run_zh.bat` | 中文 |
| `run_en.bat` | 英文 |
| `run.bat` → 选 1 / 2 | 中文 / 英文 |
| `python namecard_generator.py --lang en` | 英文 |

## 项目结构

```text
new_ncg/
├── namecard_generator.py   # 主程序（含中英文字符串与 Word 自动化）
├── requirements.txt        # Python 依赖
├── run.bat                 # 启动菜单（选中/英）
├── run_zh.bat              # 中文 UI 快捷启动
├── run_en.bat              # 英文 UI 快捷启动
├── _env.bat                # 公共：查找 Python + 修复/确保 pip
├── _launch.bat             # 公共：调用 _env、装依赖、带 --lang 启动程序
├── build.bat               # 一键打包 exe（同样走 _env.bat，避免 pip 问题）
├── assets/
│   ├── app.ico             # 窗口 / exe 图标
│   └── app.png             # 附加素材图
├── .gitignore
└── README.md               # 本说明
```

`_launch.bat` 是内部辅助脚本，一般不用天天双击。`run.bat` / `run_zh.bat` / `run_en.bat` 会调用它，这样「找 Python、修 pip、装依赖、传语言参数」只维护一份。

说明：旧测试用的 `.docx` 样例 **故意不放入** 本目录，请使用你自己的模板。

## 依赖说明

```text
PyQt5>=5.15.0,<6    # 图形界面
pywin32>=306        # Word COM 自动化
```

```bat
python -m pip install -r requirements.txt
```

若本机 `python -m pip` 报 `No module named pip`，可先执行：

```bat
python -m ensurepip --default-pip --upgrade
```

或改用已配置好的虚拟环境 / Anaconda 中的 Python。

## 常见问题

**Q: 提示找不到 WordArt 占位符？**  
A: 确认模板里姓名是 **艺术字**，且「模板中的原始名称」与艺术字文字 **完全一致**。

**Q: 生成失败 / Word 相关报错？**  
A: 确认已安装桌面版 Word；关闭模板文件后再试；任务管理器中结束残留的 `WINWORD.EXE` 后重试。

**Q: 能在 macOS / Linux 上用吗？**  
A: 不能。本工具依赖 Windows 上的 Word COM。

**Q: 能否替换普通文字而不是 WordArt？**  
A: 当前版本只处理 WordArt（`Shape.Type == 15`）。如需扩展普通文本替换，可以在此基础上继续开发。

**Q: 中文界面乱码或字体难看？**  
A: 程序优先使用「微软雅黑 UI」，找不到则回退到 Segoe UI。请确保系统含中文字体。

## 开发备注

- Word 操作放在 `QThread` 中，避免界面卡死；COM 线程内使用 `pythoncom.CoInitialize()` / `CoUninitialize()`  
- 异常路径会尽量关闭文档并退出 Word，减少残留进程  
- 打包时通过 `--add-data` 带上 `assets\app.ico`，运行时用 `resource_path()` 解析资源路径  

## 许可证与用途

本项目用于办公场景下的名签 / 席卡批量生成。请遵守你所在单位对会议材料与个人信息的使用规范。
