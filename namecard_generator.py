"""
Namecard Generator
Generate multi-page Word namecards by duplicating a template and
replacing WordArt text while preserving layout and formatting.

UI languages: Chinese (--lang zh) and English (--lang en)

Requires: Windows + Microsoft Word + PyQt5 + pywin32
"""

from __future__ import annotations

import argparse
import os
import sys
import time
import traceback
from pathlib import Path

# ---------------------------------------------------------------------------
# Platform gate (Word COM automation is Windows-only)
# ---------------------------------------------------------------------------
if sys.platform != "win32":
    raise SystemExit("This application only runs on Windows (Word COM required).")

import pythoncom
import win32com.client as win32
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

# Word / Office constants
WD_STORY = 6          # wdStory
WD_PAGE_BREAK = 7     # wdPageBreak
MSO_TEXT_EFFECT = 15  # msoTextEffect (WordArt)

# =============================================================================
# UI strings (Chinese / English)
# =============================================================================
UI = {
    "zh": {
        "app_title": "名签生成器",
        "template": "模板文档:",
        "template_ph": "选择 Word 模板文档（.docx / .doc）",
        "output": "输出文档:",
        "output_ph": "设置生成结果的保存路径",
        "browse": "浏览...",
        "placeholder": "模板中的原始名称:",
        "placeholder_ph": "请输入模板 WordArt 中当前显示的完整文字",
        "names": "替换名称列表（每行一个）:",
        "names_ph": "张三\n李四\n王五",
        "import_txt": "导入 .txt",
        "ready": "就绪。",
        "generate": "开始生成",
        "stop": "停止",
        "open_folder": "打开输出目录",
        "tip": "提示：生成前请关闭正在打开该模板的 Word。本机需安装桌面版 Microsoft Word。",
        "dlg_template": "选择模板文档",
        "dlg_output": "设置输出文档",
        "dlg_import": "导入姓名列表",
        "filter_word": "Word 文档 (*.docx *.doc)",
        "filter_docx": "Word 文档 (*.docx)",
        "filter_txt": "文本文件 (*.txt);;所有文件 (*.*)",
        "warn_title": "警告",
        "warn_template": "请选择有效的模板文档。",
        "warn_output": "请设置输出文档路径。",
        "warn_placeholder": "请输入模板中的原始名称。",
        "warn_names": "请至少输入一个替换名称。",
        "warn_outdir": "输出目录不存在：\n{path}",
        "starting": "正在启动...",
        "stopping": "正在停止...",
        "saved": "已保存：{path}",
        "done_title": "完成",
        "done_msg": "生成成功：\n{path}\n\n是否立即打开该文件？",
        "fail": "失败。",
        "err_title": "错误",
        "err_body": "{short}\n\n详细信息：\n{details}",
        "info_title": "提示",
        "info_folder": "尚未设置输出路径，或目录不存在。",
        "confirm_title": "确认退出",
        "confirm_exit": "仍在生成中，确定要退出吗？",
        "yes": "是",
        "no": "否",
        # worker status
        "w_start_word": "正在启动 Microsoft Word...",
        "w_open_tpl": "正在打开模板...",
        "w_dup": "正在复制页面...",
        "w_dup_page": "正在复制第 {cur}/{total} 页...",
        "w_replace": "正在替换姓名...",
        "w_replace_one": "正在替换 {cur}/{total}：{name}",
        "w_save": "正在保存文档...",
        "w_done": "完成。",
        "w_cancel": "用户已取消。",
        "w_missing": "未找到 WordArt 占位符「{ph}」（处理姓名「{name}」时）。",
        "w_count": "WordArt 数量与姓名数量不一致：找到 {found} 个占位符，但有 {need} 个姓名。",
    },
    "en": {
        "app_title": "Namecard Generator",
        "template": "Template:",
        "template_ph": "Select a Word template (.docx / .doc)",
        "output": "Output:",
        "output_ph": "Choose where to save the result",
        "browse": "Browse...",
        "placeholder": "Placeholder name in template:",
        "placeholder_ph": "Exact WordArt text currently in the template",
        "names": "Names (one per line):",
        "names_ph": "Zhang San\nLi Si\nWang Wu",
        "import_txt": "Import .txt",
        "ready": "Ready.",
        "generate": "Generate",
        "stop": "Stop",
        "open_folder": "Open Output Folder",
        "tip": "Tip: Keep Microsoft Word closed for the template file. Requires desktop Word on this PC.",
        "dlg_template": "Select template",
        "dlg_output": "Save output as",
        "dlg_import": "Import name list",
        "filter_word": "Word Documents (*.docx *.doc)",
        "filter_docx": "Word Document (*.docx)",
        "filter_txt": "Text Files (*.txt);;All Files (*.*)",
        "warn_title": "Warning",
        "warn_template": "Please select a valid template file.",
        "warn_output": "Please set an output path.",
        "warn_placeholder": "Please enter the placeholder name.",
        "warn_names": "Please enter at least one name.",
        "warn_outdir": "Output folder does not exist:\n{path}",
        "starting": "Starting...",
        "stopping": "Stopping...",
        "saved": "Saved: {path}",
        "done_title": "Done",
        "done_msg": "Generated successfully:\n{path}\n\nOpen the file now?",
        "fail": "Failed.",
        "err_title": "Error",
        "err_body": "{short}\n\nDetails:\n{details}",
        "info_title": "Info",
        "info_folder": "Output folder is not set or missing.",
        "confirm_title": "Confirm",
        "confirm_exit": "Generation is still running. Exit anyway?",
        "yes": "Yes",
        "no": "No",
        "w_start_word": "Starting Microsoft Word...",
        "w_open_tpl": "Opening template...",
        "w_dup": "Duplicating pages...",
        "w_dup_page": "Duplicating page {cur}/{total}...",
        "w_replace": "Replacing names...",
        "w_replace_one": "Replacing {cur}/{total}: {name}",
        "w_save": "Saving document...",
        "w_done": "Done.",
        "w_cancel": "Cancelled by user.",
        "w_missing": "WordArt placeholder '{ph}' not found (while setting name '{name}').",
        "w_count": "WordArt count mismatch: found {found} placeholders, but {need} names.",
    },
}


def get_ui(lang: str) -> dict[str, str]:
    return UI["zh" if lang.startswith("zh") else "en"]


def app_dir() -> Path:
    """Return the directory that holds resources (dev or frozen exe)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resource_path(*parts: str) -> Path:
    """Resolve a resource path for both source and PyInstaller builds."""
    base = Path(getattr(sys, "_MEIPASS", app_dir()))
    return base.joinpath(*parts)


# =============================================================================
# Worker: Word automation in a background thread
# =============================================================================
class GenerateWorker(QThread):
    progress = pyqtSignal(int, str)          # percent, status text
    finished_ok = pyqtSignal(str)            # output path
    finished_err = pyqtSignal(str)           # error message

    def __init__(
        self,
        template: str,
        output: str,
        placeholder: str,
        names: list[str],
        ui: dict[str, str],
    ):
        super().__init__()
        self.template = template
        self.output = output
        self.placeholder = placeholder
        self.names = names
        self.ui = ui
        self._stop = False

    def stop(self) -> None:
        self._stop = True

    def run(self) -> None:
        word = None
        doc = None
        t = self.ui
        pythoncom.CoInitialize()
        try:
            self.progress.emit(5, t["w_start_word"])
            word = win32.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0

            self.progress.emit(10, t["w_open_tpl"])
            doc = word.Documents.Open(os.path.abspath(self.template))

            self.progress.emit(20, t["w_dup"])
            doc.Content.Select()
            selection = word.Selection
            selection.Copy()

            total = len(self.names)
            for i in range(total - 1):
                if self._stop:
                    raise RuntimeError(t["w_cancel"])
                selection.EndKey(Unit=WD_STORY)
                selection.InsertBreak(Type=WD_PAGE_BREAK)
                selection.Paste()
                pct = 20 + int(40 * (i + 1) / max(total - 1, 1))
                self.progress.emit(
                    pct,
                    t["w_dup_page"].format(cur=i + 2, total=total),
                )

            # Save & reopen so pasted pages settle before we touch WordArt.
            # Skipping this often shifts only the last page's name.
            out = os.path.abspath(self.output)
            self.progress.emit(60, t["w_save"])
            doc.SaveAs(out, FileFormat=16)  # 16 = wdFormatXMLDocument
            doc.Close(False)
            doc = None
            time.sleep(0.3)
            doc = word.Documents.Open(out)

            self.progress.emit(65, t["w_replace"])
            shapes = self._collect_wordarts(doc, self.placeholder)
            if len(shapes) != total:
                raise RuntimeError(
                    t["w_count"].format(found=len(shapes), need=total)
                )

            for i, (shape, name) in enumerate(zip(shapes, self.names)):
                if self._stop:
                    raise RuntimeError(t["w_cancel"])
                self._apply_wordart(shape, name)
                pct = 65 + int(30 * (i + 1) / total)
                self.progress.emit(
                    pct,
                    t["w_replace_one"].format(cur=i + 1, total=total, name=name),
                )

            # Let Word finish layout on the last WordArt before saving.
            time.sleep(0.25)
            try:
                doc.Repaginate()
            except Exception:
                pass

            self.progress.emit(95, t["w_save"])
            doc.Save()
            doc.Close(False)
            doc = None
            word.Quit()
            word = None

            self.progress.emit(100, t["w_done"])
            self.finished_ok.emit(out)

        except Exception as exc:
            self.finished_err.emit(f"{exc}\n\n{traceback.format_exc()}")
        finally:
            try:
                if doc is not None:
                    doc.Close(False)
            except Exception:
                pass
            try:
                if word is not None:
                    word.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()

    @staticmethod
    def _collect_wordarts(doc, find_text: str) -> list:
        """Collect placeholder WordArts in document order (not COM collection order)."""
        target = find_text.strip()
        found = []
        for shape in doc.Shapes:
            try:
                if shape.Type != MSO_TEXT_EFFECT:
                    continue
                if shape.TextEffect.Text.strip() != target:
                    continue
                # Anchor.Start = character offset; sorts pages correctly.
                anchor = int(shape.Anchor.Start)
                found.append((anchor, float(shape.Top), float(shape.Left), shape))
            except Exception:
                continue
        found.sort(key=lambda item: (item[0], item[1], item[2]))
        return [item[3] for item in found]

    @staticmethod
    def _apply_wordart(shape, replace_text: str) -> None:
        """Set WordArt text and re-apply geometry after Word's auto-resize."""
        left, top = float(shape.Left), float(shape.Top)
        width, height = float(shape.Width), float(shape.Height)
        effect = shape.TextEffect
        font_name = effect.FontName
        font_size = effect.FontSize
        font_bold = effect.FontBold
        font_italic = effect.FontItalic
        fill_rgb = shape.Fill.ForeColor.RGB
        rotation = shape.Rotation

        effect.Text = replace_text
        effect.FontName = font_name
        effect.FontSize = font_size
        effect.FontBold = font_bold
        effect.FontItalic = font_italic
        shape.Fill.ForeColor.RGB = fill_rgb
        shape.Rotation = rotation

        # Size first, then position. Setting size after Left/Top can nudge the shape.
        shape.Width = width
        shape.Height = height
        shape.Left = left
        shape.Top = top
        # Second pass: Word often shifts once more after the first restore.
        shape.Left = left
        shape.Top = top
        time.sleep(0.05)


# =============================================================================
# Main window
# =============================================================================
class MainWindow(QMainWindow):
    def __init__(self, lang: str = "zh"):
        super().__init__()
        self.lang = "zh" if lang.startswith("zh") else "en"
        self.t = get_ui(self.lang)

        self.setWindowTitle(self.t["app_title"])
        self.setMinimumSize(640, 560)
        self.resize(720, 640)

        icon = resource_path("assets", "app.ico")
        if icon.exists():
            self.setWindowIcon(QIcon(str(icon)))

        self.worker: GenerateWorker | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        t = self.t
        root = QWidget()
        layout = QVBoxLayout(root)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        # Language switch
        lang_row = QHBoxLayout()
        lang_row.addStretch()
        self.lang_btn = QPushButton("English" if self.lang == "zh" else "中文")
        self.lang_btn.setFixedWidth(88)
        self.lang_btn.clicked.connect(self._toggle_lang)
        lang_row.addWidget(self.lang_btn)
        layout.addLayout(lang_row)

        layout.addLayout(self._path_row(
            t["template"], t["template_ph"], t["browse"],
            self._pick_template, attr="template_edit",
        ))
        layout.addLayout(self._path_row(
            t["output"], t["output_ph"], t["browse"],
            self._pick_output, attr="output_edit",
        ))

        self.placeholder_label = QLabel(t["placeholder"])
        layout.addWidget(self.placeholder_label)
        self.placeholder_edit = QLineEdit()
        self.placeholder_edit.setPlaceholderText(t["placeholder_ph"])
        layout.addWidget(self.placeholder_edit)

        names_header = QHBoxLayout()
        self.names_label = QLabel(t["names"])
        names_header.addWidget(self.names_label)
        names_header.addStretch()
        self.import_btn = QPushButton(t["import_txt"])
        self.import_btn.clicked.connect(self._import_names)
        names_header.addWidget(self.import_btn)
        layout.addLayout(names_header)

        self.names_edit = QTextEdit()
        self.names_edit.setPlaceholderText(t["names_ph"])
        layout.addWidget(self.names_edit, stretch=1)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        layout.addWidget(self.progress)

        self.status_label = QLabel(t["ready"])
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        actions = QHBoxLayout()
        self.start_btn = QPushButton(t["generate"])
        self.start_btn.clicked.connect(self._start)
        actions.addWidget(self.start_btn)

        self.stop_btn = QPushButton(t["stop"])
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self._stop)
        actions.addWidget(self.stop_btn)

        self.open_btn = QPushButton(t["open_folder"])
        self.open_btn.clicked.connect(self._open_output_folder)
        actions.addWidget(self.open_btn)
        layout.addLayout(actions)

        self.tip_label = QLabel(t["tip"])
        self.tip_label.setStyleSheet("color: #666;")
        self.tip_label.setWordWrap(True)
        layout.addWidget(self.tip_label)

        self.setCentralWidget(root)

    def _toggle_lang(self) -> None:
        """Rebuild UI in the other language; keep user-entered values."""
        template = self.template_edit.text()
        output = self.output_edit.text()
        placeholder = self.placeholder_edit.text()
        names = self.names_edit.toPlainText()
        progress = self.progress.value()

        new_lang = "en" if self.lang == "zh" else "zh"
        self.lang = new_lang
        self.t = get_ui(new_lang)
        self._build_ui()

        self.template_edit.setText(template)
        self.output_edit.setText(output)
        self.placeholder_edit.setText(placeholder)
        self.names_edit.setPlainText(names)
        self.progress.setValue(progress)
        if self.worker and self.worker.isRunning():
            self._set_busy(True)

    def _path_row(self, label, placeholder, btn_text, slot, attr):
        row = QHBoxLayout()
        row.addWidget(QLabel(label))
        edit = QLineEdit()
        edit.setPlaceholderText(placeholder)
        setattr(self, attr, edit)
        row.addWidget(edit, stretch=1)
        btn = QPushButton(btn_text)
        btn.clicked.connect(slot)
        row.addWidget(btn)
        return row

    # ----- file pickers -----
    def _pick_template(self) -> None:
        t = self.t
        path, _ = QFileDialog.getOpenFileName(
            self, t["dlg_template"], "", t["filter_word"],
        )
        if not path:
            return
        self.template_edit.setText(path)
        if not self.output_edit.text().strip():
            p = Path(path)
            self.output_edit.setText(str(p.with_name(p.stem + "_output.docx")))

    def _pick_output(self) -> None:
        t = self.t
        path, _ = QFileDialog.getSaveFileName(
            self,
            t["dlg_output"],
            self.output_edit.text().strip() or "",
            t["filter_docx"],
        )
        if not path:
            return
        if not path.lower().endswith(".docx"):
            path += ".docx"
        self.output_edit.setText(path)

    def _import_names(self) -> None:
        t = self.t
        path, _ = QFileDialog.getOpenFileName(
            self, t["dlg_import"], "", t["filter_txt"],
        )
        if not path:
            return
        try:
            text = Path(path).read_text(encoding="utf-8-sig")
        except UnicodeDecodeError:
            text = Path(path).read_text(encoding="gbk", errors="ignore")
        self.names_edit.setPlainText(text.strip())

    # ----- generation -----
    def _parse_names(self) -> list[str]:
        return [
            line.strip()
            for line in self.names_edit.toPlainText().splitlines()
            if line.strip()
        ]

    def _validate(self) -> tuple[str, str, str, list[str]] | None:
        t = self.t
        template = self.template_edit.text().strip()
        output = self.output_edit.text().strip()
        placeholder = self.placeholder_edit.text().strip()
        names = self._parse_names()

        if not template or not os.path.isfile(template):
            QMessageBox.warning(self, t["warn_title"], t["warn_template"])
            return None
        if not output:
            QMessageBox.warning(self, t["warn_title"], t["warn_output"])
            return None
        if not placeholder:
            QMessageBox.warning(self, t["warn_title"], t["warn_placeholder"])
            return None
        if not names:
            QMessageBox.warning(self, t["warn_title"], t["warn_names"])
            return None

        out_dir = os.path.dirname(os.path.abspath(output))
        if out_dir and not os.path.isdir(out_dir):
            QMessageBox.warning(
                self, t["warn_title"], t["warn_outdir"].format(path=out_dir),
            )
            return None
        return template, output, placeholder, names

    def _set_busy(self, busy: bool) -> None:
        self.start_btn.setEnabled(not busy)
        self.stop_btn.setEnabled(busy)
        self.template_edit.setEnabled(not busy)
        self.output_edit.setEnabled(not busy)
        self.placeholder_edit.setEnabled(not busy)
        self.names_edit.setEnabled(not busy)
        self.lang_btn.setEnabled(not busy)

    def _start(self) -> None:
        data = self._validate()
        if not data:
            return
        template, output, placeholder, names = data

        self.progress.setValue(0)
        self.status_label.setText(self.t["starting"])
        self._set_busy(True)

        self.worker = GenerateWorker(
            template, output, placeholder, names, self.t,
        )
        self.worker.progress.connect(self._on_progress)
        self.worker.finished_ok.connect(self._on_ok)
        self.worker.finished_err.connect(self._on_err)
        self.worker.start()

    def _stop(self) -> None:
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.status_label.setText(self.t["stopping"])

    def _on_progress(self, value: int, text: str) -> None:
        self.progress.setValue(value)
        self.status_label.setText(text)

    def _on_ok(self, path: str) -> None:
        t = self.t
        self._set_busy(False)
        self.status_label.setText(t["saved"].format(path=path))
        reply = QMessageBox.question(
            self,
            t["done_title"],
            t["done_msg"].format(path=path),
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        if reply == QMessageBox.Yes:
            os.startfile(path)  # noqa: S606 - intentional on Windows

    def _on_err(self, message: str) -> None:
        t = self.t
        self._set_busy(False)
        self.status_label.setText(t["fail"])
        short = message.splitlines()[0] if message else "Unknown error"
        QMessageBox.critical(
            self,
            t["err_title"],
            t["err_body"].format(short=short, details=message[:2000]),
        )

    def _open_output_folder(self) -> None:
        t = self.t
        path = self.output_edit.text().strip()
        folder = os.path.dirname(os.path.abspath(path)) if path else ""
        if folder and os.path.isdir(folder):
            os.startfile(folder)  # noqa: S606
        else:
            QMessageBox.information(self, t["info_title"], t["info_folder"])

    def closeEvent(self, event) -> None:  # noqa: N802 - Qt API
        t = self.t
        if self.worker and self.worker.isRunning():
            reply = QMessageBox.question(
                self,
                t["confirm_title"],
                t["confirm_exit"],
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply != QMessageBox.Yes:
                event.ignore()
                return
            self.worker.stop()
            self.worker.wait(3000)
        event.accept()


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Namecard Generator")
    parser.add_argument(
        "--lang",
        choices=("zh", "en"),
        default="zh",
        help="UI language: zh (Chinese) or en (English). Default: zh",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    # High-DPI friendly on Win10/Win11
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    app.setApplicationName("Namecard Generator")
    app.setOrganizationName("NamecardGenerator")

    font = QFont("Microsoft YaHei UI")
    if not font.exactMatch():
        font = QFont("Segoe UI")
    font.setPointSize(11)
    app.setFont(font)

    window = MainWindow(lang=args.lang)
    window.show()
    return app.exec_()


if __name__ == "__main__":
    sys.exit(main())
