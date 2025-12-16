import sys
import os
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLineEdit,
    QPushButton,
    QFileDialog,
    QMessageBox,
    QScrollArea,
    QLabel,
)
from PySide6.QtCore import Qt, QPoint
from pypdf import PdfReader, PdfWriter


class PdfRow(QWidget):
    def __init__(self, app):
        super().__init__()
        self.app = app
        self._drag_start_pos: QPoint | None = None

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        # ===== 通し番号（ドラッグハンドル）=====
        self.index_label = QLabel("1")
        self.index_label.setFixedWidth(32)
        self.index_label.setAlignment(Qt.AlignCenter)
        self.index_label.setCursor(Qt.OpenHandCursor)

        # ===== パス入力 =====
        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText("PDF ファイルパス")
        self.path_edit.editingFinished.connect(self.update_pages)

        # ===== ページ数 =====
        self.page_label = QLabel("- p")
        self.page_label.setFixedWidth(50)
        self.page_label.setAlignment(Qt.AlignCenter)

        browse_btn = QPushButton("参照")
        browse_btn.clicked.connect(self.browse)

        up_btn = QPushButton("↑")
        up_btn.clicked.connect(lambda: self.app.move_row(self, -1))

        down_btn = QPushButton("↓")
        down_btn.clicked.connect(lambda: self.app.move_row(self, 1))

        delete_btn = QPushButton("削除")
        delete_btn.clicked.connect(lambda: self.app.delete_row(self))

        layout.addWidget(self.index_label)
        layout.addWidget(self.path_edit, stretch=1)
        layout.addWidget(self.page_label)
        layout.addWidget(browse_btn)
        layout.addWidget(up_btn)
        layout.addWidget(down_btn)
        layout.addWidget(delete_btn)

        # DnDイベント
        self.index_label.mousePressEvent = self.drag_start
        self.index_label.mouseMoveEvent = self.drag_move
        self.index_label.mouseReleaseEvent = self.drag_end

    # ---------- DnD ----------
    def drag_start(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_start_pos = event.globalPosition().toPoint()
            self.index_label.setCursor(Qt.ClosedHandCursor)

    def drag_move(self, event):
        if not self._drag_start_pos:
            return

        delta = event.globalPosition().toPoint() - self._drag_start_pos
        if abs(delta.y()) > 25:
            direction = -1 if delta.y() < 0 else 1
            self.app.move_row(self, direction)
            self._drag_start_pos = event.globalPosition().toPoint()

    def drag_end(self, event):
        self._drag_start_pos = None
        self.index_label.setCursor(Qt.OpenHandCursor)

    # ---------- PDF ----------
    def browse(self):
        path, _ = QFileDialog.getOpenFileName(self, "PDF選択", "", "PDF Files (*.pdf)")
        if path:
            self.path_edit.setText(path)
            self.update_pages()

    def update_pages(self):
        try:
            reader = PdfReader(self.path())
            self.page_label.setText(f"{len(reader.pages)} p")
        except Exception:
            self.page_label.setText("- p")
        self.app.update_summary()

    def path(self) -> str:
        return self.path_edit.text().strip()

    def set_index(self, index: int):
        self.index_label.setText(str(index))


class PdfMergerApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("PDF 統合ツール")
        self.resize(1000, 600)

        self.rows: list[PdfRow] = []

        outer = QVBoxLayout(self)
        outer.setContentsMargins(20, 20, 20, 20)
        outer.setSpacing(12)

        # ===== ガイダンス =====
        guide = QLabel(
            "PDFを追加し、番号をドラッグまたは ↑↓ ボタンで順序を変更してください。"
        )
        guide.setAlignment(Qt.AlignLeft)
        outer.addWidget(guide)

        # ===== スクロールエリア =====
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        self.rows_container = QWidget()
        self.rows_layout = QVBoxLayout(self.rows_container)
        self.rows_layout.setAlignment(Qt.AlignTop)
        self.rows_layout.setSpacing(6)

        scroll.setWidget(self.rows_container)
        outer.addWidget(scroll, stretch=1)

        # ===== PDF追加 =====
        add_btn = QPushButton("＋ PDF追加")
        add_btn.clicked.connect(self.add_row)
        outer.addWidget(add_btn)

        # ===== 集計表示 =====
        self.summary_label = QLabel("PDF数: 0 / 総ページ数: 0")
        outer.addWidget(self.summary_label)

        # ===== 出力名 =====
        self.output_name = QLineEdit("merged.pdf")
        self.output_name.setPlaceholderText("出力PDFファイル名")
        outer.addWidget(self.output_name)

        # ===== 統合ボタン =====
        merge_btn = QPushButton("PDFを統合")
        merge_btn.clicked.connect(self.merge_pdfs)
        outer.addWidget(merge_btn)

        self.setStyleSheet(
            """
            QWidget { font-size: 14px; }
            QLineEdit { padding: 6px; }
            QPushButton { padding: 8px; }
            QLabel { font-weight: bold; }
            """
        )

        self.add_row()

    # ---------- 行操作 ----------
    def add_row(self):
        row = PdfRow(self)
        self.rows.append(row)
        self.rebuild_rows()

    def delete_row(self, row: PdfRow):
        if row in self.rows:
            self.rows.remove(row)
            row.setParent(None)
            self.rebuild_rows()

    def move_row(self, row: PdfRow, direction: int):
        idx = self.rows.index(row)
        new_idx = idx + direction
        if 0 <= new_idx < len(self.rows):
            self.rows[idx], self.rows[new_idx] = self.rows[new_idx], self.rows[idx]
            self.rebuild_rows()

    def rebuild_rows(self):
        while self.rows_layout.count():
            item = self.rows_layout.takeAt(0)
            if item.widget():
                item.widget().setParent(None)

        for i, row in enumerate(self.rows, start=1):
            row.set_index(i)
            self.rows_layout.addWidget(row)

        self.update_summary()

    # ---------- 集計 ----------
    def update_summary(self):
        total_pages = 0
        for r in self.rows:
            try:
                reader = PdfReader(r.path())
                total_pages += len(reader.pages)
            except Exception:
                pass

        self.summary_label.setText(
            f"PDF数: {len(self.rows)} / 総ページ数: {total_pages}"
        )

    # ---------- PDF 統合 ----------
    def merge_pdfs(self):
        paths = [r.path() for r in self.rows if r.path()]
        if not paths:
            QMessageBox.critical(self, "エラー", "PDFが指定されていません")
            return

        output = self.output_name.text().strip()
        if not output.endswith(".pdf"):
            output += ".pdf"

        writer = PdfWriter()

        try:
            for path in paths:
                if not os.path.exists(path):
                    raise FileNotFoundError(path)
                reader = PdfReader(path)
                for page in reader.pages:
                    writer.add_page(page)

            with open(output, "wb") as f:
                writer.write(f)

            QMessageBox.information(self, "完了", f"生成完了:\n{output}")

        except Exception as e:
            QMessageBox.critical(self, "エラー", str(e))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PdfMergerApp()
    window.show()
    sys.exit(app.exec())
