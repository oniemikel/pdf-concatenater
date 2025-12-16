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
    QFrame,
)
from PySide6.QtCore import Qt, QPoint
from pypdf import PdfReader, PdfWriter


class DropIndicator(QFrame):
    def __init__(self, parent):
        super().__init__(parent)
        self.setFixedHeight(3)
        self.setStyleSheet("background-color: #0078d7;")
        self.hide()


class PdfRow(QWidget):
    def __init__(self, app):
        super().__init__()
        self.app = app
        self._drag_start_pos: QPoint | None = None

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        self.index_label = QLabel("1")
        self.index_label.setFixedWidth(32)
        self.index_label.setAlignment(Qt.AlignCenter)
        self.index_label.setCursor(Qt.OpenHandCursor)

        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText("PDF ファイルパス")
        self.path_edit.editingFinished.connect(self.update_pages)

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

        self.index_label.mousePressEvent = self.drag_start
        self.index_label.mouseMoveEvent = self.drag_move
        self.index_label.mouseReleaseEvent = self.drag_end

    # ---------- DnD ----------
    def drag_start(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_start_pos = event.globalPosition().toPoint()
            self.setWindowOpacity(0.5)
            self.index_label.setCursor(Qt.ClosedHandCursor)

    def drag_move(self, event):
        if not self._drag_start_pos:
            return

        self.app.update_drop_indicator(event.globalPosition().toPoint())

    def drag_end(self, event):
        self.setWindowOpacity(1.0)
        self.index_label.setCursor(Qt.OpenHandCursor)
        self._drag_start_pos = None
        self.app.apply_drop(self)

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
        self.drop_target_index: int | None = None

        outer = QVBoxLayout(self)
        outer.setContentsMargins(20, 20, 20, 20)
        outer.setSpacing(12)

        guide = QLabel(
            "PDFを追加し、番号をドラッグまたは ↑↓ ボタンで順序を変更してください。"
        )
        outer.addWidget(guide)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        self.rows_container = QWidget()
        self.rows_layout = QVBoxLayout(self.rows_container)
        self.rows_layout.setAlignment(Qt.AlignTop)
        self.rows_layout.setSpacing(6)

        self.drop_indicator = DropIndicator(self.rows_container)
        self.rows_layout.addWidget(self.drop_indicator)

        scroll.setWidget(self.rows_container)
        outer.addWidget(scroll, stretch=1)

        add_btn = QPushButton("＋ PDF追加")
        add_btn.clicked.connect(self.add_row)
        outer.addWidget(add_btn)

        self.summary_label = QLabel("PDF数: 0 / 総ページ数: 0")
        outer.addWidget(self.summary_label)

        # ===== 出力先 =====
        out_dir_layout = QHBoxLayout()
        self.output_dir = QLineEdit(os.getcwd())
        browse_dir = QPushButton("参照")
        browse_dir.clicked.connect(self.select_output_dir)
        out_dir_layout.addWidget(self.output_dir, stretch=1)
        out_dir_layout.addWidget(browse_dir)
        outer.addLayout(out_dir_layout)

        self.output_name = QLineEdit("merged.pdf")
        outer.addWidget(self.output_name)

        merge_btn = QPushButton("PDFを統合")
        merge_btn.clicked.connect(self.merge_pdfs)
        outer.addWidget(merge_btn)

        self.add_row()

    # ---------- 行操作 ----------
    def add_row(self):
        row = PdfRow(self)
        self.rows.append(row)
        self.rebuild_rows()

    def delete_row(self, row):
        if row in self.rows:
            self.rows.remove(row)
            row.setParent(None)
            self.rebuild_rows()

    def move_row(self, row, direction):
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

        for i, row in enumerate(self.rows):
            row.set_index(i + 1)
            self.rows_layout.addWidget(row)

        self.rows_layout.addWidget(self.drop_indicator)
        self.update_summary()

    # ---------- DnD 補助 ----------
    def update_drop_indicator(self, global_pos: QPoint):
        y = self.rows_container.mapFromGlobal(global_pos).y()
        index = 0

        for i, row in enumerate(self.rows):
            if y < row.y() + row.height() // 2:
                index = i
                break
        else:
            index = len(self.rows)

        self.drop_target_index = index
        self.rows_layout.insertWidget(index, self.drop_indicator)
        self.drop_indicator.show()

    def apply_drop(self, row: PdfRow):
        self.drop_indicator.hide()
        if self.drop_target_index is None:
            return

        old = self.rows.index(row)
        new = self.drop_target_index

        if new > old:
            new -= 1

        self.rows.pop(old)
        self.rows.insert(new, row)
        self.drop_target_index = None
        self.rebuild_rows()

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

    # ---------- 出力 ----------
    def select_output_dir(self):
        d = QFileDialog.getExistingDirectory(self, "出力先ディレクトリ選択")
        if d:
            self.output_dir.setText(d)

    def merge_pdfs(self):
        paths = [r.path() for r in self.rows if r.path()]
        if not paths:
            QMessageBox.critical(self, "エラー", "PDFが指定されていません")
            return

        out_dir = self.output_dir.text().strip()
        name = self.output_name.text().strip()
        if not name.endswith(".pdf"):
            name += ".pdf"

        output = os.path.join(out_dir, name)

        writer = PdfWriter()

        try:
            for path in paths:
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
