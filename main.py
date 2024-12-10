# 项目初始化及依赖安装
# 1. 创建虚拟环境（可选但推荐）
# python3 -m venv venv
# source venv/bin/activate  # Linux/macOS
# venv\Scripts\activate  # Windows

# 2. 安装依赖
# pip install PyQt6 PyMuPDF

import os
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QTableWidget, QTableWidgetItem, QLabel, QLineEdit, QHBoxLayout, QMessageBox, QToolTip, QHeaderView, QStyledItemDelegate
from PyQt6.QtCore import Qt, QDateTime
from PyQt6.QtGui import QColor, QFont, QPen, QPainter
import fitz  # PyMuPDF

class BorderDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index):
        super().paint(painter, option, index)
        if index.column() == 3:  # 检查是否是“分割页数”列
            painter.save()
            pen = QPen(Qt.GlobalColor.black, 1)
            painter.setPen(pen)
            painter.drawLine(option.rect.bottomLeft(), option.rect.bottomRight())
            painter.restore()

class PDFSplitterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.output_folder_label = QLabel("分割后文件夹: ")
        self.output_folder_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.initUI()
        self.pdf_files = []

    def initUI(self):
        self.setWindowTitle('PDF 分割器')
        self.setGeometry(100, 100, 800, 600)

        main_layout = QVBoxLayout()

        # 添加分割前文件列表的说明和浏览文件按钮
        before_split_layout = QHBoxLayout()

        before_split_label = QLabel('待分割文件列表')
        font = QFont()
        font.setBold(True)
        before_split_label.setFont(font)

        self.browse_button = QPushButton('添加PDF文件')
        self.browse_button.setFixedWidth(100)
        self.browse_button.clicked.connect(self.browse_folder)
        self.browse_button.setStyleSheet("background-color: #4CAF50; color: white; border: none; padding: 5px;")

        # 添加清除按钮
        self.clear_button = QPushButton('清除文件列表')
        self.clear_button.setFixedWidth(100)
        self.clear_button.clicked.connect(self.clear_file_list)
        self.clear_button.setStyleSheet("background-color: #FF4500; color: white; border: none; padding: 5px;")

        before_split_layout.addWidget(before_split_label)
        before_split_layout.addWidget(self.browse_button)
        before_split_layout.addWidget(self.clear_button)

        main_layout.addLayout(before_split_layout)

        self.file_table_widget = QTableWidget()
        self.file_table_widget.setColumnCount(5)
        self.file_table_widget.setHorizontalHeaderLabels(['文件名', '大小 (KB)', '页数', '分割后文件页数', '文件路径'])
        self.file_table_widget.horizontalHeader().setStretchLastSection(True)

        header = self.file_table_widget.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.file_table_widget.setStyleSheet("""
            QHeaderView::section { background-color: #f0f0f0; }
            QTableWidget { border: 1px solid black; }
        """)

        main_layout.addWidget(self.file_table_widget)

        # 添加分割后文件列表的说明
        after_split_label = QLabel('分割后文件列表')
        after_split_label.setFont(font)
        main_layout.addWidget(after_split_label)

        self.split_result_table_widget = QTableWidget()
        self.split_result_table_widget.setColumnCount(4)
        self.split_result_table_widget.setHorizontalHeaderLabels(['文件名', '大小 (KB)', '页数', '分割前文件路径'])
        self.split_result_table_widget.horizontalHeader().setStretchLastSection(True)

        header = self.split_result_table_widget.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.split_result_table_widget.setStyleSheet("""
            QHeaderView::section { background-color: #f0f0f0; }
            QTableWidget { border: 1px solid black; }
        """)

        main_layout.addWidget(self.split_result_table_widget)

        # 添加按钮布局
        btn_layout = QHBoxLayout()

        button_width = 100

        self.split_button = QPushButton('分割 PDF')
        self.split_button.setFixedWidth(button_width)
        self.split_button.clicked.connect(self.split_pdfs)
        self.split_button.setStyleSheet("background-color: #523456; color: white; border: none; padding: 5px;")

        btn_layout.addWidget(self.output_folder_label)
        btn_layout.addWidget(self.split_button)

        main_layout.addLayout(btn_layout)

        self.setLayout(main_layout)

        self.border_delegate = BorderDelegate(self)
        self.file_table_widget.setItemDelegateForColumn(3, self.border_delegate)

    def clear_file_list(self):
        self.file_table_widget.setRowCount(0)
        self.output_folder_label.setText("分割后文件夹: ")

    def browse_folder(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self, "选择 PDF 文件", "", "PDF Files (*.pdf)")
        if file_paths:
            for file_path in file_paths:
                if not self.is_file_already_loaded(file_path):
                    self.load_pdf_file(file_path)

    def is_file_already_loaded(self, file_path):
        for row in range(self.file_table_widget.rowCount()):
            existing_file_path = self.file_table_widget.item(row, 4).text()
            if existing_file_path == file_path:
                return True
        return False

    def load_pdf_file(self, file_path):
        pdf_info = self.get_pdf_info(file_path)
        row_position = self.file_table_widget.rowCount()
        self.file_table_widget.insertRow(row_position)

        file_name_item = QTableWidgetItem(os.path.basename(file_path))
        file_name_item.setFlags(file_name_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        self.file_table_widget.setItem(row_position, 0, file_name_item)

        size_item = QTableWidgetItem(str(pdf_info['size']))
        size_item.setFlags(size_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        self.file_table_widget.setItem(row_position, 1, size_item)

        pages_item = QTableWidgetItem(str(pdf_info['pages']))
        pages_item.setFlags(pages_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        self.file_table_widget.setItem(row_position, 2, pages_item)

        split_pages_item = QTableWidgetItem(str(pdf_info['pages']))
        split_pages_item.setFlags(split_pages_item.flags() | Qt.ItemFlag.ItemIsEditable)
        split_pages_item.setBackground(QColor(255, 255, 192))
        split_pages_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        split_pages_item.setToolTip('双击编辑分割页数')
        self.file_table_widget.setItem(row_position, 3, split_pages_item)

        file_path_item = QTableWidgetItem(file_path)
        file_path_item.setFlags(file_path_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        self.file_table_widget.setItem(row_position, 4, file_path_item)

    def get_pdf_info(self, file_path):
        doc = fitz.open(file_path)
        size_kb = os.path.getsize(file_path) / 1024
        return {'size': round(size_kb, 2), 'pages': len(doc)}

    def get_valid_split_value(self, row):
        try:
            split_count = int(self.file_table_widget.item(row, 3).text())
            if split_count <= 0:
                raise ValueError("分割页数必须是正整数。")
            return split_count
        except ValueError:
            print("输入无效。使用默认分割页数 20。")
            return 20  # 默认分割页数

    def split_pdfs(self):
        if self.file_table_widget.rowCount() == 0:
            QMessageBox.warning(self, "警告", "请先添加要分割的 PDF 文件。")
            return

        folder_path = QFileDialog.getExistingDirectory(self, "选择保存分割文件的目录")
        if not folder_path:
            return

        current_time = QDateTime.currentDateTime().toString("yyyyMMdd_hhmmss")
        output_folder = os.path.join(folder_path, current_time)
        os.makedirs(output_folder, exist_ok=True)

        self.split_result_table_widget.clearContents()
        self.split_result_table_widget.setRowCount(0)

        self.output_folder_label.setText(f"分割后文件夹: {output_folder}")

        for row in range(self.file_table_widget.rowCount()):
            file_path = self.file_table_widget.item(row, 4).text()
            split_count = self.get_valid_split_value(row)
            pages = int(self.file_table_widget.item(row, 2).text())

            if split_count >= pages:
                QMessageBox.information(self, "提示", f"{os.path.basename(file_path)} 的分割页数大于或等于总页数，不进行分割。")
                self.add_to_split_result_table(file_path, "未分割")
                continue

            split_files = self.split_pdf(file_path, split_count, output_folder)
            self.add_to_split_result_table(file_path, split_files)

    def add_to_split_result_table(self, original_file_path, status):
        if status == "未分割":
            file_info = self.get_pdf_info(original_file_path)
            row_position = self.split_result_table_widget.rowCount()
            self.split_result_table_widget.insertRow(row_position)

            file_name_item = QTableWidgetItem(os.path.basename(original_file_path))
            file_name_item.setFlags(file_name_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.split_result_table_widget.setItem(row_position, 0, file_name_item)

            size_item = QTableWidgetItem(str(file_info['size']))
            size_item.setFlags(size_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.split_result_table_widget.setItem(row_position, 1, size_item)

            pages_item = QTableWidgetItem(str(file_info['pages']))
            pages_item.setFlags(pages_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.split_result_table_widget.setItem(row_position, 2, pages_item)

            original_file_path_item = QTableWidgetItem(original_file_path)
            original_file_path_item.setFlags(original_file_path_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.split_result_table_widget.setItem(row_position, 3, original_file_path_item)
        else:
            for split_file_path in status:
                file_info = self.get_pdf_info(split_file_path)
                row_position = self.split_result_table_widget.rowCount()
                self.split_result_table_widget.insertRow(row_position)

                file_name_item = QTableWidgetItem(os.path.basename(split_file_path))
                file_name_item.setFlags(file_name_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.split_result_table_widget.setItem(row_position, 0, file_name_item)

                size_item = QTableWidgetItem(str(file_info['size']))
                size_item.setFlags(size_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.split_result_table_widget.setItem(row_position, 1, size_item)

                pages_item = QTableWidgetItem(str(file_info['pages']))
                pages_item.setFlags(pages_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.split_result_table_widget.setItem(row_position, 2, pages_item)

                original_file_path_item = QTableWidgetItem(original_file_path)
                original_file_path_item.setFlags(original_file_path_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.split_result_table_widget.setItem(row_position, 3, original_file_path_item)

    def split_pdf(self, file_path, split_count, output_folder):
        doc = fitz.open(file_path)
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        split_files = []

        for start_page in range(0, len(doc), split_count):
            end_page = min(start_page + split_count, len(doc))
            output_file = f"{base_name}_part_{start_page // split_count + 1}.pdf"
            output_file_path = os.path.join(output_folder, output_file)
            output_doc = fitz.open()
            for page_num in range(start_page, end_page):
                output_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
            output_doc.save(output_file_path)
            output_doc.close()
            print(f"创建: {output_file_path}")
            split_files.append(output_file_path)

        doc.close()
        return split_files

if __name__ == '__main__':
    app = QApplication([])
    ex = PDFSplitterApp()
    ex.show()
    app.exec()