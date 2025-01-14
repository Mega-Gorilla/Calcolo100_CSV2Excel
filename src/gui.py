import sys
import pandas as pd
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QPushButton, QLabel, QFileDialog, 
                           QTableWidget, QTableWidgetItem, QLineEdit, 
                           QTextEdit, QMessageBox, QTableView)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QStandardItemModel, QStandardItem
import os
from .timecard_processor import TimecardProcessor

class LogRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, text):
        self.text_widget.append(text.strip())

    def flush(self):
        pass

class ProcessThread(QThread):
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, processor, input_file, output_file, mapping_file):
        super().__init__()
        self.processor = processor
        self.input_file = input_file
        self.output_file = output_file
        self.mapping_file = mapping_file

    def run(self):
        try:
            self.processor.process_csv(
                self.input_file,
                self.output_file,
                self.mapping_file
            )
            self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))

class TimecardGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.processor = TimecardProcessor()
        self.init_ui()

    def init_ui(self):
        """GUIの初期化"""
        self.setWindowTitle('タイムカード処理ツール')
        self.setGeometry(100, 100, 800, 600)

        # メインウィジェットとレイアウト
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout()
        main_widget.setLayout(layout)

        # マッピングファイル選択部分
        mapping_layout = QHBoxLayout()
        self.mapping_path = QLineEdit()
        self.mapping_path.setPlaceholderText('名前マッピングファイルを選択してください')
        mapping_button = QPushButton('マッピングファイル選択')
        mapping_button.clicked.connect(self.select_mapping_file)
        mapping_layout.addWidget(self.mapping_path)
        mapping_layout.addWidget(mapping_button)
        layout.addLayout(mapping_layout)

        # マッピングテーブル
        self.mapping_table = QTableWidget()
        self.mapping_table.setColumnCount(2)
        self.mapping_table.setHorizontalHeaderLabels(['カード番号', '名前'])
        # テーブルの列幅を設定
        self.mapping_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.mapping_table)

        # テーブル操作ボタン
        table_buttons_layout = QHBoxLayout()
        add_row_button = QPushButton('行を追加')
        add_row_button.clicked.connect(self.add_mapping_row)
        delete_row_button = QPushButton('選択行を削除')
        delete_row_button.clicked.connect(self.delete_mapping_row)
        save_mapping_button = QPushButton('マッピングを保存')
        save_mapping_button.clicked.connect(self.save_mapping)
        
        table_buttons_layout.addWidget(add_row_button)
        table_buttons_layout.addWidget(delete_row_button)
        table_buttons_layout.addWidget(save_mapping_button)
        layout.addLayout(table_buttons_layout)

        # 入力ファイル選択
        input_layout = QHBoxLayout()
        self.input_path = QLineEdit()
        self.input_path.setPlaceholderText('入力CSVファイルを選択してください')
        input_button = QPushButton('入力ファイル選択')
        input_button.clicked.connect(self.select_input_file)
        input_layout.addWidget(self.input_path)
        input_layout.addWidget(input_button)
        layout.addLayout(input_layout)

        # 出力ファイル選択
        output_layout = QHBoxLayout()
        self.output_path = QLineEdit()
        self.output_path.setPlaceholderText('出力Excelファイルを選択してください')
        output_button = QPushButton('出力ファイル選択')
        output_button.clicked.connect(self.select_output_file)
        output_layout.addWidget(self.output_path)
        output_layout.addWidget(output_button)
        layout.addLayout(output_layout)

        # 変換ボタン
        convert_button = QPushButton('変換開始')
        convert_button.clicked.connect(self.start_conversion)
        layout.addWidget(convert_button)

        # ログウィンドウ
        self.log_window = QTextEdit()
        self.log_window.setReadOnly(True)
        layout.addWidget(self.log_window)

        # 標準出力をログウィンドウにリダイレクト
        sys.stdout = LogRedirector(self.log_window)

    def add_mapping_row(self):
        """マッピングテーブルに新しい行を追加"""
        current_row = self.mapping_table.rowCount()
        self.mapping_table.insertRow(current_row)
        # デフォルト値を設定
        self.mapping_table.setItem(current_row, 0, QTableWidgetItem(''))
        self.mapping_table.setItem(current_row, 1, QTableWidgetItem(''))

    def delete_mapping_row(self):
        """選択された行を削除"""
        selected_rows = set(item.row() for item in self.mapping_table.selectedItems())
        for row in sorted(selected_rows, reverse=True):
            self.mapping_table.removeRow(row)

    def select_mapping_file(self):
        """マッピングファイルの選択と読み込み"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, 'マッピングファイルを選択', '', 'CSVファイル (*.csv)')
        
        if file_path:
            self.mapping_path.setText(file_path)
            self.load_mapping_file(file_path)

    def load_mapping_file(self, file_path):
        """マッピングファイルを読み込んでテーブルに表示"""
        try:
            df = pd.read_csv(file_path)
            self.mapping_table.setRowCount(len(df))
            
            for i, row in df.iterrows():
                # カード番号を4桁にゼロ埋め
                card_no = str(row[0]).zfill(4)
                self.mapping_table.setItem(i, 0, QTableWidgetItem(card_no))
                self.mapping_table.setItem(i, 1, QTableWidgetItem(str(row[1])))
                
            self.log_window.append(f"マッピングファイルを読み込みました: {len(df)}件")
        except Exception as e:
            QMessageBox.critical(self, 'エラー', f'マッピングファイルの読み込みに失敗しました: {str(e)}')

    def save_mapping(self):
        """マッピングテーブルの内容を保存"""
        try:
            data = []
            for row in range(self.mapping_table.rowCount()):
                card_no = self.mapping_table.item(row, 0).text().strip()
                name = self.mapping_table.item(row, 1).text().strip()
                if card_no and name:  # 空の行は保存しない
                    data.append([card_no, name])
            
            df = pd.DataFrame(data, columns=['カード番号', '名前'])
            df.to_csv(self.mapping_path.text(), index=False)
            self.log_window.append("マッピングファイルを保存しました")
        except Exception as e:
            QMessageBox.critical(self, 'エラー', f'マッピングファイルの保存に失敗しました: {str(e)}')

    def select_input_file(self):
        """入力ファイルの選択"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, '入力ファイルを選択', '', 'CSVファイル (*.csv)')
        if file_path:
            self.input_path.setText(file_path)

    def select_output_file(self):
        """出力ファイルの選択"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, '出力ファイルを選択', '', 'Excelファイル (*.xlsx)')
        if file_path:
            self.output_path.setText(file_path)

    def validate_inputs(self):
        """入力値の検証"""
        if not self.mapping_path.text():
            QMessageBox.warning(self, '警告', 'マッピングファイルを選択してください')
            return False
        if not self.input_path.text():
            QMessageBox.warning(self, '警告', '入力ファイルを選択してください')
            return False
        if not self.output_path.text():
            QMessageBox.warning(self, '警告', '出力ファイルを選択してください')
            return False
        return True

    def start_conversion(self):
        """変換処理の開始"""
        if not self.validate_inputs():
            return

        # 処理開始前にマッピングを保存
        self.save_mapping()
        
        self.log_window.append("変換処理を開始します...")
        
        # 処理スレッドの作成と開始
        self.thread = ProcessThread(
            self.processor,
            self.input_path.text(),
            self.output_path.text(),
            self.mapping_path.text()
        )
        self.thread.finished.connect(self.conversion_finished)
        self.thread.error.connect(self.conversion_error)
        self.thread.start()

    def conversion_finished(self):
        """変換処理完了時の処理"""
        self.log_window.append("変換処理が完了しました")
        QMessageBox.information(self, '完了', '変換処理が完了しました')

    def conversion_error(self, error_msg):
        """変換処理エラー時の処理"""
        self.log_window.append(f"エラーが発生しました: {error_msg}")
        QMessageBox.critical(self, 'エラー', f'変換処理中にエラーが発生しました: {error_msg}')

def main():
    app = QApplication(sys.argv)
    gui = TimecardGUI()
    gui.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()