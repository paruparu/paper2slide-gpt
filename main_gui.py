import sys
import os
import time
import subprocess
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog,
    QMessageBox, QHBoxLayout, QDialog, QFormLayout, QLineEdit
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, pyqtSlot, QTimer
from PyQt5.QtGui import QIcon
from dotenv import load_dotenv
import dotenv
# 環境変数をここで読み込む
load_dotenv(override=True)
import query_gui
import mkmd_gui
import md2pdf
import md2pptx
import xmltodict

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("設定")
        self.setModal(True)

        load_dotenv()
        current_output_dir = os.getenv("OUTPUT_DIR", "./output")
        current_timeout_str = os.getenv("TIMEOUT_SEC", "60")
        current_api_key = os.getenv("OPENAI_API_KEY", "")

        layout = QFormLayout()

        # OpenAI API キー設定
        self.api_key_edit = QLineEdit(current_api_key, self)
        layout.addRow("OpenAI API Key", self.api_key_edit)

        # 出力ディレクトリ設定
        self.output_dir_edit = QLineEdit(current_output_dir, self)
        browse_button = QPushButton("参照")
        browse_button.clicked.connect(self.browse_output_dir)
        dir_layout = QHBoxLayout()
        dir_layout.addWidget(self.output_dir_edit)
        dir_layout.addWidget(browse_button)
        layout.addRow("出力ディレクトリ", dir_layout)

        # タイムアウト秒数設定
        self.timeout_edit = QLineEdit(current_timeout_str, self)
        layout.addRow("タイムアウト(秒)", self.timeout_edit)

        # OK, Cancelボタン
        btn_layout = QHBoxLayout()
        ok_button = QPushButton("OK")
        cancel_button = QPushButton("キャンセル")
        ok_button.clicked.connect(self.save_settings)
        cancel_button.clicked.connect(self.reject)
        btn_layout.addWidget(ok_button)
        btn_layout.addWidget(cancel_button)
        layout.addRow(btn_layout)

        self.setLayout(layout)

    def browse_output_dir(self):
        dir_path = QFileDialog.getExistingDirectory(
            self, "出力ディレクトリを選択", self.output_dir_edit.text()
        )
        if dir_path:
            self.output_dir_edit.setText(dir_path)

    def save_settings(self):
        new_api_key = self.api_key_edit.text().strip()
        new_output_dir = self.output_dir_edit.text().strip()
        new_timeout_str = self.timeout_edit.text().strip()

        if not new_api_key:
            QMessageBox.warning(self, "警告", "OpenAI API Key を入力してください。")
            return
        if not new_output_dir:
            QMessageBox.warning(self, "警告", "有効なディレクトリを指定してください。")
            return
        if not new_timeout_str.isdigit():
            QMessageBox.warning(self, "警告", "タイムアウトには数値を入力してください。")
            return

        env_file_path = ".env"
        if not os.path.exists(env_file_path):
            with open(env_file_path, "w") as f:
                f.write("")

        env_data = dotenv.dotenv_values(env_file_path)
        env_data["OPENAI_API_KEY"] = new_api_key
        env_data["OUTPUT_DIR"] = new_output_dir
        env_data["TIMEOUT_SEC"] = new_timeout_str

        with open(env_file_path, "w", encoding="utf-8") as f:
            for k, v in env_data.items():
                f.write(f"{k}={v}\n")

        QMessageBox.information(self, "情報", "設定が保存されました。")
        self.accept()


class Worker(QThread):
    finished = pyqtSignal(str)  # (dirpath)を返す
    error = pyqtSignal(str)

    def __init__(self, pdf_path, output_dir, timeout_sec):
        super().__init__()
        self.pdf_path = pdf_path
        self.output_dir = output_dir
        self.timeout_sec = timeout_sec

    def run(self):
        start_time = time.time()
        try:
            # PDF解析してxmls生成のみ行う
            dirpath = query_gui.process_pdf(
                self.pdf_path,
                dir=self.output_dir,
                timeout_sec=self.timeout_sec,
                start_time=start_time
            )
            if (time.time() - start_time) > self.timeout_sec:
                raise Exception("タイムアウトに達しました")
        except Exception as e:
            self.error.emit(str(e))
            return
        self.finished.emit(dirpath)


class PdfWorker(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    def __init__(self, md_file, pdf_output, timeout_sec):
        super().__init__()
        self.md_file = md_file
        self.pdf_output = pdf_output
        self.timeout_sec = timeout_sec

    def run(self):
        start_time = time.time()
        try:
            md2pdf.convert_md_to_pdf(
                self.md_file, self.pdf_output,
                timeout_sec=self.timeout_sec,
                start_time=start_time
            )
        except Exception as e:
            self.error.emit(str(e))
            return
        self.finished.emit(self.pdf_output)


class PptxWorker(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    def __init__(self, md_file, pptx_output, timeout_sec):
        super().__init__()
        self.md_file = md_file
        self.pptx_output = pptx_output
        self.timeout_sec = timeout_sec

    def run(self):
        # PPTX出力にはmd2pptxにtimeoutが無いため、run中に計測してtimeoutを超えたらエラーにします。
        start_time = time.time()
        while True:
            elapsed = time.time() - start_time
            if elapsed > self.timeout_sec:
                self.error.emit("PPTX変換中にタイムアウトしました")
                return
            try:
                md2pptx.convert_md_to_pptx(
                    self.md_file, pptx_output_file=self.pptx_output
                )
                break
            except Exception as e:
                self.error.emit(str(e))
                return
        self.finished.emit(self.pptx_output)


class TitleEditDialog(QDialog):
    def __init__(self, paper_xml_path, parent=None):
        super().__init__(parent)
        self.setWindowTitle("タイトル編集")
        self.setModal(True)
        self.paper_xml_path = paper_xml_path

        with open(self.paper_xml_path, 'r', encoding='utf-8') as f:
            self.paper_data = xmltodict.parse(f.read())

        current_title = self.paper_data.get('paper', {}).get('title', 'Unknown')

        layout = QFormLayout()
        self.title_edit = QLineEdit(current_title)
        layout.addRow("タイトル:", self.title_edit)

        btn_layout = QHBoxLayout()
        ok_button = QPushButton("OK")
        cancel_button = QPushButton("キャンセル")
        ok_button.clicked.connect(self.save_title)
        cancel_button.clicked.connect(self.reject)
        btn_layout.addWidget(ok_button)
        btn_layout.addWidget(cancel_button)

        layout.addRow(btn_layout)
        self.setLayout(layout)

    def save_title(self):
        new_title = self.title_edit.text().strip()
        if not new_title:
            QMessageBox.warning(self, "警告", "タイトルを入力してください。")
            return

        # Update the in-memory data
        self.paper_data['paper']['title'] = new_title

        # Save updated XML
        xml_str = xmltodict.unparse(self.paper_data, pretty=True)
        with open(self.paper_xml_path, 'w', encoding='utf-8') as f:
            f.write(xml_str)

        self.accept()


class PDFApp(QWidget):
    def __init__(self):
        super().__init__()

        # アイコンを設定
        self.setWindowIcon(QIcon('icon.ico'))

        load_dotenv(override=True)
        self.output_dir = os.getenv("OUTPUT_DIR", "./output")
        self.timeout_sec = int(os.getenv("TIMEOUT_SEC", "60"))
        self.api_key = os.getenv("OPENAI_API_KEY", "")

        self.pdf_path = None
        self.worker = None
        self.pdf_worker = None
        self.pptx_worker = None
        self.countdown_timer = None
        self.remaining_time = self.timeout_sec

        self.generated_md_file = None
        self.dirpath = None

        self.setWindowTitle("PDF to Summary & Markdown Converter (Timeout対応)")
        self.setGeometry(300, 300, 600, 300)

        main_layout = QVBoxLayout()

        self.label = QLabel("PDFファイルを選択してください:", self)
        main_layout.addWidget(self.label, alignment=Qt.AlignCenter)

        btn_layout = QHBoxLayout()

        self.select_button = QPushButton("PDFを選択", self)
        self.select_button.clicked.connect(self.select_pdf)
        btn_layout.addWidget(self.select_button, alignment=Qt.AlignCenter)

        self.run_button = QPushButton("処理", self)
        self.run_button.setEnabled(False)
        self.run_button.clicked.connect(self.start_processing)
        btn_layout.addWidget(self.run_button, alignment=Qt.AlignCenter)

        self.settings_button = QPushButton("設定")
        self.settings_button.clicked.connect(self.open_settings)
        btn_layout.addWidget(self.settings_button, alignment=Qt.AlignCenter)

        main_layout.addLayout(btn_layout)

        self.status_label = QLabel("", self)
        main_layout.addWidget(self.status_label, alignment=Qt.AlignCenter)

        self.timeout_label = QLabel("", self)
        main_layout.addWidget(self.timeout_label, alignment=Qt.AlignCenter)

        output_btn_layout = QHBoxLayout()

        self.pdf_button = QPushButton("PDFに出力")
        self.pdf_button.setEnabled(False)
        self.pdf_button.clicked.connect(self.start_pdf_export)
        output_btn_layout.addWidget(self.pdf_button)

        self.pptx_button = QPushButton("PPTXに出力")
        self.pptx_button.setEnabled(False)
        self.pptx_button.clicked.connect(self.start_pptx_export)
        output_btn_layout.addWidget(self.pptx_button)

        main_layout.addLayout(output_btn_layout)

        self.setLayout(main_layout)

        # ★ アプリ起動時にAPIキーが無ければ設定画面を強制表示
        if not self.api_key:
            QMessageBox.warning(self, "要OpenAI API Key", "OpenAI API Keyが設定されていません。設定画面を開きます。")
            self.prompt_for_api_key()

    def prompt_for_api_key(self):
        """
        起動時にAPIキーが空の場合に呼ばれる。
        設定ダイアログを開き、ユーザーがキャンセルしたらアプリ終了。
        """
        dialog = SettingsDialog(self)
        result = dialog.exec_()
        if result == QDialog.Accepted:
            load_dotenv(override=True)
            self.api_key = os.getenv("OPENAI_API_KEY", "")
            self.output_dir = os.getenv("OUTPUT_DIR", "./output")
            self.timeout_sec = int(os.getenv("TIMEOUT_SEC", "60"))
            if not self.api_key:
                QMessageBox.critical(self, "エラー", "APIキーが設定されませんでした。アプリを終了します。")
                sys.exit(1)
        else:
            # キャンセルされた場合
            sys.exit(0)

    def select_pdf(self):
        pdf_file, _ = QFileDialog.getOpenFileName(self, "PDFを選択", "", "PDF Files (*.pdf)")
        if pdf_file:
            self.pdf_path = pdf_file
            self.status_label.setText(f"選択中のPDF: {os.path.basename(pdf_file)}")
            self.run_button.setEnabled(True)

    def start_processing(self):
        if not self.pdf_path:
            QMessageBox.critical(self, "エラー", "PDFファイルが選択されていません。")
            return

        # 改めて env を再読込
        load_dotenv(override=True)
        self.api_key = os.getenv("OPENAI_API_KEY", "")
        if not self.api_key:
            QMessageBox.critical(self, "エラー", "OpenAI API Key が設定されていません。")
            return

        self.output_dir = os.getenv("OUTPUT_DIR", "./output")
        self.timeout_sec = int(os.getenv("TIMEOUT_SEC", "60"))

        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

        self.status_label.setText("実行中…")
        self.run_button.setEnabled(False)
        self.select_button.setEnabled(False)
        self.settings_button.setEnabled(False)
        self.pdf_button.setEnabled(False)
        self.pptx_button.setEnabled(False)

        self.worker = Worker(self.pdf_path, self.output_dir, self.timeout_sec)
        self.worker.finished.connect(self.on_processing_finished)
        self.worker.error.connect(self.on_processing_error)
        self.worker.start()

        # ここでタイマー開始
        self.start_countdown()

    def start_countdown(self):
        self.stop_countdown()
        self.remaining_time = self.timeout_sec
        self.timeout_label.setText(f"タイムアウトまで {self.remaining_time} 秒")
        self.countdown_timer = QTimer(self)
        self.countdown_timer.timeout.connect(self.update_countdown)
        self.countdown_timer.start(1000)

    def stop_countdown(self):
        if self.countdown_timer:
            self.countdown_timer.stop()
            self.countdown_timer = None
        self.timeout_label.setText("")

    def update_countdown(self):
        self.remaining_time -= 1
        if self.remaining_time <= 0:
            self.timeout_label.setText("タイムアウトしました!")
            self.stop_countdown()
        else:
            self.timeout_label.setText(f"タイムアウトまで {self.remaining_time} 秒")

    @pyqtSlot(str)
    def on_processing_finished(self, dirpath):
        self.dirpath = dirpath
        self.status_label.setText(
            "処理完了！xmlsフォルダに中間ファイルが生成されました。"
        )

        # タイトル編集ダイアログを表示
        paper_xml_path = os.path.join(self.dirpath, "paper.xml")
        if os.path.exists(paper_xml_path):
            dialog = TitleEditDialog(paper_xml_path, self)
            dialog.exec_()  # ユーザーがOKを押せばtitleが更新

        # タイトル編集が完了したらMD生成
        try:
            marp_dir = os.path.join(self.output_dir, "output_marp")
            if not os.path.exists(marp_dir):
                os.makedirs(marp_dir)
            # 最新タイトルでMD生成
            start_time = time.time()
            self.generated_md_file = mkmd_gui.convert_xmls_to_md(
                self.dirpath,
                output_dir=marp_dir,
                timeout_sec=self.timeout_sec,
                start_time=start_time
            )
            QMessageBox.information(
                self, "完了", "処理が完了しました。\n'./output_marp'フォルダでMDを確認できます。"
            )
        except Exception as e:
            QMessageBox.critical(
                self, "エラー", f"MD生成中にエラーが発生:\n{e}"
            )

        self.reset_ui()
        self.pdf_button.setEnabled(True)
        self.pptx_button.setEnabled(True)

    @pyqtSlot(str)
    def on_processing_error(self, error_msg):
        QMessageBox.critical(
            self, "エラー", f"処理中にエラーが発生:\n{error_msg}"
        )
        self.reset_ui()

    def reset_ui(self):
        self.run_button.setEnabled(True if self.pdf_path else False)
        self.select_button.setEnabled(True)
        self.settings_button.setEnabled(True)
        self.stop_countdown()
        self.worker = None

    def open_settings(self):
        dialog = SettingsDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            # ここでは「設定が保存されました」のメッセージを重複表示しない
            load_dotenv(override=True)
            self.api_key = os.getenv("OPENAI_API_KEY", "")
            self.output_dir = os.getenv("OUTPUT_DIR", "./output")
            self.timeout_sec = int(os.getenv("TIMEOUT_SEC", "60"))
        else:
            # キャンセルの場合は何もしない
            pass

    def start_pdf_export(self):
        if not self.generated_md_file:
            QMessageBox.warning(self, "警告", "処理が完了していません。")
            return
        pdf_filename = self.generate_unique_filename(
            os.path.splitext(
                os.path.basename(self.generated_md_file)
            )[0],
            "pdf",
            self.output_dir
        )
        pdf_output = os.path.join(self.output_dir, pdf_filename)

        self.disable_ui_during_export()
        self.pdf_worker = PdfWorker(
            self.generated_md_file, pdf_output, self.timeout_sec
        )
        self.pdf_worker.finished.connect(self.on_pdf_finished)
        self.pdf_worker.error.connect(self.on_pdf_error)
        self.pdf_worker.start()

        # PDF変換もタイマー開始
        self.start_countdown()

    @pyqtSlot(str)
    def on_pdf_finished(self, pdf_path):
        QMessageBox.information(
            self, "PDF出力", f"PDFファイルを出力しました: {pdf_path}"
        )
        self.enable_ui_after_export()

    @pyqtSlot(str)
    def on_pdf_error(self, error_msg):
        QMessageBox.critical(
            self, "エラー", f"PDF出力中にエラー:\n{error_msg}"
        )
        self.enable_ui_after_export()

    def start_pptx_export(self):
        if not self.generated_md_file:
            QMessageBox.warning(self, "警告", "処理が完了していません。")
            return
        base_title = os.path.splitext(
            os.path.basename(self.generated_md_file)
        )[0]
        pptx_filename = self.generate_unique_filename(
            base_title, "pptx", self.output_dir
        )
        pptx_output = os.path.join(self.output_dir, pptx_filename)

        self.disable_ui_during_export()
        self.pptx_worker = PptxWorker(
            self.generated_md_file, pptx_output, self.timeout_sec
        )
        self.pptx_worker.finished.connect(self.on_pptx_finished)
        self.pptx_worker.error.connect(self.on_pptx_error)
        self.pptx_worker.start()

        # PPTX変換もタイマー開始
        self.start_countdown()

    @pyqtSlot(str)
    def on_pptx_finished(self, pptx_path):
        QMessageBox.information(
            self, "PPTX出力", f"PPTXファイルを出力しました: {pptx_path}"
        )
        self.enable_ui_after_export()

    @pyqtSlot(str)
    def on_pptx_error(self, error_msg):
        QMessageBox.critical(
            self, "エラー", f"PPTX出力中にエラー:\n{error_msg}"
        )
        self.enable_ui_after_export()

    def disable_ui_during_export(self):
        self.run_button.setEnabled(False)
        self.select_button.setEnabled(False)
        self.settings_button.setEnabled(False)
        self.pdf_button.setEnabled(False)
        self.pptx_button.setEnabled(False)
        self.status_label.setText("出力中…")

    def enable_ui_after_export(self):
        self.run_button.setEnabled(True if self.pdf_path else False)
        self.select_button.setEnabled(True)
        self.settings_button.setEnabled(True)
        self.pdf_button.setEnabled(True if self.generated_md_file else False)
        self.pptx_button.setEnabled(True if self.generated_md_file else False)
        self.stop_countdown()
        self.pdf_worker = None
        self.pptx_worker = None
        self.status_label.setText("")

    def generate_unique_filename(self, base_name, ext, target_dir):
        out_name = f"{base_name}.{ext}"
        index = 2
        while os.path.exists(os.path.join(target_dir, out_name)):
            out_name = f"{base_name}_{index}.{ext}"
            index += 1
        return out_name


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon('icon.ico'))  # タスクバーアイコンも設定
    window = PDFApp()
    window.show()
    sys.exit(app.exec_())