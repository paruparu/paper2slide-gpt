# api_settings_dialog.py
import os
from PyQt5.QtWidgets import QDialog, QFormLayout, QLineEdit, QHBoxLayout, QPushButton, QMessageBox
from dotenv import load_dotenv, dotenv_values

class APISettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("APIキー設定")
        self.setModal(True)

        # .envから読み込んだ現在のAPIキーを、デフォルト値として表示
        load_dotenv()  # .envファイルを読み込む
        current_api_key = os.getenv("OPENAI_API_KEY", "")

        layout = QFormLayout()

        self.api_key_edit = QLineEdit(current_api_key, self)
        layout.addRow("OpenAI APIキー", self.api_key_edit)

        # OK/キャンセルボタン
        btn_layout = QHBoxLayout()
        ok_button = QPushButton("OK")
        cancel_button = QPushButton("キャンセル")
        ok_button.clicked.connect(self.save_api_key)
        cancel_button.clicked.connect(self.reject)
        btn_layout.addWidget(ok_button)
        btn_layout.addWidget(cancel_button)

        layout.addRow(btn_layout)
        self.setLayout(layout)

    def save_api_key(self):
        new_api_key = self.api_key_edit.text().strip()

        if not new_api_key:
            QMessageBox.warning(self, "警告", "APIキーを入力してください。")
            return

        # APIキーを.envに書き込み
        self._write_env_var("OPENAI_API_KEY", new_api_key)
        QMessageBox.information(self, "情報", "APIキーが保存されました。")
        self.accept()

    def _write_env_var(self, key, value):
        """
        指定したkeyとvalueを.envに書き込む。
        既存の.envがあれば読み込んで上書きし、なければ新規作成する。
        """
        env_file_path = ".env"
        # 既存の.envの内容を取得（無ければ空ディクショナリ）
        env_data = {}
        if os.path.exists(env_file_path):
            env_data = dotenv_values(env_file_path)

        # 上書き or 新規追加
        env_data[key] = value

        # ファイルに書き戻し
        with open(env_file_path, "w", encoding="utf-8") as f:
            for k, v in env_data.items():
                f.write(f"{k}={v}\n")