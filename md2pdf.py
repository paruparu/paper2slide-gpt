# md2pdf.py
import time
import os
import subprocess

def convert_md_to_pdf(md_file, pdf_output_file, timeout_sec=60, start_time=None):
    if start_time is None:
        start_time = time.time()

    # コマンド構築: Marp CLIを利用
    command = [
        'marp',
        '--pdf',
        '--allow-local-files',  # ローカルファイル参照を許可
        md_file,
        '--output',
        pdf_output_file
    ]

    # 出力先ディレクトリがなければ作成
    output_dir = os.path.dirname(pdf_output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 現時点で残りタイムを計算しtimeoutに反映
    elapsed = time.time() - start_time
    remaining_timeout = timeout_sec - elapsed
    if remaining_timeout <= 0:
        raise Exception("PDF変換開始前に既にタイムアウト")

    try:
        subprocess.run(command, check=True, timeout=remaining_timeout)
        print(f"Successfully converted {md_file} to {pdf_output_file}")
    except subprocess.TimeoutExpired:
        raise Exception("PDF変換中にタイムアウトしました")
    except subprocess.CalledProcessError as e:
        # marpコマンドがエラー終了した場合
        raise Exception(f"PDF変換中にエラーが発生しました: {e}")