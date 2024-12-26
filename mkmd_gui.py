# mkmd_gui.py
import os
import glob
import xmltodict
import re
import time
from PIL import Image

def safe_filename(filename):
    return re.sub(r'[^a-zA-Z0-9_\-]', '_', filename)

def make_md(dirname, filename, output_dir='./output_marp', min_size_kb=100):
    # dirname例: ユーザー指定ディレクトリ/xmls/(entry_id)
    # filename: "paper.xml"
    path = os.path.join(dirname, filename)
    with open(path, "r", encoding="utf-8") as fin:
        xml = fin.read()
        print(f"Processing file: {filename}")

    dict_data = xmltodict.parse(xml)['paper']
    print(dict_data)

    # キーが存在しない場合のデフォルト値を設定
    title_jp = dict_data.get('title_jp', 'N/A')
    title = dict_data.get('title', 'N/A')
    print(f"Full title: {title}")
    year = dict_data.get('year', 'N/A')
    entry_id = dict_data.get('entry_id', 'N/A')
    problem = dict_data.get('problem', 'N/A')
    method = dict_data.get('method', 'N/A')
    result = dict_data.get('result', 'N/A')
    half_img_path = dict_data.get('half_img_path', None)

    # タイトルからファイル名生成
    safe_title = safe_filename(title[:14])
    print(f"Safe title (first 14 chars): {safe_title}")
    output_filename = f"{safe_title}_output.md"
    output_path = os.path.join(output_dir, output_filename)

    # 出力ディレクトリが存在しない場合は作成
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    with open(output_path, "w", encoding="utf-8") as f:
        # Write header
        f.write("---\n")
        f.write("marp: true\n")
        f.write("theme: default\n")
        f.write("size: 16:9\n")
        f.write("paginate: true\n")
        f.write('_class: ["cool-theme"]\n')
        # スタイル指定: section内でoverflowを許可する
        f.write("style: |\n")
        f.write("  section {\n")
        f.write("    overflow-y: auto;\n")
        f.write("  }\n")
        f.write("\n---\n")
        f.write(f"# {title}\n")
        f.write("\n")

        f.write('\n---\n')
        f.write('<!-- _class: title -->\n')
        f.write(f"# {title_jp}\n")
        f.write(f"{title}\n")
        f.write(f"[{year}] {entry_id}\n")
        f.write(f"__課題__ {problem}\n")
        f.write(f"__手法__ {method}\n")
        f.write(f"__結果__ {result}\n")

        # half_img_pathがある場合のみ表示
        if half_img_path and os.path.exists(half_img_path):
            relative_half_img_path = os.path.relpath(half_img_path, output_dir)
            f.write("\n---\n")
            f.write('<!-- _class: info -->\n')
            f.write(f'![width:1400]({relative_half_img_path})\n')
        else:
            print("No half_img_path or file not found, skipping half image.")

        # 画像処理
        images_dir = os.path.join(dirname, "images")
        if os.path.exists(images_dir):
            images = [img for img in os.listdir(images_dir) if img.lower().endswith(('.png', '.jpg', '.jpeg'))]
            # half.pngを除外（既に表示済）
            if 'half.png' in images:
                images.remove('half.png')

            # 画像ファイルサイズフィルタリング
            valid_images = []
            for img in images:
                img_path = os.path.join(images_dir, img)
                img_size_kb = os.path.getsize(img_path) / 1024
                print(f"Image: {img}, Size: {img_size_kb:.2f} KB")

                if img_size_kb > min_size_kb:
                    valid_images.append(img)

            for img in valid_images:
                img_path = os.path.join(images_dir, img)
                relative_img_path = os.path.relpath(img_path, output_dir)
                with Image.open(img_path) as image:
                    width, height = image.size
                    x_ratio = (1600.0 * 0.7) / float(width)
                    y_ratio = (900.0 * 0.7) / float(height)
                    ratio = min(x_ratio, y_ratio)
                    resized_width = int(ratio * width)

                    f.write("\n---\n")
                    f.write('<!-- _class: info -->\n')
                    f.write(f'![width:{resized_width}]({relative_img_path})\n')

            # 有効な画像がない場合の警告
            if not valid_images:
                print(f"Warning: No valid images found above {min_size_kb} KB")
        else:
            print("No images directory found, skipping additional images.")

    return output_path

def convert_xmls_to_md(dir_path, output_dir='./output_marp', min_size_kb=100, timeout_sec=60, start_time=None):
    if start_time is None:
        start_time = time.time()

    xml_files = glob.glob(os.path.join(dir_path, "*.xml"))

    if not xml_files:
        print(f"{dir_path} にXMLファイルが見つかりません。")
        raise FileNotFoundError(f"{dir_path} にXMLファイルがありません。")

    paper_xml_path = os.path.join(dir_path, "paper.xml")
    if not os.path.exists(paper_xml_path):
        print(f"{dir_path} にpaper.xmlが存在しません。")
        raise FileNotFoundError(f"{dir_path} にpaper.xmlが存在しません。")

    if (time.time() - start_time) > timeout_sec:
        raise Exception("タイムアウトに達しました")

    dirname, filename = os.path.split(paper_xml_path)
    if (time.time() - start_time) > timeout_sec:
        raise Exception("タイムアウトに達しました")

    md_file = make_md(dirname, filename, output_dir=output_dir, min_size_kb=min_size_kb)
    return md_file