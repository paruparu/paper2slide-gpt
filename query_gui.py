import os
import time
import fitz  # PyMuPDF
from xml.dom import minidom
import dicttoxml
from PyPDF2 import PdfReader
from dotenv import load_dotenv
from PIL import Image
import io
import openai


openai.api_key = os.getenv('OPENAI_API_KEY')

prompt = """与えられた論文の要点をまとめ、以下の項目で日本語で出力せよ。それぞれの項目は最大でも180文字以内に要約せよ。
```
論文名:タイトルの日本語訳
キーワード:この論文のキーワード
課題:この論文が解決する課題
手法:この論文が提案する手法
結果:提案手法によって得られた結果
```"""

def get_summary(metadata):
    title = metadata['title']
    if isinstance(title, list):
        title = ''.join(title)

    text = f"title: {title}\nbody: {metadata.get('abstract', 'N/A')}"
    response = openai.ChatCompletion.create(
                model='gpt-4o',
                messages=[
                    {'role': 'system', 'content': prompt},
                    {'role': 'user', 'content': text}
                ],
                temperature=0.25,
            )

    summary = response['choices'][0]['message']['content']

    summary_dict = {}
    for line in summary.split('\n'):
        if line.startswith("論文名"):
            summary_dict['title_jp'] = line[4:].strip()
        elif line.startswith("キーワード"):
            summary_dict['keywords'] = line[6:].strip()
        elif line.startswith("課題"):
            summary_dict['problem'] = line[3:].strip()
        elif line.startswith("手法"):
            summary_dict['method'] = line[3:].strip()
        elif line.startswith("結果"):
            summary_dict['result'] = line[3:].strip()

    # Check for missing fields
    required_fields = ['title_jp', 'keywords', 'problem', 'method', 'result']
    for field in required_fields:
        if field not in summary_dict:
            summary_dict[field] = "N/A"
    return summary_dict

def recoverpix(doc, item):
    xref = item[0]
    smask = item[1]

    if smask > 0:
        pix0 = fitz.Pixmap(doc.extract_image(xref)["image"])
        if pix0.alpha:
            pix0 = fitz.Pixmap(pix0, 0)
        mask = fitz.Pixmap(doc.extract_image(smask)["image"])

        try:
            pix = fitz.Pixmap(pix0, mask)
        except:
            pix = fitz.Pixmap(doc.extract_image(xref)["image"])

        ext = "pam" if pix0.n > 3 else "png"
        return {"ext": ext, "colorspace": pix.colorspace.n, "image": pix.tobytes(ext)}

    if "/ColorSpace" in doc.xref_object(xref, compressed=True):
        pix = fitz.Pixmap(doc, xref)
        pix = fitz.Pixmap(fitz.csRGB, pix)
        return {"ext": "png", "colorspace": 3, "image": pix.tobytes("png")}

    return doc.extract_image(xref)

def extract_images_from_pdf(pdf_path, imgdir="./output", min_width=400, min_height=400, relsize=0.05, abssize=2048, max_ratio=8, max_num=5):
    if not os.path.exists(imgdir):
        os.makedirs(imgdir)

    t0 = time.time()
    doc = fitz.open(pdf_path)
    page_count = doc.page_count

    xreflist = []
    images = []
    for pno in range(page_count):
        if len(images) >= max_num:
            break
        il = doc.get_page_images(pno)
        for img in il:
            xref = img[0]
            if xref in xreflist:
                continue
            width = img[2]
            height = img[3]
            if width < min_width and height < min_height:
                continue
            image = recoverpix(doc, img)

            imgdata = image["image"]

            if len(imgdata) <= abssize:
                continue

            if width / height > max_ratio or height/width > max_ratio:
                continue

            imgname = f"img{pno+1:02d}_{xref:05d}.{image['ext']}"
            images.append((imgname, pno+1, width, height))
            imgfile = os.path.join(imgdir, imgname)
            with open(imgfile, "wb") as fout:
                fout.write(imgdata)
            xreflist.append(xref)

    t1 = time.time()
    return xreflist, images

def get_half(fname, imgdir):
    pdf_file = fitz.open(fname)
    page = pdf_file[0]
    mat = fitz.Matrix(2, 2)
    pix = page.get_pixmap(matrix=mat)
    im = Image.open(io.BytesIO(pix.tobytes()))
    width, height = im.size
    box = (0, height // 20, width, (height // 2) + (height // 20))
    im_cropped = im.crop(box)
    half_img_path = os.path.join(imgdir, "half.png")
    im_cropped.save(half_img_path, "PNG")
    return half_img_path

def get_metadata_from_pdf(pdf_path):
    if not os.path.exists(pdf_path):
        return None

    with open(pdf_path, "rb") as f:
        reader = PdfReader(f)

        if reader.is_encrypted:
            try:
                reader.decrypt('')
            except:
                return None

        info = reader.metadata
        metadata = {}
        
        metadata['title'] = info.title if info.title else "Unknown"
        metadata['authors'] = info.author.split(',') if info.author else ["Unknown"]
        metadata['subject'] = info.subject if info.subject else "N/A"
        if info.producer:
            metadata['producer'] = info.producer
        if info.creation_date:
            metadata['creation_date'] = info.creation_date
        if info.modification_date:
            metadata['mod_date'] = info.modification_date

        text = ""
        for page_num in range(min(3, len(reader.pages))):
            page_text = reader.pages[page_num].extract_text()
            if page_text:
                text += page_text
        metadata['abstract'] = text[:2000]
        metadata['pdf_path'] = pdf_path

        return metadata

def convert_lists_to_strings(d):
    from PyPDF2.generic import TextStringObject
    if isinstance(d, dict):
        for key, value in d.items():
            if isinstance(value, list) and all(isinstance(i, str) for i in value):
                d[key] = ''.join(value)
            elif isinstance(value, TextStringObject):
                d[key] = str(value)
            elif isinstance(value, list):
                d[key] = [convert_lists_to_strings(item) for item in value]
            else:
                d[key] = convert_lists_to_strings(value)
    elif isinstance(d, list) and all(isinstance(i, str) for i in d):
        d = ''.join(d)
    elif isinstance(d, list):
        d = [convert_lists_to_strings(item) for item in d]
    elif 'TextStringObject' in str(type(d)):
        d = str(d)
    return d

def save_as_xml(data, filepath):
    data = convert_lists_to_strings(data)
    xml_content = dicttoxml.dicttoxml(data, attr_type=False, root=False).decode('utf-8')
    pretty_xml = minidom.parseString(xml_content).toprettyxml(indent="   ")
    with open(filepath, "w") as xml_file:
        xml_file.write(pretty_xml)

def process_pdf(pdf_file, dir='./output', timeout_sec=60, start_time=None):
    if start_time is None:
        start_time = time.time()

    if not os.path.exists(dir):
        os.makedirs(dir)

    # dir自体にもスペースがあれば置換する
    dir = dir.replace(' ', '_')

    metadata = get_metadata_from_pdf(pdf_file)
    if metadata is None:
        raise ValueError("PDFメタデータの取得に失敗")

    entry_id = os.path.splitext(os.path.basename(pdf_file))[0]
    # entry_idから空白を除去
    entry_id = entry_id.replace(' ', '_')

    # xmls以下に格納
    xmls_dir = os.path.join(dir, "xmls".replace(' ', '_'))
    if not os.path.exists(xmls_dir):
        os.makedirs(xmls_dir)

    dirpath = os.path.join(xmls_dir, entry_id)
    if not os.path.exists(dirpath):
        os.makedirs(dirpath)

    # imagesディレクトリもスペースをアンダースコアに
    images_dir = os.path.join(dirpath, "images".replace(' ', '_'))
    if not os.path.exists(images_dir):
        os.makedirs(images_dir)

    image_count = extract_images_from_pdf(pdf_file, images_dir)
    half_img_path = get_half(pdf_file, images_dir)

    # 画像やメタデータ内のパスにも空白が入る可能性があるので全て置換
    # images_dir内のファイル名にもスペースがあれば置換
    for f in os.listdir(images_dir):
        new_name = f.replace(' ', '_')
        if new_name != f:
            old_path = os.path.join(images_dir, f)
            new_path = os.path.join(images_dir, new_name)
            os.rename(old_path, new_path)

    summary_info = get_summary(metadata)

    # paper_info中のパス文字列にもスペースが残らないよう処理
    def no_space_path(p):
        if isinstance(p, list):
            return [no_space_path(x) for x in p]
        elif isinstance(p, str):
            return p.replace(' ', '_')
        return p

    paper_info = {
        'paper': {
            'title': metadata['title'],
            'authors': metadata['authors'],
            'abstract': metadata.get('abstract', 'N/A'),
            'pdf': pdf_file.replace(' ', '_'),
            'image_count': image_count,
            'images': [no_space_path(os.path.join(images_dir, f)) for f in os.listdir(images_dir)],
            'half_img_path': no_space_path(half_img_path),
            'title_jp': summary_info['title_jp'],
            'keywords': summary_info['keywords'],
            'problem': summary_info['problem'],
            'method': summary_info['method'],
            'result': summary_info['result'],
            'query': "N/A"
        }
    }

    # 再度paper.xmlに書き込む前にパス文字列をチェックし、スペースを_に変換
    xml_path = os.path.join(dirpath, "paper.xml".replace(' ', '_'))
    save_as_xml(paper_info, xml_path)
    return dirpath