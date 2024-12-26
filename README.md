```markdown
# paper2slide-gpt

論文（paper）の内容を要約し、スライド（slide）向けにテキスト生成を行うサンプルアプリケーションです。  
OpenAI API を利用してテキストを生成するため、**OpenAI の API キー** が必要です。

---

## 目次

1. [環境構築](#環境構築)
   1. [conda (Anaconda) を利用する場合](#conda-anaconda-を利用する場合)
   2. [pip を利用する場合](#pip-を利用する場合)
2. [環境変数の設定 (.env ファイル)](#環境変数の設定-env-ファイル)
3. [アプリケーションの起動](#アプリケーションの起動)
4. [ライセンス](#ライセンス)

---

## 環境構築

### 1.1. conda (Anaconda) を利用する場合

1. このリポジトリをクローン（またはダウンロード）する

   ```bash
   git clone https://github.com/your-username/paper2slide-gpt.git
   cd paper2slide-gpt
   ```

2. `environment.yml` を使って環境を構築する

   ```bash
   conda env create -f environment.yml
   ```

3. 作成された環境をアクティブ化する

   ```bash
   conda activate paper2slide-gpt
   ```
   > ※ 環境名が `paper2slide-gpt` でない場合は、`environment.yml` をご確認ください。

### 1.2. pip を利用する場合

1. このリポジトリをクローン（またはダウンロード）する

   ```bash
   git clone https://github.com/your-username/paper2slide-gpt.git
   cd paper2slide-gpt
   ```

2. Python 仮想環境 (venv) を作成・有効化（推奨）

   ```bash
   python -m venv venv
   source venv/bin/activate  # macOS / Linux の場合
   venv\Scripts\activate     # Windows の場合
   ```

3. `requirements.txt` を使って必要パッケージをインストールする

   ```bash
   pip install -r requirements.txt
   ```

---

## 環境変数の設定 (.env ファイル)

1. `.env.template` を複製して `.env` というファイルを作成する

   ```bash
   cp .env.template .env
   ```
   > Windows では `copy .env.template .env` など

2. `.env` を編集し、以下のように **OpenAI の API キー** を設定する

   ```dotenv
   OPENAI_API_KEY="sk-XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
   ```
   > `OPENAI_API_KEY` の値は必ずご自身のキーに置き換えてください。

---

## アプリケーションの起動

1. 上記の手順で環境を整えたあと、ターミナル (またはコマンドプロンプト) で `main_gui.py` を実行します。

   ```bash
   python main_gui.py
   ```

2. GUI ウィンドウ（またはコンソール）が起動し、論文要約やスライド向けテキスト生成の機能を利用できます。

---

## ライセンス

このリポジトリのライセンスに関しては、`LICENSE` ファイルをご確認ください。  
（もしライセンスを設定していない場合は、追加で記載いただくことをおすすめします。）

---