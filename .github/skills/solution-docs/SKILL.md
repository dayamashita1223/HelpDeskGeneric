---
name: solution-docs
description: "Power Platform ソリューションの README.md と導入手順書（Word docx）をセットで生成する。README テンプレート・導入手順書テンプレート・画像プレースホルダー・docx 生成スクリプトを標準化。Use when: README, readme, 導入手順書, ドキュメント作成, ソリューション説明, 配布用ドキュメント, リポジトリ説明, Word, docx, インストール手順, 手順書テンプレート"
---

# ソリューションドキュメント生成スキル

Power Platform ソリューション配布時に必要な **README.md** と **導入手順書（Word docx）** をセットで生成する。

## 生成物

| ファイル                                   | 形式     | 用途                                   |
| ------------------------------------------ | -------- | -------------------------------------- |
| `README.md`                                | Markdown | GitHub リポジトリのトップページ        |
| `{アプリ名}_ソリューション導入手順書.docx` | Word     | 導入担当者向けの詳細手順書（画像入り） |

---

## Part 1: README.md テンプレート

### セクション構成（すべて必須）

```markdown
# {アプリ名}

## アプリケーション概要

{1〜3 文でアプリの目的・特徴を説明}

## キャプチャ

<!-- スクリーンショットを追加してください -->

![キャプチャ](docs/images/01_overview.png "概要画面")

## 構成

- README.md
- {SolutionName}\_X_X_X_X.zip：アンマネージドソリューション
- {SolutionName}\_X_X_X_X_managed.zip：マネージドソリューション
- docs/{アプリ名}\_ソリューション導入手順書.docx：導入手順書

## 展開・利用に必要な条件

- Power Apps Premium ライセンス（開発者・利用者）
- {追加ライセンス要件}

## 対応言語

- 日本語

## 主な機能

- {機能 1}
- {機能 2}

## アプリ利用に必要なコネクタ

- Dataverse
- {追加コネクタ}

## インストールに必要な権限

- システム管理者（セキュリティロール）

## インストールに必要なソリューション

- 特になし

## インストール方法

1. ソリューション ZIP をダウンロード
2. [Power Apps](https://make.powerapps.com/) にサインイン
3. 対象環境を選択 → ソリューション → インポート
4. ダウンロードした ZIP を選択 → 「次へ」をクリック
5. 接続の設定を求められたら、各コネクタの接続を選択または作成
6. 環境変数の値を設定（ある場合）
7. 「インポート」をクリックし、完了まで待機（数分〜10 分程度）
8. インポート完了後、「初期設定方法」に従って設定

詳細は導入手順書（docs/{アプリ名}\_ソリューション導入手順書.docx）を参照してください。

## 初期設定方法

- {初期設定項目の箇条書き}

## マネージドソリューションとアンマネージドソリューション

2種類のソリューションを用意しています。インストールにはマネージドソリューションを使用することをおすすめします。アンマネージドは開発環境への展開や、追加の開発・カスタマイズを実施する環境で展開してください。

## FAQ

- Q. 内容や機能をカスタマイズすることは可能ですか？
  - A. 可能です。カスタマイズすることを前提にシンプルで汎用的な作りになっています
- Q. 展開パートナーはどのように見つけることができますか？
  - A. 日本マイクロソフト営業担当者までお問い合わせください

## 免責事項

{免責事項テキスト（固定文。Part 3 参照）}

{年月}吉日
```

### README 記述ルール

- **日本語** で記述する
- ファイルは **リポジトリルート** に `README.md` として配置
- キャプチャ画像は `docs/images/` に配置し相対パスで参照
- 画像がない場合は `<!-- スクリーンショットを追加してください -->` コメントを残す
- FAQ と免責事項は **固定テキスト**（Part 3 の定型文を使用）

---

## Part 2: 導入手順書（Word docx）テンプレート

### 文書構造

導入手順書は以下の固定構造で構成する。**章番号・順序は変更しない。**

```
表紙ページ
  【サンプルアプリ】
  {アプリ名}
  導入手順書
  バージョン X.X.X.X　｜　YYYY年MM月

1. アプリ概要              ← Heading 1
  1.1 主な機能             ← Heading 2
  1.2 ソリューション構成   ← Heading 2（テーブル）
  1.3 必要な条件           ← Heading 2（テーブル）

2. インストール手順        ← Heading 1（★ 固定手順）
  2.1 Power Apps ポータルにサインイン
  2.2 環境の選択
  2.3 ソリューションのインポート
  2.4 接続の設定           ← 接続テーブル
  2.5 インポートの実行

3. 追加設定                ← Heading 1
  3.1 Power Automate フローの有効化  ← フローテーブル
    フローの接続修正       ← Heading 3
    通知先メールアドレスの設定 ← Heading 3（通知フローがある場合のみ）
  3.2 セキュリティロールの割当
  3.3 カテゴリの初期設定   ← マスタテーブルがある場合
  3.4 アプリの共有
  3.5 Copilot Studio エージェントの設定（オプション）← エージェントがある場合

4. 利用方法                ← Heading 1
  4.1 アプリへのアクセス
  4.2 ナビゲーション       ← SiteMap テーブル
  4.3〜 各機能の操作手順   ← アプリの主要機能ごとに小節を追加
  4.X Copilot Studio エージェントからの操作 ← エージェントがある場合

5. Appendix                ← Heading 1
  5.1 テーブル構成         ← 各テーブルの列定義テーブル
  5.2 Power Automate フロー一覧 ← フローテーブル
  5.3 Copilot Studio エージェント一覧 ← エージェントがある場合
  5.4 チャート一覧         ← チャートがある場合
```

### 第2章「インストール手順」は固定テンプレート

**第2章はソリューションに依存しない共通手順**であり、以下のテンプレートをそのまま使う。
変数部分（`{接続テーブル}` `{環境変数の説明}` 等）のみソリューション固有の値を埋める。

#### 2.1 Power Apps ポータルにサインイン

```
1. ブラウザで https://make.powerapps.com/ にアクセスします
2. 管理者アカウントでサインインします
```

**📸 スクリーンショット**: Power Apps ポータルのサインイン画面

#### 2.2 環境の選択

```
1. 画面右上の環境セレクターをクリックします
2. ソリューションをインポートする対象環境を選択します
```

> 💡 対象環境で Dataverse が有効になっていることを確認してください。

**📸 スクリーンショット**: 環境セレクター

#### 2.3 ソリューションのインポート

```
1. 左メニューの「ソリューション」をクリックします
2. 上部メニューの「インポート」をクリックします
3. 「ソリューションをインポート」ダイアログが表示されます
4. 「参照」ボタンをクリックし、ダウンロードしたソリューション ZIP ファイルを選択します
5. 「次へ」をクリックします
```

**📸 スクリーンショット**: ソリューションインポートダイアログ

#### 2.4 接続の設定

```
1. インポート中に接続の設定を求められる場合は、接続に対してアカウントを
   選択または新規作成してください。すべての接続を設定したら「次へ」をクリックします
```

**📸 スクリーンショット**: 接続設定画面

**接続テーブル**（ソリューション固有）:

| 接続                | 用途               | 設定方法                                       |
| ------------------- | ------------------ | ---------------------------------------------- |
| Microsoft Dataverse | テーブルの読み書き | 既存の接続を選択、またはサインインして新規作成 |
| {追加接続}          | {用途}             | 既存の接続を選択、またはサインインして新規作成 |

**環境変数**（ある場合）:

```
2. 環境変数の設定を求められる場合は以下を設定してください:
   - Dataverse URL: 自環境の URL（例: https://xxxxx.crm.dynamics.com）
   - {追加の環境変数}: {説明}
```

> 💡 アプリ ID 等のソリューション内で固定される値は変更不要です。

#### 2.5 インポートの実行

```
1. 設定内容を確認し、「インポート」をクリックします
2. インポートが開始されます。完了まで数分〜10 分程度かかります
3. ソリューション「{アプリ名}」が正常にインポートされました、と表示されたら完了です
```

> 💡 インポートに失敗した場合は、エラーメッセージを確認してください。接続の認証エラーが最も多い原因です。

**📸 スクリーンショット**: インポート完了画面

### スクリーンショットプレースホルダー

画像が未準備の場合、Word 文書内に以下のプレースホルダーテキストを挿入する:

```
【 {画面名} のスクリーンショットを挿入 】
```

標準的なスクリーンショット一覧:

| 挿入箇所               | プレースホルダーテキスト                                  |
| ---------------------- | --------------------------------------------------------- |
| 2.2 環境選択           | 【 環境セレクター のスクリーンショットを挿入 】           |
| 2.3 インポート         | 【 ソリューションインポート のスクリーンショットを挿入 】 |
| 2.4 接続設定           | 【 接続設定画面 のスクリーンショットを挿入 】             |
| 2.5 インポート完了     | 【 インポート完了 のスクリーンショットを挿入 】           |
| 3.1 フロー有効化       | 【 フローの有効化 のスクリーンショットを挿入 】           |
| 3.1 接続修正           | 【 接続エラーの修正 のスクリーンショットを挿入 】         |
| 3.2 セキュリティロール | 【 セキュリティロールの割当 のスクリーンショットを挿入 】 |
| 3.4 アプリ共有         | 【 アプリの共有画面 のスクリーンショットを挿入 】         |
| 4.X 各画面             | 【 {画面名} のスクリーンショットを挿入 】                 |

### テーブル定義（Appendix 用）

各 Dataverse テーブルのスキーマを以下の形式で記載する:

| 列名     | 論理名         | 型                                                        | 説明                                 |
| -------- | -------------- | --------------------------------------------------------- | ------------------------------------ |
| {表示名} | {logical_name} | テキスト / Choice / Lookup / 日時 / 整数 / 複数行テキスト | {説明。Lookup は → 参照先テーブル名} |

### Word スタイルマッピング

| 用途       | Word スタイル                  |
| ---------- | ------------------------------ |
| 章タイトル | Heading 1                      |
| 節タイトル | Heading 2                      |
| 項タイトル | Heading 3                      |
| 本文       | Normal                         |
| 手順番号   | List Number                    |
| 箇条書き   | List Bullet                    |
| 補足（💡） | Normal（段落先頭に 💡 絵文字） |

---

## Part 3: 定型テキスト（固定・変更不可）

### FAQ（README / 手順書 共通）

```
* Q. 内容や機能をカスタマイズすることは可能ですか？
    * A. 可能です。カスタマイズすることを前提にシンプルで汎用的な作りになっています
* Q. 展開パートナーはどのように見つけることができますか？
    * A. 日本マイクロソフト営業担当者までお問い合わせください
```

### 免責事項（README / 手順書 共通。日付のみ更新）

```
本アプリ集は日本マイクロソフトが提供する無償のサンプル群です。本アプリ集をダウンロードされた方は、以下の免責事項を承諾したものとみなされます。

1.　本アプリ集（本アプリ集に付属するドキュメント及びReadmeに記載されている技術情報を含みます。以下、本「免責事項」において同じ。）は利用者に対して「現状のまま」提供されるものであり、日本マイクロソフトは、本アプリ集にプログラミング上の誤りその他の瑕疵のないこと、本アプリ集が利用者の目的に適合すること、並びに本アプリ集及びその使用が利用者または利用者以外の第三者の権利を侵害するものでないこと、その他のいかなる内容についての明示または黙示の保証を行うものではありません。

2.　日本マイクロソフトは、本アプリ集の使用に起因して、利用者に生じた損害または第三者からの請求に基づく利用者の損害について、原因の如何を問わず、一切の責任を負いません。日本マイクロソフトは、本アプリ集に関連して利用者と第三者との間に発生するいかなる紛争について、一切責任を負わないものとします。本アプリ集の利用は、利用者の責任のもとで行ってください。

3.　日本マイクロソフトは、本アプリ集の全部または一部の提供を廃止することがあります。提供の廃止によって利用者に発生した損害について、日本マイクロソフトは一切責任を負いません。

4.　日本マイクロソフトは、本アプリ集のバグ修正、補修、保守、機能追加その他のいかなる義務も負いません。本アプリ集は、不定期に更新される可能性がありますが、バグ修正等が保証されているわけではありません。本アプリ集の安定した動作を確保するためには、利用者自身が適切なテストや検証を行ってください。

5.　日本マイクロソフトは、本アプリ集に関するお問い合わせにはお答えできません。ご利用にあたっては、提供された手順書を参照し、ご自身でのインストールや利用を行ってください。

{YYYY}年{M}月吉日
```

---

## Part 4: 情報収集方法

### Dataverse API から自動取得

```python
# ソリューション基本情報
solutions?$filter=uniquename eq '{SOLUTION_NAME}'
  → friendlyname, version, publisherid

# テーブル一覧
solutioncomponents?$filter=_solutionid_value eq {sol_id} and componenttype eq 1
  → EntityDefinitions で表示名・論理名・列定義を取得

# フロー一覧
solutioncomponents?$filter=componenttype eq 29
  → workflows で名前・種類・トリガーを取得

# アプリ一覧
solutioncomponents?$filter=componenttype eq 80
  → appmodules で名前・説明を取得

# チャート一覧
savedqueryvisualizations?$filter=primaryentitytypecode eq '{entity}'

# 接続参照
connectionreferences → コネクタ一覧

# Copilot Studio エージェント
bots → エージェント名・スキーマ名

# 環境変数
environmentvariabledefinitions → 変数名・型・説明
```

### ユーザーに確認する項目

| 項目                 | 例                                         |
| -------------------- | ------------------------------------------ |
| アプリの概要説明     | 「社内ヘルプデスクの問い合わせ管理アプリ」 |
| キャプチャ画像       | スクリーンショットのファイルパス           |
| 通知先の設定方法     | ヘルプデスクメーリングリスト等             |
| マスタデータの推奨値 | カテゴリ構成の推奨例                       |

---

## Part 5: docx 生成スクリプトテンプレート

導入手順書の Word ファイルは `python-docx` で生成する。以下がスクリプトの骨格。

### 依存パッケージ

```
pip install python-docx
```

### スクリプト骨格

```python
"""
{アプリ名} 導入手順書 生成スクリプト
"""
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# ============================================================
# 設定値（ソリューションごとに変更する部分）
# ============================================================
APP_NAME = "{アプリ名}"
SOLUTION_VERSION = "1.0.0.0"
DOC_DATE = "YYYY年MM月"
OUTPUT_FILE = f"{APP_NAME}_ソリューション導入手順書.docx"

# ソリューション構成テーブル
COMPONENTS = [
    ("テーブル",              "N",  "{テーブル名一覧}"),
    ("モデル駆動型アプリ",     "1",  "{アプリ名}"),
    ("Power Automate フロー", "N",  "{フロー名一覧}"),
    # 必要に応じて追加
]

# 必要な条件テーブル
PREREQUISITES = [
    ("Power Platform 環境", "Dataverse が有効であること"),
    ("ライセンス",           "Power Apps Premium ライセンス（開発者・利用者）"),
    ("セキュリティロール",    "System Administrator（インストール実施者）"),
    ("ブラウザ",             "Microsoft Edge / Google Chrome（最新版）"),
]

# 接続テーブル
CONNECTIONS = [
    ("Microsoft Dataverse",      "テーブルの読み書き"),
    ("Office 365 Outlook",       "通知メール送信"),
]

# フローテーブル
FLOWS = [
    ("フロー名", "動作", "確認事項"),
]

# ============================================================
# ドキュメント生成
# ============================================================
doc = Document()

# --- 表紙 ---
for _ in range(6):
    doc.add_paragraph("")
p = doc.add_paragraph("【サンプルアプリ】")
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p = doc.add_paragraph(APP_NAME)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
for run in p.runs:
    run.font.size = Pt(28)
    run.bold = True
p = doc.add_paragraph("導入手順書")
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
for run in p.runs:
    run.font.size = Pt(18)
doc.add_paragraph("")
p = doc.add_paragraph(f"バージョン {SOLUTION_VERSION}　｜　{DOC_DATE}")
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.add_page_break()

# --- 目次 ---
doc.add_heading("目次", level=1)
toc_items = [
    "1. アプリ概要",
    "2. インストール手順",
    "3. 追加設定",
    "4. 利用方法",
    "5. Appendix",
]
for item in toc_items:
    doc.add_paragraph(item)
doc.add_page_break()


def add_table(doc, headers, rows, col_widths=None):
    """ヘッダー行 + データ行のテーブルを追加するヘルパー"""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    # ヘッダー
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.bold = True
    # データ
    for ri, row_data in enumerate(rows):
        for ci, val in enumerate(row_data):
            table.rows[ri + 1].cells[ci].text = str(val)
    return table


def add_screenshot_placeholder(doc, label):
    """スクリーンショットプレースホルダーを追加"""
    p = doc.add_paragraph(f"【 {label} のスクリーンショットを挿入 】")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in p.runs:
        run.font.color.rgb = RGBColor(128, 128, 128)
        run.italic = True


def add_tip(doc, text):
    """💡 補足テキストを追加"""
    doc.add_paragraph(f"💡 {text}")


# --- 1. アプリ概要 ---
doc.add_heading("1. アプリ概要", level=1)
doc.add_paragraph("{アプリの概要説明}")

doc.add_heading("1.1 主な機能", level=2)
features = [
    "{機能 1}",
    "{機能 2}",
]
for f in features:
    doc.add_paragraph(f, style="List Bullet")

doc.add_heading("1.2 ソリューション構成", level=2)
add_table(doc, ["コンポーネント", "数量", "内容"], COMPONENTS)

doc.add_heading("1.3 必要な条件", level=2)
add_table(doc, ["要件", "詳細"], PREREQUISITES)


# --- 2. インストール手順（★ 固定テンプレート） ---
doc.add_heading("2. インストール手順", level=1)
doc.add_paragraph(
    "ソリューションファイル（ZIP）を Power Apps 環境にインポートします。"
    "以下の手順に従ってください。"
)

doc.add_heading("2.1 Power Apps ポータルにサインイン", level=2)
for step in [
    "ブラウザで https://make.powerapps.com/ にアクセスします",
    "管理者アカウントでサインインします",
]:
    doc.add_paragraph(step, style="List Number")

doc.add_heading("2.2 環境の選択", level=2)
for step in [
    "画面右上の環境セレクターをクリックします",
    "ソリューションをインポートする対象環境を選択します",
]:
    doc.add_paragraph(step, style="List Number")
add_screenshot_placeholder(doc, "環境セレクター")
add_tip(doc, "対象環境で Dataverse が有効になっていることを確認してください。")

doc.add_heading("2.3 ソリューションのインポート", level=2)
for step in [
    "左メニューの「ソリューション」をクリックします",
    "上部メニューの「インポート」をクリックします",
    "「ソリューションをインポート」ダイアログが表示されます",
    "「参照」ボタンをクリックし、ダウンロードしたソリューション ZIP ファイルを選択します",
    "「次へ」をクリックします",
]:
    doc.add_paragraph(step, style="List Number")
add_screenshot_placeholder(doc, "ソリューションインポート")

doc.add_heading("2.4 接続の設定", level=2)
doc.add_paragraph(
    "インポート中に接続の設定を求められる場合は、接続に対してアカウントを"
    "選択または新規作成してください。すべての接続を設定したら「次へ」を"
    "クリックします",
    style="List Number",
)
add_screenshot_placeholder(doc, "接続設定画面")
add_table(
    doc,
    ["接続", "用途", "設定方法"],
    [(c[0], c[1], "既存の接続を選択、またはサインインして新規作成") for c in CONNECTIONS],
)
# ★ 環境変数がある場合はここに追加
doc.add_paragraph(
    "環境変数の設定を求められる場合は自環境の URL"
    "（例: https://xxxxx.crm.dynamics.com）を設定してください。"
    "アプリIDは固有のため、変更は不要です。",
    style="List Number",
)

doc.add_heading("2.5 インポートの実行", level=2)
for step in [
    "設定内容を確認し、「インポート」をクリックします",
    "インポートが開始されます。完了まで数分〜10 分程度かかります",
    f"ソリューション「{APP_NAME}」が正常にインポートされました、と表示されたら完了です",
]:
    doc.add_paragraph(step, style="List Number")
add_tip(doc, "インポートに失敗した場合は、エラーメッセージを確認してください。接続の認証エラーが最も多い原因です。")
add_screenshot_placeholder(doc, "インポート完了")


# --- 3. 追加設定 ---
doc.add_heading("3. 追加設定", level=1)
doc.add_paragraph("ソリューションインポート後、以下の設定を行ってください。")

doc.add_heading("3.1 Power Automate フローの有効化", level=2)
doc.add_paragraph("インポート後、フローが無効（Draft）状態になっている場合があります。")
for step in [
    "https://make.powerautomate.com/ にサインインします",
    f"左メニューの「ソリューション」→「{APP_NAME}」を開きます",
    "以下のフローを確認し、無効の場合は「有効にする」をクリックします",
]:
    doc.add_paragraph(step, style="List Number")
add_table(doc, ["フロー名", "動作", "確認事項"], FLOWS)
add_screenshot_placeholder(doc, "フローの有効化")

doc.add_heading("フローの接続修正", level=3)
doc.add_paragraph("フローが「接続エラー」で失敗する場合は、以下の手順で接続を修正します。")
for step in [
    "対象フローを開き「編集」をクリックします",
    "赤い警告が表示されている接続をクリックします",
    "「接続の修正」で正しいアカウントを選択します",
    "「保存」をクリックします",
]:
    doc.add_paragraph(step, style="List Number")
add_screenshot_placeholder(doc, "接続エラーの修正")

# ★ 通知先設定（通知フローがある場合のみ追加）
# doc.add_heading("通知先メールアドレスの設定", level=3)
# ...

doc.add_heading("3.2 セキュリティロールの割当", level=2)
for step in [
    "Power Platform 管理センター > 環境 > 対象環境 > 設定 > ユーザー＋アクセス許可 > セキュリティロール を開きます",
    "「Basic User」ロールがアプリに関連付けられていることを確認します",
    "アプリ利用者にセキュリティロールを割り当てます",
]:
    doc.add_paragraph(step, style="List Number")
add_tip(doc, "セキュリティグループ単位で一括割当すると効率的です。")
add_screenshot_placeholder(doc, "セキュリティロールの割当")

# ★ マスタデータ初期設定（マスタテーブルがある場合のみ追加）
# doc.add_heading("3.3 カテゴリの初期設定", level=2)
# ...

doc.add_heading("3.4 アプリの共有", level=2)
for step in [
    "make.powerapps.com > アプリ一覧を開きます",
    f"「{APP_NAME}」の「…」メニュー > 「共有」をクリックします",
    "利用者（ユーザーまたはセキュリティグループ）を追加します",
    "適切なセキュリティロールを選択します",
]:
    doc.add_paragraph(step, style="List Number")
add_screenshot_placeholder(doc, "アプリの共有画面")

# ★ Copilot Studio（エージェントがある場合のみ追加）
# doc.add_heading("3.5 Copilot Studio エージェントの設定（オプション）", level=2)
# ...


# --- 4. 利用方法 ---
doc.add_heading("4. 利用方法", level=1)

doc.add_heading("4.1 アプリへのアクセス", level=2)
for step in [
    "ブラウザで Power Apps ポータル（https://make.powerapps.com/）を開きます",
    "左メニューの「アプリ」をクリックします",
    f"「{APP_NAME}」を選択して起動します",
]:
    doc.add_paragraph(step, style="List Number")

doc.add_heading("4.2 ナビゲーション", level=2)
doc.add_paragraph("アプリの左サイドバーには以下のグループがあります。")
# ★ SiteMap に応じたナビゲーションテーブルを追加

# ★ 4.3〜 各機能の操作手順をソリューション固有で追加


# --- 5. Appendix ---
doc.add_heading("5. Appendix", level=1)

doc.add_heading("5.1 テーブル構成", level=2)
# ★ 各テーブルの列定義テーブルを追加

doc.add_heading("5.2 Power Automate フロー一覧", level=2)
# ★ フローテーブルを追加

# ★ 5.3, 5.4 等 ソリューション固有で追加


# --- 保存 ---
doc.save(OUTPUT_FILE)
print(f"✅ {OUTPUT_FILE} を生成しました")
```

### スクリプトのカスタマイズポイント

| 箇所                      | 変更内容                                                       |
| ------------------------- | -------------------------------------------------------------- |
| `APP_NAME`                | ソリューションのアプリ名                                       |
| `COMPONENTS`              | ソリューション構成テーブルの行                                 |
| `PREREQUISITES`           | 追加の前提条件（Azure OpenAI 等）                              |
| `CONNECTIONS`             | 使用するコネクタ（Copilot Studio 等を追加）                    |
| `FLOWS`                   | フロー名・動作・確認事項                                       |
| 第3章のコメントアウト部分 | マスタデータ・エージェント・通知先等、ソリューション固有の設定 |
| 第4章                     | アプリの操作説明をソリューション固有で追記                     |
| 第5章                     | テーブル定義・フロー・エージェント・チャートを追記             |

---

## Part 6: 生成手順（ワークフロー）

### Step 1: 情報収集

1. Dataverse API でソリューション構成を自動取得（テーブル・フロー・アプリ・チャート・接続・エージェント・環境変数）
2. ユーザーに不足情報を確認（概要説明、キャプチャ画像、マスタデータ推奨値）

### Step 2: README.md 生成

1. Part 1 テンプレートに沿って README.md を生成
2. ユーザーにレビューを依頼

### Step 3: 導入手順書（docx）生成

1. Part 5 のスクリプト骨格をベースに、ソリューション固有の値を埋める
2. 第2章はテンプレートのまま使用（接続テーブル・環境変数のみ変更）
3. 第3章・第4章・第5章にソリューション固有のコンテンツを追加
4. スクリプトを実行して docx を生成
5. ユーザーにレビューを依頼（スクリーンショット挿入は手動）

### Step 4: 配布パッケージ構成

```
{リポジトリ}/
├── README.md
├── {SolutionName}_X_X_X_X.zip          ← アンマネージド
├── {SolutionName}_X_X_X_X_managed.zip  ← マネージド
└── docs/
    ├── {アプリ名}_ソリューション導入手順書.docx
    └── images/                          ← スクリーンショット（任意）
```

---

## 注意事項

- README と手順書で **情報の整合性を保つ**（機能一覧・接続・前提条件等が一致すること）
- 手順書の第2章（インストール手順）は **全ソリューション共通**。ソリューション固有の変更は接続テーブルと環境変数のみ
- スクリーンショットは手順書作成後に **ユーザーが手動で挿入**。プレースホルダーで位置を示す
- `yaml.dump()` 等のシリアライザで docx を生成しない。`python-docx` API を直接使用する
- 画像は PNG 形式を推奨（Word での表示互換性が最も高い）
