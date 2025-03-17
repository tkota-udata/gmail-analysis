# gmail-analysis

## 概要

Gmail分析ツールは、特定の送信者からのメール送信パターンを分析し、視覚的なレポートを生成するPythonツールです。時間帯別、曜日別、月別の送信傾向を分析し、効果的なメール戦略のための考察を提供します。

## 機能

- Gmail APIを使用して特定の送信者からのメールデータを取得
- 時間帯別、曜日別、月別の送信分布を分析
- 時間帯×曜日のヒートマップ作成
- Claude AIを活用した送信パターンの考察と改善提案
- すべての分析結果を1ページのPDFレポートとして出力

## 前提条件

- Python 3.7以上
- Gmail APIへのアクセス権限
- Anthropic API（Claude）のアクセスキー（オプション）

## インストール方法

1. リポジトリをクローンまたはダウンロードします：

```bash
git clone https://github.com/yourusername/gmail-analysis.git
cd gmail-analysis
```

2. 必要なパッケージをインストールします：

```bash
pip install -r requirements.txt
```

3. Gmail APIの認証情報を設定します：
   - [Google Cloud Console](https://console.cloud.google.com/)で新しいプロジェクトを作成
   - Gmail APIを有効化
   - OAuth 2.0クライアントIDを作成し、認証情報をダウンロード
   - ダウンロードしたJSONファイルを`credentials.json`として保存

4. （オプション）Claude APIのキーを環境変数に設定します：

```bash
export ANTHROPIC_API_KEY=your_api_key_here
```

## 使用方法

### 基本的な使い方

```python
from gmail_analyzer import GmailAnalyzer

# 分析ツールの初期化
analyzer = GmailAnalyzer()

# 認証（初回のみブラウザが開きGoogleアカウントへのアクセスを許可する必要があります）
analyzer.authenticate()

# 特定の送信者のメールを分析（例：example@gmail.com）
sender_email = "example@gmail.com"
df = analyzer.fetch_emails_from_sender(sender_email, max_results=500)

# 分析レポートを生成
report_path = analyzer.generate_comprehensive_pdf_report(df, sender_email)
print(f"レポートが生成されました: {report_path}")
```

### コマンドラインからの実行

スクリプトをコマンドラインから直接実行することもできます：

```bash
python gmail_analyzer.py --sender example@gmail.com --max-results 500
```

オプション：
- `--sender`: 分析対象の送信者メールアドレス（必須）
- `--max-results`: 取得するメールの最大数（デフォルト: 500）
- `--output`: 出力PDFファイルのパス（省略可）

## レポートの内容

生成されるPDFレポートには以下の情報が含まれます：

1. **基本情報**：送信者、分析期間、総メール数
2. **時間帯別分布**：24時間の送信傾向をグラフ化
3. **曜日別分布**：曜日ごとの送信数をグラフ化
4. **月別分布**：月ごとの送信数をグラフ化
5. **時間帯×曜日ヒートマップ**：最も送信が多い時間帯と曜日の組み合わせを視覚化
6. **データに基づく考察**：送信パターンの分析と改善提案

## 日本語フォントの設定

日本語を含むレポートを正しく表示するには、以下のいずれかの場所に日本語フォントファイルを配置してください：

- `fonts/ipaexg.ttf`（プロジェクトディレクトリ内）
- `/usr/share/fonts/truetype/ipafont/ipag.ttf`（Linux）
- `/Library/Fonts/Arial Unicode.ttf`（macOS）
- `C:\Windows\Fonts\msgothic.ttc`（Windows）

フォントが見つからない場合、レポートは英語で生成されます。

## Claude APIの活用

Claude APIを使用すると、より詳細で個別化された考察が得られます。APIキーを設定しない場合は、デフォルトの考察が使用されます。

APIキーの設定方法：
1. [Anthropic](https://www.anthropic.com/)でAPIキーを取得
2. 環境変数に設定：`export ANTHROPIC_API_KEY=your_api_key_here`

## トラブルシューティング

### 認証エラー
- `credentials.json`ファイルが正しく配置されているか確認してください
- 初回認証後に生成される`token.json`を削除して再認証を試みてください

### PDFエラー
- 日本語フォントが正しく設定されているか確認してください
- FPDFライブラリのバージョンが最新であることを確認してください

### APIレート制限
- Gmail APIには1日あたりの使用制限があります。大量のメールを分析する場合は、複数回に分けて実行してください

## 高度な使用例

### 複数の送信者の分析

```python
senders = ["sender1@gmail.com", "sender2@gmail.com", "sender3@gmail.com"]
for sender in senders:
    df = analyzer.fetch_emails_from_sender(sender, max_results=300)
    analyzer.generate_comprehensive_pdf_report(df, sender)
```

### 特定期間のメール分析

```python
import datetime

# 日付範囲を指定してメールを取得
start_date = datetime.datetime(2023, 1, 1)
end_date = datetime.datetime(2023, 12, 31)
df = analyzer.fetch_emails_from_sender(
    "example@gmail.com", 
    max_results=1000,
    start_date=start_date,
    end_date=end_date
)
```

### カスタムレポート名の指定

```python
output_path = "reports/analysis_report_2023.pdf"
analyzer.generate_comprehensive_pdf_report(df, "example@gmail.com", output_path=output_path)
```

## ライセンス

このプロジェクトはMITライセンスの下で公開されています。詳細はLICENSEファイルを参照してください。

## 謝辞

- Google Gmail API
- Anthropic Claude API
- matplotlib, pandas, numpy
- FPDF

---

ご質問やフィードバックがありましたら、Issueを作成するか、プルリクエストを送信してください。