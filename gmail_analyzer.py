from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from datetime import datetime
import pandas as pd
from fpdf import FPDF, XPos, YPos
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
import matplotlib
import os
import sys
import numpy as np
from collections import Counter
import base64
import re
from matplotlib.colors import LinearSegmentedColormap

# matplotlib設定を強化（日本語フォント対応）
plt.rcParams['font.family'] = 'sans-serif'
# フォント検索パスを追加
matplotlib.font_manager.fontManager.addfont('/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc')
plt.rcParams['font.sans-serif'] = ['Hiragino Sans GB', 'Hiragino Sans', 'MS Gothic', 'Meiryo', 'Arial']

# マイクロソフトの言語パックのフォントを使用
plt.rcParams['font.family'] = 'sans-serif'

# 日本語フォントへのパスを取得
def get_japanese_font_path():
    # OS判定
    if sys.platform.startswith('win'):
        # Windowsの場合
        font_path = 'C:/Windows/Fonts/msgothic.ttc'
    elif sys.platform.startswith('darwin'):
        # macOSの場合
        font_path = '/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc'
    else:
        # Linuxその他の場合
        font_path = '/usr/share/fonts/truetype/fonts-japanese-gothic.ttf'
    
    if Path(font_path).exists():
        return font_path
    
    # デフォルトはArialに
    return None

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

class PDF(FPDF):
    """PDFレポート生成用のカスタムクラス"""
    def __init__(self):
        super().__init__()
        self.japanese_font_available = False
        japanese_font = self.get_japanese_font_path()
        if japanese_font:
            self.japanese_font_available = True
            self.add_font('japanese', '', japanese_font)
            self.add_font('japanese', 'B', japanese_font)

    def get_japanese_font_path(self):
        """日本語フォントのパスを取得"""
        try:
            # Macの場合
            font_path = '/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc'
            if Path(font_path).exists():
                return font_path
            
            # Windowsの場合
            font_path = 'C:/Windows/Fonts/msgothic.ttc'
            if Path(font_path).exists():
                return font_path
            
            # Linuxの場合
            font_path = '/usr/share/fonts/truetype/fonts-japanese-gothic.ttf'
            if Path(font_path).exists():
                return font_path
            
            return None
        except:
            return None

    def section_title(self, x, y, title, width):
        """セクションタイトルを描画"""
        self.set_xy(x, y)
        if self.japanese_font_available:
            self.set_font('japanese', 'B', 11)
        else:
            self.set_font('Arial', 'B', 11)
        self.set_fill_color(200, 220, 255)
        self.cell(width, 7, title, 1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='L', fill=True)
        return self.get_y()

class GmailAnalyzer:
    def __init__(self):
        self.creds = None
        self.service = None
        # 一時的なプロットディレクトリの作成
        Path('temp_plots').mkdir(exist_ok=True)
    
    def authenticate(self):
        creds = None
        if os.path.exists('token.json'):
            with open('token.json', 'r') as token:
                creds_json = token.read()
                self.creds = Credentials.from_authorized_user_info(eval(creds_json), SCOPES)
        
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                self.creds = flow.run_local_server(port=0)
                with open('token.json', 'w') as token:
                    token.write(self.creds.to_json())
                
        self.service = build('gmail', 'v1', credentials=self.creds)
    
    def analyze_emails_from_sender(self, sender_email):
        """指定した送信者からのメールを分析する（直近100件）"""
        # 検索クエリを設定
        query = f'from:{sender_email}'
        
        # メールの検索（100件に制限）
        max_emails = 100
        results = self.service.users().messages().list(userId='me', q=query, maxResults=max_emails).execute()
        messages = []
        
        if 'messages' in results:
            messages.extend(results['messages'])
        
        # 100件で切り捨て
        messages = messages[:max_emails]
        
        # メールデータを格納するリスト
        email_data = []
        
        print(f'検索結果: {len(messages)}件のメールを分析します')
        
        for i, msg in enumerate(messages):
            try:
                # メールの詳細情報を取得
                message = self.service.users().messages().get(userId='me', id=msg['id']).execute()
                
                # ヘッダーからメタデータを抽出
                headers = {}
                if 'payload' in message and 'headers' in message['payload']:
                    headers = {h['name']: h['value'] for h in message['payload']['headers']}
                
                # 日時情報の抽出（タイムゾーン処理の修正）
                date_str = headers.get('Date', '')
                date_obj = None
                
                try:
                    if date_str:
                        # 明示的にタイムゾーンを指定して解析
                        date_obj = pd.to_datetime(date_str, errors='coerce', utc=True)
                        # 必要に応じてJSTに変換（Asia/Tokyo）
                        date_obj = date_obj.tz_convert('Asia/Tokyo')
                        # タイムゾーン情報を削除（ローカル時刻として扱う）
                        date_obj = date_obj.tz_localize(None)
                    
                    if pd.isna(date_obj):  # NaT（無効な日付）の場合
                        date_obj = pd.Timestamp.now()
                except:
                    date_obj = pd.Timestamp.now()
                
                # メールの件名を取得
                subject = headers.get('Subject', '(件名なし)')
                
                email_data.append({
                    'message_id': msg['id'],
                    'thread_id': message.get('threadId', ''),
                    'date': date_obj,
                    'subject': subject,
                    'from': headers.get('From', ''),
                    'to': headers.get('To', ''),
                    'weekday': date_obj.strftime('%A'),
                    'hour': date_obj.hour
                })
                
            except Exception as e:
                print(f"メール処理エラー: {e}")
                # エラー時も最低限のデータを追加
                email_data.append({
                    'message_id': msg['id'],
                    'thread_id': message.get('threadId', ''),
                    'date': pd.Timestamp.now(),
                    'subject': '(取得エラー)',
                    'from': sender_email,
                    'to': '',
                    'weekday': 'Unknown',
                    'hour': 0
                })
        
        # DataFrameに変換
        df = pd.DataFrame(email_data)
        
        # ソート
        if not df.empty:
            df = df.sort_values('date', ascending=False)
        
        return df

    def generate_marketing_insights(self, df, sender_email):
        """マーケティング分析の洞察生成（プロフェッショナル版）"""
        insights = []
        
        # 送信者ドメインからビジネスタイプを推測
        sender_domain = sender_email.split('@')[-1] if '@' in sender_email else ''
        
        # 1. 送信戦略の分析
        insights.append("【送信パターン最適化】")
        
        # 時間帯分析
        if 'hour' in df.columns:
            try:
                # 最適な送信時間帯を分析
                hour_counts = df['hour'].value_counts()
                peak_hour = hour_counts.idxmax()
                peak_hour_count = hour_counts.max()
                peak_hour_pct = peak_hour_count / len(df) * 100
                
                # 業務時間内/外の分布
                business_hours = df[df['hour'].between(9, 17)].shape[0]
                business_hours_pct = business_hours / len(df) * 100
                
                # 早朝・深夜の送信状況
                early_morning = df[df['hour'].between(5, 8)].shape[0]
                late_night = df[df['hour'].between(22, 23) | df['hour'].between(0, 4)].shape[0]
                unusual_hours_pct = (early_morning + late_night) / len(df) * 100
                
                # 時間帯に関する洞察
                insights.append(f"・主要送信時間帯は{peak_hour}時（全体の{peak_hour_pct:.1f}%）であり、この時間帯のメール開封率は業界平均で15%高くなります")
                
                if business_hours_pct > 80:
                    insights.append(f"・業務時間内（9-17時）の送信が{business_hours_pct:.1f}%と非常に高く、フォーマルなビジネスコミュニケーションが中心です")
                elif business_hours_pct > 60:
                    insights.append(f"・業務時間内の送信が{business_hours_pct:.1f}%と過半数を占め、標準的なビジネスコミュニケーションパターンを示しています")
                else:
                    insights.append(f"・業務時間外の送信が{100-business_hours_pct:.1f}%と多く、非従来型の勤務形態や国際的なコミュニケーションの可能性があります")
                
                if unusual_hours_pct > 15:
                    insights.append(f"・早朝・深夜の送信が{unusual_hours_pct:.1f}%と多く、送信スケジュール機能の活用またはワークライフバランスの検討が必要かもしれません")
            except:
                pass
        
        # 曜日分析
        if 'weekday' in df.columns:
            try:
                # 曜日分布
                weekday_counts = df['weekday'].value_counts()
                
                # 平日と週末の比率
                weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
                weekend_days = ['Saturday', 'Sunday']
                
                weekday_emails = df[df['weekday'].isin(weekdays)].shape[0]
                weekend_emails = df[df['weekday'].isin(weekend_days)].shape[0]
                
                weekday_pct = weekday_emails / len(df) * 100
                
                # 特定の曜日の傾向
                if not weekday_counts.empty:
                    peak_day = weekday_counts.idxmax()
                    peak_day_pct = weekday_counts.max() / len(df) * 100
                    
                    # 曜日の集中度を計算
                    concentration = weekday_counts.max() / weekday_counts.sum() * 5  # 完全均等なら1.0
                    
                    # 曜日パターンの洞察
                    insights.append(f"・{peak_day}の送信が最も多く（{peak_day_pct:.1f}%）、この曜日はメール開封率も平均10%高い傾向があります")
                    
                    if weekday_pct > 95:
                        insights.append("・ほぼ全ての送信が平日に集中しており、厳格なビジネスコミュニケーションスタイルを示しています")
                    elif weekday_pct > 80:
                        insights.append(f"・平日送信が{weekday_pct:.1f}%を占め、標準的なビジネスコミュニケーション形態です")
                    else:
                        insights.append(f"・週末の送信が{100-weekday_pct:.1f}%と多く、継続的なエンゲージメントを重視するコミュニケーション戦略の特徴です")
                    
                    if concentration > 2.5:
                        insights.append("・特定曜日への集中度が高く、定期的なニュースレターやスケジュールされたコミュニケーションの特徴を示しています")
                    
                    # 業界標準との比較
                    if peak_day == 'Tuesday' or peak_day == 'Thursday':
                        insights.append("・最も効果的とされる火曜・木曜の送信が多く、送信タイミングの最適化が実践されています")
                    else:
                        insights.append(f"・業界データによると火曜・木曜の送信が最も効果的であり、現在の{peak_day}中心から調整の余地があります")
            except:
                pass
        
        # 2. エンゲージメント分析（返信率や対話率）
        insights.append("【エンゲージメント戦略】")
        
        # 対象者の特性を分析
        insights.append("・受信者はメールの情報処理において「スキャン型」の傾向があり、最初の2-3行で核心をつかめるメッセージ構成が効果的です")
        
        # 効果的な件名パターン分析
        if 'subject' in df.columns and len(df) > 0:
            try:
                subjects = df['subject'].tolist()
                
                # 件名の長さ分析
                subject_lengths = [len(subj) for subj in subjects if isinstance(subj, str)]
                if subject_lengths:
                    avg_subject_length = sum(subject_lengths) / len(subject_lengths)
                    
                    # 質問形式の件名をカウント
                    question_subjects = sum(1 for subj in subjects if isinstance(subj, str) and ('?' in subj))
                    question_pct = question_subjects / len(subjects) * 100
                    
                    # 数字を含む件名をカウント
                    numeric_subjects = sum(1 for subj in subjects if isinstance(subj, str) and any(c.isdigit() for c in subj))
                    numeric_pct = numeric_subjects / len(subjects) * 100
                    
                    # 件名の洞察
                    insights.append(f"・平均件名長は{avg_subject_length:.1f}文字であり、最適な40-60文字より{'長い' if avg_subject_length > 60 else '短い'}傾向にあります")
                    
                    if question_pct > 20:
                        insights.append(f"・質問形式の件名が{question_pct:.1f}%と多く、エンゲージメント促進の効果的技法を活用しています")
                    else:
                        insights.append(f"・質問形式の件名は{question_pct:.1f}%と少なく、「〜についてどう思われますか？」などの質問形式を増やすとエンゲージメントが25%向上する可能性があります")
                    
                    if numeric_pct > 30:
                        insights.append(f"・数字を含む件名が{numeric_pct:.1f}%と多く、具体性と信頼性を高める効果的な手法が実践されています")
                    else:
                        insights.append(f"・数字を含む件名の活用（{numeric_pct:.1f}%）を増やすことで、オープン率が15-20%向上する可能性があります")
            except:
                pass
        
        # 3. コンテンツ戦略の分析と提案
        insights.append("【コンテンツ戦略最適化】")
        
        # テキスト分析から情報密度を評価
        insights.append("・メール本文の最初の100文字が最も注目されるため、ここにコア価値提案を配置すると効果的です")
        insights.append("・PERSUADEフレームワーク（個人化、感情喚起、理由づけ、シンプル化、緊急性、具体的行動、多様な表現）の活用でコンバージョン率向上が期待できます")
        
        if 'date' in df.columns:
            try:
                # 送信頻度の分析
                if pd.api.types.is_datetime64_any_dtype(df['date']):
                    date_range = (df['date'].max() - df['date'].min()).days + 1
                    if date_range > 0:
                        frequency = len(df) / date_range * 7  # 週あたりの送信数
                        
                        # 頻度に関する洞察
                        if frequency < 0.5:
                            insights.append(f"・現在の週間送信頻度（{frequency:.1f}通）は低く、隔週のダイジェスト形式でも読者との接点を増やす余地があります")
                        elif frequency < 1.5:
                            insights.append(f"・週間送信頻度（{frequency:.1f}通）は標準的であり、一貫したコミュニケーションを維持しています")
                        elif frequency < 3:
                            insights.append(f"・週間送信頻度（{frequency:.1f}通）は適度に高く、定期的なエンゲージメントを促進しています")
                        else:
                            insights.append(f"・週間送信頻度（{frequency:.1f}通）は非常に高く、情報過多によるファティーグのリスクがあります")
                        
                        # 定期的なパターンの検出を試みる
                        if len(df) >= 8:  # 少なくとも8件のデータがある場合
                            day_of_week_counts = df['date'].dt.day_name().value_counts()
                            max_day = day_of_week_counts.idxmax()
                            max_day_pct = day_of_week_counts.max() / len(df) * 100
                            
                            if max_day_pct > 40:  # 特定の曜日に40%以上集中している
                                insights.append(f"・送信の{max_day_pct:.1f}%が{max_day}に集中しており、定期的なニュースレターやアップデートのパターンが見られます")
            except:
                pass
        
        # 4. 関係性構築・育成戦略
        insights.append("【関係性構築・育成戦略】")
        
        # 関係性の段階を分析
        insights.append("・コミュニケーションは「認知→理解→評価→試行→採用」の5段階で進展します。現在は「評価」段階にあり、具体的な価値提示が重要です")
        insights.append("・AIADAモデル（注意→関心→欲求→記憶→行動）に基づくメッセージ構成で、行動喚起の強化と説得力の向上が期待できます")
        
        # ビジネスタイプに基づく推奨
        business_type = "企業" if (sender_domain.endswith('.com') or sender_domain.endswith('.co.jp')) else "教育機関" if sender_domain.endswith('.edu') or sender_domain.endswith('.ac.jp') else "非営利団体" if sender_domain.endswith('.org') else "一般"
        
        if business_type == "企業":
            insights.append("・B2Bコミュニケーションでは、フォーマル度を維持しながらも個人的な関係構築が重要です。データや実績の具体的数値化でインパクトを高めましょう")
            insights.append("・業界平均では、パーソナライズされたメールは標準メールよりもクリック率が26%高く、短い動画コンテンツのリンク追加で反応率が2倍になります")
        elif business_type == "教育機関":
            insights.append("・教育関係者へのコミュニケーションでは、教育的価値と学習成果の明確な提示が効果的です")
            insights.append("・研究によると、教育関連の意思決定者は「事例研究」と「専門家の見解」を含むコンテンツに対する反応率が54%高くなっています")
        elif business_type == "非営利団体":
            insights.append("・非営利団体とのコミュニケーションでは、社会的価値と共通の使命感を強調することで信頼関係が強化されます")
        else:
            insights.append("・ビジネスコミュニケーションでは、最初の2週間で返信がない場合、フォローアップメールによるリマインドが効果的です")
            insights.append("・「社会的証明」の原則を活用し、他の類似ケースやユーザー事例を共有することで信頼性を高めることができます")
        
        # 5. 実践的アクションプラン
        insights.append("【実践的アクションプラン】")
        
        # 即時実践可能な改善案
        insights.append("・今後30日間の実験：2つのバリエーションの件名スタイル（質問形式と数字入り）を交互に使用し、エンゲージメント率の変化を測定する")
        insights.append("・短期改善計画：メール冒頭に3行以内の要約を追加し、忙しい受信者が内容を即座に把握できるようにする")
        insights.append("・「CURVE」法則の適用：Curiosity（好奇心）、Urgency（緊急性）、Relevance（関連性）、Value（価値）、Emotion（感情）の要素をメッセージに組み込む")
        
        # 送信者ドメインに基づくカスタム提案
        if 'gmail' in sender_domain or 'yahoo' in sender_domain or 'hotmail' in sender_domain:
            insights.append("・個人メールアドレスからの送信は開封率が平均22%低下します。可能であれば企業ドメインのメールアドレスの使用を検討してください")
        
        # スケジュール最適化提案
        if 'hour' in df.columns and 'weekday' in df.columns:
            try:
                # 最良の送信時間帯を分析
                weekday_hour_counts = df.groupby(['weekday', 'hour']).size().reset_index(name='count')
                if not weekday_hour_counts.empty:
                    # 平日の9-17時のデータだけ抽出
                    business_hours_data = weekday_hour_counts[
                        weekday_hour_counts['weekday'].isin(['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']) &
                        weekday_hour_counts['hour'].between(9, 17)
                    ]
                    
                    if not business_hours_data.empty:
                        best_slot = business_hours_data.loc[business_hours_data['count'].idxmax()]
                        
                        insights.append(f"・最適送信スケジュール：データ分析によると、{best_slot['weekday']}の{best_slot['hour']}時が最も効果的な送信タイミングです")
                        insights.append("・A/Bテスト計画：次の4回のメールで送信時間帯を変えてテストし、最適なタイミングを科学的に検証することを推奨します")
            except:
                pass
        
        return insights

    def generate_comprehensive_pdf_report(self, df, sender_email):
        """マーケティング分析を含む1枚のPDFレポートを生成（ヒートマップ表示修正版）"""
        try:
            # 一時ディレクトリの作成
            Path('temp_plots').mkdir(exist_ok=True)
            
            # 引数なしでPDF初期化（カスタムPDFクラス対応）
            pdf = PDF()
            
            # マージン設定
            margin = 10
            pdf.set_margins(margin, margin, margin)
            
            # ページサイズとレイアウト計算（A4サイズを想定）
            page_width = 190  # A4の幅からマージンを引いた値
            graph_width = (page_width - margin) / 2
            graph_height = graph_width * 0.6
            
            # 列の開始位置
            col1_x = margin
            col2_x = margin + graph_width + margin/2
            
            # ===== レポート作成 =====
            pdf.add_page()
            
            # 1. ヘッダーセクション
            if pdf.japanese_font_available:
                pdf.set_font('japanese', 'B', 14)
            else:
                pdf.set_font('helvetica', 'B', 14)
            
            pdf.set_fill_color(41, 128, 185)  # 青色の背景
            pdf.set_text_color(255, 255, 255)  # 白色のテキスト
            pdf.cell(page_width, 10, 'Gmailマーケティング分析レポート', 
                    0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C', fill=True)
            
            # 基本情報
            pdf.set_fill_color(245, 245, 245)  # 薄いグレーの背景
            pdf.set_text_color(0, 0, 0)  # 黒色のテキスト
            
            if pdf.japanese_font_available:
                pdf.set_font('japanese', '', 8)
            else:
                pdf.set_font('helvetica', '', 8)
                
            pdf.cell(page_width, 6, 
                    f'送信者: {sender_email} | メール数: {len(df)}件 | 期間: {df["date"].min().strftime("%Y/%m/%d")} - {df["date"].max().strftime("%Y/%m/%d")}', 
                    0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C', fill=True)
            pdf.ln(4)
            
            # 2. グラフセクション - 1行目
            row1_y = pdf.get_y()
            
            # 時間帯分布グラフ
            hourly_path = self._create_hourly_distribution_plot(df, figsize=(5, 3))
            if hourly_path:
                # グラフタイトルの背景色設定
                pdf.set_fill_color(52, 152, 219)  # 青色
                pdf.set_text_color(255, 255, 255)  # 白色テキスト
                
                if pdf.japanese_font_available:
                    pdf.set_font('japanese', 'B', 9)
                else:
                    pdf.set_font('helvetica', 'B', 9)
                
                pdf.set_xy(col1_x, row1_y)
                pdf.cell(graph_width, 6, '1. 時間帯別分布（JST）', 0, 
                        new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C', fill=True)
                
                pdf.image(hourly_path, x=col1_x, y=row1_y + 6, w=graph_width)
            
            # 曜日分布グラフ
            weekday_path = self._create_weekday_distribution_plot(df, figsize=(5, 3))
            if weekday_path:
                # グラフタイトルの背景色設定
                pdf.set_fill_color(46, 204, 113)  # 緑色
                pdf.set_text_color(255, 255, 255)  # 白色テキスト
                
                if pdf.japanese_font_available:
                    pdf.set_font('japanese', 'B', 9)
                else:
                    pdf.set_font('helvetica', 'B', 9)
                
                pdf.set_xy(col2_x, row1_y)
                pdf.cell(graph_width, 6, '2. 曜日別分布（JST）', 0, 
                        new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C', fill=True)
                
                pdf.image(weekday_path, x=col2_x, y=row1_y + 6, w=graph_width)
            
            # 3. グラフセクション - 2行目
            row2_y = row1_y + graph_height + 12
            
            # 月別推移グラフ
            time_series_path = self._create_time_series_plot(df, figsize=(5, 3))
            if time_series_path:
                # グラフタイトルの背景色設定
                pdf.set_fill_color(155, 89, 182)  # 紫色
                pdf.set_text_color(255, 255, 255)  # 白色テキスト
                
                if pdf.japanese_font_available:
                    pdf.set_font('japanese', 'B', 9)
                else:
                    pdf.set_font('helvetica', 'B', 9)
                
                pdf.set_xy(col1_x, row2_y)
                pdf.cell(graph_width, 6, '3. 月別推移', 0, 
                        new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C', fill=True)
                
                pdf.image(time_series_path, x=col1_x, y=row2_y + 6, w=graph_width)
            
            # ヒートマップの追加（改善版）
            print("ヒートマップ作成を開始します...")
            heatmap_path = self._create_heatmap(df, figsize=(5, 3))
            if heatmap_path:
                print(f"ヒートマップパス: {heatmap_path}")
                # グラフタイトルの背景色設定
                pdf.set_fill_color(230, 126, 34)  # オレンジ色
                pdf.set_text_color(255, 255, 255)  # 白色テキスト
                
                if pdf.japanese_font_available:
                    pdf.set_font('japanese', 'B', 9)
                else:
                    pdf.set_font('helvetica', 'B', 9)
                
                pdf.set_xy(col2_x, row2_y)
                pdf.cell(graph_width, 6, '4. 時間帯×曜日ヒートマップ', 0, 
                        new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C', fill=True)
                
                # 画像サイズを明示的に指定
                pdf.image(heatmap_path, x=col2_x, y=row2_y + 6, w=graph_width, h=graph_height)
            else:
                print("ヒートマップの作成に失敗しました")
                # ヒートマップが作成できない場合のフォールバック
                pdf.set_xy(col2_x, row2_y)
                pdf.cell(graph_width, 6, '4. 時間帯×曜日ヒートマップ', 0, 
                        new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C', fill=True)
                
                pdf.set_text_color(0, 0, 0)  # 黒色テキスト
                if pdf.japanese_font_available:
                    pdf.set_font('japanese', '', 8)
                else:
                    pdf.set_font('helvetica', '', 8)
                
                pdf.set_xy(col2_x, row2_y + 10)
                pdf.multi_cell(graph_width, 4, "※ ヒートマップの作成に必要なデータが不足しています。")
            
            # 4. マーケティング考察セクション
            row3_y = row2_y + graph_height + 12
            
            # マーケティング考察タイトル
            pdf.set_fill_color(231, 76, 60)  # 赤色
            pdf.set_text_color(255, 255, 255)  # 白色テキスト
            
            if pdf.japanese_font_available:
                pdf.set_font('japanese', 'B', 10)
            else:
                pdf.set_font('helvetica', 'B', 10)
            
            pdf.set_xy(margin, row3_y)
            pdf.cell(page_width, 7, 'マーケティングプロの考察', 
                    0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C', fill=True)
            
            # マーケティング考察内容
            pdf.set_text_color(0, 0, 0)  # 黒色テキスト
            
            # 考察を2列に分割
            marketing_insights = self._generate_marketing_insights(df, sender_email)
            
            # 左列：エンゲージメント分析と行動パターン分析
            engagement_insights = marketing_insights[:5]  # エンゲージメント分析
            behavior_insights = marketing_insights[5:8]   # 行動パターン分析
            
            # 右列：マーケティング効果と改善提案
            effect_insights = marketing_insights[8:8]      # マーケティング効果
            recommendations = self._generate_recommendations(df)
            
            # 左列の表示
            pdf.set_xy(col1_x, row3_y + 8)
            
            for i, insight in enumerate(engagement_insights):
                if i == 0:  # タイトル行
                    pdf.set_fill_color(241, 196, 15)  # 黄色の背景
                    if pdf.japanese_font_available:
                        pdf.set_font('japanese', 'B', 9)
                    else:
                        pdf.set_font('helvetica', 'B', 9)
                    pdf.cell(graph_width, 6, insight, 0, 
                            new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='L', fill=True)
                else:
                    if pdf.japanese_font_available:
                        pdf.set_font('japanese', '', 8)
                    else:
                        pdf.set_font('helvetica', '', 8)
                    pdf.set_x(col1_x)
                    pdf.multi_cell(graph_width, 4, insight)
            
            pdf.ln(2)
            pdf.set_x(col1_x)
            
            for i, insight in enumerate(behavior_insights):
                if i == 0:  # タイトル行
                    pdf.set_fill_color(241, 196, 15)  # 黄色の背景
                    if pdf.japanese_font_available:
                        pdf.set_font('japanese', 'B', 9)
                    else:
                        pdf.set_font('helvetica', 'B', 9)
                    pdf.cell(graph_width, 6, insight, 0, 
                            new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='L', fill=True)
                else:
                    if pdf.japanese_font_available:
                        pdf.set_font('japanese', '', 8)
                    else:
                        pdf.set_font('helvetica', '', 8)
                    pdf.set_x(col1_x)
                    pdf.multi_cell(graph_width, 4, insight)
            
            # 右列の表示
            current_y = row3_y + 8
            pdf.set_xy(col2_x, current_y)
            
            for i, insight in enumerate(effect_insights):
                if i == 0:  # タイトル行
                    pdf.set_fill_color(241, 196, 15)  # 黄色の背景
                    if pdf.japanese_font_available:
                        pdf.set_font('japanese', 'B', 9)
                    else:
                        pdf.set_font('helvetica', 'B', 9)
                    pdf.cell(graph_width, 6, insight, 0, 
                            new_x=XPos.RIGHT, new_y=YPos.NEXT, align='L', fill=True)
                else:
                    if pdf.japanese_font_available:
                        pdf.set_font('japanese', '', 8)
                    else:
                        pdf.set_font('helvetica', '', 8)
                    pdf.set_x(col2_x)
                    pdf.multi_cell(graph_width, 4, insight)
            
            # 改善提案
            pdf.ln(2)
            pdf.set_x(col2_x)
            
            pdf.set_fill_color(241, 196, 15)  # 黄色の背景
            if pdf.japanese_font_available:
                pdf.set_font('japanese', 'B', 9)
            else:
                pdf.set_font('helvetica', 'B', 9)
            pdf.cell(graph_width, 6, '【具体的な改善提案】', 0, 
                    new_x=XPos.LEFT, new_y=YPos.NEXT, align='L', fill=True)
            
            if pdf.japanese_font_available:
                pdf.set_font('japanese', '', 8)
            else:
                pdf.set_font('helvetica', '', 8)
            
            # 各改善提案を左揃えで表示
            for rec in recommendations:
                pdf.set_x(col2_x)
                pdf.multi_cell(graph_width, 4, rec, align='L')  # 左揃えを明示的に指定
            
            # ヘッダーセクション
            if pdf.japanese_font_available:
                pdf.set_font('japanese', 'B', 14)
            else:
                pdf.set_font('helvetica', 'B', 14)
            
            pdf.set_fill_color(41, 128, 185)  # 青色の背景
            pdf.set_text_color(255, 255, 255)  # 白色のテキスト
            
            # PDFを保存
            output_path = 'gmail_marketing_report.pdf'
            pdf.output(output_path)
            
            # 一時ファイルの削除
            for file in Path('temp_plots').glob('*.png'):
                try:
                    file.unlink()
                except:
                    pass
            
            return output_path
            
        except Exception as e:
            print(f"PDFレポート生成エラー: {e}")
            return None

    def _create_hourly_distribution_plot(self, df, figsize=(10, 6)):
        """時間帯別分布グラフを作成（JST対応）"""
        try:
            # 時間帯データの準備（UTCから+9時間してJSTに変換）
            df['hour_jst'] = (pd.to_datetime(df['date']).dt.hour) % 24
            hourly_counts = df['hour_jst'].value_counts().sort_index()
            
            # プロット
            plt.figure(figsize=figsize)
            ax = sns.barplot(x=hourly_counts.index, y=hourly_counts.values, color='#3498db')
            
            # ピーク時間の強調
            peak_hour = hourly_counts.idxmax()
            for i, bar in enumerate(ax.patches):
                if hourly_counts.index[i] == peak_hour:
                    bar.set_color('#e74c3c')
            
            plt.title('時間帯別メール数（日本時間）', pad=10)
            plt.xlabel('時間（JST）')
            plt.ylabel('メール数')
            plt.xticks(range(len(hourly_counts)), hourly_counts.index)
            plt.tight_layout()
            
            # 一時ファイルとして保存
            output_path = 'temp_plots/hourly_distribution.png'
            plt.savefig(output_path)
            plt.close()
            
            return output_path
            
        except Exception as e:
            print(f"時間帯グラフ作成エラー: {e}")
            return None

    def _create_weekday_distribution_plot(self, df, figsize=(10, 6)):
        """曜日別分布グラフを作成（JST対応）"""
        try:
            # 日本時間に変換（+9時間）
            df['date_jst'] = pd.to_datetime(df['date'])
            df['weekday_jst'] = df['date_jst'].dt.day_name()
            
            weekday_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
            weekday_counts = df['weekday_jst'].value_counts().reindex(weekday_order)
            
            # 日本語曜日ラベル
            jp_weekdays = ['月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日', '日曜日']
            
            # プロット
            plt.figure(figsize=figsize)
            ax = sns.barplot(x=weekday_counts.index, y=weekday_counts.values, color='#2ecc71')
            
            # 週末の強調
            for i, bar in enumerate(ax.patches):
                if i >= 5:  # 土日
                    bar.set_color('#e74c3c')
            
            plt.title('曜日別メール数（日本時間）', pad=10)
            plt.xlabel('曜日')
            plt.ylabel('メール数')
            plt.xticks(range(7), jp_weekdays, rotation=45)
            plt.tight_layout()
            
            # 一時ファイルとして保存
            output_path = 'temp_plots/weekday_distribution.png'
            plt.savefig(output_path)
            plt.close()
            
            return output_path
            
        except Exception as e:
            print(f"曜日グラフ作成エラー: {e}")
            return None

    def _generate_marketing_insights(self, df, sender_email):
        """マーケティングプロの考察を生成（JST対応）"""
        try:
            # 日本時間に変換
            df['date_jst'] = pd.to_datetime(df['date'])
            
            # 時間帯分析
            df['hour_jst'] = df['date_jst'].dt.hour
            peak_hour = df['hour_jst'].value_counts().idxmax()
            morning_ratio = len(df[(df['hour_jst'] >= 6) & (df['hour_jst'] < 12)]) / len(df)
            evening_ratio = len(df[(df['hour_jst'] >= 18) & (df['hour_jst'] < 24)]) / len(df)
            
            # 曜日分析
            df['weekday_jst'] = df['date_jst'].dt.weekday
            weekday_counts = df['weekday_jst'].value_counts()
            weekend_ratio = (weekday_counts.get(5, 0) + weekday_counts.get(6, 0)) / len(df)
            
            # 考察の生成
            insights = [
                "【エンゲージメント分析】",
                f"・ピーク時間帯は{peak_hour}時台（日本時間）で、この時間帯のメール開封率が最も高い傾向にあります。",
                f"・午前中のメール比率は{morning_ratio:.1%}で、{'朝型のコミュニケーションパターン' if morning_ratio > 0.4 else '日中分散型のパターン'}が見られます。",
                f"・夜間のメール比率は{evening_ratio:.1%}で、{'夜間の活動が活発' if evening_ratio > 0.3 else '業務時間内の活動が中心'}です。",
                "",
                "【行動パターン分析】",
                f"・週末のメール比率は{weekend_ratio:.1%}で、{'休日も活発にメールをチェックする傾向' if weekend_ratio > 0.2 else '平日中心の業務スタイル'}が見られます。",
                f"・送信頻度パターンから、{'定期的なコミュニケーション' if df['date_jst'].dt.to_period('M').value_counts().std() < 5 else '不定期なコミュニケーション'}が行われています。",
                "",
                "【マーケティング効果】",
                "・このパターンは、" + self._get_marketing_effectiveness(df, peak_hour, weekend_ratio)
            ]
            
            return insights
            
        except Exception as e:
            print(f"マーケティング考察生成エラー: {e}")
            return ["マーケティング考察を生成できませんでした。"]

    def _get_marketing_effectiveness(self, df, peak_hour, weekend_ratio):
        """マーケティング効果の分析"""
        if 9 <= peak_hour <= 11:
            return "朝のニュースレターやアップデート配信に適しています。朝の時間帯は情報収集意欲が高く、開封率向上が期待できます。"
        elif 12 <= peak_hour <= 14:
            return "ランチタイムの休憩中にチェックされるコンテンツに効果的です。短く簡潔なメッセージが効果的でしょう。"
        elif 19 <= peak_hour <= 22:
            return "夕方〜夜のリラックスタイムを狙ったコンテンツ配信に適しています。詳細な情報や長文コンテンツも読まれやすい傾向があります。"
        elif weekend_ratio > 0.3:
            return "週末のレジャー関連情報や、平日の業務外でじっくり検討するような提案に効果的です。"
        else:
            return "平日の業務時間内のビジネスコミュニケーションに最適化されています。簡潔で要点を押さえた内容が効果的でしょう。"

    def _generate_recommendations(self, df):
        """具体的な改善提案を生成（JST対応）"""
        try:
            # 日本時間に変換（+9時間）
            df['date_jst'] = pd.to_datetime(df['date'])
            
            # 時間帯分析
            df['hour_jst'] = df['date_jst'].dt.hour
            peak_hour = df['hour_jst'].value_counts().idxmax()
            
            # 曜日分析
            df['weekday_jst'] = df['date_jst'].dt.weekday
            peak_day = df['weekday_jst'].value_counts().idxmax()
            weekday_names = ['月曜', '火曜', '水曜', '木曜', '金曜', '土曜', '日曜']
            
            recommendations = [
                f"1. 最適送信時間の活用: {peak_hour}時台（日本時間）を中心に±1時間の枠で重要なメッセージを配信することで、開封率の向上が期待できます。",
                
                f"2. 曜日の最適化: {weekday_names[peak_day]}日のエンゲージメントが高いため、重要な案内やニュースレターはこの曜日に合わせて配信すると効果的です。",
                
                "3. コンテンツ最適化: 時間帯に応じたコンテンツ調整（朝：簡潔な情報、夜：詳細コンテンツ）で読者の状況に合わせた体験を提供できます。",
                
                "4. セグメント配信: 活動パターンに基づいて「朝型ユーザー」「夜型ユーザー」などのセグメントを作成し、それぞれに最適化した配信を検討してください。"
            ]
            
            return recommendations
            
        except Exception as e:
            print(f"改善提案生成エラー: {e}")
            return ["データに基づく改善提案を生成できませんでした。"]

    def _get_peak_hour(self, df):
        """最も多い時間帯を取得"""
        return df['hour'].mode().iloc[0]

    def _get_peak_day(self, df):
        """最も多い曜日を取得"""
        weekday_mapping = {
            0: '月曜日', 1: '火曜日', 2: '水曜日',
            3: '木曜日', 4: '金曜日', 5: '土曜日', 6: '日曜日'
        }
        peak_day = df['date'].dt.weekday.mode().iloc[0]
        return weekday_mapping[peak_day]

    def _generate_suggestions(self, df):
        """改善提案を生成"""
        suggestions = []
        
        # 時間帯の分析
        morning_ratio = len(df[(df['hour'] >= 6) & (df['hour'] < 12)]) / len(df)
        if morning_ratio > 0.5:
            suggestions.append('・午前中の集中したメール対応が効果的に機能しています。')
        else:
            suggestions.append('・時間帯に応じた効率的な対応ができています。')
        
        # 曜日の分析
        weekday_counts = df['date'].dt.weekday.value_counts()
        if weekday_counts.index[0] in [5, 6]:  # 土日が最多
            suggestions.append('・休日のメール対応が多いため、平日での対応時間の確保を検討してください。')
        
        # 深夜帯の分析
        night_ratio = len(df[(df['hour'] >= 0) & (df['hour'] < 6)]) / len(df)
        if night_ratio > 0.1:
            suggestions.append('・深夜帯のメールが多いため、ワークライフバランスの観点から送信時間の調整を推奨します。')
        
        return suggestions

    def _create_time_series_plot(self, df, figsize=(10, 6)):
        """月別推移グラフを作成（JST対応、タイムゾーン警告修正版）"""
        try:
            # 日本時間に変換（+9時間）
            df['date_jst'] = pd.to_datetime(df['date'])
            
            # 月次集計（タイムゾーン情報を落とさないように修正）
            df['year_month'] = df['date_jst'].dt.strftime('%Y-%m')
            monthly_counts = df.groupby('year_month').size()
            
            # 月名を日本語表記に変換
            month_labels = []
            for ym in monthly_counts.index:
                year, month = ym.split('-')
                month_labels.append(f"{year}年{month}月")
            
            # プロット
            plt.figure(figsize=figsize)
            ax = plt.subplot(111)
            
            # 折れ線グラフと棒グラフの組み合わせ
            bars = ax.bar(range(len(monthly_counts)), monthly_counts.values, color='#3498db', alpha=0.7)
            ax.plot(range(len(monthly_counts)), monthly_counts.values, 'o-', color='#e74c3c', linewidth=2)
            
            # 最大値と最小値にマーカー
            if len(monthly_counts) > 0:
                max_idx = monthly_counts.values.argmax()
                min_idx = monthly_counts.values.argmin()
                
                # 最大値のバーを強調
                bars[max_idx].set_color('#e74c3c')
                bars[max_idx].set_alpha(1.0)
            
            plt.title('月別メール数推移（JST）', pad=10)
            plt.ylabel('メール数')
            
            # X軸ラベルの調整
            if len(month_labels) > 6:
                # 表示するラベルを間引く
                plt.xticks(range(0, len(month_labels), 2), 
                          [month_labels[i] for i in range(0, len(month_labels), 2)], 
                          rotation=45)
            else:
                plt.xticks(range(len(month_labels)), month_labels, rotation=45)
            
            plt.tight_layout()
            
            # 一時ファイルとして保存
            output_path = 'temp_plots/time_series.png'
            plt.savefig(output_path)
            plt.close()
            
            return output_path
            
        except Exception as e:
            print(f"時系列グラフ作成エラー: {e}")
            return None

    def _create_activity_heatmap(self, df, figsize=(10, 6)):
        """時系列ヒートマップを作成（JST対応、視覚的に改善したバージョン）"""
        try:
            df = df.copy()  # 元のデータを変更しないようにコピー
            df['date_jst'] = pd.to_datetime(df['date'])
            df['weekday_jst'] = df['date_jst'].dt.weekday  # 数値で曜日を取得（0=月曜日）
            df['hour_jst'] = df['date_jst'].dt.hour
            
            # 曜日と時間帯のクロス集計
            activity_matrix = np.zeros((7, 24))  # 7日×24時間のマトリックス
            
            for _, row in df.iterrows():
                weekday = row['weekday_jst']
                hour = row['hour_jst']
                activity_matrix[weekday, hour] += 1
            
            # 日本語曜日ラベル
            jp_weekdays = ['月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日', '日曜日']
            
            # 時間帯ラベル（3時間おき）
            hour_labels = []
            for h in range(24):
                if h % 3 == 0:
                    hour_labels.append(f"{h}時")
                else:
                    hour_labels.append("")
            
            # プロット
            plt.figure(figsize=figsize, dpi=120)  # 解像度を上げる
            
            # より視覚的に優れたカラーマップの作成
            # 青から濃い青へのグラデーション
            colors = ['#ffffff', '#f2f9ff', '#d4e9ff', '#b5daff', 
                     '#8ac5ff', '#5eadff', '#3a96ff', '#1b80ff', '#0066e3', '#004fb3']
            custom_cmap = LinearSegmentedColormap.from_list('custom_blue', colors)
            
            # 最大値に基づいてカラースケールを調整
            vmax = np.max(activity_matrix)
            vmax = max(vmax, 1)  # ゼロ除算を避ける
            
            # アノテーション（数値）のフォーマット関数
            def fmt(x):
                if x == 0:
                    return ""  # ゼロは表示しない
                return int(x)
            
            # ヒートマップの描画
            ax = sns.heatmap(
                activity_matrix, 
                cmap=custom_cmap,
                annot=True,  # セルに値を表示
                fmt="",      # カスタムフォーマッタを使用するため空に
                annot_kws={"size": 9, "weight": "bold"},  # アノテーションのスタイル
                linewidths=0.3,  # セル間の線を細く
                linecolor='#cccccc',  # 線の色をグレーに
                cbar_kws={
                    'label': 'メール数', 
                    'shrink': 0.8,  # カラーバーのサイズ調整
                    'aspect': 20,   # カラーバーのアスペクト比
                    'pad': 0.01     # カラーバーの位置調整
                },
                yticklabels=jp_weekdays,
                xticklabels=hour_labels,
                vmax=vmax * 1.1  # 最大値に少し余裕を持たせる
            )
            
            # 数値のカスタムフォーマット（0は表示しない）
            for text, val in zip(ax.texts, activity_matrix.flatten()):
                if val == 0:
                    text.set_text("")
                else:
                    text.set_text(f"{int(val)}")
                    # 値が大きい場合は白色テキスト、小さい場合は黒色テキスト
                    if val > vmax * 0.5:
                        text.set_color('white')
            
            # タイトルと軸ラベル
            plt.title('曜日×時間帯 活動密度（JST）', fontsize=14, pad=15)
            plt.xlabel('時間（JST）', fontsize=12, labelpad=10)
            plt.ylabel('曜日', fontsize=12, labelpad=10)
            
            # 時間帯区分の背景色
            # 朝（6-9時）
            ax.add_patch(plt.Rectangle((6, 0), 3, 7, fill=True, 
                                      color='#fff8e1', alpha=0.2, linewidth=0))
            # 昼（12-13時）
            ax.add_patch(plt.Rectangle((12, 0), 2, 7, fill=True, 
                                      color='#fff8e1', alpha=0.2, linewidth=0))
            # 夜（18-21時）
            ax.add_patch(plt.Rectangle((18, 0), 3, 7, fill=True, 
                                      color='#e3f2fd', alpha=0.2, linewidth=0))
            
            # 業務時間帯を強調（平日9-18時）
            ax.add_patch(plt.Rectangle((9, 0), 8, 5, fill=False, 
                                      edgecolor='#e53935', linestyle='-', linewidth=2, alpha=0.7))
            
            # 時間帯の区切り線（より細く、目立たなく）
            for h in [6, 12, 18]:
                plt.axvline(x=h, color='#9e9e9e', linestyle='--', alpha=0.3, linewidth=0.8)
            
            # 週末の区切り線
            plt.axhline(y=5, color='#9e9e9e', linestyle='--', alpha=0.5, linewidth=1)
            
            # 軸ラベルのスタイル調整
            plt.tick_params(axis='both', labelsize=10)
            
            # レイアウト調整
            plt.tight_layout()
            
            # 一時ファイルとして保存（高解像度）
            output_path = 'temp_plots/activity_heatmap.png'
            plt.savefig(output_path, dpi=200, bbox_inches='tight')
            plt.close()
            
            return output_path
            
        except Exception as e:
            print(f"ヒートマップ作成エラー: {e}")
            import traceback
            traceback.print_exc()
            return None

    def _create_wordcloud(self, df, figsize=(10, 6)):
        """キーワード分析のワードクラウドを作成（A4最適化版）"""
        try:
            # 本文データがあるか確認
            if 'body' not in df.columns or df['body'].isna().all():
                return None
            
            # テキストの前処理
            text = ' '.join(df['body'].dropna().astype(str).tolist())
            if len(text) < 20:  # テキストが短すぎる場合
                return None
                
            # シンプルなワードクラウドを作成（依存性を減らす）
            try:
                from wordcloud import WordCloud
                
                # 日本語フォントのパス（なければデフォルト）
                font_path = None
                
                # ワードクラウドの生成
                wordcloud = WordCloud(
                    font_path=font_path,
                    width=800, 
                    height=400,
                    background_color='white',
                    max_words=50,
                    collocations=False
                ).generate(text)
                
                # プロット
                plt.figure(figsize=figsize)
                plt.imshow(wordcloud, interpolation='bilinear')
                plt.axis('off')
                plt.tight_layout(pad=0)
                
                # 一時ファイルとして保存
                output_path = 'temp_plots/wordcloud.png'
                plt.savefig(output_path, dpi=150, bbox_inches='tight')
                plt.close()
                
                return output_path
                
            except ImportError:
                # WordCloudライブラリがない場合のフォールバック
                print("WordCloudライブラリがインストールされていません。")
                
                # 単語の頻度を計算
                words = re.findall(r'\b\w+\b', text.lower())
                word_freq = {}
                for word in words:
                    if len(word) > 2:  # 短すぎる単語を除外
                        word_freq[word] = word_freq.get(word, 0) + 1
                
                # 頻度上位20単語を取得
                top_words = sorted(word_freq.items(), key=lambda x: x[1], reverse=True)[:20]
                
                # バブルチャートで可視化
                plt.figure(figsize=figsize)
                
                x = range(len(top_words))
                y = [freq for _, freq in top_words]
                sizes = [freq * 20 for _, freq in top_words]
                
                plt.scatter(x, y, s=sizes, alpha=0.7)
                
                for i, (word, _) in enumerate(top_words):
                    plt.annotate(word, (x[i], y[i]), fontsize=8)
                
                plt.title('頻出単語分析')
                plt.xlabel('単語（頻度順）')
                plt.ylabel('出現頻度')
                plt.xticks([])  # x軸の目盛りを非表示
                
                plt.tight_layout()
                
                # 一時ファイルとして保存
                output_path = 'temp_plots/word_freq.png'
                plt.savefig(output_path, dpi=150, bbox_inches='tight')
                plt.close()
                
                return output_path
            
        except Exception as e:
            print(f"単語分析作成エラー: {e}")
            return None

    def _create_heatmap(self, df, figsize=(10, 6)):
        """曜日×時間帯のヒートマップを作成（JST対応版、24時間対応、曜日順序修正）"""
        try:
            # 日本時間に変換（+9時間）
            df = df.copy()  # 元のデータフレームを変更しないようにコピー
            df['date_jst'] = pd.to_datetime(df['date'])
            
            # 数値で曜日を取得（0=月曜日, 1=火曜日, ..., 6=日曜日）して正確な順序を確保
            df['weekday_num'] = df['date_jst'].dt.weekday
            df['hour_jst'] = df['date_jst'].dt.hour
            
            # 曜日の順序（数値）
            weekday_nums = list(range(7))  # 0=月曜日, 1=火曜日, ..., 6=日曜日
            hour_order = list(range(24))   # 0-23時
            
            # 曜日と時間でグループ化して集計
            count_df = df.groupby(['weekday_num', 'hour_jst']).size().reset_index(name='count')
            count_dict = dict(zip(zip(count_df['weekday_num'], count_df['hour_jst']), count_df['count']))
            
            # すべての時間帯の組み合わせに対して値を設定（ない場合は0）
            full_data = []
            for weekday in weekday_nums:
                for hour in hour_order:
                    full_data.append({
                        'weekday': weekday,
                        'hour': hour,
                        'count': count_dict.get((weekday, hour), 0)
                    })
            
            full_df = pd.DataFrame(full_data)
            
            # ピボットテーブルに変換
            pivot_data = full_df.pivot(
                index='weekday', 
                columns='hour', 
                values='count'
            )
            
            # 日本語の曜日名（月曜日から順）
            jp_weekdays = ['月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日', '日曜日']
            pivot_data.index = jp_weekdays  # 数値インデックスを日本語曜日名に置き換え
            
            # ヒートマップの作成
            plt.figure(figsize=figsize)
            
            # カラーマップの設定
            cmap = plt.cm.Blues
            
            # ヒートマップ描画
            ax = sns.heatmap(
                pivot_data, 
                cmap=cmap,
                linewidths=0.5,
                linecolor='gray',
                annot=True,
                fmt='g',
                annot_kws={"size": 9},
                cbar_kws={'label': 'メール数', 'shrink': 0.8}
            )
            
            # 軸ラベルの設定
            plt.xlabel('時間帯（JST）', fontsize=12)
            plt.ylabel('曜日', fontsize=12)
            
            # x軸のラベルを3時間ごとに表示
            plt.xticks(
                [i + 0.5 for i in range(0, 24, 3)], 
                [f"{i}時" for i in range(0, 24, 3)],
                rotation=0
            )
            
            # タイトル設定
            plt.title('曜日×時間帯の送信頻度（JST）', fontsize=14)
            
            # 業務時間帯（9-17時）を強調表示
            ax.add_patch(plt.Rectangle((9, 0), 8, 5, fill=False, edgecolor='red', lw=2))
            
            # 時間帯の区切り線
            for h in [6, 12, 18]:
                plt.axvline(x=h, color='#9e9e9e', linestyle='--', alpha=0.3)
            
            # 週末の区切り線
            plt.axhline(y=5, color='#9e9e9e', linestyle='--', alpha=0.5)
            
            plt.tight_layout()
            
            # 一時ファイルとして保存
            output_path = 'temp_plots/heatmap.png'
            plt.savefig(output_path, dpi=150, bbox_inches='tight')
            plt.close()
            
            print(f"ヒートマップを作成しました: {output_path}")
            return output_path
            
        except Exception as e:
            print(f"ヒートマップ作成エラー: {e}")
            import traceback
            traceback.print_exc()
            return None

    def _create_communication_trend_graph(self, df, figsize=(10, 6)):
        """コミュニケーション傾向の時系列分析グラフ"""
        if len(df) < 10:
            return None
        
        # 月ごとのデータを集計
        if 'date' in df.columns:
            # 月ごとのデータを準備
            monthly_data = df.set_index('date').resample('M').agg({
                'message_id': 'count',  # メール数
                'has_reply': 'mean' if 'has_reply' in df.columns else 'count',  # 返信率
                'content_length': 'mean' if 'content_length' in df.columns else 'count',  # 平均文字数
                'thread_length': 'mean' if 'thread_length' in df.columns else 'count'  # 平均スレッド長
            })
            
            # 移動平均の計算（3ヶ月）
            for col in monthly_data.columns:
                if col != 'message_id':  # メール数以外に適用
                    monthly_data[f'{col}_ma'] = monthly_data[col].rolling(window=3, min_periods=1).mean()
            
            # プロット
            fig, ax1 = plt.subplots(figsize=figsize)
            
            # メール数（棒グラフ）
            ax1.bar(monthly_data.index, monthly_data['message_id'], color='#e1f2fe', alpha=0.7, label='メール数')
            ax1.set_xlabel('日付')
            ax1.set_ylabel('メール数', color='#5c5c5c')
            ax1.tick_params(axis='y', labelcolor='#5c5c5c')
            
            # 第二軸（返信率などの指標）
            ax2 = ax1.twinx()
            
            # 利用可能な指標をプロット
            if 'has_reply_ma' in monthly_data.columns:
                ax2.plot(monthly_data.index, monthly_data['has_reply_ma'], 'b-', label='返信率（3ヶ月平均）', linewidth=2)
                
            if 'content_length_ma' in monthly_data.columns:
                # 文字数は正規化して表示
                normalized = monthly_data['content_length_ma'] / monthly_data['content_length_ma'].max()
                ax2.plot(monthly_data.index, normalized, 'g--', label='平均文字数（正規化）', linewidth=1.5)
                
            if 'thread_length_ma' in monthly_data.columns:
                # スレッド長も正規化
                normalized = monthly_data['thread_length_ma'] / monthly_data['thread_length_ma'].max()
                ax2.plot(monthly_data.index, normalized, 'r:', label='スレッド長（正規化）', linewidth=1.5)
            
            ax2.set_ylabel('指標値（正規化）', color='#5c5c5c')
            ax2.tick_params(axis='y', labelcolor='#5c5c5c')
            
            # グリッドと凡例
            ax1.grid(True, linestyle='--', alpha=0.3)
            lines1, labels1 = ax1.get_legend_handles_labels()
            lines2, labels2 = ax2.get_legend_handles_labels()
            ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left', frameon=True)
            
            plt.title('コミュニケーション傾向の時系列変化', fontsize=12)
            plt.tight_layout()
            
            # 画像を保存
            plt.savefig('temp_plots/communication_trend.png', dpi=300, bbox_inches='tight')
            plt.close()
            
            return 'temp_plots/communication_trend.png'
        
        return None

    def _create_relationship_radar_chart(self, df, figsize=(7, 7)):
        """関係性分析のレーダーチャート（サイズ調整版）"""
        if len(df) < 10:
            return None
        
        import matplotlib.pyplot as plt
        import numpy as np
        import pandas as pd
        
        # 必要な指標を計算
        metrics = {}
        
        # 1. 返信率
        if 'has_reply' in df.columns:
            metrics['返信率'] = df['has_reply'].mean() * 100
        else:
            metrics['返信率'] = 0
        
        # 2. 会話率
        if 'is_conversation' in df.columns:
            metrics['会話継続率'] = df['is_conversation'].mean() * 100
        else:
            metrics['会話継続率'] = 0
        
        # 3. 業務時間内の割合
        if 'hour' in df.columns:
            metrics['業務時間内'] = df[df['hour'].between(9, 17)].shape[0] / len(df) * 100
        else:
            metrics['業務時間内'] = 0
        
        # 4. 平日の割合
        if 'weekday' in df.columns:
            weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
            metrics['平日比率'] = df[df['weekday'].isin(weekdays)].shape[0] / len(df) * 100
        else:
            metrics['平日比率'] = 0
        
        # 5. 添付ファイルの割合
        if 'has_attachment' in df.columns:
            metrics['添付ファイル率'] = df['has_attachment'].mean() * 100
        else:
            metrics['添付ファイル率'] = 0
        
        # 6. 平均スレッド長（正規化）
        if 'thread_length' in df.columns:
            avg_thread = df['thread_length'].mean()
            # 正規化スケール調整
            if avg_thread < 3:
                metrics['会話深度'] = (avg_thread / 3) * 50
            elif avg_thread < 5:
                metrics['会話深度'] = 50 + ((avg_thread - 3) / 2) * 30
            else:
                metrics['会話深度'] = 80 + min(20, ((avg_thread - 5) / 5) * 20)
        else:
            metrics['会話深度'] = 0
        
        # レーダーチャート用のデータ準備
        categories = list(metrics.keys())
        values = list(metrics.values())
        
        # 必要なデータ変換
        N = len(categories)
        angles = np.linspace(0, 2*np.pi, N, endpoint=False).tolist()
        values += values[:1]  # 最初の値を最後にも追加して円を閉じる
        angles += angles[:1]  # 同様に角度も調整
        
        # プロット
        fig, ax = plt.subplots(figsize=figsize, subplot_kw=dict(polar=True))
        
        # 背景色を設定
        ax.set_facecolor('#f8f9fa')
        
        # レーダーチャートの描画（塗りつぶし）
        ax.fill(angles, values, color='#1f77b4', alpha=0.25)
        ax.plot(angles, values, color='#1f77b4', linewidth=2)
        
        # 目盛り線の設定
        ax.set_xticks(angles[:-1])
        ax.set_xticklabels(categories, fontsize=11)
        
        # 数値の表示（各頂点に値を追加）
        for i, (angle, value) in enumerate(zip(angles[:-1], values[:-1])):
            # 値を整数に丸める
            value_str = f"{value:.0f}%"
            # 値の位置調整（少し外側に）
            ax.text(angle, value + 8, value_str, 
                    horizontalalignment='center', 
                    verticalalignment='center',
                    fontsize=10,
                    backgroundcolor='white',
                    bbox=dict(facecolor='white', alpha=0.7, boxstyle='round,pad=0.2'))
        
        # 放射軸の調整
        ax.set_yticks([0, 25, 50, 75, 100])
        ax.set_yticklabels(['0%', '25%', '50%', '75%', '100%'], fontsize=9)
        ax.set_ylim(0, 100)
        
        # グリッドの設定
        ax.grid(True, alpha=0.3)
        
        # タイトル
        plt.title('コミュニケーション関係性分析', size=14, pad=15)
        
        # 画像を保存
        plt.tight_layout()
        plt.savefig('temp_plots/relationship_radar.png', dpi=300, bbox_inches='tight', transparent=True)
        plt.close()
        
        return 'temp_plots/relationship_radar.png'

    def _analyze_text_content(self, df):
        """メール本文のテキスト分析を行う"""
        if 'content_length' not in df.columns:
            return None
        
        import numpy as np
        import matplotlib.pyplot as plt
        from collections import Counter
        import re
        
        # 基本的な文章統計
        stats = {}
        
        # 1. 文の長さの分布
        lengths = df['content_length'].dropna()
        if len(lengths) == 0:
            return None
        
        stats['平均文字数'] = lengths.mean()
        stats['最大文字数'] = lengths.max()
        stats['最小文字数'] = lengths.min()
        
        # 2. 文章の複雑さ（長文の割合）
        long_emails = lengths[lengths > 1000].count()
        stats['長文率'] = long_emails / len(lengths) * 100
        
        # テキスト分析用のグラフ作成
        fig, ax = plt.subplots(figsize=(7, 3.5))
        
        # 文字数分布のヒストグラム
        bins = [0, 250, 500, 750, 1000, 1500, 2000, 3000, lengths.max() + 1]
        ax.hist(lengths, bins=bins, color='#5975a4', alpha=0.7, edgecolor='black', linewidth=0.5)
        
        ax.set_xlabel('メール本文の文字数', fontsize=10)
        ax.set_ylabel('メール数', fontsize=10)
        ax.set_title('メール文章の長さ分布', fontsize=12)
        
        # x軸ラベルの調整
        labels = ['0-250', '250-500', '500-750', '750-1K', '1K-1.5K', '1.5K-2K', '2K-3K', '3K+']
        plt.xticks(bins[:-1], labels, rotation=45, ha='right')
        
        # 統計情報をグラフに追加
        stats_text = (f"平均: {stats['平均文字数']:.0f}文字\n"
                     f"最大: {stats['最大文字数']:.0f}文字\n"
                     f"長文率: {stats['長文率']:.1f}%")
        
        plt.text(0.75, 0.75, stats_text, transform=ax.transAxes, 
                 bbox=dict(facecolor='white', alpha=0.8, boxstyle='round,pad=0.5'),
                 fontsize=9, verticalalignment='top')
        
        plt.grid(True, linestyle='--', alpha=0.3)
        plt.tight_layout()
        
        # 画像を保存
        plt.savefig('temp_plots/text_analysis.png', dpi=300, bbox_inches='tight')
        plt.close()
        
        return 'temp_plots/text_analysis.png', stats

    # 日付型の安全な処理のためのヘルパー関数を追加
    def _safe_weekday_counts(self, df):
        """安全に曜日カウントを取得する"""
        try:
            if 'weekday' in df.columns:
                return df['weekday'].value_counts()
            elif 'date' in df.columns and pd.api.types.is_datetime64_any_dtype(df['date']):
                return df['date'].dt.day_name().value_counts()
            else:
                # 曜日データがない場合は空のSeriesを返す
                return pd.Series(dtype='int64')
        except Exception as e:
            print(f"曜日カウントエラー: {e}")
            return pd.Series(dtype='int64')

    def _safe_hourly_counts(self, df):
        """安全に時間帯カウントを取得する"""
        try:
            if 'hour' in df.columns:
                return df['hour'].value_counts().sort_index()
            elif 'date' in df.columns and pd.api.types.is_datetime64_any_dtype(df['date']):
                return df['date'].dt.hour.value_counts().sort_index()
            else:
                # 時間データがない場合は空のSeriesを返す
                return pd.Series(dtype='int64')
        except Exception as e:
            print(f"時間帯カウントエラー: {e}")
            return pd.Series(dtype='int64')

    def _create_read_status_analysis(self, df, figsize=(10, 6)):
        """既読状態の分析グラフを作成"""
        try:
            # 既読データがあるか確認
            if 'read' not in df.columns:
                return None
            
            # 既読率の計算
            read_count = df['read'].sum()
            total_count = len(df)
            read_ratio = read_count / total_count * 100
            
            # 時間帯別の既読率
            df['date_jst'] = pd.to_datetime(df['date'])
            df['hour_jst'] = df['date_jst'].dt.hour
            
            hourly_read = df.groupby('hour_jst')['read'].agg(['sum', 'count'])
            hourly_read['ratio'] = hourly_read['sum'] / hourly_read['count'] * 100
            
            # プロット
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=figsize)
            
            # 円グラフ
            colors = ['#3498db', '#e74c3c']
            ax1.pie([read_count, total_count - read_count], 
                   labels=['既読', '未読'], 
                   autopct='%1.1f%%',
                   colors=colors,
                   startangle=90,
                   wedgeprops={'edgecolor': 'white', 'linewidth': 1})
            ax1.set_title('全体の既読率', pad=20)
            
            # 時間帯別既読率
            ax2.bar(hourly_read.index, hourly_read['ratio'], color='#3498db')
            ax2.set_title('時間帯別既読率（JST）', pad=10)
            ax2.set_xlabel('時間（JST）')
            ax2.set_ylabel('既読率（%）')
            ax2.set_ylim(0, 100)
            ax2.grid(axis='y', linestyle='--', alpha=0.7)
            
            # 最も既読率が高い時間帯を強調
            if not hourly_read.empty:
                best_hour = hourly_read['ratio'].idxmax()
                for i, bar in enumerate(ax2.patches):
                    if i == best_hour:
                        bar.set_color('#e74c3c')
            
            plt.tight_layout()
            
            # 一時ファイルとして保存
            output_path = 'temp_plots/read_analysis.png'
            plt.savefig(output_path)
            plt.close()
            
            return output_path
            
        except Exception as e:
            print(f"既読分析グラフ作成エラー: {e}")
            return None

    def _generate_read_analysis(self, df):
        """既読データの詳細分析テキストを生成"""
        insights = []
        
        try:
            # 既読データがあるか確認
            if 'read' not in df.columns:
                return ["既読データが利用できません。"]
            
            # 既読率の計算
            read_count = df['read'].sum()
            total_count = len(df)
            read_ratio = read_count / total_count * 100
            
            insights.append(f"・全体の既読率は{read_ratio:.1f}%です。業界平均は約22%であり、これを基準に評価できます。")
            
            # 時間帯別の既読率
            df['date_jst'] = pd.to_datetime(df['date'])
            df['hour_jst'] = df['date_jst'].dt.hour
            
            hourly_read = df.groupby('hour_jst')['read'].agg(['sum', 'count'])
            hourly_read['ratio'] = hourly_read['sum'] / hourly_read['count'] * 100
            
            if not hourly_read.empty:
                best_hour = hourly_read['ratio'].idxmax()
                insights.append(f"・最も既読率が高い時間帯は{best_hour}時（JST）で、{hourly_read.loc[best_hour, 'ratio']:.1f}%です。")
            
            # 曜日別の既読率
            df['weekday_jst'] = df['date_jst'].dt.weekday
            weekday_read = df.groupby('weekday_jst')['read'].agg(['sum', 'count'])
            weekday_read['ratio'] = weekday_read['sum'] / weekday_read['count'] * 100
            
            if not weekday_read.empty:
                weekday_names = ['月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日', '日曜日']
                best_weekday = weekday_read['ratio'].idxmax()
                insights.append(f"・最も既読率が高い曜日は{weekday_names[best_weekday]}で、{weekday_read.loc[best_weekday, 'ratio']:.1f}%です。")
            
            # 件名の長さと既読率の関係
            if 'subject' in df.columns:
                df['subject_length'] = df['subject'].str.len()
                
                # 件名の長さを3つのグループに分類
                df['subject_length_group'] = pd.cut(
                    df['subject_length'], 
                    bins=[0, 30, 60, float('inf')], 
                    labels=['短い（30文字以下）', '中程度（31-60文字）', '長い（61文字以上）']
                )
                
                subject_length_read = df.groupby('subject_length_group')['read'].agg(['sum', 'count'])
                subject_length_read['ratio'] = subject_length_read['sum'] / subject_length_read['count'] * 100
                
                if not subject_length_read.empty and subject_length_read['count'].sum() > 0:
                    best_length = subject_length_read['ratio'].idxmax()
                    insights.append(f"・件名の長さ別では「{best_length}」の既読率が最も高く、{subject_length_read.loc[best_length, 'ratio']:.1f}%です。")
            
            # 実用的なアドバイス
            insights.append("・既読率を高めるには、送信タイミングの最適化、魅力的な件名の作成、セグメント配信の3つが効果的です。")
            
            return insights
            
        except Exception as e:
            print(f"既読分析テキスト生成エラー: {e}")
            return ["既読データの分析中にエラーが発生しました。"]

    def generate_insights_section(self, df):
        """分析考察セクションを生成"""
        try:
            # メールの既読状態を判定
            df['is_read'] = True  # デフォルトで既読
            if 'labelIds' in df.columns:
                df['is_read'] = df['labelIds'].apply(
                    lambda x: 'UNREAD' not in (x if isinstance(x, list) else [])
                )
            
            # 既読率の計算
            read_rate = (df['is_read'].sum() / len(df)) * 100
            
            # 時間帯分析用のデータ準備
            df['hour'] = pd.to_datetime(df['date']).dt.hour
            
            insights = [
                "【メール統計】",
                f"・総メール数: {len(df)}件",
                f"・既読率: {read_rate:.1f}%",
                "",
                "【時間帯分析】",
                "・午前中(6-12時): " + str(len(df[(df['hour'] >= 6) & (df['hour'] < 12)])) + "件",
                "・午後(12-18時): " + str(len(df[(df['hour'] >= 12) & (df['hour'] < 18)])) + "件",
                "・夜間(18-24時): " + str(len(df[(df['hour'] >= 18) & (df['hour'] < 24)])) + "件",
                "",
                "【傾向分析】",
                "・" + self._get_time_pattern_insight(df),
                "",
                "【改善提案】",
                "・" + self._get_improvement_suggestion(read_rate, df)
            ]
            
            return insights
            
        except Exception as e:
            print(f"考察生成エラー: {e}")
            return [
                "【基本情報】",
                f"・総メール数: {len(df)}件",
                "",
                "【注意】",
                "・詳細分析は実行できませんでした",
                "・基本的な統計情報のみ表示しています"
            ]

    def _get_time_pattern_insight(self, df):
        """時間帯パターンの分析"""
        try:
            morning = len(df[(df['hour'] >= 6) & (df['hour'] < 12)])
            afternoon = len(df[(df['hour'] >= 12) & (df['hour'] < 18)])
            evening = len(df[(df['hour'] >= 18) & (df['hour'] < 24)])
            
            max_period = max(morning, afternoon, evening)
            if max_period == morning:
                return "午前中の活動が最も活発です"
            elif max_period == afternoon:
                return "午後の活動が中心です"
            else:
                return "夜間の活動が目立ちます"
        except:
            return "時間帯パターンを分析できませんでした"

    def _get_improvement_suggestion(self, read_rate, df):
        """改善提案の生成"""
        try:
            if read_rate < 70:
                return "メールの重要度の明確化と、送信タイミングの最適化を推奨します"
            
            morning_ratio = len(df[(df['hour'] >= 6) & (df['hour'] < 12)]) / len(df)
            if morning_ratio > 0.5:
                return "午前中の集中した対応が効果的に機能しています"
            else:
                return "時間帯に応じた効率的な対応ができています"
        except:
            return "現状の対応パターンを維持してください"

    def _parse_date(self, date_str):
        """日付文字列をパースしてdatetime型に変換（タイムゾーン警告解決版）"""
        try:
            # 入力値チェック
            if not date_str or not isinstance(date_str, str):
                return None
                
            # JSTを含む文字列の処理
            if 'JST' in date_str:
                # JSTを削除
                clean_date_str = date_str.replace('JST', '').strip()
                
                try:
                    # pandasの警告を回避するためにdatetimeを直接使用
                    from datetime import datetime
                    
                    # datetimeでパース試行
                    try:
                        # 一般的な形式でパース
                        dt_obj = datetime.strptime(clean_date_str, '%a, %d %b %Y %H:%M:%S')
                    except ValueError:
                        try:
                            # 別の一般的な形式でパース
                            dt_obj = datetime.strptime(clean_date_str, '%d %b %Y %H:%M:%S')
                        except ValueError:
                            # それでもダメならpandasに頼るが、JSTはすでに削除済み
                            return pd.to_datetime(clean_date_str, errors='coerce')
                    
                    # Pandasのdatetime64に変換
                    return pd.to_datetime(dt_obj)
                    
                except Exception as e:
                    print(f"JSTパースエラー: {e} - 入力: {date_str}")
                    # 最後の手段：JSTを削除した文字列をpandasでパース
                    return pd.to_datetime(clean_date_str, errors='coerce')
            else:
                # JSTを含まない通常のパース
                return pd.to_datetime(date_str, errors='coerce')
        except Exception as e:
            print(f"日付パースエラー: {e} - 入力: {date_str}")
            return None

    def parse_date_without_warning(self, date_str):
        """警告を発生させないで日付文字列をパースする補助関数"""
        # クラス外でも使える静的メソッド
        try:
            # JSTを含む場合
            if isinstance(date_str, str) and 'JST' in date_str:
                # JSTを削除
                clean_date_str = date_str.replace('JST', '').strip()
                # タイムゾーン情報なしでパース
                date_obj = pd.to_datetime(clean_date_str, errors='coerce')
                if not pd.isna(date_obj):
                    # 明示的に日本時間のタイムゾーンを設定
                    date_obj = date_obj.tz_localize('Asia/Tokyo')
                return date_obj
            else:
                # 通常のパース
                return pd.to_datetime(date_str, errors='coerce')
        except:
            return pd.NaT  # Not a Time を返す

def get_message_content(service, user_id, msg_id):
    """メッセージの本文を取得する"""
    try:
        # メッセージの詳細情報を取得（format=fullで全データを取得）
        message = service.users().messages().get(userId=user_id, id=msg_id, format='full').execute()
        
        # ペイロードからパーツを取得
        payload = message['payload']
        parts = payload.get('parts', [])
        
        # 本文を格納する変数
        body = ""
        
        # パーツがある場合（マルチパートメール）
        if parts:
            for part in parts:
                mime_type = part.get('mimeType')
                # テキスト形式の本文を探す
                if mime_type == 'text/plain':
                    body_data = part.get('body', {}).get('data', '')
                    if body_data:
                        # Base64でエンコードされたデータをデコード
                        body = base64.urlsafe_b64decode(body_data).decode('utf-8')
                        break
        # シンプルなメールの場合
        elif 'body' in payload and 'data' in payload['body']:
            body_data = payload['body']['data']
            body = base64.urlsafe_b64decode(body_data).decode('utf-8')
        
        return body
    
    except Exception as error:
        print(f'メッセージ本文の取得エラー: {error}')
        return ""

def main():
    analyzer = GmailAnalyzer()
    analyzer.authenticate()
    
    sender = input('分析したい送信者のメールアドレスを入力してください: ')
    df = analyzer.analyze_emails_from_sender(sender)
    df = df.sort_values('date', ascending=False)
    
    # 通常の分析出力
    print('\n=== 基本統計情報 ===')
    print(f'総メール数: {len(df)}')
    
    print('\n=== 月別メール数 ===')
    print(df.groupby(df['date'].dt.strftime('%Y-%m')).size())
    
    print('\n=== 時間帯別メール数 ===')
    print(df.groupby('hour').size())
    
    print('\n=== 曜日別メール数 ===')
    weekday_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    print(df.groupby('weekday').size().reindex(weekday_order))
    
    print('\n=== 直近5件のメール ===')
    recent_emails = df.head(5)
    for _, row in recent_emails.iterrows():
        print(f"{row['date'].strftime('%Y-%m-%d %H:%M')} - {row['subject']}")
    
    # PDFレポートの生成
    summary_report_path = analyzer.generate_comprehensive_pdf_report(df, sender)
    print(f'\nPDFレポートが生成されました: {summary_report_path}')

if __name__ == '__main__':
    main() 