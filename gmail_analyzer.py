from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from datetime import datetime
import pandas as pd

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

class GmailAnalyzer:
    def __init__(self):
        self.creds = None
        self.service = None
        
    def authenticate(self):
        # トークンが既に存在する場合は再利用
        try:
            self.creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        except:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            self.creds = flow.run_local_server(port=0)
            # トークンを保存
            with open('token.json', 'w') as token:
                token.write(self.creds.to_json())
                
        self.service = build('gmail', 'v1', credentials=self.creds)
    
    def analyze_emails_from_sender(self, sender_email):
        results = []
        query = f'from:{sender_email}'
        
        response = self.service.users().messages().list(
            userId='me', q=query).execute()
            
        messages = response.get('messages', [])
        
        for message in messages:
            msg = self.service.users().messages().get(
                userId='me', id=message['id']).execute()
            
            headers = msg['payload']['headers']
            subject = next(h['value'] for h in headers if h['name'] == 'Subject')
            date = int(msg['internalDate']) / 1000
            date = datetime.fromtimestamp(date)
            
            results.append({
                'date': date,
                'subject': subject,
                'hour': date.hour,
                'weekday': date.strftime('%A')
            })
            
        return pd.DataFrame(results)

def main():
    analyzer = GmailAnalyzer()
    analyzer.authenticate()
    
    sender = input('分析したい送信者のメールアドレスを入力してください: ')
    df = analyzer.analyze_emails_from_sender(sender)
    
    df = df.sort_values('date', ascending=False)
    
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

if __name__ == '__main__':
    main() 