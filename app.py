from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
import os
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime

app = Flask(__name__)

LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET")
LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")
line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

SPREADSHEET_ID = '1mmdxzloT6rOmx7SiVT4X2PtmtcsBxivcHSoMUvjDCqc'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

credentials_info = os.getenv("GOOGLE_APPLICATION_CREDENTIALS_JSON")
if credentials_info is None:
    raise Exception("環境変数 'GOOGLE_APPLICATION_CREDENTIALS_JSON' が設定されていません。")
credentials_dict = json.loads(credentials_info)
credentials = service_account.Credentials.from_service_account_info(
    credentials_dict,
    scopes=SCOPES
)
service = build('sheets', 'v4', credentials=credentials)
sheet = service.spreadsheets()

def append_weight_data(user_id, date, weight):
    values = [[user_id, date, weight]]
    body = {'values': values}
    result = sheet.values().append(
        spreadsheetId=SPREADSHEET_ID,
        range='Sheet1!A1:C1',
        valueInputOption='USER_ENTERED',
        insertDataOption='INSERT_ROWS',
        body=body
    ).execute()
    print(f"{result.get('updates').get('updatedRows')} rows appended.")

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers.get('X-Line-Signature')
    body = request.get_data(as_text=True)
    app.logger.info(f"Request body: {body}")

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    text = event.message.text.strip()
    user_id = event.source.user_id

    if text.startswith("体重"):
        try:
            parts = text.split()
            if len(parts) != 3:
                raise ValueError("フォーマットエラー")
            _, date, weight_str = parts
            weight = float(weight_str)
            append_weight_data(user_id, date, weight)
            reply = f"{date} の体重 {weight}kg を記録しました！"
        except Exception:
            reply = "記録に失敗しました。正しいフォーマットは「体重 YYYY-MM-DD 数字」です。"
    elif text.replace('.', '', 1).isdigit():
        weight = float(text)
        date = datetime.now().strftime('%Y-%m-%d')
        append_weight_data(user_id, date, weight)
        reply = f"今日({date})の体重 {weight}kg を記録しました！"
    else:
        reply = "こんにちは！体重を記録するには「体重 YYYY-MM-DD 数字」か、数字だけ送ってください。"

    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))

@app.route("/", methods=["GET"])
def home():
    return "LINE Bot is running!"

if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port)