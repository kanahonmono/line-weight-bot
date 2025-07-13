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
if not LINE_CHANNEL_SECRET or not LINE_CHANNEL_ACCESS_TOKEN:
    raise Exception("LINE_CHANNEL_SECRETまたはLINE_CHANNEL_ACCESS_TOKENが環境変数に設定されていません。")

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

SPREADSHEET_ID = '1mmdxzloT6rOmx7SiVT4X2PtmtcsBxivcHSoMUvjDCqc'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

credentials_info = os.getenv("GOOGLE_APPLICATION_CREDENTIALS_JSON")
if not credentials_info:
    raise Exception("環境変数 'GOOGLE_APPLICATION_CREDENTIALS_JSON' が設定されていません。")
credentials_dict = json.loads(credentials_info)
credentials = service_account.Credentials.from_service_account_info(
    credentials_dict,
    scopes=SCOPES
)

service = build('sheets', 'v4', credentials=credentials)
sheet = service.spreadsheets()

def get_registered_users():
    # 登録ユーザー一覧をGoogle Sheetsから取得
    try:
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range='登録!A2:C'  # 登録シートでA列: user_id, B列: 名前, C列: モード
        ).execute()
        values = result.get('values', [])
        # user_id をキーにして {user_id: {"name": 名前, "mode": モード}} の辞書を作成
        users = {row[0]: {"name": row[1], "mode": row[2]} for row in values if len(row) >= 3}
        return users
    except Exception as e:
        print(f"登録ユーザー取得エラー: {e}")
        return {}

def is_user_registered(user_id):
    users = get_registered_users()
    return user_id in users

def register_user(user_id, name, mode):
    # 登録シートにユーザー情報を追加
    values = [[user_id, name, mode]]
    body = {'values': values}
    try:
        result = sheet.values().append(
            spreadsheetId=SPREADSHEET_ID,
            range='登録!A:C',
            valueInputOption='USER_ENTERED',
            insertDataOption='INSERT_ROWS',
            body=body
        ).execute()
        print(f"{result.get('updates', {}).get('updatedRows', 0)} rows appended to 登録.")
        return True
    except Exception as e:
        print(f"登録エラー: {e}")
        return False

def append_weight_data(user_id, date, weight):
    values = [[user_id, date, weight]]
    body = {'values': values}
    try:
        result = sheet.values().append(
            spreadsheetId=SPREADSHEET_ID,
            range='体重!A:C',  # 体重記録シート、A列: user_id, B列: 日付, C列: 体重
            valueInputOption='USER_ENTERED',
            insertDataOption='INSERT_ROWS',
            body=body
        ).execute()
        print(f"{result.get('updates', {}).get('updatedRows', 0)} rows appended to 体重.")
    except Exception as e:
        print(f"体重記録エラー: {e}")
        raise

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers.get('X-Line-Signature')
    body = request.get_data(as_text=True)
    app.logger.info(f"Request body: {body}")

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    except Exception as e:
        app.logger.error(f"Exception in handler: {e}")
        abort(500)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    text = event.message.text.strip()
    user_id = event.source.user_id

    if not is_user_registered(user_id):
        # 未登録ユーザーへの対応
        if text.startswith("登録"):
            parts = text.split()
            if len(parts) == 3:
                _, name, mode = parts
                mode = mode.lower()
                if mode not in ["親", "自分"]:
                    reply = "モードは「親」か「自分」で指定してください。例：登録 太郎 親"
                else:
                    success = register_user(user_id, name, mode)
                    if success:
                        reply = f"{name}さん、モード「{mode}」で登録が完了しました！\n体重は「体重 YYYY-MM-DD 数字」か数字だけ送ってください。"
                    else:
                        reply = "登録に失敗しました。後でもう一度試してください。"
            else:
                reply = (
                    "登録方法：\n"
                    "「登録 名前 モード」で登録してください。\n"
                    "例）登録 太郎 親\n"
                    "モードは「親」か「自分」で指定してください。"
                )
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))
        else:
            reply = (
                "はじめまして！まずは登録してください。\n"
                "登録方法：「登録 名前 モード」\n"
                "例）登録 太郎 親\n"
                "モードは「親」か「自分」で指定してください。"
            )
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))
        return

    # 登録済みユーザーの処理
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
        try:
            weight = float(text)
            date = datetime.now().strftime('%Y-%m-%d')
            append_weight_data(user_id, date, weight)
            reply = f"今日({date})の体重 {weight}kg を記録しました！"
        except Exception:
            reply = "記録に失敗しました。もう一度試してください。"
    else:
        reply = "こんにちは！体重を記録するには「体重 YYYY-MM-DD 数字」か、数字だけ送ってください。"

    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))

@app.route("/", methods=["GET"])
def home():
    return "LINE Bot is running!"

if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
