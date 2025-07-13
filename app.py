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

# LINE API設定
LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET")
LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")
if not LINE_CHANNEL_SECRET or not LINE_CHANNEL_ACCESS_TOKEN:
    raise Exception("LINE_CHANNEL_SECRETまたはLINE_CHANNEL_ACCESS_TOKENが設定されていません。")

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# Google Sheets設定
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

# LINEユーザーIDからユーザー情報取得

def get_user_info_by_line_id(line_user_id):
    range_ = "Users!A2:E"
    try:
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=range_).execute()
        values = result.get("values", [])
        for row in values:
            if len(row) >= 5 and row[4] == line_user_id:
                return {
                    "username": row[0],
                    "mode": row[1],
                    "weight_col": row[2],
                    "mode_col": row[3],
                }
        return None
    except Exception as e:
        print(f"LINE IDでユーザー取得失敗: {e}")
        return None

# ユーザー情報取得（名前指定）

def get_user_info(username):
    range_ = "Users!A2:E"
    try:
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=range_).execute()
        values = result.get("values", [])
        for row in values:
            if row[0] == username:
                return {
                    "username": row[0],
                    "mode": row[1],
                    "weight_col": row[2],
                    "mode_col": row[3],
                }
        return None
    except Exception as e:
        print(f"ユーザー情報取得エラー: {e}")
        return None

# ユーザー登録

def register_user(username, mode, weight_col, mode_col, line_user_id):
    sheet.values().append(
        spreadsheetId=SPREADSHEET_ID,
        range="Users!A:E",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": [[username, mode, weight_col, mode_col, line_user_id]]}
    ).execute()

    sheet.values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"valueInputOption": "USER_ENTERED", "data": [
            {"range": f"Weights!{weight_col}1", "values": [[f"{username}体重"]]},
            {"range": f"Weights!{mode_col}1", "values": [[f"{username}モード"]]}
        ]}
    ).execute()

# リセット（ユーザー削除）

def delete_user(username):
    range_ = "Users!A2:E"
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=range_).execute()
    values = result.get("values", [])
    new_values = [row for row in values if row[0] != username]

    sheet.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_,
        valueInputOption="USER_ENTERED",
        body={"values": new_values}
    ).execute()

# 体重記録

def append_weight_data(username, date, weight):
    user_info = get_user_info(username)
    if user_info is None:
        raise Exception(f"ユーザー情報が見つかりません: {username}")

    weight_col = user_info['weight_col']
    mode_col = user_info['mode_col']
    weights_sheet = 'Weights'

    date_col_range = f"{weights_sheet}!A2:A"
    date_result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=date_col_range).execute()
    dates = [r[0] for r in date_result.get('values', []) if r]

    if date in dates:
        row_index = dates.index(date) + 2
    else:
        row_index = len(dates) + 2
        sheet.values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{weights_sheet}!A:A",
            valueInputOption='USER_ENTERED',
            insertDataOption='INSERT_ROWS',
            body={'values': [[date]]}
        ).execute()

    sheet.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{weights_sheet}!{weight_col}{row_index}",
        valueInputOption='USER_ENTERED',
        body={'values': [[weight]]}
    ).execute()

    sheet.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{weights_sheet}!{mode_col}{row_index}",
        valueInputOption='USER_ENTERED',
        body={'values': [[user_info['mode']]]}
    ).execute()

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers.get('X-Line-Signature')
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    text = event.message.text.strip()
    parts = text.split()
    line_user_id = event.source.user_id

    try:
        if parts[0] == '体重':
            if len(parts) == 4:
                username = parts[1]
                date = parts[2]
                weight = float(parts[3])
            elif len(parts) == 3:
                username = parts[1]
                date = datetime.now().strftime('%Y-%m-%d')
                weight = float(parts[2])
            elif len(parts) == 2:
                user_info = get_user_info_by_line_id(line_user_id)
                if not user_info:
                    raise Exception("登録が必要です。管理者に依頼してください。")
                username = user_info["username"]
                date = datetime.now().strftime('%Y-%m-%d')
                weight = float(parts[1])
            else:
                raise Exception("体重コマンド形式エラー")

            append_weight_data(username, date, weight)
            reply = f"{username} さんの {date} の体重 {weight}kg を記録しました！"

        elif parts[0] == 'リセット' and len(parts) == 2:
            delete_user(parts[1])
            reply = f"{parts[1]} さんの登録を削除しました。"

        else:
            reply = "こんにちは！\n■体重記録コマンド\n体重 ユーザー名 YYYY-MM-DD 体重\n体重 ユーザー名 体重\n体重 体重\n例）体重 かなた 2025-07-13 65.5\n例）体重 かなた 65.5\n登録がまだの方は管理者に登録を依頼してください。"

    except Exception as e:
        reply = f"エラーが発生しました: {e}"

    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))

@app.route("/", methods=["GET"])
def home():
    return "LINE Bot is running!"

if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
