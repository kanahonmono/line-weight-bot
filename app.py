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

# ユーザー情報取得

def get_user_info(username):
    range_ = "Users!A2:D"
    try:
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=range_).execute()
        values = result.get("values", [])
        for row in values:
            if row[0] == username:
                return {
                    "username": username,
                    "mode": row[1],
                    "weight_col": row[2],
                    "mode_col": row[3],
                }
        return None
    except Exception as e:
        print(f"ユーザー情報取得エラー: {e}")
        return None

# 自動で空き列を見つける

def find_next_available_columns():
    range_ = "Users!B1:Z1"
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=range_).execute()
    header = result.get('values', [[]])[0]
    col_index = 1
    while col_index + 1 < 26:
        col1 = chr(ord('A') + col_index)
        col2 = chr(ord('A') + col_index + 1)
        if (len(header) <= col_index or header[col_index] == '') and \
           (len(header) <= col_index + 1 or header[col_index + 1] == ''):
            return col1, col2
        col_index += 2
    raise Exception("これ以上登録できる列がありません。")

# ユーザー登録（空き列に自動割り当て）

def register_user_auto(username, mode):
    weight_col, mode_col = find_next_available_columns()
    sheet.values().append(
        spreadsheetId=SPREADSHEET_ID,
        range="Users!A:D",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": [[username, mode, weight_col, mode_col]]}
    ).execute()

    sheet.values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"valueInputOption": "USER_ENTERED", "data": [
            {"range": f"Weights!{weight_col}1", "values": [[f"{username}体重"]]},
            {"range": f"Weights!{mode_col}1", "values": [[f"{username}モード"]]}
        ]}
    ).execute()
    return weight_col, mode_col

# 体重記録

def append_weight_data(username, date, weight):
    user_info = get_user_info(username)
    if user_info is None:
        raise Exception(f"ユーザー情報が見つかりません: {username}")

    weight_col = user_info['weight_col']
    mode_col = user_info['mode_col']
    weights_sheet = 'Weights'

    try:
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

    except Exception as e:
        print(f"体重記録エラー: {e}")
        raise

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

    try:
        if parts[0] == '登録' and len(parts) == 3:
            username = parts[1]
            mode = parts[2]
            weight_col, mode_col = register_user_auto(username, mode)
            reply = f"{username} さんを登録しました！\n体重列: {weight_col}, モード列: {mode_col}\n体重を記録するには：\n体重 {username} 体重 または 体重 {username} YYYY-MM-DD 体重"

        elif parts[0] == '体重':
            if len(parts) == 4:
                username = parts[1]
                date = parts[2]
                weight = float(parts[3])
            elif len(parts) == 3:
                username = parts[1]
                date = datetime.now().strftime('%Y-%m-%d')
                weight = float(parts[2])
            else:
                raise Exception("体重コマンド形式エラー")

            append_weight_data(username, date, weight)
            reply = f"{username} さんの {date} の体重 {weight}kg を記録しました！"
        else:
            reply = "コマンド例：\n登録 かなた 筋トレモード\n体重 かなた 65.5 または 体重 かなた 2025-07-13 65.5"
    except Exception as e:
        reply = f"エラーが発生しました: {e}"

    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))

@app.route("/", methods=["GET"])
def home():
    return "LINE Bot is running!"

if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
