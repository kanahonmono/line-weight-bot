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

# ユーザー情報取得（Usersシートから）
def get_user_info(username):
    sheet_name = 'Users'
    range_ = f"{sheet_name}!A2:D"  # ユーザー名、モード、体重列、モード列
    try:
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=range_).execute()
        values = result.get('values', [])
        for row in values:
            if len(row) >= 1 and row[0] == username:
                mode = row[1] if len(row) > 1 else ''
                weight_col = row[2] if len(row) > 2 else ''
                mode_col = row[3] if len(row) > 3 else ''
                return {
                    'username': username,
                    'mode': mode,
                    'weight_col': weight_col,
                    'mode_col': mode_col
                }
        return None
    except Exception as e:
        print(f"ユーザー情報取得エラー: {e}")
        return None

# 体重記録（Weightsシートの該当列に書き込む）
def append_weight_data(username, date, weight):
    user_info = get_user_info(username)
    if user_info is None:
        raise Exception(f"ユーザー情報が見つかりません: {username}")

    weights_sheet = 'Weights'

    if not user_info['weight_col'] or not user_info['mode_col']:
        raise Exception(f"ユーザー {username} の体重列またはモード列が設定されていません。")

    weight_col = user_info['weight_col']
    mode_col = user_info['mode_col']

    try:
        date_col_range = f"{weights_sheet}!A2:A"
        date_result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=date_col_range).execute()
        date_values = date_result.get('values', [])
        dates = [r[0] for r in date_values if r]

        if date in dates:
            row_index = dates.index(date) + 2
        else:
            row_index = len(dates) + 2
            append_date_body = {
                'values': [[date]]
            }
            sheet.values().append(
                spreadsheetId=SPREADSHEET_ID,
                range=f"{weights_sheet}!A:A",
                valueInputOption='USER_ENTERED',
                insertDataOption='INSERT_ROWS',
                body=append_date_body
            ).execute()

        weight_range = f"{weights_sheet}!{weight_col}{row_index}"
        mode_range = f"{weights_sheet}!{mode_col}{row_index}"

        body_weight = {'values': [[weight]]}
        body_mode = {'values': [[user_info['mode']]]}

        sheet.values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=weight_range,
            valueInputOption='USER_ENTERED',
            body=body_weight
        ).execute()

        sheet.values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=mode_range,
            valueInputOption='USER_ENTERED',
            body=body_mode
        ).execute()

        print(f"{username} の体重 {weight}kg を {date} に記録しました。")

    except Exception as e:
        print(f"体重記録エラー: {e}")
        raise

# ユーザー登録（Usersシートに追加）
def register_user(username, mode, weight_col, mode_col):
    sheet_name = 'Users'
    values = [[username, mode, weight_col, mode_col]]
    body = {'values': values}
    try:
        sheet.values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!A:D",
            valueInputOption='USER_ENTERED',
            insertDataOption='INSERT_ROWS',
            body=body
        ).execute()
        print(f"{username} さんを登録しました。")
    except Exception as e:
        print(f"ユーザー登録エラー: {e}")
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
    parts = text.split()

    try:
        if parts[0] == '登録':
            if len(parts) != 5:
                reply = (
                    "登録コマンドの形式:\n"
                    "登録 ユーザー名 モード(親モード/筋トレモード) 体重列 モード列\n"
                    "例）登録 かなた 筋トレモード B C"
                )
            else:
                username = parts[1]
                mode = parts[2]
                weight_col = parts[3].upper()
                mode_col = parts[4].upper()
                register_user(username, mode, weight_col, mode_col)
                reply = (
                    f"{username} さんの登録を受け付けました。\n\n"
                    "次に体重を記録するには\n"
                    f"「体重 {username} YYYY-MM-DD 体重」\n"
                    "または\n"
                    f"「体重 {username} 体重」\n"
                    "（後者は今日の日付で記録されます）"
                )
        elif parts[0] == '体重':
            if len(parts) == 4:
                username = parts[1]
                date = parts[2]
                weight = float(parts[3])
                append_weight_data(username, date, weight)
                reply = f"{username} さんの {date} の体重 {weight}kg を記録しました！"
            elif len(parts) == 3:
                username = parts[1]
                weight = float(parts[2])
                date = datetime.now().strftime('%Y-%m-%d')
                append_weight_data(username, date, weight)
                reply = f"{username} さんの今日({date})の体重 {weight}kg を記録しました！"
            else:
                reply = (
                    "体重コマンドの形式:\n"
                    "「体重 ユーザー名 YYYY-MM-DD 体重」\nまたは\n"
                    "「体重 ユーザー名 体重」"
                )
        else:
            reply = (
                "こんにちは！\n"
                "■登録例\n"
                "登録 ユーザー名 モード(親モード/筋トレモード) 体重列 モード列\n"
                "例）登録 かなた 筋トレモード B C\n\n"
                "■体重記録例\n"
                "体重 ユーザー名 YYYY-MM-DD 体重\n"
                "体重 ユーザー名 体重\n\n"
                "使い方がわからなければこのメッセージを送ってください。"
            )
    except Exception as e:
        print(f"エラー: {e}")
        reply = f"エラーが発生しました: {e}"

    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))

@app.route("/", methods=["GET"])
def home():
    return "LINE Bot is running!"

if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
