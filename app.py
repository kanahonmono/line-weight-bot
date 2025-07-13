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

# --- ユーザー情報取得 ---
def get_user_info(username):
    # Usersシートからユーザー情報を取得
    # 期待する Users シートの形式（例）
    # | ユーザー名 | 体重列 | モード列 | モード名    |
    # | -------- | ------ | ------ | --------- |
    # | 自分     | B      | C      | 筋トレモード |
    # | 母       | D      | E      | 親モード    |
    # | 父       | F      | G      | 親モード    |
    try:
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range='Users!A2:D'  # ヘッダー1行目として2行目から取得
        ).execute()
        rows = result.get('values', [])
        for row in rows:
            if len(row) >= 4 and row[0] == username:
                return {
                    'weight_col': row[1],  # 体重の列（例 "B"）
                    'mode_col': row[2],    # モード列（例 "C"）
                    'mode': row[3]         # モード名（例 "筋トレモード"）
                }
        return None
    except Exception as e:
        print(f"Error fetching user info: {e}")
        return None

# --- 日付行検索 ---
def find_date_row(date_str):
    # WeightsシートのA列（日付列）から該当日付の行番号を探す
    try:
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range='Weights!A2:A'
        ).execute()
        dates = result.get('values', [])
        for i, row in enumerate(dates, start=2):
            if row and row[0] == date_str:
                return i
        return None
    except Exception as e:
        print(f"Error finding date row: {e}")
        return None

# --- 日付行追加 ---
def append_date_row(date_str):
    # Weightsシートに日付を追加（A列のみ）
    try:
        body = {'values': [[date_str]]}
        result = sheet.values().append(
            spreadsheetId=SPREADSHEET_ID,
            range='Weights!A:A',
            valueInputOption='USER_ENTERED',
            insertDataOption='INSERT_ROWS',
            body=body
        ).execute()
        # 追加された行番号を計算
        updates = result.get('updates', {})
        updated_rows = updates.get('updatedRows', 0)
        # 末尾に追加なので行番号は今の最大行
        result_get = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range='Weights!A:A'
        ).execute()
        total_rows = len(result_get.get('values', []))
        return total_rows  # 追加された日付の行番号
    except Exception as e:
        print(f"Error appending date row: {e}")
        return None

# --- セル更新 ---
def update_cell(row, col, value):
    # 指定のシートセルに値を更新
    try:
        range_str = f"Weights!{col}{row}"
        body = {'values': [[value]]}
        sheet.values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=range_str,
            valueInputOption='USER_ENTERED',
            body=body
        ).execute()
        print(f"Updated {range_str} to {value}")
    except Exception as e:
        print(f"Error updating cell {range_str}: {e}")

# --- 体重記録 ---
def record_weight(username, date_str, weight):
    user_info = get_user_info(username)
    if not user_info:
        raise Exception(f"ユーザー情報が見つかりません: {username}")
    weight_col = user_info['weight_col']
    mode_col = user_info['mode_col']
    mode = user_info['mode']

    row = find_date_row(date_str)
    if not row:
        row = append_date_row(date_str)
        if not row:
            raise Exception("日付の追加に失敗しました。")

    # 体重記録
    update_cell(row, weight_col, weight)
    # モードも同じ行に書き込み（最新状態のため）
    update_cell(row, mode_col, mode)

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
    user_id = event.source.user_id  # 今回はユーザー名はLINE側ユーザーIDで管理しない想定

    # ここではユーザー名をLINEユーザーIDから得るのは別途必要
    # まずは「名前 登録」「体重 YYYY-MM-DD 数字」など簡易対応例
    try:
        if text.startswith("登録"):
            # 例: 登録 自分
            parts = text.split()
            if len(parts) != 2:
                reply = "登録コマンドは「登録 ユーザー名」の形式で送ってください。"
            else:
                username = parts[1]
                # 実際はUsersシートに登録処理は別途必要（ここは案内のみ）
                reply = f"{username}さんの登録を受け付けました。\n登録情報は管理者が設定してください。"
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))
            return

        if text.startswith("体重"):
            # 例: 体重 2025-07-13 65.5
            parts = text.split()
            if len(parts) != 3:
                reply = "体重記録は「体重 YYYY-MM-DD 数字」の形式で送ってください。"
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))
                return
            _, date_str, weight_str = parts
            weight = float(weight_str)
            # LINEユーザーIDからユーザー名紐付けは本来必要だが、今回は送信した名前を使う想定
            # 例: "自分"などで送信者名決める運用
            username = "自分"  # ここは仮で固定
            record_weight(username, date_str, weight)
            reply = f"{date_str}の{username}さんの体重 {weight}kgを記録しました！"
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))
            return

        # 数字だけの送信は今日の日付で自分の体重記録
        if text.replace('.', '', 1).isdigit():
            weight = float(text)
            date_str = datetime.now().strftime('%Y-%m-%d')
            username = "自分"  # 仮固定
            record_weight(username, date_str, weight)
            reply = f"今日({date_str})の{username}さんの体重 {weight}kgを記録しました！"
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))
            return

        reply = "こんにちは！\n体重記録は「体重 YYYY-MM-DD 数字」か、数字だけ送ってください。"
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))

    except Exception as e:
        print(f"Error handling message: {e}")
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text="エラーが発生しました。もう一度試してください。"))

@app.route("/", methods=["GET"])
def home():
    return "LINE Bot is running!"

if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
