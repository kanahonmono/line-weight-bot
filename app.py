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

# --- 環境変数 ---
LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET")
LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")
GOOGLE_CREDENTIALS = os.getenv("GOOGLE_APPLICATION_CREDENTIALS_JSON")
SPREADSHEET_ID = "1mmdxzloT6rOmx7SiVT4X2PtmtcsBxivcHSoMUvjDCqc"

if not (LINE_CHANNEL_SECRET and LINE_CHANNEL_ACCESS_TOKEN and GOOGLE_CREDENTIALS):
    raise Exception("環境変数が設定されていません。")

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

credentials_dict = json.loads(GOOGLE_CREDENTIALS)
credentials = service_account.Credentials.from_service_account_info(
    credentials_dict,
    scopes=["https://www.googleapis.com/auth/spreadsheets"]
)
sheet_service = build('sheets', 'v4', credentials=credentials)
sheet = sheet_service.spreadsheets()

# --- ユーザー情報取得（user_idで検索） ---
def get_user_info_by_id(user_id):
    try:
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Users!A2:D").execute()
        for idx, row in enumerate(result.get("values", []), start=2):
            # row = [username, mode, user_id]
            if len(row) >= 3 and row[2] == user_id:
                return {
                    "username": row[0],
                    "mode": row[1],
                    "row_num": idx
                }
        return None
    except Exception as e:
        print(f"ユーザー情報取得エラー: {e}")
        return None

# --- ユーザー登録 ---
def register_user(username, mode, user_id):
    # すでに登録されているか確認
    existing = get_user_info_by_id(user_id)
    if existing:
        return f"{existing['username']}さんはすでに登録されています。"
    
    # Usersシートに新規追加（username, mode, user_id）
    values = [[username, mode, user_id]]
    sheet.values().append(
        spreadsheetId=SPREADSHEET_ID,
        range="Users!A2:C",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": values}
    ).execute()
    return f"{username}さんを登録しました。"

# --- ユーザーリセット（登録情報削除＋体重記録は残す） ---
def reset_user(user_id):
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Users!A2:C").execute()
    values = result.get("values", [])
    for i, row in enumerate(values, start=2):
        if len(row) >= 3 and row[2] == user_id:
            # 登録行を空白にする
            sheet.values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=f"Users!A{i}:C{i}",
                valueInputOption="USER_ENTERED",
                body={"values": [["", "", ""]]}
            ).execute()
            return "登録情報をリセットしました。体重記録は残ります。"
    return "ユーザー登録が見つかりませんでした。"

# --- 体重記録追加（縦型） ---
def append_weight(user_id, weight, date=None):
    user_info = get_user_info_by_id(user_id)
    if not user_info:
        return "登録されていません。まず「登録 ユーザー名 モード」で登録してください。"
    
    if date is None:
        date = datetime.now().strftime('%Y-%m-%d')
    else:
        # 日付フォーマット簡易チェック（yyyy-mm-dd）
        try:
            datetime.strptime(date, '%Y-%m-%d')
        except:
            return "日付の形式が不正です。例：2025-07-13"

    # 体重記録シートに1行追加
    values = [[user_info['username'], date, weight, user_info['mode']]]
    sheet.values().append(
        spreadsheetId=SPREADSHEET_ID,
        range="Weights!A2:D",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": values}
    ).execute()

    return f"{user_info['username']}さんの体重 {weight}kg を {date} に記録しました。"

# --- LINE Bot でのメッセージ処理 ---
@app.route("/callback", methods=["POST"])
def callback():
    signature = request.headers.get("X-Line-Signature")
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return "OK"

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    text = event.message.text.strip()
    user_id = event.source.user_id
    parts = text.split()

    reply = ""
    try:
        if text == "ヘルプ":
            reply = (
                "こんにちは！\n"
                "■体重記録コマンド\n"
                "体重 65.5\n"
                "体重 2025-07-13 65.5\n"
                "■登録コマンド\n"
                "登録 ユーザー名 モード\n"
                "■リセットコマンド\n"
                "リセット\n"
                "例）体重 かなた 65.5 （名前省略は登録済みのLINE IDの場合のみ）"
            )
        elif parts[0] == "登録" and len(parts) == 3:
            username = parts[1]
            mode = parts[2]
            reply = register_user(username, mode, user_id)
        elif parts[0] == "リセット":
            reply = reset_user(user_id)
        elif parts[0] == "体重":
            if len(parts) == 2:
                # 例: 体重 65.5 （登録済みユーザー名省略）
                weight = float(parts[1])
                reply = append_weight(user_id, weight)
            elif len(parts) == 3:
                # 例: 体重 2025-07-13 65.5
                date = parts[1]
                weight = float(parts[2])
                reply = append_weight(user_id, weight, date)
            elif len(parts) == 4:
                # 例: 体重 ユーザー名 2025-07-13 65.5（名前あり・未対応、基本LINE IDで管理推奨）
                reply = "体重コマンドは「体重 体重」または「体重 日付 体重」の形式で送信してください。"
            else:
                reply = "体重コマンドの形式が正しくありません。"
        else:
            reply = "コマンドが正しくありません。ヘルプと送信してください。"
    except Exception as e:
        reply = f"エラーが発生しました: {e}"

    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))

@app.route("/", methods=["GET"])
def home():
    return "LINEダイエットBot 起動中"

if __name__ == "__main__":
    app.run(debug=True)
