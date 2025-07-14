from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage, ImageSendMessage
import os
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime, timedelta
import pandas as pd
import matplotlib.pyplot as plt

app = Flask(__name__)

# === 環境変数から設定読み込み ===
LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET")
LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")
GOOGLE_CREDENTIALS = os.getenv("GOOGLE_APPLICATION_CREDENTIALS_JSON")
SPREADSHEET_ID = "1mmdxzloT6rOmx7SiVT4X2PtmtcsBxivcHSoMUvjDCqc"
YOUR_PUBLIC_BASE_URL = os.getenv('YOUR_PUBLIC_BASE_URL')  # 画像公開URLベース

if not (LINE_CHANNEL_SECRET and LINE_CHANNEL_ACCESS_TOKEN and GOOGLE_CREDENTIALS):
    raise Exception("環境変数が足りません")

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

credentials_dict = json.loads(GOOGLE_CREDENTIALS)
credentials = service_account.Credentials.from_service_account_info(
    credentials_dict,
    scopes=["https://www.googleapis.com/auth/spreadsheets"]
)
sheet_service = build('sheets', 'v4', credentials=credentials)
sheet = sheet_service.spreadsheets()

# === ユーザー情報取得 ===
def get_user_info_by_id(user_id):
    try:
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Users!A2:E").execute()
        for row in result.get("values", []):
            if len(row) >= 5 and row[4] == user_id:
                return {
                    "username": row[0],
                    "mode": row[1],
                    "weight_col": row[2],
                    "mode_col": row[3],
                    "user_id": row[4],
                }
        return None
    except Exception as e:
        print(f"ユーザー情報取得エラー: {e}")
        return None

# === 縦型データを全件取得 ===
def get_weight_data_vertical(username):
    try:
        # 体重記録は縦型で Users!A:D（例）としている想定。実際のシート範囲に合わせてください。
        range_name = "Weights!A2:D"
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=range_name).execute()
        rows = result.get("values", [])
        # DataFrame化しフィルタ
        df = pd.DataFrame(rows, columns=["ユーザー名", "日付", "体重", "モード"])
        df = df[df["ユーザー名"] == username]
        return df
    except Exception as e:
        print(f"体重データ取得エラー: {e}")
        return None

# === 直近1か月の体重データ取得 ===
def get_last_month_weight_data(username):
    df = get_weight_data_vertical(username)
    if df is None or df.empty:
        return None
    df['日付'] = pd.to_datetime(df['日付'])
    one_month_ago = datetime.now() - timedelta(days=30)
    df_1m = df[df['日付'] >= one_month_ago].copy()
    df_1m.sort_values('日付', inplace=True)
    return df_1m

# === 体重推移グラフ作成 ===
def create_monthly_weight_graph(df, username):
    plt.figure(figsize=(8, 4))
    plt.plot(df['日付'], df['体重'].astype(float), marker='o', linestyle='-', color='blue')
    plt.title(f"{username} さんの直近1か月の体重推移")
    plt.xlabel("日付")
    plt.ylabel("体重(kg)")
    plt.grid(True)
    plt.tight_layout()
    filename = f"/tmp/{username}_weight_1month.png"
    plt.savefig(filename)
    plt.close()
    return filename

# === LINEに画像送信 ===
def send_monthly_weight_graph_to_line(user_info):
    df_1m = get_last_month_weight_data(user_info["username"])
    if df_1m is None or df_1m.empty:
        raise Exception("直近1か月の体重データが見つかりません。")
    img_path = create_monthly_weight_graph(df_1m, user_info["username"])

    if YOUR_PUBLIC_BASE_URL is None:
        raise Exception("公開URL環境変数 YOUR_PUBLIC_BASE_URL が設定されていません。")

    img_url = f"{YOUR_PUBLIC_BASE_URL}/temp/{os.path.basename(img_path)}"

    line_bot_api.push_message(
        user_info["user_id"],
        ImageSendMessage(
            original_content_url=img_url,
            preview_image_url=img_url,
        )
    )

# === Flaskコールバック ===
@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

# === メッセージ処理 ===
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    text = event.message.text.strip()
    user_id = event.source.user_id
    parts = text.split()

    try:
        if text.lower() == "ヘルプ":
            reply = "こんにちは！\n■体重記録コマンド\n体重 65.5\n体重 2025-07-13 65.5\n登録 ユーザー名 モード\nリセット\nグラフ送信"
        elif parts[0] == "グラフ送信":
            user_info = get_user_info_by_id(user_id)
            if not user_info:
                reply = "登録されていません。先に登録してください。"
            else:
                try:
                    send_monthly_weight_graph_to_line(user_info)
                    reply = "直近1か月の体重グラフを送信しました。"
                except Exception as e:
                    reply = f"グラフ送信でエラーが発生しました: {e}"
        else:
            reply = "コマンドが正しくありません。ヘルプと送ってください。"
    except Exception as e:
        reply = f"エラーが発生しました: {e}"

    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))

@app.route("/", methods=["GET"])
def home():
    return "LINEダイエットBot起動中"

if __name__ == "__main__":
    app.run(debug=True)
