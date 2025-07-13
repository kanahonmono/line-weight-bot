from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
import os
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime
import openai

app = Flask(__name__)

# LINE
LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET")
LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")
if not LINE_CHANNEL_SECRET or not LINE_CHANNEL_ACCESS_TOKEN:
    raise Exception("LINE_CHANNEL_SECRETまたはLINE_CHANNEL_ACCESS_TOKENが環境変数に設定されていません。")

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# Google Sheets
SPREADSHEET_ID = '1mmdxzloT6rOmx7SiVT4X2PtmtcsBxivcHSoMUvjDCqc'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

credentials_info = os.getenv("GOOGLE_APPLICATION_CREDENTIALS_JSON")
if not credentials_info:
    raise Exception("環境変数 'GOOGLE_APPLICATION_CREDENTIALS_JSON' が設定されていません。")
credentials_dict = json.loads(credentials_info)
credentials = service_account.Credentials.from_service_account_info(credentials_dict, scopes=SCOPES)
service = build('sheets', 'v4', credentials=credentials)
sheet = service.spreadsheets()

# OpenAI
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
if not OPENAI_API_KEY:
    raise Exception("環境変数 'OPENAI_API_KEY' が設定されていません。")
openai.api_key = OPENAI_API_KEY

# シート名・範囲
USERS_SHEET = 'Users'       # ユーザー登録用シート
WEIGHTS_SHEET = 'Weights'   # 体重記録用シート

# --- GPTで登録案内メッセージ生成 ---
def generate_registration_guide(name):
    prompt = (
        f"あなたは親切なダイエットコーチです。"
        f"ユーザー名「{name}」さんに向けて、"
        f"名前登録完了の案内とこれから何をすればいいかを分かりやすく3文以内で説明してください。"
        f"親しみやすく優しい口調でお願いします。"
    )
    response = openai.ChatCompletion.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "親切なダイエットBotの登録案内を作る"},
            {"role": "user", "content": prompt}
        ],
        max_tokens=100,
        temperature=0.7,
    )
    return response.choices[0].message.content.strip()

# --- Google Sheets 操作関数 ---
def get_users_data():
    """Usersシートの全ユーザー情報を辞書で取得"""
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=f"{USERS_SHEET}!A2:D").execute()
    values = result.get('values', [])
    users = {}
    for row in values:
        if len(row) < 4:
            continue
        name, mode, weight_col, mode_col = row[0], row[1], row[2], row[3]
        users[name] = {'mode': mode, 'weight_col': weight_col, 'mode_col': mode_col}
    return users

def register_user(name, mode):
    """Usersシートに新規ユーザー登録"""
    # Usersシートの最後の行に追加
    # 例: 名前, モード, 体重列, モード列
    # ここでは簡単に空き列は固定 or 別途管理で拡張可
    # 仮にA-D列はヘッダーなので2行目以降に追加
    # 体重列・モード列は仮にB,C,D,E,F,Gなど割り当ててください

    users = get_users_data()
    if name in users:
        return False  # すでに登録済み

    # 使われているweight_colを取得
    used_cols = [v['weight_col'] for v in users.values()]
    # A〜Z列から未使用を探す（B〜Zあたり）
    import string
    candidate_cols = list(string.ascii_uppercase[1:])  # B〜Z
    available_cols = [c for c in candidate_cols if c not in used_cols]
    if len(available_cols) < 2:
        raise Exception("空き列が足りません")

    weight_col = available_cols[0]
    mode_col = available_cols[1]

    new_row = [name, mode, weight_col, mode_col]
    append_body = {'values': [new_row]}
    sheet.values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{USERS_SHEET}!A2",
        valueInputOption='USER_ENTERED',
        insertDataOption='INSERT_ROWS',
        body=append_body
    ).execute()
    return True

def append_weight_data(name, date, weight):
    """Weightsシートに体重とモードを追加"""
    # 日付列はA列固定、以降は
    # 例: B=自分体重, C=自分モード, D=母体重, E=母モード, F=父体重, G=父モード

    users = get_users_data()
    if name not in users:
        raise Exception(f"{name}は登録されていません。")

    # 日付がすでにあるか検索
    res = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=f"{WEIGHTS_SHEET}!A2:A").execute()
    dates = res.get('values', [])
    date_str = date
    dates_flat = [d[0] for d in dates]

    # その日の日付の行を探す or 最終行+1に追加
    if date_str in dates_flat:
        row_index = dates_flat.index(date_str) + 2  # シートは1始まり、ヘッダーあり
    else:
        # 日付新規追加行
        row_index = len(dates_flat) + 2
        # 日付を追加
        sheet.values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{WEIGHTS_SHEET}!A{row_index}",
            valueInputOption='USER_ENTERED',
            body={'values': [[date_str]]}
        ).execute()

    # ユーザーの体重列、モード列を取得
    weight_col = users[name]['weight_col']
    mode_col = users[name]['mode_col']
    mode = users[name]['mode']

    # 体重とモードを書き込み
    sheet.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{WEIGHTS_SHEET}!{weight_col}{row_index}",
        valueInputOption='USER_ENTERED',
        body={'values': [[weight]]}
    ).execute()

    sheet.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{WEIGHTS_SHEET}!{mode_col}{row_index}",
        valueInputOption='USER_ENTERED',
        body={'values': [[mode]]}
    ).execute()

# --- Flask / LINE --- 
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

    users = get_users_data()

    # 登録コマンド
    if text.startswith("登録"):
        try:
            # 例: 登録 かなた 筋トレモード
            parts = text.split()
            if len(parts) != 3:
                raise ValueError("フォーマットエラー")
            _, name, mode = parts
            if mode not in ["親モード", "筋トレモード"]:
                raise ValueError("モードは「親モード」か「筋トレモード」のどちらかです。")

            if name in users:
                reply = f"{name}さんはすでに登録されています。"
            else:
                register_user(name, mode)
                guide_msg = generate_registration_guide(name)
                reply = f"{name}さんを「{mode}」で登録しました！\n\n{guide_msg}"
        except Exception as e:
            reply = f"登録に失敗しました。使い方：\n登録 名前 モード\n例: 登録 かなた 筋トレモード\nモードは「親モード」か「筋トレモード」のどちらかです。\n\n詳細: {str(e)}"

    # 体重記録コマンド
    elif text.startswith("体重"):
        try:
            # 例: 体重 2025-07-13 66
            parts = text.split()
            if len(parts) != 3:
                raise ValueError("フォーマットエラー")
            _, date, weight_str = parts
            weight = float(weight_str)

            # 名前はuser_idからは取れないのでここは簡単に「名前: user_id」形式にしておくか、
            # 実際はユーザー登録時に名前管理する工夫が必要

            # 今は仮に名前を一つだけに絞るなど工夫要（ここでは名前が一人だけの想定）
            if len(users) == 0:
                reply = "まず登録してください。例：登録 かなた 筋トレモード"
            elif len(users) == 1:
                name = list(users.keys())[0]
                append_weight_data(name, date, weight)
                reply = f"{date} の体重 {weight}kg を記録しました！"
            else:
                reply = "複数ユーザーがいます。体重記録はユーザー名を付けて送ってください。\n例: 体重 かなた 2025-07-13 66"
        except Exception as e:
            reply = f"体重記録に失敗しました。フォーマットは「体重 YYYY-MM-DD 数字」です。\n詳細: {str(e)}"

    # 複数ユーザー対応の体重記録
    elif text.startswith("体重"):
        # 例: 体重 かなた 2025-07-13 66
        parts = text.split()
        if len(parts) == 4:
            _, name, date, weight_str = parts
            try:
                weight = float(weight_str)
                users = get_users_data()
                if name not in users:
                    reply = f"{name}さんは登録されていません。先に登録してください。"
                else:
                    append_weight_data(name, date, weight)
                    reply = f"{date} の{name}さんの体重 {weight}kg を記録しました！"
            except Exception as e:
                reply = f"体重記録に失敗しました。\n詳細: {str(e)}"
        else:
            reply = "体重記録のフォーマットが違います。\n複数ユーザーの場合は「体重 名前 YYYY-MM-DD 数字」で送信してください。"

    else:
        reply = (
            "こんにちは！ダイエットBotへようこそ。\n"
            "まずは登録してください。\n"
            "登録の例: 登録 かなた 筋トレモード\n"
            "体重記録の例: 体重 2025-07-13 66\n"
            "複数ユーザーの場合: 体重 名前 YYYY-MM-DD 数字\n"
            "わからなければ「助けて」と送ってね。"
        )

    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))

@app.route("/", methods=["GET"])
def home():
    return "LINE Bot is running!"

if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
