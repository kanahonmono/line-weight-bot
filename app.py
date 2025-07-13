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

# Excel列記号を0始まりのインデックスに変換（A→0, B→1, ...）
def col_letter_to_index(letter):
    letter = letter.upper()
    result = 0
    for ch in letter:
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result - 1

def get_users():
    """Usersシートから登録情報を取得"""
    result = sheet.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range='Users!A2:D'  # ヘッダー除く
    ).execute()
    values = result.get('values', [])
    users = {}
    for row in values:
        if len(row) < 4:
            continue
        name, mode, weight_col, mode_col = row[0], row[1], row[2], row[3]
        users[name] = {
            'mode': mode,
            'weight_col': weight_col,
            'mode_col': mode_col
        }
    return users

def update_user_mode_in_users(name, mode):
    """Usersシートのモードを更新"""
    # まずUsers全体取得して何行目か探す
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range='Users!A2:D').execute()
    values = result.get('values', [])
    for i, row in enumerate(values):
        if len(row) > 0 and row[0] == name:
            update_range = f'Users!B{i+2}'  # モードはB列、2行目から開始
            sheet.values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=update_range,
                valueInputOption='USER_ENTERED',
                body={'values': [[mode]]}
            ).execute()
            return True
    return False

def append_or_update_weight(name, date_str, weight, users):
    """Weightsシートの該当日付の体重・モード列を更新、なければ追加"""
    # 全Weights取得（1行目はヘッダー）
    result = sheet.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range='Weights!A2:Z1000'  # 2行目以降
    ).execute()
    values = result.get('values', [])
    weight_col_idx = col_letter_to_index(users[name]['weight_col'])
    mode_col_idx = col_letter_to_index(users[name]['mode_col'])
    mode = users[name]['mode']

    # 日付をA列（index0）から検索
    target_row_idx = None
    for i, row in enumerate(values):
        if len(row) > 0 and row[0] == date_str:
            target_row_idx = i + 2  # Sheetsの実行行番号
            break

    if target_row_idx is None:
        # 新規行を追加
        new_row = [''] * (max(weight_col_idx, mode_col_idx) + 1)
        new_row[0] = date_str
        new_row[weight_col_idx] = str(weight)
        new_row[mode_col_idx] = mode
        # append
        sheet.values().append(
            spreadsheetId=SPREADSHEET_ID,
            range='Weights!A2',
            valueInputOption='USER_ENTERED',
            insertDataOption='INSERT_ROWS',
            body={'values': [new_row]}
        ).execute()
    else:
        # 既存行の更新
        # 更新範囲は行と列を指定する必要があるため、1セルずつ更新
        weight_cell = f"Weights!{users[name]['weight_col']}{target_row_idx}"
        mode_cell = f"Weights!{users[name]['mode_col']}{target_row_idx}"

        sheet.values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=weight_cell,
            valueInputOption='USER_ENTERED',
            body={'values': [[str(weight)]]}
        ).execute()
        sheet.values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=mode_cell,
            valueInputOption='USER_ENTERED',
            body={'values': [[mode]]}
        ).execute()

def get_status_text(users):
    today = datetime.now().strftime('%Y-%m-%d')
    # Weightsシートから今日のデータを取得
    result = sheet.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f'Weights!A2:Z1000'
    ).execute()
    values = result.get('values', [])

    # 今日の行を探す
    today_row = None
    for row in values:
        if len(row) > 0 and row[0] == today:
            today_row = row
            break

    if not today_row:
        return "今日の体重記録はまだありません。"

    texts = [f"【{today}の体重状況】"]
    for name, info in users.items():
        w_idx = col_letter_to_index(info['weight_col'])
        m_idx = col_letter_to_index(info['mode_col'])
        weight = today_row[w_idx] if w_idx < len(today_row) else "未記録"
        mode = today_row[m_idx] if m_idx < len(today_row) else info['mode']
        texts.append(f"{name}：{weight}kg（{mode}）")
    return "\n".join(texts)

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
    users = get_users()

    if text.startswith("登録"):
        # 登録 名前 モード 例）登録 かなた 親モード
        parts = text.split()
        if len(parts) != 3:
            reply = "登録コマンドは「登録 名前 モード」です。\n例）登録 かなた 親モード"
        else:
            _, name, mode = parts
            if mode not in ["親モード", "筋トレモード"]:
                reply = "モードは「親モード」か「筋トレモード」で指定してください。"
            elif name in users:
                reply = f"{name}さんはすでに登録されています。"
            else:
                # 体重列・モード列の割当を決める（Usersの空き列から自動割当するロジック）
                # 例: B/C, D/E, F/G, H/I, ... を順に割り当てる

                used_weight_cols = [u['weight_col'] for u in users.values()]
                used_mode_cols = [u['mode_col'] for u in users.values()]

                # 列はB=2, D=4, F=6... のように奇数を飛ばして割り当てる例
                # ここではB,C=2,3 D,E=4,5 F,G=6,7 で決定
                candidate_cols = [('B', 'C'), ('D', 'E'), ('F', 'G'), ('H', 'I'), ('J', 'K')]

                for w_col, m_col in candidate_cols:
                    if w_col not in used_weight_cols and m_col not in used_mode_cols:
                        weight_col = w_col
                        mode_col = m_col
                        break
                else:
                    reply = "これ以上登録できません（列の割当不足）"
                    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))
                    return

                # Usersシートに追加
                append_values = [[name, mode, weight_col, mode_col]]
                sheet.values().append(
                    spreadsheetId=SPREADSHEET_ID,
                    range='Users!A2:D',
                    valueInputOption='USER_ENTERED',
                    insertDataOption='INSERT_ROWS',
                    body={'values': append_values}
                ).execute()

                reply = (f"{name}さんを{mode}で登録しました！\n\n"
                         f"体重・モード列はWeightsシートの{weight_col}列と{mode_col}列に記録されます。\n"
                         "【次の操作】\n"
                         "・体重を記録するには「体重 YYYY-MM-DD 体重」の形式で送信してください。\n"
                         "・今日の体重なら「体重 体重」のみで記録できます。\n"
                         "・モード変更は「モード変更 名前 モード」で送信してください。\n"
                         "・「状況」で今日の記録を確認できます。")

    elif text.startswith("モード変更"):
        # モード変更 名前 モード 例）モード変更 かなた 筋トレモード
        parts = text.split()
        if len(parts) != 3:
            reply = "モード変更コマンドは「モード変更 名前 モード」です。\n例）モード変更 かなた 親モード"
        else:
            _, name, mode = parts
            if mode not in ["親モード", "筋トレモード"]:
                reply = "モードは「親モード」か「筋トレモード」で指定してください。"
            elif name not in users:
                reply = f"{name}さんはまだ登録されていません。"
            else:
                update_user_mode_in_users(name, mode)
                # Weightsシートのモード列も更新したい
                # ここでは更新省略（必要なら追加実装）
                reply = f"{name}さんのモードを{mode}に変更しました。"

    elif text.startswith("体重"):
        # 体重 [日付] [数字] または 体重 [数字]
        parts = text.split()
        if len(parts) == 3:
            _, date_str, weight_str = parts
            name = None
            reply = "登録されている名前を最初に入力してください。\n例）体重 かなた 2025-07-13 65.5"
        elif len(parts) == 2:
            _, weight_str = parts
            # 今日の日付で記録、名前必要
            name = None
            reply = "登録されている名前を最初に入力してください。\n例）体重 かなた 65.5"
        elif len(parts) == 4:
            # 例）体重 かなた 2025-07-13 65.5
            _, name, date_str, weight_str = parts
            if name not in users:
                reply = f"{name}さんは登録されていません。先に登録してください。"
            else:
                try:
                    datetime.strptime(date_str, '%Y-%m-%d')
                    weight = float(weight_str)
                    append_or_update_weight(name, date_str, weight, users)
                    reply = f"{name}さんの{date_str}の体重 {weight}kg を記録しました。"
                except Exception:
                    reply = "日付や体重の形式が正しくありません。"
        elif len(parts) == 3:
            # 例）体重 かなた 65.5 今日の日付
            _, name, weight_str = parts
            if name not in users:
                reply = f"{name}さんは登録されていません。先に登録してください。"
            else:
                try:
                    weight = float(weight_str)
                    date_str = datetime.now().strftime('%Y-%m-%d')
                    append_or_update_weight(name, date_str, weight, users)
                    reply = f"{name}さんの今日({date_str})の体重 {weight}kg を記録しました。"
                except Exception:
                    reply = "体重の形式が正しくありません。"
        else:
            reply = "体重記録の形式が違います。\n\n- 「体重 名前 YYYY-MM-DD 体重」\n- 「体重 名前 体重」\nで送信してください。"

    elif text == "状況":
        users = get_users()
        reply = get_status_text(users)

    else:
        reply = (
            "【使い方】\n"
            "1. 登録：登録 名前 モード（親モード or 筋トレモード）\n"
            "　例）登録 かなた 親モード\n"
            "2. 体重記録：\n"
            "　・体重 名前 YYYY-MM-DD 体重\n"
            "　・体重 名前 体重（今日の日付で登録）\n"
            "3. モード変更：モード変更 名前 モード\n"
            "　例）モード変更 かなた 筋トレモード\n"
            "4. 状況：今日の体重とモードを一覧表示\n"
        )

    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))

@app.route("/", methods=["GET"])
def home():
    return "LINEダイエットBotが稼働中です！"

if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
