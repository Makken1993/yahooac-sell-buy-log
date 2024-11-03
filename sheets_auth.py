import sys
import os
import os.path
import re
import appdirs
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# スコープの設定
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# アプリケーションのデータディレクトリを取得
app_name = "ヤフオク売買履歴取得ツール"
app_author = "YourCompanyName"
app_data_dir = appdirs.user_data_dir(app_name, app_author)

# ディレクトリが存在しない場合は作成
os.makedirs(app_data_dir, exist_ok=True)

# token.jsonのフルパスを生成
token_path = os.path.join(app_data_dir, 'token.json')


def get_client_secret_file():
    if getattr(sys, 'frozen', False):
        # アプリケーションが実行可能ファイルとして実行されている場合
        application_path = sys._MEIPASS
    else:
        # 通常のPythonスクリプトとして実行されている場合
        application_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(application_path, 'client_secret.json')


def get_sheets_service():
    creds = None
    try:
        if os.path.exists(token_path):
            creds = Credentials.from_authorized_user_file(token_path, SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                except Exception:
                    os.remove(token_path)
                    return get_sheets_service()  # 再帰的に呼び出して新しい認証を開始
            else:
                client_secret_file = get_client_secret_file()
                flow = InstalledAppFlow.from_client_secrets_file(client_secret_file, SCOPES)
                creds = flow.run_local_server(port=0)

            with open(token_path, 'w') as token:
                token.write(creds.to_json())

        return build('sheets', 'v4', credentials=creds)

    except HttpError as error:
        print(f"An error occurred: {error}")
        if os.path.exists(token_path):
            os.remove(token_path)
        return get_sheets_service()  # 再帰的に呼び出して新しい認証を開始

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        if os.path.exists(token_path):
            os.remove(token_path)
        return get_sheets_service()  # 再帰的に呼び出して新しい認証を開始


def extract_spreadsheet_id(url):
    match = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', url)
    if match:
        return match.group(1)
    return None


def write_to_sheet(service, spreadsheet_url, data, range_name):
    spreadsheet_id = extract_spreadsheet_id(spreadsheet_url)
    if not spreadsheet_id:
        raise ValueError("Invalid spreadsheet URL")

    body = {'values': data}
    result = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range=range_name,
        valueInputOption='USER_ENTERED', body=body).execute()
    return result


def read_from_sheet(service, spreadsheet_url, range_name):
    spreadsheet_id = extract_spreadsheet_id(spreadsheet_url)
    if not spreadsheet_id:
        raise ValueError("Invalid spreadsheet URL")

    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=range_name).execute()
    return result.get('values', [])