"""
google_sheets.py - Google スプレッドシート連携モジュール

認証方法：
  A) サービスアカウント（JSON キーファイル）  ← 会社・チーム向け
  B) OAuth 2.0（個人 Google アカウント）      ← 個人向け

依存パッケージ:
    pip install google-auth google-auth-oauthlib google-api-python-client
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import json
from pathlib import Path
from typing import Optional


# 認証情報の保存先
CREDS_DIR        = Path(__file__).parent / "credentials"
SERVICE_KEY_PATH = CREDS_DIR / "service_account.json"
OAUTH_CLIENT_PATH= CREDS_DIR / "oauth_client.json"
OAUTH_TOKEN_PATH = CREDS_DIR / "oauth_token.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]


# ── パッケージ確認 ─────────────────────────────────────────────
def check_packages() -> bool:
    try:
        import google.auth          # noqa
        import googleapiclient      # noqa
        return True
    except ImportError:
        return False


def install_hint() -> str:
    return (
        "Google Sheets 連携に必要なパッケージがありません。\n\n"
        "以下を PowerShell で実行してください:\n\n"
        "  pip install google-auth google-auth-oauthlib google-api-python-client\n\n"
        "インストール後にアプリを再起動してください。"
    )


# ── 認証 ──────────────────────────────────────────────────────
class GoogleSheetsClient:
    """
    Google Sheets API クライアント。
    サービスアカウントまたは OAuth どちらでも初期化できる。
    """

    def __init__(self, creds=None):
        if not check_packages():
            raise ImportError(install_hint())
        self.creds   = creds
        self._sheets = None
        self._drive  = None

    @classmethod
    def from_service_account(cls, key_path: str = None) -> "GoogleSheetsClient":
        """サービスアカウント JSON キーから認証"""
        from google.oauth2 import service_account

        path = Path(key_path) if key_path else SERVICE_KEY_PATH
        if not path.exists():
            raise FileNotFoundError(
                f"サービスアカウントキーが見つかりません:\n{path}\n\n"
                "Google Cloud Console からダウンロードして\n"
                f"credentials/ フォルダに「service_account.json」として保存してください。\n\n"
                "詳しくは SETUP_GUIDE.md を参照してください。")

        creds = service_account.Credentials.from_service_account_file(
            str(path), scopes=SCOPES)
        return cls(creds)

    @classmethod
    def from_oauth(cls, client_secret_path: str = None) -> "GoogleSheetsClient":
        """OAuth 2.0 でブラウザログイン認証"""
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.auth.transport.requests import Request

        client_path = Path(client_secret_path) if client_secret_path else OAUTH_CLIENT_PATH
        token_path  = OAUTH_TOKEN_PATH

        creds = None

        # 保存済みトークンがあれば再利用
        if token_path.exists():
            creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)

        # トークンが無効 or 期限切れなら再認証
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                from google.auth.transport.requests import Request
                creds.refresh(Request())
            else:
                if not client_path.exists():
                    raise FileNotFoundError(
                        f"OAuth クライアントシークレットが見つかりません:\n{client_path}\n\n"
                        "Google Cloud Console からダウンロードして\n"
                        "credentials/ フォルダに「oauth_client.json」として保存してください。\n\n"
                        "詳しくは SETUP_GUIDE.md を参照してください。")
                flow = InstalledAppFlow.from_client_secrets_file(
                    str(client_path), SCOPES)
                creds = flow.run_local_server(port=0)

            # トークンを保存（次回はブラウザ不要）
            CREDS_DIR.mkdir(exist_ok=True)
            token_path.write_text(creds.to_json(), encoding="utf-8")

        return cls(creds)

    # ── API クライアント ────────────────────────────────────────
    def _get_sheets(self):
        if self._sheets is None:
            from googleapiclient.discovery import build
            self._sheets = build("sheets", "v4", credentials=self.creds)
        return self._sheets

    def _get_drive(self):
        if self._drive is None:
            from googleapiclient.discovery import build
            self._drive = build("drive", "v3", credentials=self.creds)
        return self._drive

    # ── スプレッドシート一覧 ────────────────────────────────────
    def list_spreadsheets(self) -> list[dict]:
        """Drive から スプレッドシート一覧を取得"""
        service = self._get_drive()
        result  = service.files().list(
            q="mimeType='application/vnd.google-apps.spreadsheet'",
            fields="files(id, name, modifiedTime)",
            orderBy="modifiedTime desc",
            pageSize=50
        ).execute()
        return result.get("files", [])

    # ── シート一覧 ──────────────────────────────────────────────
    def list_sheets(self, spreadsheet_id: str) -> list[dict]:
        """スプレッドシート内のシート一覧"""
        service = self._get_sheets()
        meta = service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            fields="sheets.properties"
        ).execute()
        return [
            {"title": s["properties"]["title"],
             "index": s["properties"]["index"],
             "sheetId": s["properties"]["sheetId"]}
            for s in meta.get("sheets", [])
        ]

    # ── データ取得 ──────────────────────────────────────────────
    def get_sheet_data(self, spreadsheet_id: str,
                       sheet_name: str = None,
                       range_str: str = None) -> dict:
        """
        シートのデータを取得して { columns, rows } 形式で返す
        range_str 例: "Sheet1!A1:Z1000"
        """
        service = self._get_sheets()

        if range_str is None:
            if sheet_name:
                range_str = f"'{sheet_name}'"
            else:
                # 最初のシートを取得
                sheets = self.list_sheets(spreadsheet_id)
                if not sheets:
                    return {"columns": [], "rows": []}
                range_str = f"'{sheets[0]['title']}'"

        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_str,
            valueRenderOption="FORMATTED_VALUE"
        ).execute()

        values = result.get("values", [])
        if not values:
            return {"columns": [], "rows": []}

        # 1行目をヘッダーとして扱う
        headers  = [str(h) if h != "" else f"col_{i}"
                    for i, h in enumerate(values[0])]
        max_cols = len(headers)
        data_rows = []
        for row in values[1:]:
            # 列数が足りない行はNoneで埋める
            padded = row + [""] * (max_cols - len(row))
            data_rows.append([str(c) for c in padded[:max_cols]])

        return {"columns": headers, "rows": data_rows,
                "sheet": sheet_name or range_str,
                "spreadsheet_id": spreadsheet_id}

    # ── スプレッドシートIDをURLから抽出 ─────────────────────────
    @staticmethod
    def extract_id(url_or_id: str) -> str:
        """
        URL または ID 文字列からスプレッドシートIDを抽出
        例: https://docs.google.com/spreadsheets/d/XXXXXX/edit → XXXXXX
        """
        if "/" not in url_or_id:
            return url_or_id  # すでにIDのみ
        import re
        m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url_or_id)
        if m:
            return m.group(1)
        raise ValueError(f"スプレッドシートURLの形式が正しくありません:\n{url_or_id}")
