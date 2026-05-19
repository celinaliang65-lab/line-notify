#!/usr/bin/env python3
"""
fetch_groups.py — 執行一次即可
讀取「班群設定」工作表中的 LINE Group ID，
呼叫 LINE API 查詢群組名稱，自動填回 Google Sheets。
"""

import os
import sys
import json
import requests

from google.oauth2 import service_account
from googleapiclient.discovery import build

SPREADSHEET_ID = "15shqxW2MrEybiEjLqLbtHMymS2R2SpD9CoNN3N6i0ww"
SHEET_NAME     = "班群設定"
LINE_API       = "https://api.line.me/v2/bot/group/{group_id}/summary"
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]

COL_NAME = "班群名稱"
COL_ID   = "LINE Group ID"


def get_gsheet_service():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS", "")
    if not creds_json:
        print("❌ 未設定環境變數 GOOGLE_CREDENTIALS")
        sys.exit(1)
    creds_dict = json.loads(creds_json)
    creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return build("sheets", "v4", credentials=creds)


def get_token() -> str:
    token = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
    if not token:
        if len(sys.argv) > 1:
            return sys.argv[1]
        print("❌ 請設定環境變數 LINE_CHANNEL_ACCESS_TOKEN")
        sys.exit(1)
    return token


def col_letter(idx: int) -> str:
    """0-based 欄 index 轉欄位字母"""
    result = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        result = chr(65 + rem) + result
    return result


def fetch_group_name(token: str, group_id: str) -> str | None:
    url = LINE_API.format(group_id=group_id.strip())
    headers = {"Authorization": f"Bearer {token}"}
    try:
        resp = requests.get(url, headers=headers, timeout=10)
        if resp.status_code == 200:
            return resp.json().get("groupName", "")
        elif resp.status_code == 403:
            print(f"   ⚠️  Bot 尚未加入群組 {group_id}，請先將 Bot 加為群組成員")
        elif resp.status_code == 404:
            print(f"   ⚠️  找不到群組 {group_id}，請確認 Group ID 正確")
        else:
            print(f"   ⚠️  API 錯誤 {resp.status_code}: {resp.text}")
        return None
    except Exception as e:
        print(f"   ⚠️  連線失敗: {e}")
        return None


def run():
    token   = get_token()
    service = get_gsheet_service()

    # 讀取班群設定工作表
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=SHEET_NAME
    ).execute()
    all_rows = result.get("values", [])

    if not all_rows:
        print("❌ 工作表為空")
        sys.exit(1)

    headers = all_rows[0]

    try:
        name_col_idx = headers.index(COL_NAME)
        id_col_idx   = headers.index(COL_ID)
    except ValueError as e:
        print(f"❌ 找不到欄位：{e}")
        sys.exit(1)

    print(f"🔍 開始查詢群組名稱...\n")

    updated = 0
    skipped = 0

    # 資料從第2列起（index 1），對應 Google Sheets 第2列
    for row_idx, row in enumerate(all_rows[1:], start=2):
        # 補齊欄位
        while len(row) <= max(name_col_idx, id_col_idx):
            row.append("")

        group_id = row[id_col_idx].strip()
        existing = row[name_col_idx].strip()

        if not group_id:
            continue

        if existing and existing not in ("", COL_NAME):
            print(f"   ⏭  {group_id} → 已有名稱「{existing}」，跳過")
            skipped += 1
            continue

        print(f"   🔎 查詢 {group_id} ...", end=" ")
        name = fetch_group_name(token, group_id)

        cell_range = f"{SHEET_NAME}!{col_letter(name_col_idx)}{row_idx}"
        value      = name if name else "（查詢失敗）"

        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=cell_range,
            valueInputOption="RAW",
            body={"values": [[value]]}
        ).execute()

        if name:
            print(f"✅ {name}")
            updated += 1
        else:
            print("❌ 查詢失敗")

    print(f"\n🎉 完成！更新 {updated} 筆，略過 {skipped} 筆")


if __name__ == "__main__":
    run()
