#!/usr/bin/env python3
"""
LINE 班群自動通知系統
每天早上 06:30 執行（台灣時間），依 Google Sheets 日期推播對應班群
"""

import os
import sys
import json
import requests
import pandas as pd
from datetime import datetime, date
import zoneinfo

from google.oauth2 import service_account
from googleapiclient.discovery import build

SPREADSHEET_ID = "15shqxW2MrEybiEjLqLbtHMymS2R2SpD9CoNN3N6i0ww"
SHEET_NAME     = "提醒排程"
LINE_API_URL   = "https://api.line.me/v2/bot/message/push"
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]

COL_TYPE    = "提醒類別"
COL_NAME    = "班群名稱"
COL_GROUP   = "目標班群 ID"
COL_DATE    = "提醒日期"
COL_TITLE   = "訊息標題"
COL_CONTENT = "提醒內容細節"
COL_SENT    = "已發送"
COL_SENT_AT = "發送時間"

# 提醒類別 → badge 顏色
CATEGORY_STYLE = {
    "作業繳交": {"bg": "#E3F2FD", "text": "#1565C0"},
    "上台規範": {"bg": "#FFF3E0", "text": "#E65100"},
    "公告":     {"bg": "#E8F5E9", "text": "#1B5E20"},
}
DEFAULT_STYLE = {"bg": "#F5F5F5", "text": "#555555"}


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
        print("❌ 未設定環境變數 LINE_CHANNEL_ACCESS_TOKEN")
        sys.exit(1)
    return token


def read_sheet(service) -> tuple[list, list]:
    """讀取工作表，回傳 (headers, data_rows)"""
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=SHEET_NAME
    ).execute()
    all_rows = result.get("values", [])
    if len(all_rows) < 3:
        return [], []
    headers   = all_rows[0]   # 第1列：欄位名稱
    data_rows = all_rows[2:]  # 第2列說明列跳過，第3列起為資料
    return headers, data_rows


def rows_to_df(headers, data_rows) -> pd.DataFrame:
    """將 rows 補齊長度後轉成 DataFrame"""
    max_col = len(headers)
    padded  = [row + [""] * (max_col - len(row)) for row in data_rows]
    return pd.DataFrame(padded, columns=headers)


def parse_date(val: str):
    """解析日期字串，支援 YYYY/M/D、YYYY-M-D 格式"""
    if not val or str(val).strip() == "":
        return None
    try:
        return pd.to_datetime(val, errors="coerce").date()
    except Exception:
        return None


def col_letter(idx: int) -> str:
    """0-based 欄 index 轉欄位字母（A, B, ... Z, AA ...）"""
    result = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        result = chr(65 + rem) + result
    return result


def write_sent(service, sheet_row: int, sent_col_idx: int, sent_at_col_idx: int, now_str: str):
    """回寫已發送狀態到 Google Sheets（發送成功才呼叫）"""
    sent_range    = f"{SHEET_NAME}!{col_letter(sent_col_idx)}{sheet_row}"
    sent_at_range = f"{SHEET_NAME}!{col_letter(sent_at_col_idx)}{sheet_row}"
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={
            "valueInputOption": "RAW",
            "data": [
                {"range": sent_range,    "values": [["TRUE"]]},
                {"range": sent_at_range, "values": [[now_str]]},
            ]
        }
    ).execute()


def send_line_message(token: str, group_id: str, category: str,
                      title: str, content: str, notify_date: str) -> bool:
    # 驗證 Group ID 格式（C 開頭、33碼）
    if not group_id.startswith("C") or len(group_id) != 33:
        print(f"   ❌ Group ID 格式錯誤：[{group_id}]（應為 C 開頭共 33 碼）")
        return False

    style    = CATEGORY_STYLE.get(category, DEFAULT_STYLE)
    alt_text = f"【{category}】{title}｜{content}"

    payload = {
        "to": group_id,
        "messages": [
            {
                "type": "flex",
                "altText": alt_text,
                "contents": {
                    "type": "bubble",
                    "body": {
                        "type": "box",
                        "layout": "vertical",
                        "paddingAll": "16px",
                        "spacing": "sm",
                        "contents": [
                            {
                                "type": "box",
                                "layout": "horizontal",
                                "alignItems": "center",
                                "contents": [
                                    {
                                        "type": "box",
                                        "layout": "vertical",
                                        "backgroundColor": style["bg"],
                                        "paddingAll": "4px",
                                        "cornerRadius": "4px",
                                        "flex": 0,
                                        "contents": [
                                            {
                                                "type": "text",
                                                "text": category if category else "通知",
                                                "size": "xs",
                                                "weight": "bold",
                                                "color": style["text"]
                                            }
                                        ]
                                    },
                                    {
                                        "type": "text",
                                        "text": title if title else "班群通知",
                                        "size": "sm",
                                        "weight": "bold",
                                        "color": "#111111",
                                        "flex": 1,
                                        "margin": "sm",
                                        "wrap": True
                                    },
                                    {
                                        "type": "text",
                                        "text": notify_date,
                                        "size": "xs",
                                        "color": "#999999",
                                        "align": "end",
                                        "flex": 0
                                    }
                                ]
                            },
                            {
                                "type": "separator",
                                "margin": "sm",
                                "color": "#EEEEEE"
                            },
                            {
                                "type": "text",
                                "text": content,
                                "size": "sm",
                                "color": "#333333",
                                "wrap": True,
                                "margin": "xs"
                            }
                        ]
                    }
                }
            }
        ]
    }

    try:
        resp = requests.post(
            LINE_API_URL,
            headers={"Content-Type": "application/json",
                     "Authorization": f"Bearer {token}"},
            json=payload,
            timeout=10
        )
        if resp.status_code == 200:
            return True
        print(f"   ⚠️  LINE API 錯誤 {resp.status_code}: {resp.text}")
        return False
    except Exception as e:
        print(f"   ⚠️  連線失敗: {e}")
        return False


def run():
    token   = get_token()
    service = get_gsheet_service()
    today   = datetime.now(zoneinfo.ZoneInfo("Asia/Taipei")).date()
    now_str = datetime.now(zoneinfo.ZoneInfo("Asia/Taipei")).strftime("%Y/%m/%d %H:%M")

    print(f"📅 執行日期：{today}")

    headers, data_rows = read_sheet(service)
    if not headers:
        print("❌ 無法讀取工作表或資料為空")
        sys.exit(1)

    df = rows_to_df(headers, data_rows)
    print(f"📋 共 {len(df)} 筆資料（不含說明列）")

    # 取得已發送、發送時間的欄位 index（0-based）
    try:
        sent_col_idx    = headers.index(COL_SENT)
        sent_at_col_idx = headers.index(COL_SENT_AT)
    except ValueError as e:
        print(f"❌ 找不到欄位：{e}")
        sys.exit(1)

    # 篩選今天且未發送的資料
    targets = []
    for i, row in df.iterrows():
        raw_date = str(row.get(COL_DATE, "")).strip()
        sent     = str(row.get(COL_SENT, "")).strip()
        d        = parse_date(raw_date)
        is_empty = sent == "" or sent.lower() == "nan"

        print(f"   [{i}] 日期={d} | 已發送={repr(sent)} | 日期符合={d==today} | 發送為空={is_empty}")

        if d == today and is_empty:
            targets.append((i, row))

    if not targets:
        print("✅ 今天沒有待發送的提醒。")
        return

    print(f"📨 找到 {len(targets)} 筆待發送\n")

    success = 0
    for df_idx, row in targets:
        group_id    = str(row.get(COL_GROUP, "")).strip()
        category    = str(row.get(COL_TYPE, "")).strip()
        title       = str(row.get(COL_TITLE, "")).strip()
        content     = str(row.get(COL_CONTENT, "")).strip()
        name        = str(row.get(COL_NAME, "")).strip()
        notify_date = str(row.get(COL_DATE, "")).strip()

        if not group_id or not content:
            print(f"   ⚠️  索引 {df_idx} 缺少群組 ID 或內容，跳過")
            continue

        print(f"   → 【{name}】{group_id}")
        print(f"      類別：{category}｜標題：{title}")
        print(f"      內容：{content}")

        ok = send_line_message(token, group_id, category, title, content, notify_date)
        if ok:
            # Google Sheets 列號：標題列1 + 說明列2 + 資料從第3列起
            sheet_row = 3 + df_idx
            write_sent(service, sheet_row, sent_col_idx, sent_at_col_idx, now_str)
            print("      ✅ 成功")
            success += 1
        else:
            print("      ❌ 失敗")

    print(f"\n🎉 完成！{success}/{len(targets)} 筆發送成功")


if __name__ == "__main__":
    run()
