# LINE 班群自動提醒系統

每天早上 **06:30（台灣時間）** 自動讀取 Excel，推播當天提醒到對應 LINE 群組。

---

## 📁 專案結構

```
line-notify-system/
├── .github/
│   └── workflows/
│       └── line-notify.yml    ← GitHub Actions 排程設定
├── data/
│   └── LINE班群提醒排程.xlsx   ← 提醒資料（日常只需維護這個）
├── notify.py                  ← 主程式（每天自動執行）
├── fetch_groups.py            ← 一次性工具：自動抓取群組名稱填入 Excel
├── requirements.txt           ← Python 套件清單
└── README.md
```

---

## 🚀 設定步驟

### 1. 取得 LINE Channel Access Token

1. 前往 [LINE Developers Console](https://developers.line.biz/)
2. 建立 **Messaging API Channel**
3. 進入 Channel 頁面 → **Messaging API** 分頁
4. 複製 **Channel Access Token（長效）**

---

### 2. 將 Bot 加入班群並取得 Group ID

**加入班群：**
1. 在 LINE Developers → Messaging API 分頁，掃描 Bot 的 QR Code 加為好友
2. 在 LINE App 中，將這個 Bot 邀請加入每個班群

**取得 Group ID：**

Group ID 無法從 LINE App 直接看到，需要透過 Webhook 取得：

1. 前往 [https://webhook.site](https://webhook.site)
   取得專屬網址（例如 `https://webhook.site/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`）
2. 到 LINE Developers → Messaging API → **Webhook URL** 填入該網址
3. 開啟「Use webhook」（不需要點 Verify，出現錯誤可忽略）
4. 在班群內發任意一則訊息
5. 回到 webhook.site，點左側收到的請求，找到以下欄位：
   ```json
   "source": {
     "type": "group",
     "groupId": "C1234567890abcdef..."
   }
   ```
6. 複製 `groupId` 的值，貼入 Excel「班群設定」工作表的「LINE Group ID」欄位

---

### 3. 設定 GitHub Secret

1. 進入 GitHub repo → **Settings → Secrets and variables → Actions**
2. 點選 **New repository secret**
3. 名稱：`LINE_CHANNEL_ACCESS_TOKEN`，值：貼上第 1 步取得的 Token

---

### 4. 自動抓取群組名稱（只需執行一次）

在「班群設定」工作表填入所有 LINE Group ID 後，在本機執行：

```bash
# 方式一：設定環境變數
export LINE_CHANNEL_ACCESS_TOKEN=你的Token
python fetch_groups.py

# 方式二：直接帶入 Token
python fetch_groups.py 你的Token
```

程式會自動呼叫 LINE API，將群組名稱填回 Excel「班群名稱」欄位。  
**已有名稱的列會跳過，不會覆蓋。**

> ⚠️ 執行前請確認 Bot 已被邀請加入每個群組，否則會顯示查詢失敗。

執行完後將 Excel commit & push 回 GitHub：

```bash
git add data/LINE班群提醒排程.xlsx
git commit -m "更新班群名稱"
git push
```

---

### 5. 填寫提醒排程

開啟 `data/LINE班群提醒排程.xlsx` 的「提醒排程」工作表，填入提醒資料：

| 欄位 | 說明 | 誰填 |
|------|------|------|
| 提醒類別 | 作業繳交、上台規範、公告（影響 badge 顏色） | 你 |
| 班群名稱 | 對照「班群設定」工作表填入，純顯示用 | 你 |
| 目標班群 ID | LINE Group ID（每列可填不同班群） | 你 |
| 提醒日期 | 格式 YYYY/M/D，當天 06:30 自動發送 | 你 |
| 訊息標題 | LINE 推播卡片的標題 | 你 |
| 提醒內容細節 | LINE 推播卡片的完整內容 | 你 |
| 備註 | 補充說明，不會被發送 | 你 |
| 已發送 | ⚠️ 請勿手動修改，程式自動填入 TRUE | 程式 |
| 發送時間 | ⚠️ 請勿手動修改，程式自動填入時間戳 | 程式 |

> 💡 不同提醒可以填不同的「目標班群 ID」，程式會各自發送到對應群組。

---

### 6. Push 到 GitHub，等待自動執行

```bash
git add .
git commit -m "新增提醒排程"
git push
```

GitHub Actions 會在每天 **06:30（台灣時間）** 自動執行。

---

## 🔁 運作邏輯

```
日期 < 今天               → 跳過（過期不處理）
日期 = 今天，已發送為空   → 發送 LINE，自動填入已發送 TRUE + 發送時間
日期 = 今天，已發送為TRUE → 跳過（今天已發過，防止重複）
日期 > 今天               → 跳過（尚未到期）
```

---

## 📱 通知卡片樣式

LINE 推播為卡片格式，包含以下資訊：

| 區塊 | 內容 |
|------|------|
| 頂部左側 | 提醒類別 badge（顏色依類別自動對應） |
| 頂部右側 | 訊息標題 |
| 分隔線後 | 提醒日期（灰色小字） |
| 底部 | 提醒內容細節 |

**類別 badge 顏色對照：**

| 提醒類別 | 顏色 |
|---------|------|
| 作業繳交 | 藍色 |
| 上台規範 | 橘色 |
| 公告 | 綠色 |
| 其他（自訂） | 灰色 |

> 💡 需要新增類別顏色，請修改 `notify.py` 中的 `CATEGORY_STYLE` 字典。

---

## 🧪 手動測試

在 GitHub repo 頁面：**Actions → LINE 班群每日提醒通知 → Run workflow**

---

## ⚠️ 注意事項

- Bot 加入群組為**一般成員**即可，不需要管理員權限
- LINE Messaging API 免費版每月有 **200 則**免費推播額度，超過需付費
- Excel 更新後需 **commit & push** 到 GitHub 才會生效
- GitHub Actions 使用 UTC 時間，設定為 `22:30 UTC` = 台灣時間 `06:30`
- Webhook.site 的 Verify 失敗可忽略，不影響取得 Group ID
