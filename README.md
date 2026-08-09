# 互動座位圖

辦公室 39 席座位圖，可即時標記狀態、填寫姓名與備註，全員共用同一份資料。

**線上使用**：https://shunda1788-beep.github.io/seatmap/

## 功能

- 正／反兩個視角，資料共用（反視角為正視角旋轉 180°，由程式推導）
- 筆刷模式：選一個顏色後可連續點選座位上色，同格再點一次取消
- 顏色說明可自訂命名，並在上色時自動帶入座位標籤
- 座位可填姓名與備註，改動自動儲存
- 多人即時同步，以座位為單位處理衝突
- CSV 匯出、JSON 備份還原、列印（附備註一覽）
- `?admin=1` 才顯示管理功能，避免誤觸

## 架構

| 層 | 使用 |
|---|---|
| 前端 | 原生 HTML / CSS / JavaScript，無框架，單一檔案 |
| 後端 | Google Apps Script（Web App） |
| 資料 | Google 試算表 |
| 部署 | GitHub Pages |

改動後 0.7 秒寫回試算表，每 25 秒輪詢更新；分頁隱藏時暫停輪詢以節省 Apps Script 免費額度。

## 檔案

- `index.html` — 網頁本體，開頭 `BUILTIN_API` 設定試算表網址
- `Code.gs` — Apps Script 後端程式碼
- `setup-guide.md` — 建置與部署說明

## 設定

見 [setup-guide.md](setup-guide.md)。
