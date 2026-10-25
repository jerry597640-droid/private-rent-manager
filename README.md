# 每月收租管理系統

GitHub Pages 僅提供公開的**安全入口**。實際收租資料必須在受限存取的 Google Apps Script 網頁內處理，不得在公開 GitHub 前端放入租客資料、試算表 ID、Google token 或部署金鑰。

## 架構
- GitHub Pages: `index.html`（手機響應式入口）
- Google Apps Script: 私人 Web App（`apps-script/Code.gs` 為程式範本，需手動貼入 Apps Script）
- Google Sheets: 每月分頁，依標題定位欄位

## Google Apps Script 設定
1. 在 Apps Script 專案中貼入 `apps-script/Code.gs`。
2. 專案設定 → 指令碼屬性，新增 `SPREADSHEET_ID`，值為你自己試算表的 ID；**不要放入 GitHub**。
3. 新增 `Index.html` 並製作表單，透過 `google.script.run.getMonths/getRows/saveRent` 呼叫後端。
4. 部署為網頁應用程式，執行身分為自己，存取權限設為「只有自己」。
5. 先在測試副本驗證各月份、日期、金額與勾選，再正式使用。

## 注意
- 目前 GitHub 入口與 Apps Script 尚未自動連接；此範本不會自行修改試算表。
- 不同月份可能存在欄位名稱差異，程式遇到不符欄位會停止寫入，避免誤填。
- 參考開源專案的架構，不直接複製未審核程式。
