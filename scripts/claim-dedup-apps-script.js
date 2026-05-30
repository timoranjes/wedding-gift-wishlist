// Google Apps Script — 認購去重驗證
// 部署步驟：
// 1. 打開 Google Sheet → 擴充功能 → Apps Script
// 2. 貼上此腳本，儲存
// 3. 觸發條件 → 新增觸發條件 → 函數: onFormSubmit → 事件類型: 表單提交
// 4. 授權執行
//
// 功能：當新認購與已有認購的禮物相同時，自動刪除重複的列並發送通知

const SPREADSHEET_ID = '1kqfMw6StVhgJ-c3VN7JLDCmuhOMZZ7moMZbGauNs5cA';
const SHEET_NAME = '表單回應 1';  // 請確認實際的工作表名稱

function onFormSubmit(e) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  const newRow = e.values;  // [timestamp, gift, name, message]
  if (!newRow || newRow.length < 3) return;
  
  const newGift = newRow[1].trim();
  const newName = newRow[2].trim();
  const newRowIdx = e.range ? e.range.getRow() : -1;
  
  // 檢查是否已有相同禮物的認購
  for (let i = 1; i < data.length; i++) {
    const existingGift = String(data[i][1] || '').trim();
    const existingName = String(data[i][2] || '').trim();
    
    // 跳過自己（測試用）
    if (i + 1 === newRowIdx) continue;
    
    // 匹配邏輯：完全匹配或前綴匹配（處理 "沙發" vs "沙發 — IKEA KIVIK"）
    if (existingGift && newGift && (
      existingGift === newGift ||
      existingGift.startsWith(newGift + ' — ') ||
      existingGift.startsWith(newGift + ' - ') ||
      newGift.startsWith(existingGift + ' — ') ||
      newGift.startsWith(existingGift + ' - ')
    )) {
      // 發現重複！刪除新列
      if (newRowIdx > 1) {
        sheet.deleteRow(newRowIdx);
      }
      
      // 發送通知（可選）
      const msg = `⚠️ 認購已被拒絕\n\n` +
        `認購人：${newName}\n` +
        `禮物：${newGift}\n` +
        `原因：已被 ${existingName} 認購\n\n` +
        `每個禮物只能被認購一次。`;
      
      // 可選：發送 Email 通知
      // MailApp.sendEmail('your-email@example.com', '認購被拒絕', msg);
      
      Logger.log(msg);
      return;
    }
  }
  
  Logger.log(`✅ 新認購已接受：${newName} → ${newGift}`);
}
