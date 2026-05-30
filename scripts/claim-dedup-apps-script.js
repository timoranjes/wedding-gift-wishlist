// Google Apps Script — 認購去重驗證 v2
// 部署步驟：
// 1. 打開 Google Sheet → 擴充功能 → Apps Script
// 2. 貼上此腳本，儲存
// 3. 觸發條件 → 新增觸發條件 → 函數: onFormSubmit → 事件類型: 表單提交
// 4. 授權執行
//
// 功能：當新認購與已有認購的禮物相同時，標記重複的列並發送通知
// v2 改進：
//   - 自動偵測工作表名稱（從 e.source 取得實際觸發的 sheet）
//   - 從表頭自動偵測欄位位置（不假設固定順序）
//   - 僅完全匹配禮物名稱，不進行模糊匹配

const SPREADSHEET_ID = '1kqfMw6StVhgJ-c3VN7JLDCmuhOMZZ7moMZbGauNs5cA';

// 已知的禮物選項（完全匹配用）
const VALID_GIFTS = [
  'GLOSTAD 三座位梳化 (Knisa 深灰色)',
  'MALM 高身床架 (染白橡木, 150×200cm)',
  'KLEPPSTAD 趟門衣櫃 (白色)',
  'Mitsubishi MR-CGX33EY-GBK 冰箱',
  'HITACHI LTL-065SM00 洗衣機',
  'Toshiba MW3-SAC24SE 微波烤箱',
  'Carrier DC-22VSB 抽濕機',
  'Tefal 廚具',
];

function onFormSubmit(e) {
  if (!e || !e.values) return;

  const ss = e.source || SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = e.range ? e.range.getSheet() : ss.getSheets()[0];
  const sheetName = sheet.getName();

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  // 從表頭自動偵測欄位
  const headers = data[0].map(h => String(h).trim());
  const colMap = detectColumns(headers, sheetName);
  if (!colMap) {
    Logger.log(`⚠️ 無法偵測欄位，工作表: ${sheetName}，表頭: ${JSON.stringify(headers)}`);
    return;
  }

  const newRow = e.values;
  const newRowIdx = e.range ? e.range.getRow() : -1;
  if (newRowIdx < 0) return;

  const newName = String(newRow[colMap.nameCol] || '').trim();
  const newGift = String(newRow[colMap.giftCol] || '').trim();
  const newMsg = String(newRow[colMap.msgCol] || '').trim();

  if (!newName || !newGift) return;

  // 只對已知的禮物選項進行去重
  if (!isValidGift(newGift)) {
    Logger.log(`ℹ️ 非標準禮物選項，跳過去重: "${newGift}"`);
    return;
  }

  // 檢查是否已有相同禮物的認購
  for (let i = 1; i < data.length; i++) {
    const rowIdx = i + 1; // 1-based
    if (rowIdx === newRowIdx) continue; // 跳過自己

    const existingGift = String(data[i][colMap.giftCol] || '').trim();
    const existingName = String(data[i][colMap.nameCol] || '').trim();
    const existingMsg = String(data[i][colMap.msgCol] || '').trim();

    // 跳過已取消的列
    if (existingMsg.includes('已取消')) continue;

    // 完全匹配禮物名稱
    if (existingGift && existingGift === newGift) {
      // 發現重複！標記新列為已取消
      const cancelMsg = `⚠️ 已被 ${existingName} 認購，此筆已取消`;
      sheet.getRange(newRowIdx, colMap.msgCol + 1).setValue(cancelMsg);

      Logger.log(`⚠️ 認購被拒絕 — 認購人: ${newName}, 禮物: ${newGift}, 已被: ${existingName} 認購`);
      return;
    }
  }

  Logger.log(`✅ 新認購已接受：${newName} → ${newGift}`);
}

// 從表頭偵測欄位位置
function detectColumns(headers, sheetName) {
  let nameCol = -1, giftCol = -1, msgCol = -1;

  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    // 姓名欄位
    if (h === '姓名' || h === '你的名字' || h.includes('名字')) {
      nameCol = i;
    }
    // 禮物欄位（優先選擇第二個出現的選擇禮品，因為表單回覆 2 有兩個）
    if (h === '選擇禮品' || h === '你要認購的禮物' || h.includes('禮物')) {
      giftCol = i; // 會被覆蓋，最後一個匹配的是第二個選擇禮品
    }
    // 留言欄位
    if (h === '留言' || h === '留言／祝福（可選）' || h === '留言／祝福') {
      msgCol = i;
    }
  }

  if (nameCol < 0 || giftCol < 0 || msgCol < 0) {
    Logger.log(`⚠️ 欄位偵測失敗 — 工作表: ${sheetName}, 表頭: ${JSON.stringify(headers)}`);
    return null;
  }

  return { nameCol, giftCol, msgCol };
}

// 檢查是否為有效的禮物選項
function isValidGift(gift) {
  return VALID_GIFTS.some(v => v === gift);
}
