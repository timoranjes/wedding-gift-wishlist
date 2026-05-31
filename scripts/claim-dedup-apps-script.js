// Google Apps Script — 認購去重驗證 v3
// 部署步驟：
// 1. 打開 Google Sheet → 擴充功能 → Apps Script
// 2. 貼上此腳本，儲存
// 3. 觸發條件 → 新增觸發條件 → 函數: onFormSubmit → 事件類型: 表單提交
// 4. 授權執行
//
// v3 改進：
//   - 使用固定欄位映射（與網頁 SHEET_TABS 一致），避免 auto-detect 因重複表頭錯位
//   - e.values 欄位不足時，自動 fallback 到 sheet data 讀取
//   - 保留 auto-detect 作為未知工作表的 fallback

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

// 固定欄位映射（0-based index，與網頁 index.html 的 SHEET_TABS 一致）
const FIXED_COLUMN_MAP = {
  '表單回覆 2':     { nameCol: 1, giftCol: 4, msgCol: 3 },
  '表單回覆 1':     { nameCol: 2, giftCol: 1, msgCol: 3 },
  'Form Responses 1': { nameCol: 2, giftCol: 1, msgCol: 3 },
};

function onFormSubmit(e) {
  if (!e || !e.values) return;

  const ss = e.source || SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = e.range ? e.range.getSheet() : ss.getSheets()[0];
  const sheetName = sheet.getName();

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const headers = data[0].map(h => String(h).trim());
  const colMap = detectColumns(headers, sheetName);
  if (!colMap) {
    Logger.log(`⚠️ 無法偵測欄位，工作表: ${sheetName}，表頭: ${JSON.stringify(headers)}`);
    return;
  }

  const newRowIdx = e.range ? e.range.getRow() : -1;
  if (newRowIdx < 0) return;

  // 從 e.values 讀取新列資料；若欄位不足則 fallback 到 sheet data
  let newRow = e.values;
  const maxCol = Math.max(colMap.nameCol, colMap.giftCol, colMap.msgCol);
  if (newRow.length <= maxCol) {
    Logger.log(`⚠️ e.values 欄位不足（${newRow.length} < ${maxCol + 1}），改用 sheet data — 工作表: ${sheetName}`);
    if (newRowIdx >= 1 && newRowIdx <= data.length) {
      newRow = data[newRowIdx - 1];
    } else {
      Logger.log(`⚠️ newRowIdx ${newRowIdx} 超出範圍，放棄處理`);
      return;
    }
  }

  const newName = String(newRow[colMap.nameCol] || '').trim();
  const newGift = String(newRow[colMap.giftCol] || '').trim();
  const newMsg = String(newRow[colMap.msgCol] || '').trim();

  Logger.log(`📋 新認購 — name: "${newName}", gift: "${newGift}", msg: "${newMsg}", row: ${newRowIdx}, sheet: ${sheetName}`);

  if (!newName || !newGift) {
    Logger.log(`⚠️ 姓名或禮物為空，跳過`);
    return;
  }

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
    if (existingMsg.includes('已取消') || existingMsg.includes('CANCELLED')) continue;

    // 完全匹配禮物名稱
    if (existingGift && existingGift === newGift) {
      const cancelMsg = `⚠️ 已被 ${existingName} 認購，此筆已取消`;
      sheet.getRange(newRowIdx, colMap.msgCol + 1).setValue(cancelMsg);

      Logger.log(`⚠️ 認購被拒絕 — 認購人: ${newName}, 禮物: ${newGift}, 已被: ${existingName} 認購`);
      return;
    }
  }

  // 沒找到重複 → 標記為已接受
  sheet.getRange(newRowIdx, colMap.msgCol + 1).setValue('✅');
  Logger.log(`✅ 新認購已接受：${newName} → ${newGift}`);
}

// 從表頭偵測欄位位置（fallback，僅在固定映射找不到時使用）
function detectColumns(headers, sheetName) {
  // 優先使用固定映射
  if (FIXED_COLUMN_MAP[sheetName]) {
    return FIXED_COLUMN_MAP[sheetName];
  }

  let nameCol = -1, giftCol = -1, msgCol = -1;

  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    if (h === '姓名' || h === '你的名字' || h.includes('名字')) {
      nameCol = i;
    }
    // 遇到重複表頭時，只取 FIRST 個匹配
    if (giftCol < 0 && (h === '選擇禮品' || h === '你要認購的禮物' || h.includes('禮物'))) {
      giftCol = i;
    }
    if (h === '留言' || h === '留言／祝福（可選）' || h === '留言／祝福') {
      msgCol = i;
    }
  }

  if (nameCol < 0 || giftCol < 0 || msgCol < 0) {
    Logger.log(`⚠️ 欄位偵測失敗 — 工作表: ${sheetName}, 表頭: ${JSON.stringify(headers)}`);
    return null;
  }

  Logger.log(`⚠️ 使用 auto-detect fallback — ${sheetName}: nameCol=${nameCol}, giftCol=${giftCol}, msgCol=${msgCol}`);
  return { nameCol, giftCol, msgCol };
}

// 檢查是否為有效的禮物選項
function isValidGift(gift) {
  return VALID_GIFTS.some(v => v === gift);
}
