// 🔧 詮栩 CNC 雲端系統 - Google Apps Script 完整整合版
// 功能涵蓋：刀具入庫、查詢、加工登記與查詢，支援 RESTful API 結構與 JSON 資料格式

/**
 * 🌐 主函式：處理 GET 請求
 * 說明：根據 URL 參數中的 page 決定執行哪一個功能模組，並回傳對應資料。
 * 支援功能：刀具入庫、刀具查詢、加工登記頁面下拉資料、加工數量查詢
 */
function doGet(e) {
  const page = e.parameter.page;
  const keyword = e.parameter.keyword;

  if (page === 'inventory') return getDropdownData(e);
  if (page === 'load_tool_fields' && e.parameter.category) return getToolFieldLists(e.parameter.category);
  if (page === 'check') return handleToolCheck(keyword, e.parameter);
  
  if (page === 'production_entry') {
    if (e.parameter.type === 'last_submitted') return getLastSubmittedProduction();
    return handleProductionEntryLoad(e.parameter);
  }

  if (page === 'production_query') {
    const type = e.parameter.type;
    if (e.parameter.mode === 'latest' && e.parameter.pd) {
      return getLatestProductionByProduct(e.parameter.pd);           // ← 只帶 pd
    }
    if (type === 'export_data') return exportProductionData(e.parameter.pd); // ← 只帶 pd
    if (type === 'dropdown')    return getDropdownValues();          // 仍回傳 products
    return getProductionQuery(e.parameter);                          // 區間查詢（只用 pd）
  }

  return ContentService.createTextOutput(JSON.stringify({ error: "無效的 page 參數" }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 📬 主函式：處理 POST 請求
 * 說明：根據前端送來的資料內容，自動分派功能給「刀具入庫」、「加工登記」、「刪除最後一筆加工資料」
 * 使用 JSON 格式傳遞資料，前端用 fetch 傳送 body 即可
 */
function doPost(e) {
  const data = JSON.parse(e.postData.contents); // 將前端送來的 JSON 字串解析成物件
  // ✅ 判斷為「刀具入庫」功能（需有 category 與 spec 欄位）
  if (data.category && data.spec) {
    return handleToolInventoryPost(data);  // 執行刀具入庫處理邏輯（寫入入庫表、更新庫存）
  }
  // ✅ 判斷為「加工數量登記」功能（需有機台與產品編號）
  else if (data.mc && data.pd) {
    return addProductionData(data); // 寫入加工數據，含檢查是否重複、加上時間戳
  }
  // ✅ 判斷為「刪除最後一筆加工紀錄」功能（需明確指定 action 欄位）
  else if (data.action === 'deleteLastProduction') {
    return deleteLastProductionRow(); // 找出最後一筆資料並刪除
  }
  // ⚠️ 如果傳入格式不符合任何一種，回傳失敗（前端可接收 { success: false } 處理錯誤提示）
  return ContentService.createTextOutput(JSON.stringify({ success: false }))
    .setMimeType(ContentService.MimeType.JSON); // 回傳格式為 JSON
}


/**
 * 🏭 handleProductionEntryLoad：載入加工數量登記頁所需的下拉選單資料
 * 說明：根據傳入的 type 參數，回傳對應的下拉選單資料（機台、產品、訂單）
 * 使用方式：前端以 ?page=production_entry&type=machine / product / order 呼叫
 * 
 * @param {Object} param - URL 傳入的參數物件
 *   - type: 指定要取得哪一種下拉資料（machine / product / order）
 *   - pd: 若 type=order 時需提供產品編號（pd）
 * 
 * 回傳：ContentService.createTextOutput(JSON)，回傳對應下拉清單陣列
 */
function handleProductionEntryLoad(param) {
  const type = param.type; // 取得傳入的 type 參數（用來決定要抓哪一類資料）
  // ✅ 機台下拉選單：從 data 表 D 欄抓取
  if (type === 'machine') {
    return ContentService.createTextOutput(JSON.stringify(getMachineList()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // ✅ 產品編號下拉選單：從 data 表 E 欄抓取
  if (type === 'product') {
    return ContentService.createTextOutput(JSON.stringify(getProductList()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // ✅ 訂單編號下拉選單：需提供對應產品編號（pd），從訂單編號表中查找
  // if (type === 'order' && param.pd) {
  //   return ContentService.createTextOutput(JSON.stringify(getOrderListByProduct(param.pd)))
  //     .setMimeType(ContentService.MimeType.JSON);
  // }
  // ⚠️ 若 type 參數不符合上述任一情況，則回傳錯誤訊息
  return ContentService.createTextOutput(JSON.stringify({ error: "缺少必要參數" }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 🏭 getMachineList：取得機台清單
 * 說明：從「data」工作表 D 欄（第4欄）擷取機台名稱清單，忽略空值。
 * 回傳格式：一維陣列，如 ["機台1", "機台2", "NP20CM", ...]
 */
function getMachineList() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("data")               // 取得 data 表
    .getRange("A2:A")                     // 從第2列開始抓 D 欄（機台）
    .getValues()                          // 取得值，為二維陣列
    .flat()                               // 展平成一維陣列
    .filter(String);                      // 過濾空字串，只保留有值的機台名稱
}

/**
 * 🧾 getProductList：取得產品編號清單
 * 說明：從「data」工作表 E 欄（第5欄）擷取產品代號清單，忽略空值。
 * 回傳格式：一維陣列，如 ["Z1234", "X4567", "AZ789", ...]
 */
function getProductList() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("產品訂單編號")               // 取得 data 表
    .getRange("A2:A")                     // 從第2列開始抓 E 欄（產品編號）
    .getValues()                          // 取得值，為二維陣列
    .flat()                               // 展平成一維陣列
    .filter(String);                      // 過濾空字串，只保留有值的產品編號
}

/**
 * 📦 getOrderListByProduct：根據產品編號取得所有對應的訂單編號
 * 說明：從「訂單編號表」中，找出第1欄符合指定產品代號的列，並擷取第2欄（訂單編號）
 * 
 * 傳入參數：
 *   - pd：產品代號（例如 "Z1234"）
 * 回傳格式：一維陣列，如 ["ORD-001", "ORD-002"]
 */
// function getOrderListByProduct(pd) {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("產品訂單編號表");
//   const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues(); // 去除標題列
//   let orderNumbers = [];

//   data.forEach(row => {
//     if (row[0] === pd) {
//       orderNumbers = row.slice(1); // 從第2欄開始取到最後一欄
//     }
//   });

//   return orderNumbers.filter(String); // 移除空白或 undefined
// }

/**
 * 📊 sortProductionSheet：排序「加工數量表」資料
 * 說明：依據 A 欄（日期）與 B 欄（機台）做升冪排序，
 *       讓資料按照時間與機台順序排列，便於管理與查詢。
 */
// function sortProductionSheet() {
//   // 取得「加工數量表」工作表
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("加工數量表");
//   // 取得目前最後一列的列號，用來界定要排序的範圍
//   const lastRow = sheet.getLastRow();
//   // 取得從第2列（資料列）開始，到最後一列的 A~Z 欄範圍（保留完整欄位）
//   const range = sheet.getRange("A2:Z" + lastRow);
//   // 執行排序：
//   // 1. 先依 A 欄（日期）升冪排序
//   // 2. 若日期相同，依 B 欄（機台）升冪排序
//   range.sort([
//     { column: 1, ascending: true },  // A欄：日期
//     { column: 2, ascending: true }   // B欄：機台
//   ]);
// }

function sortProductionSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("加工數量表");
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  const data = range.getValues();

  data.sort((a, b) => {
    const dateA = parseDate(a[0]);
    const dateB = parseDate(b[0]);
    if (dateA - dateB !== 0) return dateA - dateB;

    const machineA = extractMachineCode(a[1]);
    const machineB = extractMachineCode(b[1]);
    return machineA.localeCompare(machineB);
  });

  range.setValues(data);
  SpreadsheetApp.flush(); // ← 這行會強制更新畫面顯示
}

function parseDate(input) {
  if (input instanceof Date) return input;
  if (typeof input === 'string') {
    const parts = input.split(/[\/\s:]+/);
    if (parts.length >= 3) {
      const year = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1;
      const day = parseInt(parts[2], 10);
      return new Date(year, month, day);
    }
  }
  return new Date('Invalid');
}

function extractMachineCode(str) {
  const match = str.toString().match(/\d+/);
  return match ? match[0].padStart(3, '0') : '000';
}



/**
 * 🕒 getLastSubmittedProduction：查詢「加工數量表」中最新一筆資料
 * 說明：根據第 9 欄（I 欄，時間戳記）找出最後一筆加工資料，
 *       回傳完整欄位內容與時間戳，供前端提示用。
 * 
 * 回傳格式：
 *   {
 *     dt: 日期,
 *     mc: 機台,
 *     pd: 產品,
 *     or: 訂單,
 *     hr: 小時,
 *     min: 分鐘,
 *     num: 數量,
 *     com: 備註,
 *     timestamp: 時間戳字串,
 *     rowIndex: Google Sheet 中實際列號
 *   }
 */
function getLastSubmittedProduction() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("加工數量表");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify(null))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // 時間戳改到 H 欄（index 7）
  const lastIndex = data.slice(1).reduce((latestIdx, row, i, arr) => {
    const timeA = new Date(arr[latestIdx][7]);
    const timeB = new Date(row[7]);
    return timeB > timeA ? i : latestIdx;
  }, 0) + 1;

  const lastRow = data[lastIndex];
  const rawTimestamp = new Date(lastRow[7]);
  const tz = Utilities.formatDate(rawTimestamp, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');

  const result = {
    dt:  lastRow[0], // A：日期
    mc:  lastRow[1], // B：機台
    pd:  lastRow[2], // C：產品
    hr:  lastRow[3], // D：小時（左移）
    min: lastRow[4], // E：分鐘
    num: lastRow[5], // F：數量
    com: lastRow[6], // G：備註
    timestamp: tz,   // H：時間戳
    rowIndex: lastIndex + 1
  };
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * ➕ addProductionData：新增一筆加工數據
 * 說明：寫入一筆加工登記資料（包含防重複、加上時間戳、排序），回傳成功與否結果。
 * 
 * 傳入參數（rowData）應包含：
 *   - dt：加工日期（YYYY-MM-DD）
 *   - mc：機台名稱
 *   - pd：產品編號
 *   - or：訂單編號
 *   - hr：小時
 *   - min：分鐘
 *   - num：加工數量
 *   - comments：備註
 */
function addProductionData(rowData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("加工數量表");
  const currentDate = new Date();

  // 只有標題列時
  if (sheet.getLastRow() < 2) {
    sheet.appendRow([
      rowData.dt,       // A：日期
      rowData.mc,       // B：機台
      rowData.pd,       // C：產品
      rowData.hr,       // D：小時（已左移）
      rowData.min,      // E：分鐘
      rowData.num,      // F：數量
      rowData.comments, // G：備註
      currentDate       // H：時間戳
    ]);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1, 1, 8).setHorizontalAlignment("center"); // 欄數改 8
    sortProductionSheet();
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 取既有資料（不含標題）→ 欄數改 8
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();

  // 日期格式化（yyyy/m/d）
  const inputDate = new Date(rowData.dt);
  const formattedInputDate = `${inputDate.getFullYear()}/${inputDate.getMonth() + 1}/${inputDate.getDate()}`;

  // 防重複：同一天 + 同機台 + 同產品（已去掉訂單欄）
  for (let i = 0; i < dataRange.length; i++) {
    const existingDate = new Date(dataRange[i][0]);
    const formattedExistingDate = `${existingDate.getFullYear()}/${existingDate.getMonth() + 1}/${existingDate.getDate()}`;
    if (
      formattedExistingDate === formattedInputDate &&
      dataRange[i][1] === rowData.mc && // 機台
      dataRange[i][2] === rowData.pd    // 產品
    ) {
      return ContentService.createTextOutput(JSON.stringify({ success: false }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // 寫入新資料（無訂單欄）
  sheet.appendRow([
    formattedInputDate, // A
    rowData.mc,         // B
    rowData.pd,         // C
    rowData.hr,         // D
    rowData.min,        // E
    rowData.num,        // F
    rowData.comments,   // G
    currentDate         // H
  ]);
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1, 1, 8).setHorizontalAlignment("center"); // 欄數改 8
  sortProductionSheet();
  return ContentService.createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ❌ deleteLastProductionRow：刪除「加工數量表」中最後一筆資料
 * 說明：根據時間戳（I欄）找出時間最新的一筆資料並刪除，常搭配「誤登記後立即撤回」的功能。
 * 
 * 回傳格式：
 *   - { success: true } 若刪除成功
 *   - { success: false } 若表中無資料可刪
 */
function deleteLastProductionRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("加工數量表");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify({ success: false }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // 時間戳改 H 欄（index 7）
  const lastIndex = data.slice(1).reduce((latestIdx, row, i, arr) => {
    const timeA = new Date(arr[latestIdx][7]);
    const timeB = new Date(row[7]);
    return timeB > timeA ? i : latestIdx;
  }, 0) + 2;

  sheet.deleteRow(lastIndex);
  return ContentService.createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * 🔍 getLatestProductionByProduct：查詢指定產品.訂單的最後一次加工紀錄
 * 說明：先找出該產品最後加工的日期，再擷取當天的所有加工紀錄（可能多筆），供前端顯示使用。
 * 
 * 傳入參數：
 *   - pd：產品編號（如 Z1123）
 * 
 * 回傳格式：
 *   - [[日期, 組合文字, 數量], ...]
 *     → 組合文字格式為「產品編號 . 機台名稱」
 */
function getLatestProductionByProduct(pd) {
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("加工數量表");
  const data = ws.getDataRange().getValues();
  let filteredData = [];
  let latestDate = null;

  // 找最後加工日期（只看 C 欄）
  for (let i = 1; i < data.length; i++) {
    const recordDate  = new Date(data[i][0]); // A
    const productCode = data[i][2];           // C
    if (productCode === pd) {
      if (!latestDate || recordDate > latestDate) latestDate = recordDate;
    }
  }

  if (latestDate) {
    const d = `${latestDate.getFullYear()}/${latestDate.getMonth() + 1}/${latestDate.getDate()}`;
    for (let j = 1; j < data.length; j++) {
      const rd  = new Date(data[j][0]); // A
      const pc  = data[j][2];           // C
      const mc  = data[j][1];           // B
      const qty = data[j][5];           // F（數量，原 G 左移）
      const dd  = `${rd.getFullYear()}/${rd.getMonth() + 1}/${rd.getDate()}`;
      if (pc === pd && dd === d) filteredData.push([dd, `${pc} . ${mc}`, qty]);
    }
  }
  return ContentService.createTextOutput(JSON.stringify(filteredData))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * 📅 getProductionQuery：查詢某產品在特定日期區間內的所有加工紀錄
 * 說明：篩選出符合產品編號、加工日期在指定範圍內的所有資料，格式化後回傳。
 *
 * 傳入參數 param 包含：
 *   - pd：產品編號（例如 Z1234）
 *   - dt1：起始日期（例如 2024-04-01）
 *   - dt2：結束日期（例如 2024-04-11）
 * 
 * 回傳格式：
 *   - [[日期, "產品 . 機台", 數量], ...]（JSON 格式）
 */
function getProductionQuery(param) {
  const { pd, dt1, dt2 } = param;
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("加工數量表");
  const data = ws.getDataRange().getValues();
  const filteredData = [];
  const startDate = new Date(dt1);
  const endDate   = new Date(dt2);

  for (let i = 1; i < data.length; i++) {
    const recordDate  = new Date(data[i][0]); // A
    const machineCode = data[i][1];           // B
    const productCode = data[i][2];           // C
    const quantity    = data[i][5];           // F（數量）

    if (productCode === pd && recordDate >= startDate && recordDate <= endDate) {
      const d = `${recordDate.getFullYear()}/${recordDate.getMonth() + 1}/${recordDate.getDate()}`;
      filteredData.push([d, `${productCode} . ${machineCode}`, quantity]);
    }
  }
  return ContentService.createTextOutput(JSON.stringify(filteredData))
    .setMimeType(ContentService.MimeType.JSON);
}



///////////////////////////////////////////////////////////////////////
/**
 * 🔽 匯出資料查詢：取得指定產品與訂單編號的所有資料
 */
function exportProductionData(productCode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("加工數量表");
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(`A1:G${lastRow}`).getDisplayValues(); // ← 到 G

  const filtered = data.filter((row, index) => {
    if (index === 0) return false;     // 跳過標題
    return row[2] === productCode;     // C：產品
  });

  return ContentService.createTextOutput(JSON.stringify(filtered))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 🔽 匯出用下拉選單選項（產品與訂單編號）
 */
function getDropdownValues() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("加工數量表");
  const data = sheet.getDataRange().getValues();

  const productSet = new Set();
  data.forEach((row, index) => {
    if (index === 0) return;
    if (row[2]) productSet.add(row[2]);
  });

  return ContentService.createTextOutput(JSON.stringify({
    products: [...productSet]
  })).setMimeType(ContentService.MimeType.JSON);
}