//To get the final result select runAllTasks -> Run
//-----------------------------------------------------------------------------------------------------------------------------------------------
// Connecting and Authenticating
function getAccessToken() {
  const url = "https://open.larksuite.com/open-apis/auth/v3/tenant_access_token/internal";
  const payload = { 
    app_id: 'cli_a7fab27260385010', // EDIT app_id when needed
    app_secret: 'Zg4MVcFfiOu0g09voTcpfd4WGDpA0Ly5' // EDIT app_secret when needed
  };
    try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    });

    const result = JSON.parse(response.getContentText()); 
        Logger.log("Successfully obtained tenant access token from Lark. Expires in " + result.expire + " seconds."); // Check response
  
    if (result.code !== 0) throw new Error(result.msg); 
    return result.tenant_access_token;
  } catch (e) {
    Logger.log("Error in getting access Token: " + e.message); // Check response
    return null;
  }
}
// Note 1: Testing the Access Token
function testAccessToken() {
  const token = getAccessToken();
  Logger.log("Success in getting access Token: " + token); // Check response
}
//-----------------------------------------------------------------------------------------------------------------------------------------------
// Getting data from Larkbase
function fetchLarkData() { 
  const token = getAccessToken();
  if (!token) return;

  const APP_TOKEN = "GBhdbr6g4ajxbgsGLgqlTIvegth"; // EDIT APP_TOKEN when needed
  const TABLE_ID = "tbllWVANnPigwOc0"; // EDIT TABLE_ID when needed
  const baseUrl = `https://open.larksuite.com/open-apis/bitable/v1/apps/${APP_TOKEN}/tables/${TABLE_ID}/records`; 

  let allRecords = [];
  let pageToken = "";
  let hasMore = true;

  while (hasMore) {
    const url = pageToken ? `${baseUrl}?page_token=${pageToken}` : baseUrl;

    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        Authorization: `Bearer ${token}`
      }
    });

    const result = JSON.parse(response.getContentText()); 

    if (result.code !== 0) {
      Logger.log("Eror fetching Data: " + result.msg); // Check response
      break;
    }

    allRecords = allRecords.concat(result.data.items);
    hasMore = result.data.has_more;
    pageToken = result.data.page_token || "";
  }
  return allRecords;
}

// Note 2: Testing the Access Token
function testFetchLarkData() {
  const records = fetchLarkData();
  if (records) {
    Logger.log("Records fetched: " + records.length); // Check response
    Logger.log(records[0]); // Print first recorded to check structure
  }
}

//-----------------------------------------------------------------------------------------------------------------------------------------------
// Data cleaning
function convertTimestamp(ts) {
  if (!ts || typeof ts !== "number") return "";
  const date = new Date(ts);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}

function writeDataToSheet() {
  const records = fetchLarkData();
  if (!records || records.length === 0) {
    Logger.log("No data written to Sheet"); // Check response
    return;
  }

  const sheetName = "Data";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  // If no Sheet then create New Sheet
  if (!sheet) {  
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clearContents(); // If exists then Clear
  }

  const headers = [
    "Request No.", "Status", "Submitted at", "Completed at",
    "Nội dung_Tên sản phẩm", "Nội dung_Hạng mục đầu tư",
    "Nội dung_Cơ sở kinh doanh", "Nội dung_Số lượng",
    "Nội dung_Đơn giá", "Nội dung_Tên nhà cung cấp"
  ];


  sheet.appendRow(headers);   // Headers 

const dataRows = records.map(item => { 
  const f = item.fields;
  return [
    f["Request No."] || "",
    f["Status"] || "",
    convertTimestamp(f["Submitted at"]),
    convertTimestamp(f["Completed at"]),
    f["Nội dung_Tên sản phẩm"] || "",
    f["Nội dung_Hạng mục đầu tư"] || "",
    f["Nội dung_Cơ sở kinh doanh"] || "",
    f["Nội dung_Số lượng"] || "",
    f["Nội dung_Đơn giá"] || "",
    f["Nội dung_Tên nhà cung cấp"] || ""
  ];
});

  // Copying Data_V2-Optimized to copy at the same time 
sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
  Logger.log("Fetched " + records.length + " records into 'Data' sheet");  // Check response
}

//-----------------------------------------------------------------------------------------------------------------------------------------------
// Creating Summary report
function createSummaryReport() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Data");
  const summarySheetName = "Summary"; 

  if (!dataSheet) {
    Logger.log("Sheet 'Data' doesn't exist");  // Check response
    return;
  }

  const data = dataSheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1); 
  const index = {
    requestNo: headers.indexOf("Request No."),
    status: headers.indexOf("Status"),
    quantity: headers.indexOf("Nội dung_Số lượng"),
    price: headers.indexOf("Nội dung_Đơn giá"),
    location: headers.indexOf("Nội dung_Cơ sở kinh doanh")
  };

  let totalRequests = 0;
  let totalCost = 0;
  const statusCount = {};
  const locationCost = {}; 

  rows.forEach(row => {
    totalRequests += 1;

    const status = row[index.status] || "Other";
    statusCount[status] = (statusCount[status] || 0) + 1;

    const qty = parseFloat(row[index.quantity]) || 0;
    const price = parseFloat(row[index.price]) || 0;
    const cost = qty * price;
    totalCost += cost;

    const location = row[index.location] || "Unidentified";
    locationCost[location] = (locationCost[location] || 0) + cost;
  });


  // Create or clear Summary sheet 
  let summarySheet = ss.getSheetByName(summarySheetName);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(summarySheetName);
  } else {
    summarySheet.clearContents();
  }

  let row = 1;

   summarySheet.getRange(row, 1, 1, 3).setValues([["I", "Tổng số request:", totalRequests]]);  // Tổng số request
  row++;
  summarySheet.getRange(row, 1, 1, 3).setValues([["II", "Tổng chi phí:", totalCost]]);  // Tổng chi phí
  row+=2; 
  summarySheet.getRange(row, 1, 1, 2).setValues([["III", "Số request theo Status"]]);  // Báo cáo theo Status
  row++;

  let stt = 1;
  for (let status in statusCount) {
    summarySheet.getRange(row, 1).setValue(stt++);
    summarySheet.getRange(row, 2).setValue(status);
    summarySheet.getRange(row, 3).setValue(statusCount[status]);
    row++;
}
row++; // Extra row

// Báo cáo Top 5 sản phẩm được yêu cầu nhiều nhất
summarySheet.getRange(row, 1, 1, 3).setValues([["IV", "Top 5 sản phẩm được yêu cầu nhiều nhất", "Tổng SL được yêu cầu"]]);
row++;

// Đếm số lượng theo tên sản phẩm
const productCount = {};
rows.forEach(row => {
  const name = row[headers.indexOf("Nội dung_Tên sản phẩm")] || "Không xác định";
  const qty = parseFloat(row[headers.indexOf("Nội dung_Số lượng")]) || 0;
  productCount[name] = (productCount[name] || 0) + qty;
});

// Sắp xếp và lấy top 5
const top5 = Object.entries(productCount)
                   .sort((a, b) => b[1] - a[1])
                   .slice(0, 5);

// Input Data
stt = 1;
top5.forEach(([product, totalQty]) => {
  summarySheet.getRange(row, 1).setValue(stt++);
  summarySheet.getRange(row, 2).setValue(product);
  summarySheet.getRange(row, 3).setValue(totalQty);
  row++;
});

  row++; // Extra row

  // Báo cáo theo Cơ sở kinh doanh
  summarySheet.getRange(row, 1).setValue("V");
  summarySheet.getRange(row, 2).setValue("Tổng chi phí theo Cơ sở kinh doanh");
  row++;

stt = 1;
for (let location in locationCost) {
  summarySheet.getRange(row, 1).setValue(stt++);
  summarySheet.getRange(row, 2).setValue(location);
  summarySheet.getRange(row, 3).setValue(locationCost[location]);
  row++;
}

  Logger.log("Created 'Summary'Sheet");  // Check response
}

//-----------------------------------------------------------------------------
// Note: Formatting Data sheet
function formatDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data");
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const fullRange = sheet.getRange(1, 1, lastRow, lastCol);

  // 1. General format
  fullRange.setFontSize(10); // EDIT Font size
  fullRange.setFontFamily("Arial"); // EDIT Font
  fullRange.setVerticalAlignment("middle");
  fullRange.setBorder(false, false, false, false, false, false); // Clear border of previous Data

  // 2. Column Format
  // Auto resize columns: A,C,D,G,F
  const autoCols = [1, 3, 4, 7, 6];
  autoCols.forEach(col => {
    sheet.autoResizeColumn(col);
  });

  // 3. Set width for columns B,E,H,I,J + Wrap text + Bold Header
  sheet.setColumnWidth(2, 100); // Column B
  sheet.setColumnWidth(5, 250); // Column E
  sheet.setColumnWidth(8, 135); // Column H
  sheet.setColumnWidth(9, 130); // Column I
  sheet.setColumnWidth(10, 230); // Column J
  sheet.getRange(2, 5, lastRow - 1).setWrap(true);  // Column E Wrap text
  sheet.getRange(2, 10, lastRow - 1).setWrap(true); // Column J Wrap text
  sheet.getRange(1, 1, 1, lastCol).setFontWeight("bold");  // Bold Header

  // 4. Column formatting
  const datetimeFormat = 'dd"/"mm"/"yyyy"   "hh":"mm';   // Column C and D = Format to Time
  sheet.getRange(2, 3, lastRow - 1).setNumberFormat(datetimeFormat); // Columnn C (Submitted at)
  sheet.getRange(2, 4, lastRow - 1).setNumberFormat(datetimeFormat); // Columnn D (Completed at)
  sheet.getRange(2, 9, lastRow - 1).setNumberFormat("#,##0"); // Column I (Nội dung_Đơn giá) = Number format with Thousand separator

  // 5. Create border for Data
  fullRange.setBorder(true, true, true, true, true, true);

  // Freeze row header for easy view
  sheet.setFrozenRows(1);

  // 6. Conditional formatting for Column B (Status) for easy view
  const rules = sheet.getConditionalFormatRules();

  const addStatusRule = (text, color) => {
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(text)
      .setBackground(color)
      .setRanges([sheet.getRange(`B2:B${lastRow}`)])
      .build();
    rules.push(rule);
  };
  // Colors gotten from Lark
  addStatusRule("Approved",      "#D6DFF4");
  addStatusRule("Deleted",       "#FEE7CD");
  addStatusRule("Processing",    "#CAEFFC");
  addStatusRule("Rejected",      "#FAEDC2");
  addStatusRule("Canceled",      "#C4F2EC");
  addStatusRule("Terminated",    "#FEE3E2");
  addStatusRule("Recalled",      "#EFE6FE");
  addStatusRule("Under Review",  "#D0F5CE");

  sheet.setConditionalFormatRules(rules);

  Logger.log("Formatted Data Sheet");  // Check response
}

//-----------------------------------------------------------------------------
// Note: Formatting Summary sheet
function formatSummarySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Summary");
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const fullRange = sheet.getRange(1, 1, lastRow, lastCol);

  // 1. General format
  fullRange.setFontSize(10); // EDIT Font size
  fullRange.setFontFamily("Arial"); // EDIT Font
  fullRange.setVerticalAlignment("middle");
  fullRange.setBorder(false, false, false, false, false, false); // Clear border of previous Data

  // 2. Column Format
  sheet.autoResizeColumn(1); // Auto resize Column A (№)
  sheet.getRange(1, 1, lastRow).setHorizontalAlignment("center"); // Middle aign Column A (№) for easy view
  sheet.setColumnWidth(2, 300); // Column B width = 300px
  sheet.getRange(1, 2, lastRow).setWrap(true); // Set Columb B to Wrap text
  sheet.setColumnWidth(3, 155); // Column C width = 155px
  sheet.getRange(1, 3, lastRow).setNumberFormat("#,##0"); // Column I (Aggregated Numbers) = Number format with Thousand separator

  // 3. Bold Important text
  const boldRows = [1, 2, 4, 9, 10, 17]; // EDIT Row number if needed
  boldRows.forEach(r => {
    sheet.getRange(r, 1, 1, lastCol).setFontWeight("bold");
  });
  
  // 5. Create border: EDIT Row number if needed
  const borderRanges = [
    { row: 1, numRows: 2 },       // Request + total cost
    { row: 4, numRows: 5 },       // Status table
    { row: 10, numRows: 6 },      // Top 5 product
    { row: 17, numRows: lastRow - 17 + 1 }  // Location
  ];

  borderRanges.forEach(({ row, numRows }) => {
    sheet.getRange(row, 1, numRows, lastCol).setBorder(true, true, true, true, true, true);
  });

  // 6. Moving Summary Sheet to After Data Sheet_v2 - Fixed for when Data sheet is on the right most side
  const dataSheet = ss.getSheetByName("Data");
  const summarySheet = ss.getSheetByName("Summary");

  if (dataSheet && summarySheet) {
  const dataIndex = dataSheet.getIndex();
  const summaryIndex = summarySheet.getIndex();
  const totalSheets = ss.getSheets().length;
  const targetIndex = Math.min(dataIndex + 1, totalSheets); // Prevent max Index value

  if (summaryIndex <= dataIndex) {
    ss.setActiveSheet(summarySheet);
    ss.moveActiveSheet(targetIndex);
    }
  }
  Logger.log("Formatted Summary Sheet"); // Check response
}
//-----------------------------------------------------------------------------
function runAllTasks() {
  Logger.log("Executing All Tasks"); // Check response

  writeDataToSheet();
  formatDataSheet();
  createSummaryReport();
  formatSummarySheet();

  Logger.log("Finished All Tasks"); // Check response
}

