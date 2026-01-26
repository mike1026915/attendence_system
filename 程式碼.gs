// DATABASE CONFIGURATION
const RECORD_SHEET_NAME = "Records";
const MEMBER_SHEET_NAME = "name_list";

function sanitizeUserInput(input) {
  if (typeof input !== "string") {
    return "";
  }

  let value = input.trim();

  // 1️⃣ 防止 Spreadsheet Formula Injection
  if (/^[=+\-@]/.test(value)) {
    value = "'" + value; // 加單引號，Sheets 會當純文字
  }

  // 2️⃣ 基本 XSS 防禦（HTML escape）
  value = value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");

  return value;
}

// Handle POST requests (Submitting data)
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RECORD_SHEET_NAME);

    const data = JSON.parse(e.postData.contents);

    const safeName = sanitizeUserInput(data.name);
    const safeVoicePart = sanitizeUserInput(data.voicePart);
    const safeStatus = sanitizeUserInput(data.status);
    const safeNote = sanitizeUserInput(data.note || "");

    sheet.appendRow([
      new Date(),
      safeName,
      safeVoicePart,
      safeStatus,
      safeNote,
    ]);

    return ContentService.createTextOutput(
      JSON.stringify({ result: "success" })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ result: "error", error: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Handle GET requests (Fetching data for display)
function doGet(e) {
  // Check if action is 'getImage'
  if (e.parameter.action === 'getImage') {
    try {
      // Password verification
      const providedPassword = e.parameter.password || "";
      const storedPassword = PropertiesService.getScriptProperties().getProperty('IMAGE_EXPORT_PASSWORD') || "";

      if (!providedPassword || providedPassword !== storedPassword) {
        return ContentService.createTextOutput(
          JSON.stringify({ result: "error", error: "Invalid password" })
        ).setMimeType(ContentService.MimeType.JSON);
      }

      const timezone = Session.getScriptTimeZone();
      let dateStr = "";

      // Get date from parameter or use today's date
      if (e.parameter.date) {
        dateStr = e.parameter.date;
        // Validate date format (yyyy/MM/dd)
        if (!/^\d{4}\/\d{2}\/\d{2}$/.test(dateStr)) {
          return ContentService.createTextOutput(
            JSON.stringify({ result: "error", error: "Invalid date format. Please use yyyy/MM/dd" })
          ).setMimeType(ContentService.MimeType.JSON);
        }
      } else {
        // Use today's date if not provided
        const today = new Date();
        dateStr = Utilities.formatDate(today, timezone, "yyyy/MM/dd");
      }

      // Calculate attendance data using the same logic as exportAttendanceImage
      let attendanceData;
      try {
        attendanceData = calculateAttendanceDataByDate(dateStr);
      } catch (error) {
        return ContentService.createTextOutput(
          JSON.stringify({ result: "error", error: "Failed to calculate attendance data: " + error.message })
        ).setMimeType(ContentService.MimeType.JSON);
      }

      // Generate HTML content
      const htmlContent = generateAttendanceHtml(attendanceData, dateStr);

      // Return HTML page
      return HtmlService.createHtmlOutput(htmlContent)
        .setTitle("Attendance Image - " + dateStr)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } catch (error) {
      return ContentService.createTextOutput(
        JSON.stringify({ result: "error", error: error.toString() })
      ).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Original doGet logic for fetching records and members
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 取得打卡紀錄 (用於顯示列表)
  const recordSheet = ss.getSheetByName(RECORD_SHEET_NAME);
  const recordRows = recordSheet.getDataRange().getValues();
  recordRows.shift(); // 移除標題
  const records = recordRows
    .map((row) => ({
      timestamp: row[0],
      name: sanitizeUserInput(row[1]),
      voicePart: sanitizeUserInput(row[2]),
      status: row[3],
      note: row[4],
    }))
    .reverse(); // Show newest first

  // 2. 取得成員名單 (用於下拉選單)
  const memberSheet = ss.getSheetByName(MEMBER_SHEET_NAME);
  let members = [];
  if (memberSheet) {
    const memberRows = memberSheet.getDataRange().getValues();
    memberRows.shift(); // 移除標題
    members = memberRows
      .map((row) => ({
        voicePart: sanitizeUserInput(row[0]),
        name: sanitizeUserInput(row[1]),
      }))
      .filter((item) => item.name && item.voicePart); // 過濾空行
  }

  // 回傳整合的 JSON 物件
  const result = {
    records: records,
    members: members,
  };

  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(
    ContentService.MimeType.JSON
  );
}
