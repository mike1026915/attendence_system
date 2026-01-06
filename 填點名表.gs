/**
 * Configuration Area
 */
const CONFIG = {
  sourceSheetName: "Records", // Source sheet name
  targetSheetName: "點名表", // Target sheet name
  dateCell: "F2", // [MODIFIED] Moved to F2 (Shifted right)

  // Mapping parts to column indices (0=1st data col, 1=2nd data col...)
  partMapping: {
    Conductor: 0,
    T1: 1,
    T2: 2,
    B1: 3,
    B2: 4,
  },

  // Display names for headers (Order must match indices 0 to 4)
  partHeaders: ["Conductor", "T1", "T2", "B1", "B2"],

  defaultColumnIndex: 0,

  // Status configuration
  statusMap: {
    出席: { symbol: "V", color: "#ffe599" }, // Yellow
    請假: { symbol: "△", color: "#f4c7c3" }, // Red
    晚到: { symbol: "L", color: "#cfe2f3" }, // Blue
    遲到: { symbol: "L", color: "#cfe2f3" },
  },
};

function generateAttendanceByDate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(CONFIG.sourceSheetName);
  const targetSheet = ss.getSheetByName(CONFIG.targetSheetName);

  if (!sourceSheet || !targetSheet) {
    SpreadsheetApp.getUi().alert("Sheet not found, please check names!");
    return;
  }

  // --- 1. Handle Date Logic ---
  const timezone = Session.getScriptTimeZone();
  const inputDateValue = targetSheet.getRange(CONFIG.dateCell).getValue();
  let targetDateStr = "";

  // [MODIFIED] Enhanced Date Styling
  const dateRange = targetSheet.getRange(CONFIG.dateCell);

  // Check if cell has a date
  if (inputDateValue instanceof Date) {
    targetDateStr = Utilities.formatDate(
      inputDateValue,
      timezone,
      "yyyy/MM/dd"
    );
  } else {
    // Default to today if empty
    const today = new Date();
    targetDateStr = Utilities.formatDate(today, timezone, "yyyy/MM/dd");
    dateRange.setValue(targetDateStr);
  }

  // Apply styles to date cell
  dateRange
    .setFontSize(12) // Larger font
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#fff2cc") // Light yellow background
    .setBorder(true, true, true, true, null, null); // Box border

  // --- 2. Read and Filter Data ---
  const data = sourceSheet.getDataRange().getValues();
  const memberData = {}; // Store unique member status

  // Start from row 2 (skip header)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const timestamp = row[0];

    // Filter by date
    if (timestamp instanceof Date) {
      const rowDateStr = Utilities.formatDate(
        timestamp,
        timezone,
        "yyyy/MM/dd"
      );
      if (rowDateStr !== targetDateStr) {
        continue;
      }
    } else {
      continue;
    }

    const name = row[1];
    const part = row[2];
    const status = row[3];

    // Use latest submission for duplicate names
    if (name) {
      memberData[name] = {
        name: name,
        part: part,
        status: status,
        timestamp: timestamp,
      };
    }
  }

  // --- 3. Distribute Members to Columns ---
  const columns = [[], [], [], [], []];

  Object.values(memberData).forEach((member) => {
    let colIndex = CONFIG.partMapping[member.part];
    if (colIndex === undefined) colIndex = CONFIG.defaultColumnIndex;
    if (colIndex !== null && columns[colIndex]) {
      columns[colIndex].push(member);
    }
  });

  // --- 3.1 Sort Members: "出席" first ---
  columns.forEach((colList) => {
    colList.sort((a, b) => {
      const isAPresent = a.status === "出席";
      const isBPresent = b.status === "出席";
      if (isAPresent && !isBPresent) return -1;
      if (!isAPresent && isBPresent) return 1;
      return a.timestamp - b.timestamp;
    });
  });

  // --- 4. Write to Target Sheet ---
  const headerRow = 5;
  const startRow = 6;

  // [MODIFIED] Shifted columns right (B, D, F, H, J) -> (2, 4, 6, 8, 10)
  const colPositions = [2, 4, 6, 8, 10];

  // [MODIFIED] Clear larger range and clear borders/formatting
  targetSheet
    .getRange("A5:K1000")
    .clearContent()
    .setBackground(null)
    .setBorder(false, false, false, false, false, false);

  let globalTotalCount = 0;
  let globalPresentCount = 0;

  // Calculate the maximum number of members in any column for alignment
  const maxMembers = Math.max(...columns.map((col) => col.length));

  // Define footer position (Members + 1 empty row buffer)
  // Ensure at least a few rows exist even if empty
  const rowsNeeded = Math.max(maxMembers, 1);
  const footerStartRow = startRow + rowsNeeded;

  columns.forEach((memberList, colIndex) => {
    const targetCol = colPositions[colIndex];

    // Define Ranges
    const headerRange = targetSheet.getRange(headerRow, targetCol, 1, 2);
    const bodyRange = targetSheet.getRange(startRow, targetCol, rowsNeeded, 2);
    const footerRange = targetSheet.getRange(footerStartRow, targetCol, 2, 2); // Footer is 2 rows high
    const entireBlockRange = targetSheet.getRange(
      headerRow,
      targetCol,
      footerStartRow + 2 - headerRow,
      2
    );

    // 4.1 Write Header (Row 5)
    headerRange.getCell(1, 1).setValue(CONFIG.partHeaders[colIndex]);
    headerRange
      .merge() // Merge the 2 cells for header
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBackground("#efefef");

    let colPresentCount = 0;

    // 4.2 Write Members
    memberList.forEach((member, rowIndex) => {
      const currentRow = startRow + rowIndex;
      const nameCell = targetSheet.getRange(currentRow, targetCol);
      const statusCell = targetSheet.getRange(currentRow, targetCol + 1);

      nameCell
        .setValue(member.name)
        .setFontWeight("bold")
        .setFontSize(12)
        .setFontColor("#000000");

      const config = CONFIG.statusMap[member.status];
      if (config) {
        statusCell
          .setValue(config.symbol)
          .setBackground(config.color)
          .setHorizontalAlignment("center");

        if (config.symbol === "V") {
          colPresentCount++;
          globalPresentCount++;
        }
      } else {
        statusCell.setValue(member.status);
      }

      globalTotalCount++;
    });

    // 4.3 Write Column Statistics (Footer)
    // Background for footer
    footerRange.setBackground("#f9f9f9");

    // Write "Attendance Count"
    targetSheet
      .getRange(footerStartRow, targetCol)
      .setValue(`出席: ${colPresentCount}`)
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontColor("#000000");

    // Write "Total Count"
    targetSheet
      .getRange(footerStartRow + 1, targetCol)
      .setValue(`總數: ${memberList.length}`)
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontColor("#000000");

    // --- [MODIFIED] Borders & Layout Styling ---

    // 1. Box around the entire column (Header to Footer)
    entireBlockRange.setBorder(true, true, true, true, null, null);

    // 2. Line under Header
    headerRange.setBorder(null, null, true, null, null, null);

    // 3. Line above Footer (Separator between members and stats)
    footerRange.setBorder(true, null, null, null, null, null);

    // 4. Vertical line inside the block (between Name and Status) is removed for cleaner look,
    // or we can add it back if strictly needed. For "clean", usually no vertical line is better,
    // but the status colors provide separation.
  });

  // --- 5. Update Global Statistics ---
  // [MODIFIED] Moved to H3 (Shifted Right) and styled
  const globalStatCell = targetSheet.getRange("H3");
  globalStatCell
    .setValue(`總出席: ${globalPresentCount}/${globalTotalCount}`)
    .setFontWeight("bold")
    .setFontSize(12)
    .setFontColor("#c90000")
    .setBackground("#fff2cc")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, null, null);

  // Toast notification
  SpreadsheetApp.getActive().toast(`已更新 ${targetDateStr} 的點名表`, "完成");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("點名系統")
    .addItem("更新今日/指定日期點名", "generateAttendanceByDate")
    .addToUi();
}
