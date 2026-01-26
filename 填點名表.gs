/**
 * Configuration Area
 */
const CONFIG = {
  sourceSheetName: "Records", // Source sheet name
  targetSheetName: "點名表", // Target sheet name
  memberSheetName: "name_list", // Member list sheet name
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
    晚到: { symbol: "L", color: "#cfe2f3" }, // Blue
    請假: { symbol: "△", color: "#f4c7c3" }, // Red
    遲到: { symbol: "L", color: "#cfe2f3" },
  },
};

/**
 * Convert column number to column letter (1 -> A, 2 -> B, etc.)
 */
function columnToLetter(column) {
  let temp,
    letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Normalize name by trimming whitespace
 */
function normalizeName(name) {
  return name ? name.toString().trim() : "";
}

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

    const name = normalizeName(row[1]);
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

  // --- 2.5. Read All Members from name_list ---
  const memberSheet = ss.getSheetByName(CONFIG.memberSheetName);
  const allMembersByPart = {}; // { "T1": Set(["name1", "name2"]), "T2": [...], ... }
  const allMembersSet = new Set(); // For quick lookup

  if (memberSheet) {
    const memberDataRows = memberSheet.getDataRange().getValues();
    // Start from row 2 (skip header)
    for (let i = 1; i < memberDataRows.length; i++) {
      const row = memberDataRows[i];
      const part = row[0];
      const name = normalizeName(row[1]);
      if (part && name) {
        if (!allMembersByPart[part]) {
          allMembersByPart[part] = new Set(); // Use Set for automatic deduplication
        }
        allMembersByPart[part].add(name); // Use add for automatic deduplication
        allMembersSet.add(name);
      }
    }
  }

  // --- 3. Distribute Members to Columns ---
  const columns = [[], [], [], [], []];

  // Add checked-in members
  Object.values(memberData).forEach((member) => {
    let colIndex = CONFIG.partMapping[member.part];
    if (colIndex === undefined) colIndex = CONFIG.defaultColumnIndex;
    if (colIndex !== null && columns[colIndex]) {
      columns[colIndex].push(member);
    }
  });

  // Add missing members (not checked in) to columns
  const checkedInNames = new Set(Object.keys(memberData));
  CONFIG.partHeaders.forEach((partName) => {
    const colIndex = CONFIG.partMapping[partName];
    if (colIndex !== undefined && colIndex !== null && columns[colIndex]) {
      const allMembersInPart = allMembersByPart[partName]
        ? Array.from(allMembersByPart[partName])
        : [];
      allMembersInPart.forEach((name) => {
        if (!checkedInNames.has(name)) {
          // Add missing member with null status
          columns[colIndex].push({
            name: name,
            part: partName,
            status: null,
            timestamp: null,
          });
        }
      });
    }
  });

  // --- 3.1 Sort Members: "出席" first, then "晚到", then "請假", then others ---
  columns.forEach((colList) => {
    colList.sort((a, b) => {
      // Define status priority: 出席 (1) > 晚到 (2) > 請假 (3) > others (4)
      const getStatusPriority = (status) => {
        if (status === "出席") return 1;
        if (status === "晚到") return 2;
        if (status === "請假") return 3;
        return 4;
      };

      const priorityA = getStatusPriority(a.status);
      const priorityB = getStatusPriority(b.status);
      if (priorityA !== priorityB) {
        return priorityA - priorityB;
      }

      // Same status priority, then sort by timestamp (checked-in members), missing members (null timestamp) go last
      if (a.timestamp === null && b.timestamp !== null) return 1;
      if (a.timestamp !== null && b.timestamp === null) return -1;
      if (a.timestamp !== null && b.timestamp !== null) {
        return a.timestamp - b.timestamp;
      }
      // Both are missing, sort by name
      return a.name.localeCompare(b.name);
    });
  });

  // --- 4. Write to Target Sheet ---
  const headerRow = 5;
  const startRow = 6;

  // [MODIFIED] Shifted columns right (B, D, F, H, J) -> (2, 4, 6, 8, 10)
  const colPositions = [2, 4, 6, 8, 10];

  // [MODIFIED] Clear larger range and clear borders/formatting (including missing members area)
  targetSheet
    .getRange("A5:R1000")
    .clearContent()
    .setBackground(null)
    .setBorder(false, false, false, false, false, false);

  // Also clear the missing members title area
  targetSheet
    .getRange("M3:R4")
    .clearContent()
    .setBackground(null)
    .setBorder(false, false, false, false, false, false);

  // Calculate total members count (all members, not just checked-in)
  let globalTotalCount = allMembersSet.size;
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
    const footerRange = targetSheet.getRange(footerStartRow, targetCol, 3, 2); // Footer is 3 rows high (出席, 總數, 出席率)
    const entireBlockRange = targetSheet.getRange(
      headerRow,
      targetCol,
      footerStartRow + 3 - headerRow,
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

      // Handle status display
      if (member.status === null) {
        // Missing member (not checked in) - leave status cell empty
        statusCell.setValue("");
      } else {
        const config = CONFIG.statusMap[member.status];
        if (config) {
          statusCell
            .setValue(config.symbol)
            .setBackground(config.color)
            .setHorizontalAlignment("center");

          // Count both "出席" (V) and "晚到" (L) as present
          if (config.symbol === "V" || config.symbol === "L") {
            colPresentCount++;
            globalPresentCount++;
          }
        } else {
          statusCell.setValue(member.status);
        }
      }
    });

    // 4.3 Write Column Statistics (Footer)
    // Background for footer
    footerRange.setBackground("#f9f9f9");

    // Get part name for this column
    const partName = CONFIG.partHeaders[colIndex];
    // Get total members count for this part (use actual displayed count, not just name_list)
    // This includes members from Records who may not be in name_list
    const totalMembersInPart = memberList.length;

    // Write text labels to left column (targetCol)
    targetSheet
      .getRange(footerStartRow, targetCol)
      .setValue("出席")
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontColor("#000000")
      .setHorizontalAlignment("left");

    targetSheet
      .getRange(footerStartRow + 1, targetCol)
      .setValue("總數")
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontColor("#000000")
      .setHorizontalAlignment("left");

    targetSheet
      .getRange(footerStartRow + 2, targetCol)
      .setValue("出席率")
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontColor("#000000")
      .setHorizontalAlignment("left");

    // Write numbers to right column (targetCol + 1)
    const leftColLetter = columnToLetter(targetCol + 1);

    // Write attendance count
    targetSheet
      .getRange(footerStartRow, targetCol + 1)
      .setValue(colPresentCount)
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontColor("#000000")
      .setHorizontalAlignment("center");

    // Write total count
    targetSheet
      .getRange(footerStartRow + 1, targetCol + 1)
      .setValue(totalMembersInPart)
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontColor("#000000")
      .setHorizontalAlignment("center");

    // Write attendance rate with formula
    targetSheet
      .getRange(footerStartRow + 2, targetCol + 1)
      .setFormula(
        `=ROUND(${leftColLetter}${footerStartRow}/${leftColLetter}${
          footerStartRow + 1
        }*100,0)&"%"`
      )
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontColor("#000000")
      .setHorizontalAlignment("center");

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

  // --- 6. Generate Missing Members List (Right Side) ---
  generateMissingMembersList(
    targetSheet,
    columns,
    allMembersByPart,
    checkedInNames,
    headerRow,
    startRow,
    footerStartRow
  );

  // Toast notification
  SpreadsheetApp.getActive().toast(`已更新 ${targetDateStr} 的點名表`, "完成");
}

/**
 * Calculate attendance data by date (extracted logic from generateAttendanceByDate)
 * Returns data in the same format as collectAttendanceData for use in exportAttendanceImage
 * @param {string} targetDateStr - Date string in "yyyy/MM/dd" format
 * @returns {Object} Attendance data with parts, globalStat, missingMembers, totalMissing
 */
function calculateAttendanceDataByDate(targetDateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(CONFIG.sourceSheetName);
  const memberSheet = ss.getSheetByName(CONFIG.memberSheetName);

  if (!sourceSheet) {
    throw new Error("Source sheet not found!");
  }

  const timezone = Session.getScriptTimeZone();

  // --- 1. Read and Filter Data from Records ---
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

    const name = normalizeName(row[1]);
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

  // --- 2. Read All Members from name_list ---
  const allMembersByPart = {}; // { "T1": Set(["name1", "name2"]), "T2": [...], ... }
  const allMembersSet = new Set(); // For quick lookup

  if (memberSheet) {
    const memberDataRows = memberSheet.getDataRange().getValues();
    // Start from row 2 (skip header)
    for (let i = 1; i < memberDataRows.length; i++) {
      const row = memberDataRows[i];
      const part = row[0];
      const name = normalizeName(row[1]);
      if (part && name) {
        if (!allMembersByPart[part]) {
          allMembersByPart[part] = new Set();
        }
        allMembersByPart[part].add(name);
        allMembersSet.add(name);
      }
    }
  }

  // --- 3. Distribute Members to Columns ---
  const columns = [[], [], [], [], []];

  // Add checked-in members
  Object.values(memberData).forEach((member) => {
    let colIndex = CONFIG.partMapping[member.part];
    if (colIndex === undefined) colIndex = CONFIG.defaultColumnIndex;
    if (colIndex !== null && columns[colIndex]) {
      columns[colIndex].push(member);
    }
  });

  // Add missing members (not checked in) to columns
  const checkedInNames = new Set(Object.keys(memberData));
  CONFIG.partHeaders.forEach((partName) => {
    const colIndex = CONFIG.partMapping[partName];
    if (colIndex !== undefined && colIndex !== null && columns[colIndex]) {
      const allMembersInPart = allMembersByPart[partName]
        ? Array.from(allMembersByPart[partName])
        : [];
      allMembersInPart.forEach((name) => {
        if (!checkedInNames.has(name)) {
          // Add missing member with null status
          columns[colIndex].push({
            name: name,
            part: partName,
            status: null,
            timestamp: null,
          });
        }
      });
    }
  });

  // --- 4. Sort Members: "出席" first, then "晚到", then "請假", then others ---
  columns.forEach((colList) => {
    colList.sort((a, b) => {
      // Define status priority: 出席 (1) > 晚到 (2) > 請假 (3) > others (4)
      const getStatusPriority = (status) => {
        if (status === "出席") return 1;
        if (status === "晚到") return 2;
        if (status === "請假") return 3;
        return 4;
      };

      const priorityA = getStatusPriority(a.status);
      const priorityB = getStatusPriority(b.status);
      if (priorityA !== priorityB) {
        return priorityA - priorityB;
      }

      // Same status priority, then sort by timestamp (checked-in members), missing members (null timestamp) go last
      if (a.timestamp === null && b.timestamp !== null) return 1;
      if (a.timestamp !== null && b.timestamp === null) return -1;
      if (a.timestamp !== null && b.timestamp !== null) {
        return a.timestamp - b.timestamp;
      }
      // Both are missing, sort by name
      return a.name.localeCompare(b.name);
    });
  });

  // --- 5. Calculate Statistics and Convert to Output Format ---
  let globalTotalCount = allMembersSet.size;
  let globalPresentCount = 0;

  const result = {
    parts: [],
    globalStat: "",
    missingMembers: {},
    totalMissing: 0,
  };

  // Process each part
  CONFIG.partHeaders.forEach((partName, colIndex) => {
    const memberList = columns[colIndex];
    const partData = {
      name: partName,
      members: [],
      presentCount: 0,
      totalCount: memberList.length,
    };

    let colPresentCount = 0;

    // Convert members to output format with status symbols
    memberList.forEach((member) => {
      let statusSymbol = "";

      if (member.status === null) {
        // Missing member (not checked in) - empty status
        statusSymbol = "";
      } else {
        const config = CONFIG.statusMap[member.status];
        if (config) {
          statusSymbol = config.symbol;

          // Count both "出席" (V) and "晚到" (L) as present
          if (config.symbol === "V" || config.symbol === "L") {
            colPresentCount++;
            globalPresentCount++;
          }
        } else {
          // Unknown status, use original status text
          statusSymbol = member.status;
        }
      }

      partData.members.push({
        name: member.name,
        status: statusSymbol,
      });
    });

    partData.presentCount = colPresentCount;
    result.parts.push(partData);
  });

  // Calculate global statistics
  result.globalStat = `總出席: ${globalPresentCount}/${globalTotalCount}`;

  // Calculate missing members for each part (excluding Conductor)
  const missingParts = ["T1", "T2", "B1", "B2"];
  missingParts.forEach((partName) => {
    const allMembersInPart = allMembersByPart[partName]
      ? Array.from(allMembersByPart[partName])
      : [];
    const missing = allMembersInPart.filter(
      (name) => !checkedInNames.has(name)
    );
    result.missingMembers[partName] = missing;
    result.totalMissing += missing.length;
  });

  return result;
}

/**
 * Generate Missing Members List (Displayed on the right side of the attendance sheet)
 */
function generateMissingMembersList(
  targetSheet,
  columns,
  allMembersByPart,
  checkedInNames,
  headerRow,
  startRow,
  footerStartRow
) {
  // Starting column for missing members list (Column M = 13)
  const missingStartCol = 13;

  // Clear the missing members area first
  targetSheet
    .getRange("M3:R1000")
    .clearContent()
    .setBackground(null)
    .setBorder(false, false, false, false, false, false);

  // Add title for missing members section
  const titleCell = targetSheet.getRange("M3:N3");
  titleCell
    .merge()
    .setValue("未打卡名單")
    .setFontWeight("bold")
    .setFontSize(14)
    .setFontColor("#c90000")
    .setBackground("#ffcccc")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, null, null);

  // Parts to display (excluding Conductor for simplicity, can be added if needed)
  const partsToShow = ["T1", "T2", "B1", "B2"];

  // Calculate missing members for each part
  const missingByPart = {};
  partsToShow.forEach((partName) => {
    const allMembersInPart = allMembersByPart[partName]
      ? Array.from(allMembersByPart[partName])
      : [];
    missingByPart[partName] = allMembersInPart.filter(
      (name) => !checkedInNames.has(name)
    );
  });

  // Find the maximum number of missing members across all parts
  const maxMissing = Math.max(
    ...partsToShow.map((p) => missingByPart[p].length),
    0
  );

  // If no missing members, show a message
  if (maxMissing === 0) {
    const noMissingCell = targetSheet.getRange("M5");
    noMissingCell
      .setValue("全員已打卡！")
      .setFontWeight("bold")
      .setFontSize(12)
      .setFontColor("#008000")
      .setHorizontalAlignment("center");
    return;
  }

  // Column positions for missing members (M=13, N=14, O=15, P=16 for T1, T2, B1, B2)
  const missingColPositions = [13, 14, 15, 16];

  // Write headers for each part
  partsToShow.forEach((partName, index) => {
    const col = missingColPositions[index];
    const headerCell = targetSheet.getRange(headerRow, col);
    const missingCount = missingByPart[partName].length;

    headerCell
      .setValue(`${partName} (${missingCount})`)
      .setFontWeight("bold")
      .setFontSize(11)
      .setBackground("#ffeeee")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, null, null);

    // Write missing member names
    missingByPart[partName].forEach((name, rowIndex) => {
      const currentRow = startRow + rowIndex;
      const nameCell = targetSheet.getRange(currentRow, col);

      nameCell
        .setValue(name)
        .setFontSize(11)
        .setFontColor("#666666")
        .setHorizontalAlignment("center");
    });
  });

  // Add border around the entire missing members section
  const missingFooterRow = startRow + maxMissing;
  const entireMissingRange = targetSheet.getRange(
    headerRow,
    missingColPositions[0],
    missingFooterRow - headerRow + 1,
    partsToShow.length
  );
  entireMissingRange.setBorder(true, true, true, true, true, true);

  // Add total missing count at the bottom
  const totalMissing = partsToShow.reduce(
    (sum, p) => sum + missingByPart[p].length,
    0
  );
  const totalRow = missingFooterRow + 1;
  const totalRange = targetSheet.getRange(
    totalRow,
    missingColPositions[0],
    1,
    partsToShow.length
  );
  totalRange.merge();

  targetSheet
    .getRange(totalRow, missingColPositions[0])
    .setValue(`未打卡總計: ${totalMissing} 人`)
    .setFontWeight("bold")
    .setFontSize(11)
    .setFontColor("#c90000")
    .setBackground("#fff2cc")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, null, null);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("點名系統")
    .addItem("更新今日/指定日期點名", "generateAttendanceByDate")
    .addItem("匯出點名表圖片", "exportAttendanceImage")
    .addToUi();
}

/**
 * Export attendance table as image using HTML rendering
 * Generates HTML table and converts to image using html2canvas
 */
function exportAttendanceImage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(CONFIG.targetSheetName);

  if (!targetSheet) {
    SpreadsheetApp.getUi().alert("找不到點名表！");
    return;
  }

  // Get date from the sheet
  const dateValue = targetSheet.getRange(CONFIG.dateCell).getValue();
  const timezone = Session.getScriptTimeZone();
  let dateStr = "";
  if (dateValue instanceof Date) {
    dateStr = Utilities.formatDate(dateValue, timezone, "yyyy/MM/dd");
  } else {
    dateStr = Utilities.formatDate(new Date(), timezone, "yyyy/MM/dd");
  }

  // Calculate attendance data using the same logic as generateAttendanceByDate
  let attendanceData;
  try {
    attendanceData = calculateAttendanceDataByDate(dateStr);
  } catch (error) {
    SpreadsheetApp.getUi().alert("計算出席資料時發生錯誤: " + error.message);
    return;
  }

  // Generate HTML and show dialog
  const htmlContent = generateAttendanceHtml(attendanceData, dateStr);

  const html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(950)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(html, "匯出點名表圖片");
}

/**
 * Generate HTML for attendance table with download capability
 * Styled for G-Major Choir with musical theme
 */
function generateAttendanceHtml(data, dateStr) {
  // Build table rows HTML
  let tableRowsHtml = "";

  // Find max members
  const maxMembers = Math.max(...data.parts.map((p) => p.members.length), 1);

  // Member rows
  for (let i = 0; i < maxMembers; i++) {
    let rowHtml = "<tr>";
    data.parts.forEach((part) => {
      const member = part.members[i];
      if (!member) {
        rowHtml += '<td class="name empty"></td><td class="status empty"></td>';
      } else {
        let bgClass = "";
        if (member.status === "V") bgClass = "present";
        else if (member.status === "△") bgClass = "absent";
        else if (member.status === "L") bgClass = "late";

        rowHtml += `<td class="name ${bgClass}">${member.name}</td>`;
        rowHtml += `<td class="status ${bgClass}">${member.status || ""}</td>`;
      }
    });
    rowHtml += "</tr>";
    tableRowsHtml += rowHtml;
  }

  // Stats rows
  const statsConfig = [
    { label: "出席", getValue: (p) => p.presentCount },
    { label: "總數", getValue: (p) => p.totalCount },
    {
      label: "出席率",
      getValue: (p) =>
        p.totalCount > 0
          ? Math.round((p.presentCount / p.totalCount) * 100) + "%"
          : "0%",
    },
  ];

  statsConfig.forEach((stat) => {
    let rowHtml = '<tr class="stats-row">';
    data.parts.forEach((part) => {
      rowHtml += `<td class="stat-label">${stat.label}</td>`;
      rowHtml += `<td class="stat-value">${stat.getValue(part)}</td>`;
    });
    rowHtml += "</tr>";
    tableRowsHtml += rowHtml;
  });

  // Header cells with musical icons
  const partIcons = {
    Conductor: "🎼",
    T1: "🎵",
    T2: "🎶",
    B1: "🎵",
    B2: "🎶"
  };

  let headerHtml = "";
  data.parts.forEach((part) => {
    const icon = partIcons[part.name] || "♪";
    headerHtml += `<th colspan="2">${icon} ${part.name}</th>`;
  });

  // Calculate attendance rate for color indicator
  const totalPresent = data.parts.reduce((sum, p) => sum + p.presentCount, 0);
  const totalMembers = data.parts.reduce((sum, p) => sum + p.totalCount, 0);
  const overallRate = totalMembers > 0 ? Math.round((totalPresent / totalMembers) * 100) : 0;

  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <script src="https://html2canvas.hertzen.com/dist/html2canvas.min.js"></script>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&display=swap');

    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      font-family: 'Noto Sans TC', -apple-system, BlinkMacSystemFont, sans-serif;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      padding: 15px;
      min-height: 100vh;
    }
    .controls {
      text-align: center;
      margin-bottom: 15px;
    }
    .btn {
      padding: 12px 28px;
      font-size: 14px;
      border: none;
      border-radius: 25px;
      cursor: pointer;
      margin: 0 8px;
      font-weight: 600;
      transition: all 0.3s ease;
      box-shadow: 0 4px 15px rgba(0,0,0,0.2);
    }
    .btn-download {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
    }
    .btn-download:hover {
      transform: translateY(-2px);
      box-shadow: 0 6px 20px rgba(102,126,234,0.4);
    }
    .btn-copy {
      background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
      color: white;
    }
    .btn-copy:hover {
      transform: translateY(-2px);
      box-shadow: 0 6px 20px rgba(56,239,125,0.4);
    }
    .success-msg {
      color: #38ef7d;
      font-weight: bold;
      margin-top: 10px;
      display: none;
      text-shadow: 0 0 10px rgba(56,239,125,0.5);
    }

    #capture {
      background: linear-gradient(180deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
      padding: 25px;
      display: inline-block;
      border-radius: 20px;
      box-shadow: 0 20px 60px rgba(0,0,0,0.3);
    }

    .header-section {
      text-align: center;
      margin-bottom: 20px;
    }
    .logo-title {
      font-size: 28px;
      font-weight: 700;
      color: #fff;
      text-shadow: 0 0 20px rgba(255,215,0,0.5);
      margin-bottom: 8px;
      letter-spacing: 2px;
    }
    .logo-title .g { color: #ffd700; }
    .logo-title .major { color: #fff; }
    .music-notes {
      font-size: 20px;
      margin: 0 10px;
      animation: float 2s ease-in-out infinite;
    }
    @keyframes float {
      0%, 100% { transform: translateY(0); }
      50% { transform: translateY(-5px); }
    }

    .date-badge {
      display: inline-block;
      background: linear-gradient(135deg, #ffd700 0%, #ffb700 100%);
      padding: 8px 25px;
      border-radius: 20px;
      font-size: 16px;
      font-weight: 700;
      color: #1a1a2e;
      margin-bottom: 12px;
      box-shadow: 0 4px 15px rgba(255,215,0,0.3);
    }

    .stats-banner {
      display: inline-flex;
      align-items: center;
      gap: 15px;
      background: rgba(255,255,255,0.1);
      backdrop-filter: blur(10px);
      padding: 10px 25px;
      border-radius: 15px;
      border: 1px solid rgba(255,255,255,0.2);
    }
    .stats-banner .label {
      color: rgba(255,255,255,0.8);
      font-size: 14px;
    }
    .stats-banner .value {
      font-size: 22px;
      font-weight: 700;
      color: #ffd700;
      text-shadow: 0 0 10px rgba(255,215,0,0.5);
    }
    .stats-banner .rate {
      background: ${overallRate >= 70 ? 'linear-gradient(135deg, #11998e 0%, #38ef7d 100%)' : overallRate >= 50 ? 'linear-gradient(135deg, #f093fb 0%, #f5576c 100%)' : 'linear-gradient(135deg, #eb3349 0%, #f45c43 100%)'};
      padding: 5px 15px;
      border-radius: 10px;
      color: white;
      font-weight: 700;
      font-size: 14px;
    }

    table {
      border-collapse: separate;
      border-spacing: 2px;
      font-size: 12px;
      margin-top: 15px;
    }
    th {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      font-weight: 600;
      font-size: 13px;
      padding: 10px 12px;
      border-radius: 8px 8px 0 0;
      text-shadow: 0 1px 2px rgba(0,0,0,0.2);
    }
    td {
      padding: 6px 10px;
      text-align: center;
      background: rgba(255,255,255,0.95);
      transition: all 0.2s ease;
    }
    tr:hover td:not(.empty) {
      transform: scale(1.02);
    }
    .name {
      font-weight: 600;
      min-width: 75px;
      border-radius: 5px 0 0 5px;
    }
    .status {
      min-width: 28px;
      font-weight: 700;
      border-radius: 0 5px 5px 0;
    }
    .empty {
      background: rgba(255,255,255,0.3);
    }

    .present {
      background: linear-gradient(135deg, #fff9c4 0%, #ffe082 100%);
      color: #5d4037;
    }
    .present.status { color: #2e7d32; }

    .absent {
      background: linear-gradient(135deg, #ffcdd2 0%, #ef9a9a 100%);
      color: #b71c1c;
    }

    .late {
      background: linear-gradient(135deg, #bbdefb 0%, #90caf9 100%);
      color: #1565c0;
    }

    .stats-row td {
      background: linear-gradient(135deg, #e8eaf6 0%, #c5cae9 100%);
      font-weight: 600;
      color: #3f51b5;
    }
    .stat-label { font-size: 11px; }
    .stat-value { font-size: 13px; font-weight: 700; }

    .footer-note {
      text-align: center;
      margin-top: 15px;
      color: rgba(255,255,255,0.6);
      font-size: 11px;
    }

    .preview-container {
      text-align: center;
      overflow: auto;
      max-height: 600px;
      padding: 10px;
    }
    .hint {
      text-align: center;
      color: white;
      font-size: 12px;
      margin-top: 12px;
      opacity: 0.8;
    }
  </style>
</head>
<body>
  <div class="controls">
    <button class="btn btn-download" onclick="downloadImage()">⬇️ 下載圖片</button>
    <button class="btn btn-copy" onclick="copyImage()">📋 複製圖片</button>
    <div class="success-msg" id="successMsg">✓ 已複製到剪貼簿！</div>
  </div>

  <div class="preview-container">
    <div id="capture">
      <div class="header-section">
        <div class="logo-title">
          <span class="music-notes">♪</span>
          <span class="g">G</span><span class="major">Major</span> 點名表
          <span class="music-notes">♫</span>
        </div>
        <div class="date-badge">📅 ${dateStr}</div>
        <br>
        <div class="stats-banner">
          <span class="label">出席人數</span>
          <span class="value">${totalPresent}/${totalMembers}</span>
          <span class="rate">${overallRate}%</span>
        </div>
      </div>

      <table>
        <thead>
          <tr>${headerHtml}</tr>
        </thead>
        <tbody>
          ${tableRowsHtml}
        </tbody>
      </table>

      <div class="footer-note">
        🎹 G-Major Choir • Let's sing together! 🎤
      </div>
    </div>
  </div>

  <div class="hint">提示：下載或複製後可直接貼到 FB 發文</div>

  <script>
    async function downloadImage() {
      const element = document.getElementById('capture');
      const canvas = await html2canvas(element, {
        scale: 2,
        backgroundColor: null,
        useCORS: true
      });
      const link = document.createElement('a');
      link.download = 'G-Major_點名表_${dateStr}.png';
      link.href = canvas.toDataURL('image/png');
      link.click();
    }

    async function copyImage() {
      try {
        const element = document.getElementById('capture');
        const canvas = await html2canvas(element, {
          scale: 2,
          backgroundColor: null,
          useCORS: true
        });
        canvas.toBlob(async (blob) => {
          try {
            await navigator.clipboard.write([
              new ClipboardItem({ 'image/png': blob })
            ]);
            const msg = document.getElementById('successMsg');
            msg.style.display = 'block';
            setTimeout(() => { msg.style.display = 'none'; }, 2000);
          } catch (err) {
            alert('複製失敗，請改用下載按鈕');
          }
        }, 'image/png');
      } catch (err) {
        alert('複製失敗，請改用下載按鈕');
      }
    }
  </script>
</body>
</html>
  `;
}

/**
 * Collect attendance data from the target sheet
 */
function collectAttendanceData(sheet) {
  const headerRow = 5;
  const startRow = 6;
  const colPositions = [2, 4, 6, 8, 10]; // B, D, F, H, J

  const result = {
    parts: [],
    globalStat: "",
    missingMembers: {},
    totalMissing: 0,
  };

  // Get global stat
  result.globalStat = sheet.getRange("H3").getValue() || "總出席: 0/0";

  // First pass: collect all members to find max count
  const allPartsData = [];
  let globalMaxMembers = 0;

  CONFIG.partHeaders.forEach((partName, colIndex) => {
    const targetCol = colPositions[colIndex];
    const partData = {
      name: partName,
      members: [],
      presentCount: 0,
      totalCount: 0,
    };

    // Read members (start from row 6, read until empty or hit footer)
    let row = startRow;
    while (row < 100) {
      const name = sheet.getRange(row, targetCol).getValue();
      const cellText = String(name).trim();

      // Stop if we hit the footer (starts with "出席")
      if (cellText === "出席" || cellText === "") {
        // Check if it's really empty or just this column
        const nextColVal = sheet.getRange(row, targetCol + 1).getValue();
        if (cellText === "出席" || (cellText === "" && !nextColVal)) {
          break;
        }
      }

      if (cellText === "") break;

      const status = sheet.getRange(row, targetCol + 1).getValue();

      partData.members.push({
        name: cellText,
        status: status,
      });
      row++;
    }

    globalMaxMembers = Math.max(globalMaxMembers, partData.members.length);
    allPartsData.push(partData);
  });

  // Second pass: read statistics from unified footer position
  const footerStartRow = startRow + globalMaxMembers;

  allPartsData.forEach((partData, colIndex) => {
    const targetCol = colPositions[colIndex];

    // Read stats from footer (出席 row has value in second column)
    const presentVal = sheet.getRange(footerStartRow, targetCol + 1).getValue();
    const totalVal = sheet.getRange(footerStartRow + 1, targetCol + 1).getValue();

    partData.presentCount = typeof presentVal === "number" ? presentVal : 0;
    partData.totalCount = typeof totalVal === "number" ? totalVal : partData.members.length;

    result.parts.push(partData);
  });

  // Get missing members (columns M, N, O, P)
  const missingParts = ["T1", "T2", "B1", "B2"];
  const missingCols = [13, 14, 15, 16];

  missingParts.forEach((part, index) => {
    const col = missingCols[index];
    const members = [];
    let row = startRow;

    while (row < 100) {
      const name = sheet.getRange(row, col).getValue();
      if (!name) break;
      members.push(name);
      row++;
    }

    result.missingMembers[part] = members;
    result.totalMissing += members.length;
  });

  return result;
}
