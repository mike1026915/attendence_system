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
    請假: { symbol: "△", color: "#f4c7c3" }, // Red
    晚到: { symbol: "L", color: "#cfe2f3" }, // Blue
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

  // --- 3.1 Sort Members: "出席" first, then checked-in, then missing ---
  columns.forEach((colList) => {
    colList.sort((a, b) => {
      // Present members first
      const isAPresent = a.status === "出席";
      const isBPresent = b.status === "出席";
      if (isAPresent && !isBPresent) return -1;
      if (!isAPresent && isBPresent) return 1;
      // Then sort by timestamp (checked-in members), missing members (null timestamp) go last
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

          if (config.symbol === "V") {
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
    .addToUi();
}
