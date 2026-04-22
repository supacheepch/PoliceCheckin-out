function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const loadType = (e.parameter.load || "all").toLowerCase();
    const result = {};

    // Helper: ดึงข้อมูลจากชีตแปลงเป็น JSON Array
    const fetchSheetData = (name) => {
      const sheet = ss.getSheetByName(name);
      if (!sheet) return [];
      const data = sheet.getDataRange().getDisplayValues();
      if (data.length <= 1) return [];
      const headers = data[0];
      return data.slice(1).map(row => {
        const obj = {};
        headers.forEach((h, i) => { if (h) obj[h] = row[i]; });
        return obj;
      });
    };

    // 🎯 เลือกดึงเฉพาะที่จำเป็นตามหน้าที่เรียกมา (Performance Optimization)
    if (loadType === "leaderboard") {
      result.police_users = fetchSheetData("POLICE_USER");
      result.reward_master = fetchSheetData("REWARD_MASTER");
      result.wall_of_fame = fetchSheetData("WALL_OF_FAME");
    }
    else if (loadType === "main") {
      result.police_users = fetchSheetData("POLICE_USER");
    }
    else {
      // Default: ดึงทุกอย่าง (สำหรับ Manager หรือกรณีไม่ได้ระบุ)
      result.police_users = fetchSheetData("POLICE_USER");
      result.week_master = fetchSheetData("WEEK_MASTER");
      result.rank_police = fetchSheetData("RANK_POLICE");
      result.reward_master = fetchSheetData("REWARD_MASTER");
      result.wall_of_fame = fetchSheetData("WALL_OF_FAME");
    }

    return json({ status: "success", data: result });
  } catch (err) {
    return json({ status: "error", message: err.message });
  }
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  const action = data.action; // 👈 ตัวแยก API

  if (action === "register") {
    return register(data);
  }

  if (action === "checkin") {
    return checkInOut(data);
  }

  if (action === "summary") {
    return actionSummaryTrigger(data.sheet);
  }

  if (action === "summary_reward") {
    return actionSummaryReward(data);
  }

  if (action === "updateUser") {
    return updateUser(data);
  }

  if (action === "getLog") {
    return fetchUserLogs(data);
  }

  if (action === "manageRank") {
    return manageRank(data);
  }

  return json({ status: "error", message: "Invalid action" });
}

// ============================================
// ดึง Log การลงเวลาของ User จากชีตสัปดาห์ปัจจุบัน (Helper)
// ============================================
function getUserLogsList(policeCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = getCurrentWeekSheetName();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const logsData = sheet.getDataRange().getDisplayValues();
  if (logsData.length <= 1) return [];

  const userLogs = [];
  for (let i = 1; i < logsData.length; i++) {
    if (logsData[i][1] == policeCode) {
      userLogs.push({
        type: logsData[i][3].toLowerCase(),
        date: logsData[i][4],
        time: logsData[i][5]
      });
    }
  }
  return userLogs.reverse();
}

function fetchUserLogs(data) {
  try {
    const userLogs = getUserLogsList(data.POLICE_CODE);

    // ดึงสถานะการเข้าเวรล่าสุดของคนๆ นี้จากหน้า POLICE_USER ให้ด้วย
    // เผื่อในกรณีที่ Session ในเครื่องข้ามวันแล้วสถานะค้าง
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName("POLICE_USER");
    let currentStatus = "N";
    if (userSheet) {
      const uData = userSheet.getDataRange().getDisplayValues();
      for (let r = 1; r < uData.length; r++) {
        if (uData[r][0] == data.POLICE_CODE) {
          currentStatus = uData[r][3]; // คอลัมน์ ON_DULTY
          break;
        }
      }
    }

    return json({
      status: "success",
      logs: userLogs,
      on_duty: currentStatus
    });

  } catch (e) {
    return json({ status: "error", message: e.toString() });
  }
}

function register(e) {
  try {
    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("POLICE_USER");

    const data = e;
    if (!data.policeCode) {
      return json({ status: "error", message: "Missing policeCode" });
    }

    const rows = sheet.getDataRange().getValues();

    // 🔍 เช็คซ้ำ
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] == data.policeCode) {
        return json({ status: "duplicate", message: "Police code already exists" });
      }
      if (rows[i][2] == data.policeEmail) {
        return json({ status: "duplicate", message: "Email already exists" });
      }
    }

    // ✅ insert
    sheet.appendRow([
      data.policeCode,
      data.policeName,
      data.policeEmail,
      "N", // ON_DULTY
      "N", // ACTIVE / STATUS (ต้องรอแอดมินอนุมัติ)
      "นักเรียนตำรวจฝึกหัด",
      "00:00",
      new Date(),
    ]);

    return json({ status: "success", message: "ลงทะเบียนเรียบร้อยแล้ว" });

  } catch (err) {
    return json({ status: "error", message: err.message });
  }
}


function checkInOut(e) {
  try {
    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("POLICE_USER");

    const data = e;

    if (!data.POLICE_CODE || !data.checkType) {
      return json({ status: "error", message: "Missing required fields" });
    }

    const rows = sheet.getDataRange().getValues();
    let foundRow = -1;

    // 🔍 หา row
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] == data.POLICE_CODE) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow === -1) {
      return json({ status: "error", message: "Police not found" });
    }

    const type = data.checkType.toUpperCase();

    // 👉 mapping column
    const ON_DULTY = 4; // D

    if (type === "IN") {
      sheet.getRange(foundRow, ON_DULTY).setValue("Y");
      data.STATUS = "IN";
      logWeekly(data);
      logGlobal(data);
    } else if (type === "OUT") {
      sheet.getRange(foundRow, ON_DULTY).setValue("N");
      data.STATUS = "OUT";
      logWeekly(data);
      logGlobal(data);
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const weekSheetName = getCurrentWeekSheetName();
      const weeklySheet = ss.getSheetByName(weekSheetName);

      const lastIn = getLastCheckIn(weeklySheet, data.POLICE_CODE);

      if (lastIn) {
        const now = new Date();
        const hours = calcHours(lastIn, now);

        updateWorkHours(sheet, foundRow, hours);
      }
    } else {
      return json({ status: "error", message: "Invalid checkType" });
    }
    SpreadsheetApp.flush();
    // 🔥 ดึงข้อมูล row ล่าสุด
    const updatedRow = sheet.getRange(foundRow, 1, 1, 10).getValues()[0];

    // 👉 map เป็น object (ตาม column จริงของคุณ)
    const user = {
      POLICE_CODE: updatedRow[0],
      POLICE_NAME: updatedRow[1],
      POLICE_EMAIL: updatedRow[2],
      ON_DULTY: updatedRow[3],
      ACTIVE: updatedRow[4],
      ROLE: updatedRow[5],
      WORK_TIME: updatedRow[6],
      UPDATED_AT: updatedRow[7],
      TMN_DT: updatedRow[8],
      PIC_URL: updatedRow[9] || '',
    };

    // 🔥 ดึง Log ใหม่แบบ Internal (ไม่ต้องยิง API ซ้ำ)
    const updatedLogs = getUserLogsList(data.POLICE_CODE);

    return json({
      status: "success",
      message: type === "IN" ? "checked_in" : "checked_out",
      data: user,
      logs: updatedLogs
    });

  } catch (err) {
    return json({ status: "error", message: err.message });
  }
}


function logWeekly(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();

  const { sunday, saturday } = getWeekRange(now);

  const start = Utilities.formatDate(sunday, "Asia/Bangkok", "yyyy-MM-dd");
  const end = Utilities.formatDate(saturday, "Asia/Bangkok", "yyyy-MM-dd");

  const sheetName = `${start}_to_${end}`;

  let sheet = ss.getSheetByName(sheetName);

  // ❗ ถ้าไม่มี → สร้างใหม่
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    ss.moveActiveSheet(ss.getNumSheets());

    sheet.appendRow([
      "INDEX",
      "POLICE_CODE",
      "POLICE_NAME",
      "STATUS",
      "DATE",
      "TIME"
    ]);
    sheet.setFrozenRows(1);
    // 🔥 เพิ่มเข้า WEEK_MASTER
    let master = ss.getSheetByName("WEEK_MASTER");

    if (!master) {
      master = ss.insertSheet("WEEK_MASTER");
      master.appendRow(["ID", "SHEET_NAME"]);
    }

    const masterRows = master.getDataRange().getValues();

    // 👉 กันซ้ำ sheet name
    const exists = masterRows.some(r => r[1] === sheetName);

    if (!exists) {
      const id = master.getLastRow(); // simple index
      master.appendRow([id, sheetName]);
    }
  }

  // 👉 format date + time แยก
  const dateStr = Utilities.formatDate(now, "Asia/Bangkok", "yyyy-MM-dd");
  const timeStr = Utilities.formatDate(now, "Asia/Bangkok", "HH:mm:ss");
  const index = sheet.getLastRow();
  // 👉 append log
  sheet.appendRow([
    index,
    data.POLICE_CODE,
    data.POLICE_NAME,
    data.STATUS,
    dateStr,
    timeStr
  ]);
}

function logGlobal(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "GLOBAL_LOGS";
  let sheet = ss.getSheetByName(sheetName);

  // ❗ ถ้าไม่มี → สร้างใหม่
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    ss.moveActiveSheet(ss.getNumSheets());

    sheet.appendRow([
      "INDEX",
      "POLICE_CODE",
      "POLICE_NAME",
      "STATUS",
      "DATE",
      "TIME"
    ]);
    sheet.setFrozenRows(1);
  }

  const now = new Date();
  const dateStr = Utilities.formatDate(now, "Asia/Bangkok", "yyyy-MM-dd");
  const timeStr = Utilities.formatDate(now, "Asia/Bangkok", "HH:mm:ss");
  const index = sheet.getLastRow();

  sheet.appendRow([
    index,
    data.POLICE_CODE,
    data.POLICE_NAME,
    data.STATUS,
    dateStr,
    timeStr
  ]);
}

function getWeekRange(date) {
  const d = new Date(date);

  const day = d.getDay(); // 0 = Sunday
  const sunday = new Date(d);
  sunday.setDate(d.getDate() - day);

  const saturday = new Date(sunday);
  saturday.setDate(sunday.getDate() + 6);

  return { sunday, saturday };
}

function parseSheetDateTime(d, t) {
  if (!d || !t) return null;

  let yyyy, MM, dd;
  if (d instanceof Date) {
    if (isNaN(d.getTime())) return null;
    yyyy = d.getFullYear();
    MM = String(d.getMonth() + 1).padStart(2, '0');
    dd = String(d.getDate()).padStart(2, '0');
  } else {
    const s = String(d).trim();
    if (s.includes('-')) {
      const p = s.split('-');
      if (p.length !== 3) return null;
      if (p[0].length === 4) { yyyy = p[0]; MM = p[1]; dd = p[2]; }
      else { yyyy = p[2]; MM = p[1]; dd = p[0]; }
    } else if (s.includes('/')) {
      const p = s.split('/');
      if (p.length !== 3) return null;
      const tmp = new Date(s);
      if (!isNaN(tmp.getTime()) && tmp.getFullYear() > 2000) {
        yyyy = tmp.getFullYear(); MM = String(tmp.getMonth() + 1).padStart(2, '0'); dd = String(tmp.getDate()).padStart(2, '0');
      } else {
        dd = p[0]; MM = p[1]; yyyy = p[2]; // fallback dd/MM/yyyy
      }
    } else {
      return null;
    }
  }

  let hh, mm, ss;
  if (t instanceof Date) {
    if (isNaN(t.getTime())) return null;
    hh = String(t.getHours()).padStart(2, '0');
    mm = String(t.getMinutes()).padStart(2, '0');
    ss = String(t.getSeconds()).padStart(2, '0');
  } else {
    const s = String(t).trim();
    const parts = s.split(':');
    if (parts.length < 2) return null;
    hh = String(parts[0]).padStart(2, '0');
    mm = String(parts[1]).padStart(2, '0');
    ss = String(parts[2] || '00').padStart(2, '0');
  }

  const dt = new Date(`${yyyy}-${MM}-${dd}T${hh}:${mm}:${ss}+07:00`);
  if (isNaN(dt.getTime())) return null;

  return dt;
}

function getLastCheckIn(sheet, policeCode) {
  // ใช้ getValues() ร่วมกับ parseSheetDateTime ดีกว่า เผื่อ Sheet เผลอแปลงเป็น Date Object
  const rows = sheet.getDataRange().getValues();

  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][1] == policeCode && rows[i][3] == "IN") {
      const dVal = rows[i][4];
      const tVal = rows[i][5];

      // 🔥 ฟังก์ชันนี้จะแปลงได้ทั้ง Object Date และ String "2026-04-21" แน่นอน
      const dt = parseSheetDateTime(dVal, tVal);

      return dt;
    }
  }

  return null;
}

function calcHours(inTime, outTime) {
  // ตรวจสอบว่าเป็น Object ชนิด Date ที่ถูกต้องหรือไม่
  if (!(inTime instanceof Date) || isNaN(inTime.getTime())) return 0;
  if (!(outTime instanceof Date) || isNaN(outTime.getTime())) return 0;

  const diffMs = outTime.getTime() - inTime.getTime();
  if (diffMs < 0) return 0;

  return diffMs / (1000 * 60 * 60);
}

function updateWorkHours(sheetUser, foundRow, addedHours) {
  const HOURS_COL = 7;

  // 🔴 เปลี่ยนเป็น getDisplayValue()
  const raw = sheetUser.getRange(foundRow, HOURS_COL).getDisplayValue();

  // 👉 แปลงของเดิมเป็นนาที
  const currentMinutes = timeToMinutes(raw);

  // 👉 ชั่วโมง → นาที (ใช้ Math.floor เพื่อให้ตรงกับหน้า Summary)
  const addedMinutes = Math.floor(addedHours * 60);

  const totalMinutes = currentMinutes + addedMinutes;

  // 👉 แปลงกลับเป็น HH:mm
  const formatted = minutesToTime(totalMinutes);

  sheetUser.getRange(foundRow, HOURS_COL).setValue(formatted);
}

function getCurrentWeekSheetName() {
  const now = new Date();
  const { sunday, saturday } = getWeekRange(now);

  const start = Utilities.formatDate(sunday, "Asia/Bangkok", "yyyy-MM-dd");
  const end = Utilities.formatDate(saturday, "Asia/Bangkok", "yyyy-MM-dd");

  return `${start}_to_${end}`;
}

function timeToMinutes(str) {
  // 🔴 ดักจับถ้าค่าว่าง หรือชีตพังเป็น NaN:NaN ไปแล้วให้เริ่มที่ 0 ใหม่
  if (!str || str === "NaN:NaN") return 0;

  // 🔴 บังคับแปลงเป็น String ก่อนเผื่อหลุด เพื่อให้ใช้ .split ได้ชัวร์ๆ
  const [h, m] = String(str).split(":").map(Number);

  return (h || 0) * 60 + (m || 0);
}

function minutesToTime(mins) {
  const h = Math.floor(mins / 60);
  const m = mins % 60;

  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
}

/**
 * ฟังก์ชันหลักสำหรับคำนวณเวลาเข้า-ออกงานของตำรวจ
 * @param {string} payloadSheetName ชื่อชีตข้อมูล log (เช่น '2026-04-19_to_2026-04-25')
 */
function actionSummaryTrigger(payloadSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(payloadSheetName);
  const userSheet = ss.getSheetByName("POLICE_USER");

  if (!logSheet || !userSheet) {
    return json({ status: "error", message: `ไม่พบชีตข้อมูล ${payloadSheetName} หรือชีต POLICE_USER` });
  }

  // 1. ดึงข้อมูลรายชื่อตำรวจทั้งหมดจาก POLICE_USER
  const userData = userSheet.getDataRange().getValues();
  const policeMap = {};

  for (let i = 1; i < userData.length; i++) {
    const code = userData[i][0]; // POLICE_CODE (คอลัมน์ A)
    const name = userData[i][1]; // NAME (คอลัมน์ B)
    if (code) {
      policeMap[code] = {
        name: name,
        totalMs: 0,
        lastInTime: null,
        lastKnownTime: 0,
        hasAnomaly: false
      };
    }
  }

  // 2. ดึงข้อมูล Log การเข้าออก
  // ใช้ getValues เพื่อความคงเส้นคงวา (บางที Sheet จัด Format เป็น Date Object มาให้)
  const logData = logSheet.getDataRange().getValues();

  // แปลง Log ให้อยู่ในรูป Object และคัดให้ได้ Timestamp ที่ถูกต้อง ก่อนจะนำไปจัดเรียง
  const logs = [];
  for (let r = 1; r < logData.length; r++) {
    const row = logData[r];
    const code = row[1];
    const status = row[3];

    if (code && status) { // ตรวจสอบว่าแแถวนี้มีข้อมูลจริง ไม่ใช่เซลล์ว่าง
      const timestamp = parseSheetDateTime(row[4], row[5]);

      if (!timestamp) {
        // ถ้ารูปแบบวันที่หรือเวลาพัง/อ่านไม่ได้ ให้เตะ Error กลับไปแจ้งที่หน้า Dashboard ทันที
        return json({
          status: "error",
          message: `ไม่สามารถคำนวณได้ ข้อมูลรูปแบบวันที่หรือเวลาผิดพลาด (ไฟล์ ${payloadSheetName} แถวที่ ${r + 1})`
        });
      }

      logs.push({
        index: Number(row[0]) || r, // ใช้ Index จากคอลัมน์แรกสุด หรือใช้ลำดับแถว (r) หากไม่มีค่า
        code: code,
        status: status,
        timestamp: timestamp
      });
    }
  }

  // 🔥 สำคัญ: เปลี่ยนมาเรียงลำดับด้วย Index (ลำดับการเกิดจริง) แทนการเรียงเวลา
  // เพื่อป้องกันเวลาเจอ Record นาฬิกาเพี้ยนแล้วมันสลับตำแหน่งกันจนเจ๊งคู่ IN-OUT
  logs.sort((a, b) => a.index - b.index);

  // 3. คำนวณเวลาตามเงื่อนไข (ต้องเจอ IN ก่อน OUT) -> ระบบรองรับเข้างานข้ามคืน / ข้ามวันแล้ว
  logs.forEach(log => {
    const code = log.code;
    const status = log.status;
    const currentTimestamp = log.timestamp;

    if (policeMap[code] && !isNaN(currentTimestamp.getTime())) {
      const currentMs = currentTimestamp.getTime();

      // 🕵️‍♂️ ถ้าเวลาปัจจุบัน ย้อนอดีตกลับไปก่อนเวลาล่าสุดที่เคยเจอ (ของคนๆ นี้) -> แปลว่าเวลาเรียงผิดปกติละ!
      if (policeMap[code].lastKnownTime && currentMs < policeMap[code].lastKnownTime) {
        policeMap[code].hasAnomaly = true;
      }
      policeMap[code].lastKnownTime = Math.max(policeMap[code].lastKnownTime || 0, currentMs);

      if (status === "IN") {
        // บันทึกเวลา IN ล่าสุดไว้ 
        policeMap[code].lastInTime = currentTimestamp;
      } else if (status === "OUT") {
        // ถ้าเจอ OUT และมี IN ก่อนหน้า ให้คำนวณส่วนต่าง
        if (policeMap[code].lastInTime) {
          const diffMs = currentTimestamp.getTime() - policeMap[code].lastInTime.getTime();
          if (diffMs > 0) {
            policeMap[code].totalMs += diffMs;
          } else {
            // ถ้า OUT ดันเกิดก่อน IN (ในคู่การ check-in/out) ก็ให้มาร์คว่าหลอนเหมือนกัน
            policeMap[code].hasAnomaly = true;
          }
          // เคลียร์ค่า IN ทิ้ง เพื่อรอรอบถัดไป
          policeMap[code].lastInTime = null;
        }
      }
    }
  });

  // 4. สร้างชีตสรุปผล (Summary)
  const summarySheetName = "Summary_" + payloadSheetName;
  let summarySheet = ss.getSheetByName(summarySheetName);

  // หาตำแหน่งของชีตสัปดาห์นั้น (logSheet) 
  // getIndex() สตาร์ทที่ 1, แต่เวลา insertSheet ใช้ base 0 
  // ตัวเลข logSheet.getIndex() เลยจะเป็นตำแหน่งทางขวาของ logSheet พอดีเป๊ะ
  const insertIndex = logSheet.getIndex();

  if (!summarySheet) {
    summarySheet = ss.insertSheet(summarySheetName, insertIndex);
  } else {
    summarySheet.clear();
    // ถ้าย้ายชีตที่สร้างไปแล้ว ให้ใช้คำสั่ง move เพื่อย้ายไปทางขวาต่อจากชีตต้นฉบับ
    ss.setActiveSheet(summarySheet);
    // แต่ถ้ามันอยู่คนละที่ไปไกลก็ใช้ moveActiveSheet ย้ายไปต่อตูดได้เลย
    // บวก 1 เพราะ moveActiveSheet ใช้ base 1
    const targetMovePos = logSheet.getIndex() + (summarySheet.getIndex() < logSheet.getIndex() ? 0 : 1);
    ss.moveActiveSheet(targetMovePos);
  }

  // เตรียมข้อมูลก่อนเขียนลงชีต เพื่อนำมาจัดเรียงอันดับ
  let userSummary = [];
  for (const code in policeMap) {
    const totalMinutes = Math.floor(policeMap[code].totalMs / (1000 * 60));

    // แปลงกลับเป็นชั่วโมงและนาทีตรงๆ
    const h = Math.floor(totalMinutes / 60);
    const m = totalMinutes % 60;

    // สร้างเป็นข้อความเช่น "24:45"
    const formattedText = `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
    const remark = policeMap[code].hasAnomaly ? "⚠️ เวลาผิดปกติ" : "";

    userSummary.push({
      code: code,
      name: policeMap[code].name,
      totalMinutes: totalMinutes,
      formattedText: formattedText,
      remark: remark
    });
  }

  // 🥇 จัดเรียงจากเข้าเวรมากไปน้อยสุด
  userSummary.sort((a, b) => b.totalMinutes - a.totalMinutes);

  const output = [["TOP_RANK", "POLICE_CODE", "NAME", "HOURS_SERVED", "TOTAL_MINUTES", "REMARK"]];

  userSummary.forEach((u, i) => {
    let topRankStr = "-";
    // ให้อันดับเฉพาะคนที่มีเวลาเข้าเวรมากกว่า 0
    if (u.totalMinutes > 0 && i < 6) {
      if (i === 0) topRankStr = "🏆 อันดับ 1";
      else if (i === 1) topRankStr = "🥈 อันดับ 2";
      else if (i === 2) topRankStr = "🥉 อันดับ 3";
      else topRankStr = `🏅 อันดับ ${i + 1}`;
    }

    output.push([
      topRankStr,
      u.code,
      u.name,
      u.formattedText,
      u.totalMinutes,
      u.remark
    ]);
  });

  // เตรียมพื้นที่ตาราง
  const dataRange = summarySheet.getRange(1, 1, output.length, output[0].length);

  // สำคัญมาก: ล็อค Format ของคอลัมน์ HOURS_SERVED ให้เป็น Plain Text "@" (ข้อความล้วน) ก่อนฝังข้อมูล
  // HOURS_SERVED ตอนนี้ขยับไปเป็นคอลัมน์ที่ 4 แล้ว (D)
  if (output.length > 1) {
    summarySheet.getRange(2, 4, output.length - 1, 1).setNumberFormat("@");
  }

  // วางข้อมูลทั้งหมดลงตาราง (เมื่อตั้งค่าเป็น Text แล้ว มันจะวางลงไปเป็น "24:45" ตรงๆ)
  dataRange.setValues(output);

  // 5. ส่ง JSON Response กลับไปให้ Dashboard แจ้งเตือนสถานะสำเร็จ (จำเป็นมาก!)
  return json({
    status: "success",
    message: `สรุปข้อมูลลงชีต ${summarySheetName} เรียบร้อย`
  });
}

// ============================================
// ฟังก์ชันอัปเดตข้อมูลผู้ใช้งาน (Admin Management)
// ============================================
function updateUser(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("POLICE_USER");
    if (!sheet) {
      return json({ status: "error", message: "ไม่พบชีต POLICE_USER" });
    }

    if (!data.POLICE_CODE) {
      return json({ status: "error", message: "กรุณาระบุ POLICE_CODE" });
    }

    const rows = sheet.getDataRange().getValues();
    let foundRow = -1;

    // หาแถวของ Code ที่ตรงกัน (A = POLICE_CODE)
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] == data.POLICE_CODE) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow === -1) {
      return json({ status: "error", message: "ไม่พบรหัสตำรวจในระบบ: " + data.POLICE_CODE });
    }

    // เขียนค่าทับเฉพาะช่องที่กำหนด
    // คอลัมน์ลำดับ: A=1(CODE), B=2(NAME), C=3(EMAIL), D=4(ON_DULTY), E=5(ACTIVE), F=6(ROLE), G=7(WORK_TIME), H=8(DATE/TMN_DT)
    if (data.NAME !== undefined) sheet.getRange(foundRow, 2).setValue(data.NAME);
    if (data.EMAIL !== undefined) sheet.getRange(foundRow, 3).setValue(data.EMAIL);
    // STATUS ของเราแมตช์เข้าคอลัมน์ ACTIVE
    if (data.STATUS !== undefined) sheet.getRange(foundRow, 5).setValue(data.STATUS);
    if (data.RANK !== undefined) sheet.getRange(foundRow, 6).setValue(data.RANK);
    
    // บันทึกรูปภาพลงคอลัมน์ 10 (J)
    if (data.PIC_URL !== undefined) sheet.getRange(foundRow, 10).setValue(data.PIC_URL);

    // 🕒 อัปเดต TMN_DT (วันที่ตอนแก้)
    sheet.getRange(foundRow, 9).setValue(new Date());

    return json({
      status: "success",
      message: "อัปเดตข้อมูลของบุคคลากรสำเร็จ",
      data: {
        POLICE_CODE: data.POLICE_CODE,
        NAME: data.NAME,
        STATUS: data.STATUS,
        RANK: data.RANK
      }
    });

  } catch (err) {
    return json({ status: "error", message: err.message });
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// ฟังก์ชันจัดการยศ (Rank Management)
// ============================================
function manageRank(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RANK_POLICE");
    if (!sheet) {
      return json({ status: "error", message: "ไม่พบชีต RANK_POLICE" });
    }

    if (data.subAction === "ADD") {
      if (!data.RANK_NAME) return json({ status: "error", message: "กรุณาระบุชื่อยศ" });

      const rows = sheet.getDataRange().getValues();
      let newId = 1;
      if (rows.length > 1) {
        // หาค่า ID สูงสุดแล้วบวก 1
        newId = Math.max(...rows.slice(1).map(r => parseInt(r[0]) || 0)) + 1;
      }
      sheet.appendRow([newId, data.RANK_NAME]);
      return json({ status: "success", message: "เพิ่มยศใหม่เรียบร้อยแล้ว" });

    } else if (data.subAction === "DELETE") {
      if (!data.RANK_ID) return json({ status: "error", message: "กรุณาระบุรหัสยศ" });

      const rows = sheet.getDataRange().getValues();
      let foundRow = -1;
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] == data.RANK_ID) {
          foundRow = i + 1;
          break;
        }
      }

      if (foundRow > -1) {
        sheet.deleteRow(foundRow);
        return json({ status: "success", message: "ลบยศออกจากระบบแล้ว" });
      } else {
        return json({ status: "error", message: "ไม่พบรหัสยศที่ต้องการลบ" });
      }
    }

    return json({ status: "error", message: "Invalid subAction for manageRank" });

  } catch (err) {
    return json({ status: "error", message: err.message });
  }
}

// ============================================
// ฟังก์ชันคำนวณและสรุปยอดรางวัล (Wall of Fame)
// ============================================
function actionSummaryReward(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName("GLOBAL_LOGS");
    const userSheet = ss.getSheetByName("POLICE_USER");

    if (!logSheet || !userSheet) {
      return json({ status: "error", message: "ไม่พบชีต GLOBAL_LOGS หรือ POLICE_USER" });
    }

    // 1. ดึงผู้เข้าใช้อ้างอิง
    const userData = userSheet.getDataRange().getValues();
    const policeMap = {};

    for (let i = 1; i < userData.length; i++) {
      const code = userData[i][0];
      const name = userData[i][1];
      const picUrl = userData[i][9] || ""; // คอลัมน์ที่ 10 (J: PIC_URL)

      if (code) {
        policeMap[code] = {
          name: name,
          picUrl: picUrl,
          totalMs: 0,
          lastInTime: null,
        };
      }
    }

    // 2. ตั้งค่าขอบเขตเวลา
    const startObj = new Date(data.startDate + "T00:00:00+07:00");
    const endObj = new Date(data.endDate + "T23:59:59+07:00");

    const logData = logSheet.getDataRange().getValues();
    const logs = [];

    // ลูปแค่ข้อมูลที่มี Date Time อยู่ในช่วง (Filter phase)
    for (let r = 1; r < logData.length; r++) {
      const row = logData[r];
      const code = row[1];
      const status = row[3];
      if (code && status) {
        const timestamp = parseSheetDateTime(row[4], row[5]);
        if (timestamp && timestamp.getTime() >= startObj.getTime() && timestamp.getTime() <= endObj.getTime()) {
          logs.push({
            index: Number(row[0]) || r,
            code: code,
            status: status,
            timestamp: timestamp
          });
        }
      }
    }

    logs.sort((a, b) => a.index - b.index);

    // 3. คำนวณ
    logs.forEach(log => {
      const code = log.code;
      const status = log.status;
      const currentTimestamp = log.timestamp;

      if (policeMap[code]) {
        if (status === "IN") {
          policeMap[code].lastInTime = currentTimestamp;
        } else if (status === "OUT") {
          if (policeMap[code].lastInTime) {
            const diffMs = currentTimestamp.getTime() - policeMap[code].lastInTime.getTime();
            if (diffMs > 0) {
              policeMap[code].totalMs += diffMs;
            }
            // รีเซ็ตเพื่อรอรอบต่อไป
            policeMap[code].lastInTime = null;
          }
        }
      }
    });

    // 4. บันทึกลง WALL_OF_FAME
    let wofSheet = ss.getSheetByName("WALL_OF_FAME");
    if (!wofSheet) {
      wofSheet = ss.insertSheet("WALL_OF_FAME", ss.getNumSheets());
      wofSheet.appendRow([
        "PERIOD",
        "RANK",
        "POLICE_CODE",
        "NAME",
        "TOTAL_MINUTES",
        "HOURS_SERVED",
        "CALCULATED_ON",
        "PIC_URL"
      ]);
      wofSheet.setFrozenRows(1);
    }

    // สำคัญ: บังคับหัวคอลัมน์ A (PERIOD), F (HOURS_SERVED) เป็น Text เสมอ ป้องกันการแปลผิดเป็นวันที่
    wofSheet.getRange("A:A").setNumberFormat("@");
    wofSheet.getRange("F:F").setNumberFormat("@");
    wofSheet.getRange("G:G").setNumberFormat("yyyy-mm-dd hh:mm:ss");

    // Clear Previous records of the same period
    const wofData = wofSheet.getDataRange().getValues();
    const rowsToDelete = [];
    for (let i = wofData.length - 1; i >= 1; i--) {
      if (wofData[i][0] === data.periodLabel) {
        rowsToDelete.push(i + 1); // +1 for 1-based index
      }
    }
    // ไล่ลบจากข้างล่างขึ้นข้างบน เพื่อไม่ให้บรรทัดรวน
    rowsToDelete.forEach(rowIdx => {
      wofSheet.deleteRow(rowIdx);
    });

    // ลำดับและเตรียม Payload
    let userSummary = [];
    for (const code in policeMap) {
      const totalMinutes = Math.floor(policeMap[code].totalMs / (1000 * 60));
      if (totalMinutes > 0) {
        const h = Math.floor(totalMinutes / 60);
        const m = totalMinutes % 60;
        const formattedText = `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;

        userSummary.push({
          code: code,
          name: policeMap[code].name,
          picUrl: policeMap[code].picUrl,
          totalMinutes: totalMinutes,
          formattedText: formattedText
        });
      }
    }

    userSummary.sort((a, b) => b.totalMinutes - a.totalMinutes);

    const calcOn = new Date();
    userSummary.forEach((u, i) => {
      let rankLabel = (i + 1).toString();
      wofSheet.appendRow([
        data.periodLabel,
        rankLabel,
        u.code,
        u.name,
        u.totalMinutes,
        u.formattedText,
        calcOn,
        u.picUrl
      ]);
    });

    return json({
      status: "success",
      message: `บันทึกข้อมูลโขว์ผลงานบน Wall of Fame รอบ ${data.periodLabel} เรียบร้อย`
    });
  } catch (err) {
    return json({ status: "error", message: err.message });
  }
}