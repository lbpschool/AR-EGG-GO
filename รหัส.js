/**
 * AR Salted Egg Challenge : ภารกิจพิชิตไข่เค็มดินสอพอง
 * Google Apps Script Backend (Updated with AR Camera Passthrough)
 * เชื่อมต่อกับ Google Sheet ID: 1NI1pBQd-2u1lqB2IRwcdHUjKIIwZ8MZohmMZceXx7OA
 */

const SPREADSHEET_ID = '1NI1pBQd-2u1lqB2IRwcdHUjKIIwZ8MZohmMZceXx7OA';
const SHEET_NAME = 'บันทึกผลการเรียนรู้';
const TEACHER_PIN = '123';

/**
 * รองรับการรับข้อมูลคะแนนจากภายนอก (REST API สำหรับกรณีรันเว็บภายนอก iframe)
 */
function doPost(e) {
  try {
    let data = {};
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      data = e.parameter;
    }
    const result = saveStudentScore(data);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ให้บริการ Web App
 */
function doGet(e) {
  // ตรวจสอบและสร้างโครงสร้างชีตอัตโนมัติหากยังไม่มี
  setupSpreadsheet();
  
  if (e && e.parameter && e.parameter.action === 'getTeacherData') {
    const result = getTeacherDashboardData(e.parameter.pin);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('AR Salted Egg Challenge : ภารกิจพิชิตไข่เค็มดินสอพอง')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * ดึง Spreadsheet Object
 */
function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (err) {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

/**
 * สร้างและจัดรูปแบบหัวตาราง Google Sheet อัตโนมัติ
 */
function setupSpreadsheet() {
  const ss = getSpreadsheet();
  if (!ss) return { success: false, message: 'ไม่พบ Spreadsheet' };
  
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  
  const headers = [
    'วันที่/เวลา',
    'ชื่อ-นามสกุล',
    'เลขที่',
    'ชั้น/ห้อง',
    'คะแนนรวม (/100)',
    'จำนวนดาว',
    'ด่าน 1 คัดไข่ (/10)',
    'ด่าน 2 ล้าง/เช็ด (/10)',
    'ด่าน 3 สูตร 3:1:1 (/25)',
    'ด่าน 4 พอกไข่ (/20)',
    'ด่าน 5 คลุกแกลบ (/10)',
    'ด่าน 6 บรรจุ/หมัก (/10)',
    'โบนัส ระยะเวลา (/15)',
    'เวลาที่ใช้ (วินาที)',
    'จำนวนครั้งที่ใช้คำใบ้',
    'ด่านที่ควรฝึกเพิ่ม',
    'ป้ายรางวัล (Badge)'
  ];
  
  const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const isHeaderEmpty = !currentHeaders[0] || currentHeaders[0] === '';
  
  if (isHeaderEmpty) {
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground('#1b5e20'); // สีเขียวเข้มเข้ากับธีมไข่เค็มและการเกษตร
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    headerRange.setVerticalAlignment('middle');
    sheet.setRowHeight(1, 36);
    sheet.setFrozenRows(1);
    
    // ตั้งขนาดคอลัมน์ให้อ่านง่าย
    sheet.setColumnWidth(1, 160); // วันที่
    sheet.setColumnWidth(2, 180); // ชื่อ-นามสกุล
    sheet.setColumnWidth(3, 80);  // เลขที่
    sheet.setColumnWidth(4, 100); // ห้อง
    sheet.setColumnWidth(5, 120); // คะแนนรวม
    sheet.setColumnWidth(6, 100); // ดาว
    sheet.setColumnWidth(16, 180); // ควรฝึกเพิ่ม
    sheet.setColumnWidth(17, 200); // Badge
  }
  
  return { success: true, message: 'เตรียมโครงสร้าง Google Sheet เรียบร้อยแล้ว' };
}

/**
 * บันทึกคะแนนของนักเรียนลง Google Sheet
 */
function saveStudentScore(data) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      setupSpreadsheet();
      sheet = ss.getSheetByName(SHEET_NAME);
    }
    
    const now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    const row = [
      now,
      data.studentName || 'ไม่ระบุชื่อ',
      data.studentNo || '-',
      data.studentClass || '-',
      Number(data.totalScore) || 0,
      Number(data.stars) || 0,
      Number(data.stage1Score) || 0,
      Number(data.stage2Score) || 0,
      Number(data.stage3Score) || 0,
      Number(data.stage4Score) || 0,
      Number(data.stage5Score) || 0,
      Number(data.stage6Score) || 0,
      Number(data.bonusScore) || 0,
      Number(data.timeSpent) || 0,
      Number(data.hintCount) || 0,
      data.needsPractice || 'ไม่มี',
      data.badge || 'Salted Egg Master'
    ];
    
    sheet.appendRow(row);
    
    // จัดตำแหน่งกึ่งกลางในคอลัมน์คะแนนและข้อมูลทั่วไป
    const lastRow = sheet.getLastRow();
    const range = sheet.getRange(lastRow, 1, 1, row.length);
    range.setVerticalAlignment('middle');
    
    return {
      success: true,
      message: 'บันทึกคะแนนลง Google Sheet สำเร็จ!',
      timestamp: now
    };
  } catch (err) {
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการบันทึก: ' + err.toString()
    };
  }
}

/**
 * ดึงข้อมูลสำหรับแดชบอร์ดโหมดครู
 */
function getTeacherDashboardData(pin) {
  if (pin !== TEACHER_PIN) {
    return { success: false, message: 'รหัสผ่าน PIN ไม่ถูกต้อง' };
  }
  
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) {
      return {
        success: true,
        stats: {
          totalStudents: 0,
          averageScore: 0,
          averageStars: 0,
          stageAverages: {
            s1: 0, s2: 0, s3: 0, s4: 0, s5: 0, s6: 0, bonus: 0
          },
          commonWeakPoints: []
        },
        students: []
      };
    }
    
    const values = sheet.getDataRange().getValues();
    const rows = values.slice(1); // ข้ามแถว header
    
    let totalScoreSum = 0;
    let totalStarsSum = 0;
    let s1Sum = 0, s2Sum = 0, s3Sum = 0, s4Sum = 0, s5Sum = 0, s6Sum = 0, bonusSum = 0;
    const weakMap = {};
    
    const students = rows.map(r => {
      const totalScore = Number(r[4]) || 0;
      const stars = Number(r[5]) || 0;
      const s1 = Number(r[6]) || 0;
      const s2 = Number(r[7]) || 0;
      const s3 = Number(r[8]) || 0;
      const s4 = Number(r[9]) || 0;
      const s5 = Number(r[10]) || 0;
      const s6 = Number(r[11]) || 0;
      const bonus = Number(r[12]) || 0;
      const timeSpent = Number(r[13]) || 0;
      const hints = Number(r[14]) || 0;
      const weak = r[15] || '-';
      
      totalScoreSum += totalScore;
      totalStarsSum += stars;
      s1Sum += s1; s2Sum += s2; s3Sum += s3; s4Sum += s4; s5Sum += s5; s6Sum += s6; bonusSum += bonus;
      
      if (weak && weak !== 'ไม่มี' && weak !== '-') {
        weakMap[weak] = (weakMap[weak] || 0) + 1;
      }
      
      return {
        date: r[0],
        name: r[1],
        no: r[2],
        class: r[3],
        totalScore: totalScore,
        stars: stars,
        s1: s1, s2: s2, s3: s3, s4: s4, s5: s5, s6: s6, bonus: bonus,
        timeSpent: timeSpent,
        hints: hints,
        weak: weak,
        badge: r[16]
      };
    }).reverse(); // ล่าสุดขึ้นก่อน
    
    const count = rows.length;
    const stats = {
      totalStudents: count,
      averageScore: (totalScoreSum / count).toFixed(1),
      averageStars: (totalStarsSum / count).toFixed(1),
      stageAverages: {
        s1: (s1Sum / count).toFixed(1),
        s2: (s2Sum / count).toFixed(1),
        s3: (s3Sum / count).toFixed(1),
        s4: (s4Sum / count).toFixed(1),
        s5: (s5Sum / count).toFixed(1),
        s6: (s6Sum / count).toFixed(1),
        bonus: (bonusSum / count).toFixed(1)
      },
      commonWeakPoints: Object.keys(weakMap).map(k => ({ topic: k, count: weakMap[k] }))
        .sort((a, b) => b.count - a.count)
    };
    
    return {
      success: true,
      stats: stats,
      students: students,
      sheetUrl: ss.getUrl()
    };
  } catch (err) {
    return { success: false, message: 'ดึงข้อมูลไม่สำเร็จ: ' + err.toString() };
  }
}

/**
 * ฟังก์ชันทดสอบการเชื่อมต่อ
 */
function testConnection() {
  const result = setupSpreadsheet();
  const ss = getSpreadsheet();
  Logger.log(result.message + ' | Sheet Name: ' + ss.getName());
  return result;
}
