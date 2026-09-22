// ===== 全域設定 =====
var SHEET_ID = '1eaSKqrp7iQyW2yahpSV0a3ZT4A3jVo_lZjHpQQvcfNw';

function getSpreadsheet() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function getSheet(name) {
  var sheet = getSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('找不到工作表：' + name);
  return sheet;
}

/** 只讀實際有資料的範圍，避免 getDataRange 掃到過大空白區 */
function getSheetValues_(name, maxCols) {
  var sheet = getSheet(name);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return [];
  if (maxCols && lastCol > maxCols) lastCol = maxCols;
  return sheet.getRange(1, 1, lastRow, lastCol).getValues();
}

function readSettingsMap_() {
  var values = getSheetValues_('系統設定', 2);
  var config = {};
  for (var i = 0; i < values.length; i++) {
    var key = String(values[i][0] || '').trim();
    if (!key) continue;
    config[key] = values[i][1];
  }
  return config;
}

function normalizeSettingKey_(key) {
  return String(key || '').trim().replace(/\s+/g, '');
}

function getSettingByHints_(hints) {
  var values = getSheetValues_('系統設定', 2);
  for (var i = 0; i < values.length; i++) {
    var key = normalizeSettingKey_(values[i][0]);
    for (var h = 0; h < hints.length; h++) {
      if (key.indexOf(normalizeSettingKey_(hints[h])) !== -1) {
        var val = String(values[i][1] || '').trim();
        if (val) return val;
      }
    }
  }
  return '';
}

function formatSheetDate_(d) {
  if (d instanceof Date) {
    return Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM-dd');
  }
  return String(d || '');
}

/** GET 參數可能已被 Apps Script 解碼一次，避免二次 decode 失敗 */
function parseRequestData_(raw) {
  if (raw == null || raw === '') throw new Error('缺少 data 參數');
  var text = String(raw);
  try {
    return JSON.parse(text);
  } catch (e1) {
    try {
      return JSON.parse(decodeURIComponent(text));
    } catch (e2) {
      throw new Error('無法解析提交資料');
    }
  }
}

var DEFAULT_REPORT_FOLDER_ID = ''; // 從 Google Sheet「系統設定」讀取，不寫入程式碼

// Web App 匿名模式無法使用 Drive API，改用佇列 + 觸發器
function queueFileMove(fileId) {
  var props = PropertiesService.getScriptProperties();
  var queue = props.getProperty('moveQueue');
  var list = queue ? JSON.parse(queue) : [];
  list.push(fileId);
  props.setProperty('moveQueue', JSON.stringify(list));
}

// 由時間觸發器執行（每分鐘），在 owner 權限下移動檔案
function processMoveQueue() {
  var props = PropertiesService.getScriptProperties();
  var queue = props.getProperty('moveQueue');
  if (!queue) return;
  var list = JSON.parse(queue);
  if (list.length === 0) return;

  var folderId = DEFAULT_REPORT_FOLDER_ID || getSettingByHints_(['報表資料夾', '資料夾ID']);
  if (!folderId) return; // 尚未設定資料夾 ID，跳過移動

  var folder = DriveApp.getFolderById(folderId);
  var remaining = [];
  for (var i = 0; i < list.length; i++) {
    try {
      var file = DriveApp.getFileById(list[i]);
      file.moveTo(folder);
    } catch (e) {
      remaining.push(list[i]);
    }
  }
  props.setProperty('moveQueue', JSON.stringify(remaining));
}

// 在編輯器執行一次，設定每分鐘觸發器
function setupMoveTrigger() {
  // 先清除舊的
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processMoveQueue') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('processMoveQueue')
    .timeBased()
    .everyMinutes(1)
    .create();
  Logger.log('觸發器已設定：每分鐘執行 processMoveQueue');
}

// ===== 初始化（執行一次即可） =====
function initializeSheets() {
  var ss = getSpreadsheet();

  // 建立分頁（如果不存在）
  var tabs = ['系統設定', '人員名冊', '學生名冊', '課程設定', '教學日誌', '出缺席記錄', '成績設定', '成績記錄'];
  tabs.forEach(function(name) {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
    }
  });

  // 刪除預設的 Sheet1（如果存在且不在我們的列表中）
  var defaultSheet = ss.getSheetByName('工作表1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  // 1. 系統設定
  var s1 = ss.getSheetByName('系統設定');
  s1.clear();
  s1.getRange(1, 1, 8, 2).setValues([
    ['學校名稱', '國姓國民小學'],
    ['進修部名稱', '進修部'],
    ['管理者密碼', ''],
    ['鐘點費單價', 405],
    ['每日節數', 3],
    ['上課時間', '19:00~21:00'],
    ['縣市名稱', '南投縣'],
    ['報表資料夾ID', '']
  ]);

  // 2. 人員名冊
  var s2 = ss.getSheetByName('人員名冊');
  s2.clear();
  s2.getRange(1, 1, 6, 6).setValues([
    ['姓名', '角色', '狀態', '額外費用名稱', '額外費用金額', '備註'],
    ['林思遠', '校長', '在職', '校長兼職費', 2333, '三班以下3500元的三分之二'],
    ['吳怡萱', '導師', '在職', '導師費', 4000, '比照國民小學導師費標準'],
    ['余曜男', '教師', '在職', '', '', ''],
    ['劉政勳', '教師', '在職', '', '', ''],
    ['康雲昇', '教師', '在職', '', '', '']
  ]);

  // 3. 學生名冊
  var s3 = ss.getSheetByName('學生名冊');
  s3.clear();
  s3.getRange(1, 1, 12, 2).setValues([
    ['姓名', '狀態'],
    ['阮氏彫', '在學'],
    ['阮紅妮', '在學'],
    ['阮玄莊', '在學'],
    ['范宥嫺', '在學'],
    ['黎美香', '在學'],
    ['馬銨妤', '在學'],
    ['陳錦江', '在學'],
    ['黎氏銀', '在學'],
    ['范氏燕萍', '在學'],
    ['陳氏錦秀', '在學'],
    ['阮氏雪梅', '在學']
  ]);

  // 4. 課程設定
  var s4 = ss.getSheetByName('課程設定');
  s4.clear();
  s4.getRange(1, 1, 5, 3).setValues([
    ['課程名稱', '星期', '授課教師'],
    ['國語與彈性', '一', '吳怡萱'],
    ['社會生活與彈性', '二', '劉政勳'],
    ['國語與英文', '三', '康雲昇'],
    ['數學與科學', '四', '余曜男']
  ]);

  // 5. 教學日誌
  var s5 = ss.getSheetByName('教學日誌');
  s5.clear();
  s5.getRange(1, 1, 1, 6).setValues([
    ['日期', '星期', '時間', '課程', '上課內容', '授課教師']
  ]);

  // 6. 出缺席記錄
  var s6 = ss.getSheetByName('出缺席記錄');
  s6.clear();
  s6.getRange(1, 1, 1, 3).setValues([
    ['日期', '星期', '課程']
  ]);

  // 7. 成績設定
  var s7 = ss.getSheetByName('成績設定');
  s7.clear();
  s7.getRange(1, 1, 6, 2).setValues([
    ['成績科目名稱', '類型'],
    ['國語', '學科'],
    ['數學', '學科'],
    ['社會', '學科'],
    ['自然', '學科'],
    ['英文', '學科']
  ]);

  // 8. 成績記錄
  var s8 = ss.getSheetByName('成績記錄');
  s8.clear();
  s8.getRange(1, 1, 1, 7).setValues([
    ['學年度', '學期', '學生姓名', '科目', '平時成績', '考試成績', '學期成績']
  ]);

  return '初始化完成！共建立 ' + tabs.length + ' 個分頁。';
}

// ===== Web App 入口 =====
function doGet(e) {
  var action = e.parameter.action;
  var result;

  try {
    switch (action) {
      case 'init':
        result = { success: true, message: initializeSheets() };
        break;
      case 'load_config':
        result = loadConfig();
        break;
      case 'verify_admin':
        result = verifyAdmin(e.parameter.pwd);
        break;
      case 'load_admin_config':
        result = handleAdminAction(e, loadAdminConfig);
        break;
      case 'get_dashboard':
        result = handleAdminAction(e, function() {
          return getDashboard(e.parameter.date);
        });
        break;
      case 'get_students_in_range':
        result = handleAdminAction(e, function() {
          return getStudentsInRange(e.parameter.start, e.parameter.end);
        });
        break;
      case 'export_log':
        result = handleAdminAction(e, function() {
          return exportTeachingLog(e.parameter.year, e.parameter.month);
        });
        break;
      case 'export_salary':
        result = handleAdminAction(e, function() {
          return exportSalary(e.parameter.year, e.parameter.month);
        });
        break;
      case 'export_payslip':
        result = handleAdminAction(e, function() {
          return exportPayslip(e.parameter.year, e.parameter.month);
        });
        break;
      case 'export_attendance':
        result = handleAdminAction(e, function() {
          return exportAttendance(e.parameter.start, e.parameter.end, e.parameter.students);
        });
        break;
      case 'check_date':
        result = checkDateExists(e.parameter.date);
        break;
      case 'submit_log':
        result = submitLog(parseRequestData_(e.parameter.data));
        break;
      case 'submit_attendance':
        result = submitAttendance(parseRequestData_(e.parameter.data));
        break;
      default:
        result = { success: false, error: '未知的 action: ' + action };
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var result;
  try {
    var data = {};
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    }
    var action = data.action || (e.parameter && e.parameter.action);
    switch (action) {
      case 'submit_log':
        result = submitLog(data);
        break;
      case 'submit_attendance':
        result = submitAttendance(data);
        break;
      case 'check_date':
        result = checkDateExists(data.date || (e.parameter && e.parameter.date));
        break;
      default:
        result = { success: false, error: '未知的 action: ' + action };
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== 管理者驗證 =====
function handleAdminAction(e, callback) {
  var pwd = e.parameter.pwd;
  var settings = readSettingsMap_();
  var adminPwd = String(settings['管理者密碼'] || '');
  if (pwd !== adminPwd) {
    return { success: false, error: '密碼錯誤，無權限執行此操作' };
  }
  return callback();
}

function verifyAdmin(pwd) {
  var settings = readSettingsMap_();
  var adminPwd = String(settings['管理者密碼'] || '');
  if (pwd === adminPwd) {
    return { success: true };
  }
  return { success: false, error: '密碼錯誤' };
}

// ===== Task 2: 公開 API =====

function loadConfig() {
  var settings = readSettingsMap_();
  var config = {};
  Object.keys(settings).forEach(function(key) {
    if (key === '管理者密碼') return;
    if (key === '鐘點費單價' || key === '每日節數') return;
    config[key] = settings[key];
  });

  var staffData = getSheetValues_('人員名冊', 6);
  var staff = [];
  for (var i = 1; i < staffData.length; i++) {
    if (staffData[i][2] === '在職' && staffData[i][1] !== '校長') {
      staff.push({
        name: staffData[i][0],
        role: staffData[i][1]
      });
    }
  }

  var studentData = getSheetValues_('學生名冊', 3);
  var students = [];
  for (var i = 1; i < studentData.length; i++) {
    if (studentData[i][1] === '在學') {
      students.push({ name: studentData[i][0], status: studentData[i][1] });
    }
  }

  var courseData = getSheetValues_('課程設定', 3);
  var courses = [];
  for (var i = 1; i < courseData.length; i++) {
    if (!courseData[i][0]) continue;
    courses.push({
      name: courseData[i][0],
      weekday: courseData[i][1],
      teacher: courseData[i][2]
    });
  }

  return {
    success: true,
    config: config,
    staff: staff,
    students: students,
    courses: courses
  };
}

function checkDateExists(dateStr) {
  var hasLog = false;
  var hasAtt = false;
  var logData = getSheetValues_('教學日誌', 1);
  for (var i = 1; i < logData.length; i++) {
    if (formatSheetDate_(logData[i][0]) === dateStr) { hasLog = true; break; }
  }
  var attData = getSheetValues_('出缺席記錄', 1);
  for (var i = 1; i < attData.length; i++) {
    if (formatSheetDate_(attData[i][0]) === dateStr) { hasAtt = true; break; }
  }
  return { success: true, hasLog: hasLog, hasAttendance: hasAtt };
}

function submitLog(data) {
  var sheet = getSheet('教學日誌');
  var rows = getSheetValues_('教學日誌', 6);
  var existingRow = -1;
  for (var i = 1; i < rows.length; i++) {
    if (formatSheetDate_(rows[i][0]) === data.date) {
      existingRow = i + 1;
      break;
    }
  }
  var rowData = [data.date, data.weekday, data.time, data.course, data.content, data.teacher];
  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  return { success: true };
}

function submitAttendance(data) {
  var sheet = getSheet('出缺席記錄');
  var lastCol = Math.max(sheet.getLastColumn(), 3);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  var studentNames = Object.keys(data.attendance);
  studentNames.forEach(function(name) {
    if (headers.indexOf(name) === -1) {
      var nextCol = headers.length + 1;
      sheet.getRange(1, nextCol).setValue(name);
      headers.push(name);
    }
  });

  var allData = getSheetValues_('出缺席記錄');
  var existingRow = -1;
  for (var i = 1; i < allData.length; i++) {
    if (formatSheetDate_(allData[i][0]) === data.date) {
      existingRow = i + 1;
      break;
    }
  }

  var row = [data.date, data.weekday, data.course];
  for (var c = 3; c < headers.length; c++) {
    var studentName = headers[c];
    row.push(data.attendance[studentName] || '');
  }

  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return { success: true };
}

// ===== Task 3: 管理者 API =====

function loadAdminConfig() {
  var config = readSettingsMap_();
  delete config['管理者密碼'];

  var staffData = getSheetValues_('人員名冊', 6);
  var staff = [];
  for (var i = 1; i < staffData.length; i++) {
    staff.push({
      name: staffData[i][0],
      role: staffData[i][1],
      status: staffData[i][2],
      extraFeeName: staffData[i][3] || '',
      extraFeeAmount: staffData[i][4] || '',
      note: staffData[i][5] || ''
    });
  }

  var studentData = getSheetValues_('學生名冊', 3);
  var students = [];
  for (var i = 1; i < studentData.length; i++) {
    students.push({ name: studentData[i][0], status: studentData[i][1] });
  }

  var courseData = getSheetValues_('課程設定', 3);
  var courses = [];
  for (var i = 1; i < courseData.length; i++) {
    if (!courseData[i][0]) continue;
    courses.push({
      name: courseData[i][0],
      weekday: courseData[i][1],
      teacher: courseData[i][2]
    });
  }

  return {
    success: true,
    config: config,
    staff: staff,
    students: students,
    courses: courses
  };
}

function getDashboard(dateParam) {
  var targetStr = dateParam || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');

  var studentData = getSheetValues_('學生名冊', 2);
  var allStudents = [];
  for (var s = 1; s < studentData.length; s++) {
    if (studentData[s][1] === '在學') {
      allStudents.push(studentData[s][0]);
    }
  }
  var totalStudents = allStudents.length;

  var logData = getSheetValues_('教學日誌', 6);
  var dayLogs = [];
  for (var i = 1; i < logData.length; i++) {
    if (formatSheetDate_(logData[i][0]) === targetStr) {
      dayLogs.push({
        date: targetStr,
        weekday: logData[i][1],
        time: logData[i][2],
        course: logData[i][3],
        content: logData[i][4],
        teacher: logData[i][5]
      });
    }
  }

  var attData = getSheetValues_('出缺席記錄');
  var headers = attData.length ? attData[0] : [];
  var hasAttData = false;
  var presentList = [];
  var leaveList = [];

  for (var i = 1; i < attData.length; i++) {
    if (formatSheetDate_(attData[i][0]) === targetStr) {
      hasAttData = true;
      presentList = [];
      leaveList = [];
      for (var c = 3; c < headers.length; c++) {
        var name = headers[c];
        if (allStudents.indexOf(name) === -1) continue;
        var status = attData[i][c];
        if (status === '✓') presentList.push(name);
        else if (status === '△') leaveList.push(name);
      }
    }
  }

  var markedNames = presentList.concat(leaveList);
  var absentList = [];
  if (hasAttData) {
    for (var a = 0; a < allStudents.length; a++) {
      if (markedNames.indexOf(allStudents[a]) === -1) {
        absentList.push(allStudents[a]);
      }
    }
  }

  var presentCount = presentList.length;
  var leaveCount = leaveList.length;
  var absentCount = absentList.length;
  var attendanceRate = totalStudents > 0 ? Math.round(presentCount / totalStudents * 100) : 0;

  return {
    success: true,
    date: targetStr,
    dayLogs: dayLogs,
    attendance: {
      hasData: hasAttData,
      totalStudents: totalStudents,
      presentCount: presentCount,
      leaveCount: leaveCount,
      absentCount: absentCount,
      rate: attendanceRate,
      presentList: presentList,
      leaveList: leaveList,
      absentList: absentList
    }
  };
}

// ===== Task 4: 教學日誌 XLS 生成（批次寫入） =====

function exportTeachingLog(yearStr, monthStr) {
  var config = readSettingsMap_();
  var year = parseInt(yearStr);
  var month = parseInt(monthStr);

  var logData = getSheetValues_('教學日誌', 6);
  var records = [];
  for (var i = 1; i < logData.length; i++) {
    var d = logData[i][0];
    if (!(d instanceof Date)) continue;
    if ((d.getFullYear() - 1911) === year && (d.getMonth() + 1) === month) {
      records.push({
        date: d,
        weekday: logData[i][1],
        time: logData[i][2],
        course: logData[i][3],
        content: logData[i][4],
        teacher: logData[i][5]
      });
    }
  }
  records.sort(function(a, b) { return a.date - b.date; });

  var fileName = config['縣市名稱'] + config['學校名稱'] + config['進修部名稱'] +
                 year + '年度' + month + '月教學日誌';
  var newSS = SpreadsheetApp.create(fileName);
  var ws = newSS.getActiveSheet();

  ws.setColumnWidth(1, 30);
  ws.setColumnWidth(2, 100);
  ws.setColumnWidth(3, 50);
  ws.setColumnWidth(4, 100);
  ws.setColumnWidth(5, 120);
  ws.setColumnWidth(6, 260);
  ws.setColumnWidth(7, 200);
  ws.setColumnWidth(8, 90);

  var title = config['縣市名稱'] + config['學校名稱'] + config['進修部名稱'] +
              ' ' + year + '年度' + month + '月 教學日誌';
  var headers = ['序號', '日期', '星期', '時間', '課程', '上課內容', '教師簽名', '授課教師'];
  var weekdays = ['日', '一', '二', '三', '四', '五', '六'];
  var values = [[title, '', '', '', '', '', '', ''], headers];

  for (var i = 0; i < records.length; i++) {
    var r = records[i];
    var dateStr = (r.date.getMonth() + 1) + '/' + r.date.getDate();
    var weekday = r.weekday || weekdays[r.date.getDay()];
    values.push([i + 1, dateStr, weekday, r.time, r.course, r.content, '', r.teacher]);
  }

  var signRowOffset = 3;
  var signLabelRow = ['', '進修部主任：', '', '', '校長：', '', '', ''];
  for (var pad = 0; pad < signRowOffset - 1; pad++) {
    values.push(['', '', '', '', '', '', '', '']);
  }
  values.push(signLabelRow);

  ws.getRange(1, 1, values.length, 8).setValues(values);
  ws.getRange(1, 1, 1, 8).merge()
    .setFontFamily('標楷體').setFontSize(20).setFontWeight('bold')
    .setHorizontalAlignment('center');
  ws.setRowHeight(1, 50);

  var headerRange = ws.getRange(2, 1, 1, 8);
  headerRange.setFontFamily('標楷體').setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);
  ws.setRowHeight(2, 55);

  if (records.length > 0) {
    var dataEnd = 2 + records.length;
    var dataRange = ws.getRange(3, 1, dataEnd, 8);
    dataRange.setFontFamily('標楷體').setFontSize(14).setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true);
    ws.getRange(3, 1, dataEnd, 4).setHorizontalAlignment('center');
    ws.getRange(3, 8, dataEnd, 8).setHorizontalAlignment('center');
    for (var rh = 0; rh < records.length; rh++) ws.setRowHeight(3 + rh, 55);
  }

  var signRow = 2 + records.length + signRowOffset;
  ws.getRange(signRow, 2, 1, 3).merge()
    .setFontFamily('標楷體').setFontSize(18).setFontWeight('bold');
  ws.getRange(signRow, 5, 1, 2).merge()
    .setFontFamily('標楷體').setFontSize(18).setFontWeight('bold');
  ws.setRowHeight(signRow, 60);

  var fileId = newSS.getId();
  queueFileMove(fileId);
  return {
    success: true,
    fileName: fileName,
    sheetUrl: 'https://docs.google.com/spreadsheets/d/' + fileId,
    recordCount: records.length
  };
}

// ===== Task 5: 月薪資總表（批次寫入） =====

function exportSalary(yearStr, monthStr) {
  var config = readSettingsMap_();
  var year = parseInt(yearStr);
  var month = parseInt(monthStr);
  var hourlyRate = parseInt(config['鐘點費單價']);
  var sessionsPerDay = parseInt(config['每日節數']);

  var staffData = getSheetValues_('人員名冊', 6);
  var staffList = [];
  for (var i = 1; i < staffData.length; i++) {
    if (staffData[i][2] === '在職') {
      staffList.push({
        name: staffData[i][0],
        role: staffData[i][1],
        extraFeeName: staffData[i][3] || '',
        extraFeeAmount: parseInt(staffData[i][4]) || 0,
        note: staffData[i][5] || ''
      });
    }
  }

  var logData = getSheetValues_('教學日誌', 6);
  var teacherStats = {};
  for (var i = 1; i < logData.length; i++) {
    var d = logData[i][0];
    if (!(d instanceof Date)) continue;
    if ((d.getFullYear() - 1911) === year && (d.getMonth() + 1) === month) {
      var teacher = logData[i][5];
      if (!teacherStats[teacher]) teacherStats[teacher] = { days: 0, dates: [] };
      teacherStats[teacher].days++;
      teacherStats[teacher].dates.push((d.getMonth() + 1) + '/' + d.getDate());
    }
  }

  var westYear = year + 1911;
  var fileName = config['學校名稱'] + config['進修部名稱'] + ' ' + year + '年' + month + '月支給費用';
  var newSS = SpreadsheetApp.create(fileName);
  var ws = newSS.getActiveSheet();
  var widths = [90, 120, 70, 70, 85, 80, 85, 80, 250];
  for (var w = 0; w < widths.length; w++) ws.setColumnWidth(w + 1, widths[w]);

  var lastDay = new Date(westYear, month, 0).getDate();
  var periodStr = month + '月1日～\n' + month + '月' + lastDay + '日';
  var sortedStaff = staffList.sort(function(a, b) {
    var order = { '校長': 0, '導師': 1, '進修部主任': 2, '教師': 3 };
    return (order[a.role] || 9) - (order[b.role] || 9);
  });

  var title = config['學校名稱'] + config['進修部名稱'] + '  ' + year + '年' + month + '月支給費用';
  var headers = ['姓名', '上課期間', '授課天數', '授課節數', '鐘點費單價', '合計', '額外費用', '合計', '備註'];
  var values = [[title, '', '', '', '', '', '', '', ''], headers];

  for (var i = 0; i < sortedStaff.length; i++) {
    var s = sortedStaff[i];
    var stats = teacherStats[s.name] || { days: 0, dates: [] };
    var isSchoolMaster = (s.role === '校長');
    var days = isSchoolMaster ? '' : stats.days;
    var sessions = isSchoolMaster ? '' : stats.days * sessionsPerDay;
    var rate = isSchoolMaster ? '' : hourlyRate;
    var hourlyTotal = isSchoolMaster ? '' : stats.days * sessionsPerDay * hourlyRate;
    var extraFee = s.extraFeeAmount > 0 ? s.extraFeeName + '\n' + s.extraFeeAmount : '';
    var grandTotal = (isSchoolMaster ? 0 : stats.days * sessionsPerDay * hourlyRate) + s.extraFeeAmount;
    var remark = isSchoolMaster ? '' : (stats.dates.length > 0 ? '授課日：' + stats.dates.join('、') : '本月無授課');
    values.push([s.name, periodStr, days, sessions, rate, hourlyTotal, extraFee, grandTotal, remark]);
  }

  var dataStartRow = 3;
  var totalRow = dataStartRow + sortedStaff.length;
  values.push(['合  計', '', '', '', '', '', '', '', '']);

  ws.getRange(1, 1, values.length, 9).setValues(values);
  ws.getRange(1, 1, 1, 9).merge()
    .setFontFamily('標楷體').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  ws.setRowHeight(1, 35);
  ws.getRange(2, 1, 1, 9).setFontFamily('標楷體').setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);
  ws.setRowHeight(2, 40);

  if (sortedStaff.length > 0) {
    var bodyEnd = dataStartRow + sortedStaff.length - 1;
    var body = ws.getRange(dataStartRow, 1, bodyEnd, 9);
    body.setFontFamily('標楷體').setFontSize(14).setFontWeight('bold')
      .setVerticalAlignment('middle').setWrap(true)
      .setBorder(true, true, true, true, true, true);
    ws.getRange(dataStartRow, 1, bodyEnd, 8).setHorizontalAlignment('center');
    for (var rh = 0; rh < sortedStaff.length; rh++) ws.setRowHeight(dataStartRow + rh, 55);
  }

  var lastDataRow = totalRow - 1;
  if (sortedStaff.length > 0) {
    ws.getRange(totalRow, 4).setFormula('=SUM(D' + dataStartRow + ':D' + lastDataRow + ')');
    ws.getRange(totalRow, 6).setFormula('=SUM(F' + dataStartRow + ':F' + lastDataRow + ')');
    ws.getRange(totalRow, 8).setFormula('=SUM(H' + dataStartRow + ':H' + lastDataRow + ')');
  }
  ws.getRange(totalRow, 1, 1, 9).setFontFamily('標楷體').setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);
  ws.setRowHeight(totalRow, 50);

  var signRow = totalRow + 2;
  ws.getRange(signRow, 1).setValue('承辦').setFontFamily('標楷體').setFontSize(14).setFontWeight('bold');
  ws.getRange(signRow, 3, 1, 2).merge().setValue('出納').setFontFamily('標楷體').setFontSize(14).setFontWeight('bold');
  ws.getRange(signRow, 6, 1, 2).merge().setValue('會計').setFontFamily('標楷體').setFontSize(14).setFontWeight('bold');
  ws.getRange(signRow, 8).setValue('校長').setFontFamily('標楷體').setFontSize(14).setFontWeight('bold');

  var fileId = newSS.getId();
  queueFileMove(fileId);
  return {
    success: true,
    fileName: fileName,
    sheetUrl: 'https://docs.google.com/spreadsheets/d/' + fileId
  };
}

// ===== Task 6: 薪資條（沿用結構，讀取改用 helpers） =====

function exportPayslip(yearStr, monthStr) {
  var config = readSettingsMap_();
  var year = parseInt(yearStr);
  var month = parseInt(monthStr);
  var hourlyRate = parseInt(config['鐘點費單價']);
  var sessionsPerDay = parseInt(config['每日節數']);
  var westYear = year + 1911;
  var lastDay = new Date(westYear, month, 0).getDate();
  var periodStr = month + '月1日～' + month + '月' + lastDay + '日';

  var staffData = getSheetValues_('人員名冊', 6);
  var staffList = [];
  for (var i = 1; i < staffData.length; i++) {
    if (staffData[i][2] === '在職') {
      staffList.push({
        name: staffData[i][0],
        role: staffData[i][1],
        extraFeeName: staffData[i][3] || '',
        extraFeeAmount: parseInt(staffData[i][4]) || 0,
        note: staffData[i][5] || ''
      });
    }
  }

  var logData = getSheetValues_('教學日誌', 6);
  var teacherStats = {};
  for (var i = 1; i < logData.length; i++) {
    var d = logData[i][0];
    if (!(d instanceof Date)) continue;
    if ((d.getFullYear() - 1911) === year && (d.getMonth() + 1) === month) {
      var teacher = logData[i][5];
      if (!teacherStats[teacher]) teacherStats[teacher] = { days: 0, dates: [] };
      teacherStats[teacher].days++;
      teacherStats[teacher].dates.push((d.getMonth() + 1) + '/' + d.getDate());
    }
  }

  var fileName = config['學校名稱'] + '補校支給費用 ' + year + '年' + month + '月薪資條';
  var newSS = SpreadsheetApp.create(fileName);
  var ws = newSS.getActiveSheet();
  var widths = [80, 130, 55, 55, 55, 70, 70, 70, 280];
  for (var w = 0; w < widths.length; w++) ws.setColumnWidth(w + 1, widths[w]);

  var currentRow = 1;
  var schoolName = config['學校名稱'];
  var sortedStaff = staffList.sort(function(a, b) {
    var order = { '校長': 0, '導師': 1, '進修部主任': 2, '教師': 3 };
    return (order[a.role] || 9) - (order[b.role] || 9);
  });

  for (var si = 0; si < sortedStaff.length; si++) {
    var s = sortedStaff[si];
    var stats = teacherStats[s.name] || { days: 0, dates: [] };
    var isSchoolMaster = (s.role === '校長');

    var titleSuffix = isSchoolMaster ? s.extraFeeName :
      (s.extraFeeName ? ('鐘點費與' + s.extraFeeName) : '鐘點費');
    ws.getRange(currentRow, 1, 1, 9).merge()
      .setValue(schoolName + '補校支給費用   ' + year + '年' +
        (month < 10 ? '0' : '') + month + '月' + titleSuffix)
      .setFontFamily('標楷體').setFontSize(13).setFontWeight('bold');
    ws.setRowHeight(currentRow, 28);
    currentRow++;

    var headerRow;
    if (isSchoolMaster) {
      headerRow = ['姓名', '上課期間', s.extraFeeName, '', '', '', '', '合計', '備註'];
      ws.getRange(currentRow, 1, 1, 9).setValues([headerRow]);
      ws.getRange(currentRow, 3, 1, 5).merge();
    } else if (s.extraFeeName) {
      headerRow = ['姓名', '上課期間', '天數', '節數', '單價', '合計', s.extraFeeName, '合計', '備註'];
      ws.getRange(currentRow, 1, 1, 9).setValues([headerRow]);
    } else {
      headerRow = ['姓名', '上課期間', '天數', '節數', '單價', '合計', '', '', '備註'];
      ws.getRange(currentRow, 1, 1, 9).setValues([headerRow]);
      ws.getRange(currentRow, 6, 1, 3).merge();
    }
    ws.getRange(currentRow, 1, 1, 9).setFontFamily('標楷體').setFontSize(13).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true);
    ws.setRowHeight(currentRow, 55);
    currentRow++;

    if (isSchoolMaster) {
      ws.getRange(currentRow, 1, 1, 9).setValues([[
        s.name, periodStr, s.extraFeeAmount, '', '', '', '', s.extraFeeAmount, s.note
      ]]);
      ws.getRange(currentRow, 3, 1, 5).merge();
    } else {
      var days = stats.days;
      var sessions = days * sessionsPerDay;
      var hourlyTotal = sessions * hourlyRate;
      var grandTotal = hourlyTotal + s.extraFeeAmount;
      var remarkParts = [];
      if (s.note) remarkParts.push(s.note);
      remarkParts.push(stats.dates.length > 0 ? ('授課日：' + stats.dates.join('、')) : (month + '月份無授課'));
      if (s.extraFeeName) {
        ws.getRange(currentRow, 1, 1, 9).setValues([[
          s.name, periodStr, days, sessions, hourlyRate, hourlyTotal, s.extraFeeAmount, grandTotal, remarkParts.join('\n')
        ]]);
      } else {
        ws.getRange(currentRow, 1, 1, 9).setValues([[
          s.name, periodStr, days, sessions, hourlyRate, hourlyTotal, '', '', remarkParts.join('\n')
        ]]);
        ws.getRange(currentRow, 6, 1, 3).merge();
      }
    }

    ws.getRange(currentRow, 1, 1, 9).setFontFamily('標楷體').setFontSize(13).setFontWeight('bold')
      .setVerticalAlignment('middle').setWrap(true)
      .setBorder(true, true, true, true, true, true);
    ws.getRange(currentRow, 1, 1, 8).setHorizontalAlignment('center');
    ws.setRowHeight(currentRow, 65);
    currentRow += 2;
  }

  var fileId = newSS.getId();
  queueFileMove(fileId);
  return {
    success: true,
    fileName: fileName,
    sheetUrl: 'https://docs.google.com/spreadsheets/d/' + fileId
  };
}

// ===== Task 7: 出缺席報表（批次寫入） =====

function getStudentsInRange(startStr, endStr) {
  var start = new Date(startStr);
  var end = new Date(endStr);
  end.setHours(23, 59, 59);

  var attData = getSheetValues_('出缺席記錄');
  var headers = attData.length ? attData[0] : [];
  var studentSet = {};
  for (var i = 1; i < attData.length; i++) {
    var d = attData[i][0];
    if (!(d instanceof Date)) continue;
    if (d >= start && d <= end) {
      for (var c = 3; c < headers.length; c++) {
        if (attData[i][c]) studentSet[headers[c]] = true;
      }
    }
  }

  var studentData = getSheetValues_('學生名冊', 2);
  var allStudents = [];
  for (var i = 1; i < studentData.length; i++) {
    allStudents.push({
      name: studentData[i][0],
      status: studentData[i][1],
      hasRecord: studentSet[studentData[i][0]] || false
    });
  }
  return { success: true, students: allStudents };
}

function exportAttendance(startStr, endStr, studentsStr) {
  var config = readSettingsMap_();
  var start = new Date(startStr);
  var end = new Date(endStr);
  end.setHours(23, 59, 59);
  var selectedStudents = studentsStr.split(',');

  var attData = getSheetValues_('出缺席記錄');
  var headers = attData.length ? attData[0] : [];
  var classDates = [];
  var dateRecords = {};

  for (var i = 1; i < attData.length; i++) {
    var d = attData[i][0];
    if (!(d instanceof Date)) continue;
    if (d < start || d > end) continue;

    var dateStr = Utilities.formatDate(d, 'Asia/Taipei', 'M/d');
    var weekdays = ['日', '一', '二', '三', '四', '五', '六'];
    var dateKey = dateStr + '(' + weekdays[d.getDay()] + ')';
    if (!dateRecords[dateKey]) {
      classDates.push({ key: dateKey, date: d });
      dateRecords[dateKey] = {};
    }
    for (var c = 3; c < headers.length; c++) {
      if (attData[i][c]) dateRecords[dateKey][headers[c]] = attData[i][c];
    }
  }
  classDates.sort(function(a, b) { return a.date - b.date; });

  var startRoc = (start.getFullYear() - 1911) + '/' + (start.getMonth() + 1) + '/' + start.getDate();
  var endRoc = (end.getFullYear() - 1911) + '/' + (end.getMonth() + 1) + '/' + end.getDate();
  var fileName = config['縣市名稱'] + config['學校名稱'] + config['進修部名稱'] + ' 出缺席記錄表';
  var newSS = SpreadsheetApp.create(fileName);
  var ws = newSS.getActiveSheet();
  var totalCols = 1 + classDates.length + 2;

  var title = config['縣市名稱'] + config['學校名稱'] + config['進修部名稱'] +
    ' 出缺席記錄表（' + startRoc + '～' + endRoc + '）';
  var headerRow = ['姓名'];
  for (var i = 0; i < classDates.length; i++) headerRow.push(classDates[i].key);
  headerRow.push('出席天數', '出席率');

  var titleRow = [title];
  for (var t = 1; t < totalCols; t++) titleRow.push('');
  var values = [titleRow, headerRow];

  for (var si = 0; si < selectedStudents.length; si++) {
    var name = selectedStudents[si];
    var presentDays = 0;
    var row = [name];
    for (var di = 0; di < classDates.length; di++) {
      var status = dateRecords[classDates[di].key][name] || '';
      row.push(status);
      if (status === '✓') presentDays++;
    }
    var rate = classDates.length > 0 ? Math.round(presentDays / classDates.length * 100) + '%' : '0%';
    row.push(presentDays, rate);
    values.push(row);
  }

  ws.getRange(1, 1, values.length, totalCols).setValues(values);
  ws.getRange(1, 1, 1, totalCols).merge()
    .setFontFamily('標楷體').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  ws.setRowHeight(1, 40);
  ws.getRange(2, 1, 1, totalCols).setFontFamily('標楷體').setFontSize(12).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);
  ws.setRowHeight(2, 35);
  ws.setColumnWidth(1, 80);
  for (var i = 2; i <= classDates.length + 1; i++) ws.setColumnWidth(i, 45);
  ws.setColumnWidth(totalCols - 1, 65);
  ws.setColumnWidth(totalCols, 60);

  if (selectedStudents.length > 0) {
    var bodyEnd = 2 + selectedStudents.length;
    var body = ws.getRange(3, 1, bodyEnd, totalCols);
    body.setFontFamily('標楷體').setFontSize(12).setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true);
    ws.getRange(3, 2, bodyEnd, totalCols).setHorizontalAlignment('center');
    for (var rh = 0; rh < selectedStudents.length; rh++) ws.setRowHeight(3 + rh, 30);
  }

  var fileId = newSS.getId();
  queueFileMove(fileId);
  return {
    success: true,
    fileName: fileName,
    sheetUrl: 'https://docs.google.com/spreadsheets/d/' + fileId
  };
}
