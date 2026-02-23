// ===== 全域設定 =====
var SHEET_ID = '1eaSKqrp7iQyW2yahpSV0a3ZT4A3jVo_lZjHpQQvcfNw';

function getSpreadsheet() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function getSheet(name) {
  return getSpreadsheet().getSheetByName(name);
}

// ===== Web App 入口 =====
function doGet(e) {
  var action = e.parameter.action;
  var result;

  try {
    switch (action) {
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
        result = handleAdminAction(e, getDashboard);
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
  var data = JSON.parse(e.postData.contents);
  var action = data.action || e.parameter.action;
  var result;

  try {
    switch (action) {
      case 'submit_log':
        result = submitLog(data);
        break;
      case 'submit_attendance':
        result = submitAttendance(data);
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
  var settings = getSheet('系統設定').getDataRange().getValues();
  var adminPwd = '';
  for (var i = 0; i < settings.length; i++) {
    if (settings[i][0] === '管理者密碼') {
      adminPwd = String(settings[i][1]);
      break;
    }
  }
  if (pwd !== adminPwd) {
    return { success: false, error: '密碼錯誤，無權限執行此操作' };
  }
  return callback();
}

function verifyAdmin(pwd) {
  var settings = getSheet('系統設定').getDataRange().getValues();
  var adminPwd = '';
  for (var i = 0; i < settings.length; i++) {
    if (settings[i][0] === '管理者密碼') {
      adminPwd = String(settings[i][1]);
      break;
    }
  }
  if (pwd === adminPwd) {
    return { success: true };
  }
  return { success: false, error: '密碼錯誤' };
}

// ===== Task 2: 公開 API =====

function loadConfig() {
  var settings = getSheet('系統設定').getDataRange().getValues();
  var config = {};
  for (var i = 0; i < settings.length; i++) {
    var key = settings[i][0];
    if (key === '管理者密碼') continue;
    if (key === '鐘點費單價' || key === '每日節數') continue;
    config[key] = settings[i][1];
  }

  var staffSheet = getSheet('人員名冊');
  var staffData = staffSheet.getDataRange().getValues();
  var staff = [];
  for (var i = 1; i < staffData.length; i++) {
    if (staffData[i][2] === '在職' && staffData[i][1] !== '校長') {
      staff.push({
        name: staffData[i][0],
        role: staffData[i][1]
      });
    }
  }

  var studentSheet = getSheet('學生名冊');
  var studentData = studentSheet.getDataRange().getValues();
  var students = [];
  for (var i = 1; i < studentData.length; i++) {
    if (studentData[i][1] === '在學') {
      students.push({ name: studentData[i][0], status: studentData[i][1] });
    }
  }

  var courseSheet = getSheet('課程設定');
  var courseData = courseSheet.getDataRange().getValues();
  var courses = [];
  for (var i = 1; i < courseData.length; i++) {
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

function submitLog(data) {
  var sheet = getSheet('教學日誌');
  sheet.appendRow([
    data.date,
    data.weekday,
    data.time,
    data.course,
    data.content,
    data.teacher
  ]);
  return { success: true };
}

function submitAttendance(data) {
  var sheet = getSheet('出缺席記錄');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var studentNames = Object.keys(data.attendance);
  studentNames.forEach(function(name) {
    if (headers.indexOf(name) === -1) {
      var nextCol = headers.length + 1;
      sheet.getRange(1, nextCol).setValue(name);
      headers.push(name);
    }
  });

  var row = [data.date, data.weekday, data.course];
  for (var c = 3; c < headers.length; c++) {
    var studentName = headers[c];
    row.push(data.attendance[studentName] || '');
  }

  sheet.appendRow(row);
  return { success: true };
}

// ===== Task 3: 管理者 API =====

function loadAdminConfig() {
  var settings = getSheet('系統設定').getDataRange().getValues();
  var config = {};
  for (var i = 0; i < settings.length; i++) {
    config[settings[i][0]] = settings[i][1];
  }
  delete config['管理者密碼'];

  var staffSheet = getSheet('人員名冊');
  var staffData = staffSheet.getDataRange().getValues();
  var staff = [];
  for (var i = 1; i < staffData.length; i++) {
    staff.push({
      name: staffData[i][0],
      role: staffData[i][1],
      status: staffData[i][2],
      extraFeeName: staffData[i][3] || '',
      extraFeeAmount: staffData[i][4] || 0,
      note: staffData[i][5] || ''
    });
  }

  var studentSheet = getSheet('學生名冊');
  var studentData = studentSheet.getDataRange().getValues();
  var students = [];
  for (var i = 1; i < studentData.length; i++) {
    students.push({ name: studentData[i][0], status: studentData[i][1] });
  }

  var courseSheet = getSheet('課程設定');
  var courseData = courseSheet.getDataRange().getValues();
  var courses = [];
  for (var i = 1; i < courseData.length; i++) {
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

function getDashboard() {
  var today = new Date();
  var todayStr = Utilities.formatDate(today, 'Asia/Taipei', 'yyyy-MM-dd');

  var logSheet = getSheet('教學日誌');
  var logData = logSheet.getDataRange().getValues();
  var todayLogs = [];
  for (var i = 1; i < logData.length; i++) {
    var logDate = logData[i][0];
    if (logDate instanceof Date) {
      logDate = Utilities.formatDate(logDate, 'Asia/Taipei', 'yyyy-MM-dd');
    }
    if (logDate === todayStr) {
      todayLogs.push({
        date: todayStr,
        weekday: logData[i][1],
        time: logData[i][2],
        course: logData[i][3],
        content: logData[i][4],
        teacher: logData[i][5]
      });
    }
  }

  var attSheet = getSheet('出缺席記錄');
  var attData = attSheet.getDataRange().getValues();
  var headers = attData[0];
  var todayAttendance = null;
  var presentCount = 0;
  var absentCount = 0;
  var totalStudents = 0;

  for (var i = 1; i < attData.length; i++) {
    var attDate = attData[i][0];
    if (attDate instanceof Date) {
      attDate = Utilities.formatDate(attDate, 'Asia/Taipei', 'yyyy-MM-dd');
    }
    if (attDate === todayStr) {
      todayAttendance = {};
      for (var c = 3; c < headers.length; c++) {
        var status = attData[i][c];
        if (status) {
          totalStudents++;
          if (status === '✓') presentCount++;
          else if (status === '△') absentCount++;
        }
      }
      break;
    }
  }

  var attendanceRate = totalStudents > 0 ? Math.round(presentCount / totalStudents * 100) : 0;

  return {
    success: true,
    date: todayStr,
    todayLogs: todayLogs,
    attendance: {
      hasData: todayAttendance !== null,
      presentCount: presentCount,
      absentCount: absentCount,
      totalStudents: totalStudents,
      rate: attendanceRate
    }
  };
}

// ===== Task 4: 教學日誌 XLS 生成 =====

function exportTeachingLog(yearStr, monthStr) {
  var settings = getSheet('系統設定').getDataRange().getValues();
  var config = {};
  for (var i = 0; i < settings.length; i++) config[settings[i][0]] = settings[i][1];

  var year = parseInt(yearStr);
  var month = parseInt(monthStr);

  var logSheet = getSheet('教學日誌');
  var logData = logSheet.getDataRange().getValues();
  var records = [];

  for (var i = 1; i < logData.length; i++) {
    var d = logData[i][0];
    if (!(d instanceof Date)) continue;
    var rocYear = d.getFullYear() - 1911;
    var m = d.getMonth() + 1;
    if (rocYear === year && m === month) {
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

  ws.merge(ws.getRange('A1:H1'));
  var titleCell = ws.getRange('A1');
  titleCell.setValue(config['縣市名稱'] + config['學校名稱'] + config['進修部名稱'] +
                     ' ' + year + '年度' + month + '月 教學日誌');
  titleCell.setFontFamily('標楷體').setFontSize(20).setFontWeight('bold')
    .setHorizontalAlignment('center');
  ws.setRowHeight(1, 50);

  var headers = ['序號', '日期', '星期', '時間', '課程', '上課內容', '教師簽名', '授課教師'];
  var headerRange = ws.getRange(2, 1, 1, 8);
  headerRange.setValues([headers]);
  headerRange.setFontFamily('標楷體').setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  headerRange.setBorder(true, true, true, true, true, true);
  ws.setRowHeight(2, 55);

  for (var i = 0; i < records.length; i++) {
    var row = i + 3;
    var r = records[i];
    var dateStr = (r.date.getMonth() + 1) + '/' + r.date.getDate();
    var weekdays = ['日', '一', '二', '三', '四', '五', '六'];
    var weekday = r.weekday || weekdays[r.date.getDay()];

    ws.getRange(row, 1).setValue(i + 1);
    ws.getRange(row, 2).setValue(dateStr);
    ws.getRange(row, 3).setValue(weekday);
    ws.getRange(row, 4).setValue(r.time);
    ws.getRange(row, 5).setValue(r.course);
    ws.getRange(row, 6).setValue(r.content);
    ws.getRange(row, 8).setValue(r.teacher);

    var dataRange = ws.getRange(row, 1, 1, 8);
    dataRange.setFontFamily('標楷體').setFontSize(14)
      .setVerticalAlignment('middle');
    dataRange.setBorder(true, true, true, true, true, true);
    ws.setRowHeight(row, 55);

    ws.getRange(row, 1, 1, 4).setHorizontalAlignment('center');
    ws.getRange(row, 8).setHorizontalAlignment('center');
  }

  var signRow = records.length + 5;
  ws.merge(ws.getRange(signRow, 2, 1, 3));
  ws.getRange(signRow, 2).setValue('進修部主任：')
    .setFontFamily('標楷體').setFontSize(18).setFontWeight('bold');
  ws.merge(ws.getRange(signRow, 5, 1, 2));
  ws.getRange(signRow, 5).setValue('校長：')
    .setFontFamily('標楷體').setFontSize(18).setFontWeight('bold');
  ws.setRowHeight(signRow, 60);

  var fileId = newSS.getId();
  var downloadUrl = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx';

  return {
    success: true,
    fileName: fileName + '.xlsx',
    downloadUrl: downloadUrl,
    recordCount: records.length
  };
}

// ===== Task 5: 月薪資總表 XLS 生成 =====

function exportSalary(yearStr, monthStr) {
  var settings = getSheet('系統設定').getDataRange().getValues();
  var config = {};
  for (var i = 0; i < settings.length; i++) config[settings[i][0]] = settings[i][1];

  var year = parseInt(yearStr);
  var month = parseInt(monthStr);
  var hourlyRate = parseInt(config['鐘點費單價']);
  var sessionsPerDay = parseInt(config['每日節數']);

  var staffSheet = getSheet('人員名冊');
  var staffData = staffSheet.getDataRange().getValues();
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

  var logSheet = getSheet('教學日誌');
  var logData = logSheet.getDataRange().getValues();
  var teacherStats = {};

  for (var i = 1; i < logData.length; i++) {
    var d = logData[i][0];
    if (!(d instanceof Date)) continue;
    var rocYear = d.getFullYear() - 1911;
    var m = d.getMonth() + 1;
    if (rocYear === year && m === month) {
      var teacher = logData[i][5];
      if (!teacherStats[teacher]) {
        teacherStats[teacher] = { days: 0, dates: [] };
      }
      teacherStats[teacher].days++;
      teacherStats[teacher].dates.push((d.getMonth() + 1) + '/' + d.getDate());
    }
  }

  var westYear = year + 1911;
  var fileName = config['學校名稱'] + config['進修部名稱'] + ' ' + year + '年' + month + '月支給費用';
  var newSS = SpreadsheetApp.create(fileName);
  var ws = newSS.getActiveSheet();

  ws.setColumnWidth(1, 90);
  ws.setColumnWidth(2, 120);
  ws.setColumnWidth(3, 70);
  ws.setColumnWidth(4, 70);
  ws.setColumnWidth(5, 85);
  ws.setColumnWidth(6, 80);
  ws.setColumnWidth(7, 85);
  ws.setColumnWidth(8, 80);
  ws.setColumnWidth(9, 250);

  ws.merge(ws.getRange('A1:I1'));
  ws.getRange('A1').setValue(config['學校名稱'] + config['進修部名稱'] + '  ' + year + '年' + month + '月支給費用')
    .setFontFamily('標楷體').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  ws.setRowHeight(1, 35);

  var headers = ['姓名', '上課期間', '授課天數', '授課節數', '鐘點費單價', '合計', '額外費用', '合計', '備註'];
  ws.getRange(2, 1, 1, 9).setValues([headers])
    .setFontFamily('標楷體').setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);
  ws.setRowHeight(2, 40);

  var lastDay = new Date(westYear, month, 0).getDate();
  var periodStr = month + '月1日～\n' + month + '月' + lastDay + '日';

  var sortedStaff = staffList.sort(function(a, b) {
    var order = { '校長': 0, '導師': 1, '進修部主任': 2, '教師': 3 };
    return (order[a.role] || 9) - (order[b.role] || 9);
  });

  var dataStartRow = 3;
  for (var i = 0; i < sortedStaff.length; i++) {
    var row = dataStartRow + i;
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

    ws.getRange(row, 1).setValue(s.name);
    ws.getRange(row, 2).setValue(periodStr).setWrap(true);
    ws.getRange(row, 3).setValue(days);
    ws.getRange(row, 4).setValue(sessions);
    ws.getRange(row, 5).setValue(rate);
    ws.getRange(row, 6).setValue(hourlyTotal);
    ws.getRange(row, 7).setValue(extraFee).setWrap(true);
    ws.getRange(row, 8).setValue(grandTotal);
    ws.getRange(row, 9).setValue(remark).setWrap(true);

    var dataRange = ws.getRange(row, 1, 1, 9);
    dataRange.setFontFamily('標楷體').setFontSize(14).setFontWeight('bold')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true);
    ws.getRange(row, 1, 1, 8).setHorizontalAlignment('center');
    ws.setRowHeight(row, 55);
  }

  var totalRow = dataStartRow + sortedStaff.length;
  ws.getRange(totalRow, 1).setValue('合  計').setHorizontalAlignment('center');

  var lastDataRow = totalRow - 1;
  ws.getRange(totalRow, 4).setFormula('=SUM(D' + dataStartRow + ':D' + lastDataRow + ')');
  ws.getRange(totalRow, 6).setFormula('=SUM(F' + dataStartRow + ':F' + lastDataRow + ')');
  ws.getRange(totalRow, 8).setFormula('=SUM(H' + dataStartRow + ':H' + lastDataRow + ')');

  var totalRange = ws.getRange(totalRow, 1, 1, 9);
  totalRange.setFontFamily('標楷體').setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);
  ws.setRowHeight(totalRow, 50);

  var signRow = totalRow + 2;
  ws.getRange(signRow, 1).setValue('承辦').setFontFamily('標楷體').setFontSize(14).setFontWeight('bold');
  ws.merge(ws.getRange(signRow, 3, 1, 2));
  ws.getRange(signRow, 3).setValue('出納').setFontFamily('標楷體').setFontSize(14).setFontWeight('bold');
  ws.merge(ws.getRange(signRow, 6, 1, 2));
  ws.getRange(signRow, 6).setValue('會計').setFontFamily('標楷體').setFontSize(14).setFontWeight('bold');
  ws.getRange(signRow, 8).setValue('校長').setFontFamily('標楷體').setFontSize(14).setFontWeight('bold');

  var fileId = newSS.getId();
  var downloadUrl = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx';

  return {
    success: true,
    fileName: fileName + '.xlsx',
    downloadUrl: downloadUrl
  };
}

// ===== Task 6: 薪資條 XLS 生成 =====

function exportPayslip(yearStr, monthStr) {
  var settings = getSheet('系統設定').getDataRange().getValues();
  var config = {};
  for (var i = 0; i < settings.length; i++) config[settings[i][0]] = settings[i][1];

  var year = parseInt(yearStr);
  var month = parseInt(monthStr);
  var hourlyRate = parseInt(config['鐘點費單價']);
  var sessionsPerDay = parseInt(config['每日節數']);
  var westYear = year + 1911;
  var lastDay = new Date(westYear, month, 0).getDate();
  var periodStr = month + '月1日～' + month + '月' + lastDay + '日';

  var staffSheet = getSheet('人員名冊');
  var staffData = staffSheet.getDataRange().getValues();
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

  var logSheet = getSheet('教學日誌');
  var logData = logSheet.getDataRange().getValues();
  var teacherStats = {};
  for (var i = 1; i < logData.length; i++) {
    var d = logData[i][0];
    if (!(d instanceof Date)) continue;
    if (d.getFullYear() - 1911 === year && d.getMonth() + 1 === month) {
      var teacher = logData[i][5];
      if (!teacherStats[teacher]) teacherStats[teacher] = { days: 0, dates: [] };
      teacherStats[teacher].days++;
      teacherStats[teacher].dates.push((d.getMonth() + 1) + '/' + d.getDate());
    }
  }

  var fileName = config['學校名稱'] + '補校支給費用 ' + year + '年' + month + '月薪資條';
  var newSS = SpreadsheetApp.create(fileName);
  var ws = newSS.getActiveSheet();

  ws.setColumnWidth(1, 80);
  ws.setColumnWidth(2, 130);
  ws.setColumnWidth(3, 55);
  ws.setColumnWidth(4, 55);
  ws.setColumnWidth(5, 55);
  ws.setColumnWidth(6, 70);
  ws.setColumnWidth(7, 70);
  ws.setColumnWidth(8, 70);
  ws.setColumnWidth(9, 280);

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

    ws.merge(ws.getRange(currentRow, 1, 1, 9));
    var titleSuffix = '';
    if (isSchoolMaster) {
      titleSuffix = s.extraFeeName;
    } else if (s.extraFeeName) {
      titleSuffix = '鐘點費與' + s.extraFeeName;
    } else {
      titleSuffix = '鐘點費';
    }
    ws.getRange(currentRow, 1).setValue(schoolName + '補校支給費用   ' + year + '年' +
      (month < 10 ? '0' : '') + month + '月' + titleSuffix)
      .setFontFamily('標楷體').setFontSize(13).setFontWeight('bold');
    ws.setRowHeight(currentRow, 28);
    currentRow++;

    if (isSchoolMaster) {
      var h = ['姓名', '上課期間', s.extraFeeName, '', '', '', '', '合計', '備註'];
      ws.getRange(currentRow, 1, 1, 9).setValues([h]);
      ws.merge(ws.getRange(currentRow, 3, 1, 5));
    } else if (s.extraFeeName) {
      var h = ['姓名', '上課期間', '天數', '節數', '單價', '合計', s.extraFeeName, '合計', '備註'];
      ws.getRange(currentRow, 1, 1, 9).setValues([h]);
    } else {
      var h = ['姓名', '上課期間', '天數', '節數', '單價', '合計', '', '', '備註'];
      ws.getRange(currentRow, 1, 1, 9).setValues([h]);
      ws.merge(ws.getRange(currentRow, 6, 1, 3));
    }
    ws.getRange(currentRow, 1, 1, 9).setFontFamily('標楷體').setFontSize(13).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true);
    ws.setRowHeight(currentRow, 55);
    currentRow++;

    if (isSchoolMaster) {
      ws.getRange(currentRow, 1).setValue(s.name);
      ws.getRange(currentRow, 2).setValue(periodStr);
      ws.merge(ws.getRange(currentRow, 3, 1, 5));
      ws.getRange(currentRow, 3).setValue(s.extraFeeAmount);
      ws.getRange(currentRow, 8).setValue(s.extraFeeAmount);
      ws.getRange(currentRow, 9).setValue(s.note);
    } else {
      var days = stats.days;
      var sessions = days * sessionsPerDay;
      var hourlyTotal = sessions * hourlyRate;
      var grandTotal = hourlyTotal + s.extraFeeAmount;

      ws.getRange(currentRow, 1).setValue(s.name);
      ws.getRange(currentRow, 2).setValue(periodStr);
      ws.getRange(currentRow, 3).setValue(days);
      ws.getRange(currentRow, 4).setValue(sessions);
      ws.getRange(currentRow, 5).setValue(hourlyRate);
      ws.getRange(currentRow, 6).setValue(hourlyTotal);
      if (s.extraFeeName) {
        ws.getRange(currentRow, 7).setValue(s.extraFeeAmount);
        ws.getRange(currentRow, 8).setValue(grandTotal);
      } else {
        ws.merge(ws.getRange(currentRow, 6, 1, 3));
        ws.getRange(currentRow, 6).setValue(hourlyTotal);
      }

      var remarkParts = [];
      if (s.note) remarkParts.push(s.note);
      if (stats.dates.length > 0) {
        remarkParts.push('授課日：' + stats.dates.join('、'));
      } else {
        remarkParts.push(month + '月份無授課');
      }
      ws.getRange(currentRow, 9).setValue(remarkParts.join('\n')).setWrap(true);
    }

    ws.getRange(currentRow, 1, 1, 9).setFontFamily('標楷體').setFontSize(13).setFontWeight('bold')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true);
    ws.getRange(currentRow, 1, 1, 8).setHorizontalAlignment('center');
    ws.setRowHeight(currentRow, 65);
    currentRow += 2;
  }

  var fileId = newSS.getId();
  var downloadUrl = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx';

  return {
    success: true,
    fileName: fileName + '.xlsx',
    downloadUrl: downloadUrl
  };
}

// ===== Task 7: 出缺席報表 XLS 生成 =====

function getStudentsInRange(startStr, endStr) {
  var start = new Date(startStr);
  var end = new Date(endStr);
  end.setHours(23, 59, 59);

  var attSheet = getSheet('出缺席記錄');
  var attData = attSheet.getDataRange().getValues();
  var headers = attData[0];

  var studentSet = {};
  for (var i = 1; i < attData.length; i++) {
    var d = attData[i][0];
    if (!(d instanceof Date)) continue;
    if (d >= start && d <= end) {
      for (var c = 3; c < headers.length; c++) {
        if (attData[i][c]) {
          studentSet[headers[c]] = true;
        }
      }
    }
  }

  var studentSheet = getSheet('學生名冊');
  var studentData = studentSheet.getDataRange().getValues();
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
  var settings = getSheet('系統設定').getDataRange().getValues();
  var config = {};
  for (var i = 0; i < settings.length; i++) config[settings[i][0]] = settings[i][1];

  var start = new Date(startStr);
  var end = new Date(endStr);
  end.setHours(23, 59, 59);
  var selectedStudents = studentsStr.split(',');

  var attSheet = getSheet('出缺席記錄');
  var attData = attSheet.getDataRange().getValues();
  var headers = attData[0];

  var classDates = [];
  var dateRecords = {};

  for (var i = 1; i < attData.length; i++) {
    var d = attData[i][0];
    if (!(d instanceof Date)) continue;
    if (d < start || d > end) continue;

    var dateStr = Utilities.formatDate(d, 'Asia/Taipei', 'M/d');
    var weekdays = ['日', '一', '二', '三', '四', '五', '六'];
    var weekday = weekdays[d.getDay()];
    var dateKey = dateStr + '(' + weekday + ')';

    if (!dateRecords[dateKey]) {
      classDates.push({ key: dateKey, date: d });
      dateRecords[dateKey] = {};
    }

    for (var c = 3; c < headers.length; c++) {
      if (attData[i][c]) {
        dateRecords[dateKey][headers[c]] = attData[i][c];
      }
    }
  }

  classDates.sort(function(a, b) { return a.date - b.date; });

  var startRoc = (start.getFullYear() - 1911) + '/' + (start.getMonth() + 1) + '/' + start.getDate();
  var endRoc = (end.getFullYear() - 1911) + '/' + (end.getMonth() + 1) + '/' + end.getDate();
  var fileName = config['縣市名稱'] + config['學校名稱'] + config['進修部名稱'] + ' 出缺席記錄表';
  var newSS = SpreadsheetApp.create(fileName);
  var ws = newSS.getActiveSheet();

  var totalCols = 1 + classDates.length + 2;

  ws.merge(ws.getRange(1, 1, 1, totalCols));
  ws.getRange(1, 1).setValue(config['縣市名稱'] + config['學校名稱'] + config['進修部名稱'] +
    ' 出缺席記錄表（' + startRoc + '～' + endRoc + '）')
    .setFontFamily('標楷體').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  ws.setRowHeight(1, 40);

  var headerRow = ['姓名'];
  for (var i = 0; i < classDates.length; i++) {
    headerRow.push(classDates[i].key);
  }
  headerRow.push('出席天數');
  headerRow.push('出席率');

  ws.getRange(2, 1, 1, totalCols).setValues([headerRow])
    .setFontFamily('標楷體').setFontSize(12).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);
  ws.setRowHeight(2, 35);
  ws.setColumnWidth(1, 80);
  for (var i = 2; i <= classDates.length + 1; i++) ws.setColumnWidth(i, 45);
  ws.setColumnWidth(totalCols - 1, 65);
  ws.setColumnWidth(totalCols, 60);

  for (var si = 0; si < selectedStudents.length; si++) {
    var row = si + 3;
    var name = selectedStudents[si];
    var presentDays = 0;
    var totalDays = classDates.length;

    ws.getRange(row, 1).setValue(name);
    for (var di = 0; di < classDates.length; di++) {
      var status = dateRecords[classDates[di].key][name] || '';
      ws.getRange(row, di + 2).setValue(status).setHorizontalAlignment('center');
      if (status === '✓') presentDays++;
    }
    ws.getRange(row, totalCols - 1).setValue(presentDays).setHorizontalAlignment('center');
    var rate = totalDays > 0 ? Math.round(presentDays / totalDays * 100) + '%' : '0%';
    ws.getRange(row, totalCols).setValue(rate).setHorizontalAlignment('center');

    ws.getRange(row, 1, 1, totalCols)
      .setFontFamily('標楷體').setFontSize(12)
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true);
    ws.setRowHeight(row, 30);
  }

  var fileId = newSS.getId();
  var downloadUrl = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx';

  return {
    success: true,
    fileName: fileName + '.xlsx',
    downloadUrl: downloadUrl
  };
}
