# 國姓國小進修部管理系統改版 Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Rebuild the 國姓國小進修部管理系統 with dynamic data from Google Sheets, admin/teacher permission split, automated salary XLS generation, optimized teaching log XLS layout, and custom-range attendance reports.

**Architecture:** Single-page HTML/JS frontend on GitHub Pages, Google Apps Script (GAS) backend as Web App, one Google Sheet file with multiple tabs for all config and data storage. All XLS reports generated server-side by GAS using Spreadsheet API.

**Tech Stack:** HTML/CSS/JS (no framework), Google Apps Script, Google Sheets API, GitHub Pages

---

## Task 1: Google Sheet 建立與初始資料

**目標：** 建立一個新的 Google Sheet，包含所有必要分頁和初始資料。

**Files:**
- Create: Google Sheet（手動操作，記錄 Sheet ID）

**Step 1: 建立 Google Sheet 並建立分頁**

在 Google Drive 建立新的 Google Sheet，命名為「國姓國小進修部管理系統」。建立以下分頁（Sheet tabs）：

1. `系統設定`
2. `人員名冊`
3. `學生名冊`
4. `課程設定`
5. `教學日誌`
6. `出缺席記錄`
7. `成績設定`
8. `成績記錄`

**Step 2: 填入系統設定**

`系統設定` 分頁：

| A 欄（設定項目） | B 欄（值） |
|---|---|
| 學校名稱 | 國姓國民小學 |
| 進修部名稱 | 進修部 |
| 管理者密碼 | （於 Google Sheet 系統設定中填入） |
| 鐘點費單價 | 405 |
| 每日節數 | 3 |
| 上課時間 | 19:00~21:00 |
| 縣市名稱 | 南投縣 |

**Step 3: 填入人員名冊**

`人員名冊` 分頁：

| 姓名 | 角色 | 狀態 | 額外費用名稱 | 額外費用金額 | 備註 |
|------|------|------|------------|------------|------|
| 林思遠 | 校長 | 在職 | 校長兼職費 | 2333 | 三班以下3500元的三分之二 |
| 吳怡萱 | 導師 | 在職 | 導師費 | 4000 | 比照國民小學導師費標準 |
| 余曜男 | 教師 | 在職 | | | |
| 劉政勳 | 教師 | 在職 | | | |
| 康雲昇 | 教師 | 在職 | | | |

**Step 4: 填入學生名冊**

`學生名冊` 分頁：

| 姓名 | 狀態 |
|------|------|
| 阮氏彫 | 在學 |
| 阮紅妮 | 在學 |
| 阮玄莊 | 在學 |
| 范宥嫺 | 在學 |
| 黎美香 | 在學 |
| 馬銨妤 | 在學 |
| 陳錦江 | 在學 |
| 黎氏銀 | 在學 |
| 范氏燕萍 | 在學 |
| 陳氏錦秀 | 在學 |
| 阮氏雪梅 | 在學 |

**Step 5: 填入課程設定**

`課程設定` 分頁：

| 課程名稱 | 星期 | 授課教師 |
|---------|------|---------|
| 國語與彈性 | 一 | 吳怡萱 |
| 社會生活與彈性 | 二 | 劉政勳 |
| 國語與英文 | 三 | 康雲昇 |
| 數學與科學 | 四 | 余曜男 |

**Step 6: 設定教學日誌表頭**

`教學日誌` 分頁第一行表頭：

| 日期 | 星期 | 時間 | 課程 | 上課內容 | 授課教師 |

**Step 7: 設定出缺席記錄表頭**

`出缺席記錄` 分頁第一行表頭：

| 日期 | 星期 | 課程 | （之後依學生名冊動態展開學生姓名欄） |

注意：學生欄位由 GAS 在寫入時動態處理，表頭初始只需前三欄。

**Step 8: 設定成績相關分頁表頭（預留）**

`成績設定` 分頁：

| 成績科目名稱 | 類型 |
|------------|------|
| 國語 | 學科 |
| 數學 | 學科 |
| 社會 | 學科 |
| 自然 | 學科 |
| 英文 | 學科 |

`成績記錄` 分頁第一行表頭：

| 學年度 | 學期 | 學生姓名 | 科目 | 平時成績 | 考試成績 | 學期成績 |

**Step 9: 記錄 Sheet ID**

複製 Google Sheet 的 URL，從中取出 Sheet ID（URL 中 `/d/{SHEET_ID}/edit` 的部分），記錄到專案中備用。

**驗證：** 開啟 Google Sheet 確認所有 8 個分頁都存在且資料正確。

---

## Task 2: GAS 後端 — 核心框架與公開 API

**目標：** 建立 GAS Web App 骨架，實作 `load_config`、`submit_log`、`submit_attendance` 三個公開 API。

**Files:**
- Create: `gas/Code.gs`（在 Google Apps Script 編輯器中）

**Step 1: 建立 GAS 專案**

在 Google Sheet 中選擇「擴充功能 → Apps Script」，進入 GAS 編輯器。

**Step 2: 寫入 Code.gs 核心框架**

```javascript
// ===== 全域設定 =====
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
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
```

**Step 3: 實作 loadConfig（公開，不含敏感資料）**

```javascript
function loadConfig() {
  // 系統設定（只回傳非敏感欄位）
  var settings = getSheet('系統設定').getDataRange().getValues();
  var config = {};
  for (var i = 0; i < settings.length; i++) {
    var key = settings[i][0];
    if (key === '管理者密碼') continue; // 不回傳密碼
    if (key === '鐘點費單價' || key === '每日節數') continue; // 不回傳費用資訊
    config[key] = settings[i][1];
  }

  // 人員名冊（只回傳姓名和角色，不含費用）
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

  // 學生名冊
  var studentSheet = getSheet('學生名冊');
  var studentData = studentSheet.getDataRange().getValues();
  var students = [];
  for (var i = 1; i < studentData.length; i++) {
    if (studentData[i][1] === '在學') {
      students.push({ name: studentData[i][0], status: studentData[i][1] });
    }
  }

  // 課程設定
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
```

**Step 4: 實作 submitLog**

```javascript
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
```

**Step 5: 實作 submitAttendance**

```javascript
function submitAttendance(data) {
  var sheet = getSheet('出缺席記錄');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // 如果表頭只有基本欄位，需要擴展學生欄位
  var studentNames = Object.keys(data.attendance);
  studentNames.forEach(function(name) {
    if (headers.indexOf(name) === -1) {
      // 新增學生欄
      var nextCol = headers.length + 1;
      sheet.getRange(1, nextCol).setValue(name);
      headers.push(name);
    }
  });

  // 建立該筆記錄
  var row = [data.date, data.weekday, data.course];
  // 填入每位學生的出缺席
  for (var c = 3; c < headers.length; c++) {
    var studentName = headers[c];
    row.push(data.attendance[studentName] || '');
  }

  sheet.appendRow(row);
  return { success: true };
}
```

**Step 6: 部署 GAS Web App**

1. 在 GAS 編輯器中點擊「部署 → 新增部署」
2. 類型選「網頁應用程式」
3. 說明：「進修部管理系統 API v2」
4. 執行身份：「我自己」
5. 存取權限：「任何人」
6. 點擊「部署」
7. 記錄取得的 Web App URL

**驗證：** 在瀏覽器中開啟 `{WEB_APP_URL}?action=load_config`，確認回傳 JSON 包含 staff、students、courses 且不含密碼和費用資訊。

**Step 7: Commit**

```bash
git add -A && git commit -m "docs: record GAS deployment info"
```

---

## Task 3: GAS 後端 — 管理者 API（Dashboard + Admin Config）

**目標：** 實作管理者專用 API：`load_admin_config` 和 `get_dashboard`。

**Files:**
- Modify: `gas/Code.gs`

**Step 1: 實作 loadAdminConfig**

```javascript
function loadAdminConfig() {
  // 完整系統設定
  var settings = getSheet('系統設定').getDataRange().getValues();
  var config = {};
  for (var i = 0; i < settings.length; i++) {
    config[settings[i][0]] = settings[i][1];
  }
  delete config['管理者密碼']; // 密碼不回傳

  // 完整人員名冊（含費用）
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

  // 學生名冊
  var studentSheet = getSheet('學生名冊');
  var studentData = studentSheet.getDataRange().getValues();
  var students = [];
  for (var i = 1; i < studentData.length; i++) {
    students.push({ name: studentData[i][0], status: studentData[i][1] });
  }

  // 課程設定
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
```

**Step 2: 實作 getDashboard**

```javascript
function getDashboard() {
  var today = new Date();
  var todayStr = Utilities.formatDate(today, 'Asia/Taipei', 'yyyy-MM-dd');

  // 今日教學日誌
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

  // 今日出缺席
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
```

**驗證：** 呼叫 `{URL}?action=get_dashboard&pwd={密碼}` 確認回傳今日資料（或空資料）。呼叫 `{URL}?action=load_admin_config&pwd={密碼}` 確認回傳含費用的完整人員資料。

**Step 3: 重新部署 GAS**

更新部署版本（部署 → 管理部署 → 編輯 → 版本選「新版本」→ 部署）。

---

## Task 4: GAS 後端 — 教學日誌 XLS 生成

**目標：** 實作 `export_log`，生成格式優化的教學日誌 XLS，A4 橫向、標楷體、加大簽名欄。

**Files:**
- Modify: `gas/Code.gs`

**Step 1: 實作 exportTeachingLog**

```javascript
function exportTeachingLog(yearStr, monthStr) {
  var settings = getSheet('系統設定').getDataRange().getValues();
  var config = {};
  for (var i = 0; i < settings.length; i++) config[settings[i][0]] = settings[i][1];

  var year = parseInt(yearStr);
  var month = parseInt(monthStr);

  // 篩選該月教學日誌
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

  // 依日期排序
  records.sort(function(a, b) { return a.date - b.date; });

  // 建立新的 Spreadsheet
  var fileName = config['縣市名稱'] + config['學校名稱'] + config['進修部名稱'] +
                 year + '年度' + month + '月教學日誌';
  var newSS = SpreadsheetApp.create(fileName);
  var ws = newSS.getActiveSheet();

  // 設定頁面：A4 橫向
  ws.setColumnWidth(1, 30);   // 序號
  ws.setColumnWidth(2, 100);  // 日期
  ws.setColumnWidth(3, 50);   // 星期
  ws.setColumnWidth(4, 100);  // 時間
  ws.setColumnWidth(5, 120);  // 課程
  ws.setColumnWidth(6, 260);  // 上課內容
  ws.setColumnWidth(7, 200);  // 教師簽名（加大）
  ws.setColumnWidth(8, 90);   // 授課教師

  // Row 1: 標題
  ws.merge(ws.getRange('A1:H1'));
  var titleCell = ws.getRange('A1');
  titleCell.setValue(config['縣市名稱'] + config['學校名稱'] + config['進修部名稱'] +
                     ' ' + year + '年度' + month + '月 教學日誌');
  titleCell.setFontFamily('標楷體').setFontSize(20).setFontWeight('bold')
    .setHorizontalAlignment('center');
  ws.setRowHeight(1, 50);

  // Row 2: 表頭
  var headers = ['序號', '日期', '星期', '時間', '課程', '上課內容', '教師簽名', '授課教師'];
  var headerRange = ws.getRange(2, 1, 1, 8);
  headerRange.setValues([headers]);
  headerRange.setFontFamily('標楷體').setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  headerRange.setBorder(true, true, true, true, true, true);
  ws.setRowHeight(2, 55);

  // 資料列
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
    // G 欄（教師簽名）留空
    ws.getRange(row, 8).setValue(r.teacher);

    var dataRange = ws.getRange(row, 1, 1, 8);
    dataRange.setFontFamily('標楷體').setFontSize(14)
      .setVerticalAlignment('middle');
    dataRange.setBorder(true, true, true, true, true, true);
    ws.setRowHeight(row, 55);

    // 序號、日期、星期、時間置中
    ws.getRange(row, 1, 1, 4).setHorizontalAlignment('center');
    // 授課教師置中
    ws.getRange(row, 8).setHorizontalAlignment('center');
  }

  // 簽章區（資料下方空兩行）
  var signRow = records.length + 5;
  ws.merge(ws.getRange(signRow, 2, 1, 3));
  ws.getRange(signRow, 2).setValue('進修部主任：')
    .setFontFamily('標楷體').setFontSize(18).setFontWeight('bold');
  ws.merge(ws.getRange(signRow, 5, 1, 2));
  ws.getRange(signRow, 5).setValue('校長：')
    .setFontFamily('標楷體').setFontSize(18).setFontWeight('bold');
  ws.setRowHeight(signRow, 60);

  // 取得下載連結
  var fileId = newSS.getId();
  var downloadUrl = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx';

  return {
    success: true,
    fileName: fileName + '.xlsx',
    downloadUrl: downloadUrl,
    recordCount: records.length
  };
}
```

**驗證：** 在教學日誌 Sheet 中手動加幾筆測試資料，呼叫 `{URL}?action=export_log&year=114&month=11&pwd={密碼}`，下載 XLS 確認：標楷體、欄寬合理、教師簽名欄加大、簽章只有進修部主任和校長。

**Step 2: 重新部署 GAS 並 Commit**

---

## Task 5: GAS 後端 — 月薪資總表 XLS 生成

**目標：** 實作 `export_salary`，自動從教學日誌計算每位教師授課天數、鐘點費，加上校長兼職費。

**Files:**
- Modify: `gas/Code.gs`

**Step 1: 實作 exportSalary**

```javascript
function exportSalary(yearStr, monthStr) {
  var settings = getSheet('系統設定').getDataRange().getValues();
  var config = {};
  for (var i = 0; i < settings.length; i++) config[settings[i][0]] = settings[i][1];

  var year = parseInt(yearStr);
  var month = parseInt(monthStr);
  var hourlyRate = parseInt(config['鐘點費單價']);
  var sessionsPerDay = parseInt(config['每日節數']);

  // 讀取人員名冊
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

  // 讀取教學日誌，統計每位教師該月授課天數和日期
  var logSheet = getSheet('教學日誌');
  var logData = logSheet.getDataRange().getValues();
  var teacherStats = {}; // { name: { days: N, dates: ['11/3', '11/10', ...] } }

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

  // 建立 XLS
  var westYear = year + 1911;
  var fileName = config['學校名稱'] + config['進修部名稱'] + ' ' + year + '年' + month + '月支給費用';
  var newSS = SpreadsheetApp.create(fileName);
  var ws = newSS.getActiveSheet();

  // 欄寬設定
  ws.setColumnWidth(1, 90);   // 姓名
  ws.setColumnWidth(2, 120);  // 上課期間
  ws.setColumnWidth(3, 70);   // 授課天數
  ws.setColumnWidth(4, 70);   // 授課節數
  ws.setColumnWidth(5, 85);   // 鐘點費單價
  ws.setColumnWidth(6, 80);   // 合計
  ws.setColumnWidth(7, 85);   // 額外費用
  ws.setColumnWidth(8, 80);   // 合計
  ws.setColumnWidth(9, 250);  // 備註

  // Row 1: 標題
  ws.merge(ws.getRange('A1:I1'));
  ws.getRange('A1').setValue(config['學校名稱'] + config['進修部名稱'] + '  ' + year + '年' + month + '月支給費用')
    .setFontFamily('標楷體').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  ws.setRowHeight(1, 35);

  // Row 2: 表頭
  var headers = ['姓名', '上課期間', '授課天數', '授課節數', '鐘點費單價', '合計', '額外費用', '合計', '備註'];
  ws.getRange(2, 1, 1, 9).setValues([headers])
    .setFontFamily('標楷體').setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);
  ws.setRowHeight(2, 40);

  // 計算月份期間字串
  var lastDay = new Date(westYear, month, 0).getDate();
  var periodStr = month + '月1日～\n' + month + '月' + lastDay + '日';

  // 資料列：校長排第一，其他教師依序
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

  // 合計列
  var totalRow = dataStartRow + sortedStaff.length;
  ws.getRange(totalRow, 1).setValue('合  計').setHorizontalAlignment('center');

  // 用 SUM 公式
  var lastDataRow = totalRow - 1;
  ws.getRange(totalRow, 4).setFormula('=SUM(D' + dataStartRow + ':D' + lastDataRow + ')');
  ws.getRange(totalRow, 6).setFormula('=SUM(F' + dataStartRow + ':F' + lastDataRow + ')');
  ws.getRange(totalRow, 8).setFormula('=SUM(H' + dataStartRow + ':H' + lastDataRow + ')');

  var totalRange = ws.getRange(totalRow, 1, 1, 9);
  totalRange.setFontFamily('標楷體').setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true);
  ws.setRowHeight(totalRow, 50);

  // 簽章區
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
```

**驗證：** 呼叫 `export_salary`，下載 XLS 確認：校長在第一列只有兼職費、其他教師有鐘點計算和備註授課日、合計列正確。

---

## Task 6: GAS 後端 — 薪資條 XLS 生成

**目標：** 實作 `export_payslip`，每人一個區塊，校長/導師/教師各有不同標題但統一欄位格式。

**Files:**
- Modify: `gas/Code.gs`

**Step 1: 實作 exportPayslip**

```javascript
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

  // 讀取人員名冊
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

  // 教學日誌統計
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

  // 排序：校長 → 導師 → 進修部主任 → 教師
  var sortedStaff = staffList.sort(function(a, b) {
    var order = { '校長': 0, '導師': 1, '進修部主任': 2, '教師': 3 };
    return (order[a.role] || 9) - (order[b.role] || 9);
  });

  for (var si = 0; si < sortedStaff.length; si++) {
    var s = sortedStaff[si];
    var stats = teacherStats[s.name] || { days: 0, dates: [] };
    var isSchoolMaster = (s.role === '校長');

    // 區塊標題
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

    // 表頭
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

    // 資料列
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

      // 備註：第一行放費用說明，第二行放授課日
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
    currentRow += 2; // 區塊間留一行空白
  }

  var fileId = newSS.getId();
  var downloadUrl = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx';

  return {
    success: true,
    fileName: fileName + '.xlsx',
    downloadUrl: downloadUrl
  };
}
```

**驗證：** 呼叫 `export_payslip`，確認每人一個區塊、校長只有兼職費、導師有鐘點+導師費、一般教師只有鐘點費。

---

## Task 7: GAS 後端 — 出缺席報表 XLS 生成

**目標：** 實作 `getStudentsInRange` 和 `exportAttendance`。

**Files:**
- Modify: `gas/Code.gs`

**Step 1: 實作 getStudentsInRange**

```javascript
function getStudentsInRange(startStr, endStr) {
  var start = new Date(startStr);
  var end = new Date(endStr);
  end.setHours(23, 59, 59);

  var attSheet = getSheet('出缺席記錄');
  var attData = attSheet.getDataRange().getValues();
  var headers = attData[0];

  // 找出該期間內有記錄的學生
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

  // 同時也從學生名冊載入所有在學學生
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
```

**Step 2: 實作 exportAttendance**

```javascript
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

  // 收集期間內所有上課日期
  var classDates = [];
  var dateRecords = {}; // { dateStr: { student: status } }

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

  // 排序日期
  classDates.sort(function(a, b) { return a.date - b.date; });

  // 建立 XLS
  var startRoc = (start.getFullYear() - 1911) + '/' + (start.getMonth() + 1) + '/' + start.getDate();
  var endRoc = (end.getFullYear() - 1911) + '/' + (end.getMonth() + 1) + '/' + end.getDate();
  var fileName = config['縣市名稱'] + config['學校名稱'] + config['進修部名稱'] + ' 出缺席記錄表';
  var newSS = SpreadsheetApp.create(fileName);
  var ws = newSS.getActiveSheet();

  var totalCols = 1 + classDates.length + 2; // 姓名 + 日期 + 出席天數 + 出席率

  // Row 1: 標題
  ws.merge(ws.getRange(1, 1, 1, totalCols));
  ws.getRange(1, 1).setValue(config['縣市名稱'] + config['學校名稱'] + config['進修部名稱'] +
    ' 出缺席記錄表（' + startRoc + '～' + endRoc + '）')
    .setFontFamily('標楷體').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  ws.setRowHeight(1, 40);

  // Row 2: 表頭
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

  // 資料列
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
```

**驗證：** 先手動在出缺席 Sheet 加測試資料，呼叫 `get_students_in_range` 確認回傳學生列表，再呼叫 `export_attendance` 確認 XLS 格式正確。

---

## Task 8: 前端 — 新版 index.html 骨架與權限系統

**目標：** 重寫 index.html，實作教師/管理者模式切換、動態載入 config。

**Files:**
- Modify: `index.html`

**Step 1: 建立 HTML 骨架**

完全重寫 `index.html`，包含：
- 全域 CSS（沿用現有深色主題風格）
- 教師模式：只顯示教學日誌 + 出缺席兩個 tab
- 管理者模式：顯示全部 5 個 tab（總覽/日誌/出缺席/報表/成績）
- 右上角管理者登入按鈕
- JavaScript：GAS API URL 設定、頁面初始化、權限切換邏輯
- `sessionStorage` 管理者 session

**Step 2: 實作動態載入**

頁面載入時呼叫 `load_config` API：
- 教師下拉選單從 `staff` 動態填入
- 學生出缺席列表從 `students` 動態填入
- 課程下拉選單從 `courses` 動態填入
- 根據選擇的日期自動建議星期和課程

**Step 3: 實作管理者登入**

- 點擊「管理者登入」→ 彈出密碼輸入框
- 呼叫 `verify_admin` API 驗證
- 成功後存入 `sessionStorage`，切換到管理者模式
- 切換後呼叫 `load_admin_config` 取得完整資料

**驗證：** 在瀏覽器開啟 index.html，確認教師模式只看到兩個 tab，輸入密碼後看到全部功能。

**Step 4: Commit**

```bash
git add index.html && git commit -m "feat: rebuild index.html with permission system and dynamic config loading"
```

---

## Task 9: 前端 — 教學日誌填報模組

**目標：** 重寫教學日誌表單，動態載入教師/課程選單，支援語音輸入。

**Files:**
- Modify: `index.html`（教學日誌 tab 區塊）
- Delete: `teaching_log.html`（合併進 index.html，不再用 iframe）

**Step 1: 實作教學日誌表單**

在 index.html 的教學日誌 tab 內建立：
- 日期選擇器（自動帶入今天）
- 星期顯示（自動計算）
- 時間顯示（從 config 載入）
- 課程下拉選單（動態）— 選擇課程後自動帶入對應教師
- 上課內容文字區（含語音輸入按鈕 + 快速用語）
- 授課教師下拉選單（動態，選課程後自動選取）
- 提交按鈕

**Step 2: 實作提交邏輯**

表單提交 → POST 到 GAS `submit_log` → 顯示成功/失敗訊息。

**驗證：** 填寫表單並提交，確認 Google Sheet 教學日誌分頁新增了一筆資料。

**Step 3: Commit**

```bash
git add index.html && git commit -m "feat: add dynamic teaching log form with voice input"
```

---

## Task 10: 前端 — 出缺席填報模組

**目標：** 重寫出缺席表單，動態載入學生名單。

**Files:**
- Modify: `index.html`（出缺席 tab 區塊）
- Delete: `attendance.html`（合併進 index.html）

**Step 1: 實作出缺席表單**

在 index.html 的出缺席 tab 內建立：
- 日期選擇器
- 星期顯示
- 課程選單（動態）
- 學生出缺席表格（動態從 students 生成）
- 出席/請假 checkbox（互斥邏輯）
- 提交按鈕

**Step 2: 實作提交邏輯**

POST 到 GAS `submit_attendance`。

**驗證：** 填報出缺席並提交，確認 Google Sheet 出缺席記錄分頁新增資料。

**Step 3: Commit**

```bash
git add index.html && git commit -m "feat: add dynamic attendance form"
```

---

## Task 11: 前端 — 管理者總覽儀表板

**目標：** 實作 Dashboard，顯示今日出缺席率和教學日誌。

**Files:**
- Modify: `index.html`（總覽 tab 區塊）

**Step 1: 實作 Dashboard UI**

- 三個統計卡片：出席率、出席人數、請假人數
- 今日教學日誌區塊：課程、教師、內容、時間
- 無資料時顯示「今日尚無記錄」

**Step 2: 載入 Dashboard 資料**

管理者登入後呼叫 `get_dashboard` API，動態填入卡片數據。

**驗證：** 管理者登入後確認 Dashboard 顯示今日資料（或「今日尚無記錄」）。

**Step 3: Commit**

```bash
git add index.html && git commit -m "feat: add admin dashboard with today's stats"
```

---

## Task 12: 前端 — 報表下載中心

**目標：** 實作報表下載 UI，包含教學日誌/薪資總表/薪資條/出缺席報表四種下載。

**Files:**
- Modify: `index.html`（報表 tab 區塊）

**Step 1: 實作教學日誌 & 薪資報表下載 UI**

- 選擇年份（民國）+ 月份
- 三個下載按鈕：教學日誌 XLS、薪資總表 XLS、薪資條 XLS
- 點擊後呼叫對應 GAS API，顯示「處理中」，完成後開啟下載連結

**Step 2: 實作出缺席報表下載 UI**

- 選擇起始日期 + 結束日期
- 點「查詢學生」→ 呼叫 `get_students_in_range`
- 顯示學生勾選清單（預設全選）
- 點「生成報表」→ 呼叫 `export_attendance`

**Step 3: 下載邏輯**

所有 export API 回傳 `downloadUrl`，前端用 `window.open(url)` 開啟下載。

**驗證：** 分別測試四種報表下載，確認 XLS 格式和內容正確。

**Step 4: Commit**

```bash
git add index.html && git commit -m "feat: add report download center with all 4 report types"
```

---

## Task 13: 前端 — 成績管理預留 + 清理舊檔

**目標：** 加入灰色的成績管理 tab、清理不再使用的舊檔案。

**Files:**
- Modify: `index.html`
- Delete: `teaching_log.html`
- Delete: `attendance.html`

**Step 1: 加入成績管理預留 tab**

在管理者模式的 nav 中加入「📚 成績管理」按鈕，灰色、不可點擊、顯示「開發中」標籤。

**Step 2: 清理舊檔案**

刪除已合併進 index.html 的舊檔案：
- `teaching_log.html`
- `attendance.html`

保留：
- `manual.html`（使用手冊，獨立頁面）
- `docs/`（設計文件）

**Step 3: Commit**

```bash
git rm teaching_log.html attendance.html
git add index.html
git commit -m "feat: add grades placeholder tab, remove merged legacy files"
```

---

## Task 14: 整合測試與部署

**目標：** 完整端到端測試，推送到 GitHub。

**Step 1: 端到端測試清單**

- [ ] 開啟頁面 → 教師模式，只看到日誌和出缺席
- [ ] 提交教學日誌 → Sheet 新增資料
- [ ] 提交出缺席 → Sheet 新增資料
- [ ] 管理者登入（錯誤密碼被拒絕）
- [ ] 管理者登入（正確密碼成功）
- [ ] Dashboard 顯示今日資料
- [ ] 下載教學日誌 XLS → A4 排版正確、標楷體、簽名欄夠大
- [ ] 下載薪資總表 XLS → 校長+教師都在、自動計算正確
- [ ] 下載薪資條 XLS → 每人一區塊、格式正確
- [ ] 下載出缺席報表 XLS → 勾選學生有效、日期區間正確
- [ ] 手機瀏覽 → 響應式正常
- [ ] 關閉瀏覽器重開 → 管理者 session 已失效

**Step 2: 修正測試中發現的問題**

逐一修正。

**Step 3: Push to GitHub**

```bash
git push origin main
```

**驗證：** GitHub Pages 上線後，完整跑一次測試清單。
