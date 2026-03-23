const SPREADSHEET_ID = '1Upk2sB_jL-olofQw9p6-68npCtqcvX1v0lYAW0ZuEOo';
const SHEET_NAMES = {
  USERS: 'users',
  SUMMARY: 'budget_summary'
};

const DETAIL_HEADERS = ['實支', '日期', '傳票', '類別', '購案編號', '金額', '說明'];
const SUMMARY_HEADERS = ['user_id', 'project_name', '業務費', '實支+未核銷', '餘額', '支用率'];
const USER_HEADERS = ['user_id', 'project_name', 'account', 'password', 'name', 'email'];

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('個人預算查詢網站')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (error) {
    throw createAppError_('無法開啟 Google Sheet，請確認 SPREADSHEET_ID 與權限設定。', error);
  }
}

function authenticateUser(account, password) {
  return runSafely_(function() {
    account = normalizeValue(account);
    password = normalizeValue(password);

    if (!account || !password) {
      throw createAppError_('請輸入帳號與密碼。');
    }

    const users = getUsersData_();
    const matchedUser = users.find(function(user) {
      return normalizeValue(user.account) === account && normalizeValue(user.password) === password;
    });

    if (!matchedUser) {
      throw createAppError_('帳號或密碼錯誤，請重新確認。');
    }

    return buildUserSession_(matchedUser);
  });
}

function getUsers() {
  return runSafely_(function() {
    return getUsersData_().map(function(user) {
      return {
        user_id: user.user_id,
        project_name: user.project_name,
        name: user.name,
        email: user.email
      };
    });
  });
}

function getBudgetSummary(userId) {
  return runSafely_(function() {
    return getBudgetSummaryByUserId(userId);
  });
}

function getUserBudgetDetails(userId) {
  return runSafely_(function() {
    return getUserDetails(userId);
  });
}

function getUserDashboard(userId) {
  return runSafely_(function() {
    const user = getUserById_(userId);
    return buildUserSession_(user);
  });
}

function sendMonthlyBudgetEmails() {
  return runSafely_(function() {
    const users = getUsersData_();
    const summaryRows = getBudgetSummaryRecords_();
    const summaryMap = {};

    summaryRows.forEach(function(row) {
      const key = normalizeValue(row.user_id);
      if (key) {
        summaryMap[key] = row;
      }
    });

    users.forEach(function(user) {
      const email = normalizeValue(user.email);
      if (!email) {
        return;
      }

      const userId = normalizeValue(user.user_id);
      const summary = summaryMap[userId];
      if (!summary) {
        throw createAppError_('找不到使用者 ' + userId + ' 的 budget_summary 資料，無法寄送月報。');
      }

      const projectName = normalizeValue(summary.project_name) || normalizeValue(user.project_name);
      const body = [
        user.name + ' 您好：',
        '',
        '以下是您本月的預算使用情況：',
        '使用者姓名：' + displayText_(user.name),
        '使用者 ID：' + displayText_(user.user_id),
        '計畫名稱：' + displayText_(projectName),
        '業務費：' + formatCellValue_(summary['業務費']),
        '實支+未核銷：' + formatCellValue_(summary['實支+未核銷']),
        '餘額：' + formatCellValue_(summary['餘額']),
        '支用率：' + formatRate_(summary['支用率']),
        '',
        '如需確認明細，請登入個人預算查詢網站查看。',
        '',
        '此為系統自動寄送通知，敬請參考。'
      ].join('\n');

      MailApp.sendEmail({
        to: email,
        subject: '【每月預算通知】您的預算使用情況',
        body: body
      });
    });

    return '每月預算通知已寄送完成，共寄送 ' + users.filter(function(user) {
      return normalizeValue(user.email) !== '';
    }).length + ' 筆。';
  });
}

function buildUserSession_(userRecord) {
  const userId = normalizeValue(userRecord.user_id);
  return {
    user: {
      user_id: userId,
      project_name: userRecord.project_name,
      name: userRecord.name,
      email: userRecord.email
    },
    summary: getBudgetSummaryByUserId(userId, userRecord),
    details: getUserDetails(userId)
  };
}

function getBudgetSummaryByUserId(userId, userRecord) {
  const normalizedUserId = normalizeValue(userId);
  if (!normalizedUserId) {
    throw createAppError_('缺少 user_id，無法讀取 budget_summary。');
  }

  const summaryRows = getBudgetSummaryRecords_();
  const summaryRow = summaryRows.find(function(row) {
    return normalizeValue(row.user_id) === normalizedUserId;
  });

  if (!summaryRow) {
    throw createAppError_('找不到使用者 ' + normalizedUserId + ' 的 budget_summary 資料。');
  }

  const user = userRecord || getUserById_(normalizedUserId);

  return {
    user_id: normalizedUserId,
    name: user ? user.name : '',
    project_name: normalizeValue(summaryRow.project_name) || (user ? user.project_name : ''),
    業務費: formatCellValue_(summaryRow['業務費']),
    '實支+未核銷': formatCellValue_(summaryRow['實支+未核銷']),
    餘額: formatCellValue_(summaryRow['餘額']),
    支用率: formatRate_(summaryRow['支用率'])
  };
}

function getUserDetails(userId) {
  const normalizedUserId = normalizeValue(userId);
  if (!normalizedUserId) {
    throw createAppError_('缺少 user_id，無法讀取個人明細。');
  }

  const detailSheet = getUserDetailSheet_(normalizedUserId);
  const records = getDetailRecords_(detailSheet);

  return records.sort(function(a, b) {
    return parseDateValue_(b.日期) - parseDateValue_(a.日期);
  }).map(function(record) {
    const fullDescription = displayText_(record['說明']);
    return {
      實支: displayText_(record['實支']),
      日期: formatDateValue_(record['日期']),
      傳票: displayText_(record['傳票']),
      類別: displayText_(record['類別']),
      購案編號: displayText_(record['購案編號']),
      金額: formatCellValue_(record['金額']),
      說明: fullDescription,
      說明簡短: shortenDescription(fullDescription),
      狀態: getVoucherStatus(record['傳票'])
    };
  });
}

function getVoucherStatus(voucherValue) {
  return normalizeValue(voucherValue) ? '已開傳票' : '未審';
}

function shortenDescription(description, maxLength) {
  const text = displayText_(description);
  const limit = Number(maxLength) || 24;
  if (text.length <= limit) {
    return text;
  }
  return text.slice(0, limit) + '…';
}

function getUsersData_() {
  const sheet = getUsersSheet_();
  return getSheetRecords_(sheet, USER_HEADERS);
}

function getBudgetSummaryRecords_() {
  const sheet = getBudgetSummarySheet_();
  return getSheetRecords_(sheet, SUMMARY_HEADERS);
}

function getUserById_(userId) {
  const normalizedUserId = normalizeValue(userId);
  if (!normalizedUserId) {
    throw createAppError_('缺少 user_id，無法讀取 users 資料。');
  }

  const user = getUsersData_().find(function(row) {
    return normalizeValue(row.user_id) === normalizedUserId;
  });

  if (!user) {
    throw createAppError_('找不到使用者資料：' + normalizedUserId);
  }

  return user;
}

function getUsersSheet_() {
  return getRequiredSheet_(SHEET_NAMES.USERS, '找不到 users 分頁');
}

function getBudgetSummarySheet_() {
  return getRequiredSheet_(SHEET_NAMES.SUMMARY, '找不到 budget_summary 分頁');
}

function getUserDetailSheet_(userId) {
  return getRequiredSheet_(userId, '找不到使用者分頁：' + userId);
}

function getRequiredSheet_(sheetName, errorMessage) {
  const spreadsheet = getSpreadsheet();
  if (!spreadsheet) {
    throw createAppError_('無法取得 Google Sheet 物件。');
  }

  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw createAppError_(errorMessage || ('找不到分頁：' + sheetName));
  }
  return sheet;
}

function getSheetRecords_(sheet, requiredHeaders) {
  if (!sheet) {
    throw createAppError_('讀取工作表失敗：sheet 不存在。');
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return [];
  }

  const headers = values[0].map(function(header) {
    return normalizeValue(header);
  });

  (requiredHeaders || []).forEach(function(requiredHeader) {
    if (headers.indexOf(requiredHeader) === -1) {
      throw createAppError_('工作表「' + sheet.getName() + '」缺少欄位：' + requiredHeader);
    }
  });

  return values.slice(1).filter(function(row) {
    return row.some(function(cell) {
      return normalizeValue(cell) !== '';
    });
  }).map(function(row) {
    const record = {};
    headers.forEach(function(header, index) {
      record[header] = row[index];
    });
    return record;
  });
}

function getDetailRecords_(sheet) {
  if (!sheet) {
    throw createAppError_('讀取個人明細失敗：sheet 不存在。');
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return [];
  }

  const headerRow = values[0].map(function(header) {
    return normalizeValue(header);
  });
  const missingHeaders = DETAIL_HEADERS.filter(function(header) {
    return headerRow.indexOf(header) === -1;
  });

  if (missingHeaders.length) {
    throw createAppError_('工作表「' + sheet.getName() + '」缺少欄位：' + missingHeaders.join('、'));
  }

  return values.slice(1).filter(function(row) {
    return row.some(function(cell) {
      return normalizeValue(cell) !== '';
    });
  }).map(function(row) {
    const record = {};
    DETAIL_HEADERS.forEach(function(header) {
      record[header] = row[headerRow.indexOf(header)];
    });
    return record;
  });
}

function runSafely_(callback) {
  try {
    return callback();
  } catch (error) {
    throw toUserError_(error);
  }
}

function createAppError_(message, originalError) {
  const error = new Error(message);
  error.name = 'AppError';
  if (originalError) {
    error.originalError = originalError;
    if (originalError.stack) {
      error.stack = originalError.stack;
    }
  }
  return error;
}

function toUserError_(error) {
  if (error && error.name === 'AppError') {
    return error;
  }

  const message = error && error.message
    ? '系統發生錯誤：' + error.message
    : '系統發生錯誤，請稍後再試。';
  return createAppError_(message, error);
}

function normalizeValue(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function displayText_(value) {
  return normalizeValue(value);
}

function formatCellValue_(value) {
  if (value === null || value === undefined || value === '') {
    return '';
  }
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return formatDateValue_(value);
  }
  if (typeof value === 'number') {
    return Utilities.formatString('%s', value % 1 === 0 ? value.toLocaleString('en-US') : value.toLocaleString('en-US', { maximumFractionDigits: 2 }));
  }
  return String(value);
}

function formatRate_(value) {
  if (value === null || value === undefined || value === '') {
    return '';
  }
  if (typeof value === 'number') {
    if (value <= 1) {
      return (value * 100).toFixed(2).replace(/\.00$/, '') + '%';
    }
    return value.toFixed(2).replace(/\.00$/, '') + '%';
  }
  return String(value);
}

function formatDateValue_(value) {
  const date = parseDateValue_(value);
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return displayText_(value);
  }
  return Utilities.formatDate(date, Session.getScriptTimeZone() || 'Asia/Taipei', 'yyyy/MM/dd');
}

function parseDateValue_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value;
  }
  if (typeof value === 'number') {
    return new Date(Math.round((value - 25569) * 86400 * 1000));
  }
  const text = normalizeValue(value);
  const date = text ? new Date(text) : new Date('1970-01-01');
  return isNaN(date.getTime()) ? new Date('1970-01-01') : date;
}
