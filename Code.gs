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

function authenticateUser(account, password) {
  account = normalizeValue(account);
  password = normalizeValue(password);

  if (!account || !password) {
    throw new Error('請輸入帳號與密碼。');
  }

  const users = getUsersData_();
  const matchedUser = users.find(function(user) {
    return normalizeValue(user.account) === account && normalizeValue(user.password) === password;
  });

  if (!matchedUser) {
    throw new Error('帳號或密碼錯誤，請重新確認。');
  }

  return {
    user: {
      user_id: matchedUser.user_id,
      project_name: matchedUser.project_name,
      name: matchedUser.name,
      email: matchedUser.email
    },
    summary: getBudgetSummaryByUserId(matchedUser.user_id, matchedUser),
    details: getUserDetails(matchedUser.user_id)
  };
}

function getUsers() {
  return getUsersData_().map(function(user) {
    return {
      user_id: user.user_id,
      project_name: user.project_name,
      name: user.name,
      email: user.email
    };
  });
}

function getBudgetSummary(userId) {
  return getBudgetSummaryByUserId(userId);
}

function getUserBudgetDetails(userId) {
  return getUserDetails(userId);
}

function sendMonthlyBudgetEmails() {
  const users = getUsersData_();
  const summarySheet = getSheetByName_(SHEET_NAMES.SUMMARY);
  const summaryRows = getSheetRecords_(summarySheet, SUMMARY_HEADERS);
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

    const summary = summaryMap[normalizeValue(user.user_id)];
    if (!summary) {
      return;
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
}

function getBudgetSummaryByUserId(userId, userRecord) {
  const normalizedUserId = normalizeValue(userId);
  if (!normalizedUserId) {
    throw new Error('缺少 user_id，無法讀取 summary。');
  }

  const summarySheet = getSheetByName_(SHEET_NAMES.SUMMARY);
  const summaryRows = getSheetRecords_(summarySheet, SUMMARY_HEADERS);
  const summaryRow = summaryRows.find(function(row) {
    return normalizeValue(row.user_id) === normalizedUserId;
  });

  if (!summaryRow) {
    throw new Error('找不到此使用者的預算摘要資料。');
  }

  const user = userRecord || getUsersData_().find(function(row) {
    return normalizeValue(row.user_id) === normalizedUserId;
  });

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
    throw new Error('缺少 user_id，無法讀取明細資料。');
  }

  const detailSheet = getSheetByName_(normalizedUserId);
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
  const sheet = getSheetByName_(SHEET_NAMES.USERS);
  return getSheetRecords_(sheet, USER_HEADERS);
}

function getSheetByName_(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('找不到工作表：' + sheetName);
  }
  return sheet;
}

function getSheetRecords_(sheet, requiredHeaders) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return [];
  }

  const headers = values[0].map(function(header) {
    return normalizeValue(header);
  });

  (requiredHeaders || []).forEach(function(requiredHeader) {
    if (headers.indexOf(requiredHeader) === -1) {
      throw new Error('工作表「' + sheet.getName() + '」缺少欄位：' + requiredHeader);
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
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return [];
  }

  const headerRow = values[0].map(function(header) {
    return normalizeValue(header);
  });
  const useHeaderMapping = DETAIL_HEADERS.every(function(header) {
    return headerRow.indexOf(header) !== -1;
  });

  return values.slice(1).filter(function(row) {
    return row.some(function(cell) {
      return normalizeValue(cell) !== '';
    });
  }).map(function(row) {
    const record = {};
    if (useHeaderMapping) {
      DETAIL_HEADERS.forEach(function(header) {
        record[header] = row[headerRow.indexOf(header)];
      });
    } else {
      DETAIL_HEADERS.forEach(function(header, index) {
        record[header] = row[index];
      });
    }
    return record;
  });
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
