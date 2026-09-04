const SHEET_NAME = 'طلبات العمل بالحصة';
const ATTACHMENTS_FOLDER_NAME = 'مرفقات طلبات العمل بالحصة';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('طلب للعمل بنظام الحصة')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOutput_({ success: false, message: 'صيغة الطلب غير صحيحة.' });
  }

  try {
    let result;
    if (payload.action === 'save') {
      result = saveRecord(payload.data, payload.attachments);
    } else if (payload.action === 'query') {
      result = queryRecord(payload.nationalId);
    } else {
      throw new Error('إجراء غير معروف: ' + payload.action);
    }
    return jsonOutput_({ success: true, result: result });
  } catch (err) {
    return jsonOutput_({ success: false, message: err.message || 'حدث خطأ غير متوقع.' });
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('يجب ربط المشروع بملف Google Sheets.');

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  const headers = [
    'المعرف',
    'الاسم رباعي',
    'تاريخ الميلاد',
    'الرقم القومي',
    'المؤهل الدراسي',
    'التخصص',
    'التليفون المحمول',
    'محل الإقامة تفصيليًا',
    'المادة',
    'المرفقات',
    'تاريخ الحفظ',
    'الحالة',          // جديد
    'سبب الرفض'        // جديد
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    // إذا تغيرت الأعمدة، نقوم بتحديثها
    if (current.length < headers.length || current.some((h, i) => h !== headers[i])) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1261a0')
    .setFontColor('#ffffff');
  sheet.autoResizeColumns(1, headers.length);

  getAttachmentsFolder_();
  return 'تم تجهيز الورقة ومجلد المرفقات بنجاح.';
}

function saveRecord(data, attachments) {
  validateData_(data);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('لم يتم العثور على ملف Google Sheets المرتبط.');

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    setup();
    sheet = ss.getSheetByName(SHEET_NAME);
  }

  const nationalId = String(data.nationalId || '').trim();
  const lastRow = sheet.getLastRow();
  let row = -1;
  let isUpdate = false;
  let existingStatus = '';
  let existingReason = '';

  if (lastRow >= 2) {
    const ids = sheet.getRange(2, 4, lastRow - 1, 1).getDisplayValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === nationalId) {
        row = i + 2;
        // جلب الحالة الحالية وسبب الرفض (إن وجدا) للحفاظ عليهما
        const statusRange = sheet.getRange(row, 12);
        const reasonRange = sheet.getRange(row, 13);
        existingStatus = statusRange.getDisplayValue();
        existingReason = reasonRange.getDisplayValue();
        break;
      }
    }
  }

  const existingAttachmentCell = row > 0 ? sheet.getRange(row, 10).getDisplayValue() : '';
  const attachmentLinks = saveAttachments_(attachments || []);
  const allAttachments = existingAttachmentCell
    ? existingAttachmentCell + (attachmentLinks ? '\n' + attachmentLinks : '')
    : attachmentLinks;

  const recordId = row > 0
    ? sheet.getRange(row, 1).getValue()
    : Utilities.getUuid();

  const values = [[
    recordId,
    String(data.name || '').trim(),
    String(data.birthDate || '').trim(),
    nationalId,
    String(data.qualification || '').trim(),
    String(data.specialization || '').trim(),
    String(data.phone || '').trim(),
    String(data.address || '').trim(),
    String(data.subject || '').trim(),
    allAttachments,
    new Date(),
    existingStatus,   // نُبقي الحالة القديمة
    existingReason    // نُبقي سبب الرفض القديم
  ]];

  if (row > 0) {
    isUpdate = true;
    sheet.getRange(row, 1, 1, values[0].length).setValues(values);
  } else {
    sheet.appendRow(values[0]);
    row = sheet.getLastRow();
  }

  sheet.getRange(row, 3).setNumberFormat('@');
  sheet.getRange(row, 4).setNumberFormat('@');
  sheet.getRange(row, 10).setWrap(true);
  sheet.getRange(row, 11).setNumberFormat('yyyy-mm-dd hh:mm:ss');

  return {
    success: true,
    message: isUpdate ? 'تم تحديث البيانات بنجاح.' : 'تم حفظ البيانات بنجاح.',
    recordId: recordId,
    row: row,
    attachmentLinks: attachmentLinks
  };
}

function queryRecord(nationalId) {
  nationalId = String(nationalId || '').trim();
  if (!nationalId) throw new Error('أدخل الرقم القومي أولًا.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    return { found: false, message: 'لا توجد بيانات محفوظة حتى الآن.' };
  }

  // نقرأ جميع الأعمدة (حتى العمود 13)
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getDisplayValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][3]).trim() === nationalId) {
      return {
        found: true,
        data: {
          recordId: data[i][0],
          name: data[i][1],
          birthDate: normalizeDateForInput_(data[i][2]),
          nationalId: data[i][3],
          qualification: data[i][4],
          specialization: data[i][5],
          phone: data[i][6],
          address: data[i][7],
          subject: data[i][8],
          attachments: data[i][9],
          savedAt: data[i][10],
          status: data[i][11] || '',        // الحالة
          rejectionReason: data[i][12] || '' // سبب الرفض
        }
      };
    }
  }

  return { found: false, message: 'لا توجد بيانات محفوظة لهذا الرقم القومي.' };
}

function validateData_(data) {
  if (!data || !String(data.name || '').trim()) {
    throw new Error('يرجى إدخال الاسم رباعي.');
  }
  if (!String(data.birthDate || '').trim()) {
    throw new Error('يرجى اختيار تاريخ الميلاد.');
  }
  const nationalId = String(data.nationalId || '').trim();
  if (!/^\d{14}$/.test(nationalId)) {
    throw new Error('الرقم القومي يجب أن يتكون من 14 رقمًا.');
  }
}

function saveAttachments_(attachments) {
  if (!attachments || !attachments.length) return '';

  const folder = getAttachmentsFolder_();
  const links = [];

  attachments.forEach(file => {
    if (!file || !file.name || !file.base64) return;

    const name = String(file.name);
    const mimeType = String(file.mimeType || 'application/octet-stream');
    if (!/^(application\/pdf|image\/jpeg)$/i.test(mimeType) && !/\.(pdf|jpe?g)$/i.test(name)) {
      throw new Error('المرفق غير مسموح: ' + name + '. المسموح PDF وJPG/JPEG فقط.');
    }

    const bytes = Utilities.base64Decode(String(file.base64).split(',').pop());
    const blob = Utilities.newBlob(bytes, mimeType, name);
    const created = folder.createFile(blob);
    created.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    links.push(created.getUrl());
  });

  return links.join('\n');
}

function getAttachmentsFolder_() {
  const folders = DriveApp.getFoldersByName(ATTACHMENTS_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(ATTACHMENTS_FOLDER_NAME);
}

function normalizeDateForInput_(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;

  const m = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) {
    return `${m[3]}-${String(m[2]).padStart(2, '0')}-${String(m[1]).padStart(2, '0')}`;
  }

  const date = new Date(text);
  if (!isNaN(date.getTime())) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return '';
}