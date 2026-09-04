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

async function queryRecord(){
  // 1. ظهور نافذة لإدخال الرقم القومي
  const nationalId = prompt("الرجاء إدخال الرقم القومي للاستعلام عن الطلب:", "");
  
  // إذا ضغط المستخدم "إلغاء"
  if (nationalId === null) {
    return;
  }

  const trimmedId = nationalId.trim();
  if (!trimmedId) {
    showMessage("لم تدخل أي رقم. يرجى المحاولة مرة أخرى.", true);
    return;
  }

  // التحقق من صحة الرقم (14 رقم)
  if (!/^\d{14}$/.test(trimmedId)) {
    alert("⚠️ الرقم القومي يجب أن يتكون من 14 رقمًا.");
    showMessage("الرقم القومي غير صحيح.", true);
    return;
  }

  // تحديث حقل الرقم القومي في النموذج للرجوع إليه
  document.getElementById("nationalId").value = trimmedId;

  showMessage("جاري الاستعلام عن الرقم: " + trimmedId + "...");
  try {
    const result = await callServer("query", { nationalId: trimmedId });
    
    if(!result || !result.found){
      // عرض رسالة في نافذة منبثقة عند عدم العثور
      alert("❌ لم يتم العثور على طلب بهذا الرقم القومي.");
      showMessage(result && result.message ? result.message : "لا توجد بيانات.", true);
      document.getElementById("statusDisplay").classList.remove("show");
      return;
    }

    const record = result.data;

    // تعبئة الحقول في النموذج بالبيانات المسترجعة
    fieldIds.forEach(id=>{
      if(record[id] !== undefined && record[id] !== null) {
        document.getElementById(id).value = record[id];
      }
    });

    // ============================================================
    // 2. عرض نتيجة الطلب في نافذة منبثقة (alert) بشكل منسق
    // ============================================================
    let statusMsg = `✅ تم العثور على الطلب.\n\n`;
    statusMsg += `👤 الاسم: ${record.name || 'غير محدد'}\n`;
    statusMsg += `🆔 الرقم القومي: ${record.nationalId}\n`;
    statusMsg += `📋 الحالة: ${record.status || 'لم تحدد بعد'}\n`;
    if (record.status && record.status.trim().toLowerCase() === "مرفوض" && record.rejectionReason) {
      statusMsg += `❌ سبب الرفض: ${record.rejectionReason}\n`;
    }
    alert(statusMsg);
    // ============================================================

    // عرض الحالة في واجهة الصفحة (أسفل الحقول) أيضاً
    const statusDisplay = document.getElementById("statusDisplay");
    const statusValue = document.getElementById("statusValue");
    const reasonDisplay = document.getElementById("rejectionReasonDisplay");
    if (record.status) {
      statusDisplay.classList.add("show");
      statusValue.textContent = record.status;
      if (record.status.trim().toLowerCase() === "مرفوض" && record.rejectionReason) {
        reasonDisplay.style.display = "block";
        reasonDisplay.textContent = "سبب الرفض: " + record.rejectionReason;
      } else {
        reasonDisplay.style.display = "none";
      }
    } else {
      statusDisplay.classList.remove("show");
    }

    // تحديث قائمة المرفقات المعروضة (للطباعة)
    if (record.attachments) {
      const printList = document.getElementById("filesListPrint");
      printList.innerHTML = "";
      const links = record.attachments.split('\n').filter(l=>l.trim());
      links.forEach(link=>{
        const li = document.createElement("li");
        const a = document.createElement("a");
        a.href = link;
        a.textContent = "رابط المرفق";
        a.target = "_blank";
        li.appendChild(a);
        printList.appendChild(li);
      });
    }

    showMessage("تم استدعاء البيانات بنجاح.");

  } catch(err) {
    alert("⚠️ حدث خطأ أثناء الاستعلام: " + (err.message || "خطأ غير معروف"));
    showMessage(err.message || "حدث خطأ أثناء الاستعلام.", true);
  }
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
