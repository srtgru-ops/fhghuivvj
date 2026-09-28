// السكربت مرتبط بجدول الردود المحدد، ويكتب في تبويب الردود نفسه.
const SPREADSHEET_ID = '17eZ89WlyTLX4U-xC2m60N4xF1Sg76KHCWz1VuA0r7ec';
const RESPONSE_SHEET_ID = 1053634930;
const QUESTIONS = [{"label":"الاسم كامل:","entry":"1857687044","required":true},{"label":"رقم الجوال","entry":"50109027","required":true},{"label":"البريد الإلكتروني","entry":"245833343","required":true},{"label":"الجهة / الإدارة","entry":"1947524408","required":false},{"label":"الفئة","entry":"2035022963","required":true},{"label":"هل سبق لك التسجيل في المنصة الوطنية للعمل التطوعي؟","entry":"1817991851","required":true},{"label":"هل سبق لك المشاركة في فرصة تطوعية من خلال المنصة؟","entry":"1320532521","required":true},{"label":"أتعرف على طريقة الدخول إلى المنصة","entry":"1482563673","required":true},{"label":"أعرف كيفية تحديث بياناتي الشخصية","entry":"1644725631","required":true},{"label":"أعرف كيفية البحث عن الفرص التطوعية","entry":"1230340556","required":true},{"label":"أعرف كيفية التسجيل في فرصة تطوعية","entry":"922687665","required":true},{"label":"أعرف كيفية متابعة الفرص التي سجلت بها","entry":"671403525","required":true},{"label":"أعرف كيفية توثيق الساعات التطوعية","entry":"1653180075","required":true},{"label":"أعرف كيفية الاطلاع على سجل مشاركاتي التطوعية","entry":"783944499","required":true},{"label":"أعرف كيفية الحصول على الشهادات التطوعية","entry":"1031300689","required":true},{"label":"لدي معرفة بالتحديثات الجديدة في المنصة","entry":"136243343","required":true},{"label":"ما مدى معرفتك بالتحديثات الجديدة في المنصة؟","entry":"1765637807","required":true},{"label":"ما أكثر جانب ترغب في معرفة تفاصيله خلال الورشة؟","entry":"663533952","required":true},{"label":"ما أبرز الصعوبات التي تواجهك عند استخدام المنصة؟","entry":"360709053","required":true},{"label":"ما مستوى استخدامك للمنصة؟","entry":"1209453647","required":true},{"label":"ما الذي تتوقع أن تستفيد منه بعد حضور الورشة؟","entry":"854611107","required":true},{"label":"ما الموضوع أو السؤال الذي ترغب في طرحه خلال الورشة؟","entry":"2011053240","required":false}];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ورشة تحديثات منصة العمل التطوعي')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function normalizeHeader_(value) {
  return String(value || '').replace(/[\u064B-\u065F\u0670\u0640]/g, '')
    .replace(/[أإآ]/g, 'ا').replace(/ى/g, 'ي').replace(/ة/g, 'ه')
    .replace(/[^\p{L}\p{N}]/gu, '').toLowerCase();
}

function safeCell_(value) {
  const text = String(value == null ? '' : value).trim();
  if (text.length > 5000) throw new Error('إحدى الإجابات طويلة جدًا.');
  // منع تحويل إجابات المستخدم إلى صيغ في Google Sheets.
  return /^[=+@-]/.test(text) ? "'" + text : text;
}

function saveRegistration(answers) {
  if (!Array.isArray(answers) || answers.length !== QUESTIONS.length) {
    throw new Error('بيانات التسجيل غير مكتملة. حدّثي الصفحة وحاولي مرة أخرى.');
  }
  const cleaned = answers.map(safeCell_);
  QUESTIONS.forEach((q, i) => {
    if (q.required && !cleaned[i]) throw new Error('يرجى إكمال الحقل المطلوب: ' + q.label);
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetById(RESPONSE_SHEET_ID);
    if (!sheet) throw new Error('لم يُعثر على تبويب الردود في الشيت المحدد.');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const keys = headers.map(normalizeHeader_);
    const row = headers.map(() => '');
    const timestampIndex = keys.findIndex(x => x === normalizeHeader_('طابع زمني'));
    const emailIndex = keys.findIndex(x => x === normalizeHeader_('عنوان البريد الإلكتروني'));
    if (timestampIndex < 0 || emailIndex < 0) throw new Error('عناوين جدول الردود لا تطابق النموذج.');
    row[timestampIndex] = new Date();
    row[emailIndex] = cleaned[2];
    QUESTIONS.forEach((q, i) => {
      const column = keys.indexOf(normalizeHeader_(q.label));
      if (column < 0) throw new Error('العمود المطلوب غير موجود في الشيت: ' + q.label);
      row[column] = cleaned[i];
    });
    sheet.appendRow(row);
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}
