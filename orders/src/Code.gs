/**
 * لوحة الطلبات - أرض الأجداد
 * Google Sheet + Apps Script web app (زي استوديو المحتوى)
 * الأدوار: مدير (كل حاجة + أزرار الأوامر) / مجهز (ملفات اليوم + تعليم التجهيز)
 * Claude بيكتب البيانات في الشيت ويقرا تبويب "الأوامر" وينفذ.
 */

var SHEETS = {
  ORDERS: 'الطلبات',
  DAYS: 'ملفات الأيام',
  CMDS: 'الأوامر',
  STATE: 'الملخص',
  GUIDE: 'الدليل',
  USERS: 'المستخدمين',
  LOG: 'السجل'
};

var HEADERS = {
  'الطلبات': ['رقم الطلب', 'التاريخ', 'العميل', 'المدينة', 'المصدر', 'الدفع', 'شركة الشحن', 'النوع', 'المرحلة', 'رقم البوليصة', 'تكلفة البوليصة', 'يوم الملف', 'ملاحظات', 'جهّزه', 'آخر تحديث', 'رقم أمازون', 'المنتج', 'الكمية', 'آخر موعد شحن', 'المبلغ'],
  'ملفات الأيام': ['اليوم', 'القسم', 'عدد الطلبات', 'التكلفة', 'رابط البوليصات', 'رابط الفواتير', 'رابط التجهيز', 'فواتير الفرع', 'حالة التجهيز', 'جهّزه', 'وقت التجهيز', 'المسار في الأرشيف'],
  'الأوامر': ['المعرف', 'الوقت', 'المستخدم', 'الأمر', 'التفاصيل', 'الحالة', 'الرد', 'وقت التنفيذ'],
  'الملخص': ['المفتاح', 'القيمة', 'آخر تحديث'],
  'الدليل': ['القسم', 'القاعدة'],
  'المستخدمين': ['الإيميل', 'الاسم', 'الدور', 'كلمة السر (مشفرة)', 'نشط'],
  'السجل': ['الوقت', 'المستخدم', 'الحدث']
};

var STAGES = [
  'بانتظار الإصدار',
  'بانتظار تأكيد التحويل',
  'دفع عند الاستلام',
  'استلام من الفرع',
  'اتصدرت',
  'اتجهزت',
  'تم الشحن',
  'ملغي/مرتجع'
];

var CMD_TYPES = {
  ISSUE: 'أصدر',
  EXCEPTIONAL: 'إصدار استثنائي',
  TRANSFER_OK: 'اتأكد التحويل',
  NOTE: 'ملاحظة لـ Claude'
};

/* ---------------- setup ---------------- */

function setup() {
  var ss = SS_();
  Object.keys(HEADERS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    var h = HEADERS[name];
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold')
      .setBackground('#3b2a1a').setFontColor('#f3e6c8');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  });
  var first = ss.getSheets()[0];
  if (first.getName() !== SHEETS.ORDERS && first.getLastRow() === 0 && ss.getSheets().length > Object.keys(HEADERS).length) {
    ss.deleteSheet(first);
  }
  seedGuide_();
  seedState_();
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SECRET')) props.setProperty('SECRET', Utilities.getUuid());
  ensureFolder_();
  return 'تم التجهيز';
}

function seedState_() {
  var sh = sheet_(SHEETS.STATE);
  if (sh.getLastRow() > 1) return;
  var now = new Date();
  var rows = [
    ['رصيد محفظة الشحن', '', now],
    ['رصيد المدفوعات', '', now],
    ['سطر الإصدار المنتظر', '', now],
    ['حالة الإصدار', 'لا يوجد', now],
    ['آخر تشغيل لـ Claude', '', now],
    ['رسالة من Claude', '', now]
  ];
  sh.getRange(2, 1, rows.length, 3).setValues(rows);
}

function seedGuide_() {
  var sh = sheet_(SHEETS.GUIDE);
  if (sh.getLastRow() > 1) return;
  var rules = [
    ['الإصدار', 'البوليصات بتتصدر للطلبات اللي حالتها «بإنتظار المراجعة» والمدفوعة، ما عدا التحويل البنكي والدفع عند الاستلام، والأقدم الأول.'],
    ['الإصدار', 'قبل الإصدار Claude يكتب سطر فيه عدد البوليصات وإجمالي التكلفة، ومفيش إصدار غير بعد زرار «أصدر».'],
    ['الإصدار', 'أيام الإصدار من السبت للخميس الصبح. طلبات الخميس بعد الإصدار والجمعة بتتصدر السبت. زرار «إصدار استثنائي» يصدّر في أي وقت.'],
    ['الإصدار', 'بعد الإصدار الطلب يبقى «قيد التنفيذ»، وبعد يومين «تم الشحن».'],
    ['شركات الشحن', '«لوجستي» تتحول لسحاب العادي. سمسا ودي إتش إل وأرامكس يفضلوا زي ما هم. مفيش مبرد.'],
    ['التحويل البنكي', 'مايتصدرش. بتروح رسالة لأبو فهد (الاسم والمبلغ والساعة)، وبعد «اتأكدت» يتصدر.'],
    ['الدفع عند الاستلام', 'بيستنى لحد ما يتأكد مع العميل.'],
    ['الإهداء', 'يتصدر عادي، ويتبعت تقرير برقم الطلب وجملة الإهداء وسكرين.'],
    ['الاستلام من الفرع', 'مفيش بوليصة. الفاتورة تتبعت لوحدها باسم رقم الطلب، مكتوب عليها «مدفوعة» أو «غير مدفوعة».'],
    ['أمازون', 'طلب في سلة بسعر أمازون والشحن مجاني، تحويل بنكي (مصرف الراجحي) + صورة الإيصال، البوليصة على سمسا.'],
    ['أمازون', 'بعد الإصدار: Confirm Shipment على أمازون برقم التتبع، والطلب في سلة على حالة «طلبات أمازون» + وسم «أمازون».'],
    ['أمازون', 'المجهز بياخد 3 ملفات لوحدها: بوليصات (طابعة الاستيكر) / فواتير أمازون الأصلية (طابعة عادية) / تجهيز (صورة المنتج والكمية — للقراية بس).'],
    ['الملفات', 'الأرشيف على الجهاز في E:\\بوليصات (فولدر لكل يوم + فولدر امازون)، ونسخة على الدرايف للوحة.'],
    ['التنبيهات', 'الساعة 10 بالليل تنبيه لو رصيد محفظة الشحن أو المدفوعات مش كفاية لإصدار الصبح.']
  ];
  sh.getRange(2, 1, rules.length, 2).setValues(rules);
}

function ensureFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder('لوحة الطلبات - ملفات الأيام');
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}

/** يدي حساب تاني (حساب الموصل اللي Claude بيكتب منه) صلاحية تعديل الشيت وفولدر الملفات */
function shareWithEditor(email) {
  SS_().addEditor(email);
  ensureFolder_().addEditor(email);
  return 'تمت المشاركة مع ' + email;
}

/** إضافة مستخدم: addUser('x@gmail.com','الاسم','مدير' أو 'مجهز','كلمة سر مؤقتة') */
function addUser(email, name, role, password) {
  var sh = sheet_(SHEETS.USERS);
  var data = sh.getDataRange().getValues();
  email = String(email).trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === email) {
      sh.getRange(i + 1, 2, 1, 4).setValues([[name, role, hash_(password), true]]);
      return 'اتحدث ' + email;
    }
  }
  sh.appendRow([email, name, role, hash_(password), true]);
  return 'اتضاف ' + email;
}

/* ---------------- web app ---------------- */

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('لوحة الطلبات - أرض الأجداد')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function login(email, password) {
  email = String(email || '').trim().toLowerCase();
  var rows = sheet_(SHEETS.USERS).getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (String(r[0]).toLowerCase() === email && r[4] !== false && String(r[4]).toUpperCase() !== 'FALSE') {
      if (r[3] === hash_(password)) {
        var token = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
        var user = { email: email, name: r[1], role: r[2] };
        saveSession_(token, user);
        log_(user.name, 'دخل اللوحة');
        return { ok: true, token: token, user: user };
      }
      return { ok: false, msg: 'كلمة السر غلط' };
    }
  }
  return { ok: false, msg: 'الإيميل ده مش متسجل' };
}

function changePassword(token, oldPass, newPass) {
  var u = auth_(token);
  if (!newPass || String(newPass).length < 6) return { ok: false, msg: 'كلمة السر لازم 6 حروف على الأقل' };
  var sh = sheet_(SHEETS.USERS);
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === u.email) {
      if (rows[i][3] !== hash_(oldPass)) return { ok: false, msg: 'كلمة السر القديمة غلط' };
      sh.getRange(i + 1, 4).setValue(hash_(newPass));
      log_(u.name, 'غيّر كلمة السر');
      return { ok: true };
    }
  }
  return { ok: false, msg: 'مش لاقي المستخدم' };
}

function getData(token) {
  var u = auth_(token);
  var orders = rows_(SHEETS.ORDERS);
  var days = rows_(SHEETS.DAYS).reverse();
  var state = {};
  rows_(SHEETS.STATE).forEach(function (r) { state[r['المفتاح']] = { v: r['القيمة'], t: r['آخر تحديث'] }; });
  var out = { user: u, state: state, days: days.slice(0, 60), stages: STAGES, now: new Date().toISOString() };
  if (u.role === 'مدير') {
    out.orders = orders;
    out.cmds = rows_(SHEETS.CMDS).reverse().slice(0, 50);
    out.guide = rows_(SHEETS.GUIDE);
    out.log = rows_(SHEETS.LOG).reverse().slice(0, 40);
  } else {
    // المجهز: طلبات آخر الأيام اللي ليها ملفات بس
    var recentDays = {};
    days.slice(0, 7).forEach(function (d) { recentDays[fmtDay_(d['اليوم'])] = true; });
    out.orders = orders.filter(function (o) {
      return recentDays[fmtDay_(o['يوم الملف'])] && (o['المرحلة'] === 'اتصدرت' || o['المرحلة'] === 'اتجهزت');
    });
  }
  return out;
}

function sendCommand(token, type, details) {
  var u = auth_(token);
  if (u.role !== 'مدير') throw new Error('الأمر ده للمدير بس');
  var valid = Object.keys(CMD_TYPES).map(function (k) { return CMD_TYPES[k]; });
  if (valid.indexOf(type) < 0) throw new Error('أمر غير معروف');
  var id = 'C' + Utilities.formatDate(new Date(), 'Africa/Cairo', 'yyMMddHHmmss');
  sheet_(SHEETS.CMDS).appendRow([id, new Date(), u.name, type, details || '', 'جديد', '', '']);
  if (type === CMD_TYPES.ISSUE || type === CMD_TYPES.EXCEPTIONAL) setState_('حالة الإصدار', 'اتبعت أمر «' + type + '» — مستني Claude');
  log_(u.name, 'بعت أمر: ' + type + (details ? ' — ' + details : ''));
  return { ok: true, id: id };
}

function cancelCommand(token, id) {
  var u = auth_(token);
  if (u.role !== 'مدير') throw new Error('للمدير بس');
  var sh = sheet_(SHEETS.CMDS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id && data[i][5] === 'جديد') {
      sh.getRange(i + 1, 6, 1, 2).setValues([['ملغي', 'لغاه ' + u.name]]);
      log_(u.name, 'لغى الأمر ' + id);
      return { ok: true };
    }
  }
  return { ok: false, msg: 'الأمر اتنفذ أو مش موجود' };
}

/** المجهز (أو المدير) يعلّم طلب اتجهز / يرجعه */
function markPrepared(token, orderId, done) {
  var u = auth_(token);
  var sh = sheet_(SHEETS.ORDERS);
  var data = sh.getDataRange().getValues();
  var h = data[0];
  var cId = h.indexOf('رقم الطلب'), cSt = h.indexOf('المرحلة'), cBy = h.indexOf('جهّزه'), cUp = h.indexOf('آخر تحديث');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][cId]) === String(orderId)) {
      var cur = data[i][cSt];
      if (done && cur !== 'اتصدرت') return { ok: false, msg: 'الطلب مش في مرحلة «اتصدرت»' };
      if (!done && cur !== 'اتجهزت') return { ok: false, msg: 'الطلب مش متعلّم إنه اتجهز' };
      sh.getRange(i + 1, cSt + 1).setValue(done ? 'اتجهزت' : 'اتصدرت');
      sh.getRange(i + 1, cBy + 1).setValue(done ? u.name : '');
      sh.getRange(i + 1, cUp + 1).setValue(new Date());
      log_(u.name, (done ? 'جهّز ' : 'رجّع تجهيز ') + orderId);
      return { ok: true };
    }
  }
  return { ok: false, msg: 'الطلب مش موجود' };
}

/** المجهز يعلّم ملفات يوم كلها اتجهزت */
function markDayPrepared(token, day, section) {
  var u = auth_(token);
  var sh = sheet_(SHEETS.DAYS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (fmtDay_(data[i][0]) === day && data[i][1] === section) {
      sh.getRange(i + 1, 9, 1, 3).setValues([['اتجهز', u.name, new Date()]]);
      log_(u.name, 'خلّص تجهيز ' + section + ' ' + day);
      return { ok: true };
    }
  }
  return { ok: false, msg: 'اليوم مش موجود' };
}


/** تنزيل ملف من فولدر اللوحة لمستخدم داخل (من غير ما الملف يبقى عام) */
function getFile(token, url) {
  auth_(token);
  var m = String(url).match(/[-\w]{25,}/);
  if (!m) throw new Error('رابط غير صالح');
  var file = DriveApp.getFileById(m[0]);
  var folderId = ensureFolder_().getId();
  var ok = false, ps = file.getParents();
  while (ps.hasNext()) { if (ps.next().getId() === folderId) ok = true; }
  if (!ok) throw new Error('الملف ده مش من ملفات اللوحة');
  var blob = file.getBlob();
  return { name: file.getName(), mime: blob.getContentType(), b64: Utilities.base64Encode(blob.getBytes()) };
}


/* ---------------- API (للصفحة على GitHub Pages) ---------------- */

var API_FUNCS = { login: login, changePassword: changePassword, getData: getData, sendCommand: sendCommand,
  cancelCommand: cancelCommand, markPrepared: markPrepared, markDayPrepared: markDayPrepared, getFile: getFile,
  logout: logout, requestReset: requestReset, resetPassword: resetPassword };

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    if (req.action === 'app') {
      return ContentService.createTextOutput(HtmlService.createHtmlOutputFromFile('Index').getContent())
        .setMimeType(ContentService.MimeType.TEXT);
    }
    var fn = API_FUNCS[req.action];
    if (!fn) throw new Error('أمر غير معروف');
    out = { ok: true, result: fn.apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- helpers ---------------- */

var SESSION_DAYS = 90;

function saveSession_(token, user) {
  var rec = { u: user, exp: Date.now() + SESSION_DAYS * 864e5 };
  PropertiesService.getScriptProperties().setProperty('s_' + token, JSON.stringify(rec));
  CacheService.getScriptCache().put('t_' + token, JSON.stringify(user), 21600);
  cleanSessions_();
}

function cleanSessions_() {
  var p = PropertiesService.getScriptProperties(), all = p.getProperties(), now = Date.now();
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('s_') === 0 || k.indexOf('r_') === 0) {
      try { if (JSON.parse(all[k]).exp < now) p.deleteProperty(k); } catch (e) { p.deleteProperty(k); }
    }
  });
}

function auth_(token) {
  if (!token) throw new Error('SESSION_EXPIRED');
  var cache = CacheService.getScriptCache();
  var raw = cache.get('t_' + token);
  if (raw) return JSON.parse(raw);
  var rec = PropertiesService.getScriptProperties().getProperty('s_' + token);
  if (!rec) throw new Error('SESSION_EXPIRED');
  rec = JSON.parse(rec);
  if (rec.exp < Date.now()) { PropertiesService.getScriptProperties().deleteProperty('s_' + token); throw new Error('SESSION_EXPIRED'); }
  cache.put('t_' + token, JSON.stringify(rec.u), 21600);
  return rec.u;
}

function logout(token) {
  if (token) {
    PropertiesService.getScriptProperties().deleteProperty('s_' + token);
    CacheService.getScriptCache().remove('t_' + token);
  }
  return { ok: true };
}

/* ---------- استرجاع كلمة السر بالإيميل ---------- */

function findUserRow_(email) {
  var sh = sheet_(SHEETS.USERS), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) if (String(rows[i][0]).trim().toLowerCase() === email) return { sh: sh, i: i + 1, r: rows[i] };
  return null;
}

function requestReset(email) {
  email = String(email || '').trim().toLowerCase();
  var f = findUserRow_(email);
  if (f && f.r[4] !== false && String(f.r[4]).toUpperCase() !== 'FALSE') {
    var p = PropertiesService.getScriptProperties();
    var old = p.getProperty('r_' + email);
    if (old) { try { if (Date.now() - JSON.parse(old).at < 60000) return { ok: true }; } catch (e) {} }
    var code = String(Math.floor(100000 + Math.random() * 900000));
    p.setProperty('r_' + email, JSON.stringify({ h: hash_(code), exp: Date.now() + 15 * 60000, at: Date.now(), tries: 0 }));
    MailApp.sendEmail({
      to: email,
      subject: 'كود استرجاع كلمة السر — لوحة الطلبات (أرض الأجداد)',
      htmlBody: '<div dir="rtl" style="font-family:Tahoma,Arial;font-size:15px">' +
        '<p>أهلاً ' + f.r[1] + '،</p><p>كود تغيير كلمة السر بتاع لوحة الطلبات:</p>' +
        '<p style="font-size:28px;font-weight:bold;letter-spacing:4px">' + code + '</p>' +
        '<p>الكود شغال 15 دقيقة. لو انت ماطلبتش تغيير كلمة السر، تجاهل الرسالة دي.</p></div>',
      name: 'لوحة الطلبات - أرض الأجداد'
    });
    log_(f.r[1], 'طلب كود استرجاع كلمة السر');
  }
  return { ok: true };
}

function resetPassword(email, code, newPass) {
  email = String(email || '').trim().toLowerCase();
  newPass = String(newPass || '').trim();
  if (newPass.length < 6) return { ok: false, msg: 'كلمة السر لازم 6 حروف على الأقل' };
  var p = PropertiesService.getScriptProperties(), raw = p.getProperty('r_' + email);
  if (!raw) return { ok: false, msg: 'الكود انتهى أو مش موجود، اطلب كود جديد' };
  var rec = JSON.parse(raw);
  if (rec.exp < Date.now() || rec.tries >= 5) { p.deleteProperty('r_' + email); return { ok: false, msg: 'الكود انتهى، اطلب كود جديد' }; }
  if (rec.h !== hash_(String(code || '').trim())) { rec.tries++; p.setProperty('r_' + email, JSON.stringify(rec)); return { ok: false, msg: 'الكود غلط' }; }
  var f = findUserRow_(email);
  if (!f) return { ok: false, msg: 'الإيميل مش متسجل' };
  f.sh.getRange(f.i, 4).setValue(hash_(newPass));
  p.deleteProperty('r_' + email);
  // اقفل كل الجلسات القديمة للمستخدم ده
  var all = p.getProperties();
  Object.keys(all).forEach(function (k) { if (k.indexOf('s_') === 0) { try { if (JSON.parse(all[k]).u.email === email) p.deleteProperty(k); } catch (e) {} } });
  log_(f.r[1], 'غيّر كلمة السر بكود الإيميل');
  var token = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
  var user = { email: email, name: f.r[1], role: f.r[2] };
  saveSession_(token, user);
  return { ok: true, token: token, user: user };
}

function hash_(s) {
  var salt = PropertiesService.getScriptProperties().getProperty('SECRET') || '';
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + String(s), Utilities.Charset.UTF_8);
  return b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join('');
}

function sheet_(name) { return SS_().getSheetByName(name); }

function rows_(name) {
  var sh = sheet_(name);
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getDataRange().getValues();
  var h = data.shift();
  return data.filter(function (r) { return r.join('') !== ''; }).map(function (r) {
    var o = {};
    h.forEach(function (k, j) { o[k] = r[j] instanceof Date ? r[j].toISOString() : r[j]; });
    return o;
  });
}

function setState_(key, val) {
  var sh = sheet_(SHEETS.STATE);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) { sh.getRange(i + 1, 2, 1, 2).setValues([[val, new Date()]]); return; }
  }
  sh.appendRow([key, val, new Date()]);
}

function log_(who, what) { sheet_(SHEETS.LOG).appendRow([new Date(), who, what]); }

function fmtDay_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Africa/Cairo', 'yyyy-MM-dd');
  var s = String(v);
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) return Utilities.formatDate(new Date(s), 'Africa/Cairo', 'yyyy-MM-dd');
  return s.slice(0, 10);
}

var SHEET_ID = '1RufzAkyuhbxpmESsLiLhtnznrTC37-qsb36d5Wwkdgo';
function SS_() { return SpreadsheetApp.openById(SHEET_ID); }
