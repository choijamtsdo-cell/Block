/**
 * ХАНТАЙШИР ЭКО БЛОК — Google Sheets lead receiver
 * Тохиргоо: Sheet -> Extensions -> Apps Script -> Code.gs
 * Deploy as Web App: Execute as Me, access Anyone.
 * Зорилтот Google Sheets дээр "Веб үнийн хүсэлт" табыг тусад нь үүсгэнэ.
 */
const KHANTAISHIR = Object.freeze({
  spreadsheetId: '1q8awc7qZYLhoPkWOfvptJAzcY25fNLUaJgFNdrJAquw',
  sheetName: 'Веб үнийн хүсэлт',
  dealerSheetName: 'Гэрээт борлуулагчийн хүсэлт',
  siteOrigin: 'https://choijamtsdo-cell.github.io',
  priceM3: 300000,
  headers: [
    'Огноо', 'Хүсэлтийн ID', 'Төлөв', 'Харилцагч', 'Утас',
    'Төслийн байршил', 'Блокийн төрөл', 'Блокийн задаргаа',
    'Нийт м³', 'Нийт ширхэг', 'Үнэ / м³ (₮)', 'Нийт үнэ (₮)',
    'Имэйл', 'Нэмэлт тайлбар', 'Эх сурвалж', 'UTM',
    'Мэдээлэл илгээх зөвшөөрөл'
  ]
});

const DEALER_HEADERS = [
  'Огноо', 'Хүсэлтийн ID', 'Төлөв', 'Байгууллага / дэлгүүр',
  'Холбоо барих хүн', 'Утас', 'Аймаг, хот', 'Имэйл',
  'Борлуулалтын туршлага', 'Нэмэлт мэдээлэл', 'Эх сурвалж',
  'UTM', 'Мэдээлэл илгээх зөвшөөрөл'
];

function doGet() {
  return HtmlService.createHtmlOutput(
    '<!doctype html><html><head><meta charset="utf-8"></head>' +
    '<body style="font:16px Arial;padding:24px">' +
    '<h2>Хантайшир Эко Блок — үнийн хүсэлтийн бүртгэл</h2>' +
    '<p>Сервис ажиллаж байна. Хүсэлтийг зөвхөн вебсайтаас илгээнэ.</p></body></html>'
  );
}

function setup() {
  var sh = getTargetSheet_();
  var dealerSheet = getDealerSheet_();
  SpreadsheetApp.flush();
  Logger.log('Үнийн санал: ' + sh.getParent().getUrl() + '#gid=' + sh.getSheetId());
  Logger.log('Гэрээт борлуулагч: ' + dealerSheet.getParent().getUrl() + '#gid=' + dealerSheet.getSheetId());
}

function doPost(e) {
  var requestId = '';
  try {
    var raw = String((e && e.parameter && e.parameter.payload) || '');
    if (!raw || raw.length > 12000) throw new Error('Payload size invalid');
    var q = JSON.parse(raw);
    requestId = String(q.requestId || '').slice(0, 100);
    // Old quote requests without formType continue working unchanged.
    if (q.formType && q.formType !== 'quote' && q.formType !== 'dealer') throw new Error('Invalid form type');
    var result = q.formType === 'dealer' ? saveDealer_(q) : saveLead_(q);
    return reply_({type: 'KHANTAISHIR_LEAD_ACK', ok: true,
      requestId: result.requestId, alreadySaved: result.alreadySaved});
  } catch (err) {
    console.error('Lead receiver:', err);
    return reply_({type: 'KHANTAISHIR_LEAD_ACK', ok: false,
      requestId: requestId, message: 'Хүсэлтийг бүртгэж чадсангүй.'});
  }
}

function saveLead_(q) {
  if (!q || typeof q !== 'object') throw new Error('Not an object');
  if (q.website) throw new Error('Spam field');
  if (q.consent !== true) throw new Error('Missing consent');
  var id = String(q.requestId || '');
  if (!/^[a-zA-Z0-9_-]{12,100}$/.test(id)) throw new Error('Invalid id');

  var name = cleanText_(q.name, 120);
  var phone = cleanText_(q.phone, 24);
  var digits = phone.replace(/\D/g, '');
  if (digits.length === 11 && digits.slice(0, 3) === '976') digits = digits.slice(3);
  if (!name || !/^\d{8}$/.test(digits)) throw new Error('Invalid contact');
  var place = cleanText_(q.location, 120);
  if (!place) throw new Error('Missing location');
  var volume = numberBetween_(q.volume, 0.1, 100000);
  var pieces = numberBetween_(q.pieces, 1, 10000000);
  var product = cleanText_(q.product, 100);
  var breakdown = cleanText_(q.breakdown, 1000);
  if (!product) throw new Error('Missing product');
  var email = cleanText_(q.email, 160);
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Invalid email');
  var notes = cleanText_(q.notes, 700);
  var source = cleanText_(q.source || 'Website', 150);
  var utm = cleanText_(q.utm, 500);
  var total = Math.round(volume * KHANTAISHIR.priceM3);

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Busy, try again');
  try {
    var sh = getTargetSheet_();
    var lastRow = sh.getLastRow();
    if (lastRow > 1) {
      var match = sh.getRange(2, 2, lastRow - 1, 1)
        .createTextFinder(id).matchEntireCell(true).findNext();
      if (match) return {requestId: id, alreadySaved: true};
    }
    sh.appendRow([
      new Date(), id, 'Шинэ', asCell_(name), asCell_(phone),
      asCell_(place), asCell_(product), asCell_(breakdown),
      volume, Math.round(pieces), KHANTAISHIR.priceM3, total,
      asCell_(email), asCell_(notes), asCell_(source), asCell_(utm), 'Тийм'
    ]);
    return {requestId: id, alreadySaved: false};
  } finally {
    lock.releaseLock();
  }
}

function getTargetSheet_() {
  var ss = SpreadsheetApp.openById(KHANTAISHIR.spreadsheetId);
  var sh = ss.getSheetByName(KHANTAISHIR.sheetName);
  var created = false;
  if (!sh) {
    sh = ss.insertSheet(KHANTAISHIR.sheetName);
    created = true;
  }
  var headers = KHANTAISHIR.headers;
  if (created || sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length)
      .setBackground('#0C4D38').setFontColor('#FFFFFF').setFontWeight('bold');
    sh.getRange('I:L').setNumberFormat('#,##0.####');
    sh.setColumnWidth(1, 165);
    sh.setColumnWidth(4, 210);
    sh.setColumnWidth(5, 135);
  } else {
    var first = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    if (first.join('\u0000') !== headers.join('\u0000')) {
      throw new Error('An existing tab has a different column structure; no edits were made.');
    }
  }
  return sh;
}

/**
 * Stores dealership applications separately from purchase quote leads.
 * Existing quote sheet data and columns are untouched.
 */
function saveDealer_(q) {
  if (!q || typeof q !== 'object') throw new Error('Not an object');
  if (q.website) throw new Error('Spam field');
  if (q.consent !== true) throw new Error('Missing dealer consent');
  var id = String(q.requestId || '');
  if (!/^[a-zA-Z0-9_-]{12,100}$/.test(id)) throw new Error('Invalid dealer request id');

  var organization = cleanText_(q.organization, 120);
  var name = cleanText_(q.name, 120);
  var phone = cleanText_(q.phone, 24);
  var digits = phone.replace(/\D/g, '');
  if (digits.length === 11 && digits.slice(0, 3) === '976') digits = digits.slice(3);
  var location = cleanText_(q.location, 120);
  var email = cleanText_(q.email, 160);
  var experience = cleanText_(q.experience, 150);
  var notes = cleanText_(q.notes, 700);
  var source = cleanText_(q.source || 'Website', 150);
  var utm = cleanText_(q.utm, 500);
  if (!organization || !name || !/^\d{8}$/.test(digits) || !location) {
    throw new Error('Dealer required fields are invalid');
  }
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Invalid email');

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Busy, try again');
  try {
    var sh = getDealerSheet_();
    var lastRow = sh.getLastRow();
    if (lastRow > 1) {
      var match = sh.getRange(2, 2, lastRow - 1, 1)
        .createTextFinder(id).matchEntireCell(true).findNext();
      if (match) return {requestId: id, alreadySaved: true};
    }
    sh.appendRow([
      new Date(), id, 'Шинэ', asCell_(organization), asCell_(name),
      asCell_(phone), asCell_(location), asCell_(email),
      asCell_(experience), asCell_(notes), asCell_(source), asCell_(utm), 'Тийм'
    ]);
    return {requestId: id, alreadySaved: false};
  } finally {
    lock.releaseLock();
  }
}

/** Creates dealer tab only when absent; never overwrites populated tabs. */
function getDealerSheet_() {
  var ss = SpreadsheetApp.openById(KHANTAISHIR.spreadsheetId);
  var sh = ss.getSheetByName(KHANTAISHIR.dealerSheetName);
  var created = false;
  if (!sh) {
    sh = ss.insertSheet(KHANTAISHIR.dealerSheetName);
    created = true;
  }
  if (created || sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, DEALER_HEADERS.length).setValues([DEALER_HEADERS]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, DEALER_HEADERS.length)
      .setBackground('#0C4D38').setFontColor('#FFFFFF').setFontWeight('bold');
    sh.setColumnWidth(1, 165);
    sh.setColumnWidth(4, 230);
    sh.setColumnWidth(5, 170);
    sh.setColumnWidth(6, 135);
  } else {
    var existing = sh.getRange(1, 1, 1, DEALER_HEADERS.length).getValues()[0];
    if (existing.join('\u0000') !== DEALER_HEADERS.join('\u0000')) {
      throw new Error('Dealer tab exists with different headers. No edits made.');
    }
  }
  return sh;
}

function cleanText_(v, maxLen) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001F\u007F]/g, ' ')
    .trim().slice(0, maxLen);
}

function numberBetween_(v, lo, hi) {
  var n = Number(v);
  if (!isFinite(n) || n < lo || n > hi) throw new Error('Invalid number');
  return n;
}

function asCell_(s) {
  // Google Sheets formula injection protection for untrusted visitor fields.
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function reply_(response) {
  // Respond inside the cross-origin iframe instead of depending on CORS.
  var payload = JSON.stringify(response).replace(/</g, '\\u003c');
  var target = JSON.stringify(KHANTAISHIR.siteOrigin);
  var html = '<!doctype html><html><head><meta charset="utf-8"></head><body>' +
    '<script>try{window.top.postMessage(' + payload + ',' + target +
    ');}catch(e){}</script></body></html>';
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
