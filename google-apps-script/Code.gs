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



/** Internal, authenticated Sheets menu. Public doPost never approves or sends. */
const KH_SALES = Object.freeze({
  sheet: 'Албан үнийн санал', timezone: 'Asia/Ulaanbaatar',
  headers: ['Хүсэлтийн ID','Харилцагч','Утас','Имэйл','Байршил','Бүтээгдэхүүн / задаргаа',
    'м³','Ширхэг','Үнэ / м³ (₮)','Хөнгөлөлт (₮)','Подоны тоо','Нэг подоны барьцаа (₮)',
    'НӨАТ-ын нөхцөл','Хүчинтэй хоног','Нөхцөл / тайлбар','Имэйл шалгасан (ТИЙМ)',
    'Төлөв','Албан дугаар','Баталсан','Баталсан огноо','PDF','Илгээсэн огноо','Тайлбар / алдаа']
});
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Үнийн санал')
    .addItem('1. Анхны тохиргоо', 'setupSalesWorkflow')
    .addItem('2. Сонгосон хүсэлтээс санал бэлтгэх', 'prepareSalesQuote')
    .addItem('3. Урьдчилан харах', 'previewSalesQuote')
    .addItem('4. Баталж PDF-ээр имэйлдэх', 'approveAndEmailQuote').addToUi();
}
function setupSalesWorkflow() {
  var email = khIdentity_();
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('KH_SALES_APPROVERS')) {
    // Bootstrap is restricted to the spreadsheet owner, not arbitrary sheet editors.
    var owner = DriveApp.getFileById(KHANTAISHIR.spreadsheetId).getOwner();
    if (!owner || owner.getEmail().toLowerCase() !== email) throw new Error('Анхны тохиргоог хүснэгтийн эзэмшигч ажиллуулна.');
    props.setProperty('KH_SALES_APPROVERS', email);
  }
  khAuthorize_();
  khSalesSheet_();
  if (!props.getProperty('KH_QUOTE_FOLDER')) {
    var folder = DriveApp.createFolder('Хантайшир — баталсан үнийн санал');
    props.setProperty('KH_QUOTE_FOLDER', folder.getId());
  }
  if (!props.getProperty('KH_QUOTE_RECORDS_FOLDER')) {
    props.setProperty('KH_QUOTE_RECORDS_FOLDER', DriveApp.createFolder('Хантайшир — саналын баталгаажуулалтын бүртгэл').getId());
  }
  SpreadsheetApp.getUi().alert('Тохиргоо бэлэн. Шинэ хүсэлт автоматаар илгээгдэхгүй. Борлуулалт хянаж батална.');
}
function khIdentity_() {
  var email = Session.getActiveUser().getEmail().toLowerCase();
  var effective = Session.getEffectiveUser().getEmail().toLowerCase();
  if (!email || email !== effective) throw new Error('Google Sheets-ийн өөрийн цэснээс нэвтэрсэн хэрэглэгч ажиллуулна.');
  return email;
}
function khAuthorize_() {
  var email = khIdentity_();
  var list = (PropertiesService.getScriptProperties().getProperty('KH_SALES_APPROVERS') || '')
    .split(',').map(function(s){return s.trim().toLowerCase();}).filter(Boolean);
  if (list.indexOf(email) < 0) throw new Error('Борлуулалтын батлах эрх байхгүй. Эзэмшигчтэй холбогдоно уу.');
  return email;
}
function khSalesSheet_() {
  var ss = SpreadsheetApp.openById(KHANTAISHIR.spreadsheetId);
  var sh = ss.getSheetByName(KH_SALES.sheet);
  if (!sh) sh = ss.insertSheet(KH_SALES.sheet);
  if (!sh.getLastRow()) {
    sh.getRange(1,1,1,KH_SALES.headers.length).setValues([KH_SALES.headers])
      .setBackground('#0C4D38').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1); sh.setColumnWidths(1,KH_SALES.headers.length,160);
    sh.setColumnWidth(6,320); sh.setColumnWidth(15,320);
    sh.getRange('G:L').setNumberFormat('#,##0.####');
    sh.getRange('Q:W').setBackground('#edf3ef');
    sh.getRange('K1:M1').setNote('Борлуулалт шалгаж бөглөнө. Барьцаагүй бол тоо болон үнийг 0 гэж оруулна. НӨАТ-ын нөхцөлийг нийт үнэд нийцүүлж бичнэ.');
  } else if (sh.getRange(1,1,1,KH_SALES.headers.length).getValues()[0].join('\0') !== KH_SALES.headers.join('\0')) {
    throw new Error('Албан саналын хүснэгтийн багана өөрчлөгдсөн.');
  }
  return sh;
}
function khSelected_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss || ss.getId() !== KHANTAISHIR.spreadsheetId) throw new Error('Зорилтот хүснэгтээс ажиллуулна.');
  var sh = ss.getActiveSheet(); var row = sh.getActiveRange().getRow();
  if (sh.getName() !== name || row < 2) throw new Error('«'+name+'» табын өгөгдөлтэй нэг мөрийг сонгоно уу.');
  return {sheet:sh, row:row};
}
function khLock_(work) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Өөр үйлдэл явагдаж байна. Дахин оролдоно уу.');
  try {return work();} finally {lock.releaseLock();}
}
function prepareSalesQuote() {
  khAuthorize_(); var selected = khSelected_(KHANTAISHIR.sheetName);
  var lead = selected.sheet.getRange(selected.row,1,1,17).getValues()[0];
  if (!lead[1] || lead[16] !== 'Тийм') throw new Error('Зөвшөөрөлтэй хүсэлт шаардлагатай.');
  var targetRow = khLock_(function(){
    var sh = khSalesSheet_(); var match = sh.getRange('A:A').createTextFinder(String(lead[1])).matchEntireCell(true).findNext();
    if (match) return match.getRow();
    sh.appendRow([lead[1],lead[3],lead[4],lead[12],lead[5],lead[7] || lead[6],
      lead[8],lead[9],KHANTAISHIR.priceM3,0,'','','',7,'Подоны барьцааг подон буцаалтын нөхцөлөөр тооцно.','',
      'ХЯНАЛТАД','','','','','',''].map(function(v){return typeof v==='string'?asCell_(v):v;}));
    return sh.getLastRow();
  });
  var sh = khSalesSheet_(); sh.activate(); sh.setActiveRange(sh.getRange(targetRow,1,1,16));
  SpreadsheetApp.getUi().alert('Үнэ, хөнгөлөлт, подоны тоо/барьцаа, НӨАТ-ын нөхцөл, имэйлийг шалгана. P баганад ТИЙМ гэж бичээд урьдчилан харна уу.');
}
function khReadQuote_(values) {
  function text(i,max){return cleanText_(values[i],max);}
  function number(i,min,max){if(values[i]==='' || values[i]===null) throw new Error(KH_SALES.headers[i]+' бөглөнө үү.');return numberBetween_(values[i],min,max);}
  var q = {requestId:text(0,100),name:text(1,120),phone:text(2,24),email:text(3,160),location:text(4,120),
    product:text(5,1000),volume:number(6,0.1,100000),pieces:number(7,1,10000000),
    unitPrice:number(8,1,100000000),discount:number(9,0,1000000000000),pallets:number(10,0,1000000),
    depositEach:number(11,0,100000000),tax:text(12,200),validDays:number(13,1,365),terms:text(14,1500)};
  if (!/^[a-zA-Z0-9_-]{12,100}$/.test(q.requestId) || !q.name || !q.phone || !q.location || !q.product) throw new Error('Хүсэлтийн үндсэн мэдээлэл дутуу.');
  if (!/^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/.test(q.email) || text(15,20).toUpperCase() !== 'ТИЙМ') throw new Error('Нэг хэрэглэгчийн имэйл хаягийг шалгаж, P баганад ТИЙМ гэж бичнэ үү.');
  if (!q.tax) throw new Error('НӨАТ-ын нөхцөл болон нийт үнэд багтсан эсэхийг бичнэ үү.');
  if (![q.pieces,q.pallets,q.validDays,q.discount,q.depositEach,q.unitPrice].every(Number.isInteger)) throw new Error('Ширхэг, подон, хоног болон мөнгөн дүн бүхэл тоо байна.');
  q.subtotal = Math.round(q.volume*q.unitPrice);
  if (q.discount > q.subtotal) throw new Error('Хөнгөлөлт үндсэн үнээс их байна.');
  q.goodsTotal = q.subtotal-q.discount; q.depositTotal=q.pallets*q.depositEach;
  q.total=q.goodsTotal+q.depositTotal;
  // VAT is never silently inferred: reviewer must include any payable tax in unit price.
  return q;
}
function khHash_(q) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(q),Utilities.Charset.UTF_8));
}
function khRecordFile_(id) {
  if (!/^[a-zA-Z0-9_-]{12,100}$/.test(id)) throw new Error('Invalid request ID');
  var folderId=PropertiesService.getScriptProperties().getProperty('KH_QUOTE_RECORDS_FOLDER');
  if(!folderId)throw new Error('Анхны тохиргоог ажиллуулна уу.');
  var folder=DriveApp.getFolderById(folderId);var files=folder.getFilesByName(id+'.json');
  return {folder:folder,file:files.hasNext()?files.next():null};
}
function khGetRecord_(id){var f=khRecordFile_(id).file;return f?JSON.parse(f.getBlob().getDataAsString('UTF-8')):null;}
function khPutRecord_(q,r){
  var target=khRecordFile_(q.requestId);var json=JSON.stringify(r);
  if(target.file)target.file.setContent(json);
  else target.folder.createFile(q.requestId+'.json',json,MimeType.PLAIN_TEXT);
}
function previewSalesQuote() {
  khAuthorize_(); var sel=khSelected_(KH_SALES.sheet);
  var q=khReadQuote_(sel.sheet.getRange(sel.row,1,1,23).getValues()[0]);
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(khQuoteHtml_(q,null)).setWidth(850).setHeight(650),'Урьдчилсан санал — баталгаажаагүй');
}
function approveAndEmailQuote() {
  var actor=khAuthorize_(); var sel=khSelected_(KH_SALES.sheet);
  var q=khReadQuote_(sel.sheet.getRange(sel.row,1,1,23).getValues()[0]);
  var hash=khHash_(q); var ui=SpreadsheetApp.getUi();
  if(ui.alert('Албан үнийн санал батлах',q.name+'\nИмэйл: '+q.email+'\nНийт: '+khMoney_(q.total)+' ₮\nҮүнээс подоны барьцаа: '+khMoney_(q.depositTotal)+' ₮\n\nБаталснаар PDF үүсэж энэ хаяг руу илгээгдэнэ. НӨАТ зэрэг төлөх бүх дүн нэгж үнэд туссан эсэхийг шалгасан уу?',ui.ButtonSet.YES_NO)!==ui.Button.YES)return;
  khLock_(function(){
    khAuthorize_();
    var fresh=khReadQuote_(sel.sheet.getRange(sel.row,1,1,23).getValues()[0]);
    if(khHash_(fresh)!==hash)throw new Error('Батлах үеэр мэдээлэл өөрчлөгдсөн. Дахин хянана уу.');
    var r=khGetRecord_(q.requestId);
    if(r && ['SENT','SENDING','UNKNOWN'].indexOf(r.state)>=0)throw new Error('Энэ санал илгээгдсэн эсвэл илгээлтийн төлөв шалгах шаардлагатай. Давхар илгээхийг зогсоов.');
    if(r && r.hash!==hash)throw new Error('Баталсан санал өөрчлөгдсөн. Шинэ хүсэлтийн дугаараар шинэ санал бэлтгэнэ үү.');
    if(MailApp.getRemainingDailyQuota()<1)throw new Error('Имэйл илгээх өдрийн лимит хүрсэн. Батлаагүй.');
    if(!r){
      var props=PropertiesService.getScriptProperties();var year=Utilities.formatDate(new Date(),KH_SALES.timezone,'yyyy');
      var key='KH_QUOTE_SEQUENCE_'+year;var seq=Number(props.getProperty(key)||0)+1;
      props.setProperty(key,String(seq));
      r={state:'APPROVED',number:'KH-Q-'+year+'-'+('000000'+seq).slice(-6),approvedBy:actor,approvedAt:new Date().toISOString(),hash:hash,snapshot:q};
      khPutRecord_(q,r);
    }
    khWriteStatus_(sel,r);
    try {
      if(!r.fileId){
        var folderId=PropertiesService.getScriptProperties().getProperty('KH_QUOTE_FOLDER');
        if(!folderId)throw new Error('Анхны тохиргоог ажиллуулна уу.');
        var pdf=HtmlService.createHtmlOutput(khQuoteHtml_(q,r)).getAs(MimeType.PDF).setName(r.number+'.pdf');
        var file=DriveApp.getFolderById(folderId).createFile(pdf);
        r.fileId=file.getId();khPutRecord_(q,r);khWriteStatus_(sel,r);
      }
      var attachment=DriveApp.getFileById(r.fileId).getBlob();
      r.state='SENDING';khPutRecord_(q,r);khWriteStatus_(sel,r);SpreadsheetApp.flush();
      // If this call's result is ambiguous, automatic retries are deliberately blocked.
      MailApp.sendEmail({to:q.email,subject:'Хантайшир Эко Блок — үнийн санал '+r.number,
        body:'Сайн байна уу, '+q.name+'?\n\nБорлуулалтын албанаас баталгаажуулсан '+r.number+' дугаартай үнийн саналыг хавсаргалаа.\nНийт: '+khMoney_(q.total)+' ₮\nПодоны барьцаа: '+khMoney_(q.depositTotal)+' ₮\n\nХавсралтын нөхцөл болон хүчинтэй хугацааг харна уу.\nХантайшир Эко Блок',
        name:'Хантайшир Эко Блок',attachments:[attachment]});
      r.error='';r.state='SENT';r.sentAt=new Date().toISOString();khPutRecord_(q,r);khWriteStatus_(sel,r);
    } catch(err){
      if(r.state==='SENDING')r.state='UNKNOWN';
      r.error=cleanText_(err.message,500);khPutRecord_(q,r);khWriteStatus_(sel,r);throw err;
    }
  });
  ui.alert('Баталсан PDF саналыг '+q.email+' хаяг руу илгээлээ.');
}
function khWriteStatus_(sel,r){
  sel.sheet.getRange(sel.row,17,1,7).setValues([[r.state,r.number,r.approvedBy,r.approvedAt,
    r.fileId?'https://drive.google.com/file/d/'+r.fileId+'/view':'',r.sentAt||'',r.error||'']]);
}
function khMoney_(n){return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g,',');}
function khEscape_(s){return String(s==null?'':s).replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];});}
function khQuoteHtml_(q,r){
  var e=khEscape_; var approved=r?new Date(r.approvedAt):new Date();
  var until=new Date(approved.getTime()+q.validDays*86400000);
  var date=function(d){return Utilities.formatDate(d,KH_SALES.timezone,'yyyy.MM.dd');};
  var row=function(a,b){return '<tr><td>'+e(a)+'</td><td class="amount">'+e(b)+' ₮</td></tr>';};
  return '<!doctype html><html lang="mn"><head><meta charset="utf-8"><style>@page{size:A4;margin:18mm}body{font:12px Arial,sans-serif;color:#163c30;line-height:1.55}h1{font-size:26px;margin:0}h2{font-size:21px}header{border-bottom:3px solid #198754;padding-bottom:18px}.muted{color:#66756d}.badge{font-weight:bold;color:#198754}table{width:100%;border-collapse:collapse;margin:20px 0}td,th{border-bottom:1px solid #dbe5df;padding:10px;text-align:left}.amount{text-align:right;white-space:nowrap}.total td{background:#e8f5eb;font-size:17px;font-weight:bold}.terms{white-space:pre-wrap;overflow-wrap:break-word}footer{border-top:1px solid #dbe5df;margin-top:26px;padding-top:12px;font-size:10px}</style></head><body><header><p class="badge">ХАНТАЙШИР ЭКО БЛОК</p><h1>ҮНИЙН САНАЛ</h1><p>'+e(r?r.number:'УРЬДЧИЛСАН — БАТАЛГААЖААГҮЙ')+'</p></header><p>Огноо: '+date(approved)+' · Хүчинтэй: '+date(until)+' хүртэл</p><h2>'+e(q.name)+'</h2><p>'+e(q.phone)+' · '+e(q.email)+'<br>'+e(q.location)+'</p><p><strong>Бүтээгдэхүүн:</strong> '+e(q.product)+'</p><table><tr><th>Үзүүлэлт</th><th class="amount">Дүн</th></tr><tr><td>Тоо хэмжээ</td><td class="amount">'+e(q.volume)+' м³ / '+e(q.pieces)+' ш</td></tr>'+row('Нэг м³ үнэ',khMoney_(q.unitPrice))+row('Бүтээгдэхүүний үнэ',khMoney_(q.subtotal))+row('Хөнгөлөлт',khMoney_(q.discount))+row('Бүтээгдэхүүний төлөх дүн',khMoney_(q.goodsTotal))+row('Подоны буцаан олгох барьцаа ('+q.pallets+' × '+khMoney_(q.depositEach)+' ₮)',khMoney_(q.depositTotal))+'<tr class="total"><td>Нийт төлөх дүн</td><td class="amount">'+khMoney_(q.total)+' ₮</td></tr></table><p><strong>НӨАТ:</strong> '+e(q.tax)+'</p><p class="terms"><strong>Нөхцөл:</strong> '+e(q.terms)+'</p><footer>Хүсэлтийн дугаар: '+e(q.requestId)+'<br>'+(r?'Борлуулалтын баталгаажуулалт: '+e(r.approvedBy)+' · '+date(approved):'Энэ бол урьдчилсан харагдац. Албан санал биш.')+'</footer></body></html>';
}

/** Internal smoke test: creates a clearly labelled PDF, never sends email. */
function testSalesPdfNoEmail() {
  var actor=khAuthorize_();
  var values=['KH-TEST-000000000001','ТУРШИЛТ — ХАРИЛЦАГЧ БИШ','00000000','test@example.com','Туршилтын байршил','100 × 240 × 600 мм',10,694,300000,100000,7,25000,'НӨАТ нийт үнэд багтсан — туршилтын дүн',7,'ТУРШИЛТ. Төлбөрийн болон худалдааны баримт биш. Имэйл илгээгээгүй.','ТИЙМ'];
  var q=khReadQuote_(values);
  if(q.total!==3075000)throw new Error('Calculation self-test failed');
  var record={number:'ТУРШИЛТ — АЛБАН САНАЛ БИШ',approvedAt:new Date().toISOString(),approvedBy:actor};
  var pdf=HtmlService.createHtmlOutput(khQuoteHtml_(q,record)).getAs(MimeType.PDF).setName('TEST-QUOTE-NOT-FOR-SALE.pdf');
  if(pdf.getBytes().length<1000)throw new Error('PDF conversion returned an empty file');
  var folderId=PropertiesService.getScriptProperties().getProperty('KH_QUOTE_FOLDER');
  var file=DriveApp.getFolderById(folderId).createFile(pdf);
  Logger.log('PDF_SELF_TEST_OK '+file.getUrl()+' bytes='+pdf.getBytes().length+' mailQuota='+MailApp.getRemainingDailyQuota()+'; NO_EMAIL_SENT');
}
