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
    .addItem('Борлуулалтын самбар нээх', 'openSalesDashboard')
    .addSeparator()
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
  khApproveAndSend_(sel,q,hash,actor);
  ui.alert('Баталсан PDF саналыг '+q.email+' хаяг руу илгээлээ.');
}
function khApproveAndSend_(sel,q,hash,actor) {
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


/** Sales dashboard is served only from the authenticated bound Sheet menu. */
function openSalesDashboard(){
  khAuthorize_();SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(khDashboardHtml_()).setWidth(1180).setHeight(800),'Борлуулалтын самбар');
}
function khDashboardId_(id){id=String(id||'');if(!/^[a-zA-Z0-9_-]{12,100}$/.test(id))throw new Error('Хүсэлтийн дугаар буруу.');return id;}
function khDashboardRow_(id){
  var sh=khSalesSheet_();var found=sh.getRange('A:A').createTextFinder(khDashboardId_(id)).matchEntireCell(true).findNext();
  if(!found||found.getRow()<2)throw new Error('Санал олдсонгүй.');return {sheet:sh,row:found.getRow()};
}
function salesDashboardData(){
  khAuthorize_();var sh=getTargetSheet_();var count=Math.min(200,Math.max(0,sh.getLastRow()-1));
  var drafts=khSalesSheet_();var draftCount=Math.max(0,drafts.getLastRow()-1);var states={};
  if(draftCount)drafts.getRange(2,1,draftCount,17).getValues().forEach(function(r){states[r[0]]=r[16]||'DRAFT';});
  var values=count?sh.getRange(sh.getLastRow()-count+1,1,count,17).getValues():[];
  return {rows:values.reverse().filter(function(r){return r[1]&&r[16]==='Тийм';}).map(function(r){
    return {id:String(r[1]),name:String(r[3]),phone:String(r[4]),volume:r[8],state:states[r[1]]||'NEW'};
  })};
}
function khPalletSuggestion_(product,pieces){
  var dimensions=String(product).match(/\d+/g)||[];
  dimensions=dimensions.map(Number).sort(function(a,b){return a-b;});
  // Only confirmed packing data. Never extrapolate other SKUs or mixed loads.
  if(dimensions.length===3&&dimensions.join(',')==='100,240,600'&&Number.isInteger(Number(pieces))&&Number(pieces)>0)
    return Math.ceil(Number(pieces)/100);
  return null;
}
function salesDashboardLoad(id){
  khAuthorize_();id=khDashboardId_(id);
  khLock_(function(){
    var sh=khSalesSheet_();if(sh.getRange('A:A').createTextFinder(id).matchEntireCell(true).findNext())return;
    var leads=getTargetSheet_();var found=leads.getRange('B:B').createTextFinder(id).matchEntireCell(true).findNext();
    if(!found||found.getRow()<2)throw new Error('Хүсэлт олдсонгүй.');
    var r=leads.getRange(found.getRow(),1,1,17).getValues()[0];
    if(r[16]!=='Тийм')throw new Error('Мэдээлэл илгээх зөвшөөрөл байхгүй.');
    var product=r[7]||r[6];var pallets=khPalletSuggestion_(product,r[9]);
    sh.appendRow([id,r[3],r[4],r[12],r[5],product,r[8],r[9],KHANTAISHIR.priceM3,0,
      pallets===null?'':pallets,'','',7,'Подоны барьцааг подон буцаалтын нөхцөлөөр тооцно.','','DRAFT','','','','','','']
      .map(function(v){return typeof v==='string'?asCell_(v):v;}));
  });
  return khDashboardView_(khDashboardRow_(id));
}
function khDashboardView_(sel){
  var values=sel.sheet.getRange(sel.row,1,1,23).getValues()[0].slice(0,16);
  var record=khGetRecord_(String(values[0]));var state=record?record.state:'DRAFT';var valid=false,preview='';
  // A stored approval snapshot is authoritative even when somebody edits cells later.
  var same=true;if(record){try{same=khHash_(khReadQuote_(values))===record.hash;}catch(e){same=false;}}
  try{var q=record?record.snapshot:khReadQuote_(values);preview=khQuoteHtml_(q,record);valid=same;}catch(e){}
  return {values:values,token:khHash_(values),state:state,locked:!!record,valid:valid,preview:preview,
    pdf:record&&record.fileId?'https://drive.google.com/file/d/'+record.fileId+'/view':'',
    suggestion:khPalletSuggestion_(values[5],values[7])};
}
function salesDashboardSave(payload){
  khAuthorize_();if(!payload||!Array.isArray(payload.values)||payload.values.length!==16)throw new Error('Саналын бүтэц буруу.');
  var sel=khDashboardRow_(payload.id);
  khLock_(function(){
    var old=sel.sheet.getRange(sel.row,1,1,16).getValues()[0];
    if(khHash_(old)!==payload.token)throw new Error('Өөр ажилтан мэдээллийг өөрчилсөн. Хүсэлтээ дахин нээнэ үү.');
    if(khGetRecord_(payload.id))throw new Error('Баталсан санал засварлах боломжгүй.');
    var v=payload.values.slice();v[0]=old[0];var q=khReadQuote_(v);
    v=[q.requestId,q.name,q.phone,q.email,q.location,q.product,q.volume,q.pieces,q.unitPrice,q.discount,q.pallets,q.depositEach,q.tax,q.validDays,q.terms,'ТИЙМ'];
    sel.sheet.getRange(sel.row,1,1,16).setValues([v.map(function(x){return typeof x==='string'?asCell_(x):x;})]);
    sel.sheet.getRange(sel.row,17).setValue('DRAFT');SpreadsheetApp.flush();
  });
  return khDashboardView_(sel);
}
function salesDashboardSend(payload){
  var actor=khAuthorize_();if(!payload)throw new Error('Хүсэлт дутуу.');var sel=khDashboardRow_(payload.id);
  var values=sel.sheet.getRange(sel.row,1,1,16).getValues()[0];
  if(khHash_(values)!==payload.token)throw new Error('Мэдээлэл өөрчлөгдсөн. Хүсэлтээ дахин нээж хянана уу.');
  var q=khReadQuote_(values);khApproveAndSend_(sel,q,khHash_(q),actor);
  return khDashboardView_(sel);
}

function khDashboardHtml_(){return "<!doctype html><html lang=\"mn\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\"><style>\n*{box-sizing:border-box}body{margin:0;background:#f3f6f3;color:#173b2f;font:14px Arial,sans-serif}button,input,textarea{font:inherit}button{cursor:pointer;border:1px solid #cbd9cf;border-radius:9px;padding:10px 14px;background:white;color:#173b2f}button:hover{background:#e8f2eb}button:disabled{opacity:.5;cursor:wait}.primary{background:#126a45;color:white;border-color:#126a45}.primary:hover{background:#0c5033}header{padding:22px 26px;background:white;border-bottom:1px solid #dce5dd;display:flex;justify-content:space-between;gap:20px;align-items:center}h1{font-size:23px;margin:5px 0}.brand{font-size:10px;letter-spacing:2px;color:#53806a;font-weight:bold}.sub{color:#708378;font-size:12px}.stats{display:flex;gap:20px}.stats b{font-size:23px;display:block}.stats span{font-size:11px;color:#708378}.layout{display:grid;grid-template-columns:300px 1fr;min-height:530px}aside{padding:20px;border-right:1px solid #dce5dd}.search{width:100%;margin-bottom:12px}input,textarea{border:1px solid #cedbd1;border-radius:8px;background:white;padding:10px;width:100%;color:#173b2f}input:focus,textarea:focus{outline:2px solid #6ba587;outline-offset:1px}.request{display:block;width:100%;text-align:left;padding:13px;margin:8px 0;background:white}.request.active{border-color:#167c51;box-shadow:0 0 0 1px #167c51}.request strong{display:block;margin:6px 0}.request small{display:block;color:#708378;font-size:11px}.pill{display:inline-block;font-size:10px;background:#e6f2e9;color:#176443;border-radius:20px;padding:4px 7px}.pending{background:#fff2d6;color:#88671f}.main{padding:22px 26px}.empty{text-align:center;padding:100px 20px;color:#738579}.topline{display:flex;justify-content:space-between;align-items:flex-start;gap:12px}.topline h2{font-size:20px;margin:6px 0}.grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px;margin:18px 0}label{font-size:11px;color:#5c7466;display:block}label input,label textarea{margin-top:6px;font-size:13px}.wide{grid-column:1/-1}.two{grid-column:span 2}.hint{font-size:11px;line-height:1.5;color:#718476;margin-top:6px}.calc{background:#eaf4ed;border:1px solid #d5e9dc;border-radius:10px;padding:15px;display:flex;justify-content:space-between;gap:15px}.calc strong{font-size:23px}.calc small{display:block;color:#637e6d;margin-bottom:5px}.check{display:flex;gap:9px;align-items:center;margin:16px 0;font-size:12px}.check input{width:17px;height:17px;margin:0}.actions{display:flex;gap:10px;flex-wrap:wrap;border-top:1px solid #dde7df;padding-top:16px}.notice{margin:0;padding:12px 26px;font-size:12px;line-height:1.6;background:#e7f1ea}.error{background:#ffeded;color:#a12c2c}.success{background:#e4f6ea;color:#19683d}.hidden{display:none!important}iframe{width:100%;height:450px;border:1px solid #dce5dd;border-radius:10px;background:white;margin-top:18px}.readonly{background:#eef2ef}a{color:#126a45}#status{min-height:42px}button.link{font-size:11px;padding:4px 8px;margin-top:5px}#results{max-height:610px;overflow:auto}#editor[aria-busy=true]{opacity:.65;pointer-events:none}@media(max-width:780px){.layout{grid-template-columns:1fr}.stats{display:none}aside{border-right:0;border-bottom:1px solid #ddd}#results{max-height:200px}.grid{grid-template-columns:1fr 1fr}.two{grid-column:1/-1}}\n</style></head><body><header><div><div class=\"brand\">ХАНТАЙШИР ЭКО БЛОК</div><h1>Борлуулалтын самбар</h1><div class=\"sub\">Хүсэлтээ хянаж, баталгаатай санал илгээнэ.</div></div><div class=\"stats\"><div><b id=\"countAll\">—</b><span>ХҮСЭЛТ</span></div><div><b id=\"countPending\">—</b><span>ХЯНАЛТАД</span></div><div><b id=\"countSent\">—</b><span>ИЛГЭЭСЭН</span></div></div><button id=\"refresh\">Шинэчлэх ↻</button></header><p id=\"status\" class=\"notice\" role=\"status\">Хүсэлтүүдийг ачаалж байна…</p><div class=\"layout\"><aside><input class=\"search\" id=\"search\" placeholder=\"Нэр, утас, хүсэлтийн дугаар\" aria-label=\"Хүсэлт хайх\"><div class=\"hint\">Сүүлийн 200 хүсэлт</div><div id=\"results\"></div></aside><main class=\"main\"><div id=\"empty\" class=\"empty\">Хянах хүсэлтээ зүүн талаас сонгоно уу.</div><section id=\"editor\" class=\"hidden\" aria-busy=\"false\"><div class=\"topline\"><div><div id=\"requestId\" class=\"sub\"></div><h2 id=\"customerTitle\"></h2></div><span id=\"state\" class=\"pill\"></span></div><form id=\"form\"><div class=\"grid\"><label>Харилцагч<input data-i=\"1\" required maxlength=\"120\"></label><label>Утас<input data-i=\"2\" required maxlength=\"24\"></label><label>Хүлээн авагчийн имэйл<input data-i=\"3\" type=\"email\" required maxlength=\"160\"></label><label class=\"wide\">Байршил<input data-i=\"4\" required maxlength=\"120\"></label><label class=\"wide\">Блокийн хэмжээ / задаргаа<textarea data-i=\"5\" rows=\"2\" required maxlength=\"1000\"></textarea></label><label>Тоо хэмжээ, м³<input data-i=\"6\" type=\"number\" min=\"0.1\" max=\"100000\" step=\"any\" required></label><label>Нийт ширхэг<input data-i=\"7\" type=\"number\" min=\"1\" max=\"10000000\" step=\"1\" required></label><label>Нэг м³ үнэ, ₮<input data-i=\"8\" type=\"number\" min=\"1\" step=\"1\" required></label><label>Хөнгөлөлт, ₮<input data-i=\"9\" type=\"number\" min=\"0\" step=\"1\" required></label><label>Подоны тоо<input data-i=\"10\" type=\"number\" min=\"0\" step=\"1\" required><div id=\"palletHint\" class=\"hint\"></div><button id=\"usePallet\" type=\"button\" class=\"link hidden\">Тооцоог хэрэглэх</button></label><label>Нэг подоны барьцаа, ₮<input data-i=\"11\" type=\"number\" min=\"0\" step=\"1\" required></label><label class=\"two\">НӨАТ-ын нөхцөл<input data-i=\"12\" required maxlength=\"200\" placeholder=\"Жишээ: НӨАТ нэгж үнэд багтсан\"><div class=\"hint\">Төлөх бүх татварыг нэгж үнэд тусгана. Энэ талбараас нэмэлт татвар тооцохгүй.</div></label><label>Хүчинтэй хоног<input data-i=\"13\" type=\"number\" min=\"1\" max=\"365\" step=\"1\" required></label><label class=\"wide\">Подон буцаах болон бусад нөхцөл<textarea data-i=\"14\" rows=\"2\" maxlength=\"1500\"></textarea></label></div><div class=\"calc\"><div><small>Бүтээгдэхүүн − хөнгөлөлт</small><b id=\"goods\">—</b></div><div><small>Подоны барьцаа</small><b id=\"deposit\">—</b></div><div><small>Нийт төлөх дүн</small><strong id=\"total\">—</strong></div></div><label class=\"check\"><input id=\"verified\" type=\"checkbox\" required>Имэйл хаяг, үнэ, подон болон НӨАТ-ын нөхцөлийг шалгасан.</label><div class=\"actions\"><button type=\"submit\" id=\"save\">Хадгалж урьдчилан харах</button><button type=\"button\" id=\"send\" class=\"primary\" disabled>Баталж PDF-ээр имэйлдэх →</button><a id=\"pdfLink\" class=\"hidden\" target=\"_blank\" rel=\"noopener\">Баталсан PDF нээх ↗</a></div></form><iframe id=\"preview\" class=\"hidden\" title=\"Үнийн саналын урьдчилсан харагдац\" sandbox></iframe></section></main></div><script>\nconst $=id=>document.getElementById(id),inputs=[...document.querySelectorAll('[data-i]')];let rows=[],current=null,dirty=false,suggestion=null,busy=false;\nconst labels={DRAFT:'Хяналтад',APPROVED:'Баталсан',SENT:'Илгээсэн',SENDING:'Илгээлтийг шалгах',UNKNOWN:'Илгээлтийг шалгах',NEW:'Шинэ'};\nfunction msg(text,kind=''){ $('status').textContent=text;$('status').className='notice '+kind; }\nfunction rpc(name,arg){return new Promise((resolve,reject)=>google.script.run.withSuccessHandler(resolve).withFailureHandler(e=>reject(new Error(e.message||'Алдаа гарлаа')))[name](arg));}\nfunction lock(on){busy=on;$('editor').setAttribute('aria-busy',String(on));$('refresh').disabled=on;}\nfunction money(n){return Number.isFinite(n)?Math.round(n).toLocaleString('en-US')+' ₮':'—';}\nfunction input(i){return inputs.find(x=>+x.dataset.i===i)}\nfunction showRows(){let query=$('search').value.trim().toLowerCase();$('results').replaceChildren();rows.filter(r=>[r.id,r.name,r.phone].join(' ').toLowerCase().includes(query)).forEach(r=>{let b=document.createElement('button');b.className='request'+(current?.values[0]===r.id?' active':'');let badge=document.createElement('span');badge.className='pill'+(r.state==='NEW'?' pending':'');badge.textContent=labels[r.state]||r.state;let title=document.createElement('strong');title.textContent=r.name;let id=document.createElement('small');id.textContent=r.id;let info=document.createElement('small');info.textContent=r.phone+' · '+r.volume+' м³';b.append(badge,title,id,info);b.onclick=()=>open(r.id);$('results').append(b);});if(!$('results').children.length)$('results').textContent='Тохирох хүсэлт алга.';}\nasync function refresh(){if(busy)return;lock(true);try{let data=await rpc('salesDashboardData',null);rows=data.rows;$('countAll').textContent=rows.length;$('countPending').textContent=rows.filter(r=>['NEW','DRAFT','APPROVED'].includes(r.state)).length;$('countSent').textContent=rows.filter(r=>r.state==='SENT').length;showRows();msg('Зөвхөн борлуулалтын баталсан саналыг хэрэглэгчид илгээнэ.');}catch(e){msg(e.message,'error')}finally{lock(false)}}\nasync function open(id){if(busy)return;if(dirty&&!confirm('Хадгалаагүй өөрчлөлтийг орхиод өөр хүсэлт нээх үү?'))return;lock(true);try{apply(await rpc('salesDashboardLoad',id));msg(current.locked?'Энэ санал батлагдсан тул мэдээлэл түгжигдсэн. PDF болон төлөвийг харж болно.':'Мэдээллийг шалгаад хадгалж урьдчилан харна уу.');}catch(e){msg(e.message,'error')}finally{lock(false)}}\nfunction apply(data){current=data;dirty=false;let item=rows.find(r=>r.id===data.values[0]);if(item){item.state=data.state;item.name=data.values[1]}$('countPending').textContent=rows.filter(r=>['NEW','DRAFT','APPROVED'].includes(r.state)).length;$('countSent').textContent=rows.filter(r=>r.state==='SENT').length;$('empty').classList.add('hidden');$('editor').classList.remove('hidden');$('requestId').textContent=data.values[0];$('customerTitle').textContent=data.values[1];$('state').textContent=labels[data.state]||data.state;inputs.forEach(el=>{el.value=data.values[+el.dataset.i]??'';el.disabled=data.locked});$('verified').checked=String(data.values[15]).toUpperCase()==='ТИЙМ';$('verified').disabled=data.locked;$('save').disabled=data.locked;$('send').disabled=(data.locked&&data.state!=='APPROVED')||!data.valid;$('preview').classList.toggle('hidden',!data.preview);$('preview').srcdoc=data.preview||'';$('pdfLink').classList.toggle('hidden',!data.pdf);if(data.pdf)$('pdfLink').href=data.pdf;showRows();calculate();}\nfunction calculate(){let n=i=>input(i).value===''?NaN:Number(input(i).value),goods=Math.round(n(6)*n(8))-n(9),deposit=n(10)*n(11);$('goods').textContent=money(goods);$('deposit').textContent=money(deposit);$('total').textContent=money(goods+deposit);let sizes=(input(5).value.match(/\\d+/g)||[]).map(Number).sort((a,b)=>a-b);suggestion=sizes.length===3&&sizes.join(',')==='100,240,600'&&Number.isInteger(n(7))&&n(7)>0?Math.ceil(n(7)/100):null;$('palletHint').textContent=suggestion!==null?'100 ш / подон → '+suggestion+' подон. Сүүлийн дутуу подоныг бүтнээр тооцов.':'Энэ хэмжээний подоны багтаамж баталгаажаагүй. Гараар шалгаж оруулна.';$('usePallet').classList.toggle('hidden',suggestion===null||current?.locked);}\nfunction edited(event){if(event&&event.target!==$('verified'))$('verified').checked=false;dirty=true;$('send').disabled=true;$('preview').classList.add('hidden');calculate();msg('Өөрчлөлт хадгалагдаагүй. Хадгалсны дараа батлах боломжтой.');}\ninputs.forEach(el=>el.addEventListener('input',edited));$('verified').onchange=edited;$('usePallet').onclick=()=>{input(10).value=suggestion;edited()};\n$('form').onsubmit=async e=>{e.preventDefault();if(busy||!current)return;let values=[...current.values];inputs.forEach(el=>values[+el.dataset.i]=el.type==='number'?(el.value===''?'':Number(el.value)):el.value);values[15]=$('verified').checked?'ТИЙМ':'';lock(true);try{apply(await rpc('salesDashboardSave',{id:values[0],values,token:current.token}));msg('Хадгаллаа. Урьдчилсан саналыг хянаад батална уу.','success')}catch(e){msg(e.message,'error')}finally{lock(false)}};\n$('send').onclick=async()=>{if(busy||dirty||!current||(current.locked&&current.state!=='APPROVED'))return;if(!confirm(current.values[1]+'\\nИмэйл: '+current.values[3]+'\\nНийт: '+$('total').textContent+'\\n\\nБаталж, PDF саналыг энэ хаяг руу илгээх үү?'))return;lock(true);try{apply(await rpc('salesDashboardSend',{id:current.values[0],token:current.token}));msg('Баталсан PDF саналыг имэйлээр илгээлээ.','success')}catch(e){msg(e.message,'error')}finally{lock(false)}};\n$('refresh').onclick=refresh;$('search').oninput=showRows;refresh();\n</script></body></html>\n";}
