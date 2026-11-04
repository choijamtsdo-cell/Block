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

function saveLead_(q, internalActor) {
  if(internalActor&&khAuthorize_()!==internalActor)throw new Error("Бүртгэсэн ажилтны эрх буруу.");
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
    if(internalActor)khLeadOwnerColumn_(sh);
    var lastRow = sh.getLastRow();
    if (lastRow > 1) {
      var match = sh.getRange(2, 2, lastRow - 1, 1)
        .createTextFinder(id).matchEntireCell(true).findNext();
      if (match) {
        if(internalActor&&String(sh.getRange(match.getRow(),18).getValue())!==internalActor)throw new Error('Энэ хүсэлтийг өөр ажилтан бүртгэсэн.');
        return {requestId: id, alreadySaved: true};
      }
    }
    var leadRow = [
      new Date(), id, 'Шинэ', asCell_(name), asCell_(phone),
      asCell_(place), asCell_(product), asCell_(breakdown),
      volume, Math.round(pieces), KHANTAISHIR.priceM3, total,
      asCell_(email), asCell_(notes), asCell_(source), asCell_(utm), 'Тийм'
    ];
    if(internalActor)leadRow.push(internalActor);
    sh.appendRow(leadRow);
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
        body:'Сайн байна уу, '+q.name+'?\n\nБорлуулалтын албанаас баталгаажуулсан '+r.number+' дугаартай үнийн саналыг хавсаргалаа.\nНийт: '+khMoney_(q.total)+' ₮\nПодоны барьцаа: '+khMoney_(q.depositTotal)+' ₮\n\nХавсралтын нөхцөл болон хүчинтэй хугацааг харна уу.\nХантайшир эко блок ХХК / РД:3267806 /',
        name:'Хантайшир эко блок ХХК',attachments:[attachment]});
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
/** Supplier identity is server-owned and cannot be overridden by customer input. */
function khQuoteItems_(q){
  var items=[],re=/(Гадна|Дотор)\s+(\d+)\s*мм:\s*([\d,.]+)\s*м³/g,match;
  while((match=re.exec(q.product))!==null){
    var volume=Number(match[3].replace(/,/g,''));
    if(volume>0)items.push({name:'AAC блок - '+match[1].toLowerCase()+' хана',spec:'600 × 240 × '+match[2]+' мм',volume:volume});
  }
  var sum=items.reduce(function(n,item){return n+item.volume;},0);
  if(!items.length||Math.abs(sum-q.volume)>0.00011)items=[{name:'AAC ханын блок',spec:String(q.product).replace(/\s*\([\d,.\s]+ ш\)/g,''),volume:q.volume}];
  var allocated=0;items.forEach(function(item,i){item.total=i===items.length-1?q.subtotal-allocated:Math.round(item.volume*q.unitPrice);allocated+=item.total;});
  return items;
}
function khQuoteHtml_(q,r){
  var e=khEscape_,assets=khQuoteBrandAssets_(),approved=r?new Date(r.approvedAt):new Date();
  var until=new Date(approved.getTime()+q.validDays*86400000);
  var date=function(d){return Utilities.formatDate(d,KH_SALES.timezone,'yyyy.MM.dd');};
  var money=function(n){return khMoney_(n)+' ₮';};
  var isTest=!!r&&/TEST|ТУРШИЛТ/.test(r.number);
  var state=isTest?'ТУРШИЛТ - АЛБАН САНАЛ БИШ':r?'БОРЛУУЛАЛТ БАТАЛСАН':'УРЬДЧИЛСАН - БАТАЛГААЖААГҮЙ';
  var rows=khQuoteItems_(q).map(function(item,i){return '<tr'+(i%2?' bgcolor="#f5f8f5"':'')+'><td class="index">'+('0'+(i+1)).slice(-2)+'</td><td class="description"><b>'+e(item.name)+'</b><span class="spec">'+e(item.spec)+'</span></td><td>'+e(item.volume)+'</td><td>'+khMoney_(q.unitPrice)+'</td><td><b>'+khMoney_(item.total)+'</b></td></tr>';}).join('');
  var totalRow=function(label,value){return '<tr><td class="summary-label">'+e(label)+'</td><td class="summary-value">'+e(value)+'</td></tr>';};
  var notice=isTest?'Энэ нь зөвхөн загварын туршилт. Төлбөрийн болон худалдааны баримт биш. Имэйл илгээгээгүй.':r?'Энэхүү үнийн санал нь дээрх хүчинтэй хугацаанд үйлчилнэ. Нийт төлөх дүнд бүтээгдэхүүний төлбөр болон подоны буцаан олгох барьцааг тусад нь тусгасан.':'Энэхүү тооцоолол нь урьдчилсан үнэ юм. Албан үнийн саналыг борлуулалтын алба хянаж баталгаажуулсны дараа илгээнэ.';
  return `<!doctype html><html lang="mn"><head><meta charset="utf-8"><style>
@page{size:A4;margin:14mm 16mm}*{box-sizing:border-box}body{margin:0;color:#1d2521;font:11px Arial,sans-serif;line-height:1.5;-webkit-print-color-adjust:exact;print-color-adjust:exact}table{width:100%;border-collapse:collapse}td,th{vertical-align:top}h1{margin:0;font-size:28px;letter-spacing:.3px}p{margin:5px 0}.stripe td{height:6px;line-height:6px;padding:0}.header{margin-top:19px;border-bottom:1px solid #dce7df}.header td{padding:0 0 16px}.brand{width:188px;height:93px;object-fit:contain}.heading{text-align:right;padding-top:12px!important}.subtitle{font-size:10px;font-weight:bold;color:#208e43;margin:6px 0 10px}.tag{color:#208e43;font-size:8px;font-weight:bold;padding:5px 10px;background:#f1f7f2}.meta{border-bottom:1px solid #dce7df;margin-bottom:17px}.meta td{padding:13px 0 15px;width:33%}.micro{color:#68726c;font-size:8px;font-weight:bold;letter-spacing:.3px}.meta strong{display:block;margin-top:7px;font-size:11px}.client{border-left:4px solid #208e43;margin-bottom:18px}.client td{padding:13px 14px}.client .supplier{border-left:1px solid #dce7df;width:41%}.client strong{display:block;font-size:12px;margin:7px 0}.small{font-size:9px;color:#68726c}.section{margin-bottom:10px}.section td{padding:0}.section-title{font-size:13px;font-weight:bold;border-left:5px solid #f69c22;padding-left:9px!important}.items{font-size:10px;table-layout:fixed}.items th{background:#208e43;color:white;font-size:9px;text-align:right;padding:11px 7px}.items th:nth-child(1){width:7%;text-align:left}.items th:nth-child(2){width:39%;text-align:left}.items th:nth-child(3){width:12%}.items th:nth-child(4){width:20%}.items th:nth-child(5){width:22%}.items td{border-bottom:1px solid #dce7df;padding:13px 7px;text-align:right;white-space:nowrap}.items td.description{text-align:left;white-space:normal;overflow-wrap:anywhere}.items td.index{text-align:left;color:#68726c}.spec{display:block;color:#68726c;font-size:9px;margin-top:5px;line-height:1.5}.summary{margin-top:18px}.terms-cell{width:54%;padding:14px 16px;background:#f1f7f2;overflow-wrap:anywhere}.terms-cell h3{font-size:10px;color:#208e43;margin:0 0 9px}.terms-cell p{font-size:9px;line-height:1.65}.gap{width:12px}.amount-cell{background:#f5f8f5;padding:8px 0 0}.totals td{padding:5px 11px;font-size:9px}.summary-label{color:#68726c}.summary-value{text-align:right;white-space:nowrap;font-weight:bold}.grand td{background:#208e43;color:white;padding:12px 11px;font-weight:bold}.grand .summary-value{font-size:16px}.notice{border:1px solid #dce7df;border-left:4px solid #f69c22;margin-top:17px;padding:10px 12px;font-size:9px;line-height:1.55}.notice b{display:block;color:#208e43;font-size:8px;margin-bottom:3px}.footer{margin-top:17px}.footer td{vertical-align:middle}.qr{width:72px;height:72px}.social-copy{padding-left:11px;width:45%;font-size:9px}.social-copy b{font-size:9px}.social-copy p{color:#68726c;font-size:8px}.social-copy a{color:#208e43;text-decoration:none;font-size:8px}.approval{font-size:9px;overflow-wrap:anywhere}.approval .rule{border-bottom:1px solid #a6b8aa;margin:10px 0 7px}.bottom{margin-top:13px;border-top:1px solid #dce7df}.bottom td{padding-top:9px;font-size:8px;color:#68726c}.bottom b{color:#208e43}.fineprint{font-size:8px;color:#68726c;margin-top:8px;overflow-wrap:anywhere}tr{page-break-inside:avoid}.client,.summary,.notice,.footer{page-break-inside:avoid}thead{display:table-header-group}
</style></head><body>
<table class="stripe"><tr><td bgcolor="#208e43" width="83%"></td><td bgcolor="#f69c22" width="17%"></td></tr></table>
<table class="header"><tr><td><img class="brand" src="${assets.logo}" alt="Хантайшир Эко Блок лого"></td><td class="heading"><h1>ҮНИЙН САНАЛ</h1><p class="subtitle">ХАНТАЙШИР ЭКО БЛОК / AAC</p><span class="tag">${e(state)}</span></td></tr></table>
<table class="meta"><tr><td><div class="micro">САНАЛЫН ДУГААР</div><strong>${e(r?r.number:q.requestId)}</strong></td><td><div class="micro">ОГНОО</div><strong>${date(approved)}</strong></td><td><div class="micro">ХҮЧИНТЭЙ ХУГАЦАА</div><strong>${date(until)} хүртэл (${q.validDays} хоног)</strong></td></tr></table>
<table class="client" bgcolor="#f5f8f5"><tr><td><div class="micro" style="color:#208e43">ҮНИЙН САНАЛ ХҮЛЭЭН АВАГЧ</div><strong>${e(q.name)}</strong><div class="small">${e(q.phone)} · ${e(q.email)}</div><div class="small">${e(q.location)}</div></td><td class="supplier"><div class="micro">НИЙЛҮҮЛЭГЧ</div><strong>Хантайшир эко блок ХХК</strong><div class="small"><b>РД:3267806</b></div><div class="small">Говь-Алтай аймаг, Тайшир сум</div><div class="small">Утас: 8600 0254 · 8811 3350</div></td></tr></table>
<table class="section"><tr><td class="section-title">БҮТЭЭГДЭХҮҮН, ҮНИЙН ТООЦОО</td><td style="text-align:right" class="small">Нэгж үнэ: ₮ / м³</td></tr></table>
<table class="items"><thead><tr><th bgcolor="#208e43">№</th><th bgcolor="#208e43">БҮТЭЭГДЭХҮҮН / ХЭМЖЭЭ</th><th bgcolor="#208e43">М³</th><th bgcolor="#208e43">НЭГЖ ҮНЭ</th><th bgcolor="#208e43">НИЙТ ДҮН</th></tr></thead><tbody>${rows}</tbody></table>
<table class="summary"><tr><td class="terms-cell" bgcolor="#f1f7f2"><h3>ҮНИЙН НӨХЦӨЛ</h3><p><b>НӨАТ:</b> ${e(q.tax)}</p><p>Подоны буцаан олгох барьцаа: ${e(q.pallets)} подон × ${e(money(q.depositEach))}.</p><p style="white-space:pre-wrap">${e(q.terms)}</p><p>Тооцооны нэгж: м³</p></td><td class="gap"></td><td class="amount-cell" bgcolor="#f5f8f5"><table class="totals">${totalRow('Блокийн нийт хэмжээ',q.volume+' м³')}${totalRow('Бүтээгдэхүүний үнэ',money(q.subtotal))}${totalRow('Хөнгөлөлт',money(q.discount))}${totalRow('Бүтээгдэхүүний төлбөр',money(q.goodsTotal))}${totalRow('Подоны барьцаа',money(q.depositTotal))}<tr class="grand"><td bgcolor="#208e43">НИЙТ ТӨЛӨХ ДҮН</td><td bgcolor="#208e43" class="summary-value">${e(money(q.total))}</td></tr></table></td></tr></table>
<div class="notice"><b>АНХААРАХ</b>${e(notice)}</div>
<table class="footer"><tr><td style="width:72px"><a href="https://www.facebook.com/Khantaishir.ecoblock"><img class="qr" src="${assets.qr}" alt="Facebook хуудасны QR"></a></td><td class="social-copy"><b>БИДНИЙ FACEBOOK ХУУДАС</b><p>QR код уншуулж дэлгэрэнгүй<br>мэдээлэл аваарай.</p><a href="https://www.facebook.com/Khantaishir.ecoblock">facebook.com/Khantaishir.ecoblock</a></td><td class="approval"><b>Баталгаажуулсан:</b><div class="rule"></div>${r?e(r.approvedBy):'Борлуулалтын алба баталгаажуулна.'}<p class="small">${r?date(approved)+' · Борлуулалтын цахим баталгаажуулалт':'Урьдчилсан харагдац - албан санал биш'}</p></td></tr></table>
<table class="bottom"><tr><td><b>Хантайшир эко блок ХХК / РД:3267806 /</b></td><td style="text-align:right">Тайшир, Говь-Алтай · AAC блок үйлдвэрлэл</td></tr></table><p class="fineprint">Хүсэлтийн дугаар: ${e(q.requestId)}</p>
</body></html>`;
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
  if(draftCount)drafts.getRange(2,1,draftCount,17).getValues().forEach(function(r){states[r[0]]=r[16]==='ХЯНАЛТАД'?'DRAFT':r[16]||'DRAFT';});
  var values=count?sh.getRange(sh.getLastRow()-count+1,1,count,18).getValues():[];
  return {rows:values.reverse().filter(function(r){return r[1]&&r[16]==='Тийм';}).map(function(r){
    return {id:String(r[1]),name:String(r[3]),phone:String(r[4]),volume:r[8],state:states[r[1]]||'NEW',source:r[17]?'Утасны хүсэлт':'Веб хүсэлт',createdBy:String(r[17]||'')};
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

function khDashboardHtml_(){return "<!doctype html><html lang=\"mn\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\"><style>\n*{box-sizing:border-box}body{margin:0;background:#f3f6f3;color:#173b2f;font:14px Arial,sans-serif}button,input,textarea,select{font:inherit}button{cursor:pointer;border:1px solid #cbd9cf;border-radius:9px;padding:10px 14px;background:white;color:#173b2f}button:hover{background:#e8f2eb}button:disabled{opacity:.5;cursor:wait}.primary{background:#126a45;color:white;border-color:#126a45}.primary:hover{background:#0c5033}header{padding:22px 26px;background:white;border-bottom:1px solid #dce5dd;display:flex;justify-content:space-between;gap:20px;align-items:center}h1{font-size:23px;margin:5px 0}.brand{font-size:10px;letter-spacing:2px;color:#53806a;font-weight:bold}.sub{color:#708378;font-size:12px}.stats{display:flex;gap:20px}.stats b{font-size:23px;display:block}.stats span{font-size:11px;color:#708378}.layout{display:grid;grid-template-columns:300px 1fr;min-height:530px}aside{padding:20px;border-right:1px solid #dce5dd}.search{width:100%;margin-bottom:12px}input,textarea,select{border:1px solid #cedbd1;border-radius:8px;background:white;padding:10px;width:100%;color:#173b2f}input:focus,textarea:focus{outline:2px solid #6ba587;outline-offset:1px}.request{display:block;width:100%;text-align:left;padding:13px;margin:8px 0;background:white}.request.active{border-color:#167c51;box-shadow:0 0 0 1px #167c51}.request strong{display:block;margin:6px 0}.request small{display:block;color:#708378;font-size:11px}.pill{display:inline-block;font-size:10px;background:#e6f2e9;color:#176443;border-radius:20px;padding:4px 7px}.pending{background:#fff2d6;color:#88671f}.main{padding:22px 26px}.empty{text-align:center;padding:100px 20px;color:#738579}.topline{display:flex;justify-content:space-between;align-items:flex-start;gap:12px}.topline h2{font-size:20px;margin:6px 0}.grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px;margin:18px 0}label{font-size:11px;color:#5c7466;display:block}label input,label textarea{margin-top:6px;font-size:13px}.wide{grid-column:1/-1}.two{grid-column:span 2}.hint{font-size:11px;line-height:1.5;color:#718476;margin-top:6px}.calc{background:#eaf4ed;border:1px solid #d5e9dc;border-radius:10px;padding:15px;display:flex;justify-content:space-between;gap:15px}.calc strong{font-size:23px}.calc small{display:block;color:#637e6d;margin-bottom:5px}.check{display:flex;gap:9px;align-items:center;margin:16px 0;font-size:12px}.check input{width:17px;height:17px;margin:0}.actions{display:flex;gap:10px;flex-wrap:wrap;border-top:1px solid #dde7df;padding-top:16px}.notice{margin:0;padding:12px 26px;font-size:12px;line-height:1.6;background:#e7f1ea}.error{background:#ffeded;color:#a12c2c}.success{background:#e4f6ea;color:#19683d}.hidden{display:none!important}iframe{width:100%;height:450px;border:1px solid #dce5dd;border-radius:10px;background:white;margin-top:18px}.readonly{background:#eef2ef}a{color:#126a45}#status{min-height:42px}button.link{font-size:11px;padding:4px 8px;margin-top:5px}#results{max-height:610px;overflow:auto}#editor[aria-busy=true]{opacity:.65;pointer-events:none}@media(max-width:780px){.layout{grid-template-columns:1fr}.stats{display:none}aside{border-right:0;border-bottom:1px solid #ddd}#results{max-height:200px}.grid{grid-template-columns:1fr 1fr}.two{grid-column:1/-1}}\n</style></head><body><header><div><div class=\"brand\">ХАНТАЙШИР ЭКО БЛОК</div><h1>Борлуулалтын самбар</h1><div class=\"sub\">Хүсэлтээ хянаж, баталгаатай санал илгээнэ.</div></div><div class=\"stats\"><div><b id=\"countAll\">—</b><span>ХҮСЭЛТ</span></div><div><b id=\"countPending\">—</b><span>ХЯНАЛТАД</span></div><div><b id=\"countSent\">—</b><span>ИЛГЭЭСЭН</span></div></div><button id=\"newPhone\" class=\"primary\">+ Утасны хүсэлт бүртгэх</button><button id=\"refresh\">Шинэчлэх ↻</button></header><p id=\"status\" class=\"notice\" role=\"status\">Хүсэлтүүдийг ачаалж байна…</p><div class=\"layout\"><aside><input class=\"search\" id=\"search\" placeholder=\"Нэр, утас, хүсэлтийн дугаар\" aria-label=\"Хүсэлт хайх\"><div class=\"hint\">Сүүлийн 200 хүсэлт</div><div id=\"results\"></div></aside><main class=\"main\"><section id=\"phoneEntry\" class=\"hidden\"><div class=\"topline\"><div><div class=\"brand\">УТАСНЫ ХҮСЭЛТ</div><h2>Шинэ үнийн хүсэлт бүртгэх</h2></div><button id=\"cancelPhone\" type=\"button\">Болих</button></div><p class=\"hint\">Бүртгэсэн огноо, ажилтны Google хаяг автоматаар хадгалагдана. Үнэ болон барьцааг дараагийн дэлгэцэд хянана.</p><form id=\"phoneForm\"><div class=\"grid\"><label>Харилцагч / байгууллага<input name=\"name\" required maxlength=\"120\"></label><label>Утас<input name=\"phone\" type=\"tel\" required maxlength=\"24\" placeholder=\"8 оронтой утас\"></label><label>Имэйл (байвал)<input name=\"email\" type=\"email\" maxlength=\"160\"></label><label class=\"wide\">Төслийн байршил<input name=\"location\" required maxlength=\"120\"></label><label class=\"two\">Блокийн хэмжээ<select name=\"product\"><option>600 × 240 × 300 мм</option><option>600 × 240 × 200 мм</option></select></label><label>Хэмжээ, м³<input name=\"volume\" type=\"number\" required min=\"0.1\" max=\"100000\" step=\"any\"></label><label class=\"wide\">Хэрэглэгчийн хүсэлт / тайлбар<textarea name=\"notes\" rows=\"3\" maxlength=\"700\"></textarea></label></div><label class=\"check\"><input type=\"checkbox\" name=\"consent\" required>Харилцагчийн мэдээлэл бүртгэж, үнийн саналын талаар эргэн холбогдох зөвшөөрөлтэй.</label><div class=\"actions\"><button class=\"primary\" type=\"submit\">Sheet-д бүртгэж санал бэлтгэх →</button></div></form></section><div id=\"empty\" class=\"empty\">Хянах хүсэлтээ зүүн талаас сонгоно уу.</div><section id=\"editor\" class=\"hidden\" aria-busy=\"false\"><div class=\"topline\"><div><div id=\"requestId\" class=\"sub\"></div><h2 id=\"customerTitle\"></h2></div><span id=\"state\" class=\"pill\"></span></div><form id=\"form\"><div class=\"grid\"><label>Харилцагч<input data-i=\"1\" required maxlength=\"120\"></label><label>Утас<input data-i=\"2\" required maxlength=\"24\"></label><label>Хүлээн авагчийн имэйл<input data-i=\"3\" type=\"email\" required maxlength=\"160\"></label><label class=\"wide\">Байршил<input data-i=\"4\" required maxlength=\"120\"></label><label class=\"wide\">Блокийн хэмжээ / задаргаа<textarea data-i=\"5\" rows=\"2\" required maxlength=\"1000\"></textarea></label><label>Тоо хэмжээ, м³<input data-i=\"6\" type=\"number\" min=\"0.1\" max=\"100000\" step=\"any\" required></label><label>Дотоод тооцооны ширхэг (подонд)<input data-i=\"7\" type=\"number\" min=\"1\" max=\"10000000\" step=\"1\" required></label><label>Нэг м³ үнэ, ₮<input data-i=\"8\" type=\"number\" min=\"1\" step=\"1\" required></label><label>Хөнгөлөлт, ₮<input data-i=\"9\" type=\"number\" min=\"0\" step=\"1\" required></label><label>Подоны тоо<input data-i=\"10\" type=\"number\" min=\"0\" step=\"1\" required><div id=\"palletHint\" class=\"hint\"></div><button id=\"usePallet\" type=\"button\" class=\"link hidden\">Тооцоог хэрэглэх</button></label><label>Нэг подоны барьцаа, ₮<input data-i=\"11\" type=\"number\" min=\"0\" step=\"1\" required></label><label class=\"two\">НӨАТ-ын нөхцөл<input data-i=\"12\" required maxlength=\"200\" placeholder=\"Жишээ: НӨАТ нэгж үнэд багтсан\"><div class=\"hint\">Төлөх бүх татварыг нэгж үнэд тусгана. Энэ талбараас нэмэлт татвар тооцохгүй.</div></label><label>Хүчинтэй хоног<input data-i=\"13\" type=\"number\" min=\"1\" max=\"365\" step=\"1\" required></label><label class=\"wide\">Подон буцаах болон бусад нөхцөл<textarea data-i=\"14\" rows=\"2\" maxlength=\"1500\"></textarea></label></div><div class=\"calc\"><div><small>Бүтээгдэхүүн − хөнгөлөлт</small><b id=\"goods\">—</b></div><div><small>Подоны барьцаа</small><b id=\"deposit\">—</b></div><div><small>Нийт төлөх дүн</small><strong id=\"total\">—</strong></div></div><label class=\"check\"><input id=\"verified\" type=\"checkbox\" required>Имэйл хаяг, үнэ, подон болон НӨАТ-ын нөхцөлийг шалгасан.</label><div class=\"actions\"><button type=\"submit\" id=\"save\">Хадгалж урьдчилан харах</button><button type=\"button\" id=\"send\" class=\"primary\" disabled>Баталж PDF-ээр имэйлдэх →</button><a id=\"pdfLink\" class=\"hidden\" target=\"_blank\" rel=\"noopener\">Баталсан PDF нээх ↗</a></div></form><iframe id=\"preview\" class=\"hidden\" title=\"Үнийн саналын урьдчилсан харагдац\" sandbox></iframe></section></main></div><script>\nconst $=id=>document.getElementById(id),inputs=[...document.querySelectorAll('[data-i]')];let rows=[],current=null,dirty=false,suggestion=null,busy=false,phoneDraftKey=null;\nconst labels={DRAFT:'Хяналтад',APPROVED:'Баталсан',SENT:'Илгээсэн',SENDING:'Илгээлтийг шалгах',UNKNOWN:'Илгээлтийг шалгах',NEW:'Шинэ'};\nfunction msg(text,kind=''){ $('status').textContent=text;$('status').className='notice '+kind; }\nfunction rpc(name,arg){return new Promise((resolve,reject)=>google.script.run.withSuccessHandler(resolve).withFailureHandler(e=>reject(new Error(e.message||'Алдаа гарлаа')))[name](arg));}\nfunction lock(on){busy=on;$('editor').setAttribute('aria-busy',String(on));$('refresh').disabled=on;$('newPhone').disabled=on;[...$('phoneForm').elements].forEach(el=>el.disabled=on);$('cancelPhone').disabled=on;}\nfunction money(n){return Number.isFinite(n)?Math.round(n).toLocaleString('en-US')+' ₮':'—';}\nfunction input(i){return inputs.find(x=>+x.dataset.i===i)}\nfunction showRows(){let query=$('search').value.trim().toLowerCase();$('results').replaceChildren();rows.filter(r=>[r.id,r.name,r.phone].join(' ').toLowerCase().includes(query)).forEach(r=>{let b=document.createElement('button');b.className='request'+(current?.values[0]===r.id?' active':'');let badge=document.createElement('span');badge.className='pill'+(r.state==='NEW'?' pending':'');badge.textContent=labels[r.state]||r.state;let title=document.createElement('strong');title.textContent=r.name;let id=document.createElement('small');id.textContent=r.id;let info=document.createElement('small');info.textContent=r.phone+' · '+r.volume+' м³';let source=document.createElement('small');source.textContent=r.source||'Веб хүсэлт';b.append(badge,title,id,info,source);b.onclick=()=>open(r.id);$('results').append(b);});if(!$('results').children.length)$('results').textContent='Тохирох хүсэлт алга.';}\nasync function refresh(){if(busy)return;lock(true);try{let data=await rpc('salesDashboardData',null);rows=data.rows;$('countAll').textContent=rows.length;$('countPending').textContent=rows.filter(r=>['NEW','DRAFT','APPROVED'].includes(r.state)).length;$('countSent').textContent=rows.filter(r=>r.state==='SENT').length;showRows();msg('Зөвхөн борлуулалтын баталсан саналыг хэрэглэгчид илгээнэ.');}catch(e){msg(e.message,'error')}finally{lock(false)}}\nasync function open(id){if(busy)return;if(dirty&&!confirm('Хадгалаагүй өөрчлөлтийг орхиод өөр хүсэлт нээх үү?'))return;lock(true);try{apply(await rpc('salesDashboardLoad',id));msg(current.locked?'Энэ санал батлагдсан тул мэдээлэл түгжигдсэн. PDF болон төлөвийг харж болно.':'Мэдээллийг шалгаад хадгалж урьдчилан харна уу.');}catch(e){msg(e.message,'error')}finally{lock(false)}}\nfunction apply(data){$('phoneEntry').classList.add('hidden');current=data;dirty=false;let item=rows.find(r=>r.id===data.values[0]);if(item){item.state=data.state;item.name=data.values[1]}$('countPending').textContent=rows.filter(r=>['NEW','DRAFT','APPROVED'].includes(r.state)).length;$('countSent').textContent=rows.filter(r=>r.state==='SENT').length;$('empty').classList.add('hidden');$('editor').classList.remove('hidden');$('requestId').textContent=data.values[0]+' · '+(data.source||item?.source||'Веб хүсэлт')+((data.createdBy||item?.createdBy)?' · Бүртгэсэн: '+(data.createdBy||item.createdBy):'');$('customerTitle').textContent=data.values[1];$('state').textContent=labels[data.state]||data.state;inputs.forEach(el=>{el.value=data.values[+el.dataset.i]??'';el.disabled=data.locked});$('verified').checked=String(data.values[15]).toUpperCase()==='ТИЙМ';$('verified').disabled=data.locked;$('save').disabled=data.locked;$('send').disabled=(data.locked&&data.state!=='APPROVED')||!data.valid;$('preview').classList.toggle('hidden',!data.preview);$('preview').srcdoc=data.preview||'';$('pdfLink').classList.toggle('hidden',!data.pdf);if(data.pdf)$('pdfLink').href=data.pdf;showRows();calculate();}\nfunction calculate(){let n=i=>input(i).value===''?NaN:Number(input(i).value),goods=Math.round(n(6)*n(8))-n(9),deposit=n(10)*n(11);$('goods').textContent=money(goods);$('deposit').textContent=money(deposit);$('total').textContent=money(goods+deposit);let sizes=(input(5).value.match(/\\d+/g)||[]).map(Number).sort((a,b)=>a-b);suggestion=sizes.length===3&&sizes.join(',')==='100,240,600'&&Number.isInteger(n(7))&&n(7)>0?Math.ceil(n(7)/100):null;$('palletHint').textContent=suggestion!==null?'100 ш / подон → '+suggestion+' подон. Сүүлийн дутуу подоныг бүтнээр тооцов.':'Энэ хэмжээний подоны багтаамж баталгаажаагүй. Гараар шалгаж оруулна.';$('usePallet').classList.toggle('hidden',suggestion===null||current?.locked);}\nfunction edited(event){if(event&&event.target!==$('verified'))$('verified').checked=false;dirty=true;$('send').disabled=true;$('preview').classList.add('hidden');calculate();msg('Өөрчлөлт хадгалагдаагүй. Хадгалсны дараа батлах боломжтой.');}\ninputs.forEach(el=>el.addEventListener('input',edited));$('verified').onchange=edited;$('usePallet').onclick=()=>{input(10).value=suggestion;edited()};\n$('form').onsubmit=async e=>{e.preventDefault();if(busy||!current)return;let values=[...current.values];inputs.forEach(el=>values[+el.dataset.i]=el.type==='number'?(el.value===''?'':Number(el.value)):el.value);values[15]=$('verified').checked?'ТИЙМ':'';lock(true);try{apply(await rpc('salesDashboardSave',{id:values[0],values,token:current.token}));msg('Хадгаллаа. Урьдчилсан саналыг хянаад батална уу.','success')}catch(e){msg(e.message,'error')}finally{lock(false)}};\n$('send').onclick=async()=>{if(busy||dirty||!current||(current.locked&&current.state!=='APPROVED'))return;if(!confirm(current.values[1]+'\\nИмэйл: '+current.values[3]+'\\nНийт: '+$('total').textContent+'\\n\\nБаталж, PDF саналыг энэ хаяг руу илгээх үү?'))return;lock(true);try{apply(await rpc('salesDashboardSend',{id:current.values[0],token:current.token}));msg('Баталсан PDF саналыг имэйлээр илгээлээ.','success')}catch(e){msg(e.message,'error')}finally{lock(false)}};\n\n$('newPhone').onclick=()=>{if(busy)return;if(dirty&&!confirm('Хадгалаагүй өөрчлөлтийг орхиод утасны хүсэлт бүртгэх үү?'))return;$('phoneForm').reset();phoneDraftKey=Array.from(crypto.getRandomValues(new Uint8Array(16))).map(n=>n.toString(16).padStart(2,'0')).join('');dirty=false;$('editor').classList.add('hidden');$('empty').classList.add('hidden');$('phoneEntry').classList.remove('hidden');msg('Утсаар ирсэн хүсэлтээ бүртгэнэ. Бүртгэлийн үед имэйл илгээхгүй.');};\n$('phoneForm').oninput=()=>{dirty=true};\n$('cancelPhone').onclick=()=>{if(busy)return;if(dirty&&!confirm('Бүртгээгүй мэдээллийг орхих уу?'))return;dirty=false;$('phoneEntry').classList.add('hidden');if(current)apply(current);else $('empty').classList.remove('hidden');msg('Хүсэлтээ сонгох эсвэл шинэ хүсэлт бүртгэнэ үү.');};\n$('phoneForm').onsubmit=async e=>{e.preventDefault();if(busy)return;const form=$('phoneForm'),payload={draftKey:phoneDraftKey};for(const name of ['name','phone','email','location','product','notes'])payload[name]=form.elements[name].value.trim();payload.volume=Number(form.elements.volume.value);payload.consent=form.elements.consent.checked;lock(true);try{let data=await rpc('salesDashboardCreatePhone',payload);rows.unshift({id:data.values[0],name:data.values[1],phone:data.values[2],volume:data.values[6],state:data.state,source:'Утасны хүсэлт',createdBy:data.createdBy});rows=rows.filter((r,i,all)=>all.findIndex(x=>x.id===r.id)===i).slice(0,200);$('countAll').textContent=rows.length;apply(data);phoneDraftKey=null;msg('Утасны хүсэлтийг Sheet-д бүртгэлээ. Үнэ, барьцаа болон илгээх имэйлийг шалгаад санал бэлтгэнэ үү.','success');}catch(err){msg(err.message,'error')}finally{lock(false)}};\n$('refresh').onclick=refresh;$('search').oninput=showRows;refresh();\n</script></body></html>\n";}



function khQuoteBrandAssets_(){return {"logo": "data:image/png;base64,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33777eDBg9PT003TDAaDoVDINM1QKGRZVigU4j95OwDccssty5YtY4Gh53Caeo5+RjWaCKBS+dzCIysra/r06d98880vv/ySlZUVDAYty2I3Rnx8/AknnHDRRRddddVV7dq10+qF5nBBywyNpk64l1CtbNGq/Pz8rKysnTt3lpWVCSFSUlIaN27csGHDykSFdn1r6i1aZmg0daXqIX6/AiA8P1wLDE29RfszNJq6ovSMKr5141ngqMJqVJwTrtHUN7TM0GgigKc0YXX2VJKmwl/pRbk19RNtm9JoIoO2KWn+DVR3ZrRixYq1a9cmJiZ6tkspDcO47LLLIjgtIqK5c+dKKYUQai7GQe4FBQWtWrU68cQTa9es55Vet27dypUro6OjhRBSSsuyoqKiLrzwQrWYQaQgop9++ik7OzsqKgoAOPXXNM3Y2Nj27dtHRUVFarjhwhXfffddMBhUechMaWlp27ZtMzIy3DtzT77//vvCwsKannV4pJDaztezRmcUCoViY2PPPffcmJiYGnWguLj422+/dZ9sSUnJWWed1bRp0+ofvZps3rz5999/j42NdWsVpaWlp59++pFHHunuVd2PVVJS8v3331uW5T610tLSZs2anXbaafv9eXZ2NucVuuO4ysrKWrZs2apVq7p0jBWjZcuWbd26NTo6mjdyUV4iuuiii2JjY+vSPgD88ccfGzdujI6OVknyPBqcc845ycnJVfdt586dv/zyi+cp4nft/PPPVyWHa437/paVlX3//fehUMhzjxo1anT22WdXERmhmuIzjY2NDX9f3Gpo1U+UYRixsbHp6ekNGjRISko64BoqVYmUkj988skn4CrH5mmkb9++VbdTYcuqcQ/PPPOMp333ET/44IPKflgjCgoK1HvuPsTgwYOr6FvtyM7Obty4cfgZCSGWLFlS9/bdvc3MzIyNjXVfMfV5xowZ4T8pKSlp0qRJXR4h91OBiHUZMT/66KNqni/3X0r54IMPhp/pmWeeGdk7yDz33HMVvgIPPPBARNpXfS4oKDj11FPDLzIAXHfdddU5tWHDhlX48y5dutS9n7t27eLJh+dqIOJff/1V4RnViMpyG8eNG7ff377zzjvhP0TEtLS0goKCWnSGCT8RKeXtt99e4UB/7LHHFhUVVfgrN3l5eccee2yFvQ3fWDU+ny8mJiYjI6Nt27b333//66+/7rkRqjN1fy8qkEi0r4OOP3fq1Klbt25UidF24MCBW7ZsgcrdgOGHqGxw2bx58+DBg9WfHu/i9ddff+utt0Yk9emrr77atm1beJfefPNN0zQja2SQUqqiQ+4zklK6S0pAtS9gFajHosI2KczvyjPEupyvemD4voSfQvUbLyoqCu9z+OFUg0VFRe+99174V0uWLPnhhx/qfjE98IVVj5/qRmlp6X67XR3Us9G/f/9ly5a5v1ItV1OBKysrC/85V7WqSw8Zy7JKSkrUdVCNU9iMuHbPlXqAPTnz6jrXFH4p6vKQu19b/veHH36YPHmyZx/+yjRNfhE8b4Tn8ZBSFhcXV9jb6nTGjWmaJSUlu3fv/vnnn8ePH9+jR48zzzzz8ssv/+ijjzy/qvvgWYGm5pk2qn9HjRr11VdfbdmyhY0Y5NiLuJbO3Xff/fXXX7sV4erfIffOd911l5TSYyfhZz0xMXHChAmeh6kW8OHee+89ciqPuu/uxo0bFy9efN5550VQbCAia/GBQIC1bCIKhUJRUVEeC1LtDup+XflYpaWl7mMJITgzoML2o6Oj1aWoNSx7+BDqhQHHEOeun1EhPJYFAgHV2n5PlojmzZtXUFDgtsAAAGsh06dPP++88+pyRuGwZcNj3+D7uN9uVwe+RGvWrBkxYgQAeO4IEZmm6X5iq2iKd+Mro3Z2d1W1WYs+I2JMTEx+fn4gEFD3mu9gRN4afgzcp8+Gr/ALEm7/4RfKs6dpmlFRUUrk16JL7mkuIpaUlNx3332IaBiGR6cPhUKBQMD9rrk/eDoQExPjaaSyHla/5/wmFhcXz58/f968ecOGDevbt+/111/v3qFG5+6hUj0jnNjY2LFjxwIAl3TmUcDWVoT45ptvPv3002p2y32B3H/OmDHj22+/DR/aLMsCgP79+6elpdXR+sGXfufOnfPmzXNv4W6wKXDKlCmR1TMUyjgb8ZbDjVHuGfGBC9x0z6EMw1CzKs87UIup0353RsQ33niDHz/3sXjUmD59ekFBQQ3PpgZ4NGDY91LUuk1EZIWeh2PVoHvAgppfrgq7VOsBFFyKsqdjnh3qTmU9dHc+fB/3rAVc40ytz9ejVA0ePHjNmjUegaGO654hVTbWuX/irlYZ/lyFn6NbWHr25D8Nw/D5fH6/3+fzLV26tFOnTr169eJBoO73pQKZUcVlveqqq2655RYlKgCAx3efz0dEjz/+eE075L5AUsrHH3+cW3OPcSw2L7744kceeQTqfM58xDlz5hQXF/OEUd0wNQv45JNPPCajiBD+2ldzPK0O6oF2N4hhVpQKYZU/VHPcY5n7tVTllWrUFPfB/WZWfb7bt2+fO3eu+9zZ98692rFjx/fff1/Xy1plByBMVKuNtbin/JPx48f/8MMPfCJKFtZ9nlGhnleXZw8dPJKyinG8mrgfKtj3SXAPkVV33t0396tdl3dNNbJkyZLBgwdXIXErnEmEXyvVAt9c91lX/fqoYmW8GDBvVEdhPVsdjpW2sWPHdujQobi4uO6z4Uo97JWd9tixY5OSkpTFn/sHAEKIdevWPfXUU9U/tufCPfbYYxs3buRTVVYvAAiFQjExMa+++qrqQ01PMpzXXntNHUUJPxWplZWVNWfOnLofxYPnfCP1KDMeOQTO06Nuk+coHv2gpodzx2Z4Bkr1ztfiTiUkJFRzHo2IU6ZMAZelSD05AMAaT4Xu0IgTPjZBrZ5SRNy8efN///tf3LcaLjrhErxbdR4V9zOgtoTfstr1M7wRN27ZVrunGvf1AXimWZVNw90/92xRg0nd9Sr+ec+ePS3L8vv9nvte2ZgJ++o9an/PLfbMDKrzVrqbVRLFfZV4WGMb+Lx589q3b8/O+dpdBKYCf4ZHzsO+tyEtLe3VV1/9z3/+o4zXfP5+vz8YDI4ZM6Z79+5HH310dW6P+2nYuHHjuHHjAIBVFrcFnIgGDBjAiyTX5a4rVq5cuXTpUk9r6kb6/f6ysrJZs2Zdd911dTxQOO4R3LO9jg+0e36n9DYlFIUQHse+OmIgEPjggw+qH2tLjv2nqKjo/vvv37VrF7883DgfZcSIEaeffrpSGqoDR0O2a9eu+mc9e/ZsCDMWu4+4YMGCvXv3JiUlVb/NGqGO6BnXat1g7969i4uL/X6/W8Are2b121f7VzgvqeNLVPU9DX+h6nIstzwOP4rnWOHdC38k6tgTRJwwYcKvv/7KAliJc88bHS6hw4dTDNMd1RUzTXPcuHEnnniiJ5AhHPVqZ2dnr1+//s8//1ywYEFubi4A8PIt3DEeqwFg2bJl3bt3d4eN1IJKfeAV9o+/uv32299+++1vvvlGCQw+f5/PV1xc/Nhjj3388cduiVqdfvTu3TsUCinRrQY70zRPP/30//3vf+7W6vggckBneKQ2nwvfyJkzZ77yyisJCQm1PkqFVDEJreOr5TkEVMNUop74Cy64oBYHJaL4+HhPB/hAV1xxRevWrdWWuov5cNatW8cLVLj7497B7/fv2bNn9uzZXbp0OUA9ichgpJg3b94nn3wSPo+unShy3/TwCUpd+lnNn9ddgkK1r3AVQ1Zdjh5+lJ07dz7++OPgTG0rUys9b1xl3fNsdzd41VVXNWvWrBad3L59+6RJk4YMGcKzQPddYOV78uTJ7du37969u+eH1X87apD94Zb548ePBwDTNFnVUGLWMIwZM2bMnDlT/SS8ZxDmuvnwww8//fTTcMnMznZPrJRn0lQL2FfveTk9tz8nJ2fWrFmebkOkn8LDmuLi4lAlgZvZ2dn8oe43qzLYMLVf/f3jjz9Wnw+E6IoUhYWFd9xxByJ6Avnqc5//bRDRI488kpeXx0m4UFFIS10edff4rt6gmtK0adNnnnlmxYoVbdu2dSv63E922j/++OM7duxw/6pGT1oNZIb7OW7RosXQoUN5WHePCyx++/TpU1nEurouqoulpaXsBfG4vtnK8dRTT5122mnu8b2O79LixYtXrFjhPiP1WSlMHLIyc+ZMJeEiot/8w3DHzCg8s0v1Z2TFBofSQkUzNdh3HjBnzpxdu3ZF8NARxD0R6devn+qnsmzoh61eMX369GnTpikhQfu6Xev4hEdqkOFuHHPMMZ9++mnLli152q2aRceqPGrUKPevavSG1kzPANdz/MQTT7Ru3ZrTwZTewKUO1qxZM3r0aPevKuwQbxw2bNjff/9tGIbHJiilPOqoo/r37+/pANTtXXr//feVYUpdKY5A4D4odeerr77Ky8ur4+H+wVR4ZyszH0f2Gv72228rV650K6boBGshotsTVlJS4g4Br1eoa7J06dKRI0cahqE8Q57QF80hJxgMsoXcnUIUkVks7BvlBbUVPx5jTMOGDWfNmpWQkBBy5c2ofk6ePLm4uLg6BrRwalOZRB1p6tSpAMCljZRE4ZCPvn37/v3331V0iHu/Zs2a559/XgjBRgbc1183fvz4yJZ+ys/PdxdBAQAhhGVZSUlJ0dHRPOJwsKbP59u7d2+FOcYaxn01KvvMRHzge+uttwBAeb8AwLKsQCCQmJioFF8AYH3x448/Dhdske1P7VCC7d5774V9w2F5ENHl0OsDfFP69OmzdevW8DQACItAqd0hPC3U4hEN92C1bNnyhRdegLCZHBHt2LFj1qxZtettbWSGOtJJJ5302GOPgZPlp94Bn88XCoX4K9XL8EaIqHfv3uSK9eRpPtfvu+uuuy677LLKWqjdaz9v3rzt27e75RARSSnvuuuuVq1aceagW51yB2vqSZ8H99XwjG7hvqIIHrewsJANU27F1LKsLl263HTTTeEOqkWLFnGRmAPUn1qgDKEAMHr06BUrVoSH0nLQ5CHrosYBERcvXjxmzBhliVX/BoPBuquDHsGjcjVq3ZSbbt26NW7cmGfDbg8CACxcuFDt5jHqVk3N/BnhG4cNG3b88cfzaOt+7g3D+OSTT77//vsq+jF37tzZs2fz9Mr9/pummZyczOUTFOjywENtr6nbI8rtmKaZkJDQp0+f8DoTQojVq1evWbPG3YdaHPTfBiK6pW/EmTt3bnZ2tjvfjT906dKle/fuRMRHV69HSUnJ22+/DfVJ6qvRZ/PmzSoOx/2VaZo+n6/uRVg1dYdd33xH1PguhAgGgz6fLzo6uu7l2txio+5Vad02rri4OK4awuYTtwPmjz/+8MyiqvmC1DJuqvz3QnC2HReEUdtZdbj99tsrGz5KSkruuusucN4WZZVi69uECRNSU1PDO1BhU9Vkz549M2bM8LRGRBdddFGjRo3OPfdczyXj0GH3TzSV4b4ptG9wesRhJYOjLQCAtdKEhIT27dufddZZ6enp7hoJ3DE2okJ9kvrcvfvvvz8UCqlzYYLBYIsWLe655x7O8j10fdQAAAwZMuSXX37he6RuE9eJeOihh5KSkup+j9xio+7TGo+Z6/LLLwfneXNn+WzZsiU/Px/2Hc+r84LU7N2ucNS+5JJLuKCI+ytWNbZu3epRF9S3Q4YM2b17NwsMtygioquvvvrGG2+sUceqw4wZM0pLS1mYkQMA3HzzzUTUsWPHtLQ0j5ENANj/AdVW3DSRosILvmfPnq+++srtdeTPF154ISeL8KxKTeQBwDCMjRs3/vrrrwev69UAEd9///2vvvrK5/O5I4ZZ4L366qsXX3zxgShgo6kR69evHzx4sNsqpQaHXr16devWbc+ePRE5UKSGl/B2jj76aAgzzxBRVlZW7WJ8IjAfRMQJEyYkJiaG9i1syY7xwYMHZ2Zmen6yffv2YcOGCSHcdXWEEKFQKCoqij2cEYfjZ5Raw/mTqampV111FSImJiZWmNr266+/Ll++HOqTZeNfgudR5os/a9asnJwc93ae5amk/Y4dO8K+/hXDMILBoDvb5qDhthJ4Hp6srKzevXuDEzjPG9msx1Ws622I8L+Ku+++u6CgQMX4cHRlKBQKBAKDBw/m4heHuo/lVBin41Y73A9hUVFRYWGh+mH1jxIZG0J8fPzYsWPZDqteAM4M37t375133unZv2vXrlya213fgp1+w4YNa9CgQUR65ebvv//+8ssvwTWa8GW69NJLU1JSeMvdd98NAG7rJE8AOW884l3S7JdwrXny5MmI6Db0SykTExNvuOEG3vmaa65p0KCB2+nH2z/66KODbOfxBMN43thHH300MzPTU/udQxBfe+01AGC7gebgo27TtGnTFi1axAJDORu4nsfIkSPj4+N37txZ6zCniEOVZIp4is6By3qs3o4aqRoRszt37dr12muvdec3sDAwDOPLL7/85ptvwDmZzz777Pvvv+cEchYbqtOnnnpqr169InUb3JPNjz76yDRNtxGARZTbCNa+ffsjjjjC3QLvP2vWrLoEM2hqjWecXb58+cKFCz2l8hHxyiuvTE1N5Y1+v//KK68Ep3qm2nPNmjWLFi06VJ33vJnffvvt+++/DxUZsp988skzzzwTnIrRB7PDGoZvXHZ29n//+19wBgG3Onjuuef26NEDHB0X6tPI4FE1yFmhAPZVOHhCHxUVVYueR0ZmcJ9eeeWV5OTkYDBITl6SSrx44okn+AUOhUJ9+vQBJ8hKyRhOc586daqnMnldeuUOeOeIKY9vNj09nccXPmJKSsqFF17oHmi4e3/88cfSpUvrw1TiX4jbAvvll19ysIrH5nP11Ver/Ymoa9euHusz6yXKNXUwO69eV9VnKWXPnj3BkQpqu2maxx13XN++fT2/1RwSnnrqKY7LV2ORCgh88803eZ8DvvJ2TXA7m90BUeFB2/xnYmJiXFwcbzkEtinuxDHHHMMFRcA16JNTcZ4v9DPPPLNq1Srl+kZXWcC+ffu2bNnS/Y7VXXpzCz/++OMvv/ziTsvg7dddd52nyl6nTp34KnM3VOFeNonod/jg41afOcXSrfNxxJRbZiBi+/btjznmGHf8N+88a9YsntAczP67XwfeMmDAgFWrVhmGwXUHeCMHX4wfPz4mJoa36IftEPL1119PmDAhPGlGSvm///3vxBNPDP9JLYw8Bw635Fi+fLnHTMp/Nm7cWMWmHmzblDvK6L777rvooos8RU64vOLQoUPnz58/adIkdn27LbyWZR133HEDBgyAit6xuvfKXZSQ4R6qsUY9Ex07dkxOTnYfnTW7Tz/9tKCgoJ48EP8qVMjK4sWLV65cqbar6cg111yTlJTk1jxiY2PPP/98cHRwdNYq2LRpk3v54YN8Fvxh2bJlAwcOVMtjuLWozp07X3TRRUpG1qs57L+KsrKy+++/ny3n6vFjh+vRRx/97LPPhoc2ePTa+oCakX/++eewrxTh7S1atOAFp91DZXWITNyU+9/XXnvN7/erpT/4svp8vk2bNnXq1CkrK8sd6azWdXjrrbc8Ruq694qc4rgqPF8hpUxPT+fIZXDd7NjY2M6dO6vduJ9+v3/Lli1fffVVHbukqTVExIky7hVPefutt94KYcEhXbp0IaJgMKgEBg/BH3744aHofjmcHab8avyOhEKhhg0bjhgxglxL1+jCIYcEIhoxYsSmTZvUUupqGAGAN954Qy0LxnjMKoes32Fwf3bv3s0DlxqfpZSsgp966qngmp17jFpVEPm5zPHHH89VQzyrxPh8Pi7p7u4WZyf26NGDJ4aRhY/79ddfb9261X1HeftNN93EYta9EQB4uQWlKikOaKm7evW01R/cNgGWGe6X07Kso48++qKLLnLvzDtceOGFGRkZ7nZ4LJ41a1atq0zXpf/8eeLEibzcbPhEdfjw4UceeWRkp02ayqjidfvrr7+ef/55z0aW37feequaZTIRv0e1Hgcq/CEi9u3bt6ioyB1nyBX2AIBdubU4hYjJDHenn3/++VNPPZWHXbcb0FNwkKU3e0Ei1Y1wZsyYwYqO0jy4PzfddFOF16tdu3a8zqB7Z0ScMWMGL4Dloda3mRtnh4rbRlG71hQHblm6g4zbdLNgwYJ169apr9T29u3bq3Wx3BOUqKgotdQSOFfV7/fn5OTwqr0HQUh75p67d+/mhAz3tImVjMsvv/w///nPge6PRuF58d0PQ69evXjxN3DeULaFJCYmvv766wfosSGXm712Lagnzd3Djz/+mO034XbOc8455+yzz4awaU11OhCxgjbu4TUqKmrMmDHnn38+5zpU1g8Wd6NHj1bu+wjC519YWPjZZ5/xFqX3hEKhFi1aVLiMKC+ecd1113F9ebetgJty55qob2t3s/lXmzZtatiwYV0yg9zd8Pl8u3fv/mdoLW6tn62LKogFEfmKde3atbLfXnLJJSNHjlRrvIPzekyfPv32228/CLN4twwDgN69e+fl5bFhVkX0cRLr2LFjD3RnNB5oX7cwf3jttdfciz8qgQEAY8aM4dlY3cd3D+rldcd5VthhqFwtCO/Vd999x2ZbFWdITugXIj788MMQJiSqeV6RkRnusZU577zzHnzwwbFjx7ID3L2+t7v3HTt25DTsiHTD0xkimjZt2s6dO93BM6xmnn/++SpAxf0r3rNz586jRo1Sl4+c1c5nzpzJMsNzvrXrv9/vtyzrzjvvjImJ4QredX8QDcMwTbOkpOQfU94OEffs2TNz5kx+ftwVLVu2bNmhQwf3zu4LePXVV7ds2XLt2rXq7vPjN3v27N27d2dkZBxM48+cOXOmTp3qLnKHTnhe//79W7RocdB6ovG8vOqZycvLe/zxx4koKirKnZhsWdZFF110xx13wAEQGO423c1WuPJ0FY14JiijRo164oknLMviaZb73bEs6+STT77tttvC2zyoeob7TVAbhw0b9s0333BYoUdgcP79kUce6S42HincGs+7774L+64AyrOGBx54wN1hz6Nw7rnntmrVinvulvzz5893Dzd1f4CEEKpeDdbWh4b7umpY23DHcR7uzJ49OycnhycfygFARNdff73nynsGgs6dOw8cOFBtAQCfz2ea5rRp0x566KED3W31YOTk5PTq1YtfWncuLgCceuqp7PyL7DCkqQL3dXZf9i5duhQXF7tXyEAnuYGtUkrSRPZOoRPdh65FslVEU9WKhXtLWVnZ5s2bP//886lTpy5ZsoRnup6fhEKh2NhYXhc5vLVqnleE9Qz3KBwbGzt48ODrrrvONE31woMTKyWlfOKJJw5EmRBwbuqWLVuWLVumOqa2n3766WzL8ygK7j87dOiwevVqy7KUtDMMo7CwcNasWd26dYvUA8SmMI8SVlPIVQnA3XKte1Xf+Oijj/iDuolc2eyaa66p7Cd8U6699tpBgwZxjSBwhe3Onj37IMgM9WA8/fTTGzduVIY13sjCb+TIkWw60ALjoFGhVeqLL7748ssv3S81OUX1n3/++eOPP97z2wjeMjWY8DjAY06/fv2aNm1aVlZWYaF1pTTwBLGsrGzPnj3r16/f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UYVC1Z0wTfOiiy7yLJ5aNUrIH3/88X369Bk0aBARcVq/Cqy2LKt79+7Lli1T7yGE1QvhfGzVGv/8qKOOuuCCC2pxgozb5sb/tmjR4qSTTlqyZAm4NB5218+cOZMr1bhncPpZrA58JQsKCr755hvPFWPBf/vtt9exffWZb0qHDh3mzp3rXlCTrdtTp059/vnnad9octjfrXS3I6W87LLLunTp4p7iUEWRWocEIcSaNWs++eQTzm2sXWeio6N37drFCwqA60XgPxcuXHjMMcfUsZ/Z2dn8oS4GLs9lV+2YptmwYcPx48fXsZP/WmosM/r3779u3Tq2SqlnTgjBmVavvvpqjVpzv0V9+/adMmXKli1b3I8gAPj9/m3btvXp0+fNN9+s8Idc94nTht368o033gh1eFErVFofeOCBJUuWoKvSNWcPTJw48fHHH4+NjXX/vBYH/XeCiJ9//nlmZqYqSqisCscdd1zbtm15t9rdSjXuKEnQuXPnPn368BOrqmET0Zo1axYuXHj++eeHDzdVd56IDMMoKysLBALsnPPs4OlGTU8hUgghFi5cuHDhQqibM5xtACq8GJ0cCCJyr51caxBRJei4h/saUaG5jGchw4YNUyEVh1yQH3bUzAe+atUq9lWoJfAYtik9+eSTVQQhVIhbr4+JiRk0aBBU4qabMGHCDz/84PkhAGzfvv3777/3zOb4ybj55puhbmN3uB3s2muvjY+P5/NVCg0AbN68+ddff4V6YHw4TJk8ebLHyM5/duzYkS0VdZH9sK/+16RJk3bt2pFrHUlwbGIcCuiZO8P+bqt65AYPHnzMMceE97M+CAwAQESfz8dVFPnf2uGWox5LF48MdUQ5St0eypqeLL+bnvsopbzyyivvvPNOtY8WGDWlZjKDi86Gh0JKKZs2bTpgwICaHt5zR2+77bYrr7yS77S7jjxXtOdlUjzMmDGjoKAgvEtt2rQ5++yza9qfCnvolhwNGza86KKLPF/x0SsM9NYipDqsXr2a1151b2TjSadOnTwPSR1BJ66pwm/feecdLnPrmbXsV9WwLOukk07imqke6onAYNBxE4rawi1UdlK8oomvboBL16zLRROu1Z3JyR1220Lqwx057KhAZlQ2zL3yyis///yzezVE986vvPIKj+w1GiXD79nIkSN9Pp/KuFYuE7/f/+effz733HOe9qdM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", "qr": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAjQAAAI0CAIAAACF6SRUAAANUUlEQVR4nO3cO64kTRVGUQrVGLDwmP+Q8LCYRODg8EtcQt3BqR1Za9mtm1n56K10vtda608AUPLnT58AAPyROAGQI04A5IgTADniBECOOAGQI04A5IgTADniBECOOAGQI04A5IgTADniBECOOAGQI04A5IgTADniBECOOAGQI04A5IgTADniBECOOAGQI04A5IgTADniBECOOAGQI04A5IgTADniBECOOAGQI04A5IgTADniBECOOAGQI04A5LwnD/aXv/118nAp//z7P/7nv9m5Pjt/55RT96t2zrXzqTl1fU7di8lrWDufmsl3x5cTADniBECOOAGQI04A5IgTADniBECOOAGQI04A5IgTADniBECOOAGQM7qtt2Nyu+mU2tbW5Cbe5G8/daynbrWdOlbted5Re1b9P/b7fDkBkCNOAOSIEwA54gRAjjgBkCNOAOSIEwA54gRAjjgBkCNOAOSIEwA5uW29HTayflbbWNv5OzvnfOpe1LbsbjzWjtr51Ph/7Ge+nADIEScAcsQJgBxxAiBHnADIEScAcsQJgBxxAiBHnADIEScAcsQJgJwrt/WeanL/6tRW2+SO36nrM7kdV9s0m7zvtetcuxf8zJcTADniBECOOAGQI04A5IgTADniBECOOAGQI04A5IgTADniBECOOAGQY1vvgSY332rHOuXUVtvkPuFTtwdPXZ8d9vc6fDkBkCNOAOSIEwA54gRAjjgBkCNOAOSIEwA54gRAjjgBkCNOAOSIEwA5V27rPXX/anIfrLZFNrlBt6O2GXjjPuGOG/f3Tnnq/2On+HICIEecAMgRJwByxAmAHHECIEecAMgRJwByxAmAHHECIEecAMgRJwByctt6tf2rmlMbYjfu1H3z39lx4++a/DuT/D/2+3w5AZAjTgDkiBMAOeIEQI44AZAjTgDkiBMAOeIEQI44AZAjTgDkiBMAOa+11qfPgX87tcc1uec2uWlWO+en3q/add5x4znzM19OAOSIEwA54gRAjjgBkCNOAOSIEwA54gRAjjgBkCNOAOSIEwA54gRAzvvTJ/ArTu1o7Ti1RXZqs2tyG23HjfeitlM3eQ1rJt+dHbXneceN57zDlxMAOeIEQI44AZAjTgDkiBMAOeIEQI44AZAjTgDkiBMAOeIEQI44AZAzuq331A262vbX5GbXjb998u+cUtvxq12fG9Xer8lj7fDlBECOOAGQI04A5IgTADniBECOOAGQI04A5IgTADniBECOOAGQI04A5Ixu6006tev11P29Sac2xGrbX6fcuJ924y5c7V22c/gzX04A5IgTADniBECOOAGQI04A5IgTADniBECOOAGQI04A5IgTADniBEDOY7f1TqntuZ3a0ar9rh03bvTVjjV5DSefw9pu3qljPfV+7fDlBECOOAGQI04A5IgTADniBECOOAGQI04A5IgTADniBECOOAGQI04A5Ixu653aZZrcd7pxs+uU2tZWbR/sxufwlFPnM/m7as/PKbVn4xRfTgDkiBMAOeIEQI44AZAjTgDkiBMAOeIEQI44AZAjTgDkiBMAOeIEQM7ott6Nu1W18znl1L24cWNtx+Sm4o37cqeOtWPyWZ0851Nqz9gpvpwAyBEnAHLECYAccQIgR5wAyBEnAHLECYAccQIgR5wAyBEnAHLECYCc11pr7GC1vbIbt7921K7PjtpvP3WsU2rnfONW2+R7Ufs7N/LlBECOOAGQI04A5IgTADniBECOOAGQI04A5IgTADniBECOOAGQI04A5Ixu6+24cadukuvzs9oW2Y0bdKfceJ1r7843b/T5cgIgR5wAyBEnAHLECYAccQIgR5wAyBEnAHLECYAccQIgR5wAyBEnAHLekwc7tUlV27aaPFbt70ya3Aeb/O2153DH5Lt86pxvfHeeupu3w5cTADniBECOOAGQI04A5IgTADniBECOOAGQI04A5IgTADniBECOOAGQ81prjR2stldWU7s+tW29yV24p26aPXWf8JTae1G7F5PPsy8nAHLECYAccQIgR5wAyBEnAHLECYAccQIgR5wAyBEnAHLECYAccQIg57HbeqfUNtZqe26189nx1HPeUXsOa/di8jrX7mmNLycAcsQJgBxxAiBHnADIEScAcsQJgBxxAiBHnADIEScAcsQJgBxxAiDnPXmwG3e9djz1WDsm7+kptc23Gzcnd5z6XTc+Gzeq7Rz6cgIgR5wAyBEnAHLECYAccQIgR5wAyBEnAHLECYAccQIgR5wAyBEnAHJGt/Vq202n3Lj1N2ny+kz+9trG2uQ19IzN+Ob/N3w5AZAjTgDkiBMAOeIEQI44AZAjTgDkiBMAOeIEQI44AZAjTgDkiBMAOa+11qfP4f+itrVV22HbceMW2SmT9+vU9altrH3zu7Pjxt8++S77cgIgR5wAyBEnAHLECYAccQIgR5wAyBEnAHLECYAccQIgR5wAyBEnAHLekwer7USdOp/aDlvtOu+oXcPJY924T1jbrtxx4zNWu4aTfDkBkCNOAOSIEwA54gRAjjgBkCNOAOSIEwA54gRAjjgBkCNOAOSIEwA5o9t6kztRN+5o7Tj1u2pbWzeez+S9uHEvccfkbl7tPf3m+77DlxMAOeIEQI44AZAjTgDkiBMAOeIEQI44AZAjTgDkiBMAOeIEQI44AZDzWmt9+hz+w1M3qSY3xHbUtuxu3Km78Ro6Z/6b2r3w5QRAjjgBkCNOAOSIEwA54gRAjjgBkCNOAOSIEwA54gRAjjgBkCNOAOS8Jw82ud1U21ir/fYdk/uEtf001/D3Tf72yffrxt3O2jnv8OUEQI44AZAjTgDkiBMAOeIEQI44AZAjTgDkiBMAOeIEQI44AZAjTgDkvNZaYwer7TvduLV1435abXuwds615+epm5O1+1V7l2t8OQGQI04A5IgTADniBECOOAGQI04A5IgTADniBECOOAGQI04A5IgTADnvT5/Ar6ht4k2eT22HrXas2vV56o7fjbt5k3+ndi92TN6vHb6cAMgRJwByxAmAHHECIEecAMgRJwByxAmAHHECIEecAMgRJwByxAmAnNy2Xm1ra3JD7KkbazVPPefJXcEdrvPMsZ76LvtyAiBHnADIEScAcsQJgBxxAiBHnADIEScAcsQJgBxxAiBHnADIEScAcnLbejtO7ctNHmty96y2sXbKjb+rtldW26485anv++TzM/n/6g5fTgDkiBMAOeIEQI44AZAjTgDkiBMAOeIEQI44AZAjTgDkiBMAOeIEQM5rrTV2sG/exKttmp367bXtuB1PfQ5vvKe1ezGp9ozV+HICIEecAMgRJwByxAmAHHECIEecAMgRJwByxAmAHHECIEecAMgRJwByrtzWq+3dfbPa1t+kG7fRaver9g5OPmO1TcUdk/fLlxMAOeIEQI44AZAjTgDkiBMAOeIEQI44AZAjTgDkiBMAOeIEQI44AZDz/vQJ/IraptmNe2W1nbodT70Xnuff/zs7as/85P268f8EX04A5IgTADniBECOOAGQI04A5IgTADniBECOOAGQI04A5IgTADniBEDO6LbeqY2sU06dT237q7aNVlN7DndMbtDduPVX++07Jt+dG99TX04A5IgTADniBECOOAGQI04A5IgTADniBECOOAGQI04A5IgTADniBEDOa601drAb951OqW121c5nUu23187nlNrm5I7Jrb/JDcwb+XICIEecAMgRJwByxAmAHHECIEecAMgRJwByxAmAHHECIEecAMgRJwBy3p8+gT+6cSdqcvtrcquttv01uWk2uY22Y/Je1PYAn/o88zNfTgDkiBMAOeIEQI44AZAjTgDkiBMAOeIEQI44AZAjTgDkiBMAOeIEQE5uW2/H5KbZ5I5Wbautton3zfdiR21XsPb8nDpWzVM3A305AZAjTgDkiBMAOeIEQI44AZAjTgDkiBMAOeIEQI44AZAjTgDkiBMAOVdu632z2v5VzY27i7U9t9r+Xu2Zr20P7tg5n9o5+3ICIEecAMgRJwByxAmAHHECIEecAMgRJwByxAmAHHECIEecAMgRJwByXmutsYNNbnadMrmjVdtYq21/3bhlV7uGN3INf3bj1t8OX04A5IgTADniBECOOAGQI04A5IgTADniBECOOAGQI04A5IgTADniBEDO+9Mn8CtqG1Cn3Li/d+Nm146n7vidOp8b7+nkBuZT9+4m+XICIEecAMgRJwByxAmAHHECIEecAMgRJwByxAmAHHECIEecAMgRJwByXmutsYPVduEmTW5kTW5/Tartnu2onXNtq+2p24OTavf0FF9OAOSIEwA54gRAjjgBkCNOAOSIEwA54gRAjjgBkCNOAOSIEwA54gRAzui2HgDs8OUEQI44AZAjTgDkiBMAOeIEQI44AZAjTgDkiBMAOeIEQI44AZAjTgDkiBMAOeIEQI44AZAjTgDkiBMAOeIEQI44AZAjTgDkiBMAOeIEQI44AZAjTgDkiBMAOeIEQI44AZAjTgDkiBMAOeIEQI44AZAjTgDkiBMAOeIEQI44AZAjTgDk/Asayt2lutrzdAAAAABJRU5ErkJggg=="};}


/** Authenticated phone lead registration. Calls the existing lead receiver validation. */
function salesDashboardCreatePhone(payload){
  var actor=khAuthorize_();
  if(!payload||payload.consent!==true)throw new Error('Харилцагчийн мэдээлэл бүртгэж, эргэн холбогдох зөвшөөрлийг баталгаажуулна уу.');
  var key=String(payload.draftKey||'');
  if(!/^[a-f0-9]{32}$/i.test(key))throw new Error('Бүртгэлийн түлхүүр буруу. Самбарыг дахин нээнэ үү.');
  var product=String(payload.product||'');
  var thickness=product==='600 × 240 × 300 мм'?300:product==='600 × 240 × 200 мм'?200:0;
  if(!thickness)throw new Error('Блокийн хэмжээг сонгоно уу.');
  var volume=numberBetween_(payload.volume,0.1,100000);
  var q={requestId:'KH-TEL-'+key.toUpperCase(),name:payload.name,phone:payload.phone,
    location:payload.location,email:payload.email,notes:payload.notes,product:product,
    volume:volume,pieces:Math.ceil(volume/(0.6*0.24*thickness/1000)-1e-9),
    source:'Утасны хүсэлт',consent:true,website:'',utm:'',breakdown:''};
  var result=saveLead_(q,actor);
  var view=salesDashboardLoad(result.requestId);
  view.createdBy=actor;view.source='Утасны хүсэлт';view.alreadySaved=result.alreadySaved;
  return view;
}
function khLeadOwnerColumn_(sh){
  var header=String(sh.getRange(1,18).getValue()||'');
  if(header&&header!=='Бүртгэсэн ажилтан')throw new Error('R баганад өөр мэдээлэл байна. Өөрчлөлт хийгээгүй.');
  if(!header){
    // Never overwrite a populated extra column, even when its header is blank.
    if(sh.getLastRow()>1&&sh.getRange(2,18,sh.getLastRow()-1,1).getValues().some(function(r){return r[0]!=='';}))
      throw new Error('R баганад мэдээлэл байна. Бүртгэсэн ажилтны багана үүсгээгүй.');
    sh.getRange(1,18).setValue('Бүртгэсэн ажилтан').setBackground('#0C4D38').setFontColor('#FFFFFF').setFontWeight('bold');
    sh.setColumnWidth(18,240);
  }
}
