const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
const ctx={console,Number,Date,Math,JSON,isFinite,LockService:{getScriptLock:()=>({tryLock:()=>true,releaseLock(){}})}};
vm.createContext(ctx);vm.runInContext(fs.readFileSync(__dirname+'/../Code.gs','utf8'),ctx);
let actor='sales@example.com',rows=[Array(18).fill('')],mailSent=0;
ctx.khAuthorize_=()=>{if(actor!=='sales@example.com')throw Error('No sales permission');return actor};
function range(row,col,num=1,width=1){const obj={getValue:()=>rows[row-1]?.[col-1]??'',getValues:()=>Array.from({length:num},(_,i)=>Array.from({length:width},(_,j)=>rows[row-1+i]?.[col-1+j]??'')),setValue(v){rows[row-1]??=[];rows[row-1][col-1]=v;return obj},setBackground(){return obj},setFontColor(){return obj},setFontWeight(){return obj},createTextFinder(value){return {matchEntireCell(){return this},findNext(){let index=rows.findIndex((r,i)=>i>=row-1&&i<row-1+num&&r[col-1]===value);return index<0?null:{getRow:()=>index+1}}}}};return obj}
const sheet={getLastRow:()=>rows.length,getRange:range,setColumnWidth(){},appendRow(r){rows.push([...r])}};
ctx.getTargetSheet_=()=>sheet;
ctx.salesDashboardLoad=id=>{let r=rows.find(r=>r[1]===id);return {values:[r[1],r[3],r[4],r[12],r[5],r[6],r[8],r[9]],state:'DRAFT'}};
ctx.MailApp={sendEmail(){mailSent++}};
const input={draftKey:'11223344556677889900aabbccddeeff',name:'Утасны тест',phone:'+976 99112233',email:'',location:'Говь-Алтай',volume:10,product:'600 × 240 × 300 мм',notes:'Утсаар асуусан',consent:true,createdBy:'attacker@example.com',unitPrice:1,source:'Fake'};
let result=ctx.salesDashboardCreatePhone(input);assert.equal(rows.length,2);assert.equal(rows[1][1],'KH-TEL-11223344556677889900AABBCCDDEEFF');assert.equal(rows[1][14],'Утасны хүсэлт');assert.equal(rows[1][17],'sales@example.com');assert.equal(rows[0][17],'Бүртгэсэн ажилтан');assert.equal(rows[1][8],10);assert.equal(rows[1][9],232);assert.equal(rows[1][10],300000);assert.equal(rows[1][11],3000000);assert.equal(result.createdBy,'sales@example.com');assert.equal(result.alreadySaved,false);
assert.equal(ctx.salesDashboardCreatePhone(input).alreadySaved,true);assert.equal(rows.length,2);
for(const change of [{consent:false},{draftKey:'bad'},{product:'Other'},{phone:'123'},{volume:0},{email:'invalid'}])assert.throws(()=>ctx.salesDashboardCreatePhone({...input,...change}));assert.equal(rows.length,2);
actor='outsider@example.com';assert.throws(()=>ctx.salesDashboardCreatePhone(input),/permission/);actor='sales@example.com';
let next={...input,draftKey:'ffeeddccbbaa00998877665544332211'};rows[0][17]='Existing data';assert.throws(()=>ctx.salesDashboardCreatePhone(next),/R баганад/);assert.equal(rows.length,2);rows[0][17]='';rows[1][17]='Nonempty legacy data';assert.throws(()=>ctx.salesDashboardCreatePhone(next),/R баганад/);assert.equal(rows.length,2);rows[0][17]='Бүртгэсэн ажилтан';rows[1][17]='another@example.com';assert.throws(()=>ctx.salesDashboardCreatePhone(input),/өөр ажилтан/);
ctx.saveLead_({requestId:'KH-WEB-TEST-1234567890',name:'Website test',phone:'99112233',location:'UB',volume:1,pieces:24,product:'600 × 240 × 300 мм',consent:true,email:'',createdBy:'spoof@example.com'});assert.equal(rows.length,3);assert.equal(rows[2].length,17);assert.equal(rows[2][17],undefined);assert.equal(mailSent,0);
console.log('PASS: phone registration, server identity/source/price, optional email, internal packing, retry idempotency, authorization, validation, safe audit column, public form compatibility. No real mail or Google writes.');
