const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
function output(html){return {html,title:'',setTitle(s){this.title=s;return this},addMetaTag(k,v){this.meta=[k,v];return this}}}
const c={HtmlService:{createHtmlOutput:output}};vm.createContext(c);vm.runInContext(fs.readFileSync(__dirname+'/../Code.gs','utf8'),c);
let actor='sales@example.com';c.khAuthorize_=()=>{if(!actor)throw Error('Denied');return actor};
let page=c.doGet({parameter:{sales:'1'}});assert.match(page.html,/Борлуулалтын самбар/);assert.match(page.html,/sales@example.com/);assert.match(page.html,/backToRequests/);assert.equal(page.meta[0],'viewport');
actor='';let denied=c.doGet({parameter:{sales:'1'}});assert.match(denied.html,/эрх шаардлагатай/);assert.doesNotMatch(denied.html,/salesDashboardData|discountPercent|phoneForm/);
assert.match(c.doGet().html,/Сервис ажиллаж байна/);assert.match(c.doGet({parameter:{sales:'other'}}).html,/Сервис ажиллаж байна/);
console.log('PASS authenticated mobile route, closed unauthorized page, viewport and unchanged public receiver. No Google writes or email.');
