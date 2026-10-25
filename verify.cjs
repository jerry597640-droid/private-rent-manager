const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict'),crypto=require('node:crypto');
const code=fs.readFileSync(__dirname+'/apps-script/Code.gs','utf8');
let count=0;
function fixture(loan=false){
 const h=['物件名稱','類型','每月應收日','應收金額','本月應收','實際入帳日','實收金額',...(loan?['貸款金額']:[]),'差額','是否已收','備註'];
 const values=[h,['測試物件','租金',1,100,100,'',0,...(loan?[20]:[]),-100,false,'']];
 const formulas=[h.map(()=>''),h.map(x=>x==='本月應收'?'=D2':x==='差額'?'=G2-D2':'')],writes=[];
 const sh={getName:()=> '2026-10',getLastColumn:()=>h.length,getLastRow:()=>values.length};
 sh.getRange=(r,c,n=1,w=1)=>({
   getValues:()=>values.slice(r-1,r-1+n).map(row=>row.slice(c-1,c-1+w)),
   getFormulas:()=>formulas.slice(r-1,r-1+n).map(row=>row.slice(c-1,c-1+w)),
   getDisplayValues:()=>values.slice(r-1,r-1+n).map(row=>row.slice(c-1,c-1+w).map(String)),
   canEdit:()=>true,
   setValue:v=>{writes.push({c,v});values[r-1][c-1]=v}
 });
 sh.getDataRange=()=>sh.getRange(1,1,values.length,h.length);
 const ss={getSheets:()=>[sh],getSheetByName:()=>sh,getSpreadsheetTimeZone:()=> 'Asia/Taipei'};
 const ctx={Date,PropertiesService:{getScriptProperties:()=>({getProperty:k=>k==='SPREADSHEET_ID'?'test-copy':'TEST'})},
 SpreadsheetApp:{openById:()=>ss,flush:()=>{}},LockService:{getScriptLock:()=>({waitLock:()=>{},releaseLock:()=>{}})},
 Utilities:{DigestAlgorithm:{SHA_256:'sha256'},computeDigest:(_,s)=>crypto.createHash('sha256').update(s).digest(),base64EncodeWebSafe:b=>Buffer.from(b).toString('base64url'),
 parseDate:s=>new Date(s.slice(0,10)+'T12:00:00+08:00'),formatDate:d=>new Date(d.getTime()+8*3600*1000).toISOString().slice(0,10)}};
 vm.createContext(ctx);vm.runInContext(code,ctx);
 const row=ctx.getRows('2026-10')[0];
 return {ctx,ss,sh,values,formulas,h,writes,p:{...row,month:'2026-10',date:'2026-10-26',amount:'123.45',paid:true,note:'驗收'}};
}
function check(name,fn){fn();count++;console.log('PASS '+name)}
for(const loan of [false,true])check('write only correct input columns; loan='+loan,()=>{
 const f=fixture(loan),before=f.values[1].slice();f.ctx.saveRent(f.p);
 assert.deepEqual(f.writes.map(x=>f.h[x.c-1]),['實際入帳日','實收金額','是否已收','備註']);
 f.h.forEach((h,i)=>{if(!['實際入帳日','實收金額','是否已收','備註'].includes(h))assert.equal(f.values[1][i],before[i])});
});
check('zero amount remains zero',()=>assert.equal(fixture().ctx.getRows('2026-10')[0].amount,0));
for(const [key,value] of [['amount','-1'],['amount','Infinity'],['amount','1e2'],['amount',true],['amount','1.234'],['date','2026-02-30'],['date','bad'],['paid','true'],['amount',''],['row',1],['month','2026-13'],['property','wrong'],['note','x'.repeat(501)]]){
 check('invalid '+key+' rejected without writes',()=>{const f=fixture();assert.throws(()=>f.ctx.saveRent({...f.p,[key]:value}));assert.equal(f.writes.length,0)});
}
check('changed record rejects stale save',()=>{const f=fixture();f.values[1][6]=77;assert.throws(()=>f.ctx.saveRent(f.p),/已被更新/);assert.equal(f.writes.length,0)});
check('formula notes stay text',()=>{const f=fixture();f.ctx.saveRent({...f.p,note:'=IMPORTXML("https://invalid.example","//x")'});assert.equal(f.writes[3].v[0],"'")});
check('editable-column formula is preserved',()=>{const f=fixture();f.formulas[1][6]='=100';f.p.version=f.ctx.getRows('2026-10')[0].version;assert.throws(()=>f.ctx.saveRent(f.p),/公式/);assert.equal(f.writes.length,0)});
check('duplicate headers fail closed',()=>{const f=fixture();f.values[0][2]='實收金額';assert.throws(()=>f.ctx.getRows('2026-10'),/重複/)});
check('missing headers fail closed',()=>{const f=fixture();f.values[0][6]='wrong';assert.throws(()=>f.ctx.saveRent(f.p),/缺少/);assert.equal(f.writes.length,0)});
check('blank unpaid record allowed',()=>{const f=fixture();f.ctx.saveRent({...f.p,date:'',amount:'',paid:false,note:''});assert.equal(f.writes[1].v,'')});
check('test banner is default',()=>assert.equal(fixture().ctx.getContext().testMode,true));
const dash=f=>f.ctx.dashboard_(f.ss,'2026-10','2026-10-26');
function field(f,name,value,row=1){f.values[row][f.h.indexOf(name)]=value}
check('dashboard totals and overdue are read only',()=>{const f=fixture(true),d=dash(f);assert.equal(d.summary.expected,100);assert.equal(d.summary.remaining,100);assert.equal(d.summary.overdueAmount,100);assert.equal(d.records[0].overdueDays,25);assert.equal(d.summary.afterLoan,-20);assert.equal(f.writes.length,0)});
check('monthly expected overrides base amount',()=>{const f=fixture();field(f,'應收金額',999);assert.equal(dash(f).summary.expected,100)});
check('blank monthly expected is unknown, never fallback',()=>{const f=fixture();field(f,'本月應收','');const d=dash(f);assert.equal(d.summary.expectedMissing,1);assert.equal(d.summary.rate,null);assert.equal(d.records[0].remaining,null)});
check('absent monthly column uses base amount',()=>{const f=fixture();f.h[4]='other';assert.equal(dash(f).summary.expected,100)});
check('loan absent remains unknown',()=>{const d=dash(fixture());assert.equal(d.summary.hasLoan,false);assert.equal(d.summary.afterLoan,null)});
check('invalid loan is not treated as zero',()=>{const f=fixture(true);field(f,'貸款金額','格式待修');assert.equal(dash(f).summary.afterLoan,null)});
check('overpayment never masks another row shortfall',()=>{const f=fixture();field(f,'實收金額',200);f.values.push(f.values[1].slice());f.formulas.push(f.formulas[1].slice());field(f,'實收金額',0,2);const d=dash(f);assert.equal(d.summary.actual,200);assert.equal(d.summary.expected,200);assert.equal(d.summary.remaining,100);assert.equal(d.summary.rate,50)});
check('paid checkbox does not hide unpaid balance',()=>{const f=fixture();field(f,'是否已收',true);const d=dash(f);assert.equal(d.summary.overdueCount,1);assert.equal(d.summary.reviewCount,1)});
check('fully funded unchecked payment needs review',()=>{const f=fixture();field(f,'實收金額',100);const d=dash(f);assert.equal(d.summary.overdueCount,0);assert.equal(d.summary.reviewCount,1);assert.equal(d.summary.rate,100)});
check('today is due soon and not overdue',()=>{const f=fixture();field(f,'每月應收日',26);const d=dash(f);assert.equal(d.summary.overdueCount,0);assert.equal(d.summary.dueSoonCount,1)});
check('missing due day is not invented',()=>{const f=fixture();field(f,'每月應收日','');const d=dash(f);assert.equal(d.summary.overdueCount,0);assert.equal(d.records[0].due,'');assert.equal(d.summary.reviewCount,1)});
check('month-end due date clamps including leap years',()=>{const f=fixture();assert.equal(f.ctx.dueDate_('2026-02',31),'2026-02-28');assert.equal(f.ctx.dueDate_('2024-02',31),'2024-02-29');assert.equal(f.ctx.dueDate_('2026-10',32),'')});
check('invalid actual hides rate and flags incomplete aggregate',()=>{const f=fixture();field(f,'實收金額','#REF!');const d=dash(f);assert.equal(d.summary.actualMissing,1);assert.equal(d.summary.rate,null);assert.equal(d.records[0].remaining,null)});
check('zero expected is valid and not overdue',()=>{const f=fixture();field(f,'本月應收',0);const d=dash(f);assert.equal(d.summary.expectedMissing,0);assert.equal(d.summary.remaining,0);assert.equal(d.summary.overdueCount,0);assert.equal(d.summary.rate,null)});
check('calendar gaps stay unavailable and are not previous-month comparison',()=>{const d=dash(fixture());assert.equal(d.trend.length,12);assert.equal(d.trend[0].month,'2025-11');assert.equal(d.trend[10].month,'2026-09');assert.equal(d.trend[10].available,false);assert.equal(d.comparison,null)});
check('trend compares the actual previous calendar month',()=>{const f=fixture(),p=fixture();p.sh.getName=()=> '2026-09';field(p,'實收金額',50);field(f,'實收金額',75);f.ss.getSheets=()=>[p.sh,f.sh];f.ss.getSheetByName=m=>m==='2026-09'?p.sh:f.sh;const d=dash(f);assert.equal(d.comparison.delta,25);assert.equal(d.comparison.percent,50)});
check('empty month stays valid with zero records',()=>{const f=fixture();f.values.splice(1);f.formulas.splice(1);const d=dash(f);assert.equal(d.summary.count,0);assert.equal(d.summary.rate,null);assert.equal(d.properties.length,0)});
check('dashboard rejects invalid month and duplicate optional headers',()=>{const f=fixture();assert.throws(()=>f.ctx.getDashboard('2026-13'));f.h[3]='本月應收';assert.throws(()=>dash(f),/重複/);assert.equal(f.writes.length,0)});
check('dashboard omits notes and row-version hashes',()=>{const f=fixture();field(f,'備註','private note');const d=dash(f);assert.equal(JSON.stringify(d).includes('private note'),false);assert.equal('version' in d.records[0],false)});
const html=fs.readFileSync(__dirname+'/apps-script/Index.html','utf8');
new vm.Script(html.match(/<script>([\s\S]*?)<\/script>/)[1]);
console.log(count+' backend checks passed; frontend syntax passed');
