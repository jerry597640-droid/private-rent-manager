const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict'),crypto=require('node:crypto');
const code=fs.readFileSync(__dirname+'/apps-script/Code.gs','utf8');
let count=0;
function fixture(loan=false){
 const h=['物件名稱','類型','每月應收日','應收金額','本月應收','實際入帳日','實收金額',...(loan?['貸款金額']:[]),'差額','是否已收','備註'];
 const values=[h,['測試物件','租金',1,100,100,'',0,...(loan?[20]:[]),-100,false,'']];
 const formulas=[h.map(()=>''),h.map(x=>x==='本月應收'?'=D2':x==='差額'?'=G2-D2':'')],writes=[];
 const sh={getName:()=> '2026-10',getLastColumn:()=>h.length,getLastRow:()=>2};
 sh.getRange=(r,c,n=1,w=1)=>({
   getValues:()=>values.slice(r-1,r-1+n).map(row=>row.slice(c-1,c-1+w)),
   getFormulas:()=>formulas.slice(r-1,r-1+n).map(row=>row.slice(c-1,c-1+w)),
   getDisplayValues:()=>values.slice(r-1,r-1+n).map(row=>row.slice(c-1,c-1+w).map(String)),
   canEdit:()=>true,
   setValue:v=>{writes.push({c,v});values[r-1][c-1]=v}
 });
 sh.getDataRange=()=>sh.getRange(1,1,2,h.length);
 const ss={getSheets:()=>[sh],getSheetByName:()=>sh,getSpreadsheetTimeZone:()=> 'Asia/Taipei'};
 const ctx={Date,PropertiesService:{getScriptProperties:()=>({getProperty:k=>k==='SPREADSHEET_ID'?'test-copy':'TEST'})},
 SpreadsheetApp:{openById:()=>ss,flush:()=>{}},LockService:{getScriptLock:()=>({waitLock:()=>{},releaseLock:()=>{}})},
 Utilities:{DigestAlgorithm:{SHA_256:'sha256'},computeDigest:(_,s)=>crypto.createHash('sha256').update(s).digest(),base64EncodeWebSafe:b=>Buffer.from(b).toString('base64url'),
 parseDate:s=>new Date(s.slice(0,10)+'T12:00:00+08:00'),formatDate:d=>new Date(d.getTime()+8*3600*1000).toISOString().slice(0,10)}};
 vm.createContext(ctx);vm.runInContext(code,ctx);
 const row=ctx.getRows('2026-10')[0];
 return {ctx,values,formulas,h,writes,p:{...row,month:'2026-10',date:'2026-10-26',amount:'123.45',paid:true,note:'驗收'}};
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
const html=fs.readFileSync(__dirname+'/apps-script/Index.html','utf8');
new vm.Script(html.match(/<script>([\s\S]*?)<\/script>/)[1]);
console.log(count+' backend checks passed; frontend syntax passed');