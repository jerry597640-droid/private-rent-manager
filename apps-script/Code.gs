// Deploy as yourself with access set to "Only myself".
// Keep SPREADSHEET_ID and APP_MODE (TEST/LIVE) in Script Properties only.
function book_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw Error('請先設定 SPREADSHEET_ID');
  return SpreadsheetApp.openById(id);
}
function cols_(sh) {
  const h = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(s=>s.trim()), c={};
  ['物件名稱','類型','實際入帳日','實收金額','是否已收','備註'].forEach(name=>{
    if(h.filter(x=>x===name).length!==1) throw Error(sh.getName()+' 欄位缺少或重複：'+name);
    c[name]=h.indexOf(name)+1;
  });
  return c;
}
function months_(ss) { return ss.getSheets().map(s=>s.getName()).filter(n=>/^\d{4}-(0[1-9]|1[0-2])$/.test(n)).sort(); }
function sheet_(ss,month) {
  if(typeof month!=='string'||!months_(ss).includes(month)) throw Error('月份不存在');
  return ss.getSheetByName(month);
}
function version_(values,formulas) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify([values,formulas])));
}
function getMonths() { return months_(book_()); }
function getContext() {
  const ss=book_();
  return {months:months_(ss),today:Utilities.formatDate(new Date(),ss.getSpreadsheetTimeZone(),'yyyy-MM-dd'),testMode:PropertiesService.getScriptProperties().getProperty('APP_MODE')!=='LIVE'};
}
function getRows(month) {
  const ss=book_(),sh=sheet_(ss,month),c=cols_(sh),range=sh.getDataRange(),v=range.getValues(),f=range.getFormulas(),tz=ss.getSpreadsheetTimeZone();
  return v.slice(1).map((r,i)=>{
    const dt=r[c['實際入帳日']-1];
    return {row:i+2,property:String(r[c['物件名稱']-1]||''),type:String(r[c['類型']-1]||''),
      date:dt instanceof Date?Utilities.formatDate(dt,tz,'yyyy-MM-dd'):String(dt||''),
      amount:r[c['實收金額']-1],paid:r[c['是否已收']-1]===true,note:String(r[c['備註']-1]||''),version:version_(r,f[i+1])};
  }).filter(r=>r.property.trim());
}
function validate_(p,tz) {
  if(!p||typeof p!=='object') throw Error('資料格式錯誤');
  if(typeof p.amount!=='string'&&typeof p.amount!=='number') throw Error('金額錯誤');
  const text=String(p.amount).trim();
  if(text&&!/^\d+(\.\d{1,2})?$/.test(text)) throw Error('請填非負金額，最多兩位小數');
  const amount=text===''?'':Number(text);
  if(amount!==''&&(!Number.isFinite(amount)||amount>999999999)) throw Error('金額超出範圍');
  if(typeof p.paid!=='boolean') throw Error('收款狀態錯誤');
  if(typeof p.date!=='string'||(p.date&&!/^\d{4}-\d{2}-\d{2}$/.test(p.date))) throw Error('日期錯誤');
  let date='';
  if(p.date) {
    date=Utilities.parseDate(p.date+' 12:00:00',tz,'yyyy-MM-dd HH:mm:ss');
    if(Utilities.formatDate(date,tz,'yyyy-MM-dd')!==p.date) throw Error('日期錯誤');
  }
  if(p.paid&&(amount===''||amount<=0||!date)) throw Error('已收款必須填寫金額及入帳日');
  if(typeof p.note!=='string'||p.note.length>500) throw Error('備註最多 500 字');
  const note=/^[\s]*[=+@-]/.test(p.note)||/^'/.test(p.note)?"'"+p.note:p.note;
  return {date:date,amount:amount,paid:p.paid,note:note};
}
function saveRent(p) {
  const ss=book_(),data=validate_(p,ss.getSpreadsheetTimeZone()),lock=LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh=sheet_(ss,p.month),c=cols_(sh),row=Number(p.row);
    if(!Number.isInteger(row)||row<2||row>sh.getLastRow()) throw Error('資料列錯誤');
    const range=sh.getRange(row,1,1,sh.getLastColumn()),values=range.getValues()[0],formulas=range.getFormulas()[0];
    if(String(values[c['物件名稱']-1])!==p.property||String(values[c['類型']-1])!==p.type) throw Error('物件已變動，請重新讀取');
    if(typeof p.version!=='string'||p.version!==version_(values,formulas)) throw Error('資料已被更新，請重新讀取後再儲存');
    const names=['實際入帳日','實收金額','是否已收','備註'],cells=names.map(name=>sh.getRange(row,c[name]));
    if(names.some(name=>formulas[c[name]-1])||cells.some(cell=>!cell.canEdit())) throw Error('收款欄位包含公式或禁止修改，已停止寫入');
    // Only four input cells are writable. Preserve loan/difference/formula columns.
    [data.date,data.amount,data.paid,data.note].forEach((value,i)=>cells[i].setValue(value));
    SpreadsheetApp.flush();
    return {success:true};
  } finally {lock.releaseLock();}
}
function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('棲島日常｜每月收租管理').addMetaTag('viewport','width=device-width, initial-scale=1');
}

// Dashboard reads only. No writes, cross-user cache, or client-supplied file IDs.
function number_(value,blank) {
  if(value===''||value===null||value===undefined) return blank;
  if(typeof value!=='number'&&typeof value!=='string') return null;
  const text=String(value).trim();
  if(!/^\d+(\.\d+)?$/.test(text)) return null;
  const n=Number(text);return Number.isFinite(n)?Math.round(n*100)/100:null;
}
function round_(n) { return Math.round((n+Number.EPSILON)*100)/100; }
function monthOffset_(month,offset) {
  const parts=month.split('-').map(Number),date=new Date(Date.UTC(parts[0],parts[1]-1+offset,1));
  return date.getUTCFullYear()+'-'+String(date.getUTCMonth()+1).padStart(2,'0');
}
function dueDate_(month,value) {
  if(typeof value!=='number'&&typeof value!=='string') return '';
  const text=String(value).trim();if(!/^\d{1,2}$/.test(text)) return '';
  const day=Number(text);if(day<1||day>31) return '';
  const p=month.split('-').map(Number),last=new Date(Date.UTC(p[0],p[1],0)).getUTCDate();
  return month+'-'+String(Math.min(day,last)).padStart(2,'0');
}
function dashboardRows_(ss,month,today) {
  const sh=sheet_(ss,month),required=cols_(sh),v=sh.getDataRange().getValues(),h=v[0].map(x=>String(x).trim());
  function index(name){const found=h.map((x,i)=>x===name?i:-1).filter(i=>i>=0);if(found.length>1)throw Error(month+' 欄位重複：'+name);return found.length?found[0]:-1}
  const expectedColumn=index('本月應收'),fallback=index('應收金額'),dueColumn=index('每月應收日'),loanColumn=index('貸款金額');
  const source=expectedColumn>=0?expectedColumn:fallback,todayMs=Date.parse(today+'T00:00:00Z');
  return v.slice(1).map((r,i)=>{
    const property=String(r[required['物件名稱']-1]||'').trim();if(!property)return null;
    const actual=number_(r[required['實收金額']-1],0),expected=source>=0?number_(r[source],null):null,loan=loanColumn>=0?number_(r[loanColumn],0):null;
    const paid=r[required['是否已收']-1]===true,date=r[required['實際入帳日']-1],due=dueColumn>=0?dueDate_(month,r[dueColumn]):'';
    const remaining=expected===null||actual===null?null:round_(Math.max(expected-actual,0));
    const days=due?Math.floor((todayMs-Date.parse(due+'T00:00:00Z'))/86400000):null,issues=[];
    if(expected===null)issues.push('應收金額待確認');if(actual===null)issues.push('實收金額格式待確認');
    if(remaining>0&&!due)issues.push('應收日待確認');
    if(paid&&(actual===null||actual<=0||!date||remaining>0))issues.push('已勾選但金額或日期待核對');
    if(!paid&&expected>0&&actual!==null&&actual>=expected)issues.push('金額已到位，尚未勾選');
    if(actual!==null&&expected!==null&&actual>expected)issues.push('實收超過應收，請核對');
    if(loanColumn>=0&&loan===null)issues.push('貸款金額格式待確認');
    return {row:i+2,property:property,type:String(r[required['類型']-1]||''),actual:actual,expected:expected,remaining:remaining,loan:loan,hasLoan:loanColumn>=0,paid:paid,due:due,
      overdueDays:remaining>0&&days>0?days:0,dueSoon:remaining>0&&days!==null&&days<=0&&days>=-7,issues:issues};
  }).filter(Boolean);
}
function summarize_(records) {
  const s={count:records.length,expected:0,actual:0,remaining:0,allocated:0,expectedMissing:0,actualMissing:0,remainingMissing:0,paidCount:0,overdueCount:0,overdueAmount:0,dueSoonCount:0,reviewCount:0,loan:0,loanMissing:0,hasLoan:records.length>0&&records.every(r=>r.hasLoan)};
  records.forEach(r=>{
    if(r.expected===null)s.expectedMissing++;else s.expected+=r.expected;
    if(r.actual===null)s.actualMissing++;else s.actual+=r.actual;
    if(r.remaining===null)s.remainingMissing++;else s.remaining+=r.remaining;
    if(r.expected!==null&&r.actual!==null)s.allocated+=Math.min(r.expected,r.actual);
    if(r.paid)s.paidCount++;if(r.overdueDays){s.overdueCount++;s.overdueAmount+=r.remaining}
    if(r.dueSoon)s.dueSoonCount++;if(r.issues.length)s.reviewCount++;
    if(r.loan===null)s.loanMissing++;else s.loan+=r.loan;
  });
  ['expected','actual','remaining','allocated','overdueAmount','loan'].forEach(k=>s[k]=round_(s[k]));
  s.rate=s.expected>0&&!s.expectedMissing&&!s.actualMissing?Math.round(s.allocated/s.expected*1000)/10:null;
  s.afterLoan=s.hasLoan&&!s.loanMissing&&!s.actualMissing?round_(s.actual-s.loan):null;
  return s;
}
function dashboard_(ss,month,today) {
  sheet_(ss,month);
  const available=months_(ss),records=dashboardRows_(ss,month,today),summary=summarize_(records),warnings=[];
  const trend=Array.from({length:12},(_,i)=>monthOffset_(month,i-11)).map(m=>{
    if(!available.includes(m))return {month:m,available:false};
    try{return Object.assign({month:m,available:true},m===month?summary:summarize_(dashboardRows_(ss,m,today)))}
    catch(e){warnings.push(m+' 無法彙整，請檢查欄位');return {month:m,available:false,error:true}}
  });
  const grouped=new Map();records.forEach(r=>{if(!grouped.has(r.property))grouped.set(r.property,[]);grouped.get(r.property).push(r)});
  const properties=Array.from(grouped,([name,list])=>Object.assign({name:name,firstRow:list[0].row},summarize_(list))).sort((a,b)=>b.remaining-a.remaining||a.name.localeCompare(b.name));
  const prior=trend[trend.length-2],comparable=prior.available&&!prior.actualMissing&&!summary.actualMissing;
  const comparison=comparable?{month:prior.month,delta:round_(summary.actual-prior.actual),percent:prior.actual>0?Math.round((summary.actual-prior.actual)/prior.actual*1000)/10:null}:null;
  return {month:month,today:today,generatedAt:new Date().toISOString(),summary:summary,trend:trend,properties:properties,records:records,comparison:comparison,warnings:warnings};
}
function getDashboard(month) {
  const ss=book_();return dashboard_(ss,month,Utilities.formatDate(new Date(),ss.getSpreadsheetTimeZone(),'yyyy-MM-dd'));
}

// Residents belong to this private app, never to accounting cells.
const CITY_PREFIX_='RENT_CITY_RESIDENT_V1_';
const CITY_KINDS_=['deerfox','owlbunny','catdragon','otterturtle','pandatanuki','sealpenguin','corgibee','axolotldeer'];
const CITY_COLORS_=['peach','sage','lavender','honey'];
function cityStore_() { return PropertiesService.getUserProperties(); }
function cityId_(id) { if(typeof id!=='string'||!(/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i).test(id))throw Error('島民識別資料無效');return id.toLowerCase(); }
function cityResidents_(store) {
  const values=store.getProperties();return Object.keys(values).filter(k=>k.startsWith(CITY_PREFIX_)).map(k=>{
    try{const r=JSON.parse(values[k]);if(r.id&&r.name&&CITY_KINDS_.includes(r.kind)&&CITY_COLORS_.includes(r.color))return r}catch(e){}
    throw Error('島民資料無法讀取，請稍後再試');
  }).sort((a,b)=>a.createdAt.localeCompare(b.createdAt));
}
// Optional named islands live in private configuration, separate from accounting rows.
function cityExtraHomes_() {
  const text=PropertiesService.getScriptProperties().getProperty('CITY_EXTRA_ISLANDS');if(!text)return [];
  let names;try{names=JSON.parse(text)}catch(e){throw Error('CITY_EXTRA_ISLANDS 請使用名稱陣列格式')}
  if(!Array.isArray(names)||names.length>100||names.some(n=>typeof n!=='string'||!n.trim()||Array.from(n.trim()).length>120||/[\u0000-\u001f\u007f]/.test(n)))throw Error('額外島嶼名稱設定無效');
  return [...new Set(names.map(n=>n.trim()))];
}
function cityIslandList_(properties,extra) {
  const islands=properties.map(p=>Object.assign({},p,{hasRecords:true})),known=new Set(islands.map(p=>p.name));
  extra.forEach(name=>{if(!known.has(name)){known.add(name);islands.push({name:name,hasRecords:false})}});return islands;
}
function citySnapshot_(ss,month,today,store) {
  const records=dashboardRows_(ss,month,today),groups=new Map();
  records.forEach(r=>{if(!groups.has(r.property))groups.set(r.property,[]);groups.get(r.property).push(r)});
  const properties=Array.from(groups,([name,list])=>Object.assign({name:name,firstRow:list[0].row},summarize_(list))).sort((a,b)=>a.name.localeCompare(b.name));
  return {month:month,today:today,generatedAt:new Date().toISOString(),summary:summarize_(records),properties:properties,islands:cityIslandList_(properties,cityExtraHomes_()),records:records,residents:cityResidents_(store)};
}
function getCitySnapshot(month) { const ss=book_();return citySnapshot_(ss,month,Utilities.formatDate(new Date(),ss.getSpreadsheetTimeZone(),'yyyy-MM-dd'),cityStore_()); }
function createResident_(store,p,homes) {
  if(!p||typeof p!=='object')throw Error('請填寫島民資料');
  const id=cityId_(p.id);
  if(typeof p.name!=='string'||Array.from(p.name.trim()).length<1||Array.from(p.name.trim()).length>20||/[\u0000-\u001f\u007f]/.test(p.name))throw Error('名字請填 1 至 20 個字');
  if(!CITY_KINDS_.includes(p.kind)||!CITY_COLORS_.includes(p.color))throw Error('請選擇支援的精靈外型');
  if(typeof p.home!=='string'||(p.home!==''&&!homes.includes(p.home)))throw Error('請選擇有效的居住島嶼');
  if(Array.from(p.home).length>120)throw Error('物件名稱過長，請選中央戰情室');
  const key=CITY_PREFIX_+id,old=store.getProperty(key);
  if(old){const r=JSON.parse(old);if(r.name===p.name.trim()&&r.kind===p.kind&&r.color===p.color&&r.home===p.home)return r;throw Error('新增請求已使用，請重新開啟邀請視窗')}
  if(cityResidents_(store).length>=24)throw Error('城市目前最多可入住 24 位精靈');
  const r={id:id,name:p.name.trim(),kind:p.kind,color:p.color,home:p.home,hearts:0,visits:0,createdAt:new Date().toISOString(),lastActionId:''};
  store.setProperty(key,JSON.stringify(r));return r;
}
function createCityResident(p) {
  const ss=book_(),month=p&&p.month;sheet_(ss,month);
  const homes=dashboardRows_(ss,month,Utilities.formatDate(new Date(),ss.getSpreadsheetTimeZone(),'yyyy-MM-dd')).map(r=>r.property).concat(cityExtraHomes_()),lock=LockService.getScriptLock();lock.waitLock(20000);
  try{return createResident_(cityStore_(),p,homes)}finally{lock.releaseLock()}
}
function careResident_(store,p) {
  if(!p||!['wave','fruit'].includes(p.action))throw Error('互動方式無效');
  const id=cityId_(p.id),operationId=cityId_(p.operationId),key=CITY_PREFIX_+id,text=store.getProperty(key);
  if(!text)throw Error('找不到這位島民，請重新同步');
  const r=JSON.parse(text);if(r.lastActionId===operationId)return r;
  r.hearts=Math.min(999,(Number(r.hearts)||0)+(p.action==='fruit'?2:1));r.visits=Math.min(9999,(Number(r.visits)||0)+1);r.lastActionId=operationId;r.lastSeen=new Date().toISOString();
  store.setProperty(key,JSON.stringify(r));return r;
}
function careCityResident(p) { const lock=LockService.getScriptLock();lock.waitLock(20000);try{return careResident_(cityStore_(),p)}finally{lock.releaseLock()} }


