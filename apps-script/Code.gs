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
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify([values,formulas]),Utilities.Charset.UTF_8));
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
  const text=PropertiesService.getScriptProperties().getProperty('CITY_EXTRA_ISLANDS');if(!text)return cityCustomIslands_().map(r=>r.name);
  let names;try{names=JSON.parse(text)}catch(e){throw Error('CITY_EXTRA_ISLANDS 請使用名稱陣列格式')}
  if(!Array.isArray(names)||names.length>100||names.some(n=>typeof n!=='string'||!n.trim()||Array.from(n.trim()).length>120||/[\u0000-\u001f\u007f]/.test(n)))throw Error('額外島嶼名稱設定無效');
  return [...new Set([...names.map(n=>n.trim()),...cityCustomIslands_().map(r=>r.name)])];
}
// Names and per-island shapes stay in the owner's private configuration.
function cityProfiles_() {
  const text=PropertiesService.getScriptProperties().getProperty('CITY_ISLAND_PROFILES');if(!text)return {};
  let profiles;try{profiles=JSON.parse(text)}catch(e){throw Error('CITY_ISLAND_PROFILES 請使用物件格式')}
  if(!profiles||typeof profiles!=='object'||Array.isArray(profiles)||Object.keys(profiles).length>100)throw Error('島嶼地塊設定無效');
  Object.entries(profiles).forEach(([name,p])=>{if(!name.trim()||name.length>120||!p||!Number.isInteger(p.tiles)||p.tiles<5||p.tiles>12||p.displayName!==undefined&&(typeof p.displayName!=='string'||!p.displayName.trim()||p.displayName.length>120||/[\u0000-\u001f\u007f]/.test(p.displayName)))throw Error('島嶼地塊數須為 5 至 12，名稱不可空白')});return profiles;
}
function cityIslandList_(properties,extra) {
  const islands=properties.map(p=>Object.assign({},p,{hasRecords:true})),known=new Set(islands.map(p=>p.name));
  extra.forEach(name=>{if(!known.has(name)){known.add(name);islands.push({name:name,hasRecords:false})}});
  const centralOnly=(PropertiesService.getScriptProperties().getProperty('CITY_CENTRAL_ONLY_ISLAND')||'').trim();
  const profiles=cityProfiles_();cityCustomIslands_().forEach(r=>{if(!Object.prototype.hasOwnProperty.call(profiles,r.name))profiles[r.name]={tiles:r.tiles}});islands.forEach(island=>{const profile=Object.prototype.hasOwnProperty.call(profiles,island.name)?profiles[island.name]:{};island.centralOnly=island.name===centralOnly;island.tileCount=profile.tiles||7;island.displayName=profile.displayName?.trim()||island.name});return islands;
}
function citySnapshot_(ss,month,today,store) {
  const records=dashboardRows_(ss,month,today),groups=new Map();
  records.forEach(r=>{if(!groups.has(r.property))groups.set(r.property,[]);groups.get(r.property).push(r)});
  const properties=Array.from(groups,([name,list])=>Object.assign({name:name,firstRow:list[0].row},summarize_(list))).sort((a,b)=>a.name.localeCompare(b.name));
  return {month:month,today:today,generatedAt:new Date().toISOString(),summary:summarize_(records),properties:properties,islands:cityIslandList_(properties,cityExtraHomes_()),records:records,residents:cityResidents_(store),buildings:cityBuildings_(store),characterNames:cityCharacterNames_(store),islandInfo:cityIslandInfo_(store)};
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



// Building configuration belongs to the private app, never to accounting cells.
const CITY_BUILDING_PREFIX_='RENT_CITY_BUILDING_V1_';
const CITY_BUILDING_UTF8_PREFIX_='RENT_CITY_BUILDING_V2_';
const CITY_BUILDING_TYPES_=['bungalow','villa','townhouse','duplex','apartment','tower','courtyard','shopfront'];
const CITY_BUILDING_STYLES_=['nordic','japanese','mediterranean','modern','brick','pastel'];
const CITY_BUILDING_ROOMS_=['studio','one','two','three','four','open'];
function cityBuildingKey_(home,plot){return CITY_BUILDING_UTF8_PREFIX_+Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify([home,plot]),Utilities.Charset.UTF_8))}
function cityBuildingLegacyKey_(home,plot){return CITY_BUILDING_PREFIX_+Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify([home,plot])))}
function cityBuildingValid_(r){return r&&typeof r.home==='string'&&Array.from(r.home).length<=120&&Number.isInteger(r.plot)&&(r.home===''?[13,15,17].includes(r.plot):r.plot===0)&&Number.isInteger(r.version)&&r.version>0&&(r.type===null||CITY_BUILDING_TYPES_.includes(r.type)&&CITY_BUILDING_STYLES_.includes(r.style)&&CITY_BUILDING_ROOMS_.includes(r.rooms)&&Number.isInteger(r.floors)&&r.floors>=1&&r.floors<=12&&Number.isInteger(r.rotation)&&r.rotation>=0&&r.rotation<=3)}
function cityBuildingPublic_(r){return {home:r.home,plot:r.plot,type:r.type,style:r.style,rooms:r.rooms,floors:r.floors,rotation:r.rotation,version:r.version}}
function cityBuildingRecords_(store){const values=store.getProperties(),records=new Map();for(const k of Object.keys(values).filter(k=>k.startsWith(CITY_BUILDING_PREFIX_)||k.startsWith(CITY_BUILDING_UTF8_PREFIX_))){let r;try{r=JSON.parse(values[k])}catch(e){}const modern=k.startsWith(CITY_BUILDING_UTF8_PREFIX_);if(!cityBuildingValid_(r)||k!==(modern?cityBuildingKey_(r.home,r.plot):cityBuildingLegacyKey_(r.home,r.plot)))throw Error('建築配置無法讀取，請重新同步');const id=JSON.stringify([r.home,r.plot]);if(modern||!records.has(id))records.set(id,r)}return Array.from(records.values())}
function cityBuildings_(store){return cityBuildingRecords_(store).map(cityBuildingPublic_)}
function saveCityBuilding(p){
  if(!p||typeof p!=='object'||!['place','remove'].includes(p.action))throw Error('建築操作無效');
  const ss=book_();sheet_(ss,p.month);
  const homes=dashboardRows_(ss,p.month,Utilities.formatDate(new Date(),ss.getSpreadsheetTimeZone(),'yyyy-MM-dd')).map(r=>r.property).concat(cityExtraHomes_());
  if(typeof p.home!=='string'||Array.from(p.home).length>120||p.home!==''&&!homes.includes(p.home))throw Error('請選擇有效的島嶼');
  if(!Number.isInteger(p.plot)||(p.home===''?![13,15,17].includes(p.plot):p.plot!==0))throw Error('這個地塊保留給地標、設施或步道，請選可配置基地');
  if(!Number.isInteger(p.version)||p.version<0)throw Error('建築版本無效，請重新同步');
  const operationId=cityId_(p.operationId),removing=p.action==='remove';
  const fields={home:p.home,plot:p.plot,type:removing?null:p.type,style:removing?null:p.style,rooms:removing?null:p.rooms,floors:removing?null:p.floors,rotation:removing?null:p.rotation};
  if(!cityBuildingValid_({...fields,version:1}))throw Error('請選擇支援的建築、風格、房型、樓層與朝向');
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const store=cityStore_(),key=cityBuildingKey_(p.home,p.plot),records=cityBuildingRecords_(store),old=records.find(r=>r.home===p.home&&r.plot===p.plot);
    if(old&&!cityBuildingValid_(old))throw Error('建築配置無法讀取，請重新同步');
    if(old?.lastOperationId===operationId){if(JSON.stringify({...cityBuildingPublic_(old),version:0})!==JSON.stringify({...fields,version:0}))throw Error('儲存請求已使用，請重新開啟建築工坊');return cityBuildingPublic_(old)}
    if((old?.version||0)!==p.version)throw Error('此基地配置已被更新。請關閉工坊並按「立即同步」，再重新配置');
    if(!old&&records.length>=128)throw Error('建築配置已達容量上限');
    const saved={...fields,version:p.version+1,lastOperationId:operationId};
    store.setProperty(key,JSON.stringify(saved));return cityBuildingPublic_(saved);
  }finally{lock.releaseLock()}
}

// Character display names are private preferences, separate from residents and accounting.
const CITY_NAME_PREFIX_='RENT_CITY_NAME_V1_';
const CITY_NAMED_ACTORS_=['owner','guest-fox','guest-owl','guest-catdragon','guest-otterturtle','guest-pandatanuki','guest-sealpenguin','guest-corgibee','guest-axolotldeer','guest-sprout','egg-moon','egg-garden','egg-boat'];
function cityNameTextValid_(name){return typeof name==='string'&&name===name.trim()&&Array.from(name).length>=1&&Array.from(name).length<=20&&!/[\u0000-\u001f\u007f\u202a-\u202e\u2066-\u2069]/.test(name)}
function cityNameValid_(r){return r&&typeof r.id==='string'&&(CITY_NAMED_ACTORS_.includes(r.id)||/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/.test(r.id))&&Number.isSafeInteger(r.version)&&r.version>0&&(r.name===null||cityNameTextValid_(r.name))}
function cityNamePublic_(r){return {id:r.id,name:r.name,version:r.version}}
function cityCharacterNames_(store){const values=store.getProperties();return Object.keys(values).filter(k=>k.startsWith(CITY_NAME_PREFIX_)).map(k=>{let r;try{r=JSON.parse(values[k])}catch(e){}if(!cityNameValid_(r)||k!==CITY_NAME_PREFIX_+r.id)throw Error('角色名稱無法讀取，請重新同步');return cityNamePublic_(r)}).sort((a,b)=>a.id.localeCompare(b.id))}
function saveCityCharacterName(p){
  if(!p||typeof p!=='object'||!['rename','reset'].includes(p.action))throw Error('改名操作無效');
  if(typeof p.id!=='string')throw Error('請選擇有效的角色');
  const id=CITY_NAMED_ACTORS_.includes(p.id)?p.id:cityId_(p.id),operationId=cityId_(p.operationId);
  if(!Number.isSafeInteger(p.version)||p.version<0||p.version>=Number.MAX_SAFE_INTEGER)throw Error('名稱版本無效，請重新同步');
  if(p.action==='rename'&&(typeof p.name!=='string'||/[\u0000-\u001f\u007f\u202a-\u202e\u2066-\u2069]/.test(p.name)))throw Error('名字請填 1 至 20 個字，不可含換行或控制字元');
  const name=p.action==='reset'?null:p.name.trim();if(name!==null&&!cityNameTextValid_(name))throw Error('名字請填 1 至 20 個字');
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const store=cityStore_();if(!CITY_NAMED_ACTORS_.includes(id)&&!cityResidents_(store).some(r=>r.id===id))throw Error('找不到這位角色，請重新同步');
    const key=CITY_NAME_PREFIX_+id,text=store.getProperty(key);let old;try{old=text?JSON.parse(text):null}catch(e){throw Error('角色名稱無法讀取，請重新同步')}
    if(old&&(!cityNameValid_(old)||old.id!==id))throw Error('角色名稱無法讀取，請重新同步');
    if(old?.lastOperationId===operationId){if(old.name!==name)throw Error('改名請求已使用，請重新開啟改名視窗');return cityNamePublic_(old)}
    if((old?.version||0)!==p.version)throw Error('這個角色的名字已被更新。請關閉視窗並按「立即同步」，再重新改名');
    const saved={id,name,version:p.version+1,lastOperationId:operationId};store.setProperty(key,JSON.stringify(saved));return cityNamePublic_(saved);
  }finally{lock.releaseLock()}
}


// Addresses and document references stay in private User Properties, never in public source or accounting cells.
const CITY_INFO_PREFIX_='RENT_CITY_ISLAND_INFO_V1_';
const CITY_INFO_UTF8_PREFIX_='RENT_CITY_ISLAND_INFO_V2_';
const CITY_DOC_CATEGORIES_=['insurance','parking','housing','other'];
function cityInfoText_(value,max,multiline=false){return typeof value==='string'&&Array.from(value).length<=max&&!(multiline?/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/:/[\u0000-\u001f\u007f]/).test(value)}
function cityDocumentUrl_(value){if(typeof value!=='string'||value.length>500)return null;const m=value.match(/^https:\/\/(?:drive\.google\.com\/file\/d\/([A-Za-z0-9_-]{10,200})\/(?:view|edit)|docs\.google\.com\/document\/d\/([A-Za-z0-9_-]{10,200})\/(?:view|edit))(?:\?([A-Za-z0-9_=%&.~-]*))?$/);if(!m)return null;const id=m[1]||m[2],key=(m[3]||'').match(/(?:^|&)resourcekey=([A-Za-z0-9_-]{1,200})(?:&|$)/);return {id,url:'https://drive.google.com/file/d/'+id+'/view'+(key?'?resourcekey='+key[1]:''),download:'https://drive.google.com/uc?export=download&id='+id+(key?'&resourcekey='+key[1]:'')}}
function cityInfoKey_(home){return CITY_INFO_UTF8_PREFIX_+Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,home,Utilities.Charset.UTF_8))}
function cityInfoLegacyKey_(home){return CITY_INFO_PREFIX_+Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,home))}
function cityInfoEmpty_(home){return {home,address:'',notes:'',documents:[],version:0}}
function cityInfoDocValid_(d){return d&&typeof d.id==='string'&&/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/.test(d.id)&&cityInfoText_(d.title,80)&&!!d.title.trim()&&CITY_DOC_CATEGORIES_.includes(d.category)&&['pdf','docx'].includes(d.format)&&!!cityDocumentUrl_(d.url)&&cityInfoText_(d.note,160,true)&&typeof d.removed==='boolean'}
function cityInfoValid_(r){return r&&cityInfoText_(r.home,120)&&!!r.home&&cityInfoText_(r.address,200,true)&&cityInfoText_(r.notes,500,true)&&Array.isArray(r.documents)&&r.documents.length<=24&&r.documents.every(cityInfoDocValid_)&&new Set(r.documents.map(d=>d.id)).size===r.documents.length&&new Set(r.documents.map(d=>cityDocumentUrl_(d.url).id)).size===r.documents.length&&r.documents.filter(d=>!d.removed).length<=12&&Number.isSafeInteger(r.version)&&r.version>0}
function cityInfoPublic_(r){return {home:r.home,address:r.address,notes:r.notes,documents:r.documents.map(d=>({id:d.id,title:d.title,category:d.category,format:d.format,url:d.url,note:d.note,removed:d.removed})),version:r.version}}
function cityInfoRecords_(store){const values=store.getProperties(),records=new Map();for(const k of Object.keys(values).filter(k=>k.startsWith(CITY_INFO_PREFIX_)||k.startsWith(CITY_INFO_UTF8_PREFIX_))){let r;try{r=JSON.parse(values[k])}catch(e){}const modern=k.startsWith(CITY_INFO_UTF8_PREFIX_);if(!cityInfoValid_(r)||k!==(modern?cityInfoKey_(r.home):cityInfoLegacyKey_(r.home)))throw Error('島嶼資料無法讀取，請重新同步');if(modern||!records.has(r.home))records.set(r.home,r)}return Array.from(records.values())}
function cityIslandInfo_(store){return cityInfoRecords_(store).map(cityInfoPublic_).sort((a,b)=>a.home.localeCompare(b.home))}
function saveCityIslandInfo(p){
  if(!p||typeof p!=='object'||!['profile','document','remove','restore'].includes(p.action))throw Error('島嶼資料操作無效');
  const ss=book_();sheet_(ss,p.month);const homes=dashboardRows_(ss,p.month,Utilities.formatDate(new Date(),ss.getSpreadsheetTimeZone(),'yyyy-MM-dd')).map(r=>r.property).concat(cityExtraHomes_());
  if(!cityInfoText_(p.home,120)||!p.home||!(homes.includes(p.home)||p.home==='@central'))throw Error('請選擇有效的物件島嶼');
  if(!Number.isSafeInteger(p.version)||p.version<0||p.version>=Number.MAX_SAFE_INTEGER)throw Error('資料版本無效，請重新同步');
  const operationId=cityId_(p.operationId);let fields;
  if(p.action==='profile'){if(!cityInfoText_(p.address,200,true)||!cityInfoText_(p.notes,500,true))throw Error('地址限 200 字，備註限 500 字');fields={address:p.address.trim(),notes:p.notes.trim()}}
  else if(p.action==='document'){const d=p.document||{},url=cityDocumentUrl_(d.url),id=cityId_(d.id);fields={id,title:typeof d.title==='string'?d.title.trim():d.title,category:d.category,format:d.format,url:url?.url,note:typeof d.note==='string'?d.note.trim():d.note,removed:false};if(!cityInfoDocValid_(fields))throw Error('請填寫文件名稱、類別、PDF／Word 格式及有效的 Google Drive 文件連結')}
  else fields={id:cityId_(p.id)};
  const signature=JSON.stringify({action:p.action,fields}),lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const store=cityStore_(),key=cityInfoKey_(p.home),old=cityInfoRecords_(store).find(r=>r.home===p.home);
    if(old&&(!cityInfoValid_(old)||old.home!==p.home))throw Error('島嶼資料無法讀取，請重新同步');
    if(old?.lastOperationId===operationId){if(old.lastSignature!==signature)throw Error('儲存請求已使用，請重新開啟島嶼資料');return cityInfoPublic_(old)}
    if((old?.version||0)!==p.version)throw Error('這座島的資料已被更新。請關閉視窗並按「立即同步」，再重新編輯');
    const next=old?cityInfoPublic_(old):cityInfoEmpty_(p.home);
    if(p.action==='profile'){next.address=fields.address;next.notes=fields.notes}
    else if(p.action==='document'){
      const index=next.documents.findIndex(d=>d.id===fields.id),duplicate=next.documents.find(d=>d.id!==fields.id&&cityDocumentUrl_(d.url).id===cityDocumentUrl_(fields.url).id);
      if(duplicate)throw Error(duplicate.removed?'此文件在「最近移除」中，請直接還原':'此文件已加入這座島');
      if(index>=0&&next.documents[index].removed)throw Error('請先還原這份文件，再編輯');
      if(index<0){if(next.documents.length>=24)throw Error('這座島已達 24 份文件紀錄上限');if(next.documents.filter(d=>!d.removed).length>=12)throw Error('每座島最多顯示 12 份文件');next.documents.push(fields)}else next.documents[index]=fields;
    }else{const doc=next.documents.find(d=>d.id===fields.id);if(!doc)throw Error('找不到這份文件，請重新同步');if(p.action==='restore'&&doc.removed&&next.documents.filter(d=>!d.removed).length>=12)throw Error('每座島最多顯示 12 份文件');doc.removed=p.action==='remove'}
    next.version=p.version+1;const saved={...next,lastOperationId:operationId,lastSignature:signature},encoded=JSON.stringify(saved);
    if(Utilities.newBlob(encoded).getBytes().length>8000)throw Error('這座島資料已達儲存上限，請縮短備註或文件名稱');
    store.setProperty(key,encoded);return cityInfoPublic_(next);
  }finally{lock.releaseLock()}
}


const CITY_CUSTOM_ISLAND_PREFIX_='RENT_CITY_CUSTOM_ISLAND_V1_';
function cityCustomIslands_(){const values=cityStore_().getProperties();return Object.keys(values).filter(k=>k.startsWith(CITY_CUSTOM_ISLAND_PREFIX_)).map(k=>{let r;try{r=JSON.parse(values[k])}catch(e){}if(!r||k!==CITY_CUSTOM_ISLAND_PREFIX_+cityId_(r.id)||!cityInfoText_(r.name,40)||!r.name.trim()||!Number.isInteger(r.tiles)||r.tiles<5||r.tiles>12)throw Error('自建島嶼資料無法讀取');return r}).sort((a,b)=>a.createdAt.localeCompare(b.createdAt))}
function cityNewRent_(p){if(p.createRent!==true)return null;const r=p.rent||{},text=String(r.amount??'').trim();if(!/^\d+(\.\d{1,2})?$/.test(text)||Number(text)<=0||Number(text)>999999999||!Number.isInteger(r.due)||r.due<1||r.due>31)throw Error('新增帳目請填大於 0 的每月租金與 1 至 31 的應收日');return {month:p.month,amount:Number(text),due:r.due}}
function cityRentColumns_(sh){cols_(sh);const h=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(x=>x.trim());for(const key of ['每月應收日','應收金額','本月應收','差額'])if(h.filter(x=>x===key).length>1)throw Error('收租欄位重複：'+key);if(!h.includes('每月應收日')||!h.includes('應收金額')&&!h.includes('本月應收'))throw Error('此月份缺少應收日或應收金額欄位，無法建立帳目');return h}
function cityAppendRent_(ss,saved){const sh=sheet_(ss,saved.rent.month),headers=cityRentColumns_(sh),values=sh.getDataRange().getValues(),marker='[棲島建立:'+saved.id+']',noteCol=headers.indexOf('備註'),nameCol=headers.indexOf('物件名稱'),found=values.findIndex((r,i)=>i>0&&String(r[noteCol]||'').includes(marker));if(found>=0){if(values[found][nameCol]!==saved.name)throw Error('建立識別與物件不符，請檢查帳表');return found+1}if(values.slice(1).some(r=>String(r[nameCol]||'').trim()===saved.name))throw Error('此月份已有同名物件，請檢查帳表後再試');const row=sh.getLastRow()+1;if(row>sh.getMaxRows())sh.insertRowsAfter(sh.getMaxRows(),1);const range=sh.getRange(row,1,1,headers.length);if(!range.canEdit()||range.getFormulas()[0].some(Boolean))throw Error('新增資料列禁止修改或含有公式，已停止');const next=headers.map(()=>''),set=(key,v)=>{const i=headers.indexOf(key);if(i>=0)next[i]=v};set('物件名稱',saved.name);set('類型','租金');set('每月應收日',saved.rent.due);set('應收金額',saved.rent.amount);set('本月應收',saved.rent.amount);set('是否已收',false);set('備註',marker);const letter=n=>{let s='';for(n++;n>0;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s},expected=headers.indexOf(headers.includes('本月應收')?'本月應收':'應收金額');set('差額','='+letter(headers.indexOf('實收金額'))+row+'-'+letter(expected)+row);range.setValues([next]);SpreadsheetApp.flush();return row}
function createCityIsland(p){if(!p||!cityInfoText_(p.name,40)||!p.name.trim()||/^[=+@'\-]/.test(p.name.trim())||['中央戰情室','@central','__proto__','constructor','prototype'].includes(p.name.trim())||!Number.isInteger(p.tiles)||p.tiles<5||p.tiles>12)throw Error('請填寫 1 至 40 字島名（不可使用公式開頭），並選擇 5 至 12 塊地塊');const id=cityId_(p.id),name=p.name.trim(),rent=cityNewRent_(p),ss=book_(),sh=sheet_(ss,p.month);if(rent)cityRentColumns_(sh);const signature=JSON.stringify({name,tiles:p.tiles,rent}),lock=LockService.getScriptLock();lock.waitLock(20000);try{const store=cityStore_(),key=CITY_CUSTOM_ISLAND_PREFIX_+id,previous=store.getProperty(key);let saved=previous?JSON.parse(previous):null;if(saved){if(saved.signature!==signature)throw Error('此建立請求已使用，請重新開啟新增島嶼')}else{const allNames=new Set(cityExtraHomes_());months_(ss).forEach(month=>dashboardRows_(ss,month,Utilities.formatDate(new Date(),ss.getSpreadsheetTimeZone(),'yyyy-MM-dd')).forEach(r=>allNames.add(r.property)));if(allNames.has(name)||Object.values(cityProfiles_()).some(r=>r.displayName?.trim()===name))throw Error('這個島名已存在，請換一個名稱');if(cityCustomIslands_().length>=8||allNames.size>=16)throw Error('目前最多可新增 8 座島嶼，城市最多 16 座物件島');saved={id,name,tiles:p.tiles,createdAt:new Date().toISOString(),signature,rent,rentReady:!rent};store.setProperty(key,JSON.stringify(saved))}if(saved.rent&&!saved.rentReady){try{saved.rentRow=cityAppendRent_(ss,saved);saved.rentReady=true;store.setProperty(key,JSON.stringify(saved))}catch(e){throw Error('島嶼已保存，收租帳目尚未確認：'+e.message+'。請保留此表單再次送出；不會重複新增')}}return {id,name,tiles:saved.tiles,rentCreated:!!saved.rent&&saved.rentReady,month:saved.rent?.month}}finally{lock.releaseLock()}}
