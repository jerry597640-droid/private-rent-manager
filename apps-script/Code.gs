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
  return {months:months_(ss),testMode:PropertiesService.getScriptProperties().getProperty('APP_MODE')!=='LIVE'};
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
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('每月收租管理').addMetaTag('viewport','width=device-width, initial-scale=1');
}
