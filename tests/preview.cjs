const http=require('node:http'),fs=require('node:fs'),path=require('node:path'),vm=require('node:vm'),crypto=require('node:crypto');
function createFixture(){
const headers=['物件名稱','類型','每月應收日','應收金額','本月應收','實際入帳日','實收金額','貸款金額','差額','是否已收','備註'];
const sheets=Array.from({length:10},(_,i)=>{
 const month='2026-'+String(i+1).padStart(2,'0'),latest=i===9;
 const values=[headers.slice(),
 ['青禾寓所','租金',5,18000,18000,month+'-05',18000,6000,0,true,''],
 ['晴川小居','租金',10,22000,22000,latest?month+'-10':month+'-09',latest?12000:22000,8000,0,!latest,''],
 ['光映公寓','租金',28,24000,24000,latest?'':month+'-28',latest?0:24000,9000,0,!latest,''],
 ['光映公寓','停車位',28,2500,2500,latest?'':month+'-28',latest?0:2500,0,0,!latest,''],
 ['青禾寓所','停車位',5,2500,2500,month+'-05',2500,0,0,!latest,''],
 ['山嵐居所','租金',20,20000,20000,month+'-20',latest?15000:20000,7000,0,true,''],
 ['河畔小築','租金',26,16000,16000,month+'-25',latest?16500:16000,4000,0,true,'']];
 if(i===8)values[2][6]=15000;
 const sh={getMaxRows:()=>1000,insertRowsAfter:()=>{},getName:()=>month,getLastColumn:()=>headers.length,getLastRow:()=>values.length};
 sh.getRange=(r,c,n=1,w=1)=>({getValues:()=>values.slice(r-1,r-1+n).map(v=>v.slice(c-1,c-1+w)),getDisplayValues:()=>values.slice(r-1,r-1+n).map(v=>v.slice(c-1,c-1+w).map(String)),getFormulas:()=>Array.from({length:n},()=>Array(w).fill('')),canEdit:()=>true,setValues:rows=>{rows.forEach((v,i)=>{values[r-1+i]=Array.from(v)})},setValue:value=>{values[r-1][c-1]=value}});
 sh.getDataRange=()=>sh.getRange(1,1,values.length,headers.length);return sh;
});
const ss={getSheets:()=>sheets,getSheetByName:m=>sheets.find(s=>s.getName()===m),getSpreadsheetTimeZone:()=> 'Asia/Taipei'};
const userValues={},userStore={getProperties:()=>({...userValues}),getProperty:k=>userValues[k]??null,setProperty:(k,v)=>{userValues[k]=v},deleteProperty:k=>{delete userValues[k]}};
const ctx={Date,PropertiesService:{getScriptProperties:()=>({getProperty:k=>k==='SPREADSHEET_ID'?'fake-copy':k==='APP_MODE'?'TEST':k==='CITY_ISLAND_PROFILES'?JSON.stringify({'青禾寓所':{tiles:7},'晴川小居':{tiles:9},'光映公寓':{tiles:8},'山嵐居所':{tiles:9},'河畔小築':{tiles:8},'暮光公寓':{tiles:5,displayName:'暮光4樓公寓'},'晨光小院':{tiles:7}}):k==='CITY_EXTRA_ISLANDS'?'["暮光公寓","晨光小院"]':null}),getUserProperties:()=>userStore},SpreadsheetApp:{openById:()=>ss,flush:()=>{}},LockService:{getScriptLock:()=>({waitLock:()=>{},releaseLock:()=>{}})},Utilities:{Charset:{UTF_8:'UTF-8'},newBlob:s=>({getBytes:()=>Buffer.from(s,'utf8')}),DigestAlgorithm:{SHA_256:'sha256'},computeDigest:(_,s)=>crypto.createHash('sha256').update(s).digest(),base64EncodeWebSafe:b=>Buffer.from(b).toString('base64url'),parseDate:s=>new Date(s.slice(0,10)+'T12:00:00+08:00'),formatDate:d=>new Date(d.getTime()+8*3600000).toISOString().slice(0,10)}};
vm.createContext(ctx);vm.runInContext(fs.readFileSync(path.join(__dirname,'../apps-script/Code.gs'),'utf8'),ctx);
return {ctx,ss,userStore};
}
module.exports={createFixture};
if(require.main===module){
const {ctx,ss,userStore}=createFixture();
const sheets=ss.getSheets();
const mock=`<script>function callBuilder(){let done,failed;const chain={withSuccessHandler(fn){done=fn;return chain},withFailureHandler(fn){failed=fn;return chain}};['getContext','getRows','getDashboard','saveRent','getCitySnapshot','createCityResident','careCityResident','saveCityBuilding','saveCityCharacterName','saveCityIslandInfo','createCityIsland'].forEach(method=>chain[method]=(...args)=>fetch('/api/'+method+location.search,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(args)}).then(r=>r.json()).then(x=>x.error?failed({message:x.error}):done(x.data)).catch(failed));return chain}window.google={script:{get run(){return callBuilder()}}};</script>`;
http.createServer(async(req,res)=>{
 const url=new URL(req.url,'http://localhost');res.setHeader('Cache-Control','no-store');
 if(url.pathname==='/mobile'){
  const width=[320,390,430].includes(Number(url.searchParams.get('width')))?Number(url.searchParams.get('width')):390;
  res.setHeader('Content-Type','text/html; charset=utf-8');
  res.end('<title>手機版介面驗證 · 虛構資料</title><style>body{margin:0;background:#dce3df;display:flex;justify-content:center}iframe{width:'+width+'px;height:95vh;border:0;margin:12px;background:white}</style><iframe title="手機預覽" src="/form"></iframe>');return;
 }
 if(url.pathname==='/'){res.setHeader('Content-Type','text/html; charset=utf-8');res.end('<a href="/form">開啟虛構資料預覽</a>');return}
 if(url.pathname==='/form'){res.setHeader('Content-Type','text/html; charset=utf-8');const html=fs.readFileSync(path.join(__dirname,'../apps-script/Index.html'),'utf8').replace("<?!= HtmlService.createHtmlOutputFromFile('City').getContent(); ?>",fs.readFileSync(path.join(__dirname,'../apps-script/City.html'),'utf8'));res.end(html.replace('<script>',mock+'<script>'));return}
 const method=url.pathname.slice(5);if(req.method==='POST'&&url.pathname.startsWith('/api/')&&['getContext','getRows','getDashboard','saveRent','getCitySnapshot','createCityResident','careCityResident','saveCityBuilding','saveCityCharacterName','saveCityIslandInfo','createCityIsland'].includes(method)){
  try{let body='';for await(const chunk of req)body+=chunk;const args=JSON.parse(body);let data;
   if(url.searchParams.has('error')&&['getDashboard','getCitySnapshot'].includes(method))throw Error('測試彙整失敗');
   if(url.searchParams.has('residentError')&&['createCityResident','careCityResident','saveCityBuilding','saveCityCharacterName','saveCityIslandInfo','createCityIsland'].includes(method))throw Error('測試島民儲存失敗');
   if(url.searchParams.has('saveError')&&method==='saveRent')throw Error('測試儲存失敗');
   if(method==='getContext')data={months:sheets.map(s=>s.getName()),today:'2026-10-26',testMode:true};
   else if(method==='getDashboard')data=ctx.dashboard_(ss,args[0],'2026-10-26');else if(method==='getCitySnapshot')data=ctx.citySnapshot_(ss,args[0],'2026-10-26',userStore);else data=ctx[method](...args);
   if(url.searchParams.has('empty')&&method==='getDashboard'){data.summary=ctx.summarize_([]);data.records=[];data.properties=[]}
   res.setHeader('Content-Type','application/json');res.end(JSON.stringify({data}));
  }catch(e){res.setHeader('Content-Type','application/json');res.end(JSON.stringify({error:e.message}))}return;
 }
 res.writeHead(404);res.end('Not found');
}).listen(4178,'127.0.0.1',()=>console.log('Local fictional rent dashboard ready'));

}
