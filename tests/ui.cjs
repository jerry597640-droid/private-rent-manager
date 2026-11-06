// DOM integration checks with fictional Sheets and speech adapters. No live account writes.
const fs=require('node:fs'),assert=require('node:assert/strict');
const {JSDOM,VirtualConsole}=require('jsdom');
const {createFixture}=require('./preview.cjs');
const source=fs.readFileSync(__dirname+'/../apps-script/Index.html','utf8');
const stubCity='<section id="city-view"><select id="city-month"></select></section><script>const City={loading:false,wake(){},load(){},error(){}};</script>';
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
async function until(test){for(let i=0;i<200;i++){if(test())return;await sleep(5)}throw Error('Timed out waiting for UI')}
let count=0;
async function check(name,run,options={}){
  const fixture=createFixture(),calls=[],errors=[],spoken=[],pending=[],saved={};
  if(options.setup)options.setup(fixture);
  const console=new VirtualConsole();console.on('jsdomError',e=>errors.push(e));
  let mode=options.speech||'end';
  const dom=new JSDOM(source.replace("<?!= HtmlService.createHtmlOutputFromFile('City').getContent(); ?>",stubCity),{
    url:'https://rent.example.test/',runScripts:'dangerously',pretendToBeVisual:true,virtualConsole:console,
    beforeParse(w){
      w.HTMLElement.prototype.scrollIntoView=function(){};
      w.HTMLDialogElement.prototype.showModal=function(){this.open=true};
      w.HTMLDialogElement.prototype.close=function(){this.open=false};
      if(options.storageBlocked)Object.defineProperty(w,'localStorage',{get(){throw Error('storage blocked')}});
      else if(options.outfit)w.localStorage.setItem('rent-patrol-landlord',options.outfit);
      if(options.fastTimers){const timer=w.setTimeout.bind(w);w.setTimeout=(fn,ms,...args)=>timer(fn,ms>=6000?10:ms,...args)}
      if(mode!=='absent'){
        w.SpeechSynthesisUtterance=class{constructor(text){this.text=text}};
        w.speechSynthesis={getVoices:()=>[{name:'China',lang:'zh-CN'},{name:'美佳',lang:'zh-TW'}],cancel(){saved.cancelCount=(saved.cancelCount||0)+1},speak(u){spoken.push(u);if(mode!=='hang')w.setTimeout(()=>mode==='error'?u.onerror():u.onend(),1)}};
      }
      function runner(){let ok,fail;const chain={withSuccessHandler(fn){ok=fn;return chain},withFailureHandler(fn){fail=fn;return chain}};
        ['getContext','getDashboard','getRows','saveRent'].forEach(method=>chain[method]=(...args)=>{
          calls.push({method,args});
          const reply=()=>{try{
            if(options.failure===method)throw Error('模擬服務失敗');
            let data=method==='getContext'?{months:fixture.ss.getSheets().map(s=>s.getName()),today:'2026-10-26',testMode:true}:method==='getDashboard'?fixture.ctx.dashboard_(fixture.ss,args[0],'2026-10-26'):fixture.ctx[method](...args);
            ok(JSON.parse(JSON.stringify(data)));
          }catch(e){fail(e)}};
          if(options.deferRows&&method==='getRows')pending.push(reply);else w.setTimeout(reply,0);
        });return chain;
      }
      w.google={script:{get run(){return runner()}}};
    }
  });
  const w=dom.window,d=w.document,$=id=>d.getElementById(id);
  const click=id=>$(id).click(),input=(id,value)=>{$(id).value=value;$(id).dispatchEvent(new w.Event('input',{bubbles:true}))};
  const change=(id,value)=>{$(id).value=value;$(id).dispatchEvent(new w.Event('change',{bubbles:true}))};
  try{
    await until(()=>$('dash-status').dataset.kind!=='busy');
    await run({w,d,$,click,input,change,calls,spoken,pending,saved,fixture,setSpeech:x=>mode=x});
    assert.deepEqual(errors,[]);count++;process.stdout.write('PASS '+name+'\n');
  }finally{w.close()}
}
(async()=>{
await check('dashboard starts first with authoritative totals and actual receipt dates',async({$,calls})=>{
  await until(()=>$('patrol-records').textContent.includes('尚未填寫'));
  assert.equal($('city-view').hidden,true);assert.equal($('dashboard-view').hidden,false);
  assert.equal($('dash-expected').textContent,'105,000');assert.equal($('dash-actual').textContent,'64,000');assert.equal($('dash-overdue').textContent,'15,000');
  assert.deepEqual(calls.map(c=>c.method),['getContext','getDashboard','getRows']);
});
await check('month arrows, current month and unavailable boundaries',async({$,click,change})=>{
  assert.equal($('dash-next-month').disabled,true);click('dash-prev-month');await until(()=>$('dash-month').value==='2026-09'&&!$('dash-current').disabled);
  click('dash-current');await until(()=>$('dash-month').value==='2026-10'&&!$('dash-prev-month').disabled);
  change('dash-month','2026-01');await until(()=>$('dash-status').dataset.kind==='success');assert.equal($('dash-prev-month').disabled,true);
});
await check('month switching skips absent calendar tabs',async({$,click})=>{
  click('dash-prev-month');await until(()=>$('dash-status').dataset.kind==='success');assert.equal($('dash-month').value,'2026-08');
},{setup:({ss})=>{const sheets=ss.getSheets().filter(s=>s.getName()!=='2026-09');ss.getSheets=()=>sheets}});
await check('empty month clears route, records and every patrol action',async({$})=>{
  await sleep(15);assert.ok($('patrol-dialogue').textContent.includes('尚無收租物件'));
  assert.equal($('patrol-route').children.length,0);assert.equal($('patrol-records').children.length,0);
  for(const id of ['patrol-start','patrol-stop','patrol-speak','patrol-prev','patrol-next','patrol-open','patrol-select'])assert.equal($(id).disabled,true,id);
},{setup:({ss})=>{const sh=ss.getSheetByName('2026-10');const old=sh.getDataRange;sh.getDataRange=()=>{const r=old();return {...r,getValues:()=>r.getValues().slice(0,1),getFormulas:()=>r.getFormulas().slice(0,1)}}}});
await check('dashboard read failure exposes retry while entry stays available',async({$,click})=>{
  assert.equal($('dash-status').dataset.kind,'error');assert.equal($('dash-content').hidden,true);assert.equal($('reload').disabled,false);
  click('nav-entry');await until(()=>$('record-content').hidden===false);assert.equal($('amount').disabled,false);
},{failure:'getDashboard'});
await check('receipt read failure keeps dashboard and explicitly marks unknown dates',async({$})=>{
  await until(()=>$('patrol-records').textContent.includes('請至登錄確認'));assert.equal($('dash-content').hidden,false);assert.equal($('dash-actual').textContent,'64,000');
},{failure:'getRows'});
await check('late receipt response cannot overwrite another month',async({$,click,pending})=>{
  assert.equal(pending.length,1);click('dash-prev-month');await until(()=>pending.length===2);
  pending[0]();assert.ok($('patrol-records').textContent.includes('讀取中'));
  pending[1]();assert.ok($('patrol-records').textContent.includes('2026-09'));assert.ok(!$('patrol-records').textContent.includes('2026-10'));
},{deferRows:true});
await check('patrol shows missing amounts as unknown and incomplete totals',async({$,change})=>{
  const option=[...$('patrol-select').options].find(o=>o.textContent==='光映公寓');change('patrol-select',option.value);assert.ok($('patrol-records').textContent.includes('待確認'));assert.ok($('patrol-dialogue').textContent.includes('已知部分'));assert.ok($('patrol-state').textContent.includes('需核對'));
},{setup:({ss})=>{ss.getSheetByName('2026-10').getRange(4,7).setValue('#REF!')}});
await check('checked but partially paid island still needs review',async({$,change})=>{
  const option=[...$('patrol-select').options].find(o=>o.textContent==='山嵐居所');change('patrol-select',option.value);
  assert.ok($('patrol-state').textContent.includes('需核對'));assert.ok($('patrol-records').textContent.includes('5,000'));
});
await check('individual patrol records open the exact month and row',async({$,click})=>{
  $('patrol-records').querySelectorAll('button')[1].click();await until(()=>$('record-content').hidden===false);
  assert.equal($('record-name').textContent,'光映公寓');assert.equal($('record-meta').textContent,'2026-10 · 停車位');assert.equal($('item').value,'5');
});
await check('outfits are independent and saved without writing Sheets',async({$,change,calls,w})=>{
  change('landlord-outfit','builder');change('landlady-outfit','detective');assert.equal($('patrol-couple').textContent,'👷🕵️‍♀️');
  assert.equal(w.localStorage.getItem('rent-patrol-landlord'),'builder');assert.ok(!calls.some(c=>c.method==='saveRent'));
});
await check('saved outfit restored on load',async({$})=>assert.equal($('landlord-outfit').value,'casual'),{outfit:'casual'});
await check('blocked device storage does not break wardrobe or loading',async({$,change})=>{change('landlady-outfit','formal');assert.ok($('patrol-couple').textContent.includes('👩‍🎓'))},{storageBlocked:true});
await check('one-tap tour speaks both roles at every island in Taiwan Chinese',async({$,click,spoken})=>{
  click('patrol-start');await until(()=>$('patrol-voice-note').textContent.includes('巡視完成'));assert.equal(spoken.length,10);
  assert.ok(spoken.every(u=>u.lang==='zh-TW'&&u.voice.lang==='zh-TW'));assert.equal($('patrol-position').textContent,'5 / 5 島');assert.equal($('patrol-stop').disabled,true);
});
await check('speech stop settles a hung utterance and prevents later islands',async({$,click,spoken})=>{
  click('patrol-start');assert.equal(spoken.length,1);click('patrol-stop');await sleep(20);assert.equal(spoken.length,1);assert.equal($('patrol-start').disabled,false);assert.equal($('patrol-position').textContent,'1 / 5 島');
},{speech:'hang'});
await check('leaving dashboard cancels standalone speech',async({$,click,spoken,saved})=>{
  click('patrol-speak');assert.equal($('patrol-stop').disabled,false);click('nav-entry');await until(()=>$('record-content').hidden===false);
  assert.ok(saved.cancelCount>0);assert.equal(spoken.length,1);assert.equal($('patrol-stop').disabled,true);
},{speech:'hang'});
await check('backgrounding cancels the tour',async({w,d,$,click,spoken})=>{
  click('patrol-start');Object.defineProperty(d,'hidden',{value:true,configurable:true});d.dispatchEvent(new w.Event('visibilitychange'));await sleep(15);
  assert.equal(spoken.length,1);assert.equal($('patrol-stop').disabled,true);
},{speech:'hang'});
await check('manual island selection interrupts a running tour',async({$,click,change,spoken})=>{
  click('patrol-start');change('patrol-select','2');await sleep(15);assert.equal($('patrol-position').textContent,'3 / 5 島');assert.equal(spoken.length,1);assert.equal($('patrol-start').disabled,false);
},{speech:'hang'});
await check('no speech support waits on each island instead of flashing to the end',async({$,click})=>{
  click('patrol-start');await sleep(5);assert.equal($('patrol-position').textContent,'1 / 5 島');assert.ok($('patrol-voice-note').textContent.includes('文字巡島'));
  click('patrol-stop');assert.equal($('patrol-stop').disabled,true);
},{speech:'absent'});
await check('speech errors complete through paced text fallback',async({$,click})=>{
  click('patrol-start');await until(()=>$('patrol-voice-note').textContent.includes('含文字巡島'));assert.equal($('patrol-position').textContent,'5 / 5 島');
},{speech:'error',fastTimers:true});
await check('speech watchdog recovers when the browser emits no events',async({$,click})=>{
  click('patrol-speak');await until(()=>$('patrol-voice-note').textContent.includes('無法播報'));assert.equal($('patrol-start').disabled,false);
},{speech:'hang',fastTimers:true});
await check('dirty navigation cancel keeps form; confirm discards safely',async({$,click,input})=>{
  click('nav-entry');await until(()=>$('record-content').hidden===false);input('amount','123');click('nav-dashboard');assert.equal($('confirm-dialog').open,true);
  click('confirm-cancel');assert.equal($('amount').value,'123');assert.equal($('entry-view').hidden,false);
  click('nav-dashboard');click('confirm-ok');await until(()=>$('dashboard-view').hidden===false&&$('dash-status').dataset.kind==='success');
  click('nav-entry');await until(()=>$('record-content').hidden===false);assert.equal($('amount').value,'18000');
});
await check('save requires confirmation and round-trips exact existing payload',async({$,click,input,calls,fixture,w})=>{
  click('nav-entry');await until(()=>$('record-content').hidden===false);input('date','2026-10-08');input('amount','18500');input('note','測試備註');
  $('rent-form').dispatchEvent(new w.Event('submit',{cancelable:true}));assert.equal(calls.filter(c=>c.method==='saveRent').length,0);assert.equal($('confirm-dialog').open,true);
  click('confirm-ok');await until(()=>$('status').dataset.kind==='success');const payload=calls.find(c=>c.method==='saveRent').args[0];
  assert.deepEqual(Object.keys(payload).sort(),['month','row','property','type','version','date','amount','paid','note'].sort());
  assert.equal(payload.month,'2026-10');assert.equal(payload.row,2);assert.ok(payload.version);assert.equal($('amount').value,'18500');assert.equal($('date').value,'2026-10-08');assert.equal($('save').disabled,true);
  const cells=fixture.ss.getSheetByName('2026-10').getDataRange().getValues()[1];assert.equal(cells[3],18000);assert.equal(cells[4],18000);assert.equal(cells[7],6000);
});
await check('save failure preserves input for retry',async({$,click,input,w})=>{
  click('nav-entry');await until(()=>$('record-content').hidden===false);input('amount','18500');$('rent-form').dispatchEvent(new w.Event('submit',{cancelable:true}));click('confirm-ok');await until(()=>$('status').dataset.kind==='error');assert.equal($('amount').value,'18500');assert.equal($('save').disabled,false);
},{failure:'saveRent'});
await check('paid with missing date is blocked before server write',async({$,click,input,calls,w})=>{
  click('nav-entry');await until(()=>$('record-content').hidden===false);input('date','');$('rent-form').dispatchEvent(new w.Event('submit',{cancelable:true}));assert.equal($('status').dataset.kind,'error');assert.ok(!calls.some(c=>c.method==='saveRent'));
});
process.stdout.write(count+' frontend integration checks passed\n');
})().catch(e=>{console.error(e);process.exitCode=1});
