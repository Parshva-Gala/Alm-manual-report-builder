import fs from 'node:fs';
import assert from 'node:assert/strict';
import ts from 'typescript';

// Isolated bridge tests. No DOM, Android runtime, timers or file dialogs are used.
const source=fs.readFileSync(new URL('../lib/platform.ts',import.meta.url),'utf8');
const js=ts.transpileModule(source,{compilerOptions:{target:ts.ScriptTarget.ES2022,module:ts.ModuleKind.CommonJS}}).outputText;
function harness({throws=false,native=true}={}){
  const notices=[],messages=[],timers=new Map(),frames=[];let uid=0,timerId=0;
  const window={dispatchEvent:e=>{notices.push(e.detail);return true},print:()=>messages.push({action:'browser-print'})};
  if(native)window.ESGNative={onmessage:null,postMessage:raw=>{if(throws)throw Error('Native bridge unavailable');messages.push(JSON.parse(raw))}};
  const setTimeout=(fn,ms)=>{const id=++timerId;timers.set(id,{fn,ms});return id};
  const clearTimeout=id=>timers.delete(id);
  const CustomEvent=class{constructor(type,options){this.type=type;this.detail=options.detail}};
  const m={exports:{}};
  new Function('module','exports','window','crypto','CustomEvent','setTimeout','clearTimeout','requestAnimationFrame',js)(m,m.exports,window,{randomUUID:()=>`request-${++uid}`},CustomEvent,setTimeout,clearTimeout,fn=>frames.push(fn));
  const reply=(id,fields={ok:true})=>window.ESGNative.onmessage({data:JSON.stringify({id,...fields})});
  const expire=()=>{const jobs=[...timers.values()];timers.clear();jobs.forEach(t=>t.fn())};
  const tick=async()=>{await Promise.resolve();await Promise.resolve();};
  const frame=async()=>{assert.ok(frames.length,'Expected a scheduled layout frame');frames.shift()();await tick()};
  return{...m.exports,window,notices,messages,timers,frames,reply,expire,tick,frame};
}
let passed=0;const failures=[];
async function test(name,fn){try{await fn();passed++;console.log('PASS '+name)}catch(e){failures.push({name,error:e.message});console.error('FAIL '+name+': '+e.message)}}
const noFalseSuccess=h=>assert.equal(h.notices.filter(n=>n.message==='File exported').length,0);

await test('export payload preserves exact text and valid native schema',async()=>{
  const h=harness(),text='{"bank":"Δemo", "line":"a\\nb"}\n',done=h.exportDocument('workspace.json',text);
  assert.deepEqual(h.messages[0],{id:'request-1',action:'export',name:'workspace.json',mime:'application/json',text});
  assert.equal(h.notices.length,0);h.reply('request-1');await done;assert.deepEqual(h.notices,[{message:'File exported',error:false}]);assert.equal(h.timers.size,0);
});
await test('no success before native write acknowledgement',async()=>{const h=harness(),done=h.exportDocument('x.csv','a,b','text/csv');await h.tick();noFalseSuccess(h);h.reply(h.messages[0].id,{ok:false,error:'Could not save'});await done;noFalseSuccess(h);assert.deepEqual(h.notices.at(-1),{message:'Could not save',error:true});});
await test('picker cancellation is explicit and never success',async()=>{const h=harness(),done=h.exportDocument('x.json','{}');h.reply(h.messages[0].id,{ok:false,cancelled:true});await done;assert.deepEqual(h.notices,[{message:'Export cancelled',error:false}]);assert.equal(h.timers.size,0);});
await test('native rejection produces readable error',async()=>{const h=harness(),done=h.exportDocument('x.json','{}');h.reply(h.messages[0].id,{ok:false,error:'Finish current operation first.'});await done;assert.deepEqual(h.notices,[{message:'Finish current operation first.',error:true}]);});
await test('synchronous bridge exception clears request timer',async()=>{const h=harness({throws:true});await h.exportDocument('x.json','{}');assert.equal(h.timers.size,0);assert.equal(h.notices.length,1);assert.equal(h.notices[0].error,true);noFalseSuccess(h);});
await test('unrelated reply cannot complete a pending request',async()=>{const h=harness(),done=h.exportDocument('x.json','{}');h.reply('unknown-request');await h.tick();assert.equal(h.notices.length,0);assert.equal(h.timers.size,1);h.reply(h.messages[0].id);await done;assert.equal(h.notices.length,1);});
await test('duplicate reply cannot produce duplicate success',async()=>{const h=harness(),done=h.exportDocument('x.json','{}');h.reply(h.messages[0].id);await done;h.reply(h.messages[0].id);await h.tick();assert.equal(h.notices.length,1);});
await test('concurrent requests are matched by ID out of order',async()=>{const h=harness(),a=h.exportDocument('a.json','{}'),b=h.exportDocument('b.csv','a','text/csv');assert.notEqual(h.messages[0].id,h.messages[1].id);h.reply(h.messages[1].id,{ok:false,cancelled:true});await b;h.reply(h.messages[0].id);await a;assert.deepEqual(h.notices.map(x=>x.message),['Export cancelled','File exported']);assert.equal(h.timers.size,0);});
await test('bounded timeout fails without inventing success',async()=>{const h=harness(),done=h.exportDocument('x.json','{}');assert.equal([...h.timers.values()][0].ms,300000);h.expire();await done;assert.equal(h.notices.length,1);assert.equal(h.notices[0].error,true);noFalseSuccess(h);h.reply(h.messages[0].id);await h.tick();assert.equal(h.notices.length,1);});
await test('malformed JSON reply cannot falsely complete export',async()=>{const h=harness(),done=h.exportDocument('x.json','{}');h.window.ESGNative.onmessage({data:'{broken'});await h.tick();noFalseSuccess(h);h.expire();await done;assert.ok(h.notices.some(x=>x.error));});
for(const [name,fields]of [
  ['string false',{ok:'false'}],['numeric true',{ok:1}],['missing ok',{}],['string cancelled',{ok:true,cancelled:'false'}],['object error',{ok:false,error:{message:'bad'}}]
])await test('malformed reply schema rejected: '+name,async()=>{const h=harness(),done=h.exportDocument('x.json','{}');h.reply(h.messages[0].id,fields);await h.tick();h.expire();await done;noFalseSuccess(h);assert.ok(h.notices.length);assert.ok(h.notices.every(n=>typeof n.message==='string'&&typeof n.error==='boolean'));assert.ok(h.notices.some(n=>n.error));});
await test('oversized UTF-8 export is rejected before bridge transport',async()=>{
  const h=harness(),done=h.exportDocument('large.txt','é'.repeat(2_621_441),'text/plain');await h.tick();
  // Allow a failing implementation to finish so the rest of the suite still runs.
  if(h.messages.length)h.reply(h.messages[0].id,{ok:false,error:'Native export exceeds 5 MiB'});
  await done;assert.equal(h.messages.length,0,'The web layer sent a payload over the native 5 MiB text limit');assert.ok(h.notices.some(n=>n.error));noFalseSuccess(h);
});
await test('unsupported export MIME is rejected before bridge transport',async()=>{
  const h=harness(),done=h.exportDocument('x.html','<html/>','text/html');await h.tick();if(h.messages.length)h.reply(h.messages[0].id,{ok:false,error:'Unsupported format'});await done;assert.equal(h.messages.length,0);assert.ok(h.notices.some(n=>n.error));noFalseSuccess(h);
});
await test('escaped JSON request is bounded before native transport',async()=>{
  // The text itself fits 5 MiB, but JSON expands each NUL to six characters.
  // Java rejects raw messages over (5 MiB * 2 + 8192) without a correlatable ID.
  const h=harness(),text='\0'.repeat(1_750_000),done=h.exportDocument('controls.txt',text,'text/plain');await h.tick();
  if(h.messages.length)h.reply(h.messages[0].id,{ok:false,error:'Native request too large'});
  await done;assert.equal(h.messages.length,0,'Serialized JSON exceeded the native message-character bound');assert.ok(h.notices.some(n=>n.error));noFalseSuccess(h);
});
await test('print waits for two layout frames before native request',async()=>{const h=harness(),done=h.printReport();assert.equal(h.messages.length,0);await h.frame();assert.equal(h.messages.length,0);await h.frame();assert.deepEqual(h.messages[0],{id:'request-1',action:'print'});h.reply(h.messages[0].id);await done;assert.equal(h.notices.length,0);assert.equal(h.timers.size,0);});
await test('print failure is visible; cancellation is not fake completion',async()=>{for(const fields of [{ok:false,error:'Print job failed'},{ok:false,cancelled:true}]){const h=harness(),done=h.printReport();await h.frame();await h.frame();h.reply(h.messages[0].id,fields);await done;noFalseSuccess(h);assert.equal(h.notices.length,fields.cancelled?0:1);if(!fields.cancelled)assert.equal(h.notices[0].error,true)}});
await test('print timeout is bounded and a late reply cannot turn it into success',async()=>{const h=harness(),done=h.printReport();await h.frame();await h.frame();h.expire();await done;assert.equal(h.notices.length,1);assert.equal(h.notices[0].error,true);h.reply(h.messages[0].id);await h.tick();assert.equal(h.notices.length,1);});
await test('browser-only print uses browser fallback',async()=>{const h=harness({native:false});await h.printReport();assert.deepEqual(h.messages,[{action:'browser-print'}]);assert.equal(h.timers.size,0);});
console.log(JSON.stringify({passed,failed:failures.length,failures,scope:'Pure web bridge handler tests with fake window and timers; does not validate Android file/print lifecycle or an emulator.'}));
if(failures.length)process.exitCode=1;
