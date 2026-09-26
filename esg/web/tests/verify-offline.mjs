import fs from 'node:fs';
import assert from 'node:assert/strict';
import ts from 'typescript';
import {IDBFactory, IDBObjectStore} from 'fake-indexeddb';
import * as financial from '../lib/financial.mjs';
import * as emissions from '../lib/emissions.mjs';
import * as esg from '../lib/esg.mjs';

const json=n=>JSON.parse(fs.readFileSync(new URL('../lib/'+n,import.meta.url),'utf8'));
const deps={'./financial.mjs':financial,'./emissions.mjs':emissions,'./esg.mjs':esg,'./esg-reference.json':json('esg-reference.json'),'./emissions-reference.json':json('emissions-reference.json')};
function compile(name){
  const source=fs.readFileSync(new URL('../lib/'+name+'.ts',import.meta.url),'utf8');
  const js=ts.transpileModule(source,{compilerOptions:{target:ts.ScriptTarget.ES2022,module:ts.ModuleKind.CommonJS,esModuleInterop:true}}).outputText;
  const m={exports:{}};new Function('require','module','exports',js)(id=>{if(!(id in deps))throw Error('Unexpected import '+id);return deps[id]},m,m.exports);return m.exports;
}
const workspace=compile('workspace');deps['./workspace']=workspace;
const {api,OFFLINE_DB_NAME,OFFLINE_BODY_LIMIT}=compile('offline-store');
globalThis.indexedDB=new IDBFactory();
let count=0;
async function test(name,fn){await fn();count++;console.log('PASS '+name);}
const get=()=>api('/api/workspace');
const put=(state,revision,description)=>api('/api/workspace',{method:'PUT',body:JSON.stringify({state,revision,description})});
const capture=(state,label,revision)=>api('/api/runs',{method:'POST',body:JSON.stringify({state,label,revision})});
const review=(state,revision,kind,id)=>api('/api/review',{method:'POST',body:JSON.stringify({state,revision,kind,id})});
const rejects=(promise,status)=>assert.rejects(promise,e=>e.status===status);
const derive=w=>esg.calculateEsg({...deps['./esg-reference.json'],initialData:w.esg},{country:w.country,reportDate:w.inputs.settings.reportDate},w.inputs);
const fresh=()=>{globalThis.indexedDB=new IDBFactory();};
async function raw(store,operation){
  const db=await new Promise((resolve,reject)=>{const r=indexedDB.open(OFFLINE_DB_NAME,1);r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});
  try{return await new Promise((resolve,reject)=>{const tx=db.transaction(store,'readwrite');let result;const r=operation(tx.objectStore(store));if(r)r.onsuccess=()=>result=r.result;tx.oncomplete=()=>resolve(result);tx.onabort=()=>reject(tx.error);});}finally{db.close();}
}

await test('fresh installation seeds once with revision zero and baseline parity',async()=>{
  const a=await get(), b=await get();assert.equal(a.revision,0);assert.deepEqual(a,b);assert.equal(a.runs.length,0);assert.equal(a.events.length,0);
  assert.equal(derive(a.state).summary.eligibleOrGreenUSDm,16.12);assert.equal(derive(a.state).summary.esHoldsOrDeclines,4);
});
await test('returned data is detached from persisted state',async()=>{const a=await get();a.state.name='Mutated returned object';assert.equal((await get()).state.name,'Demonstration bank');});
await test('ordinary save commits state, revision and one event',async()=>{const a=await get();a.state.name='Android verification';const b=await put(a.state,a.revision,'Changed bank name');assert.equal(b.revision,1);const c=await get();assert.equal(c.state.name,a.state.name);assert.equal(c.events.length,1);assert.equal(c.events[0].revision,1);assert.equal(c.events[0].details,'Changed bank name');});
await test('module reload recovers the saved workspace',async()=>{const reloaded=compile('offline-store');assert.equal((await reloaded.api('/api/workspace')).state.name,'Android verification');});
await test('stale revision rejects without changing state or history',async()=>{const a=await get();const before=structuredClone(a);a.state.name='Rejected draft';await rejects(put(a.state,a.revision-1),409);assert.deepEqual(await get(),before);});
await test('simultaneous writes have exactly one CAS winner',async()=>{const a=await get();const x=structuredClone(a.state),y=structuredClone(a.state);x.name='Concurrent X';y.name='Concurrent Y';const r=await Promise.allSettled([put(x,a.revision),put(y,a.revision)]);assert.equal(r.filter(x=>x.status==='fulfilled').length,1);assert.equal(r.find(x=>x.status==='rejected').reason.status,409);const b=await get();assert.equal(b.revision,a.revision+1);assert.equal(b.events.length,a.events.length+1);});
for(const body of ['null','[]','broken'])await test('reject invalid JSON envelope '+body,()=>rejects(api('/api/workspace',{method:'PUT',body}),400));
await test('missing revision is rejected',async()=>{await rejects(put((await get()).state,undefined),400);});
await test('malformed record is 422 and leaves storage unchanged',async()=>{const before=await get(),a=structuredClone(before);a.state.esg.actions=[null];await rejects(put(a.state,a.revision),422);assert.deepEqual(await get(),before);});
await test('imported duplicate IDs and incompatible model are rejected',async()=>{const a=await get(),b=structuredClone(a.state);b.inputs.borrowers[1].id=b.inputs.borrowers[0].id;await rejects(put(b,a.revision),422);b.inputs.borrowers[1].id='B002';b.modelVersion='unknown';await rejects(put(b,a.revision),422);});
await test('request limit measures UTF-8 bytes, not character count',async()=>{const a=await get(),body=JSON.stringify({state:a.state,revision:a.revision,description:'é'.repeat(800000)});assert.ok(body.length<OFFLINE_BODY_LIMIT);assert.ok(new TextEncoder().encode(body).length>OFFLINE_BODY_LIMIT);await rejects(api('/api/workspace',{method:'PUT',body}),413);});
await test('unknown routes and methods stay local',async()=>{await rejects(api('/api/unknown'),404);await rejects(api('/api/workspace',{method:'DELETE'}),405);await rejects(api('https://example.com/api/workspace'),400);});

await test('ordinary save strips forged incoming review promotion',async()=>{
  const a=await get(),t=a.state.esg.criterionAssessments[0],old=t.reviewStamp;t.evidenceRefs=['OFFLINE changed evidence'];t.reviewStamp=esg.createEsgReviewStamp({...deps['./esg-reference.json'],initialData:a.state.esg},'criterion',t.id);t.reviewRecordType='Forged independent assurance';
  const b=await put(a.state,a.revision);assert.equal(b.state.esg.criterionAssessments[0].reviewStamp,old);assert.notEqual(b.state.esg.criterionAssessments[0].reviewRecordType,'Forged independent assurance');assert.equal(derive(b.state).tests[0].validatedResult,'Pending');
});
await test('new record IDs cannot import a review stamp',async()=>{
  const a=await get(),r=structuredClone(a.state.esg.actions[0]);r.id='OFFLINE-NEW-ACTION';r.reviewStamp='forged';r.reviewRecordType='forged';a.state.esg.actions.push(r);const b=await put(a.state,a.revision);assert.equal(b.state.esg.actions.at(-1).reviewStamp,undefined);assert.equal(b.state.esg.actions.at(-1).reviewRecordType,undefined);
});
await test('explicit criterion review promotes changed evidence and saves whole draft',async()=>{const a=await get(),t=a.state.esg.criterionAssessments[0];a.state.name='Criterion review saves whole draft';const b=await review(a.state,a.revision,'criterion',t.id);assert.equal(b.revision,a.revision+1);assert.equal(b.state.name,a.state.name);assert.equal(derive(b.state).tests[0].complete,true);assert.equal(b.state.esg.criterionAssessments[0].reviewRecordType,'User-entered demonstration review');});
await test('invalid kind and unknown record do not write',async()=>{const a=await get();await rejects(review(a.state,a.revision,'constructor','x'),400);await rejects(review(a.state,a.revision,'action','missing'),404);assert.deepEqual(await get(),a);});
await test('Closed alone cannot promote a review',async()=>{const a=await get();a.state.esg.actions[0].requestedState='Closed';await rejects(review(a.state,a.revision,'action',a.state.esg.actions[0].id),422);assert.equal((await get()).revision,a.revision);});
await test('self-verification and future verification are blocked',async()=>{const a=await get(),r=a.state.esg.actions[0];r.requestedState='Closed';r.closureEvidenceRefs=['OFFLINE closure'];r.verifiedOn='2026-12-15';r.verifierId=r.ownerId;await rejects(review(a.state,a.revision,'action',r.id),422);r.verifierId='esg-assurance-2';r.verifiedOn='2030-01-01';await rejects(review(a.state,a.revision,'action',r.id),422);});
await test('valid explicit closure persists and stale review cannot overwrite it',async()=>{const a=await get(),r=a.state.esg.actions[0];r.requestedState='Closed';r.closureEvidenceRefs=['OFFLINE closure'];r.verifiedOn='2026-12-15';r.verifierId='esg-assurance-2';const b=await review(a.state,a.revision,'action',r.id);assert.equal(derive(b.state).actions[0].decisionState,'Closed');assert.equal(derive(b.state).actions[0].externalAssurance,false);await rejects(review(a.state,a.revision,'action',r.id),409);assert.equal((await get()).revision,b.revision);});
await test('readiness review uses evidence and independent owner/approver',async()=>{const a=await get(),d=derive(a.state).readiness.find(x=>x.controlReadiness==='Verified'),r=a.state.esg.readiness.find(x=>x.id===d.id);const approver=r.approverId;r.approverId=r.ownerId;await rejects(review(a.state,a.revision,'readiness',r.id),422);r.approverId=approver;r.evidenceRefs=['OFFLINE readiness'];const b=await review(a.state,a.revision,'readiness',r.id);assert.equal(derive(b.state).readiness.find(x=>x.id===r.id).controlReadiness,'Verified');});
await test('taxonomy cannot promote an overallocated claim',async()=>{const a=await get(),r=derive(a.state).taxonomy.find(x=>x.claimState==='Approved'),t=a.state.esg.taxonomyAssessments.find(x=>x.id===r.id);t.allocatedUSDm=r.drawnUSDm+1;await rejects(review(a.state,a.revision,'taxonomy',t.id),422);assert.equal((await get()).revision,a.revision);});
await test('taxonomy review preserves stored scheme under a different comparison',async()=>{const a=await get(),r=derive(a.state).taxonomy.find(x=>x.claimState==='Approved');a.state.country='Jordan';a.state.taxonomySelection='Auto';const b=await review(a.state,a.revision,'taxonomy',r.id);assert.equal(derive(b.state).taxonomy.find(x=>x.id===r.id).claimState,'Approved');const comparison=esg.calculateEsg({...deps['./esg-reference.json'],initialData:b.state.esg},{country:'Jordan',reportDate:b.state.inputs.settings.reportDate,scheme:'Auto'},b.state.inputs);assert.equal(comparison.summary.eligibleOrGreenUSDm,0);assert.equal(comparison.taxonomy.find(x=>x.id===r.id).isSchemePreview,true);});

let capturedId;
await test('capture accepts a draft/stale context without saving the workspace',async()=>{const a=await get(),state=structuredClone(a.state);state.name='Unsaved scenario';state.scenario='Physical shock';const run=await capture(state,'  Offline stress  ',0);capturedId=run.id;assert.equal(run.label,'Offline stress');const after=await get();assert.equal(after.revision,a.revision);assert.deepEqual(after.state,a.state);assert.equal(after.events.length,a.events.length+1);assert.equal(after.runs[0].id,run.id);});
await test('captured input/summary remains immutable and retrieval changes no state',async()=>{const before=await get(),run=await api('/api/runs?id='+capturedId);assert.equal(run.state.name,'Unsaved scenario');assert.equal(run.summary.scenario,'Physical shock');assert.deepEqual(run.summary,workspace.makeSummary(run.state));run.state.name='Mutating retrieval';run.summary.exposure=0;const again=await api('/api/runs?id='+capturedId);assert.equal(again.state.name,'Unsaved scenario');assert.notEqual(again.summary.exposure,0);assert.deepEqual(await get(),before);});
await test('capture carries stored reviews rather than imported stamps',async()=>{const a=await get(),t=a.state.esg.criterionAssessments[0],old=t.reviewStamp;t.evidenceRefs=['Forged capture evidence'];t.reviewStamp=esg.createEsgReviewStamp({...deps['./esg-reference.json'],initialData:a.state.esg},'criterion',t.id);const r=await capture(a.state,'Forged review test',a.revision),snapshot=await api('/api/runs?id='+r.id);assert.equal(snapshot.state.esg.criterionAssessments[0].reviewStamp,old);assert.equal(derive(snapshot.state).tests[0].validatedResult,'Pending');});
await test('run names, missing ID and unknown ID are validated',async()=>{const a=await get();await rejects(capture(a.state,' ',a.revision),400);await rejects(capture(a.state,'x'.repeat(101),a.revision),400);await rejects(api('/api/runs'),400);await rejects(api('/api/runs?id=unknown'),404);});

await test('event write failure aborts workspace and revision together',async()=>{const before=await get(),state=structuredClone(before.state);state.name='Must not persist';const original=IDBObjectStore.prototype.add;IDBObjectStore.prototype.add=function(...args){if(this.name==='events')throw new DOMException('Injected full disk','QuotaExceededError');return original.apply(this,args)};try{await rejects(put(state,before.revision),503)}finally{IDBObjectStore.prototype.add=original}assert.deepEqual(await get(),before);});
await test('event failure also rolls back a newly inserted run',async()=>{const before=await get(),original=IDBObjectStore.prototype.add;IDBObjectStore.prototype.add=function(...args){if(this.name==='events')throw new DOMException('Injected full disk','QuotaExceededError');return original.apply(this,args)};try{await rejects(capture(before.state,'Must not exist',before.revision),503)}finally{IDBObjectStore.prototype.add=original}assert.deepEqual(await get(),before);});
await test('failed review cannot leak stamp, revision or event',async()=>{const before=await get(),state=structuredClone(before.state),t=state.esg.criterionAssessments[0];t.evidenceRefs=['Attempted review'];const original=IDBObjectStore.prototype.add;IDBObjectStore.prototype.add=function(...args){if(this.name==='events')throw new DOMException('Injected disk failure','QuotaExceededError');return original.apply(this,args)};try{await rejects(review(state,before.revision,'criterion',t.id),503)}finally{IDBObjectStore.prototype.add=original}assert.deepEqual(await get(),before);});
await test('missing IndexedDB is explicit failure, not memory-only success',async()=>{const saved=globalThis.indexedDB;globalThis.indexedDB=undefined;try{await rejects(get(),503)}finally{globalThis.indexedDB=saved}});

fresh();
await test('31 captures remain stored while latest list is limited to 30',async()=>{const a=await get();let first;for(let i=0;i<31;i++){const r=await capture(a.state,'Run '+i,a.revision);first ||= r.id;}const b=await get();assert.equal(b.runs.length,30);assert.equal(b.runs[0].label,'Run 30');assert.equal(b.runs.at(-1).label,'Run 1');assert.equal((await api('/api/runs?id='+first)).label,'Run 0');assert.equal((await raw('runs',s=>s.count())),31);});
await test('history displays latest 50 without losing earlier events',async()=>{let a=await get();for(let i=0;i<21;i++)a=await put(a.state,a.revision,'Save '+i);const b=await get();assert.equal(b.events.length,50);assert.equal(b.events[0].details,'Save 20');assert.equal(await raw('events',s=>s.count()),52);});
await test('sequence keeps newest order even when clock moves backward',async()=>{const a=await get(),OriginalDate=Date;globalThis.Date=class extends OriginalDate{constructor(...args){super(...(args.length?args:['2020-01-01T00:00:00Z']))}};let run;try{run=await capture(a.state,'Clock moved back',a.revision)}finally{globalThis.Date=OriginalDate}const b=await get();assert.equal(b.runs[0].id,run.id);assert.equal(b.events[0].details,'Clock moved back');});
await test('corrupt captured summary fails without erasing data',async()=>{const a=await get(),id=a.runs[0].id,row=await raw('runs',s=>s.get(id));row.summary.exposure=-999;await raw('runs',s=>s.put(row));await rejects(api('/api/runs?id='+id),503);assert.equal((await raw('runs',s=>s.get(id))).summary.exposure,-999);});
fresh();
await test('corrupt saved workspace is never silently replaced by seed',async()=>{await get();const row=await raw('meta',s=>s.get('workspace'));row.state.modelVersion='future-version';await raw('meta',s=>s.put(row));await rejects(get(),503);assert.equal((await raw('meta',s=>s.get('workspace'))).state.modelVersion,'future-version');});
fresh();
await test('missing workspace with surviving run data is not treated as fresh install',async()=>{const a=await get();await capture(a.state,'Surviving snapshot',a.revision);await raw('meta',s=>s.delete('workspace'));await rejects(get(),503);assert.equal(await raw('meta',s=>s.get('workspace')),undefined);assert.equal(await raw('runs',s=>s.count()),1);});
console.log(JSON.stringify({passed:count,status:'All offline adapter checks passed',storage:'IndexedDB with fake-indexeddb',limits:'Newest 30 run / 50 event listings; no silent storage pruning',deviceChecksStillRequired:['Actual WebView persistence after process kill','Android file picker and export/share','OS storage exhaustion','APK upgrade persistence']}));
