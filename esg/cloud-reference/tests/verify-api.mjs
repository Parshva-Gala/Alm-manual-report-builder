import assert from 'node:assert/strict';
const base='http://localhost:5173';
const signIn=await fetch(base+'/signin-with-chatgpt?return_to=/',{redirect:'manual'});
const cookie=signIn.headers.getSetCookie().map(c=>c.split(';')[0]).join('; ');
assert.ok(cookie,'Local demo sign-in must provide a session');
async function request(path,method='GET',body,extra={}){const r=await fetch(base+path,{method,headers:{Cookie:cookie,'Content-Type':'application/json',...extra},body:body===undefined?undefined:typeof body==='string'?body:JSON.stringify(body)});let j;try{j=await r.json()}catch{j={}}return{status:r.status,j}}
let count=0;const pass=name=>{count++;console.log('PASS '+name)};
assert.equal((await fetch(base+'/api/workspace')).status,401);pass('unauthenticated workspace blocked');
const start=await request('/api/workspace');assert.equal(start.status,200);let revision=start.j.revision;const original=start.j.state;
try{
 for(const body of ['null','[]','broken']){assert.equal((await request('/api/workspace','PUT',body)).status,400);pass('invalid request envelope rejected')}
 const bad=structuredClone(original);bad.esg.actions=[null];assert.equal((await request('/api/workspace','PUT',{state:bad,revision})).status,422);pass('malformed imported record rejected');
 assert.equal((await request('/api/workspace','PUT',{state:original,revision},{'Sec-Fetch-Site':'cross-site'})).status,403);pass('cross-site writes blocked');
 const edit=structuredClone(original);edit.name='API verification bank';const saved=await request('/api/workspace','PUT',{state:edit,revision,description:'Local API verification'});assert.equal(saved.status,200);revision=saved.j.revision;pass('authenticated save succeeds');
 const stale=await request('/api/workspace','PUT',{state:original,revision:revision-1});assert.equal(stale.status,409);assert.equal((await request('/api/workspace')).j.state.name,edit.name);pass('stale save cannot overwrite newer workspace');
 const action=structuredClone(edit);const a=action.esg.actions[0];a.requestedState='Closed';const blocked=await request('/api/review','POST',{state:action,revision,kind:'action',id:a.id});assert.equal(blocked.status,422);pass('closure without evidence blocked');
 a.closureEvidenceRefs=['LOCAL-QA-CLOSURE'];a.verifierId='esg-assurance-2';a.verifiedOn='2026-12-15';const reviewed=await request('/api/review','POST',{state:action,revision,kind:'action',id:a.id});assert.equal(reviewed.status,200,JSON.stringify(reviewed.j));revision=reviewed.j.revision;assert.ok(reviewed.j.state.esg.actions[0].reviewStamp);pass('explicit evidence review creates a saved revision');
 const current=reviewed.j.state,forged=structuredClone(current);forged.esg.actions[0].reviewStamp='forged';const guard=await request('/api/workspace','PUT',{state:forged,revision});assert.equal(guard.status,200);revision=guard.j.revision;assert.equal(guard.j.state.esg.actions[0].reviewStamp,current.esg.actions[0].reviewStamp);pass('ordinary save cannot manufacture a review stamp');
 const capture=await request('/api/runs','POST',{state:original,revision,label:'QA · fixed original reference'});assert.equal(capture.status,201);const retrieved=await request('/api/runs?id='+encodeURIComponent(capture.j.id));assert.equal(retrieved.status,200);assert.equal(retrieved.j.state.name,original.name);assert.equal(retrieved.j.summary.year5Cet1,capture.j.summary.year5Cet1);pass('immutable scenario retrieval matches captured inputs');
 assert.equal((await request('/api/runs?id=unknown-run')).status,404);pass('unknown scenario is not disclosed');
}finally{const now=await request('/api/workspace');const restored=await request('/api/workspace','PUT',{state:original,revision:now.j.revision,description:'Restored demonstration after local verification'});assert.equal(restored.status,200);pass('demonstration inputs restored')}
console.log(JSON.stringify({passed:count,status:'All API checks passed',scope:'Local authenticated demo. Cross-owner scoping independently reviewed in SQL.'}));
