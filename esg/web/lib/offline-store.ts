import {newWorkspace, validateWorkspace, makeSummary, esgReference} from './workspace';
import {calculateEsg, createEsgReviewStamp} from './esg.mjs';

// One device-local workspace. No cloud requests, account identity or assurance.
// The 30/50 limits affect lists only: older immutable captures are never pruned.
export const OFFLINE_DB_NAME = 'esg-banking-offline-v1';
export const OFFLINE_BODY_LIMIT = 1_800_000;
const STORE_VERSION = 1;
const KINDS: Record<string,string> = {criterion:'criterionAssessments',action:'actions',taxonomy:'taxonomyAssessments',readiness:'readiness'};
const DERIVED: Record<string,string> = {criterion:'tests',action:'actions',taxonomy:'taxonomy',readiness:'readiness'};

class LocalError extends Error {
  constructor(public status: number, message: string) { super(message); this.name = 'OfflineWorkspaceError'; }
}
const obj = (x:any) => x !== null && typeof x === 'object' && !Array.isArray(x);
const integer = (x:any) => Number.isSafeInteger(x) && x >= 0;
const validTime = (x:any) => typeof x === 'string' && Number.isFinite(Date.parse(x));
const storageFailure = () => new LocalError(503, 'Device storage could not complete this request. Your current draft is unchanged. Export it before closing the app.');
const corruption = () => new LocalError(503, 'Saved device data is unreadable or requires migration. It has not been reset. Keep a backup before attempting recovery.');

function checked(state:any) {
  let issues:string[];
  try { issues = validateWorkspace(state); } catch { throw new LocalError(422, 'Workspace data does not match the required schema.'); }
  if (issues.length) throw new LocalError(422, issues.slice(0,6).join(' · '));
  return state;
}

function parseBody(options:any) {
  const raw = options?.body;
  if (typeof raw !== 'string') throw new LocalError(400, 'Provide a JSON request object.');
  if (new TextEncoder().encode(raw).byteLength > OFFLINE_BODY_LIMIT) throw new LocalError(413, 'Workspace is too large for this edition.');
  let parsed:any;
  try { parsed = JSON.parse(raw); } catch { throw new LocalError(400, 'Unable to read the submitted data.'); }
  if (!obj(parsed)) throw new LocalError(400, 'Provide a JSON request object.');
  return parsed;
}

function carryReviews(state:any, previous:any) {
  for (const name of Object.values(KINDS)) {
    const old = new Map(previous.esg[name].map((r:any) => [r.id,r]));
    for (const record of state.esg[name]) {
      const saved:any = old.get(record.id);
      if (saved?.reviewStamp) {
        record.reviewStamp = saved.reviewStamp;
        record.reviewRecordType = saved.reviewRecordType;
      } else {
        delete record.reviewStamp;
        delete record.reviewRecordType;
      }
    }
  }
  return state;
}

function openDatabase():Promise<IDBDatabase> {
  return new Promise((resolve,reject) => {
    if (!globalThis.indexedDB) return reject(storageFailure());
    let request:IDBOpenDBRequest, settled=false;
    try { request = indexedDB.open(OFFLINE_DB_NAME, STORE_VERSION); } catch { return reject(storageFailure()); }
    request.onupgradeneeded = () => {
      const db=request.result;
      if (!db.objectStoreNames.contains('meta')) db.createObjectStore('meta',{keyPath:'id'});
      if (!db.objectStoreNames.contains('runs')) db.createObjectStore('runs',{keyPath:'id'});
      if (!db.objectStoreNames.contains('events')) db.createObjectStore('events',{keyPath:'id'});
    };
    request.onerror = () => { if (!settled) { settled=true; reject(storageFailure()); } };
    request.onblocked = () => { if (!settled) { settled=true; reject(new LocalError(503,'Device storage is open in another app view. Close that view and retry; saved data has not been reset.')); } };
    request.onsuccess = () => {
      if (settled) { request.result.close(); return; }
      settled=true;
      request.result.onversionchange = () => request.result.close();
      resolve(request.result);
    };
  });
}

function validWorkspaceRow(row:any) {
  if (!obj(row) || row.id !== 'workspace' || row.formatVersion !== STORE_VERSION || !integer(row.revision) || !validTime(row.updatedAt)) throw corruption();
  try { checked(row.state); } catch { throw corruption(); }
}
function validRun(row:any) {
  if (!obj(row) || typeof row.id!=='string' || !row.id || typeof row.label!=='string' || !row.label.trim() || !validTime(row.saved_at) || !integer(row.sequence) || !obj(row.summary)) throw corruption();
  try { checked(row.state); } catch { throw corruption(); }
  // Persisted summaries are checked against the immutable stored input snapshot.
  if (JSON.stringify(makeSummary(row.state)) !== JSON.stringify(row.summary)) throw corruption();
}
function validEvent(row:any) {
  if (!obj(row) || typeof row.id!=='string' || !validTime(row.created_at) || typeof row.action!=='string' || typeof row.details!=='string' || !integer(row.revision) || !integer(row.sequence)) throw corruption();
}
const runMetadata = (r:any) => ({id:r.id,label:r.label,saved_at:r.saved_at,summary:r.summary});
const eventMetadata = (e:any) => ({id:e.id,created_at:e.created_at,action:e.action,details:e.details,revision:e.revision});

/** The Workbench API contract, implemented entirely in device-local IndexedDB.
 * Every operation returns only after its transaction commits; failures abort all
 * writes. Compare-and-swap and review carry-forward use the same transaction.
 */
export async function api(path:string, options?:any):Promise<any> {
  let url:URL;
  try { url=new URL(path,'https://offline.invalid'); } catch { throw new LocalError(400,'Choose a valid local operation.'); }
  if (url.origin!=='https://offline.invalid') throw new LocalError(400,'Only device-local workspace operations are supported.');
  const method=String(options?.method || 'GET').toUpperCase(), route=url.pathname;
  const allowed:Record<string,string[]>={'/api/workspace':['GET','PUT'],'/api/runs':['GET','POST'],'/api/review':['POST']};
  if (!allowed[route]) throw new LocalError(404,'This local operation is not available.');
  if (!allowed[route].includes(method)) throw new LocalError(405,'This method is not supported for the local operation.');
  const input=method==='GET'?null:parseBody(options);
  if (input) {
    checked(input.state);
    if ((route==='/api/workspace'||route==='/api/review') && !integer(input.revision)) throw new LocalError(400,'Missing workspace revision.');
    if (route==='/api/review' && (!Object.hasOwn(KINDS,input.kind)||typeof input.id!=='string')) throw new LocalError(400,'Choose a record and current revision to review.');
    if (route==='/api/runs') {
      input.label=String(input.label||'').trim();
      if (!input.label||input.label.length>100) throw new LocalError(400,'Enter a run name of 1–100 characters.');
    }
  }
  const runId=route==='/api/runs'&&method==='GET'?url.searchParams.get('id'):null;
  if (route==='/api/runs'&&method==='GET'&&!runId) throw new LocalError(400,'Choose a saved run.');
  const db=await openDatabase();
  try {
    return await new Promise((resolve,reject) => {
      let tx:IDBTransaction;
      try { tx=db.transaction(['meta','runs','events'],'readwrite'); } catch { reject(storageFailure()); return; }
      const meta=tx.objectStore('meta'), runs=tx.objectStore('runs'), events=tx.objectStore('events');
      let result:any, failure:any;
      const fail=(error:any) => { failure=error instanceof LocalError?error:storageFailure(); try { tx.abort(); } catch { reject(failure); } };
      const guard=(fn:()=>void) => { try { fn(); } catch(e) { fail(e); } };
      tx.onabort = () => reject(failure||storageFailure());
      tx.onerror = () => { failure ||= storageFailure(); };
      tx.oncomplete = () => resolve(structuredClone(result));
      const get=meta.get('workspace');
      get.onsuccess = () => guard(() => {
        let saved=get.result;
        if (saved===undefined) {
          // A missing workspace in a nonempty database is not a fresh install.
          const countRuns=runs.count(), countEvents=events.count(), counter=meta.get('sequence');
          let remaining=3;
          const init=() => guard(() => {
            if (--remaining) return;
            if (countRuns.result || countEvents.result || counter.result!==undefined) throw corruption();
            saved={id:'workspace',formatVersion:STORE_VERSION,revision:0,updatedAt:new Date().toISOString(),state:newWorkspace()};
            meta.put(saved);
            operate(saved);
          });
          countRuns.onsuccess=init;countEvents.onsuccess=init;counter.onsuccess=init;
        } else { validWorkspaceRow(saved); operate(saved); }
      });

      function operate(saved:any) {
        if (route==='/api/workspace'&&method==='GET') {
          const rr=runs.getAll(), ee=events.getAll();let remaining=2;
          const loaded=() => guard(() => {
            if (--remaining) return;
            for(const row of rr.result) validRun(row);
            for(const row of ee.result) validEvent(row);
            result={state:saved.state,revision:saved.revision,updatedAt:saved.updatedAt,runs:rr.result.sort((a:any,b:any)=>b.sequence-a.sequence).slice(0,30).map(runMetadata),events:ee.result.sort((a:any,b:any)=>b.sequence-a.sequence).slice(0,50).map(eventMetadata)};
          });rr.onsuccess=loaded;ee.onsuccess=loaded;return;
        }
        if (route==='/api/runs'&&method==='GET') {
          const rr=runs.get(runId!);rr.onsuccess=()=>guard(()=>{
            if(!rr.result)throw new LocalError(404,'This run is not available on this device.');
            validRun(rr.result);result={...runMetadata(rr.result),state:rr.result.state};
          });return;
        }
        if ((route==='/api/workspace'||route==='/api/review')&&input.revision!==saved.revision) throw new LocalError(409,'This workspace changed in another app view. Export your draft, then reload the saved version.');
        const state=carryReviews(input.state,saved.state);
        if (route==='/api/review') {
          const name=KINDS[input.kind], record=state.esg[name].find((r:any)=>r.id===input.id);
          if(!record)throw new LocalError(404,'The review record was not found.');
          const ref={...esgReference,initialData:state.esg};
          record.reviewStamp=createEsgReviewStamp(ref,input.kind,input.id);
          record.reviewRecordType='User-entered demonstration review';
          // Omit comparison scheme: review belongs to the stored activity scheme.
          const calculated:any=calculateEsg(ref,{country:state.country,reportDate:state.inputs.settings.reportDate},state.inputs);
          const derived:any=calculated[DERIVED[input.kind]].find((r:any)=>r.id===input.id);
          const accepted=derived&&(input.kind==='criterion'?derived.complete:input.kind==='action'?derived.decisionState==='Closed':input.kind==='readiness'?derived.controlReadiness==='Verified':['Approved','Activity only'].includes(derived.claimState));
          if(!accepted)throw new LocalError(422,'Review cannot be promoted: '+(derived?.reasons?.map((r:any)=>r.message).join(' · ')||'Complete the requested state and evidence first.'));
        }
        const seq=meta.get('sequence');seq.onsuccess=()=>guard(()=>{
          if(seq.result!==undefined&&(!obj(seq.result)||seq.result.id!=='sequence'||!integer(seq.result.value)))throw corruption();
          const sequence=(seq.result?.value||0)+1;
          if(!Number.isSafeInteger(sequence))throw corruption();
          const now=new Date().toISOString();
          meta.put({id:'sequence',value:sequence});
          if(route==='/api/runs') {
            const id=crypto.randomUUID(), summary=makeSummary(state);
            runs.add({id,label:input.label,saved_at:now,state,summary,sequence});
            events.add({id:crypto.randomUUID(),created_at:now,action:'Scenario captured',details:input.label,revision:integer(input.revision)?input.revision:0,sequence});
            result={id,label:input.label,saved_at:now,summary};
          } else {
            const revision=saved.revision+1;
            if(!Number.isSafeInteger(revision))throw corruption();
            meta.put({id:'workspace',formatVersion:STORE_VERSION,state,revision,updatedAt:now});
            events.add({id:crypto.randomUUID(),created_at:now,action:route==='/api/review'?'Demonstration review':'Workspace saved',details:route==='/api/review'?input.kind+' '+input.id:String(input.description||'Workspace updated').slice(0,500),revision,sequence});
            result={state,revision,updatedAt:now};
          }
        });
      }
    });
  } finally { db.close(); }
}
