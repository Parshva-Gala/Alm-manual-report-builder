/* Pure demonstration decision engine. No I/O, Site dependencies or input mutation.
 * Review fingerprints detect changed records; they are not signatures or assurance.
 */
const DAY=86400000, EPOCH=Date.UTC(1899,11,30), EPS=0.000001;
const own=(o,k)=>Object.prototype.hasOwnProperty.call(o||{},k);
const finite=x=>typeof x==='number'&&Number.isFinite(x);
const text=x=>typeof x==='string'&&x.trim().length>0;
const refs=x=>Array.isArray(x)?x.filter(text).map(s=>s.trim()):[];
const unique=a=>[...new Set(a)];
const sum=a=>a.reduce((n,x)=>n+(finite(x)?x:0),0);
const round=x=>Math.round(x*1e10)/1e10;
const reason=(code,message,details={})=>({code,message,...details});
const date=value=>{
  if(finite(value)&&Number.isInteger(value)&&value>=1&&value<150000){const ms=EPOCH+value*DAY;return{iso:new Date(ms).toISOString().slice(0,10),ms};}
  if(typeof value!=='string'||!/^\d{4}-\d{2}-\d{2}$/.test(value))return null;
  const ms=Date.parse(value+'T00:00:00Z');return Number.isFinite(ms)&&new Date(ms).toISOString().slice(0,10)===value?{iso:value,ms}:null;
};
const canonical=x=>Array.isArray(x)?x.map(canonical):x&&typeof x==='object'?Object.fromEntries(Object.keys(x).sort().map(k=>[k,canonical(x[k])])):x===undefined?null:x;
const fingerprint=x=>{const s=JSON.stringify(canonical(x));let h=14695981039346656037n;for(let i=0;i<s.length;i++){h^=BigInt(s.charCodeAt(i));h=BigInt.asUintN(64,h*1099511628211n);}return'demo-review-v1:'+h.toString(16).padStart(16,'0')+':'+s.length;};
const parts=ref=>({data:ref?.initialData||ref||{},config:ref?.configuration||{},root:ref||{}});
const records=(data,name)=>Array.isArray(data[name])?data[name]:[];
const grouped=(items,key)=>{const m=new Map();for(const x of items){const k=typeof key==='function'?key(x):x[key];if(!m.has(k))m.set(k,[]);m.get(k).push(x);}return m;};
const index=items=>{const groups=grouped(items,'id');return{groups,one:id=>groups.get(id)?.length===1?groups.get(id)[0]:null,valid:id=>text(id)&&groups.get(id)?.length===1};};
const resolveScheme=(selection,country,config)=>{
  if(selection===undefined||selection===null||selection==='Auto')return config.automaticSchemeResolution?.[country]||null;
  if(['Internal','Internal criteria','INTERNAL-DEMO-1'].includes(selection))return'INTERNAL-DEMO-1';
  if(['Jordan2026','Jordan 2026','JO-TAX-2026'].includes(selection))return'JO-TAX-2026';
  return null;
};
const relevantTests=(a,tests)=>tests.filter(t=>t.active!==false&&(t.assessmentId===a.id||(t.family==='MSS'&&t.borrowerId===a.borrowerId)));
const criterionPayload=(t,config)=>({id:t.id,assessmentId:t.assessmentId,borrowerId:t.borrowerId,criterionId:t.criterionId,family:t.family,applicability:t.applicability,enteredResult:t.enteredResult,evidenceRefs:refs(t.evidenceRefs).sort(),naRationale:t.naRationale||null,ownerId:t.ownerId,reviewerId:t.reviewerId,reviewedOn:date(t.reviewedOn)?.iso||null,expiresOn:date(t.expiresOn)?.iso||null,active:t.active!==false,revision:t.revision??1,definition:config.criteria?.find(c=>c.id===t.criterionId)||null});
const actionPayload=a=>({id:a.id,borrowerId:a.borrowerId,domain:a.domain,severity:a.severity,requiredAction:a.requiredAction,ownerId:a.ownerId,dueOn:date(a.dueOn)?.iso||null,openedOn:date(a.openedOn)?.iso||null,requestedState:a.requestedState,closureEvidenceRefs:refs(a.closureEvidenceRefs).sort(),verifierId:a.verifierId,verifiedOn:date(a.verifiedOn)?.iso||null,revision:a.revision??1});
const taxonomyPayload=(a,tests,config,scheme)=>({id:a.id,facilityId:a.facilityId,borrowerId:a.borrowerId,activityCountry:a.activityCountry,activity:a.activity,schemeId:scheme||a.schemeId,performanceYear:a.performanceYear,pathway:a.pathway,metrics:a.metrics||{},allocatedUSDm:a.allocatedUSDm,allocationEvidenceRefs:refs(a.allocationEvidenceRefs).sort(),ownerId:a.ownerId,reviewerId:a.reviewerId,reviewedOn:date(a.reviewedOn)?.iso||null,expiresOn:date(a.expiresOn)?.iso||null,approval:a.approval,revision:a.revision??1,rule:config.activityRules?.find(r=>r.schemeId===(scheme||a.schemeId)&&r.activity===a.activity)||null,cementMilestones:a.activity==='Cement'?config.cementMilestones:[],tests:relevantTests(a,tests).map(t=>criterionPayload(t,config)).sort((x,y)=>String(x.id).localeCompare(String(y.id)))});
const readinessPayload=r=>({id:r.id,country:r.country,area:r.area,requiredOutcome:r.requiredOutcome,ownerId:r.ownerId,approverId:r.approverId,internalTargetOn:date(r.internalTargetOn)?.iso||null,enteredState:r.enteredState,evidenceRefs:refs(r.evidenceRefs).sort(),reviewedOn:date(r.reviewedOn)?.iso||null,expiresOn:date(r.expiresOn)?.iso||null,sourceLocator:r.sourceLocator,authority:r.authority,revision:r.revision??1});

/** Call only after an explicit reviewer action with completed evidence/identity/date fields.
 * This returns a change-detection fingerprint, never proof of real independent assurance.
 */
export function createEsgReviewStamp(referenceInputs,kind,recordId,{scheme}={}){
  const {data,config}=parts(referenceInputs), names={criterion:'criterionAssessments',action:'actions',taxonomy:'taxonomyAssessments',readiness:'readiness'};
  if(!names[kind])return null;
  const xs=records(data,names[kind]).filter(r=>r.id===recordId);if(xs.length!==1)return null;
  const r=xs[0];
  if(kind==='criterion')return fingerprint(criterionPayload(r,config));
  if(kind==='action')return fingerprint(actionPayload(r));
  if(kind==='readiness')return fingerprint(readinessPayload(r));
  const schemeId=scheme?resolveScheme(scheme,r.activityCountry,config):r.schemeId;
  return fingerprint(taxonomyPayload(r,records(data,'criterionAssessments'),config,schemeId));
}

/** Explicitly seed review records for the existing synthetic fixture only.
 * Never use this on user edits: its purpose is fixture migration, not auto-approval.
 */
export function seedSyntheticEsgReviewStamps(referenceInputs){
  const out=structuredClone(referenceInputs),{data}=parts(out);
  for(const [kind,name]of [['criterion','criterionAssessments'],['action','actions'],['readiness','readiness'],['taxonomy','taxonomyAssessments']])for(const r of records(data,name)){
    const eligible=kind==='criterion'||kind==='action'&&r.requestedState==='Closed'||kind==='readiness'&&r.enteredState==='Implemented'||kind==='taxonomy'&&r.approval==='Approved';
    if(eligible){r.reviewStamp=createEsgReviewStamp(out,kind,r.id);r.reviewRecordType='Synthetic demonstration review';}
  }
  return out;
}

function reviewReasons(r,{ownerKey='ownerId',reviewerKey='reviewerId',actors,asOf,evidenceKey='evidenceRefs',requireStamp=null}){
  const a=[];if(!refs(r[evidenceKey]).length)a.push(reason('EVIDENCE_MISSING','Evidence reference is required.'));
  if(!text(r[ownerKey])||!actors.valid(r[ownerKey]))a.push(reason('OWNER_REQUIRED','Select a known owner identity.'));
  if(!text(r[reviewerKey])||!actors.valid(r[reviewerKey]))a.push(reason('REVIEWER_REQUIRED','Select a known reviewer identity.'));
  else if(r[ownerKey]===r[reviewerKey])a.push(reason('SELF_REVIEW','Owner and reviewer must be different identities.'));
  const reviewed=date(r.reviewedOn),expiry=date(r.expiresOn);
  if(!reviewed||!expiry)a.push(reason('REVIEW_DATES_REQUIRED','Valid review and expiry dates are required.'));
  else{
    if(reviewed.ms>asOf.ms)a.push(reason('FUTURE_REVIEW','Review date cannot follow the reporting date.'));
    if(expiry.ms<reviewed.ms)a.push(reason('DATE_ORDER','Expiry date precedes the review date.'));
    if(expiry.ms<asOf.ms)a.push(reason('EVIDENCE_EXPIRED','Evidence has expired.'));
  }
  if(requireStamp&&r.reviewStamp!==requireStamp)a.push(reason('REVIEW_REQUIRED','Record changed or has no explicit current reviewer record.'));
  return a;
}

function technicalResult(a,schemeId,rule,config,year){
  const reasons=[];const result=(status,code,message)=>({status,reasons:code?[reason(code,message)]:reasons,sourceIds:rule?.sourceIds||[],sourceLocator:rule?.sourceLocator||null});
  if(!rule||!schemeId)return result('Criteria required','UNSUPPORTED_ACTIVITY','No implemented rule exists for this activity and scheme.');
  if(!Number.isInteger(a.performanceYear)||a.performanceYear<1900||a.performanceYear>year)return result('Criteria required','PERFORMANCE_YEAR','A valid nonfuture performance year is required.');
  const m=a.metrics||{},jordan=schemeId==='JO-TAX-2026';
  if(!rule.permittedPathways.includes(a.pathway))return result('Criteria required','UNSUPPORTED_PATHWAY','This pathway is not implemented for the selected activity and scheme.');
  const bool=x=>typeof x==='boolean';
  if(a.activity==='Solar PV')return !bool(m.eligiblePvActivityConfirmed)?result('Criteria required','METRIC_MISSING','Confirm whether the activity meets the PV activity definition.'):result(m.eligiblePvActivityConfirmed?'Pass':'Fail');
  if(a.activity==='Cement'){
    if(!finite(m.intensityTco2PerTCementitious)||m.intensityTco2PerTCementitious<0)return result('Criteria required','METRIC_MISSING','A nonnegative cementitious-product intensity is required.');
    if(a.pathway==='Amber')return result(year<2040?'Amber':'Fail',year<2040?null:'SUNSET','Cement Amber requires implementation before 2040.');
    let threshold=rule.referenceTrigger;
    if(jordan){const xs=(config.cementMilestones||[]).filter(x=>x.year===a.performanceYear);if(xs.length!==1)return result('Criteria required','NO_MILESTONE_RULE','No interpolation or floor-year rule is supplied for this measurement year.');threshold=xs[0].maximumIntensity;}
    if(!finite(threshold)||threshold<=0)return result('Criteria required','INVALID_RULE','The intensity threshold requires an approved rule.');
    return{...result(m.intensityTco2PerTCementitious<=threshold?'Pass':'Fail'),threshold,unit:'tCO2 / t cementitious product'};
  }
  if(a.activity==='Building renovation'){
    if(a.pathway==='Whole building')return !bool(m.qualifyingWholeBuildingRoute)?result('Criteria required','METRIC_MISSING','Confirm the whole-building new-building route.'):result(m.qualifyingWholeBuildingRoute?'Pass':'Fail');
    const x=[m.primaryEnergyReduction,m.ghgIntensityReduction],provided=x.filter(v=>v!==null&&v!==undefined);
    if(!provided.length||provided.some(v=>!finite(v)||v<0||v>1))return result('Criteria required','METRIC_MISSING','Provide valid whole-building energy or GHG-intensity reduction fractions.');
    const threshold=rule.referenceTrigger;
    if(!finite(threshold)||threshold<=0||threshold>1)return result('Criteria required','INVALID_RULE','The improvement threshold requires an approved rule.');
    if(Math.max(...provided)<threshold)return result(provided.length===2?'Fail':'Criteria required',provided.length===2?null:'ALTERNATIVE_METRIC_UNKNOWN','The second alternative may satisfy the OR condition and has not been assessed.');
    if(!jordan)return result('Pass');
    if(year===2035)return result('Criteria required','SUNSET_INTERPRETATION','The renovation Amber boundary year 2035 requires approved interpretation.');
    return result(year<2035?'Amber':'Fail');
  }
  if(a.activity==='Road freight'){
    if(!finite(m.directTailpipeGco2PerKm)||m.directTailpipeGco2PerKm<0||!bool(m.dedicatedToFossilFuelTransport))return result('Criteria required','METRIC_MISSING','Tailpipe emissions and fossil-fuel dedication must both be assessed.');
    return result(m.directTailpipeGco2PerKm===0&&!m.dedicatedToFossilFuelTransport?'Pass':'Fail');
  }
  if(a.activity==='New building'){
    if(!bool(m.qualifyingWholeBuildingRoute))return result('Criteria required','METRIC_MISSING','Confirm a qualifying certification or EPC route.');
    if(!m.qualifyingWholeBuildingRoute)return result('Fail');
    if(a.pathway==='Amber'){if(year===2030)return result('Criteria required','SUNSET_INTERPRETATION','The new-building Amber boundary year 2030 requires approved interpretation.');return result(year<2030?'Amber':'Fail');}
    return result('Pass');
  }
  if(a.activity==='Building equipment')return !bool(m.qualifyingEquipmentRoute)?result('Criteria required','METRIC_MISSING','Confirm the eligible equipment / performance route.'):result(m.qualifyingEquipmentRoute?'Pass':'Fail');
  return result('Criteria required','UNSUPPORTED_ACTIVITY','Activity criteria are not implemented.');
}

const allowsNa=(t,a)=>t.criterionId==='CEM-FUELS'&&a.metrics?.usesAlternativeFuels!==true||t.criterionId==='CEM-CCS'&&a.metrics?.usesCcs!==true||t.criterionId==='CEM-PLAN'&&a.pathway!=='Amber'||t.criterionId==='BLD-TIMBER'&&a.metrics?.usesTimber!==true||['BLD-WATER','BLD-WASTE','BLD-LAND'].includes(t.criterionId)&&a.activity==='Building equipment'||t.criterionId==='BLD-LAND'&&a.activity==='Building renovation';

export function calculateEsg(referenceInputs,options={},sharedFinancialInputs){
  const {data,config,root}=parts(referenceInputs),validationErrors=[],warnings=[];
  const rawDate=own(options,'reportDate')?options.reportDate:sharedFinancialInputs?.settings?.reportDate??root.defaults?.asOf??root.fixtureAsOf;
  const asOf=date(rawDate),country=own(options,'country')?options.country:sharedFinancialInputs?.settings?.country??root.defaults?.bankProfile??'Egypt';
  if(!asOf)return{valid:false,context:{country,reportDate:null},validationErrors:[reason('REPORT_DATE_INVALID','Reporting date must be a valid ISO date or Excel date serial.')],borrowers:[],taxonomy:[],tests:[],actions:[],readiness:[],summary:null};
  const countryProfile=config.countryProfiles?.profiles?.find(p=>p.name===country)||null;
  if(!countryProfile)validationErrors.push(reason('COUNTRY_UNSUPPORTED','The bank jurisdiction is not configured.'));
  const forced=own(options,'scheme'),defaultScheme=resolveScheme(options.scheme??'Auto',country,config);
  if(forced&&!defaultScheme)validationErrors.push(reason('SCHEME_UNSUPPORTED','The selected taxonomy scheme is not configured.'));
  const actorIndex=index(records(data,'actors'));
  const borrowerInputs=sharedFinancialInputs&&Array.isArray(sharedFinancialInputs.borrowers)?sharedFinancialInputs.borrowers:records(data,'borrowers');
  const facilityInputs=sharedFinancialInputs&&Array.isArray(sharedFinancialInputs.facilities)?sharedFinancialInputs.facilities:records(data,'facilities');
  const borrowerIndex=index(borrowerInputs),facilityIndex=index(facilityInputs),taxInputs=records(data,'taxonomyAssessments'),taxIndex=index(taxInputs),testInputs=records(data,'criterionAssessments').filter(t=>t.active!==false),testIndex=index(testInputs),esgInputs=records(data,'esgAssessments'),esgIndex=index(esgInputs),actionInputs=records(data,'actions'),actionIndex=index(actionInputs),readyInputs=records(data,'readiness'),readyIndex=index(readyInputs);
  for(const [name,idx]of [['Borrower',borrowerIndex],['Facility',facilityIndex],['Taxonomy assessment',taxIndex],['Criterion record',testIndex],['ESG assessment',esgIndex],['Action',actionIndex],['Readiness control',readyIndex],['Actor',actorIndex]])for(const [id,xs]of idx.groups)if(!text(id)||xs.length!==1)validationErrors.push(reason('DUPLICATE_OR_MISSING_ID',`${name} ID is missing or duplicated.`,{entityId:id??null,entityType:name}));
  const byBorrowerActions=grouped(actionInputs,'borrowerId');
  const actionResults=actionInputs.map(a=>{
    const reasons=[];if(!actionIndex.valid(a.id))reasons.push(reason('ACTION_ID','Action ID is missing or duplicated.'));
    if(!borrowerIndex.valid(a.borrowerId))reasons.push(reason('BORROWER_KEY','Select a unique existing borrower.'));
    for(const [field,values]of [['domain',['E&S','Taxonomy','Climate','Data']],['severity',['Critical','High','Moderate','Low']],['requestedState',['Open','In progress','Closed']]])if(!values.includes(a[field]))reasons.push(reason('ACTION_ENUM',`Invalid action ${field}.`));
    if(!text(a.requiredAction)||!actorIndex.valid(a.ownerId))reasons.push(reason('ACTION_DETAILS','Action and a known owner are required.'));
    const opened=date(a.openedOn),due=date(a.dueOn),verified=date(a.verifiedOn);
    if(!opened||!due||due.ms<opened.ms||opened.ms>asOf.ms)reasons.push(reason('ACTION_DATES','Opening and due dates are missing, future or in the wrong order.'));
    const recordValid=!reasons.length,closureReasons=[];
    if(a.requestedState==='Closed'){
      if(!refs(a.closureEvidenceRefs).length)closureReasons.push(reason('CLOSURE_EVIDENCE','Closure evidence is required.'));
      if(!actorIndex.valid(a.verifierId))closureReasons.push(reason('VERIFIER_REQUIRED','A known verifier identity is required.'));
      else if(a.verifierId===a.ownerId)closureReasons.push(reason('SELF_VERIFICATION','Owner cannot verify their own closure.'));
      if(!verified||!opened||verified.ms<opened.ms||verified.ms>asOf.ms)closureReasons.push(reason('VERIFICATION_DATE','Verification date must be between opening and reporting dates.'));
      if(a.reviewStamp!==fingerprint(actionPayload(a)))closureReasons.push(reason('CLOSURE_REVIEW_REQUIRED','Closure has no explicit current verification record.'));
    }
    const closureVerified=recordValid&&a.requestedState==='Closed'&&!closureReasons.length;
    const decisionState=!recordValid?'Review':closureVerified?'Closed':a.requestedState==='Closed'?'Closure pending':'Open';
    const overdue=decisionState!=='Closed'&&!!due&&due.ms<asOf.ms;
    return{...a,recordValid,closureVerified,decisionState,overdue,daysOverdue:overdue?(asOf.ms-due.ms)/DAY:0,reasons:[...reasons,...closureReasons],verificationKind:closureVerified?'User-entered demonstration review':'Unverified',externalAssurance:false};
  });
  const actionsByBorrower=grouped(actionResults,'borrowerId'),esgByBorrower=grouped(esgInputs,'borrowerId');
  const esgAssessments=esgInputs.map(a=>{
    const b=borrowerIndex.one(a.borrowerId),reasons=[];let weightedScore=null,managementBand='Unrated';
    if(!esgIndex.valid(a.id)||esgByBorrower.get(a.borrowerId)?.length!==1||!b)reasons.push(reason('ESG_KEY','ESG assessment and borrower keys must be unique.'));
    const ws=(config.sectorWeights||[]).filter(w=>w.sector===b?.sector),w=ws[0],scores=[a.environmentScore,a.socialScore,a.governanceScore];
    const validWeights=ws.length===1&&[w.environment,w.social,w.governance].every(x=>finite(x)&&x>=0)&&Math.abs(w.environment+w.social+w.governance-1)<EPS;
    if(!validWeights)reasons.push(reason('SECTOR_WEIGHTS','One valid set of sector weights summing to 1 is required.'));
    if(scores.every(x=>finite(x)&&x>=0&&x<=100)&&validWeights){weightedScore=round(scores[0]*w.environment+scores[1]*w.social+scores[2]*w.governance);managementBand=weightedScore>=80?'Strong':weightedScore>=60?'Developing':'Weak';}
    else reasons.push(reason('PILLAR_SCORES','Three finite pillar scores between 0 and 100 are required.'));
    if(!actorIndex.valid(a.ownerId)||!refs(a.evidenceRefs).length)reasons.push(reason('ESG_EVIDENCE','Assessment owner and evidence are required.'));
    if(!['High','Moderate','Low'].includes(a.inherentRisk)||!['Yes','No'].includes(a.legalProhibition))reasons.push(reason('SCREENING_INCOMPLETE','Complete inherent-risk and legal-prohibition screening.'));
    if(!['Standard','Enhanced'].includes(a.diligenceCompleted)||a.inherentRisk==='High'&&a.diligenceCompleted!=='Enhanced')reasons.push(reason('DILIGENCE_INCOMPLETE','High inherent risk requires Enhanced diligence.'));
    const reviewed=date(a.reviewedOn),expiry=date(a.expiresOn);
    if(!reviewed||!expiry||reviewed.ms>asOf.ms||expiry.ms<reviewed.ms)reasons.push(reason('ESG_DATE_ORDER','Assessment dates are invalid or in the wrong order.'));
    const expired=!!expiry&&expiry.ms<asOf.ms;if(expired)reasons.push(reason('ESG_EXPIRED','Assessment evidence has expired.'));
    const acts=actionsByBorrower.get(a.borrowerId)||[];
    if(acts.some(x=>!x.recordValid))reasons.push(reason('ACTION_RECORD_REVIEW','A linked action record requires review.'));
    const unresolvedEs=acts.filter(x=>x.domain==='E&S'&&x.decisionState!=='Closed'),critical=unresolvedEs.filter(x=>x.severity==='Critical');
    const evidenceState=!reasons.length?'Current':reasons.every(x=>x.code==='ESG_EXPIRED')?'Expired':'Incomplete';
    let esDecision='Clear',gateReasons=[];
    if(a.legalProhibition==='Yes'){esDecision='Decline — prohibited';gateReasons=[reason('LEGAL_PROHIBITION','A legal prohibition cannot be offset by the management score.')];}
    else if(critical.length){esDecision='Hold — critical issue';gateReasons=[reason('CRITICAL_ES_ACTION','Resolve and independently verify the critical E&S action.',{actionIds:critical.map(x=>x.id)})];}
    else if(evidenceState!=='Current'){esDecision='Hold — evidence';gateReasons=reasons;}
    else if(unresolvedEs.length){esDecision='Conditions';gateReasons=[reason('ES_CONDITIONS','Outstanding E&S actions remain.',{actionIds:unresolvedEs.map(x=>x.id)})];}
    return{...a,borrowerName:b?.name??null,country:b?.country??null,sector:b?.sector??null,weightedScore,managementBand,evidenceState,esDecision,unresolvedCriticalEsActions:critical.length,conditionsOpen:unresolvedEs.length,overdueActions:acts.filter(x=>x.overdue).length,reasons,blockingReasons:gateReasons,nextActionIds:unresolvedEs.map(x=>x.id),authority:'Internal bank demonstration policy; not external rating or credit approval'};
  });
  const esgDerivedByBorrower=grouped(esgAssessments,'borrowerId');
  const effectiveSchemes=new Map(taxInputs.map(a=>[a.id,forced?defaultScheme:a.schemeId||defaultScheme]));
  const rulesFor=a=>(config.activityRules||[]).filter(r=>r.schemeId===effectiveSchemes.get(a.id)&&r.activity===a.activity);
  const testPairs=grouped(testInputs,t=>JSON.stringify([t.assessmentId,t.criterionId]));
  const testResults=testInputs.map(t=>{
    const a=taxIndex.one(t.assessmentId),defs=(config.criteria||[]).filter(c=>c.id===t.criterionId),def=defs[0],rules=a?rulesFor(a):[],rule=rules[0],reasons=[];
    if(!testIndex.valid(t.id)||testPairs.get(JSON.stringify([t.assessmentId,t.criterionId]))?.length!==1)reasons.push(reason('DUPLICATE_TEST','Criterion record ID or assessment/criterion pair is duplicated.'));
    if(!a||!borrowerIndex.valid(t.borrowerId)||a.borrowerId!==t.borrowerId)reasons.push(reason('TEST_PARENT_KEY','Criterion must belong to a unique matching borrower and assessment.'));
    const expected=a&&rules.length===1&&rule.requiredCriterionIds.includes(t.criterionId)&&defs.length===1&&def.family===t.family;
    if(!expected)reasons.push(reason('UNEXPECTED_CRITERION','Criterion or family does not match the implemented activity checklist.'));
    const evidenceReasons=reviewReasons(t,{actors:actorIndex,asOf,requireStamp:fingerprint(criterionPayload(t,config))});
    let validatedResult='Pending';
    if(expected&&t.applicability==='Applicable'&&t.enteredResult==='Fail')validatedResult='Fail';
    else if(!reasons.length&&!evidenceReasons.length){
      if(t.applicability==='Applicable'&&t.enteredResult==='Pass')validatedResult='Pass';
      else if(t.applicability==='N/A'&&t.enteredResult==='N/A'&&text(t.naRationale)&&allowsNa(t,a))validatedResult='N/A';
      else reasons.push(reason('APPLICABILITY_OR_RESULT','Resolve the result; N/A needs an allowed condition and rationale.'));
    }
    return{...t,validatedResult,complete:['Pass','N/A'].includes(validatedResult)&&!reasons.length&&!evidenceReasons.length,reasons:[...reasons,...evidenceReasons],sourceLocator:def?.sourceLocator??null,requirement:def?.requirement??null,reviewKind:'User-entered demonstration review',externalAssurance:false};
  });
  const testsByAssessment=grouped(testResults,'assessmentId'),enterpriseMss=[];
  for(const [key,ts]of grouped(testResults.filter(t=>t.family==='MSS'),t=>JSON.stringify([t.borrowerId,t.criterionId]))){
    const statuses=unique(ts.map(t=>t.validatedResult)),failed=ts.some(t=>t.validatedResult==='Fail');
    enterpriseMss.push({borrowerId:ts[0].borrowerId,criterionId:ts[0].criterionId,result:failed?'Fail':ts.every(t=>t.complete)?'Pass':'Pending',conflict:statuses.length>1,recordIds:ts.map(t=>t.id),sourceIds:['S10']});
  }
  const mssByBorrower=grouped(enterpriseMss,'borrowerId'),allocGroups=grouped(taxInputs,'facilityId');
  const balanceFor=f=>{
    if(!f)return null;
    if(sharedFinancialInputs){const fx=sharedFinancialInputs.settings?.fxUSDPerUnit?.[f.currency];return finite(f.drawnLocalm)&&f.drawnLocalm>=0&&finite(fx)&&fx>0?f.drawnLocalm*fx:null;}
    if(finite(f.drawnUSDm)&&f.drawnUSDm>=0)return f.drawnUSDm;
    return f.currency==='USD'&&finite(f.drawnLocalm)&&f.drawnLocalm>=0?f.drawnLocalm:null;
  };
  const taxonomy=taxInputs.map(a=>{
    const schemeId=effectiveSchemes.get(a.id),jordan=schemeId==='JO-TAX-2026',f=facilityIndex.one(a.facilityId),b=borrowerIndex.one(a.borrowerId),rules=rulesFor(a),rule=rules.length===1?rules[0]:null,reasons=[];
    const expected=rule?.requiredCriterionIds||[],schemaValid=!!rule&&unique(expected).length===expected.length&&expected.length===rule.expectedCriterionCount&&expected.length>0;
    const tech=technicalResult(a,schemeId,schemaValid?rule:null,config,+asOf.iso.slice(0,4));
    const ts=testsByAssessment.get(a.id)||[],counts=grouped(ts,'criterionId');
    const missing=expected.filter(id=>counts.get(id)?.length!==1),unexpected=ts.filter(t=>!expected.includes(t.criterionId)),failed=ts.filter(t=>t.validatedResult==='Fail');
    const borrowerMss=mssByBorrower.get(a.borrowerId)||[],mssFail=borrowerMss.filter(x=>x.result==='Fail'),mssConflict=borrowerMss.filter(x=>x.conflict);
    const complete=ts.filter(t=>t.complete&&expected.includes(t.criterionId)).length;
    const safeguards=failed.length||mssFail.length?'Fail':schemaValid&&missing.length===0&&unexpected.length===0&&ts.length===expected.length&&complete===expected.length&&mssConflict.length===0?'Pass':'Pending';
    if(missing.length)reasons.push(reason('CRITERIA_MISSING','Every expected criterion must appear exactly once.',{criterionIds:missing}));
    if(unexpected.length)reasons.push(reason('UNEXPECTED_TESTS','Remove or resolve unexpected criterion records.',{testIds:unexpected.map(t=>t.id)}));
    if(failed.length)reasons.push(reason('CRITERIA_FAILED','An applicable criterion has failed.',{criterionIds:unique(failed.map(t=>t.criterionId))}));
    if(mssFail.length||mssConflict.length)reasons.push(reason('ENTERPRISE_MSS','Resolve borrower-level safeguard failures or conflicting active evidence.',{criterionIds:unique([...mssFail,...mssConflict].map(t=>t.criterionId))}));
    if(safeguards==='Pending')reasons.push(reason('CRITERIA_REVIEW','Complete current independent review of all required tests.'));
    let classification;
    if(tech.status==='Fail'||safeguards==='Fail')classification=jordan?'Red / not aligned':'Not eligible';
    else if(tech.status==='Criteria required'||!schemaValid)classification='Criteria required';
    else if(safeguards!=='Pass')classification='Pending';
    else if(tech.status==='Amber')classification='Amber';
    else classification=jordan?'Green':'Eligible (internal)';
    const recordReasons=[];
    if(!taxIndex.valid(a.id)||!f||!b||f.borrowerId!==a.borrowerId)recordReasons.push(reason('ASSESSMENT_KEY','Assessment, facility and borrower must be unique and linked.'));
    if(!text(a.activityCountry))recordReasons.push(reason('ACTIVITY_COUNTRY','Record the activity country separately from borrower domicile.'));
    // Allocation evidence has its own guard; a zero-allocation activity can still be reviewed.
    const reviewedRecord={...a,evidenceRefs:['Activity review']};
    recordReasons.push(...reviewReasons(reviewedRecord,{actors:actorIndex,asOf,requireStamp:fingerprint(taxonomyPayload(a,testInputs,config,schemeId))}));
    const isSchemePreview=schemeId!==a.schemeId;
    if(isSchemePreview)recordReasons.push(reason('SCHEME_REVIEW_REQUIRED','This is a comparison under a different scheme; stored approval is not transferred.'));
    const recordCurrent=recordReasons.length===0;
    const allocated=a.allocatedUSDm,drawnUSDm=balanceFor(f),group=allocGroups.get(a.facilityId)||[],allocationReasons=[];
    let allocationStatus='Valid';
    if(!f||f.borrowerId!==a.borrowerId){allocationStatus='Facility key';allocationReasons.push(reason('FACILITY_KEY','A unique matching facility is required.'));}
    else if(group.some(x=>!finite(x.allocatedUSDm)||x.allocatedUSDm<0)){allocationStatus='Amount required';allocationReasons.push(reason('ALLOCATION_AMOUNT','Every allocation for this facility must be finite and nonnegative.'));}
    else if(!finite(drawnUSDm)){allocationStatus='Balance required';allocationReasons.push(reason('FACILITY_BALANCE','A valid current facility balance and FX rate are required.'));}
    else if(sum(group.map(x=>x.allocatedUSDm))>drawnUSDm+EPS){allocationStatus='Overallocated';allocationReasons.push(reason('OVERALLOCATED','Combined allocations exceed the facility drawn balance.'));}
    else if(allocated===0)allocationStatus='No allocation';
    else if(!refs(a.allocationEvidenceRefs).length){allocationStatus='Evidence required';allocationReasons.push(reason('ALLOCATION_EVIDENCE','Positive financing allocation requires evidence.'));}
    const critical=(actionsByBorrower.get(a.borrowerId)||[]).filter(x=>['E&S','Taxonomy'].includes(x.domain)&&x.severity==='Critical'&&x.decisionState!=='Closed');
    const legal=(esgDerivedByBorrower.get(a.borrowerId)||[]).some(e=>e.legalProhibition==='Yes');
    const eligible=['Green','Eligible (internal)','Amber'].includes(classification);
    let claimState='Pending';
    if(a.approval==='Rejected')claimState='Rejected';
    else if(['Red / not aligned','Not eligible'].includes(classification))claimState='Not eligible';
    else if(recordCurrent&&a.approval==='Approved'&&safeguards==='Pass'&&!critical.length&&!legal&&eligible)claimState=allocationStatus==='No allocation'?'Activity only':allocationStatus==='Valid'?'Approved':'Pending';
    if(critical.length)reasons.push(reason('CRITICAL_CLAIM_ACTION','An unresolved critical E&S or Taxonomy action blocks this claim.',{actionIds:critical.map(x=>x.id)}));
    if(legal)reasons.push(reason('LEGAL_PROHIBITION','Legal prohibition blocks the financing claim.'));
    if(a.approval!=='Approved')reasons.push(reason('APPROVAL_REQUIRED','The selected assessment requires explicit approval.'));
    const classEligible=['Green','Eligible (internal)'].includes(classification),potential=eligible&&safeguards==='Pass'&&allocationStatus==='Valid'&&!critical.length&&!legal;
    return{...a,schemeId,storedSchemeId:a.schemeId,isSchemePreview,useContext:jordan?(a.activityCountry==='Jordan'?'Jordan taxonomy':'Reference benchmark'):'Internal policy',technicalResult:tech.status,technicalDetail:tech,expectedCriterionCount:expected.length,completeCriterionCount:complete,missingCriterionIds:missing,failedCriterionIds:unique([...failed.map(t=>t.criterionId),...mssFail.map(t=>t.criterionId)]),criterionCompleteness:safeguards,classification,recordCurrent,drawnUSDm,allocationStatus,claimState,eligibleOrGreenUSDm:claimState==='Approved'&&classEligible?allocated:0,amberUSDm:claimState==='Approved'&&classification==='Amber'?allocated:0,potentialEligibleOrGreenUSDm:potential&&classEligible?allocated:0,potentialAmberUSDm:potential&&classification==='Amber'?allocated:0,reasons:[...tech.reasons,...reasons,...recordReasons,...allocationReasons],sourceIds:rule?.sourceIds||[],sourceLocator:rule?.sourceLocator||null,reviewKind:'User-entered demonstration review',externalAssurance:false};
  });
  const readiness=readyInputs.map(r=>{
    const selected=r.country===country||r.country==='Global',reasons=[];
    if(!readyIndex.valid(r.id)||!actorIndex.valid(r.ownerId)||!actorIndex.valid(r.approverId)||r.ownerId===r.approverId||!date(r.internalTargetOn)||!text(r.requiredOutcome)||!text(r.sourceLocator)||!text(r.authority))reasons.push(reason('CONTROL_RECORD','Control IDs, separate owner/approver, source, outcome and internal target must be valid.'));
    let controlReadiness='Outside selection';
    if(selected){
      if(reasons.length)controlReadiness='Review';
      else if(r.enteredState!=='Implemented')controlReadiness='Open';
      else{
        const rr=reviewReasons(r,{actors:actorIndex,asOf,reviewerKey:'approverId',requireStamp:fingerprint(readinessPayload(r))});reasons.push(...rr);
        controlReadiness=!rr.length?'Verified':rr.every(x=>x.code==='EVIDENCE_EXPIRED')?'Expired':'Evidence required';
      }
    }
    const due=date(r.internalTargetOn),overdue=selected&&controlReadiness!=='Verified'&&!!due&&due.ms<asOf.ms;
    return{...r,selected,controlReadiness,requirementApplicability:r.legalApplicability||'Applicability review',overdue,daysOverdue:overdue?(asOf.ms-due.ms)/DAY:0,reasons,completion:controlReadiness==='Verified'?1:0,reviewKind:'User-entered demonstration review',externalAssurance:false};
  });
  const taxonomyByBorrower=grouped(taxonomy,'borrowerId');
  const borrowers=borrowerInputs.map(b=>{
    const es=esgDerivedByBorrower.get(b.id)||[],tax=taxonomyByBorrower.get(b.id)||[],actions=actionsByBorrower.get(b.id)||[];
    return{...b,esg:es.length===1?es[0]:null,esDecision:es.length===1?es[0].esDecision:'Hold — evidence',weightedScore:es.length===1?es[0].weightedScore:null,managementBand:es.length===1?es[0].managementBand:'Unrated',taxonomyAssessmentIds:tax.map(t=>t.id),eligibleOrGreenUSDm:round(sum(tax.map(t=>t.eligibleOrGreenUSDm))),amberUSDm:round(sum(tax.map(t=>t.amberUSDm))),pendingClaims:tax.filter(t=>t.claimState==='Pending').length,openActionIds:actions.filter(a=>a.decisionState!=='Closed').map(a=>a.id),reasons:es.length===1?es[0].blockingReasons:[reason('ESG_ASSESSMENT_REQUIRED','One valid ESG assessment is required.')]};
  });
  const selectedControls=readiness.filter(r=>r.selected),verifiedControls=selectedControls.filter(r=>r.controlReadiness==='Verified');
  const summary={assessedBorrowers:esgAssessments.filter(a=>finite(a.weightedScore)).length,esHoldsOrDeclines:borrowers.filter(b=>['Hold — critical issue','Hold — evidence','Decline — prohibited'].includes(b.esDecision)).length,esConditions:borrowers.filter(b=>b.esDecision==='Conditions').length,esClear:borrowers.filter(b=>b.esDecision==='Clear').length,eligibleOrGreenUSDm:round(sum(taxonomy.map(t=>t.eligibleOrGreenUSDm))),amberUSDm:round(sum(taxonomy.map(t=>t.amberUSDm))),potentialEligibleOrGreenUSDm:round(sum(taxonomy.map(t=>t.potentialEligibleOrGreenUSDm))),potentialAmberUSDm:round(sum(taxonomy.map(t=>t.potentialAmberUSDm))),pendingClaims:taxonomy.filter(t=>t.claimState==='Pending').length,approvedClaims:taxonomy.filter(t=>t.claimState==='Approved').length,openOrUnverifiedActions:actionResults.filter(a=>a.decisionState!=='Closed').length,criticalUnresolvedActions:actionResults.filter(a=>a.severity==='Critical'&&a.decisionState!=='Closed').length,overdueActions:actionResults.filter(a=>a.overdue).length,applicableReadinessControls:selectedControls.length,verifiedReadinessControls:verifiedControls.length,readinessRatio:selectedControls.length?verifiedControls.length/selectedControls.length:null};
  if(taxonomy.some(t=>t.isSchemePreview))warnings.push(reason('SCHEME_COMPARISON','Changed-scheme outputs are a comparison. Only reviewed approvals under the current scheme count as approved claims.'));
  warnings.push(reason('SYNTHETIC_REVIEW_RECORDS','Reviewer identities and evidence are demonstration records, not authenticated independent assurance.'));
  return{valid:validationErrors.length===0,context:{country,reportDate:asOf.iso,defaultSchemeId:defaultScheme,schemeOverride:forced?defaultScheme:null,countryProfile,unit:'USD million',sourceReviewAsOf:config.countryProfiles?.asOf||root.sourceReviewAsOf},summary,borrowers,esgAssessments,taxonomy,tests:testResults,enterpriseMss,actions:actionResults,readiness,requirements:(countryProfile?.requirements||[]).map(r=>({...r,applicability:'Scope confirmation required'})),validationErrors,warnings,unsupportedPaths:root.unsupportedPaths||[]};
}
