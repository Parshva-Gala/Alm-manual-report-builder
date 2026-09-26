import {DEFAULT_INPUTS,calculate} from './financial.mjs';
import esgReference from './esg-reference.json';
import emissionsReference from './emissions-reference.json';
import {calculateEsg} from './esg.mjs';
import {calculateEmissions} from './emissions.mjs';
export {esgReference,emissionsReference};
export const MODEL_VERSION='2026.09 · 1.0';
export const SCENARIOS=['Baseline','Orderly transition','Delayed transition','Physical shock'];
export const COUNTRIES=['Egypt','Jordan','UAE','Global'];
export function newWorkspace():any{return structuredClone({modelVersion:MODEL_VERSION,name:'Demonstration bank',scenario:'Delayed transition',country:'Egypt',taxonomySelection:'Auto',actionsEnabled:true,carbonScale:1,physicalScale:1,actionOverride:{yearIndex:3,equity:8,funding:5,fees:.25,approval:'Approved',evidence:'SYN-ACT-3'},inputs:DEFAULT_INPUTS,esg:esgReference.initialData,emissions:emissionsReference.inputs})}
export type Workspace=ReturnType<typeof newWorkspace>;
export function validateWorkspace(w:any){
 const issues:string[]=[];
 if(!w||typeof w!=='object'||!w.inputs||!w.esg||!w.emissions)return ['Workspace data is incomplete.'];
 if(!COUNTRIES.includes(w.country)||!SCENARIOS.includes(w.scenario))issues.push('Choose a supported jurisdiction and scenario.');
 if(!['Auto','Internal','Jordan 2026'].includes(w.taxonomySelection))issues.push('Choose a supported taxonomy context.');
 if(typeof w.name!=='string'||!w.name.trim()||w.name.length>100)issues.push('Bank name must contain 1–100 characters.');
 if(w.modelVersion!==MODEL_VERSION)issues.push('This workspace uses a different model version.');
 if(typeof w.actionsEnabled!=='boolean')issues.push('Management action switch must be true or false.');
 const obj=(x:any)=>x&&typeof x==='object'&&!Array.isArray(x),str=(x:any)=>typeof x==='string',finite=(x:any)=>typeof x==='number'&&Number.isFinite(x);
 if(!obj(w.inputs.settings)||!obj(w.inputs.bank))issues.push('Bank and reporting settings are required.');
 const a=w.actionOverride;if(!obj(a)||!Number.isInteger(a.yearIndex)||a.yearIndex<1||a.yearIndex>5||!['Approved','Proposed','Rejected'].includes(a.approval)||!str(a.evidence)||['equity','funding','fees'].some(k=>!finite(a[k])||a[k]<0))issues.push('Complete the management-action controls with nonnegative amounts and a valid year.');
 for(const key of ['carbonScale','physicalScale'])if(!finite(w[key])||w[key]<0||w[key]>2)issues.push('Stress sensitivities must be between 0 and 2.');
 if(w.inputs.settings?.fxUSDPerUnit?.USD!==1)issues.push('USD-to-USD conversion must equal 1.');
 for(const k of ['cet1Target','tier1Target','totalCapitalTarget','taxRate','dividendPayout'])if(!finite(w.inputs.settings?.[k])||w.inputs.settings[k]<0||w.inputs.settings[k]>1)issues.push(k+' must be a fraction between 0 and 1.');
 if(!Number.isInteger(w.inputs.settings?.reportDate)||w.inputs.settings.reportDate<36526||w.inputs.settings.reportDate>73050)issues.push('Reporting date must be a valid Excel day between 2000 and 2100.');
 for(const k of ['borrowers','facilities','sites'])if(!Array.isArray(w.inputs[k])||w.inputs[k].length>1000)issues.push('Invalid '+k+' register.');
 for(const k of ['actors','esgAssessments','taxonomyAssessments','criterionAssessments','actions','readiness'])if(!Array.isArray(w.esg[k])||w.esg[k].length>10000)issues.push('Invalid '+k+' register.');
 for(const k of ['entities','activities','factors','operations','bankScope3'])if(!Array.isArray(w.emissions[k])||w.emissions[k].length>10000)issues.push('Invalid '+k+' register.');
 if(issues.length)return issues;
 const shapes:any[]=[[w.inputs,['borrowers','facilities','sites']],[w.esg,['actors','esgAssessments','taxonomyAssessments','criterionAssessments','actions','readiness']],[w.emissions,['entities','activities','factors','operations','bankScope3']]];
 for(const [root,names] of shapes)for(const name of names){const ids=new Set();for(const r of root[name]){if(!obj(r)){issues.push(name+' contains an invalid record.');break;}if(name!=='bankScope3'&&(!str(r.id)||!r.id.trim()||ids.has(r.id))){issues.push(name+' requires unique, nonempty record IDs.');break;}ids.add(r.id);for(const [key,value]of Object.entries(r)){if(value!==null&&typeof value==='object'&&!['metrics','evidenceRefs','allocationEvidenceRefs','closureEvidenceRefs','sourceIds'].includes(key)){issues.push(name+' '+r.id+': '+key+' must be a scalar value.');break;}if(['evidenceRefs','allocationEvidenceRefs','closureEvidenceRefs','sourceIds'].includes(key)&&(!Array.isArray(value)||value.some((x:any)=>!str(x)))){issues.push(name+' '+r.id+': '+key+' must be a list of text references.');break;}}}}
 if(issues.length)return [...new Set(issues)];
 for(const k of ['lcrTarget','nsfrTarget'])if(!finite(w.inputs.settings[k])||w.inputs.settings[k]<0||w.inputs.settings[k]>10)issues.push(k+' must be between 0 and 10.');
 for(const b of w.inputs.borrowers)if(['name','country','sector'].some(k=>!str(b[k])||!b[k].trim()))issues.push('Every customer needs a text name, market and sector.');
 for(const b of [...w.emissions.entities,...w.emissions.operations])if(!str(b.name))issues.push('Inventory names must be text.');
 for(const r of w.esg.actors)if(!str(r.displayName))issues.push('Reviewer display names must be text.');
 for(const r of w.esg.taxonomyAssessments)if(!obj(r.metrics)||Object.values(r.metrics).some(v=>v!==null&&!['number','boolean','string'].includes(typeof v)))issues.push('Activity metrics must be scalar values.');
 if(issues.length)return [...new Set(issues)];
 try{const result=calculate(w);if(result.status!=='Ready')issues.push(...result.issues.map((i:any)=>i.id+': '+i.message));}catch{issues.push('Financial data does not match the required schema.');}
 try{const r=calculateEsg({...esgReference,initialData:w.esg},{country:w.country,reportDate:w.inputs.settings.reportDate},w.inputs);if(!r.valid)issues.push(...r.validationErrors.map((x:any)=>x.message));calculateEmissions(w.emissions,w.inputs)}catch{issues.push('Evidence or inventory records do not match the required schema.');}
 return issues;
}
export function isoDate(serial:number){return new Date(Date.UTC(1899,11,30)+serial*86400000).toISOString().slice(0,10)}
export function makeSummary(w:any){const r:any=calculate(w);if(r.status!=='Ready')throw new Error('Resolve financial input issues before saving a scenario run.');const end=r.bank[4];return {modelVersion:MODEL_VERSION,scenario:w.scenario,country:w.country,actionsEnabled:w.actionsEnabled,carbonScale:w.carbonScale,physicalScale:w.physicalScale,exposure:r.opening.grossLoans,ecl:r.reportingECL.total,year5Cet1:end.cet1Ratio,minCet1:Math.min(...r.bank.map((b:any)=>b.cet1Ratio)),cumulativeProfit:r.bank.reduce((s:number,b:any)=>s+b.profitAfterTax,0),rwa:end.totalRWA,year5Cash:end.cash,bank:r.bank.map((b:any)=>({year:b.year,profitAfterTax:b.profitAfterTax,cet1Ratio:b.cet1Ratio,totalRWA:b.totalRWA,lcr:b.lcr,nsfr:b.nsfr}))}}
