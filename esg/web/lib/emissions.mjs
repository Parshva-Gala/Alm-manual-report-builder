/**
 * Pure reference engine for the reviewed core emissions methods. No dependencies,
 * I/O, clock access, hidden state or mutation of caller-owned inputs.
 *
 * calculateEmissions(emissionsInputs, sharedFinancialInputs?)
 * - First argument is emissions_reference.json.inputs (or its parent reference).
 * - Optional second argument is complete state.inputs / financial_defaults.mjs.
 *   Supplied borrower/facility/site arrays REPLACE the corresponding masters.
 *   Its settings.reportDate and settings.fxUSDPerUnit are authoritative.
 * - Exact source fields, nulls and statuses are preserved. Numeric aggregate
 *   subtotals retain the workbook contract; display and complete fields distinguish
 *   no coverage / partial totals from actual zero emissions.
 */
export const EMISSIONS_METHOD_VERSION = 'PCAF-2025-core-lending / GHG-S2-2015 / reference-v1';

const has = (o, k) => o != null && Object.prototype.hasOwnProperty.call(o, k);
const finite = x => typeof x === 'number' && Number.isFinite(x);
const present = x => x !== null && x !== undefined && !(typeof x === 'string' && x.trim() === '');
const nonnegative = x => finite(x) && x >= 0;
const sum = (rows, key) => rows.reduce((a, x) => a + (finite(x[key]) ? x[key] : 0), 0);
const ratio = (a, b) => finite(a) && finite(b) && b > 0 ? a / b : null;
const keyOf = x => present(x) ? String(x) : null;
const array = x => Array.isArray(x) ? x : [];
const corporateClasses = new Set(['Corporate private', 'Corporate listed']);
const reportClasses = new Set(['Corporate private', 'Corporate listed', 'Project finance']);
const originalValueClasses = new Set(['CRE', 'Mortgage', 'Motor']);
const designatedClasses = new Set(['Project finance', 'CRE', 'Mortgage', 'Motor']);
const supportedClasses = new Set([...reportClasses, ...originalValueClasses]);
const methodPages = {'Corporate listed':'40-46','Corporate private':'55-60','Project finance':'67-71',CRE:'77-81',Mortgage:'83-87',Motor:'90-94'};
const categoryNames = ['Purchased goods and services','Capital goods','Fuel and energy related activities','Upstream transportation','Waste generated in operations','Business travel','Employee commuting','Upstream leased assets','Downstream transportation','Processing of sold products','Use of sold products','End of life of sold products','Downstream leased assets','Franchises','Investments'];

/** ISO dates and Excel dates used in financial_defaults are both supported. */
function day(value) {
  if (finite(value)) return value >= 1 && value <= 2958465 ? Math.floor(value) : null;
  if (value instanceof Date) return Number.isFinite(value.valueOf()) ? Math.floor((value.valueOf()-Date.UTC(1899,11,30))/86400000) : null;
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}(?:T.*)?$/.test(value)) return null;
  const datePart=value.slice(0,10), d=new Date(datePart+'T00:00:00Z');
  return Number.isFinite(d.valueOf()) && d.toISOString().slice(0,10)===datePart ? Math.floor((d.valueOf()-Date.UTC(1899,11,30))/86400000) : null;
}
const year = value => {const d=day(value);return d===null?null:new Date(Date.UTC(1899,11,30)+d*86400000).getUTCFullYear();};
const dateValid = x => day(x) !== null;
const pick = (record, names) => {for(const k of names)if(has(record,k))return record[k];return null;};
function index(rows, key='id') {
  const groups=new Map();
  for(const r of rows){const k=keyOf(r[key]);if(!groups.has(k))groups.set(k,[]);groups.get(k).push(r);}
  return {groups,count:k=>(groups.get(keyOf(k))||[]).length,get:k=>{const a=groups.get(keyOf(k));return a?.length===1?a[0]:null;}};
}
function normalBorrower(b) {
  return {...b,id:pick(b,['id','borrowerId']),listing:pick(b,['listing','listed']),financialDate:pick(b,['financialDate','reportDate']),revenueUSDm:pick(b,['revenueUSDm','revenue']),equityUSDm:pick(b,['equityUSDm','equity']),debtUSDm:pick(b,['debtUSDm','debt']),evicUSDm:pick(b,['evicUSDm','evic']),evidence:pick(b,['evidence','financialEvidence']),expiry:pick(b,['expiry','financialExpiry'])};
}
function normalFacility(f) {
  return {...f,id:pick(f,['id','facilityId']),emissionsEntityId:pick(f,['emissionsEntityId','entityId']),accountingCCF:pick(f,['accountingCCF','ccf']),originalCollateralUSDm:pick(f,['originalCollateralUSDm','origCollateral']),currentCollateralUSDm:pick(f,['currentCollateralUSDm','currentCollateral'])};
}

function aggregate(rows) {
  const drawnUSDm=sum(rows,'drawnUSDm'), coveredS1S2USDm=sum(rows,'coveredS1S2USDm'), coveredS3USDm=sum(rows,'coveredS3USDm');
  const coveredS1USDm=sum(rows,'coveredS1USDm'), coveredS2USDm=sum(rows,'coveredS2USDm');
  const unknownDrawnCount=rows.filter(x=>!nonnegative(x.drawnUSDm)).length;
  const joint=rows.filter(x=>x.methodStatus==='Ready'&&x.s1Status==='Ready'&&x.s2Status==='Ready');
  const financedS1S2=sum(joint,'s1')+sum(joint,'s2'), financedS1=sum(rows,'s1'), financedS2=sum(rows,'s2'), financedS3=sum(rows,'s3');
  const coverage=x=>unknownDrawnCount?null:ratio(x,drawnUSDm);
  const display=(value,covered,availableCount)=>availableCount>0?value:null;
  const a={facilityCount:rows.length,drawnUSDm,financedS1,financedS2,financedS1S2,financedS3,coveredS1S2USDm,coveredS3USDm,coveredS1USDm,coveredS2USDm,
    coverageS1S2:coverage(coveredS1S2USDm),coverageS3:coverage(coveredS3USDm),coverageS1:coverage(coveredS1USDm),coverageS2:coverage(coveredS2USDm),
    dqS1S2:ratio(sum(rows,'weightedDQ12'),coveredS1S2USDm),dqS1:ratio(sum(rows,'weightedDQ1'),coveredS1USDm),dqS2:ratio(sum(rows,'weightedDQ2'),coveredS2USDm),dqS3:ratio(sum(rows,'weightedDQ3'),coveredS3USDm),
    economicIntensityS1S2:ratio(financedS1S2,coveredS1S2USDm),methodPending:rows.filter(x=>x.methodStatus!=='Ready').length,overlapReviewUSDm:sum(rows.filter(x=>x.overlap==='Review overlap'),'drawnUSDm'),unknownDrawnCount,
    displayFinancedS1S2:display(financedS1S2,coveredS1S2USDm,joint.length),displayFinancedS1:display(financedS1,coveredS1USDm,rows.filter(x=>finite(x.s1)).length),displayFinancedS2:display(financedS2,coveredS2USDm,rows.filter(x=>finite(x.s2)).length),displayFinancedS3:display(financedS3,coveredS3USDm,rows.filter(x=>finite(x.s3)).length),
    uncoveredS1S2USDm:unknownDrawnCount?null:drawnUSDm-coveredS1S2USDm,uncoveredS3USDm:unknownDrawnCount?null:drawnUSDm-coveredS3USDm};
  a.coverageStatusS1S2=unknownDrawnCount?'Exposure amount unknown':rows.length===0?'No exposure':a.coverageS1S2===1?'Complete':a.displayFinancedS1S2===null?'Not covered':'Partial';
  a.coverageStatusS3=unknownDrawnCount?'Exposure amount unknown':rows.length===0?'No exposure':a.coverageS3===1?'Complete':a.displayFinancedS3===null?'Not covered':'Partial';
  return a;
}

export function calculateEmissions(input, sharedFinancialInputs={}) {
  const source=input?.inputs && !input.entities ? input.inputs : (input||{});
  const shared=sharedFinancialInputs||{}, sharedSettings=shared.settings||{};
  const borrowers=array(has(shared,'borrowers')?shared.borrowers:source.borrowers).map(normalBorrower);
  const rawFacilities=array(has(shared,'facilities')?shared.facilities:source.facilities).map(normalFacility);
  const sites=array(has(shared,'sites')?shared.sites:source.sites).map(x=>({...x,id:pick(x,['id','siteId'])}));
  const rawEntities=array(source.entities).map(x=>({...x}));
  const rawActivities=array(source.activities).map(x=>({...x}));
  const rawFactors=array(source.factors).map(x=>({...x}));
  const rawOperations=array(source.operations).map(x=>({...x}));
  const rawScope3=array(source.bankScope3).map(x=>({...x}));
  const reviewDate=has(sharedSettings,'reportDate')?sharedSettings.reportDate:has(shared,'reviewDate')?shared.reviewDate:has(source.settings,'reportDate')?source.settings.reportDate:source.reviewDate;
  const reviewDay=day(reviewDate), reviewYear=year(reviewDate), controlBasis=source.controlBasis;
  const borrowerIndex=index(borrowers), rawEntityIndex=index(rawEntities), siteIndex=index(sites), factorIds=index(rawFactors), facilityIds=index(rawFacilities), activityIds=index(rawActivities), operationIds=index(rawOperations);
  const fxRows=has(sharedSettings,'fxUSDPerUnit')?Object.entries(sharedSettings.fxUSDPerUnit||{}).map(([currency,usdPerUnit])=>({currency,usdPerUnit})):array(has(shared,'fx')?shared.fx:source.fx);
  const fxIndex=index(fxRows,'currency');
  const gaps=[];
  const gap=(module,recordId,status,more={})=>gaps.push({module,recordId:recordId??null,status,...more});
  if(reviewDay===null)gap('settings','reportDate','Reporting date required',{blocking:true});
  for(const [key,rows]of borrowerIndex.groups)if(!key||rows.length!==1)gap('borrowers',key,!key?'Borrower ID required':'Duplicate borrower ID',{blocking:true});

  const factors=rawFactors.map(x=>{
    let status='Ready';
    if(!present(x.id))status='Factor ID required';
    else if(factorIds.count(x.id)!==1)status='Duplicate factor ID';
    else if(![x.activity,x.country,x.unit,x.scope,x.basis,x.source,x.version,x.dataBasis].every(present)||x.gasBasis!=='CO2e'||!nonnegative(x.kgCO2ePerUnit)||!finite(x.year)||!dateValid(x.expiry))status='Missing factor input';
    else if(reviewDay===null)status='Reporting date required';
    else if(day(x.expiry)<reviewDay)status='Expired';
    if(status!=='Ready')gap('factors',x.id,status,{blocking:true});
    return {...x,status};
  });
  const factorIndex=index(factors);
  const annualKeys=index(rawActivities.map(x=>({...x,annualKey:JSON.stringify([x.entityId,x.siteId,x.scope,x.activity,x.year])})),'annualKey');
  const activities=rawActivities.map(x=>{
    const e=rawEntityIndex.get(x.entityId), b=e?borrowerIndex.get(e.borrowerId):null, fc=factorIndex.get(x.factorId), st=siteIndex.get(x.siteId);
    const country=b?.country??null, annualKey=JSON.stringify([x.entityId,x.siteId,x.scope,x.activity,x.year]);
    let status='Ready';
    if(!present(x.id))status='Activity ID required';
    else if(activityIds.count(x.id)!==1||annualKeys.count(annualKey)!==1)status='Duplicate annual activity';
    else if(!e)status='Entity required';
    else if(!present(x.siteId)||!present(x.evidence)||!nonnegative(x.quantity)||!finite(x.year)||!dateValid(x.expiry))status='Missing activity input';
    else if(x.siteId!==x.entityId&&(!st||st.borrowerId!==e.borrowerId||st.country!==country))status='Site / entity mismatch';
    else if(reviewDay===null)status='Reporting date required';
    else if(day(x.expiry)<reviewDay)status='Expired';
    else if(!fc||fc.status!=='Ready')status='Factor required';
    else if(x.activity!==fc.activity||x.unit!==fc.unit||x.scope!==fc.scope||x.year!==fc.year||(fc.country!=='Global'&&fc.country!==country)||x.year!==year(e.inventoryDate))status='Factor / boundary mismatch';
    else if(x.scope==='S2'&&fc.basis!=='Location-based')status='Use location factor';
    else if(x.dataClass!=='Actual')status='Estimation method review';
    if(status!=='Ready')gap('activities',x.id,status,{entityId:x.entityId,borrowerId:e?.borrowerId??null,scope:x.scope,blocking:true});
    return {...x,tCO2e:status==='Ready'?x.quantity*fc.kgCO2ePerUnit/1000:null,status,factorStatus:fc?.status??'Factor required',entityCountry:country};
  });
  const activityGroups=new Map();
  for(const a of activities){const k=JSON.stringify([a.entityId,a.scope]);if(!activityGroups.has(k))activityGroups.set(k,[]);activityGroups.get(k).push(a);}

  const entities=rawEntities.map(x=>{
    const b=borrowerIndex.get(x.borrowerId), corporate=corporateClasses.has(x.assetClass), original=originalValueClasses.has(x.assetClass);
    const selectedEquityUSDm=corporate?(finite(b?.equityUSDm)?b.equityUSDm:null):x.assetClass==='Project finance'?(finite(x.projectEquityUSDm)?x.projectEquityUSDm:null):null;
    const selectedDebtUSDm=corporate?(finite(b?.debtUSDm)?b.debtUSDm:null):x.assetClass==='Project finance'?(finite(x.projectDebtUSDm)?x.projectDebtUSDm:null):null;
    const selectedEVICUSDm=x.assetClass==='Corporate listed'&&finite(b?.evicUSDm)?b.evicUSDm:null;
    const denominatorUSDm=x.assetClass==='Corporate listed'?selectedEVICUSDm:['Corporate private','Project finance'].includes(x.assetClass)?(finite(selectedEquityUSDm)&&finite(selectedDebtUSDm)?Math.max(selectedEquityUSDm,0)+selectedDebtUSDm:null):original&&finite(x.originalAssetValueUSDm)?x.originalAssetValueUSDm:null;
    const financeEvidence=corporate?b?.evidence:x.assetValueEvidence, valueDate=corporate?b?.financialDate:x.assetValueDate;
    let financeStatus='Ready';
    if(!present(x.id))financeStatus='Entity ID required';
    else if(rawEntityIndex.count(x.id)!==1)financeStatus='Duplicate entity ID';
    else if(!b)financeStatus='Borrower required';
    else if(!supportedClasses.has(x.assetClass))financeStatus='Method required';
    else if((x.assetClass==='Corporate private'&&b.listing!=='Private')||(x.assetClass==='Corporate listed'&&b.listing!=='Listed'))financeStatus='Listing / class mismatch';
    else if(!finite(denominatorUSDm)||denominatorUSDm<=0||!present(financeEvidence)||!dateValid(valueDate))financeStatus='Value evidence required';
    else if(reviewDay===null)financeStatus='Reporting date required';
    else if(corporate&&(!dateValid(b.expiry)||day(b.expiry)<reviewDay))financeStatus='Expired financial evidence';
    else if(day(valueDate)>reviewDay)financeStatus='Future value date';
    else if((['Corporate private','Project finance'].includes(x.assetClass)&&(selectedDebtUSDm<0||x.valueBasis!=='Current financials'))||(x.assetClass==='Corporate listed'&&x.valueBasis!=='Current financials')||(original&&!['Origination','Fixed fallback'].includes(x.valueBasis)))financeStatus='Value basis review';
    else if(x.valueBasis==='Current financials'&&year(valueDate)!==reviewYear)financeStatus='Stale financial period';
    const out={...x,country:b?.country??null,selectedEquityUSDm,selectedDebtUSDm,selectedEVICUSDm,denominatorUSDm,financeEvidence:financeEvidence??null,valueDate:valueDate??null,financeStatus};
    if(financeStatus!=='Ready')gap('entities',x.id,financeStatus,{borrowerId:x.borrowerId,blocking:true});
    for(const scope of ['s1','s2','s3']){
      const method=x[scope+'Method'], reported=x[scope+'Reported'], acts=activityGroups.get(JSON.stringify([x.id,scope.toUpperCase()]))||[];
      let state='Ready';
      if(method==='Not assessed')state='Not assessed';
      else if(!present(x.emissionsEvidence)||!dateValid(x.inventoryDate)||!dateValid(x.expiry))state='Evidence required';
      else if(reviewDay===null)state='Reporting date required';
      else if(day(x.expiry)<reviewDay||day(x.inventoryDate)>reviewDay||year(x.inventoryDate)!==reviewYear)state='Period / expiry review';
      else if(x[scope+'Complete']!=='Yes')state='Boundary incomplete';
      else if(scope==='s2'&&(!['Location-based','Market-based'].includes(x.s2Basis)||(method==='Activity'&&x.s2Basis!=='Location-based')))state='S2 basis mismatch';
      else if(scope==='s3'&&!present(x.s3ScreeningEvidence))state='S3 screening required';
      else if(['Reported verified','Reported unverified'].includes(method)){
        if(!reportClasses.has(x.assetClass))state='Asset method required';
        else if(!nonnegative(reported))state='Reported value required';
      }else if(scope==='s3')state='Scope 3 method required';
      else if(method==='Activity'){
        if(!acts.length||acts.some(a=>a.status!=='Ready'))state='Activity incomplete';
      }else state='Method required';
      out[scope+'Status']=state;
      out[scope]=state==='Ready'?(method==='Activity'?sum(acts,'tCO2e'):reported):null;
      out[scope+'DQ']=state==='Ready'?(method==='Reported verified'?1:method==='Reported unverified'?2:x.assetClass==='Motor'?1:2):null;
      if(state!=='Ready')gap('entityScope',x.id,state,{borrowerId:x.borrowerId,scope,blocking:true,requirement:scope!=='s3'?'Core scope':corporate?'Corporate S3 conformance':x.assetClass==='Project finance'?'S3 relevance assessment':'Outside core asset S3 method'});
    }
    return out;
  });
  const entityIndex=index(entities);
  const draws=rawFacilities.map(x=>{const fx=fxIndex.get(x.currency)?.usdPerUnit;return {...x,usdPerUnit:finite(fx)&&fx>0?fx:null,drawnUSDm:nonnegative(x.drawnLocalm)&&finite(fx)&&fx>0?x.drawnLocalm*fx:null};});
  const entityDraws=new Map();for(const x of draws)if(finite(x.drawnUSDm))entityDraws.set(x.emissionsEntityId,(entityDraws.get(x.emissionsEntityId)||0)+x.drawnUSDm);
  const facilities=draws.map(x=>{
    const e=entityIndex.get(x.emissionsEntityId), entityDrawnUSDm=entityDraws.get(x.emissionsEntityId)||0;
    let methodStatus='Ready';
    if(!present(x.id))methodStatus='Facility ID required';
    else if(facilityIds.count(x.id)!==1)methodStatus='Duplicate facility ID';
    else if(!nonnegative(x.drawnUSDm))methodStatus='Drawn value required';
    else if(!e)methodStatus='Entity required';
    else if(x.assetClass!==e.assetClass||x.borrowerId!==e.borrowerId)methodStatus='Entity / class mismatch';
    else if(e.financeStatus!=='Ready')methodStatus=e.financeStatus;
    else if(designatedClasses.has(x.assetClass)&&x.purpose!=='Designated asset')methodStatus='Purpose review';
    else if(entityDrawnUSDm>e.denominatorUSDm)methodStatus='Exposure exceeds value';
    const attribution=methodStatus==='Ready'?x.drawnUSDm/e.denominatorUSDm:null;
    const out={facilityId:x.id,borrowerId:x.borrowerId,assetClass:x.assetClass,entityId:x.emissionsEntityId,drawnUSDm:x.drawnUSDm,denominatorUSDm:e?.denominatorUSDm??null,attribution,methodStatus,expiry:e?.expiry??null,methodReference:methodPages[x.assetClass]?`PCAF 2025 pp${methodPages[x.assetClass]}`:'Method required',entityDrawnUSDm,s2Basis:e?.s2Basis??null,overlap:e?.overlap??null};
    for(const s of ['s1','s2','s3']){out[s+'Status']=e?.[s+'Status']??'Entity required';out[s+'DQ']=e?.[s+'DQ']??null;out[s]=methodStatus==='Ready'&&out[s+'Status']==='Ready'?attribution*e[s]:null;}
    out.coveredS1USDm=methodStatus==='Ready'&&out.s1Status==='Ready'?x.drawnUSDm:0;
    out.coveredS2USDm=methodStatus==='Ready'&&out.s2Status==='Ready'?x.drawnUSDm:0;
    out.coveredS3USDm=methodStatus==='Ready'&&out.s3Status==='Ready'?x.drawnUSDm:0;
    out.coveredS1S2USDm=methodStatus==='Ready'&&out.s1Status==='Ready'&&out.s2Status==='Ready'?x.drawnUSDm:0;
    out.weightedDQ1=out.coveredS1USDm>0?out.coveredS1USDm*out.s1DQ:0;
    out.weightedDQ2=out.coveredS2USDm>0?out.coveredS2USDm*out.s2DQ:0;
    out.weightedDQ3=out.coveredS3USDm>0?out.coveredS3USDm*out.s3DQ:0;
    out.weightedDQ12=out.coveredS1S2USDm>0?out.coveredS1S2USDm*Math.max(out.s1DQ,out.s2DQ):0;
    if(methodStatus!=='Ready')gap('facilities',x.id,methodStatus,{borrowerId:x.borrowerId,entityId:x.emissionsEntityId,exposureUSDm:x.drawnUSDm,blocking:true});
    if(out.overlap==='Review overlap')gap('overlap',x.id,'Review overlap',{borrowerId:x.borrowerId,entityId:x.emissionsEntityId,exposureUSDm:x.drawnUSDm,blocking:false});
    return out;
  });
  const portfolio=aggregate(facilities);
  const borrowerIds=[...new Set([...borrowers.map(x=>x.id),...facilities.map(x=>x.borrowerId)])];
  const customers=borrowerIds.map(id=>{const b=borrowerIndex.get(id),a=aggregate(facilities.filter(x=>x.borrowerId===id));return {id,name:b?.name??'Unknown borrower',country:b?.country??'Unknown',sector:b?.sector??'Unknown',...a,portfolioEmissionsShare:ratio(a.financedS1S2,portfolio.financedS1S2),portfolioExposureShare:ratio(a.drawnUSDm,portfolio.drawnUSDm)};});
  const groups=field=>[...new Set(customers.map(x=>x[field]))].map(group=>({group,...aggregate(facilities.filter(x=>(borrowerIndex.get(x.borrowerId)?.[field]??'Unknown')===group))}));

  const controlled=rawOperations.filter(x=>x.withinControl==='Yes');
  let dualReporting='Unknown';
  if(controlled.some(x=>x.marketInformation==='No'&&finite(x.contractKWh)&&x.contractKWh>0))dualReporting='Unknown';
  else if(controlled.some(x=>x.marketInformation==='Yes'))dualReporting='Yes';
  else if(controlled.length&&controlled.every(x=>x.marketInformation==='No'))dualReporting='No';
  const operationalFactor=(op,id,unit,scope,basis)=>{const fc=factorIndex.get(id);return fc?.status==='Ready'&&fc.unit===unit&&fc.scope===scope&&fc.basis===basis&&fc.year===op.year&&(fc.country==='Global'||fc.country===op.country);};
  const operations=rawOperations.map(x=>{
    const statusFor=scope=>{
      if(!['Operational control','Financial control'].includes(controlBasis))return 'Control basis required';
      if(x.withinControl==='No')return 'Outside control';
      if(x.withinControl!=='Yes')return 'Control review';
      if(!present(x.id)||operationIds.count(x.id)!==1)return 'Duplicate operation ID';
      if(reviewDay===null||!present(x.activityEvidence)||!dateValid(x.expiry)||day(x.expiry)<reviewDay||!finite(x.year)||x.year!==reviewYear)return 'Evidence / period review';
      if(scope==='S1'){
        if(!nonnegative(x.dieselLitres)||x.scope1BoundaryComplete!=='Yes'||!nonnegative(x.otherScope1Tonnes)||!present(x.otherScope1Evidence))return 'Activity required';
        return operationalFactor(x,x.fuelFactorId,'litre','S1','Combustion')?'Ready':'Factor / unit review';
      }
      if(!nonnegative(x.electricityKWh))return 'Activity required';
      return operationalFactor(x,x.locationFactorId,'kWh','S2','Location-based')?'Ready':'Factor / unit review';
    };
    const scope1Status=statusFor('S1'),scope2LBStatus=statusFor('S2');
    let scope2MBStatus='Ready';
    if(x.withinControl==='No')scope2MBStatus='Outside control';
    else if(x.withinControl!=='Yes')scope2MBStatus='Control review';
    else if(dualReporting==='No')scope2MBStatus='Not required';
    else if(dualReporting!=='Yes')scope2MBStatus='Applicability review';
    else if(scope2LBStatus!=='Ready')scope2MBStatus=scope2LBStatus;
    else if(!nonnegative(x.contractKWh)||x.contractKWh>x.electricityKWh)scope2MBStatus='Contract quantity review';
    else if(x.contractKWh>0&&(x.eightCriteriaReview!=='Pass'||!present(x.retirementEvidence)||!operationalFactor(x,x.contractFactorId,'kWh','S2','Market-based')))scope2MBStatus='Contract criteria review';
    else if(x.contractKWh===x.electricityKWh)scope2MBStatus='Ready';
    else if(x.residualAvailable==='No')scope2MBStatus='Grid fallback';
    else if(x.residualAvailable==='Yes')scope2MBStatus=operationalFactor(x,x.residualFactorId,'kWh','S2','Residual mix')?'Ready':'Residual factor review';
    else scope2MBStatus='Residual availability review';
    const scope1=scope1Status==='Ready'?x.dieselLitres*factorIndex.get(x.fuelFactorId).kgCO2ePerUnit/1000+x.otherScope1Tonnes:null;
    const scope2LB=scope2LBStatus==='Ready'?x.electricityKWh*factorIndex.get(x.locationFactorId).kgCO2ePerUnit/1000:null;
    let scope2MB=null,contractEmissions=null,contractLocationCounterfactual=null,residualEmissions=null,residualLocationCounterfactual=null;
    if(['Ready','Grid fallback'].includes(scope2MBStatus)){
      contractEmissions=x.contractKWh>0?x.contractKWh*factorIndex.get(x.contractFactorId).kgCO2ePerUnit/1000:0;
      contractLocationCounterfactual=x.contractKWh*factorIndex.get(x.locationFactorId).kgCO2ePerUnit/1000;
      const remainder=x.electricityKWh-x.contractKWh;
      residualEmissions=remainder>0?remainder*factorIndex.get(x.residualAvailable==='No'?x.locationFactorId:x.residualFactorId).kgCO2ePerUnit/1000:0;
      residualLocationCounterfactual=remainder*factorIndex.get(x.locationFactorId).kgCO2ePerUnit/1000;
      scope2MB=contractEmissions+residualEmissions;
    }
    const out={...x,scope1,scope2LB,scope2MB,scope1Status,scope2LBStatus,scope2MBStatus,contractEmissions,contractLocationCounterfactual,residualEmissions,residualLocationCounterfactual};
    for(const[scope,state]of[['S1',scope1Status],['S2 LB',scope2LBStatus],['S2 MB',scope2MBStatus]])if(!['Ready','Outside control','Grid fallback','Not required'].includes(state))gap('operations',x.id,state,{scope,blocking:true});
    return out;
  });
  const categoryIndex=index(rawScope3,'category');
  const bankScope3=categoryNames.map((name,i)=>{
    const category=i+1, count=categoryIndex.count(category), row=categoryIndex.get(category),x=row||{category,name,disposition:'Not assessed'};
    const grossInput=category===15?portfolio.financedS1S2:x.grossInput;
    let status='Ready';
    if(count>1)status='Duplicate category';
    else if(!row||x.disposition==='Not assessed')status='Not assessed';
    else if(reviewDay===null||!present(x.evidence)||!present(x.rationale)||!dateValid(x.expiry)||day(x.expiry)<reviewDay||x.year!==reviewYear||x.overlapReview!=='Pass'||x.boundaryReview!=='Pass')status='Evidence / boundary review';
    else if(x.disposition==='Included'){
      if(!present(x.method)||!nonnegative(grossInput))status='Amount / method required';
      else if(category===15)status='Modeled lending only';
    }else if(['Not applicable','Excluded justified'].includes(x.disposition))status=x.disposition;
    else status='Disposition required';
    if(!['Ready','Not applicable','Excluded justified','Modeled lending only'].includes(status))gap('bankScope3',category,status,{scope:'S3',blocking:true});
    return {...x,category,name:x.name||name,grossInput:grossInput??null,selectedGross:['Ready','Modeled lending only'].includes(status)?grossInput:null,status,investeeScope3:category===15?portfolio.financedS3:null};
  });
  for(const x of rawScope3)if(!Number.isInteger(x.category)||x.category<1||x.category>15)gap('bankScope3',x.category,'Invalid category',{blocking:true});
  const opsControlled=operations.filter(x=>x.withinControl==='Yes'), validBoundary=operations.every(x=>['Yes','No'].includes(x.withinControl));
  const completeS1=validBoundary&&opsControlled.length>0&&opsControlled.every(x=>x.scope1Status==='Ready');
  const completeLB=validBoundary&&opsControlled.length>0&&opsControlled.every(x=>x.scope2LBStatus==='Ready');
  const completeMB=dualReporting==='Yes'&&validBoundary&&opsControlled.length>0&&opsControlled.every(x=>['Ready','Grid fallback'].includes(x.scope2MBStatus));
  const inventoryComplete=completeS1&&completeLB&&dualReporting!=='Unknown'&&(dualReporting==='No'||completeMB);
  const scope1=sum(operations,'scope1'),scope2LB=sum(operations,'scope2LB'),knownMarketBasedSubtotal=sum(operations,'scope2MB'),scope2MB=dualReporting==='Yes'?knownMarketBasedSubtotal:null;
  const electricityKWh=sum(opsControlled,'electricityKWh'), contractKWh=sum(opsControlled,'contractKWh');
  const unknownKWh=opsControlled.some(x=>!nonnegative(x.electricityKWh))||!validBoundary;
  const unknownContract=opsControlled.some(x=>!nonnegative(x.contractKWh)||x.contractKWh>x.electricityKWh);
  const screenedCategories=bankScope3.filter(x=>['Ready','Not applicable','Excluded justified','Modeled lending only'].includes(x.status)).length;
  const bank={scope1,scope2LB,scope2MB,scope1And2LB:scope1+scope2LB,scope1And2MB:scope2MB===null?null:scope1+scope2MB,electricityKWh,contractKWh,contractCoverage:unknownKWh||unknownContract?null:ratio(contractKWh,electricityKWh),
    contractAttributedTonnes:sum(operations,'contractEmissions'),contractLocationCounterfactual:sum(operations,'contractLocationCounterfactual'),contractVsLocationDifference:sum(operations,'contractLocationCounterfactual')-sum(operations,'contractEmissions'),residualVsLocationUplift:sum(operations,'residualEmissions')-sum(operations,'residualLocationCounterfactual'),marketVsLocationDifference:scope2MB===null?null:scope2MB-scope2LB,
    nonInvestmentScope3:sum(bankScope3.filter(x=>x.category<15),'selectedGross'),screenedCategories,scope3Status:screenedCategories<15?'Screening incomplete':'Partial investment coverage',scope1And2Status:inventoryComplete?'Complete for listed operations':'Incomplete',dualReporting,controlBasis:controlBasis??null,knownMarketBasedSubtotal,
    completeScope1:completeS1?scope1:null,completeScope2LB:completeLB?scope2LB:null,completeScope2MB:completeMB?scope2MB:null,completeScope1And2LB:completeS1&&completeLB?scope1+scope2LB:null,completeScope1And2MB:completeS1&&completeMB?scope1+scope2MB:null,
    controlledOperationCount:opsControlled.length,scope1ReadyCount:opsControlled.filter(x=>x.scope1Status==='Ready').length,scope2LBReadyCount:opsControlled.filter(x=>x.scope2LBStatus==='Ready').length,scope2MBReadyCount:opsControlled.filter(x=>['Ready','Grid fallback'].includes(x.scope2MBStatus)).length,scope2LBCoverageKWh:unknownKWh?null:ratio(sum(opsControlled.filter(x=>x.scope2LBStatus==='Ready'),'electricityKWh'),electricityKWh),scope2MBCoverageKWh:dualReporting!=='Yes'||unknownKWh?null:ratio(sum(opsControlled.filter(x=>['Ready','Grid fallback'].includes(x.scope2MBStatus)),'electricityKWh'),electricityKWh),comparisonComplete:completeLB&&completeMB,
    category15S1S2:bankScope3[14].selectedGross,investeeScope3:portfolio.financedS3,gridFallbackOperations:operations.filter(x=>x.scope2MBStatus==='Grid fallback').map(x=>x.id)};
  const result={methodVersion:EMISSIONS_METHOD_VERSION,reviewDate:reviewDate??null,portfolio,customers,sectors:groups('sector'),countries:groups('country'),facilities,entities,activities,factors,operations,bankScope3,scope3:bankScope3,bank,gaps,
    counts:{borrowers:borrowers.length,facilities:facilities.length,entities:entities.length,sites:sites.length,factors:factors.length,activities:activities.length,operations:operations.length,scope3Categories:bankScope3.length,blockingGaps:gaps.filter(x=>x.blocking).length,overlapReviews:gaps.filter(x=>x.module==='overlap').length},
    metadata:{inventoryYear:reviewYear,numericalBasis:'Input-source labels control provenance; seeded reference inputs are synthetic.',inventoryVsScenario:'Changing a credit scenario does not restate this inventory.',fullStandardConformance:false,aggregateContract:'Metric fields are covered known subtotals. Use displayFinanced* for zero-coverage presentation and completeScope* for complete bank totals.'}};
  return result;
}

export default calculateEmissions;
