// Plain-JavaScript parity reference, outside the application. Amounts are USD millions.
// No spreadsheet engine, DOM, network, mutable singleton state or external dependency.
import { DEFAULT_INPUTS } from './financial_defaults.mjs';
export { DEFAULT_INPUTS };
const sum=(xs,key)=>xs.reduce((a,x)=>a+(key?x[key]:x),0);
const pos=x=>Math.max(0,x), min=Math.min, max=Math.max;
const finite=x=>typeof x==='number'&&Number.isFinite(x);
const clone=x=>JSON.parse(JSON.stringify(x));

/** Constant annual PD survival/amortisation buckets; recoveries discounted from default. */
export function performingECL({pd,lgd,ead,eir,amort,horizon,recoveryLag}) {
  const q=(1-pd)*(1-amort)/(1+eir);
  const tail=Math.abs(1-q)<1e-9?horizon:(1-q**horizon)/(1-q);
  return ead*(1-(1-lgd)/(1+eir)**recoveryLag)*pd/(1+eir)*tail;
}
export function accountingStage(f,b,s) {
  if(b.creditImpaired==='Yes'||f.daysPastDue>=90)return 3;
  if(f.stageOverride!=='Auto'&&f.overrideEvidence&&f.overrideReviewer)return Number(f.stageOverride);
  return f.daysPastDue>=30||f.qualitativeSICR==='Yes'||(b.pd>=f.originationPD*s.sicrRelative&&b.pd-f.originationPD>=s.sicrAbsolute)?2:1;
}

/** Recompute readiness from editable fields, never trust exported sourceStatus text. */
export function validateInputs(d,path){
  const issues=[],add=(id,message)=>issues.push({id,message});
  const nums=(o,keys,id)=>{for(const k of keys)if(!finite(o[k]))add(id,`${k} must be a finite number`);};
  const range=(x,a,b,id,k)=>{if(!finite(x)||x<a||x>b)add(id,`${k} outside ${a}..${b}`);};
  const unique=(xs,type)=>{const ids=new Set();for(const x of xs){if(!x.id||ids.has(x.id))add(x.id||type,'Missing/duplicate '+type+' ID');ids.add(x.id);}};
  unique(d.borrowers,'borrower');unique(d.facilities,'facility');unique(d.sites,'site');
  if(!d.facilities.length)add('portfolio','No facilities');
  const bs=new Map(d.borrowers.map(b=>[b.id,b]));
  const settings=d.settings;
  nums(settings,['reportDate','firstYear','sicrRelative','sicrAbsolute','taxRate','dividendPayout','cet1Target','tier1Target','totalCapitalTarget','lcrTarget','nsfrTarget'],'settings');
  for(const k of ['taxRate','dividendPayout'])range(settings[k],0,1,'settings',k);
  if(settings.sicrRelative<0||settings.sicrAbsolute<0)add('settings','SICR thresholds must be nonnegative');
  if(!settings.accountingCases.length||Math.abs(sum(settings.accountingCases,'weight')-1)>1e-6)add('accounting','Case probabilities must sum to one');
  for(const c of settings.accountingCases)for(const k of ['weight','pdMultiple','lgdMultiple'])if(!finite(c[k])||c[k]<0)add('accounting',`${k} missing or negative`);
  for(const b of d.borrowers){
    nums(b,['revenue','ebitda','maintCapex','debtService','equity','debt','evic','pd','lgd','pricedTonnes','energyShare','passThrough','margin','pdSensitivity','abatement','expiry','financialDate'],b.id);
    if(!b.name||!b.country||!b.sector||!b.listing||!b.evidence)add(b.id,'Missing borrower attributes/evidence');
    if(!(b.revenue>0&&b.ebitda>b.maintCapex&&b.debtService>0))add(b.id,'Revenue/base cash/debt service must be positive');
    for(const k of ['pd','lgd','energyShare','passThrough','margin','abatement'])range(b[k],0,1,b.id,k);
    if(b.pricedTonnes<0||b.pdSensitivity<0)add(b.id,'Negative emissions or PD sensitivity');
    if(!['Yes','No'].includes(b.creditImpaired))add(b.id,'Credit-impaired flag must be Yes/No');
    if(b.expiry<settings.reportDate)add(b.id,'Expired borrower evidence');
    const sites=d.sites.filter(x=>x.borrowerId===b.id);
    if(!sites.length)add(b.id,'No physical risk site records');
    for(const k of ['revenueShare','collateralShare'])if(Math.abs(sum(sites,k)-1)>1e-6)add(b.id,k+' must sum to one');
  }
  for(const f of d.facilities){
    if(!bs.has(f.borrowerId))add(f.id,'Missing borrower');
    const keys=['drawnLocalm','undrawnLocalm','accountingCCF','tenor','eir','originationPD','daysPastDue','originalCollateral','currentCollateral','securedRecovery','amortRate','baseRW','yield','prudentialProvisionUSDm','regulatoryCCF','defaultRW','recoveryLag','recoveryFraction','stressDefaultYear','bookedUndrawnProvisionLocalm'];nums(f,keys,f.id);
    for(const k of ['accountingCCF','eir','originationPD','securedRecovery','amortRate','regulatoryCCF','recoveryFraction'])range(f[k],0,1,f.id,k);
    for(const k of ['drawnLocalm','undrawnLocalm','prudentialProvisionUSDm','daysPastDue','baseRW','yield','defaultRW','bookedUndrawnProvisionLocalm'])if(f[k]<0)add(f.id,k+' must be nonnegative');
    if(!Number.isInteger(f.tenor)||f.tenor<1||f.tenor>30)add(f.id,'Remaining years must be integer 1..30');
    if(!Number.isInteger(f.recoveryLag)||f.recoveryLag<0||f.recoveryLag>30)add(f.id,'Recovery lag must be integer 0..30');
    if(!Number.isInteger(f.daysPastDue))add(f.id,'Days past due must be integer');
    if(!Number.isInteger(f.stressDefaultYear)||f.stressDefaultYear<0||f.stressDefaultYear>5||f.stressDefaultYear>f.tenor)add(f.id,'Default year must be 0..min(5,tenor)');
    if(f.stressDefaultYear>0&&!f.defaultScheduleSource)add(f.id,'Missing default schedule evidence');
    if(!['Yes','No'].includes(f.qualitativeSICR))add(f.id,'Qualitative SICR must be Yes/No');
    if(!['Auto','1','2','3',1,2,3].includes(f.stageOverride))add(f.id,'Invalid stage override');
    if(f.stageOverride!=='Auto'&&(!f.overrideEvidence||!f.overrideReviewer))add(f.id,'Stage override requires evidence and reviewer');
    const fx=settings.fxUSDPerUnit[f.currency];if(!finite(fx)||fx<=0)add(f.id,'Missing positive FX conversion');
    if(f.prudentialProvisionUSDm>f.drawnLocalm*fx)add(f.id,'Prudential provision exceeds drawn exposure');
    if(!f.evidence)add(f.id,'Missing facility evidence');
  }
  for(const p of d.sites){
    if(!bs.has(p.borrowerId))add(p.id,'Missing borrower');
    nums(p,['assetValue','revenueShare','hazard','vulnerability','adaptation','insurance','collateralShare','expiry'],p.id);
    for(const k of ['revenueShare','hazard','vulnerability','adaptation','insurance','collateralShare'])range(p[k],0,1,p.id,k);
    if(p.assetValue<0||!p.evidence||p.expiry<settings.reportDate)add(p.id,'Invalid asset value or missing/expired site evidence');
  }
  if(!path){add('scenario','Unknown scenario');return issues;}
  for(const [k,a]of Object.entries(path)){if(k==='demand')continue;if(!Array.isArray(a)||a.length!==5||a.some(x=>!finite(x)))add('scenario',k+' requires five numeric periods');}
  for(const k of ['interruption','damage','abatement','withdrawal','hqlaHaircut','drawRate'])for(const x of path[k]||[])range(x,0,1,'scenario',k);
  for(const k of ['carbon','capex','operatingLoss','marketLoss','rwFactor'])if((path[k]||[]).some(x=>x<0))add('scenario',k+' must be nonnegative');
  if((path.energy||[]).some(x=>x< -1))add('scenario','Energy shock below -100%');
  for(const b of d.borrowers){const a=path.demand?.[b.sector];if(!a||a.length!==5||a.some(x=>!finite(x)||x< -1))add('scenario','Missing/invalid demand path: '+b.sector);}
  for(const k of Object.keys(DEFAULT_INPUTS.bank))if(!finite(d.bank[k])||d.bank[k]<0)add('bank',k+' must be a nonnegative number');
  for(const k of ['depositRunoff','undrawnDrawFactor','inflowCap','depositASF','termFundingASF','performingLoanRSF','defaultedLoanRSF','securitiesRSF','otherAssetRSF','undrawnRSF','inflowRecognition','securitiesHQLARecognition'])range(d.bank[k],0,1,'bank',k);
  if(d.bank.tier2>d.bank.openingTermFunding)add('bank','Tier 2 exceeds opening term funding');
  if(d.actions.length!==5)add('actions','Five action periods required');
  for(const [i,a]of d.actions.entries()){for(const k of ['equity','funding','fees'])if(!finite(a[k])||a[k]<0)add('actions',`${i+1}: invalid ${k}`);if(!['Proposed','Approved','Rejected'].includes(a.approval))add('actions','Invalid approval state');}
  return issues;
}

/**
 * calculate({inputs?, scenario?, actionsEnabled?, country?, carbonScale?, physicalScale?, actionOverride?})
 * inputs is the full schema in DEFAULT_INPUTS. Clone before editing. No argument mutation.
 * actionOverride replaces the full five-year action schedule with a one-off action:
 * {yearIndex:1..5,equity,funding,fees,approval,evidence}; omission retains workbook schedule.
 * carbonScale multiplies only incremental carbon prices; physicalScale multiplies only
 * interruption/damage severities, not deposit/market/bank-loss shocks. Default scales=1.
 */
export function calculate(state={}){
  const d=clone(state.inputs||DEFAULT_INPUTS),s=d.settings,k=d.bank;
  s.firstYear=finite(s.reportDate)?new Date(Date.UTC(1899,11,30)+s.reportDate*86400000).getUTCFullYear()+1:NaN;
  const scenario=state.scenario??s.scenario,enabled=state.actionsEnabled??(s.managementActions==='On');
  const country=state.country??s.country;
  const carbonScale=state.carbonScale??1,physicalScale=state.physicalScale??1;
  let path=d.paths[scenario]?clone(d.paths[scenario]):null;
  if(path){path.carbon=path.carbon.map(x=>x*carbonScale);for(const a of ['interruption','damage'])path[a]=path[a].map(x=>x*physicalScale);}
  const custom=state.actionOverride;
  if(custom){d.actions=Array.from({length:5},()=>({equity:0,funding:0,fees:0,approval:'Proposed',evidence:''}));if(Number.isInteger(custom.yearIndex)&&custom.yearIndex>=1&&custom.yearIndex<=5)d.actions[custom.yearIndex-1]={equity:custom.equity,funding:custom.funding,fees:custom.fees,approval:custom.approval,evidence:custom.evidence};}
  const issues=validateInputs(d,path);
  if(!finite(carbonScale)||carbonScale<0||!finite(physicalScale)||physicalScale<0)issues.push({id:'controls',message:'Severity scales must be finite and nonnegative'});
  if(typeof enabled!=='boolean')issues.push({id:'controls',message:'actionsEnabled must be boolean'});
  if(custom&&(!Number.isInteger(custom.yearIndex)||custom.yearIndex<1||custom.yearIndex>5))issues.push({id:'actions',message:'Action yearIndex must be 1..5'});
  if(!['Egypt','Jordan','UAE','Global'].includes(country))issues.push({id:'country',message:'Unknown country profile'});
  const reportCurrency={Egypt:'EGP',Jordan:'JOD',UAE:'AED',Global:'USD'}[country];
  if(!finite(s.fxUSDPerUnit[reportCurrency])||s.fxUSDPerUnit[reportCurrency]<=0)issues.push({id:'country',message:'Missing positive reporting-currency FX rate'});
  // Atomic failure avoids silently presenting healthy aggregates from an incomplete book.
  if(issues.length)return {status:'Review inputs',issues,scenario,country,reportingECL:null,bank:[],borrowerProjections:[],facilityProjections:[]};
  const borrowers=new Map(d.borrowers.map(b=>[b.id,b]));
  const facilities=d.facilities.map(f=>{const fx=s.fxUSDPerUnit[f.currency],b=borrowers.get(f.borrowerId);return {...f,drawn:f.drawnLocalm*fx,undrawn:f.undrawnLocalm*fx,stage:accountingStage(f,b,s)};});
  const reportingRows=facilities.map(f=>{
    const b=borrowers.get(f.borrowerId),horizon=f.stage===1?1:f.tenor;
    const cases=s.accountingCases.map(c=>{const pd=min(1,b.pd*c.pdMultiple),lgd=min(1,b.lgd*c.lgdMultiple);
      const calc=(ead,amort)=>f.stage===3?ead*(1-max(0,1-(1-f.recoveryFraction)*c.lgdMultiple)/(1+f.eir)**f.recoveryLag):performingECL({pd,lgd,ead,eir:f.eir,amort,horizon,recoveryLag:f.recoveryLag});
      return {name:c.name,weight:c.weight,pd,lgd,drawn:calc(f.drawn,f.amortRate),undrawn:calc(f.undrawn*f.accountingCCF,0)};
    });
    const drawn=sum(cases.map(c=>c.drawn*c.weight)),undrawn=sum(cases.map(c=>c.undrawn*c.weight));
    return {facilityId:f.id,borrowerId:f.borrowerId,stage:f.stage,drawnExposure:f.drawn,undrawnExposure:f.undrawn,horizon,cases,drawn,undrawn,total:drawn+undrawn,allowanceRatio:f.drawn===0?null:drawn/f.drawn};
  });
  const eclMap=new Map(reportingRows.map(x=>[x.facilityId,x]));
  const reportingECL={drawn:sum(reportingRows,'drawn'),undrawn:sum(reportingRows,'undrawn'),total:sum(reportingRows,'total'),rows:reportingRows};
  const borrowerProjections=[],facilityProjections=[],bank=[];
  const opening={year:s.firstYear-1,cash:k.openingCash,securities:k.openingSecurities,grossLoans:sum(facilities,'drawn'),drawnAllowance:reportingECL.drawn,otherAssets:k.openingOtherAssets,deposits:k.openingDeposits,termFunding:k.openingTermFunding,undrawnProvision:reportingECL.undrawn,otherLiabilities:k.otherLiabilities,at1:k.at1,tier2:k.tier2,marketRWA:k.marketRWA,operationalRWA:k.operationalRWA};
  opening.netLoans=opening.grossLoans-opening.drawnAllowance;opening.totalAssets=opening.cash+opening.securities+opening.netLoans+opening.otherAssets;
  opening.equity=opening.totalAssets-opening.deposits-opening.termFunding-opening.undrawnProvision-opening.otherLiabilities;
  opening.totalLiabilitiesEquity=opening.deposits+opening.termFunding+opening.undrawnProvision+opening.otherLiabilities+opening.equity;
  opening.balanceDifference=opening.totalAssets-opening.totalLiabilitiesEquity;
  opening.cet1Before=opening.cet1After=opening.equity-k.capitalDeductions-k.at1;
  opening.creditRWA=sum(facilities.map(f=>max(0,f.drawn-f.prudentialProvisionUSDm)*(f.stage===3?f.defaultRW:f.baseRW)+f.undrawn*f.regulatoryCCF*f.baseRW));
  opening.totalRWA=opening.creditRWA+k.marketRWA+k.operationalRWA;
  for(let y=0;y<5;y++){
    const yearIndex=y+1,year=s.firstYear+y,prior=y?bank[y-1]:opening;
    const credit=d.borrowers.map(b=>{const sites=d.sites.filter(z=>z.borrowerId===b.id),x={borrowerId:b.id,yearIndex,year,baselineCash:b.ebitda-b.maintCapex};
      x.physicalExposure=sum(sites.map(z=>z.revenueShare*z.hazard*z.vulnerability*(1-z.adaptation)));
      x.physicalLossShare=min(1,sum(sites.map(z=>z.collateralShare*z.hazard*z.vulnerability*(1-z.adaptation)))*path.damage[y]);
      x.demandChange=b.revenue*b.margin*path.demand[b.sector][y];x.carbonCost=b.pricedTonnes*(1-b.abatement*path.abatement[y])*path.carbon[y]/1e6*(1-b.passThrough);
      x.energyCost=b.revenue*b.energyShare*path.energy[y]*(1-b.passThrough);x.interruptionLoss=b.revenue*b.margin*path.interruption[y]*x.physicalExposure;
      x.repairExpense=sum(sites.map(z=>z.assetValue*z.hazard*z.vulnerability*(1-z.adaptation)*(1-z.insurance)))*path.damage[y];
      x.ebitda=b.ebitda+x.demandChange-x.carbonCost-x.energyCost-x.interruptionLoss-x.repairExpense;x.transitionCapex=b.revenue*path.capex[y];x.availableCash=x.ebitda-b.maintCapex-x.transitionCapex;x.dscr=x.availableCash/b.debtService;
      x.cashDeterioration=max(-1,(x.baselineCash-x.availableCash)/max(.000001,x.baselineCash));
      x.pd=b.creditImpaired==='Yes'?1:b.pd===0?0:min(1,b.pd*Math.exp(min(20,b.pdSensitivity*x.cashDeterioration)));
      x.lgd=min(1,b.lgd+(1-b.lgd)*x.physicalLossShare);return x;
    });
    borrowerProjections.push(...credit);const creditMap=new Map(credit.map(x=>[x.borrowerId,x]));
    const rows=facilities.map((f,i)=>{
      const prev=y?facilityProjections[(y-1)*facilities.length+i]:null,e=eclMap.get(f.id),c=creditMap.get(f.borrowerId);
      const defaultYear=f.stage===3?0:f.stressDefaultYear===0?-1:f.stressDefaultYear,settle=defaultYear+f.recoveryLag;
      const x={facilityId:f.id,borrowerId:f.borrowerId,yearIndex,year,defaultYear,remainingLife:max(0,f.tenor-yearIndex),openingPerforming:prev?prev.closingPerforming:f.stage===3?0:f.drawn,openingDefaulted:prev?prev.closingDefaulted:f.stage===3?f.drawn:0,openingUndrawn:prev?prev.closingUndrawn:f.undrawn,pd:c.pd,lgd:c.lgd,stressedRW:f.baseRW*path.rwFactor[y],recoveryFraction:f.recoveryFraction,physicalLossShare:c.physicalLossShare};
      x.drawings=x.openingDefaulted>0||(defaultYear>=0&&defaultYear<yearIndex)||yearIndex>f.tenor?0:x.openingUndrawn*path.drawRate[y];
      x.repayments=x.openingDefaulted>0||(defaultYear>=0&&defaultYear===yearIndex)?0:yearIndex>=f.tenor?x.openingPerforming+x.drawings:min(x.openingPerforming+x.drawings,x.openingPerforming*f.amortRate);
      x.transferToDefault=defaultYear===yearIndex&&defaultYear>0?max(0,x.openingPerforming+x.drawings-x.repayments):0;
      x.cashRecoveries=defaultYear>=0&&yearIndex>=settle?(x.openingDefaulted+x.transferToDefault)*f.recoveryFraction:0;
      x.writeoffs=defaultYear>=0&&yearIndex>=settle?x.openingDefaulted+x.transferToDefault-x.cashRecoveries:0;
      x.closingPerforming=max(0,x.openingPerforming+x.drawings-x.repayments-x.transferToDefault);x.closingDefaulted=max(0,x.openingDefaulted+x.transferToDefault-x.cashRecoveries-x.writeoffs);
      x.closingUndrawn=yearIndex>=f.tenor||(defaultYear>=0&&defaultYear<=yearIndex)?0:max(0,x.openingUndrawn-x.drawings);
      x.stage=x.closingDefaulted>0?3:(f.stage===2||(x.pd>=f.originationPD*s.sicrRelative&&x.pd-f.originationPD>=s.sicrAbsolute))?2:1;
      x.openingDrawnAllowance=prev?prev.drawnAllowance:e.drawn;x.openingUndrawnProvision=prev?prev.undrawnProvision:e.undrawn;
      x.defaultOpeningAllowance=prev?prev.defaultClosingAllowance:f.stage===3?x.openingDrawnAllowance:0;
      x.defaultClosingAllowance=x.closingDefaulted>0?x.closingDefaulted*(1-f.recoveryFraction/(1+f.eir)**max(0,settle-yearIndex)):0;
      const horizon=x.stage===1?min(1,x.remainingLife):x.remainingLife;
      const ecl=(ead,amort)=>performingECL({pd:x.pd,lgd:x.lgd,ead,eir:f.eir,amort,horizon,recoveryLag:f.recoveryLag});
      x.drawnAllowance=ecl(x.closingPerforming,f.amortRate)+x.defaultClosingAllowance;x.undrawnProvision=ecl(x.closingUndrawn*f.accountingCCF,0);
      x.recoveryUnwind=x.openingDefaulted>0?max(0,x.openingDefaulted-x.defaultOpeningAllowance)*((1+f.eir)**min(1,max(0,settle-yearIndex+1))-1):0;
      x.creditCharge=x.drawnAllowance-x.openingDrawnAllowance+x.undrawnProvision-x.openingUndrawnProvision+x.writeoffs+x.recoveryUnwind;
      x.interestIncome=(x.openingPerforming+x.closingPerforming+x.transferToDefault)/2*f.yield+x.recoveryUnwind;
      const netting=f.drawn===0?0:min(1,f.prudentialProvisionUSDm/f.drawn);
      x.creditRWA=x.closingPerforming*(1-netting)*x.stressedRW+x.closingDefaulted*(1-netting)*f.defaultRW+x.closingUndrawn*f.regulatoryCCF*x.stressedRW;
      x.openingGross=x.openingPerforming+x.openingDefaulted;x.closingGross=x.closingPerforming+x.closingDefaulted;x.grossChange=x.closingGross-x.openingGross;
      x.grossReconciliation=x.grossChange-x.drawings+x.repayments+x.cashRecoveries+x.writeoffs;x.netCreditContribution=x.interestIncome-x.creditCharge;
      x.accountingEAD=x.closingGross+x.closingUndrawn*f.accountingCCF;x.netLoans=x.closingGross-x.drawnAllowance;
      x.collectionProxy30d=x.closingPerforming===0?0:((x.remainingLife<=1?x.closingPerforming:x.closingPerforming*f.amortRate)+x.closingPerforming*f.yield)/12;return x;
    });
    facilityProjections.push(...rows);const a=d.actions[y],effective=enabled&&a.approval==='Approved'&&!!a.evidence;
    const issuance=effective?a.equity:0,raised=effective?a.funding:0,fees=effective?a.fees:0;
    const b={yearIndex,year,actionEffective:effective,loanInterest:sum(rows,'interestIncome'),liquidIncome:(prior.cash+prior.securities)*k.liquidYield,feeIncome:k.annualFeeIncome,fundingExpense:prior.deposits*(k.depositRate+path.fundingSpread[y])+prior.termFunding*(k.termFundingRate+path.fundingSpread[y]),operatingExpense:k.annualOperatingCosts+fees,creditCharge:sum(rows,'creditCharge'),operatingLoss:path.operatingLoss[y],marketLoss:min(max(0,prior.otherAssets),path.marketLoss[y])+prior.securities*path.hqlaHaircut[y]};
    b.profitBeforeTax=b.loanInterest+b.liquidIncome+b.feeIncome-b.fundingExpense-b.operatingExpense-b.creditCharge-b.operatingLoss-b.marketLoss;b.tax=max(0,b.profitBeforeTax)*s.taxRate;b.profitAfterTax=b.profitBeforeTax-b.tax;b.dividends=max(0,b.profitAfterTax)*s.dividendPayout;b.retainedEarnings=b.profitAfterTax-b.dividends;
    b.securities=max(0,prior.securities*(1-path.hqlaHaircut[y]));b.grossLoans=sum(rows,'closingGross');b.drawnAllowance=sum(rows,'drawnAllowance');b.netLoans=b.grossLoans-b.drawnAllowance;b.otherAssets=max(0,prior.otherAssets-path.marketLoss[y]);b.deposits=prior.deposits*(1-path.withdrawal[y]);b.termFunding=prior.termFunding+raised;b.undrawnProvision=sum(rows,'undrawnProvision');b.otherLiabilities=k.otherLiabilities;b.equity=prior.equity+b.retainedEarnings+issuance;
    b.cash=prior.cash+b.profitAfterTax+b.creditCharge-sum(rows,'recoveryUnwind')+b.marketLoss+sum(rows,'repayments')+sum(rows,'cashRecoveries')-sum(rows,'drawings')+b.deposits-prior.deposits+raised+issuance-b.dividends;
    b.totalAssets=b.cash+b.securities+b.netLoans+b.otherAssets;b.totalLiabilitiesEquity=b.deposits+b.termFunding+b.undrawnProvision+b.otherLiabilities+b.equity;b.balanceDifference=b.totalAssets-b.totalLiabilitiesEquity;
    b.cet1After=b.equity-k.capitalDeductions-k.at1;b.at1=k.at1;b.tier2=k.tier2;b.creditRWA=sum(rows,'creditRWA');b.marketRWA=k.marketRWA;b.operationalRWA=k.operationalRWA;b.totalRWA=b.creditRWA+b.marketRWA+b.operationalRWA;
    b.hqla=max(0,b.cash)+b.securities*k.securitiesHQLARecognition;b.outflow30d=b.deposits*k.depositRunoff+sum(rows,'closingUndrawn')*k.undrawnDrawFactor;b.inflow30d=min(b.outflow30d*k.inflowCap,sum(rows,'collectionProxy30d')*k.inflowRecognition);b.netOutflow30d=max(0,b.outflow30d-b.inflow30d);
    b.lcr=b.cash<0?null:b.netOutflow30d===0?null:b.hqla/b.netOutflow30d;b.lcrStatus=b.cash<0?'Funding required':b.netOutflow30d===0?'No net outflow':'Calculated';
    b.asf=b.deposits*k.depositASF+b.termFunding*k.termFundingASF+b.equity;
    b.rsf=max(0,sum(rows,'closingPerforming')-(b.drawnAllowance-sum(rows,'defaultClosingAllowance')))*k.performingLoanRSF+max(0,sum(rows,'closingDefaulted')-sum(rows,'defaultClosingAllowance'))*k.defaultedLoanRSF+b.securities*k.securitiesRSF+b.otherAssets*k.otherAssetRSF+sum(rows,'closingUndrawn')*k.undrawnRSF;
    b.nsfr=b.rsf===0?null:b.asf/b.rsf;b.nsfrStatus=b.rsf===0?'No funding requirement':'Calculated';b.cashShortfall=max(0,-b.cash);
    b.cumulativeFunding=(prior.cumulativeFunding||0)+raised;b.cumulativeEquity=(prior.cumulativeEquity||0)+issuance;
    b.openingActionCash=(prior.cumulativeFunding||0)+(prior.cumulativeEquity||0)+(prior.cumulativeActionRetained||0);
    b.actionPretaxEffect=b.openingActionCash*k.liquidYield-(prior.cumulativeFunding||0)*(k.termFundingRate+path.fundingSpread[y])-fees;
    b.preActionProfitBeforeTax=b.profitBeforeTax-b.actionPretaxEffect;b.preActionTax=max(0,b.preActionProfitBeforeTax)*s.taxRate;b.preActionProfitAfterTax=b.preActionProfitBeforeTax-b.preActionTax;b.preActionDividends=max(0,b.preActionProfitAfterTax)*s.dividendPayout;
    b.currentActionRetained=b.retainedEarnings-(b.preActionProfitAfterTax-b.preActionDividends);b.cumulativeActionRetained=(prior.cumulativeActionRetained||0)+b.currentActionRetained;
    b.cet1Before=b.cet1After-b.cumulativeEquity-b.cumulativeActionRetained;b.preActionCash=b.cash-b.cumulativeFunding-b.cumulativeEquity-b.cumulativeActionRetained;b.preActionTermFunding=b.termFunding-b.cumulativeFunding;
    b.cet1Ratio=b.totalRWA===0?null:b.cet1After/b.totalRWA;b.preActionCet1Ratio=b.totalRWA===0?null:b.cet1Before/b.totalRWA;b.tier1Ratio=b.totalRWA===0?null:(b.cet1After+b.at1)/b.totalRWA;b.totalCapitalRatio=b.totalRWA===0?null:(b.cet1After+b.at1+b.tier2)/b.totalRWA;
    b.cet1Headroom=b.cet1After-s.cet1Target*b.totalRWA;b.totalCapitalHeadroom=b.cet1After+b.at1+b.tier2-s.totalCapitalTarget*b.totalRWA;
    b.capitalStatus=b.totalRWA===0?'No RWA':b.cet1Ratio<s.cet1Target||b.tier1Ratio<s.tier1Target||b.totalCapitalRatio<s.totalCapitalTarget?'Breach':'Within targets';
    b.liquidityStatus=b.cash<0||(b.lcr!==null&&b.lcr<s.lcrTarget)||(b.nsfr!==null&&b.nsfr<s.nsfrTarget)?'Breach':'Within targets';bank.push(b);
  }
  const currency={Egypt:'EGP',Jordan:'JOD',UAE:'AED',Global:'USD'}[country];
  const toLocalThousands=1000/s.fxUSDPerUnit[currency];
  return {status:'Ready',issues:[],scenario,country,controls:{actionsEnabled:enabled,carbonScale,physicalScale},unit:'USD millions',reportingCurrency:currency,localThousandsFactor:toLocalThousands,reportingECL,opening,bank,borrowerProjections,facilityProjections,reporting:{grossLoansLocalThousands:opening.grossLoans*toLocalThousands,eclLocalThousands:reportingECL.total*toLocalThousands}};
}
