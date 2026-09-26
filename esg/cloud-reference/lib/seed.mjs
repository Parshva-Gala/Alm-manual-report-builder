export const serial=s=>Math.round((Date.parse(s+'T00:00:00Z')-Date.UTC(1899,11,30))/86400000);
const names=['Nile Cement','Delta Solar','Alexandria Logistics','Oasis Foods','Cairo Properties','Red Sea Hotels','Amman Cement','Petra Solar','Levant Freight','Jordan Foods','Aqaba Properties','Amman Healthcare','Emirates Materials','Desert Solar','Gulf Logistics','Emirates Foods','Marina Properties','Gulf Hospitality','Nile Textiles','Delta Agriculture','Sinai Utilities','Petra Manufacturing','Gulf Technology','Cairo Mobility'];
const sectors=['Cement','Solar','Logistics','Food','Property','Hospitality','Cement','Solar','Logistics','Food','Property','Healthcare','Cement','Solar','Logistics','Food','Property','Hospitality','Manufacturing','Agriculture','Utilities','Manufacturing','Technology','Transport'];
export const borrowers=names.map((name,i)=>({id:'B'+String(i+1).padStart(3,'0'),name:'Demo '+name,country:i<6||i===18||i===19||i===20||i===23?'Egypt':i<12||i===21?'Jordan':'UAE',sector:sectors[i],listed:i%5===0?'Listed':'Private',reportDate:serial('2026-12-31'),revenue:60+i*7,ebitda:12+i*1.5,maintCapex:2+i*.15,debtService:5+i*.3,equity:40+i*3,debt:25+i*2,evic:110+i*6,pd:.012+(i%6)*.006,lgd:.28+(i%5)*.04,pricedTonnes:sectors[i]==='Cement'?120000+i*1000:sectors[i]==='Solar'?200:4000+i*400,energyShare:sectors[i]==='Cement'?.23:.07+(i%4)*.02,passThrough:.35+(i%4)*.1,margin:.32+(i%3)*.06,pdSensitivity:1.6,abatement:.2,evidence:'SYN-FIN-'+String(i+1),expiry:serial('2027-12-31'),defaultFlag:i===22?'Yes':'No'}));
export const facilities=[];export const entities=[];
for(const [i,b]of borrowers.entries()){
 const assetClass=b.listed==='Listed'?'Corporate listed':'Corporate private';
 entities.push({entityId:b.id,borrowerId:b.id,name:b.name,assetClass,country:b.country,equity:b.equity,debt:b.debt,evic:b.evic,origValue:0,gfa:0});
 const count=i<12?2:1;
 for(let j=0;j<count;j++){
  const designated=j===1;const ac=designated?(b.sector==='Solar'?'Project finance':b.sector==='Property'?'CRE':b.sector==='Logistics'?'Motor':assetClass):assetClass;
  const special=['Project finance','CRE','Motor'].includes(ac);const entityId=special?'A'+b.id.slice(1):b.id;
  if(special)entities.push({entityId,borrowerId:b.id,name:b.name+(ac==='Project finance'?' solar project':ac==='CRE'?' financed property':' financed fleet'),assetClass:ac,country:b.country,equity:20,debt:15,evic:0,origValue:30,gfa:12000});
  const drawn=(j?2.5:5)+i*.35;
  facilities.push({id:'F'+String(facilities.length+1).padStart(3,'0'),borrowerId:b.id,assetClass:ac,currency:'USD',drawnLocalm:drawn,undrawnLocalm:j?.5:1.5,ccf:.5,tenor:i%7===0?12:5+(i%5),eir:.08,origPD:i%4===0?b.pd/2.4:b.pd,dpd:i===22?120:i===9?35:0,stageOverride:'Auto',origCollateral:drawn*1.4,currentCollateral:drawn*1.35,securedRecovery:.65,amortRate:.12,baseRW:1,yield:.095,emissionsEntityId:entityId,purpose:special?'Designated asset':'General corporate',taxonomyActivity:b.sector==='Solar'?'Solar PV':b.sector==='Cement'?'Cement':b.sector==='Property'?'Building renovation':b.sector==='Logistics'?'Road freight':'Other',openingAllowance:i===22?drawn*.45:drawn*b.pd*b.lgd*2,evidence:'SYN-LOAN-'+(facilities.length+1),overrideEvidence:'',overrideReviewer:'',regCCF:.5,defaultRW:1.5,recoveryLag:2,recoveryRate:.55,qualitativeSICR:'No',stressDefaultYear:0});
 }
}
export const sites=borrowers.flatMap((b,i)=>[0,1].map(j=>({id:'S'+String(i*2+j+1).padStart(3,'0'),borrowerId:b.id,country:b.country,region:j?'Secondary location':'Main location',peril:j?'Heat':'Flood',assetValue:20+i*2,revenueShare:j?.4:.6,hazard:.2+(i%5)*.12,vulnerability:.35+(i%3)*.15,adaptation:.2,insurance:.45,collateralShare:j?.4:.6,evidence:'SYN-SITE-'+(i*2+j+1),expiry:serial('2027-12-31')})));
export const cases=['Baseline','Orderly transition','Delayed transition','Physical shock'];
export const sectorsList=[...new Set(sectors)];
