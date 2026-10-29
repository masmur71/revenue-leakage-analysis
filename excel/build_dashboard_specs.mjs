// Build reviewed inputs for the user-run native VBA runner. Never opens Excel.
import fs from 'node:fs';
import path from 'node:path';
const root=path.resolve(import.meta.dirname,'..'), out=path.join(root,'excel','dashboard-specs');
fs.mkdirSync(out,{recursive:true});
fs.copyFileSync(path.join(root,'data','excel','manifest.json'),path.join(out,'manifest.json'));
function parse(text){let rows=[],row=[],s='',q=false;for(let i=0;i<text.length;i++){const c=text[i];if(c==='"'){if(q&&text[i+1]==='"'){s+='"';i++;}else q=!q;}else if(c===','&&!q){row.push(s);s='';}else if(c==='\n'&&!q){row.push(s.replace(/\r$/,''));rows.push(row);row=[];s='';}else s+=c;}if(q)throw Error('Unclosed CSV');if(s||row.length){row.push(s.replace(/\r$/,''));rows.push(row);}return rows;}
function read(name){const [h,...r]=parse(fs.readFileSync(path.join(root,'data','excel',name+'.csv'),'utf8').replace(/^\uFEFF/,''));return r.map(x=>Object.fromEntries(h.map((k,i)=>[k,x[i]])));}
const esc=x=>/[",\r\n]/.test(String(x))?'"'+String(x).replaceAll('"','""')+'"':String(x);
function write(name,headers,rows){fs.writeFileSync(path.join(out,name+'.csv'),[headers,...rows].map(r=>r.map(esc).join(',')).join('\r\n')+'\r\n');}
const all=read('SQLReference').map(r=>r.metric_id);
const pivots=[
 ['pTrend','Executive Summary','A15','DimDate.year_month','','net_sales|target_net_sales','line'],
 ['pBranchAchievement','Executive Summary','A50','DimBranch.region_name|DimBranch.branch_name','','net_sales|target_gap|achievement_rate',''],
 ['pExecutive','Data & Calculations','AA5','','','net_sales|gross_profit|gross_margin_rate|target_net_sales|achievement_rate|price_shortfall|discount_amount|return_amount',''],
 ['pLeakBranch','Leakage Diagnostics','A10','DimBranch.branch_name','','reference_sales|price_upside|price_shortfall|discount_amount|billed_sales|return_amount|net_sales|gross_profit',''],
 ['pLeakSKU','Leakage Diagnostics','A65','DimProduct.sku_name','','price_shortfall|return_amount|return_rate|billed_sales',''],
 ['pLeakOutlet','Leakage Diagnostics','A90','DimOutlet.outlet_name','','price_shortfall|return_amount|return_rate|billed_sales',''],
 ['pReturnReasons','Leakage Diagnostics','A115','FactCommercial.return_reason','','return_amount|return_margin_reversal',''],
 ['pLeakChart','Data & Calculations','AA120','DimBranch.branch_name','','price_shortfall','bar'],
 ['pAR','Operations','A10','DimBranch.branch_name','','outstanding_ar|overdue_ar|overdue_rate|dso_90d',''],
 ['pVisits','Operations','A40','DimSalesman.salesman_name','','completed_visits|visit_completion_rate|effective_call_rate|order_conversion_rate|calls_per_salesman_day',''],
 ['pRFM','Operations','A100','FactRFM.segment','','rfm_outlets|rfm_monetary',''],
 ['pCohort','Operations','A120','FactCohort.cohort_month','FactCohort.activity_month','retention_rate',''],
 ['pAcceptanceControls','Data & Calculations','AA30','','',all.concat(['price_bridge_residual','net_bridge_residual','aging_residual']).join('|'),'']
];
write('pivots',['name','sheet','anchor','rows','columns','measures','chart'],pivots);
write('slicers',['name','field','top_row'],[
 ['rlYear','DimDate.year',3],['rlMonth','DimDate.year_month',15],['rlRegion','DimBranch.region_name',27],['rlBranch','DimBranch.branch_name',39],['rlCategory','DimProduct.category',51],['rlSKU','DimProduct.sku_name',63],['rlOutlet','DimOutlet.outlet_id',75],['rlSalesman','DimSalesman.salesman_name',87],['rlChannel','DimChannel.channel',99]
]);
write('return-reasons',['key'],[...new Set(read('FactCommercial').map(r=>r.return_reason).filter(Boolean))].map(k=>[k]));
const context=read('ReportContext')[0], cases=[], requests=[], extended=[];
const headers=['case_id','start_date','end_date','branch_id','sku_id','outlet_id','channel','salesman_id','cohort_month','activity_month','segment','date_list'];
const addCase=(id,values={})=>cases.push(headers.map(k=>k==='case_id'?id:values[k]??''));
const period={start_date:context.period_start,end_date:context.period_end};
addCase('reference',period);
for(const r of read('SQLFilterControls'))addCase(r.case_id,r);
for(const r of read('SQLSupportControls'))addCase(r.case_id,{...period,...r});
const [,...template]=parse(fs.readFileSync(path.join(root,'docs','excel','native-results.template.csv'),'utf8'));
for(const [c,m] of template)requests.push([c,m,'main']);
function extra(id,vals,checks){addCase(id,{...period,...vals});for(const [metric,value,unit='IDR'] of checks){requests.push([id,metric,'extended']);extended.push([id,metric,value,unit]);}}
const targetBlank=[['target_net_sales','','IDR'],['achievement_rate','','ratio']];
const snapshotBlank=[['outstanding_ar',''],['dso_90d','','days'],['completed_visits','','visits'],['rfm_outlets','','outlets'],['retention_rate','','ratio']];
extra('guard_product',{sku_id:'SKU001'},[...targetBlank,...snapshotBlank]);
extra('guard_outlet',{outlet_id:'O0170'},targetBlank);
extra('guard_salesman',{salesman_id:read('DimSalesman')[0].salesman_id},[...targetBlank,['rfm_outlets','','outlets'],['retention_rate','','ratio']]);
extra('guard_channel',{channel:'GT'},targetBlank);
extra('guard_partial',{start_date:'2025-01-15',end_date:'2025-01-31'},targetBlank);
extra('guard_noncontiguous',{date_list:'2025-01-01|2025-03-01'},targetBlank);
const fullMonth=read('SQLBranchMonth').find(r=>r.branch_id==='B003'&&r.month==='2025-01-01');
if(!fullMonth?.target)throw Error('Missing SQL full-month target control');
extra('guard_full_month',{start_date:'2025-01-01',end_date:'2025-01-31',branch_id:'B003'},[['target_net_sales',fullMonth.target]]);
for(const r of read('SQLFilterControls')){
 const id='residual_'+r.case_id; extra(id,r,[['price_bridge_residual','0'],['net_bridge_residual','0']]);
}
extra('snapshot_2024',{start_date:'2024-01-01',end_date:'2024-12-31'},[['outstanding_ar',read('SQLSupportControls')[0].outstanding_ar],['aging_residual','0']]);
const segments=new Map();for(const r of read('FactRFM')){const a=segments.get(r.segment)??{n:0,m:0n};a.n++;a.m+=BigInt(r.monetary.replace('.',''));segments.set(r.segment,a);}
for(const [seg,a] of segments){const value=(a.m/100n).toString()+'.'+(a.m%100n).toString().padStart(2,'0');extra('rfm_'+seg.replaceAll(' ','_'),{segment:seg},[['rfm_outlets',String(a.n),'outlets'],['rfm_monetary',value]]);}
for(const r of read('SQLCohort'))extra('cohort_'+r.cohort_month+'_'+r.activity_month,{cohort_month:r.cohort_month,activity_month:r.activity_month},[['cohort_outlets',r.cohort_outlets,'outlets'],['cohort_active_outlets',r.active_outlets,'outlets'],['retention_rate',r.retention_rate,'ratio']]);
extra('cohort_future',{cohort_month:'2025-12-01',activity_month:'2024-01-01'},[['retention_rate','','ratio']]);
write('cases',headers,cases);write('requests',['case_id','metric_id','destination'],requests);write('extended-expected',['case_id','metric_id','value','unit'],extended);
// Remaining representative native checks after the interrupted exhaustive run.
const sampleCases=[
 ['cohort_2025-01-01_2025-01-01','2025-01-01','2025-01-01'],
 ['cohort_2025-12-01_2025-12-01','2025-12-01','2025-12-01'],
 ['cohort_unobserved','2025-12-01','2025-01-01'],
 ['cohort_multi_month','2025-01-01','2025-01-01|2025-02-01']
];
write('cohort-samples',headers,sampleCases.map(([id,cohort_month,activity_month])=>headers.map(k=>({case_id:id,cohort_month,activity_month}[k]??''))));
const sampleExpected=[];
for(const [id] of sampleCases.slice(0,2))for(const r of extended.filter(r=>r[0]===id))sampleExpected.push(r);
sampleExpected.push(['cohort_unobserved','retention_rate','','ratio']);
for(const [m,u] of [['cohort_outlets','outlets'],['cohort_active_outlets','outlets'],['retention_rate','ratio']])sampleExpected.push(['cohort_multi_month',m,'',u]);
if(sampleExpected.length!==10)throw Error('Missing SQL cohort sample controls');
write('cohort-samples-expected',['case_id','metric_id','value','unit'],sampleExpected);
write('cohort-sample-requests',['case_id','metric_id'],sampleExpected.map(r=>r.slice(0,2)));
const contract=JSON.parse(fs.readFileSync(path.join(root,'excel','model_contract.json'),'utf8'));
write('tables',['name','rows'],Object.entries(contract).filter(([,v])=>v.load==='model-only').map(([k])=>[k,read(k).length]));
console.log(`Dashboard specs: ${pivots.length} pivots, ${cases.length} cases, ${requests.length} native captures.`);
