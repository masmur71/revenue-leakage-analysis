// Independent CSV checks. This does NOT execute M, DAX, Power Query or Excel.
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
const root = path.resolve(import.meta.dirname,'..');
const data = path.join(root,'data','excel');
const evidence = path.join(root,'docs','excel');
const contract = JSON.parse(fs.readFileSync(path.join(root,'excel','model_contract.json'),'utf8'));
const manifest = JSON.parse(fs.readFileSync(path.join(data,'manifest.json'),'utf8').replace(/^\uFEFF/,''));
const assert = (ok,message) => {if(!ok) throw new Error(message);};
function* csv(text) {
 let row=[], field='', quoted=false;
 text=text.replace(/^\uFEFF/,'');
 for(let i=0;i<text.length;i++) {
  const c=text[i];
  if(c==='"') {if(quoted && text[i+1]==='"'){field+='"';i++;}else quoted=!quoted;}
  else if(c===',' && !quoted){row.push(field);field='';}
  else if(c==='\n' && !quoted){row.push(field.replace(/\r$/,''));yield row;row=[];field='';}
  else field+=c;
 }
 assert(!quoted,'Unclosed CSV quotation');
 if(field.length||row.length){row.push(field.replace(/\r$/,''));yield row;}
}
function* rows(name) {
 const it=csv(fs.readFileSync(path.join(data,`${name}.csv`),'utf8'));
 const headers=it.next().value;
 for(const fields of it){assert(fields.length===headers.length,`${name}: field count`);yield Object.fromEntries(headers.map((h,i)=>[h,fields[i]]));}
}
function fraction(s) {
 assert(/^-?\d+(\.\d+)?$/.test(s),`Not a decimal: ${s}`);
 const [a,b='']=s.split('.');return [BigInt(a+b),10n**BigInt(b.length)];
}
function cents(s) {
 const [n,d]=fraction(s);assert((n*100n)%d===0n,`Sub-cent money ${s}`);return n*100n/d;
}
function rounded(n,d,places) {
 assert(d!==0n,'Zero denominator');if(d<0n){n=-n;d=-d;}
 const sign=n<0n?-1n:1n;const scaled=(n<0n?-n:n)*10n**BigInt(places);
 return sign*((scaled+d/2n)/d);
}
const ratio=(n,d)=>d===0n?null:[n,d];
const counts={};const keys={};
for(const file of manifest.files) {
 const bytes=fs.readFileSync(path.join(data,file.file));
 assert(bytes.length===file.bytes,`${file.file}: size changed`);
 assert(crypto.createHash('sha256').update(bytes).digest('hex')===file.sha256,`${file.file}: hash changed`);
}
for(const [name,spec] of Object.entries(contract)) {
 const headers=csv(fs.readFileSync(path.join(data,`${name}.csv`),'utf8')).next().value;
 assert(JSON.stringify(headers)===JSON.stringify(spec.columns.map(c=>c.name)),`${name}: schema mismatch`);
 let count=0;const seen=new Set();
 for(const row of rows(name)) {
  count++;
  for(const c of spec.columns){const v=row[c.name];if(v==='')continue;
   if(c.type==='money')cents(v);
   if(c.type==='number')fraction(v);
   if(c.type==='int')assert(/^-?\d+$/.test(v),`${name}.${c.name}: integer`);
   if(c.type==='bool')assert(v==='t'||v==='f',`${name}.${c.name}: boolean`);
   if(c.type==='date')assert(/^\d{4}-\d{2}-\d{2}$/.test(v)&&new Date(v).toISOString().slice(0,10)===v,`${name}.${c.name}: date`);
  }
  if(spec.key){const k=row[spec.key];assert(k!==''&&!seen.has(k),`${name}: null/duplicate key ${k}`);seen.add(k);}
 }
 counts[name]=count;keys[name]=seen;
}
const rels=[...csv(fs.readFileSync(path.join(root,'excel','relationships.csv'),'utf8'))];
const relHeaders=rels.shift();
const byFact={};
for(const fields of rels){const r=Object.fromEntries(relHeaders.map((h,i)=>[h,fields[i]]));(byFact[r.fact]??=[]).push(r);}
for(const [fact,rels] of Object.entries(byFact))for(const row of rows(fact))for(const r of rels)assert(keys[r.dimension].has(row[r.fact_key]),`${fact}.${r.fact_key}: orphan ${row[r.fact_key]}`);
const context=[...rows('ReportContext')][0];
assert(counts.ReportContext===1,'Context must contain one row');
const branchMonths=new Set();for(const r of rows('FactTargets')){const key=r.branch_id+'|'+r.month_start;assert(!branchMonths.has(key),'Duplicate branch-month target');branchMonths.add(key);}
const expected=new Map();
for(const r of rows('SQLReference'))expected.set(`reference|${r.metric_id}`,{value:r.value===''?null:fraction(r.value),unit:r.unit});
assert(expected.size===46,'Expected 46 reference metrics');
const cases=[...rows('SQLFilterControls')];
cases.push({case_id:'reference',start_date:context.period_start,end_date:context.period_end});
const money=['gross_sales','discount_amount','billed_sales','cogs_sold','return_amount','returned_cogs','reference_sales','price_shortfall','price_upside','invalid_promo_discount','mismatch_sales'];
const accum=cases.map(c=>({c,s:Object.fromEntries(money.map(k=>[k,0n])),units:0n,returned:0n,lines:0n,promo:0n,invalid:0n,mismatch:0n,orders:new Set(),outlets:new Set()}));
for(const e of rows('FactCommercial'))for(const a of accum){const c=a.c;
 if(e.event_date<c.start_date||e.event_date>c.end_date)continue;
 if(['branch_id','sku_id','outlet_id','channel'].some(k=>c[k]&&e[k]!==c[k]))continue;
 for(const k of money)a.s[k]+=cents(e[k]);
 a.units+=BigInt(e.sales_units);a.returned+=BigInt(e.returned_units);a.promo+=BigInt(e.promo_lines);a.invalid+=BigInt(e.invalid_promo_lines);a.mismatch+=BigInt(e.mismatch_lines);
 if(e.event_type==='Sale'){a.lines++;a.orders.add(e.sales_order_id);a.outlets.add(e.outlet_id);assert(e.return_reason==='','Sale has return reason');}
 else assert(e.event_type==='Return'&&e.sales_order_id==='','Invalid return event');
}
let comparisons=0;
function compare(actual,want,label,unit='IDR') {
 assert((actual===null)===(want===null),`${label}: NULL mismatch`);
 if(actual===null)return;
 if(['units','orders','outlets','visits'].includes(unit)){assert(actual[0]*want[1]===want[0]*actual[1],`${label}: exact count mismatch`);comparisons++;return;}
 const places=unit==='ratio'?4:2;
 assert(rounded(...actual,places)===rounded(...want,places),`${label}: display rounding mismatch`);comparisons++;
}
for(const a of accum){const s=a.s;const net=s.billed_sales-s.return_amount;const profit=net-s.cogs_sold+s.returned_cogs;
 const v=Object.fromEntries(money.map(k=>[k,[s[k],100n]]));
 Object.assign(v,{net_sales:[net,100n],net_cogs:[s.cogs_sold-s.returned_cogs,100n],gross_profit:[profit,100n],return_margin_reversal:[s.return_amount-s.returned_cogs,100n],sales_units:[a.units,1n],returned_units:[a.returned,1n],orders:[BigInt(a.orders.size),1n],active_outlets:[BigInt(a.outlets.size),1n],average_order_value:ratio(s.billed_sales,100n*BigInt(a.orders.size)),gross_margin_rate:ratio(profit,net),discount_rate:ratio(s.discount_amount,s.gross_sales),return_rate:ratio(s.return_amount,s.billed_sales),price_deviation_rate:ratio(s.gross_sales-s.reference_sales,s.reference_sales),invalid_promo_line_rate:ratio(a.invalid,a.promo),salesman_mismatch_sales:[s.mismatch_sales,100n],salesman_mismatch_rate:ratio(a.mismatch,a.lines)});
 assert(s.reference_sales+s.price_upside-s.price_shortfall-s.discount_amount===s.billed_sales,`${a.c.case_id}: price bridge`);
 if(a.c.case_id==='reference')for(const [k,x] of Object.entries(v)){const want=expected.get(`reference|${k}`);if(want)compare(x,want.value,`reference ${k}`,want.unit);}
 else for(const [k,value] of Object.entries(a.c)){if(k in v){const unit=k.endsWith('_rate')?'ratio':['orders','active_outlets'].includes(k)?k:'IDR';const want=value===''?null:fraction(value);compare(v[k],want,`${a.c.case_id} ${k}`,unit);expected.set(`${a.c.case_id}|${k}`,{value:want,unit});}}
}
// Snapshot and visit controls are separately recomputed from row-level CSVs.
const ar={outstanding_ar:0n,overdue_ar:0n,ar_not_due:0n,ar_1_30:0n,ar_31_60:0n,ar_61_90:0n,ar_over_90:0n,credit_ar:0n,cash_ar:0n,overpayment:0n,denom:0n};
for(const r of rows('FactInvoiceSnapshot')){const amount=cents(r.outstanding),days=Number(r.days_past_due);ar.outstanding_ar+=amount;ar.overpayment+=cents(r.overpayment);ar.denom+=cents(r.credit_billings_90d);if(r.payment_method==='Credit')ar.credit_ar+=amount;else ar.cash_ar+=amount;const bucket=days<=0?'ar_not_due':days<=30?'ar_1_30':days<=60?'ar_31_60':days<=90?'ar_61_90':'ar_over_90';ar[bucket]+=amount;if(days>0)ar.overdue_ar+=amount;assert(cents(r.invoice_amount)-cents(r.allocated_paid)===amount-cents(r.overpayment),'Invoice balance bridge');}
for(const [k,value] of Object.entries(ar))if(k!=='denom')compare([value,100n],expected.get(`reference|${k}`).value,k);
compare(ratio(ar.overdue_ar,ar.outstanding_ar),expected.get('reference|overdue_rate').value,'overdue_rate','ratio');
compare(context.dso_history_complete==='t'?ratio(ar.credit_ar*90n,ar.denom):null,expected.get('reference|dso_90d').value,'dso_90d','days');
let receipts=0n;for(const r of rows('FactPayments'))if(r.allocation_status==='Allocated'&&r.payment_date>=context.period_start&&r.payment_date<=context.period_end)receipts+=cents(r.payment_amount);
compare([receipts,100n],expected.get('reference|allocated_receipts').value,'allocated_receipts');
let scheduled=0n,completed=0n,effective=0n;const calls=new Set(),converted=new Set(),days=new Set();
for(const r of rows('FactVisits'))if(r.visit_date>=context.period_start&&r.visit_date<=context.period_end){scheduled++;if(r.is_completed==='t'){completed++;calls.add(r.call_key);days.add(r.salesman_day_key);if(r.is_effective_call==='t')effective++;if(r.has_attributed_order==='t')converted.add(r.call_key);}}
const visit={scheduled_visits:[scheduled,1n],completed_visits:[completed,1n],effective_calls:[effective,1n],visit_completion_rate:ratio(completed,scheduled),effective_call_rate:ratio(effective,completed),order_conversion_rate:ratio(BigInt(converted.size),BigInt(calls.size)),calls_per_salesman_day:ratio(completed,BigInt(days.size))};
for(const [k,value] of Object.entries(visit)){const x=expected.get(`reference|${k}`);compare(value,x.value,k,x.unit);}
let target=0n;for(const r of rows('FactTargets'))if(r.month_start>=context.period_start&&r.month_start<=context.period_end)target+=cents(r.target_net_sales);
const complete=context.period_start.endsWith('-01')&&new Date(new Date(context.period_end).getTime()+86400000).toISOString().slice(8,10)==='01';
compare(complete?[target,100n]:null,expected.get('reference|target_net_sales').value,'target_net_sales');
const main=accum.at(-1).s;
compare(complete?ratio(main.billed_sales-main.return_amount,target):null,expected.get('reference|achievement_rate').value,'achievement_rate','ratio');
const cohort=new Map();for(const r of rows('FactCohort')){const key=r.cohort_month+'|'+r.activity_month;const a=cohort.get(key)??{n:0n,d:0n};a.n+=BigInt(r.active_outlet);a.d++;cohort.set(key,a);}
for(const r of rows('SQLCohort')){const a=cohort.get(r.cohort_month+'|'+r.activity_month);assert(a&&a.n===BigInt(r.active_outlets)&&a.d===BigInt(r.cohort_outlets),'Cohort cell mismatch');compare(ratio(a.n,a.d),fraction(r.retention_rate),'cohort retention','ratio');}
const saleCalls=new Set();const outletRfm=new Map();
for(const r of rows('FactCommercial')){
 if(r.event_type==='Sale')saleCalls.add(r.salesman_id+'|'+r.event_date+'|'+r.outlet_id);
 if(r.event_date<context.observed_rfm_start)continue;
 const a=outletRfm.get(r.outlet_id)??{last:'',orders:new Set(),money:0n};
 a.money+=cents(r.billed_sales)-cents(r.return_amount);
 if(r.event_type==='Sale'){a.orders.add(r.sales_order_id);if(r.event_date>a.last)a.last=r.event_date;}
 outletRfm.set(r.outlet_id,a);
}
for(const r of rows('FactVisits'))assert((r.has_attributed_order==='t')===saleCalls.has(r.call_key),'Visit attribution differs from sales-call keys');
const active=[...outletRfm.values()].filter(r=>r.orders.size>0);
const frequencies=active.map(r=>r.orders.size).sort((a,b)=>a-b);
const monetaries=active.map(r=>r.money).sort((a,b)=>a<b?-1:a>b?1:0);
const score=(sorted,value)=>{let low=0,high=sorted.length;while(low<high){const mid=(low+high)>>1;if(sorted[mid]<=value)low=mid+1;else high=mid;}return Math.ceil(5*low/sorted.length);};
for(const r of rows('FactRFM')){
 const a=outletRfm.get(r.outlet_id)??{last:'',orders:new Set(),money:0n};const n=a.orders.size;
 const recency=a.last?Math.round((new Date(context.as_of_date)-new Date(a.last))/86400000):null;
 const rs=n===0?0:recency<=30?5:recency<=90?4:recency<=180?3:recency<=270?2:1;
 const f=n===0?0:score(frequencies,n),m=n===0?0:score(monetaries,a.money);
 const segment=n===0?'Inactive':rs>=4&&f>=4&&m>=4?'Champions':rs<=2?'At risk':rs>=4?'Recent developing':'Other active';
 assert(Number(r.frequency)===n&&cents(r.monetary)===a.money&&r.recency_days===(recency===null?'':String(recency)),'RFM base mismatch');
 assert(Number(r.r_score)===rs&&Number(r.f_score)===f&&Number(r.m_score)===m&&r.segment===segment,'RFM score/segment mismatch');
}
const support=[...rows('SQLSupportControls')].map(c=>({c,ar:0n,overdue:0n,credit:0n,billings:0n,scheduled:0n,completed:0n,effective:0n,calls:new Set(),converted:new Set(),days:new Set(),outlets:0n,monetary:0n}));
const matches=(r,c)=>['branch_id','outlet_id','channel'].every(k=>!c[k]||r[k]===c[k]);
for(const r of rows('FactInvoiceSnapshot'))for(const a of support)if(matches(r,a.c)){a.ar+=cents(r.outstanding);if(Number(r.days_past_due)>0)a.overdue+=cents(r.outstanding);if(r.payment_method==='Credit')a.credit+=cents(r.outstanding);a.billings+=cents(r.credit_billings_90d);}
for(const r of rows('FactVisits'))if(r.visit_date>=context.period_start&&r.visit_date<=context.period_end)for(const a of support)if(matches(r,a.c)){a.scheduled++;if(r.is_completed==='t'){a.completed++;a.calls.add(r.call_key);a.days.add(r.salesman_day_key);if(r.is_effective_call==='t')a.effective++;if(r.has_attributed_order==='t')a.converted.add(r.call_key);}}
for(const r of rows('FactRFM'))for(const a of support)if(matches(r,a.c)){a.outlets++;a.monetary+=cents(r.monetary);}
for(const a of support){const values={outstanding_ar:[a.ar,100n],overdue_ar:[a.overdue,100n],credit_ar:[a.credit,100n],credit_billings_90d:[a.billings,100n],overdue_rate:ratio(a.overdue,a.ar),dso_90d:context.dso_history_complete==='t'?ratio(a.credit*90n,a.billings):null,scheduled_visits:[a.scheduled,1n],completed_visits:[a.completed,1n],effective_calls:[a.effective,1n],visit_completion_rate:ratio(a.completed,a.scheduled),effective_call_rate:ratio(a.effective,a.completed),order_conversion_rate:ratio(BigInt(a.converted.size),BigInt(a.calls.size)),calls_per_salesman_day:ratio(a.completed,BigInt(a.days.size)),rfm_outlets:[a.outlets,1n],rfm_monetary:[a.monetary,100n]};
 for(const [k,value] of Object.entries(values)){const unit=k.endsWith('_rate')?'ratio':k.endsWith('_visits')||k==='effective_calls'||k==='rfm_outlets'?'visits':k==='dso_90d'?'days':'IDR';const want=a.c[k]===''?null:fraction(a.c[k]);compare(value,want,`${a.c.case_id} ${k}`,unit);if(k!=='credit_billings_90d')expected.set(`${a.c.case_id}|${k}`,{value:want,unit});}
}
const template=['case_id,metric_id,actual_value'];for(const key of expected.keys()){const [c,m]=key.split('|');template.push(`${c},${m},`);}
fs.writeFileSync(path.join(evidence,'native-results.template.csv'),template.join('\n')+'\n');
const cli=process.argv.slice(2);
assert(cli.length===0||(cli.length===2&&cli[0]==='--native'),'Usage: node excel/verify_exports.mjs [--native capture.csv]');
let native='NOT RUN';const nativePath=cli.length?path.resolve(cli[1]):path.join(evidence,'native-results.csv');
if(cli.length)assert(fs.existsSync(nativePath),'Requested native capture does not exist');
if(fs.existsSync(nativePath)){
 const records=[...csv(fs.readFileSync(nativePath,'utf8'))];assert(records.shift().join(',')==='case_id,metric_id,actual_value','Native result headers');const seen=new Set();
 for(const row of records){assert(row.length===3,'Native result column count');const [c,m,v]=row;const key=c+'|'+m;assert(expected.has(key)&&!seen.has(key),'Unknown/duplicate native result');seen.add(key);const want=expected.get(key);compare(v===''?null:fraction(v),want.value,`NATIVE ${key}`,want.unit);}
 assert(seen.size===expected.size,'Incomplete native results');native='NUMERIC CAPTURE MATCHED; refresh/filter evidence still requires review';
}
const report={status:'PASS: CSV/source reconciliation only',context,counts,filter_cases:cases.length-1,support_filter_cases:support.length,comparisons,native_excel:native,limitations:'M and DAX are not executed by this validator. Native model, refresh, slicers, visual QA, saved workbook and reopen still pending.'};
fs.writeFileSync(path.join(evidence,'csv-verification.json'),JSON.stringify(report,null,2)+'\n');
console.log(JSON.stringify(report,null,2));
