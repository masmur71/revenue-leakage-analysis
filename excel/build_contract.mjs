// Generate explicit import schemas and M queries. No workbook is authored here.
import fs from 'node:fs/promises';
import path from 'node:path';
const root = path.resolve(import.meta.dirname, '..');
const specs = {
 DimBranch: ['branch_id','branch_name','region_id','region_name','city','province'],
 DimOutlet: ['outlet_id','outlet_name','channel','onboard_date:date'],
 DimProduct: ['sku_id','sku_name','category'],
 DimSalesman: ['salesman_id','salesman_name','home_branch_id'],
 DimChannel: ['channel'],
 DimDate: ['date:date','year:int','quarter:int','month:int','month_name','year_month'],
 FactCommercial: ['event_id','event_type','event_date:date','branch_id','outlet_id','sku_id','salesman_id','channel','sales_order_id','sales_units:int','returned_units:int','reference_sales:money','gross_sales:money','discount_amount:money','billed_sales:money','cogs_sold:money','price_shortfall:money','price_upside:money','invalid_promo_discount:money','promo_lines:int','invalid_promo_lines:int','mismatch_lines:int','mismatch_sales:money','return_amount:money','returned_cogs:money','return_reason'],
 FactVisits: ['visit_id','visit_date:date','branch_id','outlet_id','salesman_id','channel','is_completed:bool','is_effective_call:bool','has_attributed_order:bool','call_key','salesman_day_key'],
 FactInvoiceSnapshot: ['invoice_id','invoice_date:date','branch_id','outlet_id','salesman_id','channel','payment_method','invoice_amount:money','allocated_paid:money','outstanding:money','overpayment:money','applicable_term_days:int','contractual_due_date:date','days_past_due:int','credit_billings_90d:money'],
 FactTargets: ['target_id','branch_id','month_start:date','target_net_sales:money'],
 FactPayments: ['payment_id','payment_date:date','branch_id','outlet_id','salesman_id','channel','payment_amount:money','allocation_status'],
 FactRFM: ['outlet_id','branch_id','channel','recency_days:int','frequency:int','monetary:money','r_score:int','f_score:int','m_score:int','segment','rfm_history_complete:bool'],
 FactCohort: ['cohort_cell_id','outlet_id','branch_id','channel','cohort_month:date','activity_month:date','month_age:int','active_outlet:int'],
 ReportContext: ['period_start:date','period_end:date','as_of_date:date','observed_rfm_start:date','rfm_history_complete:bool','dso_history_complete:bool'],
 SQLReference: ['metric_id','numerator:number','denominator:number','value:number','unit']
};
const keys = {DimBranch:'branch_id',DimOutlet:'outlet_id',DimProduct:'sku_id',DimSalesman:'salesman_id',DimChannel:'channel',DimDate:'date',FactCommercial:'event_id',FactVisits:'visit_id',FactInvoiceSnapshot:'invoice_id',FactTargets:'target_id',FactPayments:'payment_id',FactRFM:'outlet_id',FactCohort:'cohort_cell_id',SQLReference:'metric_id'};
const types = {text:'type text',int:'Int64.Type',date:'type date',money:'Currency.Type',number:'type number',bool:'type logical'};
const dir = path.join(root,'excel','power_query');
await fs.mkdir(dir,{recursive:true});
await fs.writeFile(path.join(dir,'pFolder.m'),'"C:\\Users\\masmu\\OneDrive\\Documents\\PersProject\\leakageanalysis\\data\\excel" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]\n');
const contract = {};
for (const [name, fields] of Object.entries(specs)) {
 const columns = fields.map(f=>{const [name,type='text']=f.split(':');return {name,type};});
 contract[name]={key:keys[name]??null,columns,load:name==='SQLReference'?'worksheet-only':'model-only'};
 const schema = columns.map(c=>`{"${c.name}", ${types[c.type]}}`).join(', ');
 await fs.writeFile(path.join(dir,`${name}.m`),`fnLoadCsv("${name}.csv", {${schema}}, ${keys[name]?`"${keys[name]}"`:'null'})\n`);
}
await fs.writeFile(path.join(root,'excel','model_contract.json'),JSON.stringify(contract,null,2)+'\n');
console.log(`PASS: generated ${Object.keys(contract).length} typed queries and folder parameter.`);
