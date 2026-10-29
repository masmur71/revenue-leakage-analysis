import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import {execFileSync} from 'node:child_process';
import {csv} from '../../excel/check_dashboard_capture.mjs';
const root=path.resolve(import.meta.dirname,'../..');
process.chdir(root);
const out='docs/qa'; fs.mkdirSync(out,{recursive:true});
const assert=(ok,message)=>{if(!ok)throw Error(message);};
const commands=[
 ['excel/verify_exports.mjs','--native','excel/dashboard-runs/20261029-225202-verification/native-results.partial.csv'],
 ['excel/review_partial_capture.mjs','excel/dashboard-runs/20261029-225202-verification'],
 ['excel/check_dashboard_capture.mjs','--cohort-samples','excel/dashboard-runs/20261030-003236-cohort-samples/cohort-samples.csv'],
 ...['ORIGINAL','ALTERNATE','RESTORED'].map(s=>['excel/check_dashboard_capture.mjs','--refresh',`excel/dashboard-runs/20261030-003422-refresh/refresh-${s}.csv`])
];
let logs='';
for(const args of commands){const result=execFileSync(process.execPath,args,{encoding:'utf8'});logs+=`COMMAND: node ${args.join(' ')}\n${result}\n`;}
fs.writeFileSync(`${out}/artifact-checks.txt`,logs.trimEnd()+'\n');
const hash=p=>crypto.createHash('sha256').update(fs.readFileSync(p)).digest('hex');
const prior=JSON.parse(fs.readFileSync('docs/excel/final-workbook-integrity.json','utf8'));
assert(hash('excel/RevenueLeakage.xlsx')===prior.sha256,'Accepted workbook changed');
const report=JSON.parse(fs.readFileSync(`${out}/reproduction.json`,'utf8').replace(/^\uFEFF/,''));
assert(report.status==='PASS','Database reproduction not complete');
assert(report.workbook_sha256===prior.sha256,'QA workbook hash differs');
const tracked=execFileSync('git',['ls-files','-z'],{encoding:'utf8'}).split('\0').filter(Boolean);
const staged=execFileSync('git',['ls-files','--others','--exclude-standard','-z'],{encoding:'utf8'}).split('\0').filter(Boolean);
const files=[...new Set([...tracked,...staged])];
const forbidden=/^(?:AGENTS\.md|STATE\.md|archieve\/|skills\/|ai-workflow\/|\.env(?:\.|$)|.*\.(?:pem|p12|pfx|key)$)/i;
assert(!files.some(p=>forbidden.test(p)),'Excluded local file or credential container in deliverables');
const patterns=[/-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----/,/\bgh[pousr]_[A-Za-z0-9]{30,}/,/\bgithub_pat_[A-Za-z0-9_]{40,}/,/\bAKIA[0-9A-Z]{16}\b/,/\bsk-(?:proj-)?[A-Za-z0-9_-]{32,}/,/(?:postgres(?:ql)?|mysql):\/\/[^\s/:]+:[^\s/@]+@/i,/(?:password|passwd|pwd|api_key|access_token)\s*[:=]\s*["'][^"'\r\n]{4,}["']/i];
let scanned=0;
for(const p of files){if(!/\.(?:md|sql|ps1|mjs|json|jsonl|txt|csv|m|dax|bas)$/.test(p))continue;const text=fs.readFileSync(p,'utf8');scanned++;for(const re of patterns)assert(!re.test(text),`Potential credential detected in ${p}; inspect locally`);}
const docs=['README.md','docs/data_dictionary.md','docs/erd.md','docs/refresh.md','docs/excel_model.md','docs/sql_catalog.md','docs/final_qa.md'];
let links=0;
for(const p of docs){const text=fs.readFileSync(p,'utf8');for(const m of text.matchAll(/\]\(([^)]+)\)/g)){const target=m[1].split('#')[0];if(!target||/^https?:/.test(target))continue;assert(path.resolve(path.dirname(p),target)===path.resolve(out,'artifacts.json')||fs.existsSync(path.resolve(path.dirname(p),target)),`Broken link in ${p}: ${target}`);links++;}}
const ddl=fs.readFileSync('sql/00_setup/analytics_schema.sql','utf8');
const names=[...ddl.matchAll(/CREATE TABLE analytics\.(\w+)/g)].map(m=>m[1]);
const dict=fs.readFileSync('docs/data_dictionary.md','utf8'),erd=fs.readFileSync('docs/erd.md','utf8');
assert(names.length===13&&names.every(n=>dict.includes(`## ${n}`)&&erd.includes(n)),'Incomplete source schema documentation');
assert((erd.match(/\|\|--o\{/g)||[]).length===31,'ERD foreign-key count differs');
const summary=csv(fs.readFileSync('docs/optimization/summary.csv','utf8')).slice(1);
assert(summary.length===18,'Optimization summary coverage differs');
for(const stage of ['baseline','call_index','unhelpful_index']){const rows=fs.readFileSync(`docs/optimization/${stage}.jsonl`,'utf8').trim().split(/\r?\n/).map(s=>JSON.parse(s));assert(rows.filter(r=>r.kind==='plan').length===18,`Plan coverage differs: ${stage}`);assert(rows.filter(r=>r.kind==='equality').length===3&&rows.filter(r=>r.kind==='equality').every(r=>r.mismatched_rows===0),`Optimization equality differs: ${stage}`);}
const partial=JSON.parse(fs.readFileSync('docs/excel/partial-capture-review.json','utf8'));
assert(partial.main.numeric_matched_rows===282&&partial.extended.matched_rows===131,'Native capture coverage differs');
const result={status:'PASS',date:new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Jakarta',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date()),timezone:'Asia/Jakarta',workbook_sha256:prior.sha256,workbook_unchanged:true,native_main:282,native_extended:131,cohort_samples:10,refresh_values:138,source_tables:13,foreign_keys:31,local_links:links,tracked_and_proposed_text_files_scanned:scanned,credential_scan:'No supported credential patterns detected; heuristic text scan, not proof of absence in every binary or Git history.',excluded_files:'No forbidden local paths in tracked/proposed deliverables',optimization_plans:54,limits:'Rechecked stored captures only; no new Excel refresh/UI/reopen test. Final Phase 4 operator attribution remains unchanged.'};
fs.writeFileSync(`${out}/artifacts.json`,JSON.stringify(result,null,2)+'\n');
console.log(`PASS: workbook unchanged; 561 native captured values reconcile; ${links} links valid; ${scanned} text files scanned; 54 stored optimization plans.`);
