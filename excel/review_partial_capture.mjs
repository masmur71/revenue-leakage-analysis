// Read existing user-run native captures; never opens Excel or resumes a macro.
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import {csv,compareCaptures} from './check_dashboard_capture.mjs';
const root=path.resolve(import.meta.dirname,'..');
const run=process.argv[2];if(!run)throw Error('Usage: node excel/review_partial_capture.mjs <run-folder>');
const folder=path.resolve(run);
const read=p=>csv(fs.readFileSync(p,'utf8'));
const [eventHeader,...events]=read(path.join(folder,'events.csv'));
if(eventHeader.join(',')!=='timestamp,event,object,detail')throw Error('Unexpected event schema');
const completed=new Set(events.filter(r=>r[1]==='CASE_CAPTURED').map(r=>r[2]));
const [headers,...actual]=read(path.join(folder,'extended-results.partial.csv'));
if(headers.join(',')!=='case_id,metric_id,actual_value')throw Error('Unexpected capture schema');
const expected=read(path.join(root,'excel','dashboard-specs','extended-expected.csv')).slice(1);
const controls=new Map(expected.map(r=>[r[0]+'|'+r[1],r]));
const seen=new Set(),failures=[],uncommitted=[];
for(const r of actual){const key=r[0]+'|'+r[1];if(seen.has(key)||!controls.has(key))throw Error('Duplicate/unknown capture: '+key);seen.add(key);if(!completed.has(r[0]))uncommitted.push(key);try{compareCaptures([controls.get(key)],[r]);}catch(e){failures.push({key,error:e.message});}}
const missing=expected.filter(r=>!seen.has(r[0]+'|'+r[1]));
const completedExtendedCases=[...new Set(actual.map(r=>r[0]))].filter(c=>completed.has(c)&&expected.filter(r=>r[0]===c).every(r=>seen.has(c+'|'+r[1])));
const [mainHeaders,...main]=read(path.join(folder,'native-results.partial.csv'));
if(mainHeaders.join(',')!=='case_id,metric_id,actual_value')throw Error('Unexpected main capture schema');
const mainKeys=new Set(main.map(r=>r[0]+'|'+r[1]));
const template=read(path.join(root,'docs','excel','native-results.template.csv')).slice(1);
const mainComplete=main.length===template.length&&mainKeys.size===main.length&&template.every(r=>mainKeys.has(r[0]+'|'+r[1])&&completed.has(r[0]));
const ref=read(path.join(root,'data','excel','SQLReference.csv')).slice(1),units=new Map(ref.map(([m,,,,u])=>[m,u]));
units.set('rfm_outlets','outlets');units.set('rfm_monetary','IDR');
const reference=new Map(ref.map(([m,,,v])=>[m,v])),caseValues=new Map();
for(const file of ['SQLFilterControls','SQLSupportControls']){const [h,...rows]=read(path.join(root,'data','excel',file+'.csv'));for(const r of rows)caseValues.set(r[0],Object.fromEntries(h.map((k,i)=>[k,r[i]])));}
const mainExpected=template.map(([c,m])=>{const value=c==='reference'?reference.get(m):caseValues.get(c)?.[m];if(value===undefined||!units.has(m))throw Error('Missing main control: '+c+'|'+m);return [c,m,value,units.get(m)];});
const mainMatched=compareCaptures(mainExpected,main);
const hash=name=>crypto.createHash('sha256').update(fs.readFileSync(path.join(folder,name))).digest('hex');
const report={run:path.relative(root,folder).replaceAll('\\','/'),action:'Read existing captures only; no Excel or VBA execution',capture_sha256:{main:hash('native-results.partial.csv'),extended:hash('extended-results.partial.csv'),events:hash('events.csv')},main:{rows:main.length,required_rows:template.length,all_expected_keys_and_case_checkpoints:mainComplete,numeric_matched_rows:mainMatched},extended:{captured_rows:actual.length,matched_rows:actual.length-failures.length,failures,uncommitted_rows:uncommitted,complete_cases:completedExtendedCases,cohort_observed_cells:completedExtendedCases.filter(c=>c.startsWith('cohort_')&&c!=='cohort_future').length,remaining_rows:missing.length,remaining_cases:[...new Set(missing.map(r=>r[0]))]},limitations:'Interrupted full verification is not complete. Matching saved rows do not establish refresh recovery, visual QA, slicer restoration or save/reopen acceptance.'};
fs.writeFileSync(path.join(root,'docs','excel','partial-capture-review.json'),JSON.stringify(report,null,2)+'\n');
console.log(JSON.stringify({...report,extended:{...report.extended,remaining_cases:report.extended.remaining_cases.length}},null,2));
if(failures.length||uncommitted.length||!mainComplete)process.exitCode=1;
