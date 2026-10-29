// Compare extended native captures using exact decimal arithmetic and display rounding.
// Successful comparison is numeric evidence only, not proof of native refresh or visual QA.
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
export function csv(text){const rows=[];let row=[],field='',quoted=false;text=text.replace(/^\uFEFF/,'');for(let i=0;i<text.length;i++){const c=text[i];if(c==='"'){if(quoted&&text[i+1]==='"'){field+='"';i++;}else quoted=!quoted;}else if(c===','&&!quoted){row.push(field);field='';}else if(c==='\n'&&!quoted){row.push(field.replace(/\r$/,''));rows.push(row);row=[];field='';}else field+=c;}if(quoted)throw Error('Unclosed CSV quotation');if(field.length||row.length){row.push(field.replace(/\r$/,''));rows.push(row);}return rows;}
export function decimal(s){if(!/^-?\d+(\.\d+)?$/.test(s))throw Error(`Invalid invariant decimal: ${s}`);const [a,b='']=s.split('.');return [BigInt(a+b),10n**BigInt(b.length)];}
export function round(value,places){let [n,d]=decimal(value);const sign=n<0n?-1n:1n;if(n<0n)n=-n;return sign*((n*10n**BigInt(places)*2n+d)/(2n*d));}
export function compareCaptures(expectedRows,actualRows){
 const expected=new Map(),seen=new Set();
 for(const r of expectedRows){if(r.length!==4)throw Error('Expected-control column count');const [c,m,v,u]=r,key=c+'|'+m;if(expected.has(key))throw Error('Duplicate expected key');expected.set(key,{v,u});}
 for(const r of actualRows){if(r.length!==3)throw Error('Capture column count');const [c,m,v]=r,key=c+'|'+m,w=expected.get(key);if(!w||seen.has(key))throw Error(`Unknown/duplicate capture: ${key}`);seen.add(key);if(v===''||w.v===''){if(v!==w.v)throw Error(`BLANK mismatch: ${key}`);}else if(['outlets','orders','visits','units'].includes(w.u)){const [a,b]=decimal(v),[c,d]=decimal(w.v);if(a*d!==c*b)throw Error(`Exact count mismatch: ${key}`);}else {const places=w.u==='ratio'?4:2;if(round(v,places)!==round(w.v,places))throw Error(`Rounded mismatch: ${key}: native=${v}, expected=${w.v}`);}}
 if(seen.size!==expected.size)throw Error(`Incomplete captures: ${seen.size}/${expected.size}`);
 return seen.size;
}
if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url)){
 const refresh=process.argv[2]==='--refresh',samples=process.argv[2]==='--cohort-samples',flag=refresh||samples;const file=process.argv[flag?3:2];if(!file||process.argv.length!==(flag?4:3))throw Error('Usage: node excel/check_dashboard_capture.mjs [--refresh|--cohort-samples] <capture.csv>');
 let expected;
 if(refresh){const [headers,...rows]=csv(fs.readFileSync(new URL('../data/excel/SQLReference.csv',import.meta.url),'utf8'));if(headers.join(',')!=='metric_id,numerator,denominator,value,unit')throw Error('Unexpected SQL reference schema');expected=[['case_id','metric_id','value','unit'],...rows.map(([m,,,v,u])=>['reference',m,v,u])];}
 else expected=csv(fs.readFileSync(new URL(samples?'./dashboard-specs/cohort-samples-expected.csv':'./dashboard-specs/extended-expected.csv',import.meta.url),'utf8'));
 const actual=csv(fs.readFileSync(file,'utf8'));
 if(expected.shift().join(',')!=='case_id,metric_id,value,unit'||actual.shift().join(',')!=='case_id,metric_id,actual_value')throw Error('Unexpected headers');
 const count=compareCaptures(expected,actual);
 console.log(`PASS: ${count} ${refresh?'refresh-reference':samples?'cohort-sample':'extended'} native numeric captures match. Native action logs, visual inspection and reopen still require review.`);
}
