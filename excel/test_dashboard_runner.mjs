// Pure-core and contract tests. Does not automate Excel or prove native VBA compilation.
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import assert from 'node:assert/strict';
import {spawnSync} from 'node:child_process';
import {csv,round,compareCaptures} from './check_dashboard_capture.mjs';
const root=path.resolve(import.meta.dirname,'..'), dir=path.join(root,'excel','dashboard-specs');
const text=fs.readFileSync(path.join(root,'excel','DashboardRunner.bas'),'utf8');
let checks=0;const check=(condition,label)=>{assert(condition,label);checks++;};
const read=n=>csv(fs.readFileSync(path.join(dir,n+'.csv'),'utf8')).slice(1);
const cases=read('cases'), requests=read('requests'), expected=read('extended-expected');
check(new Set(cases.map(r=>r[0])).size===cases.length,'Unique cases');
check(cases.every(r=>r.length===12),'Case column count');
check(read('tables').length===14&&read('tables').every(r=>/^\d+$/.test(r[1])),'14 source-count contracts');
const model=JSON.parse(fs.readFileSync(path.join(root,'excel','model_contract.json'),'utf8'));
for(const [,field] of read('slicers')){const [table,column]=field.split('.');check(model[table]?.columns.some(c=>c.name===column),'Slicer field exists');}
const measures=new Set(fs.readFileSync(path.join(root,'excel','measures.dax'),'utf8').split(/\r?\n/).filter(l=>l.includes(' := ')).map(l=>l.split(' := ')[0]));
for(const [name,sheet,anchor,rows,cols,metrics] of read('pivots')){
 check(sheet.length<=31&&/^[A-Z]+[1-9]\d*$/.test(anchor),'Valid pivot destination '+name);
 for(const field of (rows+'|'+cols).split('|').filter(Boolean)){const [table,column]=field.split('.');check(model[table]?.columns.some(c=>c.name===column),'Pivot dimension exists');}
 for(const metric of metrics.split('|'))check(measures.has(metric),'Explicit pivot measure exists');
}
const ids=new Set(cases.map(r=>r[0]));for(const [id,metric,dest] of requests){check(ids.has(id)&&measures.has(metric)&&['main','extended'].includes(dest),'Request references exist');}
const template=csv(fs.readFileSync(path.join(root,'docs','excel','native-results.template.csv'),'utf8')).slice(1);
check(JSON.stringify(requests.filter(r=>r[2]==='main').map(r=>r.slice(0,2)))===JSON.stringify(template.map(r=>r.slice(0,2))),'Main capture exact template coverage');
check(requests.filter(r=>r[2]==='extended').length===expected.length,'Extended coverage');
check(expected.filter(r=>r[0].startsWith('cohort_')&&r[1]==='retention_rate').length>200,'All observed cohort cells and future guard');
check(read('return-reasons').every(r=>r[0]!==''),'No blank return reason');
const sampleCases=read('cohort-samples'),sampleExpected=read('cohort-samples-expected'),sampleRequests=read('cohort-sample-requests');
check(sampleCases.length===4&&sampleExpected.length===10&&sampleRequests.length===10,'Only four remaining cohort cases and ten values');
check(sampleCases.every(r=>r[1]===''&&r[2]===''),'Cohort samples do not repeat irrelevant operating-date filters');
check(sampleRequests.every(([c,m])=>sampleCases.some(r=>r[0]===c)&&['cohort_outlets','cohort_active_outlets','retention_rate'].includes(m)),'Dedicated three-measure sample coverage');
check(sampleExpected.filter(r=>r[0]==='cohort_multi_month').every(r=>r[2]===''),'Cross-month cohort totals remain BLANK');
check(sampleExpected.filter(r=>r[0].startsWith('cohort_2025')).every(r=>expected.some(e=>JSON.stringify(e)===JSON.stringify(r))),'Observed samples reuse existing SQL cell controls');
check(text.includes('Private Sub VerifyExhaustive()')&&text.includes('Run Phase4_CheckCohortSamples for the remaining four small cases'),'Accidental full-suite rerun guarded');
check(fs.readFileSync(path.join(dir,'manifest.json'),'utf8')===fs.readFileSync(path.join(root,'data','excel','manifest.json'),'utf8'),'Manifest snapshot current');
check(!/\bActiveWorkbook\b|\bVBProject\b|\.Delete\b|\bShell\b|\bDisplayAlerts\b/.test(text),'No foreground binding, deletion, security or shell bypass');
check(!/ModelMeasures\.Add|ModelRelationships\.Add|\.Save\b/.test(text),'No model edits or automatic target save');
check(text.indexOf('PreflightDashboard\n    BackupTarget')>0,'Dashboard preflight before backup/mutation');
check(text.includes('AssertProbeIsolated p'),'Native probe isolated from slicers');
check(text.includes('EMPTY_PIVOT_ADOPTED')&&text.includes('p.DataFields.Count <> 0 Or p.RowFields.Count <> 0 Or p.ColumnFields.Count <> 0 Or p.PageFields.Count <> 0'),'Empty pivot adoption rejects populated layouts');
check(text.includes('p.PivotCache.WorkbookConnection.Name <> wb.Model.DataModelConnection.Name')&&text.includes('p.TableRange2.Row <> anchor.Row Or p.TableRange2.Column <> anchor.Column'),'Empty adoption requires exact anchor and model connection');
check(!text.includes('extended-expected.csv')&&!text.includes('SQLReference.csv'),'Capture code never reads expected numeric controls');
check(text.includes('wb.Queries("pFolder").Formula = original')&&text.includes('RECOVERY FAILED'),'Refresh parameter restoration/error reporting');
check(text.includes('vbDefaultButton2'),'Cloud identity defaults No');
check(text.includes('SlicerCaches.Add(wb.Model.DataModelConnection, Hierarchy(CStr(row(1))))')&&text.includes('levelName = sc.SlicerCacheLevels(1).Name')&&text.includes('Set view = sc.Slicers.Add(sh, levelName)'),'Explicit model source, actual OLAP level and minimal positional creation arguments');
check(text.includes('If sc.Slicers.Count = 0 Then')&&text.includes('SLICER_VIEW_REQUEST'),'Resume a cache without a view with detailed operation logs');
check(text.indexOf('LogEvent "SLICER_CONNECTED"')<text.indexOf('Set view = sc.Slicers.Add(sh, levelName)'),'Connect pivot tables before view creation');
check(text.indexOf('LogEvent "SLICER_VIEW_CREATED"')<text.indexOf('view.Caption ='),'Separate creation from layout');
check(text.includes('Public Sub Phase4_FinishDashboard()'),'Resume finishing without rebuilding pivot layouts');
check(text.includes('If ReadText(path) = text Then')&&text.includes('SNAPSHOT_REUSED')&&text.includes('Snapshot already exists with different contents'),'Repeated inventory reuses identical immutable snapshots and rejects conflicts');
check(text.includes('Set placement = SlicerPlacement(sh, CLng(row(2)))')&&text.includes('view.Shape.Placement = xlFreeFloating'),'Measure actual slicer geometry and prevent row-resize movement');
check(!text.includes('native-results.csv" For Output'),'Final captures only published after complete loop');
const names=[...text.matchAll(/^(?:Public|Private) (?:Sub|Function) (\w+)/gm)].map(m=>m[1]);
check(new Set(names.map(n=>n.toLowerCase())).size===names.length,'Unique procedures');
check(names.length===[...text.matchAll(/^End (?:Sub|Function)/gm)].length,'Balanced procedures');
check(Math.max(...text.split('\n').map(s=>s.length))<1023,'VBA physical line lengths');
check(round('1.005',2)===101n&&round('-1.005',2)===-101n,'Half away from zero');
check(round('0.12345',4)===1235n,'Percentage display rounding');
assert.throws(()=>compareCaptures([['count','orders','10','orders']],[['count','orders','10.1']]));checks++;
assert.throws(()=>compareCaptures([['money','net_sales','1.00','IDR']],[['money','net_sales','1.01']]));checks++;
const simulated=expected.map(([c,m,v])=>[c,m,v]);
check(compareCaptures(expected,simulated)===expected.length,'Synthetic comparator acceptance (not native evidence)');
for(const bad of [simulated.slice(1),simulated.concat([simulated[0]]),simulated.map((r,i)=>i===0?[r[0],r[1],r[2]===''?'0':'']:r)]){
 assert.throws(()=>compareCaptures(expected,bad));checks++;
}
const proc=name=>{const m=text.match(new RegExp(`Private (?:Function|Sub) ${name}\\b[\\s\\S]*?^End (?:Function|Sub)\\s*$`,'m'));assert(m,'Missing '+name);return m[0];};
let core=['SameSet','FieldNames','MeasureNames','Hierarchy','MemberName','IsoDate','ExpandNumber','InvariantValue','FolderLiteral','MFolder','Csv','ReadCsv','SafeSlicerTop','RectanglesOverlap','IsNameCharacter','IsPlainName','QuotedEnd','CanonicalFormula'].map(proc).join('\n');
core=core.replace(/(\w+) As New Collection/g,'$1: Set $1 = New Collection').replace(/\bPrivate (Function|Sub)/g,'$1').replace(/ As (?:String|Long|Variant|Date|Boolean|Collection|Double)\b/g,'').replace(/\b(Left|Right|Mid|Trim|LCase|UCase|String|Str|Format)\$/g,'$1').replace(/\bNext [A-Za-z_]\w*/g,'Next');
// VBScript has no Format function; ISO date validity is checked separately in JS below.
core=core.replace(/Format\(IsoDate, "yyyy-mm-dd"\)/g,'IsoText(IsoDate)');
core=core.replace(/ReDim data\(0 To /g,'ReDim data(');
const script=`Option Explicit
Dim tests: tests = 0
Class Collection
 Private values
 Private Sub Class_Initialize()
  Set values = CreateObject("Scripting.Dictionary")
 End Sub
 Public Sub Add(value)
  values.Add values.Count + 1, value
 End Sub
 Public Property Get Count()
  Count = values.Count
 End Property
 Public Default Property Get Item(index)
  Item = values.Item(index)
 End Property
End Class
Function ReadText(text)
 ReadText = text
End Function
Sub Check(condition, label)
 If Not condition Then Err.Raise 513, "test", label
 tests = tests + 1
End Sub
Function IsoText(d)
 IsoText = Year(d) & "-" & Right("0" & Month(d),2) & "-" & Right("0" & Day(d),2)
End Function
Function IsError(value)
 IsError = (VarType(value) = 10)
End Function
Function Str(value)
 Str = Replace(CStr(value), ",", ".")
End Function
${core}
Call Check(SameSet("|a|b", "|b|a"), "Set equivalence")
Call Check(Not SameSet("|a|b", "|a"), "Reject extra field")
Call Check(Hierarchy("DimDate.year") = "[DimDate].[year]", "Hierarchy")
Call Check(MemberName("DimBranch.branch_id", "x]y") = "[DimBranch].[branch_id].&[x]]y]", "Escape member")
Call Check(FieldNames("DimDate.year|DimBranch.branch_id") = "|[DimDate].[year]|[DimBranch].[branch_id]", "Fields")
Call Check(MeasureNames("net_sales") = "|[Measures].[net_sales]", "Measures")
Call Check(ExpandNumber(".24") = "0.24", "Leading zero")
Call Check(ExpandNumber("-.24") = "-0.24", "Negative fraction")
Call Check(ExpandNumber("1.234E-05") = "0.00001234", "Small exponent")
Call Check(ExpandNumber("-1.234E+05") = "-123400", "Large negative")
Call Check(ExpandNumber("0") = "0", "True zero")
Call Check(InvariantValue(Empty) = "", "True BLANK")
Call Check(InvariantValue(0) = "0", "Numeric zero")
Call Check(IsoText(IsoDate("2024-02-29")) = "2024-02-29", "Leap date")
Call Check(FolderLiteral(MFolder("C:\\data")) = "C:\\data", "M parameter roundtrip")
Call Check(CanonicalFormula("= SUM ( 'FactTargets'[target_net_sales] )") = CanonicalFormula("SUM(FactTargets[target_net_sales])"), "Formula canonicalization")
Dim parsed, parsedRow, errorNumber
Set parsed = ReadCsv("a,b" & vbCrLf & "x," & Csv("quoted,""value") & vbCrLf)
parsedRow = parsed(1)
Call Check(parsed.Count = 1 And parsedRow(1) = "quoted,""value", "CSV quoting roundtrip")
Set parsed = ReadCsv(ChrW(&HFEFF) & "a,b" & vbLf & "x,")
parsedRow = parsed(1)
Call Check(parsedRow(1) = "", "CSV BOM and trailing BLANK")
On Error Resume Next
Set parsed = ReadCsv("a,b" & vbLf & "wrong")
errorNumber = Err.Number: Err.Clear
On Error GoTo 0
Call Check(errorNumber <> 0, "Reject malformed CSV width")
On Error Resume Next
Set parsed = ReadCsv("a,b" & vbLf & Chr(34) & "unclosed,x")
errorNumber = Err.Number: Err.Clear
On Error GoTo 0
Call Check(errorNumber <> 0, "Reject unclosed CSV quote")
On Error Resume Next
parsedRow = IsoDate("2025-02-29")
errorNumber = Err.Number: Err.Clear
On Error GoTo 0
Call Check(errorNumber <> 0, "Reject invalid ISO date rollover")
Call Check(SafeSlicerTop(120, 150) = 174, "Slicer collision gets point-based gap")
Call Check(SafeSlicerTop(300, 150) = 300, "Existing larger gap preserved")
Call Check(SafeSlicerTop(120.5, 150.25) = 174.25, "Fractional point geometry")
Call Check(Not RectanglesOverlap(0,174,220,150,0,0,220,150), "24-point gap has no collision")
Call Check(Not RectanglesOverlap(0,150,220,150,0,0,220,150), "Touching edges have no collision")
Call Check(RectanglesOverlap(0,149,220,150,0,0,220,150), "Actual one-point overlap rejected")
WScript.Echo "PASS: " & tests & " VBA-compatible pure-core assertions; Excel not opened."
`;
const temp=fs.mkdtempSync(path.join(os.tmpdir(),'dashboard-runner-test-')),file=path.join(temp,'core.vbs');
try{
 fs.writeFileSync(file,script,'utf8');
 const result=spawnSync('C:\\Windows\\System32\\cscript.exe',['//nologo',file],{encoding:'utf8',timeout:30000,windowsHide:true});
 if(result.error)throw result.error;
 const near=(result.stdout+'\n'+result.stderr).match(/\((\d+),\s*\d+\)/);
 const rendered=near?script.split(/\r?\n/).slice(Math.max(0,Number(near[1])-2),Number(near[1])+1).join('\n'):'';
 assert.equal(result.status,0,result.stdout+'\n'+result.stderr+'\n'+rendered);
 assert.match(result.stdout,/PASS: 27 VBA-compatible pure-core assertions/,`Windows Script Host did not confirm execution. stderr=${result.stderr}`);
 process.stdout.write(result.stdout);
}finally{fs.rmSync(file,{force:true});fs.rmdirSync(temp);}
console.log(`PASS: ${checks} contract/validation assertions. PENDING: native VBA compile, dashboard execution, refresh and reconciliation.`);
