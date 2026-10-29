// Run the installer's pure VBA-compatible core in Windows Script Host.
// Small mock collections/table schemas exercise preflight, never Excel or a workbook.
// This is NOT a VBA compilation test or native Data Model execution.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {spawnSync} from 'node:child_process';
import assert from 'node:assert/strict';
const root=path.resolve(import.meta.dirname,'..');
const moduleText=fs.readFileSync(path.join(root,'excel','MeasureInstaller.bas'),'utf8');
const dax=fs.readFileSync(path.join(root,'excel','measures.dax'),'utf8');
const contract=JSON.parse(fs.readFileSync(path.join(root,'excel','model_contract.json'),'utf8'));
const quote=s=>'"'+s.replaceAll('"','""')+'"';
const specs=dax.split(/\r?\n/).filter(s=>s.trim()&&!s.trim().startsWith('//')).map(s=>{
 const m=s.match(/^([a-z_0-9]+)\s*:=\s*(.+)$/);assert(m,`Malformed source: ${s}`);return [m[1],'='+m[2]];
});
assert.equal(specs.length,66);
assert.equal(new Set(specs.map(s=>s[0])).size,66);
const home=moduleText.match(/Private Const HOME_TABLE As String = "([^"]+)"/)[1];
assert(contract[home], 'Home table must already exist in the model contract');
assert(!specs.some(s=>contract[home].columns.some(c=>c.name===s[0])), 'Home column/measure collision');
const required=moduleText.match(/Private Const REQUIRED_TABLES As String = "([^"]+)"/)[1].split(',');
assert.equal(required.length,14);
assert(required.every(n=>contract[n]&&contract[n].load==='model-only'));
assert(!/\bActiveWorkbook\b/.test(moduleText),'Never bind by foreground workbook');
assert(!/\bVBProject\b|\bShell\b|\bRefreshAll\b|\bDisplayAlerts\b|\.Delete\b/.test(moduleText),'No macros/security/refresh/deletion automation');
assert(moduleText.indexOf('If conflicts > 0 Then')<moduleText.indexOf('target.SaveCopyAs'));
assert(moduleText.indexOf('target.SaveCopyAs')<moduleText.indexOf('target.Model.ModelMeasures.Add'));
assert(moduleText.includes('text = NativeAddFormula(CStr(spec(1)), CStr(Application.International(xlListSeparator)), CStr(Application.International(xlDecimalSeparator)))'), 'Native API conversion must receive actual locale settings');
assert(moduleText.includes('ModelMeasures.Add(currentName, home, text, info'), 'Native API must receive the converted expression');
assert(!/target\.Save\b/.test(moduleText),'Target save remains manual');
assert(moduleText.includes('Application.FileDialog(4)'), 'Cloud runner must support explicit local folder selection');
assert(moduleText.includes('vbDefaultButton2'), 'Cloud target confirmation must default to No');
assert(moduleText.indexOf('"CLOUD_TARGET_CONFIRMED"')<moduleText.indexOf('target.SaveCopyAs'), 'Cloud identity must be confirmed before backup/additions');
assert(Math.max(...moduleText.split('\n').map(s=>s.length))<1023,'VBA physical line limit');

function getProcedure(name){
 const pattern=new RegExp(`Private (Function|Sub) ${name}\\b[\\s\\S]*?^End (?:Function|Sub)\\s*$`,'m');
 const found=moduleText.match(pattern);assert(found,`Procedure missing: ${name}`);return found[0];
}
const coreNames=['NewDictionary','ParseDefinitions','ValidateReferences','BracketReferences','FormatKind','NativeAddFormula','IsCloudPath','IsNameCharacter','IsPlainName','ParseDefinitionLine','QuotedEnd','MaskStrings','CanonicalFormula','SelfTestCore'];
let core=coreNames.map(getProcedure).join('\n\n');
core=core.replace(/Dim (\w+) As New Collection/g,'Dim $1: Set $1 = New Collection')
 .replace(/\bPrivate (Function|Sub)/g,'$1')
 .replace(/ As (?:Object|Collection|String|Integer|Long|Boolean|Variant)\b/g,'')
 .replace(/\b(Left|Right|Mid|Trim|LCase|UCase|Space)\$/g,'$1')
 .replace(/\bNext [A-Za-z_][A-Za-z0-9_]*/g,'Next');
const lines=[`Option Explicit
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
Class MockColumn
 Public Name
End Class
Class MockTable
 Public Name
 Public ModelTableColumns
End Class
Function MakeColumn(name)
 Dim c: Set c = New MockColumn: c.Name = name: Set MakeColumn = c
End Function
Sub Check(condition, label)
 If Not condition Then Err.Raise 513, "test", label
 tests = tests + 1
End Sub
Sub ExpectLineFailure(line)
 Dim parsed, n
 On Error Resume Next
 parsed = ParseDefinitionLine(line)
 n = Err.Number: Err.Clear
 On Error GoTo 0
 Check n <> 0, "Expected parse rejection: " & line
End Sub
Sub ExpectDefinitionsFailure(text)
 Dim actual, n
 On Error Resume Next
 Set actual = ParseDefinitions(text)
 n = Err.Number: Err.Clear
 On Error GoTo 0
 Check n <> 0, "Expected duplicate rejection"
End Sub
Sub ExpectReferenceFailure(text, tables)
 Dim definitions, n
 Set definitions = ParseDefinitions(text)
 On Error Resume Next
 ValidateReferences definitions, tables
 n = Err.Number: Err.Clear
 On Error GoTo 0
 Check n <> 0, "Expected dependency/schema rejection"
End Sub
${core}
Dim parsed, definitions, tables, table, source
SelfTestCore
Check True, "Built-in self-tests"
`];
for(const [a,b,same] of [
 [' = SUM ( FactCommercial[gross_sales] ) ',"sum('FactCommercial'[gross_sales])",true],
 ['=IF(TRUE(); 1; 0)','=IF(TRUE(),1,0)',true],
 ['="a b"','="ab"',false],['="Sale"','="sale"',false],
 ['=[a b]','=[ab]',false],["='a b'[x]","='ab'[x]",false],
 ['="a"" b"','="a""b"',false],['=SUM(T[x])','=SUM(T[y])',false],
 ['=x>=0','=x>0',false],['=SUM(T[X])','=sum(t[x])',true],
 ['=SUM(T[x])+0','=SUM(T[x])+1',false]
]) lines.push(`Check (CanonicalFormula(${quote(a)}) = CanonicalFormula(${quote(b)})) = ${same?'True':'False'}, "Canonical comparison"`);
for(const line of ['broken','name :=','9bad := 1','a b := 1','a := =1'])lines.push(`ExpectLineFailure ${quote(line)}`);
for(const [p,want] of [['https://tenant.sharepoint.com/file.xlsx',true],[' HTTPS://d.docs.live.net/file.xlsx ',true],['http://host/file.xlsx',true],['C:\\Users\\user\\OneDrive\\project\\excel',false],['\\\\server\\share\\excel',false],['',false]])lines.push(`Check IsCloudPath(${quote(p)}) = ${want?'True':'False'}, "Cloud/local path classification"`);
lines.push('ExpectDefinitionsFailure "a := 1" & vbLf & "A := 2"');
lines.push('Check IsEmpty(ParseDefinitionLine("// comment")), "Comment ignored"');
lines.push('Check IsEmpty(ParseDefinitionLine("  ")), "Blank ignored"');
lines.push('parsed = ParseDefinitionLine(ChrW(-257) & "bom := 1")');
lines.push('Check parsed(0) = "bom", "UTF-8 BOM handled"');
for(const [name,formula] of specs){
 lines.push(`parsed = ParseDefinitionLine(${quote(name+' := '+formula.slice(1))})`);
 lines.push(`Check parsed(0) = ${quote(name)} And parsed(1) = ${quote(formula)}, "Source definition ${name}"`);
 lines.push(`Check NativeAddFormula(parsed(1), ",", ".") = ${quote(formula.slice(1))}, "Native Add expression ${name}"`);
 lines.push(`Check CanonicalFormula(NativeAddFormula(parsed(1), ";", ",")) = CanonicalFormula(parsed(1)), "Localized source expression ${name}"`);
}
lines.push('Check NativeAddFormula("NOT ISCROSSFILTERED(DimProduct[sku_id])", ",", ".") = "NOT ISCROSSFILTERED(DimProduct[sku_id])", "Expression without equals preserved"');
for(const [input,list,decimal,expected] of [
 ['=CALCULATE(SUM(T[x]), T[d] >= 1, T[d] <= 30)',';',',','CALCULATE(SUM(T[x]); T[d] >= 1; T[d] <= 30)'],
 ['=CALCULATE(SUM(T[x]), T[d] >= 31, T[d] <= 60)',',',',','CALCULATE(SUM(T[x]); T[d] >= 31; T[d] <= 60)'],
 ['=IF(TRUE(), 1.5, 0)',';',',','IF(TRUE(); 1,5; 0)'],
 ['=IF(TRUE(), "a,b.1", 0)',';',',','IF(TRUE(); "a,b.1"; 0)'],
 ["=SUM('a,b.1'[x,y.1])",';',',',"SUM('a,b.1'[x,y.1])"],
 ['=IF(TRUE(), "a""b,c.1", 0)',';',',','IF(TRUE(); "a""b,c.1"; 0)']
])lines.push(`Check NativeAddFormula(${quote(input)}, ${quote(list)}, ${quote(decimal)}) = ${quote(expected)}, "Locale numeric-boundary/literal regression"`);
lines.push('source = ""');
for(const line of dax.split(/\r?\n/))lines.push(`source = source & ${quote(line)} & vbLf`);
lines.push('Set definitions = ParseDefinitions(source)','Check definitions.Count = 66, "Complete source count"','Set tables = NewDictionary()');
for(const name of required){
 lines.push('Set table = New MockTable',`table.Name = ${quote(name)}`,
  `table.ModelTableColumns = Array(${contract[name].columns.map(c=>`MakeColumn(${quote(c.name)})`).join(',')})`,
  `tables.Add ${quote(name)}, table`);
}
lines.push('ValidateReferences definitions, tables','Check True, "Complete source dependencies and columns"',
 'ExpectReferenceFailure "bad := [future]" & vbLf & "future := 1", tables',
 'ExpectReferenceFailure "bad := SUM(FactCommercial[missing_column])", tables',
 'ExpectReferenceFailure "bad := SUM(MissingTable[value])", tables',
 'Set definitions = ParseDefinitions("literal := ""MissingTable[value]""")',
 'ValidateReferences definitions, tables','Check True, "Bracket inside literal ignored"');
for(const name of specs.map(s=>s[0]).filter(n=>n.endsWith('_rate')))lines.push(`Check FormatKind(${quote(name)}) = "percentage", "Percentage ${name}"`);
for(const name of ['sales_units','returned_units','orders','active_outlets','scheduled_visits','completed_visits','effective_calls','rfm_outlets','cohort_outlets','cohort_active_outlets'])lines.push(`Check FormatKind(${quote(name)}) = "whole", "Count ${name}"`);
for(const name of ['net_sales','gross_profit','average_order_value','dso_90d','calls_per_salesman_day'])lines.push(`Check FormatKind(${quote(name)}) = "decimal", "Two decimals ${name}"`);
for(const name of ['_snapshot_supported','_target_supported','_visit_supported'])lines.push(`Check FormatKind(${quote(name)}) = "boolean", "Boolean ${name}"`);
lines.push('WScript.Echo "PASS: " & tests & " pure-core assertions; no Excel workbook opened or changed."');
const temp=fs.mkdtempSync(path.join(os.tmpdir(),'measure-installer-test-'));
try{
 const file=path.join(temp,'core-tests.vbs');
 const script=lines.join('\r\n').replace(/^(\s*)Check (.+)$/gm,'$1Call Check($2)');
 fs.writeFileSync(file,script,'utf8');
 const result=spawnSync('C:\\Windows\\System32\\cscript.exe',['//nologo',file],{encoding:'utf8',timeout:30000,windowsHide:true});
 if(result.error)throw result.error;
 const lineMatch=(result.stdout+'\n'+result.stderr).match(/\((\d+),\s*\d+\)/);
 const rendered=script.split(/\r?\n/);
 const near=lineMatch?rendered.slice(Math.max(0,Number(lineMatch[1])-3),Number(lineMatch[1])+2).join('\n'):'';
 assert.equal(result.status,0,result.stdout+'\n'+result.stderr+'\n'+near);
 process.stdout.write(result.stdout);
 console.log('PASS: fixed target binding, conflict/backup ordering, home-table collisions and 66-source contract.');
 console.log('PENDING: native VBA compilation, ModelMeasures.Add, rerun and PivotTable/SQL checks in Excel.');
}finally{
 // Only remove the exact temporary script created here; no recursive deletion.
 fs.rmSync(path.join(temp,'core-tests.vbs'),{force:true});fs.rmdirSync(temp);
}
