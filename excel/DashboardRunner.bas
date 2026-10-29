Attribute VB_Name = "DashboardRunner"
Option Explicit

' User-run only. Import into DashboardRunner.xlsm, separate from the .xlsx target.
' No database, measure, relationship, security or VBA-project changes.
Private Const TARGET As String = "RevenueLeakage.xlsx"
Private Const OWNER As String = "RevenueLeakage dashboard runner v1"
Private wb As Workbook
Private fso As Object
Private base As String
Private runPath As String
Private logHandle As Integer
Private logging As Boolean
Private stage As String
Private managed As Collection
Private specs As Collection
Private oldStatus As Variant

Public Sub Phase4_Audit()
    On Error GoTo Failed
    If Not BeginRun("audit") Then Exit Sub
    Inventory
    PreflightDashboard
    LogEvent "AUDIT_READY", "", "Preflight passed; no workbook content changed."
    FinishRun "Audit passed. Save this log; next run Phase4_BuildDashboard."
    Exit Sub
Failed:
    FailRun
End Sub

Public Sub Phase4_BuildDashboard()
    Dim r As Variant, p As PivotTable, sh As Worksheet, reason As Variant, keys As String
    On Error GoTo Failed
    If Not BeginRun("dashboard") Then Exit Sub
    Inventory
    PreflightDashboard
    BackupTarget
    Set managed = New Collection
    For Each r In specs
        stage = "dashboard " & r(0)
        Application.StatusBar = stage
        Set p = FindPivot(r)
        If p Is Nothing Then
            Set sh = GetSheet(CStr(r(1)))
            Set p = NewPivot(sh, CStr(r(2)), CStr(r(0)))
            ConfigurePivot p, CStr(r(3)), CStr(r(4)), CStr(r(5))
            LogEvent "ADDED", p.Name, p.TableRange2.Address(External:=True)
        Else
            If ReusableEmptyPivot(p, r) Then
                ConfigurePivot p, CStr(r(3)), CStr(r(4)), CStr(r(5))
                LogEvent "EMPTY_PIVOT_ADOPTED", p.Name, "Existing empty model pivot configured; name/location preserved."
            End If
            LogEvent "REUSED", p.Name, "Matched native field layout; existing caption/location preserved."
        End If
        ClearManualFilters p
        SetMembers p, "DimDate.year", "2025"
        If CStr(r(0)) = "pLeakSKU" Or CStr(r(0)) = "pLeakOutlet" Then
            EnsureTopTen p, Split(CStr(r(3)), "|")(0)
        End If
        p.ManualUpdate = False
        If CStr(r(0)) = "pReturnReasons" Then
            keys = ""
            For Each reason In ReadCsv(base & "\dashboard-specs\return-reasons.csv")
                If Len(keys) > 0 Then keys = keys & "|"
                keys = keys & CStr(reason(0))
            Next reason
            SetMembers p, "FactCommercial.return_reason", keys
        End If
        managed.Add p
        If CStr(r(6)) = "line" Then EnsureTrendChart p
        If CStr(r(6)) = "bar" Then EnsureLeakageChart p
        If p.Name = "pCohort" Or CStr(r(0)) = "pCohort" Then
            p.RowGrand = False
            p.ColumnGrand = False
        End If
    Next r
    ' Existing control pivot stays in place and participates in shared slicers.
    Set p = NamedPivot("pControls")
    If p Is Nothing Then
        Set p = NamedPivot("pcontrol")
        If Not p Is Nothing Then LogEvent "CONTROL_ALIAS", p.Name, "Existing singular control name preserved."
    End If
    If Not p Is Nothing Then
        ClearManualFilters p
        SetMembers p, "DimDate.year", "2025"
        managed.Add p
    End If
    ConnectSlicers
    BuildMethodology
    FormatOwnedSheets
    BuildCards
    Inventory
    LogEvent "DASHBOARD_BUILT", "", "Save manually; inspect every sheet. Not Phase 4 completion."
    FinishRun "Dashboard stage finished. Inspect all sheets, press Ctrl+S, then run Phase4_Verify."
    Exit Sub
Failed:
    FailRun
End Sub

Public Sub Phase4_Verify()
    MsgBox "Do not restart the exhaustive suite. Its 282 main and 131 saved extended values already reconcile. Run Phase4_CheckCohortSamples for the remaining four small cases; refresh tests are separate.", vbInformation, "Use the remaining checks"
End Sub

Private Sub VerifyExhaustive()
    Dim cases As Collection, requests As Collection, row As Variant, req As Variant
    Dim p As PivotTable, sh As Worksheet, mainFile As Integer, extraFile As Integer
    Dim mainOpen As Boolean, extraOpen As Boolean, value As Variant, n As Long
    On Error GoTo Failed
    If Not BeginRun("verification") Then Exit Sub
    Set cases = ReadCsv(base & "\dashboard-specs\cases.csv")
    Set requests = ReadCsv(base & "\dashboard-specs\requests.csv")
    PreflightVerification
    BackupTarget
    Set sh = GetSheet("Native Verification")
    If sh.Range("Z1").Value2 = "" Then sh.Range("Z1").Value2 = OWNER
    Set p = NamedPivot("rlNativeProbe")
    If p Is Nothing Then
        Set p = NewPivot(sh, "A35", "rlNativeProbe")
        ' Include all capture measures once, without depending on visible SQL values.
        For Each req In requests
            EnsureMeasure p, CStr(req(1))
        Next req
        p.DataPivotField.Orientation = xlRowField
    Else
        For Each req In requests
            EnsureMeasure p, CStr(req(1))
        Next req
    End If
    AssertProbeIsolated p
    mainFile = FreeFile
    Open runPath & "\native-results.partial.csv" For Output As #mainFile
    mainOpen = True
    extraFile = FreeFile
    Open runPath & "\extended-results.partial.csv" For Output As #extraFile
    extraOpen = True
    Print #mainFile, "case_id,metric_id,actual_value"
    Print #extraFile, "case_id,metric_id,actual_value"
    For Each row In cases
        stage = "native case " & CStr(row(0))
        Application.StatusBar = stage & " (Esc/Ctrl+Break to stop)"
        p.ClearAllFilters
        ApplyCase p, row
        If Not p.RefreshTable Then Err.Raise vbObjectError + 701, , "Pivot refresh returned False."
        For Each req In requests
            If req(0) = row(0) Then
                value = p.GetPivotData("[Measures].[" & CStr(req(1)) & "]").Value2
                If CStr(req(2)) = "main" Then
                    Print #mainFile, Csv(CStr(row(0))) & "," & Csv(CStr(req(1))) & "," & InvariantValue(value)
                Else
                    Print #extraFile, Csv(CStr(row(0))) & "," & Csv(CStr(req(1))) & "," & InvariantValue(value)
                End If
                n = n + 1
            End If
        Next req
        LogEvent "CASE_CAPTURED", CStr(row(0)), "Native values read; comparison is performed by Node validator."
        DoEvents
    Next row
    Close #mainFile
    mainOpen = False
    Close #extraFile
    extraOpen = False
    Name runPath & "\native-results.partial.csv" As runPath & "\native-results.csv"
    Name runPath & "\extended-results.partial.csv" As runPath & "\extended-results.csv"
    LogEvent "CAPTURE_COMPLETE", "", CStr(n) & " actual native values. No SQL expectations read by VBA."
    FinishRun "Native captures written. Run the comparison commands in the guide. Capture completion alone is not a PASS."
    Exit Sub
Failed:
    If mainOpen Then Close #mainFile
    If extraOpen Then Close #extraFile
    FailRun
End Sub

Public Sub Phase4_CheckCohortSamples()
    Dim sh As Worksheet, p As PivotTable, field As PivotField, cf As CubeField
    Dim row As Variant, req As Variant, requests As Collection, value As Variant
    Dim handle As Integer, opened As Boolean, file As String, n As Long, i As Long
    On Error GoTo Failed
    If Not BeginRun("cohort-samples") Then Exit Sub
    Set sh = ExistingSheet("Cohort Verification")
    Set p = NamedPivot("rlCohortProbe")
    If Not sh Is Nothing Then
        If sh.Range("Z1").Value2 <> OWNER Then Err.Raise vbObjectError + 762, , "Cohort Verification is not runner-owned."
        If p Is Nothing Then AssertEmpty sh.Range("A5:L60")
    ElseIf Not p Is Nothing Then
        Err.Raise vbObjectError + 763, , "Cohort probe exists outside its dedicated sheet."
    End If
    If Not p Is Nothing Then
        If p.Parent.Name <> "Cohort Verification" Or Not p.PivotCache.OLAP Then Err.Raise vbObjectError + 764, , "Cohort probe has invalid source/location."
        AssertProbeIsolated p
        For Each field In p.DataFields
            If InStr("|[Measures].[cohort_outlets]|[Measures].[cohort_active_outlets]|[Measures].[retention_rate]|", "|" & field.CubeField.Name & "|") = 0 Then Err.Raise vbObjectError + 765, , "Unexpected measure on cohort probe."
        Next field
        For Each cf In p.CubeFields
            If cf.CubeFieldType = xlHierarchy And cf.Orientation <> xlHidden Then
                If cf.Name <> Hierarchy("FactCohort.cohort_month") And cf.Name <> Hierarchy("FactCohort.activity_month") Then Err.Raise vbObjectError + 766, , "Unexpected cohort probe dimension."
            End If
        Next cf
    End If
    Set requests = ReadCsv(base & "\dashboard-specs\cohort-sample-requests.csv")
    BackupTarget
    Set sh = GetSheet("Cohort Verification")
    sh.Range("Z1").Value2 = OWNER
    sh.Range("A1").Value2 = "Remaining cohort checks | Four cases only | Isolated from dashboard filters"
    If p Is Nothing Then Set p = NewPivot(sh, "A10", "rlCohortProbe")
    EnsureMeasure p, "cohort_outlets"
    EnsureMeasure p, "cohort_active_outlets"
    EnsureMeasure p, "retention_rate"
    p.DataPivotField.Orientation = xlRowField
    AssertProbeIsolated p
    file = runPath & "\cohort-samples.partial.csv"
    handle = FreeFile
    Open file For Output As #handle
    opened = True
    Print #handle, "case_id,metric_id,actual_value"
    Close #handle
    opened = False
    For Each row In ReadCsv(base & "\dashboard-specs\cohort-samples.csv")
        i = i + 1
        stage = "cohort sample " & CStr(i) & "/4 " & CStr(row(0))
        Application.StatusBar = stage
        p.ClearAllFilters
        ' No operating-date filter: cohort dates are independent of DimDate.
        SetMembers p, "FactCohort.cohort_month", CStr(row(8))
        SetMembers p, "FactCohort.activity_month", CStr(row(9))
        If Not p.RefreshTable Then Err.Raise vbObjectError + 767, , "Cohort sample refresh returned False."
        handle = FreeFile
        Open file For Append As #handle
        opened = True
        For Each req In requests
            If req(0) = row(0) Then
                value = p.GetPivotData("[Measures].[" & CStr(req(1)) & "]").Value2
                Print #handle, Csv(CStr(row(0))) & "," & Csv(CStr(req(1))) & "," & InvariantValue(value)
                n = n + 1
            End If
        Next req
        Close #handle
        opened = False
        LogEvent "CASE_CAPTURED", CStr(row(0)), "Dedicated three-measure probe; file flushed after this case."
        DoEvents
    Next row
    If i <> 4 Or n <> 10 Then Err.Raise vbObjectError + 768, , "Unexpected cohort sample coverage."
    Name file As runPath & "\cohort-samples.csv"
    FinishRun "Four cohort cases captured (10 values). Compare with the --cohort-samples validator; do not rerun the exhaustive suite."
    Exit Sub
Failed:
    If opened Then Close #handle
    FailRun
End Sub

Public Sub Phase4_FinishDashboard()
    Dim row As Variant, p As PivotTable
    On Error GoTo Failed
    If Not BeginRun("finish-dashboard") Then Exit Sub
    Inventory
    PreflightDashboard
    Set managed = New Collection
    For Each row In specs
        Set p = FindPivot(row)
        If p Is Nothing Then Err.Raise vbObjectError + 757, , "Missing dashboard pivot: " & row(0) & ". Use BuildDashboard first."
        If ReusableEmptyPivot(p, row) Then Err.Raise vbObjectError + 758, , "Dashboard pivot still empty: " & p.Name
        managed.Add p
    Next row
    Set p = NamedPivot("pControls")
    If p Is Nothing Then Set p = NamedPivot("pcontrol")
    If Not p Is Nothing Then managed.Add p
    BackupTarget
    ConnectSlicers
    BuildMethodology
    FormatOwnedSheets
    BuildCards
    Inventory
    FinishRun "Dashboard finishing stage completed. Inspect all sheets and save; native numeric verification remains required."
    Exit Sub
Failed:
    FailRun
End Sub

Public Sub Phase4_TestRefresh()
    Dim original As String, originalPath As String, alt As String, missing As String
    Dim row As Variant, file As Object, changed As Boolean, expectedFailure As Boolean
    Dim failure As String, failureNumber As Long, restored As Boolean
    On Error GoTo Failed
    If Not BeginRun("refresh") Then Exit Sub
    original = wb.Queries("pFolder").Formula
    originalPath = FolderLiteral(original)
    If Not fso.FolderExists(originalPath) Then Err.Raise vbObjectError + 702, , "Original pFolder does not exist."
    For Each row In ReadCsv(base & "\dashboard-specs\tables.csv")
        If Not fso.FileExists(originalPath & "\" & row(0) & ".csv") Then Err.Raise vbObjectError + 703, , "Missing source CSV: " & row(0)
    Next row
    PreflightVerification
    BackupTarget
    WriteText runPath & "\original-pFolder.m", original
    stage = "refresh original folder"
    RefreshAndCheck "ORIGINAL"
    alt = runPath & "\alternate-source"
    fso.CreateFolder alt
    For Each file In fso.GetFolder(originalPath).Files
        If LCase$(fso.GetExtensionName(file.Name)) = "csv" Then fso.CopyFile file.Path, alt & "\" & file.Name, False
    Next file
    missing = runPath & "\intentionally-missing-source"
    If fso.FolderExists(missing) Then Err.Raise vbObjectError + 704, , "Missing-folder test directory unexpectedly exists."
    changed = True
    wb.Queries("pFolder").Formula = MFolder(alt)
    stage = "refresh alternate folder"
    RefreshAndCheck "ALTERNATE"
    wb.Queries("pFolder").Formula = MFolder(missing)
    stage = "expected missing-folder failure"
    On Error Resume Next
    wb.Model.Refresh
    failureNumber = Err.Number
    failure = Err.Description
    Err.Clear
    On Error GoTo Failed
    expectedFailure = (failureNumber <> 0)
    LogEvent "MISSING_FOLDER_OBSERVED", "", "Error=" & CStr(failureNumber) & "; " & failure
    wb.Queries("pFolder").Formula = original
    stage = "restore original folder and refresh"
    RefreshAndCheck "RESTORED"
    changed = False
    If Not expectedFailure Then
        Err.Raise vbObjectError + 705, , "Excel did not surface a synchronous missing-folder error. Recovery completed, but this test is INCONCLUSIVE; inspect Queries & Connections manually."
    End If
    LogEvent "REFRESH_SEQUENCE_COMPLETED", "", "Original parameter restored. Query UI state and post-refresh numeric verification still required."
    FinishRun "Refresh sequence finished and original pFolder restored. Inspect query errors, save and rerun Phase4_Verify."
    Exit Sub
Failed:
    failureNumber = Err.Number
    failure = Err.Description
    If changed Then
        On Error Resume Next
        Err.Clear
        wb.Queries("pFolder").Formula = original
        wb.Model.Refresh
        restored = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
        LogEvent "RECOVERY", "", "Original parameter restoration/refresh succeeded=" & CStr(restored)
        If Not restored Then failure = failure & " RECOVERY FAILED: restore pFolder from original-pFolder.m and refresh before saving."
    End If
    ReportFailure failureNumber, failure
End Sub

Private Function BeginRun(ByVal action As String) As Boolean
    Dim picker As Object, candidate As Workbook, name As Variant, t As Object, m As Object
    Dim found As Boolean, stamp As String, source As String, line As Variant, metric As String, expected As String
    logging = False
    oldStatus = Application.StatusBar
    stage = "validate runner"
    If StrComp(ThisWorkbook.Name, "DashboardRunner.xlsm", vbTextCompare) <> 0 Then Err.Raise vbObjectError + 710, , "Import into a separate DashboardRunner.xlsm workbook."
    Set fso = CreateObject("Scripting.FileSystemObject")
    base = ThisWorkbook.Path
    If InStr(1, base, "://", vbTextCompare) > 0 Or Len(base) = 0 Then base = ""
    If Len(base) > 0 Then
        If Not fso.FileExists(base & "\dashboard-specs\pivots.csv") Then base = ""
    End If
    If Len(base) = 0 Then
        Set picker = Application.FileDialog(4)
        picker.Title = "Select local project excel folder containing RevenueLeakage.xlsx and dashboard-specs"
        If picker.Show <> -1 Then Exit Function
        base = CStr(picker.SelectedItems(1))
    End If
    If Not fso.FileExists(base & "\" & TARGET) Or Not fso.FileExists(base & "\dashboard-specs\pivots.csv") Then Err.Raise vbObjectError + 711, , "Wrong local project excel folder."
    Set wb = Nothing
    For Each candidate In Application.Workbooks
        If StrComp(candidate.Name, TARGET, vbTextCompare) = 0 Then Set wb = candidate
    Next candidate
    If wb Is Nothing Then Err.Raise vbObjectError + 712, , "Open RevenueLeakage.xlsx first."
    If InStr(1, wb.FullName, "://", vbTextCompare) > 0 Then
        If MsgBox("Confirm the open cloud workbook is the Phase 4 target:" & vbCrLf & wb.FullName & vbCrLf & "Local project folder: " & base, vbYesNo + vbQuestion + vbDefaultButton2) <> vbYes Then Exit Function
    ElseIf StrComp(fso.GetAbsolutePathName(wb.FullName), fso.GetAbsolutePathName(base & "\" & TARGET), vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 713, , "Target path does not match the selected folder."
    End If
    If wb Is ThisWorkbook Then Err.Raise vbObjectError + 714, , "Runner cannot be the target."
    If wb.ReadOnly Or wb.FileFormat <> xlOpenXMLWorkbook Or Not wb.Saved Then Err.Raise vbObjectError + 715, , "Target must be writable, saved with Ctrl+S, and macro-free .xlsx."
    If Not fso.FolderExists(base & "\dashboard-runs") Then fso.CreateFolder base & "\dashboard-runs"
    stamp = Format$(Now, "yyyymmdd-hhnnss") & "-" & action
    runPath = base & "\dashboard-runs\" & stamp
    If fso.FolderExists(runPath) Then Err.Raise vbObjectError + 716, , "Run directory already exists; retry after one second."
    fso.CreateFolder runPath
    logHandle = FreeFile
    Open runPath & "\events.csv" For Output As #logHandle
    logging = True
    Print #logHandle, "timestamp,event,object,detail"
    LogEvent "START", action, "Excel " & Application.Version & "; OS=" & Application.OperatingSystem & "; target=" & wb.FullName
    If ReadText(base & "\dashboard-specs\manifest.json") <> ReadText(base & "\..\data\excel\manifest.json") Then Err.Raise vbObjectError + 744, , "Source manifest changed. Regenerate and review dashboard specs first."
    ValidateRelationships
    If wb.Model.ModelTables.Count <> 14 Then Err.Raise vbObjectError + 752, , "Expected exactly 14 model tables; SQLReference must remain worksheet-only."
    For Each name In ReadCsv(base & "\dashboard-specs\tables.csv")
        Set t = wb.Model.ModelTables(CStr(name(0)))
        If t.RecordCount <> CLng(name(1)) Then Err.Raise vbObjectError + 717, , "Model row count mismatch: " & name(0)
    Next name
    source = ReadText(base & "\measures.dax")
    For Each line In Split(Replace(source, vbCr, ""), vbLf)
        If InStr(CStr(line), " := ") > 0 Then
            metric = Trim$(Split(CStr(line), " := ")(0))
            expected = Trim$(Mid$(CStr(line), InStr(CStr(line), " := ") + 4))
            found = False
            For Each m In wb.Model.ModelMeasures
                If StrComp(m.Name, metric, vbTextCompare) = 0 Then
                    found = True
                    If CanonicalFormula(m.Formula) <> CanonicalFormula(expected) Then Err.Raise vbObjectError + 745, , "Measure formula conflict: " & metric
                End If
            Next m
            If Not found Then Err.Raise vbObjectError + 718, , "Missing measure: " & metric & ". Run the measure installer first."
        End If
    Next line
    Set specs = ReadCsv(base & "\dashboard-specs\pivots.csv")
    LogEvent "PREFLIGHT", "", "Expected model counts, relationships and measure formulas match; model definitions not rewritten."
    BeginRun = True
End Function

Private Sub ValidateRelationships()
    Dim row As Variant, rel As Object, found As Boolean, expected As Collection
    Dim reversed As Boolean, inactive As Boolean, detail As String
    stage = "validate model relationships"
    Set expected = ReadCsv(base & "\relationships.csv")
    For Each rel In wb.Model.ModelRelationships
        LogEvent "RELATIONSHIP_ACTUAL", rel.ForeignKeyTable.Name & "." & rel.ForeignKeyColumn.Name, _
            "Primary=" & rel.PrimaryKeyTable.Name & "." & rel.PrimaryKeyColumn.Name & "; active=" & CStr(rel.Active)
    Next rel
    If wb.Model.ModelRelationships.Count <> expected.Count Then Err.Raise vbObjectError + 746, , "Unexpected model relationship count."
    For Each row In expected
        found = False
        reversed = False
        inactive = False
        For Each rel In wb.Model.ModelRelationships
            If StrComp(rel.PrimaryKeyTable.Name, CStr(row(0)), vbTextCompare) = 0 And StrComp(rel.PrimaryKeyColumn.Name, CStr(row(1)), vbTextCompare) = 0 And _
                StrComp(rel.ForeignKeyTable.Name, CStr(row(2)), vbTextCompare) = 0 And StrComp(rel.ForeignKeyColumn.Name, CStr(row(3)), vbTextCompare) = 0 Then
                If rel.Active Then found = True Else inactive = True
            End If
            If StrComp(rel.PrimaryKeyTable.Name, CStr(row(2)), vbTextCompare) = 0 And StrComp(rel.PrimaryKeyColumn.Name, CStr(row(3)), vbTextCompare) = 0 And _
                StrComp(rel.ForeignKeyTable.Name, CStr(row(0)), vbTextCompare) = 0 And StrComp(rel.ForeignKeyColumn.Name, CStr(row(1)), vbTextCompare) = 0 Then reversed = True
        Next rel
        If Not found Then
            detail = "Relationship not matched: " & row(0) & "." & row(1) & " -> " & row(2) & "." & row(3)
            If inactive Then detail = detail & "; matching endpoints are INACTIVE"
            If reversed Then detail = detail & "; endpoints exist in REVERSE primary/foreign orientation"
            LogEvent "RELATIONSHIP_CONFLICT", CStr(row(2)), detail
            Err.Raise vbObjectError + 747, , detail & ". See RELATIONSHIP_ACTUAL entries in events.csv. No relationship was changed."
        End If
    Next row
End Sub

Private Sub BackupTarget()
    stage = "backup"
    wb.SaveCopyAs runPath & "\RevenueLeakage-before-" & Format$(Now, "hhnnss") & ".xlsx"
    LogEvent "BACKUP_OK", "", runPath
End Sub

Private Sub Inventory()
    Dim sh As Worksheet, p As PivotTable, cf As CubeField, chart As ChartObject
    Dim sc As SlicerCache, sp As Object, m As Object, rel As Object, connection As WorkbookConnection
    Dim filters As Variant, entry As Variant
    For Each sh In wb.Worksheets
        LogEvent "SHEET", sh.Name, "UsedRange=" & sh.UsedRange.Address
        For Each p In sh.PivotTables
            LogEvent "PIVOT", p.Name, p.TableRange2.Address(External:=True) & "; OLAP=" & p.PivotCache.OLAP & "; cache=" & CStr(p.CacheIndex)
            If p.PivotCache.OLAP Then
                For Each cf In p.CubeFields
                    If cf.Orientation <> xlHidden Then
                        LogEvent "FIELD", p.Name, cf.Name & "; orientation=" & CStr(cf.Orientation)
                        If cf.CubeFieldType = xlHierarchy Then
                            filters = cf.PivotFields(1).VisibleItemsList
                            If IsArray(filters) Then
                                For Each entry In filters
                                    LogEvent "FILTER_MEMBER", p.Name & ":" & cf.Name, CStr(entry)
                                Next entry
                            End If
                        End If
                    End If
                Next cf
            End If
        Next p
        For Each chart In sh.ChartObjects
            LogEvent "CHART", sh.Name & ":" & chart.Name, "Position=" & chart.TopLeftCell.Address
        Next chart
    Next sh
    For Each sc In wb.SlicerCaches
        LogEvent "SLICER_CACHE", sc.Name, sc.SourceName
        For Each sp In sc.PivotTables
            LogEvent "SLICER_CONNECTION", sc.Name, sp.Name
        Next sp
    Next sc
    For Each m In wb.Model.ModelMeasures
        LogEvent "MEASURE", m.Name, m.Formula
    Next m
    For Each rel In wb.Model.ModelRelationships
        LogEvent "RELATIONSHIP", rel.ForeignKeyTable.Name & "." & rel.ForeignKeyColumn.Name, rel.PrimaryKeyTable.Name & "." & rel.PrimaryKeyColumn.Name & "; active=" & rel.Active
    Next rel
    For Each connection In wb.Connections
        LogEvent "CONNECTION", connection.Name, "Type=" & CStr(connection.Type)
    Next connection
    stage = "write inventory manifest snapshot"
    WriteText runPath & "\manifest.json", ReadText(base & "\..\data\excel\manifest.json")
End Sub

Private Sub PreflightDashboard()
    Dim row As Variant, p As PivotTable, other As Variant, area As Range, sh As Worksheet
    Dim slices As Collection, item As Variant, sc As SlicerCache
    stage = "preflight dashboard destinations"
    For Each row In specs
        Set p = FindPivot(row)
        If p Is Nothing Then
            Set sh = ExistingSheet(CStr(row(1)))
            If Not sh Is Nothing Then
                Set area = Slot(sh, row)
                AssertEmpty area
            End If
        End If
    Next row
    ' Cards and methodology have explicit ownership markers; unknown contents stop the build.
    CheckOwnedArea "Executive Summary", "A1:L9"
    CheckOwnedArea "Methodology", "A1:L24"
    CheckOwnedArea "Leakage Diagnostics", "A1:L4"
    CheckOwnedArea "Operations", "A1:L4"
    Set slices = ReadCsv(base & "\dashboard-specs\slicers.csv")
    For Each item In slices
        Set sc = FindSlicerCache(CStr(item(1)))
    Next item
    PreflightVerification
End Sub

Private Function FindPivot(ByVal spec As Variant) As PivotTable
    Dim sh As Worksheet, p As PivotTable, candidate As PivotTable, n As Long
    Dim named As PivotTable
    Set named = NamedPivot(CStr(spec(0)))
    For Each sh In wb.Worksheets
        For Each p In sh.PivotTables
            If StrComp(sh.Name, CStr(spec(1)), vbTextCompare) = 0 Then
                If Matches(p, spec) Or ReusableEmptyPivot(p, spec) Then
                    Set candidate = p
                    n = n + 1
                End If
            End If
        Next p
    Next sh
    If Not named Is Nothing Then
        If (Not Matches(named, spec) And Not ReusableEmptyPivot(named, spec)) Or StrComp(named.Parent.Name, CStr(spec(1)), vbTextCompare) <> 0 Then Err.Raise vbObjectError + 720, , "Named pivot has conflicting layout: " & named.Name
    End If
    If n > 1 Then Err.Raise vbObjectError + 721, , "Ambiguous pivot candidates: " & spec(0)
    If n = 1 Then Set FindPivot = candidate
End Function

Private Function ReusableEmptyPivot(ByVal p As PivotTable, ByVal spec As Variant) As Boolean
    Dim cf As CubeField, anchor As Range
    If Not p.PivotCache.OLAP Then Exit Function
    If StrComp(p.Parent.Name, CStr(spec(1)), vbTextCompare) <> 0 Then Exit Function
    Set anchor = p.Parent.Range(CStr(spec(2)))
    If p.TableRange2.Row <> anchor.Row Or p.TableRange2.Column <> anchor.Column Then Exit Function
    If p.PivotCache.WorkbookConnection.Name <> wb.Model.DataModelConnection.Name Then Exit Function
    If p.DataFields.Count <> 0 Or p.RowFields.Count <> 0 Or p.ColumnFields.Count <> 0 Or p.PageFields.Count <> 0 Then Exit Function
    For Each cf In p.CubeFields
        If cf.Orientation <> xlHidden Then Exit Function
    Next cf
    ReusableEmptyPivot = True
End Function

Private Function Matches(ByVal p As PivotTable, ByVal spec As Variant) As Boolean
    Dim cf As CubeField, rows As String, cols As String, values As String
    If Not p.PivotCache.OLAP Then Exit Function
    For Each cf In p.CubeFields
        Select Case cf.Orientation
            Case xlRowField: rows = rows & "|" & cf.Name
            Case xlColumnField: cols = cols & "|" & cf.Name
            Case xlDataField: values = values & "|" & cf.Name
        End Select
    Next cf
    Matches = SameSet(rows, FieldNames(CStr(spec(3)))) And SameSet(cols, FieldNames(CStr(spec(4)))) And SameSet(values, MeasureNames(CStr(spec(5))))
End Function

Private Function SameSet(ByVal a As String, ByVal b As String) As Boolean
    Dim entry As Variant, countA As Long, countB As Long
    For Each entry In Split(a, "|")
        If Len(entry) > 0 Then
            countA = countA + 1
            If InStr(1, "|" & b & "|", "|" & entry & "|", vbTextCompare) = 0 Then Exit Function
        End If
    Next entry
    For Each entry In Split(b, "|")
        If Len(entry) > 0 Then countB = countB + 1
    Next entry
    SameSet = (countA = countB)
End Function

Private Function FieldNames(ByVal fields As String) As String
    Dim entry As Variant
    For Each entry In Split(fields, "|")
        If Len(entry) > 0 Then FieldNames = FieldNames & "|" & Hierarchy(CStr(entry))
    Next entry
End Function

Private Function MeasureNames(ByVal fields As String) As String
    Dim entry As Variant
    For Each entry In Split(fields, "|")
        If Len(entry) > 0 Then MeasureNames = MeasureNames & "|[Measures].[" & entry & "]"
    Next entry
End Function

Private Function NamedPivot(ByVal name As String) As PivotTable
    Dim sh As Worksheet, p As PivotTable
    For Each sh In wb.Worksheets
        For Each p In sh.PivotTables
            If StrComp(p.Name, name, vbTextCompare) = 0 Then Set NamedPivot = p
        Next p
    Next sh
End Function

Private Function ExistingSheet(ByVal name As String) As Worksheet
    Dim sh As Worksheet
    For Each sh In wb.Worksheets
        If StrComp(sh.Name, name, vbTextCompare) = 0 Then Set ExistingSheet = sh
    Next sh
End Function

Private Function GetSheet(ByVal name As String) As Worksheet
    Dim sh As Worksheet
    Set sh = ExistingSheet(name)
    If sh Is Nothing Then
        Set sh = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        sh.Name = name
    End If
    Set GetSheet = sh
End Function

Private Function Slot(ByVal sh As Worksheet, ByVal spec As Variant) As Range
    Dim first As Range, height As Long
    Set first = sh.Range(CStr(spec(2)))
    Select Case CStr(spec(0))
        Case "pLeakSKU", "pLeakOutlet": height = 22
        Case "pReturnReasons": height = 25
        Case "pVisits": height = 58
        Case "pCohort": height = 60
        Case "pAcceptanceControls": height = 80
        Case Else: height = 28
    End Select
    Set Slot = first.Resize(height, 12)
End Function

Private Sub AssertEmpty(ByVal area As Range)
    Dim p As PivotTable, shape As Shape, cell As Range
    If Application.WorksheetFunction.CountA(area) > 0 Then Err.Raise vbObjectError + 722, , "Destination contains data: " & area.Address(External:=True)
    For Each cell In area.Cells
        If cell.MergeCells Then Err.Raise vbObjectError + 753, , "Destination contains merged cells: " & area.Address(External:=True)
    Next cell
    For Each p In area.Parent.PivotTables
        If Not Application.Intersect(area, p.TableRange2) Is Nothing Then Err.Raise vbObjectError + 723, , "Destination overlaps pivot: " & p.Name
    Next p
    For Each shape In area.Parent.Shapes
        If Not Application.Intersect(area, area.Parent.Range(shape.TopLeftCell, shape.BottomRightCell)) Is Nothing Then Err.Raise vbObjectError + 724, , "Destination overlaps shape/chart: " & shape.Name
    Next shape
End Sub

Private Sub EnsureLeakageChart(ByVal p As PivotTable)
    Dim sh As Worksheet, chart As ChartObject, linked As PivotTable, found As ChartObject
    Set sh = GetSheet("Leakage Diagnostics")
    For Each chart In sh.ChartObjects
        Set linked = Nothing
        On Error Resume Next
        Set linked = chart.Chart.PivotLayout.PivotTable
        On Error GoTo 0
        If Not linked Is Nothing Then
            If linked.Name = p.Name Then
                If Not found Is Nothing Then Err.Raise vbObjectError + 748, , "Ambiguous leakage charts."
                Set found = chart
            End If
        End If
    Next chart
    If Not found Is Nothing Then Exit Sub
    AssertEmpty sh.Range("A35:L60")
    Set chart = sh.ChartObjects.Add(sh.Range("A35").Left, sh.Range("A35").Top, 650, 340)
    chart.Name = "rlLeakageChart"
    chart.Chart.SetSourceData p.TableRange1
    chart.Chart.ChartType = xlBarClustered
    chart.Chart.HasTitle = True
    chart.Chart.ChartTitle.Text = "Observed Price Shortfall by Branch (IDR)"
End Sub

Private Sub CheckOwnedArea(ByVal sheetName As String, ByVal address As String)
    Dim sh As Worksheet
    Set sh = ExistingSheet(sheetName)
    If sh Is Nothing Then Exit Sub
    If sh.Range("Z1").Value2 = OWNER Then Exit Sub
    If Len(CStr(sh.Range("Z1").Value2)) > 0 Then Err.Raise vbObjectError + 725, , "Ownership marker occupied: " & sheetName & "!Z1"
    AssertEmpty sh.Range(address)
End Sub

Private Function NewPivot(ByVal sh As Worksheet, ByVal anchor As String, ByVal name As String) As PivotTable
    Dim cache As PivotCache, conn As WorkbookConnection
    Set conn = wb.Model.DataModelConnection
    Set cache = wb.PivotCaches.Create(SourceType:=xlExternal, SourceData:=conn, Version:=xlPivotTableVersion15)
    Set NewPivot = cache.CreatePivotTable(TableDestination:=sh.Range(anchor), TableName:=name, DefaultVersion:=xlPivotTableVersion15)
End Function

Private Sub ConfigurePivot(ByVal p As PivotTable, ByVal rows As String, ByVal cols As String, ByVal measures As String)
    Dim entry As Variant, position As Long
    p.ManualUpdate = True
    For Each entry In Split(rows, "|")
        If Len(entry) > 0 Then
            position = position + 1
            With p.CubeFields(Hierarchy(CStr(entry)))
                .Orientation = xlRowField
                .Position = position
            End With
        End If
    Next entry
    position = 0
    For Each entry In Split(cols, "|")
        If Len(entry) > 0 Then
            position = position + 1
            With p.CubeFields(Hierarchy(CStr(entry)))
                .Orientation = xlColumnField
                .Position = position
            End With
        End If
    Next entry
    For Each entry In Split(measures, "|")
        EnsureMeasure p, CStr(entry)
    Next entry
    If rows = "" And cols = "" Then p.DataPivotField.Orientation = xlRowField
    p.RowAxisLayout xlTabularRow
    p.TableStyle2 = "PivotStyleMedium2"
End Sub

Private Sub ClearManualFilters(ByVal p As PivotTable)
    Dim cf As CubeField
    For Each cf In p.CubeFields
        If cf.CubeFieldType = xlHierarchy And cf.Orientation <> xlHidden Then cf.ClearManualFilter
    Next cf
End Sub

Private Sub EnsureTopTen(ByVal p As PivotTable, ByVal field As String)
    Dim filters As PivotFilters, filter As PivotFilter
    Set filters = p.CubeFields(Hierarchy(field)).PivotFields(1).PivotFilters
    If filters.Count > 0 Then
        If filters.Count <> 1 Then Err.Raise vbObjectError + 749, , "Conflicting diagnostic filters: " & p.Name
        Set filter = filters(1)
        If filter.FilterType <> xlTopCount Or filter.Value1 <> 10 Then Err.Raise vbObjectError + 750, , "Expected top-ten shortfall filter: " & p.Name
        If filter.DataCubeField.Name <> "[Measures].[price_shortfall]" Then Err.Raise vbObjectError + 751, , "Top-ten filter ranks a different metric: " & p.Name
        Exit Sub
    End If
    filters.Add Type:=xlTopCount, DataField:=p.CubeFields("[Measures].[price_shortfall]"), Value1:=10
End Sub

Private Sub EnsureMeasure(ByVal p As PivotTable, ByVal name As String)
    Dim field As PivotField
    For Each field In p.DataFields
        If StrComp(field.CubeField.Name, "[Measures].[" & name & "]", vbTextCompare) = 0 Then Exit Sub
    Next field
    Set field = p.AddDataField(p.CubeFields("[Measures].[" & name & "]"), name)
    If InStr(name, "rate") > 0 Then
        field.NumberFormat = "0.00%"
    ElseIf name = "orders" Or name = "active_outlets" Or InStr(name, "visits") > 0 Or name = "effective_calls" Or name = "rfm_outlets" Or InStr(name, "cohort_") = 1 Then
        field.NumberFormat = "#,##0"
    Else
        field.NumberFormat = "#,##0.00"
    End If
End Sub

Private Function Hierarchy(ByVal field As String) As String
    Dim parts As Variant
    parts = Split(field, ".")
    If UBound(parts) <> 1 Then Err.Raise vbObjectError + 726, , "Invalid field identifier: " & field
    Hierarchy = "[" & parts(0) & "].[" & parts(1) & "]"
End Function

Private Function MemberName(ByVal field As String, ByVal key As String) As String
    MemberName = Hierarchy(field) & ".&[" & Replace(key, "]", "]]") & "]"
End Function

Private Sub SetMembers(ByVal p As PivotTable, ByVal field As String, ByVal keys As String)
    Dim cf As CubeField, items As Variant, i As Long, member As String
    If Len(keys) = 0 Then Exit Sub
    Set cf = p.CubeFields(Hierarchy(field))
    If cf.Orientation = xlHidden Then cf.Orientation = xlPageField
    If cf.Orientation = xlPageField Then cf.EnableMultiplePageItems = True
    items = Split(keys, "|")
    For i = LBound(items) To UBound(items)
        member = CStr(items(i))
        If field = "DimDate.date" Or field = "FactCohort.cohort_month" Or field = "FactCohort.activity_month" Then member = member & "T00:00:00"
        items(i) = MemberName(field, member)
    Next i
    cf.PivotFields(1).VisibleItemsList = items
End Sub

Private Function FindSlicerCache(ByVal field As String) As SlicerCache
    Dim sc As SlicerCache, found As SlicerCache, n As Long
    For Each sc In wb.SlicerCaches
        If StrComp(sc.SourceName, Hierarchy(field), vbTextCompare) = 0 Then
            n = n + 1
            Set found = sc
        End If
    Next sc
    If n > 1 Then Err.Raise vbObjectError + 727, , "Multiple slicer caches for field " & field & "; resolve in audit before building."
    If n = 1 Then Set FindSlicerCache = found
End Function

Private Sub ConnectSlicers()
    Dim row As Variant, sc As SlicerCache, p As PivotTable, connected As Object
    Dim found As Boolean, sh As Worksheet, first As PivotTable, levelName As String, view As Slicer
    Dim placement As Range
    Set first = managed(1)
    Set sh = GetSheet("Executive Summary")
    For Each row In ReadCsv(base & "\dashboard-specs\slicers.csv")
        stage = "find slicer cache " & CStr(row(1))
        Set sc = FindSlicerCache(CStr(row(1)))
        If sc Is Nothing Then
            stage = "reserve slicer space " & CStr(row(1))
            Set placement = SlicerPlacement(sh, CLng(row(2)))
            AssertSlicerSpace placement
            stage = "create slicer cache " & CStr(row(1))
            LogEvent "SLICER_CACHE_REQUEST", CStr(row(1)), "Connection=" & wb.Model.DataModelConnection.Name & "; hierarchy=" & Hierarchy(CStr(row(1)))
            ' Bind directly to this workbook's model; let Excel allocate a unique cache name.
            Set sc = wb.SlicerCaches.Add(wb.Model.DataModelConnection, Hierarchy(CStr(row(1))))
            LogEvent "SLICER_CACHE_CREATED", sc.Name, sc.SourceName
        Else
            LogEvent "SLICER_REUSED", sc.Name, CStr(row(1))
        End If
        ' Associate the model cache with actual PivotTables before creating its view.
        For Each p In managed
            stage = "connect slicer " & CStr(row(1)) & " to " & p.Name
            found = False
            For Each connected In sc.PivotTables
                If connected.Name = p.Name Then found = True
            Next connected
            If Not found Then sc.PivotTables.AddPivotTable p
            LogEvent "SLICER_CONNECTED", sc.Name, p.Name
        Next p
        If sc.Slicers.Count = 0 Then
            stage = "resolve slicer level " & CStr(row(1))
            Set placement = SlicerPlacement(sh, CLng(row(2)))
            AssertSlicerSpace placement
            If sc.SlicerCacheLevels.Count <> 1 Then Err.Raise vbObjectError + 756, , "Expected a single attribute level: " & sc.Name
            levelName = sc.SlicerCacheLevels(1).Name
            LogEvent "SLICER_VIEW_REQUEST", sc.Name, "Level=" & levelName & "; destination=" & sh.Name & "; protected=" & CStr(sh.ProtectDrawingObjects) & "; connections=" & CStr(sc.PivotTables.Count)
            If sh.ProtectDrawingObjects Then Err.Raise vbObjectError + 759, , "Destination drawing objects are protected."
            stage = "create minimal slicer view " & CStr(row(1))
            ' Only the two documented OLAP creation arguments. Layout is a separate operation.
            Set view = sc.Slicers.Add(sh, levelName)
            LogEvent "SLICER_VIEW_CREATED", sc.Name, view.Name
            stage = "position slicer view " & CStr(row(1))
            view.Caption = CStr(Split(CStr(row(1)), ".")(1))
            view.Top = placement.Top
            view.Left = placement.Left
            view.Width = 220
            view.Height = 150
            view.Shape.Placement = xlFreeFloating
            LogEvent "SLICER_POSITIONED", sc.Name, view.Name & "; top=" & CStr(view.Top) & "; bottom=" & CStr(view.Top + view.Height)
        End If
        For Each view In sc.Slicers
            If view.Shape.Parent.Name = sh.Name Then
                If Abs(view.Left - sh.Range("N1").Left) < 1 Then view.Shape.Placement = xlFreeFloating
            End If
        Next view
        stage = "default slicer selection " & CStr(row(1))
        sc.ClearManualFilter
        If CStr(row(1)) = "DimDate.year" Then sc.VisibleSlicerItemsList = Array(MemberName("DimDate.year", "2025"))
    Next row
End Sub

Private Function SlicerPlacement(ByVal sh As Worksheet, ByVal minimumRow As Long) As Range
    Dim shape As Shape
    Dim top As Double, left As Double, right As Double, firstRow As Long, lastRow As Long, lastColumn As Long
    left = sh.Range("N1").Left
    right = sh.Range("U1").Left + sh.Range("U1").Width
    If right < left + 220 Then right = left + 220
    top = sh.Cells(minimumRow, 14).Top
    ' Read geometry from worksheet shapes directly; no cache/view parent assumptions.
    For Each shape In sh.Shapes
        If shape.Left < right And shape.Left + shape.Width > left Then
            LogEvent "SLICER_OBSTACLE", shape.Name, "top=" & CStr(shape.Top) & "; bottom=" & CStr(shape.Top + shape.Height)
            top = SafeSlicerTop(top, shape.Top + shape.Height)
        End If
    Next shape
    firstRow = minimumRow
    Do While sh.Cells(firstRow, 14).Top < top
        firstRow = firstRow + 1
        If firstRow >= sh.Rows.Count Then Err.Raise vbObjectError + 760, , "No worksheet space for slicer."
    Loop
    lastRow = firstRow
    Do While sh.Range(sh.Cells(firstRow, 14), sh.Cells(lastRow, 14)).Height < 150
        lastRow = lastRow + 1
        If lastRow >= sh.Rows.Count Then Err.Raise vbObjectError + 760, , "No worksheet space for slicer."
    Loop
    lastColumn = 21
    Do While sh.Range(sh.Cells(firstRow, 14), sh.Cells(firstRow, lastColumn)).Width < 220
        lastColumn = lastColumn + 1
        If lastColumn >= sh.Columns.Count Then Err.Raise vbObjectError + 760, , "No worksheet space for slicer."
    Loop
    Set SlicerPlacement = sh.Range(sh.Cells(firstRow, 14), sh.Cells(lastRow, lastColumn))
    LogEvent "SLICER_RESERVED", sh.Name, SlicerPlacement.Address & "; top=" & CStr(SlicerPlacement.Top) & "; bottom=" & CStr(SlicerPlacement.Top + SlicerPlacement.Height)
End Function

Private Sub AssertSlicerSpace(ByVal area As Range)
    Dim shape As Shape, p As PivotTable, cell As Range
    If Application.WorksheetFunction.CountA(area) > 0 Then Err.Raise vbObjectError + 722, , "Slicer destination contains worksheet data: " & area.Address
    For Each cell In area.Cells
        If cell.MergeCells Then Err.Raise vbObjectError + 753, , "Slicer destination contains merged cells."
    Next cell
    For Each p In area.Parent.PivotTables
        If Not Application.Intersect(area, p.TableRange2) Is Nothing Then Err.Raise vbObjectError + 723, , "Slicer destination overlaps pivot: " & p.Name
    Next p
    For Each shape In area.Parent.Shapes
        If RectanglesOverlap(area.Left, area.Top, area.Width, area.Height, shape.Left, shape.Top, shape.Width, shape.Height) Then
            Err.Raise vbObjectError + 724, , "Slicer rectangle overlaps shape: " & shape.Name
        End If
    Next shape
End Sub

Private Function RectanglesOverlap(ByVal x1 As Double, ByVal y1 As Double, ByVal w1 As Double, ByVal h1 As Double, _
    ByVal x2 As Double, ByVal y2 As Double, ByVal w2 As Double, ByVal h2 As Double) As Boolean
    RectanglesOverlap = x1 < x2 + w2 And x1 + w1 > x2 And y1 < y2 + h2 And y1 + h1 > y2
End Function

Private Function SafeSlicerTop(ByVal proposed As Double, ByVal existingBottom As Double) As Double
    SafeSlicerTop = proposed
    If existingBottom + 24 > proposed Then SafeSlicerTop = existingBottom + 24
End Function

Private Sub EnsureTrendChart(ByVal p As PivotTable)
    Dim chart As ChartObject, n As Long, linked As PivotTable, sh As Worksheet
    Set sh = p.Parent
    For Each chart In sh.ChartObjects
        Set linked = Nothing
        On Error Resume Next
        Set linked = chart.Chart.PivotLayout.PivotTable
        On Error GoTo 0
        If Not linked Is Nothing Then
            If linked.Name = p.Name Then n = n + 1
        End If
    Next chart
    If n > 1 Then Err.Raise vbObjectError + 728, , "Multiple charts for trend pivot; resolve before rerunning."
    If n = 1 Then Exit Sub
    AssertEmpty sh.Range("E15:L38")
    Set chart = sh.ChartObjects.Add(sh.Range("E15").Left, sh.Range("E15").Top, 580, 300)
    chart.Name = "rlTrendChart"
    chart.Chart.SetSourceData p.TableRange1
    chart.Chart.ChartType = xlLineMarkers
    chart.Chart.HasTitle = True
    chart.Chart.ChartTitle.Text = "Monthly Net Sales vs Target"
End Sub

Private Sub BuildCards()
    Dim sh As Worksheet, p As PivotTable, row As Variant, i As Long, source As Range, address As String
    Set sh = GetSheet("Executive Summary")
    For Each row In specs
        If CStr(row(0)) = "pExecutive" Then Set p = FindPivot(row)
    Next row
    sh.Range("Z1").Value2 = OWNER
    sh.Range("A1").Value2 = "Revenue Leakage | Operating year 2025 | Snapshot 2025-12-31 | IDR"
    sh.Range("A8").Value2 = "Blank = N/A / unsupported filter. Diagnostic indicators overlap; do not add them into recoverable loss."
    For Each row In Split("net_sales|gross_profit|gross_margin_rate|target_net_sales|achievement_rate|price_shortfall|discount_amount|return_amount", "|")
        Set source = p.GetPivotData("[Measures].[" & row & "]")
        address = "'" & Replace(source.Parent.Name, "'", "''") & "'!" & source.Address
        sh.Cells(3 + (i \ 4) * 3, 1 + (i Mod 4) * 3).Value2 = Replace(CStr(row), "_", " ")
        With sh.Cells(4 + (i \ 4) * 3, 1 + (i Mod 4) * 3)
            .Formula = "=IF(ISBLANK(" & address & "),""""," & address & ")"
            If InStr(CStr(row), "rate") > 0 Then
                .NumberFormat = "0.00%"
            Else
                .NumberFormat = "#,##0.00"
            End If
            .Font.Bold = True
            .Font.Size = 14
            .EntireColumn.ColumnWidth = 24
        End With
        i = i + 1
    Next row
End Sub

Private Sub BuildMethodology()
    Dim sh As Worksheet, lines As Variant, i As Long
    Set sh = GetSheet("Methodology")
    sh.Range("Z1").Value2 = OWNER
    lines = Array("Methodology and limitations", "Synthetic distribution data; results demonstrate methods, not claims about a real business.", _
        "Default operating period: 2025-01-01 to 2025-12-31. AR/RFM snapshot: 2025-12-31.", _
        "Commercial grain: independent sale/return event. Net sales = billed sales minus event-period returns, once.", _
        "Target grain: branch-month. No invented volume or price targets; target gap is observed net sales minus target.", _
        "Targets/achievement are blank for SKU, outlet, salesman, channel, partial or noncontiguous calendar selections.", _
        "AR is an invoice snapshot; RFM is an outlet snapshot with fixed national scores. Operating year alone does not change these snapshots.", _
        "Product filters make AR, visits, RFM and cohort unsupported; salesman filters make RFM/cohort unsupported.", _
        "Price shortfall, discounts and returns are overlapping diagnostics, not an additive recoverable opportunity.", _
        "Invalid promotion discounts are already included in discounts. Do not count them twice.", _
        "Ratios use summed components. Zero denominators and unsupported contexts stay blank, not zero.", _
        "Cohort uses fixed onboarding denominators. Future/unobserved cells stay blank; do not total activity months.", _
        "Display: IDR/AOV 2 decimals; percentage 2 decimals; counts exact; days/productivity 2 decimals.", _
        "Reconcile with round-half-away-from-zero, matching PostgreSQL and Excel ROUND.", _
        "Negative amounts and return rates above 100% remain visible when supported by the event data.", _
        "Sources: docs/kpi_dictionary.md, docs/core_analysis.md, docs/optimization_report.md and model contract.", _
        "Refresh: Data > Refresh All. Keep large fact tables in the Data Model, not worksheet rows.", _
        "Native logs and SQL reconciliation are required before Phase 4 completion.")
    For i = LBound(lines) To UBound(lines)
        sh.Cells(i + 1, 1).Value2 = lines(i)
    Next i
    sh.Columns("A").ColumnWidth = 115
    sh.Range("A1:A24").WrapText = True
    sh.Range("A1:A24").Rows.AutoFit
End Sub

Private Sub FormatOwnedSheets()
    Dim row As Variant, sh As Worksheet, p As PivotTable
    For Each row In specs
        Set p = FindPivot(row)
        If Not p Is Nothing Then
            p.TableRange1.Columns.AutoFit
            p.TableRange1.Rows(1).Interior.Color = RGB(20, 45, 65)
            p.TableRange1.Rows(1).Font.Color = vbWhite
        End If
    Next row
    Set sh = GetSheet("Executive Summary")
    sh.Range("A1:L1").Interior.Color = RGB(20, 45, 65)
    sh.Range("A1:L1").Font.Color = vbWhite
    Set sh = GetSheet("Leakage Diagnostics")
    sh.Range("Z1").Value2 = OWNER
    sh.Range("A1").Value2 = "Revenue Leakage Diagnostics | IDR"
    sh.Range("A2").Value2 = "SKU/outlet tables: top 10 by price shortfall. Discounts and returns overlap; do not add these into recoverable loss."
    Set sh = GetSheet("Operations")
    sh.Range("Z1").Value2 = OWNER
    sh.Range("A1").Value2 = "Operations | AR, Sales Visits, RFM and Cohort"
    sh.Range("A2").Value2 = "AR/RFM snapshot: 2025-12-31. Blank = unsupported / N/A. Cohort activity-month totals are disabled."
End Sub

Private Sub PreflightVerification()
    Dim sh As Worksheet, p As PivotTable
    Set sh = ExistingSheet("Native Verification")
    Set p = NamedPivot("rlNativeProbe")
    If Not sh Is Nothing Then
        If sh.Range("Z1").Value2 <> OWNER Then Err.Raise vbObjectError + 730, , "Native Verification sheet exists without runner ownership."
        If p Is Nothing Then AssertEmpty sh.Range("A5:L200")
    ElseIf Not p Is Nothing Then
        Err.Raise vbObjectError + 731, , "rlNativeProbe exists outside its dedicated sheet."
    End If
    If Not p Is Nothing Then
        If p.Parent.Name <> "Native Verification" Or Not p.PivotCache.OLAP Then Err.Raise vbObjectError + 732, , "Native probe has invalid source/location."
        AssertProbeIsolated p
    End If
End Sub

Private Sub AssertProbeIsolated(ByVal p As PivotTable)
    Dim sc As SlicerCache, connected As Object
    For Each sc In wb.SlicerCaches
        For Each connected In sc.PivotTables
            If connected.Name = p.Name Then Err.Raise vbObjectError + 733, , "Native probe must not be connected to any dashboard slicer."
        Next connected
    Next sc
End Sub

Private Sub ApplyCase(ByVal p As PivotTable, ByVal row As Variant)
    Dim dates As String, d As Date, last As Date
    If Len(CStr(row(11))) > 0 Then
        dates = CStr(row(11))
    ElseIf Len(CStr(row(1))) > 0 Then
        d = IsoDate(CStr(row(1)))
        last = IsoDate(CStr(row(2)))
        Do While d <= last
            If Len(dates) > 0 Then dates = dates & "|"
            dates = dates & Format$(d, "yyyy-mm-dd")
            d = DateAdd("d", 1, d)
        Loop
    End If
    SetMembers p, "DimDate.date", dates
    SetMembers p, "DimBranch.branch_id", CStr(row(3))
    SetMembers p, "DimProduct.sku_id", CStr(row(4))
    SetMembers p, "DimOutlet.outlet_id", CStr(row(5))
    SetMembers p, "DimChannel.channel", CStr(row(6))
    SetMembers p, "DimSalesman.salesman_id", CStr(row(7))
    SetMembers p, "FactCohort.cohort_month", CStr(row(8))
    SetMembers p, "FactCohort.activity_month", CStr(row(9))
    SetMembers p, "FactRFM.segment", CStr(row(10))
End Sub

Private Function IsoDate(ByVal text As String) As Date
    Dim parts As Variant
    parts = Split(text, "-")
    If UBound(parts) <> 2 Then Err.Raise vbObjectError + 734, , "Invalid ISO date."
    IsoDate = DateSerial(CInt(parts(0)), CInt(parts(1)), CInt(parts(2)))
    If Format$(IsoDate, "yyyy-mm-dd") <> text Then Err.Raise vbObjectError + 735, , "Invalid ISO date: " & text
End Function

Private Function InvariantValue(ByVal value As Variant) As String
    If IsError(value) Then Err.Raise vbObjectError + 736, , "Native result is an Excel error."
    If IsEmpty(value) Or IsNull(value) Then Exit Function
    If VarType(value) = vbString Then
        If Len(value) = 0 Then Exit Function
        Err.Raise vbObjectError + 737, , "Native numeric result is text."
    End If
    If Not IsNumeric(value) Then Err.Raise vbObjectError + 738, , "Native result is not numeric."
    InvariantValue = ExpandNumber(Trim$(Str$(CDbl(value))))
End Function

Private Function ExpandNumber(ByVal text As String) As String
    Dim e As Long, exponent As Long, point As Long, sign As String, digits As String
    If Left$(text, 1) = "-" Then
        sign = "-"
        text = Mid$(text, 2)
    End If
    e = InStr(1, text, "E", vbTextCompare)
    If e > 0 Then
        exponent = CLng(Mid$(text, e + 1))
        text = Left$(text, e - 1)
    End If
    point = InStr(text, ".")
    If point = 0 Then point = Len(text) + 1
    digits = Replace(text, ".", "")
    point = point - 1 + exponent
    If point <= 0 Then
        text = "0." & String$(-point, "0") & digits
    ElseIf point >= Len(digits) Then
        text = digits & String$(point - Len(digits), "0")
    Else
        text = Left$(digits, point) & "." & Mid$(digits, point + 1)
    End If
    ExpandNumber = sign & text
End Function

Private Sub RefreshAndCheck(ByVal label As String)
    Dim row As Variant, t As Object, sh As Worksheet, qt As QueryTable
    LogEvent "REFRESH_STARTED", label, wb.Queries("pFolder").Formula
    ' Model.Refresh updates model sources; SQLReference worksheet refresh is synchronous.
    wb.Model.Refresh
    For Each sh In wb.Worksheets
        For Each qt In sh.QueryTables
            If Not qt.Refresh(BackgroundQuery:=False) Then Err.Raise vbObjectError + 739, , "Worksheet query refresh returned False."
        Next qt
    Next sh
    For Each row In ReadCsv(base & "\dashboard-specs\tables.csv")
        Set t = wb.Model.ModelTables(CStr(row(0)))
        If t.RecordCount <> CLng(row(1)) Then Err.Raise vbObjectError + 740, , "Post-refresh model count mismatch: " & row(0)
        LogEvent "REFRESH_COUNT", CStr(row(0)), CStr(t.RecordCount)
    Next row
    LogEvent "REFRESH_RETURNED", label, "API returned; counts match. Inspect query UI and rerun native numeric captures."
    CaptureRefreshReference label
End Sub

Private Sub CaptureRefreshReference(ByVal label As String)
    Dim sh As Worksheet, p As PivotTable, row As Variant, req As Variant, requests As Collection
    Dim handle As Integer, opened As Boolean, value As Variant, number As Long, description As String
    On Error GoTo Failed
    Set sh = GetSheet("Native Verification")
    sh.Range("Z1").Value2 = OWNER
    Set p = NamedPivot("rlNativeProbe")
    If p Is Nothing Then Set p = NewPivot(sh, "A35", "rlNativeProbe")
    AssertProbeIsolated p
    Set requests = ReadCsv(base & "\dashboard-specs\requests.csv")
    For Each req In requests
        If req(0) = "reference" Then EnsureMeasure p, CStr(req(1))
    Next req
    p.DataPivotField.Orientation = xlRowField
    p.ClearAllFilters
    For Each row In ReadCsv(base & "\dashboard-specs\cases.csv")
        If row(0) = "reference" Then
            If CStr(row(1)) = Left$(CStr(row(1)), 4) & "-01-01" And CStr(row(2)) = Left$(CStr(row(1)), 4) & "-12-31" Then
                SetMembers p, "DimDate.year", Left$(CStr(row(1)), 4)
            Else
                ApplyCase p, row
            End If
        End If
    Next row
    If Not p.RefreshTable Then Err.Raise vbObjectError + 755, , "Refresh-control pivot returned False."
    handle = FreeFile
    Open runPath & "\refresh-" & label & ".partial.csv" For Output As #handle
    opened = True
    Print #handle, "case_id,metric_id,actual_value"
    For Each req In requests
        If req(0) = "reference" Then
            value = p.GetPivotData("[Measures].[" & CStr(req(1)) & "]").Value2
            Print #handle, "reference," & Csv(CStr(req(1))) & "," & InvariantValue(value)
        End If
    Next req
    Close #handle
    opened = False
    Name runPath & "\refresh-" & label & ".partial.csv" As runPath & "\refresh-" & label & ".csv"
    LogEvent "REFRESH_CAPTURED", label, "46 native reference metrics captured; compare with --refresh validator."
    Exit Sub
Failed:
    number = Err.Number
    description = Err.Description
    If opened Then Close #handle
    Err.Raise number, "CaptureRefreshReference", description
End Sub

Private Function FolderLiteral(ByVal formula As String) As String
    Dim text As String, closing As Long
    text = Trim$(formula)
    If Left$(text, 1) <> """" Then Err.Raise vbObjectError + 741, , "pFolder must be a text literal (optionally followed by meta); restore original query if needed."
    closing = InStr(2, text, """")
    If closing = 0 Then Err.Raise vbObjectError + 742, , "Invalid pFolder literal."
    FolderLiteral = Mid$(text, 2, closing - 2)
End Function

Private Function MFolder(ByVal path As String) As String
    MFolder = """" & Replace(path, """", """""") & """ meta [IsParameterQuery=true, Type=""Text"", IsParameterQueryRequired=true]"
End Function

Private Function ReadText(ByVal path As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    ReadText = stream.ReadText
    stream.Close
End Function

Private Sub WriteText(ByVal path As String, ByVal text As String)
    Dim stream As Object
    ' Inventory runs before and after building. An identical immutable snapshot is reusable.
    If fso.FileExists(path) Then
        If ReadText(path) = text Then
            LogEvent "SNAPSHOT_REUSED", path, "Existing contents match; no overwrite needed."
            Exit Sub
        End If
        Err.Raise vbObjectError + 761, , "Snapshot already exists with different contents: " & path
    End If
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile path, 1
    stream.Close
End Sub

Private Function ReadCsv(ByVal path As String) As Collection
    Dim text As String, c As String, field As String, row As String, quoted As Boolean
    Dim i As Long, rows As Collection, fields As Collection, data As Variant, j As Long, header As Boolean, columns As Long
    Set rows = New Collection
    Set fields = New Collection
    text = ReadText(path)
    If Left$(text, 1) = ChrW(&HFEFF) Then text = Mid$(text, 2)
    If Right$(text, 1) <> vbLf Then text = text & vbLf
    For i = 1 To Len(text)
        c = Mid$(text, i, 1)
        If c = """" Then
            If quoted And Mid$(text, i + 1, 1) = """" Then
                field = field & """"
                i = i + 1
            Else
                quoted = Not quoted
            End If
        ElseIf c = "," And Not quoted Then
            fields.Add field
            field = ""
        ElseIf c = vbLf And Not quoted Then
            If Right$(field, 1) = vbCr Then field = Left$(field, Len(field) - 1)
            fields.Add field
            If header And fields.Count <> columns Then Err.Raise vbObjectError + 754, , "CSV column count mismatch: " & path
            columns = fields.Count
            ReDim data(0 To fields.Count - 1)
            For j = 1 To fields.Count
                data(j - 1) = fields(j)
            Next j
            If header Then rows.Add data
            header = True
            Set fields = New Collection
            field = ""
        Else
            field = field & c
        End If
    Next i
    If quoted Then Err.Raise vbObjectError + 743, , "Unclosed CSV quote: " & path
    Set ReadCsv = rows
End Function

Private Function Csv(ByVal text As String) As String
    Csv = """" & Replace(text, """", """""") & """"
End Function

Private Sub LogEvent(ByVal eventName As String, ByVal objectName As String, ByVal detail As String)
    If logging Then Print #logHandle, Csv(Format$(Now, "yyyy-mm-dd hh:nn:ss")) & "," & Csv(eventName) & "," & Csv(objectName) & "," & Csv(detail)
End Sub

Private Sub FinishRun(ByVal message As String)
    LogEvent "STAGE_FINISHED", stage, message
    If logging Then Close #logHandle
    logging = False
    Application.StatusBar = oldStatus
    MsgBox message & vbCrLf & "Evidence: " & runPath & vbCrLf & "Phase 4 remains open until all native acceptance checks pass.", vbInformation, "Phase 4 runner"
End Sub

Private Sub FailRun()
    ReportFailure Err.Number, Err.Description
End Sub

Private Sub ReportFailure(ByVal number As Long, ByVal description As String)
    Dim failure As String
    failure = "Stopped at: " & stage & vbCrLf & "Error " & CStr(number) & ": " & description
    LogEvent "STOPPED", stage, failure
    If logging Then Close #logHandle
    logging = False
    Application.StatusBar = oldStatus
    MsgBox failure & vbCrLf & "Evidence: " & runPath & vbCrLf & "Partial additions may remain. Do not delete existing objects or mark Phase 4 complete.", vbExclamation, "Phase 4 runner stopped"
End Sub

Private Function IsNameCharacter(ByVal value As String) As Boolean
    Dim n As Long
    If Len(value) <> 1 Then Exit Function
    n = AscW(value)
    IsNameCharacter = (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n >= 48 And n <= 57) Or n = 95
End Function

Private Function IsPlainName(ByVal value As String) As Boolean
    Dim i As Long
    If Len(value) = 0 Then Exit Function
    If InStr("0123456789", Left$(value, 1)) > 0 Then Exit Function
    For i = 1 To Len(value)
        If Not IsNameCharacter(Mid$(value, i, 1)) Then Exit Function
    Next i
    IsPlainName = True
End Function

Private Function QuotedEnd(ByVal text As String, ByVal startAt As Long, ByVal delimiter As String) As Long
    Dim i As Long
    i = startAt + 1
    Do While i <= Len(text)
        If Mid$(text, i, 1) = delimiter Then
            If Mid$(text, i + 1, 1) = delimiter Then
                i = i + 2
            Else
                QuotedEnd = i
                Exit Function
            End If
        Else
            i = i + 1
        End If
    Loop
    Err.Raise vbObjectError + 134, , "Unclosed quoted literal or identifier."
End Function

Private Function CanonicalFormula(ByVal formula As String) As String
    Dim i As Long
    Dim ending As Long
    Dim token As String
    Dim c As String
    Dim result As String
    formula = Trim$(formula)
    If Left$(formula, 1) = "=" Then formula = Mid$(formula, 2)
    i = 1
    Do While i <= Len(formula)
        c = Mid$(formula, i, 1)
        Select Case c
            Case """"
                ending = QuotedEnd(formula, i, c)
                result = result & Mid$(formula, i, ending - i + 1)
                i = ending + 1
            Case "'"
                ending = QuotedEnd(formula, i, c)
                token = Mid$(formula, i + 1, ending - i - 1)
                If IsPlainName(token) Then
                    result = result & LCase$(token)
                Else
                    result = result & LCase$(Mid$(formula, i, ending - i + 1))
                End If
                i = ending + 1
            Case "["
                ending = InStr(i + 1, formula, "]")
                If ending = 0 Then Err.Raise vbObjectError + 135, , "Unclosed bracket in formula."
                result = result & LCase$(Mid$(formula, i, ending - i + 1))
                i = ending + 1
            Case " ", vbTab, vbCr, vbLf
                i = i + 1
            Case ";"
                result = result & ","
                i = i + 1
            Case Else
                result = result & LCase$(c)
                i = i + 1
        End Select
    Loop
    CanonicalFormula = result
End Function
