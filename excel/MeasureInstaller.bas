Attribute VB_Name = "MeasureInstaller"
Option Explicit

' Run manually from excel\MeasureInstaller.xlsm, never from the target workbook.
' Source: measures.dax. No database access, refresh, save, deletion or UI automation.
Private Const EXPECTED_COUNT As Long = 66
Private Const TARGET_NAME As String = "RevenueLeakage.xlsx"
Private Const HOME_TABLE As String = "ReportContext"
Private Const REQUIRED_TABLES As String = "DimBranch,DimChannel,DimDate,DimOutlet,DimProduct,DimSalesman,FactCommercial,FactVisits,FactInvoiceSnapshot,FactTargets,FactPayments,FactRFM,FactCohort,ReportContext"

Public Sub InstallMeasures()
    Dim target As Workbook
    Dim candidate As Workbook
    Dim specs As Collection
    Dim tables As Object
    Dim existing As Object
    Dim names As Object
    Dim modelMeasure As Object
    Dim home As Object
    Dim info As Object
    Dim fso As Object
    Dim spec As Variant
    Dim key As Variant
    Dim text As String
    Dim basePath As String
    Dim targetPath As String
    Dim logPath As String
    Dim backupPath As String
    Dim runId As String
    Dim sourcePath As String
    Dim currentName As String
    Dim stage As String
    Dim message As String
    Dim failure As String
    Dim logFile As Integer
    Dim logOpen As Boolean
    Dim added As Long
    Dim skipped As Long
    Dim conflicts As Long
    Dim i As Long
    On Error GoTo Failed

    stage = "validate runner"
    If StrComp(ThisWorkbook.Name, "MeasureInstaller.xlsm", vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 100, , "Import this module into a separate workbook named MeasureInstaller.xlsm."
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    basePath = ResolveSourceFolder(fso, ThisWorkbook.Path)
    If Len(basePath) = 0 Then Exit Sub
    EnsureFolder fso, basePath & "\measure-install-logs"
    runId = UniqueRunId(fso, basePath)
    logPath = basePath & "\measure-install-logs\" & runId & ".csv"
    logFile = FreeFile
    Open logPath For Output As #logFile
    logOpen = True
    Print #logFile, "timestamp,event,measure,detail,formula"
    WriteLog logFile, "START", "", "Native Excel installer started; runner=" & ThisWorkbook.FullName

    stage = "validate target"
    targetPath = fso.GetAbsolutePathName(basePath & "\" & TARGET_NAME)
    For Each candidate In Application.Workbooks
        If StrComp(candidate.Name, TARGET_NAME, vbTextCompare) = 0 Then
            If IsCloudPath(candidate.FullName) Then
                ' A URL cannot be equated to a local sync path by filename alone.
                ' Ask the operator to identify the open cloud workbook explicitly.
                If MsgBox("Excel reports the target as a OneDrive/SharePoint URL:" & vbCrLf & candidate.FullName & vbCrLf & vbCrLf & _
                    "Source folder: " & basePath & vbCrLf & _
                    "Is this the RevenueLeakage.xlsx workbook containing your Phase 4 model?" & vbCrLf & _
                    "Choose No if unsure. No measures have been changed.", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm cloud target") <> vbYes Then
                    WriteLog logFile, "CANCELLED", "", "Cloud workbook identity was not confirmed; no target changes."
                    Close #logFile
                    logOpen = False
                    Exit Sub
                End If
                WriteLog logFile, "CLOUD_TARGET_CONFIRMED", "", "Operator explicitly identified open target: " & candidate.FullName
            ElseIf StrComp(fso.GetAbsolutePathName(candidate.FullName), targetPath, vbTextCompare) <> 0 Then
                Err.Raise vbObjectError + 103, , "An open RevenueLeakage.xlsx has a different path: " & candidate.FullName
            End If
            Set target = candidate
        End If
    Next candidate
    If target Is Nothing Then Err.Raise vbObjectError + 104, , "Open the local target workbook first: " & targetPath
    If target Is ThisWorkbook Then Err.Raise vbObjectError + 105, , "Runner and target must be different workbooks."
    If target.ReadOnly Then Err.Raise vbObjectError + 106, , "Target workbook is read-only."
    If target.FileFormat <> xlOpenXMLWorkbook Then Err.Raise vbObjectError + 107, , "Target must be a macro-free .xlsx workbook."
    If Not target.Saved Then Err.Raise vbObjectError + 108, , "Save RevenueLeakage.xlsx with Ctrl+S, then run again."
    WriteLog logFile, "TARGET", "", target.FullName
    WriteLog logFile, "LOCALE", "", "List separator=" & Application.International(xlListSeparator) & "; decimal separator=" & Application.International(xlDecimalSeparator)

    stage = "parse source"
    sourcePath = basePath & "\measures.dax"
    text = ReadUtf8(sourcePath)
    Set specs = ParseDefinitions(text)
    If specs.Count <> EXPECTED_COUNT Then
        Err.Raise vbObjectError + 109, , "Expected 66 definitions; found " & specs.Count & ". No measures changed."
    End If
    fso.CopyFile sourcePath, basePath & "\measure-install-logs\" & runId & ".source.dax", False
    SelfTestCore

    stage = "preflight model and dependencies"
    Set tables = NewDictionary()
    For Each home In target.Model.ModelTables
        tables.Add home.Name, home
    Next home
    For Each key In Split(REQUIRED_TABLES, ",")
        If Not tables.Exists(CStr(key)) Then Err.Raise vbObjectError + 110, , "Required model table missing: " & CStr(key)
    Next key
    Set home = tables.Item(HOME_TABLE)
    ValidateReferences specs, tables
    For i = 1 To specs.Count
        spec = specs.Item(i)
        For Each key In home.ModelTableColumns
            If StrComp(key.Name, CStr(spec(0)), vbTextCompare) = 0 Then
                Err.Raise vbObjectError + 117, , "Measure name conflicts with a home-table column: " & spec(0)
            End If
        Next key
    Next i

    Set existing = NewDictionary()
    For Each modelMeasure In target.Model.ModelMeasures
        existing.Add modelMeasure.Name, modelMeasure
    Next modelMeasure
    If existing.Exists("_shapshot_supported") Then
        WriteLog logFile, "WARNING_TYPO", "_shapshot_supported", "Preserved. Canonical name is _snapshot_supported. Do not count this extra name toward the expected 66."
    End If

    ' Compare ALL existing formulas before making ANY target change.
    For i = 1 To specs.Count
        spec = specs.Item(i)
        currentName = CStr(spec(0))
        If existing.Exists(currentName) Then
            Set modelMeasure = existing.Item(currentName)
            If CanonicalFormula(CStr(modelMeasure.Formula)) <> CanonicalFormula(CStr(spec(1))) Then
                conflicts = conflicts + 1
                WriteLog logFile, "CONFLICT_ACTUAL", currentName, "Existing formula differs; preserved.", CStr(modelMeasure.Formula)
                WriteLog logFile, "CONFLICT_EXPECTED", currentName, "Resolve manually before rerunning.", CStr(spec(1))
            Else
                WriteLog logFile, "PREFLIGHT_MATCH", currentName, "Existing formula matches; existing format will be preserved.", CStr(modelMeasure.Formula)
                If StrComp(modelMeasure.AssociatedTable.Name, HOME_TABLE, vbTextCompare) <> 0 Then
                    WriteLog logFile, "WARNING_HOME_TABLE", currentName, "Existing home table preserved: " & modelMeasure.AssociatedTable.Name
                End If
            End If
        End If
    Next i
    If conflicts > 0 Then Err.Raise vbObjectError + 111, , CStr(conflicts) & " formula conflict(s). No measures changed. Read the log."
    WriteLog logFile, "PREFLIGHT_OK", "", "66 definitions validated; required tables/columns and dependency order checked."

    stage = "backup"
    EnsureFolder fso, basePath & "\measure-backups"
    backupPath = basePath & "\measure-backups\RevenueLeakage-before-measures-" & runId & ".xlsx"
    target.SaveCopyAs backupPath
    If Not fso.FileExists(backupPath) Then Err.Raise vbObjectError + 112, , "Backup was not created. No measures changed."
    If fso.GetFile(backupPath).Size = 0 Then Err.Raise vbObjectError + 113, , "Backup is empty. No measures changed."
    WriteLog logFile, "BACKUP_OK", "", backupPath

    stage = "add measures"
    For i = 1 To specs.Count
        spec = specs.Item(i)
        currentName = CStr(spec(0))
        If existing.Exists(currentName) Then
            skipped = skipped + 1
            WriteLog logFile, "SKIPPED", currentName, "Matching existing measure; no formula or format changes."
        Else
            Set info = MakeFormat(target.Model, FormatKind(currentName))
            ' Native Add supplies the assignment '=' itself; the dialog-style
            ' formula retained in specs/readback must not be passed verbatim.
            text = NativeAddFormula(CStr(spec(1)), CStr(Application.International(xlListSeparator)), CStr(Application.International(xlDecimalSeparator)))
            WriteLog logFile, "ADD_REQUEST", currentName, "Localized expression sent to native Add.", text
            Set modelMeasure = target.Model.ModelMeasures.Add(currentName, home, text, info, "Phase 4 KPI contract; source: excel/measures.dax")
            added = added + 1
            WriteLog logFile, "ADDED", currentName, "Native ModelMeasures.Add returned; format=" & FormatKind(currentName), CStr(modelMeasure.Formula)
            If CanonicalFormula(CStr(modelMeasure.Formula)) <> CanonicalFormula(CStr(spec(1))) Then
                Err.Raise vbObjectError + 114, , "Readback formula differs after adding " & currentName
            End If
            existing.Add currentName, modelMeasure
        End If
    Next i

    stage = "verify final names and formulas"
    Set names = NewDictionary()
    For Each modelMeasure In target.Model.ModelMeasures
        names.Add modelMeasure.Name, modelMeasure
    Next modelMeasure
    For i = 1 To specs.Count
        spec = specs.Item(i)
        currentName = CStr(spec(0))
        If Not names.Exists(currentName) Then Err.Raise vbObjectError + 115, , "Missing final measure: " & currentName
        Set modelMeasure = names.Item(currentName)
        If CanonicalFormula(CStr(modelMeasure.Formula)) <> CanonicalFormula(CStr(spec(1))) Then
            Err.Raise vbObjectError + 116, , "Final formula mismatch: " & currentName
        End If
    Next i
    WriteLog logFile, "DONE", "", "Expected measures=66; added=" & added & "; skipped=" & skipped & "; total native measures=" & names.Count
    WriteLog logFile, "NATIVE_ACCEPTANCE_PENDING", "", "No SQL reconciliation, refresh, slicer or save/reopen checks performed. Save target manually."
    Close #logFile
    logOpen = False
    MsgBox "Measures installed: " & added & vbCrLf & "Matching measures skipped: " & skipped & vbCrLf & _
        "All 66 expected names/formulas are present." & vbCrLf & _
        "Save RevenueLeakage.xlsx with Ctrl+S, then run again to check idempotency." & vbCrLf & _
        "Next: native PivotTable/SQL verification (stage 11)." & vbCrLf & "Log: " & logPath, vbInformation, "Measure installer complete"
    Exit Sub

Failed:
    failure = "Error " & CStr(Err.Number) & ": " & Err.Description
    message = "Stopped at: " & stage & vbCrLf & "Measure: " & currentName & vbCrLf & failure & vbCrLf & _
        "Added before stop: " & added & ". Existing measures were not overwritten or deleted."
    On Error Resume Next
    If logOpen Then
        WriteLog logFile, "FAILED", currentName, "Stage=" & stage & "; added=" & added & "; " & failure
        Close #logFile
    End If
    On Error GoTo 0
    If Len(backupPath) > 0 Then message = message & vbCrLf & "Backup path (check BACKUP_OK in log): " & backupPath
    If Len(logPath) > 0 Then message = message & vbCrLf & "Log: " & logPath
    MsgBox message & vbCrLf & "Partial additions may remain. Resolve the error before rerunning; do not mark Phase 4 complete.", vbExclamation, "Measure installer stopped"
End Sub

Private Function ResolveSourceFolder(ByVal fso As Object, ByVal runnerPath As String) As String
    Dim picker As Object
    Dim selected As String
    selected = runnerPath
    If Len(selected) > 0 And Not IsCloudPath(selected) Then
        If fso.FolderExists(selected) Then
            If fso.FileExists(selected & "\measures.dax") And fso.FileExists(selected & "\" & TARGET_NAME) Then
                ResolveSourceFolder = fso.GetAbsolutePathName(selected)
                Exit Function
            End If
        End If
    End If
    Set picker = Application.FileDialog(4) ' msoFileDialogFolderPicker; no extra reference.
    picker.Title = "Select local project excel folder containing measures.dax and RevenueLeakage.xlsx"
    picker.AllowMultiSelect = False
    If picker.Show <> -1 Then Exit Function
    selected = CStr(picker.SelectedItems(1))
    If IsCloudPath(selected) Or Not fso.FolderExists(selected) Then
        Err.Raise vbObjectError + 101, , "Select a local filesystem folder, not a OneDrive web URL."
    End If
    selected = fso.GetAbsolutePathName(selected)
    If Not fso.FileExists(selected & "\measures.dax") Or Not fso.FileExists(selected & "\" & TARGET_NAME) Then
        Err.Raise vbObjectError + 102, , "Selected folder must contain both measures.dax and RevenueLeakage.xlsx."
    End If
    ResolveSourceFolder = selected
End Function

Private Function NewDictionary() As Object
    Set NewDictionary = CreateObject("Scripting.Dictionary")
    NewDictionary.CompareMode = vbTextCompare
End Function

Private Function ReadUtf8(ByVal filePath As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadUtf8 = stream.ReadText
    stream.Close
End Function

Private Sub EnsureFolder(ByVal fso As Object, ByVal folderPath As String)
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
End Sub

Private Function UniqueRunId(ByVal fso As Object, ByVal basePath As String) As String
    Dim stamp As String
    Dim value As String
    Dim suffix As Long
    stamp = Format$(Now, "yyyymmdd-hhnnss")
    value = stamp
    Do While fso.FileExists(basePath & "\measure-install-logs\" & value & ".csv") Or _
             fso.FileExists(basePath & "\measure-install-logs\" & value & ".source.dax") Or _
             fso.FileExists(basePath & "\measure-backups\RevenueLeakage-before-measures-" & value & ".xlsx")
        suffix = suffix + 1
        value = stamp & "-" & CStr(suffix)
    Loop
    UniqueRunId = value
End Function

Private Sub WriteLog(ByVal handle As Integer, ByVal eventName As String, ByVal measureName As String, ByVal detail As String, Optional ByVal formula As String = "")
    Print #handle, CsvField(Format$(Now, "yyyy-mm-dd hh:nn:ss")) & "," & CsvField(eventName) & "," & _
        CsvField(measureName) & "," & CsvField(detail) & "," & CsvField(formula)
End Sub

Private Function CsvField(ByVal value As String) As String
    CsvField = """" & Replace(value, """", """""") & """"
End Function

Private Function ParseDefinitions(ByVal text As String) As Collection
    Dim result As New Collection
    Dim names As Object
    Dim lines As Variant
    Dim spec As Variant
    Dim i As Long
    Set names = NewDictionary()
    lines = Split(Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For i = LBound(lines) To UBound(lines)
        spec = ParseDefinitionLine(CStr(lines(i)))
        If Not IsEmpty(spec) Then
            If names.Exists(CStr(spec(0))) Then Err.Raise vbObjectError + 120, , "Duplicate source measure on line " & i + 1 & ": " & spec(0)
            names.Add CStr(spec(0)), True
            result.Add spec
        End If
    Next i
    Set ParseDefinitions = result
End Function

Private Sub ValidateReferences(ByVal specs As Collection, ByVal tables As Object)
    Dim earlier As Object
    Dim columns As Object
    Dim refs As Collection
    Dim spec As Variant
    Dim ref As Variant
    Dim table As Object
    Dim column As Object
    Dim key As Variant
    Dim i As Long
    Dim j As Long
    Set earlier = NewDictionary()
    Set columns = NewDictionary()
    For Each key In tables.Keys
        Set table = tables.Item(CStr(key))
        For Each column In table.ModelTableColumns
            columns.Add table.Name & "[" & column.Name & "]", True
        Next column
    Next key
    For i = 1 To specs.Count
        spec = specs.Item(i)
        Set refs = BracketReferences(CStr(spec(1)))
        For j = 1 To refs.Count
            ref = refs.Item(j)
            If Len(CStr(ref(0))) > 0 Then
                If Not columns.Exists(CStr(ref(0)) & "[" & CStr(ref(1)) & "]") Then
                    Err.Raise vbObjectError + 121, , "Missing source table/column for " & spec(0) & ": " & ref(0) & "[" & ref(1) & "]"
                End If
            ElseIf Not earlier.Exists(CStr(ref(1))) Then
                Err.Raise vbObjectError + 122, , "Missing/forward measure dependency for " & spec(0) & ": " & ref(1)
            End If
        Next j
        earlier.Add CStr(spec(0)), True
    Next i
End Sub

Private Function BracketReferences(ByVal formula As String) As Collection
    Dim refs As New Collection
    Dim masked As String
    Dim prefix As String
    Dim startAt As Long
    Dim endAt As Long
    Dim beforeAt As Long
    masked = MaskStrings(formula)
    startAt = InStr(1, masked, "[")
    Do While startAt > 0
        endAt = InStr(startAt + 1, masked, "]")
        If endAt = 0 Then Err.Raise vbObjectError + 123, , "Unclosed column/measure reference."
        beforeAt = startAt - 1
        Do While beforeAt > 0
            If Not IsNameCharacter(Mid$(masked, beforeAt, 1)) Then Exit Do
            beforeAt = beforeAt - 1
        Loop
        prefix = Mid$(masked, beforeAt + 1, startAt - beforeAt - 1)
        refs.Add Array(prefix, Mid$(masked, startAt + 1, endAt - startAt - 1))
        startAt = InStr(endAt + 1, masked, "[")
    Loop
    Set BracketReferences = refs
End Function

Private Function FormatKind(ByVal measureName As String) As String
    If Right$(measureName, 5) = "_rate" Then
        FormatKind = "percentage"
    Else
        Select Case measureName
            Case "_snapshot_supported", "_visit_supported", "_outlet_snapshot_supported", "_target_supported", "_target_coverage"
                FormatKind = "boolean"
            Case "_sales_lines", "_completed_calls", "_converted_calls", "_salesman_days", "sales_units", "returned_units", "active_outlets", "orders", "scheduled_visits", "completed_visits", "effective_calls", "rfm_outlets", "cohort_outlets", "cohort_active_outlets"
                FormatKind = "whole"
            Case Else
                FormatKind = "decimal"
        End Select
    End If
End Function

Private Function MakeFormat(ByVal model As Object, ByVal kind As String) As Object
    Dim info As Object
    Select Case kind
        Case "boolean"
            Set info = model.ModelFormatBoolean
        Case "whole"
            Set info = model.ModelFormatWholeNumber
            info.UseThousandSeparator = True
        Case "percentage"
            Set info = model.ModelFormatPercentageNumber
            info.DecimalPlaces = 2
            info.UseThousandSeparator = True
        Case Else
            Set info = model.ModelFormatDecimalNumber
            info.DecimalPlaces = 2
            info.UseThousandSeparator = True
    End Select
    Set MakeFormat = info
End Function

' BEGIN PURE CORE: also tested outside Excel without accessing any workbook.
Private Function NativeAddFormula(ByVal formula As String, ByVal listSeparator As String, ByVal decimalSeparator As String) As String
    Dim i As Long
    Dim ending As Long
    Dim result As String
    Dim c As String
    formula = Trim$(formula)
    If Left$(formula, 1) = "=" Then formula = Trim$(Mid$(formula, 2))
    If Len(formula) = 0 Or Left$(formula, 1) = "=" Then
        Err.Raise vbObjectError + 136, , "Native Add requires a nonempty expression without an assignment equals sign."
    End If
    If listSeparator <> "," And listSeparator <> ";" Then Err.Raise vbObjectError + 137, , "Unsupported DAX list separator."
    If decimalSeparator <> "." And decimalSeparator <> "," Then Err.Raise vbObjectError + 138, , "Unsupported DAX decimal separator."
    ' Power Pivot's native Add parses localized numeric tokens. A comma
    ' immediately after a number must not be misread as a decimal mark.
    If decimalSeparator = "," Then listSeparator = ";"
    i = 1
    Do While i <= Len(formula)
        c = Mid$(formula, i, 1)
        Select Case c
            Case """", "'"
                ending = QuotedEnd(formula, i, c)
                result = result & Mid$(formula, i, ending - i + 1)
                i = ending + 1
            Case "["
                ending = InStr(i + 1, formula, "]")
                If ending = 0 Then Err.Raise vbObjectError + 135, , "Unclosed bracket in native formula."
                result = result & Mid$(formula, i, ending - i + 1)
                i = ending + 1
            Case ","
                result = result & listSeparator
                i = i + 1
            Case "."
                result = result & decimalSeparator
                i = i + 1
            Case Else
                result = result & c
                i = i + 1
        End Select
    Loop
    NativeAddFormula = result
End Function

Private Function IsCloudPath(ByVal value As String) As Boolean
    value = LCase$(Trim$(value))
    IsCloudPath = Left$(value, 8) = "https://" Or Left$(value, 7) = "http://"
End Function

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

Private Function ParseDefinitionLine(ByVal line As String) As Variant
    Dim at As Long
    Dim measureName As String
    Dim formula As String
    line = Trim$(Replace(line, vbTab, " "))
    If Len(line) > 0 Then
        If AscW(Left$(line, 1)) = -257 Then line = Mid$(line, 2)
    End If
    If Len(line) = 0 Or Left$(line, 2) = "//" Then Exit Function
    at = InStr(1, line, ":=", vbBinaryCompare)
    If at < 2 Then Err.Raise vbObjectError + 130, , "Expected one name := expression per non-comment line."
    measureName = Trim$(Left$(line, at - 1))
    formula = Trim$(Mid$(line, at + 2))
    If Not IsPlainName(measureName) Then Err.Raise vbObjectError + 131, , "Invalid measure name: " & measureName
    If Len(formula) = 0 Then Err.Raise vbObjectError + 132, , "Empty formula for " & measureName
    If Left$(formula, 1) = "=" Then Err.Raise vbObjectError + 133, , "Source expressions must not start with an equals sign."
    ParseDefinitionLine = Array(measureName, "=" & formula)
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

Private Function MaskStrings(ByVal formula As String) As String
    Dim i As Long
    Dim ending As Long
    Dim result As String
    i = 1
    Do While i <= Len(formula)
        If Mid$(formula, i, 1) = """" Then
            ending = QuotedEnd(formula, i, """")
            result = result & Space$(ending - i + 1)
            i = ending + 1
        Else
            result = result & Mid$(formula, i, 1)
            i = i + 1
        End If
    Loop
    MaskStrings = result
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
' END PURE CORE

Private Sub SelfTestCore()
    Dim parsed As Variant
    If CanonicalFormula("= SUM ( FactCommercial[gross_sales] )") <> CanonicalFormula("sum('FactCommercial'[gross_sales])") Then
        Err.Raise vbObjectError + 140, , "Internal test failed: harmless formula formatting."
    End If
    If CanonicalFormula("=""a b""") = CanonicalFormula("=""ab""") Then
        Err.Raise vbObjectError + 141, , "Internal test failed: string spaces."
    End If
    If CanonicalFormula("=[a b]") = CanonicalFormula("=[ab]") Then
        Err.Raise vbObjectError + 142, , "Internal test failed: identifier spaces."
    End If
    parsed = ParseDefinitionLine("sample := SUM(FactCommercial[gross_sales])")
    If parsed(0) <> "sample" Or parsed(1) <> "=SUM(FactCommercial[gross_sales])" Then
        Err.Raise vbObjectError + 143, , "Internal test failed: source parser."
    End If
End Sub
