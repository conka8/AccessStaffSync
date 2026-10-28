Attribute VB_Name = "modStaffSync"
Option Compare Database
Option Explicit

' AccessStaffSync - Microsoft Access / VBA proof of concept.
' Pure DAO + late-bound Excel/FSO (no extra VBA references required).
' Intended to be used on an isolated demo .accdb first.

Private Const TIMER_SECONDS As Long = 60
Private Const READ_DELAY_SECONDS As Long = 20
Private Const MAX_EXCEL_ROWS As Long = 10001
Private Const SHEET_COLUMNS As Long = 7
Private Const CLEAR_TOKEN As String = "__CLEAR__"
Private mInTick As Boolean

' Call from the unbound form's On Timer property: =SyncTick()
Public Function SyncTick() As Boolean
    Dim lockNo As Integer
    Dim lockOpen As Boolean
    Dim inbox As String, archive As String, receipts As String
    Dim fso As Object, folder As Object, fil As Object
    Dim fileNames As Collection, fileName As Variant

    On Error GoTo Failed
    If mInTick Then Exit Function
    mInTick = True
    lockNo = FreeFile
    Open CurrentProject.Path & "\staffsync.instance.lock" For Binary Access Read Write Lock Read Write As #lockNo
    lockOpen = True

    inbox = SyncSetting("InboxPath", CurrentProject.Path & "\inbox")
    archive = SyncSetting("ArchivePath", CurrentProject.Path & "\archive")
    receipts = SyncSetting("ReceiptPath", CurrentProject.Path & "\receipts")
    Set fso = CreateObject("Scripting.FileSystemObject")
    EnsureFolder fso, inbox
    EnsureFolder fso, archive
    EnsureFolder fso, receipts

    ' Recover archive/receipt steps interrupted by a crash after DB commit.
    ReconcileCompletedFiles inbox, archive, receipts, fso

    Set fileNames = New Collection
    Set folder = fso.GetFolder(inbox)
    For Each fil In folder.Files
        If LCase$(fso.GetExtensionName(fil.Name)) = "xlsx" Then
            If Left$(fil.Name, 2) <> "~$" Then fileNames.Add CStr(fil.Name)
        End If
    Next fil

    For Each fileName In fileNames
        If Not FileAlreadyComplete(CStr(fileName)) Then
            If FileReady(fso, inbox & "\" & CStr(fileName)) Then
                ProcessWorkbook inbox & "\" & CStr(fileName), CStr(fileName), archive, receipts, fso
            End If
        Else
            ' An already-completed filename must never be silently re-used.
            WriteDiagnostic "Duplicate input filename held for review: " & CStr(fileName)
        End If
    Next fileName

    SyncTick = True
CleanExit:
    On Error Resume Next
    If lockOpen Then Close #lockNo
    mInTick = False
    Exit Function
Failed:
    WriteDiagnostic "SyncTick failed " & CStr(Err.Number) & " - " & Err.Description
    Resume CleanExit
End Function

Private Function FileReady(ByVal fso As Object, ByVal path As String) As Boolean
    On Error GoTo NotReady
    Dim delaySec As Long
    delaySec = CLng(Val(SyncSetting("ReadyDelaySeconds", CStr(READ_DELAY_SECONDS))))
    If delaySec < 0 Then delaySec = READ_DELAY_SECONDS
    If fso.FileExists(path) Then
        If fso.GetFile(path).Size > 0 Then
            FileReady = (DateDiff("s", fso.GetFile(path).DateLastModified, Now()) >= delaySec)
        End If
    End If
NotReady:
End Function

Private Sub ProcessWorkbook(ByVal path As String, ByVal fileKey As String, _
                            ByVal archive As String, ByVal receipts As String, ByVal fso As Object)
    Dim excelApp As Object, wb As Object, wsSheet As Object
    Dim lastRow As Long, lastInColumn As Long, items As Variant, i As Long, j As Long
    Dim eventId As String, staffId As String, isoAt As String
    Dim stat As String, avail As String, addr As String, consent As String
    Dim eventAt As Date, why As String, outcome As String
    Dim seen As Object, details As String
    Dim applied As Long, stale As Long, duplicates As Long, rejected As Long, locked As Long
    Dim total As Long, result As String, attemptNumber As Long
    On Error GoTo ImportFailure

    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    excelApp.AskToUpdateLinks = False
    excelApp.AutomationSecurity = 3 ' msoAutomationSecurityForceDisable
    Set wb = excelApp.Workbooks.Open(path, 0, True, , , , True)
    Set wsSheet = wb.Worksheets(1)

    For j = 1 To SHEET_COLUMNS
        If StrComp(CellText(wsSheet.Cells(1, j).Value2), ExpectedHeader(j), vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 201, , "Invalid Excel header at column " & j & _
                " (expected '" & ExpectedHeader(j) & "')"
        End If
    Next j
    ' Detect rows by all seven columns, not only EventID, so an incomplete
    ' final row cannot be silently ignored. Avoid UsedRange (formatting can
    ' expand it to the entire sheet).
    lastRow = 1
    For j = 1 To SHEET_COLUMNS
        lastInColumn = wsSheet.Cells(wsSheet.Rows.Count, j).End(-4162).Row ' xlUp
        If lastInColumn > lastRow Then lastRow = lastInColumn
    Next j
    If lastRow > MAX_EXCEL_ROWS Then Err.Raise vbObjectError + 202, , "Excel file exceeds 10,000 update rows"

    If lastRow >= 2 Then
        items = wsSheet.Range("A1:G" & lastRow).Value2
    End If
    wb.Close False
    Set wb = Nothing
    excelApp.Quit
    Set excelApp = Nothing

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1 ' TextCompare; EventIDs are case-insensitive
    total = IIf(lastRow > 1, lastRow - 1, 0)

    For i = 2 To lastRow
        why = vbNullString
        If HasErrorCell(items, i) Then
            why = "Excel error value is not permitted"
        Else
            eventId = CellText(items(i, 1))
            staffId = CellText(items(i, 2))
            isoAt = CellText(items(i, 3))
            stat = CellText(items(i, 4))
            avail = CellText(items(i, 5))
            addr = CellText(items(i, 6))
            consent = CellText(items(i, 7))
            why = ValidateRow(eventId, staffId, isoAt, stat, avail, addr, consent, eventAt)
        End If

        If Len(why) > 0 Then
            rejected = rejected + 1
            details = details & "Row " & i & ": REJECTED - " & why & vbCrLf
        ElseIf seen.Exists(eventId) Then
            duplicates = duplicates + 1
            details = details & "Row " & i & ": DUPLICATE - repeated EventID in workbook" & vbCrLf
        Else
            seen.Add eventId, True
            outcome = ApplyEvent(eventId, staffId, eventAt, stat, avail, addr, consent, fileKey)
            Select Case outcome
                Case "APPLIED": applied = applied + 1
                Case "STALE": stale = stale + 1
                Case "DUPLICATE": duplicates = duplicates + 1
                Case "LOCKED": locked = locked + 1
                Case Else
                    rejected = rejected + 1
                    details = details & "Row " & i & ": REJECTED - " & outcome & vbCrLf
            End Select
        End If
    Next i

    ' Do not finalize a batch with a transient lock. Successfully committed
    ' rows are already in tblImportEvents, so the next pass is idempotent.
    If locked > 0 Then
        WriteDiagnostic fileKey & ": " & locked & " locked records; retry on the next timer tick"
        Exit Sub
    End If

    result = "SUCCESS"
    If rejected > 0 Then result = "PARTIAL"
    SaveCompletedFile fileKey, total, applied, stale, duplicates, rejected, result, details
    ReconcileOneFile fileKey, SyncSetting("InboxPath", CurrentProject.Path & "\inbox"), archive, receipts, fso
    Exit Sub

ImportFailure:
    WriteDiagnostic fileKey & " - Excel/read error " & CStr(Err.Number) & ": " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not excelApp Is Nothing Then excelApp.Quit
End Sub

Private Function HasErrorCell(ByRef a As Variant, ByVal r As Long) As Boolean
    Dim c As Long
    For c = 1 To SHEET_COLUMNS
        If IsError(a(r, c)) Then HasErrorCell = True: Exit Function
    Next c
End Function

Private Function ExpectedHeader(ByVal colNo As Long) As String
    Select Case colNo
        Case 1: ExpectedHeader = "EventID"
        Case 2: ExpectedHeader = "StaffID"
        Case 3: ExpectedHeader = "EventTimeUTC"
        Case 4: ExpectedHeader = "JobOfferStatus"
        Case 5: ExpectedHeader = "Availability"
        Case 6: ExpectedHeader = "Address"
        Case 7: ExpectedHeader = "MessagingConsent"
    End Select
End Function

Private Function CellText(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function
    CellText = Trim$(CStr(v))
End Function

Private Function ValidateRow(ByVal eventId As String, ByVal staffId As String, _
                             ByVal isoAt As String, ByVal stat As String, ByVal avail As String, _
                             ByVal addr As String, ByVal consent As String, ByRef eventAt As Date) As String
    If Len(eventId) = 0 Or Len(eventId) > 64 Then
        ValidateRow = "EventID must have 1 to 64 characters": Exit Function
    End If
    If Len(staffId) = 0 Or Len(staffId) > 32 Then
        ValidateRow = "StaffID must have 1 to 32 characters": Exit Function
    End If
    If Not ParseUtcTimestamp(isoAt, eventAt) Then
        ValidateRow = "EventTimeUTC must be YYYY-MM-DDTHH:MM:SSZ": Exit Function
    End If
    If Len(stat) > 40 Or Len(avail) > 80 Or Len(addr) > 255 Then
        ValidateRow = "Field value exceeds maximum length": Exit Function
    End If
    If Len(consent) > 0 Then
        Select Case NormalizedConsent(consent)
            Case "YES", "NO", "TRUE", "FALSE", "1", "0", "SI"
            Case Else
                ValidateRow = "MessagingConsent must be YES/NO/TRUE/FALSE/1/0/SI": Exit Function
        End Select
    End If
    If Len(stat) + Len(avail) + Len(addr) + Len(consent) = 0 Then
        ValidateRow = "No fields to update": Exit Function
    End If
    If UCase$(consent) = CLEAR_TOKEN Then
        ValidateRow = "Consent cannot be cleared": Exit Function
    End If
End Function

' ISO 8601 UTC seconds. Date values in tblStaff *At columns are stored as UTC wall times.
Private Function ParseUtcTimestamp(ByVal s As String, ByRef value As Date) As Boolean
    On Error GoTo Invalid
    If Len(s) <> 20 Then Exit Function
    If Mid$(s, 5, 1) <> "-" Or Mid$(s, 8, 1) <> "-" Or Mid$(s, 11, 1) <> "T" _
        Or Mid$(s, 14, 1) <> ":" Or Mid$(s, 17, 1) <> ":" Or Right$(s, 1) <> "Z" Then Exit Function
    If Not (IsNumeric(Mid$(s, 1, 4)) And IsNumeric(Mid$(s, 6, 2)) And _
        IsNumeric(Mid$(s, 9, 2)) And IsNumeric(Mid$(s, 12, 2)) And _
        IsNumeric(Mid$(s, 15, 2)) And IsNumeric(Mid$(s, 18, 2))) Then Exit Function
    value = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2))) + _
            TimeSerial(CInt(Mid$(s, 12, 2)), CInt(Mid$(s, 15, 2)), CInt(Mid$(s, 18, 2)))
    ParseUtcTimestamp = (Format$(value, "yyyy-mm-dd") = Left$(s, 10) And _
                         Format$(value, "hh:nn:ss") = Mid$(s, 12, 8))
Invalid:
End Function

Private Function IsNewEnough(ByVal oldTime As Variant, ByVal newTime As Date) As Boolean
    If IsNull(oldTime) Then
        IsNewEnough = True
    Else
        IsNewEnough = (newTime >= CDate(oldTime))
    End If
End Function

' Per-event DAO transaction: update the staff row and the event ledger atomically.
Private Function ApplyEvent(ByVal eventId As String, ByVal staffId As String, _
                            ByVal eventAt As Date, ByVal stat As String, ByVal avail As String, _
                            ByVal addr As String, ByVal consent As String, ByVal sourceFile As String) As String
    Dim db As DAO.Database, work As DAO.Workspace
    Dim q As DAO.QueryDef, person As DAO.Recordset, ev As DAO.Recordset
    Dim priorOutcome As String, priorSource As String
    Dim changed As Boolean, withinTrans As Boolean
    Dim faultCode As Long, faultText As String
    On Error GoTo Failure
    Set db = CurrentDb()
    priorOutcome = ExistingEvent(db, eventId, priorSource)
    If Len(priorOutcome) > 0 Then
        If StrComp(priorSource, sourceFile, vbTextCompare) = 0 Then
            ApplyEvent = priorOutcome
        Else
            ApplyEvent = "DUPLICATE"
        End If
        Exit Function
    End If

    Set work = DBEngine.Workspaces(0)
    work.BeginTrans
    withinTrans = True
    Set q = db.CreateQueryDef("", _
        "PARAMETERS [pStaff] TEXT(32); SELECT * FROM [tblStaff] WHERE [StaffID] = [pStaff]")
    q.Parameters("pStaff") = staffId
    Set person = q.OpenRecordset(dbOpenDynaset)
    If person.EOF Then
        ApplyEvent = "UNKNOWN STAFF: " & staffId
        work.Rollback
        withinTrans = False
        Exit Function
    End If

    changed = False
    person.Edit
    If Len(stat) > 0 And IsNewEnough(person.Fields("JobOfferStatusAt").Value, eventAt) Then
        person.Fields("JobOfferStatus").Value = NullablePatch(stat)
        person.Fields("JobOfferStatusAt").Value = eventAt
        changed = True
    End If
    If Len(avail) > 0 And IsNewEnough(person.Fields("AvailabilityAt").Value, eventAt) Then
        person.Fields("Availability").Value = NullablePatch(avail)
        person.Fields("AvailabilityAt").Value = eventAt
        changed = True
    End If
    If Len(addr) > 0 And IsNewEnough(person.Fields("AddressAt").Value, eventAt) Then
        person.Fields("Address").Value = NullablePatch(addr)
        person.Fields("AddressAt").Value = eventAt
        changed = True
    End If
    If Len(consent) > 0 And IsNewEnough(person.Fields("MessagingConsentAt").Value, eventAt) Then
        person.Fields("MessagingConsent").Value = ConsentIsTrue(consent)
        person.Fields("MessagingConsentAt").Value = eventAt
        changed = True
    End If
    If changed Then
        person.Update
    Else
        person.CancelUpdate
    End If

    Set ev = db.OpenRecordset("tblImportEvents", dbOpenDynaset)
    ev.AddNew
    ev.Fields("EventID").Value = eventId
    ev.Fields("StaffID").Value = staffId
    ev.Fields("EventTimeUTC").Value = eventAt
    ev.Fields("SourceFile").Value = sourceFile
    ev.Fields("Outcome").Value = IIf(changed, "APPLIED", "STALE")
    ev.Fields("ProcessedAtLocal").Value = Now()
    ev.Update
    work.CommitTrans
    withinTrans = False
    ApplyEvent = IIf(changed, "APPLIED", "STALE")
Done:
    On Error Resume Next
    If Not person Is Nothing Then person.Close
    If Not ev Is Nothing Then ev.Close
    Exit Function
Failure:
    faultCode = Err.Number
    faultText = Err.Description
    On Error Resume Next
    If withinTrans Then work.Rollback
    Select Case faultCode
        Case 3022
            ApplyEvent = "DUPLICATE" ' Unique EventID collision, concurrent import
        Case 3188, 3197, 3218, 3260, 3050, 3045, 3263
            ApplyEvent = "LOCKED"
        Case Else
            WriteDiagnostic "ApplyEvent " & eventId & " error " & faultCode & ": " & faultText
            ApplyEvent = "LOCKED" ' Retry unexpected DB errors; never discard update
    End Select
    Resume Done
End Function

Private Function NullablePatch(ByVal textValue As String) As Variant
    If StrComp(textValue, CLEAR_TOKEN, vbTextCompare) = 0 Then
        NullablePatch = Null
    Else
        NullablePatch = textValue
    End If
End Function

Private Function NormalizedConsent(ByVal v As String) As String
    ' Accept both SI and SÍ without embedding codepage-specific VBA literals.
    NormalizedConsent = Replace(UCase$(Trim$(v)), ChrW$(205), "I")
End Function

Private Function ConsentIsTrue(ByVal v As String) As Boolean
    Select Case NormalizedConsent(v)
        Case "YES", "TRUE", "1", "SI": ConsentIsTrue = True
    End Select
End Function

Private Function ExistingEvent(ByVal db As DAO.Database, ByVal eventId As String, _
                               ByRef fromFile As String) As String
    Dim q As DAO.QueryDef, rs As DAO.Recordset
    Set q = db.CreateQueryDef("", _
        "PARAMETERS [pId] TEXT(64); SELECT [Outcome], [SourceFile] FROM [tblImportEvents] WHERE [EventID]=[pId]")
    q.Parameters("pId") = eventId
    Set rs = q.OpenRecordset(dbOpenSnapshot)
    If Not rs.EOF Then
        ExistingEvent = Nz(rs.Fields("Outcome").Value, "")
        fromFile = Nz(rs.Fields("SourceFile").Value, "")
    End If
    rs.Close
End Function

Private Sub SaveCompletedFile(ByVal fileKey As String, ByVal total As Long, _
                              ByVal applied As Long, ByVal stale As Long, _
                              ByVal duplicates As Long, ByVal rejected As Long, _
                              ByVal result As String, ByVal details As String)
    Dim db As DAO.Database, rs As DAO.Recordset
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("tblImportFiles", dbOpenDynaset)
    rs.AddNew
    rs.Fields("FileKey").Value = fileKey
    rs.Fields("CompletedAtLocal").Value = Now()
    rs.Fields("TotalRows").Value = total
    rs.Fields("AppliedRows").Value = applied
    rs.Fields("StaleRows").Value = stale
    rs.Fields("DuplicateRows").Value = duplicates
    rs.Fields("RejectedRows").Value = rejected
    rs.Fields("Result").Value = result
    rs.Fields("DetailText").Value = details
    rs.Update
    rs.Close
End Sub

Private Function FileAlreadyComplete(ByVal fileKey As String) As Boolean
    Dim db As DAO.Database, q As DAO.QueryDef, rs As DAO.Recordset
    Set db = CurrentDb()
    Set q = db.CreateQueryDef("", _
        "PARAMETERS [pFile] TEXT(255); SELECT [FileKey] FROM [tblImportFiles] WHERE [FileKey]=[pFile]")
    q.Parameters("pFile") = fileKey
    Set rs = q.OpenRecordset(dbOpenSnapshot)
    FileAlreadyComplete = Not rs.EOF
    rs.Close
End Function

Private Sub ReconcileCompletedFiles(ByVal inbox As String, ByVal archive As String, _
                                   ByVal receipts As String, ByVal fso As Object)
    Dim rs As DAO.Recordset
    Set rs = CurrentDb().OpenRecordset("SELECT [FileKey] FROM [tblImportFiles]", dbOpenSnapshot)
    Do Until rs.EOF
        ReconcileOneFile CStr(rs.Fields("FileKey").Value), inbox, archive, receipts, fso
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Sub ReconcileOneFile(ByVal fileKey As String, ByVal inbox As String, _
                              ByVal archive As String, ByVal receipts As String, ByVal fso As Object)
    Dim rs As DAO.Recordset, q As DAO.QueryDef, db As DAO.Database
    Dim dest As String, src As String, receiptPath As String, report As String
    On Error GoTo Failure
    Set db = CurrentDb()
    Set q = db.CreateQueryDef("", _
        "PARAMETERS [pFile] TEXT(255); SELECT * FROM [tblImportFiles] WHERE [FileKey]=[pFile]")
    q.Parameters("pFile") = fileKey
    Set rs = q.OpenRecordset(dbOpenSnapshot)
    If rs.EOF Then GoTo Done
    src = inbox & "\" & fileKey
    dest = archive & "\" & fileKey
    receiptPath = receipts & "\" & fileKey & ".receipt.txt"
    ' Do not overwrite a previously archived file. This detects name reuse.
    If fso.FileExists(src) Then
        If fso.FileExists(dest) Then
            WriteDiagnostic "Archived file name was reused: " & fileKey & "; manual review required"
            GoTo Done
        End If
        fso.MoveFile src, dest
    End If
    If Not fso.FileExists(dest) Then GoTo Done
    If Not fso.FileExists(receiptPath) Then
        report = "ACCESS STAFF SYNC - IMPORT RECEIPT" & vbCrLf & _
            "File: " & fileKey & vbCrLf & _
            "Result: " & rs.Fields("Result").Value & vbCrLf & _
            "CompletedAtLocal: " & Format$(rs.Fields("CompletedAtLocal").Value, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
            "TotalRows: " & rs.Fields("TotalRows").Value & vbCrLf & _
            "AppliedRows: " & rs.Fields("AppliedRows").Value & vbCrLf & _
            "StaleRows: " & rs.Fields("StaleRows").Value & vbCrLf & _
            "DuplicateRows: " & rs.Fields("DuplicateRows").Value & vbCrLf & _
            "RejectedRows: " & rs.Fields("RejectedRows").Value & vbCrLf & vbCrLf & _
            Nz(rs.Fields("DetailText").Value, "")
        WriteReceiptAtomic fso, receiptPath, report
    End If
Done:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Sub
Failure:
    WriteDiagnostic "Receipt/archive retry required for " & fileKey & " - " & Err.Description
    Resume Done
End Sub

Private Sub WriteReceiptAtomic(ByVal fso As Object, ByVal dest As String, ByVal contents As String)
    Dim tmp As String, stream As Object
    tmp = dest & ".tmp"
    Set stream = fso.CreateTextFile(tmp, True, True) ' Unicode receipt incl. Spanish text
    stream.Write contents
    stream.Close
    If fso.FileExists(dest) Then
        fso.DeleteFile tmp, True
    Else
        fso.MoveFile tmp, dest
    End If
End Sub

Private Sub EnsureFolder(ByVal fso As Object, ByVal path As String)
    If Not fso.FolderExists(path) Then fso.CreateFolder path
End Sub

Public Function SyncSetting(ByVal name As String, ByVal defaultText As String) As String
    On Error GoTo UseDefault
    Dim q As DAO.QueryDef, rs As DAO.Recordset, db As DAO.Database
    Set db = CurrentDb()
    Set q = db.CreateQueryDef("", _
        "PARAMETERS [pName] TEXT(50); SELECT [SettingValue] FROM [tblSyncConfig] WHERE [SettingName]=[pName]")
    q.Parameters("pName") = name
    Set rs = q.OpenRecordset(dbOpenSnapshot)
    If Not rs.EOF Then SyncSetting = Nz(rs.Fields(0).Value, defaultText)
    rs.Close
    If Len(SyncSetting) = 0 Then SyncSetting = defaultText
    Exit Function
UseDefault:
    SyncSetting = defaultText
End Function

Public Sub WriteDiagnostic(ByVal message As String)
    On Error Resume Next
    Dim fso As Object, log As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set log = fso.OpenTextFile(CurrentProject.Path & "\staffsync.log", 8, True, -1)
    log.WriteLine Format$(Now(), "yyyy-mm-dd hh:nn:ss") & " | " & message
    log.Close
End Sub

' Schema creation for a DEMO database only. Never run over client's original tables.
Public Sub InitializeDemo()
    Dim db As DAO.Database
    Set db = CurrentDb()
    If Not HasTable("tblStaff") Then
        db.Execute "CREATE TABLE [tblStaff] (" & _
            "[StaffID] TEXT(32) CONSTRAINT [pkStaff] PRIMARY KEY, " & _
            "[JobOfferStatus] TEXT(40), [JobOfferStatusAt] DATETIME, " & _
            "[Availability] TEXT(80), [AvailabilityAt] DATETIME, " & _
            "[Address] TEXT(255), [AddressAt] DATETIME, " & _
            "[MessagingConsent] YESNO, [MessagingConsentAt] DATETIME)", dbFailOnError
    End If
    If Not HasTable("tblImportEvents") Then
        db.Execute "CREATE TABLE [tblImportEvents] (" & _
            "[EventID] TEXT(64) CONSTRAINT [pkEvents] PRIMARY KEY, " & _
            "[StaffID] TEXT(32), [EventTimeUTC] DATETIME, " & _
            "[SourceFile] TEXT(255), [Outcome] TEXT(20), [ProcessedAtLocal] DATETIME)", dbFailOnError
    End If
    If Not HasTable("tblImportFiles") Then
        db.Execute "CREATE TABLE [tblImportFiles] (" & _
            "[FileKey] TEXT(255) CONSTRAINT [pkFiles] PRIMARY KEY, " & _
            "[CompletedAtLocal] DATETIME, [TotalRows] LONG, " & _
            "[AppliedRows] LONG, [StaleRows] LONG, " & _
            "[DuplicateRows] LONG, [RejectedRows] LONG, " & _
            "[Result] TEXT(20), [DetailText] MEMO)", dbFailOnError
    End If
    If Not HasTable("tblSyncConfig") Then
        db.Execute "CREATE TABLE [tblSyncConfig] (" & _
            "[SettingName] TEXT(50) CONSTRAINT [pkConfig] PRIMARY KEY, " & _
            "[SettingValue] MEMO)", dbFailOnError
    End If
    PutSetting "InboxPath", CurrentProject.Path & "\inbox"
    PutSetting "ArchivePath", CurrentProject.Path & "\archive"
    PutSetting "ReceiptPath", CurrentProject.Path & "\receipts"
    PutSetting "ReadyDelaySeconds", "20"
End Sub

Private Function HasTable(ByVal name As String) As Boolean
    On Error GoTo Missing
    Dim td As DAO.TableDef
    Set td = CurrentDb().TableDefs(name)
    HasTable = True
Missing:
End Function

Public Sub PutSetting(ByVal key As String, ByVal newValue As String)
    Dim db As DAO.Database, q As DAO.QueryDef, rs As DAO.Recordset
    Set db = CurrentDb()
    Set q = db.CreateQueryDef("", _
        "PARAMETERS [pKey] TEXT(50); SELECT * FROM [tblSyncConfig] WHERE [SettingName]=[pKey]")
    q.Parameters("pKey") = key
    Set rs = q.OpenRecordset(dbOpenDynaset)
    If rs.EOF Then
        rs.AddNew
        rs.Fields("SettingName").Value = key
    Else
        rs.Edit
    End If
    rs.Fields("SettingValue").Value = newValue
    rs.Update
    rs.Close
End Sub

Public Sub SeedDemoStaff(Optional ByVal howMany As Long = 4000)
    Dim db As DAO.Database, rs As DAO.Recordset, work As DAO.Workspace
    Dim i As Long, inTrans As Boolean
    On Error GoTo Fail
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("tblStaff", dbOpenTable)
    If Not (rs.BOF And rs.EOF) Then GoTo Clean
    Set work = DBEngine.Workspaces(0)
    work.BeginTrans
    inTrans = True
    For i = 1 To howMany
        rs.AddNew
        rs.Fields("StaffID").Value = Format$(i, "000000")
        rs.Fields("JobOfferStatus").Value = "Pending"
        rs.Fields("Availability").Value = "Unknown"
        rs.Fields("Address").Value = "Example address " & i
        rs.Fields("MessagingConsent").Value = False
        rs.Update
    Next i
    work.CommitTrans
    inTrans = False
Clean:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Sub
Fail:
    If inTrans Then work.Rollback
    MsgBox "SeedDemoStaff failed: " & Err.Description, vbExclamation
    Resume Clean
End Sub

' Optional: call in Access Immediate Window with: ? SyncTick()
Public Function CountImportedEvents() As Long
    CountImportedEvents = DCount("*", "tblImportEvents")
End Function

' Optional startup hook. Hidden forms continue to receive Timer events in Access.
Public Function HideSyncForm() As Boolean
    On Error Resume Next
    Forms("frmSyncTimer").Visible = False
    HideSyncForm = True
End Function
