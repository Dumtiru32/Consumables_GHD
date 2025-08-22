Attribute VB_Name = "ExportConsumables_GENT_module"
Option Compare Database
Option Explicit

'=====================================================================
'  Excel constants (late-binding – works without a reference to Excel)
'=====================================================================
Const xlDatabase        As Long = 1
Const xlTabularRow      As Long = 1
Const xlRowField        As Long = 1
Const xlDataField       As Long = 4
Const xlCenter          As Long = -4108
Const xlUp              As Long = -4162
Const xlToLeft          As Long = -4159
Const xlColumnField     As Long = 2
Const xlOpenXMLWorkbook As Long = 51          ' .xlsx
Const xlLocalSessionChanges As Long = 2      ' no overwrite prompt

'=====================================================================
'  1??  Helper – remove illegal characters from a file name
'=====================================================================
Private Function MakeSafeFileName(s As Variant) As String
    Dim txt As String, illegal As Variant, ch As Variant
    txt = CStr(s)                     ' force to a string – Null becomes ""
    illegal = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In illegal
        txt = Replace(txt, ch, "_")
    Next ch
    txt = Trim(txt)
    Do While Right$(txt, 1) = "." Or Right$(txt, 1) = " "
        txt = Left$(txt, Len(txt) - 1)
    Loop
    MakeSafeFileName = txt
End Function

'=====================================================================
'  2??  Helper – ensure a folder exists (creates it if necessary)
'=====================================================================
Private Function EnsureFolder(ByVal sPath As String) As String
    Dim f As String
    f = Trim$(sPath)
    If Right$(f, 1) <> "\" Then f = f & "\"
    If Dir(f, vbDirectory) = "" Then
        On Error Resume Next
        MkDir f
        On Error GoTo 0
    End If
    EnsureFolder = f
End Function

'=====================================================================
'  3??  Full-path builder (creates month-folder + KAM-sub-folder)
'=====================================================================
Public Function BuildFullPath(ByVal basePath As String, _
                             ByVal monthNum As String, _
                             ByVal yearNum As String, _
                             ByVal kamName As String) As String
    Dim folderPath As String, fName As String, fullPath As String
    
    fName = MakeSafeFileName(kamName)                     ' safe KAM name
    If Len(fName) = 0 Then Exit Function
    
    '--- month folder -------------------------------------------------
    folderPath = EnsureFolder(basePath) & _
                 "Consumables " & monthNum & " " & yearNum & "\"
    folderPath = EnsureFolder(folderPath)                  ' create it
    
    '--- KAM sub-folder ------------------------------------------------
    folderPath = EnsureFolder(folderPath & _
                 "Consumables " & fName & " " & monthNum & " " & yearNum & "\")
    
    '--- final file name -----------------------------------------------
    fullPath = folderPath & "Consumables " & fName & " " & monthNum & " " & Right$(yearNum, 2) & ".xlsx"
    
    If Len(fullPath) > 215 Then Exit Function
    BuildFullPath = fullPath
End Function

'=====================================================================
'  4??  Simple text-file logger (used for both Excel & PDF problems)
'=====================================================================
Public Sub LogProblem(ByVal custID As Long, _
                      ByVal custName As String, _
                      ByVal msg As String)
    Dim logFile As String, txt As String, f As Integer
    logFile = CurrentProject.Path & "\ExportLog.txt"
    txt = Format(Now, "yyyy-mm-dd hh:nn:ss") & _
          " | CustID:" & custID & _
          " | CustName:" & custName & _
          " | " & msg & vbCrLf
    On Error Resume Next
    f = FreeFile
    Open logFile For Append As #f
    Print #f, txt
    Close #f
    On Error GoTo 0
End Sub

'=====================================================================
'  5??  Return an existing worksheet or create a new one
'=====================================================================
Public Function GetOrCreateSheet(ByVal WB As Object, _
                                 ByVal SheetName As String) As Object
    Dim ws As Object
    On Error Resume Next
    Set ws = WB.Worksheets(SheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = WB.Worksheets.Add(After:=WB.Worksheets(WB.Worksheets.Count))
        ws.Name = SheetName
    End If
    Const xlSheetVisible As Long = -1
    If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    Set GetOrCreateSheet = ws
End Function

'=====================================================================
'  6??  Apply simple formatting to a range (Calibri, 11 pt)
'=====================================================================
Public Sub SetProperty(CelRan1 As String, CelRan2 As String, _
                      WB As Object, SheetName As String)
    Dim targetRange As Object, addressStr As String
    addressStr = CelRan1 & ":" & CelRan2
    Set targetRange = WB.Sheets(SheetName).Range(addressStr)
    With targetRange.Font
        .Name = "Calibri"
        .Size = 11
        .Bold = False
    End With
End Sub

'=====================================================================
'  7??  Insert company logo (if a logo exists for the Customer_Entity)
'=====================================================================
Private Sub InsertCompanyLogo(ByVal CompanyKey As Variant, ByVal xlSheet As Object)
    Dim db As DAO.Database, rsLogo As DAO.Recordset, rsAttach As DAO.Recordset2
    Dim tmpPath As String, tmpFile As String, ext As String, baseName As String
    Dim shp As Object, crit As String, msoFalse As String, msoCTrue As String, msoTrue As String
    
    Set db = CurrentDb
    
    If IsNull(CompanyKey) Or Len(Trim$(CStr(CompanyKey))) = 0 Then Exit Sub
    
    '--- 1??  Build the WHERE-clause ---------------------------------
    If IsNumeric(CompanyKey) Then
        crit = "WHERE CompanyID = " & CLng(CompanyKey)
    Else
        crit = "WHERE CompanyName = '" & Replace(CStr(CompanyKey), "'", "''") & "'"
    End If
    
    '--- 2??  Pull the attachment ------------------------------------
    Set rsLogo = db.OpenRecordset("SELECT CompanyLogo FROM Company " & crit, dbOpenDynaset)
    If rsLogo.EOF Then GoTo CleanExit
    
    If rsLogo.Fields("CompanyLogo").Type <> dbAttachment Then GoTo CleanExit
    Set rsAttach = rsLogo.Fields("CompanyLogo").Value
    If rsAttach.EOF Then GoTo CleanExit
    
    '--- 3??  Delete any existing picture named "CompanyLogo" ----------
    Dim shpDel As Object
    For Each shpDel In xlSheet.Shapes
        If shpDel.Name Like "CompanyLogo*" Then shpDel.Delete
    Next shpDel
    
    '--- 4??  Write the attachment to a temp file --------------------
    tmpPath = Environ$("TEMP") & "\"
    If Len(Dir$(tmpPath, vbDirectory)) = 0 Then tmpPath = CurrentProject.Path & "\"
    
    baseName = rsAttach.Fields("FileName").Value
    If Len(baseName) > 0 And InStrRev(baseName, ".") > 0 Then
        ext = Mid$(baseName, InStrRev(baseName, "."))
    Else
        ext = ".png"
    End If
    
    tmpFile = tmpPath & "CompanyLogo_" & Format$(Now, "yyyymmdd_hhnnss") & ext
    rsAttach.Fields("FileData").SaveToFile tmpFile
    
    '--- 5??  Insert the picture – anchor it to cell A1 ---------------
    Set shp = xlSheet.Shapes.AddPicture( _
                fileName:=tmpFile, _
                LinkToFile:=msoFalse, _
                SaveWithDocument:=msoCTrue, _
                Left:=xlSheet.Range("A1").Left, _
                Top:=xlSheet.Range("A1").Top, _
                Width:=-1, Height:=-1)                 ' -1 = keep original size
    
    shp.Name = "CompanyLogo_" & Replace(CStr(CompanyKey), " ", "_")
    
    '--- 6??  Optional – resize to fit inside the Antet block ----------
    Const maxW As Double = 150    ' max width in points (˜2 cm)
    Const maxH As Double = 50     ' max height in points (˜1 cm)
    With shp
        .LockAspectRatio = msoTrue
        If .Width > maxW Then .Width = maxW
        If .Height > maxH Then .Height = maxH
        .Placement = xlMoveAndSize
    End With
    
    '--- 7??  Clean-up -------------------------------------------------
    On Error Resume Next
    Kill tmpFile
    On Error GoTo 0
    
CleanExit:
    If Not rsAttach Is Nothing Then rsAttach.Close
    If Not rsLogo Is Nothing Then rsLogo.Close
    Set rsAttach = Nothing: Set rsLogo = Nothing: Set db = Nothing
End Sub

'=====================================================================
'  8??  Parse the Posting-Month field to a true VBA Date
'=====================================================================
Public Function ParsePostingMonthToDate(v As Variant) As Date
    On Error GoTo ErrP
    
    If IsDate(v) Then
        ParsePostingMonthToDate = DateSerial(Year(CDate(v)), month(CDate(v)), 1)
        Exit Function
    End If
    
    Dim s As String, parts() As String
    s = Trim$(CStr(v))
    If s = "" Then Err.Raise vbObjectError + 1, , "Empty Posting Month"
    
    If IsNumeric(s) And Len(s) = 6 Then
        ParsePostingMonthToDate = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 5, 2)), 1)
        Exit Function
    End If
    
    If InStr(s, "-") > 0 Then
        parts = Split(s, "-")
    ElseIf InStr(s, "/") > 0 Then
        parts = Split(s, "/")
    ElseIf InStr(s, ".") > 0 Then
        parts = Split(s, ".")
    End If
    
    If UBound(parts) >= 1 Then
        ParsePostingMonthToDate = DateSerial(CInt(parts(0)), CInt(parts(1)), 1)
        Exit Function
    End If
    
    If IsDate(s) Then
        ParsePostingMonthToDate = DateSerial(Year(CDate(s)), month(CDate(s)), 1)
        Exit Function
    End If
    
ErrP:
    Err.Raise vbObjectError + 2, , "Cannot parse Posting Month: " & CStr(v)
End Function

'=====================================================================
'  9??  Write the “Antet” rows on the *Data* sheet (exactly as in your old code)
'=====================================================================
Private Sub AddAntetRows_Data(ws As Object, _
                              postingMonth As Date, _
                              ledgerList As String, _
                              ledgerList2 As String, _
                              rsKAM As DAO.Recordset)

    Dim dFirst As Date, dLast As Date
    dFirst = DateSerial(Year(postingMonth), month(postingMonth), 1)
    dLast = DateSerial(Year(postingMonth), month(postingMonth) + 1, 0)

    With ws
        .Range("A7").Value = "Ledger"
        .Range("A8").Value = "Posting Date From"
        .Range("A9").Value = "Document Nr."
        .Range("A10").Value = "G/L Account Nr."
        .Range("A11").Value = "Segment1"
        .Range("A12").Value = "Segment2"
        .Range("A13").Value = "Segment3"
        .Range("A14").Value = "Segment4"
        .Range("A15").Value = "Segment5"
        .Range("A16").Value = "Segment6"
        .Range("A17").Value = "Segment7"
        .Range("A18").Value = "Segment8"
        .Range("A19").Value = "Segment9"
        .Range("A20").Value = "Segment10"
        .Range("A21").Value = "Include Adjustment Period"

        .Range("B7").Value = "Primary"
        .Cells(8, 2).Value = dFirst                ' B8
        .Cells(8, 2).NumberFormat = "dd/mm/yyyy"
        .Range("B11").Value = ledgerList           ' e.g. [123,456]
        .Cells(8, 5).Value = dLast                 ' E8
        .Cells(8, 5).NumberFormat = "dd/mm/yyyy"

        .Range("B16").Value = rsKAM!KAM_Name
        .Range("B14").Value = rsKAM!KAM_Segment6      ' you can replace with any field you want to show
        .Range("B21").Value = "Yes"

        .Range("C6:G6").Merge
        .Range("C6:G6").HorizontalAlignment = xlCenter
        .Range("C6:G6").Value = ledgerList2        ' e.g. BE_BU_123,BE_BU_456
        .Range("C6:G6").Font.Name = "Calibri"
        .Range("C6:G6").Font.Size = 11
        .Range("C6:G6").Font.Bold = True

        .Range("D8").Value = "Posting Date To"
    End With
End Sub

'=====================================================================
'  ??  Write the “Antet” rows on the *Overview* sheet
'=====================================================================
Private Sub AddAntetRows_Overview(ws As Object, _
                                 postingMonth As Date, _
                                 ledgerList As String, _
                                 ledgerList2 As String, _
                                 rsKAM As DAO.Recordset)

    Dim dFirst As Date, dLast As Date
    dFirst = DateSerial(Year(postingMonth), month(postingMonth), 1)
    dLast = DateSerial(Year(postingMonth), month(postingMonth) + 1, 0)

    With ws
        .Range("A7").Value = "Posting Date From"
        .Range("A8").Value = "Document Nr."
        .Range("A9").Value = "G/L Account Nr."
        .Range("A10").Value = "Segment1"
        .Range("A11").Value = "Segment2"
        .Range("A12").Value = "Segment3"
        .Range("A13").Value = "Segment4"
        .Range("A14").Value = "Segment5"
        .Range("A15").Value = "Segment6"
        .Range("A16").Value = "Segment7"
        .Range("A17").Value = "Segment8"
        .Range("A18").Value = "Segment9"
        .Range("A19").Value = "Segment10"
        .Range("A20").Value = "Include Adjustment Period"

        .Cells(7, 2).Value = dFirst                 ' B7
        .Cells(7, 2).NumberFormat = "dd/mm/yyyy"
        .Range("B10").Value = ledgerList
        .Cells(7, 5).Value = dLast                  ' E7
        .Cells(7, 5).NumberFormat = "dd/mm/yyyy"

        .Range("B15").Value = rsKAM!KAM_Name
        .Range("B13").Value = rsKAM!KAM_Segment6
        .Range("B20").Value = "Yes"

        .Range("C6:F6").Merge
        .Range("C6:F6").HorizontalAlignment = xlCenter
        .Range("C6:F6").Value = ledgerList2
        .Range("C6:F6").Font.Name = "Calibri"
        .Range("C6:F6").Font.Size = 11
        .Range("C6:F6").Font.Bold = True

        .Range("D7").Value = "Posting Date To"
        .Range("B1:I1").ColumnWidth = 32.3
    End With
End Sub

'=====================================================================
'  1??1??  MAIN ROUTINE – Export ONE workbook per KAM-segment-6 group
'=====================================================================
Public Sub ExportConsumables_GENT()
    On Error GoTo ErrHandler
    
    Dim db          As DAO.Database
    Dim rsKAM       As DAO.Recordset
    Dim rsData      As DAO.Recordset
    Dim xlApp       As Object            ' Excel.Application (late-bound)
    Dim xlWB        As Object
    Dim xlData      As Object
    Dim xlOverview  As Object
    Dim pc          As Object            ' PivotCache
    Dim pt          As Object            ' PivotTable
    Dim srcRange    As Object
    Dim destRange   As Object
    
    Dim basePath    As String
    Dim monthNum    As String, yearNum As String
    Dim fullPath    As String
    Dim maxPosting  As Date
    Dim SourceDataStr As String
    
    Dim custNames   As Variant, custName As String
    Dim i As Long, j As Long
    Dim lastRow As Long, lastCol As Long
    
    Dim ledgerList As String, ledgerList2 As String
    Dim ledgers As Collection, ledgerVal As String, tempVal As String
    
    Set db = CurrentDb()
    
    '--- 1) read all KAM records -------------------------------------------------
    Set rsKAM = db.OpenRecordset( _
        "SELECT KAM_location, KAM_Segment6, KAM_name FROM KeyAccountManager " & _
        "WHERE KAM_Segment6 Is Not Null", dbOpenSnapshot)
    If rsKAM.EOF Then
        MsgBox "No KAM-segment-6 data found.", vbExclamation
        GoTo CleanUp
    End If
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    
    Do While Not rsKAM.EOF
        '------------------------------------------------------
        '  Build base path (must exist) -------------------------
        '------------------------------------------------------
        basePath = Nz(rsKAM!KAM_location, "")
        If Len(Trim$(basePath)) = 0 Then
            LogProblem 0, "(KAM row)", "KAM_location is empty – row skipped."
            GoTo ContinueKAM
        End If
        
        '------------------------------------------------------
        '  Split the pipe-separated customer list
        '------------------------------------------------------
        custNames = Split(Nz(rsKAM!KAM_Segment6, ""), "|")
        
        '------------------------------------------------------
        '  Build the IN-list for the SELECT statement
        '------------------------------------------------------
        Dim inList As String: inList = ""
        For i = LBound(custNames) To UBound(custNames)
            custName = Trim$(custNames(i))
            If Len(custName) = 0 Then GoTo SkipCust
            inList = inList & ",'" & Replace(custName, "'", "''") & "'"
SkipCust:
        Next i
        If Len(inList) = 0 Then
            LogProblem 0, "(KAM row)", "All customer names empty after split."
            GoTo ContinueKAM
        End If
        inList = Mid$(inList, 2)                ' strip leading comma
        
        '------------------------------------------------------
        '  Pull the data for all customers in this KAM group
        '------------------------------------------------------
        Dim sqlData As String
        sqlData = "SELECT * FROM Consumables WHERE [Customer] IN (" & inList & ")"
        Set rsData = db.OpenRecordset(sqlData, dbOpenDynaset)
        If rsData.EOF Then
            LogProblem 0, "(KAM row)", "No consumables rows for this KAM."
            rsData.Close: Set rsData = Nothing
            GoTo ContinueKAM
        End If
        
        '------------------------------------------------------
        '  Determine the latest Posting-Month (used for folder name)
        '------------------------------------------------------
        maxPosting = DateSerial(1900, 1, 1)
        rsData.MoveFirst
        Do Until rsData.EOF
            If Not IsNull(rsData![Posting Month]) Then
                Dim dTmp As Date
                dTmp = ParsePostingMonthToDate(rsData![Posting Month])
                If dTmp > maxPosting Then maxPosting = dTmp
            End If
            rsData.MoveNext
        Loop
        
        If Year(maxPosting) = 1900 Then
            LogProblem 0, "(KAM row)", "Unable to determine a Posting-Month."
            rsData.Close: Set rsData = Nothing
            GoTo ContinueKAM
        End If
        
        monthNum = Format(maxPosting, "mm")
        yearNum = Format(maxPosting, "yyyy")
        
        '------------------------------------------------------
        '  Build the full path (folder + file name)
        '------------------------------------------------------
        fullPath = BuildFullPath(basePath, monthNum, yearNum, rsKAM!KAM_Name)
        If Len(fullPath) = 0 Then
            LogProblem 0, "(KAM row)", "BuildFullPath failed – skipped."
            rsData.Close: Set rsData = Nothing
            GoTo ContinueKAM
        End If
        
        '------------------------------------------------------
        '  CREATE WORKBOOK & WRITE RAW DATA
        '------------------------------------------------------
        Set xlWB = xlApp.Workbooks.Add
        Set xlData = xlWB.Worksheets(1)
        xlData.Name = "Data"
        xlData.Cells.Clear
        
        '--- header row (row 25) ---------------------------------
        For j = 0 To rsData.Fields.Count - 1
            xlData.Cells(25, j + 1).Value = rsData.Fields(j).Name
        Next j
        
        '--- copy the raw data (starts at row 26) -----------------
        rsData.MoveFirst
        xlData.Range("A26").CopyFromRecordset rsData
        
        lastRow = xlData.Cells(xlData.Rows.Count, 1).End(xlUp).row
        lastCol = xlData.Cells(25, xlData.Columns.Count).End(xlToLeft).Column
        
        '--- build the ledger-list strings -------------------------
        Set ledgers = New Collection
        rsData.MoveFirst
        Do Until rsData.EOF
            tempVal = Trim(Nz(rsData!Ledger, ""))
            If tempVal <> "" Then
                ledgerVal = Right(tempVal, 3)
                On Error Resume Next
                ledgers.Add ledgerVal, ledgerVal
                On Error GoTo 0
            End If
            rsData.MoveNext
        Loop
        
        ledgerList = "["
        ledgerList2 = ""
        For i = 1 To ledgers.Count
            ledgerList = ledgerList & ledgers(i)
            If i < ledgers.Count Then ledgerList = ledgerList & ","
            ledgerList2 = ledgerList2 & "BE_BU_" & Right$(ledgers(i), 3)
            If i < ledgers.Count Then ledgerList2 = ledgerList2 & ","
        Next i
        ledgerList = ledgerList & "]"
        
        '--- **INSERT ANTET ROWS on the Data sheet** -------------
        AddAntetRows_Data xlData, maxPosting, ledgerList, ledgerList2, rsKAM
        
        Call SetProperty("A7", "F21", xlWB, "Data")
        Call InsertCompanyLogo(rsKAM!KAM_Name, xlData)   ' <-- logo inside Antet area
        
        With xlData
            .Activate
            .Range("A25:Y25").AutoFilter
            .Range("A25:Y25").Interior.Color = RGB(201, 201, 201)
        End With
        
        '------------------------------------------------------
        '  Overview sheet (pivot destination)
        '------------------------------------------------------
        Set xlOverview = GetOrCreateSheet(xlWB, "Overview")
        
        Set srcRange = xlData.Range(xlData.Cells(25, 1), xlData.Cells(lastRow, lastCol))
        SourceDataStr = "'" & xlData.Name & "'!" & srcRange.Address(False, False)
        
        Set pc = xlWB.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SourceDataStr)
        Set destRange = xlOverview.Range("A26")
        Set pt = pc.CreatePivotTable(TableDestination:=destRange, TableName:="PivotTable1")
        pt.RowAxisLayout xlTabularRow
        
        '--- pivot field configuration (your original block) -------------
        On Error Resume Next
        With pt
            .PivotFields("Customer").Orientation = xlRowField
            .PivotFields("Segment5").Orientation = xlRowField
            .PivotFields("Segment4").Orientation = xlRowField
            .PivotFields("Segment 4 Name").Orientation = xlRowField
            .PivotFields("Document Nr.").Orientation = xlRowField
            .PivotFields("External Document Nr.").Orientation = xlRowField
            .PivotFields("Source Contact Name").Orientation = xlRowField
            .PivotFields("Description").Orientation = xlRowField
            .PivotFields("GL Account Name").Orientation = xlRowField
            .PivotFields("PO Number").Orientation = xlRowField
            .PivotFields("Posting Month").Orientation = xlColumnField
            
            .PivotFields("Amount").Orientation = xlDataField
            .PivotFields("Amount").Function = xlSum
            .PivotFields("Amount").Name = "Sum of Amount"
            .DataFields(1).NumberFormat = "#,##0.00"
            
            Dim pf As Object
            For Each pf In .PivotFields
                If pf.Orientation = xlRowField Or pf.Orientation = xlColumnField Then
                    pf.Subtotals = Array(False, False, False, False, False, False, _
                                         False, False, False, False, False, False)
                End If
            Next pf
            
            .TableStyle2 = "None"
        End With
        On Error GoTo 0
        
        '--- **INSERT ANTET ROWS on the Overview sheet** -----------------
        AddAntetRows_Overview xlOverview, maxPosting, ledgerList, ledgerList2, rsKAM
        
        Call SetProperty("A7", "F20", xlWB, "Overview")
        Call InsertCompanyLogo(rsKAM!KAM_Name, xlOverview)
        
        '------------------------------------------------------
        '  SAVE THE WORKBOOK
        '------------------------------------------------------
        On Error GoTo SaveFailed
        xlWB.SaveAs fileName:=fullPath, _
                    FileFormat:=xlOpenXMLWorkbook, _
                    ConflictResolution:=xlLocalSessionChanges
        On Error GoTo 0
        
        xlWB.Close SaveChanges:=False
        Set xlWB = Nothing
        
CloseWorkbook:
        If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
        
ContinueKAM:
        rsKAM.MoveNext
    Loop
    
CleanUp:
    If Not rsKAM Is Nothing Then rsKAM.Close: Set rsKAM = Nothing
    If Not xlApp Is Nothing Then xlApp.Quit: Set xlApp = Nothing
    Set db = Nothing
    
    MsgBox "Export finished.", vbInformation
    Exit Sub
    
'=====================================================================
'  Error handling for the SaveAs call
'=====================================================================
SaveFailed:
    Dim errMsg As String
    errMsg = "SaveAs failed – " & Err.Number & ": " & Err.Description & _
             ". FullPath='" & fullPath & "'"
    LogProblem 0, "(KAM row)", errMsg
    Err.Clear
    On Error Resume Next
    If Not xlWB Is Nothing Then xlWB.Close SaveChanges:=False
    On Error GoTo 0
    GoTo CloseWorkbook
    
'=====================================================================
'  General error handler
'=====================================================================
ErrHandler:
    MsgBox "Unexpected error " & Err.Number & ": " & Err.Description, vbCritical, "Export failed"
    Resume CleanUp
End Sub

