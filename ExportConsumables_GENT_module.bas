Option Compare Database
Option Explicit

'=====================================================================
'  Excel constants (late-binding friendly)
'=====================================================================
Const xlDatabase       As Long = 1
Const xlTabularRow     As Long = 1
Const xlRowField       As Long = 1
Const xlDataField      As Long = 4
Const xlCenter         As Long = -4108
Const xlUp             As Long = -4162
Const xlToLeft         As Long = -4159
Const xlColumnField    As Long = 2
Const xlOpenXMLWorkbook As Long = 51          ' .xlsx
Const xlLocalSessionChanges As Long = 2      ' no overwrite prompt

'=====================================================================
'  1?? Helper: remove illegal characters from a file name
'=====================================================================
Public Function CleanFileName(s As String) As String
    Dim illegal As Variant, ch As Variant
    illegal = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In illegal
        s = Replace(s, ch, "_")
    Next ch
    s = Trim(s)
    Do While Right$(s, 1) = "." Or Right$(s, 1) = " "
        s = Left$(s, Len(s) - 1)
    Loop
    CleanFileName = s
End Function

'--------------------------------------------------------------
'  Build a safe full path for the Excel workbook (unchanged logic)
'--------------------------------------------------------------
Public Function BuildFullPathSafe(savePath As String, _
                                 monthFolder As String, _
                                 yearVal As String, monthVal As String, _
                                 custName As String) As String
    Dim fName As String, folderPath As String, fullPath As String
    
    fName = CleanFileName(custName)                 ' remove illegal chars
    If Len(fName) = 0 Then Exit Function            ' nothing left after cleaning
    
    folderPath = Trim(savePath)                      ' strip any trailing space
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    monthFolder = Trim(yearVal) & Trim(monthVal)      ' compact yyyymm
    folderPath = folderPath & monthFolder & "\"
    
    If Dir(folderPath, vbDirectory) = "" Then
        On Error Resume Next
        MkDir folderPath
        If Err.Number <> 0 Then
            Err.Clear
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    fName = Left$(fName, 150)
    fullPath = folderPath & yearVal & monthVal & " Consumables " & fName & ".xlsx"
    
    If Len(fullPath) > 255 Then Exit Function
    
    BuildFullPathSafe = fullPath
End Function

'=====================================================================
'  3?? Simple text-file logger (used for both Excel & PDF problems)
'=====================================================================
Public Sub LogProblem(ByVal custID As Long, _
                      ByVal custName As String, _
                      ByVal msg As String)
    Dim logFile As String, txt As String
    logFile = CurrentProject.Path & "\ExportLog.txt"
    
    txt = Format(Now, "yyyy-mm-dd hh:nn:ss") & _
          " | CustID:" & custID & _
          " | CustName:" & custName & _
          " | " & msg & vbCrLf
    
    On Error Resume Next
    Dim f As Integer
    f = FreeFile
    Open logFile For Append As #f
    Print #f, txt
    Close #f
    On Error GoTo 0
End Sub

'=====================================================================
'  4?? Log a PDF-specific problem (just forwards to LogProblem)
'=====================================================================
Private Sub LogPDFProblem(ByVal custID As Long, _
                         ByVal custName As String, _
                         ByVal msg As String)
    LogProblem custID, custName, "PDF-merge: " & msg
End Sub

'=====================================================================
'  5?? GetOrCreateSheet – returns a Worksheet object.
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
'  6?? Set property for a range (font etc.)
'=====================================================================
Public Sub SetProperty(CelRan1 As String, CelRan2 As String, WB As Object, SheetName As String)
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
'  7?? Insert Company Logo (unchanged, only minor formatting)
'=====================================================================
Private Sub InsertCompanyLogo(ByVal CustomerEntity As String, ByVal xlOverview As Object)
    Dim db As DAO.Database
    Dim rsLogo As DAO.Recordset
    Dim rsAttach As DAO.Recordset2
    Dim tmpPath As String, tmpFile As String
    Dim ext As String, baseName As String
    Dim shp As Object
    Dim crit As String
    
    Set db = CurrentDb
    
    If Len(CustomerEntity) > 0 And IsNumeric(CustomerEntity) Then
        crit = "WHERE CompanyID = " & CLng(CustomerEntity)
    Else
        crit = "WHERE CompanyID = '" & Replace(CustomerEntity, "'", "''") & "'"
    End If
    
    Set rsLogo = db.OpenRecordset( _
        "SELECT CompanyLogo FROM Company " & crit, _
        dbOpenDynaset)
    
    If Not (rsLogo Is Nothing) Then
        If Not rsLogo.EOF Then
            If rsLogo.Fields("CompanyLogo").Type = dbAttachment Then
                Set rsAttach = rsLogo.Fields("CompanyLogo").Value
                If Not rsAttach Is Nothing Then
                    If Not rsAttach.EOF Then
                        tmpPath = Environ$("TEMP") & "\"
                        If Len(Dir$(tmpPath, vbDirectory)) = 0 Then tmpPath = CurrentProject.Path & "\"
                        
                        On Error Resume Next
                        baseName = rsAttach.Fields("FileName").Value
                        On Error GoTo 0
                        If Len(baseName) > 0 And InStrRev(baseName, ".") > 0 Then
                            ext = Mid$(baseName, InStrRev(baseName, "."))
                        Else
                            ext = ".png"
                        End If
                        
                        tmpFile = tmpPath & "CompanyLogo_" & Format$(Now, "yyyymmdd_hhnnss") & ext
                        rsAttach.Fields("FileData").SaveToFile tmpFile
                        
                        Set shp = xlOverview.Shapes.AddPicture( _
                                    tmpFile, _
                                    False, _
                                    True, _
                                    xlOverview.Range("A2").Left, _
                                    xlOverview.Range("A2").Top, _
                                    100, _
                                    50)
                        
                        If Not shp Is Nothing Then
                            Dim target As Object, maxW As Double, maxH As Double
                            Set target = xlOverview.Range("A2:A6")
                            maxW = target.Width: maxH = target.Height
                            On Error Resume Next
                            shp.LockAspectRatio = True
                            If shp.Width > maxW Then shp.Width = maxW
                            If shp.Height > maxH Then shp.Height = maxH
                            shp.Left = target.Left + (maxW - shp.Width) / 2
                            shp.Top = target.Top + (maxH - shp.Height) / 2
                            On Error GoTo 0
                        End If
                        
                        On Error Resume Next
                        Kill tmpFile
                        On Error GoTo 0
                    End If
                    rsAttach.Close
                End If
            End If
        End If
        rsLogo.Close
    End If
    
    Set rsAttach = Nothing: Set rsLogo = Nothing: Set db = Nothing
End Sub

'=====================================================================
'  8?? Pivot-order helper – returns an array of the values that appear
'=====================================================================
Private Function GetPivotOrder(ByVal xlOverview As Object) As Variant
    Dim pt As Object, arr() As String, i As Long
    
    On Error GoTo NoPivot
    Set pt = xlOverview.PivotTables("PivotTable1")
    
    With pt.DataBodyRange
        If .Rows.Count = 0 Then GoTo NoPivot
        ReDim arr(1 To .Rows.Count)
        For i = 1 To .Rows.Count
            arr(i) = CStr(.Cells(i, 1).Value)
        Next i
    End With
    
    GetPivotOrder = arr
    Exit Function
    
NoPivot:
    GetPivotOrder = Array()
End Function

'=====================================================================
'  9?? Clean the External Document Nr. so it matches the PDF file name
'=====================================================================
Private Function CleanExternalDocNr(ByVal s As String) As String
    Dim tmp As String, parts() As String
    
    tmp = Trim(s)
    
    If InStr(tmp, " ") > 0 Then
        parts = Split(tmp, " ")
        If IsNumeric(parts(0)) Then tmp = Mid(tmp, Len(parts(0)) + 2)
    End If
    
    tmp = Replace(tmp, "/", "-")
    CleanExternalDocNr = tmp
End Function

'--------------------------------------------------------------
'  Build the UNC folder that holds the source PDF files.
'--------------------------------------------------------------
Private Function BuildPDFFolder(ByVal postingMonth As Date, _
                               ByVal custName As String) As String
    Dim y As String, m As String
    
    y = Format(postingMonth, "yyyy")
    m = Format(postingMonth, "mm")
    
    BuildPDFFolder = "\\itglo.net\public\EMEA\BE-KI\DataShares\Share Boekhouding CGI Kallo\" & _
                     "Consumabiles for split\CLA\2025\" & y & m & "\" & CleanFileName(custName)
End Function

'=====================================================================
'  1??1?? Merge PDFs – **Ghostscript command-line version**
'=====================================================================
Public Function MergePDFs(ByVal pdfFiles As Variant, _
                         ByVal mergedFile As String) As Boolean
    Dim exePath As String          ' full path to the Ghostscript console exe
    Dim cmd As String
    Dim i As Long
    Dim wsh As Object
    Dim rc As Long
    
    ' ---------------------------------------------------------
    ' 1?? Locate the Ghostscript executable
    ' ---------------------------------------------------------
    exePath = "C:\Users\purceld\OneDrive - Katoen Natie\Documents\gs10051w64.exe"
    
    Debug.Print "Looking for Ghostscript exe at: " & exePath
    Debug.Print "File exists? " & (Dir$(exePath, vbNormal) <> "")
    
    If Dir$(exePath, vbNormal) = "" Then
        LogPDFProblem 0, "", "Ghostscript exe not found – cannot merge PDFs."
        Exit Function
    End If
    
    ' ---------------------------------------------------------
    ' 2?? Verify every source PDF exists
    ' ---------------------------------------------------------
    For i = LBound(pdfFiles) To UBound(pdfFiles)
        Debug.Print "Checking source PDF: " & pdfFiles(i)
        If Dir$(pdfFiles(i), vbNormal) = "" Then
            LogPDFProblem 0, "", "Source PDF not found ? " & pdfFiles(i)
            Exit Function
        End If
    Next i
    
    ' ---------------------------------------------------------
    ' 3?? Build the Ghostscript command line
    ' ---------------------------------------------------------
    '   -dBATCH -dNOPAUSE   ? run without interactive prompts
    '   -sDEVICE=pdfwrite   ? output device = PDF
    '   -sOutputFile="out.pdf" ? name of the merged PDF
    cmd = Chr$(34) & exePath & Chr$(34) & " -dBATCH -dNOPAUSE -sDEVICE=pdfwrite"
    cmd = cmd & " -sOutputFile=" & Chr$(34) & mergedFile & Chr$(34)
    
    For i = LBound(pdfFiles) To UBound(pdfFiles)
        cmd = cmd & " " & Chr$(34) & pdfFiles(i) & Chr$(34)
    Next i
    
    Debug.Print "Running Ghostscript command line:"
    Debug.Print cmd
    
    ' ---------------------------------------------------------
    ' 4?? Execute the command (synchronously)
    ' ---------------------------------------------------------
    On Error GoTo ErrHandler
    Set wsh = CreateObject("WScript.Shell")
    rc = wsh.Run(cmd, 0, True)          ' 0 = hidden window, True = wait for finish
    
    Debug.Print "Ghostscript exit code = " & rc
    
    If rc <> 0 Then GoTo ErrHandler
    
    ' ---------------------------------------------------------
    ' 5?? Verify the merged file really exists
    ' ---------------------------------------------------------
    If Dir$(mergedFile, vbNormal) = "" Then
        LogPDFProblem 0, "", "Merged PDF not found after Ghostscript run: " & mergedFile
        GoTo ErrHandler
    End If
    
    Debug.Print "Merge succeeded – file created: " & mergedFile
    MergePDFs = True
    Exit Function
    
ErrHandler:
    Debug.Print "MergePDFs VBA error " & Err.Number & ": " & Err.Description
    MergePDFs = False
End Function

'=====================================================================
'  1??2?? Parse Posting Month into a true Date value (unchanged)
'=====================================================================
Private Function ParsePostingMonthToDate(v As Variant) As Date
    On Error GoTo ErrP
    If IsDate(v) Then
        ParsePostingMonthToDate = DateSerial(Year(CDate(v)), month(CDate(v)), 1)
        Exit Function
    End If
    
    Dim s As String, parts() As String
    s = Trim(CStr(v))
    If s = "" Then Err.Raise vbObjectError + 1, , "Empty Posting Month"
    
    If IsNumeric(s) And Len(s) = 6 Then
        ParsePostingMonthToDate = DateSerial(CInt(Left(s, 4)), CInt(Mid(s, 5, 2)), 1)
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
'  1??3?? Build & run the combined GL-check query (unchanged)
'=====================================================================
Public Sub BuildAndRun_GLCombinationCheckFromTable()
    Const cDefTable   As String = "tblGLCheckDefs"
    Const cTempTable  As String = "Temp_CheckDefs"
    Const cResultQry  As String = "Check_GL_Combined_qry"

    Dim db As DAO.Database
    Dim rsDefs As DAO.Recordset
    Dim sqlBlock As String, insertSQL As String
    Dim colList As String

    colList = "ID, Ledger, [Segment8 (Dim7)] AS Segment8, Customer, Segment5, " & _
              "[GL Account Nr], Segment4, [External Document Nr], Amount"

    Set db = CurrentDb

    ' Clear the temporary table
    db.Execute "DELETE FROM " & cTempTable, dbFailOnError

    ' Open the definition table
    Set rsDefs = db.OpenRecordset( _
                "SELECT ConditionNr, Description, WhereClause " & _
                "FROM " & cDefTable & " ORDER BY ConditionNr;", dbOpenSnapshot)

    If rsDefs.EOF Then
        MsgBox "Definition table '" & cDefTable & "' is empty.", vbExclamation
        GoTo CleanExit
    End If

    ' Loop through each condition and insert matching records into Temp_CheckDefs
    Do While Not rsDefs.EOF
        
            insertSQL = "INSERT INTO " & cTempTable & " (ID, Ledger, Segment8, Customer, Segment5, [GL Account Nr], Segment4, [External Document Nr], Amount, ConditionNr, ConditionDesc) " & _
                "SELECT ID, Ledger, [Segment8 (Dim7)], Customer, Segment5, [GL Account Nr], Segment4, [External Document Nr], Amount, " & _
                rsDefs!ConditionNr & ", '" & Replace(Nz(rsDefs!Description, ""), "'", "''") & "' " & _
                "FROM Consumables WHERE " & rsDefs!whereClause

        On Error Resume Next
        db.Execute insertSQL, dbFailOnError
        On Error GoTo 0
        rsDefs.MoveNext
    Loop

    ' Open the review form
    DoCmd.OpenQuery ("consumablesTempCheckDefs_qry")

CleanExit:
    If Not rsDefs Is Nothing Then rsDefs.Close: Set rsDefs = Nothing
    Set db = Nothing
End Sub

'Split a pipe (|) list into a distinct Collection of trimmed values
'=====================================================================
Private Function SplitPipeList(ByVal s As String) As Collection
    Dim c As New Collection, arr, i As Long, it As String
    arr = Split(Replace(Nz(s, ""), vbCrLf, "|"), "|")
    For i = LBound(arr) To UBound(arr)
        it = Trim$(CStr(arr(i)))
        If it <> "" Then
            On Error Resume Next
            c.Add it, it             ' distinct by key = value
            On Error GoTo 0
        End If
    Next
    Set SplitPipeList = c
End Function

'=====================================================================
' Join a Collection into a string (already distinct by our splitter)
'=====================================================================
Private Function JoinCollectionDistinct(col As Collection, _
                                        Optional sep As String = " | ") As String
    Dim i As Long, s As String
    For i = 1 To col.Count
        If Len(s) > 0 Then s = s & sep
        s = s & col(i)
    Next
    JoinCollectionDistinct = s
End Function

'=====================================================================
' Build full path for **KAM** export and auto-create the month folder:
'   {KAM_location}\Consumables {MM} {YYYY}\Consumables {KAM_Name} {MM YY}.xlsx
'=====================================================================


Public Function BuildFullPathForKAM(kamLocation As String, _
                                    mm As String, yyyy As String, _
                                    kamName As String) As String
    Dim base As String, monthFolder As String
    Dim subFolderName As String, kamFolder As String, fileName As String

    base = Trim$(Nz(kamLocation, ""))
    If Len(base) = 0 Then Exit Function
    If Right$(base, 1) <> "\" Then base = base & "\"

    ' Create month folder
    monthFolder = base & "Consumables " & mm & " " & yyyy & "\"
    On Error Resume Next
    If Dir$(monthFolder, vbDirectory) = "" Then MkDir monthFolder
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    ' Build subfolder and file name
    subFolderName = "Consumables " & CleanFileName(kamName) & " " & mm & " " & yyyy
    kamFolder = monthFolder & subFolderName & "\"
    fileName = subFolderName & ".xlsx"

    ' Create KAM subfolder
    On Error Resume Next
    If Dir$(kamFolder, vbDirectory) = "" Then MkDir kamFolder
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    ' Return full path to file
    BuildFullPathForKAM = kamFolder & fileName

    If Len(BuildFullPathForKAM) > 255 Then BuildFullPathForKAM = ""
End Function





'=====================================================================
'  1??5?? MAIN ROUTINE – Export Excel + Merge PDFs
'=====================================================================
Public Sub ExportConsumables_GENT()
    
On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim rsKAM As DAO.Recordset          ' iterate KAMs (GHD only)
    Dim rsCust As DAO.Recordset          ' per-customer details from [Customer]
    Dim rsData As DAO.Recordset          ' combined data from [Consumables]
    Dim rsPM As DAO.Recordset            ' for Max Posting Month

    Dim xlApp As Object, xlWB As Object, xlData As Object, xlOverview As Object
    Dim pc As Object, pt As Object, destRange As Object, srcRange As Object

    Dim whereOr As String
    Dim sqlData As String
    Dim i As Long, j As Long, k As Long

    Dim lastRow As Long, lastCol As Long
    Dim SourceDataStr As String

    Dim postingMonthRaw As Variant, postingMonthVal As Date
    Dim yearVal As String, monthVal As String

    Dim ledgers As Collection, ledgerVal As String, tempVal As String
    Dim ledgerList As String, ledgerList2 As String

    Dim seg4Union As New Collection       ' union of Seg4 across all customers of the KAM
    Dim seg6Union As String      ' union of Seg6 across all customers of the KAM
    Dim custList As Collection            ' customers belonging to the KAM
    Dim custName As Variant               ' element from custList

    Dim savePath As String, fullPath As String
    Dim kamName As String, kamEntity As String, kamLocation As String

    Dim pivOrder As Variant, pdfList() As String, pdfCount As Long
    Dim mergedPDF As String, extDoc As String

    Set db = CurrentDb()

    ' 1) Only process KAMs where KAM_entity='GHD' (requirement #1)
    Set rsKAM = db.OpenRecordset( _
        "SELECT * FROM KeyAccountManager " & _
        "WHERE KAM_entity='GHD' AND Nz([KAM_Name],'')<>'' AND Nz([KAM_location],'')<>''", _
        dbOpenDynaset)

    If rsKAM.EOF Then
        MsgBox "No KAMs found with KAM_entity='GHD'.", vbInformation
        GoTo CleanExit
    End If

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False

    Do While Not rsKAM.EOF
        kamName = Nz(rsKAM!KAM_Name, "")
        kamEntity = Nz(rsKAM!KAM_entity, "")
        kamLocation = Nz(rsKAM!KAM_location, "")

        ' ---------- Build the WHERE ... clause using KAM_segment6 ----------
        whereOr = ""
        Set custList = SplitPipeList(Nz(rsKAM!KAM_segment6, ""))

        ' Reset union of Segment4 values
        Set seg4Union = New Collection

        For Each custName In custList
            ' Pull the customer's Seg4 list from [Customer]
            Set rsCust = db.OpenRecordset( _
                "SELECT * FROM Customer WHERE [Customer_Name]='" & _
                Replace(CStr(custName), "'", "''") & "'", dbOpenSnapshot)
            If Not rsCust.EOF Then
                Dim seg4s As Collection, seg4In As String, s4 As Variant
                Set seg4s = SplitPipeList(Nz(rsCust!Customer_Seg4, ""))

                seg4In = ""
                For i = 1 To seg4s.Count
                    If seg4s(i) <> "" Then
                        If seg4In <> "" Then seg4In = seg4In & ","
                        seg4In = seg4In & "'" & Replace(seg4s(i), "'", "''") & "'"
                        ' add to union collection
                        On Error Resume Next
                        seg4Union.Add seg4s(i), seg4s(i)
                        On Error GoTo 0
                    End If
                Next i

                If seg4In <> "" Then
                    If whereOr <> "" Then whereOr = whereOr & " OR "
                    whereOr = whereOr & "(" & _
                              "[Customer]='" & Replace(CStr(custName), "'", "''") & "' AND " & _
                              "[Segment4] IN (" & seg4In & "))"
                End If
            End If
            rsCust.Close: Set rsCust = Nothing
        Next custName
        
        
        'Build seg6Union from Customer.Customer_Name
        seg6Union = ""
        For Each custName In custList
            If seg6Union <> "" Then seg6Union = seg6Union & "|"
            seg6Union = seg6Union & custName
        Next custName



        ' If there is nothing to export for this KAM, continue to next
        If Len(whereOr) = 0 Then GoTo NextKAM

        sqlData = "SELECT * FROM Consumables WHERE " & whereOr
        Set rsData = db.OpenRecordset(sqlData, dbOpenDynaset)
        If rsData.EOF Then
            rsData.Close: Set rsData = Nothing
            GoTo NextKAM
        End If

        ' ---------- Determine the "management month" (max Posting Month) ----------
        Set rsPM = db.OpenRecordset( _
            "SELECT Max([Posting Month]) AS MaxPM FROM Consumables WHERE " & whereOr, dbOpenSnapshot)
        postingMonthRaw = Null
        If Not rsPM.EOF Then postingMonthRaw = rsPM!MaxPM
        rsPM.Close: Set rsPM = Nothing

        postingMonthVal = ParsePostingMonthToDate(postingMonthRaw)
        yearVal = Format(postingMonthVal, "yyyy")
        monthVal = Format(postingMonthVal, "mm")

        ' ---------- Collect distinct ledger suffixes (last 3 chars) ----------
        Set ledgers = New Collection
        rsData.MoveFirst
        Do Until rsData.EOF
            tempVal = Trim$(Nz(rsData!Ledger, ""))
            If tempVal <> "" Then
                ledgerVal = Right$(tempVal, 3)
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

        ' ---------- Create workbook & write data ----------
        Set xlWB = xlApp.Workbooks.Add

        On Error Resume Next
        xlWB.Worksheets(1).Name = "Data"
        On Error GoTo 0
        Set xlData = xlWB.Worksheets("Data")

        ' Column headers at row 25
        For j = 0 To rsData.Fields.Count - 1
            xlData.Cells(25, j + 1).Value = rsData.Fields(j).Name
        Next j

        rsData.MoveFirst
        xlData.Range("A26").CopyFromRecordset rsData

        lastRow = xlData.Cells(xlData.Rows.Count, 1).End(-4162).row    ' xlUp
        lastCol = xlData.Cells(25, xlData.Columns.Count).End(-4159).Column ' xlToLeft
        If lastRow < 26 Or lastCol < 1 Then GoTo CloseWorkbookAndNextKAM

        ' ---------- Data sheet header/formatting (preserved) ----------
        With xlData
            .Range("A1:H1").Merge
            .Range("A1:H1").HorizontalAlignment = -4108 ' xlCenter
            .Range("A1").Value = "BE_050 Analytical - DIM 7-4-3-2 per month"
            .Range("A1").Font.Name = "Calibri": .Range("A1").Font.Size = 14: .Range("A1").Font.Bold = True
            .Range("A1").ColumnWidth = 25.33

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
            .Range("B8").Value = DateSerial(Year(postingMonthVal), month(postingMonthVal), 1)
            .Range("B8").NumberFormat = "dd/mm/yyyy"
            .Range("E8").Value = DateSerial(Year(postingMonthVal), month(postingMonthVal) + 1, 0)
            .Range("E8").NumberFormat = "dd/mm/yyyy"

            .Range("B11").Value = ledgerList
            ' Requirement #6: fill B14/B16 with union of Customer_Seg4
            .Range("B14").Value = JoinCollectionDistinct(seg4Union)
            .Range("B16").Value = seg6Union

            .Range("C6:G6").Merge
            .Range("C6:G6").HorizontalAlignment = -4108 ' xlCenter
            .Range("C6:G6").Value = ledgerList2
            .Range("C6:G6").Font.Name = "Calibri": .Range("C6:G6").Font.Size = 11: .Range("C6:G6").Font.Bold = True
            .Range("D8").Value = "Posting Date To"

            .Activate
            .Range("A25:Y25").AutoFilter
            .Range("A25:Y25").Interior.Color = RGB(201, 201, 201)
        End With

        ' Borders around merged C6:G6 header
        Dim m As Integer
        For m = 1 To 4
            With xlData.Range("C6:G6").Borders(m)
                .LineStyle = 1: .Weight = 2
            End With
        Next m

        Call SetProperty("A7", "F21", xlWB, "Data")
        ' Use **KAM entity** for the logo (requirement #1/7)
        Call InsertCompanyLogo(kamEntity, xlData)

        ' ---------- Overview sheet (pivot destination at A26, preserved formatting) ----------
        Set xlOverview = GetOrCreateSheet(xlWB, "Overview")

        Set srcRange = xlData.Range(xlData.Cells(25, 1), xlData.Cells(lastRow, lastCol))
        SourceDataStr = "'" & xlData.Name & "'!" & srcRange.Address(False, False)
        Set pc = xlWB.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SourceDataStr)
        Set destRange = xlOverview.Range("A26")
        Set pt = pc.CreatePivotTable(TableDestination:=destRange, TableName:="PivotTable1")
        pt.RowAxisLayout xlTabularRow

        On Error Resume Next
        With pt
            .PivotFields("Customer").Orientation = xlRowField
            .PivotFields("Segment5").Orientation = xlRowField
            .PivotFields("Segment4").Orientation = xlRowField
            .PivotFields("Segment 4 Name").Orientation = xlRowField
            .PivotFields("Document Nr").Orientation = xlRowField
            .PivotFields("External Document Nr").Orientation = xlRowField
            .PivotFields("Source Contact Name").Orientation = xlRowField
            .PivotFields("Description").Orientation = xlRowField
            .PivotFields("GL Account Name").Orientation = xlRowField
            .PivotFields("PO Number").Orientation = xlRowField
            .PivotFields("Posting Month").Orientation = xlColumnField
            .PivotFields("Amount").Orientation = xlDataField
            .PivotFields("Amount").Function = 4     ' xlSum
            .PivotFields("Amount").Name = "Sum of Amount"
            .DataFields(1).NumberFormat = "#,##0.00"

            .PivotFields("Document Nr").Position = 5
            .PivotFields("External Document Nr").Position = 6
            .PivotFields("Document Nr").Caption = "Document Nr."
            .PivotFields("External Document Nr").Caption = "External Document Nr."

            Dim pvtField As Object
            For Each pvtField In .PivotFields
                If pvtField.Orientation = 1 Or pvtField.Orientation = 2 Then ' row or column
                    pvtField.Subtotals = Array(False, False, False, False, False, False, _
                                              False, False, False, False, False, False)
                End If
            Next pvtField

            Dim arrNames As Variant
            arrNames = Array("Customer", "Segment5", "Segment4", "Segment 4 Name")
            For k = LBound(arrNames) To UBound(arrNames)
                On Error Resume Next
                .PivotFields(arrNames(k)).Subtotals(1) = True
                On Error GoTo 0
            Next k

            .TableStyle2 = "None"
        End With
        On Error GoTo 0

        With xlOverview
            .Range("A1:H1").Merge
            .Range("A1:H1").HorizontalAlignment = -4108 ' xlCenter
            .Range("A1").Value = "BE_050 Analytical - DIM 7-4-3-2 per month"
            .Range("A1").Font.Name = "Calibri": .Range("A1").Font.Size = 14: .Range("A1").Font.Bold = True
            .Range("A1").ColumnWidth = 25.33

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

            .Range("B7").Value = DateSerial(Year(postingMonthVal), month(postingMonthVal), 1)
            .Range("B7").NumberFormat = "dd/mm/yyyy"
            .Range("E7").Value = DateSerial(Year(postingMonthVal), month(postingMonthVal) + 1, 0)
            .Range("E7").NumberFormat = "dd/mm/yyyy"

            .Range("B10").Value = ledgerList

            ' Requirement #6: fill B13 and B15 with union of Customer_Seg4
            .Range("B13").Value = JoinCollectionDistinct(seg4Union)
            .Range("B15").Value = seg6Union

            .Range("B20").Value = "Yes"

            .Range("C6:F6").Merge
            .Range("C6:F6").HorizontalAlignment = -4108 ' xlCenter
            .Range("C6:F6").Value = ledgerList2
            .Range("C6:F6").Font.Name = "Calibri": .Range("C6:F6").Font.Size = 11: .Range("C6:F6").Font.Bold = True
            .Range("D7").Value = "Posting Date To"
            .Range("B1:I1").ColumnWidth = 32.3
        End With

        ' Borders around merged C6:F6 header
        Dim n As Integer
        For n = 1 To 4
            With xlOverview.Range("C6:F6").Borders(n)
                .LineStyle = 1: .Weight = 2
            End With
        Next n

        Call SetProperty("A7", "F20", xlWB, "Overview")
        Call InsertCompanyLogo(kamEntity, xlOverview)   ' logo by KAM entity

        ' ---------- SAVE to KAM location / month folder (requirement #2 & #3) ----------
        fullPath = BuildFullPathForKAM(kamLocation, monthVal, yearVal, kamName)
        If Len(fullPath) = 0 Then
            LogProblem 0, kamName, "BuildFullPathForKAM failed – missing folder or path too long."
            GoTo CloseWorkbookAndNextKAM
        End If

        ' (second guard) ensure parent exists
        Dim targetFolder As String
        targetFolder = Left$(fullPath, InStrRev(fullPath, "\") - 1)
        If Dir(targetFolder, vbDirectory) = "" Then
            On Error Resume Next
            MkDir targetFolder
            If Err.Number <> 0 Then
                LogProblem 0, kamName, "Could not create folder '" & targetFolder & "' – " & Err.Description
                Err.Clear
                GoTo CloseWorkbookAndNextKAM
            End If
            On Error GoTo 0
        End If

        On Error GoTo SaveFailed
        xlWB.SaveAs fileName:=fullPath, FileFormat:=51, ConflictResolution:=2 ' xlOpenXMLWorkbook, xlLocalSessionChanges
        On Error GoTo 0

        DoEvents
        xlWB.Close SaveChanges:=False

        ' ---------- PDF MERGE (kept logic; adapted to multi-customer) ----------
        ' 1) Read pivot order (only displayed rows)
        pivOrder = GetPivotOrder(xlOverview)
        If Not IsArray(pivOrder) Then GoTo SkipMerge
        If UBound(pivOrder) < LBound(pivOrder) Then GoTo SkipMerge

        ' 2) For each External Document Nr., search PDF in ANY customer folder for this KAM
        pdfCount = 0
        ReDim pdfList(0)
        For i = LBound(pivOrder) To UBound(pivOrder)
            extDoc = CleanExternalDocNr(CStr(pivOrder(i)))
            If extDoc <> "" Then
                Dim candidate As String, foundOne As Boolean
                foundOne = False
                For Each custName In custList
                    candidate = BuildPDFFolder(postingMonthVal, CStr(custName)) & "\" & extDoc & ".pdf"
                    If Dir$(candidate, vbNormal) <> "" Then
                        pdfCount = pdfCount + 1
                        ReDim Preserve pdfList(1 To pdfCount)
                        pdfList(pdfCount) = candidate
                        foundOne = True
                        Exit For
                    End If
                Next custName

                If Not foundOne Then
                    LogPDFProblem 0, kamName, "Missing PDF for ExternalDocNr='" & CStr(pivOrder(i)) & "'"
                End If
            End If
        Next i

        ' 3) Merge if at least one PDF was found
        If pdfCount > 0 Then
            mergedPDF = Left$(fullPath, Len(fullPath) - 4) & ".pdf"  ' same base name as Excel
            If MergePDFs(pdfList, mergedPDF) Then
                Debug.Print "Merged PDF created: " & mergedPDF
            Else
                LogPDFProblem 0, kamName, "Merge failed – Ghostscript returned an error."
            End If
        Else
            LogPDFProblem 0, kamName, "No PDFs found for this KAM – merge skipped."
        End If

SkipMerge:
CloseWorkbookAndNextKAM:
        If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing

NextKAM:
        rsKAM.MoveNext
    Loop

CleanExit:
    If Not rsKAM Is Nothing Then rsKAM.Close: Set rsKAM = Nothing
    Set db = Nothing
    If Not xlApp Is Nothing Then xlApp.Quit: Set xlApp = Nothing
    MsgBox "KAM export completed successfully!", vbInformation
    Exit Sub

SaveFailed:
    Dim errMsg As String
    errMsg = "SaveAs failed – " & Err.Number & ": " & Err.Description & _
             ". FullPath='" & fullPath & "'"
    LogProblem 0, kamName, errMsg
    Err.Clear
    On Error Resume Next
    If Not xlWB Is Nothing Then xlWB.Close SaveChanges:=False
    On Error GoTo 0
    GoTo CloseWorkbookAndNextKAM

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "KAM Export failed"
    ' fall through to clean-up
    Resume CleanUp

CleanUp:
    On Error Resume Next
    If Not rsData Is Nothing Then rsData.Close: Set rsData = Nothing
    If Not rsKAM Is Nothing Then rsKAM.Close: Set rsKAM = Nothing
    If Not xlWB Is Nothing Then xlWB.Close False: Set xlWB = Nothing
    If Not xlApp Is Nothing Then xlApp.Quit: Set xlApp = Nothing
End Sub


