Attribute VB_Name = "SendKAMEmails_module"
Option Compare Database
Option Explicit

Sub SendKAMEmails()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rsKAM As DAO.Recordset
    Dim rsMonth As DAO.Recordset
    Dim qdf As DAO.QueryDef
    Dim MaxPostingMonth As Date
    Dim kamName As String
    Dim folderPath As String
    Dim toEmail As String, ccEmail As String, bccEmail As String
    Dim emailBody As String
    Dim outlookApp As Object, outlookMail As Object
    Dim fso As Object, folder As Object, file As Object
    Dim logFile As Object, logPath As String
    Dim sqlText As String

    Set db = CurrentDb

    ' Get latest Posting Month
    Set rsMonth = db.OpenRecordset("SELECT Max([Posting Month]) AS MaxMonth FROM [Consumables]", dbOpenSnapshot)
    If Not rsMonth.EOF Then
        MaxPostingMonth = rsMonth!MaxMonth
    End If
    rsMonth.Close
    Set rsMonth = Nothing

    ' Initialize FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Prepare log file
    logPath = Application.CurrentProject.Path & "\EmailLog.txt"
    Set logFile = fso.OpenTextFile(logPath, ForAppending, True)

    ' Initialize Outlook
    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo ErrorHandler

    ' Prepare SQL and QueryDef
    sqlText = "SELECT [KAM_Name], [KAM_toemail], [KAM_ccemail], [KAM_bccemail] " & _
              "FROM [KeyAccountManager] WHERE [KAM_entity] = 'GHD' AND [KAM_entity] <> ''"
    Set qdf = db.CreateQueryDef("", sqlText)
    Set rsKAM = qdf.OpenRecordset(dbOpenSnapshot)

    Do While Not rsKAM.EOF
        kamName = Nz(rsKAM!KAM_Name, "")
        toEmail = Trim(Nz(rsKAM!KAM_toemail, ""))
        ccEmail = Trim(Nz(rsKAM!KAM_ccemail, ""))
        bccEmail = Trim(Nz(rsKAM!KAM_bccemail, ""))

        If toEmail = "" Then
            logFile.WriteLine "? Skipped " & kamName & ": No email address provided."
            GoTo SkipToNextKAM
        End If

        folderPath = "C:\Users\purceld\Desktop\Test2\Consumables " & Format(MaxPostingMonth, "mm yyyy") & _
                     "\Consumables " & kamName & " " & Format(MaxPostingMonth, "mm yyyy")

        logFile.WriteLine "Processing KAM: " & kamName & " | Folder: " & folderPath

        If fso.FolderExists(folderPath) Then
            Set folder = fso.GetFolder(folderPath)
            Set outlookMail = outlookApp.CreateItem(0)

            emailBody = "Hello," & vbCrLf & vbCrLf & _
                        "Please find attached the consumables for " & Format(MaxPostingMonth, "mmmm yyyy") & "." & vbCrLf & vbCrLf & _
                        "If there are any requests, please send them to Boekhouding.Gent@Katoennatie.com." & vbCrLf & vbCrLf & _
                        "Kind regards,"

            With outlookMail
                .To = toEmail
                .CC = ccEmail
                .BCC = bccEmail
                .Subject = "Consumables " & Format(MaxPostingMonth, "mmmm yyyy")
                .body = emailBody

                For Each file In folder.files
                    .Attachments.Add file.Path
                Next file

                .Send
            End With

            logFile.WriteLine "? Email sent to " & kamName & " (" & toEmail & ") at " & Now
        Else
            logFile.WriteLine "?? Folder not found for " & kamName & ": " & folderPath
        End If

SkipToNextKAM:
        rsKAM.MoveNext
    Loop

    rsKAM.Close
    Set rsKAM = Nothing
    logFile.Close
    Set logFile = Nothing

    MsgBox "Emails processed. Check EmailLog.txt for details.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    If Not logFile Is Nothing Then
        logFile.WriteLine "? General error: " & Err.Description & " at " & Now
        logFile.Close
    End If
    If Not rsKAM Is Nothing Then rsKAM.Close
    If Not rsMonth Is Nothing Then rsMonth.Close
End Sub




