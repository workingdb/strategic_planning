Option Compare Database
Option Explicit
 
' ==============================
' TABLE CONFIG (VERIFY SPELLING)
' ==============================
Private Const REQUESTS_TABLE As String = "tblCapacityRequests"
Private Const PERMISSIONS_TABLE As String = "tblPermissions"
Private Const ATTACH_TABLE As String = "tblStratPlanAttachmentsSP"
Private Const RESULTS_TABLE As String = "tblDropDowns_StrategicPlanning"
 
' tblCapacityRequests fields
Private Const FLD_PK As String = "RecordID"
Private Const FLD_REQUESTOR As String = "Requestor"
Private Const FLD_RESULT As String = "capacityResults"
Private Const FLD_NOTIFIED_ON As String = "NotifiedOn"
Private Const FLD_NOTIFIED_BY As String = "NotifiedBy"
 
' tblPermissions fields
Private Const FLD_PERM_ID As String = "ID"
Private Const FLD_PERM_EMAIL As String = "userEmail"
 
' tblStratPlanAttachmentsSP fields
Private Const FLD_ATTACH_REFID As String = "referenceId"
Private Const FLD_ATTACH_URL As String = "directLink"
Private Const FLD_ATTACH_NAME As String = "attachFullFileName"
 
' Dropdown table fields
Private Const FLD_DD_ID As String = "recordId"
Private Const FLD_DD_TEXT As String = "results"
 
' ==============================
' MAIN PUBLIC FUNCTION
' ==============================
Public Function NotifyCapacityResultIfNeeded(ByVal RecordID As Long) As Boolean
    On Error GoTo ErrHandler
 
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
 
    Set db = CurrentDb
 
    sql = "SELECT [" & FLD_PK & "], [" & FLD_REQUESTOR & "], [" & FLD_RESULT & "], [" & FLD_NOTIFIED_ON & "], [" & FLD_NOTIFIED_BY & "]" & _
          " FROM [" & REQUESTS_TABLE & "]" & _
          " WHERE [" & FLD_PK & "]=" & RecordID & ";"
 
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)
 
    If rs.EOF Then GoTo Cleanup
 
    ' Must have result (works for numeric or text)
    If Len(Trim(Nz(rs.Fields(FLD_RESULT).Value, ""))) = 0 Or Nz(rs.Fields(FLD_RESULT).Value, 0) = 0 Then GoTo Cleanup
 
    ' Must NOT already have notified
    If Not IsNull(rs.Fields(FLD_NOTIFIED_ON).Value) Then GoTo Cleanup
 
    ' Requestor ID
    Dim reqId As Long
    reqId = CLng(Nz(rs.Fields(FLD_REQUESTOR).Value, 0))
    If reqId = 0 Then GoTo Cleanup
 
    ' Requestor Email (fully bracketed)
    Dim toAddr As String
    toAddr = Trim(Nz(DLookup("[" & FLD_PERM_EMAIL & "]", "[" & PERMISSIONS_TABLE & "]", "[" & FLD_PERM_ID & "]=" & reqId), ""))
    If Len(toAddr) = 0 Then GoTo Cleanup
 
    ' Result ID (stored in tblCapacityRequests.capacityResults)
    Dim resultId As Long
    resultId = CLng(Nz(rs.Fields(FLD_RESULT).Value, 0))
 
    ' Lookup visible result text from dropdown table (second column in your combo)
    Dim resultText As String
    If resultId > 0 Then
        resultText = Nz(DLookup("[" & FLD_DD_TEXT & "]", "[" & RESULTS_TABLE & "]", "[" & FLD_DD_ID & "]=" & resultId), "")
    Else
        resultText = ""
    End If
 
    If Len(Trim(resultText)) = 0 Then
        ' fallback: send stored value if lookup fails
        resultText = CStr(Nz(rs.Fields(FLD_RESULT).Value, ""))
    End If
 
    ' Attachments HTML
    Dim attachmentsHtml As String
    attachmentsHtml = BuildAttachmentsHtml_(RecordID)
 
    ' Build Email
    Dim subj As String
    subj = "Capacity Results Updated | Request ID " & RecordID
 
    Dim html As String
    html = "<html><body style='font-family:Segoe UI, Arial; font-size:11pt;'>" & _
           "<p>Hello,</p>" & _
           "<p>Your capacity request has been reviewed.</p>" & _
           "<p><b>Request ID:</b> " & HtmlEncode_(CStr(RecordID)) & "</p>" & _
           "<p><b>Result:</b><br>" & HtmlEncode_(resultText) & "</p>"
 
    If Len(attachmentsHtml) > 0 Then html = html & attachmentsHtml
 
    html = html & "<p>Thank you,<br>Capacity Team</p>" & _
                  "</body></html>"
 
    ' Send via Outlook (display while testing)
    SendOutlookHtmlMail_ toAddr, subj, html
 
    ' Stamp notification
    rs.Edit
    rs.Fields(FLD_NOTIFIED_ON).Value = Now()
    rs.Fields(FLD_NOTIFIED_BY).Value = Environ$("USERNAME")
    rs.Update
 
    NotifyCapacityResultIfNeeded = True
 
Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
 
ErrHandler:
    MsgBox "NotifyCapacityResultIfNeeded error " & Err.Number & ":" & vbCrLf & Err.Description, vbExclamation
    NotifyCapacityResultIfNeeded = False
    Resume Cleanup
End Function
 
' ==============================
' ATTACHMENT HTML BUILDER
' ==============================
Private Function BuildAttachmentsHtml_(ByVal RecordID As Long) As String
    On Error GoTo ErrHandler
 
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
 
    Set db = CurrentDb
 
    sql = "SELECT [" & FLD_ATTACH_URL & "], [" & FLD_ATTACH_NAME & "]" & _
          " FROM [" & ATTACH_TABLE & "]" & _
          " WHERE [" & FLD_ATTACH_REFID & "]=" & RecordID & ";"
 
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
 
    If rs.EOF Then
        BuildAttachmentsHtml_ = ""
        GoTo Cleanup
    End If
 
    Dim listItems As String
    listItems = ""
 
    Do While Not rs.EOF
        Dim url As String, label As String
        url = Trim(Nz(rs.Fields(FLD_ATTACH_URL).Value, ""))
        label = Trim(Nz(rs.Fields(FLD_ATTACH_NAME).Value, ""))
 
        If Len(url) > 0 Then
            If Len(label) = 0 Then label = "Open attachment"
            listItems = listItems & "<li><a href='" & HtmlAttributeEncode_(url) & "'>" & HtmlEncode_(label) & "</a></li>"
        End If
 
        rs.MoveNext
    Loop
 
    If Len(listItems) > 0 Then
        BuildAttachmentsHtml_ = "<p><b>Attachments:</b></p><ul>" & listItems & "</ul>"
    Else
        BuildAttachmentsHtml_ = ""
    End If
 
Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
 
ErrHandler:
    MsgBox "BuildAttachmentsHtml_ error " & Err.Number & ":" & vbCrLf & Err.Description, vbExclamation
    BuildAttachmentsHtml_ = ""
    Resume Cleanup
End Function
 
' ==============================
' OUTLOOK SENDER
' ==============================
Private Sub SendOutlookHtmlMail_(ByVal toAddr As String, ByVal subj As String, ByVal html As String)
    Dim olApp As Object, mail As Object
    Set olApp = CreateObject("Outlook.Application")
    Set mail = olApp.CreateItem(0)
 
    With mail
        .To = toAddr
        .Subject = subj
        .HTMLBody = html
        .Send   ' switch to .Display for testing
    End With
End Sub
 
' ==============================
' HTML helpers (ASCII safe)
' ==============================
Private Function HtmlEncode_(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, vbCrLf, "<br>")
    HtmlEncode_ = s
End Function
 
Private Function HtmlAttributeEncode_(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "'", "&#39;")
    s = Replace(s, Chr(34), "&quot;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    HtmlAttributeEncode_ = s
End Function