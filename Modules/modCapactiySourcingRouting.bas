Option Compare Database
Option Explicit
 
'========================
' CONFIG
'========================
Private Const PERMISSIONS_ID_EVELYN As Long = 304
Private Const OUTSOURCE_TEAM_TO As String = "BeardE@us.nifco.com" 'EDIT
Private Const OUTSOURCE_TEAM_CC As String = "HornT@us.nifco.com; capacityrequests@us.nifco.com"
Private Const CUTOFF_DATE As Date = #2/18/2026# 'date gate to prevent previous requests from generating emails
 
'tblCapacityRequests fields
Private Const F_RecordID As String = "RecordID"
Private Const F_RequestDate As String = "requestDate"
Private Const F_partNumber As String = "partNumber"
Private Const F_Unit As String = "Unit"
Private Const F_ProductionType As String = "productionType"
Private Const F_Planner As String = "Planner"
Private Const F_SourcingNotifiedOn As String = "SourcingNotifiedOn"
Private Const F_SourcingNotifiedBy As String = "SourcingNotifiedBy"
Private Const F_SourcingNotifedFlag As String = "SourcingNotifiedFlag"
 
'Lookup tables (adjust field names if different)
'tblUnits fields
Private Const T_Units As String = "tblUnits"
Private Const U_RecordIDField As String = "recordID"
Private Const U_UnitNameField As String = "unitName"
 'tblDropDowns_StrategicPlanning fields
Private Const T_Dropdowns As String = "tblDropDowns_StrategicPlanning"
Private Const D_RecordIDField As String = "recordId"
Private Const D_ProductionTypeField As String = "productionType" '<< change if needed
 'tblPermissions fields
Private Const T_Permissions As String = "tblPermissions"
Private Const P_IDField As String = "ID"
Private Const P_FirstNameField As String = "firstName"
Private Const P_LastNameField As String = "lastName"

'tblStratPlanAttachmentsSP
Private Const T_Attach As String = "tblStratPlanAttachmentsSP"
Private Const A_ReferenceId As String = "referenceId"
Private Const A_DirectLink As String = "directLink"

 
'========================
' PUBLIC ENTRY POINT
'========================
Public Sub HandleSourcingRouting(frm As Form)
    On Error GoTo ErrHandler
 
    Dim recId As Long
    recId = CLng(Nz(GetFieldSafe(frm, F_RecordID), 0))
    If recId = 0 Then Exit Sub
 
    'Do not re-run once already notified
    If Nz(GetFieldSafe(frm, F_SourcingNotifedFlag), False) = True Then Exit Sub
    If Not IsNull(GetFieldSafe(frm, F_SourcingNotifiedOn)) Then Exit Sub
 
    'Rules engine (add more rules later)
    If Rule_U6_Purchased(frm) Then
        Apply_U6Purchased_RouteAndNotify frm
        
    If Not IsOnOrAfterCutoff(frm) Then Exit Sub
    End If
 
ExitHere:
    Exit Sub
 
ErrHandler:
    Debug.Print "HandleSourcingRouting error " & Err.Number & ": " & Err.Description
    'Optional visible error:
    'MsgBox "Sourcing routing error: " & Err.Number & vbCrLf & Err.Description, vbExclamation
    Resume ExitHere
End Sub
 
'========================
' RULES
'========================
Private Function Rule_U6_Purchased(frm As Form) As Boolean
    Dim unitId As Variant: unitId = GetFieldSafe(frm, F_Unit)
    Dim prodId As Variant: prodId = GetFieldSafe(frm, F_ProductionType)
    If IsNull(unitId) Or IsNull(prodId) Then Exit Function
 
    Dim u6Id As Variant
    u6Id = GetUnitId("U6")
 
    Dim purchasedId As Variant
    purchasedId = GetProductionTypeId("Purchased")
 
    If IsNull(u6Id) Or IsNull(purchasedId) Then Exit Function
 
    Rule_U6_Purchased = (CLng(unitId) = CLng(u6Id) And CLng(prodId) = CLng(purchasedId))
End Function

Private Function IsOnOrAfterCutoff(frm As Form) As Boolean
    Dim d As Variant: d = GetFieldSafe(frm, F_RequestDate)
    If IsNull(d) Then
        IsOnOrAfterCutoff = False
    Else
        IsOnOrAfterCutoff = (CDate(d) >= CUTOFF_DATE)
    End If
End Function
 
'========================
' ACTION
'========================
Private Sub Apply_U6Purchased_RouteAndNotify(frm As Form)
    '1) Assign Planner to Evelyn
    If Nz(GetFieldSafe(frm, F_Planner), 0) <> PERMISSIONS_ID_EVELYN Then
        frm(F_Planner).value = PERMISSIONS_ID_EVELYN
    End If
 
    '2) Email outsource team
    SendOutsourceEmail frm
 
    '3) Stamp flags
    frm(F_SourcingNotifiedOn).value = Now()
    frm(F_SourcingNotifiedBy).value = GetBestUserName()
    frm(F_SourcingNotifedFlag).value = True
 
    If frm.Dirty Then frm.Dirty = False
End Sub

 
 
'========================
' EMAIL
'========================
Private Sub SendOutsourceEmail(frm As Form)
 
    Dim recId As Long
    recId = CLng(Nz(GetFieldSafe(frm, F_RecordID), 0))
    If recId = 0 Then Exit Sub
 
    Dim unitText As String
    unitText = GetComboDisplayText(frm, F_Unit)
 
    Dim partNo As String
    partNo = Nz(GetFieldSafe(frm, F_partNumber), "")
    If Len(partNo) = 0 Then partNo = "UnknownPart"
 
    Dim subj As String
    subj = "Outsource Capacity Requests | " & unitText & _
       " | NAM:" & partNo & _
       " | RecordID:" & recId
 
    Dim attachHtml As String
    attachHtml = BuildAttachmentsHtml(recId)
 
    Dim html As String
    html = ""
    html = html & "<html><body style='font-family:Segoe UI, Arial; font-size:10.5pt;'>"
 
    'Greeting
    html = html & "<p>Hello,</p>"
    html = html & "<p>Please see below for the details of the outsource capacity request.</p>"
 
    'Attachments (only if any)
    If Len(attachHtml) > 0 Then
        html = html & "<hr/>" & attachHtml
    End If
 
    'Full request details
    html = html & "<hr/>"
    html = html & BuildHtmlDetailsTableFromForm(frm)
 
    'Footer
    html = html & "<p style='font-size:9pt;color:gray;'>This is an automated notification from the Capacity Request System.</p>"
 
    html = html & "</body></html>"
 
    SendEmail_OutlookLateBound OUTSOURCE_TEAM_TO, OUTSOURCE_TEAM_CC, subj, html
 
End Sub
 
Private Sub SendEmail_OutlookLateBound(toList As String, ccList As String, subject As String, htmlBody As String)
    On Error GoTo ErrHandler
 
    Dim olApp As Object, mail As Object
 
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    On Error GoTo 0
 
    If olApp Is Nothing Then
        Set olApp = CreateObject("Outlook.Application")
    End If
 
    Set mail = olApp.CreateItem(0)
    With mail
        .To = toList
        .CC = ccList
        .subject = subject
        .htmlBody = htmlBody
        '.Send
        .Display 'Testing mode,update to send for auto
    End With
 
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "SendEmail_OutlookLateBound", Err.Description
End Sub
  
'========================
' ATTACHMENTS
'========================
 Private Function BuildAttachmentsHtml(capacityRequestId As Long) As String
    On Error GoTo ErrHandler
 
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim out As String
    Dim i As Long
 
    sql = "SELECT [" & A_DirectLink & "] " & _
          "FROM [" & T_Attach & "] " & _
          "WHERE [" & A_ReferenceId & "]=" & capacityRequestId & " " & _
          "AND Nz([" & A_DirectLink & "],'')<>'';"
 
    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
 
    If (rs.BOF And rs.EOF) Then
        BuildAttachmentsHtml = ""   'no attachments
        rs.Close
        Set rs = Nothing
        Set db = Nothing
        Exit Function
    End If
 
    out = "<p><b>Attachments</b></p><ul>"
 
    i = 0
    Do While Not rs.EOF
        i = i + 1
        Dim link As String
        link = Nz(rs.Fields(0).value, "")
 
        'Clickable link + show URL for copy/paste safety
        out = out & "<li>Attachment " & i & ": " & _
                    "<a href='" & HtmlAttributeEncode(link) & "'>Open</a>" & _
                    " &nbsp;(" & HtmlEncode(link) & ")</li>"
 
        rs.MoveNext
    Loop
 
    out = out & "</ul>"
 
    rs.Close
    Set rs = Nothing
    Set db = Nothing
 
    BuildAttachmentsHtml = out
    Exit Function
 
ErrHandler:
    BuildAttachmentsHtml = "<p><b>Attachments</b></p><p>(Error reading attachments.)</p>"
End Function
'========================
' LOOKUPS
'========================
Private Function GetUnitId(unitName As String) As Variant
    GetUnitId = DLookup("[" & U_RecordIDField & "]", T_Units, "[" & U_UnitNameField & "]='" & Replace(unitName, "'", "''") & "'")
End Function
 
Private Function GetProductionTypeId(prodTypeName As String) As Variant
    GetProductionTypeId = DLookup("[" & D_RecordIDField & "]", T_Dropdowns, "[" & D_ProductionTypeField & "]='" & Replace(prodTypeName, "'", "''") & "'")
End Function
 
Private Function GetPermissionsDisplayName(personId As Long) As String
    Dim fn As Variant, ln As Variant
    fn = DLookup("[" & P_FirstNameField & "]", T_Permissions, "[" & P_IDField & "]=" & personId)
    ln = DLookup("[" & P_LastNameField & "]", T_Permissions, "[" & P_IDField & "]=" & personId)
    GetPermissionsDisplayName = Trim(Nz(fn, "") & " " & Nz(ln, ""))
End Function
 
'========================
' GENERIC HELPERS
'========================
Private Function GetFieldSafe(frm As Form, fieldName As String) As Variant
    On Error GoTo ErrHandler
    GetFieldSafe = frm.Controls(fieldName).value
    Exit Function
ErrHandler:
    GetFieldSafe = Null
End Function
 
Private Function GetComboDisplayText(frm As Form, ctlName As String) As String
    On Error GoTo ErrHandler
    Dim c As Control: Set c = frm.Controls(ctlName)
 
    If c.ControlType = acComboBox Then
        GetComboDisplayText = Nz(c.Column(1), Nz(c.value, ""))
    Else
        GetComboDisplayText = Nz(c.value, "")
    End If
    Exit Function
ErrHandler:
    GetComboDisplayText = ""
End Function
 
Private Function BuildHtmlDetailsTableFromForm(frm As Form) As String
    On Error GoTo ErrHandler
 
    Dim s As String
    s = "<table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse;'>" & _
        "<tr style='background:#f3f3f3;'><th align='left'>Field</th><th align='left'>Value</th></tr>"
 
    Dim ctl As Control
    For Each ctl In frm.Controls
        Select Case ctl.ControlType
            Case acTextBox, acComboBox, acCheckBox
                Dim nm As String: nm = ctl.name
 
                'Skip internal routing fields so the email stays clean
                If LCase$(nm) <> LCase$(F_SourcingNotifiedOn) _
                   And LCase$(nm) <> LCase$(F_SourcingNotifiedBy) _
                   And LCase$(nm) <> LCase$(F_SourcingNotifedFlag) Then
 
                    Dim val As String
                    If ctl.ControlType = acComboBox Then
                        val = Nz(ctl.Column(1), Nz(ctl.value, ""))
                    ElseIf ctl.ControlType = acCheckBox Then
                        val = IIf(Nz(ctl.value, False), "Yes", "No")
                    Else
                        val = Nz(ctl.value, "")
                    End If
 
                    If Len(Trim$(val)) > 0 Then
                        s = s & "<tr><td>" & HtmlEncode(nm) & "</td><td>" & HtmlEncode(val) & "</td></tr>"
                    End If
                End If
        End Select
    Next ctl
 
    s = s & "</table>"
    BuildHtmlDetailsTableFromForm = s
    Exit Function
 
ErrHandler:
    BuildHtmlDetailsTableFromForm = "<p>(Could not render request detail table.)</p>"
End Function
 
Private Function GetBestUserName() As String
    GetBestUserName = Environ$("USERNAME")
    If Len(GetBestUserName) = 0 Then GetBestUserName = CurrentUser()
End Function
 
Private Function HtmlEncode(ByVal text As String) As String
    text = Replace(text, "&", "&amp;")
    text = Replace(text, "<", "&lt;")
    text = Replace(text, ">", "&gt;")
    text = Replace(text, """", "&quot;")
    text = Replace(text, "'", "&#39;")
    HtmlEncode = text
End Function

Private Function HtmlAttributeEncode(ByVal text As String) As String
    text = Replace(text, "&", "&amp;")
    text = Replace(text, "<", "&lt;")
    text = Replace(text, ">", "&gt;")
    text = Replace(text, """", "&quot:")
    text = Replace(text, "'", "&#39;")
    HtmlAttributeEncode = text
    

End Function