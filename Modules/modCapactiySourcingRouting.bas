Option Compare Database
Option Explicit

'========================
' CONFIG
'========================
Private Const PERMISSIONS_ID_EVELYN As Long = 304
Private Const OUTSOURCE_TEAM_TO As String = "BeardE@us.nifco.com"
Private Const OUTSOURCE_TEAM_CC As String = "HornT@us.nifco.com; capacityrequest@us.nifco.com"
Private Const CUTOFF_DATE As Date = #2/18/2026#

'--- tblCapacityRequests fields
Private Const T_Requests As String = "tblCapacityRequests"
Private Const F_RecordID As String = "RecordID"
Private Const F_RequestDate As String = "requestDate"
Private Const F_PartNumber As String = "partNumber"
Private Const F_Unit As String = "Unit"
Private Const F_ProductionType As String = "productionType"
Private Const F_Planner As String = "Planner"
Private Const F_SourcingNotifiedOn As String = "SourcingNotifiedOn"
Private Const F_SourcingNotifiedBy As String = "SourcingNotifiedBy"
Private Const F_SourcingNotifedFlag As String = "SourcingNotifiedFlag" '<< your actual column

'--- Lookup tables
' tblUnits
Private Const T_Units As String = "tblUnits"
Private Const U_RecordIDField As String = "recordID"
Private Const U_UnitNameField As String = "unitName"

' tblDropDowns_StrategicPlanning
Private Const T_Dropdowns As String = "tblDropDowns_StrategicPlanning"
Private Const D_RecordIDField As String = "recordId"
Private Const D_ProductionTypeField As String = "productionType"

' tblPermissions
Private Const T_Permissions As String = "tblPermissions"
Private Const P_IDField As String = "ID"
Private Const P_FirstNameField As String = "firstName"
Private Const P_LastNameField As String = "lastName"

' tblStratPlanAttachmentsSP
Private Const T_Attach As String = "tblStratPlanAttachmentsSP"
Private Const A_ReferenceId As String = "referenceId"
Private Const A_DirectLink As String = "directLink"

'========================
' PUBLIC ENTRY POINT (Form-based)
'========================
Public Sub HandleSourcingRouting(frm As Form)
    On Error GoTo ErrHandler

    Dim recId As Long
    recId = CLng(Nz(GetFieldSafe(frm, F_RecordID), 0))
    If recId = 0 Then Exit Sub

    'Already notified? stop.
    If Nz(GetFieldSafe(frm, F_SourcingNotifedFlag), False) = True Then Exit Sub
    If Not IsNull(GetFieldSafe(frm, F_SourcingNotifiedOn)) Then Exit Sub

    'Gate: cutoff (BEFORE any action)
    If Not IsOnOrAfterCutoff_Form(frm) Then Exit Sub

    'Gate: rule match
    If Not Rule_U6_Purchased_Form(frm) Then Exit Sub

    'Action
    Apply_U6Purchased_RouteAndNotify_Form frm

ExitHere:
    Exit Sub
ErrHandler:
    Debug.Print "HandleSourcingRouting error " & Err.Number & ": " & Err.Description
    Resume ExitHere
End Sub

'========================
' PUBLIC ENTRY POINT (Table-based, NO UI)
'========================
Public Sub HandleSourcingRoutingById(ByVal recId As Long)
    On Error GoTo ErrHandler
    If recId <= 0 Then Exit Sub

    Dim rs As DAO.Recordset
    Dim sql As String

    sql = "SELECT * FROM [" & T_Requests & "] WHERE [" & F_RecordID & "]=" & recId & ";"
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)

    If rs.EOF Then GoTo Cleanup

    'Already notified? stop.
    If Nz(rs.Fields(F_SourcingNotifedFlag).value, False) = True Then GoTo Cleanup
    If Not IsNull(rs.Fields(F_SourcingNotifiedOn).value) Then GoTo Cleanup

    'Gate: cutoff (BEFORE any action)
    If Not IsOnOrAfterCutoff_RS(rs) Then GoTo Cleanup

    'Gate: rule match
    If Not Rule_U6_Purchased_RS(rs) Then GoTo Cleanup

    'Action
    Apply_U6Purchased_RouteAndNotify_RS rs

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Sub

ErrHandler:
    Debug.Print "HandleSourcingRoutingById error " & Err.Number & ": " & Err.Description
    Resume Cleanup
End Sub

'========================
' GATES (Form-based)
'========================
Private Function Rule_U6_Purchased_Form(frm As Form) As Boolean
    Dim unitId As Variant: unitId = GetFieldSafe(frm, F_Unit)
    Dim prodId As Variant: prodId = GetFieldSafe(frm, F_ProductionType)
    If IsNull(unitId) Or IsNull(prodId) Then Exit Function

    Dim u6Id As Variant: u6Id = GetUnitId("U6")
    Dim purchasedId As Variant: purchasedId = GetProductionTypeId("Purchased")

    If IsNull(u6Id) Or IsNull(purchasedId) Then Exit Function

    Rule_U6_Purchased_Form = (CLng(unitId) = CLng(u6Id) And CLng(prodId) = CLng(purchasedId))
End Function

Private Function IsOnOrAfterCutoff_Form(frm As Form) As Boolean
    Dim d As Variant: d = GetFieldSafe(frm, F_RequestDate)
    If IsNull(d) Then
        IsOnOrAfterCutoff_Form = False
    Else
        IsOnOrAfterCutoff_Form = (CDate(d) >= CUTOFF_DATE)
    End If
End Function

'========================
' GATES (Recordset-based)
'========================
Private Function Rule_U6_Purchased_RS(rs As DAO.Recordset) As Boolean
    On Error GoTo ErrHandler

    Dim unitId As Long, prodId As Long
    unitId = CLng(Nz(rs.Fields(F_Unit).value, 0))
    prodId = CLng(Nz(rs.Fields(F_ProductionType).value, 0))
    If unitId = 0 Or prodId = 0 Then
        Rule_U6_Purchased_RS = False
        Exit Function
    End If

    Dim u6Id As Variant: u6Id = GetUnitId("U6")
    Dim purchasedId As Variant: purchasedId = GetProductionTypeId("Purchased")
    If IsNull(u6Id) Or IsNull(purchasedId) Then
        Rule_U6_Purchased_RS = False
        Exit Function
    End If

    Rule_U6_Purchased_RS = (unitId = CLng(u6Id) And prodId = CLng(purchasedId))
    Exit Function

ErrHandler:
    Rule_U6_Purchased_RS = False
End Function

Private Function IsOnOrAfterCutoff_RS(rs As DAO.Recordset) As Boolean
    On Error GoTo ErrHandler

    Dim d As Variant
    d = rs.Fields(F_RequestDate).value

    If IsNull(d) Then
        IsOnOrAfterCutoff_RS = False
    Else
        IsOnOrAfterCutoff_RS = (CDate(d) >= CUTOFF_DATE)
    End If
    Exit Function

ErrHandler:
    IsOnOrAfterCutoff_RS = False
End Function

'========================
' ACTION (Form-based)
'========================
Private Sub Apply_U6Purchased_RouteAndNotify_Form(frm As Form)
    On Error GoTo ErrHandler

    '1) Assign Planner to Evelyn
    If Nz(GetFieldSafe(frm, F_Planner), 0) <> PERMISSIONS_ID_EVELYN Then
        frm(F_Planner).value = PERMISSIONS_ID_EVELYN
    End If

    '2) Email outsource team
    SendOutsourceEmail_Form frm

    '3) Stamp flags
    frm(F_SourcingNotifiedOn).value = Now()
    frm(F_SourcingNotifiedBy).value = GetBestUserName()
    frm(F_SourcingNotifedFlag).value = True

    If frm.Dirty Then frm.Dirty = False
    Exit Sub

ErrHandler:
    Debug.Print "Apply_U6Purchased_RouteAndNotify_Form error " & Err.Number & ": " & Err.Description
End Sub

'========================
' ACTION (Recordset-based, NO UI)
'========================
Private Sub Apply_U6Purchased_RouteAndNotify_RS(rs As DAO.Recordset)
    On Error GoTo ErrHandler

    Dim recId As Long
    recId = CLng(Nz(rs.Fields(F_RecordID).value, 0))
    If recId = 0 Then Exit Sub

    'Update table (planner + stamps) FIRST to prevent duplicates if multiple users scan
    Dim sqlUpdate As String
    sqlUpdate = _
        "UPDATE [" & T_Requests & "] SET " & _
        "[" & F_Planner & "]=" & PERMISSIONS_ID_EVELYN & ", " & _
        "[" & F_SourcingNotifiedOn & "]=Now(), " & _
        "[" & F_SourcingNotifiedBy & "]='" & Replace(GetBestUserName(), "'", "''") & "', " & _
        "[" & F_SourcingNotifedFlag & "]=True " & _
        "WHERE [" & F_RecordID & "]=" & recId & ";"

    CurrentDb.Execute sqlUpdate, dbFailOnError

    'Send email
    SendOutsourceEmail_RS rs
    Exit Sub

ErrHandler:
    Debug.Print "Apply_U6Purchased_RouteAndNotify_RS error " & Err.Number & ": " & Err.Description
End Sub

'========================
' EMAIL (Form-based)
'========================
Private Sub SendOutsourceEmail_Form(frm As Form)
    Dim recId As Long
    recId = CLng(Nz(GetFieldSafe(frm, F_RecordID), 0))
    If recId = 0 Then Exit Sub

    Dim unitText As String
    unitText = GetComboDisplayText(frm, F_Unit)
    If Len(unitText) = 0 Then unitText = "U6"

    Dim partNo As String
    partNo = Nz(GetFieldSafe(frm, F_PartNumber), "")
    If Len(partNo) = 0 Then partNo = "UnknownPart"

    Dim subj As String
    subj = "Outsource Capacity Requests | " & unitText & " | " & partNo & " | RecordID " & recId

    Dim attachHtml As String
    attachHtml = BuildAttachmentsHtml(recId)

    Dim html As String
    html = ""
    html = html & "<html><body style='font-family:Segoe UI, Arial; font-size:10.5pt;'>"
    html = html & "<p>Hello,</p>"
    html = html & "<p>Please see below for the details of the outsource capacity request.</p>"

    If Len(attachHtml) > 0 Then
        html = html & "<hr/>" & attachHtml
    End If

    html = html & "<hr/>"
    html = html & BuildHtmlDetailsTableFromForm(frm)

    html = html & "<p style='font-size:9pt;color:gray;'>This is an automated notification from the Capacity Request System.</p>"
    html = html & "</body></html>"

    SendEmail_OutlookLateBound OUTSOURCE_TEAM_TO, OUTSOURCE_TEAM_CC, subj, html
End Sub

'========================
' EMAIL (Recordset-based)
'========================
Private Sub SendOutsourceEmail_RS(rs As DAO.Recordset)
    On Error GoTo ErrHandler

    Dim recId As Long
    recId = CLng(Nz(rs.Fields(F_RecordID).value, 0))
    If recId = 0 Then Exit Sub

    Dim unitName As String
    unitName = Nz(DLookup("[" & U_UnitNameField & "]", T_Units, "[" & U_RecordIDField & "]=" & CLng(Nz(rs.Fields(F_Unit).value, 0))), "U6")
    If Len(unitName) = 0 Then unitName = "U6"

    Dim partNo As String
    partNo = Nz(rs.Fields(F_PartNumber).value, "")
    If Len(partNo) = 0 Then partNo = "UnknownPart"

    Dim subj As String
    subj = "Outsource Capacity Requests | " & unitName & " | " & partNo & " | RecordID " & recId

    Dim attachHtml As String
    attachHtml = BuildAttachmentsHtml(recId)

    Dim html As String
    html = ""
    html = html & "<html><body style='font-family:Segoe UI, Arial; font-size:10.5pt;'>"
    html = html & "<p>Hello,</p>"
    html = html & "<p>Please see below for the details of the outsource capacity request.</p>"

    If Len(attachHtml) > 0 Then
        html = html & "<hr/>" & attachHtml
    End If

    html = html & "<hr/>"
    html = html & BuildHtmlDetailsTableFromRecordset(rs)

    html = html & "<p style='font-size:9pt;color:gray;'>This is an automated notification from the Capacity Request System.</p>"
    html = html & "</body></html>"

    SendEmail_OutlookLateBound OUTSOURCE_TEAM_TO, OUTSOURCE_TEAM_CC, subj, html
    Exit Sub

ErrHandler:
    Debug.Print "SendOutsourceEmail_RS error " & Err.Number & ": " & Err.Description
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
        .Send 'Production mmode
        '.Display 'Testing mode
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

    Dim rs As DAO.Recordset
    Dim sql As String
    Dim out As String
    Dim i As Long

    sql = "SELECT [" & A_DirectLink & "] " & _
          "FROM [" & T_Attach & "] " & _
          "WHERE [" & A_ReferenceId & "]=" & capacityRequestId & " " & _
          "AND Nz([" & A_DirectLink & "],'')<>'';"

    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)

    If (rs.BOF And rs.EOF) Then
        BuildAttachmentsHtml = ""
        rs.Close: Set rs = Nothing
        Exit Function
    End If

    out = "<p><b>Attachments</b></p><ul>"

    i = 0
    Do While Not rs.EOF
        i = i + 1

        Dim link As String
        link = Nz(rs.Fields(0).value, "")

        out = out & "<li>Attachment " & i & ": " & _
                    "<a href='" & HtmlAttributeEncode(link) & "'>Open</a>" & _
                    " &nbsp;(" & HtmlEncode(link) & ")</li>"

        rs.MoveNext
    Loop

    out = out & "</ul>"

    rs.Close: Set rs = Nothing
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

Private Function BuildHtmlDetailsTableFromRecordset(rs As DAO.Recordset) As String
    On Error GoTo ErrHandler
 
    Dim s As String
    s = "<table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse;'>" & _
        "<tr style='background:#f3f3f3;'><th align='left'>Field</th><th align='left'>Value</th></tr>"
 
    Dim f As DAO.Field
    For Each f In rs.Fields
 
        Dim nm As String
        nm = f.name
 
        'Skip internal fields so email stays clean
        If ShouldIncludeField(nm) Then
 
            Dim v As Variant
            v = f.value
 
            If Not IsNull(v) Then
                Dim outVal As String
                outVal = FormatFieldValue(nm, v, rs)
 
                If Len(Trim$(outVal)) > 0 Then
                    s = s & "<tr><td>" & HtmlEncode(nm) & "</td><td>" & HtmlEncode(outVal) & "</td></tr>"
                End If
            End If
 
        End If
    Next f
 
    s = s & "</table>"
 
    BuildHtmlDetailsTableFromRecordset = s
    Exit Function
 
ErrHandler:
    BuildHtmlDetailsTableFromRecordset = "<p>(Could not render request detail table.)</p>"
End Function
 
Private Function ShouldIncludeField(ByVal fieldName As String) As Boolean
    Dim n As String
    n = LCase$(fieldName)
 
    'Skip routing / audit / processing fields
    If n = LCase$(F_SourcingNotifiedOn) Then Exit Function
    If n = LCase$(F_SourcingNotifiedBy) Then Exit Function
    If n = LCase$(F_SourcingNotifedFlag) Then Exit Function
 
    'Skip claim fields used by scanner (if you added them)
    If n = "sourcingprocessingon" Then Exit Function
    If n = "sourcingprocessingby" Then Exit Function
 
    'Optional: skip Attachments reference fields if you don’t want them as raw
    'If n = "referencetable" Then Exit Function
    'If n = "referenceid" Then Exit Function
 
    ShouldIncludeField = True
End Function
 
Private Function FormatFieldValue(ByVal fieldName As String, ByVal v As Variant, ByVal rs As DAO.Recordset) As String
    On Error GoTo Fallback
 
    Dim n As String
    n = LCase$(fieldName)
 
    'Friendly lookups for ID fields
    If n = LCase$(F_Unit) Then
        Dim unitId As Long: unitId = CLng(Nz(v, 0))
        If unitId > 0 Then
            FormatFieldValue = Nz(DLookup("[" & U_UnitNameField & "]", T_Units, "[" & U_RecordIDField & "]=" & unitId), CStr(v))
        Else
            FormatFieldValue = ""
        End If
        Exit Function
    End If
 
    If n = LCase$(F_ProductionType) Then
        Dim prodId As Long: prodId = CLng(Nz(v, 0))
        If prodId > 0 Then
            FormatFieldValue = Nz(DLookup("[" & D_ProductionTypeField & "]", T_Dropdowns, "[" & D_RecordIDField & "]=" & prodId), CStr(v))
        Else
            FormatFieldValue = ""
        End If
        Exit Function
    End If
 
    If n = LCase$(F_Planner) Then
        Dim personId As Long: personId = CLng(Nz(v, 0))
        If personId > 0 Then
            FormatFieldValue = GetPermissionsDisplayName(personId)
        Else
            FormatFieldValue = ""
        End If
        Exit Function
    End If
 
    'Checkbox/boolean formatting
    If VarType(v) = vbBoolean Then
        FormatFieldValue = IIf(v, "Yes", "No")
        Exit Function
    End If
 
    'Date formatting
    If IsDate(v) Then
        FormatFieldValue = Format$(CDate(v), "yyyy-mm-dd hh:nn")
        Exit Function
    End If
 
    'Default
    FormatFieldValue = CStr(v)
    Exit Function
 
Fallback:
    FormatFieldValue = Nz(v, "")
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
    text = Replace(text, """", "&quot;")
    text = Replace(text, "'", "&#39;")
    HtmlAttributeEncode = text
End Function