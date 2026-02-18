Option Compare Database
Option Explicit
 
' ==============================
' TABLE CONFIG (VERIFY SPELLING)
' ==============================
Private Const REQUESTS_TABLE     As String = "tblCapacityRequests"
Private Const PERMISSIONS_TABLE  As String = "tblPermissions"
Private Const ATTACH_TABLE       As String = "tblStratPlanAttachmentsSP"
Private Const RESULTS_TABLE      As String = "tblDropDowns_StrategicPlanning"
Private Const CUSTOMERS_TABLE    As String = "tblCustomers"
 
' Email CC
Private Const CC_ADDR            As String = "capacityrequests@us.nifco.com"
 
' tblCapacityRequests fields (KNOWN)
Private Const FLD_PK             As String = "RecordID"
Private Const FLD_REQUESTOR      As String = "Requestor"
Private Const FLD_RESULT         As String = "capacityResults"
Private Const FLD_NOTIFIED_ON    As String = "NotifiedOn"
Private Const FLD_NOTIFIED_BY    As String = "NotifiedBy"
 
' tblPermissions fields
Private Const FLD_PERM_ID        As String = "ID"
Private Const FLD_PERM_EMAIL     As String = "userEmail"
Private Const FLD_PERM_FIRST     As String = "firstName"
 
' tblCustomers fields
Private Const FLD_CUST_ID        As String = "ID"
Private Const FLD_CUST_NAME      As String = "customerName"
 
' tblStratPlanAttachmentsSP fields
Private Const FLD_ATTACH_REFID   As String = "referenceId"
Private Const FLD_ATTACH_URL     As String = "directLink"
Private Const FLD_ATTACH_NAME    As String = "attachFullFileName"
 
' Dropdown table fields
Private Const FLD_DD_ID          As String = "recordId"
Private Const FLD_DD_RESULTS     As String = "results"         ' capacity result display text
Private Const FLD_DD_VOLTIMING   As String = "volumeTiming"    ' YES: spelled as you provided
 
' ==============================
' REQUEST FIELD CANDIDATES
' (Update ONLY if your names differ)
' ==============================
Private REQUEST_ITEM_FIELDS      As Variant
Private REQUEST_UNIT_FIELDS      As Variant
Private REQUEST_VOLUME_FIELDS    As Variant
Private REQUEST_VOLTIMING_FIELDS As Variant
Private REQUEST_SOP_FIELDS       As Variant
Private REQUEST_EOP_FIELDS       As Variant
Private REQUEST_CUSTOMER_FIELDS  As Variant
Private REQUEST_PROGRAM_FIELDS   As Variant
Private REQUEST_VEHMODEL_FIELDS  As Variant
Private REQUEST_NOTES_FIELDS     As Variant
 
Private Sub InitFieldCandidates_()
    ' Common patterns. Add/remove as needed.
    REQUEST_ITEM_FIELDS = Array("ItemNumber", "Item_Number", "Item", "Part", "PartNumber", "Part_Number")
    REQUEST_UNIT_FIELDS = Array("Unit", "UnitNo", "Unit_No", "UnitNumber", "Unit_Number")
    REQUEST_VOLUME_FIELDS = Array("Volume", "AnnualVolume", "Annual_Volume", "VolumeAmount", "Volume_Amount")
    REQUEST_VOLTIMING_FIELDS = Array("VolumeTiming", "Volume_Timing", "volumnTiming", "volumn_timing", "VolumeTimingID")
    REQUEST_SOP_FIELDS = Array("SOP", "StartOfProduction", "Start_Of_Production", "StartDate", "Start_Date")
    REQUEST_EOP_FIELDS = Array("EOP", "EndOfProduction", "End_Of_Production", "EndDate", "End_Date")
    REQUEST_CUSTOMER_FIELDS = Array("Customer", "CustomerID", "Customer_Id", "CustID", "Cust_Id")
    REQUEST_PROGRAM_FIELDS = Array("Program", "ProgramID", "Program_Id", "Platform", "PlatformID")
    REQUEST_VEHMODEL_FIELDS = Array("VehicleModel", "Vehicle_Model", "Model", "Vehicle", "VehicleName")
    REQUEST_NOTES_FIELDS = Array("Notes", "RequestNotes", "Comments", "DetailNotes", "Request_Notes")
End Sub
 
' ==============================
' MAIN PUBLIC FUNCTION
' ==============================
Public Function NotifyCapacityResultIfNeeded(ByVal RecordID As Long) As Boolean
    On Error GoTo ErrHandler
 
    InitFieldCandidates_
 
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
 
    Set db = CurrentDb
 
    sql = "SELECT * FROM [" & REQUESTS_TABLE & "] WHERE [" & FLD_PK & "]=" & RecordID & ";"
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)
 
    If rs.EOF Then GoTo Cleanup
 
    ' Must have result (works for numeric or text)
    If Len(Trim(Nz(rs.Fields(FLD_RESULT).value, ""))) = 0 Or Nz(rs.Fields(FLD_RESULT).value, 0) = 0 Then GoTo Cleanup
 
    ' Must NOT already have notified
    If Not IsNull(rs.Fields(FLD_NOTIFIED_ON).value) Then GoTo Cleanup
 
    ' Requestor ID
    Dim reqId As Long
    reqId = CLng(Nz(rs.Fields(FLD_REQUESTOR).value, 0))
    If reqId = 0 Then GoTo Cleanup
 
    ' Requestor email + first name
    Dim toAddr As String
    Dim firstName As String
 
    toAddr = Trim(Nz(DLookup("[" & FLD_PERM_EMAIL & "]", "[" & PERMISSIONS_TABLE & "]", "[" & FLD_PERM_ID & "]=" & reqId), ""))
    firstName = Trim(Nz(DLookup("[" & FLD_PERM_FIRST & "]", "[" & PERMISSIONS_TABLE & "]", "[" & FLD_PERM_ID & "]=" & reqId), ""))
 
    If Len(toAddr) = 0 Then GoTo Cleanup
    If Len(firstName) = 0 Then firstName = "there"
 
    ' Capacity Result display text
    Dim resultId As Long
    Dim resultText As String
 
    resultId = CLng(Nz(rs.Fields(FLD_RESULT).value, 0))
    If resultId > 0 Then
        resultText = Nz(DLookup("[" & FLD_DD_RESULTS & "]", "[" & RESULTS_TABLE & "]", "[" & FLD_DD_ID & "]=" & resultId), "")
    End If
    If Len(Trim(resultText)) = 0 Then resultText = CStr(Nz(rs.Fields(FLD_RESULT).value, ""))
 
    ' Pull other request fields (IDs or text)
    Dim itemNo As String, unitTxt As String, volTxt As String
    Dim sopTxt As String, eopTxt As String, programTxt As String, vehModelTxt As String, notesTxt As String
    Dim customerId As Long, customerName As String
    Dim volTimingId As Long, volTimingText As String
 
    itemNo = Nz(GetFirstFieldValue_(rs, REQUEST_ITEM_FIELDS), "")
    unitTxt = Nz(GetFirstFieldValue_(rs, REQUEST_UNIT_FIELDS), "")
    volTxt = Nz(GetFirstFieldValue_(rs, REQUEST_VOLUME_FIELDS), "")
 
    sopTxt = Nz(FormatMaybeDate_(GetFirstFieldValue_(rs, REQUEST_SOP_FIELDS)), "")
    eopTxt = Nz(FormatMaybeDate_(GetFirstFieldValue_(rs, REQUEST_EOP_FIELDS)), "")
 
    programTxt = Nz(GetFirstFieldValue_(rs, REQUEST_PROGRAM_FIELDS), "")
    vehModelTxt = Nz(GetFirstFieldValue_(rs, REQUEST_VEHMODEL_FIELDS), "")
    notesTxt = Nz(GetFirstFieldValue_(rs, REQUEST_NOTES_FIELDS), "")
 
    ' Customer (stored as ID in tblCapacityRequests ? lookup customerName)
    customerId = CLng(Nz(GetFirstFieldValue_(rs, REQUEST_CUSTOMER_FIELDS), 0))
    If customerId > 0 Then
        customerName = Trim(Nz(DLookup("[" & FLD_CUST_NAME & "]", "[" & CUSTOMERS_TABLE & "]", "[" & FLD_CUST_ID & "]=" & customerId), ""))
    Else
        customerName = ""
    End If
 
    ' Volume Timing (stored as ID in tblCapacityRequests ? lookup volumnTiming)
    volTimingId = CLng(Nz(GetFirstFieldValue_(rs, REQUEST_VOLTIMING_FIELDS), 0))
    If volTimingId > 0 Then
        volTimingText = Trim(Nz(DLookup("[" & FLD_DD_VOLTIMING & "]", "[" & RESULTS_TABLE & "]", "[" & FLD_DD_ID & "]=" & volTimingId), ""))
    Else
        volTimingText = ""
    End If
 
    ' Attachments HTML
    Dim attachmentsHtml As String
    attachmentsHtml = BuildAttachmentsHtml_(RecordID)
 
    ' Build Email
    Dim subj As String
    Dim partNumber As String
    Dim programName As String
 
        partNumber = Trim(Nz(rs.Fields("partnumber").value, ""))
        programName = Trim(Nz(rs.Fields("Program").value, ""))
 
        subj = "Capacity Result: " & resultText & _
       " | NAM " & partNumber & _
       " | Cust: " & customerName & _
       IIf(Len(programName) > 0, " | Program: " & programName, "")
 
    Dim html As String
    html = "<html><body style='font-family:Segoe UI, Arial; font-size:11pt;'>" & _
           "<p>Hello " & HtmlEncode_(firstName) & ",</p>" & _
           "<p>Your capacity request has been reviewed.</p>" & _
           BuildKeyValueHtml_("Request ID", CStr(RecordID)) & _
           BuildKeyValueHtml_("Item number", itemNo) & _
           BuildKeyValueHtml_("Unit", unitTxt) & _
           BuildKeyValueHtml_("Volume", volTxt) & _
           BuildKeyValueHtml_("Volume timing", volTimingText) & _
           BuildKeyValueHtml_("SOP", sopTxt) & _
           BuildKeyValueHtml_("EOP", eopTxt) & _
           BuildKeyValueHtml_("Customer", customerName) & _
           BuildKeyValueHtml_("Program", programTxt) & _
           BuildKeyValueHtml_("Vehicle model", vehModelTxt) & _
           BuildKeyValueHtml_("Capacity results", resultText) & _
           BuildKeyValueHtml_("Notes", notesTxt)
 
    If Len(attachmentsHtml) > 0 Then html = html & attachmentsHtml
 
    html = html & "<p>Thank you,<br>Capacity Team</p></body></html>"
 
    ' Send via Outlook
    SendOutlookHtmlMail_ toAddr, CC_ADDR, subj, html
 
    ' Stamp notification
    rs.Edit
    rs.Fields(FLD_NOTIFIED_ON).value = Now()
    rs.Fields(FLD_NOTIFIED_BY).value = Environ$("USERNAME")
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
        url = Trim(Nz(rs.Fields(FLD_ATTACH_URL).value, ""))
        label = Trim(Nz(rs.Fields(FLD_ATTACH_NAME).value, ""))
 
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
Private Sub SendOutlookHtmlMail_(ByVal toAddr As String, ByVal ccAddr As String, ByVal subj As String, ByVal html As String)
    On Error GoTo ErrHandler
 
    Dim olApp As Object, mail As Object
    Set olApp = CreateObject("Outlook.Application")
    Set mail = olApp.CreateItem(0)
 
    With mail
        .To = toAddr
        If Len(Trim(ccAddr)) > 0 Then .CC = ccAddr
        .Subject = subj
        .HTMLBody = html
        '.Send   ' change to .Display for testing
        .Display 'change to .Send for live
    End With
 
    Exit Sub
 
ErrHandler:
    MsgBox "SendOutlookHtmlMail_ error " & Err.Number & ":" & vbCrLf & Err.Description, vbExclamation
End Sub
 
' ==============================
'FIELD + FORMAT HELPERS
' ==============================
Private Function GetFirstFieldValue_(ByVal rs As DAO.Recordset, ByVal candidates As Variant) As Variant
    On Error GoTo ErrHandler
 
    Dim i As Long
    For i = LBound(candidates) To UBound(candidates)
        Dim f As String
        f = CStr(candidates(i))
        If FieldExists_(rs, f) Then
            GetFirstFieldValue_ = rs.Fields(f).value
            Exit Function
        End If
    Next i
 
    GetFirstFieldValue_ = Null
    Exit Function
 
ErrHandler:
    GetFirstFieldValue_ = Null
End Function
 
Private Function FieldExists_(ByVal rs As DAO.Recordset, ByVal fieldName As String) As Boolean
    On Error GoTo Nope
    Dim tmp As Variant
    tmp = rs.Fields(fieldName).name
    FieldExists_ = True
    Exit Function
Nope:
    FieldExists_ = False
End Function
 
Private Function FormatMaybeDate_(ByVal v As Variant) As String
    If IsNull(v) Then
        FormatMaybeDate_ = ""
    ElseIf IsDate(v) Then
        FormatMaybeDate_ = Format$(CDate(v), "m/d/yyyy")
    Else
        FormatMaybeDate_ = CStr(v)
    End If
End Function
 
Private Function BuildKeyValueHtml_(ByVal label As String, ByVal value As String) As String
    If Len(Trim(value)) = 0 Then
        BuildKeyValueHtml_ = ""
    Else
        BuildKeyValueHtml_ = "<p><b>" & HtmlEncode_(label) & ":</b> " & HtmlEncode_(value) & "</p>"
    End If
End Function
 
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
 