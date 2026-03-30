Option Compare Database
Option Explicit

Function sendNotification(sendTo As String, notType As Integer, notPriority As Integer, desc As String, emailContent As String, Optional appName As String = "", Optional appId As Variant = "", Optional multiEmail As Boolean = False, Optional customEmail As Boolean = False) As Boolean
sendNotification = True

On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()

'has this person been notified about this thing today already?
Dim rsNotifications As Recordset
Set rsNotifications = db.OpenRecordset("SELECT * from tblNotificationsSP WHERE recipientUser = '" & sendTo & "' AND notificationDescription = '" & StrQuoteReplace(desc) & "' AND sentDate > #" & Date - 1 & "#")
'NEEDS CONVERTED TO ADODB
If rsNotifications.RecordCount > 0 Then
    If rsNotifications!notificationType = 1 Then
        Dim msgTxt As String
        If rsNotifications!senderUser = Environ("username") Then
            msgTxt = "You already nudged this person today"
        Else
            msgTxt = sendTo & " has already been nudged about this today by " & rsNotifications!senderUser & ". Let's wait until tomorrow to nudge them again."
        End If
        MsgBox msgTxt, vbInformation, "Hold on a minute..."
        sendNotification = False
        Exit Function
    End If
End If

Dim strEmail
If customEmail = False Then
    Dim ITEM, sendToArr() As String
    If multiEmail Then
        sendToArr = Split(sendTo, ",")
        strEmail = ""
        For Each ITEM In sendToArr
            If ITEM = "" Then GoTo nextItem
            strEmail = strEmail & getEmail(CStr(ITEM)) & ";"
nextItem:
        Next ITEM
        If strEmail = "" Then Exit Function
        strEmail = Left(strEmail, Len(strEmail) - 1)
    Else
        strEmail = getEmail(sendTo)
    End If
Else
    strEmail = sendTo
    sendTo = Split(sendTo, "@")(0)
End If

Set rsNotifications = db.OpenRecordset("tblNotificationsSP")

With rsNotifications
    .addNew
    !recipientUser = sendTo
    !recipientEmail = strEmail
    !senderUser = Environ("username")
    !senderEmail = getEmail(Environ("username"))
    !sentDate = Now()
    !notificationType = notType
    !notificationPriority = notPriority
    !notificationDescription = desc
    !appName = appName
    !appId = appId
    !emailContent = emailContent
    .Update
End With

On Error Resume Next
rsNotifications.CLOSE
Set rsNotifications = Nothing
Set db = Nothing

Exit Function
Err_Handler:
sendNotification = False
    Call handleError("wdbGlobalFunctions", "sendNotification", err.Description, err.Number)
End Function

Function getEmail(userName As String) As String
On Error GoTo Err_Handler

getEmail = ""
On Error GoTo tryOracle
Dim db As Database
Set db = CurrentDb()
Dim rsPermissions As Recordset
Set rsPermissions = db.OpenRecordset("SELECT * from tblPermissions WHERE user = '" & userName & "'", dbOpenSnapshot)
getEmail = Nz(rsPermissions!userEmail, "")
rsPermissions.CLOSE
Set rsPermissions = Nothing

GoTo exitFunc

tryOracle:
Dim rsEmployee As Recordset
Set rsEmployee = db.OpenRecordset("SELECT FIRST_NAME, LAST_NAME, EMAIL_ADDRESS FROM APPS_XXCUS_USER_EMPLOYEES_V WHERE USER_NAME = '" & StrConv(userName, vbUpperCase) & "'", dbOpenSnapshot)
getEmail = Nz(rsEmployee!EMAIL_ADDRESS, "")
rsEmployee.CLOSE
Set rsEmployee = Nothing

exitFunc:
Set db = Nothing

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "getEmail", err.Description, err.Number)
End Function

Function generateHTML(Title As String, subTitle As String, primaryMessage As String, _
        detail1 As String, detail2 As String, detail3 As String, _
        Optional Link As String = "", _
        Optional addLines As Boolean = False, _
        Optional appName As String = "", _
        Optional appId As String = "") As String
        
On Error GoTo Err_Handler

Dim tblHeading As String, tblFooter As String, strHTMLBody As String

If Link <> "" Then
    primaryMessage = "<a href = '" & Link & "'>" & primaryMessage & "</a>"
ElseIf appId <> "" Then
    primaryMessage = "<a href = ""\\data\mdbdata\WorkingDB\build\workingdb_commands\openNotification.vbs"">" & primaryMessage & "</a>"
End If

tblHeading = "<table style=""width: 100%; margin: 0 auto; padding: 2em 3em; text-align: center; background-color: #fafafa;"">" & _
                            "<tbody>" & _
                                "<tr><td><h2 style=""color: #414141; font-size: 28px; margin-top: 0;"">" & Title & "</h2></td></tr>" & _
                                "<tr><td><p style=""color: rgb(73, 73, 73);"">" & subTitle & "</p></td></tr>" & _
                                 "<tr><td><table style=""padding: 1em; text-align: center;"">" & _
                                     "<tr><td style=""padding: 1em 1.5em; background: #FF6B00; "">" & primaryMessage & "</td></tr>" & _
                                "</table></td></tr>" & _
                            "</tbody>" & _
                        "</table>"
                        
tblFooter = "<table style=""width: 100%; margin: 0 auto; padding: 3em; background: #2b2b2b; color: rgba(255,255,255,.5);"">" & _
                        "<tbody>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: 1em; color: #c9c9c9;"">Details</td></tr>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em;"">" & detail1 & "</td></tr>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em;"">" & detail2 & "</td></tr>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em 1em 2em;"">" & detail3 & "</td></tr>" & _
                        "</tbody>" & _
                    "</table>"
                    
Dim addStuff As String
addStuff = ""
If addLines Then
    addStuff = "<table style=""max-width: 600px; margin: 0 auto; padding: 3em; background: #eaeaea; color: rgba(255,255,255,.5);"">" & _
        "<tr style=""border-collapse: collapse;""><td style=""padding: 1em;"">Extra Notes: type here...</td></tr></table>"
End If

strHTMLBody = "" & _
"<!DOCTYPE html><html lang=""en"" xmlns=""http://www.w3.org/1999/xhtml"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"">" & _
    "<head><meta charset=""utf-8""><title>Working DB Notification</title></head>" & _
    "<body style=""margin: 0 auto; Font-family: 'Montserrat', sans-serif; font-weight: 400; font-size: 15px; line-height: 1.8;"">" & addStuff & _
        "<table style=""max-width: 600px; margin: 0 auto; text-align: center; "">" & _
            "<tbody>" & _
                "<tr><td>" & tblHeading & "</td></tr>" & _
                "<tr><td>" & tblFooter & "</td></tr>" & _
                "<tr><td><p style=""color: rgb(192, 192, 192); text-align: center;"">AppName:[" & appName & "], AppId:[" & appId & "]</p></td></tr>" & _
                "<tr><td><p style=""color: rgb(192, 192, 192); text-align: center;"">This email was created by  &copy; workingDB</p></td></tr>" & _
            "</tbody>" & _
        "</table>" & _
    "</body>" & _
"</html>"

generateHTML = strHTMLBody

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "generateHTML", err.Description, err.Number)
End Function

Public Function wdbEmail(ByVal strTo As String, ByVal strCC As String, ByVal strSubject As String, body As String) As Boolean
On Error GoTo Err_Handler
wdbEmail = True
    
Dim objEmail As Object

Set objEmail = CreateObject("outlook.Application")
Set objEmail = objEmail.CreateItem(0)

With objEmail
    .To = strTo
    .CC = strCC
    .subject = strSubject
    .htmlBody = body
    .display
End With

Set objEmail = Nothing
    
Exit Function
Err_Handler:
wdbEmail = False
    Call handleError("wdbGlobalFunctions", "wdbEmail", err.Description, err.Number)
End Function

Function setSplashLoading(label As String)
On Error GoTo Err_Handler

If IsNull(TempVars!loadAmount) Then Exit Function
TempVars.Add "loadAmount", TempVars!loadAmount + 1
Form_frmSplash.lnLoading.Width = (TempVars!loadAmount / 12) * TempVars!loadWd
Form_frmSplash.lblLoading.Caption = label
Form_frmSplash.Repaint

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "setSplashLoading", err.Description, err.Number)
End Function

Function userData(data As String, Optional specificUser As String = "") As String
    On Error GoTo Err_Handler

    If specificUser = "" Then specificUser = Environ("username")

    Dim conn As ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    Set conn = CurrentProject.Connection
    
    ' Using brackets around the variable [data] and reserved word [User]
    strSQL = "SELECT [" & data & "] FROM tblPermissions WHERE [User] = '" & Replace(specificUser, "'", "''") & "'"
    
    rs.open strSQL, conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        userData = Nz(rs.Fields(0).Value, "")
    Else
        userData = ""
    End If

CleanUp:
    If rs.State = adStateOpen Then rs.CLOSE
    Set rs = Nothing
    Set conn = Nothing
    Exit Function

Err_Handler:
    Call handleError("wdbGlobalFunctions", "userData", err.Description, err.Number)
    Resume CleanUp
End Function

Function dbExecute(sql As String)
On Error GoTo Err_Handler

Dim conn As ADODB.Connection
Set conn = CurrentProject.Connection

conn.Execute sql

Set conn = Nothing

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "dbExecute", err.Description, err.Number, sql)
End Function

Public Sub registerStratPlanUpdates( _
    ByVal table As String, _
    ByVal ID As Variant, _
    ByVal column As String, _
    ByVal oldVal As Variant, _
    ByVal newVal As Variant, _
    ByVal referenceId As String, _
    ByVal formName As String, _
    Optional ByVal tag0 As Variant = "")

    On Error GoTo Err_Handler

    Dim cmd As ADODB.Command
    Dim oldText As String
    Dim newText As String
    Dim tagText As String

    ' Normalize dates
    If VarType(oldVal) = vbDate Then oldVal = Format$(oldVal, "mm/dd/yyyy")
    If VarType(newVal) = vbDate Then newVal = Format$(newVal, "mm/dd/yyyy")

    ' Normalize text values
    oldText = Left$(StrQuoteReplace(CStr(Nz(oldVal, ""))), 255)
    newText = Left$(StrQuoteReplace(CStr(Nz(newVal, ""))), 255)
    tagText = Left$(StrQuoteReplace(CStr(Nz(tag0, ""))), 255)

    ' Normalize blank ID to Null
    If Nz(ID, "") = "" Then ID = Null

    Set cmd = New ADODB.Command

    With cmd
        .ActiveConnection = CurrentProject.Connection
        .CommandType = adCmdText
        .CommandText = _
            "INSERT INTO tblStratPlan_UpdateTracking (" & _
            "tableName, tableRecordId, updatedBy, updatedDate, columnName, " & _
            "previousData, newData, referenceId, formName, dataTag0) " & _
            "VALUES (?, ?, ?, '" & Format$(Now(), "yyyy-mm-dd\Thh:nn:ss") & "', ?, ?, ?, ?, ?, ?)"

        .Parameters.Append .CreateParameter("pTableName", adVarChar, adParamInput, 100, table)

        If IsNull(ID) Then
            .Parameters.Append .CreateParameter("pTableRecordId", adInteger, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("pTableRecordId", adInteger, adParamInput, , ID)
        End If

        .Parameters.Append .CreateParameter("pUpdatedBy", adVarChar, adParamInput, 55, Environ$("username"))
        .Parameters.Append .CreateParameter("pColumnName", adVarChar, adParamInput, 100, column)
        .Parameters.Append .CreateParameter("pPreviousData", adVarChar, adParamInput, 255, oldText)
        .Parameters.Append .CreateParameter("pNewData", adVarChar, adParamInput, 255, newText)
        .Parameters.Append .CreateParameter("pReferenceId", adInteger, adParamInput, , referenceId)
        .Parameters.Append .CreateParameter("pFormName", adVarChar, adParamInput, 55, StrQuoteReplace(formName))
        .Parameters.Append .CreateParameter("pDataTag0", adVarChar, adParamInput, 55, tagText)

        .Execute , , adExecuteNoRecords
    End With

CleanExit:
    Set cmd = Nothing
    Exit Sub

Err_Handler:
    Call handleError("wdbGlobalFunctions", "registerStratPlanUpdates", err.Description, err.Number)
    Resume CleanExit

End Sub

Function logClick(modName As String, formName As String, Optional dataTag0 = "")
    On Error GoTo Err_Handler

    ' 1. Check if analytics are enabled
    If Nz(DLookup("paramVal", "tblDBinfoBE", "parameter = 'recordAnalytics'"), "False") = "False" Then Exit Function

    Dim conn As ADODB.Connection
    Set conn = CurrentProject.Connection
    
    Dim strSQL As String
    
    ' 2. Build the SQL string for Access
    ' IMPORTANT: Access ADO requires # for dates and '' for escaped quotes
    strSQL = "INSERT INTO tblAnalytics ([module], [form], [username], [dateused], [datatag0], [datatag1]) " & _
             "VALUES (" & _
             "'" & Replace(modName, "'", "''") & "', " & _
             "'" & Replace(formName, "'", "''") & "', " & _
             "'" & Environ("username") & "', " & _
             "#" & Format(Now(), "yyyy-mm-dd hh:nn:ss") & "#, " & _
             "'" & Replace(Nz(dataTag0, ""), "'", "''") & "', " & _
             "'SP" & Nz(TempVars!dbVersion, "") & "')"


    ' 3. Execute
    conn.Execute strSQL

CleanUp:
    Set conn = Nothing
    Exit Function

Err_Handler:
    Call handleError("wdbGlobalFunctions", "logClick", err.Description, err.Number)
    Resume CleanUp
End Function

Public Function StrQuoteReplace(strValue)
On Error GoTo Err_Handler

StrQuoteReplace = Replace(Nz(strValue, ""), "'", "''")

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "StrQuoteReplace", err.Description, err.Number)
End Function