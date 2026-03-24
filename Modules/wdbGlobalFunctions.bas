Option Compare Database
Option Explicit

Function setSplashLoading(label As String)
On Error GoTo Err_Handler

If IsNull(TempVars!loadAmount) Then Exit Function
TempVars.Add "loadAmount", TempVars!loadAmount + 1
Form_frmSplash.lnLoading.Width = (TempVars!loadAmount / 12) * TempVars!loadWd
Form_frmSplash.lblLoading.Caption = label
Form_frmSplash.Repaint

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "setSplashLoading", Err.Description, Err.Number)
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
        userData = Nz(rs.Fields(0).value, "")
    Else
        userData = ""
    End If

CleanUp:
    If rs.State = adStateOpen Then rs.CLOSE
    Set rs = Nothing
    Set conn = Nothing
    Exit Function

Err_Handler:
    Call handleError("wdbGlobalFunctions", "userData", Err.Description, Err.Number)
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
    Call handleError("wdbGlobalFunctions", "dbExecute", Err.Description, Err.Number, sql)
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
    Call handleError("wdbGlobalFunctions", "registerStratPlanUpdates", Err.Description, Err.Number)
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
    Call handleError("wdbGlobalFunctions", "logClick", Err.Description, Err.Number)
    Resume CleanUp
End Function

Public Function StrQuoteReplace(strValue)
On Error GoTo Err_Handler

StrQuoteReplace = Replace(Nz(strValue, ""), "'", "''")

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "StrQuoteReplace", Err.Description, Err.Number)
End Function