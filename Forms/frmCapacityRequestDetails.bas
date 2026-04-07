Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function trackUpdate()
On Error GoTo Err_Handler

If IsNull(Me.RecordID) Then Exit Function
Call registerStratPlanUpdates("tblCapacityRequests", Me.RecordID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.RecordID, Me.name)

Exit Function
Err_Handler:
    Call handleError(Me.name, "trackUpdate", err.Description, err.Number)
End Function

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", err.Description, err.Number)
End Sub

Private Sub Find_Click()
On Error GoTo Err_Handler

    On Error Resume Next
    DoCmd.GoToControl Screen.PreviousControl.name
    err.Clear
    DoCmd.RunCommand acCmdFind
    If (MacroError <> 0) Then
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub New_Click()
On Error GoTo Err_Handler

    On Error Resume Next
    DoCmd.GoToRecord , "", acNewRec
    If (MacroError <> 0) Then
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub requestType_AfterUpdate()
On Error GoTo Err_Handler

Call trackUpdate

Me.surveyPartCount.Visible = Me.requestType = 2

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub Trash_Click()
On Error GoTo Err_Handler

If MsgBox("Are you sure you want to delete this request?", vbYesNo, "Please confirm") = vbYes Then
    If Nz(Me.RecordID, 0) <> 0 Then Call registerStratPlanUpdates("tblCapacityRequestDetails", Me.RecordID, "Request", "", "Deleted", Me.RecordID, Me.name)
    dbExecute ("DELETE FROM tblCapacityRequests WHERE [recordId] = " & Me.RecordID)
    TempVars.Add "reqCapDelete", "True"
    DoCmd.Close
    If CurrentProject.AllForms("frmCapacityRequestTracker").IsLoaded Then Form_frmCapacityRequestTracker.Requery
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub copy_Click()
On Error GoTo Err_Handler

    On Error Resume Next
    DoCmd.RunCommand acCmdSelectRecord
    If (MacroError = 0) Then
        DoCmd.RunCommand acCmdCopy
    End If
    If (MacroError = 0) Then
        DoCmd.RunCommand acCmdRecordsGoToNew
    End If
    If (MacroError = 0) Then
        DoCmd.RunCommand acCmdSelectRecord
    End If
    If (MacroError = 0) Then
        DoCmd.RunCommand acCmdPaste
    End If
    If (MacroError <> 0) Then
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub mailReport_Click()
On Error GoTo Err_Handler

Dim partNums As String
partNums = findCapReqPNs(Me.RecordID, True)

Dim pnSplit() As String, item, partNumFinal As String
pnSplit = Split(partNums, ",")
partNumFinal = ""

For Each item In pnSplit
    partNumFinal = partNumFinal & "PN: " & Split(item, "|")(0) & " - Response: " & Split(item, "|")(1)
Next item

Dim body As String
body = emailContentGen("Capacity Request Results", _
    Me.requestType.column(1) & " Results", _
    "Notes: " & Replace(Me.Notes, ",", ";"), _
     partNumFinal, _
    "Requested: " & CStr(Date) & " by: " & Me.Requestor.column(1), _
    "Vehicle: " & Me.Program.column(1), _
    "Program: " & Me.Program.column(0))
Call registerStratPlanUpdates("tblCapacityRequestDetails", Me.RecordID, "Results", "", "Results Sent to Requestor", Me.RecordID, Me.name)
If sendNotification(Me.Requestor.column(2), 6, 2, "Capacity Request Results", body) Then
    Call snackBox("success", "Well Done!", Me.Requestor.column(2) & " Notified!", Me.name)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
