Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function trackUpdate()
On Error GoTo Err_Handler

If IsNull(Me.recordId) Then Exit Function
Call registerStratPlanUpdates("tblCapacityRequests", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.recordId, Me.name)

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

Private Sub Trash_Click()
On Error GoTo Err_Handler

If MsgBox("Are you sure you want to delete this request?", vbYesNo, "Please confirm") = vbYes Then
    If Nz(Me.recordId, 0) <> 0 Then Call registerStratPlanUpdates("tblCapacityRequestDetails", Me.recordId, "Request", "", "Deleted", Me.recordId, Me.name)
    dbExecute ("DELETE FROM tblCapacityRequests WHERE [recordId] = " & Me.recordId)
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

DoCmd.SendObject acReport, "Capacity Confirmation", "", "", "", "", "", "", True, ""

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
