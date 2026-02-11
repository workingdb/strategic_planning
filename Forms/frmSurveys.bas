Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.Description, Err.Number)
End Sub

Private Sub Find_Click()
On Error GoTo Err_Handler

    On Error Resume Next
    DoCmd.GoToControl Screen.PreviousControl.name
    Err.Clear
    DoCmd.RunCommand acCmdFind
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub New_Click()
On Error GoTo Err_Handler

    On Error Resume Next
    DoCmd.GoToRecord , "", acNewRec
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub Trash_Click()
On Error GoTo Err_Handler

    On Error Resume Next
    DoCmd.GoToControl Screen.PreviousControl.name
    Err.Clear
    If (Not Form.newRecord) Then
        DoCmd.RunCommand acCmdDeleteRecord
    End If
    If (Form.newRecord And Not Form.Dirty) Then
        Beep
    End If
    If (Form.newRecord And Form.Dirty) Then
        DoCmd.RunCommand acCmdUndo
    End If
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub copyRecord_Click()
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
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub
