Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mPendingSourcingRouting As Boolean

Private Sub Form_AfterInsert()
    On Error GoTo ErrHandler
 
    'New request created: queue routing once
    mPendingSourcingRouting = True
    Me.TimerInterval = 50
 
ExitHere:
    Exit Sub
ErrHandler:
    MsgBox "frmCapacityRequestDetails AfterInsert error: " & Err.Number & vbCrLf & Err.Description, vbExclamation
    Resume ExitHere
End Sub
 
Private Sub Form_Timer()
    On Error GoTo ErrHandler
 
    'Stop timer immediately
    Me.TimerInterval = 0
 
    If mPendingSourcingRouting Then
        mPendingSourcingRouting = False
 
        'Run module gates + actions (U6 + Purchased etc.)
        HandleSourcingRouting Me
    End If
 
ExitHere:
    Exit Sub
ErrHandler:
    Me.TimerInterval = 0
    mPendingSourcingRouting = False
    MsgBox "frmCapacityRequestDetails Timer error: " & Err.Number & vbCrLf & Err.Description, vbExclamation
    Resume ExitHere
End Sub

Private Sub Capacity_Results_AfterUpdate()
    If Me.Dirty Then Me.Dirty = False   ' forces save
    Call NotifyCapacityResultIfNeeded(CLng(Me.RecordID))
End Sub

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
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub


Private Sub mailReport_Click()
On Error GoTo Err_Handler

    DoCmd.SendObject acReport, "Capacity Confirmation", "", "", "", "", "", "", True, ""

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub
