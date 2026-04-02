Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
 
Private Sub capacityResults_AfterUpdate()
On Error GoTo Err_Handler

Select Case Me.ActiveControl
    Case 0 'no response
        Me.Filter = "recordId IN (SELECT requestId FROM tblCapacityRequest_partnumbers WHERE capacityResults is null)"
        Me.FilterOn = True
    Case 9999 'all
        Me.FilterOn = False
    Case Else 'specific based on ID
        Me.Filter = "recordId IN (SELECT requestId FROM tblCapacityRequest_partnumbers WHERE capacityResults = " & Me.capacityResults & ")"
        Me.FilterOn = True
End Select

Me.partNumFilt = ""

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub Customer_Label_Click()
    On Error GoTo Err_Handler
    Me.Customer.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub EOP_Label_Click()
    On Error GoTo Err_Handler
    Me.EOP.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub partNumFilt_AfterUpdate()
On Error GoTo Err_Handler

If IsNull(Me.partNumFilt) Then
    Me.FilterOn = False
Else
    Me.Filter = "recordId IN (SELECT requestId FROM tblCapacityRequest_partnumbers WHERE partNumber = '" & Me.partNumFilt & "')"
    Me.FilterOn = True
    Me.capacityResults = 9999
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub Program_Label_Click()
    On Error GoTo Err_Handler
    Me.Program.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub RecordID_Label_Click()
    On Error GoTo Err_Handler
    Me.recordId.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub Request_Date_Label_Click()
    On Error GoTo Err_Handler
    Me.Request_Date.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub Request_Type_Label_Click()
On Error GoTo Err_Handler
    Me.Request_Type.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub Requestor_Label_Click()
    On Error GoTo Err_Handler
    Me.Requestor.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub SOP_Label_Click()
    On Error GoTo Err_Handler
    Me.SOP.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub newRequest_Click()
    On Error GoTo Err_Handler
    DoCmd.OpenForm "frmCapacityRequestDetails", , , , acFormAdd
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
 
Private Sub openDetails_Click()
    On Error GoTo ErrHandler
 
    If IsNull(Me.recordId) Then Exit Sub
 
    DoCmd.OpenForm "frmCapacityRequestDetails", acNormal, , "recordId = " & Me.recordId
 
Exit Sub
ErrHandler:
    MsgBox "Open Details error " & err.Number & ":" & vbCrLf & err.Description, vbExclamation
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.Filter = "recordId IN (SELECT requestId FROM tblCapacityRequest_partnumbers WHERE capacityResults is null)"
Me.FilterOn = True

Me.capacityResults = 0

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", err.Description, err.Numbe)
End Sub
