Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function applyFilter(parameter As String)

Dim db As Database
Set db = CurrentDb()

Dim qdf As QueryDef

Set qdf = db.QueryDefs("frmCapacityRequestTracker_PT")

If parameter = "" Then
    qdf.sql = Split(qdf.sql, "c.ID")(0) & " c.ID;"
Else
    qdf.sql = Split(qdf.sql, "c.ID")(0) & " c.ID WHERE EXISTS (SELECT 1 From tblCapacityRequest_partnumbers As cp WHERE cp.requestId = cr.recordId AND " & parameter & ");"
End If

db.QueryDefs.refresh

Set qdf = Nothing
Set db = Nothing

Me.Requery

End Function
 
Private Sub capacityResults_AfterUpdate()
On Error GoTo Err_Handler

Select Case Me.ActiveControl
    Case 0 'no response
        applyFilter ("cp.capacityResults is null")
    Case 9999 'all
        applyFilter ("")
    Case Else 'specific based on ID
        applyFilter ("cp.capacityResults = " & Me.capacityResults)
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

If IsNull(Me.partNumFilt) Then 'see all
    applyFilter ("")
Else 'filter by part number
    applyFilter ("cp.partNumber = '" & Me.partNumFilt & "'")
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
    Me.RecordID.SetFocus
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
 
    If IsNull(Me.RecordID) Then Exit Sub
 
    DoCmd.OpenForm "frmCapacityRequestDetails", acNormal, , "recordId = " & Me.RecordID
 
Exit Sub
ErrHandler:
    MsgBox "Open Details error " & err.Number & ":" & vbCrLf & err.Description, vbExclamation
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

applyFilter ("cp.capacityResults is null")

Me.capacityResults = 0

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", err.Description, err.Numbe)
End Sub
