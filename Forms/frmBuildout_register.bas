Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cboOEM_AfterUpdate()
On Error GoTo ErrHandler

    Me.cboModel = Null
    Me.cboProgramCode = Null
    Me!programId = Null
    Me.cboModel.Requery
    
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub
 
Private Sub cboModel_AfterUpdate()
On Error GoTo ErrHandler

    Me.cboProgramCode = Null
    Me!programId = Null
    Me.cboProgramCode.Requery
     
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub cboProgramCode_AfterUpdate()
On Error GoTo ErrHandler
 
    Dim pid As Long
    pid = Nz(Me.cboProgramCode.value, 0) 'this is tblPrograms.ID because Bound Column=1
 
    If pid = 0 Then
        Me!programId = Null
        Me!OEM = Null
        Me!modelName = Null
        Me!modelCode = Null
        Exit Sub
    End If
 
    'If Control Source is programId, Access already set it.
    'This line is harmless and ensures it s set:
    Me!programId = pid
 
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub

Private Sub cmdSaveBuildout_Click()
On Error GoTo ErrHandler
 
    Dim missing As String
 
    If IsNull(Me.txtBuildoutDate) Then missing = missing & vbCrLf & ""  Buildout Date"
    If Len(Trim$(Nz(Me.cboOEM, ""))) = 0 Then missing = missing & vbCrLf & ""  OEM"
    If Len(Trim$(Nz(Me.cboModel, ""))) = 0 Then missing = missing & vbCrLf & ""  Model Name"
    If Nz(Me!programId, 0) = 0 Then missing = missing & vbCrLf & ""  Program Code"
 
    If Len(missing) > 0 Then
        MsgBox "Please complete the required fields before saving:" & missing, vbExclamation
        Exit Sub
    End If
 
    DoCmd.RunCommand acCmdSaveRecord
 
    MsgBox "Saved. Buildout Record ID: " & Me!RecordID, vbInformation
 
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub
 
Private Sub cmdSaveClose_Click()
On Error GoTo ErrHandler
 
    Dim missing As String
 
    If IsNull(Me.txtBuildoutDate) Then missing = missing & vbCrLf & ""  Buildout Date"
    If Len(Trim$(Nz(Me.cboOEM, ""))) = 0 Then missing = missing & vbCrLf & ""  OEM"
    If Len(Trim$(Nz(Me.cboModel, ""))) = 0 Then missing = missing & vbCrLf & ""  Model Name"
    If Nz(Me!programId, 0) = 0 Then missing = missing & vbCrLf & ""  Program Code"
 
    If Len(missing) > 0 Then
        MsgBox "Please complete the required fields before saving:" & missing, vbExclamation
        Exit Sub
    End If
 
    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.CLOSE acForm, Me.name
 
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub
 
Private Sub Form_BeforeInsert(Cancel As Integer)
On Error GoTo ErrHandler
 
    Me.createdDate = Now()
    Me.createdBy = Environ("USERNAME")
    'Me.responsibleUser = Environ("USERNAME")
    Me.receivedDate = Now()
 
Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_BeforeInsert", Err.Description, Err.Number)
End Sub

 
Private Sub cmdCancelBuildout_Click()
On Error GoTo ErrHandler
 
    If MsgBox("Cancel this buildout? Any unsaved changes will be lost.", _
              vbQuestion + vbYesNo, "Cancel Buildout") <> vbYes Then
        
        Exit Sub
    End If
 
    'If record is new and unsaved
    If Me.newRecord Then
        Me.Undo
        Exit Sub
    End If
 
    'If record already exists (you generate recordId early), delete it
        DoCmd.RunCommand acCmdDeleteRecord
        
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.Description, Err.Number)
End Sub
