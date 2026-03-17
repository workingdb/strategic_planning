Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cboOEM_AfterUpdate()
    Me.cboModel = Null
    Me.cboProgramCode = Null
    Me!programId = Null
    Me.cboModel.Requery
    
End Sub
 
Private Sub cboModel_AfterUpdate()
    Me.cboProgramCode = Null
    Me!programId = Null
    Me.cboProgramCode.Requery
     
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
 
ErrHandler:
    MsgBox "cboProgramCode_AfterUpdate error: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub cmdSaveBuildout_Click()
 
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
 
End Sub
 
Private Sub cmdSaveClose_Click()
 
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
    DoCmd.Close acForm, Me.name
 
End Sub
 
Private Sub Form_BeforeInsert(Cancel As Integer)
 
    Me.createdDate = Now()
    Me.createdBy = Environ("USERNAME")
    'Me.responsibleUser = Environ("USERNAME")
    Me.receivedDate = Now()
 
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
ErrHandler:     MsgBox "Error: " & Err.Number & " - " & Err.Description, vbExclamation
    
End Sub
