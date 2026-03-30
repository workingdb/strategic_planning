Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function trackUpdate()
On Error GoTo Err_Handler

If IsNull(Me.RecordID) Then Exit Function
Call registerStratPlanUpdates("tblCapacityRequest_partnumbers", Me.RecordID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.RecordID, Me.name)

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

    On Error Resume Next
    DoCmd.GoToControl Screen.PreviousControl.name
    err.Clear
    If (Not Form.newRecord) Then
        DoCmd.RunCommand acCmdDeleteRecord
    End If
    If (Form.newRecord And Form.Dirty) Then
        DoCmd.RunCommand acCmdUndo
    End If
    If (MacroError <> 0) Then
        MsgBox MacroError.Description, vbOKOnly, ""
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


Private Sub partNumber_AfterUpdate()
On Error GoTo Err_Handler

If Nz(Me.partNumber, "") = "" Then Exit Sub

Call trackUpdate

'find current unit
Dim db As Database
Set db = CurrentDb()
Dim invId, currentUnit As String, rsCat As Recordset
invId = Nz(idNAM(Me.partNumber, "NAM"), "")

currentUnit = ""

If invId <> "" Then
    Set rsCat = db.OpenRecordset("SELECT SEGMENT1 FROM INV_MTL_ITEM_CATEGORIES LEFT JOIN APPS_MTL_CATEGORIES_VL ON INV_MTL_ITEM_CATEGORIES.CATEGORY_ID = APPS_MTL_CATEGORIES_VL.CATEGORY_ID " & _
    "GROUP BY INV_MTL_ITEM_CATEGORIES.INVENTORY_ITEM_ID, APPS_MTL_CATEGORIES_VL.SEGMENT1, APPS_MTL_CATEGORIES_VL.STRUCTURE_ID HAVING STRUCTURE_ID = 101 AND [INVENTORY_ITEM_ID] = " & invId, dbOpenSnapshot)
    If rsCat.RecordCount > 0 Then currentUnit = Nz(rsCat!SEGMENT1, "")

    rsCat.Close
    Set rsCat = Nothing
End If

If currentUnit <> "" Then
    Dim unitId
    unitId = Nz(DLookup("recordId", "tblUnits", "unitName = '" & currentUnit & "'"), 0)
    Me.unitId = unitId
End If

Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Private Sub remove_Click()
On Error GoTo Err_Handler

If MsgBox("Are you sure you want to delete this?", vbYesNo, "Please confirm") = vbYes Then
    If Nz(Me.RecordID, 0) <> 0 Then Call registerStratPlanUpdates("tblCapacityRequestDetail_partnumbers", Me.RecordID, "Part Number", Nz(Me.partNumber, ""), "Deleted", Me.RecordID, Me.name)
    dbExecute ("DELETE FROM tblCapacityRequest_partnumbers WHERE [recordId] = " & Me.RecordID)
    Me.Requery
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub
