Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDetails_Click()
    DoCmd.OpenForm "frmBuildOutDetails", , , "[registerId]=" & Me.[RecordID]
End Sub

Private Sub cmdExposureInput_Click()
    DoCmd.OpenForm "frmBuildoutFGExposure", , , "[registerId]=" & Me.RecordID
End Sub
