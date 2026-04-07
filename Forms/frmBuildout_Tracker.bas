Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdDetails_Click()
    DoCmd.OpenForm "frmBuildout_details", , , "[registerId]=" & Me.[RecordID]
End Sub

Private Sub cmdExposureInput_Click()
    DoCmd.OpenForm "frmBuildout_exposure", , , "[registerId]=" & Me.RecordID
End Sub
