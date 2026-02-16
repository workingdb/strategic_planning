Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdOpenLink_Click()
    Dim url As String
    url = Trim(Nz(Me.directLink, ""))
 
    If Len(url) = 0 Then Exit Sub
 
    CreateObject("WScript.Shell").Run _
        "cmd /c start """" """ & url & """", 0, False
End Sub
