Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdOpenLink_Click()
 
    Dim url As String
    url = Trim(Nz(Me.directLink, ""))
 
    If Len(url) = 0 Then
        MsgBox "No link found.", vbInformation
        Exit Sub
    End If
 
    'Force Windows to open in default browser (Edge/Chrome/etc.)
    CreateObject("WScript.Shell").Run _
        "cmd /c start """" """ & url & """", 0, False
 
End Sub
