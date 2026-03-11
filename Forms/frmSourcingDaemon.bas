Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
 
Private mStartTime As Date
 
Private Sub Form_Load()
    On Error GoTo ErrHandler
 
    mStartTime = Now()
 
    'Run an immediate scan on open
    Call ScanAndRoutePendingRequests
 
    'Poll while open (every 60 seconds; adjust as desired)
    Me.TimerInterval = 60000
 
    Exit Sub
ErrHandler:
    'If something goes wrong, stop timer so it doesn t spam errors
    Me.TimerInterval = 0
End Sub
 
Private Sub Form_Timer()
    On Error GoTo ErrHandler
 
    'Stop after 2 hours per user session
    If DateDiff("n", mStartTime, Now()) >= 120 Then
        Me.TimerInterval = 0
        DoCmd.Close acForm, Me.name, acSaveNo
        Exit Sub
    End If
 
    Call ScanAndRoutePendingRequests
    Exit Sub
 
ErrHandler:
    'Fail-safe: stop timer if errors repeat
    Me.TimerInterval = 0
End Sub
