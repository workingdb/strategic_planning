Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
    On Error GoTo Err_Handler

    ' --- INITIAL SETUP ---
    TempVars.Add "loadAmount", 0
    TempVars.Add "loadWd", 8160
    TempVars.Add "dbVersion", grabVersion()
    Me.lblFrozen.Visible = False
    Call setSplashLoading("Setting up app stuff...")
    Me.lblVersion.Caption = TempVars!dbVersion
    Me.lblVersion.Visible = True

    SizeAccess 280, 280
    Me.Move -2600, -1000
    
        'make sure driver reference for SQL Server is OK
    Call RelinkSQLTables

    ' Use the ADODB version of logClick we created earlier
    Call logClick("Form_Load", Me.name)

    ' Splash Image Logic (Keeping DLookup for local settings)
    Me.Picture = "\\data\mdbdata\WorkingDB\Pictures\Splash\splash" & randomNumber(0, DLookup("splashCount", "tblDBinfoBE", "ID = 1")) & ".png"
    
    On Error Resume Next
    Me.imgUser.Picture = "\\data\mdbdata\WorkingDB\Pictures\Avatars\" & Environ("username") & ".png"
    On Error GoTo Err_Handler
    
    DoEvents
    Form_frmSplash.SetFocus
    DoEvents

    ' --- RIBBON & SHORTCUTS ---
    If CommandBars("Ribbon").height > 100 Then CommandBars.ExecuteMso "MinimizeRibbon"
    DoCmd.ShowToolbar "Ribbon", acToolbarNo

    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile "\\data\mdbdata\WorkingDB\Batch\Working DB.lnk", "\\homes\data\" & Environ("username") & "\Desktop\Working DB.lnk"
    fso.CopyFile "\\data\mdbdata\WorkingDB\build\workingdb_ghost\WorkingDB_ghost.accde", "C:\workingdb\WorkingDB_ghost.accde"
    openPath "\\data\mdbdata\WorkingDB\build\workingdb_commands\openGhost.vbs"
    On Error GoTo Err_Handler

    ' --- DATABASE LOGIC (ADODB CONVERSION) ---
    Call setSplashLoading("Doing some digging on you...")
    
    Dim conn As ADODB.Connection: Set conn = CurrentProject.Connection
    Dim rsUser As New ADODB.Recordset, rsPerm As New ADODB.Recordset

    rsUser.Open "SELECT * FROM tblUserSettings WHERE [username] = '" & Environ("username") & "'", conn, adOpenKeyset, adLockOptimistic
    If rsUser.EOF Then
        MsgBox "You need to have an account in WorkingDB to access this.", vbOKOnly, "Welcome"
        Application.Quit
    End If

    rsPerm.Open "SELECT * FROM tblPermissions WHERE [User] = '" & Environ("username") & "'", conn, adOpenKeyset, adLockOptimistic
    
    If rsPerm.EOF Then
        MsgBox "You need to have an account in WorkingDB to access this.", vbOKOnly, "Welcome"
        Application.Quit
    End If

    ' --- SET TEMPVARS ---
    TempVars.Add "dept", Nz(rsPerm!Dept, "")
    TempVars.Add "org", Nz(rsPerm!org, 4)
    TempVars.Add "smallScreen", Nz(rsUser!smallScreenMode, "False")

    ' --- THEME LOGIC ---
    If Nz(rsUser!themeId, 0) <> 0 Then
        Dim rsTheme As New ADODB.Recordset
        rsTheme.Open "SELECT * FROM tblTheme WHERE recordId = " & rsUser!themeId, conn, adOpenForwardOnly, adLockReadOnly
        If Not rsTheme.EOF Then
            TempVars.Add "themeMode", IIf(rsTheme!darkMode, "Dark", "Light")
            TempVars.Add "themePrimary", CStr(rsTheme!primaryColor)
            TempVars.Add "themeSecondary", CStr(rsTheme!secondaryColor)
            TempVars.Add "themeAccent", CStr(rsTheme!accentColor)
            TempVars.Add "themeColorLevels", CStr(rsTheme!colorLevels)
        End If
        rsTheme.Close
    End If

    ' --- FINALIZE STARTUP ---
    Call setSplashLoading("Running daily checks...")
    Call grabJoke
    
    DoCmd.OpenForm "DASHBOARD"
    Forms!DASHBOARD.Visible = False
    
    DoCmd.Close acForm, Me.name
    Call maximizeAccess
    Forms!DASHBOARD.Visible = True
    DoCmd.Maximize

CleanUp:
    If rsUser.State = adStateOpen Then rsUser.Close
    If rsPerm.State = adStateOpen Then rsPerm.Close
    Set rsUser = Nothing: Set rsPerm = Nothing
    Set conn = Nothing
    Exit Sub

Err_Handler:
    Call handleError(Me.name, "Form_Load", err.Description, err.Number)
    Resume CleanUp
End Sub

Function grabJoke()
On Error GoTo Err_Handler

Dim Joke As String
Joke = Nz(DLookup("[factText]", "tblFacts", "[factDate] = #" & Date & "#"))

TempVars.Add "joke", Joke

Exit Function
Err_Handler:
    Call handleError(Me.name, "grabJoke", err.Description, err.Number)
End Function
