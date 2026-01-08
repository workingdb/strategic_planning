Option Compare Database
Option Explicit

Function userData(data, Optional specificUser As String = "") As String
On Error GoTo Err_Handler

If specificUser = "" Then specificUser = Environ("username")

userData = Nz(DLookup("[" & data & "]", "[tblPermissions]", "[User] = '" & specificUser & "'"))

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "replaceDriveLetters", Err.DESCRIPTION, Err.number)
End Function

Public Function setTheme(setForm As Form)
On Error Resume Next

Dim scalarBack As Double, scalarFront As Double, darkMode As Boolean
Dim backBase As Long, foreBase As Long, colorLevels(4), backSecondary As Long, btnXback As Long, btnXbackShade As Long

'IF NO THEME SET, APPLY DEFAULT THEME (for Dev mode)
If Nz(TempVars!themePrimary, "") = "" Then
    TempVars.Add "themePrimary", 3355443
    TempVars.Add "themeSecondary", 0
    TempVars.Add "themeMode", "Dark"
    TempVars.Add "themeColorLevels", "1.3,1.6,1.9,2.2"
End If

darkMode = TempVars!themeMode = "Dark"

If darkMode Then
    foreBase = 16777215
    btnXback = 4342397
    scalarBack = 1.3
    scalarFront = 0.9
Else
    foreBase = 657930
    btnXback = 8947896
    scalarBack = 1.1
    scalarFront = 0.3
End If

backBase = CLng(TempVars!themePrimary)
backSecondary = CLng(TempVars!themeSecondary)

Dim colorLevArr() As String
colorLevArr = Split(TempVars!themeColorLevels, ",")

If backSecondary <> 0 Then
    colorLevels(0) = backBase
    colorLevels(1) = shadeColor(backSecondary, CDbl(colorLevArr(0)))
    colorLevels(2) = shadeColor(backBase, CDbl(colorLevArr(1)))
    colorLevels(3) = shadeColor(backSecondary, CDbl(colorLevArr(2)))
    colorLevels(4) = shadeColor(backBase, CDbl(colorLevArr(3)))
Else
    colorLevels(0) = backBase
    colorLevels(1) = shadeColor(backBase, CDbl(colorLevArr(0)))
    colorLevels(2) = shadeColor(backBase, CDbl(colorLevArr(1)))
    colorLevels(3) = shadeColor(backBase, CDbl(colorLevArr(2)))
    colorLevels(4) = shadeColor(backBase, CDbl(colorLevArr(3)))
End If

setForm.FormHeader.BackColor = colorLevels(findColorLevel(setForm.FormHeader.Tag))
setForm.Detail.BackColor = colorLevels(findColorLevel(setForm.Detail.Tag))
If Len(setForm.Detail.Tag) = 4 Then
    setForm.Detail.AlternateBackColor = colorLevels(findColorLevel(setForm.Detail.Tag) + 1)
Else
    setForm.Detail.AlternateBackColor = setForm.Detail.BackColor
End If

setForm.FormFooter.BackColor = colorLevels(findColorLevel(setForm.FormFooter.Tag))

'assuming form parts don't use tags for other uses

Dim ctl As Control, eachBtn As CommandButton
Dim classColor As String, fadeBack, fadeFore
Dim Level
Dim backCol As Long, levFore As Double
Dim disFore As Double
Dim foreLevInt As Long, maxLev As Long

For Each ctl In setForm.Controls
    If ctl.Tag Like "*.L#*" Then
        Level = findColorLevel(ctl.Tag)
        backCol = colorLevels(Level)
    Else
        GoTo nextControl
    End If
    foreLevInt = Level
    If foreLevInt > 3 Then foreLevInt = 3
    maxLev = Level + 1
    If maxLev > 4 Then maxLev = 4
    
    If darkMode Then
        foreLevInt = Level
        If foreLevInt > 3 Then foreLevInt = 3
        levFore = (1 / colorLevArr(foreLevInt)) + 0.2
        disFore = 1.4 - levFore
    Else
        levFore = (colorLevArr(foreLevInt))
        disFore = 15 - levFore
    End If

    Select Case ctl.ControlType
        Case acCommandButton, acToggleButton 'OPTIONS: cardBtn.L#, cardBtnContrastBorder.L#, btn.L#
            If Not (ctl.Tag Like "*btn*") Then GoTo skipAhead0
            ctl.BackColor = backCol
            
            If (ctl.Picture = "") Then GoTo skipAhead0
            If darkMode Then
                If InStr(ctl.Picture, "\Core_theme_light\") Then ctl.Picture = Replace(ctl.Picture, "\Core_theme_light\", "\Core\")
            Else
                If InStr(ctl.Picture, "\Core\") Then ctl.Picture = Replace(ctl.Picture, "\Core\", "\Core_theme_light\")
            End If
            
skipAhead0:
            Select Case True
                Case ctl.Tag Like "*cardBtn.L#*"
                    ctl.BorderColor = backCol
                Case ctl.Tag Like "*cardBtnContrastBorder.L#*"
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = colorLevels(maxLev)
                Case ctl.Tag Like "*btn.L#*"
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = backCol
                    
                    'fade the colors
                    fadeBack = shadeColor(backCol, scalarBack)
                    fadeFore = shadeColor(foreBase, levFore - 0.2)
                    
                    ctl.ForeColor = foreBase
                    ctl.HoverColor = fadeBack
                    ctl.PressedColor = fadeBack
                    ctl.HoverForeColor = foreBase
                    ctl.PressedForeColor = foreBase
                Case ctl.Tag Like "*btnDis.L#*" 'for disabled look
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = backCol
                    
                    'fade the colors
                    fadeBack = shadeColor(backCol, scalarBack)
                    fadeFore = shadeColor(foreBase, disFore)
                    
                    ctl.ForeColor = fadeFore
                    ctl.HoverColor = fadeBack
                    ctl.PressedColor = fadeBack
                    ctl.HoverForeColor = fadeFore
                    ctl.PressedForeColor = fadeFore
                Case ctl.Tag Like "*btnDisContrastBorder.L#*" 'for disabled look
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = colorLevels(maxLev)
                    
                    'fade the colors
                    fadeBack = shadeColor(backCol, scalarBack)
                    fadeFore = shadeColor(foreBase, disFore)
                    
                    ctl.ForeColor = fadeFore
                    ctl.HoverColor = fadeBack
                    ctl.PressedColor = fadeBack
                    ctl.HoverForeColor = fadeFore
                    ctl.PressedForeColor = fadeFore
                Case ctl.Tag Like "*btnXdis.L#*" 'for disabled look
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = btnXback
                    ctl.ForeColor = foreBase
                    ctl.BackColor = btnXback
                    
                    'fade the colors
                    fadeBack = shadeColor(btnXback, scalarBack)
                    fadeFore = shadeColor(foreBase, disFore)
                    btnXbackShade = shadeColor(btnXback, (0.1 * Level) + scalarBack)
                    
                    ctl.BackColor = btnXbackShade
                    ctl.ForeColor = fadeFore
                    ctl.HoverColor = fadeBack
                    ctl.PressedColor = fadeBack
                    ctl.HoverForeColor = fadeFore
                    ctl.PressedForeColor = fadeFore
                Case ctl.Tag Like "*btnX.L#*"
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = btnXback
                    ctl.ForeColor = foreBase
                    'fade the colors
                    fadeBack = shadeColor(btnXback, scalarBack)
                    fadeFore = shadeColor(foreBase, scalarFront)
                    btnXbackShade = shadeColor(btnXback, (0.1 * Level) + scalarBack)
                    
                    ctl.BackColor = btnXbackShade
                    ctl.HoverColor = fadeBack
                    ctl.PressedColor = fadeBack
                    ctl.HoverForeColor = fadeFore
                    ctl.PressedForeColor = fadeFore
                Case ctl.Tag Like "*btnXcontrastBorder.L#*"
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = colorLevels(maxLev)
                    ctl.ForeColor = foreBase
                    'fade the colors
                    fadeBack = shadeColor(btnXback, scalarBack)
                    fadeFore = shadeColor(foreBase, scalarFront)
                    btnXbackShade = shadeColor(btnXback, (0.1 * Level) + scalarBack)
                    
                    ctl.BackColor = btnXbackShade
                    ctl.HoverColor = fadeBack
                    ctl.PressedColor = fadeBack
                    ctl.HoverForeColor = fadeFore
                    ctl.PressedForeColor = fadeFore
                Case ctl.Tag Like "*btnContrastBorder.L#*"
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = colorLevels(maxLev)
                    ctl.ForeColor = foreBase
                    
                    'fade the colors
                    fadeBack = shadeColor(backCol, scalarBack)
                    fadeFore = shadeColor(foreBase, scalarFront)
                    
                    ctl.HoverColor = fadeBack
                    ctl.PressedColor = fadeBack
                    ctl.HoverForeColor = fadeFore
                    ctl.PressedForeColor = fadeFore
            End Select
        Case acLabel
            Select Case True
               Case ctl.Tag Like "*lbl.L#*"
                   ctl.ForeColor = shadeColor(foreBase, levFore)
               Case ctl.Tag Like "*lbl_wBack.L#*"
                   ctl.ForeColor = shadeColor(foreBase, levFore)
                   ctl.BackColor = backCol
                   If ctl.BorderStyle <> 0 Then ctl.BorderColor = backCol
            End Select
        Case acTextBox, acComboBox 'OPTIONS: txt.L#, txtBackBorder.L#, txtContrastBorder.L#
            If ctl.Tag Like "*txt*" Then
                ctl.BackColor = backCol
                ctl.ForeColor = foreBase
            End If
            
            If ctl.FormatConditions.Count = 1 Then 'special case for null value conditional formatting. Typically this is used for placeholder values
                If ctl.FormatConditions.Item(0).Expression1 Like "*IsNull*" Then
                    ctl.FormatConditions.Item(0).BackColor = backCol
                    ctl.FormatConditions.Item(0).ForeColor = foreBase
                End If
            End If
            
            Select Case True
                Case ctl.Tag Like "*txtBackBorder*"
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = backCol
                Case ctl.Tag Like "*txtContrastBorder*"
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = colorLevels(maxLev)
                Case ctl.Tag Like "*txtTransFore*"
                    ctl.ForeColor = backCol
            End Select
        Case acRectangle, acSubform 'OPTIONS: cardBox.L#, cardBoxContrastBorder.L#
            If Not ctl.name Like "sfrm*" Then ctl.BackColor = backCol
            Select Case True
                Case ctl.Tag Like "*cardBox.L#*"
                    ctl.BorderColor = backCol
                Case ctl.Tag Like "*cardBoxContrastBorder.L#*"
                    ctl.BorderColor = colorLevels(Level + 1)
            End Select
        Case acTabCtl 'OPTIONS: tab.L#, tabContrastBorder.L#
            If ctl.Tag Like "*tab*" Then
                If Level = 0 Then
                    ctl.BackColor = colorLevels(Level + 0)
                    ctl.BorderColor = backCol
                    ctl.PressedColor = backCol
                    
                    'fade the colors
                    fadeBack = shadeColor(CLng(colorLevels(Level - 1)), scalarBack)
                    fadeFore = shadeColor(foreBase, levFore - 0.6)
                    
                    ctl.HoverColor = fadeBack
                    ctl.ForeColor = fadeFore
                    
                    ctl.HoverForeColor = foreBase
                    ctl.PressedForeColor = foreBase
                Else
                    ctl.BackColor = colorLevels(Level - 1)
                    ctl.BorderColor = backCol
                    ctl.PressedColor = backCol
                    
                    'fade the colors
                    fadeBack = shadeColor(CLng(colorLevels(Level - 1)), scalarBack)
                    fadeFore = shadeColor(foreBase, levFore)
                    
                    ctl.HoverColor = fadeBack
                    ctl.ForeColor = fadeFore
                    
                    ctl.HoverForeColor = foreBase
                    ctl.PressedForeColor = foreBase
                End If
            End If
            If ctl.Tag Like "*contrastBorder*" Then
                ctl.BorderColor = colorLevels(maxLev)
            End If
        Case acImage 'OPTIONS: pic.L#
            If ctl.Tag Like "*pic*" Then ctl.BackColor = backCol
    End Select
    
nextControl:
Next

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "setTheme", Err.DESCRIPTION, Err.number)
End Function

Function findColorLevel(tagText As String) As Long
On Error GoTo Err_Handler

findColorLevel = 0
If tagText = "" Then Exit Function

findColorLevel = Mid(tagText, InStr(tagText, ".L") + 2, 1)

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "setTheme", Err.DESCRIPTION, Err.number)
End Function

Function shadeColor(inputColor As Long, scalar As Double) As Long
On Error GoTo Err_Handler

Dim tempHex, ioR, ioG, ioB

tempHex = Hex(inputColor)

If tempHex = "0" Then tempHex = "111111"

If Len(tempHex) = 1 Then tempHex = "0" & tempHex
If Len(tempHex) = 2 Then tempHex = "0" & tempHex
If Len(tempHex) = 3 Then tempHex = "0" & tempHex
If Len(tempHex) = 4 Then tempHex = "0" & tempHex
If Len(tempHex) = 5 Then tempHex = "0" & tempHex

ioR = Val("&H" & Mid(tempHex, 5, 2)) * scalar
ioG = Val("&H" & Mid(tempHex, 3, 2)) * scalar
ioB = Val("&H" & Mid(tempHex, 1, 2)) * scalar

'Debug.Print ioR & " "; ioG & " " & ioB

If ioR > 255 Then ioR = 255
If ioG > 255 Then ioG = 255
If ioB > 255 Then ioB = 255

If ioR < 0 Then ioR = 0
If ioG < 0 Then ioG = 0
If ioB < 0 Then ioB = 0

shadeColor = RGB(ioR, ioG, ioB)

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "shadeColor", Err.DESCRIPTION, Err.number)
End Function

Function convertRecords()

Dim db As Database
Dim rs As Recordset

Set db = CurrentDb
Set rs = db.OpenRecordset("Capacity request Tracker")

Dim tempString As String
Dim newID As Long

Do While Not rs.EOF
    
    tempString = ""
    newID = 0
    
    If Nz(rs!Requestor, 0) = 0 Then GoTo nextrecord
    
    newID = CLng(DLookup("[Account Manager]", "Account Managers", "ID = " & rs!Requestor))
    
    rs.Edit
    rs!Requestor = newID
    rs.Update
    
nextrecord:
    rs.MoveNext
Loop

End Function