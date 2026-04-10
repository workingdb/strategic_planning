option compare database
option explicit

global const sw_hide = 0
global const sw_shownormal = 1
global const sw_showminimized = 2
global const sw_showmaximized = 3
global const sw_restore = 9

private type rect
x1 as long
y1 as long
x2 as long
y2 as long
end type

private declare ptrsafe function getdesktopwindow lib "user32" () as long
private declare ptrsafe function getwindowrect lib "user32" (byval hwnd as long, r as rect) as long
public declare ptrsafe function iszoomed lib "user32" (byval hwnd as long) as long
private declare ptrsafe function movewindow lib "user32" alias "MoveWindow" (byval hwnd as long, byval x as long, byval y as long, byval dx as long, byval dy as long, byval frepaint as long) as long
private declare ptrsafe function showwindow lib "user32" (byval hwnd as long, byval ncmdshow as long) as long

dim appx as long, appy as long, apptop as long, appleft as long, winrect as rect

sub maximizeaccess()
on error goto err_handler

dim h as long
dim r as rect

on error resume next

h = application.hwndaccessapp
'if maximised, restore
if (iszoomed(h) = false) then showwindow h, sw_showmaximized

exit sub
err_handler:
    call handleerror("modAdminFunctions", "maximizeAccess", err.description, err.number)
end sub

public sub handleerror(modname as string, activecon as string, errdesc as string, errnum as long, optional datatag as string = "")
on error resume next

if (currentproject.path <> "C:\workingdb") then
    msgbox errdesc, vbinformation, "Error Code: " & errnum
    exit sub
end if

select case errnum
    case 70
        msgbox "Permissions Error - Check if the file is already in use.", vbinformation, "Error Code: " & errnum
    case 53
        msgbox "File Not Found", vbinformation, "Error Code: " & errnum
        exit sub
    case 3011
        msgbox "Looks like I'm having issues connecting to SharePoint. Please reopen when you can", vbinformation, "Error Code: " & errnum
    case 490, 52, 75
        msgbox "I cannot open this file or location - check if it has been moved or deleted. Or - you do not have proper access to this location", vbinformation, "Error Code: " & errnum
        exit sub
    case 3022
        msgbox "A record with this key already exists. I cannot create another!", vbinformation, "Error Code: " & errnum
    case 3167
        msgbox "Looks like you already deleted that record", vbinformation, "Error Code: " & errnum
        exit sub
    case 94
        msgbox "Hmm. Looks like something is missing. Check for an empty field", vbinformation, "Error Code: " & errnum
    case 3151
        msgbox "You're not connected to Oracle. Just FYI, Oracle connection does not work outside of VMWare.", vbinformation, "Error Code: " & errnum
        exit sub
    case 429
        if modname = "frmCatiaMacros" then
            msgbox "Looks like Catia isn't open", vbinformation, "Error Code: " & errnum
            exit sub
        else
            msgbox errdesc, vbinformation, "Error Code: " & errnum
        end if
    case 3343
        msgbox "Error. Please re-open WorkingDB to reset.", vbcritical, "Error Code: " & errnum
    case else
        msgbox errdesc, vbinformation, "Error Code: " & errnum
end select

dim strsql as string

modname = replace(nz(modname, ""), "'", "''")
errdesc = replace(nz(errdesc, ""), "'", "''")
errnum = replace(nz(errnum, ""), "'", "''")
datatag = replace(nz(datatag, ""), "'", "''")

strsql = "INSERT INTO tblErrorLog([User],Form,Active_Control,Error_Date,Error_Description,Error_Number,databaseVersion,dataTag0) VALUES ('" & _
 environ("username") & "','" & modname & "','" & nz(activecon, "") & "',#" & now & "#,'" & errdesc & "'," & errnum & ",'SP:" & nz(tempvars!dbversion, "") & "','" & datatag & "')"

dim conn as adodb.connection
set conn = currentproject.connection

conn.execute strsql

set conn = nothing

end sub

sub sizeaccess(byval dx as long, byval dy as long)
on error goto err_handler
'set size of access and center on desktop

dim h as long
dim r as rect

on error resume next

h = application.hwndaccessapp
'if maximised, restore
if (iszoomed(h)) then showwindow h, sw_restore
'
'get available desktop size
getwindowrect getdesktopwindow(), r
if ((r.x2 - r.x1) - dx) < 0 or ((r.y2 - r.y1) - dy) < 0 then
'desktop smaller than requested size
'so size to desktop
movewindow h, r.x1, r.y1, r.x2, r.y2, true
else
'adjust to requested size and center
movewindow h, _
r.x1 + ((r.x2 - r.x1) - dx) \ 2, _
r.y1 + ((r.y2 - r.y1) - dy) \ 2, _
dx, dy, true
end if

exit sub
err_handler:
    call handleerror("modAdminFunctions", "SizeAccess", err.description, err.number)
end sub

function grabversion() as string
on error goto err_handler

dim db as database
set db = currentdb()
dim rs1 as recordset
set rs1 = db.openrecordset("SELECT releaseVal FROM tblDBinfo WHERE recordId = 1", dbopensnapshot)
grabversion = rs1!releaseval
rs1.close: set rs1 = nothing
set db = nothing

exit function
err_handler:
    call handleerror("modAdminFunctions", "grabVersion", err.description, err.number)
end function
