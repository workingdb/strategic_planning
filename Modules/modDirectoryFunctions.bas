option compare database
option explicit

declare ptrsafe function shellexecute lib "shell32.dll" alias "ShellExecuteA" (byval hwnd as long, byval lpoperation as string, byval lpfile as string, byval lpparameters as string, byval lpdirectory as string, byval lpnshowcmd as long) as long

public sub openpath(path)
on error goto err_handler

createobject("Shell.Application").open cvar(path)

exit sub
err_handler:
    call handleerror("modDirectoryFunctions", "openPath", err.description, err.number)
end sub

function replacedriveletters(linkinput) as string
on error goto err_handler

replacedriveletters = linkinput

replacedriveletters = replace(replacedriveletters, "N:\", "\\ncm-fs2\data\Department\")
replacedriveletters = replace(replacedriveletters, "T:\", "\\design\data\")
replacedriveletters = replace(replacedriveletters, "S:\", "\\nas01\allshare\")

exit function
err_handler:
    call handleerror("modDirectoryFunctions", "replaceDriveLetters", err.description, err.number)
end function

function addlastslash(linkstring as string) as string
on error goto err_handler

addlastslash = linkstring
if right(addlastslash, 1) <> "\" then addlastslash = addlastslash & "\"

exit function
err_handler:
    call handleerror("modDirectoryFunctions", "addLastSlash", err.description, err.number)
end function

function folderexists(sfile as variant) as boolean
on error goto err_handler

folderexists = false
if isnull(sfile) then exit function
if dir(sfile, vbdirectory) <> "" then folderexists = true

exit function
err_handler:
    if err.number = 52 then exit function
    call handleerror("modDirectoryFunctions", "FolderExists", err.description, err.number)
end function
