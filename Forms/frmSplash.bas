attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

private sub form_load()
    on error goto err_handler

    ' --- initial setup ---
    tempvars.add "loadAmount", 0
    tempvars.add "loadWd", 8160
    tempvars.add "dbVersion", grabversion()
    me.lblfrozen.visible = false
    call setsplashloading("Setting up app stuff...")
    me.lblversion.caption = tempvars!dbversion
    me.lblversion.visible = true

    sizeaccess 280, 280
    me.move -2600, -1000
    
        'make sure driver reference for sql server is ok
    call relinksqltables

    ' use the adodb version of logclick we created earlier
    call logclick("Form_Load", me.name)

    ' splash image logic (keeping dlookup for local settings)
    me.picture = "\\data\mdbdata\WorkingDB\Pictures\Splash\splash" & randomnumber(0, dlookup("splashCount", "tblDBinfoBE", "ID = 1")) & ".png"
    
    on error resume next
    me.imguser.picture = "\\data\mdbdata\WorkingDB\Pictures\Avatars\" & environ("username") & ".png"
    on error goto err_handler
    
    doevents
    form_frmsplash.setfocus
    doevents

    ' --- ribbon & shortcuts ---
    if commandbars("Ribbon").height > 100 then commandbars.executemso "MinimizeRibbon"
    docmd.showtoolbar "Ribbon", actoolbarno

    on error resume next
    dim fso as object: set fso = createobject("Scripting.FileSystemObject")
    fso.copyfile "\\data\mdbdata\WorkingDB\Batch\Working DB.lnk", "\\homes\data\" & environ("username") & "\Desktop\Working DB.lnk"
    fso.copyfile "\\data\mdbdata\WorkingDB\build\workingdb_ghost\WorkingDB_ghost.accde", "C:\workingdb\WorkingDB_ghost.accde"
    openpath "\\data\mdbdata\WorkingDB\build\workingdb_commands\openGhost.vbs"
    on error goto err_handler

    ' --- database logic (adodb conversion) ---
    call setsplashloading("Doing some digging on you...")
    
    dim conn as adodb.connection: set conn = currentproject.connection
    dim rsuser as new adodb.recordset, rsperm as new adodb.recordset

    rsuser.open "SELECT * FROM tblUserSettings WHERE [username] = '" & environ("username") & "'", conn, adopenkeyset, adlockoptimistic
    if rsuser.eof then
        msgbox "You need to have an account in WorkingDB to access this.", vbokonly, "Welcome"
        application.quit
    end if

    rsperm.open "SELECT * FROM tblPermissions WHERE [User] = '" & environ("username") & "'", conn, adopenkeyset, adlockoptimistic
    
    if rsperm.eof then
        msgbox "You need to have an account in WorkingDB to access this.", vbokonly, "Welcome"
        application.quit
    end if

    ' --- set tempvars ---
    tempvars.add "dept", nz(rsperm!dept, "")
    tempvars.add "org", nz(rsperm!org, 4)
    tempvars.add "smallScreen", nz(rsuser!smallscreenmode, "False")

    ' --- theme logic ---
    if nz(rsuser!themeid, 0) <> 0 then
        dim rstheme as new adodb.recordset
        rstheme.open "SELECT * FROM tblTheme WHERE recordId = " & rsuser!themeid, conn, adopenforwardonly, adlockreadonly
        if not rstheme.eof then
            tempvars.add "themeMode", iif(rstheme!darkmode, "Dark", "Light")
            tempvars.add "themePrimary", cstr(rstheme!primarycolor)
            tempvars.add "themeSecondary", cstr(rstheme!secondarycolor)
            tempvars.add "themeAccent", cstr(rstheme!accentcolor)
            tempvars.add "themeColorLevels", cstr(rstheme!colorlevels)
        end if
        rstheme.close
    end if

    ' --- finalize startup ---
    call setsplashloading("Running daily checks...")
    call grabjoke
    
    docmd.openform "DASHBOARD"
    forms!dashboard.visible = false
    
    docmd.close acform, me.name
    call maximizeaccess
    forms!dashboard.visible = true
    docmd.maximize

cleanup:
    if rsuser.state = adstateopen then rsuser.close
    if rsperm.state = adstateopen then rsperm.close
    set rsuser = nothing: set rsperm = nothing
    set conn = nothing
    exit sub

err_handler:
    call handleerror(me.name, "Form_Load", err.description, err.number)
    resume cleanup
end sub

function grabjoke()
on error goto err_handler

dim joke as string
joke = nz(dlookup("[factText]", "tblFacts", "[factDate] = #" & date & "#"))

tempvars.add "joke", joke

exit function
err_handler:
    call handleerror(me.name, "grabJoke", err.description, err.number)
end function
