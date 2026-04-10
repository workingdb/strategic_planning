attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

private sub form_load()
on error goto err_handler

call settheme(me)

exit sub
err_handler:
    call handleerror(me.name, "Form_Load", err.description, err.number)
end sub

private sub find_click()
on error goto err_handler

    on error resume next
    docmd.gotocontrol screen.previouscontrol.name
    err.clear
    docmd.runcommand accmdfind
    if (macroerror <> 0) then
        msgbox macroerror.description, vbokonly, ""
    end if

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub new_click()
on error goto err_handler

    on error resume next
    docmd.gotorecord , "", acnewrec
    if (macroerror <> 0) then
        msgbox macroerror.description, vbokonly, ""
    end if

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub trash_click()
on error goto err_handler

    on error resume next
    docmd.gotocontrol screen.previouscontrol.name
    err.clear
    if (not form.newrecord) then
        docmd.runcommand accmddeleterecord
    end if
    if (form.newrecord and form.dirty) then
        docmd.runcommand accmdundo
    end if
    if (macroerror <> 0) then
        msgbox macroerror.description, vbokonly, ""
    end if

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub copyrecord_click()
on error goto err_handler

    on error resume next
    docmd.runcommand accmdselectrecord
    if (macroerror = 0) then
        docmd.runcommand accmdcopy
    end if
    if (macroerror = 0) then
        docmd.runcommand accmdrecordsgotonew
    end if
    if (macroerror = 0) then
        docmd.runcommand accmdselectrecord
    end if
    if (macroerror = 0) then
        docmd.runcommand accmdpaste
    end if
    if (macroerror <> 0) then
        msgbox macroerror.description, vbokonly, ""
    end if

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
