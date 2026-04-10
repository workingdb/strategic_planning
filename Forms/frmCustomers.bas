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

private sub saverecord_click()
on error goto err_handler

if me.dirty then me.dirty = false

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub newrecord_click()
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


private sub command6_click()

end sub
