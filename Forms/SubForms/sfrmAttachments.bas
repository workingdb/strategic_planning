attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

private sub cmdopenlink_click()
 
    dim url as string
    url = trim(nz(me.directlink, ""))
 
    if len(url) = 0 then
        msgbox "No link found.", vbinformation
        exit sub
    end if
 
    'force windows to open in default browser (edge/chrome/etc.)
    createobject("WScript.Shell").run _
        "cmd /c start """" """ & url & """", 0, false
 
end sub

private sub form_load()
on error goto err_handler

call settheme(me)

exit sub
err_handler:
    call handleerror(me.name, "Form_Load", err.description, err.numbe)
end sub
