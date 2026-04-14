attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

private sub remove_click()
on error goto err_handler

if msgbox("Are you sure you want to delete this?", vbyesno, "Please confirm") = vbyes then
    'if nz(me.recordid, 0) <> 0 then call registerstratplanupdates("tblCapacityRequestDetail_partnumbers", me.recordid, "Part Number", nz(me.partnumber, ""), "Deleted", me.recordid, me.name)
    dbexecute ("DELETE FROM tblBuildout_register_comp WHERE [recordId] = " & me.recordid)
    me.requery
end if

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
