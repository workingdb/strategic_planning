attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

function trackupdate()
on error goto err_handler

if isnull(me.recordid) then exit function
call registerstratplanupdates("tblCapacityRequest_partnumbers", me.recordid, me.activecontrol.name, me.activecontrol.oldvalue, me.activecontrol, me.recordid, me.name)

exit function
err_handler:
    call handleerror(me.name, "trackUpdate", err.description, err.number)
end function

private sub capacityresults_afterupdate()
on error goto err_handler

trackupdate

if me.dirty then me.dirty = false

if nz(me.capacityresults, 0) <> 0 then
    sqlexecute "UPDATE dbo.tblCapacityRequest_partnumbers SET responseDate = GETDATE() WHERE recordId = " & me.recordid
    
    me.requery
    call registerstratplanupdates("tblCapacityRequest_partnumbers", me.recordid, "responseDate", "", now(), me.recordid, me.name)
end if

exit sub
err_handler:
    call handleerror(me.name, "trackUpdate", err.description, err.number)
end sub

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

private sub copy_click()
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

private sub mailreport_click()
on error goto err_handler

docmd.sendobject acreport, "Capacity Confirmation", "", "", "", "", "", "", true, ""

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub


private sub partnumber_afterupdate()
on error goto err_handler

if nz(me.partnumber, "") = "" then exit sub

call trackupdate

'find current unit
dim db as database
set db = currentdb()
dim invid, currentunit as string, rscat as recordset
invid = nz(idnam(me.partnumber, "NAM"), "")

currentunit = ""

if invid <> "" then
    set rscat = db.openrecordset("SELECT SEGMENT1 FROM INV_MTL_ITEM_CATEGORIES LEFT JOIN APPS_MTL_CATEGORIES_VL ON INV_MTL_ITEM_CATEGORIES.CATEGORY_ID = APPS_MTL_CATEGORIES_VL.CATEGORY_ID " & _
    "GROUP BY INV_MTL_ITEM_CATEGORIES.INVENTORY_ITEM_ID, APPS_MTL_CATEGORIES_VL.SEGMENT1, APPS_MTL_CATEGORIES_VL.STRUCTURE_ID HAVING STRUCTURE_ID = 101 AND [INVENTORY_ITEM_ID] = " & invid, dbopensnapshot)
    if rscat.recordcount > 0 then currentunit = nz(rscat!segment1, "")

    rscat.close
    set rscat = nothing
end if

if currentunit <> "" then
    dim unitid
    unitid = nz(dlookup("recordId", "tblUnits", "unitName = '" & currentunit & "'"), 0)
    me.unitid = unitid
end if

set db = nothing

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub remove_click()
on error goto err_handler

if msgbox("Are you sure you want to delete this?", vbyesno, "Please confirm") = vbyes then
    if nz(me.recordid, 0) <> 0 then call registerstratplanupdates("tblCapacityRequestDetail_partnumbers", me.recordid, "Part Number", nz(me.partnumber, ""), "Deleted", me.recordid, me.name)
    dbexecute ("DELETE FROM tblCapacityRequest_partnumbers WHERE [recordId] = " & me.recordid)
    me.requery
end if

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
