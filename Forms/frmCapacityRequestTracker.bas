attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

function applyfilter(parameter as string)

dim db as database
set db = currentdb()

dim qdf as querydef

set qdf = db.querydefs("frmCapacityRequestTracker_PT")

if parameter = "" then
    qdf.sql = split(qdf.sql, "c.ID")(0) & " c.ID;"
else
    qdf.sql = split(qdf.sql, "c.ID")(0) & " c.ID WHERE EXISTS (SELECT 1 From tblCapacityRequest_partnumbers As cp WHERE cp.requestId = cr.recordId AND " & parameter & ");"
end if

db.querydefs.refresh

set qdf = nothing
set db = nothing

me.requery

end function
 
private sub capacityresults_afterupdate()
on error goto err_handler

select case me.activecontrol
    case 0 'no response
        applyfilter ("cp.capacityResults is null")
    case 9999 'all
        applyfilter ("")
    case else 'specific based on id
        applyfilter ("cp.capacityResults = " & me.capacityresults)
end select

me.partnumfilt = ""

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub customer_label_click()
    on error goto err_handler
    me.customer.setfocus
    docmd.runcommand accmdfiltermenu
    exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
 
private sub eop_label_click()
    on error goto err_handler
    me.eop.setfocus
    docmd.runcommand accmdfiltermenu
    exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub partnumfilt_afterupdate()
on error goto err_handler

if isnull(me.partnumfilt) then 'see all
    applyfilter ("")
else 'filter by part number
    applyfilter ("cp.partNumber = '" & me.partnumfilt & "'")
    me.capacityresults = 9999
end if

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

private sub program_label_click()
    on error goto err_handler
    me.program.setfocus
    docmd.runcommand accmdfiltermenu
    exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
 
private sub recordid_label_click()
    on error goto err_handler
    me.recordid.setfocus
    docmd.runcommand accmdfiltermenu
    exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
 
private sub request_date_label_click()
    on error goto err_handler
    me.requestdate.setfocus
    docmd.runcommand accmdfiltermenu
    exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
 
private sub request_type_label_click()
on error goto err_handler
    me.request_type.setfocus
    docmd.runcommand accmdfiltermenu
    exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
 
private sub requestor_label_click()
    on error goto err_handler
    me.requestor.setfocus
    docmd.runcommand accmdfiltermenu
    exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
 
private sub sop_label_click()
    on error goto err_handler
    me.sop.setfocus
    docmd.runcommand accmdfiltermenu
    exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
 
private sub newrequest_click()
    on error goto err_handler
    docmd.openform "frmCapacityRequestDetails", , , , acformadd
    exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
 
private sub opendetails_click()
    on error goto errhandler
 
    if isnull(me.recordid) then exit sub
 
    docmd.openform "frmCapacityRequestDetails", acnormal, , "recordId = " & me.recordid
 
exit sub
errhandler:
    msgbox "Open Details error " & err.number & ":" & vbcrlf & err.description, vbexclamation
end sub

private sub form_load()
on error goto err_handler

call settheme(me)

applyfilter ("cp.capacityResults is null")

me.capacityresults = 0

exit sub
err_handler:
    call handleerror(me.name, "Form_Load", err.description, err.numbe)
end sub
