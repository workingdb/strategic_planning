attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit
 
private sub btngenerate_click()
    on error goto err_handler
 
    dim rt as string
    dim rptname as string
    dim whereclause as string
    dim d1 as date, d2 as date
 
    'use displayed text if cboreporttype is multi-column; fallback to value if not
    rt = trim(nz(me.cboreporttype.column(1), nz(me.cboreporttype.value, "")))
 
    if rt = "" then
        msgbox "Please select a Report Type.", vbexclamation
        exit sub
    end if
 
    select case rt
 
        case "Capacity by Date Range"
            rptname = "rpt_Capacity_DateRange"
 
            if isnull(me.txtstartdate) or isnull(me.txtenddate) then
                msgbox "Please enter both Start Date and End Date.", vbexclamation
                exit sub
            end if
            if not isdate(me.txtstartdate) or not isdate(me.txtenddate) then
                msgbox "Start Date and End Date must be valid dates.", vbexclamation
                exit sub
            end if
 
            d1 = datevalue(me.txtstartdate)
            d2 = datevalue(me.txtenddate)
 
            if d2 < d1 then
                msgbox "End Date must be on or after Start Date.", vbexclamation
                exit sub
            end if
 
            whereclause = "[RequestDate] Between #" & format(d1, "yyyy-mm-dd") & _
                          "# And #" & format(d2, "yyyy-mm-dd") & "#"
 
        case "KPI Report"
            rptname = "rpt_KPI_Dashboard"
 
            if isnull(me.txtstartdate) or isnull(me.txtenddate) then
                msgbox "Please enter both Start Date and End Date.", vbexclamation
                exit sub
            end if
            if not isdate(me.txtstartdate) or not isdate(me.txtenddate) then
                msgbox "Start Date and End Date must be valid dates.", vbexclamation
                exit sub
            end if
 
            d1 = datevalue(me.txtstartdate)
            d2 = datevalue(me.txtenddate)
 
            if d2 < d1 then
                msgbox "End Date must be on or after Start Date.", vbexclamation
                exit sub
            end if
 
            'set tempvars for kpi queries
            on error resume next
            tempvars.remove "StartDate"
            tempvars.remove "EndDate"
            on error goto 0
 
            tempvars.add "StartDate", d1
            tempvars.add "EndDate", d2
 
            'open without filter
            docmd.openreport rptname, acviewpreview
            exit sub
 
        case "Capacity by NAM"
            rptname = "rpt_Capacity_ByNAM"
 
            if isnull(me.txtnam) or trim(nz(me.txtnam, "")) = "" then
                msgbox "Please enter a NAM.", vbexclamation
                exit sub
            end if
 
            whereclause = "[partNumber] = """ & replace(trim(me.txtnam), """", """""") & """"
 
        case "Requests by Sales Manager"
            rptname = "rpt_SalesManager"
 
            if isnull(me.cbosalesmanager) then
                msgbox "Please select a Sales Manager.", vbexclamation
                exit sub
            end if
 
            whereclause = "[Requestor] = " & clng(me.cbosalesmanager)
 
        case "Capacity by Program Code"
            rptname = "rpt_Capacity_ByProgramCode"
 
            if isnull(me.cboprogramcode) or trim(nz(me.cboprogramcode, "")) = "" then
                msgbox "Please select a Program Code.", vbexclamation
                exit sub
            end if
 
            whereclause = "[Program] = """ & replace(trim(me.cboprogramcode), """", """""") & """"
 
        case "Past Due Report"
            rptname = "rpt_PastDue"
 
            if isnull(me.cbounit) then
                msgbox "Please select a Unit.", vbexclamation
                exit sub
            end if
 
            whereclause = "[Unit] = " & clng(me.cbounit)
 
        case "Quote Status"
            rptname = "rpt_Awarded"   '<<< change if your report name is different
 
            if isnull(me.cboquotestatus) then
                msgbox "Please select a Quote Status.", vbexclamation
                exit sub
            end if
 
            'if cboquotestatus is a text value dropdown
            whereclause = "[quoteStatus] = " & clng(me.cboquotestatus)
 
            'if cboquotestatus is actually a numeric id combo, use this instead:
            'whereclause = "[Quote] = " & clng(me.cboquotestatus)
 
        case else
            msgbox "Report type not coded yet: " & rt, vbexclamation
            exit sub
    end select
 
    if len(whereclause) > 0 then
        docmd.openreport rptname, acviewpreview, , whereclause
    else
        docmd.openreport rptname, acviewpreview
    end if
 
    exit sub
 
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
 
 
private sub cboreporttype_afterupdate()
    on error goto err_handler
 
    dim rt as string
 
    'use displayed text if cboreporttype is multi-column; fallback to value if not
    rt = trim(nz(me.cboreporttype.column(1), nz(me.cboreporttype.value, "")))
 
    '--- hide everything first ---
    me.txtstartdate.visible = false
    me.txtenddate.visible = false
    me.txtnam.visible = false
    me.cbosalesmanager.visible = false
    me.cboprogramcode.visible = false
    me.cbounit.visible = false
    me.cboquotestatus.visible = false
 
    '--- default: disable + no tabbing ---
    me.txtstartdate.enabled = false: me.txtstartdate.tabstop = false
    me.txtenddate.enabled = false: me.txtenddate.tabstop = false
    me.txtnam.enabled = false: me.txtnam.tabstop = false
    me.cbosalesmanager.enabled = false: me.cbosalesmanager.tabstop = false
    me.cboprogramcode.enabled = false: me.cboprogramcode.tabstop = false
    me.cbounit.enabled = false: me.cbounit.tabstop = false
    me.cboquotestatus.enabled = false: me.cboquotestatus.tabstop = false
 
    'optional: clear unused values when switching report types
    me.txtstartdate = null
    me.txtenddate = null
    me.txtnam = null
    me.cbosalesmanager = null
    me.cboprogramcode = null
    me.cbounit = null
    me.cboquotestatus = null
 
    '--- show only what the selected report needs ---
    select case rt
 
        case "Capacity by Date Range", "KPI Report"
            me.txtstartdate.visible = true
            me.txtenddate.visible = true
            me.txtstartdate.enabled = true: me.txtstartdate.tabstop = true
            me.txtenddate.enabled = true: me.txtenddate.tabstop = true
 
        case "Capacity by NAM"
            me.txtnam.visible = true
            me.txtnam.enabled = true: me.txtnam.tabstop = true
 
        case "Requests by Sales Manager"
            me.cbosalesmanager.visible = true
            me.cbosalesmanager.enabled = true: me.cbosalesmanager.tabstop = true
 
        case "Capacity by Program Code"
            me.cboprogramcode.visible = true
            me.cboprogramcode.enabled = true: me.cboprogramcode.tabstop = true
 
        case "Past Due Report"
            me.cbounit.visible = true
            me.cbounit.enabled = true: me.cbounit.tabstop = true
 
        case "Quote Status"
            me.cboquotestatus.visible = true
            me.cboquotestatus.enabled = true: me.cboquotestatus.tabstop = true
 
        'no controls needed for anything else
 
    end select
 
    exit sub
 
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub
 
 
private sub form_load()
    on error goto err_handler
 
    'clear all user selections/inputs when the launcher opens
    call settheme(me)
 
    on error resume next
 
    'combos
    me.cboreporttype = null
    me.cbosalesmanager = null
    me.cboprogramcode = null
    me.cbounit = null
    me.cboquotestatus = null
 
    'text boxes
    me.txtstartdate = null
    me.txtenddate = null
    me.txtnam = null
 
    'optional: clear tempvars so nothing carries over
    tempvars.remove "StartDate"
    tempvars.remove "EndDate"
 
    'hide/disable all optional controls on load
    me.txtstartdate.visible = false
    me.txtenddate.visible = false
    me.txtnam.visible = false
    me.cbosalesmanager.visible = false
    me.cboprogramcode.visible = false
    me.cbounit.visible = false
    me.cboquotestatus.visible = false
 
    me.txtstartdate.enabled = false: me.txtstartdate.tabstop = false
    me.txtenddate.enabled = false: me.txtenddate.tabstop = false
    me.txtnam.enabled = false: me.txtnam.tabstop = false
    me.cbosalesmanager.enabled = false: me.cbosalesmanager.tabstop = false
    me.cboprogramcode.enabled = false: me.cboprogramcode.tabstop = false
    me.cbounit.enabled = false: me.cbounit.tabstop = false
    me.cboquotestatus.enabled = false: me.cboquotestatus.tabstop = false
 
    me.cboreporttype.setfocus
 
    exit sub
 
err_handler:
    call handleerror(me.name, "Form_Load", err.description, err.number)
end sub
