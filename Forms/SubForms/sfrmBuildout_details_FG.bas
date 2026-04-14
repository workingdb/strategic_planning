attribute vb_globalnamespace = false
attribute vb_creatable = true
attribute vb_predeclaredid = true
attribute vb_exposed = false
option compare database
option explicit

private sub explode_click()
on error goto err_handler

dim conn as adodb.connection
dim rs as adodb.recordset

set conn = currentproject.connection
set rs = openrecordsetreadonly(conn, "SELECT * FROM tblBuildout_register_FG WHERE registerId = " & me.registerid)

do while not rs.eof
    call explodebom(rs!partnumber)
    rs.movenext
loop

clean_exit:
    on error resume next
    if not rs is nothing then
        if rs.state = adstateopen then rs.close
    end if
    
    set rs = nothing
    set conn = nothing

exit sub
err_handler:
    call handleerror(me.name, me.activecontrol.name, err.description, err.number)
end sub

public function explodebom(partnumber as string)
on error goto err_handler

explodebom = ""

if nz(partnumber, "") = "" then exit function

dim cn as adodb.connection
dim cmd as adodb.command
dim rs as adodb.recordset
dim sql as string
dim psegment1 as string
dim porgid as long

psegment1 = "26589"
porgid = 142

sql = ""

sql = sql & "WITH bom_tree (" & vbcrlf
sql = sql & "    top_assembly," & vbcrlf
sql = sql & "    parent_item_id," & vbcrlf
sql = sql & "    parent_segment1," & vbcrlf
sql = sql & "    component_item_id," & vbcrlf
sql = sql & "    component_segment1," & vbcrlf
sql = sql & "    organization_id," & vbcrlf
sql = sql & "    bom_level," & vbcrlf
sql = sql & "    component_path," & vbcrlf
sql = sql & "    qty_per_assembly," & vbcrlf
sql = sql & "    extended_qty" & vbcrlf
sql = sql & ") AS (" & vbcrlf

sql = sql & "    SELECT" & vbcrlf
sql = sql & "        asm.segment1 AS top_assembly," & vbcrlf
sql = sql & "        asm.inventory_item_id AS parent_item_id," & vbcrlf
sql = sql & "        asm.segment1 AS parent_segment1," & vbcrlf
sql = sql & "        comp.inventory_item_id AS component_item_id," & vbcrlf
sql = sql & "        comp.segment1 AS component_segment1," & vbcrlf
sql = sql & "        asm.organization_id AS organization_id," & vbcrlf
sql = sql & "        1 AS bom_level," & vbcrlf
sql = sql & "        asm.segment1 || ' -> ' || comp.segment1 AS component_path," & vbcrlf
sql = sql & "        bic.component_quantity AS qty_per_assembly," & vbcrlf
sql = sql & "        bic.component_quantity AS extended_qty" & vbcrlf
sql = sql & "    FROM inv.mtl_system_items_b asm" & vbcrlf
sql = sql & "    JOIN apps.bom_bill_of_materials bbm" & vbcrlf
sql = sql & "      ON bbm.assembly_item_id = asm.inventory_item_id" & vbcrlf
sql = sql & "     AND bbm.organization_id = asm.organization_id" & vbcrlf
sql = sql & "     AND bbm.alternate_bom_designator IS NULL" & vbcrlf
sql = sql & "    JOIN apps.bom_inventory_components bic" & vbcrlf
sql = sql & "      ON bic.bill_sequence_id = bbm.bill_sequence_id" & vbcrlf
sql = sql & "     AND bic.disable_date IS NULL" & vbcrlf
sql = sql & "    JOIN inv.mtl_system_items_b comp" & vbcrlf
sql = sql & "      ON comp.inventory_item_id = bic.component_item_id" & vbcrlf
sql = sql & "     AND comp.organization_id = asm.organization_id" & vbcrlf
sql = sql & "    WHERE asm.segment1 = ?" & vbcrlf
sql = sql & "      AND asm.organization_id = ?" & vbcrlf

sql = sql & "    UNION ALL" & vbcrlf

sql = sql & "    SELECT" & vbcrlf
sql = sql & "        bt.top_assembly," & vbcrlf
sql = sql & "        parent.inventory_item_id," & vbcrlf
sql = sql & "        parent.segment1," & vbcrlf
sql = sql & "        comp.inventory_item_id," & vbcrlf
sql = sql & "        comp.segment1," & vbcrlf
sql = sql & "        parent.organization_id," & vbcrlf
sql = sql & "        bt.bom_level + 1," & vbcrlf
sql = sql & "        bt.component_path || ' -> ' || comp.segment1," & vbcrlf
sql = sql & "        bic.component_quantity," & vbcrlf
sql = sql & "        bt.extended_qty * bic.component_quantity" & vbcrlf
sql = sql & "    FROM bom_tree bt" & vbcrlf
sql = sql & "    JOIN inv.mtl_system_items_b parent" & vbcrlf
sql = sql & "      ON parent.inventory_item_id = bt.component_item_id" & vbcrlf
sql = sql & "     AND parent.organization_id = bt.organization_id" & vbcrlf
sql = sql & "    JOIN apps.bom_bill_of_materials bbm" & vbcrlf
sql = sql & "      ON bbm.assembly_item_id = parent.inventory_item_id" & vbcrlf
sql = sql & "     AND bbm.organization_id = parent.organization_id" & vbcrlf
sql = sql & "     AND bbm.alternate_bom_designator IS NULL" & vbcrlf
sql = sql & "    JOIN apps.bom_inventory_components bic" & vbcrlf
sql = sql & "      ON bic.bill_sequence_id = bbm.bill_sequence_id" & vbcrlf
sql = sql & "     AND bic.disable_date IS NULL" & vbcrlf
sql = sql & "    JOIN inv.mtl_system_items_b comp" & vbcrlf
sql = sql & "      ON comp.inventory_item_id = bic.component_item_id" & vbcrlf
sql = sql & "     AND comp.organization_id = parent.organization_id" & vbcrlf
sql = sql & ")" & vbcrlf

sql = sql & "SEARCH DEPTH FIRST BY parent_segment1, component_segment1 SET sort_seq" & vbcrlf
sql = sql & "CYCLE component_item_id SET is_cycle TO 'Y' DEFAULT 'N'" & vbcrlf

sql = sql & "SELECT" & vbcrlf
sql = sql & "    bt.top_assembly," & vbcrlf
sql = sql & "    bt.component_segment1 AS lowest_level_component," & vbcrlf
sql = sql & "    bt.bom_level," & vbcrlf
sql = sql & "    bt.qty_per_assembly," & vbcrlf
sql = sql & "    bt.extended_qty," & vbcrlf
sql = sql & "    bt.component_path" & vbcrlf
sql = sql & "FROM bom_tree bt" & vbcrlf
sql = sql & "WHERE bt.is_cycle = 'N'" & vbcrlf
sql = sql & "  AND NOT EXISTS (" & vbcrlf
sql = sql & "        SELECT 1" & vbcrlf
sql = sql & "        FROM apps.bom_bill_of_materials bbm2" & vbcrlf
sql = sql & "        WHERE bbm2.assembly_item_id = bt.component_item_id" & vbcrlf
sql = sql & "          AND bbm2.organization_id = bt.organization_id" & vbcrlf
sql = sql & "          AND bbm2.alternate_bom_designator IS NULL" & vbcrlf
sql = sql & "  )" & vbcrlf
sql = sql & "ORDER BY bt.component_segment1, bt.component_path"

set cn = getoracleconnection()

set cmd = new adodb.command
set cmd.activeconnection = cn
cmd.commandtype = adcmdtext
cmd.commandtext = sql

cmd.parameters.append cmd.createparameter(, advarchar, adparaminput, 50, psegment1)
cmd.parameters.append cmd.createparameter(, adinteger, adparaminput, , porgid)

set rs = cmd.execute

do until rs.eof
    debug.print rs.fields("top_assembly").value, _
                rs.fields("lowest_level_component").value, _
                rs.fields("bom_level").value, _
                rs.fields("qty_per_assembly").value, _
                rs.fields("extended_qty").value, _
                rs.fields("component_path").value
    rs.movenext
loop


    
clean_exit:
    on error resume next
    if not rs is nothing then
        if rs.state = adstateopen then rs.close
    end if
    
    set rs = nothing
    set cmd = nothing
    cn.close
    set cn = nothing
    exit function

err_handler:
    call handleerror(me.name, "explodeBOM", err.description, err.number)
end function
