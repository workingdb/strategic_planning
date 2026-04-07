Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_AfterInsert()
On Error GoTo Err_Handler

Call explodeBOM(Me.partNumber)
    
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, err.Description, err.Number)
End Sub

Public Function explodeBOM(partNumber As String)
On Error GoTo Err_Handler

explodeBOM = ""

If Nz(partNumber, "") = "" Then Exit Function

Dim cn As ADODB.Connection
Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset
Dim sql As String
Dim pSegment1 As String
Dim pOrgId As Long

pSegment1 = "26589"
pOrgId = 142

sql = ""

sql = sql & "WITH bom_tree (" & vbCrLf
sql = sql & "    top_assembly," & vbCrLf
sql = sql & "    parent_item_id," & vbCrLf
sql = sql & "    parent_segment1," & vbCrLf
sql = sql & "    component_item_id," & vbCrLf
sql = sql & "    component_segment1," & vbCrLf
sql = sql & "    organization_id," & vbCrLf
sql = sql & "    bom_level," & vbCrLf
sql = sql & "    component_path," & vbCrLf
sql = sql & "    qty_per_assembly," & vbCrLf
sql = sql & "    extended_qty" & vbCrLf
sql = sql & ") AS (" & vbCrLf

sql = sql & "    SELECT" & vbCrLf
sql = sql & "        asm.segment1 AS top_assembly," & vbCrLf
sql = sql & "        asm.inventory_item_id AS parent_item_id," & vbCrLf
sql = sql & "        asm.segment1 AS parent_segment1," & vbCrLf
sql = sql & "        comp.inventory_item_id AS component_item_id," & vbCrLf
sql = sql & "        comp.segment1 AS component_segment1," & vbCrLf
sql = sql & "        asm.organization_id AS organization_id," & vbCrLf
sql = sql & "        1 AS bom_level," & vbCrLf
sql = sql & "        asm.segment1 || ' -> ' || comp.segment1 AS component_path," & vbCrLf
sql = sql & "        bic.component_quantity AS qty_per_assembly," & vbCrLf
sql = sql & "        bic.component_quantity AS extended_qty" & vbCrLf
sql = sql & "    FROM inv.mtl_system_items_b asm" & vbCrLf
sql = sql & "    JOIN apps.bom_bill_of_materials bbm" & vbCrLf
sql = sql & "      ON bbm.assembly_item_id = asm.inventory_item_id" & vbCrLf
sql = sql & "     AND bbm.organization_id = asm.organization_id" & vbCrLf
sql = sql & "     AND bbm.alternate_bom_designator IS NULL" & vbCrLf
sql = sql & "    JOIN apps.bom_inventory_components bic" & vbCrLf
sql = sql & "      ON bic.bill_sequence_id = bbm.bill_sequence_id" & vbCrLf
sql = sql & "     AND bic.disable_date IS NULL" & vbCrLf
sql = sql & "    JOIN inv.mtl_system_items_b comp" & vbCrLf
sql = sql & "      ON comp.inventory_item_id = bic.component_item_id" & vbCrLf
sql = sql & "     AND comp.organization_id = asm.organization_id" & vbCrLf
sql = sql & "    WHERE asm.segment1 = ?" & vbCrLf
sql = sql & "      AND asm.organization_id = ?" & vbCrLf

sql = sql & "    UNION ALL" & vbCrLf

sql = sql & "    SELECT" & vbCrLf
sql = sql & "        bt.top_assembly," & vbCrLf
sql = sql & "        parent.inventory_item_id," & vbCrLf
sql = sql & "        parent.segment1," & vbCrLf
sql = sql & "        comp.inventory_item_id," & vbCrLf
sql = sql & "        comp.segment1," & vbCrLf
sql = sql & "        parent.organization_id," & vbCrLf
sql = sql & "        bt.bom_level + 1," & vbCrLf
sql = sql & "        bt.component_path || ' -> ' || comp.segment1," & vbCrLf
sql = sql & "        bic.component_quantity," & vbCrLf
sql = sql & "        bt.extended_qty * bic.component_quantity" & vbCrLf
sql = sql & "    FROM bom_tree bt" & vbCrLf
sql = sql & "    JOIN inv.mtl_system_items_b parent" & vbCrLf
sql = sql & "      ON parent.inventory_item_id = bt.component_item_id" & vbCrLf
sql = sql & "     AND parent.organization_id = bt.organization_id" & vbCrLf
sql = sql & "    JOIN apps.bom_bill_of_materials bbm" & vbCrLf
sql = sql & "      ON bbm.assembly_item_id = parent.inventory_item_id" & vbCrLf
sql = sql & "     AND bbm.organization_id = parent.organization_id" & vbCrLf
sql = sql & "     AND bbm.alternate_bom_designator IS NULL" & vbCrLf
sql = sql & "    JOIN apps.bom_inventory_components bic" & vbCrLf
sql = sql & "      ON bic.bill_sequence_id = bbm.bill_sequence_id" & vbCrLf
sql = sql & "     AND bic.disable_date IS NULL" & vbCrLf
sql = sql & "    JOIN inv.mtl_system_items_b comp" & vbCrLf
sql = sql & "      ON comp.inventory_item_id = bic.component_item_id" & vbCrLf
sql = sql & "     AND comp.organization_id = parent.organization_id" & vbCrLf
sql = sql & ")" & vbCrLf

sql = sql & "SEARCH DEPTH FIRST BY parent_segment1, component_segment1 SET sort_seq" & vbCrLf
sql = sql & "CYCLE component_item_id SET is_cycle TO 'Y' DEFAULT 'N'" & vbCrLf

sql = sql & "SELECT" & vbCrLf
sql = sql & "    bt.top_assembly," & vbCrLf
sql = sql & "    bt.component_segment1 AS lowest_level_component," & vbCrLf
sql = sql & "    bt.bom_level," & vbCrLf
sql = sql & "    bt.qty_per_assembly," & vbCrLf
sql = sql & "    bt.extended_qty," & vbCrLf
sql = sql & "    bt.component_path" & vbCrLf
sql = sql & "FROM bom_tree bt" & vbCrLf
sql = sql & "WHERE bt.is_cycle = 'N'" & vbCrLf
sql = sql & "  AND NOT EXISTS (" & vbCrLf
sql = sql & "        SELECT 1" & vbCrLf
sql = sql & "        FROM apps.bom_bill_of_materials bbm2" & vbCrLf
sql = sql & "        WHERE bbm2.assembly_item_id = bt.component_item_id" & vbCrLf
sql = sql & "          AND bbm2.organization_id = bt.organization_id" & vbCrLf
sql = sql & "          AND bbm2.alternate_bom_designator IS NULL" & vbCrLf
sql = sql & "  )" & vbCrLf
sql = sql & "ORDER BY bt.component_segment1, bt.component_path"

Set cn = GetOracleConnection()

Set cmd = New ADODB.Command
Set cmd.ActiveConnection = cn
cmd.CommandType = adCmdText
cmd.CommandText = sql

cmd.Parameters.Append cmd.CreateParameter(, adVarChar, adParamInput, 50, pSegment1)
cmd.Parameters.Append cmd.CreateParameter(, adInteger, adParamInput, , pOrgId)

Set rs = cmd.Execute

Do Until rs.EOF
    Debug.Print rs.Fields("top_assembly").Value, _
                rs.Fields("lowest_level_component").Value, _
                rs.Fields("bom_level").Value, _
                rs.Fields("qty_per_assembly").Value, _
                rs.Fields("extended_qty").Value, _
                rs.Fields("component_path").Value
    rs.MoveNext
Loop


    
Clean_Exit:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    
    Set rs = Nothing
    Set cmd = Nothing
    cn.Close
    Set cn = Nothing
    Exit Function

Err_Handler:
    Call handleError(Me.name, "explodeBOM", err.Description, err.Number)
End Function
