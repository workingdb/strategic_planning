option compare database
option explicit

public g_cnoracle as adodb.connection

public function getoracleconnection() as adodb.connection
    on error goto reconnect

    if not g_cnoracle is nothing then
        if g_cnoracle.state = adstateopen then
            set getoracleconnection = g_cnoracle
            exit function
        end if
    end if

reconnect:
    set g_cnoracle = new adodb.connection

    g_cnoracle.connectionstring = "DRIVER={Oracle in OraClient11g_home1};SERVER=ebsprd1.world;UID=WorkingDB;PWD=WorkingDB2341;DBQ=ebsprd1.world;"
    g_cnoracle.open

    set getoracleconnection = g_cnoracle
end function

public function getinventoryownerunit(byval partnumber) as string
on error goto err_handler

getinventoryownerunit = ""

if nz(partnumber, "") = "" then exit function

    dim cn as adodb.connection
    dim rs as adodb.recordset
    dim sql as string
    
    sql = "SELECT msi.SEGMENT1 AS PN, mcv.SEGMENT1 AS UnitName, mcv.SEGMENT2 AS PartType " & _
                "From inV.MTL_ITEM_CATEGORIES mic " & _
                    "JOIN INV.MTL_SYSTEM_ITEMS_B msi ON mic.INVENTORY_ITEM_ID = msi.INVENTORY_ITEM_ID " & _
                    "LEFT JOIN APPS.MTL_CATEGORIES_VL mcv ON mic.CATEGORY_ID = mcv.CATEGORY_ID " & _
                "where STRUCTURE_ID = 101 AND msi.SEGMENT1 = '" & replace(cstr(partnumber), "'", "''") & "' " & _
                "Group By msi.SEGMENT1, mcv.SEGMENT1, mcv.SEGMENT2;"

    set cn = getoracleconnection()

    set rs = new adodb.recordset
    rs.open sql, cn, adopenforwardonly, adlockreadonly

    if not rs.eof then
        getinventoryownerunit = nz(rs.fields("UnitName").value, "")
    end if

clean_exit:
    on error resume next
    if not rs is nothing then
        if rs.state = adstateopen then rs.close
    end if
    set rs = nothing
    set cn = nothing
    exit function

err_handler:
    call handleerror("modOracle", "getInventoryOwnerUnit", err.description, err.number)
    resume clean_exit
end function

public function getinventoryownertype(byval partnumber) as string
'on error goto err_handler

getinventoryownertype = ""

if nz(partnumber, "") = "" then exit function

    dim cn as adodb.connection
    dim rs as adodb.recordset
    dim sql as string
    
    sql = "SELECT msi.SEGMENT1 AS PN, mcv.SEGMENT1 AS UnitName, mcv.SEGMENT2 AS PartType " & _
                "From inV.MTL_ITEM_CATEGORIES mic " & _
                    "JOIN INV.MTL_SYSTEM_ITEMS_B msi ON mic.INVENTORY_ITEM_ID = msi.INVENTORY_ITEM_ID " & _
                    "LEFT JOIN APPS.MTL_CATEGORIES_VL mcv ON mic.CATEGORY_ID = mcv.CATEGORY_ID " & _
                "where STRUCTURE_ID = 101 AND msi.SEGMENT1 = '" & replace(cstr(partnumber), "'", "''") & "' " & _
                "Group By msi.SEGMENT1, mcv.SEGMENT1, mcv.SEGMENT2;"

    set cn = getoracleconnection()

    set rs = new adodb.recordset
    rs.open sql, cn, adopenforwardonly, adlockreadonly

    if not rs.eof then
        getinventoryownertype = nz(rs.fields("PartType").value, "")
    end if

clean_exit:
    on error resume next
    if not rs is nothing then
        if rs.state = adstateopen then rs.close
    end if
    set rs = nothing
    set cn = nothing
    exit function

err_handler:
    call handleerror("modOracle", "getInventoryOwnerType", err.description, err.number)
    resume clean_exit
end function

public function getstandardcostowner(byval partnumber) as string
on error goto err_handler

getstandardcostowner = ""

if nz(partnumber, "") = "" then exit function

    dim cn as adodb.connection
    dim rs as adodb.recordset
    dim sql as string
    
    sql = "SELECT msi.SEGMENT1 AS PN, mcv.SEGMENT3 AS Standard_Cost_Owner " & _
                "From inV.MTL_ITEM_CATEGORIES mic " & _
                    "JOIN INV.MTL_SYSTEM_ITEMS_B msi ON mic.INVENTORY_ITEM_ID = msi.INVENTORY_ITEM_ID " & _
                    "LEFT JOIN APPS.MTL_CATEGORIES_VL mcv ON mic.CATEGORY_ID = mcv.CATEGORY_ID " & _
                "where STRUCTURE_ID = 50349 AND msi.SEGMENT1 = '" & replace(cstr(partnumber), "'", "''") & "' " & _
                "Group By msi.SEGMENT1,mcv.SEGMENT3;"

    set cn = getoracleconnection()

    set rs = new adodb.recordset
    rs.open sql, cn, adopenforwardonly, adlockreadonly

    if not rs.eof then
        getstandardcostowner = nz(rs.fields("Standard_Cost_Owner").value, "")
    end if

clean_exit:
    on error resume next
    if not rs is nothing then
        if rs.state = adstateopen then rs.close
    end if
    set rs = nothing
    set cn = nothing
    exit function

err_handler:
    call handleerror("modOracle", "getStandardCostOwner", err.description, err.number)
    resume clean_exit
end function

public function getcustomer(partnumber) as string
on error goto err_handler

getcustomer = ""

if nz(partnumber, "") = "" then exit function

    dim cn as adodb.connection
    dim rs as adodb.recordset
    dim sql as string
    
    sql = "SELECT C.CUSTOMER_NAME " & _
            "From APPS.XXCUS_CUSTOMERS C " & _
                "INNER JOIN INV.MTL_CUSTOMER_ITEMS ci ON c.CUSTOMER_ID = ci.CUSTOMER_ID " & _
                "INNER JOIN INV.MTL_CUSTOMER_ITEM_XREFS ic ON ic.CUSTOMER_ITEM_ID = ci.CUSTOMER_ITEM_ID " & _
                "INNER JOIN INV.MTL_SYSTEM_ITEMS_B si ON si.INVENTORY_ITEM_ID = ic.INVENTORY_ITEM_ID " & _
            "WHERE si.SEGMENT1 = '" & replace(cstr(partnumber), "'", "''") & "' AND ci.INACTIVE_FLAG = 'N' " & _
            "Group By c.CUSTOMER_NAME;"

    set cn = getoracleconnection()

    set rs = new adodb.recordset
    rs.open sql, cn, adopenforwardonly, adlockreadonly
    
    do while not rs.eof
        getcustomer = getcustomer & nz(rs.fields("CUSTOMER_NAME").value, "") & vbnewline
        rs.movenext
    loop
    
clean_exit:
    on error resume next
    if not rs is nothing then
        if rs.state = adstateopen then rs.close
    end if
    set rs = nothing
    set cn = nothing
    exit function

err_handler:
    call handleerror("modOracle", "getCustomer", err.description, err.number)
    resume clean_exit
end function

public function findcost(byval partnumber as string, byval costtype as string, byval org as string) as double
    on error goto err_handler

    findcost = 0

    dim db as dao.database
    dim rslocal as dao.recordset
    dim orgid as long

    '--- step 1: get org id from access table
    set db = currentdb()
    set rslocal = db.openrecordset( _
        "SELECT ID FROM tblOrgs WHERE Org = '" & replace(org, "'", "''") & "'", _
        dbopensnapshot)

    if rslocal.eof then goto clean_exit

    orgid = rslocal!id

    rslocal.close
    set rslocal = nothing
    set db = nothing

    '--- step 2: query oracle using adodb
    dim cn as adodb.connection
    dim cmd as adodb.command
    dim rs as adodb.recordset

    set cn = getoracleconnection()

    set cmd = new adodb.command
    set cmd.activeconnection = cn
    cmd.commandtype = adcmdtext

    cmd.commandtext = _
        "SELECT ITEM_COST " & _
        "FROM APPS.CST_ITEM_COST_TYPE_V " & _
        "WHERE ITEM_NUMBER = ? " & _
        "AND COST_TYPE = ? " & _
        "AND ORGANIZATION_ID = ?"

    ' parameters (safer + faster)
    cmd.parameters.append cmd.createparameter("pItem", advarchar, adparaminput, 50, partnumber)
    cmd.parameters.append cmd.createparameter("pCostType", advarchar, adparaminput, 50, costtype)
    cmd.parameters.append cmd.createparameter("pOrgID", adinteger, adparaminput, , orgid)

    set rs = cmd.execute

    if not rs.eof then
        findcost = nz(rs.fields("ITEM_COST").value, 0)
    end if

clean_exit:
    on error resume next

    if not rs is nothing then if rs.state = adstateopen then rs.close

    set rs = nothing
    set cmd = nothing
    set cn = nothing

    exit function

err_handler:
    call handleerror("modOracle", "findCost", err.description, err.number)
    resume clean_exit
end function

public function getcustomername(customerid as long) as string
on error goto err_handler

getcustomername = ""

if nz(customerid, "") = "" then exit function

    dim cn as adodb.connection
    dim rs as adodb.recordset
    dim sql as string
    
    sql = "SELECT C.CUSTOMER_NAME " & _
            "From APPS.XXCUS_CUSTOMERS C " & _
            "WHERE c.CUSTOMER_ID = " & customerid & ";"

    set cn = getoracleconnection()

    set rs = new adodb.recordset
    rs.open sql, cn, adopenforwardonly, adlockreadonly
    
    if not rs.eof then
        getcustomername = nz(rs.fields("CUSTOMER_NAME").value, "")
    end if
    
clean_exit:
    on error resume next
    if not rs is nothing then
        if rs.state = adstateopen then rs.close
    end if
    set rs = nothing
    set cn = nothing
    exit function

err_handler:
    call handleerror("modOracle", "getCustomerName", err.description, err.number)
    resume clean_exit
end function

public function finddescription(byval partnumber as variant) as string
    on error goto err_handler

    finddescription = ""

    if nz(partnumber, "") = "" then exit function

    dim cn as adodb.connection
    dim cmd as adodb.command
    dim rs as adodb.recordset
    dim partvalue as string

    partvalue = cstr(partnumber)
    set cn = getoracleconnection()

    '========================================
    ' 1) check master items table first
    '========================================
    set cmd = new adodb.command
    set cmd.activeconnection = cn
    cmd.commandtype = adcmdtext
    cmd.commandtext = _
        "SELECT DESCRIPTION " & _
        "FROM INV.MTL_SYSTEM_ITEMS_B " & _
        "WHERE SEGMENT1 = ? " & _
        "AND ROWNUM = 1"

    cmd.parameters.append cmd.createparameter("pPartNumber", advarchar, adparaminput, 50, partvalue)

    set rs = cmd.execute

    if not rs.eof then
        finddescription = nz(rs.fields("DESCRIPTION").value, "")
        goto clean_exit
    end if

    rs.close
    set rs = nothing
    set cmd = nothing

    '========================================
    ' 2) check all sif tables at once
    '    using union all + rownum
    '========================================
    set cmd = new adodb.command
    set cmd.activeconnection = cn
    cmd.commandtype = adcmdtext
    cmd.commandtext = _
        "SELECT PART_DESCRIPTION " & _
        "FROM (" & _
        "    SELECT PART_DESCRIPTION " & _
        "    FROM (" & _
        "        SELECT PART_DESCRIPTION, 1 AS SRC_ORDER " & _
        "        FROM APPS.Q_SIF_NEW_ASSEMBLED_PART_V " & _
        "        WHERE NIFCO_PART_NUMBER = ? " & _
        "        UNION ALL " & _
        "        SELECT PART_DESCRIPTION, 2 AS SRC_ORDER " & _
        "        FROM APPS.Q_SIF_NEW_MOLDED_PART_V " & _
        "        WHERE NIFCO_PART_NUMBER = ? " & _
        "        UNION ALL " & _
        "        SELECT PART_DESCRIPTION, 3 AS SRC_ORDER " & _
        "        FROM APPS.Q_SIF_NEW_PURCHASING_PART_V " & _
        "        WHERE NIFCO_PART_NUMBER = ? " & _
        "        ORDER BY SRC_ORDER " & _
        "    ) " & _
        ") " & _
        "WHERE ROWNUM = 1"

    cmd.parameters.append cmd.createparameter("pPart1", advarchar, adparaminput, 50, partvalue)
    cmd.parameters.append cmd.createparameter("pPart2", advarchar, adparaminput, 50, partvalue)
    cmd.parameters.append cmd.createparameter("pPart3", advarchar, adparaminput, 50, partvalue)

    set rs = cmd.execute

    if not rs.eof then
        finddescription = nz(rs.fields("PART_DESCRIPTION").value, "")
    end if

clean_exit:
    on error resume next

    if not rs is nothing then
        if rs.state = adstateopen then rs.close
    end if

    set rs = nothing
    set cmd = nothing
    set cn = nothing

    exit function

err_handler:
    call handleerror("modOracle", "findDescription", err.description, err.number)
    resume clean_exit
end function

public function findpartrev(partnumber) as string
on error goto err_handler

findpartrev = "00"

if nz(partnumber, "") = "" then exit function

    dim cn as adodb.connection
    dim rs as adodb.recordset
    dim sql as string
    
    sql = "SELECT Max(NEW_ITEM_REVISION) As REV " & _
            "From ENG.ENG_REVISED_ITEMS RI " & _
            "INNER JOIN INV.MTL_SYSTEM_ITEMS_B SI ON RI.REVISED_ITEM_ID = SI.INVENTORY_ITEM_ID " & _
            "WHERE ((RI.IMPLEMENTATION_DATE Is Not Null) AND (SI.SEGMENT1='" & replace(cstr(partnumber), "'", "''") & "'));"

    set cn = getoracleconnection()

    set rs = new adodb.recordset
    rs.open sql, cn, adopenforwardonly, adlockreadonly
    
    if not rs.eof then
        findpartrev = nz(rs.fields("REV").value, "00")
    end if
    
clean_exit:
    on error resume next
    if not rs is nothing then
        if rs.state = adstateopen then rs.close
    end if
    set rs = nothing
    set cn = nothing
    exit function

err_handler:
    call handleerror("modOracle", "findPartRev", err.description, err.number)
    resume clean_exit
end function

function loadecotype(changenotice as string) as string
    on error goto err_handler

loadecotype = ""

if nz(changenotice, "") = "" then exit function

    dim cn as adodb.connection
    dim rs as adodb.recordset
    dim sql as string
    
    sql = "SELECT CHANGE_ORDER_TYPE_ID from ENG.ENG_ENGINEERING_CHANGES where CHANGE_NOTICE = '" & strquotereplace(changenotice) & "'"

    set cn = getoracleconnection()

    set rs = new adodb.recordset
    rs.open sql, cn, adopenforwardonly, adlockreadonly
    
    dim conn as adodb.connection
    set conn = currentproject.connection
    if not rs.eof then
        loadecotype = sqllookup(conn, "ECO_Type", "[tblOracleDropDowns]", "[ECO_Type_ID]=" & rs!change_order_type_id)
    end if
    
clean_exit:
    on error resume next
    if not rs is nothing then
        if rs.state = adstateopen then rs.close
    end if
    if not conn is nothing then
        if conn.state = adstateopen then conn.close
    end if
    set conn = nothing
    set rs = nothing
    set cn = nothing
    exit function

err_handler:
    call handleerror("modOracle", "findPartRev", err.description, err.number)
    resume clean_exit
end function

public function idnam(byval inputval as variant, byval typeval as variant) as variant
    on error resume next   ' preserve original behavior

    idnam = ""

    if nz(inputval, "") = "" then exit function

    dim cn as adodb.connection
    dim cmd as adodb.command
    dim rs as adodb.recordset

    set cn = getoracleconnection()

    set cmd = new adodb.command
    set cmd.activeconnection = cn
    cmd.commandtype = adcmdtext

    '========================================
    ' id ? return segment1
    '========================================
    if typeval = "ID" then
        cmd.commandtext = _
            "SELECT SEGMENT1 " & _
            "FROM INV.MTL_SYSTEM_ITEMS_B " & _
            "WHERE INVENTORY_ITEM_ID = ? " & _
            "AND ROWNUM = 1"

        cmd.parameters.append cmd.createparameter("pID", adinteger, adparaminput, , clng(inputval))
    end if

    '========================================
    ' nam ? return inventory_item_id
    '========================================
    if typeval = "NAM" then
        cmd.commandtext = _
            "SELECT INVENTORY_ITEM_ID " & _
            "FROM INV.MTL_SYSTEM_ITEMS_B " & _
            "WHERE SEGMENT1 = ? " & _
            "AND ROWNUM = 1"

        cmd.parameters.append cmd.createparameter("pSEG", advarchar, adparaminput, 50, cstr(inputval))
    end if

    set rs = cmd.execute

    if not rs.eof then
        if typeval = "ID" then
            idnam = nz(rs.fields("SEGMENT1").value, "")
        elseif typeval = "NAM" then
            idnam = nz(rs.fields("INVENTORY_ITEM_ID").value, "")
        end if
    end if

clean_exit:
    on error resume next

    if not rs is nothing then
        if rs.state = adstateopen then rs.close
    end if

    set rs = nothing
    set cmd = nothing
    set cn = nothing

end function

public function getdescriptionfromid(byval inventid as long) as string
    on error goto err_handler

    getdescriptionfromid = ""

    if isnull(inventid) then exit function

    on error resume next   ' preserve original behavior after initial check

    dim cn as adodb.connection
    dim cmd as adodb.command
    dim rs as adodb.recordset

    set cn = getoracleconnection()

    set cmd = new adodb.command
    set cmd.activeconnection = cn
    cmd.commandtype = adcmdtext
    cmd.commandtext = _
        "SELECT DESCRIPTION " & _
        "FROM INV.MTL_SYSTEM_ITEMS_B " & _
        "WHERE INVENTORY_ITEM_ID = ? " & _
        "AND ROWNUM = 1"

    cmd.parameters.append cmd.createparameter("pID", adinteger, adparaminput, , inventid)

    set rs = cmd.execute

    if not rs.eof then
        getdescriptionfromid = nz(rs.fields("DESCRIPTION").value, "")
    end if

clean_exit:
    on error resume next

    if not rs is nothing then
        if rs.state = adstateopen then rs.close
    end if

    set rs = nothing
    set cmd = nothing
    set cn = nothing

    exit function

err_handler:
    call handleerror("modOracle", "getDescriptionFromId", err.description, err.number)
    resume clean_exit
end function

public function getstatusfromid(byval inventid as long) as string
    on error goto err_handler

    getstatusfromid = ""

    if isnull(inventid) then exit function

    on error resume next   ' preserve original behavior after initial check

    dim cn as adodb.connection
    dim cmd as adodb.command
    dim rs as adodb.recordset

    set cn = getoracleconnection()

    set cmd = new adodb.command
    set cmd.activeconnection = cn
    cmd.commandtype = adcmdtext
    cmd.commandtext = _
        "SELECT INVENTORY_ITEM_STATUS_CODE " & _
        "FROM INV.MTL_SYSTEM_ITEMS_B " & _
        "WHERE INVENTORY_ITEM_ID = ? " & _
        "AND ROWNUM = 1"

    cmd.parameters.append cmd.createparameter("pID", adinteger, adparaminput, , inventid)

    set rs = cmd.execute

    if not rs.eof then
        getstatusfromid = nz(rs.fields("INVENTORY_ITEM_STATUS_CODE").value, "")
    end if

clean_exit:
    on error resume next

    if not rs is nothing then
        if rs.state = adstateopen then rs.close
    end if

    set rs = nothing
    set cmd = nothing
    set cn = nothing

    exit function

err_handler:
    call handleerror("modOracle", "getStatusFromId", err.description, err.number)
    resume clean_exit
end function
