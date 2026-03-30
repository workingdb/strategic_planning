Option Compare Database
Option Explicit

Public g_cnOracle As ADODB.Connection

Public Function GetOracleConnection() As ADODB.Connection
    On Error GoTo Reconnect

    If Not g_cnOracle Is Nothing Then
        If g_cnOracle.State = adStateOpen Then
            Set GetOracleConnection = g_cnOracle
            Exit Function
        End If
    End If

Reconnect:
    Set g_cnOracle = New ADODB.Connection

    g_cnOracle.ConnectionString = "DRIVER={Oracle in OraClient11g_home1};SERVER=ebsprd1.world;UID=WorkingDB;PWD=WorkingDB2341;DBQ=ebsprd1.world;"
    g_cnOracle.Open

    Set GetOracleConnection = g_cnOracle
End Function

Public Function getStandardCostOwner(ByVal partNumber) As String
On Error GoTo Err_Handler

getStandardCostOwner = ""

If Nz(partNumber, "") = "" Then Exit Function

    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT msi.SEGMENT1 AS PN, mcv.SEGMENT3 AS Standard_Cost_Owner " & _
                "From inV.MTL_ITEM_CATEGORIES mic " & _
                    "JOIN INV.MTL_SYSTEM_ITEMS_B msi ON mic.INVENTORY_ITEM_ID = msi.INVENTORY_ITEM_ID " & _
                    "LEFT JOIN APPS.MTL_CATEGORIES_VL mcv ON mic.CATEGORY_ID = mcv.CATEGORY_ID " & _
                "where STRUCTURE_ID = 50349 AND msi.SEGMENT1 = '" & Replace(CStr(partNumber), "'", "''") & "' " & _
                "Group By msi.SEGMENT1,mcv.SEGMENT3;"

    Set cn = GetOracleConnection()

    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenForwardOnly, adLockReadOnly

    If Not rs.EOF Then
        getStandardCostOwner = Nz(rs.Fields("Standard_Cost_Owner").Value, "")
    End If

Clean_Exit:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    Set rs = Nothing
    Set cn = Nothing
    Exit Function

Err_Handler:
    Call handleError("modOracle", "getStandardCostOwner", err.Description, err.Number)
    Resume Clean_Exit
End Function

Public Function getCustomer(partNumber) As String
On Error GoTo Err_Handler

getCustomer = ""

If Nz(partNumber, "") = "" Then Exit Function

    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT C.CUSTOMER_NAME " & _
            "From APPS.XXCUS_CUSTOMERS C " & _
                "INNER JOIN INV.MTL_CUSTOMER_ITEMS ci ON c.CUSTOMER_ID = ci.CUSTOMER_ID " & _
                "INNER JOIN INV.MTL_CUSTOMER_ITEM_XREFS ic ON ic.CUSTOMER_ITEM_ID = ci.CUSTOMER_ITEM_ID " & _
                "INNER JOIN INV.MTL_SYSTEM_ITEMS_B si ON si.INVENTORY_ITEM_ID = ic.INVENTORY_ITEM_ID " & _
            "WHERE si.SEGMENT1 = '" & Replace(CStr(partNumber), "'", "''") & "' AND ci.INACTIVE_FLAG = 'N' " & _
            "Group By c.CUSTOMER_NAME;"

    Set cn = GetOracleConnection()

    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        getCustomer = getCustomer & Nz(rs.Fields("CUSTOMER_NAME").Value, "") & vbNewLine
        rs.MoveNext
    Loop
    
Clean_Exit:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    Set rs = Nothing
    Set cn = Nothing
    Exit Function

Err_Handler:
    Call handleError("modOracle", "getCustomer", err.Description, err.Number)
    Resume Clean_Exit
End Function

Public Function findCost(ByVal partNumber As String, ByVal costType As String, ByVal org As String) As Double
    On Error GoTo Err_Handler

    findCost = 0

    Dim db As DAO.Database
    Dim rsLocal As DAO.Recordset
    Dim orgID As Long

    '--- STEP 1: Get Org ID from Access table
    Set db = CurrentDb()
    Set rsLocal = db.OpenRecordset( _
        "SELECT ID FROM tblOrgs WHERE Org = '" & Replace(org, "'", "''") & "'", _
        dbOpenSnapshot)

    If rsLocal.EOF Then GoTo Clean_Exit

    orgID = rsLocal!ID

    rsLocal.Close
    Set rsLocal = Nothing
    Set db = Nothing

    '--- STEP 2: Query Oracle using ADODB
    Dim cn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset

    Set cn = GetOracleConnection()

    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cn
    cmd.CommandType = adCmdText

    cmd.CommandText = _
        "SELECT ITEM_COST " & _
        "FROM APPS.CST_ITEM_COST_TYPE_V " & _
        "WHERE ITEM_NUMBER = ? " & _
        "AND COST_TYPE = ? " & _
        "AND ORGANIZATION_ID = ?"

    ' Parameters (safer + faster)
    cmd.Parameters.Append cmd.CreateParameter("pItem", adVarChar, adParamInput, 50, partNumber)
    cmd.Parameters.Append cmd.CreateParameter("pCostType", adVarChar, adParamInput, 50, costType)
    cmd.Parameters.Append cmd.CreateParameter("pOrgID", adInteger, adParamInput, , orgID)

    Set rs = cmd.Execute

    If Not rs.EOF Then
        findCost = Nz(rs.Fields("ITEM_COST").Value, 0)
    End If

Clean_Exit:
    On Error Resume Next

    If Not rs Is Nothing Then If rs.State = adStateOpen Then rs.Close

    Set rs = Nothing
    Set cmd = Nothing
    Set cn = Nothing

    Exit Function

Err_Handler:
    Call handleError("modOracle", "findCost", err.Description, err.Number)
    Resume Clean_Exit
End Function

Public Function getCustomerName(customerId As Long) As String
On Error GoTo Err_Handler

getCustomerName = ""

If Nz(customerId, "") = "" Then Exit Function

    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT C.CUSTOMER_NAME " & _
            "From APPS.XXCUS_CUSTOMERS C " & _
            "WHERE c.CUSTOMER_ID = " & customerId & ";"

    Set cn = GetOracleConnection()

    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        getCustomerName = Nz(rs.Fields("CUSTOMER_NAME").Value, "")
    End If
    
Clean_Exit:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    Set rs = Nothing
    Set cn = Nothing
    Exit Function

Err_Handler:
    Call handleError("modOracle", "getCustomerName", err.Description, err.Number)
    Resume Clean_Exit
End Function

Public Function findDescription(ByVal partNumber As Variant) As String
    On Error GoTo Err_Handler

    findDescription = ""

    If Nz(partNumber, "") = "" Then Exit Function

    Dim cn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim partValue As String

    partValue = CStr(partNumber)
    Set cn = GetOracleConnection()

    '========================================
    ' 1) Check master items table first
    '========================================
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cn
    cmd.CommandType = adCmdText
    cmd.CommandText = _
        "SELECT DESCRIPTION " & _
        "FROM INV.MTL_SYSTEM_ITEMS_B " & _
        "WHERE SEGMENT1 = ? " & _
        "AND ROWNUM = 1"

    cmd.Parameters.Append cmd.CreateParameter("pPartNumber", adVarChar, adParamInput, 50, partValue)

    Set rs = cmd.Execute

    If Not rs.EOF Then
        findDescription = Nz(rs.Fields("DESCRIPTION").Value, "")
        GoTo Clean_Exit
    End If

    rs.Close
    Set rs = Nothing
    Set cmd = Nothing

    '========================================
    ' 2) Check all SIF tables at once
    '    using UNION ALL + ROWNUM
    '========================================
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cn
    cmd.CommandType = adCmdText
    cmd.CommandText = _
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

    cmd.Parameters.Append cmd.CreateParameter("pPart1", adVarChar, adParamInput, 50, partValue)
    cmd.Parameters.Append cmd.CreateParameter("pPart2", adVarChar, adParamInput, 50, partValue)
    cmd.Parameters.Append cmd.CreateParameter("pPart3", adVarChar, adParamInput, 50, partValue)

    Set rs = cmd.Execute

    If Not rs.EOF Then
        findDescription = Nz(rs.Fields("PART_DESCRIPTION").Value, "")
    End If

Clean_Exit:
    On Error Resume Next

    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If

    Set rs = Nothing
    Set cmd = Nothing
    Set cn = Nothing

    Exit Function

Err_Handler:
    Call handleError("modOracle", "findDescription", err.Description, err.Number)
    Resume Clean_Exit
End Function

Public Function findPartRev(partNumber) As String
On Error GoTo Err_Handler

findPartRev = "00"

If Nz(partNumber, "") = "" Then Exit Function

    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT Max(NEW_ITEM_REVISION) As REV " & _
            "From ENG.ENG_REVISED_ITEMS RI " & _
            "INNER JOIN INV.MTL_SYSTEM_ITEMS_B SI ON RI.REVISED_ITEM_ID = SI.INVENTORY_ITEM_ID " & _
            "WHERE ((RI.IMPLEMENTATION_DATE Is Not Null) AND (SI.SEGMENT1='" & Replace(CStr(partNumber), "'", "''") & "'));"

    Set cn = GetOracleConnection()

    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        findPartRev = Nz(rs.Fields("REV").Value, "00")
    End If
    
Clean_Exit:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    Set rs = Nothing
    Set cn = Nothing
    Exit Function

Err_Handler:
    Call handleError("modOracle", "findPartRev", err.Description, err.Number)
    Resume Clean_Exit
End Function

Function loadECOtype(changeNotice As String) As String
    On Error GoTo Err_Handler

loadECOtype = ""

If Nz(changeNotice, "") = "" Then Exit Function

    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT CHANGE_ORDER_TYPE_ID from ENG.ENG_ENGINEERING_CHANGES where CHANGE_NOTICE = '" & StrQuoteReplace(changeNotice) & "'"

    Set cn = GetOracleConnection()

    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenForwardOnly, adLockReadOnly
    
    Dim conn As ADODB.Connection
    Set conn = CurrentProject.Connection
    If Not rs.EOF Then
        loadECOtype = SqlLookup(conn, "ECO_Type", "[tblOracleDropDowns]", "[ECO_Type_ID]=" & rs!CHANGE_ORDER_TYPE_ID)
    End If
    
Clean_Exit:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
    Set conn = Nothing
    Set rs = Nothing
    Set cn = Nothing
    Exit Function

Err_Handler:
    Call handleError("modOracle", "findPartRev", err.Description, err.Number)
    Resume Clean_Exit
End Function

Public Function idNAM(ByVal inputVal As Variant, ByVal typeVal As Variant) As Variant
    On Error Resume Next   ' preserve original behavior

    idNAM = ""

    If Nz(inputVal, "") = "" Then Exit Function

    Dim cn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset

    Set cn = GetOracleConnection()

    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cn
    cmd.CommandType = adCmdText

    '========================================
    ' ID ? return SEGMENT1
    '========================================
    If typeVal = "ID" Then
        cmd.CommandText = _
            "SELECT SEGMENT1 " & _
            "FROM INV.MTL_SYSTEM_ITEMS_B " & _
            "WHERE INVENTORY_ITEM_ID = ? " & _
            "AND ROWNUM = 1"

        cmd.Parameters.Append cmd.CreateParameter("pID", adInteger, adParamInput, , CLng(inputVal))
    End If

    '========================================
    ' NAM ? return INVENTORY_ITEM_ID
    '========================================
    If typeVal = "NAM" Then
        cmd.CommandText = _
            "SELECT INVENTORY_ITEM_ID " & _
            "FROM INV.MTL_SYSTEM_ITEMS_B " & _
            "WHERE SEGMENT1 = ? " & _
            "AND ROWNUM = 1"

        cmd.Parameters.Append cmd.CreateParameter("pSEG", adVarChar, adParamInput, 50, CStr(inputVal))
    End If

    Set rs = cmd.Execute

    If Not rs.EOF Then
        If typeVal = "ID" Then
            idNAM = Nz(rs.Fields("SEGMENT1").Value, "")
        ElseIf typeVal = "NAM" Then
            idNAM = Nz(rs.Fields("INVENTORY_ITEM_ID").Value, "")
        End If
    End If

Clean_Exit:
    On Error Resume Next

    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If

    Set rs = Nothing
    Set cmd = Nothing
    Set cn = Nothing

End Function

Public Function getDescriptionFromId(ByVal inventId As Long) As String
    On Error GoTo Err_Handler

    getDescriptionFromId = ""

    If IsNull(inventId) Then Exit Function

    On Error Resume Next   ' preserve original behavior after initial check

    Dim cn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset

    Set cn = GetOracleConnection()

    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cn
    cmd.CommandType = adCmdText
    cmd.CommandText = _
        "SELECT DESCRIPTION " & _
        "FROM INV.MTL_SYSTEM_ITEMS_B " & _
        "WHERE INVENTORY_ITEM_ID = ? " & _
        "AND ROWNUM = 1"

    cmd.Parameters.Append cmd.CreateParameter("pID", adInteger, adParamInput, , inventId)

    Set rs = cmd.Execute

    If Not rs.EOF Then
        getDescriptionFromId = Nz(rs.Fields("DESCRIPTION").Value, "")
    End If

Clean_Exit:
    On Error Resume Next

    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If

    Set rs = Nothing
    Set cmd = Nothing
    Set cn = Nothing

    Exit Function

Err_Handler:
    Call handleError("modOracle", "getDescriptionFromId", err.Description, err.Number)
    Resume Clean_Exit
End Function

Public Function getStatusFromId(ByVal inventId As Long) As String
    On Error GoTo Err_Handler

    getStatusFromId = ""

    If IsNull(inventId) Then Exit Function

    On Error Resume Next   ' preserve original behavior after initial check

    Dim cn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset

    Set cn = GetOracleConnection()

    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cn
    cmd.CommandType = adCmdText
    cmd.CommandText = _
        "SELECT INVENTORY_ITEM_STATUS_CODE " & _
        "FROM INV.MTL_SYSTEM_ITEMS_B " & _
        "WHERE INVENTORY_ITEM_ID = ? " & _
        "AND ROWNUM = 1"

    cmd.Parameters.Append cmd.CreateParameter("pID", adInteger, adParamInput, , inventId)

    Set rs = cmd.Execute

    If Not rs.EOF Then
        getStatusFromId = Nz(rs.Fields("INVENTORY_ITEM_STATUS_CODE").Value, "")
    End If

Clean_Exit:
    On Error Resume Next

    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If

    Set rs = Nothing
    Set cmd = Nothing
    Set cn = Nothing

    Exit Function

Err_Handler:
    Call handleError("modOracle", "getStatusFromId", err.Description, err.Number)
    Resume Clean_Exit
End Function