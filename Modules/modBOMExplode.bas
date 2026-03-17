Option Compare Database
Option Explicit
 
Public Sub ExplodeBOM_AllLevels()
 
    Dim db As DAO.Database
    Dim rsSeed As DAO.Recordset
    Dim rsAdd As DAO.Recordset
    Dim sqlSeed As String
    Dim sqlNext As String
    Dim rowsAdded As Long
    Dim levelNum As Long
 
    On Error GoTo ErrHandler
 
    Set db = CurrentDb
 
    db.Execute "DELETE FROM tblBOM_Exploded;", dbFailOnError
 
    sqlSeed = "SELECT Assy, ASSEMBLY_ITEM_ID, Compt, COMPONENT_ITEM_ID, Qty, assyItemType, compItemType, assyDescription, compDescription " & _
              "FROM qryBOM_L0_Unique;"
 
    Set rsSeed = db.OpenRecordset(sqlSeed, dbOpenSnapshot)
    Set rsAdd = db.OpenRecordset("tblBOM_Exploded", dbOpenDynaset)
 
    Do While Not rsSeed.EOF
 
        rsAdd.AddNew
        rsAdd!ParentNAM = Nz(rsSeed!assy, "")
        rsAdd!ParentNAM_ID = Nz(rsSeed!ASSEMBLY_ITEM_ID, 0)
 
        rsAdd!CurrentParent = Nz(rsSeed!assy, "")
        rsAdd!CurrentParent_ID = Nz(rsSeed!ASSEMBLY_ITEM_ID, 0)
 
        rsAdd!ComponentItem = Nz(rsSeed!Compt, "")
        rsAdd!ComponentItem_ID = Nz(rsSeed!COMPONENT_ITEM_ID, 0)
 
        rsAdd!BOMLevel = 0
        rsAdd!Qty = Nz(rsSeed!Qty, 0)
        rsAdd!DirectQty = Nz(rsSeed!Qty, 0)
        rsAdd!ExtendedQty = Nz(rsSeed!Qty, 0)
 
        rsAdd!CurrentParentType = Nz(rsSeed!assyItemType, "")
        rsAdd!ComponentItemType = Nz(rsSeed!compItemType, "")
 
        rsAdd!ParentDescription = Nz(rsSeed!assyDescription, "")
        rsAdd!ComponentDescription = Nz(rsSeed!compDescription, "")
 
        rsAdd!PathText = ">" & Nz(rsSeed!ASSEMBLY_ITEM_ID, 0) & ">" & Nz(rsSeed!COMPONENT_ITEM_ID, 0) & ">"
 
        rsAdd!SourceAssy = Nz(rsSeed!assy, "")
        rsAdd!SourceCompt = Nz(rsSeed!Compt, "")
 
        rsAdd.Update
 
        rsSeed.MoveNext
    Loop
 
    rsSeed.Close
    rsAdd.Close
 
    levelNum = 0
 
    Do
 
        sqlNext = ""
        sqlNext = sqlNext & "INSERT INTO tblBOM_Exploded "
        sqlNext = sqlNext & "(ParentNAM, ParentNAM_ID, CurrentParent, CurrentParent_ID, ComponentItem, ComponentItem_ID, "
        sqlNext = sqlNext & "BOMLevel, Qty, DirectQty, ExtendedQty, CurrentParentType, ComponentItemType, ParentDescription, ComponentDescription, "
        sqlNext = sqlNext & "PathText, SourceAssy, SourceCompt) "
 
        sqlNext = sqlNext & "SELECT "
        sqlNext = sqlNext & "E.ParentNAM, "
        sqlNext = sqlNext & "E.ParentNAM_ID, "
        sqlNext = sqlNext & "Q.Assy AS CurrentParent, "
        sqlNext = sqlNext & "Q.ASSEMBLY_ITEM_ID AS CurrentParent_ID, "
        sqlNext = sqlNext & "Q.Compt AS ComponentItem, "
        sqlNext = sqlNext & "Q.COMPONENT_ITEM_ID AS ComponentItem_ID, "
        sqlNext = sqlNext & "(E.BOMLevel + 1) AS NewLevel, "
        sqlNext = sqlNext & "Q.Qty AS Qty, "
        sqlNext = sqlNext & "Q.Qty AS DirectQty, "
        sqlNext = sqlNext & "Round(E.ExtendedQty * Q.Qty,5) AS ExtendedQty, "
        sqlNext = sqlNext & "Q.assyItemType AS CurrentParentType, "
        sqlNext = sqlNext & "Q.compItemType AS ComponentItemType, "
        sqlNext = sqlNext & "Q.assyDescription AS ParentDescription, "
        sqlNext = sqlNext & "Q.compDescription AS ComponentDescription, "
        sqlNext = sqlNext & "E.PathText & CStr(Q.COMPONENT_ITEM_ID) & '>' AS PathText, "
        sqlNext = sqlNext & "Q.Assy AS SourceAssy, "
        sqlNext = sqlNext & "Q.Compt AS SourceCompt "
 
        sqlNext = sqlNext & "FROM tblBOM_Exploded AS E "
        sqlNext = sqlNext & "INNER JOIN qryBOM_L0_Unique AS Q "
        sqlNext = sqlNext & "ON E.ComponentItem_ID = Q.ASSEMBLY_ITEM_ID "
 
        sqlNext = sqlNext & "WHERE E.BOMLevel = " & levelNum & " "
        sqlNext = sqlNext & "AND E.ComponentItemType In ('COMPTM','FGM','SA','PH','COMPT') "
        sqlNext = sqlNext & "AND InStr(1,E.PathText,'>' & CStr(Q.COMPONENT_ITEM_ID) & '>') = 0;"
 
        db.Execute sqlNext, dbFailOnError
        rowsAdded = db.RecordsAffected
 
        levelNum = levelNum + 1
        If levelNum > 10 Then Exit Do
 
    Loop While rowsAdded > 0
 
ExitHere:
    On Error Resume Next
    rsSeed.Close
    rsAdd.Close
    Set rsSeed = Nothing
    Set rsAdd = Nothing
    Set db = Nothing
    Exit Sub
 
ErrHandler:
    MsgBox "Error in ExplodeBOM_AllLevels: " & Err.Number & " - " & Err.Description
    Resume ExitHere
 
End Sub