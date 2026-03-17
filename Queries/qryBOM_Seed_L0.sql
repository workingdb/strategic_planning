INSERT INTO tblBOM_Exploded ( ParentNAM, ParentNAM_ID, CurrentParent, CurrentParent_ID, ComponentItem, ComponentItem_ID, BOMLevel, DirectQty, ExtendedQty, CurrentParentType, ComponentItemType, PathText, SourceAssy, SourceCompt )
SELECT Q.Assy AS ParentNAM, Q.ASSEMBLY_ITEM_ID AS ParentNAM_ID, Q.Assy AS CurrentParent, Q.ASSEMBLY_ITEM_ID AS CurrentParent_ID, Q.Compt AS ComponentItem, Q.COMPONENT_ITEM_ID AS ComponentItem_ID, 0 AS BOMLevel, Q.Qty AS DirectQty, Q.Qty AS ExtendedQty, Q.assyItemType AS CurrentParentType, Q.compItemType AS ComponentItemType, '>' & CStr([Q].[ASSEMBLY_ITEM_ID]) & '>' & CStr([Q].[COMPONENT_ITEM_ID]) & '>' AS PathText, Q.Assy AS SourceAssy, Q.Compt AS SourceCompt
FROM qryBOM_L0 AS Q;

