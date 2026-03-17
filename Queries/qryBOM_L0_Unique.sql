SELECT First(T.Assy) AS Assy, First(T.Compt) AS Compt, Min(T.IMPLEMENTATION_DATE) AS IMPLEMENTATION_DATE, Min(T.Qty) AS Qty, First(T.Inverse_Qty) AS Inverse_Qty, First(T.assyDescription) AS assyDescription, First(T.assyStatus) AS assyStatus, First(T.assyItemType) AS assyItemType, First(T.compDescription) AS compDescription, First(T.compStatus) AS compStatus, First(T.compItemType) AS compItemType, First(T.COMPONENT_ITEM_ID) AS COMPONENT_ITEM_ID, First(T.ASSEMBLY_ITEM_ID) AS ASSEMBLY_ITEM_ID, Count(*) AS RowCopies
FROM qryBOM_L0 AS T
GROUP BY T.Assy, T.Compt;

