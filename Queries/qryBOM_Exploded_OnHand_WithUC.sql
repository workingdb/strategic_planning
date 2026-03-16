SELECT Q.*, U.AssyCount, IIf(Nz(U.AssyCount,0)<=1,"Unique","Common") AS UniqueCommon
FROM qryBOM_Exploded_OnHand AS Q LEFT JOIN qryBOM_ComponentUsageCount AS U ON Q.ComponentItem = U.ComponentItem;

