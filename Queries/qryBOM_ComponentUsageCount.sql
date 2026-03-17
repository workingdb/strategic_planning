SELECT qryBOM_DistinctParentComponent.ComponentItem, Count(*) AS AssyCount
FROM qryBOM_DistinctParentComponent
GROUP BY qryBOM_DistinctParentComponent.ComponentItem;

