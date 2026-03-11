SELECT Count([RecordID]) AS TotalRequests, Avg([Responsetime]) AS AvgResponseTime, Sum(IIf([ResponseDate] Is Null,1,0)) AS OpenRequests, Sum(IIf([ResponseDate] Is Not Null,1,0)) AS RespondedRequests
FROM tblCapacityRequests
WHERE [RequestDate] >= TempVars!StartDate
  AND [RequestDate] < DateAdd("d", 1, TempVars!EndDate);

