SELECT Count(r.RecordID) AS TotalRequests, Avg(r.responseTime) AS AvgResponseTime, Sum(IIf(Nz(s.OpenPartCount,0) > 0, 1, 0)) AS OpenRequests, Sum(IIf(Nz(s.OpenPartCount,0) = 0, 1, 0)) AS RespondedRequests
FROM tblCapacityRequests AS r LEFT JOIN (SELECT
            p.requestId,
            Sum(IIf(IsNull([capacityResults]) Or Trim([capacityResults] & "") = "", 1, 0)) AS OpenPartCount
        FROM tblCapacityRequest_partnumbers AS p
        GROUP BY p.requestId
    )  AS s ON r.RecordID = s.requestId
WHERE r.requestType = 1;

