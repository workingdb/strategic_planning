SELECT DISTINCT CR.[Program]
FROM tblCapacityRequests AS CR
WHERE CR.[Program] Is Not Null

AND Trim(CR.[Program]) <> ""
ORDER BY CR.[Program];

