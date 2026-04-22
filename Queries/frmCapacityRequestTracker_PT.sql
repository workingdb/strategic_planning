SELECT
    cr.RecordID,
    cr.requestDate,
    ddsp.requestType,
    cr.SOP,
    cr.EOP,
    c.customerName,
    cr.Program,
    cr.requestor,
    cr.Notes,
    cr.responseTime
FROM
    (
        tblCapacityRequests AS cr
        LEFT JOIN tblDropDowns_StrategicPlanning AS ddsp ON cr.requestType = ddsp.recordId
    )
    LEFT JOIN tblCustomers AS c ON cr.Customer =                         c.ID WHERE EXISTS (SELECT 1 From tblCapacityRequest_partnumbers As cp WHERE cp.requestId = cr.recordId AND cp.capacityResults is null);
