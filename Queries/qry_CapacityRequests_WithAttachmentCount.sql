SELECT r.*, Nz(DCount("*","tblStratPlanAttachmentsSP",
        "referenceTable='tblCapacityRequests' AND referenceId=" & r.RecordID),0
    ) AS AttachmentCount
FROM tblCapacityRequests AS r;

