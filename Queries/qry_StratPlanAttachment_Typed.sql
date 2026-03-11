SELECT r.RecordID, r.*, a.referenceId, a.Title AS AttachmentTitle, a.directLink, a.Created AS AttachmentCreated
FROM tblCapacityRequests AS r LEFT JOIN tblStratPlanAttachmentsSP AS a ON r.RecordID = CLng(a.referenceId)
WHERE a.referenceTable = "tblCapacityRequests"
    OR a.referenceTable IS NULL;

