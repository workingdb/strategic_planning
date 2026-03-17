SELECT B.recordId, B.programId, P.OEM, P.modelName, P.modelCode
FROM tblBuildout_register_main AS B LEFT JOIN tblPrograms AS P ON B.programId = P.ID;

