CREATE PROCEDURE "TNX_ApprovalStatusDetailReport"
(
    IN AppID NVARCHAR(15),
    IN DocType NVARCHAR(30)
)
AS
BEGIN

SELECT DISTINCT
    T0."U_Doc" AS "Doc Type",
    T1."U_Userid" AS "Authorizer",
    T1."U_Stage" AS "Approval Stage",
    T0."DocNum" AS "DocEntry",
    CASE 
        WHEN T0."DataSource" = 'S' THEN 'ESS'
        ELSE 'SAP'
    END AS "Type",
    T0."U_docnum" AS "Application Num",
    T01."empID" AS "Requester SAP ID",

    CASE 
        WHEN T0."DataSource" = 'S' THEN 
            (SELECT "ExtEmpNo" FROM OHEM WHERE CAST("empID" AS VARCHAR) = T0."U_userid")
        ELSE 
            T0."U_userid"
    END AS "Requester Emp ID",

    CASE 
        WHEN T0."DataSource" = 'S' THEN 
            (SELECT 
                IFNULL("firstName",'') || ' ' || 
                IFNULL("middleName",'') || ' ' || 
                IFNULL("lastName",'')
             FROM OHEM 
             WHERE CAST("empID" AS VARCHAR)=T0."U_userid")
        ELSE 
            T0."U_userid"
    END AS "Requester Name",

    T3."BPLName" AS "Requester Branch",
    T0."U_docdate" AS "DocumentDate",
    T0."CreateDate" AS "Create Date",

    CASE 
        WHEN T1."U_statusN" IN ('R','N') THEN 'Rejected'
        WHEN T1."U_statusN" = 'S' THEN 'Submitted'
        WHEN T1."U_statusN" = 'A' THEN 'Approved'
        WHEN T1."U_statusN" = 'O' THEN 'Open'
        ELSE 'Cancel'
    END AS "Approval Doc Status",

    T1."U_Remarks" AS "Remarks",

    CASE 
        WHEN T1."U_statusN" IN ('R','N') THEN 'Rejected'
        WHEN T1."U_statusN" = 'S' THEN 'Submitted'
        WHEN T1."U_statusN" = 'A' THEN 'Approved'
        WHEN T1."U_statusN" = 'O' THEN 'Open'
        ELSE 'Cancel'
    END AS "Approval Stage Status",

    CASE 
        WHEN T1."U_Stage"='HR-STG' THEN 'HR Approver Stages'
        ELSE 'Line Manager Approver Stage'
    END AS "Approver Stage",

    T02."empID" AS "Approver SAP ID",

    (SELECT "ExtEmpNo" FROM OHEM WHERE CAST("empID" AS VARCHAR)=T1."U_EmpId") AS "Approver Emp ID",

    (SELECT 
        IFNULL("firstName",'') || ' ' || 
        IFNULL("middleName",'') || ' ' ||
        IFNULL("lastName",'')
     FROM OHEM 
     WHERE CAST("empID" AS VARCHAR)=T1."U_EmpId") AS "Approver Name",

    T1."U_TempId" AS "Template ID"

FROM "@SBO_DD" T0
INNER JOIN "@SBO_DD1" T1 ON T0."DocEntry" = T1."DocEntry"
LEFT JOIN OHEM T01 ON CAST(T01."empID" AS VARCHAR) = T0."U_userid"
LEFT JOIN OHEM T02 ON CAST(T02."empID" AS VARCHAR) = T1."U_EmpId"
LEFT JOIN HEM10 T2 ON CAST(T2."empID" AS VARCHAR) = T0."U_userid"
LEFT JOIN OBPL T3 ON T2."BPLId" = T3."BPLId"

WHERE 
    T0."U_statusN" <> 'R'
    AND T0."U_docnum" = :AppID
    AND T0."U_Doc" = :DocType

ORDER BY T0."DocEntry";

END;


