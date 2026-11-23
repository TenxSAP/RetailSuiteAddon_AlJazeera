CREATE PROCEDURE "TNX_ITEMLLIST"
(
    IN CardCode NVARCHAR(30)
)
LANGUAGE SQLSCRIPT
AS
BEGIN

    SELECT DISTINCT 
      T0."U_AppID"       AS "AppID",
    CASE 
            WHEN T0."U_Statu" = 'R' OR UPPER(T0."U_Statu") LIKE '%REJECT%' THEN 'Rejected'
            WHEN T0."U_Statu" = 'A' OR UPPER(T0."U_Statu") LIKE '%APPROV%' THEN 'Approved'
            WHEN T0."U_Statu" = 'N' OR UPPER(T0."U_Statu") LIKE '%CANCEL%' THEN 'Cancelled'
            WHEN T0."U_Statu" = 'S' OR UPPER(T0."U_Statu") LIKE '%SUBMIT%' THEN 'Submitted'
            WHEN T0."U_Statu" = 'P' OR UPPER(T0."U_Statu") LIKE '%PEND%' THEN 'Pending'
              WHEN T0."U_Statu" = 'O' OR UPPER(T0."U_Statu") LIKE '%OPEN%' THEN 'Open'
        END AS "Status",
      
        -- ? Business Partner
        T0."U_CardCode"    AS "CardCode",
        OCRD."CardName"    AS "CardName",

      
        -- ? Clean column names
       -- T0."U_BRANCH"      AS "Branch",
        T0."U_SERIES"      AS "Series",
       
          T0."DocNum"        AS "DocNum",
        T0."DocEntry"      AS "DocEntry"
        

    FROM "@SBO_LIST" T0 
    INNER JOIN "@SBO_LIST_C0" T1 
        ON T0."DocEntry" = T1."DocEntry"

    LEFT JOIN OCRD 
        ON OCRD."CardCode" = T0."U_CardCode"

    WHERE T0."U_CardCode" = :CardCode

    ORDER BY T0."DocEntry" DESC;

END;
 