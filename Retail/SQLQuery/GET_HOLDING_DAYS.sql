ALTER PROCEDURE "GET_HOLDING_DAYS"
(
    IN itemCode NVARCHAR(50),
    IN BPLId INT
)
LANGUAGE SQLSCRIPT
AS
BEGIN

    SELECT 
        COALESCE(SUM(T1."InvQty" * G."BaseQty"), 0) AS "TotalQty"
    FROM "OINV" T0
    INNER JOIN "INV1" T1 
        ON T0."DocEntry" = T1."DocEntry"
    INNER JOIN "OITM" O 
        ON T1."ItemCode" = O."ItemCode"
    INNER JOIN "UGP1" G 
        ON O."UgpEntry" = G."UgpEntry" 
        AND T1."UomEntry" = G."UomEntry"
    WHERE T1."ItemCode" = :itemCode
      AND T0."BPLId" = :BPLId  
      AND T0."DocDate" >= ADD_DAYS(CURRENT_DATE, -30);

END;
