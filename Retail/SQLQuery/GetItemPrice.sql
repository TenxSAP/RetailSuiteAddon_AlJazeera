CREATE PROCEDURE "GetItemPrice" (
    IN ITEM NVARCHAR(30),
    IN Branch NVARCHAR(30)
)
LANGUAGE SQLSCRIPT
SQL SECURITY INVOKER
AS
BEGIN
    SELECT 
        CASE 
            WHEN UGP1."BaseQty" <> UGP1."AltQty"
                THEN OITW."AvgPrice" * UGP1."BaseQty"
            ELSE 
                OITW."AvgPrice"
        END AS "AvgPrice"
    FROM 
        OITW
    INNER JOIN 
        OITM ON OITW."ItemCode" = OITM."ItemCode"
    INNER JOIN 
        OITB ON OITM."ItmsGrpCod" = OITB."ItmsGrpCod"
    INNER JOIN 
        OWHS ON OITW."WhsCode" = OWHS."WhsCode"
    INNER JOIN 
        OBCD ON OITM."ItemCode" = OBCD."ItemCode"
    INNER JOIN 
        OUGP ON OITM."UgpEntry" = OUGP."UgpEntry"
    LEFT JOIN 
        UGP1 ON OUGP."UgpEntry" = UGP1."UgpEntry" 
             AND OBCD."UomEntry" = UGP1."UomEntry"
    INNER JOIN 
        OUOM ON OUOM."UomEntry" = OBCD."UomEntry"
    LEFT JOIN 
        OCRD ON OITM."CardCode" = OCRD."CardCode"

    WHERE 
        IFNULL(OITM."ItemCode", '') = :ITEM
        AND IFNULL(OITW."WhsCode", '') = (
            SELECT "DflWhs" 
            FROM OBPL 
            WHERE "BPLId" = :Branch
        );
END;