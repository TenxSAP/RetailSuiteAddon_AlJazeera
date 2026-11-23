CREATE FUNCTION "GET_COST"
(
    IN ITEMCODE VARCHAR(100),
    IN UOMCODE VARCHAR(100),
    IN PRICELIST VARCHAR(100)
)
RETURNS basic DOUBLE
LANGUAGE SQLSCRIPT
AS
BEGIN

    SELECT 
        IFNULL((
            SELECT 
                CASE 
                    WHEN X."UomEntry" = Y."BaseUom" THEN B."AvgPrice"
                    WHEN X."UomEntry" <> Y."BaseUom" THEN B."AvgPrice" * (1 / X."AltQty") * X."BaseQty"
                    ELSE 0 
                END
            FROM OITM C
            INNER JOIN (
                SELECT 
                    W."ItemCode",
                    W."AvgPrice",
                    W."WhsCode",
                    H."BPLid"
                FROM OITW W
                INNER JOIN OWHS H ON W."WhsCode" = H."WhsCode"
                WHERE W."AvgPrice" > 0
            ) B ON C."ItemCode" = B."ItemCode"
            LEFT JOIN OUGP Y ON Y."UgpEntry" = C."UgpEntry"
            LEFT JOIN UGP1 X ON Y."UgpEntry" = X."UgpEntry"
            WHERE 
                B."ItemCode" = :ITEMCODE
                AND X."UomEntry" = :UOMCODE
                AND B."BPLid" = (
                    SELECT MIN("BPLId") 
                    FROM OBPL 
                    WHERE "U_PricelistNo" = :PRICELIST
                )
            LIMIT 1
        ), 0)
    INTO basic
    FROM DUMMY;

END;