CREATE PROCEDURE "GetDetails"
(
    IN ItemCode VARCHAR(100)
)
LANGUAGE SQLSCRIPT
SQL SECURITY INVOKER
AS
BEGIN

    SELECT
        T0."BcdCode"        AS "Barcode",
        T0."BcdName"        AS "BarcodeName",
        T0."UomEntry"       AS "UomEntry",
        T1."UomCode"        AS "UomCode",
        T0."U_STAX"         AS "Sugar Tax",
        T0."U_SUTAXAMT"     AS "Base Amount",
        ROUND(T0."U_ETAX", 3)         AS "Excise Tax",
        T0."U_CatalogNo"    AS "CatalogNo",
        T0."U_Price"        AS "Price",
        T0."U_ArbName"      AS "ArabicName",
        T0."U_PICK"         AS "PackingName",
        T0."U_POSCA"        AS "PosScan",
        T2."DocNum"         AS "Promotion Document Number"
    FROM "OBCD" T0
    INNER JOIN "OUOM" T1
        ON T0."UomEntry" = T1."UomEntry"
    LEFT JOIN "@TNX_ITEMDISC_C0" T3
        ON T0."BcdCode" = T3."U_BCODE"
       AND T3."U_Active" = 'Yes'
    LEFT JOIN "@TNX_ITEMDISC" T2
        ON T2."DocEntry" = T3."DocEntry"
       AND T2."U_STATUS" = 'Active'
    WHERE T0."ItemCode" = :ItemCode;

END;