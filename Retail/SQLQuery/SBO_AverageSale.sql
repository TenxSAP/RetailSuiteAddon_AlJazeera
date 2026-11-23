CREATE FUNCTION SBO_AverageSale()
RETURNS TABLE 
( 
    ITEMCODE NVARCHAR(215),
    ITEMNAME NVARCHAR(215),
    BARCODE NVARCHAR(215),
    ITEMARABICDESCRIPTION NVARCHAR(215),
    QUARTELYSALESQUANTITY DECIMAL(19,6),
    HALFYEARLYSALESQUANTITY DECIMAL(19,6),
    ANNUALSALESQUANTITY DECIMAL(19,6),
    MONTHLYAVERAGESALESQUANTITY DECIMAL(19,6)
)
LANGUAGE SQLSCRIPT
SQL SECURITY INVOKER 
AS
BEGIN

RETURN

SELECT
    CAST(T1."ItemCode" AS NVARCHAR(215)) AS "ITEMCODE",
    CAST(T1."ItemName" AS NVARCHAR(215)) AS "ITEMNAME",
    CAST(T1."CodeBars" AS NVARCHAR(215)) AS "BARCODE",
    CAST(T1."FrgnName" AS NVARCHAR(215)) AS "ITEMARABICDESCRIPTION",

    /* Quarterly Sales */
    CAST(
        COALESCE(
            (
                SELECT SUM(T3."InvQty")
                FROM "OINV" T2
                INNER JOIN "INV1" T3 ON T2."DocEntry" = T3."DocEntry"
                WHERE MONTH(T2."DocDate") BETWEEN 1 AND 4
                  AND YEAR(T2."DocDate") = YEAR(CURRENT_DATE)
                  AND T3."ItemCode" = T1."ItemCode" 
                  AND T3."CodeBars" = T1."CodeBars"
            ), 0
        ) AS DECIMAL(19,6)
    ) AS "QUARTELYSALESQUANTITY",

    /* Half-Yearly Sales */
    CAST(
        COALESCE(
            (
                SELECT SUM(T5."InvQty")
                FROM "OINV" T4
                INNER JOIN "INV1" T5 ON T4."DocEntry" = T5."DocEntry"
                WHERE MONTH(T4."DocDate") BETWEEN 1 AND 6
                  AND YEAR(T4."DocDate") = YEAR(CURRENT_DATE)
                  AND T5."ItemCode" = T1."ItemCode" 
                  AND T5."CodeBars" = T1."CodeBars"
            ), 0
        ) AS DECIMAL(19,6)
    ) AS "HALFYEARLYSALESQUANTITY",

    /* Annual Sales */
    CAST(
        COALESCE(
            (
                SELECT SUM(T7."InvQty")
                FROM "OINV" T6
                INNER JOIN "INV1" T7 ON T6."DocEntry" = T7."DocEntry"
                WHERE YEAR(T6."DocDate") = YEAR(CURRENT_DATE)
                  AND T7."ItemCode" = T1."ItemCode" 
                  AND T7."CodeBars" = T1."CodeBars"
            ), 0
        ) AS DECIMAL(19,6)
    ) AS "ANNUALSALESQUANTITY",

    /* Monthly Average */
    CAST(
        COALESCE(
            (
                COALESCE(
                    (
                        SELECT SUM(T9."InvQty")
                        FROM "OINV" T8
                        INNER JOIN "INV1" T9 ON T8."DocEntry" = T9."DocEntry"
                        WHERE T9."ItemCode" = T1."ItemCode" 
                          AND T9."CodeBars" = T1."CodeBars"
                    ), 0
                )
                /
                NULLIF(
                    MONTHS_BETWEEN(
                        (
                            SELECT MIN("DocDate")
                            FROM "OINM"
                            WHERE "ItemCode" = T1."ItemCode" 
                              AND "CodeBars" = T1."CodeBars"
                              AND "InQty" > 0
                        ),
                        CURRENT_DATE
                    ),
                0)
            ),
        0
        ) AS DECIMAL(19,6)
    ) AS "MONTHLYAVERAGESALESQUANTITY"

FROM "OITM" T1;

END;