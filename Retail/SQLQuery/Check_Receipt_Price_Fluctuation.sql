CREATE PROCEDURE "Check_Receipt_Price_Fluctuation"(IN DocEntry INT)
LANGUAGE SQLSCRIPT
SQL SECURITY INVOKER
AS
BEGIN

WITH "BarcodePerUOM" AS
(
    SELECT
         X."ItemCode"
        ,X."UomEntry"
        ,X."BcdCode"
        ,X."BcdName"
        ,X."BcdEntry"
    FROM
    (
        SELECT
             B."ItemCode"
            ,B."UomEntry"
            ,B."BcdCode"
            ,B."BcdName"
            ,B."BcdEntry"

            ,ROW_NUMBER() OVER
            (
                PARTITION BY
                     B."ItemCode"
                    ,B."UomEntry"
                ORDER BY
                    B."BcdEntry"
            ) AS "RN"

        FROM "OBCD" B

        WHERE IFNULL(TRIM(B."BcdCode"), '') <> ''
    ) X

    WHERE X."RN" = 1
)

SELECT *
,CASE 
    WHEN ABS("Difference") >= "MarginTolerance" AND "Difference" < 0 THEN 'Price Increase'
    WHEN ABS("Difference") >= "MarginTolerance" AND "Difference" > 0 THEN 'Price Reduction'
    ELSE 'No Fluctuation'
END As "FluctuationType"

FROM
(
    Select *
    ,Round(("OldSelling"-"OLDMAC"),3) As "OldMarginAmount"

    /* OLD MARGIN % */
    ,CASE
        WHEN IFNULL("OldSelling",0) = 0 THEN 0
        ELSE ROUND(
                (("OldSelling" - "OLDMAC") / "OldSelling") * 100
             ,3)
    END AS "OldMarginGP"

    ,Round(("OldSelling"-"NEWMAC"),3) As "NewMarginAmount"
    ,Round((("OldSelling"-"NEWMAC")-("OldSelling"-"OLDMAC")),3) "Difference"

    From
    (
        SELECT         
                 H."DocEntry"
                ,H."DocType"
                ,H."DocNum"
                ,H."DocDate"
                ,L."LineNum"
                ,L."ItemCode"
                ,L."Dscription"

                /* RECEIPT UOM */
                ,L."UomCode" AS "ReceiptUOM"

                /* ALL ITEM UOMS */
                ,U."UomCode"
                ,G."BaseQty"

                /* Converted Quantity */
                ,ROUND(
                    ((L."NumPerMsr" * L."Quantity") / NULLIF(G."BaseQty",0))
                ,3) AS "Quantity"

                /* Quantity as per UOM */
                ,ROUND(
                    ((L."NumPerMsr" * L."Quantity") / NULLIF(G."BaseQty",0))
                ,3) AS "UOMQty"

                ,L."Price"

                /* Price as per UOM */
                ,ROUND(
                    (L."Price" / NULLIF(L."NumPerMsr",0))
                    * G."BaseQty"
                ,6) AS "UOMPrice"

                ,H."UserSign"

                ,(
                    SELECT X."USER_CODE" 
                    FROM "OUSR" X 
                    WHERE X."USERID" = H."UserSign"
                 ) AS "UserCode"

                ,BP."U_PricelistNo"
                ,H."BPLId"
                ,BP."BPLName" AS "BranchCode"

                ,ROUND(
                    (IFNULL(L."U_suppprice",0) / NULLIF(L."NumPerMsr",0))
                    * G."BaseQty"
                ,6) AS "SupplierCost"

                /* OLD MAC */
                ,ROUND(
                    IFNULL(
                    (
                        SELECT A."BC"
                        FROM
                        (
                            SELECT
                            CASE
                                WHEN SUM(T0."InQty" - T0."OutQty")
                                     OVER
                                     (
                                        PARTITION BY T0."ItemCode"
                                        ORDER BY T0."DocDate",T0."CreatedBy"
                                        ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW
                                     ) <> 0
                                THEN
                                    T0."Balance"
                                    /
                                    SUM(T0."InQty" - T0."OutQty")
                                    OVER
                                    (
                                        PARTITION BY T0."ItemCode"
                                        ORDER BY T0."DocDate",T0."CreatedBy"
                                        ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW
                                    )
                                ELSE 0
                            END AS "BC",

                            ROW_NUMBER() OVER
                            (
                                ORDER BY T0."DocDate" DESC,
                                         T0."CreatedBy" DESC
                            ) AS "RN"

                            FROM OINM T0
                            WHERE T0."ItemCode" = L."ItemCode"
                            AND T0."Warehouse" = L."WhsCode"
                        ) A
                        WHERE A."RN" = 2
                    ),0)
                    * G."BaseQty"
                ,6) AS "OLDMAC"

                /* NEW MAC */
                ,ROUND(
                    IFNULL(
                    (
                        SELECT T0."AvgPrice"
                        FROM "OITW" T0
                        WHERE T0."ItemCode" = L."ItemCode"
                        AND T0."WhsCode" = L."WhsCode"
                    ),0)
                    * G."BaseQty"
                ,6) AS "NEWMAC"

                /* SELLING PRICE */
                ,ROUND(
                    (
                        IFNULL(
                        (
                            SELECT P."Price"
                            FROM OITM O
                            INNER JOIN UGP1 G1
                                ON O."UgpEntry" = G1."UgpEntry"
                            INNER JOIN OUOM U1
                                ON G1."UomEntry" = U1."UomEntry"
                            INNER JOIN ITM1 P
                                ON O."ItemCode" = P."ItemCode"
                            INNER JOIN OBCD B1
                                ON O."ItemCode" = B1."ItemCode"
                            WHERE B1."BcdCode" = I."CodeBars"
                            AND P."PriceList" = BP."U_PricelistNo"
                            AND U1."UomEntry" = L."UomEntry"
                        ),0)
                        / NULLIF(L."NumPerMsr",0)
                    )
                    * G."BaseQty"
                ,6) AS "OldSelling"

                ,(SELECT "U_MTlce" FROM OADM) AS "MarginTolerance"

                /* PURCHASE VAT */
                ,T."Rate" AS "VatRate"

                /* SALES VAT */
                ,TS."Rate" AS "SalesVatRate"

                ,I."CodeBars"
                ,U."UomEntry"

                /* =================================================
                   BARCODE

                   If this UOM is the UOM posted on the receipt:
                       1. Use IGN1.CodeBars if available
                       2. Otherwise use TOP 1 OBCD barcode

                   For all other UOMs:
                       Use TOP 1 OBCD barcode
                   ================================================= */
                ,CASE
                    WHEN U."UomEntry" = L."UomEntry"
                         AND IFNULL(TRIM(L."CodeBars"), '') <> ''
                    THEN L."CodeBars"

                    ELSE B."BcdCode"
                 END AS "BardCode"

        FROM "OIGN" H

        INNER JOIN "IGN1" L 
            ON L."DocEntry" = H."DocEntry"

        INNER JOIN "OBPL" BP 
            ON BP."BPLId" = H."BPLId"

        /* ITEM */
        INNER JOIN "OITM" I 
            ON I."ItemCode" = L."ItemCode"

        /* ALL UOMS OF ITEM */
        INNER JOIN "UGP1" G 
            ON I."UgpEntry" = G."UgpEntry"

        INNER JOIN "OUOM" U 
            ON U."UomEntry" = G."UomEntry"

        /* PURCHASE VAT */
        LEFT JOIN "OVTG" T 
            ON T."Code" = L."VatGroup"

        /* SALES VAT */
        LEFT JOIN "OVTG" TS 
            ON TS."Code" = I."VatGourpSa"

        /* =========================================================
           ONE BARCODE ONLY PER ITEM + UOM
           TOP 1 BY BcdEntry
           ========================================================= */
        LEFT JOIN "BarcodePerUOM" B
            ON B."ItemCode" = I."ItemCode"
           AND B."UomEntry" = U."UomEntry"

        WHERE H."DocEntry" = :DocEntry
          AND IFNULL(H."CANCELED",'N') = 'N'

        ORDER BY
            I."ItemCode",
            CASE 
                WHEN U."UomEntry" = L."UomEntry" THEN 0
                ELSE 1
            END,
            U."UomEntry"
    )
)
Where "Difference" <> '0';

END;