CREATE PROCEDURE "Check_GRPO_Price_Fluctuation"
(
    IN DocEntry INT
)
LANGUAGE SQLSCRIPT
SQL SECURITY INVOKER
AS
BEGIN

    WITH "BarcodePerUOM" AS
    (
        SELECT
            X."ItemCode",
            X."UomEntry",
            X."BcdCode",
            X."BcdName",
            X."BcdEntry"
        FROM
        (
            SELECT
                B."ItemCode",
                B."UomEntry",
                B."BcdCode",
                B."BcdName",
                B."BcdEntry",

                ROW_NUMBER() OVER
                (
                    PARTITION BY
                        B."ItemCode",
                        B."UomEntry"
                    ORDER BY
                        B."BcdEntry"
                ) AS "RN"

            FROM "OBCD" B

            WHERE IFNULL(TRIM(B."BcdCode"), '') <> ''
        ) X

        WHERE X."RN" = 1
    ),

    "Data" AS
    (
        SELECT
            B.*,

            CASE
                WHEN ABS(B."MarginDiff") >= B."MarginTolerance"
                     AND B."MarginDiff" < 0
                THEN 'Price Increase'

                WHEN ABS(B."MarginDiff") >= B."MarginTolerance"
                     AND B."MarginDiff" > 0
                THEN 'Price Reduction'

                ELSE 'No Fluctuation'
            END AS "FluctuationType"

        FROM
        (
            SELECT
                A.*,

                /* OLD MARGIN */
                ROUND(
                    ("SellingPrice1" - "LastGRPOPrice"),
                    3
                ) AS "OldMarginAmount",

                /* OLD MARGIN % */
                CASE
                    WHEN IFNULL(
                            ROUND(
                                "SellingPrice1" /
                                (1 + IFNULL("SalesVatRate",0)/100),
                                3
                            ),
                            0
                         ) = 0
                    THEN 0

                    ELSE ROUND(
                            (
                                (
                                    ROUND(
                                        "SellingPrice1" /
                                        (1 + IFNULL("SalesVatRate",0)/100),
                                        3
                                    )
                                    - "LastGRPOPrice"
                                )
                                /
                                ROUND(
                                    "SellingPrice1" /
                                    (1 + IFNULL("SalesVatRate",0)/100),
                                    3
                                )
                            ) * 100,
                            3
                         )
                END AS "OldMarginPercentage",

                /* NEW MARGIN */
                ROUND(
                    ("SellingPrice1" - "ConvertedPrice"),
                    3
                ) AS "NewMarginAmount",

                /* NEW MARGIN % */
                CASE
                    WHEN IFNULL(
                            ROUND(
                                "SellingPrice1" /
                                (1 + IFNULL("SalesVatRate",0)/100),
                                3
                            ),
                            0
                         ) = 0
                    THEN 0

                    ELSE ROUND(
                            (
                                (
                                    ROUND(
                                        "SellingPrice1" /
                                        (1 + IFNULL("SalesVatRate",0)/100),
                                        3
                                    )
                                    - "LastGRPOPrice"
                                )
                                /
                                ROUND(
                                    "SellingPrice1" /
                                    (1 + IFNULL("SalesVatRate",0)/100),
                                    3
                                )
                            ) * 100,
                            3
                         )
                END AS "NewMarginPercentage",

                /* MARGIN DIFFERENCE */
                ROUND(
                    ROUND(
                        ("SellingPrice1" - "ConvertedPrice"),
                        3
                    )
                    -
                    ROUND(
                        ("SellingPrice1" - "LastGRPOPrice"),
                        3
                    ),
                    3
                ) AS "MarginDiff",

                /* PRICE WITH VAT */
                ROUND(
                    "ConvertedPrice"
                    +
                    (
                        "ConvertedPrice"
                        * IFNULL("VatRate",0)
                        / 100
                    ),
                    3
                ) AS "PriceWithVAT",

                /* LAST GRPO WITH VAT */
                ROUND(
                    "LastGRPOPrice"
                    +
                    (
                        "LastGRPOPrice"
                        * IFNULL("VatRate",0)
                        / 100
                    ),
                    3
                ) AS "LastGRPOPriceWithVAT",

                /* SELLING WITHOUT VAT */
                ROUND(
                    IFNULL(
                        (
                            SELECT "TNX_SELLING_PRICE"
                            (
                                "ItemCode",
                                "UomEntry",
                                "BPLId"
                            )
                            FROM DUMMY
                        ),
                        0
                    )
                    -
                    (
                        IFNULL(
                            (
                                SELECT "TNX_SELLING_PRICE"
                                (
                                    "ItemCode",
                                    "UomEntry",
                                    "BPLId"
                                )
                                FROM DUMMY
                            ),
                            0
                        )
                        * IFNULL("VatRate",0)
                        / 100
                    ),
                    3
                ) AS "SellingPriceWithoutVAT",

                /* SELLING PRICE */
                ROUND(
                    IFNULL(
                        (
                            SELECT "TNX_SELLING_PRICE"
                            (
                                "ItemCode",
                                "UomEntry",
                                "BPLId"
                            )
                            FROM DUMMY
                        ),
                        0
                    ),
                    6
                ) AS "SellingPrice",

                "ConvertedPrice" AS "Price"

            FROM
            (
                SELECT
                     H."DocEntry"
                    ,H."DocNum"
                    ,H."DocDate"

                    ,L."LineNum"
                    ,L."ItemCode"

                    /* DESCRIPTION */
                    ,B."BcdName" AS "Dscription"

                    /* RECEIPT UOM */
                    ,L."UomCode" AS "ReceiptUOM"
                    ,L."UomEntry" AS "ReceiptUomEntry"

                    /* ALL ITEM UOMS */
                    ,U."UomCode"
                    ,U."UomEntry"

                    ,G."BaseQty"

                    ,L."Quantity"

                    /* QTY AS PER UOM */
                    ,(
                        (L."NumPerMsr" * L."Quantity")
                        / G."BaseQty"
                     ) AS "UOMQty"

                    /* ORIGINAL PRICE */
                    ,L."Price" AS "DocumentPrice"

                    /* PRICE AS PER UOM */
                    ,ROUND(
                        (
                            L."Price"
                            / NULLIF(L."NumPerMsr",0)
                        )
                        * G."BaseQty",
                        6
                     ) AS "ConvertedPrice"

                    ,H."UserSign"

                    ,(
                        SELECT U1."USER_CODE"
                        FROM "OUSR" U1
                        WHERE U1."USERID" = H."UserSign"
                     ) AS "UserCode"

                    ,BP."U_PricelistNo"

                    ,H."BPLId"

                    ,BP."BPLName" AS "BranchCode"

                    /* SUPPLIER COST */
                    ,ROUND(
                        (
                            L."U_suppprice"
                            / NULLIF(L."NumPerMsr",0)
                        )
                        * G."BaseQty",
                        6
                     ) AS "SupplierCost"

                    /* LAST GRPO PRICE */
                    ,ROUND(
                        IFNULL(
                            (
                                SELECT "TNX_LASTGRPO_PRICE"
                                (
                                    L."ItemCode",
                                    L."UomEntry",
                                    H."DocDate",
                                    H."BPLId",
                                    H."DocEntry"
                                )
                                FROM DUMMY
                            ),
                            0
                        )
                        / NULLIF(L."NumPerMsr",0)
                        * G."BaseQty",
                        6
                     ) AS "LastGRPOPrice"

                    /* SELLING PRICE */
                    ,ROUND(
                        IFNULL(
                            (
                                SELECT "TNX_SELLING_PRICE"
                                (
                                    L."ItemCode",
                                    L."UomEntry",
                                    H."BPLId"
                                )
                                FROM DUMMY
                            ),
                            0
                        )
                        / NULLIF(L."NumPerMsr",0)
                        * G."BaseQty",
                        6
                     ) AS "SellingPrice1"

                    /* MARGIN TOLERANCE */
                    ,(
                        SELECT "U_MTlce"
                        FROM OADM
                     ) AS "MarginTolerance"

                    /* PURCHASE VAT */
                    ,T."Rate" AS "VatRate"

                    /* SALES VAT */
                    ,TS."Rate" AS "SalesVatRate"

                    /* BARCODE */
                    ,CASE
                        WHEN U."UomEntry" = L."UomEntry"
                             AND IFNULL(TRIM(L."CodeBars"), '') <> ''
                        THEN L."CodeBars"

                        ELSE B."BcdCode"
                     END AS "BarCode"

                    /* CODE BARS */
                    ,CASE
                        WHEN U."UomEntry" = L."UomEntry"
                             AND IFNULL(TRIM(L."CodeBars"), '') <> ''
                        THEN L."CodeBars"

                        ELSE B."BcdCode"
                     END AS "CodeBars"
                     /*,CASE
    WHEN U."UomEntry" = L."UomEntry"
         AND IFNULL(TRIM(L."CodeBars"), '') <> ''
    THEN CASE
             WHEN RIGHT(TRIM(L."CodeBars"), 2) = '-R'
             THEN LEFT(TRIM(L."CodeBars"), LENGTH(TRIM(L."CodeBars")) - 2)
             ELSE L."CodeBars"
         END
    ELSE CASE
             WHEN RIGHT(TRIM(B."BcdCode"), 2) = '-R'
             THEN LEFT(TRIM(B."BcdCode"), LENGTH(TRIM(B."BcdCode")) - 2)
             ELSE B."BcdCode"
         END
 END AS "BarCode"


 SAME BARCODE ALSO RETURNED AS CodeBars
   FOR EXISTING PRICE UPDATE LOGIC 
,CASE
    WHEN U."UomEntry" = L."UomEntry"
         AND IFNULL(TRIM(L."CodeBars"), '') <> ''
    THEN CASE
             WHEN RIGHT(TRIM(L."CodeBars"), 2) = '-R'
             THEN LEFT(TRIM(L."CodeBars"), LENGTH(TRIM(L."CodeBars")) - 2)
             ELSE L."CodeBars"
         END
    ELSE CASE
             WHEN RIGHT(TRIM(B."BcdCode"), 2) = '-R'
             THEN LEFT(TRIM(B."BcdCode"), LENGTH(TRIM(B."BcdCode")) - 2)
             ELSE B."BcdCode"
         END
 END AS "CodeBars" */

                    /* BARCODE SOURCE */
                    ,CASE
                        WHEN U."UomEntry" = L."UomEntry"
                             AND IFNULL(TRIM(L."CodeBars"), '') <> ''
                        THEN 'GRPO'

                        WHEN B."BcdCode" IS NOT NULL
                        THEN 'ITEM MASTER'

                        ELSE 'NO BARCODE'
                     END AS "BarcodeSource"

                FROM "OPDN" H

                INNER JOIN "PDN1" L
                        ON L."DocEntry" = H."DocEntry"

                INNER JOIN "OBPL" BP
                        ON BP."BPLId" = H."BPLId"

                INNER JOIN "OITM" I
                        ON I."ItemCode" = L."ItemCode"

                /* ALL ITEM UOMS */
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

                /* ONE BARCODE PER ITEM + UOM */
                LEFT JOIN "BarcodePerUOM" B
                       ON B."ItemCode" = I."ItemCode"
                      AND B."UomEntry" = U."UomEntry"

                WHERE H."DocEntry" = :DocEntry
                  AND IFNULL(H."CANCELED",'N') = 'N'
                  
                  /* EXCLUDE THE ENTIRE ROW IF BARCODE ENDS WITH -R */
  AND NOT
  ((U."UomEntry" = L."UomEntry"
          AND IFNULL(TRIM(L."CodeBars"), '') <> ''
          AND RIGHT(TRIM(L."CodeBars"), 2) = '-R')
      OR((U."UomEntry" <> L."UomEntry"
              OR IFNULL(TRIM(L."CodeBars"), '') = ''
          )
          AND RIGHT(TRIM(IFNULL(B."BcdCode", '')), 2) = '-R'))
          ) A
        ) B
    )

    SELECT
        D.*
    FROM "Data" D

    WHERE
        /* LAST GRPO PRICE MUST EXIST */
        D."LastGRPOPrice" <> 0

        /* CURRENT ROW MUST HAVE FLUCTUATION */
        AND ABS(D."MarginDiff") >= D."MarginTolerance"

        /* =====================================================
           IMPORTANT:
           ACTUAL GRPO UOM MUST HAVE FLUCTUATION

           Only after the Receipt UOM has fluctuation,
           other UOM fluctuations are allowed.
           ===================================================== */
        AND EXISTS
        (
            SELECT 1
            FROM "Data" P

            WHERE P."DocEntry" = D."DocEntry"
              AND P."LineNum" = D."LineNum"
              AND P."ItemCode" = D."ItemCode"

              /* This is the actual GRPO UOM */
              AND P."UomEntry" = P."ReceiptUomEntry"

              /* GRPO UOM itself must have fluctuation */
              AND P."LastGRPOPrice" <> 0
              AND ABS(P."MarginDiff") >= P."MarginTolerance"
        )

    ORDER BY
        D."ItemCode",
        CASE
            WHEN D."UomEntry" = D."ReceiptUomEntry"
            THEN 0
            ELSE 1
        END,
        D."UomEntry";

END;