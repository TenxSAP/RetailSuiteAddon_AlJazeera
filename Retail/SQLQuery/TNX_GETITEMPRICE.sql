CREATE PROCEDURE "TNX_GETITEMPRICE"
(
    IN  Barcode  NVARCHAR(50),
    IN  UoM      INT,
    IN  ItemCode NVARCHAR(50)
)
LANGUAGE SQLSCRIPT
SQL SECURITY INVOKER
AS
BEGIN

    SELECT
        IFNULL(
            CASE
                WHEN (
                    SELECT "IUoMEntry"
                    FROM OITM
                    WHERE "ItemCode" = :ItemCode
                ) = :UoM
                THEN (
                    SELECT "Price"
                    FROM ITM1
                    WHERE "ItemCode" = :ItemCode
                      AND "PriceList" = 11
                )
                ELSE (
                    SELECT "Price"
                    FROM ITM9
                    WHERE "ItemCode" = :ItemCode
                      AND "PriceList" = 11
                      AND "UomEntry" = :UoM
                )
            END,
            0
        ) AS "Price"
    FROM DUMMY;

END;