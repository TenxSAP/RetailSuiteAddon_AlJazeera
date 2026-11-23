CREATE PROCEDURE "Barcode_ItemName" (IN ItemCode varchar(100), IN BarCode varchar(100)) 
language SQLScript
sql security invoker

AS 

BEGIN 

SELECT A."BcdName" 
/*|| '-' || CAST(
(SELECT CAST("BaseQty" AS integer) FROM UGP1  WHERE "UgpEntry" = (SELECT "UgpEntry" FROM OITM 
WHERE "ItemCode" = A."ItemCode") AND "UomEntry" = 
(SELECT "UomEntry"  FROM OBCD WHERE "BcdCode" = :BarCode)) AS varchar(30)) */

AS "ItemName" 
    FROM OBCD A 
    WHERE A."BcdCode" = :BarCode;
    
END;