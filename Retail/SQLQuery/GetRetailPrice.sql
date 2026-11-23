CREATE PROCEDURE "GetRetailPrice"(IN Barcode NVARCHAR(30),IN Branch NVARCHAR(30)) LANGUAGE SQLSCRIPT SQL SECURITY INVOKER AS 
Begin SELECT
  (CASE WHEN T0."UomEntry"=T1."IUoMEntry" 
 THEN (SELECT
  A."Price" 
  from ITM1 A 
  where A."PriceList"= (SELECT
  T."U_PricelistNo" 
   FROM OBPL T 
   WHERE T."BPLId" =:Branch) 
  and A."ItemCode"=T0."ItemCode") 
 ELSE (SELECT
  A."Price" 
  from ITM9 A 
  where A."PriceList"=(SELECT
  T."U_PricelistNo" 
   FROM OBPL T 
   WHERE T."BPLId" =:Branch) 
  and A."ItemCode"=T0."ItemCode" 
  and A."UomEntry"=T0."UomEntry") 
 END) as "Price" 
FROM OBCD T0 
INNER JOIN OITM T1 on T0."ItemCode"=T1."ItemCode" 
where (T0."BcdCode") = :Barcode;  
End;

 