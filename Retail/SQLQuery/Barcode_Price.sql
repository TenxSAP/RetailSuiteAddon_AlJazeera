CREATE PROCEDURE "Barcode_Price" (IN ItemCode varchar(100), IN UoM integer,IN BranchID Integer) 
language SQLScript
sql security invoker

AS 
Price Decimal;
RecordCount INTEGER;
BEGIN 
Price=0;
/*RecordCount=0;

SELECT Count(*) INTO RecordCount FROM OSPP T0 
INNER JOIN SPP1 T1 ON T0."ItemCode"=T1."ItemCode"
INNER JOIN SPP2 T2 ON T0."ItemCode"=T2."ItemCode"
Where T0."ItemCode" = :ItemCode And T0."Valid"='Y' And CURRENT_DATE Between T1."FromDate" and T1."ToDate"
And CURRENT_DATE Between T0."ValidFrom" and T0."ValidTo" ;

	If RecordCount = 0 Then*/
		
		SELECT IFNULL(
	            CASE 
	                WHEN (SELECT "IUoMEntry" FROM OITM WHERE "ItemCode" = :ItemCode) = :UoM THEN 
	                	(SELECT "Price" FROM ITM1 WHERE "ItemCode" = :ItemCode AND "PriceList" = 
	                    (SELECT T0."U_PricelistNo" FROM OBPL T0 WHERE T0."BPLId" = :BranchID)) 
	            ELSE 
	                (SELECT "Price" FROM ITM9 WHERE "ItemCode" = :ItemCode AND "PriceList" = 
	                (SELECT T0."U_PricelistNo" FROM OBPL T0 WHERE T0."BPLId" = :BranchID) AND "UomEntry" = :UoM) 
	            END
	        ,0) INTO Price
	    FROM DUMMY;
	    
/*	Else
		
		SELECT T1."Price" INTO Price  FROM OSPP T0 
		INNER JOIN SPP1 T1 ON T0."ItemCode"=T1."ItemCode"
		INNER JOIN SPP2 T2 ON T0."ItemCode"=T2."ItemCode"
		Where T0."ItemCode" = :ItemCode And T0."Valid"='Y' And CURRENT_DATE Between T1."FromDate" and T1."ToDate"
		And CURRENT_DATE Between T0."ValidFrom" and T0."ValidTo" ;
		
	End If;*/

Select :Price As "Price" from DUMMY;

END;