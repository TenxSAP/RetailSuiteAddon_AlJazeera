CREATE PROCEDURE "TNX_CategorywiseInventoryOB"(IN BranchName VARCHAR(30))
As
Begin

SELECT T2."U_VSubGroup",SUM( T1."Quantity") as "Quantity" , 
SUM(CASE WHEN T1."Quantity"<0 then T1."PostValueL" ELSE  T1."DocTotal" END) as "Total" 
FROM OIQI T0
INNER join IQI1 T1 ON T0."DocEntry" = T1."DocEntry"
INNER JOIN OITM T2 ON T2."ItemCode" = T1."ItemCode"
INNer Join OITB T3 ON T2."ItmsGrpCod" = T3."ItmsGrpCod"
where T0."BPLName" = :BranchName
Group by T2."U_VSubGroup";
END;