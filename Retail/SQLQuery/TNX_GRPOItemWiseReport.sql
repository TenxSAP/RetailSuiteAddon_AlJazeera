CREATE PROCEDURE "TNX_GRPOItemWiseReport"(IN FromDate Date,IN ToDate DATE,IN BranchName VARCHAR(30))
As
Begin

Select * from 
(SELECT 'Goods Receipt PO' as "Type", T0."BPLName", T0."DocEntry",T0."DocNum", T0."DocType", T0."DocDate", T0."DocDueDate", T0."CardCode", T0."CardName", T0."NumAtCard", 
(Case when Row_number() over (Partition by T0."DocEntry" order by T1."VisOrder" ASC)=1 then T0."VatSum" else 0 end) as "Total Tax", 
(Case when Row_number() over (Partition by T0."DocEntry" order by T1."VisOrder" ASC)=1 then T0."DocTotal"+T0."DpmAmnt" else 0 end) as "Document Total", 
T0."Comments", T1."VisOrder" as "LineNum", T1."CodeBars", T1."ItemCode", T1."Dscription", T1."Quantity", T1."UomCode", T1."NumPerMsr" as "CF", (case when T0."DocType"='S' then T1."Price" else T1."INMPrice" end) as "Price", (case when T0."DocType"='S' then T1."LineTotal" else T1."Quantity"*T1."INMPrice" end) as "LineTotal",T1."VatSum" as "LineVatAmount" FROM OPDN T0  INNER JOIN PDN1 T1 ON T0."DocEntry" = T1."DocEntry" 
WHERE T0."DocDate" >=:FromDate and  T0."DocDate" <=:ToDate and  T0."BPLName" like '%' || REPLACE(:BranchName, '''', '''''') || '%' AND T0."CANCELED"='N'


UNION ALL

SELECT 'Goods Return' as "Type", T0."BPLName", T0."DocEntry",T0."DocNum", T0."DocType", T0."DocDate", T0."DocDueDate", T0."CardCode", T0."CardName", T0."NumAtCard", 
-(Case when Row_number() over (Partition by T0."DocEntry" order by T1."VisOrder" ASC)=1 then T0."VatSum" else 0 end) as "Total Tax", 
-(Case when Row_number() over (Partition by T0."DocEntry" order by T1."VisOrder" ASC)=1 then T0."DocTotal"+T0."DpmAmnt" else 0 end) as "Document Total", 
T0."Comments", T1."VisOrder" as "LineNum",T1."CodeBars", T1."ItemCode", T1."Dscription", T1."Quantity", T1."UomCode",T1."NumPerMsr" as "CF", (case when T0."DocType"='S' then T1."Price" else T1."INMPrice" end) as "Price", -(case when T0."DocType"='S' then T1."LineTotal" else T1."Quantity"*T1."INMPrice" end) as "LineTotal", -T1."VatSum" as "LineVatAmount" FROM ORPD T0  INNER JOIN RPD1 T1 ON T0."DocEntry" = T1."DocEntry" 
WHERE T0."DocDate" >=:FromDate and  T0."DocDate" <=:ToDate and  T0."BPLName" like '%' || REPLACE(:BranchName, '''', '''''') || '%' AND T0."CANCELED"='N') A 
ORDER BY A."Type", A."BPLName" ASC, A."DocEntry" ASC, A."LineNum" ASC;
END;