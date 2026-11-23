CREATE PROCEDURE "TNX_ARInvoiceDocWiseReport"(IN FromDate Date,IN ToDate DATE,IN BranchName VARCHAR(30))
As
Begin

Select "Type", "BPLName", "DocEntry","DocNum", "DocType","DocDate", "DocDueDate", "CardCode", "CardName", "NumAtCard", SUM("Total Tax") as "Total Tax", SUM("Document Total") as "Document Total","Comments" from 
(SELECT 'AR Invoice' as "Type", T0."BPLName", T0."DocEntry",T0."DocNum", T0."DocType", T0."DocDate", T0."DocDueDate", T0."CardCode", T0."CardName", T0."NumAtCard", 
(Case when Row_number() over (Partition by T0."DocEntry" order by T1."VisOrder" ASC)=1 then T0."VatSum" else 0 end) as "Total Tax", 
(Case when Row_number() over (Partition by T0."DocEntry" order by T1."VisOrder" ASC)=1 then T0."DocTotal"+T0."DpmAmnt" else 0 end) as "Document Total", 
T0."Comments", T1."VisOrder" as "LineNum", T1."CodeBars", T1."ItemCode", T1."Dscription", T1."Quantity", T1."UomCode", (case when T0."DocType"='S' then T1."Price" else T1."INMPrice" end) as "Price", (case when T0."DocType"='S' then T1."LineTotal" else T1."Quantity"*T1."INMPrice" end) as "LineTotal",T1."VatSum" as "LineVatAmount" FROM OINV T0  INNER JOIN INV1 T1 ON T0."DocEntry" = T1."DocEntry" 
WHERE T0."DocDate" >=:FromDate and  T0."DocDate" <=:ToDate and  T0."BPLName" like '%' || REPLACE(:BranchName, '''', '''''') || '%' AND T0."CANCELED"='N'


UNION ALL

SELECT 'AR Credit Note' as "Type", T0."BPLName", T0."DocEntry",T0."DocNum", T0."DocType", T0."DocDate", T0."DocDueDate", T0."CardCode", T0."CardName", T0."NumAtCard", 
-(Case when Row_number() over (Partition by T0."DocEntry" order by T1."VisOrder" ASC)=1 then T0."VatSum" else 0 end) as "Total Tax", 
-(Case when Row_number() over (Partition by T0."DocEntry" order by T1."VisOrder" ASC)=1 then T0."DocTotal"+T0."DpmAmnt" else 0 end) as "Document Total", 
T0."Comments", T1."VisOrder" as "LineNum",T1."CodeBars", T1."ItemCode", T1."Dscription", T1."Quantity", T1."UomCode", (case when T0."DocType"='S' then T1."Price" else T1."INMPrice" end) as "Price", -(case when T0."DocType"='S' then T1."LineTotal" else T1."Quantity"*T1."INMPrice" end) as "LineTotal", -T1."VatSum" as "LineVatAmount" FROM ORIN T0  INNER JOIN RIN1 T1 ON T0."DocEntry" = T1."DocEntry" 
WHERE T0."DocDate" >=:FromDate and  T0."DocDate" <=:ToDate and  T0."BPLName" like '%' || REPLACE(:BranchName, '''', '''''') || '%' AND T0."CANCELED"='N') A 
GROUP BY "Type", "BPLName", "DocEntry","DocNum", "DocType","DocDate", "DocDueDate", "CardCode", "CardName", "NumAtCard", "Comments"
ORDER BY A."Type" DESC, A."BPLName" ASC, A."DocEntry" ASC;
END;