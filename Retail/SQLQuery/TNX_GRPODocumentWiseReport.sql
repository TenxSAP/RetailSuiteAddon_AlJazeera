
CREATE PROCEDURE "TNX_GRPODocumentWiseReport"(IN FromDate Date,IN ToDate DATE,IN BranchName VARCHAR(30))
As
Begin
Select "Type", "BPLName", "DocEntry","DocNum", "DocStatus","DocType","DocDate", "DocDueDate", "Payment Term", "CardCode", "CardName","GroupName" as "Vendor Group", "NumAtCard", SUM("Total Tax") as "Total Tax", SUM("Document Total") as "Document Total","Comments" from 
(SELECT 'Goods Receipt PO' as "Type", T0."BPLName", T0."DocEntry",T0."DocNum", T0."DocStatus",T0."DocType", T0."DocDate", T0."DocDueDate", (SELECT "PymntGroup" FROM OCTG  WHERE "GroupNum"=T0."GroupNum") as "Payment Term",T0."CardCode", T0."CardName", T3."GroupName", T0."NumAtCard", 
(Case when Row_number() over (Partition by T0."DocEntry" order by T1."VisOrder" ASC)=1 then T0."VatSum" else 0 end) as "Total Tax", 
(Case when Row_number() over (Partition by T0."DocEntry" order by T1."VisOrder" ASC)=1 then T0."DocTotal"+T0."DpmAmnt" else 0 end) as "Document Total", 
T0."Comments", T1."VisOrder" as "LineNum", T1."CodeBars", T1."ItemCode", T1."Dscription", T1."Quantity", T1."UomCode", (case when T0."DocType"='S' then T1."Price" else T1."INMPrice" end) as "Price", (case when T0."DocType"='S' then T1."LineTotal" else T1."Quantity"*T1."INMPrice" end) as "LineTotal",T1."VatSum" as "LineVatAmount" 
FROM OPDN T0  INNER JOIN PDN1 T1 ON T0."DocEntry" = T1."DocEntry" 
INNER JOIN OCRD T2 on T2."CardCode"=T0."CardCode"
INNER JOIN OCRG T3 on T3."GroupCode"=T2."GroupCode"
WHERE T0."DocDate" >=:FromDate and  T0."DocDate" <=:ToDate and  T0."BPLName" like :BranchName AND T0."CANCELED"='N'


UNION ALL

SELECT 'Goods Return' as "Type", T0."BPLName", T0."DocEntry",T0."DocNum", T0."DocStatus",T0."DocType", T0."DocDate", T0."DocDueDate", (SELECT "PymntGroup" FROM OCTG  WHERE "GroupNum"=T0."GroupNum") as "Payment Term", T0."CardCode", T0."CardName", T3."GroupName",T0."NumAtCard", 
-(Case when Row_number() over (Partition by T0."DocEntry" order by T1."VisOrder" ASC)=1 then T0."VatSum" else 0 end) as "Total Tax", 
-(Case when Row_number() over (Partition by T0."DocEntry" order by T1."VisOrder" ASC)=1 then T0."DocTotal"+T0."DpmAmnt" else 0 end) as "Document Total", 
T0."Comments", T1."VisOrder" as "LineNum",T1."CodeBars", T1."ItemCode", T1."Dscription", T1."Quantity", T1."UomCode", (case when T0."DocType"='S' then T1."Price" else T1."INMPrice" end) as "Price", -(case when T0."DocType"='S' then T1."LineTotal" else T1."Quantity"*T1."INMPrice" end) as "LineTotal", -T1."VatSum" as "LineVatAmount" 
FROM ORPD T0  INNER JOIN RPD1 T1 ON T0."DocEntry" = T1."DocEntry" 
INNER JOIN OCRD T2 on T2."CardCode"=T0."CardCode"
INNER JOIN OCRG T3 on T3."GroupCode"=T2."GroupCode"
WHERE T0."DocDate" >=:FromDate and  T0."DocDate" <=:ToDate and  T0."BPLName" like :BranchName  AND T0."CANCELED"='N') A 
GROUP BY "Type", "BPLName", "DocEntry","DocNum", "DocStatus","DocType","DocDate", "DocDueDate","Payment Term", "CardCode", "CardName", "GroupName","NumAtCard","Comments"
ORDER BY A."Type" ,A."BPLName" ASC, A."DocEntry" ASC
;
End;