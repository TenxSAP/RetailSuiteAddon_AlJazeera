CREATE PROCEDURE "TNX_StockCuntingSummrytobeposted"(IN FromDate Date,IN ToDate DATE,IN WHSE VARCHAR(30), IN GROUPNAME VARCHAR(30), IN SUBGROUP VARCHAR(30))
As
Begin

Select
	VV."CodeBars", VV."ItemCode",
	 VV."ItemName",
	 VV."Whse",
	--VV."Bin Code",
	VV."InvntryUom",
	 VV."ItmsGrpNam",  VV."U_VSubGroup",
	VV."Counted Qty",
	VV."Stock in selected Warehouse/Bin" ,
	VV."Counted Qty"-VV."Stock in selected Warehouse/Bin" as "Diff Qty",
	VV."Avgcost" , VV."LPP", 
(VV."Counted Qty"-VV."Stock in selected Warehouse/Bin" )* VV."Avgcost" as "Diff Value(By Avgcost)", (VV."Counted Qty"-VV."Stock in selected Warehouse/Bin" )* VV."LPP" as "Diff Value(By LPP)", 
	 VV."Current Sale Price" 
from (SELECT
	T1."CodeBars", T1."ItemCode",
	 T1."ItemName",
	 T5."U_WhseCode" as "Whse",
	--T5."U_BinCode" as "Bin Code",
	(Select "UomCode" from OUOM where T1."IUoMEntry"="UomEntry") as "InvntryUom",
	 T2."ItmsGrpNam",  T1."U_VSubGroup",
	 IFNULL(SUM((case when T1."IUoMEntry"=T3."U_UOMCode" 
				then T3."U_Qty" 
				else T3."U_Qty"* (SELECT
	 Y1."BaseQty"/Y1."AltQty" 
					FROM OUGP Y0 
					INNER JOIN UGP1 Y1 ON Y0."UgpEntry" = Y1."UgpEntry" 
					WHERE Y0."UgpEntry"=T1."UgpEntry" 
					and T3."U_UOMCode"=Y1."UomEntry") 
				end)),
	0) as "Counted Qty" ,
	  (select SUM("InQty"-"OutQty") from OINM T6 where T6."Warehouse"= :WHSE and T6."DocDate"< :FromDate and "ItemCode"=T1."ItemCode") as "Stock in selected Warehouse/Bin",
	 
	 (Select
	 "AvgPrice" 
		from OITW T6 
		where T6."ItemCode"=T1."ItemCode" 
		and T6."WhsCode"=:WHSE) as "Avgcost" ,

 IFNULL((Select (case when B."UomEntry"=X."UomEntry"  and X."UomEntry"=Y."BaseUom" then B."Price" when B."UomEntry"=X."UomEntry"  and X."UomEntry"<>Y."BaseUom" then B."Price"*1/X."AltQty"*X."BaseQty" when B."UomEntry"<>X."UomEntry" and X."UomEntry"<>Y."BaseUom" then B."Price"*1/X."AltQty"*X."BaseQty" when B."UomEntry"<>X."UomEntry" and X."UomEntry"=Y."BaseUom" then B."Price"else 0 end) as "LPP"
from OITM C inner join (SELECT A3."DocDate", A3."CardCode", A3."CardName",H."ItemCode",IFNULL((IFNULL(H."INMPrice", 0)/H."NumPerMsr"),0)-IFNULL((IFNULL(A1."INMPrice", 0)/A1."NumPerMsr"),0) as "Price", H."UomCode",H."UomEntry",H."NumPerMsr",Row_Number() over (partition by H."ItemCode" order by H."DocDate" Desc) as "RN" FROM "PDN1" H left join OPDN A3 on A3."DocEntry"=H."DocEntry" left join ORPC A2 on  A2."U_GRPO"=A3."DocNum"left join RPC1 A1 on A2."DocEntry"=A1."DocEntry" and A1."ItemCode"=H."ItemCode" and A1."NoInvtryMv"='Y' where IFNULL(H."U_normalqty",0)<>0)B on C."ItemCode"=B."ItemCode" left join OUGP Y on Y."UgpEntry"=C."UgpEntry"  left JOIN UGP1 X ON Y."UgpEntry" = X."UgpEntry" where "RN"=1 and B."ItemCode" = T1."ItemCode" and  X."UomEntry"= T1."IUoMEntry"),0) as "LPP" ,

	 (Select
	 "Price" 
		from ITM1 
		where "ItemCode"=T1."ItemCode" 
		and "PriceList"='1') as "Current Sale Price" 
	FROM OITM T1 
	INNER JOIN OITB T2 ON T1."ItmsGrpCod" = T2."ItmsGrpCod" 
	Left join "@ME_PSC_C0" T3 on T3."U_ItemCode"=T1."ItemCode" 
	LEft join "@ME_PSC" T5 on T5."DocEntry"=T3."DocEntry" 
	--LEFT JOIN OBIN T7 ON T7."BinCode"=T5."U_BinCode" 
	--LEFT JOIN OIBQ T6 on T6."WhsCode" = T5."U_WhseCode" 
--	and T6."ItemCode"=T1."ItemCode" 
	--and T6."BinAbs" = T7."AbsEntry" 
	Where T5."U_WhseCode" = :WHSE 
	and T2."ItmsGrpNam" like '%' || REPLACE(:GROUPNAME, '''', '''''') || '%'  and T1."U_VSubGroup" like '%' || REPLACE(:SUBGROUP, '''', '''''') || '%'
	and T1."ItemType"='I' 
	and T5."U_CountDate">=:FromDate
	and T5."U_CountDate"<=:ToDate 
	and T5."Status" = 'O' 
	GROUP BY T1."ItemCode",T1."CodeBars", 
	 T1."ItemName",
	T5."U_WhseCode",
	 T1."InvntryUom",T1."IUoMEntry",
	T2."ItmsGrpNam",  T1."U_VSubGroup"
--	T5."U_BinCode" 

	UNION ALL 

Select
	BB."CodeBars",  BB."ItemCode",
	 BB."ItemName",
	 T4."WhsCode" as "Whse",
	--T7."BinCode",
	 
	(Select "UomCode" from OUOM where BB."IUoMEntry"="UomEntry")  as "InvntryUom",
	 T."ItmsGrpNam",  BB."U_VSubGroup",
	 0 as "Counted Qty" ,
	 --T4."OnHand" 
 (select SUM("InQty"-"OutQty") from OINM T6 where T6."Warehouse"= :WHSE and T6."DocDate"<:FromDate and "ItemCode"=BB."ItemCode") 
		 as "Stock in selected Warehouse/Bin" ,
	 T4."AvgPrice" as "Avgcost" ,

IFNULL( (Select (case when B."UomEntry"=X."UomEntry"  and X."UomEntry"=Y."BaseUom" then B."Price" when B."UomEntry"=X."UomEntry"  and X."UomEntry"<>Y."BaseUom" then B."Price"*1/X."AltQty"*X."BaseQty" when B."UomEntry"<>X."UomEntry" and X."UomEntry"<>Y."BaseUom" then B."Price"*1/X."AltQty"*X."BaseQty" when B."UomEntry"<>X."UomEntry" and X."UomEntry"=Y."BaseUom" then B."Price"else 0 end) as "LPP"
from OITM C inner join (SELECT A3."DocDate", A3."CardCode", A3."CardName",H."ItemCode",IFNULL((IFNULL(H."INMPrice", 0)/H."NumPerMsr"),0)-IFNULL((IFNULL(A1."INMPrice", 0)/A1."NumPerMsr"),0) as "Price", H."UomCode",H."UomEntry",H."NumPerMsr",Row_Number() over (partition by H."ItemCode" order by H."DocDate" Desc) as "RN" FROM "PDN1" H left join OPDN A3 on A3."DocEntry"=H."DocEntry" left join ORPC A2 on  A2."U_GRPO"=A3."DocNum"left join RPC1 A1 on A2."DocEntry"=A1."DocEntry" and A1."ItemCode"=H."ItemCode" and A1."NoInvtryMv"='Y' where IFNULL(H."U_normalqty",0)<>0)B on C."ItemCode"=B."ItemCode" left join OUGP Y on Y."UgpEntry"=C."UgpEntry"  left JOIN UGP1 X ON Y."UgpEntry" = X."UgpEntry" where "RN"=1 and BB."ItemCode" = B."ItemCode" and  X."UomEntry"= BB."IUoMEntry"),0) as "LPP" ,

	 (Select
	 "Price" 
		from ITM1 
		where "ItemCode"=BB."ItemCode" 
		and "PriceList"='1') as "Current Sale Price" 
	from OITM BB 
	INNER JOIN OITB T ON BB."ItmsGrpCod" = T."ItmsGrpCod" 
	LEFT JOIN OITW T4 on T4."ItemCode"=BB."ItemCode" 
	and T4."WhsCode" = :WHSE 
	--LEFT JOIN OIBQ T6 on T6."WhsCode" = T4."WhsCode" 
	--and T6."ItemCode"=B."ItemCode" 
--	LEFT JOIN OBIN T7 ON T6."BinAbs" = T7."AbsEntry" 
	Where T4."OnHand"<>0 
	and T."ItmsGrpNam" like '%' || REPLACE(:GROUPNAME, '''', '''''') || '%'  and BB."U_VSubGroup" like '%' || REPLACE(:SUBGROUP, '''', '''''') || '%'
	and BB."ItemType"='I' 
	and BB."ItemCode" not in (SELECT
	 DISTINCT T1."ItemCode"
		FROM OITM T1 
		INNER JOIN OITB T2 ON T1."ItmsGrpCod" = T2."ItmsGrpCod" 
		left join "@ME_PSC_C0" T3 on T3."U_ItemCode"=T1."ItemCode" 
		left join "@ME_PSC" T5 on T5."DocEntry"=T3."DocEntry" 
		where T5."U_WhseCode" = :WHSE 
		and T2."ItmsGrpNam" like '%' || REPLACE(:GROUPNAME, '''', '''''') || '%'  and T1."U_VSubGroup" like '%' || REPLACE(:SUBGROUP, '''', '''''') || '%'
		and T1."ItemType"='I' 
		and T5."U_CountDate">=:FromDate 
		and T5."U_CountDate"<=:ToDate 
		and T5."Status" = 'O')) VV;
		END;