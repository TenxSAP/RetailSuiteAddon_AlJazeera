CREATE PROCEDURE TNX_RETAIL_TN
(
	in object_type nvarchar(30), 				-- SBO Object Type
	in transaction_type nchar(1),			-- [A]dd, [U]pdate, [D]elete, [C]ancel, C[L]ose
	in num_of_cols_in_key int,
	in list_of_key_cols_tab_del nvarchar(255),
	in list_of_cols_val_tab_del nvarchar(255)
)
LANGUAGE SQLSCRIPT
AS
-- Return values
error  int;				-- Result (0 for no error)
error_message nvarchar (200); 		-- Error string to be displayed
Cnt1 int;
Line int;
BlockDate Date;
begin

error := 0;
error_message := N'Ok';

--------------------------------------------------------------------------------------------------------------------------------

--	ADD	YOUR	CODE	HERE
-----------Customer Name mandatory--------
Cnt1:=0;							
   if :object_type='2' and (:transaction_type ='A' OR :transaction_type ='U') then  							
							
      Select Count(*) into Cnt1 from OCRD T0 Where T0."CardCode" = :list_of_cols_val_tab_del and IFNULL(T0."CardName",'')=''  ;							
 							
      if :Cnt1>0 then							
            error := -118;							
            error_message := 'BP Name cannot be empty';							
      end if;

      ---------------------------Mobile Number length -----------------  
    Select Count(*) into Cnt1 from OCRD T0 Where T0."CardCode" = :list_of_cols_val_tab_del and T0."CardType"='C' and T0."validFor"='Y' and
    T0."Country"='AE' and LEFT(T0."Cellular",1)='5' and LENGTH(T0."Cellular")<>9 ;							
 							
      if :Cnt1>0 then							
            error := -105;							
            error_message := 'Please enter valid Mobile number, Mobile Number length should be 9 for UAE.';							
      end if;	
---------------------------Mobile Number length -----------------  
    Select Count(*) into Cnt1 from OCRD T0 Where T0."CardCode" = :list_of_cols_val_tab_del and  T0."validFor"='Y' and
    T0."Country"='AE' and LEFT(T0."Cellular",1) NOT IN ('1','2','3','4','5','6','7','8','9')   ;							
 							
      if :Cnt1>0 then							
            error := -106;							
            error_message := 'Please enter valid Mobile number without starting 0. Please use Walk-in customer code for the transaction if no valid mobile number';							
      end if;				
      ---------------------------Mobile Number length -----------------  
      Select Count(*) into Cnt1 from OCRD T0 Where T0."CardCode" = :list_of_cols_val_tab_del and IFNULL(T0."Cellular",'')='' and T0."validFor"='Y' ;							
 							
      if :Cnt1>0 then							
            error := -174;							
            error_message := 'Mobile Number cannot be empty';							
      end if;	
      ---------------------------Mobile Number Duplication-----------------
      Select Count(*) into Cnt1 from OCRD T0 Where T0."CardCode" = :list_of_cols_val_tab_del and 
      T0."Cellular" in (select "Cellular" from OCRD where "CardCode"<>:list_of_cols_val_tab_del ) ;							
 							
      if :Cnt1>0 then							
            error := -170;							
            error_message := 'Mobile Number already exists';							
      end if;
      end if;
      
/* -------- Attachment mandatory for Supplier BP -------- */

IF :object_type = '2'
   AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN

    SELECT COUNT(*)
    INTO Cnt1
    FROM "OCRD" T0
    WHERE T0."CardCode" = :list_of_cols_val_tab_del
      AND T0."CardType" = 'S'
      AND (T0."AtcEntry" IS NULL OR T0."AtcEntry" = 0);

    IF :Cnt1 > 0 THEN
        error := -118;
        error_message := 'Attachment is mandatory for Supplier BP master data';
    END IF;

END IF;


----------------- GRPO Delivery document attachment-----------
IF :object_type = '20' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN

    DECLARE ATTACH_CNT INT;

    SELECT COUNT(*)
      INTO ATTACH_CNT
    FROM OPDN T0
    INNER JOIN OATC T1
        ON T0."AtcEntry" = T1."AbsEntry"
    INNER JOIN ATC1 T2
        ON T1."AbsEntry" = T2."AbsEntry"
    WHERE T0."DocEntry" = :list_of_cols_val_tab_del
      AND T2."FileName" LIKE '%' || T0."NumAtCard" || '%';

    IF ATTACH_CNT = 0 THEN
        error := 20;
        error_message := 'Vendor Ref No must match with at least one attachment file name';
    END IF;

END IF;



Cnt1:=0;							
   if :object_type='20' and (:transaction_type ='A' OR :transaction_type ='U') then 							
      Select Count(*) into Cnt1 from OPDN T0 INNER JOIN PDN1 T1 on T0."DocEntry" = T1."DocEntry" 
      INNER JOIN PDN12 T2 on T0."DocEntry"=T2."DocEntry"
      Where T0."DocEntry" = :list_of_cols_val_tab_del and IFNULL(T0."AtcEntry",0) = 0 ;					
      if :Cnt1>0 then							
            error := -136;							
            error_message := 'Document Attachment is Mandatory for GRPO';							
      end if;	
  end if; 
       
       -------------Item Name is mandatory--------------------
Cnt1:=0;							
   if :object_type='4' and (:transaction_type ='A' OR :transaction_type ='U') then 
      Select Count(*) into Cnt1 from OITM T0 Where T0."ItemCode" = :list_of_cols_val_tab_del and IFNULL(T0."ItemName",'')=''  ;							
 							
      if :Cnt1>0 then							
            error := -172;							
            error_message := 'Item Name cannot be empty';							
      end if;
       -------------Uom Group is manual-------------------
     Select Count(*) into Cnt1 from OITM T0 Where T0."ItemCode" = :list_of_cols_val_tab_del and T0."ItemType"<>'F' and T0."UgpEntry"='-1'  ;							
 							
      if :Cnt1>0 then							
            error := -173;							
            error_message := 'Please select appropriate Uom Group';							
      end if;
      ------------------ Item Master Barcode Validation-------------------      	
     /*Select Count(*) into Cnt1 from "OITM" T0 Where T0."ItemCode" = :list_of_cols_val_tab_del and 
     IFNULL(T0."CodeBars",'')='' AND IFNULL(T0."U_AcceptWMQty",'2') <> '1' ;
    if :Cnt1>0 then	
            error := -174;							
            error_message := 'Default BarCode field cannot be empty' ;							
    end if;*/
      -----------------Barcode duplication-------------
      Select Count(*) into Cnt1 from OITM T0 Inner join OBCD T1 on T0."ItemCode"=T1."ItemCode" 
      Where T0."ItemCode" = :list_of_cols_val_tab_del and T0."ItemType"<>'F' and
      (T1."BcdCode" in (Select distinct "BcdCode" from OBCD where "ItemCode" <> :list_of_cols_val_tab_del) OR
       
       T1."BcdCode" IN (SELECT "BcdCode" FROM OBCD WHERE "ItemCode" = :list_of_cols_val_tab_del GROUP BY "BcdCode" HAVING COUNT(*) > 1 ))  ;							
 							
      if :Cnt1>0 then
      		error := -171;							
            error_message :='Barcode already exists with another ItemCode or Duplicate barcode exists for same item' ;							
      end if;
    end if;
  /*  -----------------------------
  IF :object_type IN ('1470000113', '22', '20', '18') 
AND :transaction_type IN ('A', 'U') THEN
 
    DECLARE v_count INT;
 
    SELECT COUNT(*) INTO v_count
    FROM (
        -- PR
        SELECT 1
        FROM PRQ1 T1
        INNER JOIN OITW W ON T1."ItemCode" = W."ItemCode"
                         AND T1."WhsCode" = W."WhsCode"
        WHERE :object_type = '1470000113'
          AND W."U_BlockPurchase" = 'Y'
          AND T1."DocEntry" = :list_of_cols_val_tab_del
 
        UNION ALL
 
        -- PO
        SELECT 1
        FROM POR1 T1
        INNER JOIN OITW W ON T1."ItemCode" = W."ItemCode"
                         AND T1."WhsCode" = W."WhsCode"
        WHERE :object_type = '22'
          AND W."U_BlockPurchase" = 'Y'
          AND T1."DocEntry" = :list_of_cols_val_tab_del
 
        UNION ALL
 
        -- GRPO
        SELECT 1
        FROM PDN1 T1
        INNER JOIN OITW W ON T1."ItemCode" = W."ItemCode"
                         AND T1."WhsCode" = W."WhsCode"
        WHERE :object_type = '20'
          AND W."U_BlockPurchase" = 'Y'
          AND T1."DocEntry" = :list_of_cols_val_tab_del
 
        UNION ALL
 
        -- AP Invoice
        SELECT 1
        FROM PCH1 T1
        INNER JOIN OITW W ON T1."ItemCode" = W."ItemCode"
                         AND T1."WhsCode" = W."WhsCode"
        WHERE :object_type = '18'
          AND W."U_BlockPurchase" = 'Y'
          AND T1."DocEntry" = :list_of_cols_val_tab_del
    ) X;
 
    IF v_count > 0 THEN
        error := 50001;
        error_message := 
        'Purchase is blocked for this item in the selected warehouse. Please check Item Master → Inventory Data.';
    END IF;
 
END IF;*/
-------------------------Purchases are blocked for this branch for item in Line-------------------
Cnt1:=0;
Line:=0;							
   if :object_type='22' and (:transaction_type ='A' OR :transaction_type ='U') then  							
      Select Count(*) into Cnt1 from "OPOR" T0 INNER JOIN "POR1" T1 on T0."DocEntry"=T1."DocEntry" 
      Where T0."DocEntry" = :list_of_cols_val_tab_del and (Select "U_BlockPurchase" from OITW 
      where "ItemCode" = T1."ItemCode" and "WhsCode"=T1."WhsCode")='Yes';
    if :Cnt1>0 then	
      Select Top 1 T1."VisOrder"+1 into Line from "OPOR" T0 INNER JOIN "POR1" T1 on T0."DocEntry"=T1."DocEntry" 
      Where T0."DocEntry" = :list_of_cols_val_tab_del and (Select "U_BlockPurchase" from OITW 
      where "ItemCode" = T1."ItemCode" and "WhsCode"=T1."WhsCode")='Yes';
            error := -139;							
            error_message := 'Purchases are blocked for this branch for item in Line '|| :Line ;							
    end if;
end if;  
-------------------------GRPO are blocked for this branch for item in Line-------------------
Cnt1:=0;
Line:=0;							
   if :object_type='20' and (:transaction_type ='A' OR :transaction_type ='U') then  							
      Select Count(*) into Cnt1 from "OPDN" T0 INNER JOIN "PDN1" T1 on T0."DocEntry"=T1."DocEntry" 
      Where T0."DocEntry" = :list_of_cols_val_tab_del and (Select "U_BlockPurchase" from OITW 
      where "ItemCode" = T1."ItemCode" and "WhsCode"=T1."WhsCode")='Yes';
    if :Cnt1>0 then	
      Select Top 1 T1."VisOrder"+1 into Line from "OPOR" T0 INNER JOIN "POR1" T1 on T0."DocEntry"=T1."DocEntry" 
      Where T0."DocEntry" = :list_of_cols_val_tab_del and (Select "U_BlockPurchase" from OITW 
      where "ItemCode" = T1."ItemCode" and "WhsCode"=T1."WhsCode")='Yes';
            error := -140;							
            error_message := 'Purchases are blocked for this branch for item in Line '|| :Line ;							
    end if;
end if; 

----------------------- Outlet Amount Validation (Madhu) -----------------------
IF :object_type='SBO_OUTLETUDO' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN
 
    DECLARE HeaderAmt DECIMAL(19,6);
    DECLARE SplitTotal DECIMAL(19,6);
 
    -- Get Header Amount from @SBO_OUTLET
    SELECT IFNULL("U_AUT", 0)
    INTO HeaderAmt
    FROM "@SBO_OUTLET"
    WHERE "DocEntry" = :list_of_cols_val_tab_del;
 
    -- Get total Amount Splitting from @SBO_OUTLET_C0
    SELECT IFNULL(SUM("U_AmntSpt"), 0)
    INTO SplitTotal
    FROM "@SBO_OUTLET_C0"
    WHERE "DocEntry" = :list_of_cols_val_tab_del;
 
    -- Compare header vs detail split total

   	IF :HeaderAmt <> :SplitTotal AND ((:HeaderAmt - :SplitTotal) >0.01) THEN
        error := -309002;
        error_message := 'Amount Mismatch: Header Amount (' || :HeaderAmt ||') does not match Amount Splitting total (' || :SplitTotal || ').';
        --error_message :='Test';
    END IF;
 
END IF;
----------------------- End of Outlet Amount Validation ------------------------

----------------------- SBO_LISTUDO Branch Count Validation -----------------------
IF :object_type = 'SBO_LISTUDO'
   AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN

    DECLARE BranchCnt INT;
    DECLARE InvalidCnt INT;

    -- Get total number of branches from OBPL
    SELECT COUNT(*)
    INTO BranchCnt
    FROM OBPL
    WHERE IFNULL("Disabled", 'N') = 'N';   -- only active branches

    -- Check if U_NOU exceeds branch count
    SELECT COUNT(*)
    INTO InvalidCnt
    FROM "@SBO_LIST_C0"
    WHERE "DocEntry" = :list_of_cols_val_tab_del
   ---   AND IFNULL("U_NOU", '0')
and IFNULL(CAST(NULLIF("U_NOU",'') AS INTEGER),0) 
       > :BranchCnt;

    IF :InvalidCnt > 0 THEN
        error := -309004;
        error_message :=
            'Validation Error: No of Outlets cannot be greater than total Branches (' 
            || :BranchCnt || ').';
    END IF;

END IF;
----------------------- End Branch Count Validation -----------------------

----------------------- SBO_OUTLETUDO Outlet vs Amount Validation -----------------------
----------------------- End Outlet vs Amount Validation -----------------------

----------------- SKU Count Validation category split scree (Header Based) -----------------
/*IF (:object_type = 'TNX_CATUDO') 
   AND (:transaction_type = 'A' OR :transaction_type = 'U')
THEN

    DECLARE v_HeaderSKU INT := 0;
    DECLARE v_TotalChildSKU INT := 0;

    -- Header SKU
    SELECT COALESCE(TO_INT("U_SKUS"),0)
    INTO v_HeaderSKU
    FROM "@TNX_CATERY"
    WHERE "DocEntry" = :list_of_cols_val_tab_del;

    -- Child Sum
    SELECT COALESCE(SUM(
        CASE 
            WHEN "U_SKUCount" IS NULL OR "U_SKUCount" = '' 
            THEN 0
            ELSE TO_INT("U_SKUCount")
        END
    ),0)
    INTO v_TotalChildSKU
    FROM "@TNX_CATERY_C0"
    WHERE "DocEntry" = :list_of_cols_val_tab_del;

    IF :v_TotalChildSKU <> :v_HeaderSKU THEN
        error := -5002;
        error_message := 
        'Total SKU Count must be ' || :v_HeaderSKU ||
        '. Current total is ' || :v_TotalChildSKU || '.';
    END IF;

END IF;

*/
----------------- End SKU Validation -----------------

----------------- SKU Count Validation (Decimal Aware) -----------------
IF (:object_type = 'TNX_CATUDO') 
   AND (:transaction_type = 'A' OR :transaction_type = 'U')
THEN
    DECLARE v_HeaderSKU INT := 0;
    DECLARE v_TotalChildSKU INT := 0;

    -- 1. Get Header SKU (Handles "8" or "8.000")
    SELECT COALESCE(
        CASE 
            WHEN "U_SKUS" IS NOT NULL AND OCCURRENCES_REGEXPR('^[0-9.]+$' IN TRIM("U_SKUS")) > 0 
            THEN TO_INT(TO_DOUBLE(TRIM("U_SKUS"))) 
            ELSE 0 
        END, 0)
    INTO v_HeaderSKU
    FROM "@TNX_CATERY"
    WHERE "DocEntry" = :list_of_cols_val_tab_del;

    -- 2. Use a Table Variable to isolate and clean decimal strings
    -- The Regex now includes the dot [0-9.]
    temp_table = SELECT 
                    CASE 
                        WHEN OCCURRENCES_REGEXPR('^[0-9.]+$' IN TRIM("U_SKUCount")) > 0 
                        THEN TRIM("U_SKUCount") 
                        ELSE '0' 
                    END AS CleanValue
                 FROM "@TNX_CATERY_C0"
                 WHERE "DocEntry" = :list_of_cols_val_tab_del;

    -- 3. Sum the filtered values (Converting string -> double -> int)
    SELECT SUM(TO_INT(TO_DOUBLE(CleanValue))) 
    INTO v_TotalChildSKU 
    FROM :temp_table;

    -- 4. Final Comparison
    IF :v_TotalChildSKU <> :v_HeaderSKU THEN
        error := -5002;
        error_message := 'Total SKU Count must be ' || CAST(:v_HeaderSKU AS NVARCHAR) || 
                         '. Current total is ' || CAST(:v_TotalChildSKU AS NVARCHAR) || '.';
    END IF;

END IF;
----------------- End SKU Validation -----------------


-----------------Start validation for Items List screen based of no of skus cannot add greather than that value-------------
----------------- Start SKU Validation -----------------
IF :object_type = 'SBO_LSUDO'
   AND (:transaction_type = 'A' OR :transaction_type = 'U')
THEN

    DECLARE HeaderSKUCount INT DEFAULT 0;
    DECLARE ItemCodeCount INT DEFAULT 0;

    -- Get SKU count from Header
    SELECT IFNULL("U_SKUS", 0)
    INTO HeaderSKUCount
    FROM "@SBO_LS"
    WHERE "DocEntry" = :list_of_cols_val_tab_del;

    -- Count ONLY valid ItemCodes (ignore LineId gaps)
    SELECT COUNT("U_ICode")
    INTO ItemCodeCount
    FROM "@SBO_LS_C0"
    WHERE "DocEntry" = :list_of_cols_val_tab_del
      AND IFNULL("U_ICode", '') <> '';

    -- Validation: Exact match based on ItemCode
    IF :HeaderSKUCount <> :ItemCodeCount THEN
        error := -309003;
        error_message :=
            'SKU Validation Failed: You must add exactly '
            || :HeaderSKUCount
            || ' ItemCodes. Currently added: '
            || :ItemCodeCount;
    END IF;

END IF;
----------------- End SKU Validation -----------------
------------------------------------------------------------------------------------------------------------------------
--------------------------------------------Vendor reference Number  Mandatory and not Duplicate grpo-------------------
 Cnt1:=0;
   if :object_type='20' and (:transaction_type ='A' OR :transaction_type ='U') then  							
      Select Count(*) into Cnt1 from "OPDN" T0 Where T0."DocEntry" = :list_of_cols_val_tab_del and 
      T0."CANCELED"='N' and (( IFNULL(T0."NumAtCard",'')<>'' and UPPER(T0."NumAtCard") in (Select distinct UPPER("NumAtCard") from OPDN where "DocEntry" <> :list_of_cols_val_tab_del and "CANCELED"='N' and "CardCode" = T0."CardCode"))
      OR (IFNULL(T0."NumAtCard",'')=''));
    if :Cnt1>0 then	
            error := -2060;							
            error_message := 'Vendor Ref No. is Empty or duplicated' ;							
    end if;
end if;
----------------------------------------------------------------------------------------------------------------
/*--------------------------------------------Vendor reference Number  Mandatory and not Duplicate Purchase order------------------
 Cnt1:=0;
   if :object_type='22' and (:transaction_type ='A' OR :transaction_type ='U') then  							
      Select Count(*) into Cnt1 from "OPOR" T0 Where T0."DocEntry" = :list_of_cols_val_tab_del and 
      T0."CANCELED"='N' and (( IFNULL(T0."NumAtCard",'')<>'' and UPPER(T0."NumAtCard") in (Select distinct UPPER("NumAtCard") from OPOR where "DocEntry" <> :list_of_cols_val_tab_del and "CANCELED"='N' and "CardCode" = T0."CardCode"))
      OR (IFNULL(T0."NumAtCard",'')=''));
    if :Cnt1>0 then	
            error := -2260;							
            error_message := 'Vendor Ref No. is Empty or duplicated' ;							
    end if;
end if;
----------------------------------------------------------------------------------------------------------------*/
--------------------------------------------Vendor reference Number  Mandatory and not Duplicate Ap invoice-------------------
 Cnt1:=0;
   if :object_type='18' and (:transaction_type ='A' OR :transaction_type ='U') then  							
      Select Count(*) into Cnt1 from "OPCH" T0 Where T0."DocEntry" = :list_of_cols_val_tab_del and 
      T0."CANCELED"='N' and (( IFNULL(T0."NumAtCard",'')<>'' and UPPER(T0."NumAtCard") in (Select distinct UPPER("NumAtCard") from OPCH where "DocEntry" <> :list_of_cols_val_tab_del and "CANCELED"='N' and "CardCode" = T0."CardCode"))
      OR (IFNULL(T0."NumAtCard",'')=''));
    if :Cnt1>0 then	
            error := -1860;							
            error_message := 'Vendor Ref No. is Empty or duplicated' ;							
    end if;
end if;
----------------------------------------------------------------------------------------------------------------
--------------------------------------------Vendor reference Number  Mandatory and not Duplicate Goods Return-------------------
 Cnt1:=0;
   if :object_type='21' and (:transaction_type ='A' OR :transaction_type ='U') then  							
      Select Count(*) into Cnt1 from "ORPD" T0 Where T0."DocEntry" = :list_of_cols_val_tab_del and 
      T0."CANCELED"='N' and (( IFNULL(T0."NumAtCard",'')<>'' and UPPER(T0."NumAtCard") in (Select distinct UPPER("NumAtCard") from ORPD where "DocEntry" <> :list_of_cols_val_tab_del and "CANCELED"='N' and "CardCode" = T0."CardCode"))
      OR (IFNULL(T0."NumAtCard",'')=''));
    if :Cnt1>0 then	
            error := -2060;							
            error_message := 'Vendor Ref No. is Empty or duplicated' ;							
    end if;
end if;
----------------------------------------------------------------------------------------------------------------
--------------------------------------------Vendor reference Number  Mandatory and not Duplicate Credit memo-------------------
 Cnt1:=0;
   if :object_type='19' and (:transaction_type ='A' OR :transaction_type ='U') then  							
      Select Count(*) into Cnt1 from "ORPC" T0 Where T0."DocEntry" = :list_of_cols_val_tab_del and 
      T0."CANCELED"='N' and (( IFNULL(T0."NumAtCard",'')<>'' and UPPER(T0."NumAtCard") in (Select distinct UPPER("NumAtCard") from ORPC where "DocEntry" <> :list_of_cols_val_tab_del and "CANCELED"='N' and "CardCode" = T0."CardCode"))
      OR (IFNULL(T0."NumAtCard",'')=''));
    if :Cnt1>0 then	
            error := -2060;							
            error_message := 'Vendor Ref No. is Empty or duplicated' ;							
    end if;
end if;
----------------------------------------------------------------------------------------------------------------

--------------------------------------------------------------------------------------------------------------------------------
---------------------------------------Item master data Item Group change Validation---------------------
 Cnt1 := 0;
IF :object_type = '4'
   AND ( :transaction_type = 'U')
THEN
    SELECT COUNT(*) 
      INTO Cnt1
      FROM OITM T0
      LEFT JOIN ( SELECT T1."ItemCode",
                  T1."ItmsGrpCod" FROM AITM T1
              INNER JOIN (
                    SELECT "ItemCode",
                           MAX("LogInstanc") AS "MaxLog" FROM AITM
                     GROUP BY "ItemCode") Latest
                ON T1."ItemCode" = Latest."ItemCode"
               AND T1."LogInstanc" = Latest."MaxLog") Prev
        ON T0."ItemCode" = Prev."ItemCode"
     WHERE T0."ItemCode" = :list_of_cols_val_tab_del
       AND IFNULL(Prev."ItmsGrpCod", T0."ItmsGrpCod") <> T0."ItmsGrpCod";
    IF :Cnt1 > 0 THEN
        error := -1706;
        error_message := 'You cannot Change ItemGroup';

    END IF;
 
END IF;
-------------------------------------------------------------------------------------------------------------------------------
--------------------------------------Sales Person and Supervisor Details are mandatory in contact person Details--------------
/*Cnt1 := 0;
IF :object_type = '2' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN
    SELECT COUNT(*) INTO Cnt1
    FROM "OCRD" T0
    LEFT JOIN OCPR T1 ON T0."CardCode" = T1."CardCode"
    WHERE T0."CardCode" = :list_of_cols_val_tab_del AND T0."CardType"='S'
    AND IFNULL(T1."Name",'')<>'';
    IF :Cnt1 < 2 THEN
        error := -2001;
        error_message := ' Supplier must have at least two contact persons(Supervisor and Sales Person)Contact person tab ';
    END IF;
END IF;*/
---------------------------------------------------------------------------------------------------------------------
-------------------------------GRPO Musted be Based on Purchase order Status----------------------------------
Cnt1 := 0;
IF :object_type = '20' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN
    SELECT COUNT(*) INTO Cnt1
    FROM "PDN1" T0
    INNER JOIN "OPOR" T1 ON T0."BaseEntry" = T1."DocEntry"
    WHERE T0."DocEntry" = :list_of_cols_val_tab_del
      AND T0."BaseType" = '22'
      AND IFNULL(T1."U_POSTS",'Lock') <> 'Released';
    IF Cnt1 > 0 THEN
        error := -2002;
        error_message := 'GRPO cannot created from Locked Purchase Orders.';
    END IF;
END IF;

----------------------------------------------------------------------------------------------------------------------------
---------------------------------Incoming Payment Restrict Credit Customers from cash Transcations ------------------------
Cnt1 := 0;
IF :object_type = '24' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN
    SELECT COUNT(*) INTO Cnt1
    FROM ORCT T0
    INNER JOIN OCRD T1 ON T0."CardCode" = T1."CardCode"
    WHERE T0."DocEntry" = :list_of_cols_val_tab_del
    AND T1."GroupCode" = '104'
    AND T0."CashSum" > 0;

    IF :Cnt1 > 0 THEN
        error := -24001;
        error_message := 'Cash transactions are not allowed for Credit Customers';
    END IF;

END IF;

---------------------Outgoing Payment Restrict Credit Customers from cash --------------------------------
/*Cnt1 := 0;
IF :object_type = '46' AND  (:transaction_type = 'A' OR :transaction_type = 'U') THEN
    SELECT COUNT(*) INTO Cnt1
    FROM OVPM T0
    INNER JOIN OCRD T1 ON T0."CardCode" = T1."CardCode"
    WHERE T0."DocEntry" = :list_of_cols_val_tab_del
    AND T1."GroupCode" = '104'
    AND T0."CashSum" > 0;

    IF :Cnt1 > 0 THEN
        error := -46001;
        error_message := 'Cash payment is not allowed for Credit Customers';
    END IF;

END IF;*/
---------------------------------------- Goods Issue Qty must be equal to STR Qty ----------------------------------------------------------------------------------
IF :object_type = '60' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN

    DECLARE STRDocEntry INT;
    DECLARE STRQty DECIMAL(19,6);
    DECLARE GIQty DECIMAL(19,6);

    -- Get STR DocEntry directly from Goods Issue
    SELECT "U_INVTRREN" INTO STRDocEntry 
    FROM OIGE 
    WHERE "DocEntry" = :list_of_cols_val_tab_del;

    IF :STRDocEntry IS NOT NULL THEN

        -- Get STR Total Quantity
        SELECT IFNULL(SUM("U_MEOPNQTY"),0) INTO STRQty 
        FROM "@ME_STR_C0" 
        WHERE "DocEntry" = :STRDocEntry;

        -- Get Goods Issue Total Quantity
        SELECT IFNULL(SUM("Quantity"),0) INTO GIQty 
        FROM IGE1 
        WHERE "DocEntry" = :list_of_cols_val_tab_del;

        -- Validation
        IF :STRQty <> :GIQty THEN
            error := -1;
            error_message := 'Goods Issue Qty must be equal to STR Qty';
        END IF;

    END IF;

END IF;

---------------------------------------Procurement restics  the Vendor does not have Active Purchase Blanket Agreement--------
-------------------------------- Purchase Request Validation --------------------------------

IF :object_type = '1470000113' AND :transaction_type = 'A' THEN

    DECLARE Cnt1 INTEGER;
    Cnt1 := 0;

    SELECT COUNT(*) INTO Cnt1
    FROM OOAT T0
    INNER JOIN OPRQ T1 ON T0."BpCode" = T1."U_VCode"
    WHERE T1."DocEntry" = :list_of_cols_val_tab_del
    AND T0."Status" = 'A'
    AND T1."DocDate" BETWEEN T0."StartDate" AND T0."EndDate";

    IF :Cnt1 = 0 THEN
        error := 14701;
        error_message := 'Procurement is restricted for suppliers without an active Blanket Agreement';
    END IF;

END IF;
-------------------------------- Purchase Order Validation --------------------------------

IF :object_type = '22' AND :transaction_type = 'A' THEN

    DECLARE Cnt1 INTEGER;
    Cnt1 := 0;

    SELECT COUNT(*) INTO Cnt1
    FROM OOAT T0
    INNER JOIN OPOR T1 ON T0."BpCode" = T1."CardCode"
    WHERE T1."DocEntry" = :list_of_cols_val_tab_del
    AND T0."Status" = 'A'
    AND T1."DocDate" BETWEEN T0."StartDate" AND T0."EndDate";

    IF :Cnt1 = 0 THEN
        error := 2203;
        error_message := 'Procurement is restricted for suppliers without an active Blanket Agreement';
    END IF;

END IF;
-------------------------------- GRPO Validation --------------------------------

IF :object_type = '20' AND :transaction_type = 'A' THEN

    DECLARE Cnt1 INTEGER;
    Cnt1 := 0;

    SELECT COUNT(*) INTO Cnt1
    FROM OOAT T0
    INNER JOIN OPDN T1 ON T0."BpCode" = T1."CardCode"
    WHERE T1."DocEntry" = :list_of_cols_val_tab_del
    AND T0."Status" = 'A'
    AND T1."DocDate" BETWEEN T0."StartDate" AND T0."EndDate";

    IF :Cnt1 = 0 THEN
        error := -20009;
        error_message := 'Procurement is restricted for suppliers without an active Blanket Agreement';
    END IF;

END IF;
---------------------------- PROMOTION DUPLICATION for selected period------------------------------------------------
IF :object_type = 'TNXPCRUDO' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN

    DECLARE lv_cnt INTEGER;

    SELECT COUNT(*) INTO lv_cnt

    FROM "ALJAZEERAUAT"."@TNX_ITEMDISC" I0

    INNER JOIN "ALJAZEERAUAT"."@TNX_ITEMDISC_C0" I1

        ON I0."DocEntry" = I1."DocEntry"

    WHERE I0."U_BPLID" = (

            SELECT P0."U_TNXBRC"

            FROM "ALJAZEERAUAT"."@TNX_PCRT" P0

            WHERE P0."DocEntry" = :list_of_cols_val_tab_del

          )

      AND TRIM(I0."U_PRTYP") = (

            SELECT TRIM(P0."U_TNXPUTP")

            FROM "ALJAZEERAUAT"."@TNX_PCRT" P0

            WHERE P0."DocEntry" = :list_of_cols_val_tab_del

          )

      AND (

            I1."U_ITMCODE" IN (

                SELECT P1."U_ItemCode"

                FROM "ALJAZEERAUAT"."@TNX_PCRT_C1" P1

                WHERE P1."DocEntry" = :list_of_cols_val_tab_del

            )

            OR

            I1."U_BCODE" IN (

                SELECT P1."U_BrCode"

                FROM "ALJAZEERAUAT"."@TNX_PCRT_C1" P1

                WHERE P1."DocEntry" = :list_of_cols_val_tab_del

            )

          )

      AND

          TO_TIMESTAMP(

              TO_VARCHAR(

                  (SELECT P0."U_FDATEF"

                   FROM "ALJAZEERAUAT"."@TNX_PCRT" P0

                   WHERE P0."DocEntry" = :list_of_cols_val_tab_del),

                  'YYYY-MM-DD'

              ) || ' ' ||

              SUBSTRING(LPAD(TO_VARCHAR(IFNULL(

                  (SELECT P0."U_FTIMEF"

                   FROM "ALJAZEERAUAT"."@TNX_PCRT" P0

                   WHERE P0."DocEntry" = :list_of_cols_val_tab_del), 0)), 4, '0'), 1, 2) || ':' ||

              SUBSTRING(LPAD(TO_VARCHAR(IFNULL(

                  (SELECT P0."U_FTIMEF"

                   FROM "ALJAZEERAUAT"."@TNX_PCRT" P0

                   WHERE P0."DocEntry" = :list_of_cols_val_tab_del), 0)), 4, '0'), 3, 2),

              'YYYY-MM-DD HH24:MI'

          )
<=

          TO_TIMESTAMP(

              TO_VARCHAR(I1."U_TODATE", 'YYYY-MM-DD') || ' ' ||

              SUBSTRING(LPAD(TO_VARCHAR(IFNULL(I0."U_TOTIME", 0)), 4, '0'), 1, 2) || ':' ||

              SUBSTRING(LPAD(TO_VARCHAR(IFNULL(I0."U_TOTIME", 0)), 4, '0'), 3, 2),

              'YYYY-MM-DD HH24:MI'

          )

      AND

          TO_TIMESTAMP(

              TO_VARCHAR(

                  (SELECT P0."U_TDATEF"

                   FROM "ALJAZEERAUAT"."@TNX_PCRT" P0

                   WHERE P0."DocEntry" = :list_of_cols_val_tab_del),

                  'YYYY-MM-DD'

              ) || ' ' ||

              SUBSTRING(LPAD(TO_VARCHAR(IFNULL(

                  (SELECT P0."U_TTIMEF"

                   FROM "ALJAZEERAUAT"."@TNX_PCRT" P0

                   WHERE P0."DocEntry" = :list_of_cols_val_tab_del), 0)), 4, '0'), 1, 2) || ':' ||

              SUBSTRING(LPAD(TO_VARCHAR(IFNULL(

                  (SELECT P0."U_TTIMEF"

                   FROM "ALJAZEERAUAT"."@TNX_PCRT" P0

                   WHERE P0."DocEntry" = :list_of_cols_val_tab_del), 0)), 4, '0'), 3, 2),

              'YYYY-MM-DD HH24:MI'

          )
>=

          TO_TIMESTAMP(

              TO_VARCHAR(I1."U_FRMDATE", 'YYYY-MM-DD') || ' ' ||

              SUBSTRING(LPAD(TO_VARCHAR(IFNULL(I0."U_FRMTIME", 0)), 4, '0'), 1, 2) || ':' ||

              SUBSTRING(LPAD(TO_VARCHAR(IFNULL(I0."U_FRMTIME", 0)), 4, '0'), 3, 2),

              'YYYY-MM-DD HH24:MI'

          );

    IF :lv_cnt > 0 THEN

        error := 1;

        error_message := 'Active promotion already exists for the same branch, item/barcode, promotion type, and selected period. Price Change Request cannot be created.';

    END IF;

END IF;
 
-----------------------BARCODE DUPLICATION IN ITEM MASTER UPLOAD------------------------------------
IF :object_type = 'TNX_ITEMUP_UDO' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN
 
    DECLARE lv_barcode NVARCHAR(100);
    DECLARE lv_dupdoc INTEGER;
 
    lv_barcode := '';
    lv_dupdoc := 0;
 
    -- 1. Duplicate within document
    SELECT COUNT(*) INTO lv_dupdoc
    FROM (
        SELECT TRIM(T0."U_BarCode")
        FROM "ALJAZEERAUAT"."@TNX_ITEMUP_C2" T0
        WHERE T0."DocEntry" = :list_of_cols_val_tab_del
          AND IFNULL(TRIM(T0."U_BarCode"), '') <> ''
        GROUP BY TRIM(T0."U_BarCode")
        HAVING COUNT(*) > 1
    ) X;
 
    IF :lv_dupdoc > 0 THEN
        error := 1;
        error_message := 'Duplicate barcode found within the upload document.';
    END IF;
 
    -- 2. Duplicate in OITM / OBCD
    IF :error = 0 THEN
 
        SELECT IFNULL(MAX(T0."U_BarCode"), '') 
        INTO lv_barcode
        FROM "ALJAZEERAUAT"."@TNX_ITEMUP_C2" T0
        WHERE T0."DocEntry" = :list_of_cols_val_tab_del
          AND IFNULL(TRIM(T0."U_BarCode"), '') <> ''
          AND (
                EXISTS (
                    SELECT 1 FROM "OITM" I
                    WHERE TRIM(IFNULL(I."CodeBars", '')) = TRIM(T0."U_BarCode")
                )
                OR
                EXISTS (
                    SELECT 1 FROM "OBCD" B
                    WHERE TRIM(IFNULL(B."BcdCode", '')) = TRIM(T0."U_BarCode")
                )
              );
 
        IF :lv_barcode <> '' THEN
            error := 1;
            error_message := 'Barcode already exists in Item Master / Barcode Master: ' || :lv_barcode;
        END IF;
 
    END IF;
 
END IF;
 ---------Blanket Agreement field Mandotry in PO lines----------------------
Cnt1:=0;							
   IF :object_type='22' and (:transaction_type ='A' OR :transaction_type ='U') then 					
      Select COUNT(*) into Cnt1 
      from POR1 T0
      Where T0."DocEntry" = :list_of_cols_val_tab_del
      and T0."AgrNo" IS NULL;							
 							
      if :Cnt1>0 then							
            error := -20010;							
            error_message := 'Blanket Agreement in  lines cannot be empty';							
      end if;  
   END IF;
 -------------------------------
---------------Item already active in another promotion document-------
Cnt1 := 0;
Line :=0;

IF :object_type = 'ME_ITEMUDO'
   AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN

    SELECT count(*) into Cnt1 from  "@TNX_ITEMDISC" T1
         INNER JOIN "@TNX_ITEMDISC_C0" T0 
                 ON T0."DocEntry" = T1."DocEntry" Where T1."DocEntry" = :list_of_cols_val_tab_del and IFNULL(T0."U_BCODE",'')<>'' and
                 T0."U_BCODE" in (Select distinct T2."U_BCODE" from "@TNX_ITEMDISC_C0" T2 INNER JOIN "@TNX_ITEMDISC" T3 on T2."DocEntry" = T3."DocEntry" where T3."DocEntry" <> :list_of_cols_val_tab_del
                 AND UPPER(T3."U_STATUS") = 'ACTIVE' and ((TO_CHAR(T0."U_FRMDATE",'YYYYMMDD') ||LPAD(IFNULL(T1."U_FRMTIME",0),4,'0') >= TO_CHAR(T2."U_FRMDATE",'YYYYMMDD') ||LPAD(IFNULL(T3."U_FRMTIME",0),4,'0')
AND TO_CHAR(T0."U_FRMDATE",'YYYYMMDD') ||LPAD(IFNULL(T1."U_FRMTIME",0),4,'0') <= TO_CHAR(T2."U_TODATE",'YYYYMMDD')||LPAD(IFNULL(T3."U_TOTIME",0),4,'0')) OR (TO_CHAR(T0."U_TODATE",'YYYYMMDD') ||LPAD(IFNULL(T1."U_TOTIME",0),4,'0') >= TO_CHAR(T2."U_FRMDATE",'YYYYMMDD') ||LPAD(IFNULL(T3."U_FRMTIME",0),4,'0')
AND TO_CHAR(T0."U_TODATE",'YYYYMMDD') ||LPAD(IFNULL(T1."U_TOTIME",0),4,'0') <= TO_CHAR(T2."U_TODATE",'YYYYMMDD')||LPAD(IFNULL(T3."U_TOTIME",0),4,'0')))
AND T1."U_BPLID" = T3."U_BPLID" and IFNULL(T2."U_BCODE",'')<>'');
       IF :Cnt1 > 0 THEN
       SELECT TOP 1 T0."LineId"   into Line from  "@TNX_ITEMDISC" T1
         INNER JOIN "@TNX_ITEMDISC_C0" T0 
                 ON T0."DocEntry" = T1."DocEntry" Where T1."DocEntry" = :list_of_cols_val_tab_del and IFNULL(T0."U_BCODE",'')<>'' and
                 T0."U_BCODE" in (Select distinct T2."U_BCODE" from "@TNX_ITEMDISC_C0" T2 INNER JOIN "@TNX_ITEMDISC" T3 on T2."DocEntry" = T3."DocEntry" where T3."DocEntry" <> :list_of_cols_val_tab_del
                 AND UPPER(T3."U_STATUS") = 'ACTIVE' and ((TO_CHAR(T0."U_FRMDATE",'YYYYMMDD') ||LPAD(IFNULL(T1."U_FRMTIME",0),4,'0') >= TO_CHAR(T2."U_FRMDATE",'YYYYMMDD') ||LPAD(IFNULL(T3."U_FRMTIME",0),4,'0')
AND TO_CHAR(T0."U_FRMDATE",'YYYYMMDD') ||LPAD(IFNULL(T1."U_FRMTIME",0),4,'0') <= TO_CHAR(T2."U_TODATE",'YYYYMMDD')||LPAD(IFNULL(T3."U_TOTIME",0),4,'0')) OR (TO_CHAR(T0."U_TODATE",'YYYYMMDD') ||LPAD(IFNULL(T1."U_TOTIME",0),4,'0') >= TO_CHAR(T2."U_FRMDATE",'YYYYMMDD') ||LPAD(IFNULL(T3."U_FRMTIME",0),4,'0')
AND TO_CHAR(T0."U_TODATE",'YYYYMMDD') ||LPAD(IFNULL(T1."U_TOTIME",0),4,'0') <= TO_CHAR(T2."U_TODATE",'YYYYMMDD')||LPAD(IFNULL(T3."U_TOTIME",0),4,'0')))
AND T1."U_BPLID" = T3."U_BPLID" and IFNULL(T2."U_BCODE",'')<>'');
        error := -1203;
        error_message := 'Duplicate validity range found in previous document for Line ' || :Line;
    END IF;

END IF;
---------------------------------------------------------
---------------Item already active in another promotion document-------
Cnt1 := 0;
Line :=0;

IF :object_type = 'TNXPCRUDO'
   AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN

    SELECT count(*) into Cnt1 from  "@TNX_ITEMDISC" T1
         INNER JOIN "@TNX_ITEMDISC_C0" T0 ON T0."DocEntry" = T1."DocEntry" 
         Inner join "@TNX_PCRT_C3" A ON A."U_PDN" = T1."DocNum"
         Inner join "@TNX_PCRT" B ON B."DocEntry" = A."DocEntry"
         Where B."DocEntry" = :list_of_cols_val_tab_del and IFNULL(A."U_BrCode",'')<>'' and
                 A."U_BrCode" in (Select distinct T2."U_BCODE" from "@TNX_ITEMDISC_C0" T2 
                 INNER JOIN "@TNX_ITEMDISC" T3 on T2."DocEntry" = T3."DocEntry" where T3."DocEntry" <> :list_of_cols_val_tab_del
                 AND UPPER(T3."U_STATUS") = 'ACTIVE' and ((TO_CHAR(B."U_FDATET",'YYYYMMDD') ||LPAD(IFNULL(T1."U_FRMTIME",0),4,'0') >= TO_CHAR(T2."U_FRMDATE",'YYYYMMDD') ||LPAD(IFNULL(T3."U_FRMTIME",0),4,'0')
AND TO_CHAR(B."U_FDATET",'YYYYMMDD') ||LPAD(IFNULL(T1."U_FRMTIME",0),4,'0') <= TO_CHAR(T2."U_TODATE",'YYYYMMDD')||LPAD(IFNULL(T3."U_TOTIME",0),4,'0')) OR (TO_CHAR(B."U_TDATET",'YYYYMMDD') ||LPAD(IFNULL(T1."U_TOTIME",0),4,'0') >= TO_CHAR(T2."U_FRMDATE",'YYYYMMDD') ||LPAD(IFNULL(T3."U_FRMTIME",0),4,'0')
AND TO_CHAR(B."U_TDATET",'YYYYMMDD') ||LPAD(IFNULL(T1."U_TOTIME",0),4,'0') <= TO_CHAR(T2."U_TODATE",'YYYYMMDD')||LPAD(IFNULL(T3."U_TOTIME",0),4,'0')))
AND B."U_TNXBRC" = T3."U_BPLID" and IFNULL(T2."U_BCODE",'')<>'');
       IF :Cnt1 > 0 THEN
       SELECT TOP 1 T0."LineId"   into Line from  "@TNX_ITEMDISC" T1
            INNER JOIN "@TNX_ITEMDISC_C0" T0 ON T0."DocEntry" = T1."DocEntry" 
         Inner join "@TNX_PCRT_C3" A ON A."U_PDN" = T1."DocNum"
         Inner join "@TNX_PCRT" B ON B."DocEntry" = A."DocEntry"
         Where B."DocEntry" = :list_of_cols_val_tab_del and IFNULL(A."U_BrCode",'')<>'' and
                 A."U_BrCode" in (Select distinct T2."U_BCODE" from "@TNX_ITEMDISC_C0" T2 
                 INNER JOIN "@TNX_ITEMDISC" T3 on T2."DocEntry" = T3."DocEntry" where T3."DocEntry" <> :list_of_cols_val_tab_del
                 AND UPPER(T3."U_STATUS") = 'ACTIVE' and ((TO_CHAR(B."U_FDATET",'YYYYMMDD') ||LPAD(IFNULL(T1."U_FRMTIME",0),4,'0') >= TO_CHAR(T2."U_FRMDATE",'YYYYMMDD') ||LPAD(IFNULL(T3."U_FRMTIME",0),4,'0')
AND TO_CHAR(B."U_FDATET",'YYYYMMDD') ||LPAD(IFNULL(T1."U_FRMTIME",0),4,'0') <= TO_CHAR(T2."U_TODATE",'YYYYMMDD')||LPAD(IFNULL(T3."U_TOTIME",0),4,'0')) OR (TO_CHAR(B."U_TDATET",'YYYYMMDD') ||LPAD(IFNULL(T1."U_TOTIME",0),4,'0') >= TO_CHAR(T2."U_FRMDATE",'YYYYMMDD') ||LPAD(IFNULL(T3."U_FRMTIME",0),4,'0')
AND TO_CHAR(B."U_TDATET",'YYYYMMDD') ||LPAD(IFNULL(T1."U_TOTIME",0),4,'0') <= TO_CHAR(T2."U_TODATE",'YYYYMMDD')||LPAD(IFNULL(T3."U_TOTIME",0),4,'0')))
AND B."U_TNXBRC" = T3."U_BPLID" and IFNULL(T2."U_BCODE",'')<>'');
        error := -1204;
        error_message := 'Duplicate validity range found in previous document for Line ' || :Line;
    END IF;

END IF;

 
-----------------------------------------------------------------------------------------
/*---------Blanket Agreement field Mandotry in PR lines----------------------
Cnt1:=0;							
   if :object_type='1470000113' and (:transaction_type ='A' OR :transaction_type ='U') then 					
      Select Count(*) into Cnt1 
      from PRQ1 T0
      Where T0."DocEntry" = :list_of_cols_val_tab_del
      and IFNULL(T0."AgrNo",'')='';							
 							
      if :Cnt1>0 then							
            error := -220010;							
            error_message := 'Blanket Agreement in  lines cannot be empty';							
      end if; 
    END IF*/
------------------------------------------------------------------------------------------------------------------------
IF :object_type = 'SBO_LISTUDO' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN

    DECLARE cnt INT;

    SELECT COUNT(*) INTO cnt
    FROM "@SBO_LIST"
    WHERE "DocEntry" = :list_of_cols_val_tab_del
      AND IFNULL("U_CardCode",'') = '';

    IF cnt > 0 THEN
        error := -309001;
        error_message := 'Card Code is mandatory';
    END IF;

END IF;
IF :object_type = 'SBO_LISTUDO' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN

    DECLARE cnt INT;

    SELECT COUNT(*) INTO cnt
    FROM "@SBO_LIST_C0"
    WHERE "DocEntry" = :list_of_cols_val_tab_del
      AND IFNULL("U_Date",'') = '';

    IF cnt > 0 THEN
        error := -309002;
        error_message := 'Date cannot be empty';
    END IF;

END IF;
IF :object_type = 'SBO_LISTUDO' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN

    DECLARE cnt INT;

    SELECT COUNT(*) INTO cnt
    FROM "@SBO_LIST_C0"
    WHERE "DocEntry" = :list_of_cols_val_tab_del
   ----   AND IFNULL("U_SKU",'0') 
      and IFNULL(CAST(NULLIF("U_SKU",'') AS INTEGER),0)
      <= 0;

    IF cnt > 0 THEN
        error := -309003;
        error_message := 'SKU must be greater than 0';
    END IF;

END IF;

/*IF :object_type = 'SBO_LISTUDO' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN

    DECLARE cnt INT;

    SELECT COUNT(*) INTO cnt
    FROM "@SBO_LIST_C0"
    WHERE "DocEntry" = :list_of_cols_val_tab_del
      AND IFNULL("U_BRANCH",'') = '';

    IF cnt > 0 THEN
        error := -309004;
        error_message := 'Branch is mandatory';
    END IF;

END IF;*/
IF :object_type = 'SBO_OUTLETUDO'
   AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN
 
    DECLARE OutletCnt INT;
    DECLARE CheckedCnt INT;
 
    /* Get Outlet count from UDO Header */
    SELECT IFNULL("U_Outlet", '0')
    INTO OutletCnt
    FROM "@SBO_OUTLET"
    WHERE "DocEntry" = :list_of_cols_val_tab_del;
 
    /* Count checked Amount rows from child table */
    SELECT COUNT(*)
    INTO CheckedCnt
    FROM "@SBO_OUTLET_C0"
    WHERE "DocEntry" = :list_of_cols_val_tab_del
      AND IFNULL("U_Amount", 'N') = 'Y';
 
    /* Validation: Amount count must be exactly equal to Outlet count */
    IF :CheckedCnt <> :OutletCnt THEN
        error := -3090059;
        error_message :=
            'Selected Outlet count (' ||
            :CheckedCnt || ') must be exactly equal to Outlet count (' ||
            :OutletCnt || ').';
    END IF;
 
END IF;

IF :object_type = 'SBO_LISTUDO' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN

    DECLARE cnt INT;

    SELECT COUNT(*) INTO cnt
    FROM "@SBO_LIST_C0"
    WHERE "DocEntry" = :list_of_cols_val_tab_del
      --AND IFNULL("U_NOU",0) 
  and IFNULL(CAST(NULLIF("U_NOU",'') AS INTEGER),0)  
      <= 0;

    IF cnt > 0 THEN
        error := -309005;
        error_message := 'Number of Outlets must be greater than 0';
    END IF;

END IF;
IF :object_type = 'SBO_LISTUDO' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN

    DECLARE cnt INT;
    DECLARE BranchCnt INT;

    SELECT COUNT(*) INTO BranchCnt
    FROM OBPL
    WHERE IFNULL("Disabled",'N') = 'N';

    SELECT COUNT(*) INTO cnt
    FROM "@SBO_LIST_C0"
    WHERE "DocEntry" = :list_of_cols_val_tab_del
      and IFNULL(CAST(NULLIF("U_NOU",'') AS INTEGER),0)
  
     
      > BranchCnt;

    IF cnt > 0 THEN
        error := -309006;
        error_message := 'No of Outlets cannot exceed total branches (' || BranchCnt || ')';
    END IF;

END IF;
IF :object_type = 'SBO_LISTUDO' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN

    DECLARE cnt INT;

    SELECT COUNT(*) INTO cnt
    FROM (
        SELECT "U_SKU", "U_Date", COUNT(*) 
        FROM "@SBO_LIST_C0"
        WHERE "DocEntry" = :list_of_cols_val_tab_del
        GROUP BY "U_SKU", "U_Date"
        HAVING COUNT(*) > 1
    ) T;

    IF cnt > 0 THEN
        error := -309007;
        error_message := 'Duplicate SKU + Date not allowed';
    END IF;

END IF;
/*IF :object_type = 'SBO_LISTUDO' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN

    DECLARE cnt INT;

    SELECT COUNT(*) INTO cnt
    FROM "@SBO_LIST_C2"
    WHERE "DocEntry" = :list_of_cols_val_tab_del
      AND IFNULL("U_DESC",'') = '';

    IF cnt > 0 THEN
        error := -309008;
        error_message := 'Description is mandatory';
    END IF;

END IF;*/
IF :object_type = 'SBO_LISTUDO' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN

    DECLARE cnt INT;

    SELECT COUNT(*) INTO cnt
    FROM "@SBO_LIST_C0"
    WHERE "DocEntry" = :list_of_cols_val_tab_del
      AND IFNULL("U_Amount",0) <= 0;

    IF cnt > 0 THEN
        error := -309009;
        error_message := 'Amount must be greater than 0';
    END IF;

END IF;
---------------------------------------------------------------------
IF :object_type = 'TNX_WarehouseUDO' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN

    DECLARE HQTY DECIMAL(19,6);
    DECLARE LQTY DECIMAL(19,6);
    DECLARE DUP_COUNT INT;

    SELECT T0."U_Quantity" 
    INTO HQTY
    FROM "@TNX_WAREHOUSE" T0
    WHERE T0."DocEntry" = :list_of_cols_val_tab_del;
    
    SELECT SUM(T1."U_Quantity") 
    INTO LQTY
    FROM "@TNX_WAREHOUSE_C0" T1
    WHERE T1."DocEntry" = :list_of_cols_val_tab_del; 
    
    SELECT COUNT(*) INTO DUP_COUNT
    FROM (
        SELECT T1."U_ToWhse"
        FROM "@TNX_WAREHOUSE_C0" T1
        WHERE T1."DocEntry" = :list_of_cols_val_tab_del
        GROUP BY T1."U_ToWhse"
        HAVING COUNT(*) > 1
    ) A; 
 
    IF HQTY < LQTY THEN
        error := -118;
        error_message := 'Row quantity should not be greater than header quantity';
    END IF;
    
    IF DUP_COUNT > 0 THEN
        error := -119;
        error_message := 'Duplicate warehouse is not allowed in row level';
    END IF;

END IF;

-----------------------BARCODE DUPLICATION IN ITEM MASTER UPLOAD------------------------------------
IF :object_type = 'TNX_ITEMUP_UDO' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN
 
    DECLARE lv_barcode NVARCHAR(100);
    DECLARE lv_dupdoc INTEGER;
 
    lv_barcode := '';
    lv_dupdoc := 0;
 
    -- 1. Duplicate within document
    SELECT COUNT(*) INTO lv_dupdoc
    FROM (
        SELECT TRIM(T0."U_BarCode")
        FROM "ALJAZEERAUAT"."@TNX_ITEMUP_C2" T0
        WHERE T0."DocEntry" = :list_of_cols_val_tab_del
          AND IFNULL(TRIM(T0."U_BarCode"), '') <> ''
        GROUP BY TRIM(T0."U_BarCode")
        HAVING COUNT(*) > 1
    ) X;
 
    IF :lv_dupdoc > 0 THEN
        error := 1;
        error_message := 'Duplicate barcode found within the upload document.';
    END IF;
 
    -- 2. Duplicate in OITM / OBCD
    IF :error = 0 THEN
 
        SELECT IFNULL(MAX(T0."U_BarCode"), '') 
        INTO lv_barcode
        FROM "ALJAZEERAUAT"."@TNX_ITEMUP_C2" T0
        WHERE T0."DocEntry" = :list_of_cols_val_tab_del
          AND IFNULL(TRIM(T0."U_BarCode"), '') <> ''
          AND (
                EXISTS (
                    SELECT 1 FROM "OITM" I
                    WHERE TRIM(IFNULL(I."CodeBars", '')) = TRIM(T0."U_BarCode")
                )
                OR
                EXISTS (
                    SELECT 1 FROM "OBCD" B
                    WHERE TRIM(IFNULL(B."BcdCode", '')) = TRIM(T0."U_BarCode")
                )
              );
 
        IF :lv_barcode <> '' THEN
            error := 1;
            error_message := 'Barcode already exists in Item Master / Barcode Master: ' || :lv_barcode;
        END IF;
 
    END IF;
 
END IF;
-------------------------------------------------------------------------------------------
IF :object_type = '22' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN

    DECLARE POCOUNT INT;

    SELECT COUNT(*) 
    INTO POCOUNT
    FROM OPOR T0 
    INNER JOIN POR1 T1 ON T0."DocEntry" = T1."DocEntry"
    LEFT JOIN OITM T2 ON T1."ItemCode" = T2."ItemCode"
    WHERE T0."DocEntry" = :list_of_cols_val_tab_del 
      AND T2."U_SugarTax" = 'Y'
      AND T1."U_normalqty" <> 0 
      AND T1."U_FOCQTY" <> 0;

    IF :POCOUNT > 0 THEN
        error := -119;
        error_message := 'Normal qty and FOC qty should not be allowed together in PO for Sugar tax Items';
    END IF;

END IF;
----------------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------------------
-- Validation: Linked Goods Receipt cannot exceed original Goods Issue Quantity
------------------------------------------------------------------------------------------------------------------

IF :object_type = '59' AND (:transaction_type = 'A' OR :transaction_type = 'U') THEN
    
    DECLARE v_ItemCode NVARCHAR(50);
    DECLARE v_BaseEntry INT;
    DECLARE v_BaseQty DECIMAL(19,6);
    DECLARE v_TotalRecQty DECIMAL(19,6);
    DECLARE v_ErrorMsg NVARCHAR(200);

    -- Cursor to check each line of the Goods Receipt being added
    DECLARE CURSOR c_lines FOR 
        SELECT "ItemCode", "BaseEntry", "Quantity" 
        FROM IGN1 
        WHERE "DocEntry" = :list_of_cols_val_tab_del;

    FOR cur_row AS c_lines DO
        v_ItemCode := cur_row."ItemCode";
        v_BaseEntry := cur_row."BaseEntry"; -- This is the GI DocEntry (TargetEntry link)

        -- 1. Get the original Issued Quantity from the Goods Issue (OIGE/IGE1)
        SELECT SUM("Quantity") INTO v_BaseQty 
        FROM IGE1 
        WHERE "DocEntry" = v_BaseEntry AND "ItemCode" = v_ItemCode;

        -- 2. Sum up ALL Receipts already linked to this Goods Issue (including this one)
        -- We check TrgetEntry in IGE1 to find all OIGN documents linked to that GI
        SELECT SUM(T1."Quantity") INTO v_TotalRecQty
        FROM OIGN T0 
        INNER JOIN IGN1 T1 ON T0."DocEntry" = T1."DocEntry"
        WHERE T1."BaseEntry" = v_BaseEntry 
        AND T1."ItemCode" = v_ItemCode
        AND T1."BaseType" = 60; -- BaseType 60 refers to Goods Issue

        -- 3. Validation Check
        IF :v_TotalRecQty > :v_BaseQty THEN
            error := 101;
            v_ErrorMsg := 'Quantity for Item ' || v_ItemCode || ' exceeds Issued Quantity! (Issued: ' || CAST(v_BaseQty AS NVARCHAR) || ')';
            error_message := :v_ErrorMsg;
        END IF;

    END FOR;
END IF;
------------------------------------------------------------------------------------------------------------------------------
IF :object_type = 'TNXPCRUDO' AND (:transaction_type = 'A' OR :transaction_type = 'U') 
THEN

    IF NOT EXISTS
    (
        SELECT 1
        FROM "@TNX_PCRT_C0"
        WHERE "DocEntry" = :list_of_cols_val_tab_del
          AND IFNULL(TRIM("U_TNXICD"), '') <> ''
    )
    THEN
        error := 1;
        error_message := 'At least one Item Code must be entered.';
    END IF;

END IF;

----------------------------------------------------------------------------------------------
IF :object_type = 'TNXPCRUDO' AND (:transaction_type = 'A' OR :transaction_type = 'U')
THEN

    -- Flyer Price Changes validation
    IF EXISTS
    (
        SELECT 1
        FROM "@TNX_PCRT" H0
        INNER JOIN "@TNX_PCRT_C1" L0
            ON H0."DocEntry" = L0."DocEntry"
        INNER JOIN "@TNX_PCRT" H1
            ON H0."U_FDATEF" = H1."U_FDATEF"
           AND H0."U_TDATEF" = H1."U_TDATEF"
           AND H0."U_TNXBRC" = H1."U_TNXBRC"
           AND H0."DocEntry" <> H1."DocEntry"
        INNER JOIN "@TNX_PCRT_C1" L1
            ON H1."DocEntry" = L1."DocEntry"
        WHERE H0."DocEntry" = :list_of_cols_val_tab_del
          AND H0."U_TNXPUTP" = 'Flyer Price Changes'
          AND H1."U_TNXPUTP" = 'Flyer Price Changes'
          AND L0."U_ItemCode" = L1."U_ItemCode"
    )
    THEN
        error := -1;
        error_message := 'Flyer Price Changes: This item already exists in another document with the same From Date and To Date.';
    END IF;
    
    
        -- MOP validation
    IF EXISTS
    (
        SELECT 1
        FROM "@TNX_PCRT" H0
        INNER JOIN "@TNX_PCRT_C2" L0
            ON H0."DocEntry" = L0."DocEntry"
        INNER JOIN "@TNX_PCRT" H1
            ON H0."U_FDATEM1" = H1."U_FDATEM1"
           AND H0."U_TDATEMP" = H1."U_TDATEMP"
           AND H0."U_TNXBRC" = H1."U_TNXBRC"
           AND H0."DocEntry" <> H1."DocEntry"
        INNER JOIN "@TNX_PCRT_C2" L1
            ON H1."DocEntry" = L1."DocEntry"
        WHERE H0."DocEntry" = :list_of_cols_val_tab_del
          AND H0."U_TNXPUTP" = 'MOP'
          AND H1."U_TNXPUTP" = 'MOP'
          AND L0."U_ItemCode" = L1."U_ItemCode"
    )
    THEN
        error := -1;
        error_message := 'MOP: This item already exists in another document with the same From Date and To Date.';
    END IF;


    -- Temporary Price Reduction validation
    IF EXISTS
    (
        SELECT 1
        FROM "@TNX_PCRT" H0
        INNER JOIN "@TNX_PCRT_C3" L0
            ON H0."DocEntry" = L0."DocEntry"
        INNER JOIN "@TNX_PCRT" H1
            ON H0."U_FDATET" = H1."U_FDATET"
           AND H0."U_TDATET" = H1."U_TDATET"
           AND H0."U_TNXBRC" = H1."U_TNXBRC"
           AND H0."DocEntry" <> H1."DocEntry"
        INNER JOIN "@TNX_PCRT_C3" L1
            ON H1."DocEntry" = L1."DocEntry"
        WHERE H0."DocEntry" = :list_of_cols_val_tab_del
          AND H0."U_TNXPUTP" = 'Temporary Price Reduction'
          AND H1."U_TNXPUTP" = 'Temporary Price Reduction'
          AND L0."U_ItemCode" = L1."U_ItemCode"
    )
    THEN
        error := -1;
        error_message := 'Temporary Price Reduction: This item already exists in another document with the same From Date and To Date.';
    END IF;

END IF;
------------------------------------------------------------------------------------------------------------------------------
-- Select the return values
select :error, :error_message FROM dummy;

end;