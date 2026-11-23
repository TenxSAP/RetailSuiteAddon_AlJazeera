CREATE PROCEDURE "SBO_LOADSTOCK"(IN DocNum varchar(100))                          
AS 
ItemGrpName varchar(150);
Type1 varchar(5);
Itmgrpwise varchar(1);
temp_var_0 integer;
temp_var_1 integer;
temp_var_2 integer;
temp_var_3 integer;
temp_var_4 integer;
BEGIN

Select *,
(CASE WHEN "U_Update"='' THEN 'N' ELSE "U_Update" END)  AS "Update" 

 from "@SBO_STOCKWZD_C0" T1 Inner join  "@SBO_STOCKWZD" T2 on T2."DocEntry"=  T1."DocEntry" 
Where "U_Check"='Y' And (CASE WHEN "U_Update"='' THEN 'N' ELSE "U_Update" END) <>'Y'
 AND "DocNum"=:DocNum
;

END;
 