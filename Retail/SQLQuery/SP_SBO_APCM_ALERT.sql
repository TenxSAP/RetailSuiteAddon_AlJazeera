CREATE FUNCTION  "GET_RSP"(IN ITEMCODE VARCHAR(100),IN UOMCODE VARCHAR(100),IN PRICELIST VARCHAR(100))
RETURNS basic DOUBLE LANGUAGE SQLSCRIPT  
AS
BEGIN 
select (Case when  ifnull((Select "Price" from ITM1 where "ItemCode"=:ITEMCODE and "PriceList"=:PRICELIST and "UomEntry"=:UOMCODE),0)=0  then 
ifnull((Select "Price" from ITM9 where "ItemCode"=:ITEMCODE and "UomEntry"=:UOMCODE and "PriceList"=:PRICELIST),0)
else ifnull((Select "Price" from ITM1 where "ItemCode"=:ITEMCODE and "PriceList"=:PRICELIST and "UomEntry"=:UOMCODE),0) 
end) INTO basic from  dummy;

END;