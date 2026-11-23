SELECT DISTINCT 'TRUE' from OCRD T0 where T0."CreditLine" >1 and T0."CreditLine" < ( $[$29.0.Number] + T0."OrdersBal" + T0."DNotesBal" + T0."Balance") and T0."CardCode" = $[$4.0.0]
