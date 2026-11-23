Select "DocNum" , "DocDate","Comments" From OIGE Where "DocEntry" Not In (Select "U_MEGIEN" From "@ME_MI" Where IFNULL("U_MEGIEN",'') <> '')And "U_TYPE" = 'STR'
