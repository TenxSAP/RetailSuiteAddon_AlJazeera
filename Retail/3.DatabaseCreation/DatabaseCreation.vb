
Imports System.Windows
Imports SAPbobsCOM

Public Class DatabaseCreation
    '  test push

#Region "Declaration"
    Private objUtilities As Utilities
    Dim DBCode As String = "NewEra_Retail"
    Dim DBName As String = "New Era Retail"
    ''Dim Version As String = "v1.0.0.32401"


    Dim Version As String = "v1.0.0.32598"


#End Region

#Region "DB Creation Main"

    Public Sub New()
        objUtilities = New Utilities
    End Sub

    Public Function CreateTables() As Boolean
        Try
            objUtilities.CreateTable("ME_RETAIL", "DBCONFIG(NEWERA)TABLE", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            objUtilities.AddAlphaField("@ME_RETAIL", "VERSION", "VERSION", 100)
            Dim oRs As SAPbobsCOM.Recordset
            oRs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery("SELECT * FROM ""@ME_RETAIL"" where ""U_VERSION"" = '" & Version & "'")
            Dim iDBConfigRecordCount As Integer = oRs.RecordCount
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
            If iDBConfigRecordCount = 0 Then
                objMain.objApplication.StatusBar.SetText("Your Database will now be upgraded to " + Version + ". Please Wait... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                '------------------------------------------------------------------------------

                'objMain.objUtilities.AddFloatField("PDN1", "DISCA", "Discount Amount", SAPbobsCOM.BoFldSubTypes.st_Quantity)


                objMain.objUtilities.AddAlphaField("ORPC", "BASEEN", "Base Entry", 30)
                objMain.objUtilities.AddAlphaField("OADM", "SERIES", "SERIES", 200)
                objMain.objUtilities.AddAlphaField("OITM", "CATEGORY", "CATEGORY", 20)
                objMain.objUtilities.AddAlphaField("OITM", "UOM", "UOM", 200)
                objMain.objUtilities.AddAlphaField("OITM", "CountryOrg", "CountryOrg", 200)
                objMain.objUtilities.AddAlphaField("OITM", "AcceptWMQty", "AcceptWMQty", 200)
                objMain.objUtilities.AddAlphaField("OITM", "3SUBGROUP", "3SUBGROUP", 200)
                objMain.objUtilities.AddFloatField("OITM", "WEIGHT", "Weight", SAPbobsCOM.BoFldSubTypes.st_Quantity)
                objMain.objUtilities.AddAlphaField("OITM", "Department", "Department", 110)
                ''Ramya
                objMain.objUtilities.AddAlphaField("OITM", "U_brand", "Brand", 100)
                objMain.objUtilities.AddAlphaField("OITM", "brand", "Brand", 100)
                objMain.objUtilities.AddAlphaField("OITM", "VSubSubGroup", "Sub Sub Group", 100)
                objMain.objUtilities.AddAlphaField("OITM", "VSubSubSubGroup", "Sub Sub Sub Group", 100)
                objMain.objUtilities.AddAlphaField("OITM", "BlockPurchase  ", "Block Purchase", 100)

                'Madhu
                objMain.objUtilities.addField("OITM", "REBEGBLE", "Rebate Eligible", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
                objMain.objUtilities.addField("OUSR", "Buyer", "Buyer", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
                objMain.objUtilities.addField("ORDR", "SMS", "SMS", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
                objMain.objUtilities.addField("ORDR", "SNEMAIL", "SNEMAIL", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
                objMain.objUtilities.addField("OITM", "SCALET", "Scale Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Tenant Scale,Internal Scale", "Tenant Scale,Internal Scale", "Internal Scale")
                objMain.objUtilities.AddAlphaField("OITM", "ExpiryDays", "Expire Days", 20)
                objMain.objUtilities.addField("OITM", "WoutPo", "Without Po", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
                objMain.objUtilities.addField("OITM", "SugarTax", "SugarTax", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "N")


                objMain.objUtilities.AddFloatField("OADM", "MTlce", "Margin Tolerance", SAPbobsCOM.BoFldSubTypes.st_Quantity)
                objMain.objUtilities.AddFloatField("OPDN", "MAmount", "Margin Amout", SAPbobsCOM.BoFldSubTypes.st_Quantity)
                objMain.objUtilities.AddFloatField("OPDN", "MDif", "Margin Difference", SAPbobsCOM.BoFldSubTypes.st_Quantity)

                objMain.objUtilities.AddAlphaField("OADM", "MrTle", "Margin Tolerance", 200)
                objMain.objUtilities.AddAlphaField("OADM", "PDFP", "PDF", 200)
                objMain.objUtilities.AddAlphaField("OADM", "SIGN", "Signature", 200)
                objMain.objUtilities.AddAlphaField("OPDN", "MAmount", "Margin Amount", 200)
                objMain.objUtilities.AddAlphaField("OPDN", "MDiff", "Margin Difference", 200)

                objMain.objUtilities.AddAlphaField("OITB", "ISCODE", "Item Series Code", 20)

                objMain.objUtilities.addField("OPDN", "CFF", "Cost Fluctution", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, SAPbobsCOM.BoFldSubTypes.st_None, "1,2,3", "Created,No Cost Fluctuation,Open", "3")

                objMain.objUtilities.addField("OPDN", "CMType", "Credit Memo Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Normal,Sell Out,Support", "Normal,Sell Out,Support", "Normal")


                objMain.objUtilities.AddFloatField("OITM", "REBQTY", "Rebate Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
                objMain.objUtilities.addField("OITM", "UnderPro", "Under Promotion", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")


                objMain.objUtilities.AddAlphaField("PRQ1", "NAME", "Name", 150)
                objMain.objUtilities.AddDateField("PRQ1", "DATE", "Date", SAPbobsCOM.BoFldSubTypes.st_None)
                objMain.objUtilities.AddFloatField("PRQ1", "QNTY", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
                objMain.objUtilities.addField("PRQ1", "VRFCAN", "Verification", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
                objMain.objUtilities.AddFloatField("PRQ1", "QNTY", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)



                'Madhu
                objMain.objUtilities.AddFloatField("PDN1", "DA", "Discount Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
                objMain.objUtilities.AddFloatField("OOAT", "TAmount", "Target Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
                objMain.objUtilities.AddFloatField("OOAT", "InvAmt ", "Invoice Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)

                objMain.objUtilities.AddFloatField("PDN1", "MARGN", "Margin", SAPbobsCOM.BoFldSubTypes.st_Percentage)

                objMain.objUtilities.addField("OWHS", "EXCHK", "Expiry Check", SAPbobsCOM.BoFieldTypes.db_Alpha, 4, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")

                objMain.objUtilities.AddAlphaField("OADM", "CNWhse", "Warehouse", 20)

                objMain.objUtilities.AddAlphaField("OPDN", "GUID", "GUID", 80)


                objMain.objUtilities.AddAlphaField("OPCH", "JEEntry", "JE Entry", 15) 'Created in SAP
                objMain.objUtilities.AddAlphaField("OVPM", "JEEntry", "JE Entry", 15)
                objMain.objUtilities.addField("PCH1", "Branch", "Branch", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "0", "0", "") 'Created in SAP
                objMain.objUtilities.addField("VPM4", "Branch", "Branch", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, SAPbobsCOM.BoFldSubTypes.st_None, "0", "0", "")

                objMain.objUtilities.AddInteger("OADM", "POCDT", "PO Auto close Date", SAPbobsCOM.BoFldSubTypes.st_None, 10)
                objMain.objUtilities.AddInteger("OCRD", "POCDT", "PO Auto close Date", SAPbobsCOM.BoFldSubTypes.st_None, 10)
                objMain.objUtilities.AddDateField("OCRD", "PDATE", "PO Auto Close Date", SAPbobsCOM.BoFldSubTypes.st_None)
                objMain.objUtilities.AddAlphaField("OCRD", "SUPNUM", "Supplier Posting", 30)

                objMain.objUtilities.AddFloatField("OCRD", "LCPSKU", "Listing charge per SKU", SAPbobsCOM.BoFldSubTypes.st_Quantity)

                objMain.objUtilities.addField("OCRD", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "O,R,S,N,A", "O - Open,R - Rejected,S - Submitted,N - Cancelled,A - Approved", "O")
                objUtilities.AddAlphaField("OCRD", "AppID", "Application ID", 50)
                objMain.objUtilities.AddDateField("OCRD", "DocDate", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
                objMain.objUtilities.addField("OCRD", "MOPay", "Mode of Payment", SAPbobsCOM.BoFieldTypes.db_Alpha, 25, SAPbobsCOM.BoFldSubTypes.st_None, "Cash,Cheque,Bank Transfer", "Cash,Cheque,Bank Transfer", "Cash")



                objMain.objUtilities.AddInteger("OITM", "PcPerPac", "No. of pc. per Pack", SAPbobsCOM.BoFldSubTypes.st_None, 10)
                objMain.objUtilities.AddFloatField("OITM", "BaseUnit", "BaseUnit", SAPbobsCOM.BoFldSubTypes.st_Measurement)
                objMain.objUtilities.addField("OITM", "AcceptWMQty", "Scale Item", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, "1,2", "Yes,NO", "NO")
                objMain.objUtilities.AddFloatField("OITM", "LOOSPER", "Loss Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)

                objMain.objUtilities.addField("OITM", "UOM", "UOM Dimension", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "-,Kg,L", "-,Kg,L", "-")
                objMain.objUtilities.addField("OWHS", "Transit", "Transit", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
                objMain.objUtilities.addField("OWHS", "MNWHS", "Main Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
                objMain.objUtilities.addField("OWHS", "TYP", "Warehouse Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 25, SAPbobsCOM.BoFldSubTypes.st_None, "MW,PW,DM,SW", "Main WareHouse ,Production WareHouse,Damage WareHouse,Scrap WareHouse", "MW")

                objMain.objUtilities.AddAlphaField("OWHS", "STGBLOC", "Staging Bin Location", 100)

                objMain.objUtilities.AddAlphaField("OACT", "RENTLAG", "Renatal Agreement", 1)
                objMain.objUtilities.addField("OOAT", "AUTMAN", "Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Auto,Manual", "Auto,Manual", "Auto")

                objMain.objUtilities.addField("OOAT", "RtnSup", "Return Support", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
                objMain.objUtilities.AddFloatField("OOAT", "TSA", "Total Support Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
                objMain.objUtilities.AddAlphaMemoField("OOAT", "SRN", "Supplier Registration Number", 500)
                '  objMain.objUtilities.addField("OOAT", "TypeDis", "Type of Display", SAPbobsCOM.BoFieldTypes.db_Alpha, 25, SAPbobsCOM.BoFldSubTypes.st_None, "-,Define New", "-,Define New", "-")
                objMain.objUtilities.AddAlphaField("OOAT", "TypeDis", "Type of Display", 54)
                objMain.objUtilities.AddAlphaField("OOAT", "MobileNo", "Mobile Number", 100)
                objMain.objUtilities.AddAlphaField("OOAT", "ContactPerson", "Contact Person", 100)
                objMain.objUtilities.AddFloatField("OAT1", "QTY", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)

                objMain.objUtilities.AddAlphaField("OAT1", "ItemCode", "ItemCode", 50)
                objMain.objUtilities.AddAlphaField("OAT1", "ITNM", "ItemName", 250)


                objMain.objUtilities.AddAlphaField("OWHS", "MEEXACCT", "Issue Account", 50)
                objMain.objUtilities.AddAlphaField("OWHS", "MEENTACT", "Receipt Account", 50)

                objMain.objUtilities.addField("PDN1", "SPLITTED", "Splitted", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
                'objMain.objUtilities.AddFloatField("PDN1", "Deliver", "Delivered Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
                objMain.objUtilities.AddFloatField("PDN1", "Delivr", "Delivered Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)

                objMain.objUtilities.AddFloatField("PDN1", "GROSSQTY", "Gross Quantity", SAPbobsCOM.BoFldSubTypes.st_Sum)
                objMain.objUtilities.AddFloatField("PDN1", "GROSSTOT", "Gross Total", SAPbobsCOM.BoFldSubTypes.st_Sum)

                objMain.objUtilities.AddAlphaField("OWTR", "VehicleNo", "VehicleNo", 100)

                'BusinessPartners
                objMain.objUtilities.AddFloatField("OCRD", "Discount", "Line Discount", SAPbobsCOM.BoFldSubTypes.st_Percentage)
                objMain.objUtilities.AddAlphaField("OCRD", "SUPREGDE", "SuppReg DocEntry", 30)
                objMain.objUtilities.AddAlphaField("OCRD", "ITMLIST", "ItemList DocEntry", 30)
                'objMain.objUtilities.AddAlphaField("OCRD", "SuppSt", "Supplier Status", 50)

                objMain.objUtilities.AddAlphaField("OPRQ", "PDTUSER", "Pdt User", 20)
                objMain.objUtilities.AddAlphaField("OPRQ", "UserName", "UserName", 100)
                objMain.objUtilities.AddAlphaField("OPRQ", "CCode", "Customer Code", 50)
                objMain.objUtilities.AddAlphaField("OPRQ", "CName", "Customer Name", 100) ''' Error



                objMain.objUtilities.AddAlphaField("OWHS", "INTRWHS", "InTransitWhs", 50)
                objMain.objUtilities.AddAlphaField("OWTR", "TRTYPE", "TransactionType", 30)
                objMain.objUtilities.AddAlphaField("OWTR", "INVTRREQ", "StockRequestNo", 30)
                objMain.objUtilities.AddAlphaField("OWTR", "INVTRREN", "StockRqstEntry", 30)
                objMain.objUtilities.AddAlphaField("OWTR", "MESTS", "DocStatus", 30)

                objMain.objUtilities.AddAlphaField("WTR1", "INVTRRLN", "StockRequestLine", 30)
                objMain.objUtilities.AddAlphaField("WTR1", "INVTRREQ", "StockRequestNo", 30)
                objMain.objUtilities.AddAlphaField("WTR1", "INVTRREN", "StockRqstEntry", 30)
                objMain.objUtilities.AddFloatField("WTR1", "MEOPNQTY", "Open Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
                objMain.objUtilities.AddFloatField("WTR1", "MERCEQTY", "Received Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
                objMain.objUtilities.AddFloatField("WTR1", "MEREQQTY", "STR Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
                objMain.objUtilities.AddFloatField("WTR1", "METSTQTY", "TotalSTR Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
                objMain.objUtilities.AddAlphaField("WTR1", "Addbarcode", "AddBarcode", 30)
                objMain.objUtilities.AddAlphaField("WTR1", "Packing", "Packing", 50)  '''' Error
                objMain.objUtilities.AddAlphaField("WTR1", "MIEntry", "MI Entry", 50)
                objMain.objUtilities.AddAlphaField("WTR1", "MIDocNum", "MI DocNum", 50)
                objMain.objUtilities.AddAlphaField("WTR1", "METWHSCD", "ToWarehouseCode", 50)
                objMain.objUtilities.AddAlphaField("WTR1", "METWHSNM", "ToWarehouseName", 100)
                objMain.objUtilities.AddAlphaField("WTR1", "METBPID", "ToBranch", 50)

                objMain.objUtilities.AddAlphaField("IGE1", "INVTRRLN", "StockRequestLine", 30)
                objMain.objUtilities.AddAlphaField("IGE1", "INVTRREQ", "StockRequestNo", 30)
                objMain.objUtilities.AddAlphaField("IGE1", "INVTRREN", "StockRqstEntry", 30)
                objMain.objUtilities.AddFloatField("IGE1", "MEREQQTY", "STR Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
                objMain.objUtilities.AddFloatField("IGE1", "METSTQTY", "TotalSTR Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
                objMain.objUtilities.AddAlphaField("IGE1", "Addbarcode", "AddBarcode", 30)
                objMain.objUtilities.AddAlphaField("IGE1", "Packing", "Packing", 50)
                objMain.objUtilities.AddAlphaField("IGE1", "METWHSCD", "ToWarehouseCode", 50)
                objMain.objUtilities.AddAlphaField("IGE1", "METWHSNM", "ToWarehouseName", 100)
                objMain.objUtilities.AddAlphaField("IGE1", "METBPID", "ToBranch", 50)

                'FOC OPOR
                objMain.objUtilities.addField("POR1", "FOCGrp", "FOC Group", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "-,A,B,C,D,E,F,G,H,I,J", "-,A,B,C,D,E,F,G,H,I,J", "-")
                objMain.objUtilities.AddFloatField("POR1", "normalqty", "Normal Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
                objMain.objUtilities.AddFloatField("POR1", "FOCQTY", "FOC Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
                objMain.objUtilities.AddFloatField("POR1", "suppprice", "Supplier Price", SAPbobsCOM.BoFldSubTypes.st_Price)
                objMain.objUtilities.AddAlphaField("POR1", "item", "item", 20)
                objMain.objUtilities.AddAlphaField("POR1", "ItemRef", "ItemRef", 20)
                objMain.objUtilities.AddAlphaField("POR1", "Packing", "Packing", 50)
                objMain.objUtilities.AddAlphaField("POR1", "Barcode", "BarCode", 30)
                objMain.objUtilities.AddAlphaField("POR1", "Addbarcode", "Add Barcode", 30)
                objMain.objUtilities.AddFloatField("POR1", "Uomqty", "UOM Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
                objMain.objUtilities.AddFloatField("POR1", "RetailPrice", "RetailPrice", SAPbobsCOM.BoFldSubTypes.st_Price)
                objMain.objUtilities.AddFloatField("POR1", "DiscountAmt", "DiscountAmt", SAPbobsCOM.BoFldSubTypes.st_Price)
                objMain.objUtilities.AddFloatField("POR1", "SQTY", "Sum Of Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)

                objMain.objUtilities.AddAlphaField("POR1", "BRNM", "BarCode Name", 250)

                objMain.objUtilities.AddFloatField("POR1", "BASE", "Base", SAPbobsCOM.BoFldSubTypes.st_Sum)
                objMain.objUtilities.AddFloatField("POR1", "BAD", "Base after discount", SAPbobsCOM.BoFldSubTypes.st_Sum)
                objMain.objUtilities.AddFloatField("POR1", "EXT", "Ex Tax(50% of base)", SAPbobsCOM.BoFldSubTypes.st_Sum)
                objMain.objUtilities.AddFloatField("POR1", "DSAT", "Discount Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)

                objMain.objUtilities.AddFloatField("POR1", "SuppPriBeDis", "Supplier Price Before Disc", SAPbobsCOM.BoFldSubTypes.st_Sum)
                objMain.objUtilities.AddFloatField("POR1", "SuppDiscPer", "Supplier Discount Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)

                objMain.objUtilities.AddFloatField("POR1", "DISCA", "Disc Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)

                objMain.objUtilities.AddAlphaField("POR1", "ForeNm", "Foreign Name", 100)
                objMain.objUtilities.AddAlphaField("POR1", "ASale", "Average Sale", 50)
                objMain.objUtilities.AddAlphaField("POR1", "URemarks", "Remarks", 200)
                objMain.objUtilities.addField("POR1", "SplitKitItems", "SplitKitItems", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
                objMain.objUtilities.addField("POR1", "LineValid", "Line Validation", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")

                objMain.objUtilities.addField("POR1", "LPValidate", "Least price validation", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "N")
                objMain.objUtilities.AddFloatField("POR1", "LVNDPRC", "Least Vendor Price", SAPbobsCOM.BoFldSubTypes.st_Price)
                objMain.objUtilities.AddAlphaField("POR1", "LPVCD", "LeastPrice VendorCode", 100)
                objMain.objUtilities.AddAlphaField("POR1", "LPVNM", "LeastPrice VendorName", 150)
                objMain.objUtilities.addField("POR1", "BelowCst", "Below Cost", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
                objMain.objUtilities.AddFloatField("POR1", "HDAYS", "Holding Days", SAPbobsCOM.BoFldSubTypes.st_Quantity)
                objMain.objUtilities.AddFloatField("POR1", "AQTY", "Aging Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
                objMain.objUtilities.AddFloatField("POR1", "MNSL", "Monthly Sales", SAPbobsCOM.BoFldSubTypes.st_Price)

                objMain.objUtilities.AddFloatField("POR1", "LTOTAL", "Line Total", SAPbobsCOM.BoFldSubTypes.st_Price)
                objMain.objUtilities.AddFloatField("POR1", "VATAMT", "VAT Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
                objMain.objUtilities.AddFloatField("POR1", "GTOTAL", "Gross Total", SAPbobsCOM.BoFldSubTypes.st_Price)
                objMain.objUtilities.AddFloatField("OPOR", "SPOTOTAL", "Supplier PO Total", SAPbobsCOM.BoFldSubTypes.st_Price)

                objMain.objUtilities.AddFloatField("POR1", "UTXAMT", "Untaxed Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
                objMain.objUtilities.AddFloatField("POR1", "EXTAX", "Excise Tax", SAPbobsCOM.BoFldSubTypes.st_Price)

                objMain.objUtilities.AddFloatField("OPOR", "TUTXAMT", "Total Untaxed Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
                objMain.objUtilities.AddFloatField("OPOR", "TEXTAX", "Total Excise Tax", SAPbobsCOM.BoFldSubTypes.st_Sum)
                objMain.objUtilities.addField("OPOR", "VEGITEM", "FreshKart PR", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "N")

                objMain.objUtilities.addField("VPM1", "Posted", "Posted", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "N")
                objMain.objUtilities.addField("VPM1", "Issued", "Issued", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "N")
                objMain.objUtilities.addField("VPM1", "Approved", "Approved", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "N")
                objMain.objUtilities.AddAlphaField("VPM1", "RECNM", "Receiver Name", 150)

                objMain.objUtilities.addField("OACT", "PDCReq", "PDCReq", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "N")
                objMain.objUtilities.addField("OACT", "PCA", "Petty Cash Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "N")
                objMain.objUtilities.AddAlphaField("OACT", "DepositeA", "Deposite Account", 50)

                objMain.objUtilities.AddAlphaField("OPOR", "CardCode", "Card Code", 100)
                objMain.objUtilities.AddAlphaField("OPOR", "CardName", "Card Name", 100)

                objMain.objUtilities.AddDateField("OPOR", "VDATE", "Validity Date", SAPbobsCOM.BoFldSubTypes.st_None)
                objMain.objUtilities.AddFloatField("OPOR", "EXT", "Excise Tax total", SAPbobsCOM.BoFldSubTypes.st_Sum)
                objMain.objUtilities.AddFloatField("OPOR", "SQTY", "Sum Of Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)

                objMain.objUtilities.AddAlphaField("OBPL", "FileLoc", "File Location", 250)
                objMain.objUtilities.AddAlphaField("OBPL", "PricelistNo", "PricelistNo", 10)
                'FOC

                objMain.objUtilities.AddAlphaField("OITM", "ItemSize", "ItemSize", 20)
                objMain.objUtilities.addField("OITT", "KITITEM", "KitItem", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
                'incomingpayment
                objMain.objUtilities.AddAlphaField("ORCT", "OutNum", "Outlet Number", 100)
                '--
                objMain.objUtilities.AddAlphaField("OIGE", "GRPO", "GRPONo.", 30)
                objMain.objUtilities.AddAlphaField("OIGE", "TYPE", "Type", 30)
                objMain.objUtilities.AddAlphaField("OIGE", "SUPPLIER", "Supplier", 50)
                objMain.objUtilities.AddAlphaField("OIGE", "MIDNUM", "MIDocNum", 50)
                objMain.objUtilities.AddAlphaField("OIGE", "MIDENT", "MIEntry", 50)
                objMain.objUtilities.AddAlphaField("OIGE", "ADDON", "ADDON", 20)

                objMain.objUtilities.AddAlphaField("OIGE", "ADDON", "ADDON", 20)
                objMain.objUtilities.AddAlphaField("OIGE", "GUID", "GUID", 254)
                objMain.objUtilities.AddAlphaField("OIGE", "TOWN", "To Warehouse Name", 54)
                objMain.objUtilities.AddAlphaField("OIGE", "TOBNCH", "To Branch", 54)

                objMain.objUtilities.AddAlphaField("OOAT", "SUPREBT", "Supplier Rebate", 30)
                objMain.objUtilities.AddAlphaField("OOAT", "SUPPOST", "Supplier Posting", 30)
                objMain.objUtilities.AddAlphaField("OOAT", "PMNTDTLS", "Payment Details", 30)
                objMain.objUtilities.AddAlphaField("OOAT", "GLCODE", "AccountCode", 50)
                objMain.objUtilities.AddAlphaField("OOAT", "GLNAME", "AccountName", 100)
                objMain.objUtilities.AddAlphaField("OOAT", "AUTMAN", "AUTMAN", 100)
                objMain.objUtilities.AddAlphaField("OOAT", "BranchID", "Branch Selection", 100)
                objMain.objUtilities.AddAlphaField("OOAT", "PMNTDTLS", "Payment Terms", 100)
                objMain.objUtilities.AddAlphaField("OOAT", "PAYMENT", "Payment Details", 100)

                objMain.objUtilities.AddAlphaField("OJDT", "Remarks", "Remarks", 50)
                objMain.objUtilities.AddAlphaField("OITM", "PRCUPNO", "PriceUpdateNo", 50)
                objMain.objUtilities.AddAlphaField("OBCD", "ArbName", "Arabic Name", 50)
                objMain.objUtilities.AddAlphaField("OBCD", "PICK", "Packing ", 50)

                objMain.objUtilities.AddAlphaField("OBCD", "POSCA", "Pos Scann", 50)
                objMain.objUtilities.AddAlphaField("OBCD", "PomDocNum", "Promotion Document Number", 50)
                objMain.objUtilities.AddAlphaField("OBCD", "CatalogNo", "Catalog Number", 100)
                ' objMain.objUtilities.addField("OBCD", "POSCAN", "Pos Scan", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")


                objMain.objUtilities.AddAlphaField("OCRD", "PDTPassword", "PDT Password", 10)

                objMain.objUtilities.AddAlphaField("OUSR", "SCGWLK", "SCGW Lock", 1)
                '' objMain.objUtilities.AddAlphaField("OUSR", "Branch", "PDF", 200)
                objMain.objUtilities.AddAlphaField("INV1", "PGWPEN", "Production", 50)
                objMain.objUtilities.AddAlphaField("INV1", "PGWIEN", "Issue", 50)
                objMain.objUtilities.AddAlphaField("INV1", "PGWREN", "Receipt", 50)


                objMain.objUtilities.AddAlphaField("OWOR", "GUID", "OWOR", 254)
                objMain.objUtilities.AddAlphaField("OWOR", "PGWINVEN", "Invoice", 50)
                objMain.objUtilities.AddAlphaField("OIGE", "PGWINVEN", "Invoice", 50)
                objMain.objUtilities.AddAlphaField("OIGN", "PGWINVEN", "Invoice", 50)
                objMain.objUtilities.AddAlphaField("OIGN", "GI", "Goods Issue No", 15)
                objMain.objUtilities.AddAlphaField("OIGE", "GR", "Goods Receipt No", 15)
                'objMain.objUtilities.AddAlphaField("OIGN", "GIDoc", "Goods Issue DocNo", 15)
                'objMain.objUtilities.AddAlphaField("OIGE", "GRDoc", "Goods Receipt DocNo", 15)

                objUtilities.CreateTable("ME_STOCK", "STOCKCOUNT_TABLE", SAPbobsCOM.BoUTBTableType.bott_NoObject)
                objUtilities.AddAlphaField("@ME_STOCK", "DOCNUM", "DOCNUM", 100)
                objUtilities.AddAlphaMemoField("@ME_STOCK", "ERROR", "ERROR", 6000)
                objUtilities.AddAlphaMemoField("@ME_STOCK", "STATUS", "STATUS", 6000)

                Me.GRPODOCRp()
                objMain.GRPODOCRpUDO()
                Me.BarChk()
                Me.GRPOItmRp()
                Me.OpenDisp()
                Me.TransferRep()
                Me.Stocksumm()
                Me.StockWHSE()
                Me.APInvoiceItem()
                Me.CatWiseOB()
                Me.ExciseItems()
                Me.ARDocWise()
                Me.ARItemWise()
                Me.CreateExpiryTable()
                Me.VR_COUNTRY()
                Me.PromotiononVoucherPurchase()
                Me.PromoCostReduction()
                objMain.PromoCostReductionUDO()

                Me.BarcodeReport()
                objMain.BarcodeReportUDO()
                Me.BarcodeFilter()
                objMain.BarcodeFilterUDO()

                Me.ItemMasterUploadPreview()
                objMain.ItemMasterUploadPreviewUDO()
                STOCKDET()
                objMain.STOCKDET()
                Me.PACPPriceChange()
                objMain.PACPPriceChangeUDO()

                Me.PromoCostReversion()
                objMain.PromoCostReversionUDO()
                Me.CreatePDTROLES()
                Me.BarcodeSetup()
                Me.BarcodePrinter()
                Me.Barcode()
                Me.StandardDocuments()
                Me.ClearanceSales()
                Me.SupplierAgreementPosting()
                Me.CreatePDTStockCount()
                Me.CreatePUIGA()
                Me.CreatePUBA()
                Me.CreateStockCountGenerationwizard()
                ' Me.CreateEmployee()
                objMain.objUtilities.AddAlphaFieldLinkedTable("OHEM", "PDTRole", "PDT Role", 50, "", "ME_PDTROLES")
                objMain.objUtilities.AddAlphaField("OHEM", "PDTPassword", "PDT Password", 10)

                Me.SupRebateWizard()
                Me.SupDeposit()
                Me.PettyCashMaster()
                Me.CreateStockTransferRequest()
                Me.CreatePMAPINVTable()
                Me.CreatePMAPCMTable()
                Me.CreateSupplierAgreementTable()
                Me.CreateMaterialInward()
                Me.CreatAddonConfigScrn()
                Me.CreateAutomaticProduction()
                Me.CreateItemList()
                Me.CreateSubItemList()
                Me.CreateOutlet()
                Me.CreateSupplierRebateTable()
                Me.CreatePriceUpdateTable()
                Me.CreateFuturePriceUpdateTable()
                Me.CreatePDTSeriesConfig()
                Me.CreatePDCPostingTable()
                Me.PDTUserWiseBranchDetails()
                Me.VehicleMasterNoObject()
                Me.CreateEmployee()
                Me.CreateListReport()
                Me.BarcodeRequest()
                objMain.CreateListReportUDO()
                Me.CreateCategory()
                objMain.CreateCategoryUDO()

                Me.CreatePaymentDetailsScreen()
                Me.BuyGetFreeItem()
                Me.BuyGetLeastItem()
                Me.ItemWiseDiscount()
                Me.CreatePriceConfiguration()
                Me.PromotiononBillTotal()
                Me.BuyNGetOneFreeItem()
                Me.CreateCreditNotetable()
                Me.CREATEUPDATEPRICESERVICE()
                objMain.PLUDO()

                Me.Itemmasterdataupload()
                objMain.Itemmasterdataupload1()

                Me.CentralizedPriceChangeRequest()
                objMain.CentralizedPriceChangeRequestUDO()

                objMain.CreateCreditnoteUDO()
                'Approval

                Me.POSB()
                objMain.POSBUDO()
                Me.CreateAPPROVALTemplates()
                Me.CreateAPPROVALSTAGES()
                Me.CreateTenxLogSettinTbl()
                Me.CreateTENXLOGTable()

                Me.CreatMultipleStockTransferRequest()
                Me.CreateWarehouse()

                Me.PriceChangeRequest()
                Me.TNXBranch()
                objMain.BranchUDO()

                '----madhu 
                ' Me.CreateDraftTable()
                ' objMain.DraftUDO()

                Me.CreateDDTable()

                Me.SalesBlanketAgreementMaster()
                objMain.SalesBlanketAgreementMasterUDO()
                objMain.DDUDO()

                'ProductEnquiry
                Me.ProductEnquiry()
                objMain.ProductEnquiryUDO()

                objMain.CreateWarehouseUDO()
                objMain.CreateMultipleStockTransferRequestUDO()
                objMain.BarcodeRequestUDO()
                objMain.PreciChangeRequestUDO()
                objMain.CreateApprovalTemplatesUDO()

                objMain.CreateAPPROVALSTAGESUDO()
                objMain.CreateAddonLogUDO()
                objMain.CreateSupRebateWizardUDO()

                objMain.CreateExpiryDateUDO()

                objMain.SupDepositUDO()
                objMain.PettyCashUDO()
                objMain.CreateStockTransferRequestUDO()
                objMain.CreateStockcountWizardUDO()
                objMain.CreateSupplierAgreementUDO()
                objMain.CreateMaterialInwardUDO()
                objMain.CreateAddonConfigUDO()
                objMain.CreateAutomaticProductionUDO()
                objMain.CreatePriceConfigurationUDO()
                objMain.CreateItemListUDO()
                objMain.SubItemListUDO()
                objMain.OutletUDO()
                objMain.CreateSupplierRebateUDO()
                objMain.CreatePriceUpdateUDO()
                objMain.CreateFuturePriceUpdateUDO()
                objMain.CreatePDCPostingUDO()
                'objMain.CreateEmployeeDetailsUDO()
                objMain.CreateBarcodeUDO()
                objMain.CreateClearanceSalesUDO()
                objMain.CreateSuplAgreeUDO()
                objMain.CreatePDTStockCountUDO()
                objMain.CreatePaymentDetailsScreenUDO()
                objMain.CreatePUIGAUDO()
                objMain.CreatePUBAUDO()
                objMain.CreateBuyGetFreeItemsUDO()
                objMain.CreateBuyGetLeastItemsUDO()
                objMain.CreateItemsWiseDiscountUDO()
                ' objMain.CreatePromotiononBillTotalUDO()
                objMain.CreateBuyNGetOneFreeItemUDO()
                objMain.CreatePromotionOnVoucherPurchaseUDO()

                Me.RENTPROMO()
                objMain.RENTPROMOUDO()

                Me.LicenceNew()
                objMain.CreateLicenceNewUDO()
                Me.BrandMaster()
                objMain.BrandMasterUDO()
                Me.SubGroup()
                objMain.SubGroupUDO()
                Me.SubSubGroup()
                objMain.SubSubGroupUDO()
                Me.SubSubSubGroup()
                Me.VR_SUBGROUP()
                Me.VR_SUBSUBGROUP()
                Me.VR_SUBSUBSUBGROUP()
                objMain.SubSubSubGroupUDO()
                Me.DepartmentMaster()
                objMain.DepartmentMasterUDO()
                Messanger_Log()
                Tnx_Retail_Log()
                ApCreditMemoNoObject()
                MessageConfig()
                objMain.MessageConfigUDO()
                SMS_EMail_Config_Table()
                Configuration()
                objMain.ConfigurationUDO()
                objMain.SMS_EmailUDO()
                'Excide Duty
                Me.Item_Master_Data()
                Me.AP_Invoice()
                Me.APCreditMemo()
                Me.Generalsettings()
                Me.Freight()
                Me.AR_Invoice()
                Me.CompanyLevel()
                Me.Tenx_Temp1()
                Me.Tenx_Temp2()

                objMain.objUtilities.AddDateField("OADM", "TRBDATE", "Transactions Blocking date", SAPbobsCOM.BoFldSubTypes.st_None)

                '--------------------------------------------------------------------------------------
                'Close DB Script

                Me.RunSQLScripts()
                'Me.QueryManager()'madhu commented query manager called in main
                Me.AddUserFormAuthorizations()
                'Madhu comment 14/06/2026
                ' Me.AddAlerts()

                objUtilities.UpdateTransactionNotification()
                ' objUtilities.AddDataToNoObjectTable("ME_RETAIL", DBCode.Trim, DBName.Trim, "U_VERSION", Version)
                Dim err As String

                err = objUtilities.AddDataToNoObjectTable("ME_RETAIL", Version, Version, "U_VERSION", Version)

                If err <> "" Then
                    MsgBox(err)
                End If
                objMain.objApplication.StatusBar.SetText("Your Database has now been upgraded to Version " + Version + ".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return False
        End Try
        Return True
    End Function
#End Region

#Region "Create Tables"
    Sub VR_COUNTRY()
        objMain.objUtilities.CreateTable("VR_COUNTRY", "VR_COUNTRY", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        ''test2
    End Sub
    Sub Itemmasterdataupload()
        objUtilities.CreateTable("TNX_ITEMUP", "Item Master Data Upload", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddDateField("@TNX_ITEMUP", "DDATE", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUtilities.AddAlphaField("@TNX_ITEMUP", "STATUS", "Status", 10)
        objUtilities.AddAlphaField("@TNX_ITEMUP", "AppID", "Application ID", 50)
        objUtilities.AddAlphaField("@TNX_ITEMUP", "STAT", "Status", 10)
        objUtilities.AddAlphaField("@TNX_ITEMUP", "DocNoIL", "Document No ItemListing", 40)
        objUtilities.AddAlphaField("@TNX_ITEMUP", "AppIDIL", "ApplicationID ItemListing", 50)
        objMain.objUtilities.addField("@TNX_ITEMUP", "VRCT", "Verification Completed", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")

        objUtilities.CreateTable("TNX_ITEMUP_C0", "Item Master Upload Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "KEYC", "Key Code", 50)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "ITEMN", "Item Name", 200)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "FORIE", "Foriegn Name", 200)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "CARD", "Card Code", 50)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "ITEMG", "Item Group", 200)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "SUBGR", "Sub Group", 200)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "SSGP", "Sub Sub Group", 200)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "SUGRP", "Sub Sub Sub Group", 200)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "BRAND", "BRAND", 150)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "CATGO", "Category", 100)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "SCALE", "Scale Item", 100)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "BUNIT", "Base Unit", 50)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "COUNT", "Country", 100)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "ITEMS", "Item Size", 50)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "INVENT", "Inventory Item", 10)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "PURCH", "Purchase Item", 10)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "SALITM", "Sales Item", 10)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "CNum", "Catalog Numbere", 50)


        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "ITMCD", "Item Code", 254)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "STATS", "Status", 254)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "PStst", "Priceupdate Status", 254)
        objMain.objUtilities.AddFloatField("@TNX_ITEMUP_C0", "NOP", "No. of Pics Per Pack", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_ITEMUP_C0", "BaseUt", "Base Unit", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "UOMD", "UOM Dimension", 50)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "PstatV", "Vendor Price Update Status", 254)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "MPUS", "Ministry Price Update Status", 254)
        objMain.objUtilities.addField("@TNX_ITEMUP_C0", "LPO", "LPO", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")

        objMain.objUtilities.addField("@TNX_ITEMUP_C0", "WITPO", "Without PO", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
        objMain.objUtilities.addField("@TNX_ITEMUP_C0", "REBEGBLE", "Rebate Eligible", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
        '   objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "SuTaxAmt", "sugar Tax Amount", 254)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "VIC", "Vendor Item Code", 50)
        'objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "MARKUP", "Markup", 50)
        ' objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "MARGIN", "Margin", 50)

        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "PVAT", "Purchase VAT", 100)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C0", "SVAT", "Sales VAT", 100)



        objUtilities.CreateTable("TNX_ITEMUP_C1", "Item Master Upload UOM", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C1", "KEYCDE", "Key Code", 50)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C1", "UOMCDE", "UOM Code", 50)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C1", "UOMUNT", "UOM Units", 50)

        objUtilities.CreateTable("TNX_ITEMUP_C2", "Item Master Upload BarCode", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C2", "KEYC", "Key Code", 50)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C2", "ITEMN", "Item Name", 200)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C2", "UOMCDE", "UOM Code", 50)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C2", "BarCode", "BarCode", 50)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C2", "BDesc", "BarCode Description", 254)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C2", "CNum", "Catalog Number", 50)
        ''objUtilities.AddAlphaField("@TNX_ITEMUP_C2", "FORIE", "Foreign Name", 254)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C2", "PREFER", "Preferred", 10)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C2", "ADESC", "Arabic Description", 254)

        ''
        objMain.objUtilities.addField("@TNX_ITEMUP_C2", "STAX", "Sugar Tax", SAPbobsCOM.BoFieldTypes.db_Alpha, 4, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
        ''objMain.objUtilities.AddFloatField("TNX_ITEMUP_C2", "PRCNT", "Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_ITEMUP_C2", "SUTAXAMT", "Sugar Tax Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_ITEMUP_C2", "ETAX", "Excise Tax", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_ITEMUP_C2", "Price", "Price", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_ITEMUP_C2", "Margin", "Margin", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_ITEMUP_C2", "Markup", "Markup", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C2", "Remark", "Remark", 254)

        objMain.objUtilities.addField("OBCD", "STAX", "Sugar Tax", SAPbobsCOM.BoFieldTypes.db_Alpha, 4, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
        objMain.objUtilities.AddFloatField("OBCD", "SUTAXAMT", "Sugar Tax Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("OBCD", "ETAX", "Excise Tax", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("OBCD", "Price", "Price", SAPbobsCOM.BoFldSubTypes.st_Price)


        objUtilities.AddLinkField("@TNX_ITEMUP", "ATTPATH", "Attachment Path", 2000, SAPbobsCOM.BoFldSubTypes.st_Link)


        objMain.objUtilities.CreateTable("TNX_ITEMUP_C3", "ItemMasterBranch Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C3", "ItemC", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C3", "PriceId", "PriceList Id", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C3", "PriceLt", "Price List", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C3", "BrnchId", "Branch Id", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C3", "BrachNm", "Branch Name", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C3", "UomEtry", "Uom Entry", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C3", "UOMCode", "Uom Code", 50)
        objMain.objUtilities.AddFloatField("@TNX_ITEMUP_C3", "OldPrice", "Old Price", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_ITEMUP_C3", "NewPrice", "New price", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddDateField("@TNX_ITEMUP_C3", "Update", "Update Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_ITEMUP_C3", "EndDate", "End Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C3", "UserName", "User Code", 50)

        objMain.objUtilities.AddFloatField("@TNX_ITEMUP_C3", "Cost", "Cost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C3", "PriceList", "PriceList ID", 50)

        objUtilities.CreateTable("TNX_ITEMUP_C4", "Item Master Vendor", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C4", "Code", "Key Code", 50)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C4", "UOMCode", "UOM Code", 50)
        objUtilities.AddAlphaField("@TNX_ITEMUP_C4", "CCode", "CardCode", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C4", "PriceId", "PriceList Id", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C4", "PriceLt", "Price List", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C4", "UEntry", "Uom Entry", 50)
        objMain.objUtilities.AddFloatField("@TNX_ITEMUP_C4", "Price", "Old Price", SAPbobsCOM.BoFldSubTypes.st_Price)

        objUtilities.CreateTable("TNX_ITEMUP_C5", "Item Master Ministryprice", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C5", "Code", "Key Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C5", "PriceId", "PriceList Id", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C5", "PriceLt", "Price List", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C5", "BrnchId", "Branch Id", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C5", "BrachNm", "Branch Name", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C5", "UomEtry", "Uom Entry", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C5", "UOMCode", "UOM Code", 50)
        objMain.objUtilities.AddFloatField("@TNX_ITEMUP_C5", "Price", "Old Price", SAPbobsCOM.BoFldSubTypes.st_Price)

        objUtilities.CreateTable("TNX_ITEMUP_C6", "Item Branch", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C6", "BRANCH", "Branch ", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C6", "WAREHOUSE", "Warehouse Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C6", "CHECK", "Block Purchase", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C6", "BPLID", "Branch Id ", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMUP_C6", "WHSCODE", "Warehouse Code ", 100)




    End Sub

    Sub CentralizedPriceChangeRequest()
        objMain.objUtilities.CreateTable("TNX_CP", "Centralized Price Request", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddDateField("@TNX_CP", "FDate", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_CP", "TDate", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_CP", "Branch", "Branch", 50)
        objMain.objUtilities.AddAlphaField("@TNX_CP", "PUType", "Price Update Type", 80)
        objMain.objUtilities.AddAlphaField("@TNX_CP", "DStatus", "Doc Status", 80)
    End Sub
    Sub ProductEnquiry()
        'Product enquiry
        objUtilities.CreateTable("TNX_PRDENQUIRY", "Product Enquiry", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddDateField("@TNX_PRDENQUIRY", "ENQDATE", "Enquire Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddDateField("@TNX_PRDENQUIRY", "DDate", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_PRDENQUIRY", "Branch", "Branch", 254)
        objMain.objUtilities.addField("@TNX_PRDENQUIRY", "PEStatu", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 19, SAPbobsCOM.BoFldSubTypes.st_None, "Open,Close", "Open,Close", "Open")


        objMain.objUtilities.CreateTable("TNX_PRDENQUIRY_C0", "Product Enquiry child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PRDENQUIRY_C0", "ITMNM", "ItemName", 250)
        objMain.objUtilities.AddAlphaField("@TNX_PRDENQUIRY_C0", "Desc", "Description", 254)
        objMain.objUtilities.AddFloatField("@TNX_PRDENQUIRY_C0", "Qty", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_PRDENQUIRY_C0", "BUYER", "Buyer/manager", 150)
        objMain.objUtilities.AddAlphaField("@TNX_PRDENQUIRY_C0", "Remarks", "Remarks", 254)
        objMain.objUtilities.AddLinkField("@TNX_PRDENQUIRY_C0", "ATTCH", "Attachment", 80, SAPbobsCOM.BoFldSubTypes.st_Link)


    End Sub
    Sub CreateDDTable()
        objMain.objUtilities.CreateTable("SBO_DD", "DDTable", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.CreateTable("SBO_DD1", "DD1Table", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@SBO_DD", "AppID", "Appln ID Field", 20)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "AppStat", "Appln Status Field", 20)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "Table", "Table", 40)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "Doc", "Document", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "statusN", "Status New", 20)

        objMain.objUtilities.AddAlphaField("@SBO_DD", "docnum", "Document Number", 20)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "tempid", "Template ID", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "status", "Status", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "userid", "UserID", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "objtype", "objtype", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "Isdraft", "Isdraft", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "Remarks", "Remarks", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "Appcnt", "Application", 100)
        objMain.objUtilities.AddDateField("@SBO_DD", "docdate", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "Stage", "Stage", 100)

        objMain.objUtilities.AddAlphaField("@SBO_DD1", "EmpId", "Employee Id", 50)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "Stage", "Stage", 50)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "statusN", "Status New", 20)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "TempId", "Template ID", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "Userid", "UserID", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "status", "Status", 100)
        objMain.objUtilities.AddDateField("@SBO_DD1", "cdate", "Create Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "ctime", "Create Time", 20)
        objMain.objUtilities.AddDateField("@SBO_DD1", "Udate", "Update Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "Utime", "Update Time", 20)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "Remarks", "Remarks", 200)

    End Sub
    Sub CreateDraftTable()
        objMain.objUtilities.CreateTable("SBO_OWDD", "HeaderNeDraft", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.CreateTable("SBO_WDD1", "RowsNewDraft", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@SBO_OWDD", "AppID", "Appln ID Field", 20)
        objMain.objUtilities.AddAlphaField("@SBO_OWDD", "AppStat", "Appln Status Field", 20)
        objMain.objUtilities.AddAlphaField("@SBO_OWDD", "Table", "Table", 40)
        objMain.objUtilities.AddAlphaField("@SBO_OWDD", "Doc", "Document", 100)
        objMain.objUtilities.AddAlphaField("@SBO_OWDD", "statusN", "Status New", 20)

        objMain.objUtilities.AddAlphaField("@SBO_OWDD", "docnum", "Document Number", 20)
        objMain.objUtilities.AddAlphaField("@SBO_OWDD", "tempid", "Template ID", 100)
        objMain.objUtilities.AddAlphaField("@SBO_OWDD", "status", "Status", 100)
        objMain.objUtilities.AddAlphaField("@SBO_OWDD", "userid", "UserID", 100)
        objMain.objUtilities.AddAlphaField("@SBO_OWDD", "objtype", "objtype", 100)
        objMain.objUtilities.AddAlphaField("@SBO_OWDD", "Isdraft", "Isdraft", 100)
        objMain.objUtilities.AddAlphaField("@SBO_OWDD", "Remarks", "Remarks", 100)
        objMain.objUtilities.AddAlphaField("@SBO_OWDD", "Appcnt", "Application", 100)
        objMain.objUtilities.AddDateField("@SBO_OWDD", "docdate", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@SBO_WDD1", "EmpId", "Employee Id", 50)
        objMain.objUtilities.AddAlphaField("@SBO_WDD1", "Stage", "Stage", 50)
        objMain.objUtilities.AddAlphaField("@SBO_WDD1", "statusN", "Status New", 20)
        objMain.objUtilities.AddAlphaField("@SBO_WDD1", "TempId", "Template ID", 100)
        objMain.objUtilities.AddAlphaField("@SBO_WDD1", "Userid", "UserID", 100)
        objMain.objUtilities.AddAlphaField("@SBO_WDD1", "status", "Status", 100)
        objMain.objUtilities.AddDateField("@SBO_WDD1", "cdate", "Create Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@SBO_WDD1", "ctime", "Create Time", 20)
        objMain.objUtilities.AddDateField("@SBO_WDD1", "Udate", "Update Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@SBO_WDD1", "Utime", "Update Time", 20)
        objMain.objUtilities.AddAlphaField("@SBO_WDD1", "Remarks", "Remarks", 200)

    End Sub
    Sub CreateCreditNotetable()
        objMain.objUtilities.CreateTable("TNXCRDNT", "Credit Note Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@TNXCRDNT", "CCODE", "Card Code", 50)
        objMain.objUtilities.AddAlphaField("@TNXCRDNT", "CNAME", "Card Name", 200)
        objMain.objUtilities.AddAlphaField("@TNXCRDNT", "VREFNO", "Ref No", 80)

        objMain.objUtilities.AddDateField("@TNXCRDNT", "DOCDT", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNXCRDNT", "DOCDDT", "DocDueDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNXCRDNT", "CRDNO", "Credit No", 80)
        objMain.objUtilities.AddAlphaField("@TNXCRDNT", "BRNCH", "Branch", 80)


        objMain.objUtilities.AddFloatField("@TNXCRDNT", "TOTVT", "Total Vat", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNXCRDNT", "DOCTTL", "DocTotal", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddAlphaMemoField("@TNXCRDNT", "RMKSS", "Remarks", 64000)
        objMain.objUtilities.AddAlphaField("@TNXCRDNT", "CDOCN", "DocNum", 80)
        objMain.objUtilities.AddAlphaField("@TNXCRDNT", "GRPON", "GRPO No", 80)




        objMain.objUtilities.CreateTable("TNXCRDNT_C0", "Credit Note Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNXCRDNT_C0", "BRCD", "Bar Code", 80)
        objMain.objUtilities.AddAlphaField("@TNXCRDNT_C0", "ICODE", "Item Code", 80)
        objMain.objUtilities.AddAlphaField("@TNXCRDNT_C0", "INAME", "Item Name", 100)
        objMain.objUtilities.AddAlphaField("@TNXCRDNT_C0", "UOM", "UOM", 20)

        objMain.objUtilities.AddFloatField("@TNXCRDNT_C0", "QUANTITY", "Item Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNXCRDNT_C0", "UPRICE", "Item Price", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@TNXCRDNT_C0", "TAX", "Vat", 30)

        objMain.objUtilities.AddFloatField("@TNXCRDNT_C0", "LTOTAL", "Line Total", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddAlphaField("@TNXCRDNT_C0", "RMKS", "Remarks", 200)

        objMain.objUtilities.AddFloatField("@TNXCRDNT_C0", "Cost", "Cost", SAPbobsCOM.BoFldSubTypes.st_Price)

    End Sub

    Sub CreateExpiryTable()
        objMain.objUtilities.CreateTable("TNXEDS", "Expiry Table", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@TNXEDS", "DOCTY", "DocType", 50)
        objMain.objUtilities.AddAlphaField("@TNXEDS", "BENTRY", "Base Entry", 50)
        objMain.objUtilities.AddAlphaField("@TNXEDS", "BNO", "Base Num", 50)
        objMain.objUtilities.AddDateField("@TNXEDS", "DOCDT", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNXEDS", "GUID", "UID", 80)


        objMain.objUtilities.CreateTable("TNXEDSC0", "EXpiry Table C0", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNXEDSC0", "BRCD", "BarCode", 100)
        objMain.objUtilities.AddAlphaField("@TNXEDSC0", "ITMCD", "ItemCode", 100)
        objMain.objUtilities.AddAlphaField("@TNXEDSC0", "ITMNM", "ItemName", 250)
        objMain.objUtilities.AddAlphaField("@TNXEDSC0", "UOM", "UOM", 50)
        objMain.objUtilities.AddDateField("@TNXEDSC0", "EDATE", "Expiry Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddFloatField("@TNXEDSC0", "QTY", "Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNXEDSC0", "BQTY", "Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddInteger("@TNXEDSC0", "BL", "BaseLine", SAPbobsCOM.BoFldSubTypes.st_None, 4)
        objMain.objUtilities.AddAlphaField("@TNXEDSC0", "RMKS", "Remarks", 200)
        objMain.objUtilities.addField("@TNXEDSC0", "SR", "Split Row", SAPbobsCOM.BoFieldTypes.db_Alpha, 4, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")


    End Sub
    Sub VehicleMasterNoObject()
        objMain.objUtilities.CreateTable("SBO_VEHICLEMSTR", "Vechicle Master", SAPbobsCOM.BoUTBTableType.bott_NoObject)
    End Sub
    Sub CreateEmployee()
        objMain.objUtilities.CreateTable("TNX_Custom", "EmployeeDetails", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@TNX_Custom", "CCode", "Customer Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_Custom", "CName", "Customer Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_Custom", "Status", "Customer Status", 254)

        objMain.objUtilities.CreateTable("TNX_Custom_C0", "Employee_Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_Custom_C0", "ECode", "Emp Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_Custom_C0", "EName", "Emp Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_Custom_C0", "Phn", "Phone Number", 50)


    End Sub
    Sub CreateCategory()
        objMain.objUtilities.CreateTable("TNX_CATERY", "CategorySplit", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@TNX_CATERY", "SKUS", "No Of Skus", 50)
        objMain.objUtilities.AddFloatField("@TNX_CATERY", "Total", "Total", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddAlphaField("@TNX_CATERY", "RowID", "U_RowID", 50)
        objMain.objUtilities.AddAlphaField("@TNX_CATERY", "CardCode", "CardCode", 50)
        objMain.objUtilities.AddAlphaField("@TNX_CATERY", "BNO", "Base No", 50)


        objMain.objUtilities.CreateTable("TNX_CATERY_C0", "CategorySplit_Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_CATERY_C0", "MainCat", "Main Category", 50)
        'objMain.objUtilities.AddAlphaField("@TNX_CATERY_C0", "SKUCount", "SKUS Count", 100)
        'objMain.objUtilities.AddAlphaField("@TNX_CATERY_C0", "Charge", "Charge Per Unit", 50)tnx
        'objMain.objUtilities.AddAlphaField("@TNX_CATERY_C0", "Listing", "Listing Charges", 50)

        objMain.objUtilities.AddFloatField("@TNX_CATERY_C0", "SKUCount", "SKUS Count", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_CATERY_C0", "Charge", "Charge Per Unit", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_CATERY_C0", "Listing", "Listing Charges", SAPbobsCOM.BoFldSubTypes.st_Sum)


    End Sub
    Sub BarcodeRequest()
        objMain.objUtilities.CreateTable("TNX_BCR", "Barcode Request Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@TNX_BCR", "BRN", "Barcode Request Number", 50)
        objMain.objUtilities.AddDateField("@TNX_BCR", "BRD", "Barcode Request Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.addField("@TNX_BCR", "Barcode", "Barcode Generate", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
        objMain.objUtilities.addField("@TNX_BCR", "PDTVRFID", "PDT Verified", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
        objMain.objUtilities.addField("@TNX_BCR", "StatusN", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 19, SAPbobsCOM.BoFldSubTypes.st_None, "Open,Completed", "Open,Completed", "Open")


        objMain.objUtilities.CreateTable("TNX_BCR_C0", "Barcode Request child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_BCR_C0", "PONum", "PO Number", 50)
        objMain.objUtilities.AddAlphaField("@TNX_BCR_C0", "ItemCode", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BCR_C0", "ItemDes", "Item Description", 50)
        objMain.objUtilities.AddAlphaField("@TNX_BCR_C0", "New", "New Barcode", 50)
        objMain.objUtilities.AddAlphaField("@TNX_BCR_C0", "Packing", "Packing Information", 50)
        objMain.objUtilities.addField("@TNX_BCR_C0", "Barcode", "Barcode", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
        objMain.objUtilities.AddAlphaField("@TNX_BCR_C0", "BCODE", "Barcode", 50)
        objMain.objUtilities.AddAlphaField("@TNX_BCR_C0", "BNAME", "Barcode Name", 100)


    End Sub
    Sub CreateListReport()
        objMain.objUtilities.CreateTable("TNX_LSRPT", "ListReport", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@TNX_LSRPT", "CardCode", "Card Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_LSRPT", "CardName", "Card Name", 100)
        objMain.objUtilities.AddDateField("@TNX_LSRPT", "FDate", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_LSRPT", "ToDate", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_LSRPT", "Branch", "Branch", 100)







    End Sub

    Sub SalesBlanketAgreementMaster()
        ' objMain.objUtilities.CreateTable("TNX_SBAMR", "SalesBlanketMasHeader", SAPbobsCOM.BoUTBTableType.bott_MasterData)


        objMain.objUtilities.CreateTable("TNX_SBAMRM", "SalesBlanketNew", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@TNX_SBAMRM", "Status", "Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_SBAMRM", "AbsID", "AbsID", 30)


    End Sub
    Sub TNXBranch()
        objMain.objUtilities.CreateTable("TNXBRANCH", "Branch Master1", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@TNXBRANCH", "BCode", "Branch Code", 30)
        objMain.objUtilities.AddAlphaField("@TNXBRANCH", "BName", "Branch Name", 100)
        'objMain.objUtilities.AddAlphaField("@TNXBRANCH", "DocNum", "Doc Number", 10)

        objMain.objUtilities.CreateTable("TNXBRANCH_C0", "BranchMaster1 child", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@TNXBRANCH_C0", "EAdd", "Email Address", 100)
        objMain.objUtilities.AddAlphaField("@TNXBRANCH_C0", "Name", "Name", 100)
        objMain.objUtilities.AddAlphaField("@TNXBRANCH_C0", "MobileNo", "Mobile Number", 100)
        objMain.objUtilities.AddAlphaField("@TNXBRANCH_C0", "Email", "Email", 100)
        objMain.objUtilities.AddAlphaField("@TNXBRANCH_C0", "Active", "Active", 10)
        objMain.objUtilities.AddAlphaField("@TNXBRANCH_C0", "Remarks", "Remarks", 254)
    End Sub
    Sub RENTPROMO()

        'Header Feields RENTPROMO
        objMain.objUtilities.CreateTable("RENTPROMO", "RENTPROMO Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@RENTPROMO", "AgrNo", "AgrNo", 50)
        'AgrNo

        'Row Fields RENTPROMO
        objMain.objUtilities.CreateTable("RENTPROMO_C0", "RENTPROMO Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@RENTPROMO_C0", "ItemCode", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@RENTPROMO_C0", "ItemName", "Item Name", 250)

    End Sub
    Sub POSB()

        'Header Feields POS Bill Stock Details
        objMain.objUtilities.CreateTable("POSBIL", "POSBILL Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddDateField("@POSBIL", "DocD", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@POSBIL", "Branc", "Branch", 100)
        objMain.objUtilities.AddLinkField("@POSBIL", "EXCL", "Excel", 80, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@POSBIL", "DOCE", "DocEntry", 100)
        objMain.objUtilities.AddAlphaField("@POSBIL", "AppID", "Application ID", 100)

        'Row Fields POS Bill Stock Details
        objMain.objUtilities.CreateTable("POSB_C0", "POSBILL Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@POSB_C0", "SlNo", "Sl Number", 100)
        objMain.objUtilities.AddAlphaField("@POSB_C0", "BarC", "Barcode", 100)
        objMain.objUtilities.AddAlphaField("@POSB_C0", "DisENG", "Discription English", 100)
        objMain.objUtilities.AddAlphaField("@POSB_C0", "DisABC", "Discription Arabic", 100)
        objMain.objUtilities.AddFloatField("@POSB_C0", "PQty", "Purchase Quantity", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@POSB_C0", "VAT", "VAT", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@POSB_C0", "LTAmo", "Line Total Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@POSB_C0", "UPRICE", "Unit Price", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@POSB_C0", "FRNM", "Foreign Name", 100)
    End Sub

    Sub PriceChangeRequest()
        'HeAder Table
        objMain.objUtilities.CreateTable("TNX_PCRT", "Price Change Request", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT", "TNXBRC", "Branch", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT", "TNXPUTP", "Price Update Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT", "TNXUSR", "User", 50)
        objMain.objUtilities.AddDateField("@TNX_PCRT", "TNXDDT", "Posting Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT", "TNXSTUS", "Status", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT", "Remark", "Remarks", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT", "Flyer", "Flyer Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT", "TPRT", "TPR Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT", "Check", "With Expiry", 1)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT", "TNXTYP", "Selling Price Type", 50)
        objMain.objUtilities.AddDateField("@TNX_PCRT", "FDate", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT", "FTime", "From Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddDateField("@TNX_PCRT", "FDTT", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PCRT", "TDTT", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PCRT", "FromD", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PCRT", "ToD", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddLinkField("@TNX_PCRT", "EXLNK", "Excel", 80, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT", "AppID", "AppID", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT", "PriceLNum", "Price List Num", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT", "Check", "Check", 1)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT", "PName", "Promotion Name", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT", "PDN", "Promotion Document Name", 50)
        objMain.objUtilities.AddFloatField("@TNX_PCRT", "TNMGP", "Normal gp", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_PCRT", "TFLGP", "Flyer gp", SAPbobsCOM.BoFldSubTypes.st_Percentage)

        objMain.objUtilities.AddDateField("@TNX_PCRT", "FTIMEF", "From Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddDateField("@TNX_PCRT", "TTIMEF", "To Time", SAPbobsCOM.BoFldSubTypes.st_Time)

        objMain.objUtilities.AddDateField("@TNX_PCRT", "FDATEM", "From Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddDateField("@TNX_PCRT", "TDATEM", "To Time", SAPbobsCOM.BoFldSubTypes.st_Time)

        objMain.objUtilities.AddDateField("@TNX_PCRT", "FTIMET", "From Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddDateField("@TNX_PCRT", "TTIMET", "To Time", SAPbobsCOM.BoFldSubTypes.st_Time)

        objMain.objUtilities.AddDateField("@TNX_PCRT", "FDATEM1", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PCRT", "TDATEMP", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddDateField("@TNX_PCRT", "FDATET", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PCRT", "TDATET", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddDateField("@TNX_PCRT", "FTIMEM", "From Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddDateField("@TNX_PCRT", "TTIMEM", "To Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT", "FLDS", "Flyer Designation", 250)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT", "EXCE", "Excel No", 50)
        objMain.objUtilities.addField("@TNX_PCRT", "FOUT", "Flyer Out", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")

        objMain.objUtilities.AddDateField("@TNX_PCRT", "FDATEB", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PCRT", "TDATEB", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)

        'CostFluctuation Tab
        objMain.objUtilities.CreateTable("TNX_PCRT_C0", "Price Change Request C0", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C0", "TNXICD", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C0", "TNXNM", "Item Name", 150)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C0", "TNXGRPO", "GRPO Number", 50)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C0", "ULPRC", "Updated Last Purchase Price", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C0", "PLPRC", "Previous Last Purchase Price", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C0", "SCOST", "Supplier Cost", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C0", "OMRG", "Old Margin", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C0", "NMRG", "New Margin", SAPbobsCOM.BoFldSubTypes.st_Percentage)

        objMain.objUtilities.addField("@TNX_PCRT_C0", "MUpdated", "Marked Updated", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")

        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C0", "UOM", "UOM", 50)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C0", "ORP", "Old Retail Price", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C0", "NRP", "New Retail Price", SAPbobsCOM.BoFldSubTypes.st_Sum)

        objMain.objUtilities.addField("@TNX_PCRT_C0", "PUpdate", "Price Updated", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C0", "PUS", "Price Updated Status", 254)

        objMain.objUtilities.AddInteger("@TNX_PCRT_C0", "GRPODE", "GRPO DocEntry", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddInteger("@TNX_PCRT_C0", "GRPOLID", "GRPO LineId", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C0", "FTYP", "Fluctuation Type", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C0", "OBJT", "Object Type", 100)
        objMain.objUtilities.addField("@TNX_PCRT_C0", "SPTP", "Special Price Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "Immediate,Scheduled Update", "Immediate,Scheduled Update", "")
        objMain.objUtilities.AddDateField("@TNX_PCRT_C0", "FDATE", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PCRT_C0", "FTIME", "From Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C0", "BCode", "Bar Code", 50)
        objMain.objUtilities.addField("@TNX_PCRT_C0", "PUR", "Price Update Required", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "Yes")
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C0", "RMK", "Remarks", 250)

        objMain.objUtilities.AddFloatField("@TNX_PCRT_C0", "ULPRCVAT", "Updated Last Purchase Price With Vat", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C0", "PLPRCVAT", "Previous Last Purchase Price With VAT", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C0", "ORPVAT", "Old Retail Price without VAT", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C0", "NRPVAT", "New Retail Price without Vat", SAPbobsCOM.BoFldSubTypes.st_Sum)

        'Flyer Price Tab
        objMain.objUtilities.CreateTable("TNX_PCRT_C1", "Price Change Request C1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C1", "BrCode", "Barcode", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C1", "ItemCode", "ItemCode", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C1", "ITNM", "ItemName", 250)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C1", "ITGP", "ItemGroup", 50)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C1", "CPrice", "Current Selling Price", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C1", "NwPrice", "New Selling Price", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C1", "Cost", "Cost", SAPbobsCOM.BoFldSubTypes.st_Sum)




        objMain.objUtilities.AddFloatField("@TNX_PCRT_C1", "NSNG", "Normal Selling", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C1", "FSNG", "Flyer Selling", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C1", "NMGP", "Normal gp", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C1", "FLGP", "Flyer gp", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C1", "EXPQ", "EXP Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C1", "ACost", "Actual Cost", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C1", "ACSL", "Actual Selling", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C1", "AFSL", "Actual FSelling", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C1", "NGPV", "N gp Value", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C1", "FGPV", "Flyer gp Value", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C1", "PRLS", "Profit Loss", SAPbobsCOM.BoFldSubTypes.st_Sum)





        objMain.objUtilities.AddFloatField("@TNX_PCRT_C1", "LPrice", "Last Purchase Price", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C1", "Marg", "Margin", SAPbobsCOM.BoFldSubTypes.st_Percentage)

        objMain.objUtilities.addField("@TNX_PCRT_C1", "MUpdated", "Marked Updated", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")

        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C1", "UOM", "UOM", 50)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C1", "ORP", "Old Retail Pricet", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C1", "NRP", "New Retail Price", SAPbobsCOM.BoFldSubTypes.st_Sum)

        objMain.objUtilities.addField("@TNX_PCRT_C1", "PUpdate", "Price Updated", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C1", "PUS", "Price Updated Status", 254)

        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C1", "PUS", "Price Updated Status", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C1", "PDN", "Promotion Document Number", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C1", "PS", "Promotion Status", 254)

        objMain.objUtilities.AddDateField("@TNX_PCRT_C1", "FDATEF", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PCRT_C1", "TDATEF", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C1", "SLOT", "Sell Out", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C1", "CRNT", "Credit Note Number", 254)


        'MOP Tab
        objMain.objUtilities.CreateTable("TNX_PCRT_C2", "Price Change Request C2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C2", "BCode", "BarCode", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C2", "ItemCode", "ItemCode", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C2", "IName", "Item Name", 150)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C2", "UOM", "UOM", 20)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C2", "CQty", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C2", "CCost", "Current Cost", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C2", "CSPrice", "Current Selling Price", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C2", "CMarg", "Current Margin", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C2", "NSPrice", "New Selling Price", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C2", "NMarg", "New Margin", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C2", "NCost", "New Cost", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C2", "CName", "Competitor Name", 150)
        objMain.objUtilities.AddLinkField("@TNX_PCRT_C2", "IAttach", "Invoice Attachment", 80, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C2", "Check", "Check", 1)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C2", "LPhPrice", "Last PurchasePrice", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C2", "LGRPPrice", "Last GRPO Price", SAPbobsCOM.BoFldSubTypes.st_Sum)

        objMain.objUtilities.addField("@TNX_PCRT_C2", "MUpdated", "Marked Updated", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")


        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C2", "UOM", "UOM", 50)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C2", "ORP", "Old Retail Pricet", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C2", "NRP", "New Retail Price", SAPbobsCOM.BoFldSubTypes.st_Sum)

        objMain.objUtilities.addField("@TNX_PCRT_C2", "PUpdate", "Price Updated", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C2", "PUS", "Price Updated Status", 254)

        objMain.objUtilities.addField("@TNX_PCRT_C2", "PriceR", "Price Reversed", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C2", "PRS", "Price Reverse Status", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C2", "PDN", "Promotion Document Number", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C2", "PS", "Promotion Status", 254)


        'TPR Tab
        objMain.objUtilities.CreateTable("TNX_PCRT_C3", "Price Change Request C3", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C3", "BrCode", "BarCode", 150)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C3", "ItemCode", "ItemCode", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C3", "ItemN", "Item Name", 150)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C3", "UOM", "UOM", 20)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C3", "Qty", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C3", "ICost", "Item Cost", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C3", "NSell", "Normal Selling", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C3", "SPrice", "Special Price", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C3", "DiffSP", "Diff (SP - Cost)", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C3", "DSPNS", "Diff (SP - Normal Sell)", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C3", "TPRC", "TPR Cost", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C3", "TPRS", "TPR Selling", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C3", "FundA", "Fund Allocation", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C3", "NGP", "Normal GP", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C3", "SGP", "Special GP", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C3", "Supp", "Fund From", 150)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C3", "Remark", "Remark", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C3", "UOM1", "UOM", 20)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C3", "SLOT", "Sell Out", SAPbobsCOM.BoFldSubTypes.st_Sum)

        objMain.objUtilities.addField("@TNX_PCRT_C3", "MUpdated", "Marked Updated", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")

        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C3", "UOM", "UOM", 50)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C3", "ORP", "Old Retail Pricet", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C3", "NRP", "New Retail Price", SAPbobsCOM.BoFldSubTypes.st_Sum)

        objMain.objUtilities.addField("@TNX_PCRT_C3", "PUpdate", "Price Updated", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")


        objMain.objUtilities.addField("@TNX_PCRT_C3", "PriceR", "Price Reversed", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C3", "PRS", "Price Reverse Status", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C3", "PUS", "Price Updated Status", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C3", "PDN", "Promotion Document Number", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C3", "PS", "Promotion Status", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C3", "CRNT", "Credit Note Number", 254)


        'Every day Special Price Tab
        objMain.objUtilities.CreateTable("TNX_PCRT_C5", "Price Change Request C5", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C5", "BCode", "BarCode", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C5", "ICode", "ItemCode", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C5", "IName", "Item Name", 150)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C5", "CQTT", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C5", "Unit", "Unit", 20)

        objMain.objUtilities.AddFloatField("@TNX_PCRT_C5", "CCost", "Current Cost", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C5", "CSlng", "Current Selling", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C5", "CMarg", "Current Margin", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C5", "NCost", "New Cost", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C5", "NSlng", "New Selling", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C5", "NMrgn", "New Margin", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.addField("@TNX_PCRT_C5", "PUpdate", "Price Updated", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C5", "PUS", "Price Updated Status", 254)
        objMain.objUtilities.addField("@TNX_PCRT_C5", "MUpdated", "Marked Updated", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
        objMain.objUtilities.AddLinkField("@TNX_PCRT_C5", "IAttach", "Invoice Attachment", 80, SAPbobsCOM.BoFldSubTypes.st_Link)
        'Bulk Cost Update
        objMain.objUtilities.CreateTable("TNX_PCRT_C6", "Price Change Request C6", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C6", "BCode", "BarCode", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C6", "ITCD", "ItemCode", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C6", "ITNM", "Item Name", 150)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C6", "UOM", "UOM", 20)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C6", "NPRC", "New Price", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C6", "Base", "Base Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C6", "Tax", "Tax Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.addField("@TNX_PCRT_C6", "PUpdate", "Price Updated", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C6", "PUS", "Price Updated Status", 254)

        'Cost Fluctuation Stock Transfer
        objMain.objUtilities.CreateTable("TNX_PCRT_C7", "Price Change Request C7", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C7", "BCode", "BarCode", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C7", "ITCD", "ItemCode", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C7", "ITNM", "Item Name", 150)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C7", "UOM", "UOM", 20)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C7", "QTY", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C7", "OMAC", "OLD MAC", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C7", "NMAC", "NEW MAC", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C7", "ODS", "Old Selling", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C7", "OMA", "Old Margin Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C7", "OMG", "Old Margin GP", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C7", "NMAR", "New Margin Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C7", "DIFF", "Difference", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C7", "NSP", "New Selling Price", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C7", "NMNS", "New Margin", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C7", "NMP", "New Margin percentage", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddInteger("@TNX_PCRT_C7", "GRPODE", "GRPO DocEntry", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddInteger("@TNX_PCRT_C7", "GRPOLID", "GRPO LineId", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C7", "OBJT", "Object Type", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C7", "TNXGRPO", "GRPO Number", 50)
        objMain.objUtilities.addField("@TNX_PCRT_C7", "PUpdate", "Price Updated", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C7", "PUS", "Price Updated Status", 254)


        'Bulk Price Update Tab
        objMain.objUtilities.CreateTable("TNX_PCRT_C8", "Price Change Request C8", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C8", "BCode", "BarCode", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C8", "BName", "BarCode Name", 200)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C8", "ITCD", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C8", "UOM", "UOM", 20)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C8", "OSLP", "Old Selling price", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PCRT_C8", "NSLP", "New Selling Price", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.addField("@TNX_PCRT_C8", "PUpdate", "Price Updated", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C8", "PUS", "Price Updated Status", 254)

        'Attachments Tab
        objMain.objUtilities.CreateTable("TNX_PCRT_C4", "Price Change Request C4", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddLinkField("@TNX_PCRT_C4", "TPath", "Target Path", 80, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C4", "FName", "File Name", 200)
        objMain.objUtilities.AddDateField("@TNX_PCRT_C4", "ADate", "Attachment Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_PCRT_C4", "FText", "Free Text", 254)

    End Sub
    Sub CreatePUIGA()
        objMain.objUtilities.CreateTable("SBO_PUIGA", "PUIGA Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@SBO_PUIGA", "USRCD", "User Code", 30)
        objMain.objUtilities.AddAlphaField("@SBO_PUIGA", "USRNM", "User Name", 254)
        objMain.objUtilities.AddAlphaField("@SBO_PUIGA", "SERIES", "Series", 30)

        objMain.objUtilities.CreateTable("SBO_PUIGA_C0", "PUIGA Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@SBO_PUIGA_C0", "ITMGRP", "Item Group", 100)
        objMain.objUtilities.AddAlphaField("@SBO_PUIGA_C0", "ACCESS", "Access", 1)
        objMain.objUtilities.AddAlphaField("@SBO_PUIGA_C0", "PRICEDIT", "Price Edit", 1)
    End Sub

    Sub CREATEUPDATEPRICESERVICE()
        objMain.objUtilities.CreateTable("SBO_PL", "PriceUpdate", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.CreateTable("SBO_PL_C0", "PriceUpdate child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.CreateTable("SBO_UPDATE", "PriceUpdate child", SAPbobsCOM.BoUTBTableType.bott_NoObject)


        '---------------------------------------------------------------
        objMain.objUtilities.AddAlphaField("@SBO_PL", "PLNo", "PriceList Number", 50)
        objMain.objUtilities.AddDateField("@SBO_PL", "DocDate", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@SBO_PL", "DocSts", "DocStatus", 50)
        objMain.objUtilities.AddLinkField("@SBO_PL", "upload", "Excel Upload", 250, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddLinkField("@SBO_PL", "Success", "Success Upload", 250, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddLinkField("@SBO_PL", "Failed", "Failed Upload", 250, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.addField("@SBO_PL", "PRICEUP", "PRICEUPDATE", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
        objMain.objUtilities.AddAlphaField("@SBO_PL", "Brnch", "Branch", 50)

        objMain.objUtilities.AddAlphaMemoField("@SBO_PL", "Serstat", "Service Status", 500)

        objMain.objUtilities.AddAlphaField("@SBO_PL_C0", "ICode", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@SBO_PL_C0", "IName", "Item Name", 100)
        objMain.objUtilities.AddAlphaField("@SBO_PL_C0", "BarCode", "Bar Code", 100)
        objMain.objUtilities.AddAlphaField("@SBO_PL_C0", "UOM", "UOM Code", 20)
        objMain.objUtilities.AddAlphaField("@SBO_PL_C0", "Remarks", "Remarks", 254)
        objMain.objUtilities.AddAlphaField("@SBO_PL_C0", "PLIST", "Price List", 254)
        objMain.objUtilities.AddAlphaField("@SBO_PL_C0", "Status", "Status", 254)
        objMain.objUtilities.AddAlphaField("@SBO_PL_C0", "IUOM", "Inventory Uom", 254)
        '
        objMain.objUtilities.AddAlphaField("@SBO_PL_C0", "PLNAME", "Price list Name", 254)
        objMain.objUtilities.AddAlphaField("@SBO_PL_C0", "UNAME", "Uom Name", 254)
        objMain.objUtilities.AddFloatField("@SBO_PL_C0", "CF", "CF", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@SBO_PL_C0", "Cost", "Cost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@SBO_PL_C0", "Margin", "Margin", SAPbobsCOM.BoFldSubTypes.st_Price)

        objMain.objUtilities.AddFloatField("@SBO_PL_C0", "RSP", "RSP", SAPbobsCOM.BoFldSubTypes.st_Price)

        objMain.objUtilities.AddFloatField("@SBO_PL_C0", "GP", "GP", SAPbobsCOM.BoFldSubTypes.st_Price)



        objMain.objUtilities.AddFloatField("@SBO_PL_C0", "NPrice", "New Price", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@SBO_PL_C0", "OPrice", "Old Price", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@SBO_PL_C0", "GPrice", "Gross Price", SAPbobsCOM.BoFldSubTypes.st_Price)


        objMain.objUtilities.AddDateField("@SBO_UPDATE", "SDate", "Inv Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@SBO_UPDATE", "EDate", "Inv Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaMemoField("@SBO_UPDATE", "Status", "Status", 1000)
        objMain.objUtilities.AddAlphaField("@SBO_UPDATE", "Flag", "Flag", 200)
        objMain.objUtilities.AddAlphaField("@SBO_UPDATE", "DOCNUM", "DOCNUM", 200)

        objMain.objUtilities.AddAlphaField("@SBO_PL", "SerStats", "Service Status", 250)
        '---------------------------------------------------------------
    End Sub
    Sub CreatePUBA()
        objMain.objUtilities.CreateTable("SBO_PUBA", "PUBA Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@SBO_PUBA", "USRCD", "User Code", 30)
        objMain.objUtilities.AddAlphaField("@SBO_PUBA", "USRNM", "User Name", 254)
        objMain.objUtilities.AddAlphaField("@SBO_PUBA", "SERIES", "Series", 30)

        objMain.objUtilities.CreateTable("SBO_PUBA_C0", "PUBA Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@SBO_PUBA_C0", "BPLID", "BPL Id", 30)
        objMain.objUtilities.AddAlphaField("@SBO_PUBA_C0", "BPLCD", "BPL Code", 50)
        objMain.objUtilities.AddAlphaField("@SBO_PUBA_C0", "BPLNM", "BPL Name", 100)
        objMain.objUtilities.AddAlphaField("@SBO_PUBA_C0", "ACCESS", "Access", 1)
        objMain.objUtilities.AddAlphaField("@SBO_PUBA_C0", "PRICEDIT", "Price Edit", 1)
    End Sub



    Sub CreatePaymentDetailsScreen()
        objMain.objUtilities.CreateTable("ME_PDS", "Payment Details Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@ME_PDS", "MESPCD", "SupplierCode", 30)
        objMain.objUtilities.AddAlphaField("@ME_PDS", "MESPNM", "SupplierName", 254)
        objMain.objUtilities.AddAlphaField("@ME_PDS", "SERIES", "Series", 30)
        objMain.objUtilities.AddAlphaField("@ME_PDS", "STATUS", "DocStatus", 30)
        objMain.objUtilities.AddAlphaField("@ME_PDS", "BASENT", "BaseEntry", 30)
        objMain.objUtilities.AddAlphaField("@ME_PDS", "BASENO", "BaseDocNo", 30)
        objMain.objUtilities.AddAlphaField("@ME_PDS", "MERMKS", "Remarks", 254)
        objMain.objUtilities.AddAlphaField("@ME_PDS", "BRANCH", "Branch", 50)
        objMain.objUtilities.AddDateField("@ME_PDS", "STARTDT", "StartDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@ME_PDS", "ENDDT", "EndDate", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.CreateTable("ME_PDS_C0", "Payment Details Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddDateField("@ME_PDS_C0", "MEDATE", "Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@ME_PDS_C0", "MEACCTCD", "AccountCode", 30)
        objMain.objUtilities.AddAlphaField("@ME_PDS_C0", "MEACCTNM", "AccountName", 254)
        objMain.objUtilities.AddFloatField("@ME_PDS_C0", "MEAMT", "Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
        'objMain.objUtilities.AddAlphaField("@ME_PDS_C0", "MEJENO", "JENo", 30)
        objMain.objUtilities.AddAlphaField("@ME_PDS_C0", "APCRMEMO", "APCreditMemo", 30)
        objMain.objUtilities.AddAlphaField("@ME_PDS_C0", "MESTS", "Status", 254)
        objMain.objUtilities.AddAlphaField("@ME_PDS_C0", "MERMKS", "Remarks", 254)
        objMain.objUtilities.AddAlphaField("@ME_PDS_C0", "REF", "Referrence", 254)
    End Sub

    Sub CreatePDTStockCount()
        objMain.objUtilities.CreateTable("ME_PSC", "PDTStockCount Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@ME_PSC", "TYPE", "Type", 30)
        objMain.objUtilities.AddAlphaField("@ME_PSC", "ITMGRP", "ItemGroup", 30)
        objMain.objUtilities.AddAlphaField("@ME_PSC", "SUBGRP", "SubGroup", 30)
        objMain.objUtilities.AddAlphaField("@ME_PSC", "SBSBGRP", "SubSubGroup", 30)
        objMain.objUtilities.AddAlphaField("@ME_PSC", "PDTNo", "PDTNo", 30)
        objMain.objUtilities.AddDateField("@ME_PSC", "CountDate", "CountDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@ME_PSC", "RefNo", "RefNo", 30)
        objMain.objUtilities.AddAlphaField("@ME_PSC", "Status", "Status", 30)
        objMain.objUtilities.AddAlphaField("@ME_PSC", "Series", "Series", 30)
        objMain.objUtilities.AddAlphaField("@ME_PSC", "PDTUser", "PDTUser", 100)
        objMain.objUtilities.AddDateField("@ME_PSC", "CountTime", "CountTime", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddAlphaField("@ME_PSC", "RackNo", "RackNo", 30)
        objMain.objUtilities.AddAlphaField("@ME_PSC", "WhseCode", "WhseCode", 30)
        objMain.objUtilities.AddAlphaField("@ME_PSC", "Branch", "Branch", 30)

        objMain.objUtilities.CreateTable("ME_PSC_C0", "PDTStockCount Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@ME_PSC_C0", "Barcode", "Barcode", 50)
        objMain.objUtilities.AddAlphaField("@ME_PSC_C0", "ItemCode", "ItemCode", 50)
        objMain.objUtilities.AddAlphaField("@ME_PSC_C0", "ItemName", "ItemName", 254)
        objMain.objUtilities.AddAlphaField("@ME_PSC_C0", "UOMCode", "UomCode", 30)
        objMain.objUtilities.AddAlphaField("@ME_PSC_C0", "UomName", "UomName", 30)
        objMain.objUtilities.AddAlphaField("@ME_PSC_C0", "BaseUom", "BaseUom", 30)
        objMain.objUtilities.AddAlphaField("@ME_PSC_C0", "ItmPrUnt", "ItemPerUnit", 30)
        objMain.objUtilities.AddFloatField("@ME_PSC_C0", "Qty", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddInteger("@ME_PSC_C0", "LineNo", "Line Number", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@ME_PSC_C0", "Status", "LineStatus", 30)
        objMain.objUtilities.AddAlphaField("@ME_PSC_C0", "Baynum", "BayNumber", 100)
        objMain.objUtilities.AddAlphaField("@ME_PSC_C0", "POST", "Posting Status", 100)
        objMain.objUtilities.AddAlphaField("@ME_PSC_C0", "Remarks", "Remarks", 254)

    End Sub

    Sub CreatePDCPostingTable()
        objMain.objUtilities.CreateTable("ME_PDC", "PDC header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddDateField("@ME_PDC", "PostDate", "Posting Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@ME_PDC", "DateFrom", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@ME_PDC", "DateTo", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@ME_PDC", "SERIES", "DocSeries", 30)
        objMain.objUtilities.AddAlphaField("@ME_PDC", "BRANCH", "Branch", 100)
        objMain.objUtilities.AddAlphaField("@ME_PDC", "Remarks", "Remarks", 254)
        objMain.objUtilities.AddAlphaField("@ME_PDC", "SUPCODE", "SupplierCode", 150)
        objMain.objUtilities.addField("@ME_PDC", "Issued", "Issued", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N,A", "Yes,No,All", "")
        objMain.objUtilities.addField("@ME_PDC", "Approved", "Approved", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N,A", "Yes,No,All", "")
        objMain.objUtilities.AddAlphaField("@ME_PDC", "ChequeNo", "CheckNum", 30)
        objMain.objUtilities.AddAlphaField("@ME_PDC", "GLAcc", "GL Account", 30)
        objMain.objUtilities.AddFloatField("@ME_PDC", "ChequeSm", "ChequeSuma", SAPbobsCOM.BoFldSubTypes.st_Price)

        objMain.objUtilities.CreateTable("ME_PDC_C0", "PDC child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@ME_PDC_C0", "DocEntry", "DocEntry", 30)
        objMain.objUtilities.AddDateField("@ME_PDC_C0", "DocDate", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@ME_PDC_C0", "CardCode", "CardCode", 30)
        objMain.objUtilities.AddAlphaField("@ME_PDC_C0", "CardName", "CardName", 100)
        objMain.objUtilities.AddAlphaField("@ME_PDC_C0", "ChequeNo", "CheckNum", 30)
        objMain.objUtilities.AddDateField("@ME_PDC_C0", "DueDate", "CheckNum", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddFloatField("@ME_PDC_C0", "ChequeSm", "ChequeSuma", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@ME_PDC_C0", "AcctNum", "Account Number", 30)
        objMain.objUtilities.AddAlphaField("@ME_PDC_C0", "AcctName", "Account Name", 100)
        objMain.objUtilities.AddAlphaField("@ME_PDC_C0", "DepositA", "Diposit Account", 30)
        objMain.objUtilities.AddAlphaField("@ME_PDC_C0", "JeNo", "JE No", 30)
        objMain.objUtilities.AddAlphaField("@ME_PDC_C0", "Required", "Required", 30)
        objMain.objUtilities.addField("@ME_PDC_C0", "Issued", "Issued", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "N")
        objMain.objUtilities.addField("@ME_PDC_C0", "Approved", "Approved", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "N")
        objMain.objUtilities.AddAlphaField("@ME_PDC_C0", "RecvNm", "ReceiverName", 150)
        objMain.objUtilities.AddDateField("@ME_PDC_C0", "PostDate", "Posting Date", SAPbobsCOM.BoFldSubTypes.st_None)

    End Sub

    Sub CreateItemList()
        objMain.objUtilities.CreateTable("SBO_LIST", "Item List header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@SBO_LIST", "CardCode", "Card Code", 100)
        objMain.objUtilities.AddAlphaField("@SBO_LIST", "CardName", "Card Name", 254)
        objMain.objUtilities.AddAlphaField("@SBO_LIST", "BRANCH", "Branch", 100)
        objMain.objUtilities.AddAlphaField("@SBO_LIST", "SERIES", "Series", 50)
        objMain.objUtilities.AddAlphaField("@SBO_LIST", "Statu", "Status", 50)
        objMain.objUtilities.AddAlphaField("@SBO_LIST", "AppID", "Application ID", 50)

        objMain.objUtilities.CreateTable("SBO_LIST_C0", "Item List child 1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddDateField("@SBO_LIST_C0", "Date", "Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@SBO_LIST_C0", "SKU", "No of SKU", 50)
        objMain.objUtilities.AddFloatField("@SBO_LIST_C0", "Amount", "Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddAlphaField("@SBO_LIST_C0", "NOU", "No of Outlets", 30)
        objMain.objUtilities.AddAlphaField("@SBO_LIST_C0", "JEPosted", "JE Posted", 30)
        objMain.objUtilities.AddAlphaField("@SBO_LIST_C0", "Remarks", "Remarks", 254)
        objMain.objUtilities.AddAlphaField("@SBO_LIST_C0", "BRANCH", "Branch", 50)
        objMain.objUtilities.AddAlphaField("@SBO_LIST_C0", "Charge", "No Of Charge", 50)
        objMain.objUtilities.AddAlphaField("@SBO_LIST_C0", "VATCODE", "Vat Code", 50)
        objMain.objUtilities.AddFloatField("@SBO_LIST_C0", "VATAMT", "Vat Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@SBO_LIST_C0", "TAMNT", "Total Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)

        objMain.objUtilities.CreateTable("SBO_LIST_C1", "Item List child 2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        'objMain.objUtilities.AddAlphaField("@SBO_LIST_C1", "Outlet", "Outlet", 1)
        'objMain.objUtilities.AddAlphaField("@SBO_LIST_C1", "Outlet2", "Outlet2", 1)
        'objMain.objUtilities.AddAlphaField("@SBO_LIST_C1", "Outlet3", "Outlet3", 1)
        'objMain.objUtilities.AddAlphaField("@SBO_LIST_C1", "Outlet4", "Outlet4", 1)
        'objMain.objUtilities.AddAlphaField("@SBO_LIST_C1", "Outlet5", "Outlet5", 1)
        objMain.objUtilities.AddAlphaField("@SBO_LIST_C1", "Brach", "Brach", 100)
        objMain.objUtilities.AddDateField("@SBO_LIST_C1", "FDate", "FDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@SBO_LIST_C1", "TDate", "TDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@SBO_LIST_C1", "Active", "Active", 1)


        objMain.objUtilities.CreateTable("SBO_LIST_C2", "Item List child 3", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@SBO_LIST_C2", "DESC", "Description", 100)
        objMain.objUtilities.AddAlphaField("@SBO_LIST_C2", "JESTATUS", "JE Status", 250)
        objMain.objUtilities.AddFloatField("@SBO_LIST_C2", "AMT", "Front End Percent", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@SBO_LIST_C2", "PTG", "Percent", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddAlphaField("@SBO_LIST_C2", "JE", "JE Entry", 30)
        objMain.objUtilities.AddAlphaField("@SBO_LIST_C2", "REMARKS", "Remarks", 254)
        objMain.objUtilities.AddFloatField("@SBO_LIST_C2", "VAT", "VAT %", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddAlphaField("@SBO_LIST_C2", "VATACC", "VAT Account", 50)
        objMain.objUtilities.AddAlphaField("@SBO_LIST_C2", "BRANCH", "Branch", 50)
        objMain.objUtilities.AddAlphaField("@SBO_LIST_C2", "VATCODE", "Vat Code", 50)
        objMain.objUtilities.AddFloatField("@SBO_LIST_C2", "VATAMT", "Vat Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)

    End Sub

    Sub CreateSubItemList()
        objMain.objUtilities.CreateTable("SBO_LS", "Items", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@SBO_LS", "CardCode", "Card Code", 100)
        objMain.objUtilities.AddAlphaField("@SBO_LS", "CardName", "Card Name", 254)
        objMain.objUtilities.AddAlphaField("@SBO_LS", "LineNum", "LineNum", 30)
        objMain.objUtilities.AddAlphaField("@SBO_LS", "SERIES", "Series", 50)
        objMain.objUtilities.AddAlphaField("@SBO_LS", "SKUS", "No of SKUS", 50)
        objMain.objUtilities.AddAlphaField("@SBO_LS", "DocNumIL", "Document No Item Listing", 50)
        objMain.objUtilities.AddAlphaField("@SBO_LS", "AppIDILG", "ApplicationID Item Listing", 50)

        objMain.objUtilities.CreateTable("SBO_LS_C0", "Items C0", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@SBO_LS_C0", "ICode", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@SBO_LS_C0", "IName", "Item Name", 254)

        objMain.objUtilities.AddAlphaField("@SBO_LS_C0", "BarCode", "Bar Code", 50)
        objMain.objUtilities.AddAlphaField("@SBO_LS_C0", "IGroup", "Item Group", 254)
        objMain.objUtilities.AddAlphaField("@SBO_LS_C0", "UOM", "UOM", 50)

    End Sub
    Sub CreateOutlet()
        objMain.objUtilities.CreateTable("SBO_OUTLET", "Outlet", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@SBO_OUTLET", "CardCode", "Card Code", 100)
        objMain.objUtilities.AddAlphaField("@SBO_OUTLET", "CardName", "Card Name", 254)
        objMain.objUtilities.AddAlphaField("@SBO_OUTLET", "LineNum", "LineNum", 30)
        objMain.objUtilities.AddAlphaField("@SBO_OUTLET", "BaseNo", "BaseNo", 30)
        objMain.objUtilities.AddAlphaField("@SBO_OUTLET", "SERIES", "Series", 50)
        objMain.objUtilities.AddAlphaField("@SBO_OUTLET", "PType", "Postingtype", 50)
        objMain.objUtilities.AddFloatField("@SBO_OUTLET", "AUT", "Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddAlphaField("@SBO_OUTLET", "Outlet", "NoOfOutlet", 50)
        objMain.objUtilities.AddAlphaField("@SBO_OUTLET", "VATCODE", "Vat Code", 50)
        objMain.objUtilities.AddDateField("@SBO_OUTLET", "DocDate", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)


        objMain.objUtilities.CreateTable("SBO_Outlet_C0", "Outlet C0", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@SBO_OUTLET_C0", "Outlet", "Outlet", 50)
        objMain.objUtilities.AddAlphaField("@SBO_OUTLET_C0", "Amount", "Amount", 1)
        objMain.objUtilities.AddAlphaField("@SBO_OUTLET_C0", "JE", "JE", 50)
        objMain.objUtilities.AddAlphaField("@SBO_OUTLET_C0", "Payment", "Payment", 50)
        objMain.objUtilities.AddFloatField("@SBO_OUTLET_C0", "AmntSpt", "Amount Splitting", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddAlphaField("@SBO_OUTLET_C0", "VATCODE", "Vat Code", 50)
        objMain.objUtilities.AddFloatField("@SBO_OUTLET_C0", "VATAMT", "Vat Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)


    End Sub

    Sub CreateAutomaticProduction()
        objMain.objUtilities.CreateTable("SBO_PR", "Production Head", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddDateField("@SBO_PR", "DocDate", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@SBO_PR", "PostDate", "PostDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@SBO_PR", "SERIES", "DocSeries", 30)
        objMain.objUtilities.AddAlphaField("@SBO_PR", "BRANCH", "Branch", 30)
        objMain.objUtilities.AddAlphaField("@SBO_PR", "WHSCODE", "WhsCode", 100)
        objMain.objUtilities.AddAlphaMemoField("@SBO_PR", "HRemrk", "HRemrk", 5000)

        objMain.objUtilities.CreateTable("SBO_PR_C0", "Production Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@SBO_PR_C0", "ICode", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@SBO_PR_C0", "IName", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@SBO_PR_C0", "BarCode", "Bar Code", 100)
        objMain.objUtilities.AddAlphaField("@SBO_PR_C0", "UOM", "UOM Code", 20)
        objMain.objUtilities.AddFloatField("@SBO_PR_C0", "OPrice", "Retails Price", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@SBO_PR_C0", "Prod", "Production Entry", 20)
        objMain.objUtilities.AddAlphaField("@SBO_PR_C0", "IProd", "Issue Entry.", 20)
        objMain.objUtilities.AddAlphaField("@SBO_PR_C0", "RProd", "Receipt Entry.", 20)
        objMain.objUtilities.AddAlphaField("@SBO_PR_C0", "ProdNo", "Production No", 20)
        objMain.objUtilities.AddAlphaField("@SBO_PR_C0", "IProdNo", "Issue No.", 20)
        objMain.objUtilities.AddAlphaField("@SBO_PR_C0", "RProdNo", "Receipt No.", 20)
        objMain.objUtilities.AddAlphaField("@SBO_PR_C0", "Status", "Status", 254)
        objMain.objUtilities.AddFloatField("@SBO_PR_C0", "Qty", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        'objMain.objUtilities.AddFloatField("@SBO_PR_C0", "Qnty", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@SBO_PR_C0", "RevQnty", "Received Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        'objMain.objUtilities.AddFloatField("@SBO_PR_C0", "RejQnty", "Rejected Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@SBO_PR_C0", "RejQty", "Rejected Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        'objMain.objUtilities.AddAlphaField("@SBO_PR_C0", "Remark", "Remark", 254)
        objMain.objUtilities.AddAlphaMemoField("@SBO_PR_C0", "Remak", "Remark", 5000)


    End Sub

    Sub CreatAddonConfigScrn()
        objMain.objUtilities.CreateTable("ME_ACFG", "Addon Configuration", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "SAGCRACT", "SupAgrCreditAcct", 50)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "METWHSCD", "TransitWarehouse", 50)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "METWHSNM", "TransitWarehouseName", 150)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "MEDWHSCD", "DamageWarehouse", 50)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "MEDWHSNM", "DamageWarehouseName", 150)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "MEEXACCT", "ExitAccount", 50)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "MEENTACT", "EntryAccount", 50)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "KITMIACT", "KitItemsIssueAccount", 50)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "KITMRACT", "KitItemsReceiptAccount", 50)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "REG", "Registration Charges GL", 50)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "MEPSWD", "HanaPassword", 30)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "MEATCH", "AttchPath", 254)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "MEPATH", "CrystalRptPath", 254)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "STGBLOC", "Staging Bin Location", 100)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "RNTLACT", "RentalAccount", 50)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "MEUSRNM", "SAP UserName", 50)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "MEPWD", "SAP Password", 50)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "APCMSRS", "APCreditMemoSeries", 50)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "PDTTXCD", "PDTTaxCode", 50)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "RChar", "Supplier charges", 50)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "Least ", "Least Price", 1)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "PO ", "PO", 1)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "GRPO ", "GRPO", 1)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "Round", "Round", 50)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "SuppReb", "Supplier Rebate", 40)
        objMain.objUtilities.addField("@ME_ACFG", "INTRW", "In-Transit Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 5, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "N")
        objMain.objUtilities.addField("@ME_ACFG", "Auto", "Automatic Tax", SAPbobsCOM.BoFieldTypes.db_Alpha, 5, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "N")
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "VatCode", "VatCode", 50)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "Excise ", "Excise Duty add-on", 1)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "FOCCal ", "FOC calculation", 1)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "Supreb", "Supplier rebats", 30)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "Barg", "Barcode generatin", 30)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "APCM", "Ap Credit Memo", 30)
        objMain.objUtilities.AddAlphaField("@ME_ACFG", "PriceLst", "Price List", 30)


    End Sub

    Sub CreateStockTransferRequest()
        objMain.objUtilities.CreateTable("ME_STR", "StockTransferRequest Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@ME_STR", "TYPE", "Type", 30)
        objMain.objUtilities.AddAlphaField("@ME_STR", "MEPRN", "PurchaseReqNo", 30)
        objMain.objUtilities.AddAlphaField("@ME_STR", "MEPRE", "PurchaseReqEntry", 30)
        objMain.objUtilities.AddAlphaField("@ME_STR", "MEREMPID", "EmpId", 30)
        objMain.objUtilities.AddAlphaField("@ME_STR", "MEREQBY", "RequestedBy", 254)
        objMain.objUtilities.AddAlphaField("@ME_STR", "MEFWHSCD", "FromWarehouseCode", 50)
        objMain.objUtilities.AddAlphaField("@ME_STR", "MEFWHSNM", "WarehouseName", 100)
        objMain.objUtilities.AddAlphaField("@ME_STR", "MEFBPID", "FromBranch", 30)
        objMain.objUtilities.AddAlphaField("@ME_STR", "METWHSCD", "ToWarehouseCode", 50)
        objMain.objUtilities.AddAlphaField("@ME_STR", "METWHSNM", "WarehouseName", 100)
        objMain.objUtilities.AddAlphaField("@ME_STR", "METBPID", "ToBranch", 30)
        objMain.objUtilities.AddAlphaField("@ME_STR", "MESRS", "DocSeries", 30)
        objMain.objUtilities.AddAlphaField("@ME_STR", "MESTS", "DocStatus", 30)
        objMain.objUtilities.AddDateField("@ME_STR", "MEPSDT", "PostingDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@ME_STR", "MEDOCDT", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@ME_STR", "MERMKS", "Remarks", 254)
        objMain.objUtilities.AddAlphaField("@ME_STR", "PDTUSER", "Pdt User", 20)
        objMain.objUtilities.AddAlphaField("@ME_STR", "UserName", "UserName", 100)
        objMain.objUtilities.AddAlphaField("@ME_STR", "REMARKS", "REMARKS", 254)
        objMain.objUtilities.AddAlphaField("@ME_STR", "GINum", "GINum", 20)
        objMain.objUtilities.AddAlphaField("@ME_STR", "GIDoc", "GIDoc", 20)

        objMain.objUtilities.CreateTable("ME_STR_C0", "StockTransferRequest Child1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@ME_STR_C0", "MEBRCD", "Barcode", 100)
        objMain.objUtilities.AddAlphaField("@ME_STR_C0", "MEBRNM", "BarName", 254)
        objMain.objUtilities.AddAlphaField("@ME_STR_C0", "MEITMCD", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@ME_STR_C0", "MEITNM", "Item Name", 254)
        objMain.objUtilities.AddAlphaField("@ME_STR_C0", "MEUOM", "UOM", 100)
        objMain.objUtilities.AddAlphaField("@ME_STR_C0", "MEADBRCD", "Addl Barcode", 30)
        objMain.objUtilities.AddAlphaField("@ME_STR_C0", "MEFWSTK", "From WH SOH", 30)
        objMain.objUtilities.AddAlphaField("@ME_STR_C0", "METWHSTK", "To WH SOH", 30)
        objMain.objUtilities.AddAlphaField("@ME_STR_C0", "MEFWHSCD", "FromWarehouse", 50)
        objMain.objUtilities.AddAlphaField("@ME_STR_C0", "METWHSCD", "ToWarehouse", 50)
        objMain.objUtilities.AddFloatField("@ME_STR_C0", "MEREQQTY", "Required Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@ME_STR_C0", "MERELQTY", "Dispatched Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@ME_STR_C0", "MEOPNQTY", "Open Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@ME_STR_C0", "MERCEQTY", "Received Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@ME_STR_C0", "MELNSTS", "LineStatus", 30)
        objMain.objUtilities.AddAlphaField("@ME_STR_C0", "MEPCKNG", "Packing", 50)

        objMain.objUtilities.CreateTable("ME_STR_C1", "StockTransferRequest Child2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@ME_STR_C1", "MEATCNM", "Attachment Name", 100)
        objMain.objUtilities.AddLinkField("@ME_STR_C1", "MEATCH", "Attachment", 254, SAPbobsCOM.BoFldSubTypes.st_Link)
    End Sub
    Sub SupDeposit()
        objMain.objUtilities.CreateTable("TNX_SDPT", "Supplier Depst", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@TNX_SDPT", "GLAcct", "G/L Account", 50)
        objMain.objUtilities.AddAlphaField("@TNX_SDPT", "SupChrg", "Supplier Charges", 1)
        objMain.objUtilities.AddAlphaField("@TNX_SDPT", "Itmlst", "Item Listing", 1)
        objMain.objUtilities.AddAlphaField("@TNX_SDPT", "SupCtr", "Supplier Contract", 1)
        objMain.objUtilities.AddAlphaField("@TNX_SDPT", "VATACC", "VAT Account", 50)


    End Sub
    Sub PettyCashMaster()
        objMain.objUtilities.CreateTable("TNX_PCTM", "Petty Cash type Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@TNX_PCTM", "Type", "Type", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PCTM", "ACCT", "G/L Account", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCTM", "Name", "Account Name", 250)
        objMain.objUtilities.addField("@TNX_PCTM", "ACT", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "Yes")

    End Sub


    Sub LicenceNew()
        objMain.objUtilities.CreateTable("TNX_LICENCE", "Licence Administration New one", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "Company", "Company Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "DocNum", "Document Number", 100)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "License", "License", 254)


        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "DB", "Database Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "HKey", "Hardware Key", 100)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "Total", "Total Licenses", 10)
        objMain.objUtilities.AddDateField("@TNX_LICENCE", "SDate", "Start Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_LICENCE", "EDate", "Expiry Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "Addon", "Add-on Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "NOUSR", "No of Users", 254)

        objMain.objUtilities.CreateTable("TNX_LICENCE_C0", "License Administration Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE_C0", "Code", "Device Code", 254)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE_C0", "USERC", "User Code", 254)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE_C0", "Name", "User Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE_C0", "Sts", "Status", 1)

        'objMain.objUtilities.CreateTable("SBO_ADDON", "License Addon", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        'objMain.objUtilities.AddAlphaField("@SBO_ADDON", "ADCODE", "Addon Code", 254)
        'objMain.objUtilities.AddAlphaField("@SBO_ADDON", "ADNAM", "Addon Name", 254)
        'objMain.objUtilities.AddAlphaField("@SBO_ADDON", "USER", "Used", 1)
        'objMain.objUtilities.AddAlphaField("@SBO_ADDON", "AVBLE", "Available", 254)

    End Sub

    Sub CreateMaterialInward()
        objMain.objUtilities.CreateTable("ME_MI", "MaterialInward Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@ME_MI", "METYPE", "Type", 30)
        objMain.objUtilities.AddAlphaField("@ME_MI", "MEREMPID", "EmpId", 30)
        objMain.objUtilities.AddAlphaField("@ME_MI", "MERECBY", "ReceivedBy", 254)
        objMain.objUtilities.AddAlphaField("@ME_MI", "MESTRNO", "Stock Request No", 50)
        objMain.objUtilities.AddAlphaField("@ME_MI", "MESTREN", "STR Entry", 30)
        objMain.objUtilities.AddAlphaField("@ME_MI", "MESTS", "DocStatus", 30)
        objMain.objUtilities.AddAlphaField("@ME_MI", "MEITRNO", "ITR No", 50)
        objMain.objUtilities.AddAlphaField("@ME_MI", "MEITREN", "ITR Entry", 30)
        objMain.objUtilities.AddAlphaField("@ME_MI", "MEFBPID", "FromBranch", 30)
        objMain.objUtilities.AddAlphaField("@ME_MI", "MEFWHS", "From WH", 50)
        objMain.objUtilities.AddAlphaField("@ME_MI", "MEFWNM", "From WH Name", 254)
        objMain.objUtilities.AddAlphaField("@ME_MI", "METBPID", "ToBranch", 30)
        objMain.objUtilities.AddAlphaField("@ME_MI", "METWHS", "To WH", 50)
        objMain.objUtilities.AddAlphaField("@ME_MI", "METWHNM", "To WH Name", 254)
        objMain.objUtilities.AddAlphaField("@ME_MI", "METRWHS", "To Transit WH", 50)
        objMain.objUtilities.AddAlphaField("@ME_MI", "METRWNM", "To Transit WH Name", 254)
        objMain.objUtilities.AddAlphaField("@ME_MI", "MESRS", "DocSeries", 30)
        objMain.objUtilities.AddAlphaField("@ME_MI", "MEGINO", "GoodsIssueNo", 30)
        objMain.objUtilities.AddAlphaField("@ME_MI", "MEGRNO", "GoodsReceiptNo", 30)
        objMain.objUtilities.AddAlphaField("@ME_MI", "MEGIEN", "GoodsIssueEntry", 30)
        objMain.objUtilities.AddAlphaField("@ME_MI", "MEGREN", "GoodsReceiptEntry", 30)
        objMain.objUtilities.AddDateField("@ME_MI", "MEDATE", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@ME_MI", "PDTUSER", "Pdt User", 20)
        objMain.objUtilities.AddAlphaField("@ME_MI", "UserName", "UserName", 100)
        objMain.objUtilities.AddAlphaField("@ME_MI", "VehicleNo", "VehicleNo", 100)
        objMain.objUtilities.AddAlphaField("@ME_MI", "REMARKS", "REMARKS", 254)
        objMain.objUtilities.addField("@ME_MI", "Updated", "Updated", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Y,N", "N")

        objMain.objUtilities.CreateTable("ME_MI_C0", "MaterialInward Child1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@ME_MI_C0", "MEBRCD", "Barcode", 50)
        objMain.objUtilities.AddAlphaField("@ME_MI_C0", "MEITMCD", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@ME_MI_C0", "MEITDSC", "Item Description", 254)
        objMain.objUtilities.AddAlphaField("@ME_MI_C0", "MEUOM", "UOM Code", 50)
        objMain.objUtilities.AddAlphaField("@ME_MI_C0", "ADBRCD", "Add Barcode", 50)
        objMain.objUtilities.AddAlphaField("@ME_MI_C0", "MEPKGE", "Package", 50)

        objMain.objUtilities.AddAlphaField("@ME_MI_C0", "METYPE", "Type", 50)
        objMain.objUtilities.AddAlphaMemoField("@ME_MI_C0", "MESRBT", "SerialOrBartch", 64000)
        objMain.objUtilities.AddAlphaField("@ME_MI_C0", "CHKBIN", "CheckBin", 50)
        objMain.objUtilities.AddAlphaField("@ME_MI_C0", "METWHS", "ToWarhouse", 50)
        objMain.objUtilities.AddAlphaField("@ME_MI_C0", "METWBL", "WarehouseBinLocation", 50)
        objMain.objUtilities.AddAlphaField("@ME_MI_C0", "METBRCH", "ToBranch", 50)

        objMain.objUtilities.AddFloatField("@ME_MI_C0", "MEVARQTY", "Variance Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@ME_MI_C0", "MEREQQTY", "Dispatched Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@ME_MI_C0", "MERECQTY", "Received Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@ME_MI_C0", "MESTRNO", "STRNo", 30)
        objMain.objUtilities.AddAlphaField("@ME_MI_C0", "MESTREN", "STREntry", 30)
        objMain.objUtilities.AddInteger("@ME_MI_C0", "MESTRLN", "STR Line", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@ME_MI_C0", "MEITRNO", "ITRNo", 30)
        objMain.objUtilities.AddAlphaField("@ME_MI_C0", "MEITREN", "ITREntry", 30)
        objMain.objUtilities.AddAlphaField("@ME_MI_C0", "MEBSTYP", "BaseType", 30)
        objMain.objUtilities.AddInteger("@ME_MI_C0", "MEBLN", "BaseLine", SAPbobsCOM.BoFldSubTypes.st_None, 10)
    End Sub

    Sub CreatePMAPINVTable()
        objMain.objUtilities.CreateTable("PMAPINV", "PostMultipleAPInvoiceTemp", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement)
        objMain.objUtilities.AddInteger("@PMAPINV", "GRIDROW", "GridRow", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddInteger("@PMAPINV", "GRPOENT", "GRPOEntry", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddInteger("@PMAPINV", "GRPOROW", "GrpoRowId", SAPbobsCOM.BoFldSubTypes.st_None, 10)
    End Sub

    Sub SupRebateWizard()
        objMain.objUtilities.CreateTable("TNX_SUPP", "Supplier Rebate", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@TNX_SUPP", "BRNH", "Branch", 50)
        objMain.objUtilities.AddAlphaField("@TNX_SUPP", "PRD", "Period", 30)
        objMain.objUtilities.AddAlphaField("@TNX_SUPP", "Select", "Select", 1)
        objMain.objUtilities.AddAlphaField("@TNX_SUPP", "Sel", "Select All", 1)
        'objMain.objUtilities.AddAlphaField("@TNX_SUPP", "TOTAMT", "Total Amount", 5)
        objMain.objUtilities.AddFloatField("@TNX_SUPP", "TlAmt", "Total Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddDateField("@TNX_SUPP", "FDATE", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_SUPP", "TDATE", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_SUPP", "Crtcode", "CreationCode", 100)
        objMain.objUtilities.AddAlphaField("@TNX_SUPP", "Brand", "Brand", 50)
        objMain.objUtilities.AddAlphaField("@TNX_SUPP", "FML", "FML", 50)
        objMain.objUtilities.AddAlphaField("@TNX_SUPP", "SFML", "SFML", 50)
        objMain.objUtilities.AddDateField("@TNX_SUPP", "PDATE", "Production Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_SUPP", "Supcd", "Supplier Code", 250)

    End Sub
    Sub CreatePMAPCMTable()
        objMain.objUtilities.CreateTable("PMAPCMT", "PostMultipleAPInvoiceTemp", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement)
        objMain.objUtilities.AddInteger("@PMAPCMT", "GRIDROW", "GridRow", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddInteger("@PMAPCMT", "GRTENT", "GoodsReturynEntry", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddInteger("@PMAPCMT", "GRTROW", "GoddsReturnRowId", SAPbobsCOM.BoFldSubTypes.st_None, 10)
    End Sub

    Sub CreateSupplierAgreementTable()
        objMain.objUtilities.CreateTable("ME_SUPAG", "Supplier Agreement Hdr", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@ME_SUPAG", "SUPCD", "Supplier Code", 50)
        objMain.objUtilities.AddAlphaField("@ME_SUPAG", "SUPNM", "Supplier Name", 254)
        objMain.objUtilities.AddDateField("@ME_SUPAG", "DOR", "Date of Registration", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@ME_SUPAG", "AGGTYP", "Agreement Type", 50)
        objMain.objUtilities.AddDateField("@ME_SUPAG", "FRMDT", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@ME_SUPAG", "TODT", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@ME_SUPAG", "SERIES", "Series", 50)
        objMain.objUtilities.AddAlphaField("@ME_SUPAG", "JENO", "JounalEntryNo", 50)
        objMain.objUtilities.AddAlphaField("@ME_SUPAG", "BRANCH", "Branch", 100)
        objMain.objUtilities.AddAlphaField("@ME_SUPAG", "REMARKS", "Remarks", 254)
        objMain.objUtilities.AddFloatField("@ME_SUPAG", "OTRCHG", "OTRegCharge", SAPbobsCOM.BoFldSubTypes.st_Price)

        objMain.objUtilities.CreateTable("ME_SUPAG_C0", "Supplier Agreement Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@ME_SUPAG_C0", "PRAGGNO", "PurchaseAgrNo", 50)
        objMain.objUtilities.AddDateField("@ME_SUPAG_C0", "STRDT", "StartDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@ME_SUPAG_C0", "ENDDT", "EndDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@ME_SUPAG_C0", "STATUS", "Status", 50)
        objMain.objUtilities.AddAlphaField("@ME_SUPAG_C0", "REMARKS", "Remarks", 254)
    End Sub

    Sub CreateSupplierRebateTable()
        objMain.objUtilities.CreateTable("SBO_SUPREBATE", "Supplier Rebate Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@SBO_SUPREBATE", "BPCD", "Supplier Code", 50)
        objMain.objUtilities.AddAlphaField("@SBO_SUPREBATE", "BPNM", "Supplier Name", 254)
        objMain.objUtilities.AddAlphaField("@SBO_SUPREBATE", "SERIES", "Series", 50)
        objMain.objUtilities.AddFloatField("@SBO_SUPREBATE", "INVTLDSC", "InvTotalDiscount", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddAlphaField("@SBO_SUPREBATE", "BASENT", "BaseEntry", 50)
        objMain.objUtilities.AddAlphaField("@SBO_SUPREBATE", "BASENO", "BaseNumber", 50)
        objMain.objUtilities.AddAlphaField("@SBO_SUPREBATE", "REMARK", "Remarks", 254)
        objMain.objUtilities.AddFloatField("@SBO_SUPREBATE", "ITAmt", "ITAmt", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddAlphaField("@SBO_SUPREBATE", "Status", "Status", 50)
        objMain.objUtilities.AddAlphaField("@SBO_SUPREBATE", "AppID", "AppID", 50)


        objMain.objUtilities.CreateTable("SBO_SUPREBATE_C0", "Supplier Rebate Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddFloatField("@SBO_SUPREBATE_C0", "POAF", "PurchaseofAmtFrom", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@SBO_SUPREBATE_C0", "POAT", "PurchaseofAmtTo", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@SBO_SUPREBATE_C0", "RBTPRC", "RebatePercent", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@SBO_SUPREBATE_C0", "RAmount", "RebateAmount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@SBO_SUPREBATE_C0", "Alert", "Alert", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddAlphaField("@SBO_SUPREBATE_C0", "Brand", "Brand", 50)

        objMain.objUtilities.CreateTable("SBO_SUPREBATE_C1", "Supplier Branch Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@SBO_SUPREBATE_C1", "CHECKBOX", "Check Box", 2)
        objMain.objUtilities.AddAlphaField("@SBO_SUPREBATE_C1", "BRANCHID", "Branch Id", 50)
        objMain.objUtilities.AddAlphaField("@SBO_SUPREBATE_C1", "BRANCHNM", "Branch Name", 100)
        objMain.objUtilities.AddAlphaField("@SBO_SUPREBATE_C1", "REMARKS", "Remarks", 250)

        objMain.objUtilities.CreateTable("SBO_SUPREBATE_C2", "Supplier Branch C2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@SBO_SUPREBATE_C2", "BRCD", "BarCodex", 50)
        objMain.objUtilities.AddAlphaField("@SBO_SUPREBATE_C2", "ITMCD", "ItemCode", 100)
        objMain.objUtilities.AddAlphaField("@SBO_SUPREBATE_C2", "ITMNM", "ItemName", 200)
        objMain.objUtilities.AddFloatField("@SBO_SUPREBATE_C2", "QTY", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@SBO_SUPREBATE_C2", "UOM", "UOM", 10)



    End Sub


    Sub CreatePriceUpdateTable()
        objMain.objUtilities.CreateTable("ME_PRICEUPD", "PriceUpdate Hdr", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@ME_PRICEUPD", "ITMCD", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@ME_PRICEUPD", "ITMNM", "Item Name", 254)
        objMain.objUtilities.AddAlphaField("@ME_PRICEUPD", "APPSTS", "Approval Status", 50)
        objMain.objUtilities.AddAlphaField("@ME_PRICEUPD", "SERIES", "Series", 50)
        objMain.objUtilities.AddAlphaField("@ME_PRICEUPD", "UNITS", "Units", 50)
        objMain.objUtilities.AddAlphaField("@ME_PRICEUPD", "USRID", "UserId", 100)
        objMain.objUtilities.AddAlphaField("@ME_PRICEUPD", "Updated", "Update", 1)
        objMain.objUtilities.AddAlphaField("@ME_PRICEUPD", "ErrorMsg", "ErrorMsg", 254)
        objMain.objUtilities.AddFloatField("@ME_PRICEUPD", "NEWPRC", "New price", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@ME_PRICEUPD", "BARCD", "BarCode", 50)

        objMain.objUtilities.CreateTable("ME_PRICEUPD_C0", "Branch Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@ME_PRICEUPD_C0", "ITMCD", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@ME_PRICEUPD_C0", "BRCHID", "Branch Id", 50)
        objMain.objUtilities.AddAlphaField("@ME_PRICEUPD_C0", "BRCHNM", "Branch Name", 50)
        objMain.objUtilities.AddAlphaField("@ME_PRICEUPD_C0", "PRCID", "PriceList Id", 50)
        objMain.objUtilities.AddAlphaField("@ME_PRICEUPD_C0", "PRCLST", "Price List", 50)
        objMain.objUtilities.AddAlphaField("@ME_PRICEUPD_C0", "UOMENT", "Uom Entry", 50)
        objMain.objUtilities.AddAlphaField("@ME_PRICEUPD_C0", "UOMCD", "Uom Code", 50)
        objMain.objUtilities.AddFloatField("@ME_PRICEUPD_C0", "OLDPRC", "Old Price", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@ME_PRICEUPD_C0", "NEWPRC", "New price", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddDateField("@ME_PRICEUPD_C0", "UPDT", "Update Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@ME_PRICEUPD_C0", "ENDDT", "End Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@ME_PRICEUPD_C0", "USRCD", "User Code", 50)
        objMain.objUtilities.AddAlphaField("@ME_PRICEUPD_C0", "Packing", "Packing", 50)
    End Sub
    Sub CreatePriceConfiguration()
        objMain.objUtilities.CreateTable("SBO_PRICECONFIG", "PriceConfiguration Hdr", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddDateField("@SBO_PRICECONFIG", "DocDate", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@SBO_PRICECONFIG", "Branch", "Branch", 50)
        'objMain.objUtilities.AddFloatField("@SBO_PRICECONFIG", "Smargin", "Supplier margin", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        'objMain.objUtilities.AddFloatField("@SBO_PRICECONFIG", "Exprice", "Exprice", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@SBO_PRICECONFIG", "Selmar", "Supplier margin", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@SBO_PRICECONFIG", "Exchprce", "Exchange price", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddAlphaField("@SBO_PRICECONFIG", "VCode", "Vendor Code", 50)
        objMain.objUtilities.AddAlphaField("@SBO_PRICECONFIG", "VName", "Vendor Name", 100)
        objMain.objUtilities.AddAlphaField("@SBO_PRICECONFIG", "ItmGrp", "ItmGrp", 50)
        objMain.objUtilities.AddAlphaField("@SBO_PRICECONFIG", "FITMCODE", "FITMCODE", 50)
        objMain.objUtilities.AddAlphaField("@SBO_PRICECONFIG", "TITMCODE", "TITMCODE", 50)
        objMain.objUtilities.AddAlphaField("@SBO_PRICECONFIG", "Curren", "Curren", 50)

        objMain.objUtilities.CreateTable("SBO_PRICECONFIG_C0", "PriceConfiguration Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@SBO_PRICECONFIG_C0", "ItmCo", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@SBO_PRICECONFIG_C0", "BRCHID", "Branch Id", 50)
        objMain.objUtilities.AddAlphaField("@SBO_PRICECONFIG_C0", "UOMENT", "Uom Entry", 50)
        objMain.objUtilities.AddAlphaField("@SBO_PRICECONFIG_C0", "PRCID", "PriceList Id", 50)
        objMain.objUtilities.AddAlphaField("@SBO_PRICECONFIG_C0", "ItmNm", "Item Name", 100)
        objMain.objUtilities.AddAlphaField("@SBO_PRICECONFIG_C0", "Branch", "Branch", 50)
        objMain.objUtilities.AddAlphaField("@SBO_PRICECONFIG_C0", "Prlist", "PriceList", 50)
        objMain.objUtilities.AddAlphaField("@SBO_PRICECONFIG_C0", "UmCode", "Uom Code", 50)
        'objMain.objUtilities.AddFloatField("@SBO_PRICECONFIG_C0", "Price", "Price", SAPbobsCOM.BoFldSubTypes.st_Price)
        'objMain.objUtilities.AddFloatField("@SBO_PRICECONFIG_C0", "SMargn", "Supplier margin", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        'objMain.objUtilities.AddFloatField("@SBO_PRICECONFIG_C0", "ExchgeP", "Exprice", SAPbobsCOM.BoFldSubTypes.st_Price)
        'objMain.objUtilities.AddFloatField("@SBO_PRICECONFIG_C0", "NewPrce", "New price", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@SBO_PRICECONFIG_C0", "PriceC", "PriceC", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@SBO_PRICECONFIG_C0", "SMar", "Supplier margin", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@SBO_PRICECONFIG_C0", "ExchP", "Exprice", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@SBO_PRICECONFIG_C0", "Newpr", "New price", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddAlphaField("@SBO_PRICECONFIG_C0", "User", "User Code", 50)
    End Sub
    Sub CreateFuturePriceUpdateTable()
        objMain.objUtilities.CreateTable("ME_FPRICEUPD", "PriceUpdate Hdr", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@ME_FPRICEUPD", "PDOC", "Price Update DocNum", 50)
        objMain.objUtilities.AddAlphaField("@ME_FPRICEUPD", "ITMCD", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@ME_FPRICEUPD", "ITMNM", "Item Name", 254)
        objMain.objUtilities.AddAlphaField("@ME_FPRICEUPD", "APPSTS", "Approval Status", 50)
        objMain.objUtilities.AddAlphaField("@ME_FPRICEUPD", "SERIES", "Series", 50)
        objMain.objUtilities.AddAlphaField("@ME_FPRICEUPD", "UNITS", "Units", 50)
        objMain.objUtilities.AddAlphaField("@ME_FPRICEUPD", "USRID", "UserId", 100)
        objMain.objUtilities.AddAlphaField("@ME_FPRICEUPD", "Updated", "Update", 1)
        objMain.objUtilities.AddAlphaField("@ME_FPRICEUPD", "ErrorMsg", "ErrorMsg", 254)
        objMain.objUtilities.AddFloatField("@ME_FPRICEUPD", "NEWPRC", "New price", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@ME_FPRICEUPD", "BARCD", "BarCode", 50)

        objMain.objUtilities.CreateTable("ME_FPRICEUPD_C0", "Branch Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@ME_FPRICEUPD_C0", "ITMCD", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@ME_FPRICEUPD_C0", "BRCHID", "Branch Id", 50)
        objMain.objUtilities.AddAlphaField("@ME_FPRICEUPD_C0", "BRCHNM", "Branch Name", 50)
        objMain.objUtilities.AddAlphaField("@ME_FPRICEUPD_C0", "PRCID", "PriceList Id", 50)
        objMain.objUtilities.AddAlphaField("@ME_FPRICEUPD_C0", "PRCLST", "Price List", 50)
        objMain.objUtilities.AddAlphaField("@ME_FPRICEUPD_C0", "UOMENT", "Uom Entry", 50)
        objMain.objUtilities.AddAlphaField("@ME_FPRICEUPD_C0", "UOMCD", "Uom Code", 50)
        objMain.objUtilities.AddFloatField("@ME_FPRICEUPD_C0", "OLDPRC", "Old Price", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@ME_FPRICEUPD_C0", "NEWPRC", "New price", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddDateField("@ME_FPRICEUPD_C0", "UPDT", "Update Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@ME_FPRICEUPD_C0", "ENDDT", "End Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@ME_FPRICEUPD_C0", "USRCD", "User Code", 50)
    End Sub

    Sub CreatePDTSeriesConfig()
        objMain.objUtilities.CreateTable("MXL_PDTSERIES", "PDT Series Config", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement)
        objMain.objUtilities.AddAlphaField("@MXL_PDTSERIES", "DOCTYPE", "Document Type", 30)
        objMain.objUtilities.AddAlphaField("@MXL_PDTSERIES", "LOCCODE", "Location Code", 10)
        objMain.objUtilities.AddAlphaField("@MXL_PDTSERIES", "LOCNAME", "Location Name", 100)
        objMain.objUtilities.AddInteger("@MXL_PDTSERIES", "YEAR", "Year", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddInteger("@MXL_PDTSERIES", "SERIES", "Series", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@MXL_PDTSERIES", "SERIESNAME", "Series Name", 20)
    End Sub

    Sub PDTUserWiseBranchDetails()
        objMain.objUtilities.CreateTable("ME_PDTUBD", "PDTUserWiseBranchDetails", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement)
        objMain.objUtilities.AddAlphaField("@ME_PDTUBD", "USERID", "PDTUserId", 30)
        objMain.objUtilities.AddAlphaField("@ME_PDTUBD", "USERNM", "PDTUserName", 150)
        objMain.objUtilities.AddAlphaField("@ME_PDTUBD", "BRANCHID", "BranchId", 30)
        objMain.objUtilities.AddAlphaField("@ME_PDTUBD", "BRANCHNM", "BranchName", 254)
        objMain.objUtilities.AddAlphaField("@ME_PDTUBD", "BRNCHFNM", "BPLFrName", 254)

        objMain.objUtilities.addField("@ME_PDTUBD", "ACTIVE", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "Yes")
    End Sub

    Sub CreatePDTROLES()
        objMain.objUtilities.CreateTable("ME_PDTROLES", "PDT ROLES", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement)
    End Sub

    Sub BarcodeSetup()
        objMain.objUtilities.CreateTable("NT_BRST", "Barcode Setup", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        objMain.objUtilities.addField("@NT_BRST", "Itemgrpwise", "ItemGroupWise", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "1,2", "Specific,All", "")
        objMain.objUtilities.AddAlphaField("@NT_BRST", "Itmgrpnm", "ItemGroup Name", 250)
        objMain.objUtilities.addField("@NT_BRST", "Nodigit", "No. of Digit", SAPbobsCOM.BoFieldTypes.db_Alpha, 11, SAPbobsCOM.BoFldSubTypes.st_None, "6,13", "6,13", "13")
        objMain.objUtilities.AddAlphaField("@NT_BRST", "Itprefix", "ItemCode Prefix", 11)
        objMain.objUtilities.AddAlphaField("@NT_BRST", "Prefix", "Prefix", 11)
        objMain.objUtilities.AddAlphaField("@NT_BRST", "Start", "Start", 11)
        objMain.objUtilities.AddAlphaField("@NT_BRST", "Start1", "Start1", 11)
        objMain.objUtilities.AddAlphaField("@NT_BRST", "End", "End", 11)
        objMain.objUtilities.addField("@NT_BRST", "Scale", "Scale Item", SAPbobsCOM.BoFieldTypes.db_Alpha, 11, SAPbobsCOM.BoFldSubTypes.st_None, "1,2", "Yes,No", "2")
        objMain.objUtilities.AddAlphaField("@NT_BRST", "SubGrp", "Sub Group", 50)
        objMain.objUtilities.AddAlphaField("@NT_BRST", "SubSGrp", "SubSub Group", 50)
        objMain.objUtilities.AddAlphaField("@NT_BRST", "SubSSGp", "SubSubSub Group", 50)
        ' objMain.objUtility.addUDTLinkedField("OHEM", "CATGTYPE", "Categorytype", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, "SBO_CAT", "", "")
        'objMain.objUtilities.addUDTLinkedField("@NT_BRST", "Itmgpnm", "ItemGroup Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "52")
    End Sub

    Sub BarcodePrinter()
        objMain.objUtilities.CreateTable("NT_Barprsetup", "Barcode Printer Setup", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        objMain.objUtilities.addField("@NT_Barprsetup", "Type", "Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 11, SAPbobsCOM.BoFldSubTypes.st_None, "1,2,3", "Type1,Type2,Type3", "1") ''' Error
        objMain.objUtilities.AddAlphaField("@NT_Barprsetup", "PrinterName", "PrinterName", 30)
        objMain.objUtilities.AddAlphaField("@NT_Barprsetup", "PrinterPath", "PrinterPath", 254)
        objMain.objUtilities.AddAlphaField("@NT_Barprsetup", "BPLId", "Branch ID", 100) '''' Error
    End Sub

    Sub StandardDocuments()
        objMain.objUtilities.AddAlphaField("OBPL", "PricelistNo", "PricelistNo", 10)
        objMain.objUtilities.addField("OPLN", "BarPrice", "BarCode Price", SAPbobsCOM.BoFieldTypes.db_Alpha, 11, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "N")
        objMain.objUtilities.addField("OITB", "Nodigit", "No. of Digit", SAPbobsCOM.BoFieldTypes.db_Alpha, 11, SAPbobsCOM.BoFldSubTypes.st_None, "6,13", "6,13", "13")
        objMain.objUtilities.addField("OITB", "Itemgrpwise", "ItemGroupWise", SAPbobsCOM.BoFieldTypes.db_Alpha, 11, SAPbobsCOM.BoFldSubTypes.st_None, "1,2", "Specific,All", "1")
        objMain.objUtilities.AddAlphaField("OITM", "ShortName", "ShortName", 35)
        objMain.objUtilities.AddAlphaField("OITM", "packing", "Packing", 30)
        objMain.objUtilities.AddFloatField("OITM", "PIC", "No. of Pics Per Pack", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("OITM", "BsUnit", "Base Unit", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddDateField("ITM1", "UPDT", "UpdatedDate", SAPbobsCOM.BoFldSubTypes.st_None)
        '----Missing fields added here
        objMain.objUtilities.AddFloatField("POR1", "UnitPriceN", "Unit Price N", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("POR1", "SuppPri", "Supplier Price N", SAPbobsCOM.BoFldSubTypes.st_Sum)


        objMain.objUtilities.AddFloatField("POR1", "CNTRLWHS", "Central Warehouse Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("POR1", "ExTaDis", "Ex. Tax Dis %", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.addField("OWTR", "TrnsTyp", "Transaction Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, SAPbobsCOM.BoFldSubTypes.st_None, "-,AG,OS,NE,DI,O", "-,Aged Goods,Overstock,Near Expiry,DRA Items,Others", "-")

    End Sub

    Sub Barcode()
        objMain.objUtilities.CreateTable("ALE_BARCODE", "Bar Code Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.CreateTable("ALE_BARCODE1", "Bar Code Line", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@ALE_BARCODE", "DocEntry", "DocEntry", 20)
        objMain.objUtilities.AddAlphaField("@ALE_BARCODE", "DocNum", "Document No", 20)
        objMain.objUtilities.AddAlphaField("@ALE_BARCODE", "DocType", "Document Type", 20)
        objMain.objUtilities.AddDateField("@ALE_BARCODE", "DocDate", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@ALE_BARCODE", "CardCode", "Card Code", 15)
        objMain.objUtilities.AddAlphaField("@ALE_BARCODE", "CardName", "Card Name", 100)
        objMain.objUtilities.AddFloatField("@ALE_BARCODE", "PQty", "Print Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@ALE_BARCODE", "PriceUp", "Price Update No", 20)
        objMain.objUtilities.AddAlphaField("@ALE_BARCODE", "Branch", "Branch", 100)
        objMain.objUtilities.AddAlphaField("@ALE_BARCODE", "Printer", "Printer", 100)
        objMain.objUtilities.AddAlphaField("@ALE_BARCODE", "Type", "Type", 10)
        objMain.objUtilities.AddAlphaField("@ALE_BARCODE", "WithPrc", "With Price", 1)

        objMain.objUtilities.AddAlphaField("@ALE_BARCODE1", "ItemCode", "Item Code", 20)
        objMain.objUtilities.AddAlphaField("@ALE_BARCODE1", "ItemName", "Description", 100)
        objMain.objUtilities.AddAlphaField("@ALE_BARCODE1", "Uom", "Uom", 20)
        objMain.objUtilities.AddAlphaField("@ALE_BARCODE1", "UoMEntry", "UoMEntry", 20)
        objMain.objUtilities.AddAlphaField("@ALE_BARCODE1", "CodeBars", "Bar Code", 16)
        objMain.objUtilities.AddAlphaField("@ALE_BARCODE1", "FrgnName", "Foreign Name", 100)
        objMain.objUtilities.AddAlphaField("@ALE_BARCODE1", "ItmGrpNm", "Item Group Name", 20)
        objMain.objUtilities.AddAlphaField("@ALE_BARCODE1", "FirmName", "Manufacturer", 30)
        objMain.objUtilities.AddFloatField("@ALE_BARCODE1", "Qty", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.addField("@ALE_BARCODE1", "Print", "Print", SAPbobsCOM.BoFieldTypes.db_Alpha, 11, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "N")
        objMain.objUtilities.AddFloatField("@ALE_BARCODE1", "PrintQty", "Print Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
    End Sub

    Sub ClearanceSales()
        objMain.objUtilities.CreateTable("SBO_CLEAR", "Clearance sales", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.CreateTable("SBO_CLEAR_C0", "ClearanceSales", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddDateField("@SBO_CLEAR", "DocDate", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@SBO_CLEAR", "PostDate", "PostDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@SBO_CLEAR", "Branch", "Branch", 100)
        objMain.objUtilities.AddAlphaField("@SBO_CLEAR", "Whs", "Warehouse", 100)
        objMain.objUtilities.AddAlphaField("@SBO_CLEAR", "REMARKS", "Remarks", 254)
        objMain.objUtilities.AddAlphaField("@SBO_CLEAR", "MESTS", "Status", 254)

        objMain.objUtilities.AddAlphaField("@SBO_CLEAR_C0", "ICode", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@SBO_CLEAR_C0", "IName", "Item Name", 200)
        objMain.objUtilities.AddAlphaField("@SBO_CLEAR_C0", "BarCode", "Bar Code", 100)
        objMain.objUtilities.AddAlphaField("@SBO_CLEAR_C0", "UOM", "UOM Code", 20)
        objMain.objUtilities.AddFloatField("@SBO_CLEAR_C0", "Qty", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@SBO_CLEAR_C0", "Price", "Price", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@SBO_CLEAR_C0", "Prod", "Production No", 20)
        objMain.objUtilities.AddAlphaField("@SBO_CLEAR_C0", "RProd", "Receipt From Production", 20)

        objMain.objUtilities.AddAlphaField("@SBO_CLEAR_C0", "ProdNo", "Production No", 20)
        objMain.objUtilities.AddAlphaField("@SBO_CLEAR_C0", "RProdNo", "Receipt Production No", 20)

        objMain.objUtilities.AddAlphaField("@SBO_CLEAR_C0", "RCode", "Rct Item Code", 50)
        objMain.objUtilities.AddAlphaField("@SBO_CLEAR_C0", "RName", "Rct Item Name", 200)
        objMain.objUtilities.AddAlphaField("@SBO_CLEAR_C0", "RBarCode", "Rct Bar Code", 100)
        objMain.objUtilities.AddAlphaField("@SBO_CLEAR_C0", "RUOM", "Rct UOM Code", 20)
        objMain.objUtilities.AddFloatField("@SBO_CLEAR_C0", "RQty", "Rct Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@SBO_CLEAR_C0", "RPrice", "Rct Price", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@SBO_CLEAR_C0", "Status", "Status", 250)
        objMain.objUtilities.AddFloatField("@SBO_CLEAR_C0", "SSTK", "System Stock", SAPbobsCOM.BoFldSubTypes.st_Quantity)
    End Sub

    Sub SupplierAgreementPosting()
        objMain.objUtilities.CreateTable("ME_SUP", "Supplier Agree post", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@ME_SUP", "BPCD", "Supplier Code", 50)
        objMain.objUtilities.AddAlphaField("@ME_SUP", "BPNM", "Supplier Name", 254)
        objMain.objUtilities.AddAlphaField("@ME_SUP", "SERIES", "Seroies", 50)
        objMain.objUtilities.AddFloatField("@ME_SUP", "INVTLDSC", "InvTotalDiscount", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddAlphaField("@ME_SUP", "BASENT", "BaseEntry", 50)
        objMain.objUtilities.AddAlphaField("@ME_SUP", "BASENO", "BaseNumber", 50)
        objMain.objUtilities.AddAlphaField("@ME_SUP", "REMARK", "Remarks", 254)
        objMain.objUtilities.AddAlphaField("@ME_SUP", "BRANCH", "Branch", 100)
        objMain.objUtilities.AddDateField("@ME_SUP", "DocDate", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.addField("@ME_SUP", "MOPay", "Mode of Payment", SAPbobsCOM.BoFieldTypes.db_Alpha, 25, SAPbobsCOM.BoFldSubTypes.st_None, "Cash,Cheque,Bank Transfer", "Cash,Cheque,Bank Transfer", "Cash")


        objMain.objUtilities.CreateTable("ME_SUP_C0", "Supplier Agree Child2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@ME_SUP_C0", "DESC", "Description", 100)
        objMain.objUtilities.AddAlphaField("@ME_SUP_C0", "JESTATUS", "JE Status", 250)
        objMain.objUtilities.AddFloatField("@ME_SUP_C0", "AMT", "Front End Percent", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@ME_SUP_C0", "PTG", "Percent", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddAlphaField("@ME_SUP_C0", "JE", "JE Entry", 30)
        objMain.objUtilities.AddAlphaField("@ME_SUP_C0", "REMARKS", "Remarks", 254)

        objMain.objUtilities.CreateTable("ME_SUP_C1", "Supplier Agree Child2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@ME_SUP_C1", "DESC", "Description", 100)
        objMain.objUtilities.AddFloatField("@ME_SUP_C1", "AMT", "Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddAlphaField("@ME_SUP_C1", "PENT", "Pent", 250)
        objMain.objUtilities.AddAlphaField("@ME_SUP_C1", "REMARKS", "Remarks", 254)
        objMain.objUtilities.AddAlphaField("@ME_SUP_C1", "JENum", "JE Entry", 30)
        objMain.objUtilities.AddFloatField("@ME_SUP_C1", "VAT", "VAT %", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddAlphaField("@ME_SUP_C1", "VATACC", "VAT Account", 50)
        objMain.objUtilities.AddAlphaField("@ME_SUP_C1", "BRANCH", "Branch", 50)
        objMain.objUtilities.AddAlphaField("@ME_SUP_C1", "VATCODE", "Vat Code", 50)
        objMain.objUtilities.AddFloatField("@ME_SUP_C1", "VATAMT", "Vat Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)

        objMain.objUtilities.AddFloatField("@ME_SUP_C1", "TVAT", "Total Vat", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@ME_SUP_C1", "VAMUNT", "Total Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)

        objMain.objUtilities.CreateTable("ME_SUP_C2", "Supplier Agree Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@ME_SUP_C2", "GROUP", "Item Group", 30)
        objMain.objUtilities.AddAlphaField("@ME_SUP_C2", "BRANCH", "Branch", 30)
        objMain.objUtilities.AddFloatField("@ME_SUP_C2", "AMT", "Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@ME_SUP_C2", "FCENT", "Front End Percent", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@ME_SUP_C2", "BCENT", "Back End Percent", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddAlphaField("@ME_SUP_C2", "REMARKS", "Remarks", 254)


    End Sub
    Sub BuyGetFreeItem()
        objMain.objUtilities.CreateTable("TNX_BGFREEITMS", "Buy and Get Free Items", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddDateField("@TNX_BGFREEITMS", "FRMDATE", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_BGFREEITMS", "TODATE", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_BGFREEITMS", "FRMTIME", "From Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddDateField("@TNX_BGFREEITMS", "TOTIME", "To Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.addField("@TNX_BGFREEITMS", "STATUS", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "ACTIVE,INACTIVE", "ACTIVE,INACTIVE", "")
        objMain.objUtilities.AddAlphaField("@TNX_BGFREEITMS", "User", "User", 50)
        objMain.objUtilities.AddAlphaField("@TNX_BGFREEITMS", "PName", "PName", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGFREEITMS", "Path", "Path", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGFREEITMS", "IGroup", "IGroup", 100)
        objMain.objUtilities.addField("@TNX_BGFREEITMS", "PRMTYP", "Promotion Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Quantity,Value", "Quantity,Value", "")
        objMain.objUtilities.AddFloatField("@TNX_BGFREEITMS", "MXQBTME", "Maximum Buy Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_BGFREEITMS", "MXQGTME", "Maximum Get Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_BGFREEITMS", "BUYQTY", "Buy Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_BGFREEITMS", "GTQTY", "Get Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.CreateTable("TNX_BGFREEITMS_C0", "Buy and Get Free Items Child1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_BGFREEITMS_C0", "BCODE", "Barcode", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGFREEITMS_C0", "ITMCODE", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGFREEITMS_C0", "ITMNM", "Item Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGFREEITMS_C0", "UOM", "UOM", 10)
        objMain.objUtilities.AddFloatField("@TNX_BGFREEITMS_C0", "RTLPRICE", "Retail Price", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_BGFREEITMS_C0", "PAIDQTY", "Paid Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_BGFREEITMS_C0", "MAXFREEQTY", "Maximum Free Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_BGFREEITMS_C0", "FREEQTY", "Free Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.CreateTable("TNX_BGFREEITMS_C1", "Buy and Get Free Items Child2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_BGFREEITMS_C1", "BCODE", "Barcode", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGFREEITMS_C1", "ITMCODE", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGFREEITMS_C1", "ITMNM", "Item Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGFREEITMS_C1", "UOM", "UOM", 10)
        objMain.objUtilities.AddFloatField("@TNX_BGFREEITMS_C1", "RTLPRICE", "Retail Price", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_BGFREEITMS_C1", "PAIDQTY", "Paid Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_BGFREEITMS_C1", "MAXFREEQTY", "Maximum Free Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_BGFREEITMS_C1", "FREEQTY", "Free Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
    End Sub

    Sub BuyNGetOneFreeItem()
        objMain.objUtilities.CreateTable("TNX_BNGOFITM", "Buy N Get One Free Item", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddDateField("@TNX_BNGOFITM", "FRMDATE", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_BNGOFITM", "TODATE", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_BNGOFITM", "FRMTIME", "From Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddDateField("@TNX_BNGOFITM", "TOTIME", "To Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.addField("@TNX_BNGOFITM", "STATUS", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "ACTIVE,INACTIVE", "ACTIVE,INACTIVE", "")
        objMain.objUtilities.AddAlphaField("@TNX_BNGOFITM", "User", "User Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_BNGOFITM", "PRONME", "Promotion Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BNGOFITM", "Path", "Path", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BNGOFITM", "IGroup", "IGroup", 100)
        objMain.objUtilities.addField("@TNX_BNGOFITM", "PRMTYP", "Promotion Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Quantity,Value", "Quantity,Value", "")
        objMain.objUtilities.AddFloatField("@TNX_BNGOFITM", "MXQBTME", "Maximum Buy Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_BNGOFITM", "MXQGTME", "Maximum Get Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_BNGOFITM", "BUYQTY", "Buy Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_BNGOFITM", "GTQTY", "Get Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_BNGOFITM", "FREEQTY", "Free Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_BNGOFITM", "FREEITM", "Free Item", 100)
        objMain.objUtilities.AddFloatField("@TNX_BNGOFITM", "PURPRM", "Purchase Amount for Promotion", SAPbobsCOM.BoFldSubTypes.st_Price)


        objMain.objUtilities.CreateTable("TNX_BNGOFITM_C0", "Buy N Get One Free Item Child1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_BNGOFITM_C0", "BCODE", "Barcode", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BNGOFITM_C0", "ITMCODE", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BNGOFITM_C0", "ITMDESC", "Item Description", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BNGOFITM_C0", "UOM", "UOM", 10)
        objMain.objUtilities.AddFloatField("@TNX_BNGOFITM_C0", "PRICE", "Price", SAPbobsCOM.BoFldSubTypes.st_Price)

        objMain.objUtilities.CreateTable("TNX_BNGOFITM_C1", "Buy N Get One Free Item Child2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_BNGOFITM_C1", "BCODE", "Barcode", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BNGOFITM_C1", "ITMCODE", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BNGOFITM_C1", "ITMDESC", "Item Description", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BNGOFITM_C1", "UOM", "UOM", 10)
        objMain.objUtilities.AddFloatField("@TNX_BNGOFITM_C1", "PRICE", "Price", SAPbobsCOM.BoFldSubTypes.st_Price)

    End Sub
    Sub BuyGetLeastItem()
        objMain.objUtilities.CreateTable("TNX_BGLEASTITMS", "Buy and Get Least Items", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddDateField("@TNX_BGLEASTITMS", "FRMDATE", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_BGLEASTITMS", "TODATE", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_BGLEASTITMS", "FRMTIME", "From Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddDateField("@TNX_BGLEASTITMS", "TOTIME", "To Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.addField("@TNX_BGLEASTITMS", "STATUS", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "ACTIVE,INACTIVE", "ACTIVE,INACTIVE", "")
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS", "User", "User", 50)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS", "PName", "PName", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS", "SubGrp", "Sub Group", 50)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS", "SSbGrp", "Sub Sub Group", 50)
        objMain.objUtilities.AddLinkField("@TNX_BGLEASTITMS", "Path", "Path", 254, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS", "IGroup", "IGroup", 100)
        objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS", "PRMLMT", "Promotion Limit", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        ' objMain.objUtilities.addField("@TNX_BGLEASTITMS", "PRMTYP", "Promotion Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Quantity,Value", "Quantity,Value", "")

        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS", "PRTYP", "Promotion Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS", "BPLID", "Branch", 100)

        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS", "PRMNM", "Promotion Name", 50)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS", "SBCDE", "Search Barcodes", 50)



        objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS", "MXQBQTY", "Maximum Buy Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS", "MXQGQTY", "Maximum Get Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS", "GITMD", "Get Item Discount %", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddDateField("@TNX_BGLEASTITMS", "DOCDATE", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)


        'objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS", "MXQBTME", "Maximum Buy Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS", "MXQGTME", "Maximum Get Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS", "BUYQTY", "Buy Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS", "GTQTY", "Get Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.CreateTable("TNX_BGLEASTITMS_C0", "Buy and Get Least Items Child1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C0", "BCODE", "Barcode", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C0", "ITMCODE", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C0", "ITMNM", "Item Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C0", "UOM", "UOM", 10)
        objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS_C0", "PRICE", "Retail Price", SAPbobsCOM.BoFldSubTypes.st_Price)

        objMain.objUtilities.CreateTable("TNX_BGLEASTITMS_C1", "Buy and Get Least Items Child2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C1", "CHK", "Check", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C1", "ITMCDE", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C1", "ITMGRP", "Item Group", 100)
        'objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS_C1", "DPERC", "Discount Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)

        objMain.objUtilities.CreateTable("TNX_BGLEASTITMS_C2", "Buy and Get Least Items Child3", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C2", "CHEK", "Check", 1)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C2", "BCDE", "Branch Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C2", "BRNME", "Branch Name", 100)

        objMain.objUtilities.CreateTable("TNX_BGLEASTITMS_C3", "Buy and Get Least Items Child4", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C3", "BRCODE", "Barcode", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C3", "ITCODE", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C3", "ITNM", "Item Name", 100)

    End Sub

    Sub BuyGetLeastItem_old()
        objMain.objUtilities.CreateTable("TNX_BGLEASTITMS", "Buy and Get Least Items", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddDateField("@TNX_BGLEASTITMS", "FRMDATE", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_BGLEASTITMS", "TODATE", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_BGLEASTITMS", "FRMTIME", "From Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddDateField("@TNX_BGLEASTITMS", "TOTIME", "To Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.addField("@TNX_BGLEASTITMS", "STATUS", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "ACTIVE,INACTIVE", "ACTIVE,INACTIVE", "")
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS", "User", "User", 50)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS", "PName", "PName", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS", "Path", "Path", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS", "IGroup", "IGroup", 100)
        objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS", "PRMLMT", "Promotion Limit", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.addField("@TNX_BGLEASTITMS", "PRMTYP", "Promotion Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Quantity,Value", "Quantity,Value", "")
        objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS", "MXQBTME", "Maximum Buy Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS", "MXQGTME", "Maximum Get Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS", "BUYQTY", "Buy Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS", "GTQTY", "Get Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.CreateTable("TNX_BGLEASTITMS_C0", "Buy and Get Least Items Child1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C0", "BCODE", "Barcode", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C0", "ITMCODE", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C0", "ITMNM", "Item Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C0", "UOM", "UOM", 10)
        objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS_C0", "PRICE", "Retail Price", SAPbobsCOM.BoFldSubTypes.st_Price)

        objMain.objUtilities.CreateTable("TNX_BGLEASTITMS_C1", "Buy and Get Least Items Child2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C1", "BCODE", "Barcode", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C1", "ITMCODE", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C1", "ITMNM", "Item Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C1", "UOM", "UOM", 10)
        objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS_C1", "PRICE", "Retail Price", SAPbobsCOM.BoFldSubTypes.st_Price)

        objMain.objUtilities.CreateTable("TNX_BGLEASTITMS", "Buy and Get Least Items", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddDateField("@TNX_BGLEASTITMS", "FRMDATE", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_BGLEASTITMS", "TODATE", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_BGLEASTITMS", "FRMTIME", "From Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddDateField("@TNX_BGLEASTITMS", "TOTIME", "To Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.addField("@TNX_BGLEASTITMS", "STATUS", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "ACTIVE,INACTIVE", "ACTIVE,INACTIVE", "")
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS", "User", "User", 50)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS", "PName", "PName", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS", "SubGrp", "Sub Group", 50)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS", "SSbGrp", "Sub Sub Group", 50)
        objMain.objUtilities.AddLinkField("@TNX_BGLEASTITMS", "Path", "Path", 254, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS", "IGroup", "IGroup", 100)
        objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS", "PRMLMT", "Promotion Limit", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        ' objMain.objUtilities.addField("@TNX_BGLEASTITMS", "PRMTYP", "Promotion Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Quantity,Value", "Quantity,Value", "")

        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS", "PRTYP", "Promotion Type", 50)


        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS", "PRMNM", "Promotion Name", 50)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS", "SBCDE", "Search Barcodes", 50)



        objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS", "MXQBQTY", "Maximum Buy Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS", "GITMD", "Get Item Discount %", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddDateField("@TNX_BGLEASTITMS", "DOCDATE", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)


        'objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS", "MXQBTME", "Maximum Buy Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS", "MXQGTME", "Maximum Get Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS", "BUYQTY", "Buy Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS", "GTQTY", "Get Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.CreateTable("TNX_BGLEASTITMS_C0", "Buy and Get Least Items Child1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C0", "BCODE", "Barcode", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C0", "ITMCODE", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C0", "ITMNM", "Item Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C0", "UOM", "UOM", 10)
        objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS_C0", "PRICE", "Retail Price", SAPbobsCOM.BoFldSubTypes.st_Price)

        objMain.objUtilities.CreateTable("TNX_BGLEASTITMS_C1", "Buy and Get Least Items Child2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C1", "CHK", "Check", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C1", "ITMCDE", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C1", "ITMGRP", "Item Group", 100)
        'objMain.objUtilities.AddFloatField("@TNX_BGLEASTITMS_C1", "DPERC", "Discount Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)

        objMain.objUtilities.CreateTable("TNX_BGLEASTITMS_C2", "Buy and Get Least Items Child3", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C2", "CHEK", "Check", 1)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C2", "BCDE", "Branch Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C2", "BRNME", "Branch Name", 100)

        objMain.objUtilities.CreateTable("TNX_BGLEASTITMS_C3", "Buy and Get Least Items Child4", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C3", "BRCODE", "Barcode", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C3", "ITCODE", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_BGLEASTITMS_C3", "ITNM", "Item Name", 100)

    End Sub




    Sub ItemWiseDiscount()




        objMain.objUtilities.CreateTable("TNX_ITEMDISC", "Item Discount", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddDateField("@TNX_ITEMDISC", "FRMDATE", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_ITEMDISC", "TODATE", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_ITEMDISC", "FRMTIME", "From Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddDateField("@TNX_ITEMDISC", "TOTIME", "To Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddDateField("@TNX_ITEMDISC", "DOCDATE", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.addField("@TNX_ITEMDISC", "STATUS", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Active,Inactive", "Active,Inactive", "")
        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC", "PRMNM", "Promotion Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC", "BPLID", "Branch", 20)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC", "PRTYP", "Promotion Type", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC", "SBCDE", "Search BarCodes", 100)
        objMain.objUtilities.AddLinkField("@TNX_ITEMDISC", "Path", "Path", 254, SAPbobsCOM.BoFldSubTypes.st_Link)

        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC", "ITMGRP", "Item Group", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC", "SubGrp", "Sub Group", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC", "SubSubGrp", "Sub Sub Group", 100)

        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC", "DISCTYP", "Discount Type", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC", "PCRDoc", "PCR Document Number", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC", "Login", "Login User", 100)





        objMain.objUtilities.CreateTable("TNX_ITEMDISC_C0", "Item Discount Child1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC_C0", "BCODE", "Barcode", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC_C0", "ITMCODE", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC_C0", "ITMNM", "Item Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC_C0", "UOM", "UOM", 10)
        objMain.objUtilities.AddDateField("@TNX_ITEMDISC_C0", "FRMDATE", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_ITEMDISC_C0", "TODATE", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddFloatField("@TNX_ITEMDISC_C0", "ITMCST", "Item Cost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_ITEMDISC_C0", "MINAMTPURC", "Minimum Purchase Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_ITEMDISC_C0", "RTPRWVAT", "Retail Price With VAT", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_ITEMDISC_C0", "OFPRWVAT", "Offer Price With VAT", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_ITEMDISC_C0", "DCPRWVAT", "Discount Price With VAT", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_ITEMDISC_C0", "DPERCTG", "Discount Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)

        objMain.objUtilities.CreateTable("TNX_ITEMDISC_C1", "Item Discount Child2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC_C1", "CHK", "Check", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC_C1", "ITMCDE", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC_C1", "ITMGRP", "Item Group", 100)
        objMain.objUtilities.AddFloatField("@TNX_ITEMDISC_C1", "DPERC", "Discount Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_ITEMDISC_C1", "DISAMNT", "Discount Amounts", SAPbobsCOM.BoFldSubTypes.st_Price)


        objMain.objUtilities.CreateTable("TNX_ITEMDISC_C2", "Item Discount Child3", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC_C2", "CHK", "Check", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC_C2", "BCODE", "Branch Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC_C2", "BCHNME", "Branch Name", 100)

        objMain.objUtilities.CreateTable("TNX_ITEMDISC_C3", "Item Discount Child4", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC_C3", "BCODE", "Barcode", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC_C3", "ITMCODE", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC_C3", "ITMNM", "Item Description", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ITEMDISC_C3", "UOM", "UOM", 10)


    End Sub

    Sub PromotiononVoucherPurchase()
        objMain.objUtilities.CreateTable("TNX_PONVP", "Promotion Voucher Purchase", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@TNX_PONVP", "BPLID", "Branch", 100)

        objMain.objUtilities.AddDateField("@TNX_PONVP", "FRMDATE", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PONVP", "TODATE", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PONVP", "FRMTIME", "From Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddDateField("@TNX_PONVP", "TOTIME", "To Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddDateField("@TNX_PONVP", "DOCDATE", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.addField("@TNX_PONVP", "STATUS", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Active,Inactive", "Active,Inactive", "")
        objMain.objUtilities.AddAlphaField("@TNX_PONVP", "PName", "Promotion Name", 100)
        objMain.objUtilities.AddFloatField("@TNX_PONVP", "REDMNBAMNT", "Redemption Minimum Bill Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddAlphaField("@TNX_PONVP", "DTYPE", "Discount Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PONVP", "PTYPE", "PROMOTION Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PONVP", "IGROUP", "Item Group", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PONVP", "SUBGRP", "Sub Group", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PONVP", "SSBGRP", "Sub Sub Group", 50)


        objMain.objUtilities.CreateTable("TNX_PONVP_C0", "Promotion Voucher Child1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PONVP_C0", "VITMCDE", "Voucher Item Code", 100)
        objMain.objUtilities.AddFloatField("@TNX_PONVP_C0", "FRMAMNT", "From Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_PONVP_C0", "TOAMOUNT", "To Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@TNX_PONVP_C0", "VOHRORDIS", "Voucher or Discount", 100)
        objMain.objUtilities.AddFloatField("@TNX_PONVP_C0", "PERORAMONT", "Percentage Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_PONVP_C0", "PERCNTG", "Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_PONVP_C0", "DPRCTG", "Discount By Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_PONVP_C0", "DBAMT", "Dicount By Amount", SAPbobsCOM.BoFldSubTypes.st_Price)


        objMain.objUtilities.CreateTable("TNX_PONVP_C1", "Promotion Voucher Child2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PONVP_C1", "BARCD", "Bar Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PONVP_C1", "ITMCDE", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PONVP_C1", "ITMDESC", "Item Description", 100)

        objMain.objUtilities.CreateTable("TNX_PONVP_C2", "Promotion Voucher Child3", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PONVP_C2", "CHK", "Check", 1)
        objMain.objUtilities.AddAlphaField("@TNX_PONVP_C2", "BCHNME", "Branch Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PONVP_C2", "BCODE", "Branch Code", 100)


    End Sub

    Sub PromotiononBillTotal_OLD()
        objMain.objUtilities.CreateTable("TNX_PONBTOTL", "Promotion on Bill Total", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddDateField("@TNX_PONBTOTL", "FRMDATE", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PONBTOTL", "TODATE", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PONBTOTL", "FRMTIME", "From Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddDateField("@TNX_PONBTOTL", "TOTIME", "To Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.addField("@TNX_PONBTOTL", "STATUS", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "ACTIVE,INACTIVE", "ACTIVE,INACTIVE", "")
        objMain.objUtilities.AddAlphaField("@TNX_PONBTOTL", "User", "User", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PONBTOTL", "PName", "PName", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PONBTOTL", "Path", "Path", 100,)
        objMain.objUtilities.AddAlphaField("@TNX_PONBTOTL", "REDMNBAMNT", "Redemption Minimum Bill Amount", SAPbobsCOM.BoFldSubTypes.st_Price)


        objMain.objUtilities.CreateTable("TNX_PONBTOTL_C0", "Promotion on Bill Total Child1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PONBTOTL_C0", "IGroup", "IGroup", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PONBTOTL_C0", "sbItemGrp", "Sub Item Group", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PONBTOTL_C0", "sbbItemGrp", "Sub Sub Item Group", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PONBTOTL_C0", "Path", "Path", 100)
        objMain.objUtilities.AddFloatField("@TNX_PONBTOTL_C0", "FRMAMOUNT", "From Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_PONBTOTL_C0", "TOAMOUNT", "To Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.addField("@TNX_PONBTOTL_C0", "VOHRORDIS", "Voucher or Discount", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Voucher,Discount", "Voucher,Discount", "")
        objMain.objUtilities.AddFloatField("@TNX_PONBTOTL_C0", "PERORAMONT", "Percentage or Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_PONBTOTL_C0", "PERCNTG", "Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)

        objMain.objUtilities.CreateTable("TNX_PONBTOTL_C1", "Promotion on Bill Total Child2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PONBTOTL_C1", "IGroup", "IGroup", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PONBTOTL_C1", "sbItemGrp", "Sub Item Group", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PONBTOTL_C1", "sbbItemGrp", "Sub Sub Item Group", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PONBTOTL_C1", "Path1", "Path1", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PONBTOTL_C1", "ITMCDE", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PONBTOTL_C1", "ITMDESC", "Item Description", 100)

    End Sub

    Sub PromotiononBillTotal()
        objMain.objUtilities.CreateTable("TNX_PBILL", "Prmn BILL Total header", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_PBILL", "BPLID", "Branch", 100)

        objMain.objUtilities.AddDateField("@TNX_PBILL", "FRMDATE", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PBILL", "TODATE", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PBILL", "FRMTIME", "From Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddDateField("@TNX_PBILL", "TOTIME", "To Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.addField("@TNX_PBILL", "STATUS", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "ACTIVE,INACTIVE", "ACTIVE,INACTIVE", "")
        objMain.objUtilities.AddAlphaField("@TNX_PBILL", "PNAME", "Promotion Name", 100)
        objMain.objUtilities.AddFloatField("@TNX_PBILL", "MAXAP", "Max Amount Price", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddAlphaField("@TNX_PBILL", "ITMGRP1", "Item Group", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PBILL", "SubGrp1", "Sub Item Group", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PBILL", "SubSubGrp1", "Sub Sub Item Group", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PBILL", "SBCODE", "Search BarCodes", 100)
        objMain.objUtilities.AddLinkField("@TNX_PBILL", "PATH", "Path", 254, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_PBILL", "DISCTYP1", "Discount Type", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PBILL", "PRTYP", "Promotion Type", 100)
        objMain.objUtilities.AddDateField("@TNX_PBILL", "DOCDATE", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)


        objMain.objUtilities.CreateTable("TNX_PBILL_C0", "Prmn BILL Total C0", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PBILL_C0", "BCODE", "BAR CODE", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PBILL_C0", "ITMCODE", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PBILL_C0", "ITMNM", "Item Description", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PBILL_C0", "UOM", "UOM", 10)
        objMain.objUtilities.AddFloatField("@TNX_PBILL_C0", "PRICE", "PRICE", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_PBILL_C0", "DAMNT", "Discount by Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PBILL_C0", "DPERCENT", "Discount by Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_PBILL_C0", "DAMT", "Discount by Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddAlphaField("@TNX_PBILL_C0", "REMK", "Remarks", 100)
        objMain.objUtilities.AddFloatField("@TNX_PBILL_C0", "OFFP", "Offer Price", SAPbobsCOM.BoFldSubTypes.st_Sum)

        objMain.objUtilities.CreateTable("TNX_PBILL_C1", "Prmn BILL Total C1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PBILL_C1", "CHCK", "Checck", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PBILL_C1", "ITMMCDE", "Itemm Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PBILL_C1", "ITMMGRP", "Itemm Group", 100)

        objMain.objUtilities.CreateTable("TNX_PBILL_C2", "Prmn BILL Total C2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PBILL_C2", "CHK", "Check", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PBILL_C2", "BRCODE", "Branch Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PBILL_C2", "BCHNME", "Branch Name", 100)

        objMain.objUtilities.CreateTable("TNX_PBILL_C3", "Prnn BILL Total C3", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PBILL_C3", "BCCODE", "Baar Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PBILL_C3", "ITTMCODE", "Itemm Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PBILL_C3", "ITMNM", "Itemm Description", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PBILL_C3", "UOM", "UOOM", 100)



    End Sub

    Sub CreateStockCountGenerationwizard()
        objMain.objUtilities.CreateTable("SBO_STOCKWZD", "StockCount WizardHeader", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@SBO_STOCKWZD", "TYPE", "Type", 30)
        objMain.objUtilities.AddAlphaField("@SBO_STOCKWZD", "ITGRP", "ItemGroup", 30)
        objMain.objUtilities.AddAlphaField("@SBO_STOCKWZD", "ITEMGRP", "ItemGroup", 30)
        objMain.objUtilities.AddAlphaField("@SBO_STOCKWZD", "SBGRP", "SubGroup", 30)
        objMain.objUtilities.AddAlphaField("@SBO_STOCKWZD", "SBSBGRP", "SubSubGroup", 30)
        objMain.objUtilities.AddDateField("@SBO_STOCKWZD", "PSDT", "PostDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@SBO_STOCKWZD", "WHSE", "WhseCode", 30)
        objMain.objUtilities.AddAlphaField("@SBO_STOCKWZD", "BRANCH", "Branch", 30)
        objMain.objUtilities.AddDateField("@SBO_STOCKWZD", "FROMDT", "PostDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@SBO_STOCKWZD", "TODT", "PostDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@SBO_STOCKWZD", "INEXITM", "INEXITM", 1)
        objMain.objUtilities.AddAlphaMemoField("@SBO_STOCKWZD", "STATUS", "STATUS", 6000)
        objMain.objUtilities.AddAlphaField("@SBO_STOCKWZD", "INCNT", "Inventory counting number", 30)


        objMain.objUtilities.CreateTable("SBO_STOCKWZD_C0", "StockCountwizard Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@SBO_STOCKWZD_C0", "Check", "Check", 1)
        objMain.objUtilities.AddAlphaField("@SBO_STOCKWZD_C0", "ICode", "ItemCode", 50)
        objMain.objUtilities.AddAlphaField("@SBO_STOCKWZD_C0", "IName", "ItemName", 254)
        objMain.objUtilities.AddAlphaField("@SBO_STOCKWZD_C0", "BaseUom", "BaseUom", 30)
        objMain.objUtilities.AddFloatField("@SBO_STOCKWZD_C0", "Qty", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@SBO_STOCKWZD_C0", "Update", "Update", 1)
        objMain.objUtilities.AddAlphaField("@SBO_STOCKWZD_C0", "Flag", "StatusFlag", 50)
        objMain.objUtilities.AddFloatField("@SBO_STOCKWZD_C0", "Onhand", "OnHand", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@SBO_STOCKWZD_C0", "DOCE", "DocEntry", 50)
        objMain.objUtilities.AddAlphaField("@SBO_STOCKWZD_C0", "LINE", "LineID", 50)

    End Sub


    Sub PromoCostReduction()
        objMain.objUtilities.CreateTable("TNX_RDCTN", "Promo Cost Reduction", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddDateField("@TNX_RDCTN", "DOCDATE", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_RDCTN", "BRANCH", "Branch", 100)


        objMain.objUtilities.AddAlphaField("@TNX_RDCTN", "BYNAME", "Buyer Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_RDCTN", "SPCODE", "Supplier Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_RDCTN", "GLACNT", "G/L Account", 100)
        objMain.objUtilities.AddAlphaField("@TNX_RDCTN", "SPNAME", "Supplier Name", 100)
        objMain.objUtilities.AddDateField("@TNX_RDCTN", "PSTART", "PRMN START Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_RDCTN", "PEND", "PRMN END Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddFloatField("@TNX_RDCTN", "TOTAL", "Total", SAPbobsCOM.BoFldSubTypes.st_Price)

        objMain.objUtilities.AddAlphaField("@TNX_RDCTN", "STS", "Status", 100)

        objMain.objUtilities.AddFloatField("@TNX_RDCTN", "PRMEDT", "Alrt Bfr E Date", SAPbobsCOM.BoFldSubTypes.st_Sum)

        'objMain.objUtilities.AddInteger("@TNX_RDCTN", "PRMEDT", "Alrt Bfr Prmn E Date", 30)

        objMain.objUtilities.CreateTable("TNX_RDCTN_C0", "Promo Cost Reduction Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_RDCTN_C0", "BARCD", "BarCode", 100)

        objMain.objUtilities.AddFloatField("@TNX_RDCTN_C0", "CSTDTL", "Cost Difference Total", SAPbobsCOM.BoFldSubTypes.st_Price)


        objMain.objUtilities.AddAlphaField("@TNX_RDCTN_C0", "ITMCODE", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_RDCTN_C0", "ITMDESC", "Item Description", 100)
        objMain.objUtilities.AddAlphaField("@TNX_RDCTN_C0", "ISTQTY", "Item in Stk Qty", 100)
        objMain.objUtilities.AddFloatField("@TNX_RDCTN_C0", "OLDCST", "Old Item Cost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_RDCTN_C0", "NEWCST", "New Item Cost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_RDCTN_C0", "COSTD", "Cost Difference", SAPbobsCOM.BoFldSubTypes.st_Price)

        objMain.objUtilities.AddAlphaField("@TNX_RDCTN_C0", "JRMRK", "Journal Remark", 100)
        objMain.objUtilities.AddAlphaField("@TNX_RDCTN_C0", "BRANCH", "BRANCH", 100)
        objMain.objUtilities.AddAlphaField("@TNX_RDCTN_C0", "WRHSE", "Warehouse", 100)
        objMain.objUtilities.AddAlphaField("@TNX_RDCTN_C0", "REF1", "REF-1 ", 100)
        objMain.objUtilities.AddAlphaField("@TNX_RDCTN_C0", "REF2", "REF1-2 ", 100)
        objMain.objUtilities.AddAlphaField("@TNX_RDCTN_C0", "IREVNO", "Inv Revaluation No", 50)
        'objMain.objUtilities.addField("@TNX_RDCTN_C0", "PSTS", "Posting Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 4, SAPbobsCOM.BoFldSubTypes.st_None, "Open,Close", "Open,Close", "Open")
        objMain.objUtilities.AddAlphaField("@TNX_RDCTN_C0", "PSTS", "Posting Status", 50)
        objMain.objUtilities.AddAlphaField("@TNX_RDCTN_C0", "REVDRF", "Revaluation Doc Ref", 50)

    End Sub
    Sub BarcodeReport()
        objMain.objUtilities.CreateTable("TNX_BCRRTIM", "Barcode Report", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@TNX_BCRRTIM", "BRNCH", "Branch", 50)
    End Sub
    Sub BarcodeFilter()
        objMain.objUtilities.CreateTable("TNX_FIND", "Barcode Filter", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@TNX_FIND", "FIND", "FIND", 50)
    End Sub
    Sub ItemMasterUploadPreview()
        objMain.objUtilities.CreateTable("TNX_ITMPEVIEW", "Item Master Upload Preview", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@TNX_ITMPEVIEW", "Branch", "Branch", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ITMPEVIEW", "AppID", "Application ID", 50)
    End Sub

    Sub STOCKDET()
        'Header Feields Stock Details
        objMain.objUtilities.CreateTable("STOCKD", "STOCK DETAILS", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddDateField("@STOCKD", "DOCDT", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@STOCKD", "STATUS", "Status", 100)
        objMain.objUtilities.AddAlphaField("@STOCKD", "BName", "Branch Name", 100)

        'Row Fields Stock Details
        objMain.objUtilities.CreateTable("C_STOCKD", "STOCK DETAILS Row", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@C_STOCKD", "ItemC", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@C_STOCKD", "Disc", "Item Discription", 100)
        objMain.objUtilities.AddAlphaField("@C_STOCKD", "UOM", "UOM", 100)
        objMain.objUtilities.AddAlphaField("@C_STOCKD", "BarC", "Barcode", 100)
        objMain.objUtilities.AddFloatField("@C_STOCKD", "InStoc", "In Stock", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@C_STOCKD", "PreSto", "Present Stock", SAPbobsCOM.BoFldSubTypes.st_Price)

    End Sub
    Sub PACPPriceChange()

        objMain.objUtilities.CreateTable("TNX_PACP", "PACPPriceChangeRM Header", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddDateField("@TNX_PACP", "Date", "Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_PACP", "VName", "Vendor Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PACP", "Barcode", "Barcode", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PACP", "IName", "Item Name", 100)
        objMain.objUtilities.AddFloatField("@TNX_PACP", "Excost", "Existing cost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_PACP", "Ncost", "New cost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_PACP", "NCV", "New cost+vat", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_PACP", "NSelling", "New Selling", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_PACP", "Margin", "Margin", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@TNX_PACP", "Remark", "Remark", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PACP", "ICode", "Item Code", 100)
        objMain.objUtilities.AddLinkField("@TNX_PACP", "Excel", "Excel", 80, SAPbobsCOM.BoFldSubTypes.st_Link)


        'Attachments Tab
        'objMain.objUtilities.CreateTable("TNX_PACP_C0", "PACPPriceChangeMaster Child", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        'objMain.objUtilities.AddLinkField("@TNX_PACP_C0", "TPath", "Target Path", 80, SAPbobsCOM.BoFldSubTypes.st_Link)
        'objMain.objUtilities.AddAlphaField("@TNX_PACP_C0", "FName", "File Name", 200)
        'objMain.objUtilities.AddDateField("@TNX_PACP_C0", "ADate", "Attachment Date", SAPbobsCOM.BoFldSubTypes.st_None)
        'objMain.objUtilities.AddAlphaField("@TNX_PACP_C0", "FText", "Free Text", 254)

    End Sub

    Sub PromoCostReversion()
        objMain.objUtilities.CreateTable("TNX_REVERSION", "Pomotion Cost Revision", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@TNX_REVERSION", "TOTAL", "Total", 100)

        objMain.objUtilities.AddAlphaField("@TNX_REVERSION", "BNAME", "Buyer code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_REVERSION", "SCODE", "Supplier Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_REVERSION", "SNAME", "Supplier Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_REVERSION", "PCRNO", "Prm cost reduction no", 100)
        objMain.objUtilities.AddDateField("@TNX_REVERSION", "PDTRG", "Prm Date Range", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_REVERSION", "DOCDATE", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_REVERSION", "STATUS", "Document Status", 100)
        objMain.objUtilities.AddDateField("@TNX_REVERSION", "FRMDATE", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_REVERSION", "TODATE", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddFloatField("@TNX_REVERSION", "TSCNAMT", "Total Sup Cred Note Amt ", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddAlphaField("@TNX_REVERSION", "CNTNO", "Credit note No", 100)
        objMain.objUtilities.AddAlphaField("@TNX_REVERSION", "CNPBR", "Credit Note Post Branch ", 100)
        objMain.objUtilities.AddAlphaField("@TNX_REVERSION", "GLACT", "GLAccount", 100)

        objMain.objUtilities.CreateTable("TNX_REVERSION_C0", "Pomotion Cost Revision Child1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_REVERSION_C0", "ICODE", "Item Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_REVERSION_C0", "CSTDI", "Cost Difference", 100)


        objMain.objUtilities.AddFloatField("@TNX_REVERSION_C0", "CSTDFT", "Cost Difference Tota1", SAPbobsCOM.BoFldSubTypes.st_Price)


        objMain.objUtilities.AddAlphaField("@TNX_REVERSION_C0", "ITDSP", "Item Description", 100)
        objMain.objUtilities.AddFloatField("@TNX_REVERSION_C0", "CSQUN", "Current In Stock Quty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_REVERSION_C0", "PICOST", "Promo Item Cost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_REVERSION_C0", "RICOST", "Revision Item Cost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@TNX_REVERSION_C0", "JRRMK", "Journal Remark", 100)
        objMain.objUtilities.AddAlphaField("@TNX_REVERSION_C0", "REF1", "Reference 1", 100)
        objMain.objUtilities.AddAlphaField("@TNX_REVERSION_C0", "REF2", "Reference 2", 100)
        objMain.objUtilities.AddAlphaField("@TNX_REVERSION_C0", "WRHOS", "Warehouse", 100)
        objMain.objUtilities.AddAlphaField("@TNX_REVERSION_C0", "BRANCH", "Branch", 50)
        objMain.objUtilities.AddAlphaField("@TNX_REVERSION_C0", "PSTATUS", "Posting Status", 50)
        objMain.objUtilities.AddAlphaField("@TNX_REVERSION_C0", "INRNO", "Inventory Revaluation Number", 100)
        objMain.objUtilities.AddAlphaField("@TNX_REVERSION_C0", "RDREF", "Revalution Doc Ref", 100)
        objMain.objUtilities.CreateTable("TNX_REVERSION_C1", "Pomotion Cost Revision Child2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddFloatField("@TNX_REVERSION_C1", "OCOST", "Old Cost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_REVERSION_C1", "PRCOST", "Promotion Cost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_REVERSION_C1", "TSAMT", "Total Sales Amt", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_REVERSION_C1", "SQNTY", "Sold Quantity ", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_REVERSION_C1", "AMOUNT", "Amount", SAPbobsCOM.BoFldSubTypes.st_Measurement)
    End Sub
    Sub GRPODOCRp()
        objMain.objUtilities.CreateTable("TNX_GRPODR", "GRPO DocumentWise Report", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddDateField("@TNX_GRPODR", "FDate", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_GRPODR", "TDate", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_GRPODR", "BName", "BranchName", 50)


    End Sub

    Sub BarChk()
        objMain.objUtilities.CreateTable("TNX_BARCHK", "BarCode Check", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@TNX_BARCHK", "BName", "BarCode", 30)


    End Sub

    Sub GRPOItmRp()
        objMain.objUtilities.CreateTable("TNX_GRPOITM", "GRPO ItemWise Report", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddDateField("@TNX_GRPOITM", "FDate", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_GRPOITM", "TDate", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_GRPOITM", "BName", "BranchName", 50)


    End Sub

    Sub OpenDisp()
        objMain.objUtilities.CreateTable("TNX_OPDSPR", "Open Dispatch Report", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddDateField("@TNX_OPDSPR", "FDate", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_OPDSPR", "TDate", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)


    End Sub

    Sub TransferRep()
        objMain.objUtilities.CreateTable("TNX_TRNSRP", "Transfer Receipt Report", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddDateField("@TNX_TRNSRP", "FDate", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)


    End Sub
    Sub Stocksumm()
        objMain.objUtilities.CreateTable("TNX_STCKCNT", "GRPO ItemWise Report", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddDateField("@TNX_STCKCNT", "FDate", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_STCKCNT", "TDate", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_STCKCNT", "WHSE", "WareHouseName", 50)
        objMain.objUtilities.AddAlphaField("@TNX_STCKCNT", "GRPName", "GroupName", 50)
        objMain.objUtilities.AddAlphaField("@TNX_STCKCNT", "BName", "SubGroupName", 50)


    End Sub

    Sub StockWHSE()
        objMain.objUtilities.CreateTable("TNX_STKWHSE", "Stock By Warehouse", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@TNX_STKWHSE", "BName", "GroupName", 30)


    End Sub

    Sub APInvoiceItem()
        objMain.objUtilities.CreateTable("TNX_APINR", "AP Invoice Itemwise", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddDateField("@TNX_APINR", "FDate", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_APINR", "TDate", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_APINR", "BName", "BranchName", 50)


    End Sub

    Sub CatWiseOB()
        objMain.objUtilities.CreateTable("TNX_CATWIN", "Category Wise Inv OB", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@TNX_CATWIN", "BName", "BranchName", 30)


    End Sub

    Sub ExciseItems()
        objMain.objUtilities.CreateTable("TNX_EXITM", "Excise Items Report", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddDateField("@TNX_EXITM", "FDate", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_EXITM", "TDate", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_EXITM", "BName", "BranchName", 50)
        objMain.objUtilities.AddAlphaField("@TNX_EXITM", "VName", "VendorName", 50)


    End Sub

    Sub ARDocWise()
        objMain.objUtilities.CreateTable("TNX_ARINR", "AR Invoice DocumentWise Report", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddDateField("@TNX_ARINR", "FDate", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_ARINR", "TDate", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_ARINR", "BName", "BranchName", 50)


    End Sub

    Sub ARItemWise()
        objMain.objUtilities.CreateTable("TNX_ARITMR", "AR Invoice ItemWise Report", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddDateField("@TNX_ARITMR", "FDate", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_ARITMR", "TDate", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_ARITMR", "BName", "BranchName", 50)


    End Sub

    Sub BrandMaster()
        objMain.objUtilities.CreateTable("TNX_BRAND", "Brand Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@TNX_BRAND", "Status", "Status", 1)
    End Sub
    Sub SubGroup()
        objMain.objUtilities.CreateTable("TNX_SUBGROUP", "SubGroup Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@TNX_SUBGROUP", "IGroup", "Item Group", 50)
        objMain.objUtilities.AddAlphaField("@TNX_SUBGROUP", "SubGroup", "Sub Group", 50)
        objMain.objUtilities.AddAlphaField("@TNX_SUBGROUP", "Status", "Status", 1)
    End Sub
    Sub SubSubGroup()
        objMain.objUtilities.CreateTable("TNX_SSGROUP", "SubSubGroup Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        'objMain.objUtilities.AddAlphaField("@TNX_SSGROUP", "IGroup", "Item Group", 50)
        objMain.objUtilities.AddAlphaField("@TNX_SSGROUP", "SGroup", "Sub Group", 50)
        objMain.objUtilities.AddAlphaField("@TNX_SSGROUP", "Status", "Status", 1)
    End Sub
    Sub SubSubSubGroup()
        objMain.objUtilities.CreateTable("TNX_SSSGROUP", "SubSubSubGroup Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        'objMain.objUtilities.AddAlphaField("@TNX_SSGROUP", "IGroup", "Item Group", 50)
        objMain.objUtilities.AddAlphaField("@TNX_SSSGROUP", "SSGroup", "Sub Sub Group", 50)
        objMain.objUtilities.AddAlphaField("@TNX_SSSGROUP", "Status", "Status", 1)
    End Sub

    Sub VR_SUBGROUP()
        objMain.objUtilities.CreateTable("VR_SUBGROUP", "Sub Group", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement)
        objMain.objUtilities.AddAlphaField("@VR_SUBGROUP", "Name", "Name", 100)

        objMain.objUtilities.AddAlphaField("@VR_SUBGROUP", "ITEMGROUP", "ITEM GROUP", 100)
        objMain.objUtilities.AddAlphaField("@VR_SUBGROUP", "SUBGROUP", "SUB GROUP", 100)
    End Sub
    Sub VR_SUBSUBGROUP()
        objMain.objUtilities.CreateTable("VR_SUBSUBGROUP", "Sub Sub Group", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement)
        objMain.objUtilities.AddAlphaField("@VR_SUBSUBGROUP", "SUBGROUP", "SUBGROUP", 100)
        objMain.objUtilities.addField("@VR_SUBSUBGROUP", "WithoutPO", "WithoutPO", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
        objMain.objUtilities.AddFloatField("@VR_SUBSUBGROUP", "Margin", "Margin", SAPbobsCOM.BoFldSubTypes.st_Price)
    End Sub
    Sub VR_SUBSUBSUBGROUP()
        objMain.objUtilities.CreateTable("VR_SUBSUBSUBGROUP", "Sub Sub Sub Group", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement)
        objMain.objUtilities.AddAlphaField("@VR_SUBSUBSUBGROUP", "SUBSUBGROUP", "SUBSUB GROUP", 100)

    End Sub
    Sub DepartmentMaster()
        objMain.objUtilities.CreateTable("TNX_DMASTER", "Department Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@TNX_DMASTER", "STATUS", "Status", 1)
    End Sub

    Sub Messanger_Log()
        objUtilities.CreateTable("SBO_MSGLOG", "Messanger Log", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        objUtilities.AddDateField("@SBO_MSGLOG", "Date", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUtilities.AddAlphaField("@SBO_MSGLOG", "DocType", "Document Type", 30)
        objUtilities.AddAlphaField("@SBO_MSGLOG", "MsgType", "Message Type", 30)
        objUtilities.AddAlphaField("@SBO_MSGLOG", "DocEntry", "Document Entry", 15)
        objUtilities.AddAlphaField("@SBO_MSGLOG", "DocNum", "Document No", 15)
        objUtilities.AddAlphaField("@SBO_MSGLOG", "CardCode", "Customer Code", 15)
        objUtilities.AddAlphaField("@SBO_MSGLOG", "Status", "Status", 254)
        objUtilities.AddAlphaField("@SBO_MSGLOG", "EmailID", "Email", 100)
        objUtilities.AddAlphaField("@SBO_MSGLOG", "Mobile", "Mobile No", 15)
        objUtilities.AddAlphaField("@SBO_MSGLOG", "Subject", "Subject", 254)
        objUtilities.AddAlphaMemoField("@SBO_MSGLOG", "MSG", "Message", 5000)
        objUtilities.AddAlphaMemoField("@SBO_MSGLOG", "Response", "Response", 5000)
        objUtilities.AddAlphaField("@SBO_MSGLOG", "ResMessage", "Response Msg", 254)
    End Sub

    Sub Tnx_Retail_Log()
        objUtilities.CreateTable("TNX_RETAIL_LOG", "Tnx_Retail_Log", SAPbobsCOM.BoUTBTableType.bott_NoObject)


    End Sub
    Sub ApCreditMemoNoObject()
        objUtilities.CreateTable("SBO_APCM", "ApCreditMemo", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        objUtilities.AddDateField("@SBO_APCM", "Date", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objUtilities.AddAlphaField("@SBO_APCM", "VName", "Vendor Name", 30)
        objUtilities.AddAlphaField("@SBO_APCM", "InNumber", "Invoice Number", 30)
        objUtilities.AddAlphaField("@SBO_APCM", "Amount", "Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objUtilities.AddAlphaField("@SBO_APCM", "Status", "Status", 1)

    End Sub
    Sub MessageConfig()
        objUtilities.CreateTable("MGRCONFIG", "MessageConfig", SAPbobsCOM.BoUTBTableType.bott_Document)
        objUtilities.AddAlphaField("@MGRCONFIG", "Doc", "Document", 100)
        objUtilities.AddAlphaField("@MGRCONFIG", "Branch", "Branch", 100)
        objUtilities.AddAlphaField("@MGRCONFIG", "Cust", "Customer", 1)
        objUtilities.AddAlphaField("@MGRCONFIG", "Emp", "Employee", 1)
        objUtilities.AddAlphaField("@MGRCONFIG", "User", "User", 1)
        objUtilities.AddAlphaField("@MGRCONFIG", "MgsTemp", "MgsTemp", 50)

        objUtilities.AddAlphaField("@MGRCONFIG", "AutoS", "AutomaticSMS", 1)
        objUtilities.AddAlphaField("@MGRCONFIG", "ManualS", "ManualSMS", 1)

        objUtilities.AddAlphaField("@MGRCONFIG", "AutoE", "AutomaticEmail", 1)
        objUtilities.AddAlphaField("@MGRCONFIG", "ManualE", "ManualEmail", 1)

        objUtilities.AddAlphaField("@MGRCONFIG", "Sub", "Subject", 200)
        objUtilities.AddAlphaField("@MGRCONFIG", "Auto", "Automatic", 1)
        objUtilities.AddAlphaField("@MGRCONFIG", "Manual", "Manual", 1)
        objUtilities.AddAlphaField("@MGRCONFIG", "Status", "Status", 1)

        objUtilities.AddAlphaField("@MGRCONFIG", "CustSMS", "SMSCustomer", 1)
        objUtilities.AddAlphaField("@MGRCONFIG", "EmpSMS", "SMSEmployee", 1)
        objUtilities.AddAlphaField("@MGRCONFIG", "UserSMS", "SMSUser", 1)

        objUtilities.AddAlphaField("@MGRCONFIG", "CustEML", "EmailCustomer", 1)
        objUtilities.AddAlphaField("@MGRCONFIG", "EmpEML", "EmailEmployee", 1)
        objUtilities.AddAlphaField("@MGRCONFIG", "UserEML", "EmailUser", 1)

        objUtilities.AddAlphaMemoField("@MGRCONFIG", "Msg", "Message", 5000)
        objUtilities.AddAlphaMemoField("@MGRCONFIG", "CRN", "Report Name", 5000)
        objUtilities.AddAlphaField("@MGRCONFIG", "Att", "Attachment", 1)
        objUtilities.AddAlphaField("@MGRCONFIG", "AttM", "Attachment", 1)

        objUtilities.AddAlphaMemoField("@MGRCONFIG", "MsgSMS", "SMSMessage", 5000)
        objUtilities.AddAlphaField("@MGRCONFIG", "AttSMS", "SMSAttachment", 1)
        objUtilities.AddAlphaField("@MGRCONFIG", "AttSMSM", "SMSAttachment", 1)

        objUtilities.AddAlphaMemoField("@MGRCONFIG", "MsgEML", "EMLMessage", 5000)
        objUtilities.AddAlphaField("@MGRCONFIG", "AttEML", "EMLAttachment", 1)
        objUtilities.AddAlphaField("@MGRCONFIG", "AttEMLM", "EMLAttachment", 1)
        objUtilities.AddAlphaField("@MGRCONFIG", "ReportID", "Report Id", 50)
        objUtilities.AddAlphaField("@MGRCONFIG", "SignEML", "Email Signature", 200)


        objUtilities.CreateTable("MGRCONFIG_C0", "MassageConfig c0", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUtilities.AddAlphaField("@MGRCONFIG_C0", "UserID", "UserID", 30)
        objUtilities.AddAlphaField("@MGRCONFIG_C0", "EmpID", "Employee ID", 30)

        objUtilities.CreateTable("MGRCONFIG_C1", "MassageConfig c1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUtilities.AddAlphaField("@MGRCONFIG_C1", "UserIDS", "UserID", 30)
        objUtilities.AddAlphaField("@MGRCONFIG_C1", "EmpIDS", "Employee ID", 30)

        objUtilities.CreateTable("MGRCONFIG_C2", "MassageConfig c2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objUtilities.AddAlphaField("@MGRCONFIG_C2", "UserIDE", "UserID", 30)
        objUtilities.AddAlphaField("@MGRCONFIG_C2", "EmpIDE", "Employee ID", 30)



    End Sub
    Sub SMS_EMail_Config_Table()
        objUtilities.CreateTable("SBO_SETCONFIG", "SMS/EMail Config Table", SAPbobsCOM.BoUTBTableType.bott_Document)
        objUtilities.AddAlphaField("@SBO_SETCONFIG", "SMSUrl", "SMS URL", 240)
        objUtilities.AddAlphaField("@SBO_SETCONFIG", "SMTPsevr", "SMTP Server", 100)
        objUtilities.AddAlphaField("@SBO_SETCONFIG", "SMPTPort", "SMTP Port", 50)
        objUtilities.AddAlphaField("@SBO_SETCONFIG", "EmailFro", "Email From", 100)
        objUtilities.AddAlphaField("@SBO_SETCONFIG", "EmailPwd", "Email Password", 100)
        objUtilities.AddAlphaField("@SBO_SETCONFIG", "CC", "CC Email ID", 240)
        objUtilities.AddAlphaField("@SBO_SETCONFIG", "BCC", "BCC Email ID", 240)
        objUtilities.AddAlphaMemoField("@SBO_SETCONFIG", "Name", "Name", 1000)
        objUtilities.AddAlphaField("@SBO_SETCONFIG", "SMSUId", "SMS User ID", 100)
        objUtilities.AddAlphaField("@SBO_SETCONFIG", "SMSPwd", "SMS Password", 100)
        objUtilities.AddAlphaField("@SBO_SETCONFIG", "WUId", "Whatsapp ID", 100)
        objUtilities.AddAlphaField("@SBO_SETCONFIG", "WPwd", "Whatsapp Pwd", 100)
        objUtilities.AddAlphaField("@SBO_SETCONFIG", "WMobile", "Whatsapp Mobile", 20)
        objUtilities.AddAlphaField("@SBO_SETCONFIG", "SMobile", "SMS Mobile", 20)
        objUtilities.AddAlphaField("@SBO_SETCONFIG", "SSL", "SSL", 1)


    End Sub

    Sub Configuration()
        objUtilities.CreateTable("CONFIG", "Configuration", SAPbobsCOM.BoUTBTableType.bott_Document)
        objUtilities.AddAlphaField("@CONFIG", "SBA", "Sales Blanket Agree", 1)
        objUtilities.AddAlphaField("@CONFIG", "PBA", "Purchase Blanket Agree", 1)

        objUtilities.AddAlphaField("@CONFIG", "SO", "Sales Order", 1)
        objUtilities.AddAlphaField("@CONFIG", "SQ", "Sales Quotation", 1)
        objUtilities.AddAlphaField("@CONFIG", "DL", "Delivery", 1)
        objUtilities.AddAlphaField("@CONFIG", "ARINV", "A/R Invoice", 1)

        objUtilities.AddAlphaField("@CONFIG", "ARRINV", "A/R Reserver Inv", 1)

        objUtilities.AddAlphaField("@CONFIG", "ARRET", "Return", 1)
        objUtilities.AddAlphaField("@CONFIG", "ARC", "A/R Credit Memo", 1)

        objUtilities.AddAlphaField("@CONFIG", "PQ", "Purchase Quotation", 1)
        objUtilities.AddAlphaField("@CONFIG", "PO", "Purchase Order", 1)
        objUtilities.AddAlphaField("@CONFIG", "GRPO", "GRPO", 1)
        objUtilities.AddAlphaField("@CONFIG", "APINV", "A/P Invoice", 1)
        objUtilities.AddAlphaField("@CONFIG", "APC", "A/P Credit Memo", 1)
        objUtilities.AddAlphaField("@CONFIG", "GR", "Goods Return", 1)

        objUtilities.AddAlphaField("@CONFIG", "GRP", "Goods Receipt PO", 1)
        objUtilities.AddAlphaField("@CONFIG", "GI", "Goods Issue", 1)
        objUtilities.AddAlphaField("@CONFIG", "ITR", "Inventory Trans Req", 1)
        objUtilities.AddAlphaField("@CONFIG", "IT", "Inventory Trans ", 1)

        objUtilities.AddAlphaField("@CONFIG", "IP", "Incoming Payment ", 1)
        objUtilities.AddAlphaField("@CONFIG", "OP", "Outgoing Payment ", 1)
        objUtilities.AddAlphaField("@CONFIG", "JE", "Journal Entry", 1)
        objUtilities.AddAlphaField("@CONFIG", "PRD", "Production Order", 1)

        'SMS Config
        objUtilities.AddAlphaField("@CONFIG", "SBASMS", "Sales Blanket Agree", 1)
        objUtilities.AddAlphaField("@CONFIG", "PBASMS", "Purchase Blanket Agree", 1)

        objUtilities.AddAlphaField("@CONFIG", "SOSMS", "Sales Order", 1)
        objUtilities.AddAlphaField("@CONFIG", "SQSMS", "Sales Quotation", 1)
        objUtilities.AddAlphaField("@CONFIG", "DLSMS", "Delivery", 1)
        objUtilities.AddAlphaField("@CONFIG", "ARINVSMS", "A/R Invoice", 1)
        objUtilities.AddAlphaField("@CONFIG", "ARRETSMS", "Return", 1)
        objUtilities.AddAlphaField("@CONFIG", "ARCSMS", "A/R CreditMemo", 1)

        objUtilities.AddAlphaField("@CONFIG", "PQSMS", "Purchase Quotation", 1)
        objUtilities.AddAlphaField("@CONFIG", "POSMS", "Purchase Order", 1)
        objUtilities.AddAlphaField("@CONFIG", "GRPOSMS", "GRPO", 1)
        objUtilities.AddAlphaField("@CONFIG", "APINVSMS", "A/P Invoice", 1)
        objUtilities.AddAlphaField("@CONFIG", "APCSMS", "A/P Invoice", 1)
        objUtilities.AddAlphaField("@CONFIG", "GRSMS", "Goods Return", 1)

        objUtilities.AddAlphaField("@CONFIG", "GRPSMS", "Goods Receipt PO", 1)
        objUtilities.AddAlphaField("@CONFIG", "GISMS", "Goods Issue", 1)
        objUtilities.AddAlphaField("@CONFIG", "ITRSMS", "Inventory Trans Req", 1)
        objUtilities.AddAlphaField("@CONFIG", "ITSMS", "Inventory Trans ", 1)

        objUtilities.AddAlphaField("@CONFIG", "IPSMS", "Incoming Payment ", 1)
        objUtilities.AddAlphaField("@CONFIG", "OPSMS", "Outgoing Payment ", 1)
        objUtilities.AddAlphaField("@CONFIG", "JESMS", "Journal Entry", 1)
        objUtilities.AddAlphaField("@CONFIG", "PRDSMS", "Production Order", 1)


        'Email
        objUtilities.AddAlphaField("@CONFIG", "SBAML", "Sales Blanket Agree", 1)
        objUtilities.AddAlphaField("@CONFIG", "PBAML", "Purchase Blanket Agree", 1)

        objUtilities.AddAlphaField("@CONFIG", "SOEML", "Sales Order", 1)
        objUtilities.AddAlphaField("@CONFIG", "SQEML", "Sales Quotation", 1)
        objUtilities.AddAlphaField("@CONFIG", "DLEML", "Delivery", 1)
        objUtilities.AddAlphaField("@CONFIG", "ARINVEML", "A/R Invoice", 1)
        objUtilities.AddAlphaField("@CONFIG", "ARRETEML", "Return", 1)

        objUtilities.AddAlphaField("@CONFIG", "PQEML", "Purchase Quotation", 1)
        objUtilities.AddAlphaField("@CONFIG", "POEML", "Purchase Order", 1)
        objUtilities.AddAlphaField("@CONFIG", "GRPOEML", "GRPO", 1)
        objUtilities.AddAlphaField("@CONFIG", "APINVEML", "A/P Invoice", 1)
        objUtilities.AddAlphaField("@CONFIG", "GREML", "Goods Return", 1)

        objUtilities.AddAlphaField("@CONFIG", "GRPEML", "Goods Receipt PO", 1)
        objUtilities.AddAlphaField("@CONFIG", "GIEML", "Goods Issue", 1)
        objUtilities.AddAlphaField("@CONFIG", "ITREML", "Inventory Trans Req", 1)
        objUtilities.AddAlphaField("@CONFIG", "ITEML", "Inventory Trans ", 1)

        objUtilities.AddAlphaField("@CONFIG", "IPEML", "Incoming Payment ", 1)
        objUtilities.AddAlphaField("@CONFIG", "OPEML", "Outgoing Payment ", 1)
        objUtilities.AddAlphaField("@CONFIG", "JEEML", "Journal Entry", 1)
        objUtilities.AddAlphaField("@CONFIG", "PRDEML", "Production Order", 1)
        objUtilities.AddAlphaField("@CONFIG", "PRUP", "Price Update", 1)


    End Sub

    Sub Item_Master_Data()
        objMain.objUtilities.AddFloatField("OITM", "EXCVAT", "Excise Vat", SAPbobsCOM.BoFldSubTypes.st_Sum)

    End Sub

    Sub AP_Invoice()
        ' objMain.objUtilities.AddAlphaField("OPCH", "EXCTAX", "Excise Tax", 50)
        ' objMain.objUtilities.AddAlphaField("PCH1", "EXCTAX", "Excise Tax", 50)
        objMain.objUtilities.AddFloatField("PCH1", "EXCTAX", "Excise Tax", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddAlphaField("OPCH", "EXCJVN", "Excise JV Number", 15)
    End Sub

    Sub AR_Invoice()
        objMain.objUtilities.AddAlphaField("OINV", "EDJE", "Excise JE Number", 15)
    End Sub
    Sub APCreditMemo()
        objMain.objUtilities.AddAlphaField("ORPC", "EXCJVN1", "Excise JV Number", 15)
    End Sub

    Sub Generalsettings()
        objMain.objUtilities.AddAlphaField("OADM", "PEDCG", "Purchase excise duty Cr GL", 50)
        objMain.objUtilities.AddAlphaField("OADM", "PEDDRG", "Purchase excise duty Dr GL", 50)
        objMain.objUtilities.AddAlphaField("OADM", "SEDDG", "Sales excise duty Dr GL", 50)

    End Sub
    Sub CompanyLevel()
        'objMain.objUtilities.AddAlphaField("OADM", "LCUSTSRS", "Local Customers Series", 50)
        'objMain.objUtilities.AddAlphaField("OADM", "FCUSTSRS", "Foreign Customers Series", 50)
        'objMain.objUtilities.AddAlphaField("OADM", "FCDLCUST", "Freight Code of Local Customers", 50)
    End Sub
    Sub Freight()
        objMain.objUtilities.addField("OEXD", "BPTYPE", "Business Patner Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "LOCAL,FOREIGN", "LOCAL,FOREIGN", "")
        objMain.objUtilities.AddAlphaField("OEXD", "BPSERIES", "Business Patner Series", 50)
    End Sub
    Sub Tenx_Temp1()
        objUtilities.CreateTable("TENX_TEMP1", "Temporary Table 1", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        objUtilities.AddAlphaField("@TENX_TEMP1", "COLUMN0", "Column 0", 10)
        objUtilities.AddAlphaField("@TENX_TEMP1", "COLUMN1", "Column 1", 100)
        objUtilities.AddAlphaField("@TENX_TEMP1", "COLUMN2", "Column 2", 200)
        objUtilities.AddAlphaField("@TENX_TEMP1", "COLUMN3", "Column 3", 100)
        objUtilities.AddAlphaField("@TENX_TEMP1", "COLUMN4", "Column 4", 100)
        objUtilities.AddFloatField("@TENX_TEMP1", "COLUMN5", "Column 5", SAPbobsCOM.BoFldSubTypes.st_Sum)
    End Sub
    Sub Tenx_Temp2()
        objUtilities.CreateTable("TENX_TEMP2", "Temporary Table 2", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        objUtilities.AddAlphaField("@TENX_TEMP2", "COLUMN0", "Column 0", 10)
        objUtilities.AddAlphaField("@TENX_TEMP2", "COLUMN1", "Column 1", 100)
        objUtilities.AddAlphaField("@TENX_TEMP2", "COLUMN2", "Column 2", 200)
        objUtilities.AddAlphaField("@TENX_TEMP2", "COLUMN3", "Column 3", 100)
        objUtilities.AddAlphaField("@TENX_TEMP2", "COLUMN4", "Column 4", 100)
        objUtilities.AddFloatField("@TENX_TEMP2", "COLUMN5", "Column 5", SAPbobsCOM.BoFldSubTypes.st_Sum)
    End Sub
    'Approvals
    Sub CreateAPPROVALSTAGES()
        objMain.objUtilities.CreateTable("SBO_AST", "Approval Stages", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@SBO_AST", "SName", "Stage Name", 40)
        objMain.objUtilities.AddAlphaField("@SBO_AST", "SDesc", "Stage Description", 40)
        objMain.objUtilities.AddAlphaField("@SBO_AST", "NAP", "No. of Approvals Required", 40)
        objMain.objUtilities.AddAlphaField("@SBO_AST", "NRJ", "No. of Rejections Required", 40)
        objMain.objUtilities.AddAlphaField("@SBO_AST", "NAME", "NAME", 40)


        objMain.objUtilities.CreateTable("SBO_AST_C0", "Approval Stages", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@SBO_AST_C0", "AUTH", "AUTH", 40)
        objMain.objUtilities.AddAlphaField("@SBO_AST_C0", "EmpId", "EmpId", 50)
        objMain.objUtilities.AddAlphaField("@SBO_AST_C0", "Dept", "Dept", 30)
        objMain.objUtilities.AddAlphaField("@SBO_AST_C0", "UKey", "UKey", 30)
    End Sub

    Sub CreateAPPROVALTemplates()
        objMain.objUtilities.CreateTable("SBO_APPHDR", "Temp", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "T1", "Stage Name", 40)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "T2", "Stage Description", 40)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "NAP", "No. of Approvals Required", 40)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "C1", "Active", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "PCR", "Price Change Request", 254)



        objMain.objUtilities.CreateTable("SBO_APPREQ", "Originator", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@SBO_APPREQ", "M1_1", "Name", 40)
        objMain.objUtilities.AddAlphaField("@SBO_APPREQ", "EmpId", "EmpId", 50)
        objMain.objUtilities.AddAlphaField("@SBO_APPREQ", "M1_2", "Department", 30)

        objMain.objUtilities.CreateTable("SBO_APPAUT", "Stages", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@SBO_APPAUT", "M3_1", "M3_1", 40)
        objMain.objUtilities.AddAlphaField("@SBO_APPAUT", "M3_2", "M3_2", 50)


        objMain.objUtilities.AddAlphaField("@SBO_APPAUT", "NAMES", "Names", 30)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "Name", "Name", 50)

        objMain.objUtilities.CreateTable("SBO_APPHDR", "Conditions", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "Conds", "Always", 50)


        objMain.objUtilities.AddAlphaField("@SBO_APPREQ", "Name", "EmpId", 50)

        objMain.objUtilities.AddAlphaField("@SBO_APPREQ", "Dept", "Department", 50)

        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "Descrip", "Desciption", 100)

        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "Acive", "Active", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "Active", "Active", 1)

        'objMain.objUtilities.CreateTable("SBO_APPHDR", "Conditions", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        'objMain.objUtilities.AddButton("@SBO_APPHDR", "M3_1", "M3_1", 40)
        'objMain.objUtilities.AddButton(objform.UniqueID, "Btn_Reset", objform.Items.Item("71").Top + 230,_  objform.Items.Item("73").Left, 80, objform.Items.Item("71").Height + 10, "71", "&Reset", objform.Items.Item("71").FromPane, objform.Items.Item("71").ToPane)

        objMain.objUtilities.CreateTable("SBO_APPDOC", "Documents", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "RcReq", "Recruitment Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "RmSm", "Resume Submission", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "OL", "Offer Letter", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "VRQ", "Advertisement Request", 1) 'C'
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Price", "Price Change Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "SuppReb", "SupplierRebate", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "IMP", "ItemMasterUpload", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "ITEML", "ItemListing", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "ITEML", "ItemListing", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "BP", "BusinessPartner", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "PACPPCR", "PACPPriceChangeReq", 1)


        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "EmpMstr", "Employee Master Data", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "EmpSal", "Employee Salary Setup", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "ConNew", "Employee Contract Renewal", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "TS", "TimeSheet", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "LvRj", "Leave Rejoin Application", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "LvApp", "Leave Application", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "CompOff", "Compensatory off", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "AttSum", "Attendance Summary", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Shift", "Shift Schedule", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "AReg", "Attendance Regularisation", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "LER", "Leave Encashment Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "LPC", "Leave Period End Closing", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "LnApp", "Loans and Advances Application", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "LnPr", "Loan Pre Closure Application", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "PD", "Payments and Deductions", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "PrPre", "Payroll Pre-Process", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Claim", "Claim Application", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Review", "Employee Review / Appraisal", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "PerEval", "Performance Evaluation", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Encash", "Leave Encashment", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "EOS", "End of Service Process", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Resign", "Employee Resignation Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Clear", "Exit Clearance", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Air", "Airticket Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Doc", "Document Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "PTC", "Petty Cash Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "PTV", "Petty Cash Voucher", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "MTR", "Material Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Trans", "Branch Transfer", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "GRPMTNG", "Group Meeting", 1)


    End Sub
    Sub CreateTenxLogSettinTbl()
        objMain.objUtilities.CreateTable("TNX_LOGST", "Tenx Add on Log Settings", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@TNX_LOGST", "CMBSQLV", "Level", 15)
        objMain.objUtilities.AddAlphaField("@TNX_LOGST", "LODDAYS", "Days", 4)
        objMain.objUtilities.AddAlphaField("@TNX_LOGST", "CMBENBL", "Enable Log", 15)
    End Sub

    'No Object Table
    Sub CreateTENXLOGTable()

        objMain.objUtilities.CreateTable("TNX_TALES", "Tenx Add on Log Error Table", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        objMain.objUtilities.AddDateField("@TNX_TALES", "EntDate", "Enter Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_TALES", "ErrTime", "Error Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddAlphaField("@TNX_TALES", "DocNum", "Document No", 20)
        objMain.objUtilities.AddAlphaField("@TNX_TALES", "USERID", "Login UserId", 100)
        objMain.objUtilities.AddAlphaField("@TNX_TALES", "AdName", "Add-on Name", 200)
        objMain.objUtilities.AddAlphaField("@TNX_TALES", "Level", "Level", 15)
        objMain.objUtilities.AddAlphaField("@TNX_TALES", "Adscop", "Scope", 200)
        objMain.objUtilities.AddAlphaMemoField("@TNX_TALES", "MSGTXT", "Detail Error", 64000)
    End Sub
    Sub CreateWarehouse()
        objMain.objUtilities.CreateTable("TNX_Warehouse", "Warehouse Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@TNX_Warehouse", "LineNum", "Line Number", 30)
        objMain.objUtilities.AddAlphaField("@TNX_Warehouse", "ItemCode", "Item Code", 30)
        objMain.objUtilities.AddAlphaField("@TNX_Warehouse", "ItemName", "Item Name", 30)
        objMain.objUtilities.AddAlphaField("@TNX_Warehouse", "FWhse", "From Whse", 30)
        objMain.objUtilities.AddAlphaField("@TNX_Warehouse", "BarCode", "Barcode", 30)
        objMain.objUtilities.AddAlphaField("@TNX_Warehouse", "Series", "Series", 30)
        objMain.objUtilities.AddFloatField("@TNX_Warehouse", "Quantity", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_Warehouse", "Branch", "From Branch", 50)
        objMain.objUtilities.AddAlphaField("@TNX_Warehouse", "UOM", "Uom Code", 50)

        objMain.objUtilities.AddDateField("@TNX_Warehouse", "DocDate", "Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_Warehouse", "BEntry", "Base Entry", 50)
        objMain.objUtilities.AddAlphaField("@TNX_Warehouse", "BaseNum", "Base Number", 50)


        objMain.objUtilities.CreateTable("TNX_Warehouse_C0", "Warehouse Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_Warehouse_C0", "ToWhse", "To Whse", 50)
        objMain.objUtilities.AddFloatField("@TNX_Warehouse_C0", "Quantity", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_Warehouse_C0", "Remark", "Remark", 50)
        objMain.objUtilities.AddAlphaField("@TNX_Warehouse_C0", "STR", "STR No", 50)



    End Sub
    Sub CreatMultipleStockTransferRequest()

        objMain.objUtilities.CreateTable("TNX_MPSTR", "Multiple STR", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_MPSTR", "MEFWHSCD", "FromWarehouseCode", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR", "MEFWHSNM", "WarehouseName", 100)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR", "MEFBPID", "FromBranch", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR", "METWHSCD", "ToWarehouseCode", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR", "METWHSNM", "WarehouseName", 100)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR", "METBPID", "ToBranch", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR", "MESRS", "DocSeries", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR", "MESTS", "DocStatus", 30)
        objMain.objUtilities.AddDateField("@TNX_MPSTR", "MEPSDT", "PostingDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_MPSTR", "MEDOCDT", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR", "MERMKS", "Remarks", 254)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR", "PDTUSER", "Pdt User", 20)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR", "UserName", "UserName", 100)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR", "REMARKS", "REMARKS", 254)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR", "VCode", "Vendorcode", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR", "GRPNO", "GRPONumber", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR", "GRPDM", "GRPONum", 30)


        objMain.objUtilities.CreateTable("TNX_MPSTR_C0", "Multiple STR Child1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR_C0", "MEBRCD", "Barcode", 100)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR_C0", "MEITMCD", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR_C0", "MEITNM", "Item Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR_C0", "MEUOM", "UOM", 100)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR_C0", "MEADBRCD", "Addl Barcode", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR_C0", "MEFWSTK", "From WH SOH", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR_C0", "METWHSTK", "To WH SOH", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR_C0", "MEFWHSCD", "FromWarehouse", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR_C0", "METWHSCD", "ToWarehouse", 50)
        objMain.objUtilities.AddFloatField("@TNX_MPSTR_C0", "MEREQQTY", "Required Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_MPSTR_C0", "MERELQTY", "Dispatched Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_MPSTR_C0", "MEOPNQTY", "Open Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_MPSTR_C0", "MERCEQTY", "Received Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR_C0", "MELNSTS", "LineStatus", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR_C0", "MEPCKNG", "Packing", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR_C0", "GENT", "GRPO Entry", 100)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR_C0", "GLNE", "GRPO Line", 100)


        objMain.objUtilities.CreateTable("TNX_MPSTR_C1", "Multiple STR Child2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_MPSTR_C1", "MEATCNM", "Attachment Name", 100)
        objMain.objUtilities.AddLinkField("@TNX_MPSTR_C1", "MEATCH", "Attachment", 254, SAPbobsCOM.BoFldSubTypes.st_Link)
    End Sub


#End Region

#Region "SQL Scripts, Queries, Authorizations"
    Sub RunSQLScripts()
        Try
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("Barcode_ItemName.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("Barcode_Price.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("Barcode_Vat.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("Barcode_WithoutVatPrice.sql"))

            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_ApprovalStatusDetailReport.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_ApprovalStatusReport.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_ITEMDETAILS_NEW.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("GET_RSP.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("GET_COST.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("GetRetailPrice.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_LOAD_CENTRAL_PCR.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("SBO_AverageSale.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("SBO_RebateAmount.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("UPDATE_PRICECHANGES.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("SP_TNX_PRICE_UPDATE.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("GetItemPrice.sql"))

            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("Barcode_ItemNameMadhu.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("Barcode_PriceMadhu.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("offerprice.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_BULKPRICEUPDATE.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_SUPPLIER_REBATE_WIZARDLOADGRID.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("GetDetails.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("SBO_LOADSTOCK.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("GET_HOLDING_DAYS.sql"))

            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("Check_GRPO_Price_Fluctuation.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("Check_GR_Price_Fluctuation.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("Check_GRPO_Price_Fluctuation_APCredit.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("SBO_SP_PostTransactionNotice.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_LASTGRPO_PRICE.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("PriceList_TextFile_Generation_20251017.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("GET_ITEM_PRICING_DETAILS.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_ITEMLLIST.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("SBO_GET_TNX_ITEM_MASTER_UPLOAD.sql"))





            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_AllItemsList.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_BarcodeCheck.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_GRPODocumentWiseReport.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_GRPOItemWiseReport.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_MainWHSNearExpiryDtStock.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_OpenDispatchReport.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_TransferReceiptReport.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_StockCuntingSummrytobeposted.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_STOCKBYWAREHOUSE.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_APInvoicesItemWiseReport.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_CategorywiseInventoryOB.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_ExciseItemsReportforRebate.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_ARInvoiceDocWiseReport.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_ARInvoiceItemWiseReport.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_RETAIL_TN.sql"))

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub QueryManager()
        Try
            '  objUtilities.AddOrUpdateQuery(objMain.objCompany, "AR Invoice Document wise Report", -1, objUtilities.LoadEmbeddedSQL("ARInvoiceDocumentwiseReport.sql"))
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Al Bada STR Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_AlBadaaSTRAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Al Muraqabat 1 STR Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_AlMuraqabat1STRAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Al Muraqabat 2 STR Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_AlMuraqabat2STRAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Al Muraqabat 3 STR Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_AlMuraqabat3STRAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Al Rigga 1 STR Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_AlRigga1STRAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Al Rigga 2 STR Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_AlRigga2STRAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Al Rigga 3 STR Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_AlRigga3STRAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Al Rigga 4 STR Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_AlRigga4STRAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "BFH Expiry in 60 days", -1, objUtilities.LoadEmbeddedSQL("TNX_BFHExpiryin60days.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Below cost price in GRPO", -1, objUtilities.LoadEmbeddedSQL("TNX_BelowcostpriceinGRPO.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Cheque Alert on due date", -1, objUtilities.LoadEmbeddedSQL("TNX_ChequeAlertonduedate.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Diyafah 1 STR Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_Diyafah1STRAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Diyafah 2 STR Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_Diyafah2STRAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Emirates ID Expires in 60 Days", -1, objUtilities.LoadEmbeddedSQL("TNX_EmiratesIDExpiresin60Days.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Insurance Expiry in 60 days", -1, objUtilities.LoadEmbeddedSQL("TNX_InsuranceExpiryin60days.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Main Warehouse STR Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_MainWarehouseSTRAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Minimum Qty Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_MinimumQtyAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "OHC Expiry in 60 days", -1, objUtilities.LoadEmbeddedSQL("TNX_OHCExpiryin60days.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Passport Expiry in 60 days", -1, objUtilities.LoadEmbeddedSQL("TNX_PassportExpiryin60days.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Purchase Request Added Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_PurchaseRequestAddedAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Salahuddin STR Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_SalahuddinSTRAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Sales Order Added alert", -1, objUtilities.LoadEmbeddedSQL("TNX_SalesOrderAddedalert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Supplier Rebate Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_SupplierRebateAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Tomorrow Deliverys Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_TomorrowDeliverysAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Visa Expiry in 60 days", -1, objUtilities.LoadEmbeddedSQL("TNX_VisaExpiryin60days.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Warehouse PO Created Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_WarehousePOCreatedAlert.sql"), "Alert")
            'Cart
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "BTF Alert HO", -1, objUtilities.LoadEmbeddedSQL("TNX_BTFAlertHO.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Exchange Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_ExchangeAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Credit Limit Exceeded Vendors", -1, objUtilities.LoadEmbeddedSQL("TNX_CreditLimitExceededVendors.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Approval greater than 2000", -1, objUtilities.LoadEmbeddedSQL("TNX_Approvalgreaterthan2000.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "JE for Inventory GL", -1, objUtilities.LoadEmbeddedSQL("TNX_JEforInventoryGL.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "JE for Non Inventory GL", -1, objUtilities.LoadEmbeddedSQL("TNX_JEforNonInventoryGL.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "RSP Price Change Query", -1, objUtilities.LoadEmbeddedSQL("TNX_RSPPriceChangeQuery.sql"), "Alert")
            'Kenz
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Journal Voucher ApprovalV", -1, objUtilities.LoadEmbeddedSQL("TNX_JournalVoucherApprovalV.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "OutGoingPayment Alert vino", -1, objUtilities.LoadEmbeddedSQL("TNX_OutGoingPaymentAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "More than setup barcodes", -1, objUtilities.LoadEmbeddedSQL("TNX_Morethansetupbarcodes.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "UomPrices Query", -1, objUtilities.LoadEmbeddedSQL("TNX_UomPricesQuery.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Records Without Mobile Number", -1, objUtilities.LoadEmbeddedSQL("TNX_RecordsWithoutMobileNumber.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "SPP2 RowNumbers", -1, objUtilities.LoadEmbeddedSQL("TNX_SPP2RowNumbers.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "SPP1 RowNumbers", -1, objUtilities.LoadEmbeddedSQL("TNX_SPP1RowNumbers.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "UOM Entry", -1, objUtilities.LoadEmbeddedSQL("TNX_UOMEntry.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Promotion Near Expiry Items", -1, objUtilities.LoadEmbeddedSQL("TNX_PromotionNearExpiryItems.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Future and Back Dated Entries Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_FutureandBackDatedEntriesAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Open Purchase Request Lines", -1, objUtilities.LoadEmbeddedSQL("TNX_OpenPurchaseRequestLines.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "cco line numbers", -1, objUtilities.LoadEmbeddedSQL("TNX_ccolinenumbers.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "ITM2 Vendor Row Numbers", -1, objUtilities.LoadEmbeddedSQL("TNX_ITM2VendorRowNumbers.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Sales Report for Rectification", -1, objUtilities.LoadEmbeddedSQL("TNX_SalesReportforRectification.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Open Purchase Order Lines", -1, objUtilities.LoadEmbeddedSQL("TNX_OpenPurchaseOrderLines.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Minimum Quantity Alert_New", -1, objUtilities.LoadEmbeddedSQL("TNX_MinimumQuantityAlert_New.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Minimum Quantity Alert_New_V1", -1, objUtilities.LoadEmbeddedSQL("TNX_MinimumQuantityAlert_New_V1.sql"), "Alert")

            'Saudia
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Passport Expiry", -1, objUtilities.LoadEmbeddedSQL("TNX_PassportExpiry.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "QID Expiry", -1, objUtilities.LoadEmbeddedSQL("TNX_QIDExpiry.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Contract Expiry", -1, objUtilities.LoadEmbeddedSQL("TNX_ContractExpiry.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Birthday", -1, objUtilities.LoadEmbeddedSQL("TNX_Birthday.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Medical Card Expiry", -1, objUtilities.LoadEmbeddedSQL("TNX_MedicalCardExpiry.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "UOM Group By line value", -1, objUtilities.LoadEmbeddedSQL("TNX_UOMGroupBylinevalue.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Goods issue posted without MI", -1, objUtilities.LoadEmbeddedSQL("TNX_GoodsissuepostedwithoutMI.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "LOCKED USER", -1, objUtilities.LoadEmbeddedSQL("TNX_LOCKEDUSER.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Numbering seires", -1, objUtilities.LoadEmbeddedSQL("TNX_Numberingseries.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "BackdateGRPO", -1, objUtilities.LoadEmbeddedSQL("TNX_BackdateGRPO.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Below Cost GRPO Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_BelowCostGRPOAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "SPP1 LineNumber", -1, objUtilities.LoadEmbeddedSQL("TNX_SPP1LineNumber.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "SPP2 LineNumber", -1, objUtilities.LoadEmbeddedSQL("TNX_SPP2LineNumber.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "MissingAttachmentsinBP", -1, objUtilities.LoadEmbeddedSQL("TNX_MissingAttachmentsinBP.sql"), "Alert")

            'Seven Emirates
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Below Cost GRPO Added", -1, objUtilities.LoadEmbeddedSQL("TNX_BelowCostGRPOAdded.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Below Cost PO Added", -1, objUtilities.LoadEmbeddedSQL("TNX_BelowCostPOAdded.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Emirates ID Expiry in next 90 days", -1, objUtilities.LoadEmbeddedSQL("TNX_EmiratesIDExpiryinnext90days.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Goods Receipt Note Added Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_GoodsReceiptNoteAddedAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Goods Receipt Note_Electra Added Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_GoodsReceiptNote_ElectraAddedAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Goods Receipt Note_KHD Added Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_GoodsReceiptNote_KHDAddedAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Goods Receipt Note_TCA Added Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_GoodsReceiptNote_TCAAddedAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Goods Receipt Note_Warehouse Added Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_GoodsReceiptNote_WarehouseAddedAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Insurance Expiry in next 90 days", -1, objUtilities.LoadEmbeddedSQL("TNX_InsuranceExpiryinnext90days.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Labour Card Expiry in next 90 days", -1, objUtilities.LoadEmbeddedSQL("TNX_LabourCardExpiryinnext90days.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "License Expiry in next 90 days", -1, objUtilities.LoadEmbeddedSQL("TNX_LicenseExpiryinnext90days.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Passport Expiry in next 90 days", -1, objUtilities.LoadEmbeddedSQL("TNX_PassportExpiryinnext90days.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Purchase Order Added Alert", -1, objUtilities.LoadEmbeddedSQL("TNX_PurchaseOrderAddedAlert.sql"), "Alert")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Visa Expiry in next 90 days", -1, objUtilities.LoadEmbeddedSQL("TNX_VisaExpiryinnext90days.sql"), "Alert")




            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Goods Issue Approval", -1, objUtilities.LoadEmbeddedSQL("TNX_GoodsIssueApproval.sql"), "GoodsIssue")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Item quantity deviates from Maximum Qty", -1, objUtilities.LoadEmbeddedSQL("TNX_ItemquantitydeviatesfromMaximumQty.sql"), "ItemQuantityPO")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "PO Price more than Last Purchase price", -1, objUtilities.LoadEmbeddedSQL("TNX_POPricemorethanLastPurchaseprice.sql"), "POPrice")
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Sales Oder Credit Limit Deviation", -1, objUtilities.LoadEmbeddedSQL("TNX_SalesOderCreditLimitDeviation.sql"), "SalesOrder")
            'AL JAZEERA
            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Ap Credit Memo ", -1, objUtilities.LoadEmbeddedSQL("SP_SBO_APCM_ALERT.sql"), "Alert")

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub AddUserFormAuthorizations()
        Try

            objUtilities.AddAuthorization("ME_ACFG_Form", "Retail Configuration", "ME_ACFG_Form", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("SDPTR", "Supplier Deposit", "SDPTR", "ME_ACFG_Form", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)

            objUtilities.AddAuthorization("License", "Tenx License Administration", "License", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)

            objUtilities.AddAuthorization("ITEM", "Material Promotion Agreements", "ITEM", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("LEASTITM", "Buy-X-Get-Y Incentives", "LEASTITM", "ITEM", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("PRMONVP", "Voucher Campaign Management", "PRMONVP", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("PRMONTOTL", "Total Invoice Discount", "PRMONTOTL", "PRMONVP", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)

            objUtilities.AddAuthorization("REDCTN", "Supplier Rebate Agreements", "REDCTN", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("ME_REVERSION", "Rebate Settlement and Reversion", "ME_REVERSION", "REDCTN", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)

            objUtilities.AddAuthorization("frmBarCd", "Barcode Label Print", "frmBarCd", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("frmScale", "Scale Interface Setup", "frmScale", "frmBarCd", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)

            objUtilities.AddAuthorization("ME_PSC_Form", "Mobile Stock Counting", "ME_PSC_Form", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)

            objUtilities.AddAuthorization("ME_STR_Form", "Inter-Store Transfer Requests", "ME_STR_Form", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)

            objUtilities.AddAuthorization("TNXCRDNT_Form", "Inventory Returns and Credit", "TNXCRDNT_Form", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)

            objUtilities.AddAuthorization("PriceList", "Mass Price Update", "PriceList", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)

            objUtilities.AddAuthorization("ME_AP_Form", "Automatic Production Execution", "ME_AP_Form", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("Clear", "Promotional Clearance Runs", "Clear", "ME_AP_Form", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)

            objUtilities.AddAuthorization("ME_GPOFPRQ_Form", "PO Generation Wizard", "ME_GPOFPRQ_Form", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("ME_BPO_Form", "Purchase Order Generation Wizard", "ME_BPO_Form", "ME_GPOFPRQ_Form", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("ME_PMAPINV_Form", "Mass AP Invoice Posting Wizard", "ME_PMAPINV_Form", "ME_GPOFPRQ_Form", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("ME_PMAPCM_Form", "AP Credit Memo Wizard", "ME_PMAPCM_Form", "ME_GPOFPRQ_Form", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("ME_SCGW", "Goods Receipt and Count Generation Wizard", "ME_SCGW", "ME_GPOFPRQ_Form", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)

            objUtilities.AddAuthorization("SBO_PGW_Form", "Production Order Wizard", "SBO_PGW_Form", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)

            objUtilities.AddAuthorization("ME_PDC_Form", "Supplier Deposit Management Wizard", "ME_PDC_Form", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)

            objUtilities.AddAuthorization("APINR", "AP Invoices Itemwise Report", "APINR", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("CATWIN", "Category wise Inventory OB", "CATWIN", "APINR", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("EXITM", "Excise Items Report for Rebate", "EXITM", "APINR", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)

            objUtilities.AddAuthorization("ALLITMS", "All Items List", "ALLITMS", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("BARCHK", "Barcode Check", "BARCHK", "ALLITMS", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("GRPODWR", "GRPO Document Wise Report", "GRPODWR", "ALLITMS", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("GRPOITM", "GRPO Item Wise Report", "GRPOITM", "ALLITMS", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("MWNEDSR", "Main Warehouse Near Expiry Date Stock Report", "MWNEDSR", "ALLITMS", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("OPDSPR", "Open Dispatch Report", "OPDSPR", "ALLITMS", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("TRNSRP", "Transfer Receipt Report", "TRNSRP", "ALLITMS", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("STCKCNT", "Stock Counting Summary to be Posted", "STCKCNT", "ALLITMS", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("STKWHSE", "Stock Report by Warehouse", "STKWHSE", "ALLITMS", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)


            objUtilities.AddAuthorization("ARINR", "AR Invoice Document wise Report", "ARINR", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("ARITMR", "AR Invoice Itemwise Report", "ARITMR", "ARINR", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)



        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub AddApproval()
        Try
            objUtilities.CreateApprovalTemplateWithQueryTerm("GoodsIssue", "Goods Issue Approval", 1,
                SAPbobsCOM.ApprovalTemplatesDocumentTypeEnum.atdtGoodsIssue, 1, 20)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    'Alerts Notification
    Sub AddAlerts()
        Try
            ' QueryID = 360 (saved in Query Manager)- OUQR table
            ' Frequency = every 30 minutes
            'Start after 1 minute
            objUtilities.CreateCustomAlert("AP Credit Memo Alert", 361, 1, 1, AlertManagementFrequencyType.atfi_Days)
            objMain.objApplication.StatusBar.SetText("Alert created successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


#End Region

End Class
