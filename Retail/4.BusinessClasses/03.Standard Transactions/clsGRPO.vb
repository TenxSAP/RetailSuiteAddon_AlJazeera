Imports System.Text.RegularExpressions
Imports SAPbobsCOM

Public Class clsGRPO

#Region "Declaration"
    Public objForm As SAPbouiCOM.Form
    Dim objMatrix, objMatrix1 As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Details As SAPbouiCOM.DBDataSource
    Dim oitem, oItem1, oItem2, oItem3, oItem4 As SAPbouiCOM.Item
    Dim result As Integer
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim BaseUom As String
    Dim CalcualtionBtnFlag As Boolean = False
    Dim DirectGRPOFlag As Boolean = True
    Dim SplitCombo As SAPbouiCOM.ComboBox
    Dim SplitFlgCombo As SAPbouiCOM.ComboBox
    Dim DiscPrcnt As Double = 0
    Dim DiscUnitPrice As Double = 0
    Dim GlobalDocnum As String = 0
    Private LastDiscRow As Integer = -1
    Private LastDiscValue As String = ""
    Dim NST As String = ""
#End Region

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByVal GlobalDocnum As String)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("OPDN")
                    objMatrix = objForm.Items.Item("38").Specific
                    If pVal.ItemUID = "38" And pVal.ColUID = "U_BelowCst" And pVal.BeforeAction = True Then
                        objMain.objApplication.StatusBar.SetText("User not authorized to performe this operation...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        BubbleEvent = False
                    End If
                    If pVal.ItemUID = "38" And pVal.ColUID = "U_RetailPrice" And pVal.BeforeAction = True Then
                        objMain.objApplication.StatusBar.SetText("User not authorized to performe this operation...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        BubbleEvent = False
                    End If



                    If pVal.ItemUID = "1" And pVal.Before_Action = True Then
                        Dim RetailPrice As Double = 0
                        Dim SupplierPrice As Double = 0
                        Dim SupplierPriceStr As String = ""
                        Dim oCmbBelowCost As SAPbouiCOM.ComboBox




                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            For i As Integer = 1 To objMatrix.RowCount
                                SplitCombo = objMatrix.Columns.Item("U_SplitKitItems").Cells.Item(i).Specific
                                SplitFlgCombo = objMatrix.Columns.Item("U_SPLITTED").Cells.Item(i).Specific
                                If SplitCombo.Selected.Value = SplitFlgCombo.Selected.Value Then

                                Else
                                    objMain.objApplication.StatusBar.SetText("Validation : Split Item Not Splitted  for Line ID : " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                End If
                            Next
                        End If


                    End If

                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("OPDN")
                    objMatrix = objForm.Items.Item("38").Specific
                    If pVal.ItemUID = "38" And pVal.ColUID = "U_BelowCst" And pVal.BeforeAction = True Then
                        objMain.objApplication.StatusBar.SetText("User not authorized to performe this operation...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        BubbleEvent = False
                    End If

                    If pVal.ItemUID = "38" And pVal.BeforeAction = False And pVal.CharPressed = "40" Then
                        Try
                            If CStr(objMatrix.Columns.Item("4").Cells.Item(pVal.Row + 1).Specific.Value).Trim = "" Then
                                objMatrix.Columns.Item("4").Cells.Item(pVal.Row + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                            End If
                        Catch ex As Exception

                        End Try
                    End If

                    If pVal.ItemUID = "38" And pVal.ColUID = "U_RetailPrice" And pVal.BeforeAction = True Then
                        objMain.objApplication.StatusBar.SetText("User not authorized to performe this operation...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        BubbleEvent = False
                    End If
                    'Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    '    If pVal.ItemUID = "10000330" And pVal.BeforeAction = False Then
                    '        Dim objComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("10000330").Specific
                    '        Dim Doc1 As String = objComboBox.Selected.Description
                    '        If Doc1 = "22" Then
                    '            'Dim objBaseForm As SAPbouiCOM.Form
                    '            'Dim Purchasenumber As String
                    '            'objBaseForm = objMain.objApplication.Forms.GetForm("10021", objMain.objApplication.Forms.ActiveForm.TypeCount)

                    '            'objMatrix1 = objBaseForm.Items.Item("7").Specific

                    '            'Dim line As Integer = objMatrix1.VisualRowCount
                    '            'Purchasenumber = objMatrix1.Columns.Item("DocNum").Cells.Item(line).Specific.Value

                    '            'Dim oRsBranch As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '            'Dim Query As String = " Select T1.""Quantity"",T1.""OpenQty"",(T1.""Quantity""- T1.""OpenQty"")as ""Delieverd Quantity"",T1.""LineNum"",T1.""ItemCode"" From OPOR T0  Inner Join " &
                    '            '                                                    "POR1 T1 On T1.""DocEntry"" = T0.""DocEntry"" Where T0.""DocNum"" = '" & Purchasenumber & "'and T1.""LineStatus""='O' "
                    '            'oRsBranch.DoQuery(Query)
                    '            'objMatrix = objForm.Items.Item("38").Specific
                    '            'If oRsBranch.RecordCount > 0 Then

                    '            '    For k As Integer = 1 To oRsBranch.RecordCount
                    '            '        Dim Deliver As String = oRsBranch.Fields.Item("Delieverd Quantity").Value
                    '            '        If oRsBranch.Fields.Item("Delieverd Quantity").Value <> oRsBranch.Fields.Item("Quantity").Value Then
                    '            '            objMatrix.Columns.Item("U_Delivr").Cells.Item(k).Specific.Value = Deliver
                    '            '        End If
                    '            '        oRsBranch.MoveNext()


                    '            '    Next
                    '        End If
                    '    End If


                    '    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("OPDN")
                    objMatrix = objForm.Items.Item("38").Specific

                    'If pVal.ItemUID = "btn_MRG" And pVal.BeforeAction = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                    '    Try

                    '        For i As Integer = 1 To objMatrix.VisualRowCount

                    '            Dim ISGROSS As String = ""
                    '            If CStr(objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value).Trim <> "" Then

                    '                Dim VatGrpRate As String = "SELECT (T1.""Rate""+100)/100 FROM OITM T0  INNER JOIN OVTG T1 ON T0.""VatGourpSa"" = T1.""Code"" Where T0.""ItemCode""='" & CStr(objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value).Trim & "'"
                    '                Dim oRsVatGrpRate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                oRsVatGrpRate.DoQuery(VatGrpRate)

                    '                Dim Amm As Double = 0


                    '                If CStr(oRsVatGrpRate.Fields.Item(0).Value).Trim <> "" Then
                    '                    Amm = CDbl(CStr(oRsVatGrpRate.Fields.Item(0).Value).Trim)
                    '                End If

                    '                objComboBox = objForm.Items.Item("2001").Specific
                    '                Dim Branch As Integer = objComboBox.Value
                    '                Dim GROSS As String = "SELECT IFNULL(X.""IsGrossPrc"",'N') FROM ""OBPL"" M INNER  JOIN ""OPLN"" X ON M.""U_PricelistNo""=""ListNum"" WHERE ""BPLId""='" & Branch & "'"
                    '                Dim RSGROSS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                RSGROSS.DoQuery(GROSS)

                    '                If CStr(RSGROSS.Fields.Item(0).Value).Trim <> "" Then
                    '                    ISGROSS = (CStr(RSGROSS.Fields.Item(0).Value).Trim)
                    '                End If


                    '                'Dim NetCost As Double = 0

                    '                'If objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value = "" Then
                    '                '    NetCost = 0
                    '                'Else
                    '                '    NST = objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value
                    '                '    NST = NST.Replace(objForm.Items.Item("63").Specific.Value, "")

                    '                '    'NetCost = objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value
                    '                '    'NetCost = NetCost.Replace(objForm.Items.Item("63").Specific.Value, "")
                    '                '    NetCost = Convert.ToDouble(NST.Trim)
                    '                'End If

                    '                'DiscPrcnt = 0
                    '                'DiscPrcnt = Convert.ToDouble(objMatrix.Columns.Item("15").Cells.Item(i).Specific.Value)
                    '                'If DiscPrcnt > 0 Then
                    '                '    NST = objMatrix.Columns.Item("15").Cells.Item(i).Specific.Value
                    '                '    NST = NST.Replace(objForm.Items.Item("63").Specific.Value, "")
                    '                '    NetCost = Convert.ToDouble(NST.Trim)
                    '                'End If

                    '                'Dim RetailPrice As Double = objMatrix.Columns.Item("U_RetailPrice").Cells.Item(i).Specific.Value

                    '                'If RetailPrice > 0 And NetCost > 0 Then
                    '                '    If NetCost > RetailPrice Then
                    '                '        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_BelowCst")
                    '                '        oColumn.Cells.Item(i).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    '                '    Else
                    '                '        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_BelowCst")
                    '                '        oColumn.Cells.Item(i).Specific.Select("No", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    '                '    End If
                    '                'End If
                    '                Dim PriAfterDis As Double = 0
                    '                Dim pattern As String = "\d+(\.\d{1,4})?"

                    '                Dim match As Match = Regex.Match(CStr(objMatrix.Columns.Item("17").Cells.Item(i).Specific.Value).Trim, pattern)


                    '                If CStr(objMatrix.Columns.Item("17").Cells.Item(i).Specific.Value).Trim <> "" Then
                    '                    PriAfterDis = CDbl(match.Value)
                    '                End If

                    '                Dim RetailPrice1 As Double = 0
                    '                Dim SuppPrice As Double = 0


                    '                If CStr(objMatrix.Columns.Item("U_RetailPrice").Cells.Item(i).Specific.Value).Trim <> "" And ISGROSS = "Y" Then
                    '                    RetailPrice1 = CDbl(CStr(objMatrix.Columns.Item("U_RetailPrice").Cells.Item(i).Specific.Value).Trim)
                    '                    RetailPrice1 = RetailPrice1 / Amm
                    '                    RetailPrice1 = Math.Round(RetailPrice1, 4)
                    '                ElseIf CStr(objMatrix.Columns.Item("U_RetailPrice").Cells.Item(i).Specific.Value).Trim <> "" And ISGROSS = "N" Then
                    '                    RetailPrice1 = CDbl(CStr(objMatrix.Columns.Item("U_RetailPrice").Cells.Item(i).Specific.Value).Trim)
                    '                Else
                    '                    RetailPrice1 = CDbl(CStr(objMatrix.Columns.Item("U_RetailPrice").Cells.Item(i).Specific.Value).Trim)
                    '                End If

                    '                If PriAfterDis > RetailPrice1 Then
                    '                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_BelowCst")
                    '                    oColumn.Cells.Item(i).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    '                Else
                    '                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_BelowCst")
                    '                    oColumn.Cells.Item(i).Specific.Select("No", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    '                End If

                    '                Try



                    '                    'Dim pattern As String = "\d+(\.\d{1,4})?"

                    '                    'Dim match As Match = Regex.Match(CStr(objMatrix.Columns.Item("17").Cells.Item(i).Specific.Value).Trim, pattern)


                    '                    If CStr(objMatrix.Columns.Item("17").Cells.Item(i).Specific.Value).Trim <> "" Then
                    '                        SuppPrice = CDbl(match.Value)
                    '                    End If

                    '                    objMatrix.Columns.Item("U_MARGN").Cells.Item(i).Specific.Value = ((RetailPrice1 - SuppPrice) / RetailPrice1) * 100

                    '                Catch ex As Exception

                    '                End Try

                    '            End If
                    '        Next
                    '    Catch ex As Exception
                    '        objMain.objApplication.StatusBar.SetText(ex.Message)
                    '    End Try
                    'End If

                    If pVal.ItemUID = "btn_ED" And pVal.BeforeAction = False Then
                        Try
                            objForm.Freeze(True)




                            'Dim GetData As String = "SELECT T1.""LineNum"",T1.""ItemCode"", T1.""Dscription"", T1.""Quantity"", T1.""CodeBars"", " &
                            '                        "T1.""UomCode"" FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.""DocEntry"" = T1.""DocEntry""  " &
                            '                        " INNER JOIN OWHS T2 ON T1.""WhsCode"" = T2.""WhsCode"" WHERE  " &
                            '                        "T0.""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0) & "' And IFNULL(T2.""U_EXCHK"",'No') ='Yes' And T0.""DocType""='I' ORDER BY T1.""LineNum"""

                            'Dim oRsGetData As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            'oRsGetData.DoQuery(GetData)

                            'If oRsGetData.RecordCount > 0 Then
                            '    objMain.objExpiry.CreateForm(objForm.UniqueID, "OPDN", oDBs_Head.GetValue("DocEntry", 0), oDBs_Head.GetValue("DocNum", 0), oRsGetData)
                            'Else
                            '    objMain.objApplication.StatusBar.SetText("Expiry Date screen not applicable")
                            'End If

                            'Dim oGUID As SAPbouiCOM.Form = objMain.objApplication.Forms.GetForm("-143", 0)

                            'Dim oDbH As SAPbouiCOM.DBDataSource = oGUID.DataSources.DBDataSources.Item("OPDN")


                            If CStr(objForm.Items.Item("3").Specific.Value).Trim <> "I" Then
                                objMain.objApplication.StatusBar.SetText("Expiry Screen not applicable")
                                objForm.Freeze(False)
                                Exit Try
                            End If

                            Dim GetWCnd As String = "select IFNULL(""U_EXCHK"",'No') FROM OWHS Where ""WhsCode""='" & objMatrix.Columns.Item("24").Cells.Item(1).Specific.Value & "' And ""U_EXCHK""='Yes' And ""U_EXCHK"" IS NOT NULL And ""U_EXCHK""<>''"
                            Dim oRsGetWCnd As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetWCnd.DoQuery(GetWCnd)

                            If oRsGetWCnd.RecordCount = 0 Then
                                objMain.objApplication.StatusBar.SetText("Expiry Screen not applicable")
                                objForm.Freeze(False)
                                Exit Try
                            End If


                            Dim dt As New DataTable()
                            Dim UID As String = String.Empty

                            If CStr(objForm.Items.Item("txt_GUID").Specific.Value).Trim <> "" Then
                                UID = objForm.Items.Item("txt_GUID").Specific.Value
                                'GoTo GotoLst
                            Else

                                UID = System.Guid.NewGuid.ToString()
                                objForm.Items.Item("txt_GUID").Specific.Value = UID.Trim
                            End If

                            ' Define columns for the DataTable
                            dt.Columns.Add("LineNum", GetType(Integer))
                            dt.Columns.Add("ItemCode", GetType(String))
                            dt.Columns.Add("Dscription", GetType(String))
                            dt.Columns.Add("Quantity", GetType(Double))
                            dt.Columns.Add("CodeBars", GetType(String))
                            dt.Columns.Add("UomCode", GetType(String))
                            dt.Columns.Add("BQuantity", GetType(Double))


                            For i As Integer = 1 To objMatrix.VisualRowCount

                                If CStr(objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value).Trim <> "" Then

                                    Dim dr As DataRow = dt.NewRow
                                    dr("LineNum") = i - 1
                                    dr("ItemCode") = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value
                                    dr("Dscription") = objMatrix.Columns.Item("3").Cells.Item(i).Specific.Value
                                    dr("Quantity") = objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value
                                    dr("CodeBars") = objMatrix.Columns.Item("4").Cells.Item(i).Specific.Value
                                    dr("UomCode") = objMatrix.Columns.Item("1470002145").Cells.Item(i).Specific.Value
                                    dr("BQuantity") = objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value

                                    dt.Rows.Add(dr)

                                End If



                            Next

GotoLst:                    objMain.objExpiry.CreateFormNew(objForm.UniqueID, "OPDN", oDBs_Head.GetValue("DocEntry", 0), oDBs_Head.GetValue("DocNum", 0), dt, UID.Trim, objForm.Mode)


                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If

                    If pVal.ItemUID = "Barcode" And pVal.BeforeAction = True And objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        objMain.objBarcode.CreateForm(pVal.FormTypeCount, "GRPO", "BarCode - GRPO.rpt", objForm.Items.Item("8").Specific.Value.ToString().Trim())
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then


                        If Me.Validation(objForm.UniqueID) = False Then BubbleEvent = False

                    End If

                    If pVal.ItemUID = "Split" And pVal.BeforeAction = True And objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Try
                            result = objMain.objApplication.MessageBox("You are going to split the items. It will post goods issue and goods receipt documents.", 1, "Ok", "Cancel")
                            If (result = 1) Then
                                If objForm.Items.Item("81").Specific.Value.ToString.Trim() = "4" Then
                                    objMain.objApplication.StatusBar.SetText("Document is canceled, you can not split the items", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Else
                                    objMain.objApplication.StatusBar.SetText("Please wait, splitting the items.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    Me.GoodsIssueAndReceipt(objForm.UniqueID)
                                End If
                            End If
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText("Error:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If

                    If pVal.ItemUID = "CAL_NETQTY" And pVal.BeforeAction = False Then
                        Try
                            For i As Integer = 1 To objMatrix.VisualRowCount
                                If objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value <> "" Then
                                    Dim qStr As String = "SELECT A.""U_LOOSPER"" FROM OITM A WHERE A.""ItemCode""='" + objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.ToString() + "'"
                                    Dim rSet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    rSet.DoQuery(qStr)
                                    If rSet.RecordCount > 0 Then
                                        Dim GrossQty As Double = CDbl(objMatrix.Columns.Item("U_GROSSQTY").Cells.Item(i).Specific.Value)
                                        Dim GrossTotal As Double = CDbl(objMatrix.Columns.Item("U_GROSSTOT").Cells.Item(i).Specific.Value)
                                        If GrossQty <> 0 And GrossTotal <> 0 Then
                                            Dim LoosPercemtage As Double = CDbl(rSet.Fields.Item("U_LOOSPER").Value)
                                            Dim Quantity As Double = GrossQty * (1 - (LoosPercemtage / 100))
                                            If Quantity = 0 Then Quantity = GrossQty
                                            Dim UnitPrice As Double = GrossTotal / Quantity

                                            objMatrix.Columns.Item("15").Cells.Item(i).Specific.Value = (0).ToString()
                                            objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value = (0).ToString()
                                            '' objMatrix.Columns.Item("21").Cells.Item(i).Specific.Value = GrossTotal.ToString()
                                            objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value = Quantity.ToString()
                                            objMatrix.Columns.Item("21").Cells.Item(i).Specific.Value = GrossTotal.ToString()
                                        Else
                                            objMain.objApplication.StatusBar.SetText("Kindly enter GrossQty and GrossTotal in Line No " + i.ToString() + "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If
                                    End If
                                End If
                            Next
                        Catch ex As Exception
                            objForm.Freeze(False)
                        End Try
                    End If

                    'If pVal.ItemUID = "btn_calFOC" And pVal.BeforeAction = False Then

                    '    Dim PurchaseNum As String = objMatrix.Columns.Item("45").Cells.Item(objMatrix.VisualRowCount - 1).Specific.Value.ToString()
                    '    Dim oRsBranch As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    '    Dim Query As String = " Select T0.""DocEntry"",T0.""DocNum"",T1.""Quantity"",T1.""OpenQty"",(T1.""Quantity""- T1.""OpenQty"")as ""Delieverd Quantity"",T1.""LineNum"",T1.""ItemCode"" From OPOR T0  Inner Join " &
                    '                                                        "POR1 T1 On T1.""DocEntry"" = T0.""DocEntry"" Where CAST(T0.""DocEntry"" As VARCHAR) = '" & PurchaseNum & "'and T1.""LineStatus""='O' "
                    '    oRsBranch.DoQuery(Query)

                    '    For i As Integer = 1 To oRsBranch.RecordCount



                    '        Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.Trim
                    '        Dim SupplierPrice As String = objMatrix.Columns.Item("U_suppprice").Cells.Item(i).Specific.Value.Trim

                    '        If (ItemCode <> "" And SupplierPrice = "0.0") Then
                    '            objMain.objApplication.StatusBar.SetText("For Item " & ItemCode & " and Line No. " & i & "... Supplier Price should not be empty.....!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '            Exit Sub
                    '        End If

                    '        If oRsBranch.RecordCount > 0 Then
                    '            Dim Deliver As String = oRsBranch.Fields.Item("Delieverd Quantity").Value
                    '            If oRsBranch.Fields.Item("Delieverd Quantity").Value <> oRsBranch.Fields.Item("Quantity").Value Then
                    '                objMatrix.Columns.Item("U_Delivr").Cells.Item(i).Specific.Value = Deliver
                    '            End If


                    '            'Dim Normalquantity As Double = objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value.ToString()
                    '            'Dim openquantity As String = oRsBranch.Fields.Item("Delieverd Quantity").Value
                    '            ''If Normalquantity <> openquantity Then
                    '            'Dim quantity As String = Normalquantity + openquantity
                    '            'If quantity <0 Then
                    '            '    objMatrix.Columns.Item("U_Delivr").Cells.Item(i).Specific.Value = -(quantity)
                    '            'Else
                    '            '    objMatrix.Columns.Item("U_Delivr").Cells.Item(i).Specific.Value = (quantity)

                    '            'Dim Normalquantity As Double = objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value.ToString()
                    '            'Dim openquantity As String = oRsBranch.Fields.Item("Delieverd Quantity").Value
                    '            ''If Normalquantity <> openquantity Then
                    '            'Dim quantity As String = Normalquantity + openquantity
                    '            'If quantity <0 Then
                    '            '    objMatrix.Columns.Item("U_Delivr").Cells.Item(i).Specific.Value = -(quantity)
                    '            'Else
                    '            '    objMatrix.Columns.Item("U_Delivr").Cells.Item(i).Specific.Value = (quantity)

                    '            'End If
                    '            'End If
                    '            oRsBranch.MoveNext()
                    '        End If

                    '    Next

                    '    Me.CalculateFOC(objForm.UniqueID)

                    '    ' Me.LineValidations(objForm.UniqueID)

                    '    CalcualtionBtnFlag = True

                    '    CalcualtionBtnFlag = True

                    '    For i As Integer = 1 To objMatrix.VisualRowCount
                    '        Dim IsUomDefault As String = "Yes"
                    '        Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.ToString()
                    '        Dim UoM As String = objMatrix.Columns.Item("1470002145").Cells.Item(i).Specific.Value.ToString()
                    '        Dim BarCode As String = objMatrix.Columns.Item("4").Cells.Item(i).Specific.Value.ToString()
                    '        If UoM <> "" Then
                    '            If Me.CheckDefaultUOM(UoM, ItemCode) = True Then
                    '                IsUomDefault = "Yes"
                    '            Else
                    '                IsUomDefault = "No"
                    '            End If
                    '        End If
                    '        'Me.LoadOnItemOrUOMSelection(objForm.UniqueID, i, IsUomDefault, ItemCode, UoM, objForm.Items.Item("4").Specific.Value)

                    '        '''For Special Prices
                    '        Me.LoadOnItemOrUOMSelection_SpecialPrices(objForm.UniqueID, i, IsUomDefault, ItemCode, UoM, objForm.Items.Item("4").Specific.Value, objForm.Items.Item("10").Specific.Value, BarCode)
                    '    Next

                    'End If

                    'If pVal.ItemUID = "btn_LdItd" And pVal.BeforeAction = False Then
                    '    objMain.objApplication.StatusBar.SetText("Please wait while loading line details....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    '    'Dim PurchaseNum As String = objMatrix.Columns.Item("45").Cells.Item(objMatrix.VisualRowCount - 1).Specific.Value.ToString()

                    '    'Dim oRsBranch As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '    'Dim Query As String = " Select T0.""DocEntry"",T0.""DocNum"",T1.""Quantity"",T1.""OpenQty"",(T1.""Quantity""- T1.""OpenQty"")as ""Delieverd Quantity"",T1.""LineNum"",T1.""ItemCode"" From OPOR T0  Inner Join " &
                    '    '                                                    "POR1 T1 On T1.""DocEntry"" = T0.""DocEntry"" Where T0.""DocEntry"" = '" & PurchaseNum & "'and T1.""LineStatus""='O' "
                    '    'oRsBranch.DoQuery(Query)
                    '    'For i As Integer = 1 To oRsBranch.RecordCount
                    '    '    Dim IsUomDefault As String = "Yes"
                    '    '    Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.ToString()
                    '    '    Dim UoM As String = objMatrix.Columns.Item("1470002145").Cells.Item(i).Specific.Value.ToString()
                    '    '    If UoM <> "" Then
                    '    '        If Me.CheckDefaultUOM(UoM, ItemCode) = True Then
                    '    '            IsUomDefault = "Yes"
                    '    '        Else
                    '    '            IsUomDefault = "No"
                    '    '        End If
                    '    '    End If
                    '    '    Me.LoadOnItemOrUOMSelection(objForm.UniqueID, i, IsUomDefault, ItemCode, UoM, objForm.Items.Item("4").Specific.Value)

                    '    '    'If oRsBranch.RecordCount > 0 Then
                    '    '    'For k As Integer = 1 To oRsBranch.RecordCount
                    '    '    Dim Deliver As String = oRsBranch.Fields.Item("Delieverd Quantity").Value
                    '    '    If oRsBranch.Fields.Item("Delieverd Quantity").Value <> oRsBranch.Fields.Item("Quantity").Value Then
                    '    '        objMatrix.Columns.Item("U_Delivr").Cells.Item(i).Specific.Value = Deliver
                    '    '    End If

                    '    '    Dim Normalquantity As String = objMatrix.Columns.Item("U_normalqty").Cells.Item(i).Specific.Value.ToString()
                    '    '    Dim openquantity As String = oRsBranch.Fields.Item("OpenQty").Value
                    '    '    If Normalquantity <> openquantity Then
                    '    '        Dim quantity As String = Normalquantity - openquantity
                    '    '        If quantity < 0 Then
                    '    '            objMatrix.Columns.Item("U_Delivr").Cells.Item(i).Specific.Value = -(quantity)
                    '    '        Else
                    '    '            objMatrix.Columns.Item("U_Delivr").Cells.Item(i).Specific.Value = (quantity)

                    '    '        End If
                    '    '    End If
                    '    '    oRsBranch.MoveNext()
                    '    '    'End If
                    '    'Next

                    '    ''For Special Prices
                    '    'Me.LoadOnItemOrUOMSelection_SpecialPrices(objForm.UniqueID, i, IsUomDefault, ItemCode, UoM, objForm.Items.Item("4").Specific.Value)
                    '    '    Next
                    '    For i As Integer = 1 To objMatrix.VisualRowCount
                    '        Dim IsUomDefault As String = "Yes"
                    '        Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.ToString()
                    '        Dim UoM As String = objMatrix.Columns.Item("1470002145").Cells.Item(i).Specific.Value.ToString()
                    '        Dim BarCode As String = objMatrix.Columns.Item("4").Cells.Item(i).Specific.Value.ToString()
                    '        If UoM <> "" Then
                    '            If Me.CheckDefaultUOM(UoM, ItemCode) = True Then
                    '                IsUomDefault = "Yes"
                    '            Else
                    '                IsUomDefault = "No"
                    '            End If
                    '        End If
                    '        'Me.LoadOnItemOrUOMSelection(objForm.UniqueID, i, IsUomDefault, ItemCode, UoM, objForm.Items.Item("4").Specific.Value)

                    '        '''For Special Prices
                    '        'Me.LoadOnItemOrUOMSelection_SpecialPrices(objForm.UniqueID, i, IsUomDefault, ItemCode, UoM, objForm.Items.Item("4").Specific.Value, objForm.Items.Item("10").Specific.Value, BarCode)
                    '    Next
                    'End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("OPDN")
                    objMatrix = objForm.Items.Item("38").Specific

                    If pVal.ItemUID = "38" And pVal.ColUID = "1" And pVal.BeforeAction = False And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                        Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value.ToString()
                        If ItemCode <> "" Then
                            Try
                                Dim CheckKitItem As String = "Select ""U_KITITEM"" From OITT Where ""Code"" = '" & ItemCode & "' And ""U_KITITEM"" = 'Yes'"
                                Dim oRsCheckKitItem As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsCheckKitItem.DoQuery(CheckKitItem)
                                If oRsCheckKitItem.RecordCount > 0 Then
                                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_SplitKitItems")
                                    oColumn.Cells.Item(pVal.Row).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                End If
                            Catch ex As Exception
                            End Try
                        End If
                    End If
                    If pVal.ItemUID = "38" And pVal.ColUID = "15" And pVal.BeforeAction = False And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        If (Convert.ToDouble(objMatrix.Columns.Item("15").Cells.Item(pVal.Row).Specific.Value) > 0) Then
                            DiscPrcnt = 0

                            DiscPrcnt = Convert.ToDouble(objMatrix.Columns.Item("15").Cells.Item(pVal.Row).Specific.Value)
                            If DiscPrcnt > 0 Then
                                NST = objMatrix.Columns.Item("17").Cells.Item(pVal.Row).Specific.Value
                                NST = NST.Replace(objForm.Items.Item("63").Specific.Value, "")
                            End If

                            Dim RetailPrice As Double = objMatrix.Columns.Item("U_RetailPrice").Cells.Item(pVal.Row).Specific.Value
                            If RetailPrice > 0 And Convert.ToDouble(NST) > 0 Then

                                If Convert.ToDouble(NST) > RetailPrice Then
                                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_BelowCst")
                                    oColumn.Cells.Item(pVal.Row).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Else
                                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_BelowCst")
                                    oColumn.Cells.Item(pVal.Row).Specific.Select("No", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                End If
                            End If
                        End If
                    End If

                    If pVal.ItemUID = "38" And pVal.ColUID = "1" And pVal.BeforeAction = False Then
                        ' For i As Integer = 1 To objMatrix.VisualRowCount - 1
                        If objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            Try
                                objForm.Freeze(True)

                                objMatrix = objForm.Items.Item("38").Specific
                                Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value
                                Dim GetDocEntry As String = "Select ""U_EXCVAT"" From OITM Where ""ItemCode"" = '" & ItemCode & "'"
                                Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRs.DoQuery(GetDocEntry)
                                objMatrix.Columns.Item("U_EXCTAX").Cells.Item(pVal.Row).Specific.Value() = oRs.Fields.Item("U_EXCVAT").Value()


                            Catch ex As Exception

                            Finally
                                objForm.Freeze(False)

                            End Try

                        End If
                        'Next
                    End If

                    'If pVal.ItemUID = "38" And pVal.ColUID = "1" And pVal.BeforeAction = False Then
                    '    For i As Integer = 1 To objMatrix.VisualRowCount - 1
                    '        If objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value <> "" Then
                    '            objMatrix = objForm.Items.Item("38").Specific
                    '            Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value
                    '            Dim GetDocEntry As String = "Select ""U_EXCVAT"" From OITM Where ""ItemCode"" = '" & ItemCode & "' "
                    '            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '            oRs.DoQuery(GetDocEntry)
                    '            objMatrix.Columns.Item("U_EXCTAX").Cells.Item(i).Specific.Value() = oRs.Fields.Item("U_EXCVAT").Value()
                    '        End If
                    '    Next
                    'End If

                    If pVal.ItemUID = "38" And pVal.ColUID = "U_DSAT" And pVal.BeforeAction = False Then
                        objMatrix.Columns.Item("U_BAD").Cells.Item(pVal.Row).Specific.Value = (objMatrix.Columns.Item("U_BASE").Cells.Item(pVal.Row).Specific.Value - objMatrix.Columns.Item("U_DSAT").Cells.Item(pVal.Row).Specific.Value)
                        objMatrix.Columns.Item("U_EXT").Cells.Item(pVal.Row).Specific.Value = (objMatrix.Columns.Item("U_BASE").Cells.Item(pVal.Row).Specific.Value * 50) / 100

                        Dim bad As Double = 0
                        Dim ext As Double = 0

                        Double.TryParse(objMatrix.Columns.Item("U_BAD").Cells.Item(pVal.Row).Specific.Value, bad)
                        Double.TryParse(objMatrix.Columns.Item("U_EXT").Cells.Item(pVal.Row).Specific.Value, ext)

                        objMatrix.Columns.Item("14").Cells.Item(pVal.Row).Specific.Value = (bad + ext).ToString()
                        objMatrix.Columns.Item("U_suppprice").Cells.Item(pVal.Row).Specific.Value = (bad + ext).ToString()

                    End If

                    If pVal.ItemUID = "38" And pVal.ColUID = "4" And pVal.BeforeAction = False Then
                        If objMatrix.Columns.Item("4").Cells.Item(pVal.Row).Specific.Value.ToString().Trim <> "" Then
                            objMain.objApplication.StatusBar.SetText("Please wait while loading line details....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            'For i As Integer = 1 To objMatrix.VisualRowCount
                            Dim IsUomDefault As String = "Yes"
                            Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value.ToString()
                            Dim UoM As String = objMatrix.Columns.Item("1470002145").Cells.Item(pVal.Row).Specific.Value.ToString()
                            Dim BarCode As String = objMatrix.Columns.Item("4").Cells.Item(pVal.Row).Specific.Value.ToString()
                            If UoM <> "" Then
                                If Me.CheckDefaultUOM(UoM, ItemCode) = True Then
                                    IsUomDefault = "Yes"
                                Else
                                    IsUomDefault = "No"
                                End If
                            End If
                            'Me.LoadOnItemOrUOMSelection(objForm.UniqueID, i, IsUomDefault, ItemCode, UoM, objForm.Items.Item("4").Specific.Value)

                            '''For Special Prices
                            Me.LoadOnItemOrUOMSelection_LostFocus(objForm.UniqueID, pVal.Row, IsUomDefault, ItemCode, UoM, objForm.Items.Item("4").Specific.Value, objForm.Items.Item("10").Specific.Value, BarCode)
                            'Next
                            objMain.objApplication.StatusBar.SetText("Details loaded successfully....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                        End If
                    End If

                    If pVal.ItemUID = "38" And pVal.ColUID = "U_SuppDiscPer" And pVal.BeforeAction = False Then
                        objForm.Freeze(True)
                        'isUpdating = True
                        Dim SUPP As Double = objMatrix.Columns.Item("U_SuppPri").Cells.Item(pVal.Row).Specific.Value * objMatrix.Columns.Item("U_SuppDiscPer").Cells.Item(pVal.Row).Specific.Value / 100
                        'Dim SUPP As Double = objMatrix.Columns.Item("U_SuppPri").Cells.Item(pVal.Row).Specific.Value * objMatrix.Columns.Item("U_SuppDiscPer").Cells.Item(pVal.Row).Specific.Value / 100

                        objMatrix.Columns.Item("U_DISCA").Cells.Item(pVal.Row).Specific.Value = (SUPP).ToString()
                        ''objMatrix.Columns.Item("U_suppprice").Cells.Item(pVal.Row).Specific.Value = (objMatrix.Columns.Item("U_SuppPri").Cells.Item(pVal.Row).Specific.Value - SUPP).ToString()
                        objMatrix.Columns.Item("U_AQTY").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        'isUpdating = False
                        objForm.Freeze(False)
                    End If

                    If pVal.ItemUID = "38" And pVal.ColUID = "U_DISCA" And pVal.BeforeAction = False Then

                        If objMatrix.Columns.Item("U_DISCA").Cells.Item(pVal.Row).Specific.Value = 0 Then
                            Dim basePrice As Double = CDbl(objMatrix.Columns.Item("U_SuppPriBeDis").Cells.Item(pVal.Row).Specific.Value)

                            objMatrix.Columns.Item("U_SuppPri").Cells.Item(pVal.Row).Specific.Value = basePrice.ToString()
                            objMatrix.Columns.Item("U_suppprice").Cells.Item(pVal.Row).Specific.Value = basePrice.ToString()
                            ''objMatrix.Columns.Item("U_SuppPri").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        End If

                    End If

                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    If pVal.ItemUID = "38" AndAlso pVal.ColUID = "U_DISCA" AndAlso Not pVal.BeforeAction Then

                        Dim currValue As String = objMatrix.Columns.Item("U_DISCA").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()

                        ' Ignore duplicate validate for the same value
                        If LastDiscRow = pVal.Row AndAlso LastDiscValue = currValue Then
                            Exit Select
                        End If

                        LastDiscRow = pVal.Row
                        LastDiscValue = currValue

                        objForm.Freeze(True)

                        Try
                            Dim basePrice As Double = 0
                            Dim discAmt As Double = 0

                            Double.TryParse(objMatrix.Columns.Item("U_SuppPri").Cells.Item(pVal.Row).Specific.Value.ToString(), basePrice)
                            Double.TryParse(currValue, discAmt)

                            If basePrice > 0 And discAmt <> 0 Then
                                objMatrix.Columns.Item("U_suppprice").Cells.Item(pVal.Row).Specific.Value = (basePrice - discAmt).ToString("0.000")
                            ElseIf basePrice > 0 And discAmt = 0 Then
                                Dim NEWPRICE As Double = CDbl(objMatrix.Columns.Item("U_SuppPriBeDis").Cells.Item(pVal.Row).Specific.Value)

                                objMatrix.Columns.Item("U_suppprice").Cells.Item(pVal.Row).Specific.Value = NEWPRICE.ToString()
                            End If

                        Finally
                            objForm.Freeze(False)
                        End Try

                    End If



                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("OPDN")
                    objMatrix = objForm.Items.Item("38").Specific
                    If pVal.BeforeAction = False Then
                        CalcualtionBtnFlag = False
                        Me.AddItems(objForm.UniqueID)
                        Me.DisableCoulmns(objForm.UniqueID)
                        ' Dim Purchasenumber As String = objForm.Items.Item("8").Specific.Value

                        'Dim oRsBranch As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'Dim Query As String = " Select T1.""Quantity"",T1.""OpenQty"",(T1.""Quantity""- T1.""OpenQty"")as ""Delieverd Quantity"",T1."" LineNum""From OPOR T0  Inner Join " &
                        '                                                    "POR1 T1 On T1.""DocEntry"" = T0.""DocEntry"" Where T0.""DocNum"" = '" & GlobalDocnum & "' "

                        'oRsBranch.DoQuery(Query)
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub AddItems(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim objMatrix As SAPbouiCOM.Matrix
            objMatrix = objForm.Items.Item("38").Specific

            oItem4 = objForm.Items.Add("CAL_NETQTY", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem4.Top = objForm.Items.Item("3").Top
            oItem4.Left = objForm.Items.Item("3").Left + 80
            oItem4.Width = objForm.Items.Item("3").Width + 30
            oItem4.Height = objForm.Items.Item("3").Height
            oItem4.FromPane = objForm.Items.Item("3").FromPane
            oItem4.ToPane = objForm.Items.Item("3").ToPane
            Dim btn1 As SAPbouiCOM.Button = objForm.Items.Item("CAL_NETQTY").Specific
            btn1.Caption = "Calculate NetQty"

            oItem1 = objForm.Items.Add("btn_LdItd", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem1.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 10
            oItem1.Left = objForm.Items.Item("38").Left
            oItem1.Width = objForm.Items.Item("2").Width + 30
            oItem1.Height = objForm.Items.Item("2").Height
            oItem1.FromPane = objForm.Items.Item("38").FromPane
            oItem1.ToPane = objForm.Items.Item("38").ToPane
            Dim btn3 As SAPbouiCOM.Button = objForm.Items.Item("btn_LdItd").Specific
            btn3.Caption = "Load &Item Details"

            oItem2 = objForm.Items.Add("btn_calFOC", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem2.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 10
            oItem2.Left = objForm.Items.Item("btn_LdItd").Width + objForm.Items.Item("btn_LdItd").Left + 5
            oItem2.Width = objForm.Items.Item("2").Width + 30
            oItem2.Height = objForm.Items.Item("2").Height
            oItem2.FromPane = objForm.Items.Item("38").FromPane
            oItem2.ToPane = objForm.Items.Item("38").ToPane
            Dim btn2 As SAPbouiCOM.Button = objForm.Items.Item("btn_calFOC").Specific
            btn2.Caption = "FOC Calculation"

            objMain.objUtilities.AddButton(objForm.UniqueID, "btn_MRG", objForm.Items.Item("btn_calFOC").Top,
                                        objForm.Items.Item("btn_calFOC").Width + objForm.Items.Item("btn_calFOC").Left + 5, "150", objForm.Items.Item("2").Height, "btn_calFOC", "Calculate Margin", objForm.Items.Item("38").FromPane, objForm.Items.Item("38").ToPane)


            objMain.objUtilities.AddEditBox(objForm.UniqueID, "txt_GUID", objForm.Items.Item("4").Top, objForm.Items.Item("4").Left + objForm.Items.Item("4").Width + 5, objForm.Items.Item("4").Width, "OPDN", "U_GUID", "", objForm.Items.Item("4").Height, 96, 96)

            objForm.Items.Item("txt_GUID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objMain.objUtilities.AddButton(objForm.UniqueID, "btn_LdItd", objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 10, _
            '                            objForm.Items.Item("38").Left, "100", "116", _
            '                             "Load &Item Details", objForm.Items.Item("38").FromPane, objForm.Items.Item("38").ToPane)
            oitem = objForm.Items.Add("Split", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oitem.Top = objForm.Items.Item("2").Top
            oitem.Left = objForm.Items.Item("2").Left + objForm.Items.Item("2").Width + 5
            oitem.Width = objForm.Items.Item("2").Width + 10
            oitem.Height = objForm.Items.Item("2").Height
            oitem.FromPane = objForm.Items.Item("2").FromPane
            oitem.ToPane = objForm.Items.Item("2").ToPane
            Dim btn As SAPbouiCOM.Button = objForm.Items.Item("Split").Specific
            btn.Caption = "Split Items"

            oItem3 = objForm.Items.Add("Barcode", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem3.Top = objForm.Items.Item("10000329").Top - objForm.Items.Item("10000329").Height - 2
            oItem3.Left = objForm.Items.Item("10000329").Left
            oItem3.Width = objForm.Items.Item("10000329").Width
            oItem3.Height = objForm.Items.Item("10000329").Height
            Dim btn4 As SAPbouiCOM.Button = objForm.Items.Item("Barcode").Specific
            btn4.Caption = "Load Barcode"

            'objMain.objUtilities.AddButton(objForm.UniqueID, "btn_calFOC", objForm.Items.Item("btn_LdItd").Top,
            '                            objForm.Items.Item("btn_LdItd").Width + objForm.Items.Item("btn_LdItd").Left + 5, "100", "116",
            '                            "FOC Calculation", objForm.Items.Item("38").FromPane, objForm.Items.Item("38").ToPane)


            'objMain.objUtilities.AddButton(objForm.UniqueID, "btn_ED", objForm.Items.Item("2").Top,
            '                            objForm.Items.Item("Split").Width + objForm.Items.Item("Split").Left + 5, "100", objForm.Items.Item("2").Height, "2", "Expiry Date", objForm.Items.Item("2").FromPane, objForm.Items.Item("2").ToPane)

            Dim Validation As String = "Select ""U_GRPO""as ""GRPO"" from ""@ME_ACFG"""
            Dim oLValidation As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oLValidation.DoQuery(Validation)
            If oLValidation.Fields.Item("GRPO").Value = "Y" Then
                objForm.Items.Item("btn_calFOC").Visible = True
            Else
                objForm.Items.Item("btn_calFOC").Visible = False
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub GoodsIssueAndReceipt(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim objMatrix As SAPbouiCOM.Matrix
            objMatrix = objForm.Items.Item("38").Specific
            Dim GoodsIssuePosted As String = ""
            Dim DocNum As String = objForm.Items.Item("8").Specific.Value
            Dim Series As String = objForm.Items.Item("88").Specific.Value.ToString.Trim

            Dim Qry As String = " SELECT T1.""CodeBars"" As ""BarCode"" , T1.""InvQty"" ,  T1.""DocEntry"" , T1.""ItemCode"", T1.""Quantity"", T1.""Price"",T0.""DocDate"",T1.""WhsCode"" , ""UomCode"" FROM OPDN T0   " &
                                " INNER JOIN PDN1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
                                " WHERE T0.""DocNum""='" & DocNum & "' and T0.""Series""='" & Series & "' and T1.""U_SplitKitItems""='Yes'  "
            Dim oRsIssue As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsIssue.DoQuery(Qry)
            If oRsIssue.RecordCount > 0 Then
                oRsIssue.MoveFirst()
                For k As Integer = 1 To oRsIssue.RecordCount
                    Dim ChkItmSz As String = " Select T2.""Code"" From OITT T1 Inner Join ITT1 T2 On T1.""Code"" = T2.""Father"" " &
                                             " Inner Join OITM T3 On T3.""ItemCode"" = T2.""Code"" Where T1.""Code"" = '" & oRsIssue.Fields.Item("ItemCode").Value & "' And IFNULL(""U_ItemSize"",0) = 0  "
                    Dim oRsChkItmSz As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsChkItmSz.DoQuery(ChkItmSz)
                    If oRsChkItmSz.RecordCount > 0 Then
                        objMain.objApplication.StatusBar.SetText("Please configure itemsize for the itemcode " & oRsChkItmSz.Fields.Item(0).Value & ", splitting the items.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If
                    oRsIssue.MoveNext()
                Next
                oRsIssue.MoveFirst()
                Dim QryExist As String = "Select * From OIGE Where ""U_GRPO""='" & oRsIssue.Fields.Item("DocEntry").Value & "'  "
                Dim oRsExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsExist.DoQuery(QryExist)
                If oRsExist.RecordCount = 0 Then
                    Dim oGI As SAPbobsCOM.Documents
                    oGI = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
                    oGI.DocDate = oRsIssue.Fields.Item("DocDate").Value
                    oGI.UserFields.Fields.Item("U_TYPE").Value = "KitItem"
                    'oGI.Reference2 = DocNum
                    oGI.Reference2 = "Kit Item"

                    'oGI.JournalMemo = DocNum
                    objComboBox = objForm.Items.Item("2001").Specific
                    Dim Branch As Integer = objComboBox.Value
                    oGI.BPL_IDAssignedToInvoice = Branch
                    Try
                        oGI.UserFields.Fields.Item("U_GRPO").Value = oRsIssue.Fields.Item("DocEntry").Value.ToString
                    Catch ex As Exception
                    End Try
                    Try
                        oGI.UserFields.Fields.Item("U_BaseNum").Value = DocNum
                    Catch ex As Exception
                    End Try
                    Dim GetGISeries As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = '60' And ""Locked"" = 'N' And ""BPLId"" = '" & Branch.ToString() & "' Order By ""Series"" Desc "
                    Dim oRsGetGISeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetGISeries.DoQuery(GetGISeries)
                    Dim GISeries As String = oRsGetGISeries.Fields.Item("Series").Value
                    oGI.Series = GISeries

                    For i As Integer = 1 To oRsIssue.RecordCount
                        oGI.Lines.ItemCode = oRsIssue.Fields.Item("ItemCode").Value
                        If oRsIssue.Fields.Item("BarCode").Value <> "" Then
                            oGI.Lines.BarCode = oRsIssue.Fields.Item("BarCode").Value
                        End If
                        oGI.Lines.Quantity = oRsIssue.Fields.Item("Quantity").Value
                        oGI.Lines.UnitPrice = oRsIssue.Fields.Item("Price").Value
                        oGI.Lines.WarehouseCode = oRsIssue.Fields.Item("WhsCode").Value
                        Dim GetAddonConfigurationDtls As String = "Select ""U_KITMIACT"" From ""@ME_ACFG"" "
                        Dim oRsGetAddonConfigurationDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetAddonConfigurationDtls.DoQuery(GetAddonConfigurationDtls)
                        oGI.Lines.AccountCode = oRsGetAddonConfigurationDtls.Fields.Item("U_KITMIACT").Value
                        Dim UomEntry As Integer = Me.GetUomEntry(oRsIssue.Fields.Item("UomCode").Value)
                        Dim GetItSrBtchDet As String = " SELECT " & IsNull & "(T0.""ManBtchNum"",'N')," & IsNull & "(T0.""ManSerNum"",'N')," & IsNull & "(T0.""MngMethod"",'N') " &
                                                      " FROM OITM T0 WHERE T0.""ItemCode""='" & oRsIssue.Fields.Item("ItemCode").Value & "'"
                        Dim oRsGetItSrBtchDet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetItSrBtchDet.DoQuery(GetItSrBtchDet)

                        If Me.CheckDefaultUOM(oRsIssue.Fields.Item("UomCode").Value, oRsIssue.Fields.Item("ItemCode").Value) = True Then
                            oGI.Lines.UseBaseUnits = SAPbobsCOM.BoYesNoEnum.tNO
                            oGI.Lines.UoMEntry = UomEntry
                            If oRsGetItSrBtchDet.RecordCount > 0 Then
                                If CStr(oRsGetItSrBtchDet.Fields.Item(0).Value).Trim = "Y" Then
                                    Dim GetBatchDetails As String = " Select ""BatchNum"" From IBT1 Where ""BaseType"" = '20' And  ""BaseEntry"" = '" & oRsIssue.Fields.Item("DocEntry").Value & "' " &
                                                                    " And  ""ItemCode""  = '" & oRsIssue.Fields.Item("ItemCode").Value & "' And ""WhsCode"" = '" & oRsIssue.Fields.Item("WhsCode").Value & "' "
                                    Dim oRsGetBatchDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRsGetBatchDetails.DoQuery(GetBatchDetails)
                                    If oRsGetBatchDetails.RecordCount > 0 Then
                                        oGI.Lines.BatchNumbers.BatchNumber = oRsGetBatchDetails.Fields.Item("BatchNum").Value.Trim
                                        oGI.Lines.BatchNumbers.Quantity = CDbl(oRsIssue.Fields.Item("Quantity").Value)
                                        oGI.Lines.BatchNumbers.Add()
                                    End If
                                End If
                            End If
                        Else
                            oGI.Lines.UseBaseUnits = SAPbobsCOM.BoYesNoEnum.tYES
                            oGI.Lines.UoMEntry = UomEntry
                            If oRsGetItSrBtchDet.RecordCount > 0 Then
                                If CStr(oRsGetItSrBtchDet.Fields.Item(0).Value).Trim = "Y" Then
                                    Dim GetBatchDetails As String = " Select ""BatchNum"" From IBT1 Where ""BaseType"" = '20' And  ""BaseEntry"" = '" & oRsIssue.Fields.Item("DocEntry").Value & "' " &
                                                                    " And  ""ItemCode""  = '" & oRsIssue.Fields.Item("ItemCode").Value & "' And ""WhsCode"" = '" & oRsIssue.Fields.Item("WhsCode").Value & "' "
                                    Dim oRsGetBatchDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRsGetBatchDetails.DoQuery(GetBatchDetails)
                                    If oRsGetBatchDetails.RecordCount > 0 Then
                                        oGI.Lines.BatchNumbers.BatchNumber = oRsGetBatchDetails.Fields.Item("BatchNum").Value.Trim
                                        oGI.Lines.BatchNumbers.Quantity = CDbl(oRsIssue.Fields.Item("InvQty").Value)
                                        oGI.Lines.BatchNumbers.Add()
                                    End If
                                End If
                            End If
                        End If
                        oGI.Lines.Add()
                        oRsIssue.MoveNext()
                    Next
                    result = oGI.Add()
                    If result <> 0 Then
                        objMain.objApplication.StatusBar.SetText("Error:" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Else
                        objMain.objApplication.StatusBar.SetText("Goods issue is posted", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        GoodsIssuePosted = "Yes"
                    End If
                Else
                    GoodsIssuePosted = "Yes"
                End If
                '--Goods Receipt
                If GoodsIssuePosted = "Yes" Then
                    oRsIssue.MoveFirst()
                    QryExist = "Select * From OIGN Where ""U_GRPO""='" & oRsIssue.Fields.Item("DocEntry").Value & "'"
                    oRsExist = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsExist.DoQuery(QryExist)
                    If oRsExist.RecordCount = 0 Then
                        Dim oRsReceipt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim QryRec As String = " SELECT T3.""Code"" AS ""ItemCode"", (T1.""Quantity"" * T3.""Quantity"") ""Quantity"" ,T1.""Price"" " &
                                               " ,T0.""DocDate"",T1.""WhsCode"" As""Warehouse"" ,  (T4.""U_ItemSize"" / (Select SUM(Cast(""U_ItemSize"" As Integer))  From OITT A Inner Join ITT1 B On A.""Code"" =  B.""Father"" Inner Join OITM C On C.""ItemCode"" = B.""Code"" " &
                                               "  Where A.""Code"" = T1.""ItemCode""))  As ""ItemSize"" , T1.""LineNum"" , T1.""DocEntry""  FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
                                               " INNER JOIN OITT T2 ON T1.""ItemCode""= T2.""Code""INNER JOIN ITT1 T3 ON T2.""Code"" = T3.""Father"" " &
                                               " INNER JOIN OITM T4 ON T4.""ItemCode""=T3.""Code""WHERE T0.""DocNum""='" & DocNum & "'  " &
                                               " and T0.""Series""='" & Series & "' and T1.""U_SplitKitItems""='Yes'"
                        oRsReceipt.DoQuery(QryRec)
                        If GoodsIssuePosted = "Yes" Then
                            If oRsReceipt.RecordCount > 0 Then
                                Dim oGR As SAPbobsCOM.Documents
                                oGR = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
                                oGR.DocDate = oRsReceipt.Fields.Item("DocDate").Value
                                objComboBox = objForm.Items.Item("2001").Specific
                                Dim Branch As Integer = objComboBox.Value
                                oGR.BPL_IDAssignedToInvoice = Branch
                                Dim GetGRSeries As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = '59' And ""Locked"" = 'N' And ""BPLId"" = '" & Branch.ToString() & "' Order By ""Series"" Desc "
                                Dim oRsGetGRSeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetGRSeries.DoQuery(GetGRSeries)
                                Dim GRSeries As String = oRsGetGRSeries.Fields.Item("Series").Value
                                oGR.Series = GRSeries
                                oGR.Reference2 = "Kit Item"
                                'oGR.JournalMemo = DocNum
                                Try
                                    oGR.UserFields.Fields.Item("U_GRPO").Value = oRsIssue.Fields.Item("DocEntry").Value.ToString
                                Catch ex As Exception
                                End Try
                                Try
                                    oGR.UserFields.Fields.Item("U_BaseNum").Value = DocNum
                                Catch ex As Exception
                                End Try
                                oGR.UserFields.Fields.Item("U_TYPE").Value = "KitItem"
                                For i As Integer = 1 To oRsReceipt.RecordCount
                                    oGR.Lines.ItemCode = oRsReceipt.Fields.Item("ItemCode").Value
                                    oGR.Lines.Quantity = oRsReceipt.Fields.Item("Quantity").Value
                                    Dim UnitPrice As Double = 0
                                    UnitPrice = CDbl(oRsReceipt.Fields.Item("Price").Value) * CDbl(oRsReceipt.Fields.Item("ItemSize").Value)
                                    oGR.Lines.UnitPrice = UnitPrice
                                    oGR.Lines.WarehouseCode = oRsReceipt.Fields.Item("Warehouse").Value
                                    Dim GetAddonConfigurationDtls As String = "Select ""U_KITMRACT"" From ""@ME_ACFG"" "
                                    Dim oRsGetAddonConfigurationDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRsGetAddonConfigurationDtls.DoQuery(GetAddonConfigurationDtls)
                                    oGR.Lines.AccountCode = oRsGetAddonConfigurationDtls.Fields.Item("U_KITMRACT").Value
                                    Dim UomEntry As Integer = Me.GetUomEntry(Me.GetDefaultUomCode(oRsReceipt.Fields.Item("ItemCode").Value))
                                    Dim GetItSrBtchDet As String = " Select " & IsNull & "(T0.""ManBtchNum"",'N')," & IsNull & "(T0.""ManSerNum"",'N')," & IsNull & "(T0.""MngMethod"",'N') " &
                                                                   " FROM OITM T0 WHERE T0.""ItemCode""='" & oRsReceipt.Fields.Item("ItemCode").Value & "'"
                                    Dim oRsGetItSrBtchDet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRsGetItSrBtchDet.DoQuery(GetItSrBtchDet)
                                    oGR.Lines.UseBaseUnits = SAPbobsCOM.BoYesNoEnum.tNO
                                    oGR.Lines.UoMEntry = UomEntry
                                    If oRsGetItSrBtchDet.RecordCount > 0 Then
                                        If CStr(oRsGetItSrBtchDet.Fields.Item(0).Value).Trim = "Y" Then
                                            Dim GetBatchDetails As String = " Select ""BatchNum"" From IBT1 Where ""BaseType"" = '20' And  ""BaseEntry"" = '" & oRsIssue.Fields.Item("DocEntry").Value & "' " &
                                                                            " And  ""ItemCode""  = '" & oRsReceipt.Fields.Item("ItemCode").Value & "' And ""WhsCode"" = '" & oRsReceipt.Fields.Item("Warehouse").Value & "' "
                                            Dim oRsGetBatchDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRsGetBatchDetails.DoQuery(GetBatchDetails)
                                            If oRsGetBatchDetails.RecordCount > 0 Then
                                                oGR.Lines.BatchNumbers.BatchNumber = oRsGetBatchDetails.Fields.Item("BatchNum").Value.Trim
                                                oGR.Lines.BatchNumbers.Quantity = CDbl(oRsIssue.Fields.Item("Quantity").Value)
                                                oGR.Lines.BatchNumbers.Add()
                                            End If
                                        End If
                                    End If
                                    oGR.Lines.Add()
                                    oRsReceipt.MoveNext()
                                Next
                                result = oGR.Add()
                                If result <> 0 Then
                                    objMain.objApplication.StatusBar.SetText("Error:" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Else
                                    Dim QryRec1 As String = " SELECT  T1.""LineNum"" , T1.""DocEntry""  FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
                                              " INNER JOIN OITT T2 ON T1.""ItemCode""= T2.""Code""INNER JOIN ITT1 T3 ON T2.""Code"" = T3.""Father"" " &
                                              " INNER JOIN OITM T4 ON T4.""ItemCode""=T3.""Code""WHERE T0.""DocNum""='" & DocNum & "'  " &
                                              " and T0.""Series""='" & Series & "' and T1.""U_SplitKitItems""='Yes' Group By T1.""LineNum"" , T1.""DocEntry"" "
                                    Dim oRsReceipt1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRsReceipt1.DoQuery(QryRec1)
                                    For i As Integer = 1 To oRsReceipt1.RecordCount
                                        Dim UpdateSplitItems As String = " Update ""PDN1"" Set ""U_SPLITTED"" = 'Yes'  " &
                                                                         " Where ""DocEntry""='" & oRsReceipt1.Fields.Item("DocEntry").Value & "' And " &
                                                                         " ""LineNum"" = '" & oRsReceipt1.Fields.Item("LineNum").Value & "'  "
                                        Dim oRsUpdateSplitItems As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRsUpdateSplitItems.DoQuery(UpdateSplitItems)
                                        oRsReceipt1.MoveNext()
                                    Next
                                    objMain.objApplication.StatusBar.SetText("Goods Receipt is posted", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                End If
                            End If
                        Else
                            objMain.objApplication.StatusBar.SetText("Goods issue is not added", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                    Else
                        objMain.objApplication.StatusBar.SetText("Split Document is already posted", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                Else
                    objMain.objApplication.StatusBar.SetText("Goods issue not posted", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub DisableCoulmns(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("38").Specific
            Try
                objMatrix.Columns.Item("U_INVTRRLN").Visible = False
                objMatrix.Columns.Item("U_INVTRREQ").Visible = False
                objMatrix.Columns.Item("U_INVTRREN").Visible = False
                objMatrix.Columns.Item("U_MEOPNQTY").Visible = False
                objMatrix.Columns.Item("U_MERCEQTY").Visible = False
                objMatrix.Columns.Item("U_MEREQQTY").Visible = False
                objMatrix.Columns.Item("U_METSTQTY").Visible = False
                objMatrix.Columns.Item("U_MIEntry").Visible = False
                objMatrix.Columns.Item("U_MIDocNum").Visible = False
            Catch ex As Exception
            End Try
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub LoadOnItemOrUOMSelection(ByVal FormUID As String, ByVal RowId As Integer, ByVal IsUomDefault As String, ByVal ItemCode As String, ByVal Uom As String, ByVal Vendor As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("OPDN")
            objMatrix = objForm.Items.Item("38").Specific
            Dim SalesPrice As Double = 0
            Dim BaseQty As Double = 0
            Dim Packing As String = ""
            Dim RetailsPrice As Double = 0
            Dim AdditionalBarCode As String = ""
            Dim UOMEntry As Integer = 0
            Dim LeastVendorPrice As Double = 0
            Dim LeastVendorCode As String = ""
            Dim LeastVendorName As String = ""
            Dim LeastVendorValidation As String = ""
            Dim PCPrPack As String = objMain.objUtilities.GetPCPerItem(ItemCode)
            BaseUom = objMain.objUtilities.GetItemBaseUom(ItemCode)



            If objForm.Items.Item("2001").Specific.Value.Trim = "" Then
                objMain.objApplication.StatusBar.SetText("Please select branch at header...")
                Exit Sub
            End If

            Dim LPP As Double = 0


            If IsUomDefault = "No" Then
                UOMEntry = Me.GetUomEntry(Uom)
                'Dim GetPurchasePrice As String = " Select T0.""Price"" From ITM9 T0 Where T0.""PriceList""= " & _
                '                                 " (Select ""ListNum"" From OCRD Where ""CardCode"" = '" & Vendor & "') And " & _
                '                                 " T0.""UomEntry""='" & UOMEntry & "' And ""ItemCode""='" & ItemCode & "' "
                'Dim oRsGetPurchasePrice As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oRsGetPurchasePrice.DoQuery(GetPurchasePrice)
                'If oRsGetPurchasePrice.RecordCount > 0 Then
                '    SalesPrice = CDbl(oRsGetPurchasePrice.Fields.Item(0).Value)
                'End If

                'Dim GetPurchasePrice As String = " Select IFNULL(T3.""Price"",0) ""Price"" from SPP2 T3  " & _
                '                                 " Where T3.""ItemCode"" = '" & ItemCode & "' and T3.""CardCode"" = '" & Vendor & "' and T3.""UomEntry"" = '" & UOMEntry & "' "
                'Dim oRsGetPurchasePrice As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oRsGetPurchasePrice.DoQuery(GetPurchasePrice)
                'If oRsGetPurchasePrice.RecordCount > 0 Then
                '    SalesPrice = CDbl(oRsGetPurchasePrice.Fields.Item(0).Value)
                'Else
                '    SalesPrice = Me.getSupplierPrice_SpecialPrices(objForm.UniqueID, IsUomDefault, ItemCode, Vendor, UOMEntry, Uom)
                'End If


                Dim GetLastLPP As String = "Select IFNULL((Select (case when B.""UomEntry""=X.""UomEntry""  and X.""UomEntry""=Y.""BaseUom"" then B.""Price""  " &
                                            "when B.""UomEntry""=X.""UomEntry""  and X.""UomEntry""<>Y.""BaseUom"" then B.""Price""*1/X.""AltQty""*X.""BaseQty""  " &
                                            "when B.""UomEntry""<>X.""UomEntry"" and X.""UomEntry""<>Y.""BaseUom"" then B.""Price""*1/X.""AltQty""*X.""BaseQty""  " &
                                            "when B.""UomEntry""<>X.""UomEntry"" and X.""UomEntry""=Y.""BaseUom"" then B.""Price"" " &
                                            "else 0 end) from OITM C inner join (SELECT H.""ItemCode"",IFNULL((IFNULL(H.""INMPrice"", 0)/H.""NumPerMsr""),0)-IFNULL((IFNULL(A1.""INMPrice"", 0)/A1.""NumPerMsr""),0) as ""Price"", H.""UomCode"",H.""UomEntry"",H.""NumPerMsr"",Row_Number() over  " &
                                            "(partition by H.""ItemCode"" order by H.""DocEntry"" Desc) as ""RN"" FROM ""PDN1"" H left join OPDN A3 on A3.""DocEntry""=H.""DocEntry"" left join ORPC A2 on  A2.""U_GRPO""=A3.""DocNum"" " &
                                            "left join RPC1 A1 on A2.""DocEntry""=A1.""DocEntry"" and A1.""ItemCode""=H.""ItemCode"" and A1.""NoInvtryMv""='Y' where H.""DocDate""<CURRENT_DATE and IFNULL(H.""U_normalqty"",0)<>0)B on C.""ItemCode""=B.""ItemCode"" left join OUGP Y " &
                                            "on Y.""UgpEntry""=C.""UgpEntry""  left JOIN UGP1 X ON Y.""UgpEntry"" = X.""UgpEntry"" where ""RN""=1  " &
                                            "and  B.""ItemCode"" ='" & ItemCode & "' and  X.""UomEntry""='" & UOMEntry & "'),0) as ""LPP"" FROM DUMMY"

                Dim oRsGetLastLPP As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                oRsGetLastLPP.DoQuery(GetLastLPP)

                If oRsGetLastLPP.RecordCount > 0 Then
                    LPP = oRsGetLastLPP.Fields.Item(0).Value
                End If

                Dim GetPurchasePrice As String = " Select T0.""Price"" From ITM1 T0 Where T0.""PriceList""= " &
                                               " (Select ""ListNum"" From OCRD Where ""CardCode"" = '" & Vendor & "') And " &
                                              " ""ItemCode""='" & ItemCode & "' "
                Dim oRsGetPurchasePrice As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetPurchasePrice.DoQuery(GetPurchasePrice)
                If oRsGetPurchasePrice.RecordCount > 0 And LPP = 0 Then
                    SalesPrice = CDbl(oRsGetPurchasePrice.Fields.Item(0).Value)
                Else
                    SalesPrice = Me.getSupplierPrice_SpecialPrices(objForm.UniqueID, IsUomDefault, ItemCode, Vendor, UOMEntry, Uom)
                End If


                'Dim GetPackageCode As String = " Select Cast(Cast(T0.""BaseQty"" As Int) As Varchar)" & Concate & "'*'" & Concate & "'" & BaseUom & "'" & _
                '                               " From UGP1 T0 Inner Join OUOM T1 On  " & _
                '                               " T1.""UomEntry"" = T0.""UomEntry""  Where ""UgpEntry"" = '" & Me.GetUgpEntry(ItemCode) & "' And  " & _
                '                               " T1.""UomCode"" = '" & Uom & "' "
                'Dim oRsGetPackageCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oRsGetPackageCode.DoQuery(GetPackageCode)
                'Packing = oRsGetPackageCode.Fields.Item(0).Value
                'If PCPrPack <> "" Then
                '    Packing = PCPrPack & "*" & oRsGetPackageCode.Fields.Item(0).Value             
                'End If

                Dim GetPackageCode As String = "Select ""U_packing"" FROM OITM Where ""ItemCode""='" & ItemCode.Trim & "'"
                Dim oRsGetPackageCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetPackageCode.DoQuery(GetPackageCode)

                Packing = oRsGetPackageCode.Fields.Item(0).Value

                'If PCPrPack <> "" Then
                '    Packing = PCPrPack & "*" & oRsGetPackageCode.Fields.Item(0).Value
                'End If

                Dim GetRetailPriceBrnchWise As String = " Select T1.""Price"" From OITM T0 Inner Join ITM9 T1 On T1.""ItemCode"" = T0.""ItemCode"" " &
                                                        " Inner Join OBPL T2 On T1.""PriceList""=T2.""U_PricelistNo"" Where T0.""ItemCode"" = '" & ItemCode & "' " &
                                                        " And T2.""BPLId""= '" & objForm.Items.Item("2001").Specific.Value.Trim & "' And T1.""UomEntry""='" & UOMEntry & "'"
                Dim oRsGetRetailPriceBrnchWise As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetRetailPriceBrnchWise.DoQuery(GetRetailPriceBrnchWise)
                If oRsGetRetailPriceBrnchWise.RecordCount > 0 Then
                    RetailsPrice = oRsGetRetailPriceBrnchWise.Fields.Item("Price").Value
                Else
                    Dim GetRetailPriceBrnchWise1 As String = " Select T1.""Price"" From OITM T0 Inner Join ITM1 T1 On T1.""ItemCode"" = T0.""ItemCode"" " &
                                                                          " Inner Join OBPL T2 On T1.""PriceList""=T2.""U_PricelistNo"" Where T0.""ItemCode"" = '" & ItemCode & "' " &
                                                                          " And T2.""BPLId""= '" & objForm.Items.Item("2001").Specific.Value.Trim & "' "
                    Dim oRsGetRetailPriceBrnchWise1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetRetailPriceBrnchWise1.DoQuery(GetRetailPriceBrnchWise1)
                    If oRsGetRetailPriceBrnchWise1.RecordCount > 0 Then
                        BaseQty = Me.GetBaseQty(ItemCode, UOMEntry.ToString)
                        RetailsPrice = Math.Round((oRsGetRetailPriceBrnchWise1.Fields.Item("Price").Value * BaseQty), 2)
                    End If
                End If
                'Dim GetLeastVendorDtlsAndPrice As String = " Select Top 1 T0.""ItemCode"" , T1.""CardCode"", T1.""CardName"" , T0.""Price"" From ITM9 T0 " & _
                '                                           " Inner Join OCRD T1 On T1.""ListNum"" = T0.""PriceList""  And T1.""CardCode"" In (Select ""VendorCode"" From ITM2 Where  ITM2.""ItemCode"" = T0.""ItemCode"") Where " & _
                '                                           " ""ItemCode"" = '" & ItemCode & "' And T1.""CardType"" = 'S' And  T0.""UomEntry"" = '" & UOMEntry & "' " & _
                '                                           " Order By T0.""Price""  Asc "
                'Dim GetLeastVendorDtlsAndPrice As String = " Select  TOP 1  Round(Sum(""Price""/""NumPerMsr""),2)  , ""BaseCard"" As ""CardCode"" , ""DocEntry"" , (Select ""CardName"" From OCRD Where ""CardCode"" = ""BaseCard"" ) As ""CardName"" From PDN1 Where  ""ItemCode""= '" & ItemCode & "'  And " & _
                '                                           " ""DocDate"" > (Select ADD_DAYS (TO_DATE (CURRENT_DATE, 'YYYY-MM-DD'), -180)  From DUMMY)  " & _
                '                                           " Group by ""BaseCard"",""DocEntry"" Order By SUM(""Price""/""NumPerMsr"")"

                'Dim GetLeastVendorDtlsAndPrice As String = " Select  TOP 1  Sum(""U_suppprice""/""NumPerMsr"")  , ""BaseCard"" As ""CardCode"" , ""DocEntry"" , (Select ""CardName"" From OCRD Where ""CardCode"" = ""BaseCard"" ) As ""CardName"" From PDN1 Where  ""ItemCode""= '" & ItemCode & "'  And " &
                '                                           " ""DocDate"">(Select ADD_DAYS (TO_DATE (CURRENT_DATE, 'YYYY-MM-DD'), -180)  From DUMMY)  " &
                '                                           " Group by ""BaseCard"",""DocEntry"" , ""U_suppprice"" Order By SUM(""U_suppprice""/""NumPerMsr"")"
                Dim GetLeastVendorDtlsAndPrice As String = " Select  TOP 1  Sum(""U_suppprice""/""NumPerMsr"")  , ""BaseCard"" As ""CardCode"" , ""DocEntry"" , (Select ""CardName"" From OCRD Where ""CardCode"" = ""BaseCard"" ) As ""CardName"" From PDN1 Where  ""ItemCode""= '" & ItemCode & "'  And " &
                                                           " ""DocDate"">(Select ADD_DAYS (TO_DATE (CURRENT_DATE, 'YYYY-MM-DD'), -45)  From DUMMY)  " &
                                                           " Group by ""BaseCard"",""DocEntry"" , ""U_suppprice"" Order By SUM(""U_suppprice""/""NumPerMsr"")"
                Dim oRsGetLeastVendorDtlsAndPrice As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetLeastVendorDtlsAndPrice.DoQuery(GetLeastVendorDtlsAndPrice)
                If oRsGetLeastVendorDtlsAndPrice.RecordCount > 0 Then
                    BaseQty = Me.GetBaseQty(ItemCode, UOMEntry.ToString)
                    LeastVendorPrice = Math.Round((oRsGetLeastVendorDtlsAndPrice.Fields.Item(0).Value * BaseQty), 2)
                    LeastVendorCode = oRsGetLeastVendorDtlsAndPrice.Fields.Item("CardCode").Value
                    LeastVendorName = oRsGetLeastVendorDtlsAndPrice.Fields.Item("CardName").Value
                End If

            ElseIf IsUomDefault = "Yes" Then
                UOMEntry = Me.GetUomEntry(Me.GetDefaultUomCode(ItemCode))
                'Dim GetPurchasePriceDfltUOM As String = "Select ""Price"" From ITM1 Where ""ItemCode""='" & ItemCode & "' AND ""PriceList""= (Select ""ListNum"" From OCRD Where ""CardCode"" = '" & Vendor & "')"
                'Dim oRsGetPurchasePriceDfltUOM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oRsGetPurchasePriceDfltUOM.DoQuery(GetPurchasePriceDfltUOM)
                'SalesPrice = CDbl(oRsGetPurchasePriceDfltUOM.Fields.Item(0).Value)


                Dim GetLastLPP As String = "Select IFNULL((Select (case when B.""UomEntry""=X.""UomEntry""  and X.""UomEntry""=Y.""BaseUom"" then B.""Price""  " &
                                            "when B.""UomEntry""=X.""UomEntry""  and X.""UomEntry""<>Y.""BaseUom"" then B.""Price""*1/X.""AltQty""*X.""BaseQty""  " &
                                            "when B.""UomEntry""<>X.""UomEntry"" and X.""UomEntry""<>Y.""BaseUom"" then B.""Price""*1/X.""AltQty""*X.""BaseQty""  " &
                                            "when B.""UomEntry""<>X.""UomEntry"" and X.""UomEntry""=Y.""BaseUom"" then B.""Price"" " &
                                            "else 0 end) from OITM C inner join (SELECT H.""ItemCode"",IFNULL((IFNULL(H.""INMPrice"", 0)/H.""NumPerMsr""),0)-IFNULL((IFNULL(A1.""INMPrice"", 0)/A1.""NumPerMsr""),0) as ""Price"", H.""UomCode"",H.""UomEntry"",H.""NumPerMsr"",Row_Number() over  " &
                                            "(partition by H.""ItemCode"" order by H.""DocEntry"" Desc) as ""RN"" FROM ""PDN1"" H left join OPDN A3 on A3.""DocEntry""=H.""DocEntry"" left join ORPC A2 on  A2.""U_GRPO""=A3.""DocNum"" " &
                                            "left join RPC1 A1 on A2.""DocEntry""=A1.""DocEntry"" and A1.""ItemCode""=H.""ItemCode"" and A1.""NoInvtryMv""='Y' where H.""DocDate""<CURRENT_DATE and IFNULL(H.""U_normalqty"",0)<>0)B on C.""ItemCode""=B.""ItemCode"" left join OUGP Y " &
                                            "on Y.""UgpEntry""=C.""UgpEntry""  left JOIN UGP1 X ON Y.""UgpEntry"" = X.""UgpEntry"" where ""RN""=1  " &
                                            "and  B.""ItemCode"" ='" & ItemCode & "' and  X.""UomEntry""='" & UOMEntry & "'),0) as ""LPP"" FROM DUMMY"

                Dim oRsGetLastLPP As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                oRsGetLastLPP.DoQuery(GetLastLPP)

                If oRsGetLastLPP.RecordCount > 0 Then
                    LPP = oRsGetLastLPP.Fields.Item(0).Value
                End If

                Dim GetPurchasePrice As String = " Select IFNULL(T3.""Price"",0) ""Price"" from SPP2 T3  " &
                " Where T3.""ItemCode"" = '" & ItemCode & "' and T3.""CardCode"" = '" & Vendor & "' and T3.""UomEntry"" = '" & UOMEntry & "' "

                Dim oRsGetPurchasePrice As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetPurchasePrice.DoQuery(GetPurchasePrice)
                If oRsGetPurchasePrice.RecordCount > 0 Then
                    SalesPrice = CDbl(oRsGetPurchasePrice.Fields.Item(0).Value)
                Else
                    SalesPrice = Me.getSupplierPrice_SpecialPrices(objForm.UniqueID, IsUomDefault, ItemCode, Vendor, UOMEntry, Uom)
                End If

                'If PCPrPack <> "" Then
                '    Packing = PCPrPack & "*" & "1*" & BaseUom
                'Else
                '    Packing = "1*" & BaseUom
                'End If

                Dim GetPackageCode As String = "Select ""U_packing"" FROM OITM Where ""ItemCode""='" & ItemCode.Trim & "'"
                Dim oRsGetPackageCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetPackageCode.DoQuery(GetPackageCode)

                Packing = oRsGetPackageCode.Fields.Item(0).Value


                Dim GetRetailPriceBrnchWise As String = " Select T1.""Price"" From OITM T0 Inner Join ITM1 T1 On T1.""ItemCode"" = T0.""ItemCode"" " &
                                                        " Inner Join OBPL T2 On T1.""PriceList""=T2.""U_PricelistNo"" Where T0.""ItemCode"" = '" & ItemCode & "' " &
                                                        " And T2.""BPLId""= '" & objForm.Items.Item("2001").Specific.Value.Trim & "' "
                Dim oRsGetRetailPriceBrnchWise As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetRetailPriceBrnchWise.DoQuery(GetRetailPriceBrnchWise)
                If oRsGetRetailPriceBrnchWise.RecordCount > 0 Then
                    RetailsPrice = oRsGetRetailPriceBrnchWise.Fields.Item("Price").Value
                End If
                'Dim GetLeastVendorDtlsAndPrice As String = " Select Top 1 T0.""ItemCode"" , T1.""CardCode"", T1.""CardName"" , T0.""Price"" From ITM1 T0 " & _
                '                                           " Inner Join OCRD T1 On T1.""ListNum"" = T0.""PriceList"" And T1.""CardCode"" In (Select ""VendorCode"" From ITM2 Where  ITM2.""ItemCode"" = T0.""ItemCode"") Where " & _
                '                                           " ""ItemCode"" = '" & ItemCode & "' And T1.""CardType"" = 'S' " & _
                '                                           " Order By T0.""Price""  Asc "
                'Dim GetLeastVendorDtlsAndPrice As String = " Select TOP 1 ""Price"", ""BaseCard"" As ""CardCode"" , (Select ""CardName"" From OCRD Where ""CardCode"" = ""BaseCard"" ) As ""CardName"" " & _
                '                                           " From PDN1 Where ""ItemCode""='" & ItemCode & "' And ""DocDate"">(SELECT ADD_DAYS (TO_DATE (CURRENT_DATE, 'YYYY-MM-DD'), -180) From DUMMY) " & _
                '                                           " Order By ""Price"" Asc "
                'Dim GetLeastVendorDtlsAndPrice As String = " Select  TOP 1  Round(Sum(""U_suppprice""/""NumPerMsr""),2)  , ""BaseCard"" As ""CardCode"" , ""DocEntry"" , (Select ""CardName"" From OCRD Where ""CardCode"" = ""BaseCard"" ) As ""CardName"" From PDN1 Where  ""ItemCode""= '" & ItemCode & "'  And " &
                '                                           " ""DocDate"">(Select ADD_DAYS (TO_DATE (CURRENT_DATE, 'YYYY-MM-DD'), -180)  From DUMMY)  " &
                '                                           " Group by ""BaseCard"",""DocEntry"" , ""U_suppprice"" Order By SUM(""U_suppprice""/""NumPerMsr"")"

                Dim GetLeastVendorDtlsAndPrice As String = " Select  TOP 1  Round(Sum(""U_suppprice""/""NumPerMsr""),2)  , ""BaseCard"" As ""CardCode"" , ""DocEntry"" , (Select ""CardName"" From OCRD Where ""CardCode"" = ""BaseCard"" ) As ""CardName"" From PDN1 Where  ""ItemCode""= '" & ItemCode & "'  And " &
                                                           " ""DocDate"">(Select ADD_DAYS (TO_DATE (CURRENT_DATE, 'YYYY-MM-DD'), -45)  From DUMMY)  " &
                                                           " Group by ""BaseCard"",""DocEntry"" , ""U_suppprice"" Order By SUM(""U_suppprice""/""NumPerMsr"")"
                Dim oRsGetLeastVendorDtlsAndPrice As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetLeastVendorDtlsAndPrice.DoQuery(GetLeastVendorDtlsAndPrice)
                If oRsGetLeastVendorDtlsAndPrice.RecordCount > 0 Then
                    LeastVendorPrice = oRsGetLeastVendorDtlsAndPrice.Fields.Item(0).Value
                    LeastVendorCode = oRsGetLeastVendorDtlsAndPrice.Fields.Item("CardCode").Value
                    LeastVendorName = oRsGetLeastVendorDtlsAndPrice.Fields.Item("CardName").Value
                End If
            End If
            If objMatrix.Columns.Item("4").Cells.Item(RowId).Specific.Value = "" Then
                Dim GetBarCode As String = " SELECT Top 1 ""BcdCode"" FROM OBCD WHERE ""UomEntry"" = '" & UOMEntry & "' " &
                                           " And ""ItemCode"" = '" & objMatrix.Columns.Item("1").Cells.Item(RowId).Specific.Value & "' ORDER BY ""BcdEntry"" ASC "
                Dim oRsGetBarCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetBarCode.DoQuery(GetBarCode)
                If oRsGetBarCode.RecordCount > 0 Then
                    AdditionalBarCode = oRsGetBarCode.Fields.Item("BcdCode").Value
                Else
                    Dim GetBarCode1 As String = " SELECT Top 1 ""BcdCode"" FROM OBCD WHERE " &
                                                " ""ItemCode"" = '" & objMatrix.Columns.Item("1").Cells.Item(RowId).Specific.Value & "' ORDER BY ""BcdEntry"" ASC "
                    Dim oRsGetBarCode1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetBarCode1.DoQuery(GetBarCode1)
                    If oRsGetBarCode1.RecordCount > 0 Then
                        AdditionalBarCode = oRsGetBarCode1.Fields.Item("BcdCode").Value
                    End If
                End If
            Else
                AdditionalBarCode = objMatrix.Columns.Item("4").Cells.Item(RowId).Specific.Value
            End If

            If LPP <> 0 Then
                Try
                    objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value = LPP
                Catch ex As Exception
                End Try
            Else
                Try
                    objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value = SalesPrice
                Catch ex As Exception
                End Try
            End If




            Try
                objMatrix.Columns.Item("U_Packing").Cells.Item(RowId).Specific.Value = Packing
            Catch ex As Exception
            End Try
            Try
                objMatrix.Columns.Item("U_RetailPrice").Cells.Item(RowId).Specific.Value = RetailsPrice
            Catch ex As Exception
            End Try
            Try
                objMatrix.Columns.Item("U_Addbarcode").Cells.Item(RowId).Specific.Value = AdditionalBarCode
            Catch ex As Exception
            End Try
            Try
                objMatrix.Columns.Item("U_LVNDPRC").Cells.Item(RowId).Specific.Value = LeastVendorPrice
            Catch ex As Exception
            End Try
            Try
                objMatrix.Columns.Item("U_LPVCD").Cells.Item(RowId).Specific.Value = LeastVendorCode
            Catch ex As Exception
            End Try
            Try
                objMatrix.Columns.Item("U_LPVNM").Cells.Item(RowId).Specific.Value = LeastVendorName
            Catch ex As Exception
            End Try
            Dim NetCost As String = objMatrix.Columns.Item("14").Cells.Item(RowId).Specific.Value
            If NetCost <> "" Then
                NetCost = NetCost.Replace(objForm.Items.Item("63").Specific.Value, "")
                NetCost = Convert.ToDouble(NetCost.Trim)
                Dim LeastPrice As Double = objMatrix.Columns.Item("U_LVNDPRC").Cells.Item(RowId).Specific.Value
                If NetCost > 0 And LeastPrice > 0 Then
                    If NetCost <= LeastPrice Then
                        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LPValidate")
                        oColumn.Cells.Item(RowId).Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    Else
                        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LPValidate")
                        oColumn.Cells.Item(RowId).Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If
                End If
                '-------------------Below cost---------------------------
                DiscPrcnt = 0
                DiscPrcnt = Convert.ToDouble(objMatrix.Columns.Item("15").Cells.Item(RowId).Specific.Value)
                If DiscPrcnt > 0 Then
                    NST = objMatrix.Columns.Item("15").Cells.Item(RowId).Specific.Value
                    NST = NST.Replace(objForm.Items.Item("63").Specific.Value, "")
                    NetCost = Convert.ToDouble(NST.Trim)
                End If

                Dim RetailPrice As Double = objMatrix.Columns.Item("U_RetailPrice").Cells.Item(RowId).Specific.Value
                If RetailPrice > 0 And NetCost > 0 Then
                    'If NetCost > RetailPrice Then
                    '    Try
                    '        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_BelowCst")
                    '        oColumn.Cells.Item(RowId).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)


                    '    Catch ex As Exception

                    '    End Try

                    'End If
                    If NetCost > RetailPrice Then
                        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_BelowCst")
                        oColumn.Cells.Item(RowId).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    Else
                        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_BelowCst")
                        oColumn.Cells.Item(RowId).Specific.Select("No", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If
                End If
            End If
            '-------------------Below cost---------------------------
            objMain.objApplication.StatusBar.SetText("Operation completed successfully....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "UOM Methods"
    Public Function GetUomEntry(ByVal sUomCode As String) As Integer
        Try
            Dim GetDocEntry As String = "Select ""UomEntry"" From OUOM Where ""UomCode"" = '" & sUomCode & "' "
            Dim oRsGetDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDocEntry.DoQuery(GetDocEntry)
            Return CInt(oRsGetDocEntry.Fields.Item("UoMEntry").Value)
        Catch ex As Exception
        End Try
    End Function
    Public Function GetUgpEntry(ByVal sItemCode As String) As Integer
        Try
            Dim GetDocEntry As String = "Select ""UgpEntry"" From OITM Where ""ItemCode"" = '" & sItemCode & "' "
            Dim oRsGetDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDocEntry.DoQuery(GetDocEntry)
            Return CInt(oRsGetDocEntry.Fields.Item("UgpEntry").Value)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
    End Function
    Public Function GetUomCode(ByVal sItemCode As String)
        Try
            Dim GetUomCode1 As String = " Select ""UomCode"" From OUOM Where ""UomEntry"" = " &
                                         " (Select ""IUoMEntry"" From OITM Where ""ItemCode"" = '" & sItemCode & "') "
            Dim oRsGetUomCode1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetUomCode1.DoQuery(GetUomCode1)
            Return oRsGetUomCode1.Fields.Item("UomCode").Value
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function
    Public Function CheckDefaultUOM(ByVal sUomCode As String, ByVal sItemCode As String) As Boolean
        Try
            Dim ChkInventoryUOM As String = " Select ""IUoMEntry"" From OITM  T0 Inner Join OUOM T1 On T1.""UomEntry"" = ""IUoMEntry"" " &
                                            " Where ""ItemCode"" = '" & sItemCode & "' And T1.""UomCode"" = '" & sUomCode & "'"
            Dim oRsChkInventoryUOM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsChkInventoryUOM.DoQuery(ChkInventoryUOM)
            If oRsChkInventoryUOM.RecordCount > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Public Function GetDefaultUomCode(ByVal sItemCode As String)
        Try
            Dim GetUomCode As String = "Select ""UomCode"" From OUOM Where ""UomEntry"" = " &
                                         " (Select ""IUoMEntry"" From OITM Where ""ItemCode"" = '" & sItemCode & "') "
            Dim oRsGetUomCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetUomCode.DoQuery(GetUomCode)
            Return oRsGetUomCode.Fields.Item("UomCode").Value
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function
#End Region

    Private Sub CalculateFOC(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("38").Specific
            oDBs_Details = objForm.DataSources.DBDataSources.Item("PDN1")
            Dim FlagA As String = ""
            Dim FlagB As String = ""
            Dim FlagC As String = ""
            Dim FlagD As String = ""
            Dim FlagE As String = ""
            Dim FlagF As String = ""
            Dim FlagG As String = ""
            Dim FlagH As String = ""
            Dim FlagI As String = ""
            Dim FlagJ As String = ""
            For i As Integer = 1 To objMatrix.VisualRowCount




                Dim NetCost As Double = 0
                Dim NST As String = ""
                Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.ToString().Trim()
                Dim FOCQty As Double = objMatrix.Columns.Item("U_FOCQTY").Cells.Item(i).Specific.Value.ToString().Trim()
                Dim NormalQty As Double = objMatrix.Columns.Item("U_normalqty").Cells.Item(i).Specific.Value.ToString().Trim()
                If (FOCQty + NormalQty) <> 0 Then
                    Try
                        objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value = FOCQty + NormalQty
                    Catch ex As Exception

                    End Try
                    ''' objMatrix.Columns.Item("11").Editable = True
                    'objMatrix.GetLineData(i)
                    'oDBs_Details.SetValue("Quantity", i, FOCQty + NormalQty)
                    'objMatrix.SetLineData(i)
                    'objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value = FOCQty + NormalQty
                    'objMatrix.Columns.Item("14").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    ''objMatrix.Columns.Item("11").Editable = False
                End If
                Dim PurchasePrice As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(i).Specific.Value)
                If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "-" Then
                    Dim FOCTotal As Double = 0
                    Dim Total As Double = PurchasePrice * CDbl(NormalQty)
                    'Dim DiscountTotal As Double = PurchasePrice * CDbl(FOCQty)
                    'FOCTotal = FOCTotal + (FOCQty * PurchasePrice)
                    'Dim Ratio As Double = (FOCTotal / (FOCTotal + Total) * 100)
                    Dim PriceAfterDiscount As Double = Total / (NormalQty + FOCQty)
                    'Dim PriceAfterDiscount As Double = PurchasePrice - ((Ratio / 100) * PurchasePrice)
                    Dim Discount As Double = ((PurchasePrice - PriceAfterDiscount) / 100) * 100
                    'objMatrix.Columns.Item("15").Cells.Item(i).Specific.Value = Discount
                    If FOCQty > 0 Then
                        Try
                            Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                            oColumn.Cells.Item(i).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        Catch ex As Exception
                        End Try
                    End If
                    If PriceAfterDiscount >= 0 Then
                        Try
                            objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value = PriceAfterDiscount
                        Catch ex As Exception

                        End Try

                    End If
                ElseIf objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "A" And FlagA = "" Then
                    Dim NormalTotal As Double = 0
                    Dim FOCTotal As Double = 0
                    Dim TotalQty As Double = 0
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "A" Then
                            Dim FOCQtyA As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyA As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            If FOCQtyA > 0 Then
                                Try
                                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                                    oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Catch ex As Exception
                                End Try
                            End If
                            If (FOCQtyA + NormalQtyA) <> 0 Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyA + NormalQtyA
                            End If
                            Dim PurchasePriceA As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                            FOCTotal = FOCTotal + (FOCQtyA * PurchasePriceA)
                            NormalTotal = NormalTotal + (NormalQtyA * PurchasePriceA)
                            If FOCQtyA > 0 Or NormalQtyA > 0 Then
                                TotalQty = TotalQty + FOCQtyA + NormalQtyA
                            Else
                                objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greate then 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If
                        End If
                    Next
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                        If TotalQty > 0 Then
                            'Dim Ratio As Double = FOCTotal / TotalQty
                            Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                            If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "A" Then
                                Dim FOCQtyA As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                                Dim NormalQtyA As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                                Dim PurchasePriceA As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                                'Dim LineTotalA As Double = ((NormalQtyA + FOCQtyA) * PurchasePriceA) - ((Ratio) * (NormalQtyA + FOCQtyA))
                                'Dim Price As Double = LineTotalA / (NormalQtyA + FOCQtyA)
                                Dim Price As Double = PurchasePriceA - ((Ratio / 100) * PurchasePriceA)
                                If Price > 0 Then
                                    ' objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price
                                    objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price.ToString(System.Globalization.CultureInfo.InvariantCulture)

                                End If
                            End If
                        End If
                    Next
                    FlagA = "Completed"
                ElseIf objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "B" And FlagB = "" Then
                    Dim NormalTotal As Double = 0
                    Dim FOCTotal As Double = 0
                    Dim TotalQty As Double = 0
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "B" Then
                            Dim FOCQtyB As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyB As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            If FOCQtyB > 0 Then
                                Try
                                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                                    oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Catch ex As Exception
                                End Try
                            End If
                            If (FOCQtyB + NormalQtyB) <> 0 Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyB + NormalQtyB
                            End If
                            Dim PurchasePriceB As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                            FOCTotal = FOCTotal + (FOCQtyB * PurchasePriceB)
                            NormalTotal = NormalTotal + (NormalQtyB * PurchasePriceB)
                            If FOCQtyB > 0 Or NormalQtyB > 0 Then
                                TotalQty = TotalQty + FOCQtyB + NormalQtyB
                            Else
                                objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greate then 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If
                        End If
                    Next
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                        If TotalQty > 0 Then
                            'Dim Ratio As Double = FOCTotal / TotalQty
                            Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                            If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "B" Then
                                Dim FOCQtyB As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                                Dim NormalQtyB As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                                Dim PurchasePriceB As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                                'Dim LineTotalB As Double = ((NormalQtyB + FOCQtyB) * PurchasePriceB) - ((Ratio) * (NormalQtyB + FOCQtyB))
                                'Dim Price As Double = LineTotalB / (NormalQtyB + FOCQtyB)
                                Dim Price As Double = PurchasePriceB - ((Ratio / 100) * PurchasePriceB)
                                objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price
                            End If
                        End If
                    Next
                    FlagB = "Completed"
                ElseIf objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "C" And FlagC = "" Then
                    Dim NormalTotal As Double = 0
                    Dim FOCTotal As Double = 0
                    Dim TotalQty As Double = 0
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "C" Then
                            Dim FOCQtyC As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyC As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            If FOCQtyC > 0 Then
                                Try
                                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                                    oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Catch ex As Exception
                                End Try
                            End If
                            If (FOCQtyC + NormalQtyC) <> 0 Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyC + NormalQtyC
                            End If
                            Dim PurchasePriceC As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                            FOCTotal = FOCTotal + (FOCQtyC * PurchasePriceC)
                            NormalTotal = NormalTotal + (NormalQtyC * PurchasePriceC)
                            If FOCQtyC > 0 Or NormalQtyC > 0 Then
                                TotalQty = TotalQty + FOCQtyC + NormalQtyC
                            Else
                                objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greate then 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If
                        End If
                    Next
                    If TotalQty <= 0 Then
                        objMain.objApplication.StatusBar.SetText("All Items total quantity is 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                        'Dim Ratio As Double = FOCTotal / TotalQty
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "C" Then
                            Dim FOCQtyC As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyC As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceC As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                            'Dim LineTotalC As Double = ((NormalQtyC + FOCQtyC) * PurchasePriceC) - ((Ratio) * (NormalQtyC + FOCQtyC))
                            'Dim Price As Double = LineTotalC / (NormalQtyC + FOCQtyC)
                            Dim Price As Double = PurchasePriceC - ((Ratio / 100) * PurchasePriceC)
                            objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price
                        End If
                    Next
                    FlagC = "Completed"
                ElseIf objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "D" And FlagD = "" Then
                    Dim NormalTotal As Double = 0
                    Dim FOCTotal As Double = 0
                    Dim TotalQty As Double = 0
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "D" Then
                            Dim FOCQtyD As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyD As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            If FOCQtyD > 0 Then
                                Try
                                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                                    oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Catch ex As Exception
                                End Try
                            End If
                            If (FOCQtyD + NormalQtyD) <> 0 Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyD + NormalQtyD
                            End If
                            Dim PurchasePriceD As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                            FOCTotal = FOCTotal + (FOCQtyD * PurchasePriceD)
                            NormalTotal = NormalTotal + (NormalQtyD * PurchasePriceD)
                            If FOCQtyD > 0 Or NormalQtyD > 0 Then
                                TotalQty = TotalQty + FOCQtyD + NormalQtyD
                            Else
                                objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greate then 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If
                        End If
                    Next
                    If TotalQty <= 0 Then
                        objMain.objApplication.StatusBar.SetText("All Items total quantity is 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                        'Dim Ratio As Double = FOCTotal / TotalQty
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "D" Then
                            Dim FOCQtyD As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyD As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceD As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                            'Dim LineTotalD As Double = ((NormalQtyD + FOCQtyD) * PurchasePriceD) - ((Ratio) * (NormalQtyD + FOCQtyD))
                            'Dim Price As Double = LineTotalD / (NormalQtyD + FOCQtyD)
                            Dim Price As Double = PurchasePriceD - ((Ratio / 100) * PurchasePriceD)
                            objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price
                        End If
                    Next
                    FlagD = "Completed"
                ElseIf objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "E" And FlagE = "" Then
                    Dim NormalTotal As Double = 0
                    Dim FOCTotal As Double = 0
                    Dim TotalQty As Double = 0
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "E" Then

                            Dim FOCQtyE As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyE As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            If FOCQtyE > 0 Then
                                Try
                                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                                    oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Catch ex As Exception
                                End Try
                            End If
                            If (FOCQtyE + NormalQtyE) <> 0 Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyE + NormalQtyE
                            End If
                            Dim PurchasePriceE As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                            FOCTotal = FOCTotal + (FOCQtyE * PurchasePriceE)
                            NormalTotal = NormalTotal + (NormalQtyE * PurchasePriceE)
                            If FOCQtyE > 0 Or NormalQtyE > 0 Then
                                TotalQty = TotalQty + FOCQtyE + NormalQtyE
                            Else
                                objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greate then 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If
                        End If
                    Next
                    If TotalQty <= 0 Then
                        objMain.objApplication.StatusBar.SetText("All Items total quantity is 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                        'Dim Ratio As Double = FOCTotal / TotalQty
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "E" Then
                            Dim FOCQtyE As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyE As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceE As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                            'Dim LineTotalE As Double = ((NormalQtyE + FOCQtyE) * PurchasePriceE) - ((Ratio) * (NormalQtyE + FOCQtyE))
                            'Dim Price As Double = LineTotalE / (NormalQtyE + FOCQtyE)
                            Dim Price As Double = PurchasePriceE - ((Ratio / 100) * PurchasePriceE)
                            objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price
                        End If
                    Next
                    FlagE = "Completed"
                ElseIf objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "F" And FlagF = "" Then
                    Dim NormalTotal As Double = 0
                    Dim FOCTotal As Double = 0
                    Dim TotalQty As Double = 0
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "F" Then

                            Dim FOCQtyF As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyF As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            If FOCQtyF > 0 Then
                                Try
                                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                                    oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Catch ex As Exception
                                End Try
                            End If
                            If (FOCQtyF + NormalQtyF) <> 0 Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyF + NormalQtyF
                            End If
                            Dim PurchasePriceF As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                            FOCTotal = FOCTotal + (FOCQtyF * PurchasePriceF)
                            NormalTotal = NormalTotal + (NormalQtyF * PurchasePriceF)
                            If FOCQtyF > 0 Or NormalQtyF > 0 Then
                                TotalQty = TotalQty + FOCQtyF + NormalQtyF
                            Else
                                objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greate then 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If
                        End If
                    Next
                    If TotalQty <= 0 Then
                        objMain.objApplication.StatusBar.SetText("All Items total quantity is 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                        'Dim Ratio As Double = FOCTotal / TotalQty
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "F" Then
                            Dim FOCQtyF As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyF As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceF As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                            'Dim LineTotalF As Double = ((NormalQtyF + FOCQtyF) * PurchasePriceF) - ((Ratio) * (NormalQtyF + FOCQtyF))
                            'Dim Price As Double = LineTotalF / (NormalQtyF + FOCQtyF)
                            Dim Price As Double = PurchasePriceF - ((Ratio / 100) * PurchasePriceF)
                            objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price
                        End If
                    Next
                    FlagF = "Completed"
                ElseIf objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "G" And FlagG = "" Then
                    Dim NormalTotal As Double = 0
                    Dim FOCTotal As Double = 0
                    Dim TotalQty As Double = 0
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "G" Then
                            Dim FOCQtyG As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyG As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())

                            If FOCQtyG > 0 Then
                                Try
                                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                                    oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Catch ex As Exception
                                End Try
                            End If

                            If (FOCQtyG + NormalQtyG) <> 0 Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyG + NormalQtyG
                            End If
                            Dim PurchasePriceG As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                            FOCTotal = FOCTotal + (FOCQtyG * PurchasePriceG)
                            NormalTotal = NormalTotal + (NormalQtyG * PurchasePriceG)
                            If FOCQtyG > 0 Or NormalQtyG > 0 Then
                                TotalQty = TotalQty + FOCQtyG + NormalQtyG
                            Else
                                objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greate then 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If
                        End If
                    Next
                    If TotalQty <= 0 Then
                        objMain.objApplication.StatusBar.SetText("All Items total quantity is 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                        'Dim Ratio As Double = FOCTotal / TotalQty
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "G" Then
                            Dim FOCQtyG As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyG As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceG As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                            'Dim LineTotalG As Double = ((NormalQtyG + FOCQtyG) * PurchasePriceG) - ((Ratio) * (NormalQtyG + FOCQtyG))
                            'Dim Price As Double = LineTotalG / (NormalQtyG + FOCQtyG)
                            Dim Price As Double = PurchasePriceG - ((Ratio / 100) * PurchasePriceG)
                            objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price
                        End If
                    Next
                    FlagG = "Completed"
                ElseIf objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "H" And FlagH = "" Then
                    Dim NormalTotal As Double = 0
                    Dim FOCTotal As Double = 0
                    Dim TotalQty As Double = 0
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "H" Then
                            Dim FOCQtyH As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyH As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            If FOCQtyH > 0 Then
                                Try
                                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                                    oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Catch ex As Exception
                                End Try
                            End If
                            If (FOCQtyH + NormalQtyH) <> 0 Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyH + NormalQtyH
                            End If
                            Dim PurchasePriceH As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                            FOCTotal = FOCTotal + (FOCQtyH * PurchasePriceH)
                            NormalTotal = NormalTotal + (NormalQtyH * PurchasePriceH)
                            If FOCQtyH > 0 Or NormalQtyH > 0 Then
                                TotalQty = TotalQty + FOCQtyH + NormalQtyH
                            Else
                                objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greate then 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If
                        End If
                    Next
                    If TotalQty <= 0 Then
                        objMain.objApplication.StatusBar.SetText("All Items total quantity is 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                        'Dim Ratio As Double = FOCTotal / TotalQty
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "H" Then
                            Dim FOCQtyH As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyH As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceH As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                            'Dim LineTotalG As Double = ((NormalQtyH + FOCQtyH) * PurchasePriceH) - ((Ratio) * (NormalQtyH + FOCQtyH))
                            'Dim Price As Double = LineTotalG / (NormalQtyH + FOCQtyH)
                            Dim Price As Double = PurchasePriceH - ((Ratio / 100) * PurchasePriceH)
                            objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price
                        End If
                    Next
                    FlagH = "Completed"
                ElseIf objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "I" And FlagI = "" Then
                    Dim NormalTotal As Double = 0
                    Dim FOCTotal As Double = 0
                    Dim TotalQty As Double = 0
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "I" Then
                            Dim FOCQtyI As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyI As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            If FOCQtyI > 0 Then
                                Try
                                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                                    oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Catch ex As Exception
                                End Try
                            End If
                            If (FOCQtyI + NormalQtyI) <> 0 Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyI + NormalQtyI
                            End If
                            Dim PurchasePriceI As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                            FOCTotal = FOCTotal + (FOCQtyI * PurchasePriceI)
                            NormalTotal = NormalTotal + (NormalQtyI * PurchasePriceI)
                            If FOCQtyI > 0 Or NormalQtyI > 0 Then
                                TotalQty = TotalQty + FOCQtyI + NormalQtyI
                            Else
                                objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greate then 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If
                        End If
                    Next
                    If TotalQty <= 0 Then
                        objMain.objApplication.StatusBar.SetText("All Items total quantity is 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                        'Dim Ratio As Double = FOCTotal / TotalQty
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "I" Then

                            Dim FOCQtyI As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyI As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceI As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                            'Dim LineTotalI As Double = ((NormalQtyI + FOCQtyI) * PurchasePriceI) - ((Ratio) * (NormalQtyI + FOCQtyI))
                            'Dim Price As Double = LineTotalI / (NormalQtyI + FOCQtyI)
                            Dim Price As Double = PurchasePriceI - ((Ratio / 100) * PurchasePriceI)
                            objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price
                        End If
                    Next
                    FlagI = "Completed"
                ElseIf objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "J" And FlagJ = "" Then
                    Dim NormalTotal As Double = 0
                    Dim FOCTotal As Double = 0
                    Dim TotalQty As Double = 0
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "J" Then
                            Dim FOCQtyJ As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyJ As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            If FOCQtyJ > 0 Then
                                Try
                                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                                    oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Catch ex As Exception
                                End Try
                            End If
                            If (FOCQtyJ + NormalQtyJ) <> 0 Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyJ + NormalQtyJ
                            End If
                            Dim PurchasePriceJ As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                            FOCTotal = FOCTotal + (FOCQtyJ * PurchasePriceJ)
                            NormalTotal = NormalTotal + (NormalQtyJ * PurchasePriceJ)
                            If FOCQtyJ > 0 Or NormalQtyJ > 0 Then
                                TotalQty = TotalQty + FOCQtyJ + NormalQtyJ
                            Else
                                objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greate then 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If
                        End If
                    Next
                    If TotalQty <= 0 Then
                        objMain.objApplication.StatusBar.SetText("All Items total quantity is 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                        'Dim Ratio As Double = FOCTotal / TotalQty
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "J" Then
                            Dim FOCQtyJ As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyJ As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceJ As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                            'Dim LineTotalJ As Double = ((NormalQtyJ + FOCQtyJ) * PurchasePriceJ) - ((Ratio) * (NormalQtyJ + FOCQtyJ))
                            'Dim Price As Double = LineTotalJ / (NormalQtyJ + FOCQtyJ)
                            Dim Price As Double = PurchasePriceJ - ((Ratio / 100) * PurchasePriceJ)
                            objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price
                        End If
                    Next
                    FlagJ = "Completed"
                End If

                'MARGIN 
                'If objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value = "" Then
                '    NetCost = 0
                'Else
                '    NST = objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value
                '    NST = NST.Replace(objForm.Items.Item("63").Specific.Value, "")

                '    'NetCost = objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value
                '    'NetCost = NetCost.Replace(objForm.Items.Item("63").Specific.Value, "")
                '    NetCost = Convert.ToDouble(NST.Trim)
                'End If

                'DiscPrcnt = 0
                'DiscPrcnt = Convert.ToDouble(objMatrix.Columns.Item("15").Cells.Item(i).Specific.Value)
                'If DiscPrcnt > 0 Then
                '    NST = objMatrix.Columns.Item("15").Cells.Item(i).Specific.Value
                '    NST = NST.Replace(objForm.Items.Item("63").Specific.Value, "")
                '    NetCost = Convert.ToDouble(NST.Trim)
                'End If

                'Dim RetailPrice As Double = objMatrix.Columns.Item("U_RetailPrice").Cells.Item(i).Specific.Value

                'If RetailPrice > 0 And NetCost > 0 Then
                '    If NetCost > RetailPrice Then
                '        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_BelowCst")
                '        oColumn.Cells.Item(i).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                '    Else
                '        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_BelowCst")
                '        oColumn.Cells.Item(i).Specific.Select("No", SAPbouiCOM.BoSearchKey.psk_ByValue)
                '    End If
                'End If

                'Try
                '    Dim RetailPrice1 As Double = 0
                '    Dim SuppPrice As Double = 0


                '    If CStr(objMatrix.Columns.Item("U_RetailPrice").Cells.Item(i).Specific.Value).Trim <> "" Then
                '        RetailPrice1 = CDbl(CStr(objMatrix.Columns.Item("U_RetailPrice").Cells.Item(i).Specific.Value).Trim)
                '    End If

                '    If CStr(objMatrix.Columns.Item("U_suppprice").Cells.Item(i).Specific.Value).Trim <> "" Then
                '        SuppPrice = CDbl(CStr(objMatrix.Columns.Item("U_suppprice").Cells.Item(i).Specific.Value).Trim)
                '    End If

                '    objMatrix.Columns.Item("U_MARGN").Cells.Item(i).Specific.Value = ((RetailPrice1 - SuppPrice) / RetailPrice1) * 100

                'Catch ex As Exception

                'End Try

                'MARGIN
                Try

                    objMatrix.Columns.Item("15").Cells.Item(i).Specific.Value = 0

                    Dim DA As Double = 0

                    Dim Tot As Double = 0

                    If CStr(objMatrix.Columns.Item("U_DA").Cells.Item(i).Specific.Value).Trim <> "" Then

                        DA = CDbl(CStr(objMatrix.Columns.Item("U_DA").Cells.Item(i).Specific.Value).Trim)
                    End If

                    If CStr(objMatrix.Columns.Item("21").Cells.Item(i).Specific.Value).Trim <> "" Then

                        'Regex.Match(subjectString, @"\d+").Value  (str, "\d+(?:[.,]\d+)*")

                        Dim rawValue As String = CStr(objMatrix.Columns.Item("21").Cells.Item(i).Specific.Value).Trim()
                        rawValue = rawValue.Replace(",", "").Replace(" ", "") ' Remove thousands separators and spaces

                        Dim pattern As String = "\d+(\.\d{1,4})?"
                        Dim match As Match = Regex.Match(rawValue, pattern)

                        If match.Success Then
                            Tot = CDbl(match.Value)
                        End If
                        'commented by manohar 2000 taking as 2
                        'Dim pattern As String = "\d+(\.\d{1,4})?"

                        '     Dim match As Match = Regex.Match(CStr(objMatrix.Columns.Item("21").Cells.Item(i).Specific.Value).Trim, pattern)

                        ''Tot = CDbl(CStr(objMatrix.Columns.Item("288").Cells.Item(i).Specific.Value).Trim)

                        'Tot = match.Value


                    End If

                    objMatrix.Columns.Item("15").Cells.Item(i).Specific.Value = (DA / Tot) * 100

                Catch ex As Exception

                End Try


            Next
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub LineValidations(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim objMatrix As SAPbouiCOM.Matrix
            objMatrix = objForm.Items.Item("38").Specific

            For i As Integer = 1 To objMatrix.VisualRowCount - 1
                Dim NetCost As String = objMatrix.Columns.Item("U_suppprice").Cells.Item(i).Specific.Value
                If NetCost <> "" Then
                    NetCost = NetCost.Replace(objForm.Items.Item("63").Specific.Value, "")
                    NetCost = Convert.ToDouble(NetCost)
                    Dim LeastPrice As Double = objMatrix.Columns.Item("U_LVNDPRC").Cells.Item(i).Specific.Value
                    If LeastPrice = 0 Then
                        If objMatrix.Columns.Item("U_LVNDPRC").Visible = True Then

                            Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LPValidate")
                            oColumn.Cells.Item(i).Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        End If
                    End If
                    If NetCost > 0 And LeastPrice > 0 Then
                        If NetCost > LeastPrice Then
                            Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LPValidate")
                            oColumn.Cells.Item(i).Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        Else
                            Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LPValidate")
                            oColumn.Cells.Item(i).Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        End If
                    End If
                    '-------------------Below cost---------------------------
                    'NetCost = objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value
                    'NetCost = NetCost.Replace(objForm.Items.Item("63").Specific.Value, "")
                    'NetCost = Convert.ToDouble(NetCost.Trim)

                    If objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value = "" Then
                        NetCost = 0
                    Else
                        NetCost = objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value
                        NetCost = NetCost.Replace(objForm.Items.Item("63").Specific.Value, "")
                        NetCost = Convert.ToDouble(NetCost.Trim)
                    End If

                    DiscPrcnt = 0
                    DiscPrcnt = Convert.ToDouble(objMatrix.Columns.Item("15").Cells.Item(i).Specific.Value)
                    If DiscPrcnt > 0 Then
                        NST = objMatrix.Columns.Item("15").Cells.Item(i).Specific.Value
                        NST = NST.Replace(objForm.Items.Item("63").Specific.Value, "")
                        NetCost = Convert.ToDouble(NST.Trim)
                    End If

                    Dim RetailPrice As Double = objMatrix.Columns.Item("U_RetailPrice").Cells.Item(i).Specific.Value
                    If RetailPrice > 0 And NetCost > 0 Then
                        'If NetCost > RetailPrice Then
                        '    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_BelowCst")
                        '    oColumn.Cells.Item(i).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        'End If
                        If NetCost > RetailPrice Then
                            Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_BelowCst")
                            oColumn.Cells.Item(i).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        Else
                            Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_BelowCst")
                            oColumn.Cells.Item(i).Specific.Select("No", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("38").Specific
            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                If objMatrix.VisualRowCount > 0 Then
                    If objMatrix.Columns.Item("44").Cells.Item(1).Specific.Value = "" Then
                        If CalcualtionBtnFlag = False Then
                            Dim PObutton As String = "Select ""U_GRPO""as ""GRPO"" from ""@ME_ACFG"""
                            Dim rsPObutton As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            rsPObutton.DoQuery(PObutton)
                            'If rsPObutton.Fields.Item("GRPO").Value = "Y" Then
                            '    '    Exit Function
                            '    'Else
                            '    objMain.objApplication.StatusBar.SetText("Please run the calculation before generating document.....!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            '    Return False
                            'End If
                        End If
                    End If
                End If
            End If


            For i As Integer = 1 To objMatrix.VisualRowCount - 1
                Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.Trim
                Dim NormalQty As Double = objMatrix.Columns.Item("U_normalqty").Cells.Item(i).Specific.Value.Trim
                Dim FOCQty As Double = objMatrix.Columns.Item("U_FOCQTY").Cells.Item(i).Specific.Value.Trim

                Dim Qty As Double = objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value.Trim

                Dim GrsQt As Double = 0

                If CStr(objMatrix.Columns.Item("U_GROSSQTY").Cells.Item(i).Specific.Value).Trim <> "" Then
                    GrsQt = CDbl(objMatrix.Columns.Item("U_GROSSQTY").Cells.Item(i).Specific.Value)
                End If

                If (Qty <> (NormalQty + FOCQty)) And GrsQt = 0 Then
                    objMain.objApplication.StatusBar.SetText("For Item " & ItemCode & " And Line No. " & i & "....OverAll Qty should balance NormalQty And FOCQty.....!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

            Next
            For i As Integer = 1 To objMatrix.VisualRowCount - 1
                Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.Trim
                Dim PurchaseNum As String = objMatrix.Columns.Item("45").Cells.Item(objMatrix.VisualRowCount - 1).Specific.Value.ToString()
                Dim LineNum As String = objMatrix.Columns.Item("0").Cells.Item(i).Specific.Value

                Dim qStr As String = "SELECT DISTINCT A.""U_LOOSPER"" FROM OITM A INNER JOIN OITB B ON B.""ItmsGrpCod""=A.""ItmsGrpCod"" " +
                                                    " AND ""ItmsGrpNam""='FROZEN' AND A.""ItemCode""='" + ItemCode + "'"
                Dim rSet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rSet.DoQuery(qStr)
                If rSet.RecordCount > 0 Then
                    Dim GrossQty As Double = CDbl(objMatrix.Columns.Item("U_GROSSQTY").Cells.Item(i).Specific.Value)
                    Dim GrossTotal As Double = CDbl(objMatrix.Columns.Item("U_GROSSTOT").Cells.Item(i).Specific.Value)
                    If GrossQty <> 0 And GrossTotal <> 0 Then
                    Else
                        objMain.objApplication.StatusBar.SetText("Kindly enter GrossQty and GrossTotal in Line No " + i.ToString() + "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            Next
            For i As Integer = 1 To objMatrix.VisualRowCount - 1
                Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.Trim
                Dim PurchaseNum As String = objMatrix.Columns.Item("45").Cells.Item(objMatrix.VisualRowCount - 1).Specific.Value.ToString()
                'Dim LineNum As String = objMatrix.Columns.Item("0").Cells.Item(i).Specific.Value
                'Dim BaseLine As Integer = objMatrix.Columns.Item("46").Cells.Item(objMatrix.VisualRowCount - 1).Specific.Value.ToString() ''(objMatrix.Columns.Item("46").Cells.Item(i).Specific.Value)

                If String.IsNullOrWhiteSpace(PurchaseNum) Then
                    Continue For
                End If

                Dim BaseLineStr As String = objMatrix.Columns.Item("46").Cells.Item(i).Specific.Value.Trim()

                If String.IsNullOrWhiteSpace(BaseLineStr) Then
                    Continue For
                End If

                Dim BaseLine As Integer = CInt(BaseLineStr)

                If PurchaseNum.Trim = "" Then
                    GoTo NextStep
                End If

                Dim oRsBranch As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim Query As String = " Select T1.""ItemCode"",T1.""Quantity"",T1.""OpenQty"",T1.""LineNum"",T1.""U_normalqty""as ""Normal"",T1.""U_FOCQTY""as ""FOCQty"",T0.""DocNum"" From OPOR T0  Inner Join " &
                                                                        "POR1 T1 On T1.""DocEntry"" = T0.""DocEntry"" Where T0.""DocEntry"" = '" & PurchaseNum & "'and T1.""LineStatus""='O'and T1.""LineNum""= '" & BaseLine & "'"
                oRsBranch.DoQuery(Query)
                If oRsBranch.RecordCount > 0 Then
                    Dim PONormalQty As Double = oRsBranch.Fields.Item("Normal").Value
                    Dim POFOCQty As Double = oRsBranch.Fields.Item("FOCQty").Value
                    Dim NormalQty As Double = objMatrix.Columns.Item("U_normalqty").Cells.Item(i).Specific.Value.Trim
                    Dim FOCQty As Double = objMatrix.Columns.Item("U_FOCQTY").Cells.Item(i).Specific.Value.Trim
                    If POFOCQty > 0 Then
                        If PONormalQty <> NormalQty Then
                            objMain.objApplication.StatusBar.SetText("For Item " & ItemCode & " And Line No. " & i & ".... NormalQty And FOCQty should be same as Purchase Order " & oRsBranch.Fields.Item("DocNum").Value & "..!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                        If POFOCQty <> FOCQty Then
                            objMain.objApplication.StatusBar.SetText("For Item " & ItemCode & " And Line No. " & i & ".... NormalQty And FOCQty should be same as Purchase Order  " & oRsBranch.Fields.Item("DocNum").Value & "....!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                End If


NextStep:       Next
                Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_SplitKitItems")
                For i As Integer = 1 To objMatrix.VisualRowCount
                    Dim ChkKitItems As String = oColumn.Cells.Item(i).Specific.Value
                    If ChkKitItems = "Yes" Then
                        Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.Trim
                        Dim UOM As String = objMatrix.Columns.Item("1470002145").Cells.Item(i).Specific.Value.Trim
                        Dim chkBaseUom As String = "Select ""InvntryUom"" From OITM Where ""ItemCode"" =  '" & ItemCode & "'"
                        Dim oRschkBaseUom As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRschkBaseUom.DoQuery(chkBaseUom)
                        If oRschkBaseUom.RecordCount > 0 Then
                            If oRschkBaseUom.Fields.Item("InvntryUom").Value <> UOM Then
                                objMain.objApplication.StatusBar.SetText("For Item " & ItemCode & " and Line No. " & i & "....  please provide Base Uom !!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                        End If
                    End If
                Next

                Try

                    Dim CheckBackDOc As String = "Select ""U_GUID"" FROM ""@TNXEDS"" Where " &
                                                         """U_GUID""='" & objForm.Items.Item("txt_GUID").Specific.Value & "' And ""U_DOCTY""='OPDN'"

                    Dim oRsCheckBackDOc As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    oRsCheckBackDOc.DoQuery(CheckBackDOc)

                    Dim GetWCnd As String = "select IFNULL(""U_EXCHK"",'No') FROM OWHS Where ""WhsCode""='" & objMatrix.Columns.Item("24").Cells.Item(1).Specific.Value & "' And ""U_EXCHK""='Yes' And ""U_EXCHK"" IS NOT NULL And ""U_EXCHK""<>''"
                    Dim oRsGetWCnd As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetWCnd.DoQuery(GetWCnd)



                    If CStr(objForm.Items.Item("3").Specific.Value).Trim = "I" And oRsGetWCnd.RecordCount > 0 And (oRsCheckBackDOc.RecordCount = 0 Or CStr(CStr(objForm.Items.Item("txt_GUID").Specific.Value).Trim).Trim = "") Then
                        objMain.objApplication.StatusBar.SetText("please add Expiry Screen before adding")
                        objForm.Freeze(False)
                        Return False
                        Exit Try
                    End If



                    If oRsGetWCnd.RecordCount > 0 And (oRsCheckBackDOc.RecordCount = 0 Or CStr(CStr(objForm.Items.Item("txt_GUID").Specific.Value).Trim).Trim = "") Then
                        objMain.objApplication.StatusBar.SetText("please add Expiry Screen before adding")
                        objForm.Freeze(False)
                        Return False


                        Exit Try
                    End If



                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText(ex.Message)
                End Try

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Private Sub LoadOnItemOrUOMSelection_LostFocus(ByVal FormUID As String, ByVal RowId As Integer, ByVal IsUomDefault As String, ByVal ItemCode As String, ByVal Uom As String, ByVal Vendor As String, ByVal [Date] As String, ByVal BarCode As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("OPDN")
            objMatrix = objForm.Items.Item("38").Specific
            Dim SalesPrice As Double = 0
            Dim SppPrice As Double = 0
            Dim Packing As String = ""
            Dim RetailsPrice As Double = 0
            Dim AdditionalBarCode As String = ""
            Dim UOMEntry As Integer = 0
            Dim LeastVendorPrice As Double = 0
            Dim LeastVendorCode As String = ""
            Dim LeastVendorName As String = ""
            Dim LeastVendorValidation As String = ""
            Dim PCPrPack As String = objMain.objUtilities.GetPCPerItem(ItemCode)
            BaseUom = objMain.objUtilities.GetItemBaseUom(ItemCode)
            ''
            If objForm.Items.Item("2001").Specific.Value.Trim = "" Then
                objMain.objApplication.StatusBar.SetText("Please select branch at header...")
                Exit Sub
            End If


            If IsUomDefault = "No" Then
                UOMEntry = Me.GetUomEntry(Uom)
            ElseIf IsUomDefault = "Yes" Then
                UOMEntry = Me.GetUomEntry(Me.GetDefaultUomCode(ItemCode))
            End If

            Try

                Dim GetBase As String = " SELECT T0.""U_SUTAXAMT"" AS ""Price"" FROM ""OBCD"" T0 LEFT JOIN OITM T1 ON T1.""ItemCode""=T0.""ItemCode"" 
 WHERE T0.""ItemCode"" =  '" & ItemCode & "' and T0.""UomEntry"" = '" & UOMEntry & "' and T1.""U_SugarTax""='Y'"

                Dim oRsGetbase As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetbase.DoQuery(GetBase)



                objMatrix.Columns.Item("U_BASE").Cells.Item(RowId).Specific.Value = oRsGetbase.Fields.Item("Price").Value

                Dim Discount As Double = objMatrix.Columns.Item("U_DSAT").Cells.Item(RowId).Specific.Value
                Dim baseVal As Double = 0
                Dim discountVal As Double = 0
                Dim badVal As Double = 0
                Dim extVal As Double = 0

                Double.TryParse(objMatrix.Columns.Item("U_BASE").Cells.Item(RowId).Specific.Value, baseVal)
                Double.TryParse(Discount, discountVal)

                ' U_BAD = U_BASE - Discount
                badVal = baseVal - discountVal
                objMatrix.Columns.Item("U_BAD").Cells.Item(RowId).Specific.Value = badVal.ToString("0.00")

                ' U_EXT = 50% of U_BAD
                extVal = (baseVal * 50) / 100
                objMatrix.Columns.Item("U_EXT").Cells.Item(RowId).Specific.Value = extVal.ToString("0.00")

                ' Column 14 (Unit Price) = U_BAD + U_EXT
                objMatrix.Columns.Item("14").Cells.Item(RowId).Specific.Value = (badVal + extVal).ToString("0.00")

                Dim GetSPP1 As String = "SELECT T0.""ItemCode"" AS ""Item No."",
	 T4.""ItemName"" AS ""Description"",
	COALESCE(T2.""Price"",
	 T1.""Price"",
	 T0.""Price"") AS ""Effective Price"",
	 COALESCE(T2.""Discount"",
	 T1.""Discount"",
	 T0.""Discount"") AS ""Applied Discount %"" 
FROM ""OSPP"" T0 
LEFT JOIN ""SPP1"" T1 ON T0.""CardCode"" = T1.""CardCode"" 
AND T0.""ItemCode"" = T1.""ItemCode"" 
LEFT JOIN ""SPP2"" T2 ON T1.""CardCode"" = T2.""CardCode"" 
AND T1.""ItemCode"" = T2.""ItemCode"" 
INNER JOIN ""OCRD"" T3 ON T0.""CardCode"" = T3.""CardCode"" 
INNER JOIN ""OITM"" T4 ON T0.""ItemCode"" = T4.""ItemCode"" 
WHERE T0.""CardCode"" = '" & Vendor & "' and T0.""ItemCode""='" & ItemCode & "' and T2.""UomEntry"" = '" & UOMEntry & "'
AND TO_DATE('" & [Date] & "', 'YYYYMMDD') BETWEEN T1.""FromDate"" AND T1.""ToDate""
ORDER BY T0.""CardCode"",T0.""ItemCode"",T1.""FromDate"""

                Dim oRsGetSPP1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetSPP1.DoQuery(GetSPP1)

                SppPrice = oRsGetSPP1.Fields.Item("Effective Price").Value


                '            Dim GetSPP As String = "SELECT T1.""Price"" FROM OITM T0 Left JOIN ITM1 T1 ON T0.""ItemCode"" = T1.""ItemCode""
                'WHERE T0.""ItemCode"" = '" & ItemCode & "' AND T1.""PriceList"" = '11'"

                Dim GetSPP As String = "CALL ""TNX_GETITEMPRICE""('" & BarCode & "', " & UOMEntry & ", '" & ItemCode & "')"
                Dim oRsGetSPP As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetSPP.DoQuery(GetSPP)



                Dim GetSugar As String = "SELECT T0.""U_SugarTax"" FROM OITM T0 WHERE T0.""ItemCode"" = '" & ItemCode & "'"

                Dim oRsGetSugar As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetSugar.DoQuery(GetSugar)

                If oRsGetSugar.Fields.Item("U_SugarTax").Value = "Y" Then
                    objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value = (badVal + extVal).ToString("0.00")
                Else
                    If SppPrice <> 0 Then
                        objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value = SppPrice
                    Else
                        objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value = oRsGetSPP.Fields.Item("Price").Value
                    End If

                End If


                ''objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value = SalesPrice
            Catch ex As Exception
            End Try

            '-------------------Below cost---------------------------
            ''objMain.objApplication.StatusBar.SetText("Operation completed successfully....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub LoadOnItemOrUOMSelection_SpecialPrices(ByVal FormUID As String, ByVal RowId As Integer, ByVal IsUomDefault As String, ByVal ItemCode As String, ByVal Uom As String, ByVal Vendor As String, ByVal [Date] As String, ByVal BarCode As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("OPDN")
            objMatrix = objForm.Items.Item("38").Specific
            Dim SalesPrice As Double = 0
            Dim SppPrice As Double = 0
            Dim Packing As String = ""
            Dim RetailsPrice As Double = 0
            Dim AdditionalBarCode As String = ""
            Dim UOMEntry As Integer = 0
            Dim LeastVendorPrice As Double = 0
            Dim LeastVendorCode As String = ""
            Dim LeastVendorName As String = ""
            Dim LeastVendorValidation As String = ""
            Dim PCPrPack As String = objMain.objUtilities.GetPCPerItem(ItemCode)
            BaseUom = objMain.objUtilities.GetItemBaseUom(ItemCode)
            ''
            If objForm.Items.Item("2001").Specific.Value.Trim = "" Then
                objMain.objApplication.StatusBar.SetText("Please select branch at header...")
                Exit Sub
            End If

            If IsUomDefault = "No" Then
                UOMEntry = Me.GetUomEntry(Uom)
                'Dim GetPurchasePrice As String = " Select T0.""Price"" From ITM9 T0 Where T0.""PriceList""= " & _
                '                                 " (Select ""ListNum"" From OCRD Where ""CardCode"" = '" & Vendor & "') And " & _
                '                                 " T0.""UomEntry""='" & UOMEntry & "' And ""ItemCode""='" & ItemCode & "' "

                Dim GetPurchasePrice As String = " Select IFNULL(T3.""Price"",0) ""Price"" from SPP2 T3  " &
                " Where T3.""ItemCode"" = '" & ItemCode & "' and T3.""CardCode"" = '" & Vendor & "' and T3.""UomEntry"" = '" & UOMEntry & "' "

                Dim oRsGetPurchasePrice As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetPurchasePrice.DoQuery(GetPurchasePrice)
                If oRsGetPurchasePrice.RecordCount > 0 Then
                    SalesPrice = CDbl(oRsGetPurchasePrice.Fields.Item(0).Value)
                End If
                Dim GetPackageCode As String = " Select Cast(Cast(T0.""BaseQty"" As Int) As Varchar)" & Concate & "'*'" & Concate & "'" & BaseUom & "'" &
                                               " From UGP1 T0 Inner Join OUOM T1 On  " &
                                               " T1.""UomEntry"" = T0.""UomEntry""  Where ""UgpEntry"" = '" & Me.GetUgpEntry(ItemCode) & "' And  " &
                                               " T1.""UomCode"" = '" & Uom & "' "
                Dim oRsGetPackageCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetPackageCode.DoQuery(GetPackageCode)
                Packing = oRsGetPackageCode.Fields.Item(0).Value
                If PCPrPack <> "" Then
                    Packing = PCPrPack & "*" & oRsGetPackageCode.Fields.Item(0).Value
                End If
                'Dim GetRetailPriceBrnchWise As String = " Select T1.""Price"" From OITM T0 Inner Join ITM9 T1 On T1.""ItemCode"" = T0.""ItemCode"" " & _
                '                                        " Inner Join OBPL T2 On T1.""PriceList""=T2.""U_PricelistNo"" Where T0.""ItemCode"" = '" & ItemCode & "' " & _
                '                                        " And T2.""BPLId""= '" & objForm.Items.Item("2001").Specific.Value.Trim & "' And T1.""UomEntry""='" & UOMEntry & "'"

                Dim GetRetailPriceBrnchWise As String = " Select T0.""Price"" From SPP2 T0 Inner Join ITM1 T1 On T1.""ItemCode"" = T0.""ItemCode"" " &
                " Inner Join OBPL T2 On T1.""PriceList"" = T2.""U_PricelistNo"" " &
                " Where T0.""ItemCode"" = '" & ItemCode & "'  And T2.""BPLId""= '" & objForm.Items.Item("2001").Specific.Value.Trim & "' And T0.""UomEntry""='" & UOMEntry & "' "

                Dim oRsGetRetailPriceBrnchWise As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetRetailPriceBrnchWise.DoQuery(GetRetailPriceBrnchWise)
                If oRsGetRetailPriceBrnchWise.RecordCount > 0 Then
                    RetailsPrice = oRsGetRetailPriceBrnchWise.Fields.Item("Price").Value
                End If
                'Dim GetLeastVendorDtlsAndPrice As String = " Select Top 1 T0.""ItemCode"" , T1.""CardCode"", T1.""CardName"" , T0.""Price"" From ITM9 T0 " & _
                '                                           " Inner Join OCRD T1 On T1.""ListNum"" = T0.""PriceList""  And T1.""CardCode"" In (Select ""VendorCode"" From ITM2 Where  ITM2.""ItemCode"" = T0.""ItemCode"") Where " & _
                '                                           " ""ItemCode"" = '" & ItemCode & "' And T1.""CardType"" = 'S' And  T0.""UomEntry"" = '" & UOMEntry & "' " & _
                '                                           " Order By T0.""Price""  Asc "

                Dim GetLeastVendorDtlsAndPrice As String = " Select Top 1 T0.""ItemCode"", T1.""CardCode"", T1.""CardName"", T0.""Price""  " &
               " From SPP2 T0 Inner Join OCRD T1 On T1.""CardCode"" = T0.""CardCode""  " &
               " Where ""ItemCode"" = '" & ItemCode & "' And T1.""CardType"" = 'S' Order By T0.""Price""  Asc "

                Dim oRsGetLeastVendorDtlsAndPrice As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetLeastVendorDtlsAndPrice.DoQuery(GetLeastVendorDtlsAndPrice)
                If oRsGetLeastVendorDtlsAndPrice.RecordCount > 0 Then
                    LeastVendorPrice = oRsGetLeastVendorDtlsAndPrice.Fields.Item("Price").Value
                    LeastVendorCode = oRsGetLeastVendorDtlsAndPrice.Fields.Item("CardCode").Value
                    LeastVendorName = oRsGetLeastVendorDtlsAndPrice.Fields.Item("CardName").Value
                End If
            ElseIf IsUomDefault = "Yes" Then
                UOMEntry = Me.GetUomEntry(Me.GetDefaultUomCode(ItemCode))
                'Dim GetPurchasePriceDfltUOM As String = "Select ""Price"" From ITM1 Where ""ItemCode""='" & ItemCode & "' AND ""PriceList""= (Select ""ListNum"" From OCRD Where ""CardCode"" = '" & Vendor & "')"
                Dim GetPurchasePriceDfltUOM As String = " Select IFNULL(T3.""Price"",0) ""Price"" from SPP2 T3 " &
                " Where ""ItemCode"" = '" & ItemCode & "' and ""CardCode"" = '" & Vendor & "' and T3.""UomEntry"" = '" & UOMEntry & "' "

                Dim oRsGetPurchasePriceDfltUOM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetPurchasePriceDfltUOM.DoQuery(GetPurchasePriceDfltUOM)
                SalesPrice = CDbl(oRsGetPurchasePriceDfltUOM.Fields.Item(0).Value)
                If PCPrPack <> "" Then
                    Packing = PCPrPack & "*" & "1*" & BaseUom
                Else
                    Packing = "1*" & BaseUom
                End If
                'Dim GetRetailPriceBrnchWise As String = " Select T1.""Price"" From OITM T0 Inner Join ITM1 T1 On T1.""ItemCode"" = T0.""ItemCode"" " & _
                '                                        " Inner Join OBPL T2 On T1.""PriceList""=T2.""U_PricelistNo"" Where T0.""ItemCode"" = '" & ItemCode & "' " & _
                '                                        " And T2.""BPLId""= '" & objForm.Items.Item("2001").Specific.Value.Trim & "' "

                Dim GetRetailPriceBrnchWise As String = " Select T0.""Price"" From SPP2 T0 Inner Join ITM1 T1 On T1.""ItemCode"" = T0.""ItemCode"" " &
                " Inner Join OBPL T2 On T1.""PriceList"" = T2.""U_PricelistNo"" " &
                " Where T0.""ItemCode"" = '" & ItemCode & "'  And T2.""BPLId""= '" & objForm.Items.Item("2001").Specific.Value.Trim & "' And T0.""UomEntry""='" & UOMEntry & "' "

                Dim oRsGetRetailPriceBrnchWise As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetRetailPriceBrnchWise.DoQuery(GetRetailPriceBrnchWise)
                If oRsGetRetailPriceBrnchWise.RecordCount > 0 Then
                    RetailsPrice = oRsGetRetailPriceBrnchWise.Fields.Item("Price").Value
                End If
                'Dim GetLeastVendorDtlsAndPrice As String = " Select Top 1 T0.""ItemCode"" , T1.""CardCode"", T1.""CardName"" , T0.""Price"" From ITM1 T0 " & _
                '                                           " Inner Join OCRD T1 On T1.""ListNum"" = T0.""PriceList"" And T1.""CardCode"" In (Select ""VendorCode"" From ITM2 Where  ITM2.""ItemCode"" = T0.""ItemCode"") Where " & _
                '                                           " ""ItemCode"" = '" & ItemCode & "' And T1.""CardType"" = 'S' " & _
                '                                           " Order By T0.""Price""  Asc "

                Dim GetLeastVendorDtlsAndPrice As String = " Select Top 1 T0.""ItemCode"", T1.""CardCode"", T1.""CardName"", T0.""Price""  " &
              " From SPP2 T0 Inner Join OCRD T1 On T1.""CardCode"" = T0.""CardCode""  " &
              " Where ""ItemCode"" = '" & ItemCode & "' And T1.""CardType"" = 'S' Order By T0.""Price""  Asc "

                Dim oRsGetLeastVendorDtlsAndPrice As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetLeastVendorDtlsAndPrice.DoQuery(GetLeastVendorDtlsAndPrice)
                If oRsGetLeastVendorDtlsAndPrice.RecordCount > 0 Then
                    LeastVendorPrice = oRsGetLeastVendorDtlsAndPrice.Fields.Item("Price").Value
                    LeastVendorCode = oRsGetLeastVendorDtlsAndPrice.Fields.Item("CardCode").Value
                    LeastVendorName = oRsGetLeastVendorDtlsAndPrice.Fields.Item("CardName").Value
                End If
            End If
            If objMatrix.Columns.Item("4").Cells.Item(RowId).Specific.Value = "" Then
                Dim GetBarCode As String = " SELECT Top 1 ""BcdCode"" FROM OBCD WHERE ""UomEntry"" = '" & UOMEntry & "' " &
                                           " And ""ItemCode"" = '" & objMatrix.Columns.Item("1").Cells.Item(RowId).Specific.Value & "' ORDER BY ""BcdEntry"" ASC "
                Dim oRsGetBarCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetBarCode.DoQuery(GetBarCode)
                If oRsGetBarCode.RecordCount > 0 Then
                    AdditionalBarCode = oRsGetBarCode.Fields.Item("BcdCode").Value
                Else
                    Dim GetBarCode1 As String = " SELECT Top 1 ""BcdCode"" FROM OBCD WHERE " &
                                                " ""ItemCode"" = '" & objMatrix.Columns.Item("1").Cells.Item(RowId).Specific.Value & "' ORDER BY ""BcdEntry"" ASC "
                    Dim oRsGetBarCode1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetBarCode1.DoQuery(GetBarCode1)
                    If oRsGetBarCode1.RecordCount > 0 Then
                        AdditionalBarCode = oRsGetBarCode1.Fields.Item("BcdCode").Value
                    End If
                End If
            Else
                AdditionalBarCode = objMatrix.Columns.Item("4").Cells.Item(RowId).Specific.Value
            End If
            '            Try

            '                Dim GetBase As String = " SELECT T0.""U_SUTAXAMT"" AS ""Price"" FROM ""OBCD"" T0 LEFT JOIN OITM T1 ON T1.""ItemCode""=T0.""ItemCode"" 
            ' WHERE T0.""ItemCode"" =  '" & ItemCode & "' and T0.""UomEntry"" = '" & UOMEntry & "' and T1.""U_SugarTax""='Y'"

            '                Dim oRsGetbase As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '                oRsGetbase.DoQuery(GetBase)



            '                objMatrix.Columns.Item("U_BASE").Cells.Item(RowId).Specific.Value = oRsGetbase.Fields.Item("Price").Value

            '                Dim Discount As Double = objMatrix.Columns.Item("U_DSAT").Cells.Item(RowId).Specific.Value
            '                Dim baseVal As Double = 0
            '                Dim discountVal As Double = 0
            '                Dim badVal As Double = 0
            '                Dim extVal As Double = 0

            '                Double.TryParse(objMatrix.Columns.Item("U_BASE").Cells.Item(RowId).Specific.Value, baseVal)
            '                Double.TryParse(Discount, discountVal)

            '                ' U_BAD = U_BASE - Discount
            '                badVal = baseVal - discountVal
            '                objMatrix.Columns.Item("U_BAD").Cells.Item(RowId).Specific.Value = badVal.ToString("0.00")

            '                ' U_EXT = 50% of U_BAD
            '                extVal = (baseVal * 50) / 100
            '                objMatrix.Columns.Item("U_EXT").Cells.Item(RowId).Specific.Value = extVal.ToString("0.00")

            '                ' Column 14 (Unit Price) = U_BAD + U_EXT
            '                objMatrix.Columns.Item("14").Cells.Item(RowId).Specific.Value = (badVal + extVal).ToString("0.00")

            '                Dim GetSPP1 As String = "SELECT T0.""ItemCode"" AS ""Item No."",
            '	 T4.""ItemName"" AS ""Description"",
            '	COALESCE(T2.""Price"",
            '	 T1.""Price"",
            '	 T0.""Price"") AS ""Effective Price"",
            '	 COALESCE(T2.""Discount"",
            '	 T1.""Discount"",
            '	 T0.""Discount"") AS ""Applied Discount %"" 
            'FROM ""OSPP"" T0 
            'LEFT JOIN ""SPP1"" T1 ON T0.""CardCode"" = T1.""CardCode"" 
            'AND T0.""ItemCode"" = T1.""ItemCode"" 
            'LEFT JOIN ""SPP2"" T2 ON T1.""CardCode"" = T2.""CardCode"" 
            'AND T1.""ItemCode"" = T2.""ItemCode"" 
            'INNER JOIN ""OCRD"" T3 ON T0.""CardCode"" = T3.""CardCode"" 
            'INNER JOIN ""OITM"" T4 ON T0.""ItemCode"" = T4.""ItemCode"" 
            'WHERE T0.""CardCode"" = '" & Vendor & "' and T0.""ItemCode""='" & ItemCode & "' and T2.""UomEntry"" = '" & UOMEntry & "'
            'AND TO_DATE('" & [Date] & "', 'YYYYMMDD') BETWEEN T1.""FromDate"" AND T1.""ToDate""
            'ORDER BY T0.""CardCode"",T0.""ItemCode"",T1.""FromDate"""

            '                Dim oRsGetSPP1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '                oRsGetSPP1.DoQuery(GetSPP1)

            '                SppPrice = oRsGetSPP1.Fields.Item("Effective Price").Value


            '                '            Dim GetSPP As String = "SELECT T1.""Price"" FROM OITM T0 Left JOIN ITM1 T1 ON T0.""ItemCode"" = T1.""ItemCode""
            '                'WHERE T0.""ItemCode"" = '" & ItemCode & "' AND T1.""PriceList"" = '11'"

            '                Dim GetSPP As String = " SELECT (P.""Price"" * G.""BaseQty"") ""Price""
            '            FROM OITM O
            '            INNER JOIN UGP1 G ON O.""UgpEntry"" = G.""UgpEntry""
            '            INNER JOIN OUOM U ON G.""UomEntry"" = U.""UomEntry""
            '            INNER JOIN ITM1 P ON O.""ItemCode"" = P.""ItemCode""
            '            INNER JOIN OBCD B ON O.""ItemCode"" = B.""ItemCode""
            '            WHERE 
            '            B.""BcdCode"" = '" & BarCode & "'
            '              and O.""ItemCode"" = '" & ItemCode & "'
            '              AND P.""PriceList"" = '11'
            '              AND U.""UomEntry"" = '" & UOMEntry & "'"

            '                Dim oRsGetSPP As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '                oRsGetSPP.DoQuery(GetSPP)

            '                Dim GetSugar As String = "SELECT T0.""U_SugarTax"" FROM OITM T0 WHERE T0.""ItemCode"" = '" & ItemCode & "'"

            '                Dim oRsGetSugar As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '                oRsGetSugar.DoQuery(GetSugar)

            '                If oRsGetSugar.Fields.Item("U_SugarTax").Value = "Y" Then
            '                    objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value = (badVal + extVal).ToString("0.00")
            '                Else
            '                    If SppPrice <> 0 Then
            '                        objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value = SppPrice
            '                    Else
            '                        objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value = oRsGetSPP.Fields.Item("Price").Value
            '                    End If

            '                End If


            '                ''objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value = SalesPrice
            '            Catch ex As Exception
            '            End Try
            Try
                objMatrix.Columns.Item("U_Packing").Cells.Item(RowId).Specific.Value = Packing
            Catch ex As Exception
            End Try
            Try
                objMatrix.Columns.Item("U_RetailPrice").Cells.Item(RowId).Specific.Value = RetailsPrice
            Catch ex As Exception
            End Try
            Try
                objMatrix.Columns.Item("U_Addbarcode").Cells.Item(RowId).Specific.Value = AdditionalBarCode
            Catch ex As Exception
            End Try
            Try
                objMatrix.Columns.Item("U_LVNDPRC").Cells.Item(RowId).Specific.Value = LeastVendorPrice
            Catch ex As Exception
            End Try
            Try
                objMatrix.Columns.Item("U_LPVCD").Cells.Item(RowId).Specific.Value = LeastVendorCode
            Catch ex As Exception
            End Try
            Try
                objMatrix.Columns.Item("U_LPVNM").Cells.Item(RowId).Specific.Value = LeastVendorName
            Catch ex As Exception
            End Try
            Dim NetCost As String = objMatrix.Columns.Item("14").Cells.Item(RowId).Specific.Value
            If NetCost <> "" Then
                NetCost = NetCost.Replace(objForm.Items.Item("63").Specific.Value, "")
                NetCost = Convert.ToDouble(NetCost.Trim)
                Dim LeastPrice As Double = objMatrix.Columns.Item("U_LVNDPRC").Cells.Item(RowId).Specific.Value
                If NetCost > 0 And LeastPrice > 0 Then
                    If NetCost < LeastPrice Then
                        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LPValidate")
                        oColumn.Cells.Item(RowId).Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If
                End If
                '-------------------Below cost---------------------------
                Dim RetailPrice As Double = objMatrix.Columns.Item("U_RetailPrice").Cells.Item(RowId).Specific.Value
                If RetailPrice > 0 And NetCost > 0 Then
                    If NetCost > RetailPrice Then
                        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_BelowCst")
                        oColumn.Cells.Item(RowId).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If
                End If
            End If
            '-------------------Below cost---------------------------
            objMain.objApplication.StatusBar.SetText("Operation completed successfully....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Function getSupplierPrice_SpecialPrices(ByVal FormUID As String, ByVal IsUomDefault As String, _
                                 ByVal ItemCode As String, ByVal Vendor As String, ByVal UOMEntry As Integer, ByVal UomCode As String) As Double
        Dim SuppPrice As Double = 0
        Dim DflUomPrice As Double = 0
        Dim BaseQty As Double = 0
        Dim UnitCost As Double = 0
        Dim OtherUOMPrice As Double = 0
        Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            If IsUomDefault = "No" Then
                'Dim GetPurchasePriceDfltUOM As String = "Select ""Price"" From ITM1 Where ""ItemCode""='" & ItemCode & "' " & _
                '" AND ""PriceList""= (Select ""ListNum"" From OCRD Where ""CardCode"" = '" & Vendor & "')"

                Dim GetPurchasePriceDfltUOM As String = " Select IFNULL(T3.""Price"",0) ""Price"" from SPP2 T3  " & _
                " Where T3.""ItemCode"" = '" & ItemCode & "' and T3.""CardCode"" = '" & Vendor & "'  "

                Dim oRsGetPurchasePriceDfltUOM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetPurchasePriceDfltUOM.DoQuery(GetPurchasePriceDfltUOM)
                If oRsGetPurchasePriceDfltUOM.RecordCount > 0 Then
                    DflUomPrice = CDbl(oRsGetPurchasePriceDfltUOM.Fields.Item(0).Value)
                    BaseQty = Me.GetBaseQty(ItemCode, UOMEntry.ToString)
                    SuppPrice = BaseQty * DflUomPrice
                Else
                    'Dim GetPurchasePrice As String = " Select T0.""Price"" , T0.""UomEntry"" From ITM9 T0 Where T0.""PriceList""= " & _
                    '                                 " (Select ""ListNum"" From OCRD Where ""CardCode"" = '" & Vendor & "') And " & _
                    '                                 " ""ItemCode""='" & ItemCode & "' "

                    Dim GetPurchasePrice As String = " Select IFNULL(T3.""Price"",0) ""Price"",T3.""UomEntry"" from SPP2 T3  " & _
                    " Where T3.""ItemCode"" = '" & ItemCode & "' and T3.""CardCode"" = '" & Vendor & "'  "

                    Dim oRsGetPurchasePrice As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetPurchasePrice.DoQuery(GetPurchasePrice)
                    If oRsGetPurchasePrice.RecordCount > 0 Then
                        OtherUOMPrice = CDbl(oRsGetPurchasePrice.Fields.Item(0).Value)
                        BaseQty = Me.GetBaseQty(ItemCode, oRsGetPurchasePrice.Fields.Item("UomEntry").Value.ToString())
                        UnitCost = OtherUOMPrice / BaseQty
                        SuppPrice = UnitCost * CDbl(Me.GetBaseQty(ItemCode, UOMEntry.ToString))
                    End If
                End If
            End If

            If IsUomDefault = "Yes" Then
                'Dim GetPurchasePrice As String = " Select T0.""Price"" , T0.""UomEntry"" From ITM9 T0 Where T0.""PriceList""= " & _
                '                                                    " (Select ""ListNum"" From OCRD Where ""CardCode"" = '" & Vendor & "') And " & _
                '                                                    " ""ItemCode""='" & ItemCode & "' "

                Dim GetPurchasePrice As String = " Select IFNULL(T3.""Price"",0) ""Price"",T3.""UomEntry"" from SPP2 T3  " & _
                   " Where T3.""ItemCode"" = '" & ItemCode & "' and T3.""CardCode"" = '" & Vendor & "'  "

                Dim oRsGetPurchasePrice As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetPurchasePrice.DoQuery(GetPurchasePrice)
                If oRsGetPurchasePrice.RecordCount > 0 Then
                    OtherUOMPrice = CDbl(oRsGetPurchasePrice.Fields.Item(0).Value)
                    BaseQty = Me.GetBaseQty(ItemCode, oRsGetPurchasePrice.Fields.Item("UomEntry").Value.ToString())
                    UnitCost = OtherUOMPrice / BaseQty
                    SuppPrice = UnitCost
                End If
            End If

            Return SuppPrice
        Catch ex As Exception
            Return SuppPrice
            objMain.objApplication.StatusBar.SetText(ex.Message)
        Finally
            oRS = Nothing
        End Try
    End Function

    Public Function GetBaseQty(ByVal ItemCode As String, ByVal UomEntry As String)
        Dim BaseQty As Double = 0
        Dim oRsFetchBaseQty As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            Dim FetchBaseQty As String = " Select ""BaseQty"" From UGP1 Where ""UomEntry"" =  '" & UomEntry & "' And ""UgpEntry""  =  " &
                                         " (Select ""UgpEntry"" From OITM Where ""ItemCode"" = '" & ItemCode & "')"
            oRsFetchBaseQty.DoQuery(FetchBaseQty)

            BaseQty = oRsFetchBaseQty.Fields.Item("BaseQty").Value

            Return BaseQty
        Catch ex As Exception
            Return BaseQty
            objMain.objApplication.StatusBar.SetText(ex.Message)
        Finally
            oRsFetchBaseQty = Nothing
        End Try
    End Function
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("143", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then


                        If CStr(oDBs_Head.GetValue("U_GUID", 0)).Trim <> "" Then

                            Dim UpdateDa As String = "Update ""@TNXEDS"" SET ""U_BENTRY""='" & oDBs_Head.GetValue("DocEntry", 0) & "', " &
                                                  """U_BNO""='" & oDBs_Head.GetValue("DocNum", 0) & "' Where ""U_GUID""='" & oDBs_Head.GetValue("U_GUID", 0) & "'  And ""U_DOCTY""='OPDN'"

                            Dim oRsUpdateDa As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            oRsUpdateDa.DoQuery(UpdateDa)

                        End If


                        objMain.objApplication.MessageBox("GRPO Document is Succesfully Added")

                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Dim Validation As String = "Select ""U_GRPO""as ""GRPO"" from ""@ME_ACFG"""
                Dim oLValidation As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oLValidation.DoQuery(Validation)
                If oLValidation.Fields.Item("GRPO").Value = "Y" Then
                    objForm.Items.Item("btn_calFOC").Visible = True
                Else
                    objForm.Items.Item("btn_calFOC").Visible = False
                End If

            End If
            'If pVal.BeforeAction Then

            '    If pVal.MenuUID = "1283" Then   'Remove

            '        Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.ActiveForm

            '        If oForm.TypeEx = "143" Then    'GRPO

            '            Dim Ret As Integer
            '            Ret = objMain.objApplication.MessageBox(
            '                "Are you sure you want to remove this GRPO Draft?",
            '                2, "Yes", "No")

            '            If Ret = 2 Then
            '                BubbleEvent = False
            '            End If

            '        End If

            '    End If

            'End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
End Class
