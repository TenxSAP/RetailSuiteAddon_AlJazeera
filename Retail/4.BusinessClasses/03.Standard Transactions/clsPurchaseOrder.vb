
Imports System.Text.RegularExpressions
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Header
Imports Newtonsoft.Json.Linq

Public Class clsPurchaseOrder
    ''test
#Region "Declaration"
    Public objForm, objGRPOForm As SAPbouiCOM.Form
    Dim objMatrix, objGRPOMatrix As SAPbouiCOM.Matrix
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim BaseUom As String
    Dim CalcualtionBtnFlag As Boolean = False
    Dim strQry As String = ""
    Dim oItem, oItem1, oItem2, oItem3 As SAPbouiCOM.Item
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim GlobalDocnum As String
    Private isUpdating As Boolean = False
    Dim OrsSTRQRY, rs, rsPO As SAPbobsCOM.Recordset
    Dim Flag As String = "No"
    Private IsRunning As Boolean = False
    Private LastEditedField As String = ""
    Private TabPressed As Boolean = False
    Private LastDiscRow As Integer = -1
    Private LastDiscValue As String = ""
    Dim SPPPPrice As Double

#End Region
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    Try
                        objForm = objMain.objApplication.Forms.Item(FormUID)
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("OPOR")
                        objMatrix = objForm.Items.Item("38").Specific
                        If pVal.ItemUID = "38" And pVal.ColUID = "U_BelowCst" And pVal.BeforeAction = True Then
                            objMain.objApplication.StatusBar.SetText("User not authorized to performe this operation...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                        End If
                    Catch ex As Exception

                    End Try

                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    Try
                        objForm = objMain.objApplication.Forms.Item(FormUID)
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("OPOR")
                        objMatrix = objForm.Items.Item("38").Specific
                        If pVal.ItemUID = "38" And pVal.ColUID = "U_BelowCst" And pVal.BeforeAction = True Then
                            objMain.objApplication.StatusBar.SetText("User not authorized to performe this operation...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                        End If


                        Try
                            objForm = objMain.objApplication.Forms.Item(FormUID)
                            objMatrix = objForm.Items.Item("38").Specific


                            If pVal.ItemUID = "38" AndAlso pVal.ColUID = "4" AndAlso pVal.BeforeAction = True And pVal.CharPressed = 9 Then
                                Try
                                    objForm.Freeze(True)

                                    Dim Barcode As String = CType(objMatrix.Columns.Item("4").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.Trim()

                                    'Trigger only when Barcode is empty
                                    If String.IsNullOrWhiteSpace(Barcode) Then
                                        System.Threading.Thread.Sleep(50)
                                        objMain.objApplication.ActivateMenuItem("7425")
                                    End If

                                Catch ex As Exception

                                Finally
                                    objForm.Freeze(False)
                                End Try
                            End If

                            If pVal.ItemUID = "38" AndAlso pVal.ColUID = "1" AndAlso pVal.BeforeAction = True And pVal.CharPressed = 9 Then
                                Try
                                    objForm.Freeze(True)

                                    Dim ITcode As String = CType(objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.Trim()

                                    'Trigger only when Barcode is empty
                                    If String.IsNullOrWhiteSpace(ITcode) Then
                                        System.Threading.Thread.Sleep(50)
                                        objMain.objApplication.ActivateMenuItem("7425")
                                    End If

                                Catch ex As Exception

                                Finally
                                    objForm.Freeze(False)
                                End Try
                            End If

                        Catch ex As Exception
                            ' Handle exception
                        End Try

                        'If pVal.ItemUID = "38" AndAlso pVal.BeforeAction = True Then

                        '    If pVal.CharPressed = 9 Then   'Tab key
                        '        TabPressed = True
                        '    End If

                        '    If pVal.ColUID = "4" AndAlso TabPressed Then

                        '        System.Threading.Thread.Sleep(100)
                        '        objMain.objApplication.ActivateMenuItem("7425")

                        '        TabPressed = False

                        '    End If

                        'End If
                        If pVal.ItemUID = "38" AndAlso pVal.BeforeAction = False AndAlso pVal.CharPressed = 40 Then
                            Try
                                If CStr(objMatrix.Columns.Item("4").Cells.Item(pVal.Row + 1).Specific.Value).Trim = "" Then
                                    objMatrix.Columns.Item("4").Cells.Item(pVal.Row + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                End If
                            Catch ex As Exception
                                ' Handle or ignore boundary exceptions if on the last row
                            End Try
                        End If
                    Catch ex As Exception

                    End Try

                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    If pVal.ItemUID = "10000329" And pVal.BeforeAction = False Then
                        Dim objComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("10000329").Specific
                        Dim Doc1 As String = objComboBox.Selected.Description
                        If Doc1 = "20" Then
                            Dim objBaseForm As SAPbouiCOM.Form
                            objBaseForm = objMain.objApplication.Forms.GetForm("143", objMain.objApplication.Forms.ActiveForm.TypeCount)
                            Dim oRsBranch As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim Query As String = " Select T1.""Quantity"",T1.""OpenQty"",(T1.""Quantity""- T1.""OpenQty"")as ""Delieverd Quantity"",T1.""LineNum"",T1.""ItemCode"" From OPOR T0  Inner Join " &
                                                                                "POR1 T1 On T1.""DocEntry"" = T0.""DocEntry"" Where T0.""DocNum"" = '" & objForm.Items.Item("8").Specific.Value & "'and T1.""LineStatus""='O' "
                            oRsBranch.DoQuery(Query)
                            objMatrix = objBaseForm.Items.Item("38").Specific
                            If oRsBranch.RecordCount > 0 Then

                                For k As Integer = 1 To oRsBranch.RecordCount
                                    Dim Deliver As String = oRsBranch.Fields.Item("Delieverd Quantity").Value
                                    If oRsBranch.Fields.Item("Delieverd Quantity").Value <> oRsBranch.Fields.Item("Quantity").Value Then
                                        objMatrix.Columns.Item("U_Delivr").Cells.Item(k).Specific.Value = Deliver
                                    End If
                                    oRsBranch.MoveNext()

                                    'For k As Integer = 1 To oRsBranch.RecordCount
                                    '    Dim Deliver As String = oRsBranch.Fields.Item("Delieverd Quantity").Value
                                    '    Dim line As Integer = oRsBranch.Fields.Item("LineNum").Value
                                    '    '   line = line + 1
                                    '    If oRsBranch.Fields.Item("Delieverd Quantity").Value <> oRsBranch.Fields.Item("Quantity").Value Then
                                    '        objMatrix.Columns.Item("U_Delivr").Cells.Item(line).Specific.Value = Deliver
                                    '    End If
                                    '    oRsBranch.MoveNext()
                                    '    line = 0
                                Next
                            End If
                        End If

                    End If
                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("OPOR")
                    objMatrix = objForm.Items.Item("38").Specific
                    If pVal.ItemUID = "38" And pVal.ColUID = "0" And pVal.BeforeAction = True Then
                        If objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value <> "" Then         'DocType = Service
                            BubbleEvent = False
                            objMain.objPurchaseHistoryDetails.CreateForm(objForm.UniqueID, objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value)
                            Exit Sub
                        End If
                    End If
                    'Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    '    objForm = objMain.objApplication.Forms.Item(FormUID)
                    '    If pVal.ItemUID = "38" And pVal.ColUID = "1" And pVal.BeforeAction = False Then
                    '        For i As Integer = 1 To pVal.Row
                    '            If objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value <> "" Then
                    '                Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value
                    '                Dim GetDocEntry As String = "Select ""FrgnName"" From OITM Where ""ItemCode"" = '" & ItemCode & "' "
                    '                Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                oRs.DoQuery(GetDocEntry)
                    '                objMatrix.Columns.Item("U_ForeNm").Cells.Item(i).Specific.Value() = oRs.Fields.Item("FrgnName").Value()
                    '            End If
                    '        Next
                    '    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("OPOR")
                    objMatrix = objForm.Items.Item("38").Specific

                    If pVal.ItemUID = "btn_MRG" And pVal.BeforeAction = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Try
                            objForm.Freeze(True)
                            For i As Integer = 1 To objMatrix.VisualRowCount
                                Dim ISGROSS As String = ""

                                If CStr(objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value).Trim <> "" Then

                                    Dim VatGrpRate As String = "SELECT (T1.""Rate""+100)/100 FROM OITM T0  INNER JOIN OVTG T1 ON T0.""VatGourpSa"" = T1.""Code"" Where T0.""ItemCode""='" & CStr(objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value).Trim & "'"

                                    Dim oRsVatGrpRate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    oRsVatGrpRate.DoQuery(VatGrpRate)

                                    Dim Amm As Double = 0

                                    If CStr(oRsVatGrpRate.Fields.Item(0).Value).Trim <> "" Then
                                        Amm = CDbl(CStr(oRsVatGrpRate.Fields.Item(0).Value).Trim)
                                    End If

                                    objComboBox = objForm.Items.Item("2001").Specific
                                    Dim Branch As Integer = objComboBox.Value
                                    Dim GROSS As String = "SELECT IFNULL(X.""IsGrossPrc"",'N') FROM ""OBPL"" M INNER  JOIN ""OPLN"" X ON M.""U_PricelistNo""=""ListNum"" WHERE ""BPLId""='" & Branch & "'"
                                    Dim RSGROSS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    RSGROSS.DoQuery(GROSS)

                                    If CStr(RSGROSS.Fields.Item(0).Value).Trim <> "" Then
                                        ISGROSS = (CStr(RSGROSS.Fields.Item(0).Value).Trim)
                                    End If
                                    'Dim NetCost As Double = 0

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
                                    Dim PriAfterDis As Double = 0
                                    Dim pattern As String = "\d+(\.\d{1,4})?"

                                    Dim match As Match = Regex.Match(CStr(objMatrix.Columns.Item("17").Cells.Item(i).Specific.Value).Trim, pattern)


                                    If CStr(objMatrix.Columns.Item("17").Cells.Item(i).Specific.Value).Trim <> "" Then
                                        PriAfterDis = CDbl(match.Value)
                                    End If

                                    Dim RetailPrice1 As Double = 0
                                    Dim SuppPrice As Double = 0


                                    'If CStr(objMatrix.Columns.Item("U_RetailPrice").Cells.Item(i).Specific.Value).Trim <> "" Then
                                    '    RetailPrice1 = CDbl(CStr(objMatrix.Columns.Item("U_RetailPrice").Cells.Item(i).Specific.Value).Trim)
                                    '    RetailPrice1 = RetailPrice1 / Amm
                                    'End If
                                    If CStr(objMatrix.Columns.Item("U_RetailPrice").Cells.Item(i).Specific.Value).Trim <> "" And ISGROSS = "Y" Then
                                        RetailPrice1 = CDbl(CStr(objMatrix.Columns.Item("U_RetailPrice").Cells.Item(i).Specific.Value).Trim)
                                        RetailPrice1 = RetailPrice1 / Amm
                                        RetailPrice1 = Math.Round(RetailPrice1, 4)
                                    ElseIf CStr(objMatrix.Columns.Item("U_RetailPrice").Cells.Item(i).Specific.Value).Trim <> "" And ISGROSS = "N" Then
                                        RetailPrice1 = CDbl(CStr(objMatrix.Columns.Item("U_RetailPrice").Cells.Item(i).Specific.Value).Trim)
                                    Else
                                        RetailPrice1 = CDbl(CStr(objMatrix.Columns.Item("U_RetailPrice").Cells.Item(i).Specific.Value).Trim)
                                    End If

                                    If PriAfterDis > RetailPrice1 Then
                                        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_BelowCst")
                                        oColumn.Cells.Item(i).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                    Else
                                        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_BelowCst")
                                        oColumn.Cells.Item(i).Specific.Select("No", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                    End If

                                    Try



                                        'Dim pattern As String = "\d+(\.\d{1,4})?"

                                        'Dim match As Match = Regex.Match(CStr(objMatrix.Columns.Item("17").Cells.Item(i).Specific.Value).Trim, pattern)


                                        If CStr(objMatrix.Columns.Item("17").Cells.Item(i).Specific.Value).Trim <> "" Then
                                            SuppPrice = CDbl(match.Value)
                                        End If

                                        objMatrix.Columns.Item("U_MARGN").Cells.Item(i).Specific.Value = ((RetailPrice1 - SuppPrice) / RetailPrice1) * 100

                                    Catch ex As Exception

                                    End Try

                                End If
                            Next

                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        Finally

                            objForm.Freeze(False)
                        End Try
                        objMain.objApplication.MessageBox("Item details loaded successfully")
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                    (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        'If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        '    strQry = "SELECT *  FROM OPOR T0 WHERE T0.""U_Approval"" ='Approved' AND  T0.""DocNum"" ='" & objForm.Items.Item("8").Specific.Value & "' and  T0.""DocDate"" ='" & objForm.Items.Item("10").Specific.Value & "'"
                        '    OrsSTRQRY = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '    If OrsSTRQRY.RecordCount > 0 Then

                        '    Else
                        '        If Me.Validation(objForm.UniqueID) = False Then BubbleEvent = False
                        '    End If
                        'Else


                        If Me.Validation(objForm.UniqueID) = False Then BubbleEvent = False

                        If Me.ValidationnVadlityDate(objForm.UniqueID) = False Then
                            BubbleEvent = False
                        End If

                    End If

                    If pVal.ItemUID = "btn_Supp" And pVal.BeforeAction = False Then
                        objMatrix = objForm.Items.Item("38").Specific
                        'If objForm.Items.Item("42").Specific.Value <> "" Then
                        objForm.Items.Item("42").Specific.Value = ""
                        objForm.Items.Item("24").Specific.Value = ""

                        For i As Integer = 1 To objMatrix.VisualRowCount - 1
                            Dim str As String = objMatrix.Columns.Item("1980002193").Cells.Item(i).Specific.Value
                            Dim strPrice As String = objForm.Items.Item("22").Specific.Value
                            Dim amount As Double = CDbl(strPrice.Substring(3))
                            If str <> "" Then
                                strQry = "SELECT T2.""U_POAF""as""PurchaseP"", T2.""U_POAT""as""PurchaseAmount"", T2.""U_RBTPRC""as ""Percentage"", T2.""U_RAmount""as""Amount"",T2.""U_Alert"" as ""Alert"" FROM ""@SBO_SUPREBATE""  T0  inner join OOAT T1 on  T0.""U_BASENO"" = T1.""Number"" " &
                                    "inner join  ""@SBO_SUPREBATE_C0""  T2  on  T0.""DocEntry"" = T2.""DocEntry"" WHERE T1.""Number"" = '" & str & "' And '" & amount & "' between  T2.""U_POAF"" and T2.""U_POAT"" "

                                OrsSTRQRY = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                OrsSTRQRY.DoQuery(strQry)
                                If OrsSTRQRY.RecordCount > 0 Then
                                    If amount <> 0 Then
                                        Dim purchase As Double = OrsSTRQRY.Fields.Item("Amount").Value
                                        If purchase > 0 Then
                                            objForm.Items.Item("42").Specific.Value = purchase
                                        End If

                                        Dim purchasePer As Double = OrsSTRQRY.Fields.Item("Percentage").Value
                                        If purchasePer > 0 Then
                                            objForm.Items.Item("24").Specific.Value = purchasePer
                                        End If
                                    End If
                                End If
                                'End If


                                rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim strQry1 As String = "SELECT T2.""U_POAF""as""PurchaseFrom"", T2.""U_POAT""as""PurchaseTo"", T2.""U_RBTPRC""as ""Percentage"", T2.""U_RAmount""as""Amount"",T2.""U_Alert"" as ""Alert"" FROM ""@SBO_SUPREBATE""  T0  inner join OOAT T1 on  T0.""U_BASENO"" = T1.""Number"" " &
                                        "inner join  ""@SBO_SUPREBATE_C0""  T2  on  T0.""DocEntry"" = T2.""DocEntry"" WHERE T1.""Number"" = '" & str & "' And '" & amount & "' between  T2.""U_Alert"" and T2.""U_POAF"" "
                                rs.DoQuery(strQry1)

                                If rs.RecordCount > 0 Then
                                    Dim purchaseFrom As Double = rs.Fields.Item("PurchaseFrom").Value
                                    Dim purchaseTo As Double = rs.Fields.Item("PurchaseTo").Value
                                    Dim Amount1 As Double = rs.Fields.Item("Amount").Value
                                    Dim percentage As Double = rs.Fields.Item("Percentage").Value
                                    If Amount1 <> 0 Then
                                        objMain.objApplication.MessageBox("If you Purchase '" & purchaseFrom & "' to  '" & purchaseTo & "'  you will  get '" & Amount1 & "'  discount")
                                    Else
                                        objMain.objApplication.MessageBox("If you Purchase '" & purchaseFrom & "' to  '" & purchaseTo & "'  you will  get '" & percentage & "'% discount")
                                    End If
                                End If
                            End If
                        Next

                        'End If
                    End If
                    If pVal.ItemUID = "btn_calFOC" And pVal.BeforeAction = False Then
                        objMain.objApplication.StatusBar.SetText("Please wait while calculating FOC details....", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        Dim itemList As New List(Of String)

                        For i As Integer = 1 To objMatrix.VisualRowCount - 1
                            Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.Trim
                            Dim SupplierPrice As String = objMatrix.Columns.Item("U_suppprice").Cells.Item(i).Specific.Value.Trim
                            If (ItemCode <> "" And SupplierPrice = "0.0") Then
                                objMain.objApplication.StatusBar.SetText("For Item " & ItemCode & " and Line No. " & i & "... Supplier Price should not be empty.....!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If
                            If ItemCode <> "" Then
                                itemList.Add("'" & ItemCode & "'")
                            End If
                        Next
                        'Me.CalculateFOC(objForm.UniqueID)
                        'Me.CalculateFOC_SUAGARTAX(objForm.UniqueID)
                        If itemList.Count > 0 Then
                            Dim sQuery As String = "SELECT TOP 1 IFNULL(""U_SugarTax"",'N') FROM ""OITM"" WHERE ""ItemCode"" IN (" & String.Join(",", itemList) & ") AND ""U_SugarTax"" = 'Y'"

                            Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRS.DoQuery(sQuery)

                            ' Step 3: Decision
                            If Not oRS.EoF Then

                                ' Found at least one item with SugarTax = Y
                                Me.CalculateFOC_SUAGARTAX(objForm.UniqueID)
                            Else
                                Me.CalculateFOC(objForm.UniqueID)
                                ''objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            End If
                            Dim focusRow As Integer = objMatrix.VisualRowCount

                            If focusRow > 0 Then
                                objMatrix.Columns.Item("1").Cells.Item(focusRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            End If

                            'Else
                            '    Me.CalculateFOC(objForm.UniqueID)
                        End If
                        'Dim LineValidation As String = "Select ""U_Least""as ""Least"" from ""@ME_ACFG"""
                        'Dim oLValidation As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'oLValidation.DoQuery(LineValidation)
                        'If oLValidation.Fields.Item("Least").Value = "N" Then
                        '    Me.LineValidations(objForm.UniqueID)
                        '    'Else

                        'End If
                        ' Me.LineValidations(objForm.UniqueID)
                        ''Ramya 
                        ''Load button functionality Start
                        ''08-08-2026 Commented for testing
                        'objMain.objApplication.StatusBar.SetText("Please wait while loading line details....", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        'Dim objComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("2001").Specific
                        'If objComboBox.Value = "" Then
                        '    objMain.objApplication.StatusBar.SetText("Please select branch at the doument header...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    Exit Sub
                        'End If

                        'For i As Integer = 1 To objMatrix.VisualRowCount - 1
                        '    If objMatrix.Columns.Item("1470002145").Cells.Item(i).Specific.Value.Trim = "" Then
                        '        objMain.objApplication.StatusBar.SetText("Please select  UomCode for the items " & objMatrix.Columns.Item("0").Cells.Item(i).Specific.Value & " ...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '        Exit Sub
                        '    End If
                        'Next
                        'Try
                        '    objForm.Freeze(True)
                        '    For i As Integer = 1 To objMatrix.VisualRowCount - 1
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

                        '        'Commented by Ramya on 20-05-2026 
                        '        'Me.LoadOnItemOrUOMSelection(objForm.UniqueID, i, IsUomDefault, ItemCode, UoM, objForm.Items.Item("4").Specific.Value)
                        '        ''For Special Prices
                        '        'Me.LoadOnItemOrUOMSelection_SpecialPrices(objForm.UniqueID, i, IsUomDefault, ItemCode, UoM, objForm.Items.Item("4").Specific.Value, objForm.Items.Item("10").Specific.Value, BarCode)

                        '        'Hol`ding Days here...!
                        '        Me.UpdateHoldingDays(objForm)

                        '    Next


                        'Catch ex As Exception
                        '    objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        'Finally

                        '    objForm.Freeze(False)
                        'End Try

                        objMain.objApplication.StatusBar.SetText("FOC calculation has been completed successfully.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objMain.objApplication.MessageBox("FOC calculation has been completed successfully.")
                        ''08-08-2026 Commented for testing
                        ''Button loading fucntionality end 
                        CalcualtionBtnFlag = True
                    End If
                    'If pVal.ItemUID = "btn_LdItd" And pVal.BeforeAction = False Then
                    '    objMain.objApplication.StatusBar.SetText("Please wait while loading line details....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    '    Dim objComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("2001").Specific
                    '    If objComboBox.Value = "" Then
                    '        objMain.objApplication.StatusBar.SetText("Please select branch at the doument header...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        Exit Sub
                    '    End If

                    '    For i As Integer = 1 To objMatrix.VisualRowCount - 1
                    '        If objMatrix.Columns.Item("1470002145").Cells.Item(i).Specific.Value.Trim = "" Then
                    '            objMain.objApplication.StatusBar.SetText("Please select  UomCode for the items " & objMatrix.Columns.Item("0").Cells.Item(i).Specific.Value & " ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '            Exit Sub
                    '        End If
                    '    Next
                    '    Try
                    '        objForm.Freeze(True)
                    '        For i As Integer = 1 To objMatrix.VisualRowCount - 1
                    '            Dim IsUomDefault As String = "Yes"
                    '            Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.ToString()
                    '            Dim UoM As String = objMatrix.Columns.Item("1470002145").Cells.Item(i).Specific.Value.ToString()
                    '            Dim BarCode As String = objMatrix.Columns.Item("4").Cells.Item(i).Specific.Value.ToString()
                    '            If UoM <> "" Then
                    '                If Me.CheckDefaultUOM(UoM, ItemCode) = True Then
                    '                    IsUomDefault = "Yes"
                    '                Else
                    '                    IsUomDefault = "No"
                    '                End If
                    '            End If

                    '            ''Commented by Ramya on 20-05-2026 
                    '            'Me.LoadOnItemOrUOMSelection(objForm.UniqueID, i, IsUomDefault, ItemCode, UoM, objForm.Items.Item("4").Specific.Value)
                    '            ''For Special Prices
                    '            'Me.LoadOnItemOrUOMSelection_SpecialPrices(objForm.UniqueID, i, IsUomDefault, ItemCode, UoM, objForm.Items.Item("4").Specific.Value, objForm.Items.Item("10").Specific.Value, BarCode)

                    '            ''Hol`ding Days here...!
                    '            'Me.UpdateHoldingDays(objForm)

                    '        Next


                    '    Catch ex As Exception
                    '        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    '    Finally

                    '        objForm.Freeze(False)
                    '    End Try

                    '    'Dim totalExcise As Double = 0.0
                    '    'For i As Integer = 1 To objMatrix.VisualRowCount - 1
                    '    '    Dim exciseVal As Double = 0.0
                    '    '    Dim exciseField = objMatrix.Columns.Item("U_EXT").Cells.Item(i).Specific

                    '    '    If Not String.IsNullOrEmpty(exciseField.Value) Then
                    '    '        Double.TryParse(exciseField.Value, exciseVal)
                    '    '    End If

                    '    '    totalExcise += exciseVal
                    '    'Next

                    '    'Dim headerExcise As SAPbouiCOM.EditText = objForm.Items.Item("U_EXT").Specific
                    '    'headerExcise.Value = totalExcise.ToString("F2")

                    '    objMain.objApplication.MessageBox("Item details loaded successfully")
                    'End If
                    If pVal.ItemUID = "fldSale" And pVal.BeforeAction = False Then
                        objForm.PaneLevel = "24"
                        'Me.Items(objForm.UniqueID)
                        Me.Averagesale()
                    End If

                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("OPOR")
                    objMatrix = objForm.Items.Item("38").Specific
                    If pVal.ItemUID = "10000329" And pVal.BeforeAction = False Then
                        Me.GRPOFormLoad(objForm.UniqueID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("OPOR")
                    objMatrix = objForm.Items.Item("38").Specific
                    If pVal.BeforeAction = False Then
                        CalcualtionBtnFlag = False
                        Me.AddItems(objForm.UniqueID)
                        Me.DisableCoulmns(objForm.UniqueID)

                        'objMatrix.Columns.Item("U_ForeNm").Enable
                        'objForm.Items.Item("btn_Supp").Enabled = False
                        GlobalDocnum = objForm.Items.Item("8").Specific.Value

                        Dim LineValidation As String = "Select ""U_Least""as ""Least"" from ""@ME_ACFG"""
                        Dim oLValidation As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oLValidation.DoQuery(LineValidation)
                        If oLValidation.Fields.Item("Least").Value = "N" Then
                            'objMatrix.Columns.Item("U_LPValidate").Width = 0
                            'objMatrix.Columns.Item("U_LVNDPRC").Width = 0
                            'objMatrix.Columns.Item("U_LPVCD").Width = 0
                            'objMatrix.Columns.Item("U_LPVNM").Width = 0
                            objMatrix.Columns.Item("U_LPValidate").Visible = False
                            objMatrix.Columns.Item("U_LVNDPRC").Visible = False
                            objMatrix.Columns.Item("U_LPVCD").Visible = False
                            objMatrix.Columns.Item("U_LPVNM").Visible = False
                        Else
                            objMatrix.Columns.Item("U_LPValidate").Visible = True
                            objMatrix.Columns.Item("U_LVNDPRC").Visible = True
                            objMatrix.Columns.Item("U_LPVCD").Visible = True
                            objMatrix.Columns.Item("U_LPVNM").Visible = True
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    Try
                        If pVal.ItemUID = "U_ADSUPPR" Then
                            LastEditedField = "P"
                        End If

                        If pVal.ItemUID = "U_ADSUPAMT" Then
                            LastEditedField = "A"
                        End If

                        If pVal.ItemUID = "U_DSAT" Then
                            LastEditedField = "D"
                        End If

                        If pVal.ItemUID = "U_ExTaDis" Then
                            LastEditedField = "E"
                        End If



                        'Try
                        '    objForm = objMain.objApplication.Forms.Item(FormUID)
                        '    objMatrix = objForm.Items.Item("38").Specific


                        '    If pVal.ItemUID = "38" AndAlso pVal.ColUID = "4" AndAlso pVal.BeforeAction = False Then
                        '        Try
                        '            objForm.Freeze(True)


                        '            System.Threading.Thread.Sleep(50)


                        '            objMain.objApplication.ActivateMenuItem("7425")

                        '        Catch ex As Exception

                        '        Finally
                        '            objForm.Freeze(False)
                        '        End Try
                        '    End If

                        'Catch ex As Exception
                        '    ' Handle exception
                        'End Try



                    Catch ex As Exception

                    End Try

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    Try
                        'If pVal.ItemUID = "38" And pVal.ColUID = "1" And pVal.BeforeAction = False Then
                        '    ' For i As Integer = 1 To objMatrix.VisualRowCount - 1
                        '    If objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value <> "" Then
                        '        Try
                        '            objForm.Freeze(True)

                        '            objMatrix = objForm.Items.Item("38").Specific
                        '            Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value
                        '            Dim GetDocEntry As String = "Select ""FrgnName"" From OITM Where ""ItemCode"" = '" & ItemCode & "' "
                        '            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '            oRs.DoQuery(GetDocEntry)
                        '            objMatrix.Columns.Item("U_ForeNm").Cells.Item(pVal.Row).Specific.Value() = oRs.Fields.Item("FrgnName").Value()


                        '        Catch ex As Exception

                        '        Finally
                        '            objForm.Freeze(False)

                        '        End Try

                        '    End If
                        '    'Next
                        'End If

                        If pVal.ItemUID = "38" And pVal.ColUID = "15" And pVal.BeforeAction = False Then
                            If Flag = "No" Then
                                Try
                                    objForm.Freeze(True)

                                    Flag = "Yes"
                                    Dim strPrice As String = objMatrix.Columns.Item("14").Cells.Item(pVal.Row).Specific.Value
                                    Dim SubString As String = strPrice.Substring(3)
                                    Dim Price As Double = CDbl(SubString)
                                    Dim Quantity As Double = objMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value
                                    Dim DiscountPer As Double = objMatrix.Columns.Item("15").Cells.Item(pVal.Row).Specific.Value

                                    Dim LineTotal As Double = Price * Quantity
                                    'objMatrix.Columns.Item("15").Cells.Item(pVal.Row).Specific.Value = (DiscountAmt / LineTotal) * 100
                                    objMatrix.Columns.Item("U_DiscountAmt").Cells.Item(pVal.Row).Specific.Value = LineTotal * DiscountPer * 0.01
                                Catch ex As Exception

                                Finally
                                    objForm.Freeze(False)
                                End Try

                            Else
                                Flag = "No"
                            End If
                        End If
                        If pVal.ItemUID = "38" And pVal.ColUID = "U_DiscountAmt" And pVal.BeforeAction = False Then
                            If Flag = "No" Then
                                Try
                                    objForm.Freeze(True)
                                    Flag = "Yes"
                                    Dim strPrice As String = objMatrix.Columns.Item("14").Cells.Item(pVal.Row).Specific.Value
                                    Dim Quantity As Double = objMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value
                                    Dim DiscountAmt As Double = objMatrix.Columns.Item("U_DiscountAmt").Cells.Item(pVal.Row).Specific.Value
                                    Dim SubString As String = strPrice.Substring(3)
                                    Dim Price As Double = CDbl(SubString)
                                    Dim LineTotal As Double = Price * Quantity
                                    objMatrix.Columns.Item("15").Cells.Item(pVal.Row).Specific.Value = (DiscountAmt / LineTotal) * 100
                                    Dim CurDiscountAmt As Double = objMatrix.Columns.Item("U_DiscountAmt").Cells.Item(pVal.Row).Specific.Value
                                    objMatrix.Columns.Item("21").Cells.Item(pVal.Row).Specific.Value = (Price * Quantity) - CurDiscountAmt
                                Catch ex As Exception

                                Finally
                                    objForm.Freeze(False)
                                End Try

                            Else
                                Flag = "No"
                            End If
                        End If
                        If pVal.ItemUID = "38" And pVal.ColUID = "1" And pVal.BeforeAction = False Then
                            'For i As Integer = 1 To objMatrix.VisualRowCount - 1

                            If objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value <> "" Then
                                objForm.Freeze(True)
                                objMatrix = objForm.Items.Item("38").Specific
                                Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value
                                Dim GetDocEntry As String = "Select ""U_EXCVAT"" From OITM Where ""ItemCode"" = '" & ItemCode & "' "
                                Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRs.DoQuery(GetDocEntry)
                                objMatrix.Columns.Item("U_EXCTAX").Cells.Item(pVal.Row).Specific.Value() = oRs.Fields.Item("U_EXCVAT").Value()
                                objForm.Freeze(False)
                            End If

                            'Next
                        End If

                        If pVal.ItemUID = "38" And pVal.ColUID = "U_DSAT" And pVal.BeforeAction = False Then
                            objForm.Freeze(True)
                            objMatrix.Columns.Item("U_BAD").Cells.Item(pVal.Row).Specific.Value = (objMatrix.Columns.Item("U_BASE").Cells.Item(pVal.Row).Specific.Value - objMatrix.Columns.Item("U_DSAT").Cells.Item(pVal.Row).Specific.Value)
                            'objMatrix.Columns.Item("U_EXT").Cells.Item(pVal.Row).Specific.Value = (objMatrix.Columns.Item("U_BASE").Cells.Item(pVal.Row).Specific.Value * 50) / 100

                            Dim bad As Double = 0
                            Dim ext As Double = 0

                            Double.TryParse(objMatrix.Columns.Item("U_BAD").Cells.Item(pVal.Row).Specific.Value, bad)
                            Double.TryParse(objMatrix.Columns.Item("U_EXT").Cells.Item(pVal.Row).Specific.Value, ext)

                            objMatrix.Columns.Item("14").Cells.Item(pVal.Row).Specific.Value = (bad + ext).ToString()
                            objMatrix.Columns.Item("U_suppprice").Cells.Item(pVal.Row).Specific.Value = (bad + ext).ToString()
                            objMatrix.Columns.Item("U_SuppPriBeDis").Cells.Item(pVal.Row).Specific.Value = (bad + ext).ToString()
                            objForm.Freeze(False)
                        End If
                        'If pVal.ItemUID = "38" And pVal.ColUID = "U_DSAT" And pVal.BeforeAction = False AndAlso LastEditedField = "D" Then

                        '    objMatrix = objForm.Items.Item("38").Specific

                        '    Dim BaseAmt As Double = 0
                        '    Dim DiscAmt As Double = 0
                        '    Dim DiscPer As Double = 0

                        '    Double.TryParse(objMatrix.Columns.Item("U_BASE").Cells.Item(pVal.Row).Specific.Value.ToString(), BaseAmt)
                        '    Double.TryParse(objMatrix.Columns.Item("U_DSAT").Cells.Item(pVal.Row).Specific.Value.ToString(), DiscAmt)

                        '    If BaseAmt <> 0 Then
                        '        DiscPer = Math.Round((DiscAmt / BaseAmt) * 100, 3, MidpointRounding.AwayFromZero)
                        '    Else
                        '        DiscPer = 0
                        '    End If

                        '    objMatrix.Columns.Item("U_ExTaDis").Cells.Item(pVal.Row).Specific.Value = DiscPer.ToString("0.000")
                        '    LastEditedField = ""
                        'End If
                        'If pVal.ItemUID = "38" And pVal.ColUID = "U_ExTaDis" And pVal.BeforeAction = False AndAlso LastEditedField = "E" Then

                        '    objMatrix = objForm.Items.Item("38").Specific

                        '    Dim BaseAmt As Double = 0
                        '    Dim DiscPer As Double = 0
                        '    Dim DiscAmt As Double = 0

                        '    Double.TryParse(objMatrix.Columns.Item("U_BASE").Cells.Item(pVal.Row).Specific.Value.ToString(), BaseAmt)
                        '    Double.TryParse(objMatrix.Columns.Item("U_ExTaDis").Cells.Item(pVal.Row).Specific.Value.ToString(), DiscPer)

                        '    DiscAmt = Math.Round((BaseAmt * DiscPer) / 100, 3, MidpointRounding.AwayFromZero)

                        '    objMatrix.Columns.Item("U_DSAT").Cells.Item(pVal.Row).Specific.Value = DiscAmt.ToString("0.000")
                        '    LastEditedField = ""
                        'End If
                        If pVal.ItemUID = "38" And pVal.ColUID = "U_ExTaDis" And pVal.BeforeAction = False Then
                            objForm.Freeze(True)
                            objMatrix.Columns.Item("U_BAD").Cells.Item(pVal.Row).Specific.Value = (objMatrix.Columns.Item("U_BASE").Cells.Item(pVal.Row).Specific.Value - objMatrix.Columns.Item("U_DSAT").Cells.Item(pVal.Row).Specific.Value)
                            'objMatrix.Columns.Item("U_EXT").Cells.Item(pVal.Row).Specific.Value = (objMatrix.Columns.Item("U_BASE").Cells.Item(pVal.Row).Specific.Value * 50) / 100

                            Dim bad As Double = 0
                            Dim ext As Double = 0

                            Double.TryParse(objMatrix.Columns.Item("U_BAD").Cells.Item(pVal.Row).Specific.Value, bad)
                            Double.TryParse(objMatrix.Columns.Item("U_EXT").Cells.Item(pVal.Row).Specific.Value, ext)

                            objMatrix.Columns.Item("14").Cells.Item(pVal.Row).Specific.Value = (bad + ext).ToString()
                            objMatrix.Columns.Item("U_suppprice").Cells.Item(pVal.Row).Specific.Value = (bad + ext).ToString()
                            objMatrix.Columns.Item("U_SuppPriBeDis").Cells.Item(pVal.Row).Specific.Value = (bad + ext).ToString()
                            objForm.Freeze(False)
                        End If
                        If pVal.ItemUID = "38" And pVal.ColUID = "U_EXT" And pVal.BeforeAction = False Then
                            objForm.Freeze(True)
                            objMatrix.Columns.Item("U_BAD").Cells.Item(pVal.Row).Specific.Value = (objMatrix.Columns.Item("U_BASE").Cells.Item(pVal.Row).Specific.Value - objMatrix.Columns.Item("U_DSAT").Cells.Item(pVal.Row).Specific.Value)
                            'objMatrix.Columns.Item("U_EXT").Cells.Item(pVal.Row).Specific.Value = (objMatrix.Columns.Item("U_BASE").Cells.Item(pVal.Row).Specific.Value * 50) / 100

                            Dim bad As Double = 0
                            Dim ext As Double = 0

                            Double.TryParse(objMatrix.Columns.Item("U_BAD").Cells.Item(pVal.Row).Specific.Value, bad)
                            Double.TryParse(objMatrix.Columns.Item("U_EXT").Cells.Item(pVal.Row).Specific.Value, ext)

                            objMatrix.Columns.Item("14").Cells.Item(pVal.Row).Specific.Value = (bad + ext).ToString()
                            objMatrix.Columns.Item("U_suppprice").Cells.Item(pVal.Row).Specific.Value = (bad + ext).ToString()
                            objMatrix.Columns.Item("U_SuppPriBeDis").Cells.Item(pVal.Row).Specific.Value = (bad + ext).ToString()
                            ''Commented by Ramya(05-05-2026)
                            ''Total Calculation
                            Dim rawValue As String = CType(objMatrix.Columns.Item("14").Cells.Item(pVal.Row).Specific.Value, String)

                            ' Keep only digits and decimal point
                            Dim cleanValue As String = System.Text.RegularExpressions.Regex.Replace(rawValue, "[^0-9.]", "")

                            ''objMatrix.Columns.Item("21").Cells.Item(pVal.Row).Specific.Value = (cleanValue - objMatrix.Columns.Item("U_DSAT").Cells.Item(pVal.Row).Specific.Value).ToString()
                            objForm.Freeze(False)
                        End If

                        ' Percentage → Amount
                        ' Percentage → Amount
                        If pVal.ItemUID = "U_ADSUPPR" AndAlso pVal.BeforeAction = False AndAlso LastEditedField = "P" Then
                            objForm.Freeze(True)
                            Dim percentage As Decimal = 0D
                            Dim TBD As Decimal = 0D

                            Dim rawValue As String = CType(objForm.Items.Item("22").Specific, SAPbouiCOM.EditText).Value

                            ' Keep only digits and decimal point
                            Dim cleanValue As String = System.Text.RegularExpressions.Regex.Replace(rawValue, "[^0-9.]", "")

                            Decimal.TryParse(cleanValue, TBD)
                            Decimal.TryParse(CType(objForm.Items.Item("U_ADSUPPR").Specific, SAPbouiCOM.EditText).Value, percentage)

                            Dim result As Decimal = Math.Round((percentage / 100D) * TBD, 3)

                            CType(objForm.Items.Item("U_ADSUPAMT").Specific, SAPbouiCOM.EditText).Value = result.ToString("0.####")
                            LastEditedField = ""
                            objForm.Freeze(False)
                        End If


                        ' Amount → Percentage
                        If pVal.ItemUID = "U_ADSUPAMT" AndAlso pVal.BeforeAction = False AndAlso LastEditedField = "A" Then
                            objForm.Freeze(True)
                            Dim amount As Decimal = 0D
                            Dim TBD As Decimal = 0D

                            Dim rawValue As String = CType(objForm.Items.Item("22").Specific, SAPbouiCOM.EditText).Value

                            ' Keep only digits and decimal point
                            Dim cleanValue As String = System.Text.RegularExpressions.Regex.Replace(rawValue, "[^0-9.]", "")

                            Decimal.TryParse(cleanValue, TBD)
                            Decimal.TryParse(CType(objForm.Items.Item("U_ADSUPAMT").Specific, SAPbouiCOM.EditText).Value, amount)

                            If TBD <> 0D Then
                                Dim result As Decimal = Math.Round((amount / TBD) * 100D, 3)
                                CType(objForm.Items.Item("U_ADSUPPR").Specific, SAPbouiCOM.EditText).Value = result.ToString("0.###")
                            End If
                            LastEditedField = ""
                            objForm.Freeze(False)
                        End If

                        ''New Supplier price calculation
                        'If isUpdating Then Exit Sub
                        ''Commented by Ramya based on new requirement on 06-08-2026
                        'If pVal.ItemUID = "38" And pVal.ColUID = "U_SuppDiscPer" And pVal.BeforeAction = False Then
                        '    objForm.Freeze(True)
                        '    'isUpdating = True
                        '    Dim SUPP As Double = objMatrix.Columns.Item("U_SuppPriBeDis").Cells.Item(pVal.Row).Specific.Value * objMatrix.Columns.Item("U_SuppDiscPer").Cells.Item(pVal.Row).Specific.Value / 100
                        '    'Dim SUPP As Double = objMatrix.Columns.Item("U_SuppPri").Cells.Item(pVal.Row).Specific.Value * objMatrix.Columns.Item("U_SuppDiscPer").Cells.Item(pVal.Row).Specific.Value / 100

                        '    objMatrix.Columns.Item("U_DISCA").Cells.Item(pVal.Row).Specific.Value = (SUPP).ToString()
                        '    objMatrix.Columns.Item("U_suppprice").Cells.Item(pVal.Row).Specific.Value = (objMatrix.Columns.Item("U_SuppPriBeDis").Cells.Item(pVal.Row).Specific.Value - SUPP).ToString()
                        '    objMatrix.Columns.Item("11").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        '    'isUpdating = False
                        '    objForm.Freeze(False)
                        'End If

                        'If pVal.ItemUID = "38" And pVal.ColUID = "U_DISCA" And pVal.BeforeAction = False Then
                        '    objForm.Freeze(True)
                        '    isUpdating = True
                        '    Dim basePrice As Double = CDbl(objMatrix.Columns.Item("U_SuppPriBeDis").Cells.Item(pVal.Row).Specific.Value)
                        '    Dim discAmt As Double = CDbl(objMatrix.Columns.Item("U_DISCA").Cells.Item(pVal.Row).Specific.Value)

                        '    If basePrice <> 0 Then
                        '        Dim discPer As Double = (discAmt / basePrice) * 100

                        '        ''objMatrix.Columns.Item("U_SuppDiscPer").Cells.Item(pVal.Row).Specific.Value = discPer.ToString()
                        '        objMatrix.Columns.Item("U_suppprice").Cells.Item(pVal.Row).Specific.Value = (basePrice - discAmt).ToString()

                        '    End If
                        '    isUpdating = False
                        '    objForm.Freeze(False)
                        'End If

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

                        If pVal.ItemUID = "38" And pVal.ColUID = "1470002145" And pVal.BeforeAction = False Then
                            objForm.Freeze(True)
                            Dim ItemCode As String = Trim(objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value)
                            Dim UOM As String = Trim(objMatrix.Columns.Item("1470002145").Cells.Item(pVal.Row).Specific.Value)


                            Dim Rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            Rs.DoQuery("SELECT TOP 1 T1.""BcdCode"" FROM OBCD T1 INNER JOIN OUOM T2 ON T1.""UomEntry"" = T2.""UomEntry"" WHERE T1.""ItemCode""='" & ItemCode & "' AND T2.""UomCode""='" & UOM & "'")

                            If Rs.RecordCount > 0 Then
                                objMatrix.Columns.Item("4").Cells.Item(pVal.Row).Specific.Value = Rs.Fields.Item("BcdCode").Value
                            End If
                            objForm.Freeze(False)
                        End If



                        'Lost Focus instead of Load button
                        If pVal.ItemUID = "38" And pVal.ColUID = "4" And pVal.BeforeAction = False Then


                            If objMatrix.Columns.Item("4").Cells.Item(pVal.Row).Specific.Value.ToString().Trim <> "" Then
                                objMain.objApplication.StatusBar.SetText("Please wait while loading line details....",
                           SAPbouiCOM.BoMessageTime.bmt_Short,
                           SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                                '    For i As Integer = 1 To objMatrix.VisualRowCount - 1
                                '        If objMatrix.Columns.Item("1470002145").Cells.Item(i).Specific.Value.Trim = "" Then
                                '            objMain.objApplication.StatusBar.SetText("Please select  UomCode for the items " & objMatrix.Columns.Item("0").Cells.Item(i).Specific.Value & " ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                '            Exit Sub
                                '        End If
                                '    Next

                                Dim objComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("2001").Specific

                                If objComboBox.Value = "" Then
                                    objMain.objApplication.StatusBar.SetText("Please select branch at the document header...",
                                    SAPbouiCOM.BoMessageTime.bmt_Short,
                                    SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Exit Sub
                                End If

                                Try
                                    objForm.Freeze(True)

                                    Dim IsUomDefault As String = "Yes"

                                    Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value.ToString().Trim
                                    Dim UOM As String = objMatrix.Columns.Item("1470002145").Cells.Item(pVal.Row).Specific.Value.ToString().Trim
                                    Dim BarCode As String = objMatrix.Columns.Item("4").Cells.Item(pVal.Row).Specific.Value.ToString().Trim

                                    If UOM <> "" Then
                                        If Me.CheckDefaultUOM(UOM, ItemCode) = True Then
                                            IsUomDefault = "Yes"
                                        Else
                                            IsUomDefault = "No"
                                        End If
                                    End If

                                    ''Normal Load
                                    'Me.LoadOnItemOrUOMSelection(objForm.UniqueID, pVal.Row, IsUomDefault, ItemCode, UOM, objForm.Items.Item("4").Specific.Value)

                                    ''Special Prices
                                    'Commeted by Ramya for testing on 10142026
                                    Me.LoadOnItemOrUOMSelection_LostFocus(objForm.UniqueID, pVal.Row, IsUomDefault, ItemCode, UOM, objForm.Items.Item("4").Specific.Value, objForm.Items.Item("10").Specific.Value, BarCode)

                                    ''Me.LoadOnItemOrUOMSelection_SpecialPrices(objForm.UniqueID, pVal.Row, IsUomDefault, ItemCode, UOM, objForm.Items.Item("4").Specific.Value, objForm.Items.Item("10").Specific.Value, BarCode)

                                    ''Commeted by Ramya for testing on 10142026
                                    objMatrix.Columns.Item("U_suppprice").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    ''Holding Days
                                    'Me.UpdateHoldingDays(objForm)



                                Catch ex As Exception

                                    objMain.objApplication.StatusBar.SetText(ex.Message,
                                    SAPbouiCOM.BoMessageTime.bmt_Short,
                                    SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                                Finally

                                    objForm.Freeze(False)

                                End Try
                            End If

                            'Dim totalExcise As Double = 0.0

                            'For i As Integer = 1 To objMatrix.VisualRowCount

                            '    Dim exciseVal As Double = 0.0
                            '    Dim qty As Double = 0.0

                            '    Dim exciseField As SAPbouiCOM.EditText = objMatrix.Columns.Item("U_EXT").Cells.Item(i).Specific
                            '    Dim qtyField As SAPbouiCOM.EditText = objMatrix.Columns.Item("11").Cells.Item(i).Specific

                            '    Double.TryParse(exciseField.Value, exciseVal)
                            '    Double.TryParse(qtyField.Value, qty)

                            '    totalExcise += (exciseVal * qty)

                            'Next

                            'Dim headerExcise As SAPbouiCOM.EditText = objForm.Items.Item("U_TEXTAX").Specific
                            'headerExcise.Value = Math.Round(totalExcise, 3, MidpointRounding.AwayFromZero).ToString("0.000")

                        End If

                        If pVal.ItemUID = "38" And (pVal.ColUID = "4" Or pVal.ColUID = "11") And pVal.BeforeAction = False Then
                            objForm.Freeze(True)
                            Dim totalExcise As Double = 0

                            objMatrix = objForm.Items.Item("38").Specific

                            For i As Integer = 1 To objMatrix.VisualRowCount

                                Dim exciseVal As Double = 0
                                Dim qty As Double = 0

                                Double.TryParse(CType(objMatrix.Columns.Item("U_EXT").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value, exciseVal)
                                Double.TryParse(CType(objMatrix.Columns.Item("11").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value, qty)

                                totalExcise += (exciseVal * qty)

                            Next

                            CType(objForm.Items.Item("U_TEXTAX").Specific, SAPbouiCOM.EditText).Value = Math.Round(totalExcise, 3, MidpointRounding.AwayFromZero).ToString("0.000")
                            objForm.Freeze(False)
                        End If

                        If pVal.ItemUID = "38" And (pVal.ColUID = "U_normalqty" Or pVal.ColUID = "U_FOCQTY") And pVal.BeforeAction = False Then
                            objForm.Freeze(True)
                            objMatrix = objForm.Items.Item("38").Specific

                            Dim NormalQty As Double = 0
                            Dim FOCQty As Double = 0

                            Double.TryParse(objMatrix.Columns.Item("U_normalqty").Cells.Item(pVal.Row).Specific.Value.ToString(), NormalQty)
                            Double.TryParse(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(pVal.Row).Specific.Value.ToString(), FOCQty)

                            objMatrix.Columns.Item("U_SQTY").Cells.Item(pVal.Row).Specific.Value = (NormalQty + FOCQty).ToString()
                            'Header U_SQTY Calculation
                            Dim TotalSQTY As Double = 0

                            For i As Integer = 1 To objMatrix.VisualRowCount

                                Dim LineSQTY As Double = 0

                                Double.TryParse(objMatrix.Columns.Item("U_SQTY").Cells.Item(i).Specific.Value.ToString(), LineSQTY)

                                TotalSQTY += LineSQTY

                            Next

                            Dim oHeaderSQTY As SAPbouiCOM.EditText = objForm.Items.Item("U_SQTY").Specific
                            oHeaderSQTY.Value = TotalSQTY.ToString()
                            objForm.Freeze(False)
                        End If


                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    End Try

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

                    'Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                    '    Try
                    '        objForm = objMain.objApplication.Forms.Item(FormUID)
                    '        objMatrix = objForm.Items.Item("38").Specific


                    '        If pVal.ItemUID = "38" AndAlso pVal.ColUID = "4" AndAlso pVal.BeforeAction = True Then
                    '            Try
                    '                objForm.Freeze(True)


                    '                System.Threading.Thread.Sleep(50)


                    '                objMain.objApplication.ActivateMenuItem("7425")

                    '            Catch ex As Exception

                    '            Finally
                    '                objForm.Freeze(False)
                    '            End Try
                    '        End If

                    '    Catch ex As Exception
                    '        ' Handle exception
                    '    End Try


                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    Dim AppTemForm As SAPbouiCOM.Form
                    AppTemForm = objMain.objApplication.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = AppTemForm.ChooseFromLists.Item(CFL_Id)
                    If pVal.ItemUID = "38" AndAlso pVal.ColUID = "4" AndAlso pVal.BeforeAction = True And oCFL.UniqueID = "8" Then
                        Dim Barcode As String = CType(objMatrix.Columns.Item("4").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.Trim()

                        If Barcode = "" Then
                            BubbleEvent = False

                        End If

                    End If
                    If pVal.ItemUID = "38" AndAlso pVal.ColUID = "1" AndAlso pVal.BeforeAction = True And oCFL.UniqueID = "6" Then
                        Dim ITcode As String = CType(objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.Trim()

                        If ITcode = "" Then
                            BubbleEvent = False

                        End If

                    End If
                    'Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    '    Dim AppTemForm As SAPbouiCOM.Form
                    '    AppTemForm = objMain.objApplication.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                    '    Dim objMatrix As SAPbouiCOM.Matrix
                    '    objMatrix = AppTemForm.Items.Item("38").Specific
                    '    Dim oCFL As SAPbouiCOM.ChooseFromList
                    '    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    '    Dim CFL_Id As String
                    '    CFL_Id = CFLEvent.ChooseFromListUID
                    '    oCFL = AppTemForm.ChooseFromLists.Item(CFL_Id)
                    '    Dim oDT As SAPbouiCOM.DataTable
                    '    oDT = CFLEvent.SelectedObjects
                    '    If oCFL.UniqueID = "6" Then
                    '        Try
                    '            Dim Query As String = " SELECT T1.""ItemCode"", T1.""VendorCode"" FROM OITM T0  INNER JOIN ITM2 T1 ON T0.""ItemCode"" = T1.""ItemCode"" WHERE T1.""VendorCode""='" & AppTemForm.Items.Item("4").Specific.Value.ToString().Trim() & "'"
                    '            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '            oRsQuery.DoQuery(Query)

                    '            Dim oConditions As SAPbouiCOM.Conditions
                    '            Dim oCondition As SAPbouiCOM.Condition
                    '            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
                    '            Dim emptyCon As New SAPbouiCOM.Conditions
                    '            oChooseFromList = AppTemForm.ChooseFromLists.Item("6")
                    '            oChooseFromList.SetConditions(emptyCon)
                    '            oConditions = oChooseFromList.GetConditions()

                    '            For i As Integer = 1 To oRsQuery.RecordCount
                    '                oCondition = oConditions.Add()
                    '                oCondition.Alias = "ItemCode"
                    '                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    '                oCondition.CondVal = oRsQuery.Fields.Item(0).Value.ToString().Trim
                    '                If i <> oRsQuery.RecordCount Then
                    '                    oCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    '                End If
                    '                oRsQuery.MoveNext()
                    '            Next
                    '            oChooseFromList.SetConditions(oConditions)
                    '            'Dim GetDocEntry As String = "Select ""FrgnName"" From OITM Where ""ItemCode"" = '" & oRsQuery.Fields.Item("ItemCode").Value & "' "
                    '            'Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '            'oRs.DoQuery(GetDocEntry)
                    '            'objMatrix.Columns.Item("U_ForeNm").Cells.Item(pVal.Row).Specific.Value() = oRs.Fields.Item("FrgnName").Value()
                    '        Catch ex As Exception

                    '        End Try

                    '    End If
                    '    If oCFL.UniqueID = "2" And pVal.BeforeAction = True Then

                    '        Try
                    '            oDBs_Details = objForm.DataSources.DBDataSources.Item("POR1")
                    '            objMatrix = objForm.Items.Item("38").Specific
                    '            Dim ItemCode As String
                    '            Dim ItemCode2 As String
                    '            Dim ItemCode3 As String
                    '            For i As Integer = 1 To objMatrix.VisualRowCount
                    '                ItemCode = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.ToString()
                    '                If i = 1 Then                                   '
                    '                    ItemCode2 = ItemCode
                    '                ElseIf i > 1 Then
                    '                    ItemCode2 = ItemCode & "'" & "," & "'" & ItemCode2

                    '                End If
                    '            Next

                    '            Dim Query As String = "SELECT * from SBO_Preferredvendor() WHERE ""ITEMCODE"" IN ('" & ItemCode2 & "')"
                    '            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '            oRsQuery.DoQuery(Query)

                    '            Dim oConditions As SAPbouiCOM.Conditions
                    '            Dim oCondition As SAPbouiCOM.Condition
                    '            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
                    '            Dim emptyCon As New SAPbouiCOM.Conditions
                    '            oChooseFromList = AppTemForm.ChooseFromLists.Item("2")
                    '            oChooseFromList.SetConditions(emptyCon)
                    '            oConditions = oChooseFromList.GetConditions()

                    '            For i As Integer = 1 To oRsQuery.RecordCount
                    '                oCondition = oConditions.Add()
                    '                oCondition.Alias = "CardCode"
                    '                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    '                oCondition.CondVal = oRsQuery.Fields.Item(0).Value.ToString().Trim
                    '                If i <> oRsQuery.RecordCount Then
                    '                    oCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    '                End If
                    '                oRsQuery.MoveNext()
                    '            Next
                    '            oChooseFromList.SetConditions(oConditions)
                    '        Catch ex As Exception
                    '        End Try
                    '    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub



    Private Sub LoadOnItemOrUOMSelection(ByVal FormUID As String, ByVal RowId As Integer, ByVal IsUomDefault As String, ByVal ItemCode As String, ByVal Uom As String, ByVal Vendor As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("OPOR")
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
            If IsUomDefault = "No" Then
                UOMEntry = Me.GetUomEntry(Uom)
                'Dim GetPurchasePrice As String = " Select T0.""Price"" From ITM9 T0 Where T0.""PriceList""= " & _
                '                                 " (Select ""ListNum"" From OCRD Where ""CardCode"" = '" & Vendor & "') And " & _
                '                                 " T0.""UomEntry""='" & UOMEntry & "' And ""ItemCode""='" & ItemCode & "' "
                'Dim oRsGetPurchasePrice As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oRsGetPurchasePrice.DoQuery(GetPurchasePrice)
                'If oRsGetPurchasePrice.RecordCount > 0 Then
                '    SalesPrice = CDbl(oRsGetPurchasePrice.Fields.Item(0).Value)
                'Else
                '    SalesPrice = Me.getSupplierPrice(objForm.UniqueID, IsUomDefault, ItemCode, Vendor, UOMEntry, Uom)
                'End If
                Dim GetPurchasePrice As String = " Select IFNULL(T3.""Price"",0) ""Price"" from SPP2 T3  " &
                                                 " Where T3.""ItemCode"" = '" & ItemCode & "' and T3.""CardCode"" = '" & Vendor & "' and T3.""UomEntry"" = '" & UOMEntry & "' "
                Dim oRsGetPurchasePrice As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetPurchasePrice.DoQuery(GetPurchasePrice)
                If oRsGetPurchasePrice.RecordCount > 0 Then
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
                'If PCPrPack <> "" Then
                '    Packing = PCPrPack & "*" & oRsGetPackageCode.Fields.Item(0).Value
                'End If
                'Packing = oRsGetPackageCode.Fields.Item(0).Value


                Dim GetPackageCode As String = "Select ""U_packing"" FROM OITM Where ""ItemCode""='" & ItemCode.Trim & "'"
                Dim oRsGetPackageCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetPackageCode.DoQuery(GetPackageCode)

                Packing = oRsGetPackageCode.Fields.Item(0).Value


                Dim GetRetailPriceBrnchWise As String = " Select T1.""Price"" From OITM T0 Inner Join ITM9 T1 On T1.""ItemCode"" = T0.""ItemCode"" " &
                                                        " Inner Join OBPL T2 On T1.""PriceList""=T2.""U_PricelistNo"" Where T0.""ItemCode"" = '" & ItemCode & "' " &
                                                        " And T2.""BPLId""= '" & objForm.Items.Item("2001").Specific.Value.Trim & "' And T1.""UomEntry""='" & UOMEntry & "'"
                Dim oRsGetRetailPriceBrnchWise As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetRetailPriceBrnchWise.DoQuery(GetRetailPriceBrnchWise)
                If oRsGetRetailPriceBrnchWise.RecordCount > 0 Then
                    RetailsPrice = oRsGetRetailPriceBrnchWise.Fields.Item("Price").Value
                End If
                'Dim GetLeastVendorDtlsAndPrice As String = " Select Top 1 T0.""ItemCode"" , T1.""CardCode"", T1.""CardName"" , T0.""Price"" From ITM9 T0 " & _
                '                                          " Inner Join OCRD T1 On T1.""ListNum"" = T0.""PriceList""  And T1.""CardCode"" In (Select ""VendorCode"" From ITM2 Where  ITM2.""ItemCode"" = T0.""ItemCode"") Where " & _
                '                                          " ""ItemCode"" = '" & ItemCode & "' And T1.""CardType"" = 'S' And  T0.""UomEntry"" = '" & UOMEntry & "' " & _
                '                                          " Order By T0.""Price""  Asc "
                'Dim GetLeastVendorDtlsAndPrice As String = " Select TOP 1 ""Price"", ""BaseCard"" As ""CardCode"" , (Select ""CardName"" From OCRD Where ""CardCode"" = ""BaseCard"" ) As ""CardName"" " & _
                '                                           " From PDN1 Where ""ItemCode""='" & ItemCode & "' And ""DocDate"">(SELECT ADD_DAYS (TO_DATE (CURRENT_DATE, 'YYYY-MM-DD'), -180) From DUMMY) " & _
                '                                           " Order By ""Price""  Asc "
                'Dim GetLeastVendorDtlsAndPrice As String = " Select  TOP 1  Round(Sum(""Price""/""NumPerMsr""),2)  , ""BaseCard"" As ""CardCode"" , ""DocEntry"" , (Select ""CardName"" From OCRD Where ""CardCode"" = ""BaseCard"" ) As ""CardName"" From PDN1 Where  ""ItemCode""= '" & ItemCode & "'  And " & _
                '                                           " ""DocDate"">(Select ADD_DAYS (TO_DATE (CURRENT_DATE, 'YYYY-MM-DD'), -180)  From DUMMY)  " & _
                '                                           " Group by ""BaseCard"",""DocEntry"" Order By SUM(""Price""/""NumPerMsr"")"
                'Dim GetLeastVendorDtlsAndPrice As String = " Select  TOP 1  Sum(""U_suppprice""/""NumPerMsr"") , ""BaseCard"" As ""CardCode"" , ""DocEntry"" , (Select ""CardName"" From OCRD Where ""CardCode"" = ""BaseCard"" ) As ""CardName"" From PDN1 Where  ""ItemCode""= '" & ItemCode & "'  And " & _
                '                                           " ""DocDate"">(Select ADD_DAYS (TO_DATE (CURRENT_DATE, 'YYYY-MM-DD'), -180)  From DUMMY)  " & _
                '                                           " Group by ""BaseCard"",""DocEntry"" , ""U_suppprice"" Order By SUM(""U_suppprice""/""NumPerMsr"")"


                Dim GetLeastVendorDtlsAndPrice As String = " Select  TOP 1  Sum(""U_suppprice""/""NumPerMsr"") , ""BaseCard"" As ""CardCode"" , ""DocEntry"" , (Select ""CardName"" From OCRD Where ""CardCode"" = ""BaseCard"" ) As ""CardName"" From PDN1 Where  ""ItemCode""= '" & ItemCode & "'  And " &
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

                'If oRsGetPurchasePriceDfltUOM.RecordCount > 0 Then
                '    SalesPrice = CDbl(oRsGetPurchasePriceDfltUOM.Fields.Item(0).Value)
                'Else                     
                '    SalesPrice = Me.getSupplierPrice(objForm.UniqueID, IsUomDefault, ItemCode, Vendor, UOMEntry, Uom)
                'End If
                Dim GetPurchasePriceDfltUOM As String = " Select IFNULL(T3.""Price"",0) ""Price"" from SPP2 T3 " &
               " Where ""ItemCode"" = '" & ItemCode & "' and ""CardCode"" = '" & Vendor & "' and T3.""UomEntry"" = '" & UOMEntry & "' "

                Dim oRsGetPurchasePriceDfltUOM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetPurchasePriceDfltUOM.DoQuery(GetPurchasePriceDfltUOM)

                If oRsGetPurchasePriceDfltUOM.RecordCount > 0 Then
                    SalesPrice = CDbl(oRsGetPurchasePriceDfltUOM.Fields.Item(0).Value)
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
                '                                            " Inner Join OCRD T1 On T1.""ListNum"" = T0.""PriceList"" And T1.""CardCode"" In (Select ""VendorCode"" From ITM2 Where  ITM2.""ItemCode"" = T0.""ItemCode"") Where " & _
                '                                            " ""ItemCode"" = '" & ItemCode & "' And T1.""CardType"" = 'S' " & _
                '                                            " Order By T0.""Price""  Asc "
                'Dim GetLeastVendorDtlsAndPrice As String = " Select TOP 1 ""Price"", ""BaseCard"" As ""CardCode"" , (Select ""CardName"" From OCRD Where ""CardCode"" = ""BaseCard"" ) As ""CardName"" " & _
                '                                           " From PDN1 Where ""ItemCode""='" & ItemCode & "' And ""DocDate"">(SELECT ADD_DAYS (TO_DATE (CURRENT_DATE, 'YYYY-MM-DD'), -180) From DUMMY) " & _
                '                                           " Order By ""Price"" Asc "
                'Dim GetLeastVendorDtlsAndPrice As String = " Select  TOP 1  Round(Sum(""U_suppprice""/""NumPerMsr""),2)  , ""BaseCard"" As ""CardCode"" , ""DocEntry"" , (Select ""CardName"" From OCRD Where ""CardCode"" = ""BaseCard"" ) As ""CardName"" From PDN1 Where  ""ItemCode""= '" & ItemCode & "'  And " &
                '                                           " ""DocDate"">(Select ADD_DAYS (TO_DATE (CURRENT_DATE, 'YYYY-MM-DD'), -180)  From DUMMY)  " &
                '                                           " Group by ""BaseCard"",""DocEntry""  , ""U_suppprice"" Order By SUM(""U_suppprice""/""NumPerMsr"")"
                'Dim GetLeastVendorDtlsAndPrice As String = " Select  TOP 1  Round(Sum(""U_suppprice""/""NumPerMsr""),2)  , ""BaseCard"" As ""CardCode"" , ""DocEntry"" , (Select ""CardName"" From OCRD Where ""CardCode"" = ""BaseCard"" ) As ""CardName"" From PDN1 Where  ""ItemCode""= '" & ItemCode & "'  And " &
                '                                           " ""DocDate"">(Select ADD_DAYS (TO_DATE (CURRENT_DATE, 'YYYY-MM-DD'), -45)  From DUMMY)  " &
                '                                           " Group by ""BaseCard"",""DocEntry""  , ""U_suppprice"" Order By SUM(""U_suppprice""/""NumPerMsr"")"

                Dim GetLeastVendorDtlsAndPrice As String = " Select  TOP 1  Round(Sum(""U_suppprice""/""NumPerMsr""),2)  , ""BaseCard"" As ""CardCode"" , ""DocEntry"" , (Select ""CardName"" From OCRD Where ""CardCode"" = ""BaseCard"" ) As ""CardName"" From PDN1 Where  ""ItemCode""= '" & ItemCode & "'  And " &
                                                           " ""DocDate"">(Select ADD_DAYS (TO_DATE (CURRENT_DATE, 'YYYY-MM-DD'), -45)  From DUMMY)  " &
                                                           " Group by ""BaseCard"",""DocEntry""  , ""U_suppprice"" Order By SUM(""U_suppprice""/""NumPerMsr"")"

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

            Dim LPP As Double = 0

            Dim GetLastLPP As String = "Select IFNULL((Select (case when B.""UomEntry""=X.""UomEntry""  and X.""UomEntry""=Y.""BaseUom"" then B.""Price""  " &
                                            "when B.""UomEntry""=X.""UomEntry""  and X.""UomEntry""<>Y.""BaseUom"" then B.""Price""*1/X.""AltQty""*X.""BaseQty""  " &
                                            "when B.""UomEntry""<>X.""UomEntry"" and X.""UomEntry""<>Y.""BaseUom"" then B.""Price""*1/X.""AltQty""*X.""BaseQty""  " &
                                            "when B.""UomEntry""<>X.""UomEntry"" and X.""UomEntry""=Y.""BaseUom"" then B.""Price"" " &
                                            "else 0 end) from OITM C inner join (SELECT H.""ItemCode"",IFNULL((IFNULL(H.""INMPrice"", 0)/H.""NumPerMsr""),0)-IFNULL((IFNULL(A1.""INMPrice"", 0)/A1.""NumPerMsr""),0) as ""Price"", H.""UomCode"",H.""UomEntry"",H.""NumPerMsr"",Row_Number() over  " &
                                            "(partition by H.""ItemCode"" order by H.""DocEntry"" Desc) as ""RN"" FROM ""PDN1"" H left join OPDN A3 on A3.""DocEntry""=H.""DocEntry"" left join ORPC A2 on  A2.""U_GRPO""=A3.""DocNum"" " &
                                            "left join RPC1 A1 on A2.""DocEntry""=A1.""DocEntry"" and A1.""ItemCode""=H.""ItemCode"" and A1.""NoInvtryMv""='Y' where H.""DocDate""<'" & objForm.Items.Item("10").Specific.Value & "' and IFNULL(H.""U_normalqty"",0)<>0)B on C.""ItemCode""=B.""ItemCode"" left join OUGP Y " &
                                            "on Y.""UgpEntry""=C.""UgpEntry""  left JOIN UGP1 X ON Y.""UgpEntry"" = X.""UgpEntry"" where ""RN""=1  " &
                                            "and  B.""ItemCode"" ='" & ItemCode & "' and  X.""UomEntry""='" & UOMEntry & "'),0) as ""LPP"" FROM DUMMY"

            Dim oRsGetLastLPP As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oRsGetLastLPP.DoQuery(GetLastLPP)

            If oRsGetLastLPP.RecordCount > 0 Then
                LPP = oRsGetLastLPP.Fields.Item(0).Value
            End If

            'If LPP <> 0 Then

            '    objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value = LPP

            'Else

            '    If SalesPrice <> 0 Then
            '        objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value = SalesPrice

            '    Else
            '        If Me.CheckDefaultUOM(Uom, ItemCode) = True Then '---Same UomCode in Po & Itemmaster 
            '            IsUomDefault = "Yes"
            '            UOMEntry = Me.GetUomEntry(Me.GetDefaultUomCode(ItemCode))
            '            Dim GetPurchasePrice As String = " Select T0.""Price""From ITM1 T0 Where T0.""PriceList""= " &
            '                                                                " (Select ""ListNum"" From OCRD Where ""CardCode"" = '" & Vendor & "') And " &
            '                                                                " ""ItemCode""='" & ItemCode & "' and T0.""UomEntry"" = '" & UOMEntry & "' "
            '            Dim oRsGetPurchasePrice As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '            oRsGetPurchasePrice.DoQuery(GetPurchasePrice)
            '            If oRsGetPurchasePrice.RecordCount > 0 Then
            '                objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value = oRsGetPurchasePrice.Fields.Item(0).Value
            '            Else
            '                objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value = 0
            '            End If
            '        Else
            '            IsUomDefault = "No"
            '            Dim GetPurchasePrice As String = " Select T0.""Price"" From ITM9 T0 Where T0.""PriceList""= " &
            '                                             " (Select ""ListNum"" From OCRD Where ""CardCode"" = '" & Vendor & "') And " &
            '                                             " T0.""UomEntry""='" & UOMEntry & "' And ""ItemCode""='" & ItemCode & "' "
            '            Dim oRsGetPurchasePrice As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '            oRsGetPurchasePrice.DoQuery(GetPurchasePrice)
            '            If oRsGetPurchasePrice.RecordCount > 0 Then
            '                objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value = oRsGetPurchasePrice.Fields.Item(0).Value
            '            Else
            '                objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value = 0

            '            End If
            '        End If

            '    End If

            'End If



            objMatrix.Columns.Item("U_Packing").Cells.Item(RowId).Specific.Value = Packing

            objMatrix.Columns.Item("U_RetailPrice").Cells.Item(RowId).Specific.Value = RetailsPrice
            'objMatrix.Columns.Item("U_Addbarcode").Cells.Item(RowId).Specific.Value = AdditionalBarCode
            Dim LineValidation As String = "Select ""U_Least""as ""Least"" from ""@ME_ACFG"""
            Dim oLValidation As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oLValidation.DoQuery(LineValidation)
            If oLValidation.Fields.Item("Least").Value = "Y" Then
                objMatrix.Columns.Item("U_LVNDPRC").Cells.Item(RowId).Specific.Value = LeastVendorPrice
                objMatrix.Columns.Item("U_LPVCD").Cells.Item(RowId).Specific.Value = LeastVendorCode
                objMatrix.Columns.Item("U_LPVNM").Cells.Item(RowId).Specific.Value = LeastVendorName
                Dim NetCost As String = objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value
                If NetCost <> "" Then
                    NetCost = NetCost.Replace(objForm.Items.Item("63").Specific.Value, "")
                    NetCost = Convert.ToDouble(NetCost.Trim)
                    Dim LeastPrice As Double = objMatrix.Columns.Item("U_LVNDPRC").Cells.Item(RowId).Specific.Value
                    If LeastPrice = 0 Then
                        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LPValidate")
                        oColumn.Cells.Item(RowId).Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If
                    If NetCost > 0 And LeastPrice > 0 Then
                        If NetCost > LeastPrice Then
                            Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LPValidate")
                            oColumn.Cells.Item(RowId).Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        Else
                            Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LPValidate")
                            oColumn.Cells.Item(RowId).Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        End If
                    End If
                End If
                '-------------------Below cost---------------------------
                If objMatrix.Columns.Item("14").Cells.Item(RowId).Specific.Value = "" Then
                    NetCost = 0
                Else
                    NetCost = objMatrix.Columns.Item("14").Cells.Item(RowId).Specific.Value
                    NetCost = NetCost.Replace(objForm.Items.Item("63").Specific.Value, "")
                    NetCost = Convert.ToDouble(NetCost.Trim)
                End If
                Dim RetailPrice As Double = objMatrix.Columns.Item("U_RetailPrice").Cells.Item(RowId).Specific.Value
                If RetailPrice > 0 And NetCost > 0 Then
                    If NetCost > RetailPrice Then
                        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_BelowCst")
                        oColumn.Cells.Item(RowId).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If
                End If
            End If
            '-------------------Below cost---------------------------
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
                    Dim LeastPrice As Double = Convert.ToDouble(objMatrix.Columns.Item("U_LVNDPRC").Cells.Item(i).Specific.Value)
                    If LeastPrice = 0 Then
                        If objMatrix.Columns.Item("U_LVNDPRC").Visible = True Then
                            Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LPValidate")
                            oColumn.Cells.Item(i).Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        End If

                    End If
                    If NetCost > 0 And LeastPrice > 0 Then
                        If NetCost > LeastPrice Then
                            Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LPValidate")
                            oColumn.Cells.Item(i).Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        Else
                            Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LPValidate")
                            oColumn.Cells.Item(i).Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue)
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
                    Dim RetailPrice As Double = objMatrix.Columns.Item("U_RetailPrice").Cells.Item(i).Specific.Value
                    If RetailPrice > 0 And NetCost > 0 Then
                        If NetCost > RetailPrice Then
                            Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_BelowCst")
                            oColumn.Cells.Item(i).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub AddItems(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim objMatrix As SAPbouiCOM.Matrix
            objMatrix = objForm.Items.Item("38").Specific
            'objMain.objUtilities.AddButton(objForm.UniqueID, "btn_LdItd", objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 10, _
            '                              objForm.Items.Item("38").Left, objForm.Items.Item("2").Width + 30, "116", _
            '                               "Load &Item Details", objForm.Items.Item("38").FromPane, objForm.Items.Item("38").ToPane)

            oItem1 = objForm.Items.Add("btn_LdItd", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem1.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 10
            oItem1.Left = objForm.Items.Item("38").Left
            oItem1.Width = objForm.Items.Item("2").Width + 30
            oItem1.Height = objForm.Items.Item("2").Height
            oItem1.FromPane = objForm.Items.Item("38").FromPane
            oItem1.ToPane = objForm.Items.Item("38").ToPane
            Dim btn1 As SAPbouiCOM.Button = objForm.Items.Item("btn_LdItd").Specific
            btn1.Caption = "Load &Item Details"

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
                                        objForm.Items.Item("btn_calFOC").Width + objForm.Items.Item("btn_calFOC").Left + 5, "130", objForm.Items.Item("2").Height, "btn_calFOC", "Calculate Margin", objForm.Items.Item("38").FromPane, objForm.Items.Item("38").ToPane)


            'supplier rebate button hidden
            'oItem3 = objForm.Items.Add("btn_Supp", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'oItem3.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 10
            'oItem3.Left = objForm.Items.Item("btn_MRG").Width + objForm.Items.Item("btn_MRG").Left + 5
            'oItem3.Width = objForm.Items.Item("2").Width + 30
            'oItem3.Height = objForm.Items.Item("2").Height
            'oItem3.FromPane = objForm.Items.Item("38").FromPane
            'oItem3.ToPane = objForm.Items.Item("38").ToPane
            'oItem3.LinkTo = "btn_MRG"
            'Dim btn3 As SAPbouiCOM.Button = objForm.Items.Item("btn_Supp").Specific
            'btn3.Caption = "Supplier &Rebate"

            'oItem3 = objForm.Items.Add("btn_Supp", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'oItem3.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 10
            'oItem3.Left = objForm.Items.Item("btn_LdItd").Width + objForm.Items.Item("btn_calFOC").Width + objForm.Items.Item("btn_LdItd").Left + 10
            'oItem3.Width = objForm.Items.Item("2").Width + 30
            'oItem3.Height = objForm.Items.Item("2").Height
            'oItem3.FromPane = objForm.Items.Item("38").FromPane
            'oItem3.ToPane = objForm.Items.Item("38").ToPane
            'Dim btn3 As SAPbouiCOM.Button = objForm.Items.Item("btn_Supp").Specific
            'btn3.Caption = "Supplier &Rebate"

            'objMain.objUtilities.AddButton(objForm.UniqueID, "btn_calFOC", objForm.Items.Item("btn_LdItd").Top,
            '                               objForm.Items.Item("btn_LdItd").Width + objForm.Items.Item("btn_LdItd").Left + 5, "100", "116",
            '                               "FOC Calculation", objForm.Items.Item("38").FromPane, objForm.Items.Item("38").ToPane)

            oItem = objForm.Items.Add("fldSale", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem.Top = objForm.Items.Item("1320002137").Top
            oItem.Left = objForm.Items.Item("1320002137").Left - objForm.Items.Item("1320002137").Width + 10
            oItem.Width = objForm.Items.Item("1320002137").Width + 30
            oItem.Height = objForm.Items.Item("1320002137").Height
            oItem.AffectsFormMode = False
            Dim fld As SAPbouiCOM.Folder = objForm.Items.Item("fldSale").Specific
            fld.Caption = "Average Sale"
            fld.GroupWith("1320002137")
            fld.Pane = "1"


            oItem = objForm.Items.Add("Grid", SAPbouiCOM.BoFormItemTypes.it_GRID)
            oItem.Top = objForm.Items.Item("38").Top
            oItem.Left = objForm.Items.Item("38").Left
            oItem.Width = objForm.Items.Item("38").Width
            oItem.Height = objForm.Items.Item("38").Height
            oItem.FromPane = 24
            oItem.ToPane = 24
            objForm.DataSources.DataTables.Add("MyDataTable")

            'Me.ButtonFoc(objForm.UniqueID)
            Dim Validation As String = "Select ""U_PO""as ""PO"" from ""@ME_ACFG"""
            Dim oLValidation As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oLValidation.DoQuery(Validation)
            If oLValidation.Fields.Item("PO").Value = "Y" Then
                objForm.Items.Item("btn_calFOC").Visible = True
            Else
                objForm.Items.Item("btn_calFOC").Visible = False
            End If

            'madhu supplier po total
            ''-madhu comment suplier po total now not required as per client
            'oItem = objForm.Items.Add("lbl", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            'oItem.Top = objForm.Items.Item("86").Top + objForm.Items.Item("86").Height + 20
            'oItem.Left = objForm.Items.Item("86").Left
            'oItem.Width = objForm.Items.Item("86").Width
            'oItem.Height = 15
            'Dim lbl As SAPbouiCOM.StaticText = objForm.Items.Item("lbl").Specific
            'lbl.Caption = "Supplier PO Total"
            'oItem = objForm.Items.Add("U_SPOTOTAL", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oItem.Top = objForm.Items.Item("46").Top + objForm.Items.Item("46").Height + 20
            'oItem.Left = objForm.Items.Item("46").Left
            'oItem.Width = objForm.Items.Item("46").Width
            'oItem.Height = 15
            'Dim editbox As SAPbouiCOM.EditText = objForm.Items.Item("U_SPOTOTAL").Specific
            'editbox.DataBind.SetBound(True, "OPOR", "U_SPOTOTAL")
            'editbox.Item.Enabled = False

            ''--
            '' Total Untaxed Amount Label
            'oItem = objForm.Items.Add("lblUntx", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            'oItem.Top = objForm.Items.Item("lbl").Top + 40
            'oItem.Left = objForm.Items.Item("lbl").Left
            'oItem.Width = objForm.Items.Item("lbl").Width
            'oItem.Height = 15

            'Dim lblUntx As SAPbouiCOM.StaticText = oItem.Specific
            'lblUntx.Caption = "Total Untaxed Amount"

            '' Total Untaxed Amount Field
            'oItem = objForm.Items.Add("TUTXAMT", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oItem.Top = objForm.Items.Item("U_SPOTOTAL").Top + 40
            'oItem.Left = objForm.Items.Item("U_SPOTOTAL").Left
            'oItem.Width = objForm.Items.Item("U_SPOTOTAL").Width
            'oItem.Height = 15

            'Dim txtUntx As SAPbouiCOM.EditText = oItem.Specific
            'txtUntx.DataBind.SetBound(True, "OPOR", "U_TUTXAMT")   ' UDF name
            'txtUntx.Item.Enabled = False
            ''--
            '' Total Excise Tax Label
            'oItem = objForm.Items.Add("lblTax", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            'oItem.Top = objForm.Items.Item("lblUntx").Top + 18
            'oItem.Left = objForm.Items.Item("lblUntx").Left
            'oItem.Width = objForm.Items.Item("lblUntx").Width
            'oItem.Height = 15

            'Dim lblTax As SAPbouiCOM.StaticText = oItem.Specific
            'lblTax.Caption = "Total Excise Tax"

            '' Total Excise Tax Field
            'oItem = objForm.Items.Add("TEXTAX", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oItem.Top = objForm.Items.Item("TUTXAMT").Top + 18
            'oItem.Left = objForm.Items.Item("TUTXAMT").Left
            'oItem.Width = objForm.Items.Item("TUTXAMT").Width
            'oItem.Height = 15

            'Dim txtTax As SAPbouiCOM.EditText = oItem.Specific
            'txtTax.DataBind.SetBound(True, "OPOR", "U_TEXTAX")   ' UDF name
            'txtTax.Item.Enabled = False

            '---------------------madhu
            '-------------------- Supplier PO Total (HIDDEN as per client) --------------------
            ' '-madhu comment supplier po total now not required as per client
            ' oItem = objForm.Items.Add("lbl", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            ' oItem.Top = objForm.Items.Item("86").Top + objForm.Items.Item("86").Height + 20
            ' oItem.Left = objForm.Items.Item("86").Left
            ' oItem.Width = objForm.Items.Item("86").Width
            ' oItem.Height = 15

            ' Dim lbl As SAPbouiCOM.StaticText = objForm.Items.Item("lbl").Specific
            ' lbl.Caption = "Supplier PO Total"

            ' oItem = objForm.Items.Add("U_SPOTOTAL", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            ' oItem.Top = objForm.Items.Item("46").Top + objForm.Items.Item("46").Height + 20
            ' oItem.Left = objForm.Items.Item("46").Left
            ' oItem.Width = objForm.Items.Item("46").Width
            ' oItem.Height = 15

            ' Dim editbox As SAPbouiCOM.EditText = objForm.Items.Item("U_SPOTOTAL").Specific
            ' editbox.DataBind.SetBound(True, "OPOR", "U_SPOTOTAL")
            ' editbox.Item.Enabled = False
            '-------------------------------------------------------------------------------


            '-------------------- Total Untaxed Amount --------------------
            oItem = objForm.Items.Add("lblUntx", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Top = objForm.Items.Item("86").Top + objForm.Items.Item("86").Height + 60
            oItem.Left = objForm.Items.Item("86").Left
            oItem.Width = objForm.Items.Item("86").Width
            oItem.Height = 15

            Dim lblUntx As SAPbouiCOM.StaticText = oItem.Specific
            lblUntx.Caption = "Total Untaxed Amount"

            oItem = objForm.Items.Add("TUTXAMT", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = objForm.Items.Item("46").Top + objForm.Items.Item("46").Height + 60
            oItem.Left = objForm.Items.Item("46").Left
            oItem.Width = objForm.Items.Item("46").Width
            oItem.Height = 15

            Dim txtUntx As SAPbouiCOM.EditText = oItem.Specific
            txtUntx.DataBind.SetBound(True, "OPOR", "U_TUTXAMT")
            txtUntx.Item.Enabled = False
            '------------------------------------------------------------


            '-------------------- Total Excise Tax --------------------
            oItem = objForm.Items.Add("lblTax", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Top = objForm.Items.Item("lblUntx").Top + 18
            oItem.Left = objForm.Items.Item("lblUntx").Left
            oItem.Width = objForm.Items.Item("lblUntx").Width
            oItem.Height = 15

            Dim lblTax As SAPbouiCOM.StaticText = oItem.Specific
            lblTax.Caption = "Total Excise Tax"

            oItem = objForm.Items.Add("TEXTAX", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = objForm.Items.Item("TUTXAMT").Top + 18
            oItem.Left = objForm.Items.Item("TUTXAMT").Left
            oItem.Width = objForm.Items.Item("TUTXAMT").Width
            oItem.Height = 15

            Dim txtTax As SAPbouiCOM.EditText = oItem.Specific
            txtTax.DataBind.SetBound(True, "OPOR", "U_TEXTAX")
            txtTax.Item.Enabled = False
            '------------------------------------------------------------


            'Me.CFLFilterSTG(objForm.UniqueID)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    Private Sub DisableCoulmns(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("38").Specific
            Try
                'objForm.Items.Item("U_SPOTOTAL").Enabled = False
                'objMatrix.Columns.Item("U_INVTRRLN").Visible = False
                'objMatrix.Columns.Item("U_INVTRREQ").Visible = False
                'objMatrix.Columns.Item("U_INVTRREN").Visible = False
                'objMatrix.Columns.Item("U_MEOPNQTY").Visible = False
                'objMatrix.Columns.Item("U_MERCEQTY").Visible = False
                'objMatrix.Columns.Item("U_MEREQQTY").Visible = False
                'objMatrix.Columns.Item("U_METSTQTY").Visible = False
                'objMatrix.Columns.Item("U_MIEntry").Visible = False
                'objMatrix.Columns.Item("U_MIDocNum").Visible = False
                '---
                'objMatrix.Columns.Item("U_METWHSCD").Visible = False
                'objMatrix.Columns.Item("U_METWHSNM").Visible = False
                'objMatrix.Columns.Item("U_METBPID").Visible = False
                'objMatrix.Columns.Item("U_ItemRef").Visible = False
                'objMatrix.Columns.Item("U_Barcode").Visible = False
                'objMatrix.Columns.Item("U_Uomqty").Visible = False
                'objMatrix.Columns.Item("U_SplitKitItems").Visible = False
                'objMatrix.Columns.Item("U_LineValid").Visible = False
                'objMatrix.Columns.Item("U_BelowCst").Visible = False
            Catch ex As Exception
            End Try
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub CalculateFOCbackup(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("38").Specific
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
            For i As Integer = 1 To objMatrix.VisualRowCount - 1
                Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.ToString().Trim()
                Dim FOCQty As Double = objMatrix.Columns.Item("U_FOCQTY").Cells.Item(i).Specific.Value.ToString().Trim()
                Dim NormalQty As Double = objMatrix.Columns.Item("U_normalqty").Cells.Item(i).Specific.Value.ToString().Trim()
                If (FOCQty + NormalQty) <> 0 Then
                    objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value = FOCQty + NormalQty
                End If
                Dim PurchasePrice As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(i).Specific.Value)
                If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "-" Then
                    Dim FOCTotal As Double = 0
                    Dim Total As Double = PurchasePrice * CDbl(NormalQty)
                    Dim PriceAfterDiscount As Double = Total / (NormalQty + FOCQty)
                    Dim Discount As Double = ((PurchasePrice - PriceAfterDiscount) / 100) * 100
                    If FOCQty > 0 Then
                        Try
                            Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                            oColumn.Cells.Item(i).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        Catch ex As Exception
                        End Try
                    End If
                    If PriceAfterDiscount >= 0 Then
                        objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value = PriceAfterDiscount
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
                            Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                            If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "A" Then
                                Dim FOCQtyA As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                                Dim NormalQtyA As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                                Dim PurchasePriceA As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                                Dim Price As Double = PurchasePriceA - ((Ratio / 100) * PurchasePriceA)
                                If Price > 0 Then
                                    objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price
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
                            Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                            If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "B" Then
                                Dim FOCQtyB As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                                Dim NormalQtyB As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                                Dim PurchasePriceB As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
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
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "C" Then
                            Dim FOCQtyC As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyC As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceC As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
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
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "D" Then
                            Dim FOCQtyD As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyD As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceD As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
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
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "E" Then
                            Dim FOCQtyE As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyE As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceE As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
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
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "F" Then
                            Dim FOCQtyF As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyF As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceF As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
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
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "G" Then
                            Dim FOCQtyG As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyG As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceG As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
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
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "H" Then
                            Dim FOCQtyH As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyH As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceH As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
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
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "I" Then

                            Dim FOCQtyI As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyI As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceI As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
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
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "J" Then
                            Dim FOCQtyJ As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyJ As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceJ As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                            Dim Price As Double = PurchasePriceJ - ((Ratio / 100) * PurchasePriceJ)
                            objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price
                        End If
                    Next
                    FlagJ = "Completed"

                End If

            Next
            'madhu new logic start here


        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub CalculateFOC(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("38").Specific
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
            objForm.Freeze(True)
            For i As Integer = 1 To objMatrix.VisualRowCount - 1
                Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.ToString().Trim()
                Dim FOCQty As Double = objMatrix.Columns.Item("U_FOCQTY").Cells.Item(i).Specific.Value.ToString().Trim()
                Dim NormalQty As Double = objMatrix.Columns.Item("U_normalqty").Cells.Item(i).Specific.Value.ToString().Trim()
                If (FOCQty + NormalQty) <> 0 Then
                    objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value = FOCQty + NormalQty
                End If
                Dim PurchasePrice As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(i).Specific.Value)
                If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "-" Then
                    Dim FOCTotal As Double = 0
                    Dim Total As Double = PurchasePrice * CDbl(NormalQty)
                    Dim PriceAfterDiscount As Double = Total / (NormalQty + FOCQty)
                    Dim Discount As Double = ((PurchasePrice - PriceAfterDiscount) / 100) * 100
                    If FOCQty > 0 Then
                        Try
                            Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                            oColumn.Cells.Item(i).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        Catch ex As Exception
                        End Try
                    End If
                    If PriceAfterDiscount >= 0 Then
                        objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value = Math.Round(PriceAfterDiscount, 6) 'PriceAfterDiscount
                        ''oDocument.Lines.PriceAfterDiscount = Math.Round(PriceAfterDiscount, 6)
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
                            Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                            If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "A" Then
                                Dim FOCQtyA As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                                Dim NormalQtyA As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                                Dim PurchasePriceA As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                                Dim Price As Double = PurchasePriceA - ((Ratio / 100) * PurchasePriceA)
                                If Price > 0 Then
                                    objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price
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
                            Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                            If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "B" Then
                                Dim FOCQtyB As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                                Dim NormalQtyB As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                                Dim PurchasePriceB As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
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
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "C" Then
                            Dim FOCQtyC As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyC As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceC As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
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
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "D" Then
                            Dim FOCQtyD As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyD As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceD As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
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
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "E" Then
                            Dim FOCQtyE As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyE As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceE As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
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
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "F" Then
                            Dim FOCQtyF As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyF As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceF As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
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
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "G" Then
                            Dim FOCQtyG As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyG As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceG As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
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
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "H" Then
                            Dim FOCQtyH As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyH As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceH As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
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
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "I" Then

                            Dim FOCQtyI As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyI As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceI As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
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
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "J" Then
                            Dim FOCQtyJ As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim NormalQtyJ As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                            Dim PurchasePriceJ As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                            Dim Price As Double = PurchasePriceJ - ((Ratio / 100) * PurchasePriceJ)
                            objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price
                        End If
                    Next
                    FlagJ = "Completed"

                End If



                'madhu new logic for line total, vat amount and gross total calculation

                Dim SupplierPrice As Double = Val(objMatrix.Columns.Item("U_suppprice").Cells.Item(i).Specific.Value)
                Dim DiscPrcnt As Double = Val(objMatrix.Columns.Item("15").Cells.Item(i).Specific.Value)

                Dim NormalQuantity As Double = Val(objMatrix.Columns.Item("U_normalqty").Cells.Item(i).Specific.Value)
                Dim ExTaxofbase As Double = Val(objMatrix.Columns.Item("U_EXT").Cells.Item(i).Specific.Value)
                Dim FocQuantity As Double = Val(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(i).Specific.Value)

                Dim FinalLineTotal As Double = 0


                ' If NormalQuantity > 0 And SupplierPrice > 0 And FocQuantity = 0 And ExTaxofbase = 0 Then

                'FinalLineTotal = NormalQuantity * SupplierPrice * (1 - DiscPrcnt / 100)

                'ElseIf ExTaxofbase > 0 Or FocQuantity > 0 Then

                FinalLineTotal = ((NormalQuantity * SupplierPrice) + (ExTaxofbase * FocQuantity)) * (1 - DiscPrcnt / 100)
                ' End If


                objMatrix.Columns.Item("U_LTOTAL").Cells.Item(i).Specific.Value = FinalLineTotal

                Dim VatCode As String = ""
                Dim VatPercent As Double = 0

                Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                Dim Vatgroup As String = objMatrix.Columns.Item("18").Cells.Item(i).Specific.Value.ToString()


                If Not String.IsNullOrEmpty(Vatgroup) Then
                    oRS.DoQuery("SELECT ""Rate"" FROM OVTG WHERE ""Code"" = '" & Vatgroup & "'")
                    If Not oRS.EoF Then
                        VatPercent = CDbl(oRS.Fields.Item("Rate").Value)
                    End If
                End If


                Dim VatAmount As Double = FinalLineTotal * VatPercent / 100
                objMatrix.Columns.Item("U_VATAMT").Cells.Item(i).Specific.Value = VatAmount

                Dim GrossTotal As Double = FinalLineTotal + VatAmount
                objMatrix.Columns.Item("U_GTOTAL").Cells.Item(i).Specific.Value = GrossTotal

            Next


            Dim SupplierPOTotal As Double = 0

            For i As Integer = 1 To objMatrix.VisualRowCount - 1
                Dim GT As Double = Val(objMatrix.Columns.Item("U_GTOTAL").Cells.Item(i).Specific.Value)
                SupplierPOTotal += GT
            Next

            '---Client wants to now not required comment
            'Try
            '    Dim oEdit As SAPbouiCOM.EditText = objForm.Items.Item("U_SPOTOTAL").Specific
            '    'oEdit.Value = SupplierPOTotal
            '    oEdit.Value = Math.Round(SupplierPOTotal, 3, MidpointRounding.AwayFromZero)
            'Catch ex As Exception
            '    objForm.DataSources.DBDataSources.Item("OPOR").SetValue("U_SPOTOTAL", 0, SupplierPOTotal)
            'End Try

            'Round value
            '---Client wants to now not required comment
            'Try
            '    Dim oForm As SAPbouiCOM.Form = objForm
            '    Dim DocTotal As Double = CDbl(oForm.DataSources.DBDataSources.Item("OPOR").GetValue("DocTotal", 0))
            '    Dim RoundingValue As Double = Math.Round(SupplierPOTotal - DocTotal, 3, MidpointRounding.AwayFromZero)
            '    Dim objCheckCh As SAPbouiCOM.CheckBox = oForm.Items.Item("105").Specific
            '    If objCheckCh.Checked = False Then
            '        oForm.Items.Item("105").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            '    End If
            '    oForm.Items.Item("103").Specific.Value = RoundingValue

            'Catch ex As Exception
            '    objMain.objApplication.StatusBar.SetText("RoundDiff Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'End Try


            objForm.Freeze(False)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Private Sub CalculateFOC_SUAGARTAX(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("38").Specific
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
            For i As Integer = 1 To objMatrix.VisualRowCount - 1
                Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.ToString().Trim()
                Dim FOCQty As Double = objMatrix.Columns.Item("U_FOCQTY").Cells.Item(i).Specific.Value.ToString().Trim()
                Dim NormalQty As Double = objMatrix.Columns.Item("U_normalqty").Cells.Item(i).Specific.Value.ToString().Trim()
                If (FOCQty + NormalQty) <> 0 Then
                    objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value = FOCQty + NormalQty
                End If
                Dim PurchasePrice As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(i).Specific.Value)
                If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "-" Then
                    Dim FOCTotal As Double = 0
                    Dim Total As Double = PurchasePrice * CDbl(NormalQty)
                    Dim PriceAfterDiscount As Double = Total / (NormalQty + FOCQty)
                    Dim Discount As Double = ((PurchasePrice - PriceAfterDiscount) / 100) * 100
                    If FOCQty > 0 Then
                        Try
                            Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                            oColumn.Cells.Item(i).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        Catch ex As Exception
                        End Try
                    End If
                    If PriceAfterDiscount >= 0 Then
                        objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value = PriceAfterDiscount
                    End If
                ElseIf objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "A" And FlagA = "" Then
                    'Dim NormalTotal As Double = 0
                    'Dim FOCTotal As Double = 0
                    'Dim TotalQty As Double = 0
                    'For j As Integer = 1 To objMatrix.VisualRowCount - 1
                    '    If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "A" Then
                    '        Dim FOCQtyA As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                    '        Dim NormalQtyA As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                    '        If FOCQtyA > 0 Then
                    '            Try
                    '                Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                    '                oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    '            Catch ex As Exception
                    '            End Try
                    '        End If
                    '        If (FOCQtyA + NormalQtyA) <> 0 Then
                    '            objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyA + NormalQtyA
                    '        End If
                    '        Dim PurchasePriceA As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)

                    '        ' ✅ Use Base After Discount (2.134)
                    '        Dim BaseAfterDisc As Double = CDbl(objMatrix.Columns.Item("U_BAD").Cells.Item(j).Specific.Value)

                    '        FOCTotal = FOCTotal + (FOCQtyA * BaseAfterDisc)
                    '        NormalTotal = NormalTotal + (NormalQtyA * PurchasePriceA)
                    '        'FOCTotal = FOCTotal + (FOCQtyA * PurchasePriceA)
                    '        'NormalTotal = NormalTotal + (NormalQtyA * PurchasePriceA)
                    '        If FOCQtyA > 0 Or NormalQtyA > 0 Then
                    '            TotalQty = TotalQty + FOCQtyA + NormalQtyA
                    '        Else
                    '            objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greate then 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '            Exit Sub
                    '        End If
                    '    End If
                    'Next
                    'For j As Integer = 1 To objMatrix.VisualRowCount - 1
                    '    If TotalQty > 0 Then
                    '        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                    '        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "A" Then
                    '            Dim FOCQtyA As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                    '            Dim NormalQtyA As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                    '            Dim PurchasePriceA As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                    '            Dim Price As Double = PurchasePriceA - ((Ratio / 100) * PurchasePriceA)
                    '            Dim PerUnitFOCDiscount As Double = FOCTotal / TotalQty
                    '            Dim Price1 As Double = PurchasePriceA - PerUnitFOCDiscount
                    '            If Price > 0 Then
                    '                objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price
                    '            End If
                    '        End If
                    '    End If
                    'Next
                    'FlagA = "Completed"
                    Dim NormalTotal As Double = 0
                    Dim FOCTotal As Double = 0
                    Dim TotalQty As Double = 0
                    Dim TotalFOCQty As Double = 0

                    ' 🔹 FIRST LOOP: Get TotalQty + TotalFOCQty
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1

                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "A" Then

                            ' SAFE FOC QTY
                            Dim FOCQtyA As Double = 0
                            Dim focVal As String = objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim()
                            If focVal <> "" Then
                                FOCQtyA = CDbl(focVal)
                            End If
                            If FOCQtyA > 0 Then
                                Try
                                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                                    oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Catch ex As Exception
                                End Try
                            End If
                            ' SAFE NORMAL QTY
                            Dim NormalQtyA As Double = 0
                            Dim normVal As String = objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim()
                            If normVal <> "" Then
                                NormalQtyA = CDbl(normVal)
                            End If

                            ' ✅ Sum TOTAL FOC QTY
                            TotalFOCQty = TotalFOCQty + FOCQtyA

                            ' Total Qty
                            If (FOCQtyA + NormalQtyA) <> 0 Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyA + NormalQtyA
                                TotalQty = TotalQty + FOCQtyA + NormalQtyA
                            Else
                                objMain.objApplication.StatusBar.SetText("Qty should be > 0 in line:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If

                            Dim PurchasePriceA As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)

                            ' Normal total
                            NormalTotal = NormalTotal + (NormalQtyA * PurchasePriceA)

                        End If

                    Next

                    ' 🔹 STEP 2: Calculate FOCTotal using TOTAL FOC QTY
                    ' ⚠️ MUST be Base After Discount (2.134)
                    'Dim BaseAfterDisc As Double = CDbl(objMatrix.Columns.Item("U_BAD").Cells.Item(1).Specific.Value)
                    'Dim SupplierPC As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(1).Specific.Value)

                    'FOCTotal = TotalFOCQty * BaseAfterDisc

                    ' 🔹 STEP 3: Apply price
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1

                        Dim BaseAfterDisc As Double = CDbl(objMatrix.Columns.Item("U_BAD").Cells.Item(j).Specific.Value)
                        Dim SupplierPC As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)

                        FOCTotal = TotalFOCQty * BaseAfterDisc

                        If TotalQty > 0 Then

                            If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "A" Then

                                Dim PurchasePriceA As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)

                                ' Per unit discount
                                Dim PerUnitFOCDiscount As Double = FOCTotal / TotalQty

                                ' Rounded price
                                Dim Price As Double = (PurchasePriceA - PerUnitFOCDiscount)

                                If Price > 0 Then
                                    objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price
                                Else
                                    objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = 0
                                End If

                            End If

                        End If

                    Next

                    FlagA = "Completed"

                ElseIf objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "B" And FlagB = "" Then
                    'Dim NormalTotal As Double = 0
                    'Dim FOCTotal As Double = 0
                    'Dim TotalQty As Double = 0
                    'For j As Integer = 1 To objMatrix.VisualRowCount - 1
                    '    If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "B" Then
                    '        Dim FOCQtyB As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                    '        Dim NormalQtyB As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                    '        If FOCQtyB > 0 Then
                    '            Try
                    '                Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                    '                oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    '            Catch ex As Exception
                    '            End Try
                    '        End If
                    '        If (FOCQtyB + NormalQtyB) <> 0 Then
                    '            objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyB + NormalQtyB
                    '        End If
                    '        Dim PurchasePriceB As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                    '        FOCTotal = FOCTotal + (FOCQtyB * PurchasePriceB)
                    '        NormalTotal = NormalTotal + (NormalQtyB * PurchasePriceB)
                    '        If FOCQtyB > 0 Or NormalQtyB > 0 Then
                    '            TotalQty = TotalQty + FOCQtyB + NormalQtyB
                    '        Else
                    '            objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greate then 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '            Exit Sub
                    '        End If
                    '    End If
                    'Next
                    'For j As Integer = 1 To objMatrix.VisualRowCount - 1
                    '    If TotalQty > 0 Then
                    '        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                    '        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "B" Then
                    '            Dim FOCQtyB As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                    '            Dim NormalQtyB As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                    '            Dim PurchasePriceB As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                    '            Dim Price As Double = PurchasePriceB - ((Ratio / 100) * PurchasePriceB)
                    '            objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price
                    '        End If
                    '    End If
                    'Next
                    'FlagB = "Completed"
                    Dim NormalTotal As Double = 0
                    Dim FOCTotal As Double = 0
                    Dim TotalQty As Double = 0
                    Dim TotalFOCQty As Double = 0

                    ' 🔹 FIRST LOOP: TotalQty + TotalFOCQty
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1

                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "B" Then

                            ' ✅ SAFE FOC QTY
                            Dim FOCQtyB As Double = 0
                            Dim focVal As String = objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim()
                            If focVal <> "" Then
                                FOCQtyB = CDbl(focVal)
                            End If

                            ' ✅ SAFE NORMAL QTY
                            Dim NormalQtyB As Double = 0
                            Dim normVal As String = objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim()
                            If normVal <> "" Then
                                NormalQtyB = CDbl(normVal)
                            End If

                            ' Mark valid line if FOC exists
                            If FOCQtyB > 0 Then
                                Try
                                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                                    oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Catch ex As Exception
                                End Try
                            End If

                            ' Total Qty per row
                            If (FOCQtyB + NormalQtyB) <> 0 Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyB + NormalQtyB
                                TotalQty = TotalQty + FOCQtyB + NormalQtyB
                            Else
                                objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greater than 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If

                            ' ✅ SUM TOTAL FOC QTY
                            TotalFOCQty = TotalFOCQty + FOCQtyB

                            Dim PurchasePriceB As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)

                            ' Normal total (kept)
                            NormalTotal = NormalTotal + (NormalQtyB * PurchasePriceB)

                        End If

                    Next

                    ' 🔹 STEP 2: Calculate FOCTotal using TOTAL FOC QTY
                    ' ⚠️ MUST be Base After Discount (2.134)
                    'Dim BaseAfterDisc As Double = CDbl(objMatrix.Columns.Item("U_BAD").Cells.Item(1).Specific.Value)

                    'FOCTotal = TotalFOCQty * BaseAfterDisc

                    ' 🔹 SECOND LOOP: Apply price
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1

                        If TotalQty > 0 Then

                            Dim BaseAfterDisc As Double = CDbl(objMatrix.Columns.Item("U_BAD").Cells.Item(j).Specific.Value)

                            FOCTotal = TotalFOCQty * BaseAfterDisc

                            ' (kept for structure, not used)
                            Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)

                            If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "B" Then

                                Dim PurchasePriceB As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)

                                ' ✅ Flat per-unit discount
                                Dim PerUnitFOCDiscount As Double = FOCTotal / TotalQty

                                ' ✅ Rounded price
                                'madhu comment due removing round value
                                ' Dim Price As Double = Math.Round(PurchasePriceB - PerUnitFOCDiscount, 3, MidpointRounding.AwayFromZero)
                                Dim Price As Double = (PurchasePriceB - PerUnitFOCDiscount)

                                objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price

                            End If

                        End If

                    Next

                    FlagB = "Completed"
                ElseIf objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "C" And FlagC = "" Then
                    '    Dim NormalTotal As Double = 0
                    '    Dim FOCTotal As Double = 0
                    '    Dim TotalQty As Double = 0
                    '    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                    '        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "C" Then
                    '            Dim FOCQtyC As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                    '            Dim NormalQtyC As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                    '            If FOCQtyC > 0 Then
                    '                Try
                    '                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                    '                    oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    '                Catch ex As Exception
                    '                End Try
                    '            End If
                    '            If (FOCQtyC + NormalQtyC) <> 0 Then
                    '                objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyC + NormalQtyC
                    '            End If
                    '            Dim PurchasePriceC As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                    '            FOCTotal = FOCTotal + (FOCQtyC * PurchasePriceC)
                    '            NormalTotal = NormalTotal + (NormalQtyC * PurchasePriceC)
                    '            If FOCQtyC > 0 Or NormalQtyC > 0 Then
                    '                TotalQty = TotalQty + FOCQtyC + NormalQtyC
                    '            Else
                    '                objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greate then 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '                Exit Sub
                    '            End If
                    '        End If
                    '    Next
                    '    If TotalQty <= 0 Then
                    '        objMain.objApplication.StatusBar.SetText("All Items total quantity is 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        Exit Sub
                    '    End If
                    '    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                    '        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                    '        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "C" Then
                    '            Dim FOCQtyC As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                    '            Dim NormalQtyC As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                    '            Dim PurchasePriceC As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                    '            Dim Price As Double = PurchasePriceC - ((Ratio / 100) * PurchasePriceC)
                    '            objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price
                    '        End If
                    '    Next
                    '    FlagC = "Completed"
                    Dim NormalTotal As Double = 0
                    Dim FOCTotal As Double = 0
                    Dim TotalQty As Double = 0
                    Dim TotalFOCQty As Double = 0

                    ' 🔹 FIRST LOOP: TotalQty + TotalFOCQty
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1

                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "C" Then

                            ' ✅ SAFE FOC QTY
                            Dim FOCQtyC As Double = 0
                            Dim focVal As String = objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim()
                            If focVal <> "" Then
                                FOCQtyC = CDbl(focVal)
                            End If

                            ' ✅ SAFE NORMAL QTY
                            Dim NormalQtyC As Double = 0
                            Dim normVal As String = objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim()
                            If normVal <> "" Then
                                NormalQtyC = CDbl(normVal)
                            End If

                            ' Mark valid line if FOC exists
                            If FOCQtyC > 0 Then
                                Try
                                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                                    oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Catch ex As Exception
                                End Try
                            End If

                            ' Total Qty per row
                            If (FOCQtyC + NormalQtyC) <> 0 Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyC + NormalQtyC
                                TotalQty = TotalQty + FOCQtyC + NormalQtyC
                            Else
                                objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greater than 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If

                            ' ✅ TOTAL FOC QTY
                            TotalFOCQty = TotalFOCQty + FOCQtyC

                            Dim PurchasePriceC As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)

                            ' Normal total (kept)
                            NormalTotal = NormalTotal + (NormalQtyC * PurchasePriceC)

                        End If

                    Next

                    ' 🔹 VALIDATION
                    If TotalQty <= 0 Then
                        objMain.objApplication.StatusBar.SetText("All Items total quantity is 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If

                    ' 🔹 STEP 2: Calculate FOCTotal using TOTAL FOC QTY
                    ' ⚠️ MUST be Base After Discount (2.134)
                    'Dim BaseAfterDisc As Double = CDbl(objMatrix.Columns.Item("U_BAD").Cells.Item(1).Specific.Value)

                    'FOCTotal = TotalFOCQty * BaseAfterDisc

                    ' 🔹 SECOND LOOP: Apply price
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1

                        Dim BaseAfterDisc As Double = CDbl(objMatrix.Columns.Item("U_BAD").Cells.Item(j).Specific.Value)

                        FOCTotal = TotalFOCQty * BaseAfterDisc

                        ' (kept for structure, not used)
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)

                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "C" Then

                            Dim PurchasePriceC As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)

                            ' ✅ Flat per-unit discount
                            Dim PerUnitFOCDiscount As Double = FOCTotal / TotalQty

                            ' ✅ Rounded price (3 decimals)
                            Dim Price As Double = (PurchasePriceC - PerUnitFOCDiscount)

                            objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price

                        End If

                    Next

                    FlagC = "Completed"
                ElseIf objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "D" And FlagD = "" Then
                    Dim NormalTotal As Double = 0
                    Dim FOCTotal As Double = 0
                    Dim TotalQty As Double = 0
                    Dim TotalFOCQty As Double = 0

                    ' 🔹 FIRST LOOP: TotalQty + TotalFOCQty
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1

                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "D" Then

                            ' ✅ SAFE FOC QTY
                            Dim FOCQtyD As Double = 0
                            Dim focVal As String = objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim()
                            If focVal <> "" Then
                                FOCQtyD = CDbl(focVal)
                            End If

                            ' ✅ SAFE NORMAL QTY
                            Dim NormalQtyD As Double = 0
                            Dim normVal As String = objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim()
                            If normVal <> "" Then
                                NormalQtyD = CDbl(normVal)
                            End If

                            ' Mark valid line if FOC exists
                            If FOCQtyD > 0 Then
                                Try
                                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                                    oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Catch ex As Exception
                                End Try
                            End If

                            ' Total Qty per row
                            If (FOCQtyD + NormalQtyD) <> 0 Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyD + NormalQtyD
                                TotalQty = TotalQty + FOCQtyD + NormalQtyD
                            Else
                                objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greater than 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If

                            ' ✅ TOTAL FOC QTY
                            TotalFOCQty = TotalFOCQty + FOCQtyD

                            Dim PurchasePriceD As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)

                            ' Normal total (kept for structure)
                            NormalTotal = NormalTotal + (NormalQtyD * PurchasePriceD)

                        End If

                    Next

                    ' 🔹 VALIDATION
                    If TotalQty <= 0 Then
                        objMain.objApplication.StatusBar.SetText("All Items total quantity is 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If

                    ' 🔹 STEP 2: Calculate FOCTotal using TOTAL FOC QTY
                    ' ⚠️ MUST be Base After Discount (2.134)
                    'Dim BaseAfterDisc As Double = CDbl(objMatrix.Columns.Item("U_BAD").Cells.Item(1).Specific.Value)

                    'FOCTotal = TotalFOCQty * BaseAfterDisc

                    ' 🔹 SECOND LOOP: Apply price
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1

                        Dim BaseAfterDisc As Double = CDbl(objMatrix.Columns.Item("U_BAD").Cells.Item(j).Specific.Value)

                        FOCTotal = TotalFOCQty * BaseAfterDisc
                        ' (kept for structure, not used)
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)

                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "D" Then

                            Dim PurchasePriceD As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)

                            ' ✅ Flat per-unit discount
                            Dim PerUnitFOCDiscount As Double = FOCTotal / TotalQty

                            ' ✅ Rounded price (3 decimals)
                            Dim Price As Double = (PurchasePriceD - PerUnitFOCDiscount)

                            objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price

                        End If

                    Next

                    FlagD = "Completed"
                ElseIf objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "E" And FlagE = "" Then
                    'Dim NormalTotal As Double = 0
                    'Dim FOCTotal As Double = 0
                    'Dim TotalQty As Double = 0
                    'For j As Integer = 1 To objMatrix.VisualRowCount - 1
                    '    If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "E" Then

                    '        Dim FOCQtyE As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                    '        Dim NormalQtyE As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                    '        If FOCQtyE > 0 Then
                    '            Try
                    '                Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                    '                oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    '            Catch ex As Exception
                    '            End Try
                    '        End If
                    '        If (FOCQtyE + NormalQtyE) <> 0 Then
                    '            objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyE + NormalQtyE
                    '        End If
                    '        Dim PurchasePriceE As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                    '        FOCTotal = FOCTotal + (FOCQtyE * PurchasePriceE)
                    '        NormalTotal = NormalTotal + (NormalQtyE * PurchasePriceE)
                    '        If FOCQtyE > 0 Or NormalQtyE > 0 Then
                    '            TotalQty = TotalQty + FOCQtyE + NormalQtyE
                    '        Else
                    '            objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greate then 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '            Exit Sub
                    '        End If
                    '    End If
                    'Next
                    'If TotalQty <= 0 Then
                    '    objMain.objApplication.StatusBar.SetText("All Items total quantity is 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    Exit Sub
                    'End If
                    'For j As Integer = 1 To objMatrix.VisualRowCount - 1
                    '    Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                    '    If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "E" Then
                    '        Dim FOCQtyE As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                    '        Dim NormalQtyE As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                    '        Dim PurchasePriceE As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                    '        Dim Price As Double = PurchasePriceE - ((Ratio / 100) * PurchasePriceE)
                    '        objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price
                    '    End If
                    'Next
                    'FlagE = "Completed"
                    Dim NormalTotal As Double = 0
                    Dim FOCTotal As Double = 0
                    Dim TotalQty As Double = 0
                    Dim TotalFOCQty As Double = 0

                    ' 🔹 FIRST LOOP: TotalQty + TotalFOCQty
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1

                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "E" Then

                            ' ✅ SAFE FOC QTY
                            Dim FOCQtyE As Double = 0
                            Dim focVal As String = objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim()
                            If focVal <> "" Then
                                FOCQtyE = CDbl(focVal)
                            End If

                            ' ✅ SAFE NORMAL QTY
                            Dim NormalQtyE As Double = 0
                            Dim normVal As String = objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim()
                            If normVal <> "" Then
                                NormalQtyE = CDbl(normVal)
                            End If

                            ' Mark valid line if FOC exists
                            If FOCQtyE > 0 Then
                                Try
                                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                                    oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Catch ex As Exception
                                End Try
                            End If

                            ' Total Qty per row
                            If (FOCQtyE + NormalQtyE) <> 0 Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyE + NormalQtyE
                                TotalQty = TotalQty + FOCQtyE + NormalQtyE
                            Else
                                objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greater than 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If

                            ' ✅ TOTAL FOC QTY
                            TotalFOCQty = TotalFOCQty + FOCQtyE

                            Dim PurchasePriceE As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)

                            ' Normal total (kept for structure)
                            NormalTotal = NormalTotal + (NormalQtyE * PurchasePriceE)

                        End If

                    Next

                    ' 🔹 VALIDATION
                    If TotalQty <= 0 Then
                        objMain.objApplication.StatusBar.SetText("All Items total quantity is 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If

                    ' 🔹 STEP 2: Calculate FOCTotal using TOTAL FOC QTY
                    ' ⚠️ MUST be Base After Discount (2.134)
                    'Dim BaseAfterDisc As Double = CDbl(objMatrix.Columns.Item("U_BAD").Cells.Item(1).Specific.Value)

                    'FOCTotal = TotalFOCQty * BaseAfterDisc

                    ' 🔹 SECOND LOOP: Apply price
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1

                        Dim BaseAfterDisc As Double = CDbl(objMatrix.Columns.Item("U_BAD").Cells.Item(j).Specific.Value)

                        FOCTotal = TotalFOCQty * BaseAfterDisc

                        ' (kept for structure, not used)
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)

                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "E" Then

                            Dim PurchasePriceE As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)

                            ' ✅ Flat per-unit discount
                            Dim PerUnitFOCDiscount As Double = FOCTotal / TotalQty

                            ' ✅ Rounded price (3 decimals)
                            Dim Price As Double = (PurchasePriceE - PerUnitFOCDiscount)

                            objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price

                        End If

                    Next

                    FlagE = "Completed"
                ElseIf objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "F" And FlagF = "" Then
                    'Dim NormalTotal As Double = 0
                    'Dim FOCTotal As Double = 0
                    'Dim TotalQty As Double = 0
                    'For j As Integer = 1 To objMatrix.VisualRowCount - 1
                    '    If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "F" Then
                    '        Dim FOCQtyF As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                    '        Dim NormalQtyF As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                    '        If FOCQtyF > 0 Then
                    '            Try
                    '                Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                    '                oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    '            Catch ex As Exception
                    '            End Try
                    '        End If
                    '        If (FOCQtyF + NormalQtyF) <> 0 Then
                    '            objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyF + NormalQtyF
                    '        End If
                    '        Dim PurchasePriceF As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                    '        FOCTotal = FOCTotal + (FOCQtyF * PurchasePriceF)
                    '        NormalTotal = NormalTotal + (NormalQtyF * PurchasePriceF)
                    '        If FOCQtyF > 0 Or NormalQtyF > 0 Then
                    '            TotalQty = TotalQty + FOCQtyF + NormalQtyF
                    '        Else
                    '            objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greate then 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '            Exit Sub
                    '        End If
                    '    End If
                    'Next
                    'If TotalQty <= 0 Then
                    '    objMain.objApplication.StatusBar.SetText("All Items total quantity is 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    Exit Sub
                    'End If
                    'For j As Integer = 1 To objMatrix.VisualRowCount - 1
                    '    Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                    '    If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "F" Then
                    '        Dim FOCQtyF As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                    '        Dim NormalQtyF As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                    '        Dim PurchasePriceF As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                    '        Dim Price As Double = PurchasePriceF - ((Ratio / 100) * PurchasePriceF)
                    '        objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price
                    '    End If
                    'Next
                    'FlagF = "Completed"
                    Dim NormalTotal As Double = 0
                    Dim FOCTotal As Double = 0
                    Dim TotalQty As Double = 0
                    Dim TotalFOCQty As Double = 0

                    ' 🔹 FIRST LOOP: TotalQty + TotalFOCQty
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1

                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "F" Then

                            ' ✅ SAFE FOC QTY
                            Dim FOCQtyF As Double = 0
                            Dim focVal As String = objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim()
                            If focVal <> "" Then
                                FOCQtyF = CDbl(focVal)
                            End If

                            ' ✅ SAFE NORMAL QTY
                            Dim NormalQtyF As Double = 0
                            Dim normVal As String = objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim()
                            If normVal <> "" Then
                                NormalQtyF = CDbl(normVal)
                            End If

                            ' Mark valid line if FOC exists
                            If FOCQtyF > 0 Then
                                Try
                                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                                    oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Catch ex As Exception
                                End Try
                            End If

                            ' Total Qty per row
                            If (FOCQtyF + NormalQtyF) <> 0 Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyF + NormalQtyF
                                TotalQty = TotalQty + FOCQtyF + NormalQtyF
                            Else
                                objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greater than 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If

                            ' ✅ TOTAL FOC QTY
                            TotalFOCQty = TotalFOCQty + FOCQtyF

                            Dim PurchasePriceF As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)

                            ' Normal total (kept for structure)
                            NormalTotal = NormalTotal + (NormalQtyF * PurchasePriceF)

                        End If

                    Next

                    ' 🔹 VALIDATION
                    If TotalQty <= 0 Then
                        objMain.objApplication.StatusBar.SetText("All Items total quantity is 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If

                    ' 🔹 STEP 2: Calculate FOCTotal using TOTAL FOC QTY
                    ' ⚠️ MUST be Base After Discount (2.134)
                    'Dim BaseAfterDisc As Double = CDbl(objMatrix.Columns.Item("U_BAD").Cells.Item(1).Specific.Value)

                    'FOCTotal = TotalFOCQty * BaseAfterDisc

                    ' 🔹 SECOND LOOP: Apply price
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1

                        Dim BaseAfterDisc As Double = CDbl(objMatrix.Columns.Item("U_BAD").Cells.Item(j).Specific.Value)

                        FOCTotal = TotalFOCQty * BaseAfterDisc

                        ' (kept for structure, not used)
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)

                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "F" Then

                            Dim PurchasePriceF As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)

                            ' ✅ Flat per-unit discount
                            Dim PerUnitFOCDiscount As Double = FOCTotal / TotalQty

                            ' ✅ Rounded price (3 decimals)
                            Dim Price As Double = (PurchasePriceF - PerUnitFOCDiscount)

                            objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price

                        End If

                    Next

                    FlagF = "Completed"
                ElseIf objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "G" And FlagG = "" Then
                    'Dim NormalTotal As Double = 0
                    'Dim FOCTotal As Double = 0
                    'Dim TotalQty As Double = 0
                    'For j As Integer = 1 To objMatrix.VisualRowCount - 1
                    '    If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "G" Then
                    '        Dim FOCQtyG As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                    '        Dim NormalQtyG As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                    '        If FOCQtyG > 0 Then
                    '            Try
                    '                Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                    '                oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    '            Catch ex As Exception
                    '            End Try
                    '        End If

                    '        If (FOCQtyG + NormalQtyG) <> 0 Then
                    '            objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyG + NormalQtyG
                    '        End If
                    '        Dim PurchasePriceG As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                    '        FOCTotal = FOCTotal + (FOCQtyG * PurchasePriceG)
                    '        NormalTotal = NormalTotal + (NormalQtyG * PurchasePriceG)
                    '        If FOCQtyG > 0 Or NormalQtyG > 0 Then
                    '            TotalQty = TotalQty + FOCQtyG + NormalQtyG
                    '        Else
                    '            objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greate then 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '            Exit Sub
                    '        End If
                    '    End If
                    'Next
                    'If TotalQty <= 0 Then
                    '    objMain.objApplication.StatusBar.SetText("All Items total quantity is 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    Exit Sub
                    'End If
                    'For j As Integer = 1 To objMatrix.VisualRowCount - 1
                    '    Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                    '    If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "G" Then
                    '        Dim FOCQtyG As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                    '        Dim NormalQtyG As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                    '        Dim PurchasePriceG As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                    '        Dim Price As Double = PurchasePriceG - ((Ratio / 100) * PurchasePriceG)
                    '        objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price
                    '    End If
                    'Next
                    'FlagG = "Completed"
                    Dim NormalTotal As Double = 0
                    Dim FOCTotal As Double = 0
                    Dim TotalQty As Double = 0
                    Dim TotalFOCQty As Double = 0

                    ' 🔹 FIRST LOOP: TotalQty + TotalFOCQty
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1

                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "G" Then

                            ' ✅ SAFE FOC QTY
                            Dim FOCQtyG As Double = 0
                            Dim focVal As String = objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim()
                            If focVal <> "" Then
                                FOCQtyG = CDbl(focVal)
                            End If

                            ' ✅ SAFE NORMAL QTY
                            Dim NormalQtyG As Double = 0
                            Dim normVal As String = objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim()
                            If normVal <> "" Then
                                NormalQtyG = CDbl(normVal)
                            End If

                            ' Mark valid line if FOC exists
                            If FOCQtyG > 0 Then
                                Try
                                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                                    oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Catch ex As Exception
                                End Try
                            End If

                            ' Total Qty per row
                            If (FOCQtyG + NormalQtyG) <> 0 Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyG + NormalQtyG
                                TotalQty = TotalQty + FOCQtyG + NormalQtyG
                            Else
                                objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greater than 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If

                            ' ✅ TOTAL FOC QTY
                            TotalFOCQty = TotalFOCQty + FOCQtyG

                            Dim PurchasePriceG As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)

                            ' Normal total (kept for structure)
                            NormalTotal = NormalTotal + (NormalQtyG * PurchasePriceG)

                        End If

                    Next

                    ' 🔹 VALIDATION
                    If TotalQty <= 0 Then
                        objMain.objApplication.StatusBar.SetText("All Items total quantity is 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If

                    ' 🔹 STEP 2: Calculate FOCTotal using TOTAL FOC QTY
                    ' ⚠️ MUST be Base After Discount (2.134)
                    'Dim BaseAfterDisc As Double = CDbl(objMatrix.Columns.Item("U_BAD").Cells.Item(1).Specific.Value)

                    'FOCTotal = TotalFOCQty * BaseAfterDisc

                    ' 🔹 SECOND LOOP: Apply price
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1

                        Dim BaseAfterDisc As Double = CDbl(objMatrix.Columns.Item("U_BAD").Cells.Item(j).Specific.Value)

                        FOCTotal = TotalFOCQty * BaseAfterDisc
                        ' (kept for structure, not used)
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)

                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "G" Then

                            Dim PurchasePriceG As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)

                            ' ✅ Flat per-unit discount
                            Dim PerUnitFOCDiscount As Double = FOCTotal / TotalQty

                            ' ✅ Rounded price (3 decimals)
                            Dim Price As Double = Math.Round(PurchasePriceG - PerUnitFOCDiscount, 3, MidpointRounding.AwayFromZero)

                            objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price

                        End If

                    Next

                    FlagG = "Completed"
                ElseIf objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "H" And FlagH = "" Then
                    Dim NormalTotal As Double = 0
                    Dim FOCTotal As Double = 0
                    Dim TotalQty As Double = 0
                    Dim TotalFOCQty As Double = 0

                    ' 🔹 FIRST LOOP: TotalQty + TotalFOCQty
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1

                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "H" Then

                            ' ✅ SAFE FOC QTY
                            Dim FOCQtyH As Double = 0
                            Dim focVal As String = objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim()
                            If focVal <> "" Then
                                FOCQtyH = CDbl(focVal)
                            End If

                            ' ✅ SAFE NORMAL QTY
                            Dim NormalQtyH As Double = 0
                            Dim normVal As String = objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim()
                            If normVal <> "" Then
                                NormalQtyH = CDbl(normVal)
                            End If

                            ' Mark valid line if FOC exists
                            If FOCQtyH > 0 Then
                                Try
                                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                                    oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Catch ex As Exception
                                End Try
                            End If

                            ' Total Qty per row
                            If (FOCQtyH + NormalQtyH) <> 0 Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyH + NormalQtyH
                                TotalQty = TotalQty + FOCQtyH + NormalQtyH
                            Else
                                objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greater than 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If

                            ' ✅ TOTAL FOC QTY
                            TotalFOCQty = TotalFOCQty + FOCQtyH

                            Dim PurchasePriceH As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)

                            ' Normal total (kept for structure)
                            NormalTotal = NormalTotal + (NormalQtyH * PurchasePriceH)

                        End If

                    Next

                    ' 🔹 VALIDATION
                    If TotalQty <= 0 Then
                        objMain.objApplication.StatusBar.SetText("All Items total quantity is 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If

                    ' 🔹 STEP 2: Calculate FOCTotal using TOTAL FOC QTY
                    ' ⚠️ MUST be Base After Discount (2.134)
                    'Dim BaseAfterDisc As Double = CDbl(objMatrix.Columns.Item("U_BAD").Cells.Item(1).Specific.Value)

                    'FOCTotal = TotalFOCQty * BaseAfterDisc

                    ' 🔹 SECOND LOOP: Apply price
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1

                        Dim BaseAfterDisc As Double = CDbl(objMatrix.Columns.Item("U_BAD").Cells.Item(j).Specific.Value)

                        FOCTotal = TotalFOCQty * BaseAfterDisc
                        ' (kept for structure, not used)
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)

                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "H" Then

                            Dim PurchasePriceH As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)

                            ' ✅ Flat per-unit discount
                            Dim PerUnitFOCDiscount As Double = FOCTotal / TotalQty

                            ' ✅ Rounded price (3 decimals)
                            Dim Price As Double = (PurchasePriceH - PerUnitFOCDiscount)

                            objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price

                        End If

                    Next

                    FlagH = "Completed"
                ElseIf objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "I" And FlagI = "" Then
                    '    Dim NormalTotal As Double = 0
                    '    Dim FOCTotal As Double = 0
                    '    Dim TotalQty As Double = 0
                    '    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                    '        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "I" Then
                    '            Dim FOCQtyI As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                    '            Dim NormalQtyI As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                    '            If FOCQtyI > 0 Then
                    '                Try
                    '                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                    '                    oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    '                Catch ex As Exception
                    '                End Try
                    '            End If
                    '            If (FOCQtyI + NormalQtyI) <> 0 Then
                    '                objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyI + NormalQtyI
                    '            End If
                    '            Dim PurchasePriceI As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                    '            FOCTotal = FOCTotal + (FOCQtyI * PurchasePriceI)
                    '            NormalTotal = NormalTotal + (NormalQtyI * PurchasePriceI)
                    '            If FOCQtyI > 0 Or NormalQtyI > 0 Then
                    '                TotalQty = TotalQty + FOCQtyI + NormalQtyI
                    '            Else
                    '                objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greate then 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '                Exit Sub
                    '            End If
                    '        End If
                    '    Next
                    '    If TotalQty <= 0 Then
                    '        objMain.objApplication.StatusBar.SetText("All Items total quantity is 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        Exit Sub
                    '    End If
                    '    For j As Integer = 1 To objMatrix.VisualRowCount - 1
                    '        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                    '        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "I" Then

                    '            Dim FOCQtyI As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                    '            Dim NormalQtyI As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                    '            Dim PurchasePriceI As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                    '            Dim Price As Double = PurchasePriceI - ((Ratio / 100) * PurchasePriceI)
                    '            objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price
                    '        End If
                    '    Next
                    '    FlagI = "Completed"
                    Dim NormalTotal As Double = 0
                    Dim FOCTotal As Double = 0
                    Dim TotalQty As Double = 0
                    Dim TotalFOCQty As Double = 0

                    ' 🔹 FIRST LOOP: TotalQty + TotalFOCQty
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1

                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "I" Then

                            ' ✅ SAFE FOC QTY
                            Dim FOCQtyI As Double = 0
                            Dim focVal As String = objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim()
                            If focVal <> "" Then
                                FOCQtyI = CDbl(focVal)
                            End If

                            ' ✅ SAFE NORMAL QTY
                            Dim NormalQtyI As Double = 0
                            Dim normVal As String = objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim()
                            If normVal <> "" Then
                                NormalQtyI = CDbl(normVal)
                            End If

                            ' Mark valid line if FOC exists
                            If FOCQtyI > 0 Then
                                Try
                                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                                    oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Catch ex As Exception
                                End Try
                            End If

                            ' Total Qty per row
                            If (FOCQtyI + NormalQtyI) <> 0 Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyI + NormalQtyI
                                TotalQty = TotalQty + FOCQtyI + NormalQtyI
                            Else
                                objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greater than 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If

                            ' ✅ TOTAL FOC QTY
                            TotalFOCQty = TotalFOCQty + FOCQtyI

                            Dim PurchasePriceI As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)

                            ' Normal total (kept for structure)
                            NormalTotal = NormalTotal + (NormalQtyI * PurchasePriceI)

                        End If

                    Next

                    ' 🔹 VALIDATION
                    If TotalQty <= 0 Then
                        objMain.objApplication.StatusBar.SetText("All Items total quantity is 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If

                    ' 🔹 STEP 2: Calculate FOCTotal using TOTAL FOC QTY
                    ' ⚠️ MUST be Base After Discount (2.134)
                    'Dim BaseAfterDisc As Double = CDbl(objMatrix.Columns.Item("U_BAD").Cells.Item(1).Specific.Value)

                    'FOCTotal = TotalFOCQty * BaseAfterDisc

                    ' 🔹 SECOND LOOP: Apply price
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1

                        Dim BaseAfterDisc As Double = CDbl(objMatrix.Columns.Item("U_BAD").Cells.Item(j).Specific.Value)

                        FOCTotal = TotalFOCQty * BaseAfterDisc

                        ' (kept for structure, not used)
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)

                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "I" Then

                            Dim PurchasePriceI As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)

                            ' ✅ Flat per-unit discount
                            Dim PerUnitFOCDiscount As Double = FOCTotal / TotalQty

                            ' ✅ Rounded price (3 decimals)
                            Dim Price As Double = (PurchasePriceI - PerUnitFOCDiscount)

                            objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price

                        End If

                    Next

                    FlagI = "Completed"
                ElseIf objMatrix.Columns.Item("U_FOCGrp").Cells.Item(i).Specific.Value.ToString().Trim() = "J" And FlagJ = "" Then
                    'Dim NormalTotal As Double = 0
                    'Dim FOCTotal As Double = 0
                    'Dim TotalQty As Double = 0
                    'For j As Integer = 1 To objMatrix.VisualRowCount - 1
                    '    If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "J" Then
                    '        Dim FOCQtyJ As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                    '        Dim NormalQtyJ As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                    '        If FOCQtyJ > 0 Then
                    '            Try
                    '                Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                    '                oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    '            Catch ex As Exception
                    '            End Try
                    '        End If
                    '        If (FOCQtyJ + NormalQtyJ) <> 0 Then
                    '            objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyJ + NormalQtyJ
                    '        End If
                    '        Dim PurchasePriceJ As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                    '        FOCTotal = FOCTotal + (FOCQtyJ * PurchasePriceJ)
                    '        NormalTotal = NormalTotal + (NormalQtyJ * PurchasePriceJ)
                    '        If FOCQtyJ > 0 Or NormalQtyJ > 0 Then
                    '            TotalQty = TotalQty + FOCQtyJ + NormalQtyJ
                    '        Else
                    '            objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greate then 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '            Exit Sub
                    '        End If
                    '    End If
                    'Next
                    'If TotalQty <= 0 Then
                    '    objMain.objApplication.StatusBar.SetText("All Items total quantity is 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    Exit Sub
                    'End If
                    'For j As Integer = 1 To objMatrix.VisualRowCount - 1
                    '    Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)
                    '    If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "J" Then
                    '        Dim FOCQtyJ As Double = CDbl(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim())
                    '        Dim NormalQtyJ As Double = CDbl(objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim())
                    '        Dim PurchasePriceJ As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)
                    '        Dim Price As Double = PurchasePriceJ - ((Ratio / 100) * PurchasePriceJ)
                    '        objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price
                    '    End If
                    'Next
                    'FlagJ = "Completed"
                    Dim NormalTotal As Double = 0
                    Dim FOCTotal As Double = 0
                    Dim TotalQty As Double = 0
                    Dim TotalFOCQty As Double = 0

                    ' 🔹 FIRST LOOP: TotalQty + TotalFOCQty
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1

                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "J" Then

                            ' ✅ SAFE FOC QTY
                            Dim FOCQtyJ As Double = 0
                            Dim focVal As String = objMatrix.Columns.Item("U_FOCQTY").Cells.Item(j).Specific.Value.ToString().Trim()
                            If focVal <> "" Then
                                FOCQtyJ = CDbl(focVal)
                            End If

                            ' ✅ SAFE NORMAL QTY
                            Dim NormalQtyJ As Double = 0
                            Dim normVal As String = objMatrix.Columns.Item("U_normalqty").Cells.Item(j).Specific.Value.ToString().Trim()
                            If normVal <> "" Then
                                NormalQtyJ = CDbl(normVal)
                            End If

                            ' Mark valid line if FOC exists
                            If FOCQtyJ > 0 Then
                                Try
                                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LineValid")
                                    oColumn.Cells.Item(j).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Catch ex As Exception
                                End Try
                            End If

                            ' Total Qty per row
                            If (FOCQtyJ + NormalQtyJ) <> 0 Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.Value = FOCQtyJ + NormalQtyJ
                                TotalQty = TotalQty + FOCQtyJ + NormalQtyJ
                            Else
                                objMain.objApplication.StatusBar.SetText("Normal Quantity or FOC Quantity should be greater than 0 in line num:" & j.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If

                            ' ✅ TOTAL FOC QTY
                            TotalFOCQty = TotalFOCQty + FOCQtyJ

                            Dim PurchasePriceJ As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)

                            ' Normal total (kept for structure)
                            NormalTotal = NormalTotal + (NormalQtyJ * PurchasePriceJ)

                        End If

                    Next

                    ' 🔹 VALIDATION
                    If TotalQty <= 0 Then
                        objMain.objApplication.StatusBar.SetText("All Items total quantity is 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If

                    ' 🔹 STEP 2: Calculate FOCTotal using TOTAL FOC QTY
                    ' ⚠️ MUST be Base After Discount (2.134)
                    'Dim BaseAfterDisc As Double = CDbl(objMatrix.Columns.Item("U_BAD").Cells.Item(1).Specific.Value)

                    'FOCTotal = TotalFOCQty * BaseAfterDisc

                    ' 🔹 SECOND LOOP: Apply price
                    For j As Integer = 1 To objMatrix.VisualRowCount - 1

                        Dim BaseAfterDisc As Double = CDbl(objMatrix.Columns.Item("U_BAD").Cells.Item(j).Specific.Value)

                        FOCTotal = TotalFOCQty * BaseAfterDisc

                        ' (kept for structure, not used)
                        Dim Ratio As Double = (FOCTotal / (FOCTotal + NormalTotal) * 100)

                        If objMatrix.Columns.Item("U_FOCGrp").Cells.Item(j).Specific.Value.ToString().Trim() = "J" Then

                            Dim PurchasePriceJ As Double = CDbl(objMatrix.Columns.Item("U_suppprice").Cells.Item(j).Specific.Value)

                            ' ✅ Flat per-unit discount
                            Dim PerUnitFOCDiscount As Double = FOCTotal / TotalQty

                            ' ✅ Rounded price (3 decimals)
                            Dim Price As Double = (PurchasePriceJ - PerUnitFOCDiscount)

                            objMatrix.Columns.Item("14").Cells.Item(j).Specific.Value = Price

                        End If

                    Next

                    FlagJ = "Completed"
                End If
                '------------------madhu 29042026
                'madhu new logic for line total, vat amount and gross total calculation
                Dim SupplierPrice As Double = Val(objMatrix.Columns.Item("U_suppprice").Cells.Item(i).Specific.Value)
                Dim DiscPrcnt As Double = Val(objMatrix.Columns.Item("15").Cells.Item(i).Specific.Value)

                Dim NormalQuantity As Double = Val(objMatrix.Columns.Item("U_normalqty").Cells.Item(i).Specific.Value)
                Dim ExTaxofbase As Double = Val(objMatrix.Columns.Item("U_EXT").Cells.Item(i).Specific.Value)
                Dim FocQuantity As Double = Val(objMatrix.Columns.Item("U_FOCQTY").Cells.Item(i).Specific.Value)

                Dim FinalLineTotal As Double = 0


                ' If NormalQuantity > 0 And SupplierPrice > 0 And FocQuantity = 0 And ExTaxofbase = 0 Then

                'FinalLineTotal = NormalQuantity * SupplierPrice * (1 - DiscPrcnt / 100)

                'ElseIf ExTaxofbase > 0 Or FocQuantity > 0 Then

                FinalLineTotal = ((NormalQuantity * SupplierPrice) + (ExTaxofbase * FocQuantity)) * (1 - DiscPrcnt / 100)
                ' End If


                objMatrix.Columns.Item("U_LTOTAL").Cells.Item(i).Specific.Value = FinalLineTotal

                Dim VatCode As String = ""
                Dim VatPercent As Double = 0

                Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                Dim Vatgroup As String = objMatrix.Columns.Item("18").Cells.Item(i).Specific.Value.ToString()


                If Not String.IsNullOrEmpty(Vatgroup) Then
                    oRS.DoQuery("SELECT ""Rate"" FROM OVTG WHERE ""Code"" = '" & Vatgroup & "'")
                    If Not oRS.EoF Then
                        VatPercent = CDbl(oRS.Fields.Item("Rate").Value)
                    End If
                End If


                Dim VatAmount As Double = FinalLineTotal * VatPercent / 100
                objMatrix.Columns.Item("U_VATAMT").Cells.Item(i).Specific.Value = VatAmount

                Dim GrossTotal As Double = FinalLineTotal + VatAmount
                objMatrix.Columns.Item("U_GTOTAL").Cells.Item(i).Specific.Value = GrossTotal

            Next


            Dim SupplierPOTotal As Double = 0

            For i As Integer = 1 To objMatrix.VisualRowCount - 1
                Dim GT As Double = Val(objMatrix.Columns.Item("U_GTOTAL").Cells.Item(i).Specific.Value)
                SupplierPOTotal += GT
            Next

            '--Madhu 02052026 now not required clien
            'Try
            '    Dim oEdit As SAPbouiCOM.EditText = objForm.Items.Item("U_SPOTOTAL").Specific
            '    'oEdit.Value = SupplierPOTotal
            '    oEdit.Value = Math.Round(SupplierPOTotal, 3, MidpointRounding.AwayFromZero)
            'Catch ex As Exception
            '    objForm.DataSources.DBDataSources.Item("OPOR").SetValue("U_SPOTOTAL", 0, SupplierPOTotal)
            'End Try

            'Round value
            '--Madhu 02052026 now not required clien
            'Try
            '    Dim oForm As SAPbouiCOM.Form = objForm
            '    Dim DocTotal As Double = CDbl(oForm.DataSources.DBDataSources.Item("OPOR").GetValue("DocTotal", 0))
            '    Dim RoundingValue As Double = Math.Round(SupplierPOTotal - DocTotal, 3, MidpointRounding.AwayFromZero)
            '    Dim objCheckCh As SAPbouiCOM.CheckBox = oForm.Items.Item("105").Specific
            '    If objCheckCh.Checked = False Then
            '        oForm.Items.Item("105").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            '    End If
            '    oForm.Items.Item("103").Specific.Value = RoundingValue

            'Catch ex As Exception
            '    objMain.objApplication.StatusBar.SetText("RoundDiff Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'End Try

            '----madhu new logic
            'madhu new logic Untaxed Amount and Excise Tax of base calculation sugar Tax Yes
            Dim TotalUntaxedAmount As Double = 0
            Dim TotalExciseAmount As Double = 0

            For i As Integer = 1 To objMatrix.VisualRowCount - 1

                ' Get values safely
                Dim NormalQty As Double = Val(objMatrix.Columns.Item("U_normalqty").Cells.Item(i).Specific.Value)
                Dim BaseAfterDiscount As Double = Val(objMatrix.Columns.Item("U_BAD").Cells.Item(i).Specific.Value)
                Dim EXT As Double = Val(objMatrix.Columns.Item("U_EXT").Cells.Item(i).Specific.Value)
                Dim SAPQty As Double = Val(objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value)

                Dim UntaxedAmount As Double = (NormalQty * BaseAfterDiscount)
                objMatrix.Columns.Item("U_UTXAMT").Cells.Item(i).Specific.Value = Math.Round(UntaxedAmount, 3, MidpointRounding.AwayFromZero)
                'objMatrix.Columns.Item("U_UTXAMT").Cells.Item(i).Specific.Value = UntaxedAmount

                Dim ExciseTaxAmount As Double = SAPQty * EXT

                objMatrix.Columns.Item("U_EXTAX").Cells.Item(i).Specific.Value = Math.Round(ExciseTaxAmount, 3, MidpointRounding.AwayFromZero)
                'objMatrix.Columns.Item("U_EXTAX").Cells.Item(i).Specific.Value = ExciseTaxAmount


                TotalUntaxedAmount += UntaxedAmount
                TotalExciseAmount += ExciseTaxAmount

            Next

            ' Set Untaxed Amount
            Try
                Dim oEditUntaxed As SAPbouiCOM.EditText = objForm.Items.Item("TUTXAMT").Specific
                ' oEditUntaxed.Value = Math.Round(TotalUntaxedAmount, 3, MidpointRounding.AwayFromZero)
                oEditUntaxed.Value = TotalUntaxedAmount
            Catch ex As Exception
                objForm.DataSources.DBDataSources.Item("OPOR").SetValue("TUTXAMT", 0, Math.Round(TotalUntaxedAmount, 3, MidpointRounding.AwayFromZero))
            End Try

            ' Set Excise Tax
            Try
                Dim oEditExcise As SAPbouiCOM.EditText = objForm.Items.Item("TEXTAX").Specific
                'oEditExcise.Value = Math.Round(TotalExciseAmount, 3, MidpointRounding.AwayFromZero)
                oEditExcise.Value = TotalExciseAmount
            Catch ex As Exception
                objForm.DataSources.DBDataSources.Item("OPOR").SetValue("TEXTAX", 0, Math.Round(TotalExciseAmount, 3, MidpointRounding.AwayFromZero))
            End Try


            objForm.Freeze(False)


        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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

#End Region
    Public Sub GRPOFormLoad(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            If objMain.objApplication.Forms.ActiveForm.TypeEx = "143" Then
                objGRPOForm = objMain.objApplication.Forms.GetForm("143", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objGRPOMatrix = objGRPOForm.Items.Item("38").Specific
                For i As Integer = 1 To objGRPOMatrix.VisualRowCount
                    Dim ItemCode As String = objGRPOMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.ToString()
                    If ItemCode <> "" Then
                        Try
                            Dim CheckKitItem As String = "Select ""U_KITITEM"" From OITT Where ""Code"" = '" & ItemCode & "' And ""U_KITITEM"" = 'Yes'"
                            Dim oRsCheckKitItem As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsCheckKitItem.DoQuery(CheckKitItem)
                            If oRsCheckKitItem.RecordCount > 0 Then
                                Dim oColumn As SAPbouiCOM.Column = objGRPOMatrix.Columns.Item("U_SplitKitItems")
                                oColumn.Cells.Item(i).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            End If
                        Catch ex As Exception
                        End Try
                    End If
                Next
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub Items(ByVal FormUID As String)
        Try

            objMatrix = objForm.Items.Item("38").Specific
            'Dim ItemCode As String
            'For i As Integer = 1 To objMatrix.VisualRowCount - 1
            'ItemCode = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.Trim

            Dim objGrid As SAPbouiCOM.Grid
            objGrid = objForm.Items.Item("Grid").Specific
            Dim Details As String = "SELECT T1.""ItemCode""as ""ItemCode"", T1.""Dscription""as ""Item Description"", T1.""CodeBars"" as ""Barcode"", T1.""UomCode"" as ""UOM"" " &
                ",T1.""U_ForeNm"" as ""Item Arabic Description"",'' as ""Quartely Sales Quantity"",'' as ""Half Yearly Sales Quantity"",'' as""Annual Sales Quantity"" " &
                "FROM OPOR T0  INNER JOIN POR1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" WHERE T0.""DocNum""='" & objForm.Items.Item("8").Specific.Value & "' "
            objForm.DataSources.DataTables.Item(0).ExecuteQuery(Details)
            objGrid.DataTable = objForm.DataSources.DataTables.Item("MyDataTable")
            objGrid.AutoResizeColumns()
            objGrid.Columns.Item(0).Editable = False
            objGrid.Columns.Item(1).Editable = False
            objGrid.Columns.Item(2).Editable = False
            objGrid.Columns.Item(3).Editable = False
            If objGrid.Rows.Count = 0 Then
                objMain.objApplication.MessageBox("No records found")
                Exit Sub
            End If
            'Next

        Catch ex As Exception
            objForm.Close()
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub Averagesale()
        Try
            oDBs_Head = objForm.DataSources.DBDataSources.Item("OPOR")
            objMatrix = objForm.Items.Item("38").Specific
            Dim objGrid As SAPbouiCOM.Grid
            objGrid = objForm.Items.Item("Grid").Specific
            Dim UomCode As String
            Dim ItemCode As String
            Dim ItemCode2 As String
            Dim ItemCode3 As String
            Dim BarCode As String
            For i As Integer = 1 To objMatrix.VisualRowCount - 1

                ItemCode = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.ToString()
                UomCode = objMatrix.Columns.Item("1470002145").Cells.Item(i).Specific.Value.ToString()
                UomCode = objMatrix.Columns.Item("1470002145").Cells.Item(i).Specific.Value.ToString()
                BarCode = objMatrix.Columns.Item("4").Cells.Item(i).Specific.Value.ToString()
                If i = 1 Then
                    'ItemCode2 = "'" & ItemCode & "'"
                    ItemCode2 = ItemCode
                ElseIf i > 1 Then
                    'ItemCode2 = (ItemCode & "," & ItemCode2)
                    ItemCode2 = ItemCode & "'" & "," & "'" & ItemCode2

                End If
            Next

            'Dim Details As String = ("Call Averagesale ('" & ItemCode2 & "')")
            Dim Details As String = "SELECT ""ITEMCODE"" as ""Item Code"",""ITEMNAME"" as ""Item Description"",""BARCODE"" as ""Barcode"",""ITEMARABICDESCRIPTION"" as ""Item  ArabicDescription"" " &
                ",""QUARTELYSALESQUANTITY"" as ""Quartely SalesQuantity"",""HALFYEARLYSALESQUANTITY"" as ""Halfyearly SalesQuantity"" " &
                ",""ANNUALSALESQUANTITY"" as""Annual SalesQuantity"",""MONTHLYAVERAGESALESQUANTITY"" AS ""Monthly Average Sales"" from SBO_AverageSale() WHERE ""ITEMCODE"" IN ('" & ItemCode2 & "') " '' and ""BARCODE"" = '" & BarCode & "'"



            objForm.DataSources.DataTables.Item(0).ExecuteQuery(Details)
            objGrid.DataTable = objForm.DataSources.DataTables.Item("MyDataTable")
            objGrid.AutoResizeColumns()
            objGrid.Columns.Item(0).Editable = False
            objGrid.Columns.Item(1).Editable = False
            objGrid.Columns.Item(2).Editable = False
            objGrid.Columns.Item(3).Editable = False
            'objGrid.Columns.Item(4).Editable = False
            If objGrid.Rows.Count = 0 Then
                objMain.objApplication.MessageBox("No records found")
                Exit Sub
            End If



        Catch ex As Exception
            'objForm.Close()
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("38").Specific

            If CalcualtionBtnFlag = False Then
                Dim PObutton As String = "Select ""U_PO""as ""PO"" from ""@ME_ACFG"""
                Dim rsPObutton As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rsPObutton.DoQuery(PObutton)
                If rsPObutton.Fields.Item("PO").Value = "Y" Then
                    '    Exit Function
                    'Else
                    objMain.objApplication.StatusBar.SetText("Please run the calculation before generating document.....!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

            End If

            'madhu

            For i As Integer = 1 To objMatrix.VisualRowCount - 1
                Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.Trim
                Dim FOCQty As Double = objMatrix.Columns.Item("U_FOCQTY").Cells.Item(i).Specific.Value.ToString().Trim()
                If (ItemCode <> "" And objMatrix.Columns.Item("U_LineValid").Cells.Item(i).Specific.Value.Trim <> "Yes" And FOCQty > 0) Then
                    objMain.objApplication.StatusBar.SetText("For Item " & ItemCode & " And Line No. " & i & "... Please calculate FOC.....!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Next
            For i As Integer = 1 To objMatrix.VisualRowCount - 1
                Dim LineValidation As String = "Select ""U_Least""as ""Least"" from ""@ME_ACFG"""
                Dim oLValidation As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oLValidation.DoQuery(LineValidation)
                If oLValidation.Fields.Item("Least").Value = "Y" Then
                    Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.Trim
                    If (ItemCode <> "" And objMatrix.Columns.Item("U_LPValidate").Cells.Item(i).Specific.Value.Trim = "N") Then
                        objMain.objApplication.StatusBar.SetText("For Item " & ItemCode & " And Line No. " & i & ".... Please rechek least price.....!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            Next


            For i As Integer = 1 To objMatrix.VisualRowCount - 1
                Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.Trim
                Dim NormalQty As Double = objMatrix.Columns.Item("U_normalqty").Cells.Item(i).Specific.Value.Trim
                Dim FOCQty As Double = objMatrix.Columns.Item("U_FOCQTY").Cells.Item(i).Specific.Value.Trim
                Dim Qty As Double = objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value.Trim
                If (Qty <> (NormalQty + FOCQty)) Then
                    objMain.objApplication.StatusBar.SetText("For Item " & ItemCode & " And Line No. " & i & "....OverAll Qty should balance NormalQty And FOCQty.....!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Next
            ''If pVal.ItemUID = "38" And pVal.ColUID = "1" And pVal.BeforeAction = False Then
            'For i As Integer = 1 To objMatrix.VisualRowCount - 1
            '    If objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value <> "" Then
            '        objMatrix = objForm.Items.Item("38").Specific
            '        Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value
            '        Dim GetDocEntry As String = "Select ""FrgnName"" From OITM Where ""ItemCode"" = '" & ItemCode & "' "
            '        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '        oRs.DoQuery(GetDocEntry)
            '        objMatrix.Columns.Item("U_ForeNm").Cells.Item(i).Specific.Value() = oRs.Fields.Item("FrgnName").Value()
            '    End If
            'Next
            'End If
            'For i As Integer = 1 To objMatrix.VisualRowCount - 1
            '    Dim stQry As String = "SELECT *  FROM OPOR T0  INNER JOIN POR1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" WHERE T0.""CardCode""='" & objForm.Items.Item("4").Specific.Value & "' AND  T1.""ItemCode"" ='" & objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value & "' and  T1.""CodeBars"" ='" & objMatrix.Columns.Item("4").Cells.Item(i).Specific.Value & "'and T0.""DocStatus""='O'"
            '    rsPO = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '    rsPO.DoQuery(stQry)
            '    If rsPO.RecordCount > 0 Then
            '        objMain.objApplication.StatusBar.SetText("This Purchase order already exists in purchase order " & rsPO.Fields.Item("DocNum").Value & ".....!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '        Return False
            '    End If

            'Next
            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function
    Public Function getSupplierPrice(ByVal FormUID As String, ByVal IsUomDefault As String,
                                     ByVal ItemCode As String, ByVal Vendor As String, ByVal UOMEntry As Integer, ByVal UomCode As String) As Double
        Dim SuppPrice As Double = 0
        Dim DflUomPrice As Double = 0
        Dim BaseQty As Double = 0
        Dim UnitCost As Double = 0
        Dim OtherUOMPrice As Double = 0
        Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            If IsUomDefault = "No" Then
                Dim GetPurchasePriceDfltUOM As String = "Select ""Price"" From ITM1 Where ""ItemCode""='" & ItemCode & "' AND ""PriceList""= (Select ""ListNum"" From OCRD Where ""CardCode"" = '" & Vendor & "')"
                Dim oRsGetPurchasePriceDfltUOM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetPurchasePriceDfltUOM.DoQuery(GetPurchasePriceDfltUOM)
                If oRsGetPurchasePriceDfltUOM.RecordCount > 0 Then
                    DflUomPrice = CDbl(oRsGetPurchasePriceDfltUOM.Fields.Item(0).Value)
                    BaseQty = Me.GetBaseQty(ItemCode, UOMEntry.ToString)
                    SuppPrice = BaseQty * DflUomPrice
                Else
                    Dim GetPurchasePrice As String = " Select T0.""Price"" , T0.""UomEntry"" From ITM9 T0 Where T0.""PriceList""= " &
                                                     " (Select ""ListNum"" From OCRD Where ""CardCode"" = '" & Vendor & "') And " &
                                                     " ""ItemCode""='" & ItemCode & "' "
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
                Dim GetPurchasePrice As String = " Select T0.""Price"" , T0.""UomEntry"" From ITM9 T0 Where T0.""PriceList""= " &
                                                                    " (Select ""ListNum"" From OCRD Where ""CardCode"" = '" & Vendor & "') And " &
                                                                    " ""ItemCode""='" & ItemCode & "' "
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
    Private Sub LoadOnItemOrUOMSelection_SpecialPricesmadhu(ByVal FormUID As String, ByVal RowId As Integer, ByVal IsUomDefault As String, ByVal ItemCode As String, ByVal Uom As String, ByVal Vendor As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("OPOR")
            objMatrix = objForm.Items.Item("38").Specific
            Dim SalesPrice As Double = 0
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
                Else
                    SalesPrice = Me.getSupplierPrice_SpecialPrices(objForm.UniqueID, IsUomDefault, ItemCode, Vendor, UOMEntry, Uom)
                End If

                Dim GetPackageCode As String = " Select Cast(Cast(T0.""BaseQty"" As Int) As Varchar)" & Concate & "'*'" & Concate & "'" & BaseUom & "'" &
                                               " From UGP1 T0 Inner Join OUOM T1 On  " &
                                               " T1.""UomEntry"" = T0.""UomEntry""  Where ""UgpEntry"" = '" & Me.GetUgpEntry(ItemCode) & "' And  " &
                                               " T1.""UomCode"" = '" & Uom & "' "
                Dim oRsGetPackageCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetPackageCode.DoQuery(GetPackageCode)
                If PCPrPack <> "" Then
                    Packing = PCPrPack & "*" & oRsGetPackageCode.Fields.Item(0).Value
                End If
                Packing = oRsGetPackageCode.Fields.Item(0).Value
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
                '                                          " Inner Join OCRD T1 On T1.""ListNum"" = T0.""PriceList""  And T1.""CardCode"" In (Select ""VendorCode"" From ITM2 Where  ITM2.""ItemCode"" = T0.""ItemCode"") Where " & _
                '                                          " ""ItemCode"" = '" & ItemCode & "' And T1.""CardType"" = 'S' And  T0.""UomEntry"" = '" & UOMEntry & "' " & _
                '                                          " Order By T0.""Price""  Asc "

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

                If oRsGetPurchasePriceDfltUOM.RecordCount > 0 Then
                    SalesPrice = CDbl(oRsGetPurchasePriceDfltUOM.Fields.Item(0).Value)
                Else
                    SalesPrice = Me.getSupplierPrice_SpecialPrices(objForm.UniqueID, IsUomDefault, ItemCode, Vendor, UOMEntry, Uom)
                End If

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
                '                                            " Inner Join OCRD T1 On T1.""ListNum"" = T0.""PriceList"" And T1.""CardCode"" In (Select ""VendorCode"" From ITM2 Where  ITM2.""ItemCode"" = T0.""ItemCode"") Where " & _
                '                                            " ""ItemCode"" = '" & ItemCode & "' And T1.""CardType"" = 'S' " & _
                '                                            " Order By T0.""Price""  Asc "

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

            objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value = SalesPrice

            objMatrix.Columns.Item("U_Packing").Cells.Item(RowId).Specific.Value = Packing
            objMatrix.Columns.Item("U_RetailPrice").Cells.Item(RowId).Specific.Value = RetailsPrice
            'objMatrix.Columns.Item("U_Addbarcode").Cells.Item(RowId).Specific.Value = AdditionalBarCode
            objMatrix.Columns.Item("U_LVNDPRC").Cells.Item(RowId).Specific.Value = LeastVendorPrice
            objMatrix.Columns.Item("U_LPVCD").Cells.Item(RowId).Specific.Value = LeastVendorCode
            objMatrix.Columns.Item("U_LPVNM").Cells.Item(RowId).Specific.Value = LeastVendorName
            Dim NetCost As String = objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value
            If NetCost <> "" Then
                NetCost = NetCost.Replace(objForm.Items.Item("63").Specific.Value, "")
                NetCost = Convert.ToDouble(NetCost.Trim)
                Dim LeastPrice As Double = objMatrix.Columns.Item("U_LVNDPRC").Cells.Item(RowId).Specific.Value
                If LeastPrice = 0 Then
                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LPValidate")
                    oColumn.Cells.Item(RowId).Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)
                End If
                If NetCost > 0 And LeastPrice > 0 Then
                    If NetCost > LeastPrice Then
                        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LPValidate")
                        oColumn.Cells.Item(RowId).Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    Else
                        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LPValidate")
                        oColumn.Cells.Item(RowId).Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If
                End If
                '-------------------Below cost---------------------------
                If objMatrix.Columns.Item("14").Cells.Item(RowId).Specific.Value = "" Then
                    NetCost = 0
                Else
                    NetCost = objMatrix.Columns.Item("14").Cells.Item(RowId).Specific.Value
                    NetCost = NetCost.Replace(objForm.Items.Item("63").Specific.Value, "")
                    NetCost = Convert.ToDouble(NetCost.Trim)
                End If
                Dim RetailPrice As Double = objMatrix.Columns.Item("U_RetailPrice").Cells.Item(RowId).Specific.Value
                If RetailPrice > 0 And NetCost > 0 Then
                    If NetCost > RetailPrice Then
                        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_BelowCst")
                        oColumn.Cells.Item(RowId).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If
                End If
            End If
            '-------------------Below cost---------------------------
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Function getSupplierPrice_SpecialPricesmadhu(ByVal FormUID As String, ByVal IsUomDefault As String,
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

                Dim GetPurchasePriceDfltUOM As String = " Select IFNULL(T3.""Price"",0) ""Price"" from SPP2 T3  " &
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

                    Dim GetPurchasePrice As String = " Select IFNULL(T3.""Price"",0) ""Price"",T3.""UomEntry"" from SPP2 T3  " &
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

                Dim GetPurchasePrice As String = " Select IFNULL(T3.""Price"",0) ""Price"",T3.""UomEntry"" from SPP2 T3  " &
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
    Private Sub LoadOnItemOrUOMSelection_LostFocus(ByVal FormUID As String, ByVal RowId As Integer, ByVal IsUomDefault As String, ByVal ItemCode As String, ByVal Uom As String, ByVal Vendor As String, ByVal [Date] As String, ByVal BarCode As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("OPOR")
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
            Dim GetBARANME As String = "Select ""BcdName"" from ""OBCD"" T0 Left Join ""OITM"" T1 ON T0.""ItemCode""=T1.""ItemCode"" where T0.""BcdCode"" = '" & BarCode & "' and T0.""ItemCode"" = '" & ItemCode & "'"

            Dim oRsGetBARANME As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetBARANME.DoQuery(GetBARANME)

            objMatrix.Columns.Item("U_BRNM").Cells.Item(RowId).Specific.Value = oRsGetBARANME.Fields.Item("BcdName").Value

            If IsUomDefault = "No" Then
                UOMEntry = Me.GetUomEntry(Uom)
            ElseIf IsUomDefault = "Yes" Then
                UOMEntry = Me.GetUomEntry(Me.GetDefaultUomCode(ItemCode))
            End If
            Dim GetBase As String = " SELECT T0.""U_SUTAXAMT"" AS ""Price"",""U_ETAX"" FROM ""OBCD"" T0 LEFT JOIN OITM T1 ON T1.""ItemCode""=T0.""ItemCode"" 
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

            badVal = baseVal - discountVal
            objMatrix.Columns.Item("U_BAD").Cells.Item(RowId).Specific.Value = badVal

            'extVal = (baseVal * 50) / 100
            'objMatrix.Columns.Item("U_EXT").Cells.Item(RowId).Specific.Value = extVal
            'extVal = Math.Round((baseVal * 50) / 100, 3, MidpointRounding.AwayFromZero)
            extVal = oRsGetbase.Fields.Item("U_ETAX").Value
            objMatrix.Columns.Item("U_EXT").Cells.Item(RowId).Specific.Value = oRsGetbase.Fields.Item("U_ETAX").Value

            ' Check if the column is editable
            'Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("14")
            'oColumn.Editable = True
            objMatrix.Columns.Item("14").Cells.Item(RowId).Specific.Value = (badVal + extVal)
            'objMatrix.Columns.Item("14").Cells.Item(RowId).Specific.Value = 3

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
AND TO_DATE('" & [Date] & "', 'YYYYMMDD') >= T1.""FromDate""
AND (
        T1.""ToDate"" IS NULL
        OR TO_DATE('" & [Date] & "', 'YYYYMMDD') <= T1.""ToDate""
    )
ORDER BY T0.""CardCode"",T0.""ItemCode"",T1.""FromDate"""

            Dim oRsGetSPP1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetSPP1.DoQuery(GetSPP1)



            SppPrice = oRsGetSPP1.Fields.Item("Effective Price").Value



            Dim GetSPP As String = "CALL ""TNX_GETITEMPRICE""('" & BarCode & "', " & UOMEntry & ", '" & ItemCode & "')"
            Dim oRsGetSPP As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetSPP.DoQuery(GetSPP)

            Dim GetSugar As String = "SELECT T0.""U_SugarTax"" FROM OITM T0 WHERE T0.""ItemCode"" = '" & ItemCode & "'"

            Dim oRsGetSugar As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetSugar.DoQuery(GetSugar)

            Dim oColumnSUpp As SAPbouiCOM.Column = objMatrix.Columns.Item("U_suppprice")
            oColumnSUpp.Editable = True

            If oRsGetSugar.Fields.Item("U_SugarTax").Value = "Y" Then
                objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value = (badVal + extVal)
                objMatrix.Columns.Item("U_SuppPriBeDis").Cells.Item(RowId).Specific.Value = (badVal + extVal)
            Else
                If SppPrice <> 0 Then
                    objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value = SppPrice
                    objMatrix.Columns.Item("U_SuppPriBeDis").Cells.Item(RowId).Specific.Value = SppPrice
                Else
                    objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value = oRsGetSPP.Fields.Item("Price").Value
                    objMatrix.Columns.Item("U_SuppPriBeDis").Cells.Item(RowId).Specific.Value = oRsGetSPP.Fields.Item("Price").Value
                End If

            End If


            'Monthly Average Sales Quantity
            Dim MonthlyAvgSales As Double = 0

            Dim MonthlyAvgQry As String = "SELECT
    CAST(
        COALESCE((
                COALESCE(
                    (
                        SELECT SUM(T1.""InvQty"")
                        FROM ""OINV"" T0
                        INNER JOIN ""INV1"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                        WHERE T1.""ItemCode"" = '" & ItemCode & "'
                    ),0
                )
                /
                NULLIF(
                    MONTHS_BETWEEN(
                        (
                            SELECT MIN(""DocDate"")
                            FROM ""OINM""
                            WHERE ""ItemCode"" = '" & ItemCode & "'
                              AND ""InQty"" > 0
                        ),
                        CURRENT_DATE
                    ),
                0)
            ),
        0
        ) AS DECIMAL(19,6)
    ) AS ""MonthlyAvg""
FROM DUMMY"

            Dim oRsMonthlyAvg As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsMonthlyAvg.DoQuery(MonthlyAvgQry)

            If oRsMonthlyAvg.RecordCount > 0 Then
                Double.TryParse(oRsMonthlyAvg.Fields.Item("MonthlyAvg").Value.ToString(), MonthlyAvgSales)
            End If

            'Assign to Matrix Column
            objMatrix.Columns.Item("U_MNSL").Cells.Item(RowId).Specific.Value = MonthlyAvgSales.ToString("0.00")

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    'Ramya
    Private Sub LoadOnItemOrUOMSelection_SpecialPrices(ByVal FormUID As String, ByVal RowId As Integer, ByVal IsUomDefault As String, ByVal ItemCode As String, ByVal Uom As String, ByVal Vendor As String, ByVal [Date] As String, ByVal BarCode As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("OPOR")
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
                Else
                    SalesPrice = Me.getSupplierPrice_SpecialPrices(objForm.UniqueID, IsUomDefault, ItemCode, Vendor, UOMEntry, Uom)
                End If

                Dim GetPackageCode As String = " Select Cast(Cast(T0.""BaseQty"" As Int) As Varchar)" & Concate & "'*'" & Concate & "'" & BaseUom & "'" &
                                           " From UGP1 T0 Inner Join OUOM T1 On  " &
                                           " T1.""UomEntry"" = T0.""UomEntry""  Where ""UgpEntry"" = '" & Me.GetUgpEntry(ItemCode) & "' And  " &
                                           " T1.""UomCode"" = '" & Uom & "' "
                Dim oRsGetPackageCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetPackageCode.DoQuery(GetPackageCode)
                If PCPrPack <> "" Then
                    Packing = PCPrPack & "*" & oRsGetPackageCode.Fields.Item(0).Value
                End If
                Packing = oRsGetPackageCode.Fields.Item(0).Value
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
                '                                          " Inner Join OCRD T1 On T1.""ListNum"" = T0.""PriceList""  And T1.""CardCode"" In (Select ""VendorCode"" From ITM2 Where  ITM2.""ItemCode"" = T0.""ItemCode"") Where " & _
                '                                          " ""ItemCode"" = '" & ItemCode & "' And T1.""CardType"" = 'S' And  T0.""UomEntry"" = '" & UOMEntry & "' " & _
                '                                          " Order By T0.""Price""  Asc "

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

                If oRsGetPurchasePriceDfltUOM.RecordCount > 0 Then
                    SalesPrice = CDbl(oRsGetPurchasePriceDfltUOM.Fields.Item(0).Value)
                Else
                    SalesPrice = Me.getSupplierPrice_SpecialPrices(objForm.UniqueID, IsUomDefault, ItemCode, Vendor, UOMEntry, Uom)
                End If

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
                '                                            " Inner Join OCRD T1 On T1.""ListNum"" = T0.""PriceList"" And T1.""CardCode"" In (Select ""VendorCode"" From ITM2 Where  ITM2.""ItemCode"" = T0.""ItemCode"") Where " & _
                '                                            " ""ItemCode"" = '" & ItemCode & "' And T1.""CardType"" = 'S' " & _
                '                                            " Order By T0.""Price""  Asc "

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




            'Dim GetBARANME As String = "Select ""BcdName"" from ""OBCD"" T0 Left Join ""OITM"" T1 ON T0.""ItemCode""=T1.""ItemCode"" where T0.""BcdCode"" = '" & BarCode & "' and T0.""ItemCode"" = '" & ItemCode & "'"

            'Dim oRsGetBARANME As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRsGetBARANME.DoQuery(GetBARANME)

            'objMatrix.Columns.Item("U_BRNM").Cells.Item(RowId).Specific.Value = oRsGetBARANME.Fields.Item("BcdName").Value

            'CType(objMatrix.Columns.Item("U_Packing").Cells.Item(RowId).Specific, SAPbouiCOM.EditText).Value = Packing






            'Commented by ramya for testing  on 03-08-2026
            objMatrix.Columns.Item("U_Packing").Cells.Item(RowId).Specific.Value = Packing
            objMatrix.Columns.Item("U_RetailPrice").Cells.Item(RowId).Specific.Value = RetailsPrice
            'objMatrix.Columns.Item("U_Addbarcode").Cells.Item(RowId).Specific.Value = AdditionalBarCode
            objMatrix.Columns.Item("U_LVNDPRC").Cells.Item(RowId).Specific.Value = LeastVendorPrice
            objMatrix.Columns.Item("U_LPVCD").Cells.Item(RowId).Specific.Value = LeastVendorCode
            objMatrix.Columns.Item("U_LPVNM").Cells.Item(RowId).Specific.Value = LeastVendorName
            Dim NetCost As String = objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value
            If NetCost <> "" Then
                NetCost = NetCost.Replace(objForm.Items.Item("63").Specific.Value, "")
                NetCost = Convert.ToDouble(NetCost.Trim)
                Dim LeastPrice As Double = objMatrix.Columns.Item("U_LVNDPRC").Cells.Item(RowId).Specific.Value
                'Commented by ramya for testing
                If LeastPrice = 0 Then
                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LPValidate")
                    oColumn.Cells.Item(RowId).Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)
                End If
                If NetCost > 0 And LeastPrice > 0 Then
                    If NetCost > LeastPrice Then
                        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LPValidate")
                        oColumn.Cells.Item(RowId).Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    Else
                        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LPValidate")
                        oColumn.Cells.Item(RowId).Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If
                End If
                '-------------------Below cost---------------------------
                If objMatrix.Columns.Item("14").Cells.Item(RowId).Specific.Value = "" Then
                    NetCost = 0
                Else
                    NetCost = objMatrix.Columns.Item("14").Cells.Item(RowId).Specific.Value
                    NetCost = NetCost.Replace(objForm.Items.Item("63").Specific.Value, "")
                    NetCost = Convert.ToDouble(NetCost.Trim)
                End If
                Dim RetailPrice As Double = objMatrix.Columns.Item("U_RetailPrice").Cells.Item(RowId).Specific.Value
                If RetailPrice > 0 And NetCost > 0 Then
                    If NetCost > RetailPrice Then
                        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_BelowCst")
                        oColumn.Cells.Item(RowId).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If
                End If
            End If
            '-------------------Below cost---------------------------
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    'madhu
    '    Private Sub LoadOnItemOrUOMSelection_SpecialPrices(ByVal FormUID As String, ByVal RowId As Integer, ByVal IsUomDefault As String, ByVal ItemCode As String, ByVal Uom As String, ByVal Vendor As String, ByVal [Date] As String, ByVal BarCode As String)
    '        Try
    '            objForm = objMain.objApplication.Forms.Item(FormUID)
    '            oDBs_Head = objForm.DataSources.DBDataSources.Item("OPOR")
    '            objMatrix = objForm.Items.Item("38").Specific
    '            Dim SalesPrice As Double = 0
    '            Dim SppPrice As Double = 0
    '            Dim Packing As String = ""
    '            Dim RetailsPrice As Double = 0
    '            Dim AdditionalBarCode As String = ""
    '            Dim UOMEntry As Integer = 0
    '            Dim LeastVendorPrice As Double = 0
    '            Dim LeastVendorCode As String = ""
    '            Dim LeastVendorName As String = ""
    '            Dim LeastVendorValidation As String = ""
    '            Dim PCPrPack As String = objMain.objUtilities.GetPCPerItem(ItemCode)
    '            BaseUom = objMain.objUtilities.GetItemBaseUom(ItemCode)
    '            If IsUomDefault = "No" Then
    '                UOMEntry = Me.GetUomEntry(Uom)
    '                'Dim GetPurchasePrice As String = " Select T0.""Price"" From ITM9 T0 Where T0.""PriceList""= " & _
    '                '                                 " (Select ""ListNum"" From OCRD Where ""CardCode"" = '" & Vendor & "') And " & _
    '                '                                 " T0.""UomEntry""='" & UOMEntry & "' And ""ItemCode""='" & ItemCode & "' "

    '                Dim GetPurchasePrice As String = " Select IFNULL(T3.""Price"",0) ""Price"" from SPP2 T3  " &
    '            " Where T3.""ItemCode"" = '" & ItemCode & "' and T3.""CardCode"" = '" & Vendor & "' and T3.""UomEntry"" = '" & UOMEntry & "' "

    '                Dim oRsGetPurchasePrice As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                oRsGetPurchasePrice.DoQuery(GetPurchasePrice)
    '                If oRsGetPurchasePrice.RecordCount > 0 Then
    '                    SalesPrice = CDbl(oRsGetPurchasePrice.Fields.Item(0).Value)
    '                Else
    '                    SalesPrice = Me.getSupplierPrice_SpecialPrices(objForm.UniqueID, IsUomDefault, ItemCode, Vendor, UOMEntry, Uom)
    '                End If

    '                Dim GetPackageCode As String = " Select Cast(Cast(T0.""BaseQty"" As Int) As Varchar)" & Concate & "'*'" & Concate & "'" & BaseUom & "'" &
    '                                           " From UGP1 T0 Inner Join OUOM T1 On  " &
    '                                           " T1.""UomEntry"" = T0.""UomEntry""  Where ""UgpEntry"" = '" & Me.GetUgpEntry(ItemCode) & "' And  " &
    '                                           " T1.""UomCode"" = '" & Uom & "' "
    '                Dim oRsGetPackageCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                oRsGetPackageCode.DoQuery(GetPackageCode)
    '                If PCPrPack <> "" Then
    '                    Packing = PCPrPack & "*" & oRsGetPackageCode.Fields.Item(0).Value
    '                End If
    '                Packing = oRsGetPackageCode.Fields.Item(0).Value
    '                'Dim GetRetailPriceBrnchWise As String = " Select T1.""Price"" From OITM T0 Inner Join ITM9 T1 On T1.""ItemCode"" = T0.""ItemCode"" " & _
    '                '                                        " Inner Join OBPL T2 On T1.""PriceList""=T2.""U_PricelistNo"" Where T0.""ItemCode"" = '" & ItemCode & "' " & _
    '                '                                        " And T2.""BPLId""= '" & objForm.Items.Item("2001").Specific.Value.Trim & "' And T1.""UomEntry""='" & UOMEntry & "'"

    '                Dim GetRetailPriceBrnchWise As String = " Select T0.""Price"" From SPP2 T0 Inner Join ITM1 T1 On T1.""ItemCode"" = T0.""ItemCode"" " &
    '            " Inner Join OBPL T2 On T1.""PriceList"" = T2.""U_PricelistNo"" " &
    '            " Where T0.""ItemCode"" = '" & ItemCode & "'  And T2.""BPLId""= '" & objForm.Items.Item("2001").Specific.Value.Trim & "' And T0.""UomEntry""='" & UOMEntry & "' "

    '                Dim oRsGetRetailPriceBrnchWise As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                oRsGetRetailPriceBrnchWise.DoQuery(GetRetailPriceBrnchWise)
    '                If oRsGetRetailPriceBrnchWise.RecordCount > 0 Then
    '                    RetailsPrice = oRsGetRetailPriceBrnchWise.Fields.Item("Price").Value
    '                End If
    '                'Dim GetLeastVendorDtlsAndPrice As String = " Select Top 1 T0.""ItemCode"" , T1.""CardCode"", T1.""CardName"" , T0.""Price"" From ITM9 T0 " & _
    '                '                                          " Inner Join OCRD T1 On T1.""ListNum"" = T0.""PriceList""  And T1.""CardCode"" In (Select ""VendorCode"" From ITM2 Where  ITM2.""ItemCode"" = T0.""ItemCode"") Where " & _
    '                '                                          " ""ItemCode"" = '" & ItemCode & "' And T1.""CardType"" = 'S' And  T0.""UomEntry"" = '" & UOMEntry & "' " & _
    '                '                                          " Order By T0.""Price""  Asc "

    '                Dim GetLeastVendorDtlsAndPrice As String = " Select Top 1 T0.""ItemCode"", T1.""CardCode"", T1.""CardName"", T0.""Price""  " &
    '            " From SPP2 T0 Inner Join OCRD T1 On T1.""CardCode"" = T0.""CardCode""  " &
    '            " Where ""ItemCode"" = '" & ItemCode & "' And T1.""CardType"" = 'S' Order By T0.""Price""  Asc "

    '                Dim oRsGetLeastVendorDtlsAndPrice As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                oRsGetLeastVendorDtlsAndPrice.DoQuery(GetLeastVendorDtlsAndPrice)
    '                If oRsGetLeastVendorDtlsAndPrice.RecordCount > 0 Then
    '                    LeastVendorPrice = oRsGetLeastVendorDtlsAndPrice.Fields.Item("Price").Value
    '                    LeastVendorCode = oRsGetLeastVendorDtlsAndPrice.Fields.Item("CardCode").Value
    '                    LeastVendorName = oRsGetLeastVendorDtlsAndPrice.Fields.Item("CardName").Value
    '                End If
    '            ElseIf IsUomDefault = "Yes" Then
    '                UOMEntry = Me.GetUomEntry(Me.GetDefaultUomCode(ItemCode))
    '                'Dim GetPurchasePriceDfltUOM As String = "Select ""Price"" From ITM1 Where ""ItemCode""='" & ItemCode & "' AND ""PriceList""= (Select ""ListNum"" From OCRD Where ""CardCode"" = '" & Vendor & "')"
    '                Dim GetPurchasePriceDfltUOM As String = " Select IFNULL(T3.""Price"",0) ""Price"" from SPP2 T3 " &
    '            " Where ""ItemCode"" = '" & ItemCode & "' and ""CardCode"" = '" & Vendor & "' and T3.""UomEntry"" = '" & UOMEntry & "' "

    '                Dim oRsGetPurchasePriceDfltUOM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                oRsGetPurchasePriceDfltUOM.DoQuery(GetPurchasePriceDfltUOM)

    '                If oRsGetPurchasePriceDfltUOM.RecordCount > 0 Then
    '                    SalesPrice = CDbl(oRsGetPurchasePriceDfltUOM.Fields.Item(0).Value)
    '                Else
    '                    SalesPrice = Me.getSupplierPrice_SpecialPrices(objForm.UniqueID, IsUomDefault, ItemCode, Vendor, UOMEntry, Uom)
    '                End If

    '                If PCPrPack <> "" Then
    '                    Packing = PCPrPack & "*" & "1*" & BaseUom
    '                Else
    '                    Packing = "1*" & BaseUom
    '                End If

    '                'Dim GetRetailPriceBrnchWise As String = " Select T1.""Price"" From OITM T0 Inner Join ITM1 T1 On T1.""ItemCode"" = T0.""ItemCode"" " & _
    '                '                                        " Inner Join OBPL T2 On T1.""PriceList""=T2.""U_PricelistNo"" Where T0.""ItemCode"" = '" & ItemCode & "' " & _
    '                '                                        " And T2.""BPLId""= '" & objForm.Items.Item("2001").Specific.Value.Trim & "' "

    '                Dim GetRetailPriceBrnchWise As String = " Select T0.""Price"" From SPP2 T0 Inner Join ITM1 T1 On T1.""ItemCode"" = T0.""ItemCode"" " &
    '            " Inner Join OBPL T2 On T1.""PriceList"" = T2.""U_PricelistNo"" " &
    '            " Where T0.""ItemCode"" = '" & ItemCode & "'  And T2.""BPLId""= '" & objForm.Items.Item("2001").Specific.Value.Trim & "' And T0.""UomEntry""='" & UOMEntry & "' "

    '                Dim oRsGetRetailPriceBrnchWise As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                oRsGetRetailPriceBrnchWise.DoQuery(GetRetailPriceBrnchWise)
    '                If oRsGetRetailPriceBrnchWise.RecordCount > 0 Then
    '                    RetailsPrice = oRsGetRetailPriceBrnchWise.Fields.Item("Price").Value
    '                End If
    '                'Dim GetLeastVendorDtlsAndPrice As String = " Select Top 1 T0.""ItemCode"" , T1.""CardCode"", T1.""CardName"" , T0.""Price"" From ITM1 T0 " & _
    '                '                                            " Inner Join OCRD T1 On T1.""ListNum"" = T0.""PriceList"" And T1.""CardCode"" In (Select ""VendorCode"" From ITM2 Where  ITM2.""ItemCode"" = T0.""ItemCode"") Where " & _
    '                '                                            " ""ItemCode"" = '" & ItemCode & "' And T1.""CardType"" = 'S' " & _
    '                '                                            " Order By T0.""Price""  Asc "

    '                Dim GetLeastVendorDtlsAndPrice As String = " Select Top 1 T0.""ItemCode"", T1.""CardCode"", T1.""CardName"", T0.""Price""  " &
    '            " From SPP2 T0 Inner Join OCRD T1 On T1.""CardCode"" = T0.""CardCode""  " &
    '            " Where ""ItemCode"" = '" & ItemCode & "' And T1.""CardType"" = 'S' Order By T0.""Price""  Asc "

    '                Dim oRsGetLeastVendorDtlsAndPrice As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                oRsGetLeastVendorDtlsAndPrice.DoQuery(GetLeastVendorDtlsAndPrice)
    '                If oRsGetLeastVendorDtlsAndPrice.RecordCount > 0 Then
    '                    LeastVendorPrice = oRsGetLeastVendorDtlsAndPrice.Fields.Item("Price").Value
    '                    LeastVendorCode = oRsGetLeastVendorDtlsAndPrice.Fields.Item("CardCode").Value
    '                    LeastVendorName = oRsGetLeastVendorDtlsAndPrice.Fields.Item("CardName").Value
    '                End If
    '            End If
    '            If objMatrix.Columns.Item("4").Cells.Item(RowId).Specific.Value = "" Then
    '                Dim GetBarCode As String = " SELECT Top 1 ""BcdCode"" FROM OBCD WHERE ""UomEntry"" = '" & UOMEntry & "' " &
    '                                       " And ""ItemCode"" = '" & objMatrix.Columns.Item("1").Cells.Item(RowId).Specific.Value & "' ORDER BY ""BcdEntry"" ASC "
    '                Dim oRsGetBarCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                oRsGetBarCode.DoQuery(GetBarCode)
    '                If oRsGetBarCode.RecordCount > 0 Then
    '                    AdditionalBarCode = oRsGetBarCode.Fields.Item("BcdCode").Value
    '                Else
    '                    Dim GetBarCode1 As String = " SELECT Top 1 ""BcdCode"" FROM OBCD WHERE " &
    '                                            " ""ItemCode"" = '" & objMatrix.Columns.Item("1").Cells.Item(RowId).Specific.Value & "' ORDER BY ""BcdEntry"" ASC "
    '                    Dim oRsGetBarCode1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                    oRsGetBarCode1.DoQuery(GetBarCode1)
    '                    If oRsGetBarCode1.RecordCount > 0 Then
    '                        AdditionalBarCode = oRsGetBarCode1.Fields.Item("BcdCode").Value
    '                    End If
    '                End If
    '            Else
    '                AdditionalBarCode = objMatrix.Columns.Item("4").Cells.Item(RowId).Specific.Value
    '            End If


    '            Dim GetBase As String = " SELECT T0.""U_SUTAXAMT"" AS ""Price"" FROM ""OBCD"" T0 LEFT JOIN OITM T1 ON T1.""ItemCode""=T0.""ItemCode"" 
    ' WHERE T0.""ItemCode"" =  '" & ItemCode & "' and T0.""UomEntry"" = '" & UOMEntry & "' and T1.""U_SugarTax""='Y'"

    '            Dim oRsGetbase As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '            oRsGetbase.DoQuery(GetBase)



    '            objMatrix.Columns.Item("U_BASE").Cells.Item(RowId).Specific.Value = oRsGetbase.Fields.Item("Price").Value

    '            Dim Discount As Double = objMatrix.Columns.Item("U_DSAT").Cells.Item(RowId).Specific.Value
    '            Dim baseVal As Double = 0
    '            Dim discountVal As Double = 0
    '            Dim badVal As Double = 0
    '            Dim extVal As Double = 0

    '            Double.TryParse(objMatrix.Columns.Item("U_BASE").Cells.Item(RowId).Specific.Value, baseVal)
    '            Double.TryParse(Discount, discountVal)

    '            ' U_BAD = U_BASE - Discount
    '            badVal = baseVal - discountVal
    '            objMatrix.Columns.Item("U_BAD").Cells.Item(RowId).Specific.Value = badVal.ToString("0.00")

    '            ' U_EXT = 50% of U_BAD
    '            extVal = (baseVal * 50) / 100
    '            objMatrix.Columns.Item("U_EXT").Cells.Item(RowId).Specific.Value = extVal.ToString("0.00")

    '            ' Column 14 (Unit Price) = U_BAD + U_EXT
    '            objMatrix.Columns.Item("14").Cells.Item(RowId).Specific.Value = (badVal + extVal).ToString("0.00")

    '            Dim GetSPP1 As String = "SELECT T0.""ItemCode"" AS ""Item No."",
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
    'AND TO_DATE('" & [Date] & "', 'YYYYMMDD') >= T1.""FromDate""
    'AND (
    '        T1.""ToDate"" IS NULL
    '        OR TO_DATE('" & [Date] & "', 'YYYYMMDD') <= T1.""ToDate""
    '    )
    'ORDER BY T0.""CardCode"",T0.""ItemCode"",T1.""FromDate"""

    '            Dim oRsGetSPP1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '            oRsGetSPP1.DoQuery(GetSPP1)


    '            SppPrice = oRsGetSPP1.Fields.Item("Effective Price").Value


    '            '            Dim GetSPP As String = "SELECT T1.""Price"" FROM OITM T0 Left JOIN ITM1 T1 ON T0.""ItemCode"" = T1.""ItemCode""
    '            'WHERE T0.""ItemCode"" = '" & ItemCode & "' AND T1.""PriceList"" = '11'"

    '            Dim GetSPP As String = " SELECT (P.""Price"" * G.""BaseQty"") ""Price""
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

    '            Dim oRsGetSPP As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '            oRsGetSPP.DoQuery(GetSPP)

    '            Dim GetSugar As String = "SELECT T0.""U_SugarTax"" FROM OITM T0 WHERE T0.""ItemCode"" = '" & ItemCode & "'"

    '            Dim oRsGetSugar As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '            oRsGetSugar.DoQuery(GetSugar)

    '            If oRsGetSugar.Fields.Item("U_SugarTax").Value = "Y" Then
    '                objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value = (badVal + extVal).ToString("0.00")
    '                objMatrix.Columns.Item("U_SuppPriBeDis").Cells.Item(RowId).Specific.Value = (badVal + extVal).ToString("0.00")
    '            Else
    '                If SppPrice <> 0 Then
    '                    objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value = SppPrice
    '                    objMatrix.Columns.Item("U_SuppPriBeDis").Cells.Item(RowId).Specific.Value = SppPrice
    '                Else
    '                    objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value = oRsGetSPP.Fields.Item("Price").Value
    '                    objMatrix.Columns.Item("U_SuppPriBeDis").Cells.Item(RowId).Specific.Value = oRsGetSPP.Fields.Item("Price").Value
    '                End If

    '            End If

    '            Dim GetBARANME As String = "Select ""BcdName"" from ""OBCD"" T0 Left Join ""OITM"" T1 ON T0.""ItemCode""=T1.""ItemCode"" where T0.""BcdCode"" = '" & BarCode & "' and T0.""ItemCode"" = '" & ItemCode & "'"

    '            Dim oRsGetBARANME As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '            oRsGetBARANME.DoQuery(GetBARANME)

    '            objMatrix.Columns.Item("U_BRNM").Cells.Item(RowId).Specific.Value = oRsGetBARANME.Fields.Item("BcdName").Value





    '            'CType(objMatrix.Columns.Item("U_Packing").Cells.Item(RowId).Specific, SAPbouiCOM.EditText).Value = Packing

    '            objMatrix.Columns.Item("U_Packing").Cells.Item(RowId).Specific.Value = Packing
    '            objMatrix.Columns.Item("U_RetailPrice").Cells.Item(RowId).Specific.Value = RetailsPrice
    '            'objMatrix.Columns.Item("U_Addbarcode").Cells.Item(RowId).Specific.Value = AdditionalBarCode
    '            objMatrix.Columns.Item("U_LVNDPRC").Cells.Item(RowId).Specific.Value = LeastVendorPrice
    '            objMatrix.Columns.Item("U_LPVCD").Cells.Item(RowId).Specific.Value = LeastVendorCode
    '            objMatrix.Columns.Item("U_LPVNM").Cells.Item(RowId).Specific.Value = LeastVendorName
    '            Dim NetCost As String = objMatrix.Columns.Item("U_suppprice").Cells.Item(RowId).Specific.Value
    '            If NetCost <> "" Then
    '                NetCost = NetCost.Replace(objForm.Items.Item("63").Specific.Value, "")
    '                NetCost = Convert.ToDouble(NetCost.Trim)
    '                Dim LeastPrice As Double = objMatrix.Columns.Item("U_LVNDPRC").Cells.Item(RowId).Specific.Value
    '                If LeastPrice = 0 Then
    '                    Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LPValidate")
    '                    oColumn.Cells.Item(RowId).Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)
    '                End If
    '                If NetCost > 0 And LeastPrice > 0 Then
    '                    If NetCost > LeastPrice Then
    '                        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LPValidate")
    '                        oColumn.Cells.Item(RowId).Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)
    '                    Else
    '                        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_LPValidate")
    '                        oColumn.Cells.Item(RowId).Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue)
    '                    End If
    '                End If
    '                '-------------------Below cost---------------------------
    '                If objMatrix.Columns.Item("14").Cells.Item(RowId).Specific.Value = "" Then
    '                    NetCost = 0
    '                Else
    '                    NetCost = objMatrix.Columns.Item("14").Cells.Item(RowId).Specific.Value
    '                    NetCost = NetCost.Replace(objForm.Items.Item("63").Specific.Value, "")
    '                    NetCost = Convert.ToDouble(NetCost.Trim)
    '                End If
    '                Dim RetailPrice As Double = objMatrix.Columns.Item("U_RetailPrice").Cells.Item(RowId).Specific.Value
    '                If RetailPrice > 0 And NetCost > 0 Then
    '                    If NetCost > RetailPrice Then
    '                        Dim oColumn As SAPbouiCOM.Column = objMatrix.Columns.Item("U_BelowCst")
    '                        oColumn.Cells.Item(RowId).Specific.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
    '                    End If
    '                End If
    '            End If
    '            '-------------------Below cost---------------------------
    '        Catch ex As Exception
    '            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        End Try
    '    End Sub
    Public Function getSupplierPrice_SpecialPrices(ByVal FormUID As String, ByVal IsUomDefault As String,
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

                Dim GetPurchasePriceDfltUOM As String = " Select IFNULL(T3.""Price"",0) ""Price"" from SPP2 T3  " &
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

                    Dim GetPurchasePrice As String = " Select IFNULL(T3.""Price"",0) ""Price"",T3.""UomEntry"" from SPP2 T3  " &
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

                'madhu comment
                'Dim GetPurchasePrice As String = " Select IFNULL(T3.""Price"",0) ""Price"",T3.""UomEntry"" from SPP2 T3  " &
                '" Where T3.""ItemCode"" = '" & ItemCode & "' and T3.""CardCode"" = '" & Vendor & "'  "



                '                Dim oRsGetPurchasePrice As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                '                oRsGetPurchasePrice.DoQuery(GetPurchasePrice)
                '                If oRsGetPurchasePrice.RecordCount > 0 Then
                '                    OtherUOMPrice = CDbl(oRsGetPurchasePrice.Fields.Item(0).Value)
                '                    BaseQty = Me.GetBaseQty(ItemCode, oRsGetPurchasePrice.Fields.Item("UomEntry").Value.ToString())
                '                    UnitCost = OtherUOMPrice / BaseQty
                '                    SuppPrice = UnitCost
                '                End If
                '            End If

                '            Return SuppPrice



                Dim GetPurchasePrice As String = " SELECT IFNULL(T2U.""Price"", T0.""Price"") AS ""SpecialPrice"", " & " IFNULL(T2U.""UomEntry"", -1) AS ""DbUomEntry"" " &
                            " FROM OSPP T0 " & " INNER JOIN OCRD T1 ON T0.""CardCode"" = T1.""CardCode"" " & " LEFT JOIN SPP2 T2U ON T0.""ItemCode"" = T2U.""ItemCode"" " &
                 " AND T0.""CardCode"" = T2U.""CardCode"" " & " WHERE T1.""CardType"" = 'S' " & " AND T0.""ItemCode"" = '" & ItemCode & "' " & " AND T0.""CardCode"" = '" & Vendor & "' "

                Dim oRsGetPurchasePrice As SAPbobsCOM.Recordset =
                objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                oRsGetPurchasePrice.DoQuery(GetPurchasePrice)

                If oRsGetPurchasePrice.RecordCount > 0 Then

                    ' Get price safely
                    If Not IsDBNull(oRsGetPurchasePrice.Fields.Item("SpecialPrice").Value) Then
                        OtherUOMPrice = CDbl(oRsGetPurchasePrice.Fields.Item("SpecialPrice").Value)
                    Else
                        OtherUOMPrice = 0
                    End If

                    ' Use different variable name
                    Dim DbUomEntry As Integer = -1

                    If Not IsDBNull(oRsGetPurchasePrice.Fields.Item("DbUomEntry").Value) Then
                        DbUomEntry = CInt(oRsGetPurchasePrice.Fields.Item("DbUomEntry").Value)
                    End If

                    ' If DB has UoM use that, otherwise use method parameter UomEntry
                    Dim FinalUomEntry As Integer = If(DbUomEntry <> -1, DbUomEntry, UOMEntry)

                    BaseQty = Me.GetBaseQty(ItemCode, FinalUomEntry)

                    If BaseQty > 0 Then
                        UnitCost = OtherUOMPrice / BaseQty
                    Else
                        UnitCost = OtherUOMPrice
                    End If

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


    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("142", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        Try


                            'Me.Items(objForm.UniqueID)
                            'Me.Averagesale()

                            ' 👇 Aging Calculation
                            Me.CalculateAgingQty(oForm)
                            'madhu comment
                            Me.DisabledSupplierPOTotalButton(oForm)
                            If objForm.Items.Item("81").Specific.Value = "3" And BusinessObjectInfo.BeforeAction = False Then
                                objForm.Items.Item("btn_calFOC").Enabled = False
                            End If
                            GlobalDocnum = objForm.Items.Item("8").Specific.Value

                            objForm.Items.Item("112").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                        Catch ex As Exception

                        End Try

                    End If




                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then

                        objMain.objApplication.MessageBox("Purchase Order Document is Successfully Added")
                    End If

            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Private Sub DisabledSupplierPOTotalButton(ByVal oForm As SAPbouiCOM.Form)
        Try

            ' objForm.Items.Item("U_SPOTOTAL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("U_SPOTOTAL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("U_SPOTOTAL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("U_SPOTOTAL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("TUTXAMT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("TUTXAMT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("TUTXAMT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("TUTXAMT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("TEXTAX").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("TEXTAX").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("TEXTAX").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("TEXTAX").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

        Catch ex As Exception
        End Try
    End Sub


    'Sub ButtonFoc(ByVal FormUID As String)
    '    Try
    '        Dim Validation As String = "Select ""U_PO""as ""PO"" from ""@ME_ACFG"""
    '        Dim oLValidation As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oLValidation.DoQuery(Validation)
    '        If oLValidation.Fields.Item("PO").Value = "Y" Then
    '            objForm.Items.Item("btn_calFOC").Enabled = True
    '        Else
    '            objForm.Items.Item("btn_calFOC").Enabled = False
    '        End If
    '    Catch ex As Exception
    '        objForm.Close()
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Dim Validation As String = "Select ""U_PO""as ""PO"" from ""@ME_ACFG"""
                Dim oLValidation As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oLValidation.DoQuery(Validation)
                If oLValidation.Fields.Item("PO").Value = "Y" Then
                    objForm.Items.Item("btn_calFOC").Visible = True
                Else
                    objForm.Items.Item("btn_calFOC").Visible = False
                End If

            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    'Sub CFLFilterSTG(ByVal FormUID As String)
    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)
    '        Dim GetSelection As String = "Select ""ItemCode"" FROM OITM T0 WHERE T0.""CardCode"" = '" & objForm.Items.Item("4").Specific.value & "'"
    '        Dim oRsGetSelectionItems As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oRsGetSelectionItems.DoQuery(GetSelection)

    '        'Dim oConditions As SAPbouiCOM.Conditions
    '        'Dim oCondition As SAPbouiCOM.Condition
    '        'Dim oChooseFromList As SAPbouiCOM.ChooseFromList
    '        'Dim emptyCon As New SAPbouiCOM.Conditions
    '        'oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item("")
    '        'oChooseFromList.SetConditions(emptyCon)
    '        'oConditions = oChooseFromList.GetConditions()
    '        'If oRsGetSelectionItems.RecordCount > 0 Then
    '        '    oCondition = oConditions.Add()
    '        '    oCondition.Alias = "Code"
    '        '    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '        '    oCondition.CondVal = oRsGetSelectionItems.Fields.Item(0).Value
    '        '    oChooseFromList.SetConditions(oConditions)
    '        '    oRsGetSelectionItems.MoveNext()
    '        '    For i As Integer = 1 To oRsGetSelectionItems.RecordCount - 1
    '        '        oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
    '        '        oCondition = oConditions.Add()
    '        '        oCondition.Alias = "Code"
    '        '        oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '        '        oCondition.CondVal = oRsGetSelectionItems.Fields.Item(0).Value
    '        '        oChooseFromList.SetConditions(oConditions)
    '        '        oRsGetSelectionItems.MoveNext()
    '        '    Next
    '        'Else
    '        '    oCondition = oConditions.Add()
    '        '    oCondition.Alias = "Code"
    '        '    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '        '    oCondition.CondVal = ""
    '        '    oChooseFromList.SetConditions(oConditions)
    '        'End If

    '    Catch ex As Exception

    '    End Try
    'End Sub

    Function ValidationnVadlityDate(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)



            If objForm.Items.Item("U_VDATE").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Validity Date Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If



            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Private Sub UpdateHoldingDays(oForm As SAPbouiCOM.Form)

        Try
            If oForm.TypeEx <> "142" Then Exit Sub

            Dim oMatrix As SAPbouiCOM.Matrix = CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix)

            Dim oDBHead As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("OPOR")

            Dim BPLID As String = oDBHead.GetValue("BPLId", 0).Trim()
            Dim DOCSTATUS As String = oDBHead.GetValue("DocStatus", 0).Trim()

            For i As Integer = 1 To oMatrix.VisualRowCount

                Dim itemCode As String = CType(oMatrix.Columns.Item("1").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value

                If itemCode <> "" And DOCSTATUS = "O" Then

                    Dim holdingDays As Double = CalculateHoldingDays_Direct(itemCode, BPLID)
                    objForm.Items.Item("112").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    CType(oMatrix.Columns.Item("U_HDAYS").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value = holdingDays.ToString("0.00")

                End If

            Next

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Private Function CalculateHoldingDays_Direct(itemCode As String, BPLID As String) As Double

        Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim totalQty As Double = 0
        Dim perDay As Double = 0
        Dim avgPrice As Double = 0
        Dim finalValue As Double = 0

        'madhu loading total qty
        'Dim sqlQty As String = "SELECT COALESCE(SUM(T1.""InvQty""),0) AS ""TotalQty"" " &
        '                       "FROM ""OINV"" T0 " & "INNER JOIN ""INV1"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
        '                       "WHERE T1.""ItemCode"" = '" & itemCode & "' " & "AND YEAR(T0.""DocDate"") = YEAR(CURRENT_DATE)"

        'without uom conversation query
        'Dim sqlQty As String = "SELECT COALESCE(SUM(T1.""InvQty""),0) AS ""TotalQty"" " & "FROM ""OINV"" T0 " &
        '                       "INNER JOIN ""INV1"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
        '                      "WHERE T1.""ItemCode"" = '" & itemCode & "' " & "AND T0.""DocDate"" >= ADD_DAYS(CURRENT_DATE, -30)"

        'Uom uom conversation query
        'Dim sqlQty As String = "SELECT COALESCE(SUM(T1.""InvQty"" * G.""BaseQty""),0) AS ""TotalQty"" " &
        '               "FROM ""OINV"" T0 " &
        '               "INNER JOIN ""INV1"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
        '               "INNER JOIN ""OITM"" O ON T1.""ItemCode"" = O.""ItemCode"" " &
        '               "INNER JOIN ""UGP1"" G ON O.""UgpEntry"" = G.""UgpEntry"" AND T1.""UomEntry"" = G.""UomEntry"" " &
        '               "WHERE T1.""ItemCode"" = '" & itemCode & "' " &
        '               "AND T0.""DocDate"" >= ADD_DAYS(CURRENT_DATE, -30)"
        'oRS.DoQuery(sqlQty)


        Dim sql As String = "CALL ""GET_HOLDING_DAYS""('" & itemCode & "', " & BPLID & ")"


        oRS.DoQuery(sql)



        If Not oRS.EoF Then
            Double.TryParse(oRS.Fields.Item("TotalQty").Value.ToString(), totalQty)
        End If

        perDay = totalQty / 30

        Dim sqlAvg As String = "SELECT MAX(T1.""OnHand"") AS ""OnHand"" " & "FROM ""OWHS"" T0 " & "INNER JOIN ""OITW"" T1 ON T0.""WhsCode"" = T1.""WhsCode"" " &
                               "WHERE T1.""ItemCode"" = '" & itemCode & "' " & "AND T0.""BPLid"" = '" & BPLID & "'"

        oRS.DoQuery(sqlAvg)

        If Not oRS.EoF Then
            Double.TryParse(oRS.Fields.Item("OnHand").Value.ToString(), avgPrice)
        End If

        If perDay > 0 Then
            finalValue = avgPrice / perDay
        Else
            finalValue = 0
        End If

        Return Math.Round(finalValue, 2)

    End Function

    'ageing qty calcuation
    Private Sub CalculateAgingQty(oForm As SAPbouiCOM.Form)

        Try
            'Purchase Order Form
            If oForm.TypeEx <> "142" Then Exit Sub

            Dim oMatrix As SAPbouiCOM.Matrix = CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix)
            Dim oDBHead As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("OPOR")

            Dim DOCSTATUS As String = oDBHead.GetValue("DocStatus", 0).Trim()

            For i As Integer = 1 To oMatrix.VisualRowCount

                Dim itemCode As String = CType(oMatrix.Columns.Item("1").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value.Trim()


                'If itemCode <> "" And DOCSTATUS = "O" Then

                '    Dim agingDays As Double = GetAgingDays(itemCode)

                '    CType(oMatrix.Columns.Item("U_AQTY").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value = agingDays.ToString()

                'End If

            Next

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
    Private Function GetAgingDays(itemCode As String) As Double

        Try

            Dim oRS As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim sql As String =
        "SELECT MAX(DAYS_BETWEEN(""DocDate"", CURRENT_DATE)) AS ""AgingDays"" " &
        "FROM ""OINM"" WHERE ""ItemCode"" = '" & itemCode & "'"

            oRS.DoQuery(sql)

            If Not oRS.EoF Then

                If Not IsDBNull(oRS.Fields.Item("AgingDays").Value) Then
                    Return CDbl(oRS.Fields.Item("AgingDays").Value)
                End If

            End If

        Catch
        End Try

        Return 0

    End Function


End Class
