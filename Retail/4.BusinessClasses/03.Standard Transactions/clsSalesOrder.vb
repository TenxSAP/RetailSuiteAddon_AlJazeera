
Public Class clsSalesOrder

#Region "Declaration"
    Public objForm, objGRPOForm As SAPbouiCOM.Form
    Dim objMatrix, objGRPOMatrix As SAPbouiCOM.Matrix
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim BaseUom As String
    Dim completed As Boolean = False
    Dim strQry As String = ""
    Dim oItem, oItem1, oItem2, oItem3 As SAPbouiCOM.Item
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim GlobalDocnum As String
    Dim OrsSTRQRY, rs, rsPO As SAPbobsCOM.Recordset
    Dim Flag As String = "No"
    Dim FormID As String
    Dim ButtonCalFlag As Boolean

#End Region
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix = objForm.Items.Item("38").Specific
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                    (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                    End If

                    Dim objCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("3").Specific
                    Dim TYPE As String = objCombo.Value

                    If pVal.ItemUID = "CLEX" And pVal.BeforeAction = False And (pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE) And TYPE = "I" Then
                        Try
                            objForm.Freeze(True)

                            ButtonCalFlag = True


                            For R As Integer = 1 To objMatrix.VisualRowCount

                                If objMatrix.Columns.Item("1").Cells.Item(R).Specific.Value <> "" Then


                                    Dim GetPerPCVal As Double = 0

                                    Dim GetPerPC As String = "SELECT ""BaseQty""  FROM  ""OUGP""  T0 INNER JOIN  ""UGP1""  T1 ON T0.""UgpEntry"" = T1.""UgpEntry"" " &
                                                            " INNER JOIN  ""OITM""  T2 ON T0.""UgpEntry"" = T2.""UgpEntry"" WHERE T2.""ItemCode"" ='" & CStr(objMatrix.Columns.Item("1").Cells.Item(R).Specific.Value).Trim & "' " &
                                                            " And ""UomEntry"" IN (SELECT T0.""UomEntry"" FROM OUOM T0 WHERE T0.""UomCode"" ='" & CStr(objMatrix.Columns.Item("1470002145").Cells.Item(R).Specific.Value).Trim & "') "

                                    Dim oRsGetPerPCas As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRsGetPerPCas.DoQuery(GetPerPC)

                                    If oRsGetPerPCas.RecordCount > 0 Then
                                        GetPerPCVal = oRsGetPerPCas.Fields.Item(0).Value
                                    End If


                                    Dim Discount As Double = 0
                                    Dim Total As Double = 0


                                    If CStr(objMatrix.Columns.Item("15").Cells.Item(R).Specific.Value).Trim <> "" Then
                                        Discount = CDbl(CStr(objMatrix.Columns.Item("15").Cells.Item(R).Specific.Value).Trim)
                                    End If


                                    If CStr(objMatrix.Columns.Item("21").Cells.Item(R).Specific.Value).Trim <> "" Then
                                        Total = String.Concat(CStr(objMatrix.Columns.Item("21").Cells.Item(R).Specific.Value).Trim.Where(Function(c) Char.IsDigit(c))) 'CDbl(CStr(objMatrix.Columns.Item("21").Cells.Item(R).Specific.Value).Trim)
                                    End If


                                    ' If Discount = 100 Or Total = 0 Then
                                    'If Total = 0 Then
                                    '    Dim EXPNM As String = " Select ""ExpnsCode"" from OEXD  where ""ExpnsName"" = 'Local Excise' "
                                    '    Dim oRsEXPNM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    '    oRsEXPNM.DoQuery(EXPNM)

                                    '    Try
                                    '        Dim ExpnsCode1 As String = oRsEXPNM.Fields.Item("ExpnsCode").Value

                                    '        Dim objComboBox8 As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("111").Cells.Item(R).Specific
                                    '        Dim TaskName8 As String = objComboBox8.Selected.Value

                                    '        objComboBox8.Select("-1", SAPbouiCOM.BoSearchKey.psk_ByValue)

                                    '        objMatrix.Columns.Item("U_EXCTAX").Cells.Item(R).Specific.Value() = 0
                                    '        objMatrix.Columns.Item("112").Cells.Item(R).Specific.Value = 0


                                    '    Catch ex As Exception

                                    '    End Try

                                    'Else

                                    Dim cardcode As String = objForm.Items.Item("4").Specific.Value
                                    Dim Getdata As String = "SELECT T0.""U_BPTYPE"",T0.""ExpnsCode"",T0.""ExpnsName"" FROM OEXD T0 inner join NNM1 T1 ON T1.""Series""=T0.""U_BPSERIES"" " +
                                                                    "  inner join OCRD T2 on T2.""Series""=T1.""Series"" where  T2.""CardCode""='" & cardcode & "'"
                                    Dim oRs1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRs1.DoQuery(Getdata)
                                    Dim ExpnsCode As String = oRs1.Fields.Item("ExpnsCode").Value
                                    Dim BPType As String = oRs1.Fields.Item("U_BPTYPE").Value

                                    If BPType = "LOCAL" Then 'And (customertype = "CUS" Or customertype = "Manual") Then
                                        Dim objComboBox3 As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("111").Cells.Item(R).Specific
                                        Dim TaskName3 As String = objComboBox3.Selected.Value
                                        objComboBox3.Select(ExpnsCode, SAPbouiCOM.BoSearchKey.psk_ByValue)

                                        objMatrix = objForm.Items.Item("38").Specific
                                        Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(R).Specific.Value
                                        Dim GetDocEntry As String = "Select ""U_EXCVAT"" From OITM Where ""ItemCode"" = '" & ItemCode & "' "
                                        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRs.DoQuery(GetDocEntry)

                                        objMatrix.Columns.Item("U_EXCTAX").Cells.Item(R).Specific.Value() = oRs.Fields.Item("U_EXCVAT").Value()

                                        Dim ItemCode1 As String = objMatrix.Columns.Item("1").Cells.Item(R).Specific.Value.ToString().Trim()
                                        Dim QTY As Double = objMatrix.Columns.Item("11").Cells.Item(R).Specific.Value.ToString().Trim()
                                        Dim EXCTAX As Double = objMatrix.Columns.Item("U_EXCTAX").Cells.Item(R).Specific.Value.ToString().Trim() * GetPerPCVal
                                        Dim SubTotal As String = "ST1"
                                        If (QTY * EXCTAX) <> 0 Then
                                            Dim ST1 As Double = QTY * EXCTAX
                                            SubTotal = ST1

                                            If objMatrix.Columns.Item("U_EXCTAX").Cells.Item(R).Specific.Value <> "" Then
                                                Dim objComboBox As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("111").Cells.Item(R).Specific
                                                Dim TaskName As String = objComboBox.Selected.Value
                                                ' objComboBox.Select("Sales Excise Duty", SAPbouiCOM.BoSearchKey.psk_ByDescription)
                                                objComboBox.Select(ExpnsCode, SAPbouiCOM.BoSearchKey.psk_ByValue)

                                                Dim objComboBox1 As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("111").Cells.Item(R).Specific
                                                Dim TaskName1 As String = objComboBox.Selected.Value
                                                If objMatrix.Columns.Item("111").Cells.Item(R).Specific.Value = TaskName1 Then
                                                    objMatrix.Columns.Item("112").Cells.Item(R).Specific.Value = SubTotal
                                                End If

                                            End If
                                        Else
                                            Dim objComboBox2 As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("111").Cells.Item(R).Specific
                                            Dim TaskName As String = objComboBox2.Selected.Value
                                            objComboBox2.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                        End If

                                    Else
                                        Dim EXPNM As String = " Select ""ExpnsCode"" from OEXD  where ""U_BPTYPE"" = 'FOREIGN' "
                                        Dim oRsEXPNM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRsEXPNM.DoQuery(EXPNM)
                                        Dim ExpnsCode1 As String = oRsEXPNM.Fields.Item("ExpnsCode").Value

                                        Dim objComboBox8 As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("111").Cells.Item(R).Specific
                                        Dim TaskName8 As String = objComboBox8.Selected.Value
                                        objComboBox8.Select(ExpnsCode, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                        objMatrix = objForm.Items.Item("38").Specific
                                        Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(R).Specific.Value
                                        Dim GetDocEntry As String = "Select ""U_EXCVAT"" From OITM Where ""ItemCode"" = '" & ItemCode & "' "
                                        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRs.DoQuery(GetDocEntry)
                                        objMatrix.Columns.Item("U_EXCTAX").Cells.Item(R).Specific.Value() = oRs.Fields.Item("U_EXCVAT").Value()
                                    End If


                                End If


                                '  End If

                            Next


                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If



                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And TYPE = "I" Then
                    '    Try
                    '        If ButtonCalFlag = False Then
                    '            objMain.objApplication.StatusBar.SetText("Please click Calculate Excise button before adding")
                    '            BubbleEvent = False
                    '        End If
                    '    Catch ex As Exception

                    '    End Try
                    'End If

                    If pVal.ItemUID = "btn_Filter" And pVal.BeforeAction = False Then
                        Try
                            Dim CardCode As String = objForm.Items.Item("4").Specific.Value.ToString()
                            If CardCode <> "" Then
                                objMatrix = objForm.Items.Item("38").Specific
                                Dim rSet3 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim qStr3 As String = "DELETE FROM ""@TENX_TEMP2"""
                                rSet3.DoQuery(qStr3)

                                Dim rSet2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim qStr2 As String = "DELETE FROM ""@TENX_TEMP1"""
                                rSet2.DoQuery(qStr2)

                                Dim rSet4 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim qStr4 As String = "SELECT ROW_NUMBER() OVER (ORDER BY C.""ItemCode"") ""Code"",ROW_NUMBER() OVER (ORDER BY C.""ItemCode"") ""Name"",'N' ""Check"",C.""ItemCode"",C.""ItemName"",A.""BcdCode"" as ""BarCode"",B.""UomCode"",C.""OnHand"" " +
                                                   " FROM OITM C LEFT JOIN OBCD A ON A.""ItemCode""=C.""ItemCode"" LEFT JOIN OUOM B ON B.""UomEntry""=A.""UomEntry"" WHERE C.""validFor""='Y' AND C.""SellItem""='Y'"

                                Dim rSet5 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim qStr5 As String = "INSERT INTO ""@TENX_TEMP1"" " & qStr4.ToString().ToString() & " "
                                rSet5.DoQuery(qStr5)

                                rSet4.DoQuery(qStr4)
                                If rSet4.RecordCount > 0 Then
                                    objMain.objSearchAndFilter.CreateForm(qStr4, objForm)
                                Else
                                    objMain.objApplication.StatusBar.SetText("Data not found !!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    objForm.Freeze(False)
                                    BubbleEvent = False
                                End If
                            Else
                                objMain.objApplication.StatusBar.SetText("Before searching for an item, you need to select the customer.!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If
                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If


                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.BeforeAction = False And pVal.ItemUID = "10000330" Then
                        Dim oComboBox7 As SAPbouiCOM.ComboBox = objForm.Items.Item("10000330").Specific
                        Dim obj As String = oComboBox7.Selected.Description
                        objForm = objMain.objApplication.Forms.Item(FormUID)
                        objMatrix = objForm.Items.Item("38").Specific
                        If obj = "Purchase Order" Then
                            objMain.IsFormClosed = False
                            objMain.objDocumentDraw.CreateForm(objForm.UniqueID)
                            '   AddHandler objMain.objApplication.ItemEvent, AddressOf HandleFormClose

                            'If objMain.PODocnumentNumber <> "" Then
                            '    Dim qry As String = "call ""TNX_LOADPO""('" & objMain.PODocnumentNumber & "')"
                            '    Dim oRsGetSTRLineDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            '    oRsGetSTRLineDetails.DoQuery(qry)
                            '    If oRsGetSTRLineDetails.RecordCount > 0 Then
                            '        For i As Integer = 1 To oRsGetSTRLineDetails.RecordCount
                            '            Dim ItemCode As String = oRsGetSTRLineDetails.Fields.Item("ItemCode").Value.ToString().Trim()
                            '            Dim Quantity As Double = CDbl(oRsGetSTRLineDetails.Fields.Item("Quantity").Value)
                            '            Dim UOM As String = oRsGetSTRLineDetails.Fields.Item("UomCode").Value.ToString().Trim()
                            '            Dim Barcode As String = oRsGetSTRLineDetails.Fields.Item("CodeBars").Value.ToString().Trim()
                            '            Dim Price As Double = CDbl(oRsGetSTRLineDetails.Fields.Item("Price").Value)
                            '            Dim line As String = oRsGetSTRLineDetails.Fields.Item("LineNum").Value.ToString().Trim()
                            '            Dim docnum As String = oRsGetSTRLineDetails.Fields.Item("DocNum").Value.ToString().Trim()
                            '            Dim docEntry As String = oRsGetSTRLineDetails.Fields.Item("DocEntry").Value.ToString().Trim()

                            '            If i > objMatrix.RowCount Then objMatrix.AddRow()


                            '            ' Set Barcode or Item Code based on Barcode value

                            '            If String.IsNullOrEmpty(Barcode) And String.IsNullOrEmpty(ItemCode) Then
                            '                objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value = ItemCode

                            '                objMatrix.Columns.Item("4").Cells.Item(i).Specific.Value = Barcode

                            '            Else
                            '                objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value = ItemCode

                            '            End If

                            '            ' Set UOM or default combobox selection
                            '            Dim combobox As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("12").Cells.Item(i).Specific
                            '            If String.IsNullOrEmpty(UOM) Then
                            '                'combobox.SelectExclusive("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            '            Else
                            '                combobox.SelectExclusive("N", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            '                objMatrix.Columns.Item("1470002177").Cells.Item(i).Specific.Value = UOM
                            '            End If

                            '            ' Set Quantity, with a default of 0 if Quantity is 0
                            '            objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value = If(Quantity = 0, 0, Quantity)
                            '            objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value = If(Price = 0, 0, Price)


                            '            objMatrix.Columns.Item("U_MIDocNum").Cells.Item(i).Specific.Value = If(docnum = "", "", docnum)
                            '            objMatrix.Columns.Item("U_MIEntry").Cells.Item(i).Specific.Value = If(docEntry = "", "", docEntry)
                            '            objMatrix.Columns.Item("U_ItemRef").Cells.Item(i).Specific.Value = If(line = "", "", line)

                            '            objMatrix.AddRow()

                            '            oRsGetSTRLineDetails.MoveNext()
                            '        Next
                            '    End If

                            'End If

                        End If
                    End If

                    If pVal.ItemUID = "btn_Filter" And pVal.BeforeAction = False Then
                        Try
                            '    objMatrix = objForm.Items.Item("38").Specific
                            'Dim rSet3 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'Dim qStr3 As String = "DELETE FROM ""@TENX_TEMP2"""
                            'rSet3.DoQuery(qStr3)

                            'Dim rSet2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'Dim qStr2 As String = "DELETE FROM ""@TENX_TEMP1"""
                            'rSet2.DoQuery(qStr2)

                            'Dim rSet4 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            '    Dim qStr4 As String = "SELECT  ""ItemCode"",""ItemCode"",""ItemCode"",""ItemName"",""CodeBars"" FROM OITM WHERE ""validFor""='Y' AND ""SellItem""='Y'"

                            '    Dim rSet5 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'Dim qStr5 As String = "INSERT INTO ""@TENX_TEMP1"" " & qStr4.ToString().ToString() & " "
                            'rSet5.DoQuery(qStr5)

                            'rSet4.DoQuery(qStr4)
                            '    If rSet4.RecordCount > 0 Then
                            '        objMain.objSearchAndFilter.CreateForm(qStr4)
                            '    Else
                            '        objMain.objApplication.StatusBar.SetText("Data not found !!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            '        objForm.Freeze(False)
                            '        BubbleEvent = False
                            '    End If

                        Catch ex As Exception
                            objForm.Freeze(False)
                        End Try
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    FormID = FormUID
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("ORDR")
                    objMatrix = objForm.Items.Item("38").Specific
                    If pVal.ItemUID = "4" And pVal.BeforeAction = True Then
                        Dim cardcode As String = objForm.Items.Item("4").Specific.value
                        '   If pVal.BeforeAction = False And cardcode <> "" Then
                        ''Me.AddItems(objForm.UniqueID)
                        objForm = objMain.objApplication.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                        Dim oComboBox As SAPbouiCOM.ComboBox

                        oComboBox = CType(objForm.Items.Item("10000330").Specific, SAPbouiCOM.ComboBox)
                        oComboBox.ValidValues.Add("Purchase Order", "Purchase Order")


                        '  End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    ' Check if the second form was closed and the flag is set
                    If objMain.IsFormClosed And objMain.PODocnumentNumber <> "" Then
                        ' Call logic to add data to the Sales Order matrix
                        completed = False
                        If completed <> True Then
                            AddDataToSalesOrder(objMain.PODocnumentNumber)

                        End If

                        ' Reset the flag after processing
                        objMain.IsFormClosed = False
                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    FormID = FormUID
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("ORDR")
                    objMatrix = objForm.Items.Item("38").Specific
                    Dim cardcode As String = objForm.Items.Item("4").Specific.value
                    If pVal.BeforeAction = False And cardcode <> "" Then
                        ''Me.AddItems(objForm.UniqueID)
                        objForm = objMain.objApplication.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                        Dim oComboBox As SAPbouiCOM.ComboBox

                        oComboBox = CType(objForm.Items.Item("10000330").Specific, SAPbouiCOM.ComboBox)
                        oComboBox.ValidValues.Add("Purchase Order", "Purchase Order")

                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType


                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        ButtonCalFlag = False
                    End If
            End Select

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Private Sub AddItems(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            oItem1 = objForm.Items.Add("btn_Filter", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem1.Top = objForm.Items.Item("3").Top
            oItem1.Left = objForm.Items.Item("3").Left + 80
            oItem1.Width = objForm.Items.Item("3").Width
            oItem1.Height = objForm.Items.Item("3").Height
            oItem1.FromPane = objForm.Items.Item("3").FromPane
            oItem1.ToPane = objForm.Items.Item("3").ToPane
            Dim btn1 As SAPbouiCOM.Button = objForm.Items.Item("btn_Filter").Specific
            btn1.Caption = "Filter & Load"

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Sub GetValueFromSearchAndFilter()
        Try
            '' objForm = objMain.objApplication.Forms.Item(FormID)
            objForm = objMain.objApplication.Forms.ActiveForm
            objMatrix = objForm.Items.Item("38").Specific
            objForm.Freeze(True)
            Dim RSet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim qStr As String = "SELECT ""U_COLUMN1"" FROM ""@TENX_TEMP2"" "
            RSet.DoQuery(qStr)
            If RSet.RecordCount > 0 Then
                While Not RSet.EoF
                    objMatrix.Columns.Item("1").Cells.Item(objMatrix.VisualRowCount).Specific.Value = RSet.Fields.Item("U_COLUMN1").Value.ToString()
                    RSet.MoveNext()
                End While
            End If
            objMain.objApplication.StatusBar.SetText("Data Loaded Successfully !!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
        End Try
    End Sub

    Public Sub GetValueFromSearchAndFilter1(ByRef oForm As SAPbouiCOM.Form)
        Try
            '' objForm = objMain.objApplication.Forms.Item(FormID)
            oForm = objMain.objApplication.Forms.ActiveForm
            objMatrix = oForm.Items.Item("38").Specific
            oForm.Freeze(True)
            Dim RSet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim qStr As String = "SELECT ""U_COLUMN1"",""U_COLUMN3""  FROM ""@TENX_TEMP2"" "
            RSet.DoQuery(qStr)
            If RSet.RecordCount > 0 Then
                While Not RSet.EoF
                    Dim i As Integer = objMatrix.VisualRowCount
                    objMain.objApplication.StatusBar.SetText("Item is loading in queue for line number " + i.ToString() + ", please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value = RSet.Fields.Item("U_COLUMN1").Value.ToString()
                    objMatrix.Columns.Item("4").Cells.Item(i).Specific.Value = RSet.Fields.Item("U_COLUMN3").Value.ToString()
                    RSet.MoveNext()
                End While
            End If
            objMain.objApplication.StatusBar.SetText("Item Loaded Successfully !!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objMatrix.Columns.Item("1").Cells.Item(objMatrix.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub HandleFormClose(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        ' Check if the closed form is the one we opened
        If pVal.FormUID = FormUID AndAlso pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
            ' Remove the event handler to prevent multiple triggers
            RemoveHandler objMain.objApplication.ItemEvent, AddressOf HandleFormClose

            ' Now fetch data and populate the matrix
            If Not String.IsNullOrEmpty(objMain.PODocnumentNumber) Then
                Dim qry As String = "CALL ""TNX_LOADPO""('" & objMain.PODocnumentNumber & "')"
                Dim oRsGetSTRLineDetails As SAPbobsCOM.Recordset
                oRsGetSTRLineDetails = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetSTRLineDetails.DoQuery(qry)

                If oRsGetSTRLineDetails.RecordCount > 0 Then
                    '    PopulateMatrix(oRsGetSTRLineDetails)
                End If
            End If
        End If
    End Sub

    Private Sub PopulateMatrix(ByRef oRsGetSTRLineDetails As SAPbobsCOM.Recordset)
        For i As Integer = 1 To oRsGetSTRLineDetails.RecordCount
            Dim ItemCode As String = oRsGetSTRLineDetails.Fields.Item("ItemCode").Value.ToString().Trim()
            Dim Quantity As Double = CDbl(oRsGetSTRLineDetails.Fields.Item("Quantity").Value)
            Dim UOM As String = oRsGetSTRLineDetails.Fields.Item("UomCode").Value.ToString().Trim()
            Dim Barcode As String = oRsGetSTRLineDetails.Fields.Item("CodeBars").Value.ToString().Trim()
            Dim Price As Double = CDbl(oRsGetSTRLineDetails.Fields.Item("Price").Value)
            Dim line As String = oRsGetSTRLineDetails.Fields.Item("LineNum").Value.ToString().Trim()
            Dim docnum As String = oRsGetSTRLineDetails.Fields.Item("DocNum").Value.ToString().Trim()
            Dim docEntry As String = oRsGetSTRLineDetails.Fields.Item("DocEntry").Value.ToString().Trim()

            If i > objMatrix.RowCount Then objMatrix.AddRow()

            ' Set Barcode or Item Code based on Barcode value
            If String.IsNullOrEmpty(Barcode) And String.IsNullOrEmpty(ItemCode) Then
                objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value = ItemCode
                objMatrix.Columns.Item("4").Cells.Item(i).Specific.Value = Barcode
            Else
                objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value = ItemCode
            End If

            ' Set UOM or default combobox selection
            Dim combobox As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("12").Cells.Item(i).Specific
            If String.IsNullOrEmpty(UOM) Then
                'combobox.SelectExclusive("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)
            Else
                combobox.SelectExclusive("N", SAPbouiCOM.BoSearchKey.psk_ByValue)
                objMatrix.Columns.Item("1470002177").Cells.Item(i).Specific.Value = UOM
            End If

            ' Set Quantity & Price
            objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value = If(Quantity = 0, 0, Quantity)
            objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value = If(Price = 0, 0, Price)

            ' Set additional fields
            objMatrix.Columns.Item("U_MIDocNum").Cells.Item(i).Specific.Value = If(docnum = "", "", docnum)
            objMatrix.Columns.Item("U_MIEntry").Cells.Item(i).Specific.Value = If(docEntry = "", "", docEntry)
            objMatrix.Columns.Item("U_ItemRef").Cells.Item(i).Specific.Value = If(line = "", "", line)

            objMatrix.AddRow()
            oRsGetSTRLineDetails.MoveNext()
        Next
    End Sub




    Public Sub AddDataToSalesOrder(ByVal poNumber As String)
        Try
            If objMain.PODocnumentNumber <> "" And objMain.IsFormClosed = True Then




                Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.ActiveForm
                Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("38").Specific

                ' Query Data
                Dim qry As String = "CALL ""TNX_LOADPO""('" & poNumber & "')"
                Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRs.DoQuery(qry)

                If oRs.RecordCount > 0 Then
                    Dim docnum As String = oRs.Fields.Item("DocNum").Value.ToString().Trim()

                    objForm.Items.Item("14").Specific.value = docnum
                    Dim test As Integer = oRs.RecordCount
                    For i As Integer = 1 To oRs.RecordCount
                        Dim ItemCode As String = oRs.Fields.Item("ItemCode").Value.ToString().Trim()
                        Dim Quantity As Double = CDbl(oRs.Fields.Item("Quantity").Value)
                        Dim UOM As String = oRs.Fields.Item("UomCode").Value.ToString().Trim()
                        Dim Barcode As String = oRs.Fields.Item("CodeBars").Value.ToString().Trim()
                        Dim Price As Double = CDbl(oRs.Fields.Item("Price").Value)
                        Dim LineNum As String = oRs.Fields.Item("LineNum").Value.ToString().Trim()
                        Dim line As String = oRs.Fields.Item("LineNum").Value.ToString().Trim()
                        docnum = oRs.Fields.Item("DocNum").Value.ToString().Trim()
                        Dim docEntry As String = oRs.Fields.Item("DocEntry").Value.ToString().Trim()

                        ' Add row if necessary
                        'if i > objmatrix.RowCount Then objMatrix.AddRow()

                        ' Set values in the matrix

                        objMatrix.Columns.Item("4").Cells.Item(i).Specific.Value = Barcode
                        '     objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value = If(String.IsNullOrEmpty(ItemCode), "", ItemCode)

                        'Dim combobox As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("12").Cells.Item(i).Specific
                        'If String.IsNullOrEmpty(UOM) Then
                        '    'combobox.SelectExclusive("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        'Else
                        '    combobox.SelectExclusive("N", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        '    objMatrix.Columns.Item("1470002177").Cells.Item(i).Specific.Value = UOM
                        'End If

                        objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value = If(Quantity = 0, 1, Quantity)
                        objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value = If(Price = 0, 0, Price)
                        objMatrix.Columns.Item("U_ItemRef").Cells.Item(i).Specific.Value = If(LineNum = "", "", LineNum)
                        objMatrix.Columns.Item("U_MIDocNum").Cells.Item(i).Specific.Value = If(docnum = "", "", docnum)
                        objMatrix.Columns.Item("U_MIEntry").Cells.Item(i).Specific.Value = If(docEntry = "", "", docEntry)
                        '       objMatrix.Columns.Item("U_ItemRef").Cells.Item(i).Specific.Value = If(line = "", "", line)
                        If Barcode <> "" Then
                            objMatrix.AddRow()
                        End If

                        oRs.MoveNext()

                    Next
                    completed = True
                    objMain.PODocnumentNumber = ""
                    objMain.IsFormClosed = False
                End If
            End If
        Catch ex As Exception
            objMain.PODocnumentNumber = ""
            objMain.IsFormClosed = False
            completed = True
        End Try
    End Sub



End Class
