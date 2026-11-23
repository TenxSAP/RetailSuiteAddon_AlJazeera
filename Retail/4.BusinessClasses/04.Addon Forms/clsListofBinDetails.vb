Public Class clsListofBinDetails

#Region "        Declaration        "
    Dim objForm, PrICTForm As SAPbouiCOM.Form
    Dim PrICTMatrix As SAPbouiCOM.Matrix
    Dim oGrid As SAPbouiCOM.Grid
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim oDt As SAPbouiCOM.DataTable
    Dim oEditTextCol As SAPbouiCOM.EditTextColumn
    Dim oDBs_Head, oDBs_Details, PrICT_oDBs_Details As SAPbouiCOM.DBDataSource
    Public sItemCode As String
    Public dQuantity As Double
    Public sWarehouse As String
    Public sBuyerDB As String
    Public sTransactionType As String
    Public iPrLineNum As Integer
    Public sBaseFormUID As String
    Public sSellerDB As String
    Dim oEditText As SAPbouiCOM.EditText
#End Region

    Sub CreateForm(ByVal BaseFormUID As String, ByVal ItemCode As String, ByVal Quantity As Double, _
                   ByVal Warehouse As String, ByVal BuyerDB As String, ByVal SellerDB As String, ByVal TransactionType As String, ByVal PrLineNum As Integer)
        Try
            objMain.objUtilities.LoadForm("ListofBinDetails.xml", "ME_LBD_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_LBD_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)

            oDt = objForm.DataSources.DataTables.Add("dt1")
            oDt = objForm.DataSources.DataTables.Item("dt1")

            Me.sBaseFormUID = BaseFormUID
            Me.sItemCode = ItemCode
            Me.dQuantity = Quantity
            Me.sWarehouse = Warehouse
            Me.sBuyerDB = BuyerDB
            Me.sTransactionType = TransactionType
            Me.iPrLineNum = PrLineNum
            Me.sSellerDB = SellerDB
            If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = False Then
                Retail.MainCls.ohtLookUpForm.Add(objForm.UniqueID, BaseFormUID)
            End If

            objForm.DataSources.UserDataSources.Add("U_MEDB", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_MEWHS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_MEITM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_MEQTY", SAPbouiCOM.BoDataType.dt_QUANTITY)

            oEditText = objForm.Items.Item("5").Specific
            oEditText.DataBind.SetBound(True, "", "U_MEDB")
            objForm.Items.Item("5").AffectsFormMode = False
            If sTransactionType = "GRPO" Then
                objForm.DataSources.UserDataSources.Item("U_MEDB").Value = sBuyerDB
            ElseIf sTransactionType = "Delivery" Then
                objForm.DataSources.UserDataSources.Item("U_MEDB").Value = sSellerDB
            End If

            oEditText = objForm.Items.Item("7").Specific
            oEditText.DataBind.SetBound(True, "", "U_MEWHS")
            objForm.Items.Item("7").AffectsFormMode = False
            objForm.DataSources.UserDataSources.Item("U_MEWHS").Value = sWarehouse

            oEditText = objForm.Items.Item("9").Specific
            oEditText.DataBind.SetBound(True, "", "U_MEITM")
            objForm.Items.Item("9").AffectsFormMode = False
            objForm.DataSources.UserDataSources.Item("U_MEITM").Value = sItemCode

            oEditText = objForm.Items.Item("11").Specific
            oEditText.DataBind.SetBound(True, "", "U_MEQTY")
            objForm.Items.Item("11").AffectsFormMode = False
            objForm.DataSources.UserDataSources.Item("U_MEQTY").Value = dQuantity

            Me.LoadGrid(objForm.UniqueID, ItemCode, dQuantity, sWarehouse, BuyerDB, TransactionType)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub LoadGrid(ByVal FormUID As String, ByVal ItemCode As String, ByVal Quantity As Double, _
                   ByVal Warehouse As String, ByVal BuyerDB As String, ByVal TransactionType As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oGrid = objForm.Items.Item("3").Specific
            oGrid.DataTable = oDt
            oGrid.DataTable.Clear()
            Dim GetBinDetails As String
            If TransactionType = "GRPO" Then
                GetBinDetails = "Select 'N' 'Check' , BinCode , Convert(Float,0) 'Qty' , AbsEntry  From OBIN Where WhsCode ='" & Warehouse & "'"
                oGrid.DataTable.ExecuteQuery(GetBinDetails)
            End If

            If TransactionType = "Delivery" Then
                GetBinDetails = " Select 'N' 'Check' , (Select BinCode  From [" & sSellerDB & "].[dbo].OBIN Where AbsEntry = BinAbs) 'BinCode' , OnHandQty 'AvailableQty' , Convert(Float,0) 'Qty' , BinAbs As 'AbsEntry' From  [" & sSellerDB & "].[dbo].OIBQ " & _
                                " Where ItemCode = '" & ItemCode & "' And WhsCode = '" & Warehouse & "' "
                oGrid.DataTable.ExecuteQuery(GetBinDetails)
            End If

            oGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item("Check").Editable = True
            oGrid.Columns.Item("BinCode").Editable = False
            If TransactionType = "Delivery" Then
                oGrid.Columns.Item("AvailableQty").Editable = False
            End If
            oGrid.Columns.Item("AbsEntry").Visible = False
            oGrid.AutoResizeColumns()
            oGrid.Columns.Item("Qty").Click(0)

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    If pVal.BeforeAction = False Then
                        If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = True Then
                            Retail.MainCls.ohtLookUpForm.Remove(objForm.UniqueID)
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oGrid = objForm.Items.Item("3").Specific
                    If pVal.ItemUID = "4" And pVal.BeforeAction = False Then

                        If sTransactionType = "GRPO" Then
                            For k As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                                If oGrid.DataTable.GetValue("Check", k) = "Y" And oGrid.DataTable.GetValue("Qty", k) = 0 Then
                                    objMain.objApplication.StatusBar.SetText("Please enter the Selected Quantity for the Bin - " & oGrid.DataTable.GetValue("BinCode", k))
                                    Exit Sub
                                End If
                            Next
                        End If

                        If sTransactionType = "Delivery" Then
                            For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                                If oGrid.DataTable.GetValue("Check", i) = "Y" Then

                                    If oGrid.DataTable.GetValue("Check", i) = "Y" And oGrid.DataTable.GetValue("Qty", i) <> 0 Then
                                        If (CDbl(oGrid.DataTable.GetValue("Qty", i)) > CDbl(oGrid.DataTable.GetValue("AvailableQty", i))) Then
                                            objMain.objApplication.StatusBar.SetText("Selected Quantity cannot be greater than Available Quantity")
                                            oGrid.DataTable.SetValue("Check", i, "N")
                                            BubbleEvent = False
                                            Exit Select
                                        End If
                                    End If
                                End If
                            Next
                        End If

                        '---------GetTotlaQty----------
                        Dim TotalSelcQty As Double = 0
                        For j As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                            If oGrid.DataTable.GetValue("Check", j) = "Y" Then
                                If CDbl(oGrid.DataTable.GetValue("Qty", j)) <> 0 Then
                                    TotalSelcQty = TotalSelcQty + CDbl(oGrid.DataTable.GetValue("Qty", j))
                                End If
                            End If
                        Next
                        If TotalSelcQty = "0" Then
                            objMain.objApplication.StatusBar.SetText("Total Selected Quantity cannot be Zero")
                            BubbleEvent = False
                            Exit Select
                        End If
                        If TotalSelcQty > CDbl(objForm.Items.Item("11").Specific.Value) Then
                            objMain.objApplication.StatusBar.SetText("Total Selected Quantity cannot be greater than Item Quantity")
                            BubbleEvent = False
                            Exit Select
                        End If

                        PrICTForm = objMain.objApplication.Forms.Item(sBaseFormUID)
                        PrICTMatrix = PrICTForm.Items.Item("21").Specific
                        PrICT_oDBs_Details = PrICTForm.DataSources.DBDataSources.Item("@SaudiaPROC_C0")

                        Dim Bin_Qty As String = ""
                        For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                            If oGrid.DataTable.GetValue("Check", i) = "Y" Then
                                Bin_Qty = Bin_Qty & oGrid.DataTable.GetValue("AbsEntry", i) & "|" & oGrid.DataTable.GetValue("Qty", i) & ";"
                            End If
                        Next

                        If Bin_Qty.Trim <> "" Then
                            'Bin_Qty = Bin_Qty.Remove(Bin_Qty.LastIndexOf(";"))
                            PrICT_oDBs_Details.SetValue("LineId", PrICT_oDBs_Details.Offset, iPrLineNum)
                            PrICT_oDBs_Details.SetValue("U_MESELCT", PrICT_oDBs_Details.Offset, "Y")
                            PrICT_oDBs_Details.SetValue("U_MEITMCD", PrICT_oDBs_Details.Offset, PrICTMatrix.Columns.Item("V_0").Cells.Item(iPrLineNum).Specific.Value)
                            PrICT_oDBs_Details.SetValue("U_MEITMNM", PrICT_oDBs_Details.Offset, PrICTMatrix.Columns.Item("V_1").Cells.Item(iPrLineNum).Specific.Value)
                            PrICT_oDBs_Details.SetValue("U_MEHSN", PrICT_oDBs_Details.Offset, PrICTMatrix.Columns.Item("V_2").Cells.Item(iPrLineNum).Specific.Value)
                            PrICT_oDBs_Details.SetValue("U_MEQTY", PrICT_oDBs_Details.Offset, PrICTMatrix.Columns.Item("V_3").Cells.Item(iPrLineNum).Specific.Value)
                            PrICT_oDBs_Details.SetValue("U_MEINSTK", PrICT_oDBs_Details.Offset, PrICTMatrix.Columns.Item("V_4").Cells.Item(iPrLineNum).Specific.Value)
                            PrICT_oDBs_Details.SetValue("U_MEPBD", PrICT_oDBs_Details.Offset, PrICTMatrix.Columns.Item("V_5").Cells.Item(iPrLineNum).Specific.Value)
                            PrICT_oDBs_Details.SetValue("U_MEPAD", PrICT_oDBs_Details.Offset, PrICTMatrix.Columns.Item("V_6").Cells.Item(iPrLineNum).Specific.Value)
                            PrICT_oDBs_Details.SetValue("U_METAX", PrICT_oDBs_Details.Offset, PrICTMatrix.Columns.Item("V_7").Cells.Item(iPrLineNum).Specific.Value)
                            PrICT_oDBs_Details.SetValue("U_MECC1", PrICT_oDBs_Details.Offset, PrICTMatrix.Columns.Item("V_9").Cells.Item(iPrLineNum).Specific.Value)
                            PrICT_oDBs_Details.SetValue("U_MECC2", PrICT_oDBs_Details.Offset, PrICTMatrix.Columns.Item("V_10").Cells.Item(iPrLineNum).Specific.Value)
                            PrICT_oDBs_Details.SetValue("U_MECC3", PrICT_oDBs_Details.Offset, PrICTMatrix.Columns.Item("V_11").Cells.Item(iPrLineNum).Specific.Value)
                            PrICT_oDBs_Details.SetValue("U_MECC4", PrICT_oDBs_Details.Offset, PrICTMatrix.Columns.Item("V_12").Cells.Item(iPrLineNum).Specific.Value)
                            PrICT_oDBs_Details.SetValue("U_MECC5", PrICT_oDBs_Details.Offset, PrICTMatrix.Columns.Item("V_13").Cells.Item(iPrLineNum).Specific.Value)
                            PrICT_oDBs_Details.SetValue("U_MEPRLN", PrICT_oDBs_Details.Offset, PrICTMatrix.Columns.Item("V_14").Cells.Item(iPrLineNum).Specific.Value)
                            PrICT_oDBs_Details.SetValue("U_MEDSTRL", PrICT_oDBs_Details.Offset, PrICTMatrix.Columns.Item("V_8").Cells.Item(iPrLineNum).Specific.Value)
                            PrICT_oDBs_Details.SetValue("U_MEBWHS", PrICT_oDBs_Details.Offset, PrICTMatrix.Columns.Item("V_16").Cells.Item(iPrLineNum).Specific.Value)
                            If sTransactionType = "GRPO" Then
                                PrICT_oDBs_Details.SetValue("U_MEBBL", PrICT_oDBs_Details.Offset, Bin_Qty)
                            Else
                                PrICT_oDBs_Details.SetValue("U_MEBBL", PrICT_oDBs_Details.Offset, PrICTMatrix.Columns.Item("V_17").Cells.Item(iPrLineNum).Specific.Value)
                            End If
                            PrICT_oDBs_Details.SetValue("U_MESWHS", PrICT_oDBs_Details.Offset, PrICTMatrix.Columns.Item("V_18").Cells.Item(iPrLineNum).Specific.Value)
                            If sTransactionType = "Delivery" Then
                                PrICT_oDBs_Details.SetValue("U_MESBL", PrICT_oDBs_Details.Offset, Bin_Qty)
                            Else
                                PrICT_oDBs_Details.SetValue("U_MESBL", PrICT_oDBs_Details.Offset, PrICTMatrix.Columns.Item("V_19").Cells.Item(iPrLineNum).Specific.Value)
                            End If
                            PrICT_oDBs_Details.SetValue("U_METYP", PrICT_oDBs_Details.Offset, PrICTMatrix.Columns.Item("V_20").Cells.Item(iPrLineNum).Specific.Value)

                            PrICT_oDBs_Details.SetValue("U_MECBN", PrICT_oDBs_Details.Offset, PrICTMatrix.Columns.Item("V_21").Cells.Item(iPrLineNum).Specific.Value)
                            PrICT_oDBs_Details.SetValue("U_MEBBS", PrICT_oDBs_Details.Offset, PrICTMatrix.Columns.Item("V_22").Cells.Item(iPrLineNum).Specific.Value)
                            PrICT_oDBs_Details.SetValue("U_MESBS", PrICT_oDBs_Details.Offset, PrICTMatrix.Columns.Item("V_23").Cells.Item(iPrLineNum).Specific.Value)


                            PrICTMatrix.SetLineData(iPrLineNum)
                        End If
                        objForm.Close()
                    End If

            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
