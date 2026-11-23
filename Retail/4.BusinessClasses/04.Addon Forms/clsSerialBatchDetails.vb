
Public Class clsSerialBatchDetails

#Region "        Declaration        "
    Dim objForm, MIForm As SAPbouiCOM.Form
    Dim MIMatrix As SAPbouiCOM.Matrix
    Dim oGrid As SAPbouiCOM.Grid
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim oDt As SAPbouiCOM.DataTable
    Dim oEditTextCol As SAPbouiCOM.EditTextColumn
    Dim oDBs_Head, oDBs_Details, MI_oDBs_Details As SAPbouiCOM.DBDataSource
    Public sItemCode As String
    Public dQuantity As Double
    Public sWarehouse As String
    Public iMILineNum As Integer
    Public sBaseFormUID As String
    Public sItemType As String
    Public sBinChk As String
    Public sDispatchNo As String

    Dim oEditText As SAPbouiCOM.EditText
#End Region

    Sub CreateForm(ByVal BaseFormUID As String, ByVal ItemCode As String, ByVal Quantity As Double, _
                   ByVal Warehouse As String, ByVal PrLineNum As Integer, ByVal ItemType As String, ByVal BinChk As String, ByVal DispatchNo As String)
        Try
            objMain.objUtilities.LoadForm("SerialBatchDetails.xml", "ME_SBD_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_SBD_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)

            oDt = objForm.DataSources.DataTables.Add("dt1")
            oDt = objForm.DataSources.DataTables.Item("dt1")

            Me.sBaseFormUID = BaseFormUID
            Me.sItemCode = ItemCode
            Me.dQuantity = Quantity
            Me.sWarehouse = Warehouse
            Me.sItemType = ItemType
            Me.sBinChk = BinChk
            Me.sDispatchNo = DispatchNo
            Me.iMILineNum = PrLineNum
            If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = False Then
                Retail.MainCls.ohtLookUpForm.Add(objForm.UniqueID, BaseFormUID)
            End If

           


            objForm.DataSources.UserDataSources.Add("U_MEDB", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_MEWHS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_MEITM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_MEQTY", SAPbouiCOM.BoDataType.dt_QUANTITY)
            objForm.DataSources.UserDataSources.Add("U_MEITTYP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_MEBNCK", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            oEditText = objForm.Items.Item("5").Specific
            oEditText.DataBind.SetBound(True, "", "U_MEDB")
            objForm.Items.Item("5").AffectsFormMode = False

            'If sTransactionType = "GRPO" Then
            '    objForm.DataSources.UserDataSources.Item("U_MEDB").Value = sBuyerDB
            'ElseIf sTransactionType = "Delivery" Then
            '    objForm.DataSources.UserDataSources.Item("U_MEDB").Value = sSellerDB
            'End If

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

            oEditText = objForm.Items.Item("13").Specific
            oEditText.DataBind.SetBound(True, "", "U_MEITTYP")
            objForm.Items.Item("13").AffectsFormMode = False
            objForm.DataSources.UserDataSources.Item("U_MEITTYP").Value = sItemType

            oEditText = objForm.Items.Item("15").Specific
            oEditText.DataBind.SetBound(True, "", "U_MEBNCK")
            objForm.Items.Item("15").AffectsFormMode = False
            objForm.DataSources.UserDataSources.Item("U_MEBNCK").Value = sBinChk

            Me.LoadGrid(objForm.UniqueID, ItemCode, dQuantity, sWarehouse)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub LoadGrid(ByVal FormUID As String, ByVal ItemCode As String, ByVal Quantity As Double, _
                   ByVal Warehouse As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            oGrid = objForm.Items.Item("3").Specific
            oGrid.DataTable = oDt
            oGrid.DataTable.Clear()
            objForm.Items.Item("16").Visible = False

            Dim GetSerialNoDetails As String = " SELECT Distinct  'N' ""Check""  , IFNULL(T4.""SuppSerial"" ,  T4.""IntrSerial"") As ""SRNo"" FROM OWTR T0 INNER JOIN WTR1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" INNER JOIN OITW T2 " & _
                                               " ON T1.""ItemCode"" = T2.""ItemCode"" AND T1.""WhsCode"" = T2.""WhsCode"" LEFT JOIN SRI1 T3 " & _
                                               " ON T3.""BaseEntry""=T1.""DocEntry"" AND T3.""BaseLinNum"" = T1.""LineNum"" AND T3.""BaseType"" ='67' " & _
                                               " LEFT JOIN OSRI T4 ON T3.""ItemCode""=T4.""ItemCode"" AND T3.""SysSerial"" =T4.""SysSerial"" WHERE " & _
                                               " T1.""DocEntry"" = '" & sDispatchNo & "' And  " & _
                                               " T1.""ItemCode"" ='" & sItemCode & "' And T1.""WhsCode"" = '" & sWarehouse & "' "
            oGrid.DataTable.ExecuteQuery(GetSerialNoDetails)
            oGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item("Check").Editable = True
            oGrid.Columns.Item("SRNo").Editable = False

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
                    If pVal.ItemUID = "16" And pVal.BeforeAction = False Then
                        oGrid.DataTable.Rows.Add()
                    End If
                    If pVal.ItemUID = "17" And pVal.BeforeAction = True Then
                        For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                            If oGrid.Rows.IsSelected(i) Then
                                oGrid.DataTable.Rows.Remove(i)
                                Exit For
                            End If
                        Next
                    End If
                    If pVal.ItemUID = "4" And pVal.BeforeAction = False Then
                        MIForm = objMain.objApplication.Forms.Item(sBaseFormUID)
                        MIMatrix = MIForm.Items.Item("3").Specific
                        MI_oDBs_Details = MIForm.DataSources.DBDataSources.Item("@ME_MI_C0")
                        Dim ByrSrWiseQty As String = ""
                        Dim SlrSrndBinWiseQty As String = ""
                        If sItemType = "SerialNo" Then
                            For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                                If oGrid.DataTable.GetValue("Check", i) = "Y" Then
                                    ByrSrWiseQty = ByrSrWiseQty & oGrid.DataTable.GetValue("SRNo", i) & ";"
                                End If
                            Next
                        End If
                        MI_oDBs_Details.SetValue("LineId", MI_oDBs_Details.Offset, iMILineNum)
                        MI_oDBs_Details.SetValue("U_MEBRCD", MI_oDBs_Details.Offset, MIMatrix.Columns.Item("V_0").Cells.Item(iMILineNum).Specific.Value)
                        MI_oDBs_Details.SetValue("U_MEITMCD", MI_oDBs_Details.Offset, MIMatrix.Columns.Item("V_1").Cells.Item(iMILineNum).Specific.Value)
                        MI_oDBs_Details.SetValue("U_MEITDSC", MI_oDBs_Details.Offset, MIMatrix.Columns.Item("V_2").Cells.Item(iMILineNum).Specific.Value)
                        MI_oDBs_Details.SetValue("U_MEUOM", MI_oDBs_Details.Offset, MIMatrix.Columns.Item("V_3").Cells.Item(iMILineNum).Specific.Value)
                        MI_oDBs_Details.SetValue("U_ADBRCD", MI_oDBs_Details.Offset, MIMatrix.Columns.Item("V_4").Cells.Item(iMILineNum).Specific.Value)
                        MI_oDBs_Details.SetValue("U_MEPKGE", MI_oDBs_Details.Offset, MIMatrix.Columns.Item("V_5").Cells.Item(iMILineNum).Specific.Value)
                        MI_oDBs_Details.SetValue("U_METYPE", MI_oDBs_Details.Offset, MIMatrix.Columns.Item("V_6").Cells.Item(iMILineNum).Specific.Value)
                        MI_oDBs_Details.SetValue("U_MESRBT", MI_oDBs_Details.Offset, ByrSrWiseQty)
                        MI_oDBs_Details.SetValue("U_CHKBIN", MI_oDBs_Details.Offset, MIMatrix.Columns.Item("V_8").Cells.Item(iMILineNum).Specific.Checked)
                        MI_oDBs_Details.SetValue("U_METWHS", MI_oDBs_Details.Offset, MIMatrix.Columns.Item("V_9").Cells.Item(iMILineNum).Specific.Value)
                        MI_oDBs_Details.SetValue("U_METWBL", MI_oDBs_Details.Offset, MIMatrix.Columns.Item("V_10").Cells.Item(iMILineNum).Specific.Value)
                        MI_oDBs_Details.SetValue("U_METBRCH", MI_oDBs_Details.Offset, MIMatrix.Columns.Item("V_11").Cells.Item(iMILineNum).Specific.Value)
                        MI_oDBs_Details.SetValue("U_MEREQQTY", MI_oDBs_Details.Offset, MIMatrix.Columns.Item("V_12").Cells.Item(iMILineNum).Specific.Value)
                        MI_oDBs_Details.SetValue("U_MERECQTY", MI_oDBs_Details.Offset, MIMatrix.Columns.Item("V_13").Cells.Item(iMILineNum).Specific.Value)
                        MI_oDBs_Details.SetValue("U_MEVARQTY", MI_oDBs_Details.Offset, MIMatrix.Columns.Item("V_14").Cells.Item(iMILineNum).Specific.Value)
                        MI_oDBs_Details.SetValue("U_MESTRNO", MI_oDBs_Details.Offset, MIMatrix.Columns.Item("V_15").Cells.Item(iMILineNum).Specific.Value)
                        MI_oDBs_Details.SetValue("U_MESTREN", MI_oDBs_Details.Offset, MIMatrix.Columns.Item("V_16").Cells.Item(iMILineNum).Specific.Value)
                        MI_oDBs_Details.SetValue("U_MEITRNO", MI_oDBs_Details.Offset, MIMatrix.Columns.Item("V_17").Cells.Item(iMILineNum).Specific.Value)
                        MI_oDBs_Details.SetValue("U_MEITREN", MI_oDBs_Details.Offset, MIMatrix.Columns.Item("V_18").Cells.Item(iMILineNum).Specific.Value)
                        MI_oDBs_Details.SetValue("U_MEBSTYP", MI_oDBs_Details.Offset, MIMatrix.Columns.Item("V_19").Cells.Item(iMILineNum).Specific.Value)
                        MI_oDBs_Details.SetValue("U_MEBLN", MI_oDBs_Details.Offset, MIMatrix.Columns.Item("V_20").Cells.Item(iMILineNum).Specific.Value)
                        MIMatrix.SetLineData(iMILineNum)
                        objForm.Close()
                    End If

            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
