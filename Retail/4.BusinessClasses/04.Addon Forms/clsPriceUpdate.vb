
Public Class clsPriceUpdate

#Region "    Declaration   "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oEditText As SAPbouiCOM.EditText
    Dim oDt As SAPbouiCOM.DataTable
    Dim sItemCode, sItemName, sBaseNo, sBaseEntry As String
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim oGrid As SAPbouiCOM.Grid
    Dim sDoNum As String
    Dim FDoNum As String
    Dim BaseBarcode As String = ""
    Dim FindFlag As Boolean = False
    Dim BaseFlagOITM As Boolean = False
    Dim BaseItemCode As String = ""
#End Region
    Sub CreateForm(ByVal FormUID As String, ByVal ItemCode As String, ByVal BarCode As String, ByVal ItemName As String, ByVal IfPriceUpdateExist As String, Optional PriceUpdateDocNum As String = "", Optional ItemGroup As String = "")
        Try
            objMain.objUtilities.LoadForm("PriceUpdate.xml", "ME_PRICEUPD_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_PRICEUPD_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.PaneLevel = 1
            If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = False Then
                Retail.MainCls.ohtLookUpForm.Add(objForm.UniqueID, FormUID)
            End If
            sItemCode = ItemCode
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD_C0")
            objMatrix = objForm.Items.Item("3").Specific
            BaseFlagOITM = True
            objMatrix.CommonSetting.EnableArrowKey = True
            objForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            objComboBox = objForm.Items.Item("9").Specific
            Dim GetSeriesQty As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'ME_OPRICEUPD' And ""Locked"" = 'N' Order By ""Series"" Desc "
            objMain.objUtilities.ComboBoxLoadValuesSeries(objForm.Items.Item("9").Specific, GetSeriesQty)
            objForm.Items.Item("9").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            If IfPriceUpdateExist = "Y" Then
                objForm.Freeze(True)
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                objForm.Items.Item("10").Specific.value = PriceUpdateDocNum

                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                objForm.Items.Item("Item_1").Specific.value = BarCode

                For i As Integer = 1 To objMatrix.RowCount
                    objMatrix.GetLineData(i)
                    oDBs_Details.SetValue("U_ENDDT", oDBs_Details.Offset, "20401231")
                    objMatrix.SetLineData(i)
                Next
                'Me.CFLUOM(objForm.UniqueID, "CFL_UOM", objForm.Items.Item("5").Specific.Value)
                objForm.Freeze(False)
            End If

            If IfPriceUpdateExist = "N" Then
                sItemCode = ItemCode
                sItemName = ItemName
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                sDoNum = objMain.objUtilities.GetNextDocNum(objForm, "ME_OPRICEUPD")
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, sDoNum)
                oDBs_Head.SetValue("U_ITMCD", oDBs_Head.Offset, sItemCode)
                oDBs_Head.SetValue("U_ITMNM", oDBs_Head.Offset, sItemName)
                oDBs_Head.SetValue("U_BARCD", oDBs_Head.Offset, BarCode)
                'oDBs_Head.SetValue("U_APPSTS", oDBs_Head.Offset, "Open")
                oDBs_Head.SetValue("U_USRID", oDBs_Head.Offset, objMain.objCompany.UserName)

                Dim GetDflt As String = "Select ""DfltSeries"" From ONNM Where ""ObjectCode"" = 'ME_OPRICEUPD'"
                Dim oRsGetDflt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetDflt.DoQuery(GetDflt)

                oDBs_Head.SetValue("U_SERIES", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)
                oDBs_Head.SetValue("Series", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)

                objMatrix.Clear()
                oDBs_Details.Clear()
                objMatrix.FlushToDataSource()
                Me.SetNewLine(objForm.UniqueID)
            End If

            Me.CellMasking(objForm.UniqueID)

            oDt = objForm.DataSources.DataTables.Add("dt1")
            oDt = objForm.DataSources.DataTables.Item("dt1")


            Dim chkForAuthorization As String = " Select ""U_ACCESS"" , ""U_PRICEDIT"" FROM OUSR T0 Inner Join ""@SBO_PUIGA"" T1 On ""USER_CODE"" = ""U_USRCD"" " &
                                                " Inner Join ""@SBO_PUIGA_C0"" T2 On T1.""DocEntry"" = T2.""DocEntry"" " &
                                                " Where IFNULL(""U_PRICEDIT"",'N')='Y' And ""USER_CODE"" = '" & objMain.objCompany.UserName & "' and ""U_ITMGRP"" = '" & ItemGroup & "'"
            Dim oRschkForAuthorization As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRschkForAuthorization.DoQuery(chkForAuthorization)
            If oRschkForAuthorization.RecordCount > 0 Then
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
            End If
            objForm.EnableMenu("1304", True)
            objForm.Items.Item("21").Enabled = False
            objForm.Items.Item("14").Enabled = False

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("PriceUpdate.xml", "ME_PRICEUPD_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_PRICEUPD_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            'objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            'objForm.PaneLevel = 1
            'BaseFlagOITM = False
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD_C0")
            objMatrix = objForm.Items.Item("3").Specific
            objMatrix.CommonSetting.EnableArrowKey = True
            objForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            objComboBox = objForm.Items.Item("9").Specific
            Dim GetSeriesQty As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'ME_OPRICEUPD' And ""Locked"" = 'N' Order By ""Series"" Desc "
            objMain.objUtilities.ComboBoxLoadValuesSeries(objForm.Items.Item("9").Specific, GetSeriesQty)
            objForm.Items.Item("9").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("14").Visible = False
            Me.CellMasking(objForm.UniqueID)

            oDt = objForm.DataSources.DataTables.Add("dt1")
            oDt = objForm.DataSources.DataTables.Item("dt1")
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Sub CellMasking(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Items.Item("5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("Item_1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "2", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD_C0")
            objMatrix = objForm.Items.Item("3").Specific

            objMatrix.AddRow()
            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
            oDBs_Details.SetValue("U_ITMCD", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_BRCHID", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_BRCHNM", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_PRCID", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_PRCLST", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_UOMENT", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_UOMCD", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_OLDPRC", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_NEWPRC", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_UPDT", oDBs_Details.Offset, "20211101")
            oDBs_Details.SetValue("U_ENDDT", oDBs_Details.Offset, "20401231")
            oDBs_Details.SetValue("U_USRCD", oDBs_Details.Offset, "")

            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD_C0")
                    objMatrix = objForm.Items.Item("3").Specific

                    If pVal.ItemUID = "22" And pVal.BeforeAction = False And pVal.ItemChanged = True Then
                        Try
                            objForm.Freeze(True)

                            Dim Units As String = objForm.Items.Item("22").Specific.Value

                            Me.LoadBranchWisePrices(objForm.UniqueID, Units.Trim)

                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If

                    If pVal.ItemUID = "3" And pVal.ColUID = "V_8" And pVal.BeforeAction = False Then
                        If objMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value > 0 Then
                            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row)
                            oDBs_Details.SetValue("U_ITMCD", oDBs_Details.Offset, objMatrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_BRCHID", oDBs_Details.Offset, objMatrix.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_BRCHNM", oDBs_Details.Offset, objMatrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_PRCID", oDBs_Details.Offset, objMatrix.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_PRCLST", oDBs_Details.Offset, objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_UOMENT", oDBs_Details.Offset, objMatrix.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_UOMCD", oDBs_Details.Offset, objMatrix.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_OLDPRC", oDBs_Details.Offset, objMatrix.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_NEWPRC", oDBs_Details.Offset, objMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_UPDT", oDBs_Details.Offset, DateTime.Now.ToString("yyyyMMdd"))
                            oDBs_Details.SetValue("U_ENDDT", oDBs_Details.Offset, "20401231")
                            oDBs_Details.SetValue("U_USRCD", oDBs_Details.Offset, objMain.objCompany.UserName.ToString)
                            objMatrix.SetLineData(pVal.Row)
                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            End If
                            'ElseIf objMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value = 0 And objMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            '    oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row)
                            '    oDBs_Details.SetValue("U_ITMCD", oDBs_Details.Offset, objMatrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                            '    oDBs_Details.SetValue("U_BRCHID", oDBs_Details.Offset, objMatrix.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                            '    oDBs_Details.SetValue("U_BRCHNM", oDBs_Details.Offset, objMatrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                            '    oDBs_Details.SetValue("U_PRCID", oDBs_Details.Offset, objMatrix.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                            '    oDBs_Details.SetValue("U_PRCLST", oDBs_Details.Offset, objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                            '    oDBs_Details.SetValue("U_UOMENT", oDBs_Details.Offset, objMatrix.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value)
                            '    oDBs_Details.SetValue("U_UOMCD", oDBs_Details.Offset, objMatrix.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                            '    oDBs_Details.SetValue("U_OLDPRC", oDBs_Details.Offset, objMatrix.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value)
                            '    oDBs_Details.SetValue("U_NEWPRC", oDBs_Details.Offset, objMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value)
                            '    oDBs_Details.SetValue("U_UPDT", oDBs_Details.Offset, "")
                            '    objMatrix.SetLineData(pVal.Row)
                            '    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            '        objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            '    End If
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD_C0")
                    objMatrix = objForm.Items.Item("3").Specific
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                    '' If objForm.Items.Item("5").Specific.Value <> "" Then
                    If oCFL.UniqueID = "CFL_UOM" And pVal.BeforeAction = True Then
                        Me.CFLUOM(objForm.UniqueID, "CFL_UOM", objForm.Items.Item("5").Specific.Value)
                    End If
                    ''End If
                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        If oCFL.UniqueID = "CFL_UOM" Then
                            oDBs_Head.SetValue("U_UNITS", oDBs_Head.Offset, oDT.GetValue("UomCode", 0))
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD_C0")
                    objMatrix = objForm.Items.Item("3").Specific

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE And FindFlag = True Then
                        oDBs_Head.SetValue("U_BARCD", oDBs_Head.Offset, BaseBarcode)
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim chkIfExist As String = "Select ""DocNum"" From ""@ME_PRICEUPD"" Where ""DocEntry""=(Select MAX(""DocEntry"") From ""@ME_PRICEUPD"")"
                        Dim oRschkIfExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRschkIfExist.DoQuery(chkIfExist)
                        If oRschkIfExist.RecordCount > 0 Then
                            objForm.Freeze(True)
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            objForm.Items.Item("10").Specific.value = oRschkIfExist.Fields.Item("DocNum").Value
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)


                            objForm.Freeze(False)
                        End If
                    End If

                    If pVal.ItemUID = "16" And pVal.BeforeAction = False Then
                        objForm.PaneLevel = 1
                        '  objForm.Items.Item("13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        objMatrix.AutoResizeColumns()
                    End If

                    If pVal.ItemUID = "17" And pVal.BeforeAction = False Then
                        objForm.PaneLevel = 2
                        '  objForm.Items.Item("13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    End If

                    If pVal.ItemUID = "15" And pVal.BeforeAction = False Then
                        ' Me.UpdateVendorPrice(objForm.UniqueID)
                        Me.UpdateSpecialPriceBP(objForm.UniqueID)
                    End If

                    If pVal.ItemUID = "19" And pVal.BeforeAction = False Then
                        'Me.LoadVendorPriceList(objForm.UniqueID)
                        Me.LoadSpecialPricesBPList(objForm.UniqueID)
                    End If

                    If pVal.ItemUID = "21" And pVal.BeforeAction = False Then
                        Me.PriceAutoFill(objForm.UniqueID)
                    End If

                    If pVal.ItemUID = "13" And pVal.BeforeAction = False Then
                        'If objForm.Items.Item("12").Specific.Value.Trim = "Submitted" Then
                        '    objMain.objApplication.StatusBar.SetText("Operation stopped..Document is in submitted waiting for approval !!.... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    Exit Sub
                        '    'ElseIf objForm.Items.Item("12").Specific.Value.Trim = "Approved" Then
                        '    '    objMain.objApplication.StatusBar.SetText("Operation stopped..Document is in approved by authorizer !!.... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    '    Exit Sub
                        'End If
                        'start
                        'If objForm.Items.Item("5").Specific.Value = "" Then
                        '    If objForm.Items.Item("Item_1").Specific.Value <> "" Then
                        '        Dim BarcodeQry As String = "Select * from OITM Where ""CodeBars""='" + objForm.Items.Item("Item_1").Specific.Value + "'"
                        '        Dim orsBarcode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '        orsBarcode.DoQuery(BarcodeQry)
                        '        If orsBarcode.RecordCount > 0 Then
                        '            oDBs_Head.SetValue("U_ITMCD", oDBs_Head.Offset, orsBarcode.Fields.Item("ItemCode").Value)
                        '            oDBs_Head.SetValue("U_ITMNM", oDBs_Head.Offset, orsBarcode.Fields.Item("ItemName").Value)
                        '        End If
                        '    End If
                        'End If
                        'If objForm.Items.Item("Item_1").Specific.Value = "" Then
                        '    If objForm.Items.Item("5").Specific.Value <> "" Then
                        '        Dim BarcodeQry As String = "Select * from OITM Where ""ItemCode""='" + objForm.Items.Item("5").Specific.Value + "'"
                        '        Dim orsBarcode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '        orsBarcode.DoQuery(BarcodeQry)
                        '        If orsBarcode.RecordCount > 0 Then
                        '            oDBs_Head.SetValue("U_BARCD", oDBs_Head.Offset, orsBarcode.Fields.Item("CodeBars").Value)
                        '            oDBs_Head.SetValue("U_ITMNM", oDBs_Head.Offset, orsBarcode.Fields.Item("ItemName").Value)
                        '        End If
                        '    End If
                        'End If
                        Dim result As Integer = objMain.objApplication.MessageBox("Matrix data will get refresh. Do you want to continue?", 1, "Continue", "Cancel")
                        If (result = 1) Then
                            Me.LoadBranchWisePrices(objForm.UniqueID)
                        End If
                        'objForm.Items.Item("13").Enabled = False
                        objForm.Items.Item("21").Enabled = True
                        objForm.Items.Item("14").Enabled = False
                    End If


                    'If pVal.ItemUID = "14" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    '    If (objForm.Items.Item("12").Specific.Value.Trim = "Open" OrElse objForm.Items.Item("12").Specific.Value.Trim = "Approved") Then
                    '        Dim result As Integer = objMain.objApplication.MessageBox("Submitting document is irreversible..Confirm... Do you want to continue?", 1, "Continue", "Cancel")
                    '        If (result = 1) Then
                    '            objMain.objApplication.StatusBar.SetText("Submission is in progress...Please wait.....!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    '            Me.UpdatePrice(objForm.UniqueID)
                    '            'Me.UpdateDocStatus(objForm.UniqueID, "Submitted")
                    '            'objMain.objApplication.StatusBar.SetText("Document submitted successfully !!.... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    '        End If
                    '    ElseIf objForm.Items.Item("12").Specific.Value.Trim = "Submitted" Then
                    '        'Me.UpdateDocStatus(objForm.UniqueID, "Open")
                    '        'objMain.objApplication.StatusBar.SetText("Document is already submitted waiting for approval !!.... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        'Exit Sub
                    '        'ElseIf objForm.Items.Item("12").Specific.Value.Trim = "Approved" Then
                    '        '    objMain.objApplication.StatusBar.SetText("Document is already approved by authorizer !!.... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        '    Exit Sub
                    '    End If
                    'End If

                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD_C0")
                    objMatrix = objForm.Items.Item("3").Specific

                    If pVal.ItemUID = "3" And pVal.ColUID = "V_8" And pVal.BeforeAction = True Then


                        'If (objForm.Items.Item("12").Specific.Value.Trim = "Submitted") Then
                        '    BubbleEvent = False
                        '    objMain.objApplication.StatusBar.SetText("Price modification is not permitted when document is under " & objForm.Items.Item("12").Specific.Value.Trim & " !!.... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        'End If
                        If (objForm.Items.Item("12").Specific.Value.Trim = "Approved") Then
                            BubbleEvent = False
                            objMain.objApplication.StatusBar.SetText("Price modification is not permitted when document is under " & objForm.Items.Item("12").Specific.Value.Trim & " !!.... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD_C0")
                    objMatrix = objForm.Items.Item("3").Specific
                    If pVal.ItemUID = "3" And pVal.ColUID = "V_8" And pVal.BeforeAction = True Then
                        Dim chkForAuthorization As String = " Select ""U_ACCESS"" , ""U_PRICEDIT"" FROM OUSR T0 Inner Join ""@SBO_PUBA"" T1 On ""USER_CODE"" = ""U_USRCD"" " &
                                                            " Inner Join ""@SBO_PUBA_C0"" T2 On T1.""DocEntry"" = T2.""DocEntry"" " &
                                                            " Where IFNULL(""U_PRICEDIT"",'N')='Y' And ""USER_CODE"" = '" & objMain.objCompany.UserName & "' and ""U_BPLCD"" = '" & objMatrix.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value & "'"
                        Dim oRschkForAuthorization As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRschkForAuthorization.DoQuery(chkForAuthorization)
                        If oRschkForAuthorization.RecordCount > 0 Then

                            objMain.objApplication.StatusBar.SetText("User not authorized to performe this operation...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            objMatrix.Columns.Item("V_9").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            BubbleEvent = False
                        End If
                    End If
                    If pVal.ItemUID = "3" And pVal.ColUID = "V_10" And pVal.BeforeAction = True Then
                        Dim chkForAuthorization As String = " Select ""U_ACCESS"" , ""U_PRICEDIT"" FROM OUSR T0 Inner Join ""@SBO_PUBA"" T1 On ""USER_CODE"" = ""U_USRCD"" " &
                                                            " Inner Join ""@SBO_PUBA_C0"" T2 On T1.""DocEntry"" = T2.""DocEntry"" " &
                                                            " Where IFNULL(""U_PRICEDIT"",'N')='Y' And ""USER_CODE"" = '" & objMain.objCompany.UserName & "' and ""U_BPLCD"" = '" & objMatrix.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value & "'"
                        Dim oRschkForAuthorization As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRschkForAuthorization.DoQuery(chkForAuthorization)
                        If oRschkForAuthorization.RecordCount > 0 Then
                            objMain.objApplication.StatusBar.SetText("User not authorized to performe this operation...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                            BubbleEvent = False
                        End If
                    End If
                    If pVal.ItemUID = "1" And pVal.Before_Action = True And (objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                        'If objForm.Items.Item("5").Specific.Value <> "" Then
                        '    BaseItemCode = objForm.Items.Item("5").Specific.Value
                        '    FuturePriceUpdate(FormUID)
                        'End If

                    End If
                    If pVal.ItemUID = "1" And pVal.Before_Action = True And objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        Dim BarcodeQry As String = ""
                        Dim orsBarcode As SAPbobsCOM.Recordset
                        If objForm.Items.Item("Item_1").Specific.Value <> "" Then
                            objForm.Items.Item("5").Specific.Value = ""
                            BaseBarcode = objForm.Items.Item("Item_1").Specific.Value

                            BarcodeQry = "SELECT *  FROM ""@ME_PRICEUPD""  T0 WHERE T0.""U_ITMCD"" ='" + objForm.Items.Item("5").Specific.Value + "'"
                            orsBarcode = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            orsBarcode.DoQuery(BarcodeQry)
                            If orsBarcode.RecordCount > 0 Then

                                BarcodeQry = "SELECT T0.""ItemCode"", T1.""ItemName"" FROM OBCD T0  INNER JOIN OITM T1 ON T0.""ItemCode"" = T1.""ItemCode"" WHERE T0.""BcdCode""='" + objForm.Items.Item("Item_1").Specific.Value + "'"
                                orsBarcode = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                orsBarcode.DoQuery(BarcodeQry)
                                If orsBarcode.RecordCount > 0 Then
                                    oDBs_Head.SetValue("U_ITMCD", oDBs_Head.Offset, orsBarcode.Fields.Item("ItemCode").Value)
                                    oDBs_Head.SetValue("U_ITMNM", oDBs_Head.Offset, orsBarcode.Fields.Item("ItemName").Value)
                                    objForm.Items.Item("Item_1").Specific.Value = ""
                                    FindFlag = True
                                End If


                            Else
                                Dim ItemGrpCode As String = ""
                                If BaseFlagOITM = True Then
                                    BarcodeQry = "SELECT T0.""ItemCode"", T1.""ItemName"" ,T1.""ItmsGrpCod"" FROM OBCD T0  INNER JOIN OITM T1 ON T0.""ItemCode"" = T1.""ItemCode"" WHERE T0.""BcdCode""='" + objForm.Items.Item("Item_1").Specific.Value + "'"
                                    orsBarcode = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    orsBarcode.DoQuery(BarcodeQry)
                                    If orsBarcode.RecordCount > 0 Then
                                        oDBs_Head.SetValue("U_ITMCD", oDBs_Head.Offset, orsBarcode.Fields.Item("ItemCode").Value)
                                        oDBs_Head.SetValue("U_ITMNM", oDBs_Head.Offset, orsBarcode.Fields.Item("ItemName").Value)
                                        objForm.Items.Item("Item_1").Specific.Value = ""
                                        FindFlag = True
                                    End If
                                Else
                                    BarcodeQry = "SELECT T0.""ItemCode"", T1.""ItemName"" FROM OBCD T0  INNER JOIN OITM T1 ON T0.""ItemCode"" = T1.""ItemCode"" WHERE T0.""BcdCode""='" + objForm.Items.Item("Item_1").Specific.Value + "'"
                                    orsBarcode = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    orsBarcode.DoQuery(BarcodeQry)
                                    If orsBarcode.RecordCount > 0 Then
                                        oDBs_Head.SetValue("U_ITMCD", oDBs_Head.Offset, orsBarcode.Fields.Item("ItemCode").Value)
                                        oDBs_Head.SetValue("U_ITMNM", oDBs_Head.Offset, orsBarcode.Fields.Item("ItemName").Value)
                                        objForm.Items.Item("Item_1").Specific.Value = ""
                                        FindFlag = True
                                    Else
                                        objMain.objApplication.StatusBar.SetText("It's a New Item and New Barcode. Price is not updated Yet, Kindly use Item Master Data to load price...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                    End If
                                    'BarcodeQry = "SELECT T0.""ItemCode"", T1.""ItemName"",T1.""ItmsGrpCod"" FROM OBCD T0  INNER JOIN OITM T1 ON T0.""ItemCode"" = T1.""ItemCode"" WHERE T0.""BcdCode""='" + objForm.Items.Item("Item_1").Specific.Value + "'"
                                    'orsBarcode = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    'orsBarcode.DoQuery(BarcodeQry)
                                    'If orsBarcode.RecordCount > 0 Then
                                    '    oDBs_Head.SetValue("U_ITMCD", oDBs_Head.Offset, orsBarcode.Fields.Item("ItemCode").Value)
                                    '    oDBs_Head.SetValue("U_ITMNM", oDBs_Head.Offset, orsBarcode.Fields.Item("ItemName").Value)
                                    '    objForm.Items.Item("Item_1").Specific.Value = ""
                                    '    ItemGrpCode = orsBarcode.Fields.Item("ItmsGrpCod").Value
                                    '    FindFlag = False
                                    'End If
                                    'objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                                    ''//
                                    'objForm = objMain.objApplication.Forms.GetForm("ME_PRICEUPD_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                                    'objForm.PaneLevel = 1
                                    'If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = False Then
                                    '    Retail.MainCls.ohtLookUpForm.Add(objForm.UniqueID, FormUID)
                                    'End If
                                    'sItemCode = orsBarcode.Fields.Item("ItemCode").Value
                                    'oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD")
                                    'oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD_C0")
                                    'objMatrix = objForm.Items.Item("3").Specific

                                    'objMatrix.CommonSetting.EnableArrowKey = True

                                    'objComboBox = objForm.Items.Item("9").Specific
                                    'Dim GetSeriesQty As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'ME_OPRICEUPD' And ""Locked"" = 'N' Order By ""Series"" Desc "
                                    'objMain.objUtilities.ComboBoxLoadValuesSeries(objForm.Items.Item("9").Specific, GetSeriesQty)
                                    'objForm.Items.Item("9").DisplayDesc = True
                                    'objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly



                                    'sItemCode = orsBarcode.Fields.Item("ItemCode").Value
                                    'sItemName = orsBarcode.Fields.Item("ItemName").Value
                                    'objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                                    'sDoNum = objMain.objUtilities.GetNextDocNum(objForm, "ME_OPRICEUPD")
                                    'oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, sDoNum)
                                    'oDBs_Head.SetValue("U_ITMCD", oDBs_Head.Offset, sItemCode)
                                    'oDBs_Head.SetValue("U_ITMNM", oDBs_Head.Offset, sItemName)
                                    'oDBs_Head.SetValue("U_BARCD", oDBs_Head.Offset, BaseBarcode)
                                    ''oDBs_Head.SetValue("U_APPSTS", oDBs_Head.Offset, "Open")
                                    'oDBs_Head.SetValue("U_USRID", oDBs_Head.Offset, objMain.objCompany.UserName)

                                    'Dim GetDflt As String = "Select ""DfltSeries"" From ONNM Where ""ObjectCode"" = 'ME_OPRICEUPD'"
                                    'Dim oRsGetDflt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    'oRsGetDflt.DoQuery(GetDflt)

                                    'oDBs_Head.SetValue("U_SERIES", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)
                                    'oDBs_Head.SetValue("Series", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)

                                    'objMatrix.Clear()
                                    'oDBs_Details.Clear()
                                    'objMatrix.FlushToDataSource()
                                    'Me.SetNewLine(objForm.UniqueID)

                                    'Me.CellMasking(objForm.UniqueID)



                                    'Dim chkForAuthorization As String = " Select ""U_ACCESS"" , ""U_PRICEDIT"" FROM OUSR T0 Inner Join ""@SBO_PUIGA"" T1 On ""USER_CODE"" = ""U_USRCD"" " &
                                    '                " Inner Join ""@SBO_PUIGA_C0"" T2 On T1.""DocEntry"" = T2.""DocEntry"" " &
                                    '                " Where IFNULL(""U_PRICEDIT"",'N')='Y' And ""USER_CODE"" = '" & objMain.objCompany.UserName & "' and ""U_ITMGRP"" = '" & ItemGrpCode & "' "
                                    '''
                                    'Dim oRschkForAuthorization As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    'oRschkForAuthorization.DoQuery(chkForAuthorization)
                                    'If oRschkForAuthorization.RecordCount > 0 Then
                                    '    objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                                    'End If
                                    '' objForm.Items.Item("13").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    'objForm.Freeze(False)
                                End If


                                ''//
                            End If

                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = True And pVal.BeforeAction = False Then
                        Retail.MainCls.ohtLookUpForm.Remove(objForm.UniqueID)
                    End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub CFLUOM(ByVal FormUID As String, ByVal CFL_ID As String, ByVal ItemCode As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim GetBarCode As String = " Select T1.""UomEntry"" From  UGP1 T1 " &
                                       "  Where T1.""UgpEntry"" " &
                                       " = (Select ""UgpEntry"" From OITM Where ""ItemCode"" = '" & ItemCode & "' AND Ifnull(""IsActive"",'')='Y')"
            Dim oRsGetBarCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetBarCode.DoQuery(GetBarCode)
            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            If oRsGetBarCode.RecordCount > 0 Then
                oCondition = oConditions.Add()
                oCondition.Alias = "UomEntry"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = oRsGetBarCode.Fields.Item(0).Value
                oChooseFromList.SetConditions(oConditions)
                oRsGetBarCode.MoveNext()
                For i As Integer = 1 To oRsGetBarCode.RecordCount - 1
                    oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCondition = oConditions.Add()
                    oCondition.Alias = "UomEntry"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = oRsGetBarCode.Fields.Item(0).Value
                    oChooseFromList.SetConditions(oConditions)
                    oRsGetBarCode.MoveNext()
                Next
            Else
                Dim GetBarCode1 As String = " SELECT ""IUoMEntry""  FROM OITM WHERE " &
                                           " ""ItemCode"" = '" & ItemCode & "'"
                Dim oRsGetBarCode1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetBarCode1.DoQuery(GetBarCode1)
                oCondition = oConditions.Add()
                oCondition.Alias = "UomEntry"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = oRsGetBarCode1.Fields.Item(0).Value
                oChooseFromList.SetConditions(oConditions)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ME_PRICEUPD" And pVal.BeforeAction = False Then
                Me.CreateForm()
            End If
            If pVal.MenuUID = "Delete Row" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("ME_PRICEUPD_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix = objForm.Items.Item("3").Specific
                For i As Integer = 1 To objMatrix.VisualRowCount - 1
                    If objMatrix.IsRowSelected(i) = True Then
                        objMatrix.DeleteRow(i)
                    End If
                Next
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.Columns.Item("V_-1").Cells.Item(i).Specific.string = i
                Next
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("ME_PRICEUPD_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Me.UpdateItemMaster(objForm.UniqueID)
                        'Dim result As Integer = objMain.objApplication.MessageBox("Price updation is irreversible..Confirm... Do you want to continue?", 1, "Continue", "Cancel")
                        'If (result = 1) Then
                        'objMain.objApplication.StatusBar.SetText("Price updation is in progress...Please wait.....!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        'Me.FuturePriceUpdate(objForm.UniqueID)
                        Me.UpdatePrice(objForm.UniqueID)
                        '    objMain.objApplication.StatusBar.SetText("Operation completed successfully.....!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        'End If

                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Me.UpdateItemMaster(objForm.UniqueID)
                        'Dim result As Integer = objMain.objApplication.MessageBox("Price updation is irreversible..Confirm... Do you want to continue?", 1, "Continue", "Cancel")
                        'If (result = 1) Then
                        '    objMain.objApplication.StatusBar.SetText("Price updation is in progress...Please wait.....!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        'Me.FuturePriceUpdate(objForm.UniqueID)
                        Me.UpdatePrice(objForm.UniqueID)
                        '    objMain.objApplication.StatusBar.SetText("Operation completed successfully.....!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        'End If
                        'objForm.Items.Item("13").Enabled = False
                        objForm.Items.Item("21").Enabled = True
                        objForm.Items.Item("14").Enabled = True
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Sub UpdatePrice(ByVal FormUID As String)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        Dim oRsGetItemPriceDetails As SAPbobsCOM.Recordset = Nothing
        Dim oRsGetItemPriceLineDtls As SAPbobsCOM.Recordset = Nothing
        Dim UomCount As Integer = 0
        Try
            Dim GetItemPriceDetails As String = " Select  T1.""U_ITMCD"" , T0.""DocEntry"" As ""DocEntry"" " &
                                                " From ""@ME_PRICEUPD"" T0 Inner Join ""@ME_PRICEUPD_C0"" T1 On T1.""DocEntry"" = T0.""DocEntry"" " &
                                                " Inner Join OITM T2 ON T2.""ItemCode"" = T0.""U_ITMCD"" Where IFNULL(T1.""U_ITMCD"",'')<>'' And T1.""U_NEWPRC"" > 0  " &
                                                " And  T1.""U_ITMCD"" = '" & sItemCode & "' Group By T1.""U_ITMCD""  ,  T0.""DocEntry"" "

            'Changes based on U_UPDT Date condition with CurrentDate
            'Dim GetItemPriceDetails As String = " Select  T1.""U_ITMCD"" , T0.""DocEntry"" As ""DocEntry"" " &
            '                                    " From ""@ME_PRICEUPD"" T0 Inner Join ""@ME_PRICEUPD_C0"" T1 On T1.""DocEntry"" = T0.""DocEntry"" " &
            '                                    " Inner Join OITM T2 ON T2.""ItemCode"" = T0.""U_ITMCD"" Where IFNULL(T1.""U_ITMCD"",'')<>'' And T1.""U_NEWPRC"" > 0  " &
            '                                    " And  T1.""U_ITMCD"" = '" & sItemCode & "' AND T1.""U_UPDT"" = CURRENT_DATE Group By T1.""U_ITMCD""  ,  T0.""DocEntry"" "

            oRsGetItemPriceDetails = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetItemPriceDetails.DoQuery(GetItemPriceDetails)
            If oRsGetItemPriceDetails.RecordCount > 0 Then

                oRsGetItemPriceDetails.MoveFirst()
                Dim result As Integer = objMain.objApplication.MessageBox("Price updation is irreversible..Confirm... Do you want to continue?", 1, "Continue", "Cancel")
                If (result = 1) Then
                    objMain.objApplication.StatusBar.SetText("Price updation is in progress...Please wait.....!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    Exit Sub
                End If

                Dim oItem As SAPbobsCOM.Items = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                For i As Integer = 1 To oRsGetItemPriceDetails.RecordCount
                    If oItem.GetByKey(oRsGetItemPriceDetails.Fields.Item(0).Value) Then
                        Dim GetItemPriceLineDtls As String = " Select  T1.""U_NEWPRC"" As ""New Price"", T1.""U_OLDPRC"" As ""Old Price"" , T1.""U_ITMCD"" As ""Item Code"" , T0.""U_ITMNM"" As ""Item Name"" , " &
                                                             " T1.""U_BRCHID"" As ""Branch ID"" , T1.""U_BRCHNM"" As ""Branch Name"" , ""U_PRCID"" As ""Price List ID"" , ""U_PRCLST"" As ""Price List"" , ""U_UOMENT"" As ""UoM Entry"" , " &
                                                             " ""U_UOMCD"" As ""UoM"" , T0.""U_USRID"" As ""User Id"" , T0.""DocEntry"" , T2.""IUoMEntry"" As ""Default UoM"" " &
                                                             " From ""@ME_PRICEUPD"" T0 Inner Join ""@ME_PRICEUPD_C0"" T1 On T1.""DocEntry"" = T0.""DocEntry"" " &
                                                             " Inner Join OITM T2 ON T2.""ItemCode"" = T0.""U_ITMCD"" Where IFNULL(T1.""U_ITMCD"",'')<>'' And T1.""U_NEWPRC"" > 0 " &
                                                             " And  T1.""U_ITMCD"" = '" & oRsGetItemPriceDetails.Fields.Item(0).Value & "'"

                        'Changes based on U_UPDT Date condition with CurrentDate
                        'Dim GetItemPriceLineDtls As String = " Select  T1.""U_NEWPRC"" As ""New Price"", T1.""U_OLDPRC"" As ""Old Price"" , T1.""U_ITMCD"" As ""Item Code"" , T0.""U_ITMNM"" As ""Item Name"" , " &
                        '                                     " T1.""U_BRCHID"" As ""BranchId"" , T1.""U_BRCHNM"" As ""Branch Name"" , ""U_PRCID"" As ""PricListId"" , ""U_PRCLST"" As ""Price List"" , ""U_UOMENT"" As ""UomEntry"" , " &
                        '                                     " ""U_UOMCD"" As ""UOM"" , T0.""U_USRID"" As ""User Id"" , T0.""DocEntry"" , T2.""IUoMEntry"" As ""DefaultUOM"" " &
                        '                                     " From ""@ME_PRICEUPD"" T0 Inner Join ""@ME_PRICEUPD_C0"" T1 On T1.""DocEntry"" = T0.""DocEntry"" " &
                        '                                     " Inner Join OITM T2 ON T2.""ItemCode"" = T0.""U_ITMCD"" Where IFNULL(T1.""U_ITMCD"",'')<>'' And T1.""U_NEWPRC"" > 0 " &
                        '                                     " And  T1.""U_ITMCD"" = '" & oRsGetItemPriceDetails.Fields.Item(0).Value & "' AND T1.""U_UPDT"" =CURRENT_DATE"

                        oRsGetItemPriceLineDtls = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetItemPriceLineDtls.DoQuery(GetItemPriceLineDtls)
                        If oRsGetItemPriceLineDtls.RecordCount > 0 Then
                            oRsGetItemPriceLineDtls.MoveFirst()
                            For j As Integer = 1 To oRsGetItemPriceLineDtls.RecordCount
                                Dim UomEntry As Integer = CInt(oRsGetItemPriceLineDtls.Fields.Item("UoM Entry").Value)
                                Dim DefaultUom As Integer = CInt(oRsGetItemPriceLineDtls.Fields.Item("Default UoM").Value)
                                Dim PriceListRow As Integer = oRsGetItemPriceLineDtls.Fields.Item("Price List ID").Value
                                oItem.PriceList.SetCurrentLine(PriceListRow - 1)
                                Dim strPrlist As String = oItem.PriceList.PriceList
                                If strPrlist = oRsGetItemPriceLineDtls.Fields.Item("Price List ID").Value Then
                                    If DefaultUom = UomEntry Then
                                        oItem.PriceList.Price = CDbl(oRsGetItemPriceLineDtls.Fields.Item("New Price").Value)
                                    Else
                                        Dim UOMFoundFlag As String = ""
                                        Dim chkIfExsist As String = "Select Count(*) From ITM9 Where ""ItemCode"" = '" & oRsGetItemPriceDetails.Fields.Item(0).Value & "' And ""UomEntry"" = '" & UomEntry & "'"
                                        Dim oRschkIfExsist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRschkIfExsist.DoQuery(chkIfExsist)
                                        If oRschkIfExsist.Fields.Item(0).Value > 0 Then
                                            For k As Integer = 0 To oItem.PriceList.UoMPrices.Count - 1
                                                oItem.PriceList.UoMPrices.SetCurrentLine(k)
                                                Dim iUom As Integer = oItem.PriceList.UoMPrices.UoMEntry
                                                If iUom = UomEntry Then
                                                    oItem.PriceList.UoMPrices.Price = CDbl(oRsGetItemPriceLineDtls.Fields.Item("New Price").Value)
                                                    UOMFoundFlag = "Yes"
                                                End If
                                            Next
                                        End If
                                        If UOMFoundFlag = "" Then
                                            Dim GetUomPriceListCount As String = "Select Count(*) From ITM9 Where ""ItemCode"" = '" & oRsGetItemPriceDetails.Fields.Item(0).Value & "' And ""PriceList"" = '" & strPrlist & "'"
                                            Dim oRsGetUomPriceListCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRsGetUomPriceListCount.DoQuery(GetUomPriceListCount)
                                            If oRsGetUomPriceListCount.Fields.Item(0).Value = 0 Then
                                                UomCount = 0
                                            Else
                                                UomCount = oRsGetUomPriceListCount.Fields.Item(0).Value
                                            End If
                                            oItem.PriceList.UoMPrices.Add()
                                            oItem.PriceList.UoMPrices.SetCurrentLine(UomCount)
                                            oItem.PriceList.UoMPrices.Price = CDbl(oRsGetItemPriceLineDtls.Fields.Item("New Price").Value)
                                            oItem.PriceList.UoMPrices.UoMEntry = UomEntry
                                            oItem.PriceList.UoMPrices.PriceList = strPrlist
                                            oItem.PriceList.UoMPrices.Auto = SAPbobsCOM.BoYesNoEnum.tNO
                                            ''oItem.PriceList.UoMPrices.Currency = "AED"
                                        End If
                                    End If
                                End If
                                oRsGetItemPriceLineDtls.MoveNext()
                            Next
                            Try
                                ' oItem.UserFields.Fields.Item("U_CXS_MDFD").Value = "N"
                            Catch ex As Exception
                            End Try
                            Dim RetValue As Integer = oItem.Update
                            If RetValue <> 0 Then
                                Dim Err As String = objMain.objCompany.GetLastErrorDescription.Replace("'", "")
                                Dim PriceUpdate As String = "Update ""@ME_PRICEUPD"" Set  ""U_Updated"" = 'N' , ""U_ErrorMsg"" = '" & Err & "'  Where ""DocEntry"" = '" & oRsGetItemPriceDetails.Fields.Item(1).Value & "' "
                                Dim oRsPriceUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsPriceUpdate.DoQuery(PriceUpdate)
                            Else
                                Dim PriceUpdate As String = "Update ""@ME_PRICEUPD"" Set  ""U_Updated"" = 'Y' , ""U_ErrorMsg"" = '' Where ""DocEntry"" = '" & oRsGetItemPriceDetails.Fields.Item(1).Value & "' "
                                Dim oRsPriceUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsPriceUpdate.DoQuery(PriceUpdate)
                                'Me.UpdateDocStatus(objForm.UniqueID, "Submitted")
                                objMain.objApplication.MessageBox("Price updation completed successfully...!", 1, "Ok")
                                'objMatrix = objForm.Items.Item("3").Specific
                                'oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD_C0")

                                'Comment below
                                'Dim BCD As String = ""
                                'Dim BItem As String = ""
                                'BCD = objForm.Items.Item("Item_1").Specific.Value
                                'BItem = objForm.Items.Item("5").Specific.Value
                                'If BCD <> "" Then
                                '    Try
                                '        objForm.Freeze(True)
                                '        objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                '        objForm.Items.Item("Item_1").Specific.Value = BCD
                                '        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                '    Catch ex As Exception
                                '    Finally
                                '        objForm.Freeze(False)
                                '    End Try
                                'Else
                                '    Try
                                '        objForm.Freeze(True)
                                '        objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                '        objForm.Items.Item("5").Specific.Value = BItem
                                '        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                '    Catch ex As Exception
                                '    Finally
                                '        objForm.Freeze(False)
                                '    End Try
                                'End If
                                'For A As Integer = 1 To objMatrix.RowCount
                                '    objMatrix.GetLineData(A)
                                '    oDBs_Details.SetValue("U_OLDPRC", oDBs_Details.Offset, objMatrix.Columns.Item("V_8").Cells.Item(A).Specific.Value)
                                '    oDBs_Details.SetValue("U_NEWPRC", oDBs_Details.Offset, "0")
                                '    objMatrix.SetLineData(A)
                                'Next
                                'FDoNum = objForm.Items.Item("10").Specific.Value
                                'Try

                                '        objForm.Freeze(True)
                                '        objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                '        objForm.Items.Item("10").Specific.value = FDoNum
                                '        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                '    objForm.Freeze(False)

                                'Catch ex As Exception

                                'End Try
                                ''objForm.Refresh()

                                ''Me.LoadBranchWisePrices1(objForm.UniqueID)

                                'objMain.objApplication.StatusBar.SetText("Document submitted successfully !!.... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            End If
                        End If
                    End If
                    oRsGetItemPriceDetails.MoveNext()
                Next

                ' objForm.Items.Item("13").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                objMain.objApplication.ActivateMenuItem("1304")

                'objForm.Freeze(True)
                'objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                'objForm.Freeze(False)
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
            'Finally
            '    objMain.objApplication.MessageBox("Operation completed successfully...!", 1, "Done")
        End Try
    End Sub
    Public Sub FuturePriceUpdate(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            BaseItemCode = objForm.Items.Item("5").Specific.Value
            If BaseItemCode <> "" Then
                Dim FutPrice As String = "  SELECT T1.""U_UPDT"" AS ""LDT"",*  FROM ""@ME_PRICEUPD""  T0  INNER JOIN ""@ME_PRICEUPD_C0""  T1 ON T0.""DocEntry""=T1.""DocEntry"" WHERE T0.""U_ITMCD""  = '" & BaseItemCode & "' "
                Dim oRsFutPrice As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsFutPrice.DoQuery(FutPrice)
                If oRsFutPrice.RecordCount > 0 Then
                    Dim ret As Long
                    Dim oGeneralService As SAPbobsCOM.GeneralService
                    Dim oGeneralData As SAPbobsCOM.GeneralData
                    Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
                    Dim oCompanyService As SAPbobsCOM.CompanyService
                    Dim oChild As SAPbobsCOM.GeneralData
                    Dim oChildren As SAPbobsCOM.GeneralDataCollection
                    oCompanyService = objMain.objCompany.GetCompanyService()
                    oGeneralService = oCompanyService.GetGeneralService("ME_OFPRICEUPD")
                    oGeneralData = ((oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)))
                    oGeneralData.SetProperty("U_BARCD", oRsFutPrice.Fields.Item("U_BARCD").Value)
                    oGeneralData.SetProperty("U_ITMCD", oRsFutPrice.Fields.Item("U_ITMCD").Value)
                    oGeneralData.SetProperty("U_ITMNM", oRsFutPrice.Fields.Item("U_ITMNM").Value)
                    oGeneralData.SetProperty("U_UNITS", oRsFutPrice.Fields.Item("U_UNITS").Value)
                    oGeneralData.SetProperty("U_USRID", oRsFutPrice.Fields.Item("U_USRID").Value)
                    oGeneralData.SetProperty("U_NEWPRC", Convert.ToDouble(oRsFutPrice.Fields.Item("U_NEWPRC").Value))
                    oChildren = oGeneralData.Child("ME_FPRICEUPD_C0")

                    oRsFutPrice.MoveFirst()
                    For i As Integer = 0 To oRsFutPrice.RecordCount - 1
                        oChild = oChildren.Add
                        ''oChild.SetProperty("LineId", oRsFutPrice.Fields.Item("LineId").Value)
                        oChild.SetProperty("U_ITMCD", oRsFutPrice.Fields.Item("U_ITMCD").Value)
                        oChild.SetProperty("U_BRCHID", oRsFutPrice.Fields.Item("U_BRCHID").Value)
                        oChild.SetProperty("U_BRCHNM", oRsFutPrice.Fields.Item("U_BRCHNM").Value)
                        oChild.SetProperty("U_PRCID", oRsFutPrice.Fields.Item("U_PRCID").Value)
                        oChild.SetProperty("U_PRCLST", oRsFutPrice.Fields.Item("U_PRCLST").Value)
                        oChild.SetProperty("U_UOMENT", oRsFutPrice.Fields.Item("U_UOMENT").Value)
                        oChild.SetProperty("U_UOMCD", oRsFutPrice.Fields.Item("U_UOMCD").Value)
                        oChild.SetProperty("U_OLDPRC", oRsFutPrice.Fields.Item("U_OLDPRC").Value)
                        oChild.SetProperty("U_NEWPRC", oRsFutPrice.Fields.Item("U_NEWPRC").Value)
                        oChild.SetProperty("U_UPDT", oRsFutPrice.Fields.Item("LDT").Value)
                        oChild.SetProperty("U_ENDDT", oRsFutPrice.Fields.Item("U_ENDDT").Value)
                        oChild.SetProperty("U_USRCD", oRsFutPrice.Fields.Item("U_USRCD").Value)
                        oRsFutPrice.MoveNext()
                    Next

                    oGeneralParams = oGeneralService.Add(oGeneralData)

                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Sub UpdateItemMaster(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD")

            'Dim oItems As SAPbobsCOM.Items = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

            Dim UpdateOITM As String = " Update OITM Set ""U_PRCUPNO"" = '" & oDBs_Head.GetValue("DocNum", 0).ToString & "' Where ""ItemCode"" = '" & objForm.Items.Item("5").Specific.Value & "' "
            Dim oRsUpdateOITM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsUpdateOITM.DoQuery(UpdateOITM)
            'If oItems.GetByKey(objForm.Items.Item("5").Specific.Value) Then
            '    oItems.UserFields.Fields.Item("U_PRCUPNO").Value = oDBs_Head.GetValue("DocNum", 0).ToString
            '    oItems.Update()
            'End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Sub LoadBranchWisePrices(ByVal FormUID As String, Optional ByVal UnitsIn As String = "")
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD_C0")
            objMatrix = objForm.Items.Item("3").Specific

            Dim GetBranchList As String = "Select * FROM ( Select ""ItemCode"" , ""BranchId"" ,""BranchName"" , ""PricListId"" , ""PriceList"" , ""UomEntry"" , ""UomCode"" ,""Packing"", Sum (""OldPrice"") As ""OldPrice""  " &
                                          " From (Select AA.""ItemCode"" , T1.""BPLId"" As ""BranchId"" , T1.""BPLName"" As ""BranchName"" ,  T1.""U_PricelistNo""  As  ""PricListId"" , " &
                                          " (Select ""ListName"" From OPLN  Where ""ListNum"" =  T1.""U_PricelistNo"" ) As ""PriceList"" , AA.""UomEntry"" ,  AA.""UomCode"",AA.""Packing""," &
                                          " 0 As ""OldPrice""  From (Select T0.""ItemCode"" , C.""UomCode"" , C.""UomEntry"" ,Cast(Cast(B.""BaseQty"" As Int) As Varchar)||'*'|| Cast(Cast(B.""AltQty"" As Int) As Varchar)as ""Packing"", ROW_NUMBER() OVER (PARTITION BY B.""UomEntry"" ORDER BY B.""UomEntry"" DESC) AS UGP1RN  " &
                                          " From OITM T0 Inner Join OUGP A On A.""UgpEntry"" = T0.""UgpEntry""  Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry""  Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry""  Where T0.""ItemCode"" = '" & objForm.Items.Item("5").Specific.Value.Trim & "'   ) AA  " &
                                          " Inner Join  (Select   ROW_NUMBER() OVER (PARTITION BY ""BPLId"" ORDER BY ""BPLId"" DESC) AS BPLRN , * From OBPL  Where ""U_PricelistNo"" <> '') T1 On AA.""UGP1RN"" = ""BPLRN""  " &
                                          "            Union All " &
                                          " Select T0.""ItemCode"" ,  T2.""BPLId"" As ""BranchId"" ,  T2.""BPLName"" As ""BranchName"" , T2.""U_PricelistNo"" As  ""PricListId"" ," &
                                          " (Select ""ListName"" From OPLN  Where ""ListNum"" =  T2.""U_PricelistNo"") As ""PriceList"" ,   B.""UomEntry"" , C.""UomCode"",Cast(Cast(B.""BaseQty"" As Int) As Varchar)||'*'|| Cast(Cast(B.""AltQty"" As Int) As Varchar)as ""Packing""," &
                                          " T1.""Price""  As ""OldPrice""   From OITM T0 Inner Join OUGP A On A.""UgpEntry"" = T0.""UgpEntry""  Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry"" " &
                                          " Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry""  Inner Join ITM1 T1 On T1.""ItemCode"" = T0.""ItemCode""   And  T1.""UomEntry"" = T0.""IUoMEntry""   " &
                                          " Inner Join OBPL T2 On  T2.""U_PricelistNo"" = T1.""PriceList""  Where T0.""ItemCode"" = '" & objForm.Items.Item("5").Specific.Value.Trim & "'    And  B.""UomEntry"" =  T0.""IUoMEntry""   " &
                                          "            Union All " &
                                          " Select T0.""ItemCode"" , T2.""BPLId"" As ""BranchId"" ,  T2.""BPLName"" As ""BranchName"" , T2.""U_PricelistNo"" As  ""PricListId"" , " &
                                          " (Select ""ListName"" From OPLN  Where ""ListNum"" =  T2.""U_PricelistNo"") As ""PriceList"" ,   B.""UomEntry"" , C.""UomCode"",Cast(Cast(B.""BaseQty"" As Int) As Varchar)||'*'|| Cast(Cast(B.""AltQty"" As Int) As Varchar)as ""Packing""," &
                                          "  T3.""Price""  As ""OldPrice""   From OITM T0 Inner Join OUGP A On A.""UgpEntry"" = T0.""UgpEntry""  Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry"" " &
                                          " Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry""    Inner Join ITM9 T3 On T3.""ItemCode"" = T0.""ItemCode"" And  T3.""UomEntry"" = B.""UomEntry"" " &
                                          " Inner Join OBPL T2 On  T2.""U_PricelistNo"" = T3.""PriceList""       Where T0.""ItemCode"" = '" & objForm.Items.Item("5").Specific.Value.Trim & "') ""CompleteQty""   " &
                                          " Group By ""ItemCode"" , ""BranchId"" ,""BranchName"" , ""PricListId"" , ""PriceList"" ,  ""UomEntry"" , ""UomCode"",""Packing"")T "


            ' " Group By ""ItemCode"" , ""BranchId"" ,""BranchName"" , ""PricListId"" , ""PriceList"" ,  ""UomEntry"" , ""UomCode"",""Packing"") Where ""BranchId""<>'1' "

            If CStr(objForm.Items.Item("22").Specific.Value).Trim <> "" Then
                GetBranchList = GetBranchList & " Where  T.""UomCode""='" & CStr(objForm.Items.Item("22").Specific.Value).Trim & "'  Order By T.""BranchId"",""UomCode"""

            Else
                GetBranchList = GetBranchList & " Order By ""BranchId"",""UomCode"""

            End If

            Dim oRsGetBranchList As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetBranchList.DoQuery(GetBranchList)
            If oRsGetBranchList.RecordCount > 0 Then
                'If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                '    Me.UpdateDocStatus(objForm.UniqueID, "Open")
                'End If
                objMatrix.Clear()
                oDBs_Details.Clear()
                objMatrix.FlushToDataSource()
                For i As Integer = 1 To oRsGetBranchList.RecordCount
                    objMatrix.AddRow()
                    oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
                    oDBs_Details.SetValue("U_ITMCD", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("ItemCode").Value)
                    oDBs_Details.SetValue("U_BRCHID", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("BranchId").Value)
                    oDBs_Details.SetValue("U_BRCHNM", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("BranchName").Value)
                    oDBs_Details.SetValue("U_PRCID", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("PricListId").Value)
                    oDBs_Details.SetValue("U_PRCLST", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("PriceList").Value)
                    oDBs_Details.SetValue("U_Packing", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("Packing").Value)
                    oDBs_Details.SetValue("U_UOMENT", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("UomEntry").Value)
                    oDBs_Details.SetValue("U_UOMCD", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("UomCode").Value)
                    oDBs_Details.SetValue("U_OLDPRC", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("OldPrice").Value)
                    Dim CheckUpDate As String = " Select ""U_UPDT"" , ""U_ENDDT"" , ""U_USRCD"" From ""@ME_PRICEUPD_C0"" A Where A.""U_ITMCD"" = '" & oRsGetBranchList.Fields.Item("ItemCode").Value & "' " &
                                                " And ""U_BRCHID"" = '" & oRsGetBranchList.Fields.Item("BranchId").Value & "' And " &
                                                " ""U_PRCID""='" & oRsGetBranchList.Fields.Item("PricListId").Value & "' And ""U_UOMCD"" = '" & oRsGetBranchList.Fields.Item("UomCode").Value & "' "
                    Dim oRsCheckUpDate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsCheckUpDate.DoQuery(CheckUpDate)

                    Dim iDate As String = oRsCheckUpDate.Fields.Item("U_ENDDT").Value
                    Dim oDate As DateTime = Convert.ToDateTime(iDate)


                    If oDate.Year = "1899" Or oDate.Year = "1999" Or oDate.Year = "9999" Then
                        oDBs_Details.SetValue("U_ENDDT", oDBs_Details.Offset, "20401231")
                    Else
                        If CDate(oRsCheckUpDate.Fields.Item("U_ENDDT").Value).ToString("yyyyMMdd") <> "18991230" And CDate(oRsCheckUpDate.Fields.Item("U_ENDDT").Value).ToString("yyyyMMdd") <> "19991230" And oRsCheckUpDate.Fields.Item("U_ENDDT").Value.ToString <> "" Then
                            oDBs_Details.SetValue("U_ENDDT", oDBs_Details.Offset, CDate(oRsCheckUpDate.Fields.Item("U_ENDDT").Value).ToString("yyyyMMdd"))
                        Else
                            oDBs_Details.SetValue("U_ENDDT", oDBs_Details.Offset, "20401231")
                        End If
                    End If

                    If CDate(oRsCheckUpDate.Fields.Item("U_UPDT").Value).ToString("yyyyMMdd") <> "18991230" And CDate(oRsCheckUpDate.Fields.Item("U_UPDT").Value).ToString("yyyyMMdd") <> "19991230" And oRsCheckUpDate.Fields.Item("U_UPDT").Value.ToString <> "" Then
                        oDBs_Details.SetValue("U_UPDT", oDBs_Details.Offset, CDate(oRsCheckUpDate.Fields.Item("U_UPDT").Value).ToString("yyyyMMdd"))
                    Else
                        oDBs_Details.SetValue("U_UPDT", oDBs_Details.Offset, "20211101") 'DateTime.Now.ToString("yyyyMMdd"))
                    End If


                    If oRsCheckUpDate.Fields.Item("U_USRCD").Value <> "" Then
                        oDBs_Details.SetValue("U_USRCD", oDBs_Details.Offset, oRsCheckUpDate.Fields.Item("U_USRCD").Value.ToString)
                    Else
                        oDBs_Details.SetValue("U_USRCD", oDBs_Details.Offset, "")
                    End If
                    objMatrix.SetLineData(objMatrix.VisualRowCount)
                    oRsGetBranchList.MoveNext()
                Next
                objMatrix.AutoResizeColumns()
                objForm.Refresh()
            Else
                'If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                '    Me.UpdateDocStatus(objForm.UniqueID, "Open")
                'End If
                'GetBranchList = " Select ""ItemCode"" , ""BranchId"" ,""BranchName"" , ""PricListId"" , ""PriceList"" , ""UomEntry"" , ""UomCode"" , Sum (""OldPrice"") As ""OldPrice""  " &
                '                          " From (Select AA.""ItemCode"" , T1.""BPLId"" As ""BranchId"" , T1.""BPLName"" As ""BranchName"" ,  T1.""U_PricelistNo""  As  ""PricListId"" , " &
                '                          " (Select ""ListName"" From OPLN  Where ""ListNum"" =  T1.""U_PricelistNo"" ) As ""PriceList"" , AA.""UomEntry"" ,  AA.""UomCode""  ,  " &
                '                          " 0 As ""OldPrice""  From (Select T0.""ItemCode"" , C.""UomCode"" , C.""UomEntry"" , ROW_NUMBER() OVER (PARTITION BY B.""UomEntry"" ORDER BY B.""UomEntry"" DESC) AS UGP1RN  " &
                '                          " From OITM T0 Inner Join OUGP A On A.""UgpEntry"" = T0.""UgpEntry""  Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry""  Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry""  Where T0.""CodeBars"" = '" & objForm.Items.Item("Item_1").Specific.Value.Trim & "'   ) AA  " &
                '                          " Inner Join  (Select   ROW_NUMBER() OVER (PARTITION BY ""BPLId"" ORDER BY ""BPLId"" DESC) AS BPLRN , * From OBPL  Where ""U_PricelistNo"" <> '') T1 On AA.""UGP1RN"" = ""BPLRN""  " &
                '                          "            Union All " &
                '                          " Select T0.""ItemCode"" ,  T2.""BPLId"" As ""BranchId"" ,  T2.""BPLName"" As ""BranchName"" , T2.""U_PricelistNo"" As  ""PricListId"" ," &
                '                          " (Select ""ListName"" From OPLN  Where ""ListNum"" =  T2.""U_PricelistNo"") As ""PriceList"" ,   B.""UomEntry"" , C.""UomCode"" , " &
                '                          " T1.""Price""  As ""OldPrice""   From OITM T0 Inner Join OUGP A On A.""UgpEntry"" = T0.""UgpEntry""  Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry"" " &
                '                          " Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry""  Inner Join ITM1 T1 On T1.""ItemCode"" = T0.""ItemCode""   And  T1.""UomEntry"" = T0.""IUoMEntry""   " &
                '                          " Inner Join OBPL T2 On  T2.""U_PricelistNo"" = T1.""PriceList""  Where T0.""CodeBars"" = '" & objForm.Items.Item("Item_1").Specific.Value.Trim & "'    And  B.""UomEntry"" =  T0.""IUoMEntry""   " &
                '                          "            Union All " &
                '                          " Select T0.""ItemCode"" , T2.""BPLId"" As ""BranchId"" ,  T2.""BPLName"" As ""BranchName"" , T2.""U_PricelistNo"" As  ""PricListId"" , " &
                '                          " (Select ""ListName"" From OPLN  Where ""ListNum"" =  T2.""U_PricelistNo"") As ""PriceList"" ,   B.""UomEntry"" , C.""UomCode"" , " &
                '                          "  T3.""Price""  As ""OldPrice""   From OITM T0 Inner Join OUGP A On A.""UgpEntry"" = T0.""UgpEntry""  Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry"" " &
                '                          " Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry""    Inner Join ITM9 T3 On T3.""ItemCode"" = T0.""ItemCode"" And  T3.""UomEntry"" = B.""UomEntry"" " &
                '                          " Inner Join OBPL T2 On  T2.""U_PricelistNo"" = T3.""PriceList""       Where T0.""CodeBars"" = '" & objForm.Items.Item("Item_1").Specific.Value.Trim & "') ""CompleteQty""   " &
                '                          " Group By ""ItemCode"" , ""BranchId"" ,""BranchName"" , ""PricListId"" , ""PriceList"" ,  ""UomEntry"" , ""UomCode"" "
                GetBranchList = " Select ""ItemCode"" , ""BranchId"" ,""BranchName"" , ""PricListId"" , ""PriceList"" , ""UomEntry"" , ""UomCode"" ,""Packing"", Sum (""OldPrice"") As ""OldPrice""  " &
                                         " From (Select AA.""ItemCode"" , T1.""BPLId"" As ""BranchId"" , T1.""BPLName"" As ""BranchName"" ,  T1.""U_PricelistNo""  As  ""PricListId"" , " &
                                         " (Select ""ListName"" From OPLN  Where ""ListNum"" =  T1.""U_PricelistNo"" ) As ""PriceList"" , AA.""UomEntry"" ,  AA.""UomCode"",AA.""Packing""," &
                                         " 0 As ""OldPrice""  From (Select T0.""ItemCode"" , C.""UomCode"" , C.""UomEntry"" ,Cast(Cast(B.""BaseQty"" As Int) As Varchar)||'*'|| Cast(Cast(B.""AltQty"" As Int) As Varchar)as ""Packing"", ROW_NUMBER() OVER (PARTITION BY B.""UomEntry"" ORDER BY B.""UomEntry"" DESC) AS UGP1RN  " &
                                         " From OITM T0 Inner Join OUGP A On A.""UgpEntry"" = T0.""UgpEntry""  Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry""  Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry""  Where T0.""ItemCode"" = '" & objForm.Items.Item("5").Specific.Value.Trim & "'   ) AA  " &
                                         " Inner Join  (Select   ROW_NUMBER() OVER (PARTITION BY ""BPLId"" ORDER BY ""BPLId"" DESC) AS BPLRN , * From OBPL  Where ""U_PricelistNo"" <> '') T1 On AA.""UGP1RN"" = ""BPLRN""  " &
                                         "            Union All " &
                                         " Select T0.""ItemCode"" ,  T2.""BPLId"" As ""BranchId"" ,  T2.""BPLName"" As ""BranchName"" , T2.""U_PricelistNo"" As  ""PricListId"" ," &
                                         " (Select ""ListName"" From OPLN  Where ""ListNum"" =  T2.""U_PricelistNo"") As ""PriceList"" ,   B.""UomEntry"" , C.""UomCode"",Cast(Cast(B.""BaseQty"" As Int) As Varchar)||'*'|| Cast(Cast(B.""AltQty"" As Int) As Varchar)as ""Packing""," &
                                         " T1.""Price""  As ""OldPrice""   From OITM T0 Inner Join OUGP A On A.""UgpEntry"" = T0.""UgpEntry""  Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry"" " &
                                         " Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry""  Inner Join ITM1 T1 On T1.""ItemCode"" = T0.""ItemCode""   And  T1.""UomEntry"" = T0.""IUoMEntry""   " &
                                         " Inner Join OBPL T2 On  T2.""U_PricelistNo"" = T1.""PriceList""  Where T0.""ItemCode"" = '" & objForm.Items.Item("5").Specific.Value.Trim & "'    And  B.""UomEntry"" =  T0.""IUoMEntry""   " &
                                         "            Union All " &
                                         " Select T0.""ItemCode"" , T2.""BPLId"" As ""BranchId"" ,  T2.""BPLName"" As ""BranchName"" , T2.""U_PricelistNo"" As  ""PricListId"" , " &
                                         " (Select ""ListName"" From OPLN  Where ""ListNum"" =  T2.""U_PricelistNo"") As ""PriceList"" ,   B.""UomEntry"" , C.""UomCode"",Cast(Cast(B.""BaseQty"" As Int) As Varchar)||'*'|| Cast(Cast(B.""AltQty"" As Int) As Varchar)as ""Packing""," &
                                         "  T3.""Price""  As ""OldPrice""   From OITM T0 Inner Join OUGP A On A.""UgpEntry"" = T0.""UgpEntry""  Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry"" " &
                                         " Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry""    Inner Join ITM9 T3 On T3.""ItemCode"" = T0.""ItemCode"" And  T3.""UomEntry"" = B.""UomEntry"" " &
                                         " Inner Join OBPL T2 On  T2.""U_PricelistNo"" = T3.""PriceList""       Where T0.""ItemCode"" = '" & objForm.Items.Item("5").Specific.Value.Trim & "') ""CompleteQty""   " &
                                         " Group By ""ItemCode"" , ""BranchId"" ,""BranchName"" , ""PricListId"" , ""PriceList"" ,  ""UomEntry"" , ""UomCode"",""Packing"" "



                oRsGetBranchList = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetBranchList.DoQuery(GetBranchList)
                objMatrix.Clear()
                oDBs_Details.Clear()
                objMatrix.FlushToDataSource()
                For i As Integer = 1 To oRsGetBranchList.RecordCount
                    objMatrix.AddRow()
                    oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
                    oDBs_Details.SetValue("U_ITMCD", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("ItemCode").Value)
                    oDBs_Details.SetValue("U_BRCHID", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("BranchId").Value)
                    oDBs_Details.SetValue("U_BRCHNM", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("BranchName").Value)
                    oDBs_Details.SetValue("U_PRCID", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("PricListId").Value)
                    oDBs_Details.SetValue("U_PRCLST", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("PriceList").Value)
                    oDBs_Details.SetValue("U_UOMENT", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("UomEntry").Value)
                    oDBs_Details.SetValue("U_Packing", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("Packing").Value)
                    oDBs_Details.SetValue("U_UOMCD", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("UomCode").Value)
                    oDBs_Details.SetValue("U_OLDPRC", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("OldPrice").Value)
                    Dim CheckUpDate As String = " Select ""U_UPDT"" , ""U_ENDDT"" , ""U_USRCD"" From ""@ME_PRICEUPD_C0"" A Where A.""U_ITMCD"" = '" & oRsGetBranchList.Fields.Item("ItemCode").Value & "' " &
                                                " And ""U_BRCHID"" = '" & oRsGetBranchList.Fields.Item("BranchId").Value & "' And " &
                                                " ""U_PRCID""='" & oRsGetBranchList.Fields.Item("PricListId").Value & "' And ""U_UOMCD"" = '" & oRsGetBranchList.Fields.Item("UomCode").Value & "' "
                    Dim oRsCheckUpDate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsCheckUpDate.DoQuery(CheckUpDate)
                    If CDate(oRsCheckUpDate.Fields.Item("U_UPDT").Value).ToString("yyyyMMdd") <> "18991230" And CDate(oRsCheckUpDate.Fields.Item("U_UPDT").Value).ToString("yyyyMMdd") <> "19991230" And oRsCheckUpDate.Fields.Item("U_UPDT").Value.ToString <> "" Then
                        oDBs_Details.SetValue("U_UPDT", oDBs_Details.Offset, CDate(oRsCheckUpDate.Fields.Item("U_UPDT").Value).ToString("yyyyMMdd"))
                    Else
                        oDBs_Details.SetValue("U_UPDT", oDBs_Details.Offset, "20211101") 'DateTime.Now.ToString("yyyyMMdd"))
                    End If
                    If CDate(oRsCheckUpDate.Fields.Item("U_ENDDT").Value).ToString("yyyyMMdd") <> "18991230" And CDate(oRsCheckUpDate.Fields.Item("U_ENDDT").Value).ToString("yyyyMMdd") <> "19991230" And oRsCheckUpDate.Fields.Item("U_ENDDT").Value.ToString <> "" Then
                        oDBs_Details.SetValue("U_ENDDT", oDBs_Details.Offset, CDate(oRsCheckUpDate.Fields.Item("U_ENDDT").Value).ToString("yyyyMMdd"))
                    Else
                        oDBs_Details.SetValue("U_ENDDT", oDBs_Details.Offset, "20401231")
                    End If
                    If oRsCheckUpDate.Fields.Item("U_USRCD").Value <> "" Then
                        oDBs_Details.SetValue("U_USRCD", oDBs_Details.Offset, oRsCheckUpDate.Fields.Item("U_USRCD").Value.ToString)
                    Else
                        oDBs_Details.SetValue("U_USRCD", oDBs_Details.Offset, "")
                    End If
                    objMatrix.SetLineData(objMatrix.VisualRowCount)
                    oRsGetBranchList.MoveNext()
                Next
                ''objMatrix.AutoResizeColumns()
            End If
            ''objForm.Refresh()

            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Sub LoadBranchWisePrices1(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD_C0")
            objMatrix = objForm.Items.Item("3").Specific

            Dim GetBranchList As String = " Select ""ItemCode"" , ""BranchId"" ,""BranchName"" , ""PricListId"" , ""PriceList"" , ""UomEntry"" , ""UomCode"" , Sum (""OldPrice"") As ""OldPrice""  " &
                                          " From (Select AA.""ItemCode"" , T1.""BPLId"" As ""BranchId"" , T1.""BPLName"" As ""BranchName"" ,  T1.""U_PricelistNo""  As  ""PricListId"" , " &
                                          " (Select ""ListName"" From OPLN  Where ""ListNum"" =  T1.""U_PricelistNo"" ) As ""PriceList"" , AA.""UomEntry"" ,  AA.""UomCode""  ,  " &
                                          " 0 As ""OldPrice""  From (Select T0.""ItemCode"" , C.""UomCode"" , C.""UomEntry"" , ROW_NUMBER() OVER (PARTITION BY B.""UomEntry"" ORDER BY B.""UomEntry"" DESC) AS UGP1RN  " &
                                          " From OITM T0 Inner Join OUGP A On A.""UgpEntry"" = T0.""UgpEntry""  Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry""  Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry""  Where T0.""ItemCode"" = '" & objForm.Items.Item("5").Specific.Value.Trim & "'   ) AA  " &
                                          " Inner Join  (Select   ROW_NUMBER() OVER (PARTITION BY ""BPLId"" ORDER BY ""BPLId"" DESC) AS BPLRN , * From OBPL  Where ""U_PricelistNo"" <> '') T1 On AA.""UGP1RN"" = ""BPLRN""  " &
                                          "            Union All " &
                                          " Select T0.""ItemCode"" ,  T2.""BPLId"" As ""BranchId"" ,  T2.""BPLName"" As ""BranchName"" , T2.""U_PricelistNo"" As  ""PricListId"" ," &
                                          " (Select ""ListName"" From OPLN  Where ""ListNum"" =  T2.""U_PricelistNo"") As ""PriceList"" ,   B.""UomEntry"" , C.""UomCode"" , " &
                                          " T1.""Price""  As ""OldPrice""   From OITM T0 Inner Join OUGP A On A.""UgpEntry"" = T0.""UgpEntry""  Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry"" " &
                                          " Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry""  Inner Join ITM1 T1 On T1.""ItemCode"" = T0.""ItemCode""   And  T1.""UomEntry"" = T0.""IUoMEntry""   " &
                                          " Inner Join OBPL T2 On  T2.""U_PricelistNo"" = T1.""PriceList""  Where T0.""ItemCode"" = '" & objForm.Items.Item("5").Specific.Value.Trim & "'    And  B.""UomEntry"" =  T0.""IUoMEntry""   " &
                                          "            Union All " &
                                          " Select T0.""ItemCode"" , T2.""BPLId"" As ""BranchId"" ,  T2.""BPLName"" As ""BranchName"" , T2.""U_PricelistNo"" As  ""PricListId"" , " &
                                          " (Select ""ListName"" From OPLN  Where ""ListNum"" =  T2.""U_PricelistNo"") As ""PriceList"" ,   B.""UomEntry"" , C.""UomCode"" , " &
                                          "  T3.""Price""  As ""OldPrice""   From OITM T0 Inner Join OUGP A On A.""UgpEntry"" = T0.""UgpEntry""  Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry"" " &
                                          " Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry""    Inner Join ITM9 T3 On T3.""ItemCode"" = T0.""ItemCode"" And  T3.""UomEntry"" = B.""UomEntry"" " &
                                          " Inner Join OBPL T2 On  T2.""U_PricelistNo"" = T3.""PriceList""       Where T0.""ItemCode"" = '" & objForm.Items.Item("5").Specific.Value.Trim & "') ""CompleteQty""   " &
                                          " Group By ""ItemCode"" , ""BranchId"" ,""BranchName"" , ""PricListId"" , ""PriceList"" ,  ""UomEntry"" , ""UomCode"" "
            Dim oRsGetBranchList As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetBranchList.DoQuery(GetBranchList)
            If oRsGetBranchList.RecordCount > 0 Then
                'If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                '    Me.UpdateDocStatus(objForm.UniqueID, "Open")
                'End If
                objMatrix.Clear()
                oDBs_Details.Clear()
                objMatrix.FlushToDataSource()
                For i As Integer = 1 To oRsGetBranchList.RecordCount
                    objMatrix.AddRow()
                    oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
                    oDBs_Details.SetValue("U_ITMCD", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("ItemCode").Value)
                    oDBs_Details.SetValue("U_BRCHID", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("BranchId").Value)
                    oDBs_Details.SetValue("U_BRCHNM", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("BranchName").Value)
                    oDBs_Details.SetValue("U_PRCID", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("PricListId").Value)
                    oDBs_Details.SetValue("U_PRCLST", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("PriceList").Value)
                    oDBs_Details.SetValue("U_UOMENT", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("UomEntry").Value)
                    oDBs_Details.SetValue("U_UOMCD", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("UomCode").Value)
                    oDBs_Details.SetValue("U_OLDPRC", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("OldPrice").Value)
                    Dim CheckUpDate As String = " Select ""U_UPDT"" , ""U_ENDDT"" , ""U_USRCD"" From ""@ME_PRICEUPD_C0"" A Where A.""U_ITMCD"" = '" & oRsGetBranchList.Fields.Item("ItemCode").Value & "' " &
                                                " And ""U_BRCHID"" = '" & oRsGetBranchList.Fields.Item("BranchId").Value & "' And " &
                                                " ""U_PRCID""='" & oRsGetBranchList.Fields.Item("PricListId").Value & "' And ""U_UOMCD"" = '" & oRsGetBranchList.Fields.Item("UomCode").Value & "' "
                    Dim oRsCheckUpDate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsCheckUpDate.DoQuery(CheckUpDate)

                    Dim iDate As String = oRsCheckUpDate.Fields.Item("U_ENDDT").Value
                    Dim oDate As DateTime = Convert.ToDateTime(iDate)


                    If oDate.Year = "1899" Or oDate.Year = "1999" Or oDate.Year = "9999" Then
                        oDBs_Details.SetValue("U_ENDDT", oDBs_Details.Offset, "20401231")
                    Else
                        If CDate(oRsCheckUpDate.Fields.Item("U_ENDDT").Value).ToString("yyyyMMdd") <> "18991230" And CDate(oRsCheckUpDate.Fields.Item("U_ENDDT").Value).ToString("yyyyMMdd") <> "19991230" And oRsCheckUpDate.Fields.Item("U_ENDDT").Value.ToString <> "" Then
                            oDBs_Details.SetValue("U_ENDDT", oDBs_Details.Offset, CDate(oRsCheckUpDate.Fields.Item("U_ENDDT").Value).ToString("yyyyMMdd"))
                        Else
                            oDBs_Details.SetValue("U_ENDDT", oDBs_Details.Offset, "20401231")
                        End If
                    End If

                    If CDate(oRsCheckUpDate.Fields.Item("U_UPDT").Value).ToString("yyyyMMdd") <> "18991230" And CDate(oRsCheckUpDate.Fields.Item("U_UPDT").Value).ToString("yyyyMMdd") <> "19991230" And oRsCheckUpDate.Fields.Item("U_UPDT").Value.ToString <> "" Then
                        oDBs_Details.SetValue("U_UPDT", oDBs_Details.Offset, CDate(oRsCheckUpDate.Fields.Item("U_UPDT").Value).ToString("yyyyMMdd"))
                    Else
                        oDBs_Details.SetValue("U_UPDT", oDBs_Details.Offset, "20211101") 'DateTime.Now.ToString("yyyyMMdd"))
                    End If


                    If oRsCheckUpDate.Fields.Item("U_USRCD").Value <> "" Then
                        oDBs_Details.SetValue("U_USRCD", oDBs_Details.Offset, oRsCheckUpDate.Fields.Item("U_USRCD").Value.ToString)
                    Else
                        oDBs_Details.SetValue("U_USRCD", oDBs_Details.Offset, "")
                    End If
                    objMatrix.SetLineData(objMatrix.VisualRowCount)
                    oRsGetBranchList.MoveNext()
                Next
                'objMatrix.AutoResizeColumns()
                'objForm.Refresh()
            Else
                'If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                '    Me.UpdateDocStatus(objForm.UniqueID, "Open")
                'End If
                GetBranchList = " Select ""ItemCode"" , ""BranchId"" ,""BranchName"" , ""PricListId"" , ""PriceList"" , ""UomEntry"" , ""UomCode"" , Sum (""OldPrice"") As ""OldPrice""  " &
                                          " From (Select AA.""ItemCode"" , T1.""BPLId"" As ""BranchId"" , T1.""BPLName"" As ""BranchName"" ,  T1.""U_PricelistNo""  As  ""PricListId"" , " &
                                          " (Select ""ListName"" From OPLN  Where ""ListNum"" =  T1.""U_PricelistNo"" ) As ""PriceList"" , AA.""UomEntry"" ,  AA.""UomCode""  ,  " &
                                          " 0 As ""OldPrice""  From (Select T0.""ItemCode"" , C.""UomCode"" , C.""UomEntry"" , ROW_NUMBER() OVER (PARTITION BY B.""UomEntry"" ORDER BY B.""UomEntry"" DESC) AS UGP1RN  " &
                                          " From OITM T0 Inner Join OUGP A On A.""UgpEntry"" = T0.""UgpEntry""  Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry""  Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry""  Where T0.""CodeBars"" = '" & objForm.Items.Item("Item_1").Specific.Value.Trim & "'   ) AA  " &
                                          " Inner Join  (Select   ROW_NUMBER() OVER (PARTITION BY ""BPLId"" ORDER BY ""BPLId"" DESC) AS BPLRN , * From OBPL  Where ""U_PricelistNo"" <> '') T1 On AA.""UGP1RN"" = ""BPLRN""  " &
                                          "            Union All " &
                                          " Select T0.""ItemCode"" ,  T2.""BPLId"" As ""BranchId"" ,  T2.""BPLName"" As ""BranchName"" , T2.""U_PricelistNo"" As  ""PricListId"" ," &
                                          " (Select ""ListName"" From OPLN  Where ""ListNum"" =  T2.""U_PricelistNo"") As ""PriceList"" ,   B.""UomEntry"" , C.""UomCode"" , " &
                                          " T1.""Price""  As ""OldPrice""   From OITM T0 Inner Join OUGP A On A.""UgpEntry"" = T0.""UgpEntry""  Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry"" " &
                                          " Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry""  Inner Join ITM1 T1 On T1.""ItemCode"" = T0.""ItemCode""   And  T1.""UomEntry"" = T0.""IUoMEntry""   " &
                                          " Inner Join OBPL T2 On  T2.""U_PricelistNo"" = T1.""PriceList""  Where T0.""CodeBars"" = '" & objForm.Items.Item("Item_1").Specific.Value.Trim & "'    And  B.""UomEntry"" =  T0.""IUoMEntry""   " &
                                          "            Union All " &
                                          " Select T0.""ItemCode"" , T2.""BPLId"" As ""BranchId"" ,  T2.""BPLName"" As ""BranchName"" , T2.""U_PricelistNo"" As  ""PricListId"" , " &
                                          " (Select ""ListName"" From OPLN  Where ""ListNum"" =  T2.""U_PricelistNo"") As ""PriceList"" ,   B.""UomEntry"" , C.""UomCode"" , " &
                                          "  T3.""Price""  As ""OldPrice""   From OITM T0 Inner Join OUGP A On A.""UgpEntry"" = T0.""UgpEntry""  Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry"" " &
                                          " Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry""    Inner Join ITM9 T3 On T3.""ItemCode"" = T0.""ItemCode"" And  T3.""UomEntry"" = B.""UomEntry"" " &
                                          " Inner Join OBPL T2 On  T2.""U_PricelistNo"" = T3.""PriceList""       Where T0.""CodeBars"" = '" & objForm.Items.Item("Item_1").Specific.Value.Trim & "') ""CompleteQty""   " &
                                          " Group By ""ItemCode"" , ""BranchId"" ,""BranchName"" , ""PricListId"" , ""PriceList"" ,  ""UomEntry"" , ""UomCode"" "
                oRsGetBranchList = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetBranchList.DoQuery(GetBranchList)
                objMatrix.Clear()
                oDBs_Details.Clear()
                objMatrix.FlushToDataSource()
                For i As Integer = 1 To oRsGetBranchList.RecordCount
                    objMatrix.AddRow()
                    oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
                    oDBs_Details.SetValue("U_ITMCD", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("ItemCode").Value)
                    oDBs_Details.SetValue("U_BRCHID", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("BranchId").Value)
                    oDBs_Details.SetValue("U_BRCHNM", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("BranchName").Value)
                    oDBs_Details.SetValue("U_PRCID", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("PricListId").Value)
                    oDBs_Details.SetValue("U_PRCLST", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("PriceList").Value)
                    oDBs_Details.SetValue("U_UOMENT", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("UomEntry").Value)
                    oDBs_Details.SetValue("U_UOMCD", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("UomCode").Value)
                    oDBs_Details.SetValue("U_OLDPRC", oDBs_Details.Offset, oRsGetBranchList.Fields.Item("OldPrice").Value)
                    Dim CheckUpDate As String = " Select ""U_UPDT"" , ""U_ENDDT"" , ""U_USRCD"" From ""@ME_PRICEUPD_C0"" A Where A.""U_ITMCD"" = '" & oRsGetBranchList.Fields.Item("ItemCode").Value & "' " &
                                                " And ""U_BRCHID"" = '" & oRsGetBranchList.Fields.Item("BranchId").Value & "' And " &
                                                " ""U_PRCID""='" & oRsGetBranchList.Fields.Item("PricListId").Value & "' And ""U_UOMCD"" = '" & oRsGetBranchList.Fields.Item("UomCode").Value & "' "
                    Dim oRsCheckUpDate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsCheckUpDate.DoQuery(CheckUpDate)
                    If CDate(oRsCheckUpDate.Fields.Item("U_UPDT").Value).ToString("yyyyMMdd") <> "18991230" And CDate(oRsCheckUpDate.Fields.Item("U_UPDT").Value).ToString("yyyyMMdd") <> "19991230" And oRsCheckUpDate.Fields.Item("U_UPDT").Value.ToString <> "" Then
                        oDBs_Details.SetValue("U_UPDT", oDBs_Details.Offset, CDate(oRsCheckUpDate.Fields.Item("U_UPDT").Value).ToString("yyyyMMdd"))
                    Else
                        oDBs_Details.SetValue("U_UPDT", oDBs_Details.Offset, "20211101") 'DateTime.Now.ToString("yyyyMMdd"))
                    End If
                    If CDate(oRsCheckUpDate.Fields.Item("U_ENDDT").Value).ToString("yyyyMMdd") <> "18991230" And CDate(oRsCheckUpDate.Fields.Item("U_ENDDT").Value).ToString("yyyyMMdd") <> "19991230" And oRsCheckUpDate.Fields.Item("U_ENDDT").Value.ToString <> "" Then
                        oDBs_Details.SetValue("U_ENDDT", oDBs_Details.Offset, CDate(oRsCheckUpDate.Fields.Item("U_ENDDT").Value).ToString("yyyyMMdd"))
                    Else
                        oDBs_Details.SetValue("U_ENDDT", oDBs_Details.Offset, "20401231")
                    End If
                    If oRsCheckUpDate.Fields.Item("U_USRCD").Value <> "" Then
                        oDBs_Details.SetValue("U_USRCD", oDBs_Details.Offset, oRsCheckUpDate.Fields.Item("U_USRCD").Value.ToString)
                    Else
                        oDBs_Details.SetValue("U_USRCD", oDBs_Details.Offset, "")
                    End If
                    objMatrix.SetLineData(objMatrix.VisualRowCount)
                    oRsGetBranchList.MoveNext()
                Next
                ''objMatrix.AutoResizeColumns()
            End If
            ''objForm.Refresh()

            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub RefreshData(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD_C0")
            objMatrix = objForm.Items.Item("3").Specific

            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD_C0"), "DocEntry", oDBs_Details.GetValue("DocEntry", 0))

            objMatrix.LoadFromDataSource()
            objMatrix.AutoResizeColumns()

            objForm.Refresh()
            objForm.Freeze(False)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Private Sub LoadVendorPriceList(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("18").Specific
            objForm.Freeze(True)
            Dim UomLine As String = ""

            oGrid.DataTable = oDt
            oGrid.DataTable.Clear()

            Dim Query As String = " Select  T0.""ItemCode"" As ""Item Code""   , T2.""CardCode"" As ""Vendor Code"" , T2.""CardName""  As ""Vendor Name"" ,  " &
                                  " T2.""ListNum"" As  ""Price List ID"" ,  (Select ""ListName"" From OPLN  Where ""ListNum"" =  T2.""ListNum"") As ""PriceList"" ,  " &
                                  " B.""UomEntry"" , C.""UomCode"" ,   (" & SelectCase & " Case When B.""UomEntry"" = T0.""IUoMEntry"" Then IFNULL(T1.""Price"",0)  " &
                                  " When B.""UomEntry"" = T3.""UomEntry""    And T3.""PriceList"" = T2.""ListNum"" Then IFNULL(T3.""Price"",0) Else 0 End) As ""OldPrice""  ,  " &
                                  " CAST ('0' AS Float) As ""New Price"" , '                                  ' As ""Status"" , '      ' As ""UomLine"" ,  " &
                                  " T0.""IUoMEntry"" As ""DefaultUOM""  From OITM T0 Inner Join OUGP A On A.""UgpEntry"" = T0.""UgpEntry""    " &
                                  " Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry""   Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry""   " &
                                  " Inner Join ITM2 T4 On T4.""ItemCode"" = T0.""ItemCode""   Inner Join OCRD T2 On T2.""CardCode"" = T4.""VendorCode""    " &
                                  " Left Join ITM1 T1 On T1.""ItemCode"" = T0.""ItemCode""  And  T1.""UomEntry"" = T0.""IUoMEntry""   " &
                                  " And T2.""ListNum"" = T1.""PriceList""     Left Join ITM9 T3 On T3.""ItemCode"" = T0.""ItemCode""  " &
                                  " And  T3.""UomEntry"" = B.""UomEntry"" And T3.""PriceList"" = T2.""ListNum""   " &
                                  " Where T0.""ItemCode"" = '" & objForm.Items.Item("5").Specific.Value.Trim & "'   "
            oGrid.DataTable.ExecuteQuery(Query)
            If oGrid.Rows.Count = 0 Then
                objMain.objApplication.MessageBox("No records found")
                Exit Sub
            End If
            For i As Integer = 0 To oGrid.Columns.Count - 1
                oGrid.Columns.Item(i).Editable = False
            Next
            'oGrid.Columns.Item("PricListId").Visible = False
            'oGrid.Columns.Item("UomEntry").Visible = False
            'oGrid.Columns.Item("DefaultUOM").Visible = False
            'oGrid.Columns.Item("UomLine").Visible = False

4:          oGrid.Columns.Item("New Price").Editable = True
            Dim oEditTextCol As SAPbouiCOM.EditTextColumn
            oEditTextCol = oGrid.Columns.Item("Item Code")
            oEditTextCol.LinkedObjectType = 4

            oEditTextCol = oGrid.Columns.Item("Vendor Code")
            oEditTextCol.LinkedObjectType = 2

            Dim FirstUomLine As String = ""

            Dim Count As Integer = 0
            For i As Integer = 0 To oGrid.Rows.Count - 1
                If oGrid.DataTable.GetValue("DefaultUOM", i) <> oGrid.DataTable.GetValue("UomEntry", i) Then
                    If FirstUomLine = "" Then
                        FirstUomLine = oGrid.DataTable.GetValue("UomEntry", i)
                    End If
                    If FirstUomLine = oGrid.DataTable.GetValue("UomEntry", i) Then
                        oGrid.DataTable.SetValue("UomLine", i, Count)
                    Else
                        Count = Count + 1
                        oGrid.DataTable.SetValue("UomLine", i, Count)
                        FirstUomLine = oGrid.DataTable.GetValue("UomEntry", i)
                    End If
                End If
            Next

            oGrid.AutoResizeColumns()
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Sub UpdateVendorPrice(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("18").Specific
            Dim UomCount As Integer = 0
            Dim r As Integer = 0
            Dim oItemMaster As SAPbobsCOM.Items = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
            For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If oGrid.DataTable.GetValue("New Price", i) > 0 Then
                    If oItemMaster.GetByKey(oGrid.DataTable.GetValue("Item Code", i).ToString) Then
H:                      Dim NewPrice As Double = CDbl(oGrid.DataTable.GetValue("New Price", i).ToString)
                        If NewPrice > 0 Then
                            Dim UomEntry As Integer = CInt(oGrid.DataTable.GetValue("Uom Entry", i).ToString)
                            Dim DefaultUom As Integer = CInt(oGrid.DataTable.GetValue("DefaultUOM", i).ToString)
                            Dim PriceListRow As Integer = oGrid.DataTable.GetValue("PricListId", i)
                            oItemMaster.PriceList.SetCurrentLine(PriceListRow - 1)
                            Dim strPrlist As String = oItemMaster.PriceList.PriceList
                            If strPrlist = oGrid.DataTable.GetValue("PricListId", i).ToString Then
                                If DefaultUom = UomEntry Then
                                    oItemMaster.PriceList.Price = CDbl(oGrid.DataTable.GetValue("New Price", i).ToString)
                                Else
                                    Dim UOMFoundFlag As String = ""
                                    Dim chkIfExsist As String = "Select Count(*) From ITM9 Where ""ItemCode"" = '" & oGrid.DataTable.GetValue("Item Code", i).ToString & "' And ""UomEntry"" = '" & UomEntry & "'"
                                    Dim oRschkIfExsist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRschkIfExsist.DoQuery(chkIfExsist)
                                    If oRschkIfExsist.Fields.Item(0).Value > 0 Then
                                        For j As Integer = 0 To oItemMaster.PriceList.UoMPrices.Count - 1
                                            oItemMaster.PriceList.UoMPrices.SetCurrentLine(j)
                                            If oItemMaster.PriceList.UoMPrices.UoMEntry = UomEntry Then
                                                oItemMaster.PriceList.UoMPrices.Price = NewPrice
                                                UOMFoundFlag = "Yes"
                                            End If
                                        Next
                                    End If
                                    If UOMFoundFlag = "" Then
                                        'Dim GetUomPriceListCount As String = "Select Count(*) From ITM9 Where ""ItemCode"" = '" & oItemMaster.ItemCode & "' And ""PriceList"" = '" & strPrlist & "'"
                                        'Dim oRsGetUomPriceListCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        'oRsGetUomPriceListCount.DoQuery(GetUomPriceListCount)
                                        'If oRsGetUomPriceListCount.Fields.Item(0).Value = 0 Then
                                        '    UomCount = CInt(oGrid.DataTable.GetValue("UomLine", i))
                                        'Else
                                        '    UomCount = oRsGetUomPriceListCount.Fields.Item(0).Value
                                        'End If
                                        'oItemMaster.PriceList.UoMPrices.Add()
                                        'oItemMaster.PriceList.UoMPrices.SetCurrentLine(UomCount)
                                        'oItemMaster.PriceList.UoMPrices.Price = NewPrice
                                        'oItemMaster.PriceList.UoMPrices.UoMEntry = UomEntry
                                        'oItemMaster.PriceList.UoMPrices.PriceList = strPrlist
                                        'oItemMaster.PriceList.UoMPrices.Auto = SAPbobsCOM.BoYesNoEnum.tNO
                                        'oItemMaster.PriceList.UoMPrices.Add()

                                        Dim GetUomPriceListCount As String = "Select Count(*) From ITM9 Where ""ItemCode"" = '" & oItemMaster.ItemCode & "' And ""PriceList"" = '" & strPrlist & "'"
                                        Dim oRsGetUomPriceListCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRsGetUomPriceListCount.DoQuery(GetUomPriceListCount)
                                        If oRsGetUomPriceListCount.Fields.Item(0).Value = 0 Then
                                            UomCount = 0
                                        Else
                                            UomCount = oRsGetUomPriceListCount.Fields.Item(0).Value
                                        End If
                                        oItemMaster.PriceList.UoMPrices.Add()
                                        oItemMaster.PriceList.UoMPrices.SetCurrentLine(UomCount)
                                        oItemMaster.PriceList.UoMPrices.Price = CDbl(NewPrice)
                                        oItemMaster.PriceList.UoMPrices.UoMEntry = UomEntry
                                        oItemMaster.PriceList.UoMPrices.PriceList = strPrlist
                                        oItemMaster.PriceList.UoMPrices.Auto = SAPbobsCOM.BoYesNoEnum.tNO
                                        oItemMaster.PriceList.UoMPrices.Currency = "QR"


                                    End If
                                End If
                            End If
                            oGrid.DataTable.SetValue("Status", i, "Updated Successfully")
                        End If
                        r = i + 1
                        If r <= oGrid.DataTable.Rows.Count - 1 Then
                            If oGrid.DataTable.GetValue("Item Code", i).ToString = oGrid.DataTable.GetValue("Item Code", r).ToString Then
                                i = i + 1
                                GoTo H
                            End If
                        End If
                        oGrid.DataTable.SetValue("Status", i, "In Progress...")

                        'objMain.objApplication.MessageBox("Please wait..price updation for item " & oGrid.DataTable.GetValue("Item Code", i).ToString & "  in progress...!!")
                        Dim RetValue As Integer = oItemMaster.Update()
                        If RetValue <> 0 Then
                            oGrid.DataTable.SetValue("Status", i, "Failed")
                            objMain.objApplication.MessageBox(objMain.objCompany.GetLastErrorDescription.ToString)
                        Else
                            objMain.objApplication.StatusBar.SetText("operation completed successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Me.LoadVendorPriceList(objForm.UniqueID)
                            objMain.objApplication.MessageBox("Updated Successfully..!!!")
                        End If
                    End If
                End If
            Next
            objForm.Freeze(False)
            oGrid.AutoResizeColumns()
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    'Public Sub UpdateDocStatus(ByVal FormUID As String, ByVal Status As String)
    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)
    '        oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD")
    '        oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD_C0")
    '        objMatrix = objForm.Items.Item("3").Specific

    '        objMain.sCmp = objMain.objCompany.GetCompanyService
    '        objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OPRICEUPD")
    '        objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
    '        objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0).ToString)
    '        objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
    '        objMain.oGeneralData.SetProperty("U_APPSTS", Status)
    '        objMain.oGeneralService.Update(objMain.oGeneralData)
    '        Me.RefreshData(objForm.UniqueID)
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub
    Private Sub PriceAutoFill(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("3").Specific

            If objForm.Items.Item("20").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Value cannot be null...")
                Exit Sub
            End If

            If objForm.Items.Item("22").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Please provide UomCode...")
                Exit Sub
            End If

            For i As Integer = 1 To objMatrix.VisualRowCount
                oDBs_Details.SetValue("LineId", oDBs_Details.Offset, i)
                oDBs_Details.SetValue("U_ITMCD", oDBs_Details.Offset, objMatrix.Columns.Item("V_0").Cells.Item(i).Specific.Value)
                oDBs_Details.SetValue("U_BRCHID", oDBs_Details.Offset, objMatrix.Columns.Item("V_1").Cells.Item(i).Specific.Value)
                oDBs_Details.SetValue("U_BRCHNM", oDBs_Details.Offset, objMatrix.Columns.Item("V_2").Cells.Item(i).Specific.Value)
                oDBs_Details.SetValue("U_PRCID", oDBs_Details.Offset, objMatrix.Columns.Item("V_3").Cells.Item(i).Specific.Value)
                oDBs_Details.SetValue("U_PRCLST", oDBs_Details.Offset, objMatrix.Columns.Item("V_4").Cells.Item(i).Specific.Value)
                oDBs_Details.SetValue("U_UOMENT", oDBs_Details.Offset, objMatrix.Columns.Item("V_5").Cells.Item(i).Specific.Value)
                oDBs_Details.SetValue("U_UOMCD", oDBs_Details.Offset, objMatrix.Columns.Item("V_6").Cells.Item(i).Specific.Value)
                oDBs_Details.SetValue("U_OLDPRC", oDBs_Details.Offset, objMatrix.Columns.Item("V_7").Cells.Item(i).Specific.Value)
                If objMatrix.Columns.Item("V_6").Cells.Item(i).Specific.Value = objForm.Items.Item("22").Specific.Value Then

                    Dim chkForAuthorization As String = " Select ""U_ACCESS"" , ""U_PRICEDIT"" FROM OUSR T0 Inner Join ""@SBO_PUBA"" T1 On ""USER_CODE"" = ""U_USRCD"" " &
                                                           " Inner Join ""@SBO_PUBA_C0"" T2 On T1.""DocEntry"" = T2.""DocEntry"" " &
                                                           " Where IFNULL(""U_PRICEDIT"",'N')='Y' And ""USER_CODE"" = '" & objMain.objCompany.UserName & "' and ""U_BPLCD"" = '" & objMatrix.Columns.Item("V_1").Cells.Item(i).Specific.Value & "'"
                    Dim oRschkForAuthorization As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRschkForAuthorization.DoQuery(chkForAuthorization)
                    If oRschkForAuthorization.RecordCount > 0 Then
                        oDBs_Details.SetValue("U_NEWPRC", oDBs_Details.Offset, objMatrix.Columns.Item("V_7").Cells.Item(i).Specific.Value)
                    Else
                        oDBs_Details.SetValue("U_NEWPRC", oDBs_Details.Offset, objForm.Items.Item("20").Specific.Value)
                    End If


                    oDBs_Details.SetValue("U_UPDT", oDBs_Details.Offset, DateTime.Now.ToString("yyyyMMdd"))
                    If objMatrix.Columns.Item("V_10").Cells.Item(i).Specific.Value <> "" Then
                        oDBs_Details.SetValue("U_ENDDT", oDBs_Details.Offset, objMatrix.Columns.Item("V_10").Cells.Item(i).Specific.Value)
                    Else
                        oDBs_Details.SetValue("U_ENDDT", oDBs_Details.Offset, "20401231")
                    End If
                    oDBs_Details.SetValue("U_USRCD", oDBs_Details.Offset, objMain.objCompany.UserName)
                Else
                    Dim chkForAuthorization As String = " Select ""U_ACCESS"" , ""U_PRICEDIT"" FROM OUSR T0 Inner Join ""@SBO_PUBA"" T1 On ""USER_CODE"" = ""U_USRCD"" " &
                                                          " Inner Join ""@SBO_PUBA_C0"" T2 On T1.""DocEntry"" = T2.""DocEntry"" " &
                                                          " Where IFNULL(""U_PRICEDIT"",'N')='Y' And ""USER_CODE"" = '" & objMain.objCompany.UserName & "' and ""U_BPLCD"" = '" & objMatrix.Columns.Item("V_1").Cells.Item(i).Specific.Value & "'"
                    Dim oRschkForAuthorization As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRschkForAuthorization.DoQuery(chkForAuthorization)
                    If oRschkForAuthorization.RecordCount > 0 Then
                        oDBs_Details.SetValue("U_NEWPRC", oDBs_Details.Offset, objMatrix.Columns.Item("V_7").Cells.Item(i).Specific.Value)
                    Else
                        oDBs_Details.SetValue("U_NEWPRC", oDBs_Details.Offset, objMatrix.Columns.Item("V_8").Cells.Item(i).Specific.Value)
                    End If

                    If objMatrix.Columns.Item("V_9").Cells.Item(i).Specific.Value <> "" Then
                        oDBs_Details.SetValue("U_UPDT", oDBs_Details.Offset, objMatrix.Columns.Item("V_9").Cells.Item(i).Specific.Value)
                    Else
                        oDBs_Details.SetValue("U_UPDT", oDBs_Details.Offset, DateTime.Now.ToString("yyyyMMdd"))
                    End If
                    If objMatrix.Columns.Item("V_10").Cells.Item(i).Specific.Value <> "" Then
                        oDBs_Details.SetValue("U_ENDDT", oDBs_Details.Offset, objMatrix.Columns.Item("V_10").Cells.Item(i).Specific.Value)
                    Else
                        oDBs_Details.SetValue("U_ENDDT", oDBs_Details.Offset, "20401231")
                    End If
                    oDBs_Details.SetValue("U_USRCD", oDBs_Details.Offset, objMatrix.Columns.Item("V_11").Cells.Item(i).Specific.Value)
                End If
                objMatrix.SetLineData(i)
            Next
            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
            objForm.Items.Item("14").Enabled = True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub LoadSpecialPricesBPList(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("18").Specific
            objForm.Freeze(True)
            Dim UomLine As String = ""

            oGrid.DataTable = oDt
            oGrid.DataTable.Clear()

            Dim Query As String = "  Select T0.""ItemCode"" As ""Item Code"", T2.""CardCode"" As ""Vendor Code"", T2.""CardName"" As ""Vendor Name"", " &
            " T2.""ListNum"" As  ""Price List ID"", (Select ""ListName"" From OPLN  Where ""ListNum"" =  T2.""ListNum"") As ""Price List"", " &
            " B.""UomEntry""as ""UoM Entry"", C.""UomCode""as ""UoM Code"", ( Case  When B.""UomEntry"" = T3.""UomEntry"" Then IFNULL(T3.""Price"",0) Else 0 End) As ""Old Price"", " &
            " CAST ('0' AS Float) As ""New Price"", '' As ""Status"" , '' As ""UoM Line"", T0.""IUoMEntry"" As ""Default UoM""  " &
            " From OITM T0 Inner Join OUGP A On A.""UgpEntry"" = T0.""UgpEntry"" Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry""  " &
            " Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry"" Inner Join ITM2 T4 On T4.""ItemCode"" = T0.""ItemCode""  " &
            " Inner Join OCRD T2 On T2.""CardCode"" = T4.""VendorCode""  " &
            " Left Join SPP2 T3 On T0.""ItemCode"" = T3.""ItemCode"" and T2.""CardCode"" = T3.""CardCode"" And B.""UomEntry"" = T3.""UomEntry""  " &
            " Where T0.""ItemCode"" = '" & objForm.Items.Item("5").Specific.Value.Trim & "' "

            oGrid.DataTable.ExecuteQuery(Query)
            If oGrid.Rows.Count = 0 Then
                objMain.objApplication.MessageBox("No records found")
                Exit Sub
            End If
            For i As Integer = 0 To oGrid.Columns.Count - 1
                oGrid.Columns.Item(i).Editable = False
            Next

4:          oGrid.Columns.Item("New Price").Editable = True
            Dim oEditTextCol As SAPbouiCOM.EditTextColumn
            oEditTextCol = oGrid.Columns.Item("Item Code")
            oEditTextCol.LinkedObjectType = 4

            oEditTextCol = oGrid.Columns.Item("Vendor Code")
            oEditTextCol.LinkedObjectType = 2

            Dim FirstUomLine As String = ""

            Dim Count As Integer = 0
            For i As Integer = 0 To oGrid.Rows.Count - 1
                If oGrid.DataTable.GetValue("Default UoM", i) <> oGrid.DataTable.GetValue("UoM Entry", i) Then
                    If FirstUomLine = "" Then
                        FirstUomLine = oGrid.DataTable.GetValue("UoM Entry", i)
                    End If
                    If FirstUomLine = oGrid.DataTable.GetValue("UoM Entry", i) Then
                        oGrid.DataTable.SetValue("UoM Line", i, Count)
                    Else
                        Count = Count + 1
                        oGrid.DataTable.SetValue("UoM Line", i, Count)
                        FirstUomLine = oGrid.DataTable.GetValue("UoM Entry", i)
                    End If
                End If
            Next

            oGrid.AutoResizeColumns()
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Sub UpdateSpecialPriceBP(ByVal FormUID As String)
        Try
H:          objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("18").Specific
            Dim UomCount As Integer = 0
            Dim retV As Integer = 0
            Dim r As Integer = 0
            Dim spItemCode As String = ""
            Dim spCardCode As String = ""
            Dim qStrLocalCurr As String = "SELECT ""MainCurncy"" FROM OADM"
            Dim rSetLocalCurr As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rSetLocalCurr.DoQuery(qStrLocalCurr)
            Dim Currency As String = rSetLocalCurr.Fields.Item("MainCurncy").Value.ToString()

            Dim oSpp As SAPbobsCOM.SpecialPrices = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSpecialPrices)
            For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If oGrid.DataTable.GetValue("New Price", i) > 0 Then
                    spItemCode = oGrid.DataTable.GetValue("Item Code", i).ToString
                    spCardCode = oGrid.DataTable.GetValue("Vendor Code", i).ToString
                    Dim NewPrice As Double = CDbl(oGrid.DataTable.GetValue("New Price", i).ToString)
                    Dim UomEntry As Integer = CInt(oGrid.DataTable.GetValue("UoM Entry", i).ToString)
                    If oSpp.GetByKey(spItemCode, spCardCode) Then

                        If NewPrice > 0 Then
                            'Dim UomEntry As Integer = CInt(oGrid.DataTable.GetValue("UomEntry", i).ToString)
                            Dim DefaultUom As Integer = CInt(oGrid.DataTable.GetValue("Default UoM", i).ToString)
                            Dim PriceListRow As Integer = oGrid.DataTable.GetValue("Price List ID", i)
                            oSpp.ItemCode = spItemCode
                            oSpp.CardCode = spCardCode
                            oSpp.Currency = Currency
                            oSpp.PriceListNum = 0
                            oSpp.Valid = SAPbobsCOM.BoYesNoEnum.tNO

                            If DefaultUom = UomEntry Then

                                Dim GetUomPriceListCount As String = "Select IFNULL(""SPP2LNum"",0) From SPP2 " &
                               " Where ""ItemCode"" = '" & spItemCode & "' and ""CardCode"" = '" & spCardCode & "' and ""UomEntry"" = '" & UomEntry & "' "
                                Dim oRsGetUomPriceListCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetUomPriceListCount.DoQuery(GetUomPriceListCount)
                                If oRsGetUomPriceListCount.Fields.Item(0).Value = 0 Then
                                    UomCount = 0
                                Else
                                    UomCount = oRsGetUomPriceListCount.Fields.Item(0).Value
                                End If

                                oSpp.Price = CDbl(NewPrice)
                                oSpp.SpecialPricesDataAreas.PriceListNo = oSpp.PriceListNum
                                oSpp.SpecialPricesDataAreas.DateFrom = Now
                                oSpp.SpecialPricesDataAreas.Dateto = Now
                                oSpp.SpecialPricesDataAreas.SpecialPrice = CDbl(NewPrice)
                                oSpp.SpecialPricesDataAreas.Discount = 0

                                oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.SetCurrentLine(UomCount)
                                oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.UoMEntry = UomEntry
                                oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.PriceCurrency = Currency
                                oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.Quantity = 1
                                oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.SpecialPrice = CDbl(NewPrice)

                                'For j As Integer = 0 To oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.Count - 1

                                '    oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.SetCurrentLine(j)
                                '    If oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.UoMEntry = UomEntry Then

                                '        oSpp.Price = CDbl(NewPrice)

                                '        oSpp.SpecialPricesDataAreas.PriceListNo = oSpp.PriceListNum
                                '        oSpp.SpecialPricesDataAreas.DateFrom = Now
                                '        oSpp.SpecialPricesDataAreas.Dateto = Now
                                '        oSpp.SpecialPricesDataAreas.SpecialPrice = CDbl(NewPrice)

                                '        oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.UoMEntry = UomEntry
                                '        oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.PriceCurrency =Currency
                                '        oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.Quantity = 1
                                '        oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.SpecialPrice = CDbl(NewPrice)

                                '    End If
                                'Next

                            Else
                                Dim UOMFoundFlag As String = ""
                                'Dim chkIfExsist As String = "Select Count(*) From ITM9 Where ""ItemCode"" = '" & oGrid.DataTable.GetValue("Item Code", i).ToString & "' And ""UomEntry"" = '" & UomEntry & "'"
                                Dim chkIfExsist As String = "Select Count(*) From SPP2 " &
                                " Where ""ItemCode"" = '" & spItemCode & "' and ""CardCode"" = '" & spCardCode & "' " &
                                " And ""UomEntry"" = '" & UomEntry & "'"
                                Dim oRschkIfExsist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRschkIfExsist.DoQuery(chkIfExsist)
                                If oRschkIfExsist.Fields.Item(0).Value > 0 Then
                                    For j As Integer = 0 To oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.Count - 1

                                        'oSpp.SpecialPricesDataAreas.PriceListNo = oSpp.PriceListNum
                                        'oSpp.SpecialPricesDataAreas.DateFrom = Now
                                        'oSpp.SpecialPricesDataAreas.SpecialPrice = CDbl(NewPrice)

                                        oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.SetCurrentLine(j)
                                        If oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.UoMEntry = UomEntry Then

                                            'oSpp.Price = CDbl(NewPrice)
                                            'oSpp.SpecialPricesDataAreas.PriceListNo = oSpp.PriceListNum
                                            'oSpp.SpecialPricesDataAreas.DateFrom = Now
                                            'oSpp.SpecialPricesDataAreas.SpecialPrice = CDbl(NewPrice)

                                            oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.UoMEntry = UomEntry
                                            oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.PriceCurrency = Currency
                                            oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.Quantity = 1
                                            oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.SpecialPrice = CDbl(NewPrice)
                                            UOMFoundFlag = "Yes"
                                        End If
                                    Next
                                End If
                                If UOMFoundFlag = "" Then
                                    'Dim GetUomPriceListCount As String = "Select Count(*) From ITM9 Where ""ItemCode"" = '" & oGrid.DataTable.GetValue("Item Code", i).ToString & "' And ""PriceList"" = '" & strPrlist & "'"
                                    Dim GetUomPriceListCount As String = "Select Count(*) From SPP2 " &
                                    " Where ""ItemCode"" = '" & spItemCode & "' and ""CardCode"" = '" & spCardCode & "' "
                                    Dim oRsGetUomPriceListCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRsGetUomPriceListCount.DoQuery(GetUomPriceListCount)
                                    If oRsGetUomPriceListCount.Fields.Item(0).Value = 0 Then
                                        UomCount = 0
                                    Else
                                        UomCount = oRsGetUomPriceListCount.Fields.Item(0).Value
                                    End If

                                    'oSpp.SpecialPricesDataAreas.Add()
                                    'oSpp.SpecialPricesDataAreas.PriceListNo = oSpp.PriceListNum
                                    'oSpp.SpecialPricesDataAreas.DateFrom = Now
                                    'oSpp.SpecialPricesDataAreas.SpecialPrice = CDbl(NewPrice)

                                    oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.Add()
                                    oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.SetCurrentLine(UomCount)
                                    oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.UoMEntry = UomEntry
                                    oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.PriceCurrency = Currency
                                    'oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.UoMEntry = UomCount
                                    oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.Quantity = 1
                                    oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.SpecialPrice = CDbl(NewPrice)

                                End If
                            End If


                        End If

                        'objMain.objApplication.MessageBox("Please wait..price updation for item " & oGrid.DataTable.GetValue("Item Code", i).ToString & "  in progress...!!")
                        'oGrid.DataTable.SetValue("Status", i, "In Progress...")

                        retV = oSpp.Update()

                        If retV <> 0 Then
                            Dim test As String = objMain.objCompany.GetLastErrorDescription
                            'oGrid.DataTable.SetValue("Status", i, "Failed")
                        End If
                        oGrid.DataTable.SetValue("New Price", i, 0)

                    Else 'Add

                        oSpp.AutoUpdate = SAPbobsCOM.BoYesNoEnum.tNO
                        oSpp.CardCode = spCardCode
                        oSpp.ItemCode = spItemCode
                        oSpp.Currency = "QR"
                        oSpp.PriceListNum = 0
                        oSpp.Valid = SAPbobsCOM.BoYesNoEnum.tNO
                        oSpp.Price = CDbl(NewPrice)

                        'oSpp.SpecialPricesDataAreas.Add()
                        oSpp.SpecialPricesDataAreas.SetCurrentLine(0)
                        oSpp.SpecialPricesDataAreas.PriceListNo = oSpp.PriceListNum

                        'oSpp.SpecialPricesDataAreas.DateFrom = #11/21/2007 12:14:00 PM#
                        'oSpp.SpecialPricesDataAreas.Dateto = #11/30/2007 12:14:00 PM#

                        oSpp.SpecialPricesDataAreas.DateFrom = Now
                        'oSpp.SpecialPricesDataAreas.Dateto = Now

                        oSpp.SpecialPricesDataAreas.AutoUpdate = SAPbobsCOM.BoYesNoEnum.tNO
                        oSpp.SpecialPricesDataAreas.SpecialPrice = CDbl(NewPrice)
                        'oSpp.SpecialPricesDataAreas.Discount = 0

                        'oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.Add()
                        oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.SetCurrentLine(0)
                        oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.PriceCurrency = Currency
                        oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.UoMEntry = UomEntry
                        oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.Quantity = 1
                        oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.SpecialPrice = CDbl(NewPrice)

                        retV = oSpp.Add()

                        If retV <> 0 Then
                            Dim test As String = objMain.objCompany.GetLastErrorDescription
                            GoTo H
                            'oGrid.DataTable.SetValue("Status", i, "Failed")
                        End If

                        'oSpp.Valid = SAPbobsCOM.BoYesNoEnum.tNO
                        'retV = oSpp.Update()
                        'If retV <> 0 Then
                        '    Dim test As String = objMain.objCompany.GetLastErrorDescription
                        'End If

                    End If
                End If
            Next

            If retV <> 0 Then
                'Dim test As String = objMain.objCompany.GetLastErrorDescription
                objMain.objApplication.MessageBox(objMain.objCompany.GetLastErrorDescription.ToString)
            Else
                objMain.objApplication.StatusBar.SetText("operation completed successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Me.LoadSpecialPricesBPList(objForm.UniqueID)
                objMain.objApplication.MessageBox("Updated Successfully..!!!")
            End If

            objForm.Freeze(False)
            oGrid.AutoResizeColumns()
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Shared Sub FILTER_CFL(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef oApplication As SAPbouiCOM.Application, ByRef oCompany As SAPbobsCOM.Company, ByVal Query As String, ByVal Code As String)
        Try
            Dim oCLFEvent As SAPbouiCOM.ChooseFromListEvent = Nothing
            Dim oCell As SAPbouiCOM.Column
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFL_collection As SAPbouiCOM.ChooseFromListCollection
            Dim oCon As SAPbouiCOM.Condition = Nothing
            Dim emptyConds As SAPbouiCOM.Conditions = Nothing
            Dim oCons As SAPbouiCOM.Conditions = Nothing
            Dim oRec As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRec.DoQuery(Query)
            Dim strCFLID As String = Nothing
            oCLFEvent = CType((pVal), SAPbouiCOM.ChooseFromListEvent)
            strCFLID = oCLFEvent.ChooseFromListUID
            oCFL = oApplication.Forms.ActiveForm.ChooseFromLists.Item(strCFLID)
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            Dim rr As Integer = oRec.RecordCount

            If rr = 0 Then
                oCon = oCons.Add()
                oCon.[Alias] = Code
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = oRec.Fields.Item(Code).Value.ToString()
            Else

                For i As Integer = 0 To oRec.RecordCount - 1
                    oCon = oCons.Add()
                    oCon.[Alias] = Code
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRec.Fields.Item(Code).Value.ToString()

                    If i <> oRec.RecordCount - 1 Then
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    End If

                    oRec.MoveNext()
                Next
            End If

            oCFL.SetConditions(oCons)
        Catch ex As Exception

        End Try

    End Sub

End Class
