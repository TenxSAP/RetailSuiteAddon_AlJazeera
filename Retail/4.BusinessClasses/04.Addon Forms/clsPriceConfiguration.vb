Public Class clsPriceConfiguration

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

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("PriceConfiguration.xml", "PRICECONFIG", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("PRICECONFIG", objMain.objApplication.Forms.ActiveForm.TypeCount)          
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PRICECONFIG")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_PRICECONFIG_C0")
            Me.SetDefault(objForm.UniqueID)
            Me.CellMasking(objForm.UniqueID)
            Me.objForm.DataBrowser.BrowseBy = "DocNum"
            Me.objForm.EnableMenu("1293", True)
            'objMatrix.AutoResizeColumns()
            

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub CellMasking(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Items.Item("DocDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("VName").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("VName").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("Load").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("Auto").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Calc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("Branch").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Selmar").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Exchprce").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("VCode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("ItmGrp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("FITMCODE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("TITMCODE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)


            'objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "", Optional ByVal Series As Integer = 0)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)          
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PRICECONFIG")
            objMatrix = objForm.Items.Item("Mtx").Specific

            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "PRICECONFIGUDO"))
            oDBs_Head.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"))
            Dim oRsBranch As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim Query As String = " Select T0.""BPLId"",T1.""BPLName"",T1.""DflWhs"" From USR6 T0 Inner Join " &
                                                                " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "'Order By T0.""BPLId""  "
            'Dim Query As String = "SELECT T2.""BPLId"",T2.""BPLName"" AS ""Branch"" FROM USR6 T0  INNER JOIN OUSR T1 ON T0.""UserCode"" = T1.""USER_CODE"" " & _
            '        " INNER JOIN OBPL T2 ON T0.""BPLId"" = T2.""BPLId"" WHERE T1.""USER_CODE"" = '" & objMain.objCompany.UserName.ToString().Trim & "' "
            oRsBranch.DoQuery(Query)           
            objComboBox = objForm.Items.Item("Branch").Specific
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, Query)
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("Branch").DisplayDesc = True
            If objComboBox.ValidValues.Count <> 0 Then
                objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
            End If
            Dim oRsItemType As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim itemtype As String = "SELECT ""ItmsGrpCod"",""ItmsGrpNam"" AS ""GroupName"" FROM OITB"
            oRsItemType.DoQuery(itemtype)
            objComboBox = objForm.Items.Item("ItmGrp").Specific
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, itemtype)
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("ItmGrp").DisplayDesc = True
            Me.SetNewLine(objForm.UniqueID)
            objForm.Items.Item("Calc").Enabled = False
            objForm.Items.Item("Auto").Enabled = False
            objForm.Items.Item("1").Enabled = False
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PRICECONFIG")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_PRICECONFIG_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific

            objMatrix.AddRow()
            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
            oDBs_Details.SetValue("U_ItmCo", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_ItmNm", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_BRCHID", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_PRCID", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_UOMENT", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_UmCode", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Branch", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Prlist", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_PriceC", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_SMar", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_ExchP", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Newpr", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_User", oDBs_Details.Offset, "")

            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType                  

                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PRICECONFIG")

                    If pVal.ItemUID = "ItmGrp" And pVal.BeforeAction = False And Me.objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If objForm.Items.Item("ItmGrp").Specific.Selected.Value <> "" Then
                            objForm.Items.Item("FITMCODE").Specific.Value = ""
                            objForm.Items.Item("TITMCODE").Specific.Value = ""
                            objForm.Items.Item("Load").Enabled = True
                        End If
                    End If
                    If pVal.ItemUID = "Curren" And pVal.BeforeAction = False And Me.objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If objForm.Items.Item("Curren").Specific.Selected.Value <> "" Then
                            Dim Currency As String = objForm.Items.Item("Curren").Specific.Selected.Value
                            Dim date1 As String = objForm.Items.Item("DocDate").Specific.Value

                            Dim CheckUpDate As String = " Select ""Rate""  From ORTT  Where ""Currency"" = '" & Currency & "' " &
                                                " and  ""RateDate"" = '" & date1 & "' "
                            Dim oRsCheckUpDate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsCheckUpDate.DoQuery(CheckUpDate)
                            objForm.Items.Item("Exchprce").Specific.Value = oRsCheckUpDate.Fields.Item("Rate").Value
                            ' objForm.Items.Item("Load").Enabled = True
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PRICECONFIG")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_PRICECONFIG_C0")
                    objMatrix = objForm.Items.Item("Mtx").Specific
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                    If oCFL.UniqueID = "CFL_1" And pVal.BeforeAction = True Then
                        Me.CFLFilter(objForm.UniqueID, "CFL_1")
                    End If
                    If oCFL.UniqueID = "CFL_2" And pVal.BeforeAction = True Then
                        Me.CFLItem(objForm.UniqueID, "CFL_2", objForm.Items.Item("ItmGrp").Specific.Value, objForm.Items.Item("VCode").Specific.Value)
                        'Me.CFLFilter1(objForm.UniqueID, "CFL_2", objForm.Items.Item("ItmGrp").Specific.Value)
                    End If
                    If oCFL.UniqueID = "CFL_3" And pVal.BeforeAction = True Then
                        Me.CFLItem(objForm.UniqueID, "CFL_3", objForm.Items.Item("ItmGrp").Specific.Value, objForm.Items.Item("VCode").Specific.Value)
                        ' Me.CFLFilter1(objForm.UniqueID, "CFL_3", objForm.Items.Item("ItmGrp").Specific.Value)
                    End If
                    ''End If
                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        If oCFL.UniqueID = "CFL_1" Then
                            oDBs_Head.SetValue("U_VCode", oDBs_Head.Offset, oDT.GetValue("CardCode", 0))
                            oDBs_Head.SetValue("U_VName", oDBs_Head.Offset, oDT.GetValue("CardName", 0))
                        End If
                        If oCFL.UniqueID = "CFL_2" Then
                            oDBs_Head.SetValue("U_FITMCODE", oDBs_Head.Offset, oDT.GetValue("ItemCode", 0))
                        End If
                        If oCFL.UniqueID = "CFL_3" Then
                            oDBs_Head.SetValue("U_TITMCODE", oDBs_Head.Offset, oDT.GetValue("ItemCode", 0))
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    If pVal.ItemUID = "Load" And pVal.BeforeAction = False And Me.objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim Itemgroup As String = objForm.Items.Item("ItmGrp").Specific.Value
                        Dim vendorcode As String = objForm.Items.Item("VCode").Specific.Value


                        'Dim Query As String = "SELECT T0.""ItemCode"",T0.""ItemName"" ,C.""UomCode"",T2.""BPLName"" As ""BranchName"",T2.""BPLId"" As ""BranchId"",T2.""U_PricelistNo"" As  ""PricListId"" ,B.""UomEntry"" " &
                        '    ",(Select ""ListName"" From OPLN  Where ""ListNum"" =  T2.""U_PricelistNo"") As ""PriceList"", T1.""Price""  As ""OldPrice"" from OITM T0 " &
                        '    " Inner Join ITM1 T1 On T1.""ItemCode"" = T0.""ItemCode"" And  T1.""UomEntry"" = T0.""IUoMEntry""  INNER JOIN OUOM C ON T1.""UomEntry"" = C.""UomEntry""  Inner Join OBPL T2 On  T2.""U_PricelistNo"" = T1.""PriceList"" " &
                        '    "Inner Join OUGP A On A.""UgpEntry"" = T0.""UgpEntry""  Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry"" "
                        Dim Query As String = " Select T0.""ItemCode"" ,T0.""ItemName"", T2.""BPLId"" As ""BranchId"" ,  T2.""BPLName"" As ""BranchName"" , T2.""U_PricelistNo"" As  ""PricListId"" , " &
                                          " (Select ""ListName"" From OPLN  Where ""ListNum"" =  T2.""U_PricelistNo"") As ""PriceList"" ,   B.""UomEntry"" , C.""UomCode"" , " &
                                          "  T3.""Price""  As ""OldPrice""   From OITM T0 Inner Join OUGP A On A.""UgpEntry"" = T0.""UgpEntry""  Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry"" " &
                                          " Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry""    Inner Join ITM1 T3 On T3.""ItemCode"" = T0.""ItemCode""" &
                                          " Inner Join OBPL T2 On  T2.""U_PricelistNo"" = T3.""PriceList""  Inner Join ITM2 T4 On T4.""ItemCode"" = T0.""ItemCode"" "
                        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim Itemcode1 As String = objForm.Items.Item("FITMCODE").Specific.Value
                        Dim Itemcode2 As String = objForm.Items.Item("TITMCODE").Specific.Value
                        Query = Query + " Where T0.""ItmsGrpCod""= '" & Itemgroup & "' and T4.""VendorCode""= '" & vendorcode & "' and T0.""ItemCode""  BETWEEN '" & Itemcode1 & "' and '" & Itemcode2 & "'Order By T0.""ItemCode"" "
                        oRs.DoQuery(Query)
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PRICECONFIG")
                        oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_PRICECONFIG_C0")
                        objMatrix = objForm.Items.Item("Mtx").Specific
                        objMatrix.Clear()
                        oDBs_Details.Clear()
                        objMatrix.FlushToDataSource()
                        objMatrix.AutoResizeColumns()
                        For i As Integer = 1 To oRs.RecordCount
                            objMatrix.AddRow()

                            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
                            oDBs_Details.SetValue("U_ItmCo", oDBs_Details.Offset, oRs.Fields.Item("ItemCode").Value)
                            oDBs_Details.SetValue("U_ItmNm", oDBs_Details.Offset, oRs.Fields.Item("ItemName").Value)

                            oDBs_Details.SetValue("U_UOMENT", oDBs_Details.Offset, oRs.Fields.Item("UomEntry").Value)
                            oDBs_Details.SetValue("U_BRCHID", oDBs_Details.Offset, oRs.Fields.Item("BranchId").Value)
                            oDBs_Details.SetValue("U_PRCID", oDBs_Details.Offset, oRs.Fields.Item("PricListId").Value)

                            oDBs_Details.SetValue("U_UmCode", oDBs_Details.Offset, oRs.Fields.Item("UomCode").Value)
                            oDBs_Details.SetValue("U_Branch", oDBs_Details.Offset, oRs.Fields.Item("BranchName").Value)
                            oDBs_Details.SetValue("U_Prlist", oDBs_Details.Offset, oRs.Fields.Item("PriceList").Value)
                            oDBs_Details.SetValue("U_PriceC", oDBs_Details.Offset, oRs.Fields.Item("OldPrice").Value)
                            oDBs_Details.SetValue("U_SMar", oDBs_Details.Offset, "")
                            oDBs_Details.SetValue("U_ExchP", oDBs_Details.Offset, "")
                            oDBs_Details.SetValue("U_Newpr", oDBs_Details.Offset, "")
                            Dim CheckUpDate As String = " Select ""U_UPDT"" , ""U_ENDDT"" , ""U_USRCD"" From ""@ME_PRICEUPD_C0"" A Where A.""U_ITMCD"" = '" & oRs.Fields.Item("ItemCode").Value & "' " &
                                                " and  ""U_UOMCD"" = '" & oRs.Fields.Item("UomCode").Value & "' "
                            Dim oRsCheckUpDate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsCheckUpDate.DoQuery(CheckUpDate)
                            If oRsCheckUpDate.Fields.Item("U_USRCD").Value <> "" Then
                                oDBs_Details.SetValue("U_User", oDBs_Details.Offset, oRsCheckUpDate.Fields.Item("U_USRCD").Value.ToString)
                            Else
                                oDBs_Details.SetValue("U_User", oDBs_Details.Offset, "")
                            End If
                            objMatrix.SetLineData(objMatrix.VisualRowCount)
                            oRs.MoveNext()
                        Next
                        objForm.Items.Item("Load").Enabled = False
                        objForm.Items.Item("Auto").Enabled = True
                    End If
                    If pVal.ItemUID = "Calc" And pVal.Before_Action = False Then
                        objMatrix = objForm.Items.Item("Mtx").Specific
                        Dim SellingMargin As Double = Convert.ToDouble(objForm.Items.Item("Selmar").Specific.Value)
                        Dim s1 As Double = SellingMargin / 100
                        Dim ExchangePrice As Double = Convert.ToDouble(objForm.Items.Item("Exchprce").Specific.Value)
                        Dim s2 As Double = ExchangePrice / 100
                        For i As Integer = 1 To objMatrix.VisualRowCount
                            Dim Pricelist As Double = Convert.ToDouble(objMatrix.Columns.Item("PriceC").Cells.Item(i).Specific.Value)
                            Dim Supplier As Double = Convert.ToDouble(objMatrix.Columns.Item("SMar").Cells.Item(i).Specific.Value)
                            Dim s3 As Double = Supplier / 100
                            Dim Exchange As Double = Convert.ToDouble(objMatrix.Columns.Item("ExchP").Cells.Item(i).Specific.Value)
                            Dim s4 As Double = Exchange / 100
                            Dim test, test1 As Double
                            If s1 > 0 Then

                                test = Pricelist + ((Pricelist) * (s1))
                            Else
                                test = Pricelist + ((Pricelist) * (s3))

                            End If
                            If s2 > 0 Then

                                test1 = Pricelist * (s2)
                            Else
                                test1 = Pricelist * (s4)

                            End If
                            Dim Total As Double = test + test1
                            Dim Totalprice As Double = Math.Round(Total, 0)
                            objMatrix.Columns.Item("Newpr").Cells.Item(i).Specific.Value = Total
                        Next
                    End If
                    If pVal.ItemUID = "Auto" And pVal.Before_Action = False And Me.objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        objMatrix = objForm.Items.Item("Mtx").Specific
                        Dim SellingMargin As String = objForm.Items.Item("Selmar").Specific.Value
                        Dim ExchangePrice As String = objForm.Items.Item("Exchprce").Specific.Value

                        For i As Integer = 1 To objMatrix.VisualRowCount
                            objMatrix.Columns.Item("SMar").Cells.Item(i).Specific.Value = SellingMargin
                            objMatrix.Columns.Item("ExchP").Cells.Item(i).Specific.Value = ExchangePrice

                        Next
                        objForm.Items.Item("Auto").Enabled = False
                        objForm.Items.Item("Calc").Enabled = True
                        objForm.Items.Item("1").Enabled = True
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim chkIfExist As String = "Select ""DocNum"" From ""@SBO_PRICECONFIG"" Where ""DocEntry""=(Select MAX(""DocEntry"") From ""@SBO_PRICECONFIG"")"
                        Dim oRschkIfExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRschkIfExist.DoQuery(chkIfExist)
                        If oRschkIfExist.RecordCount > 0 Then
                            objForm.Freeze(True)
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            objForm.Items.Item("DocNum").Specific.value = oRschkIfExist.Fields.Item("DocNum").Value
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)


                            objForm.Freeze(False)
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PRICECONFIG")
                    If pVal.ItemUID = "Selmar" Or pVal.ItemUID = "Exchprce" And pVal.Before_Action = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If objForm.Items.Item("Selmar").Specific.Value <> "" Then
                            objForm.Items.Item("Auto").Enabled = True
                            objForm.Items.Item("Calc").Enabled = False
                        End If
                    End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub CFLItem(ByVal FormUID As String, ByVal CFL_ID As String, ByVal ItemGroup As String, ByVal Vendorcode As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            'Dim GetBarCode As String = " Select T0.""ItemCode"" ,T0.""ItemName"" From  OITM T0 " &
            '                           "  Where T0.""ItmsGrpCod"" " &
            '                           " =  '" & ItemGroup & "'"
            Dim GetBarCode As String = " SELECT T0.""ItemCode"", T0.""ItemName"" FROM OITM T0  INNER JOIN ITM2 T1 ON T0.""ItemCode"" = T1.""ItemCode"" WHERE T1.""VendorCode""='" & Vendorcode & "'and T0.""ItmsGrpCod"" ='" & ItemGroup & "'"

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
                oCondition.Alias = "ItemCode"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = oRsGetBarCode.Fields.Item(0).Value
                oChooseFromList.SetConditions(oConditions)
                oRsGetBarCode.MoveNext()
                For i As Integer = 1 To oRsGetBarCode.RecordCount - 1
                    oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCondition = oConditions.Add()
                    oCondition.Alias = "ItemCode"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = oRsGetBarCode.Fields.Item(0).Value
                    oChooseFromList.SetConditions(oConditions)
                    oRsGetBarCode.MoveNext()
                Next
            Else
                oCondition = oConditions.Add()
                oCondition.Alias = "ItemCode"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = ""
                oChooseFromList.SetConditions(oConditions)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    'Sub CFLFilter1(ByVal FormUID As String, ByVal CFL_ID As String, ByVal ItemGroup As String)
    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)
    '        Select Case CFL_ID
    '            Case CFL_ID
    '                Dim oConditions As SAPbouiCOM.Conditions
    '                Dim oCondition As SAPbouiCOM.Condition
    '                Dim oChooseFromList As SAPbouiCOM.ChooseFromList
    '                Dim emptyCon As New SAPbouiCOM.Conditions
    '                oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
    '                oChooseFromList.SetConditions(emptyCon)
    '                oConditions = oChooseFromList.GetConditions()
    '                oCondition = oConditions.Add()
    '                oCondition.Alias = "ItmsGrpCod"
    '                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '                oCondition.CondVal = ItemGroup
    '                oChooseFromList.SetConditions(oConditions)
    '        End Select
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ME_PRICECONFIG" And pVal.BeforeAction = False Then
                Me.CreateForm()

             ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
            End If
            'If pVal.MenuUID = "Delete Row" And pVal.BeforeAction = False Then
            '    objForm = objMain.objApplication.Forms.GetForm("ME_PRICEUPD_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            '    objMatrix = objForm.Items.Item("3").Specific
            '    For i As Integer = 1 To objMatrix.VisualRowCount - 1
            '        If objMatrix.IsRowSelected(i) = True Then
            '            objMatrix.DeleteRow(i)
            '        End If
            '    Next
            '    For i As Integer = 1 To objMatrix.VisualRowCount
            '        objMatrix.Columns.Item("V_-1").Cells.Item(i).Specific.string = i
            '    Next
            '    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            'End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("PRICECONFIG", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        'If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        '    objMatrix = objForm.Items.Item("Mtx").Specific
                        '    objMatrix.Columns.Item("SMar").Editable = False
                        '    objMatrix.Columns.Item("ExchP").Editable = False
                        '    objMatrix.Columns.Item("Newpr").Editable = False
                        '    objMatrix.Columns.Item("User").Editable = False

                        'End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                        If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                            Me.UpdatePrice(objForm.UniqueID)
                        End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                        If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                            Me.UpdatePrice(objForm.UniqueID)
                        End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub CFLFilter(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Select Case CFL_ID
                Case "CFL_1"
                    Dim oConditions As SAPbouiCOM.Conditions
                    Dim oCondition As SAPbouiCOM.Condition
                    Dim oChooseFromList As SAPbouiCOM.ChooseFromList
                    Dim emptyCon As New SAPbouiCOM.Conditions
                    oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
                    oChooseFromList.SetConditions(emptyCon)
                    oConditions = oChooseFromList.GetConditions()
                    oCondition = oConditions.Add()
                    oCondition.Alias = "CardType"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = "S"
                    oChooseFromList.SetConditions(oConditions)
            End Select
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
  
    Public Sub UpdatePrice(ByVal FormUID As String)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        Dim oRsGetItemPriceDetails As SAPbobsCOM.Recordset = Nothing
        Dim oRsGetItemPriceLineDtls As SAPbobsCOM.Recordset = Nothing
        Dim UomCount As Integer = 0
        objMatrix = objForm.Items.Item("Mtx").Specific
        For M = 1 To objMatrix.VisualRowCount
            sItemCode = objMatrix.Columns.Item("ItmCo").Cells.Item(M).Specific.Value
            Dim Docentry As String = objForm.Items.Item("DocNum").Specific.Value
            Try
                Dim GetItemPriceDetails As String = " Select  T1.""U_ItmCo"" , T0.""DocEntry"" As ""DocEntry"" " &
                                                    " From ""@SBO_PRICECONFIG"" T0 Inner Join ""@SBO_PRICECONFIG_C0"" T1 On T1.""DocEntry"" = T0.""DocEntry"" " &
                                                    " Inner Join OITM T2 ON T2.""ItemCode"" = T1.""U_ItmCo"" Where IFNULL(T1.""U_ItmCo"",'')<>'' " &
                                                    " And  T1.""U_ItmCo"" = '" & sItemCode & "' and T0.""DocNum"" = '" & Docentry & "' Group By T1.""U_ItmCo""  ,  T0.""DocEntry"" "

                oRsGetItemPriceDetails = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetItemPriceDetails.DoQuery(GetItemPriceDetails)
                If oRsGetItemPriceDetails.RecordCount > 0 Then
                    oRsGetItemPriceDetails.MoveFirst()
                    'Dim result As Integer = objMain.objApplication.MessageBox("Price updation is irreversible..Confirm... Do you want to continue?", 1, "Continue", "Cancel")
                    'If (result = 1) Then
                    objMain.objApplication.StatusBar.SetText("Price updation is in progress...Please wait.....!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    'Else
                    '    Exit Sub
                    'End If

                    Dim oItem As SAPbobsCOM.Items = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                    For i As Integer = 1 To oRsGetItemPriceDetails.RecordCount
                        If oItem.GetByKey(oRsGetItemPriceDetails.Fields.Item(0).Value) Then
                            Dim GetItemPriceLineDtls As String = " Select  T1.""U_Newpr"" As ""New Price"", T1.""U_PriceC"" As ""Old Price"" , T1.""U_ItmCo"" As ""Item Code"" , T1.""U_ItmNm"" As ""Item Name"" , " &
                                                                 "  T1.""U_BRCHID"" As ""Branch ID"",T1.""U_Branch"" As ""Branch Name""  ,""U_PRCID"" As ""Price List ID"" ,""U_Prlist"" As ""Price List"",""U_UOMENT"" As ""UoM Entry""  , " &
                                                                 " ""U_UmCode"" As ""UoM"" , T1.""U_User"" As ""User Id"" , T0.""DocEntry"" , T2.""IUoMEntry"" As ""Default UoM"" " &
                                                                 " From ""@SBO_PRICECONFIG"" T0 Inner Join ""@SBO_PRICECONFIG_C0"" T1 On T1.""DocEntry"" = T0.""DocEntry"" " &
                                                                 " Inner Join OITM T2 ON T2.""ItemCode"" = T1.""U_ItmCo"" Where IFNULL(T1.""U_ItmCo"",'')<>'' And T1.""U_Newpr"" > 0 " &
                                                                 " And  T1.""U_ItmCo"" = '" & oRsGetItemPriceDetails.Fields.Item(0).Value & "' and T0.""DocNum"" = '" & Docentry & "' "


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
                                                oItem.PriceList.UoMPrices.Currency = "AED"
                                            End If
                                        End If
                                    End If
                                    oRsGetItemPriceLineDtls.MoveNext()
                                Next
                                'Try
                                '    oItem.UserFields.Fields.Item("U_CXS_MDFD").Value = "N"
                                'Catch ex As Exception
                                'End Try
                                Dim RetValue As Integer = oItem.Update
                                '                If RetValue <> 0 Then
                                '                    Dim Err As String = objMain.objCompany.GetLastErrorDescription.Replace("'", "")
                                '                    Dim PriceUpdate As String = "Update ""@ME_PRICEUPD"" Set  ""U_Updated"" = 'N' , ""U_ErrorMsg"" = '" & Err & "'  Where ""DocEntry"" = '" & oRsGetItemPriceDetails.Fields.Item(1).Value & "' "
                                '                    Dim oRsPriceUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                '                    oRsPriceUpdate.DoQuery(PriceUpdate)
                                '                Else
                                '                    Dim PriceUpdate As String = "Update ""@ME_PRICEUPD"" Set  ""U_Updated"" = 'Y' , ""U_ErrorMsg"" = '' Where ""DocEntry"" = '" & oRsGetItemPriceDetails.Fields.Item(1).Value & "' "
                                '                    Dim oRsPriceUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                '                    oRsPriceUpdate.DoQuery(PriceUpdate)
                                '                    'Me.UpdateDocStatus(objForm.UniqueID, "Submitted")
                                '                    objMain.objApplication.MessageBox("Price updation completed successfully...!", 1, "Ok")
                                '                    'objMatrix = objForm.Items.Item("3").Specific
                                '                    'oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PRICEUPD_C0")
                                '                    Dim BCD As String = ""
                                '                    BCD = objForm.Items.Item("Item_1").Specific.Value
                                '                    Try
                                '                        objForm.Freeze(True)
                                '                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                '                        objForm.Items.Item("Item_1").Specific.Value = BCD
                                '                        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                '                    Catch ex As Exception
                                '                    Finally
                                '                        objForm.Freeze(False)
                                '                    End Try
                                '                End If
                            End If
                        End If
                        oRsGetItemPriceDetails.MoveNext()
                    Next
                    'objForm.Items.Item("13").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                End If

            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText(ex.Message)

            End Try
        Next
    End Sub
End Class

