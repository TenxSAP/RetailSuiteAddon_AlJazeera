Public Class clsSupplierAgreement

#Region "    Declaration   "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oEditText As SAPbouiCOM.EditText
    Dim oDt As SAPbouiCOM.DataTable
    Dim sSupplierCode, sSupplierName As String
    Dim objComboBox As SAPbouiCOM.ComboBox
#End Region

    Sub CreateForm(ByVal FormUID As String, ByVal CardCode As String, ByVal CardName As String, ByVal IfSRegExist As String, Optional SRegDocNum As String = "")
        Try
            objMain.objUtilities.LoadForm("SupplierAgreement.xml", "ME_SUPAG_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_SUPAG_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)

            If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = False Then
                Retail.MainCls.ohtLookUpForm.Add(objForm.UniqueID, FormUID)
            End If
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_SUPAG")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_SUPAG_C0")
            objMatrix = objForm.Items.Item("15").Specific
            objForm.Items.Item("28").Visible = False
            objForm.Items.Item("Item_1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            objComboBox = objForm.Items.Item("12").Specific
            objComboBox.ValidValues.Add("", "")
            objComboBox.ValidValues.Add("Limited", "Limited")
            objComboBox.ValidValues.Add("UnLimited", "UnLimited")
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            objMatrix.Columns.Item("V_3").ValidValues.Add("", "")
            objMatrix.Columns.Item("V_3").ValidValues.Add("Active", "Active")
            objMatrix.Columns.Item("V_3").ValidValues.Add("InActive", "InActive")
            objMatrix.Columns.Item("V_3").ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            If IfSRegExist = "Y" Then
                objForm.Freeze(True)
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                objForm.Items.Item("8").Specific.value = SRegDocNum
                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)


                objComboBox = objForm.Items.Item("10").Specific
                Dim GetSeriesQty As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'ME_OSUPAG' And ""Locked"" = 'N'"
                objMain.objUtilities.ComboBoxLoadValuesSeries(objForm.Items.Item("10").Specific, GetSeriesQty)
                objForm.Items.Item("10").DisplayDesc = True
                objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly


                objComboBox = objForm.Items.Item("31").Specific

                objMain.objUtilities.ComboBoxLoadValues(objComboBox, " Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join " & _
                                                                    " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""  ")
                objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                objForm.Items.Item("31").DisplayDesc = True

                objForm.Freeze(False)
            End If
            If IfSRegExist = "N" Then
                sSupplierCode = CardCode
                sSupplierName = CardName
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "ME_OSUPAG"))
                oDBs_Head.SetValue("U_SUPCD", oDBs_Head.Offset, sSupplierCode)
                oDBs_Head.SetValue("U_SUPNM", oDBs_Head.Offset, sSupplierName)

                objComboBox = objForm.Items.Item("10").Specific
                Dim GetSeriesQty As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'ME_OSUPAG' And ""Locked"" = 'N'"
                objMain.objUtilities.ComboBoxLoadValuesSeries(objForm.Items.Item("10").Specific, GetSeriesQty)
                objForm.Items.Item("10").DisplayDesc = True
                objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

                Dim GetDflt As String = "Select ""DfltSeries"" From ONNM Where ""ObjectCode"" = 'ME_OSUPAG'"
                Dim oRsGetDflt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetDflt.DoQuery(GetDflt)

                oDBs_Head.SetValue("U_SERIES", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)
                oDBs_Head.SetValue("Series", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)

                objComboBox = objForm.Items.Item("31").Specific
                objMain.objUtilities.ComboBoxLoadValues(objComboBox, " Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join " & _
                                                                     " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""  ")
                objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                objForm.Items.Item("31").DisplayDesc = True
                If objComboBox.ValidValues.Count <> 0 Then
                    objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
                End If


                objMatrix.Clear()
                oDBs_Details.Clear()
                objMatrix.FlushToDataSource()
                Me.SetNewLine(objForm.UniqueID)
            End If

            Me.CellMasking(objForm.UniqueID)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub CellMasking(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("24").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("24").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "2", SAPbouiCOM.BoModeVisualBehavior.mvb_True)


        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_SUPAG")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_SUPAG_C0")
            objMatrix = objForm.Items.Item("15").Specific

            objMatrix.AddRow()
            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
            oDBs_Details.SetValue("U_PRAGGNO", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_STRDT", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_ENDDT", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_STATUS", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_REMARKS", oDBs_Details.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim chkIfExist As String = "Select ""DocNum"" From ""@ME_SUPAG"" Where ""DocEntry""=(Select MAX(""DocEntry"") From ""@ME_SUPAG"")"
                        Dim oRschkIfExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRschkIfExist.DoQuery(chkIfExist)
                        If oRschkIfExist.RecordCount > 0 Then
                            objForm.Freeze(True)
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            objForm.Items.Item("8").Specific.value = oRschkIfExist.Fields.Item("DocNum").Value
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            objForm.Freeze(False)
                        End If
                    End If
                    If pVal.ItemUID = "16" And pVal.BeforeAction = False Then
                        objMain.objLoadAgreementsList.CreateForm(objForm.UniqueID, objForm.Items.Item("4").Specific.Value, "Y")
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                    If pVal.ItemUID = "29" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("24").Specific.Value = "" Then
                            Me.PostJE(objForm.UniqueID)
                        Else
                            objMain.objApplication.StatusBar.SetText("JE Posted already...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "Delete Row" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("ME_SUPAG_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix = objForm.Items.Item("15").Specific
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
            objForm = objMain.objApplication.Forms.GetForm("ME_SUPAG_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Me.UpdateBP(objForm.UniqueID)
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub UpdateBP(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim oBusinessPartners As SAPbobsCOM.BusinessPartners = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            If oBusinessPartners.GetByKey(objForm.Items.Item("4").Specific.Value) Then
                oBusinessPartners.UserFields.Fields.Item("U_SUPREGDE").Value = objForm.Items.Item("8").Specific.Value
                oBusinessPartners.Update()
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

#Region " Right Click Event"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Dim objForm As SAPbouiCOM.Form
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)
        Try
            If eventInfo.FormUID = objForm.UniqueID Then
                If (eventInfo.BeforeAction = True) Then
                    If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                        objMatrix = objForm.Items.Item("15").Specific
                        If eventInfo.ItemUID = "15" And eventInfo.ColUID = "V_-1" Then
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("Delete Row") = False Then
                                    oCreationPackage.UniqueID = "Delete Row"
                                    oCreationPackage.String = "Delete Row"
                                    oCreationPackage.Enabled = True
                                    oMenus.AddEx(oCreationPackage)
                                End If
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        ElseIf eventInfo.ItemUID = "15" And objMatrix.RowCount <= 1 Then
                            oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                            oMenus = oMenuItem.SubMenus
                            Try
                                If oMenus.Exists("Delete Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Delete Row")
                                End If
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If
                        If eventInfo.ItemUID <> "15" Then
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("Delete Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Delete Row")
                                End If
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If
                    End If
                Else
                    Try
                        oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                        oMenus = oMenuItem.SubMenus
                        If oMenus.Exists("Delete Row") = True Then
                            objMain.objApplication.Menus.RemoveEx("Delete Row")
                        End If
                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(ex.Message)
                    End Try
                End If
            End If
            ' System.Diagnostics.Process.Start()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

    Private Sub PostJE(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_SUPAG")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_SUPAG_C0")
            Dim objComboBoxBranch As SAPbouiCOM.ComboBox = objForm.Items.Item("31").Specific
            If objComboBoxBranch.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Please select the branch.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            Dim oJE As SAPbobsCOM.JournalEntries = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

            Dim GetAddonConfigurationDtls As String = "Select ""U_SAGCRACT"" From ""@ME_ACFG"" "
            Dim oRsGetAddonConfigurationDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetAddonConfigurationDtls.DoQuery(GetAddonConfigurationDtls)


            Dim sDate As String = DateTime.Now.ToString("yyyyMMdd")
            sDate = sDate.Insert("4", "-")
            sDate = sDate.Insert("7", "-")

            oJE.ReferenceDate = sDate
            oJE.DueDate = sDate
            oJE.TaxDate = sDate

            oJE.Reference2 = "SupplierAgreement"

            oJE.Lines.BPLID = objComboBoxBranch.Value
            oJE.Lines.AccountCode = oRsGetAddonConfigurationDtls.Fields.Item("U_SAGCRACT").Value
            oJE.Lines.Credit = objForm.Items.Item("18").Specific.value
            oJE.Lines.Add()

            oJE.Lines.BPLID = objComboBoxBranch.Value
            oJE.Lines.ShortName = objForm.Items.Item("4").Specific.value
            oJE.Lines.Debit = objForm.Items.Item("18").Specific.value

            oJE.Lines.Add()

            If oJE.Add = 0 Then

                Dim GetJEDocEntry As String = "Select Max(""TransId"") From OJDT Where ""TransType"" = '30'"
                Dim oRsGetJEDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetJEDocEntry.DoQuery(GetJEDocEntry)

                'Updating Payment
                objMain.sCmp = objMain.objCompany.GetCompanyService
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OSUPAG")
                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0).ToString)
                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                objMain.oGeneralData.SetProperty("U_JENO", oRsGetJEDocEntry.Fields.Item(0).Value.ToString)
                objMain.oGeneralService.Update(objMain.oGeneralData)

                objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_SUPAG"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_SUPAG_C0"), "DocEntry", oDBs_Details.GetValue("DocEntry", 0))
                objMatrix.LoadFromDataSource()
                objForm.Refresh()
            Else
                objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorCode & " : " & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
