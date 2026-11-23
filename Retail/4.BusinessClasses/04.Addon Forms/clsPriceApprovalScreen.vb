Public Class clsPriceApprovalScreen

#Region "        Declaration        "
    Dim objForm As SAPbouiCOM.Form
    Dim oGrid As SAPbouiCOM.Grid
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim oEditText As SAPbouiCOM.EditText
    Dim oDt As SAPbouiCOM.DataTable
    Dim oProgressBar As SAPbouiCOM.ProgressBar = Nothing
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("PriceApprovalScreen.xml", "ME_PRAPS_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_PRAPS_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.DataSources.UserDataSources.Add("U_APPSTS", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            objForm.DataSources.UserDataSources.Add("U_BRNCH", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            objForm.DataSources.UserDataSources.Add("U_ITMCD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_ITMNM", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            objForm.DataSources.UserDataSources.Add("U_ITMGRP", SAPbouiCOM.BoDataType.dt_LONG_TEXT)

            objComboBox = objForm.Items.Item("11").Specific
            objComboBox.DataBind.SetBound(True, "", "U_APPSTS")
            objForm.Items.Item("11").AffectsFormMode = False
            objComboBox.ValidValues.Add("", "")
            objComboBox.ValidValues.Add("All", "All")
            objComboBox.ValidValues.Add("Submitted", "Submitted")
            objComboBox.ValidValues.Add("Approved", "Approved")
            objComboBox.Select("Submitted", SAPbouiCOM.BoSearchKey.psk_ByValue)

            objComboBox = objForm.Items.Item("6").Specific
            objComboBox.DataBind.SetBound(True, "", "U_BRNCH")
            objForm.Items.Item("6").AffectsFormMode = False
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, " Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join " & _
                                                                 " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""  ")
            objForm.Items.Item("6").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            oEditText = objForm.Items.Item("ITMCD").Specific
            oEditText.DataBind.SetBound(True, "", "U_ITMCD")

            oEditText.ChooseFromListUID = "CFL_ITMCD"
            oEditText.ChooseFromListAlias = "ItemCode"
            objForm.Items.Item("ITMCD").AffectsFormMode = False

            objComboBox = objForm.Items.Item("15").Specific
            objComboBox.DataBind.SetBound(True, "", "U_ITMGRP")
            objForm.Items.Item("15").AffectsFormMode = False
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, " Select ""ItmsGrpCod"" , ""ItmsGrpNam"" From OITB ")
            objForm.Items.Item("15").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            oEditText = objForm.Items.Item("4").Specific
            oEditText.DataBind.SetBound(True, "", "U_ITMNM")

            objForm.Items.Item("4").AffectsFormMode = False
            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            oDt = objForm.DataSources.DataTables.Add("dt1")
            oDt = objForm.DataSources.DataTables.Item("dt1")

            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ME_PRAPS" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.ActiveForm
                Me.CreateForm()
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    If pVal.ItemUID = "13" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("11").Specific.Selected.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Approval Type", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            Me.LoadGrid(objForm.UniqueID, "Y")
                        End If
                    End If

                    If pVal.ItemUID = "16" And pVal.BeforeAction = False Then
                        Me.LoadGrid(objForm.UniqueID, "Y")
                    End If

                    If pVal.ItemUID = "17" And pVal.BeforeAction = False Then
                        Me.LoadGrid(objForm.UniqueID, "N")
                    End If

                    If pVal.ItemUID = "8" And pVal.BeforeAction = False Then
                        Dim result As Integer = objMain.objApplication.MessageBox("Confirm... Do you want to continue?", 1, "Continue", "Cancel")
                        If (result = 1) Then
                            Me.ApprovePrice(objForm.UniqueID)
                        End If
                    End If
                    If pVal.ItemUID = "12" And pVal.BeforeAction = False Then
                        Dim result As Integer = objMain.objApplication.MessageBox("Confirm before updating... Do you want to continue?", 1, "Continue", "Cancel")
                        If (result = 1) Then
                            ' Me.UpdatePrice(objForm.UniqueID)
                        End If
                    End If

                    'If pVal.ItemUID = "13" And pVal.BeforeAction = False Then
                    '    Me.LoadGrid(objForm.UniqueID, "Y")
                    'End If
                    'If pVal.ItemUID = "14" And pVal.BeforeAction = False Then
                    '    Me.LoadGrid(objForm.UniqueID, "N")
                    'End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oGrid = objForm.Items.Item("7").Specific

                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        If oCFL.UniqueID = "CFL_ITMCD" Then
                            objForm.DataSources.UserDataSources.Item("U_ITMCD").Value = oDT.GetValue("ItemCode", 0)
                            objForm.DataSources.UserDataSources.Item("U_ITMNM").Value = oDT.GetValue("ItemName", 0)
                            objForm.DataSources.UserDataSources.Item("U_ITMGRP").Value = oDT.GetValue("ItmsGrpCod", 0)
                        End If
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub LoadGrid(ByVal FormUID As String, ByVal Value As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oGrid = objForm.Items.Item("7").Specific
            oGrid.DataTable = oDt
            oGrid.DataTable.Clear()

            Dim Query As String = " Select '" & Value & "' as ""Check"" , T1.""U_NEWPRC"" As ""New Price"", T1.""U_OLDPRC"" As ""Old Price"",T0.""U_APPSTS"" As ""Status"", T1.""U_ITMCD"" As ""Item Code"" , T0.""U_ITMNM"" As ""Item Name"" , " & _
                                  " T1.""U_BRCHID"" As ""BranchId"" , T1.""U_BRCHNM"" As ""Branch Name"" , ""U_PRCID"" As ""PricListId"" , ""U_PRCLST"" As ""Price List"" , ""U_UOMENT"" As ""UomEntry"" , " & _
                                  " ""U_UOMCD"" As ""UOM"" , T0.""U_USRID"" As ""User Id"" , T0.""DocEntry"" , T2.""IUoMEntry"" As ""DefaultUOM"" , '                                                                                                                                       '    As ""ErrorMessage"" " & _
                                  " From ""@ME_PRICEUPD"" T0 Inner Join ""@ME_PRICEUPD_C0"" T1 On T1.""DocEntry"" = T0.""DocEntry"" " & _
                                  " Inner Join OITM T2 ON T2.""ItemCode"" = T0.""U_ITMCD""  Where IFNULL(T1.""U_ITMCD"",'')<>'' And T1.""U_NEWPRC"" > 0 "
            If objForm.Items.Item("11").Specific.Selected.Value <> "All" Then
                Query = Query & " And T0.""U_APPSTS"" = '" & objForm.Items.Item("11").Specific.Selected.Value & "'"
            End If
            If objForm.DataSources.UserDataSources.Item("U_BRNCH").Value <> "" Then
                Query = Query & " And T1.""U_BRCHNM"" = '" & objForm.Items.Item("6").Specific.Selected.Description & "'"
            End If
            If objForm.DataSources.UserDataSources.Item("U_ITMCD").Value <> "" Then
                Query = Query & " And T1.""U_ITMCD"" = '" & objForm.Items.Item("ITMCD").Specific.Value & "'"
            End If
            If objForm.DataSources.UserDataSources.Item("U_ITMGRP").Value <> "" Then
                Query = Query & " And T2.""ItmsGrpCod"" = '" & objForm.Items.Item("15").Specific.Selected.Value & "'"
            End If
            oGrid.DataTable.ExecuteQuery(Query)
            oGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item("BranchId").Visible = False
            oGrid.Columns.Item("PricListId").Visible = False
            oGrid.Columns.Item("UomEntry").Visible = False
            oGrid.Columns.Item("DocEntry").Visible = False
            oGrid.Columns.Item("DefaultUOM").Visible = False


            If oGrid.Rows.Count = 0 Then
                objMain.objApplication.MessageBox("No records found")
                Exit Sub
            End If
            For i As Integer = 0 To oGrid.Columns.Count - 1
                oGrid.Columns.Item(i).Editable = False
            Next
            oGrid.Columns.Item("Check").Editable = True
            Dim oEditTextCol As SAPbouiCOM.EditTextColumn
            oEditTextCol = oGrid.Columns.Item("Item Code")
            oEditTextCol.LinkedObjectType = 4
            oGrid.AutoResizeColumns()

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub ApprovePrice(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            oGrid = objForm.Items.Item("7").Specific
            If oGrid.DataTable.GetValue("DocEntry", 0).ToString <> "" Then
                Dim FirstValue As String = ""
                objMain.sCmp = objMain.objCompany.GetCompanyService
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OPRICEUPD")
                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                Dim CountFlag As Integer = 0
                For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                    If FirstValue = oGrid.DataTable.GetValue("DocEntry", i).ToString Then
                        CountFlag = 1
                    Else
                        CountFlag = 0
                    End If
                    If CountFlag = 0 And oGrid.DataTable.GetValue("Status", i).ToString = "Submitted" Then
                        Dim Update As String = "Update OITM Set U_CXS_MDFD = 'Y' Where ""ItemCode"" = '" & oGrid.DataTable.GetValue("Item Code", i).ToString & "' "
                        Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsUpdate.DoQuery(Update)
                        objMain.oGeneralParams.SetProperty("DocEntry", oGrid.DataTable.GetValue("DocEntry", i).ToString)
                        objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                        objMain.oGeneralData.SetProperty("U_APPSTS", "Approved")
                        objMain.oGeneralData.SetProperty("U_Updated", "N")
                        objMain.oGeneralData.SetProperty("U_ErrorMsg", "")

                        objMain.oGeneralService.Update(objMain.oGeneralData)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objMain.oGeneralData)
                        FirstValue = oGrid.DataTable.GetValue("DocEntry", i).ToString
                    End If
                Next
                objMain.objApplication.StatusBar.SetText("Approved successfully....", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Me.LoadGrid(objForm.UniqueID, "Y")
            End If
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub UpdatePrice(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("7").Specific
            Dim UomCount As Integer = 0
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Freeze(True)

            objMain.sCmp = objMain.objCompany.GetCompanyService
            objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OPRICEUPD")
            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            Dim CountFlag As Integer = 0
            Dim FirstValue As String = ""
            Dim oItemMaster As SAPbobsCOM.Items
            oProgressBar = Nothing
            Dim r As Integer = 0

            oProgressBar = objMain.objApplication.StatusBar.CreateProgressBar("Initilizing Price Updation...please wait...", oGrid.DataTable.Rows.Count - 1, True)
            oProgressBar.Value = 0
            'oProgressBar.Value = oProgressBar.Value + 1
            oProgressBar.Text = "Initilizing Price Updation...please wait..."
            For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If oGrid.DataTable.GetValue("Check", i) = "Y" Then
                    oItemMaster = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                    If oItemMaster.GetByKey(oGrid.DataTable.GetValue("Item Code", i).ToString) Then
H:                      Dim UomEntry As Integer = CInt(oGrid.DataTable.GetValue("UomEntry", i).ToString)
                        Dim DefaultUom As Integer = CInt(oGrid.DataTable.GetValue("DefaultUOM", i).ToString)
                        'For k As Integer = 0 To oItemMaster.PriceList.Count() - 1
                        Dim PriceListRow As Integer = oGrid.DataTable.GetValue("PricListId", i)
                        oItemMaster.PriceList.SetCurrentLine(PriceListRow - 1)
                        Dim strPrlist As String = oItemMaster.PriceList.PriceList
                        If strPrlist = oGrid.DataTable.GetValue("PricListId", i).ToString Then
                            If DefaultUom = UomEntry Then
                                oItemMaster.PriceList.Price = CDbl(oGrid.DataTable.GetValue("New Price", i).ToString)
                            Else
                                Dim UOMFoundFlag As String = ""
                                For j As Integer = 0 To oItemMaster.PriceList.UoMPrices.Count - 1
                                    oItemMaster.PriceList.UoMPrices.SetCurrentLine(j)
                                    Dim iUom As Integer = oItemMaster.PriceList.UoMPrices.UoMEntry
                                    If iUom = UomEntry Then
                                        oItemMaster.PriceList.UoMPrices.Price = CDbl(oGrid.DataTable.GetValue("New Price", i).ToString)
                                        UOMFoundFlag = "Yes"
                                    End If
                                Next
                                If UOMFoundFlag = "" Then
                                    'oItemMaster.PriceList.UoMPrices.Add()
                                    Dim GetUomPriceListCount As String = "Select Count(*) From ITM9 Where ""ItemCode"" = '" & oItemMaster.ItemCode & "' And ""PriceList"" = '" & strPrlist & "'"
                                    Dim oRsGetUomPriceListCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRsGetUomPriceListCount.DoQuery(GetUomPriceListCount)
                                    If oRsGetUomPriceListCount.RecordCount = 0 Then
                                        UomCount = 0
                                    Else
                                        UomCount = oRsGetUomPriceListCount.Fields.Item(0).Value
                                    End If
                                    oItemMaster.PriceList.UoMPrices.Add()
                                    oItemMaster.PriceList.UoMPrices.SetCurrentLine(UomCount)
                                    oItemMaster.PriceList.UoMPrices.Price = CDbl(oGrid.DataTable.GetValue("New Price", i).ToString)
                                    oItemMaster.PriceList.UoMPrices.UoMEntry = UomEntry
                                    oItemMaster.PriceList.UoMPrices.PriceList = strPrlist
                                    oItemMaster.PriceList.UoMPrices.Auto = SAPbobsCOM.BoYesNoEnum.tNO
                                End If
                            End If
                        End If

                        r = i + 1
                        If r <= oGrid.DataTable.Rows.Count - 1 Then
                            If oGrid.DataTable.GetValue("Item Code", i).ToString = oGrid.DataTable.GetValue("Item Code", r).ToString Then
                                oProgressBar.Value = oProgressBar.Value + i
                                i = i + 1
                                GoTo H
                            End If
                        End If

                        'Next                      
                        Dim RetValue As Integer = oItemMaster.Update()
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oItemMaster)
                        If RetValue <> 0 Then
                            Dim ErrCode As Long
                            Dim ErrMsg As String
                            objMain.objCompany.GetLastError(ErrCode, ErrMsg)
                            Dim s As String = objMain.objCompany.GetLastErrorDescription
                            oGrid.DataTable.SetValue("Status", i, "Error in Updation")
                            oGrid.DataTable.SetValue("ErrorMessage", i, ErrCode.ToString & "  " & s)

                            oProgressBar.Value = oProgressBar.Value + i

                        Else
                            objForm.Freeze(False)
                            oGrid.DataTable.SetValue("Status", i, "Updated")
                            oGrid.DataTable.SetValue("ErrorMessage", i, "")
                            oProgressBar.Value = oProgressBar.Value + i
                            objForm.Freeze(True)
                            If FirstValue = oGrid.DataTable.GetValue("DocEntry", i).ToString Then
                                CountFlag = 1
                            Else
                                CountFlag = 0
                            End If
                            If CountFlag = 0 Then
                                objMain.oGeneralParams.SetProperty("DocEntry", oGrid.DataTable.GetValue("DocEntry", i).ToString)
                                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                                objMain.oGeneralData.SetProperty("U_APPSTS", "Updated")


                                objMain.oGeneralService.Update(objMain.oGeneralData)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objMain.oGeneralData)
                                FirstValue = oGrid.DataTable.GetValue("DocEntry", i).ToString

                                'objMain.objApplication.StatusBar.SetText(i + 1 & " record Updated out of d" & oGrid.DataTable.Rows.Count & " Records", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            End If
                        End If
                    End If
                End If
            Next
            objForm.Freeze(False)
            oProgressBar.Stop()
            If Not oProgressBar Is Nothing Then
                oProgressBar.Stop()
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgressBar)
            objMain.objApplication.StatusBar.SetText("Operation completed successfully ...!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

        Catch ex As Exception
            objForm.Freeze(False)
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            oProgressBar.Stop()
            If Not oProgressBar Is Nothing Then
                oProgressBar.Stop()
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgressBar)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
