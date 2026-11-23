Public Class clsPaymentDetails

#Region "    Declaration   "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oEditText As SAPbouiCOM.EditText
    Dim oDt As SAPbouiCOM.DataTable
    Dim sSupplierCode, sSupplierName, sBaseNo, sBaseEntry As String
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim sDoNum As String
    Dim sType As String
    Dim sQty As Integer
    Public dAmount As Double
    Public sAccountCode As String
    Public sAccountName As String
    Public sDescription As String
    Public iBaseNo As String
#End Region

    Sub CreateForm(ByVal FormUID As String, ByVal IfSRegExist As String, ByVal CardCode As String, ByVal CardName As String, _
                   ByVal BaseEntry As String, ByVal BaseNo As String, ByVal FromDate As String, ToDate As String, Optional Qty As Integer = 0, _
                   Optional Amount As Double = 0, Optional SRebDocNum As String = "", Optional BranchId As String = "", _
                   Optional Type As String = "", Optional AccountCode As String = "", Optional AccountName As String = "", _
                   Optional Description As String = "")
        Try
            objMain.objUtilities.LoadForm("PaymentDetailsScreen.xml", "ME_PDS_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_PDS_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            sQty = Qty
            dAmount = Amount
            sType = Type
            iBaseNo = BaseNo
            sAccountCode = AccountCode
            sAccountName = AccountName
            sDescription = Description

            If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = False Then
                Retail.MainCls.ohtLookUpForm.Add(objForm.UniqueID, FormUID)
            End If

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PDS")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PDS_C0")
            objMatrix = objForm.Items.Item("3").Specific

            objComboBox = objForm.Items.Item("9").Specific
            Dim GetSeriesQty As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'ME_OPDS' And ""Locked"" = 'N'"
            objMain.objUtilities.ComboBoxLoadValuesSeries(objForm.Items.Item("9").Specific, GetSeriesQty)
            objForm.Items.Item("9").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            objComboBox = objForm.Items.Item("19").Specific
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, " Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join " &
                                                                 " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""  ")
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("19").DisplayDesc = True

            Me.CFLFilter("CFL_ACT")

            objComboBox = objForm.Items.Item("26").Specific
            objComboBox.ValidValues.Add("", "")
            objComboBox.ValidValues.Add("Open", "Open")
            objComboBox.ValidValues.Add("Close", "Close")
            objComboBox.ValidValues.Add("Cancel", "Cancel")
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("26").DisplayDesc = True

            If sType = "Manual" Then
                objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Else
                objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            End If


            If IfSRegExist = "Y" Then
                objForm.Freeze(True)
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                objForm.Items.Item("10").Specific.value = SRebDocNum
                objForm.Items.Item("13").Specific.value = iBaseNo
                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                objForm.Freeze(False)
            End If
            If IfSRegExist = "N" Then

                sSupplierCode = CardCode
                sSupplierName = CardName

                sBaseNo = BaseNo
                sBaseEntry = BaseEntry

                objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                sDoNum = objMain.objUtilities.GetNextDocNum(objForm, "ME_OPDS")
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, sDoNum)
                oDBs_Head.SetValue("U_STATUS", oDBs_Head.Offset, "Open")
                oDBs_Head.SetValue("U_MESPCD", oDBs_Head.Offset, sSupplierCode)
                oDBs_Head.SetValue("U_MESPNM", oDBs_Head.Offset, sSupplierName)

                oDBs_Head.SetValue("U_STARTDT", oDBs_Head.Offset, FromDate)
                oDBs_Head.SetValue("U_ENDDT", oDBs_Head.Offset, ToDate)

                oDBs_Head.SetValue("U_BASENT", oDBs_Head.Offset, sBaseEntry)
                oDBs_Head.SetValue("U_BASENO", oDBs_Head.Offset, sBaseNo)

                Dim GetDflt As String = "Select ""DfltSeries"" From ONNM Where ""ObjectCode"" = 'ME_OPDS'"
                Dim oRsGetDflt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetDflt.DoQuery(GetDflt)

                oDBs_Head.SetValue("U_SERIES", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)
                oDBs_Head.SetValue("Series", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)

                Dim objBranchCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("19").Specific
                objBranchCombo.Select(BranchId, SAPbouiCOM.BoSearchKey.psk_ByDescription)

                objMatrix.Clear()
                oDBs_Details.Clear()
                objMatrix.FlushToDataSource()
                Me.SetNewLine(objForm.UniqueID)
            End If
            Me.CellMasking(objForm.UniqueID)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub CellMasking(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Items.Item("16").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("16").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("16").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "1", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("21").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("21").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("26").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("26").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PDS")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PDS_C0")
            objMatrix = objForm.Items.Item("3").Specific

            objMatrix.AddRow()
            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
            oDBs_Details.SetValue("U_MEDATE", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_MEACCTCD", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_MEACCTNM", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_MEAMT", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_MEJENO", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_MESTS", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_MERMKS", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_REF", oDBs_Details.Offset, "")

            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PDS")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PDS_C0")
                    objMatrix = objForm.Items.Item("3").Specific

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim chkIfExist As String = "Select ""DocNum"" From ""@ME_PDS"" Where ""DocEntry""=(Select MAX(""DocEntry"") From ""@ME_PDS"")"
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
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And _
                                  (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.Validation(objForm.UniqueID) = False Then
                            BubbleEvent = False
                        End If
                    End If

                    If pVal.ItemUID = "16" And pVal.BeforeAction = False Then
                        'Me.PostJE(objForm.UniqueID)
                        Dim oComboStatus As SAPbouiCOM.ComboBox = objForm.Items.Item("26").Specific
                        If oComboStatus.Value = "Open" Then
                            Me.PostAPCreditMemo(objForm.UniqueID)
                        End If
                    End If

                    If pVal.ItemUID = "20" And pVal.BeforeAction = False Then
                        Me.LoadDetails(objForm.UniqueID, sQty)
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
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
                        If oCFL.UniqueID = "CFL_ACT" Then
                            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
                            oDBs_Details.SetValue("U_MEDATE", oDBs_Details.Offset, objMatrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_MEACCTCD", oDBs_Details.Offset, oDT.GetValue("AcctCode", 0))
                            oDBs_Details.SetValue("U_MEACCTNM", oDBs_Details.Offset, oDT.GetValue("AcctName", 0))
                            oDBs_Details.SetValue("U_MEAMT", oDBs_Details.Offset, "")
                            oDBs_Details.SetValue("U_MEJENO", oDBs_Details.Offset, "")
                            oDBs_Details.SetValue("U_MESTS", oDBs_Details.Offset, "")
                            oDBs_Details.SetValue("U_MERMKS", oDBs_Details.Offset, "")
                            oDBs_Details.SetValue("U_REF", oDBs_Details.Offset, "")
                            objMatrix.SetLineData(objMatrix.VisualRowCount)
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PDS")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PDS_C0")
                    objMatrix = objForm.Items.Item("3").Specific
                    If pVal.ItemUID = "3" And pVal.ColUID = "V_3" Then
                        If CDbl(objMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) > 0.0 Then
                            If pVal.Row = objMatrix.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID)
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

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "Delete Row" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("ME_PDS_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
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
            objForm = objMain.objApplication.Forms.GetForm("ME_PDS_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Me.UpdateBAG(objForm.UniqueID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        If objForm.Items.Item("26").Specific.Selected.Value = "Close" Then
                            Me.SetCellsEditable(objForm.UniqueID)
                        ElseIf objForm.Items.Item("26").Specific.Selected.Value = "Open" Then
                            Me.SetCellsEditable(objForm.UniqueID)
                        End If
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetCellsEditable(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PDS")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PDS_C0")
            objMatrix = objForm.Items.Item("3").Specific

            If objForm.Items.Item("26").Specific.Selected.Value = "Close" Then
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.CommonSetting.SetRowEditable(i, False)
                Next
                objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Else
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.CommonSetting.SetRowEditable(i, True)
                Next
                objMatrix.Columns.Item("V_-1").Editable = False
                objMatrix.Columns.Item("V_2").Editable = False
                objMatrix.Columns.Item("V_4").Editable = False
                objMatrix.Columns.Item("V_5").Editable = False
                If sType = "Manual" Then
                    objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                Else
                    objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                End If
            End If
            objForm.Freeze(False)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub UpdateBAG(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim UpdatePurchaseAgg As String = "Update ""OOAT"" SET ""U_PAYMENT"" = '" & sDoNum & "'  Where ""AbsID"" ='" & sBaseEntry & "'"
            Dim oRsUpdatePurchaseAgg As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsUpdatePurchaseAgg.DoQuery(UpdatePurchaseAgg)
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
                        objMatrix = objForm.Items.Item("3").Specific
                        If eventInfo.ItemUID = "3" And eventInfo.ColUID = "V_-1" Then
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
                        ElseIf eventInfo.ItemUID = "3" And objMatrix.RowCount <= 1 Then
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
                        If eventInfo.ItemUID <> "3" Then
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

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PDS")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PDS_C0")
            objMatrix = objForm.Items.Item("3").Specific

            Dim objComboBoxBranch As SAPbouiCOM.ComboBox = objForm.Items.Item("19").Specific
            If objComboBoxBranch.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Please select the branch.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            Dim sDate As String = DateTime.Now.ToString("yyyyMMdd")
            sDate = sDate.Insert("4", "-")
            sDate = sDate.Insert("7", "-")
            Dim oJE As SAPbobsCOM.JournalEntries
            For i As Integer = 1 To objMatrix.VisualRowCount
                If objMatrix.Columns.Item("V_3").Cells.Item(i).Specific.Value > 0 And objMatrix.Columns.Item("V_4").Cells.Item(i).Specific.Value = "" And
                    objMatrix.Columns.Item("V_0").Cells.Item(i).Specific.Value <> "" Then
                    oJE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                    oJE.ReferenceDate = sDate
                    oJE.DueDate = sDate
                    oJE.TaxDate = sDate
                    oJE.Memo = objMatrix.Columns.Item("V_6").Cells.Item(i).Specific.Value
                    oJE.Reference = objMatrix.Columns.Item("V_7").Cells.Item(i).Specific.Value
                    oJE.Reference2 = objForm.Items.Item("10").Specific.value & "-" & i
                    oJE.Lines.BPLID = objComboBoxBranch.Value
                    oJE.Lines.ShortName = objForm.Items.Item("5").Specific.value
                    oJE.Lines.Debit = objMatrix.Columns.Item("V_3").Cells.Item(i).Specific.Value
                    oJE.Lines.LineMemo = objMatrix.Columns.Item("V_6").Cells.Item(i).Specific.Value
                    oJE.Lines.Add()
                    oJE.Lines.BPLID = objComboBoxBranch.Value
                    oJE.Lines.AccountCode = objMatrix.Columns.Item("V_1").Cells.Item(i).Specific.Value
                    oJE.Lines.Credit = objMatrix.Columns.Item("V_3").Cells.Item(i).Specific.Value
                    oJE.Lines.LineMemo = objMatrix.Columns.Item("V_6").Cells.Item(i).Specific.Value
                    oJE.Lines.Add()
                    If oJE.Add = 0 Then
                        Dim GetJEDocEntry As String = "Select Max(""TransId"") From OJDT Where ""TransType"" = '30'"
                        Dim oRsGetJEDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetJEDocEntry.DoQuery(GetJEDocEntry)
                        objMain.sCmp = objMain.objCompany.GetCompanyService
                        objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OPDS")
                        objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                        objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                        objMain.oChildren = objMain.oGeneralData.Child("ME_PDS_C0")
                        objMain.oChildren.Item(i - 1).SetProperty("U_MEJENO", oRsGetJEDocEntry.Fields.Item(0).Value.ToString)
                        objMain.oChildren.Item(i - 1).SetProperty("U_MESTS", "Success")
                        objMain.oGeneralService.Update(objMain.oGeneralData)

                        'Me.UpdateDocStatus(objForm.UniqueID)
                    Else
                        objMain.sCmp = objMain.objCompany.GetCompanyService
                        objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OPDS")
                        objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0).ToString)
                        objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                        objMain.oChildren = objMain.oGeneralData.Child("ME_PDS_C0")
                        objMain.oChildren.Item(i - 1).SetProperty("U_MESTS", objMain.objCompany.GetLastErrorDescription)
                        objMain.oGeneralService.Update(objMain.oGeneralData)
                        ' objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorCode & " : " & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oJE)
                End If
            Next
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_PDS"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_PDS_C0"), "DocEntry", oDBs_Details.GetValue("DocEntry", 0))
            objMatrix.LoadFromDataSource()
            objForm.Refresh()
            objMatrix.AutoResizeColumns()
            Me.SetCellsEditable(objForm.UniqueID)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub PostAPCreditMemo(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim RentalDsc As String = sDescription
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PDS")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PDS_C0")
            objMatrix = objForm.Items.Item("3").Specific

            Dim objComboBoxBranch As SAPbouiCOM.ComboBox = objForm.Items.Item("19").Specific
            If objComboBoxBranch.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Please select the branch.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            Dim sDate As String = ""
           
            Dim oPurchaseCreditNotes As SAPbobsCOM.Documents
            objMain.objApplication.StatusBar.SetText("AP Credit Memo Generating...Please wait.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            objForm.Items.Item("16").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            For i As Integer = 1 To objMatrix.VisualRowCount
                If objMatrix.Columns.Item("V_3").Cells.Item(i).Specific.Value > 0 And objMatrix.Columns.Item("V_4").Cells.Item(i).Specific.Value = "" And
                    objMatrix.Columns.Item("V_0").Cells.Item(i).Specific.Value <> "" Then
                    sDate = objMatrix.Columns.Item("V_0").Cells.Item(i).Specific.Value
                    sDate = sDate.Insert("4", "-")
                    sDate = sDate.Insert("7", "-")
                    RentalDsc = objMatrix.Columns.Item("V_6").Cells.Item(i).Specific.Value
                    oPurchaseCreditNotes = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)

                    oPurchaseCreditNotes.CardCode = objForm.Items.Item("5").Specific.value

                    oPurchaseCreditNotes.DocDate = sDate
                    oPurchaseCreditNotes.DocDueDate = sDate
                    oPurchaseCreditNotes.TaxDate = sDate

                    oPurchaseCreditNotes.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                    oPurchaseCreditNotes.BPL_IDAssignedToInvoice = objComboBoxBranch.Value
                    oPurchaseCreditNotes.NumAtCard = iBaseNo & "/" & sDate & "/" & objMatrix.Columns.Item("V_-1").Cells.Item(i).Specific.Value
                    oPurchaseCreditNotes.Comments = RentalDsc
                    oPurchaseCreditNotes.JournalMemo = RentalDsc
                    oPurchaseCreditNotes.Lines.ItemDescription = RentalDsc
                    oPurchaseCreditNotes.Lines.AccountCode = objMatrix.Columns.Item("V_1").Cells.Item(i).Specific.Value
                    oPurchaseCreditNotes.Lines.LineTotal = objMatrix.Columns.Item("V_3").Cells.Item(i).Specific.Value()
                    If oPurchaseCreditNotes.Add = 0 Then
                        Dim GetJEDocEntry As String = "Select Max(""DocEntry"") From ORPC  "
                        Dim oRsGetJEDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetJEDocEntry.DoQuery(GetJEDocEntry)
                        objMain.sCmp = objMain.objCompany.GetCompanyService
                        objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OPDS")
                        objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                        objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                        objMain.oChildren = objMain.oGeneralData.Child("ME_PDS_C0")
                        objMain.oChildren.Item(i - 1).SetProperty("U_APCRMEMO", oRsGetJEDocEntry.Fields.Item(0).Value.ToString)
                        objMain.oChildren.Item(i - 1).SetProperty("U_MESTS", "Success")
                        objMain.oGeneralService.Update(objMain.oGeneralData)

                        Me.UpdateDocStatus(objForm.UniqueID, "Close")
                    Else
                        objMain.sCmp = objMain.objCompany.GetCompanyService
                        objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OPDS")
                        objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0).ToString)
                        objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                        objMain.oChildren = objMain.oGeneralData.Child("ME_PDS_C0")
                        objMain.oChildren.Item(i - 1).SetProperty("U_MESTS", objMain.objCompany.GetLastErrorDescription)
                        objMain.oGeneralService.Update(objMain.oGeneralData)
                        Me.UpdateDocStatus(objForm.UniqueID, "Open")
                        ' objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorCode & " : " & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                End If
            Next
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_PDS"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_PDS_C0"), "DocEntry", oDBs_Details.GetValue("DocEntry", 0))
            objMatrix.LoadFromDataSource()
            objForm.Refresh()
            objMatrix.AutoResizeColumns()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        Finally
            objForm.Items.Item("16").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Me.SetCellsEditable(objForm.UniqueID)
            objMain.objApplication.StatusBar.SetText("Operation sucessfully completed.... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End Try
    End Sub

    Private Sub UpdateDocStatus(ByVal FormUID As String, ByVal Status As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMain.sCmp = objMain.objCompany.GetCompanyService
            objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OPDS")
            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0).ToString)
            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
            objMain.oGeneralData.SetProperty("U_STATUS", Status)
            objMain.oGeneralService.Update(objMain.oGeneralData)
        Catch ex As Exception
        End Try
    End Sub

    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PDS")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PDS_C0")
            objMatrix = objForm.Items.Item("3").Specific
            'Dim ItemsArry(objMatrix.VisualRowCount) As String
            'For i As Integer = 1 To objMatrix.VisualRowCount - 1
            '    ItemsArry(i - 1) = objMatrix.Columns.Item("V_7").Cells.Item(i).Specific.Value
            'Next
            'If HasDuplicates(ItemsArry) = True Then
            '    Return False
            'End If
            Dim oComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("19").Specific
            If oComboBox.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Branch Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            If objMatrix.VisualRowCount > 0 Then
                If objMatrix.Columns.Item("V_1").Cells.Item(1).Specific.Value = "" Then
                    objMain.objApplication.StatusBar.SetText("Matrix Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Private Function HasDuplicates(ByVal arr As Array) As Boolean
        For i As Integer = 0 To arr.Length - 1
            If Not arr(i) Is Nothing Then
                Dim l As Integer = Array.LastIndexOf(arr, arr(i))
                If l <> i Then
                    objMain.objApplication.StatusBar.SetText(" Duplicate item found : " & objMatrix.Columns.Item("V_0").Cells.Item(l + 1).Specific.Value, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    'MessageBox.Show("Duplicate item found : " & objMatrix.Columns.Item("V_7").Cells.Item(l + 1).Specific.Value)
                    Return True
                End If
            End If
        Next
        Return False
    End Function

    Sub CFLFilter(ByVal CFL_ID As String)
        Try
            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objForm.ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            oCondition = oConditions.Add()
            oCondition.Alias = "Postable"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "Y"
            oChooseFromList.SetConditions(oConditions)
            oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCondition = oConditions.Add()
            oCondition.Alias = "U_RENTLAG"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "Y"
            oChooseFromList.SetConditions(oConditions)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadDetails(ByVal FormUID As String, ByVal Qty As Integer)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim Dt As DateTime
            Dim Count As Integer = Qty
            Dim S As String = objForm.Items.Item("21").Specific.value
            Dim D As DateTime = DateSerial(S.Substring(0, 4), S.Substring(4, 2), S.Substring(6, 2))
            Dim sYear As String = Year(D)
            Dim j As Integer = 0
            Dim iMonth As Integer

            Dim GetAddonConfigurationDtls As String = "Select ""U_RNTLACT""  , ""AcctName""  From ""@ME_ACFG""  T0 Inner Join OACT T1 On T1.""AcctCode""  = T0.""U_RNTLACT"" Where IFNULL(""U_RNTLACT"",'')<>''  "
            Dim oRsGetAddonConfigurationDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetAddonConfigurationDtls.DoQuery(GetAddonConfigurationDtls)
            If oRsGetAddonConfigurationDtls.RecordCount = 0 Then
                objMain.objApplication.StatusBar.SetText("Please configure Rental Account in Addon Configuration Screen...!")
                Exit Sub
            End If

            For i As Integer = 1 To Count
                oDBs_Details.SetValue("LineId", oDBs_Details.Offset, i)
                If j < 13 Then
                    j = j + 1
                End If
                If iMonth <= 12 Then
                    iMonth = Month(D) + j - 1
                End If
                If iMonth = 13 Then
                    iMonth = 1
                    Dt = New DateTime(sYear, CInt(iMonth), 1)
                    S = Dt.ToString("yyyyMMdd")
                    D = DateSerial(S.Substring(0, 4), S.Substring(4, 2), S.Substring(6, 2))
                    j = 1
                End If
                If i = 1 Then
                    oDBs_Details.SetValue("U_MEDATE", oDBs_Details.Offset, objForm.Items.Item("21").Specific.value)
                ElseIf i = Count Then
                    oDBs_Details.SetValue("U_MEDATE", oDBs_Details.Offset, objForm.Items.Item("23").Specific.value)
                Else
                    Dt = New DateTime(sYear, CInt(iMonth), 1)
                    oDBs_Details.SetValue("U_MEDATE", oDBs_Details.Offset, Dt.ToString("yyyyMMdd"))
                End If
                oDBs_Details.SetValue("U_MEACCTCD", oDBs_Details.Offset, sAccountCode) 'oRsGetAddonConfigurationDtls.Fields.Item("U_RNTLACT").Value)
                oDBs_Details.SetValue("U_MEACCTNM", oDBs_Details.Offset, sAccountName) 'oRsGetAddonConfigurationDtls.Fields.Item("AcctName").Value)
                oDBs_Details.SetValue("U_MEAMT", oDBs_Details.Offset, dAmount)
                oDBs_Details.SetValue("U_MEJENO", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_MESTS", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_MERMKS", oDBs_Details.Offset, sDescription)
                oDBs_Details.SetValue("U_REF", oDBs_Details.Offset, "")
                objMatrix.SetLineData(i)
                Me.SetNewLine(objForm.UniqueID)
                If iMonth = 12 Then
                    sYear = sYear + 1
                End If
            Next
            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
