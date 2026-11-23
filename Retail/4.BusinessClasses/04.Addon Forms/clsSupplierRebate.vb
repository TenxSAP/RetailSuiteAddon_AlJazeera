Public Class clsSupplierRebate

#Region "    Declaration   "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oEditText As SAPbouiCOM.EditText
    Dim oDt As SAPbouiCOM.DataTable
    Dim sSupplierCode, sSupplierName, sBaseNo, sBaseEntry As String
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim sDoNum As String
#End Region

    Sub CreateForm(ByVal FormUID As String, ByVal CardCode As String, ByVal CardName As String, ByVal BaseEntry As String, ByVal BaseNo As String, ByVal IfSRegExist As String, Optional SRebDocNum As String = "")
        Try
            objMain.objUtilities.LoadForm("SupplierRebateNew.xml", "ME_SPRB_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_SPRB_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)

            If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = False Then
                Retail.MainCls.ohtLookUpForm.Add(objForm.UniqueID, FormUID)
            End If

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_SPRB")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_SPRB_C0")
            objMatrix = objForm.Items.Item("13").Specific

            objComboBox = objForm.Items.Item("8").Specific
            Dim GetSeriesQty As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'ME_SPRBUDO' And ""Locked"" = 'N'"
            objMain.objUtilities.ComboBoxLoadValuesSeries(objForm.Items.Item("8").Specific, GetSeriesQty)
            objForm.Items.Item("8").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            If IfSRegExist = "Y" Then
                objForm.Freeze(True)
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                objForm.Items.Item("9").Specific.value = SRebDocNum
                objForm.Items.Item("4").Specific.value = CardCode
                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                objForm.Freeze(False)
            End If
            If IfSRegExist = "N" Then
                sSupplierCode = CardCode
                sSupplierName = CardName

                sBaseNo = BaseNo
                sBaseEntry = BaseEntry

                objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                sDoNum = objMain.objUtilities.GetNextDocNum(objForm, "ME_OSPRB")
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, sDoNum)
                oDBs_Head.SetValue("U_BPCD", oDBs_Head.Offset, sSupplierCode)
                oDBs_Head.SetValue("U_BPNM", oDBs_Head.Offset, sSupplierName)

                oDBs_Head.SetValue("U_BASENT", oDBs_Head.Offset, sBaseEntry)
                oDBs_Head.SetValue("U_BASENO", oDBs_Head.Offset, sBaseNo)



                Dim GetDflt As String = "Select ""DfltSeries"" From ONNM Where ""ObjectCode"" = 'ME_OSPRB'"
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
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Sub CellMasking(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Items.Item("1000001").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("1000001").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "2", SAPbouiCOM.BoModeVisualBehavior.mvb_True)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_SPRB")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_SPRB_C0")
            objMatrix = objForm.Items.Item("13").Specific

            objMatrix.AddRow()
            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
            oDBs_Details.SetValue("U_POAF", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_POAT", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_RBTPRC", oDBs_Details.Offset, "")

            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, BubbleEvent As Boolean)
        Try
            Select Case pVal.EventTypem
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_SPRB")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_SPRB_C0")
                    objMatrix = objForm.Items.Item("13").Specific

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim chkIfExist As String = "Select ""DocNum"" From ""@ME_SPRB"" Where ""DocEntry""=(Select MAX(""DocEntry"") From ""@ME_SPRB"")"
                        Dim oRschkIfExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRschkIfExist.DoQuery(chkIfExist)
                        If oRschkIfExist.RecordCount > 0 Then
                            objForm.Freeze(True)
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            objForm.Items.Item("9").Specific.value = oRschkIfExist.Fields.Item("DocNum").Value
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            objForm.Freeze(False)
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_SPRB")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_SPRB_C0")
                    objMatrix = objForm.Items.Item("13").Specific
                    If pVal.ItemUID = "13" And pVal.ColUID = "V_0" Then
                        If CDbl(objMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) > 0.0 Then
                            If pVal.Row = objMatrix.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID)
                            End If
                        End If
                    End If
                    If pVal.ItemUID = "13" And pVal.ColUID = "V_2" And pVal.BeforeAction = False Then
                        objForm.Items.Item("RAmount").Visible = False
                    End If
                    If pVal.ItemUID = "13" And pVal.ColUID = "RAmount" And pVal.BeforeAction = False Then
                        objForm.Items.Item("V_2").Visible = False
                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    'If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = True And pVal.BeforeAction = False Then
                    'Retail.MainCls.ohtLookUpForm.Remove(objForm.UniqueID)
                    'End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "Delete Row" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("ME_SPRB_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix = objForm.Items.Item("13").Specific
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
            objForm = objMain.objApplication.Forms.GetForm("ME_SPRB_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    'If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                    '    Me.UpdateBAG(objForm.UniqueID)
                    'End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Sub UpdateBAG(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim UpdatePurchaseAgg As String = "Update ""OOAT"" SET ""U_SUPREBT"" = '" & sDoNum & "'  Where ""AbsID"" ='" & sBaseEntry & "'"
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
                        objMatrix = objForm.Items.Item("13").Specific
                        If eventInfo.ItemUID = "13" And eventInfo.ColUID = "V_-1" Then
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
                        ElseIf eventInfo.ItemUID = "13" And objMatrix.RowCount <= 1 Then
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
                        If eventInfo.ItemUID <> "13" Then
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

End Class
