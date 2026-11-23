Imports SAPbobsCOM

Public Class ClsWarehouse

#Region "    Declaration   "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oEditText As SAPbouiCOM.EditText
    Dim oDt As SAPbouiCOM.DataTable
    Dim sSupplierCode, sSupplierName As String
    Dim objComboBox As SAPbouiCOM.ComboBox
#End Region
    Sub CreateForm(ByVal BaseFormUID As String, ByVal ItemCode As String, ByVal ItemName As String, ByVal UOM As String, ByVal Warehouse As String, ByVal Barcode As String, ByVal Qty As Double, ByVal BaseDocEntry As String, ByVal BaseDocNum As String, ByVal Baseline As String)

        Try
            objMain.objUtilities.LoadForm("warehouses.xml", "W_Master", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("W_Master", objMain.objApplication.Forms.ActiveForm.TypeCount)

            If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = False Then
                Retail.MainCls.ohtLookUpForm.Add(objForm.UniqueID, BaseFormUID)
            End If

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_WAREHOUSE")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_WAREHOUSE_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific

            objComboBox = objForm.Items.Item("Series").Specific
            Dim GetSeriesQty As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'TNX_WarehouseUDO' And ""Locked"" = 'N'"
            objMain.objUtilities.ComboBoxLoadValuesSeries(objForm.Items.Item("Series").Specific, GetSeriesQty)

            Dim GetDflt As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'TNX_WarehouseUDO' And ""Locked"" = 'N'"
            Dim oRsGetDflt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDflt.DoQuery(GetDflt)
            objForm.Items.Item("Series").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            Dim CHeckRecord As String = "Select ""DocNum"" From ""@TNX_WAREHOUSE"" Where ""U_BEntry""='" & BaseDocEntry.Trim & "' " &
                                      "And ""U_LineNum""='" & Baseline.Trim & "' And ""DocNum"" IS NOT NULL"

            Dim oRsCHeckRecord As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oRsCHeckRecord.DoQuery(CHeckRecord)


            Dim Branch As String = "Select T0.""BPLid"",T1.""BPLName"" from ""OWHS"" T0 Left Join ""OBPL"" T1 ON T0.""BPLid"" = ""BPLId"" where ""WhsCode""='" & Warehouse & "'
"

            Dim oRsBranch As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oRsBranch.DoQuery(Branch)

            If oRsCHeckRecord.RecordCount > 0 Then

                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE

                objForm.Items.Item("BEntry").Specific.Value = BaseDocEntry.Trim
                objForm.Items.Item("LineNum").Specific.Value = Baseline.Trim


                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)


            Else
                oDBs_Head.SetValue("U_Series", oDBs_Head.Offset, oRsGetDflt.Fields.Item(1).Value)
                oDBs_Head.SetValue("U_ItemCode", oDBs_Head.Offset, ItemCode.Trim)
                oDBs_Head.SetValue("U_ItemName", oDBs_Head.Offset, ItemName.Trim)
                oDBs_Head.SetValue("U_FWhse", oDBs_Head.Offset, Warehouse.Trim)
                oDBs_Head.SetValue("U_BarCode", oDBs_Head.Offset, Barcode.Trim)
                oDBs_Head.SetValue("U_Branch", oDBs_Head.Offset, oRsBranch.Fields.Item("BPLName").Value)
                oDBs_Head.SetValue("U_UOM", oDBs_Head.Offset, UOM.Trim)
                oDBs_Head.SetValue("U_Quantity", oDBs_Head.Offset, Qty)
                oDBs_Head.SetValue("U_DocDate", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TNX_Warehouse"))
                oDBs_Head.SetValue("U_BEntry", oDBs_Head.Offset, BaseDocEntry.Trim)
                oDBs_Head.SetValue("U_BaseNum", oDBs_Head.Offset, BaseDocNum.Trim)
                oDBs_Head.SetValue("U_LineNum", oDBs_Head.Offset, Baseline.Trim)
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

            objForm.Items.Item("ItemCode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("ItemCode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("ItemName").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("ItemName").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("FWhse").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("FWhse").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("BarCode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("BarCode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("Quantity").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Quantity").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("BEntry").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("BEntry").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("BaseNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("BaseNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("LineNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("LineNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_WAREHOUSE")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_WAREHOUSE_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific
            objMatrix.AddRow()
            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
            oDBs_Details.SetValue("U_ToWhse", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Quantity", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Remark", oDBs_Details.Offset, "")

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
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_WAREHOUSE")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_WAREHOUSE_C0")
                    objMatrix = objForm.Items.Item("Mtx").Specific
                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                    '    'If Me.Validation(objForm.UniqueID) = False Then BubbleEvent = False
                    '    If Validation(objForm.UniqueID) = False Then
                    '        BubbleEvent = False
                    '        Exit Sub
                    '    End If
                    'End If

                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                    '    If Me.Validation(objForm.UniqueID) = False Then

                    '        BubbleEvent = False
                    '        Exit Sub
                    '    End If
                    'End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then


                        Dim whseSet As New HashSet(Of String)

                        For i As Integer = 1 To objMatrix.RowCount

                            'If objMatrix.Columns.Item("ToWhse").Cells.Item(i).Specific.value <> "" Then
                            '    If objMatrix.Columns.Item("Quantity").Cells.Item(i).Specific.value = 0 Then

                            '        objMain.objApplication.StatusBar.SetText("Quantity should not be Zero ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            '        BubbleEvent = False
                            '        Exit Sub
                            '    End If


                            'End If



                            Dim toWhse As String = ""

                            Try
                                toWhse = Trim(objMatrix.Columns.Item("ToWhse").Cells.Item(i).Specific.Value)
                            Catch
                                Continue For
                            End Try

                            If String.IsNullOrEmpty(toWhse) Then Continue For

                            If whseSet.Contains(toWhse) Then
                                objMain.objApplication.StatusBar.SetText("Duplicate Warehouse found in row " & i & ".....!!!", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub

                            End If

                            whseSet.Add(toWhse)

                        Next


                    End If


                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_WAREHOUSE")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_WAREHOUSE_C0")
                    objMatrix = objForm.Items.Item("Mtx").Specific

                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If oCFL.UniqueID = "CFL_1" Then
                            '----------------------------------------
                            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row)
                            oDBs_Details.SetValue("U_ToWhse", oDBs_Details.Offset, oDT.GetValue("WhsCode", 0))
                            oDBs_Details.SetValue("U_Quantity", oDBs_Details.Offset, objMatrix.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_Remark", oDBs_Details.Offset, objMatrix.Columns.Item("Remark").Cells.Item(pVal.Row).Specific.Value)
                            objMatrix.SetLineData(pVal.Row)
                            '------------------------------------------
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
                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_WAREHOUSE")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_WAREHOUSE_C0")
                    objMatrix = objForm.Items.Item("Mtx").Specific

                    If pVal.ItemUID = "Mtx" And pVal.ColUID = "ToWhse" And pVal.BeforeAction = True And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        Dim FromwhseDetails As String = objMatrix.Columns.Item("ToWhse").Cells.Item(pVal.Row).Specific.Value
                        Dim FromWhsStock As String = oDBs_Head.GetValue("U_FWhse", 0)
                        If FromwhseDetails = FromWhsStock Then
                            'objMain.objApplication.StatusBar.SetText("Required Quantity should not be more then from warehouse stock", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            objMain.objApplication.StatusBar.SetText("From Warehouse and To Warehouse cannot be same.....!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                            BubbleEvent = False
                        End If

                    End If

                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                    If pVal.ItemUID = "Mtx" And pVal.ColUID = "Col_0" And pVal.Before_Action = False Then
                        objForm = objMain.objApplication.Forms.Item(FormUID)
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_WAREHOUSE")
                        oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_WAREHOUSE_C0")
                        objMatrix = objForm.Items.Item("Mtx").Specific
                        Dim DocumentNumber As String = objMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific.Value

                        Dim qWhs As String = "Select ""DocNum"" from ""@ME_STR"" where ""DocEntry""='" & DocumentNumber & "'"

                        Dim oRsWhs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsWhs.DoQuery(qWhs)

                        objMain.objApplication.ActivateMenuItem("ME_STR")
                        objForm = objMain.objApplication.Forms.GetForm("ME_STR_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                        objForm.Freeze(True)
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objForm.Items.Item("18").Specific.Value = oRsWhs.Fields.Item("DocNum").Value.ToString()

                        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        objForm.Freeze(False)

                    End If


            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    'Function Validation(ByVal FormUID As String)
    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)
    '        objMatrix = objForm.Items.Item("Mtx").Specific
    '        Dim whseSet As New HashSet(Of String)

    '        For i As Integer = 1 To objMatrix.VisualRowCount
    '            Dim toWhse As String = Trim(objMatrix.Columns.Item("ToWhse").Cells.Item(i).Specific.Value)

    '            If toWhse <> "" Then
    '                ' Check duplicate
    '                If whseSet.Contains(toWhse) Then
    '                    objMain.objApplication.StatusBar.SetText("Duplicate Warehouse found in row " & i & ".....!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                    Return False
    '                    ''Exit Function
    '                Else
    '                    whseSet.Add(toWhse)
    '                End If
    '            End If

    '        Next

    '        'If objMatrix.Columns.Item("LineId").Cells.Item(i).Specific.Value <> "" And objMatrix.Columns.Item("Quantity").Cells.Item(i).Specific.Value = 0 Then
    '        '    objMain.objApplication.StatusBar.SetText("Quantity must be greater than 0.....!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        '    Return False

    '        'ElseIf objMatrix.Columns.Item("LineId").Cells.Item(i).Specific.Value <> "" And objMatrix.Columns.Item("ToWhse").Cells.Item(i).Specific.Value = oDBs_Head.GetValue("U_FWhse", 0) Then
    '        '    objMain.objApplication.StatusBar.SetText("From Warehouse and To Warehouse cannot be same.....!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        '    Return False
    '        'End If


    '        Return True
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Function

    Public Function Validation(ByVal FormUID As String) As Boolean
        Try
            Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUID)
            Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx").Specific

            Dim whseSet As New HashSet(Of String)

            For i As Integer = 1 To objMatrix.RowCount

                Dim toWhse As String = ""

                Try
                    toWhse = Trim(objMatrix.Columns.Item("ToWhse").Cells.Item(i).Specific.Value)
                Catch
                    Continue For
                End Try

                If String.IsNullOrEmpty(toWhse) Then Continue For

                ' 🚨 Duplicate check
                If whseSet.Contains(toWhse) Then
                    objMain.objApplication.StatusBar.SetText(
                    "Duplicate Warehouse found in row " & i & ".....!!!",
                    SAPbouiCOM.BoMessageTime.bmt_Long,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    ''Return False
                    Exit Function
                End If

                whseSet.Add(toWhse)

            Next

            Return True

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
            Return False

        End Try
        Exit Function
    End Function

End Class
