Imports System
Imports System.Drawing.Drawing2D
Imports SAPbobsCOM
Imports SAPbouiCOM

Public Class clsSubItemList

#Region "    Declaration   "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details, oDS1 As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim matrixid As String = ""
    Dim oEditText As SAPbouiCOM.EditText
    Dim oDt As SAPbouiCOM.DataTable
    Dim sSupplierCode, sSupplierName As String
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim LostFocusFlag As Boolean = False
    Dim MATRIXS As String


#End Region

#Region "    Public Methods   "

    Sub CreateForm(ByVal FormUID As String, ByVal CardCode As String, ByVal CardName As String, ByVal RowId As Integer, ByVal NoOfSKUValue As Double, ByVal DocNumItemListingValue As Integer, ByVal ApplicationIDItemListingValue As String, ByVal IfITLSExist As String)
        Try
            ' Load form and get reference
            objMain.objUtilities.LoadForm("SubItemsList.xml", "ME_SUBITMLST_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_SUBITMLST_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            sSupplierCode = CardCode

            objForm.EnableMenu("1292", True)
            'objForm.EnableMenu("1293", True)
            objForm.EnableMenu("774", True)

            ' Keep lookup references
            If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = False Then
                Retail.MainCls.ohtLookUpForm.Add(objForm.UniqueID, FormUID)
            End If

            ' Datasources and controls
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_LS")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_LS_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific

            ' Series Combo Load
            objComboBox = CType(objForm.Items.Item("12").Specific, SAPbouiCOM.ComboBox)
            Dim GetSeriesQty As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'SBO_LSUDO' And ""Locked"" = 'N'"
            objMain.objUtilities.ComboBoxLoadValuesSeries(objForm.Items.Item("12").Specific, GetSeriesQty)
            objForm.Items.Item("12").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            ' If ITLS exists -> go to FIND mode and assign SKUS after loading
            If IfITLSExist = "Y" Then
                objForm.Freeze(True)
                Try
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    objForm.Items.Item("CardCode").Specific.Value = CardCode
                    objForm.Items.Item("LineNum").Specific.Value = RowId

                    ' Click the default search/find button (Item "1" on your form)
                    objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                    ' Assign SKUS value after the find/load
                    objForm.Items.Item("SKUS").Specific.Value = NoOfSKUValue
                    objForm.Items.Item("DocNumIL").Specific.Value = DocNumItemListingValue
                    objForm.Items.Item("AppIDILG").Specific.Value = ApplicationIDItemListingValue
                Finally
                    objForm.Freeze(False)
                End Try
            End If

            ' If ITLS not exists -> prepare new entry (ADD MODE)
            If IfITLSExist = "N" Then
                sSupplierName = CardName
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE

                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "SBO_LSUDO"))
                oDBs_Head.SetValue("U_CardCode", oDBs_Head.Offset, sSupplierCode)
                oDBs_Head.SetValue("U_CardName", oDBs_Head.Offset, sSupplierName)
                oDBs_Head.SetValue("U_LineNum", oDBs_Head.Offset, RowId)

                ' Set SKUS control value
                objForm.Items.Item("SKUS").Specific.Value = NoOfSKUValue.ToString()
                objForm.Items.Item("DocNumIL").Specific.Value = DocNumItemListingValue.ToString()
                objForm.Items.Item("AppIDILG").Specific.Value = ApplicationIDItemListingValue.ToString()

                ' Default series from ONNM
                Dim GetDflt As String = "Select ""DfltSeries"" From ONNM Where ""ObjectCode"" = 'SBO_LSUDO'"
                Dim oRsGetDflt As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                oRsGetDflt.DoQuery(GetDflt)
                If oRsGetDflt.RecordCount > 0 Then
                    Dim seriesValue = oRsGetDflt.Fields.Item(0).Value
                    oDBs_Head.SetValue("U_SERIES", oDBs_Head.Offset, seriesValue)
                    oDBs_Head.SetValue("Series", oDBs_Head.Offset, seriesValue)
                End If

                ' Clear matrix and add one blank line for entry
                objMatrix.Clear()
                oDBs_Details.Clear()
                objMatrix.FlushToDataSource()
                Me.SetNewLine(objForm.UniqueID)
            End If

            ' Apply masking (readonly/managed attributes)
            Me.CellMasking(objForm.UniqueID)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("clsSubItemList.CreateForm: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

#Region "    UI Helpers   "

    Public Sub CellMasking(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Items.Item("CardCode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("CardCode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("LineNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("LineNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "2", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("DocNumIL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNumIL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNumIL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNumIL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("AppIDILG").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("AppIDILG").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("AppIDILG").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("AppIDILG").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("clsSubItemList.CellMasking: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_LS")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_LS_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific

            objMatrix.AddRow()
            ' Important: Set values using current Offset. If offsets required different logic adjust accordingly.
            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
            oDBs_Details.SetValue("U_ICode", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_IName", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_BarCode", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_IGroup", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_UOM", oDBs_Details.Offset, "")

            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("clsSubItemList.SetNewLine: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

#Region "    Events   "

    Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If (pVal.ItemUID = "Mtx") And pVal.BeforeAction = False Then
                        MATRIXS = pVal.ItemUID
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    ' BEFORE clicking Save/Add (Item "1")
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                        (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                        'If Me.Validation(objForm.UniqueID) = False Then
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If
                    End If
                    'madhu link button open item master upload screen
                    'If pVal.ItemUID = "MasterBtn" And pVal.BeforeAction = False Then
                    '    Try

                    '        objMain.objApplication.ActivateMenuItem("ITEMUP")
                    '    Catch ex As Exception
                    '        objMain.objApplication.StatusBar.SetText("Unable to open Item Master Upload screen: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    End Try
                    'End If

                    If pVal.ItemUID = "MasterBtn" AndAlso pVal.BeforeAction = False Then
                        objMain.objApplication.ActivateMenuItem("ITEMUP")
                        System.Threading.Thread.Sleep(200)
                        Dim oItemForm As SAPbouiCOM.Form
                        oItemForm = objMain.objApplication.Forms.ActiveForm

                        Dim DocNumItemListingValue As String = objForm.Items.Item("DocNumIL").Specific.Value
                        Dim ApplicationIDItemListingValue As String = objForm.Items.Item("AppIDILG").Specific.Value

                        oItemForm.Items.Item("DocNoIL").Specific.Value = DocNumItemListingValue
                        oItemForm.Items.Item("AppIDIL").Specific.Value = ApplicationIDItemListingValue



                    End If



                    ' AFTER successful add (ActionSuccess) - move to find and display created doc
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_LS")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_LS_C0")
                    objMatrix = objForm.Items.Item("Mtx").Specific
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim chkIfExist As String = "Select ""DocNum"" From ""@SBO_LS"" Where ""DocEntry""=(Select MAX(""DocEntry"") From ""@SBO_LS"")"
                        Dim oRschkIfExist As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                        oRschkIfExist.DoQuery(chkIfExist)
                        If oRschkIfExist.RecordCount > 0 Then
                            objForm.Freeze(True)
                            Try
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                objForm.Items.Item("9").Specific.Value = oRschkIfExist.Fields.Item("DocNum").Value
                                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Finally
                                objForm.Freeze(False)
                            End Try
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_LS")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_LS_C0")
                    objMatrix = objForm.Items.Item("Mtx").Specific

                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = CType(pVal, SAPbouiCOM.IChooseFromListEvent)
                    Dim CFL_Id As String = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable = CFLEvent.SelectedObjects

                    If oCFL.UniqueID = "CFL_CFL" And pVal.BeforeAction = True Then
                        Me.CFLFilterItem(FormUID, "CFL_CFL")
                    End If

                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If oCFL.UniqueID = "CFL_CFL" Then
                            '----------------------------------------
                            oDBs_Details.SetValue("LineID", oDBs_Details.Offset, pVal.Row)
                            oDBs_Details.SetValue("U_ICode", oDBs_Details.Offset, oDT.GetValue("ItemCode", 0))
                            oDBs_Details.SetValue("U_IName", oDBs_Details.Offset, oDT.GetValue("ItemName", 0))
                            oDBs_Details.SetValue("U_BarCode", oDBs_Details.Offset, oDT.GetValue("CodeBars", 0))
                            oDBs_Details.SetValue("U_IGroup", oDBs_Details.Offset, oDT.GetValue("ItmsGrpCod", 0))
                            oDBs_Details.SetValue("U_UOM", oDBs_Details.Offset, oDT.GetValue("UgpEntry", 0))

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
            End Select

        Catch ex As Exception
            Try
                objForm.Freeze(False)
            Catch
            End Try
            objMain.objApplication.StatusBar.SetText("clsSubItemList.ItemEvent: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

#Region "    Validation   "


    Function Validation(ByVal FormUID As String) As Boolean
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix = CType(objForm.Items.Item("Mtx").Specific, SAPbouiCOM.Matrix)

            Dim NoOfSKUValue As Integer = 0
            Integer.TryParse(objForm.Items.Item("SKUS").Specific.Value, NoOfSKUValue)

            Dim skuCount As Integer = objMatrix.VisualRowCount - 1

            If skuCount <> NoOfSKUValue Then
                objMain.objApplication.StatusBar.SetText("Error: You must add exactly " & NoOfSKUValue & " SKUs. Currently added: " & skuCount, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Return False
            End If

            Return True

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("clsSubItemList.Validation: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            Return False
        End Try
    End Function


#End Region

#Region "    CFL Filter   "

    Sub CFLFilterItem(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim Query As String = " SELECT T1.""ItemCode"", T1.""VendorCode"" FROM OITM T0  INNER JOIN ITM2 T1 ON T0.""ItemCode"" = T1.""ItemCode"" WHERE T1.""VendorCode""='" & objForm.Items.Item("CardCode").Specific.Value.ToString().Trim() & "'"
            Dim oRsGetSelectionItems As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRsGetSelectionItems.DoQuery(Query)

            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()

            If oRsGetSelectionItems.RecordCount > 0 Then
                ' Add first condition
                oCondition = oConditions.Add()
                oCondition.Alias = "ItemCode"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = oRsGetSelectionItems.Fields.Item(0).Value
                oChooseFromList.SetConditions(oConditions)

                ' Move next and add OR conditions for remaining records if any
                oRsGetSelectionItems.MoveNext()
                For i As Integer = 1 To oRsGetSelectionItems.RecordCount - 1
                    oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCondition = oConditions.Add()
                    oCondition.Alias = "ItemCode"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = oRsGetSelectionItems.Fields.Item(0).Value
                    oChooseFromList.SetConditions(oConditions)
                    oRsGetSelectionItems.MoveNext()
                Next
            Else
                oCondition = oConditions.Add()
                oCondition.Alias = "ItemCode"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = ""
                oChooseFromList.SetConditions(oConditions)
            End If
        Catch ex As Exception
            ' silent swallow as original code; if you want logging enable below
            ' objMain.objApplication.StatusBar.SetText("clsSubItemList.CFLFilterItem: " & ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ITEMLIST" And pVal.BeforeAction = False Then

            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                Me.SetNewLine(objForm.UniqueID)

                'ElseIf pVal.MenuUID = "1293" And pVal.BeforeAction = False Then

                '    objMatrix = objForm.Items.Item("Mtx").Specific
                '    For i As Integer = 1 To objMatrix.VisualRowCount - 1
                '        If objMatrix.IsRowSelected(i) = True Then
                '            objMatrix.DeleteRow(i)
                '        End If
                '    Next
                '    For i As Integer = 1 To objMatrix.VisualRowCount
                '        objMatrix.Columns.Item("LineId").Cells.Item(i).Specific.Value = i
                '    Next

                'End If
                'ElseIf pVal.MenuUID = "774" And pVal.BeforeAction = False Then
                '    objForm = objMain.objApplication.Forms.GetForm("ME_SUBITMLST_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                '    objMatrix = objForm.Items.Item("Mtx").Specific

                '    Dim row As Integer = objMatrix.VisualRowCount
                '    If MATRIXS.Equals("Mtx") = True Then
                '        If objMatrix.IsRowSelected(1) <> True And objMatrix.VisualRowCount < 1 Then
                '            objMatrix.AddRow()
                '            oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
                '            objMatrix.SetLineData(objMatrix.VisualRowCount)
                '        End If
                '        If objMatrix.IsRowSelected(row) = True Then
                '            objMatrix.DeleteRow(row)
                '        Else
                '            For i As Integer = 1 To objMatrix.VisualRowCount - 1
                '                If objMatrix.IsRowSelected(i) = True Then
                '                    If i <> 1 Then
                '                        objMatrix.DeleteRow(i)
                '                    End If
                '                End If
                '            Next
                '        End If
                '    End If
                '    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                'End If
            ElseIf pVal.MenuUID = "774" And pVal.BeforeAction = False Then ' delete row
                objForm = objMain.objApplication.Forms.GetForm("ME_SUBITMLST_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix = objForm.Items.Item("Mtx").Specific
                Dim row As Integer = objMatrix.VisualRowCount
                If MATRIXS.Equals("Mtx") = True Then
                    If objMatrix.IsRowSelected(1) <> True And objMatrix.VisualRowCount < 1 Then
                        objMatrix.AddRow()
                        oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
                        objMatrix.SetLineData(objMatrix.VisualRowCount)
                    End If
                    If objMatrix.IsRowSelected(row) = True Then
                        objMatrix.DeleteRow(row)
                    Else
                        For i As Integer = 1 To objMatrix.VisualRowCount - 1

                            If objMatrix.IsRowSelected(i) = True Then
                                '  If i <> 1 Then
                                objMatrix.DeleteRow(i)
                                '   End If
                            End If
                        Next
                    End If
                    For i As Integer = 1 To objMatrix.VisualRowCount
                        objMatrix.Columns.Item("LineId").Cells.Item(i).Specific.Value = i

                    Next
                End If
            End If


            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If



        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

#End Region

#Region "    Helpers   "

    Private Function SafeGetInteger(ByVal input As String) As Integer
        Dim val As Integer = 0
        Integer.TryParse(input, val)
        Return val
    End Function

#End Region

End Class
