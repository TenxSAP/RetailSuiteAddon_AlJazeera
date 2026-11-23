Public Class clsPurchaseRequest

#Region "Declaration"
    Public objForm, objGRPOForm As SAPbouiCOM.Form
    Dim objMatrix, objGRPOMatrix As SAPbouiCOM.Matrix
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim BaseUom As String
    Dim CalcualtionBtnFlag As Boolean = False
    Dim strQry As String = ""
    Dim oItem, oItem1, oItem2, oItem3 As SAPbouiCOM.Item
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim GlobalDocnum As String
    Dim OrsSTRQRY, rs As SAPbobsCOM.Recordset
    Dim Flag As String = "No"
#End Region
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.BeforeAction = True Then
                        Me.AddItems(objform.UniqueID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "BtnItm" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.Before_Action = False Then
                        Try
                            objForm = objMain.objApplication.Forms.Item(FormUID)
                            If objForm.Items.Item("CCode").Specific.Value <> "" Then
                                objMain.objListItemPR.CreateForm(objForm.Items.Item("8").Specific.Value, objForm.Items.Item("CCode").Specific.Value)
                            Else
                                objMain.objApplication.StatusBar.SetText("Please Select Supplier", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                            End If
                        Catch ex As Exception

                        End Try
                    End If

                    If pVal.ItemUID = "fldRtl" Then
                        objForm.PaneLevel = "46"
                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    If pVal.ItemUID = "10000329" And pVal.BeforeAction = False Then
                        Dim objComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("10000329").Specific
                        Dim Doc1 As String = objComboBox.Selected.Description
                        If Doc1 = "22" Then
                            Dim objBaseForm As SAPbouiCOM.Form
                            objBaseForm = objMain.objApplication.Forms.GetForm("142", objMain.objApplication.Forms.ActiveForm.TypeCount)
                            Dim oRsBranch As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim Query As String = " Select Top 1 T1.""LineVendor"" From OPRQ T0  Inner Join " &
                                                                                "PRQ1 T1 On T1.""DocEntry"" = T0.""DocEntry"" Where T0.""DocNum"" = '" & objForm.Items.Item("8").Specific.Value & "'and T1.""LineStatus""='O' "
                            oRsBranch.DoQuery(Query)
                            objMatrix = objForm.Items.Item("38").Specific
                            If oRsBranch.RecordCount > 0 Then
                                Dim vendor As String = oRsBranch.Fields.Item("LineVendor").Value
                                Flag = 1
                                For k As Integer = 1 To objMatrix.VisualRowCount - 1
                                    Dim LVendor As String = objMatrix.Columns.Item("1470002179").Cells.Item(k).Specific.Value

                                    If Flag = 1 Then
                                        If vendor = LVendor Then
                                            Flag = 1
                                        Else
                                            Flag = 0
                                        End If
                                    Else
                                        Exit For
                                    End If
                                Next
                                If Flag = 1 Then
                                    objBaseForm.Items.Item("4").Specific.Value = vendor

                                End If
                            End If
                        End If

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
                    If oCFL.UniqueID = "CFL_Sales" And pVal.BeforeAction = True Then
                        Me.CFLFilter(objForm.UniqueID, "CFL_Sales")
                    End If
                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        If oCFL.UniqueID = "CFL_Sales" Then
                            objForm.Items.Item("CCode").Specific.Value = oDT.GetValue("CardCode", 0)
                            objForm.Items.Item("CName").Specific.Value = oDT.GetValue("CardName", 0)
                        End If
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub AddItems(ByVal FormUID As String)
        Try


            objForm = objMain.objApplication.Forms.Item(FormUID)

            'oItem = objForm.Items.Add("fldRtl", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            'oItem.Top = objForm.Items.Item("112").Top
            'oItem.Left = objForm.Items.Item("112").Left + objForm.Items.Item("112").Width + 10
            'oItem.Width = objForm.Items.Item("112").Width + 30
            'oItem.Height = objForm.Items.Item("112").Height
            'oItem.AffectsFormMode = False
            'Dim fld As SAPbouiCOM.Folder = objForm.Items.Item("fldRtl").Specific
            'fld.Caption = "Retail"
            'fld.GroupWith("112")
            'fld.Pane = "46"


            'objMain.objUtilities.AddLabel(objForm.UniqueID, "lblCus", objForm.Items.Item("2000").Top, objForm.Items.Item("78").Left + 5, objForm.Items.Item("78").Width, "Supplier Code", "84", objForm.Items.Item("78").Height, 46, 46)
            'objMain.objUtilities.AddEditBox(objForm.UniqueID, "CCode", objForm.Items.Item("75").Top, objForm.Items.Item("75").Left + 10, objForm.Items.Item("75").Width, "OPRQ", "U_CCode", "lblCus", objForm.Items.Item("75").Height, 46, 46)

            'objMain.objUtilities.AddLabel(objForm.UniqueID, "lblCN", objForm.Items.Item("78").Top + 22, objForm.Items.Item("78").Left + 5, objForm.Items.Item("78").Width + 13, "Supplier Name", "CCode", objForm.Items.Item("78").Height, 46, 46)
            'objMain.objUtilities.AddEditBox(objForm.UniqueID, "CName", objForm.Items.Item("75").Top + 22, objForm.Items.Item("75").Left + 10, objForm.Items.Item("75").Width, "OPRQ", "U_CName", "lblCN", objForm.Items.Item("75").Height, 46, 46)

            objMain.objUtilities.AddLabel(objForm.UniqueID, "lblCus", objForm.Items.Item("2000").Top + 55, objForm.Items.Item("2000").Left + 5, objForm.Items.Item("2000").Width, "Supplier Code", "2000", objForm.Items.Item("2000").Height, 46, 46)
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "CCode", objForm.Items.Item("2001").Top + 55, objForm.Items.Item("2001").Left, objForm.Items.Item("2001").Width, "OPRQ", "U_CCode", "lblCus", objForm.Items.Item("2001").Height, 46, 46)

            objMain.objUtilities.AddLabel(objForm.UniqueID, "lblCN", objForm.Items.Item("2000").Top + 75, objForm.Items.Item("2000").Left + 5, objForm.Items.Item("2000").Width, "Supplier Name", "2000", objForm.Items.Item("2000").Height, 46, 46)
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "CName", objForm.Items.Item("2001").Top + 75, objForm.Items.Item("2001").Left, objForm.Items.Item("2001").Width, "OPRQ", "U_CName", "lblCN", objForm.Items.Item("2001").Height, 46, 46)

            'objMain.objUtilities.AddButton(objForm.UniqueID, "BtnItm", objForm.Items.Item("2001").Top,
            '                             objForm.Items.Item("2001").Left + 150 + objForm.Items.Item("2").Width - 50, 100, objForm.Items.Item("2").Height, "2001", "&Listing Items", objForm.Items.Item("2").FromPane, objForm.Items.Item("2").ToPane)

            Me.AddChoosefromlist()

            Me.objForm.Items.Item("CCode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("1470000200", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    'If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                    '    objComboBox = objform.Items.Item("40").Specific
                    '    If objComboBox.Value = "C" Then
                    '        objform.Items.Item("Btn_Gen").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    '    Else
                    '        objform.Items.Item("Btn_Gen").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "1", SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                    '    End If
                    'End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                'Dim Validation As String = "Select ""U_PO""as ""PO"" from ""@ME_ACFG"""
                'Dim oLValidation As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oLValidation.DoQuery(Validation)
                'If oLValidation.Fields.Item("PO").Value = "Y" Then
                '    objForm.Items.Item("btn_calFOC").Visible = True
                'Else
                '    objForm.Items.Item("btn_calFOC").Visible = False
                'End If

            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub AddChoosefromlist()
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            Dim Count As Integer = Count + 1
            oCFLCreationParams.ObjectType = "2"
            oCFLCreationParams.UniqueID = "CFL_Sales"
            oCFL = oCFLs.Add(oCFLCreationParams)
            Dim oEditBox As SAPbouiCOM.EditText = objForm.Items.Item("CName").Specific
            oEditBox.ChooseFromListUID = "CFL_Sales"
            oEditBox.ChooseFromListAlias = "CardName"
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub CFLFilter(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Select Case CFL_ID
                Case "CFL_Sales"
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
End Class

