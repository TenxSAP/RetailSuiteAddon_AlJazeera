Public Class clsPriceUpdateBranchAuthorization

#Region "    Declaration   "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oEditText As SAPbouiCOM.EditText
    Dim oDt As SAPbouiCOM.DataTable
    Dim sUserCode, sUserName, sBaseNo, sBaseEntry As String
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim oGrid As SAPbouiCOM.Grid
    Dim sDoNum As String
#End Region


  


    Sub CreateForm(ByVal BaseFormUID As String, ByVal UserCode As String, ByVal UserName As String, ByVal IfDocExist As String, Optional ByVal BaseDocNum As String = "")
        Try
            objMain.objUtilities.LoadForm("PriceUpdateBranchAuthorization.xml", "SBO_PUBA_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("SBO_PUBA_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.PaneLevel = 1
            If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = False Then
                Retail.MainCls.ohtLookUpForm.Add(objForm.UniqueID, BaseFormUID)
            End If
            sUserCode = UserCode
            sUserName = UserName
            sDoNum = BaseDocNum
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PUBA")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_PUBA_C0")
            objMatrix = objForm.Items.Item("Item_6").Specific

            objMatrix.CommonSetting.EnableArrowKey = True

            objComboBox = objForm.Items.Item("Item_10").Specific
            Dim GetSeriesQty As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'SBO_OPUBA' And ""Locked"" = 'N' Order By ""Series"" Desc "
            objMain.objUtilities.ComboBoxLoadValuesSeries(objForm.Items.Item("Item_10").Specific, GetSeriesQty)
            objForm.Items.Item("Item_10").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly


            If IfDocExist = "Y" Then
                objForm.Freeze(True)
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                objForm.Items.Item("Item_7").Specific.value = sDoNum
                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                objForm.Freeze(False)
            End If

            If IfDocExist = "N" Then
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                sDoNum = objMain.objUtilities.GetNextDocNum(objForm, "SBO_OPUBA")
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, sDoNum)
                oDBs_Head.SetValue("U_USRCD", oDBs_Head.Offset, sUserCode)
                oDBs_Head.SetValue("U_USRNM", oDBs_Head.Offset, sUserName)

                Dim GetDflt As String = "Select ""DfltSeries"" From ONNM Where ""ObjectCode"" = 'SBO_OPUBA'"
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

            objForm.Items.Item("Item_1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("Item_3").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_3").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("Item_7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("Item_10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Item_10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "2", SAPbouiCOM.BoModeVisualBehavior.mvb_True)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PUBA")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_PUBA_C0")
            objMatrix = objForm.Items.Item("Item_6").Specific

            objMatrix.AddRow()
            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
            oDBs_Details.SetValue("U_BPLID", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_BPLCD", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_BPLNM", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_ACCESS", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_PRICEDIT", oDBs_Details.Offset, "")
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
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PUBA")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_PUBA_C0")
                    objMatrix = objForm.Items.Item("Item_6").Specific

                    If pVal.ItemUID = "Item_5" And pVal.Before_Action = False Then
                        Me.LoadData(objForm.UniqueID)
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And _
                                 (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.Validation(objForm.UniqueID) = False Then
                            BubbleEvent = False
                        End If
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim LoadSameForm As String = "SELECT MAX(""DocNum"") FROM ""@SBO_PUBA"""
                        Dim oRsLoadSameForm As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsLoadSameForm.DoQuery(LoadSameForm)
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objForm.Items.Item("Item_7").Specific.Value = oRsLoadSameForm.Fields.Item(0).Value
                        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
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

    Public Sub LoadData(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PUBA")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_PUBA_C0")
            objMatrix = objForm.Items.Item("Item_6").Specific

            Dim GetBranchDtls As String = ""
            Dim Branch As String = String.Empty
            Try
                For i As Integer = 1 To objMatrix.VisualRowCount
                    If objMatrix.Columns.Item("Col_0").Cells.Item(i).Specific.Value <> "" Then
                        Branch = Branch & "'" & objMatrix.Columns.Item("Col_0").Cells.Item(i).Specific.Value & "'" & ","
                    End If
                Next
                If Branch.Contains(",") = False Then
                    GetBranchDtls = "Select ""BPLId"" , ""BPLName"" , ""BPLFrName"" From OBPL  "
                Else
                    Branch = Branch.Remove(Branch.LastIndexOf(","))
                    GetBranchDtls = " Select ""BPLId"" , ""BPLName"" , ""BPLFrName"" From OBPL Where ""BPLId"" Not In(" & Branch & ") "
                End If
            Catch ex As Exception
            End Try
            Dim oRsGetBranchDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetBranchDtls.DoQuery(GetBranchDtls)

            If oRsGetBranchDtls.RecordCount > 0 Then
                objMatrix.Clear()
                oDBs_Details.Clear()
                objMatrix.FlushToDataSource()
                For i As Integer = 1 To oRsGetBranchDtls.RecordCount
                    objMatrix.AddRow()
                    oDBs_Details.SetValue("U_BPLID", oDBs_Details.Offset, oRsGetBranchDtls.Fields.Item("BPLId").Value)
                    oDBs_Details.SetValue("U_BPLCD", oDBs_Details.Offset, oRsGetBranchDtls.Fields.Item("BPLName").Value)
                    oDBs_Details.SetValue("U_BPLNM", oDBs_Details.Offset, oRsGetBranchDtls.Fields.Item("BPLFrName").Value)
                    oDBs_Details.SetValue("U_ACCESS", oDBs_Details.Offset, "N")
                    oDBs_Details.SetValue("U_PRICEDIT", oDBs_Details.Offset, "N")
                    objMatrix.SetLineData(objMatrix.VisualRowCount)
                    oRsGetBranchDtls.MoveNext()
                Next
                objMatrix.AutoResizeColumns()
            End If
            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PUBA")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_PUBA_C0")
            objMatrix = objForm.Items.Item("Item_6").Specific

            If objForm.Items.Item("Item_1").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("User Code  Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("Item_3").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("User Name Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

End Class
