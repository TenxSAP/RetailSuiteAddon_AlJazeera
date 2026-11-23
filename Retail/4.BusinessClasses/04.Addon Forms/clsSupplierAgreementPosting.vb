Public Class clsSupplierAgreementPosting
#Region "    Declaration   "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details, oDBs_Details1, oDBs_Details2 As SAPbouiCOM.DBDataSource
    Dim objMatrix, objMatrix1, objMatrix2 As SAPbouiCOM.Matrix
    Dim oEditText As SAPbouiCOM.EditText
    Dim oDt As SAPbouiCOM.DataTable
    Dim sSupplierCode, sSupplierName, sBaseNo, sBaseEntry As String
    Dim objComboBox, objComboBox1, objComboBox2 As SAPbouiCOM.ComboBox
    Dim GlobNum As String
    Dim sDoNum As String
    Dim BaseCardCode As String
    Dim temp As String
#End Region
    '
    Sub CreateForm(ByVal FormUID As String, ByVal CardCode As String, ByVal CardName As String, ByVal BaseEntry As String, ByVal BaseNo As String, ByVal IfSRegExist As String, Optional SRebDocNum As String = "", Optional BaseType As String = "")
        Try
            objMain.objUtilities.LoadForm("SupplierAgreementPosting.xml", "ME_SUP_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_SUP_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = False Then
                Retail.MainCls.ohtLookUpForm.Add(objForm.UniqueID, FormUID)
            End If

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_SUP")
            ' oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_SUP_C0")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_SUP_C1")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@ME_SUP_C2")

            ' objMatrix = objForm.Items.Item("Mtx").Specific
            objMatrix1 = objForm.Items.Item("Mtx1").Specific
            objMatrix2 = objForm.Items.Item("Mtx2").Specific

            ' Set Current Date
            'Dim currentDate As String = DateTime.Now.ToString("ddMMyyyy")
            'oDBs_Head.SetValue("U_DocDate", oDBs_Head.Offset, currentDate)
            oDBs_Head.SetValue("U_DocDate", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))

            'Enabling Print Layout Icon
            'Try
            '    objForm.ReportType = "ME_SUP_Form"

            '    Try
            '        CreateLayoutFolder_ResgistrationCharges()
            '    Catch ex As Exception
            '    End Try

            '    Try
            '        RemoveHandler objMain.objApplication.LayoutKeyEvent, AddressOf Me.HandleLayoutKeyEvent
            '    Catch
            '    End Try
            '    Try
            '        AddHandler objMain.objApplication.LayoutKeyEvent, AddressOf Me.HandleLayoutKeyEvent
            '    Catch ex As Exception
            '    End Try
            'Catch ex As Exception
            'End Try

            'Me.objForm.EnableMenu("1292", True)

            'Dim oRsBranch As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'Dim Query As String = "SELECT T2.""BPLId"",T2.""BPLName"" AS ""Branch"" FROM USR6 T0  INNER JOIN OUSR T1 ON T0.""UserCode"" = T1.""USER_CODE"" " &
            '        " INNER JOIN OBPL T2 ON T0.""BPLId"" = T2.""BPLId"" WHERE T1.""USER_CODE"" = '" & objMain.objCompany.UserName.ToString().Trim & "' "
            'oRsBranch.DoQuery(Query)

            'objComboBox = objForm.Items.Item("BRANCH").Specific
            'objMain.objUtilities.ComboBoxLoadValues(objComboBox, Query)
            'objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            'objForm.Items.Item("BRANCH").DisplayDesc = True
            'If objComboBox.ValidValues.Count <> 0 Then
            '    'objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
            'End If

            objComboBox = objForm.Items.Item("8").Specific
            Dim GetSeriesQty As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'ME_SUPUDO' And ""Locked"" = 'N'"
            objMain.objUtilities.ComboBoxLoadValuesSeries(objForm.Items.Item("8").Specific, GetSeriesQty)
            objForm.Items.Item("8").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.PaneLevel = 1
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
                BaseCardCode = CardCode
                sBaseNo = BaseNo
                sBaseEntry = BaseEntry
                GlobNum = BaseType
                Dim GetDflt As String = "Select ""DfltSeries"" From ONNM Where ""ObjectCode"" = 'ME_SUPUDO'"
                Dim oRsGetDflt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetDflt.DoQuery(GetDflt)
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                ' sDoNum = objMain.objUtilities.GetNextDocNum(objForm, "ME_SUPUDO")
                sDoNum = objMain.objUtilities.GetNextDocNumSeries(objForm, "ME_SUPUDO", oRsGetDflt.Fields.Item(0).Value)

                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, sDoNum)
                oDBs_Head.SetValue("U_BPCD", oDBs_Head.Offset, sSupplierCode)
                oDBs_Head.SetValue("U_BPNM", oDBs_Head.Offset, sSupplierName)

                oDBs_Head.SetValue("U_BASENT", oDBs_Head.Offset, sBaseEntry)
                oDBs_Head.SetValue("U_BASENO", oDBs_Head.Offset, sBaseNo)

                oDBs_Head.SetValue("U_SERIES", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)
                oDBs_Head.SetValue("Series", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)



                'objMatrix.Clear()
                'oDBs_Details.Clear()
                'objMatrix.FlushToDataSource()

                objMatrix1.Clear()
                oDBs_Details1.Clear()
                objMatrix1.FlushToDataSource()

                objMatrix2.Clear()
                oDBs_Details2.Clear()
                objMatrix2.FlushToDataSource()

                ' Me.SetNewLine(objForm.UniqueID)
                Me.SetNewLine1(objForm.UniqueID)
                Me.SetNewLine2(objForm.UniqueID)

            End If

            Try
                Dim objComboBox As SAPbouiCOM.ComboBox
                objComboBox = CType(objMatrix1.Columns.Item("BRANCH").Cells.Item(1).Specific, SAPbouiCOM.ComboBox)
                Dim sQuery As String = "SELECT T0.""BPLId"", T1.""BPLName"" " & "FROM USR6 T0 " & "INNER JOIN OBPL T1 ON T1.""BPLId"" = T0.""BPLId"" " &
                                       "WHERE T0.""UserCode"" = '" & objMain.objCompany.UserName & "' " & "ORDER BY T0.""BPLId"""

                objMain.objUtilities.ComboBoxLoadValues(objComboBox, sQuery)

                objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                If objComboBox.ValidValues.Count > 0 AndAlso String.IsNullOrEmpty(objComboBox.Value) Then
                    objComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                End If

            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try


            'Try
            '    objComboBox1 = objMatrix1.Columns.Item("BRANCH").Cells.Item(1).Specific  ''objForm.Items.Item("14").Specific
            '    Dim sQury As String = " Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join " &
            '                          " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""  "
            '    objMain.objUtilities.ComboBoxLoadValues(objComboBox1, sQury)
            '    objComboBox1.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            '    If objComboBox1.ValidValues.Count <> 0 Then
            '        objComboBox1.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
            '    End If
            'Catch ex As Exception

            'End Try

            Try
                objComboBox2 = objMatrix1.Columns.Item("VATCODE").Cells.Item(1).Specific  ''objForm.Items.Item("14").Specific
                ' Dim sQury As String = "SELECT ""Code"",""Name"" FROM OVTG WHERE IFNULL(""Account"",'')<>'' and ""Locked""='N'"
                '--madhu based on logic changed query
                Dim sQury As String = "SELECT ""Code"",""Name"" FROM OVTG WHERE IFNULL(""Account"",'')<>'' AND ""Locked""='N' AND UPPER(TRIM(""Name"")) NOT LIKE '%OUTPUT%'"
                objMain.objUtilities.ComboBoxLoadValues(objComboBox2, sQury)
                objComboBox2.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            Catch ex As Exception

            End Try
            objMatrix1.Columns.Item("VATAMT").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix1.Columns.Item("AMT").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

            objForm.Items.Item("Item_2").Width = 200
            objForm.Items.Item("Item_4").Width = 200
            Me.CellMasking(objForm.UniqueID)
            ' Me.DisableColumns()
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
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
    'Sub SetNewLine(ByVal FormUID As String)
    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)
    '        oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_SUP")
    '        oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_SUP_C0")
    '        objMatrix = objForm.Items.Item("Mtx").Specific

    '        objMatrix.AddRow()
    '        oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
    '        oDBs_Details.SetValue("U_DESC", oDBs_Details.Offset, "")
    '        oDBs_Details.SetValue("U_AMT", oDBs_Details.Offset, "")
    '        oDBs_Details.SetValue("U_PTG", oDBs_Details.Offset, "")
    '        oDBs_Details.SetValue("U_REMARKS", oDBs_Details.Offset, "")
    '        oDBs_Details.SetValue("U_JESTATUS", oDBs_Details.Offset, "")
    '        oDBs_Details.SetValue("U_JE", oDBs_Details.Offset, "")

    '        objMatrix.SetLineData(objMatrix.VisualRowCount)
    '        objMatrix.AutoResizeColumns()
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub
    Sub SetNewLine1(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_SUP")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_SUP_C1")
            objMatrix1 = objForm.Items.Item("Mtx1").Specific

            objMatrix1.AddRow()
            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
            oDBs_Details1.SetValue("U_DESC", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_AMT", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_VATCODE", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_VATAMT", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_BRANCH", oDBs_Details1.Offset, "")
            ' oDBs_Details1.SetValue("U_VATACC", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_REMARKS", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_JENum", oDBs_Details1.Offset, "")



            objMatrix1.SetLineData(objMatrix1.VisualRowCount)
            objMatrix1.AutoResizeColumns()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub SetNewLine2(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_SUP")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@ME_SUP_C2")
            objMatrix2 = objForm.Items.Item("Mtx2").Specific

            objMatrix2.AddRow()
            oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, objMatrix2.VisualRowCount)
            oDBs_Details2.SetValue("U_GROUP", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_AMT", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_FCENT", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_BCENT", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_BRANCH", oDBs_Details2.Offset, "")

            objMatrix2.SetLineData(objMatrix2.VisualRowCount)
            objMatrix2.AutoResizeColumns()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                'Case SAPbouiCOM.BoEventTypes.et_CLICK
                '    objForm = objMain.objApplication.Forms.Item(FormUID)
                '    objMatrix2 = objForm.Items.Item("Mtx1").Specific
                '    If pVal.ItemUID = "Mtx1" And pVal.BeforeAction = True Then
                '        If (pVal.ColUID = "DESC" Or pVal.ColUID = "V_0" Or pVal.ColUID = "VATAMT" Or pVal.ColUID = "VATCODE" Or pVal.ColUID = "BRANCH") Then
                '            If objMatrix2.Columns.Item("JENum").Cells.Item(pVal.Row).Specific.Value.ToString() <> "" Then
                '                BubbleEvent = False
                '            End If
                '        End If
                '    End If
                'Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                '    objForm = objMain.objApplication.Forms.Item(FormUID)
                '    objMatrix2 = objForm.Items.Item("Mtx1").Specific
                '    If pVal.ItemUID = "Mtx1" And pVal.BeforeAction = True Then
                '        If (pVal.ColUID = "DESC" Or pVal.ColUID = "V_0" Or pVal.ColUID = "VATAMT" Or pVal.ColUID = "VATCODE" Or pVal.ColUID = "BRANCH") Then
                '            If objMatrix2.Columns.Item("JENum").Cells.Item(pVal.Row).Specific.Value.ToString() <> "" Then
                '                BubbleEvent = False
                '            End If
                '        End If
                '    End If
                'Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED
                '    objForm = objMain.objApplication.Forms.Item(FormUID)
                '    objMatrix2 = objForm.Items.Item("Mtx1").Specific
                '    If pVal.ItemUID = "Mtx1" And pVal.BeforeAction = True Then
                '        If (pVal.ColUID = "DESC" Or pVal.ColUID = "V_0" Or pVal.ColUID = "VATAMT" Or pVal.ColUID = "VATCODE" Or pVal.ColUID = "BRANCH") Then
                '            If objMatrix2.Columns.Item("JENum").Cells.Item(pVal.Row).Specific.Value.ToString() <> "" Then
                '                BubbleEvent = False
                '            End If
                '        End If
                '    End If
                'Case SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK
                '    objForm = objMain.objApplication.Forms.Item(FormUID)
                '    objMatrix2 = objForm.Items.Item("Mtx1").Specific
                '    If pVal.ItemUID = "Mtx1" And pVal.BeforeAction = True Then
                '        If (pVal.ColUID = "DESC" Or pVal.ColUID = "V_0" Or pVal.ColUID = "VATAMT" Or pVal.ColUID = "VATCODE" Or pVal.ColUID = "BRANCH") Then
                '            If objMatrix2.Columns.Item("JENum").Cells.Item(pVal.Row).Specific.Value.ToString() <> "" Then
                '                BubbleEvent = False
                '            End If
                '        End If
                '    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix2 = objForm.Items.Item("Mtx1").Specific
                    '    If pVal.ItemUID = "Mtx1" And pVal.BeforeAction = True Then
                    '        If (pVal.ColUID = "DESC" Or pVal.ColUID = "V_0" Or pVal.ColUID = "VATAMT" Or pVal.ColUID = "VATCODE" Or pVal.ColUID = "BRANCH") Then
                    '            If objMatrix2.Columns.Item("JENum").Cells.Item(pVal.Row).Specific.Value.ToString() <> "" Then
                    '                BubbleEvent = False
                    '            End If
                    '        End If
                    '    End If
                    If pVal.ItemUID = "Mtx1" And pVal.ColUID = "VATCODE" And pVal.BeforeAction = False Then
                        Try
                            objForm.Freeze(True)
                            If objMatrix1.Columns.Item("JENum").Cells.Item(pVal.Row).Specific.Value.ToString() = "" Then
                                Dim vatCode As String = objMatrix1.Columns.Item("VATCODE").Cells.Item(pVal.Row).Specific.Value.ToString()
                                Dim Amount As String = objMatrix1.Columns.Item("AMT").Cells.Item(pVal.Row).Specific.Value.ToString()
                                Dim qStrVatPer As String = "SELECT ""Rate"" FROM OVTG WHERE ""Code""='" + vatCode + "'"
                                Dim rSetVatPer As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                rSetVatPer.DoQuery(qStrVatPer)
                                If rSetVatPer.RecordCount > 0 Then
                                    Dim Rate As Double = CDbl(rSetVatPer.Fields.Item("Rate").Value)
                                    Dim VatAmount As Double = (Rate / 100) * Amount
                                    objMatrix1.Columns.Item("VATAMT").Cells.Item(pVal.Row).Specific.Value = VatAmount.ToString()
                                End If
                                objMatrix1.FlushToDataSource()
                                objMatrix1.LoadFromDataSource()
                            End If
                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                        End Try
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_SUP")
                    ' oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_SUP_C0")
                    'objMatrix = objForm.Items.Item("Mtx").Specific

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim chkIfExist As String = "Select ""DocNum"" From ""@ME_SUP"" Where ""DocEntry""=(Select MAX(""DocEntry"") From ""@ME_SUP"")"
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
                    If pVal.ItemUID = "Post" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        '    Dim GetAddonConfigDtls As String = "Select ""U_Auto"" From ""@ME_ACFG"" "
                        '    Dim oRsGetAddonConfigDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '    oRsGetAddonConfigDtls.DoQuery(GetAddonConfigDtls)
                        '    If oRsGetAddonConfigDtls.RecordCount > 0 Then
                        '        If oRsGetAddonConfigDtls.Fields.Item(0).Value = "Y" Then
                        '            Me.PostJEWithVAT()
                        '        Else
                        '            Me.PostJE()
                        '        End If
                        '    End If
                        Me.PostJEWithVATcode()

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
                    If oCFL.UniqueID = "CFL_0" And pVal.BeforeAction = True Then
                    End If

                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                        If oCFL.UniqueID = "CFL_0" Then
                            objMatrix1.GetLineData(objMatrix1.VisualRowCount)

                            oDBs_Details1.SetValue("U_VATACC", oDBs_Details1.Offset, oDT.GetValue("AcctCode", 0))
                            objMatrix1.SetLineData(objMatrix1.VisualRowCount)
                        End If
                    End If

                    If oCFL.UniqueID = "CFL_1" And pVal.BeforeAction = True Then
                        Me.CFLFilter(FormUID, "CFL_1")

                    End If

                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                        If oCFL.UniqueID = "CFL_1" Then
                            objMatrix1.GetLineData(objMatrix1.VisualRowCount)

                            oDBs_Details1.SetValue("U_DESC", oDBs_Details1.Offset, oDT.GetValue("Name", 0))
                            objMatrix1.SetLineData(objMatrix1.VisualRowCount)
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_SUP")
                    ' oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_SUP_C0")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_SUP_C1")
                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@ME_SUP_C2")
                    'objMatrix = objForm.Items.Item("Mtx").Specific
                    'If pVal.ItemUID = "Mtx" And pVal.ColUID = "DESC" Then
                    '    If objMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim() <> "" Then
                    '        If pVal.Row = objMatrix.VisualRowCount Then
                    '            Me.SetNewLine(objForm.UniqueID)
                    '        End If
                    '    End If
                    'End If
                    'If pVal.ItemUID = "Mtx1" And pVal.ColUID = "DESC" Then
                    '    If objMatrix1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim() <> "" Then
                    '        If pVal.Row = objMatrix1.VisualRowCount Then
                    '            Me.SetNewLine1(objForm.UniqueID)
                    '        End If
                    '    End If
                    'End If
                    'If pVal.ItemUID = "Mtx2" And pVal.ColUID = "GROUP" Then
                    '    If objMatrix2.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim() <> "" Then
                    '        If pVal.Row = objMatrix2.VisualRowCount Then
                    '            Me.SetNewLine2(objForm.UniqueID)
                    '        End If
                    '    End If
                    'End If
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "Mtx1" Or pVal.ItemUID = "Mtx2" Then
                        temp = pVal.ItemUID
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
    Sub CFLFilter(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim Selection As String = "select ""Code"",""Name"",""U_GLAcct"", ""U_SupChrg"" from ""@TNX_SDPT"" where ""U_SupChrg""='Y'"
            oRs.DoQuery(Selection)

            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            If oRs.RecordCount > 0 Then
                oCondition = oConditions.Add()
                oCondition.Alias = "U_SupChrg"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = oRs.Fields.Item("U_SupChrg").Value
                oChooseFromList.SetConditions(oConditions)

            End If
        Catch ex As Exception

        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "Delete Row" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("ME_SUP_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                'objMatrix = objForm.Items.Item("Mtx").Specific
                'For i As Integer = 1 To objMatrix.VisualRowCount - 1
                '    If objMatrix.IsRowSelected(i) = True Then
                '        objMatrix.DeleteRow(i)
                '    End If
                'Next
                'For i As Integer = 1 To objMatrix.VisualRowCount
                '    objMatrix.Columns.Item("V_-1").Cells.Item(i).Specific.string = i
                'Next
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
            If pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("ME_SUP_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                Try
                    objForm.Freeze(True)
                    If temp = "Mtx1" Then
                        Me.SetNewLine1(objForm.UniqueID)
                    Else
                        Me.SetNewLine2(objForm.UniqueID)
                    End If
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                    objForm.Freeze(False)
                Catch ex As Exception
                    objForm.Freeze(False)
                End Try
            ElseIf pVal.MenuUID = "519" And pVal.BeforeAction = True Then
                Try
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_SUP")
                    ' oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_SUP_C0")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_SUP_C1")
                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@ME_SUP_C2")

                    ' objMatrix = objForm.Items.Item("Mtx").Specific
                    objMatrix1 = objForm.Items.Item("Mtx1").Specific
                    objMatrix2 = objForm.Items.Item("Mtx2").Specific


                    Dim frs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'SupplierAgreementPosting' AND ""Type"" = 'C'")

                    If frs.RecordCount = 0 Then
                        objMain.objApplication.MessageBox("SupplierAgreementPosting Sheet Layout('SupplierAgreementPosting') not found. Please upload with same name and Code as parameter", 0, "OK")
                    Else
                        objMain.objApplication.ActivateMenuItem(frs.Fields.Item(0).Value.ToString())
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_SUP")
                        Dim CrForm As SAPbouiCOM.Form
                        Dim oedt As SAPbouiCOM.EditText
                        CrForm = objMain.objApplication.Forms.ActiveForm
                        oedt = CrForm.Items.Item("1000003").Specific
                        oedt.Value = oDBs_Head.GetValue("DocEntry", 0)
                        CrForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If


                Catch ex As Exception
                End Try



            ElseIf pVal.MenuUID = "520" And pVal.BeforeAction = True Then



                Dim LayoutSelection As Integer = objMain.objApplication.MessageBox("Please select the layout for printing", 1, "LC Process", "")
                    If (LayoutSelection = 1) Then

                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_SUP")
                        ' oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_SUP_C0")
                        oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_SUP_C1")
                        oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@ME_SUP_C2")

                        ' objMatrix = objForm.Items.Item("Mtx").Specific
                        objMatrix1 = objForm.Items.Item("Mtx1").Specific
                        objMatrix2 = objForm.Items.Item("Mtx2").Specific



                        Dim frs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'SupplierAgreementPosting' AND ""Type"" = 'C'")

                    If frs.RecordCount = 0 Then
                        objMain.objApplication.MessageBox("SupplierAgreementPosting Sheet Layout('SupplierAgreementPosting') not found. Please upload with same name and Code as parameter", 0, "OK")
                    Else
                            objMain.objApplication.ActivateMenuItem(frs.Fields.Item(0).Value.ToString())
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_SUP")
                        Dim CrForm As SAPbouiCOM.Form
                            Dim oedt As SAPbouiCOM.EditText
                            CrForm = objMain.objApplication.Forms.ActiveForm
                            oedt = CrForm.Items.Item("1000003").Specific
                            oedt.Value = oDBs_Head.GetValue("DocEntry", 0)
                            CrForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        End If

                    End If
                    End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("ME_SUP_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        objMatrix1 = objForm.Items.Item("Mtx1").Specific
                        For i As Integer = 1 To objMatrix1.VisualRowCount
                            If objMatrix1.Columns.Item("JENum").Cells.Item(i).Specific.Value.ToString() <> "" Then
                                objMatrix1.CommonSetting.SetRowEditable(i, False)
                            Else
                                objMatrix1.CommonSetting.SetCellEditable(i, 1, True)
                                objMatrix1.CommonSetting.SetCellEditable(i, 2, True)
                                objMatrix1.CommonSetting.SetCellEditable(i, 3, True)
                                objMatrix1.CommonSetting.SetCellEditable(i, 5, True)
                                objMatrix1.CommonSetting.SetCellEditable(i, 6, True)
                            End If
                        Next

                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        If GlobNum = "BussinessPartner" Then
                            Me.UpdateBAGBP(objForm.UniqueID)
                        Else
                            Me.UpdateBAG(objForm.UniqueID)
                        End If
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Sub UpdateBAG(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim UpdatePurchaseAgg As String = "Update ""OOAT"" SET ""U_SUPPOST"" = '" & sDoNum & "'  Where ""AbsID"" ='" & sBaseEntry & "'"
            Dim oRsUpdatePurchaseAgg As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsUpdatePurchaseAgg.DoQuery(UpdatePurchaseAgg)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Sub UpdateBAGBP(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim UpdatePurchaseAgg As String = "Update ""OCRD"" SET ""U_SUPNUM"" = '" & sDoNum & "'  Where ""CardCode"" ='" & BaseCardCode & "'"
            Dim oRsUpdatePurchaseAgg As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsUpdatePurchaseAgg.DoQuery(UpdatePurchaseAgg)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub PostJE()
        oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_SUP")
        oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_SUP_C1")
        objMatrix = objForm.Items.Item("Mtx1").Specific
        For i As Integer = 1 To objMatrix.VisualRowCount
            Dim supplier As String = oDBs_Head.GetValue("U_BPCD", 0).Trim
            If objMatrix.Columns.Item("JENum").Cells.Item(i).Specific.Value.ToString().Trim = "" And CDbl(objMatrix.Columns.Item("V_0").Cells.Item(i).Specific.Value.ToString().Trim) <> 0 Then
                Dim Query As String = "SELECT ""U_RChar"" FROM ""@ME_ACFG"" "
                Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rs.DoQuery(Query)
                Dim DrAccount As String = rs.Fields.Item(0).Value
                Dim Amount As Double = CDbl(objMatrix.Columns.Item("V_0").Cells.Item(i).Specific.Value.ToString().Trim)
                Dim Remarks As String = objMatrix.Columns.Item("REMARKS").Cells.Item(i).Specific.Value.ToString().Trim
                Dim objComboBoxBranch As SAPbouiCOM.ComboBox = objForm.Items.Item("BRANCH").Specific
                If objComboBoxBranch.Value = "" Then
                    objMain.objApplication.StatusBar.SetText("Please select the branch.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
                Dim oJE As SAPbobsCOM.JournalEntries
                oJE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                oJE.ReferenceDate = Now
                oJE.DueDate = Now
                oJE.TaxDate = Now
                oJE.UserFields.Fields.Item("U_Remarks").Value = Remarks
                oJE.Reference = "Registration charges"
                oJE.Memo = "Registration amount"
                oJE.Lines.BPLID = objComboBoxBranch.Value
                oJE.Lines.ShortName = supplier
                oJE.Lines.Debit = Amount
                oJE.Lines.Add()

                oJE.Lines.AccountCode = DrAccount
                oJE.Lines.Credit = Amount
                oJE.Lines.BPLID = objComboBoxBranch.Value
                oJE.Lines.Add()

                If oJE.Add() = 0 Then
                    objMain.objApplication.StatusBar.SetText("Registration JE Posted Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Dim Query1 As String = "Update ""@ME_SUP_C1"" Set ""U_JENum""=(Select Max(""TransId"") FROM OJDT) Where ""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0) & "' and ""LineId""='" & i & "'"
                    rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    rs.DoQuery(Query1)
                Else
                    objMain.objApplication.StatusBar.SetText("Registration JE Posting failed" + objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
                Try
                    objMain.objApplication.ActivateMenuItem("1304")
                Catch ex As Exception

                End Try
            End If
        Next
    End Sub
#Region " Right Click Event"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Dim objForm As SAPbouiCOM.Form
        objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)

        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_SUP")
        oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_SUP_C1")
        oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@ME_SUP_C2")

        objMatrix1 = objForm.Items.Item("Mtx1").Specific
        objMatrix2 = objForm.Items.Item("Mtx2").Specific


        Try
            If eventInfo.FormUID = objForm.UniqueID Then
                If (eventInfo.BeforeAction = True) Then
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then 'SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                        '  If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Or objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then

                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then 'SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                            '  objMatrix = objForm.Items.Item("Matx").Specific
                            objForm.EnableMenu("1292", True)
                            ' objForm.EnableMenu("1293", True)
                            If eventInfo.ItemUID = "Mtx1" And eventInfo.ColUID = "LineId" Then
                                Try

                                Catch ex As Exception
                                    objMain.objApplication.StatusBar.SetText(ex.Message)
                                End Try
                            ElseIf eventInfo.ItemUID = "Mtx2" And eventInfo.ColUID = "LineId" Then
                                Try

                                Catch ex As Exception
                                    objMain.objApplication.StatusBar.SetText(ex.Message)
                                End Try
                            End If

                        End If
                    End If
                End If
            End If
            ' System.Diagnostics.Process.Start()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

    Private Sub PostJEWithVATcode()
        Try
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_SUP")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_SUP_C1")
            objMatrix = objForm.Items.Item("Mtx1").Specific
            For i As Integer = 1 To objMatrix.VisualRowCount
                Dim supplier As String = oDBs_Head.GetValue("U_BPCD", 0).Trim
                If objMatrix.Columns.Item("JENum").Cells.Item(i).Specific.Value.ToString().Trim = "" And CDbl(objMatrix.Columns.Item("AMT").Cells.Item(i).Specific.Value.ToString().Trim) <> 0 Then
                    'Dim Query As String = "SELECT ""U_VAT"",""U_VATACC"", ""U_DESC"" FROM ""@ME_SUP_C1"""
                    Dim Query As String = "SELECT ""U_AMT"",""U_VATACC"",""U_VATCODE"",""U_VATAMT"" ""U_DESC"" FROM ""@ME_SUP_C1"""
                    Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    rs.DoQuery(Query)
                    Dim Description As String = objMatrix.Columns.Item("DESC").Cells.Item(i).Specific.Value

                    Dim Query2 As String = "SELECT ""Code"",""Name"", ""U_GLAcct"", ""U_VATACC"" FROM ""@TNX_SDPT"" where ""U_GLAcct"" <>'' And ""Name"" = '" & Description & "'"
                    Dim rs1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    rs1.DoQuery(Query2)


                    If rs.RecordCount > 0 Then
                        Dim Amount As Double = CDbl(objMatrix.Columns.Item("AMT").Cells.Item(i).Specific.Value.ToString().Trim)
                        Dim VatCode As String = objMatrix.Columns.Item("VATCODE").Cells.Item(i).Specific.Value
                        Dim VatAmount As Double = CDbl(objMatrix.Columns.Item("VATAMT").Cells.Item(i).Specific.Value)

                        Dim DrAccount As String = rs1.Fields.Item("U_GLAcct").Value
                        Dim DrAccount1 As String = rs1.Fields.Item("U_VATACC").Value
                        'Dim DrAccount1 As String = objMatrix.Columns.Item("VATACC").Cells.Item(i).Specific.Value.ToString().Trim
                        Dim Remarks As String = objMatrix.Columns.Item("REMARKS").Cells.Item(i).Specific.Value.ToString().Trim
                        Dim objComboBoxBranch As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("BRANCH").Cells.Item(i).Specific
                        If objComboBoxBranch.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please select the branch.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Sub
                        End If
                        Dim oJE As SAPbobsCOM.JournalEntries
                        oJE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                        ' oJE.ReferenceDate = Now
                        oJE.AutoVAT = SAPbobsCOM.BoYesNoEnum.tYES
                        oJE.DueDate = DateAdd(DateInterval.Day, 90, Now) ''Now
                        oJE.TaxDate = Now
                        oJE.UserFields.Fields.Item("U_Remarks").Value = Remarks
                        oJE.Reference = "Registration charges"
                        oJE.Reference2 = oDBs_Head.GetValue("DocNum", 0).Trim

                        Dim DocumentDate As String = oDBs_Head.GetValue("U_DocDate", 0).Trim()
                        oJE.ReferenceDate = Date.ParseExact(DocumentDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture)

                        oJE.Memo = "Registration amount"
                        oJE.Lines.BPLID = objComboBoxBranch.Value
                        oJE.Lines.ShortName = supplier
                        oJE.Lines.Debit = Amount + VatAmount
                        oJE.Lines.Add()

                        oJE.Lines.AccountCode = DrAccount
                        oJE.Lines.Credit = Amount
                        oJE.Lines.BPLID = objComboBoxBranch.Value
                        oJE.Lines.TaxGroup = VatCode
                        oJE.Lines.Add()


                        If oJE.Add() = 0 Then
                            objMain.objApplication.StatusBar.SetText("Registration JE Posted Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Dim Query1 As String = "Update ""@ME_SUP_C1"" Set ""U_JENum""=(Select Max(""TransId"") FROM OJDT) Where ""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0) & "' and ""LineId""='" & i & "'"
                            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            rs.DoQuery(Query1)
                        Else
                            objMain.objApplication.StatusBar.SetText("Registration JE Posting failed" + objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If

                        Try
                            objMain.objApplication.ActivateMenuItem("1304")
                        Catch ex As Exception

                        End Try
                    Else
                        objMain.objApplication.StatusBar.SetText("Please Select the VatCode in Supplier Charges.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If
                End If
            Next


        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    'Enabling Print Layout Icon
    Public Sub HandleLayoutKeyEvent(ByRef eventInfo As SAPbouiCOM.LayoutKeyInfo, ByRef BubbleEvent As Boolean)

        Try

            If eventInfo.BeforeAction = False Then
                Dim f As SAPbouiCOM.Form = Nothing
                Try
                    f = objMain.objApplication.Forms.Item(eventInfo.FormUID)
                Catch
                    f = Nothing
                End Try

                If f Is Nothing Then Return

                If f.TypeEx = "ME_SUP_Form" OrElse f.TypeEx = "ME_SUP_Form" Then
                    Try
                        'Dim txtObj As SAPbouiCOM.EditText = CType(f.Items.Item("DocEntry").Specific, SAPbouiCOM.EditText)
                        'Dim docVal As String = If(txtObj IsNot Nothing, txtObj.Value.Trim(), "")
                        Dim docVal As String
                        If String.IsNullOrWhiteSpace(docVal) Then
                            ' fallback: try header datasource DocNum
                            Try
                                Dim ds As SAPbouiCOM.DBDataSource = f.DataSources.DBDataSources.Item("@ME_SUP")
                                docVal = ds.GetValue("DocEntry", 0).ToString().Trim()
                            Catch
                                docVal = ""
                            End Try
                        End If

                        If Not String.IsNullOrWhiteSpace(docVal) Then
                            eventInfo.LayoutKey = docVal
                        End If
                    Catch ex As Exception
                        ' ignore - cannot supply layout key
                    End Try
                End If
            End If
        Catch ex As Exception
            Try
                objMain.objApplication.StatusBar.SetText("LayoutKey handler error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Catch
            End Try
        End Try

    End Sub


    ' Enabling Print Layout Icon
    Public Sub CreateLayoutFolder_ResgistrationCharges()
        Try
            ' Use the connected company object from MainCls (objMain.objCompany). If not available, abort with a clear message.
            Dim companyServiceSource As SAPbobsCOM.Company = Nothing
            Try
                If objMain IsNot Nothing AndAlso objMain.objCompany IsNot Nothing Then
                    companyServiceSource = objMain.objCompany
                End If
            Catch
                companyServiceSource = Nothing
            End Try

            If companyServiceSource Is Nothing Then
                objMain.objApplication.StatusBar.SetText("CreateLayoutFolder_ResgistrationCharges aborted: not connected to a company.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return
            End If

            Dim oReportTypeService As SAPbobsCOM.ReportTypesService
            oReportTypeService = companyServiceSource.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService)

            Dim oList = oReportTypeService.GetReportTypeList()

            For i As Integer = 0 To oList.Count - 1
                If oList.Item(i).AddonFormType = "ME_SUP_Form" Then
                    Exit Sub   ' Already exists
                End If
            Next

            Dim oReportType As SAPbobsCOM.ReportType
            oReportType = oReportTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportType)

            oReportType.TypeName = "ME_SUP_Form Print"
            oReportType.AddonName = "Payroll Addon"
            oReportType.AddonFormType = "ME_SUP_Form"
            oReportType.MenuID = "ME_SUP_Form_MENU"

            oReportTypeService.AddReportType(oReportType)

        Catch ex As Exception
            Try
                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Catch
            End Try
        End Try
    End Sub

    '--Madhu Disabled method
    Private Sub DisableColumns()

        Try
            objForm.Freeze(True)

            Me.objMatrix1.Columns.Item("VATCODE").Visible = False
            Me.objMatrix1.Columns.Item("VATAMT").Visible = False



        Catch ex As Exception
            objForm.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        Finally
            objForm.Freeze(False)
        End Try

    End Sub
End Class
