Imports System.Runtime.Remoting.Channels

Public Class clsOutlets

#Region "    Declaration   "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oEditText As SAPbouiCOM.EditText
    Dim oDt As SAPbouiCOM.DataTable
    Dim sSupplierCode, sSupplierName As String
    Dim objComboBox, objComboBox1 As SAPbouiCOM.ComboBox
    Dim OrsSTRQRY, rs, rs2, rsPO As SAPbobsCOM.Recordset
    Dim LostFocusFlag As Boolean = False

    Dim objutilities As Utilities
#End Region
    Sub CreateForm(ByVal FormUID As String, ByVal CardCode As String, ByVal CardName As String, ByVal BaseNo As String, ByVal RowId As Integer, ByVal AmountValue As Double, ByVal OutletValue As Integer, ByVal VatCodeValue As String, ByVal IfITLSExist As String)
        Try
            objMain.objUtilities.LoadForm("OutletForm.xml", "Me_Outlet", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("Me_Outlet", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            objMain.objApplication.StatusBar.SetText("Data Loading Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            sSupplierCode = CardCode

            If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = False Then
                Retail.MainCls.ohtLookUpForm.Add(objForm.UniqueID, FormUID)
            End If

            objutilities = New Utilities
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_OUTLET")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_OUTLET_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific

            objComboBox = objForm.Items.Item("12").Specific
            Dim GetSeriesQty As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'SBO_OUTLETUDO' And ""Locked"" = 'N'"
            objMain.objUtilities.ComboBoxLoadValuesSeries(objForm.Items.Item("12").Specific, GetSeriesQty)
            objForm.Items.Item("12").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.DataBrowser.BrowseBy = "9"


            objForm.Items.Item("Post").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("Post").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("Post").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("VATCODE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("VATCODE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("VATCODE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("VATCODE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            '--------------------------------------------------------------
            '             CASE 1 : RECORD EXISTS
            '--------------------------------------------------------------
            If IfITLSExist = "Y" Then
                objForm.Freeze(True)
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE

                objForm.Items.Item("BaseNo").Specific.Value = BaseNo
                objForm.Items.Item("LineNum").Specific.Value = RowId

                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                objForm.Items.Item("AUT").Specific.Value = AmountValue
                objForm.Items.Item("Outlet").Specific.Value = OutletValue
                objForm.Items.Item("VATCODE").Specific.Value = VatCodeValue
                objMatrix.AutoResizeColumns()
                objForm.Freeze(False)

            End If


            '--------------------------------------------------------------
            '             CASE 2 : NEW RECORD
            '--------------------------------------------------------------
            If IfITLSExist = "N" Then

                sSupplierName = CardName
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE

                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "SBO_OUTLETUDO"))
                oDBs_Head.SetValue("U_CardCode", oDBs_Head.Offset, sSupplierCode)
                oDBs_Head.SetValue("U_CardName", oDBs_Head.Offset, sSupplierName)
                oDBs_Head.SetValue("U_LineNum", oDBs_Head.Offset, RowId)
                oDBs_Head.SetValue("U_BaseNo", oDBs_Head.Offset, BaseNo)

                objForm.Items.Item("AUT").Specific.value = AmountValue
                objForm.Items.Item("Outlet").Specific.value = OutletValue
                objForm.Items.Item("VATCODE").Specific.value = VatCodeValue

                Dim GetDflt As String = "Select ""DfltSeries"" From ONNM Where ""ObjectCode"" = 'SBO_OUTLETUDO'"
                Dim oRsGetDflt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetDflt.DoQuery(GetDflt)

                oDBs_Head.SetValue("U_SERIES", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)
                oDBs_Head.SetValue("Series", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)

                objMatrix.Clear()
                oDBs_Details.Clear()
                objMatrix.FlushToDataSource()

                Me.SetNewLine(objForm.UniqueID)

                'loading Vat code combox

                Try
                    objComboBox1 = objMatrix.Columns.Item("VATCODE").Cells.Item(1).Specific  ''objForm.Items.Item("14").Specific
                    ' Dim sQury As String = "SELECT ""Code"",""Name"" FROM OVTG WHERE IFNULL(""Account"",'')<>'' and ""Locked""='N'"
                    Dim sQury As String = "SELECT ""Code"",""Name"" FROM OVTG WHERE IFNULL(""Account"",'')<>'' AND ""Locked""='N' AND UPPER(TRIM(""Name"")) NOT LIKE '%OUTPUT%'"

                    objMain.objUtilities.ComboBoxLoadValues(objComboBox1, sQury)
                    objComboBox1.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                Catch ex As Exception


                End Try


                Dim stQry As String = "Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join " & " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId"""
                rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rs.DoQuery(stQry)

                If rs.RecordCount > 0 Then

                    For i As Integer = 1 To rs.RecordCount

                        objMatrix.Columns.Item("Outlet").Cells.Item(i).Specific.Value =
            rs.Fields.Item("BPLName").Value

                        ' ✅ Set VAT Code in Matrix Column for Item Listing screen
                        Try
                            Dim oVatCombo As SAPbouiCOM.ComboBox = CType(objMatrix.Columns.Item("VATCODE").Cells.Item(i).Specific, SAPbouiCOM.ComboBox)

                            If VatCodeValue <> "" Then
                                oVatCombo.Select(VatCodeValue, SAPbouiCOM.BoSearchKey.psk_ByValue)
                            End If
                        Catch
                        End Try

                        Dim checkBox As SAPbouiCOM.CheckBox =
            CType(objMatrix.Columns.Item("Amount").Cells.Item(i).Specific, SAPbouiCOM.CheckBox)
                        checkBox.Checked = False

                        objMatrix.Columns.Item("AmntSpt").Cells.Item(i).Specific.Value = ""

                        rs.MoveNext()
                        Me.SetNewLine(objForm.UniqueID)
                    Next
                    objMatrix.AutoResizeColumns()
                    objForm.Freeze(False)

                End If

            End If
            objMatrix.Columns.Item("AmntSpt").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix.Columns.Item("VATAMT").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

            objForm.Items.Item("BaseNo").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("BaseNo").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("BaseNo").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("BaseNo").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            Me.CellMasking(objForm.UniqueID)
            objMain.objApplication.StatusBar.SetText("Data Loading Completed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

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
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_OUTLET")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_OUTLET_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific

            objMatrix.AddRow()
            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
            oDBs_Details.SetValue("U_Outlet", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Amount", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_JE", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Payment", oDBs_Details.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
            '  objMatrix.AutoResizeColumns()
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                        objForm = objMain.objApplication.Forms.Item(FormUID)
                        'If Me.Validation(objForm) = False Then
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If

                        ''NEW validation (Outlet vs Matrix checkbox)
                        'If Me.ValidateOutletVsAmount(objForm) = False Then
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If

                        'objForm.Items.Item("Post").Enabled = True


                    End If



                    'madhu checkbox validation uncheck amount 0
                    If pVal.ItemUID = "Mtx" And pVal.ColUID = "Amount" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                        Dim oMatrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx").Specific, SAPbouiCOM.Matrix)

                        Dim oCheck As SAPbouiCOM.CheckBox = CType(oMatrix.Columns.Item("Amount").Cells.Item(pVal.Row).Specific, SAPbouiCOM.CheckBox)

                        ' 🔴 UNCHECK → Amount Split = 0
                        If oCheck.Checked = False Then
                            oMatrix.Columns.Item("AmntSpt").Cells.Item(pVal.Row).Specific.Value = 0
                        End If
                        Me.calculateSplitAmountBasedOnCheckBoxes(objForm)
                    End If




                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_OUTLET")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_OUTLET_C0")
                    objMatrix = objForm.Items.Item("Mtx").Specific
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim chkIfExist As String = "Select ""DocNum"" From ""@SBO_OUTLET"" Where ""DocEntry""=(Select MAX(""DocEntry"") From ""@SBO_OUTLET"")"
                        Dim oRschkIfExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRschkIfExist.DoQuery(chkIfExist)
                        If oRschkIfExist.RecordCount > 0 Then
                            objForm.Freeze(True)
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            objForm.Items.Item("9").Specific.value = oRschkIfExist.Fields.Item("DocNum").Value
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            objForm.Freeze(False)
                        End If
                        objMain.objApplication.ActivateMenuItem("1304")
                    End If
                    If pVal.ItemUID = "Post" And pVal.BeforeAction = False And objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Me.PostWithVATJE()
                    End If
                    If pVal.ItemUID = "INPAY" And pVal.BeforeAction = False And objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        objMain.objApplication.ActivateMenuItem("2817")
                        Dim objBaseAppTemForm_IN As SAPbouiCOM.Form
                        Dim oMatrix As SAPbouiCOM.Matrix
                        objBaseAppTemForm_IN = objMain.objApplication.Forms.GetForm("170", objMain.objApplication.Forms.ActiveForm.TypeCount)
                        objBaseAppTemForm_IN.Items.Item("57").Specific.Selected = True
                        Dim supplier As String = oDBs_Head.GetValue("U_CardCode", 0).Trim
                        objBaseAppTemForm_IN.Items.Item("5").Specific.Value = supplier
                        objBaseAppTemForm_IN.Items.Item("OutNum").Specific.Value = objForm.Items.Item("9").Specific.value

                        Dim Qry As String = " Select T0.""BPLId""From USR6 T0 Inner Join " &
                                                                               " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "'Order By T0.""BPLId""  "
                        'Dim Qry As String = "SELECT T0.""BPLId"", T0.""BPLName"" FROM OBPL T0 where  T0.""BPLName""='" & objMatrix.Columns.Item("Outlet").Cells.Item(pVal.Row).Specific.Value & "'"
                        rs2 = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        rs2.DoQuery(Qry)
                        Dim BranchId As String = rs2.Fields.Item(0).Value.ToString().Trim()
                        If BranchId <> "" Then
                            Dim ocb As SAPbouiCOM.ComboBox = objBaseAppTemForm_IN.Items.Item("1320002037").Specific
                            ocb.SelectExclusive(BranchId, SAPbouiCOM.BoSearchKey.psk_ByValue)
                        End If
                        oMatrix = objBaseAppTemForm_IN.Items.Item("71").Specific
                        objBaseAppTemForm_IN.Freeze(True)
                        objBaseAppTemForm_IN.Items.Item("234000001").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        objBaseAppTemForm_IN.Freeze(False)
                        Dim Query As String = "SELECT ""U_REG"" FROM ""@ME_ACFG"" "
                        Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        rs.DoQuery(Query)
                        Dim DrAccount As String = rs.Fields.Item(0).Value
                        Dim AppTemForm_PAY As SAPbouiCOM.Form
                        AppTemForm_PAY = objMain.objApplication.Forms.GetFormByTypeAndCount(146, 1)
                        AppTemForm_PAY.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                        AppTemForm_PAY.Items.Item("32").Specific.Value = DrAccount
                        Dim AmountQry As String = "SELECT ""U_Amount""  FROM ""@SBO_LIST""  T0 inner join  ""@SBO_LIST_C0""  T1 on T1.""DocEntry"" = T0.""DocEntry"" where ""U_CardCode""='" & oDBs_Head.GetValue("U_CardCode", 0) & "'and ""LineId""='" & oDBs_Head.GetValue("U_LineNum", 0) & "'"
                        Dim rs1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        rs1.DoQuery(AmountQry)
                        Dim Amount As Double = CDbl(rs1.Fields.Item(0).Value)
                        AppTemForm_PAY.Items.Item("38").Specific.Value = Amount
                        AppTemForm_PAY.Items.Item("12").Specific.Value = Amount

                        '   Me.Incomingpayments()
                        'Dim AppTemForm_IN As SAPbouiCOM.Form
                        'AppTemForm_IN = objMain.objApplication.Forms.GetFormByTypeAndCount(170, 1)
                        'If pVal.ItemUID = "1" Then
                        '    Dim Query1 As String = "Update ""@SBO_OUTLET_C0"" Set ""U_Payment""=(Select Max(""DocEntry"") FROM ORCT) Where ""DocNum""='" & AppTemForm_IN.Items.Item("3").Specific.Value & "' "
                        '    rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '    rs.DoQuery(Query1)
                        ' End If
                    End If
                    If pVal.BeforeAction = False And objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        objForm.Freeze(True)
                        'Not required to enable buttons here as they are auto-managed, but if you want to manage them manually, you can uncomment the line below
                        'Me.Enablingbuttons(objForm.UniqueID)
                        objForm.Freeze(False)
                    End If

                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    If pVal.ItemUID = "PType" And pVal.Before_Action = False Then
                        Dim objCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("PType").Specific
                        If objForm.Items.Item("PType").Specific.Selected.Value <> "" Then
                            Dim series As Integer = objForm.Items.Item("PType").Specific.Selected.Value
                            If objCombo.Selected.Value = "1" Then
                                objForm.Items.Item("INPAY").Enabled = True
                                ' objForm.Items.Item("Post").Enabled = False

                            Else
                                ' objForm.Items.Item("Post").Enabled = True
                                objForm.Items.Item("INPAY").Enabled = False


                            End If
                        End If
                    End If

                    'calculate amount spliting basedon vat code
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix = objForm.Items.Item("Mtx").Specific
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_OUTLET_C0")

                    If pVal.ItemUID = "Mtx" And pVal.ColUID = "VATCODE" And pVal.BeforeAction = False Then
                        Try
                            If objMatrix.Columns.Item("JE").Cells.Item(pVal.Row).Specific.Value.ToString() = "" Then
                                Dim vatCode As String = objMatrix.Columns.Item("VATCODE").Cells.Item(pVal.Row).Specific.Value.ToString()
                                Dim Amount As String = objMatrix.Columns.Item("AmntSpt").Cells.Item(pVal.Row).Specific.Value.ToString()
                                Dim qStrVatPer As String = "SELECT ""Rate"" FROM OVTG WHERE ""Code""='" + vatCode + "'"
                                Dim rSetVatPer As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                rSetVatPer.DoQuery(qStrVatPer)
                                If rSetVatPer.RecordCount > 0 Then
                                    Dim Rate As Double = CDbl(rSetVatPer.Fields.Item("Rate").Value)
                                    Dim VatAmount As Double = (Rate / 100) * Amount
                                    objMatrix.FlushToDataSource()
                                    'oDBs_Details.SetValue("U_VATAMT", oDBs_Details.Offset, VatAmount.ToString())
                                    'objMatrix.LoadFromDataSource()
                                    objMatrix.Columns.Item("VATAMT").Cells.Item(pVal.Row).Specific.Value = VatAmount.ToString()
                                    objMatrix.FlushToDataSource()
                                End If
                            End If
                        Catch ex As Exception

                        End Try
                    End If

                    '--madhu In item listing outlet posting, when the amount splitting is changed manually, the VAT is not refreshing accordingly
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                    If pVal.ItemUID = "Mtx" And pVal.ColUID = "AmntSpt" And pVal.BeforeAction = False Then
                        Try
                            objForm = objMain.objApplication.Forms.Item(FormUID)
                            objForm.Freeze(True)

                            Dim oMatrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx").Specific, SAPbouiCOM.Matrix)

                            Dim vatCode As String = oMatrix.Columns.Item("VATCODE").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()

                            Dim AmountSplitted As String = oMatrix.Columns.Item("AmntSpt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()

                            If vatCode <> "" And AmountSplitted <> "" Then

                                Dim amount As Double = CDbl(AmountSplitted)

                                Dim qStrVatPer As String = "SELECT ""Rate"" FROM OVTG WHERE ""Code""='" & vatCode & "'"

                                Dim rSetVatPer As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                rSetVatPer.DoQuery(qStrVatPer)

                                If rSetVatPer.RecordCount > 0 Then

                                    Dim rate As Double = CDbl(rSetVatPer.Fields.Item("Rate").Value)

                                    Dim vatAmount As Double = (rate / 100) * amount

                                    oMatrix.Columns.Item("VATAMT").Cells.Item(pVal.Row).Specific.Value = Math.Round(vatAmount, 3).ToString()


                                End If
                            End If

                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        Finally
                            objForm.Freeze(False)
                        End Try
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

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)

        Try
            objForm = objMain.objApplication.Forms.GetForm("Me_Outlet", objMain.objApplication.Forms.ActiveForm.TypeCount)

            Select Case BusinessObjectInfo.EventType


                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                    If BusinessObjectInfo.BeforeAction = False AndAlso BusinessObjectInfo.ActionSuccess = True Then

                        Try
                            oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_OUTLET_C0")

                            Dim jeFound As Boolean = False

                            For i As Integer = 0 To oDBs_Details.Size - 1

                                Dim jeValue As String = ""

                                Try
                                    jeValue = oDBs_Details.GetValue("U_JE", i).Trim()
                                Catch
                                End Try

                                If jeValue <> "" Then
                                    jeFound = True
                                    Exit For
                                End If

                            Next

                            Dim oMatrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx").Specific, SAPbouiCOM.Matrix)

                            ' JE Posted → Disable matrix
                            For i As Integer = 0 To oMatrix.Columns.Count - 1
                                Try
                                    oMatrix.Columns.Item(i).Editable = False
                                Catch
                                End Try
                            Next
                            If jeFound = True Then


                                ' ✅ Disable Post Button (Same Logic as Matrix)
                                Try
                                    ' objForm.Items.Item("Post").Enabled = False
                                    objForm.Items.Item("Post").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                Catch
                                End Try

                            Else


                                ' ✅ Enable Post Button
                                Try
                                    objForm.Items.Item("Post").Enabled = True
                                Catch
                                End Try

                            End If

                        Catch
                        End Try

                    End If




            End Select

        Catch ex As Exception
            Try : objForm.Freeze(False) : Catch : End Try
            objMain.objApplication.StatusBar.SetText(
            ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Function Validation(ByVal oForm As SAPbouiCOM.Form) As Boolean

        Try
            Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("Mtx").Specific
            Dim TotalSplit As Double = 0.0
            Dim AUT_Amount As Double = 0.0

            Double.TryParse(oForm.Items.Item("AUT").Specific.Value.ToString().Trim(), AUT_Amount)

            For r As Integer = 1 To oMatrix.VisualRowCount
                Dim cellVal As String = oMatrix.Columns.Item("AmntSpt").Cells.Item(r).Specific.Value
                Dim v As Double = 0.0
                Double.TryParse(cellVal.ToString().Trim(), v)
                TotalSplit += v
            Next

            TotalSplit = Math.Round(TotalSplit, 2)

            '=== VALIDATIONS ===
            If TotalSplit <> AUT_Amount Then
                Dim msg As String = $"Split amount ({TotalSplit}) must be equal to Actual Amount ({AUT_Amount})!"
                objMain.objApplication.StatusBar.SetText(msg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Return True

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Sub PostJEPostWithOutVATJE()
        oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_OUTLET")
        oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_OUTLET_C0")
        objMatrix = objForm.Items.Item("Mtx").Specific
        For i As Integer = 1 To objMatrix.VisualRowCount
            Dim supplier As String = oDBs_Head.GetValue("U_CardCode", 0).Trim
            'Dim AmountQry As String = "SELECT ""U_Amount""  FROM ""@SBO_LIST""  T0 inner join  ""@SBO_LIST_C0""  T1 on T1.""DocEntry"" = T0.""DocEntry"" where ""U_CardCode""='" & oDBs_Head.GetValue("U_CardCode", 0) & "'and ""LineId""='" & oDBs_Head.GetValue("U_LineNum", 0) & "'"
            'Dim AmountQry As String = "SELECT ""U_AmntSpt""  FROM ""@SBO_OUTLET""  T0 inner join  ""@SBO_OUTLET_C0""  T1 on T1.""DocEntry"" = T0.""DocEntry"" where ""U_AmntSpt"">0 and  ""U_CardCode""='" & oDBs_Head.GetValue("U_CardCode", 0) & "'and ""U_LineNum""='" & oDBs_Head.GetValue("U_LineNum", 0) & "'"
            'Dim rs1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'rs1.DoQuery(AmountQry)
            'Dim PayFlag As SAPbouiCOM.CheckBox = objMatrix.Columns.Item("Amount").Cells.Item(i).Specific
            Dim PayFlag As SAPbouiCOM.CheckBox = objMatrix.Columns.Item("Amount").Cells.Item(i).Specific

            ' For g As Integer = 1 To rs1.RecordCount
            If objMatrix.Columns.Item("JE").Cells.Item(i).Specific.Value.ToString().Trim = "" And objMatrix.Columns.Item("AmntSpt").Cells.Item(i).Specific.Value <> 0 And objMatrix.Columns.Item("Outlet").Cells.Item(i).Specific.Value.ToString().Trim <> "" And PayFlag.Checked = True Then
                Dim Query As String = "SELECT ""U_REG"" FROM ""@ME_ACFG"" "
                Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rs.DoQuery(Query)
                Dim DrAccount As String = rs.Fields.Item(0).Value
                Dim Qry As String = "SELECT T0.""BPLId"", T0.""BPLName"" FROM OBPL T0 where  T0.""BPLName""='" & objMatrix.Columns.Item("Outlet").Cells.Item(i).Specific.Value & "'"
                rs2 = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rs2.DoQuery(Qry)
                Dim Amount As Double = objMatrix.Columns.Item("AmntSpt").Cells.Item(i).Specific.Value
                'Dim Remarks As String = objMatrix.Columns.Item("Remarks").Cells.Item(i).Specific.Value.ToString().Trim
                If objMatrix.Columns.Item("Outlet").Cells.Item(i).Specific.Value = "" Then
                    objMain.objApplication.StatusBar.SetText("Please select the branch.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
                Dim oJE As SAPbobsCOM.JournalEntries
                oJE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                Dim sDate As String = DateTime.Now.ToString("yyyyMMdd")
                sDate = sDate.Insert("4", "-")
                sDate = sDate.Insert("7", "-")

                oJE.ReferenceDate = sDate
                oJE.DueDate = sDate
                oJE.TaxDate = sDate
                oJE.UserFields.Fields.Item("U_Remarks").Value = "Outlet"
                oJE.Reference = "Items Listing Outlet"
                oJE.Memo = "Listing amount"

                Dim test As String = rs2.Fields.Item("BPLId").Value
                oJE.Lines.BPLID = test

                ' oJE.Lines.BPLID = 
                oJE.Lines.ShortName = supplier
                oJE.Lines.Credit = Amount
                oJE.Lines.Add()

                oJE.Lines.AccountCode = DrAccount
                oJE.Lines.Debit = Amount
                oJE.Lines.BPLID = rs2.Fields.Item("BPLId").Value
                oJE.Lines.Add()

                If oJE.Add() = 0 Then
                    Dim GetJEDocEntry As String = "Select Max(""TransId"") From OJDT Where ""TransType"" = '30'"
                    Dim oRsGetJEDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetJEDocEntry.DoQuery(GetJEDocEntry)

                    objMain.objApplication.StatusBar.SetText("Registration JE Posted Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Dim Query1 As String = "Update ""@SBO_OUTLET_C0"" Set ""U_JE""=(Select Max(""TransId"") FROM OJDT) Where ""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0) & "' and ""LineId""='" & i & "'"
                    rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    rs.DoQuery(Query1)
                    ' ===========================
                    ' 🔥 SET FORM TO VIEW MODE
                    ' ===========================
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

                    ' Disable Matrix
                    Dim oMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx").Specific

                    oMatrix.AutoResizeColumns()
                    objForm.Items.Item("Mtx").Enabled = False

                    '  objForm.Items.Item("Post").Enabled = False
                    objForm.Items.Item("Post").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                Else
                    objMain.objApplication.StatusBar.SetText("Registration JE Posting failed" + objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
                'objMain.objApplication.ActivateMenuItem("1304")
            End If
        Next


        Try
            objMain.objApplication.ActivateMenuItem("1304")
        Catch ex As Exception

        End Try
    End Sub

    Sub PostWithVATJE()

        Try

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_OUTLET")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_OUTLET_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific

            For i As Integer = 1 To objMatrix.VisualRowCount

                Dim supplier As String = oDBs_Head.GetValue("U_CardCode", 0).Trim
                Dim PayFlag As SAPbouiCOM.CheckBox = objMatrix.Columns.Item("Amount").Cells.Item(i).Specific

                If objMatrix.Columns.Item("JE").Cells.Item(i).Specific.Value.ToString.Trim = "" _
            And CDbl(objMatrix.Columns.Item("AmntSpt").Cells.Item(i).Specific.Value) <> 0 _
            And objMatrix.Columns.Item("Outlet").Cells.Item(i).Specific.Value.ToString.Trim <> "" _
            And PayFlag.Checked = True Then

                    Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    rs.DoQuery("SELECT ""U_REG"" FROM ""@ME_ACFG""")

                    If rs.RecordCount = 0 Then
                        objMain.objApplication.StatusBar.SetText("Expense account not configured", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If

                    Dim DrAccount As String = rs.Fields.Item(0).Value.ToString.Trim

                    If DrAccount = "" Then
                        objMain.objApplication.StatusBar.SetText("Expense account is blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If

                    Dim rs2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    rs2.DoQuery("SELECT ""BPLId"" FROM OBPL WHERE ""BPLName""='" & objMatrix.Columns.Item("Outlet").Cells.Item(i).Specific.Value & "'")

                    If rs2.RecordCount = 0 Then
                        objMain.objApplication.StatusBar.SetText("Branch not found", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If

                    Dim BplId As Integer = rs2.Fields.Item(0).Value

                    Dim Amount As Double = CDbl(objMatrix.Columns.Item("AmntSpt").Cells.Item(i).Specific.Value)

                    Dim VatAmount As Double = 0
                    If objMatrix.Columns.Item("VATAMT").Cells.Item(i).Specific.Value.ToString.Trim <> "" Then
                        VatAmount = CDbl(objMatrix.Columns.Item("VATAMT").Cells.Item(i).Specific.Value)
                    End If

                    Dim VatCode As String = objMatrix.Columns.Item("VATCODE").Cells.Item(i).Specific.Value.ToString.Trim
                    Dim ValidVatCode As String = ""

                    If VatCode <> "" Then
                        Dim rsVat As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        rsVat.DoQuery("SELECT ""Code"" FROM OVTG WHERE ""Code""='" & VatCode & "'")

                        If rsVat.RecordCount > 0 Then
                            ValidVatCode = rsVat.Fields.Item(0).Value.ToString.Trim
                        End If
                    End If

                    Dim oJE As SAPbobsCOM.JournalEntries
                    oJE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

                    Dim sDate As Date = Date.Now

                    'oJE.ReferenceDate = sDate
                    oJE.DueDate = sDate
                    oJE.TaxDate = sDate
                    oJE.Reference = "Items Listing Outlet"

                    oJE.Reference2 = oDBs_Head.GetValue("DocNum", 0).Trim
                    Dim DocumentDate As String = oDBs_Head.GetValue("U_DocDate", 0).Trim()
                    oJE.ReferenceDate = Date.ParseExact(DocumentDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture)

                    oJE.Memo = "Listing amount"
                    oJE.UserFields.Fields.Item("U_Remarks").Value = "Outlet"
                    oJE.AutoVAT = SAPbobsCOM.BoYesNoEnum.tYES

                    ' ---------------- FIRST LINE (DEBIT - SUPPLIER) ----------------
                    oJE.Lines.BPLID = BplId
                    oJE.Lines.ShortName = supplier
                    oJE.Lines.Debit = Amount + VatAmount
                    oJE.Lines.Add()

                    ' ---------------- SECOND LINE (CREDIT - EXPENSE) ----------------
                    oJE.Lines.AccountCode = DrAccount
                    oJE.Lines.Credit = Amount
                    oJE.Lines.BPLID = BplId

                    If ValidVatCode <> "" Then
                        oJE.Lines.TaxGroup = ValidVatCode
                    End If

                    oJE.Lines.Add()

                    If oJE.Add() = 0 Then

                        Dim rsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        rsUpdate.DoQuery("Update ""@SBO_OUTLET_C0"" Set ""U_JE""=(Select Max(""TransId"") FROM OJDT) Where ""DocEntry""='" _
                                     & oDBs_Head.GetValue("DocEntry", 0) & "' and ""LineId""='" & i & "'")

                        objMain.objApplication.StatusBar.SetText("Registration JE Posted Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

                        Dim oMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx").Specific
                        oMatrix.AutoResizeColumns()
                        objForm.Items.Item("Mtx").Enabled = False
                        objForm.Items.Item("Post").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                    Else
                        objMain.objApplication.StatusBar.SetText("Registration JE Posting failed " & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If

                End If

            Next

            objMain.objApplication.ActivateMenuItem("1304")

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub


    Sub Incomingpayments()
        objMain.objApplication.ActivateMenuItem("2817")
        Dim AppTemForm_IN As SAPbouiCOM.Form
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim objForm As SAPbouiCOM.Form
        AppTemForm_IN = objMain.objApplication.Forms.GetFormByTypeAndCount(170, 1)
        AppTemForm_IN.Items.Item("57").Specific.Selected = True
        Dim supplier As String = oDBs_Head.GetValue("U_CardCode", 0).Trim
        AppTemForm_IN.Items.Item("5").Specific.Value = supplier
        Dim Qry As String = " Select T0.""BPLId""From USR6 T0 Inner Join " &
                                                                 " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "'Order By T0.""BPLId""  "
        'Dim Qry As String = "SELECT T0.""BPLId"", T0.""BPLName"" FROM OBPL T0 where  T0.""BPLName""='" & objMatrix.Columns.Item("Outlet").Cells.Item(pVal.Row).Specific.Value & "'"
        rs2 = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs2.DoQuery(Qry)
        Dim BranchId As String = rs2.Fields.Item(0).Value.ToString().Trim()
        If BranchId <> "" Then
            Dim ocb As SAPbouiCOM.ComboBox = AppTemForm_IN.Items.Item("1320002037").Specific
            ocb.SelectExclusive(BranchId, SAPbouiCOM.BoSearchKey.psk_ByValue)
        End If


        oMatrix = AppTemForm_IN.Items.Item("71").Specific
        AppTemForm_IN.Freeze(True)
        AppTemForm_IN.Items.Item("234000001").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        AppTemForm_IN.Freeze(False)
        Dim Query As String = "SELECT ""U_REG"" FROM ""@ME_ACFG"" "
        Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs.DoQuery(Query)
        Dim DrAccount As String = rs.Fields.Item(0).Value
        Dim AppTemForm_PAY As SAPbouiCOM.Form
        AppTemForm_PAY = objMain.objApplication.Forms.GetFormByTypeAndCount(146, 1)
        AppTemForm_PAY.Items.Item("32").Specific.Value = DrAccount
        Dim AmountQry As String = "SELECT ""U_Amount""  FROM ""@SBO_LIST""  T0 inner join  ""@SBO_LIST_C0""  T1 on T1.""DocEntry"" = T0.""DocEntry"" where ""U_CardCode""='" & oDBs_Head.GetValue("U_CardCode", 0) & "'and ""LineId""='" & oDBs_Head.GetValue("U_LineNum", 0) & "'"
        Dim rs1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs1.DoQuery(AmountQry)
        Dim Amount As Double = CDbl(rs1.Fields.Item(0).Value)
        AppTemForm_PAY.Items.Item("38").Specific.Value = Amount
        AppTemForm_PAY.Items.Item("12").Specific.Value = Amount
        '   Dim testb As String = AppTemForm_IN.Items.Item("3").Specific.Value

    End Sub
    Sub Enablingbuttons(ByVal FormUID As String)
        Dim PostingType As String = objForm.Items.Item("PType").Specific.Value
        If PostingType = "2" Then
            'objForm.Items.Item("Post").Enabled = True
            objForm.Items.Item("INPAY").Enabled = False
        Else
            objForm.Items.Item("INPAY").Enabled = True
            ' objForm.Items.Item("Post").Enabled = False

        End If
    End Sub

    Private Sub calculateSplitAmountBasedOnCheckBoxes(ByVal objForm As SAPbouiCOM.Form)

        Try

            objForm.Freeze(True)

            Dim objMatrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx").Specific, SAPbouiCOM.Matrix)
            Dim HeaderAmount As Double = 0
            Double.TryParse(CType(objForm.Items.Item("AUT").Specific, SAPbouiCOM.EditText).Value, HeaderAmount)
            Dim CheckedCount As Integer = 0
            ' Count checked rows
            For i As Integer = 1 To objMatrix.RowCount

                Dim chk As SAPbouiCOM.CheckBox = CType(objMatrix.Columns.Item("Amount").Cells.Item(i).Specific, SAPbouiCOM.CheckBox)
                If chk.Checked Then
                    CheckedCount += 1
                End If

            Next

            ' Validation
            If CheckedCount = 0 OrElse HeaderAmount = 0 Then

                For i As Integer = 1 To objMatrix.RowCount
                    objMatrix.Columns.Item("AmntSpt").Cells.Item(i).Specific.Value = ""
                    objMatrix.Columns.Item("VATAMT").Cells.Item(i).Specific.Value = ""

                Next
                objMain.objApplication.StatusBar.SetText("No rows selected or header amount is zero.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Exit Sub
            End If

            ' Processing Message
            objMain.objApplication.StatusBar.SetText("Please wait... Processing split amount calculation.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            Dim SplitAmount As Double = Math.Round(HeaderAmount / CheckedCount, 2)

            For i As Integer = 1 To objMatrix.RowCount
                Dim chk As SAPbouiCOM.CheckBox = CType(objMatrix.Columns.Item("Amount").Cells.Item(i).Specific, SAPbouiCOM.CheckBox)
                If chk.Checked Then

                    objMatrix.Columns.Item("AmntSpt").Cells.Item(i).Specific.Value = SplitAmount.ToString()

                    ' VAT Code
                    Dim VatCode As String = objMatrix.Columns.Item("VATCODE").Cells.Item(i).Specific.Value.ToString().Trim()

                    If VatCode <> "" Then

                        Dim Rate As Double = GetVatRate(VatCode)
                        Dim VatAmount As Double = (Rate / 100) * SplitAmount
                        objMatrix.Columns.Item("VATAMT").Cells.Item(i).Specific.Value = Math.Round(VatAmount, 2).ToString()
                    Else
                        objMatrix.Columns.Item("VATAMT").Cells.Item(i).Specific.Value = ""
                    End If
                Else

                    objMatrix.Columns.Item("AmntSpt").Cells.Item(i).Specific.Value = ""
                    objMatrix.Columns.Item("VATAMT").Cells.Item(i).Specific.Value = ""

                End If

            Next

            objMain.objApplication.StatusBar.SetText("Split Amount and VAT calculated successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            objForm.Freeze(False)
        End Try
    End Sub
    Private Function GetVatRate(ByVal VatCode As String) As Double

        Dim oRec As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        oRec.DoQuery("SELECT ""Rate"" FROM OVTG WHERE ""Code""='" + VatCode + "'")

        If Not oRec.EoF Then
            Return CDbl(oRec.Fields.Item("Rate").Value)
        End If

        Return 0

    End Function


End Class
