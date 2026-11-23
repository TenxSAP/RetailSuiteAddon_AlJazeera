Imports System.ComponentModel
Public Class clsMIGenerationWizard

#Region "        Declaration        "
    Dim objForm, objMIForm As SAPbouiCOM.Form
    Dim oGrid As SAPbouiCOM.Grid
    Dim oDt As SAPbouiCOM.DataTable
    Dim oEditText As SAPbouiCOM.EditText
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim DocFlag As String = ""
    Dim GetInTransitWhs As String
    Dim oRsGetInTransitWhs As SAPbobsCOM.Recordset
    Private WithEvents BackgroundWorker_GenerateDocumets As New System.ComponentModel.BackgroundWorker
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("MIGenerationWizard.xml", "ME_MIGNWZRD_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_MIGNWZRD_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.DataSources.UserDataSources.Add("U_FROMDT", SAPbouiCOM.BoDataType.dt_DATE)
            objForm.DataSources.UserDataSources.Add("U_TODT", SAPbouiCOM.BoDataType.dt_DATE)
            objForm.DataSources.UserDataSources.Add("U_FRBRCH", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            objForm.DataSources.UserDataSources.Add("U_TOBRCH", SAPbouiCOM.BoDataType.dt_LONG_TEXT)

            oEditText = objForm.Items.Item("22").Specific
            oEditText.DataBind.SetBound(True, "", "U_FROMDT")
            objForm.Items.Item("22").AffectsFormMode = False

            oEditText = objForm.Items.Item("4").Specific
            oEditText.DataBind.SetBound(True, "", "U_TODT")
            objForm.Items.Item("4").AffectsFormMode = False

            objComboBox = objForm.Items.Item("15").Specific
            objComboBox.DataBind.SetBound(True, "", "U_FRBRCH")
            objForm.Items.Item("15").AffectsFormMode = False
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, "Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""  ")
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            'If objComboBox.ValidValues.Count <> 0 Then
            '    objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
            'End If
            objForm.Items.Item("15").DisplayDesc = True

            objComboBox = objForm.Items.Item("1000001").Specific
            objComboBox.DataBind.SetBound(True, "", "U_TOBRCH")
            objForm.Items.Item("1000001").AffectsFormMode = False
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, "Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""  ")
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            If objComboBox.ValidValues.Count <> 0 Then
                objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
            End If
            objForm.Items.Item("1000001").DisplayDesc = True

            oDt = objForm.DataSources.DataTables.Add("dt1")
            oDt = objForm.DataSources.DataTables.Item("dt1")

            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ME_MIGNWZRD" And pVal.BeforeAction = False Then
                Me.CreateForm()
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oGrid = objForm.Items.Item("10").Specific
                    If pVal.ItemUID = "10" And pVal.ColUID = "MINo" And pVal.BeforeAction = True Then
                        BubbleEvent = False
                        objMain.objApplication.ActivateMenuItem("ME_MI")
                        objMIForm = objMain.objApplication.Forms.GetForm("ME_MI_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                        objMIForm.Freeze(True)
                        objMIForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objMIForm.Items.Item("13").Specific.value = oGrid.DataTable.GetValue("MINo", pVal.Row)
                        objMIForm.Items.Item("5").Specific.value = oGrid.DataTable.GetValue("TransferNumber", pVal.Row)
                        objMIForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        objMIForm.Freeze(False)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oGrid = objForm.Items.Item("10").Specific
                    If pVal.ItemUID = "10" And pVal.BeforeAction = False Then
                        Try
                            oGrid.Columns.Item("RowsHeader").Click(pVal.Row, SAPbouiCOM.BoCellClickType.ct_Regular)
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "9" And pVal.BeforeAction = False Then
                        'If objForm.Items.Item("22").Specific.Value = "" Then
                        '    objMain.objApplication.StatusBar.SetText("Please Enter From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        'Else
                        If objForm.Items.Item("4").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            GetInTransitWhs = "SELECT *  FROM ""@ME_ACFG""  T0 WHERE T0.""U_INTRW"" ='N' "
                            oRsGetInTransitWhs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetInTransitWhs.DoQuery(GetInTransitWhs)
                            If oRsGetInTransitWhs.RecordCount > 0 Then
                                Me.LoadGridGI(objForm.UniqueID, "Y")
                            Else
                                Me.LoadGrid(objForm.UniqueID, "Y")
                            End If

                        End If
                    End If

                    If pVal.ItemUID = "13" And pVal.BeforeAction = False Then
                        Me.LoadGrid(objForm.UniqueID, "Y")
                    End If

                    If pVal.ItemUID = "14" And pVal.BeforeAction = False Then
                        Me.LoadGrid(objForm.UniqueID, "N")
                    End If

                    If pVal.ItemUID = "11" And pVal.BeforeAction = False Then
                        DocFlag = "Issue"
                        StartBackgroundForLoading()
                    End If

                    If pVal.ItemUID = "12" And pVal.BeforeAction = False Then
                        DocFlag = "Receipt"
                        StartBackgroundForLoading()
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub LoadGrid(ByVal FormUID As String, ByVal Value As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMain.objApplication.StatusBar.SetText("Loading initiated...please wait a while...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objForm.Freeze(True)

            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oGrid = objForm.Items.Item("10").Specific
            oGrid.DataTable = oDt

            Try
                oGrid.DataTable.Clear()
            Catch ex As Exception
            End Try

            Dim Query As String = " Select '" & Value & "' as ""Check"" , ""DocEntry"" As ""MIEntry"" ,  ""DocNum"" As ""MINo"" , ""U_MEDATE"" As ""MIDate"" , " &
                                  " ""U_MEFBPID"" ""FromBranchId""  ,  (Select T1.""BPLName"" From  OBPL T1 Where T1.""BPLId"" = ""U_MEFBPID"" ) As ""FromBranch"" , " &
                                  " ""U_MEITREN"" As ""TransferId"" , ""U_MEITRNO"" As ""TransferNumber"" , ""U_METRWHS"" As ""TRWhs"" , (Select ""U_MEEXACCT"" From ""OWHS"" Where " &
                                  " ""WhsCode"" = T0.""U_METWHS"" And IFNULL(""U_MEEXACCT"",'')<>'') As ""GIAccountCode"" , (Select ""AcctName"" From OACT Where ""AcctCode"" = (Select ""U_MEEXACCT"" From ""OWHS"" Where " &
                                  " ""WhsCode"" = T0.""U_METWHS"" And IFNULL(""U_MEEXACCT"",'')<>'')) As  ""GIAccountName""  , (Select ""U_MEENTACT""  From ""OWHS""  " &
                                  " Where ""WhsCode"" = T0.""U_METRWHS"" And IFNULL(""U_MEENTACT"",'')<>'' ) As ""GRAccountCode"" , (Select ""AcctName"" From OACT Where ""AcctCode"" = (Select ""U_MEENTACT""  From ""OWHS""  " &
                                  " Where ""WhsCode"" = T0.""U_METRWHS"" And IFNULL(""U_MEENTACT"",'')<>'')) As  ""GRAccountName""   , ""U_METWHS"" As ""ToWhs"" , " &
                                  " IFNULL(""U_MEGIEN"",'') As ""IssueEntry"", IFNULL(""U_MEGINO"",'') As ""GoodsIssueNo""  , " &
                                  " IFNULL(""U_MEGREN"",'') As ""ReceiptEntry"" ,  IFNULL(""U_MEGRNO"",'') As ""GoodsReceiptNo"" , " &
                                 "  Cast('' AS NVARCHAR(5000)) As ""Status""   From ""@ME_MI"" T0 Where (IFNULL(""U_MEGINO"",'')='' Or IFNULL(""U_MEGRNO"",'')='')  And  " &
                                 "  IFNULL(""U_MESTS"",'')='Open' And  T0.""U_MEDATE"" Between  '20210101' And '" & objForm.Items.Item("4").Specific.Value & "'  "
            If objForm.DataSources.UserDataSources.Item("U_FRBRCH").Value <> "" Then
                Query = Query & " And T0.""U_MEFBPID"" = '" & objForm.Items.Item("15").Specific.Value & "' "
            End If
            If objForm.DataSources.UserDataSources.Item("U_TOBRCH").Value <> "" Then
                Query = Query & " And T0.""U_METBPID"" = '" & objForm.Items.Item("1000001").Specific.Value & "' "
            End If
            oGrid.DataTable.ExecuteQuery(Query)

            oGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item("MIEntry").Visible = False
            oGrid.Columns.Item("FromBranchId").Visible = False
            oGrid.Columns.Item("TRWhs").Visible = True
            oGrid.Columns.Item("ToWhs").Visible = True
            oGrid.Columns.Item("MINo").Editable = False
            oGrid.Columns.Item("MIDate").Editable = False
            oGrid.Columns.Item("FromBranch").Editable = False
            oGrid.Columns.Item("TransferNumber").Editable = False
            oGrid.Columns.Item("Status").Editable = False
            oGrid.Columns.Item("GIAccountCode").Editable = False
            oGrid.Columns.Item("GRAccountCode").Editable = False
            oGrid.Columns.Item("GIAccountName").Editable = False
            oGrid.Columns.Item("GRAccountName").Editable = False
            oGrid.Columns.Item("GoodsIssueNo").Editable = False
            oGrid.Columns.Item("GoodsReceiptNo").Editable = False
            oGrid.Columns.Item("IssueEntry").Editable = False
            oGrid.Columns.Item("ReceiptEntry").Editable = False
            oGrid.Columns.Item("TransferId").Editable = False
            oGrid.Columns.Item("Status").Editable = False

            If oGrid.DataTable.GetValue("TransferNumber", 0) = "" Then
                objMain.objApplication.MessageBox("No records found")
                objForm.Freeze(False)
                Exit Sub
            End If

            Dim oEditTextCol As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("GIAccountCode")
            oEditTextCol.LinkedObjectType = 1

            Dim oEditTextCol1 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("GRAccountCode")
            oEditTextCol1.LinkedObjectType = 1

            Dim oEditTextCol2 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("IssueEntry")
            oEditTextCol2.LinkedObjectType = 60

            Dim oEditTextCol3 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("ReceiptEntry")
            oEditTextCol3.LinkedObjectType = 59

            Dim oEditTextCol4 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("MINo")
            oEditTextCol4.LinkedObjectType = 17

            Dim oEditTextCol5 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("TransferId")
            oEditTextCol5.LinkedObjectType = 67


            For x As Integer = 0 To oGrid.Rows.Count - 1
                oGrid.CommonSetting.SetRowFontStyle(x + 1, SAPbouiCOM.BoFontStyle.fs_Bold)
                oGrid.CommonSetting.SetRowFontColor(x + 1, ColorTranslator.ToOle(Color.Black))
            Next

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        Finally
            oGrid.AutoResizeColumns()
            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

        End Try
    End Sub

    Sub LoadGridGI(ByVal FormUID As String, ByVal Value As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMain.objApplication.StatusBar.SetText("Loading initiated...please wait a while...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objForm.Freeze(True)

            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oGrid = objForm.Items.Item("10").Specific
            oGrid.DataTable = oDt

            Try
                oGrid.DataTable.Clear()
            Catch ex As Exception
            End Try
            ''" ""U_MEITREN"" As ""TransferId"" , ""U_MEITRNO"" As ""TransferNumber"" , ""U_METRWHS"" As ""TRWhs"" , (Select ""U_MEEXACCT"" From ""OWHS"" Where " &

            Dim Query As String = " Select '" & Value & "' as ""Check"" , ""DocEntry"" As ""MIEntry"" ,  ""DocNum"" As ""MINo"" , ""U_MEDATE"" As ""MIDate"" , " &
                                  " ""U_MEFBPID"" ""FromBranchId""  ,  (Select T1.""BPLName"" From  OBPL T1 Where T1.""BPLId"" = ""U_MEFBPID"" ) As ""FromBranch"" , " &
                                  " ""U_MEITREN"" As ""TransferId"" , ""U_MEITRNO"" As ""TransferNumber"" ,  ""U_METWHS"" As ""TRWhs"" , (Select ""U_MEEXACCT"" From ""OWHS"" Where " &
                                  " ""WhsCode"" = T0.""U_METWHS"" And IFNULL(""U_MEEXACCT"",'')<>'') As ""GIAccountCode"" , (Select ""AcctName"" From OACT Where ""AcctCode"" = (Select ""U_MEEXACCT"" From ""OWHS"" Where " &
                                  " ""WhsCode"" = T0.""U_METWHS"" And IFNULL(""U_MEEXACCT"",'')<>'')) As  ""GIAccountName""  , (Select ""U_MEENTACT""  From ""OWHS""  " &
                                  " Where ""WhsCode"" = T0.""U_METRWHS"" And IFNULL(""U_MEENTACT"",'')<>'' ) As ""GRAccountCode"" , (Select ""AcctName"" From OACT Where ""AcctCode"" = (Select ""U_MEENTACT""  From ""OWHS""  " &
                                  " Where ""WhsCode"" = T0.""U_METRWHS"" And IFNULL(""U_MEENTACT"",'')<>'')) As  ""GRAccountName""   , ""U_METWHS"" As ""ToWhs"" , " &
                                  " IFNULL(""U_MEGIEN"",'') As ""IssueEntry"", IFNULL(""U_MEGINO"",'') As ""GoodsIssueNo""  , " &
                                  " IFNULL(""U_MEGREN"",'') As ""ReceiptEntry"" ,  IFNULL(""U_MEGRNO"",'') As ""GoodsReceiptNo"" , " &
                                 "  Cast('' AS NVARCHAR(5000)) As ""Status""   From ""@ME_MI"" T0 Where (IFNULL(""U_MEGINO"",'')='' Or IFNULL(""U_MEGRNO"",'')='')  And  " &
                                 "  IFNULL(""U_MESTS"",'')='Open' And  T0.""U_MEDATE"" Between  '20210101' And '" & objForm.Items.Item("4").Specific.Value & "'  "
            If objForm.DataSources.UserDataSources.Item("U_FRBRCH").Value <> "" Then
                Query = Query & " And T0.""U_MEFBPID"" = '" & objForm.Items.Item("15").Specific.Value & "' "
            End If
            If objForm.DataSources.UserDataSources.Item("U_TOBRCH").Value <> "" Then
                Query = Query & " And T0.""U_METBPID"" = '" & objForm.Items.Item("1000001").Specific.Value & "' "
            End If
            oGrid.DataTable.ExecuteQuery(Query)

            oGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item("MIEntry").Visible = False
            oGrid.Columns.Item("FromBranchId").Visible = False
            oGrid.Columns.Item("TRWhs").Visible = True
            oGrid.Columns.Item("ToWhs").Visible = True
            oGrid.Columns.Item("MINo").Editable = False
            oGrid.Columns.Item("MIDate").Editable = False
            oGrid.Columns.Item("FromBranch").Editable = False
            oGrid.Columns.Item("TransferNumber").Editable = False
            oGrid.Columns.Item("Status").Editable = False
            oGrid.Columns.Item("GIAccountCode").Editable = False
            oGrid.Columns.Item("GRAccountCode").Editable = False
            oGrid.Columns.Item("GIAccountName").Editable = False
            oGrid.Columns.Item("GRAccountName").Editable = False
            oGrid.Columns.Item("GoodsIssueNo").Editable = False
            oGrid.Columns.Item("GoodsReceiptNo").Editable = False
            oGrid.Columns.Item("IssueEntry").Editable = False
            oGrid.Columns.Item("ReceiptEntry").Editable = False
            oGrid.Columns.Item("TransferId").Editable = False
            oGrid.Columns.Item("Status").Editable = False
            oGrid.Columns.Item("TransferId").Visible = False
            oGrid.Columns.Item("TransferNumber").Visible = False
            oGrid.Columns.Item("TRWhs").Editable = False
            oGrid.Columns.Item("ToWhs").Editable = False
            objForm.Items.Item("11").Visible = False
            If oGrid.DataTable.GetValue("TransferNumber", 0) = "" Then
                objMain.objApplication.MessageBox("No records found")
                objForm.Freeze(False)
                Exit Sub
            End If

            Dim oEditTextCol As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("GIAccountCode")
            oEditTextCol.LinkedObjectType = 1

            Dim oEditTextCol1 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("GRAccountCode")
            oEditTextCol1.LinkedObjectType = 1

            Dim oEditTextCol2 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("IssueEntry")
            oEditTextCol2.LinkedObjectType = 60

            Dim oEditTextCol3 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("ReceiptEntry")
            oEditTextCol3.LinkedObjectType = 59

            Dim oEditTextCol4 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("MINo")
            oEditTextCol4.LinkedObjectType = 17

            Dim oEditTextCol5 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("TransferId")
            oEditTextCol5.LinkedObjectType = 67


            For x As Integer = 0 To oGrid.Rows.Count - 1
                oGrid.CommonSetting.SetRowFontStyle(x + 1, SAPbouiCOM.BoFontStyle.fs_Bold)
                oGrid.CommonSetting.SetRowFontColor(x + 1, ColorTranslator.ToOle(Color.Black))
            Next

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        Finally
            oGrid.AutoResizeColumns()
            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

        End Try
    End Sub

    Private Sub PostGoodsIssue(ByVal FormUID As String)
        Dim oProgressPostGoodsIssue As SAPbouiCOM.ProgressBar = Nothing
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("10").Specific
            oProgressPostGoodsIssue = objMain.objApplication.StatusBar.CreateProgressBar("Please Wait Issue docments generating....!", oGrid.DataTable.Rows.Count - 1, False)
            oProgressPostGoodsIssue.Text = "Please Wait Issue docments generating....!"
            Dim j1 As Integer = 1
            For j As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If oGrid.DataTable.GetValue("Check", j) = "Y" And oGrid.DataTable.GetValue("GoodsIssueNo", j) = "" Then
                    Try
                        Dim GetMaterialInwardDtls As String = "Select (Select ""IUoMEntry"" From OITM Where ""ItemCode"" = ""U_MEITMCD"") As ""DfltUOM"" , *  From ""@ME_MI_C0"" Where ""DocEntry"" = '" & oGrid.DataTable.GetValue("MIEntry", j) & "' And  U_MERECQTY > 0  "
                        Dim oRsGetMaterialInwardDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetMaterialInwardDtls.DoQuery(GetMaterialInwardDtls)
                        Dim GetAddonConfigurationDtls As String = "Select ""U_MEEXACCT"" From ""OWHS"" Where ""WhsCode"" = '" & oGrid.DataTable.GetValue("ToWhs", j) & "' And IFNULL(""U_MEEXACCT"",'')<>''  "
                        Dim oRsGetAddonConfigurationDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetAddonConfigurationDtls.DoQuery(GetAddonConfigurationDtls)
                        If oRsGetAddonConfigurationDtls.RecordCount = 0 Then
                            oGrid.DataTable.SetValue("Status", j, "Please Configure Issue Account in Warehouse Master")
                        End If
                        Dim oInventoryGenExit As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
                        Try
                            oInventoryGenExit.UserFields.Fields.Item("U_TYPE").Value = "STR"
                        Catch ex As Exception
                        End Try
                        oInventoryGenExit.DocDate = oGrid.DataTable.GetValue("MIDate", j)
                        oInventoryGenExit.DocDueDate = oGrid.DataTable.GetValue("MIDate", j)
                        oInventoryGenExit.TaxDate = oGrid.DataTable.GetValue("MIDate", j)
                        Dim sDate As Date = oGrid.DataTable.GetValue("MIDate", j)
                        oInventoryGenExit.UserFields.Fields.Item("U_ADDON").Value = "Yes"
                        oInventoryGenExit.JournalMemo = "Goods Issue based on MI No " & oGrid.DataTable.GetValue("MINo", j)
                        'objComboBox = objForm.Items.Item("15").Specific
                        Dim GetSeries As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = '60' And ""Locked"" = 'N'  And ""BPLId"" = '" & oGrid.DataTable.GetValue("FromBranchId", j) & "' And  ""Indicator"" = '" & sDate.Year & "' "
                        Dim oRsGetSeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetSeries.DoQuery(GetSeries)
                        If oRsGetSeries.RecordCount > 0 Then
                            oInventoryGenExit.Series = oRsGetSeries.Fields.Item("Series").Value
                        End If
                        oInventoryGenExit.BPL_IDAssignedToInvoice = oGrid.DataTable.GetValue("FromBranchId", j) 'objComboBox.Value
                        For i As Integer = 1 To oRsGetMaterialInwardDtls.RecordCount
                            If i <> 1 Then
                                oInventoryGenExit.Lines.Add()
                            End If
                            oInventoryGenExit.Lines.ItemCode = oRsGetMaterialInwardDtls.Fields.Item("U_MEITMCD").Value
                            If oRsGetMaterialInwardDtls.Fields.Item("U_MEBRCD").Value <> "" Then
                                oInventoryGenExit.Lines.BarCode = oRsGetMaterialInwardDtls.Fields.Item("U_MEBRCD").Value
                            End If
                            oInventoryGenExit.Lines.WarehouseCode = oGrid.DataTable.GetValue("TRWhs", j)
                            Dim UomEntry As Integer = Me.GetUomEntry(oRsGetMaterialInwardDtls.Fields.Item("U_MEUOM").Value)
                            If UomEntry = CInt(oRsGetMaterialInwardDtls.Fields.Item("DfltUOM").Value) Then
                                oInventoryGenExit.Lines.UseBaseUnits = SAPbobsCOM.BoYesNoEnum.tYES
                                oInventoryGenExit.Lines.UoMEntry = UomEntry
                                If oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "BatchNo" Then
                                    oInventoryGenExit.Lines.BatchNumbers.BatchNumber = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
                                    oInventoryGenExit.Lines.BatchNumbers.Quantity = CDbl(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value)
                                    oInventoryGenExit.Lines.BatchNumbers.Add()
                                ElseIf oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "SerialNo" Then
                                    Dim TotalSlrBndBinWiseQty As String = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
                                    Dim ArrayTotalSlrBndBinWiseQty() As String = TotalSlrBndBinWiseQty.Split(";")
                                    For k As Integer = 0 To CInt(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value) - 1 'ArrayTotalSlrBndBinWiseQty.Length - 1
                                        If ArrayTotalSlrBndBinWiseQty(k) <> "" Then
                                            oInventoryGenExit.Lines.SerialNumbers.ManufacturerSerialNumber = ArrayTotalSlrBndBinWiseQty(k)
                                            oInventoryGenExit.Lines.SerialNumbers.Add()
                                        End If
                                    Next
                                End If
                            Else
                                oInventoryGenExit.Lines.UseBaseUnits = SAPbobsCOM.BoYesNoEnum.tNO
                                oInventoryGenExit.Lines.UoMEntry = UomEntry
                                If oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "BatchNo" Then
                                    Dim BaseQty As String = " Select ""BaseQty"" * " & oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value & "   From UGP1 T0 Inner Join OITM T1 On " &
                                                            " T1.""UgpEntry"" =  T0.""UgpEntry"" Where ""ItemCode"" = '" &
                                                            oRsGetMaterialInwardDtls.Fields.Item("U_MEITMCD").Value & "' And T0.""UomEntry"" = '" & UomEntry & "'"
                                    Dim oRsBaseQty As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRsBaseQty.DoQuery(BaseQty)
                                    oInventoryGenExit.Lines.BatchNumbers.BatchNumber = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
                                    oInventoryGenExit.Lines.BatchNumbers.Quantity = oRsBaseQty.Fields.Item(0).Value
                                    oInventoryGenExit.Lines.BatchNumbers.Add()
                                ElseIf oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "SerialNo" Then
                                    Dim TotalSlrBndBinWiseQty As String = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
                                    Dim ArrayTotalSlrBndBinWiseQty() As String = TotalSlrBndBinWiseQty.Split(";")
                                    For k As Integer = 0 To CInt(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value) - 1 ' ArrayTotalSlrBndBinWiseQty.Length - 1
                                        If ArrayTotalSlrBndBinWiseQty(k) <> "" Then
                                            oInventoryGenExit.Lines.SerialNumbers.ManufacturerSerialNumber = ArrayTotalSlrBndBinWiseQty(k)
                                            oInventoryGenExit.Lines.SerialNumbers.Add()
                                        End If
                                    Next
                                End If
                            End If
                            oInventoryGenExit.Lines.Quantity = CDbl(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value)
                            oInventoryGenExit.Lines.AccountCode = oRsGetAddonConfigurationDtls.Fields.Item("U_MEEXACCT").Value
                            Try
                                oInventoryGenExit.Lines.UserFields.Fields.Item("U_Addbarcode").Value = oRsGetMaterialInwardDtls.Fields.Item("U_ADBRCD").Value
                                oInventoryGenExit.Lines.UserFields.Fields.Item("U_Packing").Value = oRsGetMaterialInwardDtls.Fields.Item("U_MEPKGE").Value
                            Catch ex As Exception
                            End Try
                            oRsGetMaterialInwardDtls.MoveNext()
                        Next
                        Dim sIssueGUID As String
                        sIssueGUID = System.Guid.NewGuid.ToString()
                        Try
                            oInventoryGenExit.UserFields.Fields.Item("U_GUID").Value = sIssueGUID
                        Catch ex As Exception
                        End Try

                        'Level 1 Chk
                        Dim JounalMemo As String = "Goods Issue based on MI No " & oGrid.DataTable.GetValue("MINo", j)
                        Dim chkIfExisted As String = " Select ""DocNum"" From ""OIGE"" Where  ""U_TYPE"" = 'STR'  And IFNULL(""JrnlMemo"",'') <> 'Goods Issue' And IFNULL(""JrnlMemo"",'') = '" & JounalMemo & "'"
                        Dim oRschkIfExisted As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRschkIfExisted.DoQuery(chkIfExisted)
                        If oRschkIfExisted.Fields.Item(0).Value > 0 Then
                            oGrid.CommonSetting.SetRowFontColor(j + 1, ColorTranslator.ToOle(Color.DarkRed))
                            'oGrid.DataTable.SetValue("Status", j, "Journal Remark already existed in GINum " & oRschkIfExisted.Fields.Item("DocNum").Value & " : " & JounalMemo)
                            oGrid.DataTable.SetValue("Status", j, "Document already posted")
                            GoTo CH
                        End If
                        'Level 2 Chk                      
                        Dim chkMIInIssue As String = " Select ""DocNum"" From ""OIGE"" Where  ""U_TYPE"" = 'STR'  And IFNULL(""U_MIDNUM"",'') <> '' And  IFNULL(""U_MIDNUM"",'') = '" & oGrid.DataTable.GetValue("MINo", j) & "'   "
                        Dim oRschkMIInIssue As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRschkMIInIssue.DoQuery(chkMIInIssue)
                        If oRschkMIInIssue.Fields.Item(0).Value > 0 Then
                            oGrid.CommonSetting.SetRowFontColor(j + 1, ColorTranslator.ToOle(Color.DarkRed))
                            oGrid.DataTable.SetValue("Status", j, "Document already posted")
                            GoTo CH
                        End If

                        If oInventoryGenExit.Add() = 0 Then
                            Dim GetOrderEntry As String = "Select ""DocEntry"" , ""DocNum""   From OIGE Where ""U_GUID"" = '" & sIssueGUID & "' "
                            Dim oRsGetOrderEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetOrderEntry.DoQuery(GetOrderEntry)

                            Dim MIUpdate As String = "Update ""@ME_MI"" Set  ""U_MEGINO"" = '" & oRsGetOrderEntry.Fields.Item(1).Value.ToString & "' , ""U_MEGIEN"" = '" & oRsGetOrderEntry.Fields.Item(0).Value.ToString & "' , ""U_REMARKS"" = ' '   Where ""DocEntry"" = '" & oGrid.DataTable.GetValue("MIEntry", j) & "' "
                            Dim oRsMIUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsMIUpdate.DoQuery(MIUpdate)

                            Dim GIUpdate As String = "Update ""OIGE"" Set  ""U_MIDNUM"" = '" & oGrid.DataTable.GetValue("MINo", j) & "' , ""U_MIDENT"" = '" & oGrid.DataTable.GetValue("MIEntry", j) & "'  Where ""DocEntry"" = '" & oRsGetOrderEntry.Fields.Item(0).Value.ToString & "' "
                            Dim oRsGIUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGIUpdate.DoQuery(GIUpdate)
                            oGrid.DataTable.SetValue("GoodsIssueNo", j, oRsGetOrderEntry.Fields.Item(1).Value.ToString)
                            oGrid.DataTable.SetValue("IssueEntry", j, oRsGetOrderEntry.Fields.Item(0).Value.ToString)
                            oGrid.CommonSetting.SetRowFontColor(j + 1, ColorTranslator.ToOle(Color.DarkGreen))
                            oGrid.DataTable.SetValue("Status", j, "Issue Posted Successfully")
                        Else
                            objMain.sCmp = objMain.objCompany.GetCompanyService
                            objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OMI")
                            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                            objMain.oGeneralParams.SetProperty("DocEntry", oGrid.DataTable.GetValue("MIEntry", j))
                            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                            objMain.oGeneralData.SetProperty("U_REMARKS", "ISSUE : " & objMain.objCompany.GetLastErrorDescription.ToString)
                            objMain.oGeneralService.Update(objMain.oGeneralData)
                            oGrid.CommonSetting.SetRowFontColor(j + 1, ColorTranslator.ToOle(Color.DarkRed))
                            oGrid.DataTable.SetValue("Status", j, "ISSUE : " & objMain.objCompany.GetLastErrorDescription.ToString)
                        End If
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oInventoryGenExit)
                    Catch ex As Exception
                        oGrid.CommonSetting.SetRowFontColor(j + 1, ColorTranslator.ToOle(Color.DarkRed))
                        oGrid.DataTable.SetValue("Status", j, "ISSUE : " & ex.Message)
                    End Try
                End If
CH:             oProgressPostGoodsIssue.Value = j1
                j1 = j1 + 1
            Next
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        Finally
            oGrid.AutoResizeColumns()
            If oProgressPostGoodsIssue IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgressPostGoodsIssue)
            End If
        End Try
    End Sub

    Private Sub PostGoodsReceipts(ByVal FormUID As String)
        Dim oProgressPostGoodsReceipt As SAPbouiCOM.ProgressBar = Nothing
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("10").Specific
            oProgressPostGoodsReceipt = objMain.objApplication.StatusBar.CreateProgressBar("Please Wait Receipt docments generating....!", oGrid.DataTable.Rows.Count - 1, False)
            oProgressPostGoodsReceipt.Text = "Please Wait Receipt docments generating....!"
            Dim j1 As Integer = 1
            For j As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If oGrid.DataTable.GetValue("Check", j) = "Y" And oGrid.DataTable.GetValue("GoodsIssueNo", j) <> "" And oGrid.DataTable.GetValue("GoodsReceiptNo", j) = "" Then
                    Try


                        Dim oRsBaseQty As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim GetMaterialInwardDtls As String = "Select  (Select ""IUoMEntry"" From OITM Where ""ItemCode"" = ""U_MEITMCD"") As ""DfltUOM"" , (Select ""AbsEntry"" From OBIN Where ""BinCode"" = ""U_METWBL"") As ""BinAbsEntry"" ,  * From ""@ME_MI_C0"" Where ""DocEntry"" = '" & oGrid.DataTable.GetValue("MIEntry", j) & "'  And  U_MERECQTY > 0  "
                        Dim oRsGetMaterialInwardDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetMaterialInwardDtls.DoQuery(GetMaterialInwardDtls)
                        objComboBox = objForm.Items.Item("1000001").Specific

                        '' Dim GetAddonConfigurationDtls As String = "Select ""U_MEENTACT""  From ""OWHS"" Where ""WhsCode"" = '" & oGrid.DataTable.GetValue("GIAccountCode", j) & "' And IFNULL(""U_MEENTACT"",'')<>''  "
                        Dim GetAddonConfigurationDtls As String = "Select ""U_MEENTACT""  From ""OWHS"" Where ""WhsCode"" = '" & oGrid.DataTable.GetValue("TRWhs", j) & "' And IFNULL(""U_MEENTACT"",'')<>''  "
                        Dim oRsGetAddonConfigurationDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetAddonConfigurationDtls.DoQuery(GetAddonConfigurationDtls)
                        If oRsGetAddonConfigurationDtls.RecordCount = 0 Then
                            oGrid.DataTable.SetValue("Status", j, "Please Configure Receipt Account in Warehouse Master")
                            Exit Sub
                        End If
                        Dim oInventoryGenEntry As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
                        Try
                            oInventoryGenEntry.UserFields.Fields.Item("U_TYPE").Value = "STR"
                        Catch ex As Exception
                        End Try
                        oInventoryGenEntry.DocDate = oGrid.DataTable.GetValue("MIDate", j)
                        oInventoryGenEntry.DocDueDate = oGrid.DataTable.GetValue("MIDate", j)
                        oInventoryGenEntry.TaxDate = oGrid.DataTable.GetValue("MIDate", j)
                        Dim sDate As Date = oGrid.DataTable.GetValue("MIDate", j)
                        oInventoryGenEntry.UserFields.Fields.Item("U_ADDON").Value = "Yes"

                        objComboBox = objForm.Items.Item("1000001").Specific
                        oInventoryGenEntry.BPL_IDAssignedToInvoice = objComboBox.Value
                        oInventoryGenEntry.JournalMemo = "Goods Receipt based on MI No " & oGrid.DataTable.GetValue("MINo", j)
                        Dim GetSeries As String = " Select ""Series"" , ""SeriesName""   From NNM1 Where ""ObjectCode"" = '59' And ""Locked"" = 'N' And ""BPLId"" =  '" & objComboBox.Value & "' And ""Indicator"" = '" & sDate.Year & "' "
                        Dim oRsGetSeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetSeries.DoQuery(GetSeries)

                        oInventoryGenEntry.Series = oRsGetSeries.Fields.Item("Series").Value

                        For i As Integer = 1 To oRsGetMaterialInwardDtls.RecordCount
                            Dim binlineno As Integer = 0
                            If i <> 1 Then
                                oInventoryGenEntry.Lines.Add()
                            End If
                            'Dim GetItemValue As String = " Select ""StockPrice"" From IGE1 Where ""DocEntry"" = '" & oGrid.DataTable.GetValue("IssueEntry", j) _
                            '                             & "' And ""ItemCode""  = '" & oRsGetMaterialInwardDtls.Fields.Item("U_MEITMCD").Value & "'"
                            Dim GetItemValue As String = " Select  ABS(SUM(""TransValue"")) As ""TransValue"" FROM OINM Where ""TransType""='60' And " &
                                           " ""CreatedBy""= '" & oGrid.DataTable.GetValue("IssueEntry", j) & "' And ""ItemCode"" = '" & oRsGetMaterialInwardDtls.Fields.Item("U_MEITMCD").Value & "' Group By  ""DocLineNum""  "
                            Dim oRsGetItemValue As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetItemValue.DoQuery(GetItemValue)
                            oInventoryGenEntry.Lines.ItemCode = oRsGetMaterialInwardDtls.Fields.Item("U_MEITMCD").Value
                            If oRsGetMaterialInwardDtls.Fields.Item("U_MEBRCD").Value <> "" Then
                                oInventoryGenEntry.Lines.BarCode = oRsGetMaterialInwardDtls.Fields.Item("U_MEBRCD").Value
                            End If
                            oInventoryGenEntry.Lines.WarehouseCode = oGrid.DataTable.GetValue("ToWhs", j)
                            oInventoryGenEntry.Lines.Quantity = CDbl(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value)
                            If oRsGetItemValue.RecordCount > 0 Then
                                'oInventoryGenEntry.Lines.Price = CDbl(oRsGetItemValue.Fields.Item("StockPrice").Value)
                                'oInventoryGenEntry.Lines.UnitPrice = CDbl(oRsGetItemValue.Fields.Item("StockPrice").Value)
                                oInventoryGenEntry.Lines.LineTotal = CDbl(oRsGetItemValue.Fields.Item("TransValue").Value)
                            End If
                            oInventoryGenEntry.Lines.AccountCode = oRsGetAddonConfigurationDtls.Fields.Item("U_MEENTACT").Value
                            Dim UomEntry As Integer = Me.GetUomEntry(oRsGetMaterialInwardDtls.Fields.Item("U_MEUOM").Value)
                            If UomEntry = CInt(oRsGetMaterialInwardDtls.Fields.Item("DfltUOM").Value) Then
                                oInventoryGenEntry.Lines.UseBaseUnits = SAPbobsCOM.BoYesNoEnum.tYES
                                oInventoryGenEntry.Lines.UoMEntry = UomEntry
                                If oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "BatchNo" Then
                                    oInventoryGenEntry.Lines.BatchNumbers.BatchNumber = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
                                    oInventoryGenEntry.Lines.BatchNumbers.Quantity = CDbl(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value)
                                    oInventoryGenEntry.Lines.BatchNumbers.Add()
                                    If oRsGetMaterialInwardDtls.Fields.Item("U_CHKBIN").Value.Trim = "Y" Then
                                        oInventoryGenEntry.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = binlineno
                                        oInventoryGenEntry.Lines.BatchNumbers.BatchNumber = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
                                        oInventoryGenEntry.Lines.BinAllocations.BinAbsEntry = oRsGetMaterialInwardDtls.Fields.Item("BinAbsEntry").Value
                                        oInventoryGenEntry.Lines.BinAllocations.Quantity = CDbl(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value)
                                        oInventoryGenEntry.Lines.BinAllocations.Add()
                                    End If
                                ElseIf oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "SerialNo" Then
                                    Dim TotalSlrBndBinWiseQty As String = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
                                    Dim ArrayTotalSlrBndBinWiseQty() As String = TotalSlrBndBinWiseQty.Split(";")
                                    For k As Integer = 0 To CInt(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value) - 1 'ArrayTotalSlrBndBinWiseQty.Length - 1
                                        If ArrayTotalSlrBndBinWiseQty(k) <> "" Then
                                            oInventoryGenEntry.Lines.SerialNumbers.ManufacturerSerialNumber = ArrayTotalSlrBndBinWiseQty(k)
                                            oInventoryGenEntry.Lines.SerialNumbers.Add()
                                            If oRsGetMaterialInwardDtls.Fields.Item("U_CHKBIN").Value.Trim = "Y" Then
                                                oInventoryGenEntry.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = binlineno
                                                oInventoryGenEntry.Lines.BatchNumbers.BatchNumber = ArrayTotalSlrBndBinWiseQty(k)
                                                oInventoryGenEntry.Lines.BinAllocations.BinAbsEntry = oRsGetMaterialInwardDtls.Fields.Item("BinAbsEntry").Value
                                                oInventoryGenEntry.Lines.BinAllocations.Quantity = 1 'CDbl(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value)
                                                oInventoryGenEntry.Lines.BinAllocations.Add()
                                                binlineno = binlineno + 1
                                            End If
                                        End If
                                    Next
                                End If
                            Else
                                oInventoryGenEntry.Lines.UseBaseUnits = SAPbobsCOM.BoYesNoEnum.tNO
                                oInventoryGenEntry.Lines.UoMEntry = UomEntry
                                If oRsGetItemValue.RecordCount > 0 Then
                                    'Dim BaseQty As String = " Select ""BaseQty"" * " & oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value & "   From UGP1 T0 Inner Join OITM T1 On " & _
                                    '                        " T1.""UgpEntry"" =  T0.""UgpEntry"" Where ""ItemCode"" = '" & _
                                    '                       oRsGetMaterialInwardDtls.Fields.Item("U_MEITMCD").Value & "' And T0.""UomEntry"" = '" & UomEntry & "'"
                                    'oRsBaseQty.DoQuery(BaseQty)
                                    'oInventoryGenEntry.Lines.LineTotal = CDbl(oRsGetItemValue.Fields.Item("StockPrice").Value) * oRsBaseQty.Fields.Item(0).Value
                                    oInventoryGenEntry.Lines.LineTotal = CDbl(oRsGetItemValue.Fields.Item("TransValue").Value)
                                End If
                                If oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "BatchNo" Then
                                    oInventoryGenEntry.Lines.BatchNumbers.BatchNumber = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
                                    oInventoryGenEntry.Lines.BatchNumbers.Quantity = oRsBaseQty.Fields.Item(0).Value
                                    oInventoryGenEntry.Lines.BatchNumbers.Add()
                                    If oRsGetMaterialInwardDtls.Fields.Item("U_CHKBIN").Value.Trim = "Y" Then
                                        oInventoryGenEntry.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = binlineno
                                        oInventoryGenEntry.Lines.BatchNumbers.BatchNumber = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
                                        oInventoryGenEntry.Lines.BinAllocations.BinAbsEntry = oRsGetMaterialInwardDtls.Fields.Item("BinAbsEntry").Value
                                        oInventoryGenEntry.Lines.BinAllocations.Quantity = oRsBaseQty.Fields.Item(0).Value
                                        oInventoryGenEntry.Lines.BinAllocations.Add()
                                    End If
                                ElseIf oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "SerialNo" Then
                                    Dim TotalSlrBndBinWiseQty As String = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
                                    Dim ArrayTotalSlrBndBinWiseQty() As String = TotalSlrBndBinWiseQty.Split(";")
                                    For k As Integer = 0 To CInt(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value) - 1  ' ArrayTotalSlrBndBinWiseQty.Length - 1
                                        If ArrayTotalSlrBndBinWiseQty(k) <> "" Then
                                            oInventoryGenEntry.Lines.SerialNumbers.ManufacturerSerialNumber = ArrayTotalSlrBndBinWiseQty(k)
                                            oInventoryGenEntry.Lines.SerialNumbers.Add()
                                            If oRsGetMaterialInwardDtls.Fields.Item("U_CHKBIN").Value.Trim = "Y" Then
                                                Dim BaseQty As String = " Select ""BaseQty"" * " & oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value & "   From UGP1 T0 Inner Join OITM T1 On " &
                                                                    " T1.""UgpEntry"" =  T0.""UgpEntry"" Where ""ItemCode"" = '" &
                                            oRsGetMaterialInwardDtls.Fields.Item("U_MEITMCD").Value & "' And T0.""UomEntry"" = '" & UomEntry & "'"
                                                oRsBaseQty.DoQuery(BaseQty)
                                                oInventoryGenEntry.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = binlineno
                                                oInventoryGenEntry.Lines.BatchNumbers.BatchNumber = ArrayTotalSlrBndBinWiseQty(k)
                                                oInventoryGenEntry.Lines.BinAllocations.BinAbsEntry = oRsGetMaterialInwardDtls.Fields.Item("BinAbsEntry").Value
                                                oInventoryGenEntry.Lines.BinAllocations.Quantity = 1
                                                oInventoryGenEntry.Lines.BinAllocations.Add()
                                                binlineno = binlineno + 1
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                            Try
                                oInventoryGenEntry.Lines.UserFields.Fields.Item("U_Addbarcode").Value = oRsGetMaterialInwardDtls.Fields.Item("U_ADBRCD").Value
                                oInventoryGenEntry.Lines.UserFields.Fields.Item("U_Packing").Value = oRsGetMaterialInwardDtls.Fields.Item("U_MEPKGE").Value
                            Catch ex As Exception
                            End Try
                            binlineno = binlineno + 1
                            oRsGetMaterialInwardDtls.MoveNext()
                        Next
                        Dim sGoodsReceiptGUID As String
                        sGoodsReceiptGUID = System.Guid.NewGuid.ToString()
                        Try
                            oInventoryGenEntry.UserFields.Fields.Item("U_GUID").Value = sGoodsReceiptGUID
                        Catch ex As Exception
                        End Try

                        'Level 1 Chk
                        Dim JounalMemo As String = "Goods Receipt based on MI No " & oGrid.DataTable.GetValue("MINo", j)
                        Dim chkIfExisted As String = " Select ""DocNum"" From ""OIGN"" Where  ""U_TYPE"" = 'STR'  And IFNULL(""JrnlMemo"",'') <> 'Goods Receipt' And IFNULL(""JrnlMemo"",'') = '" & JounalMemo & "'"
                        Dim oRschkIfExisted As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRschkIfExisted.DoQuery(chkIfExisted)
                        If oRschkIfExisted.Fields.Item(0).Value > 0 Then
                            oGrid.CommonSetting.SetRowFontColor(j + 1, ColorTranslator.ToOle(Color.DarkRed))
                            oGrid.DataTable.SetValue("Status", j, "Document already posted")
                            GoTo CH
                        End If

                        'Level 2 Chk                      
                        Dim chkMIInIssue As String = " Select ""DocNum"" From ""OIGN"" Where  ""U_TYPE"" = 'STR'  And IFNULL(""U_MIDNUM"",'') <> '' And  IFNULL(""U_MIDNUM"",'') = '" & oGrid.DataTable.GetValue("MINo", j) & "'   "
                        Dim oRschkMIInIssue As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRschkMIInIssue.DoQuery(chkMIInIssue)
                        If oRschkMIInIssue.Fields.Item(0).Value > 0 Then
                            oGrid.CommonSetting.SetRowFontColor(j + 1, ColorTranslator.ToOle(Color.DarkRed))
                            oGrid.DataTable.SetValue("Status", j, "Document already posted")
                            GoTo CH
                        End If
                        If oInventoryGenEntry.Add() = 0 Then
                            Dim GetOrderEntry As String = "Select ""DocEntry"" , ""DocNum""   From OIGN  Where  ""U_GUID"" = '" & sGoodsReceiptGUID & "' "
                            Dim oRsGetOrderEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetOrderEntry.DoQuery(GetOrderEntry)

                            Dim MIUpdate As String = "Update ""@ME_MI"" Set  ""U_MEGRNO"" = '" & oRsGetOrderEntry.Fields.Item(1).Value.ToString & "' , ""U_MEGREN"" = '" & oRsGetOrderEntry.Fields.Item(0).Value.ToString & "' , ""U_REMARKS"" = ' ' , ""U_MESTS"" = 'Close'  Where ""DocEntry"" = '" & oGrid.DataTable.GetValue("MIEntry", j) & "' "
                            Dim oRsMIUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsMIUpdate.DoQuery(MIUpdate)

                            Dim GRUpdate As String = "Update ""OIGN"" Set  ""U_MIDNUM"" = '" & oGrid.DataTable.GetValue("MINo", j) & "' , ""U_MIDENT"" = '" & oGrid.DataTable.GetValue("MIEntry", j) & "'  Where ""DocEntry"" = '" & oRsGetOrderEntry.Fields.Item(0).Value.ToString & "' "
                            Dim oRsGRUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGRUpdate.DoQuery(GRUpdate)

                            oGrid.DataTable.SetValue("GoodsReceiptNo", j, oRsGetOrderEntry.Fields.Item(1).Value.ToString)
                            oGrid.DataTable.SetValue("ReceiptEntry", j, oRsGetOrderEntry.Fields.Item(0).Value.ToString)
                            oGrid.CommonSetting.SetRowFontColor(j + 1, ColorTranslator.ToOle(Color.DarkGreen))
                            oGrid.DataTable.SetValue("Status", j, "Receipt Posted Successfully")
                        Else
                            objMain.sCmp = objMain.objCompany.GetCompanyService
                            objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OMI")
                            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                            objMain.oGeneralParams.SetProperty("DocEntry", oGrid.DataTable.GetValue("MIEntry", j))
                            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                            objMain.oGeneralData.SetProperty("U_REMARKS", "Receipt : " & objMain.objCompany.GetLastErrorDescription.ToString)
                            objMain.oGeneralService.Update(objMain.oGeneralData)
                            oGrid.CommonSetting.SetRowFontColor(j + 1, ColorTranslator.ToOle(Color.DarkRed))
                            oGrid.DataTable.SetValue("Status", j, "Receipt : " & objMain.objCompany.GetLastErrorDescription.ToString)
                        End If
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oInventoryGenEntry)
                    Catch ex As Exception
                        oGrid.CommonSetting.SetRowFontColor(j + 1, ColorTranslator.ToOle(Color.DarkRed))
                        oGrid.DataTable.SetValue("Status", j, ex.Message)
                    End Try
                End If
CH:             oProgressPostGoodsReceipt.Value = j1
                j1 = j1 + 1
            Next
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        Finally
            oGrid.AutoResizeColumns()
            If oProgressPostGoodsReceipt IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgressPostGoodsReceipt)
            End If
        End Try
    End Sub

    Sub StartBackgroundForLoading(Optional ByVal BackGroundWrkr As Boolean = True)
        Try
            If BackGroundWrkr = True Then
                If BackgroundWorker_GenerateDocumets.IsBusy <> True Then
                    BackgroundWorker_GenerateDocumets.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Calling Unselecting worker Error : " & ex.Message)
        End Try
    End Sub

    Private Sub BackgroundWorker_Loading_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_GenerateDocumets.DoWork
        If DocFlag = "Issue" Then
            Me.PostGoodsIssue(objForm.UniqueID)
        End If
        If DocFlag = "Receipt" Then
            PostGoodsReceipts(objForm.UniqueID)
        End If
    End Sub

    Private Sub backgroundWorker1_Loading_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As RunWorkerCompletedEventArgs) Handles BackgroundWorker_GenerateDocumets.RunWorkerCompleted
        If (e.Error IsNot Nothing) Then
            objMain.objApplication.StatusBar.SetText("Loading error :" & e.Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        ElseIf e.Cancelled Then
            objMain.objApplication.StatusBar.SetText("Canceled", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Else
            objMain.objApplication.StatusBar.SetText("Operation completed successfully..." & e.Result, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objForm.Freeze(False)
        End If
    End Sub

    Public Function GetUomEntry(ByVal sUomCode As String) As Integer
        Try
            Dim GetDocEntry As String = "Select ""UomEntry"" From OUOM Where ""UomCode"" = '" & sUomCode & "' "
            Dim oRsGetDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDocEntry.DoQuery(GetDocEntry)
            Return CInt(oRsGetDocEntry.Fields.Item("UoMEntry").Value)
        Catch ex As Exception
        End Try
    End Function

End Class
