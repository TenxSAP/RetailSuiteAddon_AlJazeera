
Public Class clsProductionGenerationWizard

#Region "        Declaration        "
    Dim objForm, objMIForm As SAPbouiCOM.Form
    Dim oGrid As SAPbouiCOM.Grid
    Dim oDt As SAPbouiCOM.DataTable
    Dim oEditText As SAPbouiCOM.EditText
    Dim objComboBox As SAPbouiCOM.ComboBox
    Private WithEvents BackgroundWorker_GenerateDocumets As New System.ComponentModel.BackgroundWorker
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("ProductionGenerationWizard.xml", "SBO_PGW_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("SBO_PGW_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.DataSources.UserDataSources.Add("U_FROMDT", SAPbouiCOM.BoDataType.dt_DATE)
            objForm.DataSources.UserDataSources.Add("U_TODT", SAPbouiCOM.BoDataType.dt_DATE)
            objForm.DataSources.UserDataSources.Add("U_WHS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            oEditText = objForm.Items.Item("3").Specific
            oEditText.DataBind.SetBound(True, "", "U_FROMDT")
            objForm.Items.Item("3").AffectsFormMode = False

            oEditText = objForm.Items.Item("5").Specific
            oEditText.DataBind.SetBound(True, "", "U_TODT")
            objForm.Items.Item("5").AffectsFormMode = False

            objComboBox = objForm.Items.Item("Item_1").Specific
            objComboBox.DataBind.SetBound(True, "", "U_WHS")
            objForm.Items.Item("Item_1").AffectsFormMode = False
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, "Select ""WhsCode"" , ""WhsCode"" From ""OWHS"" Where ""U_Transit"" = 'No'")
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            'If objComboBox.ValidValues.Count <> 0 Then
            '    objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
            'End If

            oDt = objForm.DataSources.DataTables.Add("dt1")
            oDt = objForm.DataSources.DataTables.Item("dt1")

            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "SBO_PGW" And pVal.BeforeAction = False Then
                Me.CreateForm()
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oGrid = objForm.Items.Item("6").Specific
                    If pVal.ItemUID = "6" And pVal.BeforeAction = False Then
                        Try
                            oGrid.Columns.Item("RowsHeader").Click(pVal.Row, SAPbouiCOM.BoCellClickType.ct_Regular)
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "10" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("3").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        ElseIf objForm.Items.Item("5").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            Me.LoadGrid(objForm.UniqueID, "Y")
                        End If
                    End If

                    If pVal.ItemUID = "11" And pVal.BeforeAction = False Then
                        Me.LoadGrid(objForm.UniqueID, "Y")
                    End If

                    If pVal.ItemUID = "12" And pVal.BeforeAction = False Then
                        Me.LoadGrid(objForm.UniqueID, "N")
                    End If

                    If pVal.ItemUID = "7" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("3").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        ElseIf objForm.Items.Item("5").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            Me.GenerateProduction(objForm.UniqueID)
                        End If
                    End If

                    If pVal.ItemUID = "8" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("3").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        ElseIf objForm.Items.Item("5").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            Me.GenerateIssue(objForm.UniqueID)
                        End If
                    End If
                    If pVal.ItemUID = "9" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("3").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        ElseIf objForm.Items.Item("5").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            Me.GenerateReceipt(objForm.UniqueID)
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
            objMain.objApplication.StatusBar.SetText("Loading initiated...please wait a while...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objForm.Freeze(True)

            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oGrid = objForm.Items.Item("6").Specific
            oGrid.DataTable = oDt

            Try
                oGrid.DataTable.Clear()
            Catch ex As Exception
            End Try

            'Dim Query As String = " Select '" & Value & "' as ""Check"" ,  T0.""DocEntry"" As ""InvEntry"", T0.""DocNum"" As ""InvNo"" , TO_DATE(T0.""DocDate"") As ""InvDate"" , " &
            '                      " T1.""ItemCode"" , T1.""Dscription"" As ""ItemName"" , ""Quantity"" As ""Qty"" ,   IFNULL(""U_PGWPEN"",'') As ""Production""  , " &
            '                     " IFNULL(""U_PGWIEN"",'')  As ""Issue"" ,  Cast('' AS NVARCHAR(100)) As ""Receipt"" , Cast('' AS NVARCHAR(5000)) As ""Status""  , " &
            '                     " (Select ""BPLid"" From OITT T0 Inner Join OWHS T1 On T1.""WhsCode"" = T0.""ToWH""  Where ""Code"" = ""ItemCode"") As ""Branch"", ""LineNum"" From OINV T0 Inner Join " &
            '                      "  INV1 T1 On T1.""DocEntry"" = T0.""DocEntry"" Inner Join OITT T2 On T2.""Code"" = T1.""ItemCode"" " &
            '                      "  Where T2.""U_KITITEM"" = 'Yes' And T1.""DocDate"" Between '" & objForm.Items.Item("3").Specific.Value & "' And  '" & objForm.Items.Item("3").Specific.Value & "' "


            Dim Query As String = " Select '" & Value & "' as ""Check"" ,  T0.""DocEntry"" As ""InvEntry"", T0.""DocNum"" As ""InvNo"" , TO_DATE(T0.""DocDate"") As ""InvDate"" , " &
                                  " T1.""ItemCode"" , T1.""Dscription"" As ""ItemName"" , ""Quantity"" As ""Qty"" , T1.""WhsCode"" As ""Warehouse"" , IFNULL(""U_PGWPEN"",'') As ""Production""  , " &
                                  " IFNULL(""U_PGWIEN"",'')  As ""Issue"" , IFNULL(""U_PGWREN"",'')  As ""Receipt"" , Cast('' AS NVARCHAR(5000)) As ""Status""  , " &
                                  " ""BPLId"" As ""Branch"", ""LineNum"" , ( Select ""IssueMthd"" From OITM Where ""ItemCode""= T1.""ItemCode"" ) As ""IssueMethod"" From OINV T0 Inner Join " &
                                  "  INV1 T1 On T1.""DocEntry"" = T0.""DocEntry"" Inner Join OITT T2 On T2.""Code"" = T1.""ItemCode"" " &
                                  "  Where T2.""U_KITITEM"" = 'Yes' And T1.""DocDate"" Between '" & objForm.Items.Item("3").Specific.Value & "' And  '" & objForm.Items.Item("5").Specific.Value & "' "
            If objForm.Items.Item("Item_1").Specific.Value.Trim <> "" Then
                Query = Query & " And T1.""WhsCode""  = '" & objForm.Items.Item("Item_1").Specific.Value.Trim & "'"
            End If

            oGrid.DataTable.ExecuteQuery(Query)

            oGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item("InvEntry").Editable = False
            oGrid.Columns.Item("InvNo").Editable = False
            oGrid.Columns.Item("InvDate").Editable = False
            oGrid.Columns.Item("ItemCode").Editable = False
            oGrid.Columns.Item("ItemName").Editable = False
            oGrid.Columns.Item("Qty").Editable = False
            oGrid.Columns.Item("Warehouse").Editable = False
            oGrid.Columns.Item("Production").Editable = False
            oGrid.Columns.Item("Issue").Editable = False
            oGrid.Columns.Item("Receipt").Editable = False
            oGrid.Columns.Item("Status").Editable = False
            oGrid.Columns.Item("LineNum").Visible = False
            oGrid.Columns.Item("Branch").Visible = False
            oGrid.Columns.Item("IssueMethod").Visible = False
            If oGrid.DataTable.GetValue("ItemCode", 0) = "" Then
                objMain.objApplication.MessageBox("No records found")
                objForm.Freeze(False)
                Exit Sub
            End If
            Dim oEditTextCol As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("InvEntry")
            oEditTextCol.LinkedObjectType = 13
            Dim oEditTextCol1 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("Production")
            oEditTextCol1.LinkedObjectType = 202
            Dim oEditTextCol2 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("Issue")
            oEditTextCol2.LinkedObjectType = 60
            Dim oEditTextCol3 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("Receipt")
            oEditTextCol3.LinkedObjectType = 59
            Dim oEditTextCol4 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("ItemCode")
            oEditTextCol4.LinkedObjectType = 4

            Dim oEditTextCol5 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("Warehouse")
            oEditTextCol5.LinkedObjectType = 64

            For x As Integer = 0 To oGrid.Rows.Count - 1
                oGrid.CommonSetting.SetRowFontStyle(x + 1, SAPbouiCOM.BoFontStyle.fs_Bold)
                oGrid.CommonSetting.SetRowFontColor(x + 1, ColorTranslator.ToOle(Color.Black))
            Next

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        Finally
            oGrid.AutoResizeColumns()
            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
        End Try
    End Sub

    Private Sub GenerateProduction(ByVal FormUID As String)
        Dim oProgressPostProduction As SAPbouiCOM.ProgressBar = Nothing
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            Dim sProductionGUID As String
            oGrid = objForm.Items.Item("6").Specific
            oProgressPostProduction = objMain.objApplication.StatusBar.CreateProgressBar("Please Wait Issue docments generating....!", oGrid.DataTable.Rows.Count - 1, False)
            oProgressPostProduction.Text = "Please Wait production docments generating....!"
            Dim j1 As Integer = 1
            Dim oProductionOrders As SAPbobsCOM.ProductionOrders
            For j As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If oGrid.DataTable.GetValue("Check", j) = "Y" And oGrid.DataTable.GetValue("Production", j) = "" Then
                    oProductionOrders = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                    Dim GetSeries As String = " Select ""Series"" , ""SeriesName""   From NNM1 Where ""ObjectCode"" = '202' And ""Locked"" = 'N' And ""BPLId"" =  '" & oGrid.DataTable.GetValue("Branch", j) & "' And ""Indicator"" = '" & CDate(oGrid.DataTable.GetValue("InvDate", j)).Year & "' "
                    Dim oRsGetSeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetSeries.DoQuery(GetSeries)

                    oProductionOrders.Series = oRsGetSeries.Fields.Item("Series").Value
                    oProductionOrders.DueDate = DateTime.Now 'oGrid.DataTable.GetValue("InvDate", j)
                    oProductionOrders.ItemNo = oGrid.DataTable.GetValue("ItemCode", j)
                    oProductionOrders.PlannedQuantity = CDbl(oGrid.DataTable.GetValue("Qty", j))
                    oProductionOrders.Warehouse = oGrid.DataTable.GetValue("Warehouse", j)
                    sProductionGUID = System.Guid.NewGuid.ToString()
                    Try
                        oProductionOrders.UserFields.Fields.Item("U_GUID").Value = sProductionGUID
                    Catch ex As Exception
                    End Try

                    oProductionOrders.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotStandard

                    Dim GetBOMDetails As String = "Select * From OITT T0 Inner Join ITT1 T1 On T0.""Code""=T1.""Father"" Where T0.""Code""='" & oGrid.DataTable.GetValue("ItemCode", j) & "'"
                    Dim oRsGetBOMDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetBOMDetails.DoQuery(GetBOMDetails)
                    For i As Integer = 1 To oRsGetBOMDetails.RecordCount
                        oProductionOrders.Lines.ItemNo = oRsGetBOMDetails.Fields.Item("Code").Value
                        oProductionOrders.Lines.BaseQuantity = oRsGetBOMDetails.Fields.Item("Quantity").Value
                        oProductionOrders.Lines.Warehouse = oGrid.DataTable.GetValue("Warehouse", j)
                        oProductionOrders.Lines.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                        oProductionOrders.Lines.Add()
                        oRsGetBOMDetails.MoveNext()
                    Next

                  

                    Dim retVal As Integer = oProductionOrders.Add()
                    If retVal <> 0 Then
                        objMain.objApplication.StatusBar.SetText("Error:" & objMain.objCompany.GetLastErrorDescription.Replace("'", ""), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oGrid.DataTable.SetValue("Status", j, "Error:" & objMain.objCompany.GetLastErrorDescription.Replace("'", ""))
                    Else
                        Dim oRsGetProd As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim Production As String = "SELECT T0.""DocEntry"" FROM OWOR T0  Where ""U_GUID"" = '" & sProductionGUID & "' "
                        oRsGetProd.DoQuery(Production)

                        Dim INV1Update As String = "Update ""INV1"" Set  ""U_PGWPEN"" = '" & oRsGetProd.Fields.Item(0).Value.ToString & "'    Where ""DocEntry"" = '" & oGrid.DataTable.GetValue("InvEntry", j) & "' And ""LineNum"" = '" & oGrid.DataTable.GetValue("LineNum", j) & "' "
                        Dim oRsINV1Update As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsINV1Update.DoQuery(INV1Update)

                        oGrid.DataTable.SetValue("Production", j, oRsGetProd.Fields.Item(0).Value.ToString)
                        Dim Release As SAPbobsCOM.ProductionOrders = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                        If Release.GetByKey(oRsGetProd.Fields.Item(0).Value) = True Then
                            Release.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposReleased
                            If Release.Update() = 0 Then
                                oGrid.DataTable.SetValue("Status", j, "Production Order Released")
                            Else
                                Dim Err As String = " Production Order Release Error:'" & objMain.objCompany.GetLastErrorDescription.Replace("'", "").ToString().Substring(0, 200)
                                oGrid.DataTable.SetValue("Status", j, Err)
                                objMain.objApplication.StatusBar.SetText("Error:" & objMain.objCompany.GetLastErrorDescription.Replace("'", ""), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                        End If
                    End If
                End If
                oProgressPostProduction.Value = j1
                j1 = j1 + 1
            Next
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        Finally
            oGrid.AutoResizeColumns()
            If oProgressPostProduction IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgressPostProduction)
            End If
            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

        End Try
    End Sub

    Private Sub GenerateIssue(ByVal FormUID As String)
        Dim oProgressPostGoodsIssue As SAPbouiCOM.ProgressBar = Nothing
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim result As Integer = objMain.objApplication.MessageBox("You are going to post Issue from production. Do you want to continue?", 1, "Continue", "Cancel")
            If (result = 1) Then
                objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                Dim sIssueGUID As String
                oGrid = objForm.Items.Item("6").Specific
                oProgressPostGoodsIssue = objMain.objApplication.StatusBar.CreateProgressBar("Please Wait Issue docments generating....!", oGrid.DataTable.Rows.Count - 1, False)
                oProgressPostGoodsIssue.Text = "Please Wait Issue docments generating....!"
                Dim j1 As Integer = 1
                Dim oInventoryGenExit As SAPbobsCOM.Documents
                For j As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                    If oGrid.DataTable.GetValue("Check", j) = "Y" And oGrid.DataTable.GetValue("Production", j) <> "" And oGrid.DataTable.GetValue("Issue", j) = "" Then

                        'If oGrid.DataTable.GetValue("IssueMethod", j) = "B" Then
                        '    oGrid.DataTable.SetValue("Status", j, "Cannot Issue Backflush item...")
                        '    GoTo H
                        'End If

                        oInventoryGenExit = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
                        oInventoryGenExit.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
                        oInventoryGenExit.JournalMemo = "Issue from Production"
                        Dim GetSeries As String = " Select ""Series"" , ""SeriesName""   From NNM1 Where ""ObjectCode"" = '60' And ""Locked"" = 'N' And ""BPLId"" =  '" & oGrid.DataTable.GetValue("Branch", j) & "' And ""Indicator"" = '" & CDate(oGrid.DataTable.GetValue("InvDate", j)).Year & "' "
                        Dim oRsGetSeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetSeries.DoQuery(GetSeries)
                        oInventoryGenExit.Series = oRsGetSeries.Fields.Item("Series").Value
                        oInventoryGenExit.BPL_IDAssignedToInvoice = oGrid.DataTable.GetValue("Branch", j)

                        Dim ProdDetails As String = "Select ""LineNum"" from WOR1 A Where A.""DocEntry""='" & oGrid.DataTable.GetValue("Production", j) & "'"
                        Dim oRsProdDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsProdDetails.DoQuery(ProdDetails)

                        Dim LineCount As Integer = 0
                        For k As Integer = 1 To oRsProdDetails.RecordCount
                            If LineCount > 0 Then
                                oInventoryGenExit.Lines.Add()
                            End If
                            oInventoryGenExit.Lines.BaseEntry = oGrid.DataTable.GetValue("Production", j)
                            oInventoryGenExit.Lines.BaseLine = oRsProdDetails.Fields.Item(0).Value
                            oInventoryGenExit.Lines.BaseType = 202
                            LineCount = 1
                            oRsProdDetails.MoveNext()
                        Next
                        sIssueGUID = System.Guid.NewGuid.ToString()
                        Try
                            oInventoryGenExit.UserFields.Fields.Item("U_GUID").Value = sIssueGUID
                        Catch ex As Exception
                        End Try
                        Dim retVal As Integer = oInventoryGenExit.Add()
                        If retVal <> 0 Then
                            objMain.objApplication.StatusBar.SetText("Error:" & objMain.objCompany.GetLastErrorDescription.Replace("'", ""), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oGrid.DataTable.SetValue("Status", j, "Error:" & objMain.objCompany.GetLastErrorDescription.Replace("'", ""))
                        Else
                            Dim oRsGetIssue As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim Issue As String = "SELECT T0.""DocEntry"" FROM OIGE T0  Where ""U_GUID"" = '" & sIssueGUID & "' "
                            oRsGetIssue.DoQuery(Issue)
                            Dim INV1Update As String = "Update ""INV1"" Set  ""U_PGWIEN"" = '" & oRsGetIssue.Fields.Item(0).Value.ToString & "'    Where ""DocEntry"" = '" & oGrid.DataTable.GetValue("InvEntry", j) & "' And ""LineNum"" = '" & oGrid.DataTable.GetValue("LineNum", j) & "' "
                            Dim oRsINV1Update As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsINV1Update.DoQuery(INV1Update)
                            oGrid.DataTable.SetValue("Issue", j, oRsGetIssue.Fields.Item(0).Value.ToString)
                            oGrid.DataTable.SetValue("Status", j, "Issue From Production posted successfully")
                            objMain.objApplication.StatusBar.SetText("Issue From Production posted successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        End If
                    End If
                    oProgressPostGoodsIssue.Value = j1
                    j1 = j1 + 1
                Next
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        Finally
            oGrid.AutoResizeColumns()
            If oProgressPostGoodsIssue IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgressPostGoodsIssue)
            End If
            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
        End Try
    End Sub

    Private Sub GenerateReceipt(ByVal FormUID As String)
        Dim oProgressPostGoodsReceipt As SAPbouiCOM.ProgressBar = Nothing
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim result As Integer = objMain.objApplication.MessageBox("You are going to post Receipt from production. Do you want to continue?", 1, "Continue", "Cancel")
            If (result = 1) Then
                objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                Dim sReceiptGUID As String
                oGrid = objForm.Items.Item("6").Specific
                oProgressPostGoodsReceipt = objMain.objApplication.StatusBar.CreateProgressBar("Please Wait Receipt docments generating....!", oGrid.DataTable.Rows.Count - 1, False)
                oProgressPostGoodsReceipt.Text = "Please Wait Receipt docments generating....!"
                Dim j1 As Integer = 1
                Dim oInventoryGenEntry As SAPbobsCOM.Documents
                For j As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                    If oGrid.DataTable.GetValue("Check", j) = "Y" And oGrid.DataTable.GetValue("Production", j) <> "" And oGrid.DataTable.GetValue("Receipt", j) = "" Then
                        oInventoryGenEntry = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)

                        oInventoryGenEntry.JournalMemo = "Receipt from Production"
                        Dim GetSeries As String = " Select ""Series"" , ""SeriesName""   From NNM1 Where ""ObjectCode"" = '59' And ""Locked"" = 'N' And ""BPLId"" =  '" & oGrid.DataTable.GetValue("Branch", j) & "' And ""Indicator"" = '" & CDate(oGrid.DataTable.GetValue("InvDate", j)).Year & "' "
                        Dim oRsGetSeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetSeries.DoQuery(GetSeries)
                        oInventoryGenEntry.Series = oRsGetSeries.Fields.Item("Series").Value
                        oInventoryGenEntry.BPL_IDAssignedToInvoice = oGrid.DataTable.GetValue("Branch", j)
                        oInventoryGenEntry.Lines.BaseEntry = oGrid.DataTable.GetValue("Production", j)
                        oInventoryGenEntry.Lines.BaseType = 202
                        Dim GetItSrBtchDet As String = " SELECT " & IsNull & "(T0.""ManBtchNum"",'N')," & IsNull & "(T0.""ManSerNum"",'N')," & IsNull & "(T0.""MngMethod"",'N') " & _
                                                       " FROM OITM T0 WHERE T0.""ItemCode""='" & oGrid.DataTable.GetValue("ItemCode", j) & "'"
                        Dim oRsGetItSrBtchDet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetItSrBtchDet.DoQuery(GetItSrBtchDet)
                        If oRsGetItSrBtchDet.RecordCount > 0 Then
                            If CStr(oRsGetItSrBtchDet.Fields.Item(0).Value).Trim = "Y" Then
                                oInventoryGenEntry.Lines.BatchNumbers.BatchNumber = "SD - " & oGrid.DataTable.GetValue("InvEntry", j)
                                oInventoryGenEntry.Lines.BatchNumbers.Quantity = CDbl(oGrid.DataTable.GetValue("Qty", j))
                            End If
                        End If
                        oInventoryGenEntry.Lines.Add()
                        sReceiptGUID = System.Guid.NewGuid.ToString()
                        Try
                            oInventoryGenEntry.UserFields.Fields.Item("U_GUID").Value = sReceiptGUID
                        Catch ex As Exception
                        End Try
                        Dim retVal As Integer = oInventoryGenEntry.Add()
                        If retVal <> 0 Then
                            objMain.objApplication.StatusBar.SetText("Error:" & objMain.objCompany.GetLastErrorDescription.Replace("'", ""), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oGrid.DataTable.SetValue("Status", j, "Error:" & objMain.objCompany.GetLastErrorDescription.Replace("'", ""))
                        Else
                            Dim oRsGetIssue As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim Issue As String = "SELECT T0.""DocEntry"" FROM OIGN T0  Where ""U_GUID"" = '" & sReceiptGUID & "' "
                            oRsGetIssue.DoQuery(Issue)
                            Dim INV1Update As String = "Update ""INV1"" Set  ""U_PGWREN"" = '" & oRsGetIssue.Fields.Item(0).Value.ToString & "'    Where ""DocEntry"" = '" & oGrid.DataTable.GetValue("InvEntry", j) & "' And ""LineNum"" = '" & oGrid.DataTable.GetValue("LineNum", j) & "' "
                            Dim oRsINV1Update As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsINV1Update.DoQuery(INV1Update)
                            oGrid.DataTable.SetValue("Receipt", j, oRsGetIssue.Fields.Item(0).Value.ToString)
                            oGrid.DataTable.SetValue("Status", j, "Receipt posted successfully")
                            objMain.objApplication.StatusBar.SetText("Receipt posted successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        End If

                    End If
                Next
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        Finally
            oGrid.AutoResizeColumns()
            If oProgressPostGoodsReceipt IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgressPostGoodsReceipt)
            End If
            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
        End Try
    End Sub

End Class
