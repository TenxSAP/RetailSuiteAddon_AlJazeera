Public Class clsAutomaticProduction

#Region "       Declaration             "
    Public objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details1 As SAPbouiCOM.DBDataSource
    Dim objMatrix1 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim oDt, oDt1 As SAPbouiCOM.DataTable
    Dim objGrid As SAPbouiCOM.Grid
    Dim ProductionFlag As Boolean = False
    Dim IssueFlag As Boolean = False
    Dim oRs As SAPbobsCOM.Recordset
    Dim ReceiptFlag As Boolean = False
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("AutomaticProduction.xml", "ME_AP_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_AP_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            objForm.PaneLevel = 1
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_PR_C0")
            objMatrix1 = objForm.Items.Item("Mtx1").Specific
            objForm.Items.Item("Item_5").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oDt1 = objForm.DataSources.DataTables.Add("Test1")
            objForm.DataBrowser.BrowseBy = "DocNum"
            Me.CellsMasking(objForm.UniqueID)
            'objMatrix1.Columns.Item("RevQnty").Editable = False
            'objMatrix1.Columns.Item("RejQty").Editable = False
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ME_AP" And pVal.BeforeAction = False Then
                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "519" And pVal.BeforeAction = True Then 'Preview
                Try
                    Dim frs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'AutomaticProductionReport'")
                    'frs.DoQuery(Menuid)
                    If frs.RecordCount = 0 Then
                        objMain.objApplication.MessageBox("Report layout 'AutomaticProductionReport' not found. Please upload with same name and Code as parameter", 0, "OK")
                    Else
                        objMain.objApplication.ActivateMenuItem(frs.Fields.Item(0).Value.ToString())
                        Dim CrForm As SAPbouiCOM.Form
                        Dim oedt As SAPbouiCOM.EditText
                        CrForm = objMain.objApplication.Forms.ActiveForm
                        oedt = CrForm.Items.Item("1000003").Specific
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PR")
                        oedt.Value = oDBs_Head.GetValue("DocEntry", 0)
                        CrForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                Catch ex As Exception
                End Try
            End If
            If pVal.MenuUID = "Delete Row" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("ME_AP_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix1 = objForm.Items.Item("Mtx1").Specific
                For i As Integer = 1 To objMatrix1.VisualRowCount - 1
                    If objMatrix1.IsRowSelected(i) = True Then
                        objMatrix1.DeleteRow(i)
                    End If
                Next
                For i As Integer = 1 To objMatrix1.VisualRowCount
                    objMatrix1.Columns.Item("LineId").Cells.Item(i).Specific.string = i
                Next
                'If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub CellsMasking(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("DocDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("Load").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("Load").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("Load").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("Post").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("PostDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("PostDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("DNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("Receipt").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("Receipt").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("21").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("21").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("26").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "", Optional ByVal Series As Integer = 0)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_PR_C0")
            objMatrix1 = objForm.Items.Item("Mtx1").Specific
            objComboBox = objForm.Items.Item("20").Specific

            Dim GetSeriesQty As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'SBO_PRUDO' And ""Locked"" = 'N'"
            objMain.objUtilities.ComboBoxLoadValuesSeries(objForm.Items.Item("20").Specific, GetSeriesQty)
            objForm.Items.Item("20").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            Dim GetDflt As String = "Select ""DfltSeries"" From ONNM Where ""ObjectCode"" = 'SBO_PRUDO'"
            Dim oRsGetDflt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDflt.DoQuery(GetDflt)

            If Flag = "" Then
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNumSeries(objForm, "SBO_PRUDO", oRsGetDflt.Fields.Item(0).Value))
                oDBs_Head.SetValue("U_SERIES", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)
                oDBs_Head.SetValue("Series", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)
                oDBs_Head.SetValue("U_DocDate", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
                oDBs_Head.SetValue("U_PostDate", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
            Else
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNumSeries(objForm, "SBO_PRUDO", Series))
                oDBs_Head.SetValue("U_SERIES", oDBs_Head.Offset, Series)
                oDBs_Head.SetValue("Series", oDBs_Head.Offset, Series)
                oDBs_Head.SetValue("U_DocDate", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
                oDBs_Head.SetValue("U_PostDate", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)
            End If

            objMatrix1.Clear()
            oDBs_Details1.Clear()
            objMatrix1.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID)
            objComboBox = objForm.Items.Item("23").Specific
            Dim oRsUpdate2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim BPLId As String = " Select T0.""BPLId"",T1.""BPLName"",T1.""DflWhs"" From USR6 T0 Inner Join " &
                                                                 " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "'Order By T0.""BPLId""  "
            oRsUpdate2.DoQuery(BPLId)
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, BPLId)
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("23").DisplayDesc = True
            If objComboBox.ValidValues.Count <> 0 Then
                objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
            End If
            objMatrix1.AutoResizeColumns()
            objForm.Items.Item("DocNum").Visible = False
            objForm.Items.Item("st_cat").Visible = False
            objForm.Items.Item("Load").Visible = False
            'objForm.Items.Item("Post").Enabled = False
            'objForm.Items.Item("21").Enabled = False
            'objForm.Items.Item("Receipt").Enabled = False
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Sub ModeMask(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_PR_C0")
            objMatrix1 = objForm.Items.Item("Mtx1").Specific
            objMatrix1.Columns.Item("Qty").Editable = False
            'objMatrix1.Columns.Item("RevQnty").Editable = False
            'objMatrix1.Columns.Item("RejQty").Editable = False
            For i As Integer = 1 To objMatrix1.VisualRowCount
                'objMatrix1.Columns.Item("Qty").Editable = False
                'If objMatrix1.Columns.Item("Prod").Cells.Item(i).Specific.Value <> "" Then
                '    objMatrix1.CommonSetting.SetRowEditable(i, False)

                If objMatrix1.Columns.Item("Prod").Cells.Item(i).Specific.Value = "" Then
                    objMatrix1.CommonSetting.SetRowEditable(i, True)

                    objMatrix1.Columns.Item("IName").Editable = False
                    objMatrix1.Columns.Item("UOM").Editable = False
                    objMatrix1.Columns.Item("OPrice").Editable = False
                    objMatrix1.Columns.Item("Prod").Editable = False
                    objMatrix1.Columns.Item("RProd").Editable = False
                    objMatrix1.Columns.Item("Status").Editable = False
                    objMatrix1.Columns.Item("V_0").Editable = False
                    objMatrix1.Columns.Item("V_1").Editable = False
                    objMatrix1.Columns.Item("V_2").Editable = False
                    objMatrix1.Columns.Item("V_3").Editable = False
                    'objMatrix1.Columns.Item("RevQnty").Editable = False
                    'objMatrix1.Columns.Item("RejQty").Editable = False
                End If
              
                'If objMatrix1.Columns.Item("Prod").Cells.Item(i).Specific.Value <> "" And objMatrix1.Columns.Item("V_1").Cells.Item(i).Specific.Value <> "" Then
                '    objMatrix1.CommonSetting.SetRowEditable(i, True)
                '    objMatrix1.Columns.Item("RevQnty").Editable = True
                '    objMatrix1.Columns.Item("RejQty").Editable = True
                '    'objMatrix1.Columns.Item("IName").Editable = False
                '    'objMatrix1.Columns.Item("UOM").Editable = False
                '    'objMatrix1.Columns.Item("OPrice").Editable = False
                '    'objMatrix1.Columns.Item("Prod").Editable = False
                '    'objMatrix1.Columns.Item("RProd").Editable = False
                '    'objMatrix1.Columns.Item("Status").Editable = False
                '    'objMatrix1.Columns.Item("V_0").Editable = False
                '    'objMatrix1.Columns.Item("V_1").Editable = False
                '    'objMatrix1.Columns.Item("V_2").Editable = False
                '    'objMatrix1.Columns.Item("V_3").Editable = False
                '    'objMatrix1.Columns.Item("Qty").Editable = False
                'End If
            Next

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub SetNewLine(ByVal FormUID As String)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PR")
        oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_PR_C0")
        objMatrix1 = objForm.Items.Item("Mtx1").Specific

        objMatrix1.AddRow()
        oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
        oDBs_Details1.SetValue("U_ICode", oDBs_Details1.Offset, "")
        oDBs_Details1.SetValue("U_IName", oDBs_Details1.Offset, "")
        oDBs_Details1.SetValue("U_BarCode", oDBs_Details1.Offset, "")
        oDBs_Details1.SetValue("U_UOM", oDBs_Details1.Offset, "")
        oDBs_Details1.SetValue("U_OPrice", oDBs_Details1.Offset, "")
        oDBs_Details1.SetValue("U_Qty", oDBs_Details1.Offset, "")
        oDBs_Details1.SetValue("U_Status", oDBs_Details1.Offset, "")
        oDBs_Details1.SetValue("U_Prod", oDBs_Details1.Offset, "")
        oDBs_Details1.SetValue("U_IProd", oDBs_Details1.Offset, "")
        oDBs_Details1.SetValue("U_RProd", oDBs_Details1.Offset, "")
        oDBs_Details1.SetValue("U_ProdNo", oDBs_Details1.Offset, "")
        oDBs_Details1.SetValue("U_IProdNo", oDBs_Details1.Offset, "")
        oDBs_Details1.SetValue("U_RProdNo", oDBs_Details1.Offset, "")
        oDBs_Details1.SetValue("U_RevQnty", oDBs_Details1.Offset, "")
        oDBs_Details1.SetValue("U_RejQty", oDBs_Details1.Offset, "")
        oDBs_Details1.SetValue("U_Remak", oDBs_Details1.Offset, "")
        objMatrix1.SetLineData(objMatrix1.VisualRowCount)
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim Query As String = "SELECT MAX(""DocNum"") FROM ""@SBO_PR"""
                        Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        rs.DoQuery(Query)
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objForm.Items.Item("DNum").Specific.Value = rs.Fields.Item(0).Value
                        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                       
                        If Me.Validation(objForm.UniqueID) = False Then BubbleEvent = False
                    End If
                    
                    If pVal.ItemUID = "Item_6" And pVal.BeforeAction = False And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                        Me.ComponentsQuantity(objForm.UniqueID)
                    End If
                    If pVal.ItemUID = "Item_5" And pVal.BeforeAction = False Then
                        objForm.PaneLevel = 1
                        objMatrix1.AutoResizeColumns()
                    End If
                    'If pVal.ItemUID = "Load" And pVal.BeforeAction = False And (objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                    '    Me.LoadPrices(objForm.UniqueID)
                    'End If
                    If pVal.ItemUID = "Refresh" And pVal.BeforeAction = False Then
                        '----------Components---------------
                        objForm.Freeze(True)
                        'Dim Components As String = " SELECT T1.""ItemCode"", sum(T1.""PlannedQty"") as ""Required Qty"",(select ""OnHand"" from OITW where ""WhsCode""='" & objForm.Items.Item("26").Specific.Value & "' and ""ItemCode""=T1.""ItemCode"") as ""Qty in Bakery Whse"", " & _
                        '                           " (select ""OnHand"" from OITW where ""WhsCode""='" & objForm.Items.Item("26").Specific.Value & "' and ""ItemCode""=T1.""ItemCode"")-sum(T1.""PlannedQty"") AS ""Remaining Quanitity"" FROM OWOR T0  INNER JOIN WOR1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" " & _
                        '                           " Where T0.""DocEntry"" in (SELECT ""U_Prod"" FROM ""@SBO_PR"" P0 INNER JOIN ""@SBO_PR_C0"" P1 ON P0.""DocEntry""=P1.""DocEntry"" WHERE P0.""DocEntry""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "'  AND " & _
                        '                           " " & IsNull & "(P1.""U_Prod"",'')<>'') GROUP BY T1.""ItemCode"" "




                        '                   Dim Components As String = " SELECT T0.""ItemCode""as""Parent ItemCode"", T0.""ProdName"" as""Parent ItemName"",T1.""ItemCode"",T1.""ItemName"",T0.""DocNum""as ""Production Number""," &
                        '                           "sum(T1.""PlannedQty"") as ""Required Quantity"" " &
                        '" ,(select ""OnHand"" from OITW where ""WhsCode""='" & objForm.Items.Item("26").Specific.Value & "' and ""ItemCode""=T1.""ItemCode"") as ""Quantity in Whse"", " &
                        '                           " (select ""OnHand"" from OITW where ""WhsCode""='" & objForm.Items.Item("26").Specific.Value & "' and ""ItemCode""=T1.""ItemCode"")-sum(T1.""PlannedQty"") AS ""Remaining Quanitity""," &
                        '                             " '' as ""Components Cost"",'' as ""Production Cost"" , " &
                        '                            " (SELECT ""U_Remak"" FROM ""@SBO_PR"" P0 INNER JOIN  ""@SBO_PR_C0"" P1 ON P0.""DocEntry""=P1.""DocEntry"" WHERE P0.""DocEntry""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "' and P1.""U_ICode""=T0.""ItemCode"" ) as ""Remarks"" " &
                        '                            " FROM OWOR T0  INNER JOIN WOR1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" Where T0.""DocEntry"" in (SELECT ""U_Prod"" FROM ""@SBO_PR"" P0 INNER JOIN " &
                        '                            " ""@SBO_PR_C0"" P1 ON P0.""DocEntry""=P1.""DocEntry"" WHERE P0.""DocEntry""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "'  " &
                        '                            " AND " & IsNull & "(P1.""U_Prod"",'')<>'') GROUP BY T1.""ItemCode"" , T1.""ItemName"",T0.""DocNum"",T0.""ItemCode"", T0.""ProdName"" order by T1.""ItemCode"" "


                        'Dim Components As String = "Call ""SBO_COMPONENTDETAILS""('" & oDBs_Head.GetValue("DocEntry", oDBs_Head.Offset) & "')"

                        Dim Components As String = " SELECT T1.""ItemCode"", sum(T1.""PlannedQty"") as ""Required Qty"",(select ""OnHand"" from OITW where ""WhsCode""='" & objForm.Items.Item("26").Specific.Value & "' and ""ItemCode""=T1.""ItemCode"") as ""Qty in Bakery Whse"", " &
                                                   " (select ""OnHand"" from OITW where ""WhsCode""='" & objForm.Items.Item("26").Specific.Value & "' and ""ItemCode""=T1.""ItemCode"")-sum(T1.""PlannedQty"") AS ""Remaining Quanitity"" FROM OWOR T0  INNER JOIN WOR1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
                                                   " Where T0.""DocEntry"" in (SELECT ""U_Prod"" FROM ""@SBO_PR"" P0 INNER JOIN ""@SBO_PR_C0"" P1 ON P0.""DocEntry""=P1.""DocEntry"" WHERE P0.""DocEntry""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "'  AND " &
                                                   " " & IsNull & "(P1.""U_Prod"",'')<>'') GROUP BY T1.""ItemCode"" "

                        oDt1 = objForm.DataSources.DataTables.Item("Test1")
                        objGrid = objForm.Items.Item("Grid").Specific
                        oDt1.ExecuteQuery(Components)
                        objGrid.DataTable = oDt1
                        objGrid.AutoResizeColumns()
                        objGrid.Columns.Item(0).Editable = False
                        objGrid.Columns.Item(1).Editable = False
                        objGrid.Columns.Item(2).Editable = False
                        objGrid.Columns.Item(3).Editable = False
                        'objGrid.Columns.Item(4).Editable = False
                        'objGrid.Columns.Item(5).Editable = False
                        'If objGrid.Rows.Count = 0 Then
                        '    objMain.objApplication.MessageBox("No records found")
                        '    Exit Sub
                        'End If
                        'Dim oEditTextCol As SAPbouiCOM.EditTextColumn
                        'oEditTextCol = objGrid.Columns.Item("Parent ItemCode")
                        'oEditTextCol.LinkedObjectType = 66
                        'Dim oEditTextCol2 As SAPbouiCOM.EditTextColumn
                        'oEditTextCol2 = objGrid.Columns.Item("ItemCode")
                        'oEditTextCol2.LinkedObjectType = 4
                        'Dim oEditTextCol3 As SAPbouiCOM.EditTextColumn
                        'oEditTextCol3 = objGrid.Columns.Item("Production Entry")
                        'oEditTextCol3.LinkedObjectType = 202
                        objForm.Freeze(False)
                        '-----------------------------------
                    End If

                    If pVal.ItemUID = "Post" And pVal.BeforeAction = False And objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Dim result As Integer = objMain.objApplication.MessageBox("You are going to post production orders. Do you want to continue?", 1, "Continue", "Cancel")
                        If (result = 1) Then
                            objForm.Freeze(True)
                            Me.CreateProductionOrder(objForm.UniqueID)
                            objForm.Freeze(False)
                        End If
                    ElseIf pVal.ItemUID = "Post" And pVal.BeforeAction = False And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                        objMain.objApplication.StatusBar.SetText("Posting should be in OK Mode only", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                    If pVal.ItemUID = "Receipt" And pVal.BeforeAction = False And objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        objForm.Freeze(True)
                        Me.CreateReceipt(objForm.UniqueID)
                        objForm.Freeze(False)
                    End If

                    If pVal.ItemUID = "21" And pVal.BeforeAction = False And objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        objForm.Freeze(True)
                        Me.CreateIssueForProduction(objForm.UniqueID)
                        objForm.Freeze(False)

                    End If
                    If pVal.BeforeAction = False And objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        'If pVal.BeforeAction = False And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        'objForm.Items.Item("Post").Enabled = False
                        'objForm.Items.Item("21").Enabled = False
                        'objForm.Items.Item("Receipt").Enabled = False
                        objForm.Freeze(True)
                        Me.EnablingButtons(objForm.UniqueID)
                        objForm.Freeze(False)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PR")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_PR_C0")
                    objMatrix1 = objForm.Items.Item("Mtx1").Specific

                    If pVal.ItemUID = "Mtx1" And pVal.ColUID = "Qty" And pVal.BeforeAction = True Then
                        If (objMatrix1.Columns.Item("Prod").Cells.Item(pVal.Row).Specific.Value <> "") Then
                            BubbleEvent = False
                            objMain.objApplication.StatusBar.SetText("Quantity modification is not permitted in Line No. " & pVal.Row & "...!!.... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        End If
                    End If
                    If pVal.ItemUID = "Mtx1" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        objForm.Freeze(True)
                        Me.EnablingButtons(objForm.UniqueID)
                        objForm.Freeze(False)
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    'If pVal.ItemUID = "1" And pVal.Before_Action = "True" Then
                    '    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PR")
                    '    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_PR_C0")
                    '    objMatrix1 = objForm.Items.Item("Mtx1").Specific
                    '    Try
                    '        objForm.Freeze(True)
                    '        For i As Integer = 1 To objMatrix1.RowCount
                    '            If objMatrix1.Columns.Item("ICode").Cells.Item(i).Specific.Value = "" Then
                    '                objMatrix1.DeleteRow(i)
                    '                i = 0

                    '            End If
                    '        Next

                    '    Catch
                    '    Finally
                    '        objForm.Freeze(False)
                    '    End Try
                    '    For i As Integer = 1 To objMatrix1.RowCount
                    '        objMatrix1.Columns.Item("LineId").Cells.Item(i).Specific.Value = i
                    '        'Me.SetNewLine(objForm.UniqueID)
                    '    Next
                    '    Me.SetNewLine(objForm.UniqueID)
                    'End If
                    'Me.EnablingButtons(objForm.UniqueID)
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PR")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_PR_C0")
                    objMatrix1 = objForm.Items.Item("Mtx1").Specific
                    If pVal.ItemUID = "Mtx1" And pVal.ColUID = "BarCode" Then
                        If objMatrix1.Columns.Item("BarCode").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            Dim GetItemCode As String = " Select ""ItemCode"" , (Select A.""ItemName"" From OITM A Where  A.""ItemCode"" = B.""ItemCode"") ""ItemName"" " &
                                                        " From OBCD B Where ""BcdCode"" = '" & objMatrix1.Columns.Item("BarCode").Cells.Item(pVal.Row).Specific.Value & "'"
                            Dim oRsGetItemCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetItemCode.DoQuery(GetItemCode)

                            oDBs_Details1.SetValue("LineID", oDBs_Details1.Offset, pVal.Row)
                            oDBs_Details1.SetValue("U_ICode", oDBs_Details1.Offset, oRsGetItemCode.Fields.Item("ItemCode").Value)
                            oDBs_Details1.SetValue("U_IName", oDBs_Details1.Offset, oRsGetItemCode.Fields.Item("ItemName").Value)
                            oDBs_Details1.SetValue("U_BarCode", oDBs_Details1.Offset, objMatrix1.Columns.Item("BarCode").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_UOM", oDBs_Details1.Offset, objMatrix1.Columns.Item("UOM").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_OPrice", oDBs_Details1.Offset, objMatrix1.Columns.Item("OPrice").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_Prod", oDBs_Details1.Offset, objMatrix1.Columns.Item("Prod").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_Qty", oDBs_Details1.Offset, objMatrix1.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_RProd", oDBs_Details1.Offset, objMatrix1.Columns.Item("RProd").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_Status", oDBs_Details1.Offset, objMatrix1.Columns.Item("Status").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_IProd", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_ProdNo", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_RProdNo", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_IProdNo", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_Remak", oDBs_Details1.Offset, objMatrix1.Columns.Item("Remak").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_RevQnty", oDBs_Details1.Offset, objMatrix1.Columns.Item("RevQnty").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_RejQty", oDBs_Details1.Offset, objMatrix1.Columns.Item("RejQty").Cells.Item(pVal.Row).Specific.Value)


                            objMatrix1.SetLineData(pVal.Row)


                            If pVal.Row = objMatrix1.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID)
                            End If
                        End If
                    End If
                    If pVal.ItemUID = "Mtx1" And pVal.ColUID = "Qty" Then
                        If objMatrix1.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            Dim quantity As Double = objMatrix1.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value
                            objMatrix1.Columns.Item("RevQnty").Cells.Item(pVal.Row).Specific.Value = quantity

                        End If
                    End If
                    If pVal.ItemUID = "Mtx1" And pVal.ColUID = "RevQnty" And pVal.BeforeAction = False Then
                        If objMatrix1.Columns.Item("RevQnty").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            Dim Receivedquantity As Double = objMatrix1.Columns.Item("RevQnty").Cells.Item(pVal.Row).Specific.Value
                            Dim Plannedquantity As Double = objMatrix1.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value
                            Dim totalquantity As Double = Plannedquantity - Receivedquantity
                            If totalquantity < 0 Then
                                objMatrix1.Columns.Item("RejQty").Cells.Item(pVal.Row).Specific.Value = 0
                            Else
                                objMatrix1.Columns.Item("RejQty").Cells.Item(pVal.Row).Specific.Value = totalquantity
                            End If

                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PR")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_PR_C0")
                    objMatrix1 = objForm.Items.Item("Mtx1").Specific
                    If pVal.ItemUID = "20" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("20").Specific.Selected.Value <> "" Then
                            Dim series As Integer = objForm.Items.Item("20").Specific.Selected.Value
                            SetDefault(objForm.UniqueID, "CobmoSelect", series)
                        End If
                    End If
                    If pVal.ItemUID = "23" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("23").Specific.Selected.Value <> "" Then
                            Dim oRsUpdate2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            Dim BPLId As String = " Select T0.""BPLId"",T1.""BPLName"",T1.""DflWhs"" From USR6 T0 Inner Join " &
                                                                                 " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where  T1.""BPLName""='" & objForm.Items.Item("23").Specific.Selected.Description & "'and  T0.""UserCode"" = '" & objMain.objCompany.UserName & "'"
                            oRsUpdate2.DoQuery(BPLId)
                            objForm.Items.Item("26").Specific.Value = oRsUpdate2.Fields.Item("DflWhs").Value
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PR")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_PR_C0")
                    objMatrix1 = objForm.Items.Item("Mtx1").Specific
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    'If oCFL.UniqueID = "CFL_ITEM" Then
                    '    Me.CFLFilter(objForm.UniqueID, "CFL_ITEM")
                    'End If
                    If oCFL.UniqueID = "CFL_Barcode" Then
                        Me.CFLFilterBarcode(objForm.UniqueID, "CFL_Barcode")
                    End If
                    If oCFL.UniqueID = "CFL_WHSCD" And pVal.BeforeAction = True Then
                        Me.CFLFilter(FormUID, "CFL_WHSCD")
                    End If
                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                        'If oCFL.UniqueID = "CFL_ITEM" Then
                        '    oDBs_Details1.SetValue("LineID", oDBs_Details1.Offset, pVal.Row)
                        '    oDBs_Details1.SetValue("U_ICode", oDBs_Details1.Offset, oDT.GetValue("ItemCode", 0))
                        '    oDBs_Details1.SetValue("U_IName", oDBs_Details1.Offset, oDT.GetValue("ItemName", 0))
                        '    Dim GetBarCode As String = " SELECT Top 1 ""BcdCode"" FROM OBCD WHERE  " & _
                        '                              "  ""ItemCode"" = '" & oDT.GetValue("ItemCode", 0) & "' ORDER BY ""BcdEntry"" ASC "
                        '    Dim oRsGetBarCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '    oRsGetBarCode.DoQuery(GetBarCode)
                        '    oDBs_Details1.SetValue("U_BarCode", oDBs_Details1.Offset, oRsGetBarCode.Fields.Item("BcdCode").Value)
                        '    oDBs_Details1.SetValue("U_UOM", oDBs_Details1.Offset, objMatrix1.Columns.Item("UOM").Cells.Item(pVal.Row).Specific.Value)
                        '    oDBs_Details1.SetValue("U_OPrice", oDBs_Details1.Offset, objMatrix1.Columns.Item("OPrice").Cells.Item(pVal.Row).Specific.Value)
                        '    oDBs_Details1.SetValue("U_Prod", oDBs_Details1.Offset, "")
                        '    oDBs_Details1.SetValue("U_Qty", oDBs_Details1.Offset, "")
                        '    oDBs_Details1.SetValue("U_RProd", oDBs_Details1.Offset, "")
                        '    oDBs_Details1.SetValue("U_Status", oDBs_Details1.Offset, "")
                        '    oDBs_Details1.SetValue("U_IProd", oDBs_Details1.Offset, "")
                        '    oDBs_Details1.SetValue("U_ProdNo", oDBs_Details1.Offset, "")
                        '    oDBs_Details1.SetValue("U_RProdNo", oDBs_Details1.Offset, "")
                        '    oDBs_Details1.SetValue("U_IProdNo", oDBs_Details1.Offset, "")
                        '    objMatrix1.SetLineData(pVal.Row)
                        '    If pVal.Row = objMatrix1.VisualRowCount Then
                        '        SetNewLine(objForm.UniqueID)
                        '    End If

                        'End If
                        If oCFL.UniqueID = "CFL_BOM" Then
                            'If objMatrix1.Columns.Item("Prod").Cells.Item(pVal.Row).Specific.Value = "" Then
                            oDBs_Details1.SetValue("LineID", oDBs_Details1.Offset, pVal.Row)
                            oDBs_Details1.SetValue("U_ICode", oDBs_Details1.Offset, oDT.GetValue("Code", 0))
                            oDBs_Details1.SetValue("U_IName", oDBs_Details1.Offset, oDT.GetValue("Name", 0))
                            Dim GetBarCode As String = " SELECT Top 1 ""BcdCode"" FROM OBCD WHERE  " &
                                                  "  ""ItemCode"" = '" & oDT.GetValue("Code", 0) & "' ORDER BY ""BcdEntry"" ASC "
                            Dim oRsGetBarCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetBarCode.DoQuery(GetBarCode)
                            oDBs_Details1.SetValue("U_BarCode", oDBs_Details1.Offset, oRsGetBarCode.Fields.Item("BcdCode").Value)
                            oDBs_Details1.SetValue("U_UOM", oDBs_Details1.Offset, objMatrix1.Columns.Item("UOM").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_OPrice", oDBs_Details1.Offset, objMatrix1.Columns.Item("OPrice").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_Prod", oDBs_Details1.Offset, objMatrix1.Columns.Item("Prod").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_Qty", oDBs_Details1.Offset, objMatrix1.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_RProd", oDBs_Details1.Offset, objMatrix1.Columns.Item("RProd").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_Status", oDBs_Details1.Offset, objMatrix1.Columns.Item("Status").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_IProd", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_ProdNo", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_RProdNo", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_IProdNo", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_Remak", oDBs_Details1.Offset, objMatrix1.Columns.Item("Remak").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_RevQnty", oDBs_Details1.Offset, objMatrix1.Columns.Item("RevQnty").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_RejQty", oDBs_Details1.Offset, objMatrix1.Columns.Item("RejQty").Cells.Item(pVal.Row).Specific.Value)


                            objMatrix1.SetLineData(pVal.Row)
                            If pVal.Row = objMatrix1.VisualRowCount Then
                                SetNewLine(objForm.UniqueID)
                                Me.ModeMask(objForm.UniqueID)
                            End If
                        End If
                        'End If
                        If oCFL.UniqueID = "CFL_Barcode" Then
                            'If objMatrix1.Columns.Item("Prod").Cells.Item(pVal.Row).Specific.Value = "" Then
                            oDBs_Details1.SetValue("LineID", oDBs_Details1.Offset, pVal.Row)
                            oDBs_Details1.SetValue("U_BarCode", oDBs_Details1.Offset, oDT.GetValue("BcdCode", 0))
                            oDBs_Details1.SetValue("U_ICode", oDBs_Details1.Offset, oDT.GetValue("ItemCode", 0))
                            Dim Name As String = " SELECT ""Name""FROM OITT WHERE  " &
                                                  "  ""Code"" = '" & oDT.GetValue("ItemCode", 0) & "' "
                            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRs.DoQuery(Name)
                            oDBs_Details1.SetValue("U_IName", oDBs_Details1.Offset, oRs.Fields.Item("Name").Value)
                            oDBs_Details1.SetValue("U_UOM", oDBs_Details1.Offset, objMatrix1.Columns.Item("UOM").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_OPrice", oDBs_Details1.Offset, objMatrix1.Columns.Item("OPrice").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_Prod", oDBs_Details1.Offset, objMatrix1.Columns.Item("Prod").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_Qty", oDBs_Details1.Offset, objMatrix1.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_RProd", oDBs_Details1.Offset, objMatrix1.Columns.Item("RProd").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_Status", oDBs_Details1.Offset, objMatrix1.Columns.Item("Status").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_IProd", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_ProdNo", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_RProdNo", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_IProdNo", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_Remak", oDBs_Details1.Offset, objMatrix1.Columns.Item("Remak").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_RevQnty", oDBs_Details1.Offset, objMatrix1.Columns.Item("RevQnty").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_RejQty", oDBs_Details1.Offset, objMatrix1.Columns.Item("RejQty").Cells.Item(pVal.Row).Specific.Value)

                            objMatrix1.SetLineData(pVal.Row)
                            If pVal.Row = objMatrix1.VisualRowCount Then
                                SetNewLine(objForm.UniqueID)
                                Me.ModeMask(objForm.UniqueID)
                            End If
                            'End If
                        End If

                        If oCFL.UniqueID = "CFL_WHSCD" Then
                            oDBs_Head.SetValue("U_WHSCODE", oDBs_Head.Offset, oDT.GetValue("WhsCode", 0))
                        End If
                    End If
                    'End If
                    'Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                    '    If pVal.ItemUID = "Grid" And pVal.ColUID = "Production Number" And pVal.Before_Action = False Then
                    '        objGrid = objForm.Items.Item("Grid").Specific
                    '        Dim code As String = objGrid.DataTable.GetValue("Production Number", pVal.Row)
                    '        Dim objNewForm As SAPbouiCOM.Form
                    '        objMain.Pro.CreateForm()
                    '        objNewForm = objMain.objApplication.Forms.ActiveForm
                    '        objNewForm.Freeze(True)
                    '        objNewForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    '        objNewForm.Items.Item("Code").Specific.Value = code
                    '        objNewForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    '        objNewForm.Freeze(False)
                    '    End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub CreateProductionOrder(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_PR_C0")
            objMatrix1 = objForm.Items.Item("Mtx1").Specific
            For i As Integer = 1 To objMatrix1.RowCount
                If objMatrix1.Columns.Item("ICode").Cells.Item(i).Specific.Value.ToString().Trim() <> "" And objMatrix1.Columns.Item("Prod").Cells.Item(i).Specific.Value.ToString().Trim() = "" Then

                    objMain.objApplication.StatusBar.SetText("Production orders posting, Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Dim Prod As SAPbobsCOM.ProductionOrders = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                    Prod.ItemNo = objMatrix1.Columns.Item("ICode").Cells.Item(i).Specific.Value.ToString().Trim()
                    Dim PostDate As String = objForm.Items.Item("PostDate").Specific.Value.ToString().Trim()
                    If PostDate = "" Then
                        objMain.objApplication.SetStatusBarMessage("Please enter ther the Posting Date", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        Exit Sub
                    End If
                    PostDate = PostDate.Insert(4, "/")
                    PostDate = PostDate.Insert(7, "/")
                    Dim PDate As Date = CDate(PostDate)
                    Prod.DueDate = PDate
                    'Prod.Warehouse = oDBs_Head.GetValue("U_WHSCODE", 0)
                    Prod.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotStandard
                    Prod.PlannedQuantity = CDbl(objMatrix1.Columns.Item("Qty").Cells.Item(i).Specific.Value.ToString().Trim())
                    Dim retVal As Integer = Prod.Add()
                    If retVal <> 0 Then
                        objMain.objApplication.StatusBar.SetText("Error:" & objMain.objCompany.GetLastErrorDescription.Replace("'", ""), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Dim oRsUpdate2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim sError As String = objMain.objCompany.GetLastErrorDescription.Replace("'", "").Replace("'", "")
                        Dim Update2 As String = " UPDATE  T1 SET T1.""U_Status""='Production Order Creation Error: " & sError & "' From ""@SBO_PR_C0"" T1 WHERE T1.""DocEntry""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "' and T1.""LineId""='" & i & "'"
                        oRsUpdate2.DoQuery(Update2)
                    Else
                        Dim oRsGetProd As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim Production As String = "SELECT max(T0.""DocEntry"") FROM OWOR T0"
                        oRsGetProd.DoQuery(Production)
                        Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim Update As String = "UPDATE T1 SET T1.""U_Prod""=(SELECT max(T0.""DocEntry"") FROM OWOR T0) , T1.""U_ProdNo""=(SELECT ""DocNum"" From OWOR  Where ""DocEntry"" = (SELECT max(T0.""DocEntry"") FROM OWOR T0)) ,    T1.""U_Status""='Production Order Posted' From ""@SBO_PR_C0"" T1  WHERE T1.""DocEntry""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "' and T1.""LineId""='" & i & "'"
                        oRsUpdate.DoQuery(Update)
                        '--------Releaseing----------------
                        Dim Release As SAPbobsCOM.ProductionOrders = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                        If Release.GetByKey(oRsGetProd.Fields.Item(0).Value) = True Then
                            Release.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposReleased
                            If Release.Update() = 0 Then
                                Dim oRsUpdate2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim Update2 As String = " UPDATE  T1 SET T1.""U_Status""='Production Order Released' From ""@SBO_PR_C0"" T1 WHERE T1.""DocEntry""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "' and T1.""LineId""='" & i & "'"
                                oRsUpdate2.DoQuery(Update2)
                            Else
                                Dim oRsUpdate3 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim Update3 As String = " UPDATE  T1 SET T1.""U_Status""='Production Order Release Error:'" & objMain.objCompany.GetLastErrorDescription.Replace("'", "").ToString().Substring(0, 200) & " From ""@SBO_PR_C0"" T1 WHERE T1.""DocEntry""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "' and T1.""LineId""='" & i & "'"
                                oRsUpdate3.DoQuery(Update3)
                                objMain.objApplication.StatusBar.SetText("Error:" & objMain.objCompany.GetLastErrorDescription.Replace("'", ""), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                        End If
                    End If
                End If

            Next
            objMain.objApplication.ActivateMenuItem("1304")
            objMatrix1.AutoResizeColumns()
            objMain.objApplication.StatusBar.SetText("Production orders posted successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'objForm.Items.Item("Post").Enabled = False
            'objForm.Items.Item("21").Enabled = True
            '----------Components---------------
            objForm.Freeze(True)
            Dim Components As String = " SELECT T1.""ItemCode"", sum(T1.""PlannedQty"") as ""Required Qty"",(select ""OnHand"" from OITW where ""WhsCode""='" & objForm.Items.Item("26").Specific.Value & "' and ""ItemCode""=T1.""ItemCode"") as ""Qty in Bakery Whse"" " & _
                                       " ,(select ""OnHand"" from OITW where ""WhsCode""='" & objForm.Items.Item("26").Specific.Value & "' and ""ItemCode""=T1.""ItemCode"")-sum(T1.""PlannedQty"") AS ""Remaining Quanitity"" FROM OWOR T0  INNER JOIN WOR1 T1 ON " & _
                                       " T0.""DocEntry"" = T1.""DocEntry"" Where T0.""DocEntry"" in (SELECT ""U_Prod"" FROM ""@SBO_PR"" P0 INNER JOIN ""@SBO_PR_C0"" P1 ON P0.""DocEntry""=P1.""DocEntry"" WHERE " & _
                                       " P0.""DocEntry""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "'  AND " & IsNull & "(P1.""U_Prod"",'')<>'') GROUP BY T1.""ItemCode"" "
            oDt1 = objForm.DataSources.DataTables.Item("Test1")
            objGrid = objForm.Items.Item("Grid").Specific
            oDt1.ExecuteQuery(Components)
            objGrid.DataTable = oDt1
            objGrid.AutoResizeColumns()
            objGrid.Columns.Item(0).Editable = False
            objGrid.Columns.Item(1).Editable = False
            objGrid.Columns.Item(2).Editable = False
            objGrid.Columns.Item(3).Editable = False
            objForm.Freeze(False)
            '-----------------------------------
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub CreateIssueForProduction(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_PR_C0")
            objMatrix1 = objForm.Items.Item("Mtx1").Specific

            Dim result As Integer = objMain.objApplication.MessageBox("You are going to post IssueReceipt from production. Do you want to continue?", 1, "Continue", "Cancel")
            If (result = 1) Then
                objMatrix1 = objForm.Items.Item("Mtx1").Specific
                For i As Integer = 1 To objMatrix1.VisualRowCount
                    If objMatrix1.Columns.Item("ICode").Cells.Item(i).Specific.Value.ToString().Trim() <> "" And objMatrix1.Columns.Item("Prod").Cells.Item(i).Specific.Value.ToString().Trim() <> "" And objMatrix1.Columns.Item("RProd").Cells.Item(i).Specific.Value.ToString().Trim() = "" Then
                        objMain.objApplication.StatusBar.SetText("Please Wait Issue is posting.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        Dim ProdDetails As String = "Select ""LineNum"" from WOR1 A Where A.""DocEntry""='" & objMatrix1.Columns.Item("Prod").Cells.Item(i).Specific.Value.ToString().Trim() & "'"
                        Dim oRsProdDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsProdDetails.DoQuery(ProdDetails)

                        Dim Issue As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
                        Issue.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
                        Issue.JournalMemo = "Issue from Production"
                        Issue.Comments = "Automatic Production"
                        Issue.Reference2 = oDBs_Head.GetValue("DocNum", 0)
                        objComboBox = objForm.Items.Item("23").Specific
                        Dim GetSeries As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = '60' And ""Locked"" = 'N'  And ""BPLId"" = '" & objComboBox.Value & "'  Order By ""Series"" Desc "
                        Dim oRsGetSeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetSeries.DoQuery(GetSeries)
                        If oRsGetSeries.RecordCount > 0 Then
                            Issue.Series = oRsGetSeries.Fields.Item("Series").Value
                        End If
                        Issue.BPL_IDAssignedToInvoice = objComboBox.Value
                        Dim LineCount As Integer = 0
                        For j As Integer = 1 To oRsProdDetails.RecordCount
                            If LineCount > 0 Then
                                Issue.Lines.Add()
                            End If
                            Issue.Lines.BaseEntry = objMatrix1.Columns.Item("Prod").Cells.Item(i).Specific.Value.ToString().Trim()
                            Issue.Lines.BaseLine = oRsProdDetails.Fields.Item(0).Value
                            Issue.Lines.BaseType = 202
                            LineCount = 1
                            oRsProdDetails.MoveNext()
                        Next
                        Dim retVal As Integer = Issue.Add()
                        '-------------
                        If retVal <> 0 Then
                            Dim oRsUpdate3 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim sError As String = objMain.objCompany.GetLastErrorDescription.Replace("'", "")
                            Dim Update3 As String = " UPDATE  T1 SET T1.""U_Status""='Issue From Production Error: " & sError & "' From ""@SBO_PR_C0"" T1  WHERE T1.""DocEntry""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "' and T1.""LineId""='" & i & "'"
                            oRsUpdate3.DoQuery(Update3)
                            objMain.objApplication.StatusBar.SetText("Error:" & objMain.objCompany.GetLastErrorDescription.Replace("'", ""), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim Update As String = " UPDATE T1 SET T1.""U_IProd""=(SELECT max(T0.""DocEntry"") FROM OIGE T0),T1.""U_IProdNo""=(SELECT ""DocNum"" FROM OIGE Where ""DocEntry"" = (SELECT max(T0.""DocEntry"") FROM OIGE T0)) ,     T1.""U_Status""='Issue From Production Completed' From ""@SBO_PR_C0"" T1 WHERE T1.""DocEntry""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "' and T1.""LineId""='" & i & "'"
                            oRsUpdate.DoQuery(Update)
                            objMain.objApplication.StatusBar.SetText("Issue From Production posted successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        End If
                    End If
                    'objMatrix1.CommonSetting.SetRowEditable(i, True)
                    'objMatrix1.Columns.Item("RevQnty").Editable = True
                    'objMatrix1.Columns.Item("RejQty").Editable = True
                    'objMatrix1.Columns.Item("IName").Editable = False
                    'objMatrix1.Columns.Item("UOM").Editable = False
                    'objMatrix1.Columns.Item("OPrice").Editable = False
                    'objMatrix1.Columns.Item("Prod").Editable = False
                    'objMatrix1.Columns.Item("RProd").Editable = False
                    'objMatrix1.Columns.Item("Status").Editable = False
                    'objMatrix1.Columns.Item("V_0").Editable = False
                    'objMatrix1.Columns.Item("V_1").Editable = False
                    'objMatrix1.Columns.Item("V_2").Editable = False
                    'objMatrix1.Columns.Item("V_3").Editable = False
                Next
               
                objMain.objApplication.ActivateMenuItem("1304")
                objMatrix1.AutoResizeColumns()
            End If
            'objForm.Items.Item("21").Enabled = False
            'objForm.Items.Item("Receipt").Enabled = True
            'Else
            'objMain.objApplication.StatusBar.SetText("Please check component quantities.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub CreateReceipt(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_PR_C0")
            objMatrix1 = objForm.Items.Item("Mtx1").Specific

            'Dim Qty As String = " Select count(*) from (SELECT  (select ""OnHand"" from OITW where ""WhsCode""='201' and ""ItemCode""=T1.""ItemCode"")-sum(T1.""PlannedQty"") AS ""Diff"" " & _
            '                    " FROM OWOR T0  INNER JOIN WOR1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" " & _
            '                    " Where T0.""DocEntry"" in (SELECT P1.""U_Prod"" FROM ""@SBO_PR_C0"" P1 WHERE P1.""DocEntry""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "' AND " & IsNull & " (P1.""U_Prod"",'')<>'') " & _
            '                    " GROUP BY T1.""ItemCode"")a Where a.""Diff"" <0 "
            'Dim oRsQty As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRsQty.DoQuery(Qty)
            'If oRsQty.Fields.Item(0).Value = 0 Then
            Dim result As Integer = objMain.objApplication.MessageBox(" You are going to post Receipt from production. Do you want to continue?", 1, "Continue", "Cancel")
            If (result = 1) Then
                objMatrix1 = objForm.Items.Item("Mtx1").Specific
                For i As Integer = 1 To objMatrix1.VisualRowCount
                    If objMatrix1.Columns.Item("ICode").Cells.Item(i).Specific.Value.ToString().Trim() <> "" And objMatrix1.Columns.Item("Prod").Cells.Item(i).Specific.Value.ToString().Trim() <> "" And objMatrix1.Columns.Item("V_1").Cells.Item(i).Specific.Value.ToString().Trim() <> "" And objMatrix1.Columns.Item("V_2").Cells.Item(i).Specific.Value.ToString().Trim() <> "" And objMatrix1.Columns.Item("RProd").Cells.Item(i).Specific.Value.ToString().Trim() = "" Then
                        objMain.objApplication.StatusBar.SetText("Please Wait Receipt is posting.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        '----------
                        Dim Receipt As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
                        Receipt.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
                        Receipt.JournalMemo = "Receipt from Production"
                        Receipt.Comments = "Automatic Production"
                        Receipt.Reference2 = oDBs_Head.GetValue("DocNum", 0)
                        objComboBox = objForm.Items.Item("23").Specific
                        Receipt.BPL_IDAssignedToInvoice = objComboBox.Value

                        Dim GetSeries As String = " Select ""Series"" , ""SeriesName""   From NNM1 Where ""ObjectCode"" = '59' And ""Locked"" = 'N' And ""BPLId"" =  '" & objComboBox.Value & "' Order By ""Series"" Desc "
                        Dim oRsGetSeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetSeries.DoQuery(GetSeries)

                        Receipt.Series = oRsGetSeries.Fields.Item("Series").Value
                      
                        Receipt.Lines.BaseEntry = objMatrix1.Columns.Item("Prod").Cells.Item(i).Specific.Value.ToString().Trim()
                        Receipt.Lines.BaseType = 202

                        Dim GetItSrBtchDet As String = " SELECT " & IsNull & "(T0.""ManBtchNum"",'N')," & IsNull & "(T0.""ManSerNum"",'N')," & IsNull & "(T0.""MngMethod"",'N') " & _
                                                       " FROM OITM T0 WHERE T0.""ItemCode""='" & objMatrix1.Columns.Item("ICode").Cells.Item(i).Specific.Value.ToString().Trim() & "'"
                        Dim oRsGetItSrBtchDet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetItSrBtchDet.DoQuery(GetItSrBtchDet)
                        If oRsGetItSrBtchDet.RecordCount > 0 Then
                            If CStr(oRsGetItSrBtchDet.Fields.Item(0).Value).Trim = "Y" Then
                                Receipt.Lines.BatchNumbers.BatchNumber = "SD - " & objForm.Items.Item("DNum").Specific.Value.ToString().Trim()
                                Receipt.Lines.BatchNumbers.Quantity = CDbl(objMatrix1.Columns.Item("RevQnty").Cells.Item(i).Specific.Value.ToString().Trim())
                                'Receipt.Lines.BatchNumbers.Add()
                            End If
                        End If
                        Receipt.Lines.Quantity = objMatrix1.Columns.Item("RevQnty").Cells.Item(i).Specific.Value

                        Receipt.Lines.Add()

                        '----------Rejected------------------------
                        If objMatrix1.Columns.Item("RejQty").Cells.Item(i).Specific.Value <> 0 Then
                            Dim LineCount As Integer = 0
                            
                            Receipt.Lines.BaseEntry = objMatrix1.Columns.Item("Prod").Cells.Item(i).Specific.Value.ToString().Trim()
                            Receipt.Lines.BaseType = 202
                            Receipt.Lines.Quantity = objMatrix1.Columns.Item("RejQty").Cells.Item(i).Specific.Value
                            Receipt.Lines.TransactionType = SAPbobsCOM.BoTransactionTypeEnum.botrntReject
                            Receipt.Lines.Add()

                        End If
                        '------------------------------------------
                        Dim retVal As Integer = Receipt.Add()

                        '-------------
                        If retVal <> 0 Then
                            Dim oRsUpdate3 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            Dim Update3 As String = " UPDATE  T1 SET T1.""U_Status""='Receipt From Production Error: " & objMain.objCompany.GetLastErrorDescription.Replace("'", "") & "' From ""@SBO_PR_C0"" T1  WHERE T1.""DocEntry""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "' and T1.""LineId""='" & i & "'"
                            oRsUpdate3.DoQuery(Update3)
                            objMain.objApplication.StatusBar.SetText("Error:" & objMain.objCompany.GetLastErrorDescription.Replace("'", ""), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim Update As String = " UPDATE T1 SET T1.""U_RProd""=(SELECT max(T0.""DocEntry"") FROM OIGN T0), T1.""U_RProdNo""=(SELECT ""DocNum""  FROM OIGN Where ""DocEntry"" = (SELECT max(T0.""DocEntry"") FROM OIGN T0)) , T1.""U_Status""='Receipt From Production Completed' From ""@SBO_PR_C0"" T1 WHERE T1.""DocEntry""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "' and T1.""LineId""='" & i & "'"
                            oRsUpdate.DoQuery(Update)
                        End If
                    End If
                Next
                objMain.objApplication.StatusBar.SetText("Receipt From Production posted successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                objMain.objApplication.ActivateMenuItem("1304")
                objMatrix1.AutoResizeColumns()
            End If
            'Else
            'objMain.objApplication.StatusBar.SetText("Please check component quantities.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub LoadPrices(ByVal FormUID As String)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PR")
        oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_PR_C0")
        Dim PrgBar As SAPbouiCOM.ProgressBar
        Try
            Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx1").Specific
            PrgBar = objMain.objApplication.StatusBar.CreateProgressBar("Please wait....", objMatrix.VisualRowCount, False)
            For i As Integer = 1 To objMatrix.VisualRowCount
                PrgBar.Value = i
                PrgBar.Text = "Please wait default prices are loading."
                Dim ItemCode As String = objMatrix.Columns.Item("ICode").Cells.Item(i).Specific.Value.ToString().Trim()

                Dim GetBranchPriceList As String = "Select ""U_PricelistNo"" From ""OBPL"" Where ""BPLId""= '" & objForm.Items.Item("23").Specific.Value.Trim & "'"
                Dim oRsGetBranchPriceList As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetBranchPriceList.DoQuery(GetBranchPriceList)

                Dim PriceList As String = oRsGetBranchPriceList.Fields.Item("U_PricelistNo").Value


                Dim BarCode As String = objMatrix.Columns.Item("BarCode").Cells.Item(i).Specific.Value.ToString().Trim()
                Dim Production As String = objMatrix.Columns.Item("Prod").Cells.Item(i).Specific.Value.ToString().Trim()
                If ItemCode <> "" And BarCode = "" And Production = "" Then
                    Dim DefQuery As String = "select ""Price"" from ITM1 where ""ItemCode""='" & ItemCode & "'  and  ""PriceList"" ='" & PriceList & "'"
                    Dim oRsDefQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsDefQuery.DoQuery(DefQuery)
                    Dim DefPrice As Double = CDbl(oRsDefQuery.Fields.Item("Price").Value)

                    'Dim GPriceQuery As String = "select ""Price"" from ITM1 where ""ItemCode""='" & ItemCode & "'  and  ""PriceList"" ='3'"
                    'Dim oRsGPriceQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'oRsGPriceQuery.DoQuery(GPriceQuery)
                    'Dim GPrice As Double = CDbl(oRsGPriceQuery.Fields.Item("Price").Value)

                    Dim DetQuery As String = "select ""InvntryUom"" AS ""UOM"",""CodeBars"",""ItemName"" from OITM where ""ItemCode""='" & ItemCode & "' "
                    Dim oRsDetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsDetQuery.DoQuery(DetQuery)

                    oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, i)
                    oDBs_Details1.SetValue("U_ICode", oDBs_Details1.Offset, ItemCode)
                    oDBs_Details1.SetValue("U_IName", oDBs_Details1.Offset, oRsDetQuery.Fields.Item("ItemName").Value)
                    oDBs_Details1.SetValue("U_BarCode", oDBs_Details1.Offset, oRsDetQuery.Fields.Item("CodeBars").Value)
                    oDBs_Details1.SetValue("U_UOM", oDBs_Details1.Offset, oRsDetQuery.Fields.Item("UOM").Value)
                    oDBs_Details1.SetValue("U_OPrice", oDBs_Details1.Offset, DefPrice)
                    oDBs_Details1.SetValue("U_Qty", oDBs_Details1.Offset, objMatrix.Columns.Item("Qty").Cells.Item(i).Specific.Value)
                    oDBs_Details1.SetValue("U_Prod", oDBs_Details1.Offset, objMatrix.Columns.Item("Prod").Cells.Item(i).Specific.Value)
                    oDBs_Details1.SetValue("U_IProd", oDBs_Details1.Offset, objMatrix.Columns.Item("V_1").Cells.Item(i).Specific.Value)
                    oDBs_Details1.SetValue("U_RProd", oDBs_Details1.Offset, objMatrix.Columns.Item("RProd").Cells.Item(i).Specific.Value)
                    oDBs_Details1.SetValue("U_Status", oDBs_Details1.Offset, objMatrix.Columns.Item("Status").Cells.Item(i).Specific.Value)
                    objMatrix.SetLineData(i)

                ElseIf BarCode <> "" And Production = "" Then
                    Dim OPrice As Double = 0
                    Dim GPrice As Double = 0
                    Dim Query As String = "SELECT (select ""UomCode"" from OUOM Where ""UomEntry""=T0.""UomEntry"") AS ""UOM"",T0.""UomEntry"", T0.""ItemCode"" FROM OBCD T0 WHERE T0.""BcdCode""='" & BarCode & "' "
                    Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsQuery.DoQuery(Query)
                    ItemCode = oRsQuery.Fields.Item("ItemCode").Value
                    '--------------------------------------------
                    'Dim PriceQueryG As String = "select Top 1 ""Price"" from ITM9 where ""ItemCode""='" & ItemCode & "' and ""PriceList""='3' and ""UomEntry""='" & oRsQuery.Fields.Item("UomEntry").Value & "'"
                    'Dim oRsPriceQueryG As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'oRsPriceQueryG.DoQuery(PriceQueryG)
                    'If oRsPriceQueryG.RecordCount = 0 Then
                    '    Dim DetQueryG As String = "select T0.""ItemName"",T1.""Price"" from OITM T0 INNER JOIN ITM1 T1 ON T0.""ItemCode""=T1.""ItemCode"" where T0.""ItemCode""='" & ItemCode & "' and T1.""PriceList"" ='3' "
                    '    Dim oRsDetQueryG As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '    oRsDetQueryG.DoQuery(DetQueryG)
                    '    GPrice = CDbl(oRsDetQueryG.Fields.Item("Price").Value)
                    'Else
                    '    GPrice = CDbl(oRsPriceQueryG.Fields.Item("Price").Value)
                    'End If
                    '------------------------------------------------
                    Dim ItemName As String = "Select ""ItemName"" From OITM Where ""ItemCode""='" & ItemCode & "'"
                    Dim oRsItemName As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsItemName.DoQuery(ItemName)
                    '---------------------------------------------------------------------
                    Dim PriceQuery As String = "select Top 1 ""Price"" from ITM9 where ""ItemCode""='" & ItemCode & "' and ""PriceList""='" & PriceList & "' and ""UomEntry""='" & oRsQuery.Fields.Item("UomEntry").Value & "'"
                    Dim oRsPriceQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsPriceQuery.DoQuery(PriceQuery)
                    If oRsPriceQuery.RecordCount = 0 Then
                        Dim DefQuery As String = "select ""Price"" from ITM1 where ""ItemCode""='" & ItemCode & "'  and  ""PriceList"" ='" & PriceList & "'"
                        Dim oRsDefQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsDefQuery.DoQuery(DefQuery)
                        OPrice = oRsDefQuery.Fields.Item(0).Value
                    Else
                        OPrice = oRsPriceQuery.Fields.Item(0).Value
                    End If

                    oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, i)
                    oDBs_Details1.SetValue("U_ICode", oDBs_Details1.Offset, ItemCode)
                    oDBs_Details1.SetValue("U_IName", oDBs_Details1.Offset, oRsItemName.Fields.Item("ItemName").Value)
                    oDBs_Details1.SetValue("U_BarCode", oDBs_Details1.Offset, BarCode)
                    oDBs_Details1.SetValue("U_UOM", oDBs_Details1.Offset, oRsQuery.Fields.Item("UOM").Value)
                    oDBs_Details1.SetValue("U_OPrice", oDBs_Details1.Offset, OPrice)
                    'oDBs_Details1.SetValue("U_Qty", oDBs_Details1.Offset, "")
                    'oDBs_Details1.SetValue("U_Prod", oDBs_Details1.Offset, "")
                    'oDBs_Details1.SetValue("U_RProd", oDBs_Details1.Offset, "")
                    'oDBs_Details1.SetValue("U_Status", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_Qty", oDBs_Details1.Offset, objMatrix.Columns.Item("Qty").Cells.Item(i).Specific.Value)
                    oDBs_Details1.SetValue("U_Prod", oDBs_Details1.Offset, objMatrix.Columns.Item("Prod").Cells.Item(i).Specific.Value)
                    oDBs_Details1.SetValue("U_IProd", oDBs_Details1.Offset, objMatrix.Columns.Item("V_1").Cells.Item(i).Specific.Value)
                    oDBs_Details1.SetValue("U_RProd", oDBs_Details1.Offset, objMatrix.Columns.Item("RProd").Cells.Item(i).Specific.Value)
                    oDBs_Details1.SetValue("U_Status", oDBs_Details1.Offset, objMatrix.Columns.Item("Status").Cells.Item(i).Specific.Value)
                    objMatrix.SetLineData(i)
                End If
            Next
            PrgBar.Stop()
        Catch ex As Exception
            PrgBar.Stop()
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            GC.Collect()
        End Try
    End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("ME_AP_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Me.ComponentsQuantity(objForm.UniqueID)
                        Me.EnablingButtons(objForm.UniqueID)
                        Me.ModeMask(objForm.UniqueID)

                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Me.EnablingButtons(objForm.UniqueID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        '   Me.EnablingButtons(objForm.UniqueID)
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    'Sub CFLFilter(ByVal FormUID As String, ByVal CFL_ID As String)
    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)
    '        Dim GetFreeItems As String = "Then Select ""Code"" From OITT Where " & IsNull & "(""U_KITITEM"",'No') <> 'Yes'   "
    '        Dim oRsGetFreeItems As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oRsGetFreeItems.DoQuery(GetFreeItems)

    '        Dim oConditions As SAPbouiCOM.Conditions
    '        Dim oCondition As SAPbouiCOM.Condition
    '        Dim oChooseFromList As SAPbouiCOM.ChooseFromList
    '        Dim emptyCon As New SAPbouiCOM.Conditions
    '        oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
    '        oChooseFromList.SetConditions(emptyCon)
    '        oConditions = oChooseFromList.GetConditions()
    '        If oRsGetFreeItems.RecordCount > 0 Then
    '            oCondition = oConditions.Add()
    '            oCondition.Alias = "ItemCode"
    '            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '            oCondition.CondVal = oRsGetFreeItems.Fields.Item(0).Value
    '            oChooseFromList.SetConditions(oConditions)
    '            oRsGetFreeItems.MoveNext()
    '            For i As Integer = 1 To oRsGetFreeItems.RecordCount - 1
    '                oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
    '                oCondition = oConditions.Add()
    '                oCondition.Alias = "ItemCode"
    '                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '                oCondition.CondVal = oRsGetFreeItems.Fields.Item(0).Value
    '                oChooseFromList.SetConditions(oConditions)
    '                oRsGetFreeItems.MoveNext()
    '            Next
    '        Else
    '            oCondition = oConditions.Add()
    '            oCondition.Alias = "ItemCode"
    '            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '            oCondition.CondVal = ""
    '            oChooseFromList.SetConditions(oConditions)
    '        End If
    '    Catch ex As Exception

    '    End Try
    'End Sub
    Sub CFLFilter(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim Selection As String = "select ""WhsCode"",""WhsName"",""BPLid"" from OWHS where ""BPLid""='" & objForm.Items.Item("23").Specific.Value & "'"
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
                oCondition.Alias = "BPLid"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = oRs.Fields.Item("BPLid").Value
                oChooseFromList.SetConditions(oConditions)

            End If
        Catch ex As Exception

        End Try
    End Sub
    Sub CFLFilterBarcode(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim GetSelection As String = "SELECT T0.""ItemCode"",T1.""Name"" FROM OBCD T0  inner join OITT T1  on T0.""ItemCode""= T1.""Code"""
            Dim oRsGetSelectionItems As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetSelectionItems.DoQuery(GetSelection)

            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            If oRsGetSelectionItems.RecordCount > 0 Then
                oCondition = oConditions.Add()
                oCondition.Alias = "ItemCode"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = oRsGetSelectionItems.Fields.Item(0).Value
                oChooseFromList.SetConditions(oConditions)
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

        End Try
    End Sub
    Sub EnablingButtons(ByVal FormUID As String)
        Try
            objForm.Items.Item("Post").Enabled = False
            objForm.Items.Item("21").Enabled = False
            objForm.Items.Item("Receipt").Enabled = False

            objMatrix1 = objForm.Items.Item("Mtx1").Specific
            For i As Integer = 1 To objMatrix1.VisualRowCount - 1
                If objMatrix1.Columns.Item("Prod").Cells.Item(i).Specific.Value.ToString().Trim() <> "" Then
                    ProductionFlag = 1
                Else
                    ProductionFlag = 0

                End If
            Next
            If ProductionFlag = True Then
                objForm.Items.Item("Post").Enabled = False
                objForm.Items.Item("21").Enabled = True
            Else
                objForm.Items.Item("Post").Enabled = True
                Exit Sub
            End If
            For i As Integer = 1 To objMatrix1.VisualRowCount - 1
                If objMatrix1.Columns.Item("V_1").Cells.Item(i).Specific.Value.ToString().Trim() <> "" And objMatrix1.Columns.Item("Prod").Cells.Item(i).Specific.Value.ToString().Trim() <> "" Then
                    IssueFlag = 1
                Else
                    IssueFlag = 0
                End If
            Next
            If IssueFlag = True Then
                objForm.Items.Item("21").Enabled = False
                objForm.Items.Item("Receipt").Enabled = True
            Else
                objForm.Items.Item("21").Enabled = True
                Exit Sub
            End If
            For i As Integer = 1 To objMatrix1.VisualRowCount - 1
                If objMatrix1.Columns.Item("RProd").Cells.Item(i).Specific.Value.ToString().Trim() <> "" And objMatrix1.Columns.Item("V_1").Cells.Item(i).Specific.Value.ToString().Trim() <> "" And objMatrix1.Columns.Item("Prod").Cells.Item(i).Specific.Value.ToString().Trim() <> "" Then
                    ReceiptFlag = 1
                Else
                    ReceiptFlag = 0
                End If
            Next
            If ReceiptFlag = True Then
                objForm.Items.Item("Receipt").Enabled = False
                'objForm.Items.Item("Post").Enabled = True
            Else
                objForm.Items.Item("Receipt").Enabled = True
                Exit Sub
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub ComponentsQuantity(ByVal FormUID As String)

        '     Dim Components As String = " SELECT T0.""ItemCode""as""Parent ItemCode"", T0.""ProdName"" as""Parent ItemName"",T1.""ItemCode"",T1.""ItemName"",T0.""DocNum""as ""Production Number""," &
        '                               "sum(T1.""PlannedQty"") as ""Required Quantity"" " &
        '    " ,(select ""OnHand"" from OITW where ""WhsCode""='" & objForm.Items.Item("26").Specific.Value & "' and ""ItemCode""=T1.""ItemCode"") as ""Quantity in Whse"", " &
        '                               " (select ""OnHand"" from OITW where ""WhsCode""='" & objForm.Items.Item("26").Specific.Value & "' and ""ItemCode""=T1.""ItemCode"")-sum(T1.""PlannedQty"") AS ""Remaining Quanitity""," &
        '                                 "(SELECT T11.""Price"" FROM ""OIGE"" T10 INNER JOIN ""IGE1"" T11 ON T10.""DocEntry"" = T11.""DocEntry"" WHERE T10.""JrnlMemo"" ='Issue from Production'and T11.""ItemCode""=T1.""ItemCode"" " &
        '"and T10.""DocNum"" in (SELECT ""U_IProdNo"" FROM ""@SBO_PR"" P0 INNER JOIN  ""@SBO_PR_C0"" P1 ON P0.""DocEntry""=P1.""DocEntry""  WHERE P0.""DocEntry""= '" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "'  AND IFNULL(P1.""U_IProd"",'')<>'')) " &
        '" as ""Components Cost"",'' as ""Production Cost"" , " &
        '                                " (SELECT ""U_Remak"" FROM ""@SBO_PR"" P0 INNER JOIN  ""@SBO_PR_C0"" P1 ON P0.""DocEntry""=P1.""DocEntry"" WHERE P0.""DocEntry""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "' and P1.""U_ICode""=T0.""ItemCode"" ) as ""Remarks"" " &
        '                                " FROM OWOR T0  INNER JOIN WOR1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" Where T0.""DocEntry"" in (SELECT ""U_Prod"" FROM ""@SBO_PR"" P0 INNER JOIN " &
        '                                " ""@SBO_PR_C0"" P1 ON P0.""DocEntry""=P1.""DocEntry"" WHERE P0.""DocEntry""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "'  " &
        '                                " AND " & IsNull & "(P1.""U_Prod"",'')<>'') GROUP BY T1.""ItemCode"" , T1.""ItemName"",T0.""DocNum"",T0.""ItemCode"", T0.""ProdName"" order by T1.""ItemCode"" "


        Dim Components As String = " SELECT T1.""ItemCode"", sum(T1.""PlannedQty"") as ""Required Qty"",(select ""OnHand"" from OITW where ""WhsCode""='" & objForm.Items.Item("26").Specific.Value & "' and ""ItemCode""=T1.""ItemCode"") as ""Qty in Bakery Whse"", " &
                                                   " (select ""OnHand"" from OITW where ""WhsCode""='" & objForm.Items.Item("26").Specific.Value & "' and ""ItemCode""=T1.""ItemCode"")-sum(T1.""PlannedQty"") AS ""Remaining Quanitity"" FROM OWOR T0  INNER JOIN WOR1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
                                                   " Where T0.""DocEntry"" in (SELECT ""U_Prod"" FROM ""@SBO_PR"" P0 INNER JOIN ""@SBO_PR_C0"" P1 ON P0.""DocEntry""=P1.""DocEntry"" WHERE P0.""DocEntry""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "'  AND " &
                                                   " " & IsNull & "(P1.""U_Prod"",'')<>'') GROUP BY T1.""ItemCode"" "
        'Dim Qry As String = "Call ""SBO_COMPONENTDETAILS""('" & oDBs_Head.GetValue("DocEntry", oDBs_Head.Offset) & "')"
        oDt1 = objForm.DataSources.DataTables.Item("Test1")
        objGrid = objForm.Items.Item("Grid").Specific
        oDt1.ExecuteQuery(Components)
        objGrid.DataTable = oDt1
        objGrid.AutoResizeColumns()
        objGrid.Columns.Item(0).Editable = False
        objGrid.Columns.Item(1).Editable = False
        objGrid.Columns.Item(2).Editable = False
        objGrid.Columns.Item(3).Editable = False
        'objGrid.Columns.Item(4).Editable = False
        'objGrid.Columns.Item(5).Editable = False
        'If objGrid.Rows.Count = 0 Then
        '    objMain.objApplication.MessageBox("No records found")
        '    Exit Sub
        'End If
        'Dim oEditTextCol As SAPbouiCOM.EditTextColumn
        'oEditTextCol = objGrid.Columns.Item("Parent ItemCode")
        'oEditTextCol.LinkedObjectType = 66
        'Dim oEditTextCol2 As SAPbouiCOM.EditTextColumn
        'oEditTextCol2 = objGrid.Columns.Item("ItemCode")
        'oEditTextCol2.LinkedObjectType = 4
        'Dim oEditTextCol3 As SAPbouiCOM.EditTextColumn
        'oEditTextCol3 = objGrid.Columns.Item("Production Entry")
        'oEditTextCol3.LinkedObjectType = 202

        objForm.Freeze(False)
    End Sub
    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx1").Specific
            For i As Integer = 1 To objMatrix.VisualRowCount - 1
                Dim qty As Double = objMatrix.Columns.Item("Qty").Cells.Item(i).Specific.Value
                If qty = 0 Then
                    '        Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.Trim
                    '        If (ItemCode <> "" And objMatrix.Columns.Item("U_LPValidate").Cells.Item(i).Specific.Value.Trim = "N") Then
                    objMain.objApplication.StatusBar.SetText("Quantity should be greater than Zero in Line No. " & i & "...!!.... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                    '        End If
                End If
            Next
            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function
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
                        objMatrix1 = objForm.Items.Item("Mtx1").Specific
                        If eventInfo.ItemUID = "Mtx1" And eventInfo.ColUID = "LineId" Then
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
                        ElseIf eventInfo.ItemUID = "Mtx1" And objMatrix1.RowCount <= 1 Then
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
                        If eventInfo.ItemUID <> "Mtx1" Then
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
