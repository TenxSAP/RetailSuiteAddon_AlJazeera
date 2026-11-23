Imports System.ComponentModel

Public Class clsPostMultipleAPInvoices

#Region "        Declaration        "
    Dim objForm As SAPbouiCOM.Form
    Dim oGrid As SAPbouiCOM.Grid
    Dim oDt As SAPbouiCOM.DataTable
    Dim oEditText As SAPbouiCOM.EditText
    Dim objComboBox As SAPbouiCOM.ComboBox
    Private WithEvents BackgroundWorker_GenerateDocumets As New System.ComponentModel.BackgroundWorker
    Private lastSortedColumn As String = ""
    Private lastSortDirection As SAPbouiCOM.BoGridSortType = SAPbouiCOM.BoGridSortType.gst_Ascending

#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("PostMultipleAPInvoices.xml", "ME_PMAPINV_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_PMAPINV_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.DataSources.UserDataSources.Add("U_FROMDT", SAPbouiCOM.BoDataType.dt_DATE)
            objForm.DataSources.UserDataSources.Add("U_TODT", SAPbouiCOM.BoDataType.dt_DATE)
            objForm.DataSources.UserDataSources.Add("U_BRANCH", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            objForm.DataSources.UserDataSources.Add("U_SERIES", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            objForm.DataSources.UserDataSources.Add("U_SUPCD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_SUPNM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_TTLAMT", SAPbouiCOM.BoDataType.dt_PRICE)
            objForm.DataSources.UserDataSources.Add("U_GRPO", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            oEditText = objForm.Items.Item("19").Specific
            oEditText.DataBind.SetBound(True, "", "U_TTLAMT")
            objForm.Items.Item("19").AffectsFormMode = False
            objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            oEditText = objForm.Items.Item("5").Specific
            oEditText.DataBind.SetBound(True, "", "U_FROMDT")
            objForm.Items.Item("5").AffectsFormMode = False

            oEditText = objForm.Items.Item("7").Specific
            oEditText.DataBind.SetBound(True, "", "U_TODT")
            objForm.Items.Item("7").AffectsFormMode = False

            oEditText = objForm.Items.Item("14").Specific
            oEditText.DataBind.SetBound(True, "", "U_SUPCD")
            oEditText.ChooseFromListUID = "CFL_SPCD"
            oEditText.ChooseFromListAlias = "CardCode"
            objForm.Items.Item("14").AffectsFormMode = False

            oEditText = objForm.Items.Item("16").Specific
            oEditText.DataBind.SetBound(True, "", "U_SUPNM")
            oEditText.ChooseFromListUID = "CFL_SPNM"
            oEditText.ChooseFromListAlias = "CardName"
            objForm.Items.Item("16").AffectsFormMode = False

            oEditText = objForm.Items.Item("Item_1").Specific
            oEditText.DataBind.SetBound(True, "", "U_GRPO")
            oEditText.ChooseFromListUID = "CFL_GRPO"
            oEditText.ChooseFromListAlias = "DocEntry"
            objForm.Items.Item("Item_1").AffectsFormMode = False

            objComboBox = objForm.Items.Item("11").Specific
            objComboBox.DataBind.SetBound(True, "", "U_BRANCH")
            objForm.Items.Item("11").AffectsFormMode = False
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, "Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""  ")
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            If objComboBox.ValidValues.Count <> 0 Then
                objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
            End If

            Dim objBranchCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("11").Specific
            objComboBox = objForm.Items.Item("13").Specific
            objComboBox.DataBind.SetBound(True, "", "U_SERIES")
            objForm.Items.Item("13").AffectsFormMode = False
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = '18' And ""Locked"" = 'N'  And ""BPLId"" = '" & objBranchCombo.Value & "' Order By ""Series"" Desc ")
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            Try
                'Dim GetDflt As String = "Select ""DfltSeries"" From ONNM Where ""ObjectCode"" = '18'"
                'Dim oRsGetDflt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oRsGetDflt.DoQuery(GetDflt)
                If objComboBox.ValidValues.Count <> 0 Then
                    'objComboBox.Select(oRsGetDflt.Fields.Item(0).Value.ToString, SAPbouiCOM.BoSearchKey.psk_ByValue)
                    objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
                End If
            Catch ex As Exception
            End Try

            oDt = objForm.DataSources.DataTables.Add("dt1")
            oDt = objForm.DataSources.DataTables.Item("dt1")

            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ME_PMAPINV" And pVal.BeforeAction = False Then
                Me.CreateForm()
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "11" And pVal.BeforeAction = False Then
                        Dim objBranchCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("11").Specific
                        objComboBox = objForm.Items.Item("13").Specific
                        objComboBox.DataBind.SetBound(True, "", "U_SERIES")
                        objForm.Items.Item("13").AffectsFormMode = False
                        objMain.objUtilities.ComboBoxLoadValues(objComboBox, "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = '18' And ""Locked"" = 'N'  And ""BPLId"" = '" & objBranchCombo.Value & "' Order By ""Series"" Desc ")
                        objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                        Try
                            If objComboBox.ValidValues.Count <> 0 Then
                                objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
                            End If
                        Catch ex As Exception
                        End Try
                    End If
                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                    'If pVal.ItemUID = "9" AndAlso pVal.BeforeAction = True Then
                    '    Try
                    '        Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(pVal.FormUID)
                    '        Dim oGrid As SAPbouiCOM.Grid = oForm.Items.Item("9").Specific
                    '        Dim clickedColumn As String = pVal.ColUID
                    '        oGrid.Columns.Item(clickedColumn).TitleObject.Sortable = True
                    '        ' Skip checkbox column or invalid columns
                    '        If oGrid.Columns.Item(clickedColumn).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox Then Exit Select

                    '        ' Determine sort direction
                    '        Dim sortDirection As SAPbouiCOM.BoGridSortType

                    '        If clickedColumn = lastSortedColumn Then
                    '            If lastSortDirection = SAPbouiCOM.BoGridSortType.gst_Ascending Then
                    '                sortDirection = SAPbouiCOM.BoGridSortType.gst_Descending
                    '            Else
                    '                sortDirection = SAPbouiCOM.BoGridSortType.gst_Ascending
                    '            End If
                    '        Else
                    '            sortDirection = SAPbouiCOM.BoGridSortType.gst_Ascending
                    '        End If

                    '        ' Perform sort
                    '        oGrid.Columns.Item(clickedColumn).TitleObject.Sort(sortDirection)

                    '        ' Save last sorted info
                    '        lastSortedColumn = clickedColumn
                    '        lastSortDirection = sortDirection

                    '    Catch ex As Exception
                    '        objMain.objApplication.StatusBar.SetText("Error sorting column: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    End Try
                    'End If

                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "9" And pVal.ColUID = "Check" And pVal.BeforeAction = False Then
                        Try
                            Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(pVal.FormUID)

                            Dim Grid As SAPbouiCOM.Grid = oForm.Items.Item("9").Specific
                            Dim rowIndex As Integer = pVal.Row
                            Dim isChecked As String = (Grid.DataTable.GetValue("Check", pVal.Row))

                            Dim lightYellowOle As Integer = ColorTranslator.ToOle(Color.FromArgb(255, 255, 224)) ' LightYellow
                            lightYellowOle = ColorTranslator.ToOle(Color.YellowGreen)
                            Dim blackOle As Integer = ColorTranslator.ToOle(Color.Black)
                            Dim whiteOle As Integer = ColorTranslator.ToOle(Color.FromArgb(231, 236, 240))


                            Dim YELLOW As Integer = ColorTranslator.ToOle(Color.Yellow)
                            Dim BLUE As Integer = ColorTranslator.ToOle(Color.Blue)

                            If isChecked = "Y" Then
                                Grid.CommonSetting.SetRowBackColor(pVal.Row + 1, lightYellowOle)
                                Grid.CommonSetting.SetRowFontColor(pVal.Row + 1, BLUE)
                            ElseIf isChecked = "N" Then

                                Grid.CommonSetting.SetRowBackColor(pVal.Row + 1, whiteOle)
                                Grid.CommonSetting.SetRowFontColor(pVal.Row + 1, blackOle)
                            Else
                                Grid.CommonSetting.SetRowBackColor(pVal.Row + 1, whiteOle)
                                Grid.CommonSetting.SetRowFontColor(pVal.Row + 1, blackOle)
                            End If
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText("Error in checkbox click event: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        End Try

                    End If


                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    'If pVal.ItemUID = "8" And pVal.BeforeAction = False Then
                    '    If objForm.Items.Item("5").Specific.Value = "" Then
                    '        objMain.objApplication.StatusBar.SetText("Please Enter From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    ElseIf objForm.Items.Item("7").Specific.Value = "" Then
                    '        objMain.objApplication.StatusBar.SetText("Please Enter To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    Else
                    '        Me.LoadGrid(objForm.UniqueID, "Y")
                    '    End If
                    'End If

                    If pVal.ItemUID = "8" AndAlso pVal.BeforeAction = False Then
                        Dim sGRPODoc As String = objForm.DataSources.UserDataSources.Item("U_GRPO").ValueEx.Trim()

                        If sGRPODoc <> "" Then
                            Me.LoadGridBasedOnGRPODocEntry(objForm.UniqueID, "Y")

                        Else

                            If CType(objForm.Items.Item("5").Specific, SAPbouiCOM.EditText).Value.Trim() = "" Then
                                objMain.objApplication.StatusBar.SetText("Please Enter From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                            ElseIf CType(objForm.Items.Item("7").Specific, SAPbouiCOM.EditText).Value.Trim() = "" Then
                                objMain.objApplication.StatusBar.SetText("Please Enter To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                            Else
                                Me.LoadGrid(objForm.UniqueID, "Y")

                            End If

                        End If

                    End If

                    If pVal.ItemUID = "9" And pVal.ColUID = "Check" Then
                        Me.TotalAmount(objForm.UniqueID)
                    End If

                    If pVal.ItemUID = "17" And pVal.BeforeAction = False Then
                        Me.LoadGrid(objForm.UniqueID, "Y")
                    End If

                    If pVal.ItemUID = "18" And pVal.BeforeAction = False Then
                        Me.LoadGrid(objForm.UniqueID, "N")
                    End If


                    'If pVal.ItemUID = "3" And pVal.BeforeAction = False Then
                    '    Me.PostAPInvoice(objForm.UniqueID)
                    'End If

                    If pVal.ItemUID = "3" And pVal.BeforeAction = False Then
                        StartBackgroundForLoading()
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oGrid = objForm.Items.Item("9").Specific

                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        If (oCFL.UniqueID = "CFL_SPCD" Or oCFL.UniqueID = "CFL_SPNM") Then
                            objForm.DataSources.UserDataSources.Item("U_SUPCD").Value = oDT.GetValue("CardCode", 0)
                            objForm.DataSources.UserDataSources.Item("U_SUPNM").Value = oDT.GetValue("CardName", 0)
                        End If

                        If (oCFL.UniqueID = "CFL_GRPO") Then
                            objForm.DataSources.UserDataSources.Item("U_GRPO").ValueEx = oDT.GetValue("DocEntry", 0).ToString()
                        End If

                    End If

            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub LoadGridBasedOnGRPODocEntry(ByVal FormUID As String, ByVal Value As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            'objMain.objApplication.StatusBar.SetText("Loading initiated...please wait a while...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objForm.Freeze(True)

            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oGrid = objForm.Items.Item("9").Specific
            ' oGrid.DataTable = oDt

            'Try
            '    oGrid.DataTable.Clear()
            'Catch ex As Exception
            'End Try
            If oGrid.DataTable Is Nothing Then
                oDt = objForm.DataSources.DataTables.Add("DT1")
            Else
                oDt = objForm.DataSources.DataTables.Item("DT1")
            End If
            oGrid.DataTable = oDt

            Try
                'Dim whiteOle As Integer = ColorTranslator.ToOle(Color.White)
                Dim blackOle As Integer = ColorTranslator.ToOle(Color.Black)
                Dim whiteOle As Integer = ColorTranslator.ToOle(Color.FromArgb(231, 236, 240))

                For rowIndex As Integer = 0 To oGrid.DataTable.Rows.Count - 1

                    oGrid.CommonSetting.SetRowBackColor(rowIndex + 1, whiteOle)
                    oGrid.CommonSetting.SetRowFontColor(rowIndex + 1, blackOle)
                Next

                oGrid.DataTable.Clear()
            Catch ex As Exception
            End Try


            Dim Query As String = " Select '" & Value & "' as ""Check"", T0.""DocEntry"" As ""GRPOEntry"", ""DocNum"" As ""GRPONo"", T0.""DocDate"", T0.""CardCode"" As ""Supplier Code"", T0.""CardName"" As ""Supplier Name"", " &
                                 " ""DocTotal"", ""DiscSum"", ""NumAtCard"" As ""SupplierInvNo."", " & " CAST(NULL AS DATE) As ""Posting Date"", " &
                                  " CAST('' AS NVARCHAR(50)) As ""APInvoiceNo"", " & " CAST('' AS NVARCHAR(5000)) As ""Status"" " &
                                   " From OPDN T0 Inner Join PDN1 T1 On T1.""DocEntry"" = T0.""DocEntry"" " &
                                   " Where T0.""DocStatus"" = 'O' And T0.""DocType"" = 'I' And T1.""LineStatus"" = 'O' And T0.""CANCELED"" = 'N' "

            'If GRPO is selected
            If objForm.DataSources.UserDataSources.Item("U_GRPO").ValueEx.Trim() <> "" Then

                Query &= " And T0.""DocEntry"" = " & objForm.DataSources.UserDataSources.Item("U_GRPO").ValueEx.Trim()

            Else

                Query &= " And T0.""DocDate"" Between '" & objForm.Items.Item("5").Specific.Value &
             "' And '" & objForm.Items.Item("7").Specific.Value & "' "

                Query &= " And T0.""BPLId"" = '" & objForm.Items.Item("11").Specific.Value & "' "

            End If

            If objForm.DataSources.UserDataSources.Item("U_SUPCD").ValueEx.Trim() <> "" Then
                Query &= " And T0.""CardCode"" = '" & objForm.DataSources.UserDataSources.Item("U_SUPCD").ValueEx.Trim() & "' "
            End If

            Query &= " Group By T0.""DocEntry"", ""DocNum"", T0.""DocDate"", T0.""CardCode"", T0.""CardName"", ""DocTotal"", ""DiscSum"", ""NumAtCard"" "
            Query &= " Order By T0.""CardName"" ASC, T0.""DocNum"" ASC "

            oGrid.DataTable.ExecuteQuery(Query)


            If oGrid.Rows.Count = 0 Then
                objMain.objApplication.MessageBox("No records found")
                Exit Sub
            End If
            oGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item("GRPOEntry").Editable = False
            oGrid.Columns.Item("GRPONo").Editable = False
            oGrid.Columns.Item("DocDate").Editable = False
            oGrid.Columns.Item("Supplier Code").Editable = False
            oGrid.Columns.Item("Supplier Name").Editable = False
            oGrid.Columns.Item("DocTotal").Editable = False
            oGrid.Columns.Item("DiscSum").Editable = False
            oGrid.Columns.Item("SupplierInvNo.").Editable = False
            oGrid.Columns.Item("Posting Date").Editable = True
            oGrid.Columns.Item("APInvoiceNo").Editable = False
            oGrid.Columns.Item("Status").Editable = False

            Dim oEditTextCol As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("GRPOEntry")
            oEditTextCol.LinkedObjectType = 20

            Dim oEditTextCol1 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("Supplier Code")
            oEditTextCol1.LinkedObjectType = 2

            oEditTextCol = oGrid.Columns.Item("APInvoiceNo")
            oEditTextCol.LinkedObjectType = 18

            oGrid.AutoResizeColumns()
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            For rowIndex As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If oGrid.DataTable.GetValue("GRPOEntry", rowIndex) IsNot Nothing Then
                    Dim lightYellowOle As Integer = ColorTranslator.ToOle(Color.FromArgb(255, 255, 224)) ' LightYellow
                    lightYellowOle = ColorTranslator.ToOle(Color.YellowGreen)
                    Dim blackOle As Integer = ColorTranslator.ToOle(Color.Black)
                    Dim whiteOle As Integer = ColorTranslator.ToOle(Color.FromArgb(231, 236, 240))


                    Dim YELLOW As Integer = ColorTranslator.ToOle(Color.Yellow)
                    Dim BLUE As Integer = ColorTranslator.ToOle(Color.Blue)
                    Dim isChecked As String = (oGrid.DataTable.GetValue("Check", rowIndex))

                    If Value = "Y" Then
                        oGrid.CommonSetting.SetRowBackColor(rowIndex + 1, lightYellowOle)
                        oGrid.CommonSetting.SetRowFontColor(rowIndex + 1, BLUE)
                    ElseIf Value = "N" Then

                        oGrid.CommonSetting.SetRowBackColor(rowIndex + 1, whiteOle)
                        oGrid.CommonSetting.SetRowFontColor(rowIndex + 1, blackOle)
                    Else
                        oGrid.CommonSetting.SetRowBackColor(rowIndex + 1, whiteOle)
                        oGrid.CommonSetting.SetRowFontColor(rowIndex + 1, blackOle)
                    End If
                End If


            Next

            Me.TotalAmount(objForm.UniqueID)
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub LoadGrid(ByVal FormUID As String, ByVal Value As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            'objMain.objApplication.StatusBar.SetText("Loading initiated...please wait a while...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objForm.Freeze(True)

            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oGrid = objForm.Items.Item("9").Specific
            ' oGrid.DataTable = oDt

            'Try
            '    oGrid.DataTable.Clear()
            'Catch ex As Exception
            'End Try
            If oGrid.DataTable Is Nothing Then
                oDt = objForm.DataSources.DataTables.Add("DT1")
            Else
                oDt = objForm.DataSources.DataTables.Item("DT1")
            End If
            oGrid.DataTable = oDt

            Try
                'Dim whiteOle As Integer = ColorTranslator.ToOle(Color.White)
                Dim blackOle As Integer = ColorTranslator.ToOle(Color.Black)
                Dim whiteOle As Integer = ColorTranslator.ToOle(Color.FromArgb(231, 236, 240))

                For rowIndex As Integer = 0 To oGrid.DataTable.Rows.Count - 1

                    oGrid.CommonSetting.SetRowBackColor(rowIndex + 1, whiteOle)
                    oGrid.CommonSetting.SetRowFontColor(rowIndex + 1, blackOle)
                Next

                oGrid.DataTable.Clear()
            Catch ex As Exception
            End Try


            Dim Query As String = " Select '" & Value & "' as ""Check"", " & " T0.""DocEntry"" As ""GRPOEntry"", " & " ""DocNum"" As ""GRPONo"", " &
" T0.""DocDate"", " & " T0.""CardCode"" As ""Supplier Code"", " &
" T0.""CardName"" As ""Supplier Name"", " &
" ""DocTotal"", " &
" ""DiscSum"", " &
" ""NumAtCard"" As ""SupplierInvNo."", " &
" CAST('' AS DATE) As ""Posting Date"", " &
" CAST('' AS NVARCHAR(50)) As ""APInvoiceNo"", " &
" CAST('' AS NVARCHAR(5000)) As ""Status"" " &
" From OPDN T0 " &
" Inner Join PDN1 T1 On T1.""DocEntry"" = T0.""DocEntry"" " &
" Where ""DocStatus"" = 'O' " &
" And ""DocType"" = 'I' " &
" And ""LineStatus"" = 'O' " &
" And T0.""CANCELED"" = 'N' " &
" And T0.""DocDate"" Between '" & objForm.Items.Item("5").Specific.Value & "' And '" & objForm.Items.Item("7").Specific.Value & "' " &
" And ""BPLId"" = '" & objForm.Items.Item("11").Specific.Value & "' "

            If objForm.DataSources.UserDataSources.Item("U_SUPCD").Value <> "" Then
                Query &= " And T0.""CardCode"" = '" & objForm.Items.Item("14").Specific.Value & "' "
            End If

            Query &=
" Group By T0.""DocEntry"", ""DocNum"", T0.""DocDate"", T0.""CardCode"", T0.""CardName"", ""DocTotal"", ""DiscSum"", ""NumAtCard"" " &
" Order By T0.""CardName"" ASC, T0.""DocNum"" ASC"

            oGrid.DataTable.ExecuteQuery(Query)


            If oGrid.Rows.Count = 0 Then
                objMain.objApplication.MessageBox("No records found")
                Exit Sub
            End If
            oGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item("GRPOEntry").Editable = False
            oGrid.Columns.Item("GRPONo").Editable = False
            oGrid.Columns.Item("DocDate").Editable = False
            oGrid.Columns.Item("Supplier Code").Editable = False
            oGrid.Columns.Item("Supplier Name").Editable = False
            oGrid.Columns.Item("DocTotal").Editable = False
            oGrid.Columns.Item("DiscSum").Editable = False
            oGrid.Columns.Item("SupplierInvNo.").Editable = False
            oGrid.Columns.Item("Posting Date").Editable = True
            oGrid.Columns.Item("APInvoiceNo").Editable = False
            oGrid.Columns.Item("Status").Editable = False

            Dim oEditTextCol As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("GRPOEntry")
            oEditTextCol.LinkedObjectType = 20

            Dim oEditTextCol1 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("Supplier Code")
            oEditTextCol1.LinkedObjectType = 2

            oEditTextCol = oGrid.Columns.Item("APInvoiceNo")
            oEditTextCol.LinkedObjectType = 18

            oGrid.AutoResizeColumns()
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            For rowIndex As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If oGrid.DataTable.GetValue("GRPOEntry", rowIndex) IsNot Nothing Then
                    Dim lightYellowOle As Integer = ColorTranslator.ToOle(Color.FromArgb(255, 255, 224)) ' LightYellow
                    lightYellowOle = ColorTranslator.ToOle(Color.YellowGreen)
                    Dim blackOle As Integer = ColorTranslator.ToOle(Color.Black)
                    Dim whiteOle As Integer = ColorTranslator.ToOle(Color.FromArgb(231, 236, 240))


                    Dim YELLOW As Integer = ColorTranslator.ToOle(Color.Yellow)
                    Dim BLUE As Integer = ColorTranslator.ToOle(Color.Blue)
                    Dim isChecked As String = (oGrid.DataTable.GetValue("Check", rowIndex))

                    If Value = "Y" Then
                        oGrid.CommonSetting.SetRowBackColor(rowIndex + 1, lightYellowOle)
                        oGrid.CommonSetting.SetRowFontColor(rowIndex + 1, BLUE)
                    ElseIf Value = "N" Then

                        oGrid.CommonSetting.SetRowBackColor(rowIndex + 1, whiteOle)
                        oGrid.CommonSetting.SetRowFontColor(rowIndex + 1, blackOle)
                    Else
                        oGrid.CommonSetting.SetRowBackColor(rowIndex + 1, whiteOle)
                        oGrid.CommonSetting.SetRowFontColor(rowIndex + 1, blackOle)
                    End If
                End If


            Next

            Me.TotalAmount(objForm.UniqueID)
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objMain.objApplication.StatusBar.SetText(ex.Message)
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
        Me.PostAPInvoice(objForm.UniqueID)
    End Sub

    Private Sub PostAPInvoice(ByVal FormUID As String)
        Dim oProgressPostAPInvoice As SAPbouiCOM.ProgressBar = Nothing
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("9").Specific
            Dim oPurchaseInvoices As SAPbobsCOM.Documents
            Dim j1 As Integer = 1
            oProgressPostAPInvoice = objMain.objApplication.StatusBar.CreateProgressBar("Please Wait AP Invoice docments generating....!", oGrid.DataTable.Rows.Count - 1, False)
            oProgressPostAPInvoice.Text = "Please Wait AP Invoice docments generating....!"
            For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If oGrid.DataTable.GetValue("Check", i) = "Y" Then
                    oPurchaseInvoices = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                    Dim ExpenseTotal As Double = 0
                    Dim GetGRPOBaseDates As String = "Select ""TotalExpns"" , ""DocDate"" , ""DocDueDate"" , ""TaxDate"" , ""Rounding"" , ""RoundDif""  From ""OPDN"" Where ""DocEntry"" = '" & oGrid.DataTable.GetValue("GRPOEntry", i) & "'  "
                    Dim oRsGetGRPOBaseDates As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetGRPOBaseDates.DoQuery(GetGRPOBaseDates)
                    Dim GetGRPODetails As String = "Select ""LineNum"" , IFNULL(SUM(IFNULL(""Quantity"",0)),0) As ""Qty"" From ""PDN1"" Where ""DocEntry"" = '" & oGrid.DataTable.GetValue("GRPOEntry", i) & "' Group By ""LineNum""  "
                    Dim oRsGetGRPODetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetGRPODetails.DoQuery(GetGRPODetails)
                    If oRsGetGRPODetails.RecordCount > 0 Then
                        For k As Integer = 1 To oRsGetGRPODetails.RecordCount
                            If k = 1 Then
                                oPurchaseInvoices.CardCode = oGrid.DataTable.GetValue("Supplier Code", i)
                                If objForm.Items.Item("13").Specific.Value <> "" Then
                                    oPurchaseInvoices.Series = objForm.Items.Item("13").Specific.Value.Trim
                                End If
                                If Trim(oGrid.DataTable.GetValue("Posting Date", i).ToString()) <> "" Then
                                    Dim PostingDate As Date = CDate(oGrid.DataTable.GetValue("Posting Date", i))

                                    oPurchaseInvoices.DocDate = PostingDate
                                    oPurchaseInvoices.DocDueDate = PostingDate
                                    oPurchaseInvoices.TaxDate = PostingDate
                                Else
                                    oPurchaseInvoices.DocDate = oRsGetGRPOBaseDates.Fields.Item("DocDate").Value
                                    oPurchaseInvoices.DocDueDate = oRsGetGRPOBaseDates.Fields.Item("DocDueDate").Value
                                    oPurchaseInvoices.TaxDate = oRsGetGRPOBaseDates.Fields.Item("TaxDate").Value
                                End If
                                '--------Madhu -----Clinet Required only Posting date want
                                'oPurchaseInvoices.DocDate = oRsGetGRPOBaseDates.Fields.Item("DocDate").Value
                                'oPurchaseInvoices.DocDueDate = oRsGetGRPOBaseDates.Fields.Item("DocDueDate").Value
                                'oPurchaseInvoices.TaxDate = oRsGetGRPOBaseDates.Fields.Item("TaxDate").Value
                                Dim objBranchCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("11").Specific
                                oPurchaseInvoices.BPL_IDAssignedToInvoice = objBranchCombo.Value
                                If oRsGetGRPOBaseDates.Fields.Item("Rounding").Value = "Y" Then
                                    oPurchaseInvoices.Rounding = SAPbobsCOM.BoYesNoEnum.tYES
                                    oPurchaseInvoices.RoundingDiffAmount = CDbl(oRsGetGRPOBaseDates.Fields.Item("RoundDif").Value)
                                End If
                                If oRsGetGRPOBaseDates.Fields.Item("TotalExpns").Value > 0 Then
                                    Dim GetExpenseCodes As String = "Select ""ExpnsCode"" , ""ExpnsName"" From OEXD"
                                    Dim oRsGetExpenseCodes As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRsGetExpenseCodes.DoQuery(GetExpenseCodes)
                                    If oRsGetExpenseCodes.RecordCount > 0 Then
                                        oRsGetExpenseCodes.MoveFirst()
                                        For p As Integer = 1 To oRsGetExpenseCodes.RecordCount
                                            Dim GetPurchaseExpense As String = "Select ""DistrbMthd"", ""VatGroup"" , ""LineTotal"" , ""TotalFrgn"" , ""LineNum"" , ""DocEntry"" From PDN3 Where " &
                                            " ""DocEntry"" = '" & oGrid.DataTable.GetValue("GRPOEntry", i) & "'  " &
                                            " And ""ExpnsCode"" = '" & oRsGetExpenseCodes.Fields.Item(0).Value & "'"
                                            Dim oRsGetPurchaseExpense As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRsGetPurchaseExpense.DoQuery(GetPurchaseExpense)
                                            Dim PurchaseLineTotal As Double = 0
                                            For l As Integer = 1 To oRsGetPurchaseExpense.RecordCount
                                                Try
                                                    PurchaseLineTotal = oRsGetPurchaseExpense.Fields.Item("LineTotal").Value
                                                    Dim PurchaseFreightPerQty As Double = PurchaseLineTotal / CDbl(oRsGetGRPODetails.Fields.Item("Qty").Value)
                                                    PurchaseLineTotal = PurchaseFreightPerQty * CDbl(oRsGetGRPODetails.Fields.Item("Qty").Value)
                                                Catch ex As Exception
                                                End Try
                                                If PurchaseLineTotal > 0 Then
                                                    oPurchaseInvoices.Expenses.ExpenseCode = oRsGetExpenseCodes.Fields.Item("ExpnsCode").Value
                                                    oPurchaseInvoices.Expenses.LineTotal = Math.Round(PurchaseLineTotal, 2)
                                                    If oRsGetPurchaseExpense.Fields.Item("DistrbMthd").Value = "Q" Then
                                                        oPurchaseInvoices.Expenses.DistributionMethod = SAPbobsCOM.BoAdEpnsDistribMethods.aedm_Quantity
                                                    End If
                                                    If oRsGetPurchaseExpense.Fields.Item("DistrbMthd").Value = "T" Then
                                                        oPurchaseInvoices.Expenses.DistributionMethod = SAPbobsCOM.BoAdEpnsDistribMethods.aedm_RowTotal
                                                    End If
                                                    If oRsGetPurchaseExpense.RecordCount > 0 Then
                                                        oPurchaseInvoices.Expenses.VatGroup = oRsGetPurchaseExpense.Fields.Item("VatGroup").Value
                                                        oPurchaseInvoices.Expenses.BaseDocEntry = oRsGetPurchaseExpense.Fields.Item("DocEntry").Value
                                                        oPurchaseInvoices.Expenses.BaseDocLine = oRsGetPurchaseExpense.Fields.Item("LineNum").Value
                                                        oPurchaseInvoices.Expenses.BaseDocType = "20"
                                                    Else
                                                        oPurchaseInvoices.Expenses.VatGroup = oRsGetPurchaseExpense.Fields.Item("VatGroup").Value
                                                    End If
                                                    oPurchaseInvoices.Expenses.Add()
                                                End If
                                                oRsGetPurchaseExpense.MoveLast()
                                            Next
                                            oRsGetExpenseCodes.MoveNext()
                                        Next
                                    End If
                                End If
                            End If
                            oPurchaseInvoices.Lines.BaseEntry = oGrid.DataTable.GetValue("GRPOEntry", i)
                            oPurchaseInvoices.Lines.BaseLine = oRsGetGRPODetails.Fields.Item("LineNum").Value
                            oPurchaseInvoices.Lines.BaseType = "20"
                            oPurchaseInvoices.Lines.Add()
                            oRsGetGRPODetails.MoveNext()
                        Next
                        If oPurchaseInvoices.Add = 0 Then
                            Dim GetAPInvoiceEntry As String = "Select Max(""DocEntry"") From OPCH  "
                            Dim oRsGetAPInvoiceEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetAPInvoiceEntry.DoQuery(GetAPInvoiceEntry)
                            oGrid.DataTable.SetValue("APInvoiceNo", i, oRsGetAPInvoiceEntry.Fields.Item(0).Value.ToString)
                            oGrid.DataTable.SetValue("Status", i, "Posted Successfully")
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oPurchaseInvoices)
                        Else
                            oGrid.DataTable.SetValue("APInvoiceNo", i, "")
                            oGrid.DataTable.SetValue("Status", i, objMain.objCompany.GetLastErrorDescription)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oPurchaseInvoices)
                        End If
                    End If
                End If
                oProgressPostAPInvoice.Value = j1
                j1 = j1 + 1
            Next
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        Finally
            oGrid.AutoResizeColumns()
            If oProgressPostAPInvoice IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgressPostAPInvoice)
            End If
        End Try
    End Sub

    Public Sub TotalAmount(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim TotalSum As Double = 0
            For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If oGrid.DataTable.GetValue("Check", i) = "Y" Then
                    TotalSum = TotalSum + CDbl(oGrid.DataTable.GetValue("DocTotal", i))
                End If
            Next
            objForm.DataSources.UserDataSources.Item("U_TTLAMT").Value = TotalSum
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
