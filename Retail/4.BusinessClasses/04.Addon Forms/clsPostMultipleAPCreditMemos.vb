
Imports CrystalDecisions.[Shared].Json
Imports SAPbouiCOM
Imports System.Drawing
Imports System.Windows

Public Class clsPostMultipleAPCreditMemos

#Region "        Declaration        "
    Dim objForm, objNewForm As SAPbouiCOM.Form
    Dim oGrid As SAPbouiCOM.Grid
    Dim oDt As SAPbouiCOM.DataTable
    Dim oEditText As SAPbouiCOM.EditText
    Dim objComboBox As SAPbouiCOM.ComboBox
    Private lastSortedColumn As String = ""
    Private lastSortDirection As SAPbouiCOM.BoGridSortType = SAPbouiCOM.BoGridSortType.gst_Ascending
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("PostMultipleAPCreditMemos.xml", "ME_PMAPCM_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_PMAPCM_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.DataSources.UserDataSources.Add("U_FROMDT", SAPbouiCOM.BoDataType.dt_DATE)
            objForm.DataSources.UserDataSources.Add("U_TODT", SAPbouiCOM.BoDataType.dt_DATE)
            objForm.DataSources.UserDataSources.Add("U_BRANCH", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            objForm.DataSources.UserDataSources.Add("U_SERIES", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            objForm.DataSources.UserDataSources.Add("U_SUPCD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_SUPNM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_TTLAMT", SAPbouiCOM.BoDataType.dt_PRICE)


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
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = '19' And ""Locked"" = 'N' And ""BPLId"" = '" & objBranchCombo.Value & "'   Order By ""Series"" Desc ")
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly


            If objComboBox.ValidValues.Count <> 0 Then
                Dim GetDefaultSeries As String = "Select ""Series"" , ""SeriesName"" , ""BPLId""  From NNM1 Where ""ObjectCode"" = '19' And " &
                                                 " ""Locked"" = 'N' And ""BPLId"" = '" & objBranchCombo.Value & "' And ""Indicator"" = IFNULL((Select ""U_APCMSRS"" From ""@ME_ACFG""),0)  Order By ""Series"" Desc"
                Dim oRsGetDefaultSeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetDefaultSeries.DoQuery(GetDefaultSeries)
                Dim test As String = oRsGetDefaultSeries.Fields.Item("Series").Value.ToString
                If oRsGetDefaultSeries.RecordCount > 0 And oRsGetDefaultSeries.Fields.Item("Series").Value.ToString <> "0" Then
                    objComboBox.Select(oRsGetDefaultSeries.Fields.Item("Series").Value.ToString, SAPbouiCOM.BoSearchKey.psk_ByValue)
                Else
                    Try
                        objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
                    Catch ex As Exception

                    End Try

                End If
            End If


            oDt = objForm.DataSources.DataTables.Add("dt1")
            oDt = objForm.DataSources.DataTables.Item("dt1")

            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ME_PMAPCM" And pVal.BeforeAction = False Then
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
                        'objMain.objUtilities.ComboBoxLoadValues(objComboBox, "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = '19' And ""Locked"" = 'N' And ""BPLId"" = '" & objBranchCombo.Value & "' Order By ""Series"" Desc ")
                        'objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                        'If objComboBox.ValidValues.Count <> 0 Then
                        '    objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
                        'End If
                        objMain.objUtilities.ComboBoxLoadValues(objComboBox, "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = '19' And ""Locked"" = 'N' And ""BPLId"" = '" & objBranchCombo.Value & "'   Order By ""Series"" Desc ")
                        objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                        If objComboBox.ValidValues.Count <> 0 Then
                            Dim GetDefaultSeries As String = "Select ""Series"" , ""SeriesName"" , ""BPLId""  From NNM1 Where ""ObjectCode"" = '19' And " &
                                                             " ""Locked"" = 'N' And ""BPLId"" = '" & objBranchCombo.Value & "' And ""Indicator"" = IFNULL((Select ""U_APCMSRS"" From ""@ME_ACFG""),0)  Order By ""Series"" Desc"
                            Dim oRsGetDefaultSeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetDefaultSeries.DoQuery(GetDefaultSeries)
                            If oRsGetDefaultSeries.RecordCount > 0 And oRsGetDefaultSeries.Fields.Item("Series").Value.ToString <> "0" Then
                                objComboBox.Select(oRsGetDefaultSeries.Fields.Item("Series").Value.ToString, SAPbouiCOM.BoSearchKey.psk_ByValue)
                            Else
                                Try
                                    objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
                                Catch ex As Exception
                                End Try
                            End If
                        End If
                    End If
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

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "8" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("5").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        ElseIf objForm.Items.Item("7").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            Me.LoadGrid(objForm.UniqueID, "Y")
                        End If
                    End If
                    If pVal.ItemUID = "3" And pVal.BeforeAction = False Then
                        objForm.Items.Item("3").Enabled = False
                        Me.PostAPCreditMemo(objForm.UniqueID)
                        objForm.Items.Item("3").Enabled = True
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
                    End If
            End Select
        Catch ex As Exception
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

            Dim Query As String = " Select '" & Value & "' as ""Check"" , T0.""DocEntry"" As ""GoodsReturnEntry"" , ""DocNum"" As ""GoodsReturnNo"" , T0.""DocDate"" , T0.""CardCode"" As ""Supplier Code"" , T0.""CardName"" As ""Supplier Name"" , " &
                                  " ""DocTotal"", ""DiscSum"" , ""NumAtCard"" As ""SupplierRefNo."" ,  '          ' As ""APCreditMemo"" " &
                                  " , cast('' AS NVARCHAR(5000)) As ""Status""   " &
                                  " From ORPD T0 Inner Join RPD1 T1 On T1.""DocEntry""=T0.""DocEntry"" Where ""DocStatus""= 'O' And ""DocType""='I' And ""LineStatus"" = 'O' And T0.""CANCELED"" = 'N'  " &
                                  " And T0.""DocDate"" Between  '" & objForm.Items.Item("5").Specific.Value & "' And '" & objForm.Items.Item("7").Specific.Value & "' And   ""BPLId"" = '" & objForm.Items.Item("11").Specific.Value & "' "
            If objForm.DataSources.UserDataSources.Item("U_SUPCD").Value <> "" Then
                Query = Query & " And T0.""CardCode"" = '" & objForm.Items.Item("14").Specific.Value & "'"
            End If

            Query = Query + "Group By T0.""DocEntry"" , ""DocNum""  , T0.""DocDate"" , T0.""CardCode"" , T0.""CardName"" ,  ""DocTotal"", ""DiscSum"" , ""NumAtCard"" ORDER BY T0.""CardName"" ASC , T0.""DocNum"" ASC"

            oGrid.DataTable.ExecuteQuery(Query)
            If oGrid.Rows.Count = 0 Then
                objMain.objApplication.MessageBox("No records found")
                Exit Sub
            End If
            oGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item("GoodsReturnEntry").Editable = False
            oGrid.Columns.Item("GoodsReturnNo").Editable = False
            oGrid.Columns.Item("DocDate").Editable = False
            oGrid.Columns.Item("Supplier Code").Editable = False
            oGrid.Columns.Item("Supplier Name").Editable = False
            oGrid.Columns.Item("DocTotal").Editable = False
            oGrid.Columns.Item("DiscSum").Editable = False
            oGrid.Columns.Item("APCreditMemo").Editable = False
            oGrid.Columns.Item("SupplierRefNo.").Editable = False
            oGrid.Columns.Item("Status").Editable = False

            Dim oEditTextCol As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("GoodsReturnEntry")
            oEditTextCol.LinkedObjectType = 21

            oEditTextCol = oGrid.Columns.Item("APCreditMemo")
            oEditTextCol.LinkedObjectType = 19

            Dim oEditTextCol1 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("Supplier Code")
            oEditTextCol1.LinkedObjectType = 2

            oGrid.AutoResizeColumns()
            Dim c As Integer = oGrid.Rows.Count
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            If oGrid.Rows.Count <> 0 Then

                For rowIndex As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                    If oGrid.DataTable.GetValue("GoodsReturnEntry", rowIndex) IsNot Nothing Then


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
            End If

            Me.TotalAmount(objForm.UniqueID)
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    Private Sub PostAPCreditMemo(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("9").Specific
            Dim oPurchaseCreditNotes As SAPbobsCOM.Documents
            For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If oGrid.DataTable.GetValue("Check", i) = "Y" Then
                    objMain.objApplication.StatusBar.SetText("Posting AP  Credit Memo for line " & i + 1 & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                    oPurchaseCreditNotes = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)

                    Dim GetGoodsReturnBaseDates As String = "Select ""DocNum"" , ""DocDate"" , ""DocDueDate"" , ""TaxDate"" , ""Rounding"" , ""RoundDif"",""BPLId""  From ""ORPD"" Where ""DocEntry"" = '" & oGrid.DataTable.GetValue("GoodsReturnEntry", i) & "'  "
                    Dim oRsGetGoodsReturnBaseDates As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetGoodsReturnBaseDates.DoQuery(GetGoodsReturnBaseDates)

                    Dim GetGoodsReturnDetails As String = "Select ""LineNum"" From ""RPD1"" Where ""DocEntry"" = '" & oGrid.DataTable.GetValue("GoodsReturnEntry", i) & "'  "
                    Dim oRsGetGoodsReturnDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetGoodsReturnDetails.DoQuery(GetGoodsReturnDetails)
                    If oRsGetGoodsReturnDetails.RecordCount > 0 Then
                        For k As Integer = 1 To oRsGetGoodsReturnDetails.RecordCount
                            If k = 1 Then
                                oPurchaseCreditNotes.CardCode = oGrid.DataTable.GetValue("Supplier Code", i)
                                If objForm.Items.Item("13").Specific.Value <> "" Then
                                    oPurchaseCreditNotes.Series = objForm.Items.Item("13").Specific.Value.Trim
                                End If
                                'Dim sDate As String = DateTime.Now.ToString("yyyyMMdd")
                                'sDate = sDate.Insert("4", "-")
                                'sDate = sDate.Insert("7", "-")
                                'oPurchaseCreditNotes.DocDate = sDate
                                'oPurchaseCreditNotes.DocDueDate = sDate
                                'oPurchaseCreditNotes.TaxDate = sDate
                                oPurchaseCreditNotes.DocDate = oRsGetGoodsReturnBaseDates.Fields.Item("DocDate").Value
                                oPurchaseCreditNotes.DocDueDate = oRsGetGoodsReturnBaseDates.Fields.Item("DocDueDate").Value
                                oPurchaseCreditNotes.TaxDate = oRsGetGoodsReturnBaseDates.Fields.Item("TaxDate").Value
                                oPurchaseCreditNotes.BPL_IDAssignedToInvoice = oRsGetGoodsReturnBaseDates.Fields.Item("BPLId").Value

                                If oRsGetGoodsReturnBaseDates.Fields.Item("Rounding").Value = "Y" Then
                                    oPurchaseCreditNotes.Rounding = SAPbobsCOM.BoYesNoEnum.tYES
                                    oPurchaseCreditNotes.RoundingDiffAmount = CDbl(oRsGetGoodsReturnBaseDates.Fields.Item("RoundDif").Value)
                                End If
                                oPurchaseCreditNotes.JournalMemo = "Based on Purchase Return - " & oRsGetGoodsReturnBaseDates.Fields.Item("DocNum").Value
                            End If
                            oPurchaseCreditNotes.Lines.BaseEntry = oGrid.DataTable.GetValue("GoodsReturnEntry", i)
                            oPurchaseCreditNotes.Lines.BaseLine = oRsGetGoodsReturnDetails.Fields.Item("LineNum").Value
                            oPurchaseCreditNotes.Lines.BaseType = "21"
                            oPurchaseCreditNotes.Lines.Add()
                            oRsGetGoodsReturnDetails.MoveNext()
                        Next
                        If oPurchaseCreditNotes.Add = 0 Then
                            Dim GetAPInvoiceEntry As String = "Select Max(""DocEntry"") From ORPC  "
                            Dim oRsGetAPInvoiceEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetAPInvoiceEntry.DoQuery(GetAPInvoiceEntry)
                            oGrid.DataTable.SetValue("APCreditMemo", i, oRsGetAPInvoiceEntry.Fields.Item(0).Value.ToString)
                            oGrid.DataTable.SetValue("Status", i, "Posted Successfully")
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oPurchaseCreditNotes)
                        Else
                            oGrid.DataTable.SetValue("APCreditMemo", i, "")
                            oGrid.DataTable.SetValue("Status", i, objMain.objCompany.GetLastErrorDescription)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oPurchaseCreditNotes)
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            objForm.Items.Item("3").Enabled = True
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    Public Sub TotalAmount(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("9").Specific
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
