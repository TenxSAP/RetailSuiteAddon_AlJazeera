Public Class clsStockCountGenerationWizard

#Region "        Declaration        "
    Dim objForm, objICForm As SAPbouiCOM.Form
    Dim oGrid As SAPbouiCOM.Grid
    Dim oDt As SAPbouiCOM.DataTable
    Dim oEditText As SAPbouiCOM.EditText
    Dim objComboBox As SAPbouiCOM.ComboBox
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("StockCountGenerationWizard.xml", "ME_SCGW_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_SCGW_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objMain.objApplication.StatusBar.SetText("Please wait... Form is Initialilsing....", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            objForm.DataSources.UserDataSources.Add("U_FROMDT", SAPbouiCOM.BoDataType.dt_DATE)
            objForm.DataSources.UserDataSources.Add("U_TODT", SAPbouiCOM.BoDataType.dt_DATE)
            objForm.DataSources.UserDataSources.Add("U_CNTDT", SAPbouiCOM.BoDataType.dt_DATE)
            objForm.DataSources.UserDataSources.Add("U_PSDT", SAPbouiCOM.BoDataType.dt_DATE)
            objForm.DataSources.UserDataSources.Add("U_CNTTM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_BRANCH", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            objForm.DataSources.UserDataSources.Add("U_TYPE", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            objForm.DataSources.UserDataSources.Add("U_ITGRP", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            objForm.DataSources.UserDataSources.Add("U_SBGRP", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            objForm.DataSources.UserDataSources.Add("U_SBSBGRP", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            objForm.DataSources.UserDataSources.Add("U_WHSCD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_INEXITM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            Dim oCheckBox As SAPbouiCOM.CheckBox = objForm.Items.Item("26").Specific
            oCheckBox.DataBind.SetBound(True, "", "U_INEXITM")
            objForm.Items.Item("26").AffectsFormMode = False

            oEditText = objForm.Items.Item("5").Specific
            oEditText.DataBind.SetBound(True, "", "U_FROMDT")
            objForm.Items.Item("5").AffectsFormMode = False

            oEditText = objForm.Items.Item("7").Specific
            oEditText.DataBind.SetBound(True, "", "U_TODT")
            objForm.Items.Item("7").AffectsFormMode = False

            oEditText = objForm.Items.Item("28").Specific
            oEditText.DataBind.SetBound(True, "", "U_PSDT")
            objForm.Items.Item("28").AffectsFormMode = False

            oEditText = objForm.Items.Item("9").Specific
            oEditText.DataBind.SetBound(True, "", "U_CNTDT")
            objForm.Items.Item("9").AffectsFormMode = False

            oEditText = objForm.Items.Item("11").Specific
            oEditText.DataBind.SetBound(True, "", "U_CNTTM")
            objForm.Items.Item("11").AffectsFormMode = False

            oEditText = objForm.Items.Item("13").Specific
            oEditText.DataBind.SetBound(True, "", "U_WHSCD")
            oEditText.ChooseFromListUID = "CFL_WHS"
            oEditText.ChooseFromListAlias = "WhsCode"
            objForm.Items.Item("13").AffectsFormMode = False

            objComboBox = objForm.Items.Item("15").Specific
            objComboBox.DataBind.SetBound(True, "", "U_BRANCH")
            objForm.Items.Item("15").AffectsFormMode = False
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, "Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""  ")
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("15").DisplayDesc = True
            'If objComboBox.ValidValues.Count <> 0 Then
            '    objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
            'End If

            objComboBox = objForm.Items.Item("17").Specific
            objComboBox.DataBind.SetBound(True, "", "U_TYPE")
            objForm.Items.Item("17").AffectsFormMode = False
            objComboBox.ValidValues.Add("All", "All")
            objComboBox.ValidValues.Add("ItemGroup", "ItemGroup")
            objComboBox.ValidValues.Add("SubGroup", "SubGroup")
            objComboBox.ValidValues.Add("SubSubGroup", "SubSubGroup")
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            ' objComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

            objComboBox = objForm.Items.Item("19").Specific
            objComboBox.DataBind.SetBound(True, "", "U_ITGRP")
            objForm.Items.Item("19").AffectsFormMode = False
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, "Select ""ItmsGrpCod""  , ""ItmsGrpNam""  From OITB Order By  ""ItmsGrpNam""")
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("19").DisplayDesc = True

            objComboBox = objForm.Items.Item("21").Specific
            objComboBox.DataBind.SetBound(True, "", "U_SBGRP")
            objForm.Items.Item("21").AffectsFormMode = False
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, "SELECT T0.""Code"", T0.""Name"" FROM ""@VR_SUBGROUP"" T0 Order By  ""Name"" ")
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("21").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("21").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("21").DisplayDesc = True

            objComboBox = objForm.Items.Item("23").Specific
            objComboBox.DataBind.SetBound(True, "", "U_SBSBGRP")
            objForm.Items.Item("23").AffectsFormMode = False
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, "SELECT T0.""Code"", T0.""Name"" FROM ""@VR_SUBSUBGROUP"" T0 Order By  ""Name"" ")
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("23").DisplayDesc = True

            oDt = objForm.DataSources.DataTables.Add("dt1")
            oDt = objForm.DataSources.DataTables.Item("dt1")

            objMain.objApplication.StatusBar.SetText("Form loaded successfully.....countinue.!!.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized

            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)


        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        Finally
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            'If pVal.MenuUID = "ME_SCGW" And pVal.BeforeAction = False Then
            '    Me.CreateForm()
            'End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    If pVal.ItemUID = "24" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("5").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        ElseIf objForm.Items.Item("7").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            Me.LoadGrid(objForm.UniqueID, "Y")
                        End If
                    End If

                    'If pVal.ItemUID = "17" And pVal.BeforeAction = False Then
                    '    Me.LoadGrid(objForm.UniqueID, "Y")
                    'End If

                    'If pVal.ItemUID = "18" And pVal.BeforeAction = False Then
                    '    Me.LoadGrid(objForm.UniqueID, "N")
                    'End If

                    If pVal.ItemUID = "25" And pVal.BeforeAction = False Then
                        Me.UPDATETABLE(objForm.UniqueID)
                        'Me.PostStockCount(objForm.UniqueID)
                    End If

                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oGrid = objForm.Items.Item("3").Specific

                    If pVal.ItemUID = "17" And pVal.BeforeAction = False Then
                        Select Case objForm.Items.Item("17").Specific.Selected.Value
                            Case "All"
                                objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objComboBox = objForm.Items.Item("19").Specific
                                objComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                objForm.Items.Item("21").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("21").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objComboBox = objForm.Items.Item("21").Specific
                                objComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objComboBox = objForm.Items.Item("23").Specific
                                objComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                            Case "ItemGroup"
                                objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("21").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("21").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            Case "SubGroup"
                                objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("21").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            Case "SubSubGroup"
                                objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("21").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        End Select
                    End If
                    If pVal.ItemUID = "19" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("19").Specific.Selected.Value <> "" Then
                            Dim sQury As String = "SELECT T0.""Code"", T0.""Name"" FROM ""@VR_SUBGROUP"" T0 WHERE T0.""U_ITEMGROUP"" ='" & objForm.Items.Item("19").Specific.Selected.Value & "' "
                            objComboBox = objForm.Items.Item("21").Specific
                            objMain.objUtilities.ComboBoxLoadValues(objComboBox, sQury)
                            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                            objForm.Items.Item("21").DisplayDesc = True
                        End If
                    End If
                    If pVal.ItemUID = "21" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("21").Specific.Selected.Value <> "" Then
                            Dim sQury As String = "SELECT T0.""Code"", T0.""Name"" FROM ""@VR_SUBSUBGROUP""  T0 WHERE T0.""U_SUBGROUP"" = '" & objForm.Items.Item("21").Specific.Selected.Value & "' "
                            objComboBox = objForm.Items.Item("23").Specific
                            objMain.objUtilities.ComboBoxLoadValues(objComboBox, sQury)
                            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                            objForm.Items.Item("23").DisplayDesc = True
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oGrid = objForm.Items.Item("3").Specific
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
                        If (oCFL.UniqueID = "CFL_WHS") Then
                            objForm.DataSources.UserDataSources.Item("U_WHSCD").Value = oDT.GetValue("WhsCode", 0)
                            ' objForm.DataSources.UserDataSources.Item("U_SUPNM").Value = oDT.GetValue("CardName", 0)
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
            objForm.Freeze(True)

            objForm.Items.Item("24").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            oGrid = objForm.Items.Item("3").Specific
            oGrid.DataTable = oDt

            Try
                oGrid.DataTable.Clear()
            Catch ex As Exception
            End Try
            Dim IncludeExStockItems As SAPbouiCOM.CheckBox = objForm.Items.Item("26").Specific
            If IncludeExStockItems.Checked = False Then
                'Dim Query As String = " Select '" & Value & "' as ""Check""   , ""ItemCode"" , ""ItemName"" , SUM(""ActualQty"") As ""Qty"" , ""BaseUom""  From ( " & _
                '                 " Select    ""U_ItemCode"" As ""ItemCode"" , ""U_ItemName"" As  ""ItemName"" , ""U_Qty"" As ""Quantity"" , ""U_UomName"" As ""UomCode"" , " & _
                '                 " (Select (T11.""BaseQty"" * T1.""U_Qty"") From ""OITM"" T00  INNER JOIN ""UGP1"" T11 ON T00.""UgpEntry"" = T11.""UgpEntry"" INNER JOIN ""OUOM"" T22 ON T22.""UomEntry"" = T11.""UomEntry""  WHERE T00.""ItemCode"" =T1.""U_ItemCode"" And  T22.""UomCode"" = T1.""U_UomName"")  As ""ActualQty"" " & _
                '                 " , ""U_BaseUom"" As ""BaseUom"" From ""@ME_PSC"" T0 Inner Join ""@ME_PSC_C0"" T1 On T1.""DocEntry"" = T0.""DocEntry"" " & _
                '                 " Where  ""U_ItemCode"" <> '' And T0.""U_CountDate"" Between  '" & objForm.Items.Item("5").Specific.Value & "' And " & _
                '                 " '" & objForm.Items.Item("7").Specific.Value & "' And   ""U_Branch"" = '" & objForm.Items.Item("15").Specific.Selected.Value & "' And ""U_WhseCode"" = '" & objForm.Items.Item("13").Specific.Value & "'"
                'If objForm.Items.Item("17").Specific.Selected.Value = "All" Then
                '    Query = Query + "And T0.""U_TYPE"" = 'All'  "
                'ElseIf objForm.Items.Item("17").Specific.Selected.Value = "ItemGroup" Then
                '    Query = Query + "And T0.""U_TYPE"" = 'ItemGroup' And T0.""U_ITMGRP"" = '" & objForm.Items.Item("19").Specific.Selected.Value.Trim & "'  "
                'ElseIf objForm.Items.Item("17").Specific.Selected.Value = "SubGroup" Then
                '    Query = Query + "And T0.""U_TYPE"" = 'SubGroup' And  T0.""U_ITMGRP"" = '" & objForm.Items.Item("19").Specific.Selected.Value.Trim & "' And T0.""U_SUBGRP"" = '" & objForm.Items.Item("21").Specific.Selected.Value.Trim & "' "
                'ElseIf objForm.Items.Item("17").Specific.Selected.Value = "SubSubGroup" Then
                '    Query = Query + "And  T0.""U_TYPE"" = 'SubSubGroup' And T0.""U_ITMGRP"" = '" & objForm.Items.Item("19").Specific.Selected.Value.Trim & "' And  T0.""U_SUBGRP"" = '" & objForm.Items.Item("21").Specific.Selected.Value.Trim & "' And T0.""U_SBSBGRP"" = '" & objForm.Items.Item("23").Specific.Selected.Value.Trim & "' "
                'End If
                'Query = Query + " ) Group By ""ItemCode"" , ""ItemName"" , ""BaseUom"""
                Dim Query As String = " Select '" & Value & "' as ""Check""   , ""ItemCode"" , ""ItemName"" , SUM(""ActualQty"") As ""Qty"" , ""BaseUom""  From ( " &
                                 " Select    T1.""U_ItemCode"" As ""ItemCode"" , T1.""U_ItemName"" As  ""ItemName"" , ""U_Qty"" As ""Quantity"" , T1.""U_UomName"" As ""UomCode"" , " &
                                 " (Select (T11.""BaseQty"" * T1.""U_Qty"") From ""OITM"" T00  INNER JOIN ""UGP1"" T11 ON T00.""UgpEntry"" = T11.""UgpEntry"" INNER JOIN ""OUOM"" T22 ON T22.""UomEntry"" = T11.""UomEntry""  WHERE T00.""ItemCode"" =T1.""U_ItemCode"" And  T22.""UomCode"" = (select ""UomCode"" from OUOM where ""UomEntry""=T1.""U_UOMCode"")and T00.""InvntItem""='Y')  As ""ActualQty"" " &
                                 " , T5.""InvntryUom"" As ""BaseUom"" From ""@ME_PSC"" T0 Inner Join ""@ME_PSC_C0"" T1 On T1.""DocEntry"" = T0.""DocEntry""  inner join OITM T5 on T5.""ItemCode"" = T1.""U_ItemCode"" " &
                                 " Where T1. ""U_ItemCode"" <> '' And T0.""U_CountDate"" Between  '" & objForm.Items.Item("5").Specific.Value & "' And " &
                                 " '" & objForm.Items.Item("7").Specific.Value & "' And   ""U_Branch"" = '" & objForm.Items.Item("15").Specific.Selected.Value & "' And ""U_WhseCode"" = '" & objForm.Items.Item("13").Specific.Value & "'"
                If objForm.Items.Item("17").Specific.Selected.Value = "All" Then
                    Query = Query + "And T0.""U_TYPE"" = 'All'  "
                ElseIf objForm.Items.Item("17").Specific.Selected.Value = "ItemGroup" Then
                    Query = Query + "And T0.""U_TYPE"" = 'ItemGroup' And T0.""U_ITMGRP"" = '" & objForm.Items.Item("19").Specific.Selected.Value.Trim & "'  "
                ElseIf objForm.Items.Item("17").Specific.Selected.Value = "SubGroup" Then
                    Query = Query + "And T0.""U_TYPE"" = 'SubGroup' And  T0.""U_ITMGRP"" = '" & objForm.Items.Item("19").Specific.Selected.Value.Trim & "' And T0.""U_SUBGRP"" = '" & objForm.Items.Item("21").Specific.Selected.Value.Trim & "' "
                ElseIf objForm.Items.Item("17").Specific.Selected.Value = "SubSubGroup" Then
                    Query = Query + "And  T0.""U_TYPE"" = 'SubSubGroup' And T0.""U_ITMGRP"" = '" & objForm.Items.Item("19").Specific.Selected.Value.Trim & "' And  T0.""U_SUBGRP"" = '" & objForm.Items.Item("21").Specific.Selected.Value.Trim & "' And T0.""U_SBSBGRP"" = '" & objForm.Items.Item("23").Specific.Selected.Value.Trim & "' "
                End If
                Query = Query + " ) Group By ""ItemCode"" , ""ItemName"" , ""BaseUom"""
                oGrid.DataTable.ExecuteQuery(Query)
            End If

            If IncludeExStockItems.Checked = True Then
                'Dim Query As String = "Select  '" & Value & "' as ""Check""  , ""ItemCode"" , ""ItemName"" ,  SUM(""Qty"") As ""Qty"" , ""BaseUom""  From ( "
                'Query = Query + " Select '" & Value & "' as ""Check""  , ""ItemCode"" , ""ItemName"" , SUM(""ActualQty"") As ""Qty"" , ""BaseUom""  From ( " &
                '                      " Select  '" & Value & "' as ""Check""  ,  ""U_ItemCode"" As ""ItemCode"" , ""U_ItemName"" As  ""ItemName"" , ""U_Qty"" As ""Quantity"" , ""U_UomName"" As ""UomCode"" , " &
                '                      " (Select (T11.""BaseQty"" * T1.""U_Qty"") From ""OITM"" T00  INNER JOIN ""UGP1"" T11 ON T00.""UgpEntry"" = T11.""UgpEntry"" INNER JOIN ""OUOM"" T22 ON T22.""UomEntry"" = T11.""UomEntry""  WHERE T00.""ItemCode"" =T1.""U_ItemCode"" And  T22.""UomCode"" = (select ""UomCode"" from OUOM where ""UomEntry""=T1.""U_UOMCode"")and T00.""InvntItem""='Y')  As ""ActualQty"" " &
                '                      " , ""U_BaseUom"" As ""BaseUom"" From ""@ME_PSC"" T0 Inner Join ""@ME_PSC_C0"" T1 On T1.""DocEntry"" = T0.""DocEntry"" " &
                '                      " Where  ""U_ItemCode"" <> '' And T0.""U_CountDate"" Between  '" & objForm.Items.Item("5").Specific.Value & "' And " &
                '                      " '" & objForm.Items.Item("7").Specific.Value & "' And   ""U_Branch"" = '" & objForm.Items.Item("15").Specific.Selected.Value & "' And ""U_WhseCode"" = '" & objForm.Items.Item("13").Specific.Value & "'"
                'If objForm.Items.Item("17").Specific.Selected.Value = "All" Then
                '    Query = Query + "And T0.""U_TYPE"" = 'All'  "
                'ElseIf objForm.Items.Item("17").Specific.Selected.Value = "ItemGroup" Then
                '    Query = Query + "And T0.""U_TYPE"" = 'ItemGroup' And T0.""U_ITMGRP"" = '" & objForm.Items.Item("19").Specific.Selected.Value.Trim & "'  "
                'ElseIf objForm.Items.Item("17").Specific.Selected.Value = "SubGroup" Then
                '    Query = Query + "And T0.""U_TYPE"" = 'SubGroup' And  T0.""U_ITMGRP"" = '" & objForm.Items.Item("19").Specific.Selected.Value.Trim & "' And T0.""U_SUBGRP"" = '" & objForm.Items.Item("21").Specific.Selected.Value.Trim & "' "
                'ElseIf objForm.Items.Item("17").Specific.Selected.Value = "SubSubGroup" Then
                '    Query = Query + "And  T0.""U_TYPE"" = 'SubSubGroup' And T0.""U_ITMGRP"" = '" & objForm.Items.Item("19").Specific.Selected.Value.Trim & "' And  T0.""U_SUBGRP"" = '" & objForm.Items.Item("21").Specific.Selected.Value.Trim & "' And T0.""U_SBSBGRP"" = '" & objForm.Items.Item("23").Specific.Selected.Value.Trim & "' "
                'End If
                'Query = Query + " ) Group By ""ItemCode"" , ""ItemName"" , ""BaseUom"""
                'Query = Query + " Union  Select  '" & Value & "' as ""Check"" ,  ""ItemCode"" , ""ItemName"" , 0 As ""Qty"" , ""InvntryUom"" As ""BaseUom""    From OITM    "
                'If objForm.Items.Item("17").Specific.Selected.Value = "ItemGroup" Then
                '    Query = Query + " Where OITM.""ItmsGrpCod""  = '" & objForm.Items.Item("19").Specific.Selected.Value.Trim & "'  "
                'ElseIf objForm.Items.Item("17").Specific.Selected.Value = "SubGroup" Then
                '    Query = Query + " Where  OITM.""ItmsGrpCod"" = '" & objForm.Items.Item("19").Specific.Selected.Value.Trim & "' And OITM.""U_VSubGroup"" = '" & objForm.Items.Item("21").Specific.Selected.Value.Trim & "' "
                'ElseIf objForm.Items.Item("17").Specific.Selected.Value = "SubSubGroup" Then
                '    Query = Query + "Where   OITM.""ItmsGrpCod"" = '" & objForm.Items.Item("19").Specific.Selected.Value.Trim & "' And  OITM.""U_VSubGroup"" = '" & objForm.Items.Item("21").Specific.Selected.Value.Trim & "' And OITM.""U_VSubSubGroup"" = '" & objForm.Items.Item("23").Specific.Selected.Value.Trim & "' "
                'End If
                'Query = Query + " )AA Group By ""ItemCode"" , ""ItemName"" , ""BaseUom"""

                Dim Query As String = "Select  '" & Value & "' as ""Check""  , ""ItemCode"" , ""ItemName"" ,  SUM(""Qty"") As ""Qty"" , ""BaseUom""  From ( "
                Query = Query + " Select '" & Value & "' as ""Check""  , ""ItemCode"" , ""ItemName"" , SUM(""ActualQty"") As ""Qty"" , ""BaseUom""  From ( " &
                                      " Select  '" & Value & "' as ""Check""  , T1. ""U_ItemCode"" As ""ItemCode"" ,T1. ""U_ItemName"" As  ""ItemName"" , ""U_Qty"" As ""Quantity"" , T1.""U_UomName"" As ""UomCode"" , " &
                                      " (Select (T11.""BaseQty"" * T1.""U_Qty"") From ""OITM"" T00  INNER JOIN ""UGP1"" T11 ON T00.""UgpEntry"" = T11.""UgpEntry"" INNER JOIN ""OUOM"" T22 ON T22.""UomEntry"" = T11.""UomEntry""  WHERE T00.""ItemCode"" =T1.""U_ItemCode"" And  T22.""UomCode"" =(select ""UomCode"" from OUOM where ""UomEntry""=T1.""U_UOMCode"")and T00.""InvntItem""='Y')  As ""ActualQty"" " &
                                      " , T5.""InvntryUom"" As ""BaseUom"" From ""@ME_PSC"" T0 Inner Join ""@ME_PSC_C0"" T1 On T1.""DocEntry"" = T0.""DocEntry""    inner join OITM T5 on T5.""ItemCode"" = T1.""U_ItemCode""" &
                                      " Where  T1.""U_ItemCode"" <> '' And T0.""U_CountDate"" Between  '" & objForm.Items.Item("5").Specific.Value & "' And " &
                                      " '" & objForm.Items.Item("7").Specific.Value & "' And   ""U_Branch"" = '" & objForm.Items.Item("15").Specific.Selected.Value & "' And ""U_WhseCode"" = '" & objForm.Items.Item("13").Specific.Value & "'"
                If objForm.Items.Item("17").Specific.Selected.Value = "All" Then
                    Query = Query + "And T0.""U_TYPE"" = 'All'  "
                ElseIf objForm.Items.Item("17").Specific.Selected.Value = "ItemGroup" Then
                    Query = Query + "And T0.""U_TYPE"" = 'ItemGroup' And T0.""U_ITMGRP"" = '" & objForm.Items.Item("19").Specific.Selected.Value.Trim & "'  "
                ElseIf objForm.Items.Item("17").Specific.Selected.Value = "SubGroup" Then
                    Query = Query + "And T0.""U_TYPE"" = 'SubGroup' And  T0.""U_ITMGRP"" = '" & objForm.Items.Item("19").Specific.Selected.Value.Trim & "' And T0.""U_SUBGRP"" = '" & objForm.Items.Item("21").Specific.Selected.Value.Trim & "' "
                ElseIf objForm.Items.Item("17").Specific.Selected.Value = "SubSubGroup" Then
                    Query = Query + "And  T0.""U_TYPE"" = 'SubSubGroup' And T0.""U_ITMGRP"" = '" & objForm.Items.Item("19").Specific.Selected.Value.Trim & "' And  T0.""U_SUBGRP"" = '" & objForm.Items.Item("21").Specific.Selected.Value.Trim & "' And T0.""U_SBSBGRP"" = '" & objForm.Items.Item("23").Specific.Selected.Value.Trim & "' "
                End If
                Query = Query + " ) Group By ""ItemCode"" , ""ItemName"" , ""BaseUom"""
                Query = Query + " Union  Select  '" & Value & "' as ""Check"" , T0. ""ItemCode"" ,T0. ""ItemName"" , 0 As ""Qty"" , T0.""InvntryUom"" As ""BaseUom""    From OITM  T0  inner join OITW T1 on T0.""ItemCode"" = T1.""ItemCode"" where T0.""InvntItem""='Y' and T1.""WhsCode""='" & objForm.Items.Item("13").Specific.Value & "' and T1.""OnHand""<>0 and T0. ""ItemCode"" Not in " & "
                (select T1.""U_ItemCode"" From ""@ME_PSC"" T0 Inner Join ""@ME_PSC_C0"" T1 On T1.""DocEntry"" = T0.""DocEntry"" Where T1.""U_ItemCode"" <> '' And T0.""U_CountDate"" Between '" & objForm.Items.Item("5").Specific.Value & "'
And '" & objForm.Items.Item("7").Specific.Value & "' And   ""U_Branch"" = '" & objForm.Items.Item("15").Specific.Selected.Value & "' And ""U_WhseCode"" = '" & objForm.Items.Item("13").Specific.Value & "')  "
                If objForm.Items.Item("17").Specific.Selected.Value = "ItemGroup" Then
                    Query = Query + " AND T0.""ItmsGrpCod""  = '" & objForm.Items.Item("19").Specific.Selected.Value.Trim & "'  "
                ElseIf objForm.Items.Item("17").Specific.Selected.Value = "SubGroup" Then
                    Query = Query + "AND  T0.""ItmsGrpCod"" = '" & objForm.Items.Item("19").Specific.Selected.Value.Trim & "' And T0.""U_VSubGroup"" = '" & objForm.Items.Item("21").Specific.Selected.Value.Trim & "' "
                ElseIf objForm.Items.Item("17").Specific.Selected.Value = "SubSubGroup" Then
                    Query = Query + "AND T0.""ItmsGrpCod"" = '" & objForm.Items.Item("19").Specific.Selected.Value.Trim & "' And  T0.""U_VSubGroup"" = '" & objForm.Items.Item("21").Specific.Selected.Value.Trim & "' And T0.""U_VSubSubGroup"" = '" & objForm.Items.Item("23").Specific.Selected.Value.Trim & "' "
                End If
                Query = Query + " )AA Group By ""ItemCode"" , ""ItemName"" , ""BaseUom"""
                oGrid.DataTable.ExecuteQuery(Query)
            End If
           

            If oGrid.Rows.Count = 0 Then
                objMain.objApplication.MessageBox("No records found")
                Exit Sub
            End If

            oGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item("ItemCode").Editable = False
            oGrid.Columns.Item("ItemName").Editable = False
            oGrid.Columns.Item("Qty").Editable = False
            oGrid.Columns.Item("BaseUom").Editable = False

            Dim oEditTextCol As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("ItemCode")
            oEditTextCol.LinkedObjectType = 4

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        Finally
            objForm.Freeze(False)
            objForm.Items.Item("24").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
        End Try
        oGrid.AutoResizeColumns()
    End Sub
    Sub UPDATETABLE(ByVal FormUID As String)
        'oGrid = objForm.Items.Item("3").Specific
        'Dim Branch As String = objForm.Items.Item("15").Specific.Selected.Value
        'Dim WhsCode As String = objForm.Items.Item("13").Specific.Value
        'For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
        '    If oGrid.DataTable.GetValue("Check", i) = "Y" Then

        '        Dim ItemCode As String = oGrid.DataTable.GetValue("ItemCode", i).ToString
        '        Dim ItemName As String = oGrid.DataTable.GetValue("ItemName", i).ToString
        '        Dim CountedQuantity As String = CDbl(oGrid.DataTable.GetValue("Qty", i))
        '        Dim PostDate As String = objForm.Items.Item("28").Specific.Value.ToString().Trim()
        '        PostDate = PostDate.Insert(4, "/")
        '        PostDate = PostDate.Insert(7, "/")
        '        Dim PDate As Date = CDate(PostDate)
        '        Dim str1 As String = oGrid.DataTable.GetValue("BaseUom", i).ToString
        '        Dim GetUOM As String = "Select ""UomCode"" from OUOM where ""UomName""='" & str1 & "'"
        '        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '        oRs.DoQuery(GetUOM)
        '        Dim UoMCode As String = oRs.Fields.Item(0).Value
        '        Dim chkIfExist As String = "Insert into ""SBO_STOCKTRANS"" (""BRANCH"",""WHSECODE"",""ITEMCODE"",""ITEMNAME"",""QUANTITY"",""BASEUOM"",""SYNC"",""POSTINGDATE"") values( '" & Branch & "','" & WhsCode & "','" & ItemCode & "','" & ItemName & "','" & CountedQuantity & "','" & UoMCode & "','1','" & PostDate & "')"

        '        'Dim chkIfExist As String = "UPDATE  ""SBO_STOCKTRANS"" SET ""BRANCH""='" & Branch & "',""WHSECODE""='" & WhsCode & "',""ITEMCODE""='" & ItemCode & "',""ITEMNAME""='" & ItemName & "',""QUANTITY""='" & CountedQuantity & "',""BASEUOM""='" & UoMCode & "',""SYNC""='1',""POSTINGDATE""='" & PostDate & "'"
        '        Dim oRschkIfExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '        oRschkIfExist.DoQuery(chkIfExist)
        '    End If
        'Next
        'objMain.objApplication.StatusBar.SetText("Updated successfully......!!.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success
        'Private Sub PostStockCount(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("3").Specific
            If objForm.Items.Item("28").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("posting date cannot be empty...")
                Exit Sub
            End If
            Dim PostDate As String = objForm.Items.Item("28").Specific.Value.ToString().Trim()
            PostDate = PostDate.Insert(4, "/")
            PostDate = PostDate.Insert(7, "/")
            Dim PDate As Date = CDate(PostDate)

            Dim oCS As SAPbobsCOM.CompanyService = objMain.objCompany.GetCompanyService()
            Dim oICS As SAPbobsCOM.InventoryCountingsService = oCS.GetBusinessService(SAPbobsCOM.ServiceTypes.InventoryCountingsService)
            Dim oIC As SAPbobsCOM.InventoryCounting = oICS.GetDataInterface(SAPbobsCOM.InventoryCountingsServiceDataInterfaces.icsInventoryCounting)
            Dim oICP As SAPbobsCOM.InventoryCountingParams
            Dim oICLS As SAPbobsCOM.InventoryCountingLines
            Dim oICL As SAPbobsCOM.InventoryCountingLine
            Dim WhsCode As String = objForm.Items.Item("13").Specific.Value
            oIC.CountDate = PostDate
            oIC.BranchID = objForm.Items.Item("15").Specific.Selected.Value
            '----
            Dim oCS1 As SAPbobsCOM.CompanyService = objMain.objCompany.GetCompanyService()
            Dim oICS1 As SAPbobsCOM.InventoryCountingsService = oCS1.GetBusinessService(SAPbobsCOM.ServiceTypes.InventoryCountingsService)
            Dim oIC1 As SAPbobsCOM.InventoryCounting = oICS1.GetDataInterface(SAPbobsCOM.InventoryCountingsServiceDataInterfaces.icsInventoryCounting)
            Dim oICP1 As SAPbobsCOM.InventoryCountingParams
            Dim oICLS1 As SAPbobsCOM.InventoryCountingLines
            Dim oICL1 As SAPbobsCOM.InventoryCountingLine

            ' ---
            For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                Dim count As String = 2
                If i < count Then
                    '  For j As Integer = 0 To 4
                    If oGrid.DataTable.GetValue("Check", i) = "Y" Then
                        oICLS = oIC.InventoryCountingLines
                        oICL = oICLS.Add()
                        oICL.ItemCode = oGrid.DataTable.GetValue("ItemCode", i).ToString
                        oICL.CountedQuantity = CDbl(oGrid.DataTable.GetValue("Qty", i))
                        oICL.WarehouseCode = WhsCode
                        Dim str1 As String = oGrid.DataTable.GetValue("BaseUom", i).ToString
                        Dim GetUOM As String = "Select ""UomCode"" from OUOM where ""UomName""='" & str1 & "'"
                        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRs.DoQuery(GetUOM)
                        oICL.UoMCode = oRs.Fields.Item(0).Value
                        oICL.Counted = SAPbobsCOM.BoYesNoEnum.tYES
                    End If
                ElseIf i = 2 Then
                    oICP = oICS.Add(oIC)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oICP)
                    objMain.objApplication.ActivateMenuItem("3082")
                    objICForm = objMain.objApplication.Forms.GetForm("1474000001", objMain.objApplication.Forms.ActiveForm.TypeCount)
                    Dim WhsCode1 As String = objForm.Items.Item("13").Specific.Value
                    oIC1.CountDate = PostDate
                    oIC1.BranchID = objForm.Items.Item("15").Specific.Selected.Value
                Else
                    If oGrid.DataTable.GetValue("Check", i) = "Y" Then
                        oICLS1 = oIC1.InventoryCountingLines
                        oICL1 = oICLS.Add()
                        oICL1.ItemCode = oGrid.DataTable.GetValue("ItemCode", i).ToString
                        oICL1.CountedQuantity = CDbl(oGrid.DataTable.GetValue("Qty", i))
                        oICL1.WarehouseCode = WhsCode
                        Dim str1 As String = oGrid.DataTable.GetValue("BaseUom", i).ToString
                        Dim GetUOM As String = "Select ""UomCode"" from OUOM where ""UomName""='" & str1 & "'"
                        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRs.DoQuery(GetUOM)

                        oICL1.UoMCode = oRs.Fields.Item(0).Value
                        'oICL.UoMCode = oGrid.DataTable.GetValue("BaseUom", i).ToString
                        oICL1.Counted = SAPbobsCOM.BoYesNoEnum.tYES
                        'oGrid.CommonSetting.SetRowFontColor(i + 1, ColorTranslator.ToOle(Color.DarkBlue))
                    End If
                    '  Next
                End If

            Next
            oICP1 = oICS1.Add(oIC)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oICP1)
            'objMain.objApplication.MessageBox("Inventory Counting Posted Successfully...!!!")

            Dim GetMaxDocNum As String = "Select ""DocNum"" From ""OINC"" Where ""DocEntry""=(Select MAX(""DocEntry"") From ""OINC"")"
            Dim oRsGetMaxDocNum As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetMaxDocNum.DoQuery(GetMaxDocNum)



        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        Finally
        End Try
        oGrid.AutoResizeColumns()
    End Sub


    Private Sub PostStockCount(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("3").Specific
            If objForm.Items.Item("28").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("posting date cannot be empty...")
                Exit Sub
            End If
            Dim PostDate As String = objForm.Items.Item("28").Specific.Value.ToString().Trim()
            PostDate = PostDate.Insert(4, "/")
            PostDate = PostDate.Insert(7, "/")
            Dim PDate As Date = CDate(PostDate)

            Dim oCS As SAPbobsCOM.CompanyService = objMain.objCompany.GetCompanyService()
            Dim oICS As SAPbobsCOM.InventoryCountingsService = oCS.GetBusinessService(SAPbobsCOM.ServiceTypes.InventoryCountingsService)
            Dim oIC As SAPbobsCOM.InventoryCounting = oICS.GetDataInterface(SAPbobsCOM.InventoryCountingsServiceDataInterfaces.icsInventoryCounting)
            Dim oICP As SAPbobsCOM.InventoryCountingParams
            Dim oICLS As SAPbobsCOM.InventoryCountingLines
            Dim oICL As SAPbobsCOM.InventoryCountingLine
            Dim WhsCode As String = objForm.Items.Item("13").Specific.Value
            oIC.CountDate = PostDate
            oIC.BranchID = objForm.Items.Item("15").Specific.Selected.Value
            For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If oGrid.DataTable.GetValue("Check", i) = "Y" Then
                    'If i = 0 Then
                    '    oIC.CountDate = PostDate
                    '    oIC.BranchID = objForm.Items.Item("15").Specific.Selected.Value
                    'End If                  
                    oICLS = oIC.InventoryCountingLines
                    oICL = oICLS.Add()
                    oICL.ItemCode = oGrid.DataTable.GetValue("ItemCode", i).ToString
                    oICL.CountedQuantity = CDbl(oGrid.DataTable.GetValue("Qty", i))
                    oICL.WarehouseCode = WhsCode
                    Dim str1 As String = oGrid.DataTable.GetValue("BaseUom", i).ToString
                    Dim GetUOM As String = "Select ""UomCode"" from OUOM where ""UomName""='" & str1 & "'"
                    Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRs.DoQuery(GetUOM)

                    oICL.UoMCode = oRs.Fields.Item(0).Value
                    'oICL.UoMCode = oGrid.DataTable.GetValue("BaseUom", i).ToString
                    oICL.Counted = SAPbobsCOM.BoYesNoEnum.tYES
                    'oGrid.CommonSetting.SetRowFontColor(i + 1, ColorTranslator.ToOle(Color.DarkBlue))
                End If
            Next
            oICP = oICS.Add(oIC)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oICP)
            'objMain.objApplication.MessageBox("Inventory Counting Posted Successfully...!!!")

            Dim GetMaxDocNum As String = "Select ""DocNum"" From ""OINC"" Where ""DocEntry""=(Select MAX(""DocEntry"") From ""OINC"")"
            Dim oRsGetMaxDocNum As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetMaxDocNum.DoQuery(GetMaxDocNum)

            Me.LoadInventoryCounting(objForm.UniqueID, oRsGetMaxDocNum.Fields.Item(0).Value)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)            
        Finally
        End Try
        oGrid.AutoResizeColumns()
    End Sub

    Private Sub LoadInventoryCounting(ByVal FormUID As String, ByVal DocNum As String)
        Try
            objMain.objApplication.ActivateMenuItem("3082")
            objICForm = objMain.objApplication.Forms.GetForm("1474000001", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objICForm.Freeze(True)
            objICForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            objICForm.Items.Item("1470000016").Specific.value = DocNum
            objICForm.Items.Item("1470000001").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            objICForm.Freeze(False)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
