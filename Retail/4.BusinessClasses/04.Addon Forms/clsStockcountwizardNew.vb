Imports System.Security.Cryptography

Public Class clsStockcountwizardNew
    'Stock count dana
#Region "       Declaration             "
    Public objForm, objITRForm, objICForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details1 As SAPbouiCOM.DBDataSource
    Dim objMatrix1 As SAPbouiCOM.Matrix
    Dim objComboBox, OBJCOMBO, OBJCOMBO1 As SAPbouiCOM.ComboBox
    Dim ItemGroup, ItemSubGroup, sQury2, sQury3 As String
    Dim LostFocusFlag As Boolean = False

#End Region
    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("Stockwz.xml", "ME_SCGW", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_SCGW", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_STOCKWZD")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_STOCKWZD_C0")
            objMatrix1 = objForm.Items.Item("Mtx").Specific
            objMatrix1.AutoResizeColumns()
            objMatrix1.Clear()
            SetDefault(objForm.UniqueID)
            'oDBs_Details1.Clear()
            'objMatrix1.FlushToDataSource()
            'Akhil
            'oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "SBO_STOCKWZDUDO", "Primary"))
            'oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, Me.GetFieldValue("Select Max(""DocNum"")+1 From ""@SBO_STOCKWZD"" "))

            '  objForm.DataSources.UserDataSources.Add("U_ITEMGRP", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            'objComboBox = objForm.Items.Item("ITEMGRP").Specific
            'objComboBox.DataBind.SetBound(True, "", "U_ITEMGRP")
            'objForm.Items.Item("ITEMGRP").AffectsFormMode = False
            'objMain.objUtilities.ComboBoxLoadValues(objComboBox, "Select ""ItmsGrpCod""  , ""ItmsGrpNam""  From OITB Order By  ""ItmsGrpNam""")
            'objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            Dim sQury1 As String = "Select ""ItmsGrpCod""  , ""ItmsGrpNam""  From OITB Order By  ""ItmsGrpNam"""
            objComboBox = objForm.Items.Item("ITEMGRP").Specific
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, sQury1)
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("ITEMGRP").DisplayDesc = True

            '    Dim sQury2 As String = "SELECT T0.""Code"", T0.""Name"" FROM ""@VR_SUBGROUP"" T0 Order By  ""Name"" "
            sQury2 = "select distinct  ""U_SUBGROUP"",""U_SUBGROUP"" FROM ""@VR_SUBGROUP"""
            objComboBox = objForm.Items.Item("SBGRP").Specific
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, sQury2)
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("SBGRP").DisplayDesc = True

            ' Dim sQury3 As String = "SELECT T0.""Code"", T0.""Name"" FROM ""@VR_SUBSUBGROUP"" T0 Order By  ""Name"" "
            sQury3 = "SELECT T0.""Name"", T0.""Name"" FROM ""@VR_SUBSUBGROUP"" T0 "
            objComboBox = objForm.Items.Item("SBSBGRP").Specific
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, sQury3)
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("SBSBGRP").DisplayDesc = True

            oDBs_Head.SetValue("U_FROMDT", 0, DateTime.Now.ToString("yyyyMMdd"))
            oDBs_Head.SetValue("U_TODT", 0, DateTime.Now.ToString("yyyyMMdd"))
            oDBs_Head.SetValue("U_PSDT", 0, DateTime.Now.ToString("yyyyMMdd"))

            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            ' objForm.Items.Item("PSDT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            '  objForm.Items.Item("PSDT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            ' objForm.Items.Item("PSDT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("INEXITM").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("INEXITM").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("INEXITM").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("SBGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("SBGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("SBGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("SBSBGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("SBSBGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("SBSBGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("TYPE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("TYPE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("TYPE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("ITEMGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("ITEMGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("ITEMGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("FROMDT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("FROMDT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("FROMDT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("TODT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("TODT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("TODT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("BRANCH").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("BRANCH").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("BRANCH").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("WHSE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("WHSE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("WHSE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objMatrix1.Columns.Item("Onhand").Visible = False

            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                objForm.Items.Item("25").Enabled = False
            End If

            Me.SetNewLine(objForm.UniqueID)
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "", Optional ByVal Series As Integer = 0)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_STOCKWZD")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_STOCKWZD_C0")
            objMatrix1 = objForm.Items.Item("Mtx").Specific
            'aa
            'oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "SBO_STOCKWZDUDO", "Primary"))



            Dim Query As String = "SELECT IFNULL(MAX(""DocNum""),0) + 1 AS ""NextDocNum"" FROM ""@SBO_STOCKWZD"""

            Dim rs As SAPbobsCOM.Recordset
            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            rs.DoQuery(Query)

            If Not rs.EoF Then
                Dim nextDocNum As Integer = rs.Fields.Item("NextDocNum").Value
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, nextDocNum)
                oDBs_Head.SetValue("U_PSDT", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))

            End If



            objComboBox = objForm.Items.Item("BRANCH").Specific
            Dim oRsUpdate2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim BPLId As String = " Select T0.""BPLId"",T1.""BPLName"",T1.""DflWhs"" From USR6 T0 Inner Join " &
                                                                 " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "'Order By T0.""BPLId""  "
            oRsUpdate2.DoQuery(BPLId)
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, BPLId)
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("BRANCH").DisplayDesc = True
            If objComboBox.ValidValues.Count <> 0 Then
                objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
            End If


            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_STOCKWZD")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_STOCKWZD_C0")
            objMatrix1 = objForm.Items.Item("Mtx").Specific

            objMatrix1.AddRow()
            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
            oDBs_Details1.SetValue("U_Check", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_ICode", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_IName", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_Qty", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_BaseUom", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_Update", oDBs_Details1.Offset, "")
            objMatrix1.SetLineData(objMatrix1.VisualRowCount)
            objMatrix1.AutoResizeColumns()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ME_SCGW" And pVal.BeforeAction = False Then
                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
                Me.SetNewLine(objForm.UniqueID)
            End If
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Function GetFieldValue(ByVal query As String) As Object
        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs.DoQuery(query)
        Dim str As String = oRs.Fields.Item(0).Value.ToString()
        Return str
    End Function
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    ' oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_STOCKWZD")
                    '  oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_STOCKWZD_C0")
                    ' If pVal.ItemUID = "btnOk" And pVal.BeforeAction = False Then
                    'Me.FileGeneration(pVal)
                    ' End If

                    '                    If pVal.ItemUID = "Item_6" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or
                    'pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                    '                        Dim AppTemForm_EMD As SAPbouiCOM.Form
                    '                        objMain.objApplication.ActivateMenuItem("3082")
                    '                        AppTemForm_EMD = objMain.objApplication.Forms.GetFormByTypeAndCount("1474000001", 1)
                    '                        AppTemForm_EMD = objMain.objApplication.Forms.ActiveForm
                    '                        AppTemForm_EMD.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    '                        Dim Doc As String = objForm.Items.Item("INCNT").Specific.value
                    '                        AppTemForm_EMD.Items.Item("DocNum").Specific.Value = Doc
                    '                        AppTemForm_EMD.Items.Item("1470000001").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    '                        AppTemForm_EMD = objMain.objApplication.Forms.ActiveForm

                    '                    End If

                    If pVal.ItemUID = "Item_6" And pVal.BeforeAction = False And
   (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or
    pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or
    pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then

                        objMain.objApplication.ActivateMenuItem("3082") 'Inventory Counting

                        Dim AppTemForm_EMD As SAPbouiCOM.Form
                        AppTemForm_EMD = objMain.objApplication.Forms.ActiveForm
                        AppTemForm_EMD = objMain.objApplication.Forms.GetFormByTypeAndCount("1474000001", 1)

                        AppTemForm_EMD.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE

                        Dim Doc As String = objForm.Items.Item("INCNT").Specific.Value.ToString.Trim
                        AppTemForm_EMD.Items.Item("1470000016").Specific.Value = Doc

                        'Click Find button
                        AppTemForm_EMD.Items.Item("1470000001").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                    End If












                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                        Me.SetDefault(objForm.UniqueID)
                        Me.SetNewLine(objForm.UniqueID)



                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Validation(objForm.UniqueID) = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                    'If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True Then
                    '    If objForm.FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then objForm.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                    '    Me.SetDefault(objForm.UniqueID)
                    '    End If

                    If pVal.ItemUID = "24" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("FROMDT").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        ElseIf objForm.Items.Item("TODT").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            Me.LoadGrid(objForm.UniqueID, "Y")
                        End If
                    End If
                    If pVal.ItemUID = "25" And pVal.BeforeAction = False Then
                        Me.PostStockCount(objForm.UniqueID)
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
                    If oCFL.UniqueID = "CFL_WHS" And pVal.BeforeAction = True Then
                        Me.CFLFilter(FormUID, "CFL_WHS")
                    End If

                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        If oCFL.UniqueID = "CFL_WHS" Then
                            oDBs_Head.SetValue("U_WHSE", oDBs_Head.Offset, oDT.GetValue("WhsCode", 0))
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_STOCKWZD")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_STOCKWZD_C0")
                    objMatrix1 = objForm.Items.Item("Mtx").Specific
                    If pVal.ItemUID = "Mtx" And pVal.ColUID = "ICode" Then
                        If objMatrix1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim() <> "" Then
                            If pVal.Row = objMatrix1.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID)
                            End If
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_STOCKWZD")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_STOCKWZD_C0")
                    If objForm.Items.Item("ITEMGRP").Specific.Value <> "" Then
                        OBJCOMBO = objForm.Items.Item("ITEMGRP").Specific
                        ItemGroup = OBJCOMBO.Selected.Value
                        sQury2 = "select distinct  ""U_SUBGROUP"",""U_SUBGROUP"" FROM ""@VR_SUBGROUP"" where ""U_ITEMGROUP""='" & ItemGroup & "'"
                        objComboBox = objForm.Items.Item("SBGRP").Specific
                        objMain.objUtilities.ComboBoxLoadValues(objComboBox, sQury2)
                        objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                        objForm.Items.Item("SBGRP").DisplayDesc = True
                    End If
                    If objForm.Items.Item("SBGRP").Specific.Value <> "" Then
                        OBJCOMBO1 = objForm.Items.Item("SBGRP").Specific
                        ItemSubGroup = OBJCOMBO1.Selected.Value
                        sQury3 = "SELECT T0.""Name"", T0.""Name"" FROM ""@VR_SUBSUBGROUP"" T0  where ""U_SUBGROUP""='" & ItemSubGroup & "' "
                        objComboBox = objForm.Items.Item("SBSBGRP").Specific
                        objMain.objUtilities.ComboBoxLoadValues(objComboBox, sQury3)
                        objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                        objForm.Items.Item("SBSBGRP").DisplayDesc = True
                    End If
                    If pVal.ItemUID = "TYPE" And pVal.BeforeAction = False Then
                        Select Case objForm.Items.Item("TYPE").Specific.Selected.Value
                            Case "All"
                                objForm.Items.Item("ITEMGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("ITEMGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objComboBox = objForm.Items.Item("ITEMGRP").Specific
                                objComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                objForm.Items.Item("SBGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("SBGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objComboBox = objForm.Items.Item("SBGRP").Specific
                                objComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                objForm.Items.Item("SBSBGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("SBSBGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objComboBox = objForm.Items.Item("SBSBGRP").Specific
                                objComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                            Case "ItemGroup"
                                objForm.Items.Item("ITEMGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("SBGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("SBGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("SBSBGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("SBSBGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            Case "ItemSubGroup"
                                objForm.Items.Item("ITEMGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("SBGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("SBSBGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("SBSBGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            Case "ItemSubSubGroup"
                                objForm.Items.Item("ITEMGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("SBGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("SBSBGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        End Select
                    End If

            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub PostStockCount(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix1 = objForm.Items.Item("Mtx").Specific
            Dim Count As String = "Select Count(*) from ""@SBO_STOCKWZD_C0"" T1 Inner join  ""@SBO_STOCKWZD"" T2 on T2.""DocEntry""=
T1.""DocEntry"" Where ""U_Check""='Y' And ""DocNum""='" & objForm.Items.Item("DocNum").Specific.Value & "'"

            Dim Count3 As String = "Select Count(*) from ""@SBO_STOCKWZD_C0"" T1 Inner join  ""@SBO_STOCKWZD"" T2 on T2.""DocEntry""=
T1.""DocEntry"" Where ""U_Check""='Y'   and ""U_Update""='Y' And ""DocNum""='" & objForm.Items.Item("DocNum").Specific.Value & "'"
            Dim oRsCunt3 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsCunt3.DoQuery(Count3)
            Dim CheckCount3 As String = oRsCunt3.Fields.Item(0).Value

            If CheckCount3 = 4999 Then
                '     If CheckCount3 = 4 Then
                Me.PostStockCount1(objForm.UniqueID)
            Else


                '  Dim CheckCount As String = oRsCunt.Fields.Item(0).Value
                If objForm.Items.Item("PSDT").Specific.Value = "" Then
                    objMain.objApplication.StatusBar.SetText("posting date cannot be empty...")
                    Exit Sub
                End If
                Dim PostDate As String = objForm.Items.Item("PSDT").Specific.Value.ToString().Trim()
                PostDate = PostDate.Insert(4, "/")
                PostDate = PostDate.Insert(7, "/")
                Dim PDate As Date = CDate(PostDate)
                Dim oCS As SAPbobsCOM.CompanyService = objMain.objCompany.GetCompanyService()
                Dim oICS As SAPbobsCOM.InventoryCountingsService = oCS.GetBusinessService(SAPbobsCOM.ServiceTypes.InventoryCountingsService)
                Dim oIC As SAPbobsCOM.InventoryCounting = oICS.GetDataInterface(SAPbobsCOM.InventoryCountingsServiceDataInterfaces.icsInventoryCounting)
                Dim oICP As SAPbobsCOM.InventoryCountingParams
                Dim oICLS As SAPbobsCOM.InventoryCountingLines
                Dim oICL As SAPbobsCOM.InventoryCountingLine
                Dim WhsCode As String = objForm.Items.Item("WHSE").Specific.Value
                oIC.CountDate = PostDate
                oIC.BranchID = objForm.Items.Item("BRANCH").Specific.Selected.Value
                For i As Integer = 1 To objMatrix1.VisualRowCount
                    If 4999 >= i Then
                        '     If 4 >= i Then
                        Dim checkBox As SAPbouiCOM.CheckBox = DirectCast(Me.objMatrix1.Columns.Item("Check").Cells.Item(i).Specific, SAPbouiCOM.CheckBox)
                        Dim Check As Boolean
                        Dim checkBox1 As SAPbouiCOM.CheckBox = DirectCast(Me.objMatrix1.Columns.Item("Update").Cells.Item(i).Specific, SAPbouiCOM.CheckBox)
                        Dim check1 As Boolean
                        If (checkBox.Checked) And (Not checkBox1.Checked) Then

                            Check = True

                            oICLS = oIC.InventoryCountingLines
                            oICL = oICLS.Add()
                            Dim str2 As String = objMatrix1.Columns.Item("ICode").Cells.Item(i).Specific.Value
                            oICL.ItemCode = objMatrix1.Columns.Item("ICode").Cells.Item(i).Specific.Value
                            Dim str3 As String = objMatrix1.Columns.Item("Qty").Cells.Item(i).Specific.Value
                            Dim Status As String = objMatrix1.Columns.Item("Flag").Cells.Item(i).Specific.Value

                            '  If str3 <= 0 Then
                            If Status = "System" Then
                                oICL.CountedQuantity = 0
                            Else
                                oICL.CountedQuantity = CDbl(objMatrix1.Columns.Item("Qty").Cells.Item(i).Specific.Value)
                            End If
                            oICL.WarehouseCode = WhsCode
                            Dim str1 As String = objMatrix1.Columns.Item("BaseUom").Cells.Item(i).Specific.Value
                            Dim GetUOM As String = "Select ""UomCode"" from OUOM where ""UomName""='" & str1 & "'"
                            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRs.DoQuery(GetUOM)
                            oICL.UoMCode = oRs.Fields.Item(0).Value
                            oICL.Counted = SAPbobsCOM.BoYesNoEnum.tYES
                            Dim str As String = objMatrix1.Columns.Item("Qty").Cells.Item(i).Specific.Value
                            objMatrix1.Columns.Item("Update").Cells.Item(i).Specific.Checked = True
                            'oDBs_Details1.SetValue("Update", oDBs_Details1.Offset, "Y")

                        Else

                        End If
                    End If
                Next
                oICP = oICS.Add(oIC)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oICP)
                objMain.objApplication.MessageBox("Inventory Counting Posted Successfully...!!!")



                'inventory doc num refreence
                Dim GetMaxDocNum1 As String = "Select ""DocNum"" From ""OINC"" Where ""DocEntry""=(Select MAX(""DocEntry"") From ""OINC"")"
                Dim oRsGetMaxDocNum1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetMaxDocNum1.DoQuery(GetMaxDocNum1)
                'Me.LoadInventoryCounting(objForm.UniqueID, oRsGetMaxDocNum1.Fields.Item(0).Value)
                'Akila
                objForm.Items.Item("INCNT").Specific.Value = oRsGetMaxDocNum1.Fields.Item("DocNum").Value



                'Dim GetMaxDocNum1 As String = "Select ""DocNum"" From ""OINC"" Where ""DocEntry""=(Select MAX(""DocEntry"") From ""OINC"")"
                'Dim oRsGetMaxDocNum1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oRsGetMaxDocNum1.DoQuery(GetMaxDocNum1)
                'Me.LoadInventoryCounting(objForm.UniqueID, oRsGetMaxDocNum1.Fields.Item(0).Value)
                Me.PostStockCount1(objForm.UniqueID)
            End If
            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            objForm.Items.Item("25").Enabled = False
            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
            objMatrix1.Columns.Item("Check").Editable = False
            objMatrix1.Columns.Item("ICode").Editable = False
            objMatrix1.Columns.Item("IName").Editable = False
            objMatrix1.Columns.Item("Qty").Editable = False
            objMatrix1.Columns.Item("BaseUom").Editable = False
            objMatrix1.Columns.Item("Update").Editable = False

            objForm.Items.Item("PSDT").Enabled = False
            objForm.Items.Item("INEXITM").Enabled = False

            '-----------------------------------------
            Dim Docnum As String = objForm.Items.Item("DocNum").Specific.Value
            ' Define the SQL query as a string
            Dim PDTDoc As String
            'PDTDoc = "UPDATE ""@ME_PSC_C0"" T0 " &
            '    "SET T0.""U_POST"" = 'Posted' " &
            '    "WHERE T0.""DocEntry"" IN ( " &
            '    "    SELECT T2.""DocEntry"" " &
            '    "    FROM ""@SBO_STOCKWZD"" T0 " &
            '    "    INNER JOIN ""@SBO_STOCKWZD_C0"" T1 " &
            '    "        ON T1.""DocEntry"" = T0.""DocEntry"" " &
            '    "    INNER JOIN ""@ME_PSC"" T2 " &
            '    "        ON T1.""U_DOCE"" = T2.""DocNum"" " &
            '    "    WHERE T0.""DocNum"" = '" & Docnum & "'" &
            '             "AND IFNULL(T0.""U_ItemName"", '') <> '' " &
            '    ")"
            PDTDoc = "UPDATE ""@ME_PSC_C0"" T0 " &
         "SET T0.""U_POST"" = 'Posted' " &
         "WHERE T0.""DocEntry"" IN ( " &
         "    SELECT T2.""DocEntry"" " &
         "    FROM ""@SBO_STOCKWZD"" S0 " &
         "    INNER JOIN ""@SBO_STOCKWZD_C0"" S1 " &
         "        ON S1.""DocEntry"" = S0.""DocEntry"" " &
         "    INNER JOIN ""@ME_PSC"" T2 " &
         "        ON S1.""U_DOCE"" = T2.""DocNum"" " &
         "    WHERE S0.""DocNum"" = '" & Docnum & "'" &
         ") " &
         "AND IFNULL(T0.""U_ItemName"", '') <> '' " &
         "AND IFNULL(T0.""U_Qty"", 0) > 0" &
          "  AND IFNULL(T0.""U_Barcode"", '') <> '' "


            Dim oRsGetMaxDocNum As SAPbobsCOM.Recordset
            oRsGetMaxDocNum = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetMaxDocNum.DoQuery(PDTDoc)

            objMain.objApplication.ActivateMenuItem("1304")


            '--------------------------------------------------------



        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        Finally
        End Try
        objMatrix1.AutoResizeColumns()
    End Sub
    Public Sub PostStockCount1(ByVal FormUID As String)

        Try
            Dim WhsCode As String = objForm.Items.Item("WHSE").Specific.Value
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix1 = objForm.Items.Item("Mtx").Specific
            Dim Count1 As String = "Select Count(*) from ""@SBO_STOCKWZD_C0"" T1 Inner join  ""@SBO_STOCKWZD"" T2 on T2.""DocEntry""=
            T1.""DocEntry"" Where ""U_Check""='Y' And ""DocNum""='" & objForm.Items.Item("DocNum").Specific.Value & "'"
            Dim oRsCunt1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsCunt1.DoQuery(Count1)
            Dim CheckCount1 As String = oRsCunt1.Fields.Item(0).Value
            If CheckCount1 >= 5000 Then
                '     If CheckCount1 > 4 Then
                If objForm.Items.Item("PSDT").Specific.Value = "" Then
                    objMain.objApplication.StatusBar.SetText("posting date cannot be empty...")
                    Exit Sub
                End If
                Dim PostDate1 As String = objForm.Items.Item("PSDT").Specific.Value.ToString().Trim()
                PostDate1 = PostDate1.Insert(4, "/")
                PostDate1 = PostDate1.Insert(7, "/")
                Dim PDate1 As Date = CDate(PostDate1)
                Dim oCS1 As SAPbobsCOM.CompanyService = objMain.objCompany.GetCompanyService()
                Dim oICS1 As SAPbobsCOM.InventoryCountingsService = oCS1.GetBusinessService(SAPbobsCOM.ServiceTypes.InventoryCountingsService)
                Dim oIC1 As SAPbobsCOM.InventoryCounting = oICS1.GetDataInterface(SAPbobsCOM.InventoryCountingsServiceDataInterfaces.icsInventoryCounting)
                Dim oICP1 As SAPbobsCOM.InventoryCountingParams
                Dim oICLS1 As SAPbobsCOM.InventoryCountingLines
                Dim oICL1 As SAPbobsCOM.InventoryCountingLine
                Dim WhsCode1 As String = objForm.Items.Item("WHSE").Specific.Value
                oIC1.CountDate = PostDate1
                oIC1.BranchID = objForm.Items.Item("BRANCH").Specific.Selected.Value
                For i As Integer = 1 To objMatrix1.VisualRowCount

                    'If i = 9 Or i = 10 Then
                    If i > 4999 Then
                        ' If i > 4 Then
                        Dim checkBox As SAPbouiCOM.CheckBox = DirectCast(Me.objMatrix1.Columns.Item("Check").Cells.Item(i).Specific, SAPbouiCOM.CheckBox)
                        Dim Check As Boolean
                        Dim checkBox1 As SAPbouiCOM.CheckBox = DirectCast(Me.objMatrix1.Columns.Item("Update").Cells.Item(i).Specific, SAPbouiCOM.CheckBox)
                        Dim check1 As Boolean
                        If (checkBox.Checked) And (Not checkBox1.Checked) Then
                            Check = True
                            oICLS1 = oIC1.InventoryCountingLines
                            oICL1 = oICLS1.Add()
                            Dim str2 As String = objMatrix1.Columns.Item("ICode").Cells.Item(i).Specific.Value
                            oICL1.ItemCode = objMatrix1.Columns.Item("ICode").Cells.Item(i).Specific.Value
                            ' oICL1.CountedQuantity = CDbl(objMatrix1.Columns.Item("Qty").Cells.Item(i).Specific.Value)
                            Dim str3 As String = objMatrix1.Columns.Item("Qty").Cells.Item(i).Specific.Value
                            Dim Status As String = objMatrix1.Columns.Item("Flag").Cells.Item(i).Specific.Value
                            'If str3 <= 0 Then
                            If Status = "System" Then
                                oICL1.CountedQuantity = 0
                            Else
                                oICL1.CountedQuantity = CDbl(objMatrix1.Columns.Item("Qty").Cells.Item(i).Specific.Value)
                            End If

                            oICL1.WarehouseCode = WhsCode
                            Dim str1 As String = objMatrix1.Columns.Item("BaseUom").Cells.Item(i).Specific.Value
                            Dim GetUOM As String = "Select ""UomCode"" from OUOM where ""UomName""='" & str1 & "'"
                            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRs.DoQuery(GetUOM)
                            oICL1.UoMCode = oRs.Fields.Item(0).Value
                            oICL1.Counted = SAPbobsCOM.BoYesNoEnum.tYES
                            objMatrix1.Columns.Item("Update").Cells.Item(i).Specific.Checked = True
                        End If
                    End If

                Next
                oICP1 = oICS1.Add(oIC1)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oICP1)
                objMain.objApplication.MessageBox("Inventory Counting Posted Successfully...!!!")
                Dim GetMaxDocNum1 As String = "Select ""DocNum"" From ""OINC"" Where ""DocEntry""=(Select MAX(""DocEntry"") From ""OINC"")"
                Dim oRsGetMaxDocNum1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetMaxDocNum1.DoQuery(GetMaxDocNum1)
                Me.LoadInventoryCounting(objForm.UniqueID, oRsGetMaxDocNum1.Fields.Item(0).Value)
                'Akila
                objForm.Items.Item("INCNT").Specific.Value = oRsGetMaxDocNum1.Fields.Item("DocNum").Value



            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        Finally
        End Try
        objMatrix1.AutoResizeColumns()
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
        oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_STOCKWZD")
        oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_STOCKWZD_C0")
        objMatrix1 = objForm.Items.Item("Mtx").Specific

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix1 = objForm.Items.Item("Mtx").Specific
            If objForm.Items.Item("PSDT").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("posting date cannot be empty...")
                Exit Sub
            End If
            Dim PostDate As String = objForm.Items.Item("PSDT").Specific.Value.ToString().Trim()
            PostDate = PostDate.Insert(4, "/")
            PostDate = PostDate.Insert(7, "/")
            Dim PDate As Date = CDate(PostDate)

            Dim oCS As SAPbobsCOM.CompanyService = objMain.objCompany.GetCompanyService()
            Dim oICS As SAPbobsCOM.InventoryCountingsService = oCS.GetBusinessService(SAPbobsCOM.ServiceTypes.InventoryCountingsService)
            Dim oIC As SAPbobsCOM.InventoryCounting = oICS.GetDataInterface(SAPbobsCOM.InventoryCountingsServiceDataInterfaces.icsInventoryCounting)
            Dim oICP As SAPbobsCOM.InventoryCountingParams
            Dim oICLS As SAPbobsCOM.InventoryCountingLines
            Dim oICL As SAPbobsCOM.InventoryCountingLine
            Dim WhsCode As String = objForm.Items.Item("WHSE").Specific.Value
            oIC.CountDate = PostDate
            oIC.BranchID = objForm.Items.Item("BRANCH").Specific.Selected.Value
            '----
            Dim oCS1 As SAPbobsCOM.CompanyService = objMain.objCompany.GetCompanyService()
            Dim oICS1 As SAPbobsCOM.InventoryCountingsService = oCS1.GetBusinessService(SAPbobsCOM.ServiceTypes.InventoryCountingsService)
            Dim oIC1 As SAPbobsCOM.InventoryCounting = oICS1.GetDataInterface(SAPbobsCOM.InventoryCountingsServiceDataInterfaces.icsInventoryCounting)
            Dim oICP1 As SAPbobsCOM.InventoryCountingParams
            Dim oICLS1 As SAPbobsCOM.InventoryCountingLines
            Dim oICL1 As SAPbobsCOM.InventoryCountingLine

            ' ---
            For i As Integer = 0 To objMatrix1.VisualRowCount - 1
                Dim count As String = 2
                If i < count Then
                    '  For j As Integer = 0 To 4
                    If objMatrix1.Item.GetValue("Check", i) = "Y" Then
                        oICLS = oIC.InventoryCountingLines
                        oICL = oICLS.Add()
                        oICL.ItemCode = objMatrix1.Item.GetValue("ItemCode", i).ToString
                        oICL.CountedQuantity = CDbl(objMatrix1.DataTable.GetValue("Qty", i))
                        oICL.WarehouseCode = WhsCode
                        Dim str1 As String = objMatrix1.DataTable.GetValue("BaseUom", i).ToString
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
                    Dim WhsCode1 As String = objForm.Items.Item("WHSE").Specific.Value
                    oIC1.CountDate = PostDate
                    oIC1.BranchID = objForm.Items.Item("BRANCH").Specific.Selected.Value
                Else
                    If objMatrix1.DataTable.GetValue("Check", i) = "Y" Then
                        oICLS1 = oIC1.InventoryCountingLines
                        oICL1 = oICLS.Add()
                        oICL1.ItemCode = objMatrix1.DataTable.GetValue("ItemCode", i).ToString
                        oICL1.CountedQuantity = CDbl(objMatrix1.DataTable.GetValue("Qty", i))
                        oICL1.WarehouseCode = WhsCode
                        Dim str1 As String = objMatrix1.Item.GetValue("BaseUom", i).ToString
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
        objMatrix1.AutoResizeColumns()
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
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("ME_SCGW", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            For i As Integer = 1 To objMatrix1.VisualRowCount
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                Dim checkBox As SAPbouiCOM.CheckBox = DirectCast(Me.objMatrix1.Columns.Item("Check").Cells.Item(i).Specific, SAPbouiCOM.CheckBox)
                                Dim Check As Boolean
                                Dim checkBox1 As SAPbouiCOM.CheckBox = DirectCast(Me.objMatrix1.Columns.Item("Update").Cells.Item(i).Specific, SAPbouiCOM.CheckBox)
                                Dim check1 As Boolean
                                If (checkBox.Checked) And (checkBox1.Checked) Then
                                    objMatrix1.Columns.Item("Check").Editable = False
                                    objMatrix1.Columns.Item("Update").Editable = False
                                    objForm.Items.Item("PSDT").Enabled = False
                                    objForm.Items.Item("INEXITM").Enabled = False
                                    'objForm.Items.Item("INCNT").Enabled = True
                                    Try
                                        objForm.Items.Item("Item_6").Enabled = True
                                    Catch ex As Exception

                                    End Try




                                End If
                            Next
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_STOCKWZD")
                            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_STOCKWZD_C0")

                            'Dim invDoc As String = oDBs_Head.GetValue("U_INCNT", oDBs_Head.Offset)
                            'objForm.Items.Item("U_INCNT").Specific.Value
                            Dim invDoc As String = ""

                            invDoc = oDBs_Head.GetValue("U_INCNT", oDBs_Head.Offset)

                            If invDoc.Trim() <> "" Then
                                objForm.Items.Item("25").Enabled = False
                            Else
                                objForm.Items.Item("25").Enabled = True
                            End If



                            'Dim invDoc As String = objForm.Items.Item("INCNT").Specific.Value

                            '' If document is already posted
                            'If invDoc.Trim() <> "" Then
                            '    ' Disable POST button
                            '    objForm.Items.Item("25").Enabled = False
                            'End If









                        End If

                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub CFLFilter(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim Selection As String = "select ""WhsCode"",""WhsName"",""BPLid"" from OWHS where ""BPLid""='" & objForm.Items.Item("BRANCH").Specific.Value & "'"
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
    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            If objForm.Items.Item("PSDT").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Please select the  Posting  Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True

            '    If objForm.Items.Item("Code").Specific.Value = "" Then
            '        objMain.objApplication.StatusBar.SetText("Please Enter Task Code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '        Return False
            '    ElseIf objForm.Items.Item("Name").Specific.Value = "" Then
            '        objMain.objApplication.StatusBar.SetText("Please Enter Task Name", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '        Return False
            '    ElseIf objForm.Items.Item("Itemtyp").Specific.Value = "" Then
            '        objMain.objApplication.StatusBar.SetText("Please Select Item Type", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '        Return False
            '        'ElseIf objForm.Items.Item("Icode").Specific.Value = "" Then
            '        '    objMain.objApplication.StatusBar.SetText("Please Select Item Code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '        '    Return False
            '        'ElseIf objForm.Items.Item("GLAct").Specific.Value = "" Then
            '        '    objMain.objApplication.StatusBar.SetText("Please Select G/L Account", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '        '    Return False
            '        'ElseIf objForm.Items.Item("Typelaw").Specific.Value = "" Then
            '        '    objMain.objApplication.StatusBar.SetText("Please Select Type of law ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '        '    Return False
            '        'ElseIf objForm.Items.Item("stage").Specific.Value = "" Then
            '        '    objMain.objApplication.StatusBar.SetText("Please Select Stage Code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '        '    Return False
            '        'ElseIf objForm.Items.Item("Code").Specific.Value = "" Then
            '        '    objMain.objApplication.StatusBar.SetText("Please Enter Code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '        '    Return False
            '        'ElseIf objForm.Items.Item("Name").Specific.Value = "" Thsen
            '        '    objMain.objApplication.StatusBar.SetText("Please Enter Name", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '        '    Return False
            '        'ElseIf objForm.Items.Item("Rate").Specific.Value = "" Then
            '        '    objMain.objApplication.StatusBar.SetText("Please Enter Rate", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '        '    Return False
            '    End If
            '    Return True

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Sub LoadGrid(ByVal FormUID As String, ByVal Value As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_STOCKWZD")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_STOCKWZD_C0")
            objMatrix1 = objForm.Items.Item("Mtx").Specific
            ' oDBs_Details1.Clear()
            ' objMatrix1.Clear()
            objMatrix1.FlushToDataSource()
            '  objForm.Refresh()
            Dim IncludeExStockItems As SAPbouiCOM.CheckBox = objForm.Items.Item("INEXITM").Specific
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
                'Akhila
                'Dim rs1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'Dim Query1 As String = " Select  '" & Value & "' as ""Check""   ,""U_POST"" ,""DocNum"", ""ItemCode"" , ""ItemName"" , SUM(""ActualQty"") As ""Qty"" , ""BaseUom"" , ""Status""
                '             ,""OnHand"" From ( " &
                '                 " Select    T1.""U_ItemCode"" As ""ItemCode"" , T1.""U_ItemName"" As  ""ItemName"" , ""U_Qty"" As ""Quantity"" , T1.""U_UomName"" As ""UomCode"" , " &
                '                 " (Select (T11.""BaseQty"" * T1.""U_Qty"") From ""OITM"" T00  INNER JOIN ""UGP1"" T11 ON T00.""UgpEntry"" = T11.""UgpEntry"" INNER JOIN ""OUOM"" T22 ON T22.""UomEntry"" = T11.""UomEntry""  WHERE T00.""ItemCode"" =T1.""U_ItemCode"" And  T22.""UomCode"" = (select ""UomCode"" from OUOM where ""UomEntry""=T1.""U_UOMCode"")and T00.""InvntItem""='Y')  As ""ActualQty"" " &
                '                 " , T5.""InvntryUom"" As ""BaseUom"" ,'PDT' as ""Status""
                '                   ,(SELECT W.""OnHand"" FROM OITW W WHERE W.""ItemCode""=T5.""ItemCode"" and W.""WhsCode""='" & objForm.Items.Item("WHSE").Specific.Value & "') As ""OnHand"" ,T0.""DocNum"" As ""DocNum"" From ""@ME_PSC"" T0 Inner Join ""@ME_PSC_C0"" T1 On T1.""DocEntry"" = T0.""DocEntry""  inner join OITM T5 on T5.""ItemCode"" = T1.""U_ItemCode"" " &
                '                 " Where T1. ""U_ItemCode"" <> '' and (T1.""U_POST"" IS NULL OR T1.""U_POST"" = '') as ""U_POST"" and  T5.""InvntItem""='Y'  And T0.""U_CountDate"" Between  '" & objForm.Items.Item("FROMDT").Specific.Value & "' And " &
                '                 " '" & objForm.Items.Item("TODT").Specific.Value & "' And T0.""U_Branch"" = '" & objForm.Items.Item("BRANCH").Specific.Selected.Value & "' And T0. ""U_WhseCode"" = '" & objForm.Items.Item("WHSE").Specific.Value & "'"
                'If objForm.Items.Item("TYPE").Specific.Selected.Value = "All" Then
                '    Query1 = Query1 + "And T0.""U_TYPE"" = 'All'  "
                'ElseIf objForm.Items.Item("TYPE").Specific.Selected.Value = "ItemGroup" Then
                '    Query1 = Query1 + "And T0.""U_TYPE"" = 'ItemGroup' And T0.""U_ITMGRP"" = '" & objForm.Items.Item("ITEMGRP").Specific.Selected.Value.Trim & "'  "
                'ElseIf objForm.Items.Item("TYPE").Specific.Selected.Value = "ItemSubGroup" Then
                '    Query1 = Query1 + "And T0.""U_TYPE"" = 'ItemSubGroup' And  T0.""U_ITMGRP"" = '" & objForm.Items.Item("ITEMGRP").Specific.Selected.Value.Trim & "' And T0.""U_SUBGRP"" = '" & objForm.Items.Item("SBGRP").Specific.Selected.Value.Trim & "' "
                'ElseIf objForm.Items.Item("TYPE").Specific.Selected.Value = "ItemSubSubGroup" Then
                '    Query1 = Query1 + "And  T0.""U_TYPE"" = 'ItemSubSubGroup' And T0.""U_ITMGRP"" = '" & objForm.Items.Item("ITEMGRP").Specific.Selected.Value.Trim & "' And  T0.""U_SUBGRP"" = '" & objForm.Items.Item("SBGRP").Specific.Selected.Value.Trim & "' And T0.""U_SBSBGRP"" = '" & objForm.Items.Item("SBSBGRP").Specific.Selected.Value.Trim & "' "
                'End If
                'Query1 = Query1 + " ) Group By ""ItemCode"" , ""ItemName"" , ""BaseUom"", ""Status"",""OnHand"",""DocNum"",""U_POST"" "
                'rs1.DoQuery(Query1)
                Dim rs1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                Dim Query1 As String =
" Select '" & Value & "' as ""Check"", " &
" STRING_AGG(TO_NVARCHAR(""DocNum""), ', ') AS ""DocNum"", ""ItemCode"", ""ItemName"", SUM(""ActualQty"") As ""Qty"", ""BaseUom"", ""Status"", ""OnHand"" " &
" From ( " &
"   Select T1.""U_ItemCode"" As ""ItemCode"", " &
"          T1.""U_ItemName"" As ""ItemName"", " &
"          ""U_Qty"" As ""Quantity"", " &
"          T1.""U_UomName"" As ""UomCode"", " &
"          ( Select (T11.""BaseQty"" * T1.""U_Qty"") " &
"            From ""OITM"" T00 " &
"            INNER JOIN ""UGP1"" T11 ON T00.""UgpEntry"" = T11.""UgpEntry"" " &
"            INNER JOIN ""OUOM"" T22 ON T22.""UomEntry"" = T11.""UomEntry"" " &
"            WHERE T00.""ItemCode"" = T1.""U_ItemCode"" " &
"              And T22.""UomCode"" = (select ""UomCode"" from OUOM where ""UomEntry"" = T1.""U_UOMCode"") " &
"              And T00.""InvntItem"" = 'Y' ) As ""ActualQty"", " &
"          T5.""InvntryUom"" As ""BaseUom"", " &
"          'PDT' As ""Status"", " &
"          ( SELECT W.""OnHand"" FROM OITW W WHERE W.""ItemCode"" = T5.""ItemCode"" " &
"            And W.""WhsCode"" = '" & objForm.Items.Item("WHSE").Specific.Value & "' ) As ""OnHand"", " &
"          T0.""DocNum"" As ""DocNum"" , T1.""U_POST""  As ""U_POST"" " &
"   From ""@ME_PSC"" T0 " &
"   Inner Join ""@ME_PSC_C0"" T1 On T1.""DocEntry"" = T0.""DocEntry"" " &
"   Inner Join OITM T5 On T5.""ItemCode"" = T1.""U_ItemCode"" " &
"   Where T1.""U_ItemCode"" <> '' " &
"     And T5.""InvntItem"" = 'Y' " &
  "     And IFNULL(T1.""U_POST"",'') = '' " &
"     And T0.""U_CountDate"" Between '" & objForm.Items.Item("FROMDT").Specific.Value & "' And '" & objForm.Items.Item("TODT").Specific.Value & "' " &
"     And T0.""U_Branch"" = '" & objForm.Items.Item("BRANCH").Specific.Selected.Value & "' " &
"     And T0.""U_WhseCode"" = '" & objForm.Items.Item("WHSE").Specific.Value & "' "
                Dim oComboType As SAPbouiCOM.ComboBox = objForm.Items.Item("TYPE").Specific
                Dim sType As String = ""

                If oComboType.Selected IsNot Nothing Then
                    sType = oComboType.Selected.Value.Trim
                End If

                If String.IsNullOrEmpty(sType) Then
                    objMain.objApplication.StatusBar.SetText("Please select Type.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If

                If sType = "ItemGroup" Or sType = "ItemSubGroup" Or sType = "ItemSubSubGroup" Then
                    Dim oItemGrp As SAPbouiCOM.ComboBox = objForm.Items.Item("ITEMGRP").Specific
                    If oItemGrp.Selected Is Nothing OrElse String.IsNullOrEmpty(oItemGrp.Selected.Value) Then
                        objMain.objApplication.StatusBar.SetText("Please select Item Group.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If
                End If

                If sType = "ItemSubGroup" Or sType = "ItemSubSubGroup" Then
                    Dim oSubGrp As SAPbouiCOM.ComboBox = objForm.Items.Item("SBGRP").Specific
                    If oSubGrp.Selected Is Nothing OrElse String.IsNullOrEmpty(oSubGrp.Selected.Value) Then
                        objMain.objApplication.StatusBar.SetText("Please select Item Sub Group.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If
                End If

                If sType = "ItemSubSubGroup" Then
                    Dim oSubSubGrp As SAPbouiCOM.ComboBox = objForm.Items.Item("SBSBGRP").Specific
                    If oSubSubGrp.Selected Is Nothing OrElse String.IsNullOrEmpty(oSubSubGrp.Selected.Value) Then
                        objMain.objApplication.StatusBar.SetText("Please select Item Sub Sub Group.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If
                End If
                If objForm.Items.Item("TYPE").Specific.Selected.Value = "All" Then
                    ' Query1 = Query1 & " And T0.""U_TYPE"" = 'All' "
                ElseIf objForm.Items.Item("TYPE").Specific.Selected.Value = "ItemGroup" Then
                    Query1 = Query1 & " And T0.""U_TYPE"" = 'ItemGroup' And T0.""U_ITMGRP"" = '" & objForm.Items.Item("ITEMGRP").Specific.Selected.Value.Trim & "' "
                ElseIf objForm.Items.Item("TYPE").Specific.Selected.Value = "ItemSubGroup" Then
                    Query1 = Query1 & " And T0.""U_TYPE"" = 'ItemSubGroup' And T0.""U_ITMGRP"" = '" & objForm.Items.Item("ITEMGRP").Specific.Selected.Value.Trim & "' " &
                       " And T0.""U_SUBGRP"" = '" & objForm.Items.Item("SBGRP").Specific.Selected.Value.Trim & "' "
                ElseIf objForm.Items.Item("TYPE").Specific.Selected.Value = "ItemSubSubGroup" Then
                    Query1 = Query1 & " And T0.""U_TYPE"" = 'ItemSubSubGroup' And T0.""U_ITMGRP"" = '" & objForm.Items.Item("ITEMGRP").Specific.Selected.Value.Trim & "' " &
                       " And T0.""U_SUBGRP"" = '" & objForm.Items.Item("SBGRP").Specific.Selected.Value.Trim & "' " &
                       " And T0.""U_SBSBGRP"" = '" & objForm.Items.Item("SBSBGRP").Specific.Selected.Value.Trim & "' "
                End If

                Query1 = Query1 &
" ) Group By ""ItemCode"", ""ItemName"", ""BaseUom"", ""Status"", ""OnHand"",""U_POST"" "

                rs1.DoQuery(Query1)


                If rs1.RecordCount > 0 Then
                    objMatrix1.Clear()
                    For i As Integer = 1 To rs1.RecordCount
                        objMatrix1.AddRow()
                        oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
                        oDBs_Details1.SetValue("U_Check", oDBs_Details1.Offset, "Y")
                        'oDBs_Details1.SetValue("U_LINE", oDBs_Details1.Offset, i)
                        oDBs_Details1.SetValue("U_LINE", oDBs_Details1.Offset, i)
                        oDBs_Details1.SetValue("U_DOCE", oDBs_Details1.Offset, rs1.Fields.Item("DocNum").Value)

                        oDBs_Details1.SetValue("U_ICode", oDBs_Details1.Offset, rs1.Fields.Item("ItemCode").Value)
                        oDBs_Details1.SetValue("U_IName", oDBs_Details1.Offset, rs1.Fields.Item("ItemName").Value)
                        oDBs_Details1.SetValue("U_Qty", oDBs_Details1.Offset, rs1.Fields.Item("Qty").Value)
                        oDBs_Details1.SetValue("U_BaseUom", oDBs_Details1.Offset, rs1.Fields.Item("BaseUom").Value)
                        oDBs_Details1.SetValue("U_Update", oDBs_Details1.Offset, "")
                        oDBs_Details1.SetValue("U_Flag", oDBs_Details1.Offset, rs1.Fields.Item("Status").Value)
                        oDBs_Details1.SetValue("U_Onhand", oDBs_Details1.Offset, rs1.Fields.Item("OnHand").Value)
                        objMatrix1.SetLineData(objMatrix1.VisualRowCount)
                        rs1.MoveNext()
                    Next
                End If
                ' oGrid.DataTable.ExecuteQuery(Query)
                '  End If
            Else



                '------
                Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                Dim Query As String = "" ' "Select  '" & Value & "' as ""Check"" ""DocNum"", ""ItemCode"" , ""ItemName"" ,  SUM(""Qty"") As ""Qty"" , ""BaseUom"",""Status""  From ( "
                Query = Query + " Select '" & Value & "' as ""Check""  , ""ItemCode"" , ""ItemName"" , SUM(""ActualQty"") As ""Qty"" , ""BaseUom"", ""Status""   From ( " &
                                      " Select  '" & Value & "' as ""Check""  , T1. ""U_ItemCode"" As ""ItemCode"" ,T1. ""U_ItemName"" As  ""ItemName"" , ""U_Qty"" As ""Quantity"" , T1.""U_UomName"" As ""UomCode"" , " &
                                      " (Select (T11.""BaseQty"" * T1.""U_Qty"") From ""OITM"" T00  INNER JOIN ""UGP1"" T11 ON T00.""UgpEntry"" = T11.""UgpEntry"" INNER JOIN ""OUOM"" T22 ON T22.""UomEntry"" = T11.""UomEntry""  WHERE T00.""ItemCode"" =T1.""U_ItemCode"" And  T22.""UomCode"" =(select ""UomCode"" from OUOM where ""UomEntry""=T1.""U_UOMCode"")and T00.""InvntItem""='Y')  As ""ActualQty"" " &
                                      " , T5.""InvntryUom"" As ""BaseUom"",'PDT' as ""Status"" , STRING_AGG(TO_NVARCHAR(""DocNum""), ', ') AS ""DocNum"", T1.""U_POST"" AS ""U_POST"" From ""@ME_PSC"" T0 Inner Join ""@ME_PSC_C0"" T1 On T1.""DocEntry"" = T0.""DocEntry""    inner join OITM T5 on T5.""ItemCode"" = T1.""U_ItemCode""" &
                                      " Where  T1.""U_ItemCode"" <> '' AND   IFNULL(T1.""U_POST"",'') = '' " &
                                      "and  T5.""InvntItem""='Y'  And T0.""U_CountDate"" Between  '" & objForm.Items.Item("FROMDT").Specific.Value & "' And " &
                                      " '" & objForm.Items.Item("TODT").Specific.Value & "' And   ""U_Branch"" = '" & objForm.Items.Item("BRANCH").Specific.Selected.Value & "' And ""U_WhseCode"" = '" & objForm.Items.Item("WHSE").Specific.Value & "'"
                If objForm.Items.Item("TYPE").Specific.Selected.Value = "All" Then
                    ' Query = Query + "And T0.""U_TYPE"" = 'All'  "
                ElseIf objForm.Items.Item("TYPE").Specific.Selected.Value = "ItemGroup" Then
                    Query = Query + "And T0.""U_TYPE"" = 'ItemGroup' And T0.""U_ITMGRP"" = '" & objForm.Items.Item("ITEMGRP").Specific.Selected.Value.Trim & "'  "
                ElseIf objForm.Items.Item("TYPE").Specific.Selected.Value = "ItemSubGroup" Then
                    Query = Query + "And T0.""U_TYPE"" = 'ItemSubGroup' And  T0.""U_ITMGRP"" = '" & objForm.Items.Item("ITEMGRP").Specific.Selected.Value.Trim & "' And T0.""U_SUBGRP"" = '" & objForm.Items.Item("SBGRP").Specific.Selected.Value.Trim & "' "
                ElseIf objForm.Items.Item("TYPE").Specific.Selected.Value = "ItemSubSubGroup" Then
                    Query = Query + "And  T0.""U_TYPE"" = 'ItemSubSubGroup' And T0.""U_ITMGRP"" = '" & objForm.Items.Item("ITEMGRP").Specific.Selected.Value.Trim & "' And  T0.""U_SUBGRP"" = '" & objForm.Items.Item("SBGRP").Specific.Selected.Value.Trim & "' and T0.""U_SBSBGRP"" = '" & objForm.Items.Item("SBSBGRP").Specific.Selected.Value.Trim & "' "
                End If
                Query = Query + " ) Group By ""ItemCode"" , ""ItemName"" , ""BaseUom"",""Status"" ,""DocNum"",""U_POST"" "
                'Query = Query + " Union  Select  '" & Value & "' as ""Check"" , T0. ""ItemCode"" ,T0. ""ItemName"" ,T1.""OnHand"" As ""Qty"" , T0.""InvntryUom"" As ""BaseUom"" ,'System' as ""Status"" 
                '  ,T1.""OnHand"" As ""OnHand""   From OITM  T0  inner join OITW T1 on T0.""ItemCode"" = T1.""ItemCode"" where T0.""InvntItem""='Y' and T1.""WhsCode""='" & objForm.Items.Item("WHSE").Specific.Value & "' and T1.""OnHand""<>0 and T0. ""ItemCode"" 
                Query = Query + "  Union  Select '" & Value & "' as ""Check"" ,""ItemCode"",""ItemName"",""Qty"",""BaseUom"",""Status"" from (select T1. ""ItemCode"",T1.""ItemName"",
                       T1.""InvntryUom"" As ""BaseUom"" , (Select IFNULL((Sum(""InQty""-""OutQty"")),0) 	from OINM where ""DocDate"" <='" & objForm.Items.Item("TODT").Specific.Value & "' and ""ItemCode""=T0.""ItemCode"" 
                       and ""Warehouse"" = T2.""WhsCode"" ) as ""Qty"",'System' ""Status"" ,T1.""ItmsGrpCod"" FROM OINM T0  INNER JOIN OITM T1 ON T0.""ItemCode"" = T1.""ItemCode"" 
                       LEFT OUTER JOIN OWHS T2 ON T0.""Warehouse"" = T2.""WhsCode"" LEFT OUTER JOIN OBPL T3 ON T2.""BPLid"" = T3.""BPLId"" 
                        where T2.""WhsCode""='" & objForm.Items.Item("WHSE").Specific.Value & "' and T1.""InvntItem""='Y')ss where ""Qty""<>'0' and ""ItemCode""    
                        Not in " & " (select T1.""U_ItemCode"" From ""@ME_PSC"" T0 Inner Join ""@ME_PSC_C0"" T1 On T1.""DocEntry"" = T0.""DocEntry"" Where T1.""U_ItemCode"" <> '' And T0.""U_CountDate"" Between '" & objForm.Items.Item("FROMDT").Specific.Value & "'
                        And '" & objForm.Items.Item("TODT").Specific.Value & "' And   ""U_Branch"" = '" & objForm.Items.Item("BRANCH").Specific.Selected.Value & "' And ""U_WhseCode"" = '" & objForm.Items.Item("WHSE").Specific.Value & "')  "
                If objForm.Items.Item("TYPE").Specific.Selected.Value = "ItemGroup" Then
                    Query = Query + " AND ss.""ItmsGrpCod""  = '" & objForm.Items.Item("ITEMGRP").Specific.Selected.Value.Trim & "'  "
                ElseIf objForm.Items.Item("TYPE").Specific.Selected.Value = "ItemSubGroup" Then
                    Query = Query + "AND  ss.""ItmsGrpCod"" = '" & objForm.Items.Item("ITEMGRP").Specific.Selected.Value.Trim & "' And T0.""U_VSubGroup"" = '" & objForm.Items.Item("SBGRP").Specific.Selected.Value.Trim & "' "
                ElseIf objForm.Items.Item("TYPE").Specific.Selected.Value = "ItemSubSubGroup" Then
                    Query = Query + "AND ss.""ItmsGrpCod"" = '" & objForm.Items.Item("ITEMGRP").Specific.Selected.Value.Trim & "' And  T0.""U_VSubGroup"" = '" & objForm.Items.Item("SBGRP").Specific.Selected.Value.Trim & "' And T0.""U_VSubSubGroup"" = '" & objForm.Items.Item("SBSBGRP").Specific.Selected.Value.Trim & "' "
                End If
                Query = Query + "  Group By ""ItemCode"" , ""ItemName"" , ""BaseUom"",""Status"" ,""Qty""order by ""Status"",""ItemCode"""
                rs.DoQuery(Query)
                If rs.RecordCount > 0 Then
                    objMatrix1.Clear()
                    For i As Integer = 1 To rs.RecordCount
                        objMatrix1.AddRow()
                        oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
                        oDBs_Details1.SetValue("U_Check", oDBs_Details1.Offset, "Y")
                        'oDBs_Details1.SetValue("U_LINE", oDBs_Details1.Offset, i)
                        oDBs_Details1.SetValue("U_DOCE", oDBs_Details1.Offset, rs.Fields.Item("DocNum").Value)

                        oDBs_Details1.SetValue("U_ICode", oDBs_Details1.Offset, rs.Fields.Item("ItemCode").Value)
                        oDBs_Details1.SetValue("U_IName", oDBs_Details1.Offset, rs.Fields.Item("ItemName").Value)
                        oDBs_Details1.SetValue("U_Qty", oDBs_Details1.Offset, rs.Fields.Item("Qty").Value)
                        oDBs_Details1.SetValue("U_BaseUom", oDBs_Details1.Offset, rs.Fields.Item("BaseUom").Value)
                        oDBs_Details1.SetValue("U_Update", oDBs_Details1.Offset, "")
                        oDBs_Details1.SetValue("U_Flag", oDBs_Details1.Offset, rs.Fields.Item("Status").Value)
                        '  oDBs_Details1.SetValue("U_Onhand", oDBs_Details1.Offset, rs.Fields.Item("OnHand").Value)
                        objMatrix1.SetLineData(objMatrix1.VisualRowCount)

                        rs.MoveNext()
                    Next
                End If
            End If
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        Finally
            objForm.Freeze(False)
            '  objForm.Items.Item("24").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
        End Try

    End Sub
End Class


