'from golden'
'madhu
Public Class ClsSupplierRebateWizard

#Region "       Declaration             "
    Public objForm As SAPbouiCOM.Form
    Dim objMatrix, objMatrix1 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim oRs As SAPbobsCOM.Recordset
    Dim oDt, oDt1, oDt2 As SAPbouiCOM.DataTable
    Public oDS, oDS1 As SAPbouiCOM.DBDataSource
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Private objCompany As Object
    Dim str, str1 As String
    Dim oGrid, oGrid1, oGrid2 As SAPbouiCOM.Grid
    Public rs, RsNum As SAPbobsCOM.Recordset
    Dim LostFocusFlag As Boolean = False
    Dim objutilities As Utilities

#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("SupplierRebateWizard.xml", "B_SUPP", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("B_SUPP", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            objutilities = New Utilities
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_SUPP")
            objComboBox = objForm.Items.Item("BRNH").Specific
            Dim oRsUpdate2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim BPLId As String = " Select T0.""BPLId"",T1.""BPLName"",T1.""DflWhs"" From USR6 T0 Inner Join " &
                                                             " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "'Order By T0.""BPLId""  "
            oRsUpdate2.DoQuery(BPLId)
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, BPLId)
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("BRNH").DisplayDesc = True
            If objComboBox.ValidValues.Count <> 0 Then
                'objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
            End If
            objMain.objApplication.StatusBar.SetText("Please wait details are loading.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


            Me.LoadDepartmentGrid(objForm.UniqueID, "")
            ' Me.LoadSectionGrid(objForm.UniqueID, "")
            oDt = objForm.DataSources.DataTables.Add("dt1")
            oDt = objForm.DataSources.DataTables.Item("dt1")
            oGrid1 = objForm.Items.Item("Item_3").Specific
            ' oGrid2 = objForm.Items.Item("Item_4").Specific
            oDt1 = objForm.DataSources.DataTables.Item("DT_0")
            oDt2 = objForm.DataSources.DataTables.Item("DT_1")
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ME_SUPP" And pVal.BeforeAction = False Then
                Me.CreateForm()
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then

            End If

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "3" And pVal.BeforeAction = False Then
                        Dim From As String = objForm.Items.Item("FDATE").Specific.Value
                        Dim Todate As String = objForm.Items.Item("TDATE").Specific.Value
                        If Todate < From Then
                            objMain.objApplication.StatusBar.SetText("To Date must be greater than From date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Sub
                        End If
                        Me.LoadGrid(objForm.UniqueID, "Y")
                    End If
                    If pVal.ItemUID = "6" And pVal.BeforeAction = False Then
                        Me.LoadGrid(objForm.UniqueID, "Y")
                    End If

                    If pVal.ItemUID = "5" And pVal.BeforeAction = False Then
                        Me.LoadGrid(objForm.UniqueID, "N")
                    End If
                    If pVal.ItemUID = "4" And pVal.BeforeAction = False Then
                        Me.PostAPCreditMemo(objForm.UniqueID)
                    End If
                    'If pVal.ItemUID = "Item_9" And objForm.Items.Item("Crtcode").Specific.Value <> "" And pVal.BeforeAction = False Then
                    '    Me.LoadDepartmentGridCost(objForm.UniqueID, objForm.Items.Item("Crtcode").Specific.Value, "Y")
                    '    Me.LoadSectionGridCost(objForm.UniqueID, objForm.Items.Item("Crtcode").Specific.Value, "Y")
                    'End If
                    If pVal.ItemUID = "Item_6" And pVal.BeforeAction = False Then
                        Dim objCheckBox As SAPbouiCOM.CheckBox = objForm.Items.Item("Item_6").Specific
                        objForm.Freeze(True)
                        If objForm.Items.Item("Item_6").Specific.checked = True Then
                            Me.LoadDepartmentGrid(objForm.UniqueID, "Y")
                        ElseIf objForm.Items.Item("Item_6").Specific.checked = False Then
                            Me.LoadDepartmentGrid(objForm.UniqueID, "N")
                        End If
                        objForm.Freeze(False)
                    End If
                    'If pVal.ItemUID = "Item_7" And pVal.BeforeAction = False Then
                    '    Dim objCheckBox As SAPbouiCOM.CheckBox = objForm.Items.Item("Item_7").Specific
                    '    objForm.Freeze(True)
                    '    If objForm.Items.Item("Item_7").Specific.checked = True Then
                    '        Me.LoadSectionGrid(objForm.UniqueID, "Y")
                    '    ElseIf objForm.Items.Item("Item_7").Specific.checked = False Then
                    '        Me.LoadSectionGrid(objForm.UniqueID, "N")
                    '    End If
                    '    objForm.Freeze(False)
                    'End If
                'Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                '    If pVal.ItemUID = "Item_8" And pVal.ColUID = "Ap Credit memo" And pVal.Before_Action = False Then
                '        Dim code As String = oGrid.DataTable.GetValue("Ap Credit memo", pVal.Row)

                '        Dim GetAP As String = "Select ""DocNum"" From ORPC where ""DocEntry""='" & code & "'"
                '        Dim oRsGetAP As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                '        oRsGetAP.DoQuery(GetAP)
                '        Dim objNewForm As SAPbouiCOM.Form
                '        objMain.objApplication.ActivateMenuItem("2309")
                '        objNewForm = objMain.objApplication.Forms.ActiveForm
                '        objNewForm.Freeze(True)
                '        objNewForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                '        objNewForm.Items.Item("8").Specific.Value = oRsGetAP.Fields.Item(0).Value
                '        objNewForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                '        objNewForm.Freeze(False)
                '    End If

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
                    If oCFL.UniqueID = "CFL_SUPP" And pVal.BeforeAction = True Then
                    End If
                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        Me.CFLFilter(objForm.UniqueID, "CFL_SUPP")

                        If oCFL.UniqueID = "CFL_SUPP" Then
                            oDBs_Head.SetValue("U_Supcd", oDBs_Head.Offset, oDT.GetValue("CardCode", 0))


                        End If

                    End If



            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            '  objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub CFLFilter(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            oCondition = oConditions.Add()
            oCondition.Alias = "CardType"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "S"
            oChooseFromList.SetConditions(oConditions)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub LoadDepartmentGrid(ByVal FormUID As String, ByVal Value As String)
        Try
            objForm.Freeze(True)
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("Item_3").Specific
            oDt1 = objForm.DataSources.DataTables.Item("DT_0")
            oGrid.DataTable = oDt1
            Try
                oGrid.DataTable.Clear()
            Catch ex As Exception
            End Try
            ' Dim Department As String = "SELECT T0.""U_DEPTNME"" as ""Department Code"" ,T0.""U_SUBDEPT"" as ""Department Name"", T0.""U_SECNME"", T0.""U_FNAME"", T0.""U_SBFNME"",'" & Value & "' as ""Check""  FROM ""@TNX_MERCHIRCY""  T0"
            '''commt Dim Department As String = "SELECT distinct T0.""U_DEPTNME"" as ""Department Code"" ,T0.""U_DEPTNME"" as ""Department Name"", '" & Value & "' as ""Check""  FROM ""@TNX_MERCHIRCY""  T0"

            Dim Department As String = "Select
     distinct T0.""Code"" as ""Department Code"" ,
    T0.""Name"" as ""Department Name"",
     '' as ""Check"" 
                FROM ""@TNX_DEPARTMENT"" T0 "

            ' Dim Department As String="select T0.""PrcCode"" as ""Department Code"", T0.""PrcName""As ""Department Name"",'" & Value & "'  as ""Check"" from ""OPRC"" T0  WHERE T0.""Locked""='N' "
            ' Dim CreationCode As String = objForm.Items.Item("Crtcode").Specific.Value
            ' Dim Department As String = "select distinct T0.""U_DEPTNME"" as ""Department Code"",T0.""U_DEPTNME""as ""Department Name"",'" & Value & "' as ""Check"" from ""@TNX_MERCHIRCY"" T0 " where T0.""U_CRECDE""='" & CreationCode & "' "

            oGrid.DataTable.ExecuteQuery(Department)
            oGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.AutoResizeColumns()
            oGrid.Columns.Item("Department Code").Editable = False
            oGrid.Columns.Item("Department Name").Editable = False

            objForm.Freeze(False)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub LoadSectionGrid(ByVal FormUID As String, ByVal Value As String)
        Try
            objForm.Freeze(True)
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("Item_4").Specific
            oDt2 = objForm.DataSources.DataTables.Item("DT_1")
            oGrid.DataTable = oDt2
            Try
                oGrid.DataTable.Clear()
            Catch ex As Exception
            End Try
            Dim Section As String = "select distinct T0.""U_SECNME"" as ""Section Code"",T0.""U_SECNME""as ""Section Name"",'" & Value & "' as ""Check"" from ""@TNX_MERCHIRCY"" T0 "
            'Dim Section As String = "select distinct T0.""U_SECNME"" as ""Section Code"",T0.""U_SECNME""as ""Section Name"", ""U_FNAME""as ""Family"",""U_SBFNME"" as""Subfamily"",'" & Value & "' as ""Check"" from ""@TNX_MERCHIRCY"" T0 "

            ' Dim Section1 As String = "select distinct T0.""U_SECNME"" as ""Section Code"",T0.""U_SECNME""as ""Section Name"",'" & Value & "' as ""Check"" from ""@TNX_MERCHIRCY"" T0 where T0.""U_CRECDE""='" & CreationCode & "' "


            oGrid.DataTable.ExecuteQuery(Section)
            oGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.AutoResizeColumns()
            oGrid.Columns.Item("Section Code").Editable = False
            oGrid.Columns.Item("Section Name").Editable = False

            objForm.Freeze(False)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub LoadGrid(ByVal FormUID As String, ByVal Value As String)
        Try
            objForm.Freeze(True)
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("Item_8").Specific
            oGrid.DataTable = oDt
            oGrid1 = objForm.Items.Item("Item_3").Specific
            ' oGrid2 = objForm.Items.Item("Item_4").Specific
            Try
                oGrid.DataTable.Clear()
            Catch ex As Exception
            End Try
            Dim test, test2, Itmgrp, Itmgrp2 As String
            For i As Integer = 0 To oGrid1.Rows.Count - 1

                If oGrid1.DataTable.GetValue("Check", i) = "Y" Then

                    test = oGrid1.DataTable.GetValue("Department Code", i)
                    If i = 0 Then
                        test2 = test
                    ElseIf i > 0 Then
                        test2 = test & "'" & "," & "'" & test2
                    End If
                End If
            Next
            Dim objComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("BRNH").Specific
            Dim cardcode As String
            If objComboBox.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Please select branch at the  header...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objForm.Freeze(False)
                Exit Sub

            End If
            'For j As Integer = 0 To oGrid2.Rows.Count - 1
            '    If oGrid2.DataTable.GetValue("Check", j) = "Y" Then
            '        Itmgrp = oGrid2.DataTable.GetValue("Section Code", j)
            '        If j = 0 Then
            '            Itmgrp2 = Itmgrp
            '        ElseIf j > 0 Then
            '            Itmgrp2 = Itmgrp & "'" & "," & "'" & Itmgrp2
            '        End If
            '    End If
            'Next
            Dim Supp As String = objForm.Items.Item("Supcd").Specific.Value
            Dim FromDate As String = objForm.Items.Item("FDATE").Specific.Value
            Dim ToDate As String = objForm.Items.Item("TDATE").Specific.Value
            Dim Branch As String = objForm.Items.Item("BRNH").Specific.Value
            Dim Depat As String = test2
            ' Dim Crtcode As String = objForm.Items.Item("Crtcode").Specific.Value
            'Dim Brand As String = objForm.Items.Item("Brand").Specific.Value
            Dim SupRebate As String = ""

            ' Dim SupRebate As String = "select * from  ""SBO_RebateAmount""('" & objForm.Items.Item("FDATE").Specific.Value & "','" & objForm.Items.Item("TDATE").Specific.Value & "','" & (test2) & "','" & Itmgrp2 & "')"

            'oGrid.DataTable.ExecuteQuery(SupRebate)
            ' End If
            If test2 <> "" And Supp = "" Then

                SupRebate = "select 'Y' as ""Check"",""SUPPLIERCODE""as ""Supplier Code"",""SUPPLIERNAME""as ""Supplier Name"",""BRAND""AS ""Brand"",""DEPARTMENT"" as ""Department"",""SECTION1"" as""Section"", ""FAMILY"" as ""Family"", ""SUBFAMILY""as ""Sub Family"", ""REBATEPERCENTAGE"" as ""Rebate Percentage"",""REBATECATEGORY"" as ""Rebate Category"",""GLACCOUNT""as ""GL Account"",""CURRENCY"" as ""Currency"",""TOTALINVOICEAMOUNT"" as ""Total Invoice Amount"",""TOTALAPCREDITAMOUNT""as ""Total APCredit Amount"", " &
"(( ""TOTALINVOICEAMOUNT"")-(""TOTALAPCREDITAMOUNT""))as ""Net purchase Amount for rebate"",((""REBATEPERCENTAGE"")/100)*(( ""TOTALINVOICEAMOUNT"")-(""TOTALAPCREDITAMOUNT""))as ""Rebate Amount"",""APCREDITMEMO"" as ""Ap Credit memo"",""STATUS""as ""Status"" from ""SBO_RebateAmount""('" & objForm.Items.Item("FDATE").Specific.Value & "','" & objForm.Items.Item("TDATE").Specific.Value & "','" & objForm.Items.Item("BRNH").Specific.Value & "')where ""DEPARTMENT""in ('" & (test2) & "')"


                '                SupRebate = "select 'Y' as ""Check"",""SUPPLIERCODE""as ""Supplier Code"",""SUPPLIERNAME""as ""Supplier Name"",""BRAND""AS ""Brand"",""DEPARTMENT"" as ""Department"",""SECTION1"" as""Section"", ""FAMILY"" as ""Family"", ""SUBFAMILY""as ""Sub Family"", ""REBATEPERCENTAGE"" as ""Rebate Percentage"",""REBATECATEGORY"" as ""Rebate Category"",""GLACCOUNT""as ""GL Account"",""TOTALINVOICEAMOUNT"" as ""Total Invoice Amount"",IFnULL(""TOTALAPCREDITAMOUNT"",0)as ""Total APCredit Amount"", " &
                '" ((""TOTALINVOICEAMOUNT"")-(IFnULL(""TOTALAPCREDITAMOUNT"",0)))as ""Net purchase Amount for rebate"",((""REBATEPERCENTAGE"")/100)* ((""TOTALINVOICEAMOUNT"")-(IFnULL(""TOTALAPCREDITAMOUNT"",0)))as ""Rebate Amount"",""APCREDITMEMO"" as ""Ap Credit memo"",""STATUS""as ""Status"" from ""SBO_RebateAmount""('" & objForm.Items.Item("FDATE").Specific.Value & "','" & objForm.Items.Item("TDATE").Specific.Value & "')where ""DEPARTMENT""in ('" & (test2) & "')"


                '---from  ""SBO_RebateAmount""('" & objForm.Items.Item("FDATE").Specific.Value & "','" & objForm.Items.Item("TDATE").Specific.Value & "','" & (test2) & "','" & Itmgrp2 & "')"


                '            SupRebate = "SELECT 'Y' as ""Check"",T1.""U_BPCD"" as ""Supplier Code"",T1.""U_BPNM"" As ""Supplier Name"",T2.""U_Brand""as ""Brand"",T2.""U_Depart"" As ""Department"",(select distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"") As ""Section"" " &
                '                ",T2.""U_Family""as ""Family"",T2.""U_SFamily""as ""Sub Family"",T2.""U_RBTPRC""As""Rebate percentage"",T2.""U_ReCat""As""Rebate Category"",T2.""U_GLact"" as ""GL Account"" " &
                '",(case when T2.""U_Depart""<>'' and ifnull(T2.""U_Section"",'')=''and ifnull(T2.""U_Family"",'')=''and ifnull(T2.""U_SFamily"",'')=''Then " &
                '"(select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y on X.""DocEntry""=Y.""DocEntry"" " &
                '   "where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"")And X.""CardCode""=T3.""BpCode""and  X.""CANCELED"" ='N' And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N'and  ifnull(Y.""U_Apnum1"",'')='')) else " &
                '" (select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y on X.""DocEntry""=Y.""DocEntry"" " &
                '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
                '"  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")And ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")And ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily""))And X.""CardCode""=T3.""BpCode""And  X.""CANCELED"" ='N' And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N'and  ifnull(Y.""U_Apnum1"",'')='')end)as ""Total Invoice Amount"" " &
                ' ",(select IfNull(Sum(Y.""LineTotal""),0) from ORPC X inner join RPC1 Y on X.""DocEntry""=Y.""DocEntry"" " &
                '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
                '"  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")and ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")and ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily"")) " &
                '"and X.""CardCode""=T3.""BpCode"" And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N' and  ifnull(Y.""U_Apnum1"",'')='')as ""Total APCredit Amount"" " &
                '            ",(case when T2.""U_Depart""<>'' and ifnull(T2.""U_Section"",'')=''and ifnull(T2.""U_Family"",'')=''and ifnull(T2.""U_SFamily"",'')=''Then " &
                '"(select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y on X.""DocEntry""=Y.""DocEntry"" " &
                '   "where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"")And X.""CardCode""=T3.""BpCode""and  X.""CANCELED"" ='N' And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N'and  ifnull(Y.""U_Apnum1"",'')='')) else " &
                '" (select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y on X.""DocEntry""=Y.""DocEntry"" " &
                '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
                '"  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")And ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")And ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily""))And X.""CardCode""=T3.""BpCode""And  X.""CANCELED"" ='N' And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N'and  ifnull(Y.""U_Apnum1"",'')='')end)- " &
                ' "(select IfNull(Sum(Y.""LineTotal""),0) from ORPC X inner join RPC1 Y on X.""DocEntry""=Y.""DocEntry"" " &
                '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
                '"  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")) " &
                '"and X.""CardCode""=T3.""BpCode"" And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N' and  ifnull(Y.""U_Apnum1"",'')='')as ""Net purchase Amount for rebate"" " &
                ' ",(T2.""U_RBTPRC"")/100 *
                ' (case when T2.""U_Depart""<>'' and ifnull(T2.""U_Section"",'')=''and ifnull(T2.""U_Family"",'')=''and ifnull(T2.""U_SFamily"",'')=''Then " &
                '"(select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y on X.""DocEntry""=Y.""DocEntry"" " &
                '   "where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"")And X.""CardCode""=T3.""BpCode""and  X.""CANCELED"" ='N' And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N'and  ifnull(Y.""U_Apnum1"",'')='')) else " &
                '" (select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y on X.""DocEntry""=Y.""DocEntry"" " &
                '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
                '"  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")And ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")And ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily""))And X.""CardCode""=T3.""BpCode""And  X.""CANCELED"" ='N' And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N'and  ifnull(Y.""U_Apnum1"",'')='')end) as ""Rebate Amount"" " &
                '            ",'          ' As ""Ap Credit memo"",cast('' AS NVARCHAR(5000)) As ""Status"" FROM ""@SBO_SUPREBATE_C0"" T2 inner Join ""@SBO_SUPREBATE"" T1 on T2.""DocEntry"" = T1.""DocEntry"" inner Join ""OOAT"" T3 on T3.""Number""=T1. ""U_BASENO"" " &
                '"where '" & objForm.Items.Item("FDATE").Specific.Value & "'between  T3.""StartDate""and  T3.""EndDate""  And '" & objForm.Items.Item("TDATE").Specific.Value & "' between  T3.""StartDate""and  T3.""EndDate""  And T2.""U_ACTV""='Y' AND  T3.""Status""='A' And T2.""U_STUS""=''And T2.""U_Depart"" in ('" & test2 & "') group by  T1.""U_BPCD"",T1.""U_BPNM"" ,T2.""U_RBTPRC"",T2.""U_Depart"",T2.""U_Section"",T2.""U_GLact"",T3.""BpCode"",T2.""U_ReCat"",T2.""U_Brand"",T2.""U_Family"",T2.""U_SFamily"" " &
                '            "having (select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y On X.""DocEntry""= Y.""DocEntry"" " &
                '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"")) " &
                '"And X.""CardCode""=T3.""BpCode"" And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N'and  ifnull(Y.""U_Apnum1"",'')='')<>0 order by T1.""U_BPCD"""






            ElseIf test2 = "" And Supp <> "" Then
                '---department and section without supplier

                'madhu comment
                '                SupRebate = "select 'Y' as ""Check"",""SUPPLIERCODE""as ""Supplier Code"",""SUPPLIERNAME""as ""Supplier Name"",""BRAND""AS ""Brand"",""DEPARTMENT"" as ""Department"",""SECTION1"" as""Section"", ""FAMILY"" as ""Family"", ""SUBFAMILY""as ""Sub Family"", ""REBATEPERCENTAGE"" as ""Rebate Percentage"",""REBATECATEGORY"" as ""Rebate Category"",""GLACCOUNT""as ""GL Account"",""CURRENCY""  as ""Currency"",""TOTALINVOICEAMOUNT"" as ""Total Invoice Amount"",""TOTALAPCREDITAMOUNT""as ""Total APCredit Amount"", " &
                '"(( ""TOTALINVOICEAMOUNT"")-(""TOTALAPCREDITAMOUNT""))as ""Net purchase Amount for rebate"",((""REBATEPERCENTAGE"")/100)*(( ""TOTALINVOICEAMOUNT"")-(""TOTALAPCREDITAMOUNT""))as ""Rebate Amount"",""APCREDITMEMO"" as ""Ap Credit memo"",""STATUS""as ""Status"" from ""SBO_RebateAmount""('" & objForm.Items.Item("FDATE").Specific.Value & "','" & objForm.Items.Item("TDATE").Specific.Value & "','" & objForm.Items.Item("BRNH").Specific.Value & "')where ""DEPARTMENT""in ('" & (test2) & "') and ""SUPPLIERCODE""='" & Supp & "'"


                'working code
                '                SupRebate = "select 'Y' as ""Check"",""SUPPLIERCODE""as ""Supplier Code"",""SUPPLIERNAME""as ""Supplier Name"",""BRAND""AS ""Brand"",""DEPARTMENT"" as ""Department"",""SECTION1"" as""Section"", ""FAMILY"" as ""Family"", ""SUBFAMILY""as ""Sub Family"", ""REBATEPERCENTAGE"" as ""Rebate Percentage"",""REBATECATEGORY"" as ""Rebate Category"",""GLACCOUNT""as ""GL Account"",""CURRENCY""  as ""Currency"",""TOTALINVOICEAMOUNT"" as ""Total Invoice Amount"",""TOTALAPCREDITAMOUNT""as ""Total APCredit Amount"", " &
                '"(( ""TOTALINVOICEAMOUNT"")-(""TOTALAPCREDITAMOUNT""))as ""Net purchase Amount for rebate"",((""REBATEPERCENTAGE"")/100)*(( ""TOTALINVOICEAMOUNT"")-(""TOTALAPCREDITAMOUNT""))as ""Rebate Amount"",""APCREDITMEMO"" as ""Ap Credit memo"",""STATUS""as ""Status"" from ""SBO_RebateAmount""('" & objForm.Items.Item("FDATE").Specific.Value & "','" & objForm.Items.Item("TDATE").Specific.Value & "','" & objForm.Items.Item("BRNH").Specific.Value & "')where ""SUPPLIERCODE""='" & Supp & "'"

                SupRebate = "CALL ""TNX_SUPPLIER_REBATE_WIZARDLOADGRID""('" & objForm.Items.Item("FDATE").Specific.Value & "','" & objForm.Items.Item("TDATE").Specific.Value & "','" & objForm.Items.Item("BRNH").Specific.Value & "','" & Supp & "')"



                '                SupRebate = "select 'Y' as ""Check"",""SUPPLIERCODE""as ""Supplier Code"",""SUPPLIERNAME""as ""Supplier Name"",""BRAND""AS ""Brand"",""DEPARTMENT"" as ""Department"",""SECTION1"" as""Section"", ""FAMILY"" as ""Family"", ""SUBFAMILY""as ""Sub Family"", ""REBATEPERCENTAGE"" as ""Rebate Percentage"",""REBATECATEGORY"" as ""Rebate Category"",""GLACCOUNT""as ""GL Account"",""TOTALINVOICEAMOUNT"" as ""Total Invoice Amount"",IFnULL(""TOTALAPCREDITAMOUNT"",0)as ""Total APCredit Amount"", " &
                '" ((""TOTALINVOICEAMOUNT"")-(IFnULL(""TOTALAPCREDITAMOUNT"",0)))as ""Net purchase Amount for rebate"",((""REBATEPERCENTAGE"")/100)* ((""TOTALINVOICEAMOUNT"")-(IFnULL(""TOTALAPCREDITAMOUNT"",0)))as ""Rebate Amount"",""APCREDITMEMO"" as ""Ap Credit memo"",""STATUS""as ""Status"" from ""SBO_RebateAmount""('" & objForm.Items.Item("FDATE").Specific.Value & "','" & objForm.Items.Item("TDATE").Specific.Value & "')where ""DEPARTMENT""in ('" & (test2) & "') and ""SUPPLIERCODE""='" & Supp & "'"

                '                '            SupRebate = "SELECT 'Y' as ""Check"",T1.""U_BPCD"" as ""Supplier Code"",T1.""U_BPNM"" As ""Supplier Name"",T2.""U_Brand""as ""Brand"",T2.""U_Depart"" As ""Department"",(select distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"") As ""Section"" " &
                '                '                ",T2.""U_Family""as ""Family"",T2.""U_SFamily""as ""Sub Family"",T2.""U_RBTPRC""As""Rebate percentage"",T2.""U_ReCat""As""Rebate Category"",T2.""U_GLact"" as ""GL Account"" " &
                '",(case when T2.""U_Depart""<>'' and ifnull(T2.""U_Section"",'')=''and ifnull(T2.""U_Family"",'')=''and ifnull(T2.""U_SFamily"",'')=''Then " &
                '"(select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y on X.""DocEntry""=Y.""DocEntry"" " &
                '   "where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"")And X.""CardCode""=T3.""BpCode""and  X.""CANCELED"" ='N' And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N'and  ifnull(Y.""U_Apnum1"",'')='')) else " &
                '" (select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y on X.""DocEntry""=Y.""DocEntry"" " &
                '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
                '"  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")And ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")And ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily""))And X.""CardCode""=T3.""BpCode""And  X.""CANCELED"" ='N' And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N'and  ifnull(Y.""U_Apnum1"",'')='')end)as ""Total Invoice Amount"" " &
                ' ",(select IfNull(Sum(Y.""LineTotal""),0) from ORPC X inner join RPC1 Y on X.""DocEntry""=Y.""DocEntry"" " &
                '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
                '"  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")and ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")and ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily"")) " &
                '"and X.""CardCode""=T3.""BpCode"" And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N' and  ifnull(Y.""U_Apnum1"",'')='')as ""Total APCredit Amount"" " &
                '            ",(case when T2.""U_Depart""<>'' and ifnull(T2.""U_Section"",'')=''and ifnull(T2.""U_Family"",'')=''and ifnull(T2.""U_SFamily"",'')=''Then " &
                '"(select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y on X.""DocEntry""=Y.""DocEntry"" " &
                '   "where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"")And X.""CardCode""=T3.""BpCode""and  X.""CANCELED"" ='N' And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N'and  ifnull(Y.""U_Apnum1"",'')='')) else " &
                '" (select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y on X.""DocEntry""=Y.""DocEntry"" " &
                '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
                '"  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")And ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")And ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily""))And X.""CardCode""=T3.""BpCode""And  X.""CANCELED"" ='N' And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N'and  ifnull(Y.""U_Apnum1"",'')='')end)- " &
                ' "(select IfNull(Sum(Y.""LineTotal""),0) from ORPC X inner join RPC1 Y on X.""DocEntry""=Y.""DocEntry"" " &
                '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
                '"  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")) " &
                '"and X.""CardCode""=T3.""BpCode"" And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N' and  ifnull(Y.""U_Apnum1"",'')='')as ""Net purchase Amount for rebate"" " &
                ' ",(T2.""U_RBTPRC"")/100 *
                ' (case when T2.""U_Depart""<>'' and ifnull(T2.""U_Section"",'')=''and ifnull(T2.""U_Family"",'')=''and ifnull(T2.""U_SFamily"",'')=''Then " &
                '"(select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y on X.""DocEntry""=Y.""DocEntry"" " &
                '   "where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"")And X.""CardCode""=T3.""BpCode""and  X.""CANCELED"" ='N' And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N'and  ifnull(Y.""U_Apnum1"",'')='')) else " &
                '" (select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y on X.""DocEntry""=Y.""DocEntry"" " &
                '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
                '"  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")And ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")And ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily""))And X.""CardCode""=T3.""BpCode""And  X.""CANCELED"" ='N' And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N'and  ifnull(Y.""U_Apnum1"",'')='')end) as ""Rebate Amount"" " &
                '            ",'          ' As ""Ap Credit memo"",cast('' AS NVARCHAR(5000)) As ""Status"" FROM ""@SBO_SUPREBATE_C0"" T2 inner Join ""@SBO_SUPREBATE"" T1 on T2.""DocEntry"" = T1.""DocEntry"" inner Join ""OOAT"" T3 on T3.""Number""=T1. ""U_BASENO"" " &
                '"where  T1.""U_BPCD""='" & Supp & "' and  '" & objForm.Items.Item("FDATE").Specific.Value & "'between  T3.""StartDate""and  T3.""EndDate""  And '" & objForm.Items.Item("TDATE").Specific.Value & "' between  T3.""StartDate""and  T3.""EndDate""  And T2.""U_ACTV""='Y' AND  T3.""Status""='A' And T2.""U_STUS""=''And T2.""U_Depart"" in ('" & test2 & "') group by  T1.""U_BPCD"",T1.""U_BPNM"" ,T2.""U_RBTPRC"",T2.""U_Depart"",T2.""U_Section"",T2.""U_GLact"",T3.""BpCode"",T2.""U_ReCat"",T2.""U_Brand"",T2.""U_Family"",T2.""U_SFamily"" " &
                '            "having (select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y On X.""DocEntry""= Y.""DocEntry"" " &
                '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"")) " &
                '"And X.""CardCode""=T3.""BpCode"" And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N'and  ifnull(Y.""U_Apnum1"",'')='')<>0 order by T1.""U_BPCD"""

            End If


            oGrid.DataTable.ExecuteQuery(SupRebate)
            ' oGrid.Columns.Item("Total Invoice Amount").Visible = False
            ' oGrid.Columns.Item("Total APCredit Amount").Visible = False

            oGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item("Supplier Code").Editable = False
            oGrid.Columns.Item("Supplier Name").Editable = False
            oGrid.Columns.Item("AP Invoice Number").Editable = False
            oGrid.Columns.Item("Department").Editable = False
            oGrid.Columns.Item("Section").Editable = False
            oGrid.Columns.Item("Brand").Editable = False
            oGrid.Columns.Item("Family").Editable = False
            oGrid.Columns.Item("Sub Family").Editable = False
            oGrid.Columns.Item("Rebate Percentage").Editable = False
            ' oGrid.Columns.Item("Rebate Category").Editable = False
            oGrid.Columns.Item("Currency").Editable = False
            oGrid.Columns.Item("VatGroup").Editable = False
            oGrid.Columns.Item("Total Invoice Amount").Editable = False
            oGrid.Columns.Item("Rebate Amount").Editable = False
            oGrid.Columns.Item("Net purchase Amount for rebate").Editable = False
            oGrid.Columns.Item("Total APCredit Amount").Editable = False
            oGrid.Columns.Item("GL Account").Editable = False
            oGrid.AutoResizeColumns()

            ' End If
            ' End If
            ' Me.TotalAmount(objForm.UniqueID)

            'Dim oEditTextCol As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("APCreditMemo")
            '    oEditTextCol.LinkedObjectType = 19

            '    Dim oEditTextCol1 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("Supplier Code")
            '    oEditTextCol1.LinkedObjectType = 2


            Dim oEditTextCol As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("Ap Credit memo")
            oEditTextCol.LinkedObjectType = 19

            Dim oEditTextCol1 As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("Supplier Code")
            oEditTextCol1.LinkedObjectType = 2
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    'Private Sub LoadGrid(ByVal FormUID As String, ByVal Value As String)
    '    Try
    '        objForm.Freeze(True)
    '        objForm = objMain.objApplication.Forms.Item(FormUID)
    '        oGrid = objForm.Items.Item("Item_8").Specific
    '        ' oDt = objForm.DataSources.DataTables.Item("DT_SUPP")
    '        oGrid.DataTable = oDt
    '        oGrid1 = objForm.Items.Item("Item_3").Specific
    '        oGrid2 = objForm.Items.Item("Item_4").Specific
    '        Try
    '            oGrid.DataTable.Clear()
    '        Catch ex As Exception
    '        End Try
    '        '        Dim SupRebate As String = "SELECT '" & Value & "' as""Check"",T1.""U_BPCD""as ""Supplier Code"", T1.""U_BPNM"" As ""Supplier Name"", " &
    '        '         " T2.""U_POAF""As""Purchase From"", T2.""U_POAT""As""Purchase To"",T2.""U_GLact"" as ""GLAccount"",'' as ""Total Invoice Amount"",sum(T4.""DocTotal""-T4.""PaidToDate"" ) as ""Paid Invoice"", '' as ""Balance Amount"" , '          ' As ""APCreditMemo"",cast('' AS NVARCHAR(5000)) As ""Status""    " &
    '        '                "FROM ""@SBO_SUPREBATE_C0"" T2 inner Join ""@SBO_SUPREBATE"" T1 on  T2.""DocEntry"" = T1.""DocEntry""  inner Join ""OOAT"" T3 on T3.""Number""=T1. ""U_BASENO""  " &
    '        '         " inner Join ""OPCH"" T4 on T4.""CardCode""=T3.""BpCode"" where  T3.""StartDate"" between  '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "' " &
    '        '"And T2.""U_ACTV""='Y' and T2.""U_STUS""='' group by  T1.""U_BPCD"",T1.""U_BPNM"" ,T2.""U_POAF"",T2.""U_POAT"",T2.""U_GLact"""
    '        Dim test, test2 As String
    '        For i As Integer = 0 To oGrid1.Columns.Count - 1

    '            If oGrid1.DataTable.GetValue("Check", i) = "Y" Then
    '                'Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '                test = oGrid1.DataTable.GetValue("CostCenter Code", i)

    '                If i = 0 Then

    '                    test2 = test
    '                ElseIf i > 0 Then

    '                    test2 = test & "'" & "," & "'" & test2

    '                End If
    '                ' oRs.DoQuery(SupRebate)
    '            End If
    '            'oRs.MoveNext()
    '            ' Dim test As String = oGrid1.DataTable.GetValue("CostCenter Code", i)

    '        Next
    '        test2 = test2
    '        'For j As Integer = 0 To oGrid2.Columns.Count - 1
    '        '    If oGrid2.DataTable.GetValue("Check", j) = "Y" Then
    '        '        test2 = oGrid2.DataTable.GetValue("ItemGroup Code", j)
    '        '    End If         
    '        'Next
    '        Dim SupRebate As String = "Call ""SBO_SupplierWizard""('" & objForm.Items.Item("FDATE").Specific.Value & "','" & objForm.Items.Item("TDATE").Specific.Value & "','" & test2 & "')"
    '        oGrid.DataTable.ExecuteQuery(SupRebate)
    '        oGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
    '        oGrid.Columns.Item("Supplier Code").Editable = False
    '        oGrid.Columns.Item("Supplier Name").Editable = False
    '        oGrid.Columns.Item("Department").Editable = False
    '        oGrid.Columns.Item("Section").Editable = False
    '        oGrid.Columns.Item("Rebate Percentage").Editable = False
    '        oGrid.Columns.Item("Rebate Category").Editable = False
    '        oGrid.Columns.Item("Total Invoice Amount").Editable = False
    '        oGrid.Columns.Item("Rebate Amount").Editable = False
    '        oGrid.Columns.Item("GLAccount").Editable = False
    '        oGrid.AutoResizeColumns()
    '        Me.TotalAmount(objForm.UniqueID)
    '        objForm.Freeze(False)
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '    End Try
    ' -------

    '                SupRebate = "SELECT 'Y' as ""Check"",T1.""U_BPCD"" as ""Supplier Code"",T1.""U_BPNM"" As ""Supplier Name"",T2.""U_Brand""as ""Brand"",T2.""U_Depart"" As ""Department"",(select distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"") As ""Section"" " &
    '                    ",T2.""U_Family""as ""Family"",T2.""U_SFamily""as ""Sub Family"",T2.""U_RBTPRC""As""Rebate percentage"",T2.""U_ReCat""As""Rebate Category"",T2.""U_GLact"" as ""GL Account"" " &
    '    ",(select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y on X.""DocEntry""=Y.""DocEntry"" " &
    '    "where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
    '    "  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")and ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")and ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily""))And X.""CardCode""=T3.""BpCode""and  X.""CANCELED"" ='N' And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N'and  ifnull(Y.""U_Apnum1"",'')='')as ""Total Invoice Amount"" " &
    '     ",(select IfNull(Sum(Y.""LineTotal""),0) from ORPC X inner join RPC1 Y on X.""DocEntry""=Y.""DocEntry"" " &
    '    "where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
    '    "  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")and ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")and ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily"")) " &
    '    "and X.""CardCode""=T3.""BpCode"" And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N' and  ifnull(Y.""U_Apnum1"",'')='')as ""Total APCredit Amount"" " &
    '     ",(select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y on X.""DocEntry""=Y.""DocEntry"" " &
    '    "where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
    '    "  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")and ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")and ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily"")) " &
    '    "and X.""CardCode""=T3.""BpCode""and  X.""CANCELED"" ='N' And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N' and  ifnull(Y.""U_Apnum1"",'')='') - " &
    '     "(select IfNull(Sum(Y.""LineTotal""),0) from ORPC X inner join RPC1 Y on X.""DocEntry""=Y.""DocEntry"" " &
    '    "where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
    '    "  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")and ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")and ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily"")) " &
    '    "and X.""CardCode""=T3.""BpCode"" And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N' and  ifnull(Y.""U_Apnum1"",'')='')as ""Net purchase Amount for rebate"" " &
    '     ",(T2.""U_RBTPRC"")/100 * (select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y on X.""DocEntry""=Y.""DocEntry"" " &
    '    "where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
    '    "  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")and ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")and ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily""))And X.""CardCode""=T3.""BpCode"" and  X.""CANCELED"" ='N' And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N' and  ifnull(Y.""U_Apnum1"",'')='') as ""Rebate Amount"" " &
    '                ",'          ' As ""Ap Credit memo"",cast('' AS NVARCHAR(5000)) As ""Status"" FROM ""@SBO_SUPREBATE_C0"" T2 inner Join ""@SBO_SUPREBATE"" T1 on T2.""DocEntry"" = T1.""DocEntry"" inner Join ""OOAT"" T3 on T3.""Number""=T1. ""U_BASENO"" " &
    '    "where '" & objForm.Items.Item("FDATE").Specific.Value & "'between  T3.""StartDate""and  T3.""EndDate""  And '" & objForm.Items.Item("TDATE").Specific.Value & "' between  T3.""StartDate""and  T3.""EndDate""  And T2.""U_ACTV""='Y' AND  T3.""Status""='A' And T2.""U_STUS""=''And T2.""U_Depart"" in ('" & test2 & "') group by  T1.""U_BPCD"",T1.""U_BPNM"" ,T2.""U_RBTPRC"",T2.""U_Depart"",T2.""U_Section"",T2.""U_GLact"",T3.""BpCode"",T2.""U_ReCat"",T2.""U_Brand"",T2.""U_Family"",T2.""U_SFamily"" " &
    '                "having (select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y On X.""DocEntry""= Y.""DocEntry"" " &
    '    "where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
    '    "  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")and ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")and ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily"")) " &
    '    "And X.""CardCode""=T3.""BpCode"" And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N'and  ifnull(Y.""U_Apnum1"",'')='')<>0 order by T1.""U_BPCD"""
    '
    '  SupRebate = "SELECT 'Y' as ""Check"",T1.""U_BPCD"" as ""Supplier Code"",T1.""U_BPNM"" As ""Supplier Name"",T2.""U_Brand""as ""Brand"",T2.""U_Depart"" As ""Department"",(select distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"") As ""Section"" " &
    '                ",T2.""U_Family""as ""Family"",T2.""U_SFamily""as ""Sub Family"",T2.""U_RBTPRC""As""Rebate percentage"",T2.""U_ReCat""As""Rebate Category"",T2.""U_GLact"" as ""GL Account"" " &
    '",(select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y on X.""DocEntry""=Y.""DocEntry"" " &
    '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
    '"  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")and ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")and ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily""))And X.""CardCode""=T3.""BpCode""and  X.""CANCELED"" ='N' And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N'and  ifnull(Y.""U_Apnum1"",'')='')as ""Total Invoice Amount"" " &
    ' ",(select IfNull(Sum(Y.""LineTotal""),0) from ORPC X inner join RPC1 Y on X.""DocEntry""=Y.""DocEntry"" " &
    '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
    '"  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")and ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")and ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily"")) " &
    '"and X.""CardCode""=T3.""BpCode"" And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N' and  ifnull(Y.""U_Apnum1"",'')='')as ""Total APCredit Amount"" " &
    ' ",(select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y on X.""DocEntry""=Y.""DocEntry"" " &
    '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
    '"  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")and ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")and ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily"")) " &
    '"and X.""CardCode""=T3.""BpCode""and  X.""CANCELED"" ='N' And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N' and  ifnull(Y.""U_Apnum1"",'')='') - " &
    ' "(select IfNull(Sum(Y.""LineTotal""),0) from ORPC X inner join RPC1 Y on X.""DocEntry""=Y.""DocEntry"" " &
    '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
    '"  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")and ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")and ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily"")) " &
    '"and X.""CardCode""=T3.""BpCode"" And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N' and  ifnull(Y.""U_Apnum1"",'')='')as ""Net purchase Amount for rebate "" " &
    ' ",(T2.""U_RBTPRC"")/100 * (select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y on X.""DocEntry""=Y.""DocEntry"" " &
    '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
    '"  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")and ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")and ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily""))And X.""CardCode""=T3.""BpCode"" and  X.""CANCELED"" ='N' And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N' and  ifnull(Y.""U_Apnum1"",'')='') as ""Rebate Amount"" " &
    '            ",'          ' As ""Ap Credit memo"",cast('' AS NVARCHAR(5000)) As ""Status"" FROM ""@SBO_SUPREBATE_C0"" T2 inner Join ""@SBO_SUPREBATE"" T1 on T2.""DocEntry"" = T1.""DocEntry"" inner Join ""OOAT"" T3 on T3.""Number""=T1. ""U_BASENO"" " &
    '"where  T1.""U_BPCD""='" & Supp & "' AND '" & objForm.Items.Item("FDATE").Specific.Value & "'between  T3.""StartDate""and  T3.""EndDate""  And '" & objForm.Items.Item("TDATE").Specific.Value & "' between  T3.""StartDate""and  T3.""EndDate""  And T2.""U_ACTV""='Y' AND  T3.""Status""='A' And T2.""U_STUS""=''And T2.""U_Depart"" in ('" & test2 & "') group by  T1.""U_BPCD"",T1.""U_BPNM"" ,T2.""U_RBTPRC"",T2.""U_Depart"",T2.""U_Section"",T2.""U_GLact"",T3.""BpCode"",T2.""U_ReCat"",T2.""U_Brand"",T2.""U_Family"",T2.""U_SFamily"" " &
    '            "having (select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y On X.""DocEntry""= Y.""DocEntry"" " &
    '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
    '"  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")and ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")and ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily"")) " &
    '"And X.""CardCode""=T3.""BpCode"" And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N'and  ifnull(Y.""U_Apnum1"",'')='')<>0 order by T1.""U_BPCD"""

    ''SupRebate = "SELECT 'Y' as ""Check"",T1.""U_BPCD"" as ""Supplier Code"",T1.""U_BPNM"" As ""Supplier Name"",T2.""U_Brand""as ""Brand"",T2.""U_Depart"" As ""Department"",(select distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"") As ""Section"" " &
    '                ",T2.""U_Family""as ""Family"",T2.""U_SFamily""as ""Sub Family"",T2.""U_RBTPRC""As""Rebate percentage"",T2.""U_ReCat""As""Rebate Category"",T2.""U_GLact"" as ""GL Account"" " &
    '",(select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y on X.""DocEntry""=Y.""DocEntry"" " &
    '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
    '"  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")and ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")and ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily""))And X.""CardCode""=T3.""BpCode""and  X.""CANCELED"" ='N' And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N'and  ifnull(Y.""U_Apnum1"",'')='')as ""Total Invoice Amount"" " &
    ' ",(select IfNull(Sum(Y.""LineTotal""),0) from ORPC X inner join RPC1 Y on X.""DocEntry""=Y.""DocEntry"" " &
    '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
    '"  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")and ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")and ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily"")) " &
    '"and X.""CardCode""=T3.""BpCode"" And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N' and  ifnull(Y.""U_Apnum1"",'')='')as ""Total APCredit Amount"" " &
    ' ",(select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y on X.""DocEntry""=Y.""DocEntry"" " &
    '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
    '"  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")and ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")and ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily"")) " &
    '"and X.""CardCode""=T3.""BpCode""and  X.""CANCELED"" ='N' And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N' and  ifnull(Y.""U_Apnum1"",'')='') - " &
    ' "(select IfNull(Sum(Y.""LineTotal""),0) from ORPC X inner join RPC1 Y on X.""DocEntry""=Y.""DocEntry"" " &
    '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
    '"  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")) " &
    '"and X.""CardCode""=T3.""BpCode"" And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N' and  ifnull(Y.""U_Apnum1"",'')='')as ""Net purchase Amount for rebate"" " &
    ' ",(T2.""U_RBTPRC"")/100 * (select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y on X.""DocEntry""=Y.""DocEntry"" " &
    '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And ""U_Section""=(Select " &
    '"  distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY"" where ""U_SECNME""=T2.""U_Section"")and ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")and ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily""))And X.""CardCode""=T3.""BpCode"" and  X.""CANCELED"" ='N' And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N' and  ifnull(Y.""U_Apnum1"",'')='') as ""Rebate Amount"" " &
    '            ",'          ' As ""Ap Credit memo"",cast('' AS NVARCHAR(5000)) As ""Status"" FROM ""@SBO_SUPREBATE_C0"" T2 inner Join ""@SBO_SUPREBATE"" T1 on T2.""DocEntry"" = T1.""DocEntry"" inner Join ""OOAT"" T3 on T3.""Number""=T1. ""U_BASENO"" " &
    '"where '" & objForm.Items.Item("FDATE").Specific.Value & "'between  T3.""StartDate""and  T3.""EndDate""  And '" & objForm.Items.Item("TDATE").Specific.Value & "' between  T3.""StartDate""and  T3.""EndDate""  And T2.""U_ACTV""='Y' AND  T3.""Status""='A' And T2.""U_STUS""=''And T2.""U_Depart"" in ('" & test2 & "') group by  T1.""U_BPCD"",T1.""U_BPNM"" ,T2.""U_RBTPRC"",T2.""U_Depart"",T2.""U_Section"",T2.""U_GLact"",T3.""BpCode"",T2.""U_ReCat"",T2.""U_Brand"",T2.""U_Family"",T2.""U_SFamily"" " &
    '            "having (select IfNull(Sum(Y.""LineTotal""),0) from OpcH X inner join PCH1 Y On X.""DocEntry""= Y.""DocEntry"" " &
    '"where ""ItemCode""In(Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB""   WHERE ""ItmsGrpNam"" = T2.""U_Depart"")) " &
    '"And X.""CardCode""=T3.""BpCode"" And X.""DocDate"" between '" & objForm.Items.Item("FDATE").Specific.Value & "' And '" & objForm.Items.Item("TDATE").Specific.Value & "'and Y.""U_FOC""='N'and  ifnull(Y.""U_Apnum1"",'')='')<>0 order by T1.""U_BPCD"""

    ''End Sub
    Private Sub PostAPCreditMemobackup(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim FromDate As String = objForm.Items.Item("FDATE").Specific.Value
            Dim ToDate As String = objForm.Items.Item("TDATE").Specific.Value
            Dim objComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("BRNH").Specific
            Dim cardcode As String
            If objComboBox.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Please select branch at the  header...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            Dim PostDate As String = objForm.Items.Item("PDate").Specific.Value.ToString().Trim()
            If PostDate = "" Then
                objMain.objApplication.SetStatusBarMessage("Please enter the Posting Date", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Exit Sub
            End If
            oGrid = objForm.Items.Item("Item_8").Specific
            Dim Supplier, Supplier1 As String
            Supplier1 = ""
            Dim oPurchaseCreditNotes As SAPbobsCOM.Documents

            For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                Supplier = oGrid.DataTable.GetValue("Supplier Code", i)
                If Supplier1.Contains(Supplier) = False Then
                    Supplier1 = Supplier1 & "," & Supplier
                    oPurchaseCreditNotes = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                    oPurchaseCreditNotes.CardCode = oGrid.DataTable.GetValue("Supplier Code", i)
                    oPurchaseCreditNotes.BPL_IDAssignedToInvoice = objForm.Items.Item("BRNH").Specific.Value
                    Dim currency As String = oGrid.DataTable.GetValue("Currency", i)
                    If currency = "AED" Then
                        oPurchaseCreditNotes.DocCurrency = "AED"
                    ElseIf currency = "USD" Then
                        oPurchaseCreditNotes.DocCurrency = "USD"

                    ElseIf currency = "OMR" Then
                        oPurchaseCreditNotes.DocCurrency = "OMR"
                    Else
                        'oPurchaseCreditNotes.DocCurrency = "YER"
                        oPurchaseCreditNotes.DocCurrency = "QR"

                    End If
                    Dim sDate As String = objForm.Items.Item("PDate").Specific.Value.ToString().Trim()
                    ' Dim sDate As String = DateTime.Now.ToString("yyyyMMdd")
                    sDate = sDate.Insert("4", "-")
                    sDate = sDate.Insert("7", "-")
                    oPurchaseCreditNotes.DocDate = sDate
                    oPurchaseCreditNotes.DocDueDate = sDate
                    oPurchaseCreditNotes.TaxDate = sDate
                    oPurchaseCreditNotes.NumAtCard = "Rebate Generation"
                    oPurchaseCreditNotes.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service

                    For j As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                        If oGrid.DataTable.GetValue("Check", j) = "Y" And oGrid.DataTable.GetValue("Ap Credit memo", j) = "" Then
                            If Supplier = oGrid.DataTable.GetValue("Supplier Code", j) Then
                                Dim Depart As String = "select ""OcrCode"",""OcrName"" From OOCR where ""OcrName""='" & oGrid.DataTable.GetValue("Department", j) & "'and ""DimCode""='1' "
                                Dim Departrs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Departrs.DoQuery(Depart)
                                Dim sect As String = "select ""OcrCode"",""OcrName"" From OOCR where ""OcrName""='" & oGrid.DataTable.GetValue("Section", j) & "'and ""DimCode""='2' "
                                Dim sectrs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                sectrs.DoQuery(sect)
                                oPurchaseCreditNotes.Lines.ItemDescription = "Rebate Generation"
                                oPurchaseCreditNotes.Lines.AccountCode = oGrid.DataTable.GetValue("GL Account", j)
                                oPurchaseCreditNotes.Lines.LineTotal = oGrid.DataTable.GetValue("Rebate Amount", j)
                                oPurchaseCreditNotes.Lines.CostingCode = Departrs.Fields.Item(0).Value
                                oPurchaseCreditNotes.Lines.CostingCode2 = sectrs.Fields.Item(0).Value
                                oPurchaseCreditNotes.Lines.CostingCode2 = sectrs.Fields.Item(0).Value
                                'oPurchaseCreditNotes.Lines.UserFields.Fields.Item("U_Family").Value = oGrid.DataTable.GetValue("Family", j)
                                'oPurchaseCreditNotes.Lines.UserFields.Fields.Item("U_Subfamily").Value = oGrid.DataTable.GetValue("Sub Family", j)
                                oPurchaseCreditNotes.Lines.Add()
                            End If
                        End If
                    Next

                    If oPurchaseCreditNotes.Add = 0 Then
                        Dim GetAPInvoiceEntry1 As String = "select ""CardCode"",""DocEntry"",""DocNum"" From ORPC where ""DocEntry""=(Select Max(""DocEntry"") From ORPC ) "
                        Dim oRsGetAPInvoiceEntry1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetAPInvoiceEntry1.DoQuery(GetAPInvoiceEntry1)
                        For K As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                            If oGrid.DataTable.GetValue("Check", K) = "Y" Then
                                If oGrid.DataTable.GetValue("Supplier Code", K) = oRsGetAPInvoiceEntry1.Fields.Item(0).Value.ToString Then
                                    oGrid.DataTable.SetValue("Ap Credit memo", K, oRsGetAPInvoiceEntry1.Fields.Item(1).Value.ToString)
                                    oGrid.DataTable.SetValue("Status", K, "Posted Successfully")
                                    '--Ap invoice updation
                                    Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim oRs1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    'Dim ItmGroup1 As String = "select  distinct Y.""LineTotal"", Y.""LineNum"",Y.""ItemCode"" ,X.""DocNum""from OpcH X inner join PCH1 Y On X.""DocEntry""=Y.""DocEntry"" inner Join  ""OOAT"" T3 on T3.""BpCode""=X.""CardCode"" " &
                                    '"inner Join  ""@SBO_SUPREBATE"" T1 on T3.""Number""=T1. ""U_BASENO"" inner Join ""@SBO_SUPREBATE_C0"" T2 on T2.""DocEntry"" = T1.""DocEntry""where X.""CardCode""='" & oGrid.DataTable.GetValue("Supplier Code", K) & "' " &
                                    '"And Y.""ItemCode""In (Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB"" WHERE ""ItmsGrpNam"" = T2.""U_Depart"")  And ""U_Section""=(Select distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY""  " &
                                    '"where ""U_SECNME""=T2.""U_Section"")and ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")and ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily""))And Y.""U_FOC""='N' and X.""CANCELED"" ='N' " &
                                    ''"and  X.""DocDate"" between '" & FromDate & "'and '" & ToDate & "'and T2.""U_Depart""='" & oGrid.DataTable.GetValue("Department", K) & "'and  T2.""U_Section""='" & oGrid.DataTable.GetValue("Section", K) & "'and  T2.""U_Family""='" & oGrid.DataTable.GetValue("Family", K) & "'and  T2.""U_SFamily""='" & oGrid.DataTable.GetValue("Sub Family", K) & "' " &
                                    '"and ifnull(Y.""U_Apnum1"",'')='' "

                                    'madhu  comment

                                    '    Dim ItmGroup1 As String = "select  distinct Y.""LineTotal"", Y.""LineNum"",Y.""ItemCode"" ,X.""DocNum"",T3.""StartDate"",T3.""EndDate"",T3.""Number"" from OPCH X inner join PCH1 Y On X.""DocEntry""=Y.""DocEntry"" inner Join  ""OOAT"" T3 on T3.""BpCode""=X.""CardCode""and  Y.""AgrNo"" = T3.""AbsID"" " &
                                    '"inner Join  ""@SBO_SUPREBATE"" T1 on T3.""Number""=T1. ""U_BASENO"" inner Join ""@SBO_SUPREBATE_C0"" T2 on T2.""DocEntry"" = T1.""DocEntry""where X.""CardCode""='" & oGrid.DataTable.GetValue("Supplier Code", K) & "' " &
                                    '"And Y.""ItemCode""In (Select ""ItemCode"" from OITM T where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB"" WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And T.""U_Section""=(Case When IFNULL(T2.""U_Section"",'')<>'' Then T2.""U_Section"" Else T.""U_Section"" End)And T.""U_VSubGroup""= (Case when IFNULL(T2.""U_Family"",'')<>'' Then T2.""U_Family"" Else T.""U_VSubGroup"" End)And T.""U_VSubSubGroup""=(Case When IFNULL(T2.""U_SFamily"",'')<>'' Then T2.""U_SFamily"" Else T.""U_VSubSubGroup"" End) )And Y.""U_FOC""='N'And T2.""U_ACTV""='Y' And  T3.""Status""='A' And ifnull(T2.""U_STUS"",'')='' and X.""CANCELED"" ='N' " &
                                    '"and  X.""DocDate"" between '" & FromDate & "'and '" & ToDate & "' and T2.""U_Depart""='" & oGrid.DataTable.GetValue("Department", K) & "'" &
                                    '"and ifnull(Y.""U_Apnum1"",'')='' "

                                    Dim ItmGroup1 As String = "select distinct Y.""LineTotal"", Y.""LineNum"", Y.""ItemCode"", X.""DocNum"", T3.""StartDate"", T3.""EndDate"", T3.""Number"" from OPCH X inner join PCH1 Y On X.""DocEntry""=Y.""DocEntry"" inner join ""OOAT"" T3 on T3.""BpCode""=X.""CardCode"" and Y.""AgrNo""=T3.""AbsID"" " &
                                     "inner join ""@SBO_SUPREBATE"" T1 on T3.""Number""=T1.""U_BASENO"" inner join ""@SBO_SUPREBATE_C0"" T2 on T2.""DocEntry""=T1.""DocEntry"" " &
                                     "where X.""CardCode""='" & oGrid.DataTable.GetValue("Supplier Code", K) & "' " &
                                     "and Y.""ItemCode"" in (select T.""ItemCode"" from OITM T) and T3.""Status""='A' and ifnull(T2.""U_STUS"",'')='' and X.""CANCELED""='N' " &
                                      "and X.""DocDate"" between '" & FromDate & "' and '" & ToDate & "' and ifnull(Y.""U_Apnum1"",'')='' "



                                    oRs1.DoQuery(ItmGroup1)

                                    'Dim startdate1 As Date = Trim(oRs1.Fields.Item("StartDate").Value.ToString())
                                    'Dim Enddate1 As Date = Trim(oRs1.Fields.Item("EndDate").Value.ToString())
                                    If oRs1.RecordCount > 0 Then
                                        For p As Integer = 1 To oRs1.RecordCount
                                            Dim Update As String = " UPDATE T1 SET  ""U_Apnum1""='" & oRsGetAPInvoiceEntry1.Fields.Item(2).Value.ToString & "' FROM ""PCH1""T1 INNER JOIN OPCH T2 ON T2.""DocEntry""=T1.""DocEntry""   WHERE T2.""CardCode""='" & oGrid.DataTable.GetValue("Supplier Code", K) & "' and T2.""DocNum""='" & oRs1.Fields.Item("DocNum").Value & "'and  T2.""DocDate"" between '" & FromDate & "'and '" & ToDate & "' and T1.""U_FOC""='N' and ""LineNum""='" & oRs1.Fields.Item("LineNum").Value & "' and ifnull(T1.""U_Apnum1"",'')='' "

                                            oRsUpdate.DoQuery(Update)
                                            oRs1.MoveNext()
                                        Next
                                    End If
                                    '--Ap Credit memo updation
                                    Dim oRsUpdate1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim oRs9 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    'madhu quert comment

                                    '            Dim ItmGroup2 As String = "select  distinct Y.""LineTotal"", Y.""LineNum"",Y.""ItemCode"" ,X.""DocNum"",T3.""StartDate"",T3.""EndDate"",T3.""Number"" from ORPC X inner join RPC1 Y On X.""DocEntry""=Y.""DocEntry"" inner Join  ""OOAT"" T3 on T3.""BpCode""=X.""CardCode""and  Y.""AgrNo"" = T3.""AbsID""  " &
                                    '"inner Join  ""@SBO_SUPREBATE"" T1 on T3.""Number""=T1. ""U_BASENO"" inner Join ""@SBO_SUPREBATE_C0"" T2 on T2.""DocEntry"" = T1.""DocEntry""where X.""CardCode""='" & oGrid.DataTable.GetValue("Supplier Code", K) & "' " &
                                    '"And Y.""ItemCode""In (Select ""ItemCode"" from OITM T where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB"" WHERE ""ItmsGrpNam"" = T2.""U_Depart"")And T.""U_Section""=(Case When IFNULL(T2.""U_Section"",'')<>'' Then T2.""U_Section"" Else T.""U_Section"" End)And T.""U_VSubGroup""= (Case when IFNULL(T2.""U_Family"",'')<>'' Then T2.""U_Family"" Else T.""U_VSubGroup"" End)And T.""U_VSubSubGroup""=(Case When IFNULL(T2.""U_SFamily"",'')<>'' Then T2.""U_SFamily"" Else T.""U_VSubSubGroup"" End) )And Y.""U_FOC""='N'And T2.""U_ACTV""='Y' And  T3.""Status""='A' And ifnull(T2.""U_STUS"",'')='' and X.""CANCELED"" ='N' " &
                                    '"And  X.""DocDate"" between   '" & FromDate & "'and '" & ToDate & "' and T2.""U_Depart""='" & oGrid.DataTable.GetValue("Department", K) & "'" &
                                    '"And ifnull(Y.""U_Apnum1"",'')='' "

                                    Dim ItmGroup2 As String = "select distinct Y.""LineTotal"", Y.""LineNum"", Y.""ItemCode"", X.""DocNum"", T3.""StartDate"", T3.""EndDate"", T3.""Number"" from ORPC X inner join RPC1 Y On X.""DocEntry""=Y.""DocEntry"" inner join ""OOAT"" T3 on T3.""BpCode""=X.""CardCode"" and Y.""AgrNo""=T3.""AbsID"" " &
                                    "inner join ""@SBO_SUPREBATE"" T1 on T3.""Number""=T1.""U_BASENO"" inner join ""@SBO_SUPREBATE_C0"" T2 on T2.""DocEntry""=T1.""DocEntry"" " &
                                    "where X.""CardCode""='" & oGrid.DataTable.GetValue("Supplier Code", K) & "' " &
                                    "and Y.""ItemCode"" in (select T.""ItemCode"" from OITM T) and T3.""Status""='A' and ifnull(T2.""U_STUS"",'')='' and X.""CANCELED""='N' " &
                                    "and X.""DocDate"" between '" & FromDate & "' and '" & ToDate & "' and ifnull(Y.""U_Apnum1"",'')='' "



                                    oRs9.DoQuery(ItmGroup2)
                                    'Dim startdate As Date = Trim(oRs9.Fields.Item("StartDate").Value.ToString())
                                    'Dim Enddate As Date = Trim(oRs9.Fields.Item("EndDate").Value.ToString())
                                    If oRs9.RecordCount > 0 Then
                                        For p1 As Integer = 1 To oRs9.RecordCount
                                            Dim Update1 As String = " UPDATE T1 SET  ""U_Apnum1""='" & oRsGetAPInvoiceEntry1.Fields.Item(2).Value.ToString & "' FROM ""RPC1""T1 INNER JOIN ORPC T2 ON T2.""DocEntry""=T1.""DocEntry""   WHERE T2.""CardCode""='" & oGrid.DataTable.GetValue("Supplier Code", K) & "' and T2.""DocNum""='" & oRs9.Fields.Item("DocNum").Value & "'and  T2.""DocDate"" between '" & FromDate & "'and '" & ToDate & "' and T1.""U_FOC""='N' and ""LineNum""='" & oRs9.Fields.Item("LineNum").Value & "' and ifnull(T1.""U_Apnum1"",'')='' "

                                            oRsUpdate1.DoQuery(Update1)
                                            oRs9.MoveNext()
                                        Next

                                    End If
                                End If
                            End If
                        Next
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oPurchaseCreditNotes)


                    Else
                        'oGrid.DataTable.SetValue("Ap Credit memo", i, objMain.objCompany.GetLastErrorDescription)
                        oGrid.DataTable.SetValue("Status", i, objMain.objCompany.GetLastErrorDescription)
                        objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorCode & " : " & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oPurchaseCreditNotes)
                        Exit Sub
                    End If

                End If

            Next
            objMain.objApplication.MessageBox("AP Credit Memo Posted Successfully..!!!")
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub PostAPCreditMemo(ByVal FormUID As String)
        Dim oProgressPostAPCreditNote As SAPbouiCOM.ProgressBar = Nothing
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim FromDate As String = objForm.Items.Item("FDATE").Specific.Value
            Dim ToDate As String = objForm.Items.Item("TDATE").Specific.Value
            Dim objComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("BRNH").Specific
            Dim cardcode As String
            If objComboBox.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Please select branch at the  header...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            Dim PostDate As String = objForm.Items.Item("PDate").Specific.Value.ToString().Trim()
            If PostDate = "" Then
                objMain.objApplication.SetStatusBarMessage("Please enter the Posting Date", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Exit Sub
            End If
            oGrid = objForm.Items.Item("Item_8").Specific
            Dim Supplier, Supplier1 As String
            Supplier1 = ""
            objForm.Freeze(False)
            Dim oPurchaseCreditNotes As SAPbobsCOM.Documents
            oProgressPostAPCreditNote = objMain.objApplication.StatusBar.CreateProgressBar("Please Wait AP CreditNote docments generating....!", oGrid.DataTable.Rows.Count - 1, False)
            oProgressPostAPCreditNote.Text = "Please Wait AP CreditNote docments generating....!"
            Dim j1 As Integer = 1

            For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                Supplier = oGrid.DataTable.GetValue("Supplier Code", i)
                If Supplier1.Contains(Supplier) = False Then
                    Supplier1 = Supplier1 & "," & Supplier
                    oPurchaseCreditNotes = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)


                    'once posted second time not posted code here....!
                    'Dim oChkRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'oChkRs.DoQuery("SELECT ""DocEntry"" FROM ORPC WHERE ""CardCode""='" & Supplier & "' AND ""DocDate"" BETWEEN '" & FromDate & "' AND '" & ToDate & "' AND ""NumAtCard""='Rebate Generation'")
                    'If oChkRs.RecordCount > 0 Then
                    '    oGrid.DataTable.SetValue("Status", i, "Already Posted")
                    '    objMain.objApplication.MessageBox("AP Credit Memo already created for this Vendor and Date Range")
                    '    System.Runtime.InteropServices.Marshal.ReleaseComObject(oChkRs)
                    '    Exit Sub
                    'End If
                    'System.Runtime.InteropServices.Marshal.ReleaseComObject(oChkRs)


                    oPurchaseCreditNotes.CardCode = oGrid.DataTable.GetValue("Supplier Code", i)
                    oPurchaseCreditNotes.BPL_IDAssignedToInvoice = objForm.Items.Item("BRNH").Specific.Value
                    Dim currency As String = oGrid.DataTable.GetValue("Currency", i)
                    If currency = "AED" Then
                        oPurchaseCreditNotes.DocCurrency = "AED"
                    ElseIf currency = "USD" Then
                        oPurchaseCreditNotes.DocCurrency = "USD"
                    ElseIf currency = "OMR" Then
                        oPurchaseCreditNotes.DocCurrency = "OMR"
                    Else
                        oPurchaseCreditNotes.DocCurrency = "YER"
                    End If
                    Dim sDate As String = objForm.Items.Item("PDate").Specific.Value.ToString().Trim()
                    ' Dim sDate As String = DateTime.Now.ToString("yyyyMMdd")
                    sDate = sDate.Insert("4", "-")
                    sDate = sDate.Insert("7", "-")
                    oPurchaseCreditNotes.DocDate = sDate
                    oPurchaseCreditNotes.DocDueDate = sDate
                    oPurchaseCreditNotes.TaxDate = sDate
                    Dim invoiceNumbers As New List(Of String)

                    For x As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                        Dim DOCNUMX As String = oGrid.DataTable.GetValue("AP Invoice Number", x).ToString()

                        If Not String.IsNullOrEmpty(DOCNUMX) Then
                            invoiceNumbers.Add(DOCNUMX)
                        End If
                    Next

                    Dim distinctInvoices = invoiceNumbers.Distinct()

                    oPurchaseCreditNotes.NumAtCard = "Rebate Generation - (" & String.Join(",", distinctInvoices) & ")"
                    oPurchaseCreditNotes.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service

                    For j As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                        If oGrid.DataTable.GetValue("Check", j) = "Y" And oGrid.DataTable.GetValue("Ap Credit memo", j) = "" Then
                            If Supplier = oGrid.DataTable.GetValue("Supplier Code", j) Then
                                Dim Depart As String = "select ""OcrCode"",""OcrName"" From OOCR where ""OcrName""='" & oGrid.DataTable.GetValue("Department", j) & "'and ""DimCode""='1' "
                                Dim Departrs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Departrs.DoQuery(Depart)
                                Dim sect As String = "select ""OcrCode"",""OcrName"" From OOCR where ""OcrName""='" & oGrid.DataTable.GetValue("Section", j) & "'and ""DimCode""='2' "
                                Dim sectrs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                sectrs.DoQuery(sect)
                                oPurchaseCreditNotes.Lines.ItemDescription = "Rebate Generation"
                                oPurchaseCreditNotes.Lines.AccountCode = oGrid.DataTable.GetValue("GL Account", j)
                                oPurchaseCreditNotes.Lines.LineTotal = oGrid.DataTable.GetValue("Rebate Amount", j)
                                oPurchaseCreditNotes.Lines.CostingCode = Departrs.Fields.Item(0).Value
                                oPurchaseCreditNotes.Lines.CostingCode2 = sectrs.Fields.Item(0).Value
                                oPurchaseCreditNotes.Lines.UserFields.Fields.Item("U_Apnum1").Value = oGrid.DataTable.GetValue("AP Invoice Number", j)


                                '------Madhu 12/07/2026 VatGroup value is updating in Ap credit memo screen---------------'


                                oPurchaseCreditNotes.Lines.VatGroup = oGrid.DataTable.GetValue("VatGroup", j)

                                'oPurchaseCreditNotes.Lines.UserFields.Fields.Item("U_Subfamily").Value = oGrid.DataTable.GetValue("Sub Family", j)
                                oPurchaseCreditNotes.Lines.Add()
                            End If
                        End If
                    Next

                    If oPurchaseCreditNotes.Add = 0 Then
                        Dim GetAPInvoiceEntry1 As String = "select ""CardCode"",""DocEntry"",""DocNum"" From ORPC where ""DocEntry""=(Select Max(""DocEntry"") From ORPC ) "
                        Dim oRsGetAPInvoiceEntry1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetAPInvoiceEntry1.DoQuery(GetAPInvoiceEntry1)
                        For K As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                            If oGrid.DataTable.GetValue("Check", K) = "Y" Then
                                If oGrid.DataTable.GetValue("Supplier Code", K) = oRsGetAPInvoiceEntry1.Fields.Item(0).Value.ToString Then
                                    oGrid.DataTable.SetValue("Ap Credit memo", K, oRsGetAPInvoiceEntry1.Fields.Item(1).Value.ToString)
                                    oGrid.DataTable.SetValue("Status", K, "Posted Successfully")
                                    '--Ap invoice updation
                                    Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim oRs1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    'Dim ItmGroup1 As String = "select  distinct Y.""LineTotal"", Y.""LineNum"",Y.""ItemCode"" ,X.""DocNum""from OpcH X inner join PCH1 Y On X.""DocEntry""=Y.""DocEntry"" inner Join  ""OOAT"" T3 on T3.""BpCode""=X.""CardCode"" " &
                                    '"inner Join  ""@SBO_SUPREBATE"" T1 on T3.""Number""=T1. ""U_BASENO"" inner Join ""@SBO_SUPREBATE_C0"" T2 on T2.""DocEntry"" = T1.""DocEntry""where X.""CardCode""='" & oGrid.DataTable.GetValue("Supplier Code", K) & "' " &
                                    '"And Y.""ItemCode""In (Select ""ItemCode"" from OiTM where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB"" WHERE ""ItmsGrpNam"" = T2.""U_Depart"")  And ""U_Section""=(Select distinct ""U_SECNME"" from  ""@TNX_MERCHIRCY""  " &
                                    '"where ""U_SECNME""=T2.""U_Section"")and ""U_VSubGroup""=(Select   distinct ""U_FNAME"" from  ""@TNX_MERCHIRCY""where ""U_FNAME""=T2.""U_Family"")and ""U_VSubSubGroup""=(Select   distinct ""U_SBFNME"" from  ""@TNX_MERCHIRCY""where ""U_SBFNME"" =T2.""U_SFamily""))And Y.""U_FOC""='N' and X.""CANCELED"" ='N' " &
                                    ''"and  X.""DocDate"" between '" & FromDate & "'and '" & ToDate & "'and T2.""U_Depart""='" & oGrid.DataTable.GetValue("Department", K) & "'and  T2.""U_Section""='" & oGrid.DataTable.GetValue("Section", K) & "'and  T2.""U_Family""='" & oGrid.DataTable.GetValue("Family", K) & "'and  T2.""U_SFamily""='" & oGrid.DataTable.GetValue("Sub Family", K) & "' " &
                                    '"and ifnull(Y.""U_Apnum1"",'')='' "


                                    'Madhu 3/2/2026

                                    'Dim ItmGroup1 As String = "select  distinct Y.""LineTotal"", Y.""LineNum"",Y.""ItemCode"" ,X.""DocNum"",T3.""StartDate"",T3.""EndDate"",T3.""Number"" from OPCH X inner join PCH1 Y On X.""DocEntry""=Y.""DocEntry"" inner Join  ""OOAT"" T3 on T3.""BpCode""=X.""CardCode""and  Y.""AgrNo"" = T3.""AbsID"" " &
                                    '"inner Join  ""@SBO_SUPREBATE"" T1 on T3.""Number""=T1. ""U_BASENO"" inner Join ""@SBO_SUPREBATE_C0"" T2 on T2.""DocEntry"" = T1.""DocEntry""where X.""CardCode""='" & oGrid.DataTable.GetValue("Supplier Code", K) & "' " &
                                    '"And Y.""ItemCode""In (Select ""ItemCode"" from OITM T where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB"" WHERE ""ItmsGrpNam"" = T2.""U_Depart"") And T.""U_Section""=(Case When IFNULL(T2.""U_Section"",'')<>'' Then T2.""U_Section"" Else T.""U_Section"" End)And T.""U_VSubGroup""= (Case when IFNULL(T2.""U_Family"",'')<>'' Then T2.""U_Family"" Else T.""U_VSubGroup"" End)And T.""U_VSubSubGroup""=(Case When IFNULL(T2.""U_SFamily"",'')<>'' Then T2.""U_SFamily"" Else T.""U_VSubSubGroup"" End) )And Y.""U_FOC""='N'And T2.""U_ACTV""='Y' And  T3.""Status""='A' And ifnull(T2.""U_STUS"",'')='' and X.""CANCELED"" ='N' " &
                                    '"and  X.""DocDate"" between '" & FromDate & "'and '" & ToDate & "' and T2.""U_Depart""='" & oGrid.DataTable.GetValue("Department", K) & "'" &
                                    '"and ifnull(Y.""U_Apnum1"",'')='' "


                                    Dim ItmGroup1 As String = "select distinct Y.""LineTotal"", Y.""LineNum"", Y.""ItemCode"", X.""DocNum"", " &
                                                              "T3.""StartDate"", T3.""EndDate"", T3.""Number"" " & "from OPCH X " & "inner join PCH1 Y On X.""DocEntry"" = Y.""DocEntry"" " &
                                                               "inner Join ""OOAT"" T3 on T3.""BpCode"" = X.""CardCode"" and Y.""AgrNo"" = T3.""AbsID"" " & "inner Join ""@SBO_SUPREBATE"" T1 on T3.""Number"" = T1.""U_BASENO"" " &
                                                                 "inner Join ""@SBO_SUPREBATE_C0"" T2 on T2.""DocEntry"" = T1.""DocEntry"" " &
                                                               "where X.""CardCode"" = '" & oGrid.DataTable.GetValue("Supplier Code", K) & "' " & "and T3.""Status"" = 'A' " & "and ifnull(T2.""U_STUS"", '') = '' " &
                                                              "and X.""CANCELED"" = 'N' " & "and X.""DocDate"" between '" & FromDate & "' and '" & ToDate & "' " & "and ifnull(Y.""U_Apnum1"", '') = '' "


                                    oRs1.DoQuery(ItmGroup1)

                                    'Dim startdate1 As Date = Trim(oRs1.Fields.Item("StartDate").Value.ToString())
                                    'Dim Enddate1 As Date = Trim(oRs1.Fields.Item("EndDate").Value.ToString())
                                    If oRs1.RecordCount > 0 Then
                                        For p As Integer = 1 To oRs1.RecordCount
                                            Dim Update As String = " UPDATE T1 SET  ""U_Apnum1""='" & oRsGetAPInvoiceEntry1.Fields.Item(2).Value.ToString & "' FROM ""PCH1""T1 INNER JOIN OPCH T2 ON T2.""DocEntry""=T1.""DocEntry""   WHERE T2.""CardCode""='" & oGrid.DataTable.GetValue("Supplier Code", K) & "' and T2.""DocNum""='" & oRs1.Fields.Item("DocNum").Value & "'and  T2.""DocDate"" between '" & FromDate & "'and '" & ToDate & "' and T1.""U_FOC""='N' and ""LineNum""='" & oRs1.Fields.Item("LineNum").Value & "' and ifnull(T1.""U_Apnum1"",'')='' "

                                            Update = " UPDATE T1 SET  ""U_Apnum1""='" & oRsGetAPInvoiceEntry1.Fields.Item(2).Value.ToString & "' FROM ""PCH1""T1 INNER JOIN OPCH T2 ON T2.""DocEntry""=T1.""DocEntry""   WHERE T2.""CardCode""='" & oGrid.DataTable.GetValue("Supplier Code", K) & "' and T2.""DocNum""='" & oRs1.Fields.Item("DocNum").Value & "'and  T2.""DocDate"" between '" & FromDate & "'and '" & ToDate & "' and IFNULL(T1.""U_FOC"",'N')='N' and ""LineNum""='" & oRs1.Fields.Item("LineNum").Value & "' and ifnull(T1.""U_Apnum1"",'')='' "

                                            oRsUpdate.DoQuery(Update)
                                            oRs1.MoveNext()
                                        Next
                                    End If
                                    '--Ap Credit memo updation
                                    Dim oRsUpdate1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim oRs9 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    'madhu 3/2/2026
                                    '        Dim ItmGroup2 As String = "select  distinct Y.""LineTotal"", Y.""LineNum"",Y.""ItemCode"" ,X.""DocNum"",T3.""StartDate"",T3.""EndDate"",T3.""Number"" from ORPC X inner join RPC1 Y On X.""DocEntry""=Y.""DocEntry"" inner Join  ""OOAT"" T3 on T3.""BpCode""=X.""CardCode""and  Y.""AgrNo"" = T3.""AbsID""  " &
                                    '"inner Join  ""@SBO_SUPREBATE"" T1 on T3.""Number""=T1. ""U_BASENO"" inner Join ""@SBO_SUPREBATE_C0"" T2 on T2.""DocEntry"" = T1.""DocEntry""where X.""CardCode""='" & oGrid.DataTable.GetValue("Supplier Code", K) & "' " &
                                    '"And Y.""ItemCode""In (Select ""ItemCode"" from OITM T where (""ItmsGrpCod"") In(Select ""ItmsGrpCod"" FROM ""OITB"" WHERE ""ItmsGrpNam"" = T2.""U_Depart"")And T.""U_Section""=(Case When IFNULL(T2.""U_Section"",'')<>'' Then T2.""U_Section"" Else T.""U_Section"" End)And T.""U_VSubGroup""= (Case when IFNULL(T2.""U_Family"",'')<>'' Then T2.""U_Family"" Else T.""U_VSubGroup"" End)And T.""U_VSubSubGroup""=(Case When IFNULL(T2.""U_SFamily"",'')<>'' Then T2.""U_SFamily"" Else T.""U_VSubSubGroup"" End) )And Y.""U_FOC""='N'And T2.""U_ACTV""='Y' And  T3.""Status""='A' And ifnull(T2.""U_STUS"",'')='' and X.""CANCELED"" ='N' " &
                                    '"And  X.""DocDate"" between   '" & FromDate & "'and '" & ToDate & "' and T2.""U_Depart""='" & oGrid.DataTable.GetValue("Department", K) & "'" &
                                    '"And ifnull(Y.""U_Apnum1"",'')='' "
                                    '        oRs9.DoQuery(ItmGroup2)

                                    Dim ItmGroup2 As String = "SELECT DISTINCT Y.""LineTotal"", Y.""LineNum"", Y.""ItemCode"", X.""DocNum"", " &
                                                              "T3.""StartDate"", T3.""EndDate"", T3.""Number"" " & "FROM ORPC X " & "INNER JOIN RPC1 Y ON X.""DocEntry"" = Y.""DocEntry"" " &
                                                              "INNER JOIN ""OOAT"" T3 ON T3.""BpCode"" = X.""CardCode"" AND Y.""AgrNo"" = T3.""AbsID"" " & "INNER JOIN ""@SBO_SUPREBATE"" T1 ON T3.""Number"" = T1.""U_BASENO"" " &
                                                               "INNER JOIN ""@SBO_SUPREBATE_C0"" T2 ON T2.""DocEntry"" = T1.""DocEntry"" " & "WHERE X.""CardCode"" = '" & oGrid.DataTable.GetValue("Supplier Code", K) & "' " &
                                                               "AND T3.""Status"" = 'A' " & "AND IFNULL(T2.""U_STUS"", '') = '' " & "AND X.""CANCELED"" = 'N' " &
                                                                 "AND X.""DocDate"" BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & "AND IFNULL(Y.""U_Apnum1"", '') = '' "

                                    oRs9.DoQuery(ItmGroup2)


                                    'Dim startdate As Date = Trim(oRs9.Fields.Item("StartDate").Value.ToString())
                                    'Dim Enddate As Date = Trim(oRs9.Fields.Item("EndDate").Value.ToString())
                                    If oRs9.RecordCount > 0 Then
                                        For p1 As Integer = 1 To oRs9.RecordCount

                                            'Dim Update1 As String = " UPDATE T1 SET  ""U_Apnum1""='" & oRsGetAPInvoiceEntry1.Fields.Item(2).Value.ToString & "' FROM ""RPC1""T1 INNER JOIN ORPC T2 ON T2.""DocEntry""=T1.""DocEntry""   WHERE T2.""CardCode""='" & oGrid.DataTable.GetValue("Supplier Code", K) & "' and T2.""DocNum""='" & oRs9.Fields.Item("DocNum").Value & "'and  T2.""DocDate"" between '" & FromDate & "'and '" & ToDate & "' and T1.""U_FOC""='N' and ""LineNum""='" & oRs9.Fields.Item("LineNum").Value & "' and ifnull(T1.""U_Apnum1"",'')='' "
                                            Dim Update1 As String = " UPDATE T1 SET  ""U_Apnum1""='" & oRsGetAPInvoiceEntry1.Fields.Item(2).Value.ToString & "' FROM ""RPC1""T1 INNER JOIN ORPC T2 ON T2.""DocEntry""=T1.""DocEntry""   WHERE T2.""CardCode""='" & oGrid.DataTable.GetValue("Supplier Code", K) & "' and T2.""DocNum""='" & oRs9.Fields.Item("DocNum").Value & "'and  T2.""DocDate"" between '" & FromDate & "'and '" & ToDate & "'  and ""LineNum""='" & oRs9.Fields.Item("LineNum").Value & "' and ifnull(T1.""U_Apnum1"",'')='' "

                                            oRsUpdate1.DoQuery(Update1)
                                            oRs9.MoveNext()
                                        Next

                                    End If
                                End If
                            End If
                        Next
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oPurchaseCreditNotes)


                    Else
                        'oGrid.DataTable.SetValue("Ap Credit memo", i, objMain.objCompany.GetLastErrorDescription)
                        oGrid.DataTable.SetValue("Status", i, objMain.objCompany.GetLastErrorDescription)
                        objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorCode & " : " & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oPurchaseCreditNotes)
                        Exit Sub
                    End If

                End If
                oProgressPostAPCreditNote.Value = j1
                j1 = j1 + 1
            Next
            objForm.Freeze(False)
            objMain.objApplication.MessageBox("AP Credit Memo Posted Successfully..!!!")
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        Finally
            oGrid.AutoResizeColumns()
            If oProgressPostAPCreditNote IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgressPostAPCreditNote)
            End If
        End Try
    End Sub
    Public Sub TotalAmount(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim TotalSum As Double = 0
            For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If oGrid.DataTable.GetValue("Check", i) = "Y" Then
                    TotalSum = TotalSum + CDbl(oGrid.DataTable.GetValue("Rebate Amount", i))
                End If
            Next
            objForm.Items.Item("TlAmt").Specific.Value = TotalSum
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    'Private Sub APcreditmemoGeneration()
    '    Dim DocEntry As String = ""
    '    Dim oARCreditMemo As SAPbobsCOM.Documents
    '    Dim rsA As SAPbobsCOM.Recordset
    '    Dim rsUpdate As SAPbobsCOM.Recordset
    '    Dim rsDocEntry As SAPbobsCOM.Recordset
    '    Dim rsBATCH As SAPbobsCOM.Recordset
    '    Try
    '        Dim BATCHQUERY As String = " CALL ""SBO_ARCREDITCREATION_DOCNUM""()"
    '        rsBATCH = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        rsBATCH.DoQuery(BATCHQUERY)
    '        For k As Integer = 0 To rsBATCH.RecordCount - 1
    '            Dim NoOfTimes As Integer = rsBATCH.Fields.Item("NoOfTimes").Value
    '            DocEntry = rsBATCH.Fields.Item("DocEntry").Value
    '            For q As Integer = 1 To NoOfTimes
    '                If objCompany.InTransaction = True Then
    '                    objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
    '                End If
    '                objCompany.StartTransaction()
    '                Dim txn As String = "CALL ""SBO_ARCREDITMEMOCREATION""('" & DocEntry & "')"
    '                rsA = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                rsA.DoQuery(txn)
    '                If rsA.RecordCount > 0 Then
    '                    AddToFileIN("Started AR credit memo creation for LS DocEntry --" & DocEntry & "--" & System.DateTime.Now())

    '                    oARCreditMemo = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
    '                    oARCreditMemo.BPCode = rsA.Fields.Item("U_BPCD").Value
    '                    oARCreditMemo.BPName = rsA.Fields.Item("U_BPNM").Value
    '                    oARCreditMemo.FromDate = rsA.Fields.Item("U_FDATE").Value
    '                    oARCreditMemo.ToDate = rsA.Fields.Item("U_TDATE").Value
    '                    oARCreditMemo.InvoiceTotalDiscount = rsA.Fields.Item("U_INVTLDSC").Value
    '                    oARCreditMemo.InvoiceTotalAmount = rsA.Fields.Item("U_ITAmt").Value

    '                    'oARCreditMemo.Comments = "test"
    '                    ' oARCreditMemo.JournalMemo = rsA.Fields.Item("U_RCPT").Value
    '                    If rsA.Fields.Item("U_BRNH").Value <> "" Then
    '                        oARCreditMemo.BPL_IDAssignedToInvoice = rsA.Fields.Item("U_BRNH").Value
    '                    End If

    '                    ' If rsA.Fields.Item("RondingConsider").Value = 0 Then
    '                    'oARCreditMemo.Rounding = SAPbobsCOM.BoYesNoEnum.tYES
    '                    ' oARCreditMemo.RoundingDiffAmount = rsA.Fields.Item("U_ROUNDAMT").Value
    '                    'End If

    '                    'For I As Integer = 0 To rsA.RecordCount - 1
    '                    '    oARCreditMemo.Lines.ItemCode = rsA.Fields.Item("U_ITEMCODE").Value
    '                    '    oARCreditMemo.Lines.Quantity = rsA.Fields.Item("U_QTY").Value
    '                    '    Dim UnitPrice As Double = rsA.Fields.Item("U_UPRICE").Value
    '                    '    oARCreditMemo.Lines.UnitPrice = UnitPrice
    '                    '    oARCreditMemo.Lines.DiscountPercent = rsA.Fields.Item("DiscountPercent").Value
    '                    '    oARCreditMemo.Lines.TaxCode = "X0"
    '                    '    'oARCreditMemo.Lines.TaxCode = "X0"
    '                    '    oARCreditMemo.Lines.TaxCode = rsA.Fields.Item("U_TCODE").Value
    '                    '    oARCreditMemo.Lines.UoMEntry = rsA.Fields.Item("UomEntry").Value

    '                    '    oARCreditMemo.Lines.LineTotal = rsA.Fields.Item("U_LNTOTAL").Value
    '                    '    oARCreditMemo.Lines.UserFields.Fields.Item("U_LSLINE").Value = rsA.Fields.Item("LineId").Value.ToString().Trim()
    '                    '    oARCreditMemo.Lines.UserFields.Fields.Item("U_LSENTRY").Value = rsA.Fields.Item("DocEntry").Value.ToString().Trim()
    '                    '    oARCreditMemo.Lines.Add()
    '                    '    rsA.MoveNext()
    '                    'Next
    '                    Dim M As Integer = oARCreditMemo.Add()
    '                    If M = 0 Then
    '                        'Threading.Thread.Sleep(3000)
    '                        Dim DocEntryQry As String = " SELECT MAX(T0.""DocEntry"") FROM ""ORIN"" T0 INNER JOIN ""RIN1"" T1 ON T0.""DocEntry""=T1.""DocEntry"" " &
    '                            " Where T1.""U_LSENTRY""='" & DocEntry & "' "
    '                        rsDocEntry = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                        rsDocEntry.DoQuery(DocEntryQry)
    '                        Dim CMDocEntry As String = rsDocEntry.Fields.Item(0).Value
    '                        System.Runtime.InteropServices.Marshal.ReleaseComObject(rsDocEntry)
    '                        Me.AddToFileIN("Successfully created AR Credit Memo with DocEntry--" & CMDocEntry & "--For LS DocEntry--" & DocEntry & "--" & System.DateTime.Now())
    '                        Try
    '                            Dim UpdateQry As String = "Update ""@10X_LCN1"" SET ""U_STATUS""='" & CMDocEntry & "' WHERE ""DocEntry""='" & DocEntry & "'" &
    '                            " AND ""LineId"" IN (Select ""U_LSLINE"" From RIN1 Where ""DocEntry""='" & CMDocEntry & "' AND IFNULL(""U_LSLINE"",'')<>'') "
    '                            rsUpdate = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                            rsUpdate.DoQuery(UpdateQry)
    '                            System.Runtime.InteropServices.Marshal.ReleaseComObject(rsUpdate)
    '                        Catch ex As Exception
    '                            Me.AddToFileIN("Error while updating LS Table for LS DocEntry -- " & DocEntry & "--" & System.DateTime.Now())
    '                        End Try
    '                        Try
    '                            objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
    '                        Catch ex As Exception
    '                            Me.AddToFileIN("Error while commiting transaction for DocEntry--" & CMDocEntry & "--for LS DocEntry--" & DocEntry & "--" & System.DateTime.Now())
    '                        End Try
    '                    Else
    '                        If objCompany.InTransaction = True Then
    '                            objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
    '                        End If

    '                        Dim test As String = objCompany.GetLastErrorDescription()
    '                        Me.AddToFileIN("Error In Creation AR CREDIT MEMO for LS DocEntry--" & DocEntry & "--" & objCompany.GetLastErrorDescription() & "'" & System.DateTime.Now())
    '                        Dim ErrMessage As String = objCompany.GetLastErrorDescription()
    '                        If ErrMessage.Length() > 254 Then
    '                            ErrMessage.Substring(0, 253)
    '                        End If
    '                        Dim UpdateErrQry As String = " update ""@10X_OLCN"" SET ""U_SAPSTATUS""='" & ErrMessage & "' WHERE ""DocEntry""='" & DocEntry & "' "
    '                        Dim rsUpdateError As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                        rsUpdateError.DoQuery(UpdateErrQry)
    '                        System.Runtime.InteropServices.Marshal.ReleaseComObject(rsUpdateError)
    '                    End If
    '                End If
    '                '-------Releasing the objects
    '                Try
    '                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oARCreditMemo)
    '                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rsA)
    '                Catch ex As Exception
    '                    Me.AddToFileIN("Error while releasing the comobjects -- '" & objCompany.GetLastErrorDescription() & "'" & System.DateTime.Now())
    '                End Try
    '            Next
    '            rsBATCH.MoveNext()
    '        Next
    '    Catch ex As Exception
    '        If objCompany.InTransaction() = True Then
    '            objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
    '        End If

    '        Dim ErrMessage As String = ex.Message
    '        If ErrMessage.Length() > 254 Then
    '            ErrMessage.Substring(0, 253)
    '        End If
    '        Dim QRY As String = " update ""@10X_OLCN"" SET ""U_SAPSTATUS""='" & ErrMessage & "' WHERE ""DocEntry""='" & DocEntry & "' "
    '        Dim rscode1 As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        rscode1.DoQuery(QRY)
    '        AddToFileIN(Err.Description)
    '    Finally
    '        System.Runtime.InteropServices.Marshal.ReleaseComObject(rsBATCH)
    '    End Try
    'End Sub
    Private Sub LoadSectionGridCost(ByVal FormUID As String, ByVal Value1 As String, ByVal Value As String)
        Try
            objForm.Freeze(True)
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("Item_4").Specific
            oDt2 = objForm.DataSources.DataTables.Item("DT_1")
            oGrid.DataTable = oDt2
            Try
                oGrid.DataTable.Clear()
            Catch ex As Exception
            End Try
            ' Dim Section As String = "select T0.""Code"" as ""Section Code"",T0.""Name""as ""Section Name"",'" & Value & "' as ""Check"" from ""@VR_SECTION"" T0 "
            'Dim Section As String = "select distinct T0.""U_SECNME"" as ""Section Code"",T0.""U_SECNME""as ""Section Name"",'" & Value & "' as ""Check"" from ""@TNX_MERCHIRCY"" T0 "
            ' Dim CreationCode As String = objForm.Items.Item("Crtcode").Specific.Value
            Dim Section As String = "select distinct T0.""U_SECNME"" as ""Section Code"",T0.""U_SECNME""as ""Section Name"",T0. ""U_FNAME""as ""Family"",T0.""U_SBFNME""as ""SubbFamily"",'" & Value & "' as ""Check"" from ""@TNX_MERCHIRCY"" T0 where T0.""U_CRECDE""='" & Value1 & "' "
            'oGrid.DataTable.ExecuteQuery(Section1)



            oGrid.DataTable.ExecuteQuery(Section)
            oGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.AutoResizeColumns()
            oGrid.Columns.Item("Section Code").Editable = False
            oGrid.Columns.Item("Section Name").Editable = False

            objForm.Freeze(False)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub LoadDepartmentGridCost(ByVal FormUID As String, ByVal Value1 As String, ByVal Value As String)
        Try
            objForm.Freeze(True)
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("Item_3").Specific
            oDt1 = objForm.DataSources.DataTables.Item("DT_0")
            oGrid.DataTable = oDt1
            Try
                oGrid.DataTable.Clear()
            Catch ex As Exception
            End Try
            ' Dim Department As String = "SELECT T0.""U_DEPTNME"" as ""Department Code"" ,T0.""U_SUBDEPT"" as ""Department Name"", T0.""U_SECNME"", T0.""U_FNAME"", T0.""U_SBFNME"",'" & Value & "' as ""Check""  FROM ""@TNX_MERCHIRCY""  T0"
            ' Dim Department As String = "SELECT distinct T0.""U_DEPTNME"" as ""Department Code"" ,T0.""U_DEPTNME"" as ""Department Name"", '" & Value & "' as ""Check""  FROM ""@TNX_MERCHIRCY""  T0"


            ' Dim Department As String="select T0.""PrcCode"" as ""Department Code"",T0.""PrcName""as ""Department Name"",'" & Value & "'  as ""Check"" from ""OPRC"" T0  WHERE T0.""Locked""='N' "
            ' Dim CreationCode As String = objForm.Items.Item("Crtcode").Specific.Value
            Dim Department As String = "select distinct T0.""U_DEPTNME"" as ""Department Code"",T0.""U_DEPTNME""as ""Department Name"",'" & Value & "' as ""Check"" from ""@TNX_MERCHIRCY"" T0  where T0.""U_CRECDE""='" & Value1 & "' "

            oGrid.DataTable.ExecuteQuery(Department)
            oGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.AutoResizeColumns()
            oGrid.Columns.Item("Department Code").Editable = False
            oGrid.Columns.Item("Department Name").Editable = False

            objForm.Freeze(False)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
End Class






