
'NEW
'MM
Imports System.Collections.Specialized
Imports System.Configuration
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports SAPbobsCOM
Public Class clsItemWiseDiscount
#Region "       Declaration             "
    Public objForm, objNewForm, promoName, oEditText, objSTRForm As SAPbouiCOM.Form
    Dim objMatrix, objMatrix1, objMatrix2, objMatrix3, objMatrix4, omtx As SAPbouiCOM.Matrix
    Dim objComboBox, OBJCOMBO, OBJCOMBO1 As SAPbouiCOM.ComboBox
    Dim ItemGroup, subGroup, ItemWise, ItemSubGroup, sQury2, sQury3 As String
    Dim oDt As SAPbouiCOM.DataTable
    Dim oGrid As SAPbouiCOM.Grid
    Public oDS, oDBs_Head, oDBs_Details1, oDS1, oDS2, oDS3, oDS4 As SAPbouiCOM.DBDataSource
    Public PrgBar As SAPbouiCOM.ProgressBar
    Dim LostFocusFlag As Boolean = True
    Dim MATRIXS As String
    'Dim objBarcode As New List(Of String)
    Dim cdiscountpercentage, cofferprice As String
#End Region
    Sub CreateForm(ByVal DocNum As String)
        Try

            objMain.objUtilities.LoadForm("ItemWiseDiscounts.xml", "ITEM", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ITEM", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            oDS = objForm.DataSources.DBDataSources.Item("@TNX_ITEMDISC")
            oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMDISC_C0")
            oDS2 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMDISC_C1")
            oDS3 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMDISC_C2")
            oDS4 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMDISC_C3")
            'objForm.EnableMenu("774", True)
            Me.EnableCoulmns(objForm.UniqueID)

            'objForm.DataSources.UserDataSources.Add("U_TYPE", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            'objForm.DataSources.UserDataSources.Add("U_ITGRP", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            'objForm.DataSources.UserDataSources.Add("U_SBGRP", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            'objForm.DataSources.UserDataSources.Add("U_SBSBGRP", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            objMatrix1 = objForm.Items.Item("IMtx").Specific
            objMatrix1.AutoResizeColumns()
            objMatrix2 = objForm.Items.Item("IGMtx1").Specific
            objMatrix2.AutoResizeColumns()
            objMatrix3 = objForm.Items.Item("EMtx").Specific
            objMatrix3.AutoResizeColumns()
            objMatrix4 = objForm.Items.Item("BMtx").Specific
            objMatrix4.AutoResizeColumns()
            'Me.GetNextDocNum()
            objComboBox = objForm.Items.Item("PRTYP").Specific
            'objComboBox.DataBind.SetBound(True, "", "U_TYPE")


            objForm.Items.Item("PRTYP").AffectsFormMode = False
            objComboBox.ValidValues.Add("All", "All")
            objComboBox.ValidValues.Add("ItemWise", "Item Wise")




            'objComboBox.ValidValues.Add("SubGroup", "Sub Group")
            'objComboBox.ValidValues.Add("SubSubGroup", "Sub Sub Group")
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            ' objComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

            objForm.Items.Item("Pro").Width = objForm.Items.Item("Pro").Width + 140
            objForm.Items.Item("Item_0").Left = objForm.Items.Item("Item_0").Left + 140
            objForm.Items.Item("Item_1").Left = objForm.Items.Item("Item_1").Left + 140
            objForm.Items.Item("Pro").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            objComboBox = objForm.Items.Item("ITMGRP").Specific
            'objComboBox.DataBind.SetBound(True, "", "U_ITGRP")
            objForm.Items.Item("ITMGRP").AffectsFormMode = False
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, "Select ""ItmsGrpCod""  , ""ItmsGrpNam""  From OITB Order By  ""ItmsGrpNam""")
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("ITMGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("ITMGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("ITMGRP").DisplayDesc = True







            oDt = objForm.DataSources.DataTables.Add("dt1")
            oDt = objForm.DataSources.DataTables.Item("dt1")

            objMain.objApplication.StatusBar.SetText("Form loaded successfully.....continue.!!.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized




            objForm.Items.Item("PCRDoc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("PCRDoc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("PCRDoc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("PCRDoc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)


            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("FRMDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("FRMDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("FRMDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("Login").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Login").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Login").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Login").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            'objForm.Items.Item("TODATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("TODATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("TODATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("BRANCH").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("BRANCH").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("BRANCH").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            Me.BranchLoad(objForm.UniqueID)

            objForm.Items.Item("Path").Visible = True
            objForm.Items.Item("Upload").Visible = True
            objForm.Items.Item("Item_3").Visible = True
            'objMatrix.Items.Item("Item_3").Visible = True


            Dim oComb As SAPbouiCOM.ComboBox = objForm.Items.Item("BRANCH").Specific

            'objMain.objUtilities.ComboBoxLoadValues(oComb, "Select ""BPLId"",""BPLName"" FROM OBPL Where ""BPLId""<>'1'")
            objMain.objUtilities.ComboBoxLoadValues(oComb, "Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join " & " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId"" ")

            oComb.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("BRANCH").DisplayDesc = True


            'Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            Me.objForm.Items.Item("DOCDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Me.objForm.Items.Item("DOCDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Me.objForm.Items.Item("DOCDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Me.objForm.Items.Item("DOCDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Dim sQury As String = "Select ""ItmsGrpCod"" , ""ItmsGrpNam""   From OITB  "

            objComboBox = objForm.Items.Item("ITMGRP").Specific

            objMain.objUtilities.ComboBoxLoadValues(objComboBox, sQury)
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            'objForm.Items.Item("ITMGRP").DisplayDesc = True
            Dim ItemGroup As String
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Private Sub EnableCoulmns(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            omtx = objForm.Items.Item("IMtx").Specific
            Try
                omtx.Columns.Item("OFPRWVAT").Editable = True

            Catch ex As Exception
            End Try
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub




    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Dim DocNum As String = ""
            If pVal.MenuUID = "ME_ITEMA" And pVal.BeforeAction = False Then
                Me.CreateForm(DocNum)
                Me.SetDefault(objForm.UniqueID)
                Me.BranchLoad(objForm.UniqueID)

            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                'Me.GetNextDocNum()
                Me.SetDefault(objForm.UniqueID)
                Me.BranchLoad(objForm.UniqueID)

            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then

                Me.SetNewLine(objForm.UniqueID, MATRIXS)
                'ElseIf pVal.MenuUID = "774" And pVal.BeforeAction = False Then
                '    objForm = objMain.objApplication.Forms.GetForm("ITEM", objMain.objApplication.Forms.ActiveForm.TypeCount)
                '    objMatrix = objForm.Items.Item("IMtx").Specific
                '    objMatrix3 = objForm.Items.Item("EMtx").Specific
                '    Dim row As Integer = objMatrix.VisualRowCount
                '    If MATRIXS.Equals("IMtx") = True Then
                '        If objMatrix.IsRowSelected(1) <> True And objMatrix.VisualRowCount < 1 Then
                '            objMatrix.AddRow()
                '            oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
                '            objMatrix.SetLineData(objMatrix.VisualRowCount)
                '        End If
                '        If objMatrix.IsRowSelected(row) = True Then
                '            objMatrix.DeleteRow(row)
                '        Else
                '            For i As Integer = 1 To objMatrix.VisualRowCount - 1
                '                If objMatrix.IsRowSelected(i) = True Then
                '                    '  If i <> 1 Then
                '                    objMatrix.DeleteRow(i)
                '                    '   End If
                '                End If
                '            Next
                '        End If
                '        For i As Integer = 1 To objMatrix.VisualRowCount
                '            objMatrix.Columns.Item("LineId").Cells.Item(i).Specific.Value = i
                '            'objMatrix.FlushToDataSource()
                '        Next
                '    ElseIf MATRIXS.Equals("EMtx") = True Then
                '        If objMatrix3.IsRowSelected(1) <> True And objMatrix3.VisualRowCount < 1 Then
                '            objMatrix3.AddRow()
                '            oDS4.SetValue("LineId", oDS1.Offset, objMatrix3.VisualRowCount)
                '            objMatrix3.SetLineData(objMatrix3.VisualRowCount)
                '        End If
                '        If objMatrix3.IsRowSelected(row) = True Then
                '            objMatrix3.DeleteRow(row)
                '        Else
                '            For i As Integer = 1 To objMatrix3.VisualRowCount - 1
                '                If objMatrix3.IsRowSelected(i) = True Then
                '                    '  If i <> 1 Then
                '                    objMatrix3.DeleteRow(i)
                '                    '   End If
                '                End If
                '            Next
                '        End If
                '        For i As Integer = 1 To objMatrix3.VisualRowCount
                '            objMatrix3.Columns.Item("LineId").Cells.Item(i).Specific.Value = i
                '        Next
                '    End If
                '    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                '        objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                '    End If

                'ElseIf pVal.MenuUID = "Delete Row" And pVal.BeforeAction = False Then
                '    objForm = objMain.objApplication.Forms.GetForm("ITEM", objMain.objApplication.Forms.ActiveForm.TypeCount)
                '    objMatrix = objForm.Items.Item("IMtx").Specific
                '    For i As Integer = 1 To objMatrix.VisualRowCount - 1
                '        If objMatrix.IsRowSelected(i) = True Then
                '            objMatrix.DeleteRow(i)
                '        End If
                '    Next
                '    For i As Integer = 1 To objMatrix.VisualRowCount
                '        objMatrix.Columns.Item("LineId").Cells.Item(i).Specific.string = i
                '    Next
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    'objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                End If
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    ' item Cfl 
    Public Shared Sub CFL_Filter1(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef oApplication As SAPbouiCOM.Application, ByRef oCompany As SAPbobsCOM.Company, ByVal Query As String, ByVal Code As String, ByVal subgrp As String, ByVal subsubgrp As String, ByVal Itemcode As String)
        Try
            Dim oCLFEvent As SAPbouiCOM.ChooseFromListEvent = Nothing
            Dim oCell As SAPbouiCOM.Column
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFL_collection As SAPbouiCOM.ChooseFromListCollection
            Dim oCon As SAPbouiCOM.Condition = Nothing
            Dim emptyConds As SAPbouiCOM.Conditions = Nothing
            Dim oCons As SAPbouiCOM.Conditions = Nothing
            Dim oRec As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRec.DoQuery(Query)
            Dim strCFLID As String = Nothing
            oCLFEvent = CType((pVal), SAPbouiCOM.ChooseFromListEvent)
            strCFLID = oCLFEvent.ChooseFromListUID
            oCFL = oApplication.Forms.ActiveForm.ChooseFromLists.Item(strCFLID)
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            Dim rr As Integer = oRec.RecordCount


            If rr = 0 Then
                oCon = oCons.Add()
                oCon.[Alias] = Itemcode
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = oRec.Fields.Item(Itemcode).Value.ToString()
            Else


                For i As Integer = 0 To oRec.RecordCount - 1
                    oCon = oCons.Add()
                    oCon.[Alias] = Itemcode
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRec.Fields.Item(Itemcode).Value.ToString()


                    If i <> oRec.RecordCount - 1 Then
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    End If


                    oRec.MoveNext()
                Next
            End If


            oCFL.SetConditions(oCons)
        Catch ex As Exception

            'objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try


    End Sub


    Public Shared Sub CFL_Filter(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef oApplication As SAPbouiCOM.Application, ByRef oCompany As SAPbobsCOM.Company, ByVal Query As String, ByVal Code As String, ByVal subgrp As String, ByVal subsubgrp As String, ByVal Itemcode As String)
        Try
            Dim oCLFEvent As SAPbouiCOM.ChooseFromListEvent = Nothing
            Dim oCell As SAPbouiCOM.Column
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFL_collection As SAPbouiCOM.ChooseFromListCollection
            Dim oCon As SAPbouiCOM.Condition = Nothing
            Dim emptyConds As SAPbouiCOM.Conditions = Nothing
            Dim oCons As SAPbouiCOM.Conditions = Nothing
            Dim oRec As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRec.DoQuery(Query)
            Dim strCFLID As String = Nothing
            oCLFEvent = CType((pVal), SAPbouiCOM.ChooseFromListEvent)
            strCFLID = oCLFEvent.ChooseFromListUID
            oCFL = oApplication.Forms.ActiveForm.ChooseFromLists.Item(strCFLID)
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            Dim rr As Integer = oRec.RecordCount


            If rr = 0 Then
                oCon = oCons.Add()
                oCon.[Alias] = Code
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = oRec.Fields.Item(Code).Value.ToString()
            Else


                For i As Integer = 0 To oRec.RecordCount - 1
                    oCon = oCons.Add()
                    oCon.[Alias] = Code
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRec.Fields.Item(Code).Value.ToString()


                    If i <> oRec.RecordCount - 1 Then
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    End If


                    oRec.MoveNext()
                Next
            End If


            oCFL.SetConditions(oCons)
        Catch ex As Exception
            ' objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try


    End Sub
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If (pVal.ItemUID = "IMtx" Or pVal.ItemUID = "IGMtx1" Or pVal.ItemUID = "BMtx" Or pVal.ItemUID = "EMtx") And pVal.BeforeAction = False Then ' Or pVal.ItemUID = "Matrx5" Or pVal.ItemUID = "Matrix" Then
                        MATRIXS = pVal.ItemUID
                    End If
                    If (pVal.ItemUID = "IMtx" Or pVal.ItemUID = "IGMtx1" Or pVal.ItemUID = "BMtx" Or pVal.ItemUID = "EMtx") And pVal.BeforeAction = True Then ' Or pVal.ItemUID = "Matrx5" Or pVal.ItemUID = "Matrix" Then
                        MATRIXS = pVal.ItemUID
                    End If
                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    If pVal.ItemUID = "IGMtx1" And pVal.ColUID = "CHK" And pVal.BeforeAction = False Then
                        'If pVal.BeforeAction = False Then
                        objMatrix = Me.objForm.Items.Item("IGMtx1").Specific
                        Dim visualRowCount As Integer = Me.objMatrix.VisualRowCount
                        Dim checkBox As SAPbouiCOM.CheckBox = DirectCast(Me.objMatrix.Columns.Item("CHK").Cells.Item(1).Specific, SAPbouiCOM.CheckBox)
                        Dim Check As Boolean
                        If (Not checkBox.Checked) Then
                            Check = True
                        Else
                            Check = False
                        End If
                        For j As Integer = 1 To visualRowCount Step 1
                            checkBox = DirectCast(Me.objMatrix.Columns.Item("CHK").Cells.Item(j).Specific, SAPbouiCOM.CheckBox)
                            checkBox.Checked = Check
                        Next
                        'End If
                    End If
                    If pVal.ItemUID = "BMtx" And pVal.ColUID = "CHK" And pVal.BeforeAction = False Then
                        'If pVal.BeforeAction = False Then
                        objMatrix = Me.objForm.Items.Item("BMtx").Specific
                        Dim visualRowCount As Integer = Me.objMatrix.VisualRowCount
                        Dim checkBox As SAPbouiCOM.CheckBox = DirectCast(Me.objMatrix.Columns.Item("CHK").Cells.Item(1).Specific, SAPbouiCOM.CheckBox)
                        Dim Check As Boolean
                        If (Not checkBox.Checked) Then
                            Check = True
                        Else
                            Check = False
                        End If
                        For j As Integer = 1 To visualRowCount Step 1
                            checkBox = DirectCast(Me.objMatrix.Columns.Item("CHK").Cells.Item(j).Specific, SAPbouiCOM.CheckBox)
                            checkBox.Checked = Check
                        Next
                        'End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                         
                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And
     (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        Me.SetDefault(objForm.UniqueID)
                        'Me.GetNextDocNum()
                        'Me.BranchLoad(objForm.UniqueID)
                    End If
                    'If pVal.ItemUID = "1" AndAlso pVal.BeforeAction = True AndAlso (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                    (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                        'Dim bcode As String = oDt.GetValue("BcdCode", 0)
                        'For i As Integer = 1 To pVal.Row
                        '    If bcode = objMatrix.Columns.Item("BCODE").Cells.Item(i).Specific.Value Then
                        '        objMain.objApplication.StatusBar.SetText("This Barcode " & bcode & " Already Exist in row number " & i & "...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '        BubbleEvent = False
                        '        Exit Sub
                        '    End If
                        'Next





                        'If pVal.ItemUID = "1" AndAlso pVal.BeforeAction = True AndAlso (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Dim FromTime As String = CType(objForm.Items.Item("FRMTIME").Specific, SAPbouiCOM.EditText).Value
                        Dim ToTime As String = CType(objForm.Items.Item("TOTIME").Specific, SAPbouiCOM.EditText).Value

                        ' Dim tempDate As DateTime
                        'If oDS.GetValue("U_FRMDATE", 0) = "" Then
                        '    objMain.objApplication.StatusBar.SetText("Please Choose From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If
                        'If oDS.GetValue("U_TODATE", 0) = "" Then
                        '    objMain.objApplication.StatusBar.SetText("Please Choose To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If
                        'If oDS3.GetValue("U_CHK", 0) = "" Then
                        '    objMain.objApplication.StatusBar.SetText("Please Select Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If

                        Dim fromDate As DateTime
                        Dim toDate As DateTime

                        'If Not DateTime.TryParseExact(CType(objForm.Items.Item("FRMDATE").Specific, SAPbouiCOM.EditText).Value.Trim(), "yyyyMMdd", Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, fromDate) Then
                        '    objMain.objApplication.StatusBar.SetText("Please Choose 'From Date'. Please use format yyyyMMdd.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If
                        'If Not DateTime.TryParseExact(CType(objForm.Items.Item("TODATE").Specific, SAPbouiCOM.EditText).Value.Trim(), "yyyyMMdd", Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, toDate) Then
                        '    objMain.objApplication.StatusBar.SetText("Please Choose 'To Date'. Please use format yyyyMMdd.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If

                        If oDS.GetValue("U_FRMTIME", 0) = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter From Time", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If oDS.GetValue("U_TOTIME", 0) = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter To Time", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If





                    End If





                    '-----Akhila Overlapping code --
                    '                 If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                    '(pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                    '                     Dim FromTime As String = CType(objForm.Items.Item("FRMTIME").Specific, SAPbouiCOM.EditText).Value.Trim()
                    '                     Dim ToTime As String = CType(objForm.Items.Item("TOTIME").Specific, SAPbouiCOM.EditText).Value.Trim()
                    '                     Dim fromDate As String = objForm.Items.Item("FRMDATE").Specific.Value.Trim()
                    '                     Dim toDate As String = objForm.Items.Item("TODATE").Specific.Value.Trim()
                    '                     Dim docnum As String = objForm.Items.Item("DocNum").Specific.Value.Trim()
                    '                     Dim currentStatus As String = CType(objForm.Items.Item("STATUS").Specific, SAPbouiCOM.ComboBox).Selected.Value.Trim()
                    '                     Dim statustest As Boolean = False

                    '                     If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    '                         statustest = True
                    '                     ElseIf pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    '                         ' old status
                    '                         Dim oldStatusQuery As String = "SELECT ""U_STATUS"" FROM ""@TNX_ITEMDISC"" WHERE ""DocNum"" = '" & docnum & "'"
                    '                         Dim rsOldStatus As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                         rsOldStatus.DoQuery(oldStatusQuery)
                    '                         If rsOldStatus.RecordCount > 0 Then
                    '                             Dim oldStatus As String = rsOldStatus.Fields.Item("U_STATUS").Value.ToString().Trim()
                    '                             '  if status is being changed from Inactive to Active
                    '                             If oldStatus = "Inactive" And currentStatus = "Active" Then
                    '                                 statustest = True
                    '                             End If
                    '                         End If
                    '                     End If

                    '                     If statustest Then
                    '                         Dim matrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("IMtx").Specific, SAPbouiCOM.Matrix)
                    '                         For i As Integer = 1 To matrix.VisualRowCount
                    '                             matrix.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    '                             Dim Barcode As String = CType(matrix.Columns.Item("BCODE").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value.Trim()
                    '                             Dim UOM As String = CType(matrix.Columns.Item("UOM").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value.Trim()
                    '                             Dim itemCode As String = CType(matrix.Columns.Item("ITMCODE").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value.Trim()
                    '                             If itemCode <> "" Then
                    '                                 Dim query As String = "SELECT ""DocNum"" FROM ""@TNX_ITEMDISC"" H1 " &
                    '                 "INNER JOIN ""@TNX_ITEMDISC_C0"" C1 ON H1.""DocEntry"" = C1.""DocEntry"" " &
                    '                 "WHERE C1.""U_ITMCODE"" = '" & itemCode & "' " &
                    '                 "AND H1.""U_FRMDATE"" <= '" & fromDate & "' " &
                    '                 "AND H1.""U_TODATE"" >= '" & toDate & "' " &
                    '                 "AND H1.""U_FRMTIME"" <= '" & FromTime & "' " &
                    '                 "AND H1.""U_TOTIME"" >= '" & ToTime & "' " &
                    '                 "AND H1.""U_STATUS"" = 'Active'" &
                    '                                     "AND C1.""U_UOM""='" & UOM & "'"

                    '                                 ' "AND C1.""Code""='" & Barcode & "' " &

                    '                                 Dim recordSet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                                 recordSet.DoQuery(query)
                    '                                 Dim doc As String = recordSet.Fields.Item("DocNum").Value
                    '                                 If recordSet.Fields.Item(0).Value > 0 Then
                    '                                     objMain.objApplication.StatusBar.SetText("DocNum " & doc & " - Row " & i.ToString() & ": Item " & itemCode & " is already in another active promotion with overlapping dates.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    '                                     'objMain.objApplication.StatusBar.SetText("Row " & i.ToString() & ": Item " & itemCode & " is already in another active promotion with overlapping dates.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '                                     BubbleEvent = False
                    '                                     'Exit Sub
                    '                                 End If
                    '                             End If
                    '                         Next
                    '                     End If
                    '                 End If


                    If pVal.ItemUID = "Upload" And pVal.BeforeAction = False Then
                        Dim filePath As String = CType(objForm.Items.Item("Path").Specific, SAPbouiCOM.EditText).Value
                        If System.IO.File.Exists(filePath) Then
                            Me.LoadExcelToMatrix(objForm, filePath)
                        Else
                            objMain.objApplication.MessageBox("File not found or path is invalid.")
                        End If
                    End If
                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                    '    If Me.validateDiscountType(objForm.UniqueID) = False Then
                    '        BubbleEvent = False
                    '    End If
                    'End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        'If Me.validateBranch(objForm.UniqueID) = False Then
                        '    BubbleEvent = False
                        'End If


                    End If

                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                    '    If Me.validateDiscountpercentage(objForm.UniqueID) = False Then
                    '        BubbleEvent = False
                    '    End If
                    'End If
                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                    '    If Me.validateDiscountAmount(objForm.UniqueID) = False Then
                    '        BubbleEvent = False
                    '    End If
                    'End If



                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                    '    If Me.ValidatePromotionName(objForm.UniqueID) = False Then
                    '        BubbleEvent = False
                    '    End If
                    'End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.ValidateItem(objForm.UniqueID) = False Then
                            BubbleEvent = False
                        End If
                    End If





                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                                  (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                        Dim projtype As String = objForm.Items.Item("PRTYP").Specific.Value
                        'If Me.Validate(objForm.UniqueID, projtype) = False Then
                        '    LostFocusFlag = False
                        '    BubbleEvent = False
                        'End If

                    End If
                    If pVal.ItemUID = "btnGO" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("SBCDE").Specific.Value <> "" Then
                            Dim bcode As String = objForm.Items.Item("SBCDE").Specific.Value
                            Dim Afterchk As String = " "
                            Dim TYPE As String = objForm.Items.Item("PRTYP").Specific.Value
                            Select Case TYPE
                                Case "All"
                                    For i As Integer = 1 To objMatrix.VisualRowCount
                                        If bcode = objMatrix.Columns.Item("BCODE").Cells.Item(i).Specific.Value Then
                                            objMatrix.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            ' objForm.Items.Item("SBCDE").Specific.Value = Afterchk
                                            Exit For
                                        End If
                                        objMatrix.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                                    Next
                                Case "ItemGroup"
                                    For i As Integer = 1 To objMatrix3.VisualRowCount
                                        If bcode = objMatrix3.Columns.Item("BCODE").Cells.Item(i).Specific.Value Then
                                            objMatrix3.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            ' objForm.Items.Item("SBCDE").Specific.Value = Afterchk
                                            Exit For
                                        End If
                                    Next

                                Case "ItemWise"
                                    For i As Integer = 1 To objMatrix.VisualRowCount
                                        If bcode = objMatrix.Columns.Item("BCODE").Cells.Item(i).Specific.Value Then
                                            objMatrix.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            ' objForm.Items.Item("SBCDE").Specific.Value = Afterchk
                                            Exit For
                                        End If
                                        objMatrix.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                                    Next
                                Case "ItemWise"
                                    For i As Integer = 1 To objMatrix3.VisualRowCount
                                        If bcode = objMatrix3.Columns.Item("BCODE").Cells.Item(i).Specific.Value Then
                                            objMatrix3.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            ' objForm.Items.Item("SBCDE").Specific.Value = Afterchk
                                            Exit For
                                        End If

                                    Next
                                Case "SubGroup"
                                    For i As Integer = 1 To objMatrix3.VisualRowCount
                                        If bcode = objMatrix3.Columns.Item("BCODE").Cells.Item(i).Specific.Value Then
                                            objMatrix3.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            ' objForm.Items.Item("SBCDE").Specific.Value = Afterchk
                                            Exit For
                                        End If
                                    Next

                                Case "SubSubGroup"
                                    For i As Integer = 1 To objMatrix3.VisualRowCount
                                        If bcode = objMatrix3.Columns.Item("BCODE").Cells.Item(i).Specific.Value Then
                                            objMatrix3.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            ' objForm.Items.Item("SBCDE").Specific.Value = Afterchk
                                            Exit For
                                        End If
                                    Next
                            End Select
                        Else
                            objMain.objApplication.StatusBar.SetText("Please Enter Barcode", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                    End If




                    'madhu link button to PCR
                    If pVal.BeforeAction = False AndAlso pVal.ItemUID = "ToPCR" Then
                        Dim DocumentNumber As String = oDS.GetValue("U_PCRDoc", 0).Trim()
                        objMain.objApplication.ActivateMenuItem("PCRT")
                        objNewForm = objMain.objApplication.Forms.GetForm("PCRT", objMain.objApplication.Forms.ActiveForm.TypeCount)
                        objNewForm.Freeze(True)
                        objNewForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objNewForm.Items.Item("DocNum").Specific.Value = DocumentNumber
                        objNewForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        objNewForm.Freeze(False)
                    End If

                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "IMtx" And pVal.ColUID = "OFPRWVAT" And pVal.BeforeAction = True Then
                        Try
                            objForm.Freeze(True)
                            Dim RetPrice As Double = 0

                            Dim ObjMtx1 As SAPbouiCOM.Matrix = objForm.Items.Item(pVal.ItemUID).Specific
                            Dim ODbs_child As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Item("@TNX_ITEMDISC_C0")

                            If CStr(ObjMtx1.Columns.Item("RTPRWVAT").Cells.Item(pVal.Row).Specific.Value).Trim <> "" Then
                                RetPrice = CDbl(CStr(ObjMtx1.Columns.Item("RTPRWVAT").Cells.Item(pVal.Row).Specific.Value).Trim)
                            End If

                            Dim OfferPrice As Double = 0

                            If CStr(ObjMtx1.Columns.Item("OFPRWVAT").Cells.Item(pVal.Row).Specific.Value).Trim <> "" Then
                                OfferPrice = CDbl(CStr(ObjMtx1.Columns.Item("OFPRWVAT").Cells.Item(pVal.Row).Specific.Value).Trim)
                            End If

                            If OfferPrice > RetPrice Then
                                objMain.objApplication.StatusBar.SetText("Offer Price should be less then or equal to retail price")
                                objForm.Freeze(False)
                                BubbleEvent = False
                                Exit Try
                            End If
                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If

                    If pVal.ItemUID = "IMtx" And pVal.ColUID = "OFPRWVAT" And pVal.BeforeAction = False And pVal.ItemChanged = True Then
                        Try
                            objForm.Freeze(True)

                            Dim RetPrice As Double = 0

                            Dim ObjMtx1 As SAPbouiCOM.Matrix = objForm.Items.Item(pVal.ItemUID).Specific
                            Dim ODbs_child As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Item("@TNX_ITEMDISC_C0")

                            If CStr(ObjMtx1.Columns.Item("RTPRWVAT").Cells.Item(pVal.Row).Specific.Value).Trim <> "" Then
                                RetPrice = CDbl(CStr(ObjMtx1.Columns.Item("RTPRWVAT").Cells.Item(pVal.Row).Specific.Value).Trim)
                            End If

                            Dim OfferPrice As Double = 0

                            If CStr(ObjMtx1.Columns.Item("OFPRWVAT").Cells.Item(pVal.Row).Specific.Value).Trim <> "" Then
                                OfferPrice = CDbl(CStr(ObjMtx1.Columns.Item("OFPRWVAT").Cells.Item(pVal.Row).Specific.Value).Trim)
                            End If

                            If OfferPrice > RetPrice Then
                                objMain.objApplication.StatusBar.SetText("Offer Price should be less then or equal to retail price")
                                objForm.Freeze(False)
                                BubbleEvent = False
                                Exit Try
                            Else
                                Dim DiscountAmount As Double = RetPrice - OfferPrice
                                Dim DiscPer As Double = (DiscountAmount / RetPrice) * 100
                                ODbs_child.SetValue("LineId", ODbs_child.Offset, pVal.Row)
                                ODbs_child.SetValue("U_BCODE", ODbs_child.Offset, ObjMtx1.Columns.Item("BCODE").Cells.Item(pVal.Row).Specific.Value)
                                ODbs_child.SetValue("U_ITMCODE", ODbs_child.Offset, ObjMtx1.Columns.Item("ITMCODE").Cells.Item(pVal.Row).Specific.Value)
                                ODbs_child.SetValue("U_ITMNM", ODbs_child.Offset, ObjMtx1.Columns.Item("ITMNM").Cells.Item(pVal.Row).Specific.Value)
                                ODbs_child.SetValue("U_UOM", ODbs_child.Offset, ObjMtx1.Columns.Item("UOM").Cells.Item(pVal.Row).Specific.Value)
                                ODbs_child.SetValue("U_FRMDATE", ODbs_child.Offset, ObjMtx1.Columns.Item("FRMDATE").Cells.Item(pVal.Row).Specific.Value)
                                ODbs_child.SetValue("U_TODATE", ODbs_child.Offset, ObjMtx1.Columns.Item("TODATE").Cells.Item(pVal.Row).Specific.Value)
                                ODbs_child.SetValue("U_ITMCST", ODbs_child.Offset, ObjMtx1.Columns.Item("ITMCST").Cells.Item(pVal.Row).Specific.Value)
                                ODbs_child.SetValue("U_RTPRWVAT", ODbs_child.Offset, ObjMtx1.Columns.Item("RTPRWVAT").Cells.Item(pVal.Row).Specific.Value)
                                ODbs_child.SetValue("U_OFPRWVAT", ODbs_child.Offset, ObjMtx1.Columns.Item("OFPRWVAT").Cells.Item(pVal.Row).Specific.Value)
                                ODbs_child.SetValue("U_DCPRWVAT", ODbs_child.Offset, DiscountAmount)
                                ODbs_child.SetValue("U_DPERCTG", ODbs_child.Offset, DiscPer)
                                ObjMtx1.SetLineData(pVal.Row)
                            End If
                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If
                    'if this commented excel upload is succesful
                    'If pVal.ItemUID = "TODATE" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                    '    If objForm.Items.Item("FRMDATE").Specific.Value > objForm.Items.Item("TODATE").Specific.Value Then
                    '        objMain.objApplication.StatusBar.SetText("To Date must be after date to From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        BubbleEvent = False
                    '    End If
                    'End If
                    'If pVal.ItemUID = "IMtx" And pVal.BeforeAction = True Then
                    '    If (objForm.Items.Item("FRMDATE").Specific.Value = "") Or (objForm.Items.Item("TODATE").Specific.Value = "") Then
                    '        objMain.objApplication.StatusBar.SetText("Please Fill header fields first", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        BubbleEvent = False
                    '    End If
                    'End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    Try
                        ' If (pVal.ItemUID = "IMtx" And (pVal.ColUID = "DPERCTG" Or pVal.ColUID = "OFPRWVAT" And pVal.ColUID = "REMK")) And pVal.BeforeAction = False Then ' And pVal.ActionSuccess = False Then

                        'If (pVal.ItemUID = "IMtx" And (pVal.ColUID = "DPERCTG" Or pVal.ColUID = "OFPRWVAT")) And pVal.BeforeAction = False Then ' And pVal.ActionSuccess = False Then


                        ' If (pVal.ItemUID = "IMtx" And (pVal.ColUID = "DPERCTG" Or pVal.ColUID = "OFPRWVAT")) And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "IMtx" AndAlso (pVal.ColUID = "DPERCTG") AndAlso pVal.BeforeAction = False Then


                            Dim discountper As String = objMatrix.Columns.Item("DPERCTG").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()
                            Dim offerpri As String = objMatrix.Columns.Item("OFPRWVAT").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()
                            If discountper <> cdiscountpercentage Then
                                Dim ObjMtx1 As SAPbouiCOM.Matrix = objForm.Items.Item(pVal.ItemUID).Specific  '  Or pVal.ColUID = "DPERCTG"
                                'Dim discountprice As String = ObjMtx1.Columns.Item("DCPRWVAT").Cells.Item(pVal.Row).Specific.Value
                                Dim retailprice As String = ObjMtx1.Columns.Item("RTPRWVAT").Cells.Item(pVal.Row).Specific.Value
                                Dim discountpercentage As String = ObjMtx1.Columns.Item("DPERCTG").Cells.Item(pVal.Row).Specific.Value

                                ObjMtx1.Columns.Item("DCPRWVAT").Cells.Item(pVal.Row).Specific.Value = ((retailprice) * (discountpercentage)) / 100
                                Dim discountprice As String = ObjMtx1.Columns.Item("DCPRWVAT").Cells.Item(pVal.Row).Specific.Value
                                ObjMtx1.Columns.Item("OFPRWVAT").Cells.Item(pVal.Row).Specific.Value = (retailprice) - discountprice
                                cofferprice = ObjMtx1.Columns.Item("OFPRWVAT").Cells.Item(pVal.Row).Specific.Value
                                cdiscountpercentage = ObjMtx1.Columns.Item("DPERCTG").Cells.Item(pVal.Row).Specific.Value


                                '  ObjMtx1.Columns.Item("REMK").Cells.Item(ObjMtx1.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                            End If
                        End If



                        'commented imp percentage
                        ' If (pVal.ItemUID = "IMtx" And (pVal.ColUID = "OFPRWVAT" Or pVal.ColUID = "DPERCTG" And pVal.ColUID = "REMK")) And pVal.BeforeAction = False Then ' And pVal.ActionSuccess = False Then
                        'If (pVal.ItemUID = "IMtx" And (pVal.ColUID = "OFPRWVAT" Or pVal.ColUID = "DPERCTG")) And pVal.BeforeAction = False Then

                        '    Dim discountper As String = objMatrix.Columns.Item("DPERCTG").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()
                        '    Dim offerpri As String = objMatrix.Columns.Item("OFPRWVAT").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()
                        '    If offerpri <> cofferprice Then
                        '        Dim ObjMtx1 As SAPbouiCOM.Matrix = objForm.Items.Item(pVal.ItemUID).Specific  '  Or pVal.ColUID = "DPERCTG"
                        '        Dim retailprice As String = ObjMtx1.Columns.Item("RTPRWVAT").Cells.Item(pVal.Row).Specific.Value
                        '        Dim discountpercentage As String = ObjMtx1.Columns.Item("DPERCTG").Cells.Item(pVal.Row).Specific.Value
                        '        ObjMtx1.Columns.Item("DCPRWVAT").Cells.Item(pVal.Row).Specific.Value = (retailprice) - offerpri
                        '        Dim discountprice As String = ObjMtx1.Columns.Item("DCPRWVAT").Cells.Item(pVal.Row).Specific.Value
                        '        ObjMtx1.Columns.Item("OFPRWVAT").Cells.Item(pVal.Row).Specific.Value = (retailprice) - discountprice
                        '        cofferprice = ObjMtx1.Columns.Item("OFPRWVAT").Cells.Item(pVal.Row).Specific.Value
                        '        ObjMtx1.Columns.Item("DPERCTG").Cells.Item(pVal.Row).Specific.Value = ((discountprice) / (retailprice)) * 100



                        '        ' ObjMtx1.Columns.Item("REMK").Cells.Item(ObjMtx1.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                        '    End If
                        'End If


                        'discount amount 
                        'Importance----------------

                        'If (pVal.ItemUID = "IMtx" And (pVal.ColUID = "DCPRWVAT" Or pVal.ColUID = "OFPRWVAT" Or pVal.ColUID = "REMK")) And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "IMtx" AndAlso (pVal.ColUID = "DCPRWVAT") AndAlso pVal.BeforeAction = False Then

                            'And pVal.ActionSuccess = False Then
                            Dim ObjMtx1 As SAPbouiCOM.Matrix = objForm.Items.Item(pVal.ItemUID).Specific
                            Dim offerPriceStr As String = ObjMtx1.Columns.Item("OFPRWVAT").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()
                            'Dim existingOfferPrice As Double = 0.0
                            'If Double.TryParse(offerPriceStr, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, existingOfferPrice) Then
                            '    If existingOfferPrice > 0.0 Then
                            '        Exit Sub
                            '    End If
                            'End If
                            Dim retailPrice As Double = Convert.ToDouble(ObjMtx1.Columns.Item("RTPRWVAT").Cells.Item(pVal.Row).Specific.Value)
                            Dim discountAmount As Double = Convert.ToDouble(ObjMtx1.Columns.Item("DCPRWVAT").Cells.Item(pVal.Row).Specific.Value)
                            Dim offerPrice As Double = retailPrice - discountAmount
                            Dim discountPercentage As Double = (discountAmount / retailPrice) * 100
                            ObjMtx1.Columns.Item("OFPRWVAT").Cells.Item(pVal.Row).Specific.Value = offerPrice.ToString("0.00")
                            ObjMtx1.Columns.Item("DPERCTG").Cells.Item(pVal.Row).Specific.Value = discountPercentage.ToString("0.00")
                            cofferprice = offerPrice.ToString("0.00")
                            cdiscountpercentage = discountPercentage.ToString("0.00")
                            ObjMtx1.Columns.Item("REMK").Cells.Item(ObjMtx1.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        End If



                        'discount end amount 


                        'Copy option

                        If pVal.ItemUID = "IMtx" And pVal.ColUID = "BCODE" And pVal.BeforeAction = False Then

                            Try
                                objForm.Freeze(True)
                                Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("IMtx").Specific
                                If pVal.ItemUID = "IMtx" Then
                                    If pVal.Row = objMatrix.VisualRowCount Then
                                        Me.SetNewLine(objForm.UniqueID, "IMtx")
                                    End If
                                End If
                                objForm.Freeze(False)
                            Catch ex As Exception
                                objForm.Freeze(False)
                            End Try

                        End If





                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try






                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDS = objForm.DataSources.DBDataSources.Item("@TNX_ITEMDISC")
                    oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMDISC_C0")
                    objMatrix = objForm.Items.Item("IMtx").Specific
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects





                    'imp vvvv
                    If (oCFL.UniqueID = "CFL_2" Or oCFL.UniqueID = "CFL_ITM") And pVal.BeforeAction = True Then
                        ''Me.CFLFilter("CFL_OCRD")
                        Dim StrCFLQry As String = ""
                        'here changed

                        If objForm.Items.Item("ITMGRP").Specific.Value <> "" Then
                            StrCFLQry = "SELECT I.""ItemCode"",  B.""BcdCode"" ,I.""U_VSubGroup"",I.""U_VSubSubGroup""
                    FROM ""OBCD"" B " &
                 "INNER JOIN ""OITM"" I ON B.""ItemCode"" = I.""ItemCode"" " &
                 "WHERE   B.""ItemCode"" IS NOT NULL "


                            If CStr(objForm.Items.Item("ITMGRP").Specific.Value).Trim <> "" Then
                                StrCFLQry = StrCFLQry & "  and  I.""ItmsGrpCod"" = '" & objForm.Items.Item("ITMGRP").Specific.Value & "'"
                            End If

                            If CStr(objForm.Items.Item("SubGrp").Specific.Value).Trim <> "" Then
                                StrCFLQry = StrCFLQry & "  And  I.""U_VSubGroup"" = '" & objForm.Items.Item("SubGrp").Specific.Value & "'"
                            End If

                            If CStr(objForm.Items.Item("SubSubGrp").Specific.Value).Trim <> "" Then
                                StrCFLQry = StrCFLQry & "  And  I.""U_VSubSubGroup"" = '" & objForm.Items.Item("SubSubGrp").Specific.Value & "'"
                            End If


                        End If
                        'here changed

                        If oCFL.UniqueID = "CFL_ITM" Then
                            CFL_Filter1(pVal, objMain.objApplication, objMain.objCompany, StrCFLQry, "BcdCode", "U_VSubSubGroup", "U_VSubSubGroup", "ItemCode")
                        End If

                        If oCFL.UniqueID = "CFL_2" Then
                            CFL_Filter(pVal, objMain.objApplication, objMain.objCompany, StrCFLQry, "BcdCode", "U_VSubSubGroup", "U_VSubSubGroup", "ItemCode")
                        End If

                    End If

                    'If (oCFL.UniqueID = "CFL_2" Or oCFL.UniqueID = "CFL_ITM") And pVal.BeforeAction = True Then

                    '    Dim StrCFLQry As String = ""

                    '    ' Check if PRTYP = "All"
                    '    If objForm.Items.Item("PRTYP").Specific.Value = "All" Then

                    '        ' If ITMGRP is selected
                    '        If objForm.Items.Item("ITMGRP").Specific.Value <> "" Then

                    '            StrCFLQry = "SELECT I.""ItemCode"", B.""BcdCode"", I.""U_VSubGroup"", I.""U_VSubSubGroup"" " &
                    '    "FROM ""OBCD"" B " &
                    '    "INNER JOIN ""OITM"" I ON B.""ItemCode"" = I.""ItemCode"" " &
                    '    "WHERE B.""ItemCode"" IS NOT NULL "

                    '            ' Apply filters
                    '            If CStr(objForm.Items.Item("ITMGRP").Specific.Value).Trim <> "" Then
                    '                StrCFLQry &= " AND I.""ItmsGrpCod"" = '" & objForm.Items.Item("ITMGRP").Specific.Value & "'"
                    '            End If

                    '            If CStr(objForm.Items.Item("SubGrp").Specific.Value).Trim <> "" Then
                    '                StrCFLQry &= " AND I.""U_VSubGroup"" = '" & objForm.Items.Item("SubGrp").Specific.Value & "'"
                    '            End If

                    '            If CStr(objForm.Items.Item("SubSubGrp").Specific.Value).Trim <> "" Then
                    '                StrCFLQry &= " AND I.""U_VSubSubGroup"" = '" & objForm.Items.Item("SubSubGrp").Specific.Value & "'"
                    '            End If

                    '        End If

                    '    Else
                    '        ' ELSE condition when PRTYP is not "All"
                    '        StrCFLQry = "SELECT I.""ItemCode"", I.""ItemName"" FROM ""OITM"" I WHERE I.""ItemCode"" IS NOT NULL  "

                    '    End If

                    '    If oCFL.UniqueID = "CFL_ITM" Then
                    '        CFL_Filter1(pVal, objMain.objApplication, objMain.objCompany, StrCFLQry, "BcdCode", "U_VSubSubGroup", "U_VSubSubGroup", "ItemCode")
                    '    End If

                    '    If oCFL.UniqueID = "CFL_2" Then
                    '        CFL_Filter(pVal, objMain.objApplication, objMain.objCompany, StrCFLQry, "BcdCode", "U_VSubSubGroup", "U_VSubSubGroup", "ItemCode")
                    '    End If

                    'End If

                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If oDS.GetValue("U_FRMDATE", 0) = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Choose From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If oDS.GetValue("U_TODATE", 0) = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Choose To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If



                        If oCFL.UniqueID = "CFL_ITM" And pVal.BeforeAction = False Then

                            If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then

                                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                End If

                                oDS1.SetValue("LineId", oDS1.Offset, pVal.Row)

                                Dim Brch As String = objForm.Items.Item("BRANCH").Specific.Value
                                Dim ItemCode As String = oDT.GetValue("ItemCode", 0)

                                ' Get Barcode
                                Dim GetBarCode As String = "SELECT TOP 1 ""BcdCode"" FROM OBCD WHERE ""ItemCode"" = '" & ItemCode & "' ORDER BY ""BcdEntry"" ASC"
                                Dim oRsGetBarCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetBarCode.DoQuery(GetBarCode)

                                Dim Barcode As String = ""
                                If oRsGetBarCode.RecordCount > 0 Then
                                    Barcode = oRsGetBarCode.Fields.Item("BcdCode").Value.ToString()
                                End If

                                ' Set Item details to data source
                                oDS1.SetValue("U_BCODE", oDS1.Offset, Barcode)
                                oDS1.SetValue("U_ITMCODE", oDS1.Offset, ItemCode)
                                oDS1.SetValue("U_ITMNM", oDS1.Offset, oDT.GetValue("ItemName", 0))
                                ' oDS1.SetValue("U_UOM", oDS1.Offset, oDT.GetValue("InvntryUom", 0))


                                'ADDON-

                                Dim Itemcodenew As String = oDT.GetValue("ItemCode", 0)
                                Dim Itemuom As String = "SELECT DISTINCT A.""UomEntry"", " &
"(SELECT ""UomCode"" FROM ""OUOM"" WHERE ""UomEntry"" = A.""UomEntry"") AS ""UomName"" " &
"FROM ""OBCD"" A " &
"WHERE A.""ItemCode"" = '" & Itemcodenew & "'"

                                Dim itmcode As SAPbobsCOM.Recordset
                                itmcode = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                itmcode.DoQuery(Itemuom)
                                If itmcode.RecordCount > 0 Then
                                    Dim itm As String = itmcode.Fields.Item("UomName").Value.ToString()
                                    oDS1.SetValue("U_UOM", oDS1.Offset, itm)
                                End If






                                ' Copy dates from header DS to line DS
                                oDS1.SetValue("U_FRMDATE", oDS1.Offset, oDS.GetValue("U_FRMDATE", 0))
                                oDS1.SetValue("U_TODATE", oDS1.Offset, oDS.GetValue("U_TODATE", 0))

                                ' Update matrix row
                                objMatrix.SetLineData(pVal.Row)

                                ' === Get Item Cost ===
                                Try
                                    Dim Itemcod As String = objMatrix.Columns.Item("ITMCODE").Cells.Item(pVal.Row).Specific.Value
                                    If Not String.IsNullOrWhiteSpace(Itemcod) Then
                                        Dim query1 As String = "CALL ""GetItemPrice""('" & Itemcod & "','" & Brch.Trim & "')"
                                        Dim orset1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        orset1.DoQuery(query1)

                                        If orset1.RecordCount > 0 Then
                                            Dim Itemcost As Double = CDbl(orset1.Fields.Item("AvgPrice").Value)
                                            oDS1.SetValue("U_ITMCST", oDS1.Offset, Itemcost)
                                        Else
                                            oDS1.SetValue("U_ITMCST", oDS1.Offset, 0)
                                        End If
                                    End If
                                Catch ex As Exception
                                    'oDS1.SetValue("U_ITMCST", oDS1.Offset, 0)
                                    ' Optional: log error if needed
                                End Try

                                ' === Get Retail Price ===
                                Try
                                    If Not String.IsNullOrWhiteSpace(Barcode) Then
                                        Dim query2 As String = "CALL ""GetRetailPrice""('" & Barcode & "','" & Brch.Trim & "')"
                                        Dim orset2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        orset2.DoQuery(query2)

                                        If orset2.RecordCount > 0 Then
                                            Dim RetailPrice As Double = CDbl(orset2.Fields.Item("Price").Value)
                                            oDS1.SetValue("U_RTPRWVAT", oDS1.Offset, RetailPrice)
                                        Else
                                            oDS1.SetValue("U_RTPRWVAT", oDS1.Offset, 0)
                                        End If
                                    End If
                                Catch ex As Exception
                                    '  oDS1.SetValue("U_RTPRWVAT", oDS1.Offset, 0)
                                End Try

                                ' Set default pricing/discount fields
                                oDS1.SetValue("U_OFPRWVAT", oDS1.Offset, 0)
                                oDS1.SetValue("U_DCPRWVAT", oDS1.Offset, 0)
                                oDS1.SetValue("U_DPERCTG", oDS1.Offset, 0)

                                ' Final update of the matrix row
                                objMatrix.SetLineData(pVal.Row)

                                ' Auto-add new line if current is last row
                                If pVal.Row = objMatrix.VisualRowCount Then
                                    Me.SetNewLine(objForm.UniqueID, "IMtx")
                                End If
                            End If
                        End If


                        If oCFL.UniqueID = "CFL_Ex" Then
                            oDS4.SetValue("LineID", oDS4.Offset, pVal.Row)

                            oDS4.SetValue("U_BCODE", oDS4.Offset, oDT.GetValue("CodeBars", 0))
                            oDS4.SetValue("U_ITMCODE", oDS4.Offset, oDT.GetValue("ItemCode", 0))
                            oDS4.SetValue("U_ITMNM", oDS4.Offset, oDT.GetValue("ItemName", 0))
                            oDS4.SetValue("U_UOM", oDS4.Offset, oDT.GetValue("InvntryUom", 0))
                            objMatrix3.SetLineData(pVal.Row)
                            If objMatrix3.VisualRowCount = pVal.Row Then
                                Me.SetNewLine(objForm.UniqueID, "EMtx")
                            End If
                        End If
                        If oCFL.UniqueID = "CFL_3" And pVal.BeforeAction = False Then
                            Dim bcode As String = oDT.GetValue("BcdCode", 0)
                            Dim FromDate As String = oDS.GetValue("U_FRMDATE", 0)
                            Dim ToDate As String = oDS.GetValue("U_TODATE", 0)
                            Dim Itemcode As String = ""

                            oDS4.SetValue("U_BCODE", oDS4.Offset, oDT.GetValue("BcdCode", 0))
                            Dim itmqury As String = "SELECT T1.""ItemCode"", T1.""ItemName"", T0.""BcdCode"", T2.""UomCode"" " &
                                                    "FROM ""OBCD"" T0 INNER JOIN ""OITM"" T1 ON T0.""ItemCode"" = T1.""ItemCode"" INNER JOIN OUOM T2 ON " &
                                                        " T0.""UomEntry"" = T2.""UomEntry"" Where T0.""BcdCode"" = '" & bcode & "'"
                            Dim oRitmqury As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRitmqury.DoQuery(itmqury)
                            oDS4.SetValue("LineID", oDS4.Offset, pVal.Row)

                            oDS4.SetValue("U_ITMCODE", oDS4.Offset, oRitmqury.Fields.Item("ItemCode").Value)
                            oDS4.SetValue("U_ITMNM", oDS4.Offset, oRitmqury.Fields.Item("ItemName").Value)
                            oDS4.SetValue("U_UOM", oDS4.Offset, oRitmqury.Fields.Item("UomCode").Value)
                            oDS4.SetValue("U_BCODE", oDS4.Offset, oRitmqury.Fields.Item("BcdCode").Value)
                            'oDS1.SetValue("U_FRMDATE", oDS1.Offset, oDS.GetValue("U_FRMDATE", 0))
                            'oDS1.SetValue("U_TODATE", oDS1.Offset, oDS.GetValue("U_TODATE", 0))

                            objMatrix3.SetLineData(pVal.Row)
                            If pVal.Row = objMatrix3.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "EMtx")
                            End If
                            'If objMatrix3.VisualRowCount = pVal.Row Then
                            '    Me.SetNewLine(objForm.UniqueID, "EMtx")
                            'End If
                        End If
                    End If

                    If oCFL.UniqueID = "CFL_2" And pVal.BeforeAction = False Then
                        If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            Dim bcode As String = oDT.GetValue("BcdCode", 0)
                            Dim FromDate As String = oDS.GetValue("U_FRMDATE", 0)
                            Dim ToDate As String = oDS.GetValue("U_TODATE", 0)
                            Dim Itemcode As String = ""


                            ' Dim Itemcode As String = ""
                            oDS1.SetValue("U_BCODE", oDS1.Offset, oDT.GetValue("BcdCode", 0))
                            Dim itmqury As String = "SELECT T1.""ItemCode"", T1.""ItemName"", T0.""BcdCode"", T2.""UomCode"" " &
                                                    "FROM ""OBCD"" T0 INNER JOIN ""OITM"" T1 ON T0.""ItemCode"" = T1.""ItemCode"" INNER JOIN OUOM T2 ON " &
                                                     " T0.""UomEntry"" = T2.""UomEntry"" Where T0.""BcdCode"" = '" & bcode & "'"
                            Dim oRitmqury As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRitmqury.DoQuery(itmqury)

                            oDS1.SetValue("LineId", oDS1.Offset, pVal.Row)
                            oDS1.SetValue("U_ITMCODE", oDS1.Offset, oRitmqury.Fields.Item("ItemCode").Value)
                            oDS1.SetValue("U_ITMNM", oDS1.Offset, oRitmqury.Fields.Item("ItemName").Value)
                            oDS1.SetValue("U_UOM", oDS1.Offset, oRitmqury.Fields.Item("UomCode").Value)
                            oDS1.SetValue("U_BCODE", oDS1.Offset, oRitmqury.Fields.Item("BcdCode").Value)
                            oDS1.SetValue("U_FRMDATE", oDS1.Offset, oDS.GetValue("U_FRMDATE", 0))
                            oDS1.SetValue("U_TODATE", oDS1.Offset, oDS.GetValue("U_TODATE", 0))


                            objMatrix.SetLineData(pVal.Row)

                            Dim query As String = "Select ""U_CNWhse"" from OADM"
                            Dim orset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            orset.DoQuery(query)

                            Dim Warehouse As String = orset.Fields.Item("U_CNWhse").Value

                            Itemcode = oRitmqury.Fields.Item("ItemCode").Value

                            Dim Brch As String = objForm.Items.Item("BRANCH").Specific.Value


                            Try
                                Dim query1 As String = "CALL ""GetItemPrice""('" & Itemcode & "','" & Brch.Trim & "')"

                                Dim orset1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                orset1.DoQuery(query1)

                                If orset1.RecordCount > 0 Then
                                    Dim Itemcost As Double = orset1.Fields.Item("AvgPrice").Value
                                    oDS1.SetValue("U_ITMCST", oDS1.Offset, Itemcost)
                                Else
                                    oDS1.SetValue("U_ITMCST", oDS1.Offset, 0)
                                End If
                            Catch ex As Exception
                            End Try
                            Try
                                Dim query2 As String = "CALL ""GetRetailPrice""('" & bcode & "','" & Brch.Trim & "')"
                                Dim orset2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                orset2.DoQuery(query2)
                                If orset2.RecordCount > 0 Then
                                    Dim Retailprice As Double = orset2.Fields.Item("Price").Value
                                    oDS1.SetValue("U_RTPRWVAT", oDS1.Offset, Retailprice)
                                Else
                                    oDS1.SetValue("U_RTPRWVAT", oDS1.Offset, 0)
                                End If
                            Catch ex As Exception
                                'objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                            End Try

                            oDS1.SetValue("U_OFPRWVAT", oDS1.Offset, 0)
                            oDS1.SetValue("U_DCPRWVAT", oDS1.Offset, 0)
                            oDS1.SetValue("U_DPERCTG", oDS1.Offset, 0)
                            objMatrix.SetLineData(pVal.Row)
                            If objMatrix.VisualRowCount = pVal.Row Then
                                Me.SetNewLine(objForm.UniqueID, "IMtx")
                            End If
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    objForm = objMain.objApplication.Forms.Item(FormUID)


                    'madhu while manual inactive promotion then related PCR also need to be stTatus change to closed
                    Try

                        Dim docNum As String = ""
                        Dim status As String = ""

                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_ITEMDISC")

                        docNum = oDBs_Head.GetValue("DocNum", 0).Trim()
                        status = oDBs_Head.GetValue("U_STATUS", 0).Trim()

                        If status <> "Inactive" Then Exit Try
                        If String.IsNullOrWhiteSpace(docNum) Then Exit Try
                        Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim qry As String = "SELECT ""U_PCRDoc"" FROM ""@TNX_ITEMDISC"" WHERE ""DocNum"" = '" & docNum & "'"
                        oRS.DoQuery(qry)

                        If Not oRS.EoF Then
                            Dim pcrDocNum As String = oRS.Fields.Item("U_PCRDoc").Value.ToString()
                            If Not String.IsNullOrWhiteSpace(pcrDocNum) Then
                                Dim updatePCR As String = "UPDATE ""@TNX_PCRT"" SET ""U_TNXSTUS"" = 'C' WHERE ""DocNum"" = '" & pcrDocNum & "'"
                                Dim oRSUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRSUpdate.DoQuery(updatePCR)

                                objMain.objApplication.StatusBar.SetText("PCR Document " & pcrDocNum & " Closed Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            End If
                        End If

                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText("Error: " & ex.Message)
                    End Try


                    'madhu Please make the status read-only when the line-wise “Active" Status is set to “NO”. Only the approver user should be able to change the status back to “YES” if required.

                    Dim oMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("IMtx").Specific
                    Dim rowCount As Integer = oMatrix.RowCount
                    Dim loginUser As String = objMain.objApplication.Company.UserName.Trim().ToUpper()

                    Dim DocNumPromotion As String = ""
                    Try
                        DocNumPromotion = oDBs_Head.GetValue("DocNum", 0).Trim()
                    Catch ex As Exception
                        DocNumPromotion = ""
                    End Try

                    Dim PromotionType As String = ""
                    Try
                        PromotionType = oDBs_Head.GetValue("U_PRTYP", 0).Trim()
                    Catch ex As Exception
                        PromotionType = ""
                    End Try

                    Dim isApprover As Boolean = False
                    Dim isOriginator As Boolean = False

                    Try

                        Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        Dim sSQL As String = "SELECT DISTINCT T1.""U_Name"" " & "FROM ""@SBO_APPHDR"" T0 " & "INNER JOIN ""@SBO_APPREQ"" T1 ON T0.""Code"" = T1.""Code"" " &
                                              "WHERE T0.""U_Active"" = 'Y' " & "AND UPPER(T0.""U_PCR"") = '" & PromotionType.Replace("'", "''").ToUpper() & "' " &
                                             "AND UPPER(T1.""U_Name"") = '" & loginUser.Replace("'", "''").ToUpper() & "'"

                        oRS.DoQuery(sSQL)

                        Dim oRS1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        Dim sSQL1 As String = "SELECT DISTINCT  T4.""U_AUTH"" FROM ""@SBO_APPHDR"" T0 " &
    "INNER JOIN ""@SBO_APPAUT"" T2 ON T0.""Code"" = T2.""Code"" INNER JOIN ""@SBO_AST"" T3 ON T3.""Code"" = T2.""U_M3_1"" " &
    "INNER JOIN ""@SBO_AST_C0"" T4 ON T4.""Code"" = T3.""Code"" WHERE T0.""U_Active"" = 'Y' AND UPPER(T0.""U_PCR"") = '" & PromotionType.Replace("'", "''").ToUpper() & "' " &
    "AND UPPER(T4.""U_AUTH"") = '" & loginUser.Replace("'", "''").ToUpper() & "'"
                        oRS1.DoQuery(sSQL1)

                        If oRS.RecordCount > 0 Then
                            isOriginator = True
                        End If
                        If oRS1.RecordCount > 0 Then
                            isApprover = True
                        End If
                    Catch ex As Exception
                        isApprover = False
                        isOriginator = False
                    End Try

                    For i As Integer = 1 To rowCount

                        Dim activeValue As String = ""

                        Try

                            Dim oCell As Object = oMatrix.Columns.Item("Active").Cells.Item(i).Specific

                            If TypeOf oCell Is SAPbouiCOM.ComboBox Then
                                activeValue = CType(oCell, SAPbouiCOM.ComboBox).Value.Trim().ToUpper()
                            ElseIf TypeOf oCell Is SAPbouiCOM.EditText Then
                                activeValue = CType(oCell, SAPbouiCOM.EditText).Value.Trim().ToUpper()
                            End If

                        Catch ex As Exception
                            activeValue = ""
                        End Try

                        For j As Integer = 1 To oMatrix.Columns.Count

                            Try

                                If DocNumPromotion <> "" Then

                                    If activeValue = "NO" Then

                                        'Only Authoriser can edit

                                        If isApprover Then

                                            oMatrix.CommonSetting.SetCellEditable(i, j, True)

                                        Else

                                            oMatrix.CommonSetting.SetCellEditable(i, j, False)

                                        End If

                                    Else

                                        oMatrix.CommonSetting.SetCellEditable(i, j, True)

                                    End If

                                Else

                                    'New document
                                    oMatrix.CommonSetting.SetCellEditable(i, j, True)

                                End If

                            Catch ex As Exception
                            End Try

                        Next

                    Next

                    ' changing the status of a particular item in the active column it must be update once the procees is done else as it is return	
                    ' 	After changing sttaes to No and missed to update then it should default change back to yes 

                    If pVal.ColUID = "Active" AndAlso pVal.Row > 0 Then

                        Dim rowNo As Integer = pVal.Row

                        Dim oCombo As SAPbouiCOM.ComboBox = CType(oMatrix.Columns.Item("Active").Cells.Item(rowNo).Specific, SAPbouiCOM.ComboBox)

                        If oCombo.Value.Trim().ToUpper() = "NO" Then

                            Dim formUID1 As String = objForm.UniqueID

                            Dim t As New System.Threading.Thread(
        Sub()

            Threading.Thread.Sleep(10000)

            Try

                Dim oFrm As SAPbouiCOM.Form =
                    objMain.objApplication.Forms.Item(formUID1)

                Dim oMat As SAPbouiCOM.Matrix =
                    CType(oFrm.Items.Item("IMtx").Specific, SAPbouiCOM.Matrix)

                Dim oCombo1 As SAPbouiCOM.ComboBox =
                    CType(oMat.Columns.Item("Active").Cells.Item(rowNo).Specific, SAPbouiCOM.ComboBox)
                If oFrm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    If oCombo1.Value.Trim().ToUpper() = "NO" Then

                        oCombo1.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)

                        For col As Integer = 1 To oMat.Columns.Count

                            Try
                                oMat.CommonSetting.SetCellEditable(rowNo, col, True)
                            Catch ex As Exception
                            End Try

                        Next

                        objMain.objApplication.StatusBar.SetText("Active automatically changed back to Yes after 10 seconds.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                    End If

                End If


            Catch ex As Exception
            End Try

        End Sub)

                            t.IsBackground = True
                            t.Start()

                        End If

                    End If
                    '---Madhu new1


                    'new series add-------
                    If pVal.ItemUID = "25" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("25").Specific.Selected.Value <> "" Then
                            Dim series As Integer = objForm.Items.Item("25").Specific.Selected.Value
                            SetDefault(objForm.UniqueID, "CobmoSelect", series)
                        End If
                    End If
                    'new end series



                    'If pVal.BeforeAction = False AndAlso pVal.ItemUID = "PRTYP" Then
                    '    Dim objComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("PRTYP").Specific

                    '    If objComboBox.Value = "ALL" Then
                    '        objForm.Items.Item("ITMGRP").Enabled = False
                    '        objForm.Items.Item("SubGrp").Enabled = False
                    '        objForm.Items.Item("SubSubGrp").Enabled = False
                    '    End If
                    'End If


                    If pVal.BeforeAction = False AndAlso pVal.ItemUID = "PRTYP" Then
                        Dim objCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("PRTYP").Specific
                        Dim promoType As String = objCombo.Value.Trim()
                        ' Get combo fields
                        Dim cmbItemGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("ITMGRP").Specific
                        Dim cmbSubGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("SubGrp").Specific
                        Dim cmbSubSubGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("SubSubGrp").Specific

                        If promoType = "All" Then
                            objForm.Items.Item("ITMGRP").Enabled = False
                            objForm.Items.Item("SubGrp").Enabled = False
                            objForm.Items.Item("SubSubGrp").Enabled = False
                            cmbItemGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            cmbSubGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            cmbSubSubGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        Else
                            objForm.Items.Item("ITMGRP").Enabled = True
                            objForm.Items.Item("SubGrp").Enabled = True
                            objForm.Items.Item("SubSubGrp").Enabled = True
                            cmbItemGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            cmbSubGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            cmbSubSubGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        End If
                    End If


                    If pVal.ItemUID = "PRTYP" And pVal.BeforeAction = False Then
                        Select Case objForm.Items.Item("PRTYP").Specific.Selected.Value
                            Case "All"
                                objForm.Items.Item("ITMGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("ITMGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objComboBox = objForm.Items.Item("ITMGRP").Specific
                                ' objComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                objForm.Items.Item("SubGrp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("SubGrp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objComboBox = objForm.Items.Item("SubGrp").Specific
                                ' objComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                objForm.Items.Item("SubSubGrp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("SubSubGrp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objComboBox = objForm.Items.Item("SubSubGrp").Specific
                                'objComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

                            Case "ItemWise"
                                objForm.Items.Item("ITMGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("ITMGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objComboBox = objForm.Items.Item("ITMGRP").Specific
                                ' objComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                'NEW ADD
                                objForm.Items.Item("SubGrp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("SubSubGrp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                                'objForm.Items.Item("SubGrp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("SubGrp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objComboBox = objForm.Items.Item("SubGrp").Specific
                                ' objComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

                            'Case "ItemGroup"
                            '    objForm.Items.Item("ITMGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            '    objForm.Items.Item("SubGrp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            '    objForm.Items.Item("SubGrp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            '    objForm.Items.Item("SubSubGrp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            '    objForm.Items.Item("SubSubGrp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            Case "SubGroup"
                                objForm.Items.Item("ITMGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("SubGrp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("SubSubGrp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("SubSubGrp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            Case "SubSubGroup"
                                objForm.Items.Item("ITMGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("SubGrp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("SubSubGrp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        End Select
                    End If




                    If objForm.Items.Item("ITMGRP").Specific.Value <> "" Then
                        Dim objComboBoxItemGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("ITMGRP").Specific
                        If objComboBoxItemGroup.Selected IsNot Nothing Then
                            Dim itemGroupCode As String = objComboBoxItemGroup.Selected.Value

                            Dim sql As String = "SELECT DISTINCT T0.""U_VSubGroup"", T0.""U_VSubGroup"" " &
                            "FROM OITM T0 " &
                            "WHERE T0.""ItmsGrpCod"" = '" & itemGroupCode & "' " &
                            "AND T0.""U_VSubGroup"" IS NOT NULL " &
                            "ORDER BY T0.""U_VSubGroup"""

                            Dim objComboBoxSubGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("SubGrp").Specific

                            objMain.objUtilities.ComboBoxLoadValues(objComboBoxSubGroup, sql)
                            objComboBoxSubGroup.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                            objForm.Items.Item("SubGrp").DisplayDesc = True
                        End If
                    End If





                    If objForm.Items.Item("SubGrp").Specific.Value <> "" Then
                        Dim objComboBoxSubGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("SubGrp").Specific

                        If objComboBoxSubGroup.Selected IsNot Nothing Then
                            Dim Subgroup As String = objComboBoxSubGroup.Selected.Value

                            ' Ensure ItemGroup is defined somewhere earlier, or retrieve it here again
                            Dim objComboBoxItemGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("ITMGRP").Specific
                            Dim ItemGroup As String = ""
                            If objComboBoxItemGroup.Selected IsNot Nothing Then
                                ItemGroup = objComboBoxItemGroup.Selected.Value
                            Else
                                Exit Sub ' or handle as needed
                            End If


                            Dim sql As String = "Select DISTINCT
    ""U_VSubSubGroup"", 
    ""U_VSubSubGroup"" AS ""Description"" 
FROM OITM 
WHERE ""ItmsGrpCod"" = ' " & ItemGroup & "'
  And ""U_VSubSubGroup"" Is Not NULL 
  And ""U_VSubGroup"" = '" & Subgroup & "'  "

                            Dim objComboBoxSubSubGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("SubSubGrp").Specific

                            objMain.objUtilities.ComboBoxLoadValues(objComboBoxSubSubGroup, sql)
                            objComboBoxSubSubGroup.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                            objForm.Items.Item("SubSubGrp").DisplayDesc = True
                        End If
                    End If



                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or
   objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or
   objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                        Dim oComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("DISCTYP").Specific
                        Dim objMatrix1 As SAPbouiCOM.Matrix = CType(objForm.Items.Item("IMtx").Specific, SAPbouiCOM.Matrix)

                        If Not oComboBox.Selected Is Nothing Then
                            Dim selectedValue As String = oComboBox.Selected.Value.Trim()

                            If selectedValue = "A" Then
                                ' ✅ Discount by Amount
                                objMatrix1.Columns.Item("DCPRWVAT").Editable = True
                                objMatrix1.Columns.Item("DPERCTG").Editable = False

                            ElseIf selectedValue = "P" Then
                                ' ✅ Discount by Percentage
                                objMatrix1.Columns.Item("DPERCTG").Editable = True
                                objMatrix1.Columns.Item("DCPRWVAT").Editable = False

                            Else
                                ' ✅ None selected — disable both
                                objMatrix1.Columns.Item("DPERCTG").Editable = False
                                objMatrix1.Columns.Item("DCPRWVAT").Editable = False
                            End If
                        Else
                            ' ✅ Combo not selected — disable both
                            objMatrix1.Columns.Item("DPERCTG").Editable = False
                            objMatrix1.Columns.Item("DCPRWVAT").Editable = False
                        End If
                    End If


                    If pVal.ItemUID = "PRTYP" And pVal.BeforeAction = False Then
                        Dim type As String = objForm.Items.Item("PRTYP").Specific.Value
                        If type = "ItemWise" Then
                            objForm.Items.Item("Pro").Enabled = True
                            objForm.Items.Item("Item_0").Enabled = True
                            objForm.Items.Item("Item_1").Enabled = True
                            objForm.Items.Item("Item_3").Enabled = True 'PREV NO

                            objForm.Items.Item("BMtx").Enabled = True
                            objForm.Items.Item("Pro").Visible = True
                            objForm.Items.Item("Item_0").Visible = False
                            objForm.Items.Item("Item_1").Visible = True
                            objForm.Items.Item("Item_3").Enabled = True 'PREV NO

                            objForm.Items.Item("BMtx").Visible = True
                            objForm.Items.Item("Pro").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        ElseIf type = "ItemGroup" Then
                            objForm.Items.Item("Pro").Enabled = True
                            objForm.Items.Item("Item_0").Enabled = True
                            objForm.Items.Item("Item_1").Enabled = True
                            objForm.Items.Item("Item_3").Enabled = True 'PREV NO
                            objForm.Items.Item("BMtx").Enabled = True
                            objForm.Items.Item("Pro").Visible = False   'prev false
                            objForm.Items.Item("Item_0").Visible = True
                            objForm.Items.Item("Item_0").Left = objForm.Items.Item("Item_0").Left - 100
                            objForm.Items.Item("Item_1").Visible = True
                            objForm.Items.Item("BMtx").Visible = True
                            objForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        ElseIf type = "SubGroup" Then
                            objForm.Items.Item("Pro").Enabled = True
                            objForm.Items.Item("Item_0").Enabled = True
                            objForm.Items.Item("Item_1").Enabled = True
                            objForm.Items.Item("Item_3").Enabled = True 'PREV NO

                            objForm.Items.Item("BMtx").Enabled = True
                            objForm.Items.Item("Pro").Visible = True
                            objForm.Items.Item("Item_0").Visible = True
                            objForm.Items.Item("Item_1").Visible = True
                            objForm.Items.Item("Item_3").Enabled = True 'PREV NO

                            objForm.Items.Item("BMtx").Visible = True
                            objForm.Items.Item("Pro").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                        ElseIf type = "SubSubGroup" Then
                            objForm.Items.Item("Pro").Enabled = True
                            objForm.Items.Item("Item_0").Enabled = True
                            objForm.Items.Item("Item_1").Enabled = True
                            objForm.Items.Item("Item_3").Enabled = True 'PREV NO

                            objForm.Items.Item("BMtx").Enabled = True
                            objForm.Items.Item("Pro").Visible = True
                            objForm.Items.Item("Item_0").Visible = True
                            objForm.Items.Item("Item_1").Visible = True
                            objForm.Items.Item("Item_3").Enabled = True 'PREV NO

                            objForm.Items.Item("BMtx").Visible = True
                            objForm.Items.Item("Pro").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                            Dim grp As String = "SELECT T0.""ItmsGrpCod"", T0.""ItmsGrpNam"" FROM OITB T0"
                            Dim oRgrp As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRgrp.DoQuery(grp)
                            If oRgrp.RecordCount > 0 Then
                                objMatrix1.Clear()
                                For i As Integer = 0 To oRgrp.RecordCount - 1
                                    objMatrix1.AddRow()
                                    oDS2.SetValue("LineId", oDS2.Offset, objMatrix1.VisualRowCount)
                                    oDS2.SetValue("U_ITMCDE", oDS2.Offset, oRgrp.Fields.Item("ItmsGrpCod").Value)
                                    oDS2.SetValue("U_ITMGRP", oDS2.Offset, oRgrp.Fields.Item("ItmsGrpNam").Value)
                                    objMatrix1.SetLineData(objMatrix1.VisualRowCount)
                                    oRgrp.MoveNext()
                                Next

                            End If
                        End If
                    End If
                    If pVal.ItemUID = "STATUS" And pVal.BeforeAction = False Then
                        Dim status As String = objForm.Items.Item("STATUS").Specific.Value
                        If status = "Active" Then
                            objForm.Items.Item("FRMDATE").Enabled = True
                            objForm.Items.Item("TODATE").Enabled = True
                            objForm.Items.Item("SBCDE").Enabled = True
                            objForm.Items.Item("Path").Enabled = True
                            objForm.Items.Item("FRMTIME").Enabled = True
                            objForm.Items.Item("TOTIME").Enabled = True

                            objForm.Items.Item("PRTYP").Enabled = True
                            objForm.Items.Item("PRMNM").Enabled = True
                            objForm.Items.Item("IMtx").Enabled = True
                            objForm.Items.Item("IGMtx1").Enabled = True
                            objForm.Items.Item("EMtx").Enabled = True
                            objForm.Items.Item("BMtx").Enabled = True
                        ElseIf status = "Inactive" Then
                            objForm.Items.Item("FRMDATE").Enabled = False
                            objForm.Items.Item("TODATE").Enabled = False
                            objForm.Items.Item("SBCDE").Enabled = False
                            objForm.Items.Item("Path").Enabled = False
                            objForm.Items.Item("FRMTIME").Enabled = False
                            objForm.Items.Item("TOTIME").Enabled = False

                            'objForm.Items.Item("PRTYP").Enabled = False
                            objForm.Items.Item("PRMNM").Enabled = False
                            objForm.Items.Item("IMtx").Enabled = False
                            objForm.Items.Item("IGMtx1").Enabled = False
                            objForm.Items.Item("EMtx").Enabled = False
                            objForm.Items.Item("BMtx").Enabled = False
                        End If
                    End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("ITEM", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            If objForm.Items.Item("PRTYP").Specific.Value = "ItemWise" Then
                                objForm.Items.Item("Pro").Enabled = True
                                objForm.Items.Item("Item_0").Enabled = True
                                objForm.Items.Item("Item_1").Enabled = True
                                objForm.Items.Item("BMtx").Enabled = True
                                objForm.Items.Item("Pro").Visible = True
                                objForm.Items.Item("Item_0").Visible = True
                                objForm.Items.Item("Item_1").Visible = True
                                objForm.Items.Item("BMtx").Visible = True
                                objForm.Items.Item("Pro").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            ElseIf objForm.Items.Item("PRTYP").Specific.Value = "ItemGroupWise" Then
                                objForm.Items.Item("Pro").Enabled = False
                                objForm.Items.Item("Item_0").Enabled = True
                                objForm.Items.Item("Item_1").Enabled = True
                                objForm.Items.Item("BMtx").Enabled = True
                                objForm.Items.Item("Pro").Visible = False
                                objForm.Items.Item("Item_0").Visible = True
                                objForm.Items.Item("Item_0").Left = objForm.Items.Item("Item_0").Left - 100
                                objForm.Items.Item("Item_1").Visible = True
                                objForm.Items.Item("BMtx").Visible = True
                                objForm.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            End If
                        End If

                        Dim status As String = ""
                        Try
                            status = objForm.Items.Item("STATUS").Specific.Value
                        Catch
                        End Try

                        status = status.Trim()
                        If status = "Inactive" Then

                            Try
                                If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                                End If
                            Catch
                            End Try

                            ' 👉 Disable all fields
                            objForm.Items.Item("FRMDATE").Enabled = False
                            objForm.Items.Item("TODATE").Enabled = False
                            objForm.Items.Item("SBCDE").Enabled = False
                            objForm.Items.Item("Path").Enabled = False
                            objForm.Items.Item("FRMTIME").Enabled = False
                            objForm.Items.Item("TOTIME").Enabled = False
                            objForm.Items.Item("PRTYP").Enabled = False
                            objForm.Items.Item("PRMNM").Enabled = False
                            objForm.Items.Item("IMtx").Enabled = False
                            objForm.Items.Item("IGMtx1").Enabled = False
                            objForm.Items.Item("EMtx").Enabled = False
                            objForm.Items.Item("BMtx").Enabled = False

                        ElseIf status = "Active" Then

                            ' 👉 Keep NORMAL MODE (OK / UPDATE)
                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                            End If

                            ' 👉 Enable all fields
                            objForm.Items.Item("FRMDATE").Enabled = True
                            objForm.Items.Item("TODATE").Enabled = True
                            objForm.Items.Item("SBCDE").Enabled = True
                            objForm.Items.Item("Path").Enabled = True
                            objForm.Items.Item("FRMTIME").Enabled = True
                            objForm.Items.Item("TOTIME").Enabled = True
                            objForm.Items.Item("PRTYP").Enabled = False
                            objForm.Items.Item("PRMNM").Enabled = True
                            objForm.Items.Item("IMtx").Enabled = True
                            objForm.Items.Item("IGMtx1").Enabled = True
                            objForm.Items.Item("EMtx").Enabled = True
                            objForm.Items.Item("BMtx").Enabled = True

                        End If

                        'Maddh Users should have access to line livel inactive, but for access for Active only approver should have access to do both

                        Dim oMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("IMtx").Specific
                        Dim rowCount As Integer = oMatrix.RowCount

                        Dim loginUser As String = objMain.objApplication.Company.UserName.Trim().ToUpper()

                        Dim DocNumPromotion As String = ""
                        Dim PromotionType As String = ""

                        Try

                            Dim oDB As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Item("@TNX_ITEMDISC")

                            Try
                                DocNumPromotion = oDB.GetValue("DocNum", 0).Trim()
                            Catch ex As Exception
                                DocNumPromotion = ""
                            End Try

                            Try
                                PromotionType = oDB.GetValue("U_PRTYP", 0).Trim()
                            Catch ex As Exception
                                PromotionType = ""
                            End Try

                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try

                        '================ APPROVER CHECK =================
                        Dim isApprover As Boolean = False

                        Try

                            Dim oRS As SAPbobsCOM.Recordset =
                                objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            Dim sSQL As String = "SELECT DISTINCT T1.""U_Name"" " & "FROM ""@SBO_APPHDR"" T0 " & "INNER JOIN ""@SBO_APPREQ"" T1 ON T0.""Code"" = T1.""Code"" " &
                                                 "WHERE T0.""U_Active"" = 'Y' " & "AND UPPER(T0.""U_PCR"") = '" & PromotionType.Replace("'", "''").ToUpper() & "' " &
                                                 "AND UPPER(T1.""U_Name"") = '" & loginUser.Replace("'", "''").ToUpper() & "'"

                            oRS.DoQuery(sSQL)

                            If oRS.RecordCount > 0 Then
                                isApprover = True
                                objMain.objApplication.StatusBar.SetText("The logged-in user '" & loginUser & "' is configured as an approver. Editing is allowed.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Else
                                isApprover = False
                                objMain.objApplication.StatusBar.SetText("The logged-in user '" & loginUser & "' is not configured as an approver.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If

                        Catch ex As Exception
                            isApprover = False
                        End Try

                        '================ MATRIX ENABLE / DISABLE =================
                        For i As Integer = 1 To rowCount

                            Dim activeValue As String = ""

                            Try

                                Dim oCell As Object = oMatrix.Columns.Item("Active").Cells.Item(i).Specific

                                If TypeOf oCell Is SAPbouiCOM.ComboBox Then
                                    activeValue = CType(oCell, SAPbouiCOM.ComboBox).Value.Trim().ToUpper()
                                ElseIf TypeOf oCell Is SAPbouiCOM.EditText Then
                                    activeValue = CType(oCell, SAPbouiCOM.EditText).Value.Trim().ToUpper()
                                End If

                            Catch ex As Exception
                                activeValue = ""
                            End Try

                            For j As Integer = 1 To oMatrix.Columns.Count

                                Try

                                    If activeValue = "NO" OrElse activeValue = "N" Then

                                        If isApprover Then
                                            oMatrix.CommonSetting.SetCellEditable(i, j, True)
                                        Else
                                            oMatrix.CommonSetting.SetCellEditable(i, j, False)
                                        End If

                                    Else

                                        oMatrix.CommonSetting.SetCellEditable(i, j, True)

                                    End If

                                Catch ex As Exception
                                End Try

                            Next

                        Next
                        'Madu


                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    If Not BusinessObjectInfo.BeforeAction AndAlso BusinessObjectInfo.ActionSuccess Then

                        Dim oForm As SAPbouiCOM.Form
                        oForm = objMain.objApplication.Forms.Item(BusinessObjectInfo.FormUID)

                        System.Threading.Tasks.Task.Delay(200).ContinueWith(
            Sub()
                SetFormByStatus(oForm)
            End Sub)
                    End If


            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub SetFormByStatus(ByVal objForm As SAPbouiCOM.Form)

        Try
            objForm.Freeze(True)

            Dim status As String = ""

            Try
                status = objForm.DataSources.DBDataSources.Item("@TNX_ITEMDISC").GetValue("U_STATUS", 0).Trim()
            Catch
            End Try

            If status = "Inactive" Then

                ' Force view mode
                If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                End If

                For i As Integer = 0 To objForm.Items.Count - 1

                    Dim item As SAPbouiCOM.Item = objForm.Items.Item(i)

                    If item.UniqueID = "2" Then
                        item.Enabled = True
                    Else
                        item.Enabled = False
                    End If

                Next

            ElseIf status = "Active" Then

                For i As Integer = 0 To objForm.Items.Count - 1
                    objForm.Items.Item(i).Enabled = True
                Next

            End If

        Catch ex As Exception

        Finally
            objForm.Freeze(False)
        End Try

    End Sub


    Public Function GetFieldValue(ByVal query As String) As Object
        Try
            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery(query)
            Dim str As String = oRs.Fields.Item(0).Value.ToString()
            Return str
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Public Function GetNextDocNum()
        Dim Str As String
        Dim oRs As SAPbobsCOM.Recordset
        Dim DocNum As Integer
        oRs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Str = "SELECT IFNULL(MAX(""DocNum""),0)+1 as ""DocNum"" FROM ""@TNX_ITEMDISC"""
        'objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        'objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        Try
            oRs.DoQuery(Str)
            oRs.MoveFirst()
            If oRs.RecordCount > 0 Then
                DocNum = oRs.Fields.Item(0).Value
                oDS.SetValue("DocNum", 0, oRs.Fields.Item(0).Value)
            End If
            If DocNum = 0 Then DocNum = 1
            Return DocNum
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    'Sub SetDefault(ByVal FormUID As String)
    Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "", Optional ByVal Series As Integer = 0)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)



            oDS = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMDISC")
            oDS1 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMDISC_C0")
            oDS2 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMDISC_C1")
            oDS3 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMDISC_C2")
            oDS4 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMDISC_C3")
            objMatrix = objForm.Items.Item("IMtx").Specific
            objMatrix1 = objForm.Items.Item("IGMtx1").Specific
            objMatrix2 = objForm.Items.Item("BMtx").Specific
            objMatrix3 = objForm.Items.Item("EMtx").Specific

            objMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            objForm.Items.Item("DocNum").Enabled = False

            'doc series comment-----------
            'oDS.SetValue("DocNum", oDS.Offset, objMain.objUtilities.GetNextDocNum(objForm, "ME_ITEMUDO", "Primary"))
            ''oDS.SetValue("DocNum", oDS.Offset, objMain.objUtilities.GetNextDocNum(objForm, "ME_ITEMUDO"))

            'objForm.DataBrowser.BrowseBy = "DocNum"
            '----MAdhu changed
            oDS.SetValue("U_STATUS", oDS.Offset, "Active")
            'oDS.SetValue("U_DOCDATE", oDS.Offset, DateTime.Now.ToString("yyyyMMdd")
            Dim cmbSeries As SAPbouiCOM.ComboBox = objForm.Items.Item("25").Specific
            Dim seriesQuery As String = "SELECT ""Series"", ""SeriesName"" FROM NNM1 WHERE ""ObjectCode"" = 'ME_ITEMUDO' AND ""Locked"" = 'N'"
            objMain.objUtilities.ComboBoxLoadValuesSeries(cmbSeries, seriesQuery)
            objForm.Items.Item("25").DisplayDesc = True
            cmbSeries.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            ' Get default series
            '            Dim defaultSeriesQuery As String = "
            'SELECT T1.""Series"", T1.""SeriesName""
            'FROM ONNM T0 
            'JOIN NNM1 T1 ON T1.""ObjectCode"" = T0.""ObjectCode"" AND T1.""Series"" = T0.""DfltSeries"" 
            'WHERE T0.""ObjectCode"" IN ('ME_ITEMUDO','@TNX_ITEMDISC')"


            Dim defaultSeriesQuery As String =
    "SELECT ""Series"", ""SeriesName"", ""ObjectCode"", ""Indicator"" " &
    "FROM NNM1 " &
    "WHERE ""ObjectCode"" IN ('ME_ITEMUDO', '@TNX_ITEMDISC') " &
    "AND ""Indicator"" <> 'Default'"




            Dim oRsDflt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsDflt.DoQuery(defaultSeriesQuery)

            If Flag = "" Then
                Dim nextDocNum = objMain.objUtilities.GetNextDocNumSeries(objForm, "ME_ITEMUDO", oRsDflt.Fields.Item("Series").Value)
                oDS.SetValue("DocNum", oDS.Offset, nextDocNum)
                oDS.SetValue("U_ITEMSER", oDS.Offset, oRsDflt.Fields.Item("Series").Value)
            Else
                Dim nextDocNum = objMain.objUtilities.GetNextDocNumSeries(objForm, "ME_ITEMUDO", Series)
                oDS.SetValue("DocNum", oDS.Offset, nextDocNum)
                oDS.SetValue("U_ITEMSER", oDS.Offset, Series)
            End If

            oDS.SetValue("U_DOCDATE", oDS.Offset, DateTime.Now.ToString("yyyyMMdd"))

            Try
                Dim currentUser As String = ""
                Try
                    currentUser = If(String.IsNullOrWhiteSpace(objMain.objCompany.UserName), objMain.objCompany.UserSignature.ToString(), objMain.objCompany.UserName)
                Catch
                    Try
                        currentUser = objMain.objCompany.UserSignature.ToString()
                    Catch
                        currentUser = ""
                    End Try
                End Try

                If Not String.IsNullOrWhiteSpace(currentUser) Then
                    Try
                        oDS.SetValue("U_Login", oDS.Offset, currentUser)
                    Catch
                        ' field might not exist in this DB schema - ignore
                    End Try
                End If
            Catch ex As Exception
                ' ignore failures setting user
            End Try

            'new end series
            Me.BranchLoad(objForm.UniqueID)
            oDS.SetValue("U_PRTYP", oDS.Offset, "")
            If objForm.Items.Item("PRTYP").Specific.Value = "ItemWise" Then
                objForm.Items.Item("Pro").Enabled = True
                objForm.Items.Item("Item_0").Enabled = True
                objForm.Items.Item("Item_1").Enabled = True
                objForm.Items.Item("BMtx").Enabled = True
                objForm.Items.Item("Pro").Visible = True
                objForm.Items.Item("Item_0").Visible = False
                objForm.Items.Item("Item_1").Visible = False
                objForm.Items.Item("BMtx").Visible = True
                objForm.Items.Item("Pro").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            End If
            objMatrix2 = objForm.Items.Item("BMtx").Specific
            oDS3 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMDISC_C2")
            Dim branch As String = "SELECT T0.""BPLId"", T0.""BPLName"" FROM OBPL T0"
            Dim oRgrp As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRgrp.DoQuery(branch)
            objMatrix2.Clear()
            objMatrix2.FlushToDataSource()

            If oRgrp.RecordCount > 0 Then
                For i As Integer = 1 To oRgrp.RecordCount - 1
                    objMatrix2.AddRow()
                    oDS3.SetValue("LineId", oDS3.Offset, i)

                    'oDS3.SetValue("U_BPLId", oDS3.Offset, oRgrp.Fields.Item("BPLId").Value)
                    oDS3.SetValue("U_BCHNME", oDS3.Offset, oRgrp.Fields.Item("BPLName").Value)
                    objMatrix2.SetLineData(i)

                    oRgrp.MoveNext()
                Next
            End If
            objMatrix.Clear()
            oDS1.Clear()
            objMatrix.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "IMtx")

            objMatrix.AutoResizeColumns()
            objMatrix1.Clear()
            oDS2.Clear()
            objMatrix1.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "IGMtx1")
            objMatrix1.AutoResizeColumns()

            objMatrix2.Clear()
            oDS3.Clear()
            objMatrix2.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "BMtx")
            objMatrix2.AutoResizeColumns()

            objMatrix3.Clear()
            oDS4.Clear()
            objMatrix3.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "EMtx")
            objMatrix3.AutoResizeColumns()

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub BranchLoad(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix2 = objForm.Items.Item("BMtx").Specific
            oDS3 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMDISC_C2")
            Dim branch As String = "SELECT T0.""BPLId"", T0.""BPLName"" FROM OBPL T0"
            Dim oRgrp As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRgrp.DoQuery(branch)
            objMatrix2.Clear()
            objMatrix2.FlushToDataSource()
            If oRgrp.RecordCount > 0 Then
                For i As Integer = 1 To oRgrp.RecordCount
                    objMatrix2.AddRow()
                    oDS3.SetValue("LineId", oDS3.Offset, i)
                    oDS3.SetValue("U_BCODE", oDS3.Offset, oRgrp.Fields.Item("BPLId").Value)
                    oDS3.SetValue("U_BCHNME", oDS3.Offset, oRgrp.Fields.Item("BPLName").Value)
                    objMatrix2.SetLineData(i)
                    oRgrp.MoveNext()
                Next
            End If
        Catch ex As Exception
        End Try
    End Sub
    Sub SetNewLine(ByVal FormUID As String, ByVal MatrixUID As String)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oDS = objForm.DataSources.DBDataSources.Item("@TNX_ITEMDISC")
        oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMDISC_C0")
        oDS2 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMDISC_C1")
        oDS3 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMDISC_C2")
        oDS4 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMDISC_C3")
        objMatrix = objForm.Items.Item("IMtx").Specific
        objMatrix1 = objForm.Items.Item("IGMtx1").Specific
        objMatrix2 = objForm.Items.Item("BMtx").Specific
        objMatrix3 = objForm.Items.Item("EMtx").Specific
        Select Case MatrixUID
            Case "IMtx"
                objMatrix.AddRow()
                oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
                oDS1.SetValue("U_BCODE", oDS1.Offset, "")
                oDS1.SetValue("U_ITMCODE", oDS1.Offset, "")
                oDS1.SetValue("U_Active", oDS1.Offset, "Yes")

                oDS1.SetValue("U_ITMNM", oDS1.Offset, "")
                oDS1.SetValue("U_UOM", oDS1.Offset, "")
                oDS1.SetValue("U_FRMDATE", oDS1.Offset, "")
                oDS1.SetValue("U_TODATE", oDS1.Offset, "")
                oDS1.SetValue("U_ITMCST", oDS1.Offset, "")
                'oDS1.SetValue("U_RTLPRICE", oDS1.Offset, "")
                'oDS1.SetValue("U_OFFRPRIC", oDS1.Offset, "")
                'oDS1.SetValue("U_DISCPRICE", oDS1.Offset, "")
                'oDS1.SetValue("U_MINAMTPURC", oDS1.Offset, "")
                oDS1.SetValue("U_RTPRWVAT", oDS1.Offset, "")
                oDS1.SetValue("U_OFPRWVAT", oDS1.Offset, "")
                oDS1.SetValue("U_DCPRWVAT", oDS1.Offset, "")
                oDS1.SetValue("U_DPERCTG", oDS1.Offset, "")
                objMatrix.SetLineData(objMatrix.VisualRowCount)
                objMatrix.AutoResizeColumns()
            Case "IGMtx1"
                objMatrix1.AddRow()
                oDS2.SetValue("LineId", oDS2.Offset, objMatrix1.VisualRowCount)
                oDS2.SetValue("U_CHK", oDS2.Offset, "")
                oDS2.SetValue("U_ITMCDE", oDS2.Offset, "")
                oDS2.SetValue("U_ITMGRP", oDS2.Offset, "")
                oDS2.SetValue("U_DISAMNT", oDS2.Offset, "")
                oDS2.SetValue("U_DPERC", oDS2.Offset, "")
                objMatrix1.SetLineData(objMatrix1.VisualRowCount)
                objMatrix1.AutoResizeColumns()
            Case "EMtx"
                objMatrix3.AddRow()
                oDS4.SetValue("LineId", oDS4.Offset, objMatrix3.VisualRowCount)
                oDS4.SetValue("U_BCODE", oDS4.Offset, "")
                'oDS4.SetValue("U_CHK", oDS4.Offset, "")
                oDS4.SetValue("U_ITMCODE", oDS4.Offset, "")
                oDS4.SetValue("U_ITMNM", oDS4.Offset, "")
                oDS4.SetValue("U_UOM", oDS4.Offset, "")
                objMatrix3.AutoResizeColumns()
                objMatrix3.SetLineData(objMatrix3.VisualRowCount)
            Case "BMtx"
                objMatrix2.AddRow()
                oDS3.SetValue("LineId", oDS3.Offset, objMatrix2.VisualRowCount)
                oDS3.SetValue("U_CHK", oDS3.Offset, "")
                oDS3.SetValue("U_BCODE", oDS3.Offset, "")
                oDS3.SetValue("U_BCHNME", oDS3.Offset, "")
                objMatrix2.SetLineData(objMatrix2.VisualRowCount)
                objMatrix2.AutoResizeColumns()
        End Select
    End Sub
    Public Function Validate(ByVal FormUID As String, ByVal MatrixUID As String) As Boolean
        Try
            For ij As Integer = 1 To objMatrix.VisualRowCount
                Dim Barcode1 As String = objMatrix.Columns.Item("BCODE").Cells.Item(ij).Specific.Value
                For ik As Integer = 1 To objMatrix.VisualRowCount
                    If ij <> ik Then
                        If Barcode1 = objMatrix.Columns.Item("BCODE").Cells.Item(ik).Specific.Value Then
                            objMain.objApplication.StatusBar.SetText("Duplicate Barcode " & Barcode1 & " Exists in row number " & ik & "...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                Next
            Next




            Dim FromDate1 As String = oDS.GetValue("U_FRMDATE", 0)
            Dim ToDate1 As String = oDS.GetValue("U_TODATE", 0)
            Dim Barcode11 As String = ""
            For ij As Integer = 1 To objMatrix.VisualRowCount

                Barcode11 = objMatrix.Columns.Item("BCODE").Cells.Item(ij).Specific.Value
                Dim Qry As String = "select distinct T0.""U_BCODE"",T1.""U_FRMDATE"", T1.""U_TODATE"", T1.""DocNum"" " &
                                             "from ""@TNX_ITEMDISC_C0"" T0 inner join ""@TNX_ITEMDISC"" T1 on T0.""DocEntry""=T1.""DocEntry"" " &
                                          " where  (T1.""U_FRMDATE"" between '" & FromDate1 & "' and '" & ToDate1 & "' " &
                                         " and T0.""U_BCODE""<> ''   And T1.""U_STATUS""='Active' and T1.""DocNum"" <> '" & objForm.Items.Item("DocNum").Specific.Value & "' ) or " &
                                         "(T1.""U_TODATE"" between '" & FromDate1 & "' and '" & ToDate1 & "' and T0.""U_BCODE""<> '' And T1.""U_STATUS""='Active' and 
                                         T1.""DocNum"" <> '" & objForm.Items.Item("DocNum").Specific.Value & "' )    and IFNULL(T0.""U_Active"",'Yes') = 'Yes'"
                Dim rs2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim chkBarCode1 As String
                Dim chkDocNum1 As String
                Dim chkFromDate1 As String
                Dim chkToDate1 As String
                rs2.DoQuery(Qry)
                If rs2.RecordCount > 0 Then
                    For i As Integer = 1 To rs2.RecordCount
                        chkBarCode1 = rs2.Fields.Item("U_BCODE").Value
                        chkDocNum1 = rs2.Fields.Item("DocNum").Value
                        chkFromDate1 = Format(CType(rs2.Fields.Item("U_FRMDATE").Value, Date), "yyyyMMdd")
                        chkToDate1 = Format(CType(rs2.Fields.Item("U_TODATE").Value, Date), "yyyyMMdd")
                        If chkBarCode1 = Barcode11 And ((chkFromDate1 >= FromDate1 Or chkFromDate1 <= ToDate1) Or (chkToDate1 >= FromDate1 Or chkToDate1 <= ToDate1)) Then
                            objMain.objApplication.StatusBar.SetText("This Barcode " & chkBarCode1 & " Already Exist in Document Number  " & chkDocNum1 & "  between " & rs2.Fields.Item("U_FRMDATE").Value & "  to  " & rs2.Fields.Item("U_TODATE").Value & "  ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                        rs2.MoveNext()

                    Next
                End If
                ' Dim Barcode1 As String = objMatrix.Columns.Item("BCODE").Cells.Item(ik).Specific.Value

            Next



            'If oDS.GetValue("U_FRMDATE", 0) = "" Then
            '    objMain.objApplication.StatusBar.SetText("Please Choose From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            '    Exit Function
            'End If
            'If oDS.GetValue("U_TODATE", 0) = "" Then
            '    objMain.objApplication.StatusBar.SetText("Please Choose To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            '    Exit Function
            'End If
            ''If oDS3.GetValue("U_CHK", 0) = "" Then
            '    objMain.objApplication.StatusBar.SetText("Please Select Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            '    Exit Function
            'End If
            If oDS.GetValue("U_BPLID", 0) = "" Then
                objMain.objApplication.StatusBar.SetText("Please Select Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                Exit Function
            End If

            'If oDS1.GetValue("U_DPERCTG", 0) = "" Then
            '    objMain.objApplication.StatusBar.SetText("Please Select Discount Percentage", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            '    Exit Function
            'End If






            'If oDS.GetValue("U_FRMTIME", 0) = "" Then
            '    objMain.objApplication.StatusBar.SetText("Please Enter From Time", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            '    Exit Function
            'End If
            'If oDS.GetValue("U_TOTIME", 0) = "" Then
            '    objMain.objApplication.StatusBar.SetText("Please Enter To Time", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            '    Exit Function
            'End If
            'If oDS.GetValue("U_PRMNM", 0) = "" Then
            '    objMain.objApplication.StatusBar.SetText("Please Enter Promotion Name", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)



            '    Return False
            '    Exit Function
            'End If

            objMatrix = objForm.Items.Item("IMtx").Specific
            objMatrix1 = objForm.Items.Item("IGMtx1").Specific


            Select Case MatrixUID


                Case "ItemGroup"
                    Dim type As String = objForm.Items.Item("PRTYP").Specific.Value
                    If type = "ItemGroup" Then

                        For i As Integer = 1 To objMatrix.VisualRowCount
                            If objMatrix.Columns.Item("BCODE").Cells.Item(1).Specific.Value.ToString().Trim() = "" Then
                                'objMain.objApplication.StatusBar.SetText("Please Choose Different selection And press Load Button", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objMain.objApplication.StatusBar.SetText("Please Choose BarCode", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If


                        Next
                    End If


                    For ij As Integer = 1 To objMatrix.VisualRowCount

                        Dim Barcode1 As String = objMatrix.Columns.Item("BCODE").Cells.Item(ij).Specific.Value
                        If (objMatrix.Columns.Item("BCODE").Cells.Item(ij).Specific.Value = "") And (objMatrix.Columns.Item("OFPRWVAT").Cells.Item(ij).Specific.Value <= 0) Then
                            Exit For
                        ElseIf (objMatrix.Columns.Item("BCODE").Cells.Item(ij).Specific.Value = Barcode1) And (objMatrix.Columns.Item("OFPRWVAT").Cells.Item(ij).Specific.Value <= 0) Then
                            objMain.objApplication.StatusBar.SetText("Please Enter Offer Price with VAT greater than Zero in row number " & ij & "...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                        For ik As Integer = 1 To objMatrix.VisualRowCount
                            If ij <> ik Then
                                If Barcode1 = objMatrix.Columns.Item("BCODE").Cells.Item(ik).Specific.Value Then
                                    objMain.objApplication.StatusBar.SetText("Duplicate Barcode " & Barcode1 & " Exists in row number " & ik & "...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Return False
                                End If

                            End If
                        Next

                    Next
                    Dim FromDate As String = oDS.GetValue("U_FRMDATE", 0)
                    Dim ToDate As String = oDS.GetValue("U_TODATE", 0)
                    Dim Barcode As String = ""
                    For ij As Integer = 1 To objMatrix.VisualRowCount

                        Barcode = objMatrix.Columns.Item("BCODE").Cells.Item(ij).Specific.Value
                        Dim Qry As String = "select distinct T0.""Code"",T1.""U_FRMDATE"", T1.""U_TODATE"", T1.""DocNum"" " &
                                             "from ""@TNX_ITEMDISC_C0"" T0 inner join ""@TNX_ITEMDISC"" T1 on T0.""DocEntry""=T1.""DocEntry"" " &
                                          " where  (T1.""U_FRMDATE"" between '" & FromDate & "' and '" & ToDate & "' " &
                                         " and T0.""Code""<> ''   And T1.""U_STATUS""='Active' and T1.""DocNum"" <> '" & objForm.Items.Item("DocNum").Specific.Value & "' ) or " &
                                         "(T1.""U_TODATE"" between '" & FromDate & "' and '" & ToDate & "' and T0.""Code""<> '' And T1.""U_STATUS""='Active' and T1.""DocNum"" <> '" & objForm.Items.Item("DocNum").Specific.Value & "' )"
                        Dim rs2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim chkBarCode As String
                        Dim chkDocNum As String
                        Dim chkFromDate As String
                        Dim chkToDate As String
                        rs2.DoQuery(Qry)
                        If rs2.RecordCount > 0 Then
                            For i As Integer = 1 To rs2.RecordCount
                                chkBarCode = rs2.Fields.Item("Code").Value
                                chkDocNum = rs2.Fields.Item("DocNum").Value
                                chkFromDate = Format(CType(rs2.Fields.Item("U_FRMDATE").Value, Date), "yyyyMMdd")
                                chkToDate = Format(CType(rs2.Fields.Item("U_TODATE").Value, Date), "yyyyMMdd")
                                If chkBarCode = Barcode And ((chkFromDate >= FromDate Or chkFromDate <= ToDate) Or (chkToDate >= FromDate Or chkToDate <= ToDate)) Then
                                    objMain.objApplication.StatusBar.SetText("This Barcode " & chkBarCode & " Already Exist in Document Number  " & chkDocNum & "  between " & rs2.Fields.Item("U_FRMDATE").Value & "  to  " & rs2.Fields.Item("U_TODATE").Value & "  ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Return False
                                End If
                                rs2.MoveNext()

                            Next
                        End If
                        ' Dim Barcode1 As String = objMatrix.Columns.Item("BCODE").Cells.Item(ik).Specific.Value

                    Next

                    For j As Integer = 1 To objMatrix2.VisualRowCount
                        Dim oCheckBox As SAPbouiCOM.CheckBox
                        oCheckBox = objMatrix2.Columns.Item("CHK").Cells.Item(j).Specific
                        For k As Integer = 1 To objMatrix2.VisualRowCount
                            oCheckBox = objMatrix2.Columns.Item("CHK").Cells.Item(k).Specific
                            If oCheckBox.Checked = True Then
                                Exit For
                            End If
                        Next
                        If oCheckBox.Checked = False Then
                            objMain.objApplication.StatusBar.SetText("Please Select Atleast one Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    Next


                Case "ItemWise"

                    Dim type As String = objForm.Items.Item("PRTYP").Specific.Value
                    If type = "ItemWise" Then

                        objMatrix1 = objForm.Items.Item("IGMtx1").Specific
                        Dim oCheckBox As SAPbouiCOM.CheckBox
                        Dim flag As Integer = 0
                        For k As Integer = 1 To objMatrix1.VisualRowCount
                            oCheckBox = objMatrix1.Columns.Item("CHK").Cells.Item(k).Specific
                            If (oCheckBox.Checked = True) Or (objMatrix1.Columns.Item("DPERC").Cells.Item(k).Specific.Value > 0) Then
                                If oCheckBox.Checked = True Then
                                    If objMatrix1.Columns.Item("DPERC").Cells.Item(k).Specific.Value <= 0 Then
                                        objMain.objApplication.StatusBar.SetText("Please Enter Discount% greater than zero in Line No " & k, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Return False
                                    End If
                                ElseIf objMatrix1.Columns.Item("DPERC").Cells.Item(k).Specific.Value > 0 Then
                                    If oCheckBox.Checked = False Then
                                        objMain.objApplication.StatusBar.SetText("Please Check Item Group in Line No " & k, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Return False
                                    End If
                                End If
                            Else
                                flag = flag + 1
                                'If flag = objMatrix1.VisualRowCount Then
                                '    objMain.objApplication.StatusBar.SetText("Please Select Atleast one Item Group", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                '    Return False
                                'End If
                            End If
                        Next

                        'For i As Integer = 1 To objMatrix3.VisualRowCount
                        '    If objMatrix3.Columns.Item("BCODE").Cells.Item(1).Specific.Value.ToString().Trim() = "" Then
                        '        'objMain.objApplication.StatusBar.SetText("Please Choose Different selection And press Load Button", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '        objMain.objApplication.StatusBar.SetText("Please Choose BarCode in Exclusion Tab ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '        Return False
                        '    End If

                        'Next
                    End If
                    'Dim ICheckBox As SAPbouiCOM.CheckBox
                    'For ij As Integer = 1 To objMatrix3.VisualRowCount

                    '    Dim Barcode1 As String = objMatrix3.Columns.Item("BCODE").Cells.Item(ij).Specific.Value
                    '    For ik As Integer = 1 To objMatrix3.VisualRowCount
                    '        If ij <> ik Then
                    '            If Barcode1 = objMatrix3.Columns.Item("BCODE").Cells.Item(ik).Specific.Value Then
                    '                objMain.objApplication.StatusBar.SetText("Duplicate Barcode " & Barcode1 & " Exists in row number " & ik & "...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '                Return False
                    '            End If
                    '        End If
                    '    Next

                    'Next

                    'Dim FromDate As String = oDS.GetValue("U_FRMDATE", 0)
                    'Dim ToDate As String = oDS.GetValue("U_TODATE", 0)
                    'Dim Barcode As String = ""
                    'For ij As Integer = 1 To objMatrix3.VisualRowCount


                    '    Barcode = objMatrix3.Columns.Item("BCODE").Cells.Item(ij).Specific.Value

                    '    Dim Qry As String = "select distinct T0.""Code"",T1.""U_FRMDATE"", T1.""U_TODATE"", T1.""DocNum""  from ""@TNX_ITEMDISC_C3"" T0 inner join ""@TNX_ITEMDISC"" T1 on T0.""DocEntry""=T1.""DocEntry"" " &
                    '                                    " where  T1.""U_STATUS""='Active' and  (T1.""U_FRMDATE"" between '" & FromDate & "' and '" & ToDate & "' and  T0.""Code""<> '')  or (T1.""U_TODATE"" between '" & FromDate & "' and '" & ToDate & "' and  T0.""Code""<> '')"
                    '    Dim rs2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '    Dim chkBarCode As String
                    '    Dim chkDocNum As String
                    '    Dim chkFromDate As String
                    '    Dim chkToDate As String

                    '    rs2.DoQuery(Qry)
                    '    If rs2.RecordCount > 0 Then
                    '        For i As Integer = 1 To rs2.RecordCount
                    '            chkBarCode = rs2.Fields.Item("Code").Value
                    '            chkDocNum = rs2.Fields.Item("DocNum").Value
                    '            chkFromDate = Format(CType(rs2.Fields.Item("U_FRMDATE").Value, Date), "yyyyMMdd")
                    '            chkToDate = Format(CType(rs2.Fields.Item("U_TODATE").Value, Date), "yyyyMMdd")
                    '            If chkBarCode = Barcode And ((chkFromDate >= FromDate Or chkFromDate <= ToDate) Or (chkToDate >= FromDate Or chkToDate <= ToDate)) Then
                    '                objMain.objApplication.StatusBar.SetText("This Barcode " & Barcode & " Already Exist in Document Number  " & chkDocNum & "  between " & rs2.Fields.Item("U_FRMDATE").Value & "  to  " & rs2.Fields.Item("U_TODATE").Value & "  ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '                Return False
                    '            End If
                    '            rs2.MoveNext()
                    '        Next
                    '    End If
                    'Next
                    'For j As Integer = 1 To objMatrix2.VisualRowCount
                    '    Dim oCheckBox As SAPbouiCOM.CheckBox
                    '    oCheckBox = objMatrix2.Columns.Item("CHK").Cells.Item(j).Specific
                    '    For k As Integer = 1 To objMatrix2.VisualRowCount
                    '        oCheckBox = objMatrix2.Columns.Item("CHK").Cells.Item(k).Specific
                    '        If oCheckBox.Checked = True Then
                    '            Exit For
                    '        End If
                    '    Next
                    '    If oCheckBox.Checked = False Then
                    '        objMain.objApplication.StatusBar.SetText("Please Select Atleast one Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        Return False
                    '    End If
                    'Next
            End Select

            If objForm.Items.Item("SBCDE").Specific.Value <> "" Then
                Dim bcode As String = objForm.Items.Item("SBCDE").Specific.Value
                Dim Afterchk As String = " "
                objForm.Items.Item("SBCDE").Specific.Value = Afterchk
            End If

            Return True

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Public Sub CreateProductionOrder(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDS = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR")
            oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR_C0")
            objMatrix = objForm.Items.Item("IMtx").Specific
            PrgBar = objMain.objApplication.StatusBar.CreateProgressBar("Posting production orders, Please wait....", objMatrix.RowCount, False)
            PrgBar.Value = 0
            For i As Integer = 1 To objMatrix.RowCount
                PrgBar.Value = i
                If objMatrix.Columns.Item("ICode").Cells.Item(i).Specific.Value.ToString().Trim() <> "" Then
                    'If objMatrix.Columns.Item("Prod").Cells.Item(i).Specific.Value.ToString().Trim() = "" Then
                    Dim WarehouseCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Whs").Specific
                    Dim Whs As String = ""
                    Try
                        Whs = WarehouseCombo.Selected.Value
                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText("Please select warehouse", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                    objMain.objApplication.StatusBar.SetText("Production orders posting, Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Dim Prod As SAPbobsCOM.ProductionOrders = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)

                    Dim PostDate As String = objForm.Items.Item("PostDate").Specific.Value.ToString().Trim()
                    If PostDate = "" Then
                        objMain.objApplication.SetStatusBarMessage("Please enter ther the Posting Date", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        Exit Sub
                    End If
                    PostDate = PostDate.Insert(4, "/")
                    PostDate = PostDate.Insert(7, "/")
                    Dim PDate As Date = CDate(PostDate)
                    Prod.PostingDate = PDate
                    Prod.StartDate = PDate

                    Dim DueDate As String = objForm.Items.Item("PostDate").Specific.Value.ToString().Trim()
                    If DueDate = "" Then
                        objMain.objApplication.SetStatusBarMessage("Please enter ther the Due Date", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        Exit Sub
                    End If
                    DueDate = DueDate.Insert(4, "/")
                    DueDate = DueDate.Insert(7, "/")
                    Dim DDate As Date = CDate(DueDate)
                    Prod.DueDate = DDate

                    Prod.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotSpecial
                    Prod.ItemNo = objMatrix.Columns.Item("RCode").Cells.Item(i).Specific.Value.ToString().Trim()
                    Prod.PlannedQuantity = CDbl(objMatrix.Columns.Item("RQty").Cells.Item(i).Specific.Value.ToString().Trim())
                    Prod.Warehouse = Whs

                    Prod.Lines.ItemNo = objMatrix.Columns.Item("ICode").Cells.Item(i).Specific.Value.ToString().Trim()
                    Prod.Lines.PlannedQuantity = objMatrix.Columns.Item("Qty").Cells.Item(i).Specific.Value.ToString().Trim()
                    Prod.Lines.Warehouse = Whs
                    Prod.Lines.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                    Dim retVal As Integer = Prod.Add()
                    If retVal <> 0 Then
                        objMain.objApplication.StatusBar.SetText("Error:" & objMain.objCompany.GetLastErrorDescription.Replace("'", ""), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Dim oRsUpdate2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim sError As String = objMain.objCompany.GetLastErrorDescription.Replace("'", "").Replace("'", "")
                        Dim Update2 As String = " UPDATE  T1 SET T1.""U_Status""='Production Order Creation Error: " & sError & "' From ""@SBO_CLEAR_C0"" T1 WHERE T1.""DocEntry""='" & oDS.GetValue("DocEntry", 0) & "' and T1.""LineId""='" & i & "'"
                        oRsUpdate2.DoQuery(Update2)
                    Else
                        Dim oRsGetProd As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim Production As String = "SELECT max(T0.""DocEntry"") FROM OWOR T0"
                        oRsGetProd.DoQuery(Production)
                        Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim Update As String = "UPDATE T1 SET T1.""U_Prod""=(SELECT max(T0.""DocEntry"") FROM OWOR T0) , T1.""U_ProdNo""=(SELECT ""DocNum"" From OWOR  Where ""DocEntry"" = (SELECT max(T0.""DocEntry"") FROM OWOR T0)) ,    T1.""U_Status""='Production Order Posted' From ""@SBO_CLEAR_C0"" T1  WHERE T1.""DocEntry""='" & oDS.GetValue("DocEntry", 0) & "' and T1.""LineId""='" & i & "'"
                        oRsUpdate.DoQuery(Update)
                        '--------Releaseing----------------
                        Dim Release As SAPbobsCOM.ProductionOrders = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                        If Release.GetByKey(oRsGetProd.Fields.Item(0).Value) = True Then
                            Release.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposReleased
                            If Release.Update() = 0 Then
                                Dim oRsUpdate2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim Update2 As String = " UPDATE  T1 SET T1.""U_Status""='Production Order Released' From ""@SBO_CLEAR_C0"" T1 WHERE T1.""DocEntry""='" & oDS.GetValue("DocEntry", 0) & "' and T1.""LineId""='" & i & "'"
                                oRsUpdate2.DoQuery(Update2)
                                'Me.CreateReceipt(objForm.UniqueID, i, oDS.GetValue("U_Branch", 0))
                            Else
                                Dim oRsUpdate3 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim Update3 As String = " UPDATE  T1 SET T1.""U_Status""='Production Order Release Error:'" & objMain.objCompany.GetLastErrorDescription.Replace("'", "").ToString().Substring(0, 200) & " From ""@SBO_CLEAR_C0"" T1 WHERE T1.""DocEntry""='" & oDS.GetValue("DocEntry", 0) & "' and T1.""LineId""='" & i & "'"
                                oRsUpdate3.DoQuery(Update3)
                                objMain.objApplication.StatusBar.SetText("Error:" & objMain.objCompany.GetLastErrorDescription.Replace("'", ""), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                        End If
                    End If
                End If
                'End If
            Next
            PrgBar.Stop()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(PrgBar)
            objMain.objApplication.ActivateMenuItem("1304")
            objMatrix.AutoResizeColumns()
            objMain.objApplication.StatusBar.SetText("Production orders posted successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            objForm.Freeze(False)
            '-----------------------------------
        Catch ex As Exception
            PrgBar.Stop()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(PrgBar)

            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub CreateIssueForProduction(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDS = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR")
            oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR_C0")
            objMatrix = objForm.Items.Item("IMtx").Specific

            Dim result As Integer = objMain.objApplication.MessageBox("You are going to post Receipt from production. Do you want to continue?", 1, "Continue", "Cancel")
            If (result = 1) Then
                objMatrix = objForm.Items.Item("IMtx").Specific
                PrgBar = objMain.objApplication.StatusBar.CreateProgressBar("Posting production orders, Please wait....", objMatrix.RowCount, False)
                PrgBar.Value = 0
                For i As Integer = 1 To objMatrix.VisualRowCount
                    PrgBar.Value = i
                    If objMatrix.Columns.Item("ICode").Cells.Item(i).Specific.Value.ToString().Trim() <> "" And objMatrix.Columns.Item("Prod").Cells.Item(i).Specific.Value.ToString().Trim() <> "" And objMatrix.Columns.Item("RProd").Cells.Item(i).Specific.Value.ToString().Trim() = "" Then
                        objMain.objApplication.StatusBar.SetText("Please Wait Issue is posting.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        '----------
                        Dim ProdDetails As String = "Select ""LineNum"" from WOR1 A Where A.""DocEntry""='" & objMatrix.Columns.Item("Prod").Cells.Item(i).Specific.Value.ToString().Trim() & "'"
                        Dim oRsProdDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsProdDetails.DoQuery(ProdDetails)

                        Dim Issue As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
                        Issue.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
                        Issue.JournalMemo = "Issue from Production"
                        objComboBox = objForm.Items.Item("23").Specific
                        Issue.BPL_IDAssignedToInvoice = objComboBox.Value
                        Dim LineCount As Integer = 0
                        For j As Integer = 1 To oRsProdDetails.RecordCount
                            If LineCount > 0 Then
                                Issue.Lines.Add()
                            End If
                            Issue.Lines.BaseEntry = objMatrix.Columns.Item("Prod").Cells.Item(i).Specific.Value.ToString().Trim()
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
                            Dim Update3 As String = " UPDATE  T1 SET T1.""U_Status""='Issue From Production Error: " & sError & "' From ""@SBO_CLEAR_C0"" T1  WHERE T1.""DocEntry""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "' and T1.""LineId""='" & i & "'"
                            oRsUpdate3.DoQuery(Update3)
                            objMain.objApplication.StatusBar.SetText("Error:" & objMain.objCompany.GetLastErrorDescription.Replace("'", ""), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim Update As String = " UPDATE T1 SET T1.""U_IProd""=(SELECT max(T0.""DocEntry"") FROM OIGE T0),T1.""U_IProdNo""=(SELECT ""DocNum"" FROM OIGE Where ""DocEntry"" = (SELECT max(T0.""DocEntry"") FROM OIGE T0)) ,     T1.""U_Status""='Issue From Production Completed' From ""@SBO_CLEAR_C0"" T1 WHERE T1.""DocEntry""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "' and T1.""LineId""='" & i & "'"
                            oRsUpdate.DoQuery(Update)
                        End If
                    End If
                Next
                objMain.objApplication.StatusBar.SetText("Issue From Production posted successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                objMain.objApplication.ActivateMenuItem("1304")
                objMatrix.AutoResizeColumns()
            End If
            'Else
            'objMain.objApplication.StatusBar.SetText("Please check component quantities.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'End If
        Catch ex As Exception
            PrgBar.Stop()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(PrgBar)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Sub CreateReceipt(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDS = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR")
            oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR_C0")
            objMatrix = objForm.Items.Item("IMtx").Specific

            Dim GetSeries As String = " Select ""Series"" , ""SeriesName""   From NNM1 Where ""ObjectCode"" = '59' And ""Locked"" = 'N' And ""BPLId"" =  '" & oDS.GetValue("U_Branch", 0) & "'  Order By ""Series"" Desc "
            Dim oRsGetSeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetSeries.DoQuery(GetSeries)
            Dim DocSeries As String = oRsGetSeries.Fields.Item("Series").Value

            PrgBar = objMain.objApplication.StatusBar.CreateProgressBar("Please Wait Receipt is posting, Please wait....", objMatrix.RowCount, False)
            PrgBar.Value = 0
            For i As Integer = 1 To objMatrix.VisualRowCount
                PrgBar.Value = i
                If objMatrix.Columns.Item("ICode").Cells.Item(i).Specific.Value.ToString().Trim() <> "" And objMatrix.Columns.Item("Prod").Cells.Item(i).Specific.Value.ToString().Trim() <> "" And objMatrix.Columns.Item("RProd").Cells.Item(i).Specific.Value.ToString().Trim() = "" Then
                    Dim Receipt As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
                    Receipt.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
                    Receipt.JournalMemo = "Receipt from Production"
                    Receipt.Series = DocSeries
                    Receipt.BPL_IDAssignedToInvoice = oDS.GetValue("U_Branch", 0)
                    Receipt.Lines.BaseEntry = objMatrix.Columns.Item("Prod").Cells.Item(i).Specific.Value.ToString().Trim()
                    Receipt.Lines.BaseType = 202

                    Receipt.Lines.Add()
                    Dim retVal As Integer = Receipt.Add()
                    '-------------
                    If retVal <> 0 Then
                        Dim oRsUpdate3 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        Dim Update3 As String = " UPDATE  T1 SET T1.""U_Status""='Receipt From Production Error: " & objMain.objCompany.GetLastErrorDescription.Replace("'", "") & "' From ""@SBO_CLEAR_C0"" T1  WHERE T1.""DocEntry""='" & oDS.GetValue("DocEntry", 0) & "' and T1.""LineId""='" & i & "'"
                        oRsUpdate3.DoQuery(Update3)
                        objMain.objApplication.StatusBar.SetText("Error:" & objMain.objCompany.GetLastErrorDescription.Replace("'", ""), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Else
                        Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim Update As String = " UPDATE T1 SET T1.""U_RProd""=(SELECT max(T0.""DocEntry"") FROM OIGN T0), T1.""U_RProdNo""=(SELECT ""DocNum""  FROM OIGN Where ""DocEntry"" = (SELECT max(T0.""DocEntry"") FROM OIGN T0)) , T1.""U_Status""='Receipt From Production Completed' From ""@SBO_CLEAR_C0"" T1 WHERE T1.""DocEntry""='" & oDS.GetValue("DocEntry", 0) & "' and T1.""LineId""='" & i & "'"
                        oRsUpdate.DoQuery(Update)
                    End If
                End If
            Next

            PrgBar.Stop()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(PrgBar)
            objMain.objApplication.StatusBar.SetText("Receipt From Production posted successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objMain.objApplication.ActivateMenuItem("1304")
            objMatrix.AutoResizeColumns()
        Catch ex As Exception
            PrgBar.Stop()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(PrgBar)

            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

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
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                        objForm.EnableMenu("1292", True)
                        'objForm.EnableMenu("774", True)

                        'If eventInfo.ItemUID = "IMtx" And eventInfo.ColUID = "LineId" Then
                        '    Try
                        '        oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                        '        oMenus = oMenuItem.SubMenus
                        '        If oMenus.Exists("Delete Row") = False Then
                        '            oCreationPackage.UniqueID = "Delete Row"
                        '            oCreationPackage.String = "Delete Row"
                        '            oCreationPackage.Enabled = True
                        '            oMenus.AddEx(oCreationPackage)
                        '        End If
                        '    Catch ex As Exception
                        '        objMain.objApplication.StatusBar.SetText(ex.Message)
                        '    End Try

                        'Else
                        '    Try
                        '        oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                        '        oMenus = oMenuItem.SubMenus
                        '        If oMenus.Exists("Delete Row") = True Then
                        '            objMain.objApplication.Menus.RemoveEx("Delete Row")
                        '        End If
                        '        If oMenus.Exists("Close") = True Then
                        '            objMain.objApplication.Menus.RemoveEx("Close")
                        '        End If
                        '    Catch ex As Exception
                        '        objMain.objApplication.StatusBar.SetText(ex.Message)
                        '    End Try
                        'End If
                    End If
                End If
            End If


            ' System.Diagnostics.Process.Start()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region
    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDS = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR")
            oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR_C0")
            objMatrix = objForm.Items.Item("IMtx").Specific
            If objForm.Items.Item("PostDate").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Post Date Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If objMatrix.VisualRowCount = 0 Then
                objMain.objApplication.StatusBar.SetText("Matrix cannot be Empty")
                Return False
            ElseIf objMatrix.VisualRowCount = 1 Then
                If (objMatrix.Columns.Item("ICode").Cells.Item(1).Specific.Value.Trim = "") Then
                    objMain.objApplication.StatusBar.SetText("Matrix cannot be Empty")
                    Return False
                End If
            End If
            For i As Integer = 1 To objMatrix.VisualRowCount
                Dim BarCode As String = objMatrix.Columns.Item("BarCode").Cells.Item(i).Specific.Value.Trim
                Dim ICode As String = objMatrix.Columns.Item("ICode").Cells.Item(i).Specific.Value.Trim
                Dim IssueQty As Double = objMatrix.Columns.Item("Qty").Cells.Item(i).Specific.Value.Trim
                If (BarCode <> "" And ICode = "") Then
                    objMain.objApplication.StatusBar.SetText("Issue Itemcode should not be empty for Line No. " & i & ".... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                If (ICode <> "" And IssueQty = 0) Then
                    objMain.objApplication.StatusBar.SetText("For Item " & ICode & " and Line No. " & i & ".... IssueQty should not be zero!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Next
            For i As Integer = 1 To objMatrix.VisualRowCount
                Dim ICode As String = objMatrix.Columns.Item("ICode").Cells.Item(i).Specific.Value.Trim
                Dim RCode As String = objMatrix.Columns.Item("RCode").Cells.Item(i).Specific.Value.Trim
                Dim RQty As Double = objMatrix.Columns.Item("RQty").Cells.Item(i).Specific.Value.Trim
                If (ICode <> "" And RCode = "") Then
                    objMain.objApplication.StatusBar.SetText("Receipt Itemcode should not be empty for Line No. " & i & ".... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                If (RCode <> "" And RQty = 0) Then
                    objMain.objApplication.StatusBar.SetText("For Item " & RCode & " and Line No. " & i & ".... Receipt Qty should not be zero!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Next
            If objForm.Items.Item("19").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Remarks Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function
    'Private Sub LoadSubGroups1(objForm As SAPbouiCOM.Form, itemGroup As String)
    '    ' Log or check value for debugging
    '    ' objMain.objApplication.MessageBox("Item Group received: " & itemGroup)
    '    ' === SAFETY CHECK ===
    '    If String.IsNullOrEmpty(itemGroup) OrElse Not IsNumeric(itemGroup.Trim()) Then
    '        ' Optional: show a message or log
    '        ' objMain.objApplication.MessageBox("Invalid or empty Item Group. Exiting.")
    '        Exit Sub
    '    End If
    '    Dim objCombo1 As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("SubGrp").Specific, SAPbouiCOM.ComboBox)
    '    ' === CLEAR EXISTING COMBO VALUES ===
    '    'If objCombo1.ValidValues.Count > 0 Then
    '    '    For i As Integer = objCombo1.ValidValues.Count - 1 To 0 Step -1
    '    '        objCombo1.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
    '    '    Next
    '    'End If
    '    ' === QUERY SUBGROUPS ===
    '    Dim oRS As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
    '    ' Build query with numeric-safe filter
    '    'Dim sql As String = "SELECT DISTINCT T0.""U_VSubGroup"" " &
    '    '                "FROM OITM T0 " &
    '    '                "INNER JOIN OITB T1 ON T0.""ItmsGrpCod"" = T1.""ItmsGrpCod"" " &
    '    '                "WHERE T0.""ItmsGrpCod"" = " & itemGroup & " " & ' No quotes for numeric value
    '    '                "AND T0.""U_VSubGroup"" IS NOT NULL " &
    '    '                "ORDER BY T0.""U_VSubGroup"""



    '    Dim sql As String = "SELECT DISTINCT T0.""U_VSubGroup"" " &
    '                    "FROM OITM T0 " &
    '                    "WHERE T0.""ItmsGrpCod"" = " & itemGroup & " " & ' No quotes for numeric value
    '                    "AND T0.""U_VSubGroup"" IS NOT NULL " &
    '                    "ORDER BY T0.""U_VSubGroup"""



    '    oRS.DoQuery(sql)
    '    ' === ADD TO COMBO AND SELECT FIRST ===
    '    Dim firstValue As String = ""
    '    While Not oRS.EoF
    '        Dim subGroup As String = oRS.Fields.Item(0).Value.ToString()
    '        If Not String.IsNullOrEmpty(subGroup) Then
    '            objCombo1.ValidValues.Add(subGroup, subGroup)
    '            If firstValue = "" Then
    '                firstValue = subGroup
    '            End If
    '        End If
    '        oRS.MoveNext()
    '    End While
    '    'If objCombo1.ValidValues.Count > 0 AndAlso Not String.IsNullOrEmpty(firstValue) Then
    '    '    objCombo1.Select(firstValue, SAPbouiCOM.BoSearchKey.psk_ByValue)
    '    'End If
    'End Sub



    '2nd
    'Private Sub LoadSubGroups1(objForm As SAPbouiCOM.Form, itemGroup As String)
    '    If String.IsNullOrEmpty(itemGroup) OrElse Not IsNumeric(itemGroup.Trim()) Then Exit Sub

    '    Dim objCombo1 As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("SubGrp").Specific, SAPbouiCOM.ComboBox)

    '    ' === Clear previous values ===
    '    If objCombo1.ValidValues.Count > 0 Then
    '        For i As Integer = objCombo1.ValidValues.Count - 1 To 0 Step -1
    '            objCombo1.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
    '        Next
    '    End If

    '    ' === Query subgroups ===
    '    Dim oRS As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

    '    Dim sql As String = "SELECT DISTINCT T0.""U_VSubGroup"" " &
    '                    "FROM OITM T0 " &
    '                    "WHERE T0.""ItmsGrpCod"" = " & itemGroup & " " &
    '                    "AND T0.""U_VSubGroup"" IS NOT NULL " &
    '                    "ORDER BY T0.""U_VSubGroup"""

    '    oRS.DoQuery(sql)

    '    ' === Optional: Add default empty option first ===
    '    objCombo1.ValidValues.Add("", "")

    '    ' === Add subgroups from query ===
    '    While Not oRS.EoF
    '        Dim subGroup As String = oRS.Fields.Item(0).Value.ToString().Trim()
    '        If Not String.IsNullOrEmpty(subGroup) Then
    '            objCombo1.ValidValues.Add(subGroup, subGroup)
    '        End If
    '        oRS.MoveNext()
    '    End While

    '    ' === Leave default unselected (shows "-- Select --") ===
    '    objCombo1.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
    'End Sub

    Private Sub LoadSubGroups1(objForm As SAPbouiCOM.Form, itemGroup As String)
        If String.IsNullOrEmpty(itemGroup) OrElse Not IsNumeric(itemGroup.Trim()) Then Exit Sub

        Dim objCombo1 As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("SubGrp").Specific, SAPbouiCOM.ComboBox)

        If objCombo1.ValidValues.Count > 0 Then
            For i As Integer = objCombo1.ValidValues.Count - 1 To 0 Step -1
                Try
                    objCombo1.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Catch ex As Exception
                End Try
            Next
        End If

        Try
            objCombo1.ValidValues.Add("", "")
        Catch ex As Exception
        End Try

        Dim oRS As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sql As String = "SELECT DISTINCT T0.""U_VSubGroup"" " &
                        "FROM OITM T0 " &
                        "WHERE T0.""ItmsGrpCod"" = " & itemGroup & " " &
                        "AND T0.""U_VSubGroup"" IS NOT NULL " &
                        "ORDER BY T0.""U_VSubGroup"""

        oRS.DoQuery(sql)

        ' === Add subgroups ===
        While Not oRS.EoF
            Dim subGroup As String = oRS.Fields.Item(0).Value.ToString().Trim()
            If Not String.IsNullOrEmpty(subGroup) Then
                Dim exists As Boolean = False
                For j As Integer = 0 To objCombo1.ValidValues.Count - 1
                    If objCombo1.ValidValues.Item(j).Value = subGroup Then
                        exists = True
                        Exit For
                    End If
                Next
                If Not exists Then
                    objCombo1.ValidValues.Add(subGroup, subGroup)
                End If
            End If
            oRS.MoveNext()
        End While
        ' === Leave selection on default
        objCombo1.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
    End Sub





    Private Sub LoadSubSubGroups(objForm As SAPbouiCOM.Form, itemGroup As String, subGroup As String)
        ' === SAFETY CHECKS ===
        If String.IsNullOrEmpty(itemGroup) OrElse Not IsNumeric(itemGroup.Trim()) Then
            Exit Sub
        End If
        Dim objCombo As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("SubSubGrp").Specific, SAPbouiCOM.ComboBox)
        ' === CLEAR EXISTING VALUES ===
        'If objCombo.ValidValues.Count > 0 Then
        '    For i As Integer = objCombo.ValidValues.Count - 1 To 0 Step -1
        '        objCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
        '    Next
        'End If
        ' === FETCH SUBSUBGROUPS ===
        Dim oRS As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        'Dim sql As String = "SELECT DISTINCT T0.""U_VSubSubGroup"" " &
        '                    "FROM OITM T0 " &
        '                    "INNER JOIN OITB T1 ON T0.""ItmsGrpCod"" = T1.""ItmsGrpCod"" " &
        '                    "WHERE T0.""ItmsGrpCod"" = " & itemGroup & " " & ' No quotes for numeric
        '                    "AND T0.""U_VSubSubGroup"" IS NOT NULL " &
        '                    "ORDER BY T0.""U_VSubSubGroup"""

        Try
            objCombo.ValidValues.Add("", "")
        Catch ex As Exception
        End Try

        Dim sql As String = "SELECT DISTINCT T0.""U_VSubSubGroup"" " &
                            "FROM OITM T0 " &
                            "WHERE T0.""ItmsGrpCod"" = " & itemGroup & " " &
                            "AND T0.""U_VSubSubGroup"" IS NOT NULL " &
                            "AND T0.""U_VSubGroup""='" & subGroup & "' " &
                            "ORDER BY T0.""U_VSubSubGroup"""


        oRS.DoQuery(sql)
        Dim firstValue As String = ""
        While Not oRS.EoF
            Dim subsubGroup As String = oRS.Fields.Item(0).Value.ToString()
            If Not String.IsNullOrEmpty(subsubGroup) Then
                objCombo.ValidValues.Add(subsubGroup, subsubGroup)
                If firstValue = "" Then
                    firstValue = subsubGroup
                End If
            End If
            oRS.MoveNext()
        End While
        ' === AUTO-SELECT FIRST VALUE ===
        'If objCombo.ValidValues.Count > 0 AndAlso Not String.IsNullOrEmpty(firstValue) Then
        '    objCombo.Select(firstValue, SAPbouiCOM.BoSearchKey.psk_ByValue)
        'End If
    End Sub

    'Private Sub LoadSubGroups(objForm As SAPbouiCOM.Form, itemGroup As String)
    '    Dim oCombo As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("SubGrp").Specific, SAPbouiCOM.ComboBox)
    '    ' Clear existing values and selection
    '    'oCombo.ValidValues.Clear()
    '    'oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index) ' Reset selection if needed

    '    ' Escape single quotes in itemGroup
    '    ' Dim safeItemGroup As String = itemGroup.Replace("'", "''")

    '    ' Query subgroups
    '    Dim oRS As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
    '    Dim sql As String = "SELECT DISTINCT T0.""U_VSubGroup"" " &
    '                        "FROM OITM T0 " &
    '                        "INNER JOIN OITB T1 ON T0. ""ItmsGrpCod"" =T1. ""ItmsGrpCod"" " &
    '                        "WHERE T0.""ItmsGrpCod"" = '" & itemGroup & "' " &
    '                        "AND T0.""U_VSubGroup"" IS NOT NULL " &
    '                        "ORDER BY T0.""U_VSubGroup"""
    '    oRS.DoQuery(sql)

    '    ' Add subgroups to combo box
    '    While Not oRS.EoF
    '        Dim subGroup As String = oRS.Fields.Item(0).Value.ToString()
    '        If Not String.IsNullOrEmpty(subGroup) Then
    '            'oCombo.ValidValues.Add(oRS.Fields.Item(0).Value.ToString(), "")

    '            OBJCOMBO.ValidValues.Add(subGroup, "")
    '        End If
    '        oRS.MoveNext()
    '    End While

    'End Sub
    'Private Sub LoadSubSubGroups(oForm As SAPbouiCOM.Form, itemGroup As String, subGroup As String)
    '    Dim oCombo As SAPbouiCOM.ComboBox = oForm.Items.Item("SubSubGrp").Specific
    '    'oCombo.ValidValues.Clear()
    '    Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    Dim sql As String = "SELECT DISTINCT T0.""U_VSubSubGroup"" " &
    '                        "FROM OITM T0 " &
    '                        "INNER JOIN OITB T1 ON T0. ""ItmsGrpCod"" =T1. ""ItmsGrpCod"" " &
    '                        "WHERE T0.""ItmsGrpCod"" = '" & itemGroup & "' " &
    '                        "AND T0.""U_VSubSubGroup"" IS NOT NULL " &
    '                        "ORDER BY T0.""U_VSubSubGroup"""
    '    oRS.DoQuery(sql)

    '    While Not oRS.EoF
    '        oCombo.ValidValues.Add(oRS.Fields.Item(0).Value.ToString(), "")
    '        oRS.MoveNext()
    '    End While
    'End Sub
    Private Function ValidateUpoad(ByVal formUID As String) As Boolean
        If (objForm.Items.Item("FRMDATE").Specific.Value = "") Or (objForm.Items.Item("TODATE").Specific.Value = "") Then
            objMain.objApplication.StatusBar.SetText("Please Fill header fields first", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        If oDS.GetValue("FRMDATE", 0) = "" Then
            objMain.objApplication.StatusBar.SetText("Please Choose From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        If oDS.GetValue("TODATE", 0) = "" Then
            objMain.objApplication.StatusBar.SetText("Please Choose To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        If oDS.GetValue("BPLID", 0) = "" Then
            objMain.objApplication.StatusBar.SetText("Please Select Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        If oDS.GetValue("FRMTIME", 0) = "" Then
            objMain.objApplication.StatusBar.SetText("Please Enter From Time", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        If oDS.GetValue("TOTIME", 0) = "" Then
            objMain.objApplication.StatusBar.SetText("Please Enter To Time", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        Return True
    End Function

    Private Sub LoadExcelToMatrix_Old(objForm As SAPbouiCOM.Form, filePath As String)
        Try
            Dim oDataTable As SAPbouiCOM.DataTable
            ' Check and create DT_MATRIX if not exists
            If objForm.DataSources.DataTables.Count = 0 OrElse objForm.DataSources.DataTables.Item("DT_MTX") Is Nothing Then
                objForm.DataSources.DataTables.Add("DT_MTX")
            End If
            oDataTable = objForm.DataSources.DataTables.Item("DT_MTX")
            ' If columns are not already added, add them
            If oDataTable.Columns.Count = 0 Then
                oDataTable.Columns.Add("Barcode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                oDataTable.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                oDataTable.Columns.Add("ItemDescription", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 254)
                oDataTable.Columns.Add("UOM", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50)
                'oDataTable.Columns.Add("FromDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
                'oDataTable.Columns.Add("ToDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
                oDataTable.Columns.Add("ItemCost", SAPbouiCOM.BoFieldsType.ft_Price, 20)
                oDataTable.Columns.Add("RetailPriceWithVAT", SAPbouiCOM.BoFieldsType.ft_Price, 20)
                oDataTable.Columns.Add("OfferPriceWithVAT", SAPbouiCOM.BoFieldsType.ft_Price, 20)
                'oDataTable.Columns.Add("DiscountAmount", SAPbouiCOM.BoFieldsType.ft_Price, 20)
                'oDataTable.Columns.Add("DiscountPercentage", SAPbouiCOM.BoFieldsType.ft_Percent, 20)
            End If
            ' Ensure DataTable exists
            If objForm.DataSources.DataTables.Count = 0 OrElse objForm.DataSources.DataTables.Item("DT_MTX") Is Nothing Then
                objForm.DataSources.DataTables.Add("DT_MTX")
            End If
            Dim objMatrix1 As SAPbouiCOM.Matrix = CType(objForm.Items.Item("IMtx").Specific, SAPbouiCOM.Matrix)
            Dim oDataSource As SAPbouiCOM.DataTable = objForm.DataSources.DataTables.Item("DT_MTX")
            ' Clear rows safely
            If oDataSource.Rows.Count > 0 Then
                oDataSource.Rows.Clear()
            End If
            ' Load Excel file using OLEDB
            Dim connStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties='Excel 12.0 Xml;HDR=YES';"
            Dim dt As New DataTable()
            Using conn As New OleDb.OleDbConnection(connStr)
                conn.Open()
                Dim cmd As New OleDb.OleDbCommand("SELECT * FROM [Sheet1$]", conn)
                Dim adapter As New OleDb.OleDbDataAdapter(cmd)
                adapter.Fill(dt)
                conn.Close()
            End Using
            ' Fill SAP DataTable
            For i As Integer = 0 To dt.Rows.Count - 1
                oDataTable.Rows.Add()
                oDataTable.SetValue("Barcode", i, dt.Rows(i)("Barcode").ToString())
                oDataTable.SetValue("ItemCode", i, dt.Rows(i)("Item Code").ToString())
                oDataTable.SetValue("ItemDescription", i, dt.Rows(i)("Item Description").ToString())
                oDataTable.SetValue("UOM", i, dt.Rows(i)("UOM").ToString())
                ' oDataTable.SetValue("FRMDATE", i, dt.Rows(i)("From Date").ToString())
                'oDataTable.SetValue("TODATE", i, dt.Rows(i)("To Date").ToString())
                oDataTable.SetValue("ItemCost", i, dt.Rows(i)("Item Cost").ToString())
                oDataTable.SetValue("RetailPriceWithVAT", i, dt.Rows(i)("Retail Price With VAT").ToString())
                oDataTable.SetValue("OfferPriceWithVAT", i, dt.Rows(i)("Offer Price With VAT").ToString())
                'oDataTable.SetValue("DiscountAmount", i, dt.Rows(i)("Discount Amount").ToString())
                'oDataTable.SetValue("DiscountPercentage", i, dt.Rows(i)("Discount Percentage").ToString())
            Next
            ' Load into matrix
            objMatrix1.LoadFromDataSource()
            'Dim objMatrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx").Specific, SAPbouiCOM.Matrix)
            ' Example: binding each column of matrix to datatable
            Dim col As SAPbouiCOM.Column
            ' Clear existing rows in matrix (optional)
            If objMatrix1.RowCount > 0 Then
                For i As Integer = objMatrix1.RowCount To 1 Step -1
                    objMatrix1.DeleteRow(i)
                Next
            End If
            ' Add rows and set values manually
            For i As Integer = 0 To oDataSource.Rows.Count - 1
                objMatrix1.AddRow()
                objMatrix1.Columns.Item("LineId").Cells.Item(i + 1).Specific.Value = (i + 1).ToString()
                objMatrix1.Columns.Item("BCODE").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("Barcode", i).ToString()
                objMatrix1.Columns.Item("ITMCODE").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("ItemCode", i).ToString()
                objMatrix1.Columns.Item("ITMNM").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("ItemDescription", i).ToString()
                objMatrix1.Columns.Item("UOM").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("UOM", i).ToString()
                'objMatrix.Columns.Item("FRMDATE").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("FromDate", i).ToString("yyyy-MM-dd")
                'objMatrix.Columns.Item("TODATE").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("ToDate", i).ToString("yyyy-MM-dd")
                objMatrix1.Columns.Item("ITMCST").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("ItemCost", i).ToString()
                objMatrix1.Columns.Item("RTPRWVAT").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("RetailPriceWithVAT", i).ToString()
                objMatrix1.Columns.Item("OFPRWVAT").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("OfferPriceWithVAT", i).ToString()
                objMatrix1.Columns.Item("DCPRWVAT").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("DiscountAmount", i).ToString()
                objMatrix1.Columns.Item("DPERCTG").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("DiscountPercentage", i).ToString()
            Next
        Catch ex As Exception
            objMain.objApplication.MessageBox("Error loading Excel to matrix: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadExcelToMatrix(objForm As SAPbouiCOM.Form, filePath As String)
        Try
            objMatrix = objForm.Items.Item("IMtx").Specific
            Dim oDataTable As SAPbouiCOM.DataTable
            ' Check and create DT_MATRIX if not exists
            If objForm.DataSources.DataTables.Count = 0 OrElse objForm.DataSources.DataTables.Item("DT_MTX") Is Nothing Then
                objForm.DataSources.DataTables.Add("DT_MTX")
            End If
            oDataTable = objForm.DataSources.DataTables.Item("DT_MTX")
            ' If columns are not already added, add them
            If oDataTable.Columns.Count = 0 Then
                oDataTable.Columns.Add("Barcode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                oDataTable.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                oDataTable.Columns.Add("ItemDescription", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 254)
                oDataTable.Columns.Add("UOM", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50)
                'oDataTable.Columns.Add("FromDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
                'oDataTable.Columns.Add("ToDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
                oDataTable.Columns.Add("ItemCost", SAPbouiCOM.BoFieldsType.ft_Price, 20)
                oDataTable.Columns.Add("RetailPriceWithVAT", SAPbouiCOM.BoFieldsType.ft_Price, 20)
                oDataTable.Columns.Add("OfferPriceWithVAT", SAPbouiCOM.BoFieldsType.ft_Price, 20)
                oDataTable.Columns.Add("DiscountAmount", SAPbouiCOM.BoFieldsType.ft_Price, 20)
                oDataTable.Columns.Add("DiscountPercentage", SAPbouiCOM.BoFieldsType.ft_Percent, 20)
            End If
            ' Ensure DataTable exists
            If objForm.DataSources.DataTables.Count = 0 OrElse objForm.DataSources.DataTables.Item("DT_MTX") Is Nothing Then
                objForm.DataSources.DataTables.Add("DT_MTX")
            End If
            Dim objMatrix1 As SAPbouiCOM.Matrix = CType(objForm.Items.Item("IMtx").Specific, SAPbouiCOM.Matrix)
            Dim oDataSource As SAPbouiCOM.DataTable = objForm.DataSources.DataTables.Item("DT_MTX")
            ' Clear rows safely
            If oDataSource.Rows.Count > 0 Then
                oDataSource.Rows.Clear()
            End If
            ' Load Excel file using OLEDB
            Dim connStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties='Excel 12.0 Xml;HDR=YES';"
            Dim dt As New DataTable()
            Using conn As New OleDb.OleDbConnection(connStr)
                conn.Open()
                Dim cmd As New OleDb.OleDbCommand("SELECT * FROM [Sheet1$]", conn)
                Dim adapter As New OleDb.OleDbDataAdapter(cmd)
                adapter.Fill(dt)
                conn.Close()
            End Using
            ' Fill SAP DataTable
            For i As Integer = 0 To dt.Rows.Count - 1
                oDataTable.Rows.Add()
                oDataTable.SetValue("Barcode", i, dt.Rows(i)("Barcode").ToString())
                oDataTable.SetValue("ItemCode", i, dt.Rows(i)("Item Code").ToString())
                oDataTable.SetValue("ItemDescription", i, dt.Rows(i)("Item Description").ToString())
                oDataTable.SetValue("UOM", i, dt.Rows(i)("UOM").ToString())
                ' oDataTable.SetValue("FRMDATE", i, dt.Rows(i)("From Date").ToString())
                'oDataTable.SetValue("TODATE", i, dt.Rows(i)("To Date").ToString())
                oDataTable.SetValue("ItemCost", i, dt.Rows(i)("Item Cost").ToString())
                oDataTable.SetValue("RetailPriceWithVAT", i, dt.Rows(i)("Retail Price With VAT").ToString())
                oDataTable.SetValue("OfferPriceWithVAT", i, dt.Rows(i)("Offer Price With VAT").ToString())
                oDataTable.SetValue("DiscountAmount", i, dt.Rows(i)("Discount Amount").ToString())
                oDataTable.SetValue("DiscountPercentage", i, dt.Rows(i)("Discount Percentage").ToString())
            Next
            ' Load into matrix
            objMatrix.Clear()
            oDS1.Clear()
            objMatrix.FlushToDataSource()

            'Dim objMatrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx").Specific, SAPbouiCOM.Matrix)
            ' Example: binding each column of matrix to datatable
            'Dim col As SAPbouiCOM.Column
            '' Clear existing rows in matrix (optional)
            'If objMatrix1.RowCount > 0 Then
            '    For i As Integer = objMatrix1.RowCount To 1 Step -1
            '        objMatrix1.DeleteRow(i)
            '    Next
            'End If
            ' Add rows and set values manually
            Dim FromDateStr As String = objForm.Items.Item("FRMDATE").Specific.Value
            Dim ToDateStr As String = objForm.Items.Item("TODATE").Specific.Value

            'Dim FromDate As Date = objForm.Items.Item("FRMDATE").Specific.Value
            'Dim ToDate As Date = objForm.Items.Item("TODATE").Specific.Value

            ' Naresh
            For i As Integer = 0 To oDataSource.Rows.Count - 1
                objMatrix.AddRow()
                oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
                oDS1.SetValue("Code", oDS1.Offset, oDataSource.GetValue("Barcode", i).ToString())
                oDS1.SetValue("U_ITMCODE", oDS1.Offset, oDataSource.GetValue("ItemCode", i).ToString())
                oDS1.SetValue("U_ITMNM", oDS1.Offset, oDataSource.GetValue("ItemDescription", i).ToString())
                oDS1.SetValue("U_UOM", oDS1.Offset, oDataSource.GetValue("UOM", i).ToString())
                oDS1.SetValue("U_FRMDATE", oDS1.Offset, FromDateStr)
                oDS1.SetValue("U_TODATE", oDS1.Offset, ToDateStr)
                oDS1.SetValue("U_ITMCST", oDS1.Offset, oDataSource.GetValue("ItemCost", i).ToString())
                oDS1.SetValue("U_RTPRWVAT", oDS1.Offset, oDataSource.GetValue("RetailPriceWithVAT", i).ToString())
                oDS1.SetValue("U_OFPRWVAT", oDS1.Offset, oDataSource.GetValue("OfferPriceWithVAT", i).ToString())
                oDS1.SetValue("U_DCPRWVAT", oDS1.Offset, oDataSource.GetValue("DiscountAmount", i).ToString())
                oDS1.SetValue("U_DPERCTG", oDS1.Offset, oDataSource.GetValue("DiscountPercentage", i).ToString())
                objMatrix.SetLineData(objMatrix.VisualRowCount)

                'objMatrix1.AddRow()
                'objMatrix1.Columns.Item("LineId").Cells.Item(i + 1).Specific.Value = (i + 1).ToString()
                'objMatrix1.Columns.Item("BCODE").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("Barcode", i).ToString()
                'objMatrix1.Columns.Item("ITMCODE").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("ItemCode", i).ToString()
                'objMatrix1.Columns.Item("ITMNM").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("ItemDescription", i).ToString()
                'objMatrix1.Columns.Item("UOM").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("UOM", i).ToString()
                ''objMatrix.Columns.Item("FRMDATE").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("FromDate", i).ToString("yyyy-MM-dd")
                ''objMatrix.Columns.Item("TODATE").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("ToDate", i).ToString("yyyy-MM-dd")
                'objMatrix1.Columns.Item("ITMCST").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("ItemCost", i).ToString()
                'objMatrix1.Columns.Item("RTPRWVAT").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("RetailPriceWithVAT", i).ToString()
                'objMatrix1.Columns.Item("OFPRWVAT").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("OfferPriceWithVAT", i).ToString()
                ''objMatrix1.Columns.Item("DCPRWVAT").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("DiscountAmount", i).ToString()
                ''objMatrix1.Columns.Item("DPERCTG").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("DiscountPercentage", i).ToString()
            Next
            objMatrix.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.MessageBox("Error loading Excel to matrix: " & ex.Message)
        End Try
    End Sub
    Private Function ValidatePromotionName(ByVal formUID As String) As Boolean

        Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(formUID)
        Dim oEditText As SAPbouiCOM.EditText = objForm.Items.Item("PRMNM").Specific
        Dim promoName As String = oEditText.Value.Trim()
        If String.IsNullOrWhiteSpace(promoName) Then
            objMain.objApplication.StatusBar.SetText("Promotion Name cannot be empty.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        '  Max 10 characters
        If promoName.Length > 10 Then
            'objMain.objApplication.MessageBox("Promotion Name must be a maximum of 10 characters.")
            objMain.objApplication.StatusBar.SetText("Promotion Name must be a maximum of 10 characters.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        ' No spaces allowed
        If promoName.Contains(" ") Then
            'objMain.objApplication.MessageBox("Promotion Name Should not contain spaces.")
            objMain.objApplication.StatusBar.SetText("Promotion Name Should not contain spaces.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        '  Must be unique — check in the database
        Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs.DoQuery("SELECT COUNT(*) AS Cnt FROM ""@TNX_ITEMDISC"" WHERE ""U_PRMNM""= '" & promoName.Replace("'", "''") & "'")
        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And rs.Fields.Item("Cnt").Value > 1 Then
            objMain.objApplication.StatusBar.SetText("Promotion Name already exists. Please enter a unique name", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False

        ElseIf objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And rs.Fields.Item("Cnt").Value > 0 Then
            ' objMain.objApplication.MessageBox("Promotion Name already exists. Please enter a unique name.")
            objMain.objApplication.StatusBar.SetText("Promotion Name already exists. Please enter a unique name", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If






        Return True
    End Function
    Public Function IsItemInOverlappingPromotion(fromDate As DateTime, toDate As DateTime, fromTime As String, toTime As String, itemCode As String) As Boolean
        ' Assuming fromTime and toTime are in format "HHmm" (e.g., "1300" for 1:00 PM)
        ' U_FRMTIME and U_TOTIME must also be stored as strings or integers in the same format (not TIME type)
        'Dim query As String =
        '"SELECT COUNT(*) FROM ""@TNX_ITEMDISC_C0"" T0 " &
        '"INNER JOIN ""@TNX_ITEMDISC"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
        '"WHERE T0.""U_ITMCODE"" = '" & itemCode & "' " &
        '"AND T1.""U_STATUS"" = 'Active' " &
        '"AND ( " &
        '"    ('" & fromDate.ToString("yyyy-MM-dd") & "' <= T1.""U_TODATE"" " &
        '"     AND '" & toDate.ToString("yyyy-MM-dd") & "' >= T1.""U_FRMDATE"") " &
        '"    AND " &
        '"    (" &
        '"        '" & fromTime & "' <= T1.""U_TOTIME"" AND '" & toTime & "' >= T1.""U_FRMTIME"" " &
        '"    ) " &
        '")"
        Dim docnum As String = objForm.Items.Item("DocNum").Specific.Value
        Dim query As String = "SELECT COUNT(*) FROM ""@TNX_ITEMDISC"" H1 " &
"JOIN ""@TNX_ITEMDISC_C0"" C1 ON H1.""DocEntry"" = C1.""DocEntry"" " &
"JOIN ""@TNX_ITEMDISC"" H2 ON H1.""U_ITMCODE"" = H2.""U_ITMCODE"" " &
"JOIN ""@TNX_ITEMDISC_C0"" C2 ON H2.""DocEntry"" = C2.""DocEntry"" " &
"WHERE " &
"C1.""U_ITMCODE"" = C2.""U_ITMCODE"" " &
"AND H1.""DocEntry"" <> H2.""DocEntry"" " &
"AND H1.""U_FRMDATE"" <= H2.""U_TODATE"" " &
"AND H1.""U_TODATE"" >= H2.""U_FRMDATE"" " &
"AND H1.""U_FRMTIME"" <= H2.""U_TOTIME"" " &
"AND H1.""U_TOTIME"" >= H2.""U_FRMTIME"" " &
"AND H1.""DocNum""='" & docnum & "' "

        Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'oRS = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        oRS.DoQuery(query)
        Return oRS.RecordCount > 0 AndAlso Convert.ToInt32(oRS.Fields.Item(0).Value) > 0
    End Function
    Function validateDiscountType(ByVal FormUiD As String)
        Try

            objForm = objMain.objApplication.Forms.Item(FormUiD)
            If objForm.Items.Item("DISCTYP").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText(" Discount Type Cannot be Null", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True
        Catch ex As Exception

        End Try
    End Function

    Function validateBranch(ByVal FormUiD As String)
        Try
            'For i As Integer = 1 To objMatrix2.VisualRowCount
            '    If objMatrix2.Columns.Item("CHK").Cells.Item(1).Specific.Value.ToString().Trim() <> "" Then
            '        'objMain.objApplication.StatusBar.SetText("Please Choose Different selection And press Load Button", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '        objMain.objApplication.StatusBar.SetText("Please select branch row", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '        Return False
            '    End If
            'Next
            'Return True

            Dim oCheck As SAPbouiCOM.CheckBox, selectedCount As Integer = 0
            For i As Integer = 1 To objMatrix2.VisualRowCount
                If CType(objMatrix2.Columns.Item("CHK").Cells.Item(i).Specific, SAPbouiCOM.CheckBox).Checked Then selectedCount += 1
            Next
            If selectedCount <> 1 Then objMain.objApplication.StatusBar.SetText("Please select the branch in row", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Return False
            Return True



            ''Dim oCheck As SAPbouiCOM.CheckBox
            ''For i As Integer = 1 To objMatrix2.VisualRowCount
            ''    oCheck = CType(objMatrix2.Columns.Item("CHK").Cells.Item(i).Specific, SAPbouiCOM.CheckBox)
            ''    If oCheck.Checked = False Then
            ''        objMain.objApplication.StatusBar.SetText("Please select Branch row", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            ''        Return False
            ''    End If
            ''Next
            ' Return True
        Catch ex As Exception
        End Try
    End Function

    'Function validateDiscountpercentage(ByVal FormUiD As String)
    '    Try
    '        For i As Integer = 1 To objMatrix.VisualRowCount
    '            If objMatrix.Columns.Item("DPERCTG").Cells.Item(i).Specific.Value.ToString().Trim() = "" OrElse
    '           CDbl(objMatrix.Columns.Item("DPERCTG").Cells.Item(i).Specific.Value) = 0 Then

    '                objMain.objApplication.StatusBar.SetText("Please select Discount Percentage", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                Return False
    '            End If
    '        Next

    '        Return True
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        Return False
    '    End Try
    'End Function



    'Function validateDiscountpercentage(ByVal FormUiD As String)
    '    Try
    '        Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUiD)

    '        Dim oCombo As SAPbouiCOM.ComboBox = CType(oForm.Items.Item("DISCTYP").Specific, SAPbouiCOM.ComboBox)
    '        If oCombo.Selected IsNot Nothing AndAlso oCombo.Selected.Value = "P" Then
    '            For i As Integer = 1 To objMatrix.VisualRowCount
    '                Dim valueStr As String = objMatrix.Columns.Item("DPERCTG").Cells.Item(i).Specific.Value.ToString().Trim()
    '                If valueStr = "" OrElse CDbl(valueStr) = 0 Then
    '                    objMain.objApplication.StatusBar.SetText("Please enter a valid Discount Percentage in row " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                    Return False
    '                End If
    '            Next
    '        End If

    '        Return True
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        Return False
    '    End Try
    'End Function

    'Function validateDiscountAmount(ByVal FormUiD As String)
    '    Try
    '        Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUiD)

    '        Dim oCombo As SAPbouiCOM.ComboBox = CType(oForm.Items.Item("DISCTYP").Specific, SAPbouiCOM.ComboBox)
    '        If oCombo.Selected IsNot Nothing AndAlso oCombo.Selected.Value = "A" Then
    '            For i As Integer = 1 To objMatrix.VisualRowCount
    '                Dim valueStr As String = objMatrix.Columns.Item("DCPRWVAT").Cells.Item(i).Specific.Value.ToString().Trim()
    '                If valueStr = "" OrElse CDbl(valueStr) = 0 Then
    '                    objMain.objApplication.StatusBar.SetText("Please enter a valid Discount Amount in row " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                    Return False
    '                End If
    '            Next
    '        End If

    '        Return True
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        Return False
    '    End Try
    'End Function
    Function validateDiscountAmount(ByVal FormUiD As String)
        Try
            Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUiD)
            Dim oCombo As SAPbouiCOM.ComboBox = CType(oForm.Items.Item("DISCTYP").Specific, SAPbouiCOM.ComboBox)

            ' Only validate if Discount Type is "A"
            If oCombo.Selected IsNot Nothing AndAlso oCombo.Selected.Value = "A" Then
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                    ' Get Item Code from matrix row
                    Dim itemCode As String = objMatrix.Columns.Item("ITMCODE").Cells.Item(i).Specific.Value.ToString().Trim()
                    If itemCode <> "" Then
                        Dim valueStr As String = objMatrix.Columns.Item("DCPRWVAT").Cells.Item(i).Specific.Value.ToString().Trim()
                        If valueStr = "" OrElse CDbl(valueStr) = 0 Then
                            objMain.objApplication.StatusBar.SetText("Please enter a valid Discount Amount in row " & i & " for item " & itemCode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                Next
            End If

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function



    Function ValidateItem(ByVal FormUiD As String) As Boolean
        Try
            Dim itemcode As String = CType(objMatrix.Columns.Item("ITMCODE").Cells.Item(1).Specific, SAPbouiCOM.EditText).Value.Trim()

            If itemcode <> "" Then
                For i As Integer = 1 To objMatrix.VisualRowCount
                    If itemcode = objMatrix.Columns.Item("ITMCODE").Cells.Item(i).Specific.Value.Trim() Then
                        ' Select the row when a match is found
                        objMatrix.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Exit For
                    End If
                Next
                Return True
            Else
                objMain.objApplication.StatusBar.SetText("Please Enter ItemCode", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Function validateDiscountpercentage(ByVal FormUiD As String)
        Try
            Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUiD)
            Dim oCombo As SAPbouiCOM.ComboBox = CType(oForm.Items.Item("DISCTYP").Specific, SAPbouiCOM.ComboBox)
            If oCombo.Selected IsNot Nothing AndAlso oCombo.Selected.Value = "P" Then
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                    Dim itemCode As String = objMatrix.Columns.Item("ITMCODE").Cells.Item(i).Specific.Value.ToString().Trim()
                    If itemCode <> "" Then
                        Dim valueStr As String = objMatrix.Columns.Item("DPERCTG").Cells.Item(i).Specific.Value.ToString().Trim()

                        If valueStr = "" OrElse CDbl(valueStr) = 0 Then
                            objMain.objApplication.StatusBar.SetText("Please enter a valid Discount Percentage in row " & i & " for item " & itemCode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                Next
            End If

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function



End Class