
'new
Imports System.Collections.Specialized
Imports System.Configuration
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports SAPbobsCOM

Public Class clsBuyAndGetLeastItem

#Region "       Declaration             "
    Public objForm, objSTRForm As SAPbouiCOM.Form
    'Dim oDS, oDS1 As SAPbouiCOM.DBDataSource
    Dim objMatrix, objMatrix1, objMatrix2, objMatrix3, obj As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Public oDS, oDS1, oDS2, oDS3, oDS4 As SAPbouiCOM.DBDataSource
    Public PrgBar As SAPbouiCOM.ProgressBar
    Dim MATRIXS, ExcelFilePath As String
    Dim LostFocusFlag As Boolean = True


#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("BuyandGetLeastItems3.xml", "LEASTITM", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("LEASTITM", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            oDS = objForm.DataSources.DBDataSources.Item("@TNX_BGLEASTITMS")
            oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_BGLEASTITMS_C0")
            oDS4 = objForm.DataSources.DBDataSources.Item("@TNX_BGLEASTITMS_C3")

            objMatrix = objForm.Items.Item("Mtx").Specific

            Me.BranchLoad(objForm.UniqueID)
            Dim sQury As String = "Select ""ItmsGrpCod"" , ""ItmsGrpNam""   From OITB  "
            objComboBox = objForm.Items.Item("IGroup").Specific
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, sQury)
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("IGroup").DisplayDesc = True
            Dim ItemGroup As String
            objForm.EnableMenu("774", True)
            Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'Me.GetNextDocNum()
            Me.objForm.Items.Item("DOCDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Me.objForm.Items.Item("DOCDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Me.objForm.Items.Item("DOCDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Me.objForm.Items.Item("DOCDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            Dim oComb As SAPbouiCOM.ComboBox = objForm.Items.Item("BPLID").Specific

            'objMain.objUtilities.ComboBoxLoadValues(oComb, "Select ""BPLId"",""BPLName"" FROM OBPL Where ""BPLId""<>'1'")
            objMain.objUtilities.ComboBoxLoadValues(oComb, "Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join " & " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""   ")

            oComb.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("BPLID").DisplayDesc = True

            'oDS = Me.objForm.DataSources.DBDataSources.Add("@SBO_CLEAR")
            'oDS1 = Me.objForm.DataSources.DBDataSources.Add("@SBO_CLEAR_C0")


            'Dim oRsBranch As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'Dim Query As String = "SELECT T2.""BPLId"",T2.""BPLName"" AS ""Branch"" FROM USR6 T0  INNER JOIN OUSR T1 ON T0.""UserCode"" = T1.""USER_CODE"" " &
            '        " INNER JOIN OBPL T2 ON T0.""BPLId"" = T2.""BPLId"" WHERE T1.""USER_CODE"" = '" & objMain.objCompany.UserName.ToString().Trim & "' "
            'oRsBranch.DoQuery(Query)

            'objComboBox = objForm.Items.Item("Branch").Specific
            'objMain.objUtilities.ComboBoxLoadValues(objComboBox, Query)
            'objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            'objForm.Items.Item("Branch").DisplayDesc = True
            'Try
            '    If objComboBox.ValidValues.Count <> 0 Then
            '        objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
            '    End If
            'Catch ex As Exception
            'End Try

            objForm.Items.Item("0_U_FD").Width = objForm.Items.Item("0_U_FD").Width + 180
            objForm.Items.Item("Item_0").Left = objForm.Items.Item("Item_0").Left + 160
            'Me.SetDefault(objForm.UniqueID)
            objForm.Items.Item("0_U_FD").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            'objComboBox = objForm.Items.Item("17").Specific
            'objComboBox.ValidValues.Add("", "")
            'objComboBox.ValidValues.Add("Open", "Open")
            'objComboBox.ValidValues.Add("Close", "Close")
            'objComboBox.ValidValues.Add("Cancel", "Cancel")


            'Dim WhsCmbo As SAPbouiCOM.ComboBox = DirectCast(Me.objForm.Items.Item("Whs").Specific, SAPbouiCOM.ComboBox)
            'Dim WhsQry As String = "SELECT T0.""WhsCode"", T0.""WhsName"" FROM OWHS T0 Where  ""U_Transit"" = 'No' "
            'objMain.objUtilities.ComboBoxLoadValues(WhsCmbo, WhsQry)
            ''WhsCmbo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            'If WhsCmbo.ValidValues.Count <> 0 Then
            '    WhsCmbo.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
            'End If

            'Dim UOMCombo As SAPbouiCOM.ComboBox = DirectCast(Me.objMatrix.Columns.Item("UOM").Cells.Item(1).Specific, SAPbouiCOM.ComboBox)
            'Dim ReceiptUOMCombo As SAPbouiCOM.ComboBox = DirectCast(Me.objMatrix.Columns.Item("RUOM").Cells.Item(1).Specific, SAPbouiCOM.ComboBox)
            'Dim UOMQry As String = "SELECT DISTINCT A.""UomEntry"",(SELECT ""UomCode"" FROM OUOM WHERE ""UomEntry""=A.""UomEntry"") AS ""UomName"" FROM OBCD A "
            'objMain.objUtilities.ComboBoxLoadValues(UOMCombo, UOMQry)
            'objMain.objUtilities.ComboBoxLoadValues(ReceiptUOMCombo, UOMQry)

            'Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'Me.objForm.Items.Item("17").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'Me.objForm.Items.Item("17").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Function GetNextDocNum()
        Dim Str As String
        Dim oRs As SAPbobsCOM.Recordset
        Dim DocNum As Integer
        oRs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Str = "SELECT IFNULL(MAX(""DocNum""),0)+1 as ""DocNum"" FROM ""@TNX_BGLEASTITMS"""



        ' Str = "SELECT MAX(""DocNum"") FROM ""@TNX_BGLEASTITMS"""

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


    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ME_LEASTITMA" And pVal.BeforeAction = False Then
                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
                'Me.GetNextDocNum()
            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                Me.SetNewLine(objForm.UniqueID, MATRIXS)
                'Me.GetNextDocNum()
            ElseIf pVal.MenuUID = "774" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("LEASTITM", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix = objForm.Items.Item("Mtx").Specific

                Dim row As Integer = objMatrix.VisualRowCount
                If MATRIXS.Equals("Mtx") = True Then
                    If objMatrix.IsRowSelected(1) <> True And objMatrix.VisualRowCount < 1 Then
                        objMatrix.AddRow()
                        oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
                        objMatrix.SetLineData(objMatrix.VisualRowCount)
                    End If
                    If objMatrix.IsRowSelected(row) = True Then
                        objMatrix.DeleteRow(row)
                    Else
                        For i As Integer = 1 To objMatrix.VisualRowCount - 1
                            If objMatrix.IsRowSelected(i) = True Then
                                If i <> 1 Then
                                    objMatrix.DeleteRow(i)
                                End If
                            End If
                        Next

                    End If
                End If

                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                End If
            End If
            'Try
            '    If pVal.MenuUID = "Delete Row" And pVal.BeforeAction = False Then
            '        objForm = objMain.objApplication.Forms.GetForm("Clear", objMain.objApplication.Forms.ActiveForm.TypeCount)
            '        objMatrix = objForm.Items.Item("Mtx").Specific
            '        For i As Integer = 1 To objMatrix.VisualRowCount - 1
            '            If objMatrix.IsRowSelected(i) = True Then
            '                objMatrix.DeleteRow(i)
            '            End If
            '        Next
            '        For i As Integer = 1 To objMatrix.VisualRowCount
            '            objMatrix.Columns.Item("LineId").Cells.Item(i).Specific.string = i
            '        Next
            '        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            '    End If

            '    If pVal.MenuUID = "519" And pVal.BeforeAction = True Then 'Preview
            '        Try
            '            Dim frs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '            frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'ClearanceSalesRpt' AND ""Type"" = 'C'")
            '            If frs.RecordCount = 0 Then
            '                objMain.objApplication.MessageBox("Report layout 'ClearanceSalesRpt' not found. Please upload with same name and Code as parameter", 0, "OK")
            '            Else
            '                objMain.objApplication.ActivateMenuItem(frs.Fields.Item(0).Value.ToString())
            '                oDS = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR")
            '                Dim CrForm As SAPbouiCOM.Form
            '                Dim oedt As SAPbouiCOM.EditText
            '                CrForm = objMain.objApplication.Forms.ActiveForm
            '                oedt = CrForm.Items.Item("1000003").Specific
            '                oedt.Value = oDS.GetValue("DocEntry", 0)
            '                CrForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            '            End If
            '        Catch ex As Exception
            '        End Try
            '    End If
            '    If pVal.MenuUID = "520" And pVal.BeforeAction = True Then 'Print
            '        Try
            '            Dim frs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '            frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'ClearanceSalesRpt' AND ""Type"" = 'C'")
            '            If frs.RecordCount = 0 Then
            '                objMain.objApplication.MessageBox("Report layout 'ClearanceSalesRpt' not found. Please upload with same name and Code as parameter", 0, "OK")
            '            Else
            '                objMain.objApplication.ActivateMenuItem(frs.Fields.Item(0).Value.ToString())
            '                oDS = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR")
            '                Dim CrForm As SAPbouiCOM.Form
            '                Dim oedt As SAPbouiCOM.EditText
            '                CrForm = objMain.objApplication.Forms.ActiveForm
            '                oedt = CrForm.Items.Item("1000003").Specific
            '                oedt.Value = oDS.GetValue("DocEntry", 0)
            '                CrForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            '            End If
            '        Catch ex As Exception
            '        End Try
            '    End If

            '    If pVal.MenuUID = "Close" And pVal.BeforeAction = False Then
            '        objForm = objMain.objApplication.Forms.GetForm("Clear", objMain.objApplication.Forms.ActiveForm.TypeCount)
            '        objMatrix = objForm.Items.Item("Mtx").Specific
            '        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            '            objMain.objApplication.MessageBox("Form should be in 'OK Mode'.. ", 1, "Ok", "Cancel")
            '            Exit Sub
            '        End If
            '        Dim result As Integer = objMain.objApplication.MessageBox("Changing document status to close is irreversible..Confirm... Do you want to continue?", 1, "Continue", "Cancel")
            '        If (result = 1) Then
            '            Me.UpdateStatus(objForm.UniqueID)
            '        End If
            '    End If
            'Catch ex As Exception
            '    objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'End Try
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    'Private Sub UpdateStatus(ByVal FormUID As String)
    '    Try
    '        objForm.Freeze(True)

    '        oDS = Me.objForm.DataSources.DBDataSources.Add("@SBO_CLEAR")
    '        oDS1 = Me.objForm.DataSources.DBDataSources.Add("@SBO_CLEAR_C0")
    '        objMatrix = objForm.Items.Item("Mtx").Specific

    '        objMain.sCmp = objMain.objCompany.GetCompanyService
    '        objMain.oGeneralService = objMain.sCmp.GetGeneralService("SBO_CLEARUDO")
    '        objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
    '        objMain.oGeneralParams.SetProperty("DocEntry", oDS.GetValue("DocEntry", 0))
    '        objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
    '        objMain.oGeneralData.SetProperty("U_MESTS", "Close")
    '        objMain.oGeneralService.Update(objMain.oGeneralData)
    '        objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@SBO_CLEAR"), "DocEntry", oDS.GetValue("DocEntry", 0))
    '        objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@SBO_CLEAR_C0"), "DocEntry", oDS1.GetValue("DocEntry", 0))

    '        objMatrix.LoadFromDataSource()

    '        objForm.Refresh()
    '        objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
    '        objMain.objApplication.StatusBar.SetText("DocStatus updated successfully...!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    '        objForm.Freeze(False)
    '    Catch ex As Exception
    '        objForm.Freeze(False)
    '        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '    End Try
    'End Sub
    Public Shared Sub FILTER_CFL(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef oApplication As SAPbouiCOM.Application, ByRef oCompany As SAPbobsCOM.Company, ByVal Query As String, ByVal Code As String, ByVal ItemCode As String, ByVal Subgrp As String, ByVal subsubgrp As String)
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


        End Try


    End Sub

    Public Shared Sub FILTER_CFL1(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef oApplication As SAPbouiCOM.Application, ByRef oCompany As SAPbobsCOM.Company, ByVal Query As String, ByVal Code As String, ByVal ItemCode As String, ByVal Subgrp As String, ByVal subsubgrp As String)
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
                oCon.[Alias] = ItemCode
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = oRec.Fields.Item(ItemCode).Value.ToString()
            Else


                For i As Integer = 0 To oRec.RecordCount - 1
                    oCon = oCons.Add()
                    oCon.[Alias] = ItemCode
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRec.Fields.Item(ItemCode).Value.ToString()

                    If i <> oRec.RecordCount - 1 Then
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    End If
                    oRec.MoveNext()
                Next
            End If


            oCFL.SetConditions(oCons)
        Catch ex As Exception


        End Try


    End Sub
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try

            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If (pVal.ItemUID = "Mtx" Or pVal.ItemUID = "BMtx") And pVal.BeforeAction = False Then ' Or pVal.ItemUID = "Matrx5" Or pVal.ItemUID = "Matrix" Then
                        MATRIXS = pVal.ItemUID
                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "Mtx" And pVal.ColUID = "BCODE" And pVal.BeforeAction = False Then
                        Try
                            objForm.Freeze(True)
                            Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx").Specific
                            If pVal.ItemUID = "Mtx" Then
                                If pVal.Row = objMatrix.VisualRowCount Then
                                    Me.SetNewLine(objForm.UniqueID, "Mtx")
                                End If
                            End If
                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                        End Try

                    End If



                    '    If pVal.ItemUID = "Branch" And pVal.BeforeAction = False Then
                    '        Dim objBranchCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Branch").Specific
                    '        Dim BranchID As String
                    '        Try
                    '            BranchID = objBranchCombo.Selected.Value
                    '        Catch ex As Exception
                    '            objMain.objApplication.StatusBar.SetText("Please select Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        End Try
                    '        Dim WhsCmbo As SAPbouiCOM.ComboBox = DirectCast(Me.objForm.Items.Item("Whs").Specific, SAPbouiCOM.ComboBox)
                    '        Dim WhsQry As String = "SELECT T0.""WhsCode"", T0.""WhsName"" FROM OWHS T0 Where ""BPLid""='" & BranchID & "' And ""U_Transit"" = 'No' "
                    '        objMain.objUtilities.ComboBoxLoadValues(WhsCmbo, WhsQry)
                    '        Try
                    '            If WhsCmbo.ValidValues.Count > 0 Then
                    '                WhsCmbo.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
                    '            End If
                    '        Catch ex As Exception
                    '        End Try
                    '    End If
                    '    If pVal.ItemUID = "Whs" And pVal.BeforeAction = False And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    '        Dim objWhsCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Whs").Specific
                    '        objMatrix = objForm.Items.Item("Mtx").Specific
                    '        Dim Whs As String
                    '        Try
                    '            Whs = objWhsCombo.Selected.Value
                    '        Catch ex As Exception
                    '            objMain.objApplication.StatusBar.SetText("Please select warehouse", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        End Try
                    '        For i As Integer = 1 To objMatrix.VisualRowCount
                    '            If objMatrix.Columns.Item("ICode").Cells.Item(i).Specific.Value.ToString() <> "" Then
                    '                oDS1.SetValue("LineID", oDS1.Offset, i)
                    '                oDS1.SetValue("U_ICode", oDS1.Offset, objMatrix.Columns.Item("ICode").Cells.Item(i).Specific.Value)
                    '                oDS1.SetValue("U_IName", oDS1.Offset, objMatrix.Columns.Item("IName").Cells.Item(i).Specific.Value)
                    '                oDS1.SetValue("U_BarCode", oDS1.Offset, objMatrix.Columns.Item("BarCode").Cells.Item(i).Specific.Value)
                    '                oDS1.SetValue("U_UOM", oDS1.Offset, objMatrix.Columns.Item("UOM").Cells.Item(i).Specific.Value)
                    '                oDS1.SetValue("U_Price", oDS1.Offset, Me.GetFieldValue("SELECT T0.""AvgPrice"" FROM OITW T0 WHERE T0.""ItemCode""='" & objMatrix.Columns.Item("ICode").Cells.Item(i).Specific.Value & "' and  T0.""WhsCode"" ='" & Whs & "'"))
                    '                oDS1.SetValue("U_Qty", oDS1.Offset, objMatrix.Columns.Item("Qty").Cells.Item(i).Specific.Value)

                    '                oDS1.SetValue("U_RBarCode", oDS1.Offset, objMatrix.Columns.Item("RBarCode").Cells.Item(i).Specific.Value)
                    '                oDS1.SetValue("U_RCode", oDS1.Offset, objMatrix.Columns.Item("RCode").Cells.Item(i).Specific.Value)
                    '                oDS1.SetValue("U_RName", oDS1.Offset, objMatrix.Columns.Item("RName").Cells.Item(i).Specific.Value)
                    '                oDS1.SetValue("U_RUOM", oDS1.Offset, objMatrix.Columns.Item("RUOM").Cells.Item(i).Specific.Value)
                    '                oDS1.SetValue("U_RQty", oDS1.Offset, objMatrix.Columns.Item("RQty").Cells.Item(i).Specific.Value)

                    '                oDS1.SetValue("U_Prod", oDS1.Offset, "")
                    '                oDS1.SetValue("U_RProd", oDS1.Offset, "")
                    '                objMatrix.SetLineData(i)

                    '            End If
                    '        Next
                    '    End If
                    'Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    '    objForm = objMain.objApplication.Forms.Item(FormUID)
                    '    oDS = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR")
                    '    oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR_C0")
                    '    objMatrix = objForm.Items.Item("Mtx").Specific
                    '    Dim objBranchCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Branch").Specific
                    '    Dim BranchID As String


                    '    If pVal.ItemUID = "Mtx" And pVal.ColUID = "BarCode" Then
                    '        If objMatrix.Columns.Item("BarCode").Cells.Item(pVal.Row).Specific.Value <> "" Then
                    '            Try
                    '                BranchID = objBranchCombo.Selected.Value
                    '            Catch ex As Exception
                    '                objMain.objApplication.StatusBar.SetText("Please select Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '            End Try
                    '            Dim GetItemCode As String = " Select ""ItemCode"" ,(Select A.""InvntryUom"" From OITM A Where  A.""ItemCode"" = B.""ItemCode"") AS ""PUoMEntry"", (Select A.""ItemName"" From OITM A Where  A.""ItemCode"" = B.""ItemCode"") ""ItemName"" " &
                    '                                            " From OBCD B Where ""BcdCode"" = '" & objMatrix.Columns.Item("BarCode").Cells.Item(pVal.Row).Specific.Value & "'"
                    '            Dim oRsGetItemCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '            oRsGetItemCode.DoQuery(GetItemCode)

                    '            oDS1.SetValue("LineID", oDS1.Offset, pVal.Row)
                    '            oDS1.SetValue("U_ICode", oDS1.Offset, oRsGetItemCode.Fields.Item("ItemCode").Value)
                    '            oDS1.SetValue("U_IName", oDS1.Offset, oRsGetItemCode.Fields.Item("ItemName").Value)
                    '            oDS1.SetValue("U_BarCode", oDS1.Offset, objMatrix.Columns.Item("BarCode").Cells.Item(pVal.Row).Specific.Value)
                    '            oDS1.SetValue("U_UOM", oDS1.Offset, oRsGetItemCode.Fields.Item("PUoMEntry").Value)
                    '            oDS1.SetValue("U_Price", oDS1.Offset, Me.GetFieldValue("SELECT T0.""AvgPrice"" FROM OITW T0 WHERE T0.""ItemCode""='" & oRsGetItemCode.Fields.Item("ItemCode").Value & "' and  T0.""WhsCode"" ='" & oDS.GetValue("U_Whs", 0) & "'"))
                    '            oDS1.SetValue("U_Qty", oDS1.Offset, objMatrix.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value)
                    '            oDS1.SetValue("U_SSTK", oDS1.Offset, Me.GetFieldValue("SELECT T0.""OnHand"" FROM OITW T0 WHERE T0.""ItemCode"" ='" & oRsGetItemCode.Fields.Item("ItemCode").Value & "' and  T0.""WhsCode"" ='" & oDS.GetValue("U_Whs", 0) & "'"))

                    '            oDS1.SetValue("U_RBarCode", oDS1.Offset, objMatrix.Columns.Item("RBarCode").Cells.Item(pVal.Row).Specific.Value)
                    '            oDS1.SetValue("U_RCode", oDS1.Offset, objMatrix.Columns.Item("RCode").Cells.Item(pVal.Row).Specific.Value)
                    '            oDS1.SetValue("U_RName", oDS1.Offset, objMatrix.Columns.Item("RName").Cells.Item(pVal.Row).Specific.Value)
                    '            oDS1.SetValue("U_RUOM", oDS1.Offset, objMatrix.Columns.Item("RUOM").Cells.Item(pVal.Row).Specific.Value)
                    '            oDS1.SetValue("U_RQty", oDS1.Offset, objMatrix.Columns.Item("RQty").Cells.Item(pVal.Row).Specific.Value)

                    '            oDS1.SetValue("U_Prod", oDS1.Offset, "")
                    '            oDS1.SetValue("U_RProd", oDS1.Offset, "")
                    '            objMatrix.SetLineData(pVal.Row)

                    '            If pVal.Row = objMatrix.VisualRowCount Then
                    '                Me.SetNewLine(objForm.UniqueID)
                    '            End If
                    '        End If
                    '    End If

                    '    If pVal.ItemUID = "Mtx" And pVal.ColUID = "RBarCode" Then
                    '        If objMatrix.Columns.Item("RBarCode").Cells.Item(pVal.Row).Specific.Value <> "" Then
                    '            Try
                    '                BranchID = objBranchCombo.Selected.Value
                    '            Catch ex As Exception
                    '                objMain.objApplication.StatusBar.SetText("Please select Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '            End Try
                    '            Dim GetItemCode As String = " Select ""ItemCode"" ,(Select A.""InvntryUom"" From OITM A Where  A.""ItemCode"" = B.""ItemCode"") AS ""PUoMEntry"", (Select A.""ItemName"" From OITM A Where  A.""ItemCode"" = B.""ItemCode"") ""ItemName"" " &
                    '                                        " From OBCD B Where ""BcdCode"" = '" & objMatrix.Columns.Item("RBarCode").Cells.Item(pVal.Row).Specific.Value & "'"
                    '            Dim oRsGetItemCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '            oRsGetItemCode.DoQuery(GetItemCode)

                    '            oDS1.SetValue("LineID", oDS1.Offset, pVal.Row)
                    '            oDS1.SetValue("U_ICode", oDS1.Offset, objMatrix.Columns.Item("ICode").Cells.Item(pVal.Row).Specific.Value)
                    '            oDS1.SetValue("U_IName", oDS1.Offset, objMatrix.Columns.Item("IName").Cells.Item(pVal.Row).Specific.Value)
                    '            oDS1.SetValue("U_BarCode", oDS1.Offset, objMatrix.Columns.Item("BarCode").Cells.Item(pVal.Row).Specific.Value)
                    '            oDS1.SetValue("U_UOM", oDS1.Offset, objMatrix.Columns.Item("UOM").Cells.Item(pVal.Row).Specific.Value)
                    '            oDS1.SetValue("U_Price", oDS1.Offset, objMatrix.Columns.Item("Price").Cells.Item(pVal.Row).Specific.Value)
                    '            oDS1.SetValue("U_Qty", oDS1.Offset, objMatrix.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value)
                    '            oDS1.SetValue("U_SSTK", oDS1.Offset, Me.GetFieldValue("SELECT T0.""OnHand"" FROM OITW T0 WHERE T0.""ItemCode"" ='" & oRsGetItemCode.Fields.Item("ItemCode").Value & "' and  T0.""WhsCode"" ='" & oDS.GetValue("U_Whs", 0) & "'"))

                    '            oDS1.SetValue("U_RBarCode", oDS1.Offset, objMatrix.Columns.Item("RBarCode").Cells.Item(pVal.Row).Specific.Value)
                    '            oDS1.SetValue("U_RCode", oDS1.Offset, oRsGetItemCode.Fields.Item("ItemCode").Value)
                    '            oDS1.SetValue("U_RName", oDS1.Offset, oRsGetItemCode.Fields.Item("ItemName").Value)
                    '            oDS1.SetValue("U_RUOM", oDS1.Offset, oRsGetItemCode.Fields.Item("PUoMEntry").Value)
                    '            oDS1.SetValue("U_RQty", oDS1.Offset, objMatrix.Columns.Item("RQty").Cells.Item(pVal.Row).Specific.Value)

                    '            oDS1.SetValue("U_Prod", oDS1.Offset, "")
                    '            oDS1.SetValue("U_RProd", oDS1.Offset, "")
                    '            objMatrix.SetLineData(pVal.Row)

                    '            If pVal.Row = objMatrix.VisualRowCount Then
                    '                Me.SetNewLine(objForm.UniqueID)
                    '            End If
                    '        End If
                    '    End If
                    '    If pVal.ItemUID = "Mtx" And pVal.ColUID = "ICode" Then
                    '        If objMatrix.VisualRowCount = pVal.Row Then
                    '            If objMatrix.Columns.Item("ICode").Cells.Item(pVal.Row).Specific.Value <> "" Then
                    '                Me.SetNewLine(objForm.UniqueID)
                    '            End If
                    '        End If
                    '    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDS = objForm.DataSources.DBDataSources.Item("@TNX_BGLEASTITMS")
                    oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_BGLEASTITMS_C0")
                    objMatrix = objForm.Items.Item("Mtx").Specific
                    '  objForm = objMain.objApplication.Forms.Item(FormUID)
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    'objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                    'If oCFL.UniqueID = "CFL_2" And pVal.BeforeAction = True Then 'CHOOSE FORM TABLE UNIQUNE ID
                    '    'Dim itemCode As String = objForm.DataSources.DBDataSources.Item("@TNX_BGLEASTITMS").GetValue("U_IGroup", 0).Trim()
                    '    'Me.barcodecfl(pVal.FormUID, "CFL_2")
                    '    Me.barcodecfl(objMain.objApplication.Forms.Item(pVal.FormUID), "CFL_2")
                    'End If
                    If (oCFL.UniqueID = "CFL_2" Or oCFL.UniqueID = "CFL_ITEM") And pVal.BeforeAction = True Then
                        ''Me.CFLFilter("CFL_OCRD")
                        Dim StrCFLQry As String = ""
                        If objForm.Items.Item("IGroup").Specific.Value <> "" Then


                            '           StrCFLQry = "SELECT B.""ItemCode"",  B.""BcdCode"" ,I.""U_VSubGroup"",I.""U_VSubSubGroup"",
                            '              (SELECT ""GET_RSP""(I.""ItemCode"",""UomEntry"",1) FROM DUMMY) as ""SALESPRICE"" " &
                            '"FROM ""OBCD"" B " &
                            '"INNER JOIN ""OITM"" I ON B.""ItemCode"" = I.""ItemCode"" " &
                            '"WHERE I.""ItmsGrpCod"" = '" & objForm.Items.Item("IGroup").Specific.Value & "' and I.""U_VSubGroup"" ='" & objForm.Items.Item("SubGrp").Specific.Value & "' AND I.""U_VSubSubGroup""='" & objForm.Items.Item("SSbGrp").Specific.Value & "'  "

                            StrCFLQry = "SELECT B.""ItemCode"",  B.""BcdCode"" ,I.""U_VSubGroup"",I.""U_VSubSubGroup""
                 FROM ""OBCD"" B " &
                 "INNER JOIN ""OITM"" I ON B.""ItemCode"" = I.""ItemCode"" " &
                 "WHERE   B.""ItemCode"" IS NOT NULL "


                            If CStr(objForm.Items.Item("IGroup").Specific.Value).Trim <> "" Then
                                StrCFLQry = StrCFLQry & "  and  I.""ItmsGrpCod"" = '" & objForm.Items.Item("IGroup").Specific.Value & "'"
                            End If

                            If CStr(objForm.Items.Item("SubGrp").Specific.Value).Trim <> "" Then
                                StrCFLQry = StrCFLQry & "  And  I.""U_VSubGroup"" = '" & objForm.Items.Item("SubGrp").Specific.Value & "'"
                            End If

                            If CStr(objForm.Items.Item("SSbGrp").Specific.Value).Trim <> "" Then
                                StrCFLQry = StrCFLQry & "  And  I.""U_VSubSubGroup"" = '" & objForm.Items.Item("SSbGrp").Specific.Value & "'"
                            End If


                            ' ---


                            'StrCFLQry = "SELECT T1.""ItemCode"", T1.""ItemName"", T0.""BcdCode"", T2.""UomCode"" " &
                            '                       "FROM ""OBCD"" T0 INNER JOIN ""OITM"" T1 ON T0.""ItemCode"" = T1.""ItemCode"" INNER JOIN OUOM T2 ON " &
                            '                        " T0.""UomEntry"" = T2.""UomEntry"" Where T1.""ItmsGrpCod"" = '" & objForm.Items.Item("IGroup").Specific.Value & "'"



                        End If
                        If oCFL.UniqueID = "CFL_ITEM" Then
                            FILTER_CFL1(pVal, objMain.objApplication, objMain.objCompany, StrCFLQry, "BcdCode", "ItemCode", "U_VSubGroup", "U_VSubSubGroup")
                        End If

                        If oCFL.UniqueID = "CFL_2" Then
                            FILTER_CFL(pVal, objMain.objApplication, objMain.objCompany, StrCFLQry, "BcdCode", "ItemCode", "U_VSubGroup", "U_VSubSubGroup")
                        End If


                    End If
                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)


                        'If oCFL.UniqueID = "CFL_ITEM" And pVal.BeforeAction = False Then
                        '    oDS1.SetValue("LineId", oDS1.Offset, pVal.Row)

                        '    Dim Brch As String = objForm.Items.Item("BPLID").Specific.Value

                        '    Dim GetBarCode As String = " SELECT Top 1 ""BcdCode"" FROM OBCD WHERE  " &
                        '                                  "  ""ItemCode"" = '" & oDT.GetValue("ItemCode", 0) & "' ORDER BY ""BcdEntry"" ASC "
                        '    Dim oRsGetBarCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '    oRsGetBarCode.DoQuery(GetBarCode)
                        '    Dim Barcode As String = oRsGetBarCode.Fields.Item("BcdCode").Value.ToString()
                        '    oDS1.SetValue("Code", oDS1.Offset, Barcode)

                        '    oDS1.SetValue("U_ITMCODE", oDS1.Offset, oDT.GetValue("ItemCode", 0))
                        '    oDS1.SetValue("U_ITMNM", oDS1.Offset, oDT.GetValue("ItemName", 0))
                        '    oDS1.SetValue("U_UOM", oDS1.Offset, oDT.GetValue("InvntryUom", 0))

                        '    '  oDS1.SetValue("U_UOM", oDS1.Offset, objMatrix.Columns.Item("UOM").Cells.Item(pVal.Row).Specific.Value)

                        '    'oDS1.SetValue("U_UOM", oDS1.Offset, oDT.GetValue("UomEntry", 0))
                        '    objMatrix.SetLineData(pVal.Row)
                        '    If pVal.Row = objMatrix.VisualRowCount Then
                        '        Me.SetNewLine(objForm.UniqueID, "Mtx")
                        '    End If

                        '    'PRICE CODE
                        '    ' Get price using stored procedure
                        '    Dim query2 As String = "CALL ""GetRetailPrice""('" & Barcode & "','" & Brch.Trim & "')"
                        '    Dim orset2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '    orset2.DoQuery(query2)

                        '    If orset2.RecordCount > 0 Then
                        '        Dim Retailprice As Double = orset2.Fields.Item("Price").Value
                        '        oDS1.SetValue("U_PRICE", oDS1.Offset, Retailprice)
                        '        objMatrix.SetLineData(pVal.Row)
                        '    End If

                        '    ' Add new line ONLY ONCE after all processing
                        '    If pVal.Row = objMatrix.VisualRowCount Then
                        '        Me.SetNewLine(objForm.UniqueID, "Mtx")
                        '    End If

                        '    'PRICE CODE
                        If oCFL.UniqueID = "CFL_ITEM" And pVal.BeforeAction = False Then

                            Dim Brch As String = objForm.Items.Item("BPLID").Specific.Value
                            Dim Itemcode As String = ""


                            Dim GetBarCode As String = " SELECT Top 1 ""BcdCode"" FROM OBCD WHERE  " &
                                "  ""ItemCode"" = '" & oDT.GetValue("ItemCode", 0) & "' ORDER BY ""BcdEntry"" ASC "
                            Dim oRsGetBarCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetBarCode.DoQuery(GetBarCode)
                            Dim Barcode As String = oRsGetBarCode.Fields.Item("BcdCode").Value.ToString()
                            oDS1.SetValue("LineId", oDS1.Offset, pVal.Row)

                            oDS1.SetValue("Code", oDS1.Offset, Barcode)

                            oDS1.SetValue("U_ITMCODE", oDS1.Offset, oDT.GetValue("ItemCode", 0))
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







                            ' Optional UOM assignment from matrix, commented out
                            ' oDS1.SetValue("U_UOM", oDS1.Offset, objMatrix.Columns.Item("UOM").Cells.Item(pVal.Row).Specific.Value)
                            ' oDS1.SetValue("U_UOM", oDS1.Offset, oDT.GetValue("UomEntry", 0))

                            objMatrix.SetLineData(pVal.Row)

                            ' Get Item Cost
                            Try
                                Dim Itemcod As String = objMatrix.Columns.Item("ITMCODE").Cells.Item(1).Specific.Value
                                Dim query1 As String = "CALL ""GetItemPrice""('" & Itemcod & "','" & Brch.Trim & "')"
                                Dim orset1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                orset1.DoQuery(query1)

                                If orset1.RecordCount > 0 Then
                                    Dim Itemcost As Double = orset1.Fields.Item("AvgPrice").Value
                                    oDS1.SetValue("U_ITCST", oDS1.Offset, Itemcost)
                                Else
                                    oDS1.SetValue("U_ITCST", oDS1.Offset, 0)
                                End If
                            Catch ex As Exception
                            End Try


                            ' Get Item Cost End --------------------------------------------





                            ' === PRICE CODE ===
                            Dim query2 As String = "CALL ""GetRetailPrice""('" & Barcode & "','" & Brch.Trim & "')"
                            Dim orset2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            orset2.DoQuery(query2)

                            If orset2.RecordCount > 0 Then
                                Dim Retailprice As Double = orset2.Fields.Item("Price").Value
                                oDS1.SetValue("U_PRICE", oDS1.Offset, Retailprice)
                                objMatrix.SetLineData(pVal.Row)
                            End If

                            If pVal.Row = objMatrix.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "Mtx")
                            End If

                            'Me.SetNewLine(objForm.UniqueID, "Mtx")
                        End If
                        '---change

                        'If oCFL.UniqueID = "CFL_3" Then
                        '    oDS3.SetValue("U_BRCODE",  .Offset, oDT.GetValue("BcdCode", 0))
                        '    oDS3.SetValue("U_ITCODE", oDS3.Offset, oDT.GetValue("ItemCode", 0))
                        '    ' oDS3.SetValue("U_ITNM", oDS3.Offset, oDT.GetValue("ItemName", 0))
                        '    ' oDS4.SetValue("U_UOM", oDS4.Offset, oDT.GetValue("InvntryUom", 0))
                        '    objMatrix2.SetLineData(pVal.Row)
                        '    If objMatrix2.VisualRowCount = pVal.Row Then
                        '        Me.SetNewLine(objForm.UniqueID, "EMtx")
                        '    End If
                        'End If


                        If oCFL.UniqueID = "CFL_3" And pVal.BeforeAction = False Then
                            Dim bcode As String = oDT.GetValue("BcdCode", 0)
                            Dim FromDate As String = oDS.GetValue("U_FRMDATE", 0)
                            Dim ToDate As String = oDS.GetValue("U_TODATE", 0)
                            Dim Itemcode As String = ""
                            oDS3.SetValue("U_BRCODE", oDS3.Offset, oDT.GetValue("BcdCode", 0))
                            Dim itmqury As String = "SELECT T1.""ItemCode"", T1.""ItemName"", T0.""BcdCode"", T2.""UomCode"" " &
                                                    "FROM ""OBCD"" T0 INNER JOIN ""OITM"" T1 ON T0.""ItemCode"" = T1.""ItemCode"" INNER JOIN OUOM T2 ON " &
                                                        " T0.""UomEntry"" = T2.""UomEntry"" Where T0.""BcdCode"" = '" & bcode & "'"
                            Dim oRitmqury As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRitmqury.DoQuery(itmqury)
                            oDS3.SetValue("LineID", oDS3.Offset, pVal.Row)
                            oDS3.SetValue("U_ITCODE", oDS3.Offset, oRitmqury.Fields.Item("ItemCode").Value)
                            oDS3.SetValue("U_ITNM", oDS3.Offset, oRitmqury.Fields.Item("ItemName").Value)
                            oDS3.SetValue("U_BRCODE", oDS3.Offset, oRitmqury.Fields.Item("BcdCode").Value)

                            objMatrix2.SetLineData(pVal.Row)
                            If pVal.Row = objMatrix2.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "EMtx")
                            End If

                        End If



                        If oCFL.UniqueID = "CFL_Ex" Then
                            oDS3.SetValue("LineID", oDS3.Offset, pVal.Row)

                            oDS3.SetValue("U_BRCODE", oDS3.Offset, oDT.GetValue("CodeBars", 0))
                            oDS3.SetValue("U_ITCODE", oDS3.Offset, oDT.GetValue("ItemCode", 0))
                            oDS3.SetValue("U_ITNM", oDS3.Offset, oDT.GetValue("ItemName", 0))
                            'oDS3.SetValue("U_UOM", oDS3.Offset, oDT.GetValue("InvntryUom", 0))

                            objMatrix2.SetLineData(pVal.Row)
                            If objMatrix2.VisualRowCount = pVal.Row Then
                                Me.SetNewLine(objForm.UniqueID, "EMtx")
                            End If
                        End If

                        '-----------------------item codes------------------------
                        '               If oCFL.UniqueID = "CFL_2" And pVal.BeforeAction = False Then
                        '                   Dim Brch As String = objForm.Items.Item("BPLID").Specific.Value

                        '                   oDS1.SetValue("LineId", oDS1.Offset, pVal.Row)
                        '                   oDS1.SetValue("Code", oDS1.Offset, oDT.GetValue("BcdCode", 0))
                        '                   oDS1.SetValue("U_ITMCODE", oDS1.Offset, oDT.GetValue("ItemCode", 0))
                        '                   oDS1.SetValue("U_ITMNM", oDS1.Offset, oDT.GetValue("BcdName", 0))

                        '                   'oDS1.SetValue("U_UOM", oDS1.Offset, oDT.GetValue("UomEntry", 0))

                        '                   'Uom ------
                        '                   Dim barcode As String = oDT.GetValue("BcdCode", 0)

                        '                   Dim UOMQry As String

                        '                   UOMQry = "SELECT DISTINCT A.""UomEntry"", " &
                        '"(SELECT ""UomCode"" FROM OUOM WHERE ""UomEntry"" = A.""UomEntry"") AS ""UomName"" " &
                        '"FROM OBCD A " &
                        '"WHERE A.""BcdCode"" = '" & barcode & "'"

                        '                   Dim oRecordset22 As SAPbobsCOM.Recordset
                        '                   oRecordset22 = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '                   oRecordset22.DoQuery(UOMQry)

                        '                   Dim uomEntry As String = oRecordset22.Fields.Item("UomName").Value.ToString()
                        '                   oDS1.SetValue("U_UOM", oDS1.Offset, uomEntry)

                        '                   ' Use uomEntry and uomName as needed
                        '                   '-----------------------------------------------GET ITEM COST---------------

                        '                   Try
                        '                       Dim Itemcod As String = objMatrix.Columns.Item("ITMCODE").Cells.Item(1).Specific.Value
                        '                       Dim query1 As String = "CALL ""GetItemPrice""('" & Itemcod & "','" & Brch.Trim & "')"
                        '                       Dim orset1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '                       orset1.DoQuery(query1)

                        '                       If orset1.RecordCount > 0 Then
                        '                           Dim Itemcost As Double = orset1.Fields.Item("AvgPrice").Value
                        '                           oDS1.SetValue("U_ITCST", oDS1.Offset, Itemcost)
                        '                       Else
                        '                           oDS1.SetValue("U_ITCST", oDS1.Offset, 0)
                        '                       End If
                        '                   Catch ex As Exception
                        '                   End Try




                        '                   ' --------------------- END GET ITEM COST -------------

                        '                   'end uom

                        '                   'Price  SELLING 

                        '                   Dim GetBarCode As String = " SELECT Top 1 ""BcdCode"",""BcdName"" FROM OBCD WHERE ""ItemCode"" = '" & oDT.GetValue("ItemCode", 0) & "' ORDER BY ""BcdEntry"" ASC "
                        '                   Dim oRsGetBarCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '                   oRsGetBarCode.DoQuery(GetBarCode)
                        '                   Dim Barcode1 As String = oRsGetBarCode.Fields.Item("BcdCode").Value.ToString()
                        '                   Dim Barcodenm As String = oRsGetBarCode.Fields.Item("BcdName").Value.ToString()

                        '                   ' Dim Uom As String = oRsGetBarCode.Fields.Item("UomEntry").Value.ToString()

                        '                   Dim query2 As String = "CALL ""GetRetailPrice""('" & Barcode1 & "','" & Brch.Trim & "')"
                        '                   Dim orset2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '                   orset2.DoQuery(query2)

                        '                   If orset2.RecordCount > 0 Then
                        '                       Dim Retailprice As Double = orset2.Fields.Item("Price").Value
                        '                       oDS1.SetValue("U_PRICE", oDS1.Offset, Retailprice)
                        '                       'oDS1.SetValue("U_ITMNM", oDS1.Offset, Barcodenm)

                        '                       'oDS1.SetValue("U_UOM", oDS1.Offset, Uom)

                        '                       'objMatrix.SetLineData(pVal.Row)
                        '                   End If

                        '                   'end price
                        '                   objMatrix.SetLineData(pVal.Row)
                        '                   If pVal.Row = objMatrix.VisualRowCount Then
                        '                       Me.SetNewLine(objForm.UniqueID, "Mtx")
                        '                   End If
                        '                   'Me.SetNewLine(objForm.UniqueID, "Mtx")
                        '               End If
                        '-----------item codes-------------------------------
                        If oCFL.UniqueID = "CFL_2" And pVal.BeforeAction = False Then
                            Dim Brch As String = objForm.Items.Item("BPLID").Specific.Value

                            oDS1.SetValue("LineId", oDS1.Offset, pVal.Row)
                            oDS1.SetValue("Code", oDS1.Offset, oDT.GetValue("BcdCode", 0))
                            oDS1.SetValue("U_ITMCODE", oDS1.Offset, oDT.GetValue("ItemCode", 0))
                            oDS1.SetValue("U_ITMNM", oDS1.Offset, oDT.GetValue("BcdName", 0))

                            ' Uom ------
                            Dim barcode As String = oDT.GetValue("BcdCode", 0)
                            Dim UOMQry As String
                            Dim Itemcode As String = ""

                            '               UOMQry = "SELECT DISTINCT A.""UomEntry"", " &
                            '"(SELECT ""UomCode"" FROM OUOM WHERE ""UomEntry"" = A.""UomEntry"") AS ""UomName"" " &
                            '"FROM OBCD A " &
                            '"WHERE A.""BcdCode"" = '" & barcode & "'"


                            UOMQry = "SELECT T1.""ItemCode"", T1.""ItemName"", T0.""BcdCode"", T2.""UomEntry"", T2.""UomCode"" " &
         "FROM ""OBCD"" T0 " &
         "INNER JOIN ""OITM"" T1 ON T0.""ItemCode"" = T1.""ItemCode"" " &
         "INNER JOIN ""OUOM"" T2 ON T0.""UomEntry"" = T2.""UomEntry"" " &
         "WHERE T0.""BcdCode"" = '" & barcode & "'"

                            Dim oRecordset22 As SAPbobsCOM.Recordset
                            oRecordset22 = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordset22.DoQuery(UOMQry)

                            If oRecordset22.RecordCount > 0 Then
                                Dim uomEntry As String = oRecordset22.Fields.Item("UomCode").Value.ToString()
                                oDS1.SetValue("U_UOM", oDS1.Offset, uomEntry)
                            End If
                            Itemcode = oRecordset22.Fields.Item("ItemCode").Value

                            ' ------------------------------------ GET ITEM COST ------------------------------------
                            Try
                                'Dim Itemcod As String = objMatrix.Columns.Item("ITMCODE").Cells.Item(pVal.Row).Specific.Value
                                Dim query1 As String = "CALL ""GetItemPrice""('" & Itemcode & "','" & Brch.Trim() & "')"
                                Dim orset1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                orset1.DoQuery(query1)

                                If orset1.RecordCount > 0 Then
                                    Dim Itemcost As Double = orset1.Fields.Item("AvgPrice").Value
                                    oDS1.SetValue("U_ITCST", oDS1.Offset, Itemcost)
                                Else
                                    oDS1.SetValue("U_ITCST", oDS1.Offset, 0)
                                End If
                            Catch ex As Exception
                                ' Optionally log the error
                            End Try
                            ' ----------------------------------------------------------------------------------------

                            ' Price (SELLING)
                            Dim GetBarCode As String = "SELECT Top 1 ""BcdCode"", ""BcdName"" FROM OBCD WHERE ""ItemCode"" = '" & oDT.GetValue("ItemCode", 0) & "' ORDER BY ""BcdEntry"" ASC"



                            Dim oRsGetBarCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetBarCode.DoQuery(GetBarCode)

                            If oRsGetBarCode.RecordCount > 0 Then
                                Dim Barcode1 As String = oRsGetBarCode.Fields.Item("BcdCode").Value.ToString()
                                Dim Barcodenm As String = oRsGetBarCode.Fields.Item("BcdName").Value.ToString()

                                Dim query2 As String = "CALL ""GetRetailPrice""('" & Barcode1 & "','" & Brch.Trim() & "')"
                                Dim orset2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                orset2.DoQuery(query2)

                                If orset2.RecordCount > 0 Then
                                    Dim Retailprice As Double = orset2.Fields.Item("Price").Value
                                    oDS1.SetValue("U_PRICE", oDS1.Offset, Retailprice)
                                End If
                            End If


                            'Price ccode

                            objMatrix.SetLineData(pVal.Row)

                            If pVal.Row = objMatrix.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "Mtx")
                            End If
                        End If






















                    End If


                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    '-------LOADING FUNCTIONALITY-------------

                    'If pVal.ItemUID = "btnGO" And pVal.BeforeAction = False Then
                    '    If objForm.Items.Item("SBCDE").Specific.Value <> "" Then
                    '        Dim bcode As String = objForm.Items.Item("SBCDE").Specific.Value
                    '        Dim Afterchk As String = " "
                    '        Dim TYPE As String = objForm.Items.Item("PRTYP").Specific.Value
                    '        Select Case TYPE
                    '            Case "1"
                    '                For i As Integer = 1 To objMatrix.VisualRowCount
                    '                    If bcode = objMatrix.Columns.Item("BCODE").Cells.Item(i).Specific.Value Then
                    '                        objMatrix.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    '                        ' objForm.Items.Item("SBCDE").Specific.Value = Afterchk
                    '                        Exit For
                    '                    End If
                    '                Next
                    '                'Case "ItemGroupWise"
                    '                '    For i As Integer = 1 To objMatrix3.VisualRowCount
                    '                '        If bcode = objMatrix3.Columns.Item("BCODE").Cells.Item(i).Specific.Value Then
                    '                '            objMatrix3.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    '                '            ' objForm.Items.Item("SBCDE").Specific.Value = Afterchk
                    '                '            Exit For
                    '                '        End If
                    '                '    Next
                    '        End Select
                    '    Else
                    '        objMain.objApplication.StatusBar.SetText("Please Enter Barcode", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    End If
                    'End If
                    If pVal.ItemUID = "btnGO" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("SBCDE").Specific.Value <> "" Then
                            ' Force UI to commit value of combo box (PRTYP)
                            'objForm.Items.Item("PRTYP").Click()

                            Dim bcode As String = objForm.Items.Item("SBCDE").Specific.Value
                            Dim TYPE As String = objForm.Items.Item("PRTYP").Specific.Value

                            For i As Integer = 1 To objMatrix.VisualRowCount
                                If bcode = objMatrix.Columns.Item("BCODE").Cells.Item(i).Specific.Value Then
                                    objMatrix.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    Exit For
                                End If
                            Next
                        Else
                            objMain.objApplication.StatusBar.SetText("Please Enter Barcode", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                    End If






                    'LOAD OLD


                    '                If pVal.ItemUID = "LOAD" AndAlso pVal.BeforeAction = False AndAlso
                    '(objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or
                    ' objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or
                    ' objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                    '                    Try
                    '                        objForm.Freeze(True)

                    '                        Dim oDS As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Add("@TNX_BGLEASTITMS")
                    '                        Dim oDS2 As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Add("@TNX_BGLEASTITMS_C0")

                    '                        Dim objMatrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx").Specific, SAPbouiCOM.Matrix)
                    '                        Dim docnum As String = objForm.Items.Item("DocNum").Specific.Value
                    '                        Dim protype As String = objForm.Items.Item("PRTYP").Specific.Value
                    '                        Dim Itemgrp As String = objForm.Items.Item("IGroup").Specific.Value
                    '                        Dim subsub As String = objForm.Items.Item("SubGrp").Specific.Value
                    '                        Dim subsubgrp As String = objForm.Items.Item("SSbGrp").Specific.Value

                    '                        Dim rs1 As String =
                    '        "SELECT I.""ItemCode"", I.""ItemName"", I.""U_VSubGroup"", I.""U_VSubSubGroup"" " &
                    '        "FROM ""OBCD"" B " &
                    '        "INNER JOIN ""OITM"" I ON B.""ItemCode"" = I.""ItemCode"" " &
                    '        "WHERE I.""ItmsGrpCod"" = '" & Itemgrp & "' " &
                    '        "OR I.""U_VSubGroup"" = '" & subsub & "' " &
                    '        "OR I.""U_VSubSubGroup"" = '" & subsubgrp & "'"

                    '                        Dim ors1 As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                    '                        ors1.DoQuery(rs1)

                    '                        If objMatrix.RowCount > 0 Then
                    '                            For i As Integer = objMatrix.RowCount To 1 Step -1
                    '                                objMatrix.DeleteRow(i)
                    '                            Next
                    '                        End If
                    '                        If ors1.RecordCount > 0 Then
                    '                            For j As Integer = 0 To ors1.RecordCount - 1
                    '                                objMatrix.AddRow()
                    '                                Dim currentRow As Integer = objMatrix.VisualRowCount
                    '                                oDS2.SetValue("LineId", oDS2.Offset, currentRow.ToString())
                    '                                Dim itmCode As String = ors1.Fields.Item("ItemCode").Value.ToString()
                    '                                oDS2.SetValue("U_ITMCODE", oDS2.Offset, itmCode)
                    '                                oDS2.SetValue("U_ITMNM", oDS2.Offset, ors1.Fields.Item("ItemName").Value.ToString())

                    '                                Dim GetBarCode As String = "SELECT TOP 1 ""UomEntry"",""BcdCode"" FROM ""OBCD"" WHERE ""ItemCode"" = '" & itmCode & "' ORDER BY ""BcdEntry"" ASC"
                    '                                Dim oRsGetBarCode As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                    '                                oRsGetBarCode.DoQuery(GetBarCode)

                    '                                'If Not oRsGetBarCode.EoF Then
                    '                                'oDS2.SetValue("Code", oDS2.Offset, oRsGetBarCode.Fields.Item("BcdCode").Value.ToString())
                    '                                'End If
                    '                                If oRsGetBarCode.RecordCount > 0 Then
                    '                                    oDS2.SetValue("Code", oDS2.Offset, oRsGetBarCode.Fields.Item("BcdCode").Value.ToString())
                    '                                    oDS2.SetValue("U_UOM", oDS2.Offset, oRsGetBarCode.Fields.Item("UomEntry").Value.ToString())

                    '                                End If
                    '                                objMatrix.SetLineData(currentRow)
                    '                                ors1.MoveNext()
                    '                            Next
                    '                        Else
                    '                            Me.SetNewLine(objForm.UniqueID, "Mtx")
                    '                        End If

                    '                        objMatrix.AutoResizeColumns()
                    '                        objMatrix.FlushToDataSource()

                    '                    Catch ex As Exception
                    '                        objForm.StatusBar.SetText("Error: " & ex.Message & " - Page Load Events", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '                    Finally
                    '                        objForm.Freeze(False)
                    '                    End Try

                    '                End If

                    '-----------------------END LOAD
                    'working  load functionality

                    '                 If pVal.ItemUID = "LOAD" AndAlso pVal.BeforeAction = False AndAlso
                    '(objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or
                    ' objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or
                    ' objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                    '                     Try
                    '                         objForm.Freeze(True)

                    '                         Dim oDS As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Add("@TNX_BGLEASTITMS")
                    '                         Dim oDS2 As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Add("@TNX_BGLEASTITMS_C0")

                    '                         Dim objMatrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx").Specific, SAPbouiCOM.Matrix)
                    '                         Dim docnum As String = objForm.Items.Item("DocNum").Specific.Value
                    '                         Dim protype As String = objForm.Items.Item("PRTYP").Specific.Value
                    '                         Dim Itemgrp As String = objForm.Items.Item("IGroup").Specific.Value
                    '                         Dim subsub As String = objForm.Items.Item("SubGrp").Specific.Value
                    '                         Dim subsubgrp As String = objForm.Items.Item("SSbGrp").Specific.Value

                    '                         ' Build SQL query based on filter priority
                    '                         Dim rs1 As String = ""
                    '                         If Not String.IsNullOrEmpty(subsubgrp) Then
                    '                             rs1 = "SELECT I.""ItemCode"", I.""ItemName"", I.""U_VSubGroup"", I.""U_VSubSubGroup"" " &
                    '               "FROM ""OITM"" I " &
                    '               "WHERE I.""U_VSubSubGroup"" = '" & subsubgrp & "'"
                    '                         ElseIf Not String.IsNullOrEmpty(subsub) Then
                    '                             rs1 = "SELECT I.""ItemCode"", I.""ItemName"", I.""U_VSubGroup"", I.""U_VSubSubGroup"" " &
                    '               "FROM ""OITM"" I " &
                    '               "WHERE I.""U_VSubGroup"" = '" & subsub & "'"
                    '                         ElseIf Not String.IsNullOrEmpty(Itemgrp) Then
                    '                             rs1 = "SELECT I.""ItemCode"", I.""ItemName"", I.""U_VSubGroup"", I.""U_VSubSubGroup"" " &
                    '               "FROM ""OITM"" I " &
                    '               "WHERE I.""ItmsGrpCod"" = '" & Itemgrp & "'"
                    '                         End If

                    '                         Dim ors1 As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                    '                         ors1.DoQuery(rs1)

                    '                         ' Clear matrix rows if any
                    '                         If objMatrix.RowCount > 0 Then
                    '                             For i As Integer = objMatrix.RowCount To 1 Step -1
                    '                                 objMatrix.DeleteRow(i)
                    '                             Next
                    '                         End If

                    '                         ' Load items into matrix
                    '                         If ors1.RecordCount > 0 Then
                    '                             For j As Integer = 0 To ors1.RecordCount - 1
                    '                                 objMatrix.AddRow()
                    '                                 Dim currentRow As Integer = objMatrix.VisualRowCount
                    '                                 oDS2.SetValue("LineId", oDS2.Offset, currentRow.ToString())
                    '                                 Dim itmCode As String = ors1.Fields.Item("ItemCode").Value.ToString()
                    '                                 oDS2.SetValue("U_ITMCODE", oDS2.Offset, itmCode)
                    '                                 oDS2.SetValue("U_ITMNM", oDS2.Offset, ors1.Fields.Item("ItemName").Value.ToString())


                    '                                 Dim GetBarCode As String =
                    ' "SELECT TOP 1 T0.""UomEntry"", T0.""BcdCode"", T1.""UomCode"" " &
                    ' "FROM ""OBCD"" T0 " &
                    ' "INNER JOIN ""OUOM"" T1 ON T0.""UomEntry"" = T1.""UomEntry"" " &
                    ' "WHERE T0.""ItemCode"" = '" & itmCode & "' " &
                    ' "ORDER BY T0.""BcdEntry"" ASC"

                    '                                 ' Get barcode and UOM
                    '                                 ' Dim GetBarCode As String = "SELECT TOP 1 ""UomEntry"", ""BcdCode"" FROM ""OBCD"" WHERE ""ItemCode"" = '" & itmCode & "' ORDER BY ""BcdEntry"" ASC"
                    '                                 Dim oRsGetBarCode As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                    '                                 oRsGetBarCode.DoQuery(GetBarCode)

                    '                                 If oRsGetBarCode.RecordCount > 0 Then
                    '                                     oDS2.SetValue("Code", oDS2.Offset, oRsGetBarCode.Fields.Item("BcdCode").Value.ToString())
                    '                                     oDS2.SetValue("U_UOM", oDS2.Offset, oRsGetBarCode.Fields.Item("UomCode").Value.ToString())
                    '                                 End If

                    '                                 objMatrix.SetLineData(currentRow)
                    '                                 ors1.MoveNext()
                    '                             Next
                    '                         Else
                    '                             ' If no records found, add a new blank line
                    '                             Me.SetNewLine(objForm.UniqueID, "Mtx")
                    '                         End If

                    '                         objMatrix.AutoResizeColumns()
                    '                         objMatrix.FlushToDataSource()

                    '                     Catch ex As Exception
                    '                         objForm.StatusBar.SetText("Error: " & ex.Message & " - Page Load Events",
                    '                               SAPbouiCOM.BoMessageTime.bmt_Long,
                    '                               SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '                     Finally
                    '                         objForm.Freeze(False)
                    '                     End Try

                    '                 End If





                    If pVal.ItemUID = "LOAD" AndAlso pVal.BeforeAction = False AndAlso
(objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or
objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or
objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                        Try

                            objForm.Freeze(True)

                            Dim oDS As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Add("@TNX_BGLEASTITMS")

                            Dim oDS2 As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Add("@TNX_BGLEASTITMS_C0")

                            Dim objMatrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx").Specific, SAPbouiCOM.Matrix)

                            Dim Itemgrp As String = objForm.Items.Item("IGroup").Specific.Value

                            Dim subgrp As String = objForm.Items.Item("SubGrp").Specific.Value

                            Dim subsubgrp As String = objForm.Items.Item("SSbGrp").Specific.Value

                            Dim Branch As String = objForm.Items.Item("BPLID").Specific.Value

                            ' Build Query

                            Dim rs1 As String = ""
                            If Not String.IsNullOrEmpty(subsubgrp) Then
                                rs1 = "SELECT ""ItemCode"", ""ItemName"" FROM OITM WHERE ""U_VSubSubGroup"" = '" & subsubgrp & "'"
                            ElseIf Not String.IsNullOrEmpty(subgrp) Then
                                rs1 = "SELECT ""ItemCode"", ""ItemName"" FROM OITM WHERE ""U_VSubGroup"" = '" & subgrp & "'"
                            ElseIf Not String.IsNullOrEmpty(Itemgrp) Then
                                rs1 = "SELECT ""ItemCode"", ""ItemName"" FROM OITM WHERE ""ItmsGrpCod"" = '" & Itemgrp & "'"

                            End If

                            Dim ors1 As SAPbobsCOM.Recordset =
         objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            ors1.DoQuery(rs1)

                            ' Clear Matrix

                            If objMatrix.RowCount > 0 Then

                                For i As Integer = objMatrix.RowCount To 1 Step -1

                                    objMatrix.DeleteRow(i)

                                Next

                            End If

                            ' Load Rows

                            If ors1.RecordCount > 0 Then

                                ors1.MoveFirst()

                                For j As Integer = 0 To ors1.RecordCount - 1

                                    objMatrix.AddRow()

                                    Dim row As Integer = objMatrix.VisualRowCount

                                    Dim ItemCode As String = ors1.Fields.Item("ItemCode").Value.ToString()

                                    oDS2.SetValue("LineId", oDS2.Offset, row)

                                    oDS2.SetValue("U_ITMCODE", oDS2.Offset, ItemCode)

                                    oDS2.SetValue("U_ITMNM", oDS2.Offset, ors1.Fields.Item("ItemName").Value.ToString())

                                    '-----------------------------

                                    ' GET BARCODE / UOM

                                    '-----------------------------

                                    Dim bcQuery As String =
                 "SELECT TOP 1 T0.""BcdCode"", T1.""UomCode"" " &
                 "FROM OBCD T0 INNER JOIN OUOM T1 ON T0.""UomEntry"" = T1.""UomEntry"" " &
                 "WHERE T0.""ItemCode"" = '" & ItemCode & "' "



                                    '    Dim bcQuery As String = " SELECT Top 1 ""BcdCode"" FROM OBCD WHERE  " &

                                    '"  ""ItemCode"" = '" & ItemCode & "' ORDER BY ""BcdEntry"" ASC "


                                    Dim rsBC As SAPbobsCOM.Recordset =
                 objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    rsBC.DoQuery(bcQuery)

                                    Dim Barcode As String = ""

                                    If rsBC.RecordCount > 0 Then

                                        Barcode = rsBC.Fields.Item("BcdCode").Value.ToString()

                                        oDS2.SetValue("Code", oDS2.Offset, Barcode)

                                        oDS2.SetValue("U_UOM", oDS2.Offset, rsBC.Fields.Item("UomCode").Value.ToString())

                                    End If

                                    '-----------------------------

                                    ' PROCEDURE 1 → GetItemPrice

                                    ' Saves U_ITCST

                                    '-----------------------------

                                    Try

                                        Dim q1 As String = "CALL ""GetItemPrice""('" & ItemCode & "','" & Branch.Trim & "')"

                                        Dim rsCost As SAPbobsCOM.Recordset =
                     objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        rsCost.DoQuery(q1)
                                        If rsCost.RecordCount > 0 Then
                                            oDS2.SetValue("U_ITCST", oDS2.Offset, rsCost.Fields.Item("AvgPrice").Value.ToString())

                                        Else

                                            oDS2.SetValue("U_ITCST", oDS2.Offset, "0")

                                        End If

                                    Catch ex As Exception

                                    End Try

                                    '-----------------------------

                                    ' PROCEDURE 2 → GetRetailPrice

                                    ' Saves U_PRICE

                                    '-----------------------------

                                    Try

                                        Dim q2 As String = "CALL ""GetRetailPrice""('" & Barcode & "','" & Branch.Trim & "')"
                                        Dim rsRetail As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        rsRetail.DoQuery(q2)
                                        If rsRetail.RecordCount > 0 Then
                                            oDS2.SetValue("U_PRICE", oDS2.Offset, rsRetail.Fields.Item("Price").Value.ToString())
                                        Else
                                            oDS2.SetValue("U_PRICE", oDS2.Offset, "0")

                                        End If

                                    Catch ex As Exception

                                    End Try

                                    objMatrix.SetLineData(row)
                                    ors1.MoveNext()

                                Next

                            Else

                                Me.SetNewLine(objForm.UniqueID, "Mtx")

                            End If

                            objMatrix.AutoResizeColumns()

                            objMatrix.FlushToDataSource()

                        Catch ex As Exception
                            objForm.StatusBar.SetText("Error: " & ex.Message,
                               SAPbouiCOM.BoMessageTime.bmt_Long,
                               SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Finally
                            objForm.Freeze(False)
                        End Try

                    End If     'BUY AND GET LEAST ITEM 








                    If pVal.ItemUID = "Browse" And pVal.BeforeAction = False Then
                        Me.BrowseButton_Click(objForm.UniqueID)
                    End If

                    '       If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And
                    '(pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.ActionSuccess = True And
                                 (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then

                        'If pVal.ItemUID = "1" AndAlso pVal.BeforeAction = True AndAlso (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        'Me.GetNextDocNum()
                        'Me.SetDefault(objForm.UniqueID)


                        'Dim FromTime As String = CType(objForm.Items.Item("FRMTIME").Specific, SAPbouiCOM.EditText).Value
                        'Dim ToTime As String = CType(objForm.Items.Item("TOTIME").Specific, SAPbouiCOM.EditText).Value
                        'Dim fromDate As Date
                        'Dim toDate As Date
                        'If Not DateTime.TryParseExact(CType(objForm.Items.Item("FRMDATE").Specific, SAPbouiCOM.EditText).Value.Trim(), "yyyyMMdd", Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, fromDate) Then
                        '    objMain.objApplication.StatusBar.SetText("please select 'From Date'. Please use format yyyyMMdd.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If
                        'If Not DateTime.TryParseExact(CType(objForm.Items.Item("TODATE").Specific, SAPbouiCOM.EditText).Value.Trim(), "yyyyMMdd", Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, toDate) Then
                        '    objMain.objApplication.StatusBar.SetText("Please select 'To Date'. Please use format yyyyMMdd.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If


                        ''Get matrix And loop through item rows
                        'Dim matrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx").Specific, SAPbouiCOM.Matrix)
                        'For i As Integer = 1 To matrix.RowCount
                        '    Dim itemCode As String = CType(matrix.Columns.Item("ITMCODE").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value
                        '    If itemCode <> "" Then
                        '        If IsItemInOverlappingPromotion(fromDate, toDate, FromTime, ToTime, itemCode) Then
                        '            objMain.objApplication.StatusBar.SetText("Row " & i.ToString() & ": Item " & itemCode & " is already in another active promotion with overlapping dates.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '            BubbleEvent = False
                        '            Exit Sub
                        '        End If
                        '    End If
                        'Next
                    End If
                    'Duplicate barcode
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                                  (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                        Dim projtype As String = objForm.Items.Item("PRTYP").Specific.Value
                        If Me.Validate(objForm.UniqueID, projtype) = False Then
                            LostFocusFlag = False
                            BubbleEvent = False
                        End If
                    End If
                    'coment

                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    '    If Me.ValidatePromotionName(objForm.UniqueID) = False Then
                    '        BubbleEvent = False
                    '    End If

                    'End If






                    '----promo---------
                    '       If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then


                    '           Me.GetNextDocNum()
                    '           Me.SetDefault(objForm.UniqueID)
                    '           'Me.SetNewLine(objForm.UniqueID, "Mtx")


                    '           Dim FromTime As String = CType(objForm.Items.Item("FRMTIME").Specific, SAPbouiCOM.EditText).Value
                    '           Dim ToTime As String = CType(objForm.Items.Item("TOTIME").Specific, SAPbouiCOM.EditText).Value
                    '           Dim fromDate As Date
                    '           Dim toDate As Date
                    '           If Not DateTime.TryParseExact(CType(objForm.Items.Item("FRMDATE").Specific, SAPbouiCOM.EditText).Value.Trim(), "yyyyMMdd", Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, fromDate) Then
                    '               objMain.objApplication.StatusBar.SetText("please select 'From Date'. Please use format yyyyMMdd.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '               BubbleEvent = False
                    '               Exit Sub
                    '           End If
                    '           If Not DateTime.TryParseExact(CType(objForm.Items.Item("TODATE").Specific, SAPbouiCOM.EditText).Value.Trim(), "yyyyMMdd", Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, toDate) Then
                    '               objMain.objApplication.StatusBar.SetText("Please select 'To Date'. Please use format yyyyMMdd.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '               BubbleEvent = False
                    '               Exit Sub
                    '           End If


                    '           'Get matrix And loop through item rows
                    '           Dim matrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx").Specific, SAPbouiCOM.Matrix)
                    '           For i As Integer = 1 To matrix.RowCount
                    '               Dim itemCode As String = CType(matrix.Columns.Item("ITMCODE").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value
                    '               If itemCode <> "" Then
                    '                   If IsItemInOverlappingPromotion(fromDate, toDate, FromTime, ToTime, itemCode) Then
                    '                       objMain.objApplication.StatusBar.SetText("Row " & i.ToString() & ": Item " & itemCode & " is already in another active promotion with overlapping dates.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '                       BubbleEvent = False
                    '                       Exit Sub
                    '                   End If
                    '               End If
                    '           Next
                    '           If Me.ValidatePromotionName(objForm.UniqueID) = False Then
                    '               BubbleEvent = False
                    '           End If
                    '       End If
                    '       ' coment
                    '       If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And
                    '(pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then

                    '           Me.GetNextDocNum()

                    '           Me.SetDefault(objForm.UniqueID)
                    '       End If
                    '       '----------end promo-------------1---

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And
             (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        Me.SetDefault(objForm.UniqueID)
                        ' Me.GetNextDocNum()
                        'Me.BranchLoad(objForm.UniqueID)
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                    (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                        'If pVal.ItemUID = "1" AndAlso pVal.BeforeAction = True AndAlso (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Dim FromTime As String = CType(objForm.Items.Item("FRMTIME").Specific, SAPbouiCOM.EditText).Value
                        Dim ToTime As String = CType(objForm.Items.Item("TOTIME").Specific, SAPbouiCOM.EditText).Value
                        Dim fromDate As Date
                        Dim toDate As Date
                        If Not DateTime.TryParseExact(CType(objForm.Items.Item("FRMDATE").Specific, SAPbouiCOM.EditText).Value.Trim(), "yyyyMMdd", Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, fromDate) Then
                            objMain.objApplication.StatusBar.SetText("Please select 'From Date'. Please use format yyyyMMdd.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If Not DateTime.TryParseExact(CType(objForm.Items.Item("TODATE").Specific, SAPbouiCOM.EditText).Value.Trim(), "yyyyMMdd", Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, toDate) Then
                            objMain.objApplication.StatusBar.SetText("Please select 'To Date'. Please use format yyyyMMdd.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If

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

                        'Akhila 
                        'Max qty and max qty get item hided-------------------------
                        'If CDbl(oDS.GetValue("U_MXQBQTY", 0)) = 0 Then
                        '    objMain.objApplication.StatusBar.SetText("Please Enter Maximum Quantity Buy Item", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If

                        'If CDbl(oDS.GetValue("U_MXQGQTY", 0)) = 0 Then
                        '    objMain.objApplication.StatusBar.SetText("Please Enter Maximum Quantity Get Item", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If

                        If CDbl(oDS.GetValue("U_BUYQTY", 0)) = 0 Then
                            objMain.objApplication.StatusBar.SetText("Please Enter Buy Quantity", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If


                        If CDbl(oDS.GetValue("U_GTQTY", 0)) = 0 Then
                            objMain.objApplication.StatusBar.SetText("Please Enter Get Quantity", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If


                        If oDS.GetValue("U_BPLID", 0) = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Select Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If



                        If oDS.GetValue("U_BPLID", 0) = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Select Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If




                        'If oDS.GetValue("U_MXQBQTY", 0) = "" Then
                        '    objMain.objApplication.StatusBar.SetText("Please Enter Maximum Quantity Buy Item", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If
                        'If oDS.GetValue("U_MXQGQTY", 0) = "" Then
                        '    objMain.objApplication.StatusBar.SetText("Please Enter Maximum Quantity Get Item", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If
                        'If oDS.GetValue("U_PRMNM", 0) = "" Then
                        '    objMain.objApplication.StatusBar.SetText("Please Enter Promotion Name", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        'End If

                        'If oDS1.GetValue("U_DPERCTG", 0) = "" Then
                        '    objMain.objApplication.StatusBar.SetText("Please Enter the Discount Percentage", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        'End If

                        'For i As Integer = 1 To objMatrix.VisualRowCount
                        '    objMatrix.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                        '    If objMatrix.Columns.Item("BCODE").Cells.Item(i).Specific.Value.ToString().Trim() = "" Then
                        '        'objMain.objApplication.StatusBar.SetText("Please Choose Different selection And press Load Button", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '        objMain.objApplication.StatusBar.SetText("Please Choose  Bar Code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '        BubbleEvent = False
                        '        Exit Sub
                        '    End If
                        'Next


                    End If


                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If Me.ValidatePromotionName(objForm.UniqueID) = False Then
                            BubbleEvent = False
                        End If
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If Me.Branchloadnew(objForm.UniqueID) = False Then
                            BubbleEvent = False
                        End If

                    End If




                    'Akhila Overlapping code --

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
                    '                         Dim oldStatusQuery As String = "SELECT ""U_STATUS"" FROM ""@TNX_BGLEASTITMS"" WHERE ""DocNum"" = '" & docnum & "'"
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
                    '                         Dim matrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx").Specific, SAPbouiCOM.Matrix)
                    '                         For i As Integer = 1 To matrix.VisualRowCount
                    '                             matrix.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    '                             Dim Barcode As String = CType(matrix.Columns.Item("BCODE").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value.Trim()
                    '                             Dim UOM As String = CType(matrix.Columns.Item("UOM").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value.Trim()
                    '                             Dim itemCode As String = CType(matrix.Columns.Item("ITMCODE").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value.Trim()
                    '                             If itemCode <> "" Then
                    '                                 Dim query As String = "SELECT H1.""DocNum"" FROM ""@TNX_BGLEASTITMS"" H1 " &
                    '                 "INNER JOIN ""@TNX_BGLEASTITMS_C0"" C1 ON H1.""DocEntry"" = C1.""DocEntry"" " &
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





                    'other validations-------
                    '                 If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                    '(pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                    '                     Dim matrix11 As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx").Specific, SAPbouiCOM.Matrix)

                    '                     For i As Integer = 1 To matrix11.VisualRowCount
                    '                         'matrix11.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                    '                         Dim itemCode1 As String = CType(matrix11.Columns.Item("ITMCODE").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value.Trim()
                    '                         If String.IsNullOrEmpty(itemCode1) Then Continue For

                    '                         Dim fromTime1 As String = CType(objForm.Items.Item("FRMTIME").Specific, SAPbouiCOM.EditText).Value.Trim()
                    '                         Dim toTime1 As String = CType(objForm.Items.Item("TOTIME").Specific, SAPbouiCOM.EditText).Value.Trim()
                    '                         Dim fromDate1 As String = objForm.Items.Item("FRMDATE").Specific.Value.Trim()
                    '                         Dim toDate1 As String = objForm.Items.Item("TODATE").Specific.Value.Trim()
                    '                         Dim currentDocNum As String = objForm.Items.Item("DocNum").Specific.Value.Trim()


                    '                         Dim query As String =
                    ' "WITH ItemExistence AS ( " &
                    ' "    SELECT 'Item wise Discount' AS TableName,a.""DocNum"" as DocNum " &
                    ' "    FROM ""@TNX_ITEMDISC"" a " &
                    ' "    JOIN ""@TNX_ITEMDISC_C0"" b ON a.""DocEntry"" = b.""DocEntry"" " &
                    ' "    WHERE b.""U_ITMCODE"" = '" & itemCode1 & "' " &
                    ' "      AND a.""U_STATUS"" = 'Active' " &
                    ' "      AND a.""U_FRMDATE"" <= '" & fromDate1 & "' " &
                    ' "      AND a.""U_TODATE"" >= '" & toDate1 & "' " &
                    ' "      AND '" & toTime1 & "' <= a.""U_TOTIME"" " &
                    ' "      AND '" & fromTime1 & "' >= a.""U_FRMTIME"" " &
                    ' "    UNION " &
                    ' "    SELECT 'Promotion on bill total' AS TableName ,a.""DocNum"" as DocNum " &
                    ' "    FROM ""@TNX_PBILL"" a " &
                    ' "    JOIN ""@TNX_PBILL_C0"" b ON a.""DocEntry"" = b.""DocEntry"" " &
                    ' "    WHERE b.""U_ITMCODE"" = '" & itemCode1 & "' " &
                    ' "      AND a.""U_STATUS"" = 'Active' " &
                    ' "      AND a.""U_FRMDATE"" <= '" & fromDate1 & "' " &
                    ' "      AND a.""U_TODATE"" >= '" & toDate1 & "' " &
                    ' "      AND '" & toTime1 & "' <= a.""U_TOTIME"" " &
                    ' "      AND '" & fromTime1 & "' >= a.""U_FRMTIME"" " &
                    ' "    UNION " &
                    ' "    SELECT 'Promotion on voucher purchase' AS TableName ,a.""DocNum"" as DocNum " &
                    ' "    FROM ""@TNX_PONVP"" a " &
                    ' "    JOIN ""@TNX_PONVP_C0"" b ON a.""DocEntry"" = b.""DocEntry"" " &
                    ' "    WHERE b.""U_VITMCDE"" ='" & itemCode1 & "' " &
                    ' "      AND a.""U_STATUS"" = 'Active' " &
                    ' "      AND a.""U_FRMDATE"" <= '" & fromDate1 & "' " &
                    ' "      AND a.""U_TODATE"" >= '" & toDate1 & "' " &
                    ' "      AND'" & toTime1 & "' <= a.""U_TOTIME"" " &
                    ' "      AND '" & fromTime1 & "' >= a.""U_FRMTIME"" " &
                    ' ") " &
                    ' "SELECT TableName, DocNum FROM ItemExistence " &
                    ' "UNION ALL " &
                    ' "SELECT 'No' AS TableName,'0' As DocNum FROM DUMMY " &
                    ' "WHERE NOT EXISTS (SELECT 1 FROM ItemExistence);"

                    '                         Dim oRecordset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                         oRecordset.DoQuery(query)
                    '                         If oRecordset.RecordCount > 0 Then
                    '                             Dim conflictTable As String = oRecordset.Fields.Item("TableName").Value.ToString()
                    '                             Dim conflictDocNum As String = oRecordset.Fields.Item("DocNum").Value.ToString()
                    '                             If conflictTable <> "No" Then
                    '                                 'objMain.objApplication.StatusBar.SetText("Item Code '" & itemCode1 & " is already linked to an active promotion with overlapping period in the screen " & conflictTable & " Duplicate promotion entries are not allowed.'",
                    '                                 '                                                     SAPbouiCOM.BoMessageTime.bmt_Short,
                    '                                 '                                                     SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    '                                 objMain.objApplication.StatusBar.SetText(
                    ' "Item Code '" & itemCode1 & "' is already linked to an active promotion with overlapping period in the screen " & conflictTable &
                    ' " (DocNum: " & conflictDocNum & "). Duplicate promotion entries are not allowed.",
                    ' SAPbouiCOM.BoMessageTime.bmt_Short,
                    ' SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    '                                 BubbleEvent = False
                    '                                 Exit Sub
                    '                             End If

                    '                         End If
                    '                     Next

                    '                 End If







                    'Akhila
                    '-----------------upload-----1
                    If pVal.ItemUID = "Upload" And pVal.BeforeAction = False Then
                        ' 🔽 Get file path from the edit text field
                        Dim filePath As String = CType(objForm.Items.Item("Path").Specific, SAPbouiCOM.EditText).Value

                        If System.IO.File.Exists(filePath) Then
                            Me.LoadExcelToMatrix(objForm, filePath)
                        Else
                            objMain.objApplication.MessageBox("File not found or path is invalid.")
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.BeforeAction = False
                    Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(pVal.FormUID)
                    'If pVal.ItemUID = "PRTYP" Then
                    '    'Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(pVal.FormUID)
                    '    Dim objComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("PRTYP").Specific

                    '    If objComboBox.Value = "I" Then
                    '        objForm.Items.Item("IGroup").Enabled = True
                    '        objForm.Items.Item("SSbGrp").Enabled = True
                    '        objForm.Items.Item("SubGrp").Enabled = True
                    '    End If
                    'End If

                    'If pVal.BeforeAction = False AndAlso pVal.ItemUID = "PRTYP" Then
                    '    Dim objComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("PRTYP").Specific

                    '    If objComboBox.Value = "A" Then
                    '        objForm.Items.Item("IGroup").Enabled = False
                    '        objForm.Items.Item("SSbGrp").Enabled = False
                    '        objForm.Items.Item("SubGrp").Enabled = False
                    '    End If
                    'End If


                    If pVal.BeforeAction = False AndAlso pVal.ItemUID = "PRTYP" Then

                        Dim objCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("PRTYP").Specific
                        Dim promoType As String = objCombo.Value.Trim()
                        Dim cmbItemGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("IGroup").Specific
                        Dim cmbSubGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("SubGrp").Specific
                        Dim cmbSubSubGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("SSbGrp").Specific
                        If promoType = "A" Then
                            objForm.Items.Item("IGroup").Enabled = False
                            objForm.Items.Item("SubGrp").Enabled = False
                            objForm.Items.Item("SSbGrp").Enabled = False
                            cmbItemGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            cmbSubGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            cmbSubSubGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        ElseIf promoType = "I" Then
                            objForm.Items.Item("IGroup").Enabled = True
                            objForm.Items.Item("SubGrp").Enabled = True
                            objForm.Items.Item("SSbGrp").Enabled = True
                            cmbItemGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            cmbSubGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            cmbSubSubGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        End If
                    End If


                    'new series add-------
                    If pVal.ItemUID = "23" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("23").Specific.Selected.Value <> "" Then
                            Dim series As Integer = objForm.Items.Item("23").Specific.Selected.Value
                            SetDefault(objForm.UniqueID, "CobmoSelect", series)
                        End If
                    End If



                    'end new series-------



                    ' objComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                    'If pVal.ItemUID = "IGroup" Then
                    '    Dim cmb As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("IGroup").Specific, SAPbouiCOM.ComboBox)
                    '    Dim ItemGroup As String = cmb.Selected.Value ' Get the selected item group code
                    '    Dim rs As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                    '    Dim query As String = "SELECT T0.""U_VSubGroup"", T0.""U_VSubSubGroup"" " &
                    '                          "FROM OITM T0 " &
                    '                          "INNER JOIN OITB T1 ON T0.""ItmsGrpCod"" = T1.""ItmsGrpCod"" " &
                    '                          "WHERE T0.""ItmsGrpCod"" = '" & ItemGroup & "'"

                    '    rs.DoQuery(query)

                    '    If Not rs.EoF Then
                    '        objForm.Items.Item("SubGrp").Specific.Value = rs.Fields.Item("U_VSubGroup").Value.ToString()
                    '        objForm.Items.Item("SSbGrp").Specific.Value = rs.Fields.Item("U_VSubSubGroup").Value.ToString()
                    '    End If
                    'End If

                    'If pVal.ItemUID = "IGroup" Then
                    '    Dim cmb As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("IGroup").Specific, SAPbouiCOM.ComboBox)
                    '    If cmb.Selected IsNot Nothing Then
                    '        Dim itemGroup As String = cmb.Selected.Value
                    '        LoadSubGroups1(objForm, itemGroup)
                    '    End If
                    'End If

                    'If pVal.ItemUID = "SubGrp" Then
                    '    Dim cmbGrp As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("IGroup").Specific, SAPbouiCOM.ComboBox)
                    '    Dim cmbSub As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("SubGrp").Specific, SAPbouiCOM.ComboBox)
                    '    If cmbGrp.Selected IsNot Nothing AndAlso cmbSub.Selected IsNot Nothing Then
                    '        Dim itemGroup As String = cmbGrp.Selected.Value
                    '        Dim subGroup As String = cmbSub.Selected.Value
                    '        LoadSubSubGroups(objForm, itemGroup, subGroup)
                    '    End If
                    'End If


                    If objForm.Items.Item("IGroup").Specific.Value <> "" Then
                        Dim objComboBoxItemGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("IGroup").Specific
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
                            Dim objComboBoxItemGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("IGroup").Specific
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

                            Dim objComboBoxSubSubGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("SSbGrp").Specific
                            objMain.objUtilities.ComboBoxLoadValues(objComboBoxSubSubGroup, sql)
                            objComboBoxSubSubGroup.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                            objForm.Items.Item("SSbGrp").DisplayDesc = True
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
            objForm = objMain.objApplication.Forms.GetForm("Clear", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        If objForm.Items.Item("17").Specific.Value.Trim = "Close" Then
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                        Else
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                        End If
                    End If

            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
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

    'Sub SetDefault(ByVal FormUID As String)
    Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "", Optional ByVal Series As Integer = 0)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            oDS = Me.objForm.DataSources.DBDataSources.Add("@TNX_BGLEASTITMS")
            oDS1 = Me.objForm.DataSources.DBDataSources.Add("@TNX_BGLEASTITMS_C0")
            oDS2 = Me.objForm.DataSources.DBDataSources.Add("@TNX_BGLEASTITMS_C1")
            objMatrix = objForm.Items.Item("Mtx").Specific
            objMatrix1 = objForm.Items.Item("IGMtx1").Specific

            '--Akhila----------------------
            'oDS.SetValue("DocNum", oDS.Offset, objMain.objUtilities.GetNextDocNum(objForm, "ME_LEASTITMUDO", "Primary"))
            'objForm.DataBrowser.BrowseBy = "DocNum"
            oDS.SetValue("U_STATUS", oDS.Offset, "Active")
            'oDS.SetValue("U_DOCDATE", oDS.Offset, DateTime.Now.ToString("yyyyMMdd"))

            'Series code 
            Dim cmbSeries As SAPbouiCOM.ComboBox = objForm.Items.Item("23").Specific
            Dim seriesQuery As String = "SELECT ""Series"", ""SeriesName"" FROM NNM1 WHERE ""ObjectCode"" = 'ME_LEASTITMUDO' AND ""Locked"" = 'N'"
            objMain.objUtilities.ComboBoxLoadValuesSeries(cmbSeries, seriesQuery)
            objForm.Items.Item("23").DisplayDesc = True
            cmbSeries.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            ' Get default series
            '            Dim defaultSeriesQuery As String = "
            'SELECT T1.""Series"", T1.""SeriesName""
            'FROM ONNM T0 
            'JOIN NNM1 T1 ON T1.""ObjectCode"" = T0.""ObjectCode"" AND T1.""Series"" = T0.""DfltSeries"" 
            'WHERE T0.""ObjectCode"" IN ('ME_LEASTITMUDO','@TNX_BGLEASTITMS')"


            Dim defaultSeriesQuery As String =
    "SELECT ""Series"", ""SeriesName"", ""ObjectCode"", ""Indicator"" " &
    "FROM NNM1 " &
    "WHERE ""ObjectCode"" IN ('ME_LEASTITMUDO', '@TNX_BGLEASTITMS') " &
    "AND ""Indicator"" <> 'Default'"




            Dim oRsDflt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsDflt.DoQuery(defaultSeriesQuery)

            If Flag = "" Then
                Dim nextDocNum = objMain.objUtilities.GetNextDocNumSeries(objForm, "ME_LEASTITMUDO", oRsDflt.Fields.Item("Series").Value)
                oDS.SetValue("DocNum", oDS.Offset, nextDocNum)
                oDS.SetValue("U_BYSER", oDS.Offset, oRsDflt.Fields.Item("Series").Value)
            Else
                Dim nextDocNum = objMain.objUtilities.GetNextDocNumSeries(objForm, "ME_LEASTITMUDO", Series)
                oDS.SetValue("DocNum", oDS.Offset, nextDocNum)
                oDS.SetValue("U_BYSER", oDS.Offset, Series)
            End If

            oDS.SetValue("U_DOCDATE", oDS.Offset, DateTime.Now.ToString("yyyyMMdd"))


            Me.BranchLoad(objForm.UniqueID)



            'end series

            'End series






            oDS.SetValue("U_PRTYP", oDS.Offset, "")
            objMatrix.Clear()
            oDS1.Clear()
            objMatrix.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "Mtx")
            objMatrix.AutoResizeColumns()
            objMatrix1.Clear()
            oDS2.Clear()
            objMatrix1.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "IGMtx1")
            Me.SetNewLine(objForm.UniqueID, "EMtx")
            Me.SetNewLine(objForm.UniqueID, "BMtx")
            objMatrix1.AutoResizeColumns()
            objForm.Freeze(False)



        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal MatrixUID As String)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oDS = objForm.DataSources.DBDataSources.Item("@TNX_BGLEASTITMS")
        oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_BGLEASTITMS_C0")
        oDS2 = Me.objForm.DataSources.DBDataSources.Add("@TNX_BGLEASTITMS_C1")
        oDS3 = Me.objForm.DataSources.DBDataSources.Add("@TNX_BGLEASTITMS_C3")
        oDS4 = Me.objForm.DataSources.DBDataSources.Add("@TNX_BGLEASTITMS_C2")

        objMatrix = objForm.Items.Item("Mtx").Specific
        objMatrix1 = objForm.Items.Item("IGMtx1").Specific
        objMatrix2 = objForm.Items.Item("EMtx").Specific
        objMatrix3 = objForm.Items.Item("BMtx").Specific

        Select Case MatrixUID
            Case "Mtx"
                objMatrix.AddRow()
                oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
                oDS1.SetValue("U_BCODE", oDS1.Offset, "")
                oDS1.SetValue("U_ITMCODE", oDS1.Offset, "")
                oDS1.SetValue("U_ITMNM", oDS1.Offset, "")
                oDS1.SetValue("U_UOM", oDS1.Offset, "")
                oDS1.SetValue("U_PRICE", oDS1.Offset, "")
                oDS1.SetValue("U_ITCST", oDS1.Offset, "")

                objMatrix.AutoResizeColumns()
                objMatrix.SetLineData(objMatrix.VisualRowCount)
            Case "IGMtx1"
                objMatrix1.AddRow()
                oDS2.SetValue("LineId", oDS2.Offset, objMatrix1.VisualRowCount)
                ' oDS2.SetValue("Code", oDS2.Offset, "")
                oDS2.SetValue("U_ITMCDE", oDS2.Offset, "")
                oDS2.SetValue("U_ITMGRP", oDS2.Offset, "")

                'oDS2.SetValue("U_ITMNM", oDS2.Offset, "")
                'oDS2.SetValue("U_UOM", oDS2.Offset, "")
                'oDS2.SetValue("U_PRICE", oDS2.Offset, "")
                objMatrix1.SetLineData(objMatrix1.VisualRowCount)
                objMatrix1.AutoResizeColumns()
                objMatrix1.SetLineData(objMatrix1.VisualRowCount)
            Case "EMtx"
                objMatrix2.AddRow()
                oDS3.SetValue("LineID", oDS3.Offset, objMatrix2.VisualRowCount)
                oDS3.SetValue("U_BRCODE", oDS3.Offset, "")
                oDS3.SetValue("U_ITCODE", oDS3.Offset, "")
                oDS3.SetValue("U_ITNM", oDS3.Offset, "")
                objMatrix2.SetLineData(objMatrix2.VisualRowCount)
                objMatrix2.AutoResizeColumns()
                objMatrix2.SetLineData(objMatrix2.VisualRowCount)
                'Case "BMtx"
                '    objMatrix3.AddRow()
                '    oDS4.SetValue("LINEID", oDS4.Offset, objMatrix3.VisualRowCount)
                '    oDS4.SetValue("U_CHEK", oDS4.Offset, "")
                '    oDS4.SetValue("U_BCDE", oDS4.Offset, "")
                '    oDS4.SetValue("U_BRNME", oDS4.Offset, "")
                '    objMatrix3.SetLineData(objMatrix3.VisualRowCount)
                '    objMatrix3.AutoResizeColumns()
        End Select
    End Sub


    Sub BranchLoad(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix3 = objForm.Items.Item("BMtx").Specific
            oDS3 = Me.objForm.DataSources.DBDataSources.Add("@TNX_BGLEASTITMS_C2")
            Dim branch As String = "SELECT T0.""BPLId"", T0.""BPLName"" FROM OBPL T0"
            Dim oRgrp As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRgrp.DoQuery(branch)
            objMatrix3.Clear()
            objMatrix3.FlushToDataSource()
            If oRgrp.RecordCount > 0 Then
                For i As Integer = 1 To oRgrp.RecordCount
                    objMatrix3.AddRow()
                    oDS3.SetValue("LINEID", oDS3.Offset, i)
                    oDS3.SetValue("U_BCDE", oDS3.Offset, oRgrp.Fields.Item("BPLId").Value)
                    oDS3.SetValue("U_BRNME", oDS3.Offset, oRgrp.Fields.Item("BPLName").Value)
                    objMatrix3.SetLineData(i)
                    ' Me.SetNewLine(objForm.UniqueID, "BMtx")

                    oRgrp.MoveNext()
                Next
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
    End Sub


    Public Sub CreateProductionOrder(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDS = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR")
            oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific
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

    'Duplicate  barcode method

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
                Dim Qry As String = "select distinct T0.""Code"",T1.""U_FRMDATE"", T1.""U_TODATE"", T1.""DocNum"" " &
                                             "from ""@TNX_BGLEASTITMS_C0"" T0 inner join ""@TNX_BGLEASTITMS"" T1 on T0.""DocEntry""=T1.""DocEntry"" " &
                                          " where  (T1.""U_FRMDATE"" between '" & FromDate1 & "' and '" & ToDate1 & "' " &
                                         " and T0.""Code""<> ''   And T1.""U_STATUS""='Active' and T1.""DocNum"" <> '" & objForm.Items.Item("DocNum").Specific.Value & "' ) or " &
                                         "(T1.""U_TODATE"" between '" & FromDate1 & "' and '" & ToDate1 & "' and T0.""Code""<> '' And T1.""U_STATUS""='Active' and T1.""DocNum"" <> '" & objForm.Items.Item("DocNum").Specific.Value & "' )"
                Dim rs2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim chkBarCode1 As String
                Dim chkDocNum1 As String
                Dim chkFromDate1 As String
                Dim chkToDate1 As String
                rs2.DoQuery(Qry)
                If rs2.RecordCount > 0 Then
                    For i As Integer = 1 To rs2.RecordCount
                        chkBarCode1 = rs2.Fields.Item("Code").Value
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






            Return True

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function



    Sub CreateIssueForProduction(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDS = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR")
            oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific

            Dim result As Integer = objMain.objApplication.MessageBox("You are going to post Receipt from production. Do you want to continue?", 1, "Continue", "Cancel")
            If (result = 1) Then
                objMatrix = objForm.Items.Item("Mtx").Specific
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
            objMatrix = objForm.Items.Item("Mtx").Specific

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

                    End If
                End If
            End If
            ' System.Diagnostics.Process.Start()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDS = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR")
            oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific
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
    Private Sub BrowseButton_Click(ByVal FormUID As String)
        Dim filePath As String = ""

        ' Create an STA thread to show the file dialog
        Dim t As New Threading.Thread(Sub()
                                          Dim fileDialog As New OpenFileDialog()
                                          fileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                                          fileDialog.Filter = "All files (*.*)|*.*"
                                          fileDialog.Title = "Select a file"

                                          If fileDialog.ShowDialog() = DialogResult.OK Then
                                              filePath = fileDialog.FileName
                                          End If
                                      End Sub)

        t.SetApartmentState(Threading.ApartmentState.STA)
        t.Start()
        t.Join() ' Wait until the thread completes

        ' Set the file path to the edit text field
        If filePath <> "" Then
            Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUID)
            Dim oEditText As SAPbouiCOM.EditText = objForm.Items.Item("Path").Specific
            oEditText.Value = filePath
        End If
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
        rs.DoQuery("SELECT COUNT(*) AS Cnt FROM ""@TNX_BGLEASTITMS"" WHERE ""U_PRMNM""= '" & promoName.Replace("'", "''") & "'")
        If rs.Fields.Item("Cnt").Value > 0 Then
            ' objMain.objApplication.MessageBox("Promotion Name already exists. Please enter a unique name.")
            objMain.objApplication.StatusBar.SetText("Promotion Name already exists. Please enter a unique name", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        'If oDS.GetValue("U_FRMTIME", 0) = "" Then
        '    objMain.objApplication.StatusBar.SetText("Please Enter From Time", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '    Return False

        'End If
        'If oDS.GetValue("U_TOTIME", 0) = "" Then
        '    objMain.objApplication.StatusBar.SetText("Please Enter To Time", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '    Return False
        'End If


        For i As Integer = 1 To objMatrix.VisualRowCount
            If objMatrix.Columns.Item("BCODE").Cells.Item(1).Specific.Value.ToString().Trim() = "" Then
                'objMain.objApplication.StatusBar.SetText("Please Choose Different selection And press Load Button", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objMain.objApplication.StatusBar.SetText("Please Choose Barcode", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

        Next
        Return True
    End Function

    Function Branchloadnew(ByVal FormUiD As String)
        Try
            For j As Integer = 1 To objMatrix3.VisualRowCount
                Dim oCheckBox As SAPbouiCOM.CheckBox
                oCheckBox = objMatrix3.Columns.Item("CHEK").Cells.Item(j).Specific
                For k As Integer = 1 To objMatrix3.VisualRowCount
                    oCheckBox = objMatrix3.Columns.Item("CHEK").Cells.Item(k).Specific
                    If oCheckBox.Checked = True Then
                        Exit For
                    End If
                Next
                If oCheckBox.Checked = False Then
                    objMain.objApplication.StatusBar.SetText("Please Select Atleast one Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Next

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
















    'Private Sub LoadSubGroups(ByVal ItemGroup As String)
    '    Dim oCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("SubGrp").Specific
    '    oCombo.ValidValues.Clear()

    '    Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    Dim sql As String = "SELECT DISTINCT ""T0.U_VSubGroup"" FROM OITM T0 INNER JOIN OITB T1 ON T0.ItmsGrpCod = T1.ItmsGrpCod WHERE T1.ItmsGrpNam = T0." '& Itemgroup &'" AND T0.U_VSubGroup IS NOT NULL ORDER BY T0.U_VSubGroup"
    '    oRS.DoQuery(sql)

    '    While Not oRS.EoF
    '        oCombo.ValidValues.Add(oRS.Fields.Item(0).Value.ToString(), "")
    '        oRS.MoveNext()
    '    End While
    'End Sub

    Private Sub LoadSubGroups1(objForm As SAPbouiCOM.Form, itemGroup As String)

        If String.IsNullOrEmpty(itemGroup) OrElse Not IsNumeric(itemGroup.Trim()) Then

            'objMain.objApplication.MessageBox("Invalid or empty Item Group. Exiting.")
            Exit Sub
        End If

        Dim objCombo1 As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("SubGrp").Specific, SAPbouiCOM.ComboBox)

        ' === CLEAR EXISTING COMBO VALUES ===
        If objCombo1.ValidValues.Count > 0 Then
            For i As Integer = objCombo1.ValidValues.Count - 1 To 0 Step -1
                objCombo1.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
        End If

        ' === QUERY SUBGROUPS ===
        Dim oRS As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

        ' Build query with numeric-safe filter
        'Dim sql As String = "SELECT DISTINCT T0.""U_VSubGroup"" " &
        '                "FROM OITM T0 " &
        '                "INNER JOIN OITB T1 ON T0.""ItmsGrpCod"" = T1.""ItmsGrpCod"" " &
        '                "WHERE T0.""ItmsGrpCod"" = " & itemGroup & " " & ' No quotes for numeric value
        '                "AND T0.""U_VSubGroup"" IS NOT NULL " &
        '                "ORDER BY T0.""U_VSubGroup"""
        'oRS.DoQuery(sql)\
        Try
            objCombo1.ValidValues.Add("", "")
        Catch ex As Exception
        End Try

        Dim sql As String = "SELECT DISTINCT T0.""U_VSubGroup"" " &
                        "FROM OITM T0 " &
                        "WHERE T0.""ItmsGrpCod"" = " & itemGroup & " " & ' No quotes for numeric value
                        "AND T0.""U_VSubGroup"" IS NOT NULL " &
                        "ORDER BY T0.""U_VSubGroup"""

        oRS.DoQuery(sql)


        ' === ADD TO COMBO AND SELECT FIRST ===
        Dim firstValue As String = ""
        While Not oRS.EoF
            Dim subGroup As String = oRS.Fields.Item(0).Value.ToString()
            If Not String.IsNullOrEmpty(subGroup) Then
                objCombo1.ValidValues.Add(subGroup, subGroup)
                If firstValue = "" Then
                    firstValue = subGroup
                End If
            End If
            oRS.MoveNext()
        End While
        'If objCombo1.ValidValues.Count > 0 AndAlso Not String.IsNullOrEmpty(firstValue) Then
        '    objCombo1.Select(firstValue, SAPbouiCOM.BoSearchKey.psk_ByValue)
        'End If
    End Sub

    Private Sub LoadSubSubGroups(objForm As SAPbouiCOM.Form, itemGroup As String, subGroup As String)
        If String.IsNullOrEmpty(itemGroup) OrElse Not IsNumeric(itemGroup.Trim()) Then
            Exit Sub
        End If
        Dim objCombo As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("SSbGrp").Specific, SAPbouiCOM.ComboBox)
        If objCombo.ValidValues.Count > 0 Then
            For i As Integer = objCombo.ValidValues.Count - 1 To 0 Step -1
                objCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
        End If
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
        'If objCombo.ValidValues.Count > 0 AndAlso Not String.IsNullOrEmpty(firstValue) Then
        '    objCombo.Select(firstValue, SAPbouiCOM.BoSearchKey.psk_ByValue)
        'End If
    End Sub

    Private Sub LoadExcelToMatrix(objForm As SAPbouiCOM.Form, filePath As String)
        Try
            Dim oDataTable As SAPbouiCOM.DataTable

            ' Check and create DT_MATRIX if not exists
            If objForm.DataSources.DataTables.Count = 0 OrElse objForm.DataSources.DataTables.Item("DT_MATRIX") Is Nothing Then
                objForm.DataSources.DataTables.Add("DT_MATRIX")
            End If

            oDataTable = objForm.DataSources.DataTables.Item("DT_MATRIX")

            ' If columns are not already added, add them
            If oDataTable.Columns.Count = 0 Then
                oDataTable.Columns.Add("Barcode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                oDataTable.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                oDataTable.Columns.Add("ItemDescription", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 254)
                oDataTable.Columns.Add("UOM", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50)
                oDataTable.Columns.Add("Price", SAPbouiCOM.BoFieldsType.ft_Price, 20)
            End If

            ' Ensure DataTable exists
            If objForm.DataSources.DataTables.Count = 0 OrElse objForm.DataSources.DataTables.Item("DT_MATRIX") Is Nothing Then
                objForm.DataSources.DataTables.Add("DT_MATRIX")
            End If

            Dim objMatrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx").Specific, SAPbouiCOM.Matrix)
            Dim oDataSource As SAPbouiCOM.DataTable = objForm.DataSources.DataTables.Item("DT_MATRIX")

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
                oDataTable.SetValue("PRICE", i, dt.Rows(i)("Price").ToString())
            Next
            ' Load into matrix
            objMatrix.LoadFromDataSource()
            'Dim objMatrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx").Specific, SAPbouiCOM.Matrix)

            ' Example: binding each column of matrix to datatable
            Dim col As SAPbouiCOM.Column
            ' Clear existing rows in matrix (optional)
            If objMatrix.RowCount > 0 Then
                For i As Integer = objMatrix.RowCount To 1 Step -1
                    objMatrix.DeleteRow(i)
                Next
            End If
            ' Add rows and set values manually
            'For i As Integer = 0 To oDataSource.Rows.Count - 1
            '    objMatrix.AddRow()
            '    objMatrix.Columns.Item("LineId").Cells.Item(i + 1).Specific.Value = (i + 1).ToString()
            '    objMatrix.Columns.Item("BCODE").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("Barcode", i).ToString()
            '    objMatrix.Columns.Item("ITMCODE").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("ItemCode", i).ToString()
            '    objMatrix.Columns.Item("ITMNM").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("ItemDescription", i).ToString()
            '    objMatrix.Columns.Item("UOM").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("UOM", i).ToString()
            '    objMatrix.Columns.Item("PRICE").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("Price", i).ToString()
            'Next
            For i As Integer = 0 To oDataSource.Rows.Count - 1
                objMatrix.AddRow()
                oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
                oDS1.SetValue("Code", oDS1.Offset, oDataSource.GetValue("Barcode", i).ToString())
                oDS1.SetValue("U_ITMCODE", oDS1.Offset, oDataSource.GetValue("ItemCode", i).ToString())
                oDS1.SetValue("U_ITMNM", oDS1.Offset, oDataSource.GetValue("ItemDescription", i).ToString())
                oDS1.SetValue("U_UOM", oDS1.Offset, oDataSource.GetValue("UOM", i).ToString())
                oDS1.SetValue("U_PRICE", oDS1.Offset, oDataSource.GetValue("Price", i).ToString())
                objMatrix.SetLineData(objMatrix.VisualRowCount)
            Next
        Catch ex As Exception
            objMain.objApplication.MessageBox("Error loading Excel to matrix: " & ex.Message)
        End Try
    End Sub
    'Public Function IsItemInOverlappingPromotion(fromDate As DateTime, toDate As DateTime, fromTime As String, toTime As String, itemCode As String) As Boolean
    '    ' Assuming fromTime and toTime are in format "HHmm" (e.g., "1300" for 1:00 PM)
    '    ' U_FRMTIME and U_TOTIME must also be stored as strings or integers in the same format (not TIME type)

    '    Dim query As String =
    '    "SELECT COUNT(*) FROM ""@TNX_BGLEASTITMS_C0"" T0 " &
    '    "INNER JOIN ""@TNX_BGLEASTITMS"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
    '    "WHERE T0.""U_ITMCODE"" = '" & itemCode & "' " &
    '    "AND T1.""U_STATUS"" = 'Active' " &
    '    "AND ( " &
    '    "    ('" & fromDate.ToString("yyyy-MM-dd") & "' <= T1.""U_TODATE"" " &
    '    "     AND '" & toDate.ToString("yyyy-MM-dd") & "' >= T1.""U_FRMDATE"") " &
    '    "    AND " &
    '    "    (" &
    '    "        '" & fromTime & "' <= T1.""U_TOTIME"" AND '" & toTime & "' >= T1.""U_FRMTIME"" " &
    '    "    ) " &
    '    ")"
    '    Dim oRS As SAPbobsCOM.Recordset
    '    oRS = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
    '    oRS.DoQuery(query)

    '    Return oRS.RecordCount > 0 AndAlso Convert.ToInt32(oRS.Fields.Item(0).Value) > 0
    'End Function


End Class
