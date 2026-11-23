'
'Akhila
Imports System.Collections.Specialized
Imports System.Configuration
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Security.Authentication.ExtendedProtection
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports SAPbobsCOM
Imports SAPbouiCOM
Public Class clsPromotionOnVoucherPurchase
    ''Test
#Region "       Declaration             "
    Public objForm, objSTRForm As SAPbouiCOM.Form
    'Dim oDS, oDS1 As SAPbouiCOM.DBDataSource
    Dim objMatrix, objMatrix1, objMatrix2 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Public oDS, oDS1, oDS2, oDS3 As SAPbouiCOM.DBDataSource
    Public PrgBar As SAPbouiCOM.ProgressBar
    Dim LostFocusFlag As Boolean
    Dim MATRIXS As String

#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("PromotionOnVoucherPurchase.xml", "PRMONVP", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("PRMONVP", objMain.objApplication.Forms.ActiveForm.TypeCount)
            'objForm.Freeze(True)
            oDS = objForm.DataSources.DBDataSources.Item("@TNX_PONVP")
            objForm.EnableMenu("774", True)

            Dim sQury As String = "Select ""ItmsGrpCod"" , ""ItmsGrpNam""   From OITB  "
            objComboBox = objForm.Items.Item("IGROUP").Specific
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, sQury)
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("IGROUP").DisplayDesc = True
            Dim ItemGroup As String
            'Me.GetNextDocNum()
            Me.DisableCoulmns(objForm.UniqueID)



            Dim oComb As SAPbouiCOM.ComboBox = objForm.Items.Item("BPLID").Specific

            'objMain.objUtilities.ComboBoxLoadValues(oComb, "Select ""BPLId"",""BPLName"" FROM OBPL Where ""BPLId""<>'1'")
            objMain.objUtilities.ComboBoxLoadValues(oComb, "Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join " & " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""   ")

            oComb.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("BPLID").DisplayDesc = True



            'oDS1 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PONVP_C0")
            'objMatrix = objForm.Items.Item("ITMMtx").Specific
            'objComboBox = objMatrix.Columns.Item("VOHRORDIS")
            'objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            objForm.Items.Item("BItmDet").Width = objForm.Items.Item("BItmDet").Width + 140
            objForm.Items.Item("VRedItm").Left = objForm.Items.Item("VRedItm").Left + 140
            objForm.Items.Item("Branch").Left = objForm.Items.Item("Branch").Left + 140
            objForm.Items.Item("BItmDet").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            objComboBox = Me.objForm.Items.Item("STATUS").Specific
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

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



            Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            Me.objForm.Items.Item("DOCDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Me.objForm.Items.Item("DOCDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Me.objForm.Items.Item("DOCDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Me.objForm.Items.Item("DOCDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objMatrix = objForm.Items.Item("ITMMtx").Specific
            objMatrix.Columns.Item("VOHRORDIS").ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

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
    Private Sub DisableCoulmns(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("ITMMtx").Specific
            Try
                'objMatrix.Columns.Item("DPRCTG").Visible = False
                'objMatrix.Columns.Item("DBAMT").Visible = False
                'objMatrix.Columns.Item("FRMAMNT").Visible = False
                'objMatrix.Columns.Item("TOAMOUNT").Visible = False

                objMatrix.Columns.Item("DPRCTG").Visible = True
                objMatrix.Columns.Item("DBAMT").Visible = True
                objMatrix.Columns.Item("FRMAMNT").Visible = False
                objMatrix.Columns.Item("TOAMOUNT").Visible = False
                objMatrix.Columns.Item("VOHRORDIS").Visible = False
                objMatrix.Columns.Item("PERORAMONT").Visible = False
                objMatrix.Columns.Item("PERCNTG").Visible = False


                objMatrix.Columns.Item("DPRCTG").Visible = False
                objMatrix.Columns.Item("DBAMT").Visible = False


                ' objMatrix.Columns.Item("FRMAMNT").Visible = True
                ' objMatrix.Columns.Item("TOAMOUNT").Visible = True


            Catch ex As Exception
            End Try
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try

            If pVal.MenuUID = "ME_PRMONVPA" And pVal.BeforeAction = False Then
                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)
                Me.BranchLoad(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
                Me.BranchLoad(objForm.UniqueID)
                ' Me.GetNextDocNum()

            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                Me.SetNewLine(objForm.UniqueID, MATRIXS)
            ElseIf pVal.MenuUID = "774" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("PRMONVP", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix = objForm.Items.Item("ITMMtx").Specific
                objMatrix1 = objForm.Items.Item("VRMtx").Specific
                Dim row As Integer = objMatrix.VisualRowCount
                If MATRIXS.Equals("ITMMtx") = True Then
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
                                '  If i <> 1 Then
                                objMatrix.DeleteRow(i)
                                '   End If
                            End If
                        Next
                    End If
                    For i As Integer = 1 To objMatrix.VisualRowCount
                        objMatrix.Columns.Item("LineId").Cells.Item(i).Specific.Value = i
                        'objMatrix.FlushToDataSource()

                    Next
                    'ElseIf MATRIXS.Equals("VRMtx") = True Then
                    '    If objMatrix1.IsRowSelected(1) <> True And objMatrix1.VisualRowCount < 1 Then
                    '        objMatrix1.AddRow()
                    '        oDS2.SetValue("LineId", oDS1.Offset, objMatrix1.VisualRowCount)
                    '        objMatrix1.SetLineData(objMatrix1.VisualRowCount)
                    '    End If
                    '    If objMatrix1.IsRowSelected(row) = True Then
                    '        objMatrix1.DeleteRow(row)
                    '    Else
                    '        For i As Integer = 1 To objMatrix1.VisualRowCount - 1

                    '            If objMatrix1.IsRowSelected(i) = True Then
                    '                '  If i <> 1 Then
                    '                objMatrix1.DeleteRow(i)
                    '                '   End If
                    '            End If
                    '        Next
                    '    End If
                    '    For i As Integer = 1 To objMatrix1.VisualRowCount
                    '        objMatrix1.Columns.Item("LineId").Cells.Item(i).Specific.Value = i
                    '    Next
                    'End If
                ElseIf MATRIXS.Equals("VRMtx") = True Then

                    ' Add a row if none are selected and the matrix is empty
                    If objMatrix1.IsRowSelected(1) <> True And objMatrix1.VisualRowCount < 1 Then
                        objMatrix1.AddRow()
                        oDS2.SetValue("LineId", oDS1.Offset, objMatrix1.VisualRowCount)
                        objMatrix1.SetLineData(objMatrix1.VisualRowCount)
                    End If

                    ' Delete selected rows - loop from bottom to top to prevent index issues
                    For i As Integer = objMatrix1.VisualRowCount To 1 Step -1
                        If objMatrix1.IsRowSelected(i) = True Then
                            objMatrix1.DeleteRow(i)
                        End If
                    Next

                    ' Reassign line numbers to LineId column
                    For i As Integer = 1 To objMatrix1.VisualRowCount
                        objMatrix1.Columns.Item("LineId").Cells.Item(i).Specific.Value = i.ToString()
                    Next

                End If

                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If
            End If


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



    '-----------------FILTER_CFL-----------------------------'
    'Public Shared Sub FILTER_CFL(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef oApplication As SAPbouiCOM.Application, ByRef oCompany As SAPbobsCOM.Company, ByVal Query As String, ByVal Code As String, ByVal Subgrp As String, ByVal Subsubgrp As String)
    '    Try
    '        Dim oCLFEvent As SAPbouiCOM.ChooseFromListEvent = Nothing
    '        Dim oCell As SAPbouiCOM.Column
    '        Dim oCFL As SAPbouiCOM.ChooseFromList
    '        Dim oCFL_collection As SAPbouiCOM.ChooseFromListCollection
    '        Dim oCon As SAPbouiCOM.Condition = Nothing
    '        Dim emptyConds As SAPbouiCOM.Conditions = Nothing
    '        Dim oCons As SAPbouiCOM.Conditions = Nothing
    '        Dim oRec As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
    '        oRec.DoQuery(Query)
    '        Dim strCFLID As String = Nothing
    '        oCLFEvent = CType((pVal), SAPbouiCOM.ChooseFromListEvent)
    '        strCFLID = oCLFEvent.ChooseFromListUID
    '        oCFL = oApplication.Forms.ActiveForm.ChooseFromLists.Item(strCFLID)
    '        oCFL.SetConditions(emptyConds)
    '        oCons = oCFL.GetConditions()
    '        Dim rr As Integer = oRec.RecordCount


    '        If rr = 0 Then
    '            oCon = oCons.Add()
    '            oCon.[Alias] = Code
    '            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '            oCon.CondVal = oRec.Fields.Item(Code).Value.ToString()
    '        Else


    '            For i As Integer = 0 To oRec.RecordCount - 1
    '                oCon = oCons.Add()
    '                oCon.[Alias] = Code
    '                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '                oCon.CondVal = oRec.Fields.Item(Code).Value.ToString()


    '                If i <> oRec.RecordCount - 1 Then
    '                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
    '                End If


    '                oRec.MoveNext()
    '            Next
    '        End If

    '        oCFL.SetConditions(oCons)
    '    Catch ex As Exception
    '    End Try
    'End Sub
    '--------Barcode , item code cfl--------------------

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


    ''-----------------------












    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "ITMMtx" Or pVal.ItemUID = "VRMtx" Or pVal.ItemUID = "BMtx" Then ' Or pVal.ItemUID = "Matrx5" Or pVal.ItemUID = "Matrix" Then
                        MATRIXS = pVal.ItemUID
                    End If


                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "ITMMtx" And pVal.ColUID = "VITMCDE" And pVal.ColUID = "ITEMN" And pVal.BeforeAction = False Then
                        Try
                            objForm.Freeze(True)
                            Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("ITMMtx").Specific
                            If pVal.ItemUID = "ITMMtx" Then
                                If pVal.Row = objMatrix.VisualRowCount Then
                                    Me.SetNewLine(objForm.UniqueID, "ITMMtx")
                                End If
                            End If
                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                        End Try
                    End If


                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK


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

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And
             (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        Me.SetDefault(objForm.UniqueID)
                        Me.BranchLoad(objForm.UniqueID)
                        'Me.GetNextDocNum()
                    End If

                    '--load functionality-----------------------


                    'Load old-----------
                    '        If pVal.ItemUID = "LOAD" AndAlso pVal.BeforeAction = False AndAlso objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then


                    '            Try
                    '                objForm.Freeze(True)
                    '                Dim oDS As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Add("@TNX_PONVP")
                    '                Dim oDS2 As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Add("@TNX_PONVP_C0")
                    '                Dim objMatrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("ITMMtx").Specific, SAPbouiCOM.Matrix)
                    '                Dim docnum As String = objForm.Items.Item("DocNum").Specific.Value
                    '                Dim protype As String = objForm.Items.Item("PTYPE").Specific.Value
                    '                Dim Itemgrp As String = objForm.Items.Item("IGROUP").Specific.Value
                    '                Dim subsub As String = objForm.Items.Item("SUBGRP").Specific.Value
                    '                Dim subsubgrp As String = objForm.Items.Item("SSBGRP").Specific.Value
                    '                Dim Discounttype As String = objForm.Items.Item("DTYPE").Specific.Value
                    '                Dim DiscountAmt As String = objForm.Items.Item("DISAT").Specific.Value
                    '                'Dim Disctprctg As String = objForm.Items.Item("DISPER").Specific.Value

                    '                'Dim qrt As String = "Select H.""U_VAMT"", H.""U_DISPER"" FROM ""@TNX_PONVP"" H INNER JOIN ""@TNX_PONVP_C0"" C On H.""DocEntry"" = C.""DocEntry"" where H.""DocNum""='" & docnum & "' "
                    '                'Dim orss As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                    '                'orss.DoQuery(qrt)
                    '                'If orss.RecordCount > 0 Then
                    '                '    Dim vamnt As String = objForm.Items.Item("U_VAMT").Specific.Value
                    '                '    Dim Disper As String = objForm.Items.Item("U_DISPER").Specific.Value

                    '                'End If

                    '                ' Get discount type from combo box
                    '                Dim discountType1 As String = CType(objForm.Items.Item("DTYPE").Specific, SAPbouiCOM.ComboBox).Selected.Value.Trim()
                    '                Dim discountAmount As String = objForm.Items.Item("DISAT").Specific.Value.Trim()
                    '                Dim percentage As String = objForm.Items.Item("DISPER").Specific.Value.Trim()


                    '                Dim rs1 As String =
                    '"SELECT I.""ItemCode"", I.""ItemName"", I.""U_VSubGroup"", I.""U_VSubSubGroup"" " &
                    '"FROM ""OBCD"" B " &
                    '"INNER JOIN ""OITM"" I ON B.""ItemCode"" = I.""ItemCode"" " &
                    '"WHERE I.""ItmsGrpCod"" = '" & Itemgrp & "' " &
                    '"or I.""U_VSubGroup"" = '" & subsub & "' " &
                    '"or I.""U_VSubSubGroup"" = '" & subsubgrp & "'"
                    '                Dim ors1 As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                    '                ors1.DoQuery(rs1)
                    '                If objMatrix.RowCount > 0 Then
                    '                    For i As Integer = objMatrix.RowCount To 1 Step -1
                    '                        objMatrix.DeleteRow(i)
                    '                    Next
                    '                End If
                    '                If ors1.RecordCount > 0 Then
                    '                    For j As Integer = 0 To ors1.RecordCount - 1
                    '                        objMatrix.AddRow()
                    '                        oDS2.SetValue("LineId", oDS2.Offset, objMatrix.VisualRowCount)
                    '                        oDS2.SetValue("U_VITMCDE", oDS2.Offset, ors1.Fields.Item("ItemCode").Value.ToString())
                    '                        oDS2.SetValue("U_ITEMN", oDS2.Offset, ors1.Fields.Item("ItemName").Value.ToString())
                    '                        If discountType1 = "P" Then
                    '                            oDS2.SetValue("U_DPRCTG", oDS2.Offset, percentage)
                    '                        ElseIf discountType1 = "A" Then
                    '                            oDS2.SetValue("U_DBAMT", oDS2.Offset, discountAmount)
                    '                        End If
                    '                        objMatrix.SetLineData(objMatrix.VisualRowCount)
                    '                        ors1.MoveNext()
                    '                    Next
                    '                Else
                    '                    Me.SetNewLine(objForm.UniqueID, "ITMMtx")
                    '                End If
                    '                objMatrix.AutoResizeColumns()
                    '                objMatrix.FlushToDataSource()
                    '            Catch ex As Exception
                    '                objForm.StatusBar.SetText("Error: " & ex.Message & " - Page Load Events", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '            Finally
                    '                objForm.Freeze(False)
                    '            End Try
                    '        End If


                    '-----------------end load-----
                    '---------------------------------

                    'LOAD NEW FUNCTIONALITY-------- WORKING NEW CODE
                    '                 If pVal.ItemUID = "LOAD" AndAlso pVal.BeforeAction = False AndAlso
                    '(objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or
                    ' objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or
                    ' objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                    '                     Try
                    '                         objForm.Freeze(True)

                    '                         Dim oDS As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Add("@TNX_PONVP")
                    '                         Dim oDS2 As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Add("@TNX_PONVP_C0")
                    '                         Dim objMatrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("ITMMtx").Specific, SAPbouiCOM.Matrix)

                    '                         Dim docnum As String = objForm.Items.Item("DocNum").Specific.Value
                    '                         Dim protype As String = objForm.Items.Item("PTYPE").Specific.Value
                    '                         Dim Itemgrp As String = objForm.Items.Item("IGROUP").Specific.Value
                    '                         Dim subsub As String = objForm.Items.Item("SUBGRP").Specific.Value
                    '                         Dim subsubgrp As String = objForm.Items.Item("SSBGRP").Specific.Value
                    '                         ' Dim Discounttype As String = objForm.Items.Item("DTYPE").Specific.Value
                    '                         'Dim DiscountAmt As String = objForm.Items.Item("DISAT").Specific.Value

                    '                         ' Get discount values
                    '                         'Dim discountType1 As String = CType(objForm.Items.Item("DTYPE").Specific, SAPbouiCOM.ComboBox).Selected.Value.Trim()
                    '                         'Dim discountAmount As String = objForm.Items.Item("DISAT").Specific.Value.Trim()
                    '                         'Dim percentage As String = objForm.Items.Item("DISPER").Specific.Value.Trim()

                    '                         ' Build SQL query based on selected filter
                    '                         Dim rs1 As String = ""
                    '                         If Not String.IsNullOrEmpty(subsubgrp) Then
                    '                             rs1 = "SELECT DISTINCT I.""ItemCode"", I.""ItemName"", I.""U_VSubGroup"", I.""U_VSubSubGroup"" " &
                    '               "FROM ""OITM"" I " &
                    '               "WHERE I.""U_VSubSubGroup"" = '" & subsubgrp & "'"
                    '                         ElseIf Not String.IsNullOrEmpty(subsub) Then
                    '                             rs1 = "SELECT DISTINCT I.""ItemCode"", I.""ItemName"", I.""U_VSubGroup"", I.""U_VSubSubGroup"" " &
                    '               "FROM ""OITM"" I " &
                    '               "WHERE I.""U_VSubGroup"" = '" & subsub & "'"
                    '                         ElseIf Not String.IsNullOrEmpty(Itemgrp) Then
                    '                             rs1 = "SELECT DISTINCT I.""ItemCode"", I.""ItemName"", I.""U_VSubGroup"", I.""U_VSubSubGroup"" " &
                    '               "FROM ""OITM"" I " &
                    '               "WHERE I.""ItmsGrpCod"" = '" & Itemgrp & "'"
                    '                         End If

                    '                         Dim ors1 As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                    '                         ors1.DoQuery(rs1)

                    '                         ' Clear matrix
                    '                         If objMatrix.RowCount > 0 Then
                    '                             For i As Integer = objMatrix.RowCount To 1 Step -1
                    '                                 objMatrix.DeleteRow(i)
                    '                             Next
                    '                         End If

                    '                         ' Load new data
                    '                         If ors1.RecordCount > 0 Then
                    '                             For j As Integer = 0 To ors1.RecordCount - 1
                    '                                 objMatrix.AddRow()
                    '                                 Dim currentRow As Integer = objMatrix.VisualRowCount
                    '                                 Dim itmCode As String = ors1.Fields.Item("ItemCode").Value.ToString()

                    '                                 oDS2.SetValue("LineId", oDS2.Offset, currentRow)
                    '                                 oDS2.SetValue("U_VITMCDE", oDS2.Offset, ors1.Fields.Item("ItemCode").Value.ToString())
                    '                                 oDS2.SetValue("U_ITEMN", oDS2.Offset, ors1.Fields.Item("ItemName").Value.ToString())

                    '                                 Dim GetBarCode As String =
                    '                 "SELECT TOP 1 T0.""UomEntry"", T0.""BcdCode"", T1.""UomCode"" " &
                    '                 "FROM ""OBCD"" T0 " &
                    '                 "INNER JOIN ""OUOM"" T1 ON T0.""UomEntry"" = T1.""UomEntry"" " &
                    '                 "WHERE T0.""ItemCode"" = '" & itmCode & "' " &
                    '                 "ORDER BY T0.""BcdEntry"" ASC"

                    '                                 ' Get barcode and UOM
                    '                                 Dim oRsGetBarCode As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                    '                                 oRsGetBarCode.DoQuery(GetBarCode)

                    '                                 If oRsGetBarCode.RecordCount > 0 Then
                    '                                     oDS2.SetValue("U_BARC", oDS2.Offset, oRsGetBarCode.Fields.Item("BcdCode").Value.ToString())
                    '                                     oDS2.SetValue("U_UOM", oDS2.Offset, oRsGetBarCode.Fields.Item("UomCode").Value.ToString())
                    '                                 End If

                    '                                 'If discountType1 = "P" Then
                    '                                 '    oDS2.SetValue("U_DPRCTG", oDS2.Offset, percentage)
                    '                                 'ElseIf discountType1 = "A" Then
                    '                                 '    oDS2.SetValue("U_DBAMT", oDS2.Offset, discountAmount)
                    '                                 'End If

                    '                                 objMatrix.SetLineData(currentRow)
                    '                                 ors1.MoveNext()
                    '                             Next
                    '                         Else
                    '                             Me.SetNewLine(objForm.UniqueID, "ITMMtx")
                    '                         End If

                    '                         objMatrix.AutoResizeColumns()
                    '                         objMatrix.FlushToDataSource()

                    '                     Catch ex As Exception
                    '                         objForm.StatusBar.SetText("Error: " & ex.Message & " - Page Load Events", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '                     Finally
                    '                         objForm.Freeze(False)
                    '                     End Try

                    '                 End If

                    '--------------LOAD NEW FUNCTIOMNALITY-----

                    If pVal.ItemUID = "LOAD" AndAlso pVal.BeforeAction = False AndAlso
(objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or
objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or
objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                        Try

                            objForm.Freeze(True)

                            Dim oDS As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Add("@TNX_PONVP")

                            Dim oDS2 As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Add("@TNX_PONVP_C0")

                            Dim objMatrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("ITMMtx").Specific, SAPbouiCOM.Matrix)

                            Dim docnum As String = objForm.Items.Item("DocNum").Specific.Value

                            Dim protype As String = objForm.Items.Item("PTYPE").Specific.Value

                            Dim Itemgrp As String = objForm.Items.Item("IGROUP").Specific.Value

                            Dim subsub As String = objForm.Items.Item("SUBGRP").Specific.Value

                            Dim subsubgrp As String = objForm.Items.Item("SSBGRP").Specific.Value

                            Dim Branch As String = objForm.Items.Item("BPLID").Specific.Value

                            ' Build SQL query based on selected filter

                            Dim rs1 As String = ""

                            If Not String.IsNullOrEmpty(subsubgrp) Then

                                rs1 = "SELECT DISTINCT I.""ItemCode"", I.""ItemName"", I.""U_VSubGroup"", I.""U_VSubSubGroup"" " &
               "FROM ""OITM"" I WHERE I.""U_VSubSubGroup"" = '" & subsubgrp & "'"

                            ElseIf Not String.IsNullOrEmpty(subsub) Then

                                rs1 = "SELECT DISTINCT I.""ItemCode"", I.""ItemName"", I.""U_VSubGroup"", I.""U_VSubSubGroup"" " &
               "FROM ""OITM"" I WHERE I.""U_VSubGroup"" = '" & subsub & "'"

                            ElseIf Not String.IsNullOrEmpty(Itemgrp) Then

                                rs1 = "SELECT DISTINCT I.""ItemCode"", I.""ItemName"", I.""U_VSubGroup"", I.""U_VSubSubGroup"" " &
               "FROM ""OITM"" I WHERE I.""ItmsGrpCod"" = '" & Itemgrp & "'"

                            End If

                            Dim ors1 As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

                            ors1.DoQuery(rs1)

                            ' Clear matrix

                            If objMatrix.RowCount > 0 Then
                                For i As Integer = objMatrix.RowCount To 1 Step -1
                                    objMatrix.DeleteRow(i)
                                Next
                            End If

                            ' Load new data

                            If ors1.RecordCount > 0 Then

                                For j As Integer = 0 To ors1.RecordCount - 1

                                    objMatrix.AddRow()

                                    Dim currentRow As Integer = objMatrix.VisualRowCount

                                    Dim ItemCode As String = ors1.Fields.Item("ItemCode").Value.ToString()

                                    oDS2.SetValue("LineId", oDS2.Offset, currentRow)
                                    oDS2.SetValue("U_VITMCDE", oDS2.Offset, ItemCode)
                                    oDS2.SetValue("U_ITEMN", oDS2.Offset, ors1.Fields.Item("ItemName").Value.ToString())



                                    Dim GetBarCode As String =
                 "SELECT TOP 1 T0.""UomEntry"", T0.""BcdCode"", T1.""UomCode"" " &
                 "FROM ""OBCD"" T0 " &
                 "INNER JOIN ""OUOM"" T1 ON T0.""UomEntry"" = T1.""UomEntry"" " &
                 "WHERE T0.""ItemCode"" = '" & ItemCode & "' ORDER BY T0.""BcdEntry"" ASC"

                                    Dim oRsGetBarCode As SAPbobsCOM.Recordset =
                 CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset),
                 SAPbobsCOM.Recordset)

                                    oRsGetBarCode.DoQuery(GetBarCode)

                                    Dim Barcode As String = ""

                                    If oRsGetBarCode.RecordCount > 0 Then

                                        Barcode = oRsGetBarCode.Fields.Item("BcdCode").Value.ToString()

                                        oDS2.SetValue("U_BARC", oDS2.Offset, Barcode)

                                        oDS2.SetValue("U_UOM", oDS2.Offset, oRsGetBarCode.Fields.Item("UomCode").Value.ToString())

                                    End If


                                    Try

                                        Dim q1 As String = "CALL ""GetItemPrice""('" & ItemCode & "','" & Branch & "')"

                                        Dim rsCost As SAPbobsCOM.Recordset =
                     objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                        rsCost.DoQuery(q1)

                                        If rsCost.RecordCount > 0 Then
                                            Dim ItemCost As Double = rsCost.Fields.Item("AvgPrice").Value
                                            oDS2.SetValue("U_ITMCS", oDS2.Offset, ItemCost)
                                        Else
                                            oDS2.SetValue("U_ITMCS", oDS2.Offset, 0)
                                        End If
                                    Catch ex As Exception

                                    End Try

                                    Try

                                        Dim q2 As String = "CALL ""GetRetailPrice""('" & Barcode & "','" & Branch & "')"
                                        Dim rsRetail As SAPbobsCOM.Recordset =
                     objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                        rsRetail.DoQuery(q2)
                                        If rsRetail.RecordCount > 0 Then
                                            Dim RetailPrice As Double = rsRetail.Fields.Item("Price").Value
                                            oDS2.SetValue("U_RETPC", oDS2.Offset, RetailPrice)
                                        Else
                                            oDS2.SetValue("U_RETPC", oDS2.Offset, 0)
                                        End If

                                    Catch ex As Exception

                                    End Try


                                    objMatrix.SetLineData(currentRow)

                                    ors1.MoveNext()

                                Next

                            Else

                                Me.SetNewLine(objForm.UniqueID, "ITMMtx")

                            End If

                            objMatrix.AutoResizeColumns()

                            objMatrix.FlushToDataSource()

                        Catch ex As Exception

                            objForm.StatusBar.SetText("Error: " & ex.Message & " - Page Load Events",
         SAPbouiCOM.BoMessageTime.bmt_Long,
         SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        Finally

                            objForm.Freeze(False)

                        End Try

                    End If




                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
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
                            objMain.objApplication.StatusBar.SetText("please select 'To Date'. Please use format yyyyMMdd.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If



                        If oDS.GetValue("U_BPLID", 0) = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Select Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If



                        If oDS.GetValue("U_REDMNBAMNT", 0) = 0 Then
                            objMain.objApplication.StatusBar.SetText("Please Enter Redemption Minimum Bill Amount", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
                        If oDS.GetValue("U_PName", 0) = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter Promotion Name", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        'If oDS.GetValue("U_DTYPE", 0) = "" Then
                        '    objMain.objApplication.StatusBar.SetText("Discount Type Cannot be Null", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If

                        'Get matrix And loop through item rows
                        'Dim matrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("ITMMtx").Specific, SAPbouiCOM.Matrix)
                        'For i As Integer = 1 To matrix.RowCount
                        '    Dim itemCode As String = CType(matrix.Columns.Item("VITMCDE").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value
                        '    If itemCode <> "" Then
                        '        If IsItemInOverlappingPromotion(fromDate, toDate, FromTime, ToTime, itemCode) Then
                        '            objMain.objApplication.StatusBar.SetText("Row " & i.ToString() & ": Item " & itemCode & " is already in another active promotion with overlapping dates.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '            BubbleEvent = False
                        '            Exit Sub
                        '        End If
                        '    End If
                        'Next
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If Me.ValidatePromotionName(objForm.UniqueID) = False Then
                            BubbleEvent = False
                        End If
                    End If
                    ' --Akhila Overlapping code

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
                    '                         Dim oldStatusQuery As String = "SELECT ""U_STATUS"" FROM ""@TNX_PONVP"" WHERE ""DocNum"" = '" & docnum & "'"
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
                    '                         Dim matrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("ITMMtx").Specific, SAPbouiCOM.Matrix)
                    '                         For i As Integer = 1 To matrix.VisualRowCount
                    '                             matrix.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    '                             Dim itemCode As String = CType(matrix.Columns.Item("VITMCDE").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value.Trim()
                    '                             If itemCode <> "" Then
                    '                                 Dim query As String = "SELECT H1.""DocNum"" FROM ""@TNX_PONVP"" H1 " &
                    '                 "INNER JOIN ""@TNX_PONVP_C0"" C1 ON H1.""DocEntry"" = C1.""DocEntry"" " &
                    '                 "WHERE C1.""U_VITMCDE"" = '" & itemCode & "' " &
                    '                 "AND H1.""U_FRMDATE"" <= '" & fromDate & "' " &
                    '                 "AND H1.""U_TODATE"" >= '" & toDate & "' " &
                    '                 "AND H1.""U_FRMTIME"" <= '" & FromTime & "' " &
                    '                 "AND H1.""U_TOTIME"" >= '" & ToTime & "' " &
                    '                 "AND H1.""U_STATUS"" = 'Active'"

                    '                                 Dim recordSet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                                 recordSet.DoQuery(query)
                    '                                 Dim doc As String = recordSet.Fields.Item("DocNum").Value

                    '                                 If recordSet.Fields.Item(0).Value > 0 Then
                    '                                     objMain.objApplication.StatusBar.SetText("DocNum " & doc & " - Row " & i.ToString() & ": Item " & itemCode & " is already in another active promotion with overlapping dates.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    '                                     BubbleEvent = False
                    '                                     'Exit Sub
                    '                                 End If
                    '                             End If
                    '                         Next
                    '                     End If
                    '                 End If



                    '                 If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                    '(pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                    '                     Dim matrix11 As SAPbouiCOM.Matrix = CType(objForm.Items.Item("ITMMtx").Specific, SAPbouiCOM.Matrix)

                    '                     For i As Integer = 1 To matrix11.VisualRowCount
                    '                         'matrix11.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                    '                         Dim itemCode1 As String = CType(matrix11.Columns.Item("VITMCDE          ").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value.Trim()
                    '                         If String.IsNullOrEmpty(itemCode1) Then Continue For

                    '                         Dim fromTime1 As String = CType(objForm.Items.Item("FRMTIME").Specific, SAPbouiCOM.EditText).Value.Trim()
                    '                         Dim toTime1 As String = CType(objForm.Items.Item("TOTIME").Specific, SAPbouiCOM.EditText).Value.Trim()
                    '                         Dim fromDate1 As String = objForm.Items.Item("FRMDATE").Specific.Value.Trim()
                    '                         Dim toDate1 As String = objForm.Items.Item("TODATE").Specific.Value.Trim()
                    '                         Dim currentDocNum As String = objForm.Items.Item("DocNum").Specific.Value.Trim()

                    '                         Dim query As String =
                    ' "WITH ItemExistence AS ( " &
                    ' "     SELECT 'Buy and get least Item' AS TableName,a.""DocNum"" as DocNum " &
                    ' "    FROM ""@TNX_BGLEASTITMS"" a " &
                    ' "    JOIN ""@TNX_BGLEASTITMS_C0"" b ON a.""DocEntry"" = b.""DocEntry"" " &
                    ' "    WHERE b.""U_ITMCODE"" = '" & itemCode1 & "' " &
                    ' "      AND a.""U_STATUS"" = 'ACTIVE' " &
                    ' "      AND a.""U_FRMDATE"" <= '" & fromDate1 & "' " &
                    ' "      AND a.""U_TODATE"" >= '" & toDate1 & "' " &
                    ' "      AND '" & toTime1 & "' <= a.""U_TOTIME"" " &
                    ' "      AND '" & fromTime1 & "' >= a.""U_FRMTIME"" " &
                    ' "    UNION " &
                    ' "     Select 'Item wise Discount' AS TableName,a.""DocNum"" as DocNum " &
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





                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    '    If Me.BillAmount(objForm.UniqueID) = False Then
                    '        BubbleEvent = False
                    '    End If

                    'End If




                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    '    If Me.validateDiscountpercentage(objForm.UniqueID) = False Then
                    '        BubbleEvent = False
                    '    End If

                    ''End If

                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    '    If Me.validateDiscountAmount(objForm.UniqueID) = False Then
                    '        BubbleEvent = False
                    '    End If
                    'End If





                    'voucher amount comment------
                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    '    If Me.validateVocherAmount(objForm.UniqueID) = False Then
                    '        BubbleEvent = False
                    '    End If
                    'End If

                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    '    If Me.validateVocherpercentage(objForm.UniqueID) = False Then
                    '        BubbleEvent = False
                    '    End If
                    'End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                                  (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                        Dim projtype As String = objForm.Items.Item("PTYPE").Specific.Value
                        If Me.Duplicatebarcode(objForm.UniqueID, projtype) = False Then
                            LostFocusFlag = False
                            BubbleEvent = False
                        End If
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                                  (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        LostFocusFlag = True
                        If Me.Validate(objForm.UniqueID, "ITMMtx") = False Then
                            LostFocusFlag = False
                            BubbleEvent = False
                            'ElseIf Me.Validate(objForm.UniqueID, "IGMtx1") = False Then
                            '    LostFocusFlag = False
                            '    BubbleEvent = False
                        End If
                    End If
                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                    '    If Me.validateDiscountType(objForm.UniqueID) = False Then
                    '        BubbleEvent = False
                    '    End If
                    'End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                    'If pVal.BeforeAction = False AndAlso pVal.ItemUID = "PTYPE" Then
                    '    Dim objComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("PTYPE").Specific

                    '    If objComboBox.Value = "A" Then
                    '        objForm.Items.Item("IGROUP").Enabled = False
                    '        objForm.Items.Item("SUBGRP").Enabled = False
                    '        objForm.Items.Item("SSBGRP").Enabled = False
                    '    End If
                    'End If

                    If pVal.BeforeAction = False AndAlso pVal.ItemUID = "PTYPE" Then
                        Dim objComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("PTYPE").Specific
                        Dim promoType As String = objComboBox.Value.Trim()
                        Dim itemGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("IGROUP").Specific
                        Dim subGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("SUBGRP").Specific
                        Dim subSubGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("SSBGRP").Specific

                        If promoType = "A" Then
                            objForm.Items.Item("IGROUP").Enabled = False
                            objForm.Items.Item("SUBGRP").Enabled = False
                            objForm.Items.Item("SSBGRP").Enabled = False
                            itemGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            subGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            subSubGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        Else

                            objForm.Items.Item("IGROUP").Enabled = True
                            objForm.Items.Item("SUBGRP").Enabled = True
                            objForm.Items.Item("SSBGRP").Enabled = True
                            itemGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            subGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            subSubGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        End If

                    End If



                    If pVal.ItemUID = "STATUS" And pVal.BeforeAction = False Then
                        Dim status As String = objForm.Items.Item("STATUS").Specific.Value
                        If status = "Active" Then
                            objForm.Items.Item("FRMDATE").Enabled = True
                            objForm.Items.Item("TODATE").Enabled = True
                            objForm.Items.Item("REDMNBAMNT").Enabled = True
                            objForm.Items.Item("FRMTIME").Enabled = True
                            objForm.Items.Item("TOTIME").Enabled = True
                            objForm.Items.Item("PName").Enabled = True
                            objForm.Items.Item("ITMMtx").Enabled = True
                            objForm.Items.Item("VRMtx").Enabled = True
                            objForm.Items.Item("BMtx").Enabled = True
                        ElseIf status = "Inactive" Then
                            objForm.Items.Item("FRMDATE").Enabled = False
                            objForm.Items.Item("TODATE").Enabled = False
                            objForm.Items.Item("REDMNBAMNT").Enabled = False
                            objForm.Items.Item("FRMTIME").Enabled = False
                            objForm.Items.Item("TOTIME").Enabled = False
                            objForm.Items.Item("PName").Enabled = False
                            objForm.Items.Item("ITMMtx").Enabled = False
                            objForm.Items.Item("VRMtx").Enabled = False
                            objForm.Items.Item("BMtx").Enabled = False
                        End If
                    End If

                    'new series add-------
                    If pVal.ItemUID = "24" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("24").Specific.Selected.Value <> "" Then
                            Dim series As Integer = objForm.Items.Item("24").Specific.Selected.Value
                            SetDefault(objForm.UniqueID, "CobmoSelect", series)
                        End If
                    End If
                    'new end series


                    If pVal.BeforeAction = False AndAlso pVal.ItemUID = "DTYPE" Then
                        Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(pVal.FormUID)
                        Dim objComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("DTYPE").Specific
                        Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("ITMMtx").Specific ' Replace with actual matrix ID

                        If objComboBox.Value = "P" Then
                            objMatrix.Columns.Item("DPRCTG").Editable = True
                            objForm.Items.Item("DISAT").Enabled = False
                            objForm.Items.Item("DISPER").Enabled = True
                            objMatrix.Columns.Item("DBAMT").Editable = False

                        ElseIf objComboBox.Value = "A" Then
                            objMatrix.Columns.Item("DBAMT").Editable = True
                            objMatrix.Columns.Item("DPRCTG").Editable = False
                            objForm.Items.Item("DISAT").Enabled = True
                            objForm.Items.Item("DISPER").Enabled = False

                        End If
                    End If

                    'Vocher type:----


                    'If pVal.ItemUID = "VOHRORDIS" Then
                    If pVal.ItemUID = "ITMMtx" And pVal.ColUID = "VOHRORDIS" Then
                        Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("ITMMtx").Specific

                        Dim objComboBox As SAPbouiCOM.ComboBox = CType(objMatrix.Columns.Item("VOHRORDIS").Cells.Item(pVal.Row).Specific, SAPbouiCOM.ComboBox)

                        ' Dim objComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("VOHRORDIS").Specific

                        If objComboBox.Value = "Voucher" Then
                            objMatrix.Columns.Item("PERORAMONT").Editable = True
                            objMatrix.Columns.Item("PERCNTG").Editable = False

                        ElseIf objComboBox.Value = "Discount" Then
                            objMatrix.Columns.Item("PERCNTG").Editable = True
                            objMatrix.Columns.Item("PERORAMONT").Editable = False
                        End If
                    End If







                    If pVal.ItemUID = "PTYPE" Then
                        'Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(pVal.FormUID)
                        Dim objComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("PTYPE").Specific

                        If objComboBox.Value = "O" Then
                            objForm.Items.Item("IGROUP").Enabled = True
                            objForm.Items.Item("SUBGRP").Enabled = True
                            objForm.Items.Item("SSBGRP").Enabled = True
                        End If
                    End If
                    'If pVal.ItemUID = "IGROUP" Then
                    '    Dim cmb As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("IGROUP").Specific, SAPbouiCOM.ComboBox)
                    '    If cmb.Selected IsNot Nothing Then
                    '        Dim itemGroup As String = cmb.Selected.Value
                    '        LoadSubGroups1(objForm, itemGroup)
                    '    End If
                    'End If
                    'If pVal.ItemUID = "SUBGRP" Then
                    '    Dim cmbGrp As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("IGROUP").Specific, SAPbouiCOM.ComboBox)
                    '    Dim cmbSub As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("SUBGRP").Specific, SAPbouiCOM.ComboBox)
                    '    If cmbGrp.Selected IsNot Nothing AndAlso cmbSub.Selected IsNot Nothing Then
                    '        Dim itemGroup As String = cmbGrp.Selected.Value
                    '        Dim subGroup As String = cmbSub.Selected.Value
                    '        LoadSubSubGroups(objForm, itemGroup, subGroup)
                    '    End If
                    'End If



                    If objForm.Items.Item("IGROUP").Specific.Value <> "" Then
                        Dim objComboBoxItemGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("IGROUP").Specific
                        If objComboBoxItemGroup.Selected IsNot Nothing Then
                            Dim itemGroupCode As String = objComboBoxItemGroup.Selected.Value

                            Dim sql As String = "SELECT DISTINCT T0.""U_VSubGroup"", T0.""U_VSubGroup"" " &
                            "FROM OITM T0 " &
                            "WHERE T0.""ItmsGrpCod"" = '" & itemGroupCode & "' " &
                            "AND T0.""U_VSubGroup"" IS NOT NULL " &
                            "ORDER BY T0.""U_VSubGroup"""

                            Dim objComboBoxSubGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("SUBGRP").Specific

                            objMain.objUtilities.ComboBoxLoadValues(objComboBoxSubGroup, sql)
                            objComboBoxSubGroup.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                            objForm.Items.Item("SUBGRP").DisplayDesc = True
                        End If
                    End If





                    If objForm.Items.Item("SUBGRP").Specific.Value <> "" Then
                        Dim objComboBoxSubGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("SUBGRP").Specific

                        If objComboBoxSubGroup.Selected IsNot Nothing Then
                            Dim Subgroup As String = objComboBoxSubGroup.Selected.Value

                            ' Ensure ItemGroup is defined somewhere earlier, or retrieve it here again
                            Dim objComboBoxItemGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("IGROUP").Specific
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

                            Dim objComboBoxSubSubGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("SSBGRP").Specific

                            objMain.objUtilities.ComboBoxLoadValues(objComboBoxSubSubGroup, sql)
                            objComboBoxSubSubGroup.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                            objForm.Items.Item("SSBGRP").DisplayDesc = True
                        End If
                    End If

















                'Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                '    If pVal.ItemUID = "TODATE" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                '        If objForm.Items.Item("FRMDATE").Specific.Value > objForm.Items.Item("TODATE").Specific.Value Then
                '            objMain.objApplication.StatusBar.SetText("To Date must be after Date To From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '            BubbleEvent = False
                '        End If
                '    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDS = objForm.DataSources.DBDataSources.Item("@TNX_PONVP")
                    oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_PONVP_C0")
                    oDS2 = objForm.DataSources.DBDataSources.Item("@TNX_PONVP_C1")
                    objMatrix = objForm.Items.Item("ITMMtx").Specific
                    objMatrix1 = objForm.Items.Item("VRMtx").Specific
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    'If oCFL.UniqueID = "CFL_ITM" And pVal.BeforeAction = True Then
                    'Dim StrCFLQry As String = ""

                    ' If objForm.Items.Item("IGROUP").Specific.Value <> "" Then
                    'StrCFLQry = "SELECT B.""ItemCode"",  B.""BcdCode"" ,I.""U_VSubGroup"",I.""U_VSubSubGroup""
                    'FROM ""OBCD"" B " &
                    ' "INNER JOIN ""OITM"" I ON B.""ItemCode"" = I.""ItemCode"" " &
                    '"WHERE   B.""ItemCode"" IS NOT NULL "


                    'If CStr(objForm.Items.Item("IGROUP").Specific.Value).Trim <> "" Then
                    '    StrCFLQry = StrCFLQry & "  and  I.""ItmsGrpCod"" = '" & objForm.Items.Item("IGROUP").Specific.Value & "'"
                    'End If

                    'If CStr(objForm.Items.Item("SUBGRP").Specific.Value).Trim <> "" Then
                    '    StrCFLQry = StrCFLQry & "  And  I.""U_VSubGroup"" = '" & objForm.Items.Item("SUBGRP").Specific.Value & "'"
                    'End If

                    'If CStr(objForm.Items.Item("SSBGRP").Specific.Value).Trim <> "" Then
                    '    StrCFLQry = StrCFLQry & "  And  I.""U_VSubSubGroup"" = '" & objForm.Items.Item("SSBGRP").Specific.Value & "'"
                    'End If

                    ''    End If

                    '    FILTER_CFL(pVal, objMain.objApplication, objMain.objCompany, StrCFLQry, "ItemCode", "U_VSubGroup", "U_VSubSubGroup")


                    ''End If

                    '------------------------retail barcode uom -------------------------------------------------'
                    If (oCFL.UniqueID = "CFL_BCOD" Or oCFL.UniqueID = "CFL_ITM") And pVal.BeforeAction = True Then
                        ''Me.CFLFilter("CFL_OCRD")
                        Dim StrCFLQry As String = ""
                        'here changed

                        If objForm.Items.Item("IGROUP").Specific.Value <> "" Then
                            StrCFLQry = "SELECT I.""ItemCode"",  B.""BcdCode"" ,I.""U_VSubGroup"",I.""U_VSubSubGroup""
                    FROM ""OBCD"" B " &
                 "INNER JOIN ""OITM"" I ON B.""ItemCode"" = I.""ItemCode"" " &
                 "WHERE   B.""ItemCode"" IS NOT NULL "


                            If CStr(objForm.Items.Item("IGROUP").Specific.Value).Trim <> "" Then
                                StrCFLQry = StrCFLQry & "  and  I.""ItmsGrpCod"" = '" & objForm.Items.Item("IGROUP").Specific.Value & "'"
                            End If

                            If CStr(objForm.Items.Item("SUBGRP").Specific.Value).Trim <> "" Then
                                StrCFLQry = StrCFLQry & "  And  I.""U_VSubGroup"" = '" & objForm.Items.Item("SUBGRP").Specific.Value & "'"
                            End If

                            If CStr(objForm.Items.Item("SSBGRP").Specific.Value).Trim <> "" Then
                                StrCFLQry = StrCFLQry & "  And  I.""U_VSubSubGroup"" = '" & objForm.Items.Item("SSBGRP").Specific.Value & "'"
                            End If


                        End If
                        'here changed

                        If oCFL.UniqueID = "CFL_ITM" Then
                            CFL_Filter1(pVal, objMain.objApplication, objMain.objCompany, StrCFLQry, "BcdCode", "U_VSubSubGroup", "U_VSubSubGroup", "ItemCode")
                        End If

                        If oCFL.UniqueID = "CFL_BCOD" Then
                            CFL_Filter(pVal, objMain.objApplication, objMain.objCompany, StrCFLQry, "BcdCode", "U_VSubSubGroup", "U_VSubSubGroup", "ItemCode")
                        End If

                    End If

                    'Vocher barcode cfl 

                    If oCFL.UniqueID = "CFL_VBC" And pVal.BeforeAction = False Then

                        If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then

                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                            End If

                            Dim bcode As String = oDT.GetValue("BcdCode", 0)

                            Dim Itemcode As String = ""



                            Dim itmqury As String = "SELECT T1.""ItemCode"", T1.""ItemName"", T0.""BcdCode"", T2.""UomCode"" FROM ""OBCD"" T0 INNER JOIN ""OITM"" T1 ON T0.""ItemCode"" = T1.""ItemCode"" INNER JOIN ""OUOM"" T2 ON T0.""UomEntry"" = T2.""UomEntry"" WHERE T0.""BcdCode"" = '" & bcode & "'"

                            Dim oRitmqury As SAPbobsCOM.Recordset =
objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            oRitmqury.DoQuery(itmqury)

                            ' Set item info

                            oDS2.SetValue("LineId", oDS2.Offset, pVal.Row)

                            oDS2.SetValue("U_BARCD", oDS2.Offset, oRitmqury.Fields.Item("BcdCode").Value)

                            oDS2.SetValue("U_ITMCDE", oDS2.Offset, oRitmqury.Fields.Item("ItemCode").Value)

                            oDS2.SetValue("U_ITMDESC", oDS2.Offset, oRitmqury.Fields.Item("ItemName").Value)

                            objMatrix1.SetLineData(pVal.Row)

                            If objMatrix1.VisualRowCount = pVal.Row Then

                                Me.SetNewLine(objForm.UniqueID, "VRMtx")

                            End If

                        End If

                    End If



                    '-----------------BARCODE
                    If oCFL.UniqueID = "CFL_BCOD" And pVal.BeforeAction = False Then
                        If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            End If

                            Dim bcode As String = oDT.GetValue("BcdCode", 0)
                            Dim Itemcode As String = ""

                            oDS1.SetValue("U_BARC", oDS1.Offset, bcode)

                            Dim itmqury As String = "SELECT T1.""ItemCode"", T1.""ItemName"", T0.""BcdCode"", T2.""UomCode"" " &
                                "FROM ""OBCD"" T0 " &
                                "INNER JOIN ""OITM"" T1 ON T0.""ItemCode"" = T1.""ItemCode"" " &
                                "INNER JOIN ""OUOM"" T2 ON T0.""UomEntry"" = T2.""UomEntry"" " &
                                "WHERE T0.""BcdCode"" = '" & bcode & "'"

                            Dim oRitmqury As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRitmqury.DoQuery(itmqury)

                            ' Set item info
                            oDS1.SetValue("LineId", oDS1.Offset, pVal.Row)
                            oDS1.SetValue("U_VITMCDE", oDS1.Offset, oRitmqury.Fields.Item("ItemCode").Value)
                            oDS1.SetValue("U_ITEMN", oDS1.Offset, oRitmqury.Fields.Item("ItemName").Value)
                            oDS1.SetValue("U_UOM", oDS1.Offset, oRitmqury.Fields.Item("UomCode").Value)
                            oDS1.SetValue("U_BARC", oDS1.Offset, oRitmqury.Fields.Item("BcdCode").Value)

                            ' Get warehouse from OADM
                            Dim query As String = "SELECT ""U_CNWhse"" FROM OADM"
                            Dim orset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            orset.DoQuery(query)
                            Dim Warehouse As String = orset.Fields.Item("U_CNWhse").Value

                            Itemcode = oRitmqury.Fields.Item("ItemCode").Value
                            Dim Brch As String = objForm.Items.Item("BPLID").Specific.Value

                            ' Get Item Cost
                            Try
                                Dim query1 As String = "CALL ""GetItemPrice""('" & Itemcode & "','" & Brch.Trim & "')"
                                Dim orset1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                orset1.DoQuery(query1)

                                If orset1.RecordCount > 0 Then
                                    Dim Itemcost As Double = orset1.Fields.Item("AvgPrice").Value
                                    oDS1.SetValue("U_ITMCS", oDS1.Offset, Itemcost)
                                Else
                                    oDS1.SetValue("U_ITMCS", oDS1.Offset, 0)
                                End If
                            Catch ex As Exception
                                ' Optional: log or show error
                            End Try

                            Try
                                Dim query2 As String = "CALL ""GetRetailPrice""('" & bcode & "','" & Brch.Trim & "')"
                                Dim orset2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                orset2.DoQuery(query2)

                                If orset2.RecordCount > 0 Then
                                    Dim Retailprice As Double = orset2.Fields.Item("Price").Value
                                    oDS1.SetValue("U_RETPC", oDS1.Offset, Retailprice)
                                Else
                                    oDS1.SetValue("U_RETPC", oDS1.Offset, 0)
                                End If
                            Catch ex As Exception
                            End Try

                            objMatrix.SetLineData(pVal.Row)

                            If objMatrix.VisualRowCount = pVal.Row Then
                                Me.SetNewLine(objForm.UniqueID, "ITMMtx")
                            End If
                        End If
                    End If
                    'Working
                    'If oCFL.UniqueID = "CFL_VRI" And pVal.BeforeAction = False Then
                    '    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                    '        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    '            objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    '        End If
                    '        'objMatrix1 = objForm.Items.Item("VRMtx").Specific

                    '        oDS2.SetValue("LineId", oDS2.Offset, pVal.Row)
                    '        oDS2.SetValue("U_ITMCDE", oDS2.Offset, oDT.GetValue("ItemCode", 0))
                    '        oDS2.SetValue("U_ITMDESC", oDS2.Offset, oDT.GetValue("ItemName", 0))
                    '        objMatrix1.SetLineData(pVal.Row)
                    '        If pVal.Row = objMatrix1.VisualRowCount Then
                    '            Me.SetNewLine(objForm.UniqueID, "VRMtx")
                    '        End If
                    '    End If
                    'End If
                    If oCFL.UniqueID = "CFL_VRI" And pVal.BeforeAction = False Then
                        If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then

                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            End If

                            oDS2.SetValue("LineId", oDS2.Offset, pVal.Row)
                            oDS2.SetValue("U_ITMCDE", oDS2.Offset, oDT.GetValue("ItemCode", 0))
                            oDS2.SetValue("U_ITMDESC", oDS2.Offset, oDT.GetValue("ItemName", 0))

                            oDS2.SetValue("U_ITMCDE", oDS2.Offset, oDT.GetValue("ItemCode", 0))

                            ' Get barcode for selected ItemCode
                            Dim GetBarCode As String =
"SELECT TOP 1 ""BcdCode"" FROM OBCD WHERE ""ItemCode"" = '" & oDT.GetValue("ItemCode", 0) & "' ORDER BY ""BcdEntry"" ASC"

                            Dim oRsGetBarCode As SAPbobsCOM.Recordset =
objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetBarCode.DoQuery(GetBarCode)

                            Dim Barcode As String = ""
                            If oRsGetBarCode.RecordCount > 0 Then
                                Barcode = oRsGetBarCode.Fields.Item("BcdCode").Value.ToString()
                            End If

                            oDS2.SetValue("U_BARCD", oDS2.Offset, Barcode)
                            objMatrix1.SetLineData(pVal.Row)

                            If pVal.Row = objMatrix1.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "VRMtx")
                            End If

                        End If
                    End If

                    '-------------------BARCODE

                    '--------------------ITEM CODE SELCTIONN------------------

                    'If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                    '    If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                    '    If oCFL.UniqueID = "CFL_ITM" Then
                    '        Try

                    '            oDS1.SetValue("U_VITMCDE", oDS1.Offset, oDT.GetValue("ItemCode", 0))

                    '            Dim GetBarCode As String = " SELECT Top 1 ""BcdCode"" FROM OBCD WHERE  " &
                    '                                          "  ""ItemCode"" = '" & oDT.GetValue("ItemCode", 0) & "' ORDER BY ""BcdEntry"" ASC "
                    '            Dim oRsGetBarCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '            oRsGetBarCode.DoQuery(GetBarCode)
                    '            Dim Barcode As String = oRsGetBarCode.Fields.Item("BcdCode").Value.ToString()

                    '            'Dim Barcode As String = oDT.GetValue("CodeBars", 0)
                    '            'Dim FromDate As String = oDS.GetValue("U_FRMDATE", 0)
                    '            'Dim ToDate As String = oDS.GetValue("U_TODATE", 0)
                    '            oDS1.SetValue("LineId", oDS1.Offset, pVal.Row)
                    '            'objMatrix.AddRow()

                    '            oDS1.SetValue("U_BARC", oDS1.Offset, Barcode)

                    '            'oDS1.SetValue("Code", oDS1.Offset, oDT.GetValue("CodeBars", 0))
                    '            oDS1.SetValue("U_VITMCDE", oDS1.Offset, oDT.GetValue("ItemCode", 0))
                    '            oDS1.SetValue("U_ITEMN", oDS1.Offset, oDT.GetValue("ItemName", 0))
                    '            oDS1.SetValue("U_UOM", oDS1.Offset, oDT.GetValue("InvntryUom", 0))
                    '            'oDS1.SetValue("U_FRMDATE", oDS1.Offset, oDS.GetValue("U_FRMDATE", 0))
                    '            'oDS1.SetValue("U_TODATE", oDS1.Offset, oDS.GetValue("U_TODATE", 0))
                    '            'objMatrix.SetLineData(pVal.Row)
                    '            Dim Itemcode As String = objMatrix.Columns.Item("VITMCDE").Cells.Item(pVal.Row).Specific.Value
                    '            Barcode = objMatrix.Columns.Item("BARC").Cells.Item(pVal.Row).Specific.Value
                    '            Dim query As String = "Select ""U_CNWhse"" from OADM"
                    '            Dim orset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '            orset.DoQuery(query)
                    '            Dim Warehouse As String = orset.Fields.Item("U_CNWhse").Value
                    '            Dim Brch As String = objForm.Items.Item("BPLID").Specific.Value
                    '            Try
                    '                Dim query1 As String = "CALL ""GetItemPrice""('" & Itemcode & "','" & Brch.Trim & "')"
                    '                Dim orset1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                orset1.DoQuery(query1)
                    '                If orset1.RecordCount > 0 Then
                    '                    Dim Itemcost As Double = orset1.Fields.Item("AvgPrice").Value
                    '                    oDS1.SetValue("U_ITMCS", oDS1.Offset, Itemcost)
                    '                    'objMatrix.SetLineData(pVal.Row)
                    '                Else
                    '                    oDS1.SetValue("U_ITMCS", oDS1.Offset, 0)
                    '                End If
                    '            Catch ex As Exception
                    '            End Try

                    '            Try
                    '                Dim query2 As String = "CALL ""GetRetailPrice""('" & Barcode & "','" & Brch.Trim & "')"
                    '                Dim orset2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                orset2.DoQuery(query2)
                    '                If orset2.RecordCount > 0 Then
                    '                    Dim Retailprice As Double = orset2.Fields.Item("Price").Value
                    '                    oDS1.SetValue("U_RETPC", oDS1.Offset, Retailprice)
                    '                    'objMatrix.SetLineData(pVal.Row)
                    '                Else
                    '                    oDS1.SetValue("U_RETPC", oDS1.Offset, 0)
                    '                End If
                    '            Catch ex As Exception
                    '            End Try

                    '            objMatrix.SetLineData(pVal.Row)
                    '            'coment
                    '            If pVal.Row = objMatrix.VisualRowCount Then
                    '                Me.SetNewLine(objForm.UniqueID, "IMtx")
                    '            End If
                    '        Catch ex As Exception
                    '            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    '        End Try
                    '    End If

                    'End If

                    '----------------------END SELECTION----------


                    '------------------COMMENTD---------
                    'If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                    '    If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                    '    If oCFL.UniqueID = "CFL_ITM" Then
                    '        oDS1.SetValue("LineId", oDS1.Offset, pVal.Row)
                    '        oDS1.SetValue("U_VITMCDE", oDS1.Offset, oDT.GetValue("ItemCode", 0))
                    '        oDS1.SetValue("U_ITEMN", oDS1.Offset, oDT.GetValue("ItemName", 0))

                    '        'oDS1.SetValue("U_UOM", oDS1.Offset, oDT.GetValue("InvntryUom", 0))
                    '        objMatrix.SetLineData(pVal.Row)
                    '        If pVal.Row = objMatrix.VisualRowCount Then
                    '            Me.SetNewLine(objForm.UniqueID, "ITMMtx")
                    '        End If
                    '    End If




                    '    If oCFL.UniqueID = "CFL_VRI" Then
                    '        oDS2.SetValue("LineId", oDS2.Offset, pVal.Row)
                    '        oDS2.SetValue("U_ITMCDE", oDS2.Offset, oDT.GetValue("ItemCode", 0))
                    '        oDS2.SetValue("U_ITMDESC", oDS2.Offset, oDT.GetValue("ItemName", 0))
                    '        objMatrix1.SetLineData(pVal.Row)
                    '        If pVal.Row = objMatrix1.VisualRowCount Then
                    '            Me.SetNewLine(objForm.UniqueID, "VRMtx")
                    '        End If


                    '    End If
                    'End If

                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        End If

                        If oCFL.UniqueID = "CFL_ITM" Then
                            Try
                                oDS1.SetValue("U_VITMCDE", oDS1.Offset, oDT.GetValue("ItemCode", 0))
                                Dim GetBarCode As String = "SELECT TOP 1 ""BcdCode"" FROM OBCD WHERE ""ItemCode"" = '" & oDT.GetValue("ItemCode", 0) & "' ORDER BY ""BcdEntry"" ASC"
                                Dim oRsGetBarCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetBarCode.DoQuery(GetBarCode)
                                Dim Barcode As String = ""
                                If oRsGetBarCode.RecordCount > 0 Then
                                    Barcode = oRsGetBarCode.Fields.Item("BcdCode").Value.ToString()

                                End If

                                oDS1.SetValue("LineId", oDS1.Offset, pVal.Row)

                                oDS1.SetValue("U_BARC", oDS1.Offset, Barcode)
                                oDS1.SetValue("U_VITMCDE", oDS1.Offset, oDT.GetValue("ItemCode", 0))
                                oDS1.SetValue("U_ITEMN", oDS1.Offset, oDT.GetValue("ItemName", 0))
                                'oDS1.SetValue("U_UOM", oDS1.Offset, oDT.GetValue("InvntryUom", 0))



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







                                ' Use values directly from oDT instead of relying on matrix before SetLineData

                                Dim Itemcode As String = oDT.GetValue("ItemCode", 0).ToString()

                                ' Only overwrite barcode if matrix cell has a valid (non-empty) value

                                Dim matrixBarcode As String = objMatrix.Columns.Item("BARC").Cells.Item(pVal.Row).Specific.Value

                                If Not String.IsNullOrWhiteSpace(matrixBarcode) Then

                                    Barcode = matrixBarcode

                                End If

                                Dim query As String = "SELECT ""U_CNWhse"" FROM OADM"

                                Dim orset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                orset.DoQuery(query)

                                Dim Warehouse As String = ""

                                If orset.RecordCount > 0 Then

                                    Warehouse = orset.Fields.Item("U_CNWhse").Value.ToString()

                                End If

                                Dim Brch As String = objForm.Items.Item("BPLID").Specific.Value

                                Try

                                    Dim query1 As String = "CALL ""GetItemPrice""('" & Itemcode & "','" & Brch.Trim & "')"

                                    Dim orset1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    orset1.DoQuery(query1)
                                    If orset1.RecordCount > 0 Then
                                        Dim Itemcost As Double = 0
                                        If Not IsDBNull(orset1.Fields.Item("AvgPrice").Value) Then
                                            Itemcost = Convert.ToDouble(orset1.Fields.Item("AvgPrice").Value)
                                        End If
                                        oDS1.SetValue("U_ITMCS", oDS1.Offset, Itemcost)
                                    Else
                                        oDS1.SetValue("U_ITMCS", oDS1.Offset, 0)
                                    End If
                                Catch ex As Exception

                                End Try

                                Try
                                    Dim query2 As String = "CALL ""GetRetailPrice""('" & Barcode & "','" & Brch.Trim & "')"
                                    Dim orset2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    orset2.DoQuery(query2)
                                    If orset2.RecordCount > 0 Then
                                        Dim Retailprice As Double = 0
                                        If Not IsDBNull(orset2.Fields.Item("Price").Value) Then
                                            Retailprice = Convert.ToDouble(orset2.Fields.Item("Price").Value)
                                        End If
                                        oDS1.SetValue("U_RETPC", oDS1.Offset, Retailprice)
                                    Else
                                        oDS1.SetValue("U_RETPC", oDS1.Offset, 0)
                                    End If
                                Catch ex As Exception
                                End Try
                                objMatrix.SetLineData(pVal.Row)
                                If pVal.Row = objMatrix.VisualRowCount Then
                                    Me.SetNewLine(objForm.UniqueID, "ITMMtx")
                                End If
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                            End Try

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
            objForm = objMain.objApplication.Forms.GetForm("PRMONVP", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                            'If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            '    Dim status As String = objForm.Items.Item("STATUS").Specific.Value
                            '    If status = "Active" Then
                            '        objForm.Items.Item("FRMDATE").Enabled = True
                            '        objForm.Items.Item("TODATE").Enabled = True
                            '        objForm.Items.Item("REDMNBAMNT").Enabled = True
                            '        objForm.Items.Item("FRMTIME").Enabled = True
                            '        objForm.Items.Item("TOTIME").Enabled = True
                            '        objForm.Items.Item("PName").Enabled = True
                            '        objForm.Items.Item("ITMMtx").Enabled = True
                            '        objForm.Items.Item("VRMtx").Enabled = True
                            '        objForm.Items.Item("BMtx").Enabled = True
                            '    ElseIf status = "Inactive" Then
                            '        objForm.Items.Item("FRMDATE").Enabled = False
                            '        objForm.Items.Item("TODATE").Enabled = False
                            '        objForm.Items.Item("REDMNBAMNT").Enabled = False
                            '        objForm.Items.Item("FRMTIME").Enabled = False
                            '        objForm.Items.Item("TOTIME").Enabled = False
                            '        objForm.Items.Item("PName").Enabled = False
                            '        objForm.Items.Item("ITMMtx").Enabled = False
                            '        objForm.Items.Item("VRMtx").Enabled = False
                            '        objForm.Items.Item("BMtx").Enabled = False
                            '    End If

                            'End If
                        End If
                    End If
                    'Akhila .B
                    Dim objComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("DTYPE").Specific
                    Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("ITMMtx").Specific

                    If objComboBox.Value = "P" Then
                        objMatrix.Columns.Item("DPRCTG").Editable = True
                        objForm.Items.Item("DISAT").Enabled = False
                        objForm.Items.Item("DISPER").Enabled = True
                        objMatrix.Columns.Item("DBAMT").Editable = False

                    ElseIf objComboBox.Value = "A" Then
                        objMatrix.Columns.Item("DBAMT").Editable = True
                        objMatrix.Columns.Item("DPRCTG").Editable = False
                        objForm.Items.Item("DISAT").Enabled = True
                        objForm.Items.Item("DISPER").Enabled = False
                    End If

                    '-------------


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
    Public Function GetNextDocNum()
        Dim Str As String
        Dim oRs As SAPbobsCOM.Recordset
        Dim DocNum As Integer
        oRs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Str = "SELECT IFNULL(MAX(""DocNum""),0)+1 as ""DocNum"" FROM ""@TNX_PONVP"""
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

    'Sub SetDefault(ByVal FormUID As String)'madhu push
    Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "", Optional ByVal Series As Integer = 0)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDS = Me.objForm.DataSources.DBDataSources.Add("@TNX_PONVP")
            oDS1 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PONVP_C0")
            oDS2 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PONVP_C1")
            oDS3 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PONVP_C2")
            objMatrix = objForm.Items.Item("ITMMtx").Specific
            objMatrix1 = objForm.Items.Item("VRMtx").Specific
            objMatrix2 = objForm.Items.Item("BMtx").Specific
            objForm.Freeze(True)

            'DOC NUMBER
            'oDS.SetValue("DocNum", oDS.Offset, objMain.objUtilities.GetNextDocNum(objForm, "ME_PRMONVPUDO", "Primary"))
            'objForm.DataBrowser.BrowseBy = "DocNum"
            oDS.SetValue("U_STATUS", oDS.Offset, "Active")
            'oDS.SetValue("U_DOCDATE", oDS.Offset, DateTime.Now.ToString("yyyyMMdd"))

            'END DOC 


            ' Load series combo box
            Dim cmbSeries As SAPbouiCOM.ComboBox = objForm.Items.Item("24").Specific
            Dim seriesQuery As String = "SELECT ""Series"", ""SeriesName"" FROM NNM1 WHERE ""ObjectCode"" = 'ME_PRMONVPUDO' AND ""Locked"" = 'N'"
            objMain.objUtilities.ComboBoxLoadValuesSeries(cmbSeries, seriesQuery)
            objForm.Items.Item("24").DisplayDesc = True
            cmbSeries.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            ' Get default series
            '            Dim defaultSeriesQuery As String = "
            'SELECT T1.""Series"", T1.""SeriesName""
            'FROM ONNM T0 
            'JOIN NNM1 T1 ON T1.""ObjectCode"" = T0.""ObjectCode"" AND T1.""Series"" = T0.""DfltSeries"" 
            'WHERE T0.""ObjectCode"" IN ('ME_PRMONVPUDO','@TNX_PONVP')"

            Dim defaultSeriesQuery As String =
    "SELECT ""Series"", ""SeriesName"", ""ObjectCode"", ""Indicator"" " &
    "FROM NNM1 " &
    "WHERE ""ObjectCode"" IN ('ME_PRMONVPUDO', '@TNX_PONVP') " &
    "AND ""Indicator"" <> 'Default'"
            Dim oRsDflt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsDflt.DoQuery(defaultSeriesQuery)

            If Flag = "" Then
                Dim nextDocNum = objMain.objUtilities.GetNextDocNumSeries(objForm, "ME_PRMONVPUDO", oRsDflt.Fields.Item("Series").Value)
                oDS.SetValue("DocNum", oDS.Offset, nextDocNum)
                oDS.SetValue("U_PERSER", oDS.Offset, oRsDflt.Fields.Item("Series").Value)
            Else
                Dim nextDocNum = objMain.objUtilities.GetNextDocNumSeries(objForm, "ME_PRMONVPUDO", Series)
                oDS.SetValue("DocNum", oDS.Offset, nextDocNum)
                oDS.SetValue("U_PERSER", oDS.Offset, Series)
            End If

            oDS.SetValue("U_DOCDATE", oDS.Offset, DateTime.Now.ToString("yyyyMMdd"))





            'end series






            ' CODE = oDS.GetValue("DocNum", 0)
            objMatrix.Clear()
            oDS1.Clear()
            objMatrix.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "ITMMtx")

            objMatrix.AutoResizeColumns()

            objMatrix1.Clear()
            oDS2.Clear()
            objMatrix1.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "VRMtx")

            objMatrix1.AutoResizeColumns()

            objMatrix2.Clear()
            oDS3.Clear()
            objMatrix2.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "BMtx")

            objMatrix2.AutoResizeColumns()
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal MatrixUID As String)
        Try
            oDS = Me.objForm.DataSources.DBDataSources.Add("@TNX_PONVP")
            oDS1 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PONVP_C0")
            oDS2 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PONVP_C1")
            oDS3 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PONVP_C2")
            objMatrix = objForm.Items.Item("ITMMtx").Specific
            objMatrix1 = objForm.Items.Item("VRMtx").Specific
            objMatrix2 = objForm.Items.Item("BMtx").Specific
            Select Case MatrixUID
                Case "ITMMtx"
                    objMatrix.AddRow()
                    oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
                    oDS1.SetValue("U_FRMAMNT", oDS1.Offset, "")
                    oDS1.SetValue("U_TOAMOUNT", oDS1.Offset, "")
                    oDS1.SetValue("U_VOHRORDIS", oDS1.Offset, "Voucher")
                    oDS1.SetValue("U_VITMCDE", oDS1.Offset, "")
                    oDS1.SetValue("U_ITEMN", oDS1.Offset, "")
                    oDS1.SetValue("U_ITMCS", oDS1.Offset, "")

                    oDS1.SetValue("U_RETPC", oDS1.Offset, "")
                    oDS1.SetValue("U_UOM", oDS1.Offset, "")

                    oDS1.SetValue("U_BARC", oDS1.Offset, "")

                    oDS1.SetValue("U_PERCNTG", oDS1.Offset, "")
                    objMatrix.AutoResizeColumns()
                    objMatrix.SetLineData(objMatrix.VisualRowCount)
                Case "VRMtx"
                    objMatrix1.AddRow()
                    oDS2.SetValue("LineId", oDS2.Offset, objMatrix1.VisualRowCount)
                    oDS2.SetValue("U_BARCD", oDS2.Offset, "")

                    oDS2.SetValue("U_ITMCDE", oDS2.Offset, "")
                    oDS2.SetValue("U_ITMDESC", oDS2.Offset, "")
                    objMatrix1.SetLineData(objMatrix1.VisualRowCount)
                    objMatrix1.AutoResizeColumns()
                Case "BMtx"
                    objMatrix2.AddRow()
                    oDS3.SetValue("LineId", oDS3.Offset, objMatrix2.VisualRowCount)
                    oDS3.SetValue("U_CHK", oDS3.Offset, "")
                    oDS3.SetValue("U_BCODE", oDS3.Offset, "")
                    oDS3.SetValue("U_BCHNME", oDS3.Offset, "")

                    objMatrix2.SetLineData(objMatrix2.VisualRowCount)
                    objMatrix2.AutoResizeColumns()
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub BranchLoad(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix2 = objForm.Items.Item("BMtx").Specific
            oDS3 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PONVP_C2")
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
    'Function BillAmount(ByVal FormUiD As String)
    '    If oDS.GetValue("U_REDMNBAMNT", 0) = 0 Then
    '        objMain.objApplication.StatusBar.SetText("Please Enter Redemption Minimum Bill Amount", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        Return False
    '    End If
    'End Function


    'Function BillAmount(ByVal FormUiD As String)
    '    If Convert.ToDecimal(oDS.GetValue("U_REDMNBAMNT", 0)) = 0 Then
    '        objMain.objApplication.StatusBar.SetText("Please Enter Redemption Minimum Bill Amount", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        Return False
    '    End If


    '    'If oDS.GetValue("U_REDMNBAMNT", 0) = "" Then
    '    '    objMain.objApplication.StatusBar.SetText("Please  Enter Redemption Minimum Bill Amount", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '    '    Return False
    '    'End If
    'End Function
    Public Function Validate(ByVal FormUID As String, ByVal MatrixUID As String) As Boolean
        Try


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
            'If oDS.GetValue("U_PNAME", 0) = "" Then
            '    objMain.objApplication.StatusBar.SetText("Please Enter Promotion Name", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            '    Exit Function
            'End If




            '' For j As Integer = 1 To objMatrix.VisualRowCount
            'If objMatrix.Columns.Item("FRMAMNT").Cells.Item(1).Specific.Value.ToString().Trim() <> "" Then
            '    objMain.objApplication.StatusBar.SetText("Please select the From Amount", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False

            'End If
            '' Next

            ''  For j As Integer = 1 To objMatrix.VisualRowCount
            'If objMatrix.Columns.Item("TOAMOUNT").Cells.Item(1).Specific.Value.ToString().Trim() <> "" Then
            '    objMain.objApplication.StatusBar.SetText("Please select the tO Amount", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If
            ''Next


            objMatrix = objForm.Items.Item("ITMMtx").Specific

            '  objMatrix1 = objForm.Items.Item("IGMtx1").Specific
            'objMatrix2 = objForm.Items.Item("SHMtx").Specific
            Select Case MatrixUID


                Case "ITMMtx"
                    'Dim type As String = objForm.Items.Item("PRTYP").Specific.Value
                    'If type = "ItemWise" Then
                    For i As Integer = 1 To objMatrix.VisualRowCount
                        If objMatrix.Columns.Item("VITMCDE").Cells.Item(1).Specific.Value.ToString().Trim() = "" Then
                            'objMain.objApplication.StatusBar.SetText("Please Choose Different selection And press Load Button", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            objMain.objApplication.StatusBar.SetText("Please Choose Voucher Item Code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                        'If objMatrix.Columns.Item("ITMNME").Cells.Item(i).Specific.Value.ToString().Trim() = "" Then
                        '    objMain.objApplication.StatusBar.SetText("Please Enter Item name", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    Return False
                        'End If
                        ''Dim objchk As Integer = objMatrix.Columns.Item("SHMtx").Cells.Item(i).Specific.Value.ToString().Trim()
                        'If objMatrix.Columns.Item("COLOR").Cells.Item(i).Specific.Value.ToString().Trim() = "" Then
                        '    objMain.objApplication.StatusBar.SetText("Please Enter Color", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    Return False
                        'End If
                        'If objMatrix.Columns.Item("SIZE").Cells.Item(i).Specific.Value.ToString().Trim() = "" Then
                        '    objMain.objApplication.StatusBar.SetText("Please Enter Size", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    Return False
                        'End If
                        'If objMatrix.Columns.Item("BARCDE").Cells.Item(i).Specific.Value.ToString().Trim() = "" Then
                        '    objMain.objApplication.StatusBar.SetText("Please Enter Barcode", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    Return False
                        'End If
                        'Dim objchk As Integer = 0
                        'If objMatrix.Columns.Item("CWHINST").Cells.Item(i).Specific.Value.ToString().Trim() = objchk Then
                        '    objMain.objApplication.StatusBar.SetText("Please Enter Quantity In Central Wharehouse In Stock", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    Return False
                        'End If

                    Next
                    '  End If

                    'Case "IGMtx1"

                    '    Dim type As String = objForm.Items.Item("PRTYP").Specific.Value
                    '    If type = "ItemGroupWise" Then
                    '        For j As Integer = 1 To objMatrix1.VisualRowCount
                    '            Dim check As Integer = objMatrix1.Columns.Item("DPERC").Cells.Item(j).Specific.Value.ToString().Trim()
                    '            Dim objchk As Integer = 0
                    '            If objMatrix1.Columns.Item("DPERC").Cells.Item(j).Specific.Value.ToString().Trim() = objchk Then
                    '                objMain.objApplication.StatusBar.SetText("Please Enter Discount % greater than zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '                Return False
                    '            End If
                    '            Dim oCheckBox As SAPbouiCOM.CheckBox

                    '            oCheckBox = objMatrix1.Columns.Item("CHK").Cells.Item(j).Specific
                    '            For k As Integer = 1 To objMatrix1.VisualRowCount
                    '                oCheckBox = objMatrix1.Columns.Item("CHK").Cells.Item(k).Specific
                    '                If oCheckBox.Checked = True Then
                    '                    Exit For
                    '                End If
                    '            Next
                    '            If oCheckBox.Checked = False Then
                    '                ' If objMatrix1.Columns.Item("CHK").Cells.Item(j).Specific.Value.ToString().Trim() = check Then
                    '                objMain.objApplication.StatusBar.SetText("Please Select Atleast one Checkbox", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '                Return False
                    '                'End If
                    '            End If

                    '        Next

                    '    End If

                    For j As Integer = 1 To objMatrix2.VisualRowCount
                        'Dim check As Integer = objMatrix1.Columns.Item("DPERC").Cells.Item(j).Specific.Value.ToString().Trim()
                        'Dim objchk As Integer = 0
                        'If objMatrix1.Columns.Item("DPERC").Cells.Item(j).Specific.Value.ToString().Trim() = objchk Then
                        '    objMain.objApplication.StatusBar.SetText("Please Enter Discount % greater than zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    Return False
                        'End If
                        Dim oCheckBox As SAPbouiCOM.CheckBox

                        oCheckBox = objMatrix2.Columns.Item("CHK").Cells.Item(j).Specific
                        For k As Integer = 1 To objMatrix2.VisualRowCount
                            oCheckBox = objMatrix2.Columns.Item("CHK").Cells.Item(k).Specific
                            If oCheckBox.Checked = True Then
                                Exit For
                            End If
                        Next
                        If oCheckBox.Checked = False Then
                            ' If objMatrix1.Columns.Item("CHK").Cells.Item(j).Specific.Value.ToString().Trim() = check Then
                            objMain.objApplication.StatusBar.SetText("Please Select Atleast one Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                            'End If
                        End If

                    Next


            End Select
            Return True

        Catch ex As Exception

        End Try
    End Function


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

#Region " Right Click Event"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)

        Dim objForm As SAPbouiCOM.Form
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus

        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)
        oDS = Me.objForm.DataSources.DBDataSources.Add("@TNX_PONVP")
        oDS1 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PONVP_C0")
        oDS2 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PONVP_C1")
        oDS3 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PONVP_C2")
        objMatrix = objForm.Items.Item("ITMMtx").Specific
        objMatrix1 = objForm.Items.Item("VRMtx").Specific
        objMatrix2 = objForm.Items.Item("BMtx").Specific

        Try
            If eventInfo.FormUID = objForm.UniqueID Then
                If (eventInfo.BeforeAction = True) Then
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then 'SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                        '  If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Or objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then

                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then 'SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                            '  objMatrix = objForm.Items.Item("Matx").Specific
                            objForm.EnableMenu("1292", True)
                            'objForm.EnableMenu("1293", True)
                            objForm.EnableMenu("774", True)


                        End If
                    End If
                End If
            End If
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

    Private Function ValidatePromotionName(ByVal formUID As String) As Boolean
        Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(formUID)
        Dim oEditText As SAPbouiCOM.EditText = objForm.Items.Item("PName").Specific
        Dim promoName As String = oEditText.Value.Trim()
        '  Max 10 characters
        If promoName.Length > 10 Then
            'objMain.objApplication.MessageBox("Promotion Name must be a maximum of 10 characters.")
            objMain.objApplication.StatusBar.SetText("Promotion Name must be a maximum of 10 characters.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        ' No spaces allowed
        If promoName.Contains(" ") Then
            'objMain.objApplication.MessageBox("Promotion Name Should not contain spaces.")
            objMain.objApplication.StatusBar.SetText("Promotion Name Should not contain spaces", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If

        '  Must be unique — check in the database
        Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs.DoQuery("SELECT COUNT(*) AS Cnt FROM ""@TNX_PONVP"" WHERE ""U_PName""= '" & promoName.Replace("'", "''") & "'")
        If rs.Fields.Item("Cnt").Value > 0 Then
            'objMain.objApplication.MessageBox("Promotion Name already exists. Please enter a unique name.")
            objMain.objApplication.StatusBar.SetText("Promotion Name already exists. Please enter a unique name.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        Return True
    End Function
    Function validateDiscountType(ByVal FormUiD As String)
        Try

            objForm = objMain.objApplication.Forms.Item(FormUiD)
            If objForm.Items.Item("DTYPE").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText(" Discount Type Cannot be Null", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True
        Catch ex As Exception

        End Try
    End Function
    Private Sub LoadSubGroups1(objForm As SAPbouiCOM.Form, itemGroup As String)

        If String.IsNullOrEmpty(itemGroup) OrElse Not IsNumeric(itemGroup.Trim()) Then

            'objMain.objApplication.MessageBox("Invalid or empty Item Group. Exiting.")
            Exit Sub
        End If

        Dim objCombo1 As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("SUBGRP").Specific, SAPbouiCOM.ComboBox)

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
        Dim objCombo As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("SSBGRP").Specific, SAPbouiCOM.ComboBox)
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



    'Function validateDiscountpercentage(ByVal FormUiD As String)
    '    Try
    '        Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUiD)

    '        Dim oCombo As SAPbouiCOM.ComboBox = CType(oForm.Items.Item("DTYPE").Specific, SAPbouiCOM.ComboBox)
    '        If oCombo.Selected IsNot Nothing AndAlso oCombo.Selected.Value = "P" Then
    '            For i As Integer = 1 To objMatrix.VisualRowCount
    '                Dim valueStr As String = objMatrix.Columns.Item("DPRCTG").Cells.Item(i).Specific.Value.ToString().Trim()
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

    Function validateDiscountpercentage(ByVal FormUiD As String)
        Try
            Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUiD)
            Dim oCombo As SAPbouiCOM.ComboBox = CType(oForm.Items.Item("DTYPE").Specific, SAPbouiCOM.ComboBox)
            If oCombo.Selected IsNot Nothing AndAlso oCombo.Selected.Value = "P" Then
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                    Dim itemCode As String = objMatrix.Columns.Item("VITMCDE").Cells.Item(i).Specific.Value.ToString().Trim()
                    If itemCode <> "" Then
                        Dim valueStr As String = objMatrix.Columns.Item("DPRCTG").Cells.Item(i).Specific.Value.ToString().Trim()
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
    Function validateDiscountAmount(ByVal FormUiD As String)
        Try
            Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUiD)
            Dim oCombo As SAPbouiCOM.ComboBox = CType(oForm.Items.Item("DTYPE").Specific, SAPbouiCOM.ComboBox)
            If oCombo.Selected IsNot Nothing AndAlso oCombo.Selected.Value = "A" Then
                For i As Integer = 1 To objMatrix.VisualRowCount
                    Dim itemCode As String = objMatrix.Columns.Item("VITMCDE").Cells.Item(i).Specific.Value.ToString().Trim()
                    If itemCode <> "" Then
                        Dim valueStr As String = objMatrix.Columns.Item("DBAMT").Cells.Item(i).Specific.Value.ToString().Trim()
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


    Public Function Duplicatebarcode(ByVal FormUID As String, ByVal MatrixUID As String) As Boolean
        Try
            For ij As Integer = 1 To objMatrix.VisualRowCount



                Dim Barcode1 As String = objMatrix.Columns.Item("BARC").Cells.Item(ij).Specific.Value
                For ik As Integer = 1 To objMatrix.VisualRowCount
                    If ij <> ik Then
                        If Barcode1 = objMatrix.Columns.Item("BARC").Cells.Item(ik).Specific.Value Then
                            objMain.objApplication.StatusBar.SetText("Duplicate Barcode " & Barcode1 & " Exists in row number " & ik & "...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                Next

            Next
            Dim FromDate As String = oDS.GetValue("U_FRMDATE", 0)
            Dim ToDate As String = oDS.GetValue("U_TODATE", 0)
            Dim Barcode As String = ""
            For ijk As Integer = 1 To objMatrix.VisualRowCount

                Barcode = objMatrix.Columns.Item("BARC").Cells.Item(ijk).Specific.Value
                '              Dim Qry As String = "select distinct T0.""U_BARC"",T1.""U_FRMDATE"", T1.""U_TODATE"", T1.""DocNum"" " &
                '"from ""@TNX_PONVP_C0"" T0 inner join ""@TNX_PONVP"" T1 on T0.""DocEntry""=T1.""DocEntry"" " &
                '" where   (T1.""U_FRMDATE"" between '" & FromDate & "' and '" & ToDate & "' " &
                '" and T0.""U_BARC""<> '' And T1.""U_STATUS""='Active' and T1.""DocNum"" <> '" & objForm.Items.Item("DocNum").Specific.Value & "' ) or " &
                '"(T1.""U_TODATE"" between '" & FromDate & "' and '" & ToDate & "' and T0.""U_BARC""<> '' And T1.""U_STATUS""='Active' and T1.""DocNum"" <> '" & objForm.Items.Item("DocNum").Specific.Value & "' )"


                Dim Qry As String = "select distinct T0.""U_BARC"", T1.""U_FRMDATE"", T1.""U_TODATE"", T1.""DocNum"" " &
"from ""@TNX_PONVP_C0"" T0 inner join ""@TNX_PONVP"" T1 on T0.""DocEntry"" = T1.""DocEntry"" " &
"where (T1.""U_FRMDATE"" between '" & FromDate & "' and '" & ToDate & "' " &
"and T0.""U_BARC"" <> '' And T1.""U_STATUS"" = 'Active' and T1.""DocNum"" <> '" & objForm.Items.Item("DocNum").Specific.Value & "') or " &
"(T1.""U_TODATE"" between '" & FromDate & "' and '" & ToDate & "' " &
"and T0.""U_BARC"" <> '' And T1.""U_STATUS"" = 'Active' and T1.""DocNum"" <> '" & objForm.Items.Item("DocNum").Specific.Value & "')"

                Dim rs2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim chkBarCode As String
                Dim chkDocNum As String
                Dim chkFromDate As String
                Dim chkToDate As String

                rs2.DoQuery(Qry)
                If rs2.RecordCount > 0 Then
                    For i As Integer = 1 To rs2.RecordCount
                        chkBarCode = rs2.Fields.Item("U_BARC").Value
                        chkDocNum = rs2.Fields.Item("DocNum").Value
                        chkFromDate = Format(CType(rs2.Fields.Item("U_FRMDATE").Value, Date), "yyyyMMdd")
                        chkToDate = Format(CType(rs2.Fields.Item("U_TODATE").Value, Date), "yyyyMMdd")
                        If chkBarCode = Barcode And ((chkFromDate >= FromDate Or chkFromDate <= ToDate) Or (chkToDate >= FromDate Or chkToDate <= ToDate)) Then
                            objMain.objApplication.StatusBar.SetText("This Barcode  " & Barcode & "  Already Exist in DocNum " & chkDocNum & "  between " & rs2.Fields.Item("U_FRMDATE").Value & "  to  " & rs2.Fields.Item("U_TODATE").Value & "  ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                        rs2.MoveNext()

                    Next
                End If
            Next
            Return True

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function




    Function validateVocherAmount(ByVal FormUiD As String)
        Try
            Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUiD)
            Dim oCombo As SAPbouiCOM.ComboBox = CType(objMatrix.Columns.Item("VOHRORDIS").Cells.Item(1).Specific, SAPbouiCOM.ComboBox)

            'Dim oCombo As SAPbouiCOM.ComboBox = CType(oForm.Items.Item("DTYPE").Specific, SAPbouiCOM.ComboBox)
            If oCombo.Selected IsNot Nothing AndAlso oCombo.Selected.Value = "Voucher" Then
                For i As Integer = 1 To objMatrix.VisualRowCount
                    Dim itemCode As String = objMatrix.Columns.Item("VITMCDE").Cells.Item(i).Specific.Value.ToString().Trim()
                    If itemCode <> "" Then
                        Dim valueStr As String = objMatrix.Columns.Item("PERORAMONT").Cells.Item(i).Specific.Value.ToString().Trim()
                        If valueStr = "" OrElse CDbl(valueStr) = 0 Then
                            objMain.objApplication.StatusBar.SetText("Please enter a valid Voucher Amount in row " & i & " for item " & itemCode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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




    Function validateVocherpercentage(ByVal FormUiD As String)
        Try
            Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUiD)
            Dim oCombo As SAPbouiCOM.ComboBox = CType(objMatrix.Columns.Item("VOHRORDIS").Cells.Item(1).Specific, SAPbouiCOM.ComboBox)

            'Dim oCombo As SAPbouiCOM.ComboBox = CType(oForm.Items.Item("DTYPE").Specific, SAPbouiCOM.ComboBox)
            If oCombo.Selected IsNot Nothing AndAlso oCombo.Selected.Value = "Discount" Then
                For i As Integer = 1 To objMatrix.VisualRowCount
                    Dim itemCode As String = objMatrix.Columns.Item("VITMCDE").Cells.Item(i).Specific.Value.ToString().Trim()
                    If itemCode <> "" Then
                        Dim valueStr As String = objMatrix.Columns.Item("PERCNTG").Cells.Item(i).Specific.Value.ToString().Trim()
                        If valueStr = "" OrElse CDbl(valueStr) = 0 Then
                            objMain.objApplication.StatusBar.SetText("Please enter a valid Percentage in row " & i & " for item " & itemCode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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


























    'Function validateAmounts(ByVal FormUiD As String) As Boolean
    '    Try
    '        Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUiD)

    '        For i As Integer = 1 To objMatrix.VisualRowCount
    '            Dim fromValueStr As String = objMatrix.Columns.Item("FRMAMNT").Cells.Item(i).Specific.Value.ToString().Trim()
    '            Dim toValueStr As String = objMatrix.Columns.Item("TOAMOUNT").Cells.Item(i).Specific.Value.ToString().Trim()

    '            ' Skip the row if both From and To Amounts are empty
    '            If fromValueStr = "" AndAlso toValueStr = "" Then
    '                Continue For
    '            End If

    '            ' Validate From Amount
    '            If fromValueStr = "" OrElse Not IsNumeric(fromValueStr) OrElse CDbl(fromValueStr) = 0 Then
    '                objMain.objApplication.StatusBar.SetText("Please enter a valid From Amount in row " & i,
    '                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                Return False
    '            End If

    '            ' Validate To Amount
    '            If toValueStr = "" OrElse Not IsNumeric(toValueStr) OrElse CDbl(toValueStr) = 0 Then
    '                objMain.objApplication.StatusBar.SetText("Please enter a valid To Amount in row " & i,
    '                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                Return False
    '            End If


    '        Next

    '        Return True
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message,
    '        SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        Return False
    '    End Try
    'End Function


    Function validateAmounts(ByVal FormUiD As String) As Boolean
        Try
            Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUiD)

            For i As Integer = 1 To objMatrix.VisualRowCount
                Dim itemCode As String = objMatrix.Columns.Item("VITMCDE").Cells.Item(i).Specific.Value.ToString().Trim()
                Dim fromValueStr As String = objMatrix.Columns.Item("FRMAMNT").Cells.Item(i).Specific.Value.ToString().Trim()
                Dim toValueStr As String = objMatrix.Columns.Item("TOAMOUNT").Cells.Item(i).Specific.Value.ToString().Trim()

                ' Skip validation if VITMCDE (Item Code) is empty
                If itemCode = "" Then
                    Continue For
                End If

                ' Validate From Amount
                If fromValueStr = "" OrElse Not IsNumeric(fromValueStr) OrElse CDbl(fromValueStr) = 0 Then
                    objMain.objApplication.StatusBar.SetText("Please enter a valid From Amount in row " & i,
                    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                ' Validate To Amount
                If toValueStr = "" OrElse Not IsNumeric(toValueStr) OrElse CDbl(toValueStr) = 0 Then
                    objMain.objApplication.StatusBar.SetText("Please enter a valid To Amount in row " & i,
                    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If


            Next

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function






End Class

