
'NEW
Imports System.Collections.Specialized
Imports System.Configuration
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Public Class ClsPromoReversion

#Region "       Declaration             "
    Public objForm, objSTRForm As SAPbouiCOM.Form
    Dim objMatrix, objMatrix1, objMatrix2 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Public oDS, oDS1, oDS2, oDS3 As SAPbouiCOM.DBDataSource
    Public PrgBar As SAPbouiCOM.ProgressBar
    Dim LostFocusFlag As Boolean
    Dim MATRIXS As String

#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("PromoCostReversion.xml", "REVERSION", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("REVERSION", objMain.objApplication.Forms.ActiveForm.TypeCount)
            'objForm.Freeze(True)
            objForm.Freeze(False)
            objForm.Items.Item("Item_12").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            objForm.EnableMenu("1292", True)
            objForm.EnableMenu("774", True)

            Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            Me.objForm.Items.Item("DOCDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Me.objForm.Items.Item("DOCDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Me.objForm.Items.Item("DOCDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Me.objForm.Items.Item("DOCDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            Me.objForm.Items.Item("POST").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Me.objForm.Items.Item("POST").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try

            If pVal.MenuUID = "ME_REVERSION" And pVal.BeforeAction = False Then
                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                Me.SetNewLine(objForm.UniqueID, "RVMtx1")

            ElseIf pVal.MenuUID = "774" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("REVERSION", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix = objForm.Items.Item("RVMtx").Specific

                Dim row As Integer = objMatrix.VisualRowCount
                If MATRIXS.Equals("RVMtx") = True Then
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

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If (pVal.ItemUID = "RVMtx") And pVal.BeforeAction = False Then
                        MATRIXS = pVal.ItemUID
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And
                    (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        'If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Me.SetDefault(objForm.UniqueID)
                    End If
                    If pVal.ItemUID = "Item_15" And pVal.BeforeAction = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or
                    pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                        Dim AppTemForm_EMD As SAPbouiCOM.Form
                        objMain.objApplication.ActivateMenuItem("2561")
                        AppTemForm_EMD = objMain.objApplication.Forms.GetFormByTypeAndCount(134, 1)
                        AppTemForm_EMD = objMain.objApplication.Forms.ActiveForm
                        AppTemForm_EMD.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        Dim sercall As String = objForm.Items.Item("SCODE").Specific.value
                        AppTemForm_EMD.Items.Item("5").Specific.Value = sercall
                        AppTemForm_EMD.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        AppTemForm_EMD = objMain.objApplication.Forms.ActiveForm
                    End If


                    If pVal.ItemUID = "Item_5" And pVal.BeforeAction = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or
                        pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                        Dim AppTemForm_EMD As SAPbouiCOM.Form
                        objMain.objApplication.ActivateMenuItem("2309")
                        AppTemForm_EMD = objMain.objApplication.Forms.GetFormByTypeAndCount(181, 1)
                        AppTemForm_EMD = objMain.objApplication.Forms.ActiveForm
                        AppTemForm_EMD.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        Dim CrediNo As String = objForm.Items.Item("CNTNO").Specific.value
                        AppTemForm_EMD.Items.Item("8").Specific.Value = CrediNo
                        AppTemForm_EMD.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        AppTemForm_EMD = objMain.objApplication.Forms.ActiveForm
                    End If


                    If pVal.ItemUID = "lnk1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or
                        pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                        Dim AppTemForm_EMD As SAPbouiCOM.Form
                        objMain.objApplication.ActivateMenuItem("ME_REDCTN")
                        ' AppTemForm_EMD = objMain.objApplication.Forms.GetFormByTypeAndCount("REDCTN", 1)
                        AppTemForm_EMD = objMain.objApplication.Forms.ActiveForm
                        AppTemForm_EMD.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        Dim Doc As String = objForm.Items.Item("PCRNO").Specific.value
                        AppTemForm_EMD.Items.Item("DocNum").Specific.Value = Doc
                        AppTemForm_EMD.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        AppTemForm_EMD = objMain.objApplication.Forms.ActiveForm
                    End If







                    If pVal.ItemUID = "POST" AndAlso pVal.BeforeAction = False AndAlso (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE OrElse pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE OrElse
                                     pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then

                        Dim errMsg As String = ""
                        Dim oMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("RVMtx").Specific
                        Dim selectedRow As Integer = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder)

                        ' If no row is selected, default to first data row
                        If selectedRow <= 0 Then
                            If oMatrix.RowCount > 0 Then
                                selectedRow = 1 ' First data row (not title row)
                            Else
                                objMain.objApplication.StatusBar.SetText("Matrix is empty. No data to process.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Return
                            End If
                        End If

                        ' Get the branch name from the selected row
                        Dim branchName As String = oMatrix.Columns.Item("BRANCH").Cells.Item(selectedRow).Specific.Value.ToString().Trim()

                        ' Call the function
                        Dim success As Boolean = Me.AddAPCreditNote(errMsg, branchName)

                        If success Then
                            objMain.objApplication.StatusBar.SetText("AP Credit Note posted successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        Else
                            objMain.objApplication.StatusBar.SetText("Failed to post AP Credit Note: " & errMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.validateSpplierCode(objForm.UniqueID) = False Then
                            BubbleEvent = False
                        End If
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.validateBuyerName(objForm.UniqueID) = False Then
                            BubbleEvent = False
                        End If
                    End If
                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                    '    If Me.validateDocumentReference(objForm.UniqueID, "RVMtx") = False Then
                    '        BubbleEvent = False
                    '    End If
                    'End If
                    If pVal.ItemUID = "Load" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Me.LoadDetails(pVal, objForm.UniqueID)
                    End If

                    If pVal.ItemUID = "PRMOIN" And pVal.BeforeAction = False Then
                        Me.AddPromocostReversionInventoryPost(objForm.UniqueID)
                    End If




                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    Try
                        If pVal.ItemUID = "RVMtx" AndAlso pVal.ColUID = "CSQUN" Then
                            Dim oMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("RVMtx").Specific
                            Dim oMatrix1 As SAPbouiCOM.Matrix = objForm.Items.Item("RVMtx1").Specific

                            Dim oRecordSet As SAPbobsCOM.Recordset
                            oRecordSet = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            Dim promoRevNo As String = CType(objForm.Items.Item("PCRNO").Specific, SAPbouiCOM.EditText).Value

                            Dim query1 As String = "SELECT R.""U_ISTQTY"", R.""U_ITMCODE"" " &
                               "FROM ""@TNX_RDCTN_C0"" R " &
                               "LEFT JOIN ""@TNX_RDCTN"" H ON H.""DocEntry"" = R.""DocEntry"" " &
                               "LEFT JOIN ""@TNX_REVERSION"" V ON V.""U_PCRNO"" = H.""DocNum"" " &
                               "WHERE H.""DocNum"" = '" & promoRevNo & "'"

                            oRecordSet.DoQuery(query1)

                            If oRecordSet.RecordCount > 0 Then
                                For i As Integer = 1 To oMatrix.RowCount
                                    Dim currentStockQtyStr As String = CType(oMatrix.Columns.Item("CSQUN").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value
                                    Dim itemCode As String = CType(oMatrix.Columns.Item("ICODE").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value

                                    Dim CurrentStockQuantity As Double = 0
                                    If Not String.IsNullOrEmpty(currentStockQtyStr) AndAlso IsNumeric(currentStockQtyStr) Then
                                        CurrentStockQuantity = CDbl(currentStockQtyStr)
                                    End If

                                    oRecordSet.MoveFirst()
                                    Dim itemInStockQty As Double = 0
                                    Dim foundMatch As Boolean = False

                                    While Not oRecordSet.EoF
                                        If oRecordSet.Fields.Item("U_ITMCODE").Value.ToString().Trim().ToUpper() = itemCode.Trim().ToUpper() Then
                                            If Not String.IsNullOrEmpty(oRecordSet.Fields.Item("U_ISTQTY").Value.ToString()) AndAlso IsNumeric(oRecordSet.Fields.Item("U_ISTQTY").Value.ToString()) Then
                                                itemInStockQty = CDbl(oRecordSet.Fields.Item("U_ISTQTY").Value)
                                            End If
                                            foundMatch = True
                                            Exit While
                                        End If
                                        oRecordSet.MoveNext()
                                    End While

                                    Dim SoldQty As Double = 0
                                    If foundMatch Then
                                        SoldQty = itemInStockQty - CurrentStockQuantity
                                    End If

                                    oMatrix1.Columns.Item("SQNTY").Cells.Item(i).Specific.Value = Math.Abs(SoldQty).ToString("F2")
                                Next
                                oMatrix1.AutoResizeColumns()
                            Else
                                objMain.objApplication.StatusBar.SetText("No data found for the given Promo Reversion Number.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If
                        End If
                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText("Error during calculation: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try

                    If pVal.ItemUID = "RVMtx1" AndAlso pVal.ColUID = "SQNTY" Then
                        Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("RVMtx1").Specific
                        For rowIndex As Integer = 1 To objMatrix.RowCount
                            Try
                                Dim soldQuantityStr As String = CType(objMatrix.Columns.Item("SQNTY").Cells.Item(rowIndex).Specific, SAPbouiCOM.EditText).Value
                                Dim oldCostStr As String = CType(objMatrix.Columns.Item("OCOST").Cells.Item(rowIndex).Specific, SAPbouiCOM.EditText).Value
                                Dim promoCostStr As String = CType(objMatrix.Columns.Item("PRCOST").Cells.Item(rowIndex).Specific, SAPbouiCOM.EditText).Value

                                Dim soldQuantity As Double = If(Not String.IsNullOrEmpty(soldQuantityStr) AndAlso IsNumeric(soldQuantityStr), CDbl(soldQuantityStr), 0)
                                Dim oldCost As Double = If(Not String.IsNullOrEmpty(oldCostStr) AndAlso IsNumeric(oldCostStr), CDbl(oldCostStr), 0)
                                Dim promoCost As Double = If(Not String.IsNullOrEmpty(promoCostStr) AndAlso IsNumeric(promoCostStr), CDbl(promoCostStr), 0)

                                Dim totalSalesAmount As Double = soldQuantity * oldCost - promoCost
                                'sign remove
                                objMatrix.Columns.Item("TSAMT").Cells.Item(rowIndex).Specific.Value = Math.Abs(totalSalesAmount).ToString("F2")
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText("Error calculating Total Sales Amount on row " & rowIndex & ": " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Try
                        Next
                        objMatrix.AutoResizeColumns()
                    End If

                    If pVal.ItemUID = "RVMtx1" AndAlso pVal.ColUID = "TSAMT" Then
                        Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("RVMtx1").Specific
                        Dim TotalSupplierCreditNoteAmmount As Double = 0

                        For rowIndex As Integer = 1 To objMatrix.RowCount
                            Try
                                Dim totalSalesAmountStr As String = CType(objMatrix.Columns.Item("TSAMT").Cells.Item(rowIndex).Specific, SAPbouiCOM.EditText).Value
                                Dim totalSalesAmount As Double = If(Not String.IsNullOrEmpty(totalSalesAmountStr) AndAlso IsNumeric(totalSalesAmountStr), CDbl(totalSalesAmountStr), 0)
                                TotalSupplierCreditNoteAmmount += totalSalesAmount
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText("Error summing TSAMT on row " & rowIndex & ": " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Try
                        Next

                        CType(objForm.Items.Item("TSCNAMT").Specific, SAPbouiCOM.EditText).Value = TotalSupplierCreditNoteAmmount.ToString("F2")
                    End If


                    'Akhila
                    'If pVal.ColUID = "CSTDI" And pVal.BeforeAction = False Then
                    '    Dim TotalCostDifference As Double = 0
                    '    For i As Integer = 1 To objMatrix.RowCount

                    '        ' Calculate cost difference
                    '        Dim costDiff As Double = Math.Abs(objMatrix.Columns.Item("CSTDI").Cells.Item(i).Specific.Value)
                    '        objMatrix.Columns.Item("CSTDI").Cells.Item(i).Specific.Value = costDiff.ToString("F2")
                    '        TotalCostDifference += costDiff

                    '    Next
                    '    oDS.SetValue("U_TOTAL", oDS.Offset, TotalCostDifference.ToString("F2"))
                    'End If






                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDS = objForm.DataSources.DBDataSources.Item("@TNX_REVERSION")
                    oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_REVERSION_C0")
                    oDS2 = objForm.DataSources.DBDataSources.Item("@TNX_REVERSION_C1")
                    objMatrix = objForm.Items.Item("RVMtx").Specific
                    objMatrix1 = objForm.Items.Item("RVMtx1").Specific
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                    '    If oCFL.UniqueID = "CFL_1" And pVal.BeforeAction = True Then 'CHOOSE FORM TABLE UNIQUNE ID
                    '        Me.CFLCardCodeFilter(objForm.UniqueID, "CFL_1")
                    '    End If
                    '    'If oCFL.UniqueID = "CFL_4" And pVal.BeforeAction = True Then 'CHOOSE FORM TABLE UNIQUNE ID
                    '    '    Me.CFLCardCodeFilter1(objForm.UniqueID, "CFL_4")
                    '    'End If
                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                        '        If oCFL.UniqueID = "CFL_1" Then
                        '            oDS.SetValue("U_SCODE", oDS.Offset, oDT.GetValue("CardCode", 0))
                        '            oDS.SetValue("U_SNAME", oDS.Offset, oDT.GetValue("CardName", 0))
                        '        End If

                        '        If oCFL.UniqueID = "CFL_2" Then
                        '            oDS1.SetValue("U_ICODE", oDS1.Offset, oDT.GetValue("ItemCode", 0))
                        '            oDS1.SetValue("U_ITDSP", oDS1.Offset, oDT.GetValue("ItemName", 0))
                        '            objMatrix.SetLineData(pVal.Row)
                        '        End If

                        '        If oCFL.UniqueID = "CFL_3" Then
                        '            oDS1.SetValue("U_GLACT", oDS1.Offset, oDT.GetValue("AcctCode", 0))
                        '            objMatrix.SetLineData(pVal.Row)
                        '        End If
                        '        If oCFL.UniqueID = "CFL_4" Then
                        '            oDS.SetValue("U_BNAME", oDS.Offset, oDT.GetValue("SlpName", 0))

                        '            Dim SalesEmpCode As String = oDT.GetValue("SlpCode", 0)

                        '            Dim rs As SAPbobsCOM.Recordset
                        '            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '            Dim query As String = "SELECT ""U_GLAccount"" FROM ""OSLP"" WHERE ""SlpCode"" = '" & SalesEmpCode & "'"
                        '            rs.DoQuery(query)

                        '            If Not rs.EoF Then
                        '                Dim glAccount As String = rs.Fields.Item("U_GLAccount").Value

                        '                objMatrix = objForm.Items.Item("RVMtx").Specific

                        '                If pVal.Row = 0 Then
                        '                    'objMatrix.AddRow()
                        '                End If
                        '                Dim lastRow As Integer
                        '                lastRow = objMatrix.VisualRowCount
                        '                oDS1.SetValue("U_GLACT", oDS1.Offset, glAccount)
                        '                objMatrix.SetLineData(lastRow)
                        '                ' objMatrix.FlushToDataSource()
                        '                ' Me.SetNewLine(objForm.UniqueID, "RVMtx")


                        '                'objMain.objApplication.StatusBar.SetText("GL Account auto-filled based on Buyer Name.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        '            Else
                        '                'objMain.objApplication.StatusBar.SetText("No GL Account found for the selected Buyer.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '            End If
                        '        End If


                        '        If oCFL.UniqueID = "CFL_5" Then
                        '            oDS1.SetValue("U_INRNO", oDS1.Offset, oDT.GetValue("DocEntry", 0))
                        '            objMatrix.SetLineData(pVal.Row)
                        '        End If
                        '        If oCFL.UniqueID = "CFL_6" Then
                        '            oDS1.SetValue("U_WRHOS", oDS1.Offset, oDT.GetValue("WhsCode", 0))
                        '            objMatrix.SetLineData(pVal.Row)
                        '        End If

                        If oCFL.UniqueID = "CFL_7" Then
                            oDS.SetValue("U_PCRNO", oDS.Offset, oDT.GetValue("DocNum", 0))
                        End If
                    End If

                    '    If pVal.ItemUID = "Branch" And pVal.BeforeAction = False Then
                    '        Dim objBranchCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Branch").Specific
                    '        Dim BranchID As String
                    '        Try
                    '            BranchID = objBranchCombo.Selected.Value
                    '        Catch ex As Exception
                    '            objMain.objApplication.StatusBar.SetText("Please Select Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        End Try
                    '        Dim WhsCmbo As SAPbouiCOM.ComboBox = DirectCast(Me.objForm.Items.Item("Whs").Specific, SAPbouiCOM.ComboBox)
                    '        Dim WhsQry As String = "Select T0.""WhsCode"", T0.""WhsName"" FROM OWHS T0 Where ""BPLid""='" & BranchID & "' And ""U_Transit"" = 'No' "
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

            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Sub SetDefault(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDS = Me.objForm.DataSources.DBDataSources.Add("@TNX_REVERSION")
            oDS1 = Me.objForm.DataSources.DBDataSources.Add("@TNX_REVERSION_C0")
            oDS2 = Me.objForm.DataSources.DBDataSources.Add("@TNX_REVERSION_C1")
            objMatrix = objForm.Items.Item("RVMtx").Specific
            objMatrix1 = objForm.Items.Item("RVMtx1").Specific
            objForm.Freeze(True)

            oDS.SetValue("DocNum", oDS.Offset, objMain.objUtilities.GetNextDocNum(objForm, "ME_REVERSIONUDO", "Primary"))
            'objForm.DataBrowser.BrowseBy = "DocNum"
            oDS.SetValue("U_DOCDATE", oDS.Offset, DateTime.Now.ToString("yyyyMMdd"))
            oDS.SetValue("U_STATUS", oDS.Offset, "Active")

            objMatrix.Clear()
            oDS1.Clear()
            objMatrix.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "RVMtx")
            objMatrix.AutoResizeColumns()

            objMatrix1.Clear()
            oDS2.Clear()
            objMatrix1.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "RVMtx1")
            objMatrix1.AutoResizeColumns()

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    Sub SetNewLine(ByVal FormUID As String, ByVal MatrixUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDS = Me.objForm.DataSources.DBDataSources.Add("@TNX_REVERSION")
            oDS1 = Me.objForm.DataSources.DBDataSources.Add("@TNX_REVERSION_C0")
            oDS2 = Me.objForm.DataSources.DBDataSources.Add("@TNX_REVERSION_C1")
            objMatrix = objForm.Items.Item("RVMtx").Specific
            objMatrix1 = objForm.Items.Item("RVMtx1").Specific

            Select Case MatrixUID
                Case "RVMtx"
                    objMatrix.AddRow()
                    oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
                    oDS1.SetValue("U_ICODE", oDS1.Offset, "")
                    oDS1.SetValue("U_CSTDI", oDS1.Offset, "")

                    oDS1.SetValue("U_ITDSP", oDS1.Offset, "")
                    oDS1.SetValue("U_WRHOS", oDS1.Offset, "")
                    oDS1.SetValue("U_CSQUN", oDS1.Offset, "")
                    oDS1.SetValue("U_PICOST", oDS1.Offset, "")
                    oDS1.SetValue("U_RICOST", oDS1.Offset, "")
                    'oDS1.SetValue("U_GLACT", oDS1.Offset, "")
                    oDS1.SetValue("U_JRRMK", oDS1.Offset, "")
                    oDS1.SetValue("U_REF1", oDS1.Offset, "")
                    oDS1.SetValue("U_REF2", oDS1.Offset, "")
                    oDS1.SetValue("U_BRANCH", oDS1.Offset, "")
                    oDS1.SetValue("U_RDREF", oDS1.Offset, "")
                    oDS1.SetValue("U_INRNO", oDS1.Offset, "")
                    oDS1.SetValue("U_PSTATUS", oDS1.Offset, "")

                    objMatrix.AutoResizeColumns()
                    objMatrix.SetLineData(objMatrix.VisualRowCount)
                Case "RVMtx1"
                    objMatrix1.AddRow()
                    oDS2.SetValue("LineID", oDS2.Offset, objMatrix1.VisualRowCount)
                    oDS2.SetValue("U_OCOST", oDS2.Offset, "")
                    oDS2.SetValue("U_PRCOST", oDS2.Offset, "")
                    oDS2.SetValue("U_TSAMT", oDS2.Offset, "")
                    'oDS2.SetValue("U_TSCNAMT", oDS2.Offset, "")
                    'oDS2.SetValue("U_CNTNO", oDS2.Offset, "")
                    oDS2.SetValue("U_SQNTY", oDS2.Offset, "")
                    'oDS2.SetValue("U_CNPBR", oDS2.Offset, "")
                    objMatrix1.SetLineData(objMatrix1.VisualRowCount)
                    objMatrix1.AutoResizeColumns()
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Function validateSpplierCode(ByVal FormUiD As String)
        Try

            objForm = objMain.objApplication.Forms.Item(FormUiD)
            If objForm.Items.Item("SCODE").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Supplier Code Can't be Null", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True
        Catch ex As Exception

        End Try
    End Function
    Function validateBuyerName(ByVal FormUiD As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUiD)
            If objForm.Items.Item("BNAME").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Buyer Name Can't be Null", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True
        Catch ex As Exception

        End Try
    End Function
    Sub CFLCardCodeFilter(ByVal FormUID As String, ByVal CFL_ID As String)
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
    Sub CFLCardCodeFilter1(ByVal FormUID As String, ByVal CFL_ID As String)
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

    Function validateDocumentReference(ByVal FormUID As String, ByVal MatrixUID As String) As Boolean

        Try
            objMatrix = objForm.Items.Item("RVMtx").Specific
            '  objMatrix1 = objForm.Items.Item("IGMtx1").Specific
            'objMatrix2 = objForm.Items.Item("SHMtx").Specific
            Select Case MatrixUID
                Case "RVMtx"
                    'Dim type As String = objForm.Items.Item("PRTYP").Specific.Value
                    'If type = "ItemWise" Then
                    For i As Integer = 1 To objMatrix.VisualRowCount
                        If objMatrix.Columns.Item("RDREF").Cells.Item(1).Specific.Value.ToString().Trim() = "" Then
                            objMain.objApplication.StatusBar.SetText("Revaluation Document Reference cannot be null", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    Next
            End Select
            Return True

        Catch ex As Exception

        End Try
    End Function


    'Posting Inventory revaluation--------------
    Public Function AddPromocostReversionInventoryPost(ByRef errorMessage As String) As Boolean
        errorMessage = ""
        Try
            Dim oMaterialRevaluation As SAPbobsCOM.MaterialRevaluation
            ' Header information
            Dim objrec As SAPbobsCOM.Recordset
            objrec = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '            Dim recqry As String = "SELECT T1.""U_INRNO"",T1.""LineId"", T0.""DocNum"", T1.""U_RICOST"",T1.""U_ICODE"",T1.""U_ITDSP"",T1.""U_JRRMK"" ,T1.""U_CSTDI"",T0.""U_GLACT"", T1.""U_REF1"",T1.""U_REF2"", T1.""U_WRHOS"" FROM ""@TNX_REVERSION"" T0
            'INNER JOIN ""@TNX_REVERSION_C0"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" WHERE T1.""U_ICODE""<>''and    IFNULL(T1.""U_INRNO"",'') = '' and IFNULL(T1.""U_PSTATUS"",'')='Open'"
            Dim recqry As String = "SELECT T1.""U_INRNO"",T1.""LineId"", T0.""DocNum"", T1.""U_RICOST"",T1.""U_ICODE"",T1.""U_ITDSP"",T1.""U_JRRMK"" ,T1.""U_CSTDI"",T0.""U_GLACT"", T1.""U_REF1"",T1.""U_REF2"", T1.""U_WRHOS"" FROM ""@TNX_REVERSION"" T0
INNER JOIN ""@TNX_REVERSION_C0"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" WHERE T1.""U_ICODE""<>''and    IFNULL(T1.""U_RDREF"",'') = '' "
            objrec.DoQuery(recqry)
            If objrec.RecordCount > 0 Then
                For j As Integer = 1 To objrec.RecordCount
                    ' Dim oMaterialRevaluation As SAPbobsCOM.MaterialRevaluation
                    Dim oMaterialRevaluationSNBLines As SAPbobsCOM.SNBLines
                    oMaterialRevaluation = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
                    oMaterialRevaluation.DocDate = System.DateTime.Now
                    oMaterialRevaluation.RevalType = "P"
                    Dim Item1 As String = objrec.Fields.Item("U_ICODE").Value
                    Dim DocNum As String = objrec.Fields.Item("DocNum").Value

                    Dim InvtyNo As String = objrec.Fields.Item("U_INRNO").Value

                    Dim Comments1 As String = objrec.Fields.Item("U_REF1").Value
                    Dim Newcost1 As Double = objrec.Fields.Item("U_RICOST").Value
                    Dim Warehouse1 As String = objrec.Fields.Item("U_WRHOS").Value
                    Dim GLAccnt1 As String = objrec.Fields.Item("U_GLACT").Value
                    Dim Ref1 As String = objrec.Fields.Item("U_REF1").Value
                    Dim Ref2 As String = objrec.Fields.Item("U_REF2").Value
                    Dim Jonrlrmrk1 As String = objrec.Fields.Item("U_JRRMK").Value
                    Dim Cline As String = objrec.Fields.Item("LineId").Value
                    oMaterialRevaluation.Lines.ItemCode = Item1
                    oMaterialRevaluation.Lines.WarehouseCode = Warehouse1
                    oMaterialRevaluation.Comments = Comments1
                    oMaterialRevaluation.JournalMemo = Jonrlrmrk1
                    oMaterialRevaluation.Lines.RevaluationDecrementAccount = GLAccnt1
                    oMaterialRevaluation.Lines.RevaluationIncrementAccount = GLAccnt1
                    'oMaterialRevaluation.Lines.ActualPrice = Newcost
                    oMaterialRevaluation.Lines.Price = Newcost1
                    oMaterialRevaluation.Reference2 = DocNum
                    'oMaterialRevaluation.Reference1 = Ref1
                    oMaterialRevaluation.Lines.Add()
                    ''Dim result As Integer = oMaterialRevaluation.Add()
                    ''If result <> 0 Then
                    ''    errorMessage = objMain.objCompany.GetLastErrorDescription()
                    ''    Return False
                    ''Else
                    ''End If
                    ''Dim DocEntry As Integer
                    ''DocEntry = objMain.objCompany.GetNewObjectKey()
                    ''Dim updateQuery As String = "UPDATE ""@TNX_REVERSION_C0"" Set ""U_INRNO"" = '" & DocEntry.ToString() & "' , ""U_PSTATUS""='Posted'  WHERE ifnull(""U_INRNO"",'') = '' AND ""LineId"" = '" & Cline.ToString() & "'   and ""DocEntry"" = (SELECT ""DocEntry"" FROM ""@TNX_REVERSION""  WHERE ""DocNum"" = '" & DocNum & "')"
                    ''Dim objRecord As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    ''objRecord.DoQuery(updateQuery)
                    ''objMain.objApplication.StatusBar.SetText("Inventory Revaluation is successfully added for item " & Item1 & " with Revaluation No: " & DocEntry, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    ''objrec.MoveNext()

                    Dim result As Integer = oMaterialRevaluation.Add()
                    If result <> 0 Then
                        Dim errCode As Integer
                        Dim errMsg As String
                        objMain.objCompany.GetLastError(errCode, errMsg)
                        objMain.objApplication.MessageBox("Error " & errCode & ": " & errMsg)
                    Else
                        Dim DocEntry As Integer
                        DocEntry = objMain.objCompany.GetNewObjectKey()

                        'Dim updateQuery As String = "UPDATE ""@TNX_REVERSION_C0"" Set ""U_INRNO"" = '" & DocEntry.ToString() & "' , ""U_PSTATUS""='Posted'  WHERE ifnull(""U_INRNO"",'') = '' AND ""LineId"" = '" & Cline.ToString() & "'   and ""DocEntry"" = (SELECT ""DocEntry"" FROM ""@TNX_REVERSION""  WHERE ""DocNum"" = '" & DocNum & "')"
                        Dim updateQuery As String = "UPDATE ""@TNX_REVERSION_C0"" Set ""U_RDREF"" = '" & DocEntry.ToString() & "' , ""U_PSTATUS""='Posted'  WHERE ifnull(""U_RDREF"",'') = '' AND ""LineId"" = '" & Cline.ToString() & "'   and ""DocEntry"" = (SELECT ""DocEntry"" FROM ""@TNX_REVERSION""  WHERE ""DocNum"" = '" & DocNum & "')"
                        Dim objRecord As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objRecord.DoQuery(updateQuery)
                        ' ✅ Success Message
                        objMain.objApplication.StatusBar.SetText("Inventory Revaluation is successfully added for item " & Item1 & " with Revaluation No: " & DocEntry, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    End If

                    objrec.MoveNext()

                Next
            End If

        Catch ex As Exception
            errorMessage = ex.Message
            Return False
        End Try
    End Function




    Private Sub LoadDetails(ByVal pVal As SAPbouiCOM.ItemEvent, ByVal FormUID As String)
        Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUID)

        If objForm Is Nothing Then
            objMain.objApplication.MessageBox("Form not found.")
            Return
        End If

        Dim promoRevNo As String = CType(objForm.Items.Item("PCRNO").Specific, SAPbouiCOM.EditText).Value

        If String.IsNullOrEmpty(promoRevNo) Then
            objMain.objApplication.MessageBox("Please select a Promo Reversion Number.")
            Return
        End If

        ' ========== HEADER + MATRIX DATA LOADING ==========
        Dim oRecordSet As SAPbobsCOM.Recordset
        oRecordSet = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        ' Optimized Query with INNER JOIN
        Dim query As String = "SELECT D.""U_CSTDTL"" as ""Costdifftotal"",H.""U_BRANCH"" as ""Branch"", H.""U_TOTAL"",D.""U_COSTD"",H.""U_BYNAME"", H.""U_SPCODE"", H.""U_SPNAME"",H.""U_GLACNT"", " &
                          "D.""U_ITMCODE"", D.""U_ITMDESC"", D.""U_ISTQTY"", D.""U_OLDCST"", D.""U_NEWCST"", " &
                          "D.""U_JRMRK"", D.""U_BRANCH"", D.""U_WRHSE"", " &
                          "D.""U_REF1"", D.""U_REF2"", D.""U_IREVNO"", D.""U_PSTS"", D.""U_REVDRF"" " &
                          "FROM ""@TNX_RDCTN"" H " &
                          "INNER JOIN ""@TNX_RDCTN_C0"" D ON H.""DocEntry"" = D.""DocEntry"" " &
                          "WHERE H.""DocNum"" = '" & promoRevNo & "' And D.""U_ITMCODE"" <>''"

        oRecordSet.DoQuery(query)

        Dim branch As String = oRecordSet.Fields.Item("Branch").Value
        If oRecordSet.RecordCount = 0 Then
            objMain.objApplication.MessageBox("No records found for the selected Promo Reversion Number.")
            Return
        End If
        oRecordSet.MoveFirst()

        ' ========== POPULATE HEADER FIELDS ==========
        CType(objForm.Items.Item("BNAME").Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("U_BYNAME").Value
        CType(objForm.Items.Item("SCODE").Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("U_SPCODE").Value
        CType(objForm.Items.Item("SNAME").Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("U_SPNAME").Value
        CType(objForm.Items.Item("GLACT").Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item("U_GLACNT").Value

        ' ========== MATRIX1 LOADING ==========
        Dim oMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("RVMtx").Specific
        oMatrix.Clear()
        oMatrix.FlushToDataSource()
        objForm.Freeze(True)



        Dim Itemcode As String = oRecordSet.Fields.Item("U_ITMCODE").Value.ToString()
        Dim warehouse As String = oRecordSet.Fields.Item("U_WRHSE").Value.ToString()

        'In stak cost quantity
        'Dim oRecord As SAPbobsCOM.Recordset
        'oRecord = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'Dim query11 As String = "SELECT T0.""OnHand"", T0.""AvgPrice"" FROM OITW T0 WHERE T0.""ItemCode"" = '" & Itemcode & "' AND T0.""WhsCode"" = '" & warehouse & "'"
        'oRecord.DoQuery(query11)
        'Dim Promoitmcst As String = oRecord.Fields.Item("AvgPrice").Value.ToString()

        ' Set the matrix values




        Dim lineId As Integer = 1
        While Not oRecordSet.EoF
            oMatrix.AddRow()
            Dim rowIndex As Integer = oMatrix.VisualRowCount

            oMatrix.Columns.Item("#").Cells.Item(rowIndex).Specific.Value = lineId.ToString()
            oMatrix.Columns.Item("ICODE").Cells.Item(rowIndex).Specific.Value = oRecordSet.Fields.Item("U_ITMCODE").Value
            'objMatrix.Columns.Item("PICOST").Cells.Item(rowIndex).Specific.Value = Promoitmcst

            oMatrix.Columns.Item("ITDSP").Cells.Item(rowIndex).Specific.Value = oRecordSet.Fields.Item("U_ITMDESC").Value
            'oMatrix.Columns.Item("CSQUN").Cells.Item(rowIndex).Specific.Value = oRecordSet.Fields.Item("U_ISTQTY").Value
            oMatrix.Columns.Item("PICOST").Cells.Item(rowIndex).Specific.Value = oRecordSet.Fields.Item("U_OLDCST").Value
            oMatrix.Columns.Item("RICOST").Cells.Item(rowIndex).Specific.Value = oRecordSet.Fields.Item("U_NEWCST").Value
            'oMatrix.Columns.Item("GLACT").Cells.Item(rowIndex).Specific.Value = oRecordSet.Fields.Item("U_GLACNT").Value
            oMatrix.Columns.Item("JRRMK").Cells.Item(rowIndex).Specific.Value = oRecordSet.Fields.Item("U_JRMRK").Value
            oMatrix.Columns.Item("BRANCH").Cells.Item(rowIndex).Specific.Value = oRecordSet.Fields.Item("U_BRANCH").Value
            oMatrix.Columns.Item("WRHOS").Cells.Item(rowIndex).Specific.Value = oRecordSet.Fields.Item("U_WRHSE").Value
            oMatrix.Columns.Item("REF1").Cells.Item(rowIndex).Specific.Value = oRecordSet.Fields.Item("U_REF1").Value
            oMatrix.Columns.Item("REF2").Cells.Item(rowIndex).Specific.Value = oRecordSet.Fields.Item("U_REF2").Value
            oMatrix.Columns.Item("INRNO").Cells.Item(rowIndex).Specific.Value = oRecordSet.Fields.Item("U_IREVNO").Value
            oMatrix.Columns.Item("CSTDI").Cells.Item(rowIndex).Specific.Value = oRecordSet.Fields.Item("U_COSTD").Value
            oMatrix.Columns.Item("CSTDFT").Cells.Item(rowIndex).Specific.Value = oRecordSet.Fields.Item("Costdifftotal").Value

            ' oMatrix.Columns.Item("BRANCH").Cells.Item(rowIndex).Specific.Value = oRecordSet.Fields.Item("U_BRANCH").Value
            oMatrix.Columns.Item("BRANCH").Cells.Item(rowIndex).Specific.Value = branch

            oDS.SetValue("U_TOTAL", oDS.Offset, oRecordSet.Fields.Item("U_TOTAL").Value)


            Dim comboPSTS As SAPbouiCOM.ComboBox = CType(oMatrix.Columns.Item("PSTATUS").Cells.Item(rowIndex).Specific, SAPbouiCOM.ComboBox)
            If comboPSTS.ValidValues.Count > 0 Then
                comboPSTS.Select(oRecordSet.Fields.Item("U_PSTS").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
            End If
            oMatrix.Columns.Item("RDREF").Cells.Item(rowIndex).Specific.Value = oRecordSet.Fields.Item("U_REVDRF").Value

            lineId += 1
            oRecordSet.MoveNext()
        End While
        '========== MATRIX2 LOADING ==========
        oRecordSet.MoveFirst()
        Dim oMatrix1 As SAPbouiCOM.Matrix = objForm.Items.Item("RVMtx1").Specific
        oMatrix1.Clear()
        Dim lineId1 As Integer = 1

        While Not oRecordSet.EoF
            Dim rowIndex1 As Integer = oMatrix1.VisualRowCount + 1
            oMatrix1.AddRow()

            oMatrix1.Columns.Item("#").Cells.Item(rowIndex1).Specific.Value = lineId1.ToString()
            oMatrix1.Columns.Item("OCOST").Cells.Item(rowIndex1).Specific.Value = oRecordSet.Fields.Item("U_OLDCST").Value
            oMatrix1.Columns.Item("PRCOST").Cells.Item(rowIndex1).Specific.Value = oRecordSet.Fields.Item("U_NEWCST").Value
            ' oMatrix1.Columns.Item("BRANCH").Cells.Item(rowIndex1).Specific.Value = oRecordSet.Fields.Item("U_BRANCH").Value
            lineId1 += 1
            oRecordSet.MoveNext()
        End While
        ' === Set Branch value from matrix to header field ===
        'If oMatrix.RowCount > 0 Then
        '    Dim firstBranch As String = CType(oMatrix.Columns.Item("BRANCH").Cells.Item(1).Specific, SAPbouiCOM.EditText).Value
        '    CType(objForm.Items.Item("CNPBR").Specific, SAPbouiCOM.EditText).Value = firstBranch
        'End If

        objForm.Freeze(False)
        oMatrix.FlushToDataSource()
        objMain.objApplication.StatusBar.SetText("Data loaded successfully.!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

    End Sub

    Private Sub CalculateTotalSalesAmount(ByVal oMatrix As SAPbouiCOM.Matrix)
        Try
            For i As Integer = 1 To oMatrix.RowCount
                ' Get values from the matrix cells
                Dim soldQuantity As Double = CDbl(CType(oMatrix.Columns.Item("SOLDQTY").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value)
                Dim oldCost As Double = CDbl(CType(oMatrix.Columns.Item("OLDCOST").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value)
                Dim promoCost As Double = CDbl(CType(oMatrix.Columns.Item("PROMOCOST").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value)

                Dim totalSalesAmount As Double = soldQuantity * (oldCost - promoCost)

                ' Set the calculated value back to the matrix
                oMatrix.Columns.Item("TOTSALE").Cells.Item(i).Specific.Value = totalSalesAmount.ToString("F2")
            Next

            oMatrix.AutoResizeColumns()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error in calculating Total Sales Amount: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    'Public Function AddAPCreditNote(ByRef errorMessage As String) As Boolean
    '    Try
    '        Dim objApRs As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

    '        ' Query to fetch supplier and reversion data
    '        Dim selectedDocEntry As String = objForm.Items.Item("DocNum").Specific.Value  ' "8" = UID of Promo Cost Reduction No.

    '        Dim apQuery As String = "SELECT T0.""U_SCODE"", T1.""U_OCOST"", T1.""U_PRCOST"", T1.""U_TSAMT"", T1.""U_SQNTY"", T0.""U_GLACT"" " &
    '                     "FROM ""@TNX_REVERSION"" T0 " &
    '                     "INNER JOIN ""@TNX_REVERSION_C1"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" where ""DocNum"" = '" & selectedDocEntry & "'"


    '        objApRs.DoQuery(apQuery)

    '        If objApRs.RecordCount > 0 Then
    '            For j As Integer = 1 To objApRs.RecordCount

    '                ' Create AP Credit Note
    '                Dim oCreditNote As SAPbobsCOM.Documents = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes), SAPbobsCOM.Documents)

    '                ' Header
    '                oCreditNote.CardCode = objApRs.Fields.Item("U_SCODE").Value
    '                oCreditNote.DocDate = DateTime.Now
    '                oCreditNote.TaxDate = DateTime.Now
    '                oCreditNote.DocDueDate = DateTime.Now
    '                oCreditNote.Comments = "AP Credit Note posted via DI API from Promo Reversion"

    '                ' Decide whether to post as service or item based on GL Account
    '                Dim glAccount As String = objApRs.Fields.Item("U_GLACT").Value.ToString().Trim()
    '                Dim creditAmt As Double = Convert.ToDouble(objApRs.Fields.Item("U_TSAMT").Value)

    '                If Not String.IsNullOrEmpty(glAccount) Then
    '                    ' Service-based AP Credit Note
    '                    oCreditNote.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
    '                    oCreditNote.Lines.AccountCode = glAccount
    '                    oCreditNote.Lines.LineTotal = creditAmt
    '                Else
    '                    ' Item-based AP Credit Note (if GL not found)
    '                    oCreditNote.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
    '                    oCreditNote.Lines.ItemCode = "A00001" ' Replace with real item code if available
    '                    oCreditNote.Lines.Quantity = 1
    '                    oCreditNote.Lines.UnitPrice = creditAmt
    '                    oCreditNote.Lines.WarehouseCode = "01" ' Update if dynamic
    '                    oCreditNote.Lines.TaxCode = "O1"       ' Update if required
    '                End If

    '                oCreditNote.Lines.Add()

    '                ' Post the document
    '                Dim result As Integer = oCreditNote.Add()

    '                If result <> 0 Then
    '                    errorMessage = objMain.objCompany.GetLastErrorDescription()
    '                    Return False
    '                End If

    '                objApRs.MoveNext()
    '            Next
    '        End If

    '        Return True

    '    Catch ex As Exception
    '        errorMessage = "Exception: " & ex.Message
    '        Return False
    '    End Try
    'End Function
    'Public Function AddAPCreditNote(ByRef errorMessage As String) As Boolean
    '    Dim objAp As SAPbobsCOM.Recordset
    '    objAp = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '    Dim selectedDocEntry As String = objForm.Items.Item("DocNum").Specific.Value

    '    Dim apQuery As String = "SELECT T0.""U_SCODE"", T1.""U_OCOST"", T1.""U_PRCOST"", T1.""U_TSAMT"", T1.""U_SQNTY"", T0.""U_GLACT"" " &
    '                            "FROM ""@TNX_REVERSION"" T0 " &
    '                            "INNER JOIN ""@TNX_REVERSION_C1"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
    '                            "WHERE T0.""DocNum"" = '" & selectedDocEntry & "'"

    '    objAp.DoQuery(apQuery)

    '    If objAp.RecordCount > 0 Then
    '        For j As Integer = 1 To objAp.RecordCount
    '            Dim oAPCreditNote As SAPbobsCOM.Documents
    '            oAPCreditNote = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)

    '            ' Header
    '            oAPCreditNote.CardCode = objAp.Fields.Item("U_SCODE").Value   ' Must be a supplier
    '            oAPCreditNote.DocDate = System.DateTime.Now
    '            oAPCreditNote.DocDueDate = System.DateTime.Now
    '            oAPCreditNote.TaxDate = System.DateTime.Now
    '            oAPCreditNote.Comments = "AP Credit Note created via DI API"
    '            ' oAPCreditNote.DocCurrency = "USD"

    '            ' If you want to post by G/L Account (Service Type)
    '            oAPCreditNote.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service

    '            ' Line
    '            oAPCreditNote.Lines.AccountCode = objAp.Fields.Item("U_GLACT").Value  ' G/L Account
    '            oAPCreditNote.Lines.LineTotal = objAp.Fields.Item("U_TSAMT").Value    ' Amount
    '            oAPCreditNote.Lines.TaxCode = "O1"
    '            oAPCreditNote.Lines.Add()

    '            ' Submit
    '            Dim result As Integer = oAPCreditNote.Add()
    '            If result <> 0 Then
    '                errorMessage = objMain.objCompany.GetLastErrorDescription()
    '                Return False
    '            End If

    '            objAp.MoveNext()
    '        Next
    '    End If

    '    Return True
    'End Function

    '                    '    For i As Integer = 1 To objMatrix.VisualRowCount
    'olsd
    'Public Function AddAPCreditNote(ByRef errorMessage As String, ByVal branchName As String) As Boolean
    '    Try
    '        Dim objApRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '        ' Get branch ID by branch name (active branch)
    '        Dim rsBranch As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        Dim safeBranchName As String = branchName.Replace("'", "''")
    '        rsBranch.DoQuery("SELECT ""BPLId"" FROM ""OBPL"" WHERE ""BPLName"" = '" & safeBranchName & "' AND ""Disabled"" = 'N'")
    '        Dim branchID As Integer = 0
    '        If rsBranch.RecordCount > 0 Then
    '            branchID = Convert.ToInt32(rsBranch.Fields.Item("BPLId").Value)
    '        Else
    '            errorMessage = "No active branch found with the name '" & branchName & "'. Please check your branch setup."
    '            Return False
    '        End If
    '        Dim selectedDocEntry As String = objForm.Items.Item("DocNum").Specific.Value
    '        Dim Query As String = "SELECT T0.""U_SCODE"", T1.""U_OCOST"", T1.""U_PRCOST"", T1.""U_TSAMT"", T1.""U_SQNTY"", T0.""U_GLACT"" " &
    '                          "FROM ""@TNX_REVERSION"" T0 " &
    '                          "INNER JOIN ""@TNX_REVERSION_C1"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
    '                          "WHERE T0.""DocNum"" = '" & selectedDocEntry & "'"
    '        objApRs.DoQuery(Query)
    '        If objApRs.RecordCount > 0 Then
    '            objApRs.MoveFirst()
    '            For j As Integer = 1 To objApRs.RecordCount

    '                ' While Not objApRs.EoF
    '                Dim oCreditNote As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
    '                ' Set document type to service to allow G/L accounts
    '                oCreditNote.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
    '                ' Header Data
    '                oCreditNote.CardCode = objApRs.Fields.Item("U_SCODE").Value.ToString()
    '                oCreditNote.DocDate = DateTime.Now
    '                oCreditNote.TaxDate = DateTime.Now
    '                oCreditNote.DocDueDate = DateTime.Now
    '                oCreditNote.Comments = "AP Credit Note created via DI API"
    '                oCreditNote.BPL_IDAssignedToInvoice = branchID

    '                ' Line (G/L based)
    '                oCreditNote.Lines.AccountCode = objApRs.Fields.Item("U_GLACT").Value.ToString()
    '                oCreditNote.Lines.LineTotal = objApRs.Fields.Item("U_PRCOST").Value

    '                'Dim taxcode As String = oCreditNote.Lines.TaxCode
    '                'Dim tax As String = ""
    '                'oCreditNote.Lines.TaxCode = "X1"
    '                oCreditNote.Lines.Add()

    '                Dim result As Integer = oCreditNote.Add()
    '                If result <> 0 Then
    '                    errorMessage = objMain.objCompany.GetLastErrorDescription()
    '                    Return False
    '                Else
    '                    ' Get the newly created Credit Note DocEntry
    '                    Dim newDocEntry As String = objMain.objCompany.GetNewObjectKey()

    '                    ' Query to get the DocNum from DocEntry
    '                    Dim rsDocNum As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                    rsDocNum.DoQuery("SELECT ""DocNum"" FROM ""ORPC"" WHERE ""DocEntry"" = " & newDocEntry)

    '                    If rsDocNum.RecordCount > 0 Then
    '                        Dim creditNoteNumber As String = rsDocNum.Fields.Item("DocNum").Value.ToString()
    '                        objForm.Items.Item("CNTNO").Specific.Value = creditNoteNumber
    '                    End If
    '                End If
    '                objApRs.MoveNext()
    '            Next
    '            ' End While
    '        Else
    '            errorMessage = "No data found for selected Promo Reversion."
    '            Return False
    '        End If

    '        Return True

    '    Catch ex As Exception
    '        errorMessage = "Exception: " & ex.Message
    '        Return False
    '    End Try
    'End Function

    Public Function AddAPCreditNote(ByRef errorMessage As String, ByVal branchName As String) As Boolean
        Try
            Dim objApRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            ' Get branch ID by branch name (active branch)
            Dim rsBranch As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim safeBranchName As String = branchName.Replace("'", "''")
            rsBranch.DoQuery("SELECT ""BPLId"" FROM ""OBPL"" WHERE ""BPLId"" = '" & safeBranchName & "' AND ""Disabled"" = 'N'")
            Dim branchID As Integer = 0
            If rsBranch.RecordCount > 0 Then
                branchID = Convert.ToInt32(rsBranch.Fields.Item("BPLId").Value)
            Else
                errorMessage = "No active branch found with the name '" & branchName & "'. Please check your branch setup."
                Return False
            End If

            Dim selectedDocEntry As String = objForm.Items.Item("DocNum").Specific.Value
            Dim Query As String = "SELECT T0.""U_SCODE"", T1.""U_OCOST"", T1.""U_PRCOST"", T1.""U_TSAMT"", T1.""U_SQNTY"", T0.""U_GLACT"" " &
                              "FROM ""@TNX_REVERSION"" T0 " &
                              "INNER JOIN ""@TNX_REVERSION_C1"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
                              "WHERE T0.""DocNum"" = '" & selectedDocEntry & "'"
            objApRs.DoQuery(Query)

            If objApRs.RecordCount > 0 Then
                objApRs.MoveFirst()
                For j As Integer = 1 To objApRs.RecordCount

                    Dim oCreditNote As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                    oCreditNote.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service

                    ' Header
                    oCreditNote.CardCode = objApRs.Fields.Item("U_SCODE").Value.ToString()
                    oCreditNote.DocDate = DateTime.Now
                    oCreditNote.TaxDate = DateTime.Now
                    oCreditNote.DocDueDate = DateTime.Now
                    oCreditNote.Comments = "AP Credit Note created via DI API"
                    oCreditNote.BPL_IDAssignedToInvoice = branchID

                    ' Line
                    oCreditNote.Lines.AccountCode = objApRs.Fields.Item("U_GLACT").Value.ToString()
                    oCreditNote.Lines.LineTotal = objApRs.Fields.Item("U_PRCOST").Value
                    oCreditNote.Lines.Add()

                    Dim result As Integer = oCreditNote.Add()
                    If result <> 0 Then
                        errorMessage = objMain.objCompany.GetLastErrorDescription()
                        Return False
                    Else
                        ' Get new Credit Note DocEntry
                        Dim newDocEntry As String = objMain.objCompany.GetNewObjectKey()

                        ' Get the Credit Note number (DocNum)
                        Dim rsDocNum As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        rsDocNum.DoQuery("SELECT ""DocNum"" FROM ""ORPC"" WHERE ""DocEntry"" = " & newDocEntry)

                        If rsDocNum.RecordCount > 0 Then
                            Dim creditNoteNumber As String = rsDocNum.Fields.Item("DocNum").Value.ToString()

                            ' Display on the form
                            objForm.Items.Item("CNTNO").Specific.Value = creditNoteNumber

                            ' ✅ FIX: Store Credit Note number into UDO field
                            objForm.DataSources.DBDataSources.Item("@TNX_REVERSION").SetValue("U_CNTNO", 0, creditNoteNumber)

                            ' Optional: Save the update if needed
                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular) ' Save button
                            End If
                        End If
                    End If
                    objApRs.MoveNext()
                Next
            Else
                errorMessage = "No data found for selected Promo Reversion."
                Return False
            End If

            Return True

        Catch ex As Exception
            errorMessage = "Exception: " & ex.Message
            Return False
        End Try
    End Function


End Class




