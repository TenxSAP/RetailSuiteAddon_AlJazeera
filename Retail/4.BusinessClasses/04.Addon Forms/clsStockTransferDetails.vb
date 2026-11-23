Public Class clsStockTransferDetails

#Region "Declarations"
    Dim objForm, objBaseForm, objITRForm As SAPbouiCOM.Form
    Dim objMatrix, objBaseMatrix, objITRMatrix As SAPbouiCOM.Matrix
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objGrid As SAPbouiCOM.Grid
    Dim objGridColumn As SAPbouiCOM.GridColumn
    Dim objGridColumns As SAPbouiCOM.GridColumns
    Dim objComboGrid As SAPbouiCOM.ComboBoxColumn
    Dim objCheckBox As SAPbouiCOM.CheckBox
    Dim oDt As SAPbouiCOM.DataTable
    Dim sCardCode As String
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            objMain.objUtilities.LoadForm("StockTransferDetails.xml", "ME_STD_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_STD_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objBaseForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.DataSources.UserDataSources.Add("SelectAll", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objCheckBox = objForm.Items.Item("4").Specific
            objCheckBox.DataBind.SetBound(True, "", "SelectAll")
            objForm.Items.Item("4").AffectsFormMode = False

            oDt = objForm.DataSources.DataTables.Add("dt")
            oDt = objForm.DataSources.DataTables.Item("dt")

            Me.LoadGrid(objForm.UniqueID, "N")
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub LoadGrid(ByVal FormUID As String, ByVal CheckValue As String)
        Try
            objForm.Freeze(True)
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objGrid = objForm.Items.Item("3").Specific
            objGrid.DataTable = oDt
            objGrid.DataTable.Clear()

            Dim Query As String = " Select '" & CheckValue & "'  ""CheckBox"" , ""DocNum""  ""STRNo""  , T0.""DocEntry""  As ""STREntry""  , ""U_MEDOCDT""  As ""DocDate""  ,  T0.""U_MEFWHSCD""   ""FromWarehouse""  , " & _
                                  " T0.""U_METWHSCD""  ""ToWarehouse""  From ""@ME_STR""  T0 Inner Join ""@ME_STR_C0""  T1 On T1.""DocEntry""  = T0.""DocEntry""  " & _
                                  " Where IFNULL(""U_MEITMCD"" ,'') <> '' And ""U_MELNSTS""  = 'Open' And  ""U_MEFBPID"" = '" & objBaseForm.Items.Item("1320000098").Specific.Value.Trim & "' " & _
                                  " Group By T0.""DocNum""  , T0.""DocEntry""   , ""U_MEDOCDT""  ,  T0.""U_MEFWHSCD""   , T0.""U_METWHSCD""  "
            objGrid.DataTable.ExecuteQuery(Query)
            If objGrid.Rows.Count = 0 Then
                objMain.objApplication.MessageBox("No records found")
                Exit Sub
            End If
            For i As Integer = 0 To objGrid.Columns.Count - 1
                objGrid.Columns.Item(i).Editable = False
            Next
            objGrid.Columns.Item("STREntry").Visible = False
            objGrid.Columns.Item("CheckBox").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            objGrid.Columns.Item("CheckBox").Editable = True
            'Dim oEditTextCol As SAPbouiCOM.EditTextColumn
            'oEditTextCol = objGrid.Columns.Item("AbsID")
            'oEditTextCol.LinkedObjectType = 1250000025
            objGrid.AutoResizeColumns()
            objForm.Freeze(False)
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
                    If pVal.ItemUID = "4" And pVal.BeforeAction = False Then
                        objCheckBox = objForm.Items.Item("4").Specific
                        objForm.Freeze(True)
                        If objForm.Items.Item("4").Specific.checked = True Then
                            Me.LoadGrid(objForm.UniqueID, "Y")
                        ElseIf objForm.Items.Item("4").Specific.checked = False Then
                            Me.LoadGrid(objForm.UniqueID, "N")
                        End If
                        objForm.Freeze(False)
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False Then
                        Me.LoadMatrix(objForm.UniqueID)
                        objForm.Close()
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub LoadMatrix(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objITRForm = objMain.objApplication.Forms.GetForm("940", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objITRMatrix = objITRForm.Items.Item("23").Specific
            Dim oColumn As SAPbouiCOM.Column
            For i As Integer = 0 To objGrid.Rows.Count - 1
                If objGrid.DataTable.Columns.Item("CheckBox").Cells.Item(i).Value.Equals("Y") Then
                    If objITRForm.Items.Item("U_TRTYPE").Specific.Value = "" Then
                        objITRForm.Items.Item("U_TRTYPE").Specific.Value = "STR"
                        objITRForm.Items.Item("18").Specific.Value = objGrid.DataTable.Columns.Item("FromWarehouse").Cells.Item(i).Value
                        Dim GetInTransitWhs As String = "Select ""U_INTRWHS"" From ""OWHS""  Where ""U_INTRWHS"" <> '' And ""WhsCode"" = '" & objGrid.DataTable.Columns.Item("FromWarehouse").Cells.Item(i).Value & "' "
                        Dim oRsGetInTransitWhs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetInTransitWhs.DoQuery(GetInTransitWhs)
                        objITRForm.Items.Item("1470000101").Specific.Value = oRsGetInTransitWhs.Fields.Item("U_INTRWHS").Value
                    End If
                    Dim GetSTRLineDetails As String = " Select * From ""@ME_STR""  T0 Inner Join ""@ME_STR_C0""  " & _
                                                      " T1 On T1.""DocEntry""  = T0.""DocEntry""   Where IFNULL(""U_MEITMCD"" ,'') <> '' And ""U_MELNSTS""  = 'Open' And T0.""DocEntry"" = '" & objGrid.DataTable.Columns.Item("STREntry").Cells.Item(i).Value & "'"
                    Dim oRsGetSTRLineDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetSTRLineDetails.DoQuery(GetSTRLineDetails)
                    If oRsGetSTRLineDetails.RecordCount > 0 Then
                        oRsGetSTRLineDetails.MoveFirst()
                        For j As Integer = 1 To oRsGetSTRLineDetails.RecordCount
                            objITRMatrix.Columns.Item("1").Cells.Item(objITRMatrix.VisualRowCount).Specific.Value = oRsGetSTRLineDetails.Fields.Item("U_MEITMCD").Value
                            oColumn = objITRMatrix.Columns.Item("1001")
                            oColumn.Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Select("No", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            objITRMatrix.Columns.Item("1470001043").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = oRsGetSTRLineDetails.Fields.Item("U_MEUOM").Value
                            'objITRMatrix.Columns.Item("64").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = oRsGetSTRLineDetails.Fields.Item("U_MEBRCD").Value
                            objITRMatrix.Columns.Item("U_INVTRRLN").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = oRsGetSTRLineDetails.Fields.Item("LineId").Value
                            objITRMatrix.Columns.Item("U_INVTRREQ").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = oRsGetSTRLineDetails.Fields.Item("DocNum").Value
                            objITRMatrix.Columns.Item("U_INVTRREN").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = oRsGetSTRLineDetails.Fields.Item("DocEntry").Value
                            objITRMatrix.Columns.Item("U_METSTQTY").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = CDbl(oRsGetSTRLineDetails.Fields.Item("U_MEREQQTY").Value)
                            objITRMatrix.Columns.Item("U_MEREQQTY").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = CDbl(oRsGetSTRLineDetails.Fields.Item("U_MEOPNQTY").Value)
                            objITRMatrix.Columns.Item("U_Packing").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = oRsGetSTRLineDetails.Fields.Item("U_MEPCKNG").Value
                            objITRMatrix.Columns.Item("U_Addbarcode").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = CDbl(oRsGetSTRLineDetails.Fields.Item("U_MEADBRCD").Value)
                            Dim GetToBranch As String = "Select ""BPLName"" From ""OBPL""  Where   ""BPLId"" = '" & oRsGetSTRLineDetails.Fields.Item("U_METBPID").Value & "' "
                            Dim oRsGetToBranch As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetToBranch.DoQuery(GetToBranch)
                            objITRMatrix.Columns.Item("U_METBPID").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = oRsGetToBranch.Fields.Item("BPLName").Value
                            objITRMatrix.Columns.Item("U_METWHSCD").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = oRsGetSTRLineDetails.Fields.Item("U_METWHSCD").Value
                            objITRMatrix.Columns.Item("U_METWHSNM").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = oRsGetSTRLineDetails.Fields.Item("U_METWHSNM").Value
                            objITRMatrix.Columns.Item("10").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = CDbl(oRsGetSTRLineDetails.Fields.Item("U_MEOPNQTY").Value)
                            oRsGetSTRLineDetails.MoveNext()
                        Next
                    End If
                End If
            Next
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
