Imports System.Net.Http
Imports SAPbobsCOM

Public Class clsCategorySplit

#Region "    Declaration   "
    Dim objForm, objFormX As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix

    Dim oEditText As SAPbouiCOM.EditText
    Dim oDt, oDt1, oDataTable As SAPbouiCOM.DataTable
    Dim objGrid, oGrid As SAPbouiCOM.Grid
    Dim sSupplierCode, sSupplierName As String
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim MATRIXS As String
    Dim LostFocusFlag As Boolean = False
    Dim SKUValue, SkuValuenew, rowid, CardCode, DocNum As String
    Dim Chargevalue As Double = 0

#End Region

#Region "    Public Methods   "


    Sub CreateForm(SKUValue, row, CardCode, DocNum)

        Try

            Dim oRsCheck As SAPbobsCOM.Recordset
            Dim IfExists As String = "N"

            Dim sCheckQry As String = "SELECT ""DocEntry"" FROM ""@TNX_CATERY"" " & "WHERE ""U_SKUS"" = '" & SKUValue & "' " & "AND ""U_RowID"" = '" & row & "' " &
                                     "AND ""U_CardCode"" = '" & CardCode & "' " & "AND ""U_BNO"" = '" & DocNum & "'"


            oRsCheck = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsCheck.DoQuery(sCheckQry)

            If oRsCheck.RecordCount > 0 Then
                IfExists = "Y"
            End If

            objMain.objUtilities.LoadForm("CategorySplitScreen.xml", "Category", ResourceType.Embeded)

            objForm = objMain.objApplication.Forms.GetForm("Category", objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.EnableMenu("774", True)
            objForm.EnableMenu("1292", True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CATERY")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_CATERY_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific


            objForm.DataBrowser.BrowseBy = "DocNum"

            objForm.Freeze(True)
            If IfExists = "Y" Then

                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE

                objForm.Items.Item("SKUS").Specific.Value = SKUValue
                objForm.Items.Item("RowID").Specific.Value = row
                objForm.Items.Item("CardCode").Specific.Value = CardCode
                ''objForm.Items.Item("Item_5").Specific.Value = DocNum

                'Click Find
                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            End If
            If IfExists = "N" Then
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE

                oDBs_Head.SetValue("DocNum", 0,
                objMain.objUtilities.GetNextDocNum(objForm, "TNX_CATUDO"))
                objForm.Items.Item("SKUS").Specific.Value = SKUValue
                objForm.Items.Item("RowID").Specific.Value = row
                objForm.Items.Item("CardCode").Specific.Value = CardCode
                objForm.Items.Item("Item_5").Specific.Value = DocNum

            End If

            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("CardCode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("CardCode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("CardCode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("CardCode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("Item_5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Item_5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Item_5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("SKUS").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("SKUS").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("SKUS").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("SKUS").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            Dim objutilities As New Utilities()
            Dim objmatrix1UOM As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx").Specific



            Dim CostFluctionMatrixUOM As SAPbouiCOM.Column
            CostFluctionMatrixUOM = objmatrix1UOM.Columns.Item("MainCat")

            objutilities.MatrixComboBoxValues(
            CostFluctionMatrixUOM,
            "SELECT T0.""ItmsGrpCod"", T0.""ItmsGrpNam"" FROM OITB T0 ORDER BY T0.""ItmsGrpNam""")

            CostFluctionMatrixUOM.DisplayDesc = True
            CostFluctionMatrixUOM.ExpandType =
            SAPbouiCOM.BoExpandType.et_DescriptionOnly


            Me.LoadChargeFromOCRD(CardCode)

            SkuValuenew = SKUValue
            rowid = row

            If IfExists = "N" Then
                Me.SetNewLine(objForm.UniqueID)
            End If

            'objMatrix.Columns.Item("Charge").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            'objMatrix.Columns.Item("Listing").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix.Columns.Item("SKUCount").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objForm.Freeze(False)

        Catch ex As Exception

            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText("clsCategorySplit.CreateForm: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

#End Region


#Region "    Events   "

    Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    ' If pVal.ItemUID = "1" AndAlso pVal.BeforeAction AndAlso (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE OrElse pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then

                        Dim headerSKU As Double = 0
                        Dim totalSKU As Double = 0

                        Double.TryParse(objForm.Items.Item("SKUS").Specific.Value, headerSKU)

                        Dim oMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx").Specific

                        For i As Integer = 1 To oMatrix.VisualRowCount
                            Dim rowVal As Double = 0
                            Double.TryParse(oMatrix.Columns.Item("SKUCount").Cells.Item(i).Specific.Value, rowVal)
                            totalSKU += rowVal
                        Next

                        '-Madhu already written in transcation notification to check total SKU count with header SKU count, but added here also to avoid any bypass of that validation and to make sure data integrity
                        'If totalSKU > headerSKU Then
                        '    objMain.objApplication.StatusBar.SetText(
                        '        "Total SKU Count (" & totalSKU & ") cannot exceed Header No Of SKUs (" & headerSKU & ")",
                        '        SAPbouiCOM.BoMessageTime.bmt_Short,
                        '        SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If


                        ' Me.CalculateTotalListingUpdateInItemListInAmountField(objForm, headerSKU, objForm.Items.Item("RowID").Specific.Value)

                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUID)
                    Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx").Specific


                    If pVal.ItemUID = "Mtx" AndAlso (pVal.ColUID = "SKUCount" OrElse pVal.ColUID = "Charge") AndAlso pVal.BeforeAction = False Then

                        Dim skuCount As Double = 0
                        Dim charge As Double = 0

                        Double.TryParse(objMatrix.Columns.Item("SKUCount").Cells.Item(pVal.Row).Specific.Value, skuCount)
                        Double.TryParse(objMatrix.Columns.Item("Charge").Cells.Item(pVal.Row).Specific.Value, charge)


                        Me.Chargevalue = charge

                        Dim listing As Double = skuCount * charge
                        objMatrix.Columns.Item("Listing").Cells.Item(pVal.Row).Specific.Value = listing
                        objMatrix.Columns.Item("SKUCount").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        ' Me.CalculateTotalListing(objForm)


                    End If


                    If pVal.ItemUID = "Mtx" AndAlso pVal.ColUID = "Listing" AndAlso pVal.BeforeAction = False Then
                        If pVal.Row = objMatrix.VisualRowCount Then
                            Me.SetNewLine(objForm.UniqueID)
                        End If

                        ' Me.CalculateTotalListing(objForm)

                    End If



                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) And pVal.BeforeAction = False Then
                        Retail.MainCls.ohtLookUpForm.Remove(objForm.UniqueID)
                    End If

            End Select

        Catch ex As Exception
            Try : objForm.Freeze(False) : Catch : End Try
            objMain.objApplication.StatusBar.SetText("clsCategorySplit.ItemEvent: " & ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "Category" And pVal.BeforeAction = False Then
                Me.CreateForm(SKUValue, rowid, CardCode, DocNum)


            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                Me.SetNewLine(objForm.UniqueID)

            ElseIf pVal.MenuUID = "1293" And pVal.BeforeAction = False Then
                Me.SetNewLine(objForm.UniqueID)



            ElseIf pVal.MenuUID = "774" And pVal.BeforeAction = False Then

                objMatrix = objForm.Items.Item("Mtx").Specific

                Dim row As Integer = objMatrix.VisualRowCount

                If objMatrix.IsRowSelected(1) <> True And objMatrix.VisualRowCount < 1 Then

                    objMatrix.AddRow()

                    oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)

                    objMatrix.SetLineData(objMatrix.VisualRowCount)

                End If

                If objMatrix.IsRowSelected(row) = True Then

                    objMatrix.DeleteRow(row)

                Else

                    For i As Integer = 1 To objMatrix.VisualRowCount - 1

                        If objMatrix.IsRowSelected(i) = True Then

                            objMatrix.DeleteRow(i)

                        End If

                    Next

                End If

                For i As Integer = 1 To objMatrix.VisualRowCount

                    objMatrix.Columns.Item("LineId").Cells.Item(i).Specific.Value = i

                Next
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub


    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CATERY")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_CATERY_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific

            objMatrix.AddRow()

            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
            oDBs_Details.SetValue("U_MainCat", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_SKUCount", oDBs_Details.Offset, "")

            oDBs_Details.SetValue("U_Charge", oDBs_Details.Offset, Me.Chargevalue.ToString())

            oDBs_Details.SetValue("U_Listing", oDBs_Details.Offset, "")

            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

#End Region

    Private Sub CalculateTotalListing(ByVal objForm As SAPbouiCOM.Form)
        Try
            Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx").Specific

            Dim totalListing As Double = 0
            Dim rowListing As Double = 0

            For i As Integer = 1 To objMatrix.VisualRowCount
                Double.TryParse(objMatrix.Columns.Item("Listing").Cells.Item(i).Specific.Value, rowListing)
                totalListing += rowListing
            Next

            objForm.Items.Item("Total").Specific.Value = totalListing

            Dim totalValue As Double = totalListing

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("CalculateTotalListing Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub CalculateTotalListingUpdateInItemListInAmountField(ByVal objForm As SAPbouiCOM.Form, ByVal HeaderSKU As String, ByVal rowid As String)

        Try

            Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx").Specific
            Dim totalListing As Double = 0
            Dim rowListing As Double = 0

            For i As Integer = 1 To objMatrix.VisualRowCount
                Double.TryParse(objMatrix.Columns.Item("Listing").Cells.Item(i).Specific.Value, rowListing)
                totalListing += rowListing
            Next

            objForm.Items.Item("Total").Specific.Value = totalListing

            Dim oRs As SAPbobsCOM.Recordset =
            objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim updateQuery As String =
            "UPDATE ""@SBO_LIST_C0"" SET ""U_Amount"" = '" & totalListing & "'" &
            " WHERE ""U_SKU"" = '" & HeaderSKU & "' AND ""LineId"" = '" & rowid & "'"

            oRs.DoQuery(updateQuery)

            Dim f As SAPbouiCOM.Form

            For Each frm As SAPbouiCOM.Form In objMain.objApplication.Forms
                If frm.TypeEx = "ME_ITMLIST" Then
                    f = frm
                    Exit For
                End If
            Next

            If f Is Nothing Then
                objMain.objApplication.StatusBar.SetText("ME_ITMLIST form not found.", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            Try
                f.Freeze(True)
                Dim m As SAPbouiCOM.Matrix = f.Items.Item("MtxDT").Specific
                m.FlushToDataSource()
                Dim H As SAPbouiCOM.DBDataSource = f.DataSources.DBDataSources.Item("@SBO_LIST")
                Dim D As SAPbouiCOM.DBDataSource = f.DataSources.DBDataSources.Item("@SBO_LIST_C0")

                H.Query()
                D.Query()
                m.LoadFromDataSource()

                f.Update()
                f.Refresh()

            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText("Refresh failed: " & ex.Message,
                 SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            Finally
                f.Freeze(False)
            End Try

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("CalculateTotalListing Error: " & ex.Message,
             SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub


    ' Get Charge from OCRD.U_LCPSKU based on CardCode

    Private Sub LoadChargeFromOCRD(ByVal CardCode As String)
        Try
            Dim oRs As SAPbobsCOM.Recordset =
            objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim sQry As String =
            "SELECT ""U_LCPSKU"" FROM ""OCRD"" WHERE ""CardCode"" = '" & CardCode & "'"

            oRs.DoQuery(sQry)

            If oRs.RecordCount > 0 Then
                Double.TryParse(oRs.Fields.Item("U_LCPSKU").Value.ToString(), Me.Chargevalue)
            End If

            Dim oMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx").Specific
            Dim oDBs_Details As SAPbouiCOM.DBDataSource =
            objForm.DataSources.DBDataSources.Item("@TNX_CATERY_C0")

            For i As Integer = 1 To oMatrix.VisualRowCount
                oDBs_Details.SetValue("U_Charge", i - 1, Me.Chargevalue.ToString())
            Next

            oMatrix.LoadFromDataSource()



        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "Failed to load LCP SKU Charge: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub



End Class
