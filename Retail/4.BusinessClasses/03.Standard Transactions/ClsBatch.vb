Imports SAPbobsCOM

Public Class ClsBatch

#Region "Declarations"
    Dim objForm, objGRPOForm, objRFPForm, objGRForm, objSRForm, objSCForm As SAPbouiCOM.Form

    Dim oitem As SAPbouiCOM.Item

    Dim objMatrix, objItemMtx, objBatchMtx, objUpperMatrix, objBatch, objGRPOMatrix, objRFPMatrix, objGRMatrix, objSRMatrix, objSCMatrix As SAPbouiCOM.Matrix


#End Region

    Private Sub AddItems(ByVal FormUID As String)
        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)
            oitem = objForm.Items.Add("AutoBat", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oitem.Top = objForm.Items.Item("2").Top
            oitem.Left = objForm.Items.Item("2").Left + objForm.Items.Item("2").Width + 5
            oitem.Width = 130
            oitem.Height = objForm.Items.Item("2").Height
            oitem.LinkTo = "2"
            Dim btn As SAPbouiCOM.Button = objForm.Items.Item("AutoBat").Specific
            btn.Caption = "&Batch Auto Load"

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error at Batch AddItems Method" & ex.Message)
        Finally
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    If pVal.BeforeAction = True Then

                        Me.AddItems(objForm.UniqueID)

                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objBatch = objForm.Items.Item("3").Specific

                    If pVal.ItemUID = "AutoBat" And pVal.BeforeAction = False Then

                        '  Me.AutoBatchLoadGRPO(objForm.UniqueID)
                        Me.AutoBatchSelection(objForm.UniqueID)

                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                                    (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                    End If





            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Private Sub AutoBatchLoadGRPO(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objUpperMatrix = objForm.Items.Item("3").Specific
            objBatch = objForm.Items.Item("4").Specific




            For i As Integer = 1 To objUpperMatrix.VisualRowCount

                objUpperMatrix.Columns.Item(0).Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                objBatch.Columns.Item(15).TitleObject.Click(SAPbouiCOM.BoCellClickType.ct_Double)

                objBatch.Columns.Item(15).TitleObject.Sortable = True
                objBatch.Columns.Item(15).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)


                ' 
                ' objBatch.Columns.Item(15).TitleObject.Sortable = True

                ' objBatch.Columns.Item(15).Sortable = True
                'objBatch.Sort(objBatch.Columns.Item(15).UniqueID, SAPbouiCOM.BoGridSortType.gst_Ascending)

                objForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            Next
            objMain.objApplication.StatusBar.SetText("Operation completed successfully..!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '  Finally
        End Try
    End Sub

    Sub AutoBatchSelection(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objItemMtx = objForm.Items.Item("3").Specific
            objBatchMtx = objForm.Items.Item("4").Specific
            Dim ItemCode As String = ""
            Dim Qty As Double = 0
            Dim TotalReqQty As Double = 0
            Dim TotalSelectedQty As Double = 0
            Dim NeededQty As Double = 0
            Dim BatchQty As Double = 0
            Dim AllocBatchQty As Double = 0
            Dim ExpDate As String
            Dim ShelfExp As String
            Dim SODocNum As String
            Dim BatchNum As String
            Dim BatchAttr As String
            Dim ItemExpQry As String
            Dim ItemNotFound As String = ""
            Dim rsItemExp As SAPbobsCOM.Recordset
            Dim rsItemExpUpdate, rsReItemExpUpdate As SAPbobsCOM.Recordset
            objMain.objApplication.StatusBar.SetText("Loading Batches, please wait.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            objForm.Freeze(True)
            For i As Integer = 1 To objItemMtx.VisualRowCount
                TotalReqQty = 0
                TotalSelectedQty = 0
                NeededQty = 0
                BatchQty = 0
                AllocBatchQty = 0
                objItemMtx.Columns.Item(0).Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                TotalReqQty = objItemMtx.Columns.Item("234000021").Cells.Item(i).Specific.Value
                ItemCode = objItemMtx.Columns.Item(1).Cells.Item(i).Specific.Value
                ItemNotFound = "No"

                Dim Item As String = objItemMtx.Columns.Item("1").Cells.Item(i).Specific.Value '= rsPickQry.Fields.Item("ITEMCODE").Value And
                Dim warehouse As String = objItemMtx.Columns.Item("3").Cells.Item(i).Specific.Value '= rsPickQry.Fields.Item("WHSECODE").Value Then
                Dim releaseqty As Double = objItemMtx.Columns.Item("234000022").Cells.Item(i).Specific.Value

                If TotalReqQty > 0 Then
                    Me.AllocateBatchesByExpiry(objForm, TotalReqQty)
                End If


            Next
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub AllocateBatchesByExpiry(ByVal objForm As SAPbouiCOM.Form, ByVal RequiredQty As Double)
        Try
            objForm.Freeze(True)

            Dim RemainingQty As Double = RequiredQty
            Dim objBatchMtx As SAPbouiCOM.Matrix = objForm.Items.Item("4").Specific

            Do While RemainingQty > 0
                Dim TotalBatchLines As Integer = objBatchMtx.VisualRowCount
                Dim batchList As New List(Of Tuple(Of Date, Integer, Double)) ' (ExpiryDate, RowIndex, AvailableQty)

                ' Collect valid and available batch rows
                For j As Integer = 1 To TotalBatchLines
                    Dim ExpDateStr As String = objBatchMtx.Columns.Item("15").Cells.Item(j).Specific.Value
                    Dim ExpiryDate As Date

                    ' Try parsing the expiry date
                    If Not Date.TryParse(ExpDateStr, ExpiryDate) Then
                        If Not Date.TryParseExact(ExpDateStr, {"yyyyMMdd", "dd.MM.yyyy", "yyyy-MM-dd"}, Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, ExpiryDate) Then
                            Continue For
                        End If
                    End If

                    ' Parse quantities
                    Dim BatchQtyStr As String = objBatchMtx.Columns.Item("234000058").Cells.Item(j).Specific.Value
                    Dim AllocatedQtyStr As String = objBatchMtx.Columns.Item("234000061").Cells.Item(j).Specific.Value

                    Dim BatchQty As Double = 0
                    Dim AllocatedQty As Double = 0

                    Double.TryParse(BatchQtyStr, BatchQty)
                    Double.TryParse(AllocatedQtyStr, AllocatedQty)

                    Dim BatchAvailableQty As Double = BatchQty - AllocatedQty
                    If BatchAvailableQty > 0 Then
                        batchList.Add(New Tuple(Of Date, Integer, Double)(ExpiryDate, j, BatchAvailableQty))
                    End If
                Next

                ' Exit loop if no batches are available
                If batchList.Count = 0 Then Exit Do

                ' Sort by expiry date (ascending)
                batchList.Sort(Function(x, y) x.Item1.CompareTo(y.Item1))

                ' Allocate from earliest batch
                Dim selectedBatch = batchList(0)
                Dim rowIndex As Integer = selectedBatch.Item2
                Dim availableQty As Double = selectedBatch.Item3
                Dim allocateQty As Double = Math.Min(availableQty, RemainingQty)

                ' Assign quantity in matrix
                objBatchMtx.Columns.Item("234000062").Cells.Item(rowIndex).Specific.Value = allocateQty
                objBatchMtx.Columns.Item("234000059").Cells.Item(rowIndex).Specific.Value = allocateQty

                ' Trigger batch allocation
                objForm.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                RemainingQty -= allocateQty
            Loop

            ' Show allocation result
            If RemainingQty > 0 Then
                objMain.objApplication.StatusBar.SetText("Partial batch allocation done: " & (RequiredQty - RemainingQty) & "/" & RequiredQty, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Else
                objMain.objApplication.StatusBar.SetText("Full batch allocation completed: " & RequiredQty, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error during batch allocation: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            objForm.Freeze(False)
        End Try
    End Sub


    Sub AllocateBatchesByExpiry_1(ByVal objForm As SAPbouiCOM.Form, ByVal RequiredQty As Double)
        Try
            objForm.Freeze(True)

            Dim objBatchMtx As SAPbouiCOM.Matrix = objForm.Items.Item("4").Specific
            Dim TotalBatchLines As Integer = objBatchMtx.VisualRowCount
            Dim RemainingQty As Double = RequiredQty

            ' Create a list to hold (ExpiryDate, RowIndex, AvailableQty)
            Dim batchList As New List(Of Tuple(Of Date, Integer, Double))

            For j As Integer = 1 To TotalBatchLines
                ' Parse expiry date
                Dim ExpDateStr As String = objBatchMtx.Columns.Item("15").Cells.Item(j).Specific.Value
                Dim ExpiryDate As Date
                If Not Date.TryParse(ExpDateStr, ExpiryDate) Then
                    ' Try custom formats here (e.g., "yyyyMMdd", "dd.MM.yyyy")
                    If Not Date.TryParseExact(ExpDateStr, {"yyyyMMdd", "dd.MM.yyyy", "yyyy-MM-dd"}, Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, ExpiryDate) Then
                        Continue For ' Still failed, skip
                    End If
                End If

                ' Calculate available quantity
                Dim BatchQty As Double = CDbl(objBatchMtx.Columns.Item("234000058").Cells.Item(j).Specific.Value)
                Dim AllocatedQty As Double = CDbl(objBatchMtx.Columns.Item("234000061").Cells.Item(j).Specific.Value)
                Dim AvailableQty As Double = BatchQty - AllocatedQty

                If AvailableQty > 0 Then
                    batchList.Add(New Tuple(Of Date, Integer, Double)(ExpiryDate, j, AvailableQty))
                End If
            Next

            ' Sort by ExpiryDate ASC
            batchList.Sort(Function(x, y) x.Item1.CompareTo(y.Item1))

            ' Start allocation
            For Each batch In batchList
                Dim rowIndex As Integer = batch.Item2
                Dim availableQty As Double = batch.Item3
                Dim allocateQty As Double = Math.Min(availableQty, RemainingQty)

                ' Allocate in both quantity columns
                objBatchMtx.Columns.Item("234000062").Cells.Item(rowIndex).Specific.Value = allocateQty
                objBatchMtx.Columns.Item("234000059").Cells.Item(rowIndex).Specific.Value = allocateQty

                ' Simulate click to register
                objForm.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                RemainingQty -= allocateQty
                If RemainingQty <= 0 Then Exit For
            Next

            If RemainingQty > 0 Then
                objMain.objApplication.StatusBar.SetText("Warning: Only part of required batch quantity found (" & RequiredQty - RemainingQty & " out of " & RequiredQty & ").", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Else
                objMain.objApplication.StatusBar.SetText("Batch allocation completed successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error during batch allocation: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            objForm.Freeze(False)
        End Try
    End Sub


End Class
