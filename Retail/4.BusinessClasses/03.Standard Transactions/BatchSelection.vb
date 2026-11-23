Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports SAPbobsCOM
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Windows.Forms

Public Class clsBatchSelection

#Region "  Declaration   "
    Public objform, objListForm As SAPbouiCOM.Form
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim objItemMtx, objBatchMtx As SAPbouiCOM.Matrix
#End Region

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.BeforeAction = False Then
                        Me.AddItems(objform.UniqueID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "btnBatch" And pVal.BeforeAction = False Then
                        objMain.Validate = 1
                        objform.Items.Item("btnBatch").Enabled = False
                        Me.AutoBatchSelection(objform.UniqueID)
                        'objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False Then
                        Me.RequestedQuantity(objform.UniqueID)

                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub AddItems(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)

            objMain.objUtilities.AddButtonH(objform.UniqueID, "btnBatch", objform.Items.Item("16").Top,
            objform.Items.Item("16").Left - objform.Items.Item("16").Width - 5, 115, "16", "Auto Batch", 0, 0, objform.Items.Item("2").Height)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub AutoBatchSelection(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)
            objItemMtx = objform.Items.Item("3").Specific
            objBatchMtx = objform.Items.Item("4").Specific
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
            objform.Freeze(True)
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

                ItemExpQry = " Select Top 1 TO_CHAR(""EXPDATE"",'yyyyMMdd') As ""EXPDATE"",""DOCNUM"",""ATTR1"" From ""TNX_PICK"" Where ""ITEMCODE""='" & ItemCode & "' AND ""UPDATED""='N'   AND ""USER""='" & objMain.objCompany.UserName & "'"
                rsItemExp = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                rsItemExp.DoQuery(ItemExpQry)
                ShelfExp = rsItemExp.Fields.Item("EXPDATE").Value
                SODocNum = rsItemExp.Fields.Item("DOCNUM").Value
                BatchNum = rsItemExp.Fields.Item("ATTR1").Value
                Dim ItemExpUpdateQry As String = " Update TOP 1 ""TNX_PICK"" SET ""UPDATED""='Y' Where ""ITEMCODE""='" & ItemCode & "' AND ""DOCNUM""='" & SODocNum & "' AND ""USER""='" & objMain.objCompany.UserName & "'"
                rsItemExpUpdate = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                rsItemExpUpdate.DoQuery(ItemExpUpdateQry)
                'Dim PickQry As String = "Select Distinct C.""ItemCode"",B.""WHSECODE"",C.""OnHand"" from ""TNX_PICK"" B " &
                '                        "inner join OITW C on C.""WhsCode"" = B.""WHSECODE"" And C.""ItemCode"" = B.""ITEMCODE"""
                'Dim PickQry As String = "Select * from ""TNX_PICK"""
                'Dim rsPickQry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                'rsPickQry.DoQuery(PickQry)
                Dim Item As String = objItemMtx.Columns.Item("1").Cells.Item(i).Specific.Value '= rsPickQry.Fields.Item("ITEMCODE").Value And
                Dim warehouse As String = objItemMtx.Columns.Item("3").Cells.Item(i).Specific.Value '= rsPickQry.Fields.Item("WHSECODE").Value Then
                Dim releaseqty As Double = objItemMtx.Columns.Item("234000022").Cells.Item(i).Specific.Value
                Dim UpdatePickQty As String = "Update ""TNX_PICK"" set ""RELEASEQTY"" = '" & releaseqty & "'" &
                                              "Where ""ITEMCODE"" = '" & Item & "' And ""WHSECODE"" = '" & warehouse & "'  AND ""USER""='" & objMain.objCompany.UserName & "'"
                Dim rsUpdatePickQty As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                rsUpdatePickQty.DoQuery(UpdatePickQty)
                If TotalReqQty > 0 Then
                    Dim TotalBatchLines As Integer = objBatchMtx.VisualRowCount
                    If TotalBatchLines > 0 Then
                        objBatchMtx.Columns.Item("234000062").Cells.Item(1).Specific.Value = 0
                        Dim oColumn As SAPbouiCOM.Column = objBatchMtx.Columns.Item("15")
                        oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
                    End If

                    For j As Integer = 1 To TotalBatchLines
                        Dim CurrentBatchLines As Integer = objBatchMtx.VisualRowCount
                        If j > CurrentBatchLines Then
                            Exit For
                        End If
                        BatchQty = objBatchMtx.Columns.Item("234000058").Cells.Item(j).Specific.Value
                        AllocBatchQty = objBatchMtx.Columns.Item("234000061").Cells.Item(j).Specific.Value
                        ExpDate = objBatchMtx.Columns.Item("15").Cells.Item(j).Specific.Value
                        BatchAttr = objBatchMtx.Columns.Item("17").Cells.Item(j).Specific.Value
                        BatchQty = BatchQty - AllocBatchQty
                        If ExpDate >= ShelfExp And (BatchNum = BatchAttr) Then
                            If TotalReqQty <= BatchQty Then
                                objBatchMtx.Columns.Item("234000062").Cells.Item(j).Specific.Value = TotalReqQty
                                objBatchMtx.Columns.Item("234000059").Cells.Item(j).Specific.Value = TotalReqQty
                                objform.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                ItemNotFound = "Yes"
                                Exit For
                            ElseIf BatchQty > 0 Then
                                If TotalReqQty = 0 Then
                                    Exit For
                                Else
                                    Qty = BatchQty
                                    'TotalSelectedQty = TotalSelectedQty + Qty
                                    'TotalSelectedQty = TotalSelectedQty + (TotalReqQty - BatchQty)
                                    TotalReqQty = TotalReqQty - Qty
                                End If
                                objBatchMtx.Columns.Item("234000062").Cells.Item(j).Specific.Value = Qty
                                objBatchMtx.Columns.Item("234000059").Cells.Item(j).Specific.Value = Qty
                                objform.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                ItemNotFound = "Yes"
                                If TotalReqQty = 0 Then
                                    Exit For
                                Else
                                    j = j - 1
                                End If
                            End If
                        End If
                    Next
                    Dim warehouse1 As String = objItemMtx.Columns.Item("3").Cells.Item(i).Specific.Value
                    Dim requireqty As Double = objItemMtx.Columns.Item("234000021").Cells.Item(i).Specific.Value
                    Dim Quantity As Double = objItemMtx.Columns.Item("234000020").Cells.Item(i).Specific.Value
                    Dim releaseqty1 As Double = objItemMtx.Columns.Item("234000022").Cells.Item(i).Specific.Value
                    Dim RequestedQuant As String = "Update ""TNX_PICK"" Set ""REQUIREDQTY"" = '" & requireqty & "',""RELEASEQTY"" = '" & releaseqty1 & "' " &
                                               "Where ""ITEMCODE"" = '" & Item & "' And ""WHSECODE"" = '" & warehouse1 & "' And ""QUANTITY"" = '" & Quantity & "'  AND ""USER""='" & objMain.objCompany.UserName & "' "
                    Dim RsRequestedQuant As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                    RsRequestedQuant.DoQuery(RequestedQuant)
                    If ItemNotFound = "No" Then
                        Dim ReItemExpUpdateQry As String = " Update ""TNX_PICK"" SET ""UPDATED""='N' Where ""ITEMCODE""='" & ItemCode & "' AND ""DOCNUM""='" & SODocNum & "' AND ""USER""='" & objMain.objCompany.UserName & "'"
                        rsReItemExpUpdate = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                        rsReItemExpUpdate.DoQuery(ReItemExpUpdateQry)
                        objMain.objApplication.StatusBar.SetText("Item Batches Not found for given shelf life, for Item code: " & ItemCode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If

                End If
            Next
            objform.Freeze(False)
        Catch ex As Exception
            objform.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    Sub RequestedQuantity(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)
            objItemMtx = objform.Items.Item("3").Specific
            objMain.objApplication.StatusBar.SetText("Loading Required Quantity, please wait.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            '1234objform.Freeze(True)
            For i As Integer = 1 To objItemMtx.VisualRowCount
                Dim Item As String = objItemMtx.Columns.Item("1").Cells.Item(i).Specific.Value
                Dim warehouse1 As String = objItemMtx.Columns.Item("3").Cells.Item(i).Specific.Value
                Dim requireqty As Double = objItemMtx.Columns.Item("234000021").Cells.Item(i).Specific.Value
                Dim Quantity As Double = objItemMtx.Columns.Item("234000020").Cells.Item(i).Specific.Value
                Dim releaseqty1 As Double = objItemMtx.Columns.Item("234000022").Cells.Item(i).Specific.Value
                Dim RequestedQuant As String = "Update ""TNX_PICK"" Set ""REQUIREDQTY"" = '" & requireqty & "',""RELEASEQTY"" = '" & releaseqty1 & "' " &
                                               "Where ""ITEMCODE"" = '" & Item & "' And ""WHSECODE"" = '" & warehouse1 & "' And ""QUANTITY"" = '" & Quantity & "'  AND ""USER""='" & objMain.objCompany.UserName & "' "
                Dim RsRequestedQuant As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                RsRequestedQuant.DoQuery(RequestedQuant)
            Next
            objform.Freeze(False)
        Catch ex As Exception
            objform.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
End Class
