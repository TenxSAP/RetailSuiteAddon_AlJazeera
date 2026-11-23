Public Class clsPurchaseHistoryDetails
#Region "Declarations"
    Dim objForm, objPOForm As SAPbouiCOM.Form
    Dim objMatrix, objPOMatrix As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim objGrid As SAPbouiCOM.Grid
    Dim oDt As SAPbouiCOM.DataTable
    Dim sItemCode As String
#End Region

    Sub CreateForm(ByVal FormUID As String, ByVal ItemCode As String)
        Try
            objMain.objUtilities.LoadForm("PurchaseHistoryDetails.xml", "ME_PHD_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_PHD_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objPOForm = objMain.objApplication.Forms.Item(FormUID)

            If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = False Then
                Retail.MainCls.ohtLookUpForm.Add(objForm.UniqueID, FormUID)
            End If

            oDt = objForm.DataSources.DataTables.Add("dt")
            oDt = objForm.DataSources.DataTables.Item("dt")
            sItemCode = ItemCode
            Me.LoadGrid(objForm.UniqueID, sItemCode)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub LoadGrid(ByVal FormUID As String, ByVal sItemCode As String)
        Try
            objForm.Freeze(True)
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objGrid = objForm.Items.Item("3").Specific
            objGrid.DataTable = oDt
            objGrid.DataTable.Clear()
            Dim Query As String = ""
            If Hana = "Hana" Then
                Query = "  Select TOP 10 T0.""DocEntry"" As ""DocEntry Number"", T0.""DocNum""As ""Document Number"",TO_VARCHAR(TO_DATE(T0.""DocDate""),'DD/MM/YYYY') As ""Document Date"", " & _
                                      " T0.""CardName"" As ""Supplier"", T1.""Quantity"" As ""Overall Quantity"", ""UomCode"" As ""Pack"" , ""InvQty"" As ""Uom"" , " & _
                                      " T1.""U_suppprice"" As ""Price"" , (T1.""U_normalqty"" * T1.""U_suppprice"") As ""Amount""  , " & _
                                      "  ""U_DiscountAmt"" As ""Discount Amount"" ,  T1.""LineTotal"" As ""NetAmount"",""U_FOCQTY"" As ""FOC Quantity"" , ""Price"" As ""Net Cost""   " & _
                                      " From OPDN T0 Inner Join PDN1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" Where T1.""ItemCode"" = '" & sItemCode & "' " & _
                                      " And T0.""CANCELED"" != 'Y' Order By T0.""DocDate"" Desc"
            ElseIf Sql = "Sql" Then
                Query = "  Select TOP 10 T0.""DocEntry""As ""DocEntry Number"", T0.""DocNum""As ""Document Number"",Convert(Varchar(Max),T0.""DocDate"",103) As "" Document Date"", " & _
                                     " T0.""CardName"" As ""Supplier"", T1.""Quantity"" As ""Overall Quantity"", ""UomCode"" As ""Pack"" ,  ""InvQty"" As ""Uom"" , " & _
                                     " T1.""U_suppprice"" As ""Price"" , (T1.""U_normalqty"" * T1.""U_suppprice"") As ""Amount""  , " & _
                                     "  ""U_DiscountAmt"" As ""Discount Amount"" ,  T1.""LineTotal"" As ""NetAmount"",""U_FOCQTY"" As ""FOC Quantity"" , ""Price"" As ""Net Cost""   " & _
                                     " From OPDN T0 Inner Join PDN1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" Where T1.""ItemCode"" = '" & sItemCode & "' " & _
                                     " And T0.""CANCELED"" != 'Y' Order By T0.""DocDate"" Desc"
            End If
            objGrid.DataTable.ExecuteQuery(Query)
            If objGrid.Rows.Count = 0 Then
                objMain.objApplication.MessageBox("No records found")
                Exit Sub
            End If
            For i As Integer = 0 To objGrid.Columns.Count - 1
                objGrid.Columns.Item(i).Editable = False
            Next
            Dim oEditTextCol As SAPbouiCOM.EditTextColumn
            oEditTextCol = objGrid.Columns.Item("DocEntry Number")
            oEditTextCol.LinkedObjectType = 20

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

                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = True And pVal.BeforeAction = False Then
                        Retail.MainCls.ohtLookUpForm.Remove(objForm.UniqueID)
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


End Class
