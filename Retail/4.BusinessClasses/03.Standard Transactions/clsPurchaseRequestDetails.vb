Public Class clsPurchaseRequestDetails


#Region "Declarations"
    Dim objForm, objPOForm As SAPbouiCOM.Form
    Dim objMatrix, objPOMatrix As SAPbouiCOM.Matrix
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objGrid As SAPbouiCOM.Grid
    Dim objCheckBox As SAPbouiCOM.CheckBox
    Dim oDt As SAPbouiCOM.DataTable
    Dim sPRQEntry As String
#End Region

    Sub CreateForm(ByVal FormUID As String)
        Try
            objMain.objUtilities.LoadForm("PurchaseRequestDetails.xml", "ME_PRD_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_PRD_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objPOForm = objMain.objApplication.Forms.Item(FormUID)
            'sPRQEntry = PRQEntry

            objForm.DataSources.UserDataSources.Add("SelectAll", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objCheckBox = objForm.Items.Item("4").Specific
            objCheckBox.DataBind.SetBound(True, "", "SelectAll")
            objForm.Items.Item("4").AffectsFormMode = False

            objForm.Items.Item("5").Specific.Value = sPRQEntry


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

            Dim Query As String = " Select '" & CheckValue & "'  ""CheckBox"" ,  ""ItemCode"" , ""Quantity"" , ""Price"" , ""PQTReqDate"" , ""UomCode"" " & _
                                   " From OPRQ T0 Inner Join PRQ1 T1 On T0.""DocEntry"" =  T1.""DocEntry"" " & _
                                   " Where T0.""DocEntry"" = '" & sPRQEntry & "'"
            objGrid.DataTable.ExecuteQuery(Query)

            objGrid.Columns.Item("CheckBox").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox

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
            objPOForm = objMain.objApplication.Forms.GetForm("142", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objPOMatrix = objPOForm.Items.Item("38").Specific
            For i As Integer = 0 To objGrid.Rows.Count - 1
                If objGrid.DataTable.Columns.Item("CheckBox").Cells.Item(i).Value.Equals("Y") Then
                    objPOMatrix.Columns.Item("1").Cells.Item(objPOMatrix.VisualRowCount).Specific.Value = objGrid.DataTable.Columns.Item("ItemCode").Cells.Item(i).Value
                    Try
                        objPOMatrix.Columns.Item("11").Cells.Item(objPOMatrix.VisualRowCount - 1).Specific.Value = objGrid.DataTable.Columns.Item("Quantity").Cells.Item(i).Value
                        objPOMatrix.Columns.Item("14").Cells.Item(objPOMatrix.VisualRowCount - 1).Specific.Value = objGrid.DataTable.Columns.Item("Price").Cells.Item(i).Value
                    Catch ex As Exception
                    End Try
                End If
            Next
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("ME_PRD_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then                       
                    End If             
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then

                        objMain.objApplication.MessageBox("Purchase Request Document is Succesfully Added")
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
End Class
