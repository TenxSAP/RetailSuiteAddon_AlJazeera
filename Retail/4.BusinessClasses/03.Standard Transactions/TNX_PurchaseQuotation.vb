Public Class TNX_PurchaseQuotation



#Region "       Declaration             "
    Public objForm, objJVForm As SAPbouiCOM.Form
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDs, oDs1, oDs2, oDs3 As SAPbouiCOM.DBDataSource
    Dim objMatrix, objMatrix1 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim oDt As SAPbouiCOM.DataTable
    Dim str, str1 As String
    Public rs, frs, RsNum As SAPbobsCOM.Recordset
    Dim LostFocusFlag As Boolean = False
    Dim objutilities As Utilities
    Dim oItem As SAPbouiCOM.Item

    Dim PrgBar As SAPbouiCOM.ProgressBar


#End Region
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("OPQT")
                    objMatrix = objForm.Items.Item("38").Specific
                    If pVal.ItemUID = "38" And pVal.ColUID = "1" And pVal.BeforeAction = False Then
                        For i As Integer = 1 To objMatrix.VisualRowCount - 1
                            If objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value <> "" Then
                                objMatrix = objForm.Items.Item("38").Specific
                                Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value
                                Dim GetDocEntry As String = "Select ""U_EXCVAT"" From OITM Where ""ItemCode"" = '" & ItemCode & "' "
                                Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRs.DoQuery(GetDocEntry)
                                objMatrix.Columns.Item("U_EXCTAX").Cells.Item(i).Specific.Value() = oRs.Fields.Item("U_EXCVAT").Value()
                            End If
                        Next
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub AddItems(ByVal FormUID As String)
        Try

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


End Class