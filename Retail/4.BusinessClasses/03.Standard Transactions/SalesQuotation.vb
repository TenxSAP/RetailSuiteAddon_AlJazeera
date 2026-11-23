Public Class SalesQuotation

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
                    objForm.Freeze(True)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("OQUT")
                    oDs1 = objForm.DataSources.DBDataSources.Item("QUT1")
                    oDs2 = objForm.DataSources.DBDataSources.Item("QUT3")
                    objMatrix = objForm.Items.Item("38").Specific
                    If pVal.ItemUID = "38" And (pVal.ColUID = "1" Or pVal.ColUID = "11") And pVal.BeforeAction = False Then
                        For i As Integer = 1 To objMatrix.VisualRowCount - 1

                            If objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value <> "" Then
                                objMatrix = objForm.Items.Item("38").Specific
                                Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value
                                Dim GetDocEntry As String = "Select ""U_EXCVAT"" From OITM Where ""ItemCode"" = '" & ItemCode & "' "
                                Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRs.DoQuery(GetDocEntry)
                                objMatrix.Columns.Item("U_EXCTAX").Cells.Item(i).Specific.Value() = oRs.Fields.Item("U_EXCVAT").Value()

                                Dim ItemCode4 As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.ToString().Trim()
                                Dim QTY4 As Double = objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value.ToString().Trim()
                                Dim EXCTAX4 As Double = objMatrix.Columns.Item("U_EXCTAX").Cells.Item(i).Specific.Value.ToString().Trim()
                                Dim SubTotal4 As String = "ST3"
                                If (QTY4 * EXCTAX4) <> 0 Then
                                    Dim ST4 As Double = QTY4 * EXCTAX4
                                    SubTotal4 = ST4

                                    If objMatrix.Columns.Item("U_EXCTAX").Cells.Item(pVal.Row).Specific.Value <> "" Then
                                        Dim objComboBox As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("111").Cells.Item(pVal.Row).Specific
                                        Dim TaskName As String = objComboBox.Selected.Value
                                        objComboBox.Select("Sales Excise Duty", SAPbouiCOM.BoSearchKey.psk_ByDescription)

                                        Dim objComboBox2 As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("111").Cells.Item(pVal.Row).Specific
                                        Dim TaskName2 As String = objComboBox.Selected.Value
                                        If objMatrix.Columns.Item("111").Cells.Item(i).Specific.Value = TaskName2 Then
                                            objMatrix.Columns.Item("112").Cells.Item(i).Specific.Value = SubTotal4
                                        End If
                                    End If
                                Else
                                    Dim objComboBox As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("111").Cells.Item(pVal.Row).Specific
                                    Dim TaskName As String = objComboBox.Selected.Value
                                    objComboBox.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                End If
                            End If

                        Next
                    End If
                    objForm.Freeze(False)
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType
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
