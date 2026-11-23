Public Class clsUnitsofMeasure
#Region "  Declaration   "
    Public objForm As SAPbouiCOM.Form
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix


#End Region
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.BeforeAction = False Then
                        objMatrix = objForm.Items.Item("3").Specific
                        objMatrix.Columns.Item("UgpCode").Cells.Item(objMatrix.VisualRowCount).Specific.Value = objMain.objUtilities.GlobalItemNumber
                        objMatrix.Columns.Item("UgpName").Cells.Item(objMatrix.VisualRowCount).Specific.Value = objMain.objUtilities.GlobalItemNumber

                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    'If pVal.ItemUID = "btn_AC" And pVal.BeforeAction = False Then
                    '    oDBs_Head = objForm.DataSources.DBDataSources.Item("OUSR")

                    '    Dim CheckIfTCExist As String = "Select TO_VARCHAR (""DocNum"") From ""@SBO_PUIGA"" Where ""U_USRCD""='" & oDBs_Head.GetValue("USER_CODE", 0) & "'"
                    '    Dim oRsCheckIfTCExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '    oRsCheckIfTCExist.DoQuery(CheckIfTCExist)
                    '    If oRsCheckIfTCExist.Fields.Item(0).Value <> "" Then
                    '        objMain.objPriceUpdateItemGroupAuthorization.CreateForm(objForm.UniqueID, oDBs_Head.GetValue("USER_CODE", 0), oDBs_Head.GetValue("U_NAME", 0), "Y", oRsCheckIfTCExist.Fields.Item(0).Value)
                    '    Else
                    '        objMain.objPriceUpdateItemGroupAuthorization.CreateForm(objForm.UniqueID, oDBs_Head.GetValue("USER_CODE", 0), oDBs_Head.GetValue("U_NAME", 0), "N")
                    '    End If
                    'End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub AddItems(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            'objMain.objUtilities.AddButtonH(objForm.UniqueID, "btn_AC", objForm.Items.Item("200000118").Top, objForm.Items.Item("18").Left + 20, 150, "200000118", "Access Category", 1, 1, 20)

            'objMain.objUtilities.AddButtonH(objForm.UniqueID, "btn_BC", objForm.Items.Item("200000118").Top + 30, objForm.Items.Item("18").Left + 20, 150, "200000118", "Branch Access", 1, 1, 20)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


End Class
