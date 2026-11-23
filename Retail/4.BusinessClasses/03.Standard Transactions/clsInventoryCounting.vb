
Public Class clsInventoryCounting

#Region "Declaration"
    Public objForm As SAPbouiCOM.Form
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oCheckBox As SAPbouiCOM.CheckBox
    Dim oItem As SAPbouiCOM.Item
#End Region

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.BeforeAction = False Then
                        Me.AddItems(objForm.UniqueID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.BeforeAction = False And pVal.ItemUID = "chkBox" Then
                        Me.RunCheck(objForm.UniqueID)
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub RunCheck(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            objMatrix = objForm.Items.Item("1470000021").Specific
            oCheckBox = objForm.Items.Item("chkBox").Specific
            If oCheckBox.Checked = True Then
                If objMatrix.VisualRowCount > 0 Then
                    For i As Integer = 1 To objMatrix.VisualRowCount - 1
                        objMatrix.Columns.Item("1470000011").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    Next
                End If                
            End If
            If oCheckBox.Checked = False Then
                If objMatrix.VisualRowCount > 0 Then
                    For i As Integer = 1 To objMatrix.VisualRowCount - 1
                        objMatrix.Columns.Item("1470000011").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    Next
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            objForm.Freeze(False)
        End Try
    End Sub

    Private Sub AddItems(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.DataSources.UserDataSources.Add("U_CHK", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objMain.objUtilities.AddCheckBox(objForm.UniqueID, "chkBox", objForm.Items.Item("1470000006").Top, objForm.Items.Item("1470000006").Left + objForm.Items.Item("1470000006").Width + 5, objForm.Items.Item("1470000006").Width, "", "U_CHK", "&Select", "1470000004")
            oCheckBox = objForm.Items.Item("chkBox").Specific
            oCheckBox.DataBind.SetBound(True, "", "U_CHK")
            objForm.Items.Item("chkBox").AffectsFormMode = False

            'oItem = objForm.Items.Item("1470000006")
            'Dim sEditText As SAPbouiCOM.EditText
            'sEditText = oItem.Specific            
            'oItem.LinkTo = "chkBox"
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
