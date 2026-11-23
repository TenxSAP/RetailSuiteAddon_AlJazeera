Imports SAPbobsCOM
Public Class Cls_MainWHSE

#Region "       Declaration             "
    Public objForm As SAPbouiCOM.Form
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim objMatrix, objMatrix1 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim str, str1 As String
    Public rs, RsNum As SAPbobsCOM.Recordset
    Dim LostFocusFlag As Boolean = False
    Dim objutilities As Utilities
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("MainWHSExpDT.xml", "MWNEDSR", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("MWNEDSR", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            objutilities = New Utilities


            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "MWNEDSR" And pVal.BeforeAction = False Then
                Me.CreateForm()
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then

            End If

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try

            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUID)

                    If pVal.ItemUID = "Load" And pVal.BeforeAction = False Then

                        Dim oGrid As SAPbouiCOM.Grid
                        Dim oDataTable As SAPbouiCOM.DataTable

                        Try
                            oGrid = CType(objForm.Items.Item("Item_0").Specific, SAPbouiCOM.Grid)

                            If objForm.DataSources.DataTables.Count = 0 Then
                                objForm.DataSources.DataTables.Add("dtGrid")
                            End If

                            oDataTable = objForm.DataSources.DataTables.Item("dtGrid")

                            Dim recordSet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim query As String = "Call ""TNX_MainWHSNearExpiryDtStock""()"
                            'recordSet.DoQuery(query)
                            oDataTable.ExecuteQuery(query)

                            oGrid.DataTable = oDataTable

                            oGrid.AutoResizeColumns()

                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText("Please check the Procedure 'TNX_MainWHSNearExpiryDtStock'", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        End Try


                        objForm.Freeze(False)

                    End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

End Class
