Public Class ClsPettyCashMaster
    Public objForm, objITRForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Public rs, RsNum As SAPbobsCOM.Recordset
    Dim LostFocusFlag As Boolean = False

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("PettyCashTypeMaster.xml", "PCTMR", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("PCTMR", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCTM")

            objMatrix = objForm.Items.Item("Mtx").Specific
            objForm.EnableMenu("1292", True)
            objForm.EnableMenu("1304", True)
            ''Me.SetNewLine(objForm.UniqueID)
            objMatrix.AutoResizeColumns()
            Me.MatrixLoad(objForm.UniqueID)

            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCTM")
            objMatrix = objForm.Items.Item("Mtx").Specific

            objMatrix.AddRow()

            oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.VisualRowCount)
            oDBs_Head.SetValue("Code", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_Type", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_ACCT", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_Name", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_ACT", oDBs_Head.Offset, "Yes")

            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub MatrixLoad(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCTM")
            objMatrix = objForm.Items.Item("Mtx").Specific

            objForm.EnableMenu("1292", True)

            Dim strQuery As String = "SELECT ""Code"", ""U_Type"", ""U_ACCT"", ""U_Name"", ""U_ACT"" FROM ""@TNX_PCTM"" ORDER BY ""DocEntry"""

            Dim ors1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ors1.DoQuery(strQuery)

            If ors1.RecordCount > 0 Then

                objMatrix.Clear()

                For j As Integer = 1 To ors1.RecordCount

                    objMatrix.AddRow()

                    oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.VisualRowCount)
                    oDBs_Head.SetValue("Code", oDBs_Head.Offset, ors1.Fields.Item("Code").Value)
                    oDBs_Head.SetValue("U_Type", oDBs_Head.Offset, ors1.Fields.Item("U_Type").Value)
                    oDBs_Head.SetValue("U_ACCT", oDBs_Head.Offset, ors1.Fields.Item("U_ACCT").Value)
                    oDBs_Head.SetValue("U_Name", oDBs_Head.Offset, ors1.Fields.Item("U_Name").Value)
                    oDBs_Head.SetValue("U_ACT", oDBs_Head.Offset, ors1.Fields.Item("U_ACT").Value)

                    objMatrix.SetLineData(objMatrix.VisualRowCount)

                    ors1.MoveNext()

                Next

                objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

            Else

                Me.SetNewLine(objForm.UniqueID)

            End If

            objMatrix.FlushToDataSource()

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message & " Page Load Events", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then


                        Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        For i As Integer = 1 To objMatrix.VisualRowCount

                            Dim CurrCode As String = objMatrix.Columns.Item("Code").Cells.Item(i).Specific.Value.Trim()
                            Dim DocEntry As String = objMatrix.Columns.Item("DocEntry").Cells.Item(i).Specific.Value.Trim()

                            'Skip newly added rows
                            If DocEntry = "" Then Continue For

                            rs.DoQuery("SELECT ""Code"" FROM ""@TNX_PCTM"" WHERE ""DocEntry"" = '" & DocEntry & "'")

                            'If rs.RecordCount > 0 Then

                            '    Dim DBCode As String = rs.Fields.Item("Code").Value.ToString.Trim()

                            '    If DBCode <> CurrCode Then

                            '        objMain.objApplication.StatusBar.SetText("Code cannot be modified. Row : " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                            '        BubbleEvent = False
                            '        Exit Sub

                            '    End If

                            'End If

                        Next

                    End If


                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCTM")
                    objMatrix = objForm.Items.Item("Mtx").Specific

                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)

                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects

                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then

                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        End If

                        If oCFL.UniqueID = "CFL_0" Then

                            oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMatrix.Columns.Item("Code").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Head.SetValue("U_Type", oDBs_Head.Offset, objMatrix.Columns.Item("Type").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Head.SetValue("U_ACCT", oDBs_Head.Offset, oDT.GetValue("AcctCode", 0).ToString())
                            oDBs_Head.SetValue("U_Name", oDBs_Head.Offset, oDT.GetValue("AcctName", 0).ToString())
                            ''oDBs_Head.SetValue("U_ACT", oDBs_Head.Offset, objMatrix.Columns.Item("Active").Cells.Item(pVal.Row).Specific.Value)

                            objMatrix.SetLineData(pVal.Row)

                        End If

                    End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "PCTMR" And pVal.BeforeAction = False Then
                Me.CreateForm()

            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then


            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("PCTMR", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix = objForm.Items.Item("Mtx").Specific
                Me.SetNewLine(objForm.UniqueID)
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            ElseIf pVal.MenuUID = "1293" And pVal.BeforeAction = False Then
                objMatrix = objForm.Items.Item("Mtx").Specific
                For i As Integer = 1 To objMatrix.VisualRowCount - 1
                    If objMatrix.IsRowSelected(i) = True Then
                        objMatrix.DeleteRow(i)
                    End If
                Next
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.Columns.Item("DocEntry").Cells.Item(i).Specific.Value = i
                Next

            End If
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
End Class
