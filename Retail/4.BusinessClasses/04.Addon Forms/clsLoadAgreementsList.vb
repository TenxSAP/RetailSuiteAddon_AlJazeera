Public Class clsLoadAgreementsList

#Region "Declarations"
    Dim objForm, objBaseForm As SAPbouiCOM.Form
    Dim objMatrix, objBaseMatrix As SAPbouiCOM.Matrix
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objGrid As SAPbouiCOM.Grid
    Dim objGridColumn As SAPbouiCOM.GridColumn
    Dim objGridColumns As SAPbouiCOM.GridColumns
    Dim objComboGrid As SAPbouiCOM.ComboBoxColumn
    Dim objCheckBox As SAPbouiCOM.CheckBox
    Dim oDt As SAPbouiCOM.DataTable
    Dim sCardCode As String
#End Region

    Sub CreateForm(ByVal FormUID As String, ByVal CardCode As String, ByVal CheckValue As String)
        Try
            objMain.objUtilities.LoadForm("LoadAgreementsList.xml", "ME_LAL_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_LAL_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objBaseForm = objMain.objApplication.Forms.Item(FormUID)

            'objForm.DataSources.UserDataSources.Add("SelectAll", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            'objCheckBox = objForm.Items.Item("4").Specific
            'objCheckBox.DataBind.SetBound(True, "", "SelectAll")
            'objForm.Items.Item("4").AffectsFormMode = False
            objForm.Items.Item("5").Specific.Value() = CardCode
            'objForm.Items.Item("5").Visible = False
            oDt = objForm.DataSources.DataTables.Add("dt")
            oDt = objForm.DataSources.DataTables.Item("dt")

            sCardCode = CardCode

            Me.LoadGrid(objForm.UniqueID, sCardCode, "N")

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub LoadGrid(ByVal FormUID As String, ByVal BpCode As String, ByVal CheckValue As String)
        Try
            objForm.Freeze(True)
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objGrid = objForm.Items.Item("3").Specific
            objGrid.DataTable = oDt
            objGrid.DataTable.Clear()
            '------NareshV1
            'Dim Query As String = " Select '" & CheckValue & "'  ""CheckBox"" , ""AbsID"" , ""BpCode"" , ""BpName"" , ""StartDate"" , ""EndDate"" From ""OOAT"" Where ""BpCode"" = '" & BpCode & "' "

            Dim Query As String = " Select ""AbsID"" as ""Contract Number"" , ""BpCode"" as ""Vendor Code"", ""BpName"" as ""Vendor Name"", ""StartDate""  as ""Contract Start Date"", ""EndDate"" as ""Contract End Date"", ""TermDate""as ""Contract Termination Date"" From ""OOAT"" Where ""BpCode"" = '" & BpCode & "' "
            objGrid.DataTable.ExecuteQuery(Query)

            If objGrid.Rows.Count = 0 Then
                objMain.objApplication.MessageBox("No records found")
                Exit Sub
            End If
            For i As Integer = 0 To objGrid.Columns.Count - 1
                objGrid.Columns.Item(i).Editable = False
            Next

            'objGrid.Columns.Item("CheckBox").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            'objGrid.Columns.Item("CheckBox").Editable = True
            '-------NareshV1
            Dim oEditTextCol As SAPbouiCOM.EditTextColumn
            oEditTextCol = objGrid.Columns.Item("Contract Number")
            oEditTextCol.LinkedObjectType = 1250000025
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
                            Me.LoadGrid(objForm.UniqueID, objForm.Items.Item("5").Specific.Value(), "Y")
                        ElseIf objForm.Items.Item("4").Specific.checked = False Then
                            Me.LoadGrid(objForm.UniqueID, objForm.Items.Item("5").Specific.Value(), "N")
                        End If
                        objForm.Freeze(False)
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False Then
                        Me.LoadMatrix(objForm.UniqueID)
                        objForm.Close()
                    End If
                    '--------NareshV1
                    If pVal.ItemUID = "New" And pVal.BeforeAction = False Then
                        objMain.objApplication.ActivateMenuItem("39722")
                        objBaseForm = objMain.objApplication.Forms.GetForm("1250000102", objMain.objApplication.Forms.ActiveForm.TypeCount)
                        objBaseForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                        objBaseForm.Items.Item("1250000006").Specific.Value = objForm.Items.Item("5").Specific.Value
                    End If
                    '--------NareshV1
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub LoadMatrix(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objBaseMatrix = objBaseForm.Items.Item("15").Specific
            oDBs_Head = objBaseForm.DataSources.DBDataSources.Item("@ME_SUPAG")
            oDBs_Details = objBaseForm.DataSources.DBDataSources.Item("@ME_SUPAG_C0")

            For i As Integer = 0 To objGrid.Rows.Count - 1
                If objGrid.DataTable.Columns.Item("CheckBox").Cells.Item(i).Value.Equals("Y") Then

                    oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objBaseMatrix.VisualRowCount)
                    oDBs_Details.SetValue("U_PRAGGNO", oDBs_Details.Offset, objGrid.DataTable.Columns.Item("AbsID").Cells.Item(i).Value)
                    Dim StartDate As Date = objGrid.DataTable.Columns.Item("StartDate").Cells.Item(i).Value
                    oDBs_Details.SetValue("U_STRDT", oDBs_Details.Offset, StartDate.ToString("yyyyMMdd"))
                    Dim EndDate As Date = objGrid.DataTable.Columns.Item("EndDate").Cells.Item(i).Value
                    oDBs_Details.SetValue("U_ENDDT", oDBs_Details.Offset, EndDate.ToString("yyyyMMdd"))
                    oDBs_Details.SetValue("U_STATUS", oDBs_Details.Offset, "Active")
                    oDBs_Details.SetValue("U_REMARKS", oDBs_Details.Offset, "")
                    objBaseMatrix.SetLineData(objBaseMatrix.VisualRowCount)
                    objBaseMatrix.AddRow()
                    oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objBaseMatrix.VisualRowCount)
                    oDBs_Details.SetValue("U_PRAGGNO", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_STRDT", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_ENDDT", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_STATUS", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_REMARKS", oDBs_Details.Offset, "")
                    objBaseMatrix.SetLineData(objBaseMatrix.VisualRowCount)
                End If
            Next
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
End Class
