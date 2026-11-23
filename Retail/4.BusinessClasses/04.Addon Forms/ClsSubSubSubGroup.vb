Public Class ClsSubSubSubGroup
#Region "Declaration"
    Public objForm As SAPbouiCOM.Form
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix, objMatrix1 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim str, str1 As String
    Public rs, RsNum As SAPbobsCOM.Recordset
    Dim LostFocusFlag As Boolean = False
    Dim oGrid As SAPbouiCOM.Grid
    Dim oDt As SAPbouiCOM.DataTable
    Dim objutilities As Utilities

#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("SubSubSubGroup.xml", "SSSGMASTER", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("SSSGMASTER", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            objutilities = New Utilities
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_SSSGROUP")
            objMatrix = objForm.Items.Item("Mtx").Specific
            'oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TNX_GLAUDO"))
            objForm.Freeze(False)
            Me.MatrixLoad()
            objForm.EnableMenu("1292", True)
            '  Me.SetNewLine(objForm.UniqueID)
            ' Me.GlAccountCode()
            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "SSSGMASTER" And pVal.BeforeAction = False Then
                Me.CreateForm()
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                objMatrix = objForm.Items.Item("Mtx").Specific
                Me.SetNewLine(objForm.UniqueID)
            End If
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                'Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                '    objForm = objMain.objApplication.Forms.Item(FormUID)
                '    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_GLCCNA")
                '    objMatrix = objForm.Items.Item("Mtx_1").Specific
                '    Dim oCFL As SAPbouiCOM.ChooseFromList
                '    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal        'CHOOSE FORM LIST'
                '    Dim CFL_Id As String
                '    CFL_Id = CFLEvent.ChooseFromListUID
                '    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                '    Dim oDT As SAPbouiCOM.DataTable
                '    oDT = CFLEvent.SelectedObjects
                '    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                '    If oCFL.UniqueID = "CFL_1" And pVal.BeforeAction = True Then
                '    End If
                '    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                '        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                '        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                '        If oCFL.UniqueID = "CFL_1" Then
                '            oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value)
                '            oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMatrix.Columns.Item("Code").Cells.Item(pVal.Row).Specific.Value)
                '            oDBs_Head.SetValue("U_Type", oDBs_Head.Offset, objMatrix.Columns.Item("Type").Cells.Item(pVal.Row).Specific.Value)
                '            'oDBs_Head.SetValue("U_TNXGLC", oDBs_Head.Offset, oDT.GetValue("AcctCode", 0))
                '            'oDBs_Head.SetValue("U_TNXGLN", oDBs_Head.Offset, oDT.GetValue("AcctName", 0))
                '            oDBs_Head.SetValue("U_TNXSTUS", oDBs_Head.Offset, objMatrix.Columns.Item("TNXSTUS").Cells.Item(pVal.Row).Specific.Value)
                '            oDBs_Head.SetValue("U_GLCN", oDBs_Head.Offset, oDT.GetValue("AcctCode", 0))
                '            'oDBs_Head.SetValue("U_GLCNM", oDBs_Head.Offset, oDT.GetValue("AcctName", 0))
                '            objMatrix.SetLineData(pVal.Row)
                '        End If
                '    End If

                '  End If

                'Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                '    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_GLCCNA")
                '    objMatrix = objForm.Items.Item("Mtx_1").Specific
                '    If pVal.ItemUID = "Mtx_1" And pVal.ColUID = "TNXGLC" And pVal.BeforeAction = False Then
                '        objComboBox = objMatrix.Columns.Item("TNXGLC").Cells.Item(pVal.Row).Specific
                '        If objComboBox.Selected.Value <> "" Then
                '            Dim projectid As String = objComboBox.Selected.Value
                '            Dim str As String = "SELECT T0.""AcctCode"", T0.""AcctName"" FROM OACT T0 WHERE T0.""Postable"" = 'Y'"
                '            Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                '            rs.DoQuery(str)
                '            If rs.RecordCount > 0 Then
                '                objMatrix.Columns.Item("TNXGLN").Cells.Item(pVal.Row).Specific.Value = rs.Fields.Item(0).Value
                '            End If
                '        End If
                '    End If

                'Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                '    If pVal.ColUID = "TNXGLC" And pVal.BeforeAction = False Then
                '        objMatrix = objForm.Items.Item("Mtx_1").Specific
                '        'Me.MatrixAddrow()
                '        Dim selectedCode As String = objMatrix.Columns.Item("TNXGLC").Cells.Item(pVal.Row).Specific.Value
                '        Dim Query12 As String = "SELECT T0.""AcctCode"", T0.""AcctName"" FROM OACT T0 WHERE T0.""Postable"" = 'Y'"
                '        Dim Qry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                '        Qry.DoQuery(Query12)
                '        If Qry.RecordCount > 0 Then
                '            Dim Name As String = Qry.Fields.Item("TNXGLN").Value
                '            objMatrix.Columns.Item("TNXGLN").Cells.Item(pVal.Row).Specific.Value = Name
                '        End If
                '    End If
                'Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                '    If pVal.ColUID = "TNXGLC" And pVal.BeforeAction = False Then
                '        Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx_1").Specific
                '        Dim selectedCode As String = objMatrix.Columns.Item("TNXGLC").Cells.Item(pVal.Row).Specific.Value

                '        ' Query to get the GL Name
                '        Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                '        Dim Query As String = "SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = '" & selectedCode & "'"
                '        rs.DoQuery(Query)

                '        If rs.RecordCount > 0 Then
                '            objMatrix.Columns.Item("TNXGLN").Cells.Item(pVal.Row).Specific.Value = rs.Fields.Item("AcctName").Value
                '        End If
                '    End If
                'Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                '    If pVal.ColUID = "TNXGLC" AndAlso pVal.BeforeAction = False Then
                '        Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx_1").Specific
                '        Dim totalRows As Integer = objMatrix.RowCount

                '        ' Loop through all rows
                '        For i As Integer = 1 To totalRows
                '            Dim selectedCode As String = objMatrix.Columns.Item("TNXGLC").Cells.Item(i).Specific.Value

                '            If Not String.IsNullOrEmpty(selectedCode) Then
                '                ' Query to get the GL Name
                '                Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                '                Dim query As String = "SELECT ""AcctName"" FROM OACT WHERE ""AcctCode"" = '" & selectedCode & "'"
                '                rs.DoQuery(query)

                '                ' If GL Code exists, assign its GL Name to TNXGLN
                '                If rs.RecordCount > 0 Then
                '                    objMatrix.Columns.Item("TNXGLN").Cells.Item(i).Specific.Value = rs.Fields.Item("AcctName").Value
                '                End If
                '            End If
                '        Next
                '    End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_SSSGROUP")
            objMatrix = objForm.Items.Item("Mtx").Specific
            objMatrix.AddRow()
            oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.VisualRowCount)
            oDBs_Head.SetValue("Code", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("Name", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_SSGroup", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_Status", oDBs_Head.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
            ' objMatrix.AutoResizeColumns()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub MatrixLoad()
        Try
            oDBs_Head = objForm.DataSources.DBDataSources.Add("@TNX_SSSGROUP")
            Dim objsectMat As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx").Specific
            Dim rs1 As String = " select ""Code"", ""Name"", ""U_SSGroup"", ""U_Status"" from ""@TNX_SSSGROUP"" Order by ""DocEntry"" "
            Dim ors1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ors1.DoQuery(rs1)
            If ors1.RecordCount > 0 Then
                objsectMat.Clear()
                For j As Integer = 1 To ors1.RecordCount
                    objsectMat.AddRow()
                    oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.VisualRowCount)
                    oDBs_Head.SetValue("Code", oDBs_Head.Offset, ors1.Fields.Item("Code").Value)
                    oDBs_Head.SetValue("Name", oDBs_Head.Offset, ors1.Fields.Item("Name").Value)
                    oDBs_Head.SetValue("U_SSGroup", oDBs_Head.Offset, ors1.Fields.Item("U_SSGroup").Value)
                    oDBs_Head.SetValue("U_Status", oDBs_Head.Offset, ors1.Fields.Item("U_Status").Value)
                    objsectMat.SetLineData(objsectMat.VisualRowCount)
                    ors1.MoveNext()
                Next
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
            Else
                Me.SetNewLine(objForm.UniqueID)
            End If
            objMatrix.FlushToDataSource()
            ' objsectMat.AutoResizeColumns()
            ' objsectMat.FlushToDataSource()
            'objMatrix.LoadFromDataSource()

        Catch ex As Exception
            oDBs_Head.Freeze(False)
        End Try
    End Sub
End Class

