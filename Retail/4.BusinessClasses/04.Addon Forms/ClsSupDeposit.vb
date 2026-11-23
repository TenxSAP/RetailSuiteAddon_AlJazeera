Public Class ClsSupDeposit
    Public objForm, objITRForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
        Dim objMatrix As SAPbouiCOM.Matrix
        Dim objComboBox As SAPbouiCOM.ComboBox
        Public rs, RsNum As SAPbobsCOM.Recordset
        Dim LostFocusFlag As Boolean = False


        Sub CreateForm()
            Try
            objMain.objUtilities.LoadForm("SupplierDeposit.xml", "SDPTR", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("SDPTR", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_SDPT")

            objMatrix = objForm.Items.Item("Mtx").Specific
            objForm.EnableMenu("1292", True)
            objMatrix.AutoResizeColumns()

            Me.MatrixLoad(objForm.UniqueID)
            'Me.MatrixLoad1(objForm.UniqueID)


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
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_SDPT")
            objMatrix = objForm.Items.Item("Mtx").Specific

            objMatrix.AddRow()

            oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.VisualRowCount)
            oDBs_Head.SetValue("Code", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("Name", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_GLAcct", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_SupChrg", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_Itmlst", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_SupCtr", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_VATACC", oDBs_Head.Offset, "")


            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Sub MatrixLoad(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.EnableMenu("1292", True)
            Dim rs1 As String = "select ""Code"",""Name"",""U_GLAcct"", ""U_VATACC"", ""U_SupChrg"", ""U_Itmlst"", ""U_SupCtr"" from ""@TNX_SDPT"" order by ""DocEntry"" asc"
            Dim ors1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ors1.DoQuery(rs1)
            If ors1.RecordCount > 0 Then
                For j As Integer = 1 To ors1.RecordCount
                    objMatrix.AddRow()
                    oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.VisualRowCount)
                    oDBs_Head.SetValue("Code", oDBs_Head.Offset, ors1.Fields.Item("Code").Value)
                    oDBs_Head.SetValue("Name", oDBs_Head.Offset, ors1.Fields.Item("Name").Value)
                    oDBs_Head.SetValue("U_GLAcct", oDBs_Head.Offset, ors1.Fields.Item("U_GLAcct").Value)
                    oDBs_Head.SetValue("U_SupChrg", oDBs_Head.Offset, ors1.Fields.Item("U_SupChrg").Value)
                    oDBs_Head.SetValue("U_Itmlst", oDBs_Head.Offset, ors1.Fields.Item("U_Itmlst").Value)
                    oDBs_Head.SetValue("U_SupCtr", oDBs_Head.Offset, ors1.Fields.Item("U_SupCtr").Value)
                    oDBs_Head.SetValue("U_VATACC", oDBs_Head.Offset, ors1.Fields.Item("U_VATACC").Value)

                    objMatrix.SetLineData(objMatrix.VisualRowCount)
                    ors1.MoveNext()
                Next
            End If
            objMatrix.FlushToDataSource()
            objMatrix.LoadFromDataSource()
            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

        Catch ex As Exception
            Me.objForm.Freeze(False)
            Me.objForm.StatusBar.SetText(ex.Message & "   Page Load  Events", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "SDPTR" And pVal.BeforeAction = False Then
                Me.CreateForm()

            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then


            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("SDPTR", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix = objForm.Items.Item("Mtx").Specific
                Me.SetNewLine(objForm.UniqueID)
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            ElseIf pVal.MenuUID = "1293" And pVal.BeforeAction = False Then
                objMatrix = objForm.Items.Item("SDPTR").Specific
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


    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects

                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                    If oCFL.UniqueID = "CFL_0" And pVal.BeforeAction = True Then

                    End If
                    If oCFL.UniqueID = "CFL_1" And pVal.BeforeAction = True Then
                    End If

                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                        If oCFL.UniqueID = "CFL_0" Then

                            oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMatrix.Columns.Item("Code").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Head.SetValue("Name", oDBs_Head.Offset, objMatrix.Columns.Item("Name").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Head.SetValue("U_GLAcct", oDBs_Head.Offset, oDT.GetValue("AcctCode", 0))
                            oDBs_Head.SetValue("U_VATACC", oDBs_Head.Offset, objMatrix.Columns.Item("VATACC").Cells.Item(pVal.Row).Specific.Value)

                            'oDBs_Head.SetValue("U_SupChrg", oDBs_Head.Offset, objMatrix.Columns.Item("SupChrg").Cells.Item(pVal.Row).Specific.Checked)
                            'oDBs_Head.SetValue("U_Itmlst", oDBs_Head.Offset, objMatrix.Columns.Item("Itmlst").Cells.Item(pVal.Row).Specific.Checked)
                            'oDBs_Head.SetValue("U_SupCtr", oDBs_Head.Offset, objMatrix.Columns.Item("SupCtr").Cells.Item(pVal.Row).Specific.Checked)

                            objMatrix.SetLineData(pVal.Row)
                        End If




                        If oCFL.UniqueID = "CFL_1" Then

                            oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMatrix.Columns.Item("Code").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Head.SetValue("Name", oDBs_Head.Offset, objMatrix.Columns.Item("Name").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Head.SetValue("U_GLAcct", oDBs_Head.Offset, objMatrix.Columns.Item("GLAcct").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Head.SetValue("U_VATACC", oDBs_Head.Offset, oDT.GetValue("AcctCode", 0))

                            'oDBs_Head.SetValue("U_SupChrg", oDBs_Head.Offset, objMatrix.Columns.Item("SupChrg").Cells.Item(pVal.Row).Specific.Checked)
                            'oDBs_Head.SetValue("U_Itmlst", oDBs_Head.Offset, objMatrix.Columns.Item("Itmlst").Cells.Item(pVal.Row).Specific.Checked)
                            'oDBs_Head.SetValue("U_SupCtr", oDBs_Head.Offset, objMatrix.Columns.Item("SupCtr").Cells.Item(pVal.Row).Specific.Checked)

                            objMatrix.SetLineData(pVal.Row)
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    'objForm = objMain.objApplication.Forms.Item(FormUID)
                    'oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_COU")
                    'objMatrix = objForm.Items.Item("Mtx2").Specific
                    'If pVal.ItemUID = "Mtx2" And pVal.ColUID = "Code" Then
                    '    If objMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim() <> "" Then
                    '        If pVal.Row = objMatrix.VisualRowCount Then
                    '            'Me.SetNewLine(objForm.UniqueID)
                    '        End If
                    '    End If
                    'End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Sub MatrixLoad1(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.EnableMenu("1292", True)
            oDBs_Head.Query()
            Dim j As Integer
            For i = 0 To oDBs_Head.Size - 1
                oDBs_Head.Offset = i
                j = 0
                j = j + 1
                objMatrix.AddRow(j)
            Next i
            objMatrix.FlushToDataSource()
            objMatrix.LoadFromDataSource()

            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
        Catch ex As Exception
            Me.objForm.Freeze(False)
            Me.objForm.StatusBar.SetText(ex.Message & "   Page Load  Events", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
            Try
            objForm = objMain.objApplication.Forms.GetForm("SDPTR", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                        If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                        If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then

                    End If
            End Select

            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText(ex.Message)
            End Try
        End Sub


    Sub CFLFilter(ByVal FormUID As String, ByVal CFL_ID As String)
            Try
                objForm = objMain.objApplication.Forms.Item(FormUID)

                Dim oConditions As SAPbouiCOM.Conditions
                Dim oCondition As SAPbouiCOM.Condition
                Dim oChooseFromList As SAPbouiCOM.ChooseFromList
                Dim emptyCon As New SAPbouiCOM.Conditions
                oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
                oChooseFromList.SetConditions(emptyCon)
                oConditions = oChooseFromList.GetConditions()
                oCondition = oConditions.Add()
                oCondition.Alias = "U_Status"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = "Active"
                oChooseFromList.SetConditions(oConditions)

            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText(ex.Message)
            End Try
        End Sub


    End Class
