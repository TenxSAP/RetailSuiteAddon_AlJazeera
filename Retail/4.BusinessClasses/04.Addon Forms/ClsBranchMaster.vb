Public Class ClsBranchMaster

    Public objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix
    Public rs As SAPbobsCOM.Recordset
    Dim objutilities As Utilities

#Region "Create Form"
    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("BranchMaster.xml", "Branch", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("Branch", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            objutilities = New Utilities
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNXBRANCH")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNXBRANCH_C0")
            objMatrix = CType(objForm.Items.Item("Matrix").Specific, SAPbouiCOM.Matrix)
            objForm.Freeze(False)

            objForm.Items.Item("Code").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Code").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Code").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Code").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("BName").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("BName").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("BName").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("BName").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            SetNewLine(objForm.UniqueID)
            objForm.EnableMenu("1282", True)
            objForm.EnableMenu("1292", True)
            objForm.EnableMenu("1293", True)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

#End Region
    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNXBRANCH_C0")
            objMatrix = objForm.Items.Item("Matrix").Specific
            objMatrix.AddRow()
            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)

            oDBs_Details.SetValue("U_EAdd", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Name", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_MobileNo", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Email", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Active", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Remarks", oDBs_Details.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "Branch" And pVal.BeforeAction = False Then
                CreateForm()

            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                SetNewLine(objForm.UniqueID)

            ElseIf pVal.MenuUID = "1293" And pVal.BeforeAction = False Then

                objMatrix = objForm.Items.Item("Matrix").Specific

                For i As Integer = objMatrix.VisualRowCount To 1 Step -1
                    If objMatrix.IsRowSelected(i) Then
                        objMatrix.DeleteRow(i)
                    End If
                Next

                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.Columns.Item("LineId").Cells.Item(i).Specific.Value = i
                Next

            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If pVal.ItemUID = "1" AndAlso pVal.BeforeAction = True AndAlso objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNXBRANCH")
                        Dim BranchCode As String = oDBs_Head.GetValue("U_BCode", 0).Trim()
                        If BranchCode = "" Then
                            objMain.objApplication.StatusBar.SetText("Branch Code cannot be empty")
                            BubbleEvent = False
                            Exit Sub
                        End If

                        Dim oRs As SAPbobsCOM.Recordset
                        oRs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        oRs.DoQuery("SELECT ""Code"" FROM ""@TNXBRANCH"" WHERE ""U_BCode"" = '" & BranchCode & "'")

                        If oRs.RecordCount > 0 Then
                            objMain.objApplication.StatusBar.SetText("This Branch is already mapped to another document. One BranchBCode can have only one document.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If


                        oRs.DoQuery("SELECT IFNULL(MAX(CAST(""Code"" AS INTEGER)),0)+1 AS ""Code"" FROM ""@TNXBRANCH""")
                        oDBs_Head.SetValue("Code", 0, oRs.Fields.Item("Code").Value.ToString())

                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                    If pVal.BeforeAction = False Then
                        Dim oCFL_Event As SAPbouiCOM.IChooseFromListEvent = pVal
                        Dim oDataTable As SAPbouiCOM.DataTable = oCFL_Event.SelectedObjects

                        If oDataTable Is Nothing Then Exit Sub
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNXBRANCH")

                        If oCFL_Event.ChooseFromListUID = "CFL_0" Then

                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            End If
                            oDBs_Head.SetValue("U_BCode", 0, oDataTable.GetValue("BPLId", 0).ToString())
                            oDBs_Head.SetValue("U_BName", 0, oDataTable.GetValue("BPLName", 0).ToString())
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                    If pVal.ItemUID = "Matrix" AndAlso pVal.ColUID = "U_Remarks" AndAlso pVal.BeforeAction = False Then
                        objMatrix = objForm.Items.Item("Matrix").Specific
                        If pVal.Row = objMatrix.VisualRowCount Then
                            Dim val As String = objMatrix.Columns.Item("U_Remarks").Cells.Item(pVal.Row).Specific.Value
                            If val <> "" Then
                                SetNewLine(FormUID)
                            End If
                        End If
                    End If
            End Select

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
