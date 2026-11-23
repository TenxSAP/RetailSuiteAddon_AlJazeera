Imports SAPbobsCOM

Public Class Cls_StockCntSummary


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
            objMain.objUtilities.LoadForm("StockCntSumm.xml", "STCKCNT", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("STCKCNT", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            oDBs_Head = Me.objForm.DataSources.DBDataSources.Add("@TNX_STCKCNT")

            objutilities = New Utilities

            'Dim objTypeCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("BName").Specific

            'If objTypeCombo.ValidValues.Count > 0 Then
            '    For i As Integer = objTypeCombo.ValidValues.Count - 1 To 0 Step -1
            '        objTypeCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
            '    Next
            'End If

            'Dim ors As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'Dim Qry As String = "SELECT ""Code"", ""U_SUBGROUP"" FROM ""@VR_SUBGROUP"""
            'ors.DoQuery(Qry)

            'If ors.RecordCount > 0 Then
            '    Do While Not ors.EoF

            '        objTypeCombo.ValidValues.Add(ors.Fields.Item("Code").Value.ToString(), ors.Fields.Item("U_SUBGROUP").Value.ToString())

            '        ors.MoveNext()
            '    Loop
            'End If



            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "STCKCNT" And pVal.BeforeAction = False Then
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


                    Dim Fdate As String = oDBs_Head.GetValue("U_FDate", 0)

                    Dim Tdate As String = oDBs_Head.GetValue("U_TDate", 0)

                    Dim Branch As String = oDBs_Head.GetValue("U_BName", 0)
                    Dim GRPName As String = oDBs_Head.GetValue("U_GRPName", 0)
                    Dim WHSE As String = oDBs_Head.GetValue("U_WHSE", 0)



                    If pVal.ItemUID = "Load" And pVal.BeforeAction = False Then

                        If Me.Validation(objForm.UniqueID) = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If

                        Dim oGrid As SAPbouiCOM.Grid
                        Dim oDataTable As SAPbouiCOM.DataTable

                        Try
                            oGrid = CType(objForm.Items.Item("Item_0").Specific, SAPbouiCOM.Grid)

                            If objForm.DataSources.DataTables.Count = 0 Then
                                objForm.DataSources.DataTables.Add("dtGrid")
                            End If

                            oDataTable = objForm.DataSources.DataTables.Item("dtGrid")

                            Dim recordSet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim query As String = "Call ""TNX_StockCuntingSummrytobeposted""('" & Fdate & "','" & Tdate & "', '" & Branch & "', '" & GRPName & "', '" & WHSE & "') "
                            'recordSet.DoQuery(query)
                            oDataTable.ExecuteQuery(query)

                            oGrid.DataTable = oDataTable

                            oGrid.AutoResizeColumns()

                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText("Please check the Procedure 'TNX_StockCuntingSummrytobeposted'", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        End Try


                        objForm.Freeze(False)

                    End If



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
                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        If oCFL.UniqueID = "CFL_0" Then
                            oDBs_Head.SetValue("U_WHSE", oDBs_Head.Offset, oDT.GetValue("WhsCode", 0))
                        End If

                        If oCFL.UniqueID = "CFL_1" Then
                            oDBs_Head.SetValue("U_GRPName", oDBs_Head.Offset, oDT.GetValue("ItmsGrpNam", 0))
                        End If

                        If oCFL.UniqueID = "CFL_2" Then
                            oDBs_Head.SetValue("U_BName", oDBs_Head.Offset, oDT.GetValue("SubGroup", 0))
                        End If




                    End If


            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_STCKCNT")

            If objForm.Items.Item("FDate").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("From Date Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("TDate").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("To Date Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("WHSE").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Warehouse Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Return False
            ElseIf objForm.Items.Item("GRPName").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Group Name Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Return False
            ElseIf objForm.Items.Item("BName").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("SubGroup Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Return False
            End If

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function


End Class
