Public Class ClsSearchAndFilter
    Public objForm, objSrcForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Public rs, RsNum As SAPbobsCOM.Recordset
    Dim LostFocusFlag As Boolean = False
    Dim objGrid As SAPbouiCOM.Grid
    Dim oDBH, oDBl1 As SAPbouiCOM.DataTable


    Sub CreateForm(ByVal qSTR1 As String, ByRef SourceForm As SAPbouiCOM.Form)
        Try
            objSrcForm = SourceForm
            objMain.objUtilities.LoadForm("SearchAndFilter.xml", "SearchAndFilter", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("SearchAndFilter", objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.Freeze(True)
            objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            Try
                Dim DetQry As String = ""
                Dim rnd As New Random()
                Dim randomNumber As Integer = rnd.Next(10, 100)
                oDBl1 = objForm.DataSources.DataTables.Add("" + System.DateTime.Now().Day().ToString() + "-" + randomNumber.ToString() + "")
                oDBl1 = objForm.DataSources.DataTables.Item("" + System.DateTime.Now().Day().ToString() + "-" + randomNumber.ToString() + "")
                objGrid = objForm.Items.Item("grid").Specific
                Try
                    If objGrid.Columns.Count > 0 Then
                        objGrid.DataTable.Clear()
                    End If
                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End Try
                'DetQry = "SELECT '' as ""Check"",C.""ItemCode"",C.""ItemName"",A.""BcdCode"" ""BarCode"",B.""UomCode"",C.""OnHand"" FROM OITM C LEFT JOIN OBCD A ON A.""ItemCode""=C.""ItemCode"" " +
                '         " LEFT JOIN OUOM B ON B.""UomEntry""=A.""UomEntry"" WHERE C.""validFor""='Y' AND C.""SellItem""='Y'"
                oDBl1.ExecuteQuery(qSTR1)
                objGrid.DataTable = oDBl1
                For i As Integer = 0 To objGrid.Columns.Count - 1
                    objGrid.Columns.Item(i).Editable = False
                Next
                Dim objColumnCD As SAPbouiCOM.EditTextColumn = objGrid.Columns.Item("Code")
                objColumnCD.Visible = False

                Dim objColumnNM As SAPbouiCOM.EditTextColumn = objGrid.Columns.Item("Name")
                objColumnNM.Visible = False

                Dim objColumn As SAPbouiCOM.EditTextColumn = objGrid.Columns.Item("Check")
                objColumn.Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                objColumn.Editable = True

                Dim oColumns2 As SAPbouiCOM.EditTextColumn = objGrid.Columns.Item("ItemCode")
                oColumns2.LinkedObjectType = "4"
                objGrid.AutoResizeColumns()
                objMain.objApplication.StatusBar.SetText("Data Fetched Successfully ! ! !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

            ' End If
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

                    If pVal.ItemUID = "BTSHOW" And pVal.BeforeAction = False Then
                        Try
                            objForm.Freeze(True)
                            Try
                                Dim DetQry As String = ""
                                Dim rnd As New Random()
                                Dim randomNumber As Integer = rnd.Next(10, 100)
                                oDBl1 = objForm.DataSources.DataTables.Add("" + System.DateTime.Now().Day().ToString() + "-" + randomNumber.ToString() + "")
                                oDBl1 = objForm.DataSources.DataTables.Item("" + System.DateTime.Now().Day().ToString() + "-" + randomNumber.ToString() + "")
                                objGrid = objForm.Items.Item("grid").Specific
                                Try
                                    If objGrid.Columns.Count > 0 Then
                                        objGrid.DataTable.Clear()
                                    End If
                                Catch ex As Exception
                                    objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End Try
                                DetQry = "SELECT ""Code"",""Name"",""U_COLUMN0"" ""Check"",""U_COLUMN1"" ""ItemCode"",""U_COLUMN2"" ""ItemName"", " +
                                            " ""U_COLUMN3"" ""BarCode"",""U_COLUMN4"" ""UomCode"",""U_COLUMN5"" ""OnHand"" FROM ""@TENX_TEMP2"" ORDER BY ""Code"""
                                oDBl1.ExecuteQuery(DetQry)
                                objGrid.DataTable = oDBl1
                                For i As Integer = 0 To objGrid.Columns.Count - 1
                                    objGrid.Columns.Item(i).Editable = False
                                Next
                                Dim objColumnCD As SAPbouiCOM.EditTextColumn = objGrid.Columns.Item("Code")
                                objColumnCD.Visible = False

                                Dim objColumnNM As SAPbouiCOM.EditTextColumn = objGrid.Columns.Item("Name")
                                objColumnNM.Visible = False

                                Dim objColumn As SAPbouiCOM.EditTextColumn = objGrid.Columns.Item("Check")
                                objColumn.Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                                objColumn.Editable = True

                                Dim oColumns2 As SAPbouiCOM.EditTextColumn = objGrid.Columns.Item("ItemCode")
                                oColumns2.LinkedObjectType = "4"
                                objGrid.AutoResizeColumns()
                            Catch ex As Exception
                                objForm.Freeze(False)
                                If ex.Message.Contains("Data Table - duplicate") Then
                                Else
                                    objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If
                            End Try
                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If

                    If pVal.ItemUID = "BTSEARCH" And pVal.BeforeAction = False Then
                        Try
                            objForm.Freeze(True)
                            Dim SearchKey As String = objForm.Items.Item("ETSEARCH").Specific.Value.ToString().Trim()
                            Try
                                Dim DetQry As String = ""
                                Dim rnd As New Random()
                                Dim randomNumber As Integer = rnd.Next(10, 100)
                                oDBl1 = objForm.DataSources.DataTables.Add("" + System.DateTime.Now().Day().ToString() + "-" + randomNumber.ToString() + "")
                                oDBl1 = objForm.DataSources.DataTables.Item("" + System.DateTime.Now().Day().ToString() + "-" + randomNumber.ToString() + "")
                                objGrid = objForm.Items.Item("grid").Specific
                                Try
                                    If objGrid.Columns.Count > 0 Then
                                        objGrid.DataTable.Clear()
                                    End If
                                Catch ex As Exception
                                    objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End Try
                                DetQry = "Select ""Code"",""Name"",""U_COLUMN0"" ""Check"",""U_COLUMN1"" ""ItemCode"",""U_COLUMN2"" ""ItemName"",""U_COLUMN3"" ""BarCode"",""U_COLUMN4"" ""UomCode"",""U_COLUMN5"" ""OnHand"" FROM ""@TENX_TEMP1"" WHERE UPPER(""U_COLUMN1"") LIKE UPPER('%" & SearchKey & "%') " &
                                    " UNION  " & " SELECT ""Code"",""Name"",""U_COLUMN0"" ""Check"",""U_COLUMN1"" ""ItemCode"",""U_COLUMN2"" ""ItemName"",""U_COLUMN3"" ""BarCode"",""U_COLUMN4"" ""UomCode"",""U_COLUMN5"" ""OnHand"" FROM ""@TENX_TEMP1"" WHERE UPPER(""U_COLUMN2"") LIKE UPPER('%" & SearchKey & "%') " &
                                    " UNION  " & " SELECT ""Code"",""Name"",""U_COLUMN0"" ""Check"",""U_COLUMN1"" ""ItemCode"",""U_COLUMN2"" ""ItemName"",""U_COLUMN3"" ""BarCode"",""U_COLUMN4"" ""UomCode"",""U_COLUMN5"" ""OnHand"" FROM ""@TENX_TEMP1"" WHERE UPPER(""U_COLUMN3"") LIKE UPPER('%" & SearchKey & "%')  ORDER BY ""Code"""

                                oDBl1.ExecuteQuery(DetQry)
                                objGrid.DataTable = oDBl1
                                For i As Integer = 0 To objGrid.Columns.Count - 1
                                    objGrid.Columns.Item(i).Editable = False
                                Next
                                Dim objColumnCD As SAPbouiCOM.EditTextColumn = objGrid.Columns.Item("Code")
                                objColumnCD.Visible = False

                                Dim objColumnNM As SAPbouiCOM.EditTextColumn = objGrid.Columns.Item("Name")
                                objColumnNM.Visible = False

                                Dim objColumn As SAPbouiCOM.EditTextColumn = objGrid.Columns.Item("Check")
                                objColumn.Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                                objColumn.Editable = True

                                Dim oColumns2 As SAPbouiCOM.EditTextColumn = objGrid.Columns.Item("ItemCode")
                                oColumns2.LinkedObjectType = "4"
                                objGrid.AutoResizeColumns()
                                objForm.Freeze(False)
                            Catch ex As Exception
                                objForm.Freeze(False)
                                If ex.Message.Contains("Data Table - duplicate") Then
                                Else
                                    objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If
                            End Try
                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If

                    If pVal.ItemUID = "BTNSL" And pVal.BeforeAction = False Then
                        Try
                            objForm.Items.Item("BTSHOW").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            objForm.Close()
                            objMain.objSalesOrder.GetValueFromSearchAndFilter1(objSrcForm)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If


                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    'objMatrix = objForm.Items.Item("MTX_1").Specific
                    objGrid = objForm.Items.Item("grid").Specific
                    If pVal.ItemUID = "grid" And pVal.ColUID = "Check" And pVal.BeforeAction = False Then
                        Try
                            If pVal.Row <> -1 Then
                                Dim oCheck As String = objGrid.DataTable.GetValue("Check", objGrid.GetDataTableRowIndex(pVal.Row)).ToString()
                                Dim Code As String = objGrid.DataTable.GetValue("Code", objGrid.GetDataTableRowIndex(pVal.Row)).ToString()
                                If oCheck = "Y" Then
                                    Dim rSet2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim qStr2 As String = "DELETE FROM ""@TENX_TEMP2"" WHERE ""Code""='" & Code & "'"
                                    rSet2.DoQuery(qStr2)
                                    Dim rSet1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim qStr1 As String = "INSERT INTO ""@TENX_TEMP2"" SELECT *  FROM ""@TENX_TEMP1"" WHERE ""Code""='" & Code & "' "
                                    rSet1.DoQuery(qStr1)

                                    Dim rSet3 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim qStr3 As String = "UPDATE ""@TENX_TEMP2"" SET ""U_COLUMN0""='Y' WHERE ""Code""='" & Code & "' "
                                    rSet3.DoQuery(qStr3)

                                    Dim rSet4 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim qStr4 As String = "UPDATE ""@TENX_TEMP1"" SET ""U_COLUMN0""='Y' WHERE ""Code""='" & Code & "' "
                                    rSet4.DoQuery(qStr4)
                                ElseIf oCheck = "N" Then
                                    Dim rSet2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim qStr2 As String = "DELETE FROM ""@TENX_TEMP2"" WHERE ""Code""='" & Code & "'"
                                    rSet2.DoQuery(qStr2)

                                    Dim rSet4 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim qStr4 As String = "UPDATE ""@TENX_TEMP1"" SET ""U_COLUMN0""='N' WHERE ""Code""='" & Code & "' "
                                    rSet4.DoQuery(qStr4)
                                End If

                            End If
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = True And pVal.BeforeAction = False Then
                        Retail.MainCls.ohtLookUpForm.Remove(objForm.UniqueID)
                    End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


End Class
