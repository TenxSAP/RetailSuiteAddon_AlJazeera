Public Class Class1_FILTER
    Public objForm, objITRForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Public rs, RsNum As SAPbobsCOM.Recordset
    Dim LostFocusFlag As Boolean = False
    Dim objGrid As SAPbouiCOM.Grid
    Dim oDBH, oDBl1 As SAPbouiCOM.DataTable

    Sub CreateForm(ByVal qSTR1 As String)
        Try
            objMain.objUtilities.LoadForm("SearchAndFilter.xml", "SearchAndFilter", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("SearchAndFilter", objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.Freeze(True)
            objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            'objMatrix = objForm.Items.Item("MTX_1").Specific
            'oDBs_Details = objForm.DataSources.DBDataSources.Item("RDR1")
            'Dim RSet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'Dim qStr As String = qSTR1
            'RSet.DoQuery(qStr)
            'If RSet.RecordCount > 0 Then
            '    oDBs_Details.Clear()
            '    objMatrix.Clear()
            '    objMatrix.FlushToDataSource()
            '    Dim i = 0
            '    RSet.MoveFirst()
            '    While Not RSet.EoF
            '        oDBs_Details.InsertRecord(i)
            '        objMain.objApplication.StatusBar.SetText("Fetching Data For Line No ''" & (i + 1).ToString() & " out of Total Line '" & RSet.RecordCount.ToString().ToString().ToString() & "'", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            '        oDBs_Details.SetValue("LineNum", i, (i + 1).ToString())
            '        oDBs_Details.SetValue("ItemCode", i, RSet.Fields.Item("ItemCode").Value.ToString())
            '        oDBs_Details.SetValue("VendorNum", i, RSet.Fields.Item("ItemName").Value.ToString())
            '        oDBs_Details.SetValue("Dscription", i, RSet.Fields.Item("CodeBars").Value.ToString())
            '        i += 1
            '        RSet.MoveNext()
            '    End While
            '    oDBs_Details.SetValue("LineNum", i, (i + 1).ToString())
            '    objMatrix.LoadFromDataSource()
            '    objMatrix.AutoResizeColumns()
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
                End Try
                DetQry = "SELECT '' as ""Check"",""ItemCode"",""ItemName"",""CodeBars"" FROM OITM WHERE ""validFor""='Y' AND ""SellItem""='Y'"
                oDBl1.ExecuteQuery(DetQry)
                objGrid.DataTable = oDBl1
                For i As Integer = 1 To objGrid.Columns.Count - 1
                    objGrid.Columns.Item(i).Editable = False
                Next
                Dim objColumn As SAPbouiCOM.EditTextColumn = objGrid.Columns.Item("Check")
                objColumn.Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox

                Dim oColumns2 As SAPbouiCOM.EditTextColumn = objGrid.Columns.Item("ItemCode")
                oColumns2.LinkedObjectType = "4"
                objGrid.AutoResizeColumns()
            Catch ex As Exception
            End Try
            objMain.objApplication.StatusBar.SetText("Data Fetched Successfully ! ! !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            ' End If
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("RDR1")
                    objMatrix = objForm.Items.Item("MTX_1").Specific

                    If pVal.ItemUID = "BTSHOW" And pVal.BeforeAction = False Then
                        Try
                            objForm.Freeze(True)
                            Dim rSet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim qStr As String = "SELECT * FROM ""@TENX_TEMP2"""
                            rSet.DoQuery(qStr)
                            If rSet.RecordCount > 0 Then
                                oDBs_Details.Clear()
                                objMatrix.Clear()
                                objMatrix.FlushToDataSource()
                                Dim i = 0
                                rSet.MoveFirst()
                                While Not rSet.EoF
                                    oDBs_Details.InsertRecord(i)
                                    oDBs_Details.SetValue("LineNum", i, (i + 1).ToString())
                                    oDBs_Details.SetValue("ItemCode", i, rSet.Fields.Item("U_COLUMN1").Value.ToString())
                                    oDBs_Details.SetValue("VendorNum", i, rSet.Fields.Item("U_COLUMN2").Value.ToString())
                                    oDBs_Details.SetValue("Dscription", i, rSet.Fields.Item("U_COLUMN3").Value.ToString())
                                    i += 1
                                    rSet.MoveNext()
                                End While
                                oDBs_Details.SetValue("LineNum", i, (i + 1).ToString())
                                objMatrix.LoadFromDataSource()
                                objMatrix.AutoResizeColumns()
                            End If
                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                        End Try
                    End If
                    If pVal.ItemUID = "BTSEARCH" And pVal.BeforeAction = False Then
                        Try
                            objForm.Freeze(True)
                            Dim SearchKey As String = objForm.Items.Item("ETSEARCH").Specific.Value.ToString().Trim()
                            Dim rSet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim qStr As String = "SELECT * FROM ""@TENX_TEMP1"" WHERE UPPER(""U_COLUMN1"") LIKE UPPER('%" & SearchKey & "%') " &
                                " UNION  " & " SELECT * FROM ""@TENX_TEMP1"" WHERE UPPER(""U_COLUMN2"") LIKE UPPER('%" & SearchKey & "%') " &
                                " UNION  " & " SELECT * FROM ""@TENX_TEMP1"" WHERE UPPER(""U_COLUMN3"") LIKE UPPER('%" & SearchKey & "%') "

                            rSet.DoQuery(qStr)
                            If rSet.RecordCount > 0 Then
                                oDBs_Details.Clear()
                                objMatrix.Clear()
                                objMatrix.FlushToDataSource()
                                Dim i = 0
                                rSet.MoveFirst()
                                While Not rSet.EoF
                                    oDBs_Details.InsertRecord(i)
                                    oDBs_Details.SetValue("LineNum", i, (i + 1).ToString())
                                    oDBs_Details.SetValue("ItemCode", i, rSet.Fields.Item("U_COLUMN1").Value.ToString())
                                    oDBs_Details.SetValue("VendorNum", i, rSet.Fields.Item("U_COLUMN2").Value.ToString())
                                    oDBs_Details.SetValue("Dscription", i, rSet.Fields.Item("U_COLUMN3").Value.ToString())
                                    i += 1
                                    rSet.MoveNext()
                                End While
                                oDBs_Details.SetValue("LineNum", i, (i + 1).ToString())
                                objMatrix.LoadFromDataSource()
                                objMatrix.AutoResizeColumns()
                            Else
                                oDBs_Details.Clear()
                                objMatrix.Clear()
                                objMatrix.FlushToDataSource()
                            End If
                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                        End Try
                    End If
                    If pVal.ItemUID = "BTNSL" And pVal.BeforeAction = False Then
                        Try
                            objForm.Items.Item("BTSHOW").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            objForm.Close()
                            objMain.objSalesOrder.GetValueFromSearchAndFilter()
                        Catch ex As Exception
                            objForm.Freeze(False)
                        End Try
                    End If


                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    'objMatrix = objForm.Items.Item("MTX_1").Specific
                    objGrid = objForm.Items.Item("grid").Specific
                    If pVal.ItemUID = "grid" And pVal.ColUID = "Check" And pVal.BeforeAction = False Then
                        Try
                            If pVal.Row <> 0 Then
                                Dim oCheck As String = objGrid.DataTable.GetValue("Check", objGrid.GetDataTableRowIndex(pVal.Row)).ToString()
                                Dim Column1 As String = objGrid.DataTable.GetValue("ItemCode", objGrid.GetDataTableRowIndex(pVal.Row)).ToString()
                                If oCheck = "Y" Then
                                    Dim rSet2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim qStr2 As String = "DELETE FROM ""@TENX_TEMP2"" WHERE ""Code""='" & Column1 & "'"
                                    rSet2.DoQuery(qStr2)
                                    Dim rSet1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim qStr1 As String = "INSERT INTO ""@TENX_TEMP2"" SELECT *  FROM ""@TENX_TEMP1"" WHERE ""Code""='" & Column1 & "' "
                                    rSet1.DoQuery(qStr1)
                                ElseIf oCheck = "N" Then
                                    Dim rSet2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim qStr2 As String = "DELETE FROM ""@TENX_TEMP2"" WHERE ""Code""='" & Column1.ToString().ToString() & "'"
                                    rSet2.DoQuery(qStr2)
                                End If

                            End If
                        Catch ex As Exception

                        End Try
                    End If
                    'If pVal.ItemUID = "MTX_1" And pVal.ColUID = "CKECKBOX" And pVal.BeforeAction = False Then
                    '    Try
                    '        objForm.Freeze(True)
                    '        If pVal.Row <> 0 Then
                    '            Dim oCheck As SAPbouiCOM.CheckBox = objMatrix.Columns.Item("CKECKBOX").Cells.Item(pVal.Row).Specific
                    '            Dim Column1 As String = objMatrix.Columns.Item("COLUMN1").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()
                    '            If oCheck.Checked = True Then
                    '                Dim rSet2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                Dim qStr2 As String = "DELETE FROM ""@TENX_TEMP2"" WHERE ""Code""='" & Column1 & "'"
                    '                rSet2.DoQuery(qStr2)
                    '                Dim rSet1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                Dim qStr1 As String = "INSERT INTO ""@TENX_TEMP2"" SELECT *  FROM ""@TENX_TEMP1"" WHERE ""Code""='" & Column1 & "' "
                    '                rSet1.DoQuery(qStr1)
                    '            Else
                    '                Dim rSet2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                Dim qStr2 As String = "DELETE FROM ""@TENX_TEMP2"" WHERE ""Code""='" & Column1.ToString().ToString() & "'"
                    '                rSet2.DoQuery(qStr2)
                    '            End If
                    '        Else
                    '            For i As Integer = 1 To objMatrix.VisualRowCount
                    '                Dim oCheck As SAPbouiCOM.CheckBox = objMatrix.Columns.Item("CKECKBOX").Cells.Item(i).Specific
                    '                Dim Column1 As String = objMatrix.Columns.Item("COLUMN1").Cells.Item(i).Specific.Value.ToString().Trim()
                    '                If oCheck.Checked = True Then
                    '                    Dim rSet2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                    Dim qStr2 As String = "DELETE FROM ""@TENX_TEMP2"" WHERE ""Code""='" & Column1 & "'"
                    '                    rSet2.DoQuery(qStr2)
                    '                    Dim rSet1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                    Dim qStr1 As String = "INSERT INTO ""@TENX_TEMP2"" SELECT * FROM ""@TENX_TEMP1"" WHERE ""Code""='" & Column1 & "' "
                    '                    rSet1.DoQuery(qStr1)
                    '                Else
                    '                    Dim rSet2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                    Dim qStr2 As String = "DELETE FROM ""@TENX_TEMP2"" WHERE ""Code""='" & Column1 & "'"
                    '                    rSet2.DoQuery(qStr2)
                    '                End If
                    '            Next

                    '        End If
                    '        objForm.Freeze(False)
                    '    Catch ex As Exception
                    '        objForm.Freeze(False)
                    '    End Try
                    'End If

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
