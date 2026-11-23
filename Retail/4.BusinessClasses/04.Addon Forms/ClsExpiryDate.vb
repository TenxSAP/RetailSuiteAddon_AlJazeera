Public Class ClsExpiryDate

    Dim objForm, objBaseForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix, objBaseMatrix As SAPbouiCOM.Matrix

    Sub CreateForm(ByVal FormUID As String, ByVal DocType As String, ByVal BaseEntry As String, ByVal BaseNo As String, ByVal oRsData As SAPbobsCOM.Recordset)
        Try
            objMain.objUtilities.LoadForm("ExpiryDateScreen.xml", "TNXEDS_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("TNXEDS_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)

            If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = False Then
                Retail.MainCls.ohtLookUpForm.Add(objForm.UniqueID, FormUID)
            End If

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNXEDS")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNXEDSC0")
            objMatrix = objForm.Items.Item("Item_12").Specific


            objMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single


            objForm.Items.Item("Item_3").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_3").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("Item_7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("Item_9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("Item_5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("Item_11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)



            Dim GetDataAdd As String = "Select ""DocNum"" FROM ""@TNXEDS"" Where ""U_DOCTY""='" & DocType.Trim & "' " &
                                        "And ""U_BENTRY""='" & BaseEntry.Trim & "' And ""U_BNO""='" & BaseNo.Trim & "' And ""DocNum"" IS NOT NULL"

            Dim oRsGetDataAdd As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oRsGetDataAdd.DoQuery(GetDataAdd)

            If oRsGetDataAdd.RecordCount > 0 Then
                objForm.Items.Item("Item_9").Specific.Value = oRsGetDataAdd.Fields.Item(0).Value
                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                Exit Try
            End If


            oDBs_Head.SetValue("U_DOCTY", 0, DocType)
            oDBs_Head.SetValue("U_BENTRY", 0, BaseEntry)
            oDBs_Head.SetValue("U_BNO", 0, BaseNo)
            oDBs_Head.SetValue("U_DOCDT", 0, DateTime.Now.ToString("yyyyMMdd"))
            oDBs_Head.SetValue("DocNum", 0, objMain.objUtilities.GetNextDocNum(objForm, "OTNXEDS"))



            objMatrix.Clear()
            oDBs_Details.Clear()
            objMatrix.FlushToDataSource()

            oRsData.MoveFirst()
            For i As Integer = 1 To oRsData.RecordCount


                objMatrix.AddRow()
                oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
                oDBs_Details.SetValue("U_BRCD", oDBs_Details.Offset, oRsData.Fields.Item(4).Value)
                oDBs_Details.SetValue("U_ITMCD", oDBs_Details.Offset, oRsData.Fields.Item(1).Value)
                oDBs_Details.SetValue("U_ITMNM", oDBs_Details.Offset, oRsData.Fields.Item(2).Value)
                oDBs_Details.SetValue("U_UOM", oDBs_Details.Offset, oRsData.Fields.Item(5).Value)
                oDBs_Details.SetValue("U_EDATE", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_QTY", oDBs_Details.Offset, oRsData.Fields.Item(3).Value)
                oDBs_Details.SetValue("U_RMKS", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_BL", oDBs_Details.Offset, oRsData.Fields.Item(0).Value)
                oDBs_Details.SetValue("U_SR", oDBs_Details.Offset, "No")
                oDBs_Details.SetValue("U_BQTY", oDBs_Details.Offset, "")


                objMatrix.SetLineData(objMatrix.VisualRowCount)

                oRsData.MoveNext()
            Next

            objMatrix.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CreateFormNew(ByVal FormUID As String, ByVal DocType As String, ByVal BaseEntry As String, ByVal BaseNo As String, ByVal ListData As DataTable, ByVal GUID As String, ByVal mode As Integer)
        Try
            objMain.objUtilities.LoadForm("ExpiryDateScreen.xml", "TNXEDS_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("TNXEDS_Form", objMain.objApplication.Forms.ActiveForm.Type)

            If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = False Then
                Retail.MainCls.ohtLookUpForm.Add(objForm.UniqueID, FormUID)
            End If

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNXEDS")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNXEDSC0")
            objMatrix = objForm.Items.Item("Item_12").Specific

            objForm.EnableMenu("1281", False)
            objMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

            objForm.Items.Item("Item_1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("Item_3").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_3").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("Item_7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("Item_9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("Item_5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("Item_11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)



            'Dim GetDataAdd As String = "Select ""DocNum"" FROM ""@TNXEDS"" Where ""U_DOCTY""='" & DocType.Trim & "' " &
            '                            "And ""U_BENTRY""='" & BaseEntry.Trim & "' And ""U_BNO""='" & BaseNo.Trim & "' And ""DocNum"" IS NOT NULL"

            Dim GetDataAdd As String = "Select ""DocNum"" FROM ""@TNXEDS"" Where ""U_DOCTY""='" & DocType.Trim & "' " &
                                        "And ""U_GUID""='" & GUID.Trim & "' And ""U_GUID"" IS NOT NULL And ""DocNum"" IS NOT NULL"

            Dim oRsGetDataAdd As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oRsGetDataAdd.DoQuery(GetDataAdd)

            If oRsGetDataAdd.RecordCount > 0 Then
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                objForm.Items.Item("Item_9").Specific.Value = oRsGetDataAdd.Fields.Item(0).Value
                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                If mode <> 3 Then
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                End If
                objMatrix.AutoResizeColumns()
                Exit Try
            End If


            oDBs_Head.SetValue("U_DOCTY", 0, DocType)
            oDBs_Head.SetValue("U_BENTRY", 0, "")
            oDBs_Head.SetValue("U_BNO", 0, "")
            oDBs_Head.SetValue("U_DOCDT", 0, DateTime.Now.ToString("yyyyMMdd"))
            oDBs_Head.SetValue("DocNum", 0, objMain.objUtilities.GetNextDocNum(objForm, "OTNXEDS"))
            oDBs_Head.SetValue("U_GUID", 0, GUID.Trim)



            objMatrix.Clear()
            oDBs_Details.Clear()
            objMatrix.FlushToDataSource()


            For Each row In ListData.Rows


                objMatrix.AddRow()
                oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
                oDBs_Details.SetValue("U_BRCD", oDBs_Details.Offset, row(4))
                oDBs_Details.SetValue("U_ITMCD", oDBs_Details.Offset, row(1))
                oDBs_Details.SetValue("U_ITMNM", oDBs_Details.Offset, row(2))
                oDBs_Details.SetValue("U_UOM", oDBs_Details.Offset, row(5))
                oDBs_Details.SetValue("U_EDATE", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_QTY", oDBs_Details.Offset, row(3))
                oDBs_Details.SetValue("U_RMKS", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_BL", oDBs_Details.Offset, row(0))
                oDBs_Details.SetValue("U_SR", oDBs_Details.Offset, "No")
                oDBs_Details.SetValue("U_BQTY", oDBs_Details.Offset, row(6))


                objMatrix.SetLineData(objMatrix.VisualRowCount)


            Next

            objMatrix.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub SetNewLine(ByVal FormUID As String, ByVal Row As Integer)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNXEDS")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNXEDSC0")
            objMatrix = objForm.Items.Item("Item_12").Specific

            objMatrix.AddRow()
            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
            oDBs_Details.SetValue("U_BRCD", oDBs_Details.Offset, objMatrix.Columns.Item("Col_0").Cells.Item(Row).Specific.Value)
            oDBs_Details.SetValue("U_ITMCD", oDBs_Details.Offset, objMatrix.Columns.Item("Col_1").Cells.Item(Row).Specific.Value)
            oDBs_Details.SetValue("U_ITMNM", oDBs_Details.Offset, objMatrix.Columns.Item("Col_2").Cells.Item(Row).Specific.Value)
            oDBs_Details.SetValue("U_UOM", oDBs_Details.Offset, objMatrix.Columns.Item("Col_3").Cells.Item(Row).Specific.Value)
            oDBs_Details.SetValue("U_EDATE", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_QTY", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_RMKS", oDBs_Details.Offset, objMatrix.Columns.Item("Col_6").Cells.Item(Row).Specific.Value)
            oDBs_Details.SetValue("U_BL", oDBs_Details.Offset, objMatrix.Columns.Item("Col_7").Cells.Item(Row).Specific.Value)
            oDBs_Details.SetValue("U_SR", oDBs_Details.Offset, "Yes")
            oDBs_Details.SetValue("U_BQTY", oDBs_Details.Offset, objMatrix.Columns.Item("Col_9").Cells.Item(Row).Specific.Value)

            objMatrix.SetLineData(objMatrix.VisualRowCount)

            objMatrix.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNXEDS")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNXEDSC0")
                    objMatrix = objForm.Items.Item("Item_12").Specific

                    If pVal.ItemUID = "Item_14" And pVal.BeforeAction = False And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                        Try
                            objForm.Freeze(True)

                            Dim SelectedRow As Integer = objMatrix.GetNextSelectedRow

                            If SelectedRow = -1 Then
                                objMain.objApplication.StatusBar.SetText("Please select a row to split")
                                objForm.Freeze(False)
                                Exit Try
                            Else
                                Me.SetNewLine(objForm.UniqueID, SelectedRow)
                            End If

                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If

                    If pVal.ItemUID = "Item_13" And pVal.BeforeAction = False And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                        Try
                            objForm.Freeze(True)

                            Dim SelectedRow As Integer = objMatrix.GetNextSelectedRow

                            If SelectedRow = -1 Then
                                objMain.objApplication.StatusBar.SetText("Please select a row to delete")
                                objForm.Freeze(False)
                                Exit Try

                            ElseIf CStr(objMatrix.Columns.Item("Col_8").Cells.Item(SelectedRow).Specific.Value).Trim.ToUpper = "NO" Then
                                objMain.objApplication.StatusBar.SetText("you cannot delete base line...")
                                objForm.Freeze(False)
                                Exit Try
                            Else
                                objMatrix.DeleteRow(SelectedRow)

                                For i As Integer = 1 To objMatrix.VisualRowCount
                                    objMatrix.Columns.Item("#").Cells.Item(i).Specific.Value = i
                                Next


                            End If

                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If


                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Try
                            objForm.Freeze(True)

                            Dim chkIfExist As String = "Select ""DocNum"" From ""@TNXEDS"" Where ""DocEntry""=(Select MAX(""DocEntry"") From ""@TNXEDS"" Where ""Creator""='" & objMain.objCompany.UserName & "')"
                            Dim oRschkIfExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRschkIfExist.DoQuery(chkIfExist)
                            If oRschkIfExist.RecordCount > 0 Then
                                objForm.Freeze(True)
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                objForm.Items.Item("Item_9").Specific.value = oRschkIfExist.Fields.Item("DocNum").Value
                                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                objForm.Freeze(False)
                            End If

                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If


                    If pVal.BeforeAction = True And pVal.ItemUID = "1" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Try
                            objForm.Freeze(True)

                            If CStr(objForm.Items.Item("Item_1").Specific.Value).Trim = "" Then
                                objMain.objApplication.StatusBar.SetText("GUID cannot be left blank")
                                objForm.Freeze(False)
                                BubbleEvent = False
                                Exit Try
                            ElseIf CStr(objForm.Items.Item("Item_3").Specific.Value).Trim = "" Then
                                objMain.objApplication.StatusBar.SetText("DocType cannot be left blank")
                                objForm.Freeze(False)
                                BubbleEvent = False
                                Exit Try
                            End If


                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message)
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
