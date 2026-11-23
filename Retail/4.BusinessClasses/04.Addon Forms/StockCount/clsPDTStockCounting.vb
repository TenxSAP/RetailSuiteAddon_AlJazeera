Public Class clsPDTStockCounting

#Region "       Declaration             "
    Public objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim BaseUom As String
    Private SelectedRow As Integer = 0
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("PDTStockCounting.xml", "ME_PSC_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_PSC_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            objForm.EnableMenu("1292", False)
            objForm.EnableMenu("774", False)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PSC")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PSC_C0")
            objMatrix = objForm.Items.Item("29").Specific

            objComboBox = objForm.Items.Item("20").Specific
            objComboBox.ValidValues.Add("", "")
            objComboBox.ValidValues.Add("Open", "Open")
            objComboBox.ValidValues.Add("Close", "Close")
            objComboBox.ValidValues.Add("Cancel", "Cancel")
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("20").DisplayDesc = True

            objComboBox = objForm.Items.Item("4").Specific
            objComboBox.ValidValues.Add("All", "All")
            objComboBox.ValidValues.Add("ItemGroup", "ItemGroup")
            objComboBox.ValidValues.Add("SubGroup", "SubGroup")
            objComboBox.ValidValues.Add("SubSubGroup", "SubSubGroup")
            objForm.Items.Item("4").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            Dim sQury As String = "Select ""ItmsGrpCod"" , ""ItmsGrpNam""   From OITB  "
            objComboBox = objForm.Items.Item("6").Specific
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, sQury)
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("6").DisplayDesc = True

            'Dim sQury1 As String = "Select ""BPLId"" ,  ""BPLName"" From OBPL "
            Dim sQury1 As String = " Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join " &
                           " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""  "
            objComboBox = objForm.Items.Item("34").Specific
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, sQury1)
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("34").DisplayDesc = True

            Me.CellsMasking(objForm.UniqueID)
            objMatrix.Columns.Item("V_1").Visible = False
            objMatrix.Columns.Item("V_6").Visible = False
            objMatrix.Columns.Item("V_7").Visible = False
            objMatrix.Columns.Item("V_8").Visible = False
            objMatrix.Columns.Item("V_9").Visible = False

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub CellsMasking(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Items.Item("18").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("18").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("31").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("31").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("31").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ME_PSC" And pVal.BeforeAction = False Then
                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
            End If

            If (pVal.MenuUID = "Add Row" OrElse pVal.MenuUID = "1292") AndAlso pVal.BeforeAction = False Then
                Try
                    objForm = objMain.objApplication.Forms.ActiveForm
                    If objForm Is Nothing OrElse objForm.TypeEx <> "ME_PSC_Form" Then
                        objForm = objMain.objApplication.Forms.GetForm("ME_PSC_Form", 1)
                    End If
                Catch
                    Try
                        objForm = objMain.objApplication.Forms.GetForm("ME_PSC_Form", 1)
                    Catch
                    End Try
                End Try

                If objForm IsNot Nothing AndAlso objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE AndAlso objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                    objForm.Freeze(True)
                    Try
                        Me.SetNewLine(objForm.UniqueID, "29")
                        objMatrix = objForm.Items.Item("29").Specific
                        For i As Integer = 1 To objMatrix.VisualRowCount
                            Try
                                objMatrix.Columns.Item("V_-1").Cells.Item(i).Specific.String = i.ToString()
                            Catch
                            End Try
                        Next

                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        End If

                        Try
                            If objMatrix.VisualRowCount > 0 Then
                                objMatrix.Columns.Item("V_0").Cells.Item(objMatrix.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            End If
                        Catch
                        End Try
                    Finally
                        objForm.Freeze(False)
                    End Try
                End If
            End If

            If (pVal.MenuUID = "Delete Row" OrElse pVal.MenuUID = "774") AndAlso pVal.BeforeAction = False Then
                Try
                    objForm = objMain.objApplication.Forms.ActiveForm
                    If objForm Is Nothing OrElse objForm.TypeEx <> "ME_PSC_Form" Then
                        objForm = objMain.objApplication.Forms.GetForm("ME_PSC_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                    End If
                Catch
                    objForm = objMain.objApplication.Forms.GetForm("ME_PSC_Form", 1)
                End Try

                If objForm IsNot Nothing AndAlso objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE AndAlso objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                    objMatrix = objForm.Items.Item("29").Specific
                    Dim rowDeleted As Boolean = False
                    For i As Integer = objMatrix.VisualRowCount To 1 Step -1
                        If objMatrix.IsRowSelected(i) = True Then
                            objMatrix.DeleteRow(i)
                            rowDeleted = True
                        End If
                    Next
                    If Not rowDeleted AndAlso SelectedRow > 0 AndAlso SelectedRow <= objMatrix.VisualRowCount Then
                        objMatrix.DeleteRow(SelectedRow)
                        rowDeleted = True
                    End If
                    If rowDeleted Then
                        For i As Integer = 1 To objMatrix.VisualRowCount
                            Try
                                objMatrix.Columns.Item("V_-1").Cells.Item(i).Specific.String = i.ToString()
                            Catch
                            End Try
                        Next
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                    SelectedRow = 0
                End If
            End If

            If pVal.MenuUID = "Close" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("ME_PSC_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix = objForm.Items.Item("29").Specific
                Dim result As Integer = objMain.objApplication.MessageBox("Changing document is irreversible..Confirm... Do you want to continue?", 1, "Continue", "Cancel")
                If (result = 1) Then
                    Me.UpdateStatus(objForm.UniqueID)
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub UpdateStatus(ByVal FormUID As String)
        Try
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PSC")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PSC_C0")
            objMatrix = objForm.Items.Item("29").Specific

            objMain.sCmp = objMain.objCompany.GetCompanyService
            objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OPSC")
            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0))
            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
            objMain.oGeneralData.SetProperty("U_Status", "Close")
            objMain.oGeneralService.Update(objMain.oGeneralData)
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_PSC"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_PSC_C0"), "DocEntry", oDBs_Details.GetValue("DocEntry", 0))
            objMatrix.LoadFromDataSource()

            objForm.Refresh()
            objMain.objApplication.StatusBar.SetText("DocStatus updated successfully...!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "", Optional ByVal Series As Integer = 0)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            objForm.EnableMenu("1292", False)
            objForm.EnableMenu("774", False)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PSC")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PSC_C0")
            objMatrix = objForm.Items.Item("29").Specific

            objComboBox = objForm.Items.Item("20").Specific
            objComboBox.Select("Open", SAPbouiCOM.BoSearchKey.psk_ByValue)

            objComboBox = objForm.Items.Item("31").Specific
            Dim GetSeriesQty As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'ME_OPSC' And ""Locked"" = 'N'"
            objMain.objUtilities.ComboBoxLoadValuesSeries(objForm.Items.Item("31").Specific, GetSeriesQty)
            objForm.Items.Item("31").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            ''Dim GetDflt As String = "Select ""DfltSeries"" From ONNM Where ""ObjectCode"" = 'ME_OPSC'"
            Dim GetDflt As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'ME_OPSC' And ""Locked"" = 'N'"
            Dim oRsGetDflt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDflt.DoQuery(GetDflt)

            If Flag = "" Then
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNumSeries(objForm, "ME_OPSC", oRsGetDflt.Fields.Item(0).Value))
                oDBs_Head.SetValue("U_Series", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)
                oDBs_Head.SetValue("Series", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)
                oDBs_Head.SetValue("U_CountDate", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
            Else
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNumSeries(objForm, "ME_OPSC", Series))
                oDBs_Head.SetValue("U_Series", oDBs_Head.Offset, Series)
                oDBs_Head.SetValue("Series", oDBs_Head.Offset, Series)
                oDBs_Head.SetValue("U_Status", oDBs_Head.Offset, "Open")
                oDBs_Head.SetValue("U_CountDate", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
            End If

            objMatrix.AutoResizeColumns()

            objMatrix.Clear()
            oDBs_Details.Clear()
            objMatrix.FlushToDataSource()

            Me.SetNewLine(objForm.UniqueID, "29")

            objMatrix.AutoResizeColumns()
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal MatrixUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PSC")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PSC_C0")
            objMatrix = objForm.Items.Item("29").Specific
            Select Case MatrixUID
                Case "29"
                    objMatrix.AddRow()
                    oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
                    oDBs_Details.SetValue("U_Barcode", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_ItemCode", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_ItemName", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_UOMCode", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_UomName", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_BaseUom", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_ItmPrUnt", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_Qty", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_LineNo", oDBs_Details.Offset, "")
                    objMatrix.SetLineData(objMatrix.VisualRowCount)
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub AddBlankRows(ByVal FormUID As String, ByVal RowCount As Integer)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PSC")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PSC_C0")
            objMatrix = objForm.Items.Item("29").Specific

            Dim startRow As Integer = objMatrix.VisualRowCount + 1
            For i As Integer = 1 To RowCount
                Me.SetNewLine(FormUID, "29")
            Next

            For i As Integer = 1 To objMatrix.VisualRowCount
                Try
                    objMatrix.Columns.Item("V_-1").Cells.Item(i).Specific.String = i.ToString()
                Catch
                End Try
            Next

            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If

            Try
                If objMatrix.VisualRowCount >= startRow Then
                    objMatrix.Columns.Item("V_0").Cells.Item(startRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                End If
            Catch
            End Try

            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(RowCount.ToString() & " blank row(s) added successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            Try
                objForm.Freeze(False)
            Catch
            End Try
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PSC")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PSC_C0")
                    objMatrix = objForm.Items.Item("29").Specific
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim LoadSameForm As String = "Select MAX(""DocEntry"") FROM ""@ME_PSC"""
                        Dim oRsLoadSameForm As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsLoadSameForm.DoQuery(LoadSameForm)
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objForm.Items.Item("30").Specific.Value = oRsLoadSameForm.Fields.Item(0).Value
                        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PSC")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PSC_C0")
                    objMatrix = objForm.Items.Item("29").Specific

                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                    'If oCFL.UniqueID = "CFL_BCD" And pVal.BeforeAction = True Then
                    '    Me.CFLBarCodeCodeFilter(objForm.UniqueID, "CFL_BCD", objForm.Items.Item("4").Specific.Selected.Value)
                    'End If

                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                        If oCFL.UniqueID = "CFL_BCD" Then
                            Dim GetDetails As String = ""
                            If objForm.Items.Item("4").Specific.Selected.Value = "All" Then
                                GetDetails = " Select T0.""ItemCode"",T0.""ItemName"",T3.""ItmsGrpNam"",T0.""FrgnName"",T2.""UomCode"" ,  T1.""UomEntry"" , T0.""InvntryUom"" " &
                                          " From ""OITM"" T0 INNER JOIN ""OBCD"" T1 ON T0.""ItemCode"" =T1.""ItemCode"" INNER JOIN ""OUOM"" T2 ON T1.""UomEntry"" =T2.""UomEntry"" " &
                                          " INNER JOIN ""OITB"" T3 ON T0.""ItmsGrpCod"" =T3.""ItmsGrpCod"" WHERE T1.""BcdCode"" = '" & oDT.GetValue("BcdCode", 0) & "'  "
                            ElseIf objForm.Items.Item("4").Specific.Selected.Value = "ItemGroup" Then
                                GetDetails = " Select T0.""ItemCode"",T0.""ItemName"",T3.""ItmsGrpNam"",T0.""FrgnName"",T2.""UomCode"" ,  T1.""UomEntry"" , T0.""InvntryUom"" " &
                                          " From ""OITM"" T0 INNER JOIN ""OBCD"" T1 ON T0.""ItemCode"" =T1.""ItemCode"" INNER JOIN ""OUOM"" T2 ON T1.""UomEntry"" =T2.""UomEntry"" " &
                                          " INNER JOIN ""OITB"" T3 ON T0.""ItmsGrpCod"" =T3.""ItmsGrpCod"" WHERE T1.""BcdCode"" =  '" & oDT.GetValue("BcdCode", 0) & "' And T0.""ItmsGrpCod"" = '" & objForm.Items.Item("6").Specific.Value & "'  "
                            ElseIf objForm.Items.Item("4").Specific.Selected.Value = "SubGroup" Then
                                GetDetails = " Select T0.""ItemCode"",T0.""ItemName"",T3.""ItmsGrpNam"",T0.""FrgnName"",T2.""UomCode"" ,  T1.""UomEntry"" , T0.""InvntryUom"" " &
                                            " From ""OITM"" T0 INNER JOIN ""OBCD"" T1 ON T0.""ItemCode"" =T1.""ItemCode"" INNER JOIN ""OUOM"" T2 ON T1.""UomEntry"" =T2.""UomEntry"" " &
                                            " INNER JOIN ""OITB"" T3 ON T0.""ItmsGrpCod"" =T3.""ItmsGrpCod"" WHERE T1.""BcdCode"" =  '" & oDT.GetValue("BcdCode", 0) & "' And T0.""ItmsGrpCod"" = '" & objForm.Items.Item("6").Specific.Value & "' And   T0.""U_VSubGroup"" = '" & objForm.Items.Item("8").Specific.Value & "' "
                            ElseIf objForm.Items.Item("4").Specific.Selected.Value = "SubSubGroup" Then
                                GetDetails = " Select T0.""ItemCode"",T0.""ItemName"",T3.""ItmsGrpNam"",T0.""FrgnName"",T2.""UomCode"" ,  T1.""UomEntry"" , T0.""InvntryUom"" " &
                                             " From ""OITM"" T0 INNER JOIN ""OBCD"" T1 ON T0.""ItemCode"" =T1.""ItemCode"" INNER JOIN ""OUOM"" T2 ON T1.""UomEntry"" =T2.""UomEntry"" " &
                                             " INNER JOIN ""OITB"" T3 ON T0.""ItmsGrpCod"" =T3.""ItmsGrpCod"" WHERE T1.""BcdCode"" =  '" & oDT.GetValue("BcdCode", 0) & "' And T0.""ItmsGrpCod"" = '" & objForm.Items.Item("6").Specific.Value & "' And   T0.""U_VSubGroup"" = '" & objForm.Items.Item("8").Specific.Value & "'  " &
                                             " And   T0.""U_VSubSubGroup"" = '" & objForm.Items.Item("10").Specific.Value & "'  "
                            End If
                            Dim oRsGetDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetDetails.DoQuery(GetDetails)

                            If oRsGetDetails.RecordCount = 0 Then
                                objMain.objApplication.StatusBar.SetText("No data match with itemgroup provided..")
                                Exit Sub
                            End If

                            Dim BaseQty As String = " Select T1.""BaseQty""  from ""OITM"" T0   " &
                                                    " INNER JOIN ""UGP1""  T1 ON T0.""UgpEntry""  = T1.""UgpEntry"" INNER JOIN ""OUOM""  T2 ON T2.""UomEntry""  = T1.""UomEntry"" " &
                                                    " WHERE T0.""ItemCode""  ='" & oRsGetDetails.Fields.Item("ItemCode").Value & "'  And  T2.""UomCode""  = '" & oRsGetDetails.Fields.Item("UomCode").Value & "' "
                            Dim oRsBaseQty As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsBaseQty.DoQuery(BaseQty)

                            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row)
                            oDBs_Details.SetValue("U_Barcode", oDBs_Details.Offset, oDT.GetValue("BcdCode", 0))
                            oDBs_Details.SetValue("U_ItemCode", oDBs_Details.Offset, oRsGetDetails.Fields.Item("ItemCode").Value)
                            oDBs_Details.SetValue("U_ItemName", oDBs_Details.Offset, oRsGetDetails.Fields.Item("ItemName").Value)
                            oDBs_Details.SetValue("U_UOMCode", oDBs_Details.Offset, oRsGetDetails.Fields.Item("UomEntry").Value)
                            oDBs_Details.SetValue("U_UomName", oDBs_Details.Offset, oRsGetDetails.Fields.Item("UomCode").Value)
                            oDBs_Details.SetValue("U_BaseUom", oDBs_Details.Offset, oRsGetDetails.Fields.Item("InvntryUom").Value)
                            oDBs_Details.SetValue("U_ItmPrUnt", oDBs_Details.Offset, oRsBaseQty.Fields.Item(0).Value)
                            oDBs_Details.SetValue("U_Qty", oDBs_Details.Offset, "")
                            oDBs_Details.SetValue("U_LineNo", oDBs_Details.Offset, pVal.Row)
                            objMatrix.SetLineData(pVal.Row)

                            If pVal.Row = objMatrix.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "29")
                            End If
                        End If

                        If oCFL.UniqueID = "CFL_WHSCD" Then
                            oDBs_Head.SetValue("U_WhseCode", oDBs_Head.Offset, oDT.GetValue("WhsCode", 0))
                        End If

                    End If

                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PSC")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@ME_PSC_C0")
                    objMatrix = objForm.Items.Item("29").Specific
                    If pVal.ItemUID = "31" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("31").Specific.Selected.Value <> "" Then
                            Dim series As Integer = objForm.Items.Item("31").Specific.Selected.Value
                            SetDefault(objForm.UniqueID, "CobmoSelect", series)
                        End If
                    End If
                    If pVal.ItemUID = "34" And pVal.BeforeAction = False Then
                        objComboBox = objForm.Items.Item("34").Specific
                        Me.CFLWhsCodeFilter(objForm.UniqueID, "CFL_WHSCD", objComboBox.Value)
                    End If
                    If pVal.ItemUID = "4" And pVal.BeforeAction = False Then
                        Select Case objForm.Items.Item("4").Specific.Selected.Value
                            Case "All"
                                objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                                objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                                objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            Case "ItemGroup"
                                objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                ' objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                                objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                                objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


                            Case "SubGroup"
                                objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                            Case "SubSubGroup"
                                objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                        End Select
                    End If
                    If pVal.ItemUID = "6" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("6").Specific.Selected.Value <> "" Then
                            Dim sQury As String = "SELECT T0.""Code"", T0.""Name"" FROM ""@VR_SUBGROUP"" T0 WHERE T0.""U_ITEMGROUP"" ='" & objForm.Items.Item("6").Specific.Selected.Value & "' "
                            objComboBox = objForm.Items.Item("8").Specific
                            objMain.objUtilities.ComboBoxLoadValues(objComboBox, sQury)
                            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                            objForm.Items.Item("8").DisplayDesc = True
                        End If
                    End If
                    If pVal.ItemUID = "8" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("8").Specific.Selected.Value <> "" Then
                            Dim sQury As String = "SELECT T0.""Code"", T0.""Name"" FROM ""@VR_SUBSUBGROUP""  T0 WHERE T0.""U_SUBGROUP"" = '" & objForm.Items.Item("8").Specific.Selected.Value & "' "
                            objComboBox = objForm.Items.Item("10").Specific
                            objMain.objUtilities.ComboBoxLoadValues(objComboBox, sQury)
                            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                            objForm.Items.Item("10").DisplayDesc = True
                        End If
                    End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub CFLWhsCodeFilter(ByVal FormUID As String, ByVal CFL_ID As String, ByVal BPLid As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim GetWhsCode As String = " Select ""WhsCode"" From OWHS Where ""BPLid"" = '" & BPLid & "' And IFNULL(""U_Transit"",'') <> 'Yes' "
            Dim oRsGetWhsCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetWhsCode.DoQuery(GetWhsCode)

            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            If oRsGetWhsCode.RecordCount > 0 Then
                oCondition = oConditions.Add()
                oCondition.Alias = "WhsCode"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = oRsGetWhsCode.Fields.Item(0).Value
                oChooseFromList.SetConditions(oConditions)
                oRsGetWhsCode.MoveNext()
                For i As Integer = 1 To oRsGetWhsCode.RecordCount - 1
                    oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCondition = oConditions.Add()
                    oCondition.Alias = "WhsCode"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = oRsGetWhsCode.Fields.Item(0).Value
                    oChooseFromList.SetConditions(oConditions)
                    oRsGetWhsCode.MoveNext()
                Next
            Else
                oCondition = oConditions.Add()
                oCondition.Alias = "WhsCode"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = ""
                oChooseFromList.SetConditions(oConditions)
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    'Sub CFLEmpFilter(ByVal FormUID As String, ByVal CFL_ID As String, ByVal ItemGroup As String, ByVal SubGroup As String, ByVal SubSubGroup As String)
    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)
    '        Dim oConditions As SAPbouiCOM.Conditions
    '        Dim oCondition As SAPbouiCOM.Condition
    '        Dim oChooseFromList As SAPbouiCOM.ChooseFromList
    '        Dim emptyCon As New SAPbouiCOM.Conditions
    '        oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
    '        oChooseFromList.SetConditions(emptyCon)
    '        oConditions = oChooseFromList.GetConditions()
    '        If ItemGroup <> "" Then
    '            oCondition = oConditions.Add()
    '            oCondition.Alias = "ItmsGrpCod"
    '            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '            oCondition.CondVal = ItemGroup
    '            oChooseFromList.SetConditions(oConditions)
    '        End If
    '        If SubGroup <> "" Then
    '            oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
    '            oCondition = oConditions.Add()
    '            oCondition.Alias = "SubGroup"
    '            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '            oCondition.CondVal = SubGroup
    '            oChooseFromList.SetConditions(oConditions)
    '        End If
    '        If SubSubGroup <> "" Then
    '            oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
    '            oCondition = oConditions.Add()
    '            oCondition.Alias = "SubSubGroup"
    '            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '            oCondition.CondVal = SubSubGroup
    '            oChooseFromList.SetConditions(oConditions)
    '        End If
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub


    'Sub CFLBarCodeCodeFilter(ByVal FormUID As String, ByVal CFL_ID As String, ByVal Type As String)
    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)
    '        Dim GetBarcode As String = ""
    '        If Type = "ALL" Then
    '            Exit Sub
    '        ElseIf Type = "ItemGroup" Then
    '            If objForm.Items.Item("6").Specific.Value = "" Then
    '                Exit Sub
    '            End If
    '            GetBarcode = " Select T1.""BcdCode""  From ""OITM"" T0 INNER JOIN ""OBCD"" T1 ON T0.""ItemCode"" =T1.""ItemCode"" INNER JOIN ""OUOM"" T2 " & _
    '                         " ON T1.""UomEntry"" = T2.""UomEntry"" INNER JOIN ""OITB"" T3 ON T0.""ItmsGrpCod"" =T3.""ItmsGrpCod"" " & _
    '                         " Where  T0.""ItmsGrpCod"" = '" & objForm.Items.Item("6").Specific.Value & "' "
    '        ElseIf Type = "SubGroup" Then
    '            If objForm.Items.Item("6").Specific.Value = "" And objForm.Items.Item("8").Specific.Value = "" Then
    '                Exit Sub
    '            End If
    '            GetBarcode = " Select T1.""BcdCode""  From ""OITM"" T0 INNER JOIN ""OBCD"" T1 ON T0.""ItemCode"" =T1.""ItemCode"" INNER JOIN ""OUOM"" T2 " & _
    '                         " ON T1.""UomEntry"" = T2.""UomEntry"" INNER JOIN ""OITB"" T3 ON T0.""ItmsGrpCod"" =T3.""ItmsGrpCod"" " & _
    '                         " Where  T0.""ItmsGrpCod"" = '" & objForm.Items.Item("6").Specific.Value & "' And T0.""U_VSubGroup"" = '" & objForm.Items.Item("8").Specific.Value & "' "
    '        ElseIf Type = "SubSubGroup" Then
    '            If objForm.Items.Item("6").Specific.Value = "" And objForm.Items.Item("8").Specific.Value = "" And objForm.Items.Item("10").Specific.Value = "" Then
    '                Exit Sub
    '            End If
    '            GetBarcode = " Select T1.""BcdCode""  From ""OITM"" T0 INNER JOIN ""OBCD"" T1 ON T0.""ItemCode"" =T1.""ItemCode"" INNER JOIN ""OUOM"" T2 " & _
    '                         " ON T1.""UomEntry"" = T2.""UomEntry"" INNER JOIN ""OITB"" T3 ON T0.""ItmsGrpCod"" =T3.""ItmsGrpCod"" " & _
    '                         " Where  T0.""ItmsGrpCod"" = '" & objForm.Items.Item("6").Specific.Value & "' And T0.""U_VSubGroup"" = '" & objForm.Items.Item("8").Specific.Value & "' And T0.""U_VSubSubGroup"" = '" & objForm.Items.Item("10").Specific.Value & "' "
    '        End If
    '        Dim oRsGetBarcode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oRsGetBarcode.DoQuery(GetBarcode)

    '        Dim oConditions As SAPbouiCOM.Conditions
    '        Dim oCondition As SAPbouiCOM.Condition
    '        Dim oChooseFromList As SAPbouiCOM.ChooseFromList
    '        Dim emptyCon As New SAPbouiCOM.Conditions
    '        oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
    '        oChooseFromList.SetConditions(emptyCon)
    '        oConditions = oChooseFromList.GetConditions()
    '        If oRsGetBarcode.RecordCount > 0 Then
    '            oCondition = oConditions.Add()
    '            oCondition.Alias = "BcdCode"
    '            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '            oCondition.CondVal = oRsGetBarcode.Fields.Item(0).Value
    '            oChooseFromList.SetConditions(oConditions)
    '            oRsGetBarcode.MoveNext()
    '            For i As Integer = 1 To oRsGetBarcode.RecordCount - 1
    '                oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
    '                oCondition = oConditions.Add()
    '                oCondition.Alias = "BcdCode"
    '                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '                oCondition.CondVal = oRsGetBarcode.Fields.Item(0).Value
    '                oChooseFromList.SetConditions(oConditions)
    '                oRsGetBarcode.MoveNext()
    '            Next
    '        Else
    '            oCondition = oConditions.Add()
    '            oCondition.Alias = "BcdCode"
    '            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '            oCondition.CondVal = ""
    '            oChooseFromList.SetConditions(oConditions)
    '        End If
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub

#Region " Right Click Event"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Dim objForm As SAPbouiCOM.Form
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        Try
            objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)
            If eventInfo.FormUID = objForm.UniqueID Then
                SelectedRow = eventInfo.Row
                If (eventInfo.BeforeAction = True) Then
                    objForm.EnableMenu("1292", False)
                    objForm.EnableMenu("774", False)
                    If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                        oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                        oMenus = oMenuItem.SubMenus

                        If eventInfo.ItemUID = "29" Then
                            objMatrix = objForm.Items.Item("29").Specific
                            Try
                                If oMenus.Exists("Add Row") = False Then
                                    oCreationPackage.UniqueID = "Add Row"
                                    oCreationPackage.String = "Add Row"
                                    oCreationPackage.Enabled = True
                                    oMenus.AddEx(oCreationPackage)
                                End If
                                If eventInfo.Row > 0 AndAlso objMatrix.VisualRowCount > 1 Then
                                    If oMenus.Exists("Delete Row") = False Then
                                        oCreationPackage.UniqueID = "Delete Row"
                                        oCreationPackage.String = "Delete Row"
                                        oCreationPackage.Enabled = True
                                        oMenus.AddEx(oCreationPackage)
                                    End If
                                Else
                                    If oMenus.Exists("Delete Row") = True Then
                                        objMain.objApplication.Menus.RemoveEx("Delete Row")
                                    End If
                                End If
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        Else
                            Try
                                If oMenus.Exists("Close") = False Then
                                    oCreationPackage.UniqueID = "Close"
                                    oCreationPackage.String = "Close"
                                    oCreationPackage.Enabled = True
                                    oMenus.AddEx(oCreationPackage)
                                End If
                                If oMenus.Exists("Delete Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Delete Row")
                                End If
                                If oMenus.Exists("Add Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Add Row")
                                End If
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

End Class
