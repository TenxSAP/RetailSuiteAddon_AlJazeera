Imports SAPbobsCOM

Public Class ClsStockDetails

#Region "Decliration"

    Public objForm As SAPbouiCOM.Form

    'Dim oDs, oDS1 As SAPbouiCOM.DBDataSource

    Dim objContentMx, objMatrix As SAPbouiCOM.Matrix

    Dim objComboBox As SAPbouiCOM.ComboBox

    Public oDBs_Head, oDBs_Details, oDBs_HeaderF, oDBs_Content, oDBs_Summary, oDBs_Attacht, oDS, oDS1 As SAPbouiCOM.DBDataSource

    Public Prgbar As SAPbouiCOM.ProgressBar

    Dim objutilities As Utilities

#End Region

    Sub CreateForm()

        Try

            objMain.objUtilities.LoadForm(“StockDetails.xml”, "STOCKDE", ResourceType.Embeded)

            objForm = objMain.objApplication.Forms.GetForm("STOCKDE", objMain.objApplication.Forms.ActiveForm.TypeCount)

            'objForm.Visible = True

            objForm.EnableMenu("774", True)

            objForm.EnableMenu("1292", True)

            objForm.Freeze(True)

            objutilities = New Utilities

            oDBs_HeaderF = objForm.DataSources.DBDataSources.Item("@STOCKD")

            oDBs_Content = objForm.DataSources.DBDataSources.Item("@C_STOCKD")

            objContentMx = objForm.Items.Item("STMTX").Specific

            'chnaged

            'oDBs_HeaderF.SetValue("DocNum", oDBs_HeaderF.Offset, objMain.objUtilities.GetNextDocNum(objForm, "STOCKUDO", "Primary"))

            'oDBs_HeaderF.SetValue("U_DOCDT", 0, DateTime.Now.ToString("yyyyMMdd"))

            ' Me.SetNewLineContent(objForm.UniqueID)

            ' objForm.DataBrowser.BrowseBy = "DocNum"

            Me.SetDefault(objForm.UniqueID)

            Try

                Dim objComboBox As SAPbouiCOM.ComboBox

                objComboBox = objForm.Items.Item("BName").Specific

                Dim sQury As String = " Select T0.""BPLId"" , T1.""BPLName"" From USR6 T0 " &
                          " Inner Join OBPL T1 On T1.""BPLId"" = T0.""BPLId"" " &
                          " Where ""UserCode"" = '" & objMain.objCompany.UserName & "' " &
                          " Order By T0.""BPLId"" "

                objMain.objUtilities.ComboBoxLoadValues(objComboBox, sQury)


            Catch ex As Exception

                objMain.objApplication.StatusBar.SetText(ex.Message)

            End Try


            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("DOCDT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DOCDT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DOCDT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DOCDT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)


            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText("Successfully initialized, please procced..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception

            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)

        Try

            If pVal.MenuUID = "YD1" And pVal.BeforeAction = False Then

                Me.CreateForm()

            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then

                Me.SetDefault(objForm.UniqueID)

            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then

                Me.SetDefault(objForm.UniqueID)

                'Me.SetNewLineContent(objForm.UniqueID)

            ElseIf pVal.MenuUID = "774" And pVal.BeforeAction = False Then

                objMatrix = objForm.Items.Item("STMTX").Specific

                Dim row As Integer = objMatrix.VisualRowCount

                If objMatrix.IsRowSelected(1) <> True And objMatrix.VisualRowCount < 1 Then

                    objMatrix.AddRow()

                    oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)

                    objMatrix.SetLineData(objMatrix.VisualRowCount)

                End If

                If objMatrix.IsRowSelected(row) = True Then

                    objMatrix.DeleteRow(row)

                Else

                    For i As Integer = 1 To objMatrix.VisualRowCount - 1

                        If objMatrix.IsRowSelected(i) = True Then

                            objMatrix.DeleteRow(i)

                        End If

                    Next

                End If

                For i As Integer = 1 To objMatrix.VisualRowCount

                    objMatrix.Columns.Item("LineId").Cells.Item(i).Specific.Value = i

                Next

                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then

                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                End If

            End If

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)

        Try

            Select Case pVal.EventType

                'Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                        Me.SetDefault(objForm.UniqueID)

                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                    If pVal.ItemUID = "STMTX" And pVal.ColUID = "PreSto" Then

                        If pVal.Row = objContentMx.VisualRowCount Then

                            Me.SetNewLineContent(objForm.UniqueID)

                        End If

                    End If

                    '========================================================================='

                    'Stock Details Matrix

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    oDBs_HeaderF = objForm.DataSources.DBDataSources.Item("@STOCKD")

                    objContentMx = objForm.Items.Item("STMTX").Specific

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

                        'akila

                        'If oCFL.UniqueID = "CFL_ItemC" And pVal.BeforeAction = False Then

                        '    Dim lineId As String = objContentMx.Columns.Item("LineId").Cells.Item(pVal.Row).Specific.Value

                        '    Dim itemCode As String = oDT.GetValue("ItemCode", 0)

                        '    Me.CFLUOM(objForm.UniqueID, "CFL_UOM", lineId, itemCode)

                        'End If


                        If oCFL.UniqueID = "CFL_ItemC" Then


                            oDBs_Content.SetValue("LineId", oDBs_Content.Offset, pVal.Row)

                            oDBs_Content.SetValue("U_ItemC", oDBs_Content.Offset, oDT.GetValue("ItemCode", 0))

                            oDBs_Content.SetValue("U_Disc", oDBs_Content.Offset, oDT.GetValue("ItemName", 0))


                            'akhila

                            Dim itm As String = oDT.GetValue("ItemCode", 0)  ' 0 is the first row

                            Dim GetBarCode As String = "SELECT TOP 1 ""BcdCode"" FROM OBCD WHERE ""ItemCode"" = '" & itm & "' ORDER BY ""BcdEntry"" ASC"

                            Dim oRsGetBarCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            oRsGetBarCode.DoQuery(GetBarCode)

                            Dim Barcode As String = ""

                            If oRsGetBarCode.RecordCount > 0 Then

                                Barcode = oRsGetBarCode.Fields.Item("BcdCode").Value.ToString()

                                oDBs_Content.SetValue("U_BarC", oDBs_Content.Offset, Barcode)

                            End If

                            oDBs_Content.SetValue("U_InStoc", oDBs_Content.Offset, objContentMx.Columns.Item("InStoc").Cells.Item(pVal.Row).Specific.Value)

                            oDBs_Content.SetValue("U_PreSto", oDBs_Content.Offset, objContentMx.Columns.Item("PreSto").Cells.Item(pVal.Row).Specific.Value)

                            oDBs_Content.SetValue("U_UOM", oDBs_Content.Offset, objContentMx.Columns.Item("UOM").Cells.Item(pVal.Row).Specific.Value)


                            Me.CFLUOM(objForm.UniqueID, "CFL_UOM", itm)


                            objContentMx.SetLineData(pVal.Row)

                        End If

                        If oCFL.UniqueID = "CFL_UOM" Then

                            oDBs_Content.SetValue("LineId", oDBs_Content.Offset, pVal.Row)

                            oDBs_Content.SetValue("U_UOM", oDBs_Content.Offset, oDT.GetValue("UomCode", 0))

                            oDBs_Content.SetValue("U_ItemC", oDBs_Content.Offset, objContentMx.Columns.Item("ItemC").Cells.Item(pVal.Row).Specific.Value)

                            oDBs_Content.SetValue("U_Disc", oDBs_Content.Offset, objContentMx.Columns.Item("Disc").Cells.Item(pVal.Row).Specific.Value)

                            oDBs_Content.SetValue("U_BarC", oDBs_Content.Offset, objContentMx.Columns.Item("BarC").Cells.Item(pVal.Row).Specific.Value)

                            oDBs_Content.SetValue("U_InStoc", oDBs_Content.Offset, objContentMx.Columns.Item("InStoc").Cells.Item(pVal.Row).Specific.Value)

                            oDBs_Content.SetValue("U_PreSto", oDBs_Content.Offset, objContentMx.Columns.Item("PreSto").Cells.Item(pVal.Row).Specific.Value)

                            objContentMx.SetLineData(pVal.Row)

                        End If

                        If oCFL.UniqueID = "CFL_BC" Then

                            oDBs_Content.SetValue("LineId", oDBs_Content.Offset, pVal.Row)

                            oDBs_Content.SetValue("U_BarC", oDBs_Content.Offset, oDT.GetValue("BcdCode", 0))

                            'Akhila

                            Dim bcode As String = oDT.GetValue("BcdCode", 0)  ' 0 is the first row

                            Dim itmqury As String = "SELECT T1.""ItemCode"", T1.""ItemName"", T0.""BcdCode"", T2.""UomCode"" " &
                                                    "FROM ""OBCD"" T0 INNER JOIN ""OITM"" T1 ON T0.""ItemCode"" = T1.""ItemCode"" INNER JOIN OUOM T2 ON " &
                                                        " T0.""UomEntry"" = T2.""UomEntry"" Where T0.""BcdCode"" = '" & bcode & "'"

                            Dim oRitmqury As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            oRitmqury.DoQuery(itmqury)

                            Dim itm As String = ""

                            itm = oRitmqury.Fields.Item("ItemCode").Value.ToString()

                            oDBs_Content.SetValue("U_ItemC", oDBs_Content.Offset, itm)

                            oDBs_Content.SetValue("U_Disc", oDBs_Content.Offset, objContentMx.Columns.Item("Disc").Cells.Item(pVal.Row).Specific.Value)

                            ' oDBs_Content.SetValue("U_BarC", oDBs_Content.Offset, objContentMx.Columns.Item("BarC").Cells.Item(pVal.Row).Specific.Value)

                            oDBs_Content.SetValue("U_InStoc", oDBs_Content.Offset, objContentMx.Columns.Item("InStoc").Cells.Item(pVal.Row).Specific.Value)

                            oDBs_Content.SetValue("U_PreSto", oDBs_Content.Offset, objContentMx.Columns.Item("PreSto").Cells.Item(pVal.Row).Specific.Value)

                            oDBs_Content.SetValue("U_UOM", oDBs_Content.Offset, objContentMx.Columns.Item("UOM").Cells.Item(pVal.Row).Specific.Value)



                            objContentMx.SetLineData(pVal.Row)

                        End If

                    End If

            End Select

        Catch ex As Exception

            'cc

            'objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub



    'Sub CFLUOM(ByVal FormUID As String, ByVal CFL_ID As String, ByVal LineId As String, ByVal ItemCode As String)

    Sub CFLUOM(ByVal FormUID As String, ByVal CFL_ID As String, ByVal ItemCode As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim Selection As String =
            "SELECT DISTINCT U.""UomCode"" " &
            "FROM OITM I " &
            "INNER JOIN UGP1 G ON I.""UgpEntry"" = G.""UgpEntry"" " &
            "INNER JOIN OUOM U ON G.""UomEntry"" = U.""UomEntry"" " &
            "WHERE I.""ItemCode"" = '" & ItemCode & "' " &
            "ORDER BY U.""UomCode"""

            oRs.DoQuery(Selection)

            Dim oConditions As SAPbouiCOM.Conditions

            Dim oCondition As SAPbouiCOM.Condition

            Dim oChooseFromList As SAPbouiCOM.ChooseFromList

            Dim emptyCon As New SAPbouiCOM.Conditions

            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)

            oChooseFromList.SetConditions(emptyCon)

            oConditions = oChooseFromList.GetConditions()

            If oRs.RecordCount > 0 Then

                oCondition = oConditions.Add()

                oCondition.Alias = "UomCode"

                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL

                oCondition.CondVal = oRs.Fields.Item("UomCode").Value

                oChooseFromList.SetConditions(oConditions)

            End If

        Catch ex As Exception

        End Try

    End Sub

    'Sub CFLUOM(ByVal FormUID As String, ByVal CFL_ID As String, ByVal LineId As String, ByVal ItemCode As String)

    '    Try

    '        ' Retrieve the form

    '        Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUID)

    '        ' Query to fetch UOM Codes for the given ItemCode

    '        Dim UOMQry As String =

    '        "SELECT DISTINCT U.""UomCode"" " &

    '        "FROM OITM I " &

    '        "INNER JOIN UGP1 G ON I.""UgpEntry"" = G.""UgpEntry"" " &

    '        "INNER JOIN OUOM U ON G.""UomEntry"" = U.""UomEntry"" " &

    '        "WHERE I.""ItemCode"" = '" & ItemCode & "' " &

    '        "ORDER BY U.""UomCode"""

    '        ' Get Recordset object

    '        Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '        ' Execute the query

    '        rs.DoQuery(UOMQry)

    '        ' Get the ChooseFromList object

    '        Dim oCFL As SAPbouiCOM.ChooseFromList

    '        oCFL = objForm.ChooseFromLists.Item(CFL_ID)

    '        ' Initialize conditions for CFL filtering

    '        Dim oCons As New SAPbouiCOM.Conditions

    '        Dim oCon As SAPbouiCOM.Condition

    '        ' Check if any records were returned

    '        If rs.RecordCount > 0 Then

    '            While Not rs.EoF

    '                ' Add condition for each UOM found

    '                oCon = oCons.Add()

    '                oCon.Alias = "UomCode" ' Use the correct alias for UOM field in the form

    '                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL

    '                oCon.CondVal = rs.Fields.Item("UomCode").Value.ToString()

    '                ' If there are multiple UOMs, apply an OR relationship

    '                If Not rs.EoF Then

    '                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR

    '                End If

    '                rs.MoveNext()

    '            End While

    '            ' Apply the conditions to the CFL

    '            oCFL.SetConditions(oCons)

    '        End If

    '    Catch ex As Exception

    '        ' Handle errors

    '        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '    End Try

    'End Sub

    Sub SetNewLineContent(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_HeaderF = objForm.DataSources.DBDataSources.Item("@STOCKD")

            oDBs_Content = objForm.DataSources.DBDataSources.Item("@C_STOCKD")

            objContentMx = objForm.Items.Item("STMTX").Specific

            objContentMx.AddRow()

            oDBs_Content.SetValue("LineId", oDBs_Content.Offset, objContentMx.VisualRowCount)

            oDBs_Content.SetValue("U_ItemC", oDBs_Content.Offset, "")

            oDBs_Content.SetValue("U_Disc", oDBs_Content.Offset, "")

            oDBs_Content.SetValue("U_UOM", oDBs_Content.Offset, "")

            oDBs_Content.SetValue("U_BarC", oDBs_Content.Offset, "")

            oDBs_Content.SetValue("U_InStoc", oDBs_Content.Offset, "")

            oDBs_Content.SetValue("U_PreSto", oDBs_Content.Offset, "")

            objContentMx.SetLineData(objContentMx.VisualRowCount)

            objContentMx.AutoResizeColumns()

            objMain.objApplication.StatusBar.SetText("Matrix New line Inserted..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub

    Public Shared Sub CFL_Filter(
ByRef pVal As SAPbouiCOM.ItemEvent,
    ByVal Query As String,
ByVal Code As String)

        Try

            Dim oCFL As SAPbouiCOM.ChooseFromList

            Dim oCons As SAPbouiCOM.Conditions

            Dim oCon As SAPbouiCOM.Condition

            Dim oRec As SAPbobsCOM.Recordset

            Dim oCFL_Event As SAPbouiCOM.ChooseFromListEvent

            oRec = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset),
                    SAPbobsCOM.Recordset)

            oRec.DoQuery(Query)

            oCFL_Event = CType(pVal, SAPbouiCOM.ChooseFromListEvent)

            oCFL = objMain.objApplication.Forms.ActiveForm.ChooseFromLists.Item(oCFL_Event.ChooseFromListUID)

            ' Clear old filters

            oCFL.SetConditions(Nothing)

            ' Create new conditions

            oCons = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions)

            If oRec.RecordCount > 0 Then

                For i As Integer = 0 To oRec.RecordCount - 1

                    oCon = oCons.Add()

                    oCon.Alias = Code          ' "UomEntry"

                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL

                    oCon.CondVal = oRec.Fields.Item(Code).Value.ToString()

                    If i < oRec.RecordCount - 1 Then

                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR

                    End If

                    oRec.MoveNext()

                Next

            Else

                ' Force empty CFL

                oCon = oCons.Add()

                oCon.Alias = Code

                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL

                oCon.CondVal = "-1"

            End If

            oCFL.SetConditions(oCons)

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

    Sub SetDefault(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_HeaderF = Me.objForm.DataSources.DBDataSources.Add("@STOCKD")

            oDBs_Content = Me.objForm.DataSources.DBDataSources.Add("@C_STOCKD")

            objContentMx = Me.objForm.Items.Item("STMTX").Specific

            'Chnged

            'oDBs_HeaderF.SetValue("DocNum", oDBs_HeaderF.Offset, objMain.objUtilities.GetNextDocNum(objForm, "STOCKUDO", "Primary"))

            ' objForm.DataBrowser.BrowseBy = "DocNum"

            'Akhila

            Dim Query As String = "SELECT IFNULL(MAX(""DocNum""),0) + 1 AS ""NextDocNum"" FROM ""@STOCKD"""

            ' Dim Query As String = "SELECT IFNULL(MAX(""DocNum"") + 1 AS ""NextDocNum"" FROM ""@STOCKD"""

            Dim rs As SAPbobsCOM.Recordset

            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            rs.DoQuery(Query)

            If Not rs.EoF Then

                Dim nextDocNum As Integer = rs.Fields.Item("NextDocNum").Value

                oDBs_HeaderF.SetValue("DocNum", oDBs_HeaderF.Offset, nextDocNum)

            End If

            oDBs_HeaderF.SetValue("U_DOCDT", oDBs_HeaderF.Offset, DateTime.Now.ToString("yyyyMMdd"))

            Me.SetNewLineContent(objForm.UniqueID)

            objContentMx.AutoResizeColumns()

            objForm.Freeze(False)

        Catch ex As Exception

            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub



End Class

