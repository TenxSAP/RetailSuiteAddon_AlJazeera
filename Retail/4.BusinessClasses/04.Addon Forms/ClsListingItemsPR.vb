
Imports SAPbobsCOM

Public Class ClsListingItemsPR

#Region "        Declaration        "
    Dim objForm, objBaseForm As SAPbouiCOM.Form
    Dim oGrid As SAPbouiCOM.Grid
    Dim oDt As SAPbouiCOM.DataTable
    Dim oEditText As SAPbouiCOM.EditText
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim sFlag As Boolean = False
    Dim objCheckBox As SAPbouiCOM.CheckBox
    Public oDS, oDS1, oDS2, oDS3 As SAPbouiCOM.DBDataSource
    Dim BaseProjectNum, BaseObjectNum As String
#End Region

    Sub CreateForm(ByVal BaseDoc As String, ByVal BaseCardCode As String)
        Try
            objMain.objUtilities.LoadForm("ItemListPR.xml", "ITEMLIST", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ITEMLIST", objMain.objApplication.Forms.ActiveForm.TypeCount)


            objForm.DataSources.UserDataSources.Add("U_BaseN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_TxtFD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_CardC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_ItmType", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            objComboBox = objForm.Items.Item("ItmType").Specific
            ' objComboBox.Select("Item Description", SAPbouiCOM.BoSearchKey.psk_ByValue)
            objForm.DataSources.UserDataSources.Add("Select", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objCheckBox = objForm.Items.Item("Select").Specific
            objCheckBox.DataBind.SetBound(True, "", "Select")
            objForm.Items.Item("BaseN").Specific.Value = BaseDoc
            objForm.Items.Item("CardC").Specific.Value = BaseCardCode
            oDt = objForm.DataSources.DataTables.Add("dt1")
            oDt = objForm.DataSources.DataTables.Item("dt1")
            BaseProjectNum = BaseCardCode
            Me.LoadGrid(objForm.UniqueID, "", BaseCardCode)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            'If pVal.MenuUID = "ME_GPOFPRQ" And pVal.BeforeAction = False Then
            '    ' Me.CreateForm()
            'End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "Find" And pVal.BeforeAction = False Then
                        Dim Itemcode As String = objForm.Items.Item("TxtFD").Specific.Value
                        objComboBox = objForm.Items.Item("ItmType").Specific
                        Dim ItemType As String = objForm.Items.Item("ItmType").Specific.Value

                        If Itemcode <> "" Then
                            Me.LoadGrid1(objForm.UniqueID, "", BaseProjectNum, Itemcode, ItemType)
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)


                    If pVal.ItemUID = "Select" And pVal.BeforeAction = False Then
                        objCheckBox = objForm.Items.Item("Select").Specific
                        objForm.Freeze(True)
                        Dim Itemcode As String = objForm.Items.Item("TxtFD").Specific.Value
                        Dim ItemType As String = objForm.Items.Item("ItmType").Specific.Value
                        If objForm.Items.Item("Select").Specific.checked = True And Itemcode = "" Then
                            Me.LoadGrid(objForm.UniqueID, "Y", BaseProjectNum)
                        ElseIf objForm.Items.Item("Select").Specific.checked = True And Itemcode <> "" Then
                            Me.LoadGrid1(objForm.UniqueID, "Y", BaseProjectNum, Itemcode, ItemType)
                        ElseIf objForm.Items.Item("Select").Specific.checked = False And Itemcode = "" Then
                            Me.LoadGrid(objForm.UniqueID, "N", BaseProjectNum)
                        ElseIf objForm.Items.Item("Select").Specific.checked = False And Itemcode <> "" Then
                            Me.LoadGrid1(objForm.UniqueID, "N", BaseProjectNum, Itemcode, ItemType)
                        End If

                        objForm.Freeze(False)
                    End If


                    If pVal.ItemUID = "1" And pVal.BeforeAction = False Then

                        objBaseForm = objMain.objApplication.Forms.GetForm("1470000200", objMain.objApplication.Forms.ActiveForm.TypeCount)
                        'If objBaseForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        oDS = Me.objBaseForm.DataSources.DBDataSources.Add("PRQ1")
                            Dim objMatrix1 As SAPbouiCOM.Matrix = objBaseForm.Items.Item("38").Specific
                            objMatrix1.Clear()
                            oGrid = objForm.Items.Item("Item_2").Specific

                            objMatrix1.AddRow()
                            For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                                If oGrid.DataTable.GetValue("Check", i) = "Y" Then
                                    Dim ItemC As String = oGrid.DataTable.GetValue("Item Code", i)
                                    Dim Quant As String = oGrid.DataTable.GetValue("BarCode", i)
                                objMatrix1.Columns.Item("1").Cells.Item(objMatrix1.VisualRowCount).Specific.Value() = oGrid.DataTable.GetValue("Item Code", i)
                                objMatrix1.Columns.Item("1470002145").Cells.Item(objMatrix1.VisualRowCount - 1).Specific.Value() = oGrid.DataTable.GetValue("UOM", i)
                                objMatrix1.Columns.Item("4").Cells.Item(objMatrix1.VisualRowCount - 1).Specific.Value() = Quant

                                End If
                            Next
                            objForm.Close()
                            Exit Select
                        ' End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    Select Case pVal.ItemUID.ToUpper
                        Case "Item_2"
                            objForm.Freeze(True)
                            objForm.PaneLevel = 1
                            ' FloderName = "Earnings"
                            objForm.Settings.Grid = "Item_2"
                            objForm.Settings.Enabled = True
                            objForm.Freeze(False)

                    End Select
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub LoadGrid(ByVal FormUID As String, ByVal CheckValue As String, ByVal BaseCardCode As String)
        Try
            objForm.Freeze(True)
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("Item_2").Specific
            oGrid.DataTable = oDt
            oGrid.DataTable.Clear()
            'Dim Details As String = "SELECT '" & CheckValue & "' as""Check"", T0.""ItemCode""as ""Item Code"",T0.""ItemName""as ""Item Description"",T0.""CodeBars""as ""BarCode"" " &
            '",T2.""ItmsGrpNam""as ""Item Group"",""U_VSubGroup""as ""Sub Group"",""U_VSubSubGroup""as ""Sub Sub Group"",F.""FirmName"" as""Manufacturer""from OITM T0 INNER JOIN ITM2 T1 ON T0.""ItemCode"" = T1.""ItemCode"" " &
            '"INNER JOIN OITB T2 ON T2.""ItmsGrpCod"" = T0.""ItmsGrpCod""  Left JOIN OMRC F ON F.""FirmCode"" = T0.""FirmCode"" where T1.""VendorCode"" ='" & BaseCardCode & "'  order by T0.""ItemCode"" "
            '
            'Dim Details As String = "SELECT  distinct '" & CheckValue & "' as""Check"", T0.""ItemCode""as ""Item Code"",T0.""ItemName""as ""Item Description"",T0.""CodeBars""as ""BarCode"" " &
            '",T2.""ItmsGrpNam""as ""Item Group"",""U_VSubGroup""as ""Sub Group"",""U_VSubSubGroup""as ""Sub Sub Group"",F.""FirmName"" as""Manufacturer"",""InvntryUom""as ""UOM""from OITM T0 INNER JOIN ITM2 T1 On T0.""ItemCode"" = T1.""ItemCode"" " &
            '"INNER JOIN OITB T2 On T2.""ItmsGrpCod"" = T0.""ItmsGrpCod"" INNER JOIN OBCD T3 On T3.""ItemCode"" =T0.""ItemCode""INNER JOIN OUOM T4 On T4.""UomEntry"" =T3.""UomEntry""  Left JOIN OMRC F On F.""FirmCode"" = T0.""FirmCode"" where T1.""VendorCode"" ='" & BaseCardCode & "'  order by T0.""ItemCode"" "

            Dim Details As String = "SELECT distinct  '" & CheckValue & "' as""Check"", T0.""ItemCode""as ""Item Code"",T0.""ItemName""as ""Item Description"",(case when ""CodeBars"" in " &
            "(select  (""BcdCode"")from OBCD where  ""ItemCode"" =T0.""ItemCode"" And ""UomEntry""=T4.""UomEntry"") Then ""CodeBars""Else(select Max (""BcdCode"")from OBCD where  ""ItemCode"" =T0.""ItemCode"" And ""UomEntry""=T4.""UomEntry"") End) as ""BarCode"" " &
            ",T2.""ItmsGrpNam""as ""Item Group"",""U_VSubGroup""as ""Sub Group"",""U_VSubSubGroup""as ""Sub Sub Group"",F.""FirmName"" as""Manufacturer"",""UomCode"" as ""UOM"" " &
            "From OITM T0 INNER Join ITM2 T1 ON T0.""ItemCode"" = T1.""ItemCode"" INNER Join OITB T2 ON T2.""ItmsGrpCod"" = T0.""ItmsGrpCod"" Left Join OUOM T4 ON T0.""PUoMEntry"" = T4.""UomEntry""Left Join OMRC F ON F.""FirmCode"" = T0.""FirmCode"" where T1.""VendorCode"" ='" & BaseCardCode & "' order by T0.""ItemCode"" "

            objForm.DataSources.DataTables.Item(0).ExecuteQuery(Details)
            If oGrid.Rows.Count = 0 Then
                objMain.objApplication.MessageBox("No records found")
                Exit Sub
            End If
            oGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            For i As Integer = 0 To oGrid.Columns.Count - 1
                oGrid.Columns.Item("Item Code").Editable = False
                oGrid.Columns.Item("Item Description").Editable = False
                oGrid.Columns.Item("BarCode").Editable = False
                oGrid.Columns.Item("UOM").Editable = False
                oGrid.Columns.Item("Item Group").Editable = False
                oGrid.Columns.Item("Sub Group").Editable = False
                oGrid.Columns.Item("Sub Sub Group").Editable = False
                oGrid.Columns.Item("Manufacturer").Editable = False
            Next

            Dim oEditTextColumn As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("Item Code")
            oEditTextColumn.LinkedObjectType = 4

            'oEditTextColumn = oGrid.Columns.Item("DraftEntry")
            'oEditTextColumn.LinkedObjectType = 112

            oGrid.AutoResizeColumns()
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("ITEMLIST", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then

                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Private Sub LoadGrid1(ByVal FormUID As String, ByVal CheckValue As String, ByVal BaseCardCode As String, ByVal Itemcode As String, ByVal ItemType As String)
        Try
            Dim Details As String = ""
            If ItemType = "Item Code" Then
                '    Details = "SELECT '" & CheckValue & "' as""Check"", T0.""ItemCode""as ""Item Code"",T0.""ItemName""as ""Item Description"",T0.""CodeBars""as ""BarCode"" " &
                '",T2.""ItmsGrpNam""as ""Item Group"",""U_VSubGroup""as ""Sub Group"",""U_VSubSubGroup""as ""Sub Sub Group"",F.""FirmName"" as""Manufacturer""from OITM T0 INNER JOIN ITM2 T1 ON T0.""ItemCode"" = T1.""ItemCode"" " &
                '"INNER JOIN OITB T2 ON T2.""ItmsGrpCod"" = T0.""ItmsGrpCod""  INNER JOIN OMRC F ON F.""FirmCode"" = T0.""FirmCode"" where T1.""VendorCode"" ='" & BaseCardCode & "'and T0.""ItemCode"" like '%%" & Itemcode & "%%'  order by T0.""ItemCode"" "
                ' Details = "SELECT '" & CheckValue & "' as""Check"", T0.""ItemCode""as ""Item Code"",T0.""ItemName""as ""Item Description"",T0.""CodeBars""as ""BarCode"" from OITM T0 INNER JOIN ITM2 T1 ON T0.""ItemCode"" = T1.""ItemCode"" where T1.""VendorCode"" ='" & BaseCardCode & "' order by T0.""ItemCode"" "
                Details = "SELECT distinct  '" & CheckValue & "' as""Check"", T0.""ItemCode""as ""Item Code"",T0.""ItemName""as ""Item Description"",(case when ""CodeBars"" in " &
            "(select  (""BcdCode"")from OBCD where  ""ItemCode"" =T0.""ItemCode"" And ""UomEntry""=T4.""UomEntry"") Then ""CodeBars""Else(select Max (""BcdCode"")from OBCD where  ""ItemCode"" =T0.""ItemCode"" And ""UomEntry""=T4.""UomEntry"") End) as ""BarCode"" " &
            ",T2.""ItmsGrpNam""as ""Item Group"",""U_VSubGroup""as ""Sub Group"",""U_VSubSubGroup""as ""Sub Sub Group"",F.""FirmName"" as""Manufacturer"",""UomCode"" as ""UOM"" " &
            "From OITM T0 INNER Join ITM2 T1 ON T0.""ItemCode"" = T1.""ItemCode"" INNER Join OITB T2 ON T2.""ItmsGrpCod"" = T0.""ItmsGrpCod"" Left Join OUOM T4 ON T0.""PUoMEntry"" = T4.""UomEntry""Left Join OMRC F ON F.""FirmCode"" = T0.""FirmCode"" " &
                " where T1.""VendorCode"" ='" & BaseCardCode & "'and T0.""ItemCode"" like '%%" & Itemcode & "%%'  order by T0.""ItemCode"" "


            ElseIf ItemType = "Item Description" Then
                'Details = "SELECT '" & CheckValue & "' as""Check"", T0.""ItemCode""as ""Item Code"",T0.""ItemName""as ""Item Description"",T0.""CodeBars""as ""BarCode"" from OITM T0 INNER JOIN ITM2 T1 ON T0.""ItemCode"" = T1.""ItemCode"" where T1.""VendorCode"" ='" & BaseCardCode & "'and T0.""ItemName"" like '%%" & Itemcode & "%%' order by T0.""ItemCode"" "
                '    Details = "SELECT '" & CheckValue & "' as""Check"", T0.""ItemCode""as ""Item Code"",T0.""ItemName""as ""Item Description"",T0.""CodeBars""as ""BarCode"" " &
                '",T2.""ItmsGrpNam""as ""Item Group"",""U_VSubGroup""as ""Sub Group"",""U_VSubSubGroup""as ""Sub Sub Group"",F.""FirmName"" as""Manufacturer""from OITM T0 INNER JOIN ITM2 T1 ON T0.""ItemCode"" = T1.""ItemCode"" " &
                '"INNER JOIN OITB T2 ON T2.""ItmsGrpCod"" = T0.""ItmsGrpCod""  INNER JOIN OMRC F ON F.""FirmCode"" = T0.""FirmCode"" where T1.""VendorCode"" ='" & BaseCardCode & "'and T0.""ItemName"" like '%%" & Itemcode & "%%'  order by T0.""ItemCode"" "
                Details = "SELECT distinct  '" & CheckValue & "' as""Check"", T0.""ItemCode""as ""Item Code"",T0.""ItemName""as ""Item Description"",(case when ""CodeBars"" in " &
            "(select  (""BcdCode"")from OBCD where  ""ItemCode"" =T0.""ItemCode"" And ""UomEntry""=T4.""UomEntry"") Then ""CodeBars""Else(select Max (""BcdCode"")from OBCD where  ""ItemCode"" =T0.""ItemCode"" And ""UomEntry""=T4.""UomEntry"") End) as ""BarCode"" " &
            ",T2.""ItmsGrpNam""as ""Item Group"",""U_VSubGroup""as ""Sub Group"",""U_VSubSubGroup""as ""Sub Sub Group"",F.""FirmName"" as""Manufacturer"",""UomCode"" as ""UOM"" " &
            "From OITM T0 INNER Join ITM2 T1 ON T0.""ItemCode"" = T1.""ItemCode"" INNER Join OITB T2 ON T2.""ItmsGrpCod"" = T0.""ItmsGrpCod"" Left Join OUOM T4 ON T0.""PUoMEntry"" = T4.""UomEntry""Left Join OMRC F ON F.""FirmCode"" = T0.""FirmCode"" " &
                " where T1.""VendorCode"" ='" & BaseCardCode & "'and T0.""ItemName"" like '%%" & Itemcode & "%%'  order by T0.""ItemCode"" "

            Else
                'Details = "SELECT '" & CheckValue & "' as""Check"", T0.""ItemCode""as ""Item Code"",T0.""ItemName""as ""Item Description"",T0.""CodeBars""as ""BarCode"" from OITM T0 INNER JOIN ITM2 T1 ON T0.""ItemCode"" = T1.""ItemCode"" where T1.""VendorCode"" ='" & BaseCardCode & "'and T0.""CodeBars"" like '%%" & Itemcode & "%%' order by T0.""ItemCode"" "
                '    Details = "SELECT '" & CheckValue & "' as""Check"", T0.""ItemCode""as ""Item Code"",T0.""ItemName""as ""Item Description"",T0.""CodeBars""as ""BarCode"" " &
                '",T2.""ItmsGrpNam""as ""Item Group"",""U_VSubGroup""as ""Sub Group"",""U_VSubSubGroup""as ""Sub Sub Group"",F.""FirmName"" as""Manufacturer""from OITM T0 INNER JOIN ITM2 T1 ON T0.""ItemCode"" = T1.""ItemCode"" " &
                '
                '"INNER JOIN OITB T2 ON T2.""ItmsGrpCod"" = T0.""ItmsGrpCod""  INNER JOIN OMRC F ON F.""FirmCode"" = T0.""FirmCode"" where T1.""VendorCode"" ='" & BaseCardCode & "'and T0.""CodeBars"" like '%%" & Itemcode & "%%'  order by T0.""ItemCode"" "
                Details = "SELECT  * from (select distinct  '" & CheckValue & "' as""Check"", T0.""ItemCode""as ""Item Code"",T0.""ItemName""as ""Item Description"",(case when ""CodeBars"" in " &
            "(select  (""BcdCode"")from OBCD where  ""ItemCode"" =T0.""ItemCode"" And ""UomEntry""=T4.""UomEntry"") Then ""CodeBars""Else(select Max (""BcdCode"")from OBCD where  ""ItemCode"" =T0.""ItemCode"" And ""UomEntry""=T4.""UomEntry"") End) as ""BarCode"" " &
            ",T2.""ItmsGrpNam""as ""Item Group"",""U_VSubGroup""as ""Sub Group"",""U_VSubSubGroup""as ""Sub Sub Group"",F.""FirmName"" as""Manufacturer"",""UomCode"" as ""UOM"" " &
            "From OITM T0 INNER Join ITM2 T1 ON T0.""ItemCode"" = T1.""ItemCode"" INNER Join OITB T2 ON T2.""ItmsGrpCod"" = T0.""ItmsGrpCod"" Left Join OUOM T4 ON T0.""PUoMEntry"" = T4.""UomEntry""Left Join OMRC F ON F.""FirmCode"" = T0.""FirmCode"" " &
                " where T1.""VendorCode"" ='" & BaseCardCode & "'order by T0.""ItemCode"")s where s.""BarCode"" like '%%" & Itemcode & "%%'  "

            End If
            objForm.DataSources.DataTables.Item(0).ExecuteQuery(Details)

            If oGrid.Rows.Count = 0 Then
                objMain.objApplication.MessageBox("No records found")
                Exit Sub
            End If

            oGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            For i As Integer = 0 To oGrid.Columns.Count - 1
                oGrid.Columns.Item("Item Code").Editable = False
                oGrid.Columns.Item("Item Description").Editable = False
                oGrid.Columns.Item("BarCode").Editable = False
                oGrid.Columns.Item("UOM").Editable = False
                oGrid.Columns.Item("Item Group").Editable = False
                oGrid.Columns.Item("Sub Group").Editable = False
                oGrid.Columns.Item("Sub Sub Group").Editable = False
                oGrid.Columns.Item("Manufacturer").Editable = False
            Next
            Dim oEditTextColumn As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("Item Code")
            oEditTextColumn.LinkedObjectType = 4

            oGrid.AutoResizeColumns()
            objForm.Freeze(False)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


End Class
