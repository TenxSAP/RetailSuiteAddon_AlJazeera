Imports System.Collections.Specialized
Imports System.Configuration
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
'new 
'Madhu
Public Class clsBarcodePrinting
#Region "       Declaration             "
    Public objForm, objSTRForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details1 As SAPbouiCOM.DBDataSource
    Dim objMatrix1 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Private objMatrix As SAPbouiCOM.Matrix
    Public oDS, oDS1 As SAPbouiCOM.DBDataSource
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("BarcodePrinting.xml", "frmBarCd", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("frmBarCd", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            oDS = Me.objForm.DataSources.DBDataSources.Add("@ALE_BARCODE")
            oDS1 = Me.objForm.DataSources.DBDataSources.Add("@ALE_BARCODE1")
            Me.objForm.DataBrowser.BrowseBy = "txtDocNum"
            Me.objForm.EnableMenu("1293", True)
            'Me.GetNextDocNum()

            Me.objComboBox = DirectCast(Me.objForm.Items.Item("cmbPSize").Specific, SAPbouiCOM.ComboBox)
            Me.objComboBox.ValidValues.Add("TSC Paper", "TSCPaper")
            Me.objComboBox.ValidValues.Add("A4", "A4")
            Me.objComboBox.[Select]("TSC Paper", SAPbouiCOM.BoSearchKey.psk_ByValue)
            Dim comboBox As SAPbouiCOM.ComboBox = DirectCast(Me.objForm.Items.Item("Type").Specific, SAPbouiCOM.ComboBox)
            comboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            'comboBox.[Select]("1", SAPbouiCOM.BoSearchKey.psk_ByValue)
            Me.objMatrix = DirectCast(Me.objForm.Items.Item("matrix").Specific, SAPbouiCOM.Matrix)

            '-----------------------------------------
            Me.objForm.Items.Item("txtBP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Me.objForm.Items.Item("txtBP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            Me.objForm.Items.Item("txtBPName").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Me.objForm.Items.Item("txtBPName").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            Me.objForm.Items.Item("txtDocType").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Me.objForm.Items.Item("txtDocType").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            Me.objForm.Items.Item("txtDocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Me.objForm.Items.Item("txtDocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            Me.objForm.Items.Item("txtDocDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Me.objForm.Items.Item("txtDocDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            '---------------------
            Me.EnableColumns()
            If (Me.objMatrix.RowCount > 0) Then
                Me.objMatrix.Clear()
            End If
            Me.objForm.Items.Item("txthidden").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Me.DisableColumns()
            Me.objForm.Freeze(False)
            Me.objForm.Update()
            '-------------------------------------------------------------------
            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery("SELECT Max(T0.""DocNum"")+1 FROM ""@ALE_BARCODE""  T0")
            Me.objForm.Items.Item("txtDocNum").Specific.Value = oRs.Fields.Item(0).Value
            Me.objForm.Items.Item("txtDocDate").Specific.Value = Date.Now.ToString("yyyyMMdd")



            '-----------Loading UoM --Added by Naresh---------------------------
            Dim businessObject As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            Dim str As String = ""
            Me.objMatrix = DirectCast(Me.objForm.Items.Item("matrix").Specific, SAPbouiCOM.Matrix)
            objMatrix.AddRow()
            Dim UOMCombo As SAPbouiCOM.ComboBox = DirectCast(Me.objMatrix.Columns.Item("Col_0").Cells.Item(1).Specific, SAPbouiCOM.ComboBox)
            Dim UOMQry As String = "SELECT DISTINCT A.""UomEntry"",(SELECT ""UomCode"" FROM OUOM WHERE ""UomEntry""=A.""UomEntry"") AS ""UomName"" FROM OBCD A "
            objMain.objUtilities.ComboBoxLoadValues(UOMCombo, UOMQry)

            Me.EnableColumns()
            objMatrix.Clear()
            Me.SetNewLine()
            Me.objForm.Items.Item("PQty").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Me.DisableControls()
            '----------------------------------------
            Dim oRsBranch As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim Query As String = "SELECT T2.""BPLId"",T2.""BPLName"" AS ""Branch"" FROM USR6 T0  INNER JOIN OUSR T1 ON T0.""UserCode"" = T1.""USER_CODE"" " &
                    " INNER JOIN OBPL T2 ON T0.""BPLId"" = T2.""BPLId"" WHERE T1.""USER_CODE"" = '" & objMain.objCompany.UserName.ToString().Trim & "' "
            oRsBranch.DoQuery(Query)
            'If oRsBranch.RecordCount > 1 Then
            '    objMain.objApplication.StatusBar.SetText("Login user have more then one branch.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Exit Sub
            'End If
            objComboBox = objForm.Items.Item("Branch").Specific
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, Query)
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("Branch").DisplayDesc = True
            'If objComboBox.ValidValues.Count <> 0 Then
            '    objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
            'End If
            '----------------------------------------
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ME_BR" And pVal.BeforeAction = False Then
                Me.CreateForm()
                ' Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
                'Me.GetNextDocNum()

            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    objMatrix = Me.objForm.Items.Item("matrix").Specific
                    If pVal.ItemUID = "matrix" And (pVal.ColUID = "ItemCode" Or pVal.ColUID = "Desc" Or pVal.ColUID = "BarCode") Then
                        If pVal.Row = objMatrix.VisualRowCount Then
                            Me.SetNewLine()
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If (Microsoft.VisualBasic.CompilerServices.Operators.CompareString(pVal.ItemUID, "11", False) = 0 And Me.objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                        If pVal.BeforeAction = False Then
                            Me.PrintBarcode(objForm.UniqueID)
                        End If
                    End If
                    If pVal.ItemUID = "Load" And pVal.BeforeAction = False And Me.objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim PriceUpdate As String = "SELECT T1.""U_ICode"", T1.""U_IName"", T1.""U_BarCode"", (select ""UomEntry"" from OUOM Where ""UomCode""=T1.""U_UOM"") AS ""UOM"", T1.""U_NPrice""  FROM ""@SBO_PL"" T0 " &
                        " INNER JOIN ""@SBO_PL_C0"" T1 ON T0.""DocEntry""=T1.""DocEntry"" Where T0.""DocNum""='" & objForm.Items.Item("PriceUp").Specific.Value.ToString() & "' "
                        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRs.DoQuery(PriceUpdate)
                        oDS = Me.objForm.DataSources.DBDataSources.Add("@ALE_BARCODE")
                        oDS1 = Me.objForm.DataSources.DBDataSources.Add("@ALE_BARCODE1")
                        objMatrix = Me.objForm.Items.Item("matrix").Specific
                        objMatrix.Clear()
                        oDS1.Clear()
                        objMatrix.FlushToDataSource()
                        objMatrix.AutoResizeColumns()
                        For i As Integer = 1 To oRs.RecordCount
                            objMatrix.AddRow()
                            oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
                            oDS1.SetValue("U_ItemCode", oDS1.Offset, oRs.Fields.Item("U_ICode").Value)
                            oDS1.SetValue("U_ItemName", oDS1.Offset, oRs.Fields.Item("U_IName").Value)
                            oDS1.SetValue("U_Uom", oDS1.Offset, oRs.Fields.Item("UOM").Value)
                            oDS1.SetValue("U_CodeBars", oDS1.Offset, oRs.Fields.Item("U_BarCode").Value)
                            oDS1.SetValue("U_FrgnName", oDS1.Offset, "")
                            oDS1.SetValue("U_ItmGrpNm", oDS1.Offset, "")
                            oDS1.SetValue("U_FirmName", oDS1.Offset, "")
                            oDS1.SetValue("U_Qty", oDS1.Offset, "1")
                            oDS1.SetValue("U_Print", oDS1.Offset, "Y")
                            oDS1.SetValue("U_PrintQty", oDS1.Offset, "")
                            objMatrix.SetLineData(objMatrix.VisualRowCount)
                            oRs.MoveNext()
                        Next
                    End If
                    If pVal.ItemUID = "Update" And pVal.BeforeAction = False And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        Dim PrintQty As String = Me.objForm.Items.Item("PQty").Specific.Value.ToString().Trim()
                        For i As Integer = 1 To objMatrix.VisualRowCount
                            objMatrix.Columns.Item("PrintQty").Cells.Item(i).Specific.Value = PrintQty
                        Next
                    End If

                    If pVal.ItemUID = "LoadDet" And pVal.BeforeAction = False And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        For i As Integer = 1 To objMatrix.VisualRowCount
                            Dim ItemCode As String = objMatrix.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim()
                            Dim CodeBars As String = objMatrix.Columns.Item("BarCode").Cells.Item(i).Specific.Value.ToString().Trim()
                            Dim UOM As String = objMatrix.Columns.Item("Col_0").Cells.Item(i).Specific.Value.ToString().Trim()
                            oDS = Me.objForm.DataSources.DBDataSources.Add("@ALE_BARCODE")
                            oDS1 = Me.objForm.DataSources.DBDataSources.Add("@ALE_BARCODE1")
                            If ItemCode <> "" Then
                                ' Dim DetQuery As String = "select (select ""UomEntry"" from OUOM Where ""UomCode""=T0.""InvntryUom"") AS ""UOM"" ,T0.""CodeBars"",T0.""ItemName"",T0.""FrgnName"", (SELECT T1.""ItmsGrpNam"" FROM OITB T1 WHERE T1.""ItmsGrpCod"" =T0.""ItmsGrpCod"") AS ""ItemGroup"",(SELECT T2.""FirmName"" FROM OMRC T2 WHERE T2.""FirmCode""= T0.""FirmCode"") AS ""Manufacturer"" from OITM T0 where T0.""ItemCode""='" & ItemCode & "' "
                                'working code
                                'Dim DetQuery As String = "select T0.""IUoMEntry"" AS ""UOM"" ,T0.""CodeBars"",T0.""ItemName"",T0.""U_ShortName"", (SELECT T1.""ItmsGrpNam"" FROM OITB T1 WHERE T1.""ItmsGrpCod"" =T0.""ItmsGrpCod"") AS ""ItemGroup"",(SELECT T2.""FirmName"" FROM OMRC T2 WHERE T2.""FirmCode""= T0.""FirmCode"") AS ""Manufacturer"" from OITM T0 where T0.""ItemCode""='" & ItemCode & "' "


                                Dim DetQuery As String
                                If String.IsNullOrEmpty(UOM) Then
                                    DetQuery = "SELECT 
        T0.""IUoMEntry"" AS ""UOM"", 
        T0.""CodeBars"", 
        T0.""ItemName"",
        T0.""U_ShortName"",
        (SELECT T1.""ItmsGrpNam"" FROM OITB T1 WHERE T1.""ItmsGrpCod"" = T0.""ItmsGrpCod"") AS ""ItemGroup"",
        (SELECT T2.""FirmName"" FROM OMRC T2 WHERE T2.""FirmCode"" = T0.""FirmCode"") AS ""Manufacturer"",
        B1.""BcdCode""
    FROM 
        OITM T0
        INNER JOIN OBCD B1 ON B1.""ItemCode"" = T0.""ItemCode"" AND B1.""UomEntry"" = T0.""IUoMEntry""
    WHERE 
        T0.""ItemCode"" = '" & ItemCode & "' "
                                Else
                                    ' If UOM is selected, include the condition
                                    DetQuery = "SELECT 
        T0.""IUoMEntry"" AS ""UOM"", 
        T0.""CodeBars"", 
        T0.""ItemName"",
        T0.""U_ShortName"",
        (SELECT T1.""ItmsGrpNam"" FROM OITB T1 WHERE T1.""ItmsGrpCod"" = T0.""ItmsGrpCod"") AS ""ItemGroup"",
        (SELECT T2.""FirmName"" FROM OMRC T2 WHERE T2.""FirmCode"" = T0.""FirmCode"") AS ""Manufacturer"",
        B1.""BcdCode""
    FROM 
        OITM T0
        INNER JOIN OBCD B1 ON B1.""ItemCode"" = T0.""ItemCode"" AND B1.""UomEntry"" = T0.""IUoMEntry""
    WHERE 
        T0.""ItemCode"" = '" & ItemCode & "' 
        AND T0.""IUoMEntry"" = '" & UOM & "'"
                                End If


                                Dim oRsDetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsDetQuery.DoQuery(DetQuery)
                                oDS1.SetValue("LineId", oDS1.Offset, i)
                                oDS1.SetValue("U_ItemCode", oDS1.Offset, ItemCode)
                                oDS1.SetValue("U_ItemName", oDS1.Offset, oRsDetQuery.Fields.Item("ItemName").Value)
                                oDS1.SetValue("U_Uom", oDS1.Offset, oRsDetQuery.Fields.Item("UOM").Value)
                                'oDS1.SetValue("U_CodeBars", oDS1.Offset, oRsDetQuery.Fields.Item("CodeBars").Value)
                                oDS1.SetValue("U_CodeBars", oDS1.Offset, oRsDetQuery.Fields.Item("BcdCode").Value)

                                oDS1.SetValue("U_FrgnName", oDS1.Offset, oRsDetQuery.Fields.Item("U_ShortName").Value)
                                oDS1.SetValue("U_ItmGrpNm", oDS1.Offset, oRsDetQuery.Fields.Item("ItemGroup").Value)
                                oDS1.SetValue("U_FirmName", oDS1.Offset, oRsDetQuery.Fields.Item("Manufacturer").Value)
                                oDS1.SetValue("U_Qty", oDS1.Offset, "1")
                                oDS1.SetValue("U_Print", oDS1.Offset, "Y")
                                oDS1.SetValue("U_PrintQty", oDS1.Offset, "1")
                                objMatrix.SetLineData(i)
                                'oRsDetQuery.MoveNext()

                            ElseIf CodeBars <> "" Then
                                ' Dim DetQuery As String = "select (select ""UomEntry"" from OUOM Where ""UomCode""=T0.""InvntryUom"") AS ""UOM"" ,T0.""ItemCode"",T0.""ItemName"",T0.""FrgnName"", (SELECT T1.""ItmsGrpNam"" FROM OITB T1 WHERE T1.""ItmsGrpCod"" =T0.""ItmsGrpCod"") AS ""ItemGroup"",(SELECT T2.""FirmName"" FROM OMRC T2 WHERE T2.""FirmCode""= T0.""FirmCode"") AS ""Manufacturer"" from OITM T0 where T0.""CodeBars""='" & CodeBars & "' "
                                Dim DetQuery As String = "select T3.""UomEntry"" AS ""UOM"" ,T0.""ItemCode"",T0.""ItemName"",T0.""U_ShortName"", (SELECT T1.""ItmsGrpNam"" FROM OITB T1 WHERE T1.""ItmsGrpCod"" =T0.""ItmsGrpCod"") AS ""ItemGroup"",(SELECT T2.""FirmName"" FROM OMRC T2 WHERE T2.""FirmCode""= T0.""FirmCode"") AS ""Manufacturer""  from OBCD T3 INNER JOIN OITM T0 ON T3.""ItemCode""=T0.""ItemCode"" where T3.""BcdCode""='" & CodeBars & "' "
                                Dim oRsDetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsDetQuery.DoQuery(DetQuery)
                                oDS1.SetValue("LineId", oDS1.Offset, i)
                                oDS1.SetValue("U_ItemCode", oDS1.Offset, oRsDetQuery.Fields.Item("ItemCode").Value)
                                oDS1.SetValue("U_ItemName", oDS1.Offset, oRsDetQuery.Fields.Item("ItemName").Value)
                                oDS1.SetValue("U_Uom", oDS1.Offset, oRsDetQuery.Fields.Item("UOM").Value)
                                oDS1.SetValue("U_CodeBars", oDS1.Offset, CodeBars)
                                oDS1.SetValue("U_FrgnName", oDS1.Offset, oRsDetQuery.Fields.Item("U_ShortName").Value)
                                oDS1.SetValue("U_ItmGrpNm", oDS1.Offset, oRsDetQuery.Fields.Item("ItemGroup").Value)
                                oDS1.SetValue("U_FirmName", oDS1.Offset, oRsDetQuery.Fields.Item("Manufacturer").Value)
                                oDS1.SetValue("U_Qty", oDS1.Offset, "1")
                                oDS1.SetValue("U_Print", oDS1.Offset, "Y")
                                oDS1.SetValue("U_PrintQty", oDS1.Offset, "1")
                                objMatrix.SetLineData(i)
                                ' oRsDetQuery.MoveNext()
                            End If
                        Next

                    End If


                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And
(pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        Me.SetDefault(objForm.UniqueID)
                    End If



                    If pVal.ItemUID = "btnOk" And pVal.BeforeAction = False Then
                        'If DateFrom = "" Then
                        '    objForm.StatusBar.SetText("Please enter From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        'ElseIf DateTo = "" Then
                        '    objForm.StatusBar.SetText("Please enter To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        'ElseIf Branch = "" Then
                        '    objForm.StatusBar.SetText("Please enter Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        'Else
                        Me.PrintBarcode(objForm.UniqueID)
                        'End If
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
                    If oCFL.UniqueID = "CA" And pVal.BeforeAction = True Then
                        'Me.CFLFilter("CA")
                    End If
                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        'If oCFL.UniqueID = "CA" Then
                        '    oDBs_Details1.SetValue("U_DepositA", oDBs_Details1.Offset, oDT.GetValue("AcctCode", 0))
                        'End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    If pVal.ItemUID = "Item_4" And pVal.BeforeAction = False Then
                        Dim objUOMCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Item_4").Specific
                        Dim Qry As String = "SELECT ""BcdEntry"",""BcdCode"",IFNULL((SELECT DISTINCT ""BcdEntDft"" FROM ITM12 WHERE ""ItemCode""='" & objForm.Items.Item("txtItem").Specific.Value & "' " &
                            " AND ""UomEntry""='" & objUOMCombo.Selected.Value & "' AND IFNULL(""BcdEntDft"",0)<>0),0) AS ""DefBarCode"" FROM OBCD WHERE ""ItemCode""='" & objForm.Items.Item("txtItem").Specific.Value & "' " &
                            " AND ""UomEntry""='" & objUOMCombo.Selected.Value & "'"
                        Dim objBarcodeCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Item_5").Specific
                        objMain.objUtilities.ComboBoxLoadValues(objBarcodeCombo, Qry)

                        objBarcodeCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                        objForm.Items.Item("Item_5").DisplayDesc = True
                        If objBarcodeCombo.ValidValues.Count <> 0 Then
                            objBarcodeCombo.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
                        End If
                    End If
                    If (pVal.ItemUID = "Type" Or pVal.ItemUID = "Branch") And pVal.BeforeAction = False Then
                        Try
                            Dim objTypeCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Type").Specific
                            Dim objBranchCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Branch").Specific
                            Dim Type As String = objTypeCombo.Selected.Value
                            Dim Branch As String = objBranchCombo.Selected.Value
                            If Type <> "" And Branch <> "" Then
                                Dim Qry As String = "SELECT T0.""U_PrinterName"",T0.""U_PrinterName"" FROM ""@NT_BARPRSETUP"" T0 " &
                                    " WHERE T0.""U_BPLId"" ='" & Branch & "' and  T0.""U_Type"" ='" & Type & "'"
                                Dim objPrinterCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Printer").Specific
                                objMain.objUtilities.ComboBoxLoadValues(objPrinterCombo, Qry)

                                objPrinterCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                                objForm.Items.Item("Printer").DisplayDesc = True
                                If objPrinterCombo.ValidValues.Count <> 0 Then
                                    objPrinterCombo.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
                                End If
                            End If
                        Catch ex As Exception
                        End Try
                    End If
                    'If (Microsoft.VisualBasic.CompilerServices.Operators.CompareString(pVal.ItemUID, "matrix", False) = 0 And Microsoft.VisualBasic.CompilerServices.Operators.CompareString(pVal.ColUID, "Col_0", False) = 0 And pVal.Row <> 0) Then
                    '    Me.objForm.Freeze(True)
                    '    Dim businessObject As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                    '    Dim str As String = ""
                    '    Me.objMatrix = DirectCast(Me.objForm.Items.Item("matrix").Specific, SAPbouiCOM.Matrix)
                    '    Dim specific As SAPbouiCOM.ComboBox = DirectCast(Me.objMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific, SAPbouiCOM.ComboBox)
                    '    'specific.ExpandType = BoExpandType.et_DescriptionOnly
                    '    str = Conversions.ToString(Microsoft.VisualBasic.CompilerServices.Operators.ConcatenateObject(Microsoft.VisualBasic.CompilerServices.Operators.ConcatenateObject("SELECT DISTINCT A.""UomEntry"",(SELECT ""UomCode"" FROM OUOM WHERE ""UomEntry""=A.""UomEntry"") AS ""UomName"",  (SELECT ""IUoMEntry"" FROM OITM WHERE ""ItemCode""=A.""ItemCode"") AS ""DefUoM"" FROM OBCD A  WHERE A.""ItemCode""='", NewLateBinding.LateGet(Me.objMatrix.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific, Nothing, "value", New Object(-1) {}, Nothing, Nothing, Nothing)), "'"))
                    '    businessObject.DoQuery(str)
                    '    If (specific.ValidValues.Count > 0) Then
                    '        Dim count As Integer = specific.ValidValues.Count
                    '        For i As Integer = 1 To count Step 1
                    '            specific.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
                    '        Next

                    '    End If
                    '    If (businessObject.RecordCount > 0) Then
                    '        While Not businessObject.EoF
                    '            specific.ValidValues.Add(Conversions.ToString(businessObject.Fields.Item("UomEntry").Value), Conversions.ToString(businessObject.Fields.Item("UomName").Value))
                    '            businessObject.MoveNext()
                    '        End While
                    '        Try
                    '            specific.[Select](0, SAPbouiCOM.BoSearchKey.psk_Index)
                    '        Catch exception As System.Exception
                    '            ProjectData.SetProjectError(exception)
                    '            ProjectData.ClearProjectError()
                    '        End Try
                    '    End If
                    'End If
                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    If (Microsoft.VisualBasic.CompilerServices.Operators.CompareString(pVal.ItemUID, "matrix", False) = 0 And Microsoft.VisualBasic.CompilerServices.Operators.CompareString(pVal.ColUID, "PrintYorN", False) = 0 And CDbl(pVal.Row) = Conversions.ToDouble("0")) Then
                        objMatrix = Me.objForm.Items.Item("matrix").Specific
                        Dim visualRowCount As Integer = Me.objMatrix.VisualRowCount
                        Dim checkBox As SAPbouiCOM.CheckBox = DirectCast(Me.objMatrix.Columns.Item("PrintYorN").Cells.Item(1).Specific, SAPbouiCOM.CheckBox)
                        Dim Check As Boolean
                        If (Not checkBox.Checked) Then
                            Check = True
                        Else
                            Check = False
                        End If
                        For j As Integer = 1 To visualRowCount Step 1
                            checkBox = DirectCast(Me.objMatrix.Columns.Item("PrintYorN").Cells.Item(j).Specific, SAPbouiCOM.CheckBox)
                            checkBox.Checked = Check
                        Next
                    End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Public Function GetNextDocNum()
        'Dim Str As String
        'Dim oRs As SAPbobsCOM.Recordset
        'Dim DocNum As Integer
        'oRs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'Str = "SELECT IFNULL(MAX(""U_DocNum""),0)+1 as ""DocNum"" FROM ""@ALE_BARCODE"""
        Try
            '    oRs.DoQuery(Str)
            '    oRs.MoveFirst()
            '    If oRs.RecordCount > 0 Then
            '        DocNum = oRs.Fields.Item(0).Value
            '        oDS.SetValue("U_DocNum", 0, oRs.Fields.Item(0).Value)
            '        oDS.SetValue("U_DocDate", oDS.Offset, DateTime.Now.ToString("yyyyMMdd"))

            '    End If
            '    Me.SetNewLine()
            '    If DocNum = 0 Then DocNum = 1
            'Return DocNum
            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery("SELECT Max(T0.""DocNum"")+1 FROM ""@ALE_BARCODE""  T0")
            Me.objForm.Items.Item("txtDocNum").Specific.Value = oRs.Fields.Item(0).Value
            Me.objForm.Items.Item("txtDocDate").Specific.Value = Date.Now.ToString("yyyyMMdd")
            objMatrix.Clear()
            oDS1.Clear()
            objMatrix.FlushToDataSource()
            Me.SetNewLine()

        Catch ex As Exception
            Throw ex
        End Try
    End Function





    Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "", Optional ByVal Series As Integer = 0)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            objComboBox = objForm.Items.Item("Branch").Specific
            Dim GetSeriesQty As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'ME_OPDC' And ""Locked"" = 'N'"
            objMain.objUtilities.ComboBoxLoadValuesSeries(objForm.Items.Item("Branch").Specific, GetSeriesQty)
            objForm.Items.Item("Branch").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            oDS = Me.objForm.DataSources.DBDataSources.Add("@ALE_BARCODE")
            oDS1 = Me.objForm.DataSources.DBDataSources.Add("@ALE_BARCODE1")
            objMatrix = Me.objForm.Items.Item("matrix").Specific

            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery("SELECT Max(T0.""DocNum"")+1 FROM ""@ALE_BARCODE""  T0")
            Me.objForm.Items.Item("txtDocNum").Specific.Value = oRs.Fields.Item(0).Value
            Me.objForm.Items.Item("txtDocDate").Specific.Value = Date.Now.ToString("yyyyMMdd")



            ' oDS.SetValue("U_DocNum", oDS.Offset, objMain.objUtilities.GetNextDocNum(objForm, "UDO_BARCODE", "Primary"))
            'objForm.DataBrowser.BrowseBy = "txtDocNum"

            'oDS.SetValue("U_DocDate", oDS.Offset, DateTime.Now.ToString("yyyyMMdd"))

            objMatrix.Clear()
            oDS1.Clear()
            objMatrix.FlushToDataSource()
            Me.SetNewLine()

            objMatrix.AutoResizeColumns()

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CFLFilter(ByVal CFL_ID As String)
        Try
            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objForm.ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            oCondition = oConditions.Add()
            oCondition.Alias = "Postable"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "Y"
            oChooseFromList.SetConditions(oConditions)
        Catch ex As Exception

        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("frmBarCodeIM", objMain.objApplication.Forms.ActiveForm.TypeCount)
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

    'Sub PrintBarcode(ByVal FormUID As String)
    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)
    '        objForm.Freeze(True)

    '        Dim directoryName As String = System.Windows.Forms.Application.StartupPath
    '        For Each deleteFile In Directory.GetFiles(directoryName, "*.prn", SearchOption.TopDirectoryOnly)
    '            File.Delete(deleteFile)
    '        Next
    '        Dim GetDataQry As String = "SELECT T0.""U_Printer"", T0.""U_Branch"",(Select ""BPLName"" From OBPL Where ""BPLId""=T0.""U_Branch"") As ""BranchName"" " &
    '                                " , T1.""U_ItemCode"", T1.""U_Uom"", T1.""U_CodeBars"" " &
    '                                " , T1.""U_Qty"",T1.""U_Print"", T1.""U_PrintQty"",""U_WithPrc"",""U_Type"" " &
    '                                " FROM ""@ALE_BARCODE""  T0 INNER JOIN ""@ALE_BARCODE1""  T1 ON T0.""DocEntry""=T1.""DocEntry"" " &
    '                                " Where T0.""DocNum""='" & oDS.GetValue("DocNum", 0) & "' and IFNULL(T1.""U_ItemCode"",'')<>'' and T1.""U_Print""='Y' and IFNULL(""U_CodeBars"",'')<>''"
    '        Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        rs.DoQuery(GetDataQry)
    '        If rs.RecordCount > 0 Then
    '            Dim BranchName As String = rs.Fields.Item("BranchName").Value
    '            Dim BranchID As String = rs.Fields.Item("U_Branch").Value
    '            Dim Printer As String = rs.Fields.Item("U_Printer").Value
    '            'Dim WithPrice As String = rs.Fields.Item("U_WithPrc").Value
    '            Dim WithPriceCheck As SAPbouiCOM.CheckBox = Me.objForm.Items.Item("Item_2").Specific
    '            'Dim WithPrice As String = If(Not WithPriceCheck.Checked, "Y", "N")
    '            ' Dim WithPrice As String = If(Not WithPriceCheck.Checked, "Y", "N")
    '            '--------------------
    '            Dim WithPrice As String
    '            If WithPriceCheck.Checked = True Then
    '                WithPrice = "Y"
    '            Else
    '                WithPrice = "N"
    '            End If

    '            Dim Type As String = rs.Fields.Item("U_Type").Value
    '            For i As Integer = 1 To rs.RecordCount
    '                'Dim UOMCombo As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("Col_0").Cells.Item(i).Specific

    '                Dim PrintQty As Integer = rs.Fields.Item("U_PrintQty").Value
    '                Dim ItemCode As String = rs.Fields.Item("U_ItemCode").Value
    '                Dim Barcode As String = rs.Fields.Item("U_CodeBars").Value
    '                Dim Batch As String = ""
    '                Dim UOM As String = rs.Fields.Item("U_Uom").Value

    '                Dim ItemName As String = Me.GetFieldValue("Call ""Barcode_ItemName""('" & ItemCode & "','" & Barcode & "')")
    '                Dim Price As Double = Me.GetFieldValue("Call ""Barcode_Price""('" & ItemCode & "','" & UOM & "','" & BranchID & "')")

    '                If Type = "" Then
    '                    objMain.objApplication.StatusBar.SetText("Type is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                    Return
    '                ElseIf (Operators.CompareString(Barcode.Trim(), "", False) = 0) Then
    '                    objMain.objApplication.StatusBar.SetText("BarCode is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                    Return
    '                ElseIf PrintQty > 0 Then
    '                    If PrintQty > 0 Then
    '                        ' For P As Integer = 1 To PrintQty
    '                        'Me.Label_Printing(Type, Barcode, ItemName, PrintQty, Price, WithPrice, Printer, Batch, BranchName, ItemCode)
    '                        objMain.objUtilities.Label_Printing(Type, Barcode, ItemName, PrintQty, Price, WithPrice, Printer, Batch, BranchName, ItemCode, UOM)
    '                        ' Next
    '                    End If
    '                Else
    '                    objMain.objApplication.StatusBar.SetText("Quantity is greater than 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                    Return
    '                End If
    '                rs.MoveNext()
    '            Next
    '        Else
    '            objMain.objApplication.StatusBar.SetText("No Records selected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        End If

    '        objForm.Freeze(False)
    '    Catch ex As Exception
    '        objForm.Freeze(False)
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub

    Sub PrintBarcode(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            '=== Delete existing .prn files ===
            Dim directoryName As String = System.Windows.Forms.Application.StartupPath
            For Each deleteFile In Directory.GetFiles(directoryName, "*.prn", SearchOption.TopDirectoryOnly)
                File.Delete(deleteFile)
            Next

            '=== Get barcode printing data ===
            Dim GetDataQry As String = "SELECT T0.""U_Printer"", T0.""U_Branch"", (SELECT ""BPLName"" FROM OBPL WHERE ""BPLId"" = T0.""U_Branch"") AS ""BranchName"" " &
                           ", T1.""U_ItemCode"", T1.""U_Uom"", T1.""U_CodeBars"" " &
                           ", T1.""U_Qty"", T1.""U_Print"", T1.""U_PrintQty"", ""U_WithPrc"", ""U_Type"" " &
                           "FROM ""@ALE_BARCODE"" T0 " &
                           "INNER JOIN ""@ALE_BARCODE1"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
                           "WHERE T0.""DocNum"" = '" & oDS.GetValue("DocNum", 0) & "' " &
                           "AND IFNULL(T1.""U_ItemCode"", '') <> '' " &
                           "AND T1.""U_Print"" = 'Y' " &
                           "AND IFNULL(""U_CodeBars"", '') <> ''"

            Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rs.DoQuery(GetDataQry)

            If rs.RecordCount > 0 Then
                Dim BranchName As String = rs.Fields.Item("BranchName").Value
                Dim BranchID As String = rs.Fields.Item("U_Branch").Value
                Dim Printer As String = rs.Fields.Item("U_Printer").Value

                '=== Checkbox values ===
                Dim WithPriceCheck As SAPbouiCOM.CheckBox = Me.objForm.Items.Item("Item_2").Specific
                Dim WithPromoCheck As SAPbouiCOM.CheckBox = Me.objForm.Items.Item("WIPRC").Specific

                Dim WithPriceOnly As Boolean = WithPriceCheck.Checked
                Dim WithPromoOnly As Boolean = WithPromoCheck.Checked
                Dim WithPrice As String = If(WithPriceOnly Or WithPromoOnly, "Y", "N")

                Dim Type As String = rs.Fields.Item("U_Type").Value

                For i As Integer = 1 To rs.RecordCount
                    Dim PrintQty As Integer = rs.Fields.Item("U_PrintQty").Value
                    Dim ItemCode As String = rs.Fields.Item("U_ItemCode").Value
                    Dim Barcode As String = rs.Fields.Item("U_CodeBars").Value
                    Dim UOM As String = rs.Fields.Item("U_Uom").Value
                    Dim Batch As String = ""

                    '=== Get Item Name ===
                    Dim ItemName As String = Me.GetFieldValue("CALL ""Barcode_ItemName""('" & ItemCode & "','" & Barcode & "')")
                    Dim Price As Double
                    ' If "With Price" checked → Regular Price
                    If WithPriceOnly = True Then
                        Price = Me.GetFieldValue("CALL ""Barcode_Price""('" & ItemCode & "','" & UOM & "','" & BranchID & "')")
                    End If
                    'If "With Promotion Price" checked
                    If WithPromoOnly = True Then
                        Dim PromoPrice As Double = Me.GetFieldValue("CALL ""offerprice""('" & ItemCode & "','" & BranchID & "')")
                        If PromoPrice > 0 Then
                            Price = PromoPrice
                        End If
                    End If
                    If Type = "" Then
                        objMain.objApplication.StatusBar.SetText("Type is missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return
                    ElseIf (Barcode.Trim() = "") Then
                        objMain.objApplication.StatusBar.SetText("Barcode is missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return
                    ElseIf PrintQty <= 0 Then
                        objMain.objApplication.StatusBar.SetText("Quantity must be greater than 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return
                    End If

                    '=== Print Label ===
                    objMain.objUtilities.Label_Printing(Type, Barcode, ItemName, PrintQty, Price, WithPrice, Printer, Batch, BranchName, ItemCode, UOM)

                    rs.MoveNext()
                Next
            Else
                objMain.objApplication.StatusBar.SetText("No records selected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            objForm.Freeze(False)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub Label_Printing(ByVal Type As String, ByVal BarCode As String, ByVal ItemName As String, ByVal Quantity As Integer, ByVal UnitPrice As Double, ByVal Print_UnitPrice As String, ByVal PrinterName As String, ByVal Batch As String, ByVal BranchName As String, ByVal ItemCode As String)
        Try
            Dim PrinterPath As String = ""
            Dim PRNFile As String = ""
            If (Operators.CompareString(Type, "1", False) = 0) Then
                PrinterPath = Me.GetFieldValue("SELECT ""U_PrinterPath"" FROM ""@NT_BARPRSETUP"" WHERE ""U_Type""='1' and ""U_PrinterName""='" & PrinterName & "'")
                PRNFile = Me.Notepad_Read("Type1.txt")
            ElseIf (Operators.CompareString(Type, "2", False) = 0) Then
                PrinterPath = Me.GetFieldValue("SELECT ""U_PrinterPath"" FROM ""@NT_BARPRSETUP"" WHERE ""U_Type""='2' and ""U_PrinterName""='" & PrinterName & "'")
                PRNFile = Me.Notepad_Read("Type2.txt")
            ElseIf (Operators.CompareString(Type, "3", False) = 0) Then
                PrinterPath = Me.GetFieldValue("SELECT ""U_PrinterPath"" FROM ""@NT_BARPRSETUP"" WHERE ""U_Type""='3' and ""U_PrinterName""='" & PrinterName & "'")
                PRNFile = Me.Notepad_Read("Type3.txt")
            ElseIf (Operators.CompareString(Type, "4", False) = 0) Then
                PrinterPath = Me.GetFieldValue("SELECT ""U_PrinterPath"" FROM ""@NT_BARPRSETUP"" WHERE ""U_Type""='4' and ""U_PrinterName""='" & PrinterName & "'")
                PRNFile = Me.Notepad_Read("Type4.txt")
            End If
            If (Operators.CompareString(PrinterPath, "", False) = 0) Then
                Me.objForm.Freeze(False)
                objMain.objApplication.StatusBar.SetText("Printer Path is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return
            ElseIf (Operators.CompareString(PRNFile, "", False) <> 0) Then

                Dim StrTime As String = Format(DateTime.Now, "HHMMssfff")
                Dim StrDate As String = Format(DateTime.Now, "ddMMyy")
                '------------
                Dim FileName As String = "Print_" & objMain.objCompany.UserName & StrTime & ".prn"
                Dim FinalPRNFile As String = String.Concat(System.Windows.Forms.Application.StartupPath, "\", FileName)
                If (File.Exists(FinalPRNFile)) Then
                    File.WriteAllText(FinalPRNFile, "")
                Else
                    File.Create(FinalPRNFile).Close()
                End If
                '-----------------------------------------------
                'Dim ItemGrpQuery As String = " SELECT T1.""ItmsGrpCod""  FROM ""OITM"" T1 WHERE T1.""ItmsGrpCod"" IN (134,139) AND T1.""ItemCode"" = (SELECT T0.""ItemCode"" FROM OBCD T0 WHERE T0.""BcdCode""='" & BarCode & "')"
                'Dim oRsItemGrp As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oRsItemGrp.DoQuery(ItemGrpQuery)

                'If oRsItemGrp.RecordCount = 0 Then
                '    PRNFile = PRNFile.Replace("[VAT]", "(VAT Incl.)")
                'Else
                '    PRNFile = PRNFile.Replace("[VAT]", " ")
                'End If

                '-----------------------------------------------
                Dim BasePrice As Double = 0
                Try
                    Dim ItemQuery As String = " SELECT T0.""U_PIC"", T0.""U_BsUnit"" FROM OITM T0 Where T0.""ItemCode""='" & ItemCode & "'  and IFNULL(T0.""U_PIC"",0)<>0 and IFNULL(T0.""U_BsUnit"",0)<>0"
                    Dim oRsItem As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsItem.DoQuery(ItemQuery)
                    If oRsItem.RecordCount > 0 Then
                        Dim Pic As Double = oRsItem.Fields.Item("U_PIC").Value
                        Dim BaseUnit As Double = oRsItem.Fields.Item("U_BsUnit").Value
                        If Pic <> 0 And BaseUnit <> 0 Then
                            BasePrice = (Pic * BaseUnit) / UnitPrice
                        End If
                    End If
                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText("Base price calculation error:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End Try
                Dim ArabicName As String = Me.GetFieldValue("SELECT T0.""U_ShortName"" FROM OITM T0 WHERE T0.""ItemCode""='" & ItemCode & "'")
                PRNFile = PRNFile.Replace("[ArabicName]", ArabicName)

                '------------------------------------------------
                Try
                    PRNFile = PRNFile.Replace("[Batch]", Batch)
                Catch ex As Exception
                End Try

                PRNFile = PRNFile.Replace("[BasePrice]", BasePrice)
                PRNFile = PRNFile.Replace("[ItemCode]", ItemCode)

                PRNFile = PRNFile.Replace("[BarCode]", BarCode)
                PRNFile = PRNFile.Replace("[ItemName]", ItemName.ToUpper())
                PRNFile = PRNFile.Replace("[Quantity]", Conversions.ToString(Quantity))
                If (Operators.CompareString(Print_UnitPrice.Trim(), "Y", False) <> 0) Then
                    PRNFile = PRNFile.Replace("[Currency]", "")
                    PRNFile = PRNFile.Replace("[UnitPrice]", "")
                Else
                    PRNFile = PRNFile.Replace("[Currency]", "AED")
                    PRNFile = PRNFile.Replace("[UnitPrice]", Strings.Format(UnitPrice, "0.00"))
                End If
                PRNFile = PRNFile.Replace("[Date]", StrDate)
                PRNFile = PRNFile.Replace("[Location]", BranchName)

                If (File.Exists(FinalPRNFile)) Then
                    Using streamWriter As System.IO.StreamWriter = New System.IO.StreamWriter(FinalPRNFile, True, Encoding.UTF8)
                        streamWriter.Write(If(If(PRNFile, ""), ""))
                        streamWriter.Close()
                    End Using
                End If

                Try
                    IO.File.Copy(FinalPRNFile, PrinterPath & "\" & FileName)
                    objMain.objApplication.StatusBar.SetText("Barcode print file sent to the printer successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Thread.Sleep(300)
                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText("An error while sending to the printer. Error:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End Try

                Try
                    IO.File.Delete(FinalPRNFile)
                Catch ex As Exception
                End Try
            Else
                Me.objForm.Freeze(False)
                objMain.objApplication.StatusBar.SetText("Zebra print text is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return
            End If
        Catch exception1 As System.Exception
            Dim exception As System.Exception = exception1
            objMain.objApplication.StatusBar.SetText(exception.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
    End Sub

    Public Function Notepad_Read(ByVal QueryTextFile As String) As String
        Dim str As String = Nothing
        Try
            Dim sPath As String
            sPath = IO.Directory.GetParent(Application.ExecutablePath).ToString
            sPath = sPath & "\" & QueryTextFile

            If (Not File.Exists(sPath)) Then
                objMain.objApplication.StatusBar.SetText("File is not in this path - '" & sPath & "'", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                Dim streamReader As System.IO.StreamReader = New System.IO.StreamReader(sPath, Encoding.UTF8)
                Dim [end] As String = ""
                [end] = streamReader.ReadToEnd()
                streamReader.Close()
                str = [end]
                Return str
            End If
        Catch exception As System.Exception
            objMain.objApplication.StatusBar.SetText(exception.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
        Return str
    End Function

    Public Function GetFieldValue(ByVal query As String) As Object
        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs.DoQuery(query)
        Dim str As String = oRs.Fields.Item(0).Value.ToString()
        Return str
    End Function

    Private Sub DisableColumns()
        Me.objMatrix.Columns.Item("ItemCode").Editable = False
        Me.objMatrix.Columns.Item("Desc").Editable = False
        Me.objMatrix.Columns.Item("BarCode").Editable = False
        Me.objMatrix.Columns.Item("PartNo").Editable = False
        Me.objMatrix.Columns.Item("ItemGroup").Editable = False
        Me.objMatrix.Columns.Item("MFacturer").Editable = False
        Me.objMatrix.Columns.Item("Qty").Editable = False
    End Sub

    Private Sub DisableControls()
        Me.objForm.Items.Item("txtDocNum").Enabled = False
        Me.objForm.Items.Item("txtDocDate").Enabled = False
        Me.objForm.Items.Item("txtBP").Enabled = False
        Me.objForm.Items.Item("txtBPName").Enabled = False
        Me.objForm.Items.Item("txtDocType").Enabled = False
    End Sub

    Private Sub EnableColumns()
        Me.objMatrix.Columns.Item("ItemCode").Editable = True
        Me.objMatrix.Columns.Item("Desc").Editable = True
        Me.objMatrix.Columns.Item("BarCode").Editable = True
        Me.objMatrix.Columns.Item("PartNo").Editable = True
        Me.objMatrix.Columns.Item("ItemGroup").Editable = True
        Me.objMatrix.Columns.Item("MFacturer").Editable = True
        Me.objMatrix.Columns.Item("Qty").Editable = True
        Me.objMatrix.Columns.Item("Col_0").Editable = True
    End Sub

    Private Sub EnableControls()
        Me.objForm.Items.Item("txtDocNum").Enabled = True
        Me.objForm.Items.Item("txtDocDate").Enabled = True
        Me.objForm.Items.Item("txtBP").Enabled = True
        Me.objForm.Items.Item("txtBPName").Enabled = True
        Me.objForm.Items.Item("txtDocType").Enabled = True
    End Sub


    Sub SetNewLine()
        oDS = Me.objForm.DataSources.DBDataSources.Add("@ALE_BARCODE")
        oDS1 = Me.objForm.DataSources.DBDataSources.Add("@ALE_BARCODE1")
        objMatrix = Me.objForm.Items.Item("matrix").Specific
        objMatrix.AddRow()
        oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
        oDS1.SetValue("U_ItemCode", oDS1.Offset, "")
        oDS1.SetValue("U_ItemName", oDS1.Offset, "")
        oDS1.SetValue("U_Uom", oDS1.Offset, "")
        oDS1.SetValue("U_CodeBars", oDS1.Offset, "")
        oDS1.SetValue("U_FrgnName", oDS1.Offset, "")
        oDS1.SetValue("U_ItmGrpNm", oDS1.Offset, "")
        oDS1.SetValue("U_FirmName", oDS1.Offset, "")
        oDS1.SetValue("U_Qty", oDS1.Offset, "")
        oDS1.SetValue("U_Print", oDS1.Offset, "Y")
        oDS1.SetValue("U_PrintQty", oDS1.Offset, "")
        objMatrix.SetLineData(objMatrix.VisualRowCount)

    End Sub

End Class
