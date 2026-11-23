Imports System.Web

Public Class clsGeneratePOfromPurchaseRequest
#Region "        Declaration        "
    Dim objForm As SAPbouiCOM.Form
    Dim objApplication As SAPbouiCOM.Application
    Dim oGrid As SAPbouiCOM.Grid
    Dim oDt As SAPbouiCOM.DataTable
    Dim oEditText As SAPbouiCOM.EditText
    Dim objComboBox As SAPbouiCOM.ComboBox
    Public objCompany As SAPbobsCOM.Company
    Dim sFlag As Boolean = False
    Dim objCheckBox As SAPbouiCOM.CheckBox
#End Region
    ''test
    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("GeneratePOFromPurchaseRequest.xml", "ME_GPOFPRQ_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_GPOFPRQ_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.DataSources.UserDataSources.Add("U_FROMDT", SAPbouiCOM.BoDataType.dt_DATE)
            objForm.DataSources.UserDataSources.Add("U_TODT", SAPbouiCOM.BoDataType.dt_DATE)
            objForm.DataSources.UserDataSources.Add("U_VENDOR", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_BRANCH", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_TYPE", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_PRNO", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            objForm.DataSources.UserDataSources.Add("SelectAll", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objCheckBox = objForm.Items.Item("14").Specific
            objCheckBox.DataBind.SetBound(True, "", "SelectAll")
            objForm.Items.Item("14").AffectsFormMode = False

            oEditText = objForm.Items.Item("4").Specific
            oEditText.DataBind.SetBound(True, "", "U_FROMDT")
            objForm.Items.Item("4").AffectsFormMode = False

            oEditText = objForm.Items.Item("6").Specific
            oEditText.DataBind.SetBound(True, "", "U_TODT")
            objForm.Items.Item("6").AffectsFormMode = False

            objComboBox = objForm.Items.Item("13").Specific
            objComboBox.DataBind.SetBound(True, "", "U_BRANCH")
            objForm.Items.Item("13").AffectsFormMode = False
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, " Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join " &
                                                                 " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""  ")
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("13").DisplayDesc = True
            If objComboBox.ValidValues.Count <> 0 Then
                objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
            End If

            objComboBox = objForm.Items.Item("16").Specific
            objComboBox.DataBind.SetBound(True, "", "U_TYPE")
            objForm.Items.Item("16").AffectsFormMode = False
            objComboBox.ValidValues.Add("OPRQ", "OpenPurchaseRequest")
            objComboBox.ValidValues.Add("OPOD", "OpenPODraft")
            objComboBox.ValidValues.Add("Both", "Both")
            objForm.Items.Item("16").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            If objComboBox.ValidValues.Count <> 0 Then
                objComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            End If

            oEditText = objForm.Items.Item("8").Specific
            oEditText.DataBind.SetBound(True, "", "U_VENDOR")
            objForm.Items.Item("8").AffectsFormMode = False
            oEditText.ChooseFromListUID = "CFL_VEND"
            oEditText.ChooseFromListAlias = "CardCode"
            Me.CFLFilter(objForm.UniqueID, "CFL_VEND")


            oEditText = objForm.Items.Item("Item_3").Specific
            oEditText.DataBind.SetBound(True, "", "U_PRNO")
            objForm.Items.Item("Item_3").AffectsFormMode = False
            oEditText.ChooseFromListUID = "CFL_1"
            oEditText.ChooseFromListAlias = "DocNum"
            Me.CFLFilter2(objForm.UniqueID, "CFL_1")

            oDt = objForm.DataSources.DataTables.Add("dt1")
            oDt = objForm.DataSources.DataTables.Item("dt1")

            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ME_GPOFPRQ" And pVal.BeforeAction = False Then
                Me.CreateForm()
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "9" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("4").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        ElseIf objForm.Items.Item("6").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        ElseIf objForm.Items.Item("13").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            Me.LoadGrid(objForm.UniqueID, "Y")
                        End If
                    End If
                    If pVal.ItemUID = "14" And pVal.BeforeAction = False Then
                        objCheckBox = objForm.Items.Item("14").Specific
                        objForm.Freeze(True)
                        If objForm.Items.Item("14").Specific.checked = True Then
                            Me.LoadGrid(objForm.UniqueID, "Y")
                        ElseIf objForm.Items.Item("14").Specific.checked = False Then
                            Me.LoadGrid(objForm.UniqueID, "N")
                        End If
                        objForm.Freeze(False)
                    End If


                    If pVal.ItemUID = "3" And pVal.BeforeAction = False Then
                        Me.PostDraftPO(objForm.UniqueID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oGrid = objForm.Items.Item("10").Specific

                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_ID As String
                    CFL_ID = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_ID)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects

                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                        If (oCFL.UniqueID = "CFL_VEND") Then
                            objForm.DataSources.UserDataSources.Item("U_VENDOR").ValueEx = oDT.GetValue("CardCode", 0)
                            Dim oStaticText As SAPbouiCOM.StaticText = objForm.Items.Item("11").Specific
                            oStaticText.Caption = oDT.GetValue("CardName", 0)
                            oGrid = objForm.Items.Item("10").Specific
                            If oGrid.DataTable IsNot Nothing Then
                                oGrid.DataTable.Clear()
                            End If
                        End If
                        If (oCFL.UniqueID = "CFL_1") Then
                            objForm.DataSources.UserDataSources.Item("U_PRNO").ValueEx = oDT.GetValue("DocNum", 0)
                            oGrid = objForm.Items.Item("10").Specific
                            ''oGrid.DataTable = oDT
                            If oGrid.DataTable IsNot Nothing Then
                                oGrid.DataTable.Clear()
                            End If
                        End If

                    End If
                    If (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = True Then

                        If oCFL.UniqueID = "CFL_1" Then

                            Me.FilterAP(FormUID, "CFL_1")

                        End If

                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub LoadGrid(ByVal FormUID As String, ByVal CheckValue As String)
        Try
            objForm.Freeze(True)
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("10").Specific
            oGrid.DataTable = oDt
            oGrid.DataTable.Clear()

            Dim Query As String = ""

            If objForm.Items.Item("16").Specific.Value = "OPRQ" Then

                If objForm.Items.Item("Item_3").Specific.Value = "" Then

                    objMain.objApplication.StatusBar.SetText("Please Enter Purchase Request Number", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Else

                    Query = " Select '" & CheckValue & "'  ""Check"" , " &
            " T0.""DocEntry"" As ""DocEntry Number"" , T0.""DocNum"" As ""Document Number"" , " &
            " T1.""LineNum"" As ""Line Number"" , T1.""ItemCode"" as ""Item Code"", " &
            " T1.""Dscription"" as ""Item Description"", T1.""OpenQty"" As ""Normal Quantity"", " &
            " T1.""Quantity"" as ""Required Quantity"", " &
            " (Select ""Price"" From ""ITM1"" Where ""ItemCode"" = T1.""ItemCode"" And " &
            " ""PriceList"" = (Select ""ListNum"" From OCRD Where ""CardCode"" = '" & objForm.Items.Item("8").Specific.Value & "')) " &
            " As ""Price"" , T1.""PQTReqDate"" as ""Requested Date"", T1.""UomCode"", " &
            " IFNULL((Select TO_VARCHAR(Max(""DocEntry"")) From ""DRF1"" " &
            " Where ""BaseEntry"" = T1.""DocEntry"" And ""BaseLine"" = T1.""LineNum""+1 " &
            " And ""BaseType"" = '1470000113'),'') As ""DraftEntry"" " &
            " From OPRQ T0 Inner Join PRQ1 T1 On T0.""DocEntry"" = T1.""DocEntry"" " &
            " Where T0.""DocDate"" Between '" & objForm.Items.Item("4").Specific.Value &
            "' And '" & objForm.Items.Item("6").Specific.Value & "' And " &
            " T1.""LineVendor"" = '" & objForm.Items.Item("8").Specific.Value & "' And " &
            " T0.""BPLId"" = '" & objForm.Items.Item("13").Specific.Value & "' And " &
            " T1.""LineStatus"" = 'O' And " &
            " IFNULL((Select TO_VARCHAR(Max(""DocEntry"")) From ""DRF1"" " &
            " Where ""BaseEntry"" = T1.""DocEntry"" And ""BaseLine"" = T1.""LineNum""+1 " &
            " And ""BaseType"" = '1470000113'),'') = '' " &
            " Order By T0.""DocEntry"" "

                    Query = " Select '" & CheckValue & "'  ""Check"" , " &
             " T0.""DocEntry"" As ""DocEntry Number"" , T0.""DocNum"" As ""Document Number"" , " &
             " T1.""LineNum"" As ""Line Number"" , T1.""ItemCode"" as ""Item Code"", " &
             " T1.""Dscription"" as ""Item Description"", T1.""OpenQty"" As ""Normal Quantity"", " &
             " T1.""Quantity"" as ""Required Quantity"", " &
             " (Select ""Price"" From ""ITM1"" Where ""ItemCode"" = T1.""ItemCode"" And " &
             " ""PriceList"" = (Select ""ListNum"" From OCRD Where ""CardCode"" = '" & objForm.Items.Item("8").Specific.Value & "')) " &
             " As ""Price"" , T1.""PQTReqDate"" as ""Requested Date"", T1.""UomCode"", " &
             " IFNULL((Select TO_VARCHAR(Max(""DocEntry"")) From ""DRF1"" " &
             " Where ""BaseEntry"" = T1.""DocEntry"" And ""BaseLine"" = T1.""LineNum""+1 " &
             " And ""BaseType"" = '1470000113'),'') As ""DraftEntry"" " &
             " From OPRQ T0 Inner Join PRQ1 T1 On T0.""DocEntry"" = T1.""DocEntry"" " &
             " Where T0.""DocDate"" Between '" & objForm.Items.Item("4").Specific.Value &
             "' And '" & objForm.Items.Item("6").Specific.Value & "' And T0.""DocNum"" =  '" & objForm.Items.Item("Item_3").Specific.Value & "' and  " &
             "  (IFNULL('" & objForm.Items.Item("8").Specific.Value & "','') = '' " &
                " OR T1.""LineVendor"" = '" & objForm.Items.Item("8").Specific.Value & "') and " &
             " T0.""BPLId"" = '" & objForm.Items.Item("13").Specific.Value & "' And " &
             " T1.""LineStatus"" = 'O' And " &
             " IFNULL((Select TO_VARCHAR(Max(""DocEntry"")) From ""DRF1"" " &
             " Where ""BaseEntry"" = T1.""DocEntry"" And ""BaseLine"" = T1.""LineNum""+1 " &
             " And ""BaseType"" = '1470000113'),'') = '' " &
             " Order By T0.""DocEntry"" "

                End If

            ElseIf objForm.Items.Item("16").Specific.Value = "OPOD" Then

                Query = " Select '" & CheckValue & "'  ""Check"" , " &
            " T0.""DocEntry"" As ""DocEntry Number"" , T0.""DocNum"" As ""Document Number"" , " &
            " T1.""LineNum"" As ""Line Number"" , T1.""ItemCode"" as ""Item Code"", " &
            " T1.""Dscription"" as ""Item Description"", T1.""OpenQty"" As ""Normal Quantity"", " &
            " T1.""Quantity"" as ""Required Quantity"", " &
            " (Select ""Price"" From ""ITM1"" Where ""ItemCode"" = T1.""ItemCode"" And " &
            " ""PriceList"" = (Select ""ListNum"" From OCRD Where ""CardCode"" = '" & objForm.Items.Item("8").Specific.Value & "')) " &
            " As ""Price"" , T1.""PQTReqDate"" as ""Requested Date"", T1.""UomCode"", " &
            " IFNULL((Select TO_VARCHAR(Max(""DocEntry"")) From ""DRF1"" " &
            " Where ""BaseEntry"" = T1.""DocEntry"" And ""BaseLine"" = T1.""LineNum""+1 " &
            " And ""BaseType"" = '1470000113'),'') As ""DraftEntry"" " &
            " From OPRQ T0 Inner Join PRQ1 T1 On T0.""DocEntry"" = T1.""DocEntry"" " &
            " Where T0.""DocDate"" Between '" & objForm.Items.Item("4").Specific.Value &
            "' And '" & objForm.Items.Item("6").Specific.Value & "' And " &
             "  (IFNULL('" & objForm.Items.Item("8").Specific.Value & "','') = '' " &
            " OR T1.""LineVendor"" = '" & objForm.Items.Item("8").Specific.Value & "') and " &
            " T0.""BPLId"" = '" & objForm.Items.Item("13").Specific.Value & "' And " &
            " T1.""LineStatus"" = 'O' And " &
            " IFNULL((Select TO_VARCHAR(Max(""DocEntry"")) From ""DRF1"" " &
            " Where ""BaseEntry"" = T1.""DocEntry"" And ""BaseLine"" = T1.""LineNum"" " &
            " And ""BaseType"" = '1470000113'),'') <> '' " &
            " Order By T0.""DocEntry"" "

            ElseIf objForm.Items.Item("16").Specific.Value = "Both" Then

                Query = " Select '" & CheckValue & "'  ""Check"" , " &
            " T0.""DocEntry"" As ""DocEntry Number"" , T0.""DocNum"" As ""Document Number"" , " &
            " T1.""LineNum"" As ""Line Number"" , T1.""ItemCode"" as ""Item Code"", " &
            " T1.""Dscription"" as ""Item Description"", T1.""OpenQty"" As ""Normal Quantity"", " &
            " T1.""Quantity"" as ""Required Quantity"", " &
            " (Select ""Price"" From ""ITM1"" Where ""ItemCode"" = T1.""ItemCode"" And " &
            " ""PriceList"" = (Select ""ListNum"" From OCRD Where ""CardCode"" = '" & objForm.Items.Item("8").Specific.Value & "')) " &
            " As ""Price"" , T1.""PQTReqDate"" as ""Requested Date"", T1.""UomCode"", " &
            " IFNULL((Select TO_VARCHAR(Max(""DocEntry"")) From ""DRF1"" " &
            " Where ""BaseEntry"" = T1.""DocEntry"" And ""BaseLine"" = T1.""LineNum""+1 " &
            " And ""BaseType"" = '1470000113'),'') As ""DraftEntry"" " &
            " From OPRQ T0 Inner Join PRQ1 T1 On T0.""DocEntry"" = T1.""DocEntry"" " &
            " Where T0.""DocDate"" Between '" & objForm.Items.Item("4").Specific.Value &
            "' And '" & objForm.Items.Item("6").Specific.Value & "' And " &
             "  (IFNULL('" & objForm.Items.Item("8").Specific.Value & "','') = '' " &
            " OR T1.""LineVendor"" = '" & objForm.Items.Item("8").Specific.Value & "') and " &
            " T0.""BPLId"" = '" & objForm.Items.Item("13").Specific.Value & "' And " &
            " T1.""LineStatus"" = 'O' " &
            " Order By T0.""DocEntry"" "

            End If


            'oGrid.Data


            oGrid.DataTable.ExecuteQuery(Query)
            If oGrid.Rows.Count = 0 Then
                objMain.objApplication.MessageBox("No records found")
                Exit Sub
            End If
            oGrid.Columns.Item("Line Number").Visible = False
            oGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            For i As Integer = 0 To oGrid.Columns.Count - 1
                oGrid.Columns.Item(i).Editable = False
            Next
            For i As Integer = 0 To oGrid.Rows.Count - 1
                If oGrid.DataTable.GetValue("DraftEntry", i) <> "" Then
                    oGrid.DataTable.SetValue("Check", i, "N")
                    oGrid.CommonSetting.SetRowEditable(i + 1, False)
                Else
                    oGrid.CommonSetting.SetRowEditable(i + 1, True)
                End If
            Next

            Dim oEditTextCol As SAPbouiCOM.EditTextColumn
            oEditTextCol = oGrid.Columns.Item("DocEntry Number")
            oEditTextCol.LinkedObjectType = 1470000113

            Dim oEditTextColumn As SAPbouiCOM.EditTextColumn = oGrid.Columns.Item("Item Code")
            oEditTextColumn.LinkedObjectType = 4

            oEditTextColumn = oGrid.Columns.Item("DraftEntry")
            oEditTextColumn.LinkedObjectType = 112

            oGrid.AutoResizeColumns()
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub PostDraftPO(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("10").Specific
            Dim oPODrafts As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
            oPODrafts.DocObjectCodeEx = SAPbobsCOM.BoObjectTypes.oPurchaseOrders
            Dim DocSeries As String = ""

            Dim objBranchCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("13").Specific

            Dim GetPOSeries As String = "Select ""Series"" , ""SeriesName"" From NNM1 Where ""ObjectCode"" = '22' And ""Locked"" = 'N' And ""BPLId"" = '" & objBranchCombo.Value & "'   Order By ""Series"" Desc "
            Dim oRsGetPOSeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetPOSeries.DoQuery(GetPOSeries)
            If oRsGetPOSeries.RecordCount > 0 Then
                DocSeries = oRsGetPOSeries.Fields.Item("Series").Value
            End If

            For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If oGrid.DataTable.GetValue("Check", i) = "Y" And oGrid.DataTable.GetValue("DraftEntry", i) = "" Then
                    If DocSeries <> "" Then
                        oPODrafts.Series = DocSeries
                    End If
                    oPODrafts.CardCode = objForm.Items.Item("8").Specific.Value
                    oPODrafts.BPL_IDAssignedToInvoice = objForm.Items.Item("13").Specific.Value.Trim
                    oPODrafts.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
                    Dim sDate As String = DateTime.Now.ToString("yyyyMMdd")
                    sDate = sDate.Insert("4", "-")
                    sDate = sDate.Insert("7", "-")
                    oPODrafts.DocDate = sDate
                    oPODrafts.DocDueDate = sDate
                    oPODrafts.TaxDate = sDate

                    oPODrafts.Lines.BaseEntry = oGrid.DataTable.GetValue("DocEntry Number", i)
                    oPODrafts.Lines.BaseLine = oGrid.DataTable.GetValue("Line Number", i)
                    oPODrafts.Lines.Price = oGrid.DataTable.GetValue("Price", i)
                    oPODrafts.Lines.UserFields.Fields.Item("U_normalqty").Value = oGrid.DataTable.GetValue("Required Quantity", i)
                    oPODrafts.Lines.BaseType = "1470000113"
                    oPODrafts.Lines.Add()
                    sFlag = True
                End If
            Next
            If sFlag = True Then
                If oPODrafts.Add = 0 Then
                    objMain.objApplication.StatusBar.SetText("PO Draft Created Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    'Me.LoadGrid(objForm.UniqueID, "N")
                    Dim GetPODraftEntry As String = "Select Max(""DocEntry"") From ODRF Where ""ObjType"" = '22'  "
                    Dim oRsGetPODraftEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetPODraftEntry.DoQuery(GetPODraftEntry)
                    If oRsGetPODraftEntry.RecordCount > 0 Then
                        objMain.objApplication.OpenForm(112, "", oRsGetPODraftEntry.Fields.Item(0).Value.ToString)
                    End If
                Else
                    objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
                sFlag = False
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterAP(ByVal FormUID As String, ByVal CFL_Id As String)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item(CFL_Id)

            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition

            ' Clear old conditions
            oCFL.SetConditions(New SAPbouiCOM.Conditions)
            oCons = oCFL.GetConditions()

            ' Get values
            Dim Vendor As String = objForm.Items.Item("8").Specific.Value
            Dim bplId As String = objForm.Items.Item("13").Specific.Value
            Dim fromDate As String = objForm.Items.Item("4").Specific.Value
            Dim toDate As String = objForm.Items.Item("6").Specific.Value

            ' 🔹 Build Query (same as your CFL logic)
            Dim Query As String = "SELECT T0.""DocEntry"" FROM OPRQ T0 WHERE T0.""DocStatus"" = 'O' "

            If Vendor <> "" Then
                Query &= " AND T0.""U_VCode"" = '" & Vendor & "' "
            End If

            If bplId <> "" Then
                Query &= " AND T0.""BPLId"" = '" & bplId & "' "
            End If

            If fromDate <> "" Then
                Query &= " AND T0.""DocDate"" >= '" & fromDate & "' "
            End If

            If toDate <> "" Then
                Query &= " AND T0.""DocDate"" <= '" & toDate & "' "
            End If

            ' Execute query
            oRS.DoQuery(Query)

            ' 🔹 Apply to CFL (same OR logic like your example)
            For i As Integer = 1 To oRS.RecordCount

                oCon = oCons.Add()

                If i > 1 Then
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                End If

                oCon.Alias = "DocEntry"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = oRS.Fields.Item("DocEntry").Value

                oRS.MoveNext()
            Next

            ' If no records
            If oRS.RecordCount = 0 Then
                oCon = oCons.Add()
                oCon.Alias = "DocEntry"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = "-1"
            End If

            oCFL.SetConditions(oCons)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub

    'Sub FilterAP(ByVal FormUID As String, ByVal CFL_Id As String)

    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)

    '        Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item(CFL_Id)
    '        Dim oCons As SAPbouiCOM.Conditions
    '        Dim oCon As SAPbouiCOM.Condition

    '        ' Clear old conditions
    '        oCFL.SetConditions(New SAPbouiCOM.Conditions)
    '        oCons = oCFL.GetConditions()

    '        ' Get values
    '        Dim Vendor As String = objForm.Items.Item("8").Specific.Value
    '        Dim bplId As String = objForm.Items.Item("13").Specific.Value
    '        Dim fromDate As String = objForm.Items.Item("4").Specific.Value
    '        Dim toDate As String = objForm.Items.Item("6").Specific.Value

    '        Dim isFirst As Boolean = True

    '        ' Vendor
    '        If Not String.IsNullOrWhiteSpace(Vendor) Then
    '            oCon = oCons.Add()
    '            oCon.Alias = "U_VCode"
    '            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '            oCon.CondVal = Vendor
    '            isFirst = False
    '        End If

    '        ' BPL
    '        If Not String.IsNullOrWhiteSpace(bplId) Then
    '            oCon = oCons.Add()
    '            If Not isFirst Then oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
    '            oCon.Alias = "BPLId"
    '            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '            oCon.CondVal = bplId
    '            isFirst = False
    '        End If

    '        ' From Date
    '        If Not String.IsNullOrWhiteSpace(fromDate) Then
    '            oCon = oCons.Add()
    '            If Not isFirst Then oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

    '            oCon.Alias = "DocDate"
    '            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL
    '            oCon.CondVal = fromDate   ' <-- direct

    '            isFirst = False
    '        End If

    '        ' To Date
    '        If Not String.IsNullOrWhiteSpace(toDate) Then
    '            oCon = oCons.Add()
    '            If Not isFirst Then oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

    '            oCon.Alias = "DocDate"
    '            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_LESS_EQUAL
    '            oCon.CondVal = toDate   ' <-- direct

    '            isFirst = False
    '        End If

    '        ' Only Open Docs
    '        oCon = oCons.Add()
    '        If Not isFirst Then oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
    '        oCon.Alias = "DocStatus"
    '        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '        oCon.CondVal = "O"

    '        oCFL.SetConditions(oCons)

    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try

    'End Sub
    Sub CFLFilter(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Select Case CFL_ID
                Case "CFL_VEND"
                    Dim oConditions As SAPbouiCOM.Conditions
                    Dim oCondition As SAPbouiCOM.Condition
                    Dim oChooseFromList As SAPbouiCOM.ChooseFromList
                    Dim emptyCon As New SAPbouiCOM.Conditions
                    oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
                    oChooseFromList.SetConditions(emptyCon)
                    oConditions = oChooseFromList.GetConditions()
                    oCondition = oConditions.Add()
                    oCondition.Alias = "CardType"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = "S"
                    oChooseFromList.SetConditions(oConditions)
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub CFLFilter2(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            If CFL_ID = "CFL_1" Then

                Dim oConditions As SAPbouiCOM.Conditions
                Dim oCondition As SAPbouiCOM.Condition
                Dim oChooseFromList As SAPbouiCOM.ChooseFromList

                oChooseFromList = objForm.ChooseFromLists.Item(CFL_ID)

                ' Clear old conditions
                oChooseFromList.SetConditions(New SAPbouiCOM.Conditions)
                oConditions = oChooseFromList.GetConditions()

                ' Get values
                Dim Vendor As String = objForm.Items.Item("8").Specific.Value
                Dim bplId As String = objForm.Items.Item("13").Specific.Value
                Dim fromDate As String = objForm.Items.Item("4").Specific.Value
                Dim toDate As String = objForm.Items.Item("6").Specific.Value

                Dim isFirst As Boolean = True

                ' Vendor
                If Not String.IsNullOrWhiteSpace(Vendor) Then
                    oCondition = oConditions.Add()
                    oCondition.Alias = "CardCode"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = Vendor
                    isFirst = False
                End If

                ' BPL
                If Not String.IsNullOrWhiteSpace(bplId) Then
                    oCondition = oConditions.Add()
                    If Not isFirst Then oCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    oCondition.Alias = "BPLId"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = bplId
                    isFirst = False
                End If

                ' From Date
                If Not String.IsNullOrWhiteSpace(fromDate) Then
                    oCondition = oConditions.Add()
                    If Not isFirst Then oCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    oCondition.Alias = "DocDate"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL
                    oCondition.CondVal = Format(CDate(fromDate), "yyyyMMdd")
                    isFirst = False
                End If

                ' To Date
                If Not String.IsNullOrWhiteSpace(toDate) Then
                    oCondition = oConditions.Add()
                    If Not isFirst Then oCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    oCondition.Alias = "DocDate"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_LESS_EQUAL
                    oCondition.CondVal = Format(CDate(toDate), "yyyyMMdd")
                End If

                oChooseFromList.SetConditions(oConditions)

            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
End Class
