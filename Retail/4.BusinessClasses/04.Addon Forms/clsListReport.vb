Imports SAPbobsCOM

Public Class clsListReport



#Region "    Declaration   "
    Dim objForm, objFormX, objMIForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix

    Dim oEditText As SAPbouiCOM.EditText
    Dim oDt, oDt1, oDataTable As SAPbouiCOM.DataTable
    Dim objGrid, oGrid As SAPbouiCOM.Grid
    Dim sSupplierCode, sSupplierName As String
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim LostFocusFlag As Boolean = False
#End Region

#Region "    Public Methods   "

    Sub CreateForm()
        Try
            ' Load form and get reference
            objMain.objUtilities.LoadForm("ListReport.xml", "ListReport_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ListReport_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            'sSupplierCode = CardCode

            ' Keep lookup references
            If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = False Then
                'Retail.MainCls.ohtLookUpForm.Add(objForm.UniqueID, FormUID)
            End If

            ' Datasources and controls
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_LSRPT")
            LoadBranchCombo()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("clsSubItemList.CreateForm: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region




#Region "    Events   "

    Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)


                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                    (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then



                    End If
                    If pVal.ItemUID = "Load" And pVal.BeforeAction = False Then
                        objForm.Freeze(True)
                        Me.LoadListChargesGrid(objForm.UniqueID)
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
                    If oCFL.UniqueID = "CFL_0" And pVal.BeforeAction = True Then
                    End If

                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                        If oCFL.UniqueID = "CFL_0" Then
                            oDBs_Head.SetValue("U_CardCode", oDBs_Head.Offset, oDT.GetValue("CardCode", 0))
                            oDBs_Head.SetValue("U_CardName", oDBs_Head.Offset, oDT.GetValue("CardName", 0))

                        End If
                    End If

                    'Madhu when cardcode removing along with cardname
                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                    If pVal.ItemUID = "CardCode" And pVal.BeforeAction = False Then

                        ' 8 = Backspace, 46 = Delete
                        If pVal.CharPressed = 8 Or pVal.CharPressed = 46 Then

                            objForm = objMain.objApplication.Forms.Item(FormUID)

                            Dim oDBs_Head As SAPbouiCOM.DBDataSource =
                                objForm.DataSources.DBDataSources.Item("@TNX_LSRPT")

                            oDBs_Head.SetValue("U_CardCode", oDBs_Head.Offset, "")

                            oDBs_Head.SetValue("U_CardName", oDBs_Head.Offset, "")
                        End If

                    End If



                     'madhu  Grid Link Button

                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                    If pVal.ItemUID = "Grid" And pVal.ColUID = "Item listing Number" And pVal.Before_Action = False Then
                        objGrid = objForm.Items.Item("Grid").Specific
                        Dim code As String = objGrid.DataTable.GetValue("Item listing Number", pVal.Row)
                        Dim objNewForm As SAPbouiCOM.Form
                        objMain.objItemList.CreateForm(objNewForm, oDBs_Head.GetValue("U_CardCode", 0), oDBs_Head.GetValue("U_CardName", 0), "Y")
                        objNewForm = objMain.objApplication.Forms.ActiveForm
                        objNewForm.Freeze(True)
                        objNewForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objNewForm.Items.Item("11").Specific.Value = code
                        'objNewForm.Items.Item("CardCode").Specific.Value = oDBs_Head.GetValue("U_CardCode", 0)
                        'objNewForm.Items.Item("9").Specific.Value = oDBs_Head.GetValue("U_CardName", 0)
                        objNewForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        objNewForm.Freeze(False)
                    End If

                  'madhu end


                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = True And pVal.BeforeAction = False Then
                        Retail.MainCls.ohtLookUpForm.Remove(objForm.UniqueID)
                    End If
            End Select

        Catch ex As Exception
            Try
                objForm.Freeze(False)
            Catch
            End Try
            objMain.objApplication.StatusBar.SetText("clsListReport.ItemEvent: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)

        Try

            If pVal.MenuUID = "Report" And pVal.BeforeAction = False Then
                Me.CreateForm()
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then

            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then


            End If
            ' Me.BranchLoad(objForm.UniqueID)

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

    Private Sub LoadListChargesGrid(ByVal FormUID As String)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oGrid = CType(objForm.Items.Item("Grid").Specific, SAPbouiCOM.Grid)
            oDt1 = objForm.DataSources.DataTables.Item("DT_0")
            oGrid.DataTable = oDt1

            Try
                oGrid.DataTable.Clear()
            Catch ex As Exception
            End Try

            Dim fromDate As String = objForm.Items.Item("FDate").Specific.Value
            Dim toDate As String = objForm.Items.Item("ToDate").Specific.Value
            Dim cardCode As String = objForm.Items.Item("CardCode").Specific.Value
            Dim branch As String = objForm.Items.Item("Branch").Specific.Value

            If String.IsNullOrEmpty(fromDate) OrElse String.IsNullOrEmpty(toDate) Then
                objMain.objApplication.StatusBar.SetText(
                "Please select both From Date and To Date.",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objForm.Freeze(False)
                Exit Sub
            End If

            Dim listChargeQuery As String =
    "SELECT T1.""U_Date"" AS ""Date"", T1.""U_SKU"" AS ""No of SKUs"", T1.""U_Amount"" AS ""Amount"", T1.""U_NOU"" AS ""Number of Outlet's"", " &
    "T1.""U_JEPosted"" AS ""JE Number"", T1.""U_Remarks"" AS ""Remarks"", T1.""U_BRANCH"" AS ""Branch"", T0.""U_CardCode"" AS ""CardCode"", " &
    "T0.""DocNum"" AS ""Item listing Number"" " &
    "FROM ""@SBO_LIST"" T0 INNER JOIN ""@SBO_LIST_C0"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
    "WHERE T1.""U_Date"" BETWEEN '" & fromDate & "' AND '" & toDate & "' AND ('" & branch & "' = 'All' OR T1.""U_BRANCH"" = '" & branch & "')"

            '"SELECT " &
            '" T1.""U_Date"" AS ""Date"", " &
            '" T1.""U_SKU"" AS ""No of SKUs"", " &
            '" T1.""U_Amount"" AS ""Amount"", " &
            '" T1.""U_NOU"" AS ""Number of Outlet's"", " &
            '" T1.""U_JEPosted"" AS ""JE Number"", " &
            '" T1.""U_Remarks"" AS ""Remarks"", " &
            '" T1.""U_BRANCH"" AS ""Branch"", " &
            '" T0.""U_CardCode"" AS ""CardCode"", " &
            '" T0.""DocNum"" AS ""Item listing Number"" " &
            '"FROM ""@SBO_LIST"" T0 " &
            '"INNER JOIN ""@SBO_LIST_C0"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
            ''"WHERE T1.""U_Date"" BETWEEN '" & fromDate & "' AND '" & toDate & "'"


            If Not String.IsNullOrEmpty(cardCode) Then
                listChargeQuery &= " AND T0.""U_CardCode"" = '" & cardCode & "'"
            End If

            'If Not String.IsNullOrEmpty(branch) Then
            '    listChargeQuery &= " AND T1.""U_BRANCH"" = '" & branch & "'"
            'End If

            oGrid.DataTable.ExecuteQuery(listChargeQuery)

            oGrid.Columns.Item("CardCode").LinkedObjectType = "2"

            oGrid.Columns.Item("Item listing Number").LinkedObjectType = "SBO_LISTUDO"


            For i As Integer = 0 To oGrid.Columns.Count - 1
                oGrid.Columns.Item(i).Editable = False
            Next

            oGrid.AutoResizeColumns()
            objForm.Freeze(False)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            objForm.Freeze(False)
        End Try

    End Sub

    Private Sub LoadBranchCombo()
        Try
            Dim oCombo As SAPbouiCOM.ComboBox
            oCombo = CType(objForm.Items.Item("Branch").Specific, SAPbouiCOM.ComboBox)

            'oCombo.ValidValues.Clear()

            Dim oRs As SAPbobsCOM.Recordset =
            objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'oRs.DoQuery("SELECT ""BPLId"", ""BPLName"" FROM ""OBPL"" ORDER BY ""BPLId""")
            oRs.DoQuery("SELECT T2.""BPLId"",T2.""BPLName"" FROM USR6 T0  INNER JOIN OUSR T1 ON T0.""UserCode"" = T1.""USER_CODE"" " &
                    " INNER JOIN OBPL T2 ON T0.""BPLId"" = T2.""BPLId"" WHERE T1.""USER_CODE"" = '" & objMain.objCompany.UserName.ToString().Trim & "' ")



            While Not oRs.EoF
                Dim sBranchId As String = oRs.Fields.Item("BPLId").Value.ToString()
                Dim sBranchName As String = oRs.Fields.Item("BPLName").Value.ToString()

                If Not ComboValueExists(oCombo, sBranchId) Then
                    oCombo.ValidValues.Add(sBranchId, sBranchName)
                End If

                oRs.MoveNext()
            End While

            If oCombo.ValidValues.Count > 0 Then
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "LoadBranchCombo ERROR: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Function ComboValueExists(ByVal oCombo As SAPbouiCOM.ComboBox,
                                  ByVal value As String) As Boolean
        For i As Integer = 0 To oCombo.ValidValues.Count - 1
            If oCombo.ValidValues.Item(i).Value = value Then
                Return True
            End If
        Next
        Return False
    End Function


#End Region


End Class



