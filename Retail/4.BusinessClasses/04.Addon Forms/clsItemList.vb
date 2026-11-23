Imports SAPbobsCOM

Public Class clsItemList

#Region "    Declaration   "
    Dim objForm, objFormX As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details, oDBs_Details1, oDBs_Details2 As SAPbouiCOM.DBDataSource
    Dim objMatrix, objMatrix1, objMatrix2 As SAPbouiCOM.Matrix
    Dim oEditText As SAPbouiCOM.EditText
    Dim oDt As SAPbouiCOM.DataTable
    Dim sSupplierCode, sSupplierName As String
    Dim oGrid As SAPbouiCOM.Grid
    Dim objComboBox, objComboBox1, objComboBox2 As SAPbouiCOM.ComboBox
    Dim TERM As String
    Dim Chargevalue As Double = 0
    Dim APPIDX As String = ""

#End Region

    Sub CreateForm(ByVal FormUID As String, ByVal CardCode As String, ByVal CardName As String, ByVal IfITLSExist As String, Optional APPID As String = "")
        Try

            'If objMain IsNot Nothing AndAlso objMain.objCompany IsNot Nothing Then
            '    ' use the main connected company object instead of a newly constructed one
            '    clsItemList.objCompany = objMain.objCompany
            'End If

            objMain.objUtilities.LoadForm("ItemListingNew.xml", "ME_ITMLIST", ResourceType.Embeded)
            ' objMain.objUtilities.LoadForm("ItemlistNew.xml", "ME_Item_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_ITMLIST", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            objMain.objApplication.StatusBar.SetText("Loading Item list Form, please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            sSupplierCode = CardCode
            'objForm.EnableMenu("1293", True)

            ''Enabling Print Layout Icon
            'Try
            '    objForm.ReportType = "ME_ITMLIST"

            '    Try
            '        CreateLayoutFolder_ItemList()
            '    Catch ex As Exception
            '    End Try

            '    Try
            '        RemoveHandler objMain.objApplication.LayoutKeyEvent, AddressOf Me.HandleLayoutKeyEvent
            '    Catch
            '    End Try
            '    Try
            '        AddHandler objMain.objApplication.LayoutKeyEvent, AddressOf Me.HandleLayoutKeyEvent
            '    Catch ex As Exception
            '    End Try
            'Catch ex As Exception
            'End Try

            If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = False Then
                Retail.MainCls.ohtLookUpForm.Add(objForm.UniqueID, FormUID)
            End If
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_LIST")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_LIST_C0")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_LIST_C1")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_LIST_C2")

            objMatrix = objForm.Items.Item("MtxDT").Specific
            'objMatrix1 = objForm.Items.Item("Mtx2").Specific
            objMatrix2 = objForm.Items.Item("Mtx3").Specific

            objComboBox = objForm.Items.Item("12").Specific
            Dim GetSeriesQty As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'SBO_LISTUDO' And ""Locked"" = 'N'"
            objMain.objUtilities.ComboBoxLoadValuesSeries(objForm.Items.Item("12").Specific, GetSeriesQty)
            objForm.Items.Item("12").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            'If IfITLSExist = "Y" Then
            '    objForm.Freeze(True)
            '    objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            '    objForm.Items.Item("12").Specific.value = series
            '    objForm.Items.Item("11").Specific.value = docnum
            '    objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            '    objForm.Freeze(False)
            '    objMatrix.AutoResizeColumns()
            'End If

            If IfITLSExist = "Y" Then
                Try
                    objMain.objApplication.StatusBar.SetText("Finding item list...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    objForm.Freeze(True)
                    'Dim oSeriesCombo As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("12").Specific, SAPbouiCOM.ComboBox)
                    'If Not String.IsNullOrEmpty(series) AndAlso oSeriesCombo.ValidValues.Count > 0 Then
                    '    oSeriesCombo.Select(series, SAPbouiCOM.BoSearchKey.psk_ByValue)
                    'End If
                    'CType(objForm.Items.Item("11").Specific, SAPbouiCOM.EditText).Value = docnum
                    objForm.Items.Item("AppID").Specific.Value = APPID
                    objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                    objForm.Items.Item("Item_4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    objMatrix.AutoResizeColumns()
                    'madhu
                    Me.CellMasking(objForm.UniqueID)

                    objMain.objApplication.StatusBar.SetText("Item list loaded.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    '--For Testing Purpose Only, Please Remove
                    'objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                    objForm.Freeze(False)
                    Exit Sub
                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText("Error finding data.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    objForm.Freeze(False)
                Finally
                    objForm.Freeze(False)
                End Try


            End If

            ' objForm.Items.Item("14").DisplayDesc = True
            'objMain.objUtilities.MatrixComboBoxValues(objMatrix1.Columns.Item("Brach"), sQury)
            'objMatrix1.Columns.Item("Brach").DisplayDesc = True
            'objForm.DataBrowser.BrowseBy = "11"
            If IfITLSExist = "N" Then
                sSupplierName = CardName
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "SBO_LISTUDO"))
                oDBs_Head.SetValue("U_CardCode", oDBs_Head.Offset, sSupplierCode)
                oDBs_Head.SetValue("U_CardName", oDBs_Head.Offset, sSupplierName)

                Dim GetDflt As String = "Select ""DfltSeries"" From ONNM Where ""ObjectCode"" = 'SBO_LISTUDO'"
                Dim oRsGetDflt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetDflt.DoQuery(GetDflt)
                oDBs_Head.SetValue("U_SERIES", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)
                oDBs_Head.SetValue("Series", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)

                Me.AutoDocentryNumber()

                'If objComboBox.ValidValues.Count <> 0 Then
                '    objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
                'End If
                objForm.Items.Item("Item_4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                objMatrix.Clear()
                oDBs_Details.Clear()
                objMatrix.FlushToDataSource()

                objMatrix2.Clear()
                oDBs_Details2.Clear()
                objMatrix2.FlushToDataSource()




                Me.SetNewLine(objForm.UniqueID)
                Me.SetNewLine1(objForm.UniqueID)
            End If
            Try
                objComboBox = objMatrix.Columns.Item("BRANCH").Cells.Item(1).Specific  ''objForm.Items.Item("14").Specific
                Dim sQury As String = " Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join " &
                                      " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""  "
                objMain.objUtilities.ComboBoxLoadValues(objComboBox, sQury)
                objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                If objComboBox.ValidValues.Count <> 0 Then
                    objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
                End If
            Catch ex As Exception

            End Try

            Try
                objComboBox1 = objMatrix2.Columns.Item("BRANCH").Cells.Item(1).Specific  ''objForm.Items.Item("14").Specific
                Dim sQury As String = " Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join " &
                                      " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""  "
                objMain.objUtilities.ComboBoxLoadValues(objComboBox1, sQury)
                objComboBox1.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                If objComboBox1.ValidValues.Count <> 0 Then
                    objComboBox1.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
                End If
            Catch ex As Exception

            End Try

            Try
                objComboBox2 = objMatrix2.Columns.Item("VATCODE").Cells.Item(1).Specific  ''objForm.Items.Item("14").Specific
                Dim sQury As String = "SELECT ""Code"",""Name"" FROM OVTG WHERE IFNULL(""Account"",'')<>'' and ""Locked""='N'"

                objMain.objUtilities.ComboBoxLoadValues(objComboBox2, sQury)
                objComboBox2.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            Catch ex As Exception


            End Try

            'vatcode matrix1
            Try
                objComboBox2 = objMatrix1.Columns.Item("VATCODE").Cells.Item(1).Specific  ''objForm.Items.Item("14").Specific
                ' Dim sQury As String = "SELECT ""Code"",""Name"" FROM OVTG WHERE IFNULL(""Account"",'')<>'' and ""Locked""='N'"
                Dim sQury As String = "SELECT ""Code"",""Name"" FROM OVTG WHERE IFNULL(""Account"",'')<>'' AND ""Locked""='N' AND UPPER(TRIM(""Name"")) NOT LIKE '%OUTPUT%'"

                objMain.objUtilities.ComboBoxLoadValues(objComboBox2, sQury)
                objComboBox2.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            Catch ex As Exception

            End Try
            objForm.Items.Item("Item_4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            objForm.Items.Item("Submit").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Submit").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            APPIDX = objForm.Items.Item("AppID").Specific.Value

            ' Set default values: Status = "Open", Type = "Cost Fluctuation"
            Try
                ' Set DB datasource default values (header)
                oDBs_Head.SetValue("U_Statu", oDBs_Head.Offset, "Open")

            Catch ex As Exception
                ' ignore if fields not present
            End Try

            ' Me.AutoDocentryNumber()
            'Me.SetNewLine(objForm.UniqueID)
            ' Me.SetNewLine1(objForm.UniqueID)
            Me.CellMasking(objForm.UniqueID)
            'Loading Charge per unit form bussiness partner
            Me.LoadChargeFromOCRD(CardCode)
            objMain.objApplication.StatusBar.SetText("Form Loaded Successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub CreateForm_Find(ByVal DocNumberItemListing As String)
        Try
            objMain.objUtilities.LoadForm("ItemListingNew.xml", "ME_ITMLIST", ResourceType.Embeded)

            objForm = objMain.objApplication.Forms.GetForm("ME_ITMLIST",
               objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_LIST")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_LIST_C0")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_LIST_C1")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_LIST_C2")

            objMatrix = objForm.Items.Item("MtxDT").Specific
            objMatrix2 = objForm.Items.Item("Mtx3").Specific

            If Not String.IsNullOrEmpty(DocNumberItemListing) Then

                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                objForm.Items.Item("AppID").Specific.Value = DocNumberItemListing
                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                objForm.Items.Item("Item_4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            End If
            objForm.Freeze(False)

        Catch ex As Exception
            If objForm IsNot Nothing Then objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub CellMasking(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Items.Item("CardCode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("CardCode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "2", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("Statu").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Statu").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Statu").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Statu").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_LIST")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_LIST_C0")
            objMatrix = objForm.Items.Item("MtxDT").Specific

            objMatrix.AddRow()
            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
            oDBs_Details.SetValue("U_Date", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_SKU", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Amount", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_VATCODE", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_VATAMT", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_TAMNT", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_BRANCH", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_NOU", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_JEPosted", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Remarks", oDBs_Details.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub SetNewLine1(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_LIST")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_LIST_C2")
            objMatrix2 = objForm.Items.Item("Mtx3").Specific

            objMatrix2.AddRow()
            oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, objMatrix2.VisualRowCount)
            oDBs_Details2.SetValue("U_DESC", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_AMT", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_VATCODE", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_VATAMT", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_BRANCH", oDBs_Details.Offset, "")
            oDBs_Details2.SetValue("U_REMARKS", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_JESTATUS", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_JE", oDBs_Details2.Offset, "")


            objMatrix2.SetLineData(objMatrix2.VisualRowCount)
            objMatrix2.AutoResizeColumns()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, BubbleEvent As Boolean)
        Try



            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK
                    If (pVal.ItemUID = "MtxDT" Or pVal.ItemUID = "Mtx3") And pVal.BeforeAction = False Then
                        TERM = pVal.ItemUID
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If (pVal.ItemUID = "MtxDT" Or pVal.ItemUID = "Mtx3") And pVal.BeforeAction = False Then
                        TERM = pVal.ItemUID
                    End If


                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        '  If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        If Validation(FormUID) = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If

                        'Dim chkIfExist As String = "Select ""DocNum"" From ""@SBO_LIST"" Where ""DocEntry""=(Select MAX(""DocEntry"") From ""@SBO_LIST"")"
                        'Dim oRschkIfExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'oRschkIfExist.DoQuery(chkIfExist)
                        'If oRschkIfExist.RecordCount > 0 Then
                        '    objForm.Freeze(True)
                        '    objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        '    objForm.Items.Item("11").Specific.value = oRschkIfExist.Fields.Item("DocNum").Value
                        '    objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        '    objForm.Freeze(False)


                        'End If

                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        If Me.Validate1() = False Then

                            BubbleEvent = False
                        End If
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                        Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUID)
                        oForm.Freeze(True)
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        oForm.Items.Item("AppID").Specific.value = APPIDX
                        oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oForm.Freeze(False)

                    End If
                    'Approval method here
                    'submit
                    If pVal.ItemUID = "Submit" And pVal.BeforeAction = False And SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Try
                            Dim DocNum As String = objForm.Items.Item("AppID").Specific.value

                            Me.ApprovalTrigger("ITEML", DocNum, "@SBO_LIST", "U_AppID", "U_Statu", "ItemListing")

                        Catch ex As Exception
                        End Try

                    End If

                    If pVal.ItemUID = "PostJE" And pVal.BeforeAction = False And objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                        Me.PostJE()
                    End If
                    'If pVal.ItemUID = "Post" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    '    Me.PostJENew(objForm.UniqueID)
                    'End If
                    If pVal.ItemUID = "Post" And pVal.BeforeAction = False And objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                        Me.PostJEWithVATcode()
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
                            objMatrix2.GetLineData(objMatrix2.VisualRowCount)

                            oDBs_Details2.SetValue("U_VATACC", oDBs_Details2.Offset, oDT.GetValue("AcctCode", 0))
                            objMatrix2.SetLineData(objMatrix2.VisualRowCount)
                        End If
                    End If

                    If oCFL.UniqueID = "CFL_1" And pVal.BeforeAction = True Then
                        Me.CFLFilter(FormUID, "CFL_1")

                    End If

                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                        If oCFL.UniqueID = "CFL_1" Then
                            objMatrix2.GetLineData(objMatrix2.VisualRowCount)

                            oDBs_Details2.SetValue("U_DESC", oDBs_Details2.Offset, oDT.GetValue("Name", 0))
                            objMatrix2.SetLineData(objMatrix2.VisualRowCount)
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix2 = objForm.Items.Item("Mtx3").Specific

                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_LIST_C2")

                    If pVal.ItemUID = "Mtx3" And pVal.ColUID = "VATCODE" And pVal.BeforeAction = False Then
                        Try
                            If objMatrix2.Columns.Item("JE").Cells.Item(pVal.Row).Specific.Value.ToString() = "" Then
                                Dim vatCode As String = objMatrix2.Columns.Item("VATCODE").Cells.Item(pVal.Row).Specific.Value.ToString()
                                Dim Amount As String = objMatrix2.Columns.Item("AMT").Cells.Item(pVal.Row).Specific.Value.ToString()
                                Dim qStrVatPer As String = "SELECT ""Rate"" FROM OVTG WHERE ""Code""='" + vatCode + "'"
                                Dim rSetVatPer As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                rSetVatPer.DoQuery(qStrVatPer)
                                If rSetVatPer.RecordCount > 0 Then
                                    Dim Rate As Double = CDbl(rSetVatPer.Fields.Item("Rate").Value)
                                    Dim VatAmount As Double = (Rate / 100) * Amount
                                    objMatrix2.FlushToDataSource()
                                    oDBs_Details2.SetValue("U_VATAMT", oDBs_Details2.Offset, VatAmount.ToString())
                                    objMatrix2.LoadFromDataSource()
                                    ''objMatrix2.Columns.Item("VATAMT").Cells.Item(pVal.Row).Specific.Value = VatAmount.ToString()
                                End If
                            End If
                        Catch ex As Exception

                        End Try
                    End If

                    objMatrix1 = objForm.Items.Item("MtxDT").Specific
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_LIST_C0")

                    If pVal.ItemUID = "MtxDT" And pVal.ColUID = "VATCODE" And pVal.BeforeAction = False Then
                        Try
                            If objMatrix1.Columns.Item("Remarks").Cells.Item(pVal.Row).Specific.Value.ToString() = "" Then

                                Dim vatCode As String = objMatrix1.Columns.Item("VATCODE").Cells.Item(pVal.Row).Specific.Value.ToString()
                                Dim Amount As Double = 0
                                Double.TryParse(objMatrix1.Columns.Item("Amount").Cells.Item(pVal.Row).Specific.Value.ToString(), Amount)

                                If vatCode <> "" AndAlso Amount > 0 Then

                                    Dim qStrVatPer As String = "SELECT ""Rate"" FROM OVTG WHERE ""Code""='" + vatCode + "'"
                                    Dim rSetVatPer As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    rSetVatPer.DoQuery(qStrVatPer)

                                    If rSetVatPer.RecordCount > 0 Then

                                        Dim Rate As Double = CDbl(rSetVatPer.Fields.Item("Rate").Value)
                                        Dim VatAmount As Double = Math.Round((Rate / 100) * Amount, 2)

                                        objForm.Freeze(True)

                                        ' ✅ IMPORTANT FIX → Update through DataSource
                                        objMatrix1.FlushToDataSource()

                                        oDBs_Details1.SetValue("U_VATAMT", pVal.Row - 1, VatAmount.ToString("0.00"))

                                        ' Calculate Total Amount
                                        Dim totalAmount As Double = Math.Round(Amount + VatAmount, 2)
                                        oDBs_Details1.SetValue("U_TAMNT", pVal.Row - 1, totalAmount.ToString("0.00"))

                                        objMatrix1.LoadFromDataSource()

                                        objForm.Freeze(False)

                                        'objMatrix1.Columns.Item("U_BRANCH").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            objForm.Freeze(False)
                        End Try
                    End If





                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                    If pVal.ItemUID = "MtxDT" AndAlso (pVal.ColUID = "SKU" OrElse pVal.ColUID = "Charge" OrElse pVal.ColUID = "Remarks") AndAlso pVal.BeforeAction = False Then
                        Try


                            Dim skuCount As Double = 0
                            Dim charge As Double = 0
                            Double.TryParse(objMatrix.Columns.Item("SKU").Cells.Item(pVal.Row).Specific.Value, skuCount)
                            Double.TryParse(objMatrix.Columns.Item("Charge").Cells.Item(pVal.Row).Specific.Value, charge)
                            'Me.Chargevalue = charge
                            Dim Amount As Double = skuCount * charge
                            objMatrix.Columns.Item("Amount").Cells.Item(pVal.Row).Specific.Value = Amount

                            objMatrix.Columns.Item("Remarks").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Catch ex As Exception

                        End Try
                        'Me.CalculateTotalListing(objForm)
                    End If

                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_LIST")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_LIST_C0")
                    objMatrix = objForm.Items.Item("MtxDT").Specific
                    If pVal.ItemUID = "MtxDT" And pVal.ColUID = "#" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        'Dim RowDetailsForm As SAPbouiCOM.Form = objMain.objApplication.Forms.GetForm("1", objMain.objApplication.Forms.ActiveForm.TypeCount)
                        'RowDetailsForm.Close()

                        '--madhu
                        Dim DocNumItemListing As String = oDBs_Head.GetValue("DocNum", 0).Trim()
                        Dim ApplicationIDItemLIsting As String = oDBs_Head.GetValue("U_AppID", 0).Trim()



                        Dim CheckIfTCExist1 As String = "Select * From ""@SBO_LS"" Where ""U_CardCode""='" & oDBs_Head.GetValue("U_CardCode", 0) & "' And ""U_LineNum"" = '" & pVal.Row & "' "
                        Dim oRsCheckIfTCExist1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsCheckIfTCExist1.DoQuery(CheckIfTCExist1)

                        ' Dim CheckIfTCExist As String = "SELECT T1.""U_SKU"",T0.""DocNum"",T0.""U_AppID"" " & "FROM ""@SBO_LIST"" T0 " & "INNER JOIN ""@SBO_LIST_C0"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
                        ' "WHERE T0.""U_CardCode"" = '" & oDBs_Head.GetValue("U_CardCode", 0) & "' " & "AND T1.""LineId"" = '" & pVal.Row & "'"

                        Dim CheckIfTCExist As String = "SELECT T1.""U_SKU"",T0.""DocNum"",T0.""U_AppID"" " & "FROM ""@SBO_LIST"" T0 INNER JOIN ""@SBO_LIST_C0"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
                        "WHERE T0.""U_CardCode"" = '" & oDBs_Head.GetValue("U_CardCode", 0).Trim() & "' AND T1.""LineId"" = '" & pVal.Row & "' AND T0.""DocNum"" = '" & DocNumItemListing & "' AND T0.""U_AppID"" = '" & ApplicationIDItemLIsting & "'"


                        Dim oRsCheckIfTCExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsCheckIfTCExist.DoQuery(CheckIfTCExist)

                        Dim NoOfSKUValue As String = ""
                        If oRsCheckIfTCExist.RecordCount > 0 Then
                            NoOfSKUValue = oRsCheckIfTCExist.Fields.Item("U_SKU").Value.ToString()
                        End If

                        Dim DocNumItemListingValue As String = ""
                        If oRsCheckIfTCExist.RecordCount > 0 Then
                            DocNumItemListingValue = oRsCheckIfTCExist.Fields.Item("DocNum").Value.ToString()
                        End If

                        Dim ApplicationIDItemListingValue As String = ""
                        If oRsCheckIfTCExist.RecordCount > 0 Then
                            ApplicationIDItemListingValue = oRsCheckIfTCExist.Fields.Item("U_AppID").Value.ToString()
                        End If

                        If oRsCheckIfTCExist1.RecordCount = 0 Then
                            objMain.objSubItemList.CreateForm(objForm.UniqueID, oDBs_Head.GetValue("U_CardCode", 0), oDBs_Head.GetValue("U_CardName", 0), pVal.Row, NoOfSKUValue, DocNumItemListingValue, ApplicationIDItemListingValue, "N")
                        Else
                            objMain.objSubItemList.CreateForm(objForm.UniqueID, oDBs_Head.GetValue("U_CardCode", 0), oDBs_Head.GetValue("U_CardName", 0), pVal.Row, NoOfSKUValue, DocNumItemListingValue, ApplicationIDItemListingValue, "Y")
                        End If
                    End If

                    If pVal.ItemUID = "MtxDT" And pVal.ColUID = "NOU" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        '--Naresh
                        'Dim CheckIfTCExist As String = "SELECT * FROM ""@SBO_OUTLET"" " &
                        '           "WHERE ""U_CardCode""='" & oDBs_Head.GetValue("U_CardCode", 0) & "' " &
                        '           "AND ""U_LineNum"" = '" & pVal.Row & "'"

                        Dim CheckIfTCExist As String = "SELECT * FROM ""@SBO_OUTLET"" " &
                                   "WHERE ""U_BaseNo""='" & oDBs_Head.GetValue("DocNum", 0) & "' " &
                                   "AND ""U_LineNum"" = '" & pVal.Row & "'"

                        Dim oRsCheckIfTCExist As SAPbobsCOM.Recordset =
                       objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        oRsCheckIfTCExist.DoQuery(CheckIfTCExist)

                        Dim QueryAmount As String = "SELECT T1.""U_Amount"",T1.""U_NOU"",T1.""U_VATCODE"" " & "FROM ""@SBO_LIST"" T0 " & "INNER JOIN ""@SBO_LIST_C0"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
                                                   "WHERE T0.""DocNum"" = '" & oDBs_Head.GetValue("DocNum", 0) & "' " & "AND T1.""LineId"" = '" & pVal.Row & "'"

                        Dim oRsAmount As SAPbobsCOM.Recordset =
                        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        oRsAmount.DoQuery(QueryAmount)

                        Dim AmountValue As String = ""
                        Dim OutletValue As String = ""
                        Dim VatCodeValue As String = ""
                        If oRsAmount.RecordCount > 0 Then
                            AmountValue = oRsAmount.Fields.Item("U_Amount").Value.ToString()
                            OutletValue = oRsAmount.Fields.Item("U_NOU").Value.ToString()
                            VatCodeValue = oRsAmount.Fields.Item("U_VATCODE").Value.ToString()
                        End If


                        If oRsCheckIfTCExist.RecordCount = 0 Then
                            objMain.objOutlet.CreateForm(objForm.UniqueID, oDBs_Head.GetValue("U_CardCode", 0), oDBs_Head.GetValue("U_CardName", 0), oDBs_Head.GetValue("DocNum", 0), pVal.Row, AmountValue, OutletValue, VatCodeValue, "N")


                        Else
                            objMain.objOutlet.CreateForm(objForm.UniqueID, oDBs_Head.GetValue("U_CardCode", 0), oDBs_Head.GetValue("U_CardName", 0), oRsCheckIfTCExist.Fields.Item("U_BaseNo").Value, pVal.Row, AmountValue, OutletValue, VatCodeValue, "Y")

                        End If

                    End If
                    '===========================================================
                    ' 2)  NEW FUNCTIONALITY → DOUBLE CLICK ON U_SKU opens CategorySplit
                    '===========================================================
                    If pVal.ItemUID = "MtxDT" And pVal.ColUID = "SKU" And pVal.BeforeAction = False _
                        And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                        Dim CurrentSKU As String = objMatrix.Columns.Item("SKU").Cells.Item(pVal.Row).Specific.Value
                        '👉 Get CardCode from Header
                        Dim CardCode As String
                        CardCode = CType(objForm.Items.Item("CardCode").Specific, SAPbouiCOM.EditText).Value

                        Dim DocNum As String
                        DocNum = CType(objForm.Items.Item("11").Specific, SAPbouiCOM.EditText).Value

                        'Call your Category Split Form
                        Dim objCatSplit As New clsCategorySplit
                        objCatSplit.CreateForm(CurrentSKU, pVal.Row, CardCode, DocNum)

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

    Public Function Validate1() As Boolean

        Try

            Try
                CheckDocNumAndGenerateAutoID()
            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText(ex.Message & "   Errors in Validation Function", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message & "Errors in Validation Function", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Public Sub CheckDocNumAndGenerateAutoID()

        Try
            Dim docnumcc As String = Me.objForm.Items.Item("AppID").Specific.Value

            Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            '--- Check if DocNum already exists
            Dim checkQry As String =
            "SELECT COUNT(*) FROM ""@SBO_LIST"" WHERE ""U_AppID"" = '" & docnumcc & "'"

            rs.DoQuery(checkQry)

            If rs.Fields.Item(0).Value > 0 Then
                'Naresh
                '--- Generate new Auto App ID Number
                'Dim autoQry As String =
                '"SELECT " &
                '" 'MR' || RIGHT(YEAR(CURRENT_DATE),2) || '-' || " &
                '" LPAD( IFNULL( MAX(CAST(SUBSTRING(""U_ReqNo"",7) AS INTEGER)) + 1, 1 ), 4, '0' ) " &
                '"FROM ""@SBO_MATERIAL"" " &
                '"WHERE ""U_ReqNo"" LIKE 'MR' || RIGHT(YEAR(CURRENT_DATE),2) || '-%'"

                Dim Str As String = "SELECT 'IL' || LPAD('',6 - LENGTH(LTRIM(RTRIM(REPLACE(IFNULL(MAX(""U_AppID""), '0'), 'IL', '') + 1))),'0') || LTRIM(RTRIM(REPLACE(IFNULL(MAX(""U_AppID""), '0'), 'IL', '') + 1)) FROM ""@SBO_LIST"";"


                rs.DoQuery(Str)

                oDBs_Head.SetValue("U_AppID", 0, rs.Fields.Item(0).Value)
                Dim newReqNo As String = rs.Fields.Item(0).Value.ToString()
                objMain.objApplication.StatusBar.SetText("Duplicate detected. NewItem List Number generated: " & newReqNo, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "Problem while processing Auto ID : " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
    Sub CFLFilter(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim Selection As String = "select ""Code"",""Name"",""U_GLAcct"", ""U_Itmlst"" from ""@TNX_SDPT"" where ""U_Itmlst""='Y'"
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
                oCondition.Alias = "U_Itmlst"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = oRs.Fields.Item("U_Itmlst").Value
                oChooseFromList.SetConditions(oConditions)

            End If
        Catch ex As Exception

        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.ActiveForm
                Try
                    objForm.Freeze(True)
                    If TERM = "MtxDT" Then
                        Me.SetNewLine(objForm.UniqueID)
                    ElseIf TERM = "Mtx3" Then
                        Me.SetNewLine1(objForm.UniqueID)
                    End If
                    objForm.Freeze(False)
                Catch ex As Exception
                    objForm.Freeze(False)
                End Try

            ElseIf pVal.MenuUID = "Delete Row" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.ActiveForm
                objMatrix = objForm.Items.Item("MtxDT").Specific
                For i As Integer = 1 To objMatrix.VisualRowCount - 1
                    If objMatrix.IsRowSelected(i) = True Then
                        objMatrix.DeleteRow(i)
                    End If
                Next
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.Columns.Item("V_-1").Cells.Item(i).Specific.string = i
                Next
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            ElseIf pVal.MenuUID = "519" And pVal.BeforeAction = True Then
                Try
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_LIST")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_LIST_C0")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_LIST_C1")
                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_LIST_C2")

                    objMatrix = objForm.Items.Item("MtxDT").Specific
                    'objMatrix1 = objForm.Items.Item("Mtx2").Specific
                    objMatrix2 = objForm.Items.Item("Mtx3").Specific



                    Dim frs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'ITEMLIST' AND ""Type"" = 'C'")

                    If frs.RecordCount = 0 Then
                        objMain.objApplication.MessageBox("ITEMLIST Sheet Layout('ITEMLIST') not found. Please upload with same name and Code as parameter", 0, "OK")
                    Else
                        objMain.objApplication.ActivateMenuItem(frs.Fields.Item(0).Value.ToString())
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_LIST")
                        Dim CrForm As SAPbouiCOM.Form
                        Dim oedt As SAPbouiCOM.EditText
                        CrForm = objMain.objApplication.Forms.ActiveForm
                        oedt = CrForm.Items.Item("1000003").Specific
                        oedt.Value = oDBs_Head.GetValue("DocEntry", 0)
                        CrForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If


                Catch ex As Exception
                End Try



            ElseIf pVal.MenuUID = "520" And pVal.BeforeAction = True Then



                Dim LayoutSelection As Integer = objMain.objApplication.MessageBox("Please select the layout for printing", 1, "LC Process", "")
                If (LayoutSelection = 1) Then

                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_LIST")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_LIST_C0")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_LIST_C1")
                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_LIST_C2")

                    objMatrix = objForm.Items.Item("MtxDT").Specific
                    'objMatrix1 = objForm.Items.Item("Mtx2").Specific
                    objMatrix2 = objForm.Items.Item("Mtx3").Specific




                    Dim frs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'ITEMLIST' AND ""Type"" = 'C'")

                    If frs.RecordCount = 0 Then
                        objMain.objApplication.MessageBox("ITEMLIST Sheet Layout('ITEMLIST') not found. Please upload with same name and Code as parameter", 0, "OK")
                    Else
                        objMain.objApplication.ActivateMenuItem(frs.Fields.Item(0).Value.ToString())
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_LIST")
                        Dim CrForm As SAPbouiCOM.Form
                        Dim oedt As SAPbouiCOM.EditText
                        CrForm = objMain.objApplication.Forms.ActiveForm
                        oedt = CrForm.Items.Item("1000003").Specific
                        oedt.Value = oDBs_Head.GetValue("DocEntry", 0)
                        CrForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If

                End If


            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub PostJE()
        oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_LIST")
        oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_LIST_C0")
        objMatrix = objForm.Items.Item("MtxDT").Specific
        For i As Integer = 1 To objMatrix.VisualRowCount
            Dim supplier As String = oDBs_Head.GetValue("U_CardCode", 0).Trim
            If objMatrix.Columns.Item("JEPosted").Cells.Item(i).Specific.Value.ToString().Trim = "" And CDbl(objMatrix.Columns.Item("Amount").Cells.Item(i).Specific.Value.ToString().Trim) <> 0 Then
                Dim Query As String = "SELECT ""U_REG"" FROM ""@ME_ACFG"" "
                Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rs.DoQuery(Query)
                Dim DrAccount As String = rs.Fields.Item(0).Value
                Dim Amount As Double = CDbl(objMatrix.Columns.Item("Amount").Cells.Item(i).Specific.Value.ToString().Trim)
                Dim Remarks As String = objMatrix.Columns.Item("Remarks").Cells.Item(i).Specific.Value.ToString().Trim
                Dim objComboBoxBranch As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("BRANCH").Cells.Item(i).Specific 'objForm.Items.Item("14").Specific
                If objComboBoxBranch.Value = "" Then
                    objMain.objApplication.StatusBar.SetText("Please select the branch.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
                Dim oJE As SAPbobsCOM.JournalEntries
                oJE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                oJE.ReferenceDate = Now
                oJE.DueDate = Now
                oJE.TaxDate = Now
                oJE.UserFields.Fields.Item("U_Remarks").Value = Remarks
                oJE.Reference = "Items List"
                oJE.Memo = "Listing amount"
                oJE.Lines.BPLID = objComboBoxBranch.Value
                oJE.Lines.ShortName = supplier
                oJE.Lines.Debit = Amount
                oJE.Lines.Add()

                oJE.Lines.AccountCode = DrAccount
                oJE.Lines.Credit = Amount
                oJE.Lines.BPLID = objComboBoxBranch.Value
                oJE.Lines.Add()

                If oJE.Add() = 0 Then
                    objMain.objApplication.StatusBar.SetText("Registration JE Posted Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Dim Query1 As String = "Update ""@SBO_LIST_C0"" Set ""U_JEPosted""=(Select Max(""TransId"") FROM OJDT) Where ""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0) & "' and ""LineId""='" & i & "'"
                    rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    rs.DoQuery(Query1)
                    'objMain.objApplication.ActivateMenuItem("1304")
                Else
                    objMain.objApplication.StatusBar.SetText("Registration JE Posting failed" + objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
                ' objMain.objApplication.ActivateMenuItem("1304")

                Try
                    objMain.objApplication.ActivateMenuItem("1304")
                Catch ex As Exception

                End Try
            End If
        Next
    End Sub
    Private Sub PostJEWithVATcode()
        Try
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_LIST")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_LIST_C2")
            objMatrix2 = objForm.Items.Item("Mtx3").Specific
            For i As Integer = 1 To objMatrix2.VisualRowCount
                Dim supplier As String = oDBs_Head.GetValue("U_CardCode", 0).Trim
                If objMatrix2.Columns.Item("JE").Cells.Item(i).Specific.Value.ToString().Trim = "" And CDbl(objMatrix2.Columns.Item("AMT").Cells.Item(i).Specific.Value.ToString().Trim) <> 0 Then

                    Dim Query As String = "SELECT ""U_VAT"",""U_VATACC"", ""U_DESC"" FROM ""@SBO_LIST_C2"""
                    Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    rs.DoQuery(Query)
                    Dim Description As String = objMatrix2.Columns.Item("DESC").Cells.Item(i).Specific.Value

                    Dim Query2 As String = "SELECT ""Code"",""Name"", ""U_GLAcct"", ""U_VATACC"" FROM ""@TNX_SDPT"" where ""U_GLAcct"" <>'' And ""Name"" = '" & Description & "'"
                    Dim rs1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    rs1.DoQuery(Query2)


                    If rs.RecordCount > 0 Then
                        Dim Amount As Double = CDbl(objMatrix2.Columns.Item("AMT").Cells.Item(i).Specific.Value.ToString().Trim)
                        Dim VatCode As String = objMatrix2.Columns.Item("VATCODE").Cells.Item(i).Specific.Value
                        Dim VatAmount As Double = CDbl(objMatrix2.Columns.Item("VATAMT").Cells.Item(i).Specific.Value)

                        Dim DrAccount As String = rs1.Fields.Item("U_GLAcct").Value
                        Dim DrAccount1 As String = rs1.Fields.Item("U_VATACC").Value

                        Dim Remarks As String = objMatrix2.Columns.Item("REMARKS").Cells.Item(i).Specific.Value.ToString().Trim
                        Dim objComboBoxBranch As SAPbouiCOM.ComboBox = objMatrix2.Columns.Item("BRANCH").Cells.Item(i).Specific 'objForm.Items.Item("14").Specific
                        If objComboBoxBranch.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please select the branch.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Sub
                        End If
                        Dim oJE As SAPbobsCOM.JournalEntries
                        oJE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                        oJE.ReferenceDate = Now
                        oJE.DueDate = Now
                        oJE.TaxDate = Now
                        oJE.UserFields.Fields.Item("U_Remarks").Value = Remarks
                        oJE.Reference = "Items List"
                        oJE.Memo = "Listing amount"
                        oJE.AutoVAT = SAPbobsCOM.BoYesNoEnum.tYES
                        oJE.Lines.BPLID = objComboBoxBranch.Value
                        oJE.Lines.ShortName = supplier
                        oJE.Lines.Debit = Amount + VatAmount 'Pertotal
                        oJE.Lines.Add()

                        oJE.Lines.AccountCode = DrAccount
                        oJE.Lines.Credit = Amount
                        oJE.Lines.BPLID = objComboBoxBranch.Value
                        oJE.Lines.TaxGroup = VatCode
                        oJE.Lines.Add()

                        'oJE.Lines.AccountCode = DrAccount1
                        'oJE.Lines.Credit = Pertotal
                        'oJE.Lines.BPLID = objComboBoxBranch.Value
                        'oJE.Lines.Add()


                        If oJE.Add() = 0 Then
                            objMain.objApplication.StatusBar.SetText("Registration JE Posted Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Dim Query1 As String = " UPDATE ""@SBO_LIST_C2"" SET ""U_JE""=(Select Max(""TransId"") From OJDT Where ""TransType"" = '30'), ""U_JESTATUS""='Posted' WHERE ""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0) & "' and ""LineId""='" & i & "'"

                            'Dim Query1 As String = "Update ""@SBO_LIST_C2"" Set ""U_JE""=(Select Max(""TransId"") FROM OJDT) Where ""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0) & "' and ""LineId""='" & i & "'"
                            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            rs.DoQuery(Query1)

                        Else
                            objMain.objApplication.StatusBar.SetText("Registration JE Posting failed" + objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If

                        Try
                            objMain.objApplication.ActivateMenuItem("1304")
                        Catch ex As Exception

                        End Try
                    Else
                        objMain.objApplication.StatusBar.SetText("Please Select the VatCode in Supplier Charges.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If
                End If
            Next


        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("ME_ITMLIST", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        objMatrix1 = objForm.Items.Item("MtxDT").Specific
                        Dim BaseNo As String = objForm.Items.Item("11").Specific.Value

                        For i As Integer = 1 To objMatrix1.VisualRowCount
                            Dim oRs As SAPbobsCOM.Recordset
                            oRs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            Dim sSQL As String

                            sSQL = "SELECT C.""U_JE"" FROM ""@SBO_OUTLET"" H  INNER JOIN ""@SBO_OUTLET_C0"" C ON H.""DocEntry"" = C.""DocEntry"" " &
       "WHERE H.""U_BaseNo"" = '" & BaseNo & "' AND H.""U_LineNum"" = '" & i & "' AND IFNULL(C.""U_JE"", '') <> '' GROUP BY H.""DocEntry"", H.""U_CardCode"", C.""DocEntry"", H.""U_LineNum"", C.""U_JE"" " &
       "HAVING COUNT(C.""U_JE"") > 0"

                            oRs.DoQuery(sSQL)

                            If (objMatrix1.Columns.Item("JEPosted").Cells.Item(i).Specific.Value.ToString() <> "" Or oRs.RecordCount > 0) Then
                                objMatrix1.CommonSetting.SetRowEditable(i, False)
                            Else
                                objMatrix1.CommonSetting.SetCellEditable(i, 1, True)
                                objMatrix1.CommonSetting.SetCellEditable(i, 2, True)
                                objMatrix1.CommonSetting.SetCellEditable(i, 3, True)
                                objMatrix1.CommonSetting.SetCellEditable(i, 4, True)
                                objMatrix1.CommonSetting.SetCellEditable(i, 5, True)
                                objMatrix1.CommonSetting.SetCellEditable(i, 7, True)
                            End If
                        Next
                        objMatrix2 = objForm.Items.Item("Mtx3").Specific
                        For i As Integer = 1 To objMatrix2.VisualRowCount

                            If objMatrix2.Columns.Item("JE").Cells.Item(i).Specific.Value.ToString() <> "" Then
                                objMatrix2.CommonSetting.SetRowEditable(i, False)
                            Else
                                objMatrix2.CommonSetting.SetCellEditable(i, 1, True)
                                objMatrix2.CommonSetting.SetCellEditable(i, 2, True)
                                objMatrix2.CommonSetting.SetCellEditable(i, 3, True)
                                objMatrix2.CommonSetting.SetCellEditable(i, 5, True)
                                objMatrix2.CommonSetting.SetCellEditable(i, 6, True)
                            End If
                        Next


                    End If
                    '-------Madhu

                    ' After Approving documebt should be totally disbaled code here ...!
                    'Dim statusItemListing As String = objForm.Items.Item("Statu").Specific.Value.ToString().Trim().ToUpper()

                    'If statusItemListing = "APPROVED" OrElse statusItemListing = "A" OrElse statusItemListing = "SUBMITTED" OrElse statusItemListing = "S" OrElse statusItemListing = "REJECTED" OrElse statusItemListing = "R" Then

                    '    Try
                    '        If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                    '            objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                    '        End If
                    '    Catch
                    '    End Try

                    '    For i As Integer = 0 To objForm.Items.Count - 1

                    '        Try

                    '            Dim oItem As SAPbouiCOM.Item = objForm.Items.Item(i)
                    '            If oItem.UniqueID <> "MtxDT" AndAlso
                    '               oItem.UniqueID <> "Mtx3" AndAlso
                    '               oItem.UniqueID <> "1" AndAlso
                    '               oItem.UniqueID <> "2" AndAlso
                    '               oItem.UniqueID <> "Submit" AndAlso
                    '               oItem.Type <> SAPbouiCOM.BoFormItemTypes.it_FOLDER AndAlso
                    '               oItem.Type <> SAPbouiCOM.BoFormItemTypes.it_STATIC Then

                    '                oItem.Enabled = False

                    '            End If

                    '        Catch
                    '        End Try

                    '    Next

                    '    Try : objForm.Items.Item("1").Enabled = True : Catch : End Try
                    '    Try : objForm.Items.Item("2").Enabled = True : Catch : End Try

                    'Else

                    '    Try
                    '        objForm.Items.Item("Submit").Enabled = True
                    '    Catch
                    '    End Try

                    '    Try
                    '        If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    '            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                    '        End If
                    '    Catch
                    '    End Try

                    'End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        objForm = objMain.objApplication.Forms.GetForm("ME_ITMLIST", objMain.objApplication.Forms.ActiveForm.TypeCount)

                        'If Not Validation(objForm.UniqueID) Then
                        '    BubbleEvent = False
                        'Else

                        Me.UpdateBP(objForm.UniqueID)
                        '  End If
                    End If


            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub UpdateBP(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim oBusinessPartners As SAPbobsCOM.BusinessPartners = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            If oBusinessPartners.GetByKey(objForm.Items.Item("CardCode").Specific.Value) Then
                Dim test As String = objForm.Items.Item("11").Specific.Value
                oBusinessPartners.UserFields.Fields.Item("U_ITMLIST").Value = test
                oBusinessPartners.Update()
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

#Region " Right Click Event"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Dim objForm As SAPbouiCOM.Form
        objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)

        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_LIST")
        oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_LIST_C0")
        oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_LIST_C2")

        objMatrix1 = objForm.Items.Item("MtxDT").Specific
        objMatrix2 = objForm.Items.Item("Mtx3").Specific


        Try
            If eventInfo.FormUID = objForm.UniqueID Then
                If (eventInfo.BeforeAction = True) Then
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then 'SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                        '  If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Or objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then

                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then 'SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                            '  objMatrix = objForm.Items.Item("Matx").Specific
                            ' objForm.EnableMenu("1292", True)
                            ' objForm.EnableMenu("1293", True)
                            If eventInfo.ItemUID = "MtxDT" And eventInfo.ColUID = "LineId" Then
                                Try

                                Catch ex As Exception
                                    objMain.objApplication.StatusBar.SetText(ex.Message)
                                End Try
                            ElseIf eventInfo.ItemUID = "Mtx3" And eventInfo.ColUID = "LineId" Then
                                Try

                                Catch ex As Exception
                                    objMain.objApplication.StatusBar.SetText(ex.Message)
                                End Try
                            End If

                        End If
                    End If
                End If
            End If
            ' System.Diagnostics.Process.Start()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region
    Private Sub PostJENew(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_LIST")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_LIST_C2")
            objMatrix = objForm.Items.Item("Mtx3").Specific
            If oDBs_Head.GetValue("U_BRANCH", 0) = "" Then
                objMain.objApplication.StatusBar.SetText("Please select the branch.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            For i As Integer = 1 To objMatrix.VisualRowCount
                If CDbl(objMatrix.Columns.Item("AMT").Cells.Item(i).Specific.Value) <> 0.0 And objMatrix.Columns.Item("JE").Cells.Item(i).Specific.Value = "" And objMatrix.Columns.Item("DESC").Cells.Item(i).Specific.Value <> "" Then
                    Dim oJE As SAPbobsCOM.JournalEntries = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                    Dim GetAddonConfigurationDtls As String = "Select ""U_REG"" From ""@ME_ACFG"" "
                    Dim oRsGetAddonConfigurationDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetAddonConfigurationDtls.DoQuery(GetAddonConfigurationDtls)

                    Dim sDate As String = DateTime.Now.ToString("yyyyMMdd")
                    sDate = sDate.Insert("4", "-")
                    sDate = sDate.Insert("7", "-")

                    oJE.ReferenceDate = sDate
                    oJE.DueDate = sDate
                    oJE.TaxDate = sDate

                    oJE.Reference2 = "Item Listing"
                    Dim test As String = oDBs_Head.GetValue("U_BRANCH", 0)
                    oJE.Lines.BPLID = test
                    oJE.Lines.AccountCode = oRsGetAddonConfigurationDtls.Fields.Item("U_REG").Value
                    oJE.Lines.Credit = objMatrix.Columns.Item("AMT").Cells.Item(i).Specific.Value
                    oJE.Lines.Add()

                    oJE.Lines.BPLID = oDBs_Head.GetValue("U_Branch", 0)
                    oJE.Lines.ShortName = objForm.Items.Item("CardCode").Specific.value
                    oJE.Lines.Debit = objMatrix.Columns.Item("AMT").Cells.Item(i).Specific.Value

                    oJE.Lines.Add()

                    If oJE.Add = 0 Then

                        Dim GetJEDocEntry As String = "Select Max(""TransId"") From OJDT Where ""TransType"" = '30'"
                        Dim oRsGetJEDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetJEDocEntry.DoQuery(GetJEDocEntry)

                        Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim Update As String = " UPDATE ""@SBO_LIST_C2"" SET ""U_JE""='" & oRsGetJEDocEntry.Fields.Item(0).Value & "', ""U_JESTATUS""='Posted' WHERE ""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0) & "' and ""LineId""='" & i & "'"
                        oRsUpdate.DoQuery(Update)

                        objForm.Refresh()
                    Else
                        objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorCode & " : " & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                End If
            Next
            Try
                objMain.objApplication.ActivateMenuItem("1304")
            Catch ex As Exception

            End Try


        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    'madhu
    Private Function GetBranchCount() As Integer
        Try
            Dim oRs As SAPbobsCOM.Recordset =
            objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oRs.DoQuery("SELECT COUNT(*) AS CNT FROM OBPL")

            If oRs.RecordCount > 0 Then
                Return CInt(oRs.Fields.Item("CNT").Value)
            End If

        Catch ex As Exception
        End Try

        Return 0
    End Function

    Private Function ValidateNoOfSkus(ByVal FormUID As String) As Boolean
        Try
            Dim objForm As SAPbouiCOM.Form =
            objMain.objApplication.Forms.Item(FormUID)

            Dim objMatrix As SAPbouiCOM.Matrix =
            CType(objForm.Items.Item("MtxDT").Specific, SAPbouiCOM.Matrix)

            ' Force commit of active cell
            objMatrix.FlushToDataSource()

            Dim MaxBranches As Integer = GetBranchCount()

            Dim oDBs_Details As SAPbouiCOM.DBDataSource =
            objForm.DataSources.DBDataSources.Item("@SBO_LIST_C0")

            For i As Integer = 0 To objMatrix.VisualRowCount - 1

                Dim NouStr As String =
                oDBs_Details.GetValue("U_NOU", i).Trim()

                If NouStr = "" Then Continue For

                Dim NouValue As Integer
                If Not Integer.TryParse(NouStr, NouValue) Then Continue For

                If NouValue > MaxBranches Then

                    objMain.objApplication.StatusBar.SetText(
                    $"NOU cannot be greater than available branches ({MaxBranches}). Line : {i + 1}",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    ''objMatrix.ClearSelection()
                    'objMatrix.Columns.Item("NOU").Cells.Item(i + 1).Click(
                    'SAPbouiCOM.BoCellClickType.ct_Regular)

                    Return False
                End If
            Next

            Return True

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Private Function Validation(ByVal FormUID As String) As Boolean
        Try
            Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUID)
            Dim oMatrix As SAPbouiCOM.Matrix = CType(oForm.Items.Item("MtxDT").Specific, SAPbouiCOM.Matrix)

            'Header validation
            Dim CardCode As String = CType(oForm.Items.Item("CardCode").Specific, SAPbouiCOM.EditText).Value.Trim
            If CardCode = "" Then
                objMain.objApplication.StatusBar.SetText("Card Code is mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            'Matrix row validation
            If oMatrix.RowCount = 0 Then
                objMain.objApplication.StatusBar.SetText("Please add at least one row", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            For i As Integer = 1 To oMatrix.RowCount

                Dim DateVal As String = CType(oMatrix.Columns.Item("Date").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value.Trim
                Dim SKU As String = CType(oMatrix.Columns.Item("SKU").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value.Trim
                Dim Branch As String = CType(oMatrix.Columns.Item("BRANCH").Cells.Item(i).Specific, SAPbouiCOM.ComboBox).Value.Trim
                Dim NOU As String = CType(oMatrix.Columns.Item("NOU").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value.Trim

                If DateVal = "" Then
                    objMain.objApplication.StatusBar.SetText("Date is required in row " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                If SKU = "" OrElse Val(SKU) <= 0 Then
                    objMain.objApplication.StatusBar.SetText(" SKU Cannot be Empty in row " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                'If Branch = "" Then
                '    objMain.objApplication.StatusBar.SetText("Branch is required in row " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    Return False
                'End If

                If NOU = "" OrElse Val(NOU) <= 0 Then
                    objMain.objApplication.StatusBar.SetText("Number of Outlets must be greater than 0 in row " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

            Next

            Return True

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    ' Get Charge from OCRD.U_LCPSKU based on CardCode

    Private Sub LoadChargeFromOCRD(ByVal CardCode As String)
        Try
            Dim oRs As SAPbobsCOM.Recordset =
            objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim sQry As String =
            "SELECT ""U_LCPSKU"" FROM ""OCRD"" WHERE ""CardCode"" = '" & CardCode & "'"

            oRs.DoQuery(sQry)

            If oRs.RecordCount > 0 Then
                Double.TryParse(oRs.Fields.Item("U_LCPSKU").Value.ToString(), Me.Chargevalue)
            End If

            Dim oMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("MtxDT").Specific
            Dim oDBs_Details As SAPbouiCOM.DBDataSource =
            objForm.DataSources.DBDataSources.Item("@SBO_LIST_C0")

            For i As Integer = 1 To oMatrix.VisualRowCount
                oDBs_Details.SetValue("U_Charge", i - 1, Me.Chargevalue.ToString())
            Next

            oMatrix.LoadFromDataSource()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "Failed to load LCP SKU Charge: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub



    'Application Id
    Public Function AutoDocentryNumber()

        'Dim oRsDocNum As SAPbobsCOM.Recordset = objForm.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objForm = objMain.objApplication.Forms.Item(objForm.UniqueID)
        oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_LIST")
        Dim oRsDocNum As SAPbobsCOM.Recordset

        oRsDocNum = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim Query1 As String = "SELECT ifnull(Max(""DocNum""),0)+1 AS ""DocNum"" FROM ""@SBO_LIST"" "

        oRsDocNum.DoQuery(Query1)

        oDBs_Head.SetValue("DocNum", 0, oRsDocNum.Fields.Item(0).Value)

        Dim rsAppId As SAPbobsCOM.Recordset

        rsAppId = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim str As String = "SELECT 'IL' || LPAD('',6 - LENGTH(LTRIM(RTRIM(REPLACE(IFNULL(MAX(""U_AppID""), '0'), 'IL', '') + 1))),'0') || LTRIM(RTRIM(REPLACE(IFNULL(MAX(""U_AppID""), '0'), 'IL', '') + 1)) FROM ""@SBO_LIST"";"

        rsAppId.DoQuery(str)
        Dim appp As String = rsAppId.Fields.Item(0).Value.ToString().Trim()

        oDBs_Head.SetValue("U_AppID", oDBs_Head.Offset, appp)


    End Function


    'Approval

    Sub ApprovalTrigger(ByVal DocType As String, ByVal AppID As String, ByVal Table As String, ByVal AppIDField As String, ByVal AppStatField As String, ByVal Document As String)
        Dim ff As Integer
        Dim draftquery As String
        Dim approverid As String
        'Dim docdate As String = objForm.Items.Item("DDATE").Specific.Value

        ff = objMain.objApplication.MessageBox("Do You want Submit the Application?", 1, "Ok", "Cancel")
        If ff = 1 Then
            Dim result As Integer
            result = objMain.objApplication.MessageBox("You can not change the Application after you Sending for Approval,continue", 1, "Yes", "No")
            If result = 1 Then
                Dim rsapp As SAPbobsCOM.Recordset
                Dim posrs As SAPbobsCOM.Recordset
                rsapp = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim str As String = "SELECT T0.""Code"" As ""TemplateID"",T1.""U_Name"" As ""Originator"",S1.""Code"" As ""Stage"",S2.""U_UKey"" As ""Authorizers"" " &
                " ,S1.""U_NAP"" As ""NoOfAuth"",S1.""U_NRJ"" As ""NoOfRej"",T0.""CreateDate"",(Case When T3.""LineId""='1' Then 'S' Else 'O' End) As ""Status"" FROM ""@SBO_APPHDR"" T0 " &
                " INNER JOIN ""@SBO_APPREQ"" T1 ON T1.""Code""=T0.""Code"" " &
                " INNER JOIN ""@SBO_APPDOC"" T2 ON T2.""Code"" = T0.""Code"" " &
                " INNER JOIN ""@SBO_APPAUT"" T3 ON T3.""Code"" = T0.""Code"" " &
                " INNER JOIN ""@SBO_AST"" S1 ON T3.""U_M3_1""=S1.""Code"" " &
                " INNER JOIN ""@SBO_AST_C0"" S2 ON S1.""Code""=S2.""Code"" " &
                " WHERE ""U_Active"" = 'Y' AND T1.""U_Name"" = '" & objMain.objCompany.UserName & "' AND T2.""U_" & DocType & """  = 'Y' AND IFNULL(S2.""U_UKey"",'')<>''  order by T3.""LineId""  ;"
                rsapp.DoQuery(str)


                If rsapp.RecordCount > 0 Then
                    posrs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    'draftquery = " SELECT * FROM ""@SBO_OWDD"" WHERE ""U_statusN"" IN ('S','A') AND ""U_Isdraft"" IN ('Y','N') AND ""U_docnum"" = '" & AppID & "';"
                    draftquery = " SELECT * FROM ""@SBO_DD"" WHERE ""U_statusN"" IN ('S','A') AND ""U_Isdraft"" IN ('Y','N') AND ""U_AppID"" = '" & AppID & "';"

                    posrs.DoQuery(draftquery)
                    If posrs.RecordCount > 0 Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(posrs)
                        objMain.objApplication.StatusBar.SetText("Document Already sent for approval", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    Else
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(posrs)
                        ' Dim ofor As SAPbouiCOM.Form
                        'Me.SendMessageAlert(DocType, objMain.objCompany.UserName)

                        'ofor = New DraftProcedureStages(DocType, AppID, objMain.objCompany.UserName, rsapp, Table, AppIDField, AppStatField, Document, docdate)
                        ' ofor = Nothing
                        Dim objBaseForm As SAPbouiCOM.Form = objMain.objApplication.Forms.ActiveForm
                        objMain.ObjDraftProcedure.CreateForm(objBaseForm, DocType, AppID, objMain.objCompany.UserName, rsapp, Table, AppIDField, AppStatField, Document)
                        '.CreateForm(objForm.UniqueID, objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value)
                        'objMain.objApplication.ActivateMenuItem("1304")
                    End If
                    'objForm.Close()
                Else
                    objMain.objApplication.StatusBar.SetText("Login User Don't have Approvals Please Define Approvals and Add the Document", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        End If
    End Sub


    Public Sub HandleLayoutKeyEvent1(ByRef eventInfo As SAPbouiCOM.LayoutKeyInfo, ByRef BubbleEvent As Boolean)

        Try

            If eventInfo.BeforeAction = False Then
                Dim f As SAPbouiCOM.Form = Nothing
                Try
                    f = objMain.objApplication.Forms.Item(eventInfo.FormUID)
                Catch
                    f = Nothing
                End Try

                If f Is Nothing Then Return

                If f.TypeEx = "ME_ITMLIST" OrElse f.TypeEx = "ME_ITMLIST" Then
                    Try
                        'Dim txtObj As SAPbouiCOM.EditText = CType(f.Items.Item("DocEntry").Specific, SAPbouiCOM.EditText)
                        'Dim docVal As String = If(txtObj IsNot Nothing, txtObj.Value.Trim(), "")
                        Dim docVal As String
                        If String.IsNullOrWhiteSpace(docVal) Then
                            ' fallback: try header datasource DocNum
                            Try
                                Dim ds As SAPbouiCOM.DBDataSource = f.DataSources.DBDataSources.Item("@SBO_LIST")
                                docVal = ds.GetValue("DocEntry", 0).ToString().Trim()
                            Catch
                                docVal = ""
                            End Try
                        End If

                        If Not String.IsNullOrWhiteSpace(docVal) Then
                            eventInfo.LayoutKey = docVal
                        End If
                    Catch ex As Exception
                        ' ignore - cannot supply layout key
                    End Try
                End If
            End If
        Catch ex As Exception
            Try
                objMain.objApplication.StatusBar.SetText("LayoutKey handler error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Catch
            End Try
        End Try

    End Sub


    ' 'layout printing Item List code enabling
    Public Sub CreateLayoutFolder_ItemList1()
        Try
            ' Use the connected company object from MainCls (objMain.objCompany). If not available, abort with a clear message.
            Dim companyServiceSource As SAPbobsCOM.Company = Nothing
            Try
                If objMain IsNot Nothing AndAlso objMain.objCompany IsNot Nothing Then
                    companyServiceSource = objMain.objCompany
                End If
            Catch
                companyServiceSource = Nothing
            End Try

            If companyServiceSource Is Nothing Then
                objMain.objApplication.StatusBar.SetText("CreateLayoutFolder_ME_ITMLIST aborted: not connected to a company.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return
            End If

            Dim oReportTypeService As SAPbobsCOM.ReportTypesService
            oReportTypeService = companyServiceSource.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService)

            Dim oList = oReportTypeService.GetReportTypeList()

            For i As Integer = 0 To oList.Count - 1
                If oList.Item(i).AddonFormType = "ME_ITMLIST" Then
                    Exit Sub   ' Already exists
                End If
            Next

            Dim oReportType As SAPbobsCOM.ReportType
            oReportType = oReportTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportType)

            oReportType.TypeName = "ME_ITMLIST Print"
            oReportType.AddonName = "Payroll Addon"
            oReportType.AddonFormType = "ME_ITMLIST"
            oReportType.MenuID = "ME_ITMLIST_MENU"

            oReportTypeService.AddReportType(oReportType)

        Catch ex As Exception
            Try
                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Catch
            End Try
        End Try
    End Sub


    Public Sub HandleLayoutKeyEvent(ByRef eventInfo As SAPbouiCOM.LayoutKeyInfo, ByRef BubbleEvent As Boolean)

        Try

            If eventInfo.BeforeAction = False Then
                Dim f As SAPbouiCOM.Form = Nothing

                Try
                    f = objMain.objApplication.Forms.Item(eventInfo.FormUID)
                Catch
                    f = Nothing
                End Try

                If f Is Nothing Then Return

                If f.TypeEx = "ME_ITMLIST" Then

                    Try
                        Dim docVal As String = ""

                        ' Try getting from EditText (if exists)
                        Try
                            Dim txtObj As SAPbouiCOM.EditText =
                        CType(f.Items.Item("DocEntry").Specific, SAPbouiCOM.EditText)

                            If txtObj IsNot Nothing Then
                                docVal = txtObj.Value.Trim()
                            End If
                        Catch
                            docVal = ""
                        End Try

                        ' fallback: try header datasource DocEntry
                        If String.IsNullOrWhiteSpace(docVal) Then
                            Try
                                Dim ds As SAPbouiCOM.DBDataSource =
                            f.DataSources.DBDataSources.Item("@SBO_LIST")

                                docVal = ds.GetValue("DocEntry", 0).ToString().Trim()
                            Catch
                                docVal = ""
                            End Try
                        End If

                        If Not String.IsNullOrWhiteSpace(docVal) Then
                            eventInfo.LayoutKey = docVal
                        End If

                    Catch
                        ' ignore
                    End Try

                End If

            End If

        Catch ex As Exception
            Try
                objMain.objApplication.StatusBar.SetText(
            "LayoutKey handler error: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Catch
            End Try
        End Try

    End Sub

    Public Sub CreateLayoutFolder_ItemList()

        Try

            Dim companyServiceSource As SAPbobsCOM.Company = Nothing

            Try
                If objMain IsNot Nothing AndAlso objMain.objCompany IsNot Nothing Then
                    companyServiceSource = objMain.objCompany
                End If
            Catch
                companyServiceSource = Nothing
            End Try

            If companyServiceSource Is Nothing Then
                objMain.objApplication.StatusBar.SetText(
                "CreateLayoutFolder_ME_ITMLIST aborted: not connected to a company.",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return
            End If


            Dim oReportTypeService As SAPbobsCOM.ReportTypesService
            oReportTypeService = companyServiceSource.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService)


            ' Correct Type
            Dim oList As SAPbobsCOM.ReportTypesParams
            oList = oReportTypeService.GetReportTypeList()


            ' Check if already exists
            For i As Integer = 0 To oList.Count - 1

                If oList.Item(i).AddonFormType = "ME_ITMLIST" Then
                    Exit Sub

                End If

            Next


            Dim oReportType As SAPbobsCOM.ReportType
            oReportType = oReportTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportType)

            oReportType.TypeName = "ME_ITMLIST Print"
            oReportType.AddonName = "Payroll Addon"

            ' Important
            oReportType.AddonFormType = "ME_ITMLIST"

            ' Must match form type so default layout persists
            oReportType.MenuID = "ME_ITMLIST"

            oReportTypeService.AddReportType(oReportType)


            objMain.objApplication.StatusBar.SetText(
            "Item List Print Layout Folder Created",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception

            Try
                objMain.objApplication.StatusBar.SetText(
                ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Catch
            End Try

        End Try

    End Sub


End Class
