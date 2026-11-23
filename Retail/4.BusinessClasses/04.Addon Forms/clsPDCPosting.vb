Public Class clsPDCPosting

#Region "       Declaration             "
    Public objForm, objSTRForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details1 As SAPbouiCOM.DBDataSource
    Dim objMatrix1 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("PDCPosting.xml", "ME_PDC_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_PDC_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PDC")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_PDC_C0")
            objMatrix1 = objForm.Items.Item("Mtx").Specific
            objForm.DataBrowser.BrowseBy = "17"

            objMatrix1.Columns.Item("ChequeSum").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

            Dim sQury As String = " Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join " & _
                                  " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""  "
            objComboBox = objForm.Items.Item("19").Specific
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, sQury)
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("19").DisplayDesc = True

            If objComboBox.ValidValues.Count <> 0 Then
                objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)              
            End If

            objForm.Items.Item("27").ForeColor = ColorTranslator.ToOle(Color.DarkGreen)
            objForm.Items.Item("27").FontSize = 16
            objForm.Items.Item("27").TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD

            objComboBox = objForm.Items.Item("21").Specific
            objForm.Items.Item("21").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objComboBox = objForm.Items.Item("23").Specific
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("23").DisplayDesc = True
            Me.CellsMasking(objForm.UniqueID)
            Me.CFLSupplierFilter("CFL_SUPCD")
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub CellsMasking(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Items.Item("16").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("16").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("16").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("Item_13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("Item_13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("Item_13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("17").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("17").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)        

            objForm.Items.Item("PostJE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("PostJE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("PostJE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ME_PDC" And pVal.BeforeAction = False Then
                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.ActiveForm

                Me.SetDefault(objForm.UniqueID)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PDC")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_PDC_C0")
                    objMatrix1 = objForm.Items.Item("Mtx").Specific
                    If pVal.ItemUID = "26" And pVal.BeforeAction = False Then
                        Me.TotalChequeAmt(objForm.UniqueID)
                    End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PDC")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_PDC_C0")
                    objMatrix1 = objForm.Items.Item("Mtx").Specific

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim LoadSameForm As String = "Select MAX(""DocNum"") FROM ""@ME_PDC"""
                        Dim oRsLoadSameForm As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsLoadSameForm.DoQuery(LoadSameForm)
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objForm.Items.Item("17").Specific.Value = oRsLoadSameForm.Fields.Item(0).Value
                        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.Validation(objForm.UniqueID) = False Then BubbleEvent = False
                    End If

                    If pVal.ItemUID = "PostJE" And pVal.BeforeAction = False Then
                        Me.DocumentPosting(objForm.UniqueID)
                    End If

                    If pVal.ItemUID = "Item_13" And pVal.BeforeAction = False Then
                        objMatrix1 = objForm.Items.Item("Mtx").Specific

                        Dim DateFrom As String = oDBs_Head.GetValue("U_DateFrom", 0).ToString().Trim
                        Dim DateTo As String = oDBs_Head.GetValue("U_DateTo", 0).ToString().Trim
                        Dim Branch As String = oDBs_Head.GetValue("U_BRANCH", 0).ToString().Trim
                        Dim Issue As String = oDBs_Head.GetValue("U_Issued", 0).ToString().Trim
                        Dim Approved As String = oDBs_Head.GetValue("U_Approved", 0).ToString().Trim
                        Dim SupplierCode As String = oDBs_Head.GetValue("U_SUPCODE", 0).ToString().Trim
                        Dim ChequeNo As String = oDBs_Head.GetValue("U_ChequeNo", 0).ToString().Trim
                        Dim GLAccount As String = oDBs_Head.GetValue("U_GLAcc", 0).ToString().Trim
                        Dim Amount As String = oDBs_Head.GetValue("U_ChequeSm", 0).ToString().Trim

                        If DateFrom = "" Then
                            objMain.objApplication.StatusBar.SetText("Please enter From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        ElseIf DateTo = "" Then
                            objMain.objApplication.StatusBar.SetText("Please enter To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        ElseIf Branch = "" Then
                            objMain.objApplication.StatusBar.SetText("Please enter Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        ElseIf GLAccount = "" Then
                            objMain.objApplication.StatusBar.SetText("Please enter GL Account", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            Dim FromDate As String = objForm.Items.Item("From").Specific.Value.ToString().Trim()
                            Dim ToDate As String = objForm.Items.Item("To").Specific.Value.ToString().Trim()

                            Dim ChequeQuery As String = " select T2.""DocEntry"",T2.""DocDate"",T2.""CardCode"" ,T2.""CardName"" " &
                                                         " , T0.""DueDate"", T0.""CheckNum"", T0.""CheckSum"",T1.""AcctCode"",T1.""AcctName"", " &
                                                         " T1.""AccntntCod"" , T1.""U_DepositeA"" , T0.""U_Issued"", T0.""U_Approved""  , T0.""U_RECNM"" " &
                                                         " From VPM1 T0 inner join OACT T1 on T0.""CheckAct""=T1.""AcctCode"" Inner Join OVPM T2 On (T0.""DocNum""=T2.""DocEntry"") " &
                                                         " where T2.""Canceled"" ='N' and ifnull(T0.""U_Posted"",'N')='N' and Ifnull(T1.""U_PDCReq"",'')='Y' And T2.""BPLId"" = '" & Branch & "' And Not EXISTS(SELECT 1 FROM ""@ME_PDC_C0"" C " &
" WHERE TO_NVARCHAR(C.""U_DocEntry"") = TO_NVARCHAR(T2.""DocEntry"") " &
" And TO_NVARCHAR(C.""U_ChequeNo"") = TO_NVARCHAR(T0.""CheckNum"") " &
" And IFNULL(TRIM(C.""U_JeNo""),'') <> '') "
                            If FromDate <> "" And ToDate <> "" Then
                                ChequeQuery = ChequeQuery & "And T2.""DocDate"" Between '" & FromDate & "' And '" & ToDate & "'"
                            End If
                            If Issue <> "" And Issue <> "A" Then
                                ChequeQuery = ChequeQuery & "And T0.""U_Issued""= '" & Issue & "'"
                            End If
                            If Approved <> "" And Approved <> "A" Then
                                ChequeQuery = ChequeQuery & "And T0.""U_Approved""= '" & Approved & "'"
                            End If
                            If SupplierCode <> "" Then
                                ChequeQuery = ChequeQuery & "And T2.""CardCode""= '" & SupplierCode & "'"
                            End If
                            If ChequeNo <> "" Then
                                ChequeQuery = ChequeQuery & "And T0.""CheckNum""= '" & ChequeNo & "'"
                            End If
                            ' madhu Based on GL account loading the data 
                            If GLAccount <> "" Then
                                ChequeQuery = ChequeQuery & " And T1.""AcctCode"" = '" & GLAccount & "' "
                            End If
                            If Amount <> 0 Then
                                ChequeQuery = ChequeQuery & " And T2.""DocTotal"" = '" & Amount & "' "
                            End If
                            ChequeQuery = ChequeQuery & " group by T0.""DueDate"",T0.""CheckNum"", T0.""CheckSum"",T1.""AcctName"",T2.""DocEntry"",T2.""DocDate"",T2.""CardCode"",T2.""CardName"",T1.""AcctCode"",T1.""AccntntCod"" , T1.""U_DepositeA""  , T0.""U_Issued"" , T0.""U_Approved""  , T0.""U_RECNM"" "

                            Dim TransferQuery As String = " select T2.""DocEntry"", T2.""DocDate"", T2.""CardCode"", T2.""CardName"", " &
                                                          " ifnull(T2.""TrsfrDate"", T2.""DocDate"") As ""DueDate"", 0 As ""CheckNum"", T2.""TrsfrSum"" As ""CheckSum"", T1.""AcctCode"", T1.""AcctName"", " &
                                                          " T1.""AccntntCod"", T1.""U_DepositeA"", 'N' As ""U_Issued"", 'N' As ""U_Approved"", '' As ""U_RECNM"" " &
                                                          " From OVPM T2 Inner Join OACT T1 On T2.""TrsfrAcct"" = T1.""AcctCode"" " &
                                                          " where T2.""Canceled"" = 'N' and ifnull(T2.""TrsfrSum"", 0) > 0 And T2.""BPLId"" = '" & Branch & "' And Not EXISTS(SELECT 1 FROM ""@ME_PDC_C0"" C " &
" WHERE TO_NVARCHAR(C.""U_DocEntry"") = TO_NVARCHAR(T2.""DocEntry"") " &
" And TO_NVARCHAR(C.""U_ChequeNo"") = '0' " &
" AND IFNULL(TRIM(C.""U_JeNo""),'') <> '') "
                            If FromDate <> "" And ToDate <> "" Then
                                TransferQuery = TransferQuery & " And T2.""DocDate"" Between '" & FromDate & "' And '" & ToDate & "'"
                            End If
                            If SupplierCode <> "" Then
                                TransferQuery = TransferQuery & " And T2.""CardCode"" = '" & SupplierCode & "'"
                            End If
                            If GLAccount <> "" Then
                                TransferQuery = TransferQuery & " And T1.""AcctCode"" = '" & GLAccount & "' "
                            End If

                            If Amount <> 0 Then
                                TransferQuery = TransferQuery & " And T2.""DocTotal"" = '" & Amount & "' "
                            End If

                            ' If Cheque Number is entered, exclude bank transfers
                            If ChequeNo <> "" Then
                                TransferQuery &= " AND 1 = 0 "
                            End If

                            TransferQuery = TransferQuery & " group by T2.""DocEntry"", T2.""DocDate"", T2.""CardCode"", T2.""CardName"", ifnull(T2.""TrsfrDate"", T2.""DocDate""), T2.""TrsfrSum"", T1.""AcctCode"", T1.""AcctName"", T1.""AccntntCod"", T1.""U_DepositeA"",T2.""DocTotal"" "

                            Dim DetailsQuery As String = ChequeQuery & " UNION ALL " & TransferQuery
                            Dim rs As SAPbobsCOM.Recordset
                            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            rs.DoQuery(DetailsQuery)

                            objMatrix1.Clear()
                            oDBs_Details1.Clear()
                            objMatrix1.FlushToDataSource()
                            objMatrix1.AutoResizeColumns()
                            For i As Integer = 1 To rs.RecordCount
                                Dim DocDate As Date = CDate(rs.Fields.Item("DocDate").Value)
                                Dim DueDate As Date = CDate(rs.Fields.Item("DueDate").Value)
                                objMatrix1.AddRow()
                                oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, i)
                                oDBs_Details1.SetValue("U_DocEntry", oDBs_Details1.Offset, rs.Fields.Item("DocEntry").Value)
                                oDBs_Details1.SetValue("U_DocDate", oDBs_Details1.Offset, DocDate.ToString("yyyyMMdd"))
                                oDBs_Details1.SetValue("U_CardCode", oDBs_Details1.Offset, rs.Fields.Item("CardCode").Value)
                                oDBs_Details1.SetValue("U_CardName", oDBs_Details1.Offset, rs.Fields.Item("CardName").Value)
                                oDBs_Details1.SetValue("U_ChequeNo", oDBs_Details1.Offset, rs.Fields.Item("CheckNum").Value)
                                oDBs_Details1.SetValue("U_DueDate", oDBs_Details1.Offset, DueDate.ToString("yyyyMMdd"))
                                oDBs_Details1.SetValue("U_ChequeSm", oDBs_Details1.Offset, rs.Fields.Item("CheckSum").Value)
                                oDBs_Details1.SetValue("U_AcctNum", oDBs_Details1.Offset, rs.Fields.Item("AcctCode").Value)
                                oDBs_Details1.SetValue("U_AcctName", oDBs_Details1.Offset, rs.Fields.Item("AcctName").Value)
                                oDBs_Details1.SetValue("U_DepositA", oDBs_Details1.Offset, rs.Fields.Item("U_DepositeA").Value)
                                oDBs_Details1.SetValue("U_JeNo", oDBs_Details1.Offset, "")
                                oDBs_Details1.SetValue("U_Required", oDBs_Details1.Offset, "N")
                                If rs.Fields.Item("U_Issued").Value = "Y" Then
                                    oDBs_Details1.SetValue("U_Issued", oDBs_Details1.Offset, "Y")
                                Else
                                    oDBs_Details1.SetValue("U_Issued", oDBs_Details1.Offset, "N")
                                End If
                                If rs.Fields.Item("U_Approved").Value = "Y" Then
                                    oDBs_Details1.SetValue("U_Approved", oDBs_Details1.Offset, "Y")
                                Else
                                    oDBs_Details1.SetValue("U_Approved", oDBs_Details1.Offset, "N")
                                End If
                                oDBs_Details1.SetValue("U_RecvNm", oDBs_Details1.Offset, rs.Fields.Item("U_RECNM").Value)
                                objMatrix1.SetLineData(i)
                                rs.MoveNext()
                            Next
                        End If
                        Me.TotalChequeAmt(objForm.UniqueID)
                    End If

                    'Commented by ramya.Y for getting both bank and check payments on 09-11-2026
                    'If pVal.ItemUID = "Item_13" And pVal.BeforeAction = False Then
                    '    objMatrix1 = objForm.Items.Item("Mtx").Specific

                    '    Dim DateFrom As String = oDBs_Head.GetValue("U_DateFrom", 0).ToString().Trim
                    '    Dim DateTo As String = oDBs_Head.GetValue("U_DateTo", 0).ToString().Trim
                    '    Dim Branch As String = oDBs_Head.GetValue("U_BRANCH", 0).ToString().Trim
                    '    Dim Issue As String = oDBs_Head.GetValue("U_Issued", 0).ToString().Trim
                    '    Dim Approved As String = oDBs_Head.GetValue("U_Approved", 0).ToString().Trim
                    '    Dim SupplierCode As String = oDBs_Head.GetValue("U_SUPCODE", 0).ToString().Trim
                    '    Dim ChequeNo As String = oDBs_Head.GetValue("U_ChequeNo", 0).ToString().Trim
                    '    Dim GLAccount As String = oDBs_Head.GetValue("U_GLAcc", 0).ToString().Trim

                    '    If DateFrom = "" Then
                    '        objMain.objApplication.StatusBar.SetText("Please enter From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    ElseIf DateTo = "" Then
                    '        objMain.objApplication.StatusBar.SetText("Please enter To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    ElseIf Branch = "" Then
                    '        objMain.objApplication.StatusBar.SetText("Please enter Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    ElseIf GLAccount = "" Then
                    '        objMain.objApplication.StatusBar.SetText("Please enter GL Account", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    Else
                    '        Dim DetailsQuery As String = " select T2.""DocEntry"",T2.""DocDate"",T2.""CardCode"" ,T2.""CardName"" " & _
                    '                                     " , T0.""DueDate"", T0.""CheckNum"", T0.""CheckSum"",T1.""AcctCode"",T1.""AcctName"", " & _
                    '                                     " T1.""AccntntCod"" , T1.""U_DepositeA"" , T0.""U_Issued"", T0.""U_Approved""  , T0.""U_RECNM"" " & _
                    '                                     " From VPM1 T0 inner join OACT T1 on T0.""CheckAct""=T1.""AcctCode"" Inner Join OVPM T2 On (T0.""DocNum""=T2.""DocEntry"") " & _
                    '                                     " where T2.""Canceled"" ='N' and ifnull(T0.""U_Posted"",'N')='N' and Ifnull(T1.""U_PDCReq"",'')='Y' And T2.""BPLId"" = '" & Branch & "' "
                    '        Dim FromDate As String = objForm.Items.Item("From").Specific.Value.ToString().Trim()
                    '        Dim ToDate As String = objForm.Items.Item("To").Specific.Value.ToString().Trim()
                    '        If FromDate <> "" And ToDate <> "" Then
                    '            DetailsQuery = DetailsQuery & "And T2.""DocDate"" Between '" & FromDate & "' And '" & ToDate & "'"
                    '        End If
                    '        If Issue <> "" And Issue <> "A" Then
                    '            DetailsQuery = DetailsQuery & "And T0.""U_Issued""= '" & Issue & "'"
                    '        End If
                    '        If Approved <> "" And Approved <> "A" Then
                    '            DetailsQuery = DetailsQuery & "And T0.""U_Approved""= '" & Approved & "'"
                    '        End If
                    '        If SupplierCode <> "" Then
                    '            DetailsQuery = DetailsQuery & "And T2.""CardCode""= '" & SupplierCode & "'"
                    '        End If
                    '        If ChequeNo <> "" Then
                    '            DetailsQuery = DetailsQuery & "And T0.""CheckNum""= '" & ChequeNo & "'"
                    '        End If
                    '        ' madhu Based on GL account loading the data 
                    '        If GLAccount <> "" Then
                    '            DetailsQuery = DetailsQuery & " And T1.""AcctCode"" = '" & GLAccount & "' "
                    '        End If
                    '        Dim rs As SAPbobsCOM.Recordset
                    '        DetailsQuery = DetailsQuery + " group by T0.""DueDate"",T0.""CheckNum"", T0.""CheckSum"",T1.""AcctName"",T2.""DocEntry"",T2.""DocDate"",T2.""CardCode"",T2.""CardName"",T1.""AcctCode"",T1.""AccntntCod"" , T1.""U_DepositeA""  , T0.""U_Issued"" , T0.""U_Approved""  , T0.""U_RECNM"" "
                    '        rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '        rs.DoQuery(DetailsQuery)

                    '        objMatrix1.Clear()
                    '        oDBs_Details1.Clear()
                    '        objMatrix1.FlushToDataSource()
                    '        objMatrix1.AutoResizeColumns()
                    '        For i As Integer = 1 To rs.RecordCount
                    '            Dim DocDate As Date = CDate(rs.Fields.Item("DocDate").Value)
                    '            Dim DueDate As Date = CDate(rs.Fields.Item("DueDate").Value)
                    '            objMatrix1.AddRow()
                    '            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, i)
                    '            oDBs_Details1.SetValue("U_DocEntry", oDBs_Details1.Offset, rs.Fields.Item("DocEntry").Value)
                    '            oDBs_Details1.SetValue("U_DocDate", oDBs_Details1.Offset, DocDate.ToString("yyyyMMdd"))
                    '            oDBs_Details1.SetValue("U_CardCode", oDBs_Details1.Offset, rs.Fields.Item("CardCode").Value)
                    '            oDBs_Details1.SetValue("U_CardName", oDBs_Details1.Offset, rs.Fields.Item("CardName").Value)
                    '            oDBs_Details1.SetValue("U_ChequeNo", oDBs_Details1.Offset, rs.Fields.Item("CheckNum").Value)
                    '            oDBs_Details1.SetValue("U_DueDate", oDBs_Details1.Offset, DueDate.ToString("yyyyMMdd"))
                    '            oDBs_Details1.SetValue("U_ChequeSm", oDBs_Details1.Offset, rs.Fields.Item("CheckSum").Value)
                    '            oDBs_Details1.SetValue("U_AcctNum", oDBs_Details1.Offset, rs.Fields.Item("AcctCode").Value)
                    '            oDBs_Details1.SetValue("U_AcctName", oDBs_Details1.Offset, rs.Fields.Item("AcctName").Value)
                    '            oDBs_Details1.SetValue("U_DepositA", oDBs_Details1.Offset, rs.Fields.Item("U_DepositeA").Value)
                    '            oDBs_Details1.SetValue("U_JeNo", oDBs_Details1.Offset, "")
                    '            oDBs_Details1.SetValue("U_Required", oDBs_Details1.Offset, "N")
                    '            If rs.Fields.Item("U_Issued").Value = "Y" Then
                    '                oDBs_Details1.SetValue("U_Issued", oDBs_Details1.Offset, "Y")
                    '            Else
                    '                oDBs_Details1.SetValue("U_Issued", oDBs_Details1.Offset, "N")
                    '            End If
                    '            If rs.Fields.Item("U_Approved").Value = "Y" Then
                    '                oDBs_Details1.SetValue("U_Approved", oDBs_Details1.Offset, "Y")
                    '            Else
                    '                oDBs_Details1.SetValue("U_Approved", oDBs_Details1.Offset, "N")
                    '            End If
                    '            oDBs_Details1.SetValue("U_RecvNm", oDBs_Details1.Offset, rs.Fields.Item("U_RECNM").Value)
                    '            objMatrix1.SetLineData(i)
                    '            rs.MoveNext()
                    '        Next
                    '    End If
                    '    Me.TotalChequeAmt(objForm.UniqueID)
                    'End If
                    'Commented by ramya.Y for getting both bank and check payments on 09-11-2026

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PDC")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_PDC_C0")
                    objMatrix1 = objForm.Items.Item("Mtx").Specific
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                    If oCFL.UniqueID = "CA" And pVal.BeforeAction = True Then
                        Me.CFLFilter("CA")
                    End If
                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        If oCFL.UniqueID = "CA" Then
                            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, pVal.Row)
                            oDBs_Details1.SetValue("U_DocEntry", oDBs_Details1.Offset, objMatrix1.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())
                            oDBs_Details1.SetValue("U_DocDate", oDBs_Details1.Offset, objMatrix1.Columns.Item("DocDate").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())
                            oDBs_Details1.SetValue("U_CardCode", oDBs_Details1.Offset, objMatrix1.Columns.Item("CardCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())
                            oDBs_Details1.SetValue("U_CardName", oDBs_Details1.Offset, objMatrix1.Columns.Item("CardName").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())
                            oDBs_Details1.SetValue("U_ChequeNo", oDBs_Details1.Offset, objMatrix1.Columns.Item("ChequeNum").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())
                            oDBs_Details1.SetValue("U_DueDate", oDBs_Details1.Offset, objMatrix1.Columns.Item("DueDate").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())
                            oDBs_Details1.SetValue("U_ChequeSm", oDBs_Details1.Offset, objMatrix1.Columns.Item("ChequeSum").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())
                            oDBs_Details1.SetValue("U_AcctNum", oDBs_Details1.Offset, objMatrix1.Columns.Item("AcctNum").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())
                            oDBs_Details1.SetValue("U_AcctName", oDBs_Details1.Offset, objMatrix1.Columns.Item("AcctName").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())
                            oDBs_Details1.SetValue("U_DepositA", oDBs_Details1.Offset, oDT.GetValue("AcctCode", 0))
                            oDBs_Details1.SetValue("U_JeNo", oDBs_Details1.Offset, objMatrix1.Columns.Item("JeNo").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())
                            oDBs_Details1.SetValue("U_Required", oDBs_Details1.Offset, "N")
                            oDBs_Details1.SetValue("U_Issued", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())
                            oDBs_Details1.SetValue("U_RecvNm", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())
                            oDBs_Details1.SetValue("U_Approved", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())
                            objMatrix1.SetLineData(pVal.Row)
                        End If

                        If oCFL.UniqueID = "CFL_SUPCD" Then
                            oDBs_Head.SetValue("U_SUPCODE", oDBs_Head.Offset, oDT.GetValue("CardCode", 0))
                        End If

                        If oCFL.UniqueID = "CFL_GLACC" Then
                            oDBs_Head.SetValue("U_GLAcc", oDBs_Head.Offset, oDT.GetValue("AcctCode", 0))
                        End If

                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PDC")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_PDC_C0")
                    objMatrix1 = objForm.Items.Item("Mtx").Specific
                    If pVal.ItemUID = "16" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("16").Specific.Selected.Value <> "" Then
                            Dim series As Integer = objForm.Items.Item("16").Specific.Selected.Value
                            SetDefault(objForm.UniqueID, "CobmoSelect", series)
                        End If
                    End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "", Optional ByVal Series As Integer = 0)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PDC")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_PDC_C0")
            objMatrix1 = objForm.Items.Item("Mtx").Specific

            objComboBox = objForm.Items.Item("16").Specific
            Dim GetSeriesQty As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'ME_OPDC' And ""Locked"" = 'N'"
            objMain.objUtilities.ComboBoxLoadValuesSeries(objForm.Items.Item("16").Specific, GetSeriesQty)
            objForm.Items.Item("16").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            Dim GetDflt As String = "Select ""DfltSeries"" From ONNM Where ""ObjectCode"" = 'ME_OPDC'"
            Dim oRsGetDflt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDflt.DoQuery(GetDflt)

            If Flag = "" Then
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNumSeries(objForm, "ME_OPDC", oRsGetDflt.Fields.Item(0).Value))
                oDBs_Head.SetValue("U_SERIES", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)
                oDBs_Head.SetValue("Series", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)
                oDBs_Head.SetValue("U_PostDate", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
            Else
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNumSeries(objForm, "ME_OPDC", Series))
                oDBs_Head.SetValue("U_SERIES", oDBs_Head.Offset, Series)
                oDBs_Head.SetValue("Series", oDBs_Head.Offset, Series)
                oDBs_Head.SetValue("U_PostDate", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
            End If

            Dim oE As SAPbouiCOM.StaticText = objForm.Items.Item("27").Specific
            oE.Caption = ""

            objMatrix1.Clear()
            oDBs_Details1.Clear()
            objMatrix1.FlushToDataSource()
            ' Me.SetNewLine(objForm.UniqueID)

            objMatrix1.AutoResizeColumns()
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PDC")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_PDC_C0")
            objMatrix1 = objForm.Items.Item("Mtx").Specific

            objMatrix1.AddRow()
            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
            oDBs_Details1.SetValue("U_DocEntry", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_DocDate", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_CardCode", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_CardName", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_ChequeNo", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_DueDate", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_ChequeSm", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_AcctNum", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_AcctName", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_DepositA", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_JeNo", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_Required", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_Issued", oDBs_Details1.Offset, "N")
            oDBs_Details1.SetValue("U_RecvNm", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_Approved", oDBs_Details1.Offset, "")

            objMatrix1.SetLineData(objMatrix1.VisualRowCount)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub DocumentPosting(ByVal FormUID As String)
        Try
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PDC")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_PDC_C0")
            objMatrix1 = objForm.Items.Item("Mtx").Specific


            Try
                Dim Branch As String = oDBs_Head.GetValue("U_BRANCH", 0).ToString().Trim
                For i As Integer = 1 To objMatrix1.RowCount Step 1
                    Dim objCheck As SAPbouiCOM.CheckBox = objMatrix1.Columns.Item("Col_9").Cells.Item(i).Specific
                    Dim JENo As String = objMatrix1.Columns.Item("JeNo").Cells.Item(i).Specific.Value.ToString().Trim()
                    If objCheck.Checked = True And JENo = "" Then
                        Dim JournalEntry As SAPbobsCOM.JournalEntries
                        JournalEntry = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                        'madhu based on DocDate posting JE
                        'Dim DDate As String = objMatrix1.Columns.Item("DueDate").Cells.Item(i).Specific.Value.ToString()
                        'DDate = DDate.Insert(4, "/")
                        'DDate = DDate.Insert(7, "/")
                        'Dim DocDate As Date = CDate(DDate)

                        'madhu based on Posting date in line level

                        '==================== POST DATE ====================
                        Dim PostingDate As String = objMatrix1.Columns.Item("PostDate").Cells.Item(i).Specific.Value.ToString()

                        If PostingDate = "" Then
                            objMain.objApplication.StatusBar.SetText("Please select Posting Date...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Sub
                        End If

                        If PostingDate <> "" Then
                            PostingDate = PostingDate.Insert(4, "/")
                            PostingDate = PostingDate.Insert(7, "/")
                        End If

                        Dim DocDate As Date = CDate(PostingDate)

                        JournalEntry.DueDate = DocDate
                        JournalEntry.ReferenceDate = DocDate
                        JournalEntry.TaxDate = DocDate
                        JournalEntry.Memo = "PDC cheques Posting.ChequeNo :" & objMatrix1.Columns.Item("ChequeNum").Cells.Item(i).Specific.Value.ToString()
                        JournalEntry.Reference = "Customer Name:" & objMatrix1.Columns.Item("CardName").Cells.Item(i).Specific.Value.ToString()
                        JournalEntry.Reference2 = "Receipt No: " & objMatrix1.Columns.Item("DocEntry").Cells.Item(i).Specific.Value.ToString()

                        JournalEntry.Lines.BPLID = Branch
                        JournalEntry.Lines.AccountCode = objMatrix1.Columns.Item("AcctNum").Cells.Item(i).Specific.Value.ToString()
                        JournalEntry.Lines.Debit = objMatrix1.Columns.Item("ChequeSum").Cells.Item(i).Specific.Value.ToString()
                        JournalEntry.Lines.Add()

                        JournalEntry.Lines.BPLID = Branch
                        JournalEntry.Lines.AccountCode = objMatrix1.Columns.Item("DepositAc").Cells.Item(i).Specific.Value.ToString()
                        JournalEntry.Lines.Credit = objMatrix1.Columns.Item("ChequeSum").Cells.Item(i).Specific.Value.ToString()
                        JournalEntry.Lines.Add()

                        Dim RetVal As Integer = JournalEntry.Add()

                        If RetVal <> 0 Then
                            'objMain.objApplication.StatusBar.SetText("Journal entry posting Exception:" & objMain.objCompany.GetLastErrorDescription.ToString().Trim(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            objMain.objApplication.StatusBar.SetText("Please select Deposit Ac ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Sub
                        Else
                            objMain.objApplication.StatusBar.SetText("Journal Entry posted.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim Update As String = " UPDATE ""@ME_PDC_C0"" T1 SET T1.""U_JeNo""=(SELECT MAX(""TransId"") FROM OJDT) WHERE T1.""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0).ToString() & "' and T1.""LineId""='" & i & "'"
                            oRsUpdate.DoQuery(Update)
                            '--------------

                            Dim UpdateFlag As String = " Update T0 set ""U_Posted""='Y' , ""U_Issued"" = '" & objMatrix1.Columns.Item("V_0").Cells.Item(i).Specific.Value.ToString().Trim() & "'  ,  ""U_Approved"" = '" & objMatrix1.Columns.Item("V_2").Cells.Item(i).Specific.Value.ToString().Trim() & "' From VPM1 T0 " &
                                                     " Inner Join OVPM T2 On (T0.""DocNum""=T2.""DocEntry"") where T2.""Canceled"" ='N' and T0.""CheckNum""='" & objMatrix1.Columns.Item("ChequeNum").Cells.Item(i).Specific.Value.ToString().Trim() & "'and T2.""CardCode""='" & objMatrix1.Columns.Item("CardCode").Cells.Item(i).Specific.Value.ToString().Trim() & "' and T2.""DocEntry""='" & objMatrix1.Columns.Item("DocEntry").Cells.Item(i).Specific.Value.ToString().Trim() & "'"
                            Dim oRsUpdateFlags As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsUpdateFlags.DoQuery(UpdateFlag)


                            Dim UpdateFlag1 As String = " Update T0 set  ""U_RECNM"" = '" & objMatrix1.Columns.Item("V_1").Cells.Item(i).Specific.Value.ToString().Trim() & "'  From VPM1 T0 " &
                                                       " Inner Join OVPM T2 On (T0.""DocNum""=T2.""DocEntry"") where T2.""Canceled"" ='N' and T0.""CheckNum""='" & objMatrix1.Columns.Item("ChequeNum").Cells.Item(i).Specific.Value.ToString().Trim() & "'and T2.""CardCode""='" & objMatrix1.Columns.Item("CardCode").Cells.Item(i).Specific.Value.ToString().Trim() & "' and T2.""DocEntry""='" & objMatrix1.Columns.Item("DocEntry").Cells.Item(i).Specific.Value.ToString().Trim() & "'"
                            Dim oRsUpdateFlags1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsUpdateFlags1.DoQuery(UpdateFlag1)
                        End If
                    Else
                        If objCheck.Checked = True And JENo <> "" Then
                            objMain.objApplication.StatusBar.SetText("Journal entry already posted for this line.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If

                    End If
                Next
                objMain.objApplication.ActivateMenuItem("1304")
            Catch exception As System.Exception
            End Try
        Finally
            GC.Collect()
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

    Sub CFLSupplierFilter(ByVal CFL_ID As String)
        Try
            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objForm.ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            oCondition = oConditions.Add()
            oCondition.Alias = "CardType"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "S"
            oChooseFromList.SetConditions(oConditions)
        Catch ex As Exception

        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("ME_PDC_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Me.UpdateIssuedStatus(objForm.UniqueID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Me.UpdateIssuedStatus(objForm.UniqueID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Me.TotalChequeAmt(objForm.UniqueID)

                        objMatrix1 = objForm.Items.Item("Mtx").Specific
                        Dim hasJE As Boolean = False

                        For i As Integer = 1 To objMatrix1.RowCount
                            Dim jeNo As String = objMatrix1.Columns.Item("JeNo").Cells.Item(i).Specific.Value

                            If Not String.IsNullOrWhiteSpace(jeNo) Then
                                hasJE = True
                                Exit For
                            End If
                        Next

                        If hasJE Then
                            objForm.Items.Item("PostJE").Enabled = False
                        Else
                            objForm.Items.Item("PostJE").Enabled = True
                        End If
                        ''Test
                    End If



            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub UpdateIssuedStatus(ByVal FormUID As String)
        Try
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PDC")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_PDC_C0")
            objMatrix1 = objForm.Items.Item("Mtx").Specific

            For i As Integer = 1 To objMatrix1.RowCount Step 1
                Dim UpdateFlag As String = " Update T0 set  ""U_Issued"" = '" & objMatrix1.Columns.Item("V_0").Cells.Item(i).Specific.Value.ToString().Trim() & "'  , ""U_RECNM"" = '" & objMatrix1.Columns.Item("V_1").Cells.Item(i).Specific.Value.ToString().Trim() & "'  , ""U_Approved"" = '" & objMatrix1.Columns.Item("V_2").Cells.Item(i).Specific.Value.ToString().Trim() & "' From VPM1 T0 " & _
                                           " Inner Join OVPM T2 On (T0.""DocNum""=T2.""DocEntry"") where T2.""Canceled"" ='N' and T0.""CheckNum""='" & objMatrix1.Columns.Item("ChequeNum").Cells.Item(i).Specific.Value.ToString().Trim() & "' and T2.""CardCode""='" & objMatrix1.Columns.Item("CardCode").Cells.Item(i).Specific.Value.ToString().Trim() & "' and T2.""DocEntry""='" & objMatrix1.Columns.Item("DocEntry").Cells.Item(i).Specific.Value.ToString().Trim() & "'"
                Dim oRsUpdateFlags As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsUpdateFlags.DoQuery(UpdateFlag)              
            Next

        Catch ex As Exception
        End Try
    End Sub

    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_PDC")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_PDC_C0")
            objMatrix1 = objForm.Items.Item("Mtx").Specific
          
            If objForm.Items.Item("From").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("From Date  Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("To").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("To Date  Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            If objMatrix1.VisualRowCount = 0 Then
                objMain.objApplication.StatusBar.SetText("Matrix cannot be Empty")
                Return False
            ElseIf objMatrix1.VisualRowCount = 1 Then
                If (objMatrix1.Columns.Item("DocEntry").Cells.Item(1).Specific.Value.Trim = "") Then
                    objMain.objApplication.StatusBar.SetText("Matrix cannot be Empty")
                    Return False
                End If
            End If


            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Public Sub TotalChequeAmt(ByVal FormUID As String)
        Try
            Dim Value As SAPbouiCOM.StaticText
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix1 = objForm.Items.Item("Mtx").Specific
            Value = objForm.Items.Item("27").Specific
            Dim TtlChqAmt As Double = 0
            For i As Integer = 1 To objMatrix1.VisualRowCount
                TtlChqAmt = TtlChqAmt + CDbl(objMatrix1.GetCellSpecific("ChequeSum", i).Value)
            Next
            Value.Caption = TtlChqAmt
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub



End Class
