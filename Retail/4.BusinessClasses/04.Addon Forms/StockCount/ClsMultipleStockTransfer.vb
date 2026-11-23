
Imports System.Threading
Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class ClsMultipleStockTransfer

#Region "       Declaration             "
    Public objForm, objITRForm, objITRFormUDF As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details1, oDBs_Details2 As SAPbouiCOM.DBDataSource
    Dim objMatrix1, objMatrix2, objITRMatrix As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim LostFocusFlag As Boolean = False
    Dim BaseUom As String
#End Region

    Sub CreateForm()
        Try
            ''Test
            objMain.objUtilities.LoadForm("AutomaticMultipleSTR.xml", "TNX_POPost", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("TNX_POPost", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objMain.objApplication.StatusBar.SetText("Stock Transfer Request form is initializing, Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR_C1")

            objMatrix1 = objForm.Items.Item("3").Specific
            objMatrix2 = objForm.Items.Item("4").Specific


            objComboBox = objForm.Items.Item("21").Specific
            objComboBox.ValidValues.Add("", "")
            objComboBox.ValidValues.Add("Open", "Open")
            objComboBox.ValidValues.Add("Close", "Close")
            objComboBox.ValidValues.Add("Cancel", "Cancel")


            objMain.objUtilities.ComboBoxLoadValues(objForm.Items.Item("14").Specific, "Select ""BPLId"" ,  ""BPLName"" From OBPL  ")
            objForm.Items.Item("14").DisplayDesc = True


            'objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            'objForm.Items.Item("14").DisplayDesc = True
            'If objComboBox.ValidValues.Count <> 0 Then
            '    objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
            '    Me.CFLWhsCodeFilter(objForm.UniqueID, "CFL_FWHSCD", objComboBox.Value)
            'End If



            Me.CellsMasking(objForm.UniqueID)



            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub CellsMasking(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)



            objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("25").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("25").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("25").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("55").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("55").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("18").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("18").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("21").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("21").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)




        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "", Optional ByVal Series As Integer = 0)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR_C1")

            objMatrix1 = objForm.Items.Item("3").Specific
            objMatrix2 = objForm.Items.Item("4").Specific

            objComboBox = objForm.Items.Item("19").Specific
            Dim GetSeriesQty As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'ME_OSTR' And ""Locked"" = 'N'"
            objMain.objUtilities.ComboBoxLoadValuesSeries(objForm.Items.Item("19").Specific, GetSeriesQty)
            objForm.Items.Item("19").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            Dim GetDflt As String = "Select ""DfltSeries"" From ONNM Where ""ObjectCode"" = 'TNX_MPSTRUDO'"
            Dim oRsGetDflt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDflt.DoQuery(GetDflt)
            If Flag = "" Then
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNumSeries(objForm, "TNX_MPSTRUDO", oRsGetDflt.Fields.Item(0).Value))
                oDBs_Head.SetValue("U_MESRS", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)
                oDBs_Head.SetValue("Series", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)
                oDBs_Head.SetValue("U_MESTS", oDBs_Head.Offset, "Open")
                oDBs_Head.SetValue("U_MEPSDT", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
                oDBs_Head.SetValue("U_MEDOCDT", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
            Else
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNumSeries(objForm, "TNX_MPSTRUDO", Series))
                oDBs_Head.SetValue("U_MESRS", oDBs_Head.Offset, Series)
                oDBs_Head.SetValue("Series", oDBs_Head.Offset, Series)
                oDBs_Head.SetValue("U_MESTS", oDBs_Head.Offset, "Open")
                oDBs_Head.SetValue("U_MEPSDT", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
                oDBs_Head.SetValue("U_MEDOCDT", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
            End If

            objMatrix1.Clear()
            oDBs_Details1.Clear()
            objMatrix1.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "3")

            objMatrix2.Clear()
            oDBs_Details2.Clear()
            objMatrix2.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "4")

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal MatrixUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR_C1")

            objMatrix1 = objForm.Items.Item("3").Specific
            objMatrix2 = objForm.Items.Item("4").Specific

            Select Case MatrixUID
                Case "3"
                    objMatrix1.AddRow()
                    oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
                    oDBs_Details1.SetValue("U_MEBRCD", oDBs_Details1.Offset, "")
                    ''oDBs_Details1.SetValue("U_MEBRNM", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_MEITMCD", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_MEITNM", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_MEUOM", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_MEADBRCD", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_MEFWSTK", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_METWHSTK", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_MEFWHSCD", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_METWHSCD", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_MEREQQTY", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_MERELQTY", oDBs_Details1.Offset, "0")
                    oDBs_Details1.SetValue("U_MEOPNQTY", oDBs_Details1.Offset, "0")
                    oDBs_Details1.SetValue("U_MERCEQTY", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_MELNSTS", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_MEPCKNG", oDBs_Details1.Offset, "")
                    objMatrix1.SetLineData(objMatrix1.VisualRowCount)
                Case "4"
                    objMatrix2.AddRow()
                    oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, objMatrix2.VisualRowCount)
                    oDBs_Details2.SetValue("U_MEATCNM", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_MEATCH", oDBs_Details2.Offset, "")
                    objMatrix2.SetLineData(objMatrix2.VisualRowCount)
                    objMatrix2.AutoResizeColumns()
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "TNX_MULSTR" And pVal.BeforeAction = False Then
                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)

            ElseIf pVal.MenuUID = "Close" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("ME_STR_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix1 = objForm.Items.Item("3").Specific
                Dim result As Integer = objMain.objApplication.MessageBox("Changing document is irreversible..Confirm... Do you want to continue?", 1, "Continue", "Cancel")
                If (result = 1) Then
                    '  Me.UpdateStatus(objForm.UniqueID)
                End If
            ElseIf pVal.MenuUID = "Delete Row" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("TNX_POPost", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix1 = objForm.Items.Item("3").Specific
                For i As Integer = 1 To objMatrix1.VisualRowCount - 1
                    If objMatrix1.IsRowSelected(i) = True Then
                        objMatrix1.DeleteRow(i)
                    End If
                Next
                For i As Integer = 1 To objMatrix1.VisualRowCount
                    objMatrix1.Columns.Item("V_-1").Cells.Item(i).Specific.string = i
                Next
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    'objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                End If
            ElseIf pVal.MenuUID = "519" And pVal.BeforeAction = True Then 'Preview
                Try
                    Dim frs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'Transfer Request' AND ""Type"" = 'C'")
                    If frs.RecordCount = 0 Then
                        objMain.objApplication.MessageBox("Report layout 'Transfer Request' not found. Please upload with same name and Code as parameter", 0, "OK")
                    Else
                        objMain.objApplication.ActivateMenuItem(frs.Fields.Item(0).Value.ToString())
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_POPost")
                        Dim CrForm As SAPbouiCOM.Form
                        Dim oedt As SAPbouiCOM.EditText
                        CrForm = objMain.objApplication.Forms.ActiveForm
                        oedt = CrForm.Items.Item("1000003").Specific
                        oedt.Value = oDBs_Head.GetValue("DocEntry", 0)
                        CrForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                Catch ex As Exception
                End Try
            ElseIf pVal.MenuUID = "520" And pVal.BeforeAction = True Then 'Print
                Try
                    Dim frs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'Transfer Request' AND ""Type"" = 'C'")
                    If frs.RecordCount = 0 Then
                        objMain.objApplication.MessageBox("Report layout 'Transfer Request' not found. Please upload with same name and Code as parameter", 0, "OK")
                    Else
                        objMain.objApplication.ActivateMenuItem(frs.Fields.Item(0).Value.ToString())
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_POPost")
                        Dim CrForm As SAPbouiCOM.Form
                        Dim oedt As SAPbouiCOM.EditText
                        CrForm = objMain.objApplication.Forms.ActiveForm
                        oedt = CrForm.Items.Item("1000003").Specific
                        oedt.Value = oDBs_Head.GetValue("DocEntry", 0)
                        CrForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                Catch ex As Exception
                End Try
            End If
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub



    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR_C0")
                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR_C1")

                    objMatrix1 = objForm.Items.Item("3").Specific
                    objMatrix2 = objForm.Items.Item("4").Specific
                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                    '              (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                    '        objForm.Items.Item("42").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    '        LostFocusFlag = True
                    '        If Me.Validation(objForm.UniqueID, "Y") = False Then
                    '            LostFocusFlag = False
                    '            BubbleEvent = False
                    '        End If
                    '    End If
                    'If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    '        LostFocusFlag = False
                    '    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        Me.SetDefault(objForm.UniqueID)
                    End If
                    If pVal.ItemUID = "27" And pVal.BeforeAction = False Then
                        objForm.PaneLevel = 1
                        objMatrix1.AutoResizeColumns()
                    End If
                    If pVal.ItemUID = "28" And pVal.BeforeAction = False Then
                        objForm.PaneLevel = 2
                        objMatrix2.AutoResizeColumns()
                    End If
                    If pVal.ItemUID = "1000005" And pVal.BeforeAction = False Then
                        '  Me.GetPDFFile("Transfer Request SD.rpt", "Transfer Request SD", "DocKey@")
                    End If
                    If pVal.ItemUID = "57" And pVal.BeforeAction = False Then

                        Dim FWhse As String = objForm.Items.Item("6").Specific.Value
                        Dim GrpoNumber As String = objForm.Items.Item("GRPNO").Specific.Value
                        Me.LoadItemDetails(FWhse, GrpoNumber)
                    End If
                    If pVal.ItemUID = "Item_0" And pVal.BeforeAction = False Then

                        Me.UpdateSTR(objForm.UniqueID, oDBs_Head.GetValue("DocNum", 0), oDBs_Head.GetValue("DocEntry", 0))

                    End If
                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR_C0")
                    objMatrix1 = objForm.Items.Item("3").Specific
                    If pVal.ItemUID = "3" And pVal.ColUID = "V_-1" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                        If objMain.objApplication.Forms.ActiveForm.Title = "Row Details... - GRPO Allocation" Then
                            objMain.objApplication.Forms.ActiveForm.Close()
                        End If
                        If objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value <> "" Then         'DocType = Service
                            BubbleEvent = False
                            Dim Itemcode As String = objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value
                            Dim ItemName As String = objMatrix1.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value
                            Dim UOM As String = objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value
                            Dim BarCode As String = objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value
                            ''Dim Branch As String = objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value
                            Dim Warehouse As String = objMatrix1.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value
                            Dim Quan As String = objMatrix1.Columns.Item("V_11").Cells.Item(pVal.Row).Specific.Value
                            objMain.objwarehouse.CreateForm(objForm.UniqueID, Itemcode,
                                                               ItemName, UOM, Warehouse, BarCode,
                                                               CDbl(objMatrix1.Columns.Item("V_11").Cells.Item(pVal.Row).Specific.Value),
                                                               oDBs_Head.GetValue("DocEntry", 0), oDBs_Head.GetValue("DocNum", 0), pVal.Row
                                                              )
                        End If
                    End If


                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.ActiveForm()
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR_C0")
                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR_C1")

                    objMatrix1 = objForm.Items.Item("3").Specific
                    objMatrix2 = objForm.Items.Item("4").Specific

                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                    If oCFL.UniqueID = "CFL_UOM" And pVal.BeforeAction = True Then
                        Me.CFLUOM(objForm.UniqueID, "CFL_UOM", objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                    End If
                    'If oCFL.UniqueID = "CFL_7" And pVal.BeforeAction = True Then

                    '    Me.CFLGrPO(objForm.UniqueID, "CFL_7")

                    'End If

                    If oCFL.UniqueID = "CFL_BCD" And pVal.BeforeAction = True Then

                        Me.CFLBarCode(objForm.UniqueID, "CFL_BCD", objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value, objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                    End If

                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                        If oCFL.UniqueID = "CFL_8" Then
                            oDBs_Head.SetValue("U_VCode", oDBs_Head.Offset, oDT.GetValue("CardCode", 0))
                        End If

                        If oCFL.UniqueID = "CFL_7" And pVal.BeforeAction = False Then

                            oDBs_Head.SetValue("U_GRPNO", oDBs_Head.Offset, oDT.GetValue("DocNum", 0))
                            oDBs_Head.SetValue("U_GRPDM", oDBs_Head.Offset, oDT.GetValue("DocEntry", 0))
                            oDBs_Head.SetValue("U_VCode", oDBs_Head.Offset, oDT.GetValue("CardCode", 0))

                            Dim DocEntry As String = oDT.GetValue("DocEntry", 0)
                            Dim GetWhs As String = "Select Top 1 ""WhsCode"" from PDN1 where ""DocEntry""= '" + DocEntry + "' "
                            Dim oRsGetWhs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetWhs.DoQuery(GetWhs)

                            Dim GetWhsN As String = "Select * from OWHS where ""WhsCode""= '" + oRsGetWhs.Fields.Item("WhsCode").Value + "' "
                            Dim oRsGetWhsN As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetWhsN.DoQuery(GetWhsN)

                            oDBs_Head.SetValue("U_MEFWHSCD", oDBs_Head.Offset, oRsGetWhs.Fields.Item("WhsCode").Value)
                            oDBs_Head.SetValue("U_MEFWHSNM", oDBs_Head.Offset, oRsGetWhsN.Fields.Item("WhsName").Value)
                            oDBs_Head.SetValue("U_MEFBPID", oDBs_Head.Offset, oRsGetWhsN.Fields.Item("BPLid").Value)
                        End If


                        If oCFL.UniqueID = "CFL_FWHSCD" Then

                            oDBs_Head.SetValue("U_MEFWHSCD", oDBs_Head.Offset, oDT.GetValue("WhsCode", 0))
                            oDBs_Head.SetValue("U_MEFWHSNM", oDBs_Head.Offset, oDT.GetValue("WhsName", 0))
                            oDBs_Head.SetValue("U_MEFBPID", oDBs_Head.Offset, oDT.GetValue("BPLid", 0))
                        End If

                        If oCFL.UniqueID = "CFL_ITMCD" Then
                            Dim PCPrPack As String = objMain.objUtilities.GetPCPerItem(oDT.GetValue("ItemCode", 0))
                            BaseUom = objMain.objUtilities.GetItemBaseUom(oDT.GetValue("ItemCode", 0))
                            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, pVal.Row)
                            Dim GetStockDtls As String = " Select (Select ""UomCode"" From OUOM Where ""UomEntry"" = A.""IUoMEntry"") As ""UomCode"" , A.""IUoMEntry"" , (Select " & IsNull & "(T0.""OnHand"",0) FROM OITW T0 " &
                                             " Where T0.""ItemCode"" = A.""ItemCode"" And  T0.""WhsCode""='" & objForm.Items.Item("6").Specific.Value & "') As ""FromWhsStock"" , " &
                                             " (Select " & IsNull & "(T1.""OnHand"",0) FROM OITW T1 Where T1.""ItemCode"" = A.""ItemCode"" And  " &
                                             " T1.""WhsCode""='" & objForm.Items.Item("8").Specific.Value & "') As ""ToWhsStock"" , (Select SUM(" & IsNull & "(T2.""OnHand"",0))  " &
                                             " FROM OITW T2 Where T2.""ItemCode"" = A.""ItemCode"" ) As ""OverallStock"" " &
                                             " From OITM A  Where ""ItemCode""  = '" & oDT.GetValue("ItemCode", 0) & "'  "
                            Dim oRsGetStockDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetStockDtls.DoQuery(GetStockDtls)
                            Dim GetBarCode As String = " SELECT Top 1 ""BcdCode"" FROM OBCD WHERE ""UomEntry"" = '" & oRsGetStockDtls.Fields.Item("IUoMEntry").Value & "' " &
                                                   " And ""ItemCode"" = '" & oDT.GetValue("ItemCode", 0) & "' ORDER BY ""BcdEntry"" ASC "
                            Dim oRsGetBarCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetBarCode.DoQuery(GetBarCode)
                            If oRsGetBarCode.RecordCount > 0 Then
                                oDBs_Details1.SetValue("U_MEBRCD", oDBs_Details1.Offset, oRsGetBarCode.Fields.Item("BcdCode").Value)
                                oDBs_Details1.SetValue("U_MEADBRCD", oDBs_Details1.Offset, oRsGetBarCode.Fields.Item("BcdCode").Value)
                            Else
                                oDBs_Details1.SetValue("U_MEBRCD", oDBs_Details1.Offset, "")
                                oDBs_Details1.SetValue("U_MEADBRCD", oDBs_Details1.Offset, "")
                            End If
                            oDBs_Details1.SetValue("U_MEITMCD", oDBs_Details1.Offset, oDT.GetValue("ItemCode", 0))
                            oDBs_Details1.SetValue("U_MEITNM", oDBs_Details1.Offset, oDT.GetValue("ItemName", 0))
                            oDBs_Details1.SetValue("U_MEUOM", oDBs_Details1.Offset, oRsGetStockDtls.Fields.Item("UomCode").Value)
                            If PCPrPack <> "" Then
                                oDBs_Details1.SetValue("U_MEPCKNG", oDBs_Details1.Offset, PCPrPack & "*" & "1*" & BaseUom)
                            Else
                                oDBs_Details1.SetValue("U_MEPCKNG", oDBs_Details1.Offset, "1*" & BaseUom)
                            End If
                            oDBs_Details1.SetValue("U_MEFWSTK", oDBs_Details1.Offset, oRsGetStockDtls.Fields.Item("FromWhsStock").Value)
                            oDBs_Details1.SetValue("U_METWHSTK", oDBs_Details1.Offset, oRsGetStockDtls.Fields.Item("ToWhsStock").Value)
                            oDBs_Details1.SetValue("U_MEFWHSCD", oDBs_Details1.Offset, objForm.Items.Item("6").Specific.Value)
                            oDBs_Details1.SetValue("U_METWHSCD", oDBs_Details1.Offset, objForm.Items.Item("8").Specific.Value)
                            oDBs_Details1.SetValue("U_MEREQQTY", oDBs_Details1.Offset, "")
                            oDBs_Details1.SetValue("U_MERELQTY", oDBs_Details1.Offset, "")
                            oDBs_Details1.SetValue("U_MEOPNQTY", oDBs_Details1.Offset, "")
                            oDBs_Details1.SetValue("U_MERCEQTY", oDBs_Details1.Offset, "")
                            oDBs_Details1.SetValue("U_MELNSTS", oDBs_Details1.Offset, "Open")
                            objMatrix1.SetLineData(pVal.Row)
                            If pVal.Row = objMatrix1.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "3")
                            End If
                        End If

                        If oCFL.UniqueID = "CFL_BCD" Then
                            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, pVal.Row)
                            Dim GetStockDtls As String = " Select A.""ItemName"" , (Select ""UomCode"" From OUOM Where ""UomEntry"" = A.""IUoMEntry"") As ""UomCode"" , A.""IUoMEntry"" , A.""UgpEntry"" , (Select " & IsNull & "(T0.""OnHand"",0) FROM OITW T0 " &
                                             " Where T0.""ItemCode"" = A.""ItemCode"" And  T0.""WhsCode""='" & objForm.Items.Item("6").Specific.Value & "') As ""FromWhsStock"" , " &
                                             " (Select " & IsNull & "(T1.""OnHand"",0) FROM OITW T1 Where T1.""ItemCode"" = A.""ItemCode"" And  " &
                                             " T1.""WhsCode""='" & objForm.Items.Item("8").Specific.Value & "') As ""ToWhsStock"" , (Select SUM(" & IsNull & "(T2.""OnHand"",0))  " &
                                             " FROM OITW T2 Where T2.""ItemCode"" = A.""ItemCode"" ) As ""OverallStock"" " &
                                             " From OITM A  Where ""ItemCode""  = '" & oDT.GetValue("ItemCode", 0) & "'  "
                            Dim oRsGetStockDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetStockDtls.DoQuery(GetStockDtls)
                            oDBs_Details1.SetValue("U_MEBRCD", oDBs_Details1.Offset, oDT.GetValue("BcdCode", 0))
                            ''oDBs_Details1.SetValue("U_MEBRNM", oDBs_Details1.Offset, oDT.GetValue("BcdName", 0))
                            oDBs_Details1.SetValue("U_MEADBRCD", oDBs_Details1.Offset, oDT.GetValue("BcdCode", 0))
                            oDBs_Details1.SetValue("U_MEITMCD", oDBs_Details1.Offset, oDT.GetValue("ItemCode", 0))
                            Dim PCPrPack As String = objMain.objUtilities.GetPCPerItem(oDT.GetValue("ItemCode", 0))
                            BaseUom = objMain.objUtilities.GetItemBaseUom(oDT.GetValue("ItemCode", 0))
                            oDBs_Details1.SetValue("U_MEITNM", oDBs_Details1.Offset, oRsGetStockDtls.Fields.Item("ItemName").Value)

                            If objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value = "" Then
                                Dim GetUOMCode As String = " Select T1.""UomCode"" From OBCD T0 Inner Join OUOM T1 On T1.""UomEntry"" = T0.""UomEntry"" " &
                                                       " Where T0.""BcdCode""  = '" & oDT.GetValue("BcdCode", 0) & "' And T0.""ItemCode"" = '" & oDT.GetValue("ItemCode", 0) & "' "
                                Dim oRsGetUOMCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetUOMCode.DoQuery(GetUOMCode)
                                oDBs_Details1.SetValue("U_MEUOM", oDBs_Details1.Offset, oRsGetUOMCode.Fields.Item("UomCode").Value)
                                Dim ChkInventoryUOM As String = " Select ""IUoMEntry"" From OITM  T0 Inner Join OUOM T1 On T1.""UomEntry"" = ""IUoMEntry"" " &
                                                            " Where ""ItemCode"" = '" & oDT.GetValue("ItemCode", 0) & "' And T1.""UomCode"" = '" & oRsGetUOMCode.Fields.Item("UomCode").Value & "'"
                                Dim oRsChkInventoryUOM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsChkInventoryUOM.DoQuery(ChkInventoryUOM)
                                If oRsChkInventoryUOM.RecordCount > 0 Then
                                    If PCPrPack = "" Then
                                        oDBs_Details1.SetValue("U_MEPCKNG", oDBs_Details1.Offset, "1*" & BaseUom)
                                    Else
                                        oDBs_Details1.SetValue("U_MEPCKNG", oDBs_Details1.Offset, PCPrPack & "*" & "1*" & BaseUom)
                                    End If
                                Else
                                    Dim GetPackageCode As String = " Select Cast(Cast(T0.""BaseQty"" As Int) As Varchar)" & Concate & "'*'" & Concate & "'" & BaseUom & "'" &
                                                               "  From UGP1 T0 Inner Join OUOM T1 On  " &
                                                               " T1.""UomEntry"" = T0.""UomEntry""  Where ""UgpEntry"" = '" & oRsGetStockDtls.Fields.Item("UgpEntry").Value & "' And  " &
                                                               " T1.""UomCode"" = '" & oRsGetUOMCode.Fields.Item("UomCode").Value & "' "
                                    Dim oRsGetPackageCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRsGetPackageCode.DoQuery(GetPackageCode)
                                    oDBs_Details1.SetValue("U_MEPCKNG", oDBs_Details1.Offset, oRsGetPackageCode.Fields.Item(0).Value)
                                    If PCPrPack <> "" Then
                                        oDBs_Details1.SetValue("U_MEPCKNG", oDBs_Details1.Offset, PCPrPack & "*" & oRsGetPackageCode.Fields.Item(0).Value)
                                    End If

                                End If
                            Else
                                Dim ChkInventoryUOM As String = " Select ""UomCode"" From OUOM Where ""UomEntry"" = '" & oRsGetStockDtls.Fields.Item("IUoMEntry").Value & "' And ""UomCode"" = '" & objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value & "'  "
                                Dim oRsChkInventoryUOM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsChkInventoryUOM.DoQuery(ChkInventoryUOM)
                                If oRsChkInventoryUOM.RecordCount > 0 Then
                                    If PCPrPack = "" Then
                                        oDBs_Details1.SetValue("U_MEPCKNG", oDBs_Details1.Offset, "1*" & BaseUom)
                                    Else
                                        oDBs_Details1.SetValue("U_MEPCKNG", oDBs_Details1.Offset, PCPrPack & "*" & "1*" & BaseUom)
                                    End If
                                Else
                                    Dim GetPackageCode As String = " Select Cast(Cast(T0.""BaseQty"" As Int) As Varchar)" & Concate & "'*'" & Concate & "'" & BaseUom & "'" &
                                                              " From UGP1 T0 Inner Join OUOM T1 On  " &
                                                              " T1.""UomEntry"" = T0.""UomEntry""  Where ""UgpEntry"" = '" & oRsGetStockDtls.Fields.Item("UgpEntry").Value & "' And  " &
                                                              " T1.""UomCode"" = '" & objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value & "' "
                                    Dim oRsGetPackageCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRsGetPackageCode.DoQuery(GetPackageCode)
                                    oDBs_Details1.SetValue("U_MEPCKNG", oDBs_Details1.Offset, oRsGetPackageCode.Fields.Item(0).Value)
                                    If PCPrPack <> "" Then
                                        oDBs_Details1.SetValue("U_MEPCKNG", oDBs_Details1.Offset, PCPrPack & "*" & oRsGetPackageCode.Fields.Item(0).Value)
                                    End If
                                End If
                                oDBs_Details1.SetValue("U_MEUOM", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                            End If
                            oDBs_Details1.SetValue("U_MEFWSTK", oDBs_Details1.Offset, oRsGetStockDtls.Fields.Item("FromWhsStock").Value)
                            oDBs_Details1.SetValue("U_METWHSTK", oDBs_Details1.Offset, oRsGetStockDtls.Fields.Item("ToWhsStock").Value)
                            oDBs_Details1.SetValue("U_MEFWHSCD", oDBs_Details1.Offset, objForm.Items.Item("6").Specific.Value)
                            oDBs_Details1.SetValue("U_METWHSCD", oDBs_Details1.Offset, objForm.Items.Item("8").Specific.Value)
                            oDBs_Details1.SetValue("U_MEREQQTY", oDBs_Details1.Offset, "")
                            oDBs_Details1.SetValue("U_MERELQTY", oDBs_Details1.Offset, "")
                            oDBs_Details1.SetValue("U_MEOPNQTY", oDBs_Details1.Offset, "")
                            oDBs_Details1.SetValue("U_MERCEQTY", oDBs_Details1.Offset, "")
                            oDBs_Details1.SetValue("U_MELNSTS", oDBs_Details1.Offset, "Open")
                            objMatrix1.SetLineData(pVal.Row)
                            If pVal.Row = objMatrix1.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "3")
                            End If
                        End If

                        If oCFL.UniqueID = "CFL_UOM" Then
                            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, pVal.Row)
                            Dim ChkBarCodeUOM As String = " Select ""BcdCode"" From OBCD Where ""UomEntry"" = " &
                                                      " (Select ""UomEntry"" From OUOM Where ""UomCode"" = '" & oDT.GetValue("UomCode", 0) & "') " &
                                                      " And ""ItemCode"" = '" & objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value & "' "
                            Dim oRsChkBarCodeUOM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsChkBarCodeUOM.DoQuery(ChkBarCodeUOM)
                            If oRsChkBarCodeUOM.RecordCount = 0 Then
                                Dim GetBarCode As String = " SELECT Top 1 ""BcdCode"" FROM OBCD WHERE ""UomEntry"" = " &
                                                       " (SELECT ""IUoMEntry"" FROM OITM WHERE ""ItemCode"" =  '" & objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value & "') " &
                                                       " And ""ItemCode"" = '" & objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value & "' ORDER BY ""BcdEntry"" ASC "
                                Dim oRsGetBarCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetBarCode.DoQuery(GetBarCode)
                                If oRsGetBarCode.RecordCount > 0 Then
                                    oDBs_Details1.SetValue("U_MEBRCD", oDBs_Details1.Offset, oRsGetBarCode.Fields.Item("BcdCode").Value)
                                    oDBs_Details1.SetValue("U_MEADBRCD", oDBs_Details1.Offset, oRsGetBarCode.Fields.Item("BcdCode").Value)
                                End If
                            Else
                                oDBs_Details1.SetValue("U_MEBRCD", oDBs_Details1.Offset, oRsChkBarCodeUOM.Fields.Item("BcdCode").Value)
                                oDBs_Details1.SetValue("U_MEADBRCD", oDBs_Details1.Offset, oRsChkBarCodeUOM.Fields.Item("BcdCode").Value)
                            End If
                            oDBs_Details1.SetValue("U_MEITMCD", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_MEITNM", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_MEUOM", oDBs_Details1.Offset, oDT.GetValue("UomCode", 0))

                            Dim ChkInventoryUOM As String = " Select ""IUoMEntry"" From OITM  T0 Inner Join OUOM T1 On T1.""UomEntry"" = ""IUoMEntry"" " &
                                                          " Where ""ItemCode"" = '" & objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value & "' And T1.""UomCode"" = '" & oDT.GetValue("UomCode", 0) & "'"
                            Dim oRsChkInventoryUOM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsChkInventoryUOM.DoQuery(ChkInventoryUOM)
                            Dim PCPrPack As String = objMain.objUtilities.GetPCPerItem(objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                            BaseUom = objMain.objUtilities.GetItemBaseUom(objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                            If oRsChkInventoryUOM.RecordCount > 0 Then
                                If PCPrPack = "" Then
                                    oDBs_Details1.SetValue("U_MEPCKNG", oDBs_Details1.Offset, "1*" & BaseUom)
                                Else
                                    oDBs_Details1.SetValue("U_MEPCKNG", oDBs_Details1.Offset, PCPrPack & "*" & "1*" & BaseUom)
                                End If
                            Else
                                'Dim GetPackageCode As String = " Select Cast(Cast(T1.""BaseQty"" As Int) As Varchar)+'*Pcs' From  UGP1 T1 Inner Join OUOM T2 On T2.""UomEntry"" = T1.""UomEntry""  " & _
                                '  "  Where T1.""UgpEntry"" " & _
                                '  " = (Select ""UgpEntry"" From OITM Where ""ItemCode"" = '" & objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value & "') And  T2.""UomCode"" = '" & oDT.GetValue("UomCode", 0) & "'"
                                'Dim oRsGetPackageCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                'oRsGetPackageCode.DoQuery(GetPackageCode)

                                Dim GetPackageCode As String = " Select Cast(Cast(T0.""BaseQty"" As Int) As Varchar)" & Concate & "'*'" & Concate & "'" & BaseUom & "'" &
                                                           " From UGP1 T0 Inner Join OUOM T1 On  " &
                                                           " T1.""UomEntry"" = T0.""UomEntry""  Where ""UgpEntry"" = (Select ""UgpEntry"" From OITM Where ""ItemCode"" = '" & objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value & "') And  " &
                                                           " T1.""UomCode"" = '" & oDT.GetValue("UomCode", 0) & "' "
                                Dim oRsGetPackageCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetPackageCode.DoQuery(GetPackageCode)
                                oDBs_Details1.SetValue("U_MEPCKNG", oDBs_Details1.Offset, oRsGetPackageCode.Fields.Item(0).Value)
                                If PCPrPack <> "" Then
                                    oDBs_Details1.SetValue("U_MEPCKNG", oDBs_Details1.Offset, PCPrPack & "*" & oRsGetPackageCode.Fields.Item(0).Value)
                                End If
                            End If
                            oDBs_Details1.SetValue("U_MEFWSTK", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_METWHSTK", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_MEFWHSCD", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_METWHSCD", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_MEREQQTY", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_9").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_MERELQTY", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_10").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_MEOPNQTY", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_11").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_MERCEQTY", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_12").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_MELNSTS", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_13").Cells.Item(pVal.Row).Specific.Value)
                            objMatrix1.SetLineData(pVal.Row)
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR_C0")
                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR_C1")

                    objMatrix1 = objForm.Items.Item("3").Specific
                    objMatrix2 = objForm.Items.Item("4").Specific


                    If pVal.ItemUID = "3" And pVal.ColUID = "V_9" And pVal.BeforeAction = False And LostFocusFlag = False And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, pVal.Row)
                        oDBs_Details1.SetValue("U_MEBRCD", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                        ''oDBs_Details1.SetValue("U_MEBRNM", oDBs_Details1.Offset, objMatrix1.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MEITMCD", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MEITNM", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MEUOM", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MEADBRCD", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MEFWSTK", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_METWHSTK", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MEFWHSCD", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_METWHSCD", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MEREQQTY", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_9").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MERELQTY", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_10").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MEOPNQTY", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_9").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MERCEQTY", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_12").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MELNSTS", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_13").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MEPCKNG", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_14").Cells.Item(pVal.Row).Specific.Value)
                        objMatrix1.SetLineData(pVal.Row)
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_STR")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_STR_C0")
                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@ME_STR_C1")
                    objMatrix1 = objForm.Items.Item("3").Specific
                    objMatrix2 = objForm.Items.Item("4").Specific
                    If pVal.ItemUID = "19" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("19").Specific.Selected.Value <> "" Then
                            Dim series As Integer = objForm.Items.Item("19").Specific.Selected.Value
                            SetDefault(objForm.UniqueID, "CobmoSelect", series)
                        End If
                    End If
                    If pVal.ItemUID = "14" And pVal.BeforeAction = False Then
                        objComboBox = objForm.Items.Item("14").Specific
                        Dim objComboBox1 As SAPbouiCOM.ComboBox = objForm.Items.Item("16").Specific

                        If objComboBox.Value = objComboBox1.Value Then
                            objComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                            oDBs_Head.SetValue("U_MEFWHSCD", oDBs_Head.Offset, "")
                            oDBs_Head.SetValue("U_MEFWHSNM", oDBs_Head.Offset, "")
                            objMain.objApplication.StatusBar.SetText("From Branch and To Branch cannot be same", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        Me.CFLWhsCodeFilter(objForm.UniqueID, "CFL_FWHSCD", objComboBox.Value)
                    End If
                    If pVal.ItemUID = "16" And pVal.BeforeAction = False Then
                        objComboBox = objForm.Items.Item("16").Specific
                        Dim objComboBox1 As SAPbouiCOM.ComboBox = objForm.Items.Item("14").Specific
                        If objComboBox.Value = objComboBox1.Value Then
                            objComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                            oDBs_Head.SetValue("U_METWHSCD", oDBs_Head.Offset, "")
                            oDBs_Head.SetValue("U_METWHSNM", oDBs_Head.Offset, "")
                            objMain.objApplication.StatusBar.SetText("From Branch and To Branch cannot be same", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        Me.CFLWhsCodeFilter(objForm.UniqueID, "CFL_TWHSCD", objComboBox.Value)
                    End If
                    If pVal.ItemUID = "51" And pVal.BeforeAction = False Then
                        objComboBox = objForm.Items.Item("51").Specific
                        If objComboBox.Value = "ST" Then
                            objForm.Items.Item("53").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            objForm.Items.Item("54").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            objForm.Items.Item("55").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            objForm.Items.Item("57").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        Else
                            objForm.Items.Item("53").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            objForm.Items.Item("54").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            objForm.Items.Item("55").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            objForm.Items.Item("57").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        End If

                    End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub



    Function Validation(ByVal FormUID As String, ByVal Flag As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR_C1")

            objMatrix1 = objForm.Items.Item("3").Specific
            objMatrix2 = objForm.Items.Item("4").Specific


            If objForm.Items.Item("6").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("From Warehouse Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False

            End If

            If objMatrix1.VisualRowCount = 1 Then
                objMain.objApplication.StatusBar.SetText("Matrix Cannot be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                If Flag = "Y" Then
                    For i As Integer = 1 To objMatrix1.VisualRowCount - 1
                        If objMatrix1.Columns.Item("V_1").Cells.Item(i).Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Item Code is Empty in Line : " & i & ", Delete the line before proceeding", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    Next
                    For i As Integer = 1 To objMatrix1.VisualRowCount - 1
                        If objMatrix1.Columns.Item("V_9").Cells.Item(i).Specific.Value = "" Then
                            objMatrix1.Columns.Item("V_9").Cells.Item(i).Specific.Value = 0
                        End If
                        If objMatrix1.Columns.Item("V_9").Cells.Item(i).Specific.Value <= 0 Then
                            objMain.objApplication.StatusBar.SetText("Require Qty should not be zero or less than zero.. in Line : " & i & ", Delete the line before proceeding", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    Next

                    For i As Integer = 1 To objMatrix1.VisualRowCount - 1
                        If objMatrix1.Columns.Item("V_1").Cells.Item(i).Specific.Value <> "" And objForm.Items.Item("6").Specific.Value <> objMatrix1.Columns.Item("V_7").Cells.Item(i).Specific.Value Then
                            objMain.objApplication.StatusBar.SetText("Header From Branch should match with Item level From Brach  : " & i & ", Delete the line before proceeding", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    Next

                    For i As Integer = 1 To objMatrix1.VisualRowCount - 1
                        If objMatrix1.Columns.Item("V_1").Cells.Item(i).Specific.Value <> "" And objForm.Items.Item("8").Specific.Value <> objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value Then
                            objMain.objApplication.StatusBar.SetText("Header To Branch should match with Item level To Brach  : " & i & ", Delete the line before proceeding", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    Next

                    Dim ItemsArry(objMatrix1.VisualRowCount) As String
                    For i As Integer = 1 To objMatrix1.VisualRowCount - 1
                        ItemsArry(i - 1) = objMatrix1.Columns.Item("V_1").Cells.Item(i).Specific.Value
                    Next
                    If HasDuplicates(ItemsArry) = True Then
                        Return False
                    End If
                End If
            End If
            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Sub CFLBarCode(ByVal FormUID As String, ByVal CFL_ID As String, ByVal ItemCode As String, ByVal UomEntry As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim GetBarCode As String = " SELECT ""BcdCode"" FROM OBCD WHERE ""UomEntry"" =   (Select ""UomEntry"" From OUOM Where ""UomCode"" = '" & UomEntry & "') " &
                                   " And ""ItemCode"" = '" & ItemCode & "'"
            Dim oRsGetBarCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetBarCode.DoQuery(GetBarCode)
            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            If oRsGetBarCode.RecordCount > 0 Then
                oCondition = oConditions.Add()
                oCondition.Alias = "BcdCode"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = oRsGetBarCode.Fields.Item(0).Value
                oChooseFromList.SetConditions(oConditions)
                oRsGetBarCode.MoveNext()
                For i As Integer = 1 To oRsGetBarCode.RecordCount - 1
                    oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCondition = oConditions.Add()
                    oCondition.Alias = "BcdCode"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = oRsGetBarCode.Fields.Item(0).Value
                    oChooseFromList.SetConditions(oConditions)
                    oRsGetBarCode.MoveNext()
                Next
            Else
                Dim GetBarCode1 As String = " SELECT  ""BcdCode"" FROM OBCD WHERE ""UomEntry"" = " &
                                        " (SELECT ""IUoMEntry"" FROM OITM WHERE ""ItemCode"" =  '" & ItemCode & "') " &
                                        " And ""ItemCode"" = '" & ItemCode & "' ORDER BY ""BcdEntry"" ASC "
                Dim oRsGetBarCode1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetBarCode1.DoQuery(GetBarCode1)
                If oRsGetBarCode1.RecordCount > 0 Then
                    oCondition = oConditions.Add()
                    oCondition.Alias = "BcdCode"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = oRsGetBarCode1.Fields.Item(0).Value
                    oChooseFromList.SetConditions(oConditions)
                    oRsGetBarCode1.MoveNext()
                    For i As Integer = 1 To oRsGetBarCode1.RecordCount - 1
                        oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                        oCondition = oConditions.Add()
                        oCondition.Alias = "BcdCode"
                        oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCondition.CondVal = oRsGetBarCode1.Fields.Item(0).Value
                        oChooseFromList.SetConditions(oConditions)
                        oRsGetBarCode1.MoveNext()
                    Next
                End If
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CFLUOM(ByVal FormUID As String, ByVal CFL_ID As String, ByVal ItemCode As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim GetBarCode As String = " Select T1.""UomEntry"" From  UGP1 T1 " &
                                   "  Where T1.""UgpEntry"" " &
                                   " = (Select ""UgpEntry"" From OITM Where ""ItemCode"" = '" & ItemCode & "')"
            Dim oRsGetBarCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetBarCode.DoQuery(GetBarCode)
            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            If oRsGetBarCode.RecordCount > 0 Then
                oCondition = oConditions.Add()
                oCondition.Alias = "UomEntry"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = oRsGetBarCode.Fields.Item(0).Value
                oChooseFromList.SetConditions(oConditions)
                oRsGetBarCode.MoveNext()
                For i As Integer = 1 To oRsGetBarCode.RecordCount - 1
                    oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCondition = oConditions.Add()
                    oCondition.Alias = "UomEntry"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = oRsGetBarCode.Fields.Item(0).Value
                    oChooseFromList.SetConditions(oConditions)
                    oRsGetBarCode.MoveNext()
                Next
            Else
                Dim GetBarCode1 As String = " SELECT ""IUoMEntry""  FROM OITM WHERE " &
                                       " ""ItemCode"" = '" & ItemCode & "'"
                Dim oRsGetBarCode1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetBarCode1.DoQuery(GetBarCode1)
                oCondition = oConditions.Add()
                oCondition.Alias = "UomEntry"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = oRsGetBarCode1.Fields.Item(0).Value
                oChooseFromList.SetConditions(oConditions)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
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

    Private Sub LoadMatrix(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_STR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_STR_C0")
            objMatrix1 = objForm.Items.Item("3").Specific

            Dim GetPRDetails As String = " Select  ""ItemCode"" , ""Dscription"", ""Quantity"" , ""Price"" , ""PQTReqDate"" , ""UomCode"" " &
                                     " From OPRQ T0 Inner Join PRQ1 T1 On T0.""DocEntry"" =  T1.""DocEntry"" " &
                                     " Where T0.""DocEntry"" = '" & objForm.Items.Item("55").Specific.Value & "'"
            Dim oRsGetPRDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetPRDetails.DoQuery(GetPRDetails)

            If oRsGetPRDetails.RecordCount > 0 Then
                objMatrix1.Clear()
                oDBs_Details1.Clear()
                objMatrix1.FlushToDataSource()
                Me.SetNewLine(objForm.UniqueID, "3")
                For i As Integer = 1 To oRsGetPRDetails.RecordCount
                    oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, i)
                    oDBs_Details1.SetValue("U_MEBRCD", oDBs_Details1.Offset, "")
                    ''oDBs_Details1.SetValue("U_MEBRNM", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_MEITMCD", oDBs_Details1.Offset, oRsGetPRDetails.Fields.Item("ItemCode").Value)
                    oDBs_Details1.SetValue("U_MEITNM", oDBs_Details1.Offset, oRsGetPRDetails.Fields.Item("Dscription").Value)
                    oDBs_Details1.SetValue("U_MEUOM", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_MEADBRCD", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_MEFWSTK", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_5").Cells.Item(i).Specific.Value)
                    oDBs_Details1.SetValue("U_METWHSTK", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_6").Cells.Item(i).Specific.Value)
                    oDBs_Details1.SetValue("U_MEFWHSCD", oDBs_Details1.Offset, objForm.Items.Item("6").Specific.Value)
                    oDBs_Details1.SetValue("U_METWHSCD", oDBs_Details1.Offset, objForm.Items.Item("8").Specific.Value)
                    oDBs_Details1.SetValue("U_MEREQQTY", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_9").Cells.Item(i).Specific.Value)
                    oDBs_Details1.SetValue("U_MERELQTY", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_10").Cells.Item(i).Specific.Value)
                    oDBs_Details1.SetValue("U_MEOPNQTY", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_9").Cells.Item(i).Specific.Value)
                    oDBs_Details1.SetValue("U_MERCEQTY", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_12").Cells.Item(i).Specific.Value)
                    oDBs_Details1.SetValue("U_MELNSTS", oDBs_Details1.Offset, "Open")
                    oDBs_Details1.SetValue("U_MEPCKNG", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_14").Cells.Item(i).Specific.Value)

                    objMatrix1.SetLineData(i)
                    oRsGetPRDetails.MoveNext()
                    Me.SetNewLine(objForm.UniqueID, "3")
                Next
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub RefreshData(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR_C1")

            objMatrix1 = objForm.Items.Item("3").Specific
            objMatrix2 = objForm.Items.Item("4").Specific


            'objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_STR"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
            '    objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_STR_C0"), "DocEntry", oDBs_Details1.GetValue("DocEntry", 0))
            '    objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_STR_C1"), "DocEntry", oDBs_Details1.GetValue("DocEntry", 0))

            objMatrix1.LoadFromDataSource()
            objMatrix1.AutoResizeColumns()

            objMatrix2.LoadFromDataSource()
            objMatrix2.AutoResizeColumns()
            objForm.Refresh()
            objForm.Freeze(False)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Function HasDuplicates(ByVal arr As Array) As Boolean
        For i As Integer = 0 To arr.Length - 1
            If Not arr(i) Is Nothing Then
                Dim l As Integer = Array.LastIndexOf(arr, arr(i))
                If l <> i Then
                    objMain.objApplication.StatusBar.SetText(" Duplicate item found : " & objMatrix1.Columns.Item("V_0").Cells.Item(l + 1).Specific.Value, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return True
                End If
            End If
        Next
        Return False
    End Function

    Public Sub ModeMask(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR_C1")

            objMatrix1 = objForm.Items.Item("3").Specific
            objMatrix2 = objForm.Items.Item("4").Specific

            Dim objComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("21").Specific
            If objComboBox.Value <> "Close" Then
                For i As Integer = 1 To objMatrix1.VisualRowCount
                    If objMatrix1.Columns.Item("V_1").Cells.Item(i).Specific.Value <> "" And objMatrix1.Columns.Item("V_10").Cells.Item(i).Specific.Value > 0 Then
                        objMatrix1.CommonSetting.SetRowEditable(i, False)
                    ElseIf objMatrix1.Columns.Item("V_1").Cells.Item(i).Specific.Value <> "" Then
                        objMatrix1.CommonSetting.SetRowEditable(i, True)
                        objMatrix1.Columns.Item("V_-1").Editable = False
                        objMatrix1.Columns.Item("V_0").Editable = True
                        objMatrix1.Columns.Item("V_1").Editable = True
                        objMatrix1.Columns.Item("V_2").Editable = False
                        objMatrix1.Columns.Item("V_3").Editable = True
                        objMatrix1.Columns.Item("V_4").Editable = True
                        objMatrix1.Columns.Item("V_5").Editable = False
                        objMatrix1.Columns.Item("V_6").Editable = False
                        objMatrix1.Columns.Item("V_7").Editable = False
                        objMatrix1.Columns.Item("V_8").Editable = False
                        objMatrix1.Columns.Item("V_9").Editable = True
                        objMatrix1.Columns.Item("V_10").Editable = False
                        objMatrix1.Columns.Item("V_11").Editable = False
                        objMatrix1.Columns.Item("V_12").Editable = False
                        objMatrix1.Columns.Item("V_13").Editable = False
                        objMatrix1.Columns.Item("V_14").Editable = False
                    End If
                Next
            End If
            If objComboBox.Value = "Close" Then
                For i As Integer = 1 To objMatrix1.VisualRowCount
                    objMatrix1.CommonSetting.SetRowEditable(i, False)
                Next
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("TNX_POPost", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        objForm.Freeze(True)
                        Me.ModeMask(objForm.UniqueID)
                        objForm.Freeze(False)
                    End If
                    'Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    '    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                    '        Me.ModeMask(objForm.UniqueID)
                    '    End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            Dim objForm As SAPbouiCOM.Form
            Dim oMenuItem As SAPbouiCOM.MenuItem
            Dim oMenus As SAPbouiCOM.Menus
            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
            oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)
            If eventInfo.FormUID = objForm.UniqueID Then
                If (eventInfo.BeforeAction = True) Then
                    If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                        If eventInfo.ItemUID <> "3" Then
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("Close") = False Then
                                    oCreationPackage.UniqueID = "Close"
                                    oCreationPackage.String = "Close"
                                    oCreationPackage.Enabled = True
                                    oMenus.AddEx(oCreationPackage)
                                End If
                            Catch ex As Exception
                            End Try
                        End If
                        objMatrix1 = objForm.Items.Item("3").Specific
                        If eventInfo.ItemUID = "3" And eventInfo.ColUID = "V_-1" Then
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("Delete Row") = False Then
                                    oCreationPackage.UniqueID = "Delete Row"
                                    oCreationPackage.String = "Delete Row"
                                    oCreationPackage.Enabled = True
                                    oMenus.AddEx(oCreationPackage)
                                End If
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        ElseIf eventInfo.ItemUID <> "3" Then
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("Delete Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Delete Row")
                                End If
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        Else
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("Delete Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Delete Row")
                                End If
                                If oMenus.Exists("Close") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Close")
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

    Public Sub GetPDFFile(ByVal ReportName As String, ByVal FileName As String, ByVal Parameter As String)
        Try
            Dim cryRpt As New ReportDocument
            Dim crtableLogoninfos As New TableLogOnInfos
            Dim crtableLogoninfo As New TableLogOnInfo
            Dim crConnectionInfo As New ConnectionInfo
            Dim CrTables As Tables
            Dim CrTable As Table
            Dim CrExportOptions As CrystalDecisions.Shared.ExportOptions
            Dim CrDiskFileDestinationOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions()
            Dim CrFormatTypeOptions As New CrystalDecisions.Shared.PdfRtfWordFormatOptions

            Dim GetDBPassword As String = "Select ""U_MEPSWD"" , ""U_MEATCH"" , ""U_MEPATH"" From ""@ME_ACFG"" "
            Dim oRsGetDBPassword As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDBPassword.DoQuery(GetDBPassword)

            If oRsGetDBPassword.Fields.Item(1).Value = "" Then
                objMain.objApplication.StatusBar.SetText("Please set the Attachment Path in General Settings", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            With crConnectionInfo
                .ServerName = objMain.objCompany.Server
                .DatabaseName = objMain.objCompany.CompanyDB
                .UserID = "SYSTEM"
                .Password = oRsGetDBPassword.Fields.Item("U_MEPSWD").Value.Trim
            End With

            cryRpt.Load(oRsGetDBPassword.Fields.Item("U_MEATCH").Value.Trim & "\" & ReportName)

            CrTables = cryRpt.Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
            Next

            Dim DocNo As String = objForm.Items.Item("18").Specific.value

            CrDiskFileDestinationOptions.DiskFileName = oRsGetDBPassword.Fields.Item("U_MEPATH").Value.Trim & "\" & FileName & "_" & DocNo & ".pdf"

            If Parameter <> "" Then
                cryRpt.SetParameterValue(Parameter, oDBs_Head.GetValue("DocEntry", 0))
            End If

            CrExportOptions = cryRpt.ExportOptions
            With CrExportOptions
                .ExportDestinationType = ExportDestinationType.DiskFile
                .ExportFormatType = ExportFormatType.PortableDocFormat
                .DestinationOptions = CrDiskFileDestinationOptions
                .FormatOptions = CrFormatTypeOptions
            End With
            cryRpt.Export()
            System.Diagnostics.Process.Start(CrDiskFileDestinationOptions.DiskFileName)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub CFLGrPO(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim GetBarCode As String = " Select TOP 100 ""DocNum"" from  OPDN T0  INNER JOIN PDN1 T1 On T0.""DocEntry"" = T1.""DocEntry"" where T0.""CANCELED""='N' and ""DocNum"" not in (Select IFNULL(""U_GRPNO"",'') from ""@TNX_MPSTR"")"
            Dim oRsGetBarCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetBarCode.DoQuery(GetBarCode)
            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            If oRsGetBarCode.RecordCount > 0 Then
                oCondition = oConditions.Add()
                oCondition.Alias = "DocNum"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = oRsGetBarCode.Fields.Item(0).Value
                oChooseFromList.SetConditions(oConditions)
                oRsGetBarCode.MoveNext()
                For i As Integer = 1 To oRsGetBarCode.RecordCount - 1
                    oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCondition = oConditions.Add()
                    oCondition.Alias = "DocNum"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = oRsGetBarCode.Fields.Item(0).Value
                    oChooseFromList.SetConditions(oConditions)
                    oRsGetBarCode.MoveNext()
                Next
            Else
                oCondition = oConditions.Add()
                oCondition.Alias = "DocNum"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = ""
                oChooseFromList.SetConditions(oConditions)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub LoadItemDetails(ByVal whse As String, ByVal grpono As String)

        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR")
        oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR_C0")
        oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_MPSTR_C1")

        objMatrix1 = objForm.Items.Item("3").Specific
        objMatrix2 = objForm.Items.Item("4").Specific
        Dim GetGrpoDtls As String =
        " SELECT T0.""BPLId"",T1.""LineNum"", T1.""ItemCode"", T1.""Dscription"", T1.""UomCode"", T1.""WhsCode"",T1.""CodeBars"",T1.""Quantity"", " &
        " (Select ""UomCode"" FROM OUOM WHERE ""UomEntry"" = I.""IUoMEntry"") As ""IUomCode"", I.""IUoMEntry"", " &
        " (Select IFNULL(W1.""OnHand"",0) FROM OITW W1 WHERE W1.""ItemCode""=I.""ItemCode"" And W1.""WhsCode""='" & objForm.Items.Item("6").Specific.Value & "') AS ""FromWhsStock"", " &
        " (SELECT SUM(IFNULL(W3.""OnHand"",0)) FROM OITW W3 WHERE W3.""ItemCode""=I.""ItemCode"") AS ""OverallStock"" " &
        " FROM OPDN T0 " &
        " INNER JOIN PDN1 T1 ON T0.""DocEntry""=T1.""DocEntry"" " &
        " INNER JOIN OITM I ON I.""ItemCode""=T1.""ItemCode"" " &
        " WHERE  T1.""WhsCode""='" & whse & "' AND T0.""DocNum""='" & grpono & "'"

        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs.DoQuery(GetGrpoDtls)

        If oRs.RecordCount > 0 Then
            objMain.objApplication.StatusBar.SetText("Please wait while loading line details....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            objMatrix1.Clear()
            oDBs_Details1.Clear()
            objMatrix1.FlushToDataSource()
            objMatrix1.AutoResizeColumns()
            For i As Integer = 1 To oRs.RecordCount
                objMain.objApplication.StatusBar.SetText("Loading Stock - ItemCode: " & oRs.Fields.Item("ItemCode").Value & " and record count " & i & " out of " & oRs.RecordCount & ".", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                objMatrix1.AddRow()
                oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
                If oRs.Fields.Item("CodeBars").Value <> "" Then
                    oDBs_Details1.SetValue("U_MEBRCD", oDBs_Details1.Offset, oRs.Fields.Item("CodeBars").Value)
                    oDBs_Details1.SetValue("U_MEADBRCD", oDBs_Details1.Offset, oRs.Fields.Item("CodeBars").Value)
                Else
                    oDBs_Details1.SetValue("U_MEBRCD", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_MEADBRCD", oDBs_Details1.Offset, "")
                End If
                oDBs_Details1.SetValue("U_MEITMCD", oDBs_Details1.Offset, oRs.Fields.Item("ItemCode").Value)
                oDBs_Details1.SetValue("U_MEITNM", oDBs_Details1.Offset, oRs.Fields.Item("Dscription").Value)
                oDBs_Details1.SetValue("U_MEUOM", oDBs_Details1.Offset, oRs.Fields.Item("UomCode").Value)
                oDBs_Details1.SetValue("U_MEFWHSCD", oDBs_Details1.Offset, objForm.Items.Item("6").Specific.Value)
                oDBs_Details1.SetValue("U_MEFWSTK", oDBs_Details1.Offset, oRs.Fields.Item("FromWhsStock").Value)
                oDBs_Details1.SetValue("U_MEREQQTY", oDBs_Details1.Offset, oRs.Fields.Item("Quantity").Value)
                oDBs_Details1.SetValue("U_MERELQTY", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_MEOPNQTY", oDBs_Details1.Offset, oRs.Fields.Item("Quantity").Value)
                oDBs_Details1.SetValue("U_MERCEQTY", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_MELNSTS", oDBs_Details1.Offset, "Open")
                oDBs_Details1.SetValue("U_GENT", oDBs_Details1.Offset, oRs.Fields.Item("Quantity").Value)
                oDBs_Details1.SetValue("U_GLNE", oDBs_Details1.Offset, oRs.Fields.Item("LineNum").Value)

                objMatrix1.SetLineData(objMatrix1.VisualRowCount)

                oRs.MoveNext()
            Next


            objMain.objApplication.StatusBar.SetText("Loaded successfully.....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If

    End Sub
    'Sub UpdateSTR(ByVal FormUID As String, ByVal DocNum As String, ByVal STREntry As String)
    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)

    '        Dim GetDetails As String = "Select DISTINCT T1.""U_ToWhse"" from ""@TNX_WAREHOUSE""T0 Inner join ""@TNX_WAREHOUSE_C0"" T1 on T0.""DocEntry""=T1.""DocEntry""Inner join  ""@TNX_MPSTR""  T2 on T2.""DocEntry""=T0.""U_BEntry"" " &
    '                        "Inner join  ""@TNX_MPSTR_C0""  T3 on T3.""DocEntry""=T2.""DocEntry"" And T3.""LineId""=T0.""U_LineNum""where  T1.""U_Quantity"">0 group by ""U_ToWhse"",""U_ItemCode"" "
    '        Dim oRsWhs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oRsWhs.DoQuery(GetDetails)

    '        If oRsWhs.RecordCount > 0 Then
    '            oRsWhs.MoveFirst()
    '            While Not oRsWhs.EoF
    '                Dim toWhs As String = oRsWhs.Fields.Item("U_ToWhse").Value.ToString()

    '            End While



    '            objMain.sCmp = objMain.objCompany.GetCompanyService
    '                objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OSTR")
    '                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
    '                objMain.oGeneralParams.SetProperty("DocEntry", STREntry)
    '                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
    '                objMain.oChildren = objMain.oGeneralData.Child("ME_STR_C0")

    '                objMain.oChild = objMain.oChildren.Add()


    '                Dim LineId As Integer = CInt(oRsWhs.Fields.Item("LineId").Value)
    '                objMain.oChildren.Item(LineId - 1).SetProperty("U_MERELQTY", oRsWhs.Fields.Item("DispatchQty").Value)
    '                objMain.oChildren.Item(LineId - 1).SetProperty("U_MEOPNQTY", oRsWhs.Fields.Item("OpenQty").Value)

    '                objMain.oChildren.Item(LineId - 1).SetProperty("U_MELNSTS", "Open")

    '                objMain.oGeneralService.Update(objMain.oGeneralData)
    '                oRsGetDetails.MoveNext()


    '                objMain.oGeneralData.SetProperty("U_MESTS", "Open")
    '                Dim oParams As SAPbobsCOM.GeneralDataParams
    '                oParams = oGeneralService.Add(oGeneralData)

    '                Dim newDocEntry As String = oParams.GetProperty("DocEntry").ToString()

    '                ' objMain.oGeneralService.Add(objMain.oGeneralData)
    '                Next
    '        End If
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub
    Sub UpdateSTR(ByVal FormUID As String, ByVal DocNum As String, ByVal STREntry As String)

        Try
            '========================
            ' GET DISTINCT WAREHOUSES
            '========================
            Dim qWhs As String = "
        SELECT DISTINCT T1.""U_ToWhse"",T2.""U_MEFWHSCD"" as ""FromWhse"",T0.""U_Branch"",T0.""U_FWhse"",T2.""U_MEFWHSNM"" as ""FromWhseNM"",T3.""U_MEADBRCD""
        FROM ""@TNX_WAREHOUSE"" T0
        INNER JOIN ""@TNX_WAREHOUSE_C0"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
        INNER JOIN ""@TNX_MPSTR"" T2 ON T2.""DocEntry"" = T0.""U_BEntry""
        INNER JOIN ""@TNX_MPSTR_C0"" T3 
            ON T3.""DocEntry"" = T2.""DocEntry"" 
            AND T3.""LineId"" = T0.""U_LineNum""
        WHERE T1.""U_Quantity"" > 0 and ""U_BaseNum""='" & DocNum & "' and IFNULL(T1.""U_STR"",'') = ''
        "

            Dim oRsWhs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsWhs.DoQuery(qWhs)

            If oRsWhs.RecordCount = 0 Then
                Throw New Exception("No data found")
            End If

            oRsWhs.MoveFirst()

            '========================
            ' LOOP EACH WAREHOUSE
            '========================
            While Not oRsWhs.EoF

                Dim Branch As String = "Select T0.""BPLid"",T1.""BPLName"" from ""OWHS"" T0 Left Join ""OBPL"" T1 ON T0.""BPLid"" = ""BPLId"" where ""WhsCode""='" & oRsWhs.Fields.Item("U_ToWhse").Value.ToString() & "'"
                Dim oRsBranch As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsBranch.DoQuery(Branch)

                Dim Series As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'ME_OSTR' And ""Locked"" = 'N' AND ""Indicator"" <> 'Default'"
                Dim oRsSeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsSeries.DoQuery(Series)

                Dim toWhs As String = oRsWhs.Fields.Item("U_ToWhse").Value.ToString()
                Dim FrmWhsName As String = oRsWhs.Fields.Item("FromWhseNM").Value.ToString()
                Dim FrmWhsCode As String = oRsWhs.Fields.Item("FromWhse").Value.ToString()
                Dim FrmBranch As String = oRsWhs.Fields.Item("U_Branch").Value.ToString()
                Dim ToBranch As String = oRsBranch.Fields.Item("BPLName").Value.ToString()

                '========================
                ' CREATE GENERAL SERVICE
                '========================
                Dim oCompanyService As SAPbobsCOM.CompanyService = objMain.objCompany.GetCompanyService()
                Dim oGeneralService As SAPbobsCOM.GeneralService = oCompanyService.GetGeneralService("ME_OSTR")

                Dim oGeneralData As SAPbobsCOM.GeneralData
                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                '========================
                ' HEADER
                '========================
                oGeneralData.SetProperty("U_TYPE", "GR")
                oGeneralData.SetProperty("U_METBPID", ToBranch)
                oGeneralData.SetProperty("U_MESRS", oRsSeries.Fields.Item(1).Value)
                oGeneralData.SetProperty("U_MEFBPID", FrmBranch)
                oGeneralData.SetProperty("U_MEFWHSCD", FrmWhsCode)   ' From Warehouse
                oGeneralData.SetProperty("U_METWHSCD", toWhs)       ' To Warehouse


                Dim GETWNAME As String = " Select ""WhsName"" From OWHS Where ""WhsCode"" = '" & toWhs & "'"
                Dim oRsGETWNAME As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGETWNAME.DoQuery(GETWNAME)

                Dim ToWhsName As String = oRsGETWNAME.Fields.Item("WhsName").Value.ToString()


                oGeneralData.SetProperty("U_MEFWHSNM", FrmWhsName)   ' From Warehouse
                oGeneralData.SetProperty("U_METWHSNM", ToWhsName)


                oGeneralData.SetProperty("U_MESTS", "Open")
                oGeneralData.SetProperty("U_MEPSDT", DateTime.Now)
                oGeneralData.SetProperty("U_MEDOCDT", DateTime.Now)
                oGeneralData.SetProperty("U_MERMKS", "Created From Multiple STR")

                '========================
                ' CHILD COLLECTION
                '========================
                Dim oChildren As SAPbobsCOM.GeneralDataCollection
                oChildren = oGeneralData.Child("ME_STR_C0")

                '========================
                ' GET ITEMS FOR THIS WAREHOUSE
                '========================
                Dim qItems As String = "
            SELECT 
                T0.""U_BarCode"",
                T0.""U_ItemCode"",T0.""U_ItemName"",T0.""U_UOM"",
                SUM(T1.""U_Quantity"") AS ""Qty"",T3.""U_MEADBRCD"",
(SELECT MAX(X.""BcdName"") FROM OBCD X WHERE X.""BcdCode"" = T0.""U_BarCode"") AS ""BarcodeName""
            FROM ""@TNX_WAREHOUSE"" T0
            INNER JOIN ""@TNX_WAREHOUSE_C0"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
            INNER JOIN ""@TNX_MPSTR"" T2 ON T2.""DocEntry"" = T0.""U_BEntry""
            INNER JOIN ""@TNX_MPSTR_C0"" T3 
                ON T3.""DocEntry"" = T2.""DocEntry"" 
                AND T3.""LineId"" = T0.""U_LineNum""
            WHERE 
                T1.""U_Quantity"" > 0
                AND T1.""U_ToWhse"" = '" & toWhs & "' and  ""U_BaseNum""='" & DocNum & "' and IFNULL(T1.""U_STR"",'') = ''
            GROUP BY 
                T0.""U_BarCode"",
                T0.""U_ItemCode"",T0.""U_ItemName"",T0.""U_UOM"",T3.""U_MEADBRCD""
            "

                Dim oRsItems As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsItems.DoQuery(qItems)

                oRsItems.MoveFirst()

                '========================
                ' LOOP ITEMS → CHILD ROWS
                '========================
                While Not oRsItems.EoF

                    Dim oChild As SAPbobsCOM.GeneralData = oChildren.Add()

                    Dim qty As Double = CDbl(oRsItems.Fields.Item("Qty").Value)

                    'Dim ChkInventoryUOM As String = " Select ""UomCode"" From OUOM Where ""UomEntry"" = '" & oRsGetStockDtls.Fields.Item("IUoMEntry").Value & "' And ""UomCode"" = '" & objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value & "'  "
                    'Dim oRsChkInventoryUOM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'oRsChkInventoryUOM.DoQuery(ChkInventoryUOM)

                    oChild.SetProperty("U_MEBRCD", oRsItems.Fields.Item("U_BarCode").Value)
                    oChild.SetProperty("U_MEBRNM", oRsItems.Fields.Item("BarcodeName").Value)
                    oChild.SetProperty("U_MEITMCD", oRsItems.Fields.Item("U_ItemCode").Value)
                    oChild.SetProperty("U_MEITNM", oRsItems.Fields.Item("U_ItemName").Value)
                    oChild.SetProperty("U_MEUOM", oRsItems.Fields.Item("U_UOM").Value)
                    oChild.SetProperty("U_MEADBRCD", oRsItems.Fields.Item("U_MEADBRCD").Value)
                    oChild.SetProperty("U_MEFWHSCD", FrmWhsCode)
                    oChild.SetProperty("U_METWHSCD", toWhs)


                    Dim GetFrom As String = "Select IFNULL(W1.""OnHand"",0) ""FWHS"" FROM OITW W1 WHERE W1.""ItemCode""='" & oRsItems.Fields.Item("U_ItemCode").Value & "' And W1.""WhsCode""='" & FrmWhsCode & "'"

                    Dim oRsGetFrom As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetFrom.DoQuery(GetFrom)

                    Dim GetTo As String = "Select IFNULL(W1.""OnHand"",0) ""TWHS"" FROM OITW W1 WHERE W1.""ItemCode""='" & oRsItems.Fields.Item("U_ItemCode").Value & "' And W1.""WhsCode""='" & toWhs & "'"

                    Dim oRsGetTo As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetTo.DoQuery(GetTo)
                    Dim FromS As Double = CDbl(oRsGetFrom.Fields.Item("FWHS").Value)
                    Dim ToS As Double = CDbl(oRsGetTo.Fields.Item("TWHS").Value)
                    oChild.SetProperty("U_MEFWSTK", FromS.ToString())
                    oChild.SetProperty("U_METWHSTK", ToS.ToString())

                    oChild.SetProperty("U_MEREQQTY", qty)
                    oChild.SetProperty("U_MERELQTY", 0)
                    oChild.SetProperty("U_MEOPNQTY", qty)

                    oChild.SetProperty("U_MELNSTS", "Open")

                    oRsItems.MoveNext()
                End While

                '========================
                ' ADD STR DOCUMENT
                '========================
                Dim oParams As SAPbobsCOM.GeneralDataParams = oGeneralService.Add(oGeneralData)
                Dim docEntry As String = oParams.GetProperty("DocEntry").ToString()

                objMain.objApplication.StatusBar.SetText("STR Created for " & toWhs & " : " & docEntry,
                                                        SAPbouiCOM.BoMessageTime.bmt_Short,
                                                        SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                Dim Update As String = " UPDATE ""@TNX_WAREHOUSE_C0"" T1
SET T1.""U_STR"" = '" & docEntry & "'   
WHERE EXISTS (
    SELECT 1 
    FROM ""@TNX_WAREHOUSE"" T0
    INNER JOIN ""@TNX_MPSTR"" T2 
        ON T2.""DocEntry"" = T0.""U_BEntry""
    INNER JOIN ""@TNX_MPSTR_C0"" T3 
        ON T3.""DocEntry"" = T2.""DocEntry"" 
        AND T3.""LineId"" = T0.""U_LineNum""
    WHERE T0.""DocEntry"" = T1.""DocEntry""
        AND T1.""U_Quantity"" > 0 AND T1.""U_ToWhse"" = '" & toWhs & "' AND T0.""U_BaseNum"" = '" & DocNum & "' and IFNULL(T1.""U_STR"",'') = '')"

                Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsUpdate.DoQuery(Update)


                oRsWhs.MoveNext()
            End While

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub

End Class

