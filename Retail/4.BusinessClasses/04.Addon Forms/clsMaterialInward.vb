Imports System.Threading
Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class clsMaterialInward

#Region "       Declaration             "
    Public objForm, objSTRForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details1 As SAPbouiCOM.DBDataSource
    Dim objMatrix1 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim BaseUom As String
    Dim GetInTransitWhs As String
    Dim oRsGetInTransitWhs As SAPbobsCOM.Recordset
    Dim TRANSFLAG As String = ""
    Dim oRsTRANSFLAG As SAPbobsCOM.Recordset
#End Region

    Sub CreateForm()
        Try
            GetInTransitWhs = "SELECT *  FROM ""@ME_ACFG""  T0 WHERE T0.""U_INTRW"" ='N' or T0.""U_INTRW"" ='Y'"
            oRsGetInTransitWhs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetInTransitWhs.DoQuery(GetInTransitWhs)
            If oRsGetInTransitWhs.RecordCount > 0 Then
                '    objMain.objUtilities.LoadForm("MaterialInwardGI2.xml", "ME_MI_Form", ResourceType.Embeded)
                'Else
                objMain.objUtilities.LoadForm("MaterialInward1.xml", "ME_MI_Form", ResourceType.Embeded)
            End If

            objForm = objMain.objApplication.Forms.GetForm("ME_MI_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_MI")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_MI_C0")
            objMatrix1 = objForm.Items.Item("3").Specific

            objMatrix1.Columns.Item("V_16").Visible = False
            objMatrix1.Columns.Item("V_17").Visible = False
            objMatrix1.Columns.Item("V_18").Visible = False
            objMatrix1.Columns.Item("V_19").Visible = False
            objMatrix1.Columns.Item("V_20").Visible = False

            objComboBox = objForm.Items.Item("53").Specific
            objComboBox.ValidValues.Add("", "")
            objComboBox.ValidValues.Add("Open", "Open")
            objComboBox.ValidValues.Add("Close", "Close")
            objComboBox.ValidValues.Add("Cancel", "Cancel")
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("53").DisplayDesc = True
            objForm.Items.Item("21").Visible = False
            Dim sQury As String = "Select ""BPLId"" ,  ""BPLName"" From OBPL  "
            objComboBox = objForm.Items.Item("40").Specific
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, sQury)
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("40").DisplayDesc = True

            objMain.objUtilities.MatrixComboBoxValues(objMatrix1.Columns.Item("V_11"), sQury)
            objMatrix1.Columns.Item("V_11").DisplayDesc = True

            objComboBox = objForm.Items.Item("42").Specific
            objMain.objUtilities.ComboBoxLoadValues(objForm.Items.Item("42").Specific, "Select ""BPLId"" ,  ""BPLName"" From OBPL  ")
            objForm.Items.Item("42").DisplayDesc = True

            Me.CellsMasking(objForm.UniqueID)

            Me.CFLEmpFilter(objForm.UniqueID, "CFL_EMPID")

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ME_MI" And pVal.BeforeAction = False Then
                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "519" And pVal.BeforeAction = True Then 'Preview
                Try
                    Dim frs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'Material Inward' AND ""Type"" = 'C'")
                    If frs.RecordCount = 0 Then
                        objMain.objApplication.MessageBox("Report layout 'Material Inward' not found. Please upload with same name and Code as parameter", 0, "OK")
                    Else
                        objMain.objApplication.ActivateMenuItem(frs.Fields.Item(0).Value.ToString())
                        Dim CrForm As SAPbouiCOM.Form
                        Dim oedt As SAPbouiCOM.EditText
                        CrForm = objMain.objApplication.Forms.ActiveForm
                        oedt = CrForm.Items.Item("1000003").Specific
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_MI")
                        oedt.Value = oDBs_Head.GetValue("DocEntry", 0)
                        CrForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                Catch ex As Exception
                End Try
            ElseIf pVal.MenuUID = "520" And pVal.BeforeAction = True Then 'Print
                Try
                    Dim frs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'Material Inward' AND ""Type"" = 'C'")
                    If frs.RecordCount = 0 Then
                        objMain.objApplication.MessageBox("Report layout 'Material Inward' not found. Please upload with same name and Code as parameter", 0, "OK")
                    Else
                        objMain.objApplication.ActivateMenuItem(frs.Fields.Item(0).Value.ToString())
                        Dim CrForm As SAPbouiCOM.Form
                        Dim oedt As SAPbouiCOM.EditText
                        CrForm = objMain.objApplication.Forms.ActiveForm
                        oedt = CrForm.Items.Item("1000003").Specific
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_MI")
                        oedt.Value = oDBs_Head.GetValue("DocEntry", 0)
                        CrForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                Catch ex As Exception
                End Try
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub CellsMasking(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Items.Item("53").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("53").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("28").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("28").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("29").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("29").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("30").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("30").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("17").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("17").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("15").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("15").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("15").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("33").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("33").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("34").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("34").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("35").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("35").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("36").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("36").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("42").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("42").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("45").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("45").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("46").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("46").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("24").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("24").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("40").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("40").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("36").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("36").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("38").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("38").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("45").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("45").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "", Optional ByVal Series As Integer = 0)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_MI")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_MI_C0")
            objMatrix1 = objForm.Items.Item("3").Specific

            objComboBox = objForm.Items.Item("53").Specific
            objComboBox.Select("Open", SAPbouiCOM.BoSearchKey.psk_ByValue)

            objComboBox = objForm.Items.Item("52").Specific
            If (objComboBox.ValidValues.Count <> 0) Then
                For R As Integer = objComboBox.ValidValues.Count - 1 To 0 Step -1
                    Try
                        objComboBox.ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                    Catch ex As Exception
                    End Try
                Next
            End If

            'objComboBox.ValidValues.Add("ST", "Standard")
            objComboBox.ValidValues.Add("DP", "Dispatch")
            objForm.Items.Item("52").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objComboBox.Select("DP", SAPbouiCOM.BoSearchKey.psk_ByValue)

            objComboBox = objForm.Items.Item("15").Specific
            Dim GetSeriesQty As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'ME_OMI' And ""Locked"" = 'N'"
            objMain.objUtilities.ComboBoxLoadValuesSeries(objForm.Items.Item("15").Specific, GetSeriesQty)
            objForm.Items.Item("15").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            Dim GetDflt As String = "Select ""DfltSeries"" From ONNM Where ""ObjectCode"" = 'ME_OMI'"
            Dim oRsGetDflt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDflt.DoQuery(GetDflt)

            If Flag = "" Then
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNumSeries(objForm, "ME_OMI", oRsGetDflt.Fields.Item(0).Value))
                oDBs_Head.SetValue("U_MESRS", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)
                oDBs_Head.SetValue("Series", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)
                oDBs_Head.SetValue("U_MEDATE", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
            Else
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNumSeries(objForm, "ME_OMI", Series))
                oDBs_Head.SetValue("U_MESRS", oDBs_Head.Offset, Series)
                oDBs_Head.SetValue("Series", oDBs_Head.Offset, Series)
                oDBs_Head.SetValue("U_MESTS", oDBs_Head.Offset, "Open")
                oDBs_Head.SetValue("U_MEDATE", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
            End If

            Dim GetRequesterDtls As String = "Select T1.""empID"" From OUSR T0 Inner Join OHEM T1 On T0.""USERID""  = T1.""userId""  Where ""USER_CODE"" = '" & objMain.objCompany.UserName & "'"
            Dim oRsGetRequesterDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetRequesterDtls.DoQuery(GetRequesterDtls)
            If oRsGetRequesterDtls.RecordCount > 0 Then
                Dim GetEmployeeDtls As String = ""
                If Sql = "Sql" Then
                    GetEmployeeDtls = "Select T0.""firstName"" + ' ' + T0.""lastName"" AS ""Name""  FROM OHEM T0  Inner Join OUDP T1 ON T0.""dept"" = T1.""Code"" LEFT JOIN ""OBPL"" T2 ON T0.""BPLId""=T2.""BPLId"" Where T0.""empID""='" & oRsGetRequesterDtls.Fields.Item("empID").Value & "'"
                ElseIf Hana = "Hana" Then
                    GetEmployeeDtls = "Select T0.""firstName"" ||' '|| T0.""lastName"" AS ""Name"" FROM OHEM T0  Inner Join OUDP T1 ON T0.""dept"" = T1.""Code"" LEFT JOIN ""OBPL"" T2 ON T0.""BPLId""=T2.""BPLId"" Where T0.""empID""='" & oRsGetRequesterDtls.Fields.Item("empID").Value & "'"
                End If
                Dim oRsGetEmployeeDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetEmployeeDtls.DoQuery(GetEmployeeDtls)
                oDBs_Head.SetValue("U_MEREMPID", oDBs_Head.Offset, oRsGetRequesterDtls.Fields.Item("empID").Value)
                oDBs_Head.SetValue("U_MERECBY", oDBs_Head.Offset, oRsGetEmployeeDtls.Fields.Item(0).Value)
            End If

            objMatrix1.AutoResizeColumns()


            objMatrix1.Clear()
            oDBs_Details1.Clear()
            objMatrix1.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID)


            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_MI")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_MI_C0")
            objMatrix1 = objForm.Items.Item("3").Specific

            objMatrix1.AddRow()
            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
            oDBs_Details1.SetValue("U_MEBRCD", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_MEITMCD", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_MEITDSC", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_MEUOM", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_ADBRCD", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_MEPKGE", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_METYPE", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_MESRBT", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_CHKBIN", oDBs_Details1.Offset, "N")
            oDBs_Details1.SetValue("U_METWHS", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_METWBL", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_METBRCH", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_MEREQQTY", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_MERECQTY", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_MEVARQTY", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_MESTRNO", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_MESTREN", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_MEITRNO", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_MEITREN", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_MEBSTYP", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_MEBLN", oDBs_Details1.Offset, "")
            objMatrix1.SetLineData(objMatrix1.VisualRowCount)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    'If pVal.ItemUID = "1" And pVal.Before_Action = True Then
                    '    objForm = objMain.objApplication.Forms.Item(FormUID)
                    '    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_MI")
                    '    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_MI_C0")
                    '    objMatrix1 = objForm.Items.Item("3").Specific
                    '    Dim STRNo As String = ""
                    '    If objMatrix1.RowCount > 0 And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    '        STRNo = objMatrix1.Columns.Item("V_15").Cells.Item(1).Specific.Value
                    '        For i As Integer = 1 To objMatrix1.RowCount
                    '            If STRNo <> objMatrix1.Columns.Item("V_15").Cells.Item(i).Specific.Value Then
                    '                objMain.objApplication.StatusBar.SetText("STRNumber is different for line ID : " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '                BubbleEvent = False
                    '                Exit For
                    '            End If
                    '        Next
                    '    End If


                    'End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_MI")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_MI_C0")
                    objMatrix1 = objForm.Items.Item("3").Specific
                    If pVal.ItemUID = "3" And pVal.ColUID = "V_13" And pVal.BeforeAction = False Then
                        'If (objForm.Items.Item("9").Specific.value <> "") Then
                        oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, pVal.Row)
                        oDBs_Details1.SetValue("U_MEBRCD", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MEITMCD", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MEITDSC", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MEUOM", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_ADBRCD", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MEPKGE", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_METYPE", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MESRBT", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_CHKBIN", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Checked)
                        oDBs_Details1.SetValue("U_METWHS", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_9").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_METWBL", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_10").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_METBRCH", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_11").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MEREQQTY", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_12").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MERECQTY", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_13").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MEVARQTY", oDBs_Details1.Offset, CDbl(objMatrix1.Columns.Item("V_12").Cells.Item(pVal.Row).Specific.Value) - CDbl(objMatrix1.Columns.Item("V_13").Cells.Item(pVal.Row).Specific.Value))
                        oDBs_Details1.SetValue("U_MESTRNO", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_15").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MESTREN", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_16").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MEITRNO", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_17").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MEITREN", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_18").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MEBSTYP", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_19").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MEBLN", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_20").Cells.Item(pVal.Row).Specific.Value)
                        objMatrix1.SetLineData(pVal.Row)
                        'End If
                    End If



                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_MI")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_MI_C0")
                    objMatrix1 = objForm.Items.Item("3").Specific

                    If pVal.ItemUID = "3" And pVal.ColUID = "V_13" And pVal.BeforeAction = True Then
                        GetInTransitWhs = "SELECT *  FROM ""@ME_ACFG""  T0 WHERE T0.""U_INTRW"" ='N' "
                        oRsGetInTransitWhs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetInTransitWhs.DoQuery(GetInTransitWhs)
                        If oRsGetInTransitWhs.RecordCount > 0 Then
                            'If objForm.Items.Item("9").Specific.value <> "" Then
                            '    BubbleEvent = False
                            '    objMain.objApplication.StatusBar.SetText("Qty modification is not permitted when document is under Issued ....!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            'End If
                        Else
                            If objForm.Items.Item("9").Specific.value <> "" Then
                                BubbleEvent = False
                                objMain.objApplication.StatusBar.SetText("Qty modification is not permitted when document is under Issued ....!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If
                        End If
                    End If


                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_MI")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_MI_C0")
                    objMatrix1 = objForm.Items.Item("3").Specific
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim LoadSameForm As String = "Select MAX(""DocNum"") FROM ""@ME_MI"""
                        Dim oRsLoadSameForm As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsLoadSameForm.DoQuery(LoadSameForm)
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objForm.Items.Item("13").Specific.Value = oRsLoadSameForm.Fields.Item(0).Value
                        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.Validation(objForm.UniqueID) = False Then BubbleEvent = False
                        ' Me.GoodsReceipt(objForm.UniqueID)
                    End If
                    If pVal.ItemUID = "Goods" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        If objForm.Items.Item("11").Specific.Value <> "" Then
                            objMain.objApplication.StatusBar.SetText("Goods Receipt is already generated ....!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            Me.GoodsReceipt(objForm.UniqueID)
                        End If

                    End If
                    If pVal.ItemUID = "24" And pVal.BeforeAction = False Then
                        ''"SELECT *  FROM ""@ME_ACFG""  T0 WHERE T0.""U_INTRW"" ='N' "
                        GetInTransitWhs = "SELECT *  FROM ""@ME_ACFG""  T0 WHERE T0.""U_INTRW"" ='N' "
                        oRsGetInTransitWhs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetInTransitWhs.DoQuery(GetInTransitWhs)
                        If oRsGetInTransitWhs.RecordCount > 0 Then
                            If objForm.Items.Item("27").Specific.Value = "" Then
                                objMain.objApplication.StatusBar.SetText("Please select receiver details...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Select
                            End If
                            If objForm.Items.Item("5").Specific.Value <> "" Then
                                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    Dim ChkDispatch As String = "Select Count(*) From ""@ME_MI"" Where ""U_MEITREN"" = '" & objForm.Items.Item("46").Specific.value & "'  And ""CreateDate"" >='20220420'"
                                    Dim oRsChkDispatch As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRsChkDispatch.DoQuery(ChkDispatch)
                                    If oRsChkDispatch.Fields.Item(0).Value > 0 Then
                                        objMain.objApplication.StatusBar.SetText("Selected dispatch already closed.....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Exit Select
                                    End If
                                End If
                            End If
                            Me.LoadMatrixGI(objForm.UniqueID)

                        Else
                            If objForm.Items.Item("9").Specific.Value <> "" Then
                                objMain.objApplication.StatusBar.SetText("Issue posted...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Select
                            End If
                            If objForm.Items.Item("27").Specific.Value = "" Then
                                objMain.objApplication.StatusBar.SetText("Please select receiver details...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Select
                            End If
                            If objForm.Items.Item("5").Specific.Value <> "" Then
                                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    Dim ChkDispatch As String = "Select Count(*) From ""@ME_MI"" Where ""U_MEITREN"" = '" & objForm.Items.Item("46").Specific.value & "'"
                                    Dim oRsChkDispatch As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRsChkDispatch.DoQuery(ChkDispatch)
                                    If oRsChkDispatch.Fields.Item(0).Value > 0 Then
                                        objMain.objApplication.StatusBar.SetText("Selected dispatch already closed.....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Exit Select
                                    End If
                                End If
                            End If
                            Me.LoadMatrix(objForm.UniqueID)

                        End If

                    End If
                    If pVal.ItemUID = "1000002" And pVal.BeforeAction = False Then
                        Me.GetPDFFile("Transfer Request SD.rpt", "Material Inward", "DocKey@")
                    End If
                    If pVal.ItemUID = "Barcode" And pVal.BeforeAction = False Then
                        objMain.objBarcodeMaterailInwards.CreateForm(oDBs_Head.GetValue("DocNum", 0))
                    End If
                    If pVal.ItemUID = "49" And pVal.BeforeAction = False Then
                        Me.PressedOnLinked(objForm.UniqueID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_MI")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_MI_C0")
                    objMatrix1 = objForm.Items.Item("3").Specific
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

                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        If oCFL.UniqueID = "CFL_STR" Then
                            oDBs_Head.SetValue("U_MESTRNO", oDBs_Head.Offset, oDT.GetValue("DocNum", 0))
                            oDBs_Head.SetValue("U_MESTREN", oDBs_Head.Offset, oDT.GetValue("DocEntry", 0))
                            oDBs_Head.SetValue("U_MEFBPID", oDBs_Head.Offset, oDT.GetValue("U_MEFBPID", 0))
                            oDBs_Head.SetValue("U_MEFWHS", oDBs_Head.Offset, oDT.GetValue("U_MEFWHSCD", 0))
                            oDBs_Head.SetValue("U_MEFWNM", oDBs_Head.Offset, oDT.GetValue("U_MEFWHSNM", 0))
                            oDBs_Head.SetValue("U_MEFWNM", oDBs_Head.Offset, oDT.GetValue("U_MEFWHSNM", 0))
                            GetInTransitWhs = "Select ""U_INTRWHS"" , (Select ""WhsName"" From OWHS Where ""WhsCode"" = A.""U_INTRWHS"") As ""TrnsWhsName"" From ""OWHS"" A  Where ""U_INTRWHS"" <> '' And ""WhsCode"" = '" & oDT.GetValue("U_MEFWHSCD", 0) & "' "
                            oRsGetInTransitWhs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetInTransitWhs.DoQuery(GetInTransitWhs)
                            If oRsGetInTransitWhs.RecordCount > 0 Then
                                oDBs_Head.SetValue("U_METRWHS", oDBs_Head.Offset, oRsGetInTransitWhs.Fields.Item("U_INTRWHS").Value)
                                oDBs_Head.SetValue("U_METRWNM", oDBs_Head.Offset, oRsGetInTransitWhs.Fields.Item("TrnsWhsName").Value)
                            End If
                            ' Me.CFLFilterITR(objForm.UniqueID, "CFL_IT", oDT.GetValue("DocNum", 0))
                        End If
                        If oCFL.UniqueID = "CFL_IT" Then
                            'Dim GetInTransitWhs As String
                            'Dim oRsGetInTransitWhs As SAPbobsCOM.Recordset

                            GetInTransitWhs = "SELECT *  FROM ""@ME_ACFG""  T0 WHERE T0.""U_INTRW"" ='N' "
                            oRsGetInTransitWhs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetInTransitWhs.DoQuery(GetInTransitWhs)
                            If oRsGetInTransitWhs.RecordCount > 0 Then
                                oDBs_Head.SetValue("U_MEITRNO", oDBs_Head.Offset, oDT.GetValue("DocNum", 0))
                                oDBs_Head.SetValue("U_MEITREN", oDBs_Head.Offset, oDT.GetValue("DocEntry", 0))
                                oDBs_Head.SetValue("U_MEGINO", oDBs_Head.Offset, oDT.GetValue("DocNum", 0))
                                oDBs_Head.SetValue("U_MEGIEN", oDBs_Head.Offset, oDT.GetValue("DocEntry", 0))

                            Else
                                oDBs_Head.SetValue("U_MEITRNO", oDBs_Head.Offset, oDT.GetValue("DocNum", 0))
                                oDBs_Head.SetValue("U_MEITREN", oDBs_Head.Offset, oDT.GetValue("DocEntry", 0))

                            End If

                            oDBs_Head.SetValue("U_MEFBPID", oDBs_Head.Offset, oDT.GetValue("BPLId", 0))
                            ' oDBs_Head.SetValue("U_MEFWHS", oDBs_Head.Offset, oDT.GetValue("Filler", 0))
                            Dim GetWhsName As String = " Select ""WhsName"" From OWHS Where ""WhsCode"" = '" & oDT.GetValue("Filler", 0) & "' "
                            Dim oRsGetWhsName As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetWhsName.DoQuery(GetWhsName)
                            oDBs_Head.SetValue("U_MEFWNM", oDBs_Head.Offset, oRsGetWhsName.Fields.Item("WhsName").Value)


                            TRANSFLAG = "SELECT *  FROM ""@ME_ACFG""  T0 WHERE T0.""U_INTRW"" ='N' "
                            oRsTRANSFLAG = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsTRANSFLAG.DoQuery(TRANSFLAG)
                            If oRsTRANSFLAG.RecordCount > 0 Then
                                Dim GetDispatchLines As String = "Select Top 1 ""U_METBPID"" ,  ""U_METWHSCD"" , ""U_METWHSNM"" , ""WhsCode"" From IGE1 Where ""DocEntry"" = '" & oDT.GetValue("DocEntry", 0) & "' "
                                Dim oRsGetDispatchLines As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetDispatchLines.DoQuery(GetDispatchLines)

                                oDBs_Head.SetValue("U_METBPID", oDBs_Head.Offset, objMain.objUtilities.GetBinID(oRsGetDispatchLines.Fields.Item("U_METBPID").Value))
                                oDBs_Head.SetValue("U_METWHS", oDBs_Head.Offset, oRsGetDispatchLines.Fields.Item("U_METWHSCD").Value)
                                oDBs_Head.SetValue("U_METWHNM", oDBs_Head.Offset, oRsGetDispatchLines.Fields.Item("U_METWHSNM").Value)
                                oDBs_Head.SetValue("U_MEFWHS", oDBs_Head.Offset, oRsGetDispatchLines.Fields.Item("WhsCode").Value)

                                GetInTransitWhs = "Select ""U_INTRWHS"" , (Select ""WhsName"" From OWHS Where ""WhsCode"" = A.""U_INTRWHS"") As ""TrnsWhsName"" From ""OWHS"" A  Where ""U_INTRWHS"" <> '' And ""WhsCode"" = '" & oDT.GetValue("Filler", 0) & "' "
                                oRsGetInTransitWhs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetInTransitWhs.DoQuery(GetInTransitWhs)
                                If oRsGetInTransitWhs.RecordCount > 0 Then
                                    oDBs_Head.SetValue("U_METRWHS", oDBs_Head.Offset, oRsGetInTransitWhs.Fields.Item("U_INTRWHS").Value)
                                    oDBs_Head.SetValue("U_METRWNM", oDBs_Head.Offset, oRsGetInTransitWhs.Fields.Item("TrnsWhsName").Value)
                                End If
                            Else
                                Dim GetDispatchLines As String = "Select Top 1 ""U_METBPID"" ,  ""U_METWHSCD"" , ""U_METWHSNM"" , ""WhsCode"" From WTR1 Where ""DocEntry"" = '" & oDT.GetValue("DocEntry", 0) & "' "
                                Dim oRsGetDispatchLines As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetDispatchLines.DoQuery(GetDispatchLines)

                                oDBs_Head.SetValue("U_METBPID", oDBs_Head.Offset, objMain.objUtilities.GetBinID(oRsGetDispatchLines.Fields.Item("U_METBPID").Value))
                                oDBs_Head.SetValue("U_METWHS", oDBs_Head.Offset, oRsGetDispatchLines.Fields.Item("U_METWHSCD").Value)
                                oDBs_Head.SetValue("U_METWHNM", oDBs_Head.Offset, oRsGetDispatchLines.Fields.Item("U_METWHSNM").Value)
                                oDBs_Head.SetValue("U_MEFWHS", oDBs_Head.Offset, oRsGetDispatchLines.Fields.Item("WhsCode").Value)

                                GetInTransitWhs = "Select ""U_INTRWHS"" , (Select ""WhsName"" From OWHS Where ""WhsCode"" = A.""U_INTRWHS"") As ""TrnsWhsName"" From ""OWHS"" A  Where ""U_INTRWHS"" <> '' And ""WhsCode"" = '" & oDT.GetValue("Filler", 0) & "' "
                                oRsGetInTransitWhs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetInTransitWhs.DoQuery(GetInTransitWhs)
                                If oRsGetInTransitWhs.RecordCount > 0 Then
                                    oDBs_Head.SetValue("U_METRWHS", oDBs_Head.Offset, oRsGetInTransitWhs.Fields.Item("U_INTRWHS").Value)
                                    oDBs_Head.SetValue("U_METRWNM", oDBs_Head.Offset, oRsGetInTransitWhs.Fields.Item("TrnsWhsName").Value)
                                End If
                            End If

                        End If
                        If oCFL.UniqueID = "CFL_TWHS" Then
                            oDBs_Head.SetValue("U_METWHS", oDBs_Head.Offset, oDT.GetValue("WhsCode", 0))
                            oDBs_Head.SetValue("U_METWHNM", oDBs_Head.Offset, oDT.GetValue("WhsName", 0))
                            objComboBox = objForm.Items.Item("40").Specific
                            'Me.CFLFilterSTR(objForm.UniqueID, "CFL_STR", objComboBox.Value)
                        End If

                        If oCFL.UniqueID = "CFL_FWH" Then
                            oDBs_Head.SetValue("U_MEFWHS", oDBs_Head.Offset, oDT.GetValue("WhsCode", 0))
                            oDBs_Head.SetValue("U_MEFWNM", oDBs_Head.Offset, oDT.GetValue("WhsName", 0))
                            oDBs_Head.SetValue("U_MEFBPID", oDBs_Head.Offset, oDT.GetValue("BPLid", 0))
                        End If

                        If oCFL.UniqueID = "CFL_EMPID" Then
                            Dim GetEmployeeDtls As String = "Select T0.""firstName"" ||' '|| T0.""lastName"" AS ""Name"", T1.""Name"" AS ""Dept"",ifnull(T2.""BPLName"",'') AS ""Branch"",ifnull(T0.""CostCenter"",'') AS ""CostCenter"" FROM OHEM T0  INNER JOIN OUDP T1 ON T0.""dept"" = T1.""Code"" LEFT JOIN ""OBPL"" T2 ON T0.""BPLId""=T2.""BPLId"" Where T0.""empID""='" & oDT.GetValue("empID", 0) & "'"
                            Dim oRsGetEmployeeDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetEmployeeDtls.DoQuery(GetEmployeeDtls)
                            oDBs_Head.SetValue("U_MEREMPID", oDBs_Head.Offset, oDT.GetValue("empID", 0))
                            oDBs_Head.SetValue("U_MERECBY", oDBs_Head.Offset, oRsGetEmployeeDtls.Fields.Item(0).Value)
                        End If

                        If oCFL.UniqueID = "CFL_ITMCD" Then
                            Dim PCPrPack As String = objMain.objUtilities.GetPCPerItem(objMain.objUtilities.GetPCPerItem(oDT.GetValue("ItemCode", 0)))
                            BaseUom = objMain.objUtilities.GetItemBaseUom(oDT.GetValue("ItemCode", 0))
                            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, pVal.Row)
                            Dim GetStockDtls As String = " Select (Select ""UomCode"" From OUOM Where ""UomEntry"" = A.""IUoMEntry"") As ""UomCode"" , A.""IUoMEntry""  " &
                                                         " From OITM A  Where ""ItemCode""  = '" & oDT.GetValue("ItemCode", 0) & "'  "
                            Dim oRsGetStockDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetStockDtls.DoQuery(GetStockDtls)
                            Dim GetBarCode As String = " SELECT Top 1 ""BcdCode"" FROM OBCD WHERE ""UomEntry"" = '" & oRsGetStockDtls.Fields.Item("IUoMEntry").Value & "' " &
                                                       " And ""ItemCode"" = '" & oDT.GetValue("ItemCode", 0) & "' ORDER BY ""BcdEntry"" ASC "
                            Dim oRsGetBarCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetBarCode.DoQuery(GetBarCode)
                            If oRsGetBarCode.RecordCount > 0 Then
                                oDBs_Details1.SetValue("U_MEBRCD", oDBs_Details1.Offset, oRsGetBarCode.Fields.Item("BcdCode").Value)
                                oDBs_Details1.SetValue("U_ADBRCD", oDBs_Details1.Offset, oRsGetBarCode.Fields.Item("BcdCode").Value)
                            Else
                                oDBs_Details1.SetValue("U_MEBRCD", oDBs_Details1.Offset, "")
                                oDBs_Details1.SetValue("U_MEADBRCD", oDBs_Details1.Offset, "")
                            End If
                            oDBs_Details1.SetValue("U_MEITMCD", oDBs_Details1.Offset, oDT.GetValue("ItemCode", 0))
                            oDBs_Details1.SetValue("U_MEITDSC", oDBs_Details1.Offset, oDT.GetValue("ItemName", 0))
                            oDBs_Details1.SetValue("U_MEUOM", oDBs_Details1.Offset, oRsGetStockDtls.Fields.Item("UomCode").Value)
                            If PCPrPack <> "" Then
                                oDBs_Details1.SetValue("U_MEPKGE", oDBs_Details1.Offset, PCPrPack & "*" & "1*" & BaseUom)
                            Else
                                oDBs_Details1.SetValue("U_MEPKGE", oDBs_Details1.Offset, "1*" & BaseUom)
                            End If

                            Dim GetItSrBtchDet As String = " SELECT " & IsNull & "(T0.""ManBtchNum"",'N')," & IsNull & "(T0.""ManSerNum"",'N')," & IsNull & "(T0.""MngMethod"",'N') " &
                                                           " FROM OITM T0 WHERE T0.""ItemCode""='" & oDT.GetValue("ItemCode", 0) & "'"
                            Dim oRsGetItSrBtchDet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetItSrBtchDet.DoQuery(GetItSrBtchDet)
                            If oRsGetItSrBtchDet.RecordCount > 0 Then
                                If CStr(oRsGetItSrBtchDet.Fields.Item(0).Value).Trim = "Y" Then
                                    oDBs_Details1.SetValue("U_METYPE", oDBs_Details1.Offset, "BatchNo")
                                ElseIf CStr(oRsGetItSrBtchDet.Fields.Item(1).Value).Trim = "Y" Then
                                    oDBs_Details1.SetValue("U_METYPE", oDBs_Details1.Offset, "SerialNo")
                                End If
                                oDBs_Details1.SetValue("U_MESRBT", oDBs_Details1.Offset, "")
                                oDBs_Details1.SetValue("U_CHKBIN", oDBs_Details1.Offset, "N")
                                oDBs_Details1.SetValue("U_METWHS", oDBs_Details1.Offset, objForm.Items.Item("19").Specific.Value)
                                oDBs_Details1.SetValue("U_METWBL", oDBs_Details1.Offset, "")
                                oDBs_Details1.SetValue("U_METBRCH", oDBs_Details1.Offset, objForm.Items.Item("40").Specific.Selected.Description)
                                oDBs_Details1.SetValue("U_MEREQQTY", oDBs_Details1.Offset, "")
                                oDBs_Details1.SetValue("U_MERECQTY", oDBs_Details1.Offset, "")
                                oDBs_Details1.SetValue("U_MEVARQTY", oDBs_Details1.Offset, "")
                                oDBs_Details1.SetValue("U_MESTRNO", oDBs_Details1.Offset, "")
                                oDBs_Details1.SetValue("U_MESTREN", oDBs_Details1.Offset, "")
                                oDBs_Details1.SetValue("U_MEITRNO", oDBs_Details1.Offset, "")
                                oDBs_Details1.SetValue("U_MEITREN", oDBs_Details1.Offset, "")
                                oDBs_Details1.SetValue("U_MEBSTYP", oDBs_Details1.Offset, "")
                                oDBs_Details1.SetValue("U_MEBLN", oDBs_Details1.Offset, "")
                                objMatrix1.SetLineData(pVal.Row)
                                If pVal.Row = objMatrix1.VisualRowCount Then
                                    Me.SetNewLine(objForm.UniqueID)
                                End If
                            End If
                        End If

                        If oCFL.UniqueID = "CFL_BCD" Then
                            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, pVal.Row)
                            Dim GetStockDtls As String = " Select A.""ItemName"" , (Select ""UomCode"" From OUOM Where ""UomEntry"" = A.""IUoMEntry"") As ""UomCode"" , A.""IUoMEntry"" , A.""UgpEntry"" " &
                                                         " From OITM A  Where ""ItemCode""  = '" & oDT.GetValue("ItemCode", 0) & "'  "
                            Dim oRsGetStockDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetStockDtls.DoQuery(GetStockDtls)
                            oDBs_Details1.SetValue("U_MEBRCD", oDBs_Details1.Offset, oDT.GetValue("BcdCode", 0))
                            oDBs_Details1.SetValue("U_ADBRCD", oDBs_Details1.Offset, oDT.GetValue("BcdCode", 0))
                            oDBs_Details1.SetValue("U_MEITMCD", oDBs_Details1.Offset, oDT.GetValue("ItemCode", 0))
                            oDBs_Details1.SetValue("U_MEITDSC", oDBs_Details1.Offset, oRsGetStockDtls.Fields.Item("ItemName").Value)
                            'oDBs_Details1.SetValue("U_MEUOM", oDBs_Details1.Offset, oRsGetStockDtls.Fields.Item("UomCode").Value)
                            Dim PCPrPack As String = objMain.objUtilities.GetPCPerItem(oDT.GetValue("ItemCode", 0))
                            BaseUom = objMain.objUtilities.GetItemBaseUom(oDT.GetValue("ItemCode", 0))

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
                                        oDBs_Details1.SetValue("U_MEPKGE", oDBs_Details1.Offset, "1*" & BaseUom)
                                    Else
                                        oDBs_Details1.SetValue("U_MEPKGE", oDBs_Details1.Offset, PCPrPack & "*" & "1*" & BaseUom)
                                    End If
                                Else
                                    Dim GetPackageCode As String = " Select Cast(Cast(T0.""BaseQty"" As Int) As Varchar)" & Concate & "'*'" & Concate & "'" & BaseUom & "'" &
                                                                   "  From UGP1 T0 Inner Join OUOM T1 On  " &
                                                                   " T1.""UomEntry"" = T0.""UomEntry""  Where ""UgpEntry"" = '" & oRsGetStockDtls.Fields.Item("UgpEntry").Value & "' And  " &
                                                                   " T1.""UomCode"" = '" & oRsGetUOMCode.Fields.Item("UomCode").Value & "' "
                                    Dim oRsGetPackageCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRsGetPackageCode.DoQuery(GetPackageCode)
                                    oDBs_Details1.SetValue("U_MEPKGE", oDBs_Details1.Offset, oRsGetPackageCode.Fields.Item(0).Value)
                                    If PCPrPack <> "" Then
                                        oDBs_Details1.SetValue("U_MEPKGE", oDBs_Details1.Offset, PCPrPack & "*" & oRsGetPackageCode.Fields.Item(0).Value)
                                    End If
                                End If
                            Else
                                Dim ChkInventoryUOM As String = " Select ""UomCode"" From OUOM Where ""UomEntry"" = '" & oRsGetStockDtls.Fields.Item("IUoMEntry").Value & "' And ""UomCode"" = '" & objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value & "'  "
                                Dim oRsChkInventoryUOM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsChkInventoryUOM.DoQuery(ChkInventoryUOM)
                                If oRsChkInventoryUOM.RecordCount > 0 Then
                                    If PCPrPack = "" Then
                                        oDBs_Details1.SetValue("U_MEPKGE", oDBs_Details1.Offset, "1*" & BaseUom)
                                    Else
                                        oDBs_Details1.SetValue("U_MEPKGE", oDBs_Details1.Offset, PCPrPack & "*" & "1*" & BaseUom)
                                    End If
                                Else
                                    Dim GetPackageCode As String = " Select Cast(Cast(T0.""BaseQty"" As Int) As Varchar)" & Concate & "'*'" & Concate & "'" & BaseUom & "'" &
                                                                  " From UGP1 T0 Inner Join OUOM T1 On  " &
                                                                  " T1.""UomEntry"" = T0.""UomEntry""  Where ""UgpEntry"" = '" & oRsGetStockDtls.Fields.Item("UgpEntry").Value & "' And  " &
                                                                  " T1.""UomCode"" = '" & objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value & "' "
                                    Dim oRsGetPackageCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRsGetPackageCode.DoQuery(GetPackageCode)
                                    oDBs_Details1.SetValue("U_MEPKGE", oDBs_Details1.Offset, oRsGetPackageCode.Fields.Item(0).Value)

                                    If PCPrPack <> "" Then
                                        oDBs_Details1.SetValue("U_MEPKGE", oDBs_Details1.Offset, PCPrPack & "*" & oRsGetPackageCode.Fields.Item(0).Value)
                                    End If


                                End If
                                oDBs_Details1.SetValue("U_MEUOM", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                            End If
                            Dim GetItSrBtchDet As String = " SELECT " & IsNull & "(T0.""ManBtchNum"",'N')," & IsNull & "(T0.""ManSerNum"",'N')," & IsNull & "(T0.""MngMethod"",'N') " &
                                                           " FROM OITM T0 WHERE T0.""ItemCode""='" & oDT.GetValue("ItemCode", 0) & "'"
                            Dim oRsGetItSrBtchDet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetItSrBtchDet.DoQuery(GetItSrBtchDet)
                            If oRsGetItSrBtchDet.RecordCount > 0 Then
                                If CStr(oRsGetItSrBtchDet.Fields.Item(0).Value).Trim = "Y" Then
                                    oDBs_Details1.SetValue("U_METYPE", oDBs_Details1.Offset, "BatchNo")
                                ElseIf CStr(oRsGetItSrBtchDet.Fields.Item(1).Value).Trim = "Y" Then
                                    oDBs_Details1.SetValue("U_METYPE", oDBs_Details1.Offset, "SerialNo")
                                End If
                                oDBs_Details1.SetValue("U_MESRBT", oDBs_Details1.Offset, "")
                                oDBs_Details1.SetValue("U_CHKBIN", oDBs_Details1.Offset, "N")
                                oDBs_Details1.SetValue("U_METWHS", oDBs_Details1.Offset, objForm.Items.Item("19").Specific.Value)
                                oDBs_Details1.SetValue("U_METWBL", oDBs_Details1.Offset, "")
                                oDBs_Details1.SetValue("U_METBRCH", oDBs_Details1.Offset, objForm.Items.Item("40").Specific.Selected.Description)
                                oDBs_Details1.SetValue("U_MEREQQTY", oDBs_Details1.Offset, "")
                                oDBs_Details1.SetValue("U_MERECQTY", oDBs_Details1.Offset, "")
                                oDBs_Details1.SetValue("U_MEVARQTY", oDBs_Details1.Offset, "")
                                oDBs_Details1.SetValue("U_MESTRNO", oDBs_Details1.Offset, "")
                                oDBs_Details1.SetValue("U_MESTREN", oDBs_Details1.Offset, "")
                                oDBs_Details1.SetValue("U_MEITRNO", oDBs_Details1.Offset, "")
                                oDBs_Details1.SetValue("U_MEITREN", oDBs_Details1.Offset, "")
                                oDBs_Details1.SetValue("U_MEBSTYP", oDBs_Details1.Offset, "")
                                oDBs_Details1.SetValue("U_MEBLN", oDBs_Details1.Offset, "")
                                objMatrix1.SetLineData(pVal.Row)
                                If pVal.Row = objMatrix1.VisualRowCount Then
                                    Me.SetNewLine(objForm.UniqueID)
                                End If
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
                                    oDBs_Details1.SetValue("U_ADBRCD", oDBs_Details1.Offset, oRsGetBarCode.Fields.Item("BcdCode").Value)
                                End If
                            Else
                                oDBs_Details1.SetValue("U_MEBRCD", oDBs_Details1.Offset, oRsChkBarCodeUOM.Fields.Item("BcdCode").Value)
                                oDBs_Details1.SetValue("U_ADBRCD", oDBs_Details1.Offset, oRsChkBarCodeUOM.Fields.Item("BcdCode").Value)
                            End If
                            oDBs_Details1.SetValue("U_MEITMCD", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_MEITDSC", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_MEUOM", oDBs_Details1.Offset, oDT.GetValue("UomCode", 0))
                            Dim ChkInventoryUOM As String = " Select ""IUoMEntry"" From OITM  T0 Inner Join OUOM T1 On T1.""UomEntry"" = ""IUoMEntry"" " &
                                                            " Where ""ItemCode"" = '" & objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value & "' And T1.""UomCode"" = '" & oDT.GetValue("UomCode", 0) & "'"
                            Dim oRsChkInventoryUOM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsChkInventoryUOM.DoQuery(ChkInventoryUOM)
                            Dim PCPrPack As String = objMain.objUtilities.GetPCPerItem(objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                            BaseUom = objMain.objUtilities.GetItemBaseUom(objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)

                            If oRsChkInventoryUOM.RecordCount > 0 Then
                                If PCPrPack = "" Then
                                    oDBs_Details1.SetValue("U_MEPKGE", oDBs_Details1.Offset, "1*" & BaseUom)
                                Else
                                    oDBs_Details1.SetValue("U_MEPKGE", oDBs_Details1.Offset, PCPrPack & "*" & "1*" & BaseUom)
                                End If
                            Else
                                Dim GetPackageCode As String = " Select Cast(Cast(T0.""BaseQty"" As Int) As Varchar)" & Concate & "'*'" & Concate & "'" & BaseUom & "'" &
                                                               " From UGP1 T0 Inner Join OUOM T1 On  " &
                                                               " T1.""UomEntry"" = T0.""UomEntry""  Where ""UgpEntry"" = (Select ""UgpEntry"" From OITM Where ""ItemCode"" = '" & objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value & "') And  " &
                                                               " T1.""UomCode"" = '" & oDT.GetValue("UomCode", 0) & "' "
                                Dim oRsGetPackageCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetPackageCode.DoQuery(GetPackageCode)
                                oDBs_Details1.SetValue("U_MEPKGE", oDBs_Details1.Offset, oRsGetPackageCode.Fields.Item(0).Value)
                                If PCPrPack <> "" Then
                                    oDBs_Details1.SetValue("U_MEPKGE", oDBs_Details1.Offset, PCPrPack & "*" & oRsGetPackageCode.Fields.Item(0).Value)
                                End If
                            End If

                            oDBs_Details1.SetValue("U_METYPE", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_MESRBT", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_CHKBIN", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Checked)
                            oDBs_Details1.SetValue("U_METWHS", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_9").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_METWBL", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_10").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_METBRCH", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_11").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_MEREQQTY", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_12").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_MERECQTY", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_13").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_MEVARQTY", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_14").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_MESTRNO", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_15").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_MESTREN", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_16").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_MEITRNO", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_17").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_MEITREN", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_18").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_MEBSTYP", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_19").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_MEBLN", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_20").Cells.Item(pVal.Row).Specific.Value)
                            objMatrix1.SetLineData(pVal.Row)
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_MI")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_MI_C0")
                    objMatrix1 = objForm.Items.Item("3").Specific

                    If pVal.ItemUID = "15" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("15").Specific.Selected.Value <> "" Then
                            Dim series As Integer = objForm.Items.Item("15").Specific.Selected.Value
                            SetDefault(objForm.UniqueID, "CobmoSelect", series)
                        End If
                    End If
                    If pVal.ItemUID = "40" And pVal.BeforeAction = False Then
                        objComboBox = objForm.Items.Item("40").Specific
                        Me.CFLWhsCodeFilter(objForm.UniqueID, "CFL_TWHS", objComboBox.Value)
                    End If

                    If pVal.ItemUID = "42" And pVal.BeforeAction = False Then
                        objComboBox = objForm.Items.Item("42").Specific
                        Me.CFLWhsCodeFilter(objForm.UniqueID, "CFL_FWH", objComboBox.Value)
                    End If

                    If pVal.ItemUID = "52" And pVal.BeforeAction = False Then
                        objComboBox = objForm.Items.Item("52").Specific
                        If objComboBox.Value = "DP" Then
                            objForm.Items.Item("5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                            'objForm.Items.Item("40").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            'objForm.Items.Item("40").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                            'objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            'objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            'objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                            objForm.Items.Item("42").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            objForm.Items.Item("42").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                            objForm.Items.Item("33").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            objForm.Items.Item("33").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                            objForm.Items.Item("16").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            objForm.Items.Item("47").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            objForm.Items.Item("17").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            objForm.Items.Item("35").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        End If
                        If objComboBox.Value = "ST" Then
                            objForm.Items.Item("5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            objForm.Items.Item("5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                            ' objForm.Items.Item("40").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                            objForm.Items.Item("42").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                            objForm.Items.Item("33").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                            objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                            objForm.Items.Item("16").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            objForm.Items.Item("47").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            objForm.Items.Item("17").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            objForm.Items.Item("35").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_MI")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_MI_C0")
                    objMatrix1 = objForm.Items.Item("3").Specific
                    If pVal.ItemUID = "3" And pVal.ColUID = "V_7" And pVal.BeforeAction = False Then
                        If objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            Dim ActualQty As Double = CDbl(objMatrix1.Columns.Item("V_12").Cells.Item(pVal.Row).Specific.Value)
                            objMain.objSerialBatchDetails.CreateForm(objForm.UniqueID,
                                                                   objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value,
                                                                   ActualQty,
                                                                   objForm.Items.Item("17").Specific.Value,
                                                                   pVal.Row, objMatrix1.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value,
                                                                   objMatrix1.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Checked, objForm.Items.Item("46").Specific.Value)
                        Else
                            objMain.objApplication.StatusBar.SetText("Item type should be BatchNo or SerialNo managed ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                    End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub PressedOnLinked(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMain.objApplication.ActivateMenuItem("ME_STR")
            objSTRForm = objMain.objApplication.Forms.GetForm("ME_STR_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objSTRForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            objSTRForm.Freeze(True)
            objSTRForm.Items.Item("26").Specific.Value = objForm.Items.Item("45").Specific.Value
            objSTRForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            objSTRForm.Freeze(False)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_MI")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_MI_C0")
            objMatrix1 = objForm.Items.Item("3").Specific
            'If objForm.Items.Item("5").Specific.Value = "" Then
            '    objMain.objApplication.StatusBar.SetText("MaterialRequestNo Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False

            If objForm.Items.Item("52").Specific.Value.Trim = "" Then
                objMain.objApplication.StatusBar.SetText("Material Type Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If objForm.Items.Item("23").Specific.Value.Trim = "" Then
                objMain.objApplication.StatusBar.SetText("DocDate Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If objForm.Items.Item("53").Specific.Value.Trim = "" Then
                objMain.objApplication.StatusBar.SetText("Material Status Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If objForm.Items.Item("33").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("From Warehouse  Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("19").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("To Warehouse  Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If


            If objMatrix1.VisualRowCount = 0 Then
                objMain.objApplication.StatusBar.SetText("Matrix cannot be Empty")
                Return False
            ElseIf objMatrix1.VisualRowCount = 1 Then
                If (objMatrix1.Columns.Item("V_1").Cells.Item(1).Specific.Value.Trim = "") Then
                    objMain.objApplication.StatusBar.SetText("Matrix cannot be Empty")
                    Return False
                End If
            End If

            For i As Integer = 1 To objMatrix1.VisualRowCount
                Dim ItemCode As String = objMatrix1.Columns.Item("V_1").Cells.Item(i).Specific.Value.Trim
                Dim ReceivedQty As Double = objMatrix1.Columns.Item("V_13").Cells.Item(i).Specific.Value.Trim
                If (ItemCode <> "" And ReceivedQty = 0) Then
                    objMain.objApplication.StatusBar.SetText("For Item " & ItemCode & " and Line No. " & i & ".... ReceivedQty should not be zero!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Next


            For i As Integer = 1 To objMatrix1.VisualRowCount
                Dim ItemCode As String = objMatrix1.Columns.Item("V_1").Cells.Item(i).Specific.Value.Trim
                Dim DispatchedQty As Double = objMatrix1.Columns.Item("V_12").Cells.Item(i).Specific.Value.Trim
                Dim ReceivedQty As Double = objMatrix1.Columns.Item("V_13").Cells.Item(i).Specific.Value.Trim
                If (ItemCode <> "" And DispatchedQty > 0 And ReceivedQty > 0) Then
                    If ReceivedQty > DispatchedQty Then
                        objMain.objApplication.StatusBar.SetText("For Item " & ItemCode & " and Line No. " & i & ".... ReceivedQty should not be greater than Dispatched Qty !!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            Next

            For i As Integer = 1 To objMatrix1.VisualRowCount
                Dim ItemCode As String = objMatrix1.Columns.Item("V_1").Cells.Item(i).Specific.Value.Trim
                Dim VarianceQty As Double = objMatrix1.Columns.Item("V_14").Cells.Item(i).Specific.Value.Trim
                If (ItemCode <> "") Then
                    If VarianceQty > 0 Then
                        objMain.objApplication.StatusBar.SetText("For Item " & ItemCode & " and Line No. " & i & ".... VarianceQty should not be greater than Zero !!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            Next

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Public Sub LoadMatrixGI(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_MI")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_MI_C0")
            objMatrix1 = objForm.Items.Item("3").Specific
            Dim GetMateriaRequestDtls As String = " Select ""LineNum"" , ""CodeBars"" , ""ItemCode"" , ""Dscription"" , ""UomCode"" , ""U_Addbarcode"" , ""U_Packing"" , ""Quantity""  As ""DispatchQty"" ,T1.""U_INVTRREQ"" ,T1.""U_INVTRREN"" ,""U_INVTRRLN"" " &
                                                  " From OIGE T0 Inner Join IGE1 T1 On T1.""DocEntry"" = T0.""DocEntry"" Where T0.""DocEntry"" = '" & objForm.Items.Item("46").Specific.Value & "' and T1.""U_METWHSCD""='" & objForm.Items.Item("19").Specific.Value & "'"
            Dim oRsGetMateriaRequestDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetMateriaRequestDtls.DoQuery(GetMateriaRequestDtls)
            If oRsGetMateriaRequestDtls.RecordCount > 0 Then
                objMatrix1.Clear()
                oDBs_Details1.Clear()
                objMatrix1.FlushToDataSource()
                For i As Integer = 1 To oRsGetMateriaRequestDtls.RecordCount
                    objMatrix1.AddRow()
                    oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, i)
                    oDBs_Details1.SetValue("U_MEBRCD", oDBs_Details1.Offset, oRsGetMateriaRequestDtls.Fields.Item("U_Addbarcode").Value)
                    oDBs_Details1.SetValue("U_MEITMCD", oDBs_Details1.Offset, oRsGetMateriaRequestDtls.Fields.Item("ItemCode").Value)
                    oDBs_Details1.SetValue("U_MEITDSC", oDBs_Details1.Offset, oRsGetMateriaRequestDtls.Fields.Item("Dscription").Value)
                    oDBs_Details1.SetValue("U_MEUOM", oDBs_Details1.Offset, oRsGetMateriaRequestDtls.Fields.Item("UomCode").Value)
                    oDBs_Details1.SetValue("U_ADBRCD", oDBs_Details1.Offset, oRsGetMateriaRequestDtls.Fields.Item("U_Addbarcode").Value)
                    oDBs_Details1.SetValue("U_MEPKGE", oDBs_Details1.Offset, oRsGetMateriaRequestDtls.Fields.Item("U_Packing").Value)
                    Dim GetItSrBtchDet As String = " SELECT " & IsNull & "(T0.""ManBtchNum"",'N')," & IsNull & "(T0.""ManSerNum"",'N')," & IsNull & "(T0.""MngMethod"",'N') " &
                                                   " FROM OITM T0 WHERE T0.""ItemCode""='" & oRsGetMateriaRequestDtls.Fields.Item("ItemCode").Value & "'"
                    Dim oRsGetItSrBtchDet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetItSrBtchDet.DoQuery(GetItSrBtchDet)
                    If oRsGetItSrBtchDet.RecordCount > 0 Then
                        If CStr(oRsGetItSrBtchDet.Fields.Item(0).Value).Trim = "Y" Then
                            oDBs_Details1.SetValue("U_METYPE", oDBs_Details1.Offset, "BatchNo")
                            Dim GetBatchDetails As String = " Select ""BatchNum"" From IBT1 Where ""BaseType"" = '60' And  ""BaseEntry"" = '" & objForm.Items.Item("46").Specific.Value & "' " &
                                                            " And  ""ItemCode""  = '" & oRsGetMateriaRequestDtls.Fields.Item("ItemCode").Value & "' And ""WhsCode"" = '" & objForm.Items.Item("17").Specific.Value & "' "
                            Dim oRsGetBatchDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetBatchDetails.DoQuery(GetBatchDetails)
                            If oRsGetBatchDetails.RecordCount > 0 Then
                                oDBs_Details1.SetValue("U_MESRBT", oDBs_Details1.Offset, oRsGetBatchDetails.Fields.Item("BatchNum").Value)
                            End If
                        ElseIf CStr(oRsGetItSrBtchDet.Fields.Item(1).Value).Trim = "Y" Then
                            oDBs_Details1.SetValue("U_METYPE", oDBs_Details1.Offset, "SerialNo")
                            'Dim GetSerialNoDetails As String = " SELECT Distinct IFNULL(T4.""SuppSerial"" ,  T4.""IntrSerial"") As ""IntrSerial"" FROM OWTR T0 INNER JOIN WTR1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" INNER JOIN OITW T2 " & _
                            '                                   " ON T1.""ItemCode"" = T2.""ItemCode"" AND T1.""WhsCode"" = T2.""WhsCode"" LEFT JOIN SRI1 T3 " & _
                            '                                   " ON T3.""BaseEntry""=T1.""DocEntry"" AND T3.""BaseLinNum"" = T1.""LineNum"" AND T3.""BaseType"" ='67' " & _
                            '                                   " LEFT JOIN OSRI T4 ON T3.""ItemCode""=T4.""ItemCode"" AND T3.""SysSerial"" =T4.""SysSerial"" WHERE " & _
                            '                                   " T1.""DocEntry"" = '" & objForm.Items.Item("46").Specific.Value & "' And  " & _
                            '                                   " T1.""ItemCode"" ='" & oRsGetMateriaRequestDtls.Fields.Item("ItemCode").Value & "' And T1.""WhsCode"" = '" & objForm.Items.Item("17").Specific.Value & "' "
                            'Dim oRsGetSerialNoDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'oRsGetSerialNoDetails.DoQuery(GetSerialNoDetails)
                            'Dim SerialNoArray As String = ""
                            'If oRsGetSerialNoDetails.Fields.Item("IntrSerial").Value <> "" Then
                            '    For j As Integer = 1 To oRsGetSerialNoDetails.RecordCount
                            '        SerialNoArray = SerialNoArray & oRsGetSerialNoDetails.Fields.Item("IntrSerial").Value & ";"
                            '        oRsGetSerialNoDetails.MoveNext()
                            '    Next
                            '    oDBs_Details1.SetValue("U_MESRBT", oDBs_Details1.Offset, SerialNoArray)
                            'End If
                        Else
                            oDBs_Details1.SetValue("U_METYPE", oDBs_Details1.Offset, "")
                            oDBs_Details1.SetValue("U_MESRBT", oDBs_Details1.Offset, "")
                        End If
                    Else
                        oDBs_Details1.SetValue("U_METYPE", oDBs_Details1.Offset, "")
                        oDBs_Details1.SetValue("U_MESRBT", oDBs_Details1.Offset, "")
                    End If
                    Dim CheckBin As String = "Select ""BinActivat"" , (Select ""BinCode"" From OBIN Where ""AbsEntry"" =  ""DftBinAbs"") ""DflBin"" From OWHS Where ""WhsCode"" ='" & objForm.Items.Item("19").Specific.Value & "'  And ""BinActivat"" = 'Y'"
                    Dim oRsCheckBin As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsCheckBin.DoQuery(CheckBin)
                    If oRsCheckBin.RecordCount > 0 Then
                        oDBs_Details1.SetValue("U_CHKBIN", oDBs_Details1.Offset, oRsCheckBin.Fields.Item(0).Value)
                        oDBs_Details1.SetValue("U_METWBL", oDBs_Details1.Offset, objMain.objUtilities.GetWhsBinLoc(objForm.Items.Item("19").Specific.Value))
                    Else
                        oDBs_Details1.SetValue("U_CHKBIN", oDBs_Details1.Offset, "")
                        oDBs_Details1.SetValue("U_METWBL", oDBs_Details1.Offset, "")
                    End If
                    oDBs_Details1.SetValue("U_METWHS", oDBs_Details1.Offset, objForm.Items.Item("19").Specific.Value)
                    oDBs_Details1.SetValue("U_METBRCH", oDBs_Details1.Offset, objForm.Items.Item("40").Specific.Value)
                    oDBs_Details1.SetValue("U_MEREQQTY", oDBs_Details1.Offset, oRsGetMateriaRequestDtls.Fields.Item("DispatchQty").Value)
                    oDBs_Details1.SetValue("U_MERECQTY", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_MEVARQTY", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_MESTRNO", oDBs_Details1.Offset, oRsGetMateriaRequestDtls.Fields.Item("U_INVTRREQ").Value)
                    oDBs_Details1.SetValue("U_MESTREN", oDBs_Details1.Offset, oRsGetMateriaRequestDtls.Fields.Item("U_INVTRREN").Value)
                    oDBs_Details1.SetValue("U_MEITRNO", oDBs_Details1.Offset, objForm.Items.Item("5").Specific.Value)
                    oDBs_Details1.SetValue("U_MEITREN", oDBs_Details1.Offset, objForm.Items.Item("46").Specific.Value)
                    oDBs_Details1.SetValue("U_MEBSTYP", oDBs_Details1.Offset, "67")
                    oDBs_Details1.SetValue("U_MEBLN", oDBs_Details1.Offset, oRsGetMateriaRequestDtls.Fields.Item("LineNum").Value)
                    oDBs_Details1.SetValue("U_MESTRLN", oDBs_Details1.Offset, oRsGetMateriaRequestDtls.Fields.Item("U_INVTRRLN").Value)
                    objMatrix1.SetLineData(i)
                    oRsGetMateriaRequestDtls.MoveNext()
                Next
                objMatrix1.AutoResizeColumns()
                objMain.objApplication.StatusBar.SetText("Loaded successfully.....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub LoadMatrix(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_MI")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_MI_C0")
            objMatrix1 = objForm.Items.Item("3").Specific
            Dim GetMateriaRequestDtls As String = " Select ""LineNum"" , ""CodeBars"" , ""ItemCode"" , ""Dscription"" , ""UomCode"" , ""U_Addbarcode"" , ""U_Packing"" , ""Quantity""  As ""DispatchQty"" ,T1.""U_INVTRREQ"" ,T1.""U_INVTRREN"" ,""U_INVTRRLN"" " &
                                                  " From OWTR T0 Inner Join WTR1 T1 On T1.""DocEntry"" = T0.""DocEntry"" Where T0.""DocEntry"" = '" & objForm.Items.Item("46").Specific.Value & "'"
            Dim oRsGetMateriaRequestDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetMateriaRequestDtls.DoQuery(GetMateriaRequestDtls)
            If oRsGetMateriaRequestDtls.RecordCount > 0 Then
                objMatrix1.Clear()
                oDBs_Details1.Clear()
                objMatrix1.FlushToDataSource()
                For i As Integer = 1 To oRsGetMateriaRequestDtls.RecordCount
                    objMatrix1.AddRow()
                    oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, i)
                    oDBs_Details1.SetValue("U_MEBRCD", oDBs_Details1.Offset, oRsGetMateriaRequestDtls.Fields.Item("U_Addbarcode").Value)
                    oDBs_Details1.SetValue("U_MEITMCD", oDBs_Details1.Offset, oRsGetMateriaRequestDtls.Fields.Item("ItemCode").Value)
                    oDBs_Details1.SetValue("U_MEITDSC", oDBs_Details1.Offset, oRsGetMateriaRequestDtls.Fields.Item("Dscription").Value)
                    oDBs_Details1.SetValue("U_MEUOM", oDBs_Details1.Offset, oRsGetMateriaRequestDtls.Fields.Item("UomCode").Value)
                    oDBs_Details1.SetValue("U_ADBRCD", oDBs_Details1.Offset, oRsGetMateriaRequestDtls.Fields.Item("U_Addbarcode").Value)
                    oDBs_Details1.SetValue("U_MEPKGE", oDBs_Details1.Offset, oRsGetMateriaRequestDtls.Fields.Item("U_Packing").Value)
                    Dim GetItSrBtchDet As String = " SELECT " & IsNull & "(T0.""ManBtchNum"",'N')," & IsNull & "(T0.""ManSerNum"",'N')," & IsNull & "(T0.""MngMethod"",'N') " &
                                                   " FROM OITM T0 WHERE T0.""ItemCode""='" & oRsGetMateriaRequestDtls.Fields.Item("ItemCode").Value & "'"
                    Dim oRsGetItSrBtchDet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetItSrBtchDet.DoQuery(GetItSrBtchDet)
                    If oRsGetItSrBtchDet.RecordCount > 0 Then
                        If CStr(oRsGetItSrBtchDet.Fields.Item(0).Value).Trim = "Y" Then
                            oDBs_Details1.SetValue("U_METYPE", oDBs_Details1.Offset, "BatchNo")
                            Dim GetBatchDetails As String = " Select ""BatchNum"" From IBT1 Where ""BaseType"" = '67' And  ""BaseEntry"" = '" & objForm.Items.Item("46").Specific.Value & "' " &
                                                            " And  ""ItemCode""  = '" & oRsGetMateriaRequestDtls.Fields.Item("ItemCode").Value & "' And ""WhsCode"" = '" & objForm.Items.Item("17").Specific.Value & "' "
                            Dim oRsGetBatchDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetBatchDetails.DoQuery(GetBatchDetails)
                            If oRsGetBatchDetails.RecordCount > 0 Then
                                oDBs_Details1.SetValue("U_MESRBT", oDBs_Details1.Offset, oRsGetBatchDetails.Fields.Item("BatchNum").Value)
                            End If
                        ElseIf CStr(oRsGetItSrBtchDet.Fields.Item(1).Value).Trim = "Y" Then
                            oDBs_Details1.SetValue("U_METYPE", oDBs_Details1.Offset, "SerialNo")
                            'Dim GetSerialNoDetails As String = " SELECT Distinct IFNULL(T4.""SuppSerial"" ,  T4.""IntrSerial"") As ""IntrSerial"" FROM OWTR T0 INNER JOIN WTR1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" INNER JOIN OITW T2 " & _
                            '                                   " ON T1.""ItemCode"" = T2.""ItemCode"" AND T1.""WhsCode"" = T2.""WhsCode"" LEFT JOIN SRI1 T3 " & _
                            '                                   " ON T3.""BaseEntry""=T1.""DocEntry"" AND T3.""BaseLinNum"" = T1.""LineNum"" AND T3.""BaseType"" ='67' " & _
                            '                                   " LEFT JOIN OSRI T4 ON T3.""ItemCode""=T4.""ItemCode"" AND T3.""SysSerial"" =T4.""SysSerial"" WHERE " & _
                            '                                   " T1.""DocEntry"" = '" & objForm.Items.Item("46").Specific.Value & "' And  " & _
                            '                                   " T1.""ItemCode"" ='" & oRsGetMateriaRequestDtls.Fields.Item("ItemCode").Value & "' And T1.""WhsCode"" = '" & objForm.Items.Item("17").Specific.Value & "' "
                            'Dim oRsGetSerialNoDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'oRsGetSerialNoDetails.DoQuery(GetSerialNoDetails)
                            'Dim SerialNoArray As String = ""
                            'If oRsGetSerialNoDetails.Fields.Item("IntrSerial").Value <> "" Then
                            '    For j As Integer = 1 To oRsGetSerialNoDetails.RecordCount
                            '        SerialNoArray = SerialNoArray & oRsGetSerialNoDetails.Fields.Item("IntrSerial").Value & ";"
                            '        oRsGetSerialNoDetails.MoveNext()
                            '    Next
                            '    oDBs_Details1.SetValue("U_MESRBT", oDBs_Details1.Offset, SerialNoArray)
                            'End If
                        Else
                            oDBs_Details1.SetValue("U_METYPE", oDBs_Details1.Offset, "")
                            oDBs_Details1.SetValue("U_MESRBT", oDBs_Details1.Offset, "")
                        End If
                    Else
                        oDBs_Details1.SetValue("U_METYPE", oDBs_Details1.Offset, "")
                        oDBs_Details1.SetValue("U_MESRBT", oDBs_Details1.Offset, "")
                    End If
                    Dim CheckBin As String = "Select ""BinActivat"" , (Select ""BinCode"" From OBIN Where ""AbsEntry"" =  ""DftBinAbs"") ""DflBin"" From OWHS Where ""WhsCode"" ='" & objForm.Items.Item("19").Specific.Value & "'  And ""BinActivat"" = 'Y'"
                    Dim oRsCheckBin As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsCheckBin.DoQuery(CheckBin)
                    If oRsCheckBin.RecordCount > 0 Then
                        oDBs_Details1.SetValue("U_CHKBIN", oDBs_Details1.Offset, oRsCheckBin.Fields.Item(0).Value)
                        oDBs_Details1.SetValue("U_METWBL", oDBs_Details1.Offset, objMain.objUtilities.GetWhsBinLoc(objForm.Items.Item("19").Specific.Value))
                    Else
                        oDBs_Details1.SetValue("U_CHKBIN", oDBs_Details1.Offset, "")
                        oDBs_Details1.SetValue("U_METWBL", oDBs_Details1.Offset, "")
                    End If
                    oDBs_Details1.SetValue("U_METWHS", oDBs_Details1.Offset, objForm.Items.Item("19").Specific.Value)
                    oDBs_Details1.SetValue("U_METBRCH", oDBs_Details1.Offset, objForm.Items.Item("40").Specific.Value)
                    oDBs_Details1.SetValue("U_MEREQQTY", oDBs_Details1.Offset, oRsGetMateriaRequestDtls.Fields.Item("DispatchQty").Value)
                    oDBs_Details1.SetValue("U_MERECQTY", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_MEVARQTY", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_MESTRNO", oDBs_Details1.Offset, oRsGetMateriaRequestDtls.Fields.Item("U_INVTRREQ").Value)
                    oDBs_Details1.SetValue("U_MESTREN", oDBs_Details1.Offset, oRsGetMateriaRequestDtls.Fields.Item("U_INVTRREN").Value)
                    oDBs_Details1.SetValue("U_MEITRNO", oDBs_Details1.Offset, objForm.Items.Item("5").Specific.Value)
                    oDBs_Details1.SetValue("U_MEITREN", oDBs_Details1.Offset, objForm.Items.Item("46").Specific.Value)
                    oDBs_Details1.SetValue("U_MEBSTYP", oDBs_Details1.Offset, "67")
                    oDBs_Details1.SetValue("U_MEBLN", oDBs_Details1.Offset, oRsGetMateriaRequestDtls.Fields.Item("LineNum").Value)
                    oDBs_Details1.SetValue("U_MESTRLN", oDBs_Details1.Offset, oRsGetMateriaRequestDtls.Fields.Item("U_INVTRRLN").Value)
                    objMatrix1.SetLineData(i)
                    oRsGetMateriaRequestDtls.MoveNext()
                Next
                objMatrix1.AutoResizeColumns()
                objMain.objApplication.StatusBar.SetText("Loaded successfully.....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    'Issue and receipt based on whs act config
#Region "Issue and receipt based on whs act config"

    Public Sub GoodsIssue(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_MI")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_MI_C0")
            objMatrix1 = objForm.Items.Item("3").Specific

            Dim GetMaterialInwardDtls As String = "Select (Select ""IUoMEntry"" From OITM Where ""ItemCode"" = ""U_MEITMCD"") As ""DfltUOM"" , *  From ""@ME_MI_C0"" Where ""DocEntry"" = '" & oDBs_Head.GetValue("DocEntry", 0) & "' And  U_MERECQTY > 0  "
            Dim oRsGetMaterialInwardDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetMaterialInwardDtls.DoQuery(GetMaterialInwardDtls)

            Dim GetAddonConfigurationDtls As String = "Select ""U_MEEXACCT"" From ""OWHS"" Where ""WhsCode"" = '" & objForm.Items.Item("19").Specific.Value & "' And IFNULL(""U_MEEXACCT"",'')<>''  "
            Dim oRsGetAddonConfigurationDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetAddonConfigurationDtls.DoQuery(GetAddonConfigurationDtls)

            If oRsGetAddonConfigurationDtls.RecordCount = 0 Then
                objMain.sCmp = objMain.objCompany.GetCompanyService
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OMI")
                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                objMain.oGeneralData.SetProperty("U_REMARKS", "Please Configure Issue Account in Warehouse Master")
                objMain.oGeneralService.Update(objMain.oGeneralData)
                objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_MI"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_MI_C0"), "DocEntry", oDBs_Details1.GetValue("DocEntry", 0))
                objMatrix1.LoadFromDataSource()
                objForm.Refresh()
                objMain.objApplication.StatusBar.SetText("Please Configure Issue Account in Warehouse Master...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            Dim oInventoryGenExit As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
            Try
                oInventoryGenExit.UserFields.Fields.Item("U_TYPE").Value = "STR"
            Catch ex As Exception
            End Try

            Dim sDate As String = objForm.Items.Item("23").Specific.Value
            sDate = sDate.Insert("4", "-")
            sDate = sDate.Insert("7", "-")
            oInventoryGenExit.DocDate = sDate
            oInventoryGenExit.DocDueDate = sDate
            oInventoryGenExit.TaxDate = sDate
            oInventoryGenExit.UserFields.Fields.Item("U_ADDON").Value = "Yes"

            objComboBox = objForm.Items.Item("42").Specific
            Dim GetSeries As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = '60' And ""Locked"" = 'N'  And ""BPLId"" = '" & objComboBox.Value & "' And ""Indicator"" = Year('" & sDate & "') "
            Dim oRsGetSeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetSeries.DoQuery(GetSeries)
            If oRsGetSeries.RecordCount > 0 Then
                oInventoryGenExit.Series = oRsGetSeries.Fields.Item("Series").Value
            End If
            oInventoryGenExit.BPL_IDAssignedToInvoice = objComboBox.Value
            For i As Integer = 1 To oRsGetMaterialInwardDtls.RecordCount
                If i <> 1 Then
                    oInventoryGenExit.Lines.Add()
                End If
                oInventoryGenExit.Lines.ItemCode = oRsGetMaterialInwardDtls.Fields.Item("U_MEITMCD").Value
                If oRsGetMaterialInwardDtls.Fields.Item("U_MEBRCD").Value <> "" Then
                    oInventoryGenExit.Lines.BarCode = oRsGetMaterialInwardDtls.Fields.Item("U_MEBRCD").Value
                End If
                objComboBox = objForm.Items.Item("52").Specific

                'If objComboBox.Value = "ST" Then
                '    oInventoryGenExit.Lines.WarehouseCode = objForm.Items.Item("33").Specific.Value
                'Else
                oInventoryGenExit.Lines.WarehouseCode = objForm.Items.Item("17").Specific.Value
                'End If

                Dim UomEntry As Integer = Me.GetUomEntry(oRsGetMaterialInwardDtls.Fields.Item("U_MEUOM").Value)
                If UomEntry = CInt(oRsGetMaterialInwardDtls.Fields.Item("DfltUOM").Value) Then
                    oInventoryGenExit.Lines.UseBaseUnits = SAPbobsCOM.BoYesNoEnum.tYES
                    oInventoryGenExit.Lines.UoMEntry = UomEntry
                    If oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "BatchNo" Then
                        oInventoryGenExit.Lines.BatchNumbers.BatchNumber = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
                        oInventoryGenExit.Lines.BatchNumbers.Quantity = CDbl(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value)
                        oInventoryGenExit.Lines.BatchNumbers.Add()
                    ElseIf oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "SerialNo" Then
                        Dim TotalSlrBndBinWiseQty As String = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
                        Dim ArrayTotalSlrBndBinWiseQty() As String = TotalSlrBndBinWiseQty.Split(";")
                        For k As Integer = 0 To CInt(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value) - 1 'ArrayTotalSlrBndBinWiseQty.Length - 1
                            If ArrayTotalSlrBndBinWiseQty(k) <> "" Then
                                oInventoryGenExit.Lines.SerialNumbers.ManufacturerSerialNumber = ArrayTotalSlrBndBinWiseQty(k)
                                oInventoryGenExit.Lines.SerialNumbers.Add()
                            End If
                        Next
                    End If
                Else
                    oInventoryGenExit.Lines.UseBaseUnits = SAPbobsCOM.BoYesNoEnum.tNO
                    oInventoryGenExit.Lines.UoMEntry = UomEntry
                    If oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "BatchNo" Then
                        Dim BaseQty As String = " Select ""BaseQty"" * " & oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value & "   From UGP1 T0 Inner Join OITM T1 On " &
                                                " T1.""UgpEntry"" =  T0.""UgpEntry"" Where ""ItemCode"" = '" &
                                                oRsGetMaterialInwardDtls.Fields.Item("U_MEITMCD").Value & "' And T0.""UomEntry"" = '" & UomEntry & "'"
                        Dim oRsBaseQty As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsBaseQty.DoQuery(BaseQty)
                        oInventoryGenExit.Lines.BatchNumbers.BatchNumber = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
                        oInventoryGenExit.Lines.BatchNumbers.Quantity = oRsBaseQty.Fields.Item(0).Value
                        oInventoryGenExit.Lines.BatchNumbers.Add()
                    ElseIf oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "SerialNo" Then
                        Dim TotalSlrBndBinWiseQty As String = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
                        Dim ArrayTotalSlrBndBinWiseQty() As String = TotalSlrBndBinWiseQty.Split(";")
                        For k As Integer = 0 To CInt(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value) - 1 ' ArrayTotalSlrBndBinWiseQty.Length - 1
                            If ArrayTotalSlrBndBinWiseQty(k) <> "" Then
                                oInventoryGenExit.Lines.SerialNumbers.ManufacturerSerialNumber = ArrayTotalSlrBndBinWiseQty(k)
                                oInventoryGenExit.Lines.SerialNumbers.Add()
                            End If
                        Next
                    End If
                End If
                oInventoryGenExit.Lines.Quantity = CDbl(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value)
                oInventoryGenExit.Lines.AccountCode = oRsGetAddonConfigurationDtls.Fields.Item("U_MEEXACCT").Value
                Try
                    oInventoryGenExit.Lines.UserFields.Fields.Item("U_Addbarcode").Value = oRsGetMaterialInwardDtls.Fields.Item("U_ADBRCD").Value
                    oInventoryGenExit.Lines.UserFields.Fields.Item("U_Packing").Value = oRsGetMaterialInwardDtls.Fields.Item("U_MEPKGE").Value
                Catch ex As Exception
                End Try
                oRsGetMaterialInwardDtls.MoveNext()
            Next

            Dim sIssueGUID As String
            sIssueGUID = System.Guid.NewGuid.ToString()
            Try
                oInventoryGenExit.UserFields.Fields.Item("U_GUID").Value = sIssueGUID
            Catch ex As Exception
            End Try

            If oInventoryGenExit.Add() = 0 Then

                Dim GetOrderEntry As String = "Select ""DocEntry"" , ""DocNum""   From OIGE Where ""U_GUID"" = '" & sIssueGUID & "' "
                Dim oRsGetOrderEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetOrderEntry.DoQuery(GetOrderEntry)

                Dim MIUpdate As String = "Update ""@ME_MI"" Set  ""U_MEGINO"" = '" & oRsGetOrderEntry.Fields.Item(1).Value.ToString & "' , ""U_MEGIEN"" = '" & oRsGetOrderEntry.Fields.Item(0).Value.ToString & "' , ""U_REMARKS"" = ' '   Where ""DocEntry"" = '" & oDBs_Head.GetValue("DocEntry", 0) & "' "
                Dim oRsMIUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsMIUpdate.DoQuery(MIUpdate)

                objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_MI"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_MI_C0"), "DocEntry", oDBs_Details1.GetValue("DocEntry", 0))
                objMatrix1.LoadFromDataSource()
                objForm.Refresh()

                Dim GIUpdate As String = "Update ""OIGE"" Set  ""U_MIDNUM"" = '" & oDBs_Head.GetValue("DocNum", 0).ToString & "' , ""U_MIDENT"" = '" & oDBs_Head.GetValue("DocEntry", 0).ToString & "'  Where ""DocEntry"" = '" & oRsGetOrderEntry.Fields.Item(0).Value.ToString & "' "
                Dim oRsGIUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGIUpdate.DoQuery(GIUpdate)

                objMain.objApplication.StatusBar.SetText("Goods Receipt posting in progress......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                If objForm.Items.Item("11").Specific.Value = "" Then
                    Me.GoodsReceipt(objForm.UniqueID)
                End If
            Else
                objMain.sCmp = objMain.objCompany.GetCompanyService
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OMI")
                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                objMain.oGeneralData.SetProperty("U_REMARKS", "ISSUE : " & objMain.objCompany.GetLastErrorDescription.ToString)
                objMain.oGeneralService.Update(objMain.oGeneralData)
                objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_MI"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_MI_C0"), "DocEntry", oDBs_Details1.GetValue("DocEntry", 0))
                objMatrix1.LoadFromDataSource()
                objForm.Refresh()
                objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oInventoryGenExit)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub GoodsReceipt(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_MI")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_MI_C0")
            objMatrix1 = objForm.Items.Item("3").Specific

            Dim oRsBaseQty As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim GetMaterialInwardDtls As String = "Select  (Select ""IUoMEntry"" From OITM Where ""ItemCode"" = ""U_MEITMCD"") As ""DfltUOM"" , (Select ""AbsEntry"" From OBIN Where ""BinCode"" = ""U_METWBL"") As ""BinAbsEntry"" ,  * From ""@ME_MI_C0"" Where ""DocEntry"" = '" & oDBs_Head.GetValue("DocEntry", 0) & "'  And  U_MERECQTY > 0  "
            Dim oRsGetMaterialInwardDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetMaterialInwardDtls.DoQuery(GetMaterialInwardDtls)

            'Dim GetAddonConfigurationDtls As String = "Select ""U_MEENTACT""  From ""OWHS"" Where ""WhsCode"" = '" & objForm.Items.Item("17").Specific.Value & "' And IFNULL(""U_MEENTACT"",'')<>''  "
            'Dim oRsGetAddonConfigurationDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRsGetAddonConfigurationDtls.DoQuery(GetAddonConfigurationDtls)
            'If oRsGetAddonConfigurationDtls.RecordCount = 0 Then
            '    objMain.sCmp = objMain.objCompany.GetCompanyService
            '    objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OMI")
            '    objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            '    objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0))
            '    objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
            '    objMain.oGeneralData.SetProperty("U_REMARKS", "Please Configure Receipt Account in Warehouse Master")
            '    objMain.oGeneralService.Update(objMain.oGeneralData)
            '    objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_MI"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
            '    objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_MI_C0"), "DocEntry", oDBs_Details1.GetValue("DocEntry", 0))
            '    objMatrix1.LoadFromDataSource()
            '    objForm.Refresh()
            '    objMain.objApplication.StatusBar.SetText("Please Configure Receipt Account in Warehouse Master...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Exit Sub
            'End If
            Dim oInventoryGenEntry As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
            Try
                oInventoryGenEntry.UserFields.Fields.Item("U_TYPE").Value = "STR"
            Catch ex As Exception
            End Try

            Dim sDate As String = objForm.Items.Item("23").Specific.Value
            sDate = sDate.Insert("4", "-")
            sDate = sDate.Insert("7", "-")

            oInventoryGenEntry.DocDate = sDate
            oInventoryGenEntry.DocDueDate = sDate
            oInventoryGenEntry.TaxDate = sDate
            oInventoryGenEntry.UserFields.Fields.Item("U_ADDON").Value = "Yes"

            objComboBox = objForm.Items.Item("40").Specific

            oInventoryGenEntry.BPL_IDAssignedToInvoice = objComboBox.Value

            ' Dim GetSeries As String = " Select ""Series"" , ""SeriesName""   From NNM1 Where ""ObjectCode"" = '59' And ""Locked"" = 'N' And ""BPLId"" =  '" & objComboBox.Value & "' And ""Indicator"" = Year('" & sDate & "') "
            Dim GetSeries As String = " Select ""Series"" , ""SeriesName""   From NNM1 Where ""ObjectCode"" = '59' And ""Locked"" = 'N' "

            Dim oRsGetSeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetSeries.DoQuery(GetSeries)

            oInventoryGenEntry.Series = oRsGetSeries.Fields.Item("Series").Value

            For i As Integer = 1 To oRsGetMaterialInwardDtls.RecordCount
                Dim binlineno As Integer = 0
                If i <> 1 Then
                    oInventoryGenEntry.Lines.Add()
                End If
                'Dim GetItemValue As String = " Select ""StockPrice"" From IGE1 Where ""DocEntry"" = '" & objForm.Items.Item("29").Specific.Value _
                '                             & "' And ""ItemCode""  = '" & oRsGetMaterialInwardDtls.Fields.Item("U_MEITMCD").Value & "'"
                Dim GetItemValue As String = " Select  ABS(SUM(""TransValue"")) As ""TransValue"" FROM OINM Where ""TransType""='60' And " &
                                             " ""CreatedBy""='" & objForm.Items.Item("29").Specific.Value & "' And ""ItemCode"" = '" & oRsGetMaterialInwardDtls.Fields.Item("U_MEITMCD").Value & "' Group By  ""DocLineNum"" "
                Dim oRsGetItemValue As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetItemValue.DoQuery(GetItemValue)
                oInventoryGenEntry.Lines.ItemCode = oRsGetMaterialInwardDtls.Fields.Item("U_MEITMCD").Value
                If oRsGetMaterialInwardDtls.Fields.Item("U_MEBRCD").Value <> "" Then
                    oInventoryGenEntry.Lines.BarCode = oRsGetMaterialInwardDtls.Fields.Item("U_MEBRCD").Value
                End If
                oInventoryGenEntry.Lines.WarehouseCode = objForm.Items.Item("19").Specific.Value
                oInventoryGenEntry.Lines.Quantity = CDbl(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value)
                If oRsGetItemValue.RecordCount > 0 Then
                    'oInventoryGenEntry.Lines.Price = CDbl(oRsGetItemValue.Fields.Item("StockPrice").Value)
                    'oInventoryGenEntry.Lines.UnitPrice = CDbl(oRsGetItemValue.Fields.Item("StockPrice").Value)
                    oInventoryGenEntry.Lines.LineTotal = CDbl(oRsGetItemValue.Fields.Item("TransValue").Value)
                End If
                'oInventoryGenEntry.Lines.AccountCode = oRsGetAddonConfigurationDtls.Fields.Item("U_MEENTACT").Value
                Dim UomEntry As Integer = Me.GetUomEntry(oRsGetMaterialInwardDtls.Fields.Item("U_MEUOM").Value)
                If UomEntry = CInt(oRsGetMaterialInwardDtls.Fields.Item("DfltUOM").Value) Then
                    oInventoryGenEntry.Lines.UseBaseUnits = SAPbobsCOM.BoYesNoEnum.tYES
                    oInventoryGenEntry.Lines.UoMEntry = UomEntry
                    If oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "BatchNo" Then
                        oInventoryGenEntry.Lines.BatchNumbers.BatchNumber = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
                        oInventoryGenEntry.Lines.BatchNumbers.Quantity = CDbl(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value)
                        oInventoryGenEntry.Lines.BatchNumbers.Add()
                        If oRsGetMaterialInwardDtls.Fields.Item("U_CHKBIN").Value.Trim = "Y" Then
                            oInventoryGenEntry.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = binlineno
                            oInventoryGenEntry.Lines.BatchNumbers.BatchNumber = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
                            oInventoryGenEntry.Lines.BinAllocations.BinAbsEntry = oRsGetMaterialInwardDtls.Fields.Item("BinAbsEntry").Value
                            oInventoryGenEntry.Lines.BinAllocations.Quantity = CDbl(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value)
                            oInventoryGenEntry.Lines.BinAllocations.Add()
                        End If
                    ElseIf oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "SerialNo" Then
                        Dim TotalSlrBndBinWiseQty As String = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
                        Dim ArrayTotalSlrBndBinWiseQty() As String = TotalSlrBndBinWiseQty.Split(";")
                        For k As Integer = 0 To CInt(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value) - 1 'ArrayTotalSlrBndBinWiseQty.Length - 1
                            If ArrayTotalSlrBndBinWiseQty(k) <> "" Then
                                oInventoryGenEntry.Lines.SerialNumbers.ManufacturerSerialNumber = ArrayTotalSlrBndBinWiseQty(k)
                                oInventoryGenEntry.Lines.SerialNumbers.Add()
                                If oRsGetMaterialInwardDtls.Fields.Item("U_CHKBIN").Value.Trim = "Y" Then
                                    oInventoryGenEntry.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = binlineno
                                    oInventoryGenEntry.Lines.BatchNumbers.BatchNumber = ArrayTotalSlrBndBinWiseQty(k)
                                    oInventoryGenEntry.Lines.BinAllocations.BinAbsEntry = oRsGetMaterialInwardDtls.Fields.Item("BinAbsEntry").Value
                                    oInventoryGenEntry.Lines.BinAllocations.Quantity = 1 'CDbl(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value)
                                    oInventoryGenEntry.Lines.BinAllocations.Add()
                                    binlineno = binlineno + 1
                                End If
                            End If
                        Next
                    End If
                Else
                    oInventoryGenEntry.Lines.UseBaseUnits = SAPbobsCOM.BoYesNoEnum.tNO
                    oInventoryGenEntry.Lines.UoMEntry = UomEntry
                    If oRsGetItemValue.RecordCount > 0 Then
                        'Dim BaseQty As String = " Select ""BaseQty"" * " & oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value & "   From UGP1 T0 Inner Join OITM T1 On " & _
                        '                        " T1.""UgpEntry"" =  T0.""UgpEntry"" Where ""ItemCode"" = '" & _
                        '                       oRsGetMaterialInwardDtls.Fields.Item("U_MEITMCD").Value & "' And T0.""UomEntry"" = '" & UomEntry & "'"
                        'oRsBaseQty.DoQuery(BaseQty)
                        'oInventoryGenEntry.Lines.LineTotal = CDbl(oRsGetItemValue.Fields.Item("StockPrice").Value) * oRsBaseQty.Fields.Item(0).Value
                        oInventoryGenEntry.Lines.LineTotal = CDbl(oRsGetItemValue.Fields.Item("TransValue").Value)
                    End If

                    If oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "BatchNo" Then
                        oInventoryGenEntry.Lines.BatchNumbers.BatchNumber = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
                        oInventoryGenEntry.Lines.BatchNumbers.Quantity = oRsBaseQty.Fields.Item(0).Value
                        oInventoryGenEntry.Lines.BatchNumbers.Add()
                        If oRsGetMaterialInwardDtls.Fields.Item("U_CHKBIN").Value.Trim = "Y" Then
                            oInventoryGenEntry.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = binlineno
                            oInventoryGenEntry.Lines.BatchNumbers.BatchNumber = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
                            oInventoryGenEntry.Lines.BinAllocations.BinAbsEntry = oRsGetMaterialInwardDtls.Fields.Item("BinAbsEntry").Value
                            oInventoryGenEntry.Lines.BinAllocations.Quantity = oRsBaseQty.Fields.Item(0).Value
                            oInventoryGenEntry.Lines.BinAllocations.Add()
                        End If
                    ElseIf oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "SerialNo" Then
                        Dim TotalSlrBndBinWiseQty As String = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
                        Dim ArrayTotalSlrBndBinWiseQty() As String = TotalSlrBndBinWiseQty.Split(";")
                        For k As Integer = 0 To CInt(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value) - 1  ' ArrayTotalSlrBndBinWiseQty.Length - 1
                            If ArrayTotalSlrBndBinWiseQty(k) <> "" Then
                                oInventoryGenEntry.Lines.SerialNumbers.ManufacturerSerialNumber = ArrayTotalSlrBndBinWiseQty(k)
                                oInventoryGenEntry.Lines.SerialNumbers.Add()
                                If oRsGetMaterialInwardDtls.Fields.Item("U_CHKBIN").Value.Trim = "Y" Then
                                    Dim BaseQty As String = " Select ""BaseQty"" * " & oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value & "   From UGP1 T0 Inner Join OITM T1 On " &
                                                            " T1.""UgpEntry"" =  T0.""UgpEntry"" Where ""ItemCode"" = '" &
                                    oRsGetMaterialInwardDtls.Fields.Item("U_MEITMCD").Value & "' And T0.""UomEntry"" = '" & UomEntry & "'"
                                    oRsBaseQty.DoQuery(BaseQty)
                                    oInventoryGenEntry.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = binlineno
                                    oInventoryGenEntry.Lines.BatchNumbers.BatchNumber = ArrayTotalSlrBndBinWiseQty(k)
                                    oInventoryGenEntry.Lines.BinAllocations.BinAbsEntry = oRsGetMaterialInwardDtls.Fields.Item("BinAbsEntry").Value
                                    oInventoryGenEntry.Lines.BinAllocations.Quantity = 1
                                    oInventoryGenEntry.Lines.BinAllocations.Add()
                                    binlineno = binlineno + 1
                                End If
                            End If
                        Next
                    End If
                End If
                Try
                    oInventoryGenEntry.Lines.UserFields.Fields.Item("U_Addbarcode").Value = oRsGetMaterialInwardDtls.Fields.Item("U_ADBRCD").Value
                    oInventoryGenEntry.Lines.UserFields.Fields.Item("U_Packing").Value = oRsGetMaterialInwardDtls.Fields.Item("U_MEPKGE").Value
                Catch ex As Exception
                End Try
                binlineno = binlineno + 1
                oRsGetMaterialInwardDtls.MoveNext()
            Next
            Dim sGoodsReceiptGUID As String
            sGoodsReceiptGUID = System.Guid.NewGuid.ToString()
            Try
                oInventoryGenEntry.UserFields.Fields.Item("U_GUID").Value = sGoodsReceiptGUID
            Catch ex As Exception
            End Try
            Dim chk As String = oInventoryGenEntry.Add()
            If chk = 0 Then
                Dim GetOrderEntry As String = "Select ""DocEntry"" , ""DocNum""   From OIGN  Where  ""U_GUID"" = '" & sGoodsReceiptGUID & "' "
                Dim oRsGetOrderEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetOrderEntry.DoQuery(GetOrderEntry)

                Dim MIUpdate As String = "Update ""@ME_MI"" Set  ""U_MEGRNO"" = '" & oRsGetOrderEntry.Fields.Item(1).Value.ToString & "' , ""U_MEGREN"" = '" & oRsGetOrderEntry.Fields.Item(0).Value.ToString & "' , ""U_REMARKS"" = ' ' , ""U_MESTS"" = 'Close'  Where ""DocEntry"" = '" & oDBs_Head.GetValue("DocEntry", 0) & "' "
                Dim oRsMIUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsMIUpdate.DoQuery(MIUpdate)

                objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_MI"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_MI_C0"), "DocEntry", oDBs_Details1.GetValue("DocEntry", 0))
                objMatrix1.LoadFromDataSource()
                objForm.Refresh()
                'Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'Dim Update As String = " UPDATE T1 SET T1.""U_MEGRNO""=(SELECT max(T0.""DocEntry"") FROM OIGN T0), T1.""U_MEGREN""=(SELECT ""DocNum""  FROM OIGN Where ""DocEntry"" = (SELECT max(T0.""DocEntry"") FROM OIGN T0)) From ""@ME_MI"" T1 WHERE T1.""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0) & "' "
                'oRsUpdate.DoQuery(Update)
                Dim GRUpdate As String = "Update ""OIGN"" Set  ""U_MIDNUM"" = '" & oDBs_Head.GetValue("DocNum", 0).ToString & "' , ""U_MIDENT"" = '" & oDBs_Head.GetValue("DocEntry", 0).ToString & "'  Where ""DocEntry"" = '" & oRsGetOrderEntry.Fields.Item(0).Value.ToString & "' "
                Dim oRsGRUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGRUpdate.DoQuery(GRUpdate)


                ' Me.ModeMask(objForm.UniqueID)
            Else
                objMain.sCmp = objMain.objCompany.GetCompanyService
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OMI")
                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                objMain.oGeneralData.SetProperty("U_REMARKS", "Receipt : " & objMain.objCompany.GetLastErrorDescription.ToString)
                objMain.oGeneralService.Update(objMain.oGeneralData)
                objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_MI"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_MI_C0"), "DocEntry", oDBs_Details1.GetValue("DocEntry", 0))
                objMatrix1.LoadFromDataSource()
                objForm.Refresh()
                objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oInventoryGenEntry)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

#Region "Old Issue Receipt on Single Account Configuration"

    'Public Sub GoodsIssue(ByVal FormUID As String)
    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)
    '        oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_MI")
    '        oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_MI_C0")
    '        objMatrix1 = objForm.Items.Item("3").Specific

    '        Dim GetMaterialInwardDtls As String = "Select (Select ""IUoMEntry"" From OITM Where ""ItemCode"" = ""U_MEITMCD"") As ""DfltUOM"" , *  From ""@ME_MI_C0"" Where ""DocEntry"" = '" & oDBs_Head.GetValue("DocEntry", 0) & "' And  U_MERECQTY > 0  "
    '        Dim oRsGetMaterialInwardDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oRsGetMaterialInwardDtls.DoQuery(GetMaterialInwardDtls)

    '        Dim GetAddonConfigurationDtls As String = "Select ""U_MEEXACCT"" From ""@ME_ACFG"" "
    '        Dim oRsGetAddonConfigurationDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oRsGetAddonConfigurationDtls.DoQuery(GetAddonConfigurationDtls)

    '        Dim oInventoryGenExit As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)

    '        Try
    '            oInventoryGenExit.UserFields.Fields.Item("U_TYPE").Value = "STR"
    '        Catch ex As Exception
    '        End Try

    '        Dim sDate As String = objForm.Items.Item("23").Specific.Value
    '        sDate = sDate.Insert("4", "-")
    '        sDate = sDate.Insert("7", "-")
    '        oInventoryGenExit.DocDate = sDate
    '        oInventoryGenExit.DocDueDate = sDate
    '        oInventoryGenExit.TaxDate = sDate
    '        objComboBox = objForm.Items.Item("42").Specific
    '        Dim GetSeries As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = '60' And ""Locked"" = 'N'  And ""BPLId"" = '" & objComboBox.Value & "' "
    '        Dim oRsGetSeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oRsGetSeries.DoQuery(GetSeries)
    '        If oRsGetSeries.RecordCount > 0 Then
    '            oInventoryGenExit.Series = oRsGetSeries.Fields.Item("Series").Value
    '        End If
    '        oInventoryGenExit.BPL_IDAssignedToInvoice = objComboBox.Value
    '        For i As Integer = 1 To oRsGetMaterialInwardDtls.RecordCount
    '            If i <> 1 Then
    '                oInventoryGenExit.Lines.Add()
    '            End If
    '            oInventoryGenExit.Lines.ItemCode = oRsGetMaterialInwardDtls.Fields.Item("U_MEITMCD").Value
    '            If oRsGetMaterialInwardDtls.Fields.Item("U_MEBRCD").Value <> "" Then
    '                oInventoryGenExit.Lines.BarCode = oRsGetMaterialInwardDtls.Fields.Item("U_MEBRCD").Value
    '            End If

    '            objComboBox = objForm.Items.Item("52").Specific

    '            If objComboBox.Value = "ST" Then
    '                oInventoryGenExit.Lines.WarehouseCode = objForm.Items.Item("33").Specific.Value
    '            Else
    '                oInventoryGenExit.Lines.WarehouseCode = objForm.Items.Item("17").Specific.Value
    '            End If

    '            Dim UomEntry As Integer = Me.GetUomEntry(oRsGetMaterialInwardDtls.Fields.Item("U_MEUOM").Value)
    '            If UomEntry = CInt(oRsGetMaterialInwardDtls.Fields.Item("DfltUOM").Value) Then
    '                oInventoryGenExit.Lines.UseBaseUnits = SAPbobsCOM.BoYesNoEnum.tYES
    '                oInventoryGenExit.Lines.UoMEntry = UomEntry
    '                If oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "BatchNo" Then
    '                    oInventoryGenExit.Lines.BatchNumbers.BatchNumber = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
    '                    oInventoryGenExit.Lines.BatchNumbers.Quantity = CDbl(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value)
    '                    oInventoryGenExit.Lines.BatchNumbers.Add()
    '                ElseIf oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "SerialNo" Then
    '                    Dim TotalSlrBndBinWiseQty As String = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
    '                    Dim ArrayTotalSlrBndBinWiseQty() As String = TotalSlrBndBinWiseQty.Split(";")
    '                    For k As Integer = 0 To CInt(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value) - 1 'ArrayTotalSlrBndBinWiseQty.Length - 1
    '                        If ArrayTotalSlrBndBinWiseQty(k) <> "" Then
    '                            oInventoryGenExit.Lines.SerialNumbers.ManufacturerSerialNumber = ArrayTotalSlrBndBinWiseQty(k)
    '                            oInventoryGenExit.Lines.SerialNumbers.Add()
    '                        End If
    '                    Next
    '                End If
    '            Else
    '                oInventoryGenExit.Lines.UseBaseUnits = SAPbobsCOM.BoYesNoEnum.tNO
    '                oInventoryGenExit.Lines.UoMEntry = UomEntry
    '                If oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "BatchNo" Then
    '                    Dim BaseQty As String = " Select ""BaseQty"" * " & oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value & "   From UGP1 T0 Inner Join OITM T1 On " & _
    '                                            " T1.""UgpEntry"" =  T0.""UgpEntry"" Where ""ItemCode"" = '" & _
    '                                            oRsGetMaterialInwardDtls.Fields.Item("U_MEITMCD").Value & "' And T0.""UomEntry"" = '" & UomEntry & "'"
    '                    Dim oRsBaseQty As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                    oRsBaseQty.DoQuery(BaseQty)
    '                    oInventoryGenExit.Lines.BatchNumbers.BatchNumber = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
    '                    oInventoryGenExit.Lines.BatchNumbers.Quantity = oRsBaseQty.Fields.Item(0).Value
    '                    oInventoryGenExit.Lines.BatchNumbers.Add()
    '                ElseIf oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "SerialNo" Then
    '                    Dim TotalSlrBndBinWiseQty As String = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
    '                    Dim ArrayTotalSlrBndBinWiseQty() As String = TotalSlrBndBinWiseQty.Split(";")
    '                    For k As Integer = 0 To CInt(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value) - 1 ' ArrayTotalSlrBndBinWiseQty.Length - 1
    '                        If ArrayTotalSlrBndBinWiseQty(k) <> "" Then
    '                            oInventoryGenExit.Lines.SerialNumbers.ManufacturerSerialNumber = ArrayTotalSlrBndBinWiseQty(k)
    '                            oInventoryGenExit.Lines.SerialNumbers.Add()
    '                        End If
    '                    Next
    '                End If
    '            End If
    '            oInventoryGenExit.Lines.Quantity = CDbl(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value)
    '            oInventoryGenExit.Lines.AccountCode = oRsGetAddonConfigurationDtls.Fields.Item("U_MEEXACCT").Value
    '            Try
    '                oInventoryGenExit.Lines.UserFields.Fields.Item("U_Addbarcode").Value = oRsGetMaterialInwardDtls.Fields.Item("U_ADBRCD").Value
    '                oInventoryGenExit.Lines.UserFields.Fields.Item("U_Packing").Value = oRsGetMaterialInwardDtls.Fields.Item("U_MEPKGE").Value
    '            Catch ex As Exception
    '            End Try
    '            oRsGetMaterialInwardDtls.MoveNext()
    '        Next

    '        If oInventoryGenExit.Add() = 0 Then
    '            Dim GetOrderEntry As String = "Select ""DocEntry"" , ""DocNum""   From OIGE Where ""DocEntry"" = (Select MAX(""DocEntry"")  From OIGE)"
    '            Dim oRsGetOrderEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '            oRsGetOrderEntry.DoQuery(GetOrderEntry)

    '            objMain.sCmp = objMain.objCompany.GetCompanyService
    '            objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OMI")
    '            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
    '            objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0))
    '            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
    '            objMain.oGeneralData.SetProperty("U_MEGINO", oRsGetOrderEntry.Fields.Item(1).Value.ToString)
    '            objMain.oGeneralData.SetProperty("U_MEGIEN", oRsGetOrderEntry.Fields.Item(0).Value.ToString)
    '            objMain.oGeneralData.SetProperty("U_REMARKS", "")
    '            objMain.oGeneralService.Update(objMain.oGeneralData)
    '            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_MI"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
    '            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_MI_C0"), "DocEntry", oDBs_Details1.GetValue("DocEntry", 0))
    '            objMatrix1.LoadFromDataSource()
    '            objForm.Refresh()
    '            objMain.objApplication.StatusBar.SetText("Goods Receipt posting in progress......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '            If objForm.Items.Item("11").Specific.Value = "" Then
    '                Me.GoodsReceipt(objForm.UniqueID)
    '            Else
    '            End If
    '        Else
    '            objMain.sCmp = objMain.objCompany.GetCompanyService
    '            objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OMI")
    '            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
    '            objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0))
    '            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
    '            objMain.oGeneralData.SetProperty("U_REMARKS", "ISSUE : " & objMain.objCompany.GetLastErrorDescription.ToString)
    '            objMain.oGeneralService.Update(objMain.oGeneralData)
    '            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_MI"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
    '            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_MI_C0"), "DocEntry", oDBs_Details1.GetValue("DocEntry", 0))
    '            objMatrix1.LoadFromDataSource()
    '            objForm.Refresh()
    '            objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        End If

    '    Catch ex As Exception

    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub

    'Public Sub GoodsReceipt(ByVal FormUID As String)
    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)
    '        oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_MI")
    '        oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_MI_C0")
    '        objMatrix1 = objForm.Items.Item("3").Specific

    '        Dim oRsBaseQty As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '        Dim GetMaterialInwardDtls As String = "Select  (Select ""IUoMEntry"" From OITM Where ""ItemCode"" = ""U_MEITMCD"") As ""DfltUOM"" , (Select ""AbsEntry"" From OBIN Where ""BinCode"" = ""U_METWBL"") As ""BinAbsEntry"" ,  * From ""@ME_MI_C0"" Where ""DocEntry"" = '" & oDBs_Head.GetValue("DocEntry", 0) & "'  And  U_MERECQTY > 0  "
    '        Dim oRsGetMaterialInwardDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oRsGetMaterialInwardDtls.DoQuery(GetMaterialInwardDtls)

    '        Dim GetAddonConfigurationDtls As String = "Select ""U_MEENTACT"" , ""U_METWHSCD"" From ""@ME_ACFG"" "
    '        Dim oRsGetAddonConfigurationDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oRsGetAddonConfigurationDtls.DoQuery(GetAddonConfigurationDtls)

    '        Dim oInventoryGenEntry As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)

    '        Try
    '            oInventoryGenEntry.UserFields.Fields.Item("U_TYPE").Value = "STR"
    '        Catch ex As Exception
    '        End Try

    '        Dim sDate As String = objForm.Items.Item("23").Specific.Value
    '        sDate = sDate.Insert("4", "-")
    '        sDate = sDate.Insert("7", "-")

    '        oInventoryGenEntry.DocDate = sDate
    '        oInventoryGenEntry.DocDueDate = sDate
    '        oInventoryGenEntry.TaxDate = sDate

    '        objComboBox = objForm.Items.Item("40").Specific

    '        oInventoryGenEntry.BPL_IDAssignedToInvoice = objComboBox.Value

    '        Dim GetSeries As String = " Select ""Series"" , ""SeriesName""   From NNM1 Where ""ObjectCode"" = '59' And ""Locked"" = 'N' And ""BPLId"" =  '" & objComboBox.Value & "' "
    '        Dim oRsGetSeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oRsGetSeries.DoQuery(GetSeries)

    '        oInventoryGenEntry.Series = oRsGetSeries.Fields.Item("Series").Value

    '        For i As Integer = 1 To oRsGetMaterialInwardDtls.RecordCount
    '            Dim binlineno As Integer = 0
    '            If i <> 1 Then
    '                oInventoryGenEntry.Lines.Add()
    '            End If
    '            Dim GetItemValue As String = "Select ""StockPrice"" From IGE1 Where ""DocEntry"" = '" & objForm.Items.Item("29").Specific.Value _
    '                                         & "' And ""ItemCode""  = '" & oRsGetMaterialInwardDtls.Fields.Item("U_MEITMCD").Value & "'"
    '            Dim oRsGetItemValue As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '            oRsGetItemValue.DoQuery(GetItemValue)
    '            oInventoryGenEntry.Lines.ItemCode = oRsGetMaterialInwardDtls.Fields.Item("U_MEITMCD").Value
    '            If oRsGetMaterialInwardDtls.Fields.Item("U_MEBRCD").Value <> "" Then
    '                oInventoryGenEntry.Lines.BarCode = oRsGetMaterialInwardDtls.Fields.Item("U_MEBRCD").Value
    '            End If
    '            oInventoryGenEntry.Lines.WarehouseCode = objForm.Items.Item("19").Specific.Value
    '            oInventoryGenEntry.Lines.Quantity = CDbl(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value)
    '            If oRsGetItemValue.RecordCount > 0 Then
    '                oInventoryGenEntry.Lines.Price = CDbl(oRsGetItemValue.Fields.Item("StockPrice").Value)
    '                oInventoryGenEntry.Lines.UnitPrice = CDbl(oRsGetItemValue.Fields.Item("StockPrice").Value)
    '            End If
    '            oInventoryGenEntry.Lines.AccountCode = oRsGetAddonConfigurationDtls.Fields.Item("U_MEENTACT").Value
    '            Dim UomEntry As Integer = Me.GetUomEntry(oRsGetMaterialInwardDtls.Fields.Item("U_MEUOM").Value)
    '            If UomEntry = CInt(oRsGetMaterialInwardDtls.Fields.Item("DfltUOM").Value) Then
    '                oInventoryGenEntry.Lines.UseBaseUnits = SAPbobsCOM.BoYesNoEnum.tYES
    '                oInventoryGenEntry.Lines.UoMEntry = UomEntry
    '                If oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "BatchNo" Then
    '                    oInventoryGenEntry.Lines.BatchNumbers.BatchNumber = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
    '                    oInventoryGenEntry.Lines.BatchNumbers.Quantity = CDbl(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value)
    '                    oInventoryGenEntry.Lines.BatchNumbers.Add()
    '                    If oRsGetMaterialInwardDtls.Fields.Item("U_CHKBIN").Value.Trim = "Y" Then
    '                        oInventoryGenEntry.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = binlineno
    '                        oInventoryGenEntry.Lines.BatchNumbers.BatchNumber = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
    '                        oInventoryGenEntry.Lines.BinAllocations.BinAbsEntry = oRsGetMaterialInwardDtls.Fields.Item("BinAbsEntry").Value
    '                        oInventoryGenEntry.Lines.BinAllocations.Quantity = CDbl(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value)
    '                        oInventoryGenEntry.Lines.BinAllocations.Add()
    '                    End If
    '                ElseIf oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "SerialNo" Then
    '                    Dim TotalSlrBndBinWiseQty As String = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
    '                    Dim ArrayTotalSlrBndBinWiseQty() As String = TotalSlrBndBinWiseQty.Split(";")
    '                    For k As Integer = 0 To CInt(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value) - 1 'ArrayTotalSlrBndBinWiseQty.Length - 1
    '                        If ArrayTotalSlrBndBinWiseQty(k) <> "" Then
    '                            oInventoryGenEntry.Lines.SerialNumbers.ManufacturerSerialNumber = ArrayTotalSlrBndBinWiseQty(k)
    '                            oInventoryGenEntry.Lines.SerialNumbers.Add()
    '                            If oRsGetMaterialInwardDtls.Fields.Item("U_CHKBIN").Value.Trim = "Y" Then
    '                                oInventoryGenEntry.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = binlineno
    '                                oInventoryGenEntry.Lines.BatchNumbers.BatchNumber = ArrayTotalSlrBndBinWiseQty(k)
    '                                oInventoryGenEntry.Lines.BinAllocations.BinAbsEntry = oRsGetMaterialInwardDtls.Fields.Item("BinAbsEntry").Value
    '                                oInventoryGenEntry.Lines.BinAllocations.Quantity = 1 'CDbl(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value)
    '                                oInventoryGenEntry.Lines.BinAllocations.Add()
    '                                binlineno = binlineno + 1
    '                            End If
    '                        End If
    '                    Next
    '                End If
    '            Else
    '                oInventoryGenEntry.Lines.UseBaseUnits = SAPbobsCOM.BoYesNoEnum.tNO
    '                oInventoryGenEntry.Lines.UoMEntry = UomEntry
    '                If oRsGetItemValue.RecordCount > 0 Then
    '                    Dim BaseQty As String = " Select ""BaseQty"" * " & oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value & "   From UGP1 T0 Inner Join OITM T1 On " & _
    '                                            " T1.""UgpEntry"" =  T0.""UgpEntry"" Where ""ItemCode"" = '" & _
    '                                           oRsGetMaterialInwardDtls.Fields.Item("U_MEITMCD").Value & "' And T0.""UomEntry"" = '" & UomEntry & "'"
    '                    oRsBaseQty.DoQuery(BaseQty)
    '                    oInventoryGenEntry.Lines.LineTotal = CDbl(oRsGetItemValue.Fields.Item("StockPrice").Value) * oRsBaseQty.Fields.Item(0).Value
    '                End If

    '                If oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "BatchNo" Then
    '                    oInventoryGenEntry.Lines.BatchNumbers.BatchNumber = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
    '                    oInventoryGenEntry.Lines.BatchNumbers.Quantity = oRsBaseQty.Fields.Item(0).Value
    '                    oInventoryGenEntry.Lines.BatchNumbers.Add()
    '                    If oRsGetMaterialInwardDtls.Fields.Item("U_CHKBIN").Value.Trim = "Y" Then
    '                        oInventoryGenEntry.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = binlineno
    '                        oInventoryGenEntry.Lines.BatchNumbers.BatchNumber = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
    '                        oInventoryGenEntry.Lines.BinAllocations.BinAbsEntry = oRsGetMaterialInwardDtls.Fields.Item("BinAbsEntry").Value
    '                        oInventoryGenEntry.Lines.BinAllocations.Quantity = oRsBaseQty.Fields.Item(0).Value
    '                        oInventoryGenEntry.Lines.BinAllocations.Add()
    '                    End If
    '                ElseIf oRsGetMaterialInwardDtls.Fields.Item("U_METYPE").Value.Trim = "SerialNo" Then
    '                    Dim TotalSlrBndBinWiseQty As String = oRsGetMaterialInwardDtls.Fields.Item("U_MESRBT").Value.Trim
    '                    Dim ArrayTotalSlrBndBinWiseQty() As String = TotalSlrBndBinWiseQty.Split(";")
    '                    For k As Integer = 0 To CInt(oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value) - 1  ' ArrayTotalSlrBndBinWiseQty.Length - 1
    '                        If ArrayTotalSlrBndBinWiseQty(k) <> "" Then
    '                            oInventoryGenEntry.Lines.SerialNumbers.ManufacturerSerialNumber = ArrayTotalSlrBndBinWiseQty(k)
    '                            oInventoryGenEntry.Lines.SerialNumbers.Add()
    '                            If oRsGetMaterialInwardDtls.Fields.Item("U_CHKBIN").Value.Trim = "Y" Then
    '                                Dim BaseQty As String = " Select ""BaseQty"" * " & oRsGetMaterialInwardDtls.Fields.Item("U_MERECQTY").Value & "   From UGP1 T0 Inner Join OITM T1 On " & _
    '                                                 " T1.""UgpEntry"" =  T0.""UgpEntry"" Where ""ItemCode"" = '" & _
    '                                                  oRsGetMaterialInwardDtls.Fields.Item("U_MEITMCD").Value & "' And T0.""UomEntry"" = '" & UomEntry & "'"
    '                                oRsBaseQty.DoQuery(BaseQty)
    '                                oInventoryGenEntry.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = binlineno
    '                                oInventoryGenEntry.Lines.BatchNumbers.BatchNumber = ArrayTotalSlrBndBinWiseQty(k)
    '                                oInventoryGenEntry.Lines.BinAllocations.BinAbsEntry = oRsGetMaterialInwardDtls.Fields.Item("BinAbsEntry").Value
    '                                oInventoryGenEntry.Lines.BinAllocations.Quantity = 1 'oRsBaseQty.Fields.Item(0).Value
    '                                oInventoryGenEntry.Lines.BinAllocations.Add()
    '                                binlineno = binlineno + 1
    '                            End If
    '                        End If
    '                    Next
    '                End If
    '            End If
    '            Try
    '                oInventoryGenEntry.Lines.UserFields.Fields.Item("U_Addbarcode").Value = oRsGetMaterialInwardDtls.Fields.Item("U_ADBRCD").Value
    '                oInventoryGenEntry.Lines.UserFields.Fields.Item("U_Packing").Value = oRsGetMaterialInwardDtls.Fields.Item("U_MEPKGE").Value
    '            Catch ex As Exception
    '            End Try
    '            binlineno = binlineno + 1
    '            oRsGetMaterialInwardDtls.MoveNext()
    '        Next
    '        If oInventoryGenEntry.Add() = 0 Then
    '            Dim GetOrderEntry As String = "Select ""DocEntry"" , ""DocNum""   From OIGN Where ""DocEntry"" = (Select MAX(""DocEntry"")  From OIGN)"
    '            Dim oRsGetOrderEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '            oRsGetOrderEntry.DoQuery(GetOrderEntry)
    '            objMain.sCmp = objMain.objCompany.GetCompanyService
    '            objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OMI")
    '            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
    '            objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0))
    '            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
    '            objMain.oGeneralData.SetProperty("U_MEGRNO", oRsGetOrderEntry.Fields.Item(1).Value.ToString)
    '            objMain.oGeneralData.SetProperty("U_MEGREN", oRsGetOrderEntry.Fields.Item(0).Value.ToString)
    '            objMain.oGeneralData.SetProperty("U_MESTS", "Close")
    '            objMain.oGeneralData.SetProperty("U_REMARKS", "")
    '            objMain.oGeneralService.Update(objMain.oGeneralData)
    '            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_MI"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
    '            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_MI_C0"), "DocEntry", oDBs_Details1.GetValue("DocEntry", 0))
    '            objMatrix1.LoadFromDataSource()
    '            objForm.Refresh()
    '            Me.ModeMask(objForm.UniqueID)
    '        Else
    '            objMain.sCmp = objMain.objCompany.GetCompanyService
    '            objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OMI")
    '            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
    '            objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0))
    '            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
    '            objMain.oGeneralData.SetProperty("U_REMARKS", "Receipt : " & objMain.objCompany.GetLastErrorDescription.ToString)
    '            objMain.oGeneralService.Update(objMain.oGeneralData)
    '            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_MI"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
    '            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_MI_C0"), "DocEntry", oDBs_Details1.GetValue("DocEntry", 0))
    '            objMatrix1.LoadFromDataSource()
    '            objForm.Refresh()
    '            objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        End If
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub

#End Region

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("ME_MI_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = True Then
                        If Me.Validation(objForm.UniqueID) = False Then
                            BubbleEvent = False
                            Exit Try
                        End If
                    End If
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        'If objForm.Items.Item("9").Specific.Value = "" And objForm.Items.Item("53").Specific.Value.Trim = "Open" Then
                        '    Me.GoodsIssue(objForm.UniqueID)
                        'End If
                        'objComboBox = objForm.Items.Item("53").Specific
                        'If objComboBox.Value = "" Then
                        'Me.UpdateDocStatus(objForm.UniqueID)
                        'End If
                        'Dim UpdateMISTS As String = "Update ""@ME_MI"" Set ""U_MESTS"" = 'Open' Where ""DocEntry"" = '" & oDBs_Head.GetValue("DocEntry", 0) & "'"
                        'Dim oRsUpdateMISTS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'oRsUpdateMISTS.DoQuery(UpdateMISTS)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    'If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                    '    If objForm.Items.Item("9").Specific.Value = "" And objForm.Items.Item("53").Specific.Value.Trim = "Open" Then
                    '        Me.GoodsIssue(objForm.UniqueID)
                    '    ElseIf objForm.Items.Item("9").Specific.Value <> "" And objForm.Items.Item("11").Specific.Value = "" And objForm.Items.Item("53").Specific.Value.Trim = "Open" Then
                    '        Me.GoodsReceipt(objForm.UniqueID)
                    '    End If
                    'End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        'If objForm.Items.Item("9").Specific.Value <> "" And objForm.Items.Item("11").Specific.Value <> "" Then
                        '    'Me.ModeMask(objForm.UniqueID)
                        '    'Else
                        '    '    Me.ModeMask(objForm.UniqueID)
                        'End If
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CFLFilterSTR(ByVal FormUID As String, ByVal CFL_ID As String, ByVal BPId As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim GetFreeItems As String = " Select ""DocNum"" From ""@ME_STR"" Where ""U_METBPID"" = '" & BPId & "' "
            Dim oRsGetFreeItems As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetFreeItems.DoQuery(GetFreeItems)

            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            If oRsGetFreeItems.RecordCount > 0 Then
                oCondition = oConditions.Add()
                oCondition.Alias = "DocNum"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = oRsGetFreeItems.Fields.Item(0).Value
                oChooseFromList.SetConditions(oConditions)
                oRsGetFreeItems.MoveNext()
                For i As Integer = 1 To oRsGetFreeItems.RecordCount - 1
                    oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCondition = oConditions.Add()
                    oCondition.Alias = "DocNum"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = oRsGetFreeItems.Fields.Item(0).Value
                    oChooseFromList.SetConditions(oConditions)
                    oRsGetFreeItems.MoveNext()
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

    Sub CFLFilterITR(ByVal FormUID As String, ByVal CFL_ID As String, ByVal STRNo As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim GetFreeItems As String = " Select ""DocNum"" From ""OIGE"" Where ""U_INVTRREQ"" = '" & STRNo & "'  "
            Dim oRsGetFreeItems As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetFreeItems.DoQuery(GetFreeItems)

            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            If oRsGetFreeItems.RecordCount > 0 Then
                oCondition = oConditions.Add()
                oCondition.Alias = "DocNum"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = oRsGetFreeItems.Fields.Item(0).Value
                oChooseFromList.SetConditions(oConditions)
                oRsGetFreeItems.MoveNext()
                For i As Integer = 1 To oRsGetFreeItems.RecordCount - 1
                    oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCondition = oConditions.Add()
                    oCondition.Alias = "DocNum"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = oRsGetFreeItems.Fields.Item(0).Value
                    oChooseFromList.SetConditions(oConditions)
                    oRsGetFreeItems.MoveNext()
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

    Sub CFLEmpFilter(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            oCondition = oConditions.Add()
            oCondition.Alias = "Active"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "Y"
            oChooseFromList.SetConditions(oConditions)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub ModeMask(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            objMatrix1 = objForm.Items.Item("3").Specific
            If objForm.Items.Item("11").Specific.Value <> "" Then
                For i As Integer = 1 To objMatrix1.VisualRowCount
                    objMatrix1.CommonSetting.SetRowEditable(i, False)
                Next
                objForm.Items.Item("27").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("27").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("52").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("52").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                ' objForm.Items.Item("7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("24").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Else
                For i As Integer = 1 To objMatrix1.VisualRowCount
                    objMatrix1.CommonSetting.SetRowEditable(i, True)
                Next
                objMatrix1.Columns.Item("V_0").Editable = False
                objMatrix1.Columns.Item("V_1").Editable = False
                objMatrix1.Columns.Item("V_2").Editable = False
                objMatrix1.Columns.Item("V_3").Editable = False
                objMatrix1.Columns.Item("V_4").Editable = False

                objForm.Items.Item("27").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                objForm.Items.Item("5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                'objForm.Items.Item("7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                objForm.Items.Item("24").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            End If
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CFLWhsCodeFilter(ByVal FormUID As String, ByVal CFL_ID As String, ByVal BPLid As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim GetWhsCode As String = " Select ""WhsCode"" From OWHS Where ""BPLid"" = '" & BPLid & "' "
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

    Public Function GetUomEntry(ByVal sUomCode As String) As Integer
        Try
            Dim GetDocEntry As String = "Select ""UomEntry"" From OUOM Where ""UomCode"" = '" & sUomCode & "' "
            Dim oRsGetDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDocEntry.DoQuery(GetDocEntry)

            Return CInt(oRsGetDocEntry.Fields.Item("UoMEntry").Value)
        Catch ex As Exception
        End Try
    End Function

    Public Function GetUgpEntry(ByVal sItemCode As String) As Integer
        Try
            Dim GetDocEntry As String = "Select ""UgpEntry"" From OITM Where ""ItemCode"" = '" & sItemCode & "' "
            Dim oRsGetDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDocEntry.DoQuery(GetDocEntry)

            Return CInt(oRsGetDocEntry.Fields.Item("UgpEntry").Value)
        Catch ex As Exception
        End Try
    End Function

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
                .UserID = "manager"
                .Password = "Abcd@1234"
            End With

            cryRpt.Load(oRsGetDBPassword.Fields.Item("U_MEATCH").Value.Trim & "\" & ReportName)

            CrTables = cryRpt.Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
            Next

            Dim DocNo As String = objForm.Items.Item("13").Specific.value

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

    Function GetVarianceQty(ByVal FormUID As String, ByVal ItemCode As String, ByVal RecQty As Double) As Double
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim VarQty As Double = 0
            Dim GetVarQty As String = "Select ""Quantity"" From WTR1 Where ""DocEntry"" = '" & objForm.Items.Item("46").Specific.Value & "' And ""ItemCode"" = '" & ItemCode & "'"
            Dim oRsGetVarQty As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetVarQty.DoQuery(GetVarQty)
            If oRsGetVarQty.RecordCount > 0 Then
                VarQty = oRsGetVarQty.Fields.Item("Quantity").Value - RecQty
            End If
            Return VarQty
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

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

    Private Sub UpdateDocStatus(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMain.sCmp = objMain.objCompany.GetCompanyService
            objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OMI")
            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0))
            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
            objMain.oGeneralData.SetProperty("U_MESTS", "Open")
            objMain.oGeneralService.Update(objMain.oGeneralData)
            objMain.objUtilities.RefreshDatasourceFromDB(objForm.UniqueID, objForm.DataSources.DBDataSources.Item("@ME_MI"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
            objMain.objUtilities.RefreshDatasourceFromDB(objForm.UniqueID, objForm.DataSources.DBDataSources.Item("@ME_MI_C0"), "DocEntry", oDBs_Details1.GetValue("DocEntry", 0))
            objMatrix1.LoadFromDataSource()
            objForm.Refresh()
        Catch ex As Exception
        End Try
    End Sub

End Class
