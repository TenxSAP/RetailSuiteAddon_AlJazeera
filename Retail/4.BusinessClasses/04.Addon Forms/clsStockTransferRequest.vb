Imports System.Threading
Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class clsStockTransferRequest

#Region "       Declaration             "
    Public objForm, objITRForm, objITRFormUDF As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details1, oDBs_Details2 As SAPbouiCOM.DBDataSource
    Dim objMatrix1, objMatrix2, objITRMatrix As SAPbouiCOM.Matrix
    Dim objComboBox, objComboBox2 As SAPbouiCOM.ComboBox
    Dim LostFocusFlag As Boolean = False
    Dim BaseUom As String
#End Region

    Sub CreateForm()
        Try
            'enable print layout
            'If objMain IsNot Nothing AndAlso objMain.objCompany IsNot Nothing Then
            '    ' use the main connected company object instead of a newly constructed one
            '    clsStockTransferRequest.objCompany = objMain.objCompany
            'End If

            objMain.objUtilities.LoadForm("StockTransferRequest1.xml", "ME_STR_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_STR_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objMain.objApplication.StatusBar.SetText("Stock Transfer Request form is initializing, Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objForm.Freeze(True)

            objForm.Items.Item("27").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_STR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_STR_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@ME_STR_C1")

            objMatrix1 = objForm.Items.Item("3").Specific
            objMatrix2 = objForm.Items.Item("4").Specific

            objComboBox = objForm.Items.Item("51").Specific
            objComboBox.ValidValues.Add("GR", "GRPO")
            objComboBox.ValidValues.Add("ST", "Standard")
            objComboBox.ValidValues.Add("PR", "PurchaseRequest")
            objForm.Items.Item("51").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objComboBox.Select("ST", SAPbouiCOM.BoSearchKey.psk_ByValue)

            objComboBox = objForm.Items.Item("21").Specific
            objComboBox.ValidValues.Add("", "")
            objComboBox.ValidValues.Add("Open", "Open")
            objComboBox.ValidValues.Add("Close", "Close")
            objComboBox.ValidValues.Add("Cancel", "Cancel")

            objComboBox = objForm.Items.Item("16").Specific
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, "Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join " & " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""    ") '" Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join " & _
            '" OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""  ")
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("16").DisplayDesc = True
            If objComboBox.ValidValues.Count <> 0 Then
                objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
                Dim QUERY As String = "SELECT  T0.""WhsCode"", T0.""WhsName"" FROM ""OWHS""  T0 WHERE T0.""BPLid""='" & objComboBox.Value & "'"
                objComboBox = objForm.Items.Item("6").Specific
                objMain.objUtilities.ComboBoxLoadValues(objComboBox, QUERY) '" Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join " & _
                '" OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""  ")
                objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                objForm.Items.Item("6").DisplayDesc = True
                '  Me.CFLWhsCodeFilter(objForm.UniqueID, "CFL_FWHSCD", objComboBox.Value)
            End If

            objMain.objUtilities.ComboBoxLoadValues(objForm.Items.Item("14").Specific, "Select ""BPLId"" ,  ""BPLName"" From OBPL  ")
            objForm.Items.Item("14").DisplayDesc = True


            ''Enabling Print Layout Icon
            'Try
            '    objForm.ReportType = "ME_STR_Form"

            '    Try
            '        CreateLayoutFolder_StockTransferRequest()
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

            Me.CellsMasking(objForm.UniqueID)

            Me.CFLEmpFilter(objForm.UniqueID, "CFL_EMPID")
            'objForm.Items.Item("26").Visible = False
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

            'objForm.Items.Item("31").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            ''objForm.Items.Item("31").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("31").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("51").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("51").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("51").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("25").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("25").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("25").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("55").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("55").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("32").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("32").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("18").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("18").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("21").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("21").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("GINum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("GINum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("GIDoc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("GIDoc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("48").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("48").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("49").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("49").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("14").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("14").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("14").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("16").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("16").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("16").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("58").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("58").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("1000005").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("1000005").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("1000005").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "", Optional ByVal Series As Integer = 0)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_STR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_STR_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@ME_STR_C1")

            objMatrix1 = objForm.Items.Item("3").Specific
            objMatrix2 = objForm.Items.Item("4").Specific

            objComboBox = objForm.Items.Item("19").Specific
            Dim GetSeriesQty As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'ME_OSTR' And ""Locked"" = 'N'"
            objMain.objUtilities.ComboBoxLoadValuesSeries(objForm.Items.Item("19").Specific, GetSeriesQty)
            objForm.Items.Item("19").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly


            'Dim GetDflt As String = "Select ""DfltSeries"" From ONNM Where ""ObjectCode"" = 'ME_OSTR'"

            Dim GetDflt As String =
    "SELECT ""Series"", ""SeriesName"", ""ObjectCode"", ""Indicator"" " &
    "FROM NNM1 " &
    "WHERE ""ObjectCode"" IN ('ME_OSTR', '@ME_STR') " &
    "AND ""Indicator"" <> 'Default'"




            Dim oRsGetDflt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDflt.DoQuery(GetDflt)
            If Flag = "" Then
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNumSeries(objForm, "ME_OSTR", oRsGetDflt.Fields.Item(0).Value))
                oDBs_Head.SetValue("U_MESRS", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)
                oDBs_Head.SetValue("Series", oDBs_Head.Offset, oRsGetDflt.Fields.Item(2).Value)
                oDBs_Head.SetValue("U_MESTS", oDBs_Head.Offset, "Open")
                oDBs_Head.SetValue("U_MEPSDT", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
                oDBs_Head.SetValue("U_MEDOCDT", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
            Else
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNumSeries(objForm, "ME_OSTR", Series))
                oDBs_Head.SetValue("U_MESRS", oDBs_Head.Offset, Series)
                oDBs_Head.SetValue("Series", oDBs_Head.Offset, Series)
                oDBs_Head.SetValue("U_MESTS", oDBs_Head.Offset, "Open")
                oDBs_Head.SetValue("U_MEPSDT", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
                oDBs_Head.SetValue("U_MEDOCDT", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
            End If
            Dim GetRequesterDtls As String = "Select T1.""empID"" From OUSR T0 Inner Join OHEM T1 On T0.""USERID""  = T1.""userId""  Where ""USER_CODE"" = '" & objMain.objCompany.UserName & "'"
            Dim oRsGetRequesterDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetRequesterDtls.DoQuery(GetRequesterDtls)
            If oRsGetRequesterDtls.RecordCount > 0 Then
                Dim GetEmployeeDtls As String = ""
                If Sql = "Sql" Then
                    GetEmployeeDtls = "Select T0.""firstName"" + ' ' + T0.""lastName"" AS ""Name""  FROM OHEM T0  Inner Join OUDP T1 ON T0.""dept"" = T1.""Code"" LEFT JOIN ""OBPL"" T2 ON T0.""BPLId""=T2.""BPLId"" Where T0.""empID""='" & oRsGetRequesterDtls.Fields.Item("empID").Value & "'"
                ElseIf Hana = "Hana" Then
                    GetEmployeeDtls = "Select T0.""firstName"" " & Concate & "' '" & Concate & " T0.""lastName"" AS ""Name"" FROM OHEM T0  Inner Join OUDP T1 ON T0.""dept"" = T1.""Code"" LEFT JOIN ""OBPL"" T2 ON T0.""BPLId""=T2.""BPLId"" Where T0.""empID""='" & oRsGetRequesterDtls.Fields.Item("empID").Value & "'"
                End If
                Dim oRsGetEmployeeDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetEmployeeDtls.DoQuery(GetEmployeeDtls)
                oDBs_Head.SetValue("U_MEREMPID", oDBs_Head.Offset, oRsGetRequesterDtls.Fields.Item("empID").Value)
                oDBs_Head.SetValue("U_MEREQBY", oDBs_Head.Offset, oRsGetEmployeeDtls.Fields.Item(0).Value)
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
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_STR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_STR_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@ME_STR_C1")

            objMatrix1 = objForm.Items.Item("3").Specific
            objMatrix2 = objForm.Items.Item("4").Specific

            Select Case MatrixUID
                Case "3"
                    objMatrix1.AddRow()
                    oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
                    oDBs_Details1.SetValue("U_MEBRCD", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_MEBRNM", oDBs_Details1.Offset, "")
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
            If pVal.MenuUID = "ME_STR" And pVal.BeforeAction = False Then
                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
                Try
                    objComboBox = objForm.Items.Item("51").Specific
                    objComboBox.Select("ST", SAPbouiCOM.BoSearchKey.psk_ByValue)
                Catch ex As Exception
                End Try
                objForm.Freeze(True)
                objMatrix1.CommonSetting.SetRowEditable(1, True)
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
                objForm.Freeze(False)
            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                Me.SetNewLine(objForm.UniqueID, "3")
            ElseIf pVal.MenuUID = "Close" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("ME_STR_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix1 = objForm.Items.Item("3").Specific
                Dim result As Integer = objMain.objApplication.MessageBox("Changing document is irreversible..Confirm... Do you want to continue?", 1, "Continue", "Cancel")
                If (result = 1) Then
                    Me.UpdateStatus(objForm.UniqueID)
                End If
            ElseIf pVal.MenuUID = "Delete Row" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("ME_STR_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
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
                    frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'TransferRequest' AND ""Type"" = 'C'")
                    If frs.RecordCount = 0 Then
                        objMain.objApplication.MessageBox("Report layout 'TransferRequest' not found. Please upload with same name and Code as parameter", 0, "OK")
                    Else
                        objMain.objApplication.ActivateMenuItem(frs.Fields.Item(0).Value.ToString())
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_STR")
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
                    frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'TransferRequest' AND ""Type"" = 'C'")
                    If frs.RecordCount = 0 Then
                        objMain.objApplication.MessageBox("Report layout 'TransferRequest' not found. Please upload with same name and Code as parameter", 0, "OK")
                    Else
                        objMain.objApplication.ActivateMenuItem(frs.Fields.Item(0).Value.ToString())
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_STR")
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

    Private Sub UpdateStatus(ByVal FormUID As String)
        Try
            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_STR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_STR_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@ME_STR_C1")

            objMatrix1 = objForm.Items.Item("3").Specific
            objMatrix2 = objForm.Items.Item("4").Specific

            objMain.sCmp = objMain.objCompany.GetCompanyService
            objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OSTR")
            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0))
            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
            objMain.oGeneralData.SetProperty("U_MESTS", "Close")
            objMain.oGeneralService.Update(objMain.oGeneralData)
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_STR"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_STR_C0"), "DocEntry", oDBs_Details1.GetValue("DocEntry", 0))
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_STR_C1"), "DocEntry", oDBs_Details1.GetValue("DocEntry", 0))
            objMatrix1.LoadFromDataSource()
            objMatrix2.LoadFromDataSource()
            objForm.Refresh()
            objMain.objApplication.StatusBar.SetText("DocStatus updated successfully...!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objForm.Freeze(False)
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
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_STR")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_STR_C0")
                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@ME_STR_C1")
                    objMatrix1 = objForm.Items.Item("3").Specific
                    objMatrix2 = objForm.Items.Item("4").Specific
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                                  (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        objForm.Items.Item("42").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        LostFocusFlag = True
                        If Me.Validation(objForm.UniqueID, "Y") = False Then
                            LostFocusFlag = False
                            BubbleEvent = False
                        End If
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        LostFocusFlag = False
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
                        Me.GetPDFFile("Transfer Request SD.rpt", "Transfer Request SD", "DocKey@")
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim LoadSameForm As String = "SELECT MAX(""DocNum"") FROM ""@ME_STR"""
                        Dim oRsLoadSameForm As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsLoadSameForm.DoQuery(LoadSameForm)
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objForm.Items.Item("18").Specific.Value = oRsLoadSameForm.Fields.Item(0).Value
                        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If

                    If pVal.ItemUID = "57" And pVal.BeforeAction = False Then
                        Me.LoadMatrix(objForm.UniqueID)
                    End If

                    If pVal.ItemUID = "58" And pVal.BeforeAction = False Then
                        ''SELECT *  FROM "TESTDB"."@ME_ACFG"  T0 WHERE T0."U_INTRW" ='N
                        Dim GetInTransitWhs As String = "SELECT *  FROM ""@ME_ACFG""  T0 WHERE T0.""U_INTRW"" ='N' "
                        Dim oRsGetInTransitWhs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetInTransitWhs.DoQuery(GetInTransitWhs)
                        If oRsGetInTransitWhs.RecordCount > 0 Then

                            Dim GetOpenQty As String = "Select ""U_MEOPNQTY"" From ""@ME_STR_C0"" Where ""DocEntry"" = '" & oDBs_Head.GetValue("DocEntry", 0) & "' And ""U_MEOPNQTY"" > 0 And ""U_MEITMCD"" <> '' "
                            Dim oRsGetOpenQty As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetOpenQty.DoQuery(GetOpenQty)
                            If oRsGetOpenQty.RecordCount > 0 Then
                                Me.CopyToGoodsIssue(objForm.UniqueID)
                            Else
                                objMain.objApplication.StatusBar.SetText("Atleast there should be a single open qty to proceed.. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If



                        Else
                            GetInTransitWhs = "Select ""U_INTRWHS"" From ""OWHS""  Where ""U_INTRWHS"" <> '' And ""WhsCode"" = '" & objForm.Items.Item("6").Specific.Value & "' "
                            oRsGetInTransitWhs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetInTransitWhs.DoQuery(GetInTransitWhs)
                            If oRsGetInTransitWhs.RecordCount > 0 Then
                                Dim GetOpenQty As String = "Select ""U_MEOPNQTY"" From ""@ME_STR_C0"" Where ""DocEntry"" = '" & oDBs_Head.GetValue("DocEntry", 0) & "' And ""U_MEOPNQTY"" > 0 And ""U_MEITMCD"" <> '' "
                                Dim oRsGetOpenQty As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetOpenQty.DoQuery(GetOpenQty)
                                If oRsGetOpenQty.RecordCount > 0 Then
                                    Me.CopyToInventoryTransfer(objForm.UniqueID, oRsGetInTransitWhs.Fields.Item("U_INTRWHS").Value)
                                Else
                                    objMain.objApplication.StatusBar.SetText("Atleast there should be a single open qty to proceed.. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Exit Sub
                                End If
                            Else
                                objMain.objApplication.StatusBar.SetText("Please configure InTransit Warehouse to from Wareouse in warehouse master...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If
                        End If

                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_STR")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_STR_C0")
                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@ME_STR_C1")
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
                    If oCFL.UniqueID = "CFL_BCD" And pVal.BeforeAction = True Then
                        If objForm.Items.Item("31").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Requester Name cannot be empty...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If (objForm.Items.Item("6").Specific.Value = "" Or objForm.Items.Item("8").Specific.Value = "") Then
                            objMain.objApplication.StatusBar.SetText("From WareHouse Or To WareHouse cannot be empty...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        Me.CFLBarCode(objForm.UniqueID, "CFL_BCD", objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value, objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                    End If
                    If oCFL.UniqueID = "CFL_ITMCD" And pVal.BeforeAction = True Then
                        If objForm.Items.Item("31").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Requester Name cannot be empty...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If (objForm.Items.Item("6").Specific.Value = "" Or objForm.Items.Item("8").Specific.Value = "") Then
                            objMain.objApplication.StatusBar.SetText("From WareHouse Or To WareHouse cannot be empty...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        If oCFL.UniqueID = "CFL_PRQ" Then
                            oDBs_Head.SetValue("U_MEPRN", oDBs_Head.Offset, oDT.GetValue("DocNum", 0))
                            oDBs_Head.SetValue("U_MEPRE", oDBs_Head.Offset, oDT.GetValue("DocEntry", 0))
                        End If
                        If oCFL.UniqueID = "CFL_EMPID" Then
                            Dim GetEmployeeDtls As String = ""
                            GetEmployeeDtls = "Select T0.""firstName"" " & Concate & "' '" & Concate & " T0.""lastName"" AS ""Name"", T1.""Name"" AS ""Dept""," & IsNull & "(T2.""BPLName"",'') AS ""Branch""," & IsNull & "(T0.""CostCenter"",'') AS ""CostCenter"" FROM OHEM T0  Inner Join OUDP T1 ON T0.""dept"" = T1.""Code"" LEFT JOIN ""OBPL"" T2 ON T0.""BPLId""=T2.""BPLId"" Where T0.""empID""='" & oDT.GetValue("empID", 0) & "'"
                            Dim oRsGetEmployeeDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetEmployeeDtls.DoQuery(GetEmployeeDtls)
                            oDBs_Head.SetValue("U_MEREMPID", oDBs_Head.Offset, oDT.GetValue("empID", 0))
                            oDBs_Head.SetValue("U_MEREQBY", oDBs_Head.Offset, oRsGetEmployeeDtls.Fields.Item(0).Value)
                        End If
                        'If oCFL.UniqueID = "CFL_FWHSCD" Then
                        '    If oDT.GetValue("WhsCode", 0) = objForm.Items.Item("8").Specific.Value Then
                        '        objMain.objApplication.StatusBar.SetText("From Branch and To Branch cannot be same", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '        BubbleEvent = False
                        '        Exit Sub
                        '    End If
                        '    oDBs_Head.SetValue("U_MEFWHSCD", oDBs_Head.Offset, oDT.GetValue("WhsCode", 0))
                        '    oDBs_Head.SetValue("U_MEFWHSNM", oDBs_Head.Offset, oDT.GetValue("WhsName", 0))
                        '    oDBs_Head.SetValue("U_MEFBPID", oDBs_Head.Offset, oDT.GetValue("BPLid", 0))
                        'End If
                        If oCFL.UniqueID = "CFL_TWHSCD" Then
                            If oDT.GetValue("WhsCode", 0) = objForm.Items.Item("6").Specific.Value Then
                                objMain.objApplication.StatusBar.SetText("From Branch and To Branch cannot be same", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            oDBs_Head.SetValue("U_METWHSCD", oDBs_Head.Offset, oDT.GetValue("WhsCode", 0))
                            oDBs_Head.SetValue("U_METWHSNM", oDBs_Head.Offset, oDT.GetValue("WhsName", 0))
                            oDBs_Head.SetValue("U_METBPID", oDBs_Head.Offset, oDT.GetValue("BPLid", 0))
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
                            oDBs_Details1.SetValue("U_MEBRNM", oDBs_Details1.Offset, oDT.GetValue("BcdName", 0))
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
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_STR")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_STR_C0")
                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@ME_STR_C1")
                    objMatrix1 = objForm.Items.Item("3").Specific
                    objMatrix2 = objForm.Items.Item("4").Specific

                    If pVal.ItemUID = "3" And pVal.ColUID = "V_9" And pVal.BeforeAction = False And LostFocusFlag = False And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, pVal.Row)
                        oDBs_Details1.SetValue("U_MEBRCD", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Details1.SetValue("U_MEBRNM", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
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

                    'If pVal.ItemUID = "14" And pVal.BeforeAction = True Then
                    '    objComboBox2 = objForm.Items.Item("14").Specific

                    'End If
                    If pVal.ItemUID = "14" And pVal.BeforeAction = False Then
                        objComboBox = objForm.Items.Item("14").Specific
                        Dim objComboBox1 As SAPbouiCOM.ComboBox = objForm.Items.Item("16").Specific
                        ' Dim chk As String = objForm.Items.Item("14").Specific
                        If objComboBox.Value = objComboBox1.Value Then
                            objComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                            objMain.objApplication.StatusBar.SetText("From Branch and To Branch cannot be same", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        Else
                            oDBs_Head.SetValue("U_MEFWHSCD", oDBs_Head.Offset, "")
                            oDBs_Head.SetValue("U_MEFWHSNM", oDBs_Head.Offset, "")
                        End If
                        Dim QUERY As String = "SELECT  T0.""WhsCode"", T0.""WhsCode"" FROM ""OWHS""  T0 WHERE T0.""BPLid""='" & objComboBox.Value & "'"
                        objComboBox = objForm.Items.Item("6").Specific
                        objMain.objUtilities.ComboBoxLoadValues(objComboBox, QUERY) '" Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join " & _
                        '" OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""  ")
                        objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                        objForm.Items.Item("6").DisplayDesc = True

                        'If objComboBox.Value <> objComboBox2.Value Then
                        '    objComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                        '    oDBs_Head.SetValue("U_MEFWHSCD", oDBs_Head.Offset, "")
                        '    oDBs_Head.SetValue("U_MEFWHSNM", oDBs_Head.Offset, "")
                        '    BubbleEvent = False`
                        '    Exit Sub
                        'End If

                        '  Me.CFLWhsCodeFilter(objForm.UniqueID, "CFL_FWHSCD", objComboBox.Value)
                    End If
                    If pVal.ItemUID = "6" And pVal.BeforeAction = False Then
                        objComboBox = objForm.Items.Item("6").Specific
                        If objComboBox.Value <> "" Then
                            Dim QUERY As String = "SELECT DISTINCT  T0.""BPLid"",T0.""WhsCode"", T0.""WhsName"" FROM ""OWHS""  T0 WHERE T0.""WhsCode""='" & objComboBox.Value & "'"

                            Dim oRsChkBarCodeUOM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsChkBarCodeUOM.DoQuery(QUERY)

                            oDBs_Head.SetValue("U_MEFWHSNM", oDBs_Head.Offset, oRsChkBarCodeUOM.Fields.Item("WhsName").Value)
                            oDBs_Head.SetValue("U_MEFBPID", oDBs_Head.Offset, oRsChkBarCodeUOM.Fields.Item("BPLid").Value)

                        End If

                    End If

                    If pVal.ItemUID = "16" And pVal.BeforeAction = False Then
                        objComboBox = objForm.Items.Item("16").Specific
                        Dim objComboBox1 As SAPbouiCOM.ComboBox = objForm.Items.Item("14").Specific
                        If objComboBox.Value = objComboBox1.Value Then
                            objComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

                            objMain.objApplication.StatusBar.SetText("From Branch and To Branch cannot be same", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        Else
                            oDBs_Head.SetValue("U_METWHSCD", oDBs_Head.Offset, "")
                            oDBs_Head.SetValue("U_METWHSNM", oDBs_Head.Offset, "")
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

    Public Sub CopyToGoodsIssue(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMain.objApplication.ActivateMenuItem("3079")
            objITRForm = objMain.objApplication.Forms.GetForm("720", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objITRForm.Freeze(True)
            Dim GetInTransitWhs1 As String = "SELECT *  FROM ""@ME_ACFG""  T0 WHERE T0.""U_INTRW"" ='N' "
            Dim oRsGetInTransitWhs1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetInTransitWhs1.DoQuery(GetInTransitWhs1)
            If oRsGetInTransitWhs1.RecordCount > 0 Then
                Try
                    objITRFormUDF = objMain.objApplication.Forms.GetForm("-720", objMain.objApplication.Forms.ActiveForm.TypeCount)
                    objITRFormUDF.Items.Item("U_TRTYPE").Specific.value = "STR"

                Catch ex As Exception
                    objMain.objApplication.ActivateMenuItem("6913")
                    objITRFormUDF = objMain.objApplication.Forms.GetForm("-720", objMain.objApplication.Forms.ActiveForm.TypeCount)
                    objITRFormUDF.Items.Item("U_TRTYPE").Specific.value = "STR"

                End Try
                objITRMatrix = objITRForm.Items.Item("13").Specific
                'Dim combo1 As String = objITRForm.Items.Item("51").Specific.Value
                ' objComboBox = objITRForm.Items.Item("U_TRTYPE").Specific
                'objComboBox.Select("STR", SAPbouiCOM.BoSearchKey.psk_ByDescription)

            Else
                Try
                    objITRFormUDF = objMain.objApplication.Forms.GetForm("-720", objMain.objApplication.Forms.ActiveForm.TypeCount)
                    objITRFormUDF.Items.Item("U_TRTYPE").Specific.Value = "STR"
                Catch ex As Exception
                    objMain.objApplication.ActivateMenuItem("6913")
                End Try
                objITRMatrix = objITRForm.Items.Item("13").Specific
                objITRFormUDF.Items.Item("U_TRTYPE").Specific.Value = "STR"
            End If
            ''oDBs_Head = objForm.DataSources.DBDataSources.Item("OIGE")
            ''oDBs_Head.SetValue("BPLName", 0, objForm.Items.Item("6").Specific.Value)
            'objITRForm.Items.Item("1320000092").Enabled = True
            'objITRForm.Items.Item("1320000092").Specific.Value = objForm.Items.Item("6").Specific.Value
            'objITRForm.Items.Item("11").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            'objITRForm.Items.Item("1320000092").Enabled = False

            'Dim GetSeries As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = '60' And ""Locked"" = 'N'  And ""BPLId"" =(SELECT T0.""BPLId"" FROM OBPL T0 WHERE T0.""BPLName"" ='" & objForm.Items.Item("6").Specific.Value & "')  And  ""Indicator"" = YEAR(CURRENT_DATE) "
            'Dim oRsGetSeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRsGetSeries.DoQuery(GetSeries)
            'If oRsGetSeries.RecordCount > 0 Then
            '    objComboBox = objITRForm.Items.Item("30").Specific
            '    objComboBox.Select(oRsGetSeries.Fields.Item("SeriesName").Value, SAPbouiCOM.BoSearchKey.psk_ByDescription)
            '    ''oInventoryGenExit.Series = oRsGetSeries.Fields.Item("Series").Value
            '    ''objITRFormUDF.Items.Item("").Specific.Value = oRsGetSeries.Fields.Item("Series").Value

            'End If

            objITRFormUDF.Items.Item("U_INVTRREQ").Specific.Value = objForm.Items.Item("18").Specific.Value
            objITRFormUDF.Items.Item("U_INVTRREN").Specific.Value = objForm.Items.Item("26").Specific.Value.Trim

            '---Madhu
            Dim BranchCombo1 As SAPbouiCOM.ComboBox = objForm.Items.Item("16").Specific

            If BranchCombo1.Selected IsNot Nothing Then
                objITRFormUDF.Items.Item("U_TOWN").Specific.Value = BranchCombo1.Selected.Description
            End If

            objITRFormUDF.Items.Item("U_TOBNCH").Specific.Value = objForm.Items.Item("10").Specific.Value.Trim()



            '' objITRForm.Items.Item("18").Specific.Value = objForm.Items.Item("6").Specific.Value.Trim
            ''objITRForm.Items.Item("1470000101").Specific.Value = InTransitWarehouse
            Try
                ''objITRForm.Items.Item("41").Specific.Value = objForm.Items.Item("42").Specific.Value.Trim
            Catch ex As Exception
            End Try
            Dim oColumn As SAPbouiCOM.Column
            If objMatrix1.VisualRowCount > 0 Then
                For i As Integer = 1 To objMatrix1.VisualRowCount
                    If objMatrix1.Columns.Item("V_1").Cells.Item(i).Specific.Value <> "" And CDbl(objMatrix1.Columns.Item("V_11").Cells.Item(i).Specific.Value) > 0 Then
                        'objITRMatrix.Columns.Item("1").Cells.Item(objITRMatrix.VisualRowCount).Specific.Value = objMatrix1.Columns.Item("V_1").Cells.Item(i).Specific.Value
                        'objITRMatrix.Columns.Item("26").Cells.Item(objITRMatrix.VisualRowCount).Specific.Value = objMatrix1.Columns.Item("V_0").Cells.Item(i).Specific.Value
                        Dim Barcode As String = objMatrix1.Columns.Item("V_0").Cells.Item(i).Specific.Value.ToString().Trim()

                        If String.IsNullOrEmpty(Barcode) Then
                            Barcode = objMatrix1.Columns.Item("V_4").Cells.Item(i).Specific.Value.ToString().Trim()
                        End If

                        objITRMatrix.Columns.Item("26").Cells.Item(objITRMatrix.VisualRowCount).Specific.Value = Barcode
                        Try
                            ''oColumn = objITRMatrix.Columns.Item("1001")
                            ''oColumn.Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Select("No", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            objITRMatrix.Columns.Item("1470001043").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = objMatrix1.Columns.Item("V_3").Cells.Item(i).Specific.Value
                        Catch ex As Exception
                        End Try
                        'Try
                        '    objITRMatrix.Columns.Item("64").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = objMatrix1.Columns.Item("V_0").Cells.Item(i).Specific.Value
                        'Catch ex As Exception
                        'End Try
                        objITRMatrix.Columns.Item("U_INVTRRLN").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = objMatrix1.Columns.Item("V_-1").Cells.Item(i).Specific.Value
                        objITRMatrix.Columns.Item("U_INVTRREQ").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = objForm.Items.Item("18").Specific.Value
                        objITRMatrix.Columns.Item("U_INVTRREN").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = objForm.Items.Item("26").Specific.Value.Trim
                        objITRMatrix.Columns.Item("U_METSTQTY").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = CDbl(objMatrix1.Columns.Item("V_9").Cells.Item(i).Specific.Value)
                        objITRMatrix.Columns.Item("U_MEREQQTY").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = CDbl(objMatrix1.Columns.Item("V_11").Cells.Item(i).Specific.Value)
                        Try
                            objITRMatrix.Columns.Item("U_Packing").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = objMatrix1.Columns.Item("V_14").Cells.Item(i).Specific.Value
                        Catch ex As Exception
                        End Try

                        Try
                            objITRMatrix.Columns.Item("U_Addbarcode").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = CDbl(objMatrix1.Columns.Item("V_4").Cells.Item(i).Specific.Value)
                        Catch ex As Exception
                        End Try

                        'Try
                        '    Dim Barcode As String = objMatrix1.Columns.Item("V_0").Cells.Item(i).Specific.Value.ToString().Trim()

                        '    objITRMatrix.Columns.Item("26").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = Barcode

                        'Catch ex As Exception
                        'End Try

                        Dim BranchCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("16").Specific
                        objITRMatrix.Columns.Item("U_METBPID").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = BranchCombo.Selected.Description
                        objITRMatrix.Columns.Item("U_METWHSCD").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = objForm.Items.Item("8").Specific.Value.Trim
                        objITRMatrix.Columns.Item("U_METWHSNM").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = objForm.Items.Item("10").Specific.Value.Trim

                        objITRMatrix.Columns.Item("9").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = CDbl(objMatrix1.Columns.Item("V_11").Cells.Item(i).Specific.Value)
                        objITRMatrix.Columns.Item("15").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = objForm.Items.Item("6").Specific.Value.Trim



                        'GetInTransitWhs1 = "SELECT T0.""U_MEEXACCT"", T0.""U_MEENTACT"" FROM OWHS T0 WHERE T0.""WhsCode"" ='" & objITRMatrix.Columns.Item("15").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value & "'"
                        'oRsGetInTransitWhs1 = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'oRsGetInTransitWhs1.DoQuery(GetInTransitWhs1)
                        'If oRsGetInTransitWhs1.RecordCount > 0 Then
                        '    objITRMatrix.Columns.Item("59").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = oRsGetInTransitWhs1.Fields.Item("U_MEEXACCT").Value
                        'End If


                    End If
                Next
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            objITRForm.Freeze(False)
        End Try
    End Sub

    Public Sub CopyToInventoryTransfer(ByVal FormUID As String, ByVal InTransitWarehouse As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMain.objApplication.ActivateMenuItem("3080")
            objITRForm = objMain.objApplication.Forms.GetForm("940", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objITRMatrix = objITRForm.Items.Item("23").Specific
            objITRForm.Items.Item("U_TRTYPE").Specific.Value = "STR"
            objITRForm.Items.Item("U_INVTRREQ").Specific.Value = objForm.Items.Item("18").Specific.Value
            objITRForm.Items.Item("U_INVTRREN").Specific.Value = objForm.Items.Item("26").Specific.Value.Trim
            objITRForm.Items.Item("18").Specific.Value = objForm.Items.Item("6").Specific.Value.Trim
            objITRForm.Items.Item("1470000101").Specific.Value = InTransitWarehouse
            Try
                objITRForm.Items.Item("41").Specific.Value = objForm.Items.Item("42").Specific.Value.Trim
            Catch ex As Exception
            End Try
            Dim oColumn As SAPbouiCOM.Column
            If objMatrix1.VisualRowCount > 0 Then
                For i As Integer = 1 To objMatrix1.VisualRowCount
                    If objMatrix1.Columns.Item("V_1").Cells.Item(i).Specific.Value <> "" And CDbl(objMatrix1.Columns.Item("V_11").Cells.Item(i).Specific.Value) > 0 Then
                        objITRMatrix.Columns.Item("1").Cells.Item(objITRMatrix.VisualRowCount).Specific.Value = objMatrix1.Columns.Item("V_1").Cells.Item(i).Specific.Value
                        Try
                            oColumn = objITRMatrix.Columns.Item("1001")
                            oColumn.Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Select("No", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            objITRMatrix.Columns.Item("1470001043").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = objMatrix1.Columns.Item("V_3").Cells.Item(i).Specific.Value
                        Catch ex As Exception
                        End Try
                        'Try
                        '    objITRMatrix.Columns.Item("64").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = objMatrix1.Columns.Item("V_0").Cells.Item(i).Specific.Value
                        'Catch ex As Exception
                        'End Try
                        objITRMatrix.Columns.Item("U_INVTRRLN").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = objMatrix1.Columns.Item("V_-1").Cells.Item(i).Specific.Value
                        objITRMatrix.Columns.Item("U_INVTRREQ").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = objForm.Items.Item("18").Specific.Value
                        objITRMatrix.Columns.Item("U_INVTRREN").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = objForm.Items.Item("26").Specific.Value.Trim
                        objITRMatrix.Columns.Item("U_METSTQTY").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = CDbl(objMatrix1.Columns.Item("V_9").Cells.Item(i).Specific.Value)
                        objITRMatrix.Columns.Item("U_MEREQQTY").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = CDbl(objMatrix1.Columns.Item("V_11").Cells.Item(i).Specific.Value)
                        Try
                            objITRMatrix.Columns.Item("U_Packing").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = objMatrix1.Columns.Item("V_14").Cells.Item(i).Specific.Value
                        Catch ex As Exception
                        End Try

                        Try
                            objITRMatrix.Columns.Item("U_Addbarcode").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = CDbl(objMatrix1.Columns.Item("V_4").Cells.Item(i).Specific.Value)
                        Catch ex As Exception
                        End Try

                        Dim BranchCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("16").Specific
                        objITRMatrix.Columns.Item("U_METBPID").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = BranchCombo.Selected.Description
                        objITRMatrix.Columns.Item("U_METWHSCD").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = objForm.Items.Item("8").Specific.Value.Trim
                        objITRMatrix.Columns.Item("U_METWHSNM").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = objForm.Items.Item("10").Specific.Value.Trim

                        objITRMatrix.Columns.Item("10").Cells.Item(objITRMatrix.VisualRowCount - 1).Specific.Value = CDbl(objMatrix1.Columns.Item("V_11").Cells.Item(i).Specific.Value)


                    End If
                Next
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String, ByVal Flag As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_STR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_STR_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@ME_STR_C1")
            objMatrix1 = objForm.Items.Item("3").Specific
            objMatrix2 = objForm.Items.Item("4").Specific

            If objForm.Items.Item("31").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Requester Name  Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("6").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("From Warehouse Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("8").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("To Warehouse  Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
                    'Commented by  Ramya on 28-09-2026
                    'For i As Integer = 1 To objMatrix1.VisualRowCount - 1
                    '    If objMatrix1.Columns.Item("V_9").Cells.Item(i).Specific.Value = "" Then
                    '        objMatrix1.Columns.Item("V_9").Cells.Item(i).Specific.Value = 0
                    '    End If
                    '    If objMatrix1.Columns.Item("V_9").Cells.Item(i).Specific.Value <= 0 Then
                    '        objMain.objApplication.StatusBar.SetText("Require Quantity should not be zero or less than zero.. in Line : " & i & ", Delete the line before proceeding", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        Return False
                    '    End If
                    'Next

                    For i As Integer = 1 To objMatrix1.VisualRowCount - 1

                        Dim RequiredQuantity As Double = 0
                        Dim FromWhreHouseInStockQuantity As Double = 0

                        If objMatrix1.Columns.Item("V_9").Cells.Item(i).Specific.Value <> "" Then
                            RequiredQuantity = CDbl(objMatrix1.Columns.Item("V_9").Cells.Item(i).Specific.Value)
                        End If



                        If objMatrix1.Columns.Item("V_5").Cells.Item(i).Specific.Value <> "" Then
                            FromWhreHouseInStockQuantity = CDbl(objMatrix1.Columns.Item("V_5").Cells.Item(i).Specific.Value)
                        End If

                        ' STOCK ZERO VALIDATION
                        'If FromWhreHouseInStockQuantity <= 0 Then

                        '    objMain.objApplication.StatusBar.SetText("From Warehouse InStock Quantity is Zero in Line : " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    Return False

                        'End If

                        ' REQUIRED QTY GREATER THAN STOCK
                        'Commented by  Ramya on 28-09-2026
                        'If RequiredQuantity > FromWhreHouseInStockQuantity Then

                        '    'objMain.objApplication.StatusBar.SetText("Required Quantity cannot be greater than From Warehouse InStock Quantity in Line : " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    objMain.objApplication.StatusBar.SetText("Required Quantity exceeds Available Stock Quantity : " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    Return False

                        'End If

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
            ' Exit Try
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
                    oDBs_Details1.SetValue("U_MEBRNM", oDBs_Details1.Offset, "")
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

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_STR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_STR_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@ME_STR_C1")

            objMatrix1 = objForm.Items.Item("3").Specific
            objMatrix2 = objForm.Items.Item("4").Specific

            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_STR"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_STR_C0"), "DocEntry", oDBs_Details1.GetValue("DocEntry", 0))
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@ME_STR_C1"), "DocEntry", oDBs_Details1.GetValue("DocEntry", 0))

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
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_STR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_STR_C0")
            objMatrix1 = objForm.Items.Item("3").Specific
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
                objForm.Items.Item("58").Enabled = False
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE

            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("ME_STR_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
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
                    objForm.EnableMenu("1292", True)
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

                If f.TypeEx = "ME_STR_Form" OrElse f.TypeEx = "ME_STR_Form" Then
                    Try
                        'Dim txtObj As SAPbouiCOM.EditText = CType(f.Items.Item("DocEntry").Specific, SAPbouiCOM.EditText)
                        'Dim docVal As String = If(txtObj IsNot Nothing, txtObj.Value.Trim(), "")
                        Dim docVal As String
                        If String.IsNullOrWhiteSpace(docVal) Then
                            ' fallback: try header datasource DocNum
                            Try
                                Dim ds As SAPbouiCOM.DBDataSource = f.DataSources.DBDataSources.Item("@ME_STR_Form")
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


    ' 'layout printing PCR code enabling
    Public Sub CreateLayoutFolder_StockTransferRequest()
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
                objMain.objApplication.StatusBar.SetText("CreateLayoutFolder_StockTransferRequest aborted: not connected to a company.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return
            End If

            Dim oReportTypeService As SAPbobsCOM.ReportTypesService
            oReportTypeService = companyServiceSource.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService)

            Dim oList = oReportTypeService.GetReportTypeList()

            For i As Integer = 0 To oList.Count - 1
                If oList.Item(i).AddonFormType = "ME_STR_Form" Then
                    Exit Sub   ' Already exists
                End If
            Next

            Dim oReportType As SAPbobsCOM.ReportType
            oReportType = oReportTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportType)

            oReportType.TypeName = "ME_STR_Form Print"
            oReportType.AddonName = "Payroll Addon"
            oReportType.AddonFormType = "ME_STR_Form"
            oReportType.MenuID = "ME_STR_Form_MENU"

            oReportTypeService.AddReportType(oReportType)

        Catch ex As Exception
            Try
                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Catch
            End Try
        End Try
    End Sub

End Class
