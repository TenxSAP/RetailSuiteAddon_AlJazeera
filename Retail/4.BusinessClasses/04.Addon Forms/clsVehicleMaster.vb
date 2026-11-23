Public Class clsVehicleMaster

#Region "       Declaration             "
    Public objForm, objITRForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details1, oDBs_Details2 As SAPbouiCOM.DBDataSource
    Dim objMatrix1, objMatrix2, objITRMatrix As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim LostFocusFlag As Boolean = False
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("VehicleMaster.xml", "ME_VM_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_VM_Form_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)           
            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_VM_Form")

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

            
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "", Optional ByVal Series As Integer = 0)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_VM_Form")
            

            objComboBox = objForm.Items.Item("19").Specific
            Dim GetSeriesQty As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'ME_OSTR' And ""Locked"" = 'N'"
            objMain.objUtilities.ComboBoxLoadValuesSeries(objForm.Items.Item("19").Specific, GetSeriesQty)
            objForm.Items.Item("19").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            Dim GetDflt As String = "Select ""DfltSeries"" From ONNM Where ""ObjectCode"" = 'ME_OSTR'"
            Dim oRsGetDflt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDflt.DoQuery(GetDflt)
            If Flag = "" Then
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNumSeries(objForm, "ME_OSTR", oRsGetDflt.Fields.Item(0).Value))
                oDBs_Head.SetValue("U_MESRS", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)
                oDBs_Head.SetValue("Series", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)
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
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_VM_Form")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@ME_VM_Form_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@ME_VM_Form_C1")

            objMatrix1 = objForm.Items.Item("3").Specific
            objMatrix2 = objForm.Items.Item("4").Specific

            Select Case MatrixUID
                Case "3"
                    objMatrix1.AddRow()
                    oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
                    oDBs_Details1.SetValue("U_MEBRCD", oDBs_Details1.Offset, "")
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
            If pVal.MenuUID = "ME_VHMSTR" And pVal.BeforeAction = False Then
                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
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
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_VM_Form")
          


            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


End Class
