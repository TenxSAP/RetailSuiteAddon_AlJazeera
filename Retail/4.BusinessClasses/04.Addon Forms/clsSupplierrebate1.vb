Imports System.Security.Cryptography

Public Class clsSupplierrebate1

#Region "    Declaration   "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details, oDBs_Details1, oDBs_Details2, oDS As SAPbouiCOM.DBDataSource
    Dim objMatrix, objMatrix1, objMatrix2 As SAPbouiCOM.Matrix
    Dim oEditText As SAPbouiCOM.EditText
    Dim oDt As SAPbouiCOM.DataTable
    Dim sSupplierCode, sSupplierName, sBaseNo, sBaseEntry As String
    Dim objComboBox, objComboBox1 As SAPbouiCOM.ComboBox
    Dim sDoNum As String
    Public strCardCode As String = ""
#End Region

    Sub CreateForm(ByVal FormUID As String, ByVal CardCode As String, ByVal CardName As String, ByVal BaseEntry As String, ByVal BaseNo As String, ByVal IfSRegExist As String, Optional SRebDocNum As String = "")



        Try
            strCardCode = CardCode

            'objMain.objUtilities.LoadForm("SupplierRebateNew1.xml", "S_Rebate", ResourceType.Embeded)
            objMain.objUtilities.LoadForm("TESTRebate.xml", "S_Rebate", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("S_Rebate", objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm = objMain.objApplication.Forms.GetForm("S_Rebate", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.EnableMenu("774", True)
            objForm.EnableMenu("1292", True)

            If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = False Then
                Retail.MainCls.ohtLookUpForm.Add(objForm.UniqueID, FormUID)
            End If
            objForm.Items.Item("Status").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Status").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Status").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Status").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)


            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)


            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)


            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)


            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("Submit").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Submit").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("Submit").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("Submit").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Dim BpCode As String = CardCode
            Dim ParentDocEntry As String = BaseEntry

            '''Loading brand details-----Madhu
            Dim rs1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'Dim sGetTyp1e As String = "SELECT dISTINCT T0.""Code"", T0.""Name"" FROM ""@TNX_BRAND"" T0 UNION ALL Select 'All' AS ""ItmsGrpCod"", 'All' AS ""ItmsGrpNam"" From DUMMY"

            Dim sGetTyp1e As String = "SELECT DISTINCT IFNULL(T0.""U_BRAND"",'') AS ""Code"", " & "IFNULL(T0.""U_BRAND"",'') AS ""Name"" " & "FROM OITM T0 " &
                          "WHERE T0.""CardCode"" = '" & CardCode & "' " & "AND IFNULL(T0.""U_BRAND"",'') <> '' " & "UNION ALL " & "SELECT 'All' AS ""Code"", 'All' AS ""Name"" FROM DUMMY"
            rs1.DoQuery(sGetTyp1e)
            Dim objComboBox3 As SAPbouiCOM.Column = objMatrix.Columns.Item("Brand")
            Try
                objMain.objUtilities.MatrixComboBoxValues(objComboBox3, sGetTyp1e)

            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText(ex.Message)

            End Try
            '


            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C0")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C1")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C2")

            ' objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE

            objMatrix = objForm.Items.Item("Mtx").Specific
            objMatrix1 = objForm.Items.Item("Mtx1").Specific
            objMatrix2 = objForm.Items.Item("Item_10").Specific


            objMatrix2.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

            objComboBox = objForm.Items.Item("8").Specific
            Dim GetSeriesQty As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'SBO_SUPREBATEUDO' And ""Locked"" = 'N'"
            objMain.objUtilities.ComboBoxLoadValuesSeries(objForm.Items.Item("8").Specific, GetSeriesQty)
            objForm.Items.Item("8").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            If IfSRegExist = "Y" Then

                objForm.Freeze(True)
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                objForm.Items.Item("9").Specific.value = SRebDocNum
                objForm.Items.Item("4").Specific.value = CardCode
                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                If objMatrix2.VisualRowCount = 0 Then

                    objMatrix2.Clear()
                    oDBs_Details2.Clear()
                    objMatrix2.FlushToDataSource()

                    Me.SetNewLine2(objForm.UniqueID)

                End If

                objForm.Freeze(False)
                Exit Sub
            End If



            If IfSRegExist = "N" Then
                sSupplierCode = CardCode
                sSupplierName = CardName

                sBaseNo = BaseNo
                sBaseEntry = BaseEntry

                objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                sDoNum = objMain.objUtilities.GetNextDocNum(objForm, "SBO_SUPREBATEUDO")
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, sDoNum)
                oDBs_Head.SetValue("U_BPCD", oDBs_Head.Offset, sSupplierCode)
                oDBs_Head.SetValue("U_BPNM", oDBs_Head.Offset, sSupplierName)

                oDBs_Head.SetValue("U_BASENT", oDBs_Head.Offset, sBaseEntry)
                oDBs_Head.SetValue("U_BASENO", oDBs_Head.Offset, sBaseNo)
                Me.AutoDocentryNumber()


                Dim GetDflt As String = "Select ""DfltSeries"" From ONNM Where ""ObjectCode"" = 'SBO_SUPREBATEUDO'"
                Dim oRsGetDflt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetDflt.DoQuery(GetDflt)

                oDBs_Head.SetValue("U_SERIES", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)
                oDBs_Head.SetValue("Series", oDBs_Head.Offset, oRsGetDflt.Fields.Item(0).Value)


                objMatrix.Clear()
                oDBs_Details.Clear()
                objMatrix.FlushToDataSource()
                Me.SetNewLine(objForm.UniqueID)

                objMatrix1.Clear()
                oDBs_Details1.Clear()
                objMatrix1.FlushToDataSource()
                Me.SetNewLine1(objForm.UniqueID)


                objMatrix2.Clear()
                oDBs_Details2.Clear()
                objMatrix2.FlushToDataSource()
                Me.SetNewLine2(objForm.UniqueID)

                Try
                    Dim sQury As String = " Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join " &
                                      " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""  "
                    Dim qQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    qQuery.DoQuery(sQury)
                    objMatrix1.Clear()
                    oDBs_Details1.Clear()
                    objMatrix1.FlushToDataSource()
                    If qQuery.RecordCount > 0 Then
                        Dim i As Integer = 0
                        While Not qQuery.EoF
                            oDBs_Details1.InsertRecord(i + 1)
                            oDBs_Details1.SetValue("LineId", i, (i + 1).ToString())
                            oDBs_Details1.SetValue("U_BRANCHID", i, qQuery.Fields.Item("BPLId").Value.ToString())
                            oDBs_Details1.SetValue("U_BRANCHNM", i, qQuery.Fields.Item("BPLName").Value.ToString())
                            i = i + 1
                            qQuery.MoveNext()
                        End While
                        oDBs_Details1.SetValue("LineId", i, (i + 1).ToString())
                        objMatrix1.LoadFromDataSource()
                        Try
                            If oDBs_Details1.GetValue("U_BRANCHID", i).ToString() = "" Then
                                objMain.objUtilities.DeleteRow(objMatrix1, oDBs_Details1)
                            End If
                        Catch ex As Exception

                        End Try
                        objMatrix1.LoadFromDataSource()
                    End If
                Catch ex As Exception

                End Try
            End If


            objForm.DataBrowser.BrowseBy = "9"
            objMatrix.AutoResizeColumns()
            objMatrix1.AutoResizeColumns()
            ''Disabled cloums
            'Me.DisableColumns()
            Me.CellMasking(objForm.UniqueID)
            Me.SetDefault(objForm.UniqueID)
            Me.LoadItemsToSupplierRebate(objForm.UniqueID, BpCode, ParentDocEntry)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "", Optional ByVal Series As Integer = 0)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C0")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C1")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C2")


            objMatrix = objForm.Items.Item("Mtx").Specific
            objMatrix1 = objForm.Items.Item("Mtx1").Specific
            objMatrix2 = objForm.Items.Item("Item_10").Specific
            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "SBO_SUPREBATEUDO", "Primary"))

            Try
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then


                    oDBs_Head.SetValue("U_Status", oDBs_Head.Offset, "Open")


                    Dim cbStatus As SAPbouiCOM.ComboBox =
            CType(objForm.Items.Item("Status").Specific, SAPbouiCOM.ComboBox)

                    If cbStatus IsNot Nothing Then
                        Try
                            cbStatus.Select("Open", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        Catch
                            cbStatus.Select("Open", SAPbouiCOM.BoSearchKey.psk_ByDescription)
                        End Try
                    End If

                End If
            Catch ex As Exception

            End Try

            objMatrix1.Clear()
            objMatrix2.Clear()
            objMatrix.Clear()

            oDBs_Details1.Clear()
            oDBs_Details2.Clear()
            oDBs_Details.Clear()

            objMatrix1.FlushToDataSource()
            objMatrix2.FlushToDataSource()
            objMatrix.FlushToDataSource()

            Me.SetNewLine1(objForm.UniqueID)
            Me.SetNewLine2(objForm.UniqueID)
            Me.SetNewLine(objForm.UniqueID)

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub CellMasking(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C0")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C1")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C2")

            objMatrix = objForm.Items.Item("Mtx").Specific
            objMatrix1 = objForm.Items.Item("Mtx1").Specific
            objMatrix2 = objForm.Items.Item("Item_10").Specific

            objForm.Items.Item("1000001").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("1000001").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "2", SAPbouiCOM.BoModeVisualBehavior.mvb_True)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C0")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C1")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C2")

            objMatrix = objForm.Items.Item("Mtx").Specific
            objMatrix1 = objForm.Items.Item("Mtx1").Specific
            objMatrix2 = objForm.Items.Item("Item_10").Specific

            objMatrix.AddRow()
            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
            oDBs_Details.SetValue("U_POAF", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_POAT", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_RBTPRC", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_RAmount", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Brand", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_GLact", oDBs_Details.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub SetNewLine1(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C0")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C1")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C2")

            objMatrix = objForm.Items.Item("Mtx").Specific
            objMatrix1 = objForm.Items.Item("Mtx1").Specific
            objMatrix2 = objForm.Items.Item("Item_10").Specific

            objMatrix1.AddRow()
            oDBs_Details1.SetValue("LineId", oDBs_Details.Offset, objMatrix1.VisualRowCount)
            oDBs_Details1.SetValue("U_CHECKBOX", oDBs_Details.Offset, "")
            oDBs_Details1.SetValue("U_BRANCHID", oDBs_Details.Offset, "")
            oDBs_Details1.SetValue("U_REMARKS", oDBs_Details.Offset, "")
            objMatrix1.SetLineData(objMatrix1.VisualRowCount)
            objMatrix1.AutoResizeColumns()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine2(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C1")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C2")

            objMatrix = objForm.Items.Item("Mtx").Specific
            objMatrix1 = objForm.Items.Item("Mtx1").Specific
            objMatrix2 = objForm.Items.Item("Item_10").Specific

            objMatrix2.AddRow()
            oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, objMatrix2.VisualRowCount)

            oDBs_Details2.SetValue("U_BRCD", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_ITMCD", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_ITMNM", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_QTY", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_UOM", oDBs_Details2.Offset, "")

            objMatrix2.SetLineData(objMatrix2.VisualRowCount)

            objMatrix2.AutoResizeColumns()


        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C0")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C1")
                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C2")

                    objMatrix = objForm.Items.Item("Mtx").Specific
                    objMatrix1 = objForm.Items.Item("Mtx1").Specific
                    objMatrix2 = objForm.Items.Item("Item_10").Specific

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim chkIfExist As String = "Select ""DocNum"" From ""@SBO_SUPREBATE"" Where ""DocEntry""=(Select MAX(""DocEntry"") From ""@SBO_SUPREBATE"")"
                        Dim oRschkIfExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRschkIfExist.DoQuery(chkIfExist)
                        If oRschkIfExist.RecordCount > 0 Then
                            objForm.Freeze(True)
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            objForm.Items.Item("9").Specific.value = oRschkIfExist.Fields.Item("DocNum").Value
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            objForm.Freeze(False)
                        End If
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        If Me.Validate1() = False Then

                            BubbleEvent = False
                        End If
                    End If

                    If pVal.ItemUID = "Submit" And pVal.BeforeAction = False And SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Try
                            Dim DocNum As String = objForm.Items.Item("AppID").Specific.value
                            Me.ApprovalTrigger("SuppReb", DocNum, "@SBO_SUPREBATE", "U_AppID", "U_Status", "SupplierRebate")
                        Catch ex As Exception
                        End Try


                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C0")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C1")
                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C2")

                    objMatrix = objForm.Items.Item("Mtx").Specific
                    objMatrix1 = objForm.Items.Item("Mtx1").Specific
                    objMatrix2 = objForm.Items.Item("Item_10").Specific

                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects

                    If pVal.BeforeAction = True Then


                        If oCFL.UniqueID = "CFL_2" Then
                            If CStr(objMatrix2.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific.Value).Trim = "" Or CStr(objMatrix2.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific.Value).Trim = "" Then
                                objMain.objApplication.StatusBar.SetText("ItemCode or Barcode cannot be left blank")
                                BubbleEvent = False

                            Else
                                Me.FilterUOM(objForm.UniqueID, oCFL.UniqueID, CStr(objMatrix2.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific.Value).Trim)
                            End If
                        End If

                    End If

                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)



                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)


                        If oCFL.UniqueID = "CFL_0" Then

                            oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, pVal.Row)
                            oDBs_Details2.SetValue("U_BRCD", oDBs_Details2.Offset, oDT.GetValue("CodeBars", 0))
                            oDBs_Details2.SetValue("U_ITMCD", oDBs_Details2.Offset, oDT.GetValue("ItemCode", 0))
                            oDBs_Details2.SetValue("U_ITMNM", oDBs_Details2.Offset, oDT.GetValue("ItemName", 0))
                            oDBs_Details2.SetValue("U_QTY", oDBs_Details2.Offset, "")

                            Dim GetUOM As String = "SELECT T0.""UomCode"" FROM OUOM T0  INNER JOIN OITM T1 ON T0.""UomEntry"" = T1.""PUoMEntry"" Where  T0.""UomEntry"" ='" & oDT.GetValue("PUoMEntry", 0) & "' "
                            Dim oRsGetUOM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetUOM.DoQuery(GetUOM)
                            If oRsGetUOM.RecordCount > 0 Then
                                oDBs_Details2.SetValue("U_UOM", oDBs_Details2.Offset, oRsGetUOM.Fields.Item(0).Value)

                            Else
                                oDBs_Details2.SetValue("U_UOM", oDBs_Details2.Offset, "")

                            End If


                            objMatrix2.SetLineData(pVal.Row)

                            If pVal.Row = objMatrix2.VisualRowCount Then
                                Me.SetNewLine2(objForm.UniqueID)
                            End If

                        End If


                        If oCFL.UniqueID = "CFL_1" Then

                            Try
                                oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, pVal.Row)
                                oDBs_Details2.SetValue("U_BRCD", oDBs_Details2.Offset, oDT.GetValue("BcdCode", 0))
                                oDBs_Details2.SetValue("U_ITMCD", oDBs_Details2.Offset, oDT.GetValue("ItemCode", 0))

                                Dim GetUOM As String = "SELECT T1.""UomCode"", T2.""ItemName"" FROM OBCD T0  INNER JOIN OUOM T1 " &
                                                         "ON T0.""UomEntry"" = T1.""UomEntry"" INNER JOIN OITM T2 ON " &
                                                        "T0.""ItemCode"" = T2.""ItemCode"" WHERE T0.""BcdCode"" ='" & oDT.GetValue("BcdCode", 0) & "'"


                                'Dim GetUOM As String = "SELECT T0.""UomCode"" ,T1.""ItemName""  FROM OUOM T0  INNER JOIN OITM T1 ON " &
                                '                    "T0.""UomEntry"" = T1.""PUoMEntry"" Where  T1.""ItemCode"" ='" & oDT.GetValue("ItemCode", 0) & "' "

                                Dim oRsGetUOM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetUOM.DoQuery(GetUOM)
                                If oRsGetUOM.RecordCount > 0 Then
                                    oDBs_Details2.SetValue("U_ITMNM", oDBs_Details2.Offset, CStr(oRsGetUOM.Fields.Item(1).Value).Trim)
                                    oDBs_Details2.SetValue("U_UOM", oDBs_Details2.Offset, oRsGetUOM.Fields.Item(0).Value)


                                Else
                                    oDBs_Details2.SetValue("U_ITMNM", oDBs_Details2.Offset, "")
                                    oDBs_Details2.SetValue("U_UOM", oDBs_Details2.Offset, "")

                                End If


                                oDBs_Details2.SetValue("U_QTY", oDBs_Details2.Offset, "")
                                objMatrix2.SetLineData(pVal.Row)

                                If pVal.Row = objMatrix2.VisualRowCount Then
                                    Me.SetNewLine2(objForm.UniqueID)
                                End If

                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If


                        If oCFL.UniqueID = "CFL_2" Then

                            oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, pVal.Row)
                            oDBs_Details2.SetValue("U_BRCD", oDBs_Details2.Offset, objMatrix2.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details2.SetValue("U_ITMCD", oDBs_Details2.Offset, objMatrix2.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details2.SetValue("U_ITMNM", oDBs_Details2.Offset, objMatrix2.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details2.SetValue("U_QTY", oDBs_Details2.Offset, objMatrix2.Columns.Item("Col_3").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details2.SetValue("U_UOM", oDBs_Details2.Offset, oDT.GetValue("UomCode", 0))
                            objMatrix2.SetLineData(pVal.Row)

                            If pVal.Row = objMatrix2.VisualRowCount Then
                                Me.SetNewLine2(objForm.UniqueID)
                            End If

                        End If

                        If oCFL.UniqueID = "CFL_3" Then
                            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row)
                            oDBs_Details.SetValue("U_Brand", oDBs_Details.Offset, objMatrix.Columns.Item("Brand").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_RBTPRC", oDBs_Details.Offset, objMatrix.Columns.Item("RBTPRC").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_RAmount", oDBs_Details.Offset, objMatrix.Columns.Item("RAmount").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_GLact", oDBs_Details.Offset, oDT.GetValue("AcctCode", 0))

                            objMatrix.SetLineData(pVal.Row)

                            If pVal.Row = objMatrix.VisualRowCount Then
                                'Me.SetNewLine(objForm.UniqueID)
                            End If

                        End If


                    End If


                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C0")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C1")
                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C2")

                    objMatrix = objForm.Items.Item("Mtx").Specific
                    objMatrix1 = objForm.Items.Item("Mtx1").Specific
                    objMatrix2 = objForm.Items.Item("Item_10").Specific

                    If pVal.ItemUID = "Mtx" And pVal.ColUID = "POAF" Then
                        If CDbl(objMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) > 0.0 Then
                            If pVal.Row = objMatrix.VisualRowCount Then
                                'Me.SetNewLine(objForm.UniqueID)
                            End If
                        End If
                    End If
                    If pVal.ItemUID = "Mtx" And pVal.ColUID = "RBTPRC" Then
                        If pVal.Row = 1 Then
                            If objMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value <> 0.0 Then
                                If pVal.Row = 1 Then
                                    objMatrix.Columns.Item("RAmount").Editable = False
                                End If
                            End If
                        End If
                    End If
                    If pVal.ItemUID = "Mtx" And pVal.ColUID = "RAmount" Then
                        If pVal.Row = 1 Then
                            If objMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value <> 0.0 Then
                                objMatrix.Columns.Item("RBTPRC").Editable = False
                            End If
                        End If
                    End If
                    If objForm.Items.Item("15").Specific.Value <> 0.0 Then
                        objForm.Items.Item("ITAmt").Enabled = False
                    Else
                        objForm.Items.Item("15").Enabled = False
                    End If

              '----madhu

                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                    If pVal.ItemUID = "Mtx" AndAlso pVal.ColUID = "Brand" AndAlso pVal.BeforeAction = False Then

                        Try

                            Dim objForm As SAPbouiCOM.Form =
                                objMain.objApplication.Forms.Item(FormUID)

                            Dim oMatrix As SAPbouiCOM.Matrix =
                                objForm.Items.Item("Mtx").Specific

                            '================ GET SELECTED BRAND =================
                            Dim strBrand As String = ""

                            Try

                                Dim oCombo As SAPbouiCOM.ComboBox =
                                    oMatrix.Columns.Item("Brand").Cells.Item(pVal.Row).Specific

                                strBrand = oCombo.Selected.Value.Trim()

                            Catch
                                strBrand = ""
                            End Try

                            'IF BRAND NOT SELECTED EXIT
                            If strBrand = "" Then Exit Sub

                            '================ GET GL ACCOUNT =================
                            Dim rsGL As SAPbobsCOM.Recordset =
                                objMain.objCompany.GetBusinessObject(
                                SAPbobsCOM.BoObjectTypes.BoRecordset)

                            Dim strGLAccount As String = ""

                            rsGL.DoQuery(
                                "SELECT ""U_SuppReb"" FROM ""@ME_ACFG"" ")

                            If rsGL.RecordCount > 0 Then

                                strGLAccount =
                                    rsGL.Fields.Item("U_SuppReb").Value.ToString().Trim()

                            End If

                            '================ SET GL ACCOUNT =================
                            If strGLAccount <> "" Then

                                CType(oMatrix.Columns.Item("GLact").Cells.Item(pVal.Row).Specific,
                                      SAPbouiCOM.EditText).Value = strGLAccount

                            End If

                        Catch ex As Exception

                            objMain.objApplication.StatusBar.SetText(
                                ex.Message,
                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        End Try

                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C0")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C1")
                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C2")

                    objMatrix = objForm.Items.Item("Mtx").Specific
                    objMatrix1 = objForm.Items.Item("Mtx1").Specific
                    objMatrix2 = objForm.Items.Item("Item_10").Specific

                    If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = True And pVal.BeforeAction = False Then
                        Retail.MainCls.ohtLookUpForm.Remove(objForm.UniqueID)
                    End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                Me.SetNewLine(objForm.UniqueID)

            ElseIf pVal.MenuUID = "774" And pVal.BeforeAction = False Then

                objMatrix = objForm.Items.Item("Mtx").Specific

                Dim row As Integer = objMatrix.VisualRowCount

                If objMatrix.IsRowSelected(1) <> True And objMatrix.VisualRowCount < 1 Then

                    objMatrix.AddRow()

                    oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)

                    objMatrix.SetLineData(objMatrix.VisualRowCount)

                End If

                If objMatrix.IsRowSelected(row) = True Then

                    objMatrix.DeleteRow(row)

                Else

                    For i As Integer = 1 To objMatrix.VisualRowCount - 1

                        If objMatrix.IsRowSelected(i) = True Then

                            objMatrix.DeleteRow(i)

                        End If

                    Next

                End If

                For i As Integer = 1 To objMatrix.VisualRowCount

                    objMatrix.Columns.Item("LineId").Cells.Item(i).Specific.Value = i

                Next
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If
            End If


        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("S_Rebate", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Me.UpdateBAG(objForm.UniqueID)

                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then

                        '--Madhu
                        LoadBrandBasedOnBPCD(objForm)
                        ''' ================ LOAD GL ACCOUNT FROM Retail Configuration=================
                        'Try

                        '    Dim oMatrix As SAPbouiCOM.Matrix
                        '    Dim rsGLAccount As SAPbobsCOM.Recordset
                        '    Dim strGLAccount As String = ""

                        '    oMatrix = objForm.Items.Item("Mtx").Specific

                        '    rsGLAccount = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        '    rsGLAccount.DoQuery("SELECT ""U_SuppReb"" FROM ""@ME_ACFG"" ")

                        '    If rsGLAccount.RecordCount > 0 Then
                        '        strGLAccount = rsGLAccount.Fields.Item("U_SuppReb").Value.ToString().Trim()
                        '    End If

                        '    If strGLAccount <> "" Then

                        '        For row As Integer = 1 To oMatrix.RowCount

                        '            Try
                        '                CType(oMatrix.Columns.Item("GLact").Cells.Item(row).Specific, SAPbouiCOM.EditText).Value = strGLAccount
                        '            Catch
                        '            End Try

                        '        Next

                        '    End If

                        'Catch ex As Exception
                        '    objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        'End Try

                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            objForm.Items.Item("Submit").Enabled = False
                        End If


                        '---Madhu
                        ' After Approving document should be totally disabled
                        Dim statusSupplierRebate As String = objForm.Items.Item("Status").Specific.Value.Trim().ToUpper()
                        statusSupplierRebate = statusSupplierRebate.Replace(" ", "")

                        ' Get Login User Code
                        Dim strUserCode As String = objMain.objApplication.Company.UserName.Trim().ToUpper()

                        If statusSupplierRebate = "APPROVED" OrElse statusSupplierRebate = "A" OrElse statusSupplierRebate = "SUBMITTED" OrElse statusSupplierRebate = "S" OrElse statusSupplierRebate = "REJECTED" OrElse statusSupplierRebate = "R" Then
                            Try : objForm.Items.Item("Submit").Enabled = False : Catch : End Try
                            If strUserCode = "MANAGER" Then
                                Try
                                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                    objForm.Items.Item("Submit").Enabled = False
                                Catch
                                End Try

                            Else

                                Try
                                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                                Catch
                                End Try

                            End If

                            For i As Integer = 0 To objForm.Items.Count - 1

                                Try

                                    If strUserCode = "MANAGER" Then

                                        If objForm.Items.Item(i).UniqueID <> "Mtx" Then
                                            objForm.Items.Item(i).Enabled = False
                                        End If

                                    Else

                                        objForm.Items.Item(i).Enabled = False

                                    End If

                                Catch
                                End Try

                            Next

                            Dim oMatrix1 As SAPbouiCOM.Matrix
                            oMatrix1 = objForm.Items.Item("Mtx").Specific

                            For j As Integer = 0 To oMatrix1.Columns.Count - 1

                                Try
                                    oMatrix1.Columns.Item(j).Editable = False
                                Catch
                                End Try

                            Next

                            If strUserCode = "MANAGER" Then

                                Try
                                    oMatrix1.Columns.Item("Brand").Editable = True
                                    oMatrix1.Columns.Item("RBTPRC").Editable = True
                                    oMatrix1.Columns.Item("RAmount").Editable = True
                                    ' oMatrix1.Columns.Item("GLact").Editable = True
                                Catch
                                End Try

                            End If

                            Try : objForm.Items.Item("1").Enabled = True : Catch : End Try
                            Try : objForm.Items.Item("2").Enabled = True : Catch : End Try

                            Try : objForm.EnableMenu("1281", True) : Catch : End Try
                            Try : objForm.EnableMenu("1282", True) : Catch : End Try

                        Else

                            Try : objForm.EnableMenu("1281", True) : Catch : End Try
                            Try : objForm.Items.Item("Submit").Enabled = True : Catch : End Try

                            Try
                                If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                End If
                            Catch
                            End Try

                        End If

                    End If


            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    'Approval

    Sub ApprovalTrigger(ByVal DocType As String, ByVal AppID As String, ByVal Table As String, ByVal AppIDField As String, ByVal AppStatField As String, ByVal Document As String)
        Dim ff As Integer
        Dim draftquery As String
        Dim approverid As String
        'Dim docdate As String = objForm.Items.Item("TNXDDT").Specific.Value

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

                    draftquery = " SELECT * FROM ""@SBO_DD"" WHERE ""U_statusN"" IN ('S','A') AND ""U_Isdraft"" IN ('Y','N') AND ""U_docnum"" = '" & AppID & "';"
                    ' draftquery = " SELECT * FROM ""@SBO_DD"" WHERE ""U_statusN"" IN ('S','A') AND ""U_Isdraft"" IN ('Y','N') AND ""U_AppID"" = '" & AppID & "';"

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
                        ''.CreateForm(objForm.UniqueID, objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value)
                        objMain.objApplication.ActivateMenuItem("1304")
                    End If
                Else
                    objMain.objApplication.StatusBar.SetText("Login User Don't have Approvals Please Define Approvals and Add the Document", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
            Else Exit Sub
            End If
        End If
    End Sub


    'Application Id
    Public Function AutoDocentryNumber()

        'Dim oRsDocNum As SAPbobsCOM.Recordset = objForm.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objForm = objMain.objApplication.Forms.Item(objForm.UniqueID)
        oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE")
        Dim oRsDocNum As SAPbobsCOM.Recordset

        oRsDocNum = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim Query1 As String = "SELECT ifnull(Max(""DocNum""),0)+1 AS ""DocNum"" FROM ""@SBO_SUPREBATE"" "

        oRsDocNum.DoQuery(Query1)

        oDBs_Head.SetValue("DocNum", 0, oRsDocNum.Fields.Item(0).Value)

        Dim rsAppId As SAPbobsCOM.Recordset

        rsAppId = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim str As String = "SELECT 'SP' || LPAD('',6 - LENGTH(LTRIM(RTRIM(REPLACE(IFNULL(MAX(""U_AppID""), '0'), 'SP', '') + 1))),'0') || LTRIM(RTRIM(REPLACE(IFNULL(MAX(""U_AppID""), '0'), 'SP', '') + 1)) FROM ""@SBO_SUPREBATE"";"

        rsAppId.DoQuery(str)
        Dim appp As String = rsAppId.Fields.Item(0).Value.ToString().Trim()

        oDBs_Head.SetValue("U_AppID", oDBs_Head.Offset, appp)


    End Function

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
            "SELECT COUNT(*) FROM ""@SBO_SUPREBATE"" WHERE ""U_AppID"" = '" & docnumcc & "'"

            rs.DoQuery(checkQry)

            If rs.Fields.Item(0).Value > 0 Then


                Dim Str As String =
            "SELECT 'SP' || LPAD('',6 - LENGTH(LTRIM(RTRIM(REPLACE(IFNULL(MAX(""U_AppID""), '0'), 'SP', '') + 1))),'0') || LTRIM(RTRIM(REPLACE(IFNULL(MAX(""U_AppID""), '0'), 'SP', '') + 1)) FROM ""@SBO_SUPREBATE"";"



                rs.DoQuery(Str)

                oDBs_Head.SetValue("U_AppID", 0, rs.Fields.Item(0).Value)
                Dim newReqNo As String = rs.Fields.Item(0).Value.ToString()
                objMain.objApplication.StatusBar.SetText("Duplicate detected. New Supplier Rebate Number generated: " & newReqNo, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "Problem while processing Auto ID : " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Public Sub UpdateBAG(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim UpdatePurchaseAgg As String = "Update ""OOAT"" SET ""U_SUPREBT"" = '" & sDoNum & "'  Where ""AbsID"" ='" & sBaseEntry & "'"
            Dim oRsUpdatePurchaseAgg As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsUpdatePurchaseAgg.DoQuery(UpdatePurchaseAgg)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    'madhu loading Supplier Reabtes details

    Private Sub LoadItemsToSupplierRebate(ByVal FormUID As String, ByVal BpCode As String, ByVal ParentDocEntry As String)

        Dim objForm As SAPbouiCOM.Form =
        objMain.objApplication.Forms.Item(FormUID)

        Dim oMatrix As SAPbouiCOM.Matrix =
        CType(objForm.Items.Item("Item_10").Specific, SAPbouiCOM.Matrix)

        Dim oDBs As SAPbouiCOM.DBDataSource =
        objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C2")

        Dim oRs As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim sQuery As String = "SELECT T1.""ItemCode"", T1.""ItemName"",  T1.""PlanQty"",T1.""UomCode"",T2.""CodeBars"" " & "FROM ""OOAT"" T0 " & "INNER JOIN ""OAT1"" T1 ON T0.""AbsID"" = T1.""AgrNo"" " &
                                "INNER JOIN ""OITM"" T2 ON T1.""ItemCode"" = T2.""ItemCode"" " & "WHERE T0.""BpCode"" = '" & BpCode & "' " & "AND IFNULL(T2.""U_REBEGBLE"", 'No') = 'Yes'"


        oRs.DoQuery(sQuery)

        If oRs.RecordCount = 0 Then Exit Sub

        objMatrix2.Clear()
        oDBs_Details2.Clear()

        objMatrix2.FlushToDataSource()

        For i As Integer = 0 To oRs.RecordCount - 1

            objMatrix2.AddRow()

            oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, objMatrix2.VisualRowCount)

            oDBs_Details2.SetValue("U_ITMCD", oDBs_Details2.Offset, oRs.Fields.Item("ItemCode").Value.ToString())
            oDBs_Details2.SetValue("U_ITMNM", oDBs_Details2.Offset, oRs.Fields.Item("ItemName").Value.ToString())
            oDBs_Details2.SetValue("U_UOM", oDBs_Details2.Offset, oRs.Fields.Item("UomCode").Value.ToString())
            oDBs_Details2.SetValue("U_BRCD", oDBs_Details2.Offset, oRs.Fields.Item("CodeBars").Value.ToString())
            oDBs_Details2.SetValue("U_QTY", oDBs_Details2.Offset, oRs.Fields.Item("PlanQty").Value.ToString())

            objMatrix2.SetLineData(objMatrix2.VisualRowCount)
            oRs.MoveNext()
        Next

        'oRs.MoveFirst()
        'Dim i As Integer = 0

        'While Not oRs.EoF

        '    oDBs_Details2.InsertRecord(i)


        '    oDBs_Details2.SetValue("U_ITMCD", i, oRs.Fields.Item("ItemCode").Value.ToString())
        '    oDBs_Details2.SetValue("U_ITMNM", i, oRs.Fields.Item("ItemName").Value.ToString())
        '    oDBs_Details2.SetValue("U_UOM", i, oRs.Fields.Item("UomCode").Value.ToString())
        '    oDBs_Details2.SetValue("U_BRCD", i, oRs.Fields.Item("CodeBars").Value.ToString())
        '    oDBs_Details2.SetValue("U_QTY", i, oRs.Fields.Item("PlanQty").Value.ToString())

        '    oRs.MoveNext()
        '    i += 1


        'End While
        'oMatrix.LoadFromDataSource()
        'oMatrix.AutoResizeColumns()

    End Sub

    Private Sub LoadItemsToSupplierRebate2(ByVal FormUID As String,
                                     ByVal BpCode As String,
                                     ByVal ParentDocEntry As String)

        Dim objForm As SAPbouiCOM.Form =
        objMain.objApplication.Forms.Item(FormUID)

        Dim objMatrix2 As SAPbouiCOM.Matrix =
        CType(objForm.Items.Item("Item_10").Specific, SAPbouiCOM.Matrix)

        Dim oDBs_Details2 As SAPbouiCOM.DBDataSource =
        objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C2")

        Dim oRs As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim sQuery As String =
        "SELECT T1.""ItemCode"", T1.""ItemName"", T1.""PlanQty"", " &
        "T1.""UomCode"", T2.""CodeBars"" " &
        "FROM ""OOAT"" T0 " &
        "INNER JOIN ""OAT1"" T1 ON T0.""AbsID"" = T1.""AgrNo"" " &
        "INNER JOIN ""OITM"" T2 ON T1.""ItemCode"" = T2.""ItemCode"" " &
        "WHERE T0.""BpCode"" = '" & BpCode & "' " &
        "AND IFNULL(T2.""U_REBEGBLE"", 'No') = 'Yes'"

        oRs.DoQuery(sQuery)

        If oRs.RecordCount = 0 Then Exit Sub

        objMatrix2.Clear()
        oDBs_Details2.Clear()

        oRs.MoveFirst()

        For i As Integer = 1 To oRs.RecordCount

            oDBs_Details2.InsertRecord(oDBs_Details2.Size)
            oDBs_Details2.Offset = oDBs_Details2.Size - 1

            oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, i)
            oDBs_Details2.SetValue("U_ITMCD", oDBs_Details2.Offset, oRs.Fields.Item("ItemCode").Value.ToString())
            oDBs_Details2.SetValue("U_ITMNM", oDBs_Details2.Offset, oRs.Fields.Item("ItemName").Value.ToString())
            oDBs_Details2.SetValue("U_UOM", oDBs_Details2.Offset, oRs.Fields.Item("UomCode").Value.ToString())
            oDBs_Details2.SetValue("U_BRCD", oDBs_Details2.Offset, oRs.Fields.Item("CodeBars").Value.ToString())
            oDBs_Details2.SetValue("U_QTY", oDBs_Details2.Offset, oRs.Fields.Item("PlanQty").Value.ToString())

            'objMatrix2.SetLineData(i)

            oRs.MoveNext()
        Next

        objMatrix2.AutoResizeColumns()

    End Sub


#Region " Right Click Event"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Dim objForm As SAPbouiCOM.Form
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)

        objMatrix1 = objForm.Items.Item("Mtx1").Specific
        Try
            If eventInfo.FormUID = objForm.UniqueID Then
                If (eventInfo.BeforeAction = True) Then
                    If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                        objMatrix = objForm.Items.Item("Mtx").Specific
                        If eventInfo.ItemUID = "Mtx" And eventInfo.ColUID = "#" And objMatrix.VisualRowCount > 1 Then
                            'Try
                            '    oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                            '    oMenus = oMenuItem.SubMenus
                            '    If oMenus.Exists("Delete Row") = False Then
                            '        oCreationPackage.UniqueID = "Delete Row"
                            '        oCreationPackage.String = "Delete Row"
                            '        oCreationPackage.Enabled = True
                            '        oMenus.AddEx(oCreationPackage)
                            '    End If
                            'Catch ex As Exception
                            '    objMain.objApplication.StatusBar.SetText(ex.Message)
                            'End Try

                        ElseIf eventInfo.ItemUID = "Item_10" And eventInfo.ColUID = "#" And objMatrix2.VisualRowCount > 1 Then
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
                        ElseIf eventInfo.ItemUID = "Item_10" And eventInfo.ColUID = "#" And objMatrix2.VisualRowCount <= 1 Then
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


                        ElseIf eventInfo.ItemUID = "Mtx" And objMatrix.RowCount <= 1 Then
                            oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                            oMenus = oMenuItem.SubMenus
                            Try
                                If oMenus.Exists("Delete Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Delete Row")
                                End If
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try


                        ElseIf eventInfo.ItemUID = "Mtx1" And eventInfo.ColUID = "#" And objMatrix1.VisualRowCount > 1 Then
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
                        ElseIf eventInfo.ItemUID = "Mtx1" And objMatrix1.VisualRowCount <= 1 Then
                            oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                            oMenus = oMenuItem.SubMenus
                            Try
                                If oMenus.Exists("Delete Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Delete Row")
                                End If
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If

                    Else

                        Try
                            oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                            oMenus = oMenuItem.SubMenus
                            If oMenus.Exists("Delete Row") = True Then
                                objMain.objApplication.Menus.RemoveEx("Delete Row")
                            End If
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If

                End If

            End If

            ' System.Diagnostics.Process.Start()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

    Sub FilterUOM(ByVal FormUID As String, ByVal CFLID As String, ByVal ItemCode As String)
        Try
            Dim Query As String = "SELECT T3.""UomCode""  FROM OITM T0  INNER JOIN OUGP T1 ON T0.""UgpEntry""  = T1.""UgpEntry"" " &
                                    "INNER JOIN UGP1 T2 ON T1.""UgpEntry""  = T2.""UgpEntry""  INNER JOIN OUOM T3 ON " &
                            "T2.""UomEntry""  = T3.""UomEntry""  WHERE T0.""ItemCode""  ='" & ItemCode.Trim & "'"

            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsQuery.DoQuery(Query)

            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFLID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()

            If oRsQuery.RecordCount > 0 Then

                For i As Integer = 1 To oRsQuery.RecordCount
                    oCondition = oConditions.Add()
                    oCondition.Alias = "UomCode"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = oRsQuery.Fields.Item(0).Value.ToString().Trim
                    If i <> oRsQuery.RecordCount Then
                        oCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    End If
                    oRsQuery.MoveNext()
                Next

            Else

                oCondition = oConditions.Add()
                oCondition.Alias = "UomCode"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = ""
            End If


            oChooseFromList.SetConditions(oConditions)

        Catch ex As Exception

        End Try
    End Sub


    Private Sub DisableColumns()

        Try
            objForm.Freeze(True)

            Me.objMatrix.Columns.Item("POAF").Visible = False
            Me.objMatrix.Columns.Item("POAT").Visible = False

        Catch ex As Exception
            objForm.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        Finally
            objForm.Freeze(False)
        End Try

    End Sub

    Public Sub LoadBrandBasedOnBPCD(ByVal objForm As SAPbouiCOM.Form)

        Try
            objForm.Freeze(True)

            '================ MATRIX & COLUMN =================
            Dim oMatrixBrand As SAPbouiCOM.Matrix =
            objForm.Items.Item("Mtx").Specific

            Dim oCol As SAPbouiCOM.Column =
            oMatrixBrand.Columns.Item("Brand")

            '================ IMPORTANT: FULL CLEAR =================
            Try
                Dim i As Integer

                'SAFE CLEAR (reverse loop is important in SAP B1)
                For i = oCol.ValidValues.Count - 1 To 0 Step -1
                    oCol.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next

            Catch
            End Try

            'Force refresh of combo metadata (VERY IMPORTANT)
            oCol.DisplayDesc = True

            '================ GET BPCD =================
            Dim strBPCD As String = ""

            Try
                strBPCD = objForm.Items.Item("BPCD").Specific.Value.Trim()
            Catch
                strBPCD = ""
            End Try

            If strBPCD = "" Then
                Dim oDB As SAPbouiCOM.DBDataSource =
                objForm.DataSources.DBDataSources.Item(0)

                strBPCD = oDB.GetValue("U_BPCD", 0).Trim()
            End If

            If strBPCD = "" Then Exit Sub

            '================ QUERY =================
            Dim sGetBrand As String =
            "SELECT DISTINCT IFNULL(T0.""U_BRAND"",'') AS ""Code"" " &
            "FROM OITM T0 " &
            "WHERE T0.""CardCode"" = '" & strBPCD & "' " &
            "AND IFNULL(T0.""U_BRAND"",'') <> ''"

            Dim rsBrand As SAPbobsCOM.Recordset =
            objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            rsBrand.DoQuery(sGetBrand)

            '================ BUILD UNIQUE LIST =================
            Dim dt As New List(Of String)

            While Not rsBrand.EoF

                Dim code As String = rsBrand.Fields.Item("Code").Value.ToString().Trim()

                If code <> "" AndAlso Not dt.Contains(code) Then
                    dt.Add(code)
                End If

                rsBrand.MoveNext()

            End While

            'ADD DEFAULT
            dt.Add("All")

            '================ ADD FRESH VALUES =================
            For Each val As String In dt
                Try
                    oCol.ValidValues.Add(val, val)
                Catch
                    'ignore duplicates
                End Try
            Next

            oMatrixBrand.LoadFromDataSource()

            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(
            "Brand loaded successfully",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception

            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(
            ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

End Class

