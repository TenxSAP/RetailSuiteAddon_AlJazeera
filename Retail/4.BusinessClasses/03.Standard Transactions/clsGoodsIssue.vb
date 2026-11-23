Public Class clsGoodsIssue

#Region "        Declaration        "
    Dim objForm, AltForm, objMIForm As SAPbouiCOM.Form
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim GetInTransitWhs As String
    Dim oRsGetInTransitWhs As SAPbobsCOM.Recordset
    Dim TRANSFLAG As String = ""
    Dim oRsTRANSFLAG As SAPbobsCOM.Recordset
    Dim oItem As SAPbouiCOM.Item

    Dim oDBs_Head As SAPbouiCOM.DBDataSource
#End Region

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("OIGE")
                    objMatrix = objForm.Items.Item("13").Specific

                    If pVal.BeforeAction = False Then
                        Me.AddItems(objForm.UniqueID)
                        GetInTransitWhs = "SELECT *  FROM ""@ME_ACFG""  T0 WHERE T0.""U_INTRW"" ='N'"
                        oRsGetInTransitWhs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetInTransitWhs.DoQuery(GetInTransitWhs)
                        If oRsGetInTransitWhs.RecordCount > 0 Then

                            objForm.Items.Item("MI").Visible = False
                            objForm.Items.Item("U_MIDNUM").Visible = False
                            objForm.Items.Item("U_MIDENT").Visible = False
                        End If

                    End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix = objForm.Items.Item("13").Specific


                    If pVal.ItemUID = "lin1" And pVal.BeforeAction = False Then
                        Me.LoadMI(objForm.UniqueID)
                    End If

                    If pVal.ItemUID = "fldRtl" Then
                        objForm.PaneLevel = "46"
                    End If

                    If pVal.ItemUID = "btn_ED" And pVal.BeforeAction = False Then
                        Try
                            objForm.Freeze(True)


                            'Dim GetData As String = "SELECT T1.""LineNum"",T1.""ItemCode"", T1.""Dscription"", T1.""Quantity"", T1.""CodeBars"", " &
                            '                        "T1.""UomCode"" FROM OIGE T0  INNER JOIN IGE1 T1 ON T0.""DocEntry"" = T1.""DocEntry""  " &
                            '                        " INNER JOIN OWHS T2 ON T1.""WhsCode"" = T2.""WhsCode"" WHERE  " &
                            '                        "T0.""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0) & "' And IFNULL(T2.""U_EXCHK"",'No') ='Yes' And T0.""DocType""='I' ORDER BY T1.""LineNum"""

                            'Dim oRsGetData As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            'oRsGetData.DoQuery(GetData)

                            'If oRsGetData.RecordCount > 0 Then
                            '    objMain.objExpiry.CreateForm(objForm.UniqueID, "OIGE", oDBs_Head.GetValue("DocEntry", 0), oDBs_Head.GetValue("DocNum", 0), oRsGetData)
                            'Else
                            '    objMain.objApplication.StatusBar.SetText("Expiry Date screen not applicable")
                            'End If



                            Dim GetWCnd As String = "select IFNULL(""U_EXCHK"",'No') FROM OWHS Where ""WhsCode""='" & objMatrix.Columns.Item("15").Cells.Item(1).Specific.Value & "' And ""U_EXCHK""='Yes' And ""U_EXCHK"" IS NOT NULL And ""U_EXCHK""<>''"
                            Dim oRsGetWCnd As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetWCnd.DoQuery(GetWCnd)

                            If oRsGetWCnd.RecordCount = 0 Then
                                objMain.objApplication.StatusBar.SetText("Expiry Screen not applicable")
                                objForm.Freeze(False)
                                Exit Try
                            End If


                            Dim dt As New DataTable()
                            Dim UID As String = String.Empty

                            If CStr(objForm.Items.Item("txt_GUID").Specific.Value).Trim <> "" Then
                                UID = objForm.Items.Item("txt_GUID").Specific.Value
                                'GoTo GotoLst
                            Else

                                UID = System.Guid.NewGuid.ToString()
                                objForm.Items.Item("txt_GUID").Specific.Value = UID.Trim
                            End If

                            ' Define columns for the DataTable
                            dt.Columns.Add("LineNum", GetType(Integer))
                            dt.Columns.Add("ItemCode", GetType(String))
                            dt.Columns.Add("Dscription", GetType(String))
                            dt.Columns.Add("Quantity", GetType(Double))
                            dt.Columns.Add("CodeBars", GetType(String))
                            dt.Columns.Add("UomCode", GetType(String))
                            dt.Columns.Add("BQuantity", GetType(Double))


                            For i As Integer = 1 To objMatrix.VisualRowCount

                                If CStr(objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value).Trim <> "" Then

                                    Dim dr As DataRow = dt.NewRow
                                    dr("LineNum") = i - 1
                                    dr("ItemCode") = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value
                                    dr("Dscription") = objMatrix.Columns.Item("2").Cells.Item(i).Specific.Value
                                    dr("Quantity") = objMatrix.Columns.Item("9").Cells.Item(i).Specific.Value
                                    dr("CodeBars") = objMatrix.Columns.Item("26").Cells.Item(i).Specific.Value
                                    dr("UomCode") = objMatrix.Columns.Item("1470001034").Cells.Item(i).Specific.Value
                                    dr("BQuantity") = objMatrix.Columns.Item("9").Cells.Item(i).Specific.Value


                                    dt.Rows.Add(dr)
                                End If


                            Next

GotoLst:                    objMain.objExpiry.CreateFormNew(objForm.UniqueID, "OIGE", oDBs_Head.GetValue("DocEntry", 0), oDBs_Head.GetValue("DocNum", 0), dt, UID.Trim, objForm.Mode)

                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Validation(objForm.UniqueID) = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    If pVal.ItemUID = "2310000078" And pVal.BeforeAction = False Then
                        'Dim DocNum As String = objForm.Items.Item("DocNum").Specific.Value

                        Dim objBaseForm As SAPbouiCOM.Form
                        objBaseForm = objMain.objApplication.Forms.GetForm("721", objMain.objApplication.Forms.ActiveForm.TypeCount)
                        objBaseForm.Items.Item("U_GI").Specific.Value = objForm.Items.Item("7").Specific.Value
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("720", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Try

                            If CStr(oDBs_Head.GetValue("U_GUID", 0)).Trim <> "" Then

                                Dim UpdateDa As String = "Update ""@TNXEDS"" SET ""U_BENTRY""='" & oDBs_Head.GetValue("DocEntry", 0) & "', " &
                                              """U_BNO""='" & oDBs_Head.GetValue("DocNum", 0) & "' Where ""U_GUID""='" & oDBs_Head.GetValue("U_GUID", 0) & "'  And ""U_DOCTY""='OIGE'"

                                Dim oRsUpdateDa As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                oRsUpdateDa.DoQuery(UpdateDa)

                            End If

                            Dim GetInventoryDtls As String = " Select ""DocEntry"" As ""ITEntry"" , ""U_INVTRREN"" As ""STREntery"" From IGE1 " &
                                                             " Where ""DocEntry"" = (Select Max(""DocEntry"")  From OIGE) And " & IsNull & "(""U_INVTRREN"",'') <> '' " &
                                                             " Group by  ""DocEntry"",  ""U_INVTRREN"" "
                            Dim oRsGetInventoryDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetInventoryDtls.DoQuery(GetInventoryDtls)
                            If oRsGetInventoryDtls.RecordCount > 0 Then
                                oRsGetInventoryDtls.MoveFirst()
                                For i As Integer = 1 To oRsGetInventoryDtls.RecordCount
                                    Me.UpdateSTR(objForm.UniqueID, oRsGetInventoryDtls.Fields.Item("ITEntry").Value, oRsGetInventoryDtls.Fields.Item("STREntery").Value)
                                    oRsGetInventoryDtls.MoveNext()
                                Next
                            End If
                            If oDBs_Head.GetValue("U_INVTRREQ", 0) <> "" Then
                                'Dim oRsUpdate1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                'Dim Update1 As String = "select ""DocEntry"" from Where ""U_DocNum""= '" & oRsUpdate1.Fields.Item("BaseEntry").Value & "'"
                                'oRsUpdate1.DoQuery(Update1)
                                Dim oRsUpdate2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim Update2 As String = "Update ""@ME_STR"" Set ""U_GINum""='" & oDBs_Head.GetValue("DocNum", 0) & "',""U_GIDoc""='" & oDBs_Head.GetValue("DocEntry", 0) & "'Where ""DocNum""= '" & oDBs_Head.GetValue("U_INVTRREQ", 0) & "'"
                                oRsUpdate2.DoQuery(Update2)
                            End If
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Try
                            Dim GetInventoryDtls As String = " Select ""DocEntry"" As ""ITEntry"" , ""U_INVTRREN"" As ""STREntery"" From IGE1 " &
                                                             " Where ""DocEntry"" = (Select Max(""DocEntry"")  From OIGE) And " & IsNull & "(""U_INVTRREN"",'') <> '' " &
                                                             " Group by  ""DocEntry"",  ""U_INVTRREN"" """
                            Dim oRsGetInventoryDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetInventoryDtls.DoQuery(GetInventoryDtls)
                            If oRsGetInventoryDtls.RecordCount > 0 Then
                                oRsGetInventoryDtls.MoveFirst()
                                For i As Integer = 1 To oRsGetInventoryDtls.RecordCount
                                    Me.UpdateSTR(objForm.UniqueID, oRsGetInventoryDtls.Fields.Item("ITEntry").Value, oRsGetInventoryDtls.Fields.Item("STREntery").Value)
                                    oRsGetInventoryDtls.MoveNext()
                                Next
                            End If
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    Sub AddItems(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objMain.objUtilities.AddLabel(objForm.UniqueID, "MI", objForm.Items.Item("8").Top + 60, objForm.Items.Item("8").Left, 30, "MI No.", "8")
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "U_MIDNUM", objForm.Items.Item("7").Top + 60, objForm.Items.Item("7").Left, (objForm.Items.Item("7").Width / 2) - 1, "OIGE", "U_MIDNUM", "7")
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "U_MIDENT", objForm.Items.Item("7").Top + 60, objForm.Items.Item("7").Left + objForm.Items.Item("U_MIDNUM").Width + 1, objForm.Items.Item("7").Width / 2, "OIGE", "U_MIDENT", "7")

            objMain.objUtilities.AddLinkButton(objForm.UniqueID, "lin1", objForm.Items.Item("7").Top + 60, objForm.Items.Item("7").Left - 20, "U_MIDNUM", -1)


            objForm.Items.Item("U_MIDNUM").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("U_MIDNUM").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("U_MIDENT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("U_MIDENT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objMain.objUtilities.AddButton(objForm.UniqueID, "btn_ED", objForm.Items.Item("2").Top,
                                      objForm.Items.Item("2").Width + objForm.Items.Item("2").Left + 5, "100", objForm.Items.Item("2").Height, "2", "Expiry Date", objForm.Items.Item("2").FromPane, objForm.Items.Item("2").ToPane)

            objMain.objUtilities.AddEditBox(objForm.UniqueID, "txt_GUID", objForm.Items.Item("4").Top, objForm.Items.Item("4").Left + objForm.Items.Item("4").Width + 5, objForm.Items.Item("4").Width, "OIGE", "U_GUID", "", objForm.Items.Item("4").Height, 96, 96)

            objForm.Items.Item("txt_GUID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            oItem = objForm.Items.Add("fldRtl", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem.Top = objForm.Items.Item("1320000079").Top
            oItem.Left = objForm.Items.Item("1320000079").Left + objForm.Items.Item("1320000079").Width + 10
            oItem.Width = objForm.Items.Item("1320000079").Width + 30
            oItem.Height = objForm.Items.Item("1320000079").Height
            oItem.AffectsFormMode = False
            Dim fld As SAPbouiCOM.Folder = objForm.Items.Item("fldRtl").Specific
            fld.Caption = "Retail"
            fld.GroupWith("1320000079")
            fld.Pane = "46"


            objMain.objUtilities.AddLabel(objForm.UniqueID, "PDTUsr", objForm.Items.Item("20").Top + 100, objForm.Items.Item("20").Left + 5, objForm.Items.Item("20").Width, "PDT User", "84", objForm.Items.Item("20").Height, 46, 46)
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "PDTUSER", objForm.Items.Item("PDTUsr").Top, objForm.Items.Item("PDTUsr").Left + 120, objForm.Items.Item("PDTUsr").Width + 5, "OIGE", "U_PDTUSER", "PDTUsr", objForm.Items.Item("PDTUsr").Height, 46, 46)

            objMain.objUtilities.AddLabel(objForm.UniqueID, "UserNm", objForm.Items.Item("20").Top + 120, objForm.Items.Item("20").Left + 5, objForm.Items.Item("20").Width + 13, "User Name", "PDTUSER", objForm.Items.Item("20").Height, 46, 46)
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "UserName", objForm.Items.Item("UserNm").Top, objForm.Items.Item("UserNm").Left + 120, objForm.Items.Item("PDTUsr").Width + 5, "OIGE", "U_UserName", "UserNm", objForm.Items.Item("UserNm").Height, 46, 46)

            objMain.objUtilities.AddLabel(objForm.UniqueID, "Trty", objForm.Items.Item("20").Top + 140, objForm.Items.Item("20").Left + 5, objForm.Items.Item("20").Width + 13, "Transaction Type", "PDTUSER", objForm.Items.Item("20").Height, 46, 46)
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "TRTYPE", objForm.Items.Item("Trty").Top, objForm.Items.Item("Trty").Left + 120, objForm.Items.Item("PDTUsr").Width + 5, "OIGE", "U_TRTYPE", "Trty", objForm.Items.Item("Trty").Height, 46, 46)

            objMain.objUtilities.AddLabel(objForm.UniqueID, "SRN", objForm.Items.Item("20").Top + 160, objForm.Items.Item("20").Left + 5, objForm.Items.Item("20").Width + 20, "Stock Request Number", "PDTUSER", objForm.Items.Item("20").Height, 46, 46)
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "INVTRREQ", objForm.Items.Item("SRN").Top, objForm.Items.Item("SRN").Left + 120, objForm.Items.Item("PDTUsr").Width + 5, "OIGE", "U_INVTRREQ", "SRN", objForm.Items.Item("SRN").Height, 46, 46)

            objMain.objUtilities.AddLabel(objForm.UniqueID, "SRE", objForm.Items.Item("20").Top + 180, objForm.Items.Item("20").Left + 5, objForm.Items.Item("20").Width + 13, "Stock Request Entry", "PDTUSER", objForm.Items.Item("20").Height, 46, 46)
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "INVTRREN", objForm.Items.Item("SRE").Top, objForm.Items.Item("SRE").Left + 120, objForm.Items.Item("PDTUsr").Width + 5, "OIGE", "U_INVTRREN", "SRE", objForm.Items.Item("SRE").Height, 46, 46)


            '---Madhu To Warehouse Name and To Branch two fields
            objMain.objUtilities.AddLabel(objForm.UniqueID, "lbTAmt", objForm.Items.Item("1000001").Top + 22,
           objForm.Items.Item("1000001").Left, objForm.Items.Item("1000001").Width + 10, "To Warehouse Name", "&To Warehouse Name", 0, 0, objForm.Items.Item("1000001").Height)

            objMain.objUtilities.AddEditBox(objForm.UniqueID, "txtTAmt", objForm.Items.Item("U_GIT").Top + 20, objForm.Items.Item("U_GIT").Left + 2, 140,
            "OIGE", "U_TOWN", "U_GIT", objForm.Items.Item("U_GIT").Height, objForm.Items.Item("U_GIT").FromPane, objForm.Items.Item("U_GIT").ToPane)

            objMain.objUtilities.AddLabel(objForm.UniqueID, "lbTBranch", objForm.Items.Item("1000001").Top + 38, objForm.Items.Item("1000001").Left,
                objForm.Items.Item("1000001").Width + 10, "To Branch", "&To Branch", 0, 0, objForm.Items.Item("1000001").Height)

            objMain.objUtilities.AddEditBox(objForm.UniqueID, "txtTBrn", objForm.Items.Item("U_GIT").Top + 37, objForm.Items.Item("U_GIT").Left + 2,
            140, "OIGE", "U_TOBNCH", "U_GIT", objForm.Items.Item("U_GIT").Height, objForm.Items.Item("U_GIT").FromPane, objForm.Items.Item("U_GIT").ToPane)

            objForm.Items.Item("txtTAmt").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("txtTAmt").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("txtTAmt").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("txtTAmt").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("txtTBrn").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("txtTBrn").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("txtTBrn").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("txtTBrn").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)



        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub UpdateSTR(ByVal FormUID As String, ByVal DocEntry As String, ByVal STREntry As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim GetDetails As String = " Select ""ItemCode""  , (""U_METSTQTY"" - (""U_MEREQQTY"" - ""Quantity"")) ""DispatchQty"" , ""U_MEREQQTY"" ""STRQty"", " &
                                       " (""U_MEREQQTY"" - ""Quantity"") ""OpenQty"" , ""U_INVTRRLN"" As ""LineId"" , (""U_METSTQTY""-(""U_METSTQTY"" - (""U_MEREQQTY"" - ""Quantity""))) ""LineStatus""  From  IGE1 T1  Where ""DocEntry""  = '" & DocEntry & "' And ""Quantity"" <= ""U_MEREQQTY"""
            Dim oRsGetDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDetails.DoQuery(GetDetails)

            If oRsGetDetails.RecordCount > 0 Then
                oRsGetDetails.MoveFirst()
                objMain.sCmp = objMain.objCompany.GetCompanyService
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OSTR")
                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                objMain.oGeneralParams.SetProperty("DocEntry", STREntry)
                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                objMain.oChildren = objMain.oGeneralData.Child("ME_STR_C0")
                For i As Integer = 1 To oRsGetDetails.RecordCount
                    Dim LineId As Integer = CInt(oRsGetDetails.Fields.Item("LineId").Value)
                    objMain.oChildren.Item(LineId - 1).SetProperty("U_MERELQTY", oRsGetDetails.Fields.Item("DispatchQty").Value)
                    objMain.oChildren.Item(LineId - 1).SetProperty("U_MEOPNQTY", oRsGetDetails.Fields.Item("OpenQty").Value)
                    If oRsGetDetails.Fields.Item("OpenQty").Value = 0 Then
                        objMain.oChildren.Item(LineId - 1).SetProperty("U_MELNSTS", "Close")
                    Else
                        objMain.oChildren.Item(LineId - 1).SetProperty("U_MELNSTS", "Open")
                    End If
                    objMain.oGeneralService.Update(objMain.oGeneralData)
                    oRsGetDetails.MoveNext()
                Next
                'Dim GetOpenQty As String = "Select ""U_MEOPNQTY"" From ""@ME_STR_C0"" Where ""DocEntry"" = '" & STREntry & "' And ""U_MEOPNQTY"" > 0 And ""U_MEITMCD"" <> '' "
                'Dim oRsGetOpenQty As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oRsGetOpenQty.DoQuery(GetOpenQty)
                'If oRsGetOpenQty.RecordCount = 0 Then
                ''objMain.oGeneralData.SetProperty("U_MESTS", "Close")
                ''objMain.oGeneralService.Update(objMain.oGeneralData)
                'End If
                ''Adding the code to update the status of the stock request in the OIGE form after updating the STR lines on 31-08-2026 (RamyaY)
                Dim GetOpenQty As String = "SELECT COUNT(*) AS ""Cnt"" " &
                           "FROM ""@ME_STR_C0"" " &
                           "WHERE ""DocEntry"" = '" & STREntry & "' " &
                           "AND IFNULL(""U_MEOPNQTY"",0) > 0"

                Dim oRsGetOpenQty As SAPbobsCOM.Recordset =
                    objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                oRsGetOpenQty.DoQuery(GetOpenQty)

                If CInt(oRsGetOpenQty.Fields.Item("Cnt").Value) > 0 Then
                    objMain.oGeneralData.SetProperty("U_MESTS", "Open")
                Else
                    objMain.oGeneralData.SetProperty("U_MESTS", "Close")
                End If

                objMain.oGeneralService.Update(objMain.oGeneralData)


            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub LoadMI(ByVal FormUID As String)
        Try
            objMain.objApplication.ActivateMenuItem("ME_MI")
            objMIForm = objMain.objApplication.Forms.GetForm("ME_MI_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objMIForm.Freeze(True)
            objMIForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            objMIForm.Items.Item("13").Specific.value = objForm.Items.Item("U_MIDNUM").Specific.Value
            objMIForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            objMIForm.Freeze(False)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("13").Specific
            For i As Integer = 1 To objMatrix.VisualRowCount - 1

                Dim ItemCode, Itemcode2 As String
                ItemCode = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.ToString()
                If i = 1 Then
                    Itemcode2 = ItemCode

                ElseIf i > 1 Then
                    If Itemcode2.Contains(ItemCode) Then
                        objMain.objApplication.StatusBar.SetText("Same Itemcode is selected in LineId " & i & "...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                    Itemcode2 = ItemCode & "'" & "," & "'" & Itemcode2

                End If
            Next

            Dim CheckBackDOc As String = "Select ""U_GUID"" FROM ""@TNXEDS"" Where " &
                                                         """U_GUID""='" & CStr(objForm.Items.Item("txt_GUID").Specific.Value).Trim & "' And ""U_DOCTY""='OIGE'"

            Dim oRsCheckBackDOc As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oRsCheckBackDOc.DoQuery(CheckBackDOc)


            Dim GetWCnd As String = "select IFNULL(""U_EXCHK"",'No') FROM OWHS Where ""WhsCode""='" & objMatrix.Columns.Item("15").Cells.Item(1).Specific.Value & "' And ""U_EXCHK""='Yes' And ""U_EXCHK"" IS NOT NULL And ""U_EXCHK""<>''"
            Dim oRsGetWCnd As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetWCnd.DoQuery(GetWCnd)

            If oRsGetWCnd.RecordCount > 0 And (oRsCheckBackDOc.RecordCount = 0 Or CStr(CStr(objForm.Items.Item("txt_GUID").Specific.Value).Trim).Trim = "") Then
                objMain.objApplication.StatusBar.SetText("please add Expiry Screen before adding")
                objForm.Freeze(False)
                Return False

            End If

            Return True
        Catch ex As Exception
        End Try
    End Function

End Class
