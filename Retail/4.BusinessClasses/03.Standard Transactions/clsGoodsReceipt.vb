Public Class clsGoodsReceipt

#Region "        Declaration        "
    Dim objForm, AltForm, objMIForm As SAPbouiCOM.Form
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim GetInTransitWhs As String
    Dim oRsGetInTransitWhs, oRsGetIsseuNo As SAPbobsCOM.Recordset
    Dim TRANSFLAG As String = ""
    Dim GetIsseuNo As String = ""
    Dim oRsTRANSFLAG As SAPbobsCOM.Recordset
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim oItem As SAPbouiCOM.Item
    Dim CostBtnFlag As Boolean = False
    Dim ExecuteFlag As Boolean = False
#End Region

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objForm = objMain.objApplication.Forms.Item(FormUID)
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
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("OIGN")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("IGN1")
                    If pVal.ItemUID = "lin2" And pVal.BeforeAction = False Then
                        Me.LoadMI(objForm.UniqueID)
                    End If

                    If pVal.ItemUID = "fldRtl" Then
                        objForm.PaneLevel = "46"
                    End If


                    If pVal.ItemUID = "btn_ED" And pVal.BeforeAction = False Then
                        Try
                            objForm.Freeze(True)


                            'Dim GetData As String = "SELECT T1.""LineNum"",T1.""ItemCode"", T1.""Dscription"", T1.""Quantity"", T1.""CodeBars"", " &
                            '                        "T1.""UomCode"" FROM OIGN T0  INNER JOIN IGN1 T1 ON T0.""DocEntry"" = T1.""DocEntry""  " &
                            '                        " INNER JOIN OWHS T2 ON T1.""WhsCode"" = T2.""WhsCode"" WHERE  " &
                            '                        "T0.""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0) & "' And IFNULL(T2.""U_EXCHK"",'No') ='Yes' And T0.""DocType""='I' ORDER BY T1.""LineNum"""

                            'Dim oRsGetData As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            'oRsGetData.DoQuery(GetData)

                            'If oRsGetData.RecordCount > 0 Then
                            '    objMain.objExpiry.CreateForm(objForm.UniqueID, "OIGN", oDBs_Head.GetValue("DocEntry", 0), oDBs_Head.GetValue("DocNum", 0), oRsGetData)
                            'Else
                            '    objMain.objApplication.StatusBar.SetText("Expiry Date screen not applicable")
                            'End If


                            'If CStr(objForm.Items.Item("3").Specific.Value).Trim <> "I" Then
                            '    objMain.objApplication.StatusBar.SetText("Expiry Screen not applicable")
                            '    objForm.Freeze(False)
                            '    Exit Try
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

GotoLst:                    objMain.objExpiry.CreateFormNew(objForm.UniqueID, "OIGN", oDBs_Head.GetValue("DocEntry", 0), oDBs_Head.GetValue("DocNum", 0), dt, UID.Trim, objForm.Mode)


                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If Validation(objForm.UniqueID) = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If

                    If pVal.ItemUID = "Btn_UpdCst" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Try
                            objForm = objMain.objApplication.Forms.Item(FormUID)
                            objForm = objMain.objApplication.Forms.GetForm("721", objMain.objApplication.Forms.ActiveForm.TypeCount)
                            Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("13").Specific
                            Dim DocNum As String = objForm.Items.Item("7").Specific.Value
                            Dim Series As String = objForm.Items.Item("30").Specific.Value
                            Dim INo As String = objForm.Items.Item("U_GI").Specific.Value
                            For j As Integer = 1 To objMatrix.VisualRowCount - 1
                                'GetInTransitWhs = "SELECT T0.""Quantity"", T0.""ItemCode"" ,T0.""U_METWHSCD"",T0.""StockPrice"",T0.""NumPerMsr"",T0.""WhsCode"",T1.""DocNum"" FROM IGE1 T0  inner join OIGE T1 ON T0.""DocEntry""=T1.""DocEntry"" WHERE T0.""DocEntry"" ='" & oDBs_Details.GetValue("BaseEntry", j - 1) & "' And T0.""ItemCode""='" & oDBs_Details.GetValue("ItemCode", j - 1) & "'  "

                                GetInTransitWhs = "SELECT T1.""Price"" as ""Price"",T1.""StockPrice"",T1.""U_METWHSCD"" As ""ToWareHouse"" from OIGE T0  INNER JOIN IGE1 T1 ON T0.""DocEntry"" = T1.""DocEntry""  WHERE T0.""DocEntry"" ='" & oDBs_Details.GetValue("BaseEntry", j - 1) & "' And T1.""ItemCode""='" & oDBs_Details.GetValue("ItemCode", j - 1) & "'"
                                oRsGetInTransitWhs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetInTransitWhs.DoQuery(GetInTransitWhs)

                                If objMatrix.Columns.Item("1").Cells.Item(j).Specific.Value <> "" Then
                                    objMatrix.Columns.Item("10").Cells.Item(j).Specific.Value = oRsGetInTransitWhs.Fields.Item("StockPrice").Value
                                    objMatrix.Columns.Item("15").Cells.Item(j).Specific.Value = oRsGetInTransitWhs.Fields.Item("ToWareHouse").Value


                                End If
                            Next

                            CostBtnFlag = True
                            'GetIsseuNo = "SELECT ""U_GI"" FROM ""OIGN""  T0 WHERE T0.""DocNum"" = '" & DocNum & "' and T0.""Series"" = '" & Series & "' "
                            'oRsGetIsseuNo = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'oRsGetIsseuNo.DoQuery(GetIsseuNo)
                            'If oRsGetIsseuNo.RecordCount > 0 Then
                            '    Dim IssueNo As String = oRsGetIsseuNo.Fields.Item("U_GI").Value
                            '    Dim IssueDoc As String = oRsGetIsseuNo.Fields.Item("U_GIDoc").Value
                            '    Dim oRsUpdate2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            '    Dim Update2 As String = " UPDATE  ""OIGE""  SET ""U_GR""='" & DocNum & "'WHERE ""DocNum""='" & IssueNo & "'"
                            '    oRsUpdate2.DoQuery(Update2)
                            'End If
                            Dim oRsUpdate2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim Update2 As String = " UPDATE  ""OIGE""  SET ""U_GR""='" & DocNum & "'WHERE ""DocNum""='" & INo & "'"
                            oRsUpdate2.DoQuery(Update2)
                        Catch ex As Exception

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
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "U_MIDNUM", objForm.Items.Item("7").Top + 60, objForm.Items.Item("7").Left, (objForm.Items.Item("7").Width / 2) - 1, "OIGN", "U_MIDNUM", "7")
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "U_MIDENT", objForm.Items.Item("7").Top + 60, objForm.Items.Item("7").Left + objForm.Items.Item("U_MIDNUM").Width + 1, objForm.Items.Item("7").Width / 2, "OIGN", "U_MIDENT", "7")

            objMain.objUtilities.AddLinkButton(objForm.UniqueID, "lin2", objForm.Items.Item("7").Top + 60, objForm.Items.Item("7").Left - 20, "U_MIDNUM", -1)

            oItem = objForm.Items.Add("Btn_UpdCst", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Top = objForm.Items.Item("2").Top
            oItem.Left = objForm.Items.Item("2").Left + 65 + objForm.Items.Item("2").Height - 2
            oItem.Width = objForm.Items.Item("2").Width
            oItem.Height = objForm.Items.Item("2").Height
            ' oItem.LinkTo = "2"
            Dim btn As SAPbouiCOM.Button = objForm.Items.Item("Btn_UpdCst").Specific
            btn.Caption = "Update Cost"

            objMain.objUtilities.AddLabel(objForm.UniqueID, "U_SGI", objForm.Items.Item("1320000091").Top + 20, objForm.Items.Item("1320000091").Left - 10, objForm.Items.Item("1320000091").Width + 10, "Goods Issue No", "1320000091", objForm.Items.Item("1320000091").Height, 0, 0)
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "U_GI", objForm.Items.Item("1320000092").Top + 20, objForm.Items.Item("1320000092").Left + 30, objForm.Items.Item("1320000092").Width - 30, "OIGN", "U_GI", "1320000092", objForm.Items.Item("1320000092").Height, 0, 0)


            objForm.Items.Item("U_GI").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("U_GI").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("U_MIDNUM").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("U_MIDNUM").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("U_MIDENT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("U_MIDENT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objMain.objUtilities.AddButton(objForm.UniqueID, "btn_ED", objForm.Items.Item("2").Top,
                                      objForm.Items.Item("Btn_UpdCst").Width + objForm.Items.Item("Btn_UpdCst").Left + 5, "100", objForm.Items.Item("2").Height, "2", "Expiry Date", objForm.Items.Item("2").FromPane, objForm.Items.Item("2").ToPane)

            objMain.objUtilities.AddEditBox(objForm.UniqueID, "txt_GUID", objForm.Items.Item("4").Top, objForm.Items.Item("4").Left + objForm.Items.Item("4").Width + 5, objForm.Items.Item("4").Width, "OIGN", "U_GUID", "", objForm.Items.Item("4").Height, 96, 96)

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
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "PDTUSER", objForm.Items.Item("PDTUsr").Top, objForm.Items.Item("PDTUsr").Left + 120, objForm.Items.Item("PDTUsr").Width + 5, "OIGN", "U_PDTUSER", "PDTUsr", objForm.Items.Item("PDTUsr").Height, 46, 46)

            objMain.objUtilities.AddLabel(objForm.UniqueID, "UserNm", objForm.Items.Item("20").Top + 120, objForm.Items.Item("20").Left + 5, objForm.Items.Item("20").Width + 13, "User Name", "PDTUSER", objForm.Items.Item("20").Height, 46, 46)
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "UserName", objForm.Items.Item("UserNm").Top, objForm.Items.Item("UserNm").Left + 120, objForm.Items.Item("PDTUsr").Width + 5, "OIGN", "U_UserName", "UserNm", objForm.Items.Item("UserNm").Height, 46, 46)

            objMain.objUtilities.AddLabel(objForm.UniqueID, "Trty", objForm.Items.Item("20").Top + 140, objForm.Items.Item("20").Left + 5, objForm.Items.Item("20").Width + 13, "Transaction Type", "PDTUSER", objForm.Items.Item("20").Height, 46, 46)
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "TRTYPE", objForm.Items.Item("Trty").Top, objForm.Items.Item("Trty").Left + 120, objForm.Items.Item("PDTUsr").Width + 5, "OIGN", "U_TRTYPE", "Trty", objForm.Items.Item("Trty").Height, 46, 46)

            objMain.objUtilities.AddLabel(objForm.UniqueID, "SRN", objForm.Items.Item("20").Top + 160, objForm.Items.Item("20").Left + 5, objForm.Items.Item("20").Width + 20, "Stock Request Number", "PDTUSER", objForm.Items.Item("20").Height, 46, 46)
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "INVTRREQ", objForm.Items.Item("SRN").Top, objForm.Items.Item("SRN").Left + 120, objForm.Items.Item("PDTUsr").Width + 5, "OIGN", "U_INVTRREQ", "SRN", objForm.Items.Item("SRN").Height, 46, 46)

            objMain.objUtilities.AddLabel(objForm.UniqueID, "SRE", objForm.Items.Item("20").Top + 180, objForm.Items.Item("20").Left + 5, objForm.Items.Item("20").Width + 13, "Stock Request Entry", "PDTUSER", objForm.Items.Item("20").Height, 46, 46)
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "INVTRREN", objForm.Items.Item("SRE").Top, objForm.Items.Item("SRE").Left + 120, objForm.Items.Item("PDTUsr").Width + 5, "OIGN", "U_INVTRREN", "SRE", objForm.Items.Item("SRE").Height, 46, 46)





        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("721", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        CostBtnFlag = False

                        If CStr(oDBs_Head.GetValue("U_GUID", 0)).Trim <> "" Then

                            Dim UpdateDa As String = "Update ""@TNXEDS"" SET ""U_BENTRY""='" & oDBs_Head.GetValue("DocEntry", 0) & "', " &
                                                  """U_BNO""='" & oDBs_Head.GetValue("DocNum", 0) & "' Where ""U_GUID""='" & oDBs_Head.GetValue("U_GUID", 0) & "'  And ""U_DOCTY""='OIGN'"

                            Dim oRsUpdateDa As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            oRsUpdateDa.DoQuery(UpdateDa)

                        End If
                    End If

            End Select
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
            If oDBs_Head.GetValue("U_Type", 0) = "STR" Then
                'If objForm.Items.Item("U_TYPE").Specific.Value = "STR" Then
                If CostBtnFlag = False Then
                    objMain.objApplication.StatusBar.SetText("Please run Update Cost  before adding the document ......!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If
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
                                                         """U_GUID""='" & objForm.Items.Item("txt_GUID").Specific.Value & "' And ""U_DOCTY""='OIGN'"

            Dim oRsCheckBackDOc As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oRsCheckBackDOc.DoQuery(CheckBackDOc)

            Dim GetWCnd As String = "select IFNULL(""U_EXCHK"",'No') FROM OWHS Where ""WhsCode""='" & objMatrix.Columns.Item("15").Cells.Item(1).Specific.Value & "' And ""U_EXCHK""='Yes' And ""U_EXCHK"" IS NOT NULL And ""U_EXCHK""<>''"
            Dim oRsGetWCnd As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetWCnd.DoQuery(GetWCnd)

            If oRsGetWCnd.RecordCount > 0 And (oRsCheckBackDOc.RecordCount = 0 Or CStr(CStr(objForm.Items.Item("txt_GUID").Specific.Value).Trim).Trim = "") Then
                objMain.objApplication.StatusBar.SetText("please add Expiry Screen before adding")
                objForm.Freeze(False)
                Return False
                Exit Try
            End If

            Return True
        Catch ex As Exception
        End Try
    End Function
End Class
