Public Class ClsAPCreditMemo

    Dim objForm As SAPbouiCOM.Form
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oItem As SAPbouiCOM.Item

    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("ORPC")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("RPC1")
                    objMatrix = objForm.Items.Item("38").Specific

                    If pVal.BeforeAction = True Then
                        Me.AddItems(objForm.UniqueID)
                    End If


                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("ORPC")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("RPC1")
                    objMatrix = objForm.Items.Item("38").Specific

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Try


                            Dim oChk As SAPbouiCOM.CheckBox = objMatrix.Columns.Item("1250002121").Cells.Item(1).Specific

                            'If oChk.Checked = False And (oRsCheckBackDOc.RecordCount = 0 Or CStr(CStr(objForm.Items.Item("txt_GUID").Specific.Value).Trim).Trim = "") Then
                            '    objMain.objApplication.StatusBar.SetText("please add Expiry Screen before adding")
                            '    objForm.Freeze(False)
                            '    BubbleEvent = False

                            '    Exit Try
                            'End If

                            If oChk.Checked = True Then
                                Exit Try
                            End If


                            If CStr(objForm.Items.Item("3").Specific.Value).Trim <> "I" Then
                                Exit Try
                            End If

                            Dim GetWCnd As String = "select IFNULL(""U_EXCHK"",'No') FROM OWHS Where ""WhsCode""='" & objMatrix.Columns.Item("24").Cells.Item(1).Specific.Value & "' And ""U_EXCHK""='Yes' And ""U_EXCHK"" IS NOT NULL And ""U_EXCHK""<>''"
                            Dim oRsGetWCnd As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetWCnd.DoQuery(GetWCnd)


                            Dim CheckBackDOc As String = "Select ""U_GUID"" FROM ""@TNXEDS"" Where " &
                                                         """U_GUID""='" & objForm.Items.Item("txt_GUID").Specific.Value & "' And ""U_DOCTY""='ORPC'"

                            Dim oRsCheckBackDOc As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsCheckBackDOc.DoQuery(CheckBackDOc)

                            Dim BaseType As Double = objMatrix.Columns.Item("43").Cells.Item(1).Specific.Value

                            If BaseType = -1 And BaseType <> 21 And oRsGetWCnd.RecordCount > 0 And (oRsCheckBackDOc.RecordCount = 0 Or CStr(CStr(objForm.Items.Item("txt_GUID").Specific.Value).Trim).Trim = "") Then
                                objMain.objApplication.StatusBar.SetText("please add Expiry Screen before adding")
                                objForm.Freeze(False)
                                BubbleEvent = False
                                Exit Try
                            End If


                            If BaseType = -1 And BaseType <> 21 And oRsGetWCnd.RecordCount > 0 And (oRsCheckBackDOc.RecordCount = 0 Or CStr(CStr(objForm.Items.Item("txt_GUID").Specific.Value).Trim).Trim = "") Then
                                objMain.objApplication.StatusBar.SetText("please add Expiry Screen before adding")
                                objForm.Freeze(False)
                                BubbleEvent = False

                                Exit Try
                            End If




                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If

                    If pVal.ItemUID = "PSTEXD" And pVal.BeforeAction = False Then
                        If pVal.BeforeAction = False And objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("ORPC")
                            Dim JEN As String = objForm.Items.Item("EXCJVN1").Specific.Value
                            If String.IsNullOrEmpty(JEN) Then
                                Dim qStrBranchEbl As String = "SELECT ""CompnyName"" FROM OADM WHERE ""MltpBrnchs""='Y'"
                                Dim rSetBranchEbl As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                rSetBranchEbl.DoQuery(qStrBranchEbl)
                                If rSetBranchEbl.RecordCount > 0 Then
                                    Me.PostJEwithBranch(objForm.UniqueID)
                                Else
                                    Me.PostJE(objForm.UniqueID)
                                End If
                            Else
                                objMain.objApplication.StatusBar.SetText("JE Already Posted.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                            objMain.objApplication.ActivateMenuItem("1304")
                        End If
                    End If

                    If pVal.BeforeAction = False And pVal.ItemUID = "lin4" Then
                        Dim objForm1 As SAPbouiCOM.Form
                        Try
                            'Dim BatchNum As String = objForm.Items.Item("EXCJVN1").Specific.Value
                            'Dim CostCenterQry As String = "Select ""TransId"" From OJDT Where ""BatchNum""='" & BatchNum & "'"
                            'Dim oRsCostCenter As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'oRsCostCenter.DoQuery(CostCenterQry)
                            'Dim TransId As String = oRsCostCenter.Fields.Item("TransId").Value
                            Dim TransId As String = objForm.Items.Item("EXCJVN1").Specific.Value
                            objMain.objApplication.ActivateMenuItem("1540")
                            objForm1 = objMain.objApplication.Forms.ActiveForm
                            'objMain.objApplication.Forms.GetForm("392", objMain.objApplication.Forms.ActiveForm.TypeCount)
                            objForm1.Freeze(True)
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            objForm1.Items.Item("5").Specific.Value = TransId
                            objForm1.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                            objForm1.Freeze(False)
                        Catch ex As Exception
                        End Try
                    End If
                    If pVal.ItemUID = "btn_ED" And pVal.BeforeAction = False Then
                        Try
                            objForm.Freeze(True)


                            'Dim GetData As String = "SELECT T1.""LineNum"",T1.""ItemCode"", T1.""Dscription"", T1.""Quantity"", T1.""CodeBars"", " &
                            '                        "T1.""UomCode"" FROM ORPC T0  INNER JOIN RPC1 T1 ON T0.""DocEntry"" = T1.""DocEntry""  " &
                            '                        " INNER JOIN OWHS T2 ON T1.""WhsCode"" = T2.""WhsCode"" WHERE  " &
                            '                        "T0.""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0) & "' And IFNULL(T2.""U_EXCHK"",'No') ='Yes' " &
                            '                        "And T0.""DocType""='I'  And T1.""NoInvtryMv""='N'  And T1.""BaseType"" ='-1' and T1.""BaseType""<>'21' ORDER BY T1.""LineNum"""

                            'Dim oRsGetData As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            'oRsGetData.DoQuery(GetData)

                            'If oRsGetData.RecordCount > 0 Then
                            '    objMain.objExpiry.CreateForm(objForm.UniqueID, "ORPC", oDBs_Head.GetValue("DocEntry", 0), oDBs_Head.GetValue("DocNum", 0), oRsGetData)
                            'Else
                            '    objMain.objApplication.StatusBar.SetText("Expiry Date screen not applicable")
                            'End If

                            Dim BaseType As Double = objMatrix.Columns.Item("43").Cells.Item(1).Specific.Value

                            If BaseType <> -1 And BaseType <> 21 Then
                                objMain.objApplication.StatusBar.SetText("Expiry Screen not applicable")
                                objForm.Freeze(False)
                                Exit Try
                            End If

                            Dim oChk As SAPbouiCOM.CheckBox = objMatrix.Columns.Item("1250002121").Cells.Item(1).Specific

                            If oChk.Checked = True Then
                                objMain.objApplication.StatusBar.SetText("Expiry Screen not applicable")
                                objForm.Freeze(False)
                                Exit Try
                            End If

                            If CStr(objForm.Items.Item("3").Specific.Value).Trim <> "I" Then
                                objMain.objApplication.StatusBar.SetText("Expiry Screen not applicable")
                                objForm.Freeze(False)
                                Exit Try
                            End If

                            Dim GetWCnd As String = "select IFNULL(""U_EXCHK"",'No') FROM OWHS Where ""WhsCode""='" & objMatrix.Columns.Item("24").Cells.Item(1).Specific.Value & "' And ""U_EXCHK""='Yes' And ""U_EXCHK"" IS NOT NULL And ""U_EXCHK""<>''"
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
                                    dr("Dscription") = objMatrix.Columns.Item("3").Cells.Item(i).Specific.Value
                                    dr("Quantity") = objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value
                                    dr("CodeBars") = objMatrix.Columns.Item("4").Cells.Item(i).Specific.Value
                                    dr("UomCode") = objMatrix.Columns.Item("1470002145").Cells.Item(i).Specific.Value
                                    dr("BQuantity") = objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value
                                    dt.Rows.Add(dr)

                                End If


                            Next

GotoLst:                    objMain.objExpiry.CreateFormNew(objForm.UniqueID, "ORPC", oDBs_Head.GetValue("DocEntry", 0), oDBs_Head.GetValue("DocNum", 0), dt, UID.Trim, objForm.Mode)


                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("ORPC")
                    objMatrix = objForm.Items.Item("38").Specific
                    If pVal.ItemUID = "38" And pVal.ColUID = "1" And pVal.BeforeAction = False Then
                        For i As Integer = 1 To objMatrix.VisualRowCount - 1
                            If objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value <> "" Then
                                objMatrix = objForm.Items.Item("38").Specific
                                Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value
                                Dim GetDocEntry As String = "Select ""U_EXCVAT"" From OITM Where ""ItemCode"" = '" & ItemCode & "' "
                                Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRs.DoQuery(GetDocEntry)
                                objMatrix.Columns.Item("U_EXCTAX").Cells.Item(i).Specific.Value() = oRs.Fields.Item("U_EXCVAT").Value()
                            End If
                        Next
                    End If


            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub AddItems(ByVal FormUID As String)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        Try

            objMain.objUtilities.AddButton(objForm.UniqueID, "btn_ED", objForm.Items.Item("2").Top,
                                    objForm.Items.Item("2").Width + objForm.Items.Item("2").Left + 5, "100", objForm.Items.Item("2").Height, "2", "Expiry Date", objForm.Items.Item("2").FromPane, objForm.Items.Item("2").ToPane)


            objMain.objUtilities.AddEditBox(objForm.UniqueID, "txt_GUID", objForm.Items.Item("4").Top, objForm.Items.Item("4").Left + objForm.Items.Item("4").Width + 5, objForm.Items.Item("4").Width, "ORPC", "U_GUID", "", objForm.Items.Item("4").Height, 96, 96)

            objForm.Items.Item("txt_GUID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)


            objForm = objMain.objApplication.Forms.Item(FormUID)
            oItem = objForm.Items.Add("PSTEXD", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 10
            oItem.Left = objForm.Items.Item("38").Left + 100
            oItem.Width = objForm.Items.Item("2").Width + 30
            oItem.Height = objForm.Items.Item("2").Height
            oItem.FromPane = objForm.Items.Item("38").FromPane
            oItem.ToPane = objForm.Items.Item("38").ToPane
            Dim btn1 As SAPbouiCOM.Button = objForm.Items.Item("PSTEXD").Specific
            btn1.Caption = "Post Excise Duty"

            objForm.Items.Item("PSTEXD").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("PSTEXD").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objMain.objUtilities.AddLabel(objForm.UniqueID, "EXCJVNDP1", objForm.Items.Item("34").Top + 15, objForm.Items.Item("34").Left + 0, objForm.Items.Item("34").Width, "Excise JE Number", "EXCJVN1", objForm.Items.Item("34").Height) 'objform.Items.Item("QC").FromPane, objform.Items.Item("QC").ToPane)
            objMain.objUtilities.AddEditBox(FormUID, "EXCJVN1", objForm.Items.Item("33").Top + 15, objForm.Items.Item("33").Left + 0, objForm.Items.Item("33").Width, "ORPC", "U_EXCJVN1", "EXCJVNDP1", objForm.Items.Item("33").Height)

            objForm.Items.Item("EXCJVN1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm = objMain.objApplication.Forms.Item(FormUID)
            oItem = objForm.Items.Add("lin4", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oItem.Top = objForm.Items.Item("EXCJVN1").Top + 0
            oItem.Left = objForm.Items.Item("EXCJVN1").Left - 20
            Dim LinkBtn As SAPbouiCOM.LinkedButton
            LinkBtn = oItem.Specific
            oItem.LinkTo = "EXCJVN1"
            LinkBtn.LinkedObjectType = "-1"

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("181", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then


                        If CStr(oDBs_Head.GetValue("U_GUID", 0)).Trim <> "" Then

                            Dim UpdateDa As String = "Update ""@TNXEDS"" SET ""U_BENTRY""='" & oDBs_Head.GetValue("DocEntry", 0) & "', " &
                                                  """U_BNO""='" & oDBs_Head.GetValue("DocNum", 0) & "' Where ""U_GUID""='" & oDBs_Head.GetValue("U_GUID", 0) & "'  And ""U_DOCTY""='ORPC'"

                            Dim oRsUpdateDa As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            oRsUpdateDa.DoQuery(UpdateDa)

                        End If


                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("ORPC")
                    Dim JEN As String = objForm.Items.Item("EXCJVN1").Specific.Value
                    ' If String.IsNullOrEmpty(JEN) Then
                    If JEN <> "" Then
                        ' If EXCJVN1 is empty or null, enable PSTEXD
                        objForm.Items.Item("PSTEXD").Enabled = False
                    Else
                        ' If EXCJVN1 has a value, disable PSTEXD
                        objForm.Items.Item("PSTEXD").Enabled = True
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Private Sub PostJE(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("ORPC")
            ' Dim Branch As SAPbouiCOM.ComboBox = objForm.Items.Item("2001").Specific

            'If Branch.Value = "" Then
            '    objMain.objApplication.StatusBar.SetText("Please select the branch.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Exit Sub
            'End If
            Dim oJE As SAPbobsCOM.JournalEntries = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

            Dim ApCrdMemoSrsNum As String = oDBs_Head.GetValue("DocNum", 0).Trim

            Dim PurchaseexcisedutyCrGL, PurchaseexcisedutyDrGL As String
            Dim GL As String = "Select ""U_PEDCG"",""U_PEDDRG"",""U_SEDDG"" from ""OADM"" "
            Dim oRsGL As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGL.DoQuery(GL)
            If oRsGL.RecordCount > 0 Then
                PurchaseexcisedutyCrGL = oRsGL.Fields.Item("U_PEDCG").Value
                PurchaseexcisedutyDrGL = oRsGL.Fields.Item("U_PEDDRG").Value

                If PurchaseexcisedutyCrGL = "" Then
                    objMain.objApplication.StatusBar.SetText("Please Define Purchase exciseduty Cr GL", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                ElseIf PurchaseexcisedutyDrGL = "" Then
                    objMain.objApplication.StatusBar.SetText("Please Define Purchase exciseduty Dr GL", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
            End If

            Dim sDate As String = DateTime.Now.ToString("yyyyMMdd")
            sDate = sDate.Insert("4", "-")
            sDate = sDate.Insert("7", "-")

            '---------------------------------------------------------------------------------------------
            Dim totalTaxValue As Double = 0
            For i As Integer = 1 To objMatrix.VisualRowCount
                If objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value <> "" Then
                    '      oDs1 = objForm.DataSources.DBDataSources.Item("RCP1")
                    objMatrix = objForm.Items.Item("38").Specific
                    Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.ToString().Trim()
                    Dim QTY As Double = objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value.ToString().Trim()
                    Dim EXCTAX As Double = objMatrix.Columns.Item("U_EXCTAX").Cells.Item(i).Specific.Value.ToString().Trim()
                    Dim SubTotal As String = "ST"
                    Dim ItemPer As Double = objMatrix.Columns.Item("213").Cells.Item(i).Specific.Value.ToString().Trim()

                    'Dim SubTotal As Double = "ST1"
                    If (QTY * EXCTAX * ItemPer) <> 0 Then
                        Dim ST As Double = QTY * EXCTAX * ItemPer
                        totalTaxValue += ST
                    End If
                End If
            Next
            '--------------------------------------------------------------------------------------------------------------

            oJE.ReferenceDate = sDate
            oJE.DueDate = sDate
            oJE.TaxDate = sDate

            oJE.Reference2 = "A/P Credit Memo"

            ' oJE.Lines.BPLID = Branch.Value
            oJE.Lines.AccountCode = PurchaseexcisedutyDrGL
            oJE.Lines.Debit = totalTaxValue
            oJE.Memo = "A/P Credit Mome Posted :" & ApCrdMemoSrsNum
            oJE.Lines.Add()

            ' oJE.Lines.BPLID = Branch.Value
            oJE.Lines.AccountCode = PurchaseexcisedutyCrGL
            oJE.Lines.Credit = totalTaxValue
            oJE.Memo = "A/P Credit Mome Posted :" & ApCrdMemoSrsNum
            oJE.Lines.Add()

            If oJE.Add() = 0 Then
                Dim NewTransID As String = objMain.objCompany.GetNewObjectKey.ToString()

                Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim Update As String = " UPDATE ""ORPC"" T1 SET T1.""U_EXCJVN1""='" + NewTransID + "' WHERE T1.""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0) & "'"
                oRsUpdate.DoQuery(Update)
                objMain.objApplication.StatusBar.SetText("JE Posted Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                'objMain.objApplication.ActivateMenuItem("1304")
            Else
                objMain.objApplication.StatusBar.SetText("JE Posting failed" + objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub PostJEwithBranch(ByRef FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("ORPC")

            Dim oJE As SAPbobsCOM.JournalEntries = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

            Dim ApCrdMemoSrsNum As String = oDBs_Head.GetValue("DocNum", 0).Trim

            Dim PurchaseexcisedutyCrGL, PurchaseexcisedutyDrGL As String
            Dim GL As String = "Select ""U_PEDCG"",""U_PEDDRG"",""U_SEDDG"" from ""OADM"" "
            Dim oRsGL As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGL.DoQuery(GL)
            If oRsGL.RecordCount > 0 Then
                PurchaseexcisedutyCrGL = oRsGL.Fields.Item("U_PEDCG").Value
                PurchaseexcisedutyDrGL = oRsGL.Fields.Item("U_PEDDRG").Value

                If PurchaseexcisedutyCrGL = "" Then
                    objMain.objApplication.StatusBar.SetText("Please Define Purchase exciseduty Cr GL", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                ElseIf PurchaseexcisedutyDrGL = "" Then
                    objMain.objApplication.StatusBar.SetText("Please Define Purchase exciseduty Dr GL", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
            End If
            Dim sBrachCode As String = objForm.Items.Item("2001").Specific.Value.ToString
            Dim sDate As String = objForm.Items.Item("10").Specific.Value.ToString
            sDate = sDate.Insert("4", "-")
            sDate = sDate.Insert("7", "-")

            '---------------------------------------------------------------------------------------------
            Dim totalTaxValue1 As Double = 0
            For i As Integer = 1 To objMatrix.VisualRowCount
                If objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value <> "" Then
                    '      oDs1 = objForm.DataSources.DBDataSources.Item("RCP1")
                    objMatrix = objForm.Items.Item("38").Specific
                    Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.ToString().Trim()
                    Dim QTY As Double = objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value.ToString().Trim()
                    Dim EXCTAX As Double = objMatrix.Columns.Item("U_EXCTAX").Cells.Item(i).Specific.Value.ToString().Trim()
                    Dim ItemPer As Double = objMatrix.Columns.Item("213").Cells.Item(i).Specific.Value.ToString().Trim()

                    'Dim SubTotal As Double = "ST1"
                    If (QTY * EXCTAX * ItemPer) <> 0 Then
                        Dim ST1 As Double = QTY * EXCTAX * ItemPer
                        totalTaxValue1 += ST1
                    End If
                End If
            Next
            '--------------------------------------------------------------------------------------------------------------

            oJE.ReferenceDate = sDate
            oJE.DueDate = sDate
            oJE.TaxDate = sDate

            oJE.Reference2 = "A/P Credit Memo"

            '  oJE.Lines.BPLID = Branch.Value
            oJE.Lines.AccountCode = PurchaseexcisedutyCrGL
            oJE.Lines.Debit = totalTaxValue1
            oJE.Memo = "A/P Credit Memo Posted :" & ApCrdMemoSrsNum
            oJE.Lines.BPLID = CInt(sBrachCode)
            oJE.Lines.Add()
            '  oJE.Lines.BPLID = Branch.Value
            oJE.Lines.AccountCode = PurchaseexcisedutyDrGL
            oJE.Lines.Credit = totalTaxValue1
            oJE.Memo = "A/P Credit Memo Posted :" & ApCrdMemoSrsNum
            oJE.Lines.BPLID = CInt(sBrachCode)
            oJE.Lines.Add()

            If oJE.Add() = 0 Then
                Dim NewTransID As String = objMain.objCompany.GetNewObjectKey.ToString()

                Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim Update As String = "UPDATE ""ORPC"" T1 SET T1.""U_EXCJVN1""='" + NewTransID + "' WHERE T1.""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0) & "'"
                oRsUpdate.DoQuery(Update)
                objMain.objApplication.StatusBar.SetText("JE Posted Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objMain.objApplication.StatusBar.SetText("JE Posting failed" + objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try
    End Sub
End Class
