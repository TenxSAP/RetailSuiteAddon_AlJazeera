Public Class PurchaseInvoice

#Region "        Declaration        "
    Dim objForm, AltForm, objMIForm As SAPbouiCOM.Form
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim SplitCombo As SAPbouiCOM.ComboBox
    Dim SplitFlgCombo, objComboBox As SAPbouiCOM.ComboBox
    Dim OrsSTRQRY, rs As SAPbobsCOM.Recordset
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim strQry As String = ""
    Dim oItem, oItem1, oItem2, oItem3 As SAPbouiCOM.Item
    Dim oColumn As SAPbouiCOM.Column
#End Region

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix = objForm.Items.Item("38").Specific

                    If pVal.BeforeAction = True Then
                        Me.AddItems(objForm.UniqueID)
                    Else
                        Try
                            objMatrix = objForm.Items.Item("39").Specific
                            objComboBox = objMatrix.Columns.Item("U_Branch").Cells.Item(1).Specific  ''objForm.Items.Item("14").Specific
                            oColumn = objMatrix.Columns.Item("U_Branch")
                            Dim sQury As String = " Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join " &
                                                  " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""  "
                            objMain.objUtilities.MatrixComboBoxValues1(oColumn, sQury)
                            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                        Catch ex As Exception

                        End Try
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "1" And pVal.Before_Action = True Then
                        objForm = objMain.objApplication.Forms.Item(FormUID)
                        objMatrix = objForm.Items.Item("38").Specific
                        For i As Integer = 1 To objMatrix.RowCount
                            SplitCombo = objMatrix.Columns.Item("U_SplitKitItems").Cells.Item(i).Specific
                            SplitFlgCombo = objMatrix.Columns.Item("U_SPLITTED").Cells.Item(i).Specific
                            If SplitCombo.Selected.Value = SplitFlgCombo.Selected.Value Then

                            Else
                                objMain.objApplication.StatusBar.SetText("Validation : Split Item Not Splitted  for Line ID : " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                            End If
                        Next
                    End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("OPCH")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("PCH1")

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Try

                            'Dim CheckBackDOc As String = "Select ""U_GUID"" FROM ""@TNXEDS"" Where " &
                            '                             """U_GUID""='" & objForm.Items.Item("txt_GUID").Specific.Value & "' And ""U_DOCTY""='OPCH'"

                            'Dim oRsCheckBackDOc As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            'oRsCheckBackDOc.DoQuery(CheckBackDOc)

                            'Dim GetWCnd As String = "select IFNULL(""U_EXCHK"",'No') FROM OWHS Where ""WhsCode""='" & objMatrix.Columns.Item("24").Cells.Item(1).Specific.Value & "' And ""U_EXCHK""='Yes' And ""U_EXCHK"" IS NOT NULL And ""U_EXCHK""<>''"
                            'Dim oRsGetWCnd As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'oRsGetWCnd.DoQuery(GetWCnd)


                            'If CStr(objForm.Items.Item("3").Specific.Value).Trim = "I" And oRsGetWCnd.RecordCount > 0 And (oRsCheckBackDOc.RecordCount = 0 Or CStr(CStr(objForm.Items.Item("txt_GUID").Specific.Value).Trim).Trim = "") Then
                            '    objMain.objApplication.StatusBar.SetText("please add Expiry Screen before adding")
                            '    objForm.Freeze(False)
                            '    BubbleEvent = False
                            '    Exit Try
                            'End If


                            'If oRsGetWCnd.RecordCount > 0 And (oRsCheckBackDOc.RecordCount = 0 Or CStr(CStr(objForm.Items.Item("txt_GUID").Specific.Value).Trim).Trim = "") Then
                            '    objMain.objApplication.StatusBar.SetText("please add Expiry Screen before adding")
                            '    objForm.Freeze(False)
                            '    BubbleEvent = False
                            '    Exit Try
                            'End If



                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If

                    If pVal.ItemUID = "btn_ED" And pVal.BeforeAction = False Then
                        Try
                            objForm.Freeze(True)


                            'Dim GetData As String = "SELECT T1.""LineNum"",T1.""ItemCode"", T1.""Dscription"", T1.""Quantity"", T1.""CodeBars"", " &
                            '                        "T1.""UomCode"" FROM OPCH T0  INNER JOIN PCH1 T1 ON T0.""DocEntry"" = T1.""DocEntry""  " &
                            '                        " INNER JOIN OWHS T2 ON T1.""WhsCode"" = T2.""WhsCode"" WHERE  " &
                            '                        "T0.""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0) & "' And IFNULL(T2.""U_EXCHK"",'No') ='Yes' And T0.""DocType""='I' ORDER BY T1.""LineNum"""

                            'Dim oRsGetData As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            'oRsGetData.DoQuery(GetData)

                            'If oRsGetData.RecordCount > 0 Then
                            '    objMain.objExpiry.CreateForm(objForm.UniqueID, "OPCH", oDBs_Head.GetValue("DocEntry", 0), oDBs_Head.GetValue("DocNum", 0), oRsGetData)
                            'Else
                            '    objMain.objApplication.StatusBar.SetText("Expiry Date screen not applicable")
                            'End If


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

GotoLst:                    objMain.objExpiry.CreateFormNew(objForm.UniqueID, "OPCH", oDBs_Head.GetValue("DocEntry", 0), oDBs_Head.GetValue("DocNum", 0), dt, UID.Trim, objForm.Mode)


                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If

                    If pVal.ItemUID = "PEXD" And pVal.BeforeAction = False Then
                        If pVal.BeforeAction = False And objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("OPCH")
                            Dim JEN As String = objForm.Items.Item("EXCJVN").Specific.Value
                            If String.IsNullOrEmpty(JEN) Then
                                Dim qStrBranchEbl As String = "SELECT ""CompnyName"" FROM OADM WHERE ""MltpBrnchs""='Y'"
                                Dim rSetBranchEbl As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                rSetBranchEbl.DoQuery(qStrBranchEbl)
                                If rSetBranchEbl.RecordCount > 0 Then
                                    Me.PostJEwithBranch(objForm.UniqueID)
                                Else
                                    Me.PostJE1(objForm.UniqueID)
                                End If
                                '  objMain.objApplication.StatusBar.SetText("JV Posting Successfully.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Else
                                objMain.objApplication.StatusBar.SetText("JV Already Posted.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                            objMain.objApplication.ActivateMenuItem("1304")
                        End If
                    End If

                    If pVal.BeforeAction = False And pVal.ItemUID = "lin3" Then
                        Dim objForm1 As SAPbouiCOM.Form
                        Try
                            'Dim BatchNum As String = objForm.Items.Item("EXCJVN1").Specific.Value
                            'Dim CostCenterQry As String = "Select ""TransId"" From OJDT Where ""BatchNum""='" & BatchNum & "'"
                            'Dim oRsCostCenter As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'oRsCostCenter.DoQuery(CostCenterQry)
                            'Dim TransId As String = oRsCostCenter.Fields.Item("TransId").Value
                            Dim TransId As String = objForm.Items.Item("EXCJVN").Specific.Value
                            objMain.objApplication.ActivateMenuItem("1540")
                            objForm1 = objMain.objApplication.Forms.ActiveForm
                            'objMain.objApplication.Forms.GetForm("392", objMain.objApplication.Forms.ActiveForm.TypeCount)
                            'Test Push kahkh
                            objForm1.Freeze(True)
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            objForm1.Items.Item("5").Specific.Value = TransId
                            objForm1.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                            objForm1.Freeze(False)
                        Catch ex As Exception
                        End Try
                    End If

                    If pVal.ItemUID = "btn_Supp" And pVal.BeforeAction = False Then
                        objMatrix = objForm.Items.Item("38").Specific
                        objForm.Items.Item("42").Specific.Value = ""
                        objForm.Items.Item("24").Specific.Value = ""

                        For i As Integer = 1 To objMatrix.VisualRowCount - 1
                            Dim str As String = objMatrix.Columns.Item("1980002193").Cells.Item(i).Specific.Value
                            Dim strPrice As String = objForm.Items.Item("22").Specific.Value
                            Dim amount As Double = CDbl(strPrice.Substring(3))
                            If str <> "" Then
                                strQry = "SELECT T0.""U_INVTLDSC""as ""Percentage"",T0.""U_ITAmt"" as ""Amount"" FROM ""@SBO_SUPREBATE""  T0  inner join OOAT T1 on  T0.""U_BASENO"" = T1.""Number"" WHERE T1.""Number"" = '" & str & "'"

                                OrsSTRQRY = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                OrsSTRQRY.DoQuery(strQry)
                                If OrsSTRQRY.RecordCount > 0 Then
                                    If amount <> 0 Then
                                        Dim purchase As Double = OrsSTRQRY.Fields.Item("Amount").Value
                                        If purchase > 0 Then
                                            objForm.Items.Item("24").Specific.Value = purchase
                                        End If

                                        Dim purchasePer As Double = OrsSTRQRY.Fields.Item("Percentage").Value
                                        If purchasePer > 0 Then
                                            objForm.Items.Item("42").Specific.Value = purchasePer
                                        End If
                                    End If
                                End If
                            End If
                        Next

                    End If

                    If pVal.ItemUID = "BAlloc_JE" And pVal.BeforeAction = False And objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        objForm = objMain.objApplication.Forms.Item(FormUID)
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("OPCH")
                        oDBs_Details = objForm.DataSources.DBDataSources.Item("PCH1")
                        Me.PostJE()
                    End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("OPCH")
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

            oItem3 = objForm.Items.Add("btn_Supp", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem3.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 10
            oItem3.Left = objForm.Items.Item("38").Left
            oItem3.Width = objForm.Items.Item("2").Width + 30
            oItem3.Height = objForm.Items.Item("2").Height
            oItem3.FromPane = objForm.Items.Item("38").FromPane
            oItem3.ToPane = objForm.Items.Item("38").ToPane
            Dim btn3 As SAPbouiCOM.Button = objForm.Items.Item("btn_Supp").Specific
            btn3.Caption = "Supplier &Rebate"

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
        Try
            Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("39").Specific
            oItem = objForm.Items.Add("BAlloc_JE", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Top = objForm.Items.Item("2").Top
            oItem.Left = objForm.Items.Item("2").Left + 65 + objForm.Items.Item("2").Height - 2
            oItem.Width = objForm.Items.Item("2").Width + 50
            oItem.Height = objForm.Items.Item("2").Height
            Dim btn As SAPbouiCOM.Button = objForm.Items.Item("BAlloc_JE").Specific
            btn.Caption = "BranchAllocation JE"
            objForm.Items.Item("BAlloc_JE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("BAlloc_JE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            'objMain.objUtilities.AddButton(objForm.UniqueID, "btn_ED", objForm.Items.Item("2").Top,
            '                        objForm.Items.Item("BAlloc_JE").Width + objForm.Items.Item("BAlloc_JE").Left + 5, "100", objForm.Items.Item("2").Height, "2", "Expiry Date", objForm.Items.Item("2").FromPane, objForm.Items.Item("2").ToPane)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objMain.objUtilities.AddLabel(objForm.UniqueID, "JENO", objForm.Items.Item("86").Top + 15, objForm.Items.Item("86").Left, objForm.Items.Item("86").Width - 25, "JE Number", "22", 0, 0)
            oItem = objForm.Items.Add("JEEntry", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = objForm.Items.Item("46").Top + 15
            oItem.Left = objForm.Items.Item("46").Left  '+ objform.Items.Item("4").Width + 325
            oItem.Width = objForm.Items.Item("46").Width '+ 35
            oItem.Height = objForm.Items.Item("46").Height '+ 100
            oItem.LinkTo = "JENO"

            Dim JE As SAPbouiCOM.EditText = objForm.Items.Item("JEEntry").Specific
            JE.DataBind.SetBound(True, "OPCH", "U_JEEntry")

            oItem = objForm.Items.Add("LinK2", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oItem.Top = objForm.Items.Item("86").Top + 17
            oItem.Left = objForm.Items.Item("46").Left - 18
            oItem.Width = 12
            oItem.Height = 10
            oItem.LinkTo = "JEEntry"
            Dim LinkBtn As SAPbouiCOM.LinkedButton
            LinkBtn = objForm.Items.Item("LinK2").Specific
            LinkBtn.LinkedObject = 30

            objForm.Items.Item("JEEntry").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objMain.objUtilities.AddEditBox(objForm.UniqueID, "txt_GUID", objForm.Items.Item("4").Top, objForm.Items.Item("4").Left + objForm.Items.Item("4").Width + 5, objForm.Items.Item("4").Width, "OPCH", "U_GUID", "", objForm.Items.Item("4").Height, 96, 96)

            objForm.Items.Item("txt_GUID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)


            objForm = objMain.objApplication.Forms.Item(FormUID)
            oItem = objForm.Items.Add("PEXD", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 10
            oItem.Left = objForm.Items.Item("38").Left + 100
            oItem.Width = objForm.Items.Item("2").Width + 30
            oItem.Height = objForm.Items.Item("2").Height
            oItem.FromPane = objForm.Items.Item("38").FromPane
            oItem.ToPane = objForm.Items.Item("38").ToPane
            Dim btn1 As SAPbouiCOM.Button = objForm.Items.Item("PEXD").Specific
            btn1.Caption = "Post Excise Duty"



            'oItem = objForm.Items.Add("PEXD", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'oItem.Top = objForm.Items.Item("2349990001").Top
            'oItem.Left = objForm.Items.Item("2349990001").Left + 180
            'oItem.Width = objForm.Items.Item("2349990001").Width
            'oItem.Height = objForm.Items.Item("2349990001").Height
            'oItem.FromPane = objForm.Items.Item("2349990001").FromPane
            'oItem.ToPane = objForm.Items.Item("2349990001").ToPane
            'Dim btn1 As SAPbouiCOM.Button = objForm.Items.Item("PEXD").Specific
            'btn1.Caption = "Post Excise Duty"

            objForm.Items.Item("PEXD").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("PEXD").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objMain.objUtilities.AddLabel(objForm.UniqueID, "EXCJVNDP", objForm.Items.Item("34").Top + 15, objForm.Items.Item("34").Left + 0, objForm.Items.Item("34").Width, "Excise JE Number", "EXCJVN", objForm.Items.Item("34").Height) 'objform.Items.Item("QC").FromPane, objform.Items.Item("QC").ToPane)
            objMain.objUtilities.AddEditBox(FormUID, "EXCJVN", objForm.Items.Item("33").Top + 15, objForm.Items.Item("33").Left + 0, objForm.Items.Item("33").Width, "OPCH", "U_EXCJVN", "EXCJVNDP", objForm.Items.Item("33").Height)

            objForm.Items.Item("EXCJVN").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm = objMain.objApplication.Forms.Item(FormUID)
            oItem = objForm.Items.Add("lin3", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oItem.Top = objForm.Items.Item("EXCJVN").Top + 0
            oItem.Left = objForm.Items.Item("EXCJVN").Left - 20
            Dim LinkBtn1 As SAPbouiCOM.LinkedButton
            LinkBtn1 = oItem.Specific
            oItem.LinkTo = "EXCJVN"
            LinkBtn1.LinkedObjectType = "-1"



            'objForm.Items.Item("JEEntry").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("JEEntry").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("JEEntry").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("141", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("OPCH")


                        Dim entry As String = oDBs_Head.GetValue("U_JEEntry", oDBs_Head.Offset).ToString().Trim
                        If entry <> "" Then
                            'objMatrix.Columns.Item("U_JEEntry").Enabled = False
                            objForm.Items.Item("BAlloc_JE").Enabled = False
                            'objform.Items.Item("BAlloc_JE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            'objform.Items.Item("BAlloc_JE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        Else
                            objForm.Items.Item("BAlloc_JE").Enabled = True
                        End If

                        Dim JEN As String = objForm.Items.Item("EXCJVN").Specific.Value
                        ' If String.IsNullOrEmpty(JEN) Then
                        If JEN <> "" Then
                            ' If EXCJVN1 is empty or null, enable PSTEXD
                            objForm.Items.Item("PEXD").Enabled = False
                        Else
                            ' If EXCJVN1 has a value, disable PSTEXD
                            objForm.Items.Item("PEXD").Enabled = True
                        End If
                    End If


                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then

                        If CStr(oDBs_Head.GetValue("U_GUID", 0)).Trim <> "" Then

                            Dim UpdateDa As String = "Update ""@TNXEDS"" SET ""U_BENTRY""='" & oDBs_Head.GetValue("DocEntry", 0) & "', " &
                                              """U_BNO""='" & oDBs_Head.GetValue("DocNum", 0) & "' Where ""U_GUID""='" & oDBs_Head.GetValue("U_GUID", 0) & "'  And ""U_DOCTY""='OPCH'"

                            Dim oRsUpdateDa As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            oRsUpdateDa.DoQuery(UpdateDa)

                        End If

                        objMain.objApplication.MessageBox("Purchase Invoice Document is Succesfully Added")
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub PostJE()
        oDBs_Head = objForm.DataSources.DBDataSources.Item("OPCH")
        oDBs_Details = objForm.DataSources.DBDataSources.Item("PCH1")
        objMatrix = objForm.Items.Item("39").Specific

        Dim objComboBoxservice As SAPbouiCOM.ComboBox = objForm.Items.Item("3").Specific
        If objComboBoxservice.Value <> "S" Then
            objMain.objApplication.StatusBar.SetText("Document Type must be Selected as Service for Branch Allocation JE Posting ......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If


        For i As Integer = 1 To objMatrix.VisualRowCount

            Dim objComboBoxBranch As SAPbouiCOM.ComboBox = objForm.Items.Item("2001").Specific
            Dim linebranch As String = objMatrix.Columns.Item("U_Branch").Cells.Item(i).Specific.Value.ToString().Trim
            If linebranch = "" Then
                objMain.objApplication.StatusBar.SetText("Row Level Branch must be fill in all row for Branch Allocation JE Posting...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

        Next
        Dim Dnum As String = oDBs_Head.GetValue("DocNum", 0).Trim
        Dim serie As String = oDBs_Head.GetValue("Series", 0).Trim
        Dim allrecordchk As String = "SELECT COUNT(T1.""U_Branch"") as ""Linechk""  FROM ""OPCH"" T0 INNER JOIN ""PCH1"" T1 " &
                            " ON T0.""DocEntry"" = T1.""DocEntry"" WHERE T1.""U_Branch"" = T1.""U_Branch"" and T0.""DocNum""='" & Dnum & "'"

        Dim rs1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs1.DoQuery(allrecordchk)
        '  If rs1.RecordCount > 0 Then
        Dim totalbranch As Integer = rs1.Fields.Item("Linechk").Value
        ' End If
        Dim counntchk As String = "SELECT COUNT(T0.""BPLId"") as ""Headerchk""  FROM ""OPCH"" T0 INNER JOIN ""PCH1"" T1 " &
                                " ON T0.""DocEntry"" = T1.""DocEntry"" WHERE T1.""U_Branch"" = T0.""BPLId"" and T0.""DocNum""='" & Dnum & "'" &
                                 " GROUP BY  T0.""BPLId"""
        Dim rs2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs2.DoQuery(counntchk)
        ' If rs2.RecordCount > 0 Then
        Dim totalheaderbranch As Integer = rs2.Fields.Item("Headerchk").Value
        ' End If
        If totalbranch = totalheaderbranch Then
            objMain.objApplication.StatusBar.SetText("Row Level Branch must be Different from Header Branch for Branch Allocation JE Posting...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If

        Dim Query As String = "Call  ""TNX_APInvoiceAllocation""('" & Dnum & "','" & serie & "')"
        Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs.DoQuery(Query)
        Dim oJE As SAPbobsCOM.JournalEntries
        oJE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
        Dim dateString As String = oDBs_Head.GetValue("DocDate", 0)            '
        Dim format As String = "yyyyMMdd"
        Dim dateValue As Date = DateTime.ParseExact(dateString, format, System.Globalization.CultureInfo.InvariantCulture)
        Dim formattedDate As String = dateValue.ToString("dd.MM.yyyy")
        Dim dateString1 As String = oDBs_Head.GetValue("DocDueDate", 0)
        Dim format1 As String = "yyyyMMdd"
        Dim dateValue1 As Date = DateTime.ParseExact(dateString1, format1, System.Globalization.CultureInfo.InvariantCulture)
        Dim formattedDate1 As String = dateValue1.ToString("dd.MM.yyyy")
        Dim dateString2 As String = oDBs_Head.GetValue("TaxDate", 0)
        Dim format2 As String = "yyyyMMdd"
        Dim dateValue2 As Date = DateTime.ParseExact(dateString, format, System.Globalization.CultureInfo.InvariantCulture)
        Dim formattedDate2 As String = dateValue2.ToString("dd.MM.yyyy")



        oJE.ReferenceDate = dateValue
        oJE.DueDate = dateValue1
        oJE.TaxDate = dateValue2
        oJE.Memo = "Branch Allocation For AP Invoice - '" & oDBs_Head.GetValue("DocNum", 0) & "'"
        oJE.Reference = oDBs_Head.GetValue("DocNum", 0)
        Dim debittotal As Double = 0
        Dim credittototal As Double = 0
        For i As Integer = 0 To rs.RecordCount - 1
            oJE.Lines.AccountCode = rs.Fields.Item("AcctCode").Value
            ' If rs.Fields.Item("Credit").Value > 0 Then
            oJE.Lines.Credit = CDbl(rs.Fields.Item("Credit").Value)
                credittototal = credittototal + CDbl(rs.Fields.Item("Credit").Value)
            'End If
            'If rs.Fields.Item("Debit").Value > 0 Then
            oJE.Lines.Debit = CDbl(rs.Fields.Item("Debit").Value)
                debittotal = debittotal + CDbl(rs.Fields.Item("Debit").Value)
            'End If
            If rs.Fields.Item("Vatgroup").Value <> "" Then
                oJE.Lines.TaxGroup = rs.Fields.Item("Vatgroup").Value
            End If

            oJE.Lines.BPLID = rs.Fields.Item("BPLId").Value
            oJE.Lines.Add()
            rs.MoveNext()
        Next

        If oJE.Add() = 0 Then
            objMain.objApplication.StatusBar.SetText("Branch Allocation JE Posted Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Dim NewTransID As String = objMain.objCompany.GetNewObjectKey.ToString()
            Dim Query1 As String = "Update ""OPCH"" Set ""U_JEEntry""='" + NewTransID + "' Where ""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0) & "'"
            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rs.DoQuery(Query1)
            'objMain.objApplication.ActivateMenuItem("1304")
        Else
            objMain.objApplication.StatusBar.SetText("Branch Allocation JE Posting failed" + objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If
        ' objMain.objApplication.ActivateMenuItem("1304")
        Try
            objMain.objApplication.ActivateMenuItem("1304")
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
        ' End If

    End Sub


    Private Sub PostJE1(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("OPCH")


            '  Dim Branch As SAPbouiCOM.ComboBox = objForm.Items.Item("2001").Specific.Value

            'If Branch.Value = "" Then
            '    objMain.objApplication.StatusBar.SetText("Please select the branch.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Exit Sub
            'End If
            Dim oJE As SAPbobsCOM.JournalEntries = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

            Dim ApInvoiceSrsNum As String = oDBs_Head.GetValue("DocNum", 0).Trim

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

            'Dim sDate As String = DateTime.Now.ToString("yyyyMMdd")
            Dim sDate As String = objForm.Items.Item("10").Specific.Value.ToString
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
                    Dim ItemPer As Double = objMatrix.Columns.Item("213").Cells.Item(i).Specific.Value.ToString().Trim()

                    'Dim SubTotal As Double = "ST1"
                    If (QTY * EXCTAX * ItemPer) <> 0 Then
                        Dim ST1 As Double = QTY * EXCTAX * ItemPer
                        totalTaxValue += ST1
                    End If
                End If
            Next
            '--------------------------------------------------------------------------------------------------------------

            oJE.ReferenceDate = sDate
            oJE.DueDate = sDate
            oJE.TaxDate = sDate
            oJE.Reference2 = "A/P Invoice"

            '  oJE.Lines.BPLID = Branch.Value
            oJE.Lines.AccountCode = PurchaseexcisedutyCrGL
            oJE.Lines.Credit = totalTaxValue
            oJE.Memo = "A/P Invoice Posted :" & ApInvoiceSrsNum
            oJE.Lines.Add()
            '  oJE.Lines.BPLID = Branch.Value
            oJE.Lines.AccountCode = PurchaseexcisedutyDrGL
            oJE.Lines.Debit = totalTaxValue
            oJE.Memo = "A/P Invoice Posted :" & ApInvoiceSrsNum
            oJE.Lines.Add()

            If oJE.Add() = 0 Then
                Dim NewTransID As String = objMain.objCompany.GetNewObjectKey.ToString()

                Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim Update As String = " UPDATE ""OPCH"" T1 SET T1.""U_EXCJVN""='" + NewTransID + "' WHERE T1.""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0) & "'"
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
            oDBs_Head = objForm.DataSources.DBDataSources.Item("OPCH")

            Dim oJE As SAPbobsCOM.JournalEntries = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

            Dim APInvoiceSrsNum As String = oDBs_Head.GetValue("DocNum", 0).Trim

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

            oJE.Reference2 = "A/P Invoice"

            '  oJE.Lines.BPLID = Branch.Value
            oJE.Lines.AccountCode = PurchaseexcisedutyCrGL
            oJE.Lines.Credit = totalTaxValue1
            oJE.Memo = "A/P Invoice Posted :" & APInvoiceSrsNum
            oJE.Lines.BPLID = CInt(sBrachCode)
            oJE.Lines.Add()
            '  oJE.Lines.BPLID = Branch.Value
            oJE.Lines.AccountCode = PurchaseexcisedutyDrGL
            oJE.Lines.Debit = totalTaxValue1
            oJE.Memo = "A/P Invoice Posted :" & APInvoiceSrsNum
            oJE.Lines.BPLID = CInt(sBrachCode)
            oJE.Lines.Add()

            If oJE.Add() = 0 Then
                Dim NewTransID As String = objMain.objCompany.GetNewObjectKey.ToString()

                Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim Update As String = " UPDATE ""OPCH"" T1 SET T1.""U_EXCJVN""='" + NewTransID + "' WHERE T1.""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0) & "'"
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
End Class
