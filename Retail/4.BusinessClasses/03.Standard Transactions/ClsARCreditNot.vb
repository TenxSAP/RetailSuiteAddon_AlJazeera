Public Class ClsARCreditNot

    Dim objForm As SAPbouiCOM.Form
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oDBs_Head, oDBs_Details, oDs2 As SAPbouiCOM.DBDataSource
    Dim ButtonCalFlag As Boolean
    Dim oitem As SAPbouiCOM.Item

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("ORIN")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("RIN1")
                    oDs2 = objForm.DataSources.DBDataSources.Item("RIN3")

                    objMatrix = objForm.Items.Item("38").Specific

                    If pVal.BeforeAction = True Then
                        Me.AddItems(objForm.UniqueID)
                    End If


                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("ORIN")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("RIN1")
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
                                                         """U_GUID""='" & objForm.Items.Item("txt_GUID").Specific.Value & "' And ""U_DOCTY""='ORIN'"

                            Dim oRsCheckBackDOc As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsCheckBackDOc.DoQuery(CheckBackDOc)

                            Dim BaseType As Double = objMatrix.Columns.Item("43").Cells.Item(1).Specific.Value

                            If BaseType = -1 And oRsGetWCnd.RecordCount > 0 And BaseType <> 21 And (oRsCheckBackDOc.RecordCount = 0 Or CStr(CStr(objForm.Items.Item("txt_GUID").Specific.Value).Trim).Trim = "") Then
                                objMain.objApplication.StatusBar.SetText("please add Expiry Screen before adding")
                                objForm.Freeze(False)
                                BubbleEvent = False
                                Exit Try
                            End If


                            If BaseType = -1 And oRsGetWCnd.RecordCount > 0 And BaseType <> 21 And (oRsCheckBackDOc.RecordCount = 0 Or CStr(CStr(objForm.Items.Item("txt_GUID").Specific.Value).Trim).Trim = "") Then
                                objMain.objApplication.StatusBar.SetText("please add Expiry Screen before adding")
                                objForm.Freeze(False)
                                BubbleEvent = False

                                Exit Try
                            End If






                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If

                    If pVal.ItemUID = "btn_ED" And pVal.BeforeAction = False Then
                        Try
                            objForm.Freeze(True)


                            'Dim GetData As String = "SELECT T1.""LineNum"",T1.""ItemCode"", T1.""Dscription"", T1.""Quantity"", T1.""CodeBars"", " &
                            '                        "T1.""UomCode"" FROM ORIN T0  INNER JOIN RIN1 T1 ON T0.""DocEntry"" = T1.""DocEntry""  " &
                            '                        " INNER JOIN OWHS T2 ON T1.""WhsCode"" = T2.""WhsCode"" WHERE  " &
                            '                        "T0.""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0) & "' And IFNULL(T2.""U_EXCHK"",'No') ='Yes' " &
                            '                        "And T0.""DocType""='I' And T1.""NoInvtryMv""='N'  And T1.""BaseType"" IS NULL ORDER BY T1.""LineNum"""

                            'Dim oRsGetData As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            'oRsGetData.DoQuery(GetData)

                            'If oRsGetData.RecordCount > 0 Then
                            '    objMain.objExpiry.CreateForm(objForm.UniqueID, "ORIN", oDBs_Head.GetValue("DocEntry", 0), oDBs_Head.GetValue("DocNum", 0), oRsGetData)
                            'Else
                            '    objMain.objApplication.StatusBar.SetText("Expiry Date screen not applicable")
                            'End If



                            If CStr(objForm.Items.Item("3").Specific.Value).Trim <> "I" Then
                                objMain.objApplication.StatusBar.SetText("Expiry Screen not applicable")
                                objForm.Freeze(False)
                                Exit Try
                            End If

                            Dim BaseType As Double = objMatrix.Columns.Item("43").Cells.Item(1).Specific.Value

                            If BaseType <> -1 Then
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

GotoLst:                    objMain.objExpiry.CreateFormNew(objForm.UniqueID, "ORIN", oDBs_Head.GetValue("DocEntry", 0), oDBs_Head.GetValue("DocNum", 0), dt, UID.Trim, objForm.Mode)



                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If


                    Dim objComboBox1 As SAPbouiCOM.ComboBox = objForm.Items.Item("3").Specific
                    Dim TYPE As String = objComboBox1.Value
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And TYPE = "I" Then
                        Try
                            Dim objCheck As SAPbouiCOM.CheckBox = objMatrix.Columns.Item("1250002121").Cells.Item(1).Specific
                            If objCheck.Checked = True Then
                                ButtonCalFlag = True

                            End If
                            If ButtonCalFlag = False Then
                                objMain.objApplication.StatusBar.SetText("Please click Calculate Excise button before adding")
                                BubbleEvent = False
                            End If
                        Catch ex As Exception

                        End Try
                    End If



                    If pVal.ItemUID = "CLEX" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And TYPE = "I" Then
                        Try
                            objForm.Freeze(True)

                            ButtonCalFlag = True


                            For R As Integer = 1 To objMatrix.VisualRowCount

                                If objMatrix.Columns.Item("1").Cells.Item(R).Specific.Value <> "" Then


                                    Dim GetPerPCVal As Double = 0

                                    Dim GetPerPC As String = "SELECT ""BaseQty""  FROM  ""OUGP""  T0 INNER JOIN  ""UGP1""  T1 ON T0.""UgpEntry"" = T1.""UgpEntry"" " &
                                                            " INNER JOIN  ""OITM""  T2 ON T0.""UgpEntry"" = T2.""UgpEntry"" WHERE T2.""ItemCode"" ='" & CStr(objMatrix.Columns.Item("1").Cells.Item(R).Specific.Value).Trim & "' " &
                                                            " And ""UomEntry"" IN (SELECT T0.""UomEntry"" FROM OUOM T0 WHERE T0.""UomCode"" ='" & CStr(objMatrix.Columns.Item("1470002145").Cells.Item(R).Specific.Value).Trim & "') "

                                    Dim oRsGetPerPCas As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRsGetPerPCas.DoQuery(GetPerPC)

                                    If oRsGetPerPCas.RecordCount > 0 Then
                                        GetPerPCVal = oRsGetPerPCas.Fields.Item(0).Value
                                    End If

                                    Dim Discount As Double = 0
                                    Dim Total As Double = 0


                                    If CStr(objMatrix.Columns.Item("15").Cells.Item(R).Specific.Value).Trim <> "" Then
                                        Discount = CDbl(CStr(objMatrix.Columns.Item("15").Cells.Item(R).Specific.Value).Trim)
                                    End If


                                    If CStr(objMatrix.Columns.Item("21").Cells.Item(R).Specific.Value).Trim <> "" Then
                                        Total = String.Concat(CStr(objMatrix.Columns.Item("21").Cells.Item(R).Specific.Value).Trim.Where(Function(c) Char.IsDigit(c))) 'CDbl(CStr(objMatrix.Columns.Item("21").Cells.Item(R).Specific.Value).Trim)
                                    End If


                                    ' If Discount = 100 Or Total = 0 Then
                                    'If Total = 0 Then

                                    '    Dim EXPNM As String = " Select ""ExpnsCode"" from OEXD  where ""ExpnsName"" = 'Local Excise' "
                                    '    Dim oRsEXPNM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    '    oRsEXPNM.DoQuery(EXPNM)

                                    '    Try
                                    '        Dim ExpnsCode1 As String = oRsEXPNM.Fields.Item("ExpnsCode").Value

                                    '        Dim objComboBox8 As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("111").Cells.Item(R).Specific
                                    '        Dim TaskName8 As String = objComboBox8.Selected.Value

                                    '        objComboBox8.Select("-1", SAPbouiCOM.BoSearchKey.psk_ByValue)

                                    '        objMatrix.Columns.Item("U_EXCTAX").Cells.Item(R).Specific.Value() = 0
                                    '        objMatrix.Columns.Item("112").Cells.Item(R).Specific.Value = 0


                                    '    Catch ex As Exception

                                    '    End Try

                                    'Else

                                    Dim cardcode As String = objForm.Items.Item("4").Specific.Value
                                    Dim Getdata As String = "SELECT T0.""U_BPTYPE"",T0.""ExpnsCode"",T0.""ExpnsName"" FROM OEXD T0 inner join NNM1 T1 ON T1.""Series""=T0.""U_BPSERIES"" " +
                                                                    "  inner join OCRD T2 on T2.""Series""=T1.""Series"" where  T2.""CardCode""='" & cardcode & "'"
                                    Dim oRs1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRs1.DoQuery(Getdata)
                                    Dim ExpnsCode As String = oRs1.Fields.Item("ExpnsCode").Value
                                    Dim BPType As String = oRs1.Fields.Item("U_BPTYPE").Value

                                    If BPType = "LOCAL" Then 'And (customertype = "CUS" Or customertype = "Manual") Then
                                        Dim objComboBox3 As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("111").Cells.Item(R).Specific
                                        Dim TaskName3 As String = objComboBox3.Selected.Value
                                        objComboBox3.Select(ExpnsCode, SAPbouiCOM.BoSearchKey.psk_ByValue)

                                        objMatrix = objForm.Items.Item("38").Specific
                                        Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(R).Specific.Value
                                        Dim GetDocEntry As String = "Select ""U_EXCVAT"" From OITM Where ""ItemCode"" = '" & ItemCode & "' "
                                        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRs.DoQuery(GetDocEntry)

                                        objMatrix.Columns.Item("U_EXCTAX").Cells.Item(R).Specific.Value() = oRs.Fields.Item("U_EXCVAT").Value()

                                        Dim ItemCode1 As String = objMatrix.Columns.Item("1").Cells.Item(R).Specific.Value.ToString().Trim()
                                        Dim QTY As Double = objMatrix.Columns.Item("11").Cells.Item(R).Specific.Value.ToString().Trim()
                                        Dim EXCTAX As Double = objMatrix.Columns.Item("U_EXCTAX").Cells.Item(R).Specific.Value.ToString().Trim() * GetPerPCVal
                                        Dim SubTotal As String = "ST1"
                                        If (QTY * EXCTAX) <> 0 Then
                                            Dim ST1 As Double = QTY * EXCTAX
                                            SubTotal = ST1

                                            If objMatrix.Columns.Item("U_EXCTAX").Cells.Item(R).Specific.Value <> "" Then
                                                Dim objComboBox As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("111").Cells.Item(R).Specific
                                                Dim TaskName As String = objComboBox.Selected.Value
                                                ' objComboBox.Select("Sales Excise Duty", SAPbouiCOM.BoSearchKey.psk_ByDescription)
                                                objComboBox.Select(ExpnsCode, SAPbouiCOM.BoSearchKey.psk_ByValue)

                                                ' Dim objComboBox As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("111").Cells.Item(R).Specific
                                                Dim TaskName1 As String = objComboBox.Selected.Value
                                                If objMatrix.Columns.Item("111").Cells.Item(R).Specific.Value = TaskName1 Then
                                                    objMatrix.Columns.Item("112").Cells.Item(R).Specific.Value = SubTotal
                                                End If

                                            End If
                                        Else
                                            Dim objComboBox As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("111").Cells.Item(R).Specific
                                            Dim TaskName As String = objComboBox.Selected.Value
                                            objComboBox.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                        End If

                                    Else
                                        Dim EXPNM As String = " Select ""ExpnsCode"" from OEXD  where ""U_BPTYPE"" = 'FOREIGN' "
                                        Dim oRsEXPNM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRsEXPNM.DoQuery(EXPNM)
                                        Dim ExpnsCode1 As String = oRsEXPNM.Fields.Item("ExpnsCode").Value

                                        Dim objComboBox8 As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("111").Cells.Item(R).Specific
                                        Dim TaskName8 As String = objComboBox8.Selected.Value
                                        objComboBox8.Select(ExpnsCode, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                        objMatrix = objForm.Items.Item("38").Specific
                                        Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(R).Specific.Value
                                        Dim GetDocEntry As String = "Select ""U_EXCVAT"" From OITM Where ""ItemCode"" = '" & ItemCode & "' "
                                        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRs.DoQuery(GetDocEntry)
                                        objMatrix.Columns.Item("U_EXCTAX").Cells.Item(R).Specific.Value() = oRs.Fields.Item("U_EXCVAT").Value()
                                    End If


                                End If


                                ' End If

                            Next


                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
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


            objMain.objUtilities.AddEditBox(objForm.UniqueID, "txt_GUID", objForm.Items.Item("4").Top, objForm.Items.Item("4").Left + objForm.Items.Item("4").Width + 5, objForm.Items.Item("4").Width, "ORIN", "U_GUID", "", objForm.Items.Item("4").Height, 96, 96)

            objForm.Items.Item("txt_GUID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)


            oItem = objForm.Items.Add("CLEX", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 10
            oItem.Left = objForm.Items.Item("38").Left + 100
            oItem.Width = objForm.Items.Item("2").Width + 60
            oItem.Height = objForm.Items.Item("2").Height
            oItem.FromPane = objForm.Items.Item("38").FromPane
            oItem.ToPane = objForm.Items.Item("38").ToPane
            Dim btn2 As SAPbouiCOM.Button = objForm.Items.Item("CLEX").Specific

            btn2.Caption = "Calculate Excise"


        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("179", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then


                        If CStr(oDBs_Head.GetValue("U_GUID", 0)).Trim <> "" Then

                            Dim UpdateDa As String = "Update ""@TNXEDS"" SET ""U_BENTRY""='" & oDBs_Head.GetValue("DocEntry", 0) & "', " &
                                                  """U_BNO""='" & oDBs_Head.GetValue("DocNum", 0) & "' Where ""U_GUID""='" & oDBs_Head.GetValue("U_GUID", 0) & "'  And ""U_DOCTY""='ORIN'"

                            Dim oRsUpdateDa As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            oRsUpdateDa.DoQuery(UpdateDa)

                        End If


                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
