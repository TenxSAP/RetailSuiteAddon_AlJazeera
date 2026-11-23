Public Class ClsOutgoingPayments

    Public objform, objBaseForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim oItem As SAPbouiCOM.Item
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim j, linechk As Integer
    Dim objmatrix As SAPbouiCOM.Matrix
    Dim ItemCodechk As String
    Dim LostFocusFlag As Boolean = True
    Dim oColumn As SAPbouiCOM.Column
    'Dim qualtychk As Double

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.BeforeAction = True Then
                        Me.AddItems(objform.UniqueID)
                    Else
                        Try
                            objmatrix = objform.Items.Item("71").Specific
                            objComboBox = objMatrix.Columns.Item("U_Branch").Cells.Item(1).Specific  ''objForm.Items.Item("14").Specific
                            oColumn = objMatrix.Columns.Item("U_Branch")
                            Dim sQury As String = " Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join " &
                                                  " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""  "
                            objMain.objUtilities.MatrixComboBoxValues1(oColumn, sQury)
                            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                        Catch ex As Exception

                        End Try
                    End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                                (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                        'If Me.Validate(objform.UniqueID) = False Then
                        '    LostFocusFlag = False
                        '    BubbleEvent = False
                        'End If

                    End If
                    'If pVal.ItemUID = "CBChar_JE" And pVal.BeforeAction = False And objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    '    objform = objMain.objApplication.Forms.Item(FormUID)
                    '    oDBs_Head = objform.DataSources.DBDataSources.Item("OVPM")
                    '    'oDBs_Details = objform.DataSources.DBDataSources.Item("PCH1")
                    '    Me.PostJE(objform.UniqueID)
                    'End If
                    If pVal.ItemUID = "BAlloc_JE" And pVal.BeforeAction = False And objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        objform = objMain.objApplication.Forms.Item(FormUID)
                        oDBs_Head = objform.DataSources.DBDataSources.Item("OVPM")
                        oDBs_Details = objform.DataSources.DBDataSources.Item("VPM4")
                        Me.BranchPostJE()
                    End If


            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub PostJE(FormUID)
        objform = objMain.objApplication.Forms.Item(FormUID)
        oDBs_Head = objform.DataSources.DBDataSources.Item("OVPM")

        If oDBs_Head.GetValue("TrsfrSum", 0) < 1 Then
            objMain.objApplication.StatusBar.SetText("Bank Transfer Amount must be greater than zero.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
            Dim chk As String = oDBs_Head.GetValue("U_BankCharge", 0)
        ElseIf (oDBs_Head.GetValue("U_BankCharge", 0) = "") Or (oDBs_Head.GetValue("U_BankCharge", 0) = "0.0") Then
            objMain.objApplication.StatusBar.SetText("Bank Charge Amount must be greater than zero.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If


        Dim Dnum As String = oDBs_Head.GetValue("DocNum", 0).Trim
        Dim serie As String = oDBs_Head.GetValue("Series", 0).Trim
        ' If objform.Items.Item("JEEntry").Specific.Value = "" Then
        Dim Query As String = "Call  ""TNX_BankCharges""('" & Dnum & "','" & serie & "')"
        Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs.DoQuery(Query)
        Dim oJE As SAPbobsCOM.JournalEntries
        oJE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
        Dim dateString As String = oDBs_Head.GetValue("DocDate", 0)
        Dim format As String = "yyyyMMdd"
        Dim dateValue As DateTime = DateTime.ParseExact(dateString, format, System.Globalization.CultureInfo.InvariantCulture)
        Dim formattedDate As String = dateValue.ToString("dd.MM.yyyy")
        Dim dateString1 As String = oDBs_Head.GetValue("DocDueDate", 0)
        Dim format1 As String = "yyyyMMdd"
        Dim dateValue1 As DateTime = DateTime.ParseExact(dateString1, format1, System.Globalization.CultureInfo.InvariantCulture)
        Dim formattedDate1 As String = dateValue1.ToString("dd.MM.yyyy")
        Dim dateString2 As String = oDBs_Head.GetValue("TaxDate", 0)
        Dim format2 As String = "yyyyMMdd"
        Dim dateValue2 As DateTime = DateTime.ParseExact(dateString, format, System.Globalization.CultureInfo.InvariantCulture)
        Dim formattedDate2 As String = dateValue2.ToString("dd.MM.yyyy")


        oJE.ReferenceDate = dateValue
        oJE.DueDate = dateValue1
        oJE.TaxDate = dateValue2
        oJE.Memo = "Bank Charges JE for O/P '" & oDBs_Head.GetValue("DocNum", 0) & "'"
        oJE.Reference = oDBs_Head.GetValue("DocNum", 0)
        oJE.AutoVAT = SAPbobsCOM.BoYesNoEnum.tYES
        Dim debittotal As Double = 0
        Dim credittototal As Double = 0
        For i As Integer = 0 To rs.RecordCount - 1

            oJE.Lines.AccountCode = rs.Fields.Item("TrsfrAcct").Value
            'If rs.Fields.Item("Credit").Value > 0 Then
            oJE.Lines.Credit = CDbl(rs.Fields.Item("Credit").Value)
                credittototal = credittototal + CDbl(rs.Fields.Item("Credit").Value)
            'End If
            'If rs.Fields.Item("Debit").Value > 0 Then
            oJE.Lines.Debit = CDbl(rs.Fields.Item("Debit").Value)
                debittotal = debittotal + CDbl(rs.Fields.Item("Debit").Value)
            'End If
            oJE.Lines.BPLID = rs.Fields.Item("BPLId").Value
            If rs.Fields.Item("VatGroup").Value <> "" Then
                oJE.Lines.TaxGroup = rs.Fields.Item("VatGroup").Value
            End If
            oJE.Lines.Add()
            rs.MoveNext()
        Next

        If oJE.Add() = 0 Then
            objMain.objApplication.StatusBar.SetText("Bank Charges JE for O/P Posted Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Dim Query1 As String = "Update ""OVPM"" Set ""U_JEEntry""=(Select Max(""TransId"") FROM OJDT) Where ""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0) & "'"
            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rs.DoQuery(Query1)

            'objMain.objApplication.ActivateMenuItem("1304")
        Else
            objMain.objApplication.StatusBar.SetText("Bank Charges JE for O/P Posting failed" + objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If
        ' objMain.objApplication.ActivateMenuItem("1304")
        Try
            objMain.objApplication.ActivateMenuItem("1304")
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
        ' End If

    End Sub

    Sub BranchPostJE()
        oDBs_Head = objform.DataSources.DBDataSources.Item("OVPM")
        oDBs_Details = objform.DataSources.DBDataSources.Item("VPM4")
        objmatrix = objform.Items.Item("71").Specific

        Dim Dnum As String = oDBs_Head.GetValue("DocNum", 0).Trim
        Dim serie As String = oDBs_Head.GetValue("Series", 0).Trim


        Dim doctypechk As String = "SELECT ""DocType"" from ""OVPM"" WHERE ""DocNum"" = '" & Dnum & "' and ""Series""='" & serie & "'"

        Dim rs3 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs3.DoQuery(doctypechk)
        '  If rs1.RecordCount > 0 Then
        Dim dtype As String = rs3.Fields.Item("DocType").Value


        If dtype <> "A" Then
            objMain.objApplication.StatusBar.SetText("Document Type must be Selected as Account Type for Branch Allocation JE Posting ......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If


        For i As Integer = 1 To objmatrix.VisualRowCount

            ' Dim objComboBoxBranch As SAPbouiCOM.ComboBox = objform.Items.Item("1320000151").Specific
            Dim linebranch As String = objmatrix.Columns.Item("U_Branch").Cells.Item(i).Specific.Value.ToString().Trim
            If linebranch = "" Then
                objMain.objApplication.StatusBar.SetText("Row Level Branch must be fill in all row for Branch Allocation JE Posting...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

        Next

        Dim allrecordchk As String = "SELECT COUNT(T1.""U_Branch"") as ""Linechk""  FROM ""OVPM"" T0 INNER JOIN ""VPM4"" T1 " &
                            " ON T0.""DocEntry"" = T1.""DocNum"" WHERE T1.""U_Branch"" = T1.""U_Branch"" and T0.""DocNum""='" & Dnum & "'"

        Dim rs1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs1.DoQuery(allrecordchk)
        '  If rs1.RecordCount > 0 Then
        Dim totalbranch As Integer = rs1.Fields.Item("Linechk").Value
        ' End If
        Dim counntchk As String = "SELECT COUNT(T0.""BPLId"") as ""Headerchk""  FROM ""OVPM"" T0 INNER JOIN ""VPM4"" T1 " &
                                " ON T0.""DocEntry"" = T1.""DocNum"" WHERE T1.""U_Branch"" = T0.""BPLId"" and T0.""DocNum""='" & Dnum & "'" &
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


        Dim Query As String = "Call  ""TNX_OPBranchJEAllocation""('" & Dnum & "','" & serie & "')"
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
        oJE.Memo = "Branch Allocation For O/P - '" & oDBs_Head.GetValue("DocNum", 0) & "'"
        oJE.Reference = oDBs_Head.GetValue("DocNum", 0)
        Dim debittotal As Double = 0
        Dim credittototal As Double = 0
        For i As Integer = 0 To rs.RecordCount - 1

            oJE.Lines.AccountCode = rs.Fields.Item("AcctCode").Value
            If rs.Fields.Item("Credit").Value > 0 Then
                oJE.Lines.Credit = CDbl(rs.Fields.Item("Credit").Value)
                credittototal = credittototal + CDbl(rs.Fields.Item("Credit").Value)
            End If
            If rs.Fields.Item("Debit").Value > 0 Then
                oJE.Lines.Debit = CDbl(rs.Fields.Item("Debit").Value)
                debittotal = debittotal + CDbl(rs.Fields.Item("Debit").Value)
            End If
            oJE.Lines.BPLID = rs.Fields.Item("BPLId").Value
            oJE.Lines.Add()
            rs.MoveNext()
        Next

        If oJE.Add() = 0 Then
            objMain.objApplication.StatusBar.SetText("Branch Allocation JE Posted Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Dim Query1 As String = "Update ""OVPM"" Set ""U_JEEntry""=(Select Max(""TransId"") FROM OJDT) Where ""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0) & "'"
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

    Private Sub AddItems(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)

            'oItem = objform.Items.Add("CBChar_JE", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'oItem.Top = objform.Items.Item("2").Top
            'oItem.Left = objform.Items.Item("2").Left + 65 + objform.Items.Item("2").Height - 2
            'oItem.Width = objform.Items.Item("2").Width + 80
            'oItem.Height = objform.Items.Item("2").Height
            '' oItem.LinkTo = "2"
            'Dim btn As SAPbouiCOM.Button = objform.Items.Item("CBChar_JE").Specific
            'btn.Caption = "Create BankCharges JE"

            oItem = objform.Items.Add("BAlloc_JE", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Top = objform.Items.Item("2").Top
            oItem.Left = objform.Items.Item("2").Left + 80
            oItem.Width = objform.Items.Item("2").Width + 50
            oItem.Height = objform.Items.Item("2").Height
            '' oItem.LinkTo = "CBChar_JE"
            Dim btn1 As SAPbouiCOM.Button = objform.Items.Item("BAlloc_JE").Specific
            btn1.Caption = "BranchAllocation JE"


            objMain.objUtilities.AddLabel(objform.UniqueID, "JENO", objform.Items.Item("53").Top + 18, objform.Items.Item("53").Left, objform.Items.Item("53").Width - 25, "JE Number", "22", 0, 0)
            oItem = objform.Items.Add("JEEntry", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = objform.Items.Item("52").Top + 18
            oItem.Left = objform.Items.Item("52").Left  '+ objform.Items.Item("4").Width + 325
            oItem.Width = objform.Items.Item("52").Width '+ 35
            oItem.Height = objform.Items.Item("52").Height '+ 100
            oItem.LinkTo = "JENO"

            Dim JE As SAPbouiCOM.EditText = objform.Items.Item("JEEntry").Specific
            JE.DataBind.SetBound(True, "OVPM", "U_JEEntry")

            oItem = objform.Items.Add("LinK", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oItem.Top = objform.Items.Item("53").Top + 18
            oItem.Left = objform.Items.Item("53").Left + objform.Items.Item("JENO").Width + 28
            oItem.Width = 12
            oItem.Height = 10
            oItem.LinkTo = "JEEntry"
            Dim LinkBtn As SAPbouiCOM.LinkedButton
            LinkBtn = objform.Items.Item("LinK").Specific
            LinkBtn.LinkedObject = 30

            'objform.Items.Item("CBChar_JE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objform.Items.Item("CBChar_JE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objform.Items.Item("BAlloc_JE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objform.Items.Item("BAlloc_JE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)



            objform.Items.Item("JEEntry").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objform.Items.Item("JEEntry").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objform.Items.Item("JEEntry").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objform.Items.Item("JEEntry").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)



        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objform = objMain.objApplication.Forms.GetForm("426", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        oDBs_Head = objform.DataSources.DBDataSources.Item("OVPM")

                        Dim Dnum As String = oDBs_Head.GetValue("DocNum", 0).Trim
                        Dim serie As String = oDBs_Head.GetValue("Series", 0).Trim


                        Dim doctypechk As String = "SELECT ""DocType"" from ""OVPM"" WHERE ""DocNum"" = '" & Dnum & "' and ""Series""='" & serie & "'"

                        Dim rs1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        rs1.DoQuery(doctypechk)
                        '  If rs1.RecordCount > 0 Then
                        Dim dtype As String = rs1.Fields.Item("DocType").Value


                        If dtype = "A" Then
                            objform.Items.Item("BAlloc_JE").Enabled = True
                            'objform.Items.Item("CBChar_JE").Enabled = False
                        ElseIf dtype = "S" Then
                            objform.Items.Item("BAlloc_JE").Enabled = False
                            'objform.Items.Item("CBChar_JE").Enabled = True
                        Else
                            objform.Items.Item("BAlloc_JE").Enabled = False
                            'objform.Items.Item("CBChar_JE").Enabled = False
                        End If
                        Dim entry As String = oDBs_Head.GetValue("U_JEEntry", 0)
                        If entry <> "" Then
                            'objform.Items.Item("CBChar_JE").Enabled = False
                            objform.Items.Item("BAlloc_JE").Enabled = False
                            'Else

                            ' objform.Items.Item("CBChar_JE").Enabled = True
                            ' objform.Items.Item("BAlloc_JE").Enabled = True
                        End If



                    End If



            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Function Validate(ByVal FormUID As String) As Boolean
        Try

            objform = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objform.DataSources.DBDataSources.Item("OPOR")

            For i As Integer = 1 To objmatrix.VisualRowCount - 1
                Dim objMatrix As SAPbouiCOM.Matrix = objform.Items.Item("38").Specific
                Dim quant As String = objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value
                Dim Itemcode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value
                Dim DocEntry As String = objMatrix.Columns.Item("U_SWEntry").Cells.Item(i).Specific.Value
                Dim quantity As Double = Convert.ToDouble(quant)

                If (DocEntry = "") Then
                    objMain.objApplication.StatusBar.SetText("Please Select Quantity in Shopwise Allocation First in row number " & i & "... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                Dim Query As String = "Select IFNULL(Sum(T0.""U_Qty""),0) as ""OverallQuantity""  from ""@TNX_WHSALOC_C0"" T0 inner join " &
                               " ""@TNX_WHSALOC"" T1 on T0.""DocEntry""=T1.""DocEntry"" " &
                                " where T1.""DocEntry"" = '" & DocEntry & "'" &
                                       " and  T0.""U_Qty"" <> '0' "
                Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rs.DoQuery(Query)


                If rs.RecordCount > 0 Then

                    Dim overallquan As String = rs.Fields.Item("OverallQuantity").Value()
                    Dim overallquantity As Double = Convert.ToDouble(overallquan)

                    If (overallquantity = 0) Then
                        objMain.objApplication.StatusBar.SetText("Please Select Quantity in Shopwise Allocation First in row number " & i & "... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        ' objMatrix.Columns.Item("14").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Return False

                    End If
                    '  rs.MoveNext()

                End If

            Next
            If objform.Items.Item("ItmImg").Specific.Picture <> "" Then
                Dim imagechk As String = objform.Items.Item("ItmImg").Specific.Picture
                Dim Afterchk As String = " "
                objform.Items.Item("ItmImg").Specific.Picture = Afterchk
            End If
            Return True

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function


End Class
