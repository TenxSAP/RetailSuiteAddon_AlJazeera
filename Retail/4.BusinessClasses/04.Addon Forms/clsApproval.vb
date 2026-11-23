Public Class clsApproval

#Region "       Declaration             "
    Public objForm, objSTRForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details1 As SAPbouiCOM.DBDataSource
    Dim objMatrix1 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim objCheckBox, objCheckBox1 As SAPbouiCOM.CheckBox
    Dim str, str1, str2, str3, Query, str4 As String
    Public rs, RsNum As SAPbobsCOM.Recordset
    Dim LostFocusFlag As Boolean = False
    Dim objutilities As Utilities
    Dim oDt As SAPbouiCOM.DataTable
    Dim oGrid As SAPbouiCOM.Grid
    Dim r1, r2, r3 As RadioButton
    Dim column, column1 As SAPbouiCOM.EditTextColumn

#End Region

    Sub CreateForm()

        objMain.objUtilities.LoadForm("Approval1.xml", "ME_ASR", ResourceType.Embeded)
        objForm = objMain.objApplication.Forms.GetForm("ME_ASR", objMain.objApplication.Forms.ActiveForm.TypeCount)
        objForm.Items.Item("Item_71").TextStyle = 1
        'objForm.Items.Item("Item_1").TextStyle = 1
        'objForm.Items.Item("Item_38").TextStyle = 1
        'objForm.Items.Item("Item_46").TextStyle = 1
        'objForm.Items.Item("Item_66").TextStyle = 1
        'objForm.Items.Item("Item_71").TextStyle = 1
        'objForm.Items.Item("Item_21").TextStyle = 0
        'objForm.Items.Item("Item_105").TextStyle = 1

        objForm.Items.Item("31").TextStyle = 1
        objForm.Items.Item("EMP").Visible = False
        objForm.Items.Item("Item_2").Visible = False

        'objForm.Items.Item("EMP").Visible = True
        'objForm.Items.Item("Item_2").Visible = True


        'objForm.Items.Item("Item_39").Visible = False
        'objForm.Items.Item("UD_EMD").Visible = False
        'objForm.Items.Item("Item_41").Visible = False
        'objForm.Items.Item("UD_ESS").Visible = False
        'objForm.Items.Item("Item_47").Visible = False
        ' objForm.Items.Item("Item_51").Visible = False
        'objForm.Items.Item("UD_AS").Visible = False

        '   objForm.Items.Item("Item_19").Visible = False
        '        objForm.Items.Item("UD_PPP").Visible = False
        LoadUDBranchCombo()
        objForm.Items.Item("Item_0").Visible = False
        objForm.Items.Item("AUT_1").Visible = False
        ' objForm.Items.Item("UD_TS").Visible = False
        '  'objForm.DataSources.UserDataSources.Item("EMP_1").e
        objCheckBox = objForm.Items.Item("UD_PEN").Specific
        objCheckBox.Checked = True

    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ME_ASR" And pVal.BeforeAction = False Then
                Me.CreateForm()
                'Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                'Me.SetDefault(objForm.UniqueID)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)

        Select Case pVal.EventType
            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                objForm = objMain.objApplication.Forms.Item(FormUID)

                If pVal.ItemUID = "13" And pVal.BeforeAction = False And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                    Dim Pending As String = objForm.DataSources.UserDataSources.Item("UD_PEN").Value
                    Dim Submitted As String = Pending
                    Dim Approved As String = objForm.DataSources.UserDataSources.Item("UD_APP").Value
                    Dim Rejected As String = objForm.DataSources.UserDataSources.Item("UD_REJ").Value
                    Dim Canceled As String = objForm.DataSources.UserDataSources.Item("UD_CAN").Value
                    Dim RecruitmentRequest As String = objForm.DataSources.UserDataSources.Item("UD_REC").Value
                    Dim ResumeSubmission As String = objForm.DataSources.UserDataSources.Item("UD_RES").Value
                    Dim OfferLetter As String = objForm.DataSources.UserDataSources.Item("UD_OL").Value
                    Dim EmployeeMasterData As String = objForm.DataSources.UserDataSources.Item("UD_EMD").Value
                    Dim EmployeeSalarySetup As String = objForm.DataSources.UserDataSources.Item("UD_ESS").Value
                    Dim EmployeeContractRenewal1 As String = objForm.DataSources.UserDataSources.Item("UD_ECR").Value
                    Dim Timesheet As String = objForm.DataSources.UserDataSources.Item("UD_TS").Value
                    Dim LeaveRejoinApplication As String = objForm.DataSources.UserDataSources.Item("UD_LRA").Value
                    Dim LeaveApplication As String = objForm.DataSources.UserDataSources.Item("UD_LA").Value
                    Dim CompensatoryOff As String = objForm.DataSources.UserDataSources.Item("UD_COFF").Value
                    Dim AttendanceSummary As String = objForm.DataSources.UserDataSources.Item("UD_AS").Value
                    Dim ShiftShedule As String = objForm.DataSources.UserDataSources.Item("UD_SS").Value
                    Dim AttendanceRegularisation As String = objForm.DataSources.UserDataSources.Item("UD_AR").Value
                    Dim LeaveEncashmentrequest As String = objForm.DataSources.UserDataSources.Item("UD_ECReq").Value
                    Dim LeavePeriodEndClc As String = objForm.DataSources.UserDataSources.Item("UD_LPE").Value
                    Dim LoansandAdvanceApplication As String = objForm.DataSources.UserDataSources.Item("UD_LAA").Value
                    Dim LoanPreClosureApplication As String = objForm.DataSources.UserDataSources.Item("UD_LPCA").Value
                    Dim PaymentandDeduction As String = objForm.DataSources.UserDataSources.Item("UD_PD").Value
                    Dim PayrollPreProcess As String = objForm.DataSources.UserDataSources.Item("UD_PPP").Value
                    Dim ClaimApplication As String = objForm.DataSources.UserDataSources.Item("UD_CA").Value
                    Dim EmployeeReveiwAppraisal As String = objForm.DataSources.UserDataSources.Item("UD_ERA").Value
                    Dim LeaveEncashment As String = objForm.DataSources.UserDataSources.Item("UD_LE").Value
                    Dim EndofServiceProcess As String = objForm.DataSources.UserDataSources.Item("UD_ESP").Value
                    Dim EmployeeResignationRequest As String = objForm.DataSources.UserDataSources.Item("UD_ERR").Value
                    Dim ExitClearance As String = objForm.DataSources.UserDataSources.Item("UD_ECL").Value
                    Dim AirticketRequest As String = objForm.DataSources.UserDataSources.Item("UD_AIR").Value
                    Dim DocumentRequest As String = objForm.DataSources.UserDataSources.Item("UD_DR").Value
                    Dim PettyCashRequest As String = objForm.DataSources.UserDataSources.Item("UD_PCR").Value
                    Dim PettyCashVoucher As String = objForm.DataSources.UserDataSources.Item("UD_PCV").Value
                    Dim MaterialRequest As String = objForm.DataSources.UserDataSources.Item("UD_MR").Value
                    Dim BranchTransfer As String = objForm.DataSources.UserDataSources.Item("UD_BT").Value
                    Dim ORG_1 As String = objForm.DataSources.UserDataSources.Item("ORG_1").Value
                    Dim AUT_1 As String = objForm.DataSources.UserDataSources.Item("AUT_1").Value
                    Dim TEM_1 As String = objForm.DataSources.UserDataSources.Item("TEM_1").Value
                    Dim EMP_1 As String = objForm.DataSources.UserDataSources.Item("EMP_1").Value
                    Dim OVERTIME As String = objForm.DataSources.UserDataSources.Item("UD_OVR").Value
                    Dim ALTERNATEWEEKOFF As String = objForm.DataSources.UserDataSources.Item("UD_AWOF").Value
                    Dim Draftestimation As String = objForm.DataSources.UserDataSources.Item("UD_DRAFT").Value
                    Dim ADVERTISEMENT As String = objForm.DataSources.UserDataSources.Item("UD_VRQ").Value

                    Dim LOADESS As String = objForm.DataSources.UserDataSources.Item("UD_LOADESS").Value
                    Dim fromdate As String = objForm.DataSources.UserDataSources.Item("RDATE_1").Value
                    Dim todate As String = objForm.DataSources.UserDataSources.Item("RDATE_2").Value
                    Dim BusinessPartner As String = objForm.DataSources.UserDataSources.Item("UD_BP").Value
                    Dim SupplierRebate As String = objForm.DataSources.UserDataSources.Item("UD_SR").Value
                    Dim PriceChangeRequest As String = objForm.DataSources.UserDataSources.Item("UD_PCR").Value
                    Dim ItemMasterdataUpload As String = objForm.DataSources.UserDataSources.Item("UD_IMP").Value
                    Dim itemlist As String = objForm.DataSources.UserDataSources.Item("UD_ITEM").Value
                    Dim PACPPriceChangeRequest As String = objForm.DataSources.UserDataSources.Item("UD_PACP").Value

                    ' Fetch selected Branch ID from UD_BNCH ComboBox safely
                    Dim BranchId As String = ""
                    Dim BranchName As String = ""

                    Try
                        Dim oComboBranch As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("UD_BNCH").Specific, SAPbouiCOM.ComboBox)
                        If oComboBranch.Selected IsNot Nothing Then
                            BranchId = oComboBranch.Selected.Value.ToString().Trim()
                            BranchName = oComboBranch.Selected.Description.ToString().Trim()
                        End If
                    Catch ex As Exception
                        BranchId = ""
                        BranchName = ""
                    End Try

                    Dim Query1 As String
                    'If LOADESS = "Y" Then
                    If (Pending = "" Or Pending = "N") And (Rejected = "" Or Rejected = "N") And (Canceled = "" Or Canceled = "N") And (Approved = "" Or Approved = "N") Then
                        objMain.objApplication.StatusBar.SetText("Invalid Selection", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub

                    ElseIf (itemlist = "" Or itemlist = "N") And (RecruitmentRequest = "" Or RecruitmentRequest = "N") And (BusinessPartner = "" Or BusinessPartner = "N") And (SupplierRebate = "" Or SupplierRebate = "N") And (PriceChangeRequest = "" Or PriceChangeRequest = "N") And (ADVERTISEMENT = "" Or ADVERTISEMENT = "N") And (ALTERNATEWEEKOFF = "" Or ALTERNATEWEEKOFF = "N") And (OVERTIME = "" Or OVERTIME = "N") And (BranchTransfer = "" Or BranchTransfer = "N") And (MaterialRequest = "" Or MaterialRequest = "N") And (PettyCashVoucher = "" Or PettyCashVoucher = "N") And (PettyCashRequest = "" Or PettyCashRequest = "N") And (DocumentRequest = "" Or DocumentRequest = "N") And (AirticketRequest = "" Or AirticketRequest = "N") And (ExitClearance = "" Or ExitClearance = "N") And (EmployeeResignationRequest = "" Or EmployeeResignationRequest = "N") And (LeaveEncashment = "" Or LeaveEncashment = "N") And (EmployeeReveiwAppraisal = "" Or EmployeeReveiwAppraisal = "N") And (ClaimApplication = "" Or ClaimApplication = "N") And (PayrollPreProcess = "" Or PayrollPreProcess = "N") And (PaymentandDeduction = "" Or PaymentandDeduction = "N") And (LoanPreClosureApplication = "" Or LoanPreClosureApplication = "N") And (LoansandAdvanceApplication = "" Or LoansandAdvanceApplication = "N") And (LeavePeriodEndClc = "" Or LeavePeriodEndClc = "N") And (LeaveEncashmentrequest = "" Or LeaveEncashmentrequest = "N") And (AttendanceRegularisation = "" Or AttendanceRegularisation = "N") And (ShiftShedule = "" Or ShiftShedule = "N") And (CompensatoryOff = "" Or CompensatoryOff = "N") And (LeaveApplication = "" Or LeaveApplication = "N") And (LeaveRejoinApplication = "" Or LeaveRejoinApplication = "N") And (EndofServiceProcess = "" Or EndofServiceProcess = "N") And (Timesheet = "" Or Timesheet = "N") And (EmployeeContractRenewal1 = "" Or EmployeeContractRenewal1 = "N") And (EmployeeSalarySetup = "" Or EmployeeSalarySetup = "N") And (EmployeeMasterData = "" Or EmployeeMasterData = "N") And (OfferLetter = "" Or OfferLetter = "N") And (ResumeSubmission = "" Or ResumeSubmission = "N") And (AttendanceSummary = "" Or AttendanceSummary = "N") And (Draftestimation = "" Or Draftestimation = "N") And (ItemMasterdataUpload = "" Or ItemMasterdataUpload = "N") And (PACPPriceChangeRequest = "" Or PACPPriceChangeRequest = "N") Then
                        objMain.objApplication.StatusBar.SetText("Please Select Atleast one document", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub

                    Else
                        ' Updated stored procedure call string to pass BranchId as the 12th parameter
                        Query1 = "CALL ""TNX_ApprovalStatusReport""('" & SupplierRebate & "','" & PriceChangeRequest & "','" & ItemMasterdataUpload & "','" & BusinessPartner & "','" & itemlist & "','" & PACPPriceChangeRequest & "','" & Pending & "','" & Approved & "','" & Rejected & "','" & Canceled & "', '" & Submitted & "','" & BranchId & "')"

                        objMain.ObjGRID.CreateForm(Query1)
                        ' objMain.ObjGRID.GRID(Query1)
                        ' objMain.ObjGRID.GRID(FormUID

                    End If
                End If
                objCheckBox1 = objForm.Items.Item("UD_LOADESS").Specific
                If pVal.ItemUID = "UD_LOADESS" And objCheckBox1.Checked = True Then
                    Try
                        objForm.Items.Item("EMP").Visible = True
                        objForm.Items.Item("Item_2").Visible = True


                    Catch EX As Exception
                    End Try
                    'Dim oItem As SAPbouiCOM.Item = objForm.Items.Item("Item_22")

                    'oItem.Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                ElseIf pVal.ItemUID = "UD_LOADESS" And objCheckBox1.Checked = False Then



                    Dim oItem As SAPbouiCOM.Item = objForm.Items.Item("Item_22")

                    oItem.Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    objForm.Items.Item("Item_2").Visible = False
                    objForm.Items.Item("EMP").Visible = False
                    objCheckBox1.Checked = False


                    '  objForm.Items.Item("EMP"). 
                    'Dim oItem As Item = oForm.Items.Item("AnyOtherItem")

                    ' oItem.Click(BoCellClickType.ct_Regular)
                End If
                    'End If

                    '    If objCheckBox1.Checked = True Then
                    '            Try
                    '                objForm.Items.Item("EMP").Visible = False
                    '            Catch EX As Exception
                    '            End Try
                    '            Dim oItem As SAPbouiCOM.Item = objForm.Items.Item("Item_22")

                    '            oItem.Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    '        Else
                    '            objForm.Items.Item("EMP").Visible = True
                    '            'Dim oItem As SAPbouiCOM.Item = objForm.Items.Item("Item_22")

                    '            'oItem.Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    '            '  objForm.Items.Item("EMP"). 
                    '            'Dim oItem As Item = oForm.Items.Item("AnyOtherItem")

                    '            ' oItem.Click(BoCellClickType.ct_Regular)
                    '        End If
                    ''    End Ifxx

              '  End If

            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                objForm = objMain.objApplication.Forms.Item(FormUID)
                Dim oCFL As SAPbouiCOM.ChooseFromList
                Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                Dim CFL_Id As String
                CFL_Id = CFLEvent.ChooseFromListUID
                oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                Dim oDT As SAPbouiCOM.DataTable
                oDT = CFLEvent.SelectedObjects
                objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                ' If oCFL.UniqueID = "CFL_1" And pVal.BeforeAction = True Then
                '    ' Me.CFLFilter("CFL_Sale2")
                'End If
                If (Not oDT Is Nothing) And pVal.BeforeAction = False Then
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                    If oCFL.UniqueID = "CFL_1" Then
                        'objForm.Items.Item("ORG").Specific.value = oDT.GetValue("U_NAME", 0)
                        objForm.DataSources.UserDataSources.Item("ORG_1").Value = oDT.GetValue("USER_CODE", 0)
                    End If



                    If oCFL.UniqueID = "CFL_4" Then
                        objForm.DataSources.UserDataSources.Item("AUT_1").Value = oDT.GetValue("USER_CODE", 0)

                    End If
                    If oCFL.UniqueID = "CFL_5" Then
                        objForm.DataSources.UserDataSources.Item("TEM_1").Value = oDT.GetValue("Code", 0)
                    End If
                    If oCFL.UniqueID = "CFL_6" Then
                        objForm.DataSources.UserDataSources.Item("TEM_2").Value = oDT.GetValue("Code", 0)
                    End If
                    If oCFL.UniqueID = "CFL_7" Then
                        objForm.DataSources.UserDataSources.Item("EMP_1").Value = oDT.GetValue("empID", 0)
                    End If
                    If oCFL.UniqueID = "CFL_8" Then
                        objForm.DataSources.UserDataSources.Item("EMP_2").Value = oDT.GetValue("empID", 0)
                    End If
                End If
        End Select
    End Sub

    Private Sub LoadUDBranchCombo()

        Try

            Dim oCombo As SAPbouiCOM.ComboBox
            oCombo = CType(objForm.Items.Item("UD_BNCH").Specific, SAPbouiCOM.ComboBox)

            'Optional
            'oCombo.ValidValues.Clear()

            Dim oRs As SAPbobsCOM.Recordset =
                objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'oRs.DoQuery("SELECT T2.""BPLId"", T2.""BPLName"" " &
            '            "FROM USR6 T0 " &
            '            "INNER JOIN OUSR T1 ON T0.""UserCode"" = T1.""USER_CODE"" " &
            '            "INNER JOIN OBPL T2 ON T0.""BPLId"" = T2.""BPLId"" " &
            '            "WHERE T1.""USER_CODE"" = '" & objMain.objCompany.UserName.Trim() & "' " &
            '            "ORDER BY T2.""BPLId"" ")
            oRs.DoQuery("SELECT T0.""BPLId"", T0.""BPLName"" " &
            "FROM OBPL T0 " &
            "WHERE T0.""Disabled"" = 'N' " &
            "ORDER BY T0.""BPLId""")

            While Not oRs.EoF

                Dim sBranchId As String = oRs.Fields.Item("BPLId").Value.ToString()
                Dim sBranchName As String = oRs.Fields.Item("BPLName").Value.ToString()

                If Not ComboValueExists(oCombo, sBranchId) Then
                    oCombo.ValidValues.Add(sBranchId, sBranchName)
                End If

                oRs.MoveNext()

            End While

            If oCombo.ValidValues.Count > 0 Then
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            End If

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
                "LoadUDBranchCombo ERROR : " & ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

    Private Function ComboValueExists(ByVal oCombo As SAPbouiCOM.ComboBox, ByVal value As String) As Boolean
        For i As Integer = 0 To oCombo.ValidValues.Count - 1
            If oCombo.ValidValues.Item(i).Value = value Then
                Return True
            End If
        Next
        Return False
    End Function

End Class