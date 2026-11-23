Public Class clsPurchaseBlanketAgreement
    Public objform As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim oItem As SAPbouiCOM.Item
    Dim objMatrix, objGRPOMatrix As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try

            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.BeforeAction = True Then
                        Me.AddItems(objform.UniqueID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "Btn_SpReb" And pVal.BeforeAction = False Then
                        oDBs_Head = objform.DataSources.DBDataSources.Item("OOAT")

                        Dim BpCode As String = oDBs_Head.GetValue("BpCode", 0).Trim()
                        Dim DocEntry As String = oDBs_Head.GetValue("AbsID", 0).Trim()

                        Dim CheckIfTCExist As String = "Select " & IsNull & "(""U_SUPREBT"",'') ""U_SUPREBT"" From ""OOAT"" Where ""Number""='" & oDBs_Head.GetValue("Number", 0) & "' And  IFNULL(""U_SUPREBT"",'') <> '' And ""Series"" = '" & oDBs_Head.GetValue("Series", 0) & "'"
                        Dim oRsCheckIfTCExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsCheckIfTCExist.DoQuery(CheckIfTCExist)
                        If oRsCheckIfTCExist.Fields.Item("U_SUPREBT").Value <> "" Then
                            objMain.objSupplierRebate.CreateForm(objform.UniqueID, oDBs_Head.GetValue("BpCode", 0), oDBs_Head.GetValue("BpName", 0), oDBs_Head.GetValue("AbsID", 0), oDBs_Head.GetValue("Number", 0), "Y", oRsCheckIfTCExist.Fields.Item("U_SUPREBT").Value)
                            'objMain.objSupplierRebate2.CreateForm(objform.UniqueID, oDBs_Head.GetValue("BpCode", 0), oDBs_Head.GetValue("BpName", 0), oDBs_Head.GetValue("AbsID", 0), oDBs_Head.GetValue("Number", 0), "Y", oRsCheckIfTCExist.Fields.Item("U_SUPREBT").Value)

                        Else
                            objMain.objSupplierRebate.CreateForm(objform.UniqueID, oDBs_Head.GetValue("BpCode", 0), oDBs_Head.GetValue("BpName", 0), oDBs_Head.GetValue("AbsID", 0), oDBs_Head.GetValue("Number", 0), "N")
                            'objMain.objSupplierRebate2.CreateForm(objform.UniqueID, oDBs_Head.GetValue("BpCode", 0), oDBs_Head.GetValue("BpName", 0), oDBs_Head.GetValue("AbsID", 0), oDBs_Head.GetValue("Number", 0), "N")


                        End If

                    End If
                    '----Nareshv1
                    If pVal.ItemUID = "Btn_Sup" And pVal.BeforeAction = False Then
                        oDBs_Head = objform.DataSources.DBDataSources.Item("OOAT")
                        Dim CheckIfTCExist As String = "Select " & IsNull & "(""U_SUPPOST"",'') ""U_SUPPOST"" From ""OOAT"" Where ""Number""='" & oDBs_Head.GetValue("Number", 0) & "' And  IFNULL(""U_SUPPOST"",'') <> '' "
                        Dim oRsCheckIfTCExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsCheckIfTCExist.DoQuery(CheckIfTCExist)
                        If oRsCheckIfTCExist.Fields.Item("U_SUPPOST").Value <> "" Then
                            objMain.objSupplAgreememntPosting.CreateForm(objform.UniqueID, oDBs_Head.GetValue("BpCode", 0), oDBs_Head.GetValue("BpName", 0), oDBs_Head.GetValue("AbsID", 0), oDBs_Head.GetValue("Number", 0), "Y", oRsCheckIfTCExist.Fields.Item("U_SUPPOST").Value, "BlanketAgreement")
                        Else
                            objMain.objSupplAgreememntPosting.CreateForm(objform.UniqueID, oDBs_Head.GetValue("BpCode", 0), oDBs_Head.GetValue("BpName", 0), oDBs_Head.GetValue("AbsID", 0), oDBs_Head.GetValue("Number", 0), "N", "", "BlanketAgreement")
                        End If
                    End If

                    'Madhu opening item lising screen
                    If pVal.ItemUID = "ItmList" And pVal.BeforeAction = False Then
                        oDBs_Head = objform.DataSources.DBDataSources.Item("OCRD")

                        'madhu
                        If Me.ValidateItemListingButton() = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If

                        ' Dim objDocCombo As SAPbouiCOM.ComboBox = objform.Items.Item("SupplierSta").Specific

                        Dim CheckIfTCExist As String = "Select " & IsNull & "(""U_ITMLIST"",'') ""U_ITMLIST"" From ""OCRD"" Where ""CardCode""='" & oDBs_Head.GetValue("CardCode", 0) & "'"
                        'Dim oRsCheckIfTCExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'oRsCheckIfTCExist.DoQuery(CheckIfTCExist)
                        'If oRsCheckIfTCExist.Fields.Item("U_ITMLIST").Value <> "" Then
                        '    objMain.objItemList.CreateForm(objform.UniqueID, oDBs_Head.GetValue("CardCode", 0), oDBs_Head.GetValue("CardName", 0), "Y")
                        'Else
                        '    objMain.objItemList.CreateForm(objform.UniqueID, oDBs_Head.GetValue("CardCode", 0), oDBs_Head.GetValue("CardName", 0), "N")
                        'End If
                        CheckIfTCExist = " CALL  ""TNX_ITEMLLIST""('" & oDBs_Head.GetValue("CardCode", 0) & "')"
                        Dim oRsCheckIfTCExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim CARDCODEX As String = oDBs_Head.GetValue("CardCode", 0)
                        oRsCheckIfTCExist.DoQuery(CheckIfTCExist)
                        '  If oRsCheckIfTCExist.Fields.Item("U_ITMLIST").Value <> "" Then

                        If oRsCheckIfTCExist.RecordCount > 0 Then
                            objMain.objApplication.StatusBar.SetText($"Item list found for {CARDCODEX}. Opening list...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                            objMain.objDocumentDraw_ITM.CreateForm(objform.UniqueID, oDBs_Head.GetValue("CardCode", 0), oDBs_Head.GetValue("CardName", 0))
                            objform.Items.Item("ItmList").Enabled = True

                            ' objMain.objItemList.CreateForm(objform.UniqueID, oDBs_Head.GetValue("CardCode", 0), oDBs_Head.GetValue("CardName", 0), "Y")
                        Else
                            objMain.objApplication.StatusBar.SetText($"No item list available for {CARDCODEX}. Creating new item list...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                            '  objMain.objDocumentDraw_ITM.CreateForm(objform.UniqueID, oDBs_Head.GetValue("CardCode", 0), oDBs_Head.GetValue("CardName", 0))
                            objMain.objItemList.CreateForm(objform.UniqueID, oDBs_Head.GetValue("CardCode", 0), oDBs_Head.GetValue("CardName", 0), "N")
                            objform.Items.Item("ItmList").Enabled = True
                        End If
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRsCheckIfTCExist)
                        oRsCheckIfTCExist = Nothing


                    End If


                    'If pVal.ItemUID = "1250000026" And pVal.BeforeAction = False Then
                    '    Me.AddChooseFromListItms(FormUID, "6")
                    'End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                        Dim CFL_Id As String
                        CFL_Id = CFLEvent.ChooseFromListUID
                        oCFL = objform.ChooseFromLists.Item(CFL_Id)
                        Dim oDT As SAPbouiCOM.DataTable
                        oDT = CFLEvent.SelectedObjects
                        objform = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        If oCFL.UniqueID = "10" And pVal.BeforeAction = True Then
                            Try
                                Dim Query As String = " SELECT T1.""ItemCode"", T1.""VendorCode"" FROM OITM T0  INNER JOIN ITM2 T1 ON T0.""ItemCode"" = T1.""ItemCode"" WHERE T1.""VendorCode""='" & objform.Items.Item("1250000006").Specific.Value.ToString().Trim() & "'"
                                Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsQuery.DoQuery(Query)

                                Dim oConditions As SAPbouiCOM.Conditions
                                Dim oCondition As SAPbouiCOM.Condition
                                Dim oChooseFromList As SAPbouiCOM.ChooseFromList
                                Dim emptyCon As New SAPbouiCOM.Conditions
                                oChooseFromList = objform.ChooseFromLists.Item("6")
                                oChooseFromList.SetConditions(emptyCon)
                                oConditions = oChooseFromList.GetConditions()

                                For i As Integer = 1 To oRsQuery.RecordCount
                                    oCondition = oConditions.Add()
                                    oCondition.Alias = "ItemCode"
                                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                    oCondition.CondVal = oRsQuery.Fields.Item(0).Value.ToString().Trim
                                    If i <> oRsQuery.RecordCount Then
                                        oCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                                    End If
                                    oRsQuery.MoveNext()
                                Next
                                oChooseFromList.SetConditions(oConditions)

                            Catch ex As Exception

                            End Try

                        End If
                    Catch ex As Exception

                    End Try
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub AddItems(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)

            objMain.objUtilities.AddButtonH(objform.UniqueID, "Btn_SpReb", objform.Items.Item("1250000002").Top, objform.Items.Item("1250000002").Left + objform.Items.Item("1250000002").Width + 5, objform.Items.Item("1250000002").Width + 20, "1250000002", "Supplier &Rebate", 0, 0, objform.Items.Item("1250000002").Height)


            objform.Items.Item("Btn_SpReb").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objform.Items.Item("Btn_SpReb").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objform.Items.Item("Btn_SpReb").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            '-----NareshV1
            objMain.objUtilities.AddButtonH(objform.UniqueID, "Btn_Sup", objform.Items.Item("1250000002").Top, objform.Items.Item("1250000002").Left + objform.Items.Item("1250000002").Width + objform.Items.Item("Btn_SpReb").Width + 10, objform.Items.Item("1250000002").Width + 60, "1250000002", "Supplier &Charges", 0, 0, objform.Items.Item("1250000002").Height)
            '------NareshV1

            'Madhu ItemListing Buuton
            'objMain.objUtilities.AddButton(objform.UniqueID, "ItmList", objform.Items.Item("540002072").Top - 50, objform.Items.Item("540002072").Left - 130, objform.Items.Item("137").Width, objform.Items.Item("137").Height, "540002072", "Item &Listing", 46, 46)

            objMain.objUtilities.AddButtonH(objform.UniqueID, "ItmList", objform.Items.Item("Btn_Sup").Top, objform.Items.Item("Btn_Sup").Left + objform.Items.Item("Btn_Sup").Width + 10, objform.Items.Item("Btn_Sup").Width, "Btn_Sup", "Item &Listing", 0, 0, objform.Items.Item("Btn_Sup").Height)


            objform.Items.Item("Btn_Sup").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objform.Items.Item("Btn_Sup").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objform.Items.Item("Btn_Sup").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            '--- Target Amount Label
            ' objMain.objUtilities.AddLabel(objform.UniqueID, "lbTAmt", objform.Items.Item("1250000016").Top + 70,
            'objform.Items.Item("1250000016").Left - 169, objform.Items.Item("1250000016").Width, "Target Amount", "&Target Amount", 0, 0, objform.Items.Item("1250000016").Height)

            ' objMain.objUtilities.AddEditBox(objform.UniqueID, "txtTAmt", objform.Items.Item("1720000067").Top, objform.Items.Item("1720000067").Left + objform.Items.Item("1720000067").Width + 1030, 140,
            ' "OOAT", "U_TAmount", "1720000067", objform.Items.Item("1720000067").Height, objform.Items.Item("1720000067").FromPane, objform.Items.Item("1720000067").ToPane)

            objMain.objUtilities.AddLabel(objform.UniqueID, "lbTAmt", objform.Items.Item("1250000016").Top + 18,
         objform.Items.Item("1250000016").Left - 180, objform.Items.Item("1250000016").Width, "Target Amount", "&Target Amount", 0, 0, objform.Items.Item("1250000016").Height)

            objMain.objUtilities.AddEditBox(objform.UniqueID, "txtTAmt", objform.Items.Item("1720000067").Top - 52, objform.Items.Item("1720000067").Left + objform.Items.Item("1720000067").Width - 144, +144,
            "OOAT", "U_TAmount", "1720000067", objform.Items.Item("1720000067").Height, objform.Items.Item("1720000067").FromPane, objform.Items.Item("1720000067").ToPane)


            'oItem = objform.Items.Add("fldSale", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            'oItem.Top = objform.Items.Item("1320000072").Top
            'oItem.Left = objform.Items.Item("1320000072").Left - objform.Items.Item("1320000072").Width + 10
            'oItem.Width = objform.Items.Item("1320000072").Width + 30
            'oItem.Height = objform.Items.Item("1320000072").Height
            'oItem.AffectsFormMode = False
            'Dim fld As SAPbouiCOM.Folder = objform.Items.Item("fldSale").Specific
            'fld.Caption = "Average"
            'fld.GroupWith("1320000072")
            'fld.Pane = "1"

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objform = objMain.objApplication.Forms.GetForm("1250000102", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then

                        objMain.objApplication.MessageBox("Purchase Blanketagreement Document is Succesfully Added")
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    'Private Function AddChooseFromListItms(ByVal FormUID As String, ByVal CFL_ID As String) As Boolean
    '    Dim bOk As Boolean = True
    '    Try
    '        objform = objMain.objApplication.Forms.Item(FormUID)


    '        'Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
    '        'Dim oCons As SAPbouiCOM.Conditions
    '        'Dim oCon As SAPbouiCOM.Condition
    '        'oCFLs = objform.ChooseFromLists
    '        'Dim oCFL As SAPbouiCOM.ChooseFromList
    '        'Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
    '        'oCFLCreationParams = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
    '        'oCFLCreationParams.MultiSelection = False
    '        'oCFLCreationParams.ObjectType = "5"
    '        'oCFLCreationParams.UniqueID = CFL_ID
    '        'oCFL = oCFLs.Add(oCFLCreationParams)
    '        'oCons = oCFL.GetConditions()
    '        'oCon = oCons.Add()
    '        'If oCFL.UniqueID = "10" And pVal.BeforeAction = True Then
    '        '    Try
    '        Dim Query As String = " SELECT T1.""ItemCode"", T1.""VendorCode"" FROM OITM T0  INNER JOIN ITM2 T1 ON T0.""ItemCode"" = T1.""ItemCode"" WHERE T1.""VendorCode""='" & objform.Items.Item("1250000006").Specific.Value.ToString().Trim() & "'"
    '        Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oRsQuery.DoQuery(Query)

    '        Dim oConditions As SAPbouiCOM.Conditions
    '        Dim oCondition As SAPbouiCOM.Condition
    '        Dim oChooseFromList As SAPbouiCOM.ChooseFromList
    '        Dim emptyCon As New SAPbouiCOM.Conditions
    '        oChooseFromList = objform.ChooseFromLists.Item("6")
    '        oChooseFromList.SetConditions(emptyCon)
    '        oConditions = oChooseFromList.GetConditions()

    '        For i As Integer = 1 To oRsQuery.RecordCount
    '            oCondition = oConditions.Add()
    '            oCondition.Alias = "ItemCode"
    '            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '            oCondition.CondVal = oRsQuery.Fields.Item(0).Value.ToString().Trim
    '            If i <> oRsQuery.RecordCount Then
    '                oCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
    '            End If
    '            oRsQuery.MoveNext()
    '        Next
    '        oChooseFromList.SetConditions(oConditions)

    '    Catch ex As Exception

    '    End Try

    '    '    End If

    '    'Catch ex As Exception
    '    bOk = False
    '    ' objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '    ' End Try
    '    Return bOk
    'End Function

    Public Function ValidateItemListingButton() As Boolean
        Try
            oDBs_Head = objform.DataSources.DBDataSources.Item("OCRD")
            Dim CardCode As String = oDBs_Head.GetValue("CardCode", 0).Trim()
            If CardCode = "" Then
                objMain.objApplication.SetStatusBarMessage("Please select a cardcode", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Return True

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            ex.Message & " - Errors in Validation Function",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function


End Class
