Imports System.Security.Cryptography
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports SAPbobsCOM

Public Class clsBusinessPartners
    Public objform As SAPbouiCOM.Form
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oItem As SAPbouiCOM.Item
    Dim objComboBox As SAPbouiCOM.ComboBox
    Public objsubform As SAPbouiCOM.Form

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
                    If pVal.ItemUID = "Btn_SReg" And pVal.BeforeAction = False Then
                        oDBs_Head = objform.DataSources.DBDataSources.Item("OCRD")
                        Dim CheckIfTCExist As String = "Select IFNULL(""U_SUPREGDE"",'')""U_SUPREGDE"" From ""OCRD"" Where ""CardCode""='" & oDBs_Head.GetValue("CardCode", 0) & "'"
                        Dim oRsCheckIfTCExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsCheckIfTCExist.DoQuery(CheckIfTCExist)
                        '-----NareshV1--------------------
                        'If oRsCheckIfTCExist.Fields.Item("U_SUPREGDE").Value <> "" Then
                        '    objMain.objSupplierAgreement.CreateForm(objform.UniqueID, oDBs_Head.GetValue("CardCode", 0), oDBs_Head.GetValue("CardName", 0), "Y", oRsCheckIfTCExist.Fields.Item("U_SUPREGDE").Value)
                        'Else
                        'objMain.objSupplierAgreement.CreateForm(objform.UniqueID, oDBs_Head.GetValue("CardCode", 0), oDBs_Head.GetValue("CardName", 0), "N")
                        'End If

                        objMain.objLoadAgreementsList.CreateForm(objform.UniqueID, oDBs_Head.GetValue("CardCode", 0), "Y")
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRsCheckIfTCExist)
                        oRsCheckIfTCExist = Nothing
                        '-----NareshV1----------------------
                    End If

                    'Madhu
                    If pVal.ItemUID = "Submit" And pVal.BeforeAction = False Then
                        oDBs_Head = objform.DataSources.DBDataSources.Item("OCRD")
                        Dim cardCode As String = oDBs_Head.GetValue("CardCode", 0).Trim().Replace("'", "''")

                        Dim CheckIfTCExist As String = "Select ""U_Status"" As ""Status"" From ""OCRD"" Where ""CardCode""='" & cardCode & "' And ""CardType"" ='S'"
                        Dim oRsCheckIfTCExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsCheckIfTCExist.DoQuery(CheckIfTCExist)
                        If oRsCheckIfTCExist.Fields.Item("Status").Value = "O" Then
                            Dim AppID As String = oDBs_Head.GetValue("CardCode", 0)

                            Me.ApprovalTrigger("BP", AppID, "OCRD", "CardCode", "U_Status", "BusinessPartner")
                        End If
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRsCheckIfTCExist)
                        oRsCheckIfTCExist = Nothing

                    End If



                    'madhu
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objform.DataSources.DBDataSources.Item("OCRD")
                    Dim cardCode As String = oDBs_Head.GetValue("CardCode", 0).Trim().Replace("'", "''")

                    Dim BPStatus As String = "Select ""U_Status"" As ""Status"" From ""OCRD"" Where ""CardCode""='" & cardCode & "'And ""CardType"" ='S'"
                    Dim oRsBPStatus As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsBPStatus.DoQuery(BPStatus)

                    Dim rsapp As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim str As String = "SELECT T0.""Code"" As ""TemplateID"",T1.""U_Name"" As ""Originator"",S1.""Code"" As ""Stage"",S2.""U_UKey"" As ""Authorizers"" " &
                    " ,S1.""U_NAP"" As ""NoOfAuth"",S1.""U_NRJ"" As ""NoOfRej"",T0.""CreateDate"",(Case When T3.""LineId""='1' Then 'S' Else 'O' End) As ""Status"" FROM ""@SBO_APPHDR"" T0 " &
                    " INNER JOIN ""@SBO_APPREQ"" T1 ON T1.""Code""=T0.""Code"" " &
                    " INNER JOIN ""@SBO_APPDOC"" T2 ON T2.""Code"" = T0.""Code"" " &
                    " INNER JOIN ""@SBO_APPAUT"" T3 ON T3.""Code"" = T0.""Code"" " &
                    " INNER JOIN ""@SBO_AST"" S1 ON T3.""U_NAMES""=S1.""Code"" " &
                    " INNER JOIN ""@SBO_AST_C0"" S2 ON S1.""Code""=S2.""Code"" " &
                    " WHERE ""U_Active"" = 'Y' AND S2.""U_UKey"" = '" & objMain.objCompany.UserName & "' AND T2.""U_BP""='Y' AND IFNULL(S2.""U_UKey"",'')<>''  order by T3.""LineId""  ;"
                    rsapp.DoQuery(str)

                    If rsapp.RecordCount > 0 Then

                    End If
                    Dim DocType As String = "BP"
                    Dim rsapp1 As SAPbobsCOM.Recordset
                    rsapp1 = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    Dim str1 As String = "SELECT T0.""Code"" As ""TemplateID"",T1.""U_Name"" As ""Originator"",S1.""Code"" As ""Stage"",S2.""U_UKey"" As ""Authorizers"" " &
            " ,S1.""U_NAP"" As ""NoOfAuth"",S1.""U_NRJ"" As ""NoOfRej"",T0.""CreateDate"",(Case When T3.""LineId""='1' Then 'S' Else 'O' End) As ""Status"" FROM ""@SBO_APPHDR"" T0 " &
            " INNER JOIN ""@SBO_APPREQ"" T1 ON T1.""Code""=T0.""Code"" " &
            " INNER JOIN ""@SBO_APPDOC"" T2 ON T2.""Code"" = T0.""Code"" " & " INNER JOIN ""@SBO_APPAUT"" T3 ON T3.""Code"" = T0.""Code"" " &
            " INNER JOIN ""@SBO_AST"" S1 ON T3.""U_M3_1""=S1.""Code"" " & " INNER JOIN ""@SBO_AST_C0"" S2 ON S1.""Code""=S2.""Code"" " &
            " WHERE ""U_Active"" = 'Y' " & " AND T1.""U_Name"" = '" & objMain.objCompany.UserName & "' " &
            " AND T2.""U_" & DocType & """ = 'Y' " & " AND IFNULL(S2.""U_UKey"",'')<>'' " & " ORDER BY T3.""LineId"""

                    rsapp1.DoQuery(str1)

                    If rsapp1.RecordCount < 0 Then
                        If (pVal.ItemUID = "10002044" Or pVal.ItemUID = "10002046") And pVal.BeforeAction = True And objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            BubbleEvent = False
                            objMain.objApplication.StatusBar.SetText("You cannot change the status, please submit for approval", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                    End If

                    If (pVal.ItemUID = "10002044" Or pVal.ItemUID = "10002046") And pVal.BeforeAction = True And (oRsBPStatus.Fields.Item("Status").Value = "O") And (objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        BubbleEvent = False
                        objMain.objApplication.StatusBar.SetText("You cannot change the status, Please submit it for approval", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                    If (pVal.ItemUID = "10002044" Or pVal.ItemUID = "10002046") And pVal.BeforeAction = True And rsapp.RecordCount = 0 And (oRsBPStatus.Fields.Item("Status").Value = "S") And (objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        BubbleEvent = False
                        objMain.objApplication.StatusBar.SetText("You cannot change the status, It is submitted for approval", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                    If (pVal.ItemUID = "1") And pVal.BeforeAction = True And rsapp.RecordCount = 0 And (oRsBPStatus.Fields.Item("Status").Value = "S") And (objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        BubbleEvent = False
                        objMain.objApplication.StatusBar.SetText("You cannot update the Business partner, After approval only you can update the business partner", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                    If (pVal.ItemUID = "1") And pVal.BeforeAction = True And (oRsBPStatus.Fields.Item("Status").Value = "N") And (objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        BubbleEvent = False
                        objMain.objApplication.StatusBar.SetText("You cannot update the Business partner, It is rejected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                    'madhu

                    If pVal.ItemUID = "ItmList" And pVal.BeforeAction = False Then
                        Try


                            oDBs_Head = objform.DataSources.DBDataSources.Item("OCRD")
                            objform.Items.Item("ItmList").Enabled = False
                            'madhu
                            If Me.ValidateItemListingButton() = False Then
                                BubbleEvent = False
                                Exit Sub
                            End If

                            ' Dim objDocCombo As SAPbouiCOM.ComboBox = objform.Items.Item("SupplierSta").Specific

                            Dim CheckIfTCExist As String = "Select " & IsNull & "(""U_ITMLIST"",'') ""U_ITMLIST"" From ""OCRD"" Where ""CardCode""='" & oDBs_Head.GetValue("CardCode", 0) & "'"
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

                        Catch ex As Exception
                            objform.Items.Item("ItmList").Enabled = True

                        End Try
                    End If

                    If pVal.ItemUID = "fldRtl" Then
                        objform.PaneLevel = "46"
                        'Me.Matter(objform.UniqueID)
                    End If
                    If pVal.ItemUID = "Btn_Sup" And pVal.BeforeAction = False Then
                        oDBs_Head = objform.DataSources.DBDataSources.Item("OCRD")
                        Dim CheckIfTCExist As String = "Select " & IsNull & "(""U_SUPNUM"",'') ""U_SUPNUM"" From ""OCRD"" Where ""CardCode""='" & oDBs_Head.GetValue("CardCode", 0) & "' And  IFNULL(""U_SUPNUM"",'') <> '' "
                        Dim oRsCheckIfTCExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsCheckIfTCExist.DoQuery(CheckIfTCExist)
                        If oRsCheckIfTCExist.Fields.Item("U_SUPNUM").Value <> "" Then
                            objMain.objSupplAgreememntPosting.CreateForm(objform.UniqueID, oDBs_Head.GetValue("CardCode", 0), oDBs_Head.GetValue("CardName", 0), "", "", "Y", oRsCheckIfTCExist.Fields.Item("U_SUPNUM").Value, "BussinessPartner")
                        Else
                            objMain.objSupplAgreememntPosting.CreateForm(objform.UniqueID, oDBs_Head.GetValue("CardCode", 0), oDBs_Head.GetValue("CardName", 0), "", "", "N", "", "BussinessPartner")
                        End If
                    End If
            End Select
        Catch ex As Exception
            objform.Items.Item("ItmList").Enabled = True
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
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

    Private Sub AddItems(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)

            oItem = objform.Items.Add("fldRtl", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem.Top = objform.Items.Item("1320002081").Top
            oItem.Left = objform.Items.Item("1320002081").Left + objform.Items.Item("1320002081").Width + 10
            oItem.Width = objform.Items.Item("1320002081").Width + 30
            oItem.Height = objform.Items.Item("1320002081").Height
            oItem.AffectsFormMode = False
            Dim fld As SAPbouiCOM.Folder = objform.Items.Item("fldRtl").Specific
            fld.Caption = "Retail"
            fld.GroupWith("1320002081")
            fld.Pane = "46"
            'objMain.objUtilities.AddButtonH(objform.UniqueID, "Btn_SReg", objform.Items.Item("2").Top, objform.Items.Item("2").Left + objform.Items.Item("2").Width + 5, objform.Items.Item("2").Width + 30, "2", "Supplier &Contract", 0, 0, objform.Items.Item("2").Height)

            'oItem = objform.Items.Add("ItmList", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'oItem.Top = objform.Items.Item("2").Top
            'oItem.Left = objform.Items.Item("2").Left + objform.Items.Item("2").Width + 10 + objform.Items.Item("Btn_SReg").Width
            'oItem.Width = objform.Items.Item("2").Width + 30
            'oItem.Height = objform.Items.Item("2").Height
            'oItem.FromPane = objform.Items.Item("2").FromPane
            'oItem.ToPane = objform.Items.Item("2").ToPane
            'Dim btn1 As SAPbouiCOM.Button = objform.Items.Item("ItmList").Specific
            'btn1.Caption = "Item &Listing"
            'objform.Items.Item("ItmList").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            ' objform.Items.Item("Btn_SReg").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            'objMain.objUtilities.AddLabel(objform.UniqueID, "lblSuppSt", objform.Items.Item("42").Top + 22, objform.Items.Item("42").Left, objform.Items.Item("42").Width, "Supplier Status", "42", objform.Items.Item("42").Height, objform.Items.Item("42").FromPane, objform.Items.Item("42").ToPane)
            'objMain.objUtilities.AddComboBox(objform.UniqueID, "SuppSt", objform.Items.Item("41").Top + 22, objform.Items.Item("41").Left, objform.Items.Item("41").Width, objform.Items.Item("41").Height, "OCRD", "U_SuppSt", objform.Items.Item("41").FromPane, objform.Items.Item("41").ToPane)
            'Dim objBtn As SAPbouiCOM.ComboBox = objform.Items.Item("SuppSt").Specific
            'objBtn.ValidValues.Add("Lead", "Lead")
            'objBtn.ValidValues.Add("Sent to Finance", "Sent to Finance")
            'objBtn.ValidValues.Add("Workflow in progress", "Workflow in progress")
            'objBtn.ValidValues.Add("Approved", "Approved")

            objMain.objUtilities.AddLabel(objform.UniqueID, "lblDis", objform.Items.Item("78").Top, objform.Items.Item("78").Left + 5, objform.Items.Item("78").Width, "Line Discount", "84", objform.Items.Item("78").Height, 46, 46)
            objMain.objUtilities.AddEditBox(objform.UniqueID, "Discount", objform.Items.Item("75").Top, objform.Items.Item("75").Left + 10, objform.Items.Item("75").Width, "OCRD", "U_Discount", "lblDis", objform.Items.Item("75").Height, 46, 46)

            objMain.objUtilities.AddLabel(objform.UniqueID, "lblSup", objform.Items.Item("78").Top + 22, objform.Items.Item("78").Left + 5, objform.Items.Item("78").Width + 13, "Supplier Aggrement Number", "Discount", objform.Items.Item("78").Height, 46, 46)
            objMain.objUtilities.AddEditBox(objform.UniqueID, "SUPREGDE", objform.Items.Item("75").Top + 22, objform.Items.Item("75").Left + 10, objform.Items.Item("75").Width, "OCRD", "U_SUPREGDE", "lblSup", objform.Items.Item("75").Height, 46, 46)

            objMain.objUtilities.AddLabel(objform.UniqueID, "lblItm", objform.Items.Item("78").Top + 44, objform.Items.Item("78").Left + 5, objform.Items.Item("78").Width + 13, "Item List Number", "Discount", objform.Items.Item("78").Height, 46, 46)
            objMain.objUtilities.AddEditBox(objform.UniqueID, "ITMLIST", objform.Items.Item("75").Top + 44, objform.Items.Item("75").Left + 10, objform.Items.Item("75").Width, "OCRD", "U_ITMLIST", "lblItm", objform.Items.Item("75").Height, 46, 46)

            objMain.objUtilities.AddLabel(objform.UniqueID, "lblPO", objform.Items.Item("78").Top + 66, objform.Items.Item("78").Left + 5, objform.Items.Item("78").Width + 13, "PO Auto Close Date", "Discount", objform.Items.Item("78").Height, 46, 46)
            objMain.objUtilities.AddEditBox(objform.UniqueID, "PDATE", objform.Items.Item("75").Top + 66, objform.Items.Item("75").Left + 10, objform.Items.Item("75").Width, "OCRD", "U_PDATE", "lblPO", objform.Items.Item("75").Height, 46, 46)

            objMain.objUtilities.AddButton(objform.UniqueID, "ItmList", objform.Items.Item("540002072").Top - 50, objform.Items.Item("540002072").Left - 130, objform.Items.Item("137").Width, objform.Items.Item("137").Height, "540002072", "Item &Listing", 46, 46)
            objMain.objUtilities.AddButton(objform.UniqueID, "Btn_SReg", objform.Items.Item("540002072").Top - 50, objform.Items.Item("540002072").Left, objform.Items.Item("137").Width, objform.Items.Item("137").Height, "540002072", "Supplier &Contract", 46, 46)

            objMain.objUtilities.AddButton(objform.UniqueID, "Btn_Sup", objform.Items.Item("540002072").Top - 50, objform.Items.Item("540002072").Left - 260, objform.Items.Item("137").Width, objform.Items.Item("137").Height, "540002072", "Supplier &Charges", 46, 46)

            objMain.objUtilities.AddButton(objform.UniqueID, "Submit", objform.Items.Item("2").Top, objform.Items.Item("2").Left + objform.Items.Item("2").Width + 5, objform.Items.Item("2").Width, objform.Items.Item("2").Height, "2", "Submit", 0, 0)

            '---madhu payment Terms tab added fields
            objMain.objUtilities.AddLabel(objform.UniqueID, "Doc", objform.Items.Item("231000009").Top - 128,
            objform.Items.Item("231000009").Left, objform.Items.Item("231000009").Width, "Document Date", "&Document Date", objform.Items.Item("231000009").Height, objform.Items.Item("231000009").FromPane, objform.Items.Item("282").ToPane)
            objMain.objUtilities.AddEditBox(objform.UniqueID, "DocDate", objform.Items.Item("231000010").Top + 18,
            objform.Items.Item("231000010").Left, objform.Items.Item("280").Width + 144, "OCRD", "U_DocDate", "231000010", objform.Items.Item("231000010").Height, 6, 6)


            objMain.objUtilities.AddLabel(objform.UniqueID, "lblMOP", objform.Items.Item("231000009").Top - 110,
                objform.Items.Item("231000009").Left, objform.Items.Item("231000009").Width, "Mode of Payment",
                "&Mode of Payment", objform.Items.Item("231000009").Height, objform.Items.Item("231000009").FromPane, objform.Items.Item("282").ToPane)

            objMain.objUtilities.AddComboBox(objform.UniqueID, "cbxMOP", objform.Items.Item("231000010").Top - 110, objform.Items.Item("231000010").Left, objform.Items.Item("280").Width + 144, objform.Items.Item("231000010").Height, "OCRD", "U_MOPay", 6, 6)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)

        Try

            Dim objform As SAPbouiCOM.Form = Nothing
            Dim objsubform As SAPbouiCOM.Form = Nothing

            Try
                objform = objMain.objApplication.Forms.GetForm("134", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Catch
                Exit Sub
            End Try

            Select Case BusinessObjectInfo.EventType

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                    If BusinessObjectInfo.BeforeAction = False AndAlso BusinessObjectInfo.ActionSuccess = True Then

                        'Get Sub Form (-134)
                        For i As Integer = 0 To objMain.objApplication.Forms.Count - 1

                            Dim oFrm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(i)

                            If oFrm.TypeEx = "-134" Then
                                objsubform = oFrm
                                Exit For
                            End If

                        Next

                        Try

                            Dim objComboBox As SAPbouiCOM.ComboBox = objform.Items.Item("40").Specific
                            If objComboBox.Value = "C" Then

                                objform.Items.Item("Btn_SReg").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objform.Items.Item("ItmList").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            Else

                                objform.Items.Item("Btn_SReg").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objform.Items.Item("ItmList").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            End If

                        Catch
                        End Try

                        'Submit Button Visibility
                        Try
                            Me.CheckSubmitButtonVisibility(BusinessObjectInfo.FormUID)
                        Catch
                        End Try

                        'Disbaled  Status is submitted and Approved
                        Dim status As String = ""

                        Try
                            status = oDBs_Head.GetValue("U_Status", 0).Trim().ToUpper()
                        Catch
                        End Try

                        status = status.Replace(" ", "")

                        If status = "APPROVED" OrElse status = "A" OrElse status = "SUBMITTED" OrElse status = "S" OrElse status = "REJECTED" OrElse status = "R" Then

                            Try : objform.Items.Item("Submit").Enabled = False : Catch : End Try

                            Try
                                objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                            Catch
                            End Try
                            ' Disable all items
                            For i As Integer = 0 To objform.Items.Count - 1
                                Try
                                    objform.Items.Item(i).Enabled = False
                                Catch
                                End Try
                            Next

                            If objsubform IsNot Nothing Then

                                'Disable Fields
                                objsubform.Items.Item("U_POCDT").Enabled = False
                                objsubform.Items.Item("U_PDATE").Enabled = False
                                objsubform.Items.Item("U_SUPNUM").Enabled = False
                                objsubform.Items.Item("U_Discount").Enabled = False
                                objsubform.Items.Item("U_SUPREGDE").Enabled = False
                                objsubform.Items.Item("U_ITMLIST").Enabled = False
                                objsubform.Items.Item("U_PDTPassword").Enabled = False
                                objsubform.Items.Item("U_AppID").Enabled = False
                                objsubform.Items.Item("U_DocDate").Enabled = False

                                'Only Status Enabled
                                objsubform.Items.Item("U_Status").Enabled = True

                            End If

                            'Enable System Buttons
                            Try : objform.Items.Item("1").Enabled = True : Catch : End Try
                            Try : objform.Items.Item("2").Enabled = True : Catch : End Try

                            'Enable Menus
                            Try : objform.EnableMenu("1281", True) : Catch : End Try
                            Try : objform.EnableMenu("1282", True) : Catch : End Try

                        Else

                            'Normal Mode
                            Try : objform.Items.Item("Submit").Enabled = True : Catch : End Try

                            Try : objform.EnableMenu("1281", True) : Catch : End Try
                            Try : objform.EnableMenu("1282", True) : Catch : End Try

                            objsubform.Items.Item("U_POCDT").Enabled = True
                            objsubform.Items.Item("U_PDATE").Enabled = True
                            objsubform.Items.Item("U_SUPNUM").Enabled = True
                            objsubform.Items.Item("U_Discount").Enabled = True
                            objsubform.Items.Item("U_SUPREGDE").Enabled = True
                            objsubform.Items.Item("U_ITMLIST").Enabled = True
                            objsubform.Items.Item("U_PDTPassword").Enabled = True
                            objsubform.Items.Item("U_AppID").Enabled = True
                            objsubform.Items.Item("U_DocDate").Enabled = True
                        End If

                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                    If BusinessObjectInfo.BeforeAction = False AndAlso BusinessObjectInfo.ActionSuccess = True Then

                        objMain.objApplication.MessageBox("Business Partner is Successfully Added")

                    End If

            End Select

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub
    'Madhu
    Sub ApprovalTrigger(ByVal DocType As String, ByVal AppID As String, ByVal Table As String, ByVal AppIDField As String, ByVal AppStatField As String, ByVal Document As String)
        Dim ff As Integer
        Dim draftquery As String
        'Dim docdate As String = Me.m_SBO_Form.Items.Item("20").Specific.Value

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
                    posrs.DoQuery(draftquery)
                    If posrs.RecordCount > 0 Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(posrs)
                        objMain.objApplication.StatusBar.SetText("Business partner already sent for approval", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    Else
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(posrs)
                        'Dim ofor As SBO_Form
                        'ofor = New DraftProcedureStages(Me.m_ParentAddon, Me, DocType, AppID, Me.SBO_Company.UserName, rsapp, Table, AppIDField, AppStatField, Document, docdate)
                        'ofor = Nothing
                        'objMain.objDraft.ItemEvent(FormUID, pVal, BubbleEvent)

                        ' objform.objDraft.CreateForm(DocType, AppID, objMain.objCompany.UserName, rsapp, Table, AppIDField, AppStatField, Document)
                        'objMain.objApplication.ActivateMenuItem("1304")
                        Dim objBaseForm As SAPbouiCOM.Form = objMain.objApplication.Forms.ActiveForm
                        objMain.ObjDraftProcedure.CreateForm(objBaseForm, DocType, AppID, objMain.objCompany.UserName, rsapp, Table, AppIDField, AppStatField, Document)
                        ''.CreateForm(objForm.UniqueID, objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value)
                        'objMain.objApplication.ActivateMenuItem("1304")
                    End If
                Else
                    objMain.objApplication.StatusBar.SetText("Login User Don't have Approvals Please Define Approvals and Add the Document", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        End If
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "1282" And pVal.BeforeAction = False Then

                Dim DocType As String = "BP"
                Dim rsapp As SAPbobsCOM.Recordset
                rsapp = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                Dim str As String = "SELECT T0.""Code"" As ""TemplateID"",T1.""U_Name"" As ""Originator"",S1.""Code"" As ""Stage"",S2.""U_UKey"" As ""Authorizers"" " &
            " ,S1.""U_NAP"" As ""NoOfAuth"",S1.""U_NRJ"" As ""NoOfRej"",T0.""CreateDate"",(Case When T3.""LineId""='1' Then 'S' Else 'O' End) As ""Status"" FROM ""@SBO_APPHDR"" T0 " &
            " INNER JOIN ""@SBO_APPREQ"" T1 ON T1.""Code""=T0.""Code"" " &
            " INNER JOIN ""@SBO_APPDOC"" T2 ON T2.""Code"" = T0.""Code"" " & " INNER JOIN ""@SBO_APPAUT"" T3 ON T3.""Code"" = T0.""Code"" " &
            " INNER JOIN ""@SBO_AST"" S1 ON T3.""U_M3_1""=S1.""Code"" " & " INNER JOIN ""@SBO_AST_C0"" S2 ON S1.""Code""=S2.""Code"" " &
            " WHERE ""U_Active"" = 'Y' " & " AND T1.""U_Name"" = '" & objMain.objCompany.UserName & "' " &
            " AND T2.""U_" & DocType & """ = 'Y' " & " AND IFNULL(S2.""U_UKey"",'')<>'' " & " ORDER BY T3.""LineId"""

                rsapp.DoQuery(str)

                If rsapp.RecordCount > 0 Then
                    Me.StatusChange(objform.UniqueID)
                End If

            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub StatusChange(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)
            objform.Items.Item("10002045").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Private Sub CheckSubmitButtonVisibility(ByVal FormUID As String)

        Try
            Dim objForm As SAPbouiCOM.Form =
        objMain.objApplication.Forms.Item(FormUID)

            Dim DocType As String = "BP"

            Dim rsapp As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim str As String = "SELECT IFNULL(T2.""U_" & DocType & """,'') AS ""DocFlag"" " & "FROM ""@SBO_APPHDR"" T0 " &
            "INNER JOIN ""@SBO_APPREQ"" T1 ON T1.""Code""=T0.""Code"" " & "INNER JOIN ""@SBO_APPDOC"" T2 ON T2.""Code"" = T0.""Code"" " &
            "WHERE T0.""U_Active"" = 'Y' " & "AND T1.""U_Name"" = '" & objMain.objCompany.UserName & "' " & "AND IFNULL(T2.""U_" & DocType & """,'') <> ''"


            rsapp.DoQuery(str)

            Dim hasUnchecked As Boolean = False

            If rsapp.RecordCount > 0 Then
                rsapp.MoveFirst()

                While Not rsapp.EoF
                    Dim flag As String = ""
                    Try
                        flag = rsapp.Fields.Item("DocFlag").Value.ToString().Trim().ToUpper()
                    Catch
                        flag = ""
                    End Try

                    If flag <> "Y" Then
                        hasUnchecked = True
                        Exit While
                    End If

                    rsapp.MoveNext()
                End While
            Else
                hasUnchecked = True
            End If


            If hasUnchecked Then
                objForm.Items.Item("Submit").Visible = False
                objForm.Items.Item("Submit").Enabled = False
            Else

                objForm.Items.Item("Submit").Visible = True
                objForm.Items.Item("Submit").Enabled = True
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

End Class
