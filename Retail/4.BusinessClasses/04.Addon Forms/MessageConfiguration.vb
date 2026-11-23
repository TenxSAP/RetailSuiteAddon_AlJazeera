Imports System.Net
Imports System.Net.Mail

Public Class MessageConfiguration
    Private oForm, objForm As SAPbouiCOM.Form
    Dim oMatrix0, oMatrix1, oMatrix2 As SAPbouiCOM.Matrix
    Dim oDBHeader, oDBDetail0, oDBDetail1, oDBDetail2, oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oItem, oItemUP, oItemUP1 As SAPbobsCOM.Items
    Dim Error_Code As Integer = 1
    Dim Error_Msg As String = ""

    Private Sub DataSourcesLoad(ByRef oForm As SAPbouiCOM.Form)
        Try
            oMatrix0 = oForm.Items.Item("Mtx").Specific
            oMatrix1 = oForm.Items.Item("MtxSMS").Specific
            oMatrix2 = oForm.Items.Item("MtxEML").Specific
            oDBHeader = oForm.DataSources.DBDataSources.Item("@MGRCONFIG")
            oDBDetail0 = oForm.DataSources.DBDataSources.Item("@MGRCONFIG_C0")
            oDBDetail1 = oForm.DataSources.DBDataSources.Item("@MGRCONFIG_C1")
            oDBDetail2 = oForm.DataSources.DBDataSources.Item("@MGRCONFIG_C2")

        Catch ex As Exception
            objMain.objUtilities.Write_Log(ex.Message)
        End Try
    End Sub
    Sub FormLoad(ByRef oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            Dim objRedioAuto As SAPbouiCOM.OptionBtn = oForm.Items.Item("Auto").Specific
            objRedioAuto.GroupWith("Manual")
            Dim objRedioAutoSMS As SAPbouiCOM.OptionBtn = oForm.Items.Item("AutoS").Specific
            objRedioAutoSMS.GroupWith("ManualS")
            Dim objRedioAutoEmail As SAPbouiCOM.OptionBtn = oForm.Items.Item("AutoE").Specific
            objRedioAutoEmail.GroupWith("ManualE")

            Dim objRedioAutoAtt As SAPbouiCOM.OptionBtn = oForm.Items.Item("Att").Specific
            objRedioAutoAtt.GroupWith("AttM")
            Dim objRedioAutoAttEML As SAPbouiCOM.OptionBtn = oForm.Items.Item("AttEML").Specific
            objRedioAutoAttEML.GroupWith("AttEMLM")

            SetDefault()

            'Form.Items.Item("SMS").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            'oForm.Items.Item("WhatsApp").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            'oForm.Items.Item("WMsg").Click(SAPbouiCOM.BoCellClickType.ct_Regular)o

            Try
                oForm.Items.Item("WhatsApp").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Catch ex As Exception

            End Try

            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case "MGRCONFIG"
                    If pVal.BeforeAction = True Then
                        objMain.objUtilities.LoadForm("MessageConfigurations.xml", "MGRCONFIG", ResourceType.Embeded)
                    Else
                        oForm = objMain.objApplication.Forms.ActiveForm
                        oForm.Items.Item("Doc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        Try
                            Dim objDocCombo As SAPbouiCOM.ComboBox = oForm.Items.Item("Doc").Specific
                            objDocCombo.ValidValues.Add("", "")
                            ''  objDocCombo.ValidValues.Add("SO1", "Sales Order 1")
                            ''objDocCombo.ValidValues.Add("SBA", "Sales Blanket Agreement")
                            objDocCombo.ValidValues.Add("SQ", "Sales Quotation")
                            objDocCombo.ValidValues.Add("SO", "Sales Order")
                            objDocCombo.ValidValues.Add("ARINV", "A/R Invoice")
                            objDocCombo.ValidValues.Add("ARRINV", "A/R Reserve Invoice")
                            objDocCombo.ValidValues.Add("DL", "Delivery")
                            objDocCombo.ValidValues.Add("ARC", "A/R Credit memo")
                            objDocCombo.ValidValues.Add("ARRET", "Return")

                            objDocCombo.ValidValues.Add("PQ", "Purchase Quotation")
                            objDocCombo.ValidValues.Add("PO", "Purchase Order")
                            objDocCombo.ValidValues.Add("GRPO", "Goods Receipt PO")
                            objDocCombo.ValidValues.Add("APINV", "A/P Invoice")
                            objDocCombo.ValidValues.Add("APC", "A/P Credit memo")
                            objDocCombo.ValidValues.Add("GR", "Goods Return")

                            objDocCombo.ValidValues.Add("GRP", "Goods Receipt")
                            objDocCombo.ValidValues.Add("GI", "Goods Issue")
                            objDocCombo.ValidValues.Add("ITR", "Inventory Transfer Request")
                            objDocCombo.ValidValues.Add("IT", "Inventory Transfer")

                            objDocCombo.ValidValues.Add("IP", "Incoming Payment")
                            objDocCombo.ValidValues.Add("OP", "Outgoing Payment")
                            objDocCombo.ValidValues.Add("JE", "Journal Entry")
                            objDocCombo.ValidValues.Add("PRD", "Production Order")
                            objDocCombo.ValidValues.Add("PR", "Payment Reminders")
                            objDocCombo.ValidValues.Add("PU", "Price Update")
                            objDocCombo.ValidValues.Add("PCR", "Price Change Request")

                        Catch ex As Exception

                        End Try
                        DataSourcesLoad(oForm)
                        FormLoad(oForm)
                    End If
                Case "1292" ''''''''''''''''''''''''''''''''''''''''''  Add Row
                    oForm = objMain.objApplication.Forms.ActiveForm
                    If pVal.BeforeAction = False Then
                        DataSourcesLoad(oForm)
                        oForm.Freeze(True)

                        oForm.Freeze(False)
                    End If
                Case "1293" ''''''''''''''''''''''''''''''''''''''''''  Delete Row
                    oForm = objMain.objApplication.Forms.ActiveForm
                    If pVal.BeforeAction = False Then
                        oForm.Freeze(True)
                        DataSourcesLoad(oForm)

                        oForm.Freeze(False)
                    End If
                Case "1281" '''''''''''''''''''''''''''''''''''''''''''  Find Mode
                    oForm = objMain.objApplication.Forms.ActiveForm
                    If pVal.BeforeAction = False Then
                    End If
                Case "1282" '''''''''''''''''''''''''''''''''''''''''''  Add Mode
                    oForm = objMain.objApplication.Forms.ActiveForm
                    If pVal.BeforeAction = False Then
                        DataSourcesLoad(oForm)
                        FormLoad(oForm)
                    End If
                Case "1290", "1288", "1289", "1291"  ''''''''''''''''''' First Next Prev Last Record
                    oForm = objMain.objApplication.Forms.ActiveForm
                    If pVal.BeforeAction = False Then
                        DataSourcesLoad(oForm)
                    End If

            End Select

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            oForm = objMain.objApplication.Forms.Item(pVal.FormUID)
            DataSourcesLoad(oForm)
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    Select Case pVal.ItemUID
                        Case "Doc"
                            Try
                                If pVal.BeforeAction = False And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    oForm.Freeze(True)
                                    Dim Doc As String = oDBHeader.GetValue("U_Doc", oDBHeader.Offset)

                                    Dim oRsExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim ExistQry As String = "Select * From ""@MGRCONFIG"" Where ""U_Doc""='" & Doc & "' and ""U_Status""='Y'"
                                    oRsExist.DoQuery(ExistQry)
                                    If oRsExist.RecordCount > 0 Then
                                        oDBHeader.SetValue("U_Doc", 0, "")
                                        objMain.objApplication.StatusBar.SetText("Active document already exist", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If

                                    DefaultMessage(Doc)
                                    oForm.Freeze(False)
                                End If
                            Catch ex As Exception
                                oForm.Freeze(False)
                            End Try
                    End Select
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    Select Case pVal.ItemUID
                        Case "Mtx"
                            Select Case pVal.ColUID
                                Case "EmpID", "UserID"
                                    Try
                                        If pVal.BeforeAction = False And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                            oForm.Freeze(True)
                                            Dim EmpID As String = oMatrix0.Columns.Item("EmpID").Cells.Item(pVal.Row).Specific.Value.ToString()
                                            Dim UserID As String = oMatrix0.Columns.Item("UserID").Cells.Item(pVal.Row).Specific.Value.ToString()
                                            If oMatrix0.VisualRowCount = pVal.Row Then
                                                If EmpID = "" AndAlso UserID = "" Then
                                                Else
                                                    objMain.objUtilities.AddRow(oMatrix0, oDBDetail0)
                                                End If
                                            End If
                                            oForm.Freeze(False)
                                        End If
                                    Catch ex As Exception
                                        oForm.Freeze(False)
                                    End Try
                            End Select
                        Case "MtxSMS"
                            Select Case pVal.ColUID
                                Case "EmpID", "UserID"
                                    Try
                                        If pVal.BeforeAction = False And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                            oForm.Freeze(True)
                                            Dim EmpID As String = oMatrix1.Columns.Item("EmpID").Cells.Item(pVal.Row).Specific.Value.ToString()
                                            Dim UserID As String = oMatrix1.Columns.Item("UserID").Cells.Item(pVal.Row).Specific.Value.ToString()
                                            If oMatrix1.VisualRowCount = pVal.Row Then
                                                If EmpID = "" AndAlso UserID = "" Then
                                                Else
                                                    objMain.objUtilities.AddRow(oMatrix1, oDBDetail1)
                                                End If
                                            End If
                                            oForm.Freeze(False)
                                        End If
                                    Catch ex As Exception
                                        oForm.Freeze(False)
                                    End Try
                            End Select
                        Case "MtxEML"
                            Select Case pVal.ColUID
                                Case "EmpID", "UserID"
                                    Try
                                        If pVal.BeforeAction = False And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                            oForm.Freeze(True)
                                            Dim EmpID As String = oMatrix2.Columns.Item("EmpID").Cells.Item(pVal.Row).Specific.Value.ToString()
                                            Dim UserID As String = oMatrix2.Columns.Item("UserID").Cells.Item(pVal.Row).Specific.Value.ToString()
                                            If oMatrix2.VisualRowCount = pVal.Row Then
                                                If EmpID = "" AndAlso UserID = "" Then
                                                Else
                                                    objMain.objUtilities.AddRow(oMatrix2, oDBDetail2)
                                                End If
                                            End If
                                            oForm.Freeze(False)
                                        End If
                                    Catch ex As Exception
                                        oForm.Freeze(False)
                                    End Try
                            End Select
                    End Select
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = oForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If oCFL.UniqueID = "CFL_USR" Then
                            oMatrix0.Columns.Item("UserID").Cells.Item(pVal.Row).Specific.Value = oDT.GetValue("USERID", 0).ToString()
                        End If
                        If oCFL.UniqueID = "CFL_EMP" Then
                            oMatrix0.Columns.Item("EmpID").Cells.Item(pVal.Row).Specific.Value = oDT.GetValue("empID", 0).ToString()
                        End If

                        If oCFL.UniqueID = "CFL_USRS" Then
                            oMatrix1.Columns.Item("UserID").Cells.Item(pVal.Row).Specific.Value = oDT.GetValue("USERID", 0).ToString()
                        End If
                        If oCFL.UniqueID = "CFL_EMPS" Then
                            oMatrix1.Columns.Item("EmpID").Cells.Item(pVal.Row).Specific.Value = oDT.GetValue("empID", 0).ToString()
                        End If

                        If oCFL.UniqueID = "CFL_USRE" Then
                            oMatrix2.Columns.Item("UserID").Cells.Item(pVal.Row).Specific.Value = oDT.GetValue("USERID", 0).ToString()
                        End If
                        If oCFL.UniqueID = "CFL_EMPE" Then
                            oMatrix2.Columns.Item("EmpID").Cells.Item(pVal.Row).Specific.Value = oDT.GetValue("empID", 0).ToString()
                        End If
                    End If
            End Select
        Catch ex As Exception
            oForm.Freeze(False)
            If (ex.Message = "Form - Invalid Form") Then
            Else
                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        End Try
    End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = objMain.objApplication.Forms.ActiveForm
            DataSourcesLoad(oForm)
            oForm.Freeze(True)

            '
            objForm = objMain.objApplication.Forms.GetForm("MGRCONFIG", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Dim DocNum As String = oDBs_Head.GetValue("U_Doc", 0)
                        Dim oUtil As New Utilities()
                        oUtil.MarketingEmail(DocNum)


                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                    End If
            End Select
            '
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    Private Function Validation() As Boolean
        Try


        Catch ex As Exception
        End Try

        Return True
    End Function
    Sub SetDefault()
        Try
            Dim Query As String
            If objMain.objCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                Query = "SELECT " + IsNull + "(MAX(TO_INTEGER (""DocNum"")),0)+1 FROM ""@MGRCONFIG"""
            Else
                Query = "SELECT ISNULL(MAX(CAST(""DocNum"" AS int)),0)+1 FROM ""@MGRCONFIG"""
            End If

            Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rs.DoQuery(Query)
            oDBHeader.SetValue("DocNum", 0, rs.Fields.Item(0).Value)

            objMain.objUtilities.AddRow(oMatrix0, oDBDetail0)
            objMain.objUtilities.AddRow(oMatrix1, oDBDetail1)
            objMain.objUtilities.AddRow(oMatrix2, oDBDetail2)

        Catch ex As Exception
        End Try
    End Sub
    Sub DefaultMessage(ByVal Doc As String)
        Try

            If Doc = "SO" Then
                Dim Msg As String
                Dim Qstr As String = "CALL ""TENX_MESSENGER_MSG_CONF"" ('DefaultMessage','" + Doc + "')"
                Dim Rset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset.DoQuery(Qstr)
                If Rset.RecordCount > 0 Then
                    Msg = Rset.Fields.Item(0).Value.ToString()
                End If

                oDBHeader.SetValue("U_Msg", 0, Msg)

                Msg = ""
                Msg = "ORDER CONFIRMATION" & vbCrLf
                Msg = Msg + "{CardName}" & vbCrLf
                Msg = Msg + "Order Summary " & vbCrLf
                Msg = Msg + "Order Number : " + "{DocNum}" & vbCrLf
                Msg = Msg + "Order Date : " + "{TaxDate}" & vbCrLf
                Msg = Msg + "Product/SKU's Count : " + "{ItemCount}" & vbCrLf
                Msg = Msg + "Total before Discount : " + "{Total}" & vbCrLf
                Msg = Msg + "Vat Amount : " + "{VatSum}" & vbCrLf
                Msg = Msg + "Total Amount : " + "{DocTotal}" & vbCrLf
                Msg = Msg + "Thank You... Your Sales order has been placed."
                oDBHeader.SetValue("U_MsgSMS", 0, Msg)
                oDBHeader.SetValue("U_MsgEML", 0, Msg)
            ElseIf Doc = "SQ" Then
                Dim Msg As String
                Dim Qstr As String = "CALL ""TENX_MESSENGER_MSG_CONF"" ('DefaultMessage','" + Doc + "')"
                Dim Rset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset.DoQuery(Qstr)
                If Rset.RecordCount > 0 Then
                    Msg = Rset.Fields.Item(0).Value.ToString()
                End If

                oDBHeader.SetValue("U_Msg", 0, Msg)

                Msg = ""
                Msg = "QUOTATION CONFIRMATION" & vbCrLf
                Msg = Msg + "{CardName}" & vbCrLf
                Msg = Msg + "Order Summary " & vbCrLf
                Msg = Msg + "Order Number : " + "{DocNum}" & vbCrLf
                Msg = Msg + "Order Date : " + "{TaxDate}" & vbCrLf
                Msg = Msg + "Product/SKU's Count : " + "{ItemCount}" & vbCrLf
                Msg = Msg + "Total before Discount : " + "{Total}" & vbCrLf
                Msg = Msg + "Vat Amount : " + "{VatSum}" & vbCrLf
                Msg = Msg + "Total Amount : " + "{DocTotal}" & vbCrLf
                Msg = Msg + "Thank You... Your Sales Quotation has been placed."

                oDBHeader.SetValue("U_MsgSMS", 0, Msg)
                oDBHeader.SetValue("U_MsgEML", 0, Msg)

            ElseIf Doc = "ARINV" Then
                Dim Msg As String
                Dim Qstr As String = "CALL ""TENX_MESSENGER_MSG_CONF"" ('DefaultMessage','" + Doc + "')"
                Dim Rset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset.DoQuery(Qstr)
                If Rset.RecordCount > 0 Then
                    Msg = Rset.Fields.Item(0).Value.ToString()
                End If

                oDBHeader.SetValue("U_Msg", 0, Msg)

                Msg = ""
                Msg = "INVOICE CONFIRMATION " & vbCrLf
                Msg = Msg + "{CardName}" & vbCrLf
                Msg = Msg + "Order Summary " & vbCrLf
                Msg = Msg + "Order Number : " + "{DocNum}" & vbCrLf
                Msg = Msg + "Order Date : " + "{TaxDate}" & vbCrLf
                Msg = Msg + "Product/SKU's Count : " + "{ItemCount}" & vbCrLf
                Msg = Msg + "Total before Discount : " + "{Total}" & vbCrLf
                Msg = Msg + "Vat Amount : " + "{VatSum}" & vbCrLf
                Msg = Msg + "Total Amount : " + "{DocTotal}" & vbCrLf
                Msg = Msg + "Thank you. "

                oDBHeader.SetValue("U_MsgSMS", 0, Msg)
                oDBHeader.SetValue("U_MsgEML", 0, Msg)

            ElseIf Doc = "DL" Then
                Dim Msg As String
                Dim Qstr As String = "CALL ""TENX_MESSENGER_MSG_CONF"" ('DefaultMessage','" + Doc + "')"
                Dim Rset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset.DoQuery(Qstr)
                If Rset.RecordCount > 0 Then
                    Msg = Rset.Fields.Item(0).Value.ToString()
                End If

                oDBHeader.SetValue("U_Msg", 0, Msg)

                Msg = ""
                Msg = "DELIVERY CONFIRMATION " & vbCrLf
                Msg = Msg + "{CardName}" & vbCrLf
                Msg = Msg + "Order Summary " & vbCrLf
                Msg = Msg + "Order Number : " + "{DocNum}" & vbCrLf
                Msg = Msg + "Order Date : " + "{TaxDate}" & vbCrLf
                Msg = Msg + "Product/SKU's Count : " + "{ItemCount}" & vbCrLf
                Msg = Msg + "Total before Discount : " + "{Total}" & vbCrLf
                Msg = Msg + "Vat Amount : " + "{VatSum}" & vbCrLf
                Msg = Msg + "Total Amount : " + "{DocTotal}" & vbCrLf
                Msg = Msg + "Thank You... Your Delivery Note has been placed."

                oDBHeader.SetValue("U_MsgSMS", 0, Msg)
                oDBHeader.SetValue("U_MsgEML", 0, Msg)

            ElseIf Doc = "ARC" Then
                Dim Msg As String
                Dim Qstr As String = "CALL ""TENX_MESSENGER_MSG_CONF"" ('DefaultMessage','" + Doc + "')"
                Dim Rset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset.DoQuery(Qstr)
                If Rset.RecordCount > 0 Then
                    Msg = Rset.Fields.Item(0).Value.ToString()
                End If

                oDBHeader.SetValue("U_Msg", 0, Msg)

                Msg = ""
                Msg = "CREDIT MEMO CONFIRMATION " & vbCrLf
                Msg = Msg + "{CardName}" & vbCrLf
                Msg = Msg + "Order Summary " & vbCrLf
                Msg = Msg + "Order Number : " + "{DocNum}" & vbCrLf
                Msg = Msg + "Order Date : " + "{TaxDate}" & vbCrLf
                Msg = Msg + "Product/SKU's Count : " + "{ItemCount}" & vbCrLf
                Msg = Msg + "Total before Discount : " + "{Total}" & vbCrLf
                Msg = Msg + "Vat Amount : " + "{VatSum}" & vbCrLf
                Msg = Msg + "Total Amount : " + "{DocTotal}" & vbCrLf
                Msg = Msg + "Thank You... Your A/R Credit Note has been placed."

                oDBHeader.SetValue("U_MsgSMS", 0, Msg)
                oDBHeader.SetValue("U_MsgEML", 0, Msg)

            ElseIf Doc = "ARRET" Then
                Dim Msg As String
                Dim Qstr As String = "CALL ""TENX_MESSENGER_MSG_CONF"" ('DefaultMessage','" + Doc + "')"
                Dim Rset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset.DoQuery(Qstr)
                If Rset.RecordCount > 0 Then
                    Msg = Rset.Fields.Item(0).Value.ToString()
                End If

                oDBHeader.SetValue("U_Msg", 0, Msg)

                Msg = ""
                Msg = "RETURN CONFIRMATION " & vbCrLf
                Msg = Msg + "{CardName}" & vbCrLf
                Msg = Msg + "Order Summary " & vbCrLf
                Msg = Msg + "Order Number : " + "{DocNum}" & vbCrLf
                Msg = Msg + "Order Date : " + "{TaxDate}" & vbCrLf
                Msg = Msg + "Product/SKU's Count : " + "{ItemCount}" & vbCrLf
                Msg = Msg + "Total before Discount : " + "{Total}" & vbCrLf
                Msg = Msg + "Vat Amount : " + "{VatSum}" & vbCrLf
                Msg = Msg + "Total Amount : " + "{DocTotal}" & vbCrLf
                Msg = Msg + "Thank You... Your Return Document has been placed."

                oDBHeader.SetValue("U_MsgSMS", 0, Msg)
                oDBHeader.SetValue("U_MsgEML", 0, Msg)

            ElseIf Doc = "PO" Then
                Dim Msg As String
                Dim Qstr As String = "CALL ""TENX_MESSENGER_MSG_CONF"" ('DefaultMessage','" + Doc + "')"
                Dim Rset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset.DoQuery(Qstr)
                If Rset.RecordCount > 0 Then
                    Msg = Rset.Fields.Item(0).Value.ToString()
                End If

                oDBHeader.SetValue("U_Msg", 0, Msg)

                Msg = ""
                Msg = "PO CONFIRMATION " & vbCrLf
                Msg = Msg + "{CardName}" & vbCrLf
                Msg = Msg + "Order Summary " & vbCrLf
                Msg = Msg + "Order Number : " + "{DocNum}" & vbCrLf
                Msg = Msg + "Order Date : " + "{TaxDate}" & vbCrLf
                Msg = Msg + "Product/SKU's Count : " + "{ItemCount}" & vbCrLf
                Msg = Msg + "Total before Discount : " + "{Total}" & vbCrLf
                Msg = Msg + "Vat Amount : " + "{VatSum}" & vbCrLf
                Msg = Msg + "Total Amount : " + "{DocTotal}" & vbCrLf
                Msg = Msg + "Thank You... Your Purchase order has been placed."

                oDBHeader.SetValue("U_MsgSMS", 0, Msg)
                oDBHeader.SetValue("U_MsgEML", 0, Msg)

            ElseIf Doc = "PQ" Then
                Dim Msg As String
                Dim Qstr As String = "CALL ""TENX_MESSENGER_MSG_CONF"" ('DefaultMessage','" + Doc + "')"
                Dim Rset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset.DoQuery(Qstr)
                If Rset.RecordCount > 0 Then
                    Msg = Rset.Fields.Item(0).Value.ToString()
                End If

                oDBHeader.SetValue("U_Msg", 0, Msg)

                Msg = ""
                Msg = "PQ CONFIRMATION " & vbCrLf
                Msg = Msg + "{CardName}" & vbCrLf
                Msg = Msg + "Order Summary " & vbCrLf
                Msg = Msg + "Order Number : " + "{DocNum}" & vbCrLf
                Msg = Msg + "Order Date : " + "{TaxDate}" & vbCrLf
                Msg = Msg + "Product/SKU's Count : " + "{ItemCount}" & vbCrLf
                Msg = Msg + "Total before Discount : " + "{Total}" & vbCrLf
                Msg = Msg + "Vat Amount : " + "{VatSum}" & vbCrLf
                Msg = Msg + "Total Amount : " + "{DocTotal}" & vbCrLf
                Msg = Msg + "Thank You... Your Purchase Quotaion has been placed."

                oDBHeader.SetValue("U_MsgSMS", 0, Msg)
                oDBHeader.SetValue("U_MsgEML", 0, Msg)

            ElseIf Doc = "GRPO" Then
                Dim Msg As String
                Dim Qstr As String = "CALL ""TENX_MESSENGER_MSG_CONF"" ('DefaultMessage','" + Doc + "')"
                Dim Rset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset.DoQuery(Qstr)
                If Rset.RecordCount > 0 Then
                    Msg = Rset.Fields.Item(0).Value.ToString()
                End If

                oDBHeader.SetValue("U_Msg", 0, Msg)

                Msg = ""
                Msg = "GRPO CONFIRMATION " & vbCrLf
                Msg = Msg + "{CardName}" & vbCrLf
                Msg = Msg + "Order Summary " & vbCrLf
                Msg = Msg + "Order Number : " + "{DocNum}" & vbCrLf
                Msg = Msg + "Order Date : " + "{TaxDate}" & vbCrLf
                Msg = Msg + "Product/SKU's Count : " + "{ItemCount}" & vbCrLf
                Msg = Msg + "Total before Discount : " + "{Total}" & vbCrLf
                Msg = Msg + "Vat Amount : " + "{VatSum}" & vbCrLf
                Msg = Msg + "Total Amount : " + "{DocTotal}" & vbCrLf
                Msg = Msg + "Thank You... Your Goods Return Against PO has been placed."

                oDBHeader.SetValue("U_MsgSMS", 0, Msg)
                oDBHeader.SetValue("U_MsgEML", 0, Msg)

            ElseIf Doc = "GR" Then
                Dim Msg As String
                Dim Qstr As String = "CALL ""TENX_MESSENGER_MSG_CONF"" ('DefaultMessage','" + Doc + "')"
                Dim Rset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset.DoQuery(Qstr)
                If Rset.RecordCount > 0 Then
                    Msg = Rset.Fields.Item(0).Value.ToString()
                End If

                oDBHeader.SetValue("U_Msg", 0, Msg)

                Msg = ""
                Msg = "GOODS RETURN CONFIRMATION " & vbCrLf
                Msg = Msg + "{CardName}" & vbCrLf
                Msg = Msg + "Order Summary " & vbCrLf
                Msg = Msg + "Order Number : " + "{DocNum}" & vbCrLf
                Msg = Msg + "Order Date : " + "{TaxDate}" & vbCrLf
                Msg = Msg + "Product/SKU's Count : " + "{ItemCount}" & vbCrLf
                Msg = Msg + "Total before Discount : " + "{Total}" & vbCrLf
                Msg = Msg + "Vat Amount : " + "{VatSum}" & vbCrLf
                Msg = Msg + "Total Amount : " + "{DocTotal}" & vbCrLf
                Msg = Msg + "Thank You... Your Goods Return has been placed."

                oDBHeader.SetValue("U_MsgSMS", 0, Msg)
                oDBHeader.SetValue("U_MsgEML", 0, Msg)


            ElseIf Doc = "ARRINV" Then
                Dim Msg As String
                Dim Qstr As String = "CALL ""TENX_MESSENGER_MSG_CONF"" ('DefaultMessage','" + Doc + "')"
                Dim Rset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset.DoQuery(Qstr)
                If Rset.RecordCount > 0 Then
                    Msg = Rset.Fields.Item(0).Value.ToString()
                End If

                oDBHeader.SetValue("U_Msg", 0, Msg)

                Msg = ""
                Msg = "RESERVE INVOICE CONFIRMATION " & vbCrLf
                Msg = Msg + "{CardName}" & vbCrLf
                Msg = Msg + "Order Summary " & vbCrLf
                Msg = Msg + "Order Number : " + "{DocNum}" & vbCrLf
                Msg = Msg + "Order Date : " + "{TaxDate}" & vbCrLf
                Msg = Msg + "Product/SKU's Count : " + "{ItemCount}" & vbCrLf
                Msg = Msg + "Total before Discount : " + "{Total}" & vbCrLf
                Msg = Msg + "Vat Amount : " + "{VatSum}" & vbCrLf
                Msg = Msg + "Total Amount : " + "{DocTotal}" & vbCrLf
                Msg = Msg + "Thank You... Your Reserve Sales Invoice has been placed"

                oDBHeader.SetValue("U_MsgSMS", 0, Msg)
                oDBHeader.SetValue("U_MsgEML", 0, Msg)

            ElseIf Doc = "APINV" Then
                Dim Msg As String
                Dim Qstr As String = "CALL ""TENX_MESSENGER_MSG_CONF"" ('DefaultMessage','" + Doc + "')"
                Dim Rset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset.DoQuery(Qstr)
                If Rset.RecordCount > 0 Then
                    Msg = Rset.Fields.Item(0).Value.ToString()
                End If

                oDBHeader.SetValue("U_Msg", 0, Msg)

                Msg = ""
                Msg = "A/P INVOICE CONFIRMATION " & vbCrLf
                Msg = Msg + "{CardName}" & vbCrLf
                Msg = Msg + "Order Summary " & vbCrLf
                Msg = Msg + "Order Number : " + "{DocNum}" & vbCrLf
                Msg = Msg + "Order Date : " + "{TaxDate}" & vbCrLf
                Msg = Msg + "Product/SKU's Count : " + "{ItemCount}" & vbCrLf
                Msg = Msg + "Total before Discount : " + "{Total}" & vbCrLf
                Msg = Msg + "Vat Amount : " + "{VatSum}" & vbCrLf
                Msg = Msg + "Total Amount : " + "{DocTotal}" & vbCrLf
                Msg = Msg + "Thank You... Your Purchase Invoice has been placed."

                oDBHeader.SetValue("U_MsgSMS", 0, Msg)
                oDBHeader.SetValue("U_MsgEML", 0, Msg)

            ElseIf Doc = "APC" Then
                Dim Msg As String
                Dim Qstr As String = "CALL ""TENX_MESSENGER_MSG_CONF"" ('DefaultMessage','" + Doc + "')"
                Dim Rset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset.DoQuery(Qstr)
                If Rset.RecordCount > 0 Then
                    Msg = Rset.Fields.Item(0).Value.ToString()
                End If

                oDBHeader.SetValue("U_Msg", 0, Msg)

                Msg = ""
                Msg = "A/P CREDIT MEMO CONFIRMATION " & vbCrLf
                Msg = Msg + "{CardName}" & vbCrLf
                Msg = Msg + "Order Summary " & vbCrLf
                Msg = Msg + "Order Number : " + "{DocNum}" & vbCrLf
                Msg = Msg + "Order Date : " + "{TaxDate}" & vbCrLf
                Msg = Msg + "Product/SKU's Count : " + "{ItemCount}" & vbCrLf
                Msg = Msg + "Total before Discount : " + "{Total}" & vbCrLf
                Msg = Msg + "Vat Amount : " + "{VatSum}" & vbCrLf
                Msg = Msg + "Total Amount : " + "{DocTotal}" & vbCrLf
                Msg = Msg + "Thank You... Your A/P Credit Note has been placed."

                oDBHeader.SetValue("U_MsgSMS", 0, Msg)
                oDBHeader.SetValue("U_MsgEML", 0, Msg)

            ElseIf Doc = "GRP" Then
                Dim Msg As String
                Dim Qstr As String = "CALL ""TENX_MESSENGER_MSG_CONF"" ('DefaultMessage','" + Doc + "')"
                Dim Rset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset.DoQuery(Qstr)
                If Rset.RecordCount > 0 Then
                    Msg = Rset.Fields.Item(0).Value.ToString()
                End If

                oDBHeader.SetValue("U_Msg", 0, Msg)

                Msg = ""
                Msg = "GOODS RECEIPT CONFIRMATION " & vbCrLf
                Msg = Msg + "{CardName}" & vbCrLf
                Msg = Msg + "Order Summary " & vbCrLf
                Msg = Msg + "Order Number : " + "{DocNum}" & vbCrLf
                Msg = Msg + "Order Date : " + "{TaxDate}" & vbCrLf
                Msg = Msg + "Product/SKU's Count : " + "{ItemCount}" & vbCrLf
                Msg = Msg + "Total before Discount : " + "{Total}" & vbCrLf
                Msg = Msg + "Vat Amount : " + "{VatSum}" & vbCrLf
                Msg = Msg + "Total Amount : " + "{DocTotal}" & vbCrLf
                Msg = Msg + "Thank you."

                oDBHeader.SetValue("U_MsgSMS", 0, Msg)
                oDBHeader.SetValue("U_MsgEML", 0, Msg)

            ElseIf Doc = "GI" Then
                Dim Msg As String
                Dim Qstr As String = "CALL ""TENX_MESSENGER_MSG_CONF"" ('DefaultMessage','" + Doc + "')"
                Dim Rset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset.DoQuery(Qstr)
                If Rset.RecordCount > 0 Then
                    Msg = Rset.Fields.Item(0).Value.ToString()
                End If

                oDBHeader.SetValue("U_Msg", 0, Msg)

                Msg = ""
                Msg = "GOODS ISSUE CONFIRMATION " & vbCrLf
                Msg = Msg + "{CardName}" & vbCrLf
                Msg = Msg + "Order Summary " & vbCrLf
                Msg = Msg + "Order Number : " + "{DocNum}" & vbCrLf
                Msg = Msg + "Order Date : " + "{TaxDate}" & vbCrLf
                Msg = Msg + "Product/SKU's Count : " + "{ItemCount}" & vbCrLf
                Msg = Msg + "Total before Discount : " + "{Total}" & vbCrLf
                Msg = Msg + "Vat Amount : " + "{VatSum}" & vbCrLf
                Msg = Msg + "Total Amount : " + "{DocTotal}" & vbCrLf
                Msg = Msg + "Thank you."

                oDBHeader.SetValue("U_MsgSMS", 0, Msg)
                oDBHeader.SetValue("U_MsgEML", 0, Msg)

            ElseIf Doc = "ITR" Then
                Dim Msg As String
                Dim Qstr As String = "CALL ""TENX_MESSENGER_MSG_CONF"" ('DefaultMessage','" + Doc + "')"
                Dim Rset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset.DoQuery(Qstr)
                If Rset.RecordCount > 0 Then
                    Msg = Rset.Fields.Item(0).Value.ToString()
                End If

                oDBHeader.SetValue("U_Msg", 0, Msg)

                Msg = ""
                Msg = "INVENTORY TRANSFER REQUEST CONFIRMATION " & vbCrLf
                Msg = Msg + "{CardName}" & vbCrLf
                Msg = Msg + "Order Summary " & vbCrLf
                Msg = Msg + "Order Number : " + "{DocNum}" & vbCrLf
                Msg = Msg + "Order Date : " + "{TaxDate}" & vbCrLf
                Msg = Msg + "Product/SKU's Count : " + "{ItemCount}" & vbCrLf
                Msg = Msg + "Total before Discount : " + "{Total}" & vbCrLf
                Msg = Msg + "Vat Amount : " + "{VatSum}" & vbCrLf
                Msg = Msg + "Total Amount : " + "{DocTotal}" & vbCrLf
                Msg = Msg + "Thank you."

                oDBHeader.SetValue("U_MsgSMS", 0, Msg)
                oDBHeader.SetValue("U_MsgEML", 0, Msg)

            ElseIf Doc = "IT" Then
                Dim Msg As String
                Dim Qstr As String = "CALL ""TENX_MESSENGER_MSG_CONF"" ('DefaultMessage','" + Doc + "')"
                Dim Rset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset.DoQuery(Qstr)
                If Rset.RecordCount > 0 Then
                    Msg = Rset.Fields.Item(0).Value.ToString()
                End If

                oDBHeader.SetValue("U_Msg", 0, Msg)

                Msg = ""
                Msg = "INVENTORY TRANSFER CONFIRMATION " & vbCrLf
                Msg = Msg + "{CardName}" & vbCrLf
                Msg = Msg + "Order Summary " & vbCrLf
                Msg = Msg + "Order Number : " + "{DocNum}" & vbCrLf
                Msg = Msg + "Order Date : " + "{TaxDate}" & vbCrLf
                Msg = Msg + "Product/SKU's Count : " + "{ItemCount}" & vbCrLf
                Msg = Msg + "Total before Discount : " + "{Total}" & vbCrLf
                Msg = Msg + "Vat Amount : " + "{VatSum}" & vbCrLf
                Msg = Msg + "Total Amount : " + "{DocTotal}" & vbCrLf
                Msg = Msg + "Thank you."

                oDBHeader.SetValue("U_MsgSMS", 0, Msg)
                oDBHeader.SetValue("U_MsgEML", 0, Msg)

            ElseIf Doc = "IP" Then
                Dim Msg As String
                Dim Qstr As String = "CALL ""TENX_MESSENGER_MSG_CONF"" ('DefaultMessage','" + Doc + "')"
                Dim Rset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset.DoQuery(Qstr)
                If Rset.RecordCount > 0 Then
                    Msg = Rset.Fields.Item(0).Value.ToString()
                End If

                oDBHeader.SetValue("U_Msg", 0, Msg)

                Msg = ""
                Msg = "INCOMING PAYMENT CONFIRMATION " & vbCrLf
                Msg = Msg + "{CardName}" & vbCrLf
                Msg = Msg + "Order Summary " & vbCrLf
                Msg = Msg + "Order Number : " + "{DocNum}" & vbCrLf
                Msg = Msg + "Order Date : " + "{TaxDate}" & vbCrLf
                Msg = Msg + "Product/SKU's Count : " + "{ItemCount}" & vbCrLf
                Msg = Msg + "Total before Discount : " + "{Total}" & vbCrLf
                Msg = Msg + "Vat Amount : " + "{VatSum}" & vbCrLf
                Msg = Msg + "Total Amount : " + "{DocTotal}" & vbCrLf
                Msg = Msg + "Incoming payment is created."

                oDBHeader.SetValue("U_MsgSMS", 0, Msg)
                oDBHeader.SetValue("U_MsgEML", 0, Msg)

            ElseIf Doc = "OP" Then
                Dim Msg As String
                Dim Qstr As String = "CALL ""TENX_MESSENGER_MSG_CONF"" ('DefaultMessage','" + Doc + "')"
                Dim Rset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset.DoQuery(Qstr)
                If Rset.RecordCount > 0 Then
                    Msg = Rset.Fields.Item(0).Value.ToString()
                End If

                oDBHeader.SetValue("U_Msg", 0, Msg)

                Msg = "OUTGOING PAYMENT CONFIRMATION " & vbCrLf
                Msg = Msg + "{CardName}" & vbCrLf
                Msg = Msg + "Order Summary " & vbCrLf
                Msg = Msg + "Order Number : " + "{DocNum}" & vbCrLf
                Msg = Msg + "Order Date : " + "{TaxDate}" & vbCrLf
                Msg = Msg + "Product/SKU's Count : " + "{ItemCount}" & vbCrLf
                Msg = Msg + "Total before Discount : " + "{Total}" & vbCrLf
                Msg = Msg + "Vat Amount : " + "{VatSum}" & vbCrLf
                Msg = Msg + "Total Amount : " + "{DocTotal}" & vbCrLf
                Msg = Msg + "Thank you."

                oDBHeader.SetValue("U_MsgSMS", 0, Msg)
                oDBHeader.SetValue("U_MsgEML", 0, Msg)

            ElseIf Doc = "JE" Then
                Dim Msg As String
                Dim Qstr As String = "CALL ""TENX_MESSENGER_MSG_CONF"" ('DefaultMessage','" + Doc + "')"
                Dim Rset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset.DoQuery(Qstr)
                If Rset.RecordCount > 0 Then
                    Msg = Rset.Fields.Item(0).Value.ToString()
                End If

                oDBHeader.SetValue("U_Msg", 0, Msg)

                Msg = "JOURNAL ENTRY CONFIRMATION " & vbCrLf
                Msg = Msg + "{CardName}" & vbCrLf
                Msg = Msg + "Order Summary " & vbCrLf
                Msg = Msg + "Order Number : " + "{DocNum}" & vbCrLf
                Msg = Msg + "Order Date : " + "{TaxDate}" & vbCrLf
                Msg = Msg + "Product/SKU's Count : " + "{ItemCount}" & vbCrLf
                Msg = Msg + "Total before Discount : " + "{Total}" & vbCrLf
                Msg = Msg + "Vat Amount : " + "{VatSum}" & vbCrLf
                Msg = Msg + "Total Amount : " + "{DocTotal}" & vbCrLf
                Msg = Msg + "Thank you."

                oDBHeader.SetValue("U_MsgSMS", 0, Msg)
                oDBHeader.SetValue("U_MsgEML", 0, Msg)
            ElseIf Doc = "PRD" Then
                Dim Msg As String
                Dim Qstr As String = "CALL ""TENX_MESSENGER_MSG_CONF"" ('DefaultMessage','" + Doc + "')"
                Dim Rset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset.DoQuery(Qstr)
                If Rset.RecordCount > 0 Then
                    Msg = Rset.Fields.Item(0).Value.ToString()
                End If

                oDBHeader.SetValue("U_Msg", 0, Msg)

                Msg = "PRODUCTION ORDER CONFIRMATION " & vbCrLf
                Msg = Msg + "{CardName}" & vbCrLf
                Msg = Msg + "Order Summary " & vbCrLf
                Msg = Msg + "Order Number : " + "{DocNum}" & vbCrLf
                Msg = Msg + "Order Date : " + "{TaxDate}" & vbCrLf
                Msg = Msg + "Product/SKU's Count : " + "{ItemCount}" & vbCrLf
                Msg = Msg + "Total before Discount : " + "{Total}" & vbCrLf
                Msg = Msg + "Vat Amount : " + "{VatSum}" & vbCrLf
                Msg = Msg + "Total Amount : " + "{DocTotal}" & vbCrLf
                Msg = Msg + "Thank you."

                oDBHeader.SetValue("U_MsgSMS", 0, Msg)
                oDBHeader.SetValue("U_MsgEML", 0, Msg)

                'Madhu PCR email
            ElseIf Doc = "PCR" Then
                Dim Msg As String
                Dim Qstr As String = "CALL ""TENX_MESSENGER_MSG_CONF"" ('DefaultMessage','" + Doc + "')"
                Dim Rset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Rset.DoQuery(Qstr)
                If Rset.RecordCount > 0 Then
                    Msg = Rset.Fields.Item(0).Value.ToString()
                End If

                oDBHeader.SetValue("U_Msg", 0, Msg)

                Msg = "PRICE CHANGE REQUEST " & vbCrLf
                Msg = Msg + "{CardName}" & vbCrLf
                Msg = Msg + "Order Summary " & vbCrLf
                Msg = Msg + "Order Number : " + "{DocNum}" & vbCrLf
                Msg = Msg + "Order Date : " + "{TaxDate}" & vbCrLf
                Msg = Msg + "Product/SKU's Count : " + "{ItemCount}" & vbCrLf
                Msg = Msg + "Total before Discount : " + "{Total}" & vbCrLf
                Msg = Msg + "Vat Amount : " + "{VatSum}" & vbCrLf
                Msg = Msg + "Total Amount : " + "{DocTotal}" & vbCrLf
                Msg = Msg + "Thank you."

                oDBHeader.SetValue("U_MsgSMS", 0, Msg)
                oDBHeader.SetValue("U_MsgEML", 0, Msg)


            ElseIf Doc = "" Then
                Dim Msg As String = ""
                oDBHeader.SetValue("U_Msg", 0, Msg)
                oDBHeader.SetValue("U_MsgSMS", 0, Msg)
                oDBHeader.SetValue("U_MsgEML", 0, Msg)
            End If
        Catch ex As Exception

        End Try
    End Sub
End Class
