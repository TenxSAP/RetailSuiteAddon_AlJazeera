Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Mail
Imports System.Text
Public Class EmailHistory
    Private oForm As SAPbouiCOM.Form
    Dim oDBHeader, oDBDetail1 As SAPbouiCOM.DataTable
    Dim oDBH, oDBDS1 As SAPbouiCOM.DBDataSource
    Dim oItem, oItemUP, oItemUP1 As SAPbobsCOM.Items
    Dim Error_Code As Integer = 1
    Dim Error_Msg As String = ""
    Dim objGrid As SAPbouiCOM.Grid
    Dim Responce1 As String = ""
    Private sISNULL As String = "IFNULL"
    Sub FormLoad(ByRef oForm As SAPbouiCOM.Form)

        Try
            oForm.Freeze(True)
            Dim rnd As New Random()
            Dim randomNumber As Integer = rnd.Next(10, 100)
            oDBDetail1 = oForm.DataSources.DataTables.Add("" + randomNumber.ToString() + "")
            oDBDetail1 = oForm.DataSources.DataTables.Item("" + randomNumber.ToString() + "")
            objGrid = oForm.Items.Item("grid").Specific
            Try
                If objGrid.Columns.Count > 0 Then
                    objGrid.DataTable.Clear()
                End If
            Catch ex As Exception
                objMain.objUtilities.Write_Log(ex.Message)
            End Try
            oForm.Items.Item("Item_5").Specific.value = objMain.objCompany.GetDBServerDate().ToString("yyyyMMdd")
            oForm.Items.Item("Item_6").Specific.value = objMain.objCompany.GetDBServerDate().ToString("yyyyMMdd")
            Dim DetailsQry As String = "CALL ""TENX_MESSENGER_EMAIL_HISTORY"" ('FormLoad','','','" + objMain.objCompany.GetDBServerDate().ToString("yyyyMMdd") + "','" + objMain.objCompany.GetDBServerDate().ToString("yyyyMMdd") + "')"

            oDBDetail1.ExecuteQuery(DetailsQry)
            objGrid.DataTable = oDBDetail1
            objGrid.AutoResizeColumns()

            For i As Integer = 0 To objGrid.Columns.Count - 1
                objGrid.Columns.Item(i).Editable = False
            Next
            'Dim oColumns3 As SAPbouiCOM.EditTextColumn = objGrid.Columns.Item("Customer")
            'oColumns3.LinkedObjectType = "2"

            Dim objRadio As SAPbouiCOM.OptionBtn = oForm.Items.Item("All").Specific
            objRadio.GroupWith("Success")
            objRadio.GroupWith("Failed")
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            objMain.objUtilities.Write_Log(ex.Message)
            objMain.objApplication.StatusBar.SetText("Problem Loading Values to Matrix  " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case "EMLLog"
                    If pVal.BeforeAction = True Then
                        objMain.objUtilities.LoadForm("EmailLog.xml", "EMLLog", ResourceType.Embeded)
                    Else
                        oForm = objMain.objApplication.Forms.ActiveForm
                        FormLoad(oForm)

                    End If
            End Select

        Catch ex As Exception
            objMain.objUtilities.Write_Log(ex.Message)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            oForm = objMain.objApplication.Forms.Item(pVal.FormUID)

            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Select Case pVal.ItemUID
                        Case "All"
                            Try
                                Try
                                    If pVal.BeforeAction = False Then

                                        oForm.Freeze(True)
                                        Dim rnd As New Random()
                                        Dim randomNumber As Integer = rnd.Next(10, 100)
                                        oDBDetail1 = oForm.DataSources.DataTables.Add("" + randomNumber.ToString() + "")
                                        oDBDetail1 = oForm.DataSources.DataTables.Item("" + randomNumber.ToString() + "")
                                        objGrid = oForm.Items.Item("grid").Specific
                                        Try
                                            If objGrid.Columns.Count > 0 Then
                                                objGrid.DataTable.Clear()
                                            End If
                                        Catch ex As Exception
                                            objMain.objUtilities.Write_Log(ex.Message)
                                        End Try
                                        oDBH = oForm.DataSources.DBDataSources.Item("OITM")
                                        Dim FROMDT As String = oDBH.GetValue("CreateDate", oDBH.Offset)
                                        If FROMDT = "" Then
                                            FROMDT = "19900101"
                                        End If
                                        Dim TODATE As String = oDBH.GetValue("UpdateDate", oDBH.Offset)
                                        If TODATE = "" Then
                                            TODATE = System.DateTime.Now.ToString("yyyyMMdd")
                                        End If
                                        Dim DetailsQry As String = "CALL ""TENX_MESSENGER_EMAIL_HISTORY"" ('FormLoad','','','" + FROMDT + "','" + TODATE + "')"
                                        oDBDetail1.ExecuteQuery(DetailsQry)
                                        objGrid.DataTable = oDBDetail1

                                        For i As Integer = 0 To objGrid.Columns.Count - 1
                                            objGrid.Columns.Item(i).Editable = False
                                        Next

                                        'Dim oColumns3 As SAPbouiCOM.EditTextColumn = objGrid.Columns.Item("Customer")
                                        'oColumns3.LinkedObjectType = "2"
                                        objGrid.AutoResizeColumns()

                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    oForm.Freeze(False)
                                    objMain.objUtilities.Write_Log(ex.Message)
                                    objMain.objApplication.StatusBar.SetText("Problem Loading Values to Matrix  " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End Try
                            Catch ex As Exception
                                objMain.objUtilities.Write_Log(ex.Message)
                                oForm.Freeze(False)
                            End Try
                        Case "Success"
                            Try
                                Try
                                    If pVal.BeforeAction = False Then
                                        oForm.Freeze(True)
                                        Dim rnd As New Random()
                                        Dim randomNumber As Integer = rnd.Next(10, 100)
                                        oDBDetail1 = oForm.DataSources.DataTables.Add("" + randomNumber.ToString() + "")
                                        oDBDetail1 = oForm.DataSources.DataTables.Item("" + randomNumber.ToString() + "")
                                        objGrid = oForm.Items.Item("grid").Specific
                                        Try
                                            If objGrid.Columns.Count > 0 Then
                                                objGrid.DataTable.Clear()
                                            End If
                                        Catch ex As Exception
                                            objMain.objUtilities.Write_Log(ex.Message)
                                        End Try
                                        oDBH = oForm.DataSources.DBDataSources.Item("OITM")
                                        Dim FROMDT As String = oDBH.GetValue("CreateDate", oDBH.Offset)
                                        If FROMDT = "" Then
                                            FROMDT = "19900101"
                                        End If
                                        Dim TODATE As String = oDBH.GetValue("UpdateDate", oDBH.Offset)
                                        If TODATE = "" Then
                                            TODATE = System.DateTime.Now.ToString("yyyyMMdd")
                                        End If
                                        Dim DetailsQry As String = "CALL ""TENX_MESSENGER_EMAIL_HISTORY"" ('ItemEvent_Success','','','" + FROMDT + "','" + TODATE + "')"
                                        oDBDetail1.ExecuteQuery(DetailsQry)
                                        objGrid.DataTable = oDBDetail1

                                        For i As Integer = 0 To objGrid.Columns.Count - 1
                                            objGrid.Columns.Item(i).Editable = False
                                        Next

                                        'Dim oColumns3 As SAPbouiCOM.EditTextColumn = objGrid.Columns.Item("Customer")
                                        'oColumns3.LinkedObjectType = "2"
                                        objGrid.AutoResizeColumns()

                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    oForm.Freeze(False)
                                    objMain.objUtilities.Write_Log(ex.Message)
                                    objMain.objApplication.StatusBar.SetText("Problem Loading Values to Matrix  " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End Try
                            Catch ex As Exception
                                objMain.objUtilities.Write_Log(ex.Message)
                                oForm.Freeze(False)
                            End Try
                        Case "Failed"
                            Try
                                Try
                                    If pVal.BeforeAction = False Then
                                        oForm.Freeze(True)
                                        Dim rnd As New Random()
                                        Dim randomNumber As Integer = rnd.Next(10, 100)
                                        oDBDetail1 = oForm.DataSources.DataTables.Add("" + randomNumber.ToString() + "")
                                        oDBDetail1 = oForm.DataSources.DataTables.Item("" + randomNumber.ToString() + "")
                                        objGrid = oForm.Items.Item("grid").Specific
                                        Try
                                            If objGrid.Columns.Count > 0 Then
                                                objGrid.DataTable.Clear()
                                            End If
                                        Catch ex As Exception
                                            objMain.objUtilities.Write_Log(ex.Message)
                                        End Try
                                        oDBH = oForm.DataSources.DBDataSources.Item("OITM")
                                        Dim FROMDT As String = oDBH.GetValue("CreateDate", oDBH.Offset)
                                        If FROMDT = "" Then
                                            FROMDT = "19900101"
                                        End If
                                        Dim TODATE As String = oDBH.GetValue("UpdateDate", oDBH.Offset)
                                        If TODATE = "" Then
                                            TODATE = System.DateTime.Now.ToString("yyyyMMdd")
                                        End If
                                        Dim DetailsQry As String = "CALL ""TENX_MESSENGER_EMAIL_HISTORY"" ('ItemEvent_Failed','','','" + FROMDT + "','" + TODATE + "')"
                                        oDBDetail1.ExecuteQuery(DetailsQry)
                                        objGrid.DataTable = oDBDetail1

                                        For i As Integer = 0 To objGrid.Columns.Count - 1
                                            If i = 0 Then
                                                objGrid.Columns.Item(i).Editable = True
                                            Else
                                                objGrid.Columns.Item(i).Editable = False
                                            End If
                                        Next
                                        Dim objColumn As SAPbouiCOM.EditTextColumn = objGrid.Columns.Item("Check")
                                        objColumn.Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox

                                        Dim oColumns3 As SAPbouiCOM.EditTextColumn = objGrid.Columns.Item("Customer")
                                        oColumns3.LinkedObjectType = "2"
                                        objGrid.AutoResizeColumns()
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    oForm.Freeze(False)
                                    objMain.objUtilities.Write_Log(ex.Message)
                                    objMain.objApplication.StatusBar.SetText("Problem Loading Values to Matrix  " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End Try
                            Catch ex As Exception
                                objMain.objUtilities.Write_Log(ex.Message)
                                oForm.Freeze(False)
                            End Try
                        Case "BTNSEND"
                            Try
                                If pVal.BeforeAction = False Then
                                    oForm.Freeze(True)
                                    Dim Failed As SAPbouiCOM.OptionBtn = oForm.Items.Item("Failed").Specific
                                    If Failed.Selected = True Then
                                        objGrid = oForm.Items.Item("grid").Specific
                                        Try
                                            For i As Integer = 0 To objGrid.Rows.Count
                                                Dim oCheck As String = objGrid.DataTable.GetValue("Check", objGrid.GetDataTableRowIndex(i)).ToString()
                                                If oCheck = "Y" Then
                                                    Dim DocType As String = objGrid.DataTable.GetValue("Document Type", objGrid.GetDataTableRowIndex(i)).ToString()
                                                    Dim DocEntry As String = objGrid.DataTable.GetValue("Document Entry", objGrid.GetDataTableRowIndex(i)).ToString()
                                                    MarketingEmail(DocEntry, DocType)
                                                End If
                                            Next
                                        Catch ex As Exception
                                            objMain.objUtilities.Write_Log(ex.Message)
                                        End Try
                                    End If
                                    oForm.Freeze(False)
                                End If
                            Catch ex As Exception
                                objMain.objUtilities.Write_Log(ex.Message)
                                oForm.Freeze(False)
                            End Try
                    End Select

            End Select
        Catch ex As Exception
            oForm.Freeze(False)
            If (ex.Message = "Form - Invalid Form") Then
            Else
                objMain.objUtilities.Write_Log(ex.Message)
                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        End Try
    End Sub

    Sub MarketingEmail(ByVal DocEntry As String, ByVal DocType As String)
        Dim DetQry As String = ""
        Dim WMSGTXT As String = ""
        Dim MNTXT As String = ""
        If DocType = "SO" Then
            MNTXT = "ORDER CONFIRMATION"
            WMSGTXT = "Thank You... Your Sales order has been placed."
            DetQry = "CALL ""TENX_MESSENGER_EMAIL_HISTORY"" ('MarketingEmail','SO','" + DocEntry + "')"
        ElseIf DocType = "PO" Then
            MNTXT = "PO CONFIRMATION"
            WMSGTXT = "Thank You... Your Purchase order has been placed."
            DetQry = "CALL ""TENX_MESSENGER_EMAIL_HISTORY"" ('MarketingEmail','PO','" + DocEntry + "')"
        ElseIf DocType = "SQ" Then
            MNTXT = "QUOTATION CONFIRMATION"
            WMSGTXT = "Thank You... Your Sales Quotation has been placed."
            DetQry = "CALL ""TENX_MESSENGER_EMAIL_HISTORY"" ('MarketingEmail','SQ','" + DocEntry + "')"
        ElseIf DocType = "DL" Then
            MNTXT = "DELIVERY CONFIRMATION"
            WMSGTXT = "Thank You... Your Delivery Note has been placed."
            DetQry = "CALL ""TENX_MESSENGER_EMAIL_HISTORY"" ('MarketingEmail','DL','" + DocEntry + "')"
        ElseIf DocType = "ARINV" Then
            MNTXT = "INVOICE CONFIRMATION"
            WMSGTXT = "Thank You... Your Sales Invoice has been placed."
            DetQry = "CALL ""TENX_MESSENGER_EMAIL_HISTORY"" ('MarketingEmail','ARINV','" + DocEntry + "')"
        ElseIf DocType = "ARC" Then
            MNTXT = "CREDIT MEMO CONFIRMATION "
            WMSGTXT = "Thank You... Your A/R Credit Note has been placed."
            DetQry = "CALL ""TENX_MESSENGER_EMAIL_HISTORY"" ('MarketingEmail','ARC','" + DocEntry + "')"
        ElseIf DocType = "ARRET" Then
            MNTXT = "RESERVE INVOICE CONFIRMATION"
            WMSGTXT = "Thank You... Your Sales Invoice has been placed."
            DetQry = "CALL ""TENX_MESSENGER_EMAIL_HISTORY"" ('MarketingEmail','ARRET','" + DocEntry + "')"
        ElseIf DocType = "PQ" Then
            MNTXT = "PQ CONFIRMATION "
            WMSGTXT = "Thank You... Your Purchase Quotaion has been placed."
            DetQry = "CALL ""TENX_MESSENGER_EMAIL_HISTORY"" ('MarketingEmail','PQ','" + DocEntry + "')"
        ElseIf DocType = "APINV" Then
            MNTXT = "A/P INVOICE CONFIRMATION "
            WMSGTXT = "Thank You... Your Purchase Invoice has been placed."
            DetQry = "CALL ""TENX_MESSENGER_EMAIL_HISTORY"" ('MarketingEmail','APINV','" + DocEntry + "')"
        ElseIf DocType = "GRPO" Then
            MNTXT = "GRPO CONFIRMATION "
            WMSGTXT = "Thank You... Your Goods Return Against PO has been placed."
            DetQry = "CALL ""TENX_MESSENGER_EMAIL_HISTORY"" ('MarketingEmail','GRPO','" + DocEntry + "')"
        ElseIf DocType = "GR" Then
            MNTXT = "GOODS RETURN CONFIRMATION "
            WMSGTXT = "Thank You... Your Goods Return has been placed."
            DetQry = "CALL ""TENX_MESSENGER_EMAIL_HISTORY"" ('MarketingEmail','GR','" + DocEntry + "')"
        ElseIf DocType = "APC" Then
            MNTXT = "A/P CREDIT MEMO CONFIRMATION "
            WMSGTXT = "Thank You... Your A/P Credit Note has been placed."
            DetQry = "CALL ""TENX_MESSENGER_EMAIL_HISTORY"" ('MarketingEmail','APC','" + DocEntry + "')"
        ElseIf DocType = "IP" Then
            MNTXT = "INCOMING PAYMENT CONFIRMATION "
            WMSGTXT = "Thank You... Your Incomming Payment has been placed."
            DetQry = "CALL ""TENX_MESSENGER_EMAIL_HISTORY"" ('MarketingEmail','IP','" + DocEntry + "')"
        ElseIf DocType = "OP" Then
            MNTXT = "OUTGOING PAYMENT CONFIRMATION "
            WMSGTXT = "Thank You... Your Outgoing payment  been placed."
            DetQry = "CALL ""TENX_MESSENGER_EMAIL_HISTORY"" ('MarketingEmail','OP','" + DocEntry + "')"
        ElseIf DocType = "GRP" Then
            MNTXT = "GOODS RECEIPT CONFIRMATION "
            WMSGTXT = "Thank You... Your Good Receipt has been placed."
            DetQry = "CALL ""TENX_MESSENGER_EMAIL_HISTORY"" ('MarketingEmail','GRP','" + DocEntry + "')"
        ElseIf DocType = "GI" Then
            MNTXT = "GOODS ISSUE CONFIRMATION "
            WMSGTXT = "Thank You... Your Good Issue has been placed."
            DetQry = "CALL ""TENX_MESSENGER_EMAIL_HISTORY"" ('MarketingEmail','GI','" + DocEntry + "')"
        ElseIf DocType = "ITR" Then
            MNTXT = "INVENTORY TRANSFER REQUEST CONFIRMATION"
            WMSGTXT = "Inventory Transfer Request is created."
            DetQry = "CALL ""TENX_MESSENGER_EMAIL_HISTORY"" ('MarketingEmail','ITR','" + DocEntry + "')"
        ElseIf DocType = "IT" Then
            MNTXT = "INVENTORY TRANSFER CONFIRMATION"
            WMSGTXT = "Thank You... Your Inventory Tranfer has been placed."
            DetQry = "CALL ""TENX_MESSENGER_EMAIL_HISTORY"" ('MarketingEmail','IT','" + DocEntry + "')"
        ElseIf DocType = "JE" Then
            MNTXT = "JOURNAL ENTRY CONFIRMATION"
            WMSGTXT = "Thank You... Your JE has been placed."
            DetQry = "CALL ""TENX_MESSENGER_EMAIL_HISTORY"" ('MarketingEmail','JE','" + DocEntry + "')"
        End If

        Dim oRsDet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRsDet.DoQuery(DetQry)
        Dim Msg As String = objMain.objUtilities.getMessage(DocType, "Email")
        If Msg <> "" Then
            If oRsDet.RecordCount > 0 Then
                Dim Result As String = ""
                Dim ToEmail As String = oRsDet.Fields.Item("Email").Value

                Dim TotalBeforDis As Double = oRsDet.Fields.Item("DocTotal").Value + oRsDet.Fields.Item("DiscSum").Value - oRsDet.Fields.Item("VatSum").Value
                Dim CardName As String = oRsDet.Fields.Item("CardName").Value.ToString()
                Dim DocNum1 As String = oRsDet.Fields.Item("DocNum").Value.ToString()
                Dim TaxDate As String = oRsDet.Fields.Item("TaxDate").Value.ToString()
                Dim ItemCount As String = oRsDet.Fields.Item("ItemCount").Value.ToString()
                Dim TotalBeforDis1 As String = TotalBeforDis.ToString()
                Dim VatSum As String = oRsDet.Fields.Item("VatSum").Value.ToString()
                Dim DocTotal As String = oRsDet.Fields.Item("DocTotal").Value.ToString()


                Msg = Msg.Replace("{CardName}", oRsDet.Fields.Item("CardName").Value.ToString())
                Msg = Msg.Replace("{DocNum}", oRsDet.Fields.Item("DocNum").Value.ToString())
                Msg = Msg.Replace("{TaxDate}", oRsDet.Fields.Item("TaxDate").Value.ToString())
                Msg = Msg.Replace("{ItemCount}", oRsDet.Fields.Item("ItemCount").Value.ToString())
                Msg = Msg.Replace("{Total}", TotalBeforDis.ToString())
                Msg = Msg.Replace("{VatSum}", oRsDet.Fields.Item("VatSum").Value.ToString())
                Msg = Msg.Replace("{DocTotal}", oRsDet.Fields.Item("DocTotal").Value.ToString())
                If ToEmail <> "" Then
                    Result = Me.SendEmail_TEST(ToEmail, MNTXT, WMSGTXT, oRsDet.Fields.Item("DocEntry").Value.ToString(), DocType, CardName, DocNum1, TaxDate, ItemCount, TotalBeforDis1, VatSum, DocTotal)
                    If Result = "Success" Then
                        objMain.objUtilities.AddDataToLogTable(DocType, oRsDet.Fields.Item("TaxDate").Value.ToString(), oRsDet.Fields.Item("CardCode").Value.ToString(), "Success", Msg, ToEmail, "", DocNum1, oRsDet.Fields.Item("DocEntry").Value.ToString(), Result, "Email")
                    Else
                        objMain.objUtilities.AddDataToLogTable(DocType, oRsDet.Fields.Item("TaxDate").Value.ToString(), oRsDet.Fields.Item("CardCode").Value.ToString(), "Failed", Msg, ToEmail, "", DocNum1, oRsDet.Fields.Item("DocEntry").Value.ToString(), Result, "Email")
                    End If
                Else
                    objMain.objUtilities.AddDataToLogTable(DocType, oRsDet.Fields.Item("TaxDate").Value.ToString(), oRsDet.Fields.Item("CardCode").Value.ToString(), "Failed", Msg, ToEmail, "", DocNum1, oRsDet.Fields.Item("DocEntry").Value.ToString(), Result, "Email")
                End If

                Dim AddContactsQry As String = "CALL ""TENX_MESSENGER_MAIN"" ('EmailUserEmp','" + DocType + "')"
                Dim oRsAddContacts As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsAddContacts.DoQuery(AddContactsQry)
                If oRsAddContacts.RecordCount > 0 Then
                    For i As Integer = 1 To oRsAddContacts.RecordCount
                        If oRsAddContacts.Fields.Item(0).Value.ToString() <> "" Then
                            Result = Me.SendEmail_TEST(oRsAddContacts.Fields.Item(0).Value.ToString(), MNTXT, WMSGTXT, oRsDet.Fields.Item("DocEntry").Value.ToString(), DocType, CardName, DocNum1, TaxDate, ItemCount, TotalBeforDis1, VatSum, DocTotal)
                            If Result = "Success" Then
                                objMain.objUtilities.UpdateDataToLogTable(DocType, oRsDet.Fields.Item("DocEntry").Value.ToString(), "Success", "Email", "")
                            Else
                                objMain.objUtilities.UpdateDataToLogTable(DocType, oRsDet.Fields.Item("DocEntry").Value.ToString(), "Failed", "Email", objMain.Error_Message)
                            End If
                        End If
                        oRsAddContacts.MoveNext()
                    Next
                End If
            End If
        Else
            objMain.objApplication.StatusBar.SetText("Message Configuration Not Set in Message Template", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If
    End Sub
    Function SendEmail_TEST(ByVal ToEmail As String, ByVal mainTxt As String, ByVal WMSGTXT As String, ByVal DocEntry As String, ByVal DocType As String, ByVal CardName As String, ByVal DocNum1 As String, ByVal TaxDate As String, ItemCount As String, ByVal TotalBeforDis1 As String, ByVal VatSum As String, ByVal DocTotal As String) As String
        Try
            If ToEmail = "" Then
                objMain.objUtilities.Write_Log("No Email ID for this customer")
            Else
                Dim EmailQry As String = "SELECT T0.""U_SMTPsevr"", T0.""U_SMPTPort"", T0.""U_EmailFro"", T0.""U_EmailPwd"", T0.""U_CC"", T0.""U_BCC"", T0.""U_SSL"" FROM ""@SBO_SETCONFIG""  T0"
                Dim oRsEmail As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsEmail.DoQuery(EmailQry)
                Dim ServerHost As String = oRsEmail.Fields.Item("U_SMTPsevr").Value.ToString()
                Dim PortNo As String = oRsEmail.Fields.Item("U_SMPTPort").Value.ToString()
                Dim FromMail As String = oRsEmail.Fields.Item("U_EmailFro").Value.ToString()
                Dim EmailPassword As String = oRsEmail.Fields.Item("U_EmailPwd").Value.ToString()
                ServicePointManager.SecurityProtocol = DirectCast(3072, System.Net.SecurityProtocolType) 'TLS 1.2
                Dim SMTPServer As New SmtpClient
                Dim Message As New MailMessage(FromMail, ToEmail)
                If oRsEmail.Fields.Item("U_SSL").Value = "Y" Then
                    SMTPServer.EnableSsl = True
                Else
                    SMTPServer.EnableSsl = False
                End If
                SMTPServer.Host = ServerHost
                SMTPServer.Port = PortNo
                SMTPServer.DeliveryMethod = SmtpDeliveryMethod.Network
                SMTPServer.UseDefaultCredentials = False
                SMTPServer.Credentials = New Net.NetworkCredential(FromMail, EmailPassword)
                Dim GetMessageTempDtls As String = "Select ""U_Sub"" From ""@MGRCONFIG"" Where ""U_Doc"" = '" & DocType & "' And IFNULL(""U_Status"",'N') = 'Y' And Ifnull(""U_Sub"",'N') <> 'N' "
                Dim oRsGetMessageTempDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetMessageTempDtls.DoQuery(GetMessageTempDtls)
                If oRsGetMessageTempDtls.RecordCount > 0 Then
                    Message.Subject = oRsGetMessageTempDtls.Fields.Item(0).Value
                Else
                    objMain.objApplication.StatusBar.SetText("Please provide subject in message template ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return "Fail"
                    Exit Function
                End If
                Dim Attachment As String = Nothing
                Dim Base64 As String = Nothing
                Dim AttachmentNew As Attachment
                'Dim Sign As String = Nothing
                'Dim sigReader As StreamReader
                Dim AtttM As String = objMain.objUtilities.AttachmnetSelection(DocType, "Email")
                If AtttM = 1 Then
                    Dim AttPath As String = objMain.objUtilities.GetAttachmnetPath()
                    Dim CrystalID As String = objMain.objUtilities.GetCrystalID(DocType, "Email")
                    Dim Signature As String = objMain.objUtilities.GetSignature(DocType, "Email")
                    'objMain.objUtilities.GetManualAttachment(DocEntry, CrystalID, DocType, "")
                    'Dim Attachment1 As String = AttPath & DocType & "_" & DocEntry & "_" & System.DateTime.Now.ToString("yyyyMMdd") & ".pdf"
                    'Message.Attachments.Add(New System.Net.Mail.Attachment(Attachment1))
                    'Attachment = "Please find the attached document. "
                    'Base64 = objMain.ConvertFileToBase64(AttPath + Signature)

                    'AttachmentNew = New Attachment(AttPath + Signature)
                    'AttachmentNew.ContentId = "NewImage"

                    'Sign = ReadSignature(AttPath)
                    'sigReader = New StreamReader(Sign)
                Else
                    Try
                        If objMain.objUtilities.CheckAttachment_EMAIL(DocType, "Email") = True Then
                            Dim temp As String = objMain.objUtilities.getAttachment(DocType, DocEntry)
                            Message.Attachments.Add(New System.Net.Mail.Attachment(temp))
                            Attachment = "Please find the attached document. "

                        End If
                    Catch ex As Exception
                        objMain.objUtilities.Write_Log(ex.Message)
                    End Try
                End If
                Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString("<table> <tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'><b>" + mainTxt + "<b> <br></td></tr>" _
                                                                                          + "<tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'><b>" + CardName + "<b><br></td></tr>" _
                                                                                          + "<tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'>Order Summary </td></tr>" _
                                                                                          + "<tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'>Order Number : <b>" + DocNum1 + "<b></td></tr>" _
                                                                                          + "<tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'>Order Date : <b>" + TaxDate + " <b></td></tr>" _
                                                                                          + "<tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'>Product/SKU's Count : <b>" + ItemCount + "<b></td></tr>" _
                                                                                          + "<tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'>Total before Discount : <b>" + TotalBeforDis1 + "<b></td></tr>" _
                                                                                          + "<tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'>Vat Amount : <b>" + VatSum + "<b> </td></tr>" _
                                                                                          + "<tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'>Total Amount : <b>" + DocTotal + "<b></td></tr>" _
                                                                                          + "<tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'>" + WMSGTXT + "<br><br></td></tr>" _
                                                                                          + "<tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'>" + Attachment + "<br><br></td></tr></table>
                                                                                          <div><img src='cid:NewImage' alt='Email Image' /></div>
                                                                                          <table><tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'>This is system auto generated mail. Please do not reply. </td></tr></table>", Nothing, "text/html")

                Message.AlternateViews.Add(htmlView)
                ' Message.Attachments.Add(AttachmentNew)
                SMTPServer.Send(Message)

                Message.Dispose()
                objMain.objUtilities.Write_Log("E-Mail Sent successfully to -  " & ToEmail)
                objMain.objApplication.StatusBar.SetText("E-Mail Sent successfully to -  " & ToEmail, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Return "Success"
            End If
        Catch ex As Exception
            objMain.Error_Message = "Error In Email Sending:" & ex.Message
            objMain.objUtilities.Write_Log("Error In Email Sending:" & ex.Message)
            objMain.objApplication.StatusBar.SetText("Error In Email Sending:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
    End Function

End Class

