Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Mail
Imports System.Text

Public Class MessengerConfiguration
    Private oForm As SAPbouiCOM.Form
    Dim oDBHeader, oDBDetail1 As SAPbouiCOM.DBDataSource
    Dim oItem, oItemUP, oItemUP1 As SAPbobsCOM.Items
    Dim Error_Code As Integer = 1
    Dim Error_Msg As String = ""
    Private sISNULL As String = "IFNULL"
    Private Sub DataSourcesLoad(ByRef oForm As SAPbouiCOM.Form)
        Try
            oDBHeader = oForm.DataSources.DBDataSources.Item("@SBO_SETCONFIG")
        Catch ex As Exception
            objMain.objUtilities.Write_Log(ex.Message)
        End Try
    End Sub
    Sub FormLoad(ByRef oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            Dim Query As String = "SELECT * FROM ""@SBO_SETCONFIG"" "
            Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rs.DoQuery(Query)
            If rs.RecordCount = 0 Then
                oDBHeader.SetValue("DocNum", 0, GetNextDocNum())
            Else
                objMain.objApplication.ActivateMenuItem("1290")
            End If
            oForm.Items.Item("Email").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case "SBO_SETCONFIG"
                    If pVal.BeforeAction = True Then
                        objMain.objUtilities.LoadForm("SMS_Email_Configuration.xml", "SBO_SETCONFIG", ResourceType.Embeded)
                    Else
                        oForm = objMain.objApplication.Forms.ActiveForm
                        DataSourcesLoad(oForm)
                        FormLoad(oForm)
                        oForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        oForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
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
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Select Case pVal.ItemUID
                        Case "TestWhats"
                            Try
                                If pVal.BeforeAction = False And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    oForm.Freeze(True)
                                    Dim mobile As String = ""
                                    Dim Msg As String = "Test Message from WhatsApp Configuration Form"
                                    Dim Link As String = oForm.Items.Item("Item_40").Specific.Value
                                    Dim Authencate As String = oForm.Items.Item("Item_7").Specific.Value
                                    mobile = oForm.Items.Item("Item_8").Specific.Value
                                    mobile = mobile.Replace(" ", "")
                                    mobile = mobile.Replace("-", "")
                                    Dim result2 As String = mobile
                                    If mobile.Length = 10 AndAlso Left(mobile, 1) <> 0 Then
                                        mobile = "+91" + mobile
                                    ElseIf mobile.Length > 9 Then
                                        result2 = mobile.Substring(mobile.Length - 9)
                                        mobile = "+971" + result2
                                    End If

                                    SendWhatsApp(Link, Authencate, mobile)
                                    oForm.Freeze(False)
                                End If
                            Catch ex As Exception
                                oForm.Freeze(False)
                            End Try
                        Case "TestSMS"
                            Try
                                If pVal.BeforeAction = False And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    oForm.Freeze(True)
                                    Dim SMSURL As String = ""
                                    Dim SMSUID As String = ""
                                    Dim Password As String = ""
                                    Dim mobile As String = ""

                                    Dim Msg As String = "Test Message from SMS Configuration Form"

                                    SMSURL = oForm.Items.Item("Item_13").Specific.Value
                                    SMSUID = oForm.Items.Item("Item_1").Specific.Value
                                    Password = oForm.Items.Item("Item_3").Specific.Value
                                    mobile = oForm.Items.Item("40").Specific.Value

                                    SendSMS(mobile, Msg)
                                    oForm.Freeze(False)
                                End If
                            Catch ex As Exception
                                oForm.Freeze(False)
                            End Try
                        Case "TestEmail"
                            Try
                                If pVal.BeforeAction = False And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    oForm.Freeze(True)
                                    Dim Msg As String = "Test Message from E-Mail Configuration Form"
                                    SendEmail(Msg)
                                    oForm.Freeze(False)
                                End If
                            Catch ex As Exception
                                oForm.Freeze(False)
                            End Try
                    End Select

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
    Public Function GetNextDocNum() As Integer
        Dim Str As String
        Dim oRs As SAPbobsCOM.Recordset
        Dim DocNum As Integer
        oRs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If objMain.objCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            Str = "SELECT " + IsNull + "(MAX(TO_INTEGER(""DocNum"")),0)+1 FROM ""@SBO_SETCONFIG"""
        Else
            Str = "SELECT ISNULL(MAX(CAST(""DocNum"" AS int)),0)+1 FROM ""@SBO_SETCONFIG"""
        End If

        Try
            oRs.DoQuery(Str)
            oRs.MoveFirst()
            If oRs.RecordCount > 0 Then
                DocNum = oRs.Fields.Item(0).Value
            End If
            If DocNum = 0 Then DocNum = 1
            Return DocNum
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function SendWhatsApp(ByVal Link As String, ByVal Authorizer As String, ByVal mobile As String) As String
        Dim Responce As String = ""
        Try
            SendFormData(Link, Authorizer, mobile).Wait()
        Catch ex As Exception

        End Try
        Return Responce
    End Function
    Function SendSMS(ByVal toMobileNo As String, ByVal Message As String) As String
        Dim SMSResult As String
        Dim request As HttpWebRequest
        Dim response As HttpWebResponse = Nothing
        Dim url As String
        Try
            toMobileNo = toMobileNo.Replace(" ", "")
            toMobileNo = toMobileNo.Replace("-", "")
            Dim result As String = toMobileNo
            If toMobileNo.Length = 10 AndAlso Left(toMobileNo, 1) <> 0 Then
                toMobileNo = "91" + toMobileNo
            ElseIf toMobileNo.Length > 9 Then
                result = toMobileNo.Substring(toMobileNo.Length - 9)
                toMobileNo = "971" + result
            End If


            'url = "http://46.101.248.41:1401/send?username=kenz&password=kEn@z!&from=kenz hyper&to=971523706510&content=TestSMS"
            'url = "http://46.101.248.41:1401/send?username=UserName&password=Password&from=SenderID&to=MobileNo&content=Message"

            Dim CredQry As String = "Select ""U_SMSUrl"",""U_SMSUId"",""U_SMSPwd"" From ""@SBO_SETCONFIG"" "
            Dim oRsCred As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsCred.DoQuery(CredQry)

            Dim UserName As String = oRsCred.Fields.Item("U_SMSUid").Value
            Dim Password As String = oRsCred.Fields.Item("U_SMSPwd").Value

            url = "https://api.brio.ae/send?username=kenz&password=kEn@z1Bc&from=KENZ HYPR&to=" + toMobileNo + "&content=" + Message + ""
            ''
            request = DirectCast(WebRequest.Create(url), HttpWebRequest)

            response = DirectCast(request.GetResponse(), HttpWebResponse)

            Dim enc As Encoding = System.Text.Encoding.GetEncoding(1252)

            Dim ResponseStr As String
            Using loResponseStream As StreamReader = New StreamReader(response.GetResponseStream(), enc)
                ResponseStr = loResponseStream.ReadToEnd()
            End Using
            response.Close()
            If ResponseStr.Contains("Success") Then
                objMain.objApplication.StatusBar.SetText("SMS Sent To:" & toMobileNo & "Message:" & Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                SMSResult = "Success"
            End If
            Return SMSResult
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function
    Function SendEmail(ByVal mainTxt As String) As String
        Try

            Dim EmailQry As String = "SELECT T0.""U_SMTPsevr"", T0.""U_SMPTPort"", T0.""U_EmailFro"", T0.""U_EmailPwd"", T0.""U_CC"", T0.""U_BCC"", T0.""U_SSL"" FROM ""@SBO_SETCONFIG""  T0"
            Dim oRsEmail As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsEmail.DoQuery(EmailQry)
            Dim ServerHost As String = oRsEmail.Fields.Item("U_SMTPsevr").Value.ToString()
            Dim PortNo As String = oRsEmail.Fields.Item("U_SMPTPort").Value.ToString()
            Dim FromMail As String = oRsEmail.Fields.Item("U_EmailFro").Value.ToString()
            Dim EmailPassword As String = oRsEmail.Fields.Item("U_EmailPwd").Value.ToString()
            Dim ToEmail As String = oRsEmail.Fields.Item("U_CC").Value.ToString()
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

            Message.Subject = "TEST"

            Dim Attachment As String = Nothing
            Message.Body = mainTxt
            SMTPServer.Send(Message)
            Message.Dispose()
            Try
                objMain.objUtilities.Write_Log("Email Sent Successfully to - " & ToEmail & "- On -" & Today.ToString("dd-MM-yyyy") & " ")
            Catch ex As Exception

            End Try
            objMain.objApplication.StatusBar.SetText("E-Mail Sent successfully to -  " & ToEmail, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Return "Success"
        Catch ex As Exception
            objMain.objUtilities.Write_Log("Error In Email Sending:" & ex.Message)
            objMain.objApplication.StatusBar.SetText("Error In Email Sending:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        End Try
    End Function
    Async Function SendFormData(ByVal Link As String, ByVal Authorizer As String, ByVal mobile As String) As Task
        Dim httpClient As New HttpClient()
        Dim url As String = "" + Link + ""
        httpClient.DefaultRequestHeaders.Add("Authorization", "" + Authorizer + "")

        Dim formData As New List(Of KeyValuePair(Of String, String))()
        formData.Add(New KeyValuePair(Of String, String)("to", "" + mobile + ""))
        formData.Add(New KeyValuePair(Of String, String)("type", "template"))
        formData.Add(New KeyValuePair(Of String, String)("template[name]", "order_confirmation_2"))
        formData.Add(New KeyValuePair(Of String, String)("template[language][code]", "en"))
        formData.Add(New KeyValuePair(Of String, String)("template[components][1][type]", "body"))
        formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][0][type]", "text"))
        formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][0][text]", "Company ABC"))
        formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][1][type]", "text"))
        formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][1][text]", "12345"))
        formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][2][type]", "text"))
        formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][2][text]", "2024-01-01"))
        formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][3][type]", "text"))
        formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][3][text]", "2"))
        formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][4][type]", "text"))
        formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][4][text]", "12067"))
        formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][5][type]", "text"))
        formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][5][text]", "120"))
        formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][6][type]", "text"))
        formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][6][text]", "23456"))

        Dim content As New FormUrlEncodedContent(formData)

        Try
            Dim response As HttpResponseMessage = Await httpClient.PostAsync(url, content)

            If response.IsSuccessStatusCode Then
                Dim responseContent As String = Await response.Content.ReadAsStringAsync()
                Console.WriteLine("Response Content: " & responseContent)
                objMain.objUtilities.ShowSucessMessage("WhatsApp msg Sent Successfully !!!!")
            Else
                Console.WriteLine("Failed to send form data. Status code: " & response.StatusCode)
                objMain.objUtilities.ShowErrorMessage("Failed to send form data !!!!")
            End If
        Catch ex As Exception
            Console.WriteLine("An error occurred: " & ex.Message)
        End Try
    End Function
End Class
