
Imports System.Configuration
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Reflection
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TreeView
Imports SAPbobsCOM

Public Class Utilities
    Public strLastErrorCode As String
    Public strLastError As String
    Private objForm As SAPbouiCOM.Form
    Dim oItem As SAPbouiCOM.Item
    'Public objCompanyService As SAPbobsCOM.CompanyService
    'Public oCompanyAdminInfo As SAPbobsCOM.AdminInfo
    Dim TempCode As String
    Dim BankFileName As String
    Public GlobalItemNumber As String
    Dim UIConn As Boolean = False
    Dim RunTime As String = Now.ToString("HH_mm_ss") : Dim time As String = ""

    Private sISNULL As String = "IFNULL"


    Public Sub Write_Log(ByVal status As String)
        Dim fs As FileStream
        Dim objWriter As System.IO.StreamWriter
        Dim chatlog As String
        Try
            Dim diYest As DirectoryInfo = New DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + "ErrorLog\ItemCreation_" & Today.AddDays(-1).ToString("yyyyMMdd") & "")
            If diYest.Exists Then
                diYest.Delete(True)
            End If
            If time = "" Then time = Today.ToString("yyyyMMdd") & "\" & Runtime
            Dim di As DirectoryInfo = New DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + "ErrorLog\ItemCreation_" & Today.ToString("yyyyMMdd") & "")
            If di.Exists Then
            Else
                di.Create()
            End If
            chatlog = AppDomain.CurrentDomain.BaseDirectory + "ErrorLog\ItemCreation_" & time & ".txt"
            If File.Exists(chatlog) Then
            Else
                fs = New FileStream(chatlog, FileMode.Create, FileAccess.Write)
                fs.Close()
            End If
            objWriter = New System.IO.StreamWriter(chatlog, True)
            If status <> "" Then objWriter.WriteLine(Now & " : " & status)
            objWriter.Close()
        Catch ex As Exception
            ShowErrorMessage(ex.Message)
        End Try
    End Sub

    Public Class RSDBObj
        Public CompanyDB As String
        Public DBInstance As String
        Public Password As String
        Public UserName As String
    End Class

#Region " Get Application "
    Public Function GetApplication() As SAPbouiCOM.Application
        Dim objApp As SAPbouiCOM.Application
        Try
            Dim objSboGuiApi As New SAPbouiCOM.SboGuiApi
            Dim strConnectionString As String = String.Empty
            If strConnectionString = "" Then
                If Environment.GetCommandLineArgs().Length = 1 Then
                    strConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"
                Else
                    strConnectionString = Environment.GetCommandLineArgs.GetValue(1)
                End If
            End If
            'strConnectionString = Environment.GetCommandLineArgs.GetValue(1)
            objSboGuiApi = New SAPbouiCOM.SboGuiApi
            objSboGuiApi.Connect(strConnectionString)
            objApp = objSboGuiApi.GetApplication()
            Return objApp
        Catch ex As Exception
            If (ex.Message = "UI API Server - Server is Down") Then
                If UIConn = False Then
                    For Each prog As Process In Process.GetProcesses
                        If prog.ProcessName = "SAPbouiCOM" Then
                            prog.Kill()
                        End If
                    Next
                    UIConn = True
                End If
            End If
            MessageBox.Show(ex.Message)
            System.Windows.Forms.Application.Exit()
            Return Nothing
        End Try
    End Function
#End Region

#Region " Get Company "
    Public Function GetCompany(ByVal SBOApplication As SAPbouiCOM.Application) As SAPbobsCOM.Company
        Dim objCompany As SAPbobsCOM.Company

        Dim strCookie As String
        Dim strCookieContext As String

        Try
            objCompany = New SAPbobsCOM.Company
            strCookie = objCompany.GetContextCookie
            strCookieContext = SBOApplication.Company.GetConnectionContext(strCookie)
            objCompany.SetSboLoginContext(strCookieContext)
            If objCompany.Connect <> 0 Then
                strLastError = "Connection Error"
                SBOApplication.StatusBar.SetText("Connection Error", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return Nothing
            End If

            Return objCompany
        Catch ex As Exception
            SBOApplication.MessageBox(ex.Message)
            Return Nothing
        End Try
    End Function
#End Region

    ''' <summary>
    ''' This will take the XML string and modify the form left/width, top/height to match the current client font size
    ''' </summary>
    ''' <param name="Original">Original XML String</param>
    ''' <returns>modified xml string</returns>
    ''' 
    Public Function ModifySize(ByVal Original As String) As String
        Dim oTypes As New System.Collections.Specialized.OrderedDictionary()
        Dim dFontHeightRatio As Double = Math.Round(objMain.objApplication.GetFormItemDefaultHeight(SAPbouiCOM.BoFormSizeableItemTypes.fsit_EDIT) / 14.0, 2)          'Ratio is based on Edit text item. 14.00 is the reference Height that i created the forms in
        Dim dFontWidthRatio As Double = Math.Round(objMain.objApplication.GetFormItemDefaultWidth(SAPbouiCOM.BoFormSizeableItemTypes.fsit_EDIT) / 120.0, 2)    'Ratio is based on Edit text item. 80.00 is the reference Width that i created the forms in
        oTypes.Add("left", dFontWidthRatio)
        oTypes.Add("width", dFontWidthRatio)
        oTypes.Add("top", dFontHeightRatio)
        oTypes.Add("height", dFontHeightRatio)

        For Each Type As String In oTypes.Keys
            Dim i As Integer = 0
            Dim Ratio As Double = oTypes(Type)
            While (i < Original.Length)
                i = Original.IndexOf(Type + "=""", i)
                If (i > 0) Then
                    Dim iNextApos As Integer = Original.IndexOf("""", i + Type.Length + 2)
                    Dim sContent As String = Original.Substring(i + Type.Length + 2, iNextApos - (i + Type.Length + 2))
                    Dim iContent As Integer = 0
                    If (Integer.TryParse(sContent, iContent)) Then
                        Original = Original.Substring(0, i) + Type + "=""" + Convert.ToInt16(iContent * Ratio).ToString() + Original.Substring(iNextApos)
                    End If
                    i = iNextApos
                Else
                    i = Original.Length
                End If
            End While
        Next
        Return Original
    End Function


#Region " Load Form "
    'Public Sub LoadForm(ByVal XMLFile As String, ByVal FormType As String, Optional ByVal FileType As ResourceType = ResourceType.Content)
    '    Try
    '        Dim AppAssemblty As Assembly = Assembly.GetExecutingAssembly()
    '        Dim sExecutingAssemblyNmae As String = AppAssemblty.GetName().Name.ToString()
    '        Dim xmldoc As New Xml.XmlDocument
    '        XMLFile = sExecutingAssemblyNmae + "." + XMLFile

    '        '// load the form to the SBO application in one batch
    '        Dim Streaming As System.IO.Stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(XMLFile)
    '        Dim StreamRead As New System.IO.StreamReader(Streaming, True)
    '        xmldoc.LoadXml(StreamRead.ReadToEnd)
    '        StreamRead.Close()

    '        If Not xmldoc.SelectSingleNode("//form") Is Nothing Then
    '            xmldoc.SelectSingleNode("//form").Attributes.GetNamedItem("uid").Value = xmldoc.SelectSingleNode("//form").Attributes.GetNamedItem("uid").Value & "_" & objMain.objApplication.Forms.Count
    '            'Dim a As String = xmldoc.InnerXml
    '            'objMain.objApplication.LoadBatchActions(xmldoc.InnerXml)
    '            Dim a As String = xmldoc.InnerXml
    '            a = ModifySize(a)
    '            objMain.objApplication.LoadBatchActions(a)

    '        End If
    '    Catch ex As Exception
    '        objMain.objApplication.MessageBox(ex.Message)
    '    End Try
    'End Sub
    Public Sub LoadForm(ByVal XMLFile As String, ByVal FormType As String, Optional ByVal FileType As ResourceType = ResourceType.Content)
        Try
            Dim AppAssemblty As Assembly = Assembly.GetExecutingAssembly()
            Dim sExecutingAssemblyNmae As String = AppAssemblty.GetName().Name.ToString()
            Dim xmldoc As New Xml.XmlDocument
            XMLFile = sExecutingAssemblyNmae + "." + XMLFile

            '// load the form to the TNX application in one batch
            Dim Streaming As System.IO.Stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(XMLFile)
            Dim StreamRead As New System.IO.StreamReader(Streaming, True)
            xmldoc.LoadXml(StreamRead.ReadToEnd)
            StreamRead.Close()

            If Not xmldoc.SelectSingleNode("//form") Is Nothing Then
                xmldoc.SelectSingleNode("//form").Attributes.GetNamedItem("uid").Value = xmldoc.SelectSingleNode("//form").Attributes.GetNamedItem("uid").Value & "_" & objMain.objApplication.Forms.Count
                Dim a As String = xmldoc.InnerXml
                objMain.objApplication.LoadBatchActions(xmldoc.InnerXml)
            End If
        Catch ex As Exception
            objMain.objApplication.MessageBox(ex.Message)
        End Try
    End Sub
#End Region

#Region "Add Column - ## Not Used "
    Private Sub addCol(ByVal strTab As String, ByVal strCol As String, ByVal strDesc As String, ByVal nType As Integer, Optional ByVal nEditSize As Integer = 10, Optional ByVal nSubType As Integer = 0)

        Dim oUFields As SAPbobsCOM.UserFieldsMD
        Dim nError As Integer

        oUFields = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        oUFields.TableName = strTab
        oUFields.Name = strCol
        oUFields.Type = nType
        oUFields.SubType = nSubType
        oUFields.Description = strDesc
        oUFields.EditSize = nEditSize
        nError = oUFields.Add()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields)
        GC.Collect()
        GC.WaitForPendingFinalizers()
        If nError <> 0 Then
            'MsgBox(strCol & " table could not be added")
        End If
    End Sub
#End Region

#Region "Create Table"
    Public Function CreateTable(ByVal TableName As String, ByVal TableDescription As String, ByVal TableType As SAPbobsCOM.BoUTBTableType) As Boolean
        Dim intRetCode As Integer
        Dim objUserTableMD As SAPbobsCOM.UserTablesMD
        objUserTableMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
        Try
            If (Not objUserTableMD.GetByKey(TableName)) Then
                objMain.objApplication.StatusBar.SetText("Creating table... [@" & TableName & "]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objUserTableMD.TableName = TableName
                objUserTableMD.TableDescription = TableDescription
                objUserTableMD.TableType = TableType
                intRetCode = objUserTableMD.Add()
                If (intRetCode = 0) Then
                    Return True
                Else
                    'Vj Added for testing///////////////
                    Dim lret As Integer
                    Dim sret As String = String.Empty
                    objMain.objCompany.GetLastError(lret, sret)
                    objMain.objApplication.MessageBox(lret & " : " & sret)
                    '//////////////////Done
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            objMain.objApplication.MessageBox(ex.Message)
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserTableMD)
            GC.Collect()
        End Try
    End Function

#End Region

    Public Function GetNextDocNumSeries(ByRef Objform As SAPbouiCOM.Form, ByVal UDOName As String, ByVal Series As String) As Integer
        Dim Str As String
        Dim oRs As SAPbobsCOM.Recordset
        Dim DocNum As Integer
        oRs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Str = "Select ""Series"" From NNM1 Where ""ObjectCode"" = '" & UDOName & "' And ""Series""= '" & Series & "'"
        Try
            oRs.DoQuery(Str)
            oRs.MoveFirst()
            If oRs.RecordCount > 0 Then
                DocNum = Objform.BusinessObject.GetNextSerialNumber(oRs.Fields.Item(0).Value, UDOName)
            End If
            If DocNum = 0 Then DocNum = 1
            Return DocNum
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("DN: " + ex.Message)
        End Try
    End Function


    'Public Function licence1(ByVal Addonname As String, ByVal DBName As String, ByVal User As String) As Boolean
    '    If objMain.objCompany.Connected Then
    '        Dim rsConfig1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        Dim query As String = "Select T0.""U_License"" as ""licence"", ""U_HKey"" as ""Hardwarekey"",T0.""U_Company"", T0.""U_DB"" As ""Dbname"",T0.""U_Addon"" as ""Addon"", T0.""U_Total"" as ""Total""," &
    '                          "T1.""U_Code"" as ""User"", T1.""U_Sts"" as ""Status"", " &
    '                          "T0.""U_SDate"" as ""fromdate"", T0.""U_EDate"" as ""Todate"" " &
    '                          "from ""@TNX_LICENCE"" T0 " &
    '                          "INNER JOIN ""@TNX_LICENCE_C0"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
    '                          "WHERE T1.""U_Sts"" = 'Y' AND T1.""U_Code"" ='" & User & "' AND T0.""U_Addon"" ='" & Addonname & "' " &
    '                          "AND T0.""U_DB"" ='" & DBName & "' and CURRENT_DATE BETWEEN T0.""U_SDate"" AND T0.""U_EDate"""
    '        rsConfig1.DoQuery(query)
    '        Dim licence As String = rsConfig1.Fields.Item("licence").Value
    '        Dim FDate As DateTime = rsConfig1.Fields.Item("fromdate").Value
    '        Dim toDate As DateTime = rsConfig1.Fields.Item("Todate").Value
    '        Dim currentDate As DateTime = DateTime.Now
    '        Dim user2 As String = rsConfig1.Fields.Item("User").Value
    '        Dim Addonname1 As String = rsConfig1.Fields.Item("Addon").Value
    '        Dim dbname1 As String = rsConfig1.Fields.Item("Dbname").Value
    '        If rsConfig1.RecordCount > 0 And licence <> "" Then
    '            Dim licenseStatus As String = rsConfig1.Fields.Item("Status").Value
    '            If licenseStatus = "Y" Then
    '                Dim License, Database, Company, HKey, TotalLicense As String
    '                Dim StartDate, ExpDate As Date
    '                License = rsConfig1.Fields.Item("licence").Value
    '                Company = rsConfig1.Fields.Item("U_Company").Value
    '                Database = rsConfig1.Fields.Item("Dbname").Value
    '                HKey = rsConfig1.Fields.Item("Hardwarekey").Value
    '                TotalLicense = rsConfig1.Fields.Item("Total").Value
    '                Dim SDate As String = CDate(rsConfig1.Fields.Item("fromdate").Value).ToString("yyyyMMdd")
    '                Dim EDate As String = CDate(rsConfig1.Fields.Item("Todate").Value).ToString("yyyyMMdd")
    '                SDate = SDate.Insert(4, "/")
    '                SDate = SDate.Insert(7, "/")
    '                EDate = EDate.Insert(4, "/")
    '                EDate = EDate.Insert(7, "/")
    '                'StartDate = CDate(SDate)
    '                'ExpDate = CDate(EDate)
    '                'Dim SDate As String = rsConfig1.Fields.Item("fromdate").Value
    '                'Dim EDate As String = rsConfig1.Fields.Item("Todate").Value
    '                ' Correct date formats
    '                StartDate = SDate
    '                ExpDate = EDate

    '                Dim uniKey As String = "10X_SAP_B!_"
    '                Dim strKey As String = StartDate.ToString("dd-MMM-yyyy") & "|" & ExpDate.ToString("dd-MMM-yyyy") & "|" & Company & "|" & TotalLicense & "|" & Database & "|" & HKey & "|" & Addonname
    '                Dim TestLicense As String = GetSHA1HashData(strKey + GetSHA1HashData(uniKey))
    '                Dim GenLicense As String = GetSHA1HashData(strKey + GetSHA1HashData(uniKey))
    '                ' Check if the license is valid
    '                If License <> GetSHA1HashData(strKey + GetSHA1HashData(uniKey)) Then
    '                    objMain.objApplication.StatusBar.SetText("Invalid license file", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                    Return False
    '                Else
    '                    '  Me.SBO_Application.StatusBar.SetText("Licence Tenx is valid", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    '                    Return True
    '                End If
    '            End If
    '        Else
    '            objMain.objApplication.StatusBar.SetText("License Tenx is invalid", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Return False
    '        End If
    '    End If
    'End Function
    'Private Function GetSHA1HashData(ByVal data As String) As String
    '    Dim sha2 As SHA1 = sha2.Create()
    '    Dim hashData As Byte() = sha2.ComputeHash(Encoding.[Default].GetBytes(data))
    '    Dim returnValue As System.Text.StringBuilder = New StringBuilder()
    '    For i As Integer = 0 To hashData.Length - 1
    '        returnValue.Append(hashData(i).ToString())
    '    Next
    '    Return returnValue.ToString()
    'End Function

    Private Function GetSHA1HashData(ByVal data As String) As String

        Dim sha2 As SHA1 = sha2.Create()
        Dim hashData As Byte() = sha2.ComputeHash(Encoding.[Default].GetBytes(data))
        Dim returnValue As System.Text.StringBuilder = New StringBuilder()
        For i As Integer = 0 To hashData.Length - 1
            returnValue.Append(hashData(i).ToString())
        Next
        Return returnValue.ToString()
    End Function


    Public Sub ComboBoxLoadValuesSeries(ByVal objCombo As SAPbouiCOM.ComboBox, ByVal QueryAsValueAndDescription As String)
        Try
            If (objCombo.ValidValues.Count <> 0) Then
                For R As Integer = objCombo.ValidValues.Count - 1 To 0 Step -1
                    Try
                        objCombo.ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                    Catch ex As Exception
                    End Try
                Next
            End If

            If objCombo.ValidValues.Count = 0 Then
                Dim objRecSet
                objRecSet = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecSet.DoQuery(QueryAsValueAndDescription)
                objRecSet.MoveFirst()
                'objCombo.ValidValues.Add("", "")
                While Not objRecSet.EoF
                    Try
                        objCombo.ValidValues.Add(objRecSet.Fields.Item(0).Value, objRecSet.Fields.Item(1).Value)
                    Catch ex As Exception
                    End Try
                    objRecSet.MoveNext()
                End While
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub ComboBoxLoadValues1(ByVal objCombo As SAPbouiCOM.ComboBox, ByVal QueryAsValueAndDescription As String)
        Try
            If (objCombo.ValidValues.Count <> 0) Then
                For R As Integer = objCombo.ValidValues.Count - 1 To 0 Step -1
                    Try
                        objCombo.ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                    Catch ex As Exception
                    End Try
                Next
            End If

            If objCombo.ValidValues.Count = 0 Then
                Dim objRecSet
                objRecSet = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecSet.DoQuery(QueryAsValueAndDescription)
                objRecSet.MoveFirst()
                While Not objRecSet.EoF
                    Try
                        objCombo.ValidValues.Add(objRecSet.Fields.Item(0).Value, objRecSet.Fields.Item(1).Value)
                    Catch ex As Exception
                    End Try
                    objRecSet.MoveNext()
                End While
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

#Region " LoadValidValues "
#Region " LoadValidValues "
    Public Sub AddValidValue(ByVal FormUID As String, ByVal FormType As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim GetDocNum As String = "Select ""DocNum"" , ""U_VSPMATID"" , ""U_VSPITCL"" From ""@VSP_DDCS""  Where ""U_VSPFRMID"" = '" & FormType & "'  And ""U_VSPACTV"" = 'Y'"
            Dim oRsGetDocNum As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDocNum.DoQuery(GetDocNum)

            If oRsGetDocNum.RecordCount > 0 Then
                oRsGetDocNum.MoveFirst()
                For i As Integer = 1 To oRsGetDocNum.RecordCount
                    Try
                        Dim GetDetails As String = "Select T1.""U_VSPVALUS"" , T1.""U_VSPDESC"" From ""@VSP_DDCS"" T0  Inner Join ""@VSP_DDCS_C0"" T1  On T0.""DocEntry"" = T1.""DocEntry"" " &
                        "Where ""DocNum"" = '" & oRsGetDocNum.Fields.Item(0).Value & "' And ""U_VSPVALUS"" <> '' And ""U_VSPACTV"" = 'Y' "
                        Dim oRsGetDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetDetails.DoQuery(GetDetails)

                        If oRsGetDetails.RecordCount > 0 Then
                            If oRsGetDocNum.Fields.Item(1).Value <> "" Then
                                oRsGetDetails.MoveFirst()

                                Dim objMatrix As SAPbouiCOM.Matrix
                                Dim oColumn As SAPbouiCOM.Column
                                objMatrix = objForm.Items.Item(oRsGetDocNum.Fields.Item("U_VSPMATID").Value).Specific
                                oColumn = objMatrix.Columns.Item(oRsGetDocNum.Fields.Item("U_VSPITCL").Value)

                                If (oColumn.ValidValues.Count <> 0) Then
                                    For R As Integer = oColumn.ValidValues.Count - 1 To 0 Step -1
                                        Try
                                            oColumn.ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                                        Catch ex As Exception
                                        End Try
                                    Next
                                End If

                                oColumn.ValidValues.Add("", "")
                                While Not oRsGetDetails.EoF
                                    Try
                                        oColumn.ValidValues.Add(oRsGetDetails.Fields.Item("U_VSPVALUS").Value, oRsGetDetails.Fields.Item("U_VSPDESC").Value)
                                    Catch ex As Exception
                                    End Try
                                    oRsGetDetails.MoveNext()
                                End While

                            Else
                                oRsGetDetails.MoveFirst()

                                Dim objCombo As SAPbouiCOM.ComboBox = objForm.Items.Item(oRsGetDocNum.Fields.Item("U_VSPITCL").Value).Specific

                                If (objCombo.ValidValues.Count <> 0) Then
                                    For R As Integer = objCombo.ValidValues.Count - 1 To 0 Step -1
                                        Try
                                            objCombo.ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                                        Catch ex As Exception
                                        End Try
                                    Next
                                End If

                                objCombo.ValidValues.Add("", "")
                                While Not oRsGetDetails.EoF
                                    Try
                                        objCombo.ValidValues.Add(oRsGetDetails.Fields.Item("U_VSPVALUS").Value, oRsGetDetails.Fields.Item("U_VSPDESC").Value)
                                    Catch ex As Exception
                                    End Try
                                    oRsGetDetails.MoveNext()
                                End While
                            End If
                        End If
                    Catch ex As Exception
                    End Try
                    oRsGetDocNum.MoveNext()
                Next
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region
#End Region

#Region "Fields Creation"
    Public Sub AddAlphaField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, Optional ByVal DefaultValue As String = "")
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, Size, SAPbobsCOM.BoFldSubTypes.st_None, "", "", DefaultValue)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Public Sub AddAlphaField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, Size, SAPbobsCOM.BoFldSubTypes.st_None, ValidValues, ValidDescriptions, SetValidValue)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub AddRow(pMatrix As SAPbouiCOM.Matrix, oDBDetail1 As SAPbouiCOM.DBDataSource)
        Try
            pMatrix.FlushToDataSource()
            If pMatrix.VisualRowCount > 0 Then
                oDBDetail1.InsertRecord(pMatrix.VisualRowCount)
                pMatrix.LoadFromDataSource()
            End If
            Dim i = 0, loopTo As Integer = oDBDetail1.Size

            While i <= loopTo
                oDBDetail1.SetValue("LineId", i, Convert.ToString(i + 1))
                i += 1
            End While
            pMatrix.LoadFromDataSource()
        Catch ex As Exception
            ' ' CreateErrorlog(ClsSBOAddOn.SBOompany.UserName.ToString(), "Namespace:TAG_Scheme, Class: ClsItemGroupMaster,Function:ReceiveData,Error:" & ex.Message, DateTime.Now, clsUIGeneric.strErrorLogFilePath)

        End Try
    End Sub
    Public Sub DeleteRow(pMatrix As SAPbouiCOM.Matrix, oDBDetail1 As SAPbouiCOM.DBDataSource)
        Try
            pMatrix.FlushToDataSource()
            oDBDetail1.RemoveRecord(oDBDetail1.Size - 1)
            Dim i = 0, loopTo As Integer = oDBDetail1.Size

            While i <= loopTo
                oDBDetail1.SetValue("LineId", i, Convert.ToString(i + 1))
                i += 1
            End While
            pMatrix.Clear()
            pMatrix.LoadFromDataSource()
        Catch ex As Exception
            ' clsUIGeneric.CreateErrorlog(ClsSBOAddOn.SBOompany.UserName.ToString(), "Namespace:TAG_Scheme, Class: ClsItemGroupMaster.cs,Function:ReceiveData,Error:" & ex.Message, DateTime.Now, clsUIGeneric.strErrorLogFilePath)

        End Try
    End Sub


    Public Sub addField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
        Dim intLoop As Integer
        Dim strValue, strDesc As Array
        Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
        objUserFieldMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        Try
            strValue = ValidValues.Split(Convert.ToChar(","))
            strDesc = ValidDescriptions.Split(Convert.ToChar(","))
            If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                Throw New Exception("Invalid Valid Values")
            End If
            If (Not isColumnExist(TableName, ColumnName)) Then
                objMain.objApplication.StatusBar.SetText("Creating field...[" & ColumnName & "] of table [" & TableName & "]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objUserFieldMD.TableName = TableName
                objUserFieldMD.Name = ColumnName
                objUserFieldMD.Description = ColDescription
                objUserFieldMD.Type = FieldType
                If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                    objUserFieldMD.Size = Size
                Else
                    objUserFieldMD.EditSize = Size
                End If
                objUserFieldMD.SubType = SubType
                objUserFieldMD.DefaultValue = SetValidValue
                If strValue.Length > 1 Then
                    For intLoop = 0 To strValue.GetLength(0) - 1
                        objUserFieldMD.ValidValues.Value = strValue(intLoop)
                        objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                        objUserFieldMD.ValidValues.Add()
                    Next
                End If
                If (objUserFieldMD.Add() <> 0) Then
                    objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription)
                End If
                'Else
                '    objMain.objApplication.StatusBar.SetText("Creating field...[" & ColumnName & "] of table [" & TableName & "]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                '    objUserFieldMD.TableName = TableName
                '    objUserFieldMD.Name = ColumnName
                '    objUserFieldMD.Description = ColDescription
                '    objUserFieldMD.Type = FieldType
                '    If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                '        objUserFieldMD.Size = Size
                '    Else
                '        objUserFieldMD.EditSize = Size
                '    End If
                '    objUserFieldMD.SubType = SubType
                '    objUserFieldMD.DefaultValue = SetValidValue
                '    For intLoop = 0 To strValue.GetLength(0) - 1
                '        objUserFieldMD.ValidValues.Value = strValue(intLoop)
                '        objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                '        objUserFieldMD.ValidValues.Add()
                '    Next
                '    If (objUserFieldMD.Update() <> 0) Then
                '        objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription)
                '    End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            GC.Collect()
        End Try
    End Sub

    Public Sub AddAlphaFieldLinkedTable(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, Optional ByVal DefaultValue As String = "", Optional ByVal sLinkedTable As String = "")
        Try
            addFieldLinkedTable(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, Size, SAPbobsCOM.BoFldSubTypes.st_None, "", "", DefaultValue, sLinkedTable)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub addFieldLinkedTable(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String, ByVal sLinkedTable As String)
        Dim intLoop As Integer
        Dim strValue, strDesc As Array
        Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
        objUserFieldMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        Try
            strValue = ValidValues.Split(Convert.ToChar(","))
            strDesc = ValidDescriptions.Split(Convert.ToChar(","))
            If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                Throw New Exception("Invalid Valid Values")
            End If

            If (Not isColumnExist(TableName, ColumnName)) Then
                objMain.objApplication.StatusBar.SetText("Creating field...[" & ColumnName & "] of table [" & TableName & "]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objUserFieldMD.TableName = TableName
                objUserFieldMD.Name = ColumnName
                objUserFieldMD.Description = ColDescription
                objUserFieldMD.Type = FieldType

                If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                    objUserFieldMD.Size = Size
                Else
                    objUserFieldMD.EditSize = Size
                End If
                objUserFieldMD.SubType = SubType
                objUserFieldMD.DefaultValue = SetValidValue
                If strValue.Length > 1 Then
                    For intLoop = 0 To strValue.GetLength(0) - 1
                        objUserFieldMD.ValidValues.Value = strValue(intLoop)
                        objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                        objUserFieldMD.ValidValues.Add()

                    Next
                End If

                objUserFieldMD.LinkedTable = sLinkedTable

                If (objUserFieldMD.Add() <> 0) Then
                    objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription)
                End If

            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message)

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            GC.Collect()
        End Try
    End Sub
    Public Sub addUDTLinkedField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal UDTLink As String, ByVal UDOLink As String, ByVal SysUDOLink As String)
        Dim intLoop As Integer
        Dim strValue, strDesc As Array
        Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
        objUserFieldMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        Try
            If (Not isColumnExist(TableName, ColumnName)) Then
                objMain.objApplication.StatusBar.SetText("Creating field...[" & ColumnName & "] of table [" & TableName & "]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objUserFieldMD.TableName = TableName
                objUserFieldMD.Name = ColumnName
                objUserFieldMD.Description = ColDescription
                objUserFieldMD.Type = FieldType

                If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                    objUserFieldMD.Size = Size
                Else
                    objUserFieldMD.EditSize = Size
                End If
                'objUserFieldMD.SubType = SubType
                If UDTLink <> "" Then
                    objUserFieldMD.LinkedTable = UDTLink
                End If
                If UDOLink <> "" Then
                    objUserFieldMD.LinkedUDO = UDOLink
                End If
                If SysUDOLink <> "" Then
                    objUserFieldMD.LinkedSystemObject = SysUDOLink
                End If
                If (objUserFieldMD.Add() <> 0) Then
                    objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            GC.Collect()
        End Try
    End Sub
    Private Function isColumnExist(ByVal TableName As String, ByVal ColumnName As String) As Boolean
        Dim objRecordSet As SAPbobsCOM.Recordset
        objRecordSet = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            objRecordSet.DoQuery("SELECT COUNT(*) FROM CUFD WHERE UPPER(""TableID"") = '" & TableName.Trim.ToUpper & "' AND UPPER(""AliasID"") = '" & ColumnName.Trim.ToUpper & "'")
            If (Convert.ToInt16(objRecordSet.Fields.Item(0).Value) <> 0) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecordSet)
            GC.Collect()
        End Try
    End Function

    Public Sub AddFloatField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal SubType As SAPbobsCOM.BoFldSubTypes)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Float, 0, SubType, "", "", "")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Public Sub AddDateField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal SubType As SAPbobsCOM.BoFldSubTypes)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Date, 0, SubType, "", "", "")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Public Sub AddAlphaMemoField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer)

        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Memo, Size, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Public Sub AddInteger(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal Size As Integer)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Numeric, Size, SubType, "", "", "")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Public Sub UniqueIDField(ByVal TableName As String, ByVal FieldName As String, ByVal oCompany As SAPbobsCOM.Company)

        '//****************************************************************************
        '// The UserKeysMD represents a meta-data object that allows you
        '// to add\remove user defined keys.
        '//****************************************************************************

        Dim oUserKeysMD As SAPbobsCOM.UserKeysMD

        '//flag
        Dim bFlagFirst As Boolean

        bFlagFirst = True

        '//****************************************************************************
        '// In any meta-data operation there should be no other object "alive"
        '// but the meta-data object, otherwise the operation will fail.
        '// This restriction is intended to prevent collisions.
        '//****************************************************************************

        '// The meta-data object must be initialized with a
        '// regular UserKeys object
        oUserKeysMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys)

        '// Set the table name and the key name
        'oUserKeysMD.TableName = "OCRD" '// BP table
        'oUserKeysMD.KeyName = "BE_MyKey1"

        oUserKeysMD.TableName = TableName '// BP table
        oUserKeysMD.KeyName = FieldName


        '//*******************************************
        '// Add a column to a key button:
        '//-------------------------------------------
        '// To add an additional column to
        '// the key, an additional element must be
        '// created in the Elements collection.
        '// The Add method of the Elements collection
        '// must be used only as of the second element.

        '// Do not use the Add method for the first element
        If bFlagFirst = True Then
            bFlagFirst = False
        Else
            '// Add an item to the Elements collection
            oUserKeysMD.Elements.Add()
            strLastErrorCode = oCompany.GetLastErrorCode()
            strLastError = oCompany.GetLastErrorDescription()
        End If

        '// Set the column's alias
        oUserKeysMD.Elements.ColumnAlias = FieldName

        '// Determine whether the key is unique or not
        'oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES
        oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tNO

        '// Add the key
        oUserKeysMD.Add()
        strLastErrorCode = oCompany.GetLastErrorCode()
        strLastError = oCompany.GetLastErrorDescription()
        'If (oUserKeysMD <> DBNull.Value) Then
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD)

        'End If
    End Sub
    Public Sub AddLinkField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Memo, Size, SubType, "", "", "")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub AddImageField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, 254, SAPbobsCOM.BoFldSubTypes.st_Image, "", "", "")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

#Region " Add Data to No Object Table"

    Public Function AddDataToNoObjectTable(ByVal TableName As String,
                                       ByVal Code As String,
                                       ByVal Name As String,
                                       Optional ByVal UDFName1 As String = "",
                                       Optional ByVal UDFValue1 As String = "") As String
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim lReturn As Integer
        Dim ErrorString As String = ""

        Try
            oUserTable = objMain.objCompany.UserTables.Item(TableName)

            'Check if record already exists
            If oUserTable.GetByKey(Code) Then

                'Update existing record
                oUserTable.Name = Name

                If UDFName1 <> "" Then
                    oUserTable.UserFields.Fields.Item(UDFName1).Value = UDFValue1
                End If

                lReturn = oUserTable.Update()

            Else

                'Add new record
                oUserTable.Code = Code
                oUserTable.Name = Name

                If UDFName1 <> "" Then
                    oUserTable.UserFields.Fields.Item(UDFName1).Value = UDFValue1
                End If

                lReturn = oUserTable.Add()

            End If

            'Check error
            If lReturn <> 0 Then
                objMain.objCompany.GetLastError(lReturn, ErrorString)
            End If

        Catch ex As Exception
            ErrorString = ex.Message
        Finally
            If Not oUserTable Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable)
                oUserTable = Nothing
            End If
        End Try

        Return ErrorString
    End Function
    Public Function AddDataToNoObjectTable1(ByVal TableName As String, ByVal Code As String, ByVal Name As String, Optional ByVal UDFName1 As String = "", Optional ByVal UDFValue1 As String = "")
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim lReturn As Integer
        Dim ErrorString As String
        oUserTable = objMain.objCompany.UserTables.Item(TableName)

        If oUserTable.GetByKey(Code) = True Then
            'Set default, mandatory fields
            oUserTable.Code = Code
            oUserTable.Name = Name

            'Set user field
            If UDFName1 <> String.Empty Then oUserTable.UserFields.Fields.Item(UDFName1).Value = UDFValue1
            oUserTable.Update()
            If lReturn <> 0 Then
                objMain.objCompany.GetLastError(lReturn, ErrorString)
                Return (ErrorString)
            End If
        Else

            oUserTable.Code = Code
            oUserTable.Name = Name

            'Set user field
            If UDFName1 <> String.Empty Then oUserTable.UserFields.Fields.Item(UDFName1).Value = UDFValue1
            oUserTable.Add()
            If lReturn <> 0 Then
                objMain.objCompany.GetLastError(lReturn, ErrorString)
                Return (ErrorString)
            End If

        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable)
        Return ("")
    End Function
#End Region

#Region "add menus with xml"
    Public Sub LoadFromXML(ByRef FileName As String)

        Dim oXmlDoc As Xml.XmlDocument
        oXmlDoc = New Xml.XmlDocument
        '// load the content of the XML File
        Dim sPath As String
        sPath = IO.Directory.GetParent(Application.ExecutablePath).ToString
        oXmlDoc.Load(sPath & "\" & FileName)
        '// load the form to the SBO application in one batch
        objMain.objApplication.LoadBatchActions(oXmlDoc.InnerXml)
        sPath = objMain.objApplication.GetLastBatchResults()

    End Sub
#End Region

#Region " Check if Form Exists - ## Not Used "
    Public Function FormExist(ByVal FormUID As String) As Boolean
        Dim intLoop As Integer

        For intLoop = objMain.objApplication.Forms.Count - 1 To 0 Step -1
            If Trim(FormUID) = Trim(objMain.objApplication.Forms.Item(intLoop).UniqueID) Then
                Return True
            End If
        Next
        Return False
    End Function
#End Region

#Region " Get MaxCode "
    Public Function getMaxCode(ByVal sTable As String) As String
        Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim MaxCode As Integer
        Dim sCode As String
        Dim strSQL As String
        Try
            strSQL = "SELECT MAX(CAST(""Code"" AS INT)) AS Code From """ & sTable & """"
            oRS.DoQuery(strSQL)
            If Convert.ToString(oRS.Fields.Item(0).Value).Length > 0 Then
                MaxCode = oRS.Fields.Item(0).Value + 1
            Else
                MaxCode = 10000001
            End If
            sCode = MaxCode
            Return sCode

        Catch ex As Exception
            Throw ex
        Finally
            oRS = Nothing
        End Try
    End Function
#End Region

#Region " UDO Document Numbering "
    Public Function GetNextDocNum(ByRef Objform As SAPbouiCOM.Form, ByVal UDOName As String, Optional ByVal SeriesName As String = "Primary") As Integer
        Dim Str As String
        Dim oRs As SAPbobsCOM.Recordset
        Dim DocNum As Integer
        oRs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Str = "Select ""Series"" From NNM1 Where ""ObjectCode"" = '" & UDOName & "' and ""SeriesName"" = '" & SeriesName & "'"
        Try
            oRs.DoQuery(Str)
            oRs.MoveFirst()
            If oRs.RecordCount > 0 Then
                DocNum = Objform.BusinessObject.GetNextSerialNumber(oRs.Fields.Item(0).Value, UDOName)
            End If
            If DocNum = 0 Then DocNum = 1
            Return DocNum
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("DN: " + ex.Message)
        End Try
    End Function
#End Region

    Public Function GetDocumentEntry(ByVal sDocNum As String, ByVal sDocSeries As String, ByVal DocTable As String)
        Try

            Dim GetDocEntry As String = "Select ""DocEntry"" From """ & DocTable & """ Where ""DocNum"" = '" & sDocNum & "' And ""Series"" = '" & sDocSeries & "' "
            Dim oRsGetDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDocEntry.DoQuery(GetDocEntry)

            Return oRsGetDocEntry.Fields.Item("DocEntry").Value.ToString
        Catch ex As Exception
        End Try
    End Function

    Public Function GetUserName(ByVal UserCode As String)
        Try

            Dim UserName As String = "Select ""U_NAME"" From OUSR Where ""USER_CODE"" = '" & UserCode & "' "
            Dim oRsUserName As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsUserName.DoQuery(UserName)

            Return oRsUserName.Fields.Item("U_Name").Value.ToString
        Catch ex As Exception
        End Try
    End Function

#Region " Load DataSource from DB "
    Public Sub RefreshDatasourceFromDB(ByVal FormUID As String, ByRef oDBs_Head As SAPbouiCOM.DBDataSource, ByVal ConditionAlias As String, ByVal ConditionValue As String)
        Try
            Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUID)
            Dim oConditions As SAPbouiCOM.Conditions = New SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            oCondition = oConditions.Add()
            oCondition.Alias = ConditionAlias
            oCondition.ComparedAlias = ConditionAlias
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = ConditionValue
            oDBs_Head.Query(oConditions)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

#Region "      ***Load Values to ComboBoxes***       "
    Public Sub ComboBoxLoadValues(ByVal objCombo As SAPbouiCOM.ComboBox, ByVal QueryAsValueAndDescription As String)
        Try
            If (objCombo.ValidValues.Count <> 0) Then
                For R As Integer = objCombo.ValidValues.Count - 1 To 0 Step -1
                    Try
                        objCombo.ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                    Catch ex As Exception
                    End Try
                Next
            End If

            If objCombo.ValidValues.Count = 0 Then
                Dim objRecSet
                objRecSet = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecSet.DoQuery(QueryAsValueAndDescription)
                objRecSet.MoveFirst()
                objCombo.ValidValues.Add("", "")
                While Not objRecSet.EoF
                    Try
                        objCombo.ValidValues.Add(objRecSet.Fields.Item(0).Value, objRecSet.Fields.Item(1).Value)
                    Catch ex As Exception
                    End Try
                    objRecSet.MoveNext()
                End While
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub ComboBoxLoadValuesNEW(ByVal objCombo As SAPbouiCOM.ComboBox,
                              ByVal QueryAsValueAndDescription As String)

        Try

            Dim objRecSet As SAPbobsCOM.Recordset =
            objMain.objCompany.GetBusinessObject(
                SAPbobsCOM.BoObjectTypes.BoRecordset)

            objRecSet.DoQuery(QueryAsValueAndDescription)

            While Not objRecSet.EoF

                Dim Code As String =
                objRecSet.Fields.Item(0).Value.ToString().Trim()

                Dim Description As String =
                objRecSet.Fields.Item(1).Value.ToString().Trim()

                If Code <> "" Then

                    Dim Exists As Boolean = False

                    For i As Integer = 0 To objCombo.ValidValues.Count - 1

                        If objCombo.ValidValues.Item(i).Value.ToString().Trim() = Code Then
                            Exists = True
                            Exit For
                        End If

                    Next

                    If Not Exists Then
                        objCombo.ValidValues.Add(Code, Description)
                    End If

                End If

                objRecSet.MoveNext()

            End While

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
            "ComboBoxLoadValues Error: " & ex.Message
        )

        End Try

    End Sub

    Public Sub MatrixComboBoxValues(ByVal oColumn As SAPbouiCOM.Column, ByVal QueryAsValueAndDescription As String)
        Try
            If oColumn.ValidValues.Count = 0 Then
                Dim objRecSet
                objRecSet = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecSet.DoQuery(QueryAsValueAndDescription)
                objRecSet.MoveFirst()
                oColumn.ValidValues.Add("", "")
                While Not objRecSet.EoF
                    Try
                        oColumn.ValidValues.Add(objRecSet.Fields.Item(0).Value, objRecSet.Fields.Item(1).Value)
                    Catch ex As Exception
                    End Try
                    objRecSet.MoveNext()
                End While
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Sub MatrixComboBoxValues1(ByVal oColumn As SAPbouiCOM.Column, ByVal QueryAsValueAndDescription As String)
        Try
            Dim objRecSet
            objRecSet = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRecSet.DoQuery(QueryAsValueAndDescription)
            objRecSet.MoveFirst()
            oColumn.ValidValues.Add("", "")
            While Not objRecSet.EoF
                Try
                    oColumn.ValidValues.Add(objRecSet.Fields.Item(0).Value, objRecSet.Fields.Item(1).Value)
                Catch ex As Exception
                End Try
                objRecSet.MoveNext()
            End While
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

    Function SetAttachMentFile()
        Dim ShowFolderBrowserThread As Threading.Thread
        Try
            'Me.ShowFolderBrowser()
            ShowFolderBrowserThread = New Threading.Thread(AddressOf ShowFolderBrowser)
            If ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Unstarted Then
                ShowFolderBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA)
                ShowFolderBrowserThread.Start()
                'ElseIf ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Stopped Then
                '    ShowFolderBrowserThread.Start()
                '    ShowFolderBrowserThread.Join()
            End If
            While ShowFolderBrowserThread.ThreadState = Threading.ThreadState.Running
                System.Windows.Forms.Application.DoEvents()
            End While
            If BankFileName <> "" Then
                Return BankFileName
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'oApplication.MessageBox("FileFile Method Failed : " & ex.Message)
        End Try
        Return ""
    End Function

    Public Sub ShowFolderBrowser()
        Dim MyProcs() As System.Diagnostics.Process
        BankFileName = ""
        Dim OpenFile As New OpenFileDialog
        Try
            OpenFile.Multiselect = False
            OpenFile.Filter = "All files(*.)|*.*" '   "|*.*"
            Dim filterindex As Integer = 0
            Try
                filterindex = 0
            Catch ex As Exception
            End Try
            OpenFile.FilterIndex = filterindex
            OpenFile.RestoreDirectory = True
            MyProcs = Process.GetProcessesByName("SAP Business One")
            'If MyProcs.Length = 2 Then
            'For i As Integer = 0 To MyProcs.Length - 2
            Dim MyWindow As New clsWindowWrapper(MyProcs(0).MainWindowHandle)
            Dim ret As DialogResult = OpenFile.ShowDialog(MyWindow)
            If ret = DialogResult.OK Then
                BankFileName = OpenFile.FileName
                OpenFile.Dispose()
            Else
                System.Windows.Forms.Application.ExitThread()
            End If
            'Next
            'End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            BankFileName = ""
        Finally
            OpenFile.Dispose()
        End Try
    End Sub

#Region " Checking DataType"
    Public Function IsAlpha(ByVal str As String) As Boolean
        Try
            Dim i As Integer
            For i = 0 To str.Length - 1
                If Not Char.IsLetter(str, i) Then
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception

        End Try
    End Function

    Public Function IsNumeric(ByVal str As String) As Boolean
        Try
            Dim i As Integer
            If str.Contains(".") = True Then
                Return False
            End If
            For i = 0 To str.Length - 1
                If Not Char.IsNumber(str, i) Then
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception

        End Try
    End Function

    Public Function IsFloat(ByVal str As String) As Boolean
        Try
            Dim i As Integer
            If str.Substring(0, 1) = "." Then
                Return False
            End If
            If str.Contains(".") = False Then
                Return False
            Else
                str = str.Remove(str.LastIndexOfAny("."), 1)
            End If
            For i = 0 To str.Length - 1
                If Not Char.IsNumber(str, i) Then
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception

        End Try
    End Function

    Function IsPercentage(ByVal str As String) As Boolean
        Try
            Dim i As Integer
            If str.Contains(".") = True Then
                str = str.Remove(str.LastIndexOfAny("."), 1)
                If str.Contains(".") = True Then
                    Return False
                End If
            End If
            For i = 0 To str.Length - 1
                If Not Char.IsNumber(str, i) Then
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception

        End Try
    End Function
#End Region

#Region " Adding Items To Forms"

    Public Sub AddLabel(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer,
                        ByVal iWidth As Integer, ByVal iCaption As String, ByVal iLink As String, Optional ByVal iHeight As Integer = 14,
                        Optional ByVal iFromPane As Integer = 0, Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.LinkTo = iLink
        oItem.Specific.Caption = iCaption

    End Sub


    Public Sub AddEditBox(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer,
                          ByVal iWidth As Integer, ByVal TableName As String, ByVal UdFName As String,
                          ByVal LinkTo As String, Optional ByVal iHeight As Integer = 14, Optional ByVal iFromPane As Integer = 0,
                          Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.Height = iHeight
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.LinkTo = LinkTo
        oItem.Height = iHeight
        oItem.Specific.DataBind.SetBound(True, TableName, UdFName)
    End Sub

    Public Sub AddExtendedEditBox(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer,
                                  ByVal iLeft As Integer, ByVal iWidth As Integer, ByVal TableName As String,
                                  ByVal UdFName As String, ByVal LinkTo As String, Optional ByVal iFromPane As Integer = 0,
                                  Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.Height = 80
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.LinkTo = LinkTo
        oItem.Specific.DataBind.SetBound(True, TableName, UdFName)
    End Sub
    Public Sub AddComboBox(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer,
                             ByVal iWidth As Integer, ByVal iHight As Integer, ByVal TableName As String, ByVal UdFName As String, Optional ByVal iFromPane As Integer = 0,
                             Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)

        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.Height = iHight
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.Specific.DataBind.SetBound(True, TableName, UdFName)
    End Sub

    Public Sub AddComboBox2(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer,
                       ByVal iWidth As Integer, ByVal iHight As Integer, ByVal TableName As String, ByVal UdFName As String,
                       Optional ByVal iFromPane As Integer = 0, Optional ByVal iToPane As Integer = 0)

        Dim objForm As SAPbouiCOM.Form
        Dim oItem As SAPbouiCOM.Item
        Dim oCombo As SAPbouiCOM.ComboBox

        objForm = objMain.objApplication.Forms.Item(FormUID)

        '❗ Always remove if exists (to avoid conflict)
        If objForm.Items.Exists(ItemUID) Then
            objForm.Items.Remove(ItemUID)
        End If

        '👉 Create fresh ComboBox
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)

        '👉 Set position
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.Height = iHight
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane

        '👉 Bind
        oCombo = CType(oItem.Specific, SAPbouiCOM.ComboBox)
        oCombo.DataBind.SetBound(True, TableName, UdFName)

    End Sub
    Public Sub AddComboBox1(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer,
                           ByVal iWidth As Integer, ByVal TableName As String, ByVal UdFName As String, ByVal Link As String, Optional ByVal iFromPane As Integer = 0,
                           Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.LinkTo = Link
        oItem.Specific.DataBind.SetBound(True, TableName, UdFName)
    End Sub


    Public Sub AddFolder(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer, ByVal iWidth As Integer,
                         ByVal UdFName As String, ByVal Caption As String, ByVal AliasName As String,
                         ByVal GroupItem As String)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_FOLDER)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        Dim oFolder As SAPbouiCOM.Folder
        oFolder = oItem.Specific
        oFolder.Caption = Caption
        'oFolder.DataBind.SetBound(True, "", AliasName)
        oFolder.GroupWith(GroupItem)
    End Sub

    Public Sub AddButtonH(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer, ByVal iWidth As Integer,
                         ByVal LinkTo As String, ByVal Caption As String, Optional ByVal iFromPane As Integer = 0,
                                  Optional ByVal iToPane As Integer = 0, Optional ByVal iHeight As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_BUTTON)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.Height = iHeight
        oItem.LinkTo = LinkTo
        Dim btn As SAPbouiCOM.Button = objForm.Items.Item(ItemUID).Specific
        btn.Caption = Caption
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
    End Sub

    Public Sub AddButton(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer, ByVal iWidth As Integer,
                          ByVal iHight As Integer, ByVal LinkTo As String, ByVal Caption As String, Optional ByVal iFromPane As Integer = 0,
                                  Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_BUTTON)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.Height = iHight
        oItem.LinkTo = LinkTo
        Dim btn As SAPbouiCOM.Button = objForm.Items.Item(ItemUID).Specific
        btn.Caption = Caption
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
    End Sub

    Public Sub AddLinkButton(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer,
                             ByVal iLeft As Integer, ByVal iLinkTo As String, ByVal LinkedObject As Integer)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
        oItem.Top = iTop
        oItem.Left = iLeft
        Dim LinkBtn As SAPbouiCOM.LinkedButton
        LinkBtn = oItem.Specific
        oItem.LinkTo = iLinkTo
        If LinkedObject <> 0 Then
            LinkBtn.LinkedObject = LinkedObject
        End If
    End Sub

    Public Sub AddCheckBox(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer,
                           ByVal iWidth As Integer, ByVal TableName As String, ByVal UdFName As String, ByVal iCaption As String, Optional ByVal sLinkTo As String = "", Optional ByVal iHeight As Integer = 14, Optional ByVal iFromPane As Integer = 0,
                                  Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.Height = iHeight
        oItem.Specific.DataBind.SetBound(True, TableName, UdFName)
        oItem.Specific.Caption = iCaption
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.LinkTo = sLinkTo
    End Sub
#End Region

#Region "Conversions"
    Function RupeesToWord(ByVal MyNumber As String) As String
        Dim Temp As String
        Dim Rupees As String = String.Empty
        Dim Paisa As String = String.Empty
        Dim DecimalPlace As String = String.Empty
        Dim iCount As String = String.Empty
        Dim Hundred As String = String.Empty
        Dim Words As String = String.Empty

        Dim ValidateNumber As String = MyNumber

        Dim place(9) As String
        place(0) = " Thousand "
        place(2) = " Lakh "
        place(4) = " Crore "
        place(6) = " Hundred "
        place(8) = " Kharab "
        If ValidateNumber.Length > 9 Then
            If ValidateNumber.Length = 10 And ValidateNumber.Substring(1, ValidateNumber.Length - 1) = "0" Then
                place(4) = " Crore "
                place(6) = " Hundred Crore "
            ElseIf ValidateNumber.Length = 11 And (ValidateNumber.Substring(2, ValidateNumber.Length - 2) = "0") Then
                place(4) = " Crore "
                place(6) = " Hundred "
            Else
                place(4) = " Crore "
                place(6) = " Hundred "
            End If
        End If

        On Error Resume Next
        ' Convert MyNumber to a string, trimming extra spaces.
        MyNumber = Trim(Str(MyNumber))

        ' Find decimal place.
        DecimalPlace = InStr(MyNumber, ".")

        ' If we find decimal place...
        If DecimalPlace > 0 Then
            ' Convert Paisa
            Temp = Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2)
            Paisa = " and " & ConvertTens(Temp) & " Paisa"

            ' Strip off paisa from remainder to convert.
            MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
        End If

        '===============================================================
        Dim TM As String  ' If MyNumber between Rs.1 To 99 Only.
        TM = Right(MyNumber, 2)

        If Len(MyNumber) > 0 And Len(MyNumber) <= 2 Then
            If Len(TM) = 1 Then
                Words = ConvertDigit(TM)
                RupeesToWord = "Rupees " & Words & Paisa & " Only"

                Exit Function

            Else
                If Len(TM) = 2 Then
                    Words = ConvertTens(TM)
                    RupeesToWord = "Rupees " & Words & Paisa & " Only"
                    Exit Function

                End If
            End If
        End If
        '===============================================================


        ' Convert last 3 digits of MyNumber to ruppees in word.
        Hundred = ConvertHundreds(Right(MyNumber, 3))
        ' Strip off last three digits
        MyNumber = Left(MyNumber, Len(MyNumber) - 3)

        iCount = 0
        Do While MyNumber <> ""
            'Strip last two digits
            Temp = Right(MyNumber, 2)
            If Len(MyNumber) = 1 Then


                If Trim(Words) = "Thousand" Or
                Trim(Words) = "Lakh  Thousand" Or
                Trim(Words) = "Lakh" Or
                Trim(Words) = "Crore" Or
                Trim(Words) = "Crore  Lakh  Thousand" Or
                Trim(Words) = "Hundred  Crore  Lakh  Thousand" Or
                Trim(Words) = "Hundred" Or
                Trim(Words) = "Kharab  Hundred  Crore  Lakh  Thousand" Or
                Trim(Words) = "Kharab" Then

                    Words = ConvertDigit(Temp) & place(iCount)
                    MyNumber = Left(MyNumber, Len(MyNumber) - 1)

                Else

                    Words = ConvertDigit(Temp) & place(iCount) & Words
                    MyNumber = Left(MyNumber, Len(MyNumber) - 1)

                End If
            Else

                If Trim(Words) = "Thousand" Or
                   Trim(Words) = "Lakh  Thousand" Or
                   Trim(Words) = "Lakh" Or
                   Trim(Words) = "Crore" Or
                   Trim(Words) = "Crore  Lakh  Thousand" Or
                   Trim(Words) = "Hundred  Crore  Lakh  Thousand" Or
                   Trim(Words) = "Hundred" Then


                    Words = ConvertTens(Temp) & place(iCount)


                    MyNumber = Left(MyNumber, Len(MyNumber) - 2)
                Else

                    '=================================================================
                    ' if only Lakh, Crore, Arab, Kharab

                    If Trim(ConvertTens(Temp) & place(iCount)) = "Lakh" Or
                       Trim(ConvertTens(Temp) & place(iCount)) = "Crore" Or
                       Trim(ConvertTens(Temp) & place(iCount)) = "Hundred" Then

                        Words = Words
                        MyNumber = Left(MyNumber, Len(MyNumber) - 2)
                    Else
                        Words = ConvertTens(Temp) & place(iCount) & Words
                        MyNumber = Left(MyNumber, Len(MyNumber) - 2)
                    End If

                End If
            End If

            iCount = iCount + 2
        Loop

        RupeesToWord = "Rupees " & Words & Hundred & Paisa & " Only"

    End Function

    Private Function ConvertHundreds(ByVal MyNumber As String) As String
        Dim Result As String = String.Empty
        'Return String.Empty
        ' Exit if there is nothing to convert.
        If Val(MyNumber) = 0 Then
            Return Nothing
            'Exit Function
        End If



        ' Append leading zeros to number.
        MyNumber = Right("000" & MyNumber, 3)

        ' Do we have a hundreds place digit to convert?
        If Left(MyNumber, 1) <> "0" Then
            Result = ConvertDigit(Left(MyNumber, 1)) & " Hundred "
        End If

        ' Do we have a tens place digit to convert?
        If Mid(MyNumber, 2, 1) <> "0" Then
            Result = Result & ConvertTens(Mid(MyNumber, 2))
        Else
            ' If not, then convert the ones place digit.
            Result = Result & ConvertDigit(Mid(MyNumber, 3))
        End If

        ConvertHundreds = Trim(Result)
    End Function

    Private Function ConvertTens(ByVal MyTens As String) As String
        Dim Result As String = String.Empty

        ' Is value between 10 and 19?
        If Val(Left(MyTens, 1)) = 1 Then
            Select Case Val(MyTens)
                Case 10 : Result = "Ten"
                Case 11 : Result = "Eleven"
                Case 12 : Result = "Twelve"
                Case 13 : Result = "Thirteen"
                Case 14 : Result = "Fourteen"
                Case 15 : Result = "Fifteen"
                Case 16 : Result = "Sixteen"
                Case 17 : Result = "Seventeen"
                Case 18 : Result = "Eighteen"
                Case 19 : Result = "Nineteen"
                Case Else
            End Select
        Else
            ' .. otherwise it's between 20 and 99.
            Select Case Val(Left(MyTens, 1))
                Case 2 : Result = "Twenty "
                Case 3 : Result = "Thirty "
                Case 4 : Result = "Forty "
                Case 5 : Result = "Fifty "
                Case 6 : Result = "Sixty "
                Case 7 : Result = "Seventy "
                Case 8 : Result = "Eighty "
                Case 9 : Result = "Ninety "
                Case Else
            End Select

            ' Convert ones place digit.
            Result = Result & ConvertDigit(Right(MyTens, 1))
        End If

        ConvertTens = Result
    End Function

    Private Function ConvertDigit(ByVal MyDigit As String) As String
        Select Case Val(MyDigit)
            Case 1 : ConvertDigit = "One"
            Case 2 : ConvertDigit = "Two"
            Case 3 : ConvertDigit = "Three"
            Case 4 : ConvertDigit = "Four"
            Case 5 : ConvertDigit = "Five"
            Case 6 : ConvertDigit = "Six"
            Case 7 : ConvertDigit = "Seven"
            Case 8 : ConvertDigit = "Eight"
            Case 9 : ConvertDigit = "Nine"
            Case Else : ConvertDigit = ""
        End Select
    End Function
#End Region

    Public Sub AddLinkTableField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, ByVal LinkedTableName As String)
        Try

            Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
            objUserFieldMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            Try

                If (Not isColumnExist(TableName, ColumnName)) Then
                    objMain.objApplication.StatusBar.SetText("Creating field...[" & ColumnName & "] of table [" & TableName & "]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objUserFieldMD.TableName = TableName
                    objUserFieldMD.Name = ColumnName
                    objUserFieldMD.Description = ColDescription
                    objUserFieldMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
                    objUserFieldMD.LinkedTable = LinkedTableName

                    If (objUserFieldMD.Add() <> 0) Then
                        objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription)
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
                GC.Collect()
            End Try
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub AddElements(ByVal FormUID As String, ByVal TableName As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            oItem = objForm.Items.Add("lbl", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 5
            oItem.Left = objForm.Items.Item("38").Left
            oItem.Width = objForm.Items.Item("3").Width
            oItem.Height = 15
            Dim lbl As SAPbouiCOM.StaticText = objForm.Items.Item("lbl").Specific
            lbl.Caption = "Total Quantity"

            oItem = objForm.Items.Add("TtlQty", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 5
            oItem.Left = objForm.Items.Item("38").Left + objForm.Items.Item("lbl").Width
            oItem.Width = objForm.Items.Item("3").Width
            oItem.Height = 15

            oItem = objForm.Items.Add("Src", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 5
            oItem.Left = objForm.Items.Item("38").Left + objForm.Items.Item("TtlQty").Width + objForm.Items.Item("lbl").Width + 4
            oItem.Width = 16
            oItem.Height = 15

            Dim btn As SAPbouiCOM.Button = objForm.Items.Item("Src").Specific
            btn.Caption = "..."

            Dim oEditText As SAPbouiCOM.EditText = objForm.Items.Item("TtlQty").Specific
            oEditText.DataBind.SetBound(True, TableName, "U_QUANTITY")
            objForm.Items.Item("TtlQty").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("TtlQty").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.DataSources.UserDataSources.Add("U_DflWhs", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            objForm.DataSources.UserDataSources.Add("U_Branch", SAPbouiCOM.BoDataType.dt_LONG_TEXT)

            oItem = objForm.Items.Add("lblDflWhs", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 22
            oItem.Left = objForm.Items.Item("38").Left
            oItem.Width = objForm.Items.Item("3").Width
            oItem.Height = 15
            Dim lbl1 As SAPbouiCOM.StaticText = objForm.Items.Item("lblDflWhs").Specific
            lbl1.Caption = "Warehouse"

            oItem = objForm.Items.Add("edtBranch", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oItem.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 22
            oItem.Left = objForm.Items.Item("38").Left + objForm.Items.Item("lblDflWhs").Width
            oItem.Width = objForm.Items.Item("4").Width
            oItem.Height = 15
            Dim oComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("edtBranch").Specific
            oComboBox.DataBind.SetBound(True, "", "U_Branch")
            objMain.objUtilities.ComboBoxLoadValues(oComboBox, "Select ""BPLId"" ,  ""BPLName"" From OBPL")
            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            oItem.DisplayDesc = True

            oItem = objForm.Items.Add("edtDflWhs", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oItem.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 22
            oItem.Left = objForm.Items.Item("edtBranch").Left + objForm.Items.Item("edtBranch").Width + 5
            oItem.Width = objForm.Items.Item("4").Width
            oItem.Height = 15
            oComboBox = objForm.Items.Item("edtDflWhs").Specific
            oComboBox.DataBind.SetBound(True, "", "U_DflWhs")
            objMain.objUtilities.ComboBoxLoadValues(oComboBox, "Select ""WhsCode"" , ""WhsName"" From ""OWHS"" ")
            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            oItem.DisplayDesc = True

            oItem = objForm.Items.Add("Src1", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 22
            oItem.Left = objForm.Items.Item("edtDflWhs").Left + objForm.Items.Item("edtDflWhs").Width + 2
            oItem.Width = 16
            oItem.LinkTo = "edtDflWhs"
            oItem.Height = 15
            Dim btn1 As SAPbouiCOM.Button = objForm.Items.Item("Src1").Specific
            btn1.Caption = "..."
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Function GetBranch(ByVal WhsCode As String) As String
        Try
            Dim BranchDtl As String = "Select ""BPLId"" , ""BPLName"" ,  ""BPLFrName"" , ""DfltResWhs"" From OBPL Where ""DfltResWhs"" = '" & WhsCode & "' "
            Dim oRsBranchDtl As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsBranchDtl.DoQuery(BranchDtl)
            If oRsBranchDtl.RecordCount > 0 Then
                Return oRsBranchDtl.Fields.Item("BPLFrName").Value.ToString
            Else
                Return Nothing
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return Nothing
        End Try
    End Function

    Public Function GetPCPerItem(ByVal ItemCode As String) As String
        Try
            Dim GetPCPrItem As String = "Select ""U_PIC"" From OITM Where ""ItemCode"" = '" & ItemCode & "' And ""U_PIC"" > 1 "
            Dim oRsGetPCPrItem As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetPCPrItem.DoQuery(GetPCPrItem)
            If oRsGetPCPrItem.RecordCount > 0 Then
                Return oRsGetPCPrItem.Fields.Item("U_PIC").Value.ToString
            Else
                Return ""
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return Nothing
        End Try
    End Function

    Public Function GetItemBaseUom(ByVal ItemCode As String) As String
        Try
            Dim GetPCPrItem As String = "Select (Select ""UomCode"" From OUOM Where ""UomEntry"" = OITM.""IUoMEntry"") As ""BaseUomCode"" From OITM Where ""ItemCode"" = '" & ItemCode & "' "
            Dim oRsGetPCPrItem As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetPCPrItem.DoQuery(GetPCPrItem)
            If oRsGetPCPrItem.RecordCount > 0 Then
                Return oRsGetPCPrItem.Fields.Item("BaseUomCode").Value.ToString
            Else
                Return ""
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return Nothing
        End Try
    End Function

    Public Function GetWhsBinLoc(ByVal WhsCode As String) As String
        Try
            Dim GetWhBinLoc As String = "Select ""U_STGBLOC"" From OWHS Where ""WhsCode"" = '" & WhsCode & "' And ""U_STGBLOC""<>'' "
            Dim oRsGetWhBinLoc As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetWhBinLoc.DoQuery(GetWhBinLoc)
            If oRsGetWhBinLoc.RecordCount > 0 Then
                Return oRsGetWhBinLoc.Fields.Item("U_STGBLOC").Value.ToString
            Else
                Return ""
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return Nothing
        End Try
    End Function

    Public Function GetBinID(ByVal BPLId As String) As String
        Try
            Dim GetBnId As String = "Select ""BPLId"" ,  ""BPLName"" From OBPL Where ""BPLName"" = '" & BPLId & "' "
            Dim oRsGetBnId As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetBnId.DoQuery(GetBnId)
            If oRsGetBnId.RecordCount > 0 Then
                Return oRsGetBnId.Fields.Item("BPLId").Value.ToString
            Else
                Return ""
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return Nothing
        End Try
    End Function

    Sub CreateMySimpleForm1()
        Try
            'Dim sIssueSupplier As String = oDBDSHeader.GetValue("CardCode", 0).Trim
            'If sIssueSupplier.Equals("") Then
            '    oCBF.StatusBarErrorMsg("Supplier Code should be left empty ....")
            '    Exit Try
            'End If
            'Dim sIssueProject As String = oDBDSHeader.GetValue("Project", 0).Trim
            'If sIssueProject.Equals("") Then
            '    oCBF.StatusBarErrorMsg("Project Code should be left empty ....")
            '    Exit Try
            'End If
            Dim CP As SAPbouiCOM.FormCreationParams = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)

            Dim oItem As SAPbouiCOM.Item
            Dim oGrid As SAPbouiCOM.Grid
            Dim oButton As SAPbouiCOM.Button
            Dim ocheckbox As SAPbouiCOM.CheckBox

            CP.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            CP.UniqueID = "SOPOMRDetails"
            CP.FormType = "200"

            objForm = objMain.objApplication.Forms.AddEx(CP)
            ' Set form width and height 
            objForm.Height = 300
            objForm.Width = 600
            objForm.Title = "PO ItemList"
            '' Add a Grid item to the form 
            oItem = objForm.Items.Add("MyGrid", SAPbouiCOM.BoFormItemTypes.it_GRID)
            ' Set the grid dimentions and position 
            oItem.Left = 20
            oItem.Top = 20
            oItem.Width = 550
            oItem.Height = 200
            ' Set the grid data 
            oGrid = oItem.Specific
            objForm.DataSources.DataTables.Add("MyDataTable")
            ' Add OK Button 
            oItem = objForm.Items.Add("ok", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 20
            oItem.Top = 230
            oItem.Width = 65
            oItem.Height = 20
            oButton = oItem.Specific
            oButton.Caption = "Ok"
            ' Add CANCEL Button 
            oItem = objForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 90
            oItem.Top = 230
            oItem.Width = 65
            oItem.Height = 20
            oButton = oItem.Specific
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto

            objForm.DataSources.UserDataSources.Add("Accept", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oItem = objForm.Items.Add("c_select", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = 20
            oItem.Top = 5
            oItem.Width = 120
            oItem.Height = 15
            oItem.Enabled = True
            oItem.Visible = True
            oItem.DisplayDesc = False
            oItem.AffectsFormMode = False
            ocheckbox = oItem.Specific
            ocheckbox.Caption = "Select All"
            ocheckbox.ValOn = "Y"
            ocheckbox.ValOff = "N"
            ocheckbox.DataBind.SetBound(True, "", "Accept")
            objForm.Items.Item("c_select").Visible = True

            Dim sQuery As String = " Select 'N' Selected, a.Code,c.ItemName,a.Quantity,b.Project from ITT1 a inner join OITT B on a.Father= b.Code " &
                    " inner join OITM c on a.Code=c.ItemCode where a.U_MakTyp='P' and B.U_Stages='5' and " &
                    " b.Project='" & objForm.Items.Item("157").Specific.value & "' and a.code not in " &
                    " (Select a.ItemCode from POR1 a inner join OPOR b on a.Docentry=b.Docentry " &
                    " where b.Project='" & objForm.Items.Item("157").Specific.value & "' and b.Canceled='N' And b.DocType <>'S') "
            objForm.DataSources.DataTables.Item(0).ExecuteQuery(sQuery)

            oGrid.DataTable = objForm.DataSources.DataTables.Item("MyDataTable")
            oGrid.AutoResizeColumns()
            oGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(2).Editable = False
            oGrid.Columns.Item(3).Editable = True
            oGrid.Columns.Item(4).Editable = False
            objForm.Visible = True
        Catch ex As Exception
            objForm.Close()
            objMain.objApplication.StatusBar.SetText("Item Details Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    'Public Sub Label_Printing(ByVal Type As String, ByVal BarCode As String, ByVal ItemName As String, ByVal Quantity As Integer, ByVal UnitPrice As Double, ByVal Print_UnitPrice As String, ByVal PrinterName As String, ByVal Batch As String, ByVal BranchName As String, ByVal ItemCode As String, ByVal UOMEntry As String)
    '    Try
    '        Dim PrinterPath As String = ""
    '        Dim PRNFile As String = ""
    '        If Type = "1" Then
    '            PrinterPath = Me.GetFieldValue("SELECT ""U_PrinterPath"" FROM ""@NT_BARPRSETUP"" WHERE ""U_Type""='1' and ""U_PrinterName""='" & PrinterName & "'")
    '            PRNFile = Me.Notepad_Read("Type1.txt")
    '        ElseIf Type = "2" Then
    '            PrinterPath = Me.GetFieldValue("SELECT ""U_PrinterPath"" FROM ""@NT_BARPRSETUP"" WHERE ""U_Type""='2' and ""U_PrinterName""='" & PrinterName & "'")
    '            PRNFile = Me.Notepad_Read("Type2.txt")
    '        ElseIf Type = "3" Then
    '            PrinterPath = Me.GetFieldValue("SELECT ""U_PrinterPath"" FROM ""@NT_BARPRSETUP"" WHERE ""U_Type""='3' and ""U_PrinterName""='" & PrinterName & "'")
    '            PRNFile = Me.Notepad_Read("Type3.txt")

    '        ElseIf Type = "4" Then
    '            PrinterPath = Me.GetFieldValue("SELECT ""U_PrinterPath"" FROM ""@NT_BARPRSETUP"" WHERE ""U_Type""='4' and ""U_PrinterName""='" & PrinterName & "'")
    '            PRNFile = Me.Notepad_Read("Type4.txt")
    '        End If
    '        If PrinterPath = "" Then
    '            Me.objForm.Freeze(False)
    '            objMain.objApplication.StatusBar.SetText("Printer Path is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Return
    '        ElseIf PRNFile <> "" Then

    '            Dim StrTime As String = Format(DateTime.Now, "HHMMssfff")
    '            Dim StrDate As String = Format(DateTime.Now, "ddMMyy")
    '            '------------
    '            Dim FileName As String = "Print_" & objMain.objCompany.UserName & StrTime & ".prn"
    '            Dim FinalPRNFile As String = String.Concat(System.Windows.Forms.Application.StartupPath, "\", FileName)
    '            If (File.Exists(FinalPRNFile)) Then
    '                File.WriteAllText(FinalPRNFile, "")
    '            Else
    '                File.Create(FinalPRNFile).Close()
    '            End If
    '            '-----------------------------------------------
    '            'Dim ItemGrpQuery As String = " SELECT T1.""ItmsGrpCod""  FROM ""OITM"" T1 WHERE T1.""ItmsGrpCod"" IN (134,139) AND T1.""ItemCode"" = (SELECT T0.""ItemCode"" FROM OBCD T0 WHERE T0.""BcdCode""='" & BarCode & "')"
    '            'Dim oRsItemGrp As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '            'oRsItemGrp.DoQuery(ItemGrpQuery)

    '            'If oRsItemGrp.RecordCount = 0 Then
    '            '    PRNFile = PRNFile.Replace("[VAT]", "(VAT Incl.)")
    '            'Else
    '            '    PRNFile = PRNFile.Replace("[VAT]", " ")
    '            'End If

    '            '-----------------------------------------------
    '            Dim BasePrice As Double = 0
    '            Try
    '                Dim ItemQuery As String = " SELECT T0.""U_PIC"", T0.""U_BsUnit"" FROM OITM T0 Where T0.""ItemCode""='" & ItemCode & "'  and IFNULL(T0.""U_PIC"",0)<>0 and IFNULL(T0.""U_BsUnit"",0)<>0"
    '                Dim oRsItem As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                oRsItem.DoQuery(ItemQuery)
    '                If oRsItem.RecordCount > 0 Then
    '                    Dim BaseQtyQry As String = "SELECT T1.""BaseQty"" FROM OUGP T0  " & _
    '                                             " INNER JOIN UGP1 T1 ON T0.""UgpEntry"" = T1.""UgpEntry"" " & _
    '                                             " INNER JOIN OITM T2 ON T0.""UgpEntry"" = T2.""UgpEntry""  " & _
    '                                             " WHERE T1.""UomEntry"" ='" & UOMEntry & "' and T2.""ItemCode""='" & ItemCode & "'"
    '                    Dim oRsBaseQty As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                    oRsBaseQty.DoQuery(BaseQtyQry)
    '                    Dim BaseQty As Double
    '                    If oRsBaseQty.RecordCount > 0 Then
    '                        BaseQty = oRsBaseQty.Fields.Item(0).Value
    '                    End If
    '                    Dim Pic As Double = oRsItem.Fields.Item("U_PIC").Value
    '                    Dim BaseUnit As Double = oRsItem.Fields.Item("U_BsUnit").Value
    '                    If Pic <> 0 And BaseUnit <> 0 Then
    '                        BasePrice = UnitPrice / ((BaseUnit) / 1000)
    '                        BasePrice = UnitPrice / ((BaseQty * BaseUnit) * Pic / 1000)
    '                    End If
    '                End If
    '            Catch ex As Exception
    '                objMain.objApplication.StatusBar.SetText("Base price calculation error:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            End Try
    '            Dim ArabicName As String = Me.GetFieldValue("SELECT T0.""FrgnName"" FROM OITM T0 WHERE T0.""ItemCode""='" & ItemCode & "'")
    '            PRNFile = PRNFile.Replace("[ArabicName]", ArabicName)

    '            '------------------------------------------------
    '            Try
    '                PRNFile = PRNFile.Replace("[Batch]", Batch)
    '            Catch ex As Exception
    '            End Try

    '            PRNFile = PRNFile.Replace("[BasePrice]", BasePrice)
    '            PRNFile = PRNFile.Replace("[ItemCode]", ItemCode)

    '            PRNFile = PRNFile.Replace("[BarCode]", BarCode)
    '            PRNFile = PRNFile.Replace("[ItemName]", ItemName.ToUpper())
    '            PRNFile = PRNFile.Replace("[Quantity]", Quantity.ToString().Trim())
    '            If Print_UnitPrice.Trim() = "Y" Then
    '                PRNFile = PRNFile.Replace("[Currency]", "")
    '                PRNFile = PRNFile.Replace("[UnitPrice]", "")
    '            Else
    '                PRNFile = PRNFile.Replace("[Currency]", "AED")
    '                PRNFile = PRNFile.Replace("[UnitPrice]", Strings.Format(UnitPrice, "0.00"))
    '            End If
    '            PRNFile = PRNFile.Replace("[Date]", StrDate)
    '            PRNFile = PRNFile.Replace("[Location]", BranchName)

    '            If (File.Exists(FinalPRNFile)) Then
    '                Using streamWriter As System.IO.StreamWriter = New System.IO.StreamWriter(FinalPRNFile, True, Encoding.UTF8)
    '                    streamWriter.Write(If(If(PRNFile, ""), ""))
    '                    streamWriter.Close()
    '                End Using
    '            End If

    '            Try
    '                IO.File.Copy(FinalPRNFile, PrinterPath & "\" & FileName)
    '                'IO.File.Move(FinalPRNFile, PrinterPath)
    '                objMain.objApplication.StatusBar.SetText("Barcode print file sent to the printer successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    '                Thread.Sleep(300)
    '            Catch ex As Exception
    '                objMain.objApplication.StatusBar.SetText("An error while sending to the printer. Error:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            End Try

    '            Try
    '                IO.File.Delete(FinalPRNFile)
    '            Catch ex As Exception
    '            End Try
    '        Else
    '            Me.objForm.Freeze(False)
    '            objMain.objApplication.StatusBar.SetText("Zebra print text is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Return
    '        End If
    '    Catch exception1 As System.Exception
    '        Dim exception As System.Exception = exception1
    '        objMain.objApplication.StatusBar.SetText(exception.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '    End Try
    'End Sub

    Public Function GetFieldValue(ByVal query As String) As Object
        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs.DoQuery(query)
        Dim str As String = oRs.Fields.Item(0).Value.ToString()
        Return str
    End Function

    Public Function Notepad_Read(ByVal QueryTextFile As String) As String
        Dim str As String = Nothing
        Try
            Dim sPath As String
            sPath = IO.Directory.GetParent(Application.ExecutablePath).ToString
            sPath = sPath & "\" & QueryTextFile

            If (Not File.Exists(sPath)) Then
                objMain.objApplication.StatusBar.SetText("File is not in this path - '" & sPath & "'", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                Dim streamReader As System.IO.StreamReader = New System.IO.StreamReader(sPath, Encoding.UTF8)
                Dim [end] As String = ""
                [end] = streamReader.ReadToEnd()
                streamReader.Close()
                str = [end]
                Return str
            End If
        Catch exception As System.Exception
            objMain.objApplication.StatusBar.SetText(exception.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
        Return str
    End Function

#Region "        NumberToArabic        "
    'Private Function toPersianNumber(ByVal input As String, ByVal vallength As Int32) As String        
    '    Dim persian As String() = New String(9) {"۰", "١", "٢", "٣", "٤", "٥", "٦", "٧", "٨", "٩"}
    '    For k As Integer = 0 To vallength - 1
    '        input = input.Replace(input.Substring(k, vallength - 1).ToString(), persian(input.Substring(k, vallength - 1)))
    '    Next
    '    input = input.Replace(" ", "")
    '    input = Regex.Replace(input, "\s+", "")
    '    input = input.Replace(" ", "")
    '    Return input
    'End Function

    Private Function toArabicPrice(ByVal Price As String) As String
        'Dim sPrice As String
        'Dim vallong As Int32 = 0
        'Dim value As String = Price
        'value = value.Replace(" ", "")
        'vallong = value.Length
        'Dim charArray As Char() = value.ToCharArray()
        ''Array.Reverse(charArray) ' for right to left format
        'sPrice = toPersianNumber(New String(charArray), Convert.ToInt32(vallong))
        'value = sPrice
        'value = value.Replace(" ", "")
        'sPrice = value.Trim()
        Dim arrayList As New List(Of String)()
        For i As Integer = 1 To Price.Length  'Step -1
            arrayList.Add(ToPersianNumber(Price(i - 1)) + "‎")
        Next
        Return String.Join("", arrayList)
    End Function

    Private Function ToPersianNumber(input As String) As String
        'Dim persian As String() = New String(9) {"۰", "۱", "۲", "۳", "۴", "۵", "۶", "۷", "۸", "۹"}
        Dim persian As String() = New String(9) {"۰", "١", "٢", "٣", "٤", "٥", "٦", "٧", "٨", "٩"}
        For j As Integer = 0 To persian.Length - 1
            input = input.Replace(j.ToString(), persian(j))
        Next
        Return input
    End Function
#End Region

    'Public Sub Label_Printing(ByVal Type As String, ByVal BarCode As String, ByVal ItemName As String, ByVal Quantity As Integer, ByVal UnitPrice As Double, ByVal Print_UnitPrice As String, ByVal PrinterName As String, ByVal Batch As String, ByVal BranchName As String, ByVal ItemCode As String, ByVal UOMEntry As String)
    '    Try
    '        Dim PrinterPath As String = ""
    '        Dim PRNFile As String = ""
    '        If Type = "1" Then
    '            PrinterPath = Me.GetFieldValue("SELECT ""U_PrinterPath"" FROM ""@NT_BARPRSETUP"" WHERE ""U_Type""='1' and ""U_PrinterName""='" & PrinterName & "'")
    '            PRNFile = Me.Notepad_Read("Type1.txt")
    '        ElseIf Type = "2" Then
    '            PrinterPath = Me.GetFieldValue("SELECT ""U_PrinterPath"" FROM ""@NT_BARPRSETUP"" WHERE ""U_Type""='2' and ""U_PrinterName""='" & PrinterName & "'")
    '            PRNFile = Me.Notepad_Read("Type2.txt")
    '        ElseIf Type = "3" Then
    '            PrinterPath = Me.GetFieldValue("SELECT ""U_PrinterPath"" FROM ""@NT_BARPRSETUP"" WHERE ""U_Type""='3' and ""U_PrinterName""='" & PrinterName & "'")
    '            PRNFile = Me.Notepad_Read("Type3.txt")

    '        ElseIf Type = "4" Then
    '            PrinterPath = Me.GetFieldValue("SELECT ""U_PrinterPath"" FROM ""@NT_BARPRSETUP"" WHERE ""U_Type""='4' and ""U_PrinterName""='" & PrinterName & "'")
    '            PRNFile = Me.Notepad_Read("Type4.txt")
    '        End If
    '        If PrinterPath = "" Then
    '            Me.objForm.Freeze(False)
    '            objMain.objApplication.StatusBar.SetText("Printer Path is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Return
    '        ElseIf PRNFile <> "" Then

    '            Dim StrTime As String = Format(DateTime.Now, "HHMMssfff")
    '            Dim StrDate As String = Format(DateTime.Now, "ddMMyy")
    '            '------------
    '            Dim FileName As String = "Print_" & objMain.objCompany.UserName & StrTime & ".prn"
    '            Dim FinalPRNFile As String = String.Concat(System.Windows.Forms.Application.StartupPath, "\", FileName)
    '            If (File.Exists(FinalPRNFile)) Then
    '                File.WriteAllText(FinalPRNFile, "")
    '            Else
    '                File.Create(FinalPRNFile).Close()
    '            End If
    '            '-----------------------------------------------
    '            'Dim ItemGrpQuery As String = " SELECT T1.""ItmsGrpCod""  FROM ""OITM"" T1 WHERE T1.""ItmsGrpCod"" IN (134,139) AND T1.""ItemCode"" = (SELECT T0.""ItemCode"" FROM OBCD T0 WHERE T0.""BcdCode""='" & BarCode & "')"
    '            'Dim oRsItemGrp As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '            'oRsItemGrp.DoQuery(ItemGrpQuery)

    '            'If oRsItemGrp.RecordCount = 0 Then
    '            '    PRNFile = PRNFile.Replace("[VAT]", "(VAT Incl.)")
    '            'Else
    '            '    PRNFile = PRNFile.Replace("[VAT]", " ")
    '            'End If

    '            '-----------------------------------------------
    '            Dim BasePrice As Double = 0
    '            Dim oRsExp As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '            Try
    '                ''Dim ItemQuery As String = " SELECT T0.""U_PIC"", T0.""U_BsUnit"", T0.""U_BaseUnit"" FROM OITM T0 Where T0.""ItemCode""='" & ItemCode & "'  and IFNULL(T0.""U_PIC"",0)<>0 and IFNULL(T0.""U_BaseUnit"",0)<>0"

    '                Dim ItemQuery As String = " SELECT T0.""U_PcPerPac"", T0.""U_BsUnit"", T0.""U_BaseUnit"" FROM OITM T0 Where T0.""ItemCode""='" & ItemCode & "'  and IFNULL(T0.""U_PcPerPac"",0)<>0 and IFNULL(T0.""U_BaseUnit"",0)<>0"
    '                Dim oRsItem As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                oRsItem.DoQuery(ItemQuery)

    '                Dim ExpQuery As String = " SELECT ADD_DAYS(CURRENT_DATE,T0.""U_Life"") As ""ExpDate"" FROM OITM T0 Where T0.""ItemCode""='" & ItemCode & "'  and IFNULL(T0.""U_Life"",0)<>0 "
    '                oRsExp.DoQuery(ExpQuery)

    '                If oRsItem.RecordCount > 0 Then
    '                    Dim BaseQtyQry As String = "SELECT T1.""BaseQty"" FROM OUGP T0  " &
    '                                             " INNER JOIN UGP1 T1 ON T0.""UgpEntry"" = T1.""UgpEntry"" " &
    '                                             " INNER JOIN OITM T2 ON T0.""UgpEntry"" = T2.""UgpEntry""  " &
    '                                             " WHERE T1.""UomEntry"" ='" & UOMEntry & "' and T2.""ItemCode""='" & ItemCode & "'"
    '                    Dim oRsBaseQty As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                    oRsBaseQty.DoQuery(BaseQtyQry)
    '                    Dim BaseQty As Double
    '                    If oRsBaseQty.RecordCount > 0 Then
    '                        BaseQty = oRsBaseQty.Fields.Item(0).Value
    '                    End If
    '                    '' Dim Pic As Double = Convert.ToDouble(oRsItem.Fields.Item("U_PIC").Value)
    '                    Dim Pic As Double = Convert.ToDouble(oRsItem.Fields.Item("U_PcPerPac").Value)
    '                    'Dim BaseUnit As Double = oRsItem.Fields.Item("U_BsUnit").Value
    '                    Dim BaseUnit As Double = 0
    '                    BaseUnit = Convert.ToDouble(oRsItem.Fields.Item("U_BaseUnit").Value)
    '                    If Pic <> 0 And BaseUnit <> 0 Then
    '                        BasePrice = UnitPrice / ((BaseUnit) / 1000)
    '                        BasePrice = UnitPrice / ((BaseQty * BaseUnit) * Pic / 1000)
    '                        BasePrice = Math.Round(BasePrice, 2)
    '                    End If
    '                End If
    '            Catch ex As Exception
    '                objMain.objApplication.StatusBar.SetText("Base price calculation error:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            End Try

    '            Try

    '            Catch ex As Exception

    '            End Try
    '            Dim ArabicName As String = Me.GetFieldValue("SELECT T0.""FrgnName"" FROM OITM T0 WHERE T0.""ItemCode""='" & ItemCode & "'")
    '            PRNFile = PRNFile.Replace("[ArabicName]", ArabicName)

    '            '------------------------------------------------
    '            Try
    '                PRNFile = PRNFile.Replace("[Batch]", Batch)
    '            Catch ex As Exception
    '            End Try
    '            PRNFile = PRNFile.Replace("[BasePrice]", BasePrice)
    '            PRNFile = PRNFile.Replace("[ArabicBasePrice]", toArabicPrice(BasePrice.ToString))
    '            PRNFile = PRNFile.Replace("[ItemCode]", ItemCode)
    '            PRNFile = PRNFile.Replace("[BarCode]", BarCode)
    '            PRNFile = PRNFile.Replace("[ItemName]", ItemName.ToUpper())
    '            PRNFile = PRNFile.Replace("[Quantity]", Quantity.ToString().Trim())

    '            If Print_UnitPrice.Trim() = "Y" Then
    '                PRNFile = PRNFile.Replace("[Currency]", "")
    '                PRNFile = PRNFile.Replace("[UnitPrice]", "")
    '            Else
    '                PRNFile = PRNFile.Replace("[Currency]", "AED")
    '                PRNFile = PRNFile.Replace("[UnitPrice]", Strings.Format(UnitPrice, "0.00"))
    '                PRNFile = PRNFile.Replace("[ArabicPrice]", toArabicPrice(Strings.Format(UnitPrice, "0.00"))) 'toArabicPrice(UnitPrice.ToString))
    '            End If
    '            Try
    '                ''[UnitPrice]
    '                '' PRNFile = PRNFile.Replace("[UnitPrice]", UnitPrice.ToString().Trim())
    '                PRNFile = PRNFile.Replace("[QRKGLTR]", BasePrice.ToString().Trim())
    '            Catch ex As Exception

    '            End Try
    '            PRNFile = PRNFile.Replace("[Date]", StrDate)
    '            PRNFile = PRNFile.Replace("[Location]", BranchName)
    '            If oRsExp.RecordCount > 0 Then
    '                Dim ExpDate As Date = oRsExp.Fields.Item(0).Value
    '                PRNFile = PRNFile.Replace("[ProdDate]", Date.Now.ToString("ddMMyyyy"))
    '                PRNFile = PRNFile.Replace("[ExpDate]", ExpDate.ToString("ddMMyyyy"))
    '            End If

    '            If (File.Exists(FinalPRNFile)) Then
    '                Using streamWriter As System.IO.StreamWriter = New System.IO.StreamWriter(FinalPRNFile, True, Encoding.UTF8)
    '                    streamWriter.Write(If(If(PRNFile, ""), ""))
    '                    streamWriter.Close()
    '                End Using
    '            End If

    '            Try
    '                IO.File.Copy(FinalPRNFile, PrinterPath & "\" & FileName)
    '                'IO.File.Move(FinalPRNFile, PrinterPath)
    '                objMain.objApplication.StatusBar.SetText("Barcode print file sent to the printer successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    '                Thread.Sleep(300)
    '            Catch ex As Exception
    '                objMain.objApplication.StatusBar.SetText("An error while sending to the printer. Error:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            End Try

    '            Try
    '                IO.File.Delete(FinalPRNFile)
    '            Catch ex As Exception
    '            End Try
    '        Else
    '            Me.objForm.Freeze(False)
    '            objMain.objApplication.StatusBar.SetText("Zebra print text is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Return
    '        End If
    '    Catch exception1 As System.Exception
    '        Dim exception As System.Exception = exception1
    '        objMain.objApplication.StatusBar.SetText(exception.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '    End Try
    'End Sub
    Public Sub Label_Printing(ByVal Type As String, ByVal BarCode As String, ByVal ItemName As String, ByVal Quantity As Integer, ByVal UnitPrice As Double, ByVal Print_UnitPrice As String, ByVal PrinterName As String, ByVal Batch As String, ByVal BranchName As String, ByVal ItemCode As String, ByVal UOMEntry As String)
        Try
            Dim PrinterPath As String = ""
            Dim PRNFile As String = ""
            If Type = "1" Then
                PrinterPath = Me.GetFieldValue("SELECT ""U_PrinterPath"" FROM ""@NT_BARPRSETUP"" WHERE ""U_Type""='1' and ""U_PrinterName""='" & PrinterName & "'")
                PRNFile = Me.Notepad_Read("Type1.txt")
            ElseIf Type = "2" Then
                PrinterPath = Me.GetFieldValue("SELECT ""U_PrinterPath"" FROM ""@NT_BARPRSETUP"" WHERE ""U_Type""='2' and ""U_PrinterName""='" & PrinterName & "'")
                PRNFile = Me.Notepad_Read("Type2.txt")
            ElseIf Type = "3" Then
                PrinterPath = Me.GetFieldValue("SELECT ""U_PrinterPath"" FROM ""@NT_BARPRSETUP"" WHERE ""U_Type""='3' and ""U_PrinterName""='" & PrinterName & "'")
                PRNFile = Me.Notepad_Read("Type3.txt")

            ElseIf Type = "4" Then
                PrinterPath = Me.GetFieldValue("SELECT ""U_PrinterPath"" FROM ""@NT_BARPRSETUP"" WHERE ""U_Type""='4' and ""U_PrinterName""='" & PrinterName & "'")
                PRNFile = Me.Notepad_Read("Type4.txt")
            End If
            If PrinterPath = "" Then
                Me.objForm.Freeze(False)
                objMain.objApplication.StatusBar.SetText("Printer Path is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return
            ElseIf PRNFile <> "" Then
                'Changed
                ' Dim StrTime As String = Format(DateTime.Now, "HH:mm")

                Dim StrTime As String = Format(DateTime.Now, "HHMMssfff")
                Dim StrDate As String = Format(DateTime.Now, "dd/MM/yy")

                'Dim StrDate As String = Format(DateTime.Now, "ddMMyy")
                '------------
                Dim FileName As String = "Print_" & objMain.objCompany.UserName & StrTime & ".prn"
                Dim FinalPRNFile As String = String.Concat(System.Windows.Forms.Application.StartupPath, "\", FileName)
                If (File.Exists(FinalPRNFile)) Then
                    File.WriteAllText(FinalPRNFile, "")
                Else
                    File.Create(FinalPRNFile).Close()
                End If

                Dim BasePrice As Double = 0
                Dim BaseUnit As Double = 0
                Dim BaseUOM As String = ""
                Dim BaseUnitPrice As Double = 0
                Try

                    '----------------------------------

                    ' 1. Get Item details

                    '----------------------------------

                    Dim ItemQuery As String = " SELECT T0.""U_PcPerPac"", T0.""U_BsUnit"", T0.""U_BaseUnit"", " &
                " IFNULL(T0.""U_WEIGHT"",0) as ""U_WEIGHT"", T0.""U_UOM"" " &
                " FROM OITM T0 " &
                " WHERE T0.""ItemCode""='" & ItemCode & "' AND IFNULL(T0.""U_WEIGHT"",0) > 0.00001 "
                    Dim oRsItem As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsItem.DoQuery(ItemQuery)




                    If oRsItem.RecordCount > 0 Then
                        BaseUnitPrice = UnitPrice   ' ✅ the passed-in unit price directly
                        ' 3. Calculate base price per KG (or UOM)

                        BaseUnit = CDbl(oRsItem.Fields.Item("U_WEIGHT").Value)
                        BaseUOM = oRsItem.Fields.Item("U_UOM").Value.ToString()
                        If BaseUnit > 0 And BaseUnitPrice > 0 Then
                            BasePrice = BaseUnitPrice * (1000 / BaseUnit)
                            BasePrice = Math.Round(BasePrice, 2)
                        Else
                            BasePrice = 0
                        End If
                        '----------------------------------
                        ' 4. Show result in status bar
                        '----------------------------------
                        Dim priceType As String = If(UnitPrice = 0, "NO PRICE", "SELECTED PRICE")
                        objMain.objApplication.StatusBar.SetText("Unit Price based on " & priceType & ": " & BasePrice.ToString(),
                                   SAPbouiCOM.BoMessageTime.bmt_Short,
                                   SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Else
                        objMain.objApplication.StatusBar.SetText(" Item not found or Weight = 0",
                                   SAPbouiCOM.BoMessageTime.bmt_Short,
                                   SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                    End If

                Catch ex As Exception

                    objMain.objApplication.StatusBar.SetText(" Unit price calculation error: " & ex.Message,
                               SAPbouiCOM.BoMessageTime.bmt_Short,
                               SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                End Try

                Dim ArabicName As String = Me.GetFieldValue("SELECT T0.""U_ShortName"" FROM OITM T0 WHERE T0.""ItemCode""='" & ItemCode & "'")
                PRNFile = PRNFile.Replace("[ArabicName]", ArabicName)
                Try
                    Dim VatAmount As Double = Me.GetFieldValue("Call ""Barcode_Vat""('" & ItemCode & "','" & UOMEntry & "')")
                    Dim WithoutVatPrice As Double = Me.GetFieldValue("Call ""Barcode_WithoutVatPrice""('" & ItemCode & "','" & UOMEntry & "')")

                    Dim formattedVatAmount As String = If(IsNothing(VatAmount), "0.00", Strings.Format(VatAmount, "0.00"))
                    Dim formattedWithoutVatPrice As String = If(IsNothing(WithoutVatPrice), "0.00", Strings.Format(WithoutVatPrice, "0.00"))

                    If formattedVatAmount <> "0.00" AndAlso formattedVatAmount <> "" AndAlso Not String.IsNullOrEmpty(formattedVatAmount) Then
                        PRNFile = PRNFile.Replace("[VatAmount]", formattedVatAmount)
                    Else
                        PRNFile = PRNFile.Replace("[VatAmount]", "")
                    End If
                    If formattedWithoutVatPrice <> "0.00" AndAlso formattedWithoutVatPrice <> "" AndAlso Not String.IsNullOrEmpty(formattedWithoutVatPrice) Then
                        PRNFile = PRNFile.Replace("[WithoutVatPrice]", formattedWithoutVatPrice)
                    Else
                        PRNFile = PRNFile.Replace("[WithoutVatPrice]", "")
                    End If

                Catch ex As Exception

                End Try
                '------------------------------------------------
                Try
                    PRNFile = PRNFile.Replace("[Batch]", Batch)
                    PRNFile = PRNFile.Replace("[ItemCode]", ItemCode)
                Catch ex As Exception
                End Try

                PRNFile = PRNFile.Replace("[ArabicBasePrice]", ConvertToArabicNumber(BasePrice.ToString))
                PRNFile = PRNFile.Replace("[123]", BarCode)
                PRNFile = PRNFile.Replace("[BarCode]", BarCode)
                PRNFile = PRNFile.Replace("[ItemName]", ItemName.ToUpper())
                PRNFile = PRNFile.Replace("[Quantity]", Quantity.ToString().Trim())
                If BasePrice <> 0 Then
                    Dim formattedBasePrice As String = If(IsNothing(BasePrice), "0.00", Strings.Format(BasePrice, "0.00"))
                    PRNFile = PRNFile.Replace("[BasePrice]", formattedBasePrice)
                Else
                    PRNFile = PRNFile.Replace("[BasePrice]", "")
                End If

                If Print_UnitPrice.Trim() = "N" Then
                    PRNFile = PRNFile.Replace("[Currency]", "")
                    PRNFile = PRNFile.Replace("[UnitPrice]", "")
                    PRNFile = PRNFile.Replace("[AP]", "")
                    PRNFile = PRNFile.Replace("[ABP]", "")
                    PRNFile = PRNFile.Replace("[BasePrice]", "")

                Else
                    'PRNFile = PRNFile.Replace("[Currency]", "QR")
                    'PRNFile = PRNFile.Replace("[Currency]", "ر.ق")

                    ' PRNFile = PRNFile.Replace("[UnitPrice]", Strings.Format(UnitPrice, "0.00"))

                    PRNFile = PRNFile.Replace("[Currency]", "OMR") ' CHANGED HERE (QR → OMR)
                    PRNFile = PRNFile.Replace("[UnitPrice]", UnitPrice.ToString("F3")) '  CHANGED FORMAT (2 → 3 decimals)

                    Dim ArabicPrice As String = ConvertToArabicNumber(Strings.Format(UnitPrice, "0.00"))
                    Dim ArabicBasePrice As String = ConvertToArabicNumber(Strings.Format(BasePrice, "0.00"))
                    '  PRNFile = PRNFile.Replace("[AP]", ArabicPrice)
                    Dim formattedUnitPrice As String = If(IsNothing(UnitPrice), "0.00", Strings.Format(UnitPrice, "0.00"))
                    Dim formattedBasePrice As String = If(IsNothing(BasePrice), "0.00", Strings.Format(BasePrice, "0.00"))

                    'PRNFile = PRNFile.Replace("[ABP]", ArabicBasePrice.ToString().Trim)
                    'PRNFile = PRNFile.Replace("[BasePrice]", BasePrice.ToString().Trim())
                    If formattedUnitPrice <> "0.00" AndAlso formattedUnitPrice <> "" AndAlso Not String.IsNullOrEmpty(formattedUnitPrice) Then
                        PRNFile = PRNFile.Replace("[AP]", ArabicPrice)
                    End If

                    If formattedBasePrice <> "0.00" AndAlso formattedBasePrice <> "" AndAlso Not String.IsNullOrEmpty(formattedBasePrice) Then
                        PRNFile = PRNFile.Replace("[ABP]", ArabicBasePrice)
                        PRNFile = PRNFile.Replace("[BasePrice]", formattedBasePrice)
                        PRNFile = PRNFile.Replace("[BasePrice]", Strings.Format(BasePrice, "0.00"))
                        PRNFile = PRNFile.Replace("[BasePrice]", Decimal.Parse(formattedBasePrice).ToString("F2"))


                    Else
                        PRNFile = PRNFile.Replace("[ABP]", "")
                        PRNFile = PRNFile.Replace("[BasePrice]", "")
                    End If
                End If
                Try
                    PRNFile = PRNFile.Replace("[BU]", BaseUOM.ToString().Trim())
                    PRNFile = PRNFile.Replace("[QRKGLTR]", BasePrice.ToString().Trim())


                    PRNFile = PRNFile.Replace("[Time]", DateTime.Now().ToString("HH:mm"))

                Catch ex As Exception

                End Try
                PRNFile = PRNFile.Replace("[Date]", StrDate)
                PRNFile = PRNFile.Replace("[Location]", BranchName)

                Dim oRsExp As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim ExpQuery As String = " SELECT ADD_DAYS(CURRENT_DATE,T0.""U_Life"") As ""ExpDate"" FROM OITM T0 Where T0.""ItemCode""='" & ItemCode & "'  and IFNULL(T0.""U_Life"",0)<>0 "
                oRsExp.DoQuery(ExpQuery)

                If oRsExp.RecordCount > 0 Then
                    Dim ExpDate As Date = oRsExp.Fields.Item(0).Value

                    PRNFile = PRNFile.Replace("[ExpDate]", ExpDate.ToString("ddMMyyyy"))
                End If

                If (File.Exists(FinalPRNFile)) Then
                    Using streamWriter As System.IO.StreamWriter = New System.IO.StreamWriter(FinalPRNFile, True, Encoding.UTF8)
                        streamWriter.Write(If(If(PRNFile, ""), ""))
                        streamWriter.Close()
                    End Using
                End If

                Try
                    IO.File.Copy(FinalPRNFile, PrinterPath & "\" & FileName)
                    'IO.File.Move(FinalPRNFile, PrinterPath)
                    objMain.objApplication.StatusBar.SetText("Barcode print file sent to the printer successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Thread.Sleep(300)
                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText("An error while sending to the printer. Error:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End Try

                Try
                    IO.File.Delete(FinalPRNFile)
                Catch ex As Exception
                End Try
            Else
                Me.objForm.Freeze(False)
                objMain.objApplication.StatusBar.SetText("Zebra print text is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return
            End If
        Catch exception1 As System.Exception
            Dim exception As System.Exception = exception1
            objMain.objApplication.StatusBar.SetText(exception.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
    End Sub

    Public Function ConvertToArabicNumber(ByVal price As String) As String
        Dim arabicDigits() As Char = {"٠"c, "١"c, "٢"c, "٣"c, "٤"c, "٥"c, "٦"c, "٧"c, "٨"c, "٩"c}
        Dim result As String = ""

        For Each ch As Char In price
            If Char.IsDigit(ch) Then
                result &= arabicDigits(Convert.ToInt32(ch.ToString()))
            Else
                result &= ch ' Keep decimal points, commas, currency symbols, etc.
            End If
        Next

        Return result
    End Function

    '------Naresh
    Public Function ExecuteSQLScript(company As SAPbobsCOM.Company, sqlScript As String) As Boolean
        Try
            Dim rs As SAPbobsCOM.Recordset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            rs.DoQuery(sqlScript)



            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
            rs = Nothing
            GC.Collect()
            Return True
        Catch ex As Exception
            ' Log the error or handle accordingly
            'MsgBox("Error executing SQL: " & ex.Message)
            Return False
        End Try
    End Function

    Public Function LoadEmbeddedSQL(resourceName As String) As String
        Try


            Dim assembly = System.Reflection.Assembly.GetExecutingAssembly()
            Dim sExecutingAssemblyNmae As String = assembly.GetName().Name.ToString()
            Dim xmldoc As New Xml.XmlDocument
            resourceName = sExecutingAssemblyNmae + "." + resourceName

            Using stream = assembly.GetManifestResourceStream(resourceName)
                If stream Is Nothing Then Throw New Exception("Embedded SQL resource not found: " & resourceName)
                Using reader As New IO.StreamReader(stream)
                    Return reader.ReadToEnd()
                End Using
            End Using
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error loading SQL: " & resourceName & " - " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False

        End Try
    End Function

    Public Function ReadTextFile(filePath As String) As String
        If Not IO.File.Exists(filePath) Then
            Throw New FileNotFoundException("SQL file not found: " & filePath)
        End If
        Return IO.File.ReadAllText(filePath)
    End Function

    Public Function AddOrUpdateQuery(company As SAPbobsCOM.Company, queryName As String, categoryId As Integer, sqlText As String, ByVal Type As String) As Boolean
        Try
            Dim oUserQuery As SAPbobsCOM.UserQueries = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries)

            ' Check if query already exists
            Dim rs As SAPbobsCOM.Recordset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim Qry As String = "SELECT ""IntrnalKey"", ""QCategory"" FROM OUQR WHERE ""QName"" = '" & queryName & "'"
            rs.DoQuery(Qry)
            Dim key As Integer = CInt(rs.Fields.Item("IntrnalKey").Value)
            Dim categoryID1 As Integer = CInt(rs.Fields.Item("QCategory").Value)

            If rs.RecordCount > 0 Then
                ' Update existing query


                If oUserQuery.GetByKey(key, categoryID1) Then
                    oUserQuery.Query = sqlText
                    If oUserQuery.Update() <> 0 Then
                        Throw New Exception("Error updating query: " & company.GetLastErrorDescription())
                    Else
                        If Type = "Alert" Then
                            Dim rs1 As SAPbobsCOM.Recordset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim Qry1 As String = "SELECT ""IntrnalKey"" FROM OUQR WHERE ""QName"" = '" & queryName & "'"
                            rs1.DoQuery(Qry1)
                            If rs1.RecordCount > 0 Then
                                Dim key1 As Integer = CInt(rs1.Fields.Item("IntrnalKey").Value)

                                Dim rsAlertCheck As SAPbobsCOM.Recordset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim alertCheckQuery As String = "SELECT T0.""Code"", T0.""Name"" FROM OALT T0 where T0.""Name"" = '" & queryName & "'"
                                rsAlertCheck.DoQuery(alertCheckQuery)

                                If rsAlertCheck.RecordCount = 0 Then

                                    CreateCustomAlert(queryName, key1, 1, 30, AlertManagementFrequencyType.atfi_Minutes)
                                End If
                            End If
                        Else
                            CreateApprovalTemplateWithQueryTerm("GoodsIssue", "Goods Issue Approval", 1, SAPbobsCOM.ApprovalTemplatesDocumentTypeEnum.atdtGoodsIssue, 1, categoryID1)
                            CreateApprovalTemplateWithQueryTerm("POPrice", "PO Price more than Last purchase price", 1, SAPbobsCOM.ApprovalTemplatesDocumentTypeEnum.atdtPurchaseOrder, 1, categoryID1)
                            CreateApprovalTemplateWithQueryTerm("SalesOrder", "Sales Order Credit Limit Deviation", 1, SAPbobsCOM.ApprovalTemplatesDocumentTypeEnum.atdtOrder, 1, categoryID1)
                            CreateApprovalTemplateWithQueryTerm("ItemQtyMx", "Item Quantity deviates from Maximum Quantity", 1, SAPbobsCOM.ApprovalTemplatesDocumentTypeEnum.atdtPurchaseOrder, 1, categoryID1)

                        End If
                    End If
                End If
            Else
                ' Add new query
                oUserQuery.QueryCategory = categoryId
                oUserQuery.Query = sqlText
                oUserQuery.QueryDescription = queryName
                If oUserQuery.Add() <> 0 Then
                    Throw New Exception("Error adding query: " & company.GetLastErrorDescription())
                Else
                    If Type = "Alert" Then
                        Dim rs1 As SAPbobsCOM.Recordset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim Qry1 As String = "SELECT ""IntrnalKey"" FROM OUQR WHERE ""QName"" = '" & queryName & "'"
                        rs1.DoQuery(Qry1)
                        If rs1.RecordCount > 0 Then
                            Dim key1 As Integer = CInt(rs1.Fields.Item("IntrnalKey").Value)

                            Dim rsAlertCheck As SAPbobsCOM.Recordset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim alertCheckQuery As String = "SELECT T0.""Code"", T0.""Name"" FROM OALT T0 where T0.""Name"" = '" & queryName & "'"
                            rsAlertCheck.DoQuery(alertCheckQuery)

                            If rsAlertCheck.RecordCount = 0 Then

                                CreateCustomAlert(queryName, key1, 1, 30, AlertManagementFrequencyType.atfi_Minutes)
                            End If
                        End If
                    Else
                        CreateApprovalTemplateWithQueryTerm("GoodsIssue", "Goods Issue Approval", 1, SAPbobsCOM.ApprovalTemplatesDocumentTypeEnum.atdtGoodsIssue, 1, categoryID1)
                        CreateApprovalTemplateWithQueryTerm("POPrice", "PO Price more than Last purchase price", 1, SAPbobsCOM.ApprovalTemplatesDocumentTypeEnum.atdtPurchaseOrder, 1, categoryID1)
                        CreateApprovalTemplateWithQueryTerm("SalesOrder", "Sales Order Credit Limit Deviation", 1, SAPbobsCOM.ApprovalTemplatesDocumentTypeEnum.atdtOrder, 1, categoryID1)
                        CreateApprovalTemplateWithQueryTerm("ItemQtyMx", "Item Quantity deviates from Maximum Quantity", 1, SAPbobsCOM.ApprovalTemplatesDocumentTypeEnum.atdtPurchaseOrder, 1, categoryID1)

                    End If
                End If
            End If
            ' Clean up
            ' Marshal.ReleaseComObject(rs)
            ' Marshal.ReleaseComObject(oUserQuery)
            GC.Collect()
            objMain.objApplication.StatusBar.SetText("Query successfully added or updated in Query Manager:" & queryName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error while adding query:" & ex.Message & "Query:" & queryName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    'Public Sub AddAuthorization(ByVal authCode As String, ByVal authName As String, ByVal formUID As String, ByVal permission As SAPbobsCOM.BoUPTOptions)
    '    Dim oUserPermission As SAPbobsCOM.UserPermissionTree
    '    oUserPermission = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree), SAPbobsCOM.UserPermissionTree)

    '    ' Check if authorization already exists
    '    If oUserPermission.GetByKey(authCode) Then
    '        ' Optionally update existing
    '        oUserPermission.Name = authName
    '        oUserPermission.Options = permission
    '        If oUserPermission.Update() <> 0 Then
    '            MsgBox("Failed to update auth: " & objMain.objCompany.GetLastErrorDescription())
    '        End If
    '    Else
    '        ' Create new authorization
    '        oUserPermission.PermissionID = authCode
    '        oUserPermission.Name = authName
    '        oUserPermission.Options = permission
    '        oUserPermission.UserPermissionForms.FormType = formUID
    '        oUserPermission.UserPermissionForms.Add()

    '        If oUserPermission.Add() <> 0 Then
    '            MsgBox("Failed to add auth: " & objMain.objCompany.GetLastErrorDescription())
    '        End If
    '    End If

    '    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserPermission)
    '    oUserPermission = Nothing
    '    GC.Collect()
    'End Sub




    Public Sub AddAuthorization(ByVal authCode As String, ByVal authName As String, ByVal formUID As String, ByVal ParentForm As String, ByVal permission As SAPbobsCOM.BoUPTOptions)
        Dim oUserPermission As SAPbobsCOM.UserPermissionTree
        oUserPermission = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree), SAPbobsCOM.UserPermissionTree)

        ' Check if authorization already exists
        If oUserPermission.GetByKey(authCode) Then
            ' Optionally update existing
            oUserPermission.Name = authName
            oUserPermission.Options = permission
            oUserPermission.ParentID = ParentForm
            If oUserPermission.Update() <> 0 Then
                MsgBox("Failed to update auth: " & objMain.objCompany.GetLastErrorDescription())
            End If
        Else
            ' Create new authorization
            oUserPermission.PermissionID = authCode
            oUserPermission.Name = authName
            oUserPermission.Options = permission
            oUserPermission.ParentID = ParentForm
            oUserPermission.UserPermissionForms.FormType = formUID
            oUserPermission.UserPermissionForms.Add()

            If oUserPermission.Add() <> 0 Then
                MsgBox("Failed to add auth: " & objMain.objCompany.GetLastErrorDescription())
            End If
        End If

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserPermission)
        oUserPermission = Nothing
        GC.Collect()
    End Sub
    Public Function CreateCustomAlert(ByVal alertName As String, ByVal queryId As Integer, ByVal recipientUserCode As String, ByVal frequencyInterval As Integer, ByVal frequencyType As AlertManagementFrequencyType) As Boolean
        Try
            ' Get the Company Service
            Dim oCmpSrv As CompanyService = objMain.objCompany.GetCompanyService()

            ' Get the Alert Management Service
            Dim oAlertService As AlertManagementService = DirectCast(oCmpSrv.GetBusinessService(ServiceTypes.AlertManagementService), AlertManagementService)

            ' Create a new Alert Management object
            Dim oAlertManagement As AlertManagement = DirectCast(oAlertService.GetDataInterface(AlertManagementServiceDataInterfaces.atsdiAlertManagement), AlertManagement)
            ' Set alert properties
            oAlertManagement.Name = alertName
            oAlertManagement.Priority = AlertManagementPriorityEnum.atp_Low ' You can change priority as needed
            oAlertManagement.Active = BoYesNoEnum.tYES
            oAlertManagement.QueryID = queryId ' Assign the ID of your user-defined query
            oAlertManagement.FrequencyInterval = frequencyInterval
            oAlertManagement.FrequencyType = frequencyType
            oAlertManagement.SaveHistory = BoYesNoEnum.tYES ' Example: Save alert history

            Dim oRecipient As AlertManagementRecipient = oAlertManagement.AlertManagementRecipients.Add()
            oRecipient.UserCode = recipientUserCode
            oRecipient.SendInternal = BoYesNoEnum.tYES ' Send as an internal message in SAP Business One
            oRecipient.SendEmail = BoYesNoEnum.tYES ' Send as email
            ' oRecipient.EmailAddress = "recipient@example.com" ' Optionally specify an email address

            'oAlertService.AddAlert(oAlertManagement)
            oAlertService.AddAlertManagement(oAlertManagement)

            objMain.objApplication.StatusBar.SetText("Query successfully added or updated in Query Manager", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            ' MessageBox.Show("Alert '" & alertName & "' created successfully!")
            Return True

        Catch ex As Exception
            ' MessageBox.Show("Error creating alert: " & ex.Message)
            Return False
        End Try
    End Function

    Public Function CreateApprovalTemplateWithQueryTerm(ByVal templateName As String, ByVal templateDescription As String, ByVal originatorUserID As Integer, ByVal documentType As SAPbobsCOM.ApprovalTemplatesDocumentTypeEnum, ByVal approvalStageCode As Integer, ByVal userQueryID As Integer) As Integer
        ' Declare objects
        Dim oCmpSrv As SAPbobsCOM.CompanyService = Nothing
        Dim oApprovalTemplatesService As SAPbobsCOM.ApprovalTemplatesService = Nothing
        Dim oApprovalTemplate As SAPbobsCOM.ApprovalTemplate = Nothing
        Dim oApprovalTemplateStage As SAPbobsCOM.ApprovalTemplateStage = Nothing
        Dim oApprovalTemplateTerm As SAPbobsCOM.ApprovalTemplateTerm = Nothing
        Dim oApprovalTemplateParams As SAPbobsCOM.ApprovalTemplateParams = Nothing

        Dim result As Integer = -1 ' Default return value for error

        Try
            ' 1. Get CompanyService and ApprovalTemplatesService
            oCmpSrv = objMain.objCompany.GetCompanyService()
            oApprovalTemplatesService = CType(oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ApprovalTemplatesService), SAPbobsCOM.ApprovalTemplatesService)

            ' 2. Create Approval Template object
            oApprovalTemplate = CType(oApprovalTemplatesService.GetDataInterface(SAPbobsCOM.ApprovalTemplatesServiceDataInterfaces.atsdiApprovalTemplate), SAPbobsCOM.ApprovalTemplate)

            ' 3. Set general template properties
            oApprovalTemplate.Name = templateName
            oApprovalTemplate.Remarks = templateDescription
            oApprovalTemplate.IsActive = SAPbobsCOM.BoYesNoEnum.tYES ' Activate the template

            ' 4. Add originators
            oApprovalTemplate.ApprovalTemplateUsers.Add()
            oApprovalTemplate.ApprovalTemplateUsers.Item(0).UserID = originatorUserID

            ' 5. Add documents to be approved
            oApprovalTemplate.ApprovalTemplateDocuments.Add()
            oApprovalTemplate.ApprovalTemplateDocuments.Item(0).DocumentType = documentType

            ' 6. Add approval stages
            oApprovalTemplateStage = oApprovalTemplate.ApprovalTemplateStages.Add()
            oApprovalTemplateStage.ApprovalStageCode = approvalStageCode

            ' 7. Add terms including a user query
            oApprovalTemplate.UseTerms = SAPbobsCOM.BoYesNoEnum.tYES ' Enable terms

            ' Add a term based on a User Query
            Dim oApprovalTemplateQuery As SAPbobsCOM.ApprovalTemplateQuery = oApprovalTemplate.ApprovalTemplateQueries.Add()
            oApprovalTemplateQuery.QueryID = userQueryID
            ' 8. Add the approval template
            oApprovalTemplateParams = oApprovalTemplatesService.AddApprovalTemplate(oApprovalTemplate)

            ' Return the code of the newly created approval template
            result = oApprovalTemplateParams.Code

        Catch ex As Exception
            ' MessageBox.Show("Error creating approval template with query term: " & objMain.objCompany.GetLastErrorDescription())
            ' Handle the exception as needed
        Finally
            ' Release COM objects to prevent memory leaks
            If oApprovalTemplateParams IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oApprovalTemplateParams)
            If oApprovalTemplateTerm IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oApprovalTemplateTerm)
            If oApprovalTemplateStage IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oApprovalTemplateStage)
            If oApprovalTemplate IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oApprovalTemplate)
            If oApprovalTemplatesService IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oApprovalTemplatesService)
            If oCmpSrv IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oCmpSrv)
        End Try

        Return result
    End Function

    Public Function AddErrorLog(ByVal FormUID As String, ByVal DocNum As String, ByVal Method As String, ByVal Messg As String)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        ' Dim Donum As String = objMain.objUtilities.GetNextDocNum(objForm, objForm.UniqueID, "Primary")
        Dim userId As String = objMain.objCompany.UserName
        Dim formTitle As String = objForm.TypeEx

        Dim rand As New Random()
        Dim Namer As String = rand.Next(1, 100)
        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs.DoQuery("SELECT ""U_CMBENBL"" FROM ""@TNX_LOGST"" ORDER BY ""U_CMBENBL"" DESC LIMIT 1")
        Dim chkEnbl As String = oRs.Fields.Item(0).Value

        If chkEnbl = "1" Then
            'objMain.objUtilities.AddDataToLogError(Namer, DateTime.Now.ToString("yyyy/MM/dd"), DateTime.Now.ToString("HH:mm:ss"), DateTime.Now.ToString("HH:mm:ss"), Donum, userId, formTitle, "Error", "scope", Messg)
            Me.AddDataToLogError(Namer, DateTime.Now.ToString("yyyy/MM/dd"), DateTime.Now.ToString("HH:mm:ss"), DocNum, userId, formTitle, "Error", Method, Messg)
            Me.CreatAutoDeleteLogError()

        End If

    End Function
    Sub CreatAutoDeleteLogError()
        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs.DoQuery("SELECT ""U_CMBENBL"",""U_LODDAYS"" FROM ""@TNX_LOGST"" ORDER BY ""U_CMBENBL"" DESC LIMIT 1")
        Dim chkEnbl As Integer = oRs.Fields.Item(0).Value
        Dim lgdays As Integer = oRs.Fields.Item(1).Value

        If chkEnbl = 1 And lgdays > 0 Then
            Dim oRs1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim tales As String = "SELECT count(*) from ""@TNX_TALES"" where ""U_EntDate""<ADD_DAYS(CURRENT_DATE, (-1*" & lgdays & "))"
            oRs1.DoQuery(tales)
            Dim cntrd As Integer = oRs1.Fields.Item(0).Value
            If cntrd > 0 Then
                'Dim qrydlt As String = "DELETE from ""@TNX_TALES"" where ""U_EntDate""<ADD_DAYS(CURRENT_DATE, (-1*lgdays))"
                'Command.ExecuteQuery(qrydlt)
                Dim oRs2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRs2.DoQuery("DELETE from ""@TNX_TALES"" where ""U_EntDate""<ADD_DAYS(CURRENT_DATE, (-1*" & lgdays & "))")
            End If
        End If

    End Sub
    Public Function AddDataToLogError(ByVal Name As String, Optional ByVal EnterDate As Date = Nothing, Optional ByVal ErrorTime As DateTime = Nothing, Optional ByVal DocNUm As String = "", Optional ByVal Logid As String = "", Optional ByVal AddonName As String = "", Optional ByVal Level As String = "", Optional ByVal Scope As String = "", Optional ByVal Message As String = "")
        Try
            Dim oUserTable As SAPbobsCOM.UserTable
            Dim lReturn As Integer
            Dim ErrorString As String
            oUserTable = objMain.objCompany.UserTables.Item("TNX_TALES")
            Dim Code As String = Me.getMaxCode("@TNX_TALES")
            If oUserTable.GetByKey(Code) = False Then
                oUserTable.Code = Code
                oUserTable.Name = Code
                oUserTable.UserFields.Fields.Item("U_EntDate").Value = EnterDate
                oUserTable.UserFields.Fields.Item("U_ErrTime").Value = ErrorTime
                oUserTable.UserFields.Fields.Item("U_DocNum").Value = DocNUm
                oUserTable.UserFields.Fields.Item("U_USERID").Value = Logid
                oUserTable.UserFields.Fields.Item("U_AdName").Value = AddonName
                oUserTable.UserFields.Fields.Item("U_Level").Value = Level
                oUserTable.UserFields.Fields.Item("U_Adscop").Value = Scope
                oUserTable.UserFields.Fields.Item("U_MSGTXT").Value = Message

                lReturn = oUserTable.Add()
                If lReturn <> 0 Then
                    objMain.objCompany.GetLastError(lReturn, ErrorString)
                    Return (ErrorString)
                End If
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable)
            Return ("")
        Catch ex As Exception
            MessageBox.Show("Error in Add to log table:" & ex.Message)
        End Try
    End Function

    Public Sub ShowErrorMessage(ByVal oText As String)
        Message(oText, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    End Sub

    Private Sub Message(oText As String, smt_Error As SAPbouiCOM.BoStatusBarMessageType)
        Throw New NotImplementedException()
    End Sub

    Public Sub ShowWarningMessage(ByVal oText As String)
        Message(oText, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    End Sub
    Public Sub ShowSucessMessage(ByVal oText As String)
        Message(oText, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    End Sub
    Public Class ActivityRptObj
        Public name As String
        Public type As String
        Public value As List(Of List(Of String))
    End Class
    Public Class ImageViewModel
        Public BsImg As String
        Public DocNum As String
        Public CardName As String
    End Class


#Region "Messenger API"
    Public Function AttachmnetSelection(ByVal DocType As String, ByVal MsgType As String) As String
        Dim MsgQry As String = "Select ""U_AttEMLM"" From ""@MGRCONFIG"" Where ""U_Doc""='" + DocType + "' and ""U_Status""='Y'"
        Dim oRsMsg As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRsMsg.DoQuery(MsgQry)
        Dim Att As String = ""
        If oRsMsg.RecordCount > 0 Then
            If MsgType = "WhatsApp" Then
                Att = oRsMsg.Fields.Item("U_AttM").Value
            ElseIf MsgType = "Email" Then
                Att = oRsMsg.Fields.Item("U_AttEMLM").Value
            End If
        End If
        Return Att
    End Function

    Public Function GetAttachmnetPath() As String
        Dim MsgQry As String = "Select ""U_PDFP"" From ""OADM"""
        Dim oRsMsg As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRsMsg.DoQuery(MsgQry)
        Dim AttPath As String = ""
        If oRsMsg.RecordCount > 0 Then
            AttPath = oRsMsg.Fields.Item("U_PDFP").Value
        End If

        If Not String.IsNullOrEmpty(AttPath) Then

            Try
                ' 👉 Check if path exists in SERVER
                If Not Directory.Exists(AttPath) Then
                    Directory.CreateDirectory(AttPath)
                End If

            Catch ex As Exception
                objMain.objApplication.SetStatusBarMessage("Invalid/No access to path: " & ex.Message,
                                                       SAPbouiCOM.BoMessageTime.bmt_Short,
                                                       SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        Else
            objMain.objApplication.SetStatusBarMessage("Attachment path is empty in OADM",
                                                   SAPbouiCOM.BoMessageTime.bmt_Short,
                                                   SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If

        Return AttPath
    End Function
    Public Function GetCrystalReportIDday() As String
        Dim MsgQry As String = "Select ""U_PRCRID"" From ""OADM"""
        Dim oRsMsg As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRsMsg.DoQuery(MsgQry)
        Dim CRID As String = ""
        If oRsMsg.RecordCount > 0 Then
            CRID = oRsMsg.Fields.Item("U_PRCRID").Value
        End If
        Return CRID
    End Function
    Public Function GetCrystalReportIDValue() As String
        Dim MsgQry As String = "Select ""U_CSCRID"" From ""OADM"""
        Dim oRsMsg As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRsMsg.DoQuery(MsgQry)
        Dim CRID As String = ""
        If oRsMsg.RecordCount > 0 Then
            CRID = oRsMsg.Fields.Item("U_CSCRID").Value
        End If
        Return CRID
    End Function
    Public Function GetCrystalID(ByVal DocType As String, ByVal MsgType As String) As String
        Dim MsgQry As String = "Select ""U_ReportID"" From ""@MGRCONFIG"" Where ""U_Doc""='" + DocType + "' and ""U_Status""='Y'"
        Dim oRsMsg As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRsMsg.DoQuery(MsgQry)
        Dim CrystalID As String = ""
        If oRsMsg.RecordCount > 0 Then
            If MsgType = "WhatsApp" Then
                CrystalID = oRsMsg.Fields.Item("U_ReportID").Value
            ElseIf MsgType = "Email" Then
                CrystalID = oRsMsg.Fields.Item("U_ReportID").Value
            End If
        End If
        Return CrystalID
    End Function
    Public Function GetSignature(ByVal DocType As String, ByVal MsgType As String) As String
        Dim MsgQry As String = "SELECT IFNULL(""U_SIGN"",'') as ""U_SIGN"" From ""OADM"""
        Dim oRsMsg As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRsMsg.DoQuery(MsgQry)
        Dim EmailSign As String = ""
        If oRsMsg.RecordCount > 0 Then
            If MsgType = "WhatsApp" Then
                EmailSign = oRsMsg.Fields.Item("U_SIGN").Value
            ElseIf MsgType = "Email" Then
                EmailSign = oRsMsg.Fields.Item("U_SIGN").Value
            End If
        Else
            Dim MsgQry1 As String = "Select ""U_SignEML"" From ""@MGRCONFIG"" Where ""U_Doc""='" + DocType + "' and ""U_Status""='Y'"
            Dim oRsMsg1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsMsg1.DoQuery(MsgQry1)
            If oRsMsg1.RecordCount > 0 Then
                If MsgType = "WhatsApp" Then
                    EmailSign = oRsMsg1.Fields.Item("U_SignEML").Value
                ElseIf MsgType = "Email" Then
                    EmailSign = oRsMsg1.Fields.Item("U_SignEML").Value
                End If
            End If
        End If
        Return EmailSign
    End Function
    Public Function GetSignaturePCR(ByVal DocType As String, ByVal MsgType As String) As String

        Dim EmailSign As String = ""

        Dim MsgQry1 As String = "Select ""U_SignEML"" From ""@MGRCONFIG"" Where ""U_Doc""='" + DocType + "' and ""U_Status""='Y'"
            Dim oRsMsg1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsMsg1.DoQuery(MsgQry1)
            If oRsMsg1.RecordCount > 0 Then
                If MsgType = "WhatsApp" Then
                    EmailSign = oRsMsg1.Fields.Item("U_SignEML").Value
                ElseIf MsgType = "Email" Then
                    EmailSign = oRsMsg1.Fields.Item("U_SignEML").Value
                End If
            End If

        Return EmailSign
    End Function
    Public Function MessageTrigger(ByVal DocType As String, ByVal MsgType As String) As String
        Dim MsgQry As String = "Select ""U_Manual"",""U_ManualS"",""U_ManualE"" From ""@MGRCONFIG"" Where ""U_Doc""='" + DocType + "' and ""U_Status""='Y'"
        Dim oRsMsg As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRsMsg.DoQuery(MsgQry)
        Dim Manual As String = ""
        If oRsMsg.RecordCount > 0 Then
            If MsgType = "WhatsApp" Then
                Manual = oRsMsg.Fields.Item("U_Manual").Value
            ElseIf MsgType = "SMS" Then
                Manual = oRsMsg.Fields.Item("U_ManualS").Value
            ElseIf MsgType = "Email" Then
                Manual = oRsMsg.Fields.Item("U_ManualE").Value
            End If
        End If
        Return Manual
    End Function
    Public Function DocTrigger(ByVal DocTypeField As String) As String
        Try
            Dim ConfigQry As String = "Select " & sISNULL & "(" & DocTypeField & ",'Y') From ""@CONFIG"""
            Dim oRsConfig As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsConfig.DoQuery(ConfigQry)
            If oRsConfig.RecordCount > 0 Then
                Return oRsConfig.Fields.Item(0).Value
            End If
        Catch ex As Exception
            ShowErrorMessage(ex.Message)
        End Try

    End Function
    Public Function DocTriggerSMS(ByVal DocTypeField As String) As String
        Try
            Dim ConfigQry As String = "Select " & sISNULL & "(" & DocTypeField & ",'Y') From ""@CONFIG"""
            Dim oRsConfig As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsConfig.DoQuery(ConfigQry)
            If oRsConfig.RecordCount > 0 Then
                Return oRsConfig.Fields.Item(0).Value
            End If
        Catch ex As Exception
            ShowErrorMessage(ex.Message)
        End Try

    End Function
    Public Function DocTriggerEMAIL(ByVal DocTypeFieldEmail As String) As String
        Try
            Dim ConfigQry As String = "Select " & sISNULL & "(" & DocTypeFieldEmail & ",'Y') From ""@CONFIG"""
            Dim oRsConfig As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsConfig.DoQuery(ConfigQry)
            If oRsConfig.RecordCount > 0 Then
                Return oRsConfig.Fields.Item(0).Value
            End If
        Catch ex As Exception
            ShowErrorMessage(ex.Message)
        End Try

    End Function
    Public Sub Document_Update(ByVal Documnt As String, ByVal DocumntNo As Integer)
        Try
            Dim DocEntry As Integer = 0
            Dim DocEQry As String = ""
            Dim oRsDocE As SAPbobsCOM.Recordset
            Dim oDocuments As SAPbobsCOM.Documents
            Dim oOrdersLines As SAPbobsCOM.Document_Lines
            Dim odtLines As New Data.DataTable
            Dim intRetCode As Integer = -1
            Dim strErrMsg As String = ""

            If Documnt = "SO" Then
                DocEQry = "Select ""DocEntry"" From ""ORDR"" Where ""DocNum""='" + DocumntNo.ToString() + "' "
                oRsDocE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsDocE.DoQuery(DocEQry)
                DocEntry = Convert.ToInt32(oRsDocE.Fields.Item("DocEntry").Value)
                oDocuments = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                oDocuments.GetByKey(DocEntry)
            ElseIf Documnt = "PO" Then
                DocEQry = "Select ""DocEntry"" From ""OPOR"" Where ""DocNum""='" + DocumntNo.ToString() + "' "
                oRsDocE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsDocE.DoQuery(DocEQry)
                DocEntry = Convert.ToInt32(oRsDocE.Fields.Item("DocEntry").Value)
                oDocuments = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                oDocuments.GetByKey(DocEntry)
            ElseIf Documnt = "SQ" Then
                DocEQry = "Select ""DocEntry"" From ""OQUT"" Where ""DocNum""='" + DocumntNo.ToString() + "' "
                oRsDocE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsDocE.DoQuery(DocEQry)
                DocEntry = Convert.ToInt32(oRsDocE.Fields.Item("DocEntry").Value)
                oDocuments = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations)
                oDocuments.GetByKey(DocEntry)
            ElseIf Documnt = "DL" Then
                DocEQry = "Select ""DocEntry"" From ""ODLN"" Where ""DocNum""='" + DocumntNo.ToString() + "' "
                oRsDocE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsDocE.DoQuery(DocEQry)
                DocEntry = Convert.ToInt32(oRsDocE.Fields.Item("DocEntry").Value)
                oDocuments = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                oDocuments.GetByKey(DocEntry)
            ElseIf Documnt = "ARINV" Then
                DocEQry = "Select ""DocEntry"" From ""OINV"" Where ""DocNum""='" + DocumntNo.ToString() + "' "
                oRsDocE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsDocE.DoQuery(DocEQry)
                DocEntry = Convert.ToInt32(oRsDocE.Fields.Item("DocEntry").Value)
                oDocuments = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                oDocuments.GetByKey(DocEntry)
            ElseIf Documnt = "ARRINV" Then
                DocEQry = "Select ""DocEntry"" From ""OINV"" Where ""DocNum""='" + DocumntNo.ToString() + "' "
                oRsDocE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsDocE.DoQuery(DocEQry)
                DocEntry = Convert.ToInt32(oRsDocE.Fields.Item("DocEntry").Value)
                oDocuments = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                oDocuments.GetByKey(DocEntry)
            ElseIf Documnt = "ARC" Then
                DocEQry = "Select ""DocEntry"" From ""ORIN"" Where ""DocNum""='" + DocumntNo.ToString() + "' "
                oRsDocE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsDocE.DoQuery(DocEQry)
                DocEntry = Convert.ToInt32(oRsDocE.Fields.Item("DocEntry").Value)
                oDocuments = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                oDocuments.GetByKey(DocEntry)
            ElseIf Documnt = "ARRET" Then
                DocEQry = "Select ""DocEntry"" From ""OINV"" Where ""DocNum""='" + DocumntNo.ToString() + "' "
                oRsDocE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsDocE.DoQuery(DocEQry)
                DocEntry = Convert.ToInt32(oRsDocE.Fields.Item("DocEntry").Value)
                oDocuments = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCorrectionInvoiceReversal)
                oDocuments.GetByKey(DocEntry)

            ElseIf Documnt = "PQ" Then
                DocEQry = "Select ""DocEntry"" From ""OPQT"" Where ""DocNum""='" + DocumntNo.ToString() + "' "
                oRsDocE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsDocE.DoQuery(DocEQry)
                DocEntry = Convert.ToInt32(oRsDocE.Fields.Item("DocEntry").Value)
                oDocuments = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseQuotations)
                oDocuments.GetByKey(DocEntry)
            ElseIf Documnt = "APINV" Then
                DocEQry = "Select ""DocEntry"" From ""OPCH"" Where ""DocNum""='" + DocumntNo.ToString() + "' "
                oRsDocE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsDocE.DoQuery(DocEQry)
                DocEntry = Convert.ToInt32(oRsDocE.Fields.Item("DocEntry").Value)
                oDocuments = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                oDocuments.GetByKey(DocEntry)
            ElseIf Documnt = "GRPO" Then
                DocEQry = "Select ""DocEntry"" From ""OPDN"" Where ""DocNum""='" + DocumntNo.ToString() + "' "
                oRsDocE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsDocE.DoQuery(DocEQry)
                DocEntry = Convert.ToInt32(oRsDocE.Fields.Item("DocEntry").Value)
                oDocuments = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                oDocuments.GetByKey(DocEntry)
            ElseIf Documnt = "GR" Then
                DocEQry = "Select ""DocEntry"" From ""ORPD"" Where ""DocNum""='" + DocumntNo.ToString() + "' "
                oRsDocE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsDocE.DoQuery(DocEQry)
                DocEntry = Convert.ToInt32(oRsDocE.Fields.Item("DocEntry").Value)
                oDocuments = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseReturns)
                oDocuments.GetByKey(DocEntry)
            ElseIf Documnt = "APC" Then
                DocEQry = "Select ""DocEntry"" From ""ORPC"" Where ""DocNum""='" + DocumntNo.ToString() + "' "
                oRsDocE = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsDocE.DoQuery(DocEQry)
                DocEntry = Convert.ToInt32(oRsDocE.Fields.Item("DocEntry").Value)
                oDocuments = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                oDocuments.GetByKey(DocEntry)
            End If
            Try
                oDocuments.UserFields.Fields.Item("U_WFLG").Value = "YES"
            Catch ex As Exception
                ShowErrorMessage(ex.Message)
            End Try
            oDocuments.Update()

        Catch ex As Exception
            ShowErrorMessage(ex.Message)
        End Try
    End Sub
    Public Function getMessage(ByVal DocType As String, ByVal MessageType As String) As String
        Dim MsgQry As String = "Select ""U_Msg"",""U_MsgSMS"",""U_MsgEML"" From ""@MGRCONFIG"" Where ""U_Doc""='" + DocType + "' and ""U_Status""='Y'"
        Dim oRsMsg As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRsMsg.DoQuery(MsgQry)
        Dim Msg As String = ""

        If oRsMsg.RecordCount > 0 Then
            If MessageType = "WhatsApp" Then
                Msg = oRsMsg.Fields.Item("U_Msg").Value
            ElseIf MessageType = "SMS" Then
                Msg = oRsMsg.Fields.Item("U_MsgSMS").Value
            ElseIf MessageType = "Email" Then
                Msg = oRsMsg.Fields.Item("U_MsgEML").Value
            End If
        End If
        Return Msg
    End Function
    Public Function GetTemplateName(ByVal DocType As String, ByVal MessageType As String) As String
        Dim TemplateQry As String = "Select ""U_MgsTemp"" From ""@MGRCONFIG"" Where ""U_Doc""='" + DocType + "' and ""U_Status""='Y'"
        Dim oRsTemplate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRsTemplate.DoQuery(TemplateQry)
        Dim Template As String = ""

        If oRsTemplate.RecordCount > 0 Then
            If MessageType = "WhatsApp" Then
                Template = oRsTemplate.Fields.Item("U_MgsTemp").Value
            End If
        End If
        Return Template
    End Function
    Public Function CheckAttachment(ByVal DocType As String, ByVal MsgType As String) As Boolean
        Dim Attachement As Boolean = False
        Dim MsgQry As String = "Select ""U_Att"" From ""@MGRCONFIG"" Where ""U_Doc"" = '" + DocType + "' And IFNULL(""U_Status"",'N') = 'Y' And IFNULL(""U_Att"",'N') = 'Y'"
        Dim oRsMsg As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRsMsg.DoQuery(MsgQry)
        Dim Manual As String = ""
        If oRsMsg.RecordCount > 0 Then
            If MsgType = "WhatsApp" Then
                Attachement = True
            End If
        End If
        Return Attachement
    End Function
    Public Function CheckAttachment_EMAIL(ByVal DocType As String, ByVal MsgType As String) As Boolean
        Dim Attachement As Boolean = False
        Dim MsgQry As String = "Select ""U_AttEML"" From ""@MGRCONFIG"" Where ""U_Doc"" = '" + DocType + "' And IFNULL(""U_Status"",'N') = 'Y' And IFNULL(""U_AttEML"",'N') = 'Y'"
        Dim oRsMsg As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRsMsg.DoQuery(MsgQry)
        Dim Manual As String = ""
        If oRsMsg.RecordCount > 0 Then
            If MsgType = "Email" Then
                Attachement = True
            End If
        End If
        Return Attachement
    End Function
    Public Function AddDataToLogTable(ByVal Name As String, Optional ByVal EmailDate As Date = Nothing, Optional ByVal CardCode As String = "", Optional ByVal Status As String = "", Optional ByVal Message As String = "", Optional ByVal Email As String = "", Optional ByVal Mobile As String = "", Optional ByVal DocNum As String = "", Optional ByVal DocEntry As String = "", Optional ByVal Response As String = "", Optional ByVal MsgType As String = "")
        Try
            Dim oUserTable As SAPbobsCOM.UserTable
            Dim lReturn As Integer
            Dim ErrorString As String
            oUserTable = objMain.objCompany.UserTables.Item("SBO_MSGLOG")
            Dim Code As String = Me.getMaxCode("@SBO_MSGLOG")
            If oUserTable.GetByKey(Code) = False Then
                oUserTable.Code = Code
                oUserTable.Name = Code

                oUserTable.UserFields.Fields.Item("U_DocNum").Value = DocNum
                oUserTable.UserFields.Fields.Item("U_DocEntry").Value = DocEntry
                oUserTable.UserFields.Fields.Item("U_MsgType").Value = MsgType
                oUserTable.UserFields.Fields.Item("U_DocType").Value = Name
                oUserTable.UserFields.Fields.Item("U_Date").Value = EmailDate
                oUserTable.UserFields.Fields.Item("U_CardCode").Value = CardCode
                oUserTable.UserFields.Fields.Item("U_Status").Value = Status
                oUserTable.UserFields.Fields.Item("U_MSG").Value = Message
                oUserTable.UserFields.Fields.Item("U_EmailID").Value = Email
                oUserTable.UserFields.Fields.Item("U_Mobile").Value = Mobile
                oUserTable.UserFields.Fields.Item("U_Response").Value = Response
                oUserTable.UserFields.Fields.Item("U_ResMessage").Value = Response

                lReturn = oUserTable.Add()
                If lReturn <> 0 Then
                    objMain.objCompany.GetLastError(lReturn, ErrorString)
                    Return (ErrorString)
                End If
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable)
            Return ("")
        Catch ex As Exception
            ShowErrorMessage(ex.Message)
            objMain.objUtilities.Write_Log("Error in Add to log table:" & ex.Message)
        End Try
    End Function
    Public Function UpdateDataToLogTable(ByVal DocType As String, ByVal DocEntry As String, ByVal Flag As String, ByVal MsgType As String, ByVal Response As String)
        Try
            Dim sCode As String
            Dim strSQL As String = "SELECT TOP 1 ""Code"" FROM ""@SBO_MSGLOG"" WHERE ""U_DocType""='" + DocType + "' AND ""U_DocEntry""='" + DocEntry + "' AND ""U_MsgType""='" + MsgType + "'"
            Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery(strSQL)
            If oRS.RecordCount > 0 Then
                sCode = oRS.Fields.Item(0).Value.ToString().TrimEnd()
                Dim oUserTable As SAPbobsCOM.UserTable
                Dim lReturn As Integer
                Dim ErrorString As String
                oUserTable = objMain.objCompany.UserTables.Item("SBO_MSGLOG")
                If oUserTable.GetByKey(sCode) = True Then
                    If Flag = "Success" Then
                        oUserTable.UserFields.Fields.Item("U_Status").Value = "Success"
                    Else
                        oUserTable.UserFields.Fields.Item("U_Status").Value = "Failed"
                        oUserTable.UserFields.Fields.Item("U_Response").Value = Response
                    End If
                    lReturn = oUserTable.Update()
                    If lReturn <> 0 Then
                        objMain.objCompany.GetLastError(lReturn, ErrorString)
                        Return (ErrorString)
                    End If
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable)
                Return ("")
            End If
        Catch ex As Exception
            ShowErrorMessage(ex.Message)
            objMain.objUtilities.Write_Log("Error in Add to log table:" & ex.Message)
        End Try
    End Function
    Public Function getMaxCode1(ByVal sTable As String) As String
        Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim MaxCode As Integer
        Dim sCode As String
        Dim strSQL As String
        Try
            strSQL = "SELECT MAX(" & sISNULL & "(""Code"",0))+1 FROM """ & sTable & """"
            oRS.DoQuery(strSQL)
            If oRS.RecordCount > 0 Then
                MaxCode = oRS.Fields.Item(0).Value
            Else
                MaxCode = 1
            End If
            sCode = MaxCode
            Return sCode
        Catch ex As Exception
            ShowErrorMessage(ex.Message)
            objMain.objUtilities.Write_Log("Error in Next Code:" & ex.Message)
        Finally
            oRS = Nothing
        End Try
    End Function
    Sub AddToFileIN(ByVal contents As String)
        Try
            Using TextWriter As StreamWriter = New StreamWriter(System.Configuration.ConfigurationSettings.AppSettings("LogFile") & "LogFile_IN" & Today.ToString("dd-MM-yyyy") & ".txt", True)
                TextWriter.WriteLine(contents)
            End Using
        Catch ex As Exception
            ShowErrorMessage(ex.Message)
            objMain.objUtilities.Write_Log("AddToFile:" + ex.Message + System.DateTime.Now())
        End Try
    End Sub
    Function getAttachment(ByVal DocType As String, ByVal DocEntry As String) As String
        Dim PDFPath As String
        Try
            Dim Qry As String = ""
            If DocType = "SO" Then
                Qry = "SELECT T2.""trgtPath"" ||'\'|| T2.""FileName""||'.'|| T2.""FileExt""  FROM ""ORDR""  T0 " &
                " INNER JOIN ""OATC"" T1 ON T0.""AtcEntry""=T1.""AbsEntry"" " &
                " INNER JOIN ""ATC1"" T2 ON T2.""AbsEntry""=T1.""AbsEntry"" Where ""DocEntry""='" & DocEntry & "' " &
                " And T2.""FileExt""='pdf' Order By T2.""Line"" DESC "
            ElseIf DocType = "ARINV" Then
                Qry = "SELECT T2.""trgtPath"" ||'\'|| T2.""FileName""||'.'|| T2.""FileExt""  FROM ""OINV""  T0 " &
                " INNER JOIN ""OATC"" T1 ON T0.""AtcEntry""=T1.""AbsEntry"" " &
                " INNER JOIN ""ATC1"" T2 ON T2.""AbsEntry""=T1.""AbsEntry"" Where ""DocEntry""='" & DocEntry & "' " &
                " And T2.""FileExt""='pdf' Order By T2.""Line"" DESC "
            ElseIf DocType = "ARRINV" Then
                Qry = "SELECT T2.""trgtPath"" ||'\'|| T2.""FileName""||'.'|| T2.""FileExt""  FROM ""OINV""  T0 " &
                " INNER JOIN ""OATC"" T1 ON T0.""AtcEntry""=T1.""AbsEntry"" " &
                " INNER JOIN ""ATC1"" T2 ON T2.""AbsEntry""=T1.""AbsEntry"" Where ""DocEntry""='" & DocEntry & "' " &
                " And T2.""FileExt""='pdf' Order By T2.""Line"" DESC "
            ElseIf DocType = "ARC" Then
                Qry = "SELECT T2.""trgtPath"" ||'\'|| T2.""FileName""||'.'|| T2.""FileExt""  FROM ""ORIN""  T0 " &
                " INNER JOIN ""OATC"" T1 ON T0.""AtcEntry""=T1.""AbsEntry"" " &
                " INNER JOIN ""ATC1"" T2 ON T2.""AbsEntry""=T1.""AbsEntry"" Where ""DocEntry""='" & DocEntry & "' " &
                " And T2.""FileExt""='pdf' Order By T2.""Line"" DESC "
            ElseIf DocType = "DL" Then
                Qry = "SELECT T2.""trgtPath"" ||'\'|| T2.""FileName""||'.'|| T2.""FileExt""  FROM ""ODLN""  T0 " &
                " INNER JOIN ""OATC"" T1 ON T0.""AtcEntry""=T1.""AbsEntry"" " &
                " INNER JOIN ""ATC1"" T2 ON T2.""AbsEntry""=T1.""AbsEntry"" Where ""DocEntry""='" & DocEntry & "' " &
                " And T2.""FileExt""='pdf' Order By T2.""Line"" DESC "
            ElseIf DocType = "ARRET" Then
                Qry = "SELECT T2.""trgtPath"" ||'\'|| T2.""FileName""||'.'|| T2.""FileExt""  FROM ""ORDN""  T0 " &
                " INNER JOIN ""OATC"" T1 ON T0.""AtcEntry""=T1.""AbsEntry"" " &
                " INNER JOIN ""ATC1"" T2 ON T2.""AbsEntry""=T1.""AbsEntry"" Where ""DocEntry""='" & DocEntry & "' " &
                " And T2.""FileExt""='pdf' Order By T2.""Line"" DESC "
            ElseIf DocType = "SQ" Then
                Qry = "SELECT T2.""trgtPath"" ||'\'|| T2.""FileName""||'.'|| T2.""FileExt""  FROM ""OQUT""  T0 " &
                " INNER JOIN ""OATC"" T1 ON T0.""AtcEntry""=T1.""AbsEntry"" " &
                " INNER JOIN ""ATC1"" T2 ON T2.""AbsEntry""=T1.""AbsEntry"" Where ""DocEntry""='" & DocEntry & "' " &
                " And T2.""FileExt""='pdf' Order By T2.""Line"" DESC "
            ElseIf DocType = "SBA" Then
                Qry = "SELECT T2.""trgtPath"" ||'\'|| T2.""FileName""||'.'|| T2.""FileExt""  FROM ""OOAT""  T0 " &
                " INNER JOIN ""OATC"" T1 ON T0.""AtchEntry""=T1.""AbsEntry"" " &
                " INNER JOIN ""ATC1"" T2 ON T2.""AbsEntry""=T1.""AbsEntry"" Where ""AbsID""='" & DocEntry & "' " &
                " And T2.""FileExt""='pdf' Order By T2.""Line"" DESC "
                '------
            ElseIf DocType = "PO" Then
                Qry = "SELECT T2.""trgtPath"" ||'\'|| T2.""FileName""||'.'|| T2.""FileExt""  FROM ""OPOR""  T0 " &
                " INNER JOIN ""OATC"" T1 ON T0.""AtcEntry""=T1.""AbsEntry"" " &
                " INNER JOIN ""ATC1"" T2 ON T2.""AbsEntry""=T1.""AbsEntry"" Where ""DocEntry""='" & DocEntry & "' " &
                " And T2.""FileExt""='pdf' Order By T2.""Line"" DESC "
            ElseIf DocType = "PQ" Then
                Qry = "SELECT T2.""trgtPath"" ||'\'|| T2.""FileName""||'.'|| T2.""FileExt""  FROM ""OPQT""  T0 " &
                " INNER JOIN ""OATC"" T1 ON T0.""AtcEntry""=T1.""AbsEntry"" " &
                " INNER JOIN ""ATC1"" T2 ON T2.""AbsEntry""=T1.""AbsEntry"" Where ""DocEntry""='" & DocEntry & "' " &
                " And T2.""FileExt""='pdf' Order By T2.""Line"" DESC "
            ElseIf DocType = "APINV" Then
                Qry = "SELECT T2.""trgtPath"" ||'\'|| T2.""FileName""||'.'|| T2.""FileExt""  FROM ""OPCH""  T0 " &
                " INNER JOIN ""OATC"" T1 ON T0.""AtchEntry""=T1.""AbsEntry"" " &
                " INNER JOIN ""ATC1"" T2 ON T2.""AbsEntry""=T1.""AbsEntry"" Where ""AbsID""='" & DocEntry & "' " &
                " And T2.""FileExt""='pdf' Order By T2.""Line"" DESC "
            ElseIf DocType = "GR" Then
                Qry = "SELECT T2.""trgtPath"" ||'\'|| T2.""FileName""||'.'|| T2.""FileExt""  FROM ""ORPD""  T0 " &
                " INNER JOIN ""OATC"" T1 ON T0.""AtcEntry""=T1.""AbsEntry"" " &
                " INNER JOIN ""ATC1"" T2 ON T2.""AbsEntry""=T1.""AbsEntry"" Where ""DocEntry""='" & DocEntry & "' " &
                " And T2.""FileExt""='pdf' Order By T2.""Line"" DESC "
            ElseIf DocType = "GRPO" Then
                Qry = "SELECT T2.""trgtPath"" ||'\'|| T2.""FileName""||'.'|| T2.""FileExt""  FROM ""OPDN""  T0 " &
                " INNER JOIN ""OATC"" T1 ON T0.""AtchEntry""=T1.""AbsEntry"" " &
                " INNER JOIN ""ATC1"" T2 ON T2.""AbsEntry""=T1.""AbsEntry"" Where ""AbsID""='" & DocEntry & "' " &
                " And T2.""FileExt""='pdf' Order By T2.""Line"" DESC "
            ElseIf DocType = "APC" Then
                Qry = "SELECT T2.""trgtPath"" ||'\'|| T2.""FileName""||'.'|| T2.""FileExt""  FROM ""OPRC""  T0 " &
                " INNER JOIN ""OATC"" T1 ON T0.""AtchEntry""=T1.""AbsEntry"" " &
                " INNER JOIN ""ATC1"" T2 ON T2.""AbsEntry""=T1.""AbsEntry"" Where ""AbsID""='" & DocEntry & "' " &
                " And T2.""FileExt""='pdf' Order By T2.""Line"" DESC "
            ElseIf DocType = "PBA" Then
                Qry = "SELECT T2.""trgtPath"" ||'\'|| T2.""FileName""||'.'|| T2.""FileExt""  FROM ""OOAT""  T0 " &
                " INNER JOIN ""OATC"" T1 ON T0.""AtchEntry""=T1.""AbsEntry"" " &
                " INNER JOIN ""ATC1"" T2 ON T2.""AbsEntry""=T1.""AbsEntry"" Where ""AbsID""='" & DocEntry & "' " &
                " And T2.""FileExt""='pdf' Order By T2.""Line"" DESC "
            ElseIf DocType = "IP" Then
                Qry = "SELECT T2.""trgtPath"" ||'\'|| T2.""FileName""||'.'|| T2.""FileExt""  FROM ""ORCT""  T0 " &
                " INNER JOIN ""OATC"" T1 ON T0.""AtcEntry""=T1.""AbsEntry"" " &
                " INNER JOIN ""ATC1"" T2 ON T2.""AbsEntry""=T1.""AbsEntry"" Where ""DocEntry""='" & DocEntry & "' " &
                " And T2.""FileExt""='pdf' Order By T2.""Line"" DESC "
            ElseIf DocType = "OP" Then
                Qry = "SELECT T2.""trgtPath"" ||'\'|| T2.""FileName""||'.'|| T2.""FileExt""  FROM ""OVPM""  T0 " &
                " INNER JOIN ""OATC"" T1 ON T0.""AtcEntry""=T1.""AbsEntry"" " &
                " INNER JOIN ""ATC1"" T2 ON T2.""AbsEntry""=T1.""AbsEntry"" Where ""DocEntry""='" & DocEntry & "' " &
                " And T2.""FileExt""='pdf' Order By T2.""Line"" DESC "
            ElseIf DocType = "GRP" Then
                Qry = "SELECT T2.""trgtPath"" ||'\'|| T2.""FileName""||'.'|| T2.""FileExt""  FROM ""OIGN""  T0 " &
                " INNER JOIN ""OATC"" T1 ON T0.""AtcEntry""=T1.""AbsEntry"" " &
                " INNER JOIN ""ATC1"" T2 ON T2.""AbsEntry""=T1.""AbsEntry"" Where ""DocEntry""='" & DocEntry & "' " &
                " And T2.""FileExt""='pdf' Order By T2.""Line"" DESC "
            ElseIf DocType = "GI" Then
                Qry = "SELECT T2.""trgtPath"" ||'\'|| T2.""FileName""||'.'|| T2.""FileExt""  FROM ""OIGE""  T0 " &
                " INNER JOIN ""OATC"" T1 ON T0.""AtcEntry""=T1.""AbsEntry"" " &
                " INNER JOIN ""ATC1"" T2 ON T2.""AbsEntry""=T1.""AbsEntry"" Where ""DocEntry""='" & DocEntry & "' " &
                " And T2.""FileExt""='pdf' Order By T2.""Line"" DESC "
            ElseIf DocType = "ITR" Then
                Qry = "SELECT T2.""trgtPath"" ||'\'|| T2.""FileName""||'.'|| T2.""FileExt""  FROM ""OWTQ""  T0 " &
                " INNER JOIN ""OATC"" T1 ON T0.""AtcEntry""=T1.""AbsEntry"" " &
                " INNER JOIN ""ATC1"" T2 ON T2.""AbsEntry""=T1.""AbsEntry"" Where ""DocEntry""='" & DocEntry & "' " &
                " And T2.""FileExt""='pdf' Order By T2.""Line"" DESC "
            ElseIf DocType = "IT" Then
                Qry = "SELECT T2.""trgtPath"" ||'\'|| T2.""FileName""||'.'|| T2.""FileExt""  FROM ""OWTR""  T0 " &
                " INNER JOIN ""OATC"" T1 ON T0.""AtcEntry""=T1.""AbsEntry"" " &
                " INNER JOIN ""ATC1"" T2 ON T2.""AbsEntry""=T1.""AbsEntry"" Where ""DocEntry""='" & DocEntry & "' " &
                " And T2.""FileExt""='pdf' Order By T2.""Line"" DESC "
            ElseIf DocType = "JE" Then
                Qry = "SELECT T2.""trgtPath"" ||'\'|| T2.""FileName""||'.'|| T2.""FileExt""  FROM ""OJDT""  T0 " &
                " INNER JOIN ""OATC"" T1 ON T0.""AtcEntry""=T1.""AbsEntry"" " &
                " INNER JOIN ""ATC1"" T2 ON T2.""AbsEntry""=T1.""AbsEntry"" Where  T0.""TransId""='" & DocEntry & "' " &
                " And T2.""FileExt""='pdf' Order By T2.""Line"" DESC "
            End If
            Dim rsPath As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsPath.DoQuery(Qry)

            If rsPath.RecordCount > 0 Then
                If rsPath.Fields.Item(0).Value <> "" Then
                    PDFPath = rsPath.Fields.Item(0).Value
                Else
                    objMain.objApplication.StatusBar.SetText("Please configure the attachemnt path in User defaults and Print preferences", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            Else
                objMain.objApplication.StatusBar.SetText("Please configure the attachemnt path in User defaults and Print preferences", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            Return PDFPath
        Catch ex As Exception
            ShowErrorMessage(ex.Message)
        End Try
    End Function


    Public Async Function GetManualAttachment(gmDocEntry As String, rptDoc As String, printDoc As String) As Task

        Try
            Dim gwBaseUrl = ConfigurationManager.AppSettings("GatewayAPIUrl")
            Dim securedwebserviceurl = gwBaseUrl.ToString() & "/login"
            ServicePointManager.Expect100Continue = True
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11
            ' Skip validation of SSL/TLS certificate
            ServicePointManager.ServerCertificateValidationCallback = Function() True

            Dim httpWebRequest = CType(WebRequest.Create(securedwebserviceurl), HttpWebRequest)
            httpWebRequest.ContentType = "application/json"
            httpWebRequest.Method = "POST"
            '-----------------------
            Dim ci = New RSDBObj With {
                                        .CompanyDB = ConfigurationManager.AppSettings("CompanyDB"),
                                        .UserName = ConfigurationManager.AppSettings("UserName"),
                                        .DBInstance = ConfigurationManager.AppSettings("DBInstance"),
                                        .Password = ConfigurationManager.AppSettings("Password")
                                        }
            Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(ci)
            '-----------------------

            Using streamWriter = New StreamWriter(httpWebRequest.GetRequestStream())

                streamWriter.Write(json)
                streamWriter.Flush()
                streamWriter.Close()
            End Using

            httpWebRequest.ProtocolVersion = HttpVersion.Version10
            Dim httpResponse = CType(httpWebRequest.GetResponse(), HttpWebResponse)
            Dim responseFromServer As String = ""
            Using streamReader = New StreamReader(httpResponse.GetResponseStream())
                responseFromServer = streamReader.ReadToEnd()
            End Using

            For i As Integer = 0 To httpResponse.Headers.Count - 1
                Dim n As String = httpResponse.Headers.GetKey(i)
                If Not Equals(n, "Set-Cookie") Then Continue For

                Dim value As String = httpResponse.Headers.[Get](i)
                Dim ValueB As String = value
                Me.getBS64(value, gmDocEntry, rptDoc, printDoc)

            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
    Private Async Function getBS64(cc As String, gmDocEntry As String, rptDoc As String, printDoc As String) As Task
        Dim lstRF As List(Of ActivityRptObj) = New List(Of ActivityRptObj)()
        Dim pdfObj As ImageViewModel = New ImageViewModel()

        Dim str3 As List(Of String) = New List(Of String)()
        Dim ss3 As List(Of List(Of String)) = New List(Of List(Of String))()
        Try

            str3.Add(gmDocEntry)
            ss3.Add(str3)
            Dim keyName = ""
            If Equals(printDoc, "STR1") OrElse Equals(printDoc, "STR2") Then
                keyName = "@num"
            Else
                keyName = "DocKey@"
            End If
            Dim m3 As ActivityRptObj = New ActivityRptObj() With {
                                                                   .name = keyName,
                                                                   .type = "xsd:decimal"
                                                                 }
            m3.value = ss3
            lstRF.Add(m3)
            Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(lstRF)
            Dim gwBaseUrl = ConfigurationManager.AppSettings("GatewayAPIUrl")
            Dim securedwebserviceurl = gwBaseUrl.ToString() & "/rs/v1/ExportPDFData?DocCode=".ToString() & rptDoc
            ServicePointManager.Expect100Continue = True
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11
            ' Skip validation of SSL/TLS certificate
            ServicePointManager.ServerCertificateValidationCallback = Function() True

            Dim container As CookieContainer = New CookieContainer()
            'container.Add(new Uri("https://10.12.80.213:60000/rs/v1/ExportPDFData?DocCode="+rptDoc), new Cookie("SESSION", "value"));
            For Each singleCookie In cc.Split("="c)
                Dim match As Match = Regex.Match(singleCookie, "(.+?)=(.+?);")
                Dim match2 As Match = Regex.Match(singleCookie, "=(.+?);")
                Dim k = match2.Groups(1).Value

                'container.Add(new Uri(gwBaseUrl+"/rs/v1/ExportPDFData?DocCode=" + rptDoc), new Cookie("SESSION", k));
            Next

            Dim keyStrList1 = cc.Split("="c)
            Dim tokenKey1 = keyStrList1(1)

            Dim keyStrList2 = tokenKey1.Split(";"c)
            Dim tokenKey2 = keyStrList2(0)



            container.Add(New Uri(gwBaseUrl.ToString() & "/rs/v1/ExportPDFData?DocCode=".ToString() & rptDoc), New Cookie("Session", tokenKey2))

            Dim httpWebRequest = CType(WebRequest.Create(securedwebserviceurl), HttpWebRequest)

            httpWebRequest.CookieContainer = container

            httpWebRequest.Headers.Add(HttpRequestHeader.Cookie, "Session=" & tokenKey2)


            httpWebRequest.ContentType = "application/json"
            httpWebRequest.Accept = "application/json"
            httpWebRequest.Method = "POST"



            Using streamWriter = New StreamWriter(httpWebRequest.GetRequestStream())
                ' string json = "{\"CompanyDB\":\"ASC_LM\"," +
                '"\"UserName\":\"manager\", \"Password\":\"1234\"}";
                streamWriter.Write(json)
                streamWriter.Flush()
                streamWriter.Close()
            End Using



            httpWebRequest.ProtocolVersion = HttpVersion.Version10
            Dim httpResponse = CType(httpWebRequest.GetResponse(), HttpWebResponse)
            Using streamReader = New StreamReader(httpResponse.GetResponseStream())
                Dim responseFromServer As String = streamReader.ReadToEnd()
                pdfObj.BsImg = responseFromServer
                Me.SaveBS64toPDF(pdfObj, printDoc, gmDocEntry)
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Function


    Function getAttachment_up(ByVal DocType As String, ByVal DocEntry As String) As String
        Dim PDFPath As String
        Try
            Dim Qry As String = ""
            If DocType = "SO" Then
                Qry = "SELECT T2.""trgtPath"" ||'\'|| T2.""FileName""||'.'|| T2.""FileExt""  FROM ""ORDR""  T0 " &
                " INNER JOIN ""OATC"" T1 ON T0.""AtcEntry""=T1.""AbsEntry"" " &
                " INNER JOIN ""ATC1"" T2 ON T2.""AbsEntry""=T1.""AbsEntry"" Where ""DocEntry""='" & DocEntry & "' " &
                " And T2.""FileExt""='pdf' Order By T2.""Line"" DESC "
            End If
            Dim rsPath As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsPath.DoQuery(Qry)

            If rsPath.RecordCount > 0 Then
                If rsPath.Fields.Item(0).Value <> "" Then
                    PDFPath = rsPath.Fields.Item(0).Value
                Else
                    objMain.objApplication.StatusBar.SetText("Please configure the attachemnt path in User defaults and Print preferences", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            Else
                objMain.objApplication.StatusBar.SetText("Please configure the attachemnt path in User defaults and Print preferences", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            Return PDFPath
        Catch ex As Exception
            ShowErrorMessage(ex.Message)
        End Try
    End Function
    Public Function ConvertFileToBase64(ByVal fileName As String) As String
        Return Convert.ToBase64String(System.IO.File.ReadAllBytes(fileName))
    End Function

    Public Async Function GetManualAttachment(gmDocEntry As String, rptDoc As String, printDoc As String, gmObjType As String) As Task

        Try
            Dim gwBaseUrl = ConfigurationManager.AppSettings("GatewayAPIUrl")


            Dim securedwebserviceurl = gwBaseUrl.ToString() & "/login"
            ServicePointManager.Expect100Continue = True
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11
            ' Skip validation of SSL/TLS certificate
            ServicePointManager.ServerCertificateValidationCallback = Function() True

            Dim httpWebRequest = CType(WebRequest.Create(securedwebserviceurl), HttpWebRequest)
            httpWebRequest.ContentType = "application/json"
            httpWebRequest.Method = "POST"
            '-----------------------
            Dim ci = New RSDBObj With {
                                        .CompanyDB = objMain.objCompany.CompanyDB.ToString(), '' ConfigurationManager.AppSettings("CompanyDB"),
                                        .UserName = ConfigurationManager.AppSettings("UserName"),
                                        .DBInstance = ConfigurationManager.AppSettings("DBInstance"),
                                        .Password = ConfigurationManager.AppSettings("Password")
                                        }
            Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(ci)
            '-----------------------

            Using streamWriter = New StreamWriter(httpWebRequest.GetRequestStream())

                streamWriter.Write(json)
                streamWriter.Flush()
                streamWriter.Close()
            End Using

            httpWebRequest.ProtocolVersion = HttpVersion.Version10
            Dim httpResponse = CType(httpWebRequest.GetResponse(), HttpWebResponse)
            Using streamReader = New StreamReader(httpResponse.GetResponseStream())
                Dim responseFromServer As String = streamReader.ReadToEnd()
            End Using

            For i As Integer = 0 To httpResponse.Headers.Count - 1
                Dim n As String = httpResponse.Headers.GetKey(i)
                If Not Equals(n, "Set-Cookie") Then Continue For

                Dim value As String = httpResponse.Headers.[Get](i)
                Dim ValueB As String = value
                Me.getBS64(value, gmDocEntry, rptDoc, printDoc, gmObjType)

            Next
        Catch ex As Exception
            ShowErrorMessage(ex.Message)
        End Try
    End Function
    Private Async Function getBS64(cc As String, gmDocEntry As String, rptDoc As String, printDoc As String, gmObjType As String) As Task
        Dim lstRF As List(Of ActivityRptObj) = New List(Of ActivityRptObj)()
        Dim pdfObj As ImageViewModel = New ImageViewModel()


        ' Dim str3 As List(Of String) = New List(Of String)()
        Dim ss3 As List(Of List(Of String)) = New List(Of List(Of String))()
        Dim ss4 As List(Of List(Of String)) = New List(Of List(Of String))()
        Try
            If printDoc = "SQ" Or printDoc = "ARRINV" Then

                Dim items As String() = gmDocEntry.Split(","c)
                For Each item As String In items
                    ss3.Add(New List(Of String) From {item})
                Next

                Dim items1 As String() = gmObjType.Split(","c)
                For Each item1 As String In items1
                    ss4.Add(New List(Of String) From {item1})
                Next


                Dim keyName = ""
                Dim keyName1 = ""
                If Equals(printDoc, "STR1") OrElse Equals(printDoc, "STR2") Then
                    keyName = "@num"
                Else
                    keyName = "DocKey@"
                End If
                If Equals(printDoc, "STR1") OrElse Equals(printDoc, "STR2") Then
                    keyName1 = "@num"
                Else
                    keyName1 = "ObjectId@"
                End If
                Dim m3 As ActivityRptObj = New ActivityRptObj() With {
                                                                           .name = keyName,
                                                                           .type = "xsd:string",
                                                                           .value = ss3
                                                                         }
                Dim m4 As ActivityRptObj = New ActivityRptObj() With {
                                                                            .name = keyName1,
                                                                            .type = "xsd:string",
                                                                            .value = ss4
                                                                         }
                lstRF.Add(m3)
                lstRF.Add(m4)
            Else
                ss3.Add(New List(Of String) From {gmDocEntry})
                Dim keyName = ""
                If Equals(printDoc, "STR1") OrElse Equals(printDoc, "STR2") Then
                    keyName = "@num"
                Else
                    keyName = "DocKey@"
                End If
                Dim m3 As ActivityRptObj = New ActivityRptObj() With {
                                                                       .name = keyName,
                                                                       .type = "xsd:decimal"
                                                                     }
                m3.value = ss3
                lstRF.Add(m3)
            End If

            Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(lstRF)
            Dim gwBaseUrl = ConfigurationManager.AppSettings("GatewayAPIUrl")
            Dim securedwebserviceurl = gwBaseUrl.ToString() & "/rs/v1/ExportPDFData?DocCode=".ToString() & rptDoc
            ServicePointManager.Expect100Continue = True
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11

            ServicePointManager.ServerCertificateValidationCallback = Function() True

            Dim container As CookieContainer = New CookieContainer()
            For Each singleCookie In cc.Split("="c)
                Dim match As Match = Regex.Match(singleCookie, "(.+?)=(.+?);")
                Dim match2 As Match = Regex.Match(singleCookie, "=(.+?);")
                Dim k = match2.Groups(1).Value
            Next

            Dim keyStrList1 = cc.Split("="c)
            Dim tokenKey1 = keyStrList1(1)

            Dim keyStrList2 = tokenKey1.Split(";"c)
            Dim tokenKey2 = keyStrList2(0)

            container.Add(New Uri(gwBaseUrl.ToString() & "/rs/v1/ExportPDFData?DocCode=".ToString() & rptDoc), New Cookie("SESSION", tokenKey2))

            Dim httpWebRequest = CType(WebRequest.Create(securedwebserviceurl), HttpWebRequest)

            httpWebRequest.CookieContainer = container


            httpWebRequest.ContentType = "application/json"
            httpWebRequest.Method = "POST"



            Using streamWriter = New StreamWriter(httpWebRequest.GetRequestStream())
                ' string json = "{\"CompanyDB\":\"ASC_LM\"," +
                '"\"UserName\":\"manager\", \"Password\":\"1234\"}";
                streamWriter.Write(json)
                streamWriter.Flush()
                streamWriter.Close()
            End Using



            httpWebRequest.ProtocolVersion = HttpVersion.Version10
            Dim httpResponse = CType(httpWebRequest.GetResponse(), HttpWebResponse)
            Using streamReader = New StreamReader(httpResponse.GetResponseStream())
                Dim responseFromServer As String = streamReader.ReadToEnd()
                pdfObj.BsImg = responseFromServer
                Me.SaveBS64toPDF(pdfObj, printDoc, gmDocEntry)
            End Using
        Catch ex As Exception
            ShowErrorMessage(ex.Message)
            ShowErrorMessage(ex.Message)
        End Try

    End Function
    Public Sub SaveBS64toPDF(base64ToImage As ImageViewModel, docType As String, DocEntry As String)
        Try
            Dim bytes = Convert.FromBase64String(base64ToImage.BsImg)
            Dim spearator = {" "c}
            Dim AttPath As String = GetAttachmnetPath()
            Dim fullOutputPath = AttPath & docType & "_" & DocEntry & "_" & System.DateTime.Now.ToString("yyyyMMdd") & ".pdf"
            If File.Exists(fullOutputPath) Then
                File.Delete(fullOutputPath)
            End If
            Dim stream As FileStream = New FileStream(fullOutputPath, FileMode.CreateNew)
            Dim writer As BinaryWriter = New BinaryWriter(stream)
            writer.Write(bytes)
            '     AttPath = fullOutputPath
            writer.Close()
        Catch ex As Exception
            ShowErrorMessage(ex.Message)
        End Try

    End Sub

    Public Async Function GetManualAttachment_2(gCardCode As String, gDateType As String, aSDate As String, rptDoc As String, printDoc As String) As Task

        Try
            Dim gwBaseUrl = ConfigurationManager.AppSettings("GatewayAPIUrl")
            Dim securedwebserviceurl = gwBaseUrl.ToString() & "/login"
            ServicePointManager.Expect100Continue = True
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11
            ' Skip validation of SSL/TLS certificate
            ServicePointManager.ServerCertificateValidationCallback = Function() True

            Dim httpWebRequest = CType(WebRequest.Create(securedwebserviceurl), HttpWebRequest)
            httpWebRequest.ContentType = "application/json"
            httpWebRequest.Method = "POST"
            '-----------------------
            Dim ci = New RSDBObj With {
                                        .CompanyDB = objMain.objCompany.CompanyDB.ToString(), '' ConfigurationManager.AppSettings("CompanyDB"),
                                        .UserName = ConfigurationManager.AppSettings("UserName"),
                                        .DBInstance = ConfigurationManager.AppSettings("DBInstance"),
                                        .Password = ConfigurationManager.AppSettings("Password")
                                        }
            Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(ci)
            '-----------------------

            Using streamWriter = New StreamWriter(httpWebRequest.GetRequestStream())

                streamWriter.Write(json)
                streamWriter.Flush()
                streamWriter.Close()
            End Using

            httpWebRequest.ProtocolVersion = HttpVersion.Version10
            Dim httpResponse = CType(httpWebRequest.GetResponse(), HttpWebResponse)
            Using streamReader = New StreamReader(httpResponse.GetResponseStream())
                Dim responseFromServer As String = streamReader.ReadToEnd()
            End Using

            For i As Integer = 0 To httpResponse.Headers.Count - 1
                Dim n As String = httpResponse.Headers.GetKey(i)
                If Not Equals(n, "Set-Cookie") Then Continue For

                Dim value As String = httpResponse.Headers.[Get](i)
                Dim ValueB As String = value
                Me.getBS64_2(value, gCardCode, gDateType, aSDate, rptDoc, printDoc)

            Next
        Catch ex As Exception
            ShowErrorMessage(ex.Message)
        End Try
    End Function
    Private Async Function getBS64_2(cc As String, gCardCode As String, gDateType As String, aSDate As String, rptDoc As String, printDoc As String) As Task
        Dim lstRF As List(Of ActivityRptObj) = New List(Of ActivityRptObj)()
        Dim pdfObj As ImageViewModel = New ImageViewModel()


        Dim ss3 As List(Of List(Of String)) = New List(Of List(Of String))()
        Dim ss4 As List(Of List(Of String)) = New List(Of List(Of String))()
        Dim ss5 As List(Of List(Of String)) = New List(Of List(Of String))()

        Try
            Dim items As String() = gCardCode.Split(","c)
            For Each item As String In items
                ss3.Add(New List(Of String) From {item})
            Next

            Dim items1 As String() = gDateType.Split(","c)
            For Each item1 As String In items1
                ss4.Add(New List(Of String) From {item1})
            Next
            Dim items2 As String() = aSDate.Split(","c)
            For Each item2 As String In items2
                ss5.Add(New List(Of String) From {item2})
            Next

            Dim keyName = ""
            Dim keyName1 = ""
            Dim keyName2 = ""
            If Equals(printDoc, "STR1") OrElse Equals(printDoc, "STR2") Then
                keyName = "@num"
            Else
                keyName = "CardCode@"
            End If
            If Equals(printDoc, "STR1") OrElse Equals(printDoc, "STR2") Then
                keyName1 = "@num"
            Else
                keyName1 = "DateType@"
            End If
            If Equals(printDoc, "STR1") OrElse Equals(printDoc, "STR2") Then
                keyName2 = "@num"
            Else
                keyName2 = "AsDate@"
            End If
            Dim m3 As ActivityRptObj = New ActivityRptObj() With {
                                                                   .name = keyName,
                                                                   .type = "xsd:string",
                                                                   .value = ss3
                                                                 }
            Dim m4 As ActivityRptObj = New ActivityRptObj() With {
                                                                    .name = keyName1,
                                                                    .type = "xsd:string",
                                                                    .value = ss4
                                                                 }
            Dim m5 As ActivityRptObj = New ActivityRptObj() With {
                                                                    .name = keyName2,
                                                                    .type = "xsd:string",
                                                                    .value = ss5
                                                                 }
            lstRF.Add(m3)
            lstRF.Add(m4)
            lstRF.Add(m5)
            Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(lstRF)
            Dim gwBaseUrl = ConfigurationManager.AppSettings("GatewayAPIUrl")
            Dim securedwebserviceurl = gwBaseUrl.ToString() & "/rs/v1/ExportPDFData?DocCode=".ToString() & rptDoc
            ServicePointManager.Expect100Continue = True
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11
            ' Skip validation of SSL/TLS certificate
            ServicePointManager.ServerCertificateValidationCallback = Function() True

            Dim container As CookieContainer = New CookieContainer()
            'container.Add(new Uri("https://10.12.80.213:60000/rs/v1/ExportPDFData?DocCode="+rptDoc), new Cookie("SESSION", "value"));
            For Each singleCookie In cc.Split("="c)
                Dim match As Match = Regex.Match(singleCookie, "(.+?)=(.+?);")
                Dim match2 As Match = Regex.Match(singleCookie, "=(.+?);")
                Dim k = match2.Groups(1).Value

                'container.Add(new Uri(gwBaseUrl+"/rs/v1/ExportPDFData?DocCode=" + rptDoc), new Cookie("SESSION", k));
            Next

            Dim keyStrList1 = cc.Split("="c)
            Dim tokenKey1 = keyStrList1(1)

            Dim keyStrList2 = tokenKey1.Split(";"c)
            Dim tokenKey2 = keyStrList2(0)

            container.Add(New Uri(gwBaseUrl.ToString() & "/rs/v1/ExportPDFData?DocCode=".ToString() & rptDoc), New Cookie("SESSION", tokenKey2))

            Dim httpWebRequest = CType(WebRequest.Create(securedwebserviceurl), HttpWebRequest)

            httpWebRequest.CookieContainer = container


            httpWebRequest.ContentType = "application/json"
            httpWebRequest.Method = "POST"



            Using streamWriter = New StreamWriter(httpWebRequest.GetRequestStream())
                ' string json = "{\"CompanyDB\":\"ASC_LM\"," +
                '"\"UserName\":\"manager\", \"Password\":\"1234\"}";
                streamWriter.Write(json)
                streamWriter.Flush()
                streamWriter.Close()
            End Using



            httpWebRequest.ProtocolVersion = HttpVersion.Version10
            Dim httpResponse = CType(httpWebRequest.GetResponse(), HttpWebResponse)
            Using streamReader = New StreamReader(httpResponse.GetResponseStream())
                Dim responseFromServer As String = streamReader.ReadToEnd()
                pdfObj.BsImg = responseFromServer
                Me.SaveBS64toPDF(pdfObj, printDoc, gCardCode)
            End Using
        Catch ex As Exception
            ShowErrorMessage(ex.Message)
        End Try

    End Function
    Sub UpdateTransactionNotification()
        Try

            Dim oRecordset As Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim procedureName As String = "SBO_SP_TRANSACTIONNOTIFICATION"

            ' 1. Get existing procedure source
            Dim DBName As String = objMain.objCompany.CompanyDB
            Dim Qry As String = "SELECT ""DEFINITION"" FROM ""SYS"".""PROCEDURES"" WHERE ""SCHEMA_NAME"" = '" & DBName & "' AND ""PROCEDURE_NAME"" = 'SBO_SP_TRANSACTIONNOTIFICATION'"
            oRecordset.DoQuery(Qry)

            If oRecordset.EoF Then
                Throw New Exception("Procedure not found.")
            End If

            Dim originalCode As String = oRecordset.Fields.Item(0).Value.ToString()
            If originalCode.Contains("TNX_RETAIL_TN") Then
                Exit Sub
            End If
            ' 2. Your code to append
            Dim codeToInsert As String = "
    
    if :error = 0 then
        CALL TNX_RETAIL_TN(:object_type, :transaction_type, :num_of_cols_in_key, :list_of_key_cols_tab_del, :list_of_cols_val_tab_del, :error, :error_message);
    end if;
select :error, :error_message FROM dummy;
    "

            ' 3. Insert before the final "select :error, :error_message FROM dummy;"
            Dim finalCode As String = ""
            If originalCode.Contains("select :error, :error_message FROM dummy;") Then
                finalCode = originalCode.Replace("select :error, :error_message FROM dummy;", codeToInsert)
            Else
                Throw New Exception("""select :error, :error_message FROM dummy;"" not found in procedure definition.")
            End If
            finalCode = finalCode.Replace("CREATE", "ALTER")

            oRecordset.DoQuery(finalCode)
            objMain.objApplication.StatusBar.SetText("Transaction Notification Procedure updated successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    'madhu Email notification code

    Public Sub MarketingEmail(ByVal DocNum As String)
        Dim Result As String = ""
        Dim query As String = "SELECT T3.""U_UserIDE"", T3.""U_EmpIDE"", T1.""E_Mail"" AS ""User_Email"", T2.""email"" AS ""Employee_Email"" " &
                          "FROM ""@MGRCONFIG"" T0 " &
                          "INNER JOIN ""@MGRCONFIG_C2"" T3 ON T0.""DocEntry"" = T3.""DocEntry"" " &
                          "INNER JOIN ""OUSR"" T1 ON T1.""USERID"" = T3.""U_UserIDE"" " &
                          "INNER JOIN ""OHEM"" T2 ON T2.""Code"" = T3.""U_EmpIDE"" " &
                          "WHERE T0.""U_Doc"" = 'PCR'"

        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs.DoQuery(query)

        Dim UserEmail As String = oRs.Fields.Item("User_Email").Value.ToString().Trim()
        Dim EmployeeEmail As String = oRs.Fields.Item("Employee_Email").Value.ToString().Trim()

        If UserEmail <> "" And EmployeeEmail <> "" Then
            Result = Me.SendEmail_TEST(UserEmail & ";" & EmployeeEmail, DocNum)
        ElseIf UserEmail <> "" Then
            Result = Me.SendEmail_TEST(UserEmail, DocNum)
        ElseIf EmployeeEmail <> "" Then
            Result = Me.SendEmail_TEST(EmployeeEmail, DocNum)
        Else
            objMain.objApplication.StatusBar.SetText("No valid email found for DocNum: " & DocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If

        If Result = "Success" Then
            objMain.objApplication.StatusBar.SetText("Email sent successfully for " & DocNum & " !!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Else
            objMain.objApplication.StatusBar.SetText("Email sending failed for " & DocNum & " !!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If
    End Sub


    Function SendEmail_TEST(ByVal ToEmail As String, ByVal DocNum1 As String) As String
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
                ' Dim Message As New MailMessage(FromMail, ToEmail)

                Dim Message As New MailMessage()
                Message.From = New MailAddress(FromMail)

                For Each email As String In ToEmail.Split(";"c)
                    If email.Trim() <> "" Then
                        Message.To.Add(email.Trim())
                    End If
                Next

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

                Message.Subject = "Price Change Request"

                Dim Attachment As String = "Price Change Request is added successfully"
                Dim Base64 As String = Nothing



                Dim htmlBody As String = "<table>" &
                         "<tr><td style='font-family:Times New Roman; font-size:14px;'>Document Number : <b>" & DocNum1 & "</b></td></tr>" &
                         "<tr><td style='font-size:14px; font-weight:normal; font-family:Times New Roman;'>" & Attachment & "

                          </td></tr>" &
                         "</table>" &
                         "<table>" &
                         "<tr><td style='font-size:14px; font-weight:normal; font-family:Times New Roman;'>This is system auto-generated mail. Please do not reply.</td></tr>" &
                         "</table>"

                Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(htmlBody, Nothing, "text/html")



                Message.AlternateViews.Add(htmlView)
                SMTPServer.Send(Message)

                Message.Dispose()
                objMain.objUtilities.Write_Log("E-Mail Sent successfully to -  " & ToEmail)
                objMain.objApplication.StatusBar.SetText("E-Mail Sent successfully to -  " & ToEmail, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Return "Success"
            End If
        Catch ex As Exception
            objMain.objUtilities.Write_Log("Error In Email Sending:" & ex.Message)
            objMain.objApplication.StatusBar.SetText("Error In Email Sending:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
    End Function

#End Region
End Class

Public Enum ResourceType
    Embeded
    Content
End Enum



