Imports System.Configuration
Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports Sap.Data.Hana
Imports TaskScheduler
Imports Attachment = System.Net.Mail.Attachment


Public Class clsPACPPriceChange

#Region "Decliration"

    Public objForm As SAPbouiCOM.Form
    Dim objMatrix, objMatrix1 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Public oDBs_Head, oDBs_Details, oDBs_Details1, oDBs_Summary, oDBs_Attacht, oDS, oDS1 As SAPbouiCOM.DBDataSource
    Public Prgbar As SAPbouiCOM.ProgressBar
    Dim objutilities As Utilities
    Public TERM As String
    Private FileName As String = ""
    Private CurrentFormUID As String = ""

#End Region
    Sub CreateForm()

        Try

            objMain.objUtilities.LoadForm("PACPPriceChange.xml", "PACP", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("PACP", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.EnableMenu("774", True)
            objForm.EnableMenu("1292", True)

            objForm.Freeze(True)
            objutilities = New Utilities
            InitializeBackgroundWorker()

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PACP")
            objMatrix = objForm.Items.Item("Mtx").Specific

            ' Me.SetNewLineContent(objForm.UniqueID)
            Me.Matrixload(objForm.UniqueID)
            'objForm.Items.Item("Mtx").Enabled = False

            'Dim oCol As SAPbouiCOM.Column
            'For i As Integer = 0 To objMatrix.Columns.Count - 1
            '    oCol = objMatrix.Columns.Item(i)
            '    oCol.Editable = False
            'Next

            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Successfully initialized, please procced..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "PACP" And pVal.BeforeAction = False Then
                CreateForm()

            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("PACP", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix = objForm.Items.Item("Mtx").Specific
                SetNewLineContent(objForm.UniqueID)

            ElseIf pVal.MenuUID = "774" And pVal.BeforeAction = False Then

                objMatrix = objForm.Items.Item("Mtx").Specific
                Dim row As Integer = objMatrix.VisualRowCount

                If objMatrix.IsRowSelected(1) <> True And objMatrix.VisualRowCount < 1 Then
                    objMatrix.AddRow()
                    oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.VisualRowCount)
                    objMatrix.SetLineData(objMatrix.VisualRowCount)
                End If

                If objMatrix.IsRowSelected(row) = True Then
                    objMatrix.DeleteRow(row)
                Else
                    For i As Integer = 1 To objMatrix.VisualRowCount - 1
                        If objMatrix.IsRowSelected(i) = True Then
                            objMatrix.DeleteRow(i)
                        End If
                    Next
                End If

                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.Columns.Item("DocEntry").Cells.Item(i).Specific.Value = i
                Next

                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)

        Try

            Select Case pVal.EventType

                '------Madhu when Mutiples Matrixes are code here ...!
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    'If pVal.ItemUID = "Mtx" Or pVal.ItemUID = "Mtx1" Then
                    '    TERM = pVal.ItemUID
                    'End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If pVal.ItemUID = "Upload" And pVal.BeforeAction = False Then

                        StartExcelUpload(FormUID)

                    End If


            End Select

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("PACP", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then


                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then


                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLineContent(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PACP")
            objMatrix = objForm.Items.Item("Mtx").Specific
            objMatrix.AddRow()

            oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.VisualRowCount)
            oDBs_Head.SetValue("U_Date", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
            oDBs_Head.SetValue("U_VName", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_Barcode", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_IName", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("Code", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("Name", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_Excost", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_Ncost", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_NCV", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_NSelling", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_Margin", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_Remark", oDBs_Head.Offset, "")

            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()

            objMain.objApplication.StatusBar.SetText("Matrix New line Inserted..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub


    Public Sub Matrixload(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("Mtx").Specific
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PACP")

            Dim rs1 As String = "SELECT ""DocEntry"", ""U_Date"", ""U_VName"", ""Code"", ""Name"", ""U_ICode"", ""U_Excost"", ""U_Ncost"", ""U_NCV"", ""U_NSelling"", ""U_Margin"", ""U_Remark"" FROM ""@TNX_PACP"" ORDER BY ""DocEntry"" ASC"

            Dim ors1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ors1.DoQuery(rs1)

            objMatrix.Clear()
            oDBs_Head.Clear()

            If ors1.RecordCount > 0 Then
                For j As Integer = 1 To ors1.RecordCount
                    objMatrix.AddRow()

                    oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.VisualRowCount)
                    Dim dtVal As String = ""
                    If Not IsDBNull(ors1.Fields.Item("U_Date").Value) AndAlso ors1.Fields.Item("U_Date").Value IsNot Nothing Then
                        Try
                            Dim dt As Date = CDate(ors1.Fields.Item("U_Date").Value)
                            dtVal = dt.ToString("yyyyMMdd")
                        Catch
                            dtVal = ors1.Fields.Item("U_Date").Value.ToString().Trim()
                        End Try
                    End If
                    oDBs_Head.SetValue("U_Date", oDBs_Head.Offset, dtVal)
                    oDBs_Head.SetValue("U_VName", oDBs_Head.Offset, If(IsDBNull(ors1.Fields.Item("U_VName").Value), "", ors1.Fields.Item("U_VName").Value.ToString()))
                    oDBs_Head.SetValue("Code", oDBs_Head.Offset, If(IsDBNull(ors1.Fields.Item("Code").Value), "", ors1.Fields.Item("Code").Value.ToString()))
                    oDBs_Head.SetValue("Name", oDBs_Head.Offset, If(IsDBNull(ors1.Fields.Item("Name").Value), "", ors1.Fields.Item("Name").Value.ToString()))
                    oDBs_Head.SetValue("U_ICode", oDBs_Head.Offset, If(IsDBNull(ors1.Fields.Item("U_ICode").Value), "", ors1.Fields.Item("U_ICode").Value.ToString()))
                    oDBs_Head.SetValue("U_Excost", oDBs_Head.Offset, If(IsDBNull(ors1.Fields.Item("U_Excost").Value), "0", ors1.Fields.Item("U_Excost").Value.ToString()))
                    oDBs_Head.SetValue("U_Ncost", oDBs_Head.Offset, If(IsDBNull(ors1.Fields.Item("U_Ncost").Value), "0", ors1.Fields.Item("U_Ncost").Value.ToString()))
                    oDBs_Head.SetValue("U_NCV", oDBs_Head.Offset, If(IsDBNull(ors1.Fields.Item("U_NCV").Value), "0", ors1.Fields.Item("U_NCV").Value.ToString()))
                    oDBs_Head.SetValue("U_NSelling", oDBs_Head.Offset, If(IsDBNull(ors1.Fields.Item("U_NSelling").Value), "0", ors1.Fields.Item("U_NSelling").Value.ToString()))
                    oDBs_Head.SetValue("U_Margin", oDBs_Head.Offset, If(IsDBNull(ors1.Fields.Item("U_Margin").Value), "0", ors1.Fields.Item("U_Margin").Value.ToString()))
                    oDBs_Head.SetValue("U_Remark", oDBs_Head.Offset, If(IsDBNull(ors1.Fields.Item("U_Remark").Value), "", ors1.Fields.Item("U_Remark").Value.ToString()))

                    objMatrix.SetLineData(objMatrix.VisualRowCount)
                    ors1.MoveNext()
                Next
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(ors1)
            ors1 = Nothing

            objMatrix.AutoResizeColumns()
            objMatrix.FlushToDataSource()
            objMatrix.LoadFromDataSource()
            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

        Catch ex As Exception
            Me.objForm.Freeze(False)
            Me.objForm.StatusBar.SetText(ex.Message & "   Page Load Events", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    '
#Region "Excel Upload"

    Private BackgroundWorker_Excel As New System.ComponentModel.BackgroundWorker

    Public Sub InitializeBackgroundWorker()

        Try
            RemoveHandler BackgroundWorker_Excel.DoWork, AddressOf BackgroundWorker_Excel_DoWork

            AddHandler BackgroundWorker_Excel.DoWork, AddressOf BackgroundWorker_Excel_DoWork

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub

    Public Function ValidateExcelPath(ByVal FormUID As String) As Boolean

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim FilePath As String = ""

            Try
                FilePath = CType(objForm.Items.Item("Excel").Specific, SAPbouiCOM.EditText).Value.Trim()
            Catch ex As Exception
                FilePath = ""
            End Try

            If String.IsNullOrWhiteSpace(FilePath) Then
                objMain.objApplication.StatusBar.SetText("Please select Excel file path before clicking Upload.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            If Not IO.File.Exists(FilePath) Then
                objMain.objApplication.StatusBar.SetText("Selected Excel file does not exist.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Return True

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message)

            Return False

        End Try

    End Function

    Public Sub StartExcelUpload(ByVal FormUID As String)

        Try

            CurrentFormUID = FormUID
            objForm = objMain.objApplication.Forms.Item(FormUID)

            If ValidateExcelPath(FormUID) = False Then
                Exit Sub
            End If

            Dim FilePath As String = CType(objForm.Items.Item("Excel").Specific, SAPbouiCOM.EditText).Value.Trim()

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PACP")

            Try
                oDBs_Head.SetValue("U_Excel", oDBs_Head.Offset, FilePath)
            Catch ex As Exception
            End Try

            If BackgroundWorker_Excel.IsBusy = False Then
                BackgroundWorker_Excel.RunWorkerAsync(FilePath)
                objMain.objApplication.StatusBar.SetText("Excel Upload Started...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else

                objMain.objApplication.StatusBar.SetText("Upload already running. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            End If

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

    Private Sub BackgroundWorker_Excel_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs)

        Try

            Dim FilePath As String = ""

            If e.Argument IsNot Nothing Then
                FilePath = e.Argument.ToString.Trim()
            End If

            If String.IsNullOrWhiteSpace(FilePath) Then

                objMain.objApplication.StatusBar.SetText(
                "Excel path is empty.",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Exit Sub

            End If

            If Not IO.File.Exists(FilePath) Then

                objMain.objApplication.StatusBar.SetText(
                "Excel file not found.",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Exit Sub

            End If


            ReadExcelData(FilePath)

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
            "Upload Error : " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

#End Region

    Private Function GetSafeDecimalString(ByVal dt As DataTable, ByVal rowIndex As Integer, ByVal colName As String) As String
        If Not dt.Columns.Contains(colName) OrElse IsDBNull(dt.Rows(rowIndex)(colName)) Then
            Return "0"
        End If
        Dim rawStr As String = dt.Rows(rowIndex)(colName).ToString().Trim().Replace(",", "")
        If String.IsNullOrWhiteSpace(rawStr) Then
            Return "0"
        End If
        Dim val As Double
        If Double.TryParse(rawStr, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, val) Then
            Return val.ToString(Globalization.CultureInfo.InvariantCulture)
        End If
        If Double.TryParse(rawStr, Globalization.NumberStyles.Any, Globalization.CultureInfo.CurrentCulture, val) Then
            Return val.ToString(Globalization.CultureInfo.InvariantCulture)
        End If
        Return "0"
    End Function

    Public Sub ReadExcelData(ByVal FilePath As String)

        Dim ConnString As String = ""
        Dim Query As String = ""
        Dim dtExcel As New DataTable

        Try
            If Not String.IsNullOrEmpty(CurrentFormUID) Then
                Try
                    objForm = objMain.objApplication.Forms.Item(CurrentFormUID)
                Catch
                    objForm = objMain.objApplication.Forms.ActiveForm
                End Try
            Else
                objForm = objMain.objApplication.Forms.ActiveForm
            End If

            objMatrix = objForm.Items.Item("Mtx").Specific
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PACP")

            '==================================================
            ' 1. EXCEL CONNECTION & LOADING
            '==================================================
            If Path.GetExtension(FilePath).ToUpper() = ".XLSX" Then
                ConnString = "Provider=Microsoft.ACE.OLEDB.12.0;" &
                         "Data Source=" & FilePath & ";" &
                         "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';"
            Else
                ConnString = "Provider=Microsoft.Jet.OLEDB.4.0;" &
                         "Data Source=" & FilePath & ";" &
                         "Extended Properties='Excel 8.0;HDR=YES;IMEX=1';"
            End If

            Using Conn As New OleDbConnection(ConnString)
                Conn.Open()
                Dim dtSchema As DataTable = Conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                Dim SheetName As String = ""
                For Each row As DataRow In dtSchema.Rows
                    Dim tName As String = row("TABLE_NAME").ToString()
                    If tName.EndsWith("$") OrElse tName.EndsWith("$'") Then
                        SheetName = tName
                        Exit For
                    End If
                Next
                If String.IsNullOrEmpty(SheetName) AndAlso dtSchema.Rows.Count > 0 Then
                    SheetName = dtSchema.Rows(0)("TABLE_NAME").ToString()
                End If

                If SheetName.StartsWith("'") AndAlso SheetName.EndsWith("'") Then
                    SheetName = SheetName.Substring(1, SheetName.Length - 2)
                End If
                If Not SheetName.EndsWith("$") Then
                    SheetName = SheetName & "$"
                End If

                Query = "SELECT * FROM [" & SheetName & "]"

                Using da As New OleDbDataAdapter(Query, Conn)
                    da.Fill(dtExcel)
                End Using
                Conn.Close()
            End Using

            If dtExcel.Rows.Count = 0 Then
                objMain.objApplication.StatusBar.SetText("No records found in Excel.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            '==================================================
            ' 2. DIRECT BACKEND HANA DATABASE PROCESSING LOOP
            '==================================================
            Dim oRec As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim maxDocEntry As Integer = 0
            Dim oRsMax As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                oRsMax.DoQuery("SELECT IFNULL(MAX(""DocEntry""), 0) FROM ""@TNX_PACP""")
                If oRsMax.RecordCount > 0 Then
                    Integer.TryParse(oRsMax.Fields.Item(0).Value.ToString(), maxDocEntry)
                End If
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRsMax)
                oRsMax = Nothing
            End Try

            For i As Integer = 0 To dtExcel.Rows.Count - 1

                Dim barcodeVal As String = ""
                If dtExcel.Columns.Contains("BarCode") AndAlso Not IsDBNull(dtExcel.Rows(i)("BarCode")) Then
                    barcodeVal = dtExcel.Rows(i)("BarCode").ToString().Trim()
                End If

                ' Skip if there's no barcode key
                If String.IsNullOrEmpty(barcodeVal) Then Continue For

                Dim barcodeClean As String = barcodeVal.Replace("'", "''")

                Dim excelDate As String = ""
                If dtExcel.Columns.Contains("Date") AndAlso Not IsDBNull(dtExcel.Rows(i)("Date")) Then
                    Dim rawDate As Object = dtExcel.Rows(i)("Date")
                    If TypeOf rawDate Is DateTime Then
                        excelDate = CDate(rawDate).ToString("yyyyMMdd")
                    Else
                        Dim strDate As String = rawDate.ToString().Trim()
                        Dim parsedDate As DateTime
                        Dim dateFormats() As String = {"ddMMyyyy", "dd/MM/yyyy", "dd-MM-yyyy", "yyyyMMdd", "yyyy-MM-dd", "yyyy/MM/dd", "MM/dd/yyyy", "M/d/yyyy", "d/M/yyyy"}
                        If DateTime.TryParseExact(strDate, dateFormats, Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, parsedDate) Then
                            excelDate = parsedDate.ToString("yyyyMMdd")
                        ElseIf DateTime.TryParse(strDate, parsedDate) Then
                            excelDate = parsedDate.ToString("yyyyMMdd")
                        End If
                    End If
                End If

                If String.IsNullOrEmpty(excelDate) Then
                    excelDate = DateTime.Now.ToString("yyyyMMdd")
                End If

                ' Gather variables safely from Excel row and handle quotes safely
                Dim vName As String = If(dtExcel.Columns.Contains("Vendor-Name") AndAlso Not IsDBNull(dtExcel.Rows(i)("Vendor-Name")), dtExcel.Rows(i)("Vendor-Name").ToString().Trim().Replace("'", "''"), "")
                Dim iName As String = If(dtExcel.Columns.Contains("Item Name") AndAlso Not IsDBNull(dtExcel.Rows(i)("Item Name")), dtExcel.Rows(i)("Item Name").ToString().Trim().Replace("'", "''"), "")
                Dim iCode As String = If(dtExcel.Columns.Contains("Item Code") AndAlso Not IsDBNull(dtExcel.Rows(i)("Item Code")), dtExcel.Rows(i)("Item Code").ToString().Trim().Replace("'", "''"), "")
                Dim itemName As String = If(String.IsNullOrWhiteSpace(iName), barcodeClean, iName)

                Dim exCost As String = GetSafeDecimalString(dtExcel, i, "Existing cost")
                Dim nCost As String = GetSafeDecimalString(dtExcel, i, "New cost")
                Dim ncv As String = GetSafeDecimalString(dtExcel, i, "New cost+vat")
                Dim nSelling As String = GetSafeDecimalString(dtExcel, i, "New Selling")
                Dim margin As String = GetSafeDecimalString(dtExcel, i, "Margin")
                Dim remark As String = If(dtExcel.Columns.Contains("Remark") AndAlso Not IsDBNull(dtExcel.Rows(i)("Remark")), dtExcel.Rows(i)("Remark").ToString().Trim().Replace("'", "''"), "")

                ' HANA COMPLIANT CHECK: See if record already exists in the table
                oRec.DoQuery("SELECT COUNT(*) FROM ""@TNX_PACP"" WHERE ""Code"" = '" & barcodeClean & "'")
                Dim recordExists As Boolean = False
                If oRec.RecordCount > 0 Then
                    recordExists = (Integer.Parse(oRec.Fields.Item(0).Value.ToString()) > 0)
                End If

                Dim sqlQuery As String = ""
                If recordExists Then
                    ' HANA COMPLIANT UPDATE
                    sqlQuery = "UPDATE ""@TNX_PACP"" SET " &
                               """Name"" = '" & itemName & "', " &
                               """U_Date"" = '" & excelDate & "', " &
                               """U_VName"" = '" & vName & "', " &
                               """U_Barcode"" = '" & barcodeClean & "', " &
                               """U_IName"" = '" & iName & "', " &
                               """U_ICode"" = '" & iCode & "', " &
                               """U_Excost"" = " & exCost & ", " &
                               """U_Ncost"" = " & nCost & ", " &
                               """U_NCV"" = " & ncv & ", " &
                               """U_NSelling"" = " & nSelling & ", " &
                               """U_Margin"" = " & margin & ", " &
                               """U_Remark"" = '" & remark & "' " &
                               "WHERE ""Code"" = '" & barcodeClean & "'"
                Else
                    ' HANA COMPLIANT INSERT WITH DocEntry
                    maxDocEntry += 1
                    sqlQuery = "INSERT INTO ""@TNX_PACP"" (""DocEntry"", ""Code"", ""Name"", ""U_Date"", ""U_VName"", ""U_Barcode"", ""U_IName"", ""U_ICode"", ""U_Excost"", ""U_Ncost"", ""U_NCV"", ""U_NSelling"", ""U_Margin"", ""U_Remark"") VALUES (" &
                               maxDocEntry & ", '" & barcodeClean & "', '" & itemName & "', '" & excelDate & "', '" & vName & "', '" & barcodeClean & "', '" & iName & "', '" & iCode & "', " & exCost & ", " & nCost & ", " & ncv & ", " & nSelling & ", " & margin & ", '" & remark & "')"
                End If

                ' Execute direct backend HANA query commit
                oRec.DoQuery(sqlQuery)
            Next

            ' Clean up backend recordset object
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec)
            oRec = Nothing

            '==================================================
            ' 3. REFRESH SCREEN DISPLAY FROM REFRESHED TABLE
            '==================================================
            Try
                objForm.Freeze(True)
                objMatrix.Clear()
                oDBs_Head.Clear()
                oDBs_Head.Query()

                For r As Integer = 0 To oDBs_Head.Size - 1
                    Try
                        oDBs_Head.SetValue("DocEntry", r, (r + 1).ToString())
                    Catch
                    End Try
                Next

                objMatrix.LoadFromDataSource()
                objMatrix.AutoResizeColumns()
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
            Finally
                Try
                    objForm.Freeze(False)
                Catch
                End Try
            End Try

            objMain.objApplication.StatusBar.SetText("Excel Upload Completed ....", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Excel Upload Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            Try
                objForm.Freeze(False)
            Catch
            End Try
        End Try

    End Sub



End Class




