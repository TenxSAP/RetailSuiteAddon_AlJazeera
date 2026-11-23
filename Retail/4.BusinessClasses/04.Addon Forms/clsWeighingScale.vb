Imports System.Collections.Specialized
Imports System.Configuration
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices

Public Class clsWeighingScale
#Region "       Declaration             "
    Public objForm, objSTRForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details1 As SAPbouiCOM.DBDataSource
    Dim objMatrix1 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
#End Region
    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("WeighingScaleTextfileGen.xml", "frmScale", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("frmScale", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            Dim oRsBranch As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim Query As String = "SELECT T2.""BPLId"",T2.""BPLName"" AS ""Branch"" FROM USR6 T0  INNER JOIN OUSR T1 ON T0.""UserCode"" = T1.""USER_CODE"" " &
                    " INNER JOIN OBPL T2 ON T0.""BPLId"" = T2.""BPLId"" WHERE T1.""USER_CODE"" = '" & objMain.objCompany.UserName.ToString().Trim & "' "
            oRsBranch.DoQuery(Query)

            objComboBox = objForm.Items.Item("Branch").Specific
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, Query)
            Try
                objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                objForm.Items.Item("Branch").DisplayDesc = True
                If objComboBox.ValidValues.Count <> 0 Then
                    objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
                End If
            Catch ex As Exception
            End Try

            Dim UOMCmbo As SAPbouiCOM.ComboBox = DirectCast(Me.objForm.Items.Item("Group").Specific, SAPbouiCOM.ComboBox)
            ' Dim UOMQry As String = "SELECT T0.""ItmsGrpCod"", T0.""ItmsGrpNam"" FROM OITB T0"

            Dim UOMQry As String = "SELECT CAST(""ItmsGrpCod"" AS NVARCHAR) AS ""ItmsGrpCod"", ""ItmsGrpNam"" FROM OITB
            UNION ALL 
            Select 'All' AS ""ItmsGrpCod"", 'All' AS ""ItmsGrpNam"" From DUMMY"
            objMain.objUtilities.ComboBoxLoadValues(UOMCmbo, UOMQry)
            UOMCmbo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            If UOMCmbo.ValidValues.Count <> 0 Then
                UOMCmbo.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
            End If

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ME_WG" And pVal.BeforeAction = False Then
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
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    'If pVal.ItemUID = "btnOk" And pVal.BeforeAction = False Then
                    '    If HasPendingPriceUpdates() Then
                    '        objMain.objApplication.StatusBar.SetText(
                    '        "Some prices are still updating. Please wait until all price updates are completed before generating the file.",
                    '        SAPbouiCOM.BoMessageTime.bmt_Long,
                    '        SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        Exit Sub
                    '    End If
                    '    Me.FileGeneration(pVal)
                    'End If
                    If pVal.ItemUID = "btnOk" AndAlso pVal.BeforeAction = False Then

                        Dim PendingDocNum As String = ""

                        If HasPendingPriceUpdates(PendingDocNum) Then

                            Dim choice As Integer = objMain.objApplication.MessageBox("Price updates are still pending for Document No: " & PendingDocNum &
        ". Do you want to continue with file generation?", 1, "Cancel", "Continue")

                            If choice = 1 Then
                                Exit Sub
                            ElseIf choice = 2 Then
                                Me.FileGeneration(pVal)
                            End If

                        Else
                            Me.FileGeneration(pVal)
                        End If

                    End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Function HasPendingPriceUpdates(ByRef PendingDocNum As String) As Boolean

        Dim rs As SAPbobsCOM.Recordset = Nothing
        PendingDocNum = ""

        Try
            Dim Query As String = "
        SELECT TOP 1 P.""DocEntry"", P.""DocNum""
        FROM ""@TNX_PCRT"" P
        WHERE P.""U_TNXSTUS"" = 'A'
        AND (
            EXISTS (SELECT 1 FROM ""@TNX_PCRT_C0"" C0 WHERE C0.""DocEntry"" = P.""DocEntry"" AND C0.""U_PUpdate"" = 'No' AND C0.""U_BCode"" <> '')
            OR EXISTS (SELECT 1 FROM ""@TNX_PCRT_C5"" C5 WHERE C5.""DocEntry"" = P.""DocEntry"" AND C5.""U_PUpdate"" = 'No' AND C5.""U_BCode"" <> '')
            OR EXISTS (SELECT 1 FROM ""@TNX_PCRT_C6"" C6 WHERE C6.""DocEntry"" = P.""DocEntry"" AND C6.""U_PUpdate"" = 'No' AND C6.""U_BCode"" <> '')
            OR EXISTS (SELECT 1 FROM ""@TNX_PCRT_C7"" C7 WHERE C7.""DocEntry"" = P.""DocEntry"" AND C7.""U_PUpdate"" = 'No' AND C7.""U_BCode"" <> '')
            OR EXISTS (SELECT 1 FROM ""@TNX_PCRT_C8"" C8 WHERE C8.""DocEntry"" = P.""DocEntry"" AND C8.""U_PUpdate"" = 'No' AND C8.""U_BCode"" <> '')
        )
        ORDER BY P.""DocNum"""

            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rs.DoQuery(Query)

            If Not rs.EoF Then
                PendingDocNum = rs.Fields.Item("DocNum").Value.ToString()
                Return True
            End If

            Return False

        Catch ex As Exception
            Throw New Exception("Price update validation failed: " & ex.Message)

        Finally
            If rs IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
                rs = Nothing
            End If
        End Try

    End Function

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("frmBarCodeIM", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Function Notepad_Read(ByVal QueryTextFile As String) As String
        Dim str As String = Nothing
        Try
            Dim sPath As String
            sPath = IO.Directory.GetParent(Application.ExecutablePath).ToString
            sPath = sPath & "\" & QueryTextFile

            If (Not File.Exists(sPath)) Then
                objMain.objApplication.StatusBar.SetText("File Is Not In this path - '" & sPath & "'", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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

    Public Function GetFieldValue(ByVal query As String) As Object
        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs.DoQuery(query)
        Dim str As String = oRs.Fields.Item(0).Value.ToString()
        Return str
    End Function
    Private Sub FileGeneration_old(ByVal pVal As SAPbouiCOM.ItemEvent)

        Dim str As String = Nothing
        Dim str1 As String = Nothing
        Dim str2 As String = Nothing

        Dim businessObject As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(300), SAPbobsCOM.Recordset)
        Dim recordset As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(300), SAPbobsCOM.Recordset)

        Try
            Dim BranchID As String
            Dim ItemGroup As String
            Dim ItemGroupName As String
            Dim FilterType As String
            Dim ScaleType As String

            ' Branch
            Dim objBranchComb As SAPbouiCOM.ComboBox = objForm.Items.Item("Branch").Specific
            Try
                BranchID = objBranchComb.Selected.Value
            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText("Please select the branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End Try

            ' Item Group
            Dim objGroupCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Group").Specific
            Try
                ItemGroup = objGroupCombo.Selected.Value
                ItemGroupName = objGroupCombo.Selected.Description
            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText("Please select item group", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End Try

            ' Filter Type
            Dim objFilterCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Date").Specific
            Try
                FilterType = objFilterCombo.Selected.Value
            Catch ex As Exception
                FilterType = "2"
            End Try

            ' ✅ ONLY Scale Type (MANDATORY)
            Try
                Dim objScaleCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("SType").Specific
                ScaleType = objScaleCombo.Selected.Value
            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText("Please select Scale Type", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End Try

            objMain.objApplication.StatusBar.SetText("Please wait... Price Detail text file generation in progress", 20, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            ' ✅ Call Procedure
            Dim query As String
            query = "CALL ""PriceList_TextFile_Generation_20251017""('" & BranchID & "','" & ItemGroup & "','" & FilterType & "','" & ScaleType & "')"

            businessObject.DoQuery(query)

            If businessObject.RecordCount <= 0 Then
                objMain.objApplication.StatusBar.SetText("No data found for selected Scale Type", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Exit Sub
            End If

            ' File Path
            recordset.DoQuery("SELECT IFNULL(""U_FileLoc"", '') AS ""Location"" FROM OBPL WHERE ""BPLId"" = '" & BranchID & "'")
            recordset.MoveFirst()

            If recordset.Fields.Item("Location").Value.ToString().Trim() = "" Then
                objMain.objApplication.StatusBar.SetText("Please update Text File Location in branch details", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            str1 = recordset.Fields.Item("Location").Value.ToString()

            Dim recordCount As Integer = businessObject.RecordCount
            Dim num As Integer = 1

            Do
                Dim rateValue As Decimal = Convert.ToDecimal(businessObject.Fields.Item("Rate").Value)
                str2 = rateValue.ToString("F3")

                Dim strArrays() As String = {
                businessObject.Fields.Item("FLAG").Value.ToString(),
                ",", businessObject.Fields.Item("EANCode").Value.ToString(),
                ",", businessObject.Fields.Item("PLU").Value.ToString(),
                ",", businessObject.Fields.Item("ItemName").Value.ToString(),
                ",", businessObject.Fields.Item("arabic name").Value.ToString(),
                ",", businessObject.Fields.Item("UomCode").Value.ToString(),
                ",", str2,
                ",", businessObject.Fields.Item("Department").Value.ToString(),
                ",", businessObject.Fields.Item("Expire Days").Value.ToString()
            }

                Dim pathArray() As String = str1.Split(";"c)
                For Each path As String In pathArray
                    path = path.Trim()
                    If Not String.IsNullOrEmpty(path) Then
                        '
                        Dim timeStamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")

                        Dim fileSuffix As String
                        If ScaleType = "Internal Scale" Then
                            fileSuffix = "Internal"
                        ElseIf ScaleType = "Tenant Scale" Then
                            fileSuffix = "Tenant"
                        Else
                            fileSuffix = ScaleType
                        End If

                        Dim finalFileName As String = ItemGroupName & "_" & fileSuffix & "_" & timeStamp

                        Me.write_log(String.Concat(strArrays), path, finalFileName)

                        'Me.write_log(String.Concat(strArrays), path, num, ItemGroupName)
                    End If
                Next

                businessObject.MoveNext()
                num += 1

            Loop While num <= recordCount

            objMain.objApplication.StatusBar.SetText("Price Detail text file generation completed successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error during file generation: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try

    End Sub


    'Public Sub write_log(ByVal status As String, ByVal Location As String, ByVal Time As Integer, ByVal ItemGroupName As String)
    '    Dim fileStream As System.IO.FileStream
    '    Dim str As String = Location
    '    Try
    '        Try
    '            Dim directoryInfo As System.IO.DirectoryInfo = New System.IO.DirectoryInfo(String.Concat(str, "Scale"))
    '            'Dim str1 As String = String.Concat(str, "Scale.txt")
    '            'Dim str1 As String = String.Concat(str, ItemGroupName & ".txt")
    '            Dim str1 As String = String.Concat(str, "PLU.txt")
    '            If (Time = 1) Then
    '                If (Not File.Exists(str1)) Then
    '                    fileStream = New System.IO.FileStream(str1, FileMode.Create, FileAccess.Write)
    '                    fileStream.Close()
    '                Else
    '                    File.Delete(str1)
    '                    fileStream = New System.IO.FileStream(str1, FileMode.Create, FileAccess.Write)
    '                    fileStream.Close()
    '                End If
    '            End If
    '            Dim streamWriter As System.IO.StreamWriter = New System.IO.StreamWriter(str1, True)
    '            If (Operators.CompareString(status, "", False) <> 0) Then
    '                streamWriter.WriteLine(status)
    '            End If
    '            streamWriter.Close()
    '        Catch exception1 As System.Exception
    '            objMain.objApplication.StatusBar.SetText("Text file generation error:" & exception1.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        End Try
    '    Finally
    '        GC.Collect()
    '        GC.WaitForPendingFinalizers()
    '    End Try
    'End Sub

    Private Sub FileGeneration1(ByVal pVal As SAPbouiCOM.ItemEvent)

        Dim str1 As String = Nothing
        Dim str2 As String = Nothing

        Dim businessObject As SAPbobsCOM.Recordset =
            DirectCast(objMain.objCompany.GetBusinessObject(300), SAPbobsCOM.Recordset)

        Dim recordset As SAPbobsCOM.Recordset =
            DirectCast(objMain.objCompany.GetBusinessObject(300), SAPbobsCOM.Recordset)

        Try
            Dim BranchID As String
            Dim ItemGroup As String
            Dim ItemGroupName As String
            Dim FilterType As String
            Dim ScaleType As String

            ' Branch
            Dim objBranchComb As SAPbouiCOM.ComboBox = objForm.Items.Item("Branch").Specific
            BranchID = objBranchComb.Selected.Value

            ' Item Group
            Dim objGroupCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Group").Specific
            ItemGroup = objGroupCombo.Selected.Value
            ItemGroupName = objGroupCombo.Selected.Description

            ' Filter Type
            Dim objFilterCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Date").Specific
            Try
                FilterType = objFilterCombo.Selected.Value
            Catch
                FilterType = "2"
            End Try

            ' Scale Type
            Dim objScaleCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("SType").Specific
            ScaleType = objScaleCombo.Selected.Value

            objMain.objApplication.StatusBar.SetText("Please wait... File generation in progress",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            ' Procedure Call
            Dim query As String =
                "CALL ""PriceList_TextFile_Generation_20251017""('" &
                BranchID & "','" & ItemGroup & "','" & FilterType & "','" & ScaleType & "')"

            businessObject.DoQuery(query)

            If businessObject.RecordCount <= 0 Then
                objMain.objApplication.StatusBar.SetText("No data found",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Exit Sub
            End If

            ' Get File Pathmadhu
            recordset.DoQuery("SELECT IFNULL(""U_FileLoc"", '') AS ""Location"" FROM OBPL WHERE ""BPLId"" = '" & BranchID & "'")
            recordset.MoveFirst()

            str1 = recordset.Fields.Item("Location").Value.ToString()

            ' ✅ Generate file name ONLY ONCE
            Dim timeStamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")

            Dim fileSuffix As String

            Select Case ScaleType
                Case "Internal Scale"
                    fileSuffix = "Internal_Scale"

                Case "Tenant Scale"
                    fileSuffix = "Tenant_Scale"

                Case "All Scale"
                    fileSuffix = "All_Scale"

                Case Else
                    fileSuffix = ScaleType
            End Select

            Dim finalFileName As String = ItemGroupName & "_" & fileSuffix & "_" & timeStamp

            ' Loop
            Do While Not businessObject.EoF

                Dim rateValue As Decimal = Convert.ToDecimal(businessObject.Fields.Item("Rate").Value)
                str2 = rateValue.ToString("F3")

                Dim line As String = String.Concat({
                    businessObject.Fields.Item("FLAG").Value.ToString(), ",",
                    businessObject.Fields.Item("EANCode").Value.ToString(), ",",
                    businessObject.Fields.Item("PLU").Value.ToString(), ",",
                    businessObject.Fields.Item("ItemName").Value.ToString(), ",",
                    businessObject.Fields.Item("arabic name").Value.ToString(), ",",
                     businessObject.Fields.Item("Department").Value.ToString(), ",",
                    businessObject.Fields.Item("Expire Days").Value.ToString()
                })

                Dim pathArray() As String = str1.Split(";"c)

                For Each path As String In pathArray
                    path = path.Trim()
                    If path <> "" Then
                        write_log(line, path, finalFileName)
                    End If
                Next

                businessObject.MoveNext()

            Loop

            objMain.objApplication.StatusBar.SetText("Price Detail text file generation completed successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error: " & ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

Private Sub FileGeneration(ByVal pVal As SAPbouiCOM.ItemEvent)
    Dim businessObject As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(300), SAPbobsCOM.Recordset)
    Dim recordset As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(300), SAPbobsCOM.Recordset)

    Try
        Dim BranchID As String = DirectCast(objForm.Items.Item("Branch").Specific, SAPbouiCOM.ComboBox).Selected.Value
        Dim objGroupCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Group").Specific
        Dim ItemGroup As String = objGroupCombo.Selected.Value
        Dim ItemGroupName As String = objGroupCombo.Selected.Description
        
        Dim FilterType As String = "2"
        Try : FilterType = DirectCast(objForm.Items.Item("Date").Specific, SAPbouiCOM.ComboBox).Selected.Value : Catch : End Try
        
        Dim ScaleType As String = DirectCast(objForm.Items.Item("SType").Specific, SAPbouiCOM.ComboBox).Selected.Value

        objMain.objApplication.StatusBar.SetText("Please wait... File generation in progress", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        businessObject.DoQuery($"CALL ""PriceList_TextFile_Generation_20251017""('{BranchID}','{ItemGroup}','{FilterType}','{ScaleType}')")

        If businessObject.RecordCount <= 0 Then
            objMain.objApplication.StatusBar.SetText("No data found", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Exit Sub
        End If

        recordset.DoQuery($"SELECT IFNULL(""U_FileLoc"", '') AS ""Location"" FROM OBPL WHERE ""BPLId"" = '{BranchID}'")
        recordset.MoveFirst()
        Dim str1 As String = recordset.Fields.Item("Location").Value.ToString()

        Dim fileSuffix As String = If(ScaleType.EndsWith("Scale"), ScaleType.Replace(" ", "_"), ScaleType)
            'Dim finalFileName As String = $"{ItemGroupName}_{fileSuffix}_{DateTime.Now:yyyyMMdd_HHmmss}"
            Dim finalFileName As String = "CAS"
            Dim pathArray() As String = str1.Split(";"c)
            '================ DELETE OLD FILE ==================
            'Delete existing file before generating a new one
            Try
                For Each path As String In pathArray

                    If Not String.IsNullOrWhiteSpace(path) Then

                        Dim filePath As String = IO.Path.Combine(path.Trim(), finalFileName & ".csv")

                        If IO.File.Exists(filePath) Then
                            IO.File.Delete(filePath)
                        End If

                    End If

                Next
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
            '===================================================
            For Each path As String In pathArray

                If String.IsNullOrWhiteSpace(path) Then Continue For

                Dim fullPath As String = IO.Path.Combine(path.Trim(), finalFileName & ".csv")

                Using sw As New IO.StreamWriter(fullPath, False, New System.Text.UTF8Encoding(False))

                    businessObject.MoveFirst()

                    Do While Not businessObject.EoF

                        Dim rateVal As Decimal = 0
                        Decimal.TryParse(businessObject.Fields.Item("Rate").Value.ToString(), rateVal)

                        Dim fields() As String = {
                GetVal(businessObject, "Department"),
                GetVal(businessObject, "PLU"),
                GetVal(businessObject, "PLU Type"),
                GetVal(businessObject, "ItemCode"),
                GetVal(businessObject, "ItemName"),
                rateVal.ToString("0.000"),
                GetVal(businessObject, "Expire Days")
            }

                        sw.WriteLine(String.Join(",", fields))

                        businessObject.MoveNext()

                    Loop

                End Using

            Next


            '            Do While Not businessObject.EoF
            '            Dim rateVal As Decimal = 0
            '            Decimal.TryParse(businessObject.Fields.Item("Rate").Value?.ToString(), rateVal)
            '                Dim fields() As String = {
            '                GetVal(businessObject, "Department"),
            '    GetVal(businessObject, "PLU"),
            '    GetVal(businessObject, "PLU Type"),
            '    GetVal(businessObject, "ItemCode"),
            '    GetVal(businessObject, "ItemName"),
            '    rateVal.ToString("0.000"),
            '    GetVal(businessObject, "Expire Days")
            '}

            '                'GetVal(businessObject, "FLAG"),
            '                'GetVal(businessObject, "EANCode"),
            '                'GetVal(businessObject, "PLU"),
            '                'GetVal(businessObject, "arabic name"),
            '                'GetVal(businessObject, "ItemName"),
            '                'ChrW(&H202A) & rateVal.ToString("0.000"),
            '                'GetVal(businessObject, "Department"),
            '                'GetVal(businessObject, "Expire Days") & ChrW(&H202C)
            '                Dim line As String = String.Join(",", fields)

            '            For Each path As String In pathArray
            '                If Not String.IsNullOrWhiteSpace(path) Then write_log(line, path.Trim(), finalFileName)
            '            Next

            '            businessObject.MoveNext()
            '        Loop

            objMain.objApplication.StatusBar.SetText("Price Detail text file generation completed successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

    Catch ex As Exception
        objMain.objApplication.StatusBar.SetText("Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    End Try
End Sub

Private Function GetVal(ByVal rs As SAPbobsCOM.Recordset, ByVal fieldName As String) As String
    Try
        Return If(rs.Fields.Item(fieldName).Value?.ToString().Trim(), "")
    Catch
        Return ""
    End Try
End Function
    Public Sub write_log(ByVal content As String,
                     ByVal path As String,
                     ByVal fileName As String)
        Try
            If Not IO.Directory.Exists(path) Then
                IO.Directory.CreateDirectory(path)
            End If

            Dim fullPath As String = IO.Path.Combine(path, fileName & ".csv")

            Using sw As New IO.StreamWriter(fullPath, True, New System.Text.UTF8Encoding(False))
                sw.WriteLine(content)
            End Using

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Write error: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    ' Public Sub write_log(ByVal content As String, ByVal path As String, ByVal fileName As String)
    '     Try
    '         If Not IO.Directory.Exists(path) Then IO.Directory.CreateDirectory(path)
    '         Dim fullPath As String = IO.Path.Combine(path, fileName & ".csv")

    '         Using sw As New IO.StreamWriter(fullPath, True, New System.Text.UTF8Encoding(False))
    '             sw.WriteLine(content)
    '         End Using
    '     Catch ex As Exception
    '         objMain.objApplication.StatusBar.SetText("Write error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '     End Try
    ' End Sub
    'Public Sub write_log1(ByVal content As String, ByVal path As String, ByVal fileName As String)

    '    Try
    '        If Not IO.Directory.Exists(path) Then
    '            IO.Directory.CreateDirectory(path)
    '        End If

    '        Dim fullPath As String = IO.Path.Combine(path, fileName & ".txt")

    '        Using sw As New IO.StreamWriter(fullPath, True)
    '            sw.WriteLine(content)
    '        End Using

    '        'Using sw As New StreamWriter(fullPath, True, New UTF8Encoding(False))
    '        '    sw.WriteLine(content)
    '        'End Using


    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText("Write error: " & ex.Message,
    '        SAPbouiCOM.BoMessageTime.bmt_Short,
    '        SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '    End Try

    'End Sub

    'Public Sub write_log(ByVal content As String, ByVal path As String, ByVal fileName As String)
    '    Try
    '        If Not IO.Directory.Exists(path) Then IO.Directory.CreateDirectory(path)
    '        Dim fullPath As String = IO.Path.Combine(path, fileName & ".txt")

    '        ' UTF8Encoding(False) prevents BOM issues while preserving Arabic text correctly
    '        Using sw As New IO.StreamWriter(fullPath, True, New System.Text.UTF8Encoding(False))
    '            sw.WriteLine(content)
    '        End Using
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText("Write error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '    End Try
    'End Sub
    Private Function GetCsvVal(ByVal rs As SAPbobsCOM.Recordset, ByVal fieldName As String) As String
        Try
            Dim val As String = If(rs.Fields.Item(fieldName).Value?.ToString().Trim(), "")
            If val.Contains(",") OrElse val.Contains("""") OrElse val.Contains(vbCr) OrElse val.Contains(vbLf) Then
                Return """" & val.Replace("""", """""") & """"
            End If
            Return val
        Catch
            Return ""
        End Try
    End Function

End Class
