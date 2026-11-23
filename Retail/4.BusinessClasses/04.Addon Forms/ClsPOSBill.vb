


Imports System.Drawing.Drawing2D
Imports System.Data
Imports System.Data.OleDb
Imports SAPbobsCOM

Public Class ClsPOSBill

#Region "Declaration"

    Public objForm As SAPbouiCOM.Form
    Dim objContentMx, objMatrix As SAPbouiCOM.Matrix

    Private WithEvents BackgroundWorker_CSVLoad As New System.ComponentModel.BackgroundWorker
    Private WithEvents BackgroundWorker_PriceUpdateService As New System.ComponentModel.BackgroundWorker

    Dim objComboBox As SAPbouiCOM.ComboBox

    Public oDBs_Head, oDBs_Details, oDBs_HeaderF, oDBs_Content, oDBs_Summary, oDBs_Attacht, oDS, oDS1 As SAPbouiCOM.DBDataSource

    Public Prgbar As SAPbouiCOM.ProgressBar
    Dim objutilities As Utilities

#End Region

#Region "Create Form"

    Sub CreateForm(ByVal DocEntry As String, ByVal ApplicationIDPOS As String)
        Try
            objMain.objUtilities.LoadForm("POSBill.xml", "POSBILL", ResourceType.Embeded)

            objForm = objMain.objApplication.Forms.GetForm("POSBILL", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.EnableMenu("774", True)
            objForm.EnableMenu("1292", True)
            objForm.Freeze(True)

            objutilities = New Utilities
            oDBs_HeaderF = objForm.DataSources.DBDataSources.Item("@POSBIL")
            oDBs_Content = objForm.DataSources.DBDataSources.Item("@POSB_C0")
            objContentMx = objForm.Items.Item("Matrix_1").Specific

            If IsAppIDExists(ApplicationIDPOS) Then
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                CType(objForm.Items.Item("AppID").Specific, SAPbouiCOM.EditText).Value = ApplicationIDPOS
                objForm.Items.Item("1").Click()
            Else
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE

                Dim newDocNum As String = objMain.objUtilities.GetNextDocNum(objForm, "POSBILLUDO", "Primary")

                oDBs_HeaderF.SetValue("DocNum", 0, newDocNum)
                oDBs_HeaderF.SetValue("U_DocD", 0, DateTime.Now.ToString("yyyyMMdd"))
                oDBs_HeaderF.SetValue("U_DOCE", 0, DocEntry)
                oDBs_HeaderF.SetValue("U_AppID", 0, ApplicationIDPOS)

                If objContentMx.VisualRowCount = 0 Then
                    Me.SetDefault(objForm.UniqueID)
                End If
            End If
            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("DocD").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocD").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocD").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocD").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Freeze(False)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

#End Region

#Region "Check AppID Exists"

    Private Function IsAppIDExists(ByVal AppID As String) As Boolean
        Try
            Dim rs As SAPbobsCOM.Recordset =
                objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim sql As String =
                "SELECT ""DocEntry"" FROM ""@POSBIL"" WHERE ""U_AppID"" = '" & AppID & "'"

            rs.DoQuery(sql)

            Return rs.RecordCount > 0

        Catch
            Return False
        End Try
    End Function

#End Region

#Region "Menu Event"

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try


            If pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)

            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then

                Me.SetDefault(objForm.UniqueID)

            ElseIf pVal.MenuUID = "774" And pVal.BeforeAction = False Then
                objMatrix = objForm.Items.Item("Matrix_1").Specific
                Dim row As Integer = objMatrix.VisualRowCount
                If objMatrix.IsRowSelected(1) <> True And objMatrix.VisualRowCount < 1 Then
                    objMatrix.AddRow()

                    oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)

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

                    objMatrix.Columns.Item("LineId").Cells.Item(i).Specific.Value = i

                Next

                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then

                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                End If

            End If



            If pVal.MenuUID = "519" And pVal.BeforeAction = True Then
                Try
                    Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oDBs_HeaderF = objForm.DataSources.DBDataSources.Item("@POSBIL")
                    Dim PosBill As String = "PosBill"
                    Dim query As String = "SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = '" & PosBill & "'"
                    rs.DoQuery(query)
                    If rs.RecordCount = 0 Then
                        objMain.objApplication.MessageBox("Report layout not found.", 0, "OK")
                    Else
                        objMain.objApplication.ActivateMenuItem(rs.Fields.Item(0).Value.ToString())
                        Dim CrForm As SAPbouiCOM.Form = objMain.objApplication.Forms.ActiveForm
                        Dim oedt As SAPbouiCOM.EditText = CrForm.Items.Item("1000003").Specific
                        oedt.Value = oDBs_HeaderF.GetValue("DocEntry", 0).Trim()
                        CrForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If

                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText(ex.Message)
                End Try
            End If


            If pVal.MenuUID = "520" And pVal.BeforeAction = True Then
                Try
                    Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oDBs_HeaderF = objForm.DataSources.DBDataSources.Item("@POSBIL")
                    Dim PosBill As String = "PosBill"
                    Dim query As String = "SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = '" & PosBill & "'"
                    rs.DoQuery(query)
                    If rs.RecordCount = 0 Then
                        objMain.objApplication.MessageBox("Report layout not found.", 0, "OK")
                    Else
                        objMain.objApplication.ActivateMenuItem(rs.Fields.Item(0).Value.ToString())
                        Dim CrForm As SAPbouiCOM.Form = objMain.objApplication.Forms.ActiveForm
                        Dim oedt As SAPbouiCOM.EditText = CrForm.Items.Item("1000003").Specific
                        oedt.Value = oDBs_HeaderF.GetValue("DocEntry", 0).Trim()
                        CrForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If

                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText(ex.Message)
                End Try
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

#End Region

#Region "Item Event"

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)

        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                    If pVal.ItemUID = "Matrix_1" And pVal.ColUID = "LTAmo" Then
                        If pVal.Row = objContentMx.VisualRowCount Then
                            Dim val As String =
                                objContentMx.Columns.Item("BarC").Cells.Item(pVal.Row).Specific.Value

                            'If Not String.IsNullOrWhiteSpace(val) Then
                            '    Me.SetNewLineContent(objForm.UniqueID)
                            'End If
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And (objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.ActionSuccess = True Then

                        Try
                            objForm = objMain.objApplication.Forms.Item(FormUID)

                            oDBs_HeaderF = objForm.DataSources.DBDataSources.Item("@POSBIL")
                            oDBs_Content = objForm.DataSources.DBDataSources.Item("@POSB_C0")
                            objMatrix = objForm.Items.Item("Matrix_1").Specific

                            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            oRs.DoQuery("SELECT MAX(""DocNum"") FROM ""@POSBIL""")

                            Dim lastDocNum As String = ""
                            If oRs.RecordCount > 0 Then
                                lastDocNum = oRs.Fields.Item(0).Value.ToString()
                            End If

                            If Not String.IsNullOrWhiteSpace(lastDocNum) Then

                                objForm.Freeze(True)

                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                CType(objForm.Items.Item("DocNum").Specific, SAPbouiCOM.EditText).Value = lastDocNum
                                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                objForm.Freeze(False)
                                objMain.objApplication.StatusBar.SetText("Document " & lastDocNum & " added and reloaded.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                            End If

                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try

                    End If

                    ' EXCEL LOAD
                    If pVal.ItemUID = "Item_6" AndAlso pVal.BeforeAction = True Then

                        Dim filePath As String = ""

                        Try
                            filePath = objForm.Items.Item("Item_3").Specific.value
                        Catch
                        End Try

                        If Not String.IsNullOrWhiteSpace(filePath) Then

                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@POSBIL")
                            oDBs_Head.SetValue("U_EXCL", 0, filePath)

                            Me.LoadExcelToMatrix(objForm, filePath)

                        Else
                            objMain.objApplication.StatusBar.SetText("Please select Excel file")
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Content = objForm.DataSources.DBDataSources.Item("@POSB_C0")
                    objContentMx = objForm.Items.Item("Matrix_1").Specific

                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    oCFL = objForm.ChooseFromLists.Item(CFLEvent.ChooseFromListUID)

                    Dim oDT As SAPbouiCOM.DataTable = CFLEvent.SelectedObjects

                    If (Not oDT Is Nothing) And pVal.BeforeAction = False Then

                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        End If

                        If oCFL.UniqueID = "CFL_BC" Then
                            oDBs_Content.SetValue("U_BarC", oDBs_Content.Offset, oDT.GetValue("BcdCode", 0))
                            oDBs_Content.SetValue("U_DisENG", oDBs_Content.Offset, oDT.GetValue("BcdName", 0))
                            oDBs_Content.SetValue("U_DisABC", oDBs_Content.Offset, oDT.GetValue("U_ArbName", 0))
                            oDBs_Content.SetValue("U_FRNM", oDBs_Content.Offset, oDT.GetValue("U_ArbName", 0))
                            oDBs_Content.SetValue("U_SlNo", oDBs_Content.Offset, pVal.Row.ToString())


                            objContentMx.SetLineData(pVal.Row)
                        End If

                    End If

            End Select

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub

#End Region

#Region "Excel Load SAFE FIX"

    Private Sub LoadExcelToMatrix(objForm As SAPbouiCOM.Form, filePath As String)
        Try

            Dim dt As New DataTable()

            Dim connStr As String =
                "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath &
                ";Extended Properties='Excel 12.0 Xml;HDR=YES';"

            Using conn As New OleDbConnection(connStr)
                conn.Open()
                Dim cmd As New OleDbCommand("SELECT * FROM [Sheet1$]", conn)
                Dim adapter As New OleDbDataAdapter(cmd)
                adapter.Fill(dt)
                conn.Close()
            End Using

            objMatrix = objForm.Items.Item("Matrix_1").Specific
            oDBs_Content = objForm.DataSources.DBDataSources.Item("@POSB_C0")

            objForm.Freeze(True)

            objMatrix.Clear()
            oDBs_Content.Clear()

            For i As Integer = 0 To dt.Rows.Count - 1

                oDBs_Content.InsertRecord(i)

                oDBs_Content.SetValue("LineId", i, (i + 1).ToString())

                SafeSet("U_SlNo", i, (i + 1).ToString())
                SafeSet("U_BarC", i, dt.Rows(i)("BarCode").ToString())
                SafeSet("U_DisENG", i, dt.Rows(i)("Description (English)").ToString())
                SafeSet("U_DisABC", i, dt.Rows(i)("Description (Arabic)").ToString())
                SafeSet("U_PQty", i, dt.Rows(i)("Purchase Quantity").ToString())
                SafeSet("U_VAT", i, dt.Rows(i)("VAT").ToString())
                SafeSet("U_LTAmo", i, dt.Rows(i)("Line Total Amount").ToString())
                SafeSet("U_UPRICE", i, dt.Rows(i)("Unit Price").ToString())
                SafeSet("U_FRNM", i, dt.Rows(i)("Foreign Name").ToString())

            Next

            objMatrix.LoadFromDataSource()
            objMatrix.AutoResizeColumns()

            objForm.Freeze(False)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.MessageBox("Excel Error: " & ex.Message)
        End Try
    End Sub


    Private Sub SafeSet(field As String, row As Integer, value As String)
        Try
            If oDBs_Content.Fields.Item(field) IsNot Nothing Then
                oDBs_Content.SetValue(field, row, value)
            End If
        Catch
        End Try
    End Sub

#End Region

#Region "Set New Line"

    Sub SetNewLineContent(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Content = objForm.DataSources.DBDataSources.Item("@POSB_C0")
            objContentMx = objForm.Items.Item("Matrix_1").Specific

            'If objContentMx.VisualRowCount > 0 Then
            '    Dim lastVal As String =
            '        objContentMx.Columns.Item("BarC").Cells.Item(objContentMx.VisualRowCount).Specific.Value

            '    If String.IsNullOrWhiteSpace(lastVal) Then Exit Sub
            'End If

            objContentMx.AddRow()
            oDBs_Content.SetValue("LineId", oDBs_Content.Offset, objContentMx.VisualRowCount)
            oDBs_Content.SetValue("U_SlNo", oDBs_Content.Offset, "")
            oDBs_Content.SetValue("U_BarC", oDBs_Content.Offset, "")
            oDBs_Content.SetValue("U_DisENG", oDBs_Content.Offset, "")
            oDBs_Content.SetValue("U_DisABC", oDBs_Content.Offset, "")
            oDBs_Content.SetValue("U_PQty", oDBs_Content.Offset, "")
            oDBs_Content.SetValue("U_VAT", oDBs_Content.Offset, "")
            oDBs_Content.SetValue("U_LTAmo", oDBs_Content.Offset, "")
            oDBs_Content.SetValue("U_UPRICE", oDBs_Content.Offset, "")
            oDBs_Content.SetValue("U_FRNM", oDBs_Content.Offset, "")

            objContentMx.SetLineData(objContentMx.VisualRowCount)
            objContentMx.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

#End Region

#Region "Set Default"

    Sub SetDefault(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_HeaderF = objForm.DataSources.DBDataSources.Item("@POSBIL")
            oDBs_Content = objForm.DataSources.DBDataSources.Item("@POSB_C0")

            objContentMx = objForm.Items.Item("Matrix_1").Specific

            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                If String.IsNullOrWhiteSpace(oDBs_HeaderF.GetValue("DocNum", 0).Trim()) Then
                    oDBs_HeaderF.SetValue("DocNum", 0,
                        objMain.objUtilities.GetNextDocNum(objForm, "POSBILLUDO", "Primary"))
                End If

            End If

            oDBs_HeaderF.SetValue("U_DocD", 0, DateTime.Now.ToString("yyyyMMdd"))

            'If objContentMx.VisualRowCount > 0 Then
            Me.SetNewLineContent(FormUID)
            'End If

            objContentMx.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


#End Region

End Class
