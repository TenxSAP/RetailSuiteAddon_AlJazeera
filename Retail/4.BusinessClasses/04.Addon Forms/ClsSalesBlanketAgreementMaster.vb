Public Class ClsSalesBlanketAgreementMaster

    Public objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix
    Public rs As SAPbobsCOM.Recordset
    Dim objutilities As Utilities
    Private AbsIDSalesBlanketScreen As String

    '👉 GLOBAL STORAGE (FormUID → AbsID)
    Public Shared AbsIDMap As New Dictionary(Of String, String)

    Public Sub CreateForm()
        CreateForm("")
    End Sub

#Region "Create Form"
    Sub CreateForm(ByVal AbsID As String)
        Try
            AbsIDSalesBlanketScreen = AbsID

            objMain.objUtilities.LoadForm("SalesBlanketAgreementMaster.xml", "SBAMASTER", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("SBAMASTER", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            '👉 STORE AbsID against FormUID
            If AbsIDMap.ContainsKey(objForm.UniqueID) Then
                AbsIDMap(objForm.UniqueID) = AbsID
            Else
                AbsIDMap.Add(objForm.UniqueID, AbsID)
            End If

            objutilities = New Utilities
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_SBAMRM")
            objMatrix = CType(objForm.Items.Item("Mtx").Specific, SAPbouiCOM.Matrix)

            objForm.Freeze(False)

            Me.Matrixload(objForm.UniqueID)
            objForm.EnableMenu("774", True)
            objForm.EnableMenu("1292", True)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_SBAMRM")
            objMatrix = objForm.Items.Item("Mtx").Specific

            objMatrix.AddRow()

            oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.VisualRowCount)
            oDBs_Head.SetValue("Code", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("Name", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_Status", oDBs_Head.Offset, "")

            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "SBAMASTER" And pVal.BeforeAction = False Then
                CreateForm()

            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("SBAMASTER", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix = objForm.Items.Item("Mtx").Specific
                SetNewLine(objForm.UniqueID)

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
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    '-----------Madhu 18.05.2025 ----------------
                    If pVal.ItemUID = "1" AndAlso pVal.ActionSuccess = True Then
                        ' Close the Base Form
                        objForm.Close()
                    End If

            End Select

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

#Region "Form Data Event"
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("SBAMASTER", objMain.objApplication.Forms.ActiveForm.TypeCount)

            Select Case BusinessObjectInfo.EventType

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD,
                     SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                    If BusinessObjectInfo.BeforeAction = False AndAlso BusinessObjectInfo.ActionSuccess = True Then

                        objForm.Freeze(True)

                        Try
                            objMatrix = objForm.Items.Item("Mtx").Specific

                            Dim oRS As SAPbobsCOM.Recordset
                            oRS = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            Dim AbsIDValue As String = ""

                            If AbsIDMap.ContainsKey(objForm.UniqueID) Then
                                AbsIDValue = AbsIDMap(objForm.UniqueID)
                            End If

                            For i As Integer = 1 To objMatrix.VisualRowCount
                                Dim codeVal As String = objMatrix.Columns.Item("Code").Cells.Item(i).Specific.Value

                                If codeVal <> "" AndAlso AbsIDValue <> "" Then
                                    Dim strSQL As String =
                                        "UPDATE OOAT SET ""U_TypeDis"" = '" & codeVal & "' WHERE ""AbsID"" = '" & AbsIDValue & "'"

                                    oRS.DoQuery(strSQL)
                                End If
                            Next


                            Dim oFormOOAT As SAPbouiCOM.Form = objMain.objApplication.Forms.GetForm("1250000100", 1)
                            oFormOOAT.Select()

                            objMain.objApplication.ActivateMenuItem("1304")


                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Finally
                            objForm.Freeze(False)
                        End Try

                    End If

            End Select

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

    Public Sub Matrixload(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim rs1 As String = "select ""DocEntry"", ""Code"", ""Name"", ""U_Status"" from ""@TNX_SBAMRM"" order by ""DocEntry"" asc"

            Dim ors1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ors1.DoQuery(rs1)

            If ors1.RecordCount > 0 Then
                For j As Integer = 1 To ors1.RecordCount
                    objMatrix.AddRow()

                    oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.VisualRowCount)
                    oDBs_Head.SetValue("Code", oDBs_Head.Offset, ors1.Fields.Item("Code").Value)
                    oDBs_Head.SetValue("Name", oDBs_Head.Offset, ors1.Fields.Item("Name").Value)
                    oDBs_Head.SetValue("U_Status", oDBs_Head.Offset, ors1.Fields.Item("U_Status").Value)

                    objMatrix.SetLineData(objMatrix.VisualRowCount)
                    ors1.MoveNext()
                Next
            End If

            objMatrix.AutoResizeColumns()
            objMatrix.FlushToDataSource()
            objMatrix.LoadFromDataSource()
            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

        Catch ex As Exception
            Me.objForm.Freeze(False)
            Me.objForm.StatusBar.SetText(ex.Message & "   Page Load Events", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

End Class

