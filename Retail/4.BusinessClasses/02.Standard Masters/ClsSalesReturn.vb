Imports SAPbobsCOM

Public Class ClsSalesReturn

    Public objForm, objJVForm As SAPbouiCOM.Form
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDs, oDs1, oDs2, oDs3 As SAPbouiCOM.DBDataSource
    Dim objMatrix, objMatrix1 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim oDt As SAPbouiCOM.DataTable
    Dim str, str1 As String
    Public rs, frs, RsNum As SAPbobsCOM.Recordset
    Dim LostFocusFlag As Boolean = False
    Dim objutilities As Utilities
    Dim oItem, oItem1, oItem2, oItem3 As SAPbouiCOM.Item
    Dim FormID As String

    Dim PrgBar As SAPbouiCOM.ProgressBar

    Dim ButtonCalFlag As Boolean = False
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType





                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    FormID = FormUID
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("ORDN")
                    objMatrix = objForm.Items.Item("38").Specific
                    If pVal.BeforeAction = True Then
                        Me.AddItems(objForm.UniqueID)
                    End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix = objForm.Items.Item("38").Specific
                    objComboBox = objForm.Items.Item("3").Specific
                    Dim TYPE As String = objComboBox.Value
                    If pVal.ItemUID = "CLEX" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And TYPE = "I" Then
                        Try
                            objForm.Freeze(True)

                            ButtonCalFlag = True


                            For R As Integer = 1 To objMatrix.VisualRowCount

                                If objMatrix.Columns.Item("1").Cells.Item(R).Specific.Value <> "" Then

                                    Dim GetPerPCVal As Double = 0

                                    Dim GetPerPC As String = "SELECT ""BaseQty""  FROM  ""OUGP""  T0 INNER JOIN  ""UGP1""  T1 ON T0.""UgpEntry"" = T1.""UgpEntry"" " &
                                                            " INNER JOIN  ""OITM""  T2 ON T0.""UgpEntry"" = T2.""UgpEntry"" WHERE T2.""ItemCode"" ='" & CStr(objMatrix.Columns.Item("1").Cells.Item(R).Specific.Value).Trim & "' " &
                                                            " And ""UomEntry"" IN (SELECT T0.""UomEntry"" FROM OUOM T0 WHERE T0.""UomCode"" ='" & CStr(objMatrix.Columns.Item("1470002145").Cells.Item(R).Specific.Value).Trim & "') "

                                    Dim oRsGetPerPCas As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                                    oRsGetPerPCas.DoQuery(GetPerPC)

                                    If oRsGetPerPCas.RecordCount > 0 Then
                                        GetPerPCVal = oRsGetPerPCas.Fields.Item(0).Value
                                    End If


                                    Dim Discount As Double = 0
                                    Dim Total As Double = 0


                                    If CStr(objMatrix.Columns.Item("15").Cells.Item(R).Specific.Value).Trim <> "" Then
                                        Discount = CDbl(CStr(objMatrix.Columns.Item("15").Cells.Item(R).Specific.Value).Trim)
                                    End If


                                    If CStr(objMatrix.Columns.Item("21").Cells.Item(R).Specific.Value).Trim <> "" Then
                                        Total = String.Concat(CStr(objMatrix.Columns.Item("21").Cells.Item(R).Specific.Value).Trim.Where(Function(c) Char.IsDigit(c))) 'CDbl(CStr(objMatrix.Columns.Item("21").Cells.Item(R).Specific.Value).Trim)
                                    End If

                                    'If Total = 0 Then
                                    '    'If Discount = 100 Or Total = 0 Then

                                    '    Dim EXPNM As String = " Select ""ExpnsCode"" from OEXD  where ""ExpnsName"" = 'Local Excise' "
                                    '    Dim oRsEXPNM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    '    oRsEXPNM.DoQuery(EXPNM)

                                    '    Try
                                    '        Dim ExpnsCode1 As String = oRsEXPNM.Fields.Item("ExpnsCode").Value

                                    '        Dim objComboBox8 As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("111").Cells.Item(R).Specific
                                    '        Dim TaskName8 As String = objComboBox8.Selected.Value

                                    '        objComboBox8.Select("-1", SAPbouiCOM.BoSearchKey.psk_ByValue)

                                    '        objMatrix.Columns.Item("U_EXCTAX").Cells.Item(R).Specific.Value() = 0
                                    '        objMatrix.Columns.Item("112").Cells.Item(R).Specific.Value = 0


                                    '    Catch ex As Exception

                                    '    End Try

                                    'Else

                                    Dim cardcode As String = objForm.Items.Item("4").Specific.Value
                                    Dim Getdata As String = "SELECT T0.""U_BPTYPE"",T0.""ExpnsCode"",T0.""ExpnsName"" FROM OEXD T0 inner join NNM1 T1 ON T1.""Series""=T0.""U_BPSERIES"" " +
                                                                    "  inner join OCRD T2 on T2.""Series""=T1.""Series"" where  T2.""CardCode""='" & cardcode & "'"
                                    Dim oRs1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRs1.DoQuery(Getdata)
                                    Dim ExpnsCode As String = oRs1.Fields.Item("ExpnsCode").Value
                                    Dim BPType As String = oRs1.Fields.Item("U_BPTYPE").Value

                                    If BPType = "LOCAL" Then 'And (customertype = "CUS" Or customertype = "Manual") Then
                                        Dim objComboBox3 As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("111").Cells.Item(R).Specific
                                        Dim TaskName3 As String = objComboBox3.Selected.Value
                                        objComboBox3.Select(ExpnsCode, SAPbouiCOM.BoSearchKey.psk_ByValue)

                                        objMatrix = objForm.Items.Item("38").Specific
                                        Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(R).Specific.Value
                                        Dim GetDocEntry As String = "Select ""U_EXCVAT"" From OITM Where ""ItemCode"" = '" & ItemCode & "' "
                                        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRs.DoQuery(GetDocEntry)

                                        objMatrix.Columns.Item("U_EXCTAX").Cells.Item(R).Specific.Value() = oRs.Fields.Item("U_EXCVAT").Value()

                                        Dim ItemCode1 As String = objMatrix.Columns.Item("1").Cells.Item(R).Specific.Value.ToString().Trim()
                                        Dim QTY As Double = objMatrix.Columns.Item("11").Cells.Item(R).Specific.Value.ToString().Trim()
                                        Dim EXCTAX As Double = objMatrix.Columns.Item("U_EXCTAX").Cells.Item(R).Specific.Value.ToString().Trim() * GetPerPCVal
                                        Dim SubTotal As String = "ST1"
                                        If (QTY * EXCTAX) <> 0 Then
                                            Dim ST1 As Double = QTY * EXCTAX
                                            SubTotal = ST1

                                            If objMatrix.Columns.Item("U_EXCTAX").Cells.Item(R).Specific.Value <> "" Then
                                                Dim objComboBox As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("111").Cells.Item(R).Specific
                                                Dim TaskName As String = objComboBox.Selected.Value
                                                ' objComboBox.Select("Sales Excise Duty", SAPbouiCOM.BoSearchKey.psk_ByDescription)
                                                objComboBox.Select(ExpnsCode, SAPbouiCOM.BoSearchKey.psk_ByValue)

                                                Dim objComboBox1 As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("111").Cells.Item(R).Specific
                                                Dim TaskName1 As String = objComboBox.Selected.Value
                                                If objMatrix.Columns.Item("111").Cells.Item(R).Specific.Value = TaskName1 Then
                                                    objMatrix.Columns.Item("112").Cells.Item(R).Specific.Value = SubTotal
                                                End If

                                            End If
                                        Else
                                            Dim objComboBox As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("111").Cells.Item(R).Specific
                                            Dim TaskName As String = objComboBox.Selected.Value
                                            objComboBox.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                        End If

                                    Else
                                        Dim EXPNM As String = " Select ""ExpnsCode"" from OEXD  where ""U_BPTYPE"" = 'FOREIGN' "
                                        Dim oRsEXPNM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRsEXPNM.DoQuery(EXPNM)
                                        Dim ExpnsCode1 As String = oRsEXPNM.Fields.Item("ExpnsCode").Value

                                        Dim objComboBox8 As SAPbouiCOM.ComboBox = objMatrix.Columns.Item("111").Cells.Item(R).Specific
                                        Dim TaskName8 As String = objComboBox8.Selected.Value
                                        objComboBox8.Select(ExpnsCode, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                        objMatrix = objForm.Items.Item("38").Specific
                                        Dim ItemCode As String = objMatrix.Columns.Item("1").Cells.Item(R).Specific.Value
                                        Dim GetDocEntry As String = "Select ""U_EXCVAT"" From OITM Where ""ItemCode"" = '" & ItemCode & "' "
                                        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRs.DoQuery(GetDocEntry)
                                        objMatrix.Columns.Item("U_EXCTAX").Cells.Item(R).Specific.Value() = oRs.Fields.Item("U_EXCVAT").Value()
                                    End If


                                End If


                                '    End If

                            Next


                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And TYPE = "I" Then
                        Try
                            If ButtonCalFlag = False Then
                                objMain.objApplication.StatusBar.SetText("Please click Calculate Excise button before adding")
                                BubbleEvent = False
                            End If
                        Catch ex As Exception

                        End Try
                    End If





            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub AddItems(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            oItem = objForm.Items.Add("CLEX", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Top = objForm.Items.Item("38").Top + objForm.Items.Item("38").Height + 10
            oItem.Left = objForm.Items.Item("38").Left + 100
            oItem.Width = objForm.Items.Item("2").Width + 60
            oItem.Height = objForm.Items.Item("2").Height
            oItem.FromPane = objForm.Items.Item("38").FromPane
            oItem.ToPane = objForm.Items.Item("38").ToPane
            Dim btn2 As SAPbouiCOM.Button = objForm.Items.Item("CLEX").Specific

            btn2.Caption = "Calculate Excise"

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType


                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        ButtonCalFlag = False
                    End If
            End Select

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
