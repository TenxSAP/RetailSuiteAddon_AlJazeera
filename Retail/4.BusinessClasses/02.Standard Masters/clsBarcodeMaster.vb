Imports SAPbobsCOM
Imports SAPbouiCOM
Public Class clsBarcodeMaster
    Public objform, objListForm As SAPbouiCOM.Form
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oItem As SAPbouiCOM.Item
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim StrQry, oAutoBarcode As String
    Dim oRec As SAPbobsCOM.Recordset
    Dim Startt As Int64 = 0
    Dim Endd As Integer = 0
    Dim objMatrix, objBarcodeMatrix As SAPbouiCOM.Matrix
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
                    If pVal.ItemUID = "Btn_Gen" And pVal.BeforeAction = False Then


                        Me.Generate_Barcode(objform.UniqueID, objMain.InteNum)
                        'Me.StandardBarcodeGenerate()
                    End If


                    If pVal.ItemUID = "Btn_Gen1" And pVal.Before_Action = False Then
                        'Dim ofor As SAPbouiCOM.Form
                        ' objform = New cls_GetDetails()
                        Dim itemcode As String = objform.Items.Item("1470000005").Specific.Value
                        'Dim itemcode As String = CType(objform.Items.Item("ItemCode").Specific, SAPbouiCOM.EditText).Value
                        objMain.objGet.CreateForm(itemcode)


                        'ofor = Nothing
                        'Me.GetDetails(objform.UniqueID)
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub AddItems(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)

            oItem = objform.Items.Add("Btn_Gen", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Top = objform.Items.Item("1470000002").Top
            oItem.Left = objform.Items.Item("1470000002").Left + objform.Items.Item("1470000002").Width + 5
            oItem.Width = objform.Items.Item("1470000002").Width + 30
            oItem.Height = objform.Items.Item("1470000002").Height
            Dim btn1 As SAPbouiCOM.Button = objform.Items.Item("Btn_Gen").Specific
            btn1.Caption = "Generate Barcode"


            'Akhilas
            oItem = objform.Items.Add("Btn_Gen1", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Top = objform.Items.Item("Btn_Gen").Top
            oItem.Left = objform.Items.Item("Btn_Gen").Left + objform.Items.Item("Btn_Gen").Width + 5
            oItem.Width = objform.Items.Item("Btn_Gen").Width + 30
            oItem.Height = objform.Items.Item("Btn_Gen").Height
            Dim btn12 As SAPbouiCOM.Button = objform.Items.Item("Btn_Gen1").Specific
            btn12.Caption = "Get Details"




        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    'akhila

    Sub GetDetails(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)
            Dim CP As SAPbouiCOM.FormCreationParams = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            Dim oItem As SAPbouiCOM.Item
            Dim oGrid As SAPbouiCOM.Grid
            Dim oButton As SAPbouiCOM.Button

            CP.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            CP.UniqueID = "GETDE" & objMain.objApplication.Forms.Count
            CP.FormType = "GETDE"

            objListForm = objMain.objApplication.Forms.AddEx(CP)
            objListForm.Height = 320
            objListForm.Width = 600
            objListForm.Title = "Get Details"
            oItem = objListForm.Items.Add("Grid", SAPbouiCOM.BoFormItemTypes.it_GRID)
            oItem.Left = 20
            oItem.Top = 20
            oItem.Width = 550
            oItem.Height = 200
            oGrid = oItem.Specific
            objListForm.DataSources.DataTables.Add("MyDataTable")

            oItem = objListForm.Items.Add("ok", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 20
            oItem.Top = 230
            oItem.Width = 65
            oItem.Height = 20
            oButton = oItem.Specific
            oButton.Caption = "Ok"

            oItem = objListForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 90
            oItem.Top = 230
            oItem.Width = 65
            oItem.Height = 20
            oButton = oItem.Specific
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto
            Dim sQuery As String = "SELECT * FROM ""OBCD"""

            objListForm.DataSources.DataTables.Item(0).ExecuteQuery(sQuery)





            objListForm.DataSources.DataTables.Item(0).ExecuteQuery(sQuery)


            oGrid.DataTable = objListForm.DataSources.DataTables.Item("MyDataTable")
            oGrid.AutoResizeColumns()
            If oGrid.Rows.Count = 0 Then
                objMain.objApplication.MessageBox("No records found")
                Exit Sub
            End If
            For i As Integer = 0 To oGrid.Columns.Count - 1
                oGrid.Columns.Item(i).Editable = False
            Next
            objListForm.Visible = True
        Catch ex As Exception
            objListForm.Close()
            objMain.objApplication.StatusBar.SetText("Loading Branch Wise Pricelist Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub



    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objform = objMain.objApplication.Forms.GetForm("134", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        objComboBox = objform.Items.Item("40").Specific
                        If objComboBox.Value = "C" Then
                            objform.Items.Item("Btn_Gen").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                        Else
                            objform.Items.Item("Btn_Gen").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "1", SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        End If
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    Public Sub Generate_Barcode(ByVal FormUID As String, ByVal weightItem As Integer)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)
            Dim specific As Matrix = DirectCast(objMain.objApplication.Forms.ActiveForm.Items.Item("1470000013").Specific, Matrix)
            Dim businessObject As Recordset = DirectCast(objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset), Recordset)
            Dim item As String = objform.Items.Item("1470000005").Specific.Value
            Dim str1 As String = ""

            '---------------------------------------------------------------
            '---------------------------------------------------------------

            Dim DigitsQry As String = "SELECT Case When ""U_AcceptWMQty""='1' Then '6' When ""U_AcceptWMQty""='2' Then '13' End FROM OITM WHERE ""ItemCode"" = '" & item & "'"
            Dim oRsDigits As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRsDigits.DoQuery(DigitsQry)

            If weightItem <> 1 Then '------------13-------------------------------------

                'Dim ItemGroupQuery1 As String = "SELECT * FROM ""@NT_BRST""  T2 WHERE T2.""U_Itmgrpnm"" IN " &
                '"(SELECT (SELECT T1.""ItmsGrpNam"" FROM OITB T1 WHERE T1.""ItmsGrpCod""=T0.""ItmsGrpCod"") FROM OITM T0 WHERE T0.""ItemCode""='" & item & "') " &
                '" And  T2.""U_Scale""=(SELECT ""U_AcceptWMQty"" FROM ""OITM"" Where ""ItemCode""='" & item & "') and  T2.""U_SubGrp"" =(SELECT ""U_VSubGroup"" FROM ""OITM"" Where ""ItemCode""='" & item & "')"

                Dim ItemGroupQuery As String = "SELECT * FROM ""@NT_BRST""  T2 WHERE T2.""U_Itmgrpnm"" IN " &
                "(SELECT (SELECT T1.""ItmsGrpNam"" FROM OITB T1 WHERE T1.""ItmsGrpCod""=T0.""ItmsGrpCod"") FROM OITM T0 WHERE T0.""ItemCode""='" & item & "') " &
                " And  T2.""U_Scale""=(SELECT ""U_AcceptWMQty"" FROM ""OITM"" Where ""ItemCode""='" & item & "') "

                Dim oRsItemGroup As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                oRsItemGroup.DoQuery(ItemGroupQuery)
                'If oRsItemGroup.RecordCount = 0 Then
                '    Me.StandardBarcodeGenerate()
                'Else
                Dim NewBarcode As String = ""
                Dim ItemGroup As String = ""
                Dim SubGroup As String = ""
                If oRsItemGroup.RecordCount = 0 Then
                    ItemGroup = "ALL"
                Else
                    ItemGroup = oRsItemGroup.Fields.Item("U_Itmgrpnm").Value
                    '' SubGroup = oRsItemGroup.Fields.Item("U_SubGrp").Value
                End If
                Dim Barcode As String = "SELECT T0.""U_Prefix"", T0.""U_Itprefix"", T0.""U_Start"",T0.""U_Start""+1 AS ""NextNo"" FROM ""@NT_BRST""  T0 WHERE T0.""U_Itmgrpnm"" ='" & ItemGroup & "' and ""U_Scale""='2' "
                'Dim Barcode As String = "SELECT T0.""U_Prefix"", T0.""U_Itprefix"", T0.""U_Start"",T0.""U_Start""+1 AS ""NextNo"" FROM ""@NT_BRST""  T0 WHERE T0.""U_Itmgrpnm"" ='" & ItemGroup & "' and T0.""U_SubGrp"" ='" & SubGroup & "'and ""U_Scale""='2' "
                Dim oRsBarcode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                oRsBarcode.DoQuery(Barcode)
                NewBarcode = oRsBarcode.Fields.Item("U_Prefix").Value.ToString().Trim() & oRsBarcode.Fields.Item("U_Itprefix").Value.ToString().Trim() & oRsBarcode.Fields.Item("U_Start").Value.ToString().Trim()
                '------Checkdigit calculation-------------
                Dim str2 As String = Me.GetFieldValue("CALL ""EAN13_Check_Digit_Calculation""('" & NewBarcode & "')")
                If str2.Trim() = "" Then
                    objMain.objApplication.StatusBar.SetText("Check Digit should not be null. Please check the function - EAN13_Check_Digit_Calculation ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                Else
                    str1 = String.Concat((NewBarcode), str2)
                    If (Strings.Len(str1) <> 13) Then
                        objMain.objApplication.StatusBar.SetText("BarCode should be 13-Digit. Please check the function - EAN13_Check_Digit_Calculation ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                    Else
                        specific.Columns.Item("1470000003").Cells.Item(specific.VisualRowCount).Specific.value = str1.Trim()
                        'Dim str As String = objform.Items.Item("1470000006").Specific.Value
                        'specific.Columns.Item("1470000005").Cells.Item(specific.VisualRowCount).Specific.value = objform.Items.Item("1470000006").Specific.Value

                        If (objMain.objApplication.Forms.ActiveForm.Mode = BoFormMode.fm_UPDATE_MODE) Then
                            objMain.objApplication.Forms.ActiveForm.Items.Item("1470000001").Click(BoCellClickType.ct_Regular)
                        End If
                    End If
                End If
                '------Check digit calculation----------
                '------Updating Next Number-------------
                Dim NextNo As String = oRsBarcode.Fields.Item("NextNo").Value.ToString().Trim()
                If NextNo.Length < 9 Then
                    'For k As Integer = 1 To 9 - NextNo.Length
                    For k As Integer = 1 To 8 - NextNo.Length

                        NextNo = "0" & NextNo
                    Next
                End If
                'Dim NextBarcode As String = "UPDATE ""@NT_BRST"" T0 SET T0.""U_Start""='" & NextNo & "' WHERE T0.""U_Itmgrpnm"" ='" & ItemGroup & "' and T0.""U_SubGrp"" ='" & SubGroup & "' and ""U_Scale""='2' "
                Dim NextBarcode As String = "UPDATE ""@NT_BRST"" T0 SET T0.""U_Start""='" & NextNo & "' WHERE T0.""U_Itmgrpnm"" ='" & ItemGroup & "' and ""U_Scale""='2' "
                Dim oRsNextBarcode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                oRsNextBarcode.DoQuery(NextBarcode)

            Else
                'Dim str As String = "CALL ""SBO_ProcessBarCode""('" & item & "')"
                'businessObject.DoQuery(str)
                'If businessObject.RecordCount > 0 Then
                '    Dim ItemGroupQuery As String = "SELECT ""ItmsGrpNam"" FROM OITM A INNER JOIN OITB B ON B.""ItmsGrpCod""=A.""ItmsGrpCod"" WHERE A.""ItemCode""='" & item & "' "
                '    Dim oRsItemGroup As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                '    oRsItemGroup.DoQuery(ItemGroupQuery)
                '    If oRsItemGroup.RecordCount > 0 Then
                '        str1 = businessObject.Fields.Item("NextBarCode").Value
                '        Dim NextNo As Integer = CInt(str1) + 1
                '        specific.Columns.Item("1470000003").Cells.Item(specific.VisualRowCount).Specific.value = NextNo.ToString().Trim()
                '        If (objMain.objApplication.Forms.ActiveForm.Mode = BoFormMode.fm_UPDATE_MODE) Then
                '            objMain.objApplication.Forms.ActiveForm.Items.Item("1470000001").Click(BoCellClickType.ct_Regular)
                '        End If

                '        Dim NextBarcode As String = "UPDATE ""@NT_BRST"" T0 SET T0.""U_Start""=T0.""U_Start""+1 WHERE T0.""U_Itmgrpnm"" ='" & oRsItemGroup.Fields.Item(0).Value.ToString() & "' and ""U_Scale""='1' "
                '        Dim oRsNextBarcode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                '        oRsNextBarcode.DoQuery(NextBarcode)
                '    End If
                'Else
                Dim ItemGroupQuery As String = "SELECT * FROM ""@NT_BRST""  T2 WHERE T2.""U_Itmgrpnm"" IN " &
                                               "(SELECT (SELECT T1.""ItmsGrpNam"" FROM OITB T1 WHERE T1.""ItmsGrpCod""=T0.""ItmsGrpCod"") FROM OITM T0 WHERE T0.""ItemCode""='" & item & "') " &
                                               " And  T2.""U_Scale""=(SELECT ""U_AcceptWMQty"" FROM ""OITM"" Where ""ItemCode""='" & item & "') "

                Dim oRsItemGroup As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                oRsItemGroup.DoQuery(ItemGroupQuery)
                Dim NewBarcode As String = ""
                Dim ItemGroup As String = ""
                Dim SubGroup As String = ""
                If oRsItemGroup.RecordCount = 0 Then
                    ItemGroup = "ALL"
                Else
                    ItemGroup = oRsItemGroup.Fields.Item("U_Itmgrpnm").Value
                End If
                Dim str3 As String = "CALL ""SBO_ProcessBarCode""('" & item & "')"
                businessObject.DoQuery(str3)
                If businessObject.RecordCount > 0 Then
                    Dim str4 As String = businessObject.Fields.Item("NextBarCode").Value
                    specific.Columns.Item("1470000003").Cells.Item(specific.VisualRowCount).Specific.value = str4.ToString().Trim()
                    If (objMain.objApplication.Forms.ActiveForm.Mode = BoFormMode.fm_UPDATE_MODE) Then
                        objMain.objApplication.Forms.ActiveForm.Items.Item("1470000001").Click(BoCellClickType.ct_Regular)
                    End If

                    Dim NextBarcode As String = "UPDATE ""@NT_BRST"" T0 SET T0.""U_Start""=T0.""U_Start""+1 WHERE T0.""U_Itmgrpnm"" ='" & ItemGroup & "' and ""U_Scale""='1' "
                    Dim oRsNextBarcode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oRsNextBarcode.DoQuery(NextBarcode)
                End If
            End If

            'End If

        Catch exception1 As System.Exception
            objMain.objApplication.StatusBar.SetText("Error:" & exception1.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub StandardBarcodeGenerate()
        Try
            Try
                objform = objMain.objApplication.Forms.ActiveForm

                StrQry = "SELECT * FROM ""@NT_BRST"" order by ""Code"" "
                oRec = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                oRec.DoQuery(StrQry)
                If oRec.RecordCount > 0 Then
                    Dim strlen, RevStr As String
                    Dim AryCalc() As String
                    Dim calcvalue As Double = 0
                    strlen = oRec.Fields.Item("U_Itprefix").Value & oRec.Fields.Item("U_Start1").Value.ToString()
                    RevStr = StrReverse(strlen)
                    strlen = strlen.Length

                    Dim JumpNumber As Integer = 3
                    For i As Integer = 0 To RevStr.Length - 1
                        If JumpNumber = 3 Then
                            calcvalue = calcvalue + (Convert.ToDouble(RevStr.Substring(i, 1)) * JumpNumber)
                            JumpNumber = 1
                        Else
                            calcvalue = calcvalue + (Convert.ToDouble(RevStr.Substring(i, 1)) * JumpNumber)
                            JumpNumber = 3
                        End If
                    Next
                    calcvalue = calcvalue
                    AryCalc = Convert.ToString(calcvalue).Split(".")
                    Try
                        calcvalue = calcvalue / Convert.ToDouble(strlen)

                        AryCalc = Convert.ToString(calcvalue).Split(".")
                        If Convert.ToDouble(AryCalc(1)) > 0 Then
                            calcvalue = Convert.ToDouble(strlen) - Convert.ToDouble(AryCalc(1).Substring(0, 1))

                        End If
                    Catch ex As Exception
                        calcvalue = 0
                    End Try

                    'If Convert.ToInt32(strlen) = 1 Then
                    '    oAutoBarcode = oRec.Fields.Item("U_Itprefix").Value & oRec.Fields.Item("U_Start1").Value & "0" & oRec.Fields.Item("U_End").Value
                    'Else

                    'End If
                    If Convert.ToDouble(calcvalue) > 0 Then
                        oAutoBarcode = oRec.Fields.Item("U_Itprefix").Value & oRec.Fields.Item("U_Start1").Value & calcvalue.ToString()
                    Else
                        oAutoBarcode = oRec.Fields.Item("U_Itprefix").Value & oRec.Fields.Item("U_Start1").Value & calcvalue.ToString()
                        objMain.objApplication.StatusBar.SetText("13 th Digit is zero, Kindly Generate Again..." & oAutoBarcode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oAutoBarcode = "0"
                    End If

                    ''objform.Items.Item("1470002287").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                    objBarcodeMatrix = objform.Items.Item("1470000013").Specific

                    strlen = oAutoBarcode
                    strlen = strlen.Length
                    If strlen = 13 Then
                        objBarcodeMatrix.Columns.Item("1470000003").Cells.Item(objBarcodeMatrix.RowCount).Specific.value = oAutoBarcode
                    Else
                        objMain.objApplication.StatusBar.SetText("Length of the Barcode is morethan 13 digit on ..." & oAutoBarcode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                    Startt = (Convert.ToUInt64(oRec.Fields.Item("U_Start1").Value) + 13)
                    'Startt = (Convert.ToUInt64(oRec.Fields.Item("U_Start1").Value) + 6)


                    'Endd = Convert.ToUInt64(oRec.Fields.Item("U_End").Value) + 2
                    StrQry = "UPDATE ""@NT_BRST"" SET ""U_Start1"" ='" & Startt & "',""U_End"" ='" & Endd & "' Where ""Code""='" & oRec.Fields.Item("Code").Value & "' "
                    oRec = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oRec.DoQuery(StrQry)
                Else
                    objMain.objApplication.StatusBar.SetText("Barcode Setup details missing for this ItemCode or ItemGroup...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
                Try
                    StrQry = "SELECT *  FROM OBCD T0 WHERE T0.""BcdCode"" ='" & oAutoBarcode & "'"
                    oRec = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oRec.DoQuery(StrQry)
                    If oRec.RecordCount > 0 Then
                        objMain.objApplication.StatusBar.SetText("This Barcode already in Use, Kindly Generate Again..." & oAutoBarcode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        objBarcodeMatrix.DeleteRow(objBarcodeMatrix.RowCount - 1)
                    End If
                Catch ex As Exception

                End Try

            Catch ex As Exception

            End Try
        Catch ex As Exception

        End Try
    End Sub
    Public Function GetFieldValue(ByVal query As String) As Object
        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs.DoQuery(query)
        Dim str As String = oRs.Fields.Item(0).Value.ToString()
        Return str
    End Function

End Class

