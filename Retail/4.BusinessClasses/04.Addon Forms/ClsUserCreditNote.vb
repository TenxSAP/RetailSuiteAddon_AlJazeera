Imports System.Net
Imports SAPbobsCOM
Imports SAPbouiCOM
Imports System.Text.Json
Imports System.Text.Json.Serialization
Imports Newtonsoft.Json.Linq
Imports System.Configuration
Imports System.IO
Imports System.Text
Imports Newtonsoft.Json
Public Class ClsUserCreditNote

    Public objForm As SAPbouiCOM.Form
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix


    Public Property pVal As Object

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("UserCreditNote.xml", "TNXCRDNT_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("TNXCRDNT_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNXCRDNT")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNXCRDNT_C0")
            objMatrix = objForm.Items.Item("mtx").Specific

            objMatrix.SelectionMode = BoMatrixSelect.ms_Single

            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "OTNXCRDNT"))
            oDBs_Head.SetValue("U_DOCDT", 0, DateTime.Now.ToString("yyyyMMdd"))

            objMain.objUtilities.MatrixComboBoxValues(objMatrix.Columns.Item("TAX"), "SELECT T0.""Code"", T0.""Name""  FROM OVTG T0 WHERE T0.""Inactive""='N' and T0.""Category""='I'")

            Me.SetNewLine(objForm.UniqueID)

            objMain.objUtilities.ComboBoxLoadValues(objForm.Items.Item("Item_17").Specific, "SELECT T0.""BPLId"", T0.""BPLName"" FROM OBPL T0 INNER JOIN USR6 T1 ON T0.""BPLId""=T1.""BPLId"" Where T1.""UserCode""='" & objMain.objCompany.UserName & "'")

            objForm.Items.Item("Item_17").DisplayDesc = True

            Dim oCOm As SAPbouiCOM.ComboBox = objForm.Items.Item("Item_17").Specific
            oCOm.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            objMatrix.Columns.Item("Col_2").Editable = False
            objMatrix.Columns.Item("INAME").Editable = False

            objForm.Items.Item("Item_12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("Item_13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("CNAME").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("CNAME").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)




            objForm.Items.Item("Item_3").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_3").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Item_3").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)



            objForm.Items.Item("CCODE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("CCODE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("CCODE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("DocDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("Item_5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Item_5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("Item_8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("Item_19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_19").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)



            objForm.Items.Item("Item_17").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_17").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Item_17").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("Item_23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Item_23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)



            objForm.Freeze(False)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "TNXCRDNT" And pVal.BeforeAction = False Then

                Me.CreateForm()

            ElseIf pVal.MenuUID = "Delete Row" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("TNXCRDNT_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix = objForm.Items.Item("mtx").Specific
                For i As Integer = 1 To objMatrix.VisualRowCount - 1
                    If objMatrix.IsRowSelected(i) = True Then
                        objMatrix.DeleteRow(i)
                    End If
                Next
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.Columns.Item("#").Cells.Item(i).Specific.string = i
                Next
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                End If

            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm
                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNXCRDNT")
                oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNXCRDNT_C0")
                objMatrix = objForm.Items.Item("mtx").Specific

                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "OTNXCRDNT"))
                oDBs_Head.SetValue("U_DOCDT", 0, DateTime.Now.ToString("yyyyMMdd"))

                objMain.objUtilities.MatrixComboBoxValues(objMatrix.Columns.Item("TAX"), "SELECT T0.""Code"", T0.""Name""  FROM OVTG T0 WHERE T0.""Inactive""='N' and T0.""Category""='I'")

                objMatrix.Clear()
                oDBs_Details.Clear()
                objMatrix.FlushToDataSource()

                Me.SetNewLine(objForm.UniqueID)

            End If


            If pVal.MenuUID = "519" And pVal.BeforeAction = True Then 'Preview
                Try
                    Dim frs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'Credit Note' AND ""Type"" = 'C'")
                    If frs.RecordCount = 0 Then
                        objMain.objApplication.MessageBox("Report layout 'Credit Note' not found. Please upload with same name and Code as parameter", 0, "OK")
                    Else
                        objMain.objApplication.ActivateMenuItem(frs.Fields.Item(0).Value.ToString())

                        Dim CrForm As SAPbouiCOM.Form
                        Dim oedt As SAPbouiCOM.EditText
                        CrForm = objMain.objApplication.Forms.ActiveForm
                        oedt = CrForm.Items.Item("1000003").Specific
                        oedt.Value = oDBs_Head.GetValue("U_CRDNO", 0)
                        CrForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                Catch ex As Exception
                End Try

            End If

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType


                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNXCRDNT")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNXCRDNT_C0")
                    objMatrix = objForm.Items.Item("mtx").Specific

                    If pVal.ItemChanged = True And pVal.ColUID = "TAX" Then
                        Try
                            Dim Total As Double = 0
                            Dim VatTotal As Double = 0

                            For i As Integer = 1 To objMatrix.VisualRowCount

                                If CStr(objMatrix.Columns.Item("ICODE").Cells.Item(i).Specific.Value).Trim <> "" Then

                                    Dim Qty As Double = 0
                                    Dim UP As Double = 0
                                    Dim Vatper As Double = 0
                                    Dim Ltotal As Double = 0

                                    Dim GetPer As String = "Select IFNULL(""Rate"", 0) FROM OVTG Where ""Code""='" & objMatrix.Columns.Item("TAX").Cells.Item(i).Specific.Value & "'"
                                    Dim oRsGetPer As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRsGetPer.DoQuery(GetPer)

                                    If oRsGetPer.RecordCount > 0 Then
                                        Vatper = oRsGetPer.Fields.Item(0).Value
                                    Else
                                        Vatper = 0
                                    End If

                                    If CStr(objMatrix.Columns.Item("QUANTITY").Cells.Item(i).Specific.Value).Trim <> "" Then
                                        Qty = CDbl(CStr(objMatrix.Columns.Item("QUANTITY").Cells.Item(i).Specific.Value).Trim)
                                    End If

                                    If CStr(objMatrix.Columns.Item("UPRICE").Cells.Item(i).Specific.Value).Trim <> "" Then
                                        UP = CDbl(CStr(objMatrix.Columns.Item("UPRICE").Cells.Item(i).Specific.Value).Trim)
                                    End If

                                    Ltotal = (Qty * UP) + ((Qty * UP) * Vatper * 0.01)

                                    If i = pVal.Row Then

                                        oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row)
                                        oDBs_Details.SetValue("U_ICODE", oDBs_Details.Offset, objMatrix.Columns.Item("ICODE").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Details.SetValue("U_INAME", oDBs_Details.Offset, objMatrix.Columns.Item("INAME").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Details.SetValue("U_BRCD", oDBs_Details.Offset, objMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Details.SetValue("U_UOM", oDBs_Details.Offset, objMatrix.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Details.SetValue("U_QUANTITY", oDBs_Details.Offset, objMatrix.Columns.Item("QUANTITY").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Details.SetValue("U_UPRICE", oDBs_Details.Offset, objMatrix.Columns.Item("UPRICE").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Details.SetValue("U_TAX", oDBs_Details.Offset, objMatrix.Columns.Item("TAX").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Details.SetValue("U_LTOTAL", oDBs_Details.Offset, Ltotal)
                                        oDBs_Details.SetValue("U_RMKS", oDBs_Details.Offset, objMatrix.Columns.Item("Col_3").Cells.Item(pVal.Row).Specific.Value)

                                        objMatrix.SetLineData(pVal.Row)

                                    End If

                                    Total = Total + Ltotal

                                    VatTotal = VatTotal + ((Qty * UP) * Vatper * 0.01)



                                End If
                            Next

                            oDBs_Head.SetValue("U_TOTVT", 0, VatTotal)
                            oDBs_Head.SetValue("U_DOCTTL", 0, Total)
                        Catch ex As Exception

                        End Try
                    End If



                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNXCRDNT")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNXCRDNT_C0")
                    objMatrix = objForm.Items.Item("mtx").Specific

                    'If pVal.ItemUID = "mtx" And pVal.ColUID = "Col_00000" And pVal.BeforeAction = False Then
                    '    Try

                    '        Dim GetData As String = "SELECT T1.""ItemCode"", T1.""ItemName"" FROM OBCD T0  INNER JOIN OITM T1 " &
                    '                                "ON T0.""ItemCode"" = T1.""ItemCode"" WHERE T0.""BcdCode"" ='" & CStr(objMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific.Value).Trim & "'"

                    '        Dim oRsGetData As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '        oRsGetData.DoQuery(GetData)

                    '        If oRsGetData.RecordCount > 0 Then

                    '            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row)
                    '            oDBs_Details.SetValue("U_ICODE", oDBs_Details.Offset, oRsGetData.Fields.Item(0).Value)
                    '            oDBs_Details.SetValue("U_INAME", oDBs_Details.Offset, oRsGetData.Fields.Item(1).Value)
                    '            oDBs_Details.SetValue("U_BRCD", oDBs_Details.Offset, CStr(objMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific.Value).Trim)
                    '            oDBs_Details.SetValue("U_UOM", oDBs_Details.Offset, objMatrix.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific.Value)
                    '            oDBs_Details.SetValue("U_QUANTITY", oDBs_Details.Offset, 1)
                    '            oDBs_Details.SetValue("U_UPRICE", oDBs_Details.Offset, "")
                    '            oDBs_Details.SetValue("U_TAX", oDBs_Details.Offset, "")
                    '            oDBs_Details.SetValue("U_LTOTAL", oDBs_Details.Offset, "")
                    '            oDBs_Details.SetValue("U_RMKS", oDBs_Details.Offset, "")

                    '            objMatrix.SetLineData(pVal.Row)

                    '            If pVal.Row = objMatrix.VisualRowCount Then
                    '                Me.SetNewLine(objForm.UniqueID)
                    '            End If

                    '        End If
                    '    Catch ex As Exception

                    '    End Try
                    'End If

                    If pVal.ItemChanged = True And pVal.BeforeAction = False And (pVal.ColUID = "QUANTITY" Or pVal.ColUID = "UPRICE") Then

                        Try


                            Dim Total As Double = 0
                            Dim VatTotal As Double = 0

                            For i As Integer = 1 To objMatrix.VisualRowCount

                                If CStr(objMatrix.Columns.Item("ICODE").Cells.Item(i).Specific.Value).Trim <> "" Then

                                    Dim Qty As Double = 0
                                    Dim UP As Double = 0
                                    Dim Vatper As Double = 0
                                    Dim Ltotal As Double = 0

                                    Dim GetPer As String = "Select IFNULL(""Rate"", 0) FROM OVTG Where ""Code""='" & objMatrix.Columns.Item("TAX").Cells.Item(i).Specific.Value & "'"
                                    Dim oRsGetPer As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRsGetPer.DoQuery(GetPer)

                                    If oRsGetPer.RecordCount > 0 Then
                                        Vatper = oRsGetPer.Fields.Item(0).Value
                                    Else
                                        Vatper = 0
                                    End If

                                    If CStr(objMatrix.Columns.Item("QUANTITY").Cells.Item(i).Specific.Value).Trim <> "" Then
                                        Qty = CDbl(CStr(objMatrix.Columns.Item("QUANTITY").Cells.Item(i).Specific.Value).Trim)
                                    End If

                                    If CStr(objMatrix.Columns.Item("UPRICE").Cells.Item(i).Specific.Value).Trim <> "" Then
                                        UP = CDbl(CStr(objMatrix.Columns.Item("UPRICE").Cells.Item(i).Specific.Value).Trim)
                                    End If

                                    Ltotal = (Qty * UP) + ((Qty * UP) * Vatper * 0.01)

                                    If i = pVal.Row Then

                                        oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row)
                                        oDBs_Details.SetValue("U_ICODE", oDBs_Details.Offset, objMatrix.Columns.Item("ICODE").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Details.SetValue("U_INAME", oDBs_Details.Offset, objMatrix.Columns.Item("INAME").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Details.SetValue("U_BRCD", oDBs_Details.Offset, objMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Details.SetValue("U_UOM", oDBs_Details.Offset, objMatrix.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Details.SetValue("U_QUANTITY", oDBs_Details.Offset, objMatrix.Columns.Item("QUANTITY").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Details.SetValue("U_UPRICE", oDBs_Details.Offset, objMatrix.Columns.Item("UPRICE").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Details.SetValue("U_TAX", oDBs_Details.Offset, objMatrix.Columns.Item("TAX").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Details.SetValue("U_LTOTAL", oDBs_Details.Offset, Ltotal)
                                        oDBs_Details.SetValue("U_RMKS", oDBs_Details.Offset, objMatrix.Columns.Item("Col_3").Cells.Item(pVal.Row).Specific.Value)

                                        objMatrix.SetLineData(pVal.Row)

                                    End If

                                    Total = Total + Ltotal

                                    VatTotal = VatTotal + ((Qty * UP) * Vatper * 0.01)



                                End If
                            Next

                            oDBs_Head.SetValue("U_TOTVT", 0, VatTotal)
                            oDBs_Head.SetValue("U_DOCTTL", 0, Total)


                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try

                    End If


                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNXCRDNT")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNXCRDNT_C0")
                    objMatrix = objForm.Items.Item("mtx").Specific




                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Try

                            If CStr(objForm.Items.Item("CCODE").Specific.Value).Trim = "" Then
                                objMain.objApplication.StatusBar.SetText("Please select Supplier code")
                                BubbleEvent = False
                                Exit Try
                            ElseIf CStr(objForm.Items.Item("Item_17").Specific.Value).Trim = "" Then
                                objMain.objApplication.StatusBar.SetText("Please select Branch")
                                BubbleEvent = False
                                Exit Try

                            ElseIf CStr(objForm.Items.Item("DocDate").Specific.Value).Trim = "" Then
                                objMain.objApplication.StatusBar.SetText("Please select DocDate")
                                BubbleEvent = False
                                Exit Try
                            ElseIf CStr(objForm.Items.Item("Item_5").Specific.Value).Trim = "" Then
                                objMain.objApplication.StatusBar.SetText("Please select DocDue Date")
                                BubbleEvent = False
                                Exit Try
                            ElseIf objMatrix.VisualRowCount = 1 Then
                                objMain.objApplication.StatusBar.SetText("Please select atleast one item")
                                BubbleEvent = False
                                Exit Try
                            End If

                            For i As Integer = 1 To objMatrix.VisualRowCount

                                If CStr(objMatrix.Columns.Item("ICODE").Cells.Item(i).Specific.Value).Trim <> "" Then


                                    If CStr(objMatrix.Columns.Item("Col_1").Cells.Item(i).Specific.Value).Trim = "" Then

                                        objMain.objApplication.StatusBar.SetText("Please select UOM Code in Line " & i & "")
                                        BubbleEvent = False
                                        Exit Try

                                    ElseIf CStr(objMatrix.Columns.Item("TAX").Cells.Item(i).Specific.Value).Trim = "" Then

                                        objMain.objApplication.StatusBar.SetText("Please select VatCode in Line " & i & "")
                                        BubbleEvent = False
                                        Exit Try

                                    ElseIf CStr(objMatrix.Columns.Item("QUANTITY").Cells.Item(i).Specific.Value).Trim = "" Then

                                        objMain.objApplication.StatusBar.SetText("Quanity should not be null in Line " & i & "")
                                        BubbleEvent = False
                                        Exit Try

                                    ElseIf CDbl(CStr(objMatrix.Columns.Item("QUANTITY").Cells.Item(i).Specific.Value).Trim) <= 0 Then

                                        objMain.objApplication.StatusBar.SetText("Quantity should be greater than 0 in Line " & i & "")
                                        BubbleEvent = False
                                        Exit Try

                                    ElseIf CStr(objMatrix.Columns.Item("UPRICE").Cells.Item(i).Specific.Value).Trim = "" Then

                                        objMain.objApplication.StatusBar.SetText("Price should not be null in Line " & i & "")
                                        BubbleEvent = False
                                        Exit Try

                                    ElseIf CDbl(CStr(objMatrix.Columns.Item("UPRICE").Cells.Item(i).Specific.Value).Trim) <= 0 Then

                                        objMain.objApplication.StatusBar.SetText("Price should be greater than 0 in Line " & i & "")
                                        BubbleEvent = False
                                        Exit Try

                                    End If



                                End If
                            Next


                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Try
                            objForm.Freeze(True)

                            Dim chkIfExist As String = "Select ""DocNum"" From ""@TNXCRDNT"" Where ""DocEntry""=(Select MAX(""DocEntry"") From ""@TNXCRDNT"" Where ""Creator""='" & objMain.objCompany.UserName & "')"
                            Dim oRschkIfExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRschkIfExist.DoQuery(chkIfExist)
                            If oRschkIfExist.RecordCount > 0 Then
                                objForm.Freeze(True)
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                objForm.Items.Item("DocNum").Specific.value = oRschkIfExist.Fields.Item("DocNum").Value
                                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                objForm.Freeze(False)
                            End If

                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If


                    If pVal.ItemUID = "Item_15" And pVal.BeforeAction = False Then
                        Try
                            objForm.Freeze(True)
                            Me.CreateCreditNote(objForm.UniqueID)

                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If


                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNXCRDNT")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNXCRDNT_C0")
                    objMatrix = objForm.Items.Item("mtx").Specific
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal        'CHOOSE FORM LIST'
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects

                    If oCFL.UniqueID = "CFL_1" And pVal.BeforeAction = True Then
                        Me.CFLVendorsFilter(objForm.UniqueID, oCFL.UniqueID)
                    End If

                    If oCFL.UniqueID = "CFL_2" And pVal.BeforeAction = True Then
                        Me.CFLPurchaseItems(objForm.UniqueID, oCFL.UniqueID)
                    End If

                    If pVal.BeforeAction = True And oCFL.UniqueID = "CFL_UOM" Then
                        If CStr(objMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific.Value).Trim = "" Or CStr(objMatrix.Columns.Item("ICODE").Cells.Item(pVal.Row).Specific.Value).Trim = "" Then
                            objMain.objApplication.StatusBar.SetText("Item code cannot be left blank")
                            BubbleEvent = False
                            Exit Sub
                        End If
                        Me.FilterUOM(objForm.UniqueID, oCFL.UniqueID, CStr(objMatrix.Columns.Item("ICODE").Cells.Item(pVal.Row).Specific.Value).Trim)
                    End If


                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)





                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)


                        If oCFL.UniqueID = "CFL_UOM" Then
                            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row)
                            oDBs_Details.SetValue("U_ICODE", oDBs_Details.Offset, objMatrix.Columns.Item("ICODE").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_INAME", oDBs_Details.Offset, objMatrix.Columns.Item("INAME").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_BRCD", oDBs_Details.Offset, objMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_UOM", oDBs_Details.Offset, oDT.GetValue("UomCode", 0))
                            oDBs_Details.SetValue("U_QUANTITY", oDBs_Details.Offset, objMatrix.Columns.Item("QUANTITY").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_UPRICE", oDBs_Details.Offset, objMatrix.Columns.Item("UPRICE").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_TAX", oDBs_Details.Offset, objMatrix.Columns.Item("TAX").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_LTOTAL", oDBs_Details.Offset, objMatrix.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_RMKS", oDBs_Details.Offset, objMatrix.Columns.Item("Col_3").Cells.Item(pVal.Row).Specific.Value)

                            objMatrix.SetLineData(pVal.Row)

                            If pVal.Row = objMatrix.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID)
                            End If


                        End If

                        If oCFL.UniqueID = "CFL_1" Then
                            oDBs_Head.SetValue("U_CCODE", oDBs_Head.Offset, oDT.GetValue("CardCode", 0))
                            oDBs_Head.SetValue("U_CNAME", oDBs_Head.Offset, oDT.GetValue("CardName", 0))
                        End If

                        If oCFL.UniqueID = "CFL_2" Then
                            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row)
                            oDBs_Details.SetValue("U_ICODE", oDBs_Details.Offset, oDT.GetValue("ItemCode", 0))
                            oDBs_Details.SetValue("U_INAME", oDBs_Details.Offset, oDT.GetValue("ItemName", 0))
                            oDBs_Details.SetValue("U_BRCD", oDBs_Details.Offset, oDT.GetValue("CodeBars", 0))

                            Dim GetUOM As String = "SELECT T0.""UomCode"" FROM OUOM T0 WHERE T0.""UomEntry"" =(SELECT T0.""PUoMEntry"" FROM OITM T0 WHERE T0.""ItemCode""='" & oDT.GetValue("ItemCode", 0) & "')"
                            Dim oRsGetUOM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetUOM.DoQuery(GetUOM)

                            If oRsGetUOM.RecordCount > 0 Then
                                oDBs_Details.SetValue("U_UOM", oDBs_Details.Offset, oRsGetUOM.Fields.Item(0).Value)
                            Else
                                oDBs_Details.SetValue("U_UOM", oDBs_Details.Offset, "")
                            End If

                            oDBs_Details.SetValue("U_QUANTITY", oDBs_Details.Offset, 1)
                            oDBs_Details.SetValue("U_UPRICE", oDBs_Details.Offset, 0)
                            oDBs_Details.SetValue("U_TAX", oDBs_Details.Offset, "")
                            oDBs_Details.SetValue("U_LTOTAL", oDBs_Details.Offset, "")
                            oDBs_Details.SetValue("U_RMKS", oDBs_Details.Offset, "")


                            objMatrix.SetLineData(pVal.Row)

                            If pVal.Row = objMatrix.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID)
                            End If
                        End If

                        If oCFL.UniqueID = "CFL_BC" Then
                            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row)
                            oDBs_Details.SetValue("U_ICODE", oDBs_Details.Offset, oDT.GetValue("ItemCode", 0))
                            oDBs_Details.SetValue("U_BRCD", oDBs_Details.Offset, oDT.GetValue("BcdCode", 0))

                            'Dim GetUOM As String = "SELECT T0.""UomCode"",T1.""ItemName"" FROM OUOM T0 INNER JOIN OITM T1 ON T0.""UomEntry""=T1.""PUoMEntry"" " &
                            '"WHERE T1.""ItemCode"" ='" & oDT.GetValue("ItemCode", 0) & "'"

                            'Dim GetUOM As String = "SELECT T0.""UomEntry"", T1.""UomCode"" FROM OBCD T0  INNER JOIN OUOM T1 ON " &
                            '                     "T0.""UomEntry"" = T1.""UomEntry"" WHERE T0.""BcdCode"" ='" & CStr(oDT.GetValue("BcdCode", 0)).Trim & "'"

                            Dim GetUOM As String = "SELECT T1.""UomCode"", T2.""ItemName"" FROM OBCD T0  INNER JOIN OUOM T1 " &
                                                         "ON T0.""UomEntry"" = T1.""UomEntry"" INNER JOIN OITM T2 ON " &
                                                        "T0.""ItemCode"" = T2.""ItemCode"" WHERE T0.""BcdCode"" ='" & oDT.GetValue("BcdCode", 0) & "'"

                            Dim oRsGetUOM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetUOM.DoQuery(GetUOM)

                            If oRsGetUOM.RecordCount > 0 Then
                                oDBs_Details.SetValue("U_UOM", oDBs_Details.Offset, oRsGetUOM.Fields.Item(0).Value)
                                oDBs_Details.SetValue("U_INAME", oDBs_Details.Offset, oRsGetUOM.Fields.Item(1).Value)

                            Else
                                oDBs_Details.SetValue("U_UOM", oDBs_Details.Offset, "")
                                oDBs_Details.SetValue("U_INAME", oDBs_Details.Offset, "")

                            End If

                            oDBs_Details.SetValue("U_QUANTITY", oDBs_Details.Offset, 1)
                            oDBs_Details.SetValue("U_UPRICE", oDBs_Details.Offset, 0)
                            oDBs_Details.SetValue("U_TAX", oDBs_Details.Offset, "")
                            oDBs_Details.SetValue("U_LTOTAL", oDBs_Details.Offset, "")
                            oDBs_Details.SetValue("U_RMKS", oDBs_Details.Offset, "")


                            objMatrix.SetLineData(pVal.Row)

                            If pVal.Row = objMatrix.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID)
                            End If
                        End If
                    End If

            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNXCRDNT")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNXCRDNT_C0")
            objMatrix = objForm.Items.Item("mtx").Specific


            objMatrix.AddRow()

            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
            oDBs_Details.SetValue("U_BRCD", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_ICODE", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_INAME", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_UOM", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_QUANTITY", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_UPRICE", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_TAX", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_LTOTAL", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_RMKS", oDBs_Details.Offset, "")

            objMatrix.SetLineData(objMatrix.VisualRowCount)

            objMatrix.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub CreateCreditNote(ByVal FormUID As String)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNXCRDNT")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNXCRDNT_C0")
            objMatrix = objForm.Items.Item("mtx").Specific

            objForm.Freeze(True)


            Dim oCompany As SAPbobsCOM.Company = New SAPbobsCOM.Company()


            oCompany.Server = objMain.objCompany.Server
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
            oCompany.CompanyDB = objMain.objCompany.CompanyDB
            'oCompany.UserName = "manager"
            'oCompany.Password = "Nes@1234"
            oCompany.UserName = "B1i"
            oCompany.Password = "1234"
            'oCompany.DbUserName = "SYSTEM"
            'oCompany.DbPassword = "NewEr@987"

            'If oCompany.Connect() <> 0 Then
            '    objMain.objApplication.StatusBar.SetText("Error connecting to SAP Business One: " & oCompany.GetLastErrorDescription())
            '    objForm.Freeze(False)
            '    Exit Try
            'End If

            Dim GeteData As String = "Select *,(select ""UomEntry"" from ouom where ""UomCode""=""U_UOM"") as ""UOMENTRY"" FROM ""@TNXCRDNT"" A INNER JOIN ""@TNXCRDNT_C0"" B  " &
                                  "ON A.""DocEntry""=B.""DocEntry"" Where A.""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0) & "' " &
                                  "And (A.""U_CRDNO"" Is NULL Or A.""U_CRDNO""='') And B.""U_ICODE"" IS NOT NULL And B.""U_ICODE""<>'' " &
                                  "And B.""U_BRCD"" Is Not NULL And B.""U_BRCD""<>''"

            Dim oRsgetData As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oRsgetData.DoQuery(GeteData)

            If oRsgetData.RecordCount = 0 Then
                objMain.objApplication.StatusBar.SetText("No Data avialble to POST", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objForm.Freeze(False)
                Exit Try
            End If


            Dim CheckDoc As String = "Select * FROM ""ORPC"" WHERE ""U_BASEEN""='" & oDBs_Head.GetValue("DocEntry", 0) & "' And ""U_BASEEN"" IS NOT NULL AND ""U_BASEEN""<>''"
            Dim oRsCheckDoc As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsCheckDoc.DoQuery(CheckDoc)

            If oRsCheckDoc.RecordCount > 0 Then
                objMain.objApplication.StatusBar.SetText("Document already created")
                objForm.Freeze(False)
                Exit Try
            End If

            ' Step 3: Call Service Layer function to create AP Credit Note
            Dim result = PostAPCreditNoteToServiceLayer(oDBs_Head, oRsgetData)
            If result.Success Then
                Dim newDocEntry As String = result.DocEntry
                Dim newDocNum As String = result.DocNum

                ' Show success message or proceed with further logic
                objMain.objApplication.StatusBar.SetText("AP Credit Note created successfully. DocNum: " & newDocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Dim Updatedata As String = "Update ""@TNXCRDNT"" Set ""U_CRDNO""='" & newDocEntry & "',""U_CDOCN""= " &
                "(Select ""DocNum"" FROM ORPC Where ""DocEntry""='" & newDocEntry & "')   Where ""DocEntry""='" & oDBs_Head.GetValue("DocEntry", 0) & "'"
                Dim oRsGetUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetUpdate.DoQuery(Updatedata)


                objMain.objApplication.StatusBar.SetText("AP Credit Note posted successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


            Else
                ' Handle or display the error message
                'objMain.objApplication.StatusBar.SetText("Error adding AP Credit Note: " & oCompany.GetLastErrorDescription())

                objMain.objApplication.StatusBar.SetText("Failed to create AP Credit Note. Error: " & result.ErrorMessage)
            End If

            'Dim oAPCreditNote As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)



            'oAPCreditNote.CardCode = oRsgetData.Fields.Item("U_CCODE").Value
            'oAPCreditNote.DocDate = oRsgetData.Fields.Item("U_DOCDT").Value
            'oAPCreditNote.DocDueDate = oRsgetData.Fields.Item("U_DOCDDT").Value
            'oAPCreditNote.TaxDate = oRsgetData.Fields.Item("U_DOCDT").Value

            'oAPCreditNote.Comments = oRsgetData.Fields.Item("U_RMKSS").Value  ' "Based On:" & oDBs_Head.GetValue("DocEntry", 0) & ""

            'oAPCreditNote.NumAtCard = oRsgetData.Fields.Item("U_VREFNO").Value
            'oAPCreditNote.BPL_IDAssignedToInvoice = objForm.Items.Item("Item_17").Specific.Value
            'oAPCreditNote.UserFields.Fields.Item("U_BASEEN").Value = oDBs_Head.GetValue("DocEntry", 0)
            'oAPCreditNote.UserFields.Fields.Item("U_GRPO").Value = oDBs_Head.GetValue("U_GRPON", 0)


            'oRsgetData.MoveFirst()

            'For i As Integer = 1 To oRsgetData.RecordCount

            '    If i > 1 Then
            '        oAPCreditNote.Lines.Add()
            '    End If

            '    oAPCreditNote.Lines.ItemCode = oRsgetData.Fields.Item("U_ICODE").Value
            '    oAPCreditNote.Lines.Quantity = oRsgetData.Fields.Item("U_QUANTITY").Value
            '    oAPCreditNote.Lines.UnitPrice = oRsgetData.Fields.Item("U_UPRICE").Value
            '    oAPCreditNote.Lines.VatGroup = oRsgetData.Fields.Item("U_TAX").Value

            '    Dim GetUOMEntry As String = "SELECT T0.""UomEntry"" FROM OUOM T0 WHERE T0.""UomCode""='" & oRsgetData.Fields.Item("U_UOM").Value & "'"
            '    Dim oRsGetUOMEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            '    oRsGetUOMEntry.DoQuery(GetUOMEntry)

            '    If oRsGetUOMEntry.RecordCount > 0 Then
            '        oAPCreditNote.Lines.UoMEntry = oRsGetUOMEntry.Fields.Item(0).Value
            '    End If



            '    oAPCreditNote.Lines.WithoutInventoryMovement = SAPbobsCOM.BoYesNoEnum.tYES

            '    oRsgetData.MoveNext()
            'Next


            'If oAPCreditNote.Add() <> 0 Then
            'If Not success Then
            '    objMain.objApplication.StatusBar.SetText("Error adding AP Credit Note: " & oCompany.GetLastErrorDescription())

            'Else




            'End If

            'oCompany.Disconnect()

            objMain.objApplication.ActivateMenuItem("1304")

            objForm.Freeze(False)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub CFLVendorsFilter(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            oCondition = oConditions.Add()
            oCondition.Alias = "CardType"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "S"
            oChooseFromList.SetConditions(oConditions)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub CFLPurchaseItems(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            oCondition = oConditions.Add()
            oCondition.Alias = "PrchseItem"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "Y"
            oChooseFromList.SetConditions(oConditions)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("TNXCRDNT_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then

                        If CStr(objForm.Items.Item("Item_8").Specific.Value).Trim <> "" Then

                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                        Else

                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

                        End If
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FilterUOM(ByVal FormUID As String, ByVal CFLID As String, ByVal ItemCode As String)
        Try
            Dim Query As String = "SELECT T3.""UomCode""  FROM OITM T0  INNER JOIN OUGP T1 ON T0.""UgpEntry""  = T1.""UgpEntry"" " &
                                    "INNER JOIN UGP1 T2 ON T1.""UgpEntry""  = T2.""UgpEntry""  INNER JOIN OUOM T3 ON " &
                            "T2.""UomEntry""  = T3.""UomEntry""  WHERE T0.""ItemCode""  ='" & ItemCode.Trim & "'"

            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsQuery.DoQuery(Query)

            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFLID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()

            If oRsQuery.RecordCount > 0 Then

                For i As Integer = 1 To oRsQuery.RecordCount
                    oCondition = oConditions.Add()
                    oCondition.Alias = "UomCode"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = oRsQuery.Fields.Item(0).Value.ToString().Trim
                    If i <> oRsQuery.RecordCount Then
                        oCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    End If
                    oRsQuery.MoveNext()
                Next

            Else

                oCondition = oConditions.Add()
                oCondition.Alias = "UomCode"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = ""
            End If


            oChooseFromList.SetConditions(oConditions)

        Catch ex As Exception

        End Try
    End Sub

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            Dim objForm As SAPbouiCOM.Form
            Dim oMenuItem As SAPbouiCOM.MenuItem
            Dim oMenus As SAPbouiCOM.Menus
            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
            oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNXCRDNT")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNXCRDNT_C0")
            objMatrix = objForm.Items.Item("mtx").Specific

            If eventInfo.FormUID = objForm.UniqueID Then
                If (eventInfo.BeforeAction = True) Then
                    If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE And CStr(oDBs_Head.GetValue("U_CRDNO", 0)).Trim = "" Then

                        If eventInfo.ItemUID = "mtx" And eventInfo.ColUID = "#" Then
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("Delete Row") = False Then
                                    oCreationPackage.UniqueID = "Delete Row"
                                    oCreationPackage.String = "Delete Row"
                                    oCreationPackage.Enabled = True
                                    oMenus.AddEx(oCreationPackage)
                                End If

                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        Else
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("Delete Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Delete Row")
                                End If
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try

                        End If

                    Else

                        Try
                            oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                            oMenus = oMenuItem.SubMenus
                            If oMenus.Exists("Delete Row") = True Then
                                objMain.objApplication.Menus.RemoveEx("Delete Row")
                            End If
                            If oMenus.Exists("Close") = True Then
                                objMain.objApplication.Menus.RemoveEx("Close")
                            End If
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    ' -----------------------------
    ' Company Session Login
    ' -----------------------------
    Private Function CompanySession() As String
        Try
            Dim baseUrl As String = ConfigurationManager.AppSettings("IP")
            Dim loginUrl As String = baseUrl & "/b1s/v1/Login"

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            ServicePointManager.ServerCertificateValidationCallback = Function() True

            Dim loginPayload = New With {
            .CompanyDB = objMain.objCompany.CompanyDB,
            .UserName = ConfigurationManager.AppSettings("ID"),
            .Password = ConfigurationManager.AppSettings("Password")
        }

            Dim docJson As String = JsonConvert.SerializeObject(loginPayload)
            Dim requestData As Byte() = Encoding.UTF8.GetBytes(docJson)

            Dim request = CType(WebRequest.Create(loginUrl), HttpWebRequest)
            request.ContentType = "application/json"
            request.Method = "POST"
            request.ProtocolVersion = HttpVersion.Version11
            request.KeepAlive = False
            request.ContentLength = requestData.Length
            request.Timeout = 30000
            request.ServicePoint.Expect100Continue = False

            Using stream = request.GetRequestStream()
                stream.Write(requestData, 0, requestData.Length)
            End Using

            Using response = CType(request.GetResponse(), HttpWebResponse)
                For i As Integer = 0 To response.Headers.Count - 1
                    If response.Headers.GetKey(i) = "Set-Cookie" Then
                        Return response.Headers.Get(i)
                    End If
                Next
            End Using

        Catch ex As WebException
            Dim errorDetails As String = ""
            If ex.Response IsNot Nothing Then
                Using reader As New StreamReader(ex.Response.GetResponseStream())
                    errorDetails = reader.ReadToEnd()
                End Using
            Else
                errorDetails = "No response from server: " & ex.Message
            End If
            objMain.objApplication.StatusBar.SetText("WebException: " & errorDetails)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Exception: " & ex.Message)
        End Try

        Return ""
    End Function

    ' -----------------------------
    ' Post Credit Note
    ' -----------------------------
    Private Function PostAPCreditNoteToServiceLayer(header As SAPbouiCOM.DBDataSource, linesRs As SAPbobsCOM.Recordset) As APPostingResult
        Dim cookieString As String = ""

        Try
            Dim baseUrl As String = ConfigurationManager.AppSettings("IP")
            Dim postUrl As String = baseUrl & "/b1s/v1/PurchaseCreditNotes"

            cookieString = CompanySession()
            If String.IsNullOrEmpty(cookieString) Then
                Return New APPostingResult With {
                .Success = False,
                .ErrorMessage = "Login failed, no session cookie."
            }
            End If

            Dim sessionVal As String = cookieString.Split("="c)(1).Split(";"c)(0)
            Dim container As New CookieContainer()
            container.Add(New Uri(postUrl), New Cookie("B1SESSION", sessionVal))

            Dim BPLID As Integer = objForm.Items.Item("Item_17").Specific.Value
            Dim document = New Dictionary(Of String, Object) From {
            {"CardCode", header.GetValue("U_CCODE", 0).ToString().Trim()},
            {"DocDate", header.GetValue("U_DOCDT", 0).ToString()},
            {"DocDueDate", header.GetValue("U_DOCDDT", 0).ToString()},
            {"TaxDate", header.GetValue("U_DOCDT", 0).ToString()},
            {"Comments", header.GetValue("U_RMKSS", 0).ToString().Trim()},
            {"NumAtCard", header.GetValue("U_VREFNO", 0).ToString().Trim()},
            {"BPL_IDAssignedToInvoice", BPLID},
            {"U_BASEEN", header.GetValue("DocEntry", 0).ToString().Trim()},
            {"U_GRPO", header.GetValue("U_GRPON", 0).ToString().Trim()},
            {"U_UserName", objMain.objCompany.UserName.ToString().Trim()}
        }

            ' Prepare document lines
            Dim lines As New List(Of Object)()
            linesRs.MoveFirst()
            Do Until linesRs.EoF
                lines.Add(New With {
                .ItemCode = linesRs.Fields.Item("U_ICODE").Value.ToString(),
                .Quantity = Convert.ToDecimal(linesRs.Fields.Item("U_QUANTITY").Value),
                .UnitPrice = Convert.ToDecimal(linesRs.Fields.Item("U_UPRICE").Value),
                .VatGroup = linesRs.Fields.Item("U_TAX").Value.ToString(),
                .UoMEntry = linesRs.Fields.Item("UOMENTRY").Value.ToString(),
                .WithoutInventoryMovement = "tYES",
                .U_suppprice = linesRs.Fields.Item("U_Cost").Value
            })
                linesRs.MoveNext()
            Loop
            document.Add("DocumentLines", lines)

            ' Serialize document and send request
            Dim docJson As String = JsonConvert.SerializeObject(document)
            Dim requestData As Byte() = Encoding.UTF8.GetBytes(docJson)

            Dim request = CType(WebRequest.Create(postUrl), HttpWebRequest)
            request.Method = "POST"
            request.ContentType = "application/json"
            request.ContentLength = requestData.Length
            request.CookieContainer = container
            request.ProtocolVersion = HttpVersion.Version11
            request.KeepAlive = False
            request.ServicePoint.Expect100Continue = False

            Using stream = request.GetRequestStream()
                stream.Write(requestData, 0, requestData.Length)
            End Using

            Using response = CType(request.GetResponse(), HttpWebResponse)
                Using sr As New StreamReader(response.GetResponseStream())
                    Dim result = sr.ReadToEnd()
                    Dim parsed = JObject.Parse(result)
                    LogoutFromServiceLayer(cookieString)

                    Return New APPostingResult With {
                    .Success = True,
                    .DocEntry = parsed("DocEntry").ToString(),
                    .DocNum = parsed("DocNum").ToString(),
                    .ErrorMessage = ""
                }
                End Using
            End Using

        Catch ex As WebException
            objForm.Freeze(False)
            Dim errorDetails As String = ""
            If ex.Response IsNot Nothing Then
                Using reader As New StreamReader(ex.Response.GetResponseStream())
                    errorDetails = reader.ReadToEnd()
                End Using
            Else
                errorDetails = ex.Message
            End If

            Return New APPostingResult With {
            .Success = False,
            .DocEntry = "",
            .DocNum = "",
            .ErrorMessage = "WebException: " & errorDetails
        }

        Catch ex As Exception
            objForm.Freeze(False)
            Return New APPostingResult With {
            .Success = False,
            .DocEntry = "",
            .DocNum = "",
            .ErrorMessage = "General Exception: " & ex.Message
        }
        End Try
    End Function

    ' -----------------------------
    ' Logout from Service Layer
    ' -----------------------------
    Private Sub LogoutFromServiceLayer(sessionCookie As String)
        Try
            Dim baseUrl As String = ConfigurationManager.AppSettings("IP")
            Dim logoutUrl As String = baseUrl & "/b1s/v1/Logout"

            Dim sessionVal As String = sessionCookie.Split("="c)(1).Split(";"c)(0)

            Dim request = CType(WebRequest.Create(logoutUrl), HttpWebRequest)
            request.Method = "POST"
            request.ContentType = "application/json"
            request.ProtocolVersion = HttpVersion.Version11
            request.KeepAlive = False

            Dim container As New CookieContainer()
            container.Add(New Uri(logoutUrl), New Cookie("B1SESSION", sessionVal))
            request.CookieContainer = container

            Using response = CType(request.GetResponse(), HttpWebResponse)
                If response.StatusCode = HttpStatusCode.NoContent Then
                    objMain.objApplication.StatusBar.SetText("Logged out successfully.")
                Else
                    objMain.objApplication.StatusBar.SetText("Logout response status: " & response.StatusCode.ToString())
                End If
            End Using

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Logout failed: " & ex.Message)
        End Try
    End Sub


    'Private Function CompanySession() As String
    '    Try
    '        Dim baseUrl As String = ConfigurationManager.AppSettings("IP") ' Example: https://192.168.1.7:50000
    '        Dim loginUrl As String = baseUrl & "/b1s/v1/Login"

    '        ServicePointManager.Expect100Continue = True
    '        ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11
    '        ServicePointManager.ServerCertificateValidationCallback = Function() True

    '        Dim loginPayload = New With {
    '        .CompanyDB = objMain.objCompany.CompanyDB,
    '        .UserName = ConfigurationManager.AppSettings("ID"),
    '        .Password = ConfigurationManager.AppSettings("Password")
    '    }

    '        Dim request = CType(WebRequest.Create(loginUrl), HttpWebRequest)
    '        request.ContentType = "application/json"
    '        request.Method = "POST"
    '        request.ProtocolVersion = HttpVersion.Version10
    '        request.KeepAlive = False
    '        request.Timeout = 30000
    '        ' request.ReadWriteTimeout = 30000
    '        Dim docJson As String = Newtonsoft.Json.JsonConvert.SerializeObject(loginPayload)
    '        request.ServicePoint.Expect100Continue = False ' Can help with slow servers

    '        Using sw = New StreamWriter(request.GetRequestStream())
    '            sw.Write(JsonConvert.SerializeObject(loginPayload))
    '        End Using

    '        Dim response = CType(request.GetResponse(), HttpWebResponse)

    '        For i As Integer = 0 To response.Headers.Count - 1
    '            If response.Headers.GetKey(i) = "Set-Cookie" Then
    '                Return response.Headers.Get(i)
    '            End If
    '        Next

    '    Catch ex As WebException
    '        Dim errorDetails As String = ""

    '        If ex.Response IsNot Nothing Then
    '            Using reader As New StreamReader(ex.Response.GetResponseStream())
    '                errorDetails = reader.ReadToEnd()
    '            End Using
    '        Else
    '            errorDetails = "No response from server: " & ex.Message
    '        End If
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText("Exception: " & ex.Message)
    '    End Try

    '    Return ""
    'End Function
    'Private Function PostAPCreditNoteToServiceLayer(header As SAPbouiCOM.DBDataSource, linesRs As SAPbobsCOM.Recordset) As APPostingResult
    '    Dim cookieString As String = ""
    '    Try
    '        Dim baseUrl As String = ConfigurationManager.AppSettings("IP") ' Ex: https://192.168.1.7:50000
    '        Dim postUrl As String = baseUrl & "/b1s/v1/PurchaseCreditNotes"

    '        cookieString = CompanySession()
    '        If String.IsNullOrEmpty(cookieString) Then
    '            Return New APPostingResult With {
    '            .Success = False,
    '            .ErrorMessage = "Login failed, no session cookie."
    '        }

    '        End If

    '        Dim sessionVal As String = cookieString.Split("="c)(1).Split(";"c)(0)
    '        Dim container As New CookieContainer()
    '        container.Add(New Uri(postUrl), New Cookie("B1SESSION", sessionVal))
    '        Dim BPLID As Integer = objForm.Items.Item("Item_17").Specific.Value
    '        Dim document = New Dictionary(Of String, Object) From {
    '    {"CardCode", header.GetValue("U_CCODE", 0).ToString().Trim()},
    '    {"DocDate", header.GetValue("U_DOCDT", 0).ToString()},
    '    {"DocDueDate", header.GetValue("U_DOCDDT", 0).ToString()},
    '    {"TaxDate", header.GetValue("U_DOCDT", 0).ToString()},
    '    {"Comments", header.GetValue("U_RMKSS", 0).ToString().Trim()},
    '    {"NumAtCard", header.GetValue("U_VREFNO", 0).ToString().Trim()},
    '    {"BPL_IDAssignedToInvoice", BPLID},
    '    {"U_BASEEN", header.GetValue("DocEntry", 0).ToString().Trim()},
    '    {"U_GRPO", header.GetValue("U_GRPON", 0).ToString().Trim()},
    '    {"U_UserName", objMain.objCompany.UserName.ToString().Trim()}
    '}

    '        Dim lines As New List(Of Object)()
    '        linesRs.MoveFirst()
    '        Do Until linesRs.EoF
    '            lines.Add(New With {
    '        .ItemCode = linesRs.Fields.Item("U_ICODE").Value.ToString(),
    '        .Quantity = Convert.ToDecimal(linesRs.Fields.Item("U_QUANTITY").Value),
    '        .UnitPrice = Convert.ToDecimal(linesRs.Fields.Item("U_UPRICE").Value),
    '        .VatGroup = linesRs.Fields.Item("U_TAX").Value.ToString(),
    '        .UoMEntry = linesRs.Fields.Item("UOMENTRY").Value.ToString(),
    '        .WithoutInventoryMovement = "tYES",
    '        .U_suppprice = linesRs.Fields.Item("U_Cost").Value
    '    })
    '            linesRs.MoveNext()
    '        Loop
    '        document.Add("DocumentLines", lines)

    '        Dim request = CType(WebRequest.Create(postUrl), HttpWebRequest)
    '        request.ContentType = "application/json"
    '        request.Method = "POST"
    '        request.CookieContainer = container
    '        request.ProtocolVersion = HttpVersion.Version10
    '        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
    '        request.ServicePoint.Expect100Continue = False ' Can help with slow servers

    '        Using sw = New StreamWriter(request.GetRequestStream())
    '            sw.Write(JsonConvert.SerializeObject(document))
    '        End Using

    '        Dim jsonPayload As String = JsonConvert.SerializeObject(document)


    '        Dim docJson As String = Newtonsoft.Json.JsonConvert.SerializeObject(document)


    '        Dim response = CType(request.GetResponse(), HttpWebResponse)
    '        Using sr As New StreamReader(response.GetResponseStream())
    '            Dim result = sr.ReadToEnd()
    '            Dim parsed = JObject.Parse(result)
    '            LogoutFromServiceLayer(cookieString)

    '            Return New APPostingResult With {
    '            .Success = True,
    '            .DocEntry = parsed("DocEntry").ToString(),
    '            .DocNum = parsed("DocNum").ToString(),
    '            .ErrorMessage = ""}

    '        End Using



    '    Catch ex As WebException
    '        objForm.Freeze(False)
    '        Dim errorDetails As String = ""
    '        If ex.Response IsNot Nothing Then
    '            Using reader As New StreamReader(ex.Response.GetResponseStream())
    '                errorDetails = reader.ReadToEnd()
    '            End Using
    '        Else
    '            errorDetails = ex.Message
    '        End If

    '        ' Add inner exception details if present
    '        Dim innerDetails As String = ""
    '        If ex.InnerException IsNot Nothing Then
    '            innerDetails = " | InnerException: " & ex.InnerException.Message
    '        End If

    '        Return New APPostingResult With {
    '        .Success = False,
    '        .DocEntry = "",
    '        .DocNum = "",
    '        .ErrorMessage = "WebException: " & errorDetails & innerDetails
    '    }

    '    Catch ex As Exception
    '        objForm.Freeze(False)
    '        Return New APPostingResult With {
    '        .Success = False,
    '        .DocEntry = "",
    '        .DocNum = "",
    '        .ErrorMessage = "General Exception: " & ex.Message
    '    }
    '    End Try
    'End Function

    'Private Sub LogoutFromServiceLayer(sessionCookie As String)
    '    Try
    '        Dim baseUrl As String = ConfigurationManager.AppSettings("IP")
    '        Dim logoutUrl As String = baseUrl & "/b1s/v1/Logout"

    '        ' Extract session value from cookie string
    '        Dim sessionVal As String = sessionCookie.Split("="c)(1).Split(";"c)(0)

    '        ServicePointManager.Expect100Continue = True
    '        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11
    '        ServicePointManager.ServerCertificateValidationCallback = Function() True

    '        Dim request = CType(WebRequest.Create(logoutUrl), HttpWebRequest)
    '        request.Method = "POST"
    '        request.ContentType = "application/json"
    '        request.ProtocolVersion = HttpVersion.Version10

    '        ' Set session cookie
    '        Dim container As New CookieContainer()
    '        container.Add(New Uri(logoutUrl), New Cookie("B1SESSION", sessionVal))
    '        request.CookieContainer = container

    '        ' Get response (logout doesn't return content)
    '        Dim response = CType(request.GetResponse(), HttpWebResponse)
    '        Try

    '            Using reader As New StreamReader(response.GetResponseStream())
    '                Dim responseText As String = reader.ReadToEnd()
    '                ' Now you have the full response in responseText
    '                objMain.objApplication.StatusBar.SetText("Response: " & responseText)
    '            End Using
    '        Catch ex As WebException
    '            ' Handle error response here if needed
    '            If ex.Response IsNot Nothing Then
    '                Using reader As New StreamReader(ex.Response.GetResponseStream())
    '                    Dim errorResponse As String = reader.ReadToEnd()
    '                    objMain.objApplication.StatusBar.SetText("WebException: " & errorResponse)
    '                End Using
    '            Else
    '                objMain.objApplication.StatusBar.SetText("WebException: " & ex.Message)
    '            End If
    '        End Try
    '        If response.StatusCode = HttpStatusCode.NoContent Then
    '            objMain.objApplication.StatusBar.SetText("Logged out successfully.")
    '        Else
    '            objMain.objApplication.StatusBar.SetText("Logout response status: " & response.StatusCode.ToString())
    '        End If

    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText("Logout failed: " & ex.Message)
    '    End Try
    'End Sub

    Public Class APPostingResult
        Public Property Success As Boolean
        Public Property DocEntry As String
        Public Property DocNum As String
        Public Property ErrorMessage As String
    End Class

End Class
