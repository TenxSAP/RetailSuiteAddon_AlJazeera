Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TreeView
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports SAPbobsCOM
Imports SAPbouiCOM

Public Class clsItemMaster

#Region "  Declaration   "
    Public objform, objListForm, objBarcodeForm As SAPbouiCOM.Form
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim objMatrix, objBarcodeMatrix As SAPbouiCOM.Matrix
    Dim StrQry, oAutoBarcode, Var As String
    Dim oRec As SAPbobsCOM.Recordset
    Dim Startt As Int64 = 0
    Dim Endd As Integer = 0
    Dim oItem As SAPbouiCOM.Item
#End Region
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.BeforeAction = False Then
                        Me.AddItems(objform.UniqueID)
                        'madhu comment
                        'Dim oCombo As SAPbouiCOM.ComboBox = CType(objform.Items.Item("U_REBEGBLE").Specific, SAPbouiCOM.ComboBox)
                        'Dim oQtyItem As SAPbouiCOM.Item = objform.Items.Item("U_REBQTY")
                        'oCombo.Select("No", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        'oQtyItem.Enabled = False
                    End If




                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    'If pVal.ItemUID = "Btn_GRBR" And pVal.BeforeAction = False Then
                    '    Me.Generate_Barcode(objform.UniqueID)
                    '    'Me.StandardBarcodeGenerate()
                    'End If
                    objMain.objUtilities.GlobalItemNumber = objform.Items.Item("5").Specific.Value
                    If pVal.ItemUID = "Btn_GRBR" And pVal.Before_Action = False And objform.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Try
                            Dim businessObject As Recordset = DirectCast(objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset), Recordset)
                            Dim item As String = objform.Items.Item("5").Specific.Value
                            Dim str1 As String = ""
                            Dim DigitsQry As String = "SELECT Case When ""U_AcceptWMQty""='1' Then '6' When ""U_AcceptWMQty""='2' Then '13' End FROM OITM WHERE ""ItemCode"" = '" & item & "'"
                            Dim oRsDigits As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                            oRsDigits.DoQuery(DigitsQry)

                            If oRsDigits.Fields.Item(0).Value <> "6" Then '------------13-------------------------------------

                                Dim ItemGroupQuery As String = "SELECT * FROM ""@NT_BRST""  T2 WHERE T2.""U_Itmgrpnm"" IN " &
                                "(SELECT (SELECT T1.""ItmsGrpNam"" FROM OITB T1 WHERE T1.""ItmsGrpCod""=T0.""ItmsGrpCod"") FROM OITM T0 WHERE T0.""ItemCode""='" & item & "') " &
                                " And  T2.""U_Scale""=(SELECT ""U_AcceptWMQty"" FROM ""OITM"" Where ""ItemCode""='" & item & "') and  T2.""U_SubGrp"" =(SELECT ""U_VSubGroup"" FROM ""OITM"" Where ""ItemCode""='" & item & "')"
                                Dim oRsItemGroup As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                                oRsItemGroup.DoQuery(ItemGroupQuery)
                                Dim NewBarcode As String = ""
                                Dim ItemGroup As String = ""
                                Dim SubGroup As String = ""
                                If oRsItemGroup.RecordCount = 0 Then
                                    ItemGroup = "ALL"
                                Else
                                    ItemGroup = oRsItemGroup.Fields.Item("U_Itmgrpnm").Value
                                    SubGroup = oRsItemGroup.Fields.Item("U_SubGrp").Value
                                End If
                                'Dim Barcode As String = "SELECT T0.""U_Prefix"", T0.""U_Itprefix"", T0.""U_Start"",T0.""U_Start""+1 AS ""NextNo"" FROM ""@NT_BRST""  T0 WHERE T0.""U_Itmgrpnm"" ='" & ItemGroup & "' and ""U_Scale""='2' "
                                Dim Barcode As String = "SELECT T0.""U_Prefix"", T0.""U_Itprefix"", T0.""U_Start"",T0.""U_Start""+1 AS ""NextNo"" FROM ""@NT_BRST""  T0 WHERE T0.""U_Itmgrpnm"" ='" & ItemGroup & "' and T0.""U_SubGrp"" ='" & SubGroup & "'and ""U_Scale""='2' "
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
                                        oAutoBarcode = str1.Trim()
                                        objform.Items.Item("1470002287").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        objBarcodeForm = objMain.objApplication.Forms.GetForm("1470000020", pVal.FormTypeCount)
                                        objBarcodeMatrix = objBarcodeForm.Items.Item("1470000013").Specific
                                        objBarcodeMatrix.Columns.Item("1470000003").Cells.Item(1).Specific.value = oAutoBarcode

                                        ' specific.Columns.Item("1470000003").Cells.Item(specific.VisualRowCount).Specific.value = str1.Trim()
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
                                Dim NextBarcode As String = "UPDATE ""@NT_BRST"" T0 SET T0.""U_Start""='" & NextNo & "' WHERE T0.""U_Itmgrpnm"" ='" & ItemGroup & "' and T0.""U_SubGrp"" ='" & SubGroup & "' and ""U_Scale""='2' "
                                Dim oRsNextBarcode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                                oRsNextBarcode.DoQuery(NextBarcode)
                            Else
                                Dim str As String = "CALL ""SBO_ProcessBarCode""('" & item & "')"
                                businessObject.DoQuery(str)
                                If businessObject.RecordCount > 0 Then
                                    '    str1 = BusinessObject.Fields.Item("NextBarCode").Value
                                    '    ' specific.Columns.Item("1470000003").Cells.Item(specific.VisualRowCount).Specific.value = str1.Trim()
                                    '    If (objMain.objApplication.Forms.ActiveForm.Mode = BoFormMode.fm_UPDATE_MODE) Then
                                    '        objMain.objApplication.Forms.ActiveForm.Items.Item("1470000001").Click(BoCellClickType.ct_Regular)
                                    '    End If
                                    'Else
                                    '    objMain.objApplication.StatusBar.SetText("No records are found. Please check the BarCode Setup", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                    'End If
                                    '            StrQry = "SELECT *  FROM OBCD T0 WHERE T0.""ItemCode"" ='" & objform.Items.Item("5").Specific.Value & "'"
                                    '            oRec = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                                    '            oRec.DoQuery(StrQry)
                                    '            If oRec.RecordCount = 0 Then
                                    '                StrQry = "SELECT * FROM ""@NT_BRST"""
                                    '                oRec = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                                    '                oRec.DoQuery(StrQry)
                                    '                If oRec.RecordCount > 0 Then
                                    '                    Dim strlen, RevStr As String
                                    '                    Dim AryCalc() As String
                                    '                    Dim calcvalue As Double = 0
                                    '                    strlen = oRec.Fields.Item("U_Start1").Value.ToString()
                                    '                    RevStr = StrReverse(strlen)
                                    '                    strlen = strlen.Length

                                    '                    Dim JumpNumber As Integer = 3
                                    '                    For i As Integer = 0 To RevStr.Length - 1
                                    '                        If JumpNumber = 3 Then
                                    '                            calcvalue = calcvalue + (Convert.ToDouble(RevStr.Substring(i, 1)) * JumpNumber)
                                    '                            JumpNumber = 1
                                    '                        Else
                                    '                            calcvalue = calcvalue + (Convert.ToDouble(RevStr.Substring(i, 1)) * JumpNumber)
                                    '                            JumpNumber = 3
                                    '                        End If
                                    '                    Next
                                    '                    calcvalue = calcvalue
                                    '                    AryCalc = Convert.ToString(calcvalue).Split(".")
                                    '                    Try
                                    '                        calcvalue = calcvalue / Convert.ToDouble(strlen)

                                    '                        AryCalc = Convert.ToString(calcvalue).Split(".")
                                    '                        If Convert.ToDouble(AryCalc(1)) > 0 Then
                                    '                            calcvalue = Convert.ToDouble(strlen) - Convert.ToDouble(AryCalc(1))
                                    '                        End If
                                    '                    Catch ex As Exception
                                    '                        calcvalue = 0
                                    '                    End Try

                                    '                    'If Convert.ToInt32(strlen) = 1 Then
                                    '                    '    oAutoBarcode = oRec.Fields.Item("U_Itprefix").Value & oRec.Fields.Item("U_Start1").Value & "0" & oRec.Fields.Item("U_End").Value
                                    '                    'Else

                                    '                    'End If
                                    oAutoBarcode = businessObject.Fields.Item("NextBarCode").Value
                                    objform.Items.Item("1470002287").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    objBarcodeForm = objMain.objApplication.Forms.GetForm("1470000020", pVal.FormTypeCount)
                                    objBarcodeMatrix = objBarcodeForm.Items.Item("1470000013").Specific
                                    objBarcodeMatrix.Columns.Item("1470000003").Cells.Item(1).Specific.value = oAutoBarcode

                                    ' Startt = Convert.ToUInt64(oRec.Fields.Item("U_Start1").Value) + 13

                                    'Endd = Convert.ToUInt64(oRec.Fields.Item("U_End").Value) + 2
                                    '    StrQry = "UPDATE ""@NT_BRST"" SET ""U_Start1"" ='" & Startt & "',""U_End"" ='" & Endd & "' Where ""Code""='" & oRec.Fields.Item("Code").Value & "' "
                                    '    oRec = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                                    '    oRec.DoQuery(StrQry)
                                    'Else
                                    '    objMain.objApplication.StatusBar.SetText("Barcode Setup details missing for this ItemCode or ItemGroup...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If

                                '                Else
                                'objMain.objApplication.StatusBar.SetText("Barcode Already Exist...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                        Catch ex As Exception

                        End Try
                    End If


                    'madhu
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                    If pVal.ItemUID = "U_REBEGBLE" And pVal.BeforeAction = False Then
                        objform = objMain.objApplication.Forms.Item(FormUID)
                        Dim oCombo As SAPbouiCOM.ComboBox = CType(objform.Items.Item("U_REBEGBLE").Specific, SAPbouiCOM.ComboBox)
                        Dim oQtyItem As SAPbouiCOM.Item = objform.Items.Item("U_REBQTY")
                        Dim oQty As SAPbouiCOM.EditText = CType(oQtyItem.Specific, SAPbouiCOM.EditText)
                        If oCombo.Selected.Value = "Yes" Then
                            oQtyItem.Enabled = True
                        Else
                            oQtyItem.Enabled = False
                        End If
                    End If



                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    Try
                        objform = objMain.objApplication.Forms.Item(FormUID)
                        'If pVal.ItemUID = "107" And pVal.Before_Action = False And objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        '    Sql = "SELECT T0.""ItemCode"" FROM OBCD T0 WHERE T0.""BcdCode"" ='" & objform.Items.Item("107").Specific.Value & "'"
                        '    oRec = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                        '    oRec.DoQuery(Sql)
                        '    If oRec.RecordCount > 0 Then
                        '        objform.Items.Item("107").Specific.Value = ""
                        '        objform.Items.Item("5").Specific.Value = oRec.Fields.Item("ItemCode").Value
                        '        objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        '    End If
                        'End If
                    Catch ex As Exception

                    End Try
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    'objBarcodeForm = objMain.objApplication.Forms.GetForm("10010196", pVal.FormTypeCount)
                    'objBarcodeMatrix = objBarcodeForm.Items.Item("5").Specific

                    If pVal.ItemUID = "1470002287" And pVal.BeforeAction = False Then
                        Try

                            If CStr(objform.Items.Item("U_AcceptWMQty").Specific.Value).Trim <> "" Then
                                objMain.InteNum = CInt(CStr(objform.Items.Item("U_AcceptWMQty").Specific.Value).Trim)
                            End If

                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If
                    If pVal.ItemUID = "Btn_BrPrc" And pVal.BeforeAction = False Then
                        'Me.LoadBranchWisePriceList(objform.UniqueID)

                        Dim itemCode As String = objform.Items.Item("5").Specific.Value
                        Me.LOADBRANCHWISEPRICELIST(objform.UniqueID, itemCode)
                    End If

                    If pVal.ItemUID = "Btn_BR" And pVal.BeforeAction = False Then
                        Dim ItemCode As String = objform.Items.Item("5").Specific.Value
                        objMain.objBarcodeIM.CreateForm(objform.UniqueID, ItemCode)
                    End If
                    'If pVal.BeforeAction = False And pVal.ItemUID = "10002059" Then
                    '    objMain.objApplication.ActivateMenuItem("44544")
                    '    Dim objBaseForm As SAPbouiCOM.Form
                    '    objBaseForm = objMain.objApplication.Forms.GetForm("10010196", objMain.objApplication.Forms.ActiveForm.TypeCount)
                    '    'objBaseForm.Items.Item("5").Specific.Value = oDBs_Head.GetValue("U_CLCode", oDBs_Head.Offset)
                    '    'objBaseForm.Items.Item("234000181").Specific.Value = oDBs_Head.GetValue("Code", oDBs_Head.Offset)
                    'End If
                    'If pVal.ItemUID = "Btn_GRBR" And pVal.BeforeAction = False Then
                    '    Me.Generate_Barcode(objform.UniqueID)
                    '    'Me.StandardBarcodeGenerate()
                    'End If
                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                    '    If Me.Validation(objform.UniqueID) = False Then BubbleEvent = False
                    'End If
                    If pVal.ItemUID = "Btn_PrUpdt" And pVal.BeforeAction = False Then
                        oDBs_Head = objform.DataSources.DBDataSources.Item("OITM")

                        Dim chkForAuthorization As String = " Select ""U_ACCESS"" , ""U_PRICEDIT"" FROM OUSR T0 Inner Join ""@SBO_PUIGA"" T1 On ""USER_CODE"" = ""U_USRCD"" " &
                                                            " Inner Join ""@SBO_PUIGA_C0"" T2 On T1.""DocEntry"" = T2.""DocEntry"" " &
                                                            " Where IFNULL(""U_ACCESS"",'N')='Y' And ""USER_CODE"" = '" & objMain.objCompany.UserName & "' and ""U_ITMGRP"" = '" & oDBs_Head.GetValue("ItmsGrpCod", 0) & "'"
                        Dim oRschkForAuthorization As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRschkForAuthorization.DoQuery(chkForAuthorization)
                        If oRschkForAuthorization.RecordCount > 0 Then
                            objMain.objApplication.StatusBar.SetText("User not authorized to update price..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Sub
                        End If

                        Dim CheckIfTCExist As String = "Select TO_VARCHAR (""DocNum"") As ""U_PRCUPNO"" From ""@ME_PRICEUPD"" Where ""U_ITMCD""='" & oDBs_Head.GetValue("ItemCode", 0) & "'"
                        Dim oRsCheckIfTCExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsCheckIfTCExist.DoQuery(CheckIfTCExist)
                        If oRsCheckIfTCExist.Fields.Item("U_PRCUPNO").Value <> "" Then
                            objMain.objPriceUpdate.CreateForm(objform.UniqueID, oDBs_Head.GetValue("ItemCode", 0), oDBs_Head.GetValue("CodeBars", 0), oDBs_Head.GetValue("ItemName", 0), "Y", oRsCheckIfTCExist.Fields.Item("U_PRCUPNO").Value, oDBs_Head.GetValue("ItmsGrpCod", 0))
                        Else
                            objMain.objPriceUpdate.CreateForm(objform.UniqueID, oDBs_Head.GetValue("ItemCode", 0), oDBs_Head.GetValue("CodeBars", 0), oDBs_Head.GetValue("ItemName", 0), "N", "", oDBs_Head.GetValue("ItmsGrpCod", 0))
                        End If
                    End If
                    If pVal.ItemUID = "fldRtl" Then
                        objform.PaneLevel = "48"
                        'Me.Matter(objform.UniqueID)
                    End If
                    If pVal.ItemUID = "fldPrice" And pVal.BeforeAction = False Then
                        objform.PaneLevel = "49"
                        'Me.Items(objForm.UniqueID)
                        Me.PriceDetails()
                    End If
                    'Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    '    objform = objMain.objApplication.Forms.Item(FormUID)
                    '    If pVal.ItemUID = "7" And pVal.BeforeAction = False And objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                    '        Me.PriceDetails()

                    '    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub PriceDetails()
        Try
            oDBs_Head = objform.DataSources.DBDataSources.Item("OITM")
            ''objMatrix = objform.Items.Item("38").Specific
            Dim objGrid As SAPbouiCOM.Grid
            objGrid = objform.Items.Item("Grid").Specific
            Dim BranchID As String
            Dim ItemCode As String
            Dim ItemCode2 As String
            Dim ItemCode3 As String
            Dim User As String = objMain.objCompany.UserName

            Dim ItemGroupQuery As String = "SELECT * FROM USR6 where ""UserCode""='" & User & "'"
            Dim oRsItemGroup As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRsItemGroup.DoQuery(ItemGroupQuery)

            ItemCode = objform.Items.Item("5").Specific.Value
            Dim BranchIDs As String = ""

            While oRsItemGroup.EoF = False

                If BranchIDs <> "" Then
                    BranchIDs &= ","
                End If

                BranchIDs &= oRsItemGroup.Fields.Item("BPLId").Value.ToString()

                oRsItemGroup.MoveNext()

            End While

            Dim Details As String = "CALL ""GET_ITEM_PRICE_DETAILS""('" & ItemCode & "','" & BranchIDs & "')"

            objform.DataSources.DataTables.Item(0).ExecuteQuery(Details)


            objGrid.DataTable = objform.DataSources.DataTables.Item("MyDataTable")
            objGrid.AutoResizeColumns()
            objGrid.Columns.Item(0).Editable = False
            objGrid.Columns.Item(1).Editable = False
            objGrid.Columns.Item(2).Editable = False
            objGrid.Columns.Item(3).Editable = False
            objGrid.Columns.Item(4).Editable = False
            objGrid.Columns.Item(5).Editable = False
            objGrid.Columns.Item(6).Editable = False
            objGrid.Columns.Item(7).Editable = False
            objGrid.Columns.Item(8).Editable = False
            objGrid.Columns.Item(9).Editable = False
            objGrid.Columns.Item(10).Editable = False
            objGrid.Columns.Item(11).Editable = False
            objGrid.Columns.Item(12).Editable = False
            objGrid.Columns.Item(13).Editable = False
            objGrid.Columns.Item(14).Editable = False
            If objGrid.Rows.Count = 0 Then
                objMain.objApplication.MessageBox("No records found")
                Exit Sub
            End If

        Catch ex As Exception
            'objForm.Close()
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "1282" And pVal.BeforeAction = False Then


                If objform.Items.Item("U_REBEGBLE").Specific.Value = "No" Then
                    Me.objform.Items.Item("U_REBQTY").Enabled = False
                Else
                    Me.objform.Items.Item("U_REBQTY").Enabled = True
                End If

            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
    Private Sub AddItems(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)

            objMain.objUtilities.AddButtonH(objform.UniqueID, "Btn_BrPrc", objform.Items.Item("2").Top,
                                            objform.Items.Item("2").Left + objform.Items.Item("2").Width + 5, objform.Items.Item("2").Width + 30, "2", "&Retail Price", 0, 0, objform.Items.Item("2").Height)
            objform.Items.Item("Btn_BrPrc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objform.Items.Item("Btn_BrPrc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objMain.objUtilities.AddButtonH(objform.UniqueID, "Btn_PrUpdt", objform.Items.Item("Btn_BrPrc").Top,
                                            objform.Items.Item("Btn_BrPrc").Left + objform.Items.Item("Btn_BrPrc").Width + 5, objform.Items.Item("Btn_BrPrc").Width + 30, "Btn_BrPrc", "&Price Update", 0, 0, objform.Items.Item("2").Height)
            objform.Items.Item("Btn_PrUpdt").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objform.Items.Item("Btn_PrUpdt").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objMain.objUtilities.AddButtonH(objform.UniqueID, "Btn_BR", objform.Items.Item("2").Top,
                                            objform.Items.Item("Btn_PrUpdt").Left + objform.Items.Item("Btn_PrUpdt").Width + 5, objform.Items.Item("Btn_PrUpdt").Width, "2", "&Load Barcode", 0, 0, objform.Items.Item("2").Height)
            objform.Items.Item("Btn_BR").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objform.Items.Item("Btn_BR").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objMain.objUtilities.AddButtonH(objform.UniqueID, "Btn_GRBR", objform.Items.Item("2").Top,
            '                                objform.Items.Item("Btn_BR").Left + objform.Items.Item("Btn_BR").Width + 5, objform.Items.Item("Btn_BR").Width, "2", "&Generate Barcode", 0, 0, objform.Items.Item("2").Height)
            'objform.Items.Item("Btn_GRBR").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objform.Items.Item("Btn_GRBR").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objform.Items.Item("Btn_GRBR").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            oItem = objform.Items.Add("fldRtl", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem.Top = objform.Items.Item("1320002069").Top
            oItem.Left = objform.Items.Item("1320002069").Left + objform.Items.Item("1320002069").Width + 10
            oItem.Width = objform.Items.Item("1320002069").Width + 30
            oItem.Height = objform.Items.Item("1320002069").Height
            oItem.AffectsFormMode = False
            Dim fld As SAPbouiCOM.Folder = objform.Items.Item("fldRtl").Specific
            fld.Caption = "Retail"
            fld.GroupWith("1320002069")
            fld.Pane = "48"

            objMain.objUtilities.AddLabel(objform.UniqueID, "lblPC", objform.Items.Item("76").Top, objform.Items.Item("76").Left + 5, objform.Items.Item("76").Width, "Price Update Number", "84", objform.Items.Item("76").Height, 48, 48)
            objMain.objUtilities.AddEditBox(objform.UniqueID, "PRCUPNO", objform.Items.Item("80").Top, objform.Items.Item("80").Left + 10, objform.Items.Item("80").Width, "OITM", "U_PRCUPNO", "lblPC", objform.Items.Item("80").Height, 48, 48)

            objMain.objUtilities.AddLabel(objform.UniqueID, "lblNC", objform.Items.Item("76").Top + 22, objform.Items.Item("76").Left + 5, objform.Items.Item("76").Width, "No of PC per Pack", "84", objform.Items.Item("76").Height, 48, 48)
            objMain.objUtilities.AddEditBox(objform.UniqueID, "PcPerPac", objform.Items.Item("80").Top + 22, objform.Items.Item("80").Left + 10, objform.Items.Item("80").Width, "OITM", "U_PcPerPac", "lblNC", objform.Items.Item("80").Height, 48, 48)

            objMain.objUtilities.AddLabel(objform.UniqueID, "lblBC", objform.Items.Item("76").Top + 44, objform.Items.Item("76").Left + 5, objform.Items.Item("76").Width, "Base Unit", "84", objform.Items.Item("76").Height, 48, 48)
            objMain.objUtilities.AddEditBox(objform.UniqueID, "BaseUnit", objform.Items.Item("80").Top + 44, objform.Items.Item("80").Left + 10, objform.Items.Item("80").Width, "OITM", "U_BaseUnit", "lblBC", objform.Items.Item("80").Height, 48, 48)

            objMain.objUtilities.AddLabel(objform.UniqueID, "lbUOMN", objform.Items.Item("76").Top + 66, objform.Items.Item("76").Left + 5, objform.Items.Item("76").Width, "Uom Dimension", "84", objform.Items.Item("76").Height, 48, 48)
            objMain.objUtilities.AddComboBox(objform.UniqueID, "UOM", objform.Items.Item("80").Top + 66, objform.Items.Item("80").Left + 10, objform.Items.Item("80").Width, objform.Items.Item("80").Height, "OITM", "U_UOM", 48, 48)


            objMain.objUtilities.AddLabel(objform.UniqueID, "lblIS", objform.Items.Item("76").Top + 88, objform.Items.Item("76").Left + 5, objform.Items.Item("76").Width, "Item Size", "84", objform.Items.Item("76").Height, 48, 48)
            objMain.objUtilities.AddEditBox(objform.UniqueID, "ItemSize", objform.Items.Item("80").Top + 88, objform.Items.Item("80").Left + 10, objform.Items.Item("80").Width, "OITM", "U_ItemSize", "lblIS", objform.Items.Item("80").Height, 48, 48)

            objMain.objUtilities.AddLabel(objform.UniqueID, "EXCVATDP", objform.Items.Item("68").Top + 15, objform.Items.Item("68").Left + 0, objform.Items.Item("68").Width, "Excise Vat", "EXCVAT", objform.Items.Item("68").Height, 2, 2)
            objMain.objUtilities.AddEditBox(FormUID, "EXCVAT", objform.Items.Item("69").Top + 15, objform.Items.Item("69").Left + 0, objform.Items.Item("69").Width, "OITM", "U_EXCVAT", "EXCVATDP", objform.Items.Item("69").Height, 2, 2)


            Dim pane As Integer = objform.Items.Item("161").FromPane

            'objMain.objUtilities.AddLabel(objform.UniqueID, "U_lblREB", objform.Items.Item("161").Top + 20, objform.Items.Item("161").Left, objform.Items.Item("161").Width, "Rebate Eligible", "161", objform.Items.Item("161").Height, pane, pane)
            'objMain.objUtilities.AddComboBox(objform.UniqueID, "U_REBEGBLE", objform.Items.Item("U_lblREB").Top, objform.Items.Item("U_lblREB").Left + objform.Items.Item("U_lblREB").Width + 10, +objform.Items.Item("U_lblREB").Width, objform.Items.Item("U_lblREB").Height, "OITM", "U_REBEGBLE", pane, pane)
            'objMain.objUtilities.AddLabel(objform.UniqueID, "U_lblRBQ", objform.Items.Item("161").Top + 50, objform.Items.Item("161").Left, objform.Items.Item("161").Width, "Rebate Qty", "161", objform.Items.Item("161").Height, pane, pane)
            'objMain.objUtilities.AddEditBox(objform.UniqueID, "U_REBQTY", objform.Items.Item("162").Top + 50, objform.Items.Item("162").Left, objform.Items.Item("162").Width, "OITM", "U_REBQTY", "U_lblRBQ", objform.Items.Item("162").Height, pane, pane)

            Dim baseTop As Integer = objform.Items.Item("161").Top
            Dim lblHeight As Integer = objform.Items.Item("161").Height

            ' Rebate Eligible
            Try


                objMain.objUtilities.AddLabel(objform.UniqueID, "U_lblREB", baseTop + lblHeight, objform.Items.Item("161").Left, objform.Items.Item("161").Width - 10, "Rebate Eligible", "161", lblHeight, pane, pane)
                objMain.objUtilities.AddComboBox(objform.UniqueID, "U_REBEGBLE", baseTop + lblHeight, objform.Items.Item("162").Left, objform.Items.Item("162").Width, objform.Items.Item("162").Height, "OITM", "U_REBEGBLE", pane, pane)
            Catch ex As Exception

            End Try
            ' Rebate Qty (next line using height again)
            Try


                objMain.objUtilities.AddLabel(objform.UniqueID, "U_lblRBQ", baseTop + (lblHeight * 2), objform.Items.Item("161").Left, objform.Items.Item("161").Width, "Rebate Qty", "161", lblHeight, pane, pane)
                objMain.objUtilities.AddEditBox(objform.UniqueID, "U_REBQTY", baseTop + (lblHeight * 2), objform.Items.Item("162").Left, objform.Items.Item("162").Width, "OITM", "U_REBQTY", "U_lblRBQ", objform.Items.Item("162").Height, pane, pane) ' Try : objform.Items.Item("U_REBQTY").Enabled = (objform.DataSources.DBDataSources.Item("OITM").GetValue("U_REBEGBLE", 0).Trim() <> "No") : Catch : End Try
            Catch ex As Exception

            End Try
            'Try
            '    Try : objform.Items.Item("U_lblUP") : Catch : objMain.objUtilities.AddLabel(objform.UniqueID, "U_lblUP", baseTop + (lblHeight * 3) - 40, objform.Items.Item("161").Left + 1466, objform.Items.Item("161").Width, "Under Promotion", "161", lblHeight, pane, pane) : End Try

            '    Try : objform.Items.Item("UnderPro") : Catch : objMain.objUtilities.AddComboBox(objform.UniqueID, "UnderPro", baseTop + (lblHeight * 3) - 40, objform.Items.Item("162").Left + 464, objform.Items.Item("162").Width, objform.Items.Item("162").Height, "OITM", "U_UNDERPRO", pane, pane) : End Try
            'Catch ex As Exception


            'End Try
            Dim oRef As SAPbouiCOM.Item = objform.Items.Item("10002050")

            'Rebate Eligible
            'objMain.objUtilities.AddLabel(objform.UniqueID, "U_lblRBE", oRef.Top - 40, oRef.Left, 120, "Rebate Eligible", "254000001", lblHeight, pane, pane)
            'objMain.objUtilities.AddComboBox(objform.UniqueID, "U_REBEGBLE", oRef.Top - 40, oRef.Left + 195, objform.Items.Item("162").Width, objform.Items.Item("162").Height, "OITM", "U_REBEGBLE", pane, pane)

            'Under Promotion
            Try
                'objMain.objUtilities.AddLabel(objform.UniqueID, "U_lblUP", oRef.Top - 50, oRef.Left - 10, 150, "Under Promotion", "10002050", lblHeight, pane, pane)

                'objMain.objUtilities.AddComboBox(objform.UniqueID, "UnderPro", oRef.Top - 50, oRef.Left + 195, objform.Items.Item("162").Width, objform.Items.Item("162").Height, "OITM", "U_UNDERPRO", pane, pane)
                objMain.objUtilities.AddLabel(objform.UniqueID, "U_lblUP", oRef.Top - 50, oRef.Left + 5, 150, "Under Promotion", "10002050", lblHeight, pane, pane)

                objMain.objUtilities.AddComboBox(objform.UniqueID, "UnderPro", oRef.Top - 50, oRef.Left + 195, objform.Items.Item("162").Width, objform.Items.Item("162").Height, "OITM", "U_UNDERPRO", pane, pane)
            Catch ex As Exception

            End Try


            'Dim oRefItem As SAPbouiCOM.Item = objform.Items.Item("1320002069") 'Attachments tab

            'oItem = objform.Items.Add("fldPrice", SAPbouiCOM.BoFormItemTypes.it_FOLDER)

            'oItem.Top = oRefItem.Top
            '''oItem.Left = oRefItem.Left + oRefItem.Width + 10
            'oItem.Left = Math.Min(
            '    oRefItem.Left + oRefItem.Width + 5,
            '    objform.Width - oRefItem.Width - 30)
            'oItem.Width = oRefItem.Width
            'oItem.Height = oRefItem.Height
            'oItem.AffectsFormMode = False

            'Dim fld1 As SAPbouiCOM.Folder = oItem.Specific
            'fld1.Caption = "Price Details"
            'fld1.GroupWith("1320002069")
            'fld1.Pane = 49

            'oItem = objform.Items.Add("Grid", SAPbouiCOM.BoFormItemTypes.it_GRID)
            'oItem.Top = objform.Items.Item("28").Top
            'oItem.Left = objform.Items.Item("28").Left

            'oItem.Width = objform.Items.Item("28").Width
            'oItem.Height = objform.Items.Item("28").Height
            'oItem.FromPane = 49
            'oItem.ToPane = 49
            'objform.DataSources.DataTables.Add("MyDataTable")

            'For Each itm As SAPbouiCOM.Item In objform.Items
            '    If itm.Type = SAPbouiCOM.BoFormItemTypes.it_FOLDER Then
            '        Try
            '            itm.Width = 5
            '        Catch
            '        End Try
            '    End If
            'Next


            Dim oRefItem As SAPbouiCOM.Item = objform.Items.Item("1320002069") ' Attachments tab

            oItem = objform.Items.Add("fldPrice", SAPbouiCOM.BoFormItemTypes.it_FOLDER)

            oItem.Top = oRefItem.Top
            oItem.Left = Math.Min(
                oRefItem.Left + oRefItem.Width + 5,
                objform.Width - oRefItem.Width - 30)

            oItem.Width = oRefItem.Width
            oItem.Height = oRefItem.Height
            oItem.AffectsFormMode = False

            Dim fld1 As SAPbouiCOM.Folder = CType(oItem.Specific, SAPbouiCOM.Folder)
            fld1.Caption = "Price Details"
            fld1.GroupWith("1320002069")
            fld1.Pane = 49


            Dim FolderUIDs As String() = {"163", "3", "4", "26", "27", "1880000001", "11", "1470002136", "1470002137", "1470002138", "1470002139", "1470002140", "1470002221", "1980000483", "fldRtl", "9", "1320002069", "fldPrice"}

            Dim LeftPos As Integer = objform.Items.Item("163").Left

            For Each uid As String In FolderUIDs
                Try
                    Dim itm As SAPbouiCOM.Item = objform.Items.Item(uid)

                    itm.Left = LeftPos

                    If uid = "fldPrice" Then
                        itm.Width = 80
                    Else
                        itm.Width = 55
                    End If

                    LeftPos += itm.Width + 2

                Catch
                End Try
            Next



            oItem = objform.Items.Add("Grid", SAPbouiCOM.BoFormItemTypes.it_GRID)
            oItem.Top = objform.Items.Item("28").Top - 200
            oItem.Left = objform.Items.Item("28").Left

            oItem.Width = objform.Items.Item("28").Width
            oItem.Height = objform.Items.Item("28").Height
            oItem.FromPane = 49
            oItem.ToPane = 49
            objform.DataSources.DataTables.Add("MyDataTable")




            'Dim msg As String = ""

            'For Each itm As SAPbouiCOM.Item In objform.Items
            '    If itm.Type = SAPbouiCOM.BoFormItemTypes.it_FOLDER Then
            '        Dim f As SAPbouiCOM.Folder = CType(itm.Specific, SAPbouiCOM.Folder)
            '        msg &= itm.UniqueID & "  -->  " & f.Caption & vbCrLf
            '    End If
            'Next

            'MessageBox.Show(msg)



            ' objform.Items.Item("U_REBQTY").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objform.Items.Item("Btn_GRBR").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objMain.objUtilities.AddLabel(objform.UniqueID, "lblWC", objform.Items.Item("76").Top + 66, objform.Items.Item("76").Left + 5, objform.Items.Item("76").Width, "Weigh Item", "84", objform.Items.Item("76").Height, 48, 48)
            'objMain.objUtilities.AddEditBox(objform.UniqueID, "AcceptWMQty", objform.Items.Item("80").Top + 66, objform.Items.Item("80").Left + 10, objform.Items.Item("80").Width, "OITM", "U_AcceptWMQty", "lblWC", objform.Items.Item("80").Height, 48, 48)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub LOADBRANCHWISEPRICELIST(FORMUID As String, ITEMCODE As String)

        Try
            objform = objMain.objApplication.Forms.Item(FORMUID)

            Dim CP As SAPbouiCOM.FormCreationParams = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            Dim OITEM As SAPbouiCOM.Item
            Dim OGRID As SAPbouiCOM.Grid
            Dim OBUTTON As SAPbouiCOM.Button

            CP.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            CP.UniqueID = "LBWPLD_FORM_" & objMain.objApplication.Forms.Count
            CP.FormType = "LBWPLD_FORM"

            objListForm = objMain.objApplication.Forms.AddEx(CP)
            objListForm.Height = 320
            objListForm.Width = 600
            objListForm.Title = "LOAD BRANCH WISE PRICELIST DETAILS"

            ' GRID
            OITEM = objListForm.Items.Add("MYGRID", SAPbouiCOM.BoFormItemTypes.it_GRID)
            OITEM.Left = 20
            OITEM.Top = 20
            OITEM.Width = 550
            OITEM.Height = 200
            OGRID = OITEM.Specific

            objListForm.DataSources.DataTables.Add("MYDATATABLE")

            ' BUTTONS
            OITEM = objListForm.Items.Add("OK", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            OITEM.Left = 20
            OITEM.Top = 230
            OITEM.Width = 65
            OITEM.Height = 20
            OBUTTON = OITEM.Specific
            OBUTTON.Caption = "OK"

            OITEM = objListForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            OITEM.Left = 90
            OITEM.Top = 230
            OITEM.Width = 65
            OITEM.Height = 20

            If String.IsNullOrEmpty(ITEMCODE) Then
                objMain.objApplication.MessageBox("ITEMCODE IS EMPTY")
                Exit Sub
            End If

            Dim RS As SAPbobsCOM.Recordset
            RS = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim SQUERY As String = ""
            SQUERY = "CALL ""GET_ITEM_PRICING_DETAILS"" ('" & ITEMCODE & "')"

            RS.DoQuery(SQUERY)
            Dim DT As SAPbouiCOM.DataTable = objListForm.DataSources.DataTables.Item("MYDATATABLE")
            DT.Rows.Clear()

            For I As Integer = 0 To RS.Fields.Count - 1
                If DT.Columns.Count <= I Then
                    DT.Columns.Add(RS.Fields.Item(I).Name, SAPbouiCOM.BoFieldsType.ft_Text)
                End If
            Next
            While Not RS.EoF
                DT.Rows.Add()
                For I As Integer = 0 To RS.Fields.Count - 1
                    DT.SetValue(I, DT.Rows.Count - 1, RS.Fields.Item(I).Value.ToString())
                Next
                RS.MoveNext()
            End While

            OGRID.DataTable = DT
            OGRID.AutoResizeColumns()
            OGRID.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None

            For I As Integer = 0 To OGRID.Columns.Count - 1
                OGRID.Columns.Item(I).Editable = False
            Next

            OGRID.Item.Enabled = False

            If OGRID.Rows.Count = 0 Then
                objMain.objApplication.MessageBox("NO RECORDS FOUND")
                Exit Sub
            End If
            objListForm.Visible = True
        Catch EX As Exception

            If objListForm IsNot Nothing Then
                objListForm.Close()
            End If

            objMain.objApplication.StatusBar.SetText("LOADING BRANCH WISE PRICELIST FAILED: " & EX.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        End Try

    End Sub
    Private Sub FileGeneration(ByVal pVal As SAPbouiCOM.ItemEvent)
        Dim str As String = Nothing
        Dim str1 As String = Nothing
        Dim businessObject As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(300), SAPbobsCOM.Recordset)
        Dim recordset As SAPbobsCOM.Recordset = DirectCast(objMain.objCompany.GetBusinessObject(300), SAPbobsCOM.Recordset)
        Dim str2 As String = Nothing
        Try
            Try
                objMain.objApplication.StatusBar.SetText("Pls Wait.......Price Detail text file generation inprogress", 20, 1)
                Dim BranchQry As String = "SELECT T2.""BPLId"",T2.""BPLName"" AS ""Branch"" FROM USR6 T0  INNER JOIN OUSR T1 ON T0.""UserCode"" = T1.""USER_CODE"" " &
                  " INNER JOIN OBPL T2 ON T0.""BPLId"" = T2.""BPLId"" WHERE T1.""USER_CODE"" = '" & objMain.objCompany.UserName.ToString().Trim & "' "
                Dim oRsBranch As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsBranch.DoQuery(BranchQry)
                Dim BranchID As String = oRsBranch.Fields.Item("BPLId").Value
                businessObject.DoQuery("Call ""PriceList_TextFile_Generation""('" & BranchID & "') ")
                If (businessObject.RecordCount > 0) Then
                    str1 = ""
                    str = Nothing

                    recordset.DoQuery("Select ifnull(""U_FileLoc"",'') as ""Location"" From OBPL Where ""BPLId""='" & BranchID & "'")
                    recordset.MoveFirst()

                    If recordset.Fields.Item(0).Value <> "" Then
                        str1 = recordset.Fields.Item(0).Value.ToString()
                        Dim recordCount As Integer = businessObject.RecordCount
                        Dim num As Integer = 1
                        Do
                            str2 = Nothing
                            If (Not Operators.ConditionalCompareObjectGreater(businessObject.Fields.Item("Rate1").Value, 0, False)) Then
                                str2 = String.Concat(Conversions.ToString(Conversions.ToInteger(businessObject.Fields.Item("Rate").Value)), ".")
                            Else
                                Dim type As System.Type = GetType(Math)
                                Dim objectValue(2) As Object

                                objectValue(0) = RuntimeHelpers.GetObjectValue(businessObject.Fields.Item("Rate").Value)
                                objectValue(1) = 2
                                objectValue(2) = MidpointRounding.AwayFromZero
                                Dim objArray As Object() = objectValue
                                Dim flagArray() As Boolean = {True, False, False}
                                Dim obj As Object = NewLateBinding.LateGet(Nothing, type, "Round", objArray, Nothing, Nothing, flagArray)
                                str2 = Conversions.ToString(obj)
                            End If
                            str = Nothing
                            Dim strArrays() As String = {Conversions.ToString(businessObject.Fields.Item("EANCode").Value), ",", Conversions.ToString(businessObject.Fields.Item("ITEMNAME").Value), ",", Conversions.ToString(businessObject.Fields.Item("UomCode").Value), ",", str2, ","}
                            Me.write_log(String.Concat(strArrays), str1, num)
                            businessObject.MoveNext()
                            num = num + 1
                        Loop While num <= recordCount
                        objMain.objApplication.StatusBar.SetText("Price Detail text file generation Completed Successfully", 20, 3)
                    Else
                        objMain.objApplication.StatusBar.SetText("Please Update Text File Location in Company Detail", 20, 2)
                        Return
                    End If
                End If
            Catch exception As System.Exception
                ProjectData.SetProjectError(exception)
                ProjectData.ClearProjectError()
            End Try
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub
    Public Sub write_log(ByVal status As String, ByVal Location As String, ByVal Time As Integer)
        Dim fileStream As System.IO.FileStream
        Dim str As String = Location
        Try
            Try
                Dim directoryInfo As System.IO.DirectoryInfo = New System.IO.DirectoryInfo(String.Concat(str, "Scale"))
                Dim str1 As String = String.Concat(str, "Scale.txt")
                If (Time = 1) Then
                    If (Not File.Exists(str1)) Then
                        fileStream = New System.IO.FileStream(str1, FileMode.Create, FileAccess.Write)
                        fileStream.Close()
                    Else
                        File.Delete(str1)
                        fileStream = New System.IO.FileStream(str1, FileMode.Create, FileAccess.Write)
                        fileStream.Close()
                    End If
                End If
                Dim streamWriter As System.IO.StreamWriter = New System.IO.StreamWriter(str1, True)
                If (Operators.CompareString(status, "", False) <> 0) Then
                    streamWriter.WriteLine(status)
                End If
                streamWriter.Close()
            Catch exception1 As System.Exception
                ProjectData.SetProjectError(exception1)
                Dim exception As System.Exception = exception1
                Interaction.MsgBox(String.Concat("createlog :", exception.ToString()), MsgBoxStyle.OkOnly, Nothing)
                ProjectData.ClearProjectError()
            End Try
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objform = objMain.objApplication.Forms.GetForm("150", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        'Me.Items(objform.UniqueID)
                        objMain.objUtilities.GlobalItemNumber = objform.Items.Item("5").Specific.Value
                        Dim val As String = ""

                        Try
                            Try : val = objform.DataSources.DBDataSources.Item("OITM").GetValue("U_REBEGBLE", 0).Trim() : Catch : End Try
                            Try : objform.Items.Item("U_REBQTY").Enabled = (val <> "No") : Catch : End Try
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                        Try
                            Try : Me.objform.Items.Item("U_REBQTY").Enabled = (val <> "No") : Catch : End Try
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try

                        '---MAdhu Based on Promotion documnet status in header and Active  field in cchild under promotion field is Yes or No
                        'Try
                        '    objform.Items.Item("U_lblUP")
                        'Catch
                        '    objMain.objUtilities.AddLabel(objform.UniqueID, "U_lblUP", 364, objform.Items.Item("161").Left + 1470, 100, "Under Promotion", "161", 15, 0, 0)
                        'End Try

                        'Try
                        '    objform.Items.Item("UnderPro")
                        'Catch
                        '    objMain.objUtilities.AddComboBox(objform.UniqueID, "UnderPro", 364, objform.Items.Item("162").Left + 466, objform.Items.Item("162").Width, objform.Items.Item("162").Height, "OITM", "U_UNDERPRO", 0, 0)
                        'End Try
                        Dim oRef As SAPbouiCOM.Item = objform.Items.Item("10002050")
                        Dim pane As Integer = 6
                        Dim iTop As Integer = oRef.Top - 50
                        Dim lblHeight As Integer = objform.Items.Item("161").Height


                        Try
                            objform.Items.Item("U_lblUP").Top = oRef.Top - 25
                            objform.Items.Item("U_lblUP").Left = oRef.Left + 150
                        Catch
                            objMain.objUtilities.AddLabel(objform.UniqueID, "U_lblUP", oRef.Top - 25, oRef.Left + 150, 120, "Under Promotion", "10002050", lblHeight, 6, 6)
                        End Try

                        Try
                            objform.Items.Item("UnderPro").Top = oRef.Top - 25
                            objform.Items.Item("UnderPro").Left = oRef.Left + 345
                        Catch
                            objMain.objUtilities.AddComboBox(objform.UniqueID, "UnderPro", oRef.Top - 25, oRef.Left + 345, objform.Items.Item("162").Width, objform.Items.Item("162").Height, "OITM", "U_UNDERPRO", 6, 6)
                        End Try

                        Me.UpdateUnderPromotion(objform)


                        'If val = "No" Then
                        '    Me.objform.Items.Item("U_REBQTY").Enabled = False
                        'Else
                        '    Me.objform.Items.Item("U_REBQTY").Enabled = True
                        'End If

                        '---madhu Based on Promotion documnet status in header and Active  field in cchild under promotion field is Yes or No
                        'Try

                        '    Dim value As String = ""
                        '    Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        '    Dim sql As String = "SELECT CASE " & "WHEN EXISTS ( " & "SELECT 1 " & "FROM ""OBCD"" B " & "INNER JOIN ""@TNX_ITEMDISC_C0"" C0 ON B.""BcdCode"" = C0.""U_BCODE"" " &
                        '           "INNER JOIN ""@TNX_ITEMDISC"" H ON H.""DocEntry"" = C0.""DocEntry"" " & "WHERE B.""ItemCode"" = '" & objform.Items.Item("5").Specific.Value.Trim() & "' " &
                        '            "AND C0.""U_Active"" = 'Yes' " & "AND H.""U_STATUS"" = 'Active' " & ") THEN 'Yes' ELSE 'No' END AS ""UnderPro"" FROM DUMMY"

                        '    oRs.DoQuery(sql)

                        '    Try
                        '        value = oRs.Fields.Item("UnderPro").Value.ToString().Trim()
                        '    Catch

                        '    End Try

                        '    Try
                        '        Dim oCombo As SAPbouiCOM.ComboBox = CType(objform.Items.Item("UnderPro").Specific, SAPbouiCOM.ComboBox)
                        '        ' oCombo.Select(value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                        '        If value = "Yes" Then
                        '            oCombo.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        '        Else
                        '            oCombo.Select("No", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        '        End If
                        '    Catch
                        '        ' UnderPro item not available on form
                        '    End Try

                        'Catch ex As Exception
                        '    objMain.objApplication.StatusBar.SetText(ex.Message)
                        'End Try

                    End If





                    If BusinessObjectInfo.BeforeAction = False AndAlso BusinessObjectInfo.ActionSuccess = True Then

                        Try

                            objMain.objUtilities.GlobalItemNumber = objform.Items.Item("5").Specific.Value

                            Try
                                Dim oGrid As SAPbouiCOM.Item = objform.Items.Item("Grid")
                                Me.PriceDetails()
                            Catch
                                'Grid not available
                            End Try

                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try

                    End If


                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        objMain.objUtilities.GlobalItemNumber = objform.Items.Item("5").Specific.Value
                        If objform.Items.Item("U_REBEGBLE").Specific.Value = "No" Then
                            Me.objform.Items.Item("U_REBQTY").Enabled = False
                        Else
                            Me.objform.Items.Item("U_REBQTY").Enabled = True
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then

                        objMain.objApplication.MessageBox("Item Master Data is Successfully Added")
                    End If



            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Sub Generate_Barcode(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)
            ' Dim specific As Matrix = DirectCast(objMain.objApplication.Forms.ActiveForm.Items.Item("1470000013").Specific, Matrix)
            Dim businessObject As Recordset = DirectCast(objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset), Recordset)
            Dim item As String = objform.Items.Item("5").Specific.Value
            Dim str1 As String = ""

            '---------------------------------------------------------------
            '---------------------------------------------------------------

            Dim DigitsQry As String = "SELECT Case When ""U_AcceptWMQty""='1' Then '6' When ""U_AcceptWMQty""='2' Then '13' End FROM OITM WHERE ""ItemCode"" = '" & item & "'"
            Dim oRsDigits As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRsDigits.DoQuery(DigitsQry)

            If oRsDigits.Fields.Item(0).Value <> "6" Then '------------13-------------------------------------

                Dim ItemGroupQuery As String = "SELECT * FROM ""@NT_BRST""  T2 WHERE T2.""U_Itmgrpnm"" IN " &
                "(SELECT (SELECT T1.""ItmsGrpNam"" FROM OITB T1 WHERE T1.""ItmsGrpCod""=T0.""ItmsGrpCod"") FROM OITM T0 WHERE T0.""ItemCode""='" & item & "') " &
                " And  T2.""U_Scale""=(SELECT ""U_AcceptWMQty"" FROM ""OITM"" Where ""ItemCode""='" & item & "') and  T2.""U_SubGrp"" =(SELECT ""U_VSubGroup"" FROM ""OITM"" Where ""ItemCode""='" & item & "')"
                Dim oRsItemGroup As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                oRsItemGroup.DoQuery(ItemGroupQuery)
                Dim NewBarcode As String = ""
                Dim ItemGroup As String = ""
                Dim SubGroup As String = ""
                If oRsItemGroup.RecordCount = 0 Then
                    ItemGroup = "ALL"
                Else
                    ItemGroup = oRsItemGroup.Fields.Item("U_Itmgrpnm").Value
                    SubGroup = oRsItemGroup.Fields.Item("U_SubGrp").Value
                End If
                'Dim Barcode As String = "SELECT T0.""U_Prefix"", T0.""U_Itprefix"", T0.""U_Start"",T0.""U_Start""+1 AS ""NextNo"" FROM ""@NT_BRST""  T0 WHERE T0.""U_Itmgrpnm"" ='" & ItemGroup & "' and ""U_Scale""='2' "
                Dim Barcode As String = "SELECT T0.""U_Prefix"", T0.""U_Itprefix"", T0.""U_Start"",T0.""U_Start""+1 AS ""NextNo"" FROM ""@NT_BRST""  T0 WHERE T0.""U_Itmgrpnm"" ='" & ItemGroup & "' and T0.""U_SubGrp"" ='" & SubGroup & "'and ""U_Scale""='2' "
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
                        'specific.Columns.Item("1470000003").Cells.Item(specific.VisualRowCount).Specific.value = str1.Trim()
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
                Dim NextBarcode As String = "UPDATE ""@NT_BRST"" T0 SET T0.""U_Start""='" & NextNo & "' WHERE T0.""U_Itmgrpnm"" ='" & ItemGroup & "' and T0.""U_SubGrp"" ='" & SubGroup & "' and ""U_Scale""='2' "
                Dim oRsNextBarcode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                oRsNextBarcode.DoQuery(NextBarcode)

            Else
                Dim str As String = "CALL ""SBO_ProcessBarCode""('" & item & "')"
                businessObject.DoQuery(str)
                If businessObject.RecordCount > 0 Then
                    str1 = businessObject.Fields.Item("NextBarCode").Value
                    ' specific.Columns.Item("1470000003").Cells.Item(specific.VisualRowCount).Specific.value = str1.Trim()
                    If (objMain.objApplication.Forms.ActiveForm.Mode = BoFormMode.fm_UPDATE_MODE) Then
                        objMain.objApplication.Forms.ActiveForm.Items.Item("1470000001").Click(BoCellClickType.ct_Regular)
                    End If
                Else
                    objMain.objApplication.StatusBar.SetText("No records are found. Please check the BarCode Setup", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                End If

            End If

        Catch exception1 As System.Exception
            objMain.objApplication.StatusBar.SetText("Error:" & exception1.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Function GetFieldValue(ByVal query As String) As Object
        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs.DoQuery(query)
        Dim str As String = oRs.Fields.Item(0).Value.ToString()
        Return str
    End Function

    Private Sub UpdateUnderPromotion(ByVal objForm As SAPbouiCOM.Form)

        Try

            Dim itemCode As String = CType(objForm.Items.Item("5").Specific, SAPbouiCOM.EditText).Value.Trim()

            If itemCode = "" Then Exit Sub

            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim sql As String = "SELECT CASE " & "WHEN EXISTS ( " & "SELECT 1 " & "FROM ""OBCD"" B " &
                "INNER JOIN ""@TNX_ITEMDISC_C0"" C0 ON B.""BcdCode"" = C0.""U_BCODE"" " & "INNER JOIN ""@TNX_ITEMDISC"" H ON H.""DocEntry"" = C0.""DocEntry"" " &
                "WHERE B.""ItemCode"" = '" & itemCode & "' " & "AND C0.""U_Active"" = 'Yes' " & "AND H.""U_STATUS"" = 'Active' " & ") THEN 'Yes' ELSE 'No' END AS ""UnderPro"" FROM DUMMY"

            oRs.DoQuery(sql)

            Dim value As String = "No"

            If oRs.RecordCount > 0 Then
                value = oRs.Fields.Item("UnderPro").Value.ToString().Trim()
            End If


            Dim oCombo As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("UnderPro").Specific, SAPbouiCOM.ComboBox)

            If value = "Yes" Then
                oCombo.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
            Else
                oCombo.Select("No", SAPbouiCOM.BoSearchKey.psk_ByValue)
            End If

            objForm.Update()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub

End Class
