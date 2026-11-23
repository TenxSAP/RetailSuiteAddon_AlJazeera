
'from soudia

Imports System.Collections.Specialized
Imports System.Configuration
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices

Public Class clsBarcodeIM

#Region "       Declaration             "
    Public objForm, objSTRForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details1 As SAPbouiCOM.DBDataSource
    Dim objMatrix1 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
#End Region
    'Madhu neww
    Sub CreateForm(ByVal FormUID As String, ByVal ItemCode As String)
        Try
            objMain.objUtilities.LoadForm("BarcodeIM.xml", "frmBarCodeIM", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("frmBarCodeIM", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            'Dim sQury As String = " Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join " & _
            '                      " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""  "

            Dim oRsBranch As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim Query As String = "SELECT T2.""BPLId"",T2.""BPLName"" AS ""Branch"" FROM USR6 T0  INNER JOIN OUSR T1 ON T0.""UserCode"" = T1.""USER_CODE"" " &
                    " INNER JOIN OBPL T2 ON T0.""BPLId"" = T2.""BPLId"" WHERE T1.""USER_CODE"" = '" & objMain.objCompany.UserName.ToString().Trim & "' "
            oRsBranch.DoQuery(Query)

            'If oRsBranch.RecordCount > 1 Then
            '    objMain.objApplication.StatusBar.SetText("Login user have more then one branch.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Exit Sub
            'End If
            objComboBox = objForm.Items.Item("BR").Specific
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, Query)
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("BR").DisplayDesc = True
            If objComboBox.ValidValues.Count <> 0 Then
                Try
                    objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
                Catch ex As Exception
                End Try
            End If
            objForm.Items.Item("txtItem").Specific.Value = ItemCode
            'Print Type
            Dim TypeCombo As SAPbouiCOM.ComboBox = DirectCast(Me.objForm.Items.Item("Type").Specific, SAPbouiCOM.ComboBox)
            TypeCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            If TypeCombo.ValidValues.Count <> 0 Then
                TypeCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            End If

            Dim UOMCmbo As SAPbouiCOM.ComboBox = DirectCast(Me.objForm.Items.Item("Item_4").Specific, SAPbouiCOM.ComboBox)
            Dim UOMQry As String = "SELECT DISTINCT A.""UomEntry"",(SELECT ""UomCode"" FROM OUOM WHERE ""UomEntry""=A.""UomEntry"") AS ""UomName"",  (SELECT ""IUoMEntry"" FROM OITM WHERE ""ItemCode""=A.""ItemCode"") AS ""DefUoM"" FROM OBCD A WHERE A.""ItemCode""='" & ItemCode & "'"
            objMain.objUtilities.ComboBoxLoadValues(UOMCmbo, UOMQry)
            UOMCmbo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            If UOMCmbo.ValidValues.Count <> 0 Then
                UOMCmbo.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
            End If

            Dim checkBox As SAPbouiCOM.CheckBox = DirectCast(Me.objForm.Items.Item("Item_6").Specific, SAPbouiCOM.CheckBox)
            checkBox.Checked = True

            '-----Batch Details------------------
            Dim BatchQuery As String = "SELECT T0.""DistNumber"", T0.""DistNumber"" FROM OBTN T0 WHERE T0.""ItemCode""='" & ItemCode & "'"
            Dim BatchCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("BT").Specific
            objMain.objUtilities.ComboBoxLoadValues(BatchCombo, BatchQuery)
            BatchCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
            objForm.Items.Item("txtPrintQ").Specific.Value = "1"

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ME_BR" And pVal.BeforeAction = False Then
                'Me.CreateForm()
                'Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
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

                    If pVal.ItemUID = "btnOk" And pVal.BeforeAction = False Then
                        'If DateFrom = "" Then
                        '    objForm.StatusBar.SetText("Please enter From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        'ElseIf DateTo = "" Then
                        '    objForm.StatusBar.SetText("Please enter To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        'ElseIf Branch = "" Then
                        '    objForm.StatusBar.SetText("Please enter Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        'Else
                        Me.PrintBarcode(objForm.UniqueID)
                        'End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                    If oCFL.UniqueID = "CA" And pVal.BeforeAction = True Then
                        'Me.CFLFilter("CA")
                    End If
                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        'If oCFL.UniqueID = "CA" Then
                        '    oDBs_Details1.SetValue("U_DepositA", oDBs_Details1.Offset, oDT.GetValue("AcctCode", 0))
                        'End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    If pVal.ItemUID = "Item_4" And pVal.BeforeAction = False Then
                        Dim objUOMCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Item_4").Specific
                        Dim Qry As String = "SELECT ""BcdEntry"",""BcdCode"",IFNULL((SELECT DISTINCT ""BcdEntDft"" FROM ITM12 WHERE ""ItemCode""='" & objForm.Items.Item("txtItem").Specific.Value & "' " &
                            " AND ""UomEntry""='" & objUOMCombo.Selected.Value & "' AND IFNULL(""BcdEntDft"",0)<>0),0) AS ""DefBarCode"" FROM OBCD WHERE ""ItemCode""='" & objForm.Items.Item("txtItem").Specific.Value & "' " &
                            " AND ""UomEntry""='" & objUOMCombo.Selected.Value & "'"
                        Dim objBarcodeCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Item_5").Specific
                        objMain.objUtilities.ComboBoxLoadValues(objBarcodeCombo, Qry)

                        objBarcodeCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                        objForm.Items.Item("Item_5").DisplayDesc = True
                        If objBarcodeCombo.ValidValues.Count <> 0 Then
                            objBarcodeCombo.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
                        End If

                    End If
                    If (pVal.ItemUID = "Type" Or pVal.ItemUID = "BR") And pVal.BeforeAction = False Then
                        Dim objTypeCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Type").Specific
                        Dim objBranchCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("BR").Specific
                        Dim Type As String = objTypeCombo.Selected.Value
                        Dim Branch As String = objBranchCombo.Selected.Value
                        If Type <> "" And Branch <> "" Then
                            Dim Qry As String = "SELECT T0.""U_PrinterName"",T0.""U_PrinterName"" FROM ""@NT_BARPRSETUP"" T0 " &
                                " WHERE T0.""U_BPLId"" ='" & Branch & "' and  T0.""U_Type"" ='" & Type & "'"
                            Dim objPrinterCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Printer").Specific
                            objMain.objUtilities.ComboBoxLoadValues(objPrinterCombo, Qry)

                            objPrinterCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                            objForm.Items.Item("Printer").DisplayDesc = True
                            If objPrinterCombo.ValidValues.Count > 1 Then
                                objPrinterCombo.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
                            End If
                        End If
                    End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "", Optional ByVal Series As Integer = 0)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            objComboBox = objForm.Items.Item("BR").Specific
            Dim GetSeriesQty As String = "Select ""Series"" , ""SeriesName""  From NNM1 Where ""ObjectCode"" = 'ME_OPDC' And ""Locked"" = 'N'"
            objMain.objUtilities.ComboBoxLoadValuesSeries(objForm.Items.Item("BR").Specific, GetSeriesQty)
            objForm.Items.Item("BR").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CFLFilter(ByVal CFL_ID As String)
        Try
            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objForm.ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            oCondition = oConditions.Add()
            oCondition.Alias = "Postable"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "Y"
            oChooseFromList.SetConditions(oConditions)
        Catch ex As Exception

        End Try
    End Sub

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

    Sub PrintBarcode(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            Dim directoryName As String = System.Windows.Forms.Application.StartupPath
            For Each deleteFile In Directory.GetFiles(directoryName, "*.prn", SearchOption.TopDirectoryOnly)
                File.Delete(deleteFile)
            Next

            Dim TypeCobmo As SAPbouiCOM.ComboBox = Me.objForm.Items.Item("Type").Specific
            Dim BarcodeCombo As SAPbouiCOM.ComboBox = DirectCast(Me.objForm.Items.Item("Item_5").Specific, SAPbouiCOM.ComboBox)
            Dim UOMCombo As SAPbouiCOM.ComboBox = DirectCast(Me.objForm.Items.Item("Item_4").Specific, SAPbouiCOM.ComboBox)
            Dim BranchCombo As SAPbouiCOM.ComboBox = Me.objForm.Items.Item("BR").Specific
            Dim BatchCombo As SAPbouiCOM.ComboBox = Me.objForm.Items.Item("BT").Specific
            Dim PrinterCombo As SAPbouiCOM.ComboBox = Me.objForm.Items.Item("Printer").Specific

            Dim WithPriceCheck As SAPbouiCOM.CheckBox = Me.objForm.Items.Item("Item_6").Specific

            Dim PrintQty As Integer = objForm.Items.Item("txtPrintQ").Specific.Value
            Dim ItemCode As String = Me.objForm.Items.Item("txtItem").Specific.Value.ToString().Trim()
            ' Dim WithPrice As String = If(Not WithPriceCheck.Checked, "Y", "N")

            Dim WithPrice As String
            If WithPriceCheck.Checked = True Then
                WithPrice = "Y"
            Else
                WithPrice = "N"
            End If

            Dim Type As String
            Try
                Type = TypeCobmo.Selected.Value
            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText("Please select barcode print type", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

            Dim Batch As String = BatchCombo.Selected.Value

            Dim BranchID As String = BranchCombo.Selected.Value
            Dim Barcode As String = BarcodeCombo.Selected.Description
            Dim BranchName As String = BranchCombo.Selected.Description
            Dim UOM As String = UOMCombo.Selected.Value
            Dim Printer As String = PrinterCombo.Selected.Value


            Dim ItemName As String = Me.GetFieldValue("Call ""Barcode_ItemName""('" & ItemCode & "','" & Barcode & "')")
            Dim Price As Double = Me.GetFieldValue("Call ""Barcode_Price""('" & ItemCode & "','" & UOM & "','" & BranchID & "')")

            If Type = "" Then
                objMain.objApplication.StatusBar.SetText("Type is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return
            ElseIf (Operators.CompareString(Barcode.Trim(), "", False) = 0) Then
                objMain.objApplication.StatusBar.SetText("BarCode is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return
            ElseIf PrintQty > 0 Then
                'Me.Label_Printing(Type, Barcode, ItemName, PrintQty, Price, WithPrice, Printer, Batch, BranchName, ItemCode)
                'For P As Integer = 1 To PrintQty
                objMain.objUtilities.Label_Printing(Type, Barcode, ItemName, PrintQty, Price, WithPrice, Printer, Batch, BranchName, ItemCode, UOM)
                'Next
            Else
                objMain.objApplication.StatusBar.SetText("Quantity is greater than 0", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return
            End If

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub Label_Printing(ByVal Type As String, ByVal BarCode As String, ByVal ItemName As String, ByVal Quantity As Integer, ByVal UnitPrice As Double, ByVal Print_UnitPrice As String, ByVal PrinterName As String, ByVal Batch As String, ByVal BranchName As String, ByVal ItemCode As String)
        Try
            Dim PrinterPath As String = ""
            Dim PRNFile As String = ""
            If (Operators.CompareString(Type, "1", False) = 0) Then
                PrinterPath = Me.GetFieldValue("SELECT ""U_PrinterPath"" FROM ""@NT_BARPRSETUP"" WHERE ""U_Type""='1' and ""U_PrinterName""='" & PrinterName & "'")
                PRNFile = Me.Notepad_Read("Type1.txt")
            ElseIf (Operators.CompareString(Type, "2", False) = 0) Then
                PrinterPath = Me.GetFieldValue("SELECT ""U_PrinterPath"" FROM ""@NT_BARPRSETUP"" WHERE ""U_Type""='2' and ""U_PrinterName""='" & PrinterName & "'")
                PRNFile = Me.Notepad_Read("Type2.txt")
            ElseIf (Operators.CompareString(Type, "3", False) = 0) Then
                PrinterPath = Me.GetFieldValue("SELECT ""U_PrinterPath"" FROM ""@NT_BARPRSETUP"" WHERE ""U_Type""='3' and ""U_PrinterName""='" & PrinterName & "'")
                PRNFile = Me.Notepad_Read("Type3.txt")

            ElseIf (Operators.CompareString(Type, "4", False) = 0) Then
                PrinterPath = Me.GetFieldValue("SELECT ""U_PrinterPath"" FROM ""@NT_BARPRSETUP"" WHERE ""U_Type""='4' and ""U_PrinterName""='" & PrinterName & "'")
                PRNFile = Me.Notepad_Read("Type4.txt")
            End If
            If (Operators.CompareString(PrinterPath, "", False) = 0) Then
                Me.objForm.Freeze(False)
                objMain.objApplication.StatusBar.SetText("Printer Path is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return
            ElseIf (Operators.CompareString(PRNFile, "", False) <> 0) Then

                Dim StrTime As String = Format(DateTime.Now, "HHMMssfff")
                Dim StrDate As String = Format(DateTime.Now, "ddMMyy")
                '------------
                Dim FileName As String = "Print_" & objMain.objCompany.UserName & StrTime & ".prn"
                Dim FinalPRNFile As String = String.Concat(System.Windows.Forms.Application.StartupPath, "\", FileName)
                If (File.Exists(FinalPRNFile)) Then
                    File.WriteAllText(FinalPRNFile, "")
                Else
                    File.Create(FinalPRNFile).Close()
                End If
                '-----------------------------------------------
                'Dim ItemGrpQuery As String = " SELECT T1.""ItmsGrpCod""  FROM ""OITM"" T1 WHERE T1.""ItmsGrpCod"" IN (134,139) AND T1.""ItemCode"" = (SELECT T0.""ItemCode"" FROM OBCD T0 WHERE T0.""BcdCode""='" & BarCode & "')"
                'Dim oRsItemGrp As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oRsItemGrp.DoQuery(ItemGrpQuery)

                'If oRsItemGrp.RecordCount = 0 Then
                '    PRNFile = PRNFile.Replace("[VAT]", "(VAT Incl.)")
                'Else
                '    PRNFile = PRNFile.Replace("[VAT]", " ")
                'End If

                '-----------------------------------------------
                Dim BasePrice As Double = 0
                Try
                    Dim ItemQuery As String = " SELECT T0.""U_PIC"", T0.""U_BsUnit"" FROM OITM T0 Where T0.""ItemCode""='" & ItemCode & "'  and IFNULL(T0.""U_PIC"",0)<>0 and IFNULL(T0.""U_BsUnit"",0)<>0"
                    Dim oRsItem As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsItem.DoQuery(ItemQuery)
                    If oRsItem.RecordCount > 0 Then
                        Dim Pic As Double = oRsItem.Fields.Item("U_PIC").Value
                        Dim BaseUnit As Double = oRsItem.Fields.Item("U_BsUnit").Value
                        If Pic <> 0 And BaseUnit <> 0 Then
                            BasePrice = (Pic * BaseUnit) / UnitPrice
                        End If
                    End If
                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText("Base price calculation error:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End Try
                Dim ArabicName As String = Me.GetFieldValue("SELECT T0.""FrgnName"" FROM OITM T0 WHERE T0.""ItemCode""='" & ItemCode & "'")
                PRNFile = PRNFile.Replace("[ArabicName]", ArabicName)

                '------------------------------------------------
                Try
                    PRNFile = PRNFile.Replace("[Batch]", Batch)
                Catch ex As Exception
                End Try

                PRNFile = PRNFile.Replace("[BasePrice]", BasePrice)
                PRNFile = PRNFile.Replace("[ItemCode]", ItemCode)

                PRNFile = PRNFile.Replace("[BarCode]", BarCode)
                PRNFile = PRNFile.Replace("[ItemName]", ItemName.ToUpper())
                PRNFile = PRNFile.Replace("[Quantity]", Conversions.ToString(Quantity))
                If (Operators.CompareString(Print_UnitPrice.Trim(), "Y", False) <> 0) Then
                    PRNFile = PRNFile.Replace("[Currency]", "")
                    PRNFile = PRNFile.Replace("[UnitPrice]", "")
                Else
                    PRNFile = PRNFile.Replace("[Currency]", "AED")
                    PRNFile = PRNFile.Replace("[UnitPrice]", Strings.Format(UnitPrice, "0.00"))
                End If
                PRNFile = PRNFile.Replace("[Date]", StrDate)
                PRNFile = PRNFile.Replace("[Location]", BranchName)

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

    Public Function GetFieldValue(ByVal query As String) As Object
        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs.DoQuery(query)
        Dim str As String = oRs.Fields.Item(0).Value.ToString()
        Return str
    End Function

End Class
