Imports System.Collections.Specialized
Imports System.Configuration
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices

Public Class clsClearanceSales

#Region "       Declaration             "
    Public objForm, objSTRForm As SAPbouiCOM.Form
    'Dim oDS, oDS1 As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Public oDS, oDS1 As SAPbouiCOM.DBDataSource
    Public PrgBar As SAPbouiCOM.ProgressBar

#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("ClearanceSales.xml", "Clear", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("Clear", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            oDS = Me.objForm.DataSources.DBDataSources.Add("@SBO_CLEAR")
            oDS1 = Me.objForm.DataSources.DBDataSources.Add("@SBO_CLEAR_C0")


            Dim oRsBranch As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim Query As String = "SELECT T2.""BPLId"",T2.""BPLName"" AS ""Branch"" FROM USR6 T0  INNER JOIN OUSR T1 ON T0.""UserCode"" = T1.""USER_CODE"" " & _
                    " INNER JOIN OBPL T2 ON T0.""BPLId"" = T2.""BPLId"" WHERE T1.""USER_CODE"" = '" & objMain.objCompany.UserName.ToString().Trim & "' "
            oRsBranch.DoQuery(Query)

            objComboBox = objForm.Items.Item("Branch").Specific
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, Query)
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("Branch").DisplayDesc = True
            Try
                If objComboBox.ValidValues.Count <> 0 Then
                    objComboBox.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
                End If
            Catch ex As Exception
            End Try
          

            Me.SetDefault(objForm.UniqueID)

            objComboBox = objForm.Items.Item("17").Specific
            objComboBox.ValidValues.Add("", "")
            objComboBox.ValidValues.Add("Open", "Open")
            objComboBox.ValidValues.Add("Close", "Close")
            objComboBox.ValidValues.Add("Cancel", "Cancel")


            Dim WhsCmbo As SAPbouiCOM.ComboBox = DirectCast(Me.objForm.Items.Item("Whs").Specific, SAPbouiCOM.ComboBox)
            Dim WhsQry As String = "SELECT T0.""WhsCode"", T0.""WhsName"" FROM OWHS T0 Where  ""U_Transit"" = 'No' "
            objMain.objUtilities.ComboBoxLoadValues(WhsCmbo, WhsQry)
            'WhsCmbo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            If WhsCmbo.ValidValues.Count <> 0 Then
                WhsCmbo.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
            End If

            Dim UOMCombo As SAPbouiCOM.ComboBox = DirectCast(Me.objMatrix.Columns.Item("UOM").Cells.Item(1).Specific, SAPbouiCOM.ComboBox)
            Dim ReceiptUOMCombo As SAPbouiCOM.ComboBox = DirectCast(Me.objMatrix.Columns.Item("RUOM").Cells.Item(1).Specific, SAPbouiCOM.ComboBox)
            Dim UOMQry As String = "SELECT DISTINCT A.""UomEntry"",(SELECT ""UomCode"" FROM OUOM WHERE ""UomEntry""=A.""UomEntry"") AS ""UomName"" FROM OBCD A "
            objMain.objUtilities.ComboBoxLoadValues(UOMCombo, UOMQry)
            objMain.objUtilities.ComboBoxLoadValues(ReceiptUOMCombo, UOMQry)

            Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            Me.objForm.Items.Item("17").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Me.objForm.Items.Item("17").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ME_CLEAR" And pVal.BeforeAction = False Then
                Me.CreateForm()
                'Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                'Me.SetDefault(objForm.UniqueID)
            End If
            Try
                If pVal.MenuUID = "Delete Row" And pVal.BeforeAction = False Then
                    objForm = objMain.objApplication.Forms.GetForm("Clear", objMain.objApplication.Forms.ActiveForm.TypeCount)
                    objMatrix = objForm.Items.Item("Mtx").Specific
                    For i As Integer = 1 To objMatrix.VisualRowCount - 1
                        If objMatrix.IsRowSelected(i) = True Then
                            objMatrix.DeleteRow(i)
                        End If
                    Next
                    For i As Integer = 1 To objMatrix.VisualRowCount
                        objMatrix.Columns.Item("LineId").Cells.Item(i).Specific.string = i
                    Next
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If

                If pVal.MenuUID = "519" And pVal.BeforeAction = True Then 'Preview
                    Try
                        Dim frs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'ClearanceSalesRpt' AND ""Type"" = 'C'")
                        If frs.RecordCount = 0 Then
                            objMain.objApplication.MessageBox("Report layout 'ClearanceSalesRpt' not found. Please upload with same name and Code as parameter", 0, "OK")
                        Else
                            objMain.objApplication.ActivateMenuItem(frs.Fields.Item(0).Value.ToString())
                            oDS = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR")
                            Dim CrForm As SAPbouiCOM.Form
                            Dim oedt As SAPbouiCOM.EditText
                            CrForm = objMain.objApplication.Forms.ActiveForm
                            oedt = CrForm.Items.Item("1000003").Specific
                            oedt.Value = oDS.GetValue("DocEntry", 0)
                            CrForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        End If
                    Catch ex As Exception
                    End Try
                End If
                If pVal.MenuUID = "520" And pVal.BeforeAction = True Then 'Print
                    Try
                        Dim frs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'ClearanceSalesRpt' AND ""Type"" = 'C'")
                        If frs.RecordCount = 0 Then
                            objMain.objApplication.MessageBox("Report layout 'ClearanceSalesRpt' not found. Please upload with same name and Code as parameter", 0, "OK")
                        Else
                            objMain.objApplication.ActivateMenuItem(frs.Fields.Item(0).Value.ToString())
                            oDS = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR")
                            Dim CrForm As SAPbouiCOM.Form
                            Dim oedt As SAPbouiCOM.EditText
                            CrForm = objMain.objApplication.Forms.ActiveForm
                            oedt = CrForm.Items.Item("1000003").Specific
                            oedt.Value = oDS.GetValue("DocEntry", 0)
                            CrForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        End If
                    Catch ex As Exception
                    End Try
                End If

                If pVal.MenuUID = "Close" And pVal.BeforeAction = False Then
                    objForm = objMain.objApplication.Forms.GetForm("Clear", objMain.objApplication.Forms.ActiveForm.TypeCount)
                    objMatrix = objForm.Items.Item("Mtx").Specific
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        objMain.objApplication.MessageBox("Form should be in 'OK Mode'.. ", 1, "Ok", "Cancel")
                        Exit Sub
                    End If
                    Dim result As Integer = objMain.objApplication.MessageBox("Changing document status to close is irreversible..Confirm... Do you want to continue?", 1, "Continue", "Cancel")
                    If (result = 1) Then
                        Me.UpdateStatus(objForm.UniqueID)
                    End If
                End If
            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub UpdateStatus(ByVal FormUID As String)
        Try
            objForm.Freeze(True)

             oDS = Me.objForm.DataSources.DBDataSources.Add("@SBO_CLEAR")
            oDS1 = Me.objForm.DataSources.DBDataSources.Add("@SBO_CLEAR_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific

            objMain.sCmp = objMain.objCompany.GetCompanyService
            objMain.oGeneralService = objMain.sCmp.GetGeneralService("SBO_CLEARUDO")
            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            objMain.oGeneralParams.SetProperty("DocEntry", oDS.GetValue("DocEntry", 0))
            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
            objMain.oGeneralData.SetProperty("U_MESTS", "Close")
            objMain.oGeneralService.Update(objMain.oGeneralData)
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@SBO_CLEAR"), "DocEntry", oDS.GetValue("DocEntry", 0))
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@SBO_CLEAR_C0"), "DocEntry", oDS1.GetValue("DocEntry", 0))

            objMatrix.LoadFromDataSource()

            objForm.Refresh()
            objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
            objMain.objApplication.StatusBar.SetText("DocStatus updated successfully...!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    If pVal.ItemUID = "Branch" And pVal.BeforeAction = False Then
                        Dim objBranchCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Branch").Specific
                        Dim BranchID As String
                        Try
                            BranchID = objBranchCombo.Selected.Value
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText("Please select Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                        Dim WhsCmbo As SAPbouiCOM.ComboBox = DirectCast(Me.objForm.Items.Item("Whs").Specific, SAPbouiCOM.ComboBox)
                        Dim WhsQry As String = "SELECT T0.""WhsCode"", T0.""WhsName"" FROM OWHS T0 Where ""BPLid""='" & BranchID & "' And ""U_Transit"" = 'No' "
                        objMain.objUtilities.ComboBoxLoadValues(WhsCmbo, WhsQry)
                        Try
                            If WhsCmbo.ValidValues.Count > 0 Then
                                WhsCmbo.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
                            End If
                        Catch ex As Exception
                        End Try                  
                    End If
                    If pVal.ItemUID = "Whs" And pVal.BeforeAction = False And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim objWhsCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Whs").Specific
                        objMatrix = objForm.Items.Item("Mtx").Specific
                        Dim Whs As String
                        Try
                            Whs = objWhsCombo.Selected.Value
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText("Please select warehouse", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                        For i As Integer = 1 To objMatrix.VisualRowCount
                            If objMatrix.Columns.Item("ICode").Cells.Item(i).Specific.Value.ToString() <> "" Then
                                oDS1.SetValue("LineID", oDS1.Offset, i)
                                oDS1.SetValue("U_ICode", oDS1.Offset, objMatrix.Columns.Item("ICode").Cells.Item(i).Specific.Value)
                                oDS1.SetValue("U_IName", oDS1.Offset, objMatrix.Columns.Item("IName").Cells.Item(i).Specific.Value)
                                oDS1.SetValue("U_BarCode", oDS1.Offset, objMatrix.Columns.Item("BarCode").Cells.Item(i).Specific.Value)
                                oDS1.SetValue("U_UOM", oDS1.Offset, objMatrix.Columns.Item("UOM").Cells.Item(i).Specific.Value)
                                oDS1.SetValue("U_Price", oDS1.Offset, Me.GetFieldValue("SELECT T0.""AvgPrice"" FROM OITW T0 WHERE T0.""ItemCode""='" & objMatrix.Columns.Item("ICode").Cells.Item(i).Specific.Value & "' and  T0.""WhsCode"" ='" & Whs & "'"))
                                oDS1.SetValue("U_Qty", oDS1.Offset, objMatrix.Columns.Item("Qty").Cells.Item(i).Specific.Value)

                                oDS1.SetValue("U_RBarCode", oDS1.Offset, objMatrix.Columns.Item("RBarCode").Cells.Item(i).Specific.Value)
                                oDS1.SetValue("U_RCode", oDS1.Offset, objMatrix.Columns.Item("RCode").Cells.Item(i).Specific.Value)
                                oDS1.SetValue("U_RName", oDS1.Offset, objMatrix.Columns.Item("RName").Cells.Item(i).Specific.Value)
                                oDS1.SetValue("U_RUOM", oDS1.Offset, objMatrix.Columns.Item("RUOM").Cells.Item(i).Specific.Value)
                                oDS1.SetValue("U_RQty", oDS1.Offset, objMatrix.Columns.Item("RQty").Cells.Item(i).Specific.Value)

                                oDS1.SetValue("U_Prod", oDS1.Offset, "")
                                oDS1.SetValue("U_RProd", oDS1.Offset, "")
                                objMatrix.SetLineData(i)

                            End If
                        Next
                    End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDS = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR")
                    oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR_C0")
                    objMatrix = objForm.Items.Item("Mtx").Specific
                    Dim objBranchCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Branch").Specific
                    Dim BranchID As String


                    If pVal.ItemUID = "Mtx" And pVal.ColUID = "BarCode" Then
                        If objMatrix.Columns.Item("BarCode").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            Try
                                BranchID = objBranchCombo.Selected.Value
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText("Please select Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Try
                            Dim GetItemCode As String = " Select ""ItemCode"" ,(Select A.""InvntryUom"" From OITM A Where  A.""ItemCode"" = B.""ItemCode"") AS ""PUoMEntry"", (Select A.""ItemName"" From OITM A Where  A.""ItemCode"" = B.""ItemCode"") ""ItemName"" " & _
                                                            " From OBCD B Where ""BcdCode"" = '" & objMatrix.Columns.Item("BarCode").Cells.Item(pVal.Row).Specific.Value & "'"
                            Dim oRsGetItemCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetItemCode.DoQuery(GetItemCode)

                            oDS1.SetValue("LineID", oDS1.Offset, pVal.Row)
                            oDS1.SetValue("U_ICode", oDS1.Offset, oRsGetItemCode.Fields.Item("ItemCode").Value)
                            oDS1.SetValue("U_IName", oDS1.Offset, oRsGetItemCode.Fields.Item("ItemName").Value)
                            oDS1.SetValue("U_BarCode", oDS1.Offset, objMatrix.Columns.Item("BarCode").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_UOM", oDS1.Offset, oRsGetItemCode.Fields.Item("PUoMEntry").Value)
                            oDS1.SetValue("U_Price", oDS1.Offset, Me.GetFieldValue("SELECT T0.""AvgPrice"" FROM OITW T0 WHERE T0.""ItemCode""='" & oRsGetItemCode.Fields.Item("ItemCode").Value & "' and  T0.""WhsCode"" ='" & oDS.GetValue("U_Whs", 0) & "'"))
                            oDS1.SetValue("U_Qty", oDS1.Offset, objMatrix.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_SSTK", oDS1.Offset, Me.GetFieldValue("SELECT T0.""OnHand"" FROM OITW T0 WHERE T0.""ItemCode"" ='" & oRsGetItemCode.Fields.Item("ItemCode").Value & "' and  T0.""WhsCode"" ='" & oDS.GetValue("U_Whs", 0) & "'"))

                            oDS1.SetValue("U_RBarCode", oDS1.Offset, objMatrix.Columns.Item("RBarCode").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_RCode", oDS1.Offset, objMatrix.Columns.Item("RCode").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_RName", oDS1.Offset, objMatrix.Columns.Item("RName").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_RUOM", oDS1.Offset, objMatrix.Columns.Item("RUOM").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_RQty", oDS1.Offset, objMatrix.Columns.Item("RQty").Cells.Item(pVal.Row).Specific.Value)

                            oDS1.SetValue("U_Prod", oDS1.Offset, "")
                            oDS1.SetValue("U_RProd", oDS1.Offset, "")
                            objMatrix.SetLineData(pVal.Row)

                            If pVal.Row = objMatrix.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID)
                            End If
                        End If
                    End If

                    If pVal.ItemUID = "Mtx" And pVal.ColUID = "RBarCode" Then
                        If objMatrix.Columns.Item("RBarCode").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            Try
                                BranchID = objBranchCombo.Selected.Value
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText("Please select Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Try
                            Dim GetItemCode As String = " Select ""ItemCode"" ,(Select A.""InvntryUom"" From OITM A Where  A.""ItemCode"" = B.""ItemCode"") AS ""PUoMEntry"", (Select A.""ItemName"" From OITM A Where  A.""ItemCode"" = B.""ItemCode"") ""ItemName"" " & _
                                                        " From OBCD B Where ""BcdCode"" = '" & objMatrix.Columns.Item("RBarCode").Cells.Item(pVal.Row).Specific.Value & "'"
                            Dim oRsGetItemCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetItemCode.DoQuery(GetItemCode)

                            oDS1.SetValue("LineID", oDS1.Offset, pVal.Row)
                            oDS1.SetValue("U_ICode", oDS1.Offset, objMatrix.Columns.Item("ICode").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_IName", oDS1.Offset, objMatrix.Columns.Item("IName").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_BarCode", oDS1.Offset, objMatrix.Columns.Item("BarCode").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_UOM", oDS1.Offset, objMatrix.Columns.Item("UOM").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_Price", oDS1.Offset, objMatrix.Columns.Item("Price").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_Qty", oDS1.Offset, objMatrix.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_SSTK", oDS1.Offset, Me.GetFieldValue("SELECT T0.""OnHand"" FROM OITW T0 WHERE T0.""ItemCode"" ='" & oRsGetItemCode.Fields.Item("ItemCode").Value & "' and  T0.""WhsCode"" ='" & oDS.GetValue("U_Whs", 0) & "'"))

                            oDS1.SetValue("U_RBarCode", oDS1.Offset, objMatrix.Columns.Item("RBarCode").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_RCode", oDS1.Offset, oRsGetItemCode.Fields.Item("ItemCode").Value)
                            oDS1.SetValue("U_RName", oDS1.Offset, oRsGetItemCode.Fields.Item("ItemName").Value)
                            oDS1.SetValue("U_RUOM", oDS1.Offset, oRsGetItemCode.Fields.Item("PUoMEntry").Value)
                            oDS1.SetValue("U_RQty", oDS1.Offset, objMatrix.Columns.Item("RQty").Cells.Item(pVal.Row).Specific.Value)

                            oDS1.SetValue("U_Prod", oDS1.Offset, "")
                            oDS1.SetValue("U_RProd", oDS1.Offset, "")
                            objMatrix.SetLineData(pVal.Row)

                            If pVal.Row = objMatrix.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID)
                            End If
                        End If
                    End If
                    If pVal.ItemUID = "Mtx" And pVal.ColUID = "ICode" Then
                        If objMatrix.VisualRowCount = pVal.Row Then
                            If objMatrix.Columns.Item("ICode").Cells.Item(pVal.Row).Specific.Value <> "" Then
                                Me.SetNewLine(objForm.UniqueID)
                            End If
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDS = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR")
                    oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR_C0")
                    objMatrix = objForm.Items.Item("Mtx").Specific
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    'If oCFL.UniqueID = "CFL_ITEM" Then
                    '    Me.CFLFilter(objForm.UniqueID, "CFL_ITEM")
                    'End If

                    Dim objBranchCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Branch").Specific
                    Dim BranchID As String
                    Try
                        BranchID = objBranchCombo.Selected.Value
                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText("Please select Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try

                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If oCFL.UniqueID = "CFL_ITEM" Then
                            oDS1.SetValue("LineID", oDS1.Offset, pVal.Row)
                            oDS1.SetValue("U_ICode", oDS1.Offset, oDT.GetValue("ItemCode", 0))
                            oDS1.SetValue("U_IName", oDS1.Offset, oDT.GetValue("ItemName", 0))
                            oDS1.SetValue("U_BarCode", oDS1.Offset, oDT.GetValue("CodeBars", 0))
                            oDS1.SetValue("U_UOM", oDS1.Offset, oDT.GetValue("InvntryUom", 0))
                            oDS1.SetValue("U_Price", oDS1.Offset, Me.GetFieldValue("SELECT T0.""AvgPrice"" FROM OITW T0 WHERE T0.""ItemCode""='" & oDT.GetValue("ItemCode", 0) & "' and  T0.""WhsCode"" ='" & oDS.GetValue("U_Whs", 0) & "'"))
                            oDS1.SetValue("U_Qty", oDS1.Offset, objMatrix.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_SSTK", oDS1.Offset, Me.GetFieldValue("SELECT T0.""OnHand"" FROM OITW T0 WHERE T0.""ItemCode"" ='" & oDT.GetValue("ItemCode", 0) & "' and  T0.""WhsCode"" ='" & oDS.GetValue("U_Whs", 0) & "'"))

                            oDS1.SetValue("U_RBarCode", oDS1.Offset, objMatrix.Columns.Item("RBarCode").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_RCode", oDS1.Offset, objMatrix.Columns.Item("RCode").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_RName", oDS1.Offset, objMatrix.Columns.Item("RName").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_RUOM", oDS1.Offset, objMatrix.Columns.Item("RUOM").Cells.Item(pVal.Row).Specific.Value)
                            'oDS1.SetValue("U_RPrice", oDS1.Offset, objMatrix.Columns.Item("RPrice").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_RQty", oDS1.Offset, objMatrix.Columns.Item("RQty").Cells.Item(pVal.Row).Specific.Value)

                            oDS1.SetValue("U_Prod", oDS1.Offset, "")
                            oDS1.SetValue("U_RProd", oDS1.Offset, "")

                            objMatrix.SetLineData(pVal.Row)
                            If pVal.Row = objMatrix.VisualRowCount Then
                                SetNewLine(objForm.UniqueID)
                            End If
                        End If
                        If oCFL.UniqueID = "CFL_ITM" Then
                            oDS1.SetValue("LineID", oDS1.Offset, pVal.Row)
                            oDS1.SetValue("U_ICode", oDS1.Offset, objMatrix.Columns.Item("ICode").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_IName", oDS1.Offset, objMatrix.Columns.Item("IName").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_BarCode", oDS1.Offset, objMatrix.Columns.Item("BarCode").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_UOM", oDS1.Offset, objMatrix.Columns.Item("UOM").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_Price", oDS1.Offset, objMatrix.Columns.Item("Price").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_Qty", oDS1.Offset, objMatrix.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_SSTK", oDS1.Offset, objMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific.Value)

                            oDS1.SetValue("U_RBarCode", oDS1.Offset, oDT.GetValue("CodeBars", 0))
                            oDS1.SetValue("U_RCode", oDS1.Offset, oDT.GetValue("ItemCode", 0))
                            oDS1.SetValue("U_RName", oDS1.Offset, oDT.GetValue("ItemName", 0))
                            oDS1.SetValue("U_RUOM", oDS1.Offset, oDT.GetValue("InvntryUom", 0))
                            'oDS1.SetValue("U_RPrice", oDS1.Offset, "")
                            oDS1.SetValue("U_RQty", oDS1.Offset, "")

                            oDS1.SetValue("U_Prod", oDS1.Offset, "")
                            oDS1.SetValue("U_RProd", oDS1.Offset, "")
                            objMatrix.SetLineData(pVal.Row)
                            If pVal.Row = objMatrix.VisualRowCount Then
                                SetNewLine(objForm.UniqueID)
                            End If
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "Post" And pVal.BeforeAction = False Then
                        Me.CreateProductionOrder(objForm.UniqueID)

                    End If
                    If pVal.ItemUID = "Receipt" And pVal.BeforeAction = False Then
                        Me.CreateReceipt(objForm.UniqueID)
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.Validation(objForm.UniqueID) = False Then BubbleEvent = False
                    End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("Clear", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        If objForm.Items.Item("17").Specific.Value.Trim = "Close" Then
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

    Public Function GetFieldValue(ByVal query As String) As Object
        Try
            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery(query)
            Dim str As String = oRs.Fields.Item(0).Value.ToString()
            Return str
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Sub SetDefault(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)


            oDS = Me.objForm.DataSources.DBDataSources.Add("@SBO_CLEAR")
            oDS1 = Me.objForm.DataSources.DBDataSources.Add("@SBO_CLEAR_C0")

            objMatrix = objForm.Items.Item("Mtx").Specific
            oDS.SetValue("DocNum", oDS.Offset, objMain.objUtilities.GetNextDocNum(objForm, "SBO_CLEARUDO", "Primary"))
            objForm.DataBrowser.BrowseBy = "DocNum"
            oDS.SetValue("U_MESTS", oDS.Offset, "Open")
            objMatrix.Clear()
            oDS1.Clear()
            objMatrix.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID)

            objMatrix.AutoResizeColumns()
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oDS = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR")
        oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR_C0")
        objMatrix = objForm.Items.Item("Mtx").Specific

        objMatrix.AddRow()
        oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
        oDS1.SetValue("U_ICode", oDS1.Offset, "")
        oDS1.SetValue("U_IName", oDS1.Offset, "")
        oDS1.SetValue("U_BarCode", oDS1.Offset, "")
        oDS1.SetValue("U_UOM", oDS1.Offset, "")
        oDS1.SetValue("U_Price", oDS1.Offset, "")
        oDS1.SetValue("U_Qty", oDS1.Offset, "")

        oDS1.SetValue("U_RBarCode", oDS1.Offset, "")
        oDS1.SetValue("U_RCode", oDS1.Offset, "")
        oDS1.SetValue("U_RName", oDS1.Offset, "")
        oDS1.SetValue("U_RUOM", oDS1.Offset, "")
        oDS1.SetValue("U_RPrice", oDS1.Offset, "")
        oDS1.SetValue("U_RQty", oDS1.Offset, "")

        oDS1.SetValue("U_Prod", oDS1.Offset, "")
        oDS1.SetValue("U_RProd", oDS1.Offset, "")
        objMatrix.SetLineData(objMatrix.VisualRowCount)
    End Sub

    Public Sub CreateProductionOrder(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDS = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR")
            oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific
            PrgBar = objMain.objApplication.StatusBar.CreateProgressBar("Posting production orders, Please wait....", objMatrix.RowCount, False)
            PrgBar.Value = 0
            For i As Integer = 1 To objMatrix.RowCount
                PrgBar.Value = i
                If objMatrix.Columns.Item("ICode").Cells.Item(i).Specific.Value.ToString().Trim() <> "" Then
                    'If objMatrix.Columns.Item("Prod").Cells.Item(i).Specific.Value.ToString().Trim() = "" Then
                    Dim WarehouseCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("Whs").Specific
                    Dim Whs As String = ""
                    Try
                        Whs = WarehouseCombo.Selected.Value
                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText("Please select warehouse", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                    objMain.objApplication.StatusBar.SetText("Production orders posting, Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Dim Prod As SAPbobsCOM.ProductionOrders = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)

                    Dim PostDate As String = objForm.Items.Item("PostDate").Specific.Value.ToString().Trim()
                    If PostDate = "" Then
                        objMain.objApplication.SetStatusBarMessage("Please enter ther the Posting Date", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        Exit Sub
                    End If
                    PostDate = PostDate.Insert(4, "/")
                    PostDate = PostDate.Insert(7, "/")
                    Dim PDate As Date = CDate(PostDate)
                    Prod.PostingDate = PDate
                    Prod.StartDate = PDate

                    Dim DueDate As String = objForm.Items.Item("PostDate").Specific.Value.ToString().Trim()
                    If DueDate = "" Then
                        objMain.objApplication.SetStatusBarMessage("Please enter ther the Due Date", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        Exit Sub
                    End If
                    DueDate = DueDate.Insert(4, "/")
                    DueDate = DueDate.Insert(7, "/")
                    Dim DDate As Date = CDate(DueDate)
                    Prod.DueDate = DDate

                    Prod.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotSpecial
                    Prod.ItemNo = objMatrix.Columns.Item("RCode").Cells.Item(i).Specific.Value.ToString().Trim()
                    Prod.PlannedQuantity = CDbl(objMatrix.Columns.Item("RQty").Cells.Item(i).Specific.Value.ToString().Trim())
                    Prod.Warehouse = Whs

                    Prod.Lines.ItemNo = objMatrix.Columns.Item("ICode").Cells.Item(i).Specific.Value.ToString().Trim()
                    Prod.Lines.PlannedQuantity = objMatrix.Columns.Item("Qty").Cells.Item(i).Specific.Value.ToString().Trim()
                    Prod.Lines.Warehouse = Whs
                    Prod.Lines.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                    Dim retVal As Integer = Prod.Add()
                    If retVal <> 0 Then
                        objMain.objApplication.StatusBar.SetText("Error:" & objMain.objCompany.GetLastErrorDescription.Replace("'", ""), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Dim oRsUpdate2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim sError As String = objMain.objCompany.GetLastErrorDescription.Replace("'", "").Replace("'", "")
                        Dim Update2 As String = " UPDATE  T1 SET T1.""U_Status""='Production Order Creation Error: " & sError & "' From ""@SBO_CLEAR_C0"" T1 WHERE T1.""DocEntry""='" & oDS.GetValue("DocEntry", 0) & "' and T1.""LineId""='" & i & "'"
                        oRsUpdate2.DoQuery(Update2)
                    Else
                        Dim oRsGetProd As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim Production As String = "SELECT max(T0.""DocEntry"") FROM OWOR T0"
                        oRsGetProd.DoQuery(Production)
                        Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim Update As String = "UPDATE T1 SET T1.""U_Prod""=(SELECT max(T0.""DocEntry"") FROM OWOR T0) , T1.""U_ProdNo""=(SELECT ""DocNum"" From OWOR  Where ""DocEntry"" = (SELECT max(T0.""DocEntry"") FROM OWOR T0)) ,    T1.""U_Status""='Production Order Posted' From ""@SBO_CLEAR_C0"" T1  WHERE T1.""DocEntry""='" & oDS.GetValue("DocEntry", 0) & "' and T1.""LineId""='" & i & "'"
                        oRsUpdate.DoQuery(Update)
                        '--------Releaseing----------------
                        Dim Release As SAPbobsCOM.ProductionOrders = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                        If Release.GetByKey(oRsGetProd.Fields.Item(0).Value) = True Then
                            Release.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposReleased
                            If Release.Update() = 0 Then
                                Dim oRsUpdate2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim Update2 As String = " UPDATE  T1 SET T1.""U_Status""='Production Order Released' From ""@SBO_CLEAR_C0"" T1 WHERE T1.""DocEntry""='" & oDS.GetValue("DocEntry", 0) & "' and T1.""LineId""='" & i & "'"
                                oRsUpdate2.DoQuery(Update2)
                                'Me.CreateReceipt(objForm.UniqueID, i, oDS.GetValue("U_Branch", 0))
                            Else
                                Dim oRsUpdate3 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim Update3 As String = " UPDATE  T1 SET T1.""U_Status""='Production Order Release Error:'" & objMain.objCompany.GetLastErrorDescription.Replace("'", "").ToString().Substring(0, 200) & " From ""@SBO_CLEAR_C0"" T1 WHERE T1.""DocEntry""='" & oDS.GetValue("DocEntry", 0) & "' and T1.""LineId""='" & i & "'"
                                oRsUpdate3.DoQuery(Update3)
                                objMain.objApplication.StatusBar.SetText("Error:" & objMain.objCompany.GetLastErrorDescription.Replace("'", ""), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                        End If
                    End If
                End If
                'End If
            Next
            PrgBar.Stop()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(PrgBar)
            objMain.objApplication.ActivateMenuItem("1304")
            objMatrix.AutoResizeColumns()
            objMain.objApplication.StatusBar.SetText("Production orders posted successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            objForm.Freeze(False)
            '-----------------------------------
        Catch ex As Exception
            PrgBar.Stop()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(PrgBar)

            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub CreateIssueForProduction(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDS = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR")
            oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific

            Dim result As Integer = objMain.objApplication.MessageBox("You are going to post Receipt from production. Do you want to continue?", 1, "Continue", "Cancel")
            If (result = 1) Then
                objMatrix = objForm.Items.Item("Mtx").Specific
                PrgBar = objMain.objApplication.StatusBar.CreateProgressBar("Posting production orders, Please wait....", objMatrix.RowCount, False)
                PrgBar.Value = 0
                For i As Integer = 1 To objMatrix.VisualRowCount
                    PrgBar.Value = i
                    If objMatrix.Columns.Item("ICode").Cells.Item(i).Specific.Value.ToString().Trim() <> "" And objMatrix.Columns.Item("Prod").Cells.Item(i).Specific.Value.ToString().Trim() <> "" And objMatrix.Columns.Item("RProd").Cells.Item(i).Specific.Value.ToString().Trim() = "" Then
                        objMain.objApplication.StatusBar.SetText("Please Wait Issue is posting.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        '----------
                        Dim ProdDetails As String = "Select ""LineNum"" from WOR1 A Where A.""DocEntry""='" & objMatrix.Columns.Item("Prod").Cells.Item(i).Specific.Value.ToString().Trim() & "'"
                        Dim oRsProdDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsProdDetails.DoQuery(ProdDetails)

                        Dim Issue As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
                        Issue.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
                        Issue.JournalMemo = "Issue from Production"
                        objComboBox = objForm.Items.Item("23").Specific
                        Issue.BPL_IDAssignedToInvoice = objComboBox.Value
                        Dim LineCount As Integer = 0
                        For j As Integer = 1 To oRsProdDetails.RecordCount
                            If LineCount > 0 Then
                                Issue.Lines.Add()
                            End If
                            Issue.Lines.BaseEntry = objMatrix.Columns.Item("Prod").Cells.Item(i).Specific.Value.ToString().Trim()
                            Issue.Lines.BaseLine = oRsProdDetails.Fields.Item(0).Value
                            Issue.Lines.BaseType = 202
                            LineCount = 1
                            oRsProdDetails.MoveNext()
                        Next
                        Dim retVal As Integer = Issue.Add()
                        '-------------
                        If retVal <> 0 Then
                            Dim oRsUpdate3 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim sError As String = objMain.objCompany.GetLastErrorDescription.Replace("'", "")
                            Dim Update3 As String = " UPDATE  T1 SET T1.""U_Status""='Issue From Production Error: " & sError & "' From ""@SBO_CLEAR_C0"" T1  WHERE T1.""DocEntry""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "' and T1.""LineId""='" & i & "'"
                            oRsUpdate3.DoQuery(Update3)
                            objMain.objApplication.StatusBar.SetText("Error:" & objMain.objCompany.GetLastErrorDescription.Replace("'", ""), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim Update As String = " UPDATE T1 SET T1.""U_IProd""=(SELECT max(T0.""DocEntry"") FROM OIGE T0),T1.""U_IProdNo""=(SELECT ""DocNum"" FROM OIGE Where ""DocEntry"" = (SELECT max(T0.""DocEntry"") FROM OIGE T0)) ,     T1.""U_Status""='Issue From Production Completed' From ""@SBO_CLEAR_C0"" T1 WHERE T1.""DocEntry""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "' and T1.""LineId""='" & i & "'"
                            oRsUpdate.DoQuery(Update)
                        End If
                    End If
                Next
                objMain.objApplication.StatusBar.SetText("Issue From Production posted successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                objMain.objApplication.ActivateMenuItem("1304")
                objMatrix.AutoResizeColumns()
            End If
            'Else
            'objMain.objApplication.StatusBar.SetText("Please check component quantities.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'End If
        Catch ex As Exception
            PrgBar.Stop()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(PrgBar)

            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub CreateReceipt(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDS = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR")
            oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific

            Dim GetSeries As String = " Select ""Series"" , ""SeriesName""   From NNM1 Where ""ObjectCode"" = '59' And ""Locked"" = 'N' And ""BPLId"" =  '" & oDS.GetValue("U_Branch", 0) & "'  Order By ""Series"" Desc "
            Dim oRsGetSeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetSeries.DoQuery(GetSeries)
            Dim DocSeries As String = oRsGetSeries.Fields.Item("Series").Value

            PrgBar = objMain.objApplication.StatusBar.CreateProgressBar("Please Wait Receipt is posting, Please wait....", objMatrix.RowCount, False)
            PrgBar.Value = 0
            For i As Integer = 1 To objMatrix.VisualRowCount
                PrgBar.Value = i
                If objMatrix.Columns.Item("ICode").Cells.Item(i).Specific.Value.ToString().Trim() <> "" And objMatrix.Columns.Item("Prod").Cells.Item(i).Specific.Value.ToString().Trim() <> "" And objMatrix.Columns.Item("RProd").Cells.Item(i).Specific.Value.ToString().Trim() = "" Then
                    Dim Receipt As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
                    Receipt.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
                    Receipt.JournalMemo = "Receipt from Production"
                    Receipt.Series = DocSeries
                    Receipt.BPL_IDAssignedToInvoice = oDS.GetValue("U_Branch", 0)
                    Receipt.Lines.BaseEntry = objMatrix.Columns.Item("Prod").Cells.Item(i).Specific.Value.ToString().Trim()
                    Receipt.Lines.BaseType = 202

                    Receipt.Lines.Add()
                    Dim retVal As Integer = Receipt.Add()
                    '-------------
                    If retVal <> 0 Then
                        Dim oRsUpdate3 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        Dim Update3 As String = " UPDATE  T1 SET T1.""U_Status""='Receipt From Production Error: " & objMain.objCompany.GetLastErrorDescription.Replace("'", "") & "' From ""@SBO_CLEAR_C0"" T1  WHERE T1.""DocEntry""='" & oDS.GetValue("DocEntry", 0) & "' and T1.""LineId""='" & i & "'"
                        oRsUpdate3.DoQuery(Update3)
                        objMain.objApplication.StatusBar.SetText("Error:" & objMain.objCompany.GetLastErrorDescription.Replace("'", ""), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Else
                        Dim oRsUpdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim Update As String = " UPDATE T1 SET T1.""U_RProd""=(SELECT max(T0.""DocEntry"") FROM OIGN T0), T1.""U_RProdNo""=(SELECT ""DocNum""  FROM OIGN Where ""DocEntry"" = (SELECT max(T0.""DocEntry"") FROM OIGN T0)) , T1.""U_Status""='Receipt From Production Completed' From ""@SBO_CLEAR_C0"" T1 WHERE T1.""DocEntry""='" & oDS.GetValue("DocEntry", 0) & "' and T1.""LineId""='" & i & "'"
                        oRsUpdate.DoQuery(Update)
                    End If
                End If
            Next

            PrgBar.Stop()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(PrgBar)
            objMain.objApplication.StatusBar.SetText("Receipt From Production posted successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objMain.objApplication.ActivateMenuItem("1304")
            objMatrix.AutoResizeColumns()
        Catch ex As Exception
            PrgBar.Stop()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(PrgBar)

            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region " Right Click Event"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Dim objForm As SAPbouiCOM.Form
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)
        Try
            If eventInfo.FormUID = objForm.UniqueID Then
                If (eventInfo.BeforeAction = True) Then
                    If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                        objMatrix = objForm.Items.Item("Mtx").Specific
                        If eventInfo.ItemUID <> "Mtx" Then
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("Close") = False Then
                                    oCreationPackage.UniqueID = "Close"
                                    oCreationPackage.String = "Close"
                                    oCreationPackage.Enabled = True
                                    oMenus.AddEx(oCreationPackage)
                                End If
                            Catch ex As Exception
                            End Try
                        End If
                        If eventInfo.ItemUID = "Mtx" And eventInfo.ColUID = "LineId" Then
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
                        ElseIf eventInfo.ItemUID = "Mtx" And objMatrix.RowCount <= 1 Then
                            oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                            oMenus = oMenuItem.SubMenus
                            Try
                                If oMenus.Exists("Delete Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Delete Row")
                                End If
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If
                        If eventInfo.ItemUID <> "Mtx" Then
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
            ' System.Diagnostics.Process.Start()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDS = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR")
            oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_CLEAR_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific
            If objForm.Items.Item("PostDate").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Post Date Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

           
            If objMatrix.VisualRowCount = 0 Then
                objMain.objApplication.StatusBar.SetText("Matrix cannot be Empty")
                Return False
            ElseIf objMatrix.VisualRowCount = 1 Then
                If (objMatrix.Columns.Item("ICode").Cells.Item(1).Specific.Value.Trim = "") Then
                    objMain.objApplication.StatusBar.SetText("Matrix cannot be Empty")
                    Return False
                End If
            End If

            For i As Integer = 1 To objMatrix.VisualRowCount
                Dim BarCode As String = objMatrix.Columns.Item("BarCode").Cells.Item(i).Specific.Value.Trim
                Dim ICode As String = objMatrix.Columns.Item("ICode").Cells.Item(i).Specific.Value.Trim
                Dim IssueQty As Double = objMatrix.Columns.Item("Qty").Cells.Item(i).Specific.Value.Trim
                If (BarCode <> "" And ICode = "") Then
                    objMain.objApplication.StatusBar.SetText("Issue Itemcode should not be empty for Line No. " & i & ".... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                If (ICode <> "" And IssueQty = 0) Then
                    objMain.objApplication.StatusBar.SetText("For Item " & ICode & " and Line No. " & i & ".... IssueQty should not be zero!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Next

            For i As Integer = 1 To objMatrix.VisualRowCount
                Dim ICode As String = objMatrix.Columns.Item("ICode").Cells.Item(i).Specific.Value.Trim
                Dim RCode As String = objMatrix.Columns.Item("RCode").Cells.Item(i).Specific.Value.Trim
                Dim RQty As Double = objMatrix.Columns.Item("RQty").Cells.Item(i).Specific.Value.Trim
                If (ICode <> "" And RCode = "") Then
                    objMain.objApplication.StatusBar.SetText("Receipt Itemcode should not be empty for Line No. " & i & ".... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                If (RCode <> "" And RQty = 0) Then
                    objMain.objApplication.StatusBar.SetText("For Item " & RCode & " and Line No. " & i & ".... Receipt Qty should not be zero!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Next

            If objForm.Items.Item("19").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Remarks Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If


            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

End Class
