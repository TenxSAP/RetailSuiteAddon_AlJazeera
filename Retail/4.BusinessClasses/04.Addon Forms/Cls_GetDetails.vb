
Imports System.Configuration
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Mail
Imports System.Runtime.CompilerServices.RuntimeHelpers
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports Microsoft.VisualBasic.ApplicationServices
Imports SAPbobsCOM
Imports SAPbouiCOM
Public Class cls_GetDetails

#Region "       Declaration             "
    Public objForm, objNewForm, objListForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details, oDS1 As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim rs As SAPbobsCOM.Recordset
    Dim objComboGrid As SAPbouiCOM.ComboBoxColumn

    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim BaseUom, Sucess, failed As String
    Public WithEvents Timer As New Timers.Timer
    Private WithEvents BackgroundWorker_SalaryJV As New System.ComponentModel.BackgroundWorker
    Private WithEvents BackgroundWorker_SalarySingleJVBranchwise As New System.ComponentModel.BackgroundWorker
    Dim DocNum, Series As String
    Dim oPrg As SAPbouiCOM.ProgressBar
    Public oMessageService As SAPbobsCOM.MessagesService
    Public oCmpSrv As SAPbobsCOM.CompanyService
    Private formUID As String = "PriceList" ' Set this globally when the form is opened
    Dim oGrid, oSubGrid As SAPbouiCOM.Grid
    Dim oDt As SAPbouiCOM.DataTable

    Private WithEvents BackgroundWorker_CSVLoad As New System.ComponentModel.BackgroundWorker

#End Region
    Sub CreateForm(itemcode)
        Try
            objMain.objUtilities.LoadForm("GetBarcode.xml", "GETDE", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("GETDE", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)


            Try

                Try
                    Dim oGrid As SAPbouiCOM.Grid = CType(objForm.Items.Item("Item_0").Specific, SAPbouiCOM.Grid)

                    Dim oDataTable As SAPbouiCOM.DataTable
                    If Not DataTableExists(objForm, "DT_0") Then
                        oDataTable = objForm.DataSources.DataTables.Add("DT_0")
                    Else
                        oDataTable = objForm.DataSources.DataTables.Item("DT_0")
                        oDataTable.Clear()
                    End If

                    Dim str As String = "CALL ""GetDetails"" ('" & itemcode & "')"
                    oDataTable.ExecuteQuery(str)
                    oGrid.DataTable = oDataTable
                    For i As Integer = 0 To oGrid.Columns.Count - 1
                        oGrid.Columns.Item(i).Editable = False
                    Next

                    If oGrid.Columns.Item("Promotion Document Number") IsNot Nothing Then
                        oGrid.Columns.Item("Promotion Document Number").LinkedObjectType = "ME_ITEMUDO"
                    End If

                    oGrid.Columns.Item("ArabicName").Editable = True
                    oGrid.Columns.Item("PackingName").Editable = True
                    oGrid.Columns.Item("PosScan").Editable = True


                    oGrid.Columns.Item("Sugar Tax").Editable = True
                    'oGrid.Columns.Item("Sugar Tax Amount").Editable = True
                    'oGrid.Columns.Item("Excise Tax").Editable = True
                    'oGrid.Columns.Item("Price").Editable = True

                    oGrid.Columns.Item("PosScan").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                    oGrid.Columns.Item("Sugar Tax").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox

                    Dim objComboGrid1 As SAPbouiCOM.ComboBoxColumn
                    objComboGrid1 = CType(oGrid.Columns.Item("Sugar Tax"), SAPbouiCOM.ComboBoxColumn)

                    objComboGrid1.ValidValues.Add("Yes", "Yes")
                    objComboGrid1.ValidValues.Add("No", "No")
                    Dim objComboGrid As SAPbouiCOM.ComboBoxColumn
                    objComboGrid = CType(oGrid.Columns.Item("PosScan"), SAPbouiCOM.ComboBoxColumn)

                    objComboGrid.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description

                    objComboGrid.ValidValues.Add("Yes", "Yes")
                    objComboGrid.ValidValues.Add("No", "No")

                    oGrid.LoadFromDataSource()
                    oGrid.AutoResizeColumns()
                Catch ex As Exception

                End Try


                objForm.Freeze(False)
            Catch ex As Exception

                If objForm IsNot Nothing Then
                    objForm.Freeze(False)
                End If
                objMain.objApplication.StatusBar.SetText(ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            End Try
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    ' If pVal.ItemUID = "1" And pVal.BeforeAction = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) And pVal.ActionSuccess = True Then

                    ' Me.SetDefault(objForm.UniqueID)
                    ' End If
                    If pVal.ItemUID = "1" AndAlso pVal.BeforeAction = False AndAlso pVal.ActionSuccess = True Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE OrElse
                           pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE OrElse pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            Dim oGrid As SAPbouiCOM.Grid = CType(objForm.Items.Item("Item_0").Specific, SAPbouiCOM.Grid)
                            Dim oDT As SAPbouiCOM.DataTable = oGrid.DataTable
                            For i As Integer = 0 To oDT.Rows.Count - 1
                                'Dim itemCode As String = oDT.GetValue("ItemCode", i).ToString()
                                Dim arabicName As String = oDT.GetValue("ArabicName", i).ToString()
                                Dim packingName As String = oDT.GetValue("PackingName", i).ToString()
                                Dim posScan As String = oDT.GetValue("PosScan", i).ToString()
                                Dim SugarTax As String = oDT.GetValue("Sugar Tax", i).ToString()
                                Dim SugarTaxAmount As String = oDT.GetValue("Base Amount", i).ToString()
                                Dim ExciseTax As String = oDT.GetValue("Excise Tax", i).ToString()
                                Dim catalogNo As String = oDT.GetValue("CatalogNo", i).ToString()
                                Dim Price As String = oDT.GetValue("Price", i).ToString()
                                Dim Barcode As String = oDT.GetValue("Barcode", i).ToString()
                                If posScan = "Yes" Then
                                    posScan = "Yes"
                                ElseIf posScan = "No" Then
                                    posScan = "No"
                                End If
                                If SugarTax = "Yes" Then
                                    SugarTax = "Yes"
                                ElseIf SugarTax = "No" Then
                                    SugarTax = "No"
                                End If


                                '          Dim StrQry As String = "UPDATE OBCD " &
                                ' "SET ""U_ArbName"" = '" & arabicName & "', " &
                                ' """U_PICK"" = '" & packingName & "', " &
                                ' """U_POSCA"" = '" & posScan & "' " &
                                '", ""U_STAX""='" & SugarTax & "' , ""U_SUTAXAMT""='" & SugarTaxAmount & "' ,""U_ETAX""='" & ExciseTax & "' ,""U_Price""='" & Price & "'  " &
                                ' "WHERE   ""BcdCode"" = '" & Barcode & "' "

                                Dim StrQry As String = "UPDATE OBCD " &
                       "SET ""U_ArbName"" = '" & arabicName & "', " &
                       """U_PICK"" = '" & packingName & "', " &
                       """U_POSCA"" = '" & posScan & "', " &
                       """U_STAX"" = '" & SugarTax & "' " &
                       "WHERE   ""BcdCode"" = '" & Barcode & "' "
                                ' Execute query using Recordset
                                Dim oRec As SAPbobsCOM.Recordset
                                oRec = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRec.DoQuery(StrQry)

                            Next
                            objMain.objApplication.StatusBar.SetText("Details have been successfully updated",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                        End If
                    End If

                    '-----Madhu Link button GetDetails to Promotion screen
                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                    If pVal.ItemUID = "Item_0" And pVal.ColUID = "Promotion Document Number" And pVal.Before_Action = False Then
                        oGrid = objForm.Items.Item("Item_0").Specific

                        Dim DocumentNumber As String = oGrid.DataTable.GetValue("Promotion Document Number", pVal.Row)
                        objMain.objApplication.ActivateMenuItem("ME_ITEMA")
                        objNewForm = objMain.objApplication.Forms.GetForm("ITEM", objMain.objApplication.Forms.ActiveForm.TypeCount)
                        objNewForm.Freeze(True)
                        objNewForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objNewForm.Items.Item("DocNum").Specific.Value = DocumentNumber
                        objNewForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        objNewForm.Freeze(False)
                    End If


            End Select

        Catch ex As Exception
            Try : objForm.Freeze(False) : Catch : End Try
            'objMain.objApplication.StatusBar.SetText("clsBarcodeRequest.ItemEvent: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Function DataTableExists(ByVal oForm As SAPbouiCOM.Form, ByVal DTID As String) As Boolean
        For i As Integer = 0 To oForm.DataSources.DataTables.Count - 1
            If oForm.DataSources.DataTables.Item(i).UniqueID = DTID Then
                Return True
            End If
        Next
        Return False
    End Function





End Class
