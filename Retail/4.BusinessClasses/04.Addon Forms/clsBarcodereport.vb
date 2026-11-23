
Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports SAPbobsCOM
Imports System.Configuration
Imports System.Runtime.CompilerServices.RuntimeHelpers
Imports System.Data.OleDb
Imports Microsoft.VisualBasic.ApplicationServices
Imports System.Data.SqlClient

Imports System.Net
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Net.Http
Imports System.Net.Mail
Public Class clsBarcodereport

#Region "       Declaration             "
    Public objForm, objNewForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details, oDS1 As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim rs As SAPbobsCOM.Recordset
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
    Dim oGrid As SAPbouiCOM.Grid
    Dim oDt As SAPbouiCOM.DataTable

    Private WithEvents BackgroundWorker_CSVLoad As New System.ComponentModel.BackgroundWorker

#End Region
    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("ItemMasterBarcode.xml", "BCode", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("BCode", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            objForm.Freeze(True)
            LoadBranchCombo()
            objForm.Freeze(False)
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "BCReport" And pVal.BeforeAction = False Then
                Me.CreateForm()

            End If


        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub



    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx <> "BCode" Then Exit Sub

            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If pVal.BeforeAction = False Then

                        objForm = objMain.objApplication.Forms.Item(FormUID)
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_BCRRTIM")

                        If pVal.ItemUID = "Load" Then
                            LoadGrid(FormUID)
                        End If

                    End If

            End Select

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Private Sub LoadBranchCombo()
        Try
            Dim oCombo As SAPbouiCOM.ComboBox
            oCombo = CType(objForm.Items.Item("BRNCH").Specific, SAPbouiCOM.ComboBox)

            'oCombo.ValidValues.Clear()

            Dim oRs As SAPbobsCOM.Recordset =
           objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            ' oRs.DoQuery("SELECT ""BPLId"", ""BPLName"" FROM ""OBPL"" ORDER BY ""BPLId""")
            oRs.DoQuery("SELECT T2.""BPLId"",T2.""BPLName"" FROM USR6 T0  INNER JOIN OUSR T1 ON T0.""UserCode"" = T1.""USER_CODE"" " &
                    " INNER JOIN OBPL T2 ON T0.""BPLId"" = T2.""BPLId"" WHERE T1.""USER_CODE"" = '" & objMain.objCompany.UserName.ToString().Trim & "' ")



            While Not oRs.EoF
                Dim sBranchId As String = oRs.Fields.Item("BPLId").Value.ToString()
                Dim sBranchName As String = oRs.Fields.Item("BPLName").Value.ToString()

                If Not ComboValueExists(oCombo, sBranchId) Then
                    oCombo.ValidValues.Add(sBranchId, sBranchName)
                End If

                oRs.MoveNext()
            End While

            If oCombo.ValidValues.Count > 0 Then
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("LoadBranchCombo ERROR: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Function ComboValueExists(ByVal oCombo As SAPbouiCOM.ComboBox,
                             ByVal value As String) As Boolean
        For i As Integer = 0 To oCombo.ValidValues.Count - 1
            If oCombo.ValidValues.Item(i).Value = value Then
                Return True
            End If
        Next
        Return False
    End Function

    Sub LoadGrid(ByVal FormUID As String)

        Dim oGrid As SAPbouiCOM.Grid = Nothing
        Dim oDataTable As SAPbouiCOM.DataTable = Nothing

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            'Get Grid
            oGrid = CType(objForm.Items.Item("Grid").Specific, SAPbouiCOM.Grid)

            'Create or Get DataTable
            If Not DataTableExists(objForm, "DT_0") Then
                oDataTable = objForm.DataSources.DataTables.Add("DT_0")
            Else
                oDataTable = objForm.DataSources.DataTables.Item("DT_0")
                oDataTable.Clear()
            End If

            'Query
            Dim str As String = "SELECT * FROM ""OBCD"""

            'Execute Query
            oDataTable.ExecuteQuery(str)

            'Bind Grid
            oGrid.DataTable = oDataTable

            'Make Grid Read Only
            For i As Integer = 0 To oGrid.Columns.Count - 1
                oGrid.Columns.Item(i).Editable = False
            Next

            'Auto Resize
            oGrid.AutoResizeColumns()

            objForm.Freeze(False)

        Catch ex As Exception

            If objForm IsNot Nothing Then
                objForm.Freeze(False)
            End If

            objMain.objApplication.StatusBar.SetText(ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

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
