
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
Public Class clsCentralizedPriceChangeRequest

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
            objMain.objUtilities.LoadForm("CentralizedPriceChangeRequest.xml", "CPRT", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("CPRT", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CP")
            ' oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_PL_C0")

            objForm.Freeze(True)
            '  objForm.DataBrowser.BrowseBy = "DocNum"


            ''----------------------------------
            objForm.EnableMenu("1293", True)
            objForm.EnableMenu("1282", True)
            '----------------------------------


            LoadBranchCombo()
            oDBs_Head.SetValue("U_DStatus", oDBs_Head.Offset, "Open")


            objForm.Freeze(False)


            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "CPCR" And pVal.BeforeAction = False Then
                Me.CreateForm()
                ' Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                ' Me.SetDefault(objForm.UniqueID)
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
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_CP")
                    ' oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_PL_C0")

                    oGrid = objForm.Items.Item("Grid").Specific


                    ''If pVal.ItemUID = "Load" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                    If pVal.ItemUID = "Load" And pVal.BeforeAction = True Then

                        'Validation for from date and to date mantoary and price Update type also

                        Dim PriceUpdateType As String = oDBs_Head.GetValue("U_PUType", 0).Trim()
                        Dim FDate As String = oDBs_Head.GetValue("U_FDate", 0).Trim()
                        Dim TDate As String = oDBs_Head.GetValue("U_TDate", 0).Trim()



                        If String.IsNullOrEmpty(FDate) Then
                            objMain.objApplication.StatusBar.SetText("Please select From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If

                        If String.IsNullOrEmpty(TDate) Then
                            objMain.objApplication.StatusBar.SetText("Please select To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If

                        If (FDate > TDate) Then
                            objMain.objApplication.StatusBar.SetText("From date should be less than To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If

                        If String.IsNullOrEmpty(PriceUpdateType) Then
                            objMain.objApplication.StatusBar.SetText("Please select Price Update Type....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        Me.Central_PRC(objForm.UniqueID)

                    End If


                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                    'If pVal.ItemUID = "Grid" And pVal.ColUID = "Price Change Request Number" And pVal.Before_Action = False Then
                    '    oGrid = objForm.Items.Item("Grid").Specific
                    '    Dim code As String = oGrid.DataTable.GetValue("Price Change Request Number", pVal.Row)
                    '    objMain.ObjPriceChangeRequest.CreateForm(code)
                    'End If

                    If pVal.ItemUID = "Grid" And pVal.ColUID = "Price Change Request Number" And pVal.Before_Action = False Then
                        oGrid = objForm.Items.Item("Grid").Specific

                        Dim DocumentNumber As String = oGrid.DataTable.GetValue("Price Change Request Number", pVal.Row)
                        'Dim DocumentNumber As String = oDBs_Head.GetValue("Price Change Request Number", 0).Trim()
                        objMain.objApplication.ActivateMenuItem("PCRT")
                        objNewForm = objMain.objApplication.Forms.GetForm("PCRT", objMain.objApplication.Forms.ActiveForm.TypeCount)
                        objNewForm.Freeze(True)
                        objNewForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objNewForm.Items.Item("DocNum").Specific.Value = DocumentNumber
                        objNewForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        objNewForm.Freeze(False)
                    End If

            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub LoadBranchCombo()
        Try
            Dim oCombo As SAPbouiCOM.ComboBox
            oCombo = CType(objForm.Items.Item("Branch").Specific, SAPbouiCOM.ComboBox)

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
    Sub Central_PRC(ByVal FormUID As String)

        objForm = objMain.objApplication.Forms.Item(FormUID)

        Dim FDate As String = oDBs_Head.GetValue("U_FDate", 0).Trim()
        Dim TDate As String = oDBs_Head.GetValue("U_TDate", 0).Trim()
        Dim PUType As String = oDBs_Head.GetValue("U_PUType", 0).Trim()
        Dim Branch As String = oDBs_Head.GetValue("U_Branch", 0).Trim()
        Dim DStatus As String = oDBs_Head.GetValue("U_DStatus", 0).Trim()

        Dim oGrid As SAPbouiCOM.Grid =
        CType(objForm.Items.Item("Grid").Specific, SAPbouiCOM.Grid)

        Dim oDataTable As SAPbouiCOM.DataTable


        If Not DataTableExists(objForm, "Grid") Then
            oDataTable = objForm.DataSources.DataTables.Add("Grid")
        Else
            oDataTable = objForm.DataSources.DataTables.Item("Grid")
            oDataTable.Clear()
        End If


        Dim str As String =
        "CALL ""TNX_LOAD_CENTRAL_PCR""('" & FDate & "','" & TDate & "','" & PUType & "','" & Branch & "','" & DStatus & "');"

        oDataTable.ExecuteQuery(str)

        oGrid.DataTable = oDataTable

        If oGrid.Columns.Item("Price Change Request Number") IsNot Nothing Then
            oGrid.Columns.Item("Price Change Request Number").LinkedObjectType = "TNXPCRUDO"
        End If

        ' Make grid read-only
        For i As Integer = 0 To oGrid.Columns.Count - 1
            oGrid.Columns.Item(i).Editable = False
        Next

        ' Optional: auto resize columns
        oGrid.AutoResizeColumns()

    End Sub


    Private Function DataTableExists(ByVal oForm As SAPbouiCOM.Form, ByVal DTID As String) As Boolean
        For i As Integer = 0 To oForm.DataSources.DataTables.Count - 1
            If oForm.DataSources.DataTables.Item(i).UniqueID = DTID Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Function ComboValueExists(ByVal oCombo As SAPbouiCOM.ComboBox,
                             ByVal value As String) As Boolean
        For i As Integer = 0 To oCombo.ValidValues.Count - 1
            If oCombo.ValidValues.Item(i).Value = value Then
                Return True
            End If
        Next
        Return False
    End Function
End Class
