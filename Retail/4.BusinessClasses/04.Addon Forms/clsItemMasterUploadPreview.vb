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
Public Class clsItemMasterUploadPreview

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
    Sub CreateForm(ByVal ApplicationIDITMUploadPreview As String)
        Try
            objMain.objUtilities.LoadForm("IMPUploadPreview.xml", "IMPPerview", ResourceType.Embeded)

            objForm = objMain.objApplication.Forms.GetForm("IMPPerview", objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.Freeze(True)

            If IsAppIDExists(ApplicationIDITMUploadPreview) Then

                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE

                CType(objForm.Items.Item("AppID").Specific, SAPbouiCOM.EditText).Value = ApplicationIDITMUploadPreview

                objForm.Items.Item("1").Click()

            Else

                objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE

                Dim oDBs_Head As SAPbouiCOM.DBDataSource
                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_ITMPEVIEW")

                oDBs_Head.SetValue("U_AppID", 0, ApplicationIDITMUploadPreview)

            End If

            Dim oDT As SAPbouiCOM.DataTable
            oDT = objForm.DataSources.DataTables.Item("DT_0")

            Dim query As String = "CALL ""SBO_GET_TNX_ITEM_MASTER_UPLOAD""('" & ApplicationIDITMUploadPreview & "')"
            oDT.ExecuteQuery(query)

            Dim oGrid As SAPbouiCOM.Grid
            oGrid = objForm.Items.Item("Grid").Specific
            oGrid.DataTable = oDT

            Try
                oGrid.Columns.Item("Line No").Width = 40
            Catch
            End Try


            For i As Integer = 0 To oGrid.Columns.Count - 1
                oGrid.Columns.Item(i).Editable = False
            Next

            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Freeze(False)

        Catch ex As Exception
            Try
                objForm.Freeze(False)
            Catch
            End Try

            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ITMUP" And pVal.BeforeAction = False Then
                ' Me.CreateForm()

            End If


        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx <> "IMPPerview" Then Exit Sub

            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
            End Select

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try

            objForm = objMain.objApplication.Forms.GetForm("IMPPerview", objMain.objApplication.Forms.ActiveForm.TypeCount)

            Select Case BusinessObjectInfo.EventType

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                    If BusinessObjectInfo.BeforeAction = False AndAlso BusinessObjectInfo.ActionSuccess = True Then


                        objForm.Freeze(True)

                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_ITMPEVIEW")



                        objForm.Freeze(False)

                    End If


            End Select

        Catch ex As Exception
            Try : objForm.Freeze(False) : Catch : End Try
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Function IsAppIDExists(ByVal AppID As String) As Boolean
        Try
            Dim rs As SAPbobsCOM.Recordset =
                objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim sql As String =
                "SELECT ""DocEntry"" FROM ""@TNX_ITMPEVIEW"" WHERE ""U_AppID"" = '" & AppID & "'"

            rs.DoQuery(sql)

            Return rs.RecordCount > 0

        Catch
            Return False
        End Try
    End Function

End Class
