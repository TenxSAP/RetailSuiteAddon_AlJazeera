

'New
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
Public Class PriceUpdateService



#Region "       Declaration             "
    Public objForm As SAPbouiCOM.Form
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

    Private WithEvents BackgroundWorker_CSVLoad As New System.ComponentModel.BackgroundWorker

#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("PriceUpdate_1.xml", "PriceList", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("PriceList", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PL")
            oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_PL_C0")

            objForm.Freeze(True)
            objForm.DataBrowser.BrowseBy = "DocNum"
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("DocDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("DocSts").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocSts").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("Load").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("Load").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("Load").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("PU").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("PU").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("Submit").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("Submit").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            '----------------------------------
            objForm.EnableMenu("1293", True)
            objForm.EnableMenu("1282", True)
            '----------------------------------
            'Dim Query As String = "Select ""USER_CODE"" From OUSR Where ""USERID"" in (SELECT T1.""UserID""  FROM OWTM T0  INNER JOIN WTM1 T1 ON T0.""WtmCode"" = T1.""WtmCode"" where ""U_Price""='Yes') and ""USER_CODE""='" & objMain.objCompany.UserName.ToString().Trim & "'"
            'rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'rs.DoQuery(Query)
            'If rs.RecordCount > 0 Then
            '    objForm.Items.Item("PU").Enabled = True
            'Else
            '    objForm.Items.Item("PU").Enabled = False
            'End If
            '  objForm.Items.Item("1").Enabled = False

            '----------------------------------
            Try
                AddHandler Timer.Elapsed, AddressOf Refresh1
                Timer.Interval = 3000
                Timer.Enabled = True
            Catch ex As Exception
                Timer.Stop()

            End Try

            objForm.Items.Item("PU").Enabled = False
            ' objForm.Items.Item("Submit").Enabled = False
            objForm.Freeze(False)


            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub Refresh1()
        Try


            Dim exists As Boolean = False

            For i As Integer = 0 To objMain.objApplication.Forms.Count - 1
                Dim f = objMain.objApplication.Forms.Item(i)
                If f.UniqueID = formUID Or f.TypeEx = "PriceList" And f.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    exists = True
                    Exit For
                End If
            Next

            If Not exists Then
                Timer.Stop()
                RemoveHandler Timer.Elapsed, AddressOf Refresh1
                Return
            End If

            objForm = objMain.objApplication.Forms.Item(formUID)

            If objForm IsNot Nothing AndAlso objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE AndAlso objForm.TypeEx = "PriceList" Then

                Dim str As String = objForm.Items.Item("DocNum").Specific.value
                Dim QUERY As String = "SELECT * FROM ""@SBO_PL"" T0 INNER JOIN ""@SBO_PL_C0"" T1 ON T0.""DocEntry""=T1.""DocEntry"" WHERE ""DocNum""='" & str & "' AND IFNULL(""U_Remarks"",'')<>'Price Updated Successfully' AND T0.""U_DocSts""<>'Closed'"

                Dim rsHeader, GETDOCUMNENT, GETDOCUMNENT1 As SAPbobsCOM.Recordset
                rsHeader = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rsHeader.DoQuery(QUERY)

                If rsHeader.RecordCount <> 0 Then
                    Try
                        Dim GETDOCENTRY As String = "SELECT ""U_DOCNUM"" FROM ""@SBO_UPDATE"""
                        GETDOCUMNENT = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        GETDOCUMNENT1 = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        GETDOCUMNENT.DoQuery(GETDOCENTRY)
                        If GETDOCUMNENT.RecordCount > 0 Then
                            If GETDOCUMNENT.Fields.Item("U_DOCNUM").Value = str Then
                                objMain.objApplication.ActivateMenuItem("1304")
                            End If
                        End If
                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText("Refresh1 Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                End If
            End If
        Catch ex As Exception
            Timer.Stop()
            RemoveHandler Timer.Elapsed, AddressOf Refresh1
            objMain.objApplication.StatusBar.SetText("Form not accessible. Timer stopped.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Public Sub Refresh10()

        objForm = objMain.objApplication.Forms.Item("PriceList") ' if you have the UID

        ' If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE And objForm.Items.Item("DocSts").Specific.Value.ToString().Trim() = "Submitted" Then
        If objForm IsNot Nothing And objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE And objForm.TypeEx = "PriceList" Then



            'Dim PriceList5 As String = oDS.GetValue("U_PriceUp", 0).ToString().Trim
            'If PriceList5 = "1" Then
            Dim str As String = objForm.Items.Item("DocNum").Specific.value
            Dim QUERY As String = "SELECT * FROM ""@SBO_PL"" T0 inner join ""@SBO_PL_C0"" T1 ON T0.""DocEntry""=T1.""DocEntry"" WHERE ""DocNum""='" & str & "' and ifnull(""U_Remarks"",'')<>'Price Updated Successfully' and T0.""U_DocSts""<>'Closed' "
            Dim rsHeader, GETDOCUMNENT, GETDOCUMNENT1 As SAPbobsCOM.Recordset
            rsHeader = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsHeader.DoQuery(QUERY)
            If rsHeader.RecordCount <> 0 Then
                Try
                    Dim GETDOCENTRY As String = " SELECT ""U_DOCNUM"" FROM ""@SBO_UPDATE"""
                    GETDOCUMNENT = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    GETDOCUMNENT1 = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    GETDOCUMNENT.DoQuery(GETDOCENTRY)
                    If GETDOCUMNENT.RecordCount > 0 Then

                        ' Dim x As String = GETDOCUMNENT.Fields.Item("U_STATUS").Value.ToString()

                        ' If GETDOCUMNENT.Fields.Item("U_DOCNUM").Value = str And x <> "Completed" And objForm.TypeEx = "E_CMIP" Then
                        If GETDOCUMNENT.Fields.Item("U_DOCNUM").Value = str And objForm.TypeEx = "PriceList" Then

                            objMain.objApplication.ActivateMenuItem("1304")

                            'Else
                            '    ' Dim x As String = GETDOCUMNENT.Fields.Item("U_STATUS").Value.ToString()
                            '    If x = "Completed" Then
                            '            objMain.objApplication.StatusBar.SetText("Document Posting Completed SuccessFully..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            '            Dim GETDOCENTRY1 As String = " UPDATE ""@ME_MESSAGE"" SET ""U_DOCNUM""='' ,""U_STATUS""='',""U_DOCTYP""='' ,""U_USER""=''  "
                            '            GETDOCUMNENT1.DoQuery(GETDOCENTRY1)
                            '            objMain.objApplication.ActivateMenuItem("1304")
                            '        End If
                        End If
                    End If
                Catch ex As Exception
                End Try
            End If

        End If

    End Sub


    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "PriceList" And pVal.BeforeAction = False Then
                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
            End If
            If pVal.MenuUID = "Delete Row" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("ME_PSC_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix = objForm.Items.Item("29").Specific
                For i As Integer = 1 To objMatrix.VisualRowCount
                    If objMatrix.IsRowSelected(i) = True Then
                        objMatrix.DeleteRow(i)
                    End If
                Next
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.Columns.Item("LineId").Cells.Item(i).Specific.string = i
                Next
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
            If pVal.MenuUID = "Close" And pVal.BeforeAction = False Then
            End If

            If pVal.MenuUID = "1292" And pVal.BeforeAction = False Then

                'For i As Integer = 1 To objMatrix.VisualRowCount
                '    objMatrix.Columns.Item("V_-1").Cells.Item(i).Specific.string = i
                'Next
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    'objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "", Optional ByVal Series As Integer = 0)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PL")
            oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_PL_C0")

            objMatrix = objForm.Items.Item("Mtx").Specific

            Dim GetDflt As String = "Select ""DfltSeries"" From ONNM Where ""ObjectCode"" = 'SBO_PLUDO'"
            GetDflt = "SELECT MAX(T0.""Series"") FROM NNM1 T0 WHERE T0.""ObjectCode"" ='SBO_PLUDO' 
and  T0.""Indicator""=(SELECT T1.""Indicator"" FROM ""OFPR""  T1 WHERE  current_date between T1.""F_RefDate"" and  T1.""T_RefDate"")
 "


            Dim oRsGetDflt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDflt.DoQuery(GetDflt)


            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNumSeries(objForm, "SBO_PLUDO", oRsGetDflt.Fields.Item(0).Value))
            oDBs_Head.SetValue("U_DocSts", oDBs_Head.Offset, "Open")
            oDBs_Head.SetValue("U_DocDate", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))

            objMatrix.AutoResizeColumns()

            objMatrix.Clear()
            oDS1.Clear()
            objMatrix.FlushToDataSource()

            Me.SetNewLine()

            objMatrix.AutoResizeColumns()
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine()
        objMatrix = objForm.Items.Item("Mtx").Specific
        objMatrix.AddRow()
        oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
        oDS1.SetValue("U_ICode", oDS1.Offset, "")
        oDS1.SetValue("U_IName", oDS1.Offset, "")
        oDS1.SetValue("U_BarCode", oDS1.Offset, "")
        oDS1.SetValue("U_UOM", oDS1.Offset, "")
        oDS1.SetValue("U_GPrice", oDS1.Offset, "")
        oDS1.SetValue("U_OPrice", oDS1.Offset, "")
        oDS1.SetValue("U_NPrice", oDS1.Offset, "")
        'oDS1.SetValue("U_GP", oDS1.Offset, "")
        oDS1.SetValue("U_Remarks", oDS1.Offset, "")
        objMatrix.SetLineData(objMatrix.VisualRowCount)

    End Sub
    Function Validation(ByVal FormUID As String)
        Try
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PL")
            oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_PL_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific
            If objMatrix.VisualRowCount = 0 Then
                objMain.objApplication.StatusBar.SetText("Matrix cannot be Empty")
                Return False
            ElseIf objMatrix.VisualRowCount = 1 Then
                If (objMatrix.Columns.Item("BarCode").Cells.Item(1).Specific.Value.Trim = "") Then
                    objMain.objApplication.StatusBar.SetText("Matrix cannot be Empty")
                    Return False
                ElseIf (objMatrix.Columns.Item("LineId").Cells.Item(1).Specific.Value.Trim = "") Then
                    objMain.objApplication.StatusBar.SetText("Matrix cannot be Empty")
                    Return False
                End If
            End If


            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function



    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PL")
                    oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_PL_C0")

                    objMatrix = objForm.Items.Item("Mtx").Specific

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If Validation(objForm.UniqueID) = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim LoadSameForm As String = "select ""DocNum""  from ""@SBO_PL""  where ""DocEntry""= (Select MAX(""DocEntry"") FROM ""@SBO_PL"")"
                        Dim oRsLoadSameForm As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsLoadSameForm.DoQuery(LoadSameForm)
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objForm.Items.Item("DocNum").Specific.Value = oRsLoadSameForm.Fields.Item(0).Value
                        objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                    If pVal.ItemUID = "3" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or
                        pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                        Dim str As String = objForm.Items.Item("upload").Specific.value
                        If str <> "" Then
                            Dim FilePathCSV As String = str
                            '   Me.ReadCSV(FilePathCSV)
                            ' Me.Read13(FilePathCSV)
                            Me.StartBackgroundForLoading()
                        End If
                    End If
                    If pVal.ItemUID = "PU" And pVal.BeforeAction = False And objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Dim result As Integer = objMain.objApplication.MessageBox("You are going to update item prices. Do you want to continue?", 1, "Continue", "Cancel")
                        Dim oRsPriceList As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim oRsSubmit As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        If (result = 1) Then

                            objMatrix = objForm.Items.Item("Mtx").Specific
                            If objMatrix.VisualRowCount > 1 Then

                                objForm.Items.Item("PU").Enabled = False
                                Me.StartBackgroundForLoadingPrice()
                            Else
                                Try
                                    objForm.Items.Item("PU").Enabled = False
                                    Dim PriceList As String = "SELECT * FROM ""@SBO_UPDATE"" "
                                    oRsPriceList = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim UpdateQuery As String = "UPDATE ""@SBO_PL"" SET ""U_PRICEUP""='Yes' Where ""DocNum""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "'"
                                    oRsSubmit.DoQuery(UpdateQuery)
                                    oRsPriceList.DoQuery(PriceList)

                                    objMatrix.Columns.Item("UOM").Editable = False
                                    objMatrix.Columns.Item("NPrice").Editable = False
                                    objMatrix.Columns.Item("BarCode").Editable = False
                                    objMatrix.Columns.Item("ICode").Editable = False
                                    objMatrix.Columns.Item("Remarks").Editable = False

                                    Dim UpdateQuery1 As String
                                    ' "C:\Logs\Price Update\RestartService (1) 1 1.bat"
                                    'Try
                                    '    System.Diagnostics.Process.Start("C:\Logs\Price Update\PriceUpdate_Service.bat")
                                    'Catch ex As Exception

                                    'End Try
                                    If oRsPriceList.RecordCount > 0 Then

                                        UpdateQuery1 = "UPDATE ""@SBO_UPDATE"" SET ""U_Flag""='1', ""U_Status""='Price update is Started',""U_DOCNUM""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "' "

                                    Else

                                        UpdateQuery1 = "INSERT INTO ""@SBO_UPDATE"" (""Code"", ""Name"", ""U_Flag"", ""U_SDate"", ""U_Status"",""U_DOCNUM"") VALUES ('1', '1', '1', CURDATE(), 'Process started','" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "' );"
                                    End If

                                    'Dim code As String = oRsPriceList.Fields.Item("Code").Value
                                    oRsSubmit = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    'Dim UpdateQuery As String = "UPDATE ""@SBO_UPDATE"" SET ""U_Flag""='1', ""U_Status""='Price update is Started' Where ""Code""='" & code & "'"
                                    oRsSubmit.DoQuery(UpdateQuery1)
                                    '   System.Diagnostics.Process.Start("C:\Logs\Price Update\PriceUpdate_Service.bat")

                                    '   objMain.objApplication.ActivateMenuItem("1304")
                                Catch ex As Exception
                                    objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                                End Try
                            End If
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objMatrix = objForm.Items.Item("Mtx").Specific
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    If oCFL.UniqueID = "CFL_UOM" Then
                        Dim ItemCode As String = objMatrix.Columns.Item("ICode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim
                        '     Me.CFLFilter("CFL_UOM", ItemCode)
                    End If
                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If oCFL.UniqueID = "CFL_UOM" Then
                            'oDS1.SetValue("U_UOM", oDS1.Offset, oDT.GetValue("UomCode", 0))
                            'Dim test As String = oDT.GetValue("UomCode", 0)
                            'Try
                            '    objMatrix.Columns.Item("UOM").Cells.Item(pVal.Row).Specific.Value = oDT.GetValue("UomCode", 0)
                            'Catch ex As Exception
                            'End Try
                            '----------------------------------------
                            oDS1.SetValue("LineID", oDS1.Offset, pVal.Row)
                            oDS1.SetValue("U_ICode", oDS1.Offset, objMatrix.Columns.Item("ICode").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_IName", oDS1.Offset, objMatrix.Columns.Item("IName").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_BarCode", oDS1.Offset, objMatrix.Columns.Item("BarCode").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_UOM", oDS1.Offset, oDT.GetValue("UomCode", 0))
                            oDS1.SetValue("U_NPrice", oDS1.Offset, objMatrix.Columns.Item("NPrice").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_Remarks", oDS1.Offset, objMatrix.Columns.Item("Remarks").Cells.Item(pVal.Row).Specific.Value)
                            objMatrix.SetLineData(pVal.Row)
                            '------------------------------------------
                        End If
                        If oCFL.UniqueID = "CFL_ITEM" Then
                            oDS1.SetValue("LineID", oDS1.Offset, pVal.Row)
                            oDS1.SetValue("U_ICode", oDS1.Offset, oDT.GetValue("ItemCode", 0))
                            oDS1.SetValue("U_IName", oDS1.Offset, oDT.GetValue("ItemName", 0))
                            oDS1.SetValue("U_BarCode", oDS1.Offset, objMatrix.Columns.Item("BarCode").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_NPrice", oDS1.Offset, objMatrix.Columns.Item("NPrice").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_Remarks", oDS1.Offset, objMatrix.Columns.Item("Remarks").Cells.Item(pVal.Row).Specific.Value)
                            objMatrix.SetLineData(pVal.Row)
                        End If
                    End If


                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub CFLWhsCodeFilter(ByVal FormUID As String, ByVal CFL_ID As String, ByVal BPLid As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim GetWhsCode As String = " Select ""WhsCode"" From OWHS Where ""BPLid"" = '" & BPLid & "' And IFNULL(""U_Transit"",'') <> 'Yes' "
            Dim oRsGetWhsCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetWhsCode.DoQuery(GetWhsCode)

            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            If oRsGetWhsCode.RecordCount > 0 Then
                oCondition = oConditions.Add()
                oCondition.Alias = "WhsCode"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = oRsGetWhsCode.Fields.Item(0).Value
                oChooseFromList.SetConditions(oConditions)
                oRsGetWhsCode.MoveNext()
                For i As Integer = 1 To oRsGetWhsCode.RecordCount - 1
                    oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCondition = oConditions.Add()
                    oCondition.Alias = "WhsCode"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = oRsGetWhsCode.Fields.Item(0).Value
                    oChooseFromList.SetConditions(oConditions)
                    oRsGetWhsCode.MoveNext()
                Next
            Else
                oCondition = oConditions.Add()
                oCondition.Alias = "WhsCode"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = ""
                oChooseFromList.SetConditions(oConditions)
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
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
                        If eventInfo.ItemUID <> "29" Then
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
                        objMatrix = objForm.Items.Item("29").Specific
                        If eventInfo.ItemUID = "29" And eventInfo.ColUID = "V_-1" Then
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("Delete Row") = False Then
                                    oCreationPackage.UniqueID = "Delete Row"
                                    oCreationPackage.String = "Delete Row"
                                    oCreationPackage.Enabled = True
                                    oMenus.AddEx(oCreationPackage)
                                End If
                                If oMenus.Exists("Add Row") = False Then
                                    oCreationPackage.UniqueID = "Add Row"
                                    oCreationPackage.String = "Add Row"
                                    oCreationPackage.Enabled = True
                                    oMenus.AddEx(oCreationPackage)
                                End If
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        ElseIf eventInfo.ItemUID = "29" And objMatrix.RowCount <= 1 Then
                            oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                            oMenus = oMenuItem.SubMenus
                            Try
                                If oMenus.Exists("Delete Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Delete Row")
                                End If
                                If oMenus.Exists("Add Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Add Row")
                                End If
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If
                        If eventInfo.ItemUID <> "29" Then
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("Delete Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Delete Row")
                                End If
                                If oMenus.Exists("Add Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Add Row")
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
                        If oMenus.Exists("Add Row") = True Then
                            objMain.objApplication.Menus.RemoveEx("Add Row")
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

    Sub CFLFilter(ByVal CFL_ID As String, ByVal ItemCode As String)
        Try
            Dim Query As String = " SELECT (select ""UomCode"" from OUOM Where ""UomEntry""=T1.""UomEntry"")  FROM ""UGP1""  T1 Where T1.""UgpEntry""=(SELECT T0.""UgpEntry"" FROM OITM T0 WHERE T0.""ItemCode""='" & ItemCode & "')"
            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsQuery.DoQuery(Query)

            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objForm.ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()

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
            oChooseFromList.SetConditions(oConditions)
        Catch ex As Exception

        End Try
    End Sub

    Sub Read13c(ByVal PackingList As String)

        Dim barcode, Pricelist, nPrice As String
        Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx").Specific
        objMatrix.Clear()
        Dim oRS As SAPbobsCOM.Recordset
        oRS = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

        Dim createTableSQL As String = "
CREATE COLUMN TABLE ""TNX_PRICEUPDATE""
(
    ""BARCODE"" NVARCHAR(100) COMMENT 'Barcode of the item',
    ""PRICELIST"" NVARCHAR(100) COMMENT 'Price list name',
    ""PRICE"" DECIMAL(11,5) CS_FIXED COMMENT 'Item price',
    ""CREATEDBY"" NVARCHAR(100) COMMENT 'Created by user',
    ""LINENUM"" INTEGER COMMENT 'Row number in file'

)
UNLOAD PRIORITY 5 AUTO MERGE;
"
        Try
            oRS.DoQuery(createTableSQL)
            objMain.objApplication.StatusBar.SetText("Table TNX_PRICEUPDATE created successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
        End Try
        Dim CreatedBy As String = objMain.objCompany.UserSignature

        Dim connStr As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & PackingList & ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';")
        connStr.Open()

        ' Get sheets
        Dim dtSheets As DataTable = connStr.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
        If dtSheets Is Nothing OrElse dtSheets.Rows.Count = 0 Then
            objMain.objApplication.StatusBar.SetText("No sheets found.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If

        Dim firstSheetName As String = dtSheets.Rows(0)("TABLE_NAME").ToString()

        ' Load data
        Dim query As String = "SELECT * FROM [" & firstSheetName & "]"
        Dim adap As New OleDbDataAdapter(query, connStr)
        Dim dtNew As New DataTable()
        adap.Fill(dtNew)

        connStr.Close()
        Try
            Dim QRY As String = "DELETE  FROM  ""TNX_PRICEUPDATE"" WHERE ""CREATEDBY""  ='" & CreatedBy & "'"
            oRS.DoQuery(QRY)
        Catch ex As Exception

        End Try

        For i As Integer = 0 To dtNew.Rows.Count - 1
            If dtNew.Rows(i).Item(0) Is DBNull.Value OrElse String.IsNullOrWhiteSpace(dtNew.Rows(i).Item(0).ToString()) Then
                Continue For
            End If

            ' Read values
            barcode = dtNew.Rows(i).Item(0).ToString().Trim().Replace("'", "''")
            Pricelist = dtNew.Rows(i).Item(1).ToString().Trim().Replace("'", "''")
            nPrice = dtNew.Rows(i).Item(2).ToString().Trim().Replace("'", "''")
            Dim lineNumber As Integer = i + 1

            ' One INSERT per row
            Dim sql As String = String.Format(
            "INSERT INTO ""TNX_PRICEUPDATE"" (""BARCODE"", ""PRICELIST"", ""PRICE"", ""CREATEDBY"", ""LINENUM"") VALUES ('{0}', '{1}', '{2}', '{3}', '{4}')",
            barcode, Pricelist, nPrice, CreatedBy, lineNumber)

            oRS.DoQuery(sql)
        Next

        objMain.objApplication.StatusBar.SetText("Upload completed.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Try

            query = "Call ""TNX_BULKPRICEUPDATE""('" & CreatedBy & "');"

            Dim oRecordset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordset.DoQuery(query)
            Dim itemCode, uom, IUOM, itemName, UomName, ConversionRate, cost, PriceListName, rsp As String
            'Dim ConversionRate, cost, rsp As Double

            If oRecordset.RecordCount > 0 Then
                For I As Integer = 1 To oRecordset.RecordCount
                    objMatrix.AddRow()
                    itemCode = oRecordset.Fields.Item("ItemCode").Value.ToString()

                    oDS1.SetValue("LineId", oDS1.Offset, I)
                    oDS1.SetValue("U_ICode", oDS1.Offset, oRecordset.Fields.Item("ItemCode").Value.ToString())
                    oDS1.SetValue("U_IName", oDS1.Offset, oRecordset.Fields.Item("ItemName").Value.ToString())
                    oDS1.SetValue("U_BarCode", oDS1.Offset, oRecordset.Fields.Item("BARCODE").Value)
                    oDS1.SetValue("U_UOM", oDS1.Offset, oRecordset.Fields.Item("UomEntry").Value.ToString())
                    oDS1.SetValue("U_PLIST", oDS1.Offset, oRecordset.Fields.Item("PRICELIST").Value)
                    oDS1.SetValue("U_NPrice", oDS1.Offset, oRecordset.Fields.Item("PRICE").Value)
                    oDS1.SetValue("U_Remarks", oDS1.Offset, oRecordset.Fields.Item("REMARK").Value)
                    oDS1.SetValue("U_IUOM", oDS1.Offset, oRecordset.Fields.Item("IUoMEntry").Value.ToString())
                    oDS1.SetValue("U_UNAME", oDS1.Offset, oRecordset.Fields.Item("UomName").Value.ToString())
                    oDS1.SetValue("U_CF", oDS1.Offset, oRecordset.Fields.Item("ConversionRate").Value.ToString())
                    oDS1.SetValue("U_PLNAME", oDS1.Offset, oRecordset.Fields.Item("PriceListName").Value.ToString())
                    oDS1.SetValue("U_Cost", oDS1.Offset, oRecordset.Fields.Item("Cost").Value.ToString())
                    oDS1.SetValue("U_Margin", oDS1.Offset, oRecordset.Fields.Item("margin").Value)
                    oDS1.SetValue("U_RSP", oDS1.Offset, oRecordset.Fields.Item("RSP").Value)
                    objMatrix.SetLineData(objMatrix.VisualRowCount)

                    '   DisplayProgressBar(oRecordset.RecordCount, I, "Importing CSV")
                    oRecordset.MoveNext()

                Next

            End If
            '    oPrg.Stop()
            objMain.objApplication.StatusBar.SetText("CSV data successfully loaded to matrix.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


        Catch ex As Exception
            '    If oPrg IsNot Nothing Then oPrg.Stop()
            objMain.objApplication.StatusBar.SetText("CSV Import Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
    Sub Read13(ByVal PackingList As String)

        Dim barcode, Pricelist, nPrice As String
        Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx").Specific
        objMatrix.Clear()
        Dim oRecordset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim oRS As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim dtNew As New DataTable()
        ' === Create table if not exists ===
        Dim createTableSQL As String = "
CREATE COLUMN TABLE ""TNX_PRICEUPDATE""
(
    ""BARCODE"" NVARCHAR(100),
    ""PRICELIST"" NVARCHAR(100),
    ""PRICE"" DECIMAL(11,5) CS_FIXED,
    ""CREATEDBY"" NVARCHAR(100),
    ""LINENUM"" INTEGER
)
UNLOAD PRIORITY 5 AUTO MERGE;
"
        Try
            oRS.DoQuery(createTableSQL)
        Catch ex As Exception
            If Not ex.Message.ToLower().Contains("already exists") AndAlso Not ex.Message.ToLower().Contains("cannot use duplicate table name") Then
                objMain.objApplication.StatusBar.SetText("Error creating table: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        End Try

        Dim CreatedBy As String = objMain.objCompany.UserSignature

        objMain.objApplication.StatusBar.SetText("Reading Excel & preparing upload...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        ' === Read Excel ===
        Try
            Dim connStr As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & PackingList & ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';")
            connStr.Open()

            Dim dtSheets As DataTable = connStr.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
            If dtSheets Is Nothing OrElse dtSheets.Rows.Count = 0 Then
                objMain.objApplication.StatusBar.SetText("No sheets found in Excel.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                connStr.Close()
                Exit Sub
            End If

            Dim firstSheetName As String = dtSheets.Rows(0)("TABLE_NAME").ToString()
            Dim adap As New OleDbDataAdapter("SELECT * FROM [" & firstSheetName & "]", connStr)

            adap.Fill(dtNew)
            connStr.Close()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error reading Excel: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End Try
        ' === Delete old ===
        Try
            oRS.DoQuery("DELETE FROM ""TNX_PRICEUPDATE"" WHERE ""CREATEDBY"" = '" & CreatedBy & "'")
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error deleting old records: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End Try

        objMain.objApplication.StatusBar.SetText("Inserting records into SAP HANA...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        ' === Bulk Insert using UNION ALL ===
        Try
            Dim sb As New Text.StringBuilder
            Dim rowsAdded As Boolean = False

            For i As Integer = 0 To dtNew.Rows.Count - 1

                If dtNew.Rows(i).Item(0) Is DBNull.Value _
       OrElse String.IsNullOrWhiteSpace(dtNew.Rows(i).Item(0).ToString()) Then Continue For

                barcode = dtNew.Rows(i).Item(0).ToString().Trim().Replace("'", "''")
                Pricelist = dtNew.Rows(i).Item(1).ToString().Trim().Replace("'", "''")
                nPrice = dtNew.Rows(i).Item(2).ToString().Trim().Replace("'", "''")
                Dim lineNumber As Integer = i + 1

                If Not rowsAdded Then
                    sb.AppendLine("INSERT INTO ""TNX_PRICEUPDATE""")
                    sb.AppendLine("(")
                    sb.AppendLine("""BARCODE"", ""PRICELIST"", ""PRICE"", ""CREATEDBY"", ""LINENUM""")
                    sb.AppendLine(")")
                    sb.Append("SELECT ")
                    rowsAdded = True
                Else
                    sb.Append(" UNION ALL SELECT ")
                End If

                sb.AppendFormat("'{0}', '{1}', {2}, '{3}', {4} FROM DUMMY", barcode, Pricelist, nPrice, CreatedBy, lineNumber)

            Next



            If rowsAdded Then
                oRS.DoQuery(sb.ToString())
                objMain.objApplication.StatusBar.SetText("Upload completed successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objMain.objApplication.StatusBar.SetText("No valid rows found for upload.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Exit Sub
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Required Columns not found in sheet 1 ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End Try
        objMain.objApplication.StatusBar.SetText("Filling Matrix Please Wait", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        ' === Call stored proc ===
        Try


            Dim query As String = "CALL ""TNX_BULKPRICEUPDATE""('" & CreatedBy & "');"
            oRecordset.DoQuery(query)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error loading data call TNX_BULKPRICEUPDATE(" & CreatedBy & ") : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End Try
        Try
            If oRecordset.RecordCount > 0 Then
                objForm.Freeze(True)

                objMatrix.Clear()
                objMatrix.AddRow(oRecordset.RecordCount)  ' Pre-add all rows once

                Dim lineNum As Integer = 1

                While Not oRecordset.EoF
                    oDS1.SetValue("LineId", oDS1.Offset, lineNum)
                    oDS1.SetValue("U_ICode", oDS1.Offset, oRecordset.Fields.Item("ItemCode").Value.ToString())
                    oDS1.SetValue("U_IName", oDS1.Offset, oRecordset.Fields.Item("ItemName").Value.ToString())
                    oDS1.SetValue("U_BarCode", oDS1.Offset, oRecordset.Fields.Item("BARCODE").Value)
                    oDS1.SetValue("U_UOM", oDS1.Offset, oRecordset.Fields.Item("UomEntry").Value.ToString())
                    oDS1.SetValue("U_PLIST", oDS1.Offset, oRecordset.Fields.Item("PRICELIST").Value)
                    oDS1.SetValue("U_NPrice", oDS1.Offset, oRecordset.Fields.Item("PRICE").Value)
                    oDS1.SetValue("U_Remarks", oDS1.Offset, oRecordset.Fields.Item("REMARK").Value)
                    oDS1.SetValue("U_IUOM", oDS1.Offset, oRecordset.Fields.Item("IUoMEntry").Value.ToString())
                    oDS1.SetValue("U_UNAME", oDS1.Offset, oRecordset.Fields.Item("UomName").Value.ToString())
                    oDS1.SetValue("U_CF", oDS1.Offset, oRecordset.Fields.Item("ConversionRate").Value.ToString())
                    oDS1.SetValue("U_PLNAME", oDS1.Offset, oRecordset.Fields.Item("PriceListName").Value.ToString())
                    oDS1.SetValue("U_Cost", oDS1.Offset, oRecordset.Fields.Item("Cost").Value.ToString())
                    oDS1.SetValue("U_Margin", oDS1.Offset, oRecordset.Fields.Item("margin").Value)
                    oDS1.SetValue("U_RSP", oDS1.Offset, oRecordset.Fields.Item("RSP").Value)

                    objMatrix.SetLineData(lineNum)
                    oRecordset.MoveNext()
                    lineNum += 1
                End While

                objForm.Freeze(False)
            End If

            objMain.objApplication.StatusBar.SetText("Matrix filled successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Matrix  Import Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            objForm.Freeze(False)
        End Try

        ' Optional: force cleanup
        GC.Collect()

    End Sub


    Sub Read13X(ByVal PackingList As String)

        Dim barcode, Pricelist, nPrice As String
        Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx").Specific
        objMatrix.Clear()
        Dim oRS As SAPbobsCOM.Recordset
        oRS = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

        Dim createTableSQL As String = "
CREATE COLUMN TABLE ""TNX_PRICEUPDATE""
(
    ""BARCODE"" NVARCHAR(100) COMMENT 'Barcode of the item',
    ""PRICELIST"" NVARCHAR(100) COMMENT 'Price list name',
    ""PRICE"" DECIMAL(11,5) CS_FIXED COMMENT 'Item price',
    ""CREATEDBY"" NVARCHAR(100) COMMENT 'Created by user',
    ""LINENUM"" INTEGER COMMENT 'Row number in file'

)
UNLOAD PRIORITY 5 AUTO MERGE;
"
        Try
            oRS.DoQuery(createTableSQL)
            objMain.objApplication.StatusBar.SetText("Table TNX_PRICEUPDATE created successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
        End Try
        Dim CreatedBy As String = objMain.objCompany.UserSignature

        Dim connStr As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & PackingList & ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';")
        connStr.Open()

        ' Get sheets
        Dim dtSheets As DataTable = connStr.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
        If dtSheets Is Nothing OrElse dtSheets.Rows.Count = 0 Then
            objMain.objApplication.StatusBar.SetText("No sheets found.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If

        Dim firstSheetName As String = dtSheets.Rows(0)("TABLE_NAME").ToString()

        ' Load data
        Dim query As String = "SELECT * FROM [" & firstSheetName & "]"
        Dim adap As New OleDbDataAdapter(query, connStr)
        Dim dtNew As New DataTable()
        adap.Fill(dtNew)

        connStr.Close()
        Try


            Dim QRY As String = "DELETE  FROM  ""TNX_PRICEUPDATE"" WHERE ""CREATEDBY""  ='" & CreatedBy & "'"
            oRS.DoQuery(QRY)
        Catch ex As Exception

        End Try
        ' Prepare Recordset object
        ' Dim oRS As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

        ' Loop rows — starting from row 0 (skip header handled by HDR=YES)
        For i As Integer = 0 To dtNew.Rows.Count - 1
            If dtNew.Rows(i).Item(0) Is DBNull.Value OrElse String.IsNullOrWhiteSpace(dtNew.Rows(i).Item(0).ToString()) Then
                Continue For
            End If

            ' Read values
            barcode = dtNew.Rows(i).Item(0).ToString().Trim().Replace("'", "''")
            Pricelist = dtNew.Rows(i).Item(1).ToString().Trim().Replace("'", "''")
            nPrice = dtNew.Rows(i).Item(2).ToString().Trim().Replace("'", "''")


            ' One INSERT per row
            Dim sql As String = String.Format(
            "INSERT INTO ""TNX_PRICEUPDATE"" (""BARCODE"", ""PRICELIST"", ""PRICE"", ""CREATEDBY"") VALUES ('{0}', '{1}', '{2}', '{3}')",
            barcode, Pricelist, nPrice, CreatedBy)

            oRS.DoQuery(sql)
        Next

        objMain.objApplication.StatusBar.SetText("Upload completed.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)





        'Dim rows As New List(Of String)
        'Using sr As New StreamReader(PackingList)
        '    While Not sr.EndOfStream
        '        Dim line As String = sr.ReadLine()
        '        If Not String.IsNullOrWhiteSpace(line) Then
        '            rows.Add(line)
        '        End If
        '    End While
        'End Using

        'If rows.Count <= 1 Then
        '    Throw New Exception("No data rows found.")
        'End If

        'Try
        '    For i As Integer = 1 To rows.Count - 1 ' Skip header
        '        Dim rowValues As String() = rows(i).Split(","c)
        '        If rowValues.Length >= 3 Then
        '        barcode = rowValues(0).Trim().Replace("'", "")
        '        Pricelist = rowValues(1).Trim().Replace("'", "")
        '        nPrice = rowValues(2).Trim().Replace("'", "")

        '        ' ✅ Build ONE INSERT at a time
        '        Dim singleSql As String = String.Format(
        '    "INSERT INTO ""TNX_PRICEUPDATE"" (""BARCODE"", ""PRICELIST"", ""PRICE"", ""CREATEDBY"") " &
        '    "VALUES ('{0}', '{1}', '{2}', '{3}')",
        '    barcode, Pricelist, nPrice, user
        ')

        '            ' ✅ Run immediately
        '            oRS.DoQuery(singleSql)
        '        End If
        '    Next

        'Catch ex As Exception
        '    objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        'End Try


        Try

            'query = "SELECT T1.""ItemCode"", B.""UomEntry"",T1.""ItemName"",""IUoMEntry"",(select ""GET_COST""(T1.""ItemCode"",B.""UomEntry"") from dummy) AS ""COST"",(select ""GET_COST""(T1.""ItemCode"",B.""UomEntry"",) from dummy) AS ""COST""  FROM ""OITM"" T1  INNER JOIN ""OBCD"" B ON T1.""ItemCode"" = B.""ItemCode"" WHERE  B.""BcdCode"" ='" & barcode & "'"
            '  query = "SELECT T1.""ItemCode"", B.""UomEntry"",T1.""ItemName"" FROM ""OITM"" T1  INNER JOIN ""OBCD"" B ON T1.""ItemCode"" = B.""ItemCode"" WHERE  B.""BcdCode"" ='" & barcode & "'"
            'query = "Call ""TNX_ITEMDETAILS""('" & barcode & "','" & Pricelist & "');"
            query = "Call ""TNX_BULKPRICEUPDATE""('" & CreatedBy & "');"

            Dim oRecordset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordset.DoQuery(query)
            Dim itemCode, uom, IUOM, itemName, UomName, ConversionRate, cost, PriceListName, rsp As String
            'Dim ConversionRate, cost, rsp As Double

            If oRecordset.RecordCount > 0 Then
                For I As Integer = 1 To oRecordset.RecordCount
                    objMatrix.AddRow()
                    itemCode = oRecordset.Fields.Item("ItemCode").Value.ToString()
                    If itemCode IsNot Nothing Then


                        uom = oRecordset.Fields.Item("UomEntry").Value.ToString()
                        IUOM = oRecordset.Fields.Item("IUoMEntry").Value.ToString()
                        itemName = oRecordset.Fields.Item("ItemName").Value.ToString()
                        UomName = oRecordset.Fields.Item("UomName").Value.ToString()
                        ConversionRate = oRecordset.Fields.Item("ConversionRate").Value.ToString()
                        cost = oRecordset.Fields.Item("Cost").Value.ToString()
                        PriceListName = oRecordset.Fields.Item("PriceListName").Value.ToString()
                        rsp = oRecordset.Fields.Item("RSP").Value
                        barcode = oRecordset.Fields.Item("BARCODE").Value
                        Pricelist = oRecordset.Fields.Item("PRICELIST").Value
                        nPrice = oRecordset.Fields.Item("PRICE").Value


                        oDS1.SetValue("LineId", oDS1.Offset, I)
                        oDS1.SetValue("U_ICode", oDS1.Offset, itemCode)
                        oDS1.SetValue("U_IName", oDS1.Offset, itemName)
                        oDS1.SetValue("U_BarCode", oDS1.Offset, barcode)
                        oDS1.SetValue("U_UOM", oDS1.Offset, uom)
                        oDS1.SetValue("U_PLIST", oDS1.Offset, Pricelist)
                        oDS1.SetValue("U_NPrice", oDS1.Offset, nPrice)
                        oDS1.SetValue("U_Remarks", oDS1.Offset, "")
                        oDS1.SetValue("U_IUOM", oDS1.Offset, IUOM)
                        oDS1.SetValue("U_UNAME", oDS1.Offset, UomName)
                        oDS1.SetValue("U_CF", oDS1.Offset, ConversionRate)
                        oDS1.SetValue("U_PLNAME", oDS1.Offset, PriceListName)
                        oDS1.SetValue("U_Cost", oDS1.Offset, cost)
                        Dim margin As Double = 0
                        margin = nPrice - cost
                        oDS1.SetValue("U_Margin", oDS1.Offset, margin)
                        oDS1.SetValue("U_RSP", oDS1.Offset, rsp)


                    Else
                        oDS1.SetValue("LineId", oDS1.Offset, I)
                        oDS1.SetValue("U_ICode", oDS1.Offset, "")
                        oDS1.SetValue("U_IName", oDS1.Offset, "")
                        oDS1.SetValue("U_UOM", oDS1.Offset, "")
                        oDS1.SetValue("U_Remarks", oDS1.Offset, "")
                        oDS1.SetValue("U_IUOM", oDS1.Offset, "")
                        oDS1.SetValue("U_BarCode", oDS1.Offset, barcode)
                        oDS1.SetValue("U_PLIST", oDS1.Offset, Pricelist)
                        oDS1.SetValue("U_NPrice", oDS1.Offset, nPrice)
                        oDS1.SetValue("U_Remarks", oDS1.Offset, "Barcode Not Found/ Item Not found")
                        oDS1.SetValue("U_IUOM", oDS1.Offset, "")
                        oDS1.SetValue("U_UNAME", oDS1.Offset, "")
                        oDS1.SetValue("U_CF", oDS1.Offset, 0)
                        oDS1.SetValue("U_PLNAME", oDS1.Offset, "")
                        oDS1.SetValue("U_Cost", oDS1.Offset, 0)
                        oDS1.SetValue("U_Margin", oDS1.Offset, 0)
                        oDS1.SetValue("U_RSP", oDS1.Offset, 0)

                    End If
                    objMatrix.SetLineData(objMatrix.VisualRowCount)

                    '    DisplayProgressBar(oRecordset.RecordCount, I, "Importing CSV")

                    ' objMatrix.AddRow()
                Next

            End If
            '    oPrg.Stop()
            objMain.objApplication.StatusBar.SetText("CSV data successfully loaded to matrix.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


        Catch ex As Exception
            '    If oPrg IsNot Nothing Then oPrg.Stop()
            objMain.objApplication.StatusBar.SetText("CSV Import Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
    Sub StartBackgroundForLoading(Optional ByVal BackGroundWrkr As Boolean = True)
        Try
            If BackGroundWrkr = True Then
                If BackgroundWorker_CSVLoad.IsBusy <> True Then
                    BackgroundWorker_CSVLoad.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Calling Unselecting worker Error : " & ex.Message)
        End Try
    End Sub

    Private Sub BackgroundWorker_Loading_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_CSVLoad.DoWork

        Dim Pathxl As String = oDBs_Head.GetValue("U_upload", 0)
        If Pathxl <> "" Then
            Me.Read13(Pathxl)
        Else
            objMain.objApplication.StatusBar.SetText("Please Select Vendor..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If

    End Sub
    Sub DisplayProgressBar(totalRecords As Integer, currentRecord As Integer, Optional taskName As String = "")
        If totalRecords = 0 Then Exit Sub

        Dim progressBarLength As Integer = 50
        Dim progressPercentage As Double = (currentRecord / totalRecords) * 100
        Dim filledLength As Integer = CInt((currentRecord / totalRecords) * progressBarLength)

        Dim progressBar As String = New String("█"c, filledLength).PadRight(progressBarLength, " "c)
        Dim message As String = If(taskName <> "", "[" & taskName & "] ", "") &
                            vbCr & progressBar & " " & Math.Round(progressPercentage, 2) & "%"


        objMain.objApplication.StatusBar.SetText(message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

    End Sub

    Sub StartBackgroundForLoadingPrice(Optional ByVal BackGroundWrkr As Boolean = True)
        Try
            If BackGroundWrkr = True Then
                If BackgroundWorker_SalaryJV.IsBusy <> True Then
                    BackgroundWorker_SalaryJV.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Calling Unselecting worker Error : " & ex.Message)
        End Try
    End Sub
    Private Sub BackgroundWorker_Loading_DoWork1(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_SalaryJV.DoWork
        oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PL")
        oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_PL_C0")

        objMatrix = objForm.Items.Item("Mtx").Specific
        For i As Integer = 1 To objMatrix.VisualRowCount
            DisplayProgressBar(objMatrix.VisualRowCount, i, "Updating Items")
            ' objMain.objApplication.StatusBar.SetText("Updating records " & i & " out of " & objMatrix.VisualRowCount & "...  ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            Dim item As String = objMatrix.Columns.Item("ICode").Cells.Item(i).Specific.Value.ToString().Trim()
            Dim UOM As String = objMatrix.Columns.Item("UOM").Cells.Item(i).Specific.Value.ToString().Trim()
            Dim NPrice As Double = CDbl(objMatrix.Columns.Item("NPrice").Cells.Item(i).Specific.Value)
            Dim PriceList As String = objMatrix.Columns.Item("PLIST").Cells.Item(i).Specific.Value.ToString
            Dim barcode As String = objMatrix.Columns.Item("BarCode").Cells.Item(i).Specific.Value.ToString
            Dim status As String = objMatrix.Columns.Item("Status").Cells.Item(i).Specific.Value.ToString
            Dim Remarks As String = objMatrix.Columns.Item("Remarks").Cells.Item(i).Specific.Value.ToString
            'Yes' ,""U_Remarks"" ='Price Updated Successfully'
            If item <> "" And status <> "Yes" And Remarks <> "Price Updated Successfully" Then
                ' objMain.objApplication.StatusBar.SetText("Prices are updating for ItemCode:" + item, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                'PrgBar.Value = i
                'PrgBar.Text = "Updating Price for ItemCode:" + item
                Try
                    Me.UpdatePrice(item, UOM, PriceList, NPrice, barcode, objForm.Items.Item("DocNum").Specific.Value.ToString().Trim(), i, oDBs_Head.GetValue("DocEntry", oDBs_Head.Offset))

                Catch ex As Exception
                    '  PrgBar.Stop()
                    objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    objForm.Items.Item("PU").Enabled = True

                End Try
            End If
        Next
        objMain.objApplication.StatusBar.SetText("Process Completed SuccessFully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        If Sucess <> "" Then
            Dim UpdateQuery1 As String = "UPDATE ""@SBO_PL"" SET ""U_Success""='" & Sucess & "' Where ""DocNum""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "'"
            Dim oRsSubmit1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oRsSubmit1.DoQuery(UpdateQuery1)
        End If
        If failed <> "" Then
            Dim UpdateQuery1 As String = "UPDATE ""@SBO_PL"" SET ""U_Failed""='" & failed & "' Where ""DocNum""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "'"
            Dim oRsSubmit1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oRsSubmit1.DoQuery(UpdateQuery1)
        End If
        '--------------------------------------
        Dim UpdateQuery As String = "UPDATE ""@SBO_PL"" SET ""U_DocSts""='Closed' Where ""DocNum""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "'"
        Dim oRsSubmit As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        oRsSubmit.DoQuery(UpdateQuery)

        '--------------------------------------
        ' PrgBar.Stop()
        ' System.Runtime.InteropServices.Marshal.ReleaseComObject(PrgBar)
        objMain.objApplication.ActivateMenuItem("1304")


    End Sub

    Sub UpdatePrice(ByVal ItemCode As String, ByVal UOM As String, ByVal PriceList As String, ByVal SalesPrice As Double, ByVal barcode As String, ByVal DocNum As String, ByVal pvalrow As String, ByVal DocEntry As String)
        Try
            Dim oItemMaster As SAPbobsCOM.Items = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
            Dim oRsPriceList As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PL")
            oDS1 = objForm.DataSources.DBDataSources.Item("@SBO_PL_C0")
            Dim UomCount As Integer = 0


            Dim PriceListQry1 As String = "SELECT T0.""ListName"" FROM OPLN T0 WHERE T0.""ListNum"" = " & PriceList & ""
            oRsPriceList.DoQuery(PriceListQry1)
            Dim PriceList13 As String = oRsPriceList.Fields.Item(0).Value
            objMain.objApplication.StatusBar.SetText("Prices are updating for ItemCode: " + ItemCode + " for  PriceList " + PriceList13 + "....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            oItemMaster.GetByKey(ItemCode)

            Dim UOMQuery As String = "select ""UomEntry"" from OUOM Where ""UomEntry""='" & UOM & "'"
            Dim oRsUOMQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsUOMQuery.DoQuery(UOMQuery)

            '--------------------------
            Dim UomEntry As Integer = CInt(oRsUOMQuery.Fields.Item(0).Value)
            Dim DetQuery As String = "select (select ""UomEntry"" from OUOM Where TO_CHAR(""UomEntry"") = TO_CHAR(""IUoMEntry"")) from OITM where ""ItemCode""='" & ItemCode & "' "
            Dim oRsDetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsDetQuery.DoQuery(DetQuery)
            For i As Integer = 0 To oItemMaster.PriceList.Count() - 1
                oItemMaster.PriceList.SetCurrentLine(i)
                Dim strPrlist As String = oItemMaster.PriceList.PriceList
                If strPrlist = PriceList Then
                    If oRsDetQuery.Fields.Item(0).Value = UomEntry Then
                        oItemMaster.PriceList.Price = SalesPrice
                        Exit For
                    Else
                        Dim UOMFoundFlag As String = ""
                        For j As Integer = 0 To oItemMaster.PriceList.UoMPrices.Count - 1
                            oItemMaster.PriceList.UoMPrices.SetCurrentLine(j)
                            If oItemMaster.PriceList.UoMPrices.UoMEntry = UomEntry Then
                                oItemMaster.PriceList.UoMPrices.Price = SalesPrice
                                UOMFoundFlag = "Yes"
                                Exit For
                            End If

                        Next
                        If UOMFoundFlag = "" Then


                            Dim GetUomPriceListCount As String = "Select Count(*) From ITM9 Where ""ItemCode"" = '" & ItemCode & "' And ""PriceList"" = '" & strPrlist & "'"
                            Dim oRsGetUomPriceListCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetUomPriceListCount.DoQuery(GetUomPriceListCount)
                            If oRsGetUomPriceListCount.Fields.Item(0).Value = 0 Then
                                UomCount = 0
                            Else
                                UomCount = oRsGetUomPriceListCount.Fields.Item(0).Value
                            End If
                            oItemMaster.PriceList.UoMPrices.Add()
                            oItemMaster.PriceList.UoMPrices.SetCurrentLine(UomCount)
                            oItemMaster.PriceList.UoMPrices.Price = SalesPrice
                            oItemMaster.PriceList.UoMPrices.UoMEntry = UomEntry
                            oItemMaster.PriceList.UoMPrices.Auto = SAPbobsCOM.BoYesNoEnum.tNO

                            ' oItemMaster.PriceList.UoMPrices.Add()

                        End If
                        Exit For
                    End If
                End If

            Next



            Dim RetValue As Integer = oItemMaster.Update()

            '-----------------------------------------

            If RetValue <> 0 Then
                Dim ErrCode As Long
                Dim ErrMsg As String
                objMain.objCompany.GetLastError(ErrCode, ErrMsg)
                objMain.objApplication.StatusBar.SetText("Unable to update price" & ErrCode & ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Me.WriteLog(barcode, ItemCode, UomEntry, PriceList, SalesPrice, ErrCode, ErrMsg, "ERROR", DocNum)
                Dim oRsSubmit4 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim UpdateQuery4 As String = "UPDATE ""@SBO_PL_C0"" SET ""U_Status""='No' ,""U_Remarks"" ='" & ErrMsg & "'  Where ""DocEntry""='" & DocEntry & "' and ""LineId""='" & pvalrow & "'"
                oRsSubmit4.DoQuery(UpdateQuery4)
            Else
                Me.WriteLog(barcode, ItemCode, UomEntry, PriceList, SalesPrice, "Success", ItemCode + "-Item price updated successfully", "", DocNum)
                objMain.objApplication.StatusBar.SetText(ItemCode + "-Item price updated successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Dim oRsSubmit4 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim UpdateQuery4 As String = "UPDATE ""@SBO_PL_C0"" SET ""U_Status""='Yes' ,""U_Remarks"" ='Price Updated Successfully'  Where ""DocEntry""='" & DocEntry & "' and ""LineId""='" & pvalrow & "'"

                oRsSubmit4.DoQuery(UpdateQuery4)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub WriteLog(barcode As String, itemCode As String, uom As String, priceList As String, newPrice As Decimal, errorCode As String, errorMessage As String, Type As String, docnum As String)
        ' Build log file path using config + today's date
        Dim logDirectory As String = ""
        Dim logFilePath As String = ""
        If Type = "ERROR" Then
            logDirectory = ConfigurationManager.AppSettings("ERRLogFile")
            logFilePath = Path.Combine(logDirectory, "" & docnum & " - Error_Log" & Today.ToString("dd-MM-yyyy") & ".csv")
            failed = logFilePath
        Else
            logDirectory = ConfigurationManager.AppSettings("LogFile")

            logFilePath = Path.Combine(logDirectory, "" & docnum & " - Sucess_Log" & Today.ToString("dd-MM-yyyy") & ".csv")
            Sucess = logFilePath
        End If

        ' Ensure log directory exists
        If Not Directory.Exists(logDirectory) Then
            Directory.CreateDirectory(logDirectory)
        End If

        ' Check if file exists and has content
        Dim writeHeader As Boolean = Not File.Exists(logFilePath) OrElse New FileInfo(logFilePath).Length = 0

        ' Open file and write header if needed
        Using sw As StreamWriter = File.AppendText(logFilePath)

            If writeHeader Then
                sw.WriteLine("Barcode,ItemCode,UOM,Price List,New Price,Error Code,Error Message,Date,Time")
            End If

            ' Format log line
            Dim currentDate As String = DateTime.Now.ToString("yyyy-MM-dd")
            Dim currentTime As String = DateTime.Now.ToString("HH:mm:ss")
            Dim logEntry As String = String.Format("""{0}"",""{1}"",""{2}"",""{3}"",""{4:F2}"",""{5}"",""{6}"",""{7}"",""{8}""",
                                                   barcode, itemCode, uom, priceList, newPrice, errorCode, errorMessage, currentDate, currentTime)
            sw.WriteLine(logEntry)
        End Using
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("PriceList", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Dim DocNum As String = oDBs_Head.GetValue("DocNum", 0)
                        Me.MarketingEmail(DocNum)

                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub




    Sub MarketingEmail(ByVal DocNum As String)
        Dim Result As String = ""
        Dim query As String = "SELECT T3.""U_UserIDE"", T3.""U_EmpIDE"", T1.""E_Mail"" AS ""User_Email"", T2.""email"" AS ""Employee_Email"" " &
                          "FROM ""@MGRCONFIG"" T0 " &
                          "INNER JOIN ""@MGRCONFIG_C2"" T3 ON T0.""DocEntry"" = T3.""DocEntry"" " &
                          "INNER JOIN ""OUSR"" T1 ON T1.""USERID"" = T3.""U_UserIDE"" " &
                          "INNER JOIN ""OHEM"" T2 ON T2.""Code"" = T3.""U_EmpIDE"" " &
                          "WHERE T0.""U_Doc"" = 'PU'"

        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs.DoQuery(query)

        Dim UserEmail As String = oRs.Fields.Item("User_Email").Value.ToString().Trim()
        Dim EmployeeEmail As String = oRs.Fields.Item("Employee_Email").Value.ToString().Trim()

        If UserEmail <> "" And EmployeeEmail <> "" Then
            Result = Me.SendEmail_TEST(UserEmail & ";" & EmployeeEmail, DocNum)
        ElseIf UserEmail <> "" Then
            Result = Me.SendEmail_TEST(UserEmail, DocNum)
        ElseIf EmployeeEmail <> "" Then
            Result = Me.SendEmail_TEST(EmployeeEmail, DocNum)
        Else
            objMain.objApplication.StatusBar.SetText("No valid email found for DocNum: " & DocNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If

        If Result = "Success" Then
            objMain.objApplication.StatusBar.SetText("Email sent successfully for " & DocNum & " !!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Else
            objMain.objApplication.StatusBar.SetText("Email sending failed for " & DocNum & " !!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If
    End Sub


    Function SendEmail_TEST(ByVal ToEmail As String, ByVal DocNum1 As String) As String
        Try
            If ToEmail = "" Then
                objMain.objUtilities.Write_Log("No Email ID for this customer")
            Else
                Dim EmailQry As String = "SELECT T0.""U_SMTPsevr"", T0.""U_SMPTPort"", T0.""U_EmailFro"", T0.""U_EmailPwd"", T0.""U_CC"", T0.""U_BCC"", T0.""U_SSL"" FROM ""@SBO_SETCONFIG""  T0"
                Dim oRsEmail As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsEmail.DoQuery(EmailQry)
                Dim ServerHost As String = oRsEmail.Fields.Item("U_SMTPsevr").Value.ToString()
                Dim PortNo As String = oRsEmail.Fields.Item("U_SMPTPort").Value.ToString()
                Dim FromMail As String = oRsEmail.Fields.Item("U_EmailFro").Value.ToString()
                Dim EmailPassword As String = oRsEmail.Fields.Item("U_EmailPwd").Value.ToString()
                ServicePointManager.SecurityProtocol = DirectCast(3072, System.Net.SecurityProtocolType) 'TLS 1.2
                Dim SMTPServer As New SmtpClient
                ' Dim Message As New MailMessage(FromMail, ToEmail)

                Dim Message As New MailMessage()
                Message.From = New MailAddress(FromMail)

                For Each email As String In ToEmail.Split(";"c)
                    If email.Trim() <> "" Then
                        Message.To.Add(email.Trim())
                    End If
                Next

                If oRsEmail.Fields.Item("U_SSL").Value = "Y" Then
                    SMTPServer.EnableSsl = True
                Else
                    SMTPServer.EnableSsl = False
                End If
                SMTPServer.Host = ServerHost
                SMTPServer.Port = PortNo
                SMTPServer.DeliveryMethod = SmtpDeliveryMethod.Network
                SMTPServer.UseDefaultCredentials = False
                SMTPServer.Credentials = New Net.NetworkCredential(FromMail, EmailPassword)

                Message.Subject = "Price Update"

                Dim Attachment As String = "Price Update is added successfully"
                Dim Base64 As String = Nothing



                Dim htmlBody As String = "<table>" &
                         "<tr><td style='font-family:Times New Roman; font-size:14px;'>Document Number : <b>" & DocNum1 & "</b></td></tr>" &
                         "<tr><td style='font-size:14px; font-weight:normal; font-family:Times New Roman;'>" & Attachment & "<br><br></td></tr>" &
                         "</table>" &
                         "<table>" &
                         "<tr><td style='font-size:14px; font-weight:normal; font-family:Times New Roman;'>This is system auto-generated mail. Please do not reply.</td></tr>" &
                         "</table>"

                Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(htmlBody, Nothing, "text/html")



                Message.AlternateViews.Add(htmlView)
                SMTPServer.Send(Message)

                Message.Dispose()
                objMain.objUtilities.Write_Log("E-Mail Sent successfully to -  " & ToEmail)
                objMain.objApplication.StatusBar.SetText("E-Mail Sent successfully to -  " & ToEmail, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Return "Success"
            End If
        Catch ex As Exception
            objMain.objUtilities.Write_Log("Error In Email Sending:" & ex.Message)
            objMain.objApplication.StatusBar.SetText("Error In Email Sending:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
    End Function
End Class
