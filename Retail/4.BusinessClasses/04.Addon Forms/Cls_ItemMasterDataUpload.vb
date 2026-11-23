
'AAkk

Imports System.CodeDom.Compiler
Imports System.Collections.Generic
Imports System.Collections.Specialized
Imports System.Configuration
Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports System.IO.Packaging
Imports System.Net
Imports System.Net.Http
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports System.Xml
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Sap.Data.Hana
Imports SAPbobsCOM
Imports SAPbouiCOM
Imports DataTable = System.Data.DataTable
''New
Public Class Cls_ItemMasterDataUpload

#Region "       Declaration             "
    Public objForm, objBarcodeForm, objNewForm, objSTRForm As SAPbouiCOM.Form
    Private oForm As SAPbouiCOM.Form
    Dim OtherCompany As SAPbobsCOM.Company = New SAPbobsCOM.Company
    Dim objComboBox As SAPbouiCOM.ComboBox
    'Dim oDS, oDS1 As SAPbouiCOM.DBDataSource
    Dim objMatrix, objMatrix1, objMatrix2, objMatrix3, objMatrix4, objMatrix5, objMatrix6 As SAPbouiCOM.Matrix
    Dim objComboBox1, objComboBox2, objComboBox3 As SAPbouiCOM.ComboBox
    Public oDS, oDS1, oDS2, oDS3, oDS4, oDS5, oDS6, oDS7 As SAPbouiCOM.DBDataSource
    Public PrgBar As SAPbouiCOM.ProgressBar
    Dim MATRIXS, sublocchk, sublocnchk As String
    Dim subloc, lvlcde1, lvlcde2 As String
    Dim j, Totalstockchk, Totalchk, mainwarehouse As Integer
    Dim oGrid As SAPbouiCOM.Grid
    Dim column, column1, column2 As SAPbouiCOM.EditTextColumn
    Dim str, str1, str2, str3, Query, str4 As String
    Public rs, RsNum As SAPbobsCOM.Recordset
    Dim LostFocusFlag As Boolean = False
    Dim objutilities As Utilities
    Public filename, Calquery As String
    Public sCmp As SAPbobsCOM.CompanyService
    Dim oDt As SAPbouiCOM.DataTable
    Dim r1, r2, r3 As RadioButton
    Dim oItem, oItemUP, oItemUP1, OBCD As SAPbobsCOM.Items
    Dim Error_Code As Integer = 1
    Dim objBarcodeMatrix As SAPbouiCOM.Matrix
    Dim StrQry, oAutoBarcode, Barnm, Var As String
    Dim oRec As SAPbobsCOM.Recordset
    Dim Startt As Int64 = 0
    Dim Endd As Integer = 0
    Private WithEvents BackgroundWorker_BulkData As New System.ComponentModel.BackgroundWorker
    Private WithEvents BackgroundWorker_GenerateDocumets As New System.ComponentModel.BackgroundWorker
    Dim dictUOM As New Dictionary(Of String, String)
    Dim baseUrl = ConfigurationManager.AppSettings("ServiceLayerBaseUrl")
    Dim UserName = ConfigurationManager.AppSettings("ServiceLayerUser")
    Dim Password = ConfigurationManager.AppSettings("ServiceLayerPassword")
    'Dim CompanyDB = ConfigurationManager.AppSettings("CompanyDB")
    Dim CompanyDB As String = objMain.objCompany.CompanyDB

#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("ItemMasterDataNew.xml", "ITEMUP", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ITEMUP", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            objutilities = New Utilities
            oDS = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP")
            oDS1 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C0")
            oDS2 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C1")
            oDS3 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C2")
            oDS4 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C3")
            oDS5 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C4")
            oDS6 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C5")
            oDS7 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C6")

            objForm.EnableMenu("1292", True)
            objForm.EnableMenu("774", False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)


            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)


            objForm.Items.Item("ATTPATH").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("ATTPATH").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("ATTPATH").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("ATTPATH").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("ATTPATH").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)


            objForm.Items.Item("DDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("STAT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("STAT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("STAT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("STAT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("DocNoIL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNoIL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNoIL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNoIL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("AppIDIL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("AppIDIL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("AppIDIL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("AppIDIL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("Item_11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("Item_13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Item_13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            'objForm.Items.Item("Item_11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("Item_11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("Preview").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Preview").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("Item_15").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            Dim Query12 As String = "select ""ItmsGrpCod"" ,""ItmsGrpNam"" from ""OITB"" "

            Query12 = "select ""ItmsGrpNam"" ,""ItmsGrpNam"" from ""OITB"" "
            Dim Qry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Qry.DoQuery(Query12)
            Dim objmatrix1 As SAPbouiCOM.Matrix = objForm.Items.Item("ItmDetMtx").Specific
            objMain.objUtilities.MatrixComboBoxValues(objmatrix1.Columns.Item("ITEMG"), Query12)
            objmatrix1.Columns.Item("ITEMG").ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objmatrix1.Columns.Item("ITEMG").DisplayDesc = True

            ' Purchase VAT (Input Tax)
            Dim qPVAT As String = "SELECT ""Code"", ""Name"" FROM ""OVTG"" WHERE ""Category"" = 'I' AND ifnull(""Inactive"",'N')='N'"
            objMain.objUtilities.MatrixComboBoxValues(objmatrix1.Columns.Item("PVAT"), qPVAT)

            ' Sales VAT (Output Tax)
            Dim qSVAT As String = "SELECT ""Code"", ""Name"" FROM ""OVTG"" WHERE ""Category"" = 'O' AND ifnull(""Inactive"",'N')='N'"
            objMain.objUtilities.MatrixComboBoxValues(objmatrix1.Columns.Item("SVAT"), qSVAT)

            ' Common display settings
            objmatrix1.Columns.Item("PVAT").ExpandType = SAPbouiCOM.BoExpandType.et_ValueDescription
            objmatrix1.Columns.Item("PVAT").DisplayDesc = True

            objmatrix1.Columns.Item("SVAT").ExpandType = SAPbouiCOM.BoExpandType.et_ValueDescription
            objmatrix1.Columns.Item("SVAT").DisplayDesc = True
            Dim objmatrixBrn As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx5").Specific
            Dim objmatrixVen As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx6").Specific
            Dim objmatrixMini As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx7").Specific

            objutilities.MatrixComboBoxValues(objmatrixBrn.Columns.Item("BrnchId"), "SELECT ""BPLId""  ,""BPLName""from ""OBPL""  union all select 0 as  ""BPLId"" ,'ALL' as""BPLName"" from Dummy")
            objutilities.MatrixComboBoxValues(objmatrixBrn.Columns.Item("PriceId"), "SELECT Top 20 ""ListNum"",""ListName"" from ""OPLN"" order by  ""ListName""asc")
            objutilities.MatrixComboBoxValues(objmatrixVen.Columns.Item("PriceId"), "SELECT Top 20 ""ListNum"",""ListName"" from ""OPLN"" order by  ""ListName""asc")
            objutilities.MatrixComboBoxValues(objmatrixMini.Columns.Item("PriceId"), "SELECT Top 20 ""ListNum"",""ListName"" from ""OPLN"" order by  ""ListName""asc")
            objutilities.MatrixComboBoxValues(objmatrix1.Columns.Item("COUNT"), "SELECT ""Code"",""Name"" from ""@VR_COUNTRY"" order by  ""Code""")





            Me.SetDefault(objForm.UniqueID)

            ''ShowFolderBrowser()
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ITEMUP" And pVal.BeforeAction = False Then
                Me.CreateForm()

            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)

            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                Me.SetNewLine(objForm.UniqueID, MATRIXS)
                '' Me.SetNewLineNewBranch(objForm.UniqueID, MATRIXS)
                ''
            ElseIf pVal.MenuUID = "Delete Row" And pVal.BeforeAction = False Then

                Dim objMatrix As SAPbouiCOM.Matrix = Nothing
                Dim oDBs_Details As SAPbouiCOM.DBDataSource = Nothing

                objForm = objMain.objApplication.Forms.GetForm("ITEMUP", objMain.objApplication.Forms.ActiveForm.TypeCount)

                ' --------------------------
                ' Identify Matrix + Datasource
                ' --------------------------
                Select Case MATRIXS

                    Case "ItmDetMtx"
                        objMatrix = CType(objForm.Items.Item("ItmDetMtx").Specific, SAPbouiCOM.Matrix)
                        oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C0")

                    Case "UOMMtx"
                        objMatrix = CType(objForm.Items.Item("UOMMtx").Specific, SAPbouiCOM.Matrix)
                        oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C1")

                    Case "BarMtx"
                        objMatrix = CType(objForm.Items.Item("BarMtx").Specific, SAPbouiCOM.Matrix)
                        oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C2")

                    Case "Mtx5"
                        objMatrix = CType(objForm.Items.Item("Mtx5").Specific, SAPbouiCOM.Matrix)
                        oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C3")

                    Case "Mtx6"
                        objMatrix = CType(objForm.Items.Item("Mtx6").Specific, SAPbouiCOM.Matrix)
                        oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C4")

                    Case "Mtx7"
                        objMatrix = CType(objForm.Items.Item("Mtx7").Specific, SAPbouiCOM.Matrix)
                        oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C5")

                End Select

                ' --------------------------
                ' Validation
                ' --------------------------
                If objMatrix Is Nothing Then Exit Sub

                If objMatrix.VisualRowCount = 0 Then
                    objMain.objApplication.MessageBox("No rows to delete.")
                    Exit Sub
                End If

                Dim rowDeleted As Boolean = False

                ' --------------------------
                ' ADD MODE
                ' --------------------------
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                    For i As Integer = objMatrix.VisualRowCount To 1 Step -1
                        If objMatrix.IsRowSelected(i) Then
                            objMatrix.DeleteRow(i)
                            rowDeleted = True
                        End If
                    Next

                    If rowDeleted Then
                        objMatrix.FlushToDataSource()
                    End If

                    ' --------------------------
                    ' UPDATE MODE
                    ' --------------------------
                ElseIf objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE OrElse objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then

                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If

                    objMatrix.FlushToDataSource()

                    For i As Integer = objMatrix.VisualRowCount To 1 Step -1
                        If objMatrix.IsRowSelected(i) Then
                            Dim dsIndex As Integer = i - 1
                            If dsIndex < oDBs_Details.Size Then
                                oDBs_Details.RemoveRecord(dsIndex)
                                rowDeleted = True
                            End If
                        End If
                    Next

                    If rowDeleted Then
                        objMatrix.LoadFromDataSource()
                    End If
                End If

                ' --------------------------
                ' No selection
                ' --------------------------
                If Not rowDeleted Then
                    objMain.objApplication.MessageBox("Please select a row to delete.")
                    Exit Sub
                End If

                ' --------------------------
                ' Re-number LineId
                ' --------------------------
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.Columns.Item("LineId").Cells.Item(i).Specific.Value = i.ToString()
                Next

            End If
            'If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
            '    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            'End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Sub StartBackgroundForLoading(Optional ByVal BackGroundWrkr As Boolean = True)
        Try
            If BackGroundWrkr = True Then
                If BackgroundWorker_BulkData.IsBusy <> True Then
                    BackgroundWorker_BulkData.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Calling Unselecting worker Error :  " & ex.Message)
        End Try
    End Sub

    Private Sub backgroundWorker1_Loading_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker_BulkData.RunWorkerCompleted
        If (e.Error IsNot Nothing) Then
            objMain.objApplication.StatusBar.SetText("upload error :" & e.Error.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        ElseIf e.Cancelled Then
            objMain.objApplication.StatusBar.SetText("Canceled", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Else
            'Me.SBO_Application.StatusBar.SetText("File uploaded successfully ..." & e.Result, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If
    End Sub

    Private Sub BackgroundWorker_Loading_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_BulkData.DoWork
        Dim DocEntry As String = objForm.Items.Item("DocNum").Specific.Value
        '' Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.GetForm("ITEMUP", objMain.objApplication.Forms.ActiveForm.TypeCount)

        'Dim FormId As SAPbouiCOM.Form = CType(objMain.objApplication.Forms.GetForm("ITEMUP", objMain.objApplication.Forms.ActiveForm.TypeCount), SAPbouiCOM.Form)
        If DocEntry <> "" Then
            Me.Item_Creation(DocEntry)
        End If

    End Sub


    Sub ShowFolderBrowser()
        Dim srcfilenamewithfullpath As String
        Dim MyTest1 As New OpenFileDialog
        With MyTest1.DefaultExt = ".xlsx"
            MyTest1.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            'MyTest1.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
            MyTest1.FilterIndex = 2
        End With
        If MyTest1.ShowDialog() = DialogResult.OK Then
            Try
                srcfilenamewithfullpath = MyTest1.FileName
            Catch ex As IO.IOException
                objMain.objApplication.MessageBox(ex.Message)
                Exit Sub
            End Try
            System.Windows.Forms.Application.ExitThread()
            Me.ItemMasterDetailsExcelUpload(srcfilenamewithfullpath)
        End If
        Exit Sub
    End Sub


    'Akhila Bandoju -----Date 18-12-24--ItemMasterDetailsExcelUpload------------------


    Sub ItemMasterDetailsExcelUpload(ByVal FilePath As String)
        Try
            'Matrix setup
            Dim objmatrix As SAPbouiCOM.Matrix = objForm.Items.Item("ItmDetMtx").Specific
            Dim Objmtx1 As SAPbouiCOM.Matrix = objForm.Items.Item("UOMMtx").Specific
            Dim Objmtx2 As SAPbouiCOM.Matrix = objForm.Items.Item("BarMtx").Specific
            Dim Objmtx3 As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx5").Specific
            Dim Objmtx4 As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx6").Specific
            Dim Objmtx5 As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx7").Specific
            oDS = Me.objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP")
            oDS1 = Me.objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C0")
            oDS2 = Me.objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C1")
            oDS3 = Me.objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C2")
            oDS4 = Me.objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C3")
            oDS5 = Me.objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C4")
            oDS6 = Me.objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C5")
            ' Declare variables to hold values from the Excel data

            Dim Keycode1, UomCode, UomUnits As String
            Dim Keycode2, Item, UomCode1, Barcodex, Catalog, BarDespn, arabicdesc, prererd, BPCatalogNumber As String
            Dim SUGARTAXAMT, excisetax, PriceX As Double
            Dim BranchId, BranchName, pricelist, pricelist1, pricelistId, Uomentry, newprice, SUGARTAX, rebate, WithoutPO, SALESTAX, PURCHASETAX As String
            Dim KeyCode, ItemName, ForeignName, CardCode, ItemGroup, SubGroup, SubSubgroup, SubSubsubGroup, Brand1, Category, ScaleItem, CountryOfOrgin, BaseUnit, ItemSize, InventoryItem, SaleItem, PurchaseItem, NOOfpcs, baseunits, UomDimesion As String
            ' Setup the connection string
            Dim dictBPL As Dictionary(Of String, String) = LoadBPLDict()
            Dim dictPriceList As Dictionary(Of String, String) = LoadPriceListDict()

            Dim connStr As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FilePath + ";Extended Properties=ImportMixedTypes=Text;Excel 12.0 Xml;HDR=YES;IMEX=1;")
            ' Open the connection
            connStr.Open()
            ' Get the sheet names from the Excel file
            Dim dtSheets As System.Data.DataTable = connStr.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
            ' List to hold sheet names
            Dim sheetName As New List(Of String)
            sheetName.Clear()
            ' Loop through the DataTable and add sheet names to the list
            For Each row As DataRow In dtSheets.Rows
                sheetName.Add(row("TABLE_NAME").ToString())
            Next
            ' Define the sheet name we want to read from
            Dim piSheetName As String = "ItemDetails$".Trim() ' Sheet names in Excel are usually suffixed with $
            ' Determine which sheet to read from
            Dim sheetToRead As String = Nothing
            If sheetName.Contains(piSheetName) Then
                sheetToRead = piSheetName
            Else
                objMain.objApplication.StatusBar.SetText("Sheet 'ItemDetails' not found in the Excel file...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                connStr.Close()
                Exit Sub
            End If
            connStr.Close()
            ' Reopen the connection for data reading
            connStr.Open()
            ' Now build the SQL query to read the data from the selected sheet
            Dim query As String = "SELECT * FROM [" & sheetToRead & "]"
            Dim adap As New OleDbDataAdapter(query, connStr)
            ' DataTable to hold the data from the sheet
            Dim dtNew As New System.Data.DataTable()
            adap.Fill(dtNew)
            Dim Itemdetails As String = ("ItemDetails$").Trim()
            Dim UOM As String = ("UOM$").Trim()
            Dim Barcode As String = ("BarCode$").Trim()
            Dim PriceupdateBranch As String = ("Branch$").Trim()
            Dim PriceupdateVendor As String = ("Vendor$").Trim()
            Dim PriceupdateMinistryprice As String = ("MinistryPrice$").Trim()

            Dim desiredSheetName As String = ""
            Dim ds As New DataSet()
            ' Iterate through the desired sheet names and try to read data
            If sheetName.Contains(Itemdetails) Then
                ' Read data from the sheet
                Dim adapItemDet As New OleDbDataAdapter("SELECT * FROM [" & Itemdetails & "]", connStr)
                'Dim ds As New DataSet()
                ' adap.Fill(ds)
                Dim ItmDetdt As New System.Data.DataTable()
                adapItemDet.Fill(ItmDetdt)
            End If
            If sheetName.Contains(UOM) Then
                ' Read data from the sheet
                Dim adapUOM As New OleDbDataAdapter("SELECT * FROM [" & UOM & "]", connStr)
                'Dim ds As New DataSet()
                ' adap.Fill(ds)
                Dim UOM1 As New System.Data.DataTable()
                adapUOM.Fill(UOM1)
            End If
            If sheetName.Contains(Barcode) Then
                ' Read data from the sheet
                Dim adapBarcode As New OleDbDataAdapter("SELECT * FROM [" & Barcode & "]", connStr)
                'Dim ds As New DataSet()
                ' adap.Fill(ds)
                Dim Barcode1 As New System.Data.DataTable()
                adapBarcode.Fill(Barcode1)
            End If

            If sheetName.Contains(PriceupdateBranch) Then
                ' Read data from the sheet
                Dim adapPriceupdateBranch As New OleDbDataAdapter("SELECT * FROM [" & PriceupdateBranch & "]", connStr)
                'Dim ds As New DataSet()
                ' adap.Fill(ds)
                Dim PriceupdateBranch1 As New System.Data.DataTable()
                adapPriceupdateBranch.Fill(PriceupdateBranch1)
            End If

            If sheetName.Contains(PriceupdateVendor) Then
                ' Read data from the sheet
                Dim adapPriceupdatevendor As New OleDbDataAdapter("SELECT * FROM [" & PriceupdateVendor & "]", connStr)
                'Dim ds As New DataSet()
                ' adap.Fill(ds)
                Dim Priceupdatevendor1 As New System.Data.DataTable()
                adapPriceupdatevendor.Fill(Priceupdatevendor1)
            End If
            If sheetName.Contains(PriceupdateMinistryprice) Then
                ' Read data from the sheet
                Dim adapPriceupdatMiniprice As New OleDbDataAdapter("SELECT * FROM [" & PriceupdateMinistryprice & "]", connStr)
                'Dim ds As New DataSet()
                ' adap.Fill(ds)
                Dim PriceupdateMinipr1 As New System.Data.DataTable()
                adapPriceupdatMiniprice.Fill(PriceupdateMinipr1)
            End If

            ' Process the data - for each row in the DataTable

            objMain.objApplication.StatusBar.SetText("Uploading Item Details..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            If sheetName.Contains("ItemDetails$") Then
                For i = 0 To dtNew.Rows.Count - 1
                    KeyCode = dtNew.Rows(i).Item(0).ToString()
                    ''ItemName = dtNew.Rows(i).Item(1).ToString()
                    ItemName = UCase(dtNew.Rows(i).Item(1).ToString())
                    ForeignName = dtNew.Rows(i).Item(2).ToString()
                    CardCode = dtNew.Rows(i).Item(3).ToString()
                    ItemGroup = dtNew.Rows(i).Item(4).ToString()
                    SubGroup = dtNew.Rows(i).Item(5).ToString()
                    SubSubgroup = dtNew.Rows(i).Item(6).ToString()

                    '-----Start -------Madhu------------------sub and sub sub group need validation at the time of item uploding
                    'Dim rsValidate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'Dim strValidateQry As String = "SELECT COUNT(*) AS ""Cnt"" " & "FROM ""@VR_SUBSUBGROUP"" " & "WHERE ""U_SUBGROUP"" = '" & SubGroup.Replace("'", "''") & "' " &
                    '                               "AND ""Name"" = '" & SubSubgroup.Replace("'", "''") & "'"

                    'rsValidate.DoQuery(strValidateQry)
                    'If rsValidate.RecordCount > 0 Then
                    '    If CInt(rsValidate.Fields.Item("Cnt").Value) = 0 Then
                    '        objMain.objApplication.StatusBar.SetText("Excel Row " & (i + 2).ToString() & " : Please select correct Sub Sub Group '" & SubSubgroup & "' for Sub Group '" & SubGroup & "'", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        Exit Sub
                    '    End If
                    'End If
                    '-----End -------Madhu------------------sub and sub sub group need validation at the time of item uploding

                    SubSubsubGroup = dtNew.Rows(i).Item(7).ToString()
                    Brand1 = dtNew.Rows(i).Item(8).ToString()
                    Category = dtNew.Rows(i).Item(9).ToString()
                    ScaleItem = dtNew.Rows(i).Item(10).ToString()
                    CountryOfOrgin = dtNew.Rows(i).Item(11).ToString()
                    BaseUnit = dtNew.Rows(i).Item(12).ToString()
                    ItemSize = dtNew.Rows(i).Item(13).ToString()
                    InventoryItem = dtNew.Rows(i).Item(14).ToString()
                    SaleItem = dtNew.Rows(i).Item(15).ToString()
                    PurchaseItem = dtNew.Rows(i).Item(16).ToString()
                    NOOfpcs = dtNew.Rows(i).Item(17).ToString()
                    'BaseUnit = dtNew.Rows(i).Item(18).ToString()
                    ''BPCatalogNumber = dtNew.Rows(i).Item(18).ToString()
                    UomDimesion = dtNew.Rows(i).Item(18).ToString()
                    SUGARTAX = dtNew.Rows(i).Item(19).ToString()
                    '   SUGARTAXAMT = dtNew.Rows(i).Item(21).ToString()
                    rebate = dtNew.Rows(i).Item(20).ToString()
                    WithoutPO = dtNew.Rows(i).Item(21).ToString()
                    SALESTAX = dtNew.Rows(i).Item(22).ToString()
                    PURCHASETAX = dtNew.Rows(i).Item(23).ToString()

                    'If i > objmatrix.RowCount Then
                    '    objmatrix.AddRow()
                    'End If
                    If objmatrix.VisualRowCount < i + 1 Then
                        objmatrix.AddRow()
                    End If
                    Dim GetCntry As String = "Select  T0.""Code"" FROM ""@VR_COUNTRY"" T0 WHERE T0.""Code"" ='" & CountryOfOrgin & "' "
                    Dim oRsgetCntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsgetCntry.DoQuery(GetCntry)

                    ' Set values into matrix
                    oDS1.SetValue("LineId", oDS1.Offset, objmatrix.VisualRowCount)
                    oDS1.SetValue("U_KEYC", oDS1.Offset, KeyCode)
                    oDS1.SetValue("U_ITEMN", oDS1.Offset, ItemName)
                    oDS1.SetValue("U_FORIE", oDS1.Offset, ForeignName)
                    oDS1.SetValue("U_CARD", oDS1.Offset, CardCode)
                    oDS1.SetValue("U_ITEMG", oDS1.Offset, ItemGroup)
                    oDS1.SetValue("U_SUBGR", oDS1.Offset, SubGroup)
                    oDS1.SetValue("U_SSGP", oDS1.Offset, SubSubgroup)
                    oDS1.SetValue("U_SUGRP", oDS1.Offset, SubSubsubGroup)
                    oDS1.SetValue("U_BRAND", oDS1.Offset, Brand1)
                    oDS1.SetValue("U_CATGO", oDS1.Offset, Category)
                    oDS1.SetValue("U_SCALE", oDS1.Offset, ScaleItem)
                    oDS1.SetValue("U_COUNT", oDS1.Offset, oRsgetCntry.Fields.Item(0).Value)
                    oDS1.SetValue("U_BUNIT", oDS1.Offset, BaseUnit)
                    oDS1.SetValue("U_ITEMS", oDS1.Offset, ItemSize)
                    oDS1.SetValue("U_INVENT", oDS1.Offset, InventoryItem)
                    oDS1.SetValue("U_SALITM", oDS1.Offset, SaleItem)
                    oDS1.SetValue("U_PURCH", oDS1.Offset, PurchaseItem)
                    oDS1.SetValue("U_NOP", oDS1.Offset, NOOfpcs)
                    'oDS1.SetValue("U_BaseUt", oDS1.Offset, BaseUnit)
                    ''oDS1.SetValue("U_CNum", oDS1.Offset, BPCatalogNumber)
                    If UomDimesion = "" Then
                        oDS1.SetValue("U_UOMD", oDS1.Offset, "-")
                    Else
                        oDS1.SetValue("U_UOMD", oDS1.Offset, UomDimesion)
                    End If
                    If UCase(SUGARTAX) = "YES" Then
                        oDS1.SetValue("U_SugarTax", oDS1.Offset, "Yes")
                    Else
                        oDS1.SetValue("U_SugarTax", oDS1.Offset, "No")
                    End If
                    ' oDS1.SetValue("U_SugarTax", oDS1.Offset, SUGARTAX)
                    '   oDS1.SetValue("U_SuTaxAmt", oDS1.Offset, SUGARTAXAMT)
                    If UCase(rebate) = "YES" Then
                        oDS1.SetValue("U_REBEGBLE", oDS1.Offset, "Yes")
                    Else
                        oDS1.SetValue("U_REBEGBLE", oDS1.Offset, "No")
                    End If
                    '  oDS1.SetValue("U_REBEGBLE", oDS1.Offset, rebate)
                    If UCase(WithoutPO) = "YES" Then
                        oDS1.SetValue("U_WITPO", oDS1.Offset, "Yes")
                    Else
                        oDS1.SetValue("U_WITPO", oDS1.Offset, "No")
                    End If

                    oDS1.SetValue("U_PVAT", oDS1.Offset, PURCHASETAX)
                    oDS1.SetValue("U_SVAT", oDS1.Offset, SALESTAX)


                    objmatrix.SetLineData(objmatrix.VisualRowCount)
                Next
                connStr.Close()
                'sheet 2 
                objMain.objApplication.StatusBar.SetText("Uploading UOM Details..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                If sheetName.Contains(UOM) Then
                    Dim adapUOM As New OleDbDataAdapter("SELECT * FROM [" & UOM & "]", connStr)
                    Dim UOM1 As New System.Data.DataTable()
                    adapUOM.Fill(UOM1)
                    ' Process the data for UOM sheet
                    objMatrix1.Clear()
                    oDS2.Clear()
                    objMatrix1.FlushToDataSource()
                    Me.SetNewLine(objForm.UniqueID, "UOMMtx")
                    For i = 0 To UOM1.Rows.Count - 1
                        Keycode1 = UOM1.Rows(i).Item(0).ToString()
                        UomCode = UOM1.Rows(i).Item(1).ToString()
                        UomUnits = UOM1.Rows(i).Item(2).ToString()
                        UomCode = GetUOMCode(UomCode)

                        'If i > Objmtx1.RowCount Then
                        '    Objmtx1.AddRow()
                        'End If
                        If Objmtx1.VisualRowCount < i + 1 Then
                            Objmtx1.AddRow()
                        End If
                        ' Set values into matrix for UOM sheet

                        oDS2.SetValue("LineId", oDS2.Offset, Objmtx1.VisualRowCount)
                        oDS2.SetValue("U_KEYCDE", oDS2.Offset, Keycode1)
                        oDS2.SetValue("U_UOMCDE", oDS2.Offset, UomCode)
                        oDS2.SetValue("U_UOMUNT", oDS2.Offset, UomUnits)
                        Objmtx1.SetLineData(Objmtx1.VisualRowCount)
                    Next
                End If
                connStr.Close()
                'Sheet Barcode 3
                objMain.objApplication.StatusBar.SetText("Uploading Barcode Details..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                If sheetName.Contains(Barcode) Then
                    Dim adapUOM As New OleDbDataAdapter("SELECT * FROM [" & Barcode & "]", connStr)
                    Dim Barcode1 As New System.Data.DataTable()
                    adapUOM.Fill(Barcode1)
                    objMatrix2.Clear()
                    oDS3.Clear()
                    objMatrix2.FlushToDataSource()
                    Me.SetNewLine(objForm.UniqueID, "BarMtx")

                    ' Process the data for UOM sheet
                    For i = 0 To Barcode1.Rows.Count - 1
                        Keycode2 = Barcode1.Rows(i).Item(0).ToString()
                        Item = Barcode1.Rows(i).Item(1).ToString()
                        UomCode1 = Barcode1.Rows(i).Item(2).ToString()
                        Barcodex = Barcode1.Rows(i).Item(3).ToString()
                        Catalog = Barcode1.Rows(i).Item(4).ToString()
                        arabicdesc = Barcode1.Rows(i).Item(5).ToString()

                        prererd = Barcode1.Rows(i).Item(6).ToString()
                        SUGARTAX = Barcode1.Rows(i).Item(7).ToString()

                        If IsDBNull(Barcode1.Rows(i).Item(8)) Then
                            SUGARTAXAMT = 0
                        Else
                            SUGARTAXAMT = CDbl(Barcode1.Rows(i).Item(8))
                        End If    '
                        ' excisetax = Barcode1.Rows(i).Item(9).ToString()
                        'madhu
                        'Dim excisetax As Double = 0

                        If Barcode1.Rows(i).Item(9) IsNot Nothing AndAlso
                          Not IsDBNull(Barcode1.Rows(i).Item(9)) AndAlso
                            Not String.IsNullOrWhiteSpace(Barcode1.Rows(i).Item(9).ToString()) Then

                            Double.TryParse(Barcode1.Rows(i).Item(9).ToString(), excisetax)

                        End If

                        'If i > Objmtx2.RowCount Then
                        '    Objmtx2.AddRow()
                        'End If
                        Item = GetUOMCode(Item)

                        If Objmtx2.VisualRowCount < i + 1 Then
                            Objmtx2.AddRow()
                        End If
                        oDS3.SetValue("LineId", oDS3.Offset, Objmtx2.VisualRowCount)
                        oDS3.SetValue("U_KEYC", oDS3.Offset, Keycode2)
                        ' oDS3.SetValue("U_ITEMN", oDS3.Offset, Item)Item
                        oDS3.SetValue("U_UOMCDE", oDS3.Offset, Item)
                        oDS3.SetValue("U_BarCode", oDS3.Offset, UomCode1)
                        oDS3.SetValue("U_BDesc", oDS3.Offset, Barcodex)
                        oDS3.SetValue("U_CNum", oDS3.Offset, Catalog)
                        oDS3.SetValue("U_ADESC", oDS3.Offset, arabicdesc)
                        oDS3.SetValue("U_PREFER", oDS3.Offset, prererd)
                        '
                        If UCase(SUGARTAX) = "YES" Then
                            oDS3.SetValue("U_STAX", oDS3.Offset, "Yes")
                        Else
                            oDS3.SetValue("U_STAX", oDS3.Offset, "No")
                        End If

                        ' oDS3.SetValue("U_STAX", oDS3.Offset, SUGARTAX)

                        oDS3.SetValue("U_SUTAXAMT", oDS3.Offset, SUGARTAXAMT)
                        oDS3.SetValue("U_ETAX", oDS3.Offset, excisetax)
                        oDS3.SetValue("U_Margin", oDS3.Offset, 0)
                        oDS3.SetValue("U_Markup", oDS3.Offset, 0)


                        Objmtx2.SetLineData(Objmtx2.VisualRowCount)
                    Next
                End If
                connStr.Close()
                objMain.objApplication.StatusBar.SetText("Uploading 
Price  Update Details..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                'sheet4
                If sheetName.Contains(PriceupdateBranch) Then
                    Dim adapPriceupdateBranch As New OleDbDataAdapter("SELECT * FROM [" & PriceupdateBranch & "]", connStr)
                    Dim PriceupdateBranch1 As New System.Data.DataTable()
                    adapPriceupdateBranch.Fill(PriceupdateBranch1)
                    objMatrix3.Clear()
                    oDS4.Clear()
                    objMatrix3.FlushToDataSource()
                    Me.SetNewLine(objForm.UniqueID, "Mtx5")
                    '' Me.SetNewLineNewBranch(objForm.UniqueID, MATRIXS)

                    ' Process the data for Branch
                    For i = 0 To PriceupdateBranch1.Rows.Count - 1
                        Keycode1 = PriceupdateBranch1.Rows(i).Item(0).ToString()
                        BranchId = PriceupdateBranch1.Rows(i).Item(1).ToString()
                        '   BranchName = PriceupdateBranch1.Rows(i).Item(2).ToString()
                        BranchName = If(Not String.IsNullOrEmpty(BranchName), BranchName, If(dictBPL.ContainsKey(BranchId), dictBPL(BranchId), ""))
                        pricelistId = PriceupdateBranch1.Rows(i).Item(3).ToString()
                        '  pricelist = PriceupdateBranch1.Rows(i).Item(4).ToString()
                        pricelist = If(Not String.IsNullOrEmpty(pricelist), pricelist, If(dictPriceList.ContainsKey(pricelistId), dictPriceList(pricelistId), ""))
                        Uomentry = PriceupdateBranch1.Rows(i).Item(5).ToString()
                        UomCode = PriceupdateBranch1.Rows(i).Item(6).ToString()
                        Dim price As String = PriceupdateBranch1.Rows(i).Item(7).ToString()
                        Dim Cost As String = PriceupdateBranch1.Rows(i).Item(8).ToString()
                        Dim pricelistID1 As String = PriceupdateBranch1.Rows(i).Item(9).ToString()
                        ''pricelist1 = If(Not String.IsNullOrEmpty(pricelist1), pricelist1, If(dictPriceList.ContainsKey(pricelistID1), dictPriceList(pricelistID1), ""))


                        UomCode = GetUOMCode(UomCode)
                        Dim oRsUOM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        If Not String.IsNullOrEmpty(UomCode) Then

                            Dim qry As String = "SELECT ""UomEntry"" FROM OUOM WHERE UPPER(""UomCode"")='" & UomCode.Trim().ToUpper() & "'"
                            oRsUOM.DoQuery(qry)
                            If oRsUOM.RecordCount > 0 Then
                                Uomentry = oRsUOM.Fields.Item("UomEntry").Value.ToString()
                            End If

                        End If
                        If Objmtx3.VisualRowCount < i + 1 Then
                            Objmtx3.AddRow()
                        End If
                        ' Set values into matrix for UOM sheet
                        oDS4.SetValue("LineId", oDS4.Offset, Objmtx3.VisualRowCount)
                        oDS4.SetValue("U_ItemC", oDS4.Offset, Keycode1)
                        oDS4.SetValue("U_BrnchId", oDS4.Offset, BranchId)
                        oDS4.SetValue("U_BrachNm", oDS4.Offset, BranchName)
                        oDS4.SetValue("U_PriceLt", oDS4.Offset, pricelist)
                        oDS4.SetValue("U_PriceId", oDS4.Offset, pricelistId)
                        oDS4.SetValue("U_UomEtry", oDS4.Offset, Uomentry)
                        oDS4.SetValue("U_UOMCode", oDS4.Offset, UomCode)
                        oDS4.SetValue("U_OldPrice", oDS4.Offset, price)
                        oDS4.SetValue("U_Cost", oDS4.Offset, Cost)
                        oDS4.SetValue("U_PriceList", oDS4.Offset, pricelistID1)
                        oDS4.SetValue("U_NewPrice", oDS4.Offset, "")
                        oDS4.SetValue("U_Update", oDS4.Offset, "20211101")
                        oDS4.SetValue("U_EndDate", oDS4.Offset, "20401231")
                        oDS4.SetValue("U_UserName", oDS4.Offset, "")
                        Objmtx3.SetLineData(Objmtx3.VisualRowCount)
                    Next
                End If
                connStr.Close()
                'sheet5
                objMain.objApplication.StatusBar.SetText("Uploading Price Update Details..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                If sheetName.Contains(PriceupdateVendor) Then
                    Dim adapPriceupdatevendor As New OleDbDataAdapter("SELECT * FROM [" & PriceupdateVendor & "]", connStr)
                    Dim Priceupdatevendor1 As New System.Data.DataTable()
                    adapPriceupdatevendor.Fill(Priceupdatevendor1)
                    objMatrix4.Clear()
                    oDS5.Clear()
                    objMatrix4.FlushToDataSource()
                    Me.SetNewLine(objForm.UniqueID, "Mtx6")

                    For i = 0 To Priceupdatevendor1.Rows.Count - 1
                        Keycode1 = Priceupdatevendor1.Rows(i).Item(0).ToString()
                        CardCode = Priceupdatevendor1.Rows(i).Item(1).ToString()
                        pricelistId = Priceupdatevendor1.Rows(i).Item(2).ToString()
                        pricelist = If(pricelist.Trim() <> "", pricelist, If(dictPriceList.ContainsKey(pricelistId), dictPriceList(pricelistId), ""))
                        ' pricelist = Priceupdatevendor1.Rows(i).Item(3).ToString()
                        Uomentry = Priceupdatevendor1.Rows(i).Item(4).ToString()
                        UomCode = Priceupdatevendor1.Rows(i).Item(5).ToString()
                        UomCode = GetUOMCode(UomCode)

                        Dim price As String = Priceupdatevendor1.Rows(i).Item(6).ToString()
                        If Objmtx4.VisualRowCount < i + 1 Then
                            Objmtx4.AddRow()
                        End If
                        ' Set values into matrix for UOM sheet
                        oDS5.SetValue("LineId", oDS5.Offset, Objmtx4.VisualRowCount)
                        oDS5.SetValue("U_Code", oDS5.Offset, Keycode1)
                        oDS5.SetValue("U_CCode", oDS5.Offset, CardCode)
                        oDS5.SetValue("U_PriceId", oDS5.Offset, pricelistId)
                        oDS5.SetValue("U_PriceLt", oDS5.Offset, pricelist)
                        oDS5.SetValue("U_UEntry", oDS5.Offset, Uomentry)
                        oDS5.SetValue("U_UOMCode", oDS5.Offset, UomCode)
                        oDS5.SetValue("U_Price", oDS5.Offset, price)
                        Objmtx4.SetLineData(Objmtx4.VisualRowCount)
                    Next
                End If
                connStr.Close()
                'sheet6
                objMain.objApplication.StatusBar.SetText("Uploading Ministry  Details..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)


                If sheetName.Contains(PriceupdateMinistryprice) Then
                    Dim adapPriceupdatMiniprice As New OleDbDataAdapter("SELECT * FROM [" & PriceupdateMinistryprice & "]", connStr)
                    Dim PriceupdateMinipr1 As New System.Data.DataTable()
                    adapPriceupdatMiniprice.Fill(PriceupdateMinipr1)
                    objMatrix5.Clear()
                    oDS6.Clear()
                    objMatrix5.FlushToDataSource()
                    Me.SetNewLine(objForm.UniqueID, "Mtx7")

                    For i = 0 To PriceupdateMinipr1.Rows.Count - 1
                        Keycode1 = PriceupdateMinipr1.Rows(i).Item(0).ToString()
                        BranchId = PriceupdateMinipr1.Rows(i).Item(1).ToString()
                        BranchName = PriceupdateMinipr1.Rows(i).Item(2).ToString()
                        pricelistId = PriceupdateMinipr1.Rows(i).Item(3).ToString()
                        pricelist = PriceupdateMinipr1.Rows(i).Item(4).ToString()
                        Uomentry = PriceupdateMinipr1.Rows(i).Item(5).ToString()
                        UomCode = PriceupdateMinipr1.Rows(i).Item(6).ToString()
                        Dim price As String = PriceupdateMinipr1.Rows(i).Item(7).ToString()
                        BranchName = If(BranchName.Trim() <> "", BranchName, If(dictBPL.ContainsKey(BranchId), dictBPL(BranchId), ""))

                        pricelist = If(pricelist.Trim() <> "", pricelist, If(dictPriceList.ContainsKey(pricelistId), dictPriceList(pricelistId), ""))

                        If Objmtx5.VisualRowCount < i + 1 Then
                            Objmtx5.AddRow()
                        End If
                        UomCode = GetUOMCode(UomCode)

                        ' Set values into matrix for UOM sheet
                        oDS6.SetValue("LineId", oDS6.Offset, Objmtx5.VisualRowCount)
                        oDS6.SetValue("U_Code", oDS6.Offset, Keycode1)
                        oDS6.SetValue("U_BrnchId", oDS6.Offset, BranchId)
                        oDS6.SetValue("U_BrachNm", oDS6.Offset, BranchName)
                        oDS6.SetValue("U_PriceLt", oDS6.Offset, pricelist)
                        oDS6.SetValue("U_PriceId", oDS6.Offset, pricelistId)
                        oDS6.SetValue("U_UomEtry", oDS6.Offset, Uomentry)
                        oDS6.SetValue("U_UOMCode", oDS6.Offset, UomCode)
                        oDS6.SetValue("U_Price", oDS6.Offset, price)
                        Objmtx5.SetLineData(Objmtx5.VisualRowCount)
                    Next
                End If
                connStr.Close()
            End If
            Me.LoadMatrix_BPLID()
            Call CalculateMarginAndMarkup_Fast()

            objMain.objApplication.StatusBar.SetText("Uploading Completed..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            objForm.Items.Item("Item_11").Enabled = True
        Finally
            objForm.Items.Item("Item_11").Enabled = True
            ' Any cleanup code if necessary
        End Try
    End Sub



    Private Function LoadUOMDict() As Dictionary(Of String, String)

        If dictUOM IsNot Nothing Then Return dictUOM ' already loaded

        dictUOM = New Dictionary(Of String, String)

        Dim rsUOM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rsUOM.DoQuery("SELECT ""UomCode"", ""UomName"" FROM OUOM")

        While Not rsUOM.EoF
            Dim code As String = rsUOM.Fields.Item("UomCode").Value.ToString().ToUpper().Trim()
            Dim name As String = rsUOM.Fields.Item("UomName").Value.ToString().ToUpper().Trim()

            ' exact
            dictUOM(code) = code
            dictUOM(name) = code

            ' remove numbers (BAG1 → BAG)
            Dim clean As String = System.Text.RegularExpressions.Regex.Replace(code, "\d", "")
            dictUOM(clean) = code

            rsUOM.MoveNext()
        End While

        Return dictUOM

    End Function




    '========================
    ' 🔥 FUNCTION: GET CORRECT UOM
    '========================
    Private Function GetUOMCode(ByVal inputUOM As String) As String

        If String.IsNullOrWhiteSpace(inputUOM) Then Return ""

        Dim dict = LoadUOMDict()

        Dim key As String = inputUOM.ToUpper().Trim()

        ' 1. Exact match
        If dict.ContainsKey(key) Then Return dict(key)

        ' 2. Remove numbers (BG1 → BG)
        Dim cleaned As String = System.Text.RegularExpressions.Regex.Replace(key, "\d", "")
        If dict.ContainsKey(cleaned) Then Return dict(cleaned)

        ' 3. StartsWith match (BG → BAG1)
        For Each kvp In dict
            If kvp.Key.StartsWith(cleaned) Then
                Return kvp.Value
            End If
        Next

        '========================
        ' 🔥 SQL FALLBACK (OPTIONAL BUT SAFE)
        '========================
        Try
            Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rs.DoQuery("SELECT ""TNX_GET_UOM_CODE""('" & key & "') FROM DUMMY")

            Dim dbUOM As String = rs.Fields.Item(0).Value.ToString()
            If Not String.IsNullOrEmpty(dbUOM) Then Return dbUOM
        Catch
        End Try

        ' 4. final fallback
        Return inputUOM

    End Function





    ' 🔥 BPL Dictionary
    Private Function LoadBPLDict() As Dictionary(Of String, String)

        Dim dict As New Dictionary(Of String, String)
        Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        rs.DoQuery("SELECT ""BPLId"", ""BPLName"" FROM OBPL")

        While Not rs.EoF
            dict(rs.Fields.Item("BPLId").Value.ToString()) = rs.Fields.Item("BPLName").Value.ToString()
            rs.MoveNext()
        End While

        Return dict
    End Function


    ' 🔥 Price List Dictionary
    Private Function LoadPriceListDict() As Dictionary(Of String, String)

        Dim dict As New Dictionary(Of String, String)
        Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        rs.DoQuery("SELECT ""ListNum"", ""ListName"" FROM OPLN")

        While Not rs.EoF
            dict(rs.Fields.Item("ListNum").Value.ToString()) = rs.Fields.Item("ListName").Value.ToString()
            rs.MoveNext()
        End While

        Return dict
    End Function
    Private Sub CalculateMarginAndMarkup_Fast1()

        Try
            Dim objmatrix As SAPbouiCOM.Matrix = objForm.Items.Item("ItmDetMtx").Specific
            Dim Objmtx3 As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx5").Specific

            '================= STEP 1: BUILD DICTIONARY =================
            Dim PriceDict As New Dictionary(Of String, Double)

            For j As Integer = 1 To Objmtx3.VisualRowCount

                Objmtx3.GetLineData(j)

                Dim Key As String = oDS4.GetValue("U_ItemC", oDS4.Offset).Trim()
                Dim priceStr As String = oDS4.GetValue("U_OldPrice", oDS4.Offset).Trim()

                If Key <> "" AndAlso priceStr <> "" Then
                    If Not PriceDict.ContainsKey(Key) Then
                        PriceDict.Add(Key, CDbl(priceStr))
                    End If
                End If

            Next

            '================= STEP 2: LOOP ITEM MATRIX =================
            For i As Integer = 1 To objmatrix.VisualRowCount

                objmatrix.GetLineData(i)

                Dim KeyCode As String = oDS1.GetValue("U_KEYC", oDS1.Offset).Trim()
                Dim PurchaseVATCode As String = oDS1.GetValue("U_PVAT", oDS1.Offset).Trim()
                Dim SugarTaxAmtStr As String = oDS1.GetValue("U_SuTaxAmt", oDS1.Offset).Trim()

                Dim RetailPriceWithVAT As Double = 0
                Dim LastPurchaseWithVAT As Double = 0
                Dim Margin As Double = 0
                Dim Markup As Double = 0
                Dim VatPercent As Double = 0
                Dim SugarAmt As Double = 0

                '--- Sugar Tax ---
                If SugarTaxAmtStr <> "" Then
                    SugarAmt = CDbl(SugarTaxAmtStr)
                End If

                '--- VAT from OVTG ---
                If PurchaseVATCode <> "" Then
                    Dim oRsVAT As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsVAT.DoQuery("SELECT ""Rate"" FROM OVTG WHERE ""Code"" = '" & PurchaseVATCode & "'")

                    If oRsVAT.RecordCount > 0 Then
                        VatPercent = CDbl(oRsVAT.Fields.Item("Rate").Value)
                    End If
                End If

                '--- GET PRICE FROM DICTIONARY (NO LOOP) ---
                If PriceDict.ContainsKey(KeyCode) Then
                    RetailPriceWithVAT = PriceDict(KeyCode)
                End If

                '--- CALCULATIONS ---
                LastPurchaseWithVAT = SugarAmt + (SugarAmt * VatPercent / 100)

                If RetailPriceWithVAT > 0 Then
                    Margin = ((RetailPriceWithVAT - LastPurchaseWithVAT) / RetailPriceWithVAT) * 100
                End If

                If LastPurchaseWithVAT > 0 Then
                    Markup = ((RetailPriceWithVAT - LastPurchaseWithVAT) / LastPurchaseWithVAT) * 100
                End If
                objmatrix.GetLineData(i)
                '--- SET VALUES ---
                oDS1.SetValue("U_Margin", oDS1.Offset, Margin.ToString())
                oDS1.SetValue("U_Markup", oDS1.Offset, Markup.ToString())

                objmatrix.SetLineData(i)

            Next

        Catch ex As Exception
            objForm.Items.Item("Item_11").Enabled = True
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    'madhu backup 18/04/2025
    Private Sub CalculateMarginAndMarkup_Fastbacku()

        Try
            Dim objmatrix As SAPbouiCOM.Matrix = objForm.Items.Item("BarMtx").Specific
            Dim Objmtx3 As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx5").Specific
            Dim objMatrixItem As SAPbouiCOM.Matrix = objForm.Items.Item("ItmDetMtx").Specific

            oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C0") ' Item Master
            oDS3 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C2") ' Barcode Matrix
            oDS4 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C3") ' Price Matrix

            Dim VatCodeMap As New Dictionary(Of String, String)
            Dim ItemSugarMap As New Dictionary(Of String, String)
            Dim PriceDict As New Dictionary(Of String, Double)
            Dim VatDict As New Dictionary(Of String, Double)

            Const PURCHASE_LIST As String = "11"
            Const RETAIL_LIST As String = "1"

            '================= STEP 1: ITEM → VAT + SUGAR =================
            For x As Integer = 1 To objMatrixItem.VisualRowCount
                objMatrixItem.GetLineData(x)

                Dim key As String = oDS1.GetValue("U_KEYC", oDS1.Offset).Trim()
                Dim vat As String = oDS1.GetValue("U_PVAT", oDS1.Offset).Trim()
                Dim sugar As String = oDS1.GetValue("U_SugarTax", oDS1.Offset).Trim().ToUpper()

                If key <> "" Then

                    If Not VatCodeMap.ContainsKey(key) Then
                        VatCodeMap.Add(key, vat)
                    End If

                    If Not ItemSugarMap.ContainsKey(key) Then
                        ItemSugarMap.Add(key, sugar)
                    End If

                End If
            Next

            '================= STEP 2: PRICE DICTIONARY =================
            For j As Integer = 1 To Objmtx3.VisualRowCount
                Objmtx3.GetLineData(j)

                Dim itemCode As String = oDS4.GetValue("U_ItemC", oDS4.Offset).Trim()
                Dim uom As String = oDS4.GetValue("U_UOMCODE", oDS4.Offset).Trim()
                Dim priceList As String = oDS4.GetValue("U_PriceId", oDS4.Offset).Trim()

                Dim priceVal As Double
                If Double.TryParse(oDS4.GetValue("U_OldPrice", oDS4.Offset), priceVal) Then

                    Dim key As String = itemCode & "|" & uom & "|" & priceList

                    If Not PriceDict.ContainsKey(key) Then
                        PriceDict.Add(key, priceVal)
                    End If

                End If
            Next

            '================= STEP 3: VAT =================
            Dim oRsVAT As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsVAT.DoQuery("SELECT ""Code"", ""Rate"" FROM OVTG")

            While Not oRsVAT.EoF
                VatDict(oRsVAT.Fields.Item("Code").Value.ToString()) = CDbl(oRsVAT.Fields.Item("Rate").Value)
                oRsVAT.MoveNext()
            End While
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRsVAT)

            '================= STEP 4: MAIN LOOP =================
            For i As Integer = 1 To objmatrix.VisualRowCount
                If i = 1 Then
                    Dim test As String = "0"
                End If
                objmatrix.GetLineData(i)

                Dim itemCode As String = oDS3.GetValue("U_KEYC", oDS3.Offset).Trim()
                Dim uom As String = oDS3.GetValue("U_UOMCDE", oDS3.Offset).Trim()
                Dim sugarFlag As String = oDS3.GetValue("U_STAX", oDS3.Offset).Trim().ToUpper()

                ' 👉 Get Item-level sugar flag
                Dim itemSugarFlag As String = ""
                If ItemSugarMap.ContainsKey(itemCode) Then
                    itemSugarFlag = ItemSugarMap(itemCode)
                End If

                ' 👉 Get VAT
                Dim vatCode As String = ""
                If VatCodeMap.ContainsKey(itemCode) Then
                    vatCode = VatCodeMap(itemCode)
                End If

                Dim purchasePrice As Decimal = 0D
                Dim retailPrice As Decimal = 0D
                Dim vatPercent As Decimal = 0D
                vatPercent = If(VatDict.ContainsKey(vatCode), VatDict(vatCode), 0)

                '================= 🔥 BOTH CONDITIONS =================
                If itemSugarFlag = "YES" AndAlso sugarFlag = "YES" Then

                    Dim sugarTax As Decimal = 0D
                    Decimal.TryParse(oDS3.GetValue("U_SUTAXAMT", oDS3.Offset), sugarTax)


                    sugarTax = Math.Round(sugarTax, 3)

                    Dim exciseTax As Decimal = Math.Round(sugarTax * 0.5D, 3)
                    purchasePrice = Math.Round(sugarTax + exciseTax, 3)

                    oDS3.SetValue("U_ETAX", oDS3.Offset, exciseTax.ToString())
                    oDS3.SetValue("U_Price", oDS3.Offset, purchasePrice.ToString())

                Else
                    Dim purchaseKey As String = itemCode & "|" & uom & "|" & PURCHASE_LIST
                    If PriceDict.ContainsKey(purchaseKey) Then
                        purchasePrice = PriceDict(purchaseKey)
                    End If
                End If

                '================= RETAIL =================
                Dim retailKey As String = itemCode & "|" & uom & "|" & RETAIL_LIST
                If PriceDict.ContainsKey(retailKey) Then
                    retailPrice = PriceDict(retailKey)
                End If
                purchasePrice = Math.Round(purchasePrice, 3)
                retailPrice = Math.Round(retailPrice, 3)
                vatPercent = Math.Round(vatPercent, 2)

                '================= VAT CALCULATION =================
                Dim vatAmount As Decimal = Math.Round(purchasePrice * vatPercent / 100D, 5)
                Dim purchaseWithVAT As Decimal = Math.Round(purchasePrice + vatAmount, 3)

                '================= CALCULATION =================
                Dim margin As Double = 0
                Dim markup As Double = 0

                If retailPrice > 0 Then
                    margin = Math.Round(((retailPrice - purchaseWithVAT) / retailPrice) * 100, 3)
                End If

                If purchaseWithVAT > 0 Then
                    markup = Math.Round(((retailPrice - purchaseWithVAT) / purchaseWithVAT) * 100, 3)
                End If

                '================= VAT =================
                'purchasePrice = Math.Round(purchasePrice, 3)

                'Dim purchaseWithVAT As Double = Math.Round(purchasePrice * (1 + vatPercent / 100), 5)
                '' Dim purchaseWithVAT As Double = purchasePrice * (1 + vatPercent / 100)

                ''================= CALCULATION =================
                'Dim margin As Double = 0
                'Dim markup As Double = 0

                'If retailPrice > 0 Then
                '    margin = ((retailPrice - purchaseWithVAT) / retailPrice) * 100
                'End If

                'If purchaseWithVAT > 0 Then
                '    markup = ((retailPrice - purchaseWithVAT) / purchaseWithVAT) * 100
                'End If

                '================= SAVE =================
                oDS3.SetValue("U_Margin", oDS3.Offset, Math.Round(margin, 3).ToString())
                oDS3.SetValue("U_Markup", oDS3.Offset, Math.Round(markup, 3).ToString())

                objmatrix.SetLineData(i)

            Next

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub


    Private Sub CalculateMarginAndMarkup_Fast()

        Try
            Dim objmatrix As SAPbouiCOM.Matrix = objForm.Items.Item("BarMtx").Specific
            Dim Objmtx3 As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx5").Specific
            Dim objMatrixItem As SAPbouiCOM.Matrix = objForm.Items.Item("ItmDetMtx").Specific

            oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C0") ' Item Master
            oDS3 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C2") ' Barcode Matrix
            oDS4 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C3") ' Price Matrix

            Dim VatCodeMap As New Dictionary(Of String, String)
            Dim ItemSugarMap As New Dictionary(Of String, String)
            Dim PriceDict As New Dictionary(Of String, Double)
            Dim VatDict As New Dictionary(Of String, Double)

            Const PURCHASE_LIST As String = "11"
            Const RETAIL_LIST As String = "1"

            '================= STEP 1: ITEM → VAT + SUGAR =================
            For x As Integer = 1 To objMatrixItem.VisualRowCount
                objMatrixItem.GetLineData(x)

                Dim key As String = oDS1.GetValue("U_KEYC", oDS1.Offset).Trim()
                Dim vat As String = oDS1.GetValue("U_PVAT", oDS1.Offset).Trim()
                Dim sugar As String = oDS1.GetValue("U_SugarTax", oDS1.Offset).Trim().ToUpper()

                If key <> "" Then

                    If Not VatCodeMap.ContainsKey(key) Then
                        VatCodeMap.Add(key, vat)
                    End If

                    If Not ItemSugarMap.ContainsKey(key) Then
                        ItemSugarMap.Add(key, sugar)
                    End If

                End If
            Next

            '================= STEP 2: PRICE DICTIONARY =================
            For j As Integer = 1 To Objmtx3.VisualRowCount
                Objmtx3.GetLineData(j)

                Dim itemCode As String = oDS4.GetValue("U_ItemC", oDS4.Offset).Trim()
                Dim uom As String = oDS4.GetValue("U_UOMCODE", oDS4.Offset).Trim()
                'Dim priceList As String = oDS4.GetValue("U_PriceId", oDS4.Offset).Trim()

                'Dim priceVal As Double
                'If Double.TryParse(oDS4.GetValue("U_OldPrice", oDS4.Offset), priceVal) Then

                '    Dim key As String = itemCode & "|" & uom & "|" & priceList

                '    If Not PriceDict.ContainsKey(key) Then
                '        PriceDict.Add(key, priceVal)
                '    End If

                'End If

                ' -------- Price List 1 --------
                Dim priceVal1 As Double
                If Double.TryParse(oDS4.GetValue("U_OldPrice", oDS4.Offset), priceVal1) Then

                    Dim key1 As String = itemCode & "|" & uom & "|1"

                    If Not PriceDict.ContainsKey(key1) Then
                        PriceDict.Add(key1, priceVal1)
                    End If

                End If

                ' -------- Price List 11 --------
                Dim priceVal11 As Double
                If Double.TryParse(oDS4.GetValue("U_Cost", oDS4.Offset), priceVal11) Then

                    Dim key11 As String = itemCode & "|" & uom & "|11"

                    If Not PriceDict.ContainsKey(key11) Then
                        PriceDict.Add(key11, priceVal11)
                    End If

                End If

            Next

            '================= STEP 3: VAT =================
            Dim oRsVAT As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsVAT.DoQuery("SELECT ""Code"", ""Rate"" FROM OVTG")

            While Not oRsVAT.EoF
                VatDict(oRsVAT.Fields.Item("Code").Value.ToString()) = CDbl(oRsVAT.Fields.Item("Rate").Value)
                oRsVAT.MoveNext()
            End While
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRsVAT)

            '==================For getting item Group ====================
            Dim ItemGroupMap As New Dictionary(Of String, String)

            'FIRST TAB MATRIX
            Dim SubSubGroupMap As New Dictionary(Of String, String)

            'FIRST TAB
            For i As Integer = 1 To objMatrixItem.VisualRowCount

                objMatrixItem.GetLineData(i)

                Dim KeyCode As String = oDS1.GetValue("U_KEYC", oDS1.Offset).Trim()
                Dim SubSubGroup As String = oDS1.GetValue("U_SSGP", oDS1.Offset).Trim()

                If Not SubSubGroupMap.ContainsKey(KeyCode) Then
                    SubSubGroupMap.Add(KeyCode, SubSubGroup)
                End If

            Next

            '================= STEP 4: MAIN LOOP =================
            For i As Integer = 1 To objmatrix.VisualRowCount
                If i = 1 Then
                    Dim test As String = "0"
                End If
                objmatrix.GetLineData(i)

                Dim itemCode As String = oDS3.GetValue("U_KEYC", oDS3.Offset).Trim()
                Dim uom As String = oDS3.GetValue("U_UOMCDE", oDS3.Offset).Trim()
                Dim sugarFlag As String = oDS3.GetValue("U_STAX", oDS3.Offset).Trim().ToUpper()

                ' 👉 Get Item-level sugar flag
                Dim itemSugarFlag As String = ""
                If ItemSugarMap.ContainsKey(itemCode) Then
                    itemSugarFlag = ItemSugarMap(itemCode)
                End If

                ' 👉 Get VAT
                Dim vatCode As String = ""
                If VatCodeMap.ContainsKey(itemCode) Then
                    vatCode = VatCodeMap(itemCode)
                End If

                Dim purchasePrice As Decimal = 0D
                Dim retailPrice As Decimal = 0D
                Dim vatPercent As Decimal = 0D
                vatPercent = If(VatDict.ContainsKey(vatCode), VatDict(vatCode), 0)

                '================= 🔥 BOTH CONDITIONS =================
                If itemSugarFlag = "YES" AndAlso sugarFlag = "YES" Then

                    Dim sugarTax As Decimal = 0D
                    Decimal.TryParse(oDS3.GetValue("U_SUTAXAMT", oDS3.Offset), sugarTax)

                    sugarTax = Math.Round(sugarTax, 3)

                    Dim exciseTax As Decimal = 0D
                    Decimal.TryParse(oDS3.GetValue("U_ETAX", oDS3.Offset), exciseTax)

                    exciseTax = Math.Round(exciseTax, 3)

                    ''Dim exciseTax As Decimal = Math.Round(sugarTax * 0.5D, 3)
                    purchasePrice = Math.Round(sugarTax + exciseTax, 3)

                    ''oDS3.SetValue("U_ETAX", oDS3.Offset, exciseTax.ToString())
                    oDS3.SetValue("U_Price", oDS3.Offset, purchasePrice.ToString())

                Else
                    Dim purchaseKey As String = itemCode & "|" & uom & "|" & PURCHASE_LIST
                    If PriceDict.ContainsKey(purchaseKey) Then
                        purchasePrice = PriceDict(purchaseKey)
                    End If
                End If

                '================= RETAIL =================
                Dim retailKey As String = itemCode & "|" & uom & "|" & RETAIL_LIST
                If PriceDict.ContainsKey(retailKey) Then
                    retailPrice = PriceDict(retailKey)
                End If
                purchasePrice = Math.Round(purchasePrice, 3)
                retailPrice = Math.Round(retailPrice, 3)
                vatPercent = Math.Round(vatPercent, 2)

                '================= VAT CALCULATION =================
                Dim vatAmount As Decimal = Math.Round(purchasePrice * vatPercent / 100D, 5)
                Dim purchaseWithVAT As Decimal = Math.Round(purchasePrice + vatAmount, 3)

                '================= CALCULATION =================
                Dim margin As Double = 0
                Dim markup As Double = 0

                If retailPrice > 0 Then
                    margin = Math.Round(((retailPrice - purchaseWithVAT) / retailPrice) * 100, 3)
                End If

                If purchaseWithVAT > 0 Then
                    markup = Math.Round(((retailPrice - purchaseWithVAT) / purchaseWithVAT) * 100, 3)
                End If

                '================= VAT =================
                'purchasePrice = Math.Round(purchasePrice, 3)

                'Dim purchaseWithVAT As Double = Math.Round(purchasePrice * (1 + vatPercent / 100), 5)
                '' Dim purchaseWithVAT As Double = purchasePrice * (1 + vatPercent / 100)

                ''================= CALCULATION =================
                'Dim margin As Double = 0
                'Dim markup As Double = 0

                'If retailPrice > 0 Then
                '    margin = ((retailPrice - purchaseWithVAT) / retailPrice) * 100
                'End If

                'If purchaseWithVAT > 0 Then
                '    markup = ((retailPrice - purchaseWithVAT) / purchaseWithVAT) * 100
                'End If

                '================= SAVE =================
                oDS3.SetValue("U_Margin", oDS3.Offset, Math.Round(margin, 3).ToString())
                oDS3.SetValue("U_Markup", oDS3.Offset, Math.Round(markup, 3).ToString())


                '===================Remarks getting ===============================

                Dim MarginMap As New Dictionary(Of String, Decimal)

                Dim rsMargin As SAPbobsCOM.Recordset
                rsMargin = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                Dim qry As String = "SELECT ""Name"", ""U_Margin"" FROM ""@VR_SUBSUBGROUP"""

                rsMargin.DoQuery(qry)

                While rsMargin.EoF = False

                    Dim SubSubGroupNew As String = rsMargin.Fields.Item("Name").Value.ToString().Trim()
                    Dim MarginREM As Decimal = 0D

                    Decimal.TryParse(rsMargin.Fields.Item("U_Margin").Value.ToString(), MarginREM)

                    If Not MarginMap.ContainsKey(SubSubGroupNew) Then
                        MarginMap.Add(SubSubGroupNew, MarginREM)
                    End If

                    rsMargin.MoveNext()

                End While

                Dim SubSubGroup As String = ""
                Dim AllowedMargin As Decimal = 0D

                If SubSubGroupMap.ContainsKey(itemCode) Then
                    SubSubGroup = SubSubGroupMap(itemCode)
                End If

                If MarginMap.ContainsKey(SubSubGroup) Then
                    AllowedMargin = MarginMap(SubSubGroup)
                End If

                Dim Remarks As String = ""

                If margin < AllowedMargin Then
                    Remarks = "Margin is not meeting"
                Else
                    Remarks = ""
                End If

                oDS3.SetValue("U_Remark", oDS3.Offset, Remarks)

                objmatrix.SetLineData(i)

            Next

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Private Sub CalculateMarginAndMarkup_LostFocus1()

        Try
            Dim objmatrix As SAPbouiCOM.Matrix = objForm.Items.Item("BarMtx").Specific
            Dim Objmtx3 As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx5").Specific
            Dim objMatrixItem As SAPbouiCOM.Matrix = objForm.Items.Item("ItmDetMtx").Specific

            oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C0") ' Item Master
            oDS3 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C2") ' Barcode Matrix
            oDS4 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C3") ' Price Matrix

            Dim VatCodeMap As New Dictionary(Of String, String)
            Dim ItemSugarMap As New Dictionary(Of String, String)
            Dim PriceDict As New Dictionary(Of String, Double)
            Dim VatDict As New Dictionary(Of String, Double)

            Const PURCHASE_LIST As String = "11"
            Const RETAIL_LIST As String = "1"

            '================= STEP 1: ITEM → VAT + SUGAR =================
            For x As Integer = 1 To objMatrixItem.VisualRowCount
                objMatrixItem.GetLineData(x)

                Dim key As String = oDS1.GetValue("U_KEYC", oDS1.Offset).Trim()
                Dim vat As String = oDS1.GetValue("U_PVAT", oDS1.Offset).Trim()
                Dim sugar As String = oDS1.GetValue("U_SugarTax", oDS1.Offset).Trim().ToUpper()

                If key <> "" Then

                    If Not VatCodeMap.ContainsKey(key) Then
                        VatCodeMap.Add(key, vat)
                    End If

                    If Not ItemSugarMap.ContainsKey(key) Then
                        ItemSugarMap.Add(key, sugar)
                    End If

                End If
            Next

            '================= STEP 2: PRICE DICTIONARY =================
            For j As Integer = 1 To Objmtx3.VisualRowCount
                Objmtx3.GetLineData(j)

                Dim itemCode As String = oDS4.GetValue("U_ItemC", oDS4.Offset).Trim()
                Dim uom As String = oDS4.GetValue("U_UOMCODE", oDS4.Offset).Trim()
                'Dim priceList As String = oDS4.GetValue("U_PriceId", oDS4.Offset).Trim()

                'Dim priceVal As Double
                'If Double.TryParse(oDS4.GetValue("U_OldPrice", oDS4.Offset), priceVal) Then

                '    Dim key As String = itemCode & "|" & uom & "|" & priceList

                '    If Not PriceDict.ContainsKey(key) Then
                '        PriceDict.Add(key, priceVal)
                '    End If

                'End If

                ' -------- Price List 1 --------
                Dim priceVal1 As Double
                If Double.TryParse(oDS4.GetValue("U_OldPrice", oDS4.Offset), priceVal1) Then

                    Dim key1 As String = itemCode & "|" & uom & "|1"

                    If Not PriceDict.ContainsKey(key1) Then
                        PriceDict.Add(key1, priceVal1)
                    End If

                End If

                ' -------- Price List 11 --------
                Dim priceVal11 As Double
                If Double.TryParse(oDS4.GetValue("U_Cost", oDS4.Offset), priceVal11) Then

                    Dim key11 As String = itemCode & "|" & uom & "|11"

                    If Not PriceDict.ContainsKey(key11) Then
                        PriceDict.Add(key11, priceVal11)
                    End If

                End If

            Next

            '================= STEP 3: VAT =================
            Dim oRsVAT As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsVAT.DoQuery("SELECT ""Code"", ""Rate"" FROM OVTG")

            While Not oRsVAT.EoF
                VatDict(oRsVAT.Fields.Item("Code").Value.ToString()) = CDbl(oRsVAT.Fields.Item("Rate").Value)
                oRsVAT.MoveNext()
            End While
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRsVAT)

            '==================For getting item Group ====================
            Dim ItemGroupMap As New Dictionary(Of String, String)

            'FIRST TAB MATRIX
            Dim SubSubGroupMap As New Dictionary(Of String, String)

            'FIRST TAB
            For i As Integer = 1 To objMatrixItem.VisualRowCount

                objMatrixItem.GetLineData(i)

                Dim KeyCode As String = oDS1.GetValue("U_KEYC", oDS1.Offset).Trim()
                Dim SubSubGroup As String = oDS1.GetValue("U_SSGP", oDS1.Offset).Trim()

                If Not SubSubGroupMap.ContainsKey(KeyCode) Then
                    SubSubGroupMap.Add(KeyCode, SubSubGroup)
                End If

            Next

            '================= STEP 4: MAIN LOOP =================
            For i As Integer = 1 To objmatrix.VisualRowCount
                If i = 1 Then
                    Dim test As String = "0"
                End If
                objmatrix.GetLineData(i)

                Dim itemCode As String = oDS3.GetValue("U_KEYC", oDS3.Offset).Trim()
                Dim uom As String = oDS3.GetValue("U_UOMCDE", oDS3.Offset).Trim()
                Dim sugarFlag As String = oDS3.GetValue("U_STAX", oDS3.Offset).Trim().ToUpper()

                ' 👉 Get Item-level sugar flag
                Dim itemSugarFlag As String = ""
                If ItemSugarMap.ContainsKey(itemCode) Then
                    itemSugarFlag = ItemSugarMap(itemCode)
                End If

                ' 👉 Get VAT
                Dim vatCode As String = ""
                If VatCodeMap.ContainsKey(itemCode) Then
                    vatCode = VatCodeMap(itemCode)
                End If

                Dim purchasePrice As Decimal = 0D
                Dim retailPrice As Decimal = 0D
                Dim vatPercent As Decimal = 0D
                vatPercent = If(VatDict.ContainsKey(vatCode), VatDict(vatCode), 0)

                '================= 🔥 BOTH CONDITIONS =================
                If itemSugarFlag = "YES" AndAlso sugarFlag = "YES" Then

                    Dim sugarTax As Decimal = 0D
                    Dim exciseTax As Decimal = 0D
                    Decimal.TryParse(oDS3.GetValue("U_SUTAXAMT", oDS3.Offset), sugarTax)
                    Decimal.TryParse(oDS3.GetValue("U_ETAX", oDS3.Offset), exciseTax)

                    sugarTax = Math.Round(sugarTax, 3)
                    exciseTax = Math.Round(exciseTax, 3)

                    ''Dim exciseTax As Decimal = Math.Round(sugarTax * 0.5D, 3)
                    purchasePrice = Math.Round(sugarTax + exciseTax, 3)

                    ''oDS3.SetValue("U_ETAX", oDS3.Offset, exciseTax.ToString())
                    oDS3.SetValue("U_Price", oDS3.Offset, purchasePrice.ToString())

                Else
                    Dim purchaseKey As String = itemCode & "|" & uom & "|" & PURCHASE_LIST
                    If PriceDict.ContainsKey(purchaseKey) Then
                        purchasePrice = PriceDict(purchaseKey)
                    End If
                End If

                '================= RETAIL =================
                Dim retailKey As String = itemCode & "|" & uom & "|" & RETAIL_LIST
                If PriceDict.ContainsKey(retailKey) Then
                    retailPrice = PriceDict(retailKey)
                End If
                purchasePrice = Math.Round(purchasePrice, 3)
                retailPrice = Math.Round(retailPrice, 3)
                vatPercent = Math.Round(vatPercent, 2)

                '================= VAT CALCULATION =================
                Dim vatAmount As Decimal = Math.Round(purchasePrice * vatPercent / 100D, 5)
                Dim purchaseWithVAT As Decimal = Math.Round(purchasePrice + vatAmount, 3)

                '================= CALCULATION =================
                Dim margin As Double = 0
                Dim markup As Double = 0

                If retailPrice > 0 Then
                    margin = Math.Round(((retailPrice - purchaseWithVAT) / retailPrice) * 100, 3)
                End If

                If purchaseWithVAT > 0 Then
                    markup = Math.Round(((retailPrice - purchaseWithVAT) / purchaseWithVAT) * 100, 3)
                End If

                '================= VAT =================
                'purchasePrice = Math.Round(purchasePrice, 3)

                'Dim purchaseWithVAT As Double = Math.Round(purchasePrice * (1 + vatPercent / 100), 5)
                '' Dim purchaseWithVAT As Double = purchasePrice * (1 + vatPercent / 100)

                ''================= CALCULATION =================
                'Dim margin As Double = 0
                'Dim markup As Double = 0

                'If retailPrice > 0 Then
                '    margin = ((retailPrice - purchaseWithVAT) / retailPrice) * 100
                'End If

                'If purchaseWithVAT > 0 Then
                '    markup = ((retailPrice - purchaseWithVAT) / purchaseWithVAT) * 100
                'End If

                '================= SAVE =================
                oDS3.SetValue("U_Margin", oDS3.Offset, Math.Round(margin, 3).ToString())
                oDS3.SetValue("U_Markup", oDS3.Offset, Math.Round(markup, 3).ToString())


                '===================Remarks getting ===============================

                Dim MarginMap As New Dictionary(Of String, Decimal)

                Dim rsMargin As SAPbobsCOM.Recordset
                rsMargin = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                Dim qry As String = "SELECT ""Name"", ""U_Margin"" FROM ""@VR_SUBSUBGROUP"""

                rsMargin.DoQuery(qry)

                While rsMargin.EoF = False

                    Dim SubSubGroupNew As String = rsMargin.Fields.Item("Name").Value.ToString().Trim()
                    Dim MarginREM As Decimal = 0D

                    Decimal.TryParse(rsMargin.Fields.Item("U_Margin").Value.ToString(), MarginREM)

                    If Not MarginMap.ContainsKey(SubSubGroupNew) Then
                        MarginMap.Add(SubSubGroupNew, MarginREM)
                    End If

                    rsMargin.MoveNext()

                End While

                Dim SubSubGroup As String = ""
                Dim AllowedMargin As Decimal = 0D

                If SubSubGroupMap.ContainsKey(itemCode) Then
                    SubSubGroup = SubSubGroupMap(itemCode)
                End If

                If MarginMap.ContainsKey(SubSubGroup) Then
                    AllowedMargin = MarginMap(SubSubGroup)
                End If

                Dim Remarks As String = ""

                If margin < AllowedMargin Then
                    Remarks = "Margin is not meeting"
                Else
                    Remarks = ""
                End If

                oDS3.SetValue("U_Remark", oDS3.Offset, Remarks)

                objmatrix.SetLineData(i)

            Next

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Private Sub CalculateMarginAndMarkup_LostFocus1(ByVal RowNo As Integer)

        Try
            Dim objmatrix As SAPbouiCOM.Matrix = objForm.Items.Item("BarMtx").Specific
            Dim Objmtx3 As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx5").Specific
            Dim objMatrixItem As SAPbouiCOM.Matrix = objForm.Items.Item("ItmDetMtx").Specific

            oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C0") ' Item Master
            oDS3 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C2") ' Barcode Matrix
            oDS4 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C3") ' Price Matrix

            Dim VatCodeMap As New Dictionary(Of String, String)
            Dim ItemSugarMap As New Dictionary(Of String, String)
            Dim PriceDict As New Dictionary(Of String, Double)
            Dim VatDict As New Dictionary(Of String, Double)

            Const PURCHASE_LIST As String = "11"
            Const RETAIL_LIST As String = "1"

            '================= STEP 1: ITEM → VAT + SUGAR =================
            For x As Integer = 1 To objMatrixItem.VisualRowCount
                objMatrixItem.GetLineData(x)

                Dim key As String = oDS1.GetValue("U_KEYC", oDS1.Offset).Trim()
                Dim vat As String = oDS1.GetValue("U_PVAT", oDS1.Offset).Trim()
                Dim sugar As String = oDS1.GetValue("U_SugarTax", oDS1.Offset).Trim().ToUpper()

                If key <> "" Then

                    If Not VatCodeMap.ContainsKey(key) Then
                        VatCodeMap.Add(key, vat)
                    End If

                    If Not ItemSugarMap.ContainsKey(key) Then
                        ItemSugarMap.Add(key, sugar)
                    End If

                End If
            Next

            '================= STEP 2: PRICE DICTIONARY =================
            For j As Integer = 1 To Objmtx3.VisualRowCount
                Objmtx3.GetLineData(j)

                Dim itemCode As String = oDS4.GetValue("U_ItemC", oDS4.Offset).Trim()
                Dim uom As String = oDS4.GetValue("U_UOMCODE", oDS4.Offset).Trim()
                'Dim priceList As String = oDS4.GetValue("U_PriceId", oDS4.Offset).Trim()

                'Dim priceVal As Double
                'If Double.TryParse(oDS4.GetValue("U_OldPrice", oDS4.Offset), priceVal) Then

                '    Dim key As String = itemCode & "|" & uom & "|" & priceList

                '    If Not PriceDict.ContainsKey(key) Then
                '        PriceDict.Add(key, priceVal)
                '    End If

                'End If

                ' -------- Price List 1 --------
                Dim priceVal1 As Double
                If Double.TryParse(oDS4.GetValue("U_OldPrice", oDS4.Offset), priceVal1) Then

                    Dim key1 As String = itemCode & "|" & uom & "|1"

                    If Not PriceDict.ContainsKey(key1) Then
                        PriceDict.Add(key1, priceVal1)
                    End If

                End If

                ' -------- Price List 11 --------
                Dim priceVal11 As Double
                If Double.TryParse(oDS4.GetValue("U_Cost", oDS4.Offset), priceVal11) Then

                    Dim key11 As String = itemCode & "|" & uom & "|11"

                    If Not PriceDict.ContainsKey(key11) Then
                        PriceDict.Add(key11, priceVal11)
                    End If

                End If

            Next

            '================= STEP 3: VAT =================
            Dim oRsVAT As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsVAT.DoQuery("SELECT ""Code"", ""Rate"" FROM OVTG")

            While Not oRsVAT.EoF
                VatDict(oRsVAT.Fields.Item("Code").Value.ToString()) = CDbl(oRsVAT.Fields.Item("Rate").Value)
                oRsVAT.MoveNext()
            End While
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRsVAT)

            '==================For getting item Group ====================
            Dim ItemGroupMap As New Dictionary(Of String, String)

            'FIRST TAB MATRIX
            Dim SubSubGroupMap As New Dictionary(Of String, String)

            'FIRST TAB
            For ItemRow As Integer = 1 To objMatrixItem.VisualRowCount

                objMatrixItem.GetLineData(ItemRow)

                Dim KeyCode As String = oDS1.GetValue("U_KEYC", oDS1.Offset).Trim()
                Dim ItemSubSubGroup As String = oDS1.GetValue("U_SSGP", oDS1.Offset).Trim()

                If Not SubSubGroupMap.ContainsKey(KeyCode) Then
                    SubSubGroupMap.Add(KeyCode, ItemSubSubGroup)
                End If

            Next

            '================= STEP 4: MAIN LOOP =================
            objmatrix.GetLineData(RowNo)

            Dim CurrItemCode As String = oDS3.GetValue("U_KEYC", oDS3.Offset).Trim()
            Dim CurrUom As String = oDS3.GetValue("U_UOMCDE", oDS3.Offset).Trim()
            Dim sugarFlag As String = oDS3.GetValue("U_STAX", oDS3.Offset).Trim().ToUpper()




            ' 👉 Get Item-level sugar flag
            Dim itemSugarFlag As String = ""
            If ItemSugarMap.ContainsKey(CurrItemCode) Then
                itemSugarFlag = ItemSugarMap(CurrItemCode)
            End If

            ' 👉 Get VAT
            Dim vatCode As String = ""
            If VatCodeMap.ContainsKey(CurrItemCode) Then
                vatCode = VatCodeMap(CurrItemCode)
            End If

            Dim purchasePrice As Decimal = 0D
            Dim retailPrice As Decimal = 0D
            Dim vatPercent As Decimal = 0D
            vatPercent = If(VatDict.ContainsKey(vatCode), VatDict(vatCode), 0)

            '================= 🔥 BOTH CONDITIONS =================
            If itemSugarFlag = "YES" AndAlso sugarFlag = "YES" Then

                Dim sugarTax As Decimal = 0D
                Dim exciseTax As Decimal = 0D
                Decimal.TryParse(oDS3.GetValue("U_SUTAXAMT", oDS3.Offset), sugarTax)
                Decimal.TryParse(oDS3.GetValue("U_ETAX", oDS3.Offset), exciseTax)

                sugarTax = Math.Round(sugarTax, 3)
                exciseTax = Math.Round(exciseTax, 3)

                ''Dim exciseTax As Decimal = Math.Round(sugarTax * 0.5D, 3)
                purchasePrice = Math.Round(sugarTax + exciseTax, 3)

                ''oDS3.SetValue("U_ETAX", oDS3.Offset, exciseTax.ToString())
                oDS3.SetValue("U_Price", oDS3.Offset, purchasePrice.ToString())

            Else
                Dim purchaseKey As String = CurrItemCode & "|" & CurrUom & "|" & PURCHASE_LIST
                If PriceDict.ContainsKey(purchaseKey) Then
                    purchasePrice = PriceDict(purchaseKey)
                End If
            End If

            '================= RETAIL =================
            Dim retailKey As String = CurrItemCode & "|" & CurrUom & "|" & RETAIL_LIST
            If PriceDict.ContainsKey(retailKey) Then
                retailPrice = PriceDict(retailKey)
            End If
            purchasePrice = Math.Round(purchasePrice, 3)
            retailPrice = Math.Round(retailPrice, 3)
            vatPercent = Math.Round(vatPercent, 2)

            '================= VAT CALCULATION =================
            Dim vatAmount As Decimal = Math.Round(purchasePrice * vatPercent / 100D, 5)
            Dim purchaseWithVAT As Decimal = Math.Round(purchasePrice + vatAmount, 3)

            '================= CALCULATION =================
            Dim margin As Double = 0
            Dim markup As Double = 0

            If retailPrice > 0 Then
                margin = Math.Round(((retailPrice - purchaseWithVAT) / retailPrice) * 100, 3)
            End If

            If purchaseWithVAT > 0 Then
                markup = Math.Round(((retailPrice - purchaseWithVAT) / purchaseWithVAT) * 100, 3)
            End If

            '================= VAT =================
            'purchasePrice = Math.Round(purchasePrice, 3)

            'Dim purchaseWithVAT As Double = Math.Round(purchasePrice * (1 + vatPercent / 100), 5)
            '' Dim purchaseWithVAT As Double = purchasePrice * (1 + vatPercent / 100)

            ''================= CALCULATION =================
            'Dim margin As Double = 0
            'Dim markup As Double = 0

            'If retailPrice > 0 Then
            '    margin = ((retailPrice - purchaseWithVAT) / retailPrice) * 100
            'End If

            'If purchaseWithVAT > 0 Then
            '    markup = ((retailPrice - purchaseWithVAT) / purchaseWithVAT) * 100
            'End If

            '================= SAVE =================
            oDS3.SetValue("U_Margin", oDS3.Offset, Math.Round(margin, 3).ToString())
            oDS3.SetValue("U_Markup", oDS3.Offset, Math.Round(markup, 3).ToString())


            '===================Remarks getting ===============================

            Dim MarginMap As New Dictionary(Of String, Decimal)

            Dim rsMargin As SAPbobsCOM.Recordset
            rsMargin = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim qry As String = "SELECT ""Name"", ""U_Margin"" FROM ""@VR_SUBSUBGROUP"""

            rsMargin.DoQuery(qry)

            While rsMargin.EoF = False

                Dim SubSubGroupNew As String = rsMargin.Fields.Item("Name").Value.ToString().Trim()
                Dim MarginREM As Decimal = 0D

                Decimal.TryParse(rsMargin.Fields.Item("U_Margin").Value.ToString(), MarginREM)

                If Not MarginMap.ContainsKey(SubSubGroupNew) Then
                    MarginMap.Add(SubSubGroupNew, MarginREM)
                End If

                rsMargin.MoveNext()

            End While

            Dim SubSubGroup As String = ""
            Dim AllowedMargin As Decimal = 0D

            If SubSubGroupMap.ContainsKey(CurrItemCode) Then
                SubSubGroup = SubSubGroupMap(CurrItemCode)
            End If

            If MarginMap.ContainsKey(SubSubGroup) Then
                AllowedMargin = MarginMap(SubSubGroup)
            End If

            Dim Remarks As String = ""

            If margin < AllowedMargin Then
                Remarks = "Margin is not meeting"
            Else
                Remarks = ""
            End If

            oDS3.SetValue("U_Remark", oDS3.Offset, Remarks)

            objmatrix.SetLineData(RowNo)



        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
    Private Sub CalculateMarginAndMarkup_Fast_41()

        Try
            Dim objmatrix As SAPbouiCOM.Matrix = objForm.Items.Item("BarMtx").Specific
            Dim Objmtx3 As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx5").Specific

            oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C0") ' Item Master
            oDS3 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C2") ' Bar Matrix
            oDS4 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C3") ' Price Matrix
            Dim objMatrixItem As SAPbouiCOM.Matrix = objForm.Items.Item("ItmDetMtx").Specific
            Dim VatCodeMap As New Dictionary(Of String, String)

            '================= STEP 1: BUILD ITEM → VAT MAP =================

            For x As Integer = 1 To objMatrixItem.VisualRowCount
                objMatrixItem.GetLineData(x)
                Dim key As String = oDS1.GetValue("U_KEYC", oDS1.Offset).Trim()
                Dim vat As String = oDS1.GetValue("U_PVAT", oDS1.Offset).Trim()
                Dim sugartax As String = oDS1.GetValue("U_SugarTax", oDS1.Offset).Trim()
                If key <> "" And sugartax = "Yes" Then
                    If Not VatCodeMap.ContainsKey(key) Then
                        VatCodeMap.Add(key, vat)
                    End If
                End If
            Next

            '================= STEP 2: PRICE DICTIONARY =================
            Dim PriceDict As New Dictionary(Of String, Double)

            For j As Integer = 1 To Objmtx3.VisualRowCount

                Objmtx3.GetLineData(j)


                Dim itemCode As String = oDS4.GetValue("U_ItemC", oDS4.Offset).Trim()
                Dim uom As String = oDS4.GetValue("U_UOMCODE", oDS4.Offset).Trim()
                Dim priceList As String = oDS4.GetValue("U_PriceId", oDS4.Offset).Trim()

                Dim priceVal As Double
                If Double.TryParse(oDS4.GetValue("U_OldPrice", oDS4.Offset), priceVal) Then

                    If itemCode <> "" AndAlso uom <> "" AndAlso priceList <> "" Then
                        Dim key As String = itemCode & "|" & uom & "|" & priceList

                        If Not PriceDict.ContainsKey(key) Then
                            PriceDict.Add(key, priceVal)
                        End If
                    End If

                End If

            Next

            '================= STEP 3: VAT % =================
            Dim VatDict As New Dictionary(Of String, Double)

            Dim oRsVAT As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsVAT.DoQuery("SELECT ""Code"", ""Rate"" FROM OVTG")

            While Not oRsVAT.EoF
                VatDict(oRsVAT.Fields.Item("Code").Value.ToString()) = CDbl(oRsVAT.Fields.Item("Rate").Value)
                oRsVAT.MoveNext()
            End While

            '================= STEP 4: MAIN LOOP =================
            For i As Integer = 1 To objmatrix.VisualRowCount

                objmatrix.GetLineData(i)
                ' ✅ ONLY THIS OFFSET

                Dim itemCode As String = oDS3.GetValue("U_KEYC", oDS3.Offset).Trim()
                Dim uom As String = oDS3.GetValue("U_UOMCDE", oDS3.Offset).Trim()
                Dim sugarFlag As String = oDS3.GetValue("U_STAX", oDS3.Offset).Trim()

                ' 👉 Get VAT from mapping
                Dim vatCode As String = ""
                If VatCodeMap.ContainsKey(itemCode) Then
                    vatCode = VatCodeMap(itemCode)
                End If

                Dim sugarTax As Double = 0
                Dim exciseTax As Double = 0
                Dim purchasePrice As Double = 0

                '================= SUGAR TAX =================
                If sugarFlag = "Yes" Then

                    Double.TryParse(oDS3.GetValue("U_SUTAXAMT", oDS3.Offset), sugarTax)

                    exciseTax = sugarTax * 0.5
                    purchasePrice = sugarTax + exciseTax

                    oDS3.SetValue("U_ETAX", oDS3.Offset, exciseTax.ToString())
                    oDS3.SetValue("U_Price", oDS3.Offset, purchasePrice.ToString())

                Else
                    Dim purchaseKey As String = itemCode & "|" & uom & "|11"

                    If PriceDict.ContainsKey(purchaseKey) Then
                        purchasePrice = PriceDict(purchaseKey)
                    End If
                End If

                '================= RETAIL =================
                Dim retailPrice As Double = 0
                Dim retailKey As String = itemCode & "|" & uom & "|1"

                If PriceDict.ContainsKey(retailKey) Then
                    retailPrice = PriceDict(retailKey)
                End If

                '================= VAT =================
                Dim vatPercent As Double = 0
                If VatDict.ContainsKey(vatCode) Then
                    vatPercent = VatDict(vatCode)
                End If

                Dim purchaseWithVAT As Double = purchasePrice * (1 + vatPercent / 100)


                '================= CALCULATIONS =================
                Dim margin As Double = 0
                Dim markup As Double = 0
                Dim diff As Double = retailPrice - purchaseWithVAT

                If retailPrice > 0 Then
                    margin = (diff / retailPrice) * 100
                End If

                If purchasePrice > 0 Then
                    markup = (diff / purchaseWithVAT) * 100
                End If


                'If retailPrice > 0 Then
                '    margin = ((retailPrice - purchaseWithVAT) / retailPrice) * 100
                'End If

                'If purchaseWithVAT > 0 Then
                '    markup = ((retailPrice - purchaseWithVAT) / purchaseWithVAT) * 100
                'End If

                '================= SAVE =================
                oDS3.SetValue("U_Margin", oDS3.Offset, Math.Round(margin, 3).ToString())
                oDS3.SetValue("U_Markup", oDS3.Offset, Math.Round(markup, 3).ToString())

                objmatrix.SetLineData(i)

            Next

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
    Private Sub CalculateMarginAndMarkup_Fast_13()

        Try
            Dim objmatrix As SAPbouiCOM.Matrix = objForm.Items.Item("ItmDetMtx").Specific
            Dim Objmtx3 As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx5").Specific
            oDS = Me.objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP")
            oDS1 = Me.objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C0")
            oDS2 = Me.objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C1")
            oDS3 = Me.objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C2")
            oDS4 = Me.objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C3")
            oDS5 = Me.objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C4")
            oDS6 = Me.objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C5")
            '================= STEP 1: BUILD PRICE DICTIONARY =================
            Dim PriceDict As New Dictionary(Of String, Double)

            For j As Integer = 1 To Objmtx3.VisualRowCount

                Objmtx3.GetLineData(j)

                Dim Key As String = oDS4.GetValue("U_ItemC", oDS4.Offset).Trim()
                Dim priceStr As String = oDS4.GetValue("U_OldPrice", oDS4.Offset).Trim()

                If Key <> "" AndAlso priceStr <> "" Then
                    If Not PriceDict.ContainsKey(Key) Then
                        PriceDict.Add(Key, CDbl(priceStr))
                    End If
                End If

            Next

            '================= STEP 2: BUILD VAT DICTIONARY (OPTIMIZED) =================
            Dim VatDict As New Dictionary(Of String, Double)

            Dim oRsVAT As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsVAT.DoQuery("SELECT ""Code"", ""Rate"" FROM OVTG")

            While Not oRsVAT.EoF
                VatDict(oRsVAT.Fields.Item("Code").Value.ToString()) = CDbl(oRsVAT.Fields.Item("Rate").Value)
                oRsVAT.MoveNext()
            End While

            '================= STEP 3: LOOP ITEM MATRIX =================
            For i As Integer = 1 To objmatrix.VisualRowCount

                objmatrix.GetLineData(i)

                Dim KeyCode As String = oDS1.GetValue("U_KEYC", oDS1.Offset).Trim()
                Dim PurchaseVATCode As String = oDS1.GetValue("U_PVAT", oDS1.Offset).Trim()
                Dim SugarTaxAmtStr As String = oDS1.GetValue("U_SuTaxAmt", oDS1.Offset).Trim()

                Dim RetailPriceWithVAT As Double = 0
                Dim LastPurchaseWithVAT As Double = 0
                Dim Margin As Double = 0
                Dim Markup As Double = 0
                Dim VatPercent As Double = 0
                Dim PurchasePrice As Double = 0

                '--- Purchase Price (Sugar Tax included as per your logic) ---
                If SugarTaxAmtStr <> "" Then
                    PurchasePrice = CDbl(SugarTaxAmtStr)
                End If

                '--- Get VAT % from Dictionary ---
                If VatDict.ContainsKey(PurchaseVATCode) Then
                    VatPercent = VatDict(PurchaseVATCode)
                End If

                '--- Get Retail Price from Dictionary ---
                If PriceDict.ContainsKey(KeyCode) Then
                    RetailPriceWithVAT = PriceDict(KeyCode)
                End If

                '================= CORRECT CALCULATION =================

                ' Purchase WITH VAT
                LastPurchaseWithVAT = PurchasePrice * (1 + VatPercent / 100)

                ' Difference
                Dim Diff As Double = RetailPriceWithVAT - LastPurchaseWithVAT

                ' Margin
                If RetailPriceWithVAT > 0 Then
                    Margin = (Diff / RetailPriceWithVAT) * 100
                End If

                ' Markup
                If LastPurchaseWithVAT > 0 Then
                    Markup = (Diff / LastPurchaseWithVAT) * 100
                End If

                '================= SET VALUES =================
                oDS1.SetValue("U_Margin", oDS1.Offset, Math.Round(Margin, 2).ToString())
                oDS1.SetValue("U_Markup", oDS1.Offset, Math.Round(Markup, 2).ToString())

                objmatrix.SetLineData(i)

            Next

        Catch ex As Exception
            objForm.Items.Item("Item_11").Enabled = True
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    'Akhila Bandoju
    '----Automatic Creation of item master data based on Screen--30-12-2024----
    'Private Sub Item_Creation_BackUp(ByVal DocEntry As String)
    '    Try
    '        Dim oForm As SAPbouiCOM.Form
    '        Dim addedGroupCode As String
    '        Dim newitemcreated As String
    '        Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("ItmDetMtx").Specific

    '        Dim qStr As String = "select T0.""DocEntry"",T0.""LineId"",T0.""U_KEYC"", T0.""U_ITEMN"",T0.""U_FORIE"",T0.""U_CARD"",T0.""U_ITEMG"",T0.""U_SUBGR"",T0.""U_SSGP"",T0.""U_SUGRP"",T0.""U_BRAND"" ,T0.""U_CATGO"", T0.""U_COUNT"" ,T0.""U_BUNIT"",T0.""U_SCALE"",
    '        T0.""U_ITEMS"", T0.""U_INVENT"",T0.""U_SALITM"" ,T0.""U_PURCH"",T0.""U_NOP"",T0.""U_BaseUt"",T0.""U_UOMD"",T0.""U_CNum"" from ""@TNX_ITEMUP"" T inner join ""@TNX_ITEMUP_C0"" T0 on T.""DocEntry""= T0.""DocEntry"" where T.""DocNum""='" & DocEntry & "' and IFNULL(""U_STATS"",'')<>'Item Master is Succesfully Created' ORDER BY T0.""LineId""  "
    '        Dim rSet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        rSet.DoQuery(qStr)
    '        Dim DOCENTRY11 As String = ""
    '        If rSet.RecordCount > 0 Then
    '            For i = 1 To rSet.RecordCount

    '                If i = 13 Then
    '                    Dim T As String = 0
    '                End If
    '                objMain.objApplication.StatusBar.SetText("Creating Items " & i & " out of " & rSet.RecordCount & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

    '                Dim Itemcodeas As String = objMatrix.Columns.Item("ITMCD").Cells.Item(i).Specific.Value.ToString()
    '                '' Dim Itemcodeas As String = ""
    '                Dim Ik As String = objMatrix.Columns.Item("KEYC").Cells.Item(i).Specific.Value.ToString()
    '                objMain.objApplication.StatusBar.SetText("Creating Item for key code --, " & Itemcodeas & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '                Dim LINE As Integer
    '                If Itemcodeas = "" And Ik <> "" Then
    '                    Try
    '                        ' objMain.objCompany.StartTransaction()
    '                        oItem = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
    '                        DOCENTRY11 = rSet.Fields.Item("DocEntry").Value.ToString()

    '                        Dim ItemCode As String = objMatrix.Columns.Item("KEYC").Cells.Item(i).Specific.Value.ToString()
    '                        Dim ItemName As String = objMatrix.Columns.Item("ITEMN").Cells.Item(i).Specific.Value.ToString()
    '                        Dim CardCode As String = objMatrix.Columns.Item("CARD").Cells.Item(i).Specific.Value.ToString()
    '                        Dim FORIE As String = objMatrix.Columns.Item("FORIE").Cells.Item(i).Specific.Value.ToString()
    '                        Dim FItemCode As String
    '                        FItemCode = ItemCode
    '                        LINE = rSet.Fields.Item("LineId").Value.ToString()
    '                        Dim keycod As String = rSet.Fields.Item("U_KEYC").Value.ToString()
    '                        Dim ItemNme As String = rSet.Fields.Item("U_ITEMN").Value.ToString()
    '                        Dim ForineName As String = rSet.Fields.Item("U_FORIE").Value.ToString()
    '                        Dim CardCode1 As String = rSet.Fields.Item("U_CARD").Value.ToString()
    '                        Dim ItemGrp As String = rSet.Fields.Item("U_ITEMG").Value.ToString()
    '                        Dim SubGrp As String = rSet.Fields.Item("U_SUBGR").Value.ToString()
    '                        Dim SubSubgroup1 As String = rSet.Fields.Item("U_SSGP").Value.ToString()
    '                        Dim Brand As String = rSet.Fields.Item("U_BRAND").Value.ToString()
    '                        Dim Category1 As String = rSet.Fields.Item("U_CATGO").Value.ToString()
    '                        Dim ScaleItem1 As String = rSet.Fields.Item("U_SCALE").Value.ToString()
    '                        Dim CountryOfOrgin1 As String = rSet.Fields.Item("U_COUNT").Value.ToString()
    '                        Dim BaseUnit1 As String = rSet.Fields.Item("U_BUNIT").Value.ToString()
    '                        Dim ItemSize1 As String = rSet.Fields.Item("U_ITEMS").Value.ToString()
    '                        Dim InventoryItem1 As String = rSet.Fields.Item("U_INVENT").Value.ToString()
    '                        Dim SaleItem1 As String = rSet.Fields.Item("U_SALITM").Value.ToString()
    '                        Dim PurchaseItem As String = rSet.Fields.Item("U_PURCH").Value.ToString()
    '                        Dim subsubgrp As String = rSet.Fields.Item("U_SUGRP").Value.ToString()
    '                        Dim Noofpcs As String = rSet.Fields.Item("U_NOP").Value.ToString()
    '                        Dim BaseUnit As String = rSet.Fields.Item("U_BaseUt").Value.ToString()
    '                        Dim UomDimension As String = rSet.Fields.Item("U_UOMD").Value.ToString()
    '                        Dim CatlogNumber As String = rSet.Fields.Item("U_CNum").Value.ToString()
    '                        Dim seris As String = "Select ""U_SERIES"" as ""Series"" FROM ""OADM"" "
    '                        Dim rSeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                        rSeries.DoQuery(seris)
    '                        If rSeries.RecordCount > 0 And rSeries.Fields.Item("Series").Value IsNot Nothing AndAlso rSeries.Fields.Item("Series").Value.ToString() <> "" AndAlso Convert.ToInt32(rSeries.Fields.Item("Series").Value) <> 0 Then
    '                            oItem.Series = Convert.ToInt32(rSeries.Fields.Item("Series").Value)
    '                        End If
    '                        Dim obicod As String = "Select ""ObjectCode"" FROM ""NNM1"" where ""Series"" = '" & rSeries.Fields.Item("Series").Value & "' "
    '                        Dim robj As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                        robj.DoQuery(obicod)
    '                        Dim objcode As String = robj.Fields.Item("ObjectCode").Value
    '                        Dim ItmIO As String = " Select  (""NextNumber"") as ""NextNumber""  from ""NNM1"" where ""ObjectCode""='" & objcode & "' and ""Series"" = '" & rSeries.Fields.Item("Series").Value & "' "
    '                        Dim oRsItm As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                        oRsItm.DoQuery(ItmIO)
    '                        Dim BARCODE As String
    '                        Dim defaultBARCODE As String
    '                        ''UOM creation
    '                        Dim oGrpParams As SAPbobsCOM.UnitOfMeasurementGroupParams
    '                        Dim NewItemCode1 As String = OtherCompany.GetNewObjectKey.ToString()
    '                        Try
    '                            Dim oCompService As SAPbobsCOM.CompanyService = objMain.objCompany.GetCompanyService()
    '                            Dim oUoMGroupService As SAPbobsCOM.UnitOfMeasurementGroupsService = CType(oCompService.GetBusinessService(SAPbobsCOM.ServiceTypes.UnitOfMeasurementGroupsService), SAPbobsCOM.UnitOfMeasurementGroupsService)
    '                            Dim oUoMGroup As SAPbobsCOM.UnitOfMeasurementGroup = CType(oUoMGroupService.GetDataInterface(SAPbobsCOM.UnitOfMeasurementGroupsServiceDataInterfaces.uomgsUnitOfMeasurementGroup), SAPbobsCOM.UnitOfMeasurementGroup)
    '                            Dim nextNumber As String = oRsItm.Fields.Item("NextNumber").Value


    '                            Dim oUoMGroupParams As SAPbobsCOM.UnitOfMeasurementGroupParams = CType(oUoMGroupService.GetDataInterface(SAPbobsCOM.UnitOfMeasurementGroupsServiceDataInterfaces.uomgsUnitOfMeasurementGroupParams), SAPbobsCOM.UnitOfMeasurementGroupParams)

    '                            oUoMGroupParams.Code = nextNumber

    '                            ' Try to get the UoM Group
    '                            Try
    '                                oUoMGroupService.Delete(oUoMGroupParams)

    '                            Catch ex As Exception

    '                            End Try
    '                            ' If found, proceed to delete

    '                            Dim Uomqry As String = "select T2.""UomEntry"",""U_UOMUNT""  from ""@TNX_ITEMUP"" T0 inner join ""@TNX_ITEMUP_C1"" T1 On T0.""DocEntry""= T1.""DocEntry"" inner join ""OUOM"" T2  on T2.""UomCode""=T1.""U_UOMCDE"" " &
    '                                "where T0.""DocNum""='" & DocEntry & "'and ""U_KEYCDE"" ='" & FItemCode & "' order by ""U_UOMUNT"" asc"
    '                            Dim rsuomqry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                            rsuomqry.DoQuery(Uomqry)
    '                            If rsuomqry.RecordCount > 0 Then

    '                                Dim base As Double = rsuomqry.Fields.Item("U_UOMUNT").Value
    '                                Dim baseentry As Integer = rsuomqry.Fields.Item("UomEntry").Value
    '                                For j As Integer = 1 To rsuomqry.RecordCount
    '                                    If j = 1 Then
    '                                        oUoMGroup.Code = nextNumber
    '                                        oUoMGroup.Name = nextNumber
    '                                        oUoMGroup.BaseUoM = baseentry
    '                                        rsuomqry.MoveNext()
    '                                    Else
    '                                        Dim def1 As SAPbobsCOM.UoMGroupDefinition = oUoMGroup.GroupDefinitions.Add()
    '                                        Dim alternate As String = rsuomqry.Fields.Item("UomEntry").Value.ToString

    '                                        Dim alternateQty As String = rsuomqry.Fields.Item("U_UOMUNT").Value.ToString
    '                                        'def1.AlternateQuantity = rsuomqry.Fields.Item("U_UOMUNT").Value
    '                                        'def1.BaseQuantity = base
    '                                        def1.AlternateQuantity = base
    '                                        def1.BaseQuantity = rsuomqry.Fields.Item("U_UOMUNT").Value
    '                                        def1.AlternateUoM = rsuomqry.Fields.Item("UomEntry").Value


    '                                        rsuomqry.MoveNext()
    '                                    End If
    '                                Next
    '                            Else


    '                            End If

    '                            'Dim baseUoM As Integer = 1
    '                            'oUoMGroup.BaseUoM = baseUoM

    '                            'Dim def1 As SAPbobsCOM.UoMGroupDefinition
    '                            'def1.AlternateUoM = 1 ' Alternate UoM Code (e.g., "BOX")
    '                            'def1.BaseQuantity = 1 ' Conversion to Base UoM (e.g., 1 BOX = 10 PCS)
    '                            'def1.AlternateQuantity = 1
    '                            'oUoMGroup.GroupDefinitions.Add()
    '                            oGrpParams = oUoMGroupService.Add(oUoMGroup)
    '                            '    oGrpParams.
    '                            addedGroupCode = oGrpParams.AbsEntry
    '                            objMain.objApplication.StatusBar.SetText("UoM Group added successfully with Code: " & addedGroupCode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

    '                            objMain.objApplication.StatusBar.SetText("UOM Updating in current database", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    '                        Catch ex As Exception

    '                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                        End Try

    '                        If objMain.objCompany.InTransaction Then objMain.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

    '                        objMain.objCompany.StartTransaction()
    '                        oItem.ItemName = ItemName
    '                        oItem.ForeignName = ForineName
    '                        oItem.SalesVATGroup = "X0"
    '                        oItem.PurchaseVATGroup = "X1"
    '                        'oItem.ItemsGroupCode = "206"
    '                        'oItem.UoMGroupEntry = "399723"

    '                        oItem.UoMGroupEntry = addedGroupCode


    '                        oItem.PreferredVendors.BPCode = CardCode1
    '                        '  oItem.PreferredVendors.SetCurrentLine(0)
    '                        oItem.PreferredVendors.Add()

    '                        Dim Uomtest As String = "select top 1 T2.""UomEntry"",""U_UOMUNT""  from ""@TNX_ITEMUP"" T0 inner join ""@TNX_ITEMUP_C1"" T1 On T0.""DocEntry""= T1.""DocEntry"" inner join ""OUOM"" T2  on T2.""UomCode""=T1.""U_UOMCDE"" " &
    '                               "where T0.""DocNum""='" & DocEntry & "'and ""U_KEYCDE"" ='" & FItemCode & "' order by ""U_UOMUNT"" asc"
    '                        Dim rsuomqrytest As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                        rsuomqrytest.DoQuery(Uomtest)
    '                        Dim uom As Integer = 0
    '                        If rsuomqrytest.RecordCount > 0 Then
    '                            uom = rsuomqrytest.Fields.Item("UomEntry").Value
    '                            Dim base As Double = rsuomqrytest.Fields.Item("U_UOMUNT").Value
    '                        End If


    '                        'oItem.UnitOfMeasurements.UoMEntry = 93
    '                        'oItem.UnitOfMeasurements.
    '                        'oItem.UnitOfMeasurements.Add()
    '                        oItem.DefaultSalesUoMEntry = uom
    '                        ''oItem.UoMGroupEntry = oRsItm.Fields.Item("NextNumber").Value
    '                        oItem.InventoryUoMEntry = uom
    '                        oItem.SalesItem = SAPbobsCOM.BoYesNoEnum.tYES
    '                        oItem.PurchaseItem = SAPbobsCOM.BoYesNoEnum.tYES
    '                        oItem.InventoryItem = SAPbobsCOM.BoYesNoEnum.tYES
    '                        oItem.DefaultPurchasingUoMEntry = uom
    '                        oItem.UserFields.Fields.Item("U_BRAND").Value = Brand
    '                        oItem.UserFields.Fields.Item("U_VSubGroup").Value = SubGrp
    '                        oItem.UserFields.Fields.Item("U_VSubSubGroup").Value = SubSubgroup1
    '                        oItem.UserFields.Fields.Item("U_CATEGORY").Value = Category1
    '                        oItem.UserFields.Fields.Item("U_PcPerPac").Value = Noofpcs
    '                        oItem.UserFields.Fields.Item("U_BaseUnit").Value = BaseUnit1
    '                        oItem.UserFields.Fields.Item("U_ItemSize").Value = ItemSize1
    '                        oItem.UserFields.Fields.Item("U_CountryOrg").Value = CountryOfOrgin1
    '                        oItem.UserFields.Fields.Item("U_UOM").Value = UomDimension
    '                        If ScaleItem1 = "Yes" Then
    '                            oItem.UserFields.Fields.Item("U_AcceptWMQty").Value = "1" ' Yes as string
    '                        Else
    '                            oItem.UserFields.Fields.Item("U_AcceptWMQty").Value = "2" ' No as string
    '                        End If
    '                        oItem.UserFields.Fields.Item("U_3SUBGROUP").Value = subsubgrp



    '                        Dim barcodestest As String = "select T2.""UomEntry"", ""U_KEYC"",""U_UOMCDE"",""U_BarCode"",""U_BDesc"",T1.""DocEntry"",T1.""LineId"" from ""@TNX_ITEMUP"" T0 inner join ""@TNX_ITEMUP_C2"" T1 On T0.""DocEntry""= T1.""DocEntry"" inner join ""OUOM"" T2  on T2.""UomCode""=T1.""U_UOMCDE"" " &
    '                               "where T0.""DocNum""='" & DocEntry & "'and ""U_KEYC"" ='" & FItemCode & "'"
    '                        Dim barcoders As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                        barcoders.DoQuery(barcodestest)

    '                        If barcoders.RecordCount = 0 Then
    '                            defaultBARCODE = "No"

    '                            BARCODE = Me.New_Barcode_Generation(oRsItm.Fields.Item("NextNumber").Value, ItemName)

    '                            If BARCODE <> "" Then
    '                                BARCODE = BARCODE
    '                                oItem.BarCodes.BarCode = BARCODE
    '                                oItem.BarCodes.FreeText = ItemName
    '                                oItem.BarCodes.UoMEntry = uom
    '                                oItem.BarCodes.SetCurrentLine(0)
    '                                oItem.BarCodes.Add()
    '                            Else
    '                                rSet.MoveNext()

    '                            End If

    '                        Else
    '                            'For p As Integer = 1 To barcoders.RecordCount
    '                            For p As Integer = 0 To barcoders.RecordCount - 1
    '                                Dim barcodest As String = barcoders.Fields.Item("U_BarCode").Value
    '                                Dim barcodename As String = barcoders.Fields.Item("U_BDesc").Value
    '                                Dim barcodeuom As Integer = barcoders.Fields.Item("UomEntry").Value
    '                                Dim LineId As Integer = barcoders.Fields.Item("LineId").Value
    '                                'If barcodest <> "" And barcodename <> "" And barcodeuom <> 0 Then
    '                                If barcodest = "" And barcodename <> "" And barcodeuom <> 0 Then
    '                                    BARCODE = Me.New_Barcode_GenerationWithBarcodeDescription(oRsItm.Fields.Item("NextNumber").Value, barcodename)
    '                                    If BARCODE <> "" Then
    '                                        BARCODE = BARCODE
    '                                        oItem.BarCodes.BarCode = BARCODE
    '                                        oItem.BarCodes.FreeText = barcodename
    '                                        oItem.BarCodes.UoMEntry = barcodeuom
    '                                        oItem.BarCodes.SetCurrentLine(p)
    '                                        oItem.BarCodes.Add()
    '                                        StrQry = "UPDATE ""@TNX_ITEMUP_C2"" SET ""U_BarCode"" ='" & BARCODE & "' Where ""DocEntry""='" & barcoders.Fields.Item("DocEntry").Value & "'and ""LineId""='" & LineId & "' "
    '                                        oRec = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                                        oRec.DoQuery(StrQry)
    '                                    End If
    '                                Else
    '                                    oItem.BarCodes.BarCode = barcodest
    '                                    oItem.BarCodes.FreeText = barcodename
    '                                    oItem.BarCodes.UoMEntry = barcodeuom
    '                                    oItem.BarCodes.SetCurrentLine(p)
    '                                    oItem.BarCodes.Add()
    '                                End If
    '                                barcoders.MoveNext()
    '                            Next

    '                        End If

    '                        Dim barcodeentry As Integer

    '                        For o As Integer = 0 To oItem.BarCodes.Count - 1
    '                            oItem.BarCodes.SetCurrentLine(o)

    '                            If oItem.BarCodes.BarCode = BARCODE Then
    '                                barcodeentry = oItem.BarCodes.AbsEntry

    '                                Exit For
    '                            End If
    '                        Next
    '                        ' If barcodeentry <> 0 Then
    '                        '    oItem.BarCode = BARCODE
    '                        'Else
    '                        '    oItem.BarCode = BARCODE

    '                        'End If

    '                        oItem.Mainsupplier = CardCode1

    '                        Dim qStr2 As String = "Select ""ItmsGrpCod"" FROM ""OITB"" WHERE ""ItmsGrpNam""='" + ItemGrp + "' "
    '                        ''   qStr2 = "Select ""ItmsGrpCod"" FROM ""OITB"" WHERE ""ItmsGrpCod""='" + ItemGrp + "' "

    '                        Dim rSet2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                        rSet2.DoQuery(qStr2)
    '                        If rSet2.RecordCount > 0 Then
    '                            oItem.ItemsGroupCode = Convert.ToInt32(rSet2.Fields.Item("ItmsGrpCod").Value)
    '                        End If
    '                        ' Me.New_Barcode_Generation(oRsItm.Fields.Item("NextNumber").Value, ItemName)
    '                        Error_Code = oItem.Add
    '                        If Error_Code <> 0 Then
    '                            Dim lErrCode As Integer
    '                            Dim sErrMsg As String = objMain.objCompany.GetLastErrorDescription()

    '                            objMain.objCompany.GetLastError(lErrCode, sErrMsg)
    '                            objMain.objApplication.StatusBar.SetText(sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                            Dim qSt As String = "UPDATE ""@TNX_ITEMUP_C0"" SET""U_STATS""='" & sErrMsg.ToString.Replace("'", " ") & "' WHERE  ""LineId"" = '" & LINE & "' and ""DocEntry""='" & DOCENTRY11 & "'  "
    '                            Dim rSe As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                            rSe.DoQuery(qSt)

    '                        Else
    '                            Dim NewItemCode As String = objMain.objCompany.GetNewObjectKey.ToString()
    '                            Try
    '                                Dim qStr8 As String = "UPDATE OITM SET ""OneBOneRec""='Y' WHERE ""ItemCode""='" + NewItemCode + "'"
    '                                Dim rSet8 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                                rSet8.DoQuery(qStr8)

    '                            Catch ex As Exception
    '                                If objMain.objCompany.InTransaction Then objMain.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

    '                                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                            End Try
    '                            StrQry = "SELECT * FROM ""@NT_BRST"""
    '                            oRec = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                            oRec.DoQuery(StrQry)
    '                            If oRec.RecordCount > 0 Then
    '                                Startt = Convert.ToUInt64(oRec.Fields.Item("U_Start1").Value) + 13

    '                                ' Endd = Convert.ToUInt64(oRec.Fields.Item("U_End").Value) + 2
    '                                Dim StrQry1 As String = "UPDATE ""@NT_BRST"" SET ""U_Start1"" ='" & Startt & "',""U_End"" ='" & Endd & "' Where ""Code""='" & oRec.Fields.Item("Code").Value & "' "
    '                                Dim oRsT As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                                If defaultBARCODE = "No" Then
    '                                    oRsT.DoQuery(StrQry1)
    '                                End If


    '                            End If

    '                            Dim qSt As String = "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_ITMCD"" = '" & NewItemCode & "',""U_STATS""='Item Master is Succesfully Created' WHERE  ""LineId"" = '" & LINE & "' and ""DocEntry""='" & DOCENTRY11 & "'  "
    '                            Dim rSe As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                            rSe.DoQuery(qSt)
    '                            objMain.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
    '                            newitemcreated = NewItemCode
    '                            objMain.objApplication.StatusBar.SetText("adding default barcode..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '                            '-------
    '                            Me.defaultbarcode(NewItemCode, keycod, DocEntry)
    '                            Me.BussinesspartnerCatlogNumbercreation(NewItemCode, keycod, DocEntry, CardCode1, CatlogNumber)
    '                            objMain.objApplication.StatusBar.SetText("Item Code is Sucessfully Created as '" + NewItemCode + "' ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

    '                        End If
    '                    Catch ex As Exception
    '                        If objMain.objCompany.InTransaction Then objMain.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

    '                        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                        Dim qSt As String = "UPDATE ""@TNX_ITEMUP_C0"" SET""U_STATS""='" & ex.Message.ToString.Replace("'", " ") & "' WHERE  ""LineId"" = '" & LINE & "' and ""DocEntry""='" & DOCENTRY11 & "'  "
    '                        Dim rSe As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                        rSe.DoQuery(qSt)
    '                    End Try


    '                End If
    '                i = i
    '                rSet.MoveNext()
    '            Next

    '        End If
    '    Catch ex As Exception
    '        If objMain.objCompany.InTransaction Then objMain.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
    '        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '    End Try
    'End Sub

    Private Sub Item_Creation(ByVal DocEntry As String)
        Try
            Dim oForm As SAPbouiCOM.Form
            Dim addedGroupCode As String
            Dim newitemcreated As String
            Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("ItmDetMtx").Specific

            Dim qStr As String = "select T0.""DocEntry"",T0.""LineId"",T0.""U_KEYC"", T0.""U_ITEMN"",T0.""U_FORIE"",T0.""U_CARD"",T0.""U_ITEMG"",T0.""U_SUBGR"",T0.""U_SSGP"",T0.""U_SUGRP"",T0.""U_BRAND"" ,T0.""U_CATGO"", T0.""U_COUNT"" ,T0.""U_BUNIT"",T0.""U_SCALE"",
            T0.""U_ITEMS"", T0.""U_INVENT"",T0.""U_SALITM"" ,T0.""U_PURCH"",T0.""U_NOP"",T0.""U_BaseUt"",T0.""U_UOMD"",T0.""U_SugarTax"",T0.""U_PVAT"",T0.""U_SVAT"",T0.""U_REBEGBLE"",T0.""U_WITPO"",T0.""U_SuTaxAmt"",T0.""U_VIC"",T0.""U_MARKUP"",T0.""U_MARGIN"",T0.""U_CNum"" from ""@TNX_ITEMUP"" T inner join ""@TNX_ITEMUP_C0"" T0 on T.""DocEntry""= T0.""DocEntry"" where T.""DocNum""='" & DocEntry & "' and IFNULL(""U_STATS"",'')<>'Item Master is Succesfully Created' ORDER BY T0.""LineId""  "
            Dim rSet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rSet.DoQuery(qStr)
            Dim DOCENTRY11 As String = ""
            If rSet.RecordCount > 0 Then

                For i = 1 To rSet.RecordCount


                    objMain.objApplication.StatusBar.SetText("Creating Items " & i & " out of " & rSet.RecordCount & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                    Dim Itemcodeas As String = objMatrix.Columns.Item("ITMCD").Cells.Item(i).Specific.Value.ToString()
                    '' Dim Itemcodeas As String = ""
                    Dim Ik As String = objMatrix.Columns.Item("KEYC").Cells.Item(i).Specific.Value.ToString()
                    objMain.objApplication.StatusBar.SetText("Creating Item for key code --, " & Itemcodeas & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Dim LINE As Integer
                    If Itemcodeas = "" And Ik <> "" Then
                        Try
                            ' objMain.objCompany.StartTransaction()
                            oItem = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

                            Dim ItemCode As String = objMatrix.Columns.Item("KEYC").Cells.Item(i).Specific.Value.ToString()
                            Dim ItemName As String = objMatrix.Columns.Item("ITEMN").Cells.Item(i).Specific.Value.ToString()
                            Dim CardCode As String = objMatrix.Columns.Item("CARD").Cells.Item(i).Specific.Value.ToString()
                            Dim FORIE As String = objMatrix.Columns.Item("FORIE").Cells.Item(i).Specific.Value.ToString()
                            Dim FItemCode As String
                            FItemCode = ItemCode
                            DOCENTRY11 = rSet.Fields.Item("DocEntry").Value.ToString()

                            LINE = rSet.Fields.Item("LineId").Value.ToString()
                            Dim keycod As String = rSet.Fields.Item("U_KEYC").Value.ToString()
                            Dim ItemNme As String = rSet.Fields.Item("U_ITEMN").Value.ToString()
                            Dim ForineName As String = rSet.Fields.Item("U_FORIE").Value.ToString()
                            Dim CardCode1 As String = rSet.Fields.Item("U_CARD").Value.ToString()
                            Dim ItemGrp As String = rSet.Fields.Item("U_ITEMG").Value.ToString()
                            Dim SubGrp As String = rSet.Fields.Item("U_SUBGR").Value.ToString()
                            Dim SubSubgroup1 As String = rSet.Fields.Item("U_SSGP").Value.ToString()
                            Dim Brand As String = rSet.Fields.Item("U_BRAND").Value.ToString()
                            Dim Category1 As String = rSet.Fields.Item("U_CATGO").Value.ToString()
                            Dim ScaleItem1 As String = rSet.Fields.Item("U_SCALE").Value.ToString()
                            Dim CountryOfOrgin1 As String = rSet.Fields.Item("U_COUNT").Value.ToString()
                            Dim BaseUnit1 As String = rSet.Fields.Item("U_BUNIT").Value.ToString()
                            Dim ItemSize1 As String = rSet.Fields.Item("U_ITEMS").Value.ToString()
                            Dim InventoryItem1 As String = rSet.Fields.Item("U_INVENT").Value.ToString()
                            Dim SaleItem1 As String = rSet.Fields.Item("U_SALITM").Value.ToString()
                            Dim PurchaseItem As String = rSet.Fields.Item("U_PURCH").Value.ToString()
                            Dim subsubgrp As String = rSet.Fields.Item("U_SUGRP").Value.ToString()
                            Dim Noofpcs As String = rSet.Fields.Item("U_NOP").Value.ToString()
                            Dim BaseUnit As String = rSet.Fields.Item("U_BaseUt").Value.ToString()
                            Dim UomDimension As String = rSet.Fields.Item("U_UOMD").Value.ToString()

                            Dim SugarTax As String = rSet.Fields.Item("U_SugarTax").Value.ToString()
                            Dim SugarTaxAmount As String = rSet.Fields.Item("U_SuTaxAmt").Value.ToString()
                            Dim VendorItemCode As String = rSet.Fields.Item("U_VIC").Value.ToString()
                            Dim Markup As String = rSet.Fields.Item("U_MARKUP").Value.ToString()
                            Dim Margin As String = rSet.Fields.Item("U_MARGIN").Value.ToString()
                            Dim REBATE As String = rSet.Fields.Item("U_REBEGBLE").Value.ToString()
                            Dim WithPo As String = rSet.Fields.Item("U_WITPO").Value.ToString()
                            Dim CatlogNumber As String = rSet.Fields.Item("U_CNum").Value.ToString()
                            'Dim seris As String = "Select ""U_SERIES"" as ""Series"" FROM ""OADM"" "
                            Dim SALESVAT As String = rSet.Fields.Item("U_SVAT").Value.ToString()
                            Dim PURCHASEVAT As String = rSet.Fields.Item("U_PVAT").Value.ToString()

                            Dim seris As String = "Select ""U_ISCODE"" as ""Series"" FROM ""OITB""  where ""ItmsGrpNam""='" & ItemGrp & "' "

                            Dim rSeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            rSeries.DoQuery(seris)
                            If rSeries.RecordCount = 0 Then
                                objMain.objApplication.StatusBar.SetText("Series not defined for Item Group.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Dim msg As String = "Series not defined for Item Group."

                                Dim qSt As String = "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_STATS""='" & msg.Replace("'", " ") &
                    "' WHERE ""LineId""='" & LINE & "' AND ""DocEntry""='" & DOCENTRY11 & "'"
                                Dim ERR As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                'BubbleEvent = False
                                ERR.DoQuery(qSt)
                                Continue For
                            End If
                            If rSeries.RecordCount > 0 And rSeries.Fields.Item("Series").Value IsNot Nothing AndAlso rSeries.Fields.Item("Series").Value.ToString() <> "" AndAlso Convert.ToInt32(rSeries.Fields.Item("Series").Value) <> 0 Then
                                oItem.Series = Convert.ToInt32(rSeries.Fields.Item("Series").Value)

                            Else

                                '   objMain.objApplication.StatusBar.SetText("please defined in OITM table series with same in OITB table Item Series Code and In OITB table Item Group Name defined same in Item Master upload screen Item Group along with in item master data item group frop down values is same...!.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objMain.objApplication.StatusBar.SetText("Please add the series to the Item Group and ensure it matches the Item Series Code in OITB and the Item Group in Item Master....!.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Dim msg As String = "Please add the series to the Item Group and ensure it matches the Item Series Code in OITB and the Item Group in Item Master."

                                Dim qSt As String = "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_STATS""='" & msg.Replace("'", " ") &
                    "' WHERE ""LineId""='" & LINE & "' AND ""DocEntry""='" & DOCENTRY11 & "'"
                                Dim ERR As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                'BubbleEvent = False
                                ERR.DoQuery(qSt)
                                ' Exit Sub
                                Continue For
                            End If

                            Dim obicod As String = "Select ""ObjectCode"" FROM ""NNM1"" where ""Series"" = '" & rSeries.Fields.Item("Series").Value & "' "
                            Dim robj As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            robj.DoQuery(obicod)
                            Dim objcode As String = ""
                            If robj.RecordCount > 0 Then
                                objcode = robj.Fields.Item("ObjectCode").Value.ToString()
                            Else
                                ' Handle error
                                Continue For
                            End If
                            Dim ItmIO As String = " Select  (""NextNumber"") as ""NextNumber""  from ""NNM1"" where ""ObjectCode""='" & objcode & "' and ""Series"" = '" & rSeries.Fields.Item("Series").Value & "' "
                            'Dim ItmIO As String = " Select  (""NextNumber"") as ""NextNumber""  from ""NNM1"" where  ""Series"" = '" & rSeries.Fields.Item("Series").Value & "' "
                            ItmIO = "SELECT ""BeginStr"" || LPAD(CAST(""NextNumber"" AS NVARCHAR), ""NumSize"", '0') || ifnull(""EndStr"",'') AS ""NextNumber"" FROM ""NNM1"" WHERE ""ObjectCode""='" & objcode & "' AND ""Series""='" & rSeries.Fields.Item("Series").Value & "'"

                            Dim oRsItm As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsItm.DoQuery(ItmIO)
                            Dim BARCODE As String
                            Dim defaultBARCODE As String
                            ''UOM creation
                            Dim oGrpParams As SAPbobsCOM.UnitOfMeasurementGroupParams
                            Dim NewItemCode1 As String = OtherCompany.GetNewObjectKey.ToString()
                            Try
                                Dim oCompService As SAPbobsCOM.CompanyService = objMain.objCompany.GetCompanyService()
                                Dim oUoMGroupService As SAPbobsCOM.UnitOfMeasurementGroupsService = CType(oCompService.GetBusinessService(SAPbobsCOM.ServiceTypes.UnitOfMeasurementGroupsService), SAPbobsCOM.UnitOfMeasurementGroupsService)
                                Dim oUoMGroup As SAPbobsCOM.UnitOfMeasurementGroup = CType(oUoMGroupService.GetDataInterface(SAPbobsCOM.UnitOfMeasurementGroupsServiceDataInterfaces.uomgsUnitOfMeasurementGroup), SAPbobsCOM.UnitOfMeasurementGroup)
                                Dim nextNumber As String = oRsItm.Fields.Item("NextNumber").Value


                                Dim oUoMGroupParams As SAPbobsCOM.UnitOfMeasurementGroupParams = CType(oUoMGroupService.GetDataInterface(SAPbobsCOM.UnitOfMeasurementGroupsServiceDataInterfaces.uomgsUnitOfMeasurementGroupParams), SAPbobsCOM.UnitOfMeasurementGroupParams)

                                oUoMGroupParams.Code = nextNumber

                                ' Try to get the UoM Group
                                Try
                                    oUoMGroupService.Delete(oUoMGroupParams)

                                Catch ex As Exception

                                End Try
                                ' If found, proceed to delete

                                Dim Uomqry As String = "select T2.""UomEntry"",""U_UOMUNT""  from ""@TNX_ITEMUP"" T0 inner join ""@TNX_ITEMUP_C1"" T1 On T0.""DocEntry""= T1.""DocEntry"" inner join ""OUOM"" T2  on T2.""UomCode""=T1.""U_UOMCDE"" " &
                                    "where T0.""DocNum""='" & DocEntry & "'and ""U_KEYCDE"" ='" & FItemCode & "' order by ""U_UOMUNT"" asc"
                                Dim rsuomqry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                rsuomqry.DoQuery(Uomqry)
                                If rsuomqry.RecordCount = 0 Then
                                    objMain.objApplication.StatusBar.SetText("UOM Code Not Found ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)


                                End If
                                If rsuomqry.RecordCount > 0 Then

                                    Dim base As Double = rsuomqry.Fields.Item("U_UOMUNT").Value
                                    Dim baseentry As Integer = rsuomqry.Fields.Item("UomEntry").Value
                                    For j As Integer = 1 To rsuomqry.RecordCount
                                        If j = 1 Then
                                            oUoMGroup.Code = nextNumber
                                            oUoMGroup.Name = nextNumber
                                            oUoMGroup.BaseUoM = baseentry
                                            rsuomqry.MoveNext()
                                        Else
                                            Dim def1 As SAPbobsCOM.UoMGroupDefinition = oUoMGroup.GroupDefinitions.Add()
                                            Dim alternate As String = rsuomqry.Fields.Item("UomEntry").Value.ToString

                                            Dim alternateQty As String = rsuomqry.Fields.Item("U_UOMUNT").Value.ToString
                                            'def1.AlternateQuantity = rsuomqry.Fields.Item("U_UOMUNT").Value
                                            'def1.BaseQuantity = base
                                            def1.AlternateQuantity = base
                                            def1.BaseQuantity = rsuomqry.Fields.Item("U_UOMUNT").Value
                                            def1.AlternateUoM = rsuomqry.Fields.Item("UomEntry").Value


                                            rsuomqry.MoveNext()
                                        End If
                                    Next
                                Else


                                End If

                                'Dim baseUoM As Integer = 1
                                'oUoMGroup.BaseUoM = baseUoM

                                'Dim def1 As SAPbobsCOM.UoMGroupDefinition
                                'def1.AlternateUoM = 1 ' Alternate UoM Code (e.g., "BOX")
                                'def1.BaseQuantity = 1 ' Conversion to Base UoM (e.g., 1 BOX = 10 PCS)
                                'def1.AlternateQuantity = 1
                                'oUoMGroup.GroupDefinitions.Add()
                                oGrpParams = oUoMGroupService.Add(oUoMGroup)
                                '    oGrpParams.
                                addedGroupCode = oGrpParams.AbsEntry
                                objMain.objApplication.StatusBar.SetText("UoM Group added successfully with Code: " & addedGroupCode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                                objMain.objApplication.StatusBar.SetText("UOM Updating in current database", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Catch ex As Exception
                                objForm.Items.Item("Item_13").Enabled = True
                                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Try

                            If objMain.objCompany.InTransaction Then objMain.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

                            objMain.objCompany.StartTransaction()
                            oItem.ItemName = ItemName
                            oItem.ForeignName = ForineName

                            'oItem.SugarTax = SugarTax
                            'oItem.SugarTaxAmount = SugarTaxAmount
                            'oItem.VendorItemCode = VendorItemCode
                            'oItem.Markup = Markup
                            'oItem.Margin = Margin

                            Dim qStr2 As String = "Select ""Code"" FROM ""OVTG"" WHERE ""Code""='" + PURCHASEVAT + "' AND ""Category"" ='I' AND ifnull(""Inactive"",'N')='N'"
                            ''   qStr2 = "Select ""ItmsGrpCod"" FROM ""OITB"" WHERE ""ItmsGrpCod""='" + ItemGrp + "' "

                            Dim rSet2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            rSet2.DoQuery(qStr2)
                            If rSet2.RecordCount > 0 Then
                                oItem.PurchaseVATGroup = (rSet2.Fields.Item("Code").Value)
                            Else
                                oItem.PurchaseVATGroup = "X1"
                            End If
                            Dim qStr21 As String = "Select ""Code"" FROM ""OVTG"" WHERE ""Code""='" + SALESVAT + "' AND ""Category"" ='O' AND ifnull(""Inactive"",'N')='N'"
                            ''   qStr2 = "Select ""ItmsGrpCod"" FROM ""OITB"" WHERE ""ItmsGrpCod""='" + ItemGrp + "' "

                            Dim rSet21 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            rSet21.DoQuery(qStr21)
                            If rSet2.RecordCount > 0 Then
                                oItem.SalesVATGroup = (rSet21.Fields.Item("Code").Value)
                            Else
                                oItem.SalesVATGroup = "X0"
                            End If
                            'oItem.SalesVATGroup = "X0"
                            'oItem.PurchaseVATGroup = "X1"
                            '  oItem.ItemsGroupCode = ItemGrp
                            'oItem.UoMGroupEntry = "399723"

                            oItem.UoMGroupEntry = addedGroupCode


                            oItem.PreferredVendors.BPCode = CardCode1
                            '  oItem.PreferredVendors.SetCurrentLine(0)
                            oItem.PreferredVendors.Add()

                            Dim Uomtest As String = "select top 1 T2.""UomEntry"",""U_UOMUNT""  from ""@TNX_ITEMUP"" T0 inner join ""@TNX_ITEMUP_C1"" T1 On T0.""DocEntry""= T1.""DocEntry"" inner join ""OUOM"" T2  on T2.""UomCode""=T1.""U_UOMCDE"" " &
                                   "where T0.""DocNum""='" & DocEntry & "'and ""U_KEYCDE"" ='" & FItemCode & "' order by ""U_UOMUNT"" asc"
                            Dim rsuomqrytest As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            rsuomqrytest.DoQuery(Uomtest)
                            Dim uom As Integer = 0
                            If rsuomqrytest.RecordCount > 0 Then
                                uom = rsuomqrytest.Fields.Item("UomEntry").Value
                                Dim base As Double = rsuomqrytest.Fields.Item("U_UOMUNT").Value
                            End If


                            'oItem.UnitOfMeasurements.UoMEntry = 93
                            'oItem.UnitOfMeasurements.
                            'oItem.UnitOfMeasurements.Add()
                            oItem.DefaultSalesUoMEntry = uom
                            ''oItem.UoMGroupEntry = oRsItm.Fields.Item("NextNumber").Value
                            oItem.InventoryUoMEntry = uom
                            oItem.SalesItem = SAPbobsCOM.BoYesNoEnum.tYES
                            oItem.PurchaseItem = SAPbobsCOM.BoYesNoEnum.tYES
                            oItem.InventoryItem = SAPbobsCOM.BoYesNoEnum.tYES
                            oItem.DefaultPurchasingUoMEntry = uom
                            oItem.UserFields.Fields.Item("U_BRAND").Value = Brand
                            oItem.UserFields.Fields.Item("U_VSubGroup").Value = SubGrp
                            oItem.UserFields.Fields.Item("U_VSubSubGroup").Value = SubSubgroup1
                            oItem.UserFields.Fields.Item("U_CATEGORY").Value = Category1
                            oItem.UserFields.Fields.Item("U_PcPerPac").Value = Noofpcs
                            oItem.UserFields.Fields.Item("U_BaseUnit").Value = BaseUnit1
                            oItem.UserFields.Fields.Item("U_ItemSize").Value = ItemSize1
                            oItem.UserFields.Fields.Item("U_CountryOrg").Value = CountryOfOrgin1
                            oItem.UserFields.Fields.Item("U_UOM").Value = UomDimension
                            'oItem.UserFields.Fields.Item("U_SugarTax").Value = SugarTax

                            If SugarTax.Trim().ToUpper() = "YES" Then
                                oItem.UserFields.Fields.Item("U_SugarTax").Value = "Y"
                            Else
                                oItem.UserFields.Fields.Item("U_SugarTax").Value = "N"
                            End If

                            If WithPo.Trim().ToUpper() = "YES" Then
                                oItem.UserFields.Fields.Item("U_WoutPo").Value = "Yes"
                            Else
                                oItem.UserFields.Fields.Item("U_WoutPo").Value = "No"
                            End If

                            oItem.UserFields.Fields.Item("U_SuTaxAmt").Value = SugarTaxAmount
                            oItem.UserFields.Fields.Item("U_REBEGBLE").Value = REBATE
                            ''oItem.UserFields.Fields.Item("U_WITPO").Value = WithPo


                            If ScaleItem1 = "Yes" Then
                                oItem.UserFields.Fields.Item("U_AcceptWMQty").Value = "1" ' Yes as string
                            Else
                                oItem.UserFields.Fields.Item("U_AcceptWMQty").Value = "2" ' No as string
                            End If
                            oItem.UserFields.Fields.Item("U_3SUBGROUP").Value = subsubgrp



                            Dim barcodestest As String = "select T2.""UomEntry"", ""U_KEYC"",""U_UOMCDE"",""U_ADESC"",""U_BarCode"",""U_BDesc"",T1.""DocEntry"",T1.""LineId"" , ""U_BarCode"", ""U_STAX"", ""U_SUTAXAMT"", ""U_ETAX"", ""U_Price"" from ""@TNX_ITEMUP"" T0 inner join ""@TNX_ITEMUP_C2"" T1 On T0.""DocEntry""= T1.""DocEntry"" inner join ""OUOM"" T2  on T2.""UomCode""=T1.""U_UOMCDE"" " &
                                   "where T0.""DocNum""='" & DocEntry & "'and ""U_KEYC"" ='" & FItemCode & "'"
                            Dim barcoders As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            barcoders.DoQuery(barcodestest)

                            If barcoders.RecordCount = 0 Then
                                defaultBARCODE = "No"

                                BARCODE = Me.New_Barcode_Generation(oRsItm.Fields.Item("NextNumber").Value, ItemName)

                                If BARCODE <> "" Then
                                    BARCODE = BARCODE
                                    oItem.BarCodes.BarCode = BARCODE
                                    oItem.BarCodes.FreeText = ItemName
                                    oItem.BarCodes.UoMEntry = uom
                                    oItem.BarCodes.SetCurrentLine(0)
                                    oItem.BarCodes.Add()
                                    'Else
                                    '    rSet.MoveNext()

                                End If

                            Else
                                'For p As Integer = 1 To barcoders.RecordCount
                                For p As Integer = 0 To barcoders.RecordCount - 1
                                    Dim barcodest As String = barcoders.Fields.Item("U_BarCode").Value
                                    Dim barcodename As String = barcoders.Fields.Item("U_BDesc").Value
                                    Dim barcodeuom As Integer = barcoders.Fields.Item("UomEntry").Value
                                    Dim LineId As Integer = barcoders.Fields.Item("LineId").Value
                                    'If barcodest <> "" And barcodename <> "" And barcodeuom <> 0 Then
                                    If barcodest = "" And barcodename <> "" And barcodeuom <> 0 Then
                                        BARCODE = Me.New_Barcode_GenerationWithBarcodeDescription(oRsItm.Fields.Item("NextNumber").Value, barcodename)
                                        If BARCODE <> "" Then
                                            BARCODE = BARCODE
                                            oItem.BarCodes.BarCode = BARCODE
                                            oItem.BarCodes.FreeText = barcodename
                                            oItem.BarCodes.UoMEntry = barcodeuom


                                            oItem.BarCodes.SetCurrentLine(p)
                                            oItem.BarCodes.Add()
                                            StrQry = "UPDATE ""@TNX_ITEMUP_C2"" SET ""U_BarCode"" ='" & BARCODE & "' Where ""DocEntry""='" & barcoders.Fields.Item("DocEntry").Value & "'and ""LineId""='" & LineId & "' "
                                            oRec = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRec.DoQuery(StrQry)
                                        End If
                                    Else
                                        oItem.BarCodes.BarCode = barcodest
                                        oItem.BarCodes.FreeText = barcodename
                                        oItem.BarCodes.UoMEntry = barcodeuom

                                        oItem.BarCodes.SetCurrentLine(p)
                                        oItem.BarCodes.Add()
                                    End If
                                    barcoders.MoveNext()
                                Next

                            End If

                            Dim barcodeentry As Integer

                            For o As Integer = 0 To oItem.BarCodes.Count - 1
                                oItem.BarCodes.SetCurrentLine(o)

                                If oItem.BarCodes.BarCode = BARCODE Then
                                    barcodeentry = oItem.BarCodes.AbsEntry

                                    Exit For
                                End If
                            Next

                            oItem.Mainsupplier = CardCode1
                            '--- ADD THIS BLOCK HERE ---
                            If VendorItemCode <> "" Then
                                oItem.SupplierCatalogNo = VendorItemCode
                            End If

                            Dim qStr2X As String = "Select ""ItmsGrpCod"" FROM ""OITB"" WHERE ""ItmsGrpNam""='" + ItemGrp + "' "
                            ''   qStr2 = "Select ""ItmsGrpCod"" FROM ""OITB"" WHERE ""ItmsGrpCod""='" + ItemGrp + "' "

                            Dim rSet2X As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            rSet2X.DoQuery(qStr2X)
                            If rSet2X.RecordCount > 0 Then
                                oItem.ItemsGroupCode = Convert.ToInt32(rSet2X.Fields.Item("ItmsGrpCod").Value)
                            End If
                            ' Me.New_Barcode_Generation(oRsItm.Fields.Item("NextNumber").Value, ItemName)
                            Error_Code = oItem.Add
                            If Error_Code <> 0 Then
                                Dim lErrCode As Integer
                                Dim sErrMsg As String = objMain.objCompany.GetLastErrorDescription()

                                objMain.objCompany.GetLastError(lErrCode, sErrMsg)
                                objMain.objApplication.StatusBar.SetText(sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Dim qSt As String = "UPDATE ""@TNX_ITEMUP_C0"" SET""U_STATS""='" & sErrMsg.ToString.Replace("'", " ") & "' WHERE  ""LineId"" = '" & LINE & "' and ""DocEntry""='" & DOCENTRY11 & "'  "
                                Dim rSe As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                rSe.DoQuery(qSt)

                            Else
                                Dim NewItemCode As String = objMain.objCompany.GetNewObjectKey.ToString()
                                Try
                                    Dim qStr8 As String = "UPDATE OITM SET ""OneBOneRec""='Y' WHERE ""ItemCode""='" + NewItemCode + "'"
                                    Dim rSet8 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    rSet8.DoQuery(qStr8)

                                Catch ex As Exception
                                    objForm.Items.Item("Item_13").Enabled = True
                                    If objMain.objCompany.InTransaction Then objMain.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)

                                    objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End Try
                                StrQry = "SELECT * FROM ""@NT_BRST"""
                                oRec = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRec.DoQuery(StrQry)
                                If oRec.RecordCount > 0 Then
                                    Startt = Convert.ToUInt64(oRec.Fields.Item("U_Start1").Value) + 13

                                    ' Endd = Convert.ToUInt64(oRec.Fields.Item("U_End").Value) + 2
                                    Dim StrQry1 As String = "UPDATE ""@NT_BRST"" SET ""U_Start1"" ='" & Startt & "',""U_End"" ='" & Endd & "' Where ""Code""='" & oRec.Fields.Item("Code").Value & "' "
                                    Dim oRsT As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    If defaultBARCODE = "No" Then
                                        oRsT.DoQuery(StrQry1)
                                    End If


                                End If

                                Dim qSt As String = "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_ITMCD"" = '" & NewItemCode & "',""U_STATS""='Item Master is Succesfully Created' WHERE  ""LineId"" = '" & LINE & "' and ""DocEntry""='" & DOCENTRY11 & "'  "
                                Dim rSe As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                rSe.DoQuery(qSt)
                                objMain.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                newitemcreated = NewItemCode
                                objMain.objApplication.StatusBar.SetText("adding default barcode..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                '-------
                                Me.defaultbarcode(NewItemCode, keycod, DocEntry)

                                ''Me.BussinesspartnerCatlogNumbercreation(NewItemCode, keycod, DocEntry, CardCode1, CatlogNumber)
                                ''Dim Barcode As String = objForm.Items.Item("DocEntry").Specific.Value

                                Dim CAT As String = "select ""LineId"",""U_CNum"",""U_BarCode"" from ""@TNX_ITEMUP"" T0 inner join ""@TNX_ITEMUP_C2"" T1 On T0.""DocEntry""= T1.""DocEntry"" " &
                                   "where T0.""DocNum""='" & DocEntry & "'and ""U_KEYC"" ='" & FItemCode & "'"
                                Dim ORSCAT As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                ORSCAT.DoQuery(CAT)

                                For p As Integer = 1 To ORSCAT.RecordCount

                                    Dim BarcodeNew As String = ORSCAT.Fields.Item("U_BarCode").Value
                                    Dim CATNUM As String = ORSCAT.Fields.Item("U_CNum").Value
                                    Dim LineID As String = ORSCAT.Fields.Item("LineId").Value

                                    Me.BussinesspartnerCatlogNumbercreation(NewItemCode, keycod, DOCENTRY11, CardCode1, BarcodeNew, CATNUM, LineID)
                                    ORSCAT.MoveNext()
                                Next

                                objMain.objApplication.StatusBar.SetText("Item Code is Sucessfully Created as '" + NewItemCode + "' ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                            End If
                        Catch ex As Exception
                            If objMain.objCompany.InTransaction Then objMain.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            objForm.Items.Item("Item_13").Enabled = True
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Dim qSt As String = "UPDATE ""@TNX_ITEMUP_C0"" SET""U_STATS""='" & ex.Message.ToString.Replace("'", " ") & "' WHERE  ""LineId"" = '" & LINE & "' and ""DocEntry""='" & DOCENTRY11 & "'  "
                            Dim rSe As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            rSe.DoQuery(qSt)
                        End Try


                    End If
                    i = i
                    rSet.MoveNext()
                Next


                UpdateWarehouseBlockPurchase(DocEntry)
                Me.SyncBarcodeUDFs(DocEntry)
                Me.UpdateSugarTax_UOMWise(DocEntry)
            End If
        Catch ex As Exception
            objForm.Items.Item("Item_13").Enabled = True
            If objMain.objCompany.InTransaction Then objMain.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
    End Sub


    ' FULL method + helpers: Item_Creationworking + Create_UOM_Group + AssignBarCodes
    Private Sub Item_Creation1(ByVal DocEntry As String)
        Try
            Dim rSet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim qStr As String =
            "SELECT T0.""DocEntry"", T0.""LineId"", T0.""U_KEYC"", T0.""U_ITEMN"", T0.""U_FORIE"", T0.""U_CARD"", " &
            "T0.""U_ITEMG"", T0.""U_SUBGR"", T0.""U_SSGP"", T0.""U_SUGRP"", T0.""U_BRAND"", T0.""U_CATGO"", " &
            "T0.""U_COUNT"", T0.""U_BUNIT"", T0.""U_SCALE"", T0.""U_ITEMS"", T0.""U_INVENT"", T0.""U_SALITM"", " &
            "T0.""U_PURCH"", T0.""U_NOP"", T0.""U_BaseUt"", T0.""U_UOMD"", T0.""U_CNum"" " &
            "FROM ""@TNX_ITEMUP"" T INNER JOIN ""@TNX_ITEMUP_C0"" T0 ON T.""DocEntry"" = T0.""DocEntry"" " &
            "WHERE T.""DocNum"" = '" & DocEntry & "' AND IFNULL(T0.""U_STATS"",'') <> 'Item Master is Succesfully Created' " &
            "ORDER BY T0.""LineId"" "

            rSet.DoQuery(qStr)

            If rSet.RecordCount = 0 Then
                objMain.objApplication.StatusBar.SetText("No pending item lines found for DocNum: " & DocEntry, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Exit Sub
            End If

            rSet.MoveFirst()

            For rowIndex As Integer = 1 To rSet.RecordCount
                Try
                    objMain.objApplication.StatusBar.SetText("Processing item line " & rowIndex & " of " & rSet.RecordCount, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                    ' Read values directly from recordset (fixes mismatch with matrix index)
                    Dim DB_DocEntry As String = If(rSet.Fields.Item("DocEntry").Value IsNot Nothing, rSet.Fields.Item("DocEntry").Value.ToString(), "")
                    Dim DB_LineId As String = If(rSet.Fields.Item("LineId").Value IsNot Nothing, rSet.Fields.Item("LineId").Value.ToString(), "")
                    Dim keyCode_db As String = If(rSet.Fields.Item("U_KEYC").Value IsNot Nothing, rSet.Fields.Item("U_KEYC").Value.ToString(), "").Trim()
                    Dim itemName_db As String = If(rSet.Fields.Item("U_ITEMN").Value IsNot Nothing, rSet.Fields.Item("U_ITEMN").Value.ToString(), "").Trim()
                    Dim forie_db As String = If(rSet.Fields.Item("U_FORIE").Value IsNot Nothing, rSet.Fields.Item("U_FORIE").Value.ToString(), "").Trim()
                    Dim cardcode_db As String = If(rSet.Fields.Item("U_CARD").Value IsNot Nothing, rSet.Fields.Item("U_CARD").Value.ToString(), "").Trim()
                    Dim itemGrp As String = If(rSet.Fields.Item("U_ITEMG").Value IsNot Nothing, rSet.Fields.Item("U_ITEMG").Value.ToString(), "").Trim()
                    Dim subGrp As String = If(rSet.Fields.Item("U_SUBGR").Value IsNot Nothing, rSet.Fields.Item("U_SUBGR").Value.ToString(), "").Trim()
                    Dim subsub As String = If(rSet.Fields.Item("U_SSGP").Value IsNot Nothing, rSet.Fields.Item("U_SSGP").Value.ToString(), "").Trim()
                    Dim brand As String = If(rSet.Fields.Item("U_BRAND").Value IsNot Nothing, rSet.Fields.Item("U_BRAND").Value.ToString(), "").Trim()
                    Dim category1 As String = If(rSet.Fields.Item("U_CATGO").Value IsNot Nothing, rSet.Fields.Item("U_CATGO").Value.ToString(), "").Trim()
                    Dim scaleItem As String = If(rSet.Fields.Item("U_SCALE").Value IsNot Nothing, rSet.Fields.Item("U_SCALE").Value.ToString(), "").Trim()
                    Dim countryOrigin As String = If(rSet.Fields.Item("U_COUNT").Value IsNot Nothing, rSet.Fields.Item("U_COUNT").Value.ToString(), "").Trim()
                    Dim baseUnit1 As String = If(rSet.Fields.Item("U_BUNIT").Value IsNot Nothing, rSet.Fields.Item("U_BUNIT").Value.ToString(), "").Trim()
                    Dim itemSize1 As String = If(rSet.Fields.Item("U_ITEMS").Value IsNot Nothing, rSet.Fields.Item("U_ITEMS").Value.ToString(), "").Trim()
                    Dim inventoryItem1 As String = If(rSet.Fields.Item("U_INVENT").Value IsNot Nothing, rSet.Fields.Item("U_INVENT").Value.ToString(), "").Trim()
                    Dim saleItem1 As String = If(rSet.Fields.Item("U_SALITM").Value IsNot Nothing, rSet.Fields.Item("U_SALITM").Value.ToString(), "").Trim()
                    Dim purchaseItem As String = If(rSet.Fields.Item("U_PURCH").Value IsNot Nothing, rSet.Fields.Item("U_PURCH").Value.ToString(), "").Trim()
                    Dim noOfPcs As String = If(rSet.Fields.Item("U_NOP").Value IsNot Nothing, rSet.Fields.Item("U_NOP").Value.ToString(), "").Trim()
                    Dim baseUnit As String = If(rSet.Fields.Item("U_BaseUt").Value IsNot Nothing, rSet.Fields.Item("U_BaseUt").Value.ToString(), "").Trim()
                    Dim uomDimension As String = If(rSet.Fields.Item("U_UOMD").Value IsNot Nothing, rSet.Fields.Item("U_UOMD").Value.ToString(), "").Trim()
                    Dim catalogNumber As String = If(rSet.Fields.Item("U_CNum").Value IsNot Nothing, rSet.Fields.Item("U_CNum").Value.ToString(), "").Trim()

                    ' If keyCode is empty — skip (can't create item without ItemCode)
                    If String.IsNullOrEmpty(keyCode_db) Then
                        ' update status to show missing key code
                        Dim updMissingKey As String = "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_STATS""='Skipped - Missing KeyCode' WHERE ""DocEntry""='" & DB_DocEntry & "' AND ""LineId""='" & DB_LineId & "'"
                        Dim rsUpdMissing As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        rsUpdMissing.DoQuery(updMissingKey)
                        objMain.objApplication.StatusBar.SetText("Skipped a line due to missing KeyCode. DocEntry:" & DB_DocEntry & " LineId:" & DB_LineId, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        rSet.MoveNext()
                        Continue For
                    End If

                    ' Check whether item already exists in OITM by ItemCode (if it exists, just update status)
                    Dim chk As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    chk.DoQuery("SELECT ""ItemCode"" FROM OITM WHERE ""ItemCode"" = '" & keyCode_db.Replace("'", "''") & "'")
                    If chk.RecordCount > 0 Then
                        ' Already exists: update status
                        Dim updExist As String = "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_ITMCD"" = '" & keyCode_db.Replace("'", "''") & "', ""U_STATS""='Item Master is Succesfully Created' WHERE ""DocEntry""='" & DB_DocEntry & "' AND ""LineId""='" & DB_LineId & "'"
                        Dim rsUpdExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        rsUpdExist.DoQuery(updExist)
                        objMain.objApplication.StatusBar.SetText("Item already exists: " & keyCode_db & ". Status updated.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        rSet.MoveNext()
                        Continue For
                    End If

                    ' Prepare UoM Group creation & get group entry
                    Dim addedGroupCode As Integer = 0
                    Try
                        addedGroupCode = Create_UOM_Group(DocEntry, keyCode_db)
                    Catch exU As Exception
                        addedGroupCode = 0
                        objMain.objApplication.StatusBar.SetText("UoM creation error for " & keyCode_db & ": " & exU.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try

                    ' Build Item object and set details
                    Dim oItem As SAPbobsCOM.Items = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                    ' If your environment expects to set Series or other fields use existing logic from your code:
                    Try
                        ' Optional: set Series if you use series based numbering (preserve your earlier logic)
                        Dim serisSQL As String = "Select ""U_SERIES"" as ""Series"" FROM ""OADM"""
                        Dim rSeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        rSeries.DoQuery(serisSQL)
                        If rSeries.RecordCount > 0 AndAlso rSeries.Fields.Item("Series").Value IsNot Nothing AndAlso rSeries.Fields.Item("Series").Value.ToString() <> "" AndAlso Convert.ToInt32(rSeries.Fields.Item("Series").Value) <> 0 Then
                            oItem.Series = Convert.ToInt32(rSeries.Fields.Item("Series").Value)
                        End If
                    Catch exS As Exception
                        ' ignore series issues, continue
                    End Try

                    ' Assign properties (keep naming consistent with your old code)
                    oItem.ItemCode = keyCode_db
                    oItem.ItemName = itemName_db
                    oItem.ForeignName = forie_db
                    oItem.SalesVATGroup = "X0"
                    oItem.PurchaseVATGroup = "X1"

                    If Not String.IsNullOrEmpty(cardcode_db) Then
                        oItem.PreferredVendors.BPCode = cardcode_db
                        oItem.PreferredVendors.Add()
                    End If

                    If addedGroupCode <> 0 Then
                        oItem.UoMGroupEntry = addedGroupCode
                    End If

                    ' Determine UoM (pick top UoM entry if present)
                    Dim uomEntry As Integer = 0
                    Try
                        Dim uq As String = "select top 1 T2.""UomEntry"", T1.""U_UOMUNT"" from ""@TNX_ITEMUP"" T0 inner join ""@TNX_ITEMUP_C1"" T1 On T0.""DocEntry""= T1.""DocEntry"" inner join ""OUOM"" T2 on T2.""UomCode""=T1.""U_UOMCDE"" where T0.""DocNum""='" & DocEntry & "' and ""U_KEYCDE"" ='" & keyCode_db.Replace("'", "''") & "' order by ""U_UOMUNT"" asc"
                        Dim rsu As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        rsu.DoQuery(uq)
                        If rsu.RecordCount > 0 Then
                            uomEntry = Convert.ToInt32(rsu.Fields.Item("UomEntry").Value)
                        End If
                    Catch exUo As Exception
                        uomEntry = 0
                    End Try

                    If uomEntry <> 0 Then
                        oItem.DefaultSalesUoMEntry = uomEntry
                        oItem.InventoryUoMEntry = uomEntry
                        oItem.DefaultPurchasingUoMEntry = uomEntry
                    End If

                    ' Basic flags
                    oItem.SalesItem = SAPbobsCOM.BoYesNoEnum.tYES
                    oItem.PurchaseItem = SAPbobsCOM.BoYesNoEnum.tYES
                    oItem.InventoryItem = SAPbobsCOM.BoYesNoEnum.tYES

                    ' Map UDFs (preserve original names)
                    Try
                        oItem.UserFields.Fields.Item("U_BRAND").Value = brand
                        oItem.UserFields.Fields.Item("U_VSubGroup").Value = subGrp
                        oItem.UserFields.Fields.Item("U_VSubSubGroup").Value = subsub
                        oItem.UserFields.Fields.Item("U_CATEGORY").Value = category1
                        oItem.UserFields.Fields.Item("U_PcPerPac").Value = noOfPcs
                        oItem.UserFields.Fields.Item("U_BaseUnit").Value = baseUnit1
                        oItem.UserFields.Fields.Item("U_ItemSize").Value = itemSize1
                        oItem.UserFields.Fields.Item("U_CountryOrg").Value = countryOrigin
                        oItem.UserFields.Fields.Item("U_UOM").Value = uomDimension
                        If LCase(scaleItem) = "yes" Then
                            oItem.UserFields.Fields.Item("U_AcceptWMQty").Value = "1"
                        Else
                            oItem.UserFields.Fields.Item("U_AcceptWMQty").Value = "2"
                        End If
                        oItem.UserFields.Fields.Item("U_3SUBGROUP").Value = subsub
                    Catch exUDF As Exception
                        ' continue even if UDFs fail
                    End Try

                    ' Start transaction, add item, and commit (preserve transaction safety)
                    Try
                        If Not objMain.objCompany.InTransaction Then
                            objMain.objCompany.StartTransaction()
                        End If

                        Dim Error_Code As Integer = oItem.Add()
                        If Error_Code <> 0 Then
                            Dim lErrCode As Integer
                            Dim sErrMsg As String = ""
                            objMain.objCompany.GetLastError(lErrCode, sErrMsg)
                            ' Update child table with error
                            Dim qStErr As String = "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_STATS""='" & sErrMsg.Replace("'", " ") & "' WHERE ""DocEntry""='" & DB_DocEntry & "' AND ""LineId""='" & DB_LineId & "'"
                            Dim rSeErr As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            rSeErr.DoQuery(qStErr)
                            ' rollback
                            If objMain.objCompany.InTransaction Then objMain.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            objMain.objApplication.StatusBar.SetText("Error creating Item " & keyCode_db & ": " & sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            rSet.MoveNext()
                            Continue For
                        Else
                            ' Commit transaction
                            If objMain.objCompany.InTransaction Then objMain.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        End If
                    Catch exAdd As Exception
                        If objMain.objCompany.InTransaction Then objMain.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        Dim qStErr2 As String = "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_STATS""='" & exAdd.Message.Replace("'", " ") & "' WHERE ""DocEntry""='" & DB_DocEntry & "' AND ""LineId""='" & DB_LineId & "'"
                        Dim rSeErr2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        rSeErr2.DoQuery(qStErr2)
                        objMain.objApplication.StatusBar.SetText("Exception adding item " & keyCode_db & ": " & exAdd.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        rSet.MoveNext()
                        Continue For
                    End Try

                    ' If we reach here, Item was added successfully. Get new item code (system generated or same code)
                    Dim newItemCode As String = keyCode_db
                    Try
                        newItemCode = objMain.objCompany.GetNewObjectKey().ToString()
                    Catch exGet As Exception
                        ' ignore; we'll still use the keyCode_db
                        newItemCode = keyCode_db
                    End Try

                    ' Add default barcode(s) and update child barcode table
                    Try
                        AssignBarCodes(oItem, DocEntry, keyCode_db)
                    Catch exBar As Exception
                        ' update status (non-fatal)
                        objMain.objApplication.StatusBar.SetText("Barcode assign error for " & keyCode_db & ": " & exBar.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try

                    ' Run post-create updates similar to your original code
                    Try
                        Dim qStr8 As String = "UPDATE OITM SET ""OneBOneRec""='Y' WHERE ""ItemCode""='" & newItemCode.Replace("'", "''") & "'"
                        Dim rSet8 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        rSet8.DoQuery(qStr8)
                    Catch exUpd As Exception
                        ' ignore
                    End Try

                    ' Update status to success
                    Dim qSt As String = "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_ITMCD"" = '" & newItemCode.Replace("'", "''") & "', ""U_STATS""='Item Master is Succesfully Created' WHERE ""DocEntry""='" & DB_DocEntry & "' AND ""LineId""='" & DB_LineId & "'"
                    Dim rSe As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    rSe.DoQuery(qSt)

                    ' Optionally call your extra helper functions if you use them (kept names to match your code)
                    Try
                        Me.defaultbarcode(newItemCode, keyCode_db, DocEntry)
                    Catch ex As Exception
                        ' ignore if not exist
                    End Try
                    Try
                        ''Me.BussinesspartnerCatlogNumbercreation(newItemCode, keyCode_db, DocEntry, cardcode_db, catalogNumber)
                        '''Dim DocEntry1 As String = objForm.Items.Item("DocEntry").Specific.Value
                        ''Me.BussinesspartnerCatlogNumbercreation(newItemCode, keyCode_db, DB_DocEntry, cardcode_db)

                        Dim CAT As String = "select * from ""@TNX_ITEMUP"" T0 inner join ""@TNX_ITEMUP_C2"" T1 On T0.""DocEntry""= T1.""DocEntry"" " &
                                   "where T0.""DocNum""='" & DocEntry & "'and ""U_KEYC"" ='" & keyCode_db & "'"
                        Dim ORSCAT As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ORSCAT.DoQuery(CAT)

                        For p As Integer = 1 To ORSCAT.RecordCount

                            Dim BarcodeNew As String = ORSCAT.Fields.Item("U_BarCode").Value

                            ''Me.BussinesspartnerCatlogNumbercreation(newItemCode, keyCode_db, DB_DocEntry, cardcode_db, BarcodeNew)
                        Next



                    Catch ex As Exception
                        ' ignore if not exist
                    End Try

                    objMain.objApplication.StatusBar.SetText("Item Created: " & newItemCode & " for Key: " & keyCode_db, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                Catch exLine As Exception
                    ' Per-line exception handling, update child row and continue
                    Try
                        Dim DB_DocEntry_E As String = If(rSet.Fields.Item("DocEntry").Value IsNot Nothing, rSet.Fields.Item("DocEntry").Value.ToString(), "")
                        Dim DB_LineId_E As String = If(rSet.Fields.Item("LineId").Value IsNot Nothing, rSet.Fields.Item("LineId").Value.ToString(), "")
                        Dim qStErr As String = "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_STATS""='" & exLine.Message.Replace("'", " ") & "' WHERE ""DocEntry""='" & DB_DocEntry_E & "' AND ""LineId""='" & DB_LineId_E & "'"
                        Dim rSeErr As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        rSeErr.DoQuery(qStErr)
                    Catch exUpdEx As Exception
                        ' ignore
                    End Try
                    objMain.objApplication.StatusBar.SetText("Error processing row: " & exLine.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Finally
                    rSet.MoveNext()
                End Try
            Next

            objMain.objApplication.StatusBar.SetText("Item creation run completed for DocNum: " & DocEntry, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            If objMain.objCompany.InTransaction Then objMain.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            objMain.objApplication.StatusBar.SetText("Fatal: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    ' ============================
    ' Helper: Create_UOM_Group
    ' - Creates UoM group from @TNX_ITEMUP_C1 entries for the given DocNum & KeyCode
    ' - Returns AbsEntry (Integer) of created group or 0 on failure / not found
    ' ============================
    Private Function Create_UOM_Group(ByVal DocNum As String, ByVal KeyCode As String) As Integer
        Try
            Dim rsuomqry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim Uomqry As String = "select T2.""UomEntry"",T1.""U_UOMUNT"" from ""@TNX_ITEMUP"" T0 inner join ""@TNX_ITEMUP_C1"" T1 On T0.""DocEntry""= T1.""DocEntry"" inner join ""OUOM"" T2 on T2.""UomCode""=T1.""U_UOMCDE"" " &
                                "where T0.""DocNum""='" & DocNum.Replace("'", "''") & "' and ""U_KEYCDE"" ='" & KeyCode.Replace("'", "''") & "' order by ""U_UOMUNT"" asc"
            rsuomqry.DoQuery(Uomqry)
            If rsuomqry.RecordCount = 0 Then
                Return 0
            End If

            Dim compService As SAPbobsCOM.CompanyService = objMain.objCompany.GetCompanyService()
            Dim uomSvc As SAPbobsCOM.UnitOfMeasurementGroupsService = CType(compService.GetBusinessService(SAPbobsCOM.ServiceTypes.UnitOfMeasurementGroupsService), SAPbobsCOM.UnitOfMeasurementGroupsService)
            Dim uomGroup As SAPbobsCOM.UnitOfMeasurementGroup = CType(uomSvc.GetDataInterface(SAPbobsCOM.UnitOfMeasurementGroupsServiceDataInterfaces.uomgsUnitOfMeasurementGroup), SAPbobsCOM.UnitOfMeasurementGroup)

            ' Use NextNumber from NNM1 if needed - keep your original behavior lightly simplified:
            Dim nn As String = Now.Ticks.ToString() ' fallback unique code if NextNumber not desired
            Try
                ' Try existing logic to get NextNumber for items (if you want to match your old code)
                Dim seris As String = "Select ""U_SERIES"" as ""Series"" FROM ""OADM"""
                Dim rSeries As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rSeries.DoQuery(seris)
                If rSeries.RecordCount > 0 AndAlso rSeries.Fields.Item("Series").Value IsNot Nothing Then
                    Dim obicod As String = "Select ""ObjectCode"" FROM ""NNM1"" where ""Series"" = '" & rSeries.Fields.Item("Series").Value & "' "
                    Dim robj As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    robj.DoQuery(obicod)
                    If robj.RecordCount > 0 Then
                        Dim objcode As String = robj.Fields.Item("ObjectCode").Value
                        Dim ItmIO As String = " Select (""NextNumber"") as ""NextNumber"" from ""NNM1"" where ""ObjectCode""='" & objcode & "' and ""Series"" = '" & rSeries.Fields.Item("Series").Value & "'"
                        Dim oRsItm As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsItm.DoQuery(ItmIO)
                        If oRsItm.RecordCount > 0 Then
                            nn = oRsItm.Fields.Item("NextNumber").Value.ToString()
                        End If
                    End If
                End If
            Catch exnn As Exception
                ' fallback to tick value above
            End Try

            uomGroup.Code = nn
            uomGroup.Name = nn

            rsuomqry.MoveFirst()
            ' base UoM is first (smallest) - matches your earlier logic ordering by U_UOMUNT asc
            Dim baseEntry As Integer = Convert.ToInt32(rsuomqry.Fields.Item("UomEntry").Value)
            Dim baseQty As Double = Convert.ToDouble(rsuomqry.Fields.Item("U_UOMUNT").Value)
            uomGroup.BaseUoM = baseEntry
            rsuomqry.MoveNext()

            While rsuomqry.EoF = False
                Dim def As SAPbobsCOM.UoMGroupDefinition = uomGroup.GroupDefinitions.Add()
                def.AlternateUoM = rsuomqry.Fields.Item("UomEntry").Value
                def.AlternateQuantity = baseQty
                def.BaseQuantity = rsuomqry.Fields.Item("U_UOMUNT").Value
                rsuomqry.MoveNext()
            End While

            Dim grpParams As SAPbobsCOM.UnitOfMeasurementGroupParams = uomSvc.Add(uomGroup)
            If grpParams IsNot Nothing Then
                Return Convert.ToInt32(grpParams.AbsEntry)
            Else
                Return 0
            End If

        Catch ex As Exception
            ' return 0 on any error to continue item creation without UoM group
            'objMain.objApplication.StatusBar.SetText("Create_UOM_Group error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoMessageType.smt_Warning)
            Return 0
        End Try
    End Function

    ' ============================
    ' Helper: AssignBarCodes
    ' - Adds barcode rows to the created oItem based on @TNX_ITEMUP_C2 child rows
    ' - Updates the child table with generated barcode if it was generated here
    ' ============================
    Private Sub AssignBarCodes(ByRef oItem As SAPbobsCOM.Items, ByVal DocNum As String, ByVal KeyCode As String)
        Try
            Dim barcodestest As String = "select T2.""UomEntry"", T1.""U_KEYC"", T1.""U_UOMCDE"", T1.""U_BarCode"", T1.""U_BDesc"", T1.""DocEntry"", T1.""LineId"" " &
                                     "from ""@TNX_ITEMUP"" T0 inner join ""@TNX_ITEMUP_C2"" T1 On T0.""DocEntry""= T1.""DocEntry"" inner join ""OUOM"" T2 on T2.""UomCode""=T1.""U_UOMCDE"" " &
                                     "where T0.""DocNum""='" & DocNum.Replace("'", "''") & "' and ""U_KEYC"" ='" & KeyCode.Replace("'", "''") & "'"

            Dim barcoders As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            barcoders.DoQuery(barcodestest)

            Dim defaultBARCODE As String = ""
            Dim generatedBarcode As String = ""

            If barcoders.RecordCount = 0 Then
                defaultBARCODE = "No"
                ' generate default barcode using your existing function if present
                Try
                    ' If New_Barcode_Generation exists:
                    generatedBarcode = Me.New_Barcode_Generation(Now.Ticks.ToString(), oItem.ItemName)
                Catch ex As Exception
                    generatedBarcode = ""
                End Try
                If generatedBarcode <> "" Then
                    oItem.BarCodes.BarCode = generatedBarcode
                    oItem.BarCodes.FreeText = oItem.ItemName
                    ' pick first UoM if available on item
                    Try
                        If oItem.DefaultSalesUoMEntry <> 0 Then
                            oItem.BarCodes.UoMEntry = oItem.DefaultSalesUoMEntry
                        End If
                    Catch ex As Exception
                    End Try
                    oItem.BarCodes.SetCurrentLine(0)
                    oItem.BarCodes.Add()
                End If
            Else
                ' iterate child barcode rows
                barcoders.MoveFirst()
                Dim p As Integer = 0
                While barcoders.EoF = False
                    Dim barcodest As String = If(barcoders.Fields.Item("U_BarCode").Value IsNot Nothing, barcoders.Fields.Item("U_BarCode").Value.ToString(), "")
                    Dim barcodename As String = If(barcoders.Fields.Item("U_BDesc").Value IsNot Nothing, barcoders.Fields.Item("U_BDesc").Value.ToString(), "")
                    Dim barcodeuom As Integer = If(barcoders.Fields.Item("UomEntry").Value IsNot Nothing, Convert.ToInt32(barcoders.Fields.Item("UomEntry").Value), 0)
                    Dim childDocEntry As String = barcoders.Fields.Item("DocEntry").Value.ToString()
                    Dim childLineId As String = barcoders.Fields.Item("LineId").Value.ToString()

                    If String.IsNullOrEmpty(barcodest) AndAlso barcodename <> "" AndAlso barcodeuom <> 0 Then
                        ' generate barcode using your helper if available
                        Try
                            generatedBarcode = Me.New_Barcode_GenerationWithBarcodeDescription(Now.Ticks.ToString(), barcodename)
                        Catch ex As Exception
                            generatedBarcode = ""
                        End Try
                        If generatedBarcode <> "" Then
                            oItem.BarCodes.BarCode = generatedBarcode
                            oItem.BarCodes.FreeText = barcodename
                            oItem.BarCodes.UoMEntry = barcodeuom
                            oItem.BarCodes.SetCurrentLine(p)
                            oItem.BarCodes.Add()
                            ' update child with generated barcode
                            Dim StrQry As String = "UPDATE ""@TNX_ITEMUP_C2"" SET ""U_BarCode"" = '" & generatedBarcode.Replace("'", "''") & "' Where ""DocEntry""='" & childDocEntry.Replace("'", "''") & "' and ""LineId""='" & childLineId.Replace("'", "''") & "'"
                            Dim oRec As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRec.DoQuery(StrQry)
                        End If
                    Else
                        ' Use existing barcode value
                        oItem.BarCodes.BarCode = barcodest
                        oItem.BarCodes.FreeText = barcodename
                        If barcodeuom <> 0 Then
                            oItem.BarCodes.UoMEntry = barcodeuom
                        End If
                        oItem.BarCodes.SetCurrentLine(p)
                        oItem.BarCodes.Add()
                    End If

                    p += 1
                    barcoders.MoveNext()
                End While
            End If

        Catch ex As Exception
            Throw New Exception("AssignBarCodes error: " & ex.Message)
        End Try
    End Sub
    Private Sub UpdateSugarTax_UOMWise1(ByVal DocEntry As String)

        Dim UomCount As Integer = 0

        Try
            Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim Query As String =
        "SELECT * FROM ""@TNX_ITEMUP"" T0 " &
        "INNER JOIN ""@TNX_ITEMUP_C0"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
        "WHERE T0.""DocNum"" = '" & DocEntry & "' AND T1.""U_ITMCD"" <> ''"

            rs.DoQuery(Query)

            If rs.RecordCount = 0 Then Exit Sub

            For i As Integer = 1 To rs.RecordCount

                Dim ItemCode As String = rs.Fields.Item("U_ITMCD").Value
                Dim KeyCode As String = rs.Fields.Item("U_KEYC").Value

                '===============================
                ' GET SUGAR TAX PRICE (C2 TABLE)
                '===============================
                Dim GetPriceQuery As String =
            "SELECT DISTINCT " &
            "(SELECT ""UomEntry"" FROM OUOM WHERE ""UomCode"" = T2.""U_UOMCDE"")AS ""UomEntry"", " &
            "T3.""IUoMEntry"" AS ""DefaultUOM"", " &
            "T2.""U_Price"" AS ""NewPrice"" " &
            "FROM ""@TNX_ITEMUP"" T0 " &
            "INNER JOIN ""@TNX_ITEMUP_C2"" T2 ON T0.""DocEntry"" = T2.""DocEntry"" " &
            "INNER JOIN ""@TNX_ITEMUP_C0"" T1 ON T1.""DocEntry"" = T0.""DocEntry"" AND T1.""U_KEYC"" = T2.""U_KEYC"" " &
            "INNER JOIN OITM T3 ON T3.""ItemCode"" = T1.""U_ITMCD"" " &
            "WHERE T0.""DocNum"" = '" & DocEntry & "' " &
            "AND T2.""U_KEYC"" = '" & KeyCode & "' " &
            "AND T2.""U_Price"" > 0 " &
            "AND T2.""U_STAX"" = 'Yes' " &
            "AND T1.""U_SugarTax"" = 'Yes'"

                Dim rsPrice As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rsPrice.DoQuery(GetPriceQuery)

                If rsPrice.RecordCount > 0 Then

                    rsPrice.MoveFirst()

                    For j As Integer = 1 To rsPrice.RecordCount

                        Dim UomEntry As Integer = CInt(rsPrice.Fields.Item("UomEntry").Value)
                        Dim DefaultUom As Integer = CInt(rsPrice.Fields.Item("DefaultUOM").Value)
                        ' Dim BranchId As Integer = CInt(rsPrice.Fields.Item("BranchId").Value)
                        Dim NewPrice As Double = CDbl(rsPrice.Fields.Item("NewPrice").Value)

                        Dim oItem As SAPbobsCOM.Items = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

                        If oItem.GetByKey(ItemCode) Then

                            '===============================
                            ' GET BRANCH PRICE LISTS
                            '===============================
                            Dim oRsBranch As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsBranch.DoQuery("SELECT ""U_PricelistNo"" FROM ""OBPL"" WHERE ""U_PricelistNo"" <> ''")

                            For r As Integer = 1 To oRsBranch.RecordCount

                                Dim PriceListId As Integer = CInt(oRsBranch.Fields.Item(0).Value)

                                oItem.PriceList.SetCurrentLine(PriceListId - 1)
                                Dim CurrentPriceList As String = oItem.PriceList.PriceList

                                If CurrentPriceList = PriceListId Then

                                    '===============================
                                    ' DEFAULT UOM
                                    '===============================
                                    If DefaultUom = UomEntry Then

                                        oItem.PriceList.Price = NewPrice

                                    Else
                                        '===============================
                                        ' CHECK ITM9 (IMPORTANT)
                                        '===============================
                                        '    Dim chkQuery As String =
                                        '"SELECT COUNT(*) FROM ITM9 " &
                                        '"WHERE ""ItemCode"" = '" & ItemCode & "' " &
                                        '"AND ""UomEntry"" = '" & UomEntry & "' " &
                                        '"AND ""PriceList"" = '" & CurrentPriceList & "'"

                                        '    Dim rsChk As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        '    rsChk.DoQuery(chkQuery)

                                        Dim Found As Boolean = False

                                        '    If rsChk.Fields.Item(0).Value > 0 Then

                                        ' UPDATE EXISTING
                                        For x As Integer = 0 To oItem.PriceList.UoMPrices.Count - 1
                                            oItem.PriceList.UoMPrices.SetCurrentLine(x)

                                            If oItem.PriceList.UoMPrices.UoMEntry = UomEntry Then
                                                oItem.PriceList.UoMPrices.Price = NewPrice
                                                Found = True
                                                Exit For
                                            End If
                                        Next

                                        '      End If

                                        ' ADD NEW
                                        If Not Found Then
                                            'oItem.PriceList.UoMPrices.Add()
                                            'oItem.PriceList.UoMPrices.SetCurrentLine(oItem.PriceList.UoMPrices.Count - 1)
                                            'oItem.PriceList.UoMPrices.UoMEntry = UomEntry
                                            'oItem.PriceList.UoMPrices.Price = NewPrice
                                            'oItem.PriceList.UoMPrices.PriceList = CurrentPriceList
                                            'oItem.PriceList.UoMPrices.Auto = SAPbobsCOM.BoYesNoEnum.tNO
                                            Dim GetUomPriceListCount As String = "Select Count(*) From ITM9 Where ""ItemCode"" = '" & ItemCode & "' And ""PriceList"" = '" & CurrentPriceList & "'"
                                            Dim oRsGetUomPriceListCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRsGetUomPriceListCount.DoQuery(GetUomPriceListCount)
                                            If oRsGetUomPriceListCount.Fields.Item(0).Value = 0 Then
                                                UomCount = 0
                                            Else
                                                UomCount = oRsGetUomPriceListCount.Fields.Item(0).Value
                                            End If
                                            oItem.PriceList.UoMPrices.Add()
                                            oItem.PriceList.UoMPrices.SetCurrentLine(UomCount)
                                            oItem.PriceList.UoMPrices.Price = NewPrice
                                            oItem.PriceList.UoMPrices.UoMEntry = UomEntry
                                            oItem.PriceList.UoMPrices.Auto = SAPbobsCOM.BoYesNoEnum.tNO
                                            ' oItemMaster.PriceList.UoMPrices.Add()
                                        End If

                                    End If
                                End If

                                oRsBranch.MoveNext()
                            Next

                            '===============================
                            ' UPDATE ITEM
                            '===============================
                            Dim RetVal As Integer = oItem.Update()

                            If RetVal <> 0 Then
                                Dim ErrMsg As String = objMain.objCompany.GetLastErrorDescription()

                                Dim qErr As String =
                            "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_PStst""='" & ErrMsg & "' " &
                            "WHERE ""LineId""='" & i & "' AND ""DocEntry""=(" &
                            "SELECT ""DocEntry"" FROM ""@TNX_ITEMUP"" WHERE ""DocNum""='" & DocEntry & "')"

                                Dim rsErr As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                rsErr.DoQuery(qErr)

                            Else
                                Dim qSuc As String =
                            "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_PStst""='Sugar Tax /Supplier Price is updated successfully' " &
                            "WHERE ""LineId""='" & i & "' AND ""DocEntry""=(" &
                            "SELECT ""DocEntry"" FROM ""@TNX_ITEMUP"" WHERE ""DocNum""='" & DocEntry & "')"

                                Dim rsSuc As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                rsSuc.DoQuery(qSuc)
                            End If

                        End If

                        rsPrice.MoveNext()
                    Next
                End If

                rs.MoveNext()
            Next

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Private Sub UpdateSugarTax_UOMWise(ByVal DocEntry As String)

        Dim UomCount As Integer = 0

        Try
            Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim query As String =
        "SELECT * FROM ""@TNX_ITEMUP"" T0 " &
        "INNER JOIN ""@TNX_ITEMUP_C0"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
        "WHERE T0.""DocNum"" = '" & DocEntry & "' AND T1.""U_ITMCD"" <> '' " &
      "AND ( IFNULL(T1.""U_MPUS"",'') <> 'Sugar Tax /Supplier Price is updated successfully (PriceList 11)')"

            rs.DoQuery(query)

            If rs.RecordCount = 0 Then
                objMain.objApplication.StatusBar.SetText("No records found for processing.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Exit Sub
            End If

            While Not rs.EoF

                Dim itemCode As String = rs.Fields.Item("U_ITMCD").Value
                Dim keyCode As String = rs.Fields.Item("U_KEYC").Value

                ' 🔄 Progress Message
                objMain.objApplication.StatusBar.SetText("Processing Item: " & itemCode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                Dim priceQuery As String =
            "SELECT DISTINCT " &
            "(SELECT ""UomEntry"" FROM OUOM WHERE ""UomCode"" = T2.""U_UOMCDE"") AS ""UomEntry"", " &
            "T3.""IUoMEntry"" AS ""DefaultUOM"", " &
            "T2.""U_Price"" AS ""NewPrice"" " &
            "FROM ""@TNX_ITEMUP"" T0 " &
            "INNER JOIN ""@TNX_ITEMUP_C2"" T2 ON T0.""DocEntry"" = T2.""DocEntry"" " &
            "INNER JOIN ""@TNX_ITEMUP_C0"" T1 ON T1.""DocEntry"" = T0.""DocEntry"" AND T1.""U_KEYC"" = T2.""U_KEYC"" " &
            "INNER JOIN OITM T3 ON T3.""ItemCode"" = T1.""U_ITMCD"" " &
            "WHERE T0.""DocNum"" = '" & DocEntry & "' " &
            "AND T2.""U_KEYC"" = '" & keyCode & "' " &
            "AND T2.""U_Price"" > 0 " &
            "AND T2.""U_STAX"" = 'Yes' " &
            "AND T1.""U_SugarTax"" = 'Yes'
        "

                priceQuery = "SELECT DISTINCT (SELECT ""UomEntry"" FROM OUOM WHERE ""UomCode"" = T2.""U_UOMCDE"") AS ""UomEntry"", " &
"T3.""IUoMEntry"" AS ""DefaultUOM"", T2.""U_Price"" AS ""NewPrice"" FROM ""@TNX_ITEMUP"" T0 " &
"INNER JOIN ""@TNX_ITEMUP_C2"" T2 ON T0.""DocEntry"" = T2.""DocEntry"" " &
"INNER JOIN ""@TNX_ITEMUP_C0"" T1 ON T1.""DocEntry"" = T0.""DocEntry"" AND T1.""U_KEYC"" = T2.""U_KEYC"" " &
"INNER JOIN OITM T3 ON T3.""ItemCode"" = T1.""U_ITMCD"" WHERE T0.""DocNum"" = '" & DocEntry & "' " &
"AND T2.""U_KEYC"" = '" & keyCode & "' AND T3.""ItemCode"" = '" & itemCode & "' AND T2.""U_Price"" > 0 AND T2.""U_STAX"" = 'Yes' AND T1.""U_SugarTax"" = 'Yes' " &
"UNION ALL " &
"SELECT DISTINCT T1.""U_UomEtry"" AS ""UomEntry"",T3.""IUoMEntry"" AS ""DefaultUOM"",T1.""U_OldPrice"" AS ""NewPrice"" FROM ""@TNX_ITEMUP"" T0 " &
"INNER JOIN ""@TNX_ITEMUP_C3"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" INNER JOIN ""@TNX_ITEMUP_C0"" T2 ON T0.""DocEntry"" = T2.""DocEntry"" " &
"INNER JOIN OITM T3 ON T3.""ItemCode"" = T2.""U_ITMCD""  WHERE T0.""DocNum"" = '" & DocEntry & "' AND T1.""U_ItemC"" = '" & keyCode & "' " &
"AND T1.""U_OldPrice"" > 0 AND T1.""U_UomEtry"" <> '' AND T1.""U_PriceId"" = '11' AND T1.""U_ItemC"" IN (" &
"SELECT DISTINCT T2.""U_KEYC"" " &
"FROM ""@TNX_ITEMUP"" T0 INNER JOIN ""@TNX_ITEMUP_C2"" T2 ON T0.""DocEntry"" = T2.""DocEntry"" " &
"INNER JOIN ""@TNX_ITEMUP_C0"" T1 ON T1.""DocEntry"" = T0.""DocEntry"" AND T1.""U_KEYC"" = T2.""U_KEYC"" " &
"WHERE T0.""DocNum"" = '" & DocEntry & "' AND T3.""ItemCode"" = '" & itemCode & "' AND T2.""U_STAX"" = 'No' AND T1.""U_SugarTax"" = 'No')" &
"UNION ALL " &
"SELECT DISTINCT T1.""U_UomEtry"" AS ""UomEntry"",T3.""IUoMEntry"" AS ""DefaultUOM"",T1.""U_Cost"" AS ""NewPrice"" FROM ""@TNX_ITEMUP"" T0 " &
"INNER JOIN ""@TNX_ITEMUP_C3"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" INNER JOIN ""@TNX_ITEMUP_C0"" T2 ON T0.""DocEntry"" = T2.""DocEntry"" " &
"INNER JOIN OITM T3 ON T3.""ItemCode"" = T2.""U_ITMCD""  WHERE T0.""DocNum"" = '" & DocEntry & "' AND T1.""U_ItemC"" = '" & keyCode & "' " &
"AND T1.""U_Cost"" > 0 AND T1.""U_UomEtry"" <> '' AND T1.""U_PriceList"" = '11' AND T1.""U_ItemC"" IN (" &
"SELECT DISTINCT T2.""U_KEYC"" " &
"FROM ""@TNX_ITEMUP"" T0 INNER JOIN ""@TNX_ITEMUP_C2"" T2 ON T0.""DocEntry"" = T2.""DocEntry"" " &
"INNER JOIN ""@TNX_ITEMUP_C0"" T1 ON T1.""DocEntry"" = T0.""DocEntry"" AND T1.""U_KEYC"" = T2.""U_KEYC"" " &
"WHERE T0.""DocNum"" = '" & DocEntry & "' AND T3.""ItemCode"" = '" & itemCode & "' AND T2.""U_STAX"" = 'No' AND T1.""U_SugarTax"" = 'No')
"

                Dim rsPrice As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rsPrice.DoQuery(priceQuery)

                If rsPrice.RecordCount > 0 Then

                    Dim oItem As SAPbobsCOM.Items = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

                    If oItem.GetByKey(itemCode) Then

                        While Not rsPrice.EoF

                            Dim uomEntry As Integer = CInt(rsPrice.Fields.Item("UomEntry").Value)
                            Dim defaultUom As Integer = CInt(rsPrice.Fields.Item("DefaultUOM").Value)
                            Dim newPrice As Double = CDbl(rsPrice.Fields.Item("NewPrice").Value)

                            ' ✅ ONLY PRICE LIST 11
                            For i As Integer = 0 To oItem.PriceList.Count - 1
                                oItem.PriceList.SetCurrentLine(i)

                                If oItem.PriceList.PriceList = 11 Then

                                    If defaultUom = uomEntry Then
                                        oItem.PriceList.Price = newPrice
                                    Else
                                        Dim found As Boolean = False

                                        For x As Integer = 0 To oItem.PriceList.UoMPrices.Count - 1
                                            oItem.PriceList.UoMPrices.SetCurrentLine(x)

                                            If oItem.PriceList.UoMPrices.UoMEntry = uomEntry Then
                                                oItem.PriceList.UoMPrices.Price = newPrice
                                                found = True
                                                Exit For
                                            End If
                                        Next

                                        If Not found Then
                                            Dim countQuery As String =
                                        "SELECT COUNT(*) FROM ITM9 WHERE ""ItemCode""='" & itemCode & "' AND ""PriceList""='11'"

                                            Dim rsCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            rsCount.DoQuery(countQuery)

                                            UomCount = If(rsCount.Fields.Item(0).Value = 0, 0, rsCount.Fields.Item(0).Value)

                                            oItem.PriceList.UoMPrices.Add()
                                            oItem.PriceList.UoMPrices.SetCurrentLine(UomCount)
                                            oItem.PriceList.UoMPrices.UoMEntry = uomEntry
                                            oItem.PriceList.UoMPrices.Price = newPrice
                                            oItem.PriceList.UoMPrices.Auto = SAPbobsCOM.BoYesNoEnum.tNO
                                        End If
                                    End If

                                    Exit For
                                End If
                            Next

                            rsPrice.MoveNext()
                        End While

                        ' 🔄 Updating Item
                        objMain.objApplication.StatusBar.SetText("Updating Item: " & itemCode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        Dim retVal As Integer = oItem.Update()
                        Dim statusMsg As String

                        Dim xml As String = oItem.GetAsXML()
                        Dim xmlDoc As New XmlDocument()
                        xmlDoc.LoadXml(xml)
                        Dim jsonBody As String = JsonConvert.SerializeXmlNode(xmlDoc, Newtonsoft.Json.Formatting.Indented, True)

                        If retVal <> 0 Then
                            statusMsg = objMain.objCompany.GetLastErrorDescription()

                            objMain.objApplication.StatusBar.SetText("Error updating Item: " & itemCode & " - " & statusMsg, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            statusMsg = "SSugar Tax /Supplier Price is updated successfully (PriceList 11)"

                            objMain.objApplication.StatusBar.SetText("Success: " & itemCode & " updated", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        End If

                        Dim updateStatusQuery As String =
                    "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_MPUS""='" & statusMsg.Replace("'", "''") & "' " &
                    "WHERE ""LineId""='" & rs.Fields.Item("LineId").Value & "' AND ""DocEntry""=(" &
                    "SELECT ""DocEntry"" FROM ""@TNX_ITEMUP"" WHERE ""DocNum""='" & DocEntry & "')"

                        Dim rsStatus As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        rsStatus.DoQuery(updateStatusQuery)

                    End If
                End If

                rs.MoveNext()
            End While

            ' ✅ Final Completion Message
            objMain.objApplication.StatusBar.SetText("Sugar Tax update completed for all items.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    'Public Async Function UpdateSugarTax_UOMWise(ByVal DocEntry As String) As Task

    '    Dim rs As SAPbobsCOM.Recordset = Nothing
    '    Dim rsPrice As SAPbobsCOM.Recordset = Nothing
    '    Dim rsStatus As SAPbobsCOM.Recordset = Nothing

    '    Try

    '        ' ==========================================
    '        ' LOGIN TO SERVICE LAYER
    '        ' ==========================================
    '        Dim sessionId As String = Await CompanySession()

    '        If String.IsNullOrEmpty(sessionId) Then

    '            objMain.objApplication.StatusBar.SetText(
    '        "Service Layer Login Failed",
    '        SAPbouiCOM.BoMessageTime.bmt_Short,
    '        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '            Exit Function

    '        End If

    '        ' ==========================================
    '        ' GET ITEMS
    '        ' ==========================================
    '        rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '        Dim query As String =
    '    "SELECT * FROM ""@TNX_ITEMUP"" T0 " &
    '    "INNER JOIN ""@TNX_ITEMUP_C0"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
    '    "WHERE T0.""DocNum"" = '" & DocEntry & "' " &
    '    "AND T1.""U_ITMCD"" <> '' " &
    '    "AND IFNULL(T1.""U_MPUS"", '') <> 'Sugar Tax /Supplier Price is updated successfully (PriceList 11)'"

    '        rs.DoQuery(query)

    '        If rs.RecordCount = 0 Then

    '            objMain.objApplication.StatusBar.SetText(
    '        "No records found for processing.",
    '        SAPbouiCOM.BoMessageTime.bmt_Short,
    '        SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

    '            Exit Function

    '        End If

    '        While Not rs.EoF

    '            Dim itemCode As String = rs.Fields.Item("U_ITMCD").Value.ToString.Trim()
    '            Dim keyCode As String = rs.Fields.Item("U_KEYC").Value.ToString.Trim()

    '            objMain.objApplication.StatusBar.SetText(
    '        "Processing Item : " & itemCode,
    '        SAPbouiCOM.BoMessageTime.bmt_Short,
    '        SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

    '            ' ==========================================
    '            ' GET PRICE DETAILS
    '            ' ==========================================
    '            rsPrice = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '            Dim priceQuery As String =
    '        "SELECT DISTINCT " &
    '        "(SELECT ""UomEntry"" FROM OUOM WHERE ""UomCode"" = T2.""U_UOMCDE"") AS ""UomEntry"", " &
    '        "T3.""IUoMEntry"" AS ""DefaultUOM"", " &
    '        "T2.""U_Price"" AS ""NewPrice"" " &
    '        "FROM ""@TNX_ITEMUP"" T0 " &
    '        "INNER JOIN ""@TNX_ITEMUP_C2"" T2 ON T0.""DocEntry"" = T2.""DocEntry"" " &
    '        "INNER JOIN ""@TNX_ITEMUP_C0"" T1 ON T1.""DocEntry"" = T0.""DocEntry"" AND T1.""U_KEYC"" = T2.""U_KEYC"" " &
    '        "INNER JOIN OITM T3 ON T3.""ItemCode"" = T1.""U_ITMCD"" " &
    '        "WHERE T0.""DocNum"" = '" & DocEntry & "' " &
    '        "AND T2.""U_KEYC"" = '" & keyCode & "' " &
    '        "AND T2.""U_Price"" > 0 " &
    '        "AND T2.""U_STAX"" = 'Yes' " &
    '        "AND T1.""U_SugarTax"" = 'Yes' " &
    '        "UNION ALL " &
    '        "SELECT DISTINCT " &
    '        "T1.""U_UomEtry"" AS ""UomEntry"", " &
    '        "T3.""IUoMEntry"" AS ""DefaultUOM"", " &
    '        "T1.""U_OldPrice"" AS ""NewPrice"" " &
    '        "FROM ""@TNX_ITEMUP"" T0 " &
    '        "INNER JOIN ""@TNX_ITEMUP_C3"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
    '        "INNER JOIN ""@TNX_ITEMUP_C0"" T2 ON T0.""DocEntry"" = T2.""DocEntry"" " &
    '        "INNER JOIN OITM T3 ON T3.""ItemCode"" = T2.""U_ITMCD"" " &
    '        "WHERE T0.""DocNum"" = '" & DocEntry & "' " &
    '        "AND T1.""U_ItemC"" = '" & keyCode & "' " &
    '        "AND T1.""U_OldPrice"" > 0 " &
    '        "AND T1.""U_UomEtry"" <> '' " &
    '        "AND T1.""U_PriceId"" = '11' " &
    '        "AND T1.""U_ItemC"" IN ( " &
    '        "SELECT DISTINCT T2.""U_KEYC"" " &
    '        "FROM ""@TNX_ITEMUP"" T0 " &
    '        "INNER JOIN ""@TNX_ITEMUP_C2"" T2 ON T0.""DocEntry"" = T2.""DocEntry"" " &
    '        "INNER JOIN ""@TNX_ITEMUP_C0"" T1 ON T1.""DocEntry"" = T0.""DocEntry"" AND T1.""U_KEYC"" = T2.""U_KEYC"" " &
    '        "WHERE T0.""DocNum"" = '" & DocEntry & "' " &
    '        "AND T2.""U_STAX"" = 'No' " &
    '        "AND T1.""U_SugarTax"" = 'No')"

    '            rsPrice.DoQuery(priceQuery)

    '            If rsPrice.RecordCount > 0 Then

    '                ' ==========================================
    '                ' BUILD SERVICE LAYER PAYLOAD
    '                ' ==========================================
    '                Dim itemPrices As New List(Of Object)
    '                Dim uomPrices As New List(Of Object)

    '                Dim basePrice As Double = 0

    '                While Not rsPrice.EoF

    '                    Dim uomEntry As Integer = CInt(rsPrice.Fields.Item("UomEntry").Value)
    '                    Dim defaultUom As Integer = CInt(rsPrice.Fields.Item("DefaultUOM").Value)
    '                    Dim newPrice As Double = CDbl(rsPrice.Fields.Item("NewPrice").Value)

    '                    ' DEFAULT UOM PRICE
    '                    If defaultUom = uomEntry Then

    '                        basePrice = newPrice

    '                    Else

    '                        uomPrices.Add(New Dictionary(Of String, Object) From {
    '                        {"UoMEntry", uomEntry},
    '                        {"Price", newPrice},
    '                        {"Auto", "tNO"}
    '                    })

    '                    End If

    '                    rsPrice.MoveNext()

    '                End While

    '                ' PRICE LIST 11
    '                itemPrices.Add(New Dictionary(Of String, Object) From {
    '                {"PriceList", 11},
    '                {"Price", basePrice},
    '                {"UoMPrices", uomPrices}
    '            })

    '                Dim payload As New Dictionary(Of String, Object) From {
    '                {"ItemPrices", itemPrices}
    '            }

    '                Dim jsonBody As String =
    '            JsonConvert.SerializeObject(payload)

    '                Dim url As String =
    '            baseUrl & "Items('" & itemCode & "')"

    '                objMain.objApplication.StatusBar.SetText(
    '            "Updating Item : " & itemCode,
    '            SAPbouiCOM.BoMessageTime.bmt_Short,
    '            SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

    '                ' ==========================================
    '                ' PATCH ITEM
    '                ' ==========================================
    '                Dim response = Await PatchItem(sessionId, url, jsonBody)

    '                Dim statusMsg As String = ""

    '                If response.StatusCode = HttpStatusCode.NoContent OrElse
    '               response.StatusCode = HttpStatusCode.OK Then

    '                    statusMsg = "Sugar Tax /Supplier Price is updated successfully (PriceList 11)"

    '                    objMain.objApplication.StatusBar.SetText(
    '                "Success : " & itemCode & " updated",
    '                SAPbouiCOM.BoMessageTime.bmt_Short,
    '                SAPbouiCOM.BoStatusBarMessageType.smt_Success)

    '                Else

    '                    statusMsg = response.ResponseBody.Replace("'", " ")

    '                    objMain.objApplication.StatusBar.SetText(
    '                "Error updating Item : " & itemCode & " - " & statusMsg,
    '                SAPbouiCOM.BoMessageTime.bmt_Long,
    '                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '                End If

    '                ' ==========================================
    '                ' UPDATE STATUS
    '                ' ==========================================
    '                rsStatus = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '                Dim updateStatusQuery As String =
    '            "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_MPUS""='" & statusMsg.Replace("'", "''") & "' " &
    '            "WHERE ""LineId""='" & rs.Fields.Item("LineId").Value & "' " &
    '            "AND ""DocEntry""=(" &
    '            "SELECT ""DocEntry"" FROM ""@TNX_ITEMUP"" WHERE ""DocNum""='" & DocEntry & "')"

    '                rsStatus.DoQuery(updateStatusQuery)

    '            End If

    '            rs.MoveNext()

    '        End While

    '        objMain.objApplication.StatusBar.SetText(
    '    "Sugar Tax update completed for all items.",
    '    SAPbouiCOM.BoMessageTime.bmt_Short,
    '    SAPbouiCOM.BoStatusBarMessageType.smt_Success)

    '    Catch ex As Exception

    '        objMain.objApplication.StatusBar.SetText(
    '    "Error : " & ex.Message,
    '    SAPbouiCOM.BoMessageTime.bmt_Long,
    '    SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '    Finally

    '        If rs IsNot Nothing Then
    '            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
    '        End If

    '        If rsPrice IsNot Nothing Then
    '            System.Runtime.InteropServices.Marshal.ReleaseComObject(rsPrice)
    '        End If

    '        If rsStatus IsNot Nothing Then
    '            System.Runtime.InteropServices.Marshal.ReleaseComObject(rsStatus)
    '        End If

    '    End Try

    'End Function

    'Public Async Function UpdateSugarTax_UOMWise(ByVal DocEntry As String) As Task

    '    Dim rs As SAPbobsCOM.Recordset = Nothing
    '    Dim rsPrice As SAPbobsCOM.Recordset = Nothing
    '    Dim rsUpdate As SAPbobsCOM.Recordset = Nothing

    '    Try

    '        ' ===========================
    '        ' 1. LOGIN TO SERVICE LAYER
    '        ' ===========================
    '        Dim sessionId As String = Await CompanySession()

    '        If String.IsNullOrEmpty(sessionId) Then
    '            objMain.objApplication.SetStatusBarMessage(
    '        "SL Login Failed",
    '        SAPbouiCOM.BoMessageTime.bmt_Short,
    '        True)

    '            Exit Function
    '        End If

    '        ' ===========================
    '        ' 2. GET ITEMS TO PROCESS
    '        ' ===========================
    '        rs = DirectCast(
    '    objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset),
    '    SAPbobsCOM.Recordset)

    '        Dim query As String =
    '    "SELECT T0.""DocEntry"", T1.""LineId"", T1.""U_ITMCD"", T1.""U_KEYC"" " &
    '    "FROM ""@TNX_ITEMUP"" T0 " &
    '    "INNER JOIN ""@TNX_ITEMUP_C0"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
    '    "WHERE T0.""DocNum"" = '" & DocEntry & "' " &
    '    "AND IFNULL(T1.""U_ITMCD"", '') <> '' " &
    '    "AND IFNULL(T1.""U_MPUS"", '') <> 'Sugar Tax /Supplier Price is updated successfully (PriceList 11)'"

    '        rs.DoQuery(query)

    '        If rs.RecordCount = 0 Then

    '            objMain.objApplication.SetStatusBarMessage(
    '        "No records found for processing.",
    '        SAPbouiCOM.BoMessageTime.bmt_Short,
    '        False)

    '            Exit Function

    '        End If

    '        rsUpdate = DirectCast(
    '    objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset),
    '    SAPbobsCOM.Recordset)

    '        Dim totalItems As Integer = rs.RecordCount
    '        Dim currentItem As Integer = 0

    '        ' ===========================
    '        ' 3. LOOP ITEMS
    '        ' ===========================
    '        While Not rs.EoF

    '            currentItem += 1

    '            Dim itemCode As String = rs.Fields.Item("U_ITMCD").Value.ToString().Trim()
    '            Dim keyCode As String = rs.Fields.Item("U_KEYC").Value.ToString().Trim()
    '            Dim lineId As String = rs.Fields.Item("LineId").Value.ToString()
    '            Dim docEntryVal As String = rs.Fields.Item("DocEntry").Value.ToString()

    '            objMain.objApplication.SetStatusBarMessage(
    '            $"Processing {currentItem}/{totalItems} : {itemCode}",
    '            SAPbouiCOM.BoMessageTime.bmt_Short,
    '            False)

    '            ' ===========================
    '            ' 4. GET PRICE DETAILS
    '            ' ===========================
    '            rsPrice = DirectCast(
    '            objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset),
    '            SAPbobsCOM.Recordset)

    '            Dim priceQuery As String =
    '        "SELECT DISTINCT " &
    '        "(SELECT ""UomEntry"" FROM OUOM WHERE ""UomCode"" = T2.""U_UOMCDE"") AS ""UomEntry"", " &
    '        "T3.""IUoMEntry"" AS ""DefaultUOM"", " &
    '        "T2.""U_Price"" AS ""NewPrice"" " &
    '        "FROM ""@TNX_ITEMUP"" T0 " &
    '        "INNER JOIN ""@TNX_ITEMUP_C2"" T2 ON T0.""DocEntry"" = T2.""DocEntry"" " &
    '        "INNER JOIN ""@TNX_ITEMUP_C0"" T1 ON T1.""DocEntry"" = T0.""DocEntry"" AND T1.""U_KEYC"" = T2.""U_KEYC"" " &
    '        "INNER JOIN OITM T3 ON T3.""ItemCode"" = T1.""U_ITMCD"" " &
    '        "WHERE T0.""DocNum"" = '" & DocEntry & "' " &
    '        "AND T2.""U_KEYC"" = '" & keyCode & "' " &
    '        "AND T2.""U_Price"" > 0 " &
    '        "AND T2.""U_STAX"" = 'Yes' " &
    '        "AND T1.""U_SugarTax"" = 'Yes'
    '            "

    '            priceQuery = "SELECT DISTINCT (SELECT ""UomEntry"" FROM OUOM WHERE ""UomCode"" = T2.""U_UOMCDE"") AS ""UomEntry"", " &
    '            "T3.""IUoMEntry"" AS ""DefaultUOM"", T2.""U_Price"" AS ""NewPrice"" FROM ""@TNX_ITEMUP"" T0 " &
    '            "INNER JOIN ""@TNX_ITEMUP_C2"" T2 ON T0.""DocEntry"" = T2.""DocEntry"" " &
    '            "INNER JOIN ""@TNX_ITEMUP_C0"" T1 ON T1.""DocEntry"" = T0.""DocEntry"" AND T1.""U_KEYC"" = T2.""U_KEYC"" " &
    '            "INNER JOIN OITM T3 ON T3.""ItemCode"" = T1.""U_ITMCD"" WHERE T0.""DocNum"" = '" & DocEntry & "' " &
    '            "AND T2.""U_KEYC"" = '" & keyCode & "' AND T2.""U_Price"" > 0 AND T2.""U_STAX"" = 'Yes' AND T1.""U_SugarTax"" = 'Yes' " &
    '            "UNION ALL " &
    '            "SELECT DISTINCT T1.""U_UomEtry"" AS ""UomEntry"",T3.""IUoMEntry"" AS ""DefaultUOM"",T1.""U_OldPrice"" AS ""NewPrice"" FROM ""@TNX_ITEMUP"" T0 " &
    '            "INNER JOIN ""@TNX_ITEMUP_C3"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" INNER JOIN ""@TNX_ITEMUP_C0"" T2 ON T0.""DocEntry"" = T2.""DocEntry"" " &
    '            "INNER JOIN OITM T3 ON T3.""ItemCode"" = T2.""U_ITMCD""  WHERE T0.""DocNum"" = '" & DocEntry & "' AND T1.""U_ItemC"" = '" & keyCode & "' " &
    '            "AND T1.""U_OldPrice"" > 0 AND T1.""U_UomEtry"" <> '' AND T1.""U_PriceId"" = '11' AND T1.""U_ItemC"" IN (" &
    '            "SELECT DISTINCT T2.""U_KEYC"" " &
    '            "FROM ""@TNX_ITEMUP"" T0 INNER JOIN ""@TNX_ITEMUP_C2"" T2 ON T0.""DocEntry"" = T2.""DocEntry"" " &
    '            "INNER JOIN ""@TNX_ITEMUP_C0"" T1 ON T1.""DocEntry"" = T0.""DocEntry"" AND T1.""U_KEYC"" = T2.""U_KEYC"" " &
    '            "WHERE T0.""DocNum"" = '" & DocEntry & "' AND T2.""U_STAX"" = 'No' AND T1.""U_SugarTax"" = 'No')"

    '            rsPrice.DoQuery(priceQuery)

    '            Dim dt As DataTable = ConvertRecordsetToDataTable(rsPrice)

    '            If dt Is Nothing OrElse dt.Rows.Count = 0 Then

    '                rs.MoveNext()
    '                Continue While

    '            End If

    '            ' ===========================
    '            ' 5. BUILD ITEM PRICE PAYLOAD
    '            ' ===========================
    '            Dim itemPricesList As New List(Of Object)

    '            Dim priceListGroups = dt.AsEnumerable().
    '            GroupBy(Function(r) Convert.ToInt32(r("PriceList")))

    '            For Each plGroup In priceListGroups

    '                Dim currentPriceList As Integer = plGroup.Key

    '                ' DEFAULT UOM PRICE
    '                Dim baseRow = plGroup.
    '                FirstOrDefault(Function(r)
    '                                   Return Convert.ToInt32(r("UomEntry")) =
    '                                          Convert.ToInt32(r("IUOM"))
    '                               End Function)

    '                ' UOM PRICES
    '                Dim uomPrices = plGroup.
    '                Where(Function(r)
    '                          Return Convert.ToInt32(r("UomEntry")) <>
    '                                 Convert.ToInt32(r("IUOM"))
    '                      End Function).
    '                GroupBy(Function(r) Convert.ToInt32(r("UomEntry"))).
    '                Select(Function(g) g.First()).
    '                Select(Function(r) New With {
    '                    .UoMEntry = Convert.ToInt32(r("UomEntry")),
    '                    .Price = Convert.ToDouble(r("Price")),
    '                    .Auto = "tNO"
    '                }).ToList()

    '                ' REMOVE DUPLICATES
    '                uomPrices = uomPrices.
    '                GroupBy(Function(x) x.UoMEntry).
    '                Select(Function(g) g.First()).
    '                ToList()

    '                ' REMOVE INVALID UOM
    '                uomPrices = uomPrices.
    '                Where(Function(x) x.UoMEntry > 0).
    '                ToList()

    '                Dim priceObj As New Dictionary(Of String, Object)

    '                priceObj.Add("PriceList", currentPriceList)

    '                ' BASE PRICE
    '                If baseRow IsNot Nothing Then
    '                    priceObj("Price") = Convert.ToDouble(baseRow("Price"))
    '                End If

    '                ' UOM PRICES
    '                If uomPrices IsNot Nothing AndAlso uomPrices.Count > 0 Then
    '                    priceObj("UoMPrices") = uomPrices
    '                End If

    '                itemPricesList.Add(priceObj)

    '            Next

    '            ' ===========================
    '            ' 6. PATCH TO SERVICE LAYER
    '            ' ===========================
    '            Dim payload As New Dictionary(Of String, Object) From {
    '                {"ItemPrices", itemPricesList}
    '            }

    '            Dim jsonBody As String =
    '            JsonConvert.SerializeObject(payload)

    '            Dim url As String =
    '            $"{baseUrl}Items('{itemCode}')"

    '            Dim response =
    '            Await PatchItem(sessionId, url, jsonBody)

    '            ' ===========================
    '            ' 7. UPDATE STATUS
    '            ' ===========================
    '            Dim statusMsg As String = ""

    '            If response.StatusCode = HttpStatusCode.OK OrElse
    '               response.StatusCode = HttpStatusCode.NoContent Then

    '                statusMsg = "Sugar Tax /Supplier Price is updated successfully (PriceList 11)"

    '                objMain.objApplication.SetStatusBarMessage(
    '                $"Success : {itemCode} updated",
    '                SAPbouiCOM.BoMessageTime.bmt_Short,
    '                False)

    '            Else

    '                statusMsg = response.ResponseBody.Replace("'", " ").Trim()

    '                objMain.objApplication.SetStatusBarMessage(
    '                $"Error updating Item : {itemCode} - {statusMsg}",
    '                SAPbouiCOM.BoMessageTime.bmt_Medium,
    '                True)

    '            End If

    '            If statusMsg.Length > 250 Then
    '                statusMsg = statusMsg.Substring(0, 250)
    '            End If

    '            rsUpdate.DoQuery(
    '            $"UPDATE ""@TNX_ITEMUP_C0"" 
    '          SET ""U_MPUS"" = '{statusMsg}' 
    '          WHERE ""DocEntry"" = '{docEntryVal}' 
    '          AND ""LineId"" = '{lineId}'")

    '            rs.MoveNext()

    '            ' RELEASE OBJECT
    '            If rsPrice IsNot Nothing Then
    '                System.Runtime.InteropServices.Marshal.ReleaseComObject(rsPrice)
    '                rsPrice = Nothing
    '            End If

    '        End While

    '        objMain.objApplication.SetStatusBarMessage(
    '    "Sugar Tax update completed for all items.",
    '    SAPbouiCOM.BoMessageTime.bmt_Short,
    '    False)

    '        objMain.objApplication.ActivateMenuItem("1304")

    '    Catch ex As Exception

    '        objMain.objApplication.SetStatusBarMessage(
    '    $"Error : {ex.Message}",
    '    SAPbouiCOM.BoMessageTime.bmt_Medium,
    '    True)

    '    Finally

    '        If rs IsNot Nothing Then
    '            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
    '        End If

    '        If rsPrice IsNot Nothing Then
    '            System.Runtime.InteropServices.Marshal.ReleaseComObject(rsPrice)
    '        End If

    '        If rsUpdate IsNot Nothing Then
    '            System.Runtime.InteropServices.Marshal.ReleaseComObject(rsUpdate)
    '        End If

    '    End Try

    'End Function
    Sub UpdateSUGARTAX13(ByVal DocEntry As String)

        Try
            '===============================
            ' STEP 1: GET PRICE LIST ID
            '===============================
            Dim PriceListId As Integer = -1
            Dim rsPL As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim PriceListName As String = "AL JAZEERA PURCHASE"
            PriceListName = "11"
            PriceListName = PriceListName.Replace("'", "''")

            rsPL.DoQuery("SELECT ""ListNum"" FROM OPLN WHERE ""ListNum""='" & PriceListName & "'")

            If rsPL.RecordCount = 0 Then
                objMain.objApplication.StatusBar.SetText("Price List not found", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            PriceListId = CInt(rsPL.Fields.Item("ListNum").Value)

            '===============================
            ' STEP 2: GET ITEMS + SUGAR TAX + VAT
            '===============================
            Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            DocEntry = DocEntry.Replace("'", "''")

            ' Join C0 and C2 where Code matches, fetch Sugar Tax and VAT
            Dim Query As String = "SELECT T1.""U_ITMCD"", T1.""U_SuTaxAmt"",""U_Price"" " &
                              "FROM ""@TNX_ITEMUP"" T0 " &
                              "INNER JOIN ""@TNX_ITEMUP_C0"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
                              "INNER JOIN ""@TNX_ITEMUP_C2"" T2 ON T1.""DocEntry"" = T2.""DocEntry"" " &
                              "WHERE T0.""DocNum"" = '" & DocEntry & "' AND T1.""U_ITMCD"" IS NOT NULL AND T1.""U_ITMCD"" <> ''"

            rs.DoQuery(Query)
            If rs.RecordCount = 0 Then Exit Sub

            '===============================
            ' STEP 3: UPDATE ITEM PRICES
            '===============================
            Dim oItem As SAPbobsCOM.Items

            While Not rs.EoF

                Dim ItemCode As String = rs.Fields.Item("U_ITMCD").Value.ToString()
                Dim SugarTax As Double = If(IsDBNull(rs.Fields.Item("U_Price").Value), 0, CDbl(rs.Fields.Item("U_Price").Value))

                ' Calculate price with VAT
                Dim PriceWithVAT As Double = 0 'SugarTax * (1 + VATPercent / 100)
                PriceWithVAT = SugarTax

                ' Update SAP Item PriceList
                oItem = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                If oItem.GetByKey(ItemCode) Then
                    For p As Integer = 0 To oItem.PriceList.Count - 1
                        oItem.PriceList.SetCurrentLine(p)
                        If oItem.PriceList.PriceList = PriceListId Then
                            oItem.PriceList.Price = PriceWithVAT
                            Exit For
                        End If
                    Next

                    Dim RetVal As Integer = oItem.Update
                    If RetVal <> 0 Then
                        objMain.objApplication.StatusBar.SetText("Error updating " & ItemCode & ": " & objMain.objCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Else
                        objMain.objApplication.StatusBar.SetText("Price updated for " & ItemCode & " with VAT", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    End If
                End If

                ' Update oDS3 datasource (TNX_ITEMUP_C2)
                '     oDS3.SetValue("U_Price", oDS3.Offset, PriceWithVAT)

                rs.MoveNext()
            End While

        Catch ex As Exception
            objForm.Items.Item("Item_13").Enabled = True
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub


    Sub UpdateSUGARTAX1(ByVal DocEntry As String)

        Try
            '===============================
            ' 🔹 STEP 1: GET PRICE LIST ID
            '===============================
            Dim PriceListId As Integer = -1

            Dim rsPL As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim PriceListName As String = "AL JAZEERA PURCHASE"
            PriceListName = PriceListName.Replace("'", "''")

            Dim qPL As String = "SELECT ""ListNum"" FROM OPLN WHERE ""ListName""='" & PriceListName & "'"

            rsPL.DoQuery(qPL)

            If rsPL.RecordCount = 0 Then
                objMain.objApplication.StatusBar.SetText("Price List not found", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            PriceListId = CInt(rsPL.Fields.Item("ListNum").Value)

            '===============================
            ' 🔹 STEP 2: GET ITEMS + SUGARTAX
            '===============================
            Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            DocEntry = DocEntry.Replace("'", "''")

            Dim Query As String = "SELECT T1.""U_ITMCD"", T1.""U_SuTaxAmt"" " &
                              "FROM ""@TNX_ITEMUP"" T0 " &
                              "INNER JOIN ""@TNX_ITEMUP_C0"" T1 ON T0.""DocEntry""=T1.""DocEntry"" " &
                              "WHERE T0.""DocNum""='" & DocEntry & "' AND T1.""U_ITMCD""<>''"

            rs.DoQuery(Query)

            If rs.RecordCount = 0 Then Exit Sub

            '===============================
            ' 🔹 STEP 3: UPDATE PRICE
            '===============================
            Dim oItem As SAPbobsCOM.Items

            For i As Integer = 1 To rs.RecordCount

                Dim ItemCode As String = rs.Fields.Item("U_ITMCD").Value.ToString()

                Dim SugarTax As Double = 0
                If Not IsDBNull(rs.Fields.Item("U_SuTaxAmt").Value) Then
                    SugarTax = CDbl(rs.Fields.Item("U_SuTaxAmt").Value)
                End If

                oItem = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

                If oItem.GetByKey(ItemCode) Then

                    ' 🔥 Update selected PriceList
                    For p As Integer = 0 To oItem.PriceList.Count - 1

                        oItem.PriceList.SetCurrentLine(p)

                        If oItem.PriceList.PriceList = PriceListId Then

                            ' ✅ ONLY SugarTax as price (Base Price)
                            oItem.PriceList.Price = SugarTax

                            Exit For
                        End If

                    Next

                    Dim RetVal As Integer = oItem.Update

                    If RetVal <> 0 Then
                        Dim ErrMsg As String = objMain.objCompany.GetLastErrorDescription()
                        objMain.objApplication.StatusBar.SetText("Error: " & ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Else
                        objMain.objApplication.StatusBar.SetText("Price updated for " & ItemCode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    End If

                End If

                rs.MoveNext()

            Next

        Catch ex As Exception
            objForm.Items.Item("Item_13").Enabled = True
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    'Public Sub New_Barcode_Generation(ByVal ItemCode As String, ByVal Itemname As String)
    '    Try
    '        ' Check if ItemCode is valid or not empty before proceeding
    '        If String.IsNullOrEmpty(ItemCode) Then
    '            objMain.objApplication.StatusBar.SetText("ItemCode is empty or invalid", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Return
    '        End If

    '        ' Query the OBCD table to check for the item code
    '        StrQry = "SELECT * FROM OBCD T0 WHERE T0.""ItemCode"" ='" & ItemCode & "'"
    '        oRec = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oRec.DoQuery(StrQry)

    '        If oRec.RecordCount = 0 Then
    '            ' Query the @NT_BRST table if no record found for ItemCode in OBCD
    '            StrQry = "SELECT * FROM ""@NT_BRST"""
    '            oRec = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '            oRec.DoQuery(StrQry)

    '            If oRec.RecordCount > 0 Then
    '                ' Generate the barcode based on the retrieved information
    '                Dim strlen As String = oRec.Fields.Item("U_Start1").Value.ToString()
    '                Dim RevStr As String = StrReverse(strlen)
    '                strlen = strlen.Length.ToString()

    '                ' Implement the checksum calculation
    '                Dim calcvalue As Double = 0
    '                Dim JumpNumber As Integer = 3
    '                For i As Integer = 0 To RevStr.Length - 1
    '                    If JumpNumber = 3 Then
    '                        calcvalue += Convert.ToDouble(RevStr.Substring(i, 1)) * JumpNumber
    '                        JumpNumber = 1
    '                    Else
    '                        calcvalue += Convert.ToDouble(RevStr.Substring(i, 1)) * JumpNumber
    '                        JumpNumber = 3
    '                    End If
    '                Next

    '                ' Handle potential errors in the checksum calculation
    '                Try
    '                    calcvalue = calcvalue / Convert.ToDouble(strlen)
    '                    Dim AryCalc() As String = Convert.ToString(calcvalue).Split(".")
    '                    If Convert.ToDouble(AryCalc(1)) > 0 Then
    '                        calcvalue = Convert.ToDouble(strlen) - Convert.ToDouble(AryCalc(1))
    '                    End If
    '                Catch ex As Exception
    '                    calcvalue = 0
    '                End Try

    '                ' Create the barcode
    '                Dim oAutoBarcode As String = oRec.Fields.Item("U_Itprefix").Value & oRec.Fields.Item("U_Start1").Value & calcvalue.ToString()

    '                ' Update the U_Start1 value after barcode creation
    '                Dim Startt As Long = Convert.ToUInt64(oRec.Fields.Item("U_Start1").Value) + 13
    '                StrQry = "UPDATE ""@NT_BRST"" SET ""U_Start1"" ='" & Startt & "' WHERE ""Code""='" & oRec.Fields.Item("Code").Value & "'"
    '                oRec = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                oRec.DoQuery(StrQry)

    '                ' Add the barcode to the BarCode service
    '                Try
    '                    Dim oUOMRecord As String = "SELECT ""UomEntry"" FROM OUOM WHERE ""UomEntry"" = '3'"
    '                    Dim oUOMRecord1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                    oUOMRecord1.DoQuery(oUOMRecord)

    '                    If oUOMRecord1.RecordCount = 0 Then
    '                        objMain.objApplication.StatusBar.SetText("UOM Entry not found", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                        Return
    '                    End If

    '                    Dim oCompanyService As SAPbobsCOM.ICompanyService = objMain.objCompany.GetCompanyService()
    '                    Dim oBarCodesService As SAPbobsCOM.BarCodesService = CType(oCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.BarCodesService), SAPbobsCOM.BarCodesService)
    '                    Dim oBarCode As SAPbobsCOM.BarCode = CType(oBarCodesService.GetDataInterface(SAPbobsCOM.BarCodesServiceDataInterfaces.bsBarCode), SAPbobsCOM.BarCode)
    '                    oBarCode.ItemNo = ItemCode
    '                    oBarCode.BarCode = oAutoBarcode
    '                    oBarCode.UoMEntry = 3
    '                    oBarCode.FreeText = "test4"

    '                    ' Add barcode to the service
    '                    oBarCodesService.Add(oBarCode)

    '                Catch ex As Exception
    '                    objMain.objApplication.StatusBar.SetText("Error adding barcode: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                End Try
    '            Else
    '                objMain.objApplication.StatusBar.SetText("Barcode Setup details missing for this ItemCode or ItemGroup...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            End If
    '        Else
    '            objMain.objApplication.StatusBar.SetText("Barcode Already Exists for the ItemCode...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        End If
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText("An error occurred: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '    End Try
    'End Sub


    Public Sub New_Barcode_Generation4(ByVal ItemCode As String, ByVal Itemname As String)
        Try
            StrQry = "SELECT *  FROM OBCD T0 WHERE T0.""ItemCode"" ='" & ItemCode & "'"
            oRec = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery(StrQry)
            If oRec.RecordCount = 0 Then
                StrQry = "SELECT * FROM ""@NT_BRST"""
                oRec = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRec.DoQuery(StrQry)
                If oRec.RecordCount > 0 Then
                    Dim strlen, RevStr As String
                    Dim AryCalc() As String
                    Dim calcvalue As Double = 0
                    strlen = oRec.Fields.Item("U_Start1").Value.ToString()
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
                            calcvalue = Convert.ToDouble(strlen) - Convert.ToDouble(AryCalc(1))
                        End If
                    Catch ex As Exception
                        calcvalue = 0
                    End Try
                    oAutoBarcode = oRec.Fields.Item("U_Itprefix").Value & oRec.Fields.Item("U_Start1").Value & calcvalue.ToString()
                    ''oItem.BarCode = oAutoBarcode
                    Startt = Convert.ToUInt64(oRec.Fields.Item("U_Start1").Value) + 13
                    StrQry = "UPDATE ""@NT_BRST"" SET ""U_Start1"" ='" & Startt & "',""U_End"" ='" & Endd & "' Where ""Code""='" & oRec.Fields.Item("Code").Value & "' "
                    oRec = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRec.DoQuery(StrQry)
                    ''----barcode--
                    Try

                        'Dim oUOMRecord As String = "SELECT ""UomEntry"" FROM OUOM WHERE ""UomEntry"" = '3' "
                        'Dim oUOMRecord1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'oUOMRecord1.DoQuery(oUOMRecord)
                        'If oUOMRecord1.RecordCount = 0 Then
                        '    objMain.objApplication.StatusBar.SetText("Uom can not be empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        'End If
                        Dim oCompanyService As SAPbobsCOM.ICompanyService = objMain.objCompany.GetCompanyService()
                        Dim oBarCodesService As SAPbobsCOM.BarCodesService = CType(oCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.BarCodesService), SAPbobsCOM.BarCodesService)
                        Dim oBarCode As SAPbobsCOM.BarCode = CType(oBarCodesService.GetDataInterface(SAPbobsCOM.BarCodesServiceDataInterfaces.bsBarCode), SAPbobsCOM.BarCode)
                        oBarCode.ItemNo = ItemCode
                        oBarCode.BarCode = oAutoBarcode
                        oBarCode.UoMEntry = 3
                        oBarCode.FreeText = Itemname
                        '' Add Barcode
                        oBarCodesService.Add(oBarCode)
                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText("Uom can not be empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try

                    ''----barode--


                    'Dim itmc As String = "SELECT ""ItemCode"" FROM ""OITM"" WHERE ""ItemCode"" = '" & ItemCode & "'")
                    'Dim oItemRec As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'oItemRec.DoQuery(itmc)
                    'Try
                    '    If oItemRec.RecordCount = 0 Then
                    '        objMain.objApplication.StatusBar.SetText("ItemCode '" & ItemCode & "' not found in OITM.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        Exit Sub
                    '    End If
                    'Catch ex As Exception
                    'Dim oBarcode1 As SAPbobsCOM.ItFreeTextemBarCodes
                    'Dim oCompService As SAPbobsCOM.ICompanyService = objMain.objCompany.GetCompanyService()
                    'Dim oBarCodesService As SAPbobsCOM.BarCodesService = CType(oCompService.GetBusinessService(SAPbobsCOM.ServiceTypes.BarCodesService), SAPbobsCOM.BarCodesService)
                    'Dim oBarCode As SAPbobsCOM.BarCode = CType(oBarCodesService.GetDataInterface(SAPbobsCOM.BarCodesServiceDataInterfaces.bsBarCode), SAPbobsCOM.BarCode)
                    '' Set Barcode Text or                     'End Try

                    'oBarCode.ItemNo = ItemCode
                    'oBarCode.BarCode = oAutoBarcode
                    'oBarCode.UoMEntry = -1
                    'oBarCode.FreeText = "test4"
                    '' Add Barcode
                    'oBarCodesService.Add(oBarCode)

                Else
                    objMain.objApplication.StatusBar.SetText("Barcode Setup details missing for this ItemCode or ItemGroup...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If

            Else
                objMain.objApplication.StatusBar.SetText("Barcode Already Exist...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
    End Sub
    Public Function New_Barcode_Generation(ByVal ItemCode As String, ByVal Itemname As String) As String
        Try
            StrQry = "SELECT *  FROM OBCD T0 WHERE T0.""ItemCode"" ='" & ItemCode & "'"
            oRec = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery(StrQry)
            If oRec.RecordCount = 0 Then
                StrQry = "SELECT * FROM ""@NT_BRST"""
                oRec = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRec.DoQuery(StrQry)
                If oRec.RecordCount > 0 Then
                    Dim strlen, RevStr As String
                    Dim AryCalc() As String
                    Dim calcvalue As Double = 0
                    strlen = oRec.Fields.Item("U_Start1").Value.ToString()
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
                            calcvalue = Convert.ToDouble(strlen) - Convert.ToDouble(AryCalc(1))
                        End If
                    Catch ex As Exception
                        calcvalue = 0
                    End Try
                    ' oAutoBarcode = oRec.Fields.Item("U_Itprefix").Value & oRec.Fields.Item("U_Start1").Value & calcvalue.ToString()
                    If Convert.ToDouble(calcvalue) > 0 Then
                        oAutoBarcode = oRec.Fields.Item("U_Itprefix").Value & oRec.Fields.Item("U_Start1").Value & calcvalue.ToString()
                    Else

                        oAutoBarcode = oRec.Fields.Item("U_Itprefix").Value & oRec.Fields.Item("U_Start1").Value & calcvalue.ToString()
                        objMain.objApplication.StatusBar.SetText("13 th Digit is zero, Kindly Generate Again..." & oAutoBarcode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oAutoBarcode = "0"
                        Return ""
                    End If
                    strlen = oAutoBarcode
                    strlen = strlen.Length
                    If strlen = 13 Then
                        oAutoBarcode = oAutoBarcode
                    Else
                        Return ""
                        objMain.objApplication.StatusBar.SetText("Length of the Barcode is morethan 13 digit on ..." & oAutoBarcode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                    Startt = Convert.ToUInt64(oRec.Fields.Item("U_Start1").Value) + 13

                    'Endd = Convert.ToUInt64(oRec.Fields.Item("U_End").Value) + 2
                    'StrQry = "UPDATE ""@NT_BRST"" SET ""U_Start1"" ='" & Startt & "',""U_End"" ='" & Endd & "' Where ""Code""='" & oRec.Fields.Item("Code").Value & "' "
                    'oRec = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'oRec.DoQuery(StrQry)
                    Return oAutoBarcode
                    'Else
                    '    objMain.objApplication.StatusBar.SetText("Barcode Setup details missing for this ItemCode or ItemGroup...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    'End If
                    Try
                        StrQry = "SELECT *  FROM OBCD T0 WHERE T0.""BcdCode"" ='" & oAutoBarcode & "'"
                        oRec = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRec.DoQuery(StrQry)
                        If oRec.RecordCount > 0 Then
                            objMain.objApplication.StatusBar.SetText("This Barcode already in Use, Kindly Generate Again..." & oAutoBarcode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return ""
                        End If
                    Catch ex As Exception
                        Return ""
                        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    End Try

                    ''oItem.BarCode = oAutoBarcode
                    Return oAutoBarcode
                Else
                    objMain.objApplication.StatusBar.SetText("Barcode Setup details missing for this ItemCode or ItemGroup...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return ""
                End If

            Else
                objMain.objApplication.StatusBar.SetText("Barcode Already Exist...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return ""
            End If
        Catch ex As Exception
            Return ""
            objForm.Items.Item("Item_13").Enabled = True
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
    End Function

    Public Function New_Barcode_GenerationWithBarcodeDescription(ByVal ItemCode As String, ByVal Itemname As String) As String
        Try
            StrQry = "SELECT *  FROM OBCD T0 WHERE T0.""ItemCode"" ='" & ItemCode & "'"
            oRec = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery(StrQry)
            If oRec.RecordCount = 0 Then
                StrQry = "SELECT * FROM ""@NT_BRST"""
                oRec = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRec.DoQuery(StrQry)
                If oRec.RecordCount > 0 Then
                    Dim strlen, RevStr As String
                    Dim AryCalc() As String
                    Dim calcvalue As Double = 0
                    'strlen = oRec.Fields.Item("U_Start1").Value.ToString()
                    strlen = oRec.Fields.Item("U_Itprefix").Value & oRec.Fields.Item("U_Start1").Value.ToString()
                    RevStr = StrReverse(strlen)
                    strlen = strlen.Length
                    Dim baselength As Integer = strlen.Length
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
                        If AryCalc.Length > 1 AndAlso Convert.ToDouble(AryCalc(1)) > 0 Then
                            ''     If Convert.ToDouble(AryCalc(1)) > 0 Then
                            calcvalue = Convert.ToDouble(strlen) - Convert.ToDouble(AryCalc(1).Substring(0, 1))
                            If calcvalue > 9 Then
                                calcvalue = calcvalue Mod 10
                            End If
                            If calcvalue = 0 Then
                                calcvalue = 1
                            End If
                        End If
                    Catch ex As Exception
                        calcvalue = 0
                    End Try
                    If Convert.ToDouble(calcvalue) > 0 Then
                        oAutoBarcode = oRec.Fields.Item("U_Itprefix").Value & oRec.Fields.Item("U_Start1").Value & calcvalue.ToString()
                    Else

                        oAutoBarcode = oRec.Fields.Item("U_Itprefix").Value & oRec.Fields.Item("U_Start1").Value & calcvalue.ToString()
                        objMain.objApplication.StatusBar.SetText("13 th Digit is zero, Kindly Generate Again..." & oAutoBarcode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oAutoBarcode = "0"
                        Return ""
                    End If
                    strlen = oAutoBarcode
                    strlen = strlen.Length
                    If strlen = 13 Then
                        oAutoBarcode = oAutoBarcode
                    Else
                        'Return ""
                        objMain.objApplication.StatusBar.SetText("Length of the Barcode is morethan 13 digit on ..." & oAutoBarcode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                    Startt = Convert.ToUInt64(oRec.Fields.Item("U_Start1").Value) + 13

                    'Endd = Convert.ToUInt64(oRec.Fields.Item("U_End").Value) + 2
                    StrQry = "UPDATE ""@NT_BRST"" SET ""U_Start1"" ='" & Startt & "',""U_End"" ='" & Endd & "' Where ""Code""='" & oRec.Fields.Item("Code").Value & "' "
                    oRec = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRec.DoQuery(StrQry)
                    Return oAutoBarcode
                    'Else
                    '    objMain.objApplication.StatusBar.SetText("Barcode Setup details missing for this ItemCode or ItemGroup...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    'End If
                    Try
                        StrQry = "SELECT *  FROM OBCD T0 WHERE T0.""BcdCode"" ='" & oAutoBarcode & "'"
                        oRec = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRec.DoQuery(StrQry)
                        If oRec.RecordCount > 0 Then
                            objMain.objApplication.StatusBar.SetText("This Barcode already in Use, Kindly Generate Again..." & oAutoBarcode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return ""
                        End If
                    Catch ex As Exception
                        Return ""
                        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    End Try

                    ''oItem.BarCode = oAutoBarcode
                    Return oAutoBarcode
                Else
                    objMain.objApplication.StatusBar.SetText("Barcode Setup details missing for this ItemCode or ItemGroup...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return ""
                End If

            Else
                objMain.objApplication.StatusBar.SetText("Barcode Already Exist...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return ""
            End If
        Catch ex As Exception
            objForm.Items.Item("Item_13").Enabled = True
            Return ""
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
    End Function
    Public Sub Generate_Barcode(ByVal FormUID As String, ByVal ItemCode As String, ByVal scaleitm As String, ByVal subgrp As String, ByVal subsubgrp As String, ByVal Itemgrp As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim rSet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ' Dim specific As Matrix = DirectCast(objMain.objApplication.Forms.ActiveForm.Items.Item("1470000013").Specific, Matrix)
            Dim businessObject As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'Dim item As String = objForm.Items.Item("5").Specific.Value
            Dim item As String = ItemCode
            Dim str1 As String = ""
            Dim scale As String = scaleitm
            Dim subgrp1 As String = subgrp
            Dim subsubgrp1 As String = subsubgrp
            Dim Itemgrp1 As String = Itemgrp
            If scale = 2 Then


                '---------------------------------------------------------------
                '---------------------------------------------------------------

                'Dim DigitsQry As String = "Select Case When ""U_AcceptWMQty""='1' Then '6' When ""U_AcceptWMQty""='2' Then '13' End FROM OITM WHERE ""ItemCode"" = '" & item & "'"
                'Dim oRsDigits As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oRsDigits.DoQuery(DigitsQry)

                'If oRsDigits.Fields.Item(0).Value <> "6" Then '------------13-------------------------------------2

                'Dim ItemGroupQuery As String = "SELECT * FROM ""@NT_BRST""  T2 WHERE T2.""U_Itmgrpnm"" IN " &
                '"(SELECT (SELECT T1.""ItmsGrpNam"" FROM OITB T1 WHERE T1.""ItmsGrpCod""=T0.""ItmsGrpCod"") FROM OITM T0 WHERE T0.""ItemCode""='" & item & "') " &
                '" And  T2.""U_Scale""=(SELECT ""U_AcceptWMQty"" FROM ""OITM"" Where ""ItemCode""='" & item & "') and  T2.""U_SubGrp"" =(SELECT ""U_VSubGroup"" FROM ""OITM"" Where ""ItemCode""='" & item & "')"
                'Dim oRsItemGroup As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oRsItemGroup.DoQuery(ItemGroupQuery)
                Dim NewBarcode As String = ""
                'Dim ItemGroup As String = ""
                'Dim SubGroup As String = ""
                'If oRsItemGroup.RecordCount = 0 Then
                '    ItemGroup = "ALL"
                'Else
                '    ItemGroup = oRsItemGroup.Fields.Item("U_Itmgrpnm").Value
                '    SubGroup = oRsItemGroup.Fields.Item("U_SubGrp").Value
                'End If
                Dim Barcode As String = "SELECT T0.""U_Prefix"", T0.""U_Itprefix"", T0.""U_Start"",T0.""U_Start""+1 AS ""NextNo"" FROM ""@NT_BRST""  T0 WHERE T0.""U_Itmgrpnm"" ='" & Itemgrp1 & "' and T0.""U_SubGrp"" ='" & subgrp1 & "'and ""U_Scale""='2' "
                Dim oRsBarcode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
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
                Dim NextBarcode As String = "UPDATE ""@NT_BRST"" T0 SET T0.""U_Start""='" & NextNo & "' WHERE T0.""U_Itmgrpnm"" ='" & Itemgrp1 & "' and T0.""U_SubGrp"" ='" & subgrp1 & "' and ""U_Scale""='2' "
                Dim oRsNextBarcode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
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
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try

            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK

                    '                    If (pVal.ItemUID = "ItmDetMtx" Or pVal.ItemUID = "UOMMtx" Or pVal.ItemUID = "BarMtx") And pVal.BeforeAction = False Then
                    If (pVal.ItemUID = "ItmDetMtx" Or pVal.ItemUID = "UOMMtx" Or pVal.ItemUID = "BarMtx" Or pVal.ItemUID = "Mtx5" Or pVal.ItemUID = "Mtx6" Or pVal.ItemUID = "Mtx7") And pVal.BeforeAction = False Then
                        MATRIXS = pVal.ItemUID
                    End If

                    If pVal.ItemUID = "Item_11" And pVal.BeforeAction = False Then


                        objForm.Items.Item("Item_11").Enabled = False
                        Dim FileSaveWithPath As String = objForm.Items.Item("ATTPATH").Specific.Value

                        If FileSaveWithPath <> "" Then

                            Me.ItemMasterDetailsExcelUpload(FileSaveWithPath)

                        Else
                            objMain.objApplication.StatusBar.SetText("Please Select File Path...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Sub
                            objForm.Items.Item("Item_11").Enabled = True
                        End If
                        objForm.Items.Item("Item_11").Enabled = True
                    End If

                    'Calcultions ItemCode tab 
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                    ' Markup & Margin Calculation
                    'If pVal.ItemUID = "ItmDetMtx" AndAlso (pVal.ColUID = "Selling" Or pVal.ColUID = "Cost") AndAlso pVal.BeforeAction = False AndAlso pVal.Row > 0 Then

                    '    Dim objMatrix As SAPbouiCOM.Matrix
                    '    objMatrix = objForm.Items.Item("ItmDetMtx").Specific

                    '    Dim selling As Double = 0
                    '    Dim cost As Double = 0

                    '    Double.TryParse(objMatrix.Columns.Item("Selling").Cells.Item(pVal.Row).Specific.Value, selling)
                    '    Double.TryParse(objMatrix.Columns.Item("Cost").Cells.Item(pVal.Row).Specific.Value, cost)

                    '    objForm.Freeze(True)

                    '    ' -------- MARKUP --------
                    '    If cost > 0 Then
                    '        objMatrix.Columns.Item("MARKUP").Cells.Item(pVal.Row).Specific.Value = Math.Round(((selling - cost) / cost) * 100, 2).ToString()
                    '    Else
                    '        objMatrix.Columns.Item("MARKUP").Cells.Item(pVal.Row).Specific.Value = "0"
                    '    End If

                    '    ' -------- MARGIN --------
                    '    If selling > 0 Then
                    '        objMatrix.Columns.Item("MARGIN").Cells.Item(pVal.Row).Specific.Value = Math.Round(((selling - cost) / selling) * 100, 2).ToString()
                    '    Else
                    '        objMatrix.Columns.Item("MARGIN").Cells.Item(pVal.Row).Specific.Value = "0"
                    '    End If

                    '    objForm.Freeze(False)

                    'End If
                    If pVal.ItemUID = "BarMtx" And pVal.BeforeAction = False Then
                        Select Case pVal.ColUID

                            Case "SUTAXAMT", "ETAX", "Price"
                                CalculateMarginAndMarkup_LostFocus1(pVal.Row)

                        End Select



                    End If

                    'If pVal.ItemUID = "BarMtx" And (pVal.ColUID = "SUTAXAMT" Or pVal.ColUID = "ETAX") And pVal.BeforeAction = False Then
                    '    Dim oMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("BarMtx").Specific
                    '    Dim Row As Integer = pVal.Row

                    '    Dim SUTAXAMT As Double = 0
                    '    Dim ETAX As Double = 0

                    '    Double.TryParse(oMatrix.Columns.Item("SUTAXAMT").Cells.Item(Row).Specific.Value, SUTAXAMT)
                    '    Double.TryParse(oMatrix.Columns.Item("ETAX").Cells.Item(Row).Specific.Value, ETAX)

                    '    Dim TotalPrice As Double = SUTAXAMT + ETAX

                    '    oMatrix.Columns.Item("Price").Cells.Item(Row).Specific.Value = TotalPrice.ToString()

                    'End If

                    If pVal.ItemUID = "Mtx5" And (pVal.ColUID = "OldPrice") And pVal.BeforeAction = False Then

                        CalculateMarginAndMarkup_LostFocus1()

                    End If

                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    Dim objmatrixBr As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx5").Specific
                    If pVal.ItemUID = "Mtx5" And pVal.ColUID = "BrnchId" And pVal.BeforeAction = False Then
                        objComboBox = objmatrixBr.Columns.Item("BrnchId").Cells.Item(pVal.Row).Specific
                        If objComboBox.Selected.Value <> "" Then
                            Try
                                objForm.Freeze(True)

                                Dim BrachId As String = objComboBox.Selected.Value
                                Dim qStr2 As String = "Select ""BPLName"" FROM ""OBPL"" WHERE ""BPLId""='" + BrachId + "' "
                                Dim rsBN As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                rsBN.DoQuery(qStr2)
                                If rsBN.RecordCount > 0 Then
                                    objmatrixBr.Columns.Item("BrachNm").Cells.Item(pVal.Row).Specific.Value = rsBN.Fields.Item(0).Value
                                End If
                                objForm.Freeze(False)
                            Catch ex As Exception
                                objForm.Freeze(False)
                            End Try
                        End If
                    End If
                    If pVal.ItemUID = "Mtx5" And pVal.ColUID = "PriceId" And pVal.BeforeAction = False Then
                        objComboBox = objmatrixBr.Columns.Item("PriceId").Cells.Item(pVal.Row).Specific
                        If objComboBox.Selected.Value <> "" Then
                            Try
                                objForm.Freeze(True)
                                Dim PriceId As String = objComboBox.Selected.Value
                                Dim qStr3 As String = "Select ""ListName"" FROM ""OPLN"" WHERE ""ListNum""='" + PriceId + "' "
                                Dim rsPt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                rsPt.DoQuery(qStr3)
                                If rsPt.RecordCount > 0 Then
                                    objmatrixBr.Columns.Item("PriceLt").Cells.Item(pVal.Row).Specific.Value = rsPt.Fields.Item(0).Value
                                End If
                                objForm.Freeze(False)
                            Catch ex As Exception
                                objForm.Freeze(False)
                            End Try
                        End If
                    End If
                    Dim objmatrixvn As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx6").Specific
                    If pVal.ItemUID = "Mtx6" And pVal.ColUID = "PriceId" And pVal.BeforeAction = False Then
                        objComboBox = objmatrixvn.Columns.Item("PriceId").Cells.Item(pVal.Row).Specific
                        If objComboBox.Selected.Value <> "" Then
                            Try
                                objForm.Freeze(True)
                                Dim PriceId As String = objComboBox.Selected.Value
                                Dim qStr3 As String = "Select ""ListName"" FROM ""OPLN"" WHERE ""ListNum""='" + PriceId + "' "
                                Dim rsPt1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                rsPt1.DoQuery(qStr3)
                                If rsPt1.RecordCount > 0 Then
                                    objmatrixvn.Columns.Item("PriceLt").Cells.Item(pVal.Row).Specific.Value = rsPt1.Fields.Item(0).Value
                                End If
                                objForm.Freeze(False)
                            Catch ex As Exception
                                objForm.Freeze(False)
                            End Try
                        End If
                    End If
                    Dim objmatrixmin As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx7").Specific
                    If pVal.ItemUID = "Mtx7" And pVal.ColUID = "PriceId" And pVal.BeforeAction = False Then
                        objComboBox = objmatrixmin.Columns.Item("PriceId").Cells.Item(pVal.Row).Specific
                        If objComboBox.Selected.Value <> "" Then
                            Try
                                objForm.Freeze(True)
                                Dim PriceId As String = objComboBox.Selected.Value
                                Dim qStr3 As String = "Select ""ListName"" FROM ""OPLN"" WHERE ""ListNum""='" + PriceId + "' "
                                Dim rsPt2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                rsPt2.DoQuery(qStr3)
                                If rsPt2.RecordCount > 0 Then
                                    objmatrixmin.Columns.Item("PriceLt").Cells.Item(pVal.Row).Specific.Value = rsPt2.Fields.Item(0).Value
                                End If
                                objForm.Freeze(False)
                            Catch ex As Exception
                                objForm.Freeze(False)
                            End Try
                        End If
                    End If

                    Dim objmatrix1 As SAPbouiCOM.Matrix = objForm.Items.Item("ItmDetMtx").Specific
                    If pVal.ItemUID = "ItmDetMtx" And pVal.ColUID = "ITEMG" And pVal.BeforeAction = False Then
                        objComboBox = objmatrix1.Columns.Item("ITEMG").Cells.Item(pVal.Row).Specific

                        If objComboBox.Selected.Value <> "" Then
                            Try
                                objForm.Freeze(True)

                                Dim item As String = objComboBox.Selected.Value
                                Dim qStr2 As String = "Select ""ItmsGrpCod"" FROM ""OITB"" WHERE ""ItmsGrpNam""='" + item + "' "
                                Dim rsuomqry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                rsuomqry.DoQuery(qStr2)

                                If rsuomqry.RecordCount > 0 Then
                                    Dim base As String = rsuomqry.Fields.Item("ItmsGrpCod").Value
                                    ClearCombo(objmatrix1.Columns.Item("SUBGR"), pVal.Row)
                                    ClearCombo(objmatrix1.Columns.Item("SSGP"), pVal.Row)
                                    oDS1 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C0")

                                    objmatrix1.GetLineData(pVal.Row)
                                    oDS1.SetValue("U_SUBGR", oDS1.Offset, "")
                                    oDS1.SetValue("U_SSGP", oDS1.Offset, "")
                                    objmatrix1.SetLineData(pVal.Row)
                                    Dim Qury1 As String = "SELECT T0.""U_SUBGROUP"" , T0.""U_SUBGROUP""  FROM ""@VR_SUBGROUP""  T0 where ""U_ITEMGROUP"" ='" & base & "' "
                                    Dim grpcod As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    grpcod.DoQuery(Qury1)
                                    ' Clear SubGroup & SubSubGroup

                                    objMain.objUtilities.MatrixComboBoxValues(objmatrix1.Columns.Item("SUBGR"), Qury1)
                                    objmatrix1.Columns.Item("SUBGR").ExpandType = SAPbouiCOM.BoExpandType.et_ValueDescription
                                    objmatrix1.Columns.Item("SUBGR").DisplayDesc = True
                                    objForm.Freeze(False)
                                End If

                                objForm.Freeze(False)
                            Catch ex As Exception
                                objForm.Freeze(False)
                            End Try

                        End If


                    End If

                    If pVal.ItemUID = "ItmDetMtx" And pVal.ColUID = "SUBGR" And pVal.BeforeAction = False Then
                        Try
                            objForm.Freeze(True)
                            objComboBox = objmatrix1.Columns.Item("SUBGR").Cells.Item(pVal.Row).Specific
                            objComboBox1 = objmatrix1.Columns.Item("ITEMG").Cells.Item(pVal.Row).Specific
                            If objComboBox.Selected.Value <> "" And objComboBox1.Selected.Value <> "" Then
                                Dim item As String = objComboBox1.Selected.Value
                                Dim subgRP As String = objComboBox.Selected.Value
                                Dim Qur1 As String = "SELECT T0.""Name"", T0.""Name"" FROM ""@VR_SUBSUBGROUP""  T0 where ""U_SUBGROUP""='" & subgRP & "'  "
                                Dim grpcod1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                grpcod1.DoQuery(Qur1)
                                ClearCombo(objmatrix1.Columns.Item("SSGP"), pVal.Row)
                                oDS1 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C0")

                                ' Dim oSub As SAPbouiCOM.ComboBox = objmatrix1.Columns.Item("SUBGR").Cells.Item(pVal.Row).Specific
                                objmatrix1.GetLineData(pVal.Row)
                                oDS1.SetValue("U_SSGP", oDS1.Offset, "")
                                objmatrix1.SetLineData(pVal.Row)

                                '    Dim subGrp As String = oSub.Selected.Value.Replace("'", "''")


                                objMain.objUtilities.MatrixComboBoxValues(objmatrix1.Columns.Item("SSGP"), Qur1)
                                objmatrix1.Columns.Item("SSGP").ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                                objmatrix1.Columns.Item("SSGP").DisplayDesc = True
                                objForm.Freeze(False)
                            End If
                        Catch ex As Exception
                            objForm.Freeze(False)
                        End Try
                    End If

                    If pVal.ItemUID = "ItmDetMtx" And pVal.ColUID = "PVAT" And pVal.BeforeAction = False Then

                        CalculateMarginAndMarkup_Fast()

                    End If


                    If pVal.ItemUID = "BarMtx" And pVal.ColUID = "STAX" And pVal.BeforeAction = False Then

                        CalculateMarginAndMarkup_Fast()

                    End If
                    If pVal.ItemUID = "ItmDetMtx" And pVal.ColUID = "SugarTax" And pVal.BeforeAction = False Then
                        Try
                            objForm.Freeze(True)

                            Dim oCombo As SAPbouiCOM.ComboBox
                            oCombo = objmatrix1.Columns.Item("SugarTax").Cells.Item(pVal.Row).Specific

                            ' 🔹 Find column index manually
                            Dim colIndex As Integer = -1
                            For i As Integer = 0 To objmatrix1.Columns.Count - 1
                                If objmatrix1.Columns.Item(i).UniqueID = "SuTaxAmt" Then
                                    colIndex = i
                                    Exit For
                                End If
                            Next

                            If colIndex <> -1 AndAlso oCombo.Selected IsNot Nothing Then
                                If oCombo.Selected.Value.ToUpper = "YES" Then
                                    objmatrix1.CommonSetting.SetCellEditable(pVal.Row, colIndex, True)
                                ElseIf oCombo.Selected.Value.ToUpper = "NO" Then
                                    oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C0")
                                    objmatrix1.GetLineData(pVal.Row)
                                    oDS1.SetValue("U_SuTaxAmt", oDS1.Offset, "0")
                                    objmatrix1.SetLineData(pVal.Row)
                                    objmatrix1.CommonSetting.SetCellEditable(pVal.Row, colIndex, False)

                                End If
                            End If

                        Catch ex As Exception
                            objForm.Freeze(False)
                            ' Handle error if needed
                        Finally
                            objForm.Freeze(False)
                        End Try
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDS = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP")
                    oDS1 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C0")
                    oDS2 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C1")
                    oDS3 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C2")

                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    Dim objmatrix1 As SAPbouiCOM.Matrix = objForm.Items.Item("ItmDetMtx").Specific


                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                    If oCFL.UniqueID = "CFL_VEND" And pVal.BeforeAction = True Then

                        Me.CFLFilter(objForm.UniqueID, "CFL_VEND")

                    End If
                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        If oCFL.UniqueID = "CFL_VEND" Then
                            objmatrix1.GetLineData(pVal.Row)
                            oDS1.SetValue("U_CARD", oDS1.Offset, oDT.GetValue("CardCode", 0))
                            objmatrix1.SetLineData(pVal.Row)
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Dim oMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("ItmDetMtx").Specific
                        Dim rsValidate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        For i As Integer = 1 To oMatrix.VisualRowCount

                            Dim SubGroup As String = ""
                            Dim SubSubGroup As String = ""

                            SubGroup = Trim(CType(oMatrix.Columns.Item("SUBGR").Cells.Item(i).Specific, SAPbouiCOM.ComboBox).Value)
                            SubSubGroup = Trim(CType(oMatrix.Columns.Item("SSGP").Cells.Item(i).Specific, SAPbouiCOM.ComboBox).Value)

                            If SubGroup <> "" AndAlso SubSubGroup <> "" Then

                                Dim strValidateQry As String =
                                    "SELECT COUNT(*) AS ""Cnt"" " &
                                    "FROM ""@VR_SUBSUBGROUP"" " &
                                    "WHERE ""U_SUBGROUP"" = '" & SubGroup.Replace("'", "''") & "' " &
                                    "AND ""Name"" = '" & SubSubGroup.Replace("'", "''") & "'"

                                rsValidate.DoQuery(strValidateQry)

                                If CInt(rsValidate.Fields.Item("Cnt").Value) = 0 Then

                                    objMain.objApplication.StatusBar.SetText("Please select correct Sub Sub Group '" & SubSubGroup & "' for Sub Group '" & SubGroup & "' at Row " & i, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                                    BubbleEvent = False
                                    Exit Sub

                                End If

                            End If

                        Next
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And
                         (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then

                        Me.SetDefault(objForm.UniqueID)
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        DeleteEmptyRows_MatrixWise(objForm)
                        If Me.Validate1() = False Then

                            BubbleEvent = False
                        End If
                    End If

                    If pVal.ItemUID = "1" AndAlso pVal.BeforeAction = True AndAlso (objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                        Dim IsMarginIssue As Boolean = False

                        Dim objmatrixBr As SAPbouiCOM.Matrix = objForm.Items.Item("BarMtx").Specific

                        For i As Integer = 1 To objmatrixBr.RowCount

                            Dim Remarks As String = ""

                            Try
                                Remarks = CType(objmatrixBr.Columns.Item("Col_0").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value.Trim()
                            Catch ex As Exception
                            End Try

                            If Remarks = "Margin is not meeting" Then
                                IsMarginIssue = True
                                Exit For
                            End If

                        Next

                        If IsMarginIssue = True Then

                            Dim Ret As Integer

                            Ret = objMain.objApplication.MessageBox("Margin is below the defined level. Do you want to proceed?", 2, "Yes", "No")

                            If Ret = 2 Then
                                BubbleEvent = False
                                Exit Sub
                            End If

                        End If

                    End If


                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                              (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        LostFocusFlag = True
                        DeleteEmptyRows_MatrixWise(objForm)
                        If Me.Validate(objForm.UniqueID, MATRIXS) = False Then
                            LostFocusFlag = False
                            BubbleEvent = False
                        End If
                        If Me.Validation(objForm.UniqueID) = False Then BubbleEvent = False

                        If Not ValidateItemGroup(objForm) Then
                            BubbleEvent = False
                            Exit Sub
                        End If

                        '------Madhu Sub abd sub sub group valiadtion
                        'If Not ValidateSubAndSubSubGroup(objForm) Then
                        '    BubbleEvent = False
                        '    Exit Sub
                        'End If

                        Dim status As String = objForm.Items.Item("STAT").Specific.value
                        Dim str As String = objForm.Items.Item("Item_19").Specific.value

                        If (status = "OPEN" Or status = "O") And str = "Yes" Then

                            'ONLY SUBMIT ENABLED
                            Try : objForm.Items.Item("Submit").Enabled = True : Catch : End Try
                        Else
                            Try : objForm.Items.Item("Submit").Enabled = False : Catch : End Try
                            ''objMain.objApplication.ActivateMenuItem("1304")

                        End If

                    End If


                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And
                              (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then


                        Dim status As String = objForm.Items.Item("STAT").Specific.value
                        Dim str As String = objForm.Items.Item("Item_19").Specific.value

                        If (status = "OPEN" Or status = "O") And str = "Yes" Then

                            'ONLY SUBMIT ENABLED
                            Try : objForm.Items.Item("Submit").Enabled = True : Catch : End Try
                        Else
                            Try : objForm.Items.Item("Submit").Enabled = False : Catch : End Try
                            ''objMain.objApplication.ActivateMenuItem("1304")

                        End If

                    End If




                    'madhu Preview button code here...@
                    If pVal.ItemUID = "Preview" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Dim ApplicationIDITMUploadPreview As String = oDS.GetValue("U_AppID", 0).Trim()

                        objMain.objItemMasterUploadPreview.CreateForm(ApplicationIDITMUploadPreview)

                    End If

                    'submit
                    If pVal.ItemUID = "Submit" And pVal.BeforeAction = False And SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Try
                            Dim DocNum As String = objForm.Items.Item("AppID").Specific.value
                            ' Dim PriceChangeRequestType As String = CType(objForm.Items.Item("TNXPUTP").Specific, SAPbouiCOM.ComboBox).Selected.Value
                            Dim DOCSTATUS As String = oDS.GetValue("U_STAT", oDS.Offset)
                            '----Madhu when Rejected document is resubmitted then status will be open instead of rejected that's why added open status also
                            If DOCSTATUS = "O" Then
                                Me.ApprovalTrigger("IMP", DocNum, "@TNX_ITEMUP", "U_AppID", "U_STAT", "ItemMasterUpload")
                            ElseIf DOCSTATUS = "R" Then

                                Dim rs As SAPbobsCOM.Recordset
                                rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                Dim NewAppID As String = ""
                                Dim Cnt As Integer = 1

                                'CHECK HOW MANY TIMES RESUBMITTED
                                rs.DoQuery("SELECT COUNT(*) + 1 AS ""CNT"" FROM ""@SBO_DD"" WHERE ""U_docnum"" LIKE '" & DocNum & "_%'")

                                If rs.RecordCount > 0 Then
                                    Cnt = rs.Fields.Item("CNT").Value
                                End If

                                NewAppID = DocNum & "_" & Cnt

                                'ONLY SEND NEW APP ID TO APPROVAL
                                Me.ApprovalTrigger("IMP", NewAppID, "@TNX_ITEMUP", "U_AppID", "U_STAT", "ItemMasterUpload")

                            ElseIf DOCSTATUS = "S" Then

                            ElseIf DOCSTATUS = "A" Then

                            End If

                        Catch ex As Exception
                        End Try

                    End If


                    'madhu link button to Item Listing
                    If pVal.BeforeAction = False AndAlso pVal.ItemUID = "Item_21" Then

                        Dim DocumentNumber As String = oDS.GetValue("U_AppIDIL", 0).Trim()

                        objMain.objItemList.CreateForm_Find(DocumentNumber)

                    End If


                    If pVal.ItemUID = "Item_13" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Dim DocEntry As String = objForm.Items.Item("DocNum").Specific.Value
                        objForm.Items.Item("Item_13").Enabled = False
                        '' Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.GetForm("ITEMUP", objMain.objApplication.Forms.ActiveForm.TypeCount)

                        'Dim FormId As SAPbouiCOM.Form = CType(objMain.objApplication.Forms.GetForm("ITEMUP", objMain.objApplication.Forms.ActiveForm.TypeCount), SAPbouiCOM.Form)
                        If DocEntry <> "" Then
                            Me.StartBackgroundForLoading()
                            '   Me.Item_Creation(DocEntry)
                        End If
                        objForm.Items.Item("Item_13").Enabled = True
                        objForm.Freeze(False)

                    End If


                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And
                         (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Me.ItemMasterDetailsExcelUpload(objForm.UniqueID)
                    End If
                    If pVal.ItemUID = "Price" And pVal.BeforeAction = False And
                         (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                        Dim DocEntry As String = objForm.Items.Item("DocNum").Specific.Value
                        Dim STATUS As String = objForm.Items.Item("STAT").Specific.Value.Trim().ToUpper()

                        If DocEntry <> "" AndAlso (STATUS = "A" OrElse STATUS = "APPROVED") Then

                            objMain.objApplication.StatusBar.SetText("Please Wait Prices are updating...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            Me.StartBackgroundForLoading1()
                        End If
                    End If
                    If pVal.ItemUID = "Download" And pVal.BeforeAction = False Then
                        '' Dim sPath As String

                        'sPath = IO.Directory.GetParent(System.Windows.Forms.Application.ExecutablePath).ToString
                        'sPath = sPath & "\AttendanceSummaryTemplate.xlsx"
                        Dim fullPath As String = objForm.Items.Item("ATTPATH").Specific.Value
                        Process.Start(fullPath)
                        Dim indexLoc As Integer = fullPath.LastIndexOf("\")
                        Dim path As String = fullPath.Substring(0, indexLoc)
                        Dim filename As String = fullPath.Substring(indexLoc + 1, fullPath.Length - indexLoc - 1)

                        Dim saveThread As New Threading.Thread(Sub()
                                                                   Dim saveFileDialog As New SaveFileDialog()
                                                                   saveFileDialog.InitialDirectory = "E:\\"
                                                                   saveFileDialog.FileName = filename
                                                                   saveFileDialog.Filter = "Excel Files|*.xlsx"

                                                                   If saveFileDialog.ShowDialog() = DialogResult.OK Then
                                                                       IO.File.Copy(fullPath, saveFileDialog.FileName, True)
                                                                   End If
                                                               End Sub)
                        saveThread.SetApartmentState(Threading.ApartmentState.STA)
                        saveThread.Start()

                    End If

                    'If pVal.ItemUID = "Item_11" And pVal.BeforeAction = False Then
                    '    Dim FileSaveWithPath As String = objForm.Items.Item("ATTPATH").Specific.Value

                    '    If FileSaveWithPath <> "" Then
                    '        Me.ItemMasterDetailsExcelUpload(FileSaveWithPath)
                    '    Else
                    '        objMain.objApplication.StatusBar.SetText("Please Select File Path...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        Exit Sub

                    '    End If

                    'End If
                    '          objForm.Freeze(False)

            End Select

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Function ValidateItemGroup1(ByVal objForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim oMatrix As SAPbouiCOM.Matrix =
            CType(objForm.Items.Item("ItmDetMtx").Specific, SAPbouiCOM.Matrix)

            Dim oRS As SAPbobsCOM.Recordset =
            objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            For i As Integer = 1 To oMatrix.RowCount

                Dim cmb As SAPbouiCOM.ComboBox =
                CType(oMatrix.Columns.Item("ITEMG").Cells.Item(i).Specific, SAPbouiCOM.ComboBox)

                If cmb.Selected IsNot Nothing Then

                    Dim selValue As String = cmb.Selected.Value

                    Dim qry As String =
                    "SELECT COUNT(*) FROM ""OITB"" WHERE ""ItmsGrpNam"" = '" & selValue.Replace("'", "''") & "'"

                    oRS.DoQuery(qry)

                    If oRS.Fields.Item(0).Value = 0 Then
                        objMain.objApplication.StatusBar.SetText("Invalid Item Group selected at row " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        Return False
                    End If

                Else
                    objMain.objApplication.StatusBar.SetText("Item Group is mandatory at row " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    Return False
                End If
            Next

            Return True

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            Return False
        End Try
    End Function
    Private Function ValidateItemGroup(ByVal objForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim oMatrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("ItmDetMtx").Specific, SAPbouiCOM.Matrix)

            Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            For i As Integer = 1 To oMatrix.RowCount

                Dim itemCodeCell As SAPbouiCOM.EditText = CType(oMatrix.Columns.Item("CARD").Cells.Item(i).Specific, SAPbouiCOM.EditText)

                If String.IsNullOrWhiteSpace(itemCodeCell.Value) Then
                    Continue For
                End If

                Dim cmb As SAPbouiCOM.ComboBox = CType(oMatrix.Columns.Item("ITEMG").Cells.Item(i).Specific, SAPbouiCOM.ComboBox)

                If cmb.Selected Is Nothing OrElse String.IsNullOrWhiteSpace(cmb.Selected.Value) Then
                    objMain.objApplication.StatusBar.SetText("Item Group is mandatory at row " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    Return False
                End If

                Dim selValue As String = cmb.Selected.Value

                Dim qry As String = "SELECT COUNT(*) FROM ""OITB"" WHERE ""ItmsGrpNam"" = '" & selValue.Replace("'", "''") & "'"

                oRS.DoQuery(qry)

                If oRS.Fields.Item(0).Value = 0 Then
                    objMain.objApplication.StatusBar.SetText("Invalid Item Group selected at row " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    Return False
                End If

            Next

            Return True

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            Return False
        End Try
    End Function


    Sub ClearCombo(col As SAPbouiCOM.Column, row As Integer)
        Try

            ' Remove all existing values
            While col.ValidValues.Count > 0
                col.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
            End While

            '    ' Add blank value
            '    col.ValidValues.Add("", "")

            '' Set blank as selected for current row
            'Dim oCombo As SAPbouiCOM.ComboBox = col.Cells.Item(row).Specific
            '    oCombo.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)

        Catch ex As Exception
            objForm.Freeze(False)
        End Try
    End Sub
    Sub defaultbarcode(ByVal NewItemCode As String, ByVal FItemCode As String, ByVal docentry As String)
        Try


            oItem = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
            Dim item As String = NewItemCode


            ' Retrieve the Item Master Data record to update
            If oItem.GetByKey(NewItemCode) Then ' Replace with the actual Item Code
                ' Set the Bar Code
                Dim BARCODE As String
                Dim barcodestest As String = "select T2.""UomEntry"", ""U_KEYC"",""U_UOMCDE"",""U_BarCode"",""U_BDesc""  from ""@TNX_ITEMUP"" T0 inner join ""@TNX_ITEMUP_C2"" T1 On T0.""DocEntry""= T1.""DocEntry"" inner join ""OUOM"" T2  on T2.""UomCode""=T1.""U_UOMCDE"" " &
                                       "where T0.""DocNum""='" & docentry & "'and ""U_KEYC"" ='" & FItemCode & "' and  ""U_PREFER"" ='Y' "
                Dim barcoders As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                barcoders.DoQuery(barcodestest)
                If barcoders.RecordCount > 0 Then
                    Dim barcodest As String = barcoders.Fields.Item("U_BarCode").Value
                    Dim barcodename As String = barcoders.Fields.Item("U_BDesc").Value
                    Dim barcodeuom As Integer = barcoders.Fields.Item("UomEntry").Value
                    oItem.BarCode = barcodest
                    'If barcodest <> "" Then
                    '    oItem.BarCode = barcodest
                    'Else
                    '    Dim barcodes As String = "select * from ""OITM"" where ""ItemCode""='" & NewItemCode & "'"
                    '    Dim barcoderrs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '    barcoderrs.DoQuery(barcodes)
                    'End If

                    ' Update the record
                    If oItem.Update() <> 0 Then
                        Dim lErrCode As Integer
                        Dim sErrMsg As String = objMain.objCompany.GetLastErrorDescription()

                        objMain.objCompany.GetLastError(lErrCode, sErrMsg)
                        objMain.objApplication.StatusBar.SetText(sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    Else
                        objMain.objApplication.StatusBar.SetText(" '" + NewItemCode + "' - Barcode updated successfully for  the Item '" + NewItemCode + "' ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


                    End If
                Else
                    objMain.objApplication.StatusBar.SetText("Prefered Barcode  not found.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
                '  objMain.objApplication.StatusBar.SetText("Item not found.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            End If
        Catch ex As Exception
            objForm.Items.Item("Item_13").Enabled = True
            objForm.Freeze(False)
        End Try
    End Sub
    Sub BussinesspartnerCatlogNumbercreation1(ByVal NewItemCode As String, ByVal FItemCode As String, ByVal docentry As String, ByVal CardCode1 As String, ByVal CatlogNumber As String)
        Try


            Dim oBP As SAPbobsCOM.BusinessPartners
            oBP = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            Dim item As String = NewItemCode

            Dim oAltCatNum As SAPbobsCOM.AlternateCatNum = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAlternateCatNum)
            oAltCatNum.ItemCode = NewItemCode
            oAltCatNum.CardCode = CardCode1

            oAltCatNum.Substitute = CatlogNumber
            Dim result As Integer = oAltCatNum.Add()
        Catch ex As Exception

        End Try
    End Sub
    ''Commented by ramya for barcode wise catalog creation
    Sub BussinesspartnerCatlogNumbercreation2(ByVal NewItemCode As String, ByVal FItemCode As String, ByVal docentry As String, ByVal CardCode1 As String, ByVal CatlogNumber As String)
        Try
            ' ✅ Validate inputs
            If String.IsNullOrEmpty(NewItemCode) OrElse String.IsNullOrEmpty(CardCode1) Then
                Exit Sub
            End If
            If CatlogNumber <> "" Then



                Dim oAltCatNum As SAPbobsCOM.AlternateCatNum = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAlternateCatNum)

                oAltCatNum.ItemCode = NewItemCode
                oAltCatNum.CardCode = CardCode1
                oAltCatNum.Substitute = CatlogNumber

                Dim result As Integer = oAltCatNum.Add()

                If result <> 0 Then
                    Dim errCode As Integer
                    Dim errMsg As String = ""
                    objMain.objCompany.GetLastError(errCode, errMsg)

                    objMain.objApplication.StatusBar.SetText(errMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    objMain.objApplication.StatusBar.SetText("Catalog number added successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
            End If
        Catch ex As Exception
            objForm.Items.Item("Item_13").Enabled = True
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub BussinesspartnerCatlogNumbercreation3(ByVal NewItemCode As String,
                                         ByVal FItemCode As String,
                                         ByVal docentry As String,
                                         ByVal CardCode1 As String)

        Try

            ' Validate
            If String.IsNullOrEmpty(NewItemCode) OrElse String.IsNullOrEmpty(CardCode1) Then
                Exit Sub
            End If

            Dim qStr As String = ""

            qStr = "SELECT DISTINCT ""U_CNum"" " &
               "FROM ""@TNX_ITEMUP_C2"" " &
               "WHERE ""DocEntry""='" & docentry & "' " &
               "AND ""U_KEYC""='" & FItemCode & "' " &
               "AND IFNULL(""U_CNum"",'') <> '' "

            Dim rs As SAPbobsCOM.Recordset
            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            rs.DoQuery(qStr)

            If rs.RecordCount > 0 Then

                While rs.EoF = False

                    Dim CatlogNumber As String = rs.Fields.Item("U_CNum").Value.ToString.Trim()

                    If CatlogNumber <> "" Then

                        Dim oAltCatNum As SAPbobsCOM.AlternateCatNum
                        oAltCatNum = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAlternateCatNum)

                        oAltCatNum.ItemCode = NewItemCode
                        oAltCatNum.CardCode = CardCode1
                        oAltCatNum.Substitute = CatlogNumber

                        Dim result As Integer = oAltCatNum.Add()

                        If result <> 0 Then

                            Dim errCode As Integer
                            Dim errMsg As String = ""

                            objMain.objCompany.GetLastError(errCode, errMsg)

                            objMain.objApplication.StatusBar.SetText(errMsg,
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        End If

                    End If

                    rs.MoveNext()

                End While

            End If

        Catch ex As Exception

            objForm.Items.Item("Item_13").Enabled = True
            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

    Sub BussinesspartnerCatlogNumbercreation(ByVal NewItemCode As String,
                                         ByVal FItemCode As String,
                                         ByVal docentry As String,
                                         ByVal CardCode1 As String, ByVal BarcodeNew As String, ByVal CATNUM As String, ByVal LineId As Integer)

        Try

            If String.IsNullOrEmpty(NewItemCode) Then Exit Sub

            Dim qStr As String = ""

            qStr = "SELECT ""LineId"",""U_BarCode"", ""U_CNum"" " &
               "FROM ""@TNX_ITEMUP_C2"" " &
               "WHERE ""DocEntry""='" & docentry & "' " &
               "AND ""U_KEYC""='" & FItemCode & "' " &
               "AND ""U_BarCode""='" & BarcodeNew & "' " &
               "AND ""U_CNum""='" & CATNUM & "' " &
               "AND ""LineId""='" & LineId & "' " &
               "AND IFNULL(""U_CNum"",'') <> '' " &
               "AND IFNULL(""U_BarCode"",'') <> '' "

            Dim rs As SAPbobsCOM.Recordset
            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            rs.DoQuery(qStr)

            If rs.RecordCount > 0 Then

                While rs.EoF = False

                    Dim Barcode As String = rs.Fields.Item("U_BarCode").Value.ToString.Trim()
                    Dim CatlogNumber As String = rs.Fields.Item("U_CNum").Value.ToString.Trim()
                    ''Dim Arabic As String = rs.Fields.Item("U_ADESC").Value.ToString.Trim()

                    If Barcode <> "" And CatlogNumber <> "" Then

                        '=========================================
                        ' ADD INTO OSCN
                        '=========================================
                        Dim oAltCatNum As SAPbobsCOM.AlternateCatNum
                        oAltCatNum = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAlternateCatNum)

                        oAltCatNum.ItemCode = NewItemCode
                        oAltCatNum.CardCode = CardCode1
                        oAltCatNum.Substitute = CatlogNumber

                        Dim result As Integer = oAltCatNum.Add()

                        If result <> 0 Then

                            Dim errCode As Integer
                            Dim errMsg As String = ""

                            objMain.objCompany.GetLastError(errCode, errMsg)

                        End If

                        '=========================================
                        ' UPDATE OBCD
                        '=========================================
                        Dim qUpdate As String = ""

                        qUpdate = "UPDATE ""OBCD"" " &
                              "SET ""U_CatalogNo""='" & CatlogNumber.Replace("'", "''") & "' " &
                              "WHERE ""BcdCode""='" & Barcode.Replace("'", "''") & "' " &
                              "AND ""ItemCode""='" & NewItemCode.Replace("'", "''") & "' "

                        Dim rsUpdate As SAPbobsCOM.Recordset
                        rsUpdate = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        rsUpdate.DoQuery(qUpdate)

                    End If

                    rs.MoveNext()

                End While

            End If

        Catch ex As Exception

            objForm.Items.Item("Item_13").Enabled = True
            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)

        Try
            objForm = objMain.objApplication.Forms.GetForm("ITEMUP",
                    objMain.objApplication.Forms.ActiveForm.TypeCount)

            Select Case BusinessObjectInfo.EventType

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                    If BusinessObjectInfo.BeforeAction = False AndAlso
                   BusinessObjectInfo.ActionSuccess = True Then


                        objForm.Freeze(True)

                        oDS = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP")

                        Dim status As String = ""
                        Dim STR As String = ""
                        Dim docEntry As String = ""
                        Dim User As String = objMain.objCompany.UserName

                        Try : status = objForm.Items.Item("STAT").Specific.Value : Catch : End Try
                        Try : STR = objForm.Items.Item("Item_19").Specific.Value : Catch : End Try
                        Try : status = If(status, "").Trim().ToUpper() : Catch : End Try
                        Try : docEntry = oDS.GetValue("DocEntry", 0).Trim() : Catch : End Try

                        status = status.Trim().ToUpper()

                        '====================================================
                        ' RESET ALL CONTROLS (VERY IMPORTANT)
                        '====================================================

                        ' Buttons OFF
                        Try : objForm.Items.Item("Submit").Enabled = False : Catch : End Try
                        Try : objForm.Items.Item("Item_11").Enabled = False : Catch : End Try 'Upload

                        Dim barcodestest As String = "Select T1.""U_AUTH"" from ""@SBO_AST"" T0 Inner Join ""@SBO_AST_C0"" T1 ON T0.""Code""=T1.""Code"" where T0.""Code""='PCR'"
                        Dim barcoders As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        barcoders.DoQuery(barcodestest)

                        Dim IsAuthorized As Boolean = False

                        While barcoders.EoF = False

                            If User.Trim().ToUpper() = barcoders.Fields.Item("U_AUTH").Value.ToString().Trim().ToUpper() Then
                                IsAuthorized = True
                                Exit While
                            End If

                            barcoders.MoveNext()

                        End While

                        If IsAuthorized = True Then
                            Try : objForm.Items.Item("Item_13").Enabled = True : Catch : End Try
                        Else
                            Try : objForm.Items.Item("Item_13").Enabled = False : Catch : End Try
                        End If
                        'Item Create
                        Try : objForm.Items.Item("Price").Enabled = False : Catch : End Try
                        Try : objForm.Items.Item("Download").Enabled = True : Catch : End Try

                        ' MATRICES RESET (ENABLE FIRST)
                        Try : objForm.Items.Item("ItmDetMtx").Enabled = True : Catch : End Try
                        Try : objForm.Items.Item("UOMMtx").Enabled = True : Catch : End Try
                        Try : objForm.Items.Item("BarMtx").Enabled = True : Catch : End Try
                        Try : objForm.Items.Item("Mtx5").Enabled = True : Catch : End Try
                        Try : objForm.Items.Item("Mtx6").Enabled = True : Catch : End Try
                        Try : objForm.Items.Item("Mtx7").Enabled = True : Catch : End Try


                        If (status = "OPEN" Or status = "O") And STR = "Yes" Then

                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

                            'ONLY SUBMIT ENABLED
                            Try : objForm.Items.Item("Submit").Enabled = True : Catch : End Try

                        Else
                            Try : objForm.Items.Item("Submit").Enabled = False : Catch : End Try

                        End If

                        '---------------- APPROVED ----------------
                        If status = "APPROVED" Or status = "A" Then

                            'ITEM CREATE ENABLED
                            Try : objForm.Items.Item("Item_13").Enabled = True : Catch : End Try
                            Try : objForm.Items.Item("Submit").Enabled = False : Catch : End Try
                            Try : objForm.Items.Item("ItmDetMtx").Enabled = False : Catch : End Try
                            Try : objForm.Items.Item("UOMMtx").Enabled = False : Catch : End Try
                            Try : objForm.Items.Item("BarMtx").Enabled = False : Catch : End Try
                            Try : objForm.Items.Item("Mtx5").Enabled = False : Catch : End Try
                            Try : objForm.Items.Item("Mtx6").Enabled = False : Catch : End Try
                            Try : objForm.Items.Item("Mtx7").Enabled = False : Catch : End Try

                        End If

                        '---------------- SUBMITTED ----------------
                        If status = "SUBMITTED" Or status = "S" Then

                            'ALL STAYS DISABLED (DEFAULT RESET ALREADY HANDLES)
                            Try : objForm.Items.Item("Submit").Enabled = False : Catch : End Try

                            Try : objForm.Items.Item("ItmDetMtx").Enabled = False : Catch : End Try
                            Try : objForm.Items.Item("UOMMtx").Enabled = False : Catch : End Try
                            Try : objForm.Items.Item("BarMtx").Enabled = False : Catch : End Try
                            Try : objForm.Items.Item("Mtx5").Enabled = False : Catch : End Try
                            Try : objForm.Items.Item("Mtx6").Enabled = False : Catch : End Try
                            Try : objForm.Items.Item("Mtx7").Enabled = False : Catch : End Try

                        End If
                        '---------------- REJECTED ----------------
                        If status = "REJECTED" Or status = "R" Then
                            Try : objForm.Items.Item("Submit").Enabled = True : Catch : End Try
                        End If

                        '====================================================
                        ' ALL ITEMS CREATED CHECK
                        '====================================================
                        If docEntry <> "" Then

                            Dim rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            rs.DoQuery("SELECT COUNT(*) CNT FROM ""@TNX_ITEMUP_C0"" " &
                           "WHERE ""DocEntry""='" & docEntry & "' " &
                           "AND IFNULL(""U_STATS"",'') <> 'Item Master is Succesfully Created' And ""U_KEYC""<>''")

                            If rs.Fields.Item("CNT").Value = 0 Then

                                'PRICE UPDATE ENABLED
                                Try : objForm.Items.Item("Price").Enabled = True : Catch : End Try

                                'UPLOAD + CREATE DISABLED
                                Try : objForm.Items.Item("Item_11").Enabled = False : Catch : End Try
                                Try : objForm.Items.Item("Item_13").Enabled = False : Catch : End Try

                                'ONLY DISABLE ITMCD COLUMN (NOT FULL MATRIX)
                                Try
                                    Dim oMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("ItmDetMtx").Specific
                                    oMatrix.Columns.Item("ITMCD").Editable = False
                                Catch
                                End Try
                                'MATRICES DISABLED
                                Try : objForm.Items.Item("ItmDetMtx").Enabled = False : Catch : End Try
                                Try : objForm.Items.Item("UOMMtx").Enabled = False : Catch : End Try
                                Try : objForm.Items.Item("BarMtx").Enabled = False : Catch : End Try
                                Try : objForm.Items.Item("Mtx5").Enabled = False : Catch : End Try
                                Try : objForm.Items.Item("Mtx6").Enabled = False : Catch : End Try
                                Try : objForm.Items.Item("Mtx7").Enabled = False : Catch : End Try

                            End If

                        End If


                        'Madhu Once prices are updated price button should be disblaed

                        Try
                                Dim ors1 As SAPbobsCOM.Recordset
                                ors1 = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                Dim rs1 As String

                                rs1 = "SELECT COUNT(*) AS ""CNT"" FROM ""@TNX_ITEMUP_C0"" " & "WHERE ""DocEntry"" = '" & docEntry & "' " & "AND IFNULL(TRIM(""U_PStst""),'') <> 'Price updated Successfully'"
                            objForm.Freeze(False)
                            ors1.DoQuery(rs1)
                                    If ors1.Fields.Item("CNT").Value = 0 Then
                                        objForm.Items.Item("Price").Enabled = False
                                    Else
                                        objForm.Items.Item("Price").Enabled = True
                                    End If

                                Catch ex As Exception
                                    objMain.objApplication.StatusBar.SetText(ex.Message)
                                    objForm.Freeze(False)
                                End Try

                    End If
            End Select

            ' End Select

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
            objForm.Freeze(False)
        End Try

    End Sub


    Private Sub SetViewModeBasedOnItemCode(ByVal objForm As SAPbouiCOM.Form)

        Try
            objForm.Freeze(True)
            Dim oMatrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("ItmDetMtx").Specific, SAPbouiCOM.Matrix)
            Dim allHaveItemCode As Boolean = True
            Dim hasDataRow As Boolean = False
            For i As Integer = 1 To oMatrix.VisualRowCount
                Dim itemCode As String = ""

                Try
                    itemCode = oMatrix.Columns.Item("ITMCD").Cells.Item(i).Specific.Value
                Catch
                End Try

                If i = oMatrix.VisualRowCount And String.IsNullOrWhiteSpace(itemCode) Then
                    Continue For
                End If

                If String.IsNullOrWhiteSpace(itemCode) Then
                    allHaveItemCode = False
                    Exit For
                End If

                hasDataRow = True
            Next

            If allHaveItemCode And hasDataRow Then

                objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE

                Try : objForm.Items.Item("1").Enabled = True : Catch : End Try
                Try : objForm.Items.Item("2").Enabled = True : Catch : End Try
                Try : objForm.Items.Item("Price").Enabled = True : Catch : End Try
                Try : objForm.EnableMenu("1281", True) : Catch : End Try
                Try : objForm.EnableMenu("1282", True) : Catch : End Try

            End If

        Catch ex As Exception
        Finally
            objForm.Freeze(False)
        End Try

    End Sub


    Sub CFLFilter(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Select Case CFL_ID
                Case "CFL_VEND"
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
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    Sub SetDefault(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            objMatrix = objForm.Items.Item("ItmDetMtx").Specific
            objMatrix1 = objForm.Items.Item("UOMMtx").Specific
            objMatrix2 = objForm.Items.Item("BarMtx").Specific
            objMatrix3 = objForm.Items.Item("Mtx5").Specific
            oDS = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP")
            oDS1 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C0")
            oDS2 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C1")
            oDS3 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C2")
            oDS4 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C3")
            Dim rs1 As String = ("SELECT IFNULL(MAX(""DocNum"" + 0), 0) + 1 AS ""DocNum"" FROM ""@TNX_ITEMUP"";")
            Dim ors1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ors1.DoQuery(rs1)
            If ors1.RecordCount > 0 Then

                oDS.SetValue("DocNum", 0, ors1.Fields.Item("DocNum").Value)
            Else
                oDS.SetValue("DocNum", 0, 1)
            End If
            oDS.SetValue("U_DDATE", 0, DateTime.Now.ToString("yyyyMMdd"))
            oDS.SetValue("U_STAT", 0, "O")
            Me.AutoDocentryNumber()
            objMatrix.Clear()
            oDS1.Clear()
            objMatrix.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "ItmDetMtx")
            objMatrix.AutoResizeColumns()

            objMatrix1.Clear()
            oDS2.Clear()
            objMatrix1.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "UOMMtx")
            objMatrix1.AutoResizeColumns()

            objMatrix2.Clear()
            oDS3.Clear()
            objMatrix2.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "BarMtx")
            objMatrix2.AutoResizeColumns()

            objMatrix3.Clear()
            oDS4.Clear()
            objMatrix3.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "Mtx5")
            objMatrix3.AutoResizeColumns()

            objMatrix4.Clear()
            oDS5.Clear()
            objMatrix4.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "Mtx6")
            objMatrix4.AutoResizeColumns()

            objMatrix5.Clear()
            oDS6.Clear()
            objMatrix5.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "Mtx7")
            objMatrix5.AutoResizeColumns()
            Me.AutoDocentryNumber()

            objMatrix6.Clear()
            oDS7.Clear()
            objMatrix6.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "Item_7")
            objMatrix6.AutoResizeColumns()
            Me.AutoDocentryNumber()
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal MatrixUID As String)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oDS = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP")
        oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C0")
        oDS2 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C1")
        oDS3 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C2")
        oDS4 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C3")
        oDS5 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C4")
        oDS6 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C5")
        oDS7 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C6")

        objMatrix = objForm.Items.Item("ItmDetMtx").Specific
        objMatrix1 = objForm.Items.Item("UOMMtx").Specific
        objMatrix2 = objForm.Items.Item("BarMtx").Specific
        objMatrix3 = objForm.Items.Item("Mtx5").Specific
        objMatrix4 = objForm.Items.Item("Mtx6").Specific
        objMatrix5 = objForm.Items.Item("Mtx7").Specific
        objMatrix6 = objForm.Items.Item("Item_7").Specific



        Select Case MatrixUID
            Case "ItmDetMtx"
                objMatrix.AddRow()
                oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
                oDS1.SetValue("U_KEYC", oDS1.Offset, "")
                oDS1.SetValue("U_ITEMN", oDS1.Offset, "")
                oDS1.SetValue("U_CARD", oDS1.Offset, "")
                oDS1.SetValue("U_FORIE", oDS1.Offset, "")
                oDS1.SetValue("U_ITEMG", oDS1.Offset, "")
                oDS1.SetValue("U_SUBGR", oDS1.Offset, "")
                oDS1.SetValue("U_SSGP", oDS1.Offset, "")
                oDS1.SetValue("U_SUGRP", oDS1.Offset, "")
                oDS1.SetValue("U_BRAND", oDS1.Offset, "")
                oDS1.SetValue("U_CATGO", oDS1.Offset, "")
                oDS1.SetValue("U_SCALE", oDS1.Offset, "")
                oDS1.SetValue("U_COUNT", oDS1.Offset, "")
                oDS1.SetValue("U_BUNIT", oDS1.Offset, "")
                oDS1.SetValue("U_ITEMS", oDS1.Offset, "")
                oDS1.SetValue("U_INVENT", oDS1.Offset, "")
                oDS1.SetValue("U_SALITM", oDS1.Offset, "")
                oDS1.SetValue("U_PURCH", oDS1.Offset, "")
                oDS1.SetValue("U_STATS", oDS1.Offset, "")
                oDS1.SetValue("U_ITMCD", oDS1.Offset, "")

                oDS1.SetValue("U_SugarTax", oDS1.Offset, "")
                oDS1.SetValue("U_SuTaxAmt", oDS1.Offset, "")

                oDS1.SetValue("U_VIC", oDS1.Offset, "")
                oDS1.SetValue("U_MARKUP", oDS1.Offset, "")
                oDS1.SetValue("U_MARGIN", oDS1.Offset, "")

                objMatrix.SetLineData(objMatrix.VisualRowCount)

            Case "UOMMtx"
                objMatrix1.AddRow()
                oDS2.SetValue("LineId", oDS2.Offset, objMatrix1.VisualRowCount)
                oDS2.SetValue("U_KEYCDE", oDS2.Offset, "")
                oDS2.SetValue("U_UOMCDE", oDS2.Offset, "")
                oDS2.SetValue("U_UOMUNT", oDS2.Offset, "")
                objMatrix1.SetLineData(objMatrix1.VisualRowCount)

            Case "BarMtx"
                objMatrix2.AddRow()
                oDS3.SetValue("LineId", oDS3.Offset, objMatrix2.VisualRowCount)
                oDS3.SetValue("U_KEYC", oDS3.Offset, "")
                oDS3.SetValue("U_ITEMN", oDS3.Offset, "")
                oDS3.SetValue("U_UOMCDE", oDS3.Offset, "")
                oDS3.SetValue("U_BarCode", oDS3.Offset, "")
                oDS3.SetValue("U_BDesc", oDS3.Offset, "")
                oDS3.SetValue("U_PREFER", oDS3.Offset, "")
                oDS3.SetValue("U_ADESC", oDS3.Offset, "")
                '
                oDS3.SetValue("U_STAX", oDS3.Offset, "")
                oDS3.SetValue("U_SUTAXAMT", oDS3.Offset, 0)
                oDS3.SetValue("U_ETAX", oDS3.Offset, 0)
                oDS3.SetValue("U_Price", oDS3.Offset, 0)
                oDS3.SetValue("U_Margin", oDS3.Offset, 0)
                oDS3.SetValue("U_Markup", oDS3.Offset, 0)
                objMatrix2.SetLineData(objMatrix2.VisualRowCount)
            Case "Mtx5"
                objMatrix3.AddRow()
                oDS4.SetValue("LineId", oDS4.Offset, objMatrix3.VisualRowCount)
                oDS4.SetValue("U_ItemC", oDS4.Offset, "")
                oDS4.SetValue("U_BrnchId", oDS4.Offset, "")
                oDS4.SetValue("U_BrachNm", oDS4.Offset, "")
                oDS4.SetValue("U_PriceId", oDS4.Offset, "")
                oDS4.SetValue("U_PriceLt", oDS4.Offset, "")
                oDS4.SetValue("U_UomEtry", oDS4.Offset, "")
                oDS4.SetValue("U_UOMCode", oDS4.Offset, "")
                oDS4.SetValue("U_OldPrice", oDS4.Offset, "")
                oDS4.SetValue("U_NewPrice", oDS4.Offset, "")
                oDS4.SetValue("U_Cost", oDS4.Offset, "")
                oDS4.SetValue("U_PriceList", oDS4.Offset, "")
                oDS4.SetValue("U_Update", oDS4.Offset, "20211101")
                oDS4.SetValue("U_EndDate", oDS4.Offset, "20401231")
                oDS4.SetValue("U_UserName", oDS4.Offset, "")
                objMatrix3.SetLineData(objMatrix3.VisualRowCount)
            Case "Mtx6"
                objMatrix4.AddRow()
                oDS5.SetValue("LineId", oDS5.Offset, objMatrix4.VisualRowCount)
                oDS5.SetValue("U_Code", oDS5.Offset, "")
                oDS5.SetValue("U_CCode", oDS5.Offset, "")
                oDS5.SetValue("U_UOMCode", oDS5.Offset, "")
                oDS5.SetValue("U_Price", oDS5.Offset, "")
                oDS5.SetValue("U_PriceId", oDS5.Offset, "")
                oDS5.SetValue("U_PriceLt", oDS5.Offset, "")
                oDS5.SetValue("U_UEntry", oDS5.Offset, "")
                oDS5.SetValue("U_Price", oDS5.Offset, "")
                objMatrix4.SetLineData(objMatrix4.VisualRowCount)
            Case "Mtx7"
                objMatrix5.AddRow()
                oDS6.SetValue("LineId", oDS6.Offset, objMatrix5.VisualRowCount)
                oDS6.SetValue("U_Code", oDS6.Offset, "")
                oDS6.SetValue("U_BrnchId", oDS6.Offset, "")
                oDS6.SetValue("U_BrachNm", oDS6.Offset, "")
                oDS6.SetValue("U_PriceId", oDS6.Offset, "")
                oDS6.SetValue("U_PriceLt", oDS6.Offset, "")
                oDS6.SetValue("U_UomEtry", oDS6.Offset, "")
                oDS6.SetValue("U_UOMCode", oDS6.Offset, "")
                oDS6.SetValue("U_Price", oDS6.Offset, "")
                objMatrix5.SetLineData(objMatrix5.VisualRowCount)

            Case "Item_7"
                objMatrix6.AddRow()
                oDS7.SetValue("LineId", oDS7.Offset, objMatrix6.VisualRowCount)
                oDS7.SetValue("U_BRANCH", oDS7.Offset, "")
                oDS7.SetValue("U_Warehouse", oDS7.Offset, "")
                oDS7.SetValue("U_Check", oDS7.Offset, "")
                oDS7.SetValue("U_BPLID", oDS7.Offset, "")
                oDS7.SetValue("U_WHSCODE", oDS7.Offset, "")
                objMatrix6.SetLineData(objMatrix6.VisualRowCount)

        End Select

    End Sub

    Private Sub LoadMatrix_BPLID()

        ' --- 1. Get recordset
        Dim oRS As SAPbobsCOM.Recordset
        oRS = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        ' --- 2. Get DBDataSource and Matrix
        oDS7 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C6")
        objMatrix6 = objForm.Items.Item("Item_7").Specific

        ' --- 3. SQL Query
        Dim sql As String
        sql = "SELECT T1.""BPLId"", T1.""BPLName"", T0.""WhsName"", T0.""WhsCode"" " &
          "FROM OWHS T0 INNER JOIN OBPL T1 ON T0.""BPLid"" = T1.""BPLId"" " &
          "WHERE IFNULL(T0.""Inactive"",'N') = 'N' AND IFNULL(T1.""Disabled"",'N') = 'N' " &
          "ORDER BY T1.""BPLId"" "
        sql = "SELECT DISTINCT  T1.""BPLId"", T1.""BPLName"" " &
          "FROM OWHS T0 INNER JOIN OBPL T1 ON T0.""BPLid"" = T1.""BPLId"" " &
          "WHERE IFNULL(T0.""Inactive"",'N') = 'N' AND IFNULL(T1.""Disabled"",'N') = 'N' " &
          "ORDER BY T1.""BPLId"" "

        oRS.DoQuery(sql)

        ' --- 4. Clear Matrix and DataSource
        objMatrix6.Clear()
        oDS7.Clear()
        objMatrix6.FlushToDataSource() ' ensures DBDataSource sync

        ' --- 5. Loop through recordset using For
        Dim rowCount As Integer = oRS.RecordCount
        For i As Integer = 1 To oRS.RecordCount

            ' Add row in matrix
            objMatrix6.AddRow()
            Dim rowIndex As Integer = objMatrix6.VisualRowCount

            ' Get fields into variables
            Dim varBPLId As String = If(oRS.Fields.Item("BPLId").Value Is Nothing, "", oRS.Fields.Item("BPLId").Value)
            Dim varBPLName As String = If(oRS.Fields.Item("BPLName").Value Is Nothing, "", oRS.Fields.Item("BPLName").Value)
            'Dim varWhsName As String = If(oRS.Fields.Item("WhsName").Value Is Nothing, "", oRS.Fields.Item("WhsName").Value)
            'Dim varWhsCode As String = If(oRS.Fields.Item("WhsCode").Value Is Nothing, "", oRS.Fields.Item("WhsCode").Value)
            Dim varCheck As String = "N"

            ' --- 6. Set DBDataSource offset
            objMatrix6.GetLineData(i)

            ' Populate DBDataSource from variables
            oDS7.SetValue("LineId", oDS7.Offset, rowIndex)
            oDS7.SetValue("U_BPLID", oDS7.Offset, varBPLId)
            oDS7.SetValue("U_BRANCH", oDS7.Offset, varBPLName)
            'oDS7.SetValue("U_WHSCODE", oDS7.Offset, varWhsCode)
            'oDS7.SetValue("U_WAREHOUSE", oDS7.Offset, varWhsName)
            oDS7.SetValue("U_CHECK", oDS7.Offset, varCheck)

            ' Refresh matrix row
            objMatrix6.SetLineData(i)

            ' Move to next record in recordset
            oRS.MoveNext()
        Next

        ' --- 7. Optional: Add an empty row for user input
        'AddEmptyRowToMatrix()

    End Sub
    ' --- Helper method: Add one empty/default row
    Private Sub AddEmptyRowToMatrix()
        objMatrix6.AddRow()
        Dim rowIndex As Integer = objMatrix6.VisualRowCount

        oDS7.SetValue("LineId", oDS7.Offset, rowIndex)
        oDS7.SetValue("U_BPLID", oDS7.Offset, "")
        oDS7.SetValue("U_BRANCH", oDS7.Offset, "")
        oDS7.SetValue("U_WHSCODE", oDS7.Offset, "")
        oDS7.SetValue("U_WAREHOUSE", oDS7.Offset, "")
        oDS7.SetValue("U_CHECK", oDS7.Offset, "N")

        objMatrix6.SetLineData(rowIndex)
    End Sub
    Sub SetNewLineNewBranch(ByVal FormUID As String, ByVal Mtrixid As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDS = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP")
            oDS4 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C3")
            objMatrix = objForm.Items.Item("Mtx5").Specific

            objMatrix.AddRow()
            oDS4.SetValue("LineId", oDS4.Offset, objMatrix.VisualRowCount)
            oDS4.SetValue("U_ItemC", oDS4.Offset, "")
            oDS4.SetValue("U_BrnchId", oDS4.Offset, "")
            oDS4.SetValue("U_BrachNm", oDS4.Offset, "")
            oDS4.SetValue("U_PriceId", oDS4.Offset, "")
            oDS4.SetValue("U_PriceLt", oDS4.Offset, "")
            oDS4.SetValue("U_UomEtry", oDS4.Offset, "")
            oDS4.SetValue("U_UOMCode", oDS4.Offset, "")
            oDS4.SetValue("U_OldPrice", oDS4.Offset, "")
            oDS4.SetValue("U_NewPrice", oDS4.Offset, "")
            oDS4.SetValue("U_Update", oDS4.Offset, "20211101")
            oDS4.SetValue("U_EndDate", oDS4.Offset, "20401231")
            oDS4.SetValue("U_UserName", oDS4.Offset, "")

            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Function Validate(ByVal FormUID As String, ByVal MatrixUID As String) As Boolean
        Try

            'If oDS.GetValue("U_STYLNO", 0) = "" Then
            '    objMain.objApplication.StatusBar.SetText("Please Choose Style No...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            '    Exit Function
            'End If
            'objMatrix = objForm.Items.Item("Mtx").Specific
            'objMatrix1 = objForm.Items.Item("SMtx").Specific
            objMatrix = objForm.Items.Item("ItmDetMtx").Specific

            If objMatrix.VisualRowCount = 0 Then
                objMain.objApplication.StatusBar.SetText("Matrix cannot be Empty")
                Return False
            ElseIf objMatrix.VisualRowCount = 1 Then
                If (objMatrix.Columns.Item("KEYC").Cells.Item(1).Specific.Value.Trim = "") Then
                    objMain.objApplication.StatusBar.SetText("Matrix cannot be Empty")
                    Return False
                End If
            End If

            Select Case MatrixUID
            End Select
            Return True

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try
    End Function


    'Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
    '    Dim objForm As SAPbouiCOM.Form
    '    Dim oMenuItem As SAPbouiCOM.MenuItem
    '    Dim oMenus As SAPbouiCOM.Menus
    '    Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
    '    oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
    '    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
    '    objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)
    '    Try
    '        If eventInfo.FormUID = objForm.UniqueID Then
    '            If (eventInfo.BeforeAction = True) Then
    '                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
    '                    objForm.EnableMenu("1292", True)
    '                    objForm.EnableMenu("774", True)


    '                End If
    '            End If
    '        End If
    '        ' System.Diagnostics.Process.Start()
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub
    'Sub priceupdate(ByVal FormUID As String, ByVal ItemCode As String, ByVal ItemName As String)

    '    Try

    '        objForm = objMain.objApplication.Forms.Item(FormUID)

    '        Dim oGeneralService As SAPbobsCOM.GeneralService
    '        Dim oGeneralData As SAPbobsCOM.GeneralData
    '        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
    '        Dim oCompanyService As SAPbobsCOM.CompanyService
    '        Dim oChild As SAPbobsCOM.GeneralData
    '        Dim oChildren As SAPbobsCOM.GeneralDataCollection
    '        oCompanyService = objMain.objCompany.GetCompanyService()
    '        oGeneralService = oCompanyService.GetGeneralService("ME_OPRICEUPD")
    '        oGeneralData = ((oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)))
    '        Dim GetDflt As String = "Select ""DfltSeries"" From ONNM Where ""ObjectCode"" = 'ME_OPRICEUPD'"
    '        Dim oRsGetDflt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oRsGetDflt.DoQuery(GetDflt)
    '        Dim str As String = oRsGetDflt.Fields.Item(0).Value
    '        oGeneralData.SetProperty("Series", "490")
    '        oGeneralData.SetProperty("U_SERIES", str)
    '        oGeneralData.SetProperty("U_ITMCD", ItemCode)
    '        oGeneralData.SetProperty("U_ITMNM", ItemName)
    '        oGeneralData.SetProperty("U_USRID", objMain.objCompany.UserName)

    '        'oGeneralData.SetProperty("U_UNITS", oRsFutPrice.Fields.Item("U_UNITS").Value)
    '        'oGeneralData.SetProperty("U_USRID", oRsFutPrice.Fields.Item("U_USRID").Value)
    '        'oGeneralData.SetProperty("U_NEWPRC", Convert.ToDouble(oRsFutPrice.Fields.Item("U_NEWPRC").Value))
    '        oChildren = oGeneralData.Child("ME_PRICEUPD_C0")

    '        Dim GetBranchList As String = " Select ""ItemCode"" , ""BranchId"" ,""BranchName"" , ""PricListId"" , ""PriceList"" , ""UomEntry"" , ""UomCode"" , Sum (""OldPrice"") As ""OldPrice""  " &
    '                                     " From (Select AA.""ItemCode"" , T1.""BPLId"" As ""BranchId"" , T1.""BPLName"" As ""BranchName"" ,  T1.""U_PricelistNo""  As  ""PricListId"" , " &
    '                                     " (Select ""ListName"" From OPLN  Where ""ListNum"" =  T1.""U_PricelistNo"" ) As ""PriceList"" , AA.""UomEntry"" ,  AA.""UomCode""  ,  " &
    '                                     " 0 As ""OldPrice""  From (Select T0.""ItemCode"" , C.""UomCode"" , C.""UomEntry"" , ROW_NUMBER() OVER (PARTITION BY B.""UomEntry"" ORDER BY B.""UomEntry"" DESC) AS UGP1RN  " &
    '                                     " From OITM T0 Inner Join OUGP A On A.""UgpEntry"" = T0.""UgpEntry""  Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry""  Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry""  Where T0.""ItemCode"" = '" & ItemCode & "' AND B.""IsActive""='Y'  ) AA  " &
    '                                     " Inner Join  (Select   ROW_NUMBER() OVER (PARTITION BY ""BPLId"" ORDER BY ""BPLId"" DESC) AS BPLRN , * From OBPL  Where ""U_PricelistNo"" <> '') T1 On AA.""UGP1RN"" = ""BPLRN""  " &
    '                                     "            Union All " &
    '                                     " Select T0.""ItemCode"" ,  T2.""BPLId"" As ""BranchId"" ,  T2.""BPLName"" As ""BranchName"" , T2.""U_PricelistNo"" As  ""PricListId"" ," &
    '                                     " (Select ""ListName"" From OPLN  Where ""ListNum"" =  T2.""U_PricelistNo"") As ""PriceList"" ,   B.""UomEntry"" , C.""UomCode"" , " &
    '                                     " T1.""Price""  As ""OldPrice""   From OITM T0 Inner Join OUGP A On A.""UgpEntry"" = T0.""UgpEntry""  Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry"" " &
    '                                     " Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry""  Inner Join ITM1 T1 On T1.""ItemCode"" = T0.""ItemCode""   And  T1.""UomEntry"" = T0.""IUoMEntry""   " &
    '                                     " Inner Join OBPL T2 On  T2.""U_PricelistNo"" = T1.""PriceList""  Where T0.""ItemCode"" = '" & ItemCode & "'    And  B.""UomEntry"" =  T0.""IUoMEntry""  AND B.""IsActive""='Y'  " &
    '                                     "            Union All " &
    '                                     " Select T0.""ItemCode"" , T2.""BPLId"" As ""BranchId"" ,  T2.""BPLName"" As ""BranchName"" , T2.""U_PricelistNo"" As  ""PricListId"" , " &
    '                                     " (Select ""ListName"" From OPLN  Where ""ListNum"" =  T2.""U_PricelistNo"") As ""PriceList"" ,   B.""UomEntry"" , C.""UomCode"" , " &
    '                                     "  T3.""Price""  As ""OldPrice""   From OITM T0 Inner Join OUGP A On A.""UgpEntry"" = T0.""UgpEntry""  Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry"" " &
    '                                     " Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry""    Inner Join ITM9 T3 On T3.""ItemCode"" = T0.""ItemCode"" And  T3.""UomEntry"" = B.""UomEntry"" " &
    '                                     " Inner Join OBPL T2 On  T2.""U_PricelistNo"" = T3.""PriceList""       Where T0.""ItemCode"" = '" & ItemCode & "' AND B.""IsActive""='Y') ""CompleteQty""   " &
    '                                     " Group By ""ItemCode"" , ""BranchId"" ,""BranchName"" , ""PricListId"" , ""PriceList"" ,  ""UomEntry"" , ""UomCode"" "
    '        Dim oRsGetBranchList As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oRsGetBranchList.DoQuery(GetBranchList)
    '        If oRsGetBranchList.RecordCount > 0 Then
    '            For i As Integer = 0 To oRsGetBranchList.RecordCount - 1
    '                oChild = oChildren.Add
    '                oChild.SetProperty("U_ITMCD", oRsGetBranchList.Fields.Item("ItemCode").Value)
    '                Dim Br As String = oRsGetBranchList.Fields.Item("BranchId").Value
    '                Dim BrName As String = oRsGetBranchList.Fields.Item("BranchName").Value
    '                oChild.SetProperty("U_BRCHID", Br)
    '                oChild.SetProperty("U_BRCHNM", BrName)
    '                oChild.SetProperty("U_PRCID", oRsGetBranchList.Fields.Item("PricListId").Value)
    '                oChild.SetProperty("U_PRCLST", oRsGetBranchList.Fields.Item("PriceList").Value)
    '                Dim Uome As String = oRsGetBranchList.Fields.Item("UomEntry").Value
    '                Dim uomc As String = oRsGetBranchList.Fields.Item("UomCode").Value
    '                oChild.SetProperty("U_UOMENT", Uome)
    '                oChild.SetProperty("U_UOMCD", uomc)
    '                oChild.SetProperty("U_OLDPRC", oRsGetBranchList.Fields.Item("OldPrice").Value)
    '                oChild.SetProperty("U_NEWPRC", 0)
    '                Dim CheckUpDate As String = " Select ""U_UPDT"" , ""U_ENDDT"" , ""U_USRCD"" From ""@ME_PRICEUPD_C0"" A Where A.""U_ITMCD"" = '" & oRsGetBranchList.Fields.Item("ItemCode").Value & "' " &
    '                                            " And ""U_BRCHID"" = '" & oRsGetBranchList.Fields.Item("BranchId").Value & "' And " &
    '                                            " ""U_PRCID""='" & oRsGetBranchList.Fields.Item("PricListId").Value & "' And ""U_UOMCD"" = '" & oRsGetBranchList.Fields.Item("UomCode").Value & "' "
    '                Dim oRsCheckUpDate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                oRsCheckUpDate.DoQuery(CheckUpDate)
    '                If CDate(oRsCheckUpDate.Fields.Item("U_UPDT").Value).ToString("yyyyMMdd") <> "18991230" And CDate(oRsCheckUpDate.Fields.Item("U_UPDT").Value).ToString("yyyyMMdd") <> "19991230" And oRsCheckUpDate.Fields.Item("U_UPDT").Value.ToString <> "" Then
    '                    oChild.SetProperty("U_UPDT", CDate(oRsCheckUpDate.Fields.Item("U_UPDT").Value).ToString("yyyyMMdd"))
    '                Else
    '                    oChild.SetProperty("U_UPDT", "20211101") 'DateTime.Now.ToString("yyyyMMdd"))
    '                End If
    '                If CDate(oRsCheckUpDate.Fields.Item("U_ENDDT").Value).ToString("yyyyMMdd") <> "18991230" And CDate(oRsCheckUpDate.Fields.Item("U_ENDDT").Value).ToString("yyyyMMdd") <> "19991230" And oRsCheckUpDate.Fields.Item("U_ENDDT").Value.ToString <> "" Then
    '                    oChild.SetProperty("U_ENDDT", CDate(oRsCheckUpDate.Fields.Item("U_ENDDT").Value).ToString("yyyyMMdd"))
    '                Else
    '                    oChild.SetProperty("U_ENDDT", "20401231")
    '                End If
    '                If oRsCheckUpDate.Fields.Item("U_USRCD").Value <> "" Then
    '                    oChild.SetProperty("U_USRCD", oRsCheckUpDate.Fields.Item("U_USRCD").Value.ToString)
    '                Else
    '                    oChild.SetProperty("U_USRCD", "")
    '                End If

    '                oRsGetBranchList.MoveNext()
    '            Next

    '            oGeneralParams = oGeneralService.Add(oGeneralData)
    '        End If
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub
    Sub UpdatePrice(ByVal DocEntry As String)
        Dim UomCount As Integer = 0
        Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim Query3 As String = "SELECT * FROM ""@TNX_ITEMUP"" T0 INNER JOIN ""@TNX_ITEMUP_C0""  T1 ON T0.""DocEntry""=T1.""DocEntry"" where ""DocNum""='" & DocEntry & "'and ""U_ITMCD""<>''"

        rs.DoQuery(Query3)

        If rs.RecordCount > 0 Then
            Dim oItem As SAPbobsCOM.Items = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
            For i As Integer = 1 To rs.RecordCount
                Dim FItemCode As String = rs.Fields.Item("U_ITMCD").Value
                Dim FKeyCode As String = rs.Fields.Item("U_KEYC").Value
                If oItem.GetByKey(FItemCode) Then
                    Dim GetItemPriceLineDtls As String = "select distinct ""U_UomEtry"" As ""UomEntry"",T3.""IUoMEntry"" As ""DefaultUOM"",""U_PriceId"" As ""PricListId"",""U_OldPrice"" As ""New Price"" from ""@TNX_ITEMUP"" T0 inner join ""@TNX_ITEMUP_C3"" T1 On T0.""DocEntry""= T1.""DocEntry""  inner join ""@TNX_ITEMUP_C0"" T2 On T0.""DocEntry""= T2.""DocEntry"" Inner Join OITM T3 ON T3.""ItemCode"" = T2.""U_ITMCD"" " &
                                    "where T0.""DocNum""='" & DocEntry & "'and ""U_ItemC"" ='" & FKeyCode & "'and ""U_OldPrice"">0 and ""U_UomEtry""<>'' and ""U_PriceId""<>'' "
                    Dim oRsGetItemPriceLineDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetItemPriceLineDtls.DoQuery(GetItemPriceLineDtls)

                    If oRsGetItemPriceLineDtls.RecordCount > 0 Then
                        oRsGetItemPriceLineDtls.MoveFirst()
                        For j As Integer = 1 To oRsGetItemPriceLineDtls.RecordCount
                            Dim UomEntry As Integer = CInt(oRsGetItemPriceLineDtls.Fields.Item("UomEntry").Value)
                            Dim DefaultUom As Integer = CInt(oRsGetItemPriceLineDtls.Fields.Item("DefaultUOM").Value)
                            Dim PriceListRow As Integer = oRsGetItemPriceLineDtls.Fields.Item("PricListId").Value
                            oItem.PriceList.SetCurrentLine(PriceListRow - 1)
                            Dim strPrlist As String = oItem.PriceList.PriceList
                            If strPrlist = oRsGetItemPriceLineDtls.Fields.Item("PricListId").Value Then
                                If DefaultUom = UomEntry Then
                                    oItem.PriceList.Price = CDbl(oRsGetItemPriceLineDtls.Fields.Item("New Price").Value)

                                Else
                                    Dim UOMFoundFlag As String = ""
                                    Dim chkIfExsist As String = "Select Count(*) From ITM9 Where ""ItemCode"" = '" & FItemCode & "' And ""UomEntry"" = '" & UomEntry & "'"
                                    Dim oRschkIfExsist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRschkIfExsist.DoQuery(chkIfExsist)
                                    If oRschkIfExsist.Fields.Item(0).Value > 0 Then
                                        For k As Integer = 0 To oItem.PriceList.UoMPrices.Count - 1
                                            oItem.PriceList.UoMPrices.SetCurrentLine(k)
                                            Dim iUom As Integer = oItem.PriceList.UoMPrices.UoMEntry
                                            If iUom = UomEntry Then
                                                oItem.PriceList.UoMPrices.Price = CDbl(oRsGetItemPriceLineDtls.Fields.Item("New Price").Value)

                                                UOMFoundFlag = "Yes"
                                            End If
                                        Next
                                    End If
                                    If UOMFoundFlag = "" Then
                                        Dim GetUomPriceListCount As String = "Select Count(*) From ITM9 Where ""ItemCode"" = '" & FItemCode & "' And ""PriceList"" = '" & strPrlist & "'"
                                        Dim oRsGetUomPriceListCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRsGetUomPriceListCount.DoQuery(GetUomPriceListCount)
                                        If oRsGetUomPriceListCount.Fields.Item(0).Value = 0 Then
                                            UomCount = 0
                                        Else
                                            UomCount = oRsGetUomPriceListCount.Fields.Item(0).Value
                                        End If
                                        oItem.PriceList.UoMPrices.Add()
                                        oItem.PriceList.UoMPrices.SetCurrentLine(UomCount)
                                        oItem.PriceList.UoMPrices.Price = CDbl(oRsGetItemPriceLineDtls.Fields.Item("New Price").Value)
                                        oItem.PriceList.UoMPrices.UoMEntry = UomEntry
                                        oItem.PriceList.UoMPrices.PriceList = strPrlist
                                        oItem.PriceList.UoMPrices.Auto = SAPbobsCOM.BoYesNoEnum.tNO
                                        oItem.PriceList.UoMPrices.Currency = "QR"
                                    End If
                                End If
                            End If
                            oRsGetItemPriceLineDtls.MoveNext()
                        Next
                    End If

                    Dim RetValue As Integer = oItem.Update
                    If RetValue <> 0 Then
                        Dim Err As String = objMain.objCompany.GetLastErrorDescription.Replace("'", "")
                        Dim qSt1 As String = "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_PStst""='Priceupdate Error:'" & objMain.objCompany.GetLastErrorDescription.Replace("'", "").ToString().Substring(0, 200) & "  WHERE  ""LineId"" = '" & i & "' and ""DocEntry""=(Select ""DocEntry"" from  ""@TNX_ITEMUP"" where ""DocNum""= '" & DocEntry & "')"
                        Dim rSe1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        rSe1.DoQuery(qSt1)
                    Else
                        Dim qSt As String = "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_PStst""='Priceupdated Succesfully' WHERE  ""LineId"" = '" & i & "' and ""DocEntry""=(Select ""DocEntry"" from  ""@TNX_ITEMUP"" where ""DocNum""= '" & DocEntry & "')"
                        Dim rSe As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        rSe.DoQuery(qSt)
                        objMain.objApplication.StatusBar.SetText("Price are updating succesfully for" & FItemCode & " ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)

                    End If

                End If
                rs.MoveNext()
            Next
        End If
    End Sub
    Sub UpdatePriceForMultipleBranchesbakup(ByVal DocEntry As String)
        Dim UomCount As Integer = 0
        Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim Query3 As String = "SELECT * FROM ""@TNX_ITEMUP"" T0 INNER JOIN ""@TNX_ITEMUP_C0""  T1 ON T0.""DocEntry""=T1.""DocEntry"" where ""DocNum""='" & DocEntry & "'and ""U_ITMCD""<>''"

        rs.DoQuery(Query3)

        If rs.RecordCount > 0 Then

            For i As Integer = 1 To rs.RecordCount
                Dim FItemCode As String = rs.Fields.Item("U_ITMCD").Value
                Dim FKeyCode As String = rs.Fields.Item("U_KEYC").Value
                'Dim oItem As SAPbobsCOM.Items = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                'If oItem.GetByKey(FItemCode) Then
                Dim GetItemPriceLineDtls As String = "select distinct ""U_UomEtry"" As ""UomEntry"",T3.""IUoMEntry"" As ""DefaultUOM"",""U_PriceId"" As ""PricListId"",""U_BrnchId""as ""BranchId"",""U_OldPrice"" As ""New Price"" from ""@TNX_ITEMUP"" T0 inner join ""@TNX_ITEMUP_C3"" T1 On T0.""DocEntry""= T1.""DocEntry""  inner join ""@TNX_ITEMUP_C0"" T2 On T0.""DocEntry""= T2.""DocEntry"" " &
                    "Inner Join OITM T3 ON T3.""ItemCode"" = T2.""U_ITMCD""" &
                                    "where T0.""DocNum""='" & DocEntry & "'and ""U_ItemC"" ='" & FKeyCode & "'and ""U_OldPrice"">0 and ""U_UomEtry""<>'' and ""U_PriceId""<>'' "

                Dim oRsGetItemPriceLineDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetItemPriceLineDtls.DoQuery(GetItemPriceLineDtls)

                If oRsGetItemPriceLineDtls.RecordCount > 0 Then
                    oRsGetItemPriceLineDtls.MoveFirst()
                    For j As Integer = 1 To oRsGetItemPriceLineDtls.RecordCount
                        Dim UomEntry As Integer = CInt(oRsGetItemPriceLineDtls.Fields.Item("UomEntry").Value)
                        Dim DefaultUom As Integer = CInt(oRsGetItemPriceLineDtls.Fields.Item("DefaultUOM").Value)
                        Dim BranchId As Integer = CInt(oRsGetItemPriceLineDtls.Fields.Item("BranchId").Value)
                        Dim PriceListRow As Integer = oRsGetItemPriceLineDtls.Fields.Item("PricListId").Value
                        Dim oItem As SAPbobsCOM.Items = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                        If oItem.GetByKey(FItemCode) Then
                            If BranchId = 0 Then
                                Dim oRsGetBranchDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim GetBranch As String = "Select ""U_PricelistNo""as ""PricListId"" from ""OBPL"" where   ""U_PricelistNo"" <>''"
                                oRsGetBranchDtls.DoQuery(GetBranch)

                                For r As Integer = 1 To oRsGetBranchDtls.RecordCount
                                    Dim PriceListRow1 As Integer = oRsGetBranchDtls.Fields.Item("PricListId").Value
                                    oItem.PriceList.SetCurrentLine(PriceListRow1 - 1)
                                    Dim strPrlist1 As String = oItem.PriceList.PriceList
                                    If strPrlist1 = oRsGetBranchDtls.Fields.Item("PricListId").Value Then
                                        If DefaultUom = UomEntry Then
                                            oItem.PriceList.Price = CDbl(oRsGetItemPriceLineDtls.Fields.Item("New Price").Value)

                                        Else
                                            Dim UOMFoundFlag As String = ""
                                            Dim chkIfExsist As String = "Select Count(*) From ITM9 Where ""ItemCode"" = '" & FItemCode & "' And ""UomEntry"" = '" & UomEntry & "'"
                                            Dim oRschkIfExsist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRschkIfExsist.DoQuery(chkIfExsist)
                                            If oRschkIfExsist.Fields.Item(0).Value > 0 Then
                                                For k As Integer = 0 To oItem.PriceList.UoMPrices.Count - 1
                                                    oItem.PriceList.UoMPrices.SetCurrentLine(k)
                                                    Dim iUom As Integer = oItem.PriceList.UoMPrices.UoMEntry
                                                    If iUom = UomEntry Then
                                                        oItem.PriceList.UoMPrices.Price = CDbl(oRsGetItemPriceLineDtls.Fields.Item("New Price").Value)

                                                        UOMFoundFlag = "Yes"
                                                    End If
                                                Next
                                            End If
                                            If UOMFoundFlag = "" Then
                                                Dim GetUomPriceListCount As String = "Select Count(*) From ITM9 Where ""ItemCode"" = '" & FItemCode & "' And ""PriceList"" = '" & strPrlist1 & "'"
                                                Dim oRsGetUomPriceListCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                oRsGetUomPriceListCount.DoQuery(GetUomPriceListCount)
                                                If oRsGetUomPriceListCount.Fields.Item(0).Value = 0 Then
                                                    UomCount = 0
                                                Else
                                                    UomCount = oRsGetUomPriceListCount.Fields.Item(0).Value
                                                End If
                                                oItem.PriceList.UoMPrices.Add()
                                                oItem.PriceList.UoMPrices.SetCurrentLine(UomCount)
                                                oItem.PriceList.UoMPrices.Price = CDbl(oRsGetItemPriceLineDtls.Fields.Item("New Price").Value)
                                                oItem.PriceList.UoMPrices.UoMEntry = UomEntry
                                                oItem.PriceList.UoMPrices.PriceList = strPrlist1
                                                oItem.PriceList.UoMPrices.Auto = SAPbobsCOM.BoYesNoEnum.tNO
                                                oItem.PriceList.UoMPrices.Currency = "QR"
                                            End If
                                        End If
                                    End If
                                    oRsGetBranchDtls.MoveNext()
                                Next
                            Else
                                'oItem.PriceList.SetCurrentLine(PriceListRow - 1)
                                oItem.PriceList.SetCurrentLine(PriceListRow)
                                Dim strPrlist As String = oItem.PriceList.PriceList
                                If strPrlist = oRsGetItemPriceLineDtls.Fields.Item("PricListId").Value Then
                                    If DefaultUom = UomEntry Then
                                        oItem.PriceList.Price = CDbl(oRsGetItemPriceLineDtls.Fields.Item("New Price").Value)

                                    Else
                                        Dim UOMFoundFlag As String = ""
                                        Dim chkIfExsist As String = "Select Count(*) From ITM9 Where ""ItemCode"" = '" & FItemCode & "' And ""UomEntry"" = '" & UomEntry & "'"
                                        Dim oRschkIfExsist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRschkIfExsist.DoQuery(chkIfExsist)
                                        If oRschkIfExsist.Fields.Item(0).Value > 0 Then
                                            For k As Integer = 0 To oItem.PriceList.UoMPrices.Count - 1
                                                oItem.PriceList.UoMPrices.SetCurrentLine(k)
                                                Dim iUom As Integer = oItem.PriceList.UoMPrices.UoMEntry
                                                If iUom = UomEntry Then
                                                    oItem.PriceList.UoMPrices.Price = CDbl(oRsGetItemPriceLineDtls.Fields.Item("New Price").Value)

                                                    UOMFoundFlag = "Yes"
                                                End If
                                            Next
                                        End If
                                        If UOMFoundFlag = "" Then
                                            Dim GetUomPriceListCount As String = "Select Count(*) From ITM9 Where ""ItemCode"" = '" & FItemCode & "' And ""PriceList"" = '" & strPrlist & "'"
                                            Dim oRsGetUomPriceListCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRsGetUomPriceListCount.DoQuery(GetUomPriceListCount)
                                            If oRsGetUomPriceListCount.Fields.Item(0).Value = 0 Then
                                                UomCount = 0
                                            Else
                                                UomCount = oRsGetUomPriceListCount.Fields.Item(0).Value
                                            End If
                                            oItem.PriceList.UoMPrices.Add()
                                            oItem.PriceList.UoMPrices.SetCurrentLine(UomCount)
                                            oItem.PriceList.UoMPrices.Price = CDbl(oRsGetItemPriceLineDtls.Fields.Item("New Price").Value)
                                            oItem.PriceList.UoMPrices.UoMEntry = UomEntry
                                            oItem.PriceList.UoMPrices.PriceList = strPrlist
                                            oItem.PriceList.UoMPrices.Auto = SAPbobsCOM.BoYesNoEnum.tNO
                                            oItem.PriceList.UoMPrices.Currency = "QR"
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        Dim RetValue As Integer = oItem.Update
                        If RetValue <> 0 Then
                            Dim ErrMsg As String = objMain.objCompany.GetLastErrorDescription.ToString
                            If ErrMsg.Length > 200 Then
                                ErrMsg = ErrMsg.Substring(0, 200)
                            End If
                            Dim qSt1 As String = "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_PStst""='" & ErrMsg & "'  WHERE  ""LineId"" = '" & i & "' and ""DocEntry""=(Select ""DocEntry"" from  ""@TNX_ITEMUP"" where ""DocNum""= '" & DocEntry & "')"
                            Dim rSe1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            rSe1.DoQuery(qSt1)
                        Else
                            Dim qSt As String = "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_PStst""='Priceupdated Succesfully' WHERE  ""LineId"" = '" & i & "' and ""DocEntry""=(Select ""DocEntry"" from  ""@TNX_ITEMUP"" where ""DocNum""= '" & DocEntry & "')"
                            Dim rSe As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            rSe.DoQuery(qSt)
                            objMain.objApplication.StatusBar.SetText("Price are updating succesfully for" & FItemCode & " ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)

                        End If
                        oRsGetItemPriceLineDtls.MoveNext()
                    Next
                End If
                rs.MoveNext()
            Next
        End If
    End Sub

    Sub UpdatePriceForMultipleBranches13(ByVal DocEntry As String)
        ' Initialize Recordsets
        Dim rsItems As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim rsUomDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            ' Get the DocEntry from the UI DataSource
            Dim DOCENTRYX As String = oDS.GetValue("DocEntry", oDS.Offset).Trim()

            ' 1. Get Distinct Items and Keys from the Header Lines
            Dim itemQuery As String = $"SELECT DISTINCT T1.""U_ITMCD"", T1.""U_KEYC"" FROM ""@TNX_ITEMUP_C0"" T1 " &
                                  $"WHERE T1.""DocEntry"" = '{DOCENTRYX}' AND T1.""U_ITMCD"" <> ''"
            rsItems.DoQuery(itemQuery)


            ' Outer Loop: Iterate through unique Items
            For iRow As Integer = 0 To rsItems.RecordCount - 1
                Dim sItemCode As String = rsItems.Fields.Item("U_ITMCD").Value.ToString().Trim()
                Dim KEY As String = rsItems.Fields.Item("U_KEYC").Value.ToString().Trim()

                ' 2. Get all price lines for this specific Key/Item
                ' We use LTRIM/RTRIM to ensure spaces in the UDT don't break the query
                'Dim lineQry As String = $"SELECT T1.""U_UOMCode"", T1.""U_PriceId"", T1.""U_OldPrice"", T3.""IUoMEntry"", T1.""U_UomEtry"" " &
                '                    $"FROM ""@TNX_ITEMUP_C3"" T1 " &
                '                    $"INNER JOIN OITM T3 ON T3.""ItemCode"" = '{sItemCode}' " &
                '                    $"WHERE T1.""DocEntry"" = '{DOCENTRYX}' AND LTRIM(RTRIM(T1.""U_ItemC"")) = LTRIM(RTRIM('{KEY}'))"


                'rsUomDetails.DoQuery(lineQry)
                ' This query joins OUOM to find the real UomEntry based on the UomCode provided in your UDT.
                ' We use a LEFT JOIN so that if the Code isn't found, we don't lose the row.
                Dim lineQry As String = $"SELECT T1.""U_UOMCode"", T1.""U_PriceId"", T1.""U_OldPrice"", " &
                        $"T3.""IUoMEntry"", " &
                        $"IFNULL(T4.""UomEntry"", T1.""U_UomEtry"") AS ""FinalUomEntry"" " &
                        $"FROM ""@TNX_ITEMUP_C3"" T1 " &
                        $"INNER JOIN OITM T3 ON T3.""ItemCode"" = '{sItemCode}' " &
                        $"LEFT JOIN OUOM T4 ON T4.""UomCode"" = T1.""U_UOMCode"" " &
                        $"WHERE T1.""DocEntry"" = '{DOCENTRYX}' " &
                        $"AND LTRIM(RTRIM(T1.""U_ItemC"")) = LTRIM(RTRIM('{KEY}'))"

                rsUomDetails.DoQuery(lineQry)
                ' Inner Loop: Process each UoM row for this item
                For jRow As Integer = 0 To rsUomDetails.RecordCount - 1

                    ' Re-create Item Object for every line to ensure the version is fresh for the .Update()
                    Dim oItem As SAPbobsCOM.Items = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

                    If oItem.GetByKey(sItemCode) Then
                        Dim targetPL As Integer = CInt(rsUomDetails.Fields.Item("U_PriceId").Value)
                        Dim uomCode As String = rsUomDetails.Fields.Item("U_UOMCode").Value.ToString().Trim()
                        Dim targetUomEntry As Integer = CInt(rsUomDetails.Fields.Item("FinalUomEntry").Value)
                        Dim inventoryUom As Integer = CInt(rsUomDetails.Fields.Item("IUoMEntry").Value)
                        Dim newPrice As Double = CDbl(rsUomDetails.Fields.Item("U_OldPrice").Value)
                        Dim isLineChanged As Boolean = False

                        ' 3. Navigate PriceLists collection
                        For pListIdx As Integer = 0 To oItem.PriceList.Count - 1
                            oItem.PriceList.SetCurrentLine(pListIdx)

                            If oItem.PriceList.PriceList = targetPL Then
                                ' CASE A: It is the Inventory UoM (Updates ITM1)
                                If targetUomEntry = inventoryUom Then
                                    oItem.PriceList.Price = newPrice
                                    isLineChanged = True
                                Else
                                    ' CASE B: It is a specific UoM (Updates ITM9)
                                    Dim uomRowFound As Boolean = False

                                    ' Search existing UoM Price lines
                                    For uomIdx As Integer = 0 To oItem.PriceList.UoMPrices.Count - 1
                                        oItem.PriceList.UoMPrices.SetCurrentLine(uomIdx)
                                        If oItem.PriceList.UoMPrices.UoMEntry = targetUomEntry Then
                                            oItem.PriceList.UoMPrices.Price = newPrice
                                            ' Force Manual: Stops SAP from calculating price based on Base UoM
                                            oItem.PriceList.UoMPrices.Auto = SAPbobsCOM.BoYesNoEnum.tNO
                                            uomRowFound = True
                                            isLineChanged = True
                                            Exit For
                                        End If
                                    Next

                                    ' If the UoM is NOT in ITM9 yet, we must ADD it
                                    If Not uomRowFound Then
                                        oItem.PriceList.UoMPrices.Add()
                                        oItem.PriceList.UoMPrices.UoMEntry = targetUomEntry
                                        oItem.PriceList.UoMPrices.Price = newPrice
                                        oItem.PriceList.UoMPrices.Auto = SAPbobsCOM.BoYesNoEnum.tNO
                                        isLineChanged = True
                                    End If
                                End If
                                Exit For
                            End If
                        Next

                        ' 4. ✅ UPDATE IMMEDIATELY AFTER EACH LINE
                        If isLineChanged Then
                            If oItem.Update() <> 0 Then
                                ' ERROR: Red Status Bar
                                objMain.objApplication.StatusBar.SetText($"FAILED: {sItemCode} | {uomCode} | {objMain.objCompany.GetLastErrorDescription()}",
                                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Else
                                ' SUCCESS: Green Status Bar showing Price & UoM
                                objMain.objApplication.StatusBar.SetText($"SUCCESS: {sItemCode} | {uomCode} | Updated to {newPrice}",
                                SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            End If
                        End If
                    End If

                    ' Release the Item Object to avoid memory leaks
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem)
                    rsUomDetails.MoveNext()
                Next

                rsItems.MoveNext()
            Next

            objMain.objApplication.StatusBar.SetText("Price Update completed successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Critical Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            ' Cleanup Recordsets
            If rsItems IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(rsItems)
            If rsUomDetails IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(rsUomDetails)
        End Try
    End Sub

    Sub UpdatePriceForMultipleBranches(ByVal DocEntry As String)
        Try
            Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim Query3 As String = "SELECT * FROM ""@TNX_ITEMUP"" T0 INNER JOIN ""@TNX_ITEMUP_C0"" T1 ON T0.""DocEntry""=T1.""DocEntry"" WHERE ""DocNum""='" & DocEntry & "' AND ""U_ITMCD""<>'' AND (IFNULL(T1.""U_PStst"",'') <> 'Priceupdated Succesfully')"
            rs.DoQuery(Query3)

            If rs.RecordCount = 0 Then Exit Sub

            ' Cache branch price lists once (used in both BranchId=200 and else cases)
            Dim oRsGetBranchDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetBranchDtls.DoQuery("SELECT ""U_PricelistNo"" AS ""PricListId"" FROM ""OBPL"" WHERE ""U_PricelistNo"" <> ''")

            Dim branchPriceLists As New List(Of Integer)
            While Not oRsGetBranchDtls.EoF
                branchPriceLists.Add(CInt(oRsGetBranchDtls.Fields.Item("PricListId").Value))
                oRsGetBranchDtls.MoveNext()
            End While

            Dim Linenum As Integer = 0
            While Not rs.EoF
                Dim FItemCode As String = rs.Fields.Item("U_ITMCD").Value.ToString()
                Dim FKeyCode As String = rs.Fields.Item("U_KEYC").Value.ToString()
                Linenum = rs.Fields.Item("LineId").Value.ToString()

                Dim GetItemPriceLineDtls As String =
                "WITH B AS (" &
                "SELECT T1.""U_UOMCode"", T1.""U_UomEtry"", T3.""IUoMEntry"" AS ""DU"", T1.""U_PriceId"" AS ""PI"", T1.""U_BrnchId"" AS ""BI"", T1.""U_OldPrice"" AS ""NP"" " &
                "FROM ""@TNX_ITEMUP"" T0 " &
                "JOIN ""@TNX_ITEMUP_C3"" T1 ON T0.""DocEntry""=T1.""DocEntry"" " &
                "JOIN ""@TNX_ITEMUP_C0"" T2 ON T0.""DocEntry""=T2.""DocEntry"" " &
                "JOIN OITM T3 ON T3.""ItemCode""=T2.""U_ITMCD"" " &
                "WHERE T0.""DocNum""='" & DocEntry & "' AND T1.""U_ItemC""='" & FKeyCode & "' AND T1.""U_OldPrice"">0 AND T1.""U_UomEtry""<>'') " &
                "SELECT DISTINCT IFNULL(U.""UomEntry"", B.""U_UomEtry"") AS ""UomEntry"", B.""DU"" AS ""DefaultUOM"", B.""PI"" AS ""PricListId"", B.""BI"" AS ""BranchId"", B.""NP"" AS ""New Price"" " &
                "FROM B LEFT JOIN OUOM U ON U.""UomCode""=B.""U_UOMCode"" WHERE B.""PI""<>'' " &
                "UNION " &
                "SELECT DISTINCT IFNULL(U.""UomEntry"", B.""U_UomEtry""), B.""DU"", BR.""U_PricelistNo"", B.""BI"", (SELECT MAX(""NP"") FROM B WHERE ""PI""='1') " &
                "FROM B LEFT JOIN OUOM U ON U.""UomCode""=B.""U_UOMCode"" CROSS JOIN OBPL BR WHERE BR.""U_PricelistNo""<>'' ORDER BY 3"

                GetItemPriceLineDtls = "WITH B AS (SELECT T1.""U_UOMCode"" AS ""UC"", T1.""U_UomEtry"" AS ""UE"", T3.""IUoMEntry"" AS ""DU"", T1.""U_PriceId"" AS ""PI"", T1.""U_BrnchId"" AS ""BI"", T1.""U_OldPrice"" AS ""NP"" " &
    "FROM ""@TNX_ITEMUP"" T0 JOIN ""@TNX_ITEMUP_C3"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" JOIN ""@TNX_ITEMUP_C0"" T2 ON T0.""DocEntry"" = T2.""DocEntry"" JOIN OITM T3 ON T3.""ItemCode"" = T2.""U_ITMCD"" " &
    "WHERE T0.""DocNum"" = '" & DocEntry & "' AND T1.""U_ItemC"" = '10' AND T1.""U_OldPrice"" > 0 AND T1.""U_UomEtry"" <> '') " &
    "SELECT DISTINCT IFNULL(U.""UomEntry"", B.""UE"") AS ""UomEntry"", B.""DU"" AS ""DefaultUOM"", B.""PI"" AS ""PricListId"", B.""BI"" AS ""BranchId"", B.""NP"" AS ""New Price"" FROM B LEFT JOIN OUOM U ON U.""UomCode"" = B.""UC"" WHERE B.""PI"" <> '' " &
    "UNION SELECT DISTINCT IFNULL(U.""UomEntry"", B.""UE""), B.""DU"", BR.""U_PricelistNo"", B.""BI"", B1.""NP"" FROM B LEFT JOIN OUOM U ON U.""UomCode"" = B.""UC"" CROSS JOIN OBPL BR " &
    "LEFT JOIN B B1 ON B1.""PI"" = '1' AND IFNULL(U.""UomEntry"", B.""UE"") = IFNULL((SELECT U2.""UomEntry"" FROM OUOM U2 WHERE U2.""UomCode"" = B1.""UC""), B1.""UE"") WHERE BR.""U_PricelistNo"" <> '' ORDER BY 3"

                Dim oRsGetItemPriceLineDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetItemPriceLineDtls.DoQuery(GetItemPriceLineDtls)

                If oRsGetItemPriceLineDtls.RecordCount > 0 Then
                    Dim oItem As SAPbobsCOM.Items = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

                    If oItem.GetByKey(FItemCode) Then
                        While Not oRsGetItemPriceLineDtls.EoF
                            Dim UomEntry As Integer = CInt(oRsGetItemPriceLineDtls.Fields.Item("UomEntry").Value)
                            Dim DefaultUom As Integer = CInt(oRsGetItemPriceLineDtls.Fields.Item("DefaultUOM").Value)
                            Dim BranchId As Integer = CInt(oRsGetItemPriceLineDtls.Fields.Item("BranchId").Value)
                            Dim PriceListId As Integer = CInt(oRsGetItemPriceLineDtls.Fields.Item("PricListId").Value)
                            Dim NewPrice As Double = CDbl(oRsGetItemPriceLineDtls.Fields.Item("New Price").Value)

                            For Each PriceListRow As Integer In branchPriceLists
                                ' For BranchId <> 200, only process matching price list
                                If BranchId <> 200 AndAlso PriceListRow <> PriceListId Then Continue For

                                oItem.PriceList.SetCurrentLine(PriceListRow - 1)
                                Dim strPrlist As String = oItem.PriceList.PriceList.ToString()

                                ' For BranchId = 200, process all; otherwise check if price list matches
                                If BranchId <> 200 AndAlso strPrlist <> PriceListId.ToString() Then Continue For

                                If DefaultUom = UomEntry Then
                                    oItem.PriceList.Price = NewPrice
                                Else
                                    UpdateUomPrice(oItem, FItemCode, UomEntry, NewPrice, strPrlist)
                                End If
                            Next

                            oRsGetItemPriceLineDtls.MoveNext()
                        End While

                        ' Update item once after all price changes
                        Dim RetValue As Integer = oItem.Update()
                        UpdatePriceStatus(DocEntry, Linenum, RetValue, FItemCode)
                    End If
                End If

                '    lineNum += 1
                rs.MoveNext()
            End While

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub UpdateUomPrice(oItem As SAPbobsCOM.Items, itemCode As String, uomEntry As Integer, newPrice As Double, priceList As String)
        Dim UOMFoundFlag As Boolean = False

        ' Check if UOM exists for this item
        Dim oRschkIfExsist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRschkIfExsist.DoQuery("SELECT COUNT(*) FROM ITM9 WHERE ""ItemCode"" = '" & itemCode & "' AND ""UomEntry"" = '" & uomEntry & "'")

        If oRschkIfExsist.Fields.Item(0).Value > 0 Then
            For k As Integer = 0 To oItem.PriceList.UoMPrices.Count - 1
                oItem.PriceList.UoMPrices.SetCurrentLine(k)
                If oItem.PriceList.UoMPrices.UoMEntry = uomEntry Then
                    oItem.PriceList.UoMPrices.Price = newPrice
                    UOMFoundFlag = True
                    Exit For
                End If
            Next
        End If

        If Not UOMFoundFlag Then
            Dim oRsGetUomPriceListCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ' oRsGetUomPriceListCount.DoQuery("SELECT COUNT(*) FROM ITM9 WHERE ""ItemCode"" = '" & itemCode & "' AND ""PriceList"" = '" & priceList & "'")
            oRsGetUomPriceListCount.DoQuery("SELECT COUNT(*) FROM ITM9 WHERE ""ItemCode"" = '" & itemCode & "'  And ""UomEntry"" = '" & uomEntry & "'")

            Dim UomCount As Integer = CInt(oRsGetUomPriceListCount.Fields.Item(0).Value)

            'oItem.PriceList.UoMPrices.Add()
            'oItem.PriceList.UoMPrices.SetCurrentLine(UomCount)
            'oItem.PriceList.UoMPrices.Price = newPrice
            'oItem.PriceList.UoMPrices.UoMEntry = uomEntry
            'oItem.PriceList.UoMPrices.PriceList = priceList
            'oItem.PriceList.UoMPrices.Auto = SAPbobsCOM.BoYesNoEnum.tNO
            ' Set current line FIRST

            oItem.PriceList.UoMPrices.SetCurrentLine(UomCount)

            ' Assign values
            oItem.PriceList.UoMPrices.Price = newPrice
            oItem.PriceList.UoMPrices.UoMEntry = uomEntry
            oItem.PriceList.UoMPrices.PriceList = priceList
            oItem.PriceList.UoMPrices.Auto = SAPbobsCOM.BoYesNoEnum.tNO

            ' Add new row ONLY if more rows needed
            oItem.PriceList.UoMPrices.Add()

        End If
    End Sub

    Private Sub UpdatePriceStatus(docEntry As String, lineId As Integer, retValue As Integer, itemCode As String)
        Dim statusMsg As String
        If retValue <> 0 Then
            statusMsg = objMain.objCompany.GetLastErrorDescription.ToString()
            If statusMsg.Length > 200 Then statusMsg = statusMsg.Substring(0, 200)
            statusMsg = statusMsg.Replace("'", "")
        Else
            statusMsg = "Priceupdated Succesfully"
            objMain.objApplication.StatusBar.SetText("Price are updating succesfully for " & itemCode, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
        End If

        Dim rSe As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rSe.DoQuery("UPDATE ""@TNX_ITEMUP_C0"" SET ""U_PStst""='" & statusMsg & "' WHERE ""LineId"" = '" & lineId & "' AND ""DocEntry""=(SELECT ""DocEntry"" FROM ""@TNX_ITEMUP"" WHERE ""DocNum""= '" & docEntry & "')")
    End Sub



    Public Sub UpdateSpecialPriceBP(ByVal DocEntry As String)
        Try
            Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim Query3 As String = "SELECT * FROM ""@TNX_ITEMUP"" T0 INNER JOIN ""@TNX_ITEMUP_C0""  T1 ON T0.""DocEntry""=T1.""DocEntry"" where ""DocNum""='" & DocEntry & "'and ""U_ITMCD""<>'' AND ( IFNULL(T1.""U_PstatV"",'') <> 'Priceupdated Succesfully')"
            rs.DoQuery(Query3)

            If rs.RecordCount > 0 Then
                For i As Integer = 1 To rs.RecordCount
                    Dim FItemCode As String = rs.Fields.Item("U_ITMCD").Value
                    Dim FKeyCode As String = rs.Fields.Item("U_KEYC").Value

                    Dim GetItemPriceLineDtls As String = "Select distinct ""U_CCode"" As ""CardCode"",T3.""IUoMEntry"" As ""DefaultUOM"",""U_PriceId"" As ""PricListId"",""U_Price"" As ""New Price"",""U_UEntry""as ""UomEntry"" from ""@TNX_ITEMUP"" T0 inner join ""@TNX_ITEMUP_C4"" T1 On T0.""DocEntry""= T1.""DocEntry""  inner join ""@TNX_ITEMUP_C0"" T2 On T0.""DocEntry""= T2.""DocEntry"" " &
                    "Inner Join OITM T3 ON T3.""ItemCode"" = T2.""U_ITMCD"" Inner Join OUGP A On A.""UgpEntry"" = T3.""UgpEntry"" Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry""   Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry"" " &
                                        "where T0.""DocNum""='" & DocEntry & "'and ""U_Code"" ='" & FKeyCode & "'and ""U_Price"">0 and ""U_PriceId"" <>'' "
                    Dim oRsGetItemPriceLineDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetItemPriceLineDtls.DoQuery(GetItemPriceLineDtls)
                    If oRsGetItemPriceLineDtls.RecordCount > 0 Then
                        ''  oRsGetItemPriceLineDtls.MoveFirst()
                        Dim UomCount As Integer = 0
                        Dim retV As Integer = 0
                        Dim r As Integer = 0
                        Dim spItemCode As String = ""
                        Dim spCardCode As String = ""

                        For k As Integer = 1 To oRsGetItemPriceLineDtls.RecordCount
                            Dim oSpp As SAPbobsCOM.SpecialPrices = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSpecialPrices)
                            spItemCode = FItemCode
                            spCardCode = oRsGetItemPriceLineDtls.Fields.Item("CardCode").Value
                            Dim NewPrice As Double = CDbl(oRsGetItemPriceLineDtls.Fields.Item("New Price").Value)
                            Dim UomEntry As Integer = CInt(oRsGetItemPriceLineDtls.Fields.Item("UomEntry").Value)
                            If oSpp.GetByKey(spItemCode, spCardCode) Then
                                If NewPrice > 0 Then
                                    Dim DefaultUom As Integer = CInt(oRsGetItemPriceLineDtls.Fields.Item("DefaultUOM").Value)
                                    Dim PriceListRow As Integer = oRsGetItemPriceLineDtls.Fields.Item("PricListId").Value
                                    oSpp.ItemCode = spItemCode
                                    oSpp.CardCode = spCardCode
                                    ' oSpp.Currency = "QR"
                                    oSpp.PriceListNum = 0
                                    oSpp.Valid = SAPbobsCOM.BoYesNoEnum.tNO

                                    If DefaultUom = UomEntry Then

                                        Dim GetUomPriceListCount As String = "Select IFNULL(""SPP2LNum"",0) From SPP2 " &
                               " Where ""ItemCode"" = '" & spItemCode & "' and ""CardCode"" = '" & spCardCode & "' and ""UomEntry"" = '" & UomEntry & "' "
                                        Dim oRsGetUomPriceListCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRsGetUomPriceListCount.DoQuery(GetUomPriceListCount)
                                        If oRsGetUomPriceListCount.Fields.Item(0).Value = 0 Then
                                            UomCount = 0
                                        Else
                                            UomCount = oRsGetUomPriceListCount.Fields.Item(0).Value
                                        End If

                                        oSpp.Price = CDbl(NewPrice)
                                        oSpp.SpecialPricesDataAreas.PriceListNo = oSpp.PriceListNum
                                        oSpp.SpecialPricesDataAreas.DateFrom = Now
                                        oSpp.SpecialPricesDataAreas.Dateto = Now
                                        oSpp.SpecialPricesDataAreas.SpecialPrice = CDbl(NewPrice)
                                        oSpp.SpecialPricesDataAreas.Discount = 0

                                        oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.SetCurrentLine(UomCount)
                                        oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.UoMEntry = UomEntry
                                        ' oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.PriceCurrency = "QR"
                                        oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.Quantity = 1
                                        oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.SpecialPrice = CDbl(NewPrice)

                                    Else
                                        Dim UOMFoundFlag As String = ""
                                        Dim chkIfExsist As String = "Select Count(*) From SPP2 " &
                                " Where ""ItemCode"" = '" & spItemCode & "' and ""CardCode"" = '" & spCardCode & "' " &
                                " And ""UomEntry"" = '" & UomEntry & "'"
                                        Dim oRschkIfExsist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRschkIfExsist.DoQuery(chkIfExsist)
                                        If oRschkIfExsist.Fields.Item(0).Value > 0 Then
                                            For j As Integer = 0 To oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.Count - 1
                                                oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.SetCurrentLine(j)
                                                If oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.UoMEntry = UomEntry Then
                                                    oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.UoMEntry = UomEntry
                                                    ' oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.PriceCurrency = "QR"
                                                    oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.Quantity = 1
                                                    oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.SpecialPrice = CDbl(NewPrice)
                                                    UOMFoundFlag = "Yes"
                                                End If
                                            Next
                                        End If
                                        If UOMFoundFlag = "" Then
                                            Dim GetUomPriceListCount As String = "Select Count(*) From SPP2 " &
                                    " Where ""ItemCode"" = '" & spItemCode & "' and ""CardCode"" = '" & spCardCode & "' "
                                            Dim oRsGetUomPriceListCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRsGetUomPriceListCount.DoQuery(GetUomPriceListCount)
                                            If oRsGetUomPriceListCount.Fields.Item(0).Value = 0 Then
                                                UomCount = 0
                                            Else
                                                UomCount = oRsGetUomPriceListCount.Fields.Item(0).Value
                                            End If
                                            oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.Add()
                                            oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.SetCurrentLine(UomCount)
                                            oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.UoMEntry = UomEntry
                                            ' oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.PriceCurrency = "QR"
                                            oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.Quantity = 1
                                            oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.SpecialPrice = CDbl(NewPrice)

                                        End If
                                    End If
                                End If
                                retV = oSpp.Update()
                                If retV <> 0 Then
                                    Dim test As String = objMain.objCompany.GetLastErrorDescription
                                End If
                            Else 'Add

                                oSpp.AutoUpdate = SAPbobsCOM.BoYesNoEnum.tNO
                                oSpp.CardCode = spCardCode
                                oSpp.ItemCode = spItemCode
                                ' oSpp.Currency = "QR"
                                oSpp.PriceListNum = 0
                                oSpp.Valid = SAPbobsCOM.BoYesNoEnum.tNO
                                oSpp.Price = CDbl(NewPrice)

                                'oSpp.SpecialPricesDataAreas.Add()
                                oSpp.SpecialPricesDataAreas.SetCurrentLine(0)
                                oSpp.SpecialPricesDataAreas.PriceListNo = oSpp.PriceListNum

                                'oSpp.SpecialPricesDataAreas.DateFrom = #11/21/2007 12:14:00 PM#
                                'oSpp.SpecialPricesDataAreas.Dateto = #11/30/2007 12:14:00 PM#

                                oSpp.SpecialPricesDataAreas.DateFrom = Now
                                'oSpp.SpecialPricesDataAreas.Dateto = Now

                                oSpp.SpecialPricesDataAreas.AutoUpdate = SAPbobsCOM.BoYesNoEnum.tNO
                                oSpp.SpecialPricesDataAreas.SpecialPrice = CDbl(NewPrice)
                                'oSpp.SpecialPricesDataAreas.Discount = 0

                                'oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.Add()
                                oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.SetCurrentLine(0)
                                ' oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.PriceCurrency = "QR"
                                oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.UoMEntry = UomEntry
                                oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.Quantity = 1
                                oSpp.SpecialPricesDataAreas.SpecialPricesQuantityAreas.SpecialPrice = CDbl(NewPrice)

                                retV = oSpp.Add()
                                If retV <> 0 Then
                                    Dim test As String = objMain.objCompany.GetLastErrorDescription
                                    Dim qSt As String = "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_PstatV""='" & test & "' WHERE  ""LineId"" = '" & i & "' and ""DocEntry""=(Select ""DocEntry"" from  ""@TNX_ITEMUP"" where ""DocNum""= '" & DocEntry & "')"
                                    Dim rSe As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    rSe.DoQuery(qSt)
                                Else
                                    Dim qSt As String = "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_PstatV""='Priceupdated Succesfully' WHERE  ""LineId"" = '" & i & "' and ""DocEntry""=(Select ""DocEntry"" from  ""@TNX_ITEMUP"" where ""DocNum""= '" & DocEntry & "')"
                                    Dim rSe As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    rSe.DoQuery(qSt)
                                    objMain.objApplication.StatusBar.SetText("Prices are updating succesfully for " & FItemCode & " ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)

                                End If
                            End If
                            oRsGetItemPriceLineDtls.MoveNext()
                        Next
                    End If
                    rs.MoveNext()
                Next
            End If


        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub UpdatMinistryPriceold(ByVal DocEntry As String)
        Dim UomCount As Integer = 0
        Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim Query3 As String = "SELECT * FROM ""@TNX_ITEMUP"" T0 INNER JOIN ""@TNX_ITEMUP_C0""  T1 ON T0.""DocEntry""=T1.""DocEntry"" where ""DocNum""='" & DocEntry & "'and ""U_ITMCD""<>''"

        rs.DoQuery(Query3)

        If rs.RecordCount > 0 Then
            For i As Integer = 1 To rs.RecordCount
                Dim FItemCode As String = rs.Fields.Item("U_ITMCD").Value
                Dim FKeyCode As String = rs.Fields.Item("U_KEYC").Value
                Dim oItem As SAPbobsCOM.Items = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                If oItem.GetByKey(FItemCode) Then
                    Dim GetItemPriceLineDtls As String = "select distinct ""U_UomEtry"" As ""UomEntry"",T3.""IUoMEntry"" As ""DefaultUOM"",""U_PriceId"" As ""PricListId"",""U_Price"" As ""New Price"" from ""@TNX_ITEMUP"" T0 inner join ""@TNX_ITEMUP_C5"" T1 On T0.""DocEntry""= T1.""DocEntry""  inner join ""@TNX_ITEMUP_C0"" T2 On T0.""DocEntry""= T2.""DocEntry"" " &
        "Inner Join OITM T3 ON T3.""ItemCode"" = T2.""U_ITMCD"" Inner Join OUGP A On A.""UgpEntry"" = T3.""UgpEntry"" Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry""   Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry"" " &
                                    "where T0.""DocNum""='" & DocEntry & "'and ""U_Code"" ='" & FKeyCode & "'and ""U_Price"">0 and ""U_UomEtry""<>'' and ""U_PriceId""<>'' "
                    Dim oRsGetItemPriceLineDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetItemPriceLineDtls.DoQuery(GetItemPriceLineDtls)

                    If oRsGetItemPriceLineDtls.RecordCount > 0 Then
                        oRsGetItemPriceLineDtls.MoveFirst()
                        For j As Integer = 1 To oRsGetItemPriceLineDtls.RecordCount
                            Dim UomEntry As Integer = CInt(oRsGetItemPriceLineDtls.Fields.Item("UomEntry").Value)
                            Dim DefaultUom As Integer = CInt(oRsGetItemPriceLineDtls.Fields.Item("DefaultUOM").Value)
                            Dim PRList As String = "Select ""ListNum"" From ""OPLN"" Where ""ListNum"" = '100'"
                            Dim oRsPRList As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsPRList.DoQuery(PRList)
                            Dim PriceListRow As Integer = oRsPRList.Fields.Item(0).Value   'MinistryPrice'
                            oItem.PriceList.SetCurrentLine(PriceListRow - 1)
                            Dim strPrlist As String = oItem.PriceList.PriceList
                            If strPrlist = oRsGetItemPriceLineDtls.Fields.Item("PricListId").Value Then
                                If DefaultUom = UomEntry Then
                                    oItem.PriceList.Price = CDbl(oRsGetItemPriceLineDtls.Fields.Item("New Price").Value)

                                Else
                                    Dim UOMFoundFlag As String = ""
                                    Dim chkIfExsist As String = "Select Count(*) From ITM9 Where ""ItemCode"" = '" & FItemCode & "' And ""UomEntry"" = '" & UomEntry & "'"
                                    Dim oRschkIfExsist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRschkIfExsist.DoQuery(chkIfExsist)
                                    If oRschkIfExsist.Fields.Item(0).Value > 0 Then
                                        For k As Integer = 0 To oItem.PriceList.UoMPrices.Count - 1
                                            oItem.PriceList.UoMPrices.SetCurrentLine(k)
                                            Dim iUom As Integer = oItem.PriceList.UoMPrices.UoMEntry
                                            If iUom = UomEntry Then
                                                oItem.PriceList.UoMPrices.Price = CDbl(oRsGetItemPriceLineDtls.Fields.Item("New Price").Value)

                                                UOMFoundFlag = "Yes"
                                            End If
                                        Next
                                    End If
                                    If UOMFoundFlag = "" Then
                                        Dim GetUomPriceListCount As String = "Select Count(*) From ITM9 Where ""ItemCode"" = '" & FItemCode & "' And ""PriceList"" = '" & strPrlist & "'"
                                        Dim oRsGetUomPriceListCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRsGetUomPriceListCount.DoQuery(GetUomPriceListCount)
                                        If oRsGetUomPriceListCount.Fields.Item(0).Value = 0 Then
                                            UomCount = 0
                                        Else
                                            UomCount = oRsGetUomPriceListCount.Fields.Item(0).Value
                                        End If
                                        oItem.PriceList.UoMPrices.Add()
                                        oItem.PriceList.UoMPrices.SetCurrentLine(UomCount)
                                        oItem.PriceList.UoMPrices.Price = CDbl(oRsGetItemPriceLineDtls.Fields.Item("New Price").Value)
                                        oItem.PriceList.UoMPrices.UoMEntry = UomEntry
                                        oItem.PriceList.UoMPrices.PriceList = strPrlist
                                        oItem.PriceList.UoMPrices.Auto = SAPbobsCOM.BoYesNoEnum.tNO
                                        oItem.PriceList.UoMPrices.Currency = "QR"
                                    End If
                                End If
                            End If
                            oRsGetItemPriceLineDtls.MoveNext()
                        Next
                    End If

                    Dim RetValue As Integer = oItem.Update
                    If RetValue <> 0 Then
                        Dim Err As String = objMain.objCompany.GetLastErrorDescription.Replace("'", "")
                        Dim qSt1 As String = "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_MPUS""='" & Err & "' WHERE  ""LineId"" = '" & i & "' and ""DocEntry""=(Select ""DocEntry"" from  ""@TNX_ITEMUP"" where ""DocNum""= '" & DocEntry & "')"
                        Dim rSe1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        rSe1.DoQuery(qSt1)
                    Else
                        Dim qSt As String = "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_MPUS""='Priceupdated Succesfully' WHERE  ""LineId"" = '" & i & "' and ""DocEntry""=(Select ""DocEntry"" from  ""@TNX_ITEMUP"" where ""DocNum""= '" & DocEntry & "')"
                        Dim rSe As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        rSe.DoQuery(qSt)
                    End If
                End If
                rs.MoveNext()
            Next
        End If


    End Sub
    Sub UpdatMinistryPrice(ByVal DocEntry As String)
        Dim UomCount As Integer = 0
        Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim Query3 As String = "SELECT * FROM ""@TNX_ITEMUP"" T0 INNER JOIN ""@TNX_ITEMUP_C0""  T1 ON T0.""DocEntry""=T1.""DocEntry"" where ""DocNum""='" & DocEntry & "'and ""U_ITMCD""<>''"

        rs.DoQuery(Query3)

        If rs.RecordCount > 0 Then
            For i As Integer = 1 To rs.RecordCount
                Dim FItemCode As String = rs.Fields.Item("U_ITMCD").Value
                Dim FKeyCode As String = rs.Fields.Item("U_KEYC").Value
                Dim GetItemPriceLineDtls As String = "select distinct ""U_UomEtry"" As ""UomEntry"",T3.""IUoMEntry"" As ""DefaultUOM"",""U_PriceId"" As ""PricListId"",""U_Price"" As ""New Price"" from ""@TNX_ITEMUP"" T0 inner join ""@TNX_ITEMUP_C5"" T1 On T0.""DocEntry""= T1.""DocEntry""  inner join ""@TNX_ITEMUP_C0"" T2 On T0.""DocEntry""= T2.""DocEntry"" " &
        "Inner Join OITM T3 ON T3.""ItemCode"" = T2.""U_ITMCD"" Inner Join OUGP A On A.""UgpEntry"" = T3.""UgpEntry"" Inner Join UGP1 B On B.""UgpEntry"" = A.""UgpEntry""   Inner Join OUOM C On B.""UomEntry"" = C.""UomEntry"" " &
                                    "where T0.""DocNum""='" & DocEntry & "'and ""U_Code"" ='" & FKeyCode & "'and ""U_Price"">0 and ""U_UomEtry""<>'' and ""U_PriceId""<>'' "
                Dim oRsGetItemPriceLineDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetItemPriceLineDtls.DoQuery(GetItemPriceLineDtls)

                If oRsGetItemPriceLineDtls.RecordCount > 0 Then
                    oRsGetItemPriceLineDtls.MoveFirst()
                    For j As Integer = 1 To oRsGetItemPriceLineDtls.RecordCount
                        Dim UomEntry As Integer = CInt(oRsGetItemPriceLineDtls.Fields.Item("UomEntry").Value)
                        Dim DefaultUom As Integer = CInt(oRsGetItemPriceLineDtls.Fields.Item("DefaultUOM").Value)
                        Dim oItem As SAPbobsCOM.Items = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                        If oItem.GetByKey(FItemCode) Then
                            Dim PRList As String = "Select ""ListNum"" From ""OPLN"" Where ""ListNum"" = '100'"
                            Dim oRsPRList As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsPRList.DoQuery(PRList)
                            Dim PriceListRow As Integer = oRsPRList.Fields.Item(0).Value   'MinistryPrice'
                            oItem.PriceList.SetCurrentLine(PriceListRow - 1)
                            Dim strPrlist As String = oItem.PriceList.PriceList
                            If strPrlist = oRsGetItemPriceLineDtls.Fields.Item("PricListId").Value Then
                                If DefaultUom = UomEntry Then
                                    oItem.PriceList.Price = CDbl(oRsGetItemPriceLineDtls.Fields.Item("New Price").Value)

                                Else
                                    Dim UOMFoundFlag As String = ""
                                    Dim chkIfExsist As String = "Select Count(*) From ITM9 Where ""ItemCode"" = '" & FItemCode & "' And ""UomEntry"" = '" & UomEntry & "'"
                                    Dim oRschkIfExsist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRschkIfExsist.DoQuery(chkIfExsist)
                                    If oRschkIfExsist.Fields.Item(0).Value > 0 Then
                                        For k As Integer = 0 To oItem.PriceList.UoMPrices.Count - 1
                                            oItem.PriceList.UoMPrices.SetCurrentLine(k)
                                            Dim iUom As Integer = oItem.PriceList.UoMPrices.UoMEntry
                                            If iUom = UomEntry Then
                                                oItem.PriceList.UoMPrices.Price = CDbl(oRsGetItemPriceLineDtls.Fields.Item("New Price").Value)

                                                UOMFoundFlag = "Yes"
                                            End If
                                        Next
                                    End If
                                    If UOMFoundFlag = "" Then
                                        Dim GetUomPriceListCount As String = "Select Count(*) From ITM9 Where ""ItemCode"" = '" & FItemCode & "' And ""PriceList"" = '" & strPrlist & "'"
                                        Dim oRsGetUomPriceListCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRsGetUomPriceListCount.DoQuery(GetUomPriceListCount)
                                        If oRsGetUomPriceListCount.Fields.Item(0).Value = 0 Then
                                            UomCount = 0
                                        Else
                                            UomCount = oRsGetUomPriceListCount.Fields.Item(0).Value
                                        End If
                                        oItem.PriceList.UoMPrices.Add()
                                        oItem.PriceList.UoMPrices.SetCurrentLine(UomCount)
                                        oItem.PriceList.UoMPrices.Price = CDbl(oRsGetItemPriceLineDtls.Fields.Item("New Price").Value)
                                        oItem.PriceList.UoMPrices.UoMEntry = UomEntry
                                        oItem.PriceList.UoMPrices.PriceList = strPrlist
                                        oItem.PriceList.UoMPrices.Auto = SAPbobsCOM.BoYesNoEnum.tNO
                                        oItem.PriceList.UoMPrices.Currency = "QR"
                                    End If
                                End If
                            End If
                            Dim RetValue As Integer = oItem.Update
                            If RetValue <> 0 Then
                                Dim Err As String = objMain.objCompany.GetLastErrorDescription.Replace("'", "")
                                Dim qSt1 As String = "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_MPUS""='" & Err & "' WHERE  ""LineId"" = '" & i & "' and ""DocEntry""=(Select ""DocEntry"" from  ""@TNX_ITEMUP"" where ""DocNum""= '" & DocEntry & "')"
                                Dim rSe1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                rSe1.DoQuery(qSt1)
                            Else
                                Dim qSt As String = "UPDATE ""@TNX_ITEMUP_C0"" SET ""U_MPUS""='Priceupdated Succesfully' WHERE  ""LineId"" = '" & i & "' and ""DocEntry""=(Select ""DocEntry"" from  ""@TNX_ITEMUP"" where ""DocNum""= '" & DocEntry & "')"
                                Dim rSe As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                rSe.DoQuery(qSt)
                            End If

                        End If
                        oRsGetItemPriceLineDtls.MoveNext()
                    Next
                End If
                rs.MoveNext()
            Next
        End If


    End Sub
    Sub StartBackgroundForLoading1(Optional ByVal BackGroundWrkr As Boolean = True)
        Try
            If BackGroundWrkr = True Then
                If BackgroundWorker_GenerateDocumets.IsBusy <> True Then
                    BackgroundWorker_GenerateDocumets.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Calling Unselecting worker Error : " & ex.Message)
        End Try
    End Sub
    Private Sub BackgroundWorker_Loading_DoWork1(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_GenerateDocumets.DoWork
        Dim DocEntry As String = objForm.Items.Item("DocNum").Specific.Value
        ' Me.UpdatePrice(DocEntry)
        Try
            objMain.objApplication.StatusBar.SetText("Price Updating started...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            '    Me.UpdateWarehouseBlockPurchase(DocEntry)
            Me.UpdateSugarTax_UOMWise(DocEntry)

            objMain.objApplication.StatusBar.SetText("Branch Price Updating started...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
            Me.UpdatePriceForMultipleBranches_SL(DocEntry)

            '  Me.UpdatePriceForMultipleBranches(DocEntry)

            objMain.objApplication.StatusBar.SetText("Vendor Price Updating started...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)

            Me.UpdateSpecialPriceBP(DocEntry)

            objMain.objApplication.StatusBar.SetText("Operation Completed Successfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
            objMain.objApplication.ActivateMenuItem("1304")
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error: " & ex.Message, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Error)
        End Try
        ' Me.UpdatMinistryPrice(DocEntry)
    End Sub

    Public Sub SyncBarcodeUDFs(ByVal sDocNum As String)
        Dim oRec As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        ' 1. The Selection Query
        Dim sSQL As String = "SELECT T2.""U_ITMCD"", T1.""U_BarCode"",""U_ADESC"", T1.""U_STAX"", T1.""U_SUTAXAMT"", T1.""U_ETAX"", T1.""U_Price"" " &
                         "FROM ""@TNX_ITEMUP"" T0 " &
                         "INNER JOIN ""@TNX_ITEMUP_C2"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
                         "INNER JOIN ""@TNX_ITEMUP_C0"" T2 ON T0.""DocEntry"" = T2.""DocEntry"" AND T1.""U_KEYC"" = T2.""U_KEYC"" " &
                         "WHERE T0.""DocNum"" = '" & sDocNum & "' AND T1.""U_BarCode"" <> '' and  T2.""U_ITMCD""<>''"

        oRec.DoQuery(sSQL)

        While Not oRec.EoF
            Dim sBarcode As String = oRec.Fields.Item("U_BarCode").Value.ToString()
            Dim sSTAX As String = oRec.Fields.Item("U_STAX").Value.ToString()
            If UCase(sSTAX) = "YES" Then
                sSTAX = "Yes"
            Else
                sSTAX = "No"
            End If
            Dim dTaxAmt As Double = CDbl(oRec.Fields.Item("U_SUTAXAMT").Value)
            Dim sETAX As String = oRec.Fields.Item("U_ETAX").Value.ToString()
            Dim dPrice As Double = CDbl(oRec.Fields.Item("U_Price").Value)
            Dim ADESC As String = (oRec.Fields.Item("U_ADESC").Value.ToString())
            ' 2. Construct the SQL Update for OBCD
            Dim sqlUpdate As String = "UPDATE OBCD SET " &
                                 """U_STAX"" = '" & sSTAX & "', " &
                                 """U_SUTAXAMT"" = " & dTaxAmt & ", " &
                                 """U_ETAX"" = '" & sETAX & "', " &
                                 """U_Price"" = " & dPrice & " " &
                                 ",""U_ArbName"" = '" & ADESC & "' " &
                                 "WHERE ""BcdCode"" = '" & sBarcode & "'"

            ' Execute Update
            Dim oUpdateRec As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oUpdateRec.DoQuery(sqlUpdate)

            ' 3. Update UI DataSource (oDS3) if needed
            ' This ensures your screen shows the values you just processed


            oRec.MoveNext()
        End While


    End Sub
    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window
        Private _hwnd As IntPtr
        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub
        Public ReadOnly Property Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property
    End Class

    Public Function GetFieldValue(ByVal query As String) As Object
        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs.DoQuery(query)
        Dim str As String = oRs.Fields.Item(0).Value.ToString()
        Return str
    End Function

    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDS = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP")
            oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP_C2")
            oDS5 = Me.objForm.DataSources.DBDataSources.Add("@TNX_ITEMUP_C4")

            Dim objMatrixItm As Matrix = objForm.Items.Item("ItmDetMtx").Specific
            objMatrix = objForm.Items.Item("BarMtx").Specific
            Dim objMatrixVendor As Matrix = objForm.Items.Item("Mtx6").Specific


            For i As Integer = 1 To objMatrix.VisualRowCount
                Dim BarCode As String = objMatrix.Columns.Item("BDesc").Cells.Item(i).Specific.Value.Trim
                Dim ICode As String = objMatrix.Columns.Item("KEYC").Cells.Item(i).Specific.Value.Trim
                Dim RevStr As String = BarCode.Replace(" ", "")
                Dim strlen As String = RevStr.Length.ToString()

                If (BarCode <> "" And ICode <> "") And strlen > 250 Then
                    objMain.objApplication.StatusBar.SetText("For Item " & ICode & " and Line No. " & i & ".... Barcode Description Exceeds more characters!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Next

            For j As Integer = 1 To objMatrixItm.VisualRowCount
                Dim ICode1 As String = objMatrixItm.Columns.Item("KEYC").Cells.Item(j).Specific.Value.Trim
                Dim Carcode1 As String = objMatrixItm.Columns.Item("CARD").Cells.Item(j).Specific.Value.Trim

                For i As Integer = 1 To objMatrixVendor.VisualRowCount
                    'Dim BarCode As String = objMatrix.Columns.Item("BDesc").Cells.Item(i).Specific.Value.Trim
                    Dim ICode As String = objMatrixVendor.Columns.Item("Code").Cells.Item(i).Specific.Value.Trim
                    Dim Carcode As String = objMatrixVendor.Columns.Item("CCode").Cells.Item(i).Specific.Value.Trim

                    If (ICode <> "" And Carcode <> "" And ICode = ICode1) Then
                        If Carcode1 <> Carcode Then
                            objMain.objApplication.StatusBar.SetText("For Item " & ICode & " and Line No. " & i & ".... CardCode is not matching!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                Next
            Next

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    'Application Id
    Public Function AutoDocentryNumber()

        'Dim oRsDocNum As SAPbobsCOM.Recordset = objForm.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objForm = objMain.objApplication.Forms.Item(objForm.UniqueID)
        oDS = objForm.DataSources.DBDataSources.Item("@TNX_ITEMUP")
        Dim oRsDocNum As SAPbobsCOM.Recordset

        oRsDocNum = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim Query1 As String = "SELECT ifnull(Max(""DocNum""),0)+1 AS ""DocNum"" FROM ""@TNX_ITEMUP"" "

        oRsDocNum.DoQuery(Query1)

        oDS.SetValue("DocNum", 0, oRsDocNum.Fields.Item(0).Value)

        Dim rsAppId As SAPbobsCOM.Recordset

        rsAppId = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim str As String = "SELECT 'IMP' || LPAD('',6 - LENGTH(LTRIM(RTRIM(REPLACE(IFNULL(MAX(""U_AppID""), '0'), 'IMP', '') + 1))),'0') || LTRIM(RTRIM(REPLACE(IFNULL(MAX(""U_AppID""), '0'), 'IMP', '') + 1)) FROM ""@TNX_ITEMUP"";"

        rsAppId.DoQuery(str)
        Dim appp As String = rsAppId.Fields.Item(0).Value.ToString().Trim()

        oDS.SetValue("U_AppID", oDS.Offset, appp)


    End Function

    Public Sub CheckDocNumAndGenerateAutoID()

        Try
            Dim docnumcc As String = Me.objForm.Items.Item("AppID").Specific.Value

            Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            '--- Check if DocNum already exists
            Dim checkQry As String =
            "SELECT COUNT(*) FROM ""@TNX_ITEMUP"" WHERE ""U_AppID"" = '" & docnumcc & "'"

            rs.DoQuery(checkQry)

            If rs.Fields.Item(0).Value > 0 Then
                'Naresh
                '--- Generate new Auto App ID Number
                'Dim autoQry As String =
                '"SELECT " &
                '" 'MR' || RIGHT(YEAR(CURRENT_DATE),2) || '-' || " &
                '" LPAD( IFNULL( MAX(CAST(SUBSTRING(""U_ReqNo"",7) AS INTEGER)) + 1, 1 ), 4, '0' ) " &
                '"FROM ""@SBO_MATERIAL"" " &
                '"WHERE ""U_ReqNo"" LIKE 'MR' || RIGHT(YEAR(CURRENT_DATE),2) || '-%'"

                Dim str As String = "SELECT 'IMP' || LPAD('',6 - LENGTH(LTRIM(RTRIM(REPLACE(IFNULL(MAX(""U_AppID""), '0'), 'IMP', '') + 1))),'0') || LTRIM(RTRIM(REPLACE(IFNULL(MAX(""U_AppID""), '0'), 'IMP', '') + 1)) FROM ""@TNX_ITEMUP"";"


                rs.DoQuery(str)

                oDS.SetValue("U_AppID", 0, rs.Fields.Item(0).Value)
                Dim newReqNo As String = rs.Fields.Item(0).Value.ToString()
                objMain.objApplication.StatusBar.SetText("Duplicate detected. New Item Upload Number generated: " & newReqNo, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "Problem while processing Auto ID : " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
    Public Function Validate1() As Boolean

        Try

            Try
                CheckDocNumAndGenerateAutoID()
            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText(ex.Message & "   Errors in Validation Function", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message & "Errors in Validation Function", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)

        Try

            Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(eventInfo.FormUID)
            Dim oMenuItem As SAPbouiCOM.MenuItem
            Dim oMenus As SAPbouiCOM.Menus
            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams

            oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING

            If eventInfo.BeforeAction = True Then

                If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then

                    Dim isMatrix As Boolean = False
                    Dim validMatrixClick As Boolean = False

                    '--------------------------------------
                    ' ALL YOUR MATRICES
                    '--------------------------------------
                    Select Case eventInfo.ItemUID

                        Case "ItmDetMtx", "UOMMtx", "BarMtx", "Mtx5", "Mtx6", "Mtx7"
                            isMatrix = True

                    End Select

                    '--------------------------------------
                    ' COLUMN CHECK (LineId only)
                    '--------------------------------------
                    If isMatrix AndAlso (eventInfo.ColUID = "LineId" OrElse eventInfo.ColUID = "#") Then
                        validMatrixClick = True
                    End If

                    '====================================================
                    ' CASE 1: RIGHT CLICK ON VALID MATRIX ROW (LineId)
                    '====================================================
                    If validMatrixClick Then

                        Try
                            oMenuItem = objMain.objApplication.Menus.Item("1280")
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

                        '====================================================
                        ' CASE 2: NOT A VALID MATRIX CELL
                        ' REMOVE DELETE ROW
                        '====================================================

                        Try
                            oMenuItem = objMain.objApplication.Menus.Item("1280")
                            oMenus = oMenuItem.SubMenus

                            If oMenus.Exists("Delete Row") Then
                                objMain.objApplication.Menus.RemoveEx("Delete Row")
                            End If

                        Catch
                        End Try

                    End If

                    '====================================================
                    ' ALWAYS ENSURE CLEAN STATE FOR OTHER MENU (optional)
                    '====================================================
                    Try
                        oMenuItem = objMain.objApplication.Menus.Item("1280")
                        oMenus = oMenuItem.SubMenus

                        If oMenus.Exists("Close") Then
                            objMain.objApplication.Menus.RemoveEx("Close")
                        End If

                    Catch
                    End Try

                End If
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub
    'Approval

    Sub ApprovalTrigger(ByVal DocType As String, ByVal AppID As String, ByVal Table As String, ByVal AppIDField As String, ByVal AppStatField As String, ByVal Document As String)
        Dim ff As Integer
        Dim draftquery As String
        Dim approverid As String
        Dim docdate As String = objForm.Items.Item("DDATE").Specific.Value

        ff = objMain.objApplication.MessageBox("Do You want Submit the Application?", 1, "Ok", "Cancel")
        If ff = 1 Then
            Dim result As Integer
            result = objMain.objApplication.MessageBox("You can not change the Application after you Sending for Approval,continue", 1, "Yes", "No")
            If result = 1 Then
                Dim rsapp As SAPbobsCOM.Recordset
                Dim posrs As SAPbobsCOM.Recordset
                rsapp = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim str As String = "SELECT T0.""Code"" As ""TemplateID"",T1.""U_Name"" As ""Originator"",S1.""Code"" As ""Stage"",S2.""U_UKey"" As ""Authorizers"" " &
                " ,S1.""U_NAP"" As ""NoOfAuth"",S1.""U_NRJ"" As ""NoOfRej"",T0.""CreateDate"",(Case When T3.""LineId""='1' Then 'S' Else 'O' End) As ""Status"" FROM ""@SBO_APPHDR"" T0 " &
                " INNER JOIN ""@SBO_APPREQ"" T1 ON T1.""Code""=T0.""Code"" " &
                " INNER JOIN ""@SBO_APPDOC"" T2 ON T2.""Code"" = T0.""Code"" " &
                " INNER JOIN ""@SBO_APPAUT"" T3 ON T3.""Code"" = T0.""Code"" " &
                " INNER JOIN ""@SBO_AST"" S1 ON T3.""U_M3_1""=S1.""Code"" " &
                " INNER JOIN ""@SBO_AST_C0"" S2 ON S1.""Code""=S2.""Code"" " &
                " WHERE ""U_Active"" = 'Y' AND T1.""U_Name"" = '" & objMain.objCompany.UserName & "' AND T2.""U_" & DocType & """  = 'Y' AND IFNULL(S2.""U_UKey"",'')<>''  order by T3.""LineId""  ;"
                rsapp.DoQuery(str)


                If rsapp.RecordCount > 0 Then
                    posrs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    'draftquery = " SELECT * FROM ""@SBO_OWDD"" WHERE ""U_statusN"" IN ('S','A') AND ""U_Isdraft"" IN ('Y','N') AND ""U_docnum"" = '" & AppID & "';"
                    draftquery = " SELECT * FROM ""@SBO_DD"" WHERE ""U_statusN"" IN ('S','A') AND ""U_Isdraft"" IN ('Y','N') AND ""U_AppID"" = '" & AppID & "';"

                    posrs.DoQuery(draftquery)
                    If posrs.RecordCount > 0 Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(posrs)
                        objMain.objApplication.StatusBar.SetText("Document Already sent for approval", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    Else
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(posrs)
                        ' Dim ofor As SAPbouiCOM.Form
                        'Me.SendMessageAlert(DocType, objMain.objCompany.UserName)

                        'ofor = New DraftProcedureStages(DocType, AppID, objMain.objCompany.UserName, rsapp, Table, AppIDField, AppStatField, Document, docdate)
                        ' ofor = Nothing
                        Dim objBaseForm As SAPbouiCOM.Form = objMain.objApplication.Forms.ActiveForm
                        objMain.ObjDraftProcedure.CreateForm(objBaseForm, DocType, AppID, objMain.objCompany.UserName, rsapp, Table, AppIDField, AppStatField, Document)
                        '.CreateForm(objForm.UniqueID, objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value)
                        'objMain.objApplication.ActivateMenuItem("1304")
                    End If
                    'objForm.Close()
                Else
                    objMain.objApplication.StatusBar.SetText("Login User Don't have Approvals Please Define Approvals and Add the Document", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        End If
    End Sub


    Public Sub UpdateWarehouseBlockPurchase1(oRs As SAPbobsCOM.Recordset)
        ' Expects oRs to contain: ItemCode, WhsCode, U_BlockPurchase

        Dim oItem As SAPbobsCOM.Items = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

        Try
            oRs.MoveFirst()
            While Not oRs.EoF
                Dim sItemCode As String = oRs.Fields.Item("ItemCode").Value.ToString().Trim()
                Dim BPLID As String = oRs.Fields.Item("BPLID").Value.ToString().Trim()
                Dim sBlockValue As String = oRs.Fields.Item("U_BlockPurchase").Value.ToString().Trim() ' Expected "Y" or "N"
                Dim sWhsCode As String = ""
                ' 1. Load the Item Master
                If oItem.GetByKey(sItemCode) Then
                    Dim bFound As Boolean = False

                    ' 2. Locate the specific Warehouse row for this Item
                    For i As Integer = 0 To oItem.Warehouses.Count - 1
                        oItem.Warehouses.SetCurrentLine(i)

                        If oItem.Warehouses.WarehouseCode = sWhsCode Then
                            ' 3. Update the User-Defined Field on the Warehouse line
                            oItem.Warehouses.UserFields.Fields.Item("U_BlockPurchase").Value = sBlockValue
                            bFound = True
                            Exit For
                        End If
                    Next

                    If bFound Then
                        ' 4. Commit the change
                        If oItem.Update() <> 0 Then
                            Dim errCode As Integer : Dim errMsg As String = ""
                            objMain.objCompany.GetLastError(errCode, errMsg)
                            objMain.objApplication.StatusBar.SetText($"Item {sItemCode} Whs {sWhsCode} Update Failed: {errMsg}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            objMain.objApplication.StatusBar.SetText($"Updated BlockPurchase for {sItemCode} in {sWhsCode}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        End If
                    Else
                        objMain.objApplication.StatusBar.SetText($"Warehouse {sWhsCode} not assigned to Item {sItemCode}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If
                End If

                oRs.MoveNext()
            End While

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Exception: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            ' Cleanup to prevent memory leaks
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem)
        End Try
    End Sub

    '  Public member 'Warehouses' on type 'IItems' not found.
    Public Sub UpdateWarehouseBlockPurchaseold(ByVal docentry As String)
        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim quert As String = ("SELECT T2.""U_ITMCD"" AS ""ItemCode"", T1.""U_BPLID"" AS ""BPLID"", T1.""U_CHECK"" AS ""U_BlockPurchase"" " &
           "FROM ""@TNX_ITEMUP"" T0 INNER JOIN ""@TNX_ITEMUP_C0"" T2 ON T0.""DocEntry""=T2.""DocEntry"" " &
           "INNER JOIN ""@TNX_ITEMUP_C6"" T1 ON T0.""DocEntry""=T1.""DocEntry""  " &
           "WHERE T0.""DocNum""='" & docentry & "' AND IFNULL(T2.""U_ITMCD"", '')<>''")
        oRs.DoQuery(quert)
        Dim oItem As SAPbobsCOM.Items = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        Dim oWhsRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            If oRs.RecordCount = 0 Then Exit Sub

            oRs.MoveFirst()

            While Not oRs.EoF

                Dim sItemCode As String = oRs.Fields.Item("ItemCode").Value.ToString().Trim()
                Dim BPLID As String = oRs.Fields.Item("BPLID").Value.ToString().Trim()
                Dim sBlockValue As String = oRs.Fields.Item("U_BlockPurchase").Value.ToString().Trim()

                ' 🔹 1. Get ACTIVE warehouses for the branch
                Dim sQuery As String = $"SELECT ""WhsCode"" FROM ""OWHS"" 
                                    WHERE ""BPLid"" = '{BPLID}' 
                                    AND IFNULL(""Inactive"", 'N') = 'N' 
                                    AND IFNULL(""Locked"", 'N') = 'N'"

                oWhsRs.DoQuery(sQuery)

                ' 🔹 2. Load Item Master
                If oItem.GetByKey(sItemCode) Then

                    Dim bAnyUpdate As Boolean = False

                    ' 🔹 3. Loop warehouses from branch
                    While Not oWhsRs.EoF

                        Dim sWhsCode As String = oWhsRs.Fields.Item("WhsCode").Value.ToString().Trim()

                        ' 🔹 4. Match with item warehouses (OITW)
                        For i As Integer = 0 To oItem.WhsInfo.Count - 1
                            oItem.WhsInfo.SetCurrentLine(i)

                            If oItem.WhsInfo.WarehouseCode = sWhsCode Then
                                If sBlockValue = "Y" Then
                                    sBlockValue = "Yes"
                                Else
                                    sBlockValue = "No"
                                End If
                                oItem.WhsInfo.UserFields.Fields.Item("U_BlockPurchase").Value = sBlockValue
                                bAnyUpdate = True
                                Exit For
                            End If
                        Next

                        oWhsRs.MoveNext()
                    End While

                    ' 🔹 5. Update ONCE per item
                    If bAnyUpdate Then
                        If oItem.Update() <> 0 Then
                            Dim errCode As Integer
                            Dim errMsg As String = ""
                            objMain.objCompany.GetLastError(errCode, errMsg)

                            objMain.objApplication.StatusBar.SetText(
                            $"❌ Item {sItemCode} Update Failed: {errMsg}",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        Else
                            objMain.objApplication.StatusBar.SetText(
                            $"✅ Updated Active  for Item {sItemCode}",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        End If
                    Else
                        objMain.objApplication.StatusBar.SetText(
                        $"⚠ No matching warehouses for Item {sItemCode}",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If

                Else
                    objMain.objApplication.StatusBar.SetText(
                    $"⚠ Item not found: {sItemCode}",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                End If

                oRs.MoveNext()
            End While

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "❌ Exception: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        Finally
            ' 🔹 Cleanup COM objects
            If oItem IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem)
            End If

            If oWhsRs IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oWhsRs)
            End If
        End Try

    End Sub
    Public Sub UpdateWarehouseBlockPurchase(ByVal docentry As String)
        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oWhsRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oItem As SAPbobsCOM.Items = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

        Try
            ' 1. Get Distinct Items from the document first
            Dim itemQuery As String = $"SELECT DISTINCT T2.""U_ITMCD"" FROM ""@TNX_ITEMUP_C0"" T2 WHERE T2.""DocEntry"" = (SELECT ""DocEntry"" FROM ""@TNX_ITEMUP"" WHERE ""DocNum"" = '{docentry}')"
            oRs.DoQuery(itemQuery)

            If oRs.RecordCount = 0 Then Exit Sub

            While Not oRs.EoF
                Dim sItemCode As String = oRs.Fields.Item("U_ITMCD").Value.ToString().Trim()

                ' 2. For THIS item, find ALL branches where U_CHECK = 'Y'
                Dim branchQuery As String = "SELECT T1.""U_BPLID"" " &
                                      "FROM ""@TNX_ITEMUP"" T0 " &
                                      "INNER JOIN ""@TNX_ITEMUP_C6"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
                                      "WHERE T0.""DocNum"" = '" & docentry & "' AND T1.""U_CHECK"" = 'Y'"

                ' 3. Get all Warehouse Codes belonging to those specific branches
                Dim noWhseList As New HashSet(Of String)
                Dim sqlWhs As String = $"SELECT ""WhsCode"" FROM ""OWHS"" WHERE CAST(""BPLid"" AS VARCHAR) IN ({branchQuery})"
                oWhsRs.DoQuery(sqlWhs)

                While Not oWhsRs.EoF
                    noWhseList.Add(oWhsRs.Fields.Item("WhsCode").Value.ToString().Trim().ToUpper())
                    oWhsRs.MoveNext()
                End While

                ' 4. Load Item and Overwrite ALL lines
                If oItem.GetByKey(sItemCode) Then
                    Dim bChanged As Boolean = False

                    For i As Integer = 0 To oItem.WhsInfo.Count - 1
                        oItem.WhsInfo.SetCurrentLine(i)
                        Dim currentWhs As String = oItem.WhsInfo.WarehouseCode.Trim().ToUpper()

                        ' Logic: If it's in our "No" list, set No. Otherwise, FORCE Yes.
                        Dim finalValue As String = If(noWhseList.Contains(currentWhs), "No", "Yes")

                        ' Set value (this overwrites previous "Yes" or "No" regardless)
                        oItem.WhsInfo.UserFields.Fields.Item("U_BlockPurchase").Value = finalValue
                        bChanged = True
                    Next

                    If bChanged Then
                        If oItem.Update() <> 0 Then
                            Dim errCode As Integer : Dim errMsg As String = ""
                            objMain.objCompany.GetLastError(errCode, errMsg)
                            objMain.objApplication.StatusBar.SetText($"❌ {sItemCode}: {errMsg}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            objMain.objApplication.StatusBar.SetText($"✅ Updated {sItemCode}: {noWhseList.Count} whse(s) set to 'No', rest to 'Yes'", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        End If
                    End If
                End If

                oRs.MoveNext()
            End While

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("❌ Error: " & ex.Message)
        Finally
            If oItem IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem)
            If oWhsRs IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oWhsRs)
            If oRs IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
        End Try
    End Sub
    Public Sub UpdateWarehouseBlockPurchase1(ByVal docentry As String)

        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oWhsRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oItem As SAPbobsCOM.Items = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

        Try
            ' 1. Pre-fetch ALL active warehouses and their Branch IDs into a Dictionary
            ' This avoids running a query inside the While loop.
            Dim whsBranchMap As New Dictionary(Of String, String)
            oWhsRs.DoQuery("SELECT ""WhsCode"", ""BPLid"" FROM ""OWHS"" WHERE IFNULL(""Inactive"", 'N') = 'N' AND IFNULL(""Locked"", 'N') = 'N'")

            While Not oWhsRs.EoF
                Dim wCode As String = oWhsRs.Fields.Item("WhsCode").Value.ToString().Trim().ToUpper()
                Dim bId As String = oWhsRs.Fields.Item("BPLid").Value.ToString().Trim()
                whsBranchMap(wCode) = bId
                oWhsRs.MoveNext()
            End While

            ' 2. Fetch the Item Update data
            Dim query As String = "SELECT T2.""U_ITMCD"" AS ""ItemCode"", T1.""U_BPLID"" AS ""BPLID"", T1.""U_CHECK"" AS ""U_BlockPurchase"" " &
                              "FROM ""@TNX_ITEMUP"" T0 " &
                              "INNER JOIN ""@TNX_ITEMUP_C0"" T2 ON T0.""DocEntry"" = T2.""DocEntry"" " &
                              "INNER JOIN ""@TNX_ITEMUP_C6"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
                              "WHERE T0.""DocNum"" = '" & docentry & "' " &
                              "AND IFNULL(T2.""U_ITMCD"", '') <> ''"

            oRs.DoQuery(query)
            If oRs.RecordCount = 0 Then Exit Sub

            oRs.MoveFirst()
            While Not oRs.EoF
                Dim sItemCode As String = oRs.Fields.Item("ItemCode").Value.ToString().Trim()
                Dim targetBPLID As String = oRs.Fields.Item("BPLID").Value.ToString().Trim()
                Dim sBlockValue As String = oRs.Fields.Item("U_BlockPurchase").Value.ToString().Trim()

                If oItem.GetByKey(sItemCode) Then
                    Dim bAnyUpdate As Boolean = False

                    ' 3. Loop through Item Warehouse Lines
                    For i As Integer = 0 To oItem.WhsInfo.Count - 1
                        oItem.WhsInfo.SetCurrentLine(i)

                        Dim currentWhs As String = oItem.WhsInfo.WarehouseCode.Trim().ToUpper()
                        Dim newValue As String = "Yes" ' Default value

                        ' Logic: If BlockValue is 'Y' AND this warehouse belongs to the target branch, set to 'No'
                        If sBlockValue = "Y" Then
                            If whsBranchMap.ContainsKey(currentWhs) AndAlso whsBranchMap(currentWhs) = targetBPLID Then
                                newValue = "No"
                            End If
                        End If

                        oItem.WhsInfo.UserFields.Fields.Item("U_BlockPurchase").Value = newValue
                        bAnyUpdate = True
                    Next

                    If bAnyUpdate Then
                        If oItem.Update() <> 0 Then
                            Dim errCode As Integer
                            Dim errMsg As String = ""
                            objMain.objCompany.GetLastError(errCode, errMsg)
                            objMain.objApplication.StatusBar.SetText($"❌ {sItemCode}: {errMsg}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            objMain.objApplication.StatusBar.SetText($"✅ Force Updated Item {sItemCode}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        End If
                    End If
                End If

                oRs.MoveNext()
            End While

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("❌ Exception: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            ' Cleanup COM objects
            If oItem IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem)
            If oWhsRs IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oWhsRs)
            If oRs IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
            GC.Collect()
        End Try
    End Sub
    Private Sub DeleteEmptyRows_MatrixWise(ByRef objForm As SAPbouiCOM.Form)

        Try
            objForm.Freeze(True)

            Dim oMatrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("ItmDetMtx").Specific, SAPbouiCOM.Matrix)

            ' Loop from bottom to top
            For i As Integer = oMatrix.VisualRowCount To 1 Step -1

                Dim isEmpty As Boolean = True

                ' Check important columns (adjust as per your need)
                If Trim(CType(oMatrix.Columns.Item("KEYC").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value) <> "" Then isEmpty = False
                If Trim(CType(oMatrix.Columns.Item("ITEMN").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value) <> "" Then isEmpty = False
                If Trim(CType(oMatrix.Columns.Item("CARD").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value) <> "" Then isEmpty = False

                ' If row is empty → delete
                If isEmpty AndAlso oMatrix.VisualRowCount > 1 Then
                    oMatrix.DeleteRow(i)
                End If

            Next

            ' Re-number LineId
            For i As Integer = 1 To oMatrix.VisualRowCount
                CType(oMatrix.Columns.Item("LineId").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value = i.ToString()
            Next

            ' Sync back to datasource AFTER deletion
            oMatrix.FlushToDataSource()

            ' Ensure at least one row
            If oMatrix.VisualRowCount = 0 Then
                oMatrix.AddRow()
            End If

            objForm.Freeze(False)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Error: " & ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
    Public Async Function UpdatePriceForMultipleBranches_S_13L(ByVal DocEntry As String) As Task
        Dim rs As SAPbobsCOM.Recordset = Nothing
        Dim rsUpdate As SAPbobsCOM.Recordset = Nothing

        Try
            ' 1. LOGIN
            Dim sessionId As String = Await CompanySession()
            If String.IsNullOrEmpty(sessionId) Then
                objMain.objApplication.SetStatusBarMessage("SL Login Failed", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Exit Function
            End If

            ' 2. FETCH DATA
            rs = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            rs.DoQuery($"CALL ""SP_GET_ITEM_PRICE_LINE_DTLS""('{DocEntry}')")

            Dim dt As DataTable = ConvertRecordsetToDataTable(rs)
            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                objMain.objApplication.SetStatusBarMessage("No data found to update.", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                Exit Function
            End If

            rsUpdate = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            ' 3. GROUPING
            Dim itemGroups = dt.AsEnumerable().GroupBy(Function(r) r.Field(Of String)("ItemCode")).ToList()
            Dim totalItems As Integer = itemGroups.Count
            Dim currentItem As Integer = 0

            ' 4. PROCESS ITEMS
            For Each itemGroup In itemGroups
                currentItem += 1
                Dim itemCode As String = itemGroup.Key

                ' Update Status Bar Progress
                objMain.objApplication.SetStatusBarMessage($"Processing Item {currentItem} of {totalItems}: {itemCode}...", SAPbouiCOM.BoMessageTime.bmt_Short, False)

                Dim itemPricesList As New List(Of Object)()
                Dim priceListGroups = itemGroup.GroupBy(Function(r) Convert.ToInt32(r("PriceList")))

                For Each plGroup In priceListGroups
                    Dim basePriceRow = plGroup.FirstOrDefault(Function(x) Convert.ToInt32(x("UomEntry")) = Convert.ToInt32(x("IUOM")))

                    Dim uomPricesList = plGroup.
                    Where(Function(x) Convert.ToInt32(x("UomEntry")) <> Convert.ToInt32(x("IUOM"))).
                    Select(Function(x) New With {
                        .UoMEntry = Convert.ToInt32(x("UomEntry")),
                        .Price = Convert.ToDouble(x("Price")),
                        .Auto = "tNO"
                    }).ToList()

                    Dim priceObj As New Dictionary(Of String, Object) From {{"PriceList", plGroup.Key}}

                    If basePriceRow IsNot Nothing Then priceObj("Price") = Convert.ToDouble(basePriceRow("BasePrice"))
                    If uomPricesList.Any() Then priceObj("UoMPrices") = uomPricesList

                    itemPricesList.Add(priceObj)
                Next

                ' 5. SERVICE LAYER CALL
                Dim payload As New Dictionary(Of String, Object) From {{"ItemPrices", itemPricesList}}
                Dim jsonBody As String = JsonConvert.SerializeObject(payload)
                Dim url As String = $"{baseUrl}Items('{itemCode}')"

                Dim response = Await PatchItem(sessionId, url, jsonBody)

                ' 6. DETERMINE STATUS MESSAGE
                Dim statusMsg As String
                Dim isError As Boolean = False

                If response.StatusCode = HttpStatusCode.OK OrElse response.StatusCode = HttpStatusCode.NoContent Then
                    statusMsg = "Price updated Successfully"
                Else
                    statusMsg = response.ResponseBody.Replace("'", " ").Trim()
                    If statusMsg.Length > 250 Then statusMsg = statusMsg.Substring(0, 250)
                    isError = True
                End If

                ' 7. UPDATE UDT & FINAL STATUS
                For Each row In itemGroup
                    rsUpdate.DoQuery($"UPDATE ""@TNX_ITEMUP_C0"" SET ""U_PStst""='{statusMsg}' WHERE ""LineId""='{row("LineId")}' AND ""DocEntry""='{row("DocEntry")}'")
                Next

                If isError Then
                    objMain.objApplication.SetStatusBarMessage($"Error updating {itemCode}: {statusMsg}", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End If
            Next

            objMain.objApplication.SetStatusBarMessage("Price Update Process Completed.", SAPbouiCOM.BoMessageTime.bmt_Short, False)

        Catch ex As Exception
            objMain.objApplication.SetStatusBarMessage($"Critical Error: {ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        Finally
            If rs IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
            If rsUpdate IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(rsUpdate)
            GC.Collect()
        End Try
    End Function

    Public Async Function UpdatePriceForMultipleBranches_SL(ByVal DocEntry As String) As Task

        Dim rs As SAPbobsCOM.Recordset = Nothing
        Dim rsSP As SAPbobsCOM.Recordset = Nothing
        Dim rsUpdate As SAPbobsCOM.Recordset = Nothing

        Try
            ' ===========================
            ' 1. LOGIN TO SERVICE LAYER
            ' ===========================
            Dim sessionId As String = Await CompanySession()
            If String.IsNullOrEmpty(sessionId) Then
                objMain.objApplication.SetStatusBarMessage("SL Login Failed", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Exit Function
            End If

            ' ===========================
            ' 2. GET ITEMS TO PROCESS
            ' ===========================
            rs = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            Dim query As String =
            "SELECT T0.""DocEntry"", T1.""LineId"", T1.""U_ITMCD"",""U_KEYC"" " &
            "FROM ""@TNX_ITEMUP"" T0 " &
            "INNER JOIN ""@TNX_ITEMUP_C0"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
            "WHERE T0.""DocNum"" = '" & DocEntry & "' " &
            "AND IFNULL(T1.""U_ITMCD"", '') <> '' " &
            "AND IFNULL(T1.""U_PStst"", '') <> 'Price updated Successfully'"

            rs.DoQuery(query)

            If rs.RecordCount = 0 Then
                objMain.objApplication.SetStatusBarMessage("No items to process.", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                Exit Function
            End If

            rsUpdate = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            Dim totalItems As Integer = rs.RecordCount
            Dim currentItem As Integer = 0

            ' ===========================
            ' 3. LOOP EACH ITEM
            ' ===========================
            While Not rs.EoF

                currentItem += 1

                Dim itemCode As String = rs.Fields.Item("U_ITMCD").Value.ToString()
                Dim lineId As String = rs.Fields.Item("LineId").Value.ToString()
                Dim docEntryVal As String = rs.Fields.Item("DocEntry").Value.ToString()

                Dim KEYC As String = rs.Fields.Item("U_KEYC").Value.ToString()
                objMain.objApplication.SetStatusBarMessage(
                $"Processing {currentItem}/{totalItems} : {itemCode}",
                SAPbouiCOM.BoMessageTime.bmt_Short, False)

                ' ===========================
                ' 4. CALL STORED PROCEDURE (PER ITEM)
                ' ===========================
                rsSP = DirectCast(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

                rsSP.DoQuery($"CALL ""SP_GET_ITEM_PRICE_LINE_DTLS""('{DocEntry}','{KEYC}')")

                Dim dt As DataTable = ConvertRecordsetToDataTable(rsSP)

                If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                    rs.MoveNext()
                    Continue While
                End If

                ' ===========================
                ' 5. BUILD PRICE LIST PAYLOAD
                ' ===========================

                Dim itemPricesList As New List(Of Object)

                Dim priceListGroups = dt.AsEnumerable().
    GroupBy(Function(r) Convert.ToInt32(r("PriceList")))

                For Each plGroup In priceListGroups

                    Dim currentPriceList As Integer = plGroup.Key

                    ' 🔹 Base Price (Default UOM only)
                    Dim baseRow = plGroup.
        FirstOrDefault(Function(r) Convert.ToInt32(r("UomEntry")) = Convert.ToInt32(r("IUOM")))

                    ' 🔹 UOM Prices (DEDUP + REMOVE BASE UOM)
                    Dim uomPrices = plGroup.
        Where(Function(r) Convert.ToInt32(r("UomEntry")) <> Convert.ToInt32(r("IUOM"))).
        GroupBy(Function(r) Convert.ToInt32(r("UomEntry"))).
        Select(Function(g) g.First()).
        Select(Function(r) New With {
            .UoMEntry = Convert.ToInt32(r("UomEntry")),
            .Price = Convert.ToDouble(r("Price")),
            .Auto = "tNO"
        }).ToList()

                    ' 🔴 FINAL DEDUP SAFETY (VERY IMPORTANT)
                    uomPrices = uomPrices.
        GroupBy(Function(x) x.UoMEntry).
        Select(Function(g) g.First()).
        ToList()

                    ' 🔴 REMOVE INVALID UOM (0 or negative)
                    uomPrices = uomPrices.Where(Function(x) x.UoMEntry > 0).ToList()

                    Dim priceObj As New Dictionary(Of String, Object)
                    priceObj.Add("PriceList", currentPriceList)

                    ' ✅ Add Base Price
                    If baseRow IsNot Nothing Then
                        priceObj("Price") = Convert.ToDouble(baseRow("Price"))
                    End If

                    ' ✅ Add UOM Prices only if valid
                    If uomPrices IsNot Nothing AndAlso uomPrices.Count > 0 Then
                        priceObj("UoMPrices") = uomPrices
                    End If

                    itemPricesList.Add(priceObj)

                Next
                ' ===========================
                ' 6. PATCH TO SERVICE LAYER
                ' ===========================
                Dim payload As New Dictionary(Of String, Object) From {
                {"ItemPrices", itemPricesList}
            }

                Dim jsonBody As String = JsonConvert.SerializeObject(payload)
                Dim url As String = $"{baseUrl}Items('{itemCode}')"

                Dim response = Await PatchItem(sessionId, url, jsonBody)

                ' ===========================
                ' 7. UPDATE STATUS
                ' ===========================
                Dim statusMsg As String

                If response.StatusCode = HttpStatusCode.OK OrElse response.StatusCode = HttpStatusCode.NoContent Then
                    statusMsg = "Price updated Successfully"
                Else
                    statusMsg = response.ResponseBody.Replace("'", " ").Trim()
                End If

                If statusMsg.Length > 250 Then
                    statusMsg = statusMsg.Substring(0, 250)
                End If

                rsUpdate.DoQuery(
                $"UPDATE ""@TNX_ITEMUP_C0"" 
                  SET ""U_PStst"" = '{statusMsg}' 
                  WHERE ""DocEntry"" = '{docEntryVal}' 
                  AND ""LineId"" = '{lineId}'")

                rs.MoveNext()

                ' Release SP recordset each loop (important)
                If rsSP IsNot Nothing Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rsSP)
                    rsSP = Nothing
                End If

            End While

            objMain.objApplication.SetStatusBarMessage("Price Update Completed.", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            objMain.objApplication.ActivateMenuItem("1304")
        Catch ex As Exception
            objMain.objApplication.SetStatusBarMessage($"Error: {ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Medium, True)

        Finally
            If rs IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
            If rsSP IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(rsSP)
            If rsUpdate IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(rsUpdate)
        End Try

    End Function
    ' =========================================
    ' RECORDSET → DATATABLE
    ' =========================================
    Private Function ConvertRecordsetToDataTable(rs As SAPbobsCOM.Recordset) As DataTable

        Dim dt As New DataTable()

        ' CREATE COLUMNS
        For i As Integer = 0 To rs.Fields.Count - 1
            dt.Columns.Add(rs.Fields.Item(i).Name)
        Next

        ' FILL ROWS
        While Not rs.EoF
            Dim dr As DataRow = dt.NewRow()

            For i As Integer = 0 To rs.Fields.Count - 1
                dr(i) = rs.Fields.Item(i).Value
            Next

            dt.Rows.Add(dr)
            rs.MoveNext()
        End While

        Return dt

    End Function



    Private Async Function CompanySession() As Task(Of String)

        Try
            ' ===========================
            ' SSL FIX (MANDATORY)
            ' ===========================
            ServicePointManager.Expect100Continue = True
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11
            ServicePointManager.ServerCertificateValidationCallback = Function() True


            Dim securedwebserviceurl = baseUrl & "Login"

            Dim request = CType(WebRequest.Create(securedwebserviceurl), HttpWebRequest)
            request.Method = "POST"
            request.ContentType = "application/json"

            ' ===========================
            ' REQUEST BODY
            ' ===========================
            Dim json = JsonConvert.SerializeObject(New With {
            .CompanyDB = objMain.objCompany.CompanyDB,
            .UserName = UserName,
            .Password = Password
        })

            Using sw = New IO.StreamWriter(request.GetRequestStream())
                sw.Write(json)
            End Using

            ' ===========================
            ' RESPONSE
            ' ===========================
            request.ProtocolVersion = HttpVersion.Version10
            'Dim httpResponse = CType(request.GetResponse(), HttpWebResponse)
            Dim response = CType(request.GetResponse(), HttpWebResponse)

            Dim cookieHeader As String = response.Headers("Set-Cookie")

            If String.IsNullOrEmpty(cookieHeader) Then
                Throw New Exception("Login failed: No session cookie returned")
            End If

            Dim sessionId As String = cookieHeader.Split(";"c)(0).Split("="c)(1)

            Return sessionId

        Catch ex As WebException

            ' ===========================
            ' HANDLE WEB ERRORS (VERY IMPORTANT)
            ' ===========================
            Dim errorMsg As String = ex.Message

            If ex.Response IsNot Nothing Then
                Using resp = CType(ex.Response, HttpWebResponse)
                    Using sr As New IO.StreamReader(resp.GetResponseStream())
                        errorMsg = sr.ReadToEnd()
                    End Using
                End Using
            End If

            System.IO.File.AppendAllText("C:\Temp\SL_Error.txt",
            "LOGIN WEB ERROR: " & errorMsg & " " & DateTime.Now & Environment.NewLine)

            Return ""

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)



            Return ""

        End Try

    End Function


    Private Async Function PatchItem(sessionId As String, url As String, jsonBody As String) As Task(Of (StatusCode As HttpStatusCode, ResponseBody As String))
        Try
            ' ===========================
            ' SSL FIX (MANDATORY - SAME AS LOGIN)
            ' ===========================
            ServicePointManager.Expect100Continue = True
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls
            ServicePointManager.ServerCertificateValidationCallback = Function() True

            ' Initialize Request
            Dim request = CType(WebRequest.Create(url), HttpWebRequest)
            request.Method = "PATCH" ' Service Layer requires PATCH for updates
            request.ContentType = "application/json"
            request.ProtocolVersion = HttpVersion.Version10

            ' ===========================
            ' SESSION COOKIE
            ' ===========================
            request.Headers.Add("Cookie", "B1SESSION=" & sessionId)

            ' ===========================
            ' WRITE JSON BODY
            ' ===========================
            Dim data = Encoding.UTF8.GetBytes(jsonBody)
            request.ContentLength = data.Length
            Using stream = request.GetRequestStream()
                stream.Write(data, 0, data.Length)
            End Using

            ' ===========================
            ' GET RESPONSE
            ' ===========================
            Using response = CType(Await request.GetResponseAsync(), HttpWebResponse)
                Using reader As New IO.StreamReader(response.GetResponseStream())
                    Dim body = reader.ReadToEnd()
                    Return (response.StatusCode, body)
                End Using
            End Using

        Catch ex As WebException
            ' HANDLE SERVICE LAYER ERRORS
            Dim errorMsg As String = ex.Message
            Dim statusCode As HttpStatusCode = HttpStatusCode.InternalServerError

            If ex.Response IsNot Nothing Then
                Using resp = CType(ex.Response, HttpWebResponse)
                    statusCode = resp.StatusCode
                    Using sr As New IO.StreamReader(resp.GetResponseStream())
                        errorMsg = sr.ReadToEnd()
                    End Using
                End Using
            End If

            Return (statusCode, errorMsg)
        Catch ex As Exception
            Return (HttpStatusCode.InternalServerError, ex.Message)
        End Try
    End Function
    ' =========================================
    ' PATCH
    ' =========================================
    'Private Async Function PatchItem(sessionId As String, url As String, jsonBody As String
    ') As Task(Of (StatusCode As HttpStatusCode, ResponseBody As String))

    '    Try
    '        request.ProtocolVersion = HttpVersion.Version10

    '        Using handler As New HttpClientHandler()
    '            handler.CookieContainer = New CookieContainer()
    '            handler.CookieContainer.Add(New Uri(url), New Cookie("B1SESSION", sessionId))

    '            Using client As New HttpClient(handler)

    '                Dim request As New HttpRequestMessage(New HttpMethod("PATCH"), url)
    '                request.Content = New StringContent(jsonBody, Encoding.UTF8, "application/json")

    '                Dim response = Await client.SendAsync(request)
    '                Dim body = Await response.Content.ReadAsStringAsync()

    '                Return (response.StatusCode, body)

    '            End Using
    '        End Using

    '    Catch ex As Exception
    '        Return (HttpStatusCode.InternalServerError, ex.Message)
    '    End Try

    'End Function

    Private Function ValidateSubAndSubSubGroup(ByVal objForm As SAPbouiCOM.Form) As Boolean

        Try

            Dim oMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("ItmDetMtx").Specific

            For i As Integer = 1 To oMatrix.VisualRowCount

                Dim oSubGrp As SAPbouiCOM.ComboBox = Nothing
                Dim oSubSubGrp As SAPbouiCOM.ComboBox = Nothing

                Try
                    oSubGrp = oMatrix.Columns.Item("SUBGR").Cells.Item(i).Specific
                    oSubSubGrp = oMatrix.Columns.Item("SSGP").Cells.Item(i).Specific
                Catch ex As Exception
                    Continue For
                End Try

                Dim subGrp As String = ""
                Dim subSubGrp As String = ""

                Try
                    subGrp = oSubGrp.Value.Trim()
                Catch ex As Exception
                End Try

                Try
                    subSubGrp = oSubSubGrp.Value.Trim()
                Catch ex As Exception
                End Try

                'Skip empty rows
                If subGrp = "" OrElse subSubGrp = "" Then
                    Continue For
                End If

                Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                Dim strQry As String =
                    "SELECT COUNT(*) ""Cnt"" " &
                    "FROM ""@VR_SUBSUBGROUP"" " &
                    "WHERE ""U_SUBGROUP"" = '" & subGrp.Replace("'", "''") & "' " &
                    "AND ""Name"" = '" & subSubGrp.Replace("'", "''") & "'"

                rs.DoQuery(strQry)

                If CInt(rs.Fields.Item("Cnt").Value) = 0 Then

                    objMain.objApplication.StatusBar.SetText(
                        "Please select correct Sub Sub Group for Row " & i,
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    Return False

                End If

            Next

            Return True

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
                ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            Return False

        End Try

    End Function



End Class


