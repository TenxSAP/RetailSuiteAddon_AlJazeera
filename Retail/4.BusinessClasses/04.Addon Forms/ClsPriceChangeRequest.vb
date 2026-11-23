Imports System.Configuration
Imports System.Data.OleDb
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports Sap.Data.Hana
Imports TaskScheduler
Imports Attachment = System.Net.Mail.Attachment

'test
Public Class ClsPriceChangeRequest
#Region "       Declaration             "
    Public objForm, objNewForm, objFormGrid, objSubForm As SAPbouiCOM.Form
    Dim oGridSub As SAPbouiCOM.Grid
    Dim PARENT_FORM_1 As String
    Dim ChildModalForm As Boolean = False
    Dim oDBs_Head, oDBs_Details, oDBs_Details1, oDBs_Details2, oDBs_Details3, oDBs_Details4, oDBs_Details5, oDBs_Details6, oDBs_Details7, oDBs_Details8, oDBs_Details9 As SAPbouiCOM.DBDataSource
    Dim objMatrix1, objMatrix2, objMatrix3, objMatrix4, objMatrix5, objMatrix6, objMatrix7, objMatrix8, objMatrix9 As SAPbouiCOM.Matrix
    Dim objComboBox, objComboBox1 As SAPbouiCOM.ComboBox
    Dim str, str1 As String
    Public rs, RsNum As SAPbobsCOM.Recordset
    Dim LostFocusFlag As Boolean = False
    Dim oGrid As SAPbouiCOM.Grid
    Dim oDt As SAPbouiCOM.DataTable
    Dim objutilities As Utilities
    Dim matrixid As String = ""
    Dim BaseUom, Sucess, failed As String
    Public WithEvents Timer As New Timers.Timer
    Private WithEvents BackgroundWorker_SalaryJV As New System.ComponentModel.BackgroundWorker
    Private WithEvents BackgroundWorker_SalarySingleJVBranchwise As New System.ComponentModel.BackgroundWorker
    Dim DocNum, Series As String
    Dim oPrg As SAPbouiCOM.ProgressBar
    Public oMessageService As SAPbobsCOM.MessagesService
    Public oCmpSrv As SAPbobsCOM.CompanyService
    Private formUID As String = "PCRT" ' Set this globally when the form is opened
    Private WithEvents BackgroundWorker_CSVLoad As New System.ComponentModel.BackgroundWorker
    Private WithEvents BackgroundWorker_PriceUpdateService As New System.ComponentModel.BackgroundWorker

    Public objDataHelper As dataHelper = New dataHelper
    Public CompanyDB As String = System.Configuration.ConfigurationSettings.AppSettings("companyDB")
    Public UserName As String = System.Configuration.ConfigurationSettings.AppSettings("CompanyUserName")
    Public Password As String = System.Configuration.ConfigurationSettings.AppSettings("companyPassword")
    Public Shared objCompany As SAPbobsCOM.Company = New SAPbobsCOM.Company()
    Public BaseURL As String = System.Configuration.ConfigurationSettings.AppSettings("ServiceLayerBaseUrl")
    Public Property ServiceLayerCon As String
    Public cnfMsg As (FtpStatusCode As HttpStatusCode, ResponseBody As String)
    Public ConStr As String = ""
    Private PriceUpdateCompleted As Boolean = False
    Private dtMatrix3Result As DataTable = Nothing
    Private csvInserted As Integer = 0
    Private csvSkipped As Integer = 0
    Private IsAutoMove As Boolean = False
    Private m_FormUID As String = ""
    Public m_Row As Integer = 0
    Private IsGPCalculated As Boolean = False
    Private DeleteItemUID As String = ""
    Private DeleteMatrix As String = ""
    Private DeleteRow As Integer = 0
    Private m_ColUID As String = ""

#End Region

    Sub CreateForm(ByVal DocNum As String)
        Try

            If objMain IsNot Nothing AndAlso objMain.objCompany IsNot Nothing Then
                ' use the main connected company object instead of a newly constructed one
                ClsPriceChangeRequest.objCompany = objMain.objCompany
            End If

            objMain.objUtilities.LoadForm("PriceChangeRequestNew1.xml", "PCRT", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("PCRT", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            objutilities = New Utilities
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C1")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C2")
            oDBs_Details4 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C3")
            oDBs_Details5 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C4")
            oDBs_Details6 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C5")
            oDBs_Details7 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C6")
            oDBs_Details8 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C7")
            oDBs_Details9 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C8")
            objMatrix1 = objForm.Items.Item("Mtx1").Specific
            objMatrix2 = objForm.Items.Item("Mtx2").Specific
            objMatrix3 = objForm.Items.Item("Mtx3").Specific
            objMatrix4 = objForm.Items.Item("Mtx4").Specific
            objMatrix5 = objForm.Items.Item("Mtx5").Specific
            objMatrix6 = objForm.Items.Item("Mtx6").Specific
            objMatrix7 = objForm.Items.Item("Mtx7").Specific
            objMatrix8 = objForm.Items.Item("Mtx8").Specific
            objMatrix9 = objForm.Items.Item("Mtx9").Specific

            'If DocNum <> "" Then
            '    objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            '    'objForm.Items.Item("DocNum").Specific.Value = DocNum
            '    objForm.Items.Item("AppID").Specific.Value = DocNum
            '    objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            '    objForm.Freeze(False)
            '    Exit Sub
            'End If
            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TNXPCRUDO"))
            oDBs_Head.SetValue("U_TNXDDT", 0, DateTime.Now.ToString("yyyyMMdd"))

            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("TNXDDT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("TNXDDT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("TNXDDT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("TNXDDT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("TNXSTUS").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("TNXSTUS").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("TNXSTUS").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("TNXSTUS").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("Submit").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Submit").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("Item_28").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Item_28").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Item_28").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Item_28").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("PriceLNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("PriceLNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("PriceLNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("PriceLNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)



            'Cost Flucluction Tab UOM
            Dim objmatrix1UOM As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx1").Specific
            Dim CostFluctionMatrixUOM As SAPbouiCOM.Column
            CostFluctionMatrixUOM = objmatrix1UOM.Columns.Item("UOM")
            objutilities.MatrixComboBoxValues(CostFluctionMatrixUOM, "SELECT T0.""UomEntry"", T0.""UomName"" FROM OUOM T0 ORDER BY T0.""UomCode""")
            CostFluctionMatrixUOM.DisplayDesc = True
            CostFluctionMatrixUOM.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            'Flyer price changes Tab UOM
            Dim objmatrix2UOM As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx2").Specific
            Dim FlyerPriceChangesMatrixUOM As SAPbouiCOM.Column
            FlyerPriceChangesMatrixUOM = objmatrix2UOM.Columns.Item("UOM")
            objutilities.MatrixComboBoxValues(FlyerPriceChangesMatrixUOM, "SELECT T0.""UomEntry"", T0.""UomName"" FROM OUOM T0 ORDER BY T0.""UomCode""")
            FlyerPriceChangesMatrixUOM.DisplayDesc = True
            FlyerPriceChangesMatrixUOM.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            'MOP Tab
            Dim objmatrix3UOM As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx3").Specific
            Dim MOPMatrixUOM As SAPbouiCOM.Column
            MOPMatrixUOM = objmatrix3UOM.Columns.Item("UOM")
            objutilities.MatrixComboBoxValues(MOPMatrixUOM, "SELECT T0.""UomEntry"", T0.""UomName"" FROM OUOM T0 ORDER BY T0.""UomCode""")
            MOPMatrixUOM.DisplayDesc = True
            MOPMatrixUOM.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            'Temporay Price Reduction Tab UOM
            Dim objmatrix4UOM As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx4").Specific
            Dim TemporayPriceReductionMatrixUOM As SAPbouiCOM.Column
            TemporayPriceReductionMatrixUOM = objmatrix4UOM.Columns.Item("UOM1")
            objutilities.MatrixComboBoxValues(TemporayPriceReductionMatrixUOM, "SELECT T0.""UomEntry"", T0.""UomName"" FROM OUOM T0 ORDER BY T0.""UomCode""")
            TemporayPriceReductionMatrixUOM.DisplayDesc = True
            TemporayPriceReductionMatrixUOM.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                Try
                    Dim oChk As SAPbouiCOM.CheckBox
                    oChk = CType(objForm.Items.Item("Check").Specific, SAPbouiCOM.CheckBox)

                    ' Default checked
                    oChk.Checked = True

                    ' Disable dates by default
                    objForm.Items.Item("FDTT").Enabled = False
                    objForm.Items.Item("TDTT").Enabled = False
                Catch
                    ' ignore if checkbox missing
                End Try
            End If
            ''objForm.EnableMenu("5890", True)
            objForm.EnableMenu("1292", True)
            objForm.EnableMenu("1293", False)
            objForm.EnableMenu("1296", True)
            objForm.EnableMenu("1295", True)
            objForm.Freeze(False)

            'Enabling Print Layout Icon
            'Try
            '    objForm.ReportType = "TNX_PCRT"

            '    Try
            '        CreateLayoutFolder_TNX_PCRT()
            '    Catch ex As Exception
            '    End Try

            '    Try
            '        RemoveHandler objMain.objApplication.LayoutKeyEvent, AddressOf Me.HandleLayoutKeyEvent
            '    Catch
            '    End Try
            '    Try
            '        AddHandler objMain.objApplication.LayoutKeyEvent, AddressOf Me.HandleLayoutKeyEvent
            '    Catch ex As Exception
            '    End Try
            'Catch ex As Exception
            'End Try
            Me.DisableColumns()
            Me.SetDefault(objForm.UniqueID)
            LoadBranchCombo()
            LoadStatusCombo()
            Me.AutoDocentryNumber()

            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            Try : objForm.Freeze(False) : Catch : End Try
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Sub CreateForm_Find(ByVal DocNumberPCR As String)
        Try

            If objMain IsNot Nothing AndAlso objMain.objCompany IsNot Nothing Then
                ' use the main connected company object instead of a newly constructed one
                ClsPriceChangeRequest.objCompany = objMain.objCompany
            End If

            objMain.objUtilities.LoadForm("PriceChangeRequestNew1.xml", "PCRT", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("PCRT", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            objutilities = New Utilities
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C1")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C2")
            oDBs_Details4 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C3")
            oDBs_Details5 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C4")
            oDBs_Details6 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C5")
            oDBs_Details7 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C6")
            oDBs_Details8 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C7")
            oDBs_Details9 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C8")
            objMatrix1 = objForm.Items.Item("Mtx1").Specific
            objMatrix2 = objForm.Items.Item("Mtx2").Specific
            objMatrix3 = objForm.Items.Item("Mtx3").Specific
            objMatrix4 = objForm.Items.Item("Mtx4").Specific
            objMatrix5 = objForm.Items.Item("Mtx5").Specific
            objMatrix6 = objForm.Items.Item("Mtx6").Specific
            objMatrix7 = objForm.Items.Item("Mtx7").Specific
            objMatrix8 = objForm.Items.Item("Mtx8").Specific
            objMatrix9 = objForm.Items.Item("Mtx9").Specific

            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("TNXDDT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("TNXDDT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("TNXDDT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("TNXDDT").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("TNXSTUS").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("TNXSTUS").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("TNXSTUS").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("TNXSTUS").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("AppID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("Submit").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Submit").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("Item_28").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Item_28").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Item_28").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("Item_28").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("PriceLNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("PriceLNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("PriceLNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("PriceLNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)


            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE

            objForm.Items.Item("AppID").Specific.Value = DocNumberPCR

            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Me.LoadBranchCombofind(DocNumberPCR)

            objForm.Freeze(False)
        Catch ex As Exception
            Try : objForm.Freeze(False) : Catch : End Try
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub CreateFormNEW(ByVal DocEntry As String)
        Try
            objMain.objUtilities.LoadForm("FlyerSort.xml", "ME_FLYER", ResourceType.Embeded)
            objSubForm = objMain.objApplication.Forms.GetForm("ME_FLYER", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objSubForm.Select()
            Dim oGrid As SAPbouiCOM.Grid
            Dim oDT As SAPbouiCOM.DataTable
            oGrid = objSubForm.Items.Item("Item_0").Specific
            oDT = objSubForm.DataSources.DataTables.Item("DT_0")
            Dim sQuery As String = "SELECT T0.""U_BrCode"" ""BarCode"", T0.""U_ItemCode"" ""ItemCode"", T0.""U_ITNM"" ""ItemName"",T1.""FrgnName"" ""Foreign Name"", T0.""U_ITGP"" ""Item Group"",T0.""U_Cost"" ""Cost"", T0.""U_NSNG"" ""Normal Selling"",
(SELECT ROUND(T0.""U_NSNG"" / (1 + IFNULL(VT.""Rate"",0) / 100), 3)
 FROM OITM IT
 INNER JOIN OVTG VT ON IT.""VatGourpSa"" = VT.""Code""
 WHERE IT.""ItemCode"" = T0.""U_ItemCode"") ""Normal Selling Without VAT"",
(SELECT ROUND(T0.""U_NSNG"" * IFNULL(VT.""Rate"",0) / (100 + IFNULL(VT.""Rate"",0)), 3)
   FROM OITM IT 
     INNER JOIN OVTG VT ON IT.""VatGourpSa"" = VT.""Code""
     WHERE IT.""ItemCode"" = T0.""U_ItemCode"") ""VAT Amount"",
T0.""U_FSNG"" ""Flyer Selling"", 
T0.""U_NMGP"" ""Normal gp%"", T0.""U_FLGP"" ""Flyer gp%"", T0.""U_EXPQ"" ""EXP Qty"", T0.""U_ACost"" ""Actual-Cost"", T0.""U_ACSL"" ""Actual /Selling"", T0.""U_AFSL"" ""Actual F/Selling"", 
T0.""U_NGPV"" ""N gp Value"", T0.""U_FGPV"" ""Flyer-gp Value"", T0.""U_PRLS"" ""Profit Loss"", T0.""U_MUpdated"" ""Marked Updated"", T0.""U_UOM"" ""UOM"", T0.""U_PUpdate"" ""Price Updated"", T0.""U_PUS"" ""Price Update Status"", T0.""U_PDN"" ""Promotion Document Number"", T0.""U_PS"" ""Promotion Status"" FROM ""@TNX_PCRT_C1"" T0 LEFT JOIN OITM T1 ON T0.""U_ItemCode""=T1.""ItemCode"" WHERE T0.""DocEntry"" = '" + DocEntry + "' ORDER BY T0.""U_ITGP"",T0.""U_ITNM"" ASC"
            oGrid.DataTable.ExecuteQuery(sQuery)
            For i As Integer = 0 To oGrid.Columns.Count - 1
                oGrid.Columns.Item(i).Editable = False
            Next

            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objSubForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "PCRT" And pVal.BeforeAction = False Then
                Me.CreateForm(DocNum)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                objMatrix1 = objForm.Items.Item("Mtx1").Specific
                objMatrix2 = objForm.Items.Item("Mtx2").Specific
                objMatrix3 = objForm.Items.Item("Mtx3").Specific
                objMatrix4 = objForm.Items.Item("Mtx4").Specific
                objMatrix5 = objForm.Items.Item("Mtx5").Specific
                objMatrix6 = objForm.Items.Item("Mtx6").Specific
                objMatrix7 = objForm.Items.Item("Mtx7").Specific
                objMatrix8 = objForm.Items.Item("Mtx8").Specific
                objMatrix9 = objForm.Items.Item("Mtx9").Specific
                ' Me.SetNewLine(objForm.UniqueID)
                LoadStatusCombo()
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                objMatrix1 = objForm.Items.Item("Mtx1").Specific
                objMatrix2 = objForm.Items.Item("Mtx2").Specific
                objMatrix3 = objForm.Items.Item("Mtx3").Specific
                objMatrix4 = objForm.Items.Item("Mtx4").Specific
                objMatrix5 = objForm.Items.Item("Mtx5").Specific
                objMatrix6 = objForm.Items.Item("Mtx6").Specific
                objMatrix7 = objForm.Items.Item("Mtx7").Specific
                objMatrix8 = objForm.Items.Item("Mtx8").Specific
                objMatrix9 = objForm.Items.Item("Mtx9").Specific
                Me.SetNewLine1(objForm.UniqueID)
                Me.SetNewLine2(objForm.UniqueID)
                Me.SetNewLine3(objForm.UniqueID)
                Me.SetNewLine4(objForm.UniqueID)
                Me.SetNewLine5(objForm.UniqueID)
                Me.SetNewLine6(objForm.UniqueID)
                Me.SetNewLine7(objForm.UniqueID)
                Me.SetNewLine8(objForm.UniqueID)
                Me.SetNewLine9(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1293" AndAlso pVal.BeforeAction = False Then

                Dim oMatrix As SAPbouiCOM.Matrix = Nothing
                Dim oDB As SAPbouiCOM.DBDataSource = Nothing
                Dim BarcodeField As String = ""

                If DeleteMatrix = "Mtx6" Then
                    oMatrix = CType(objForm.Items.Item("Mtx6").Specific, SAPbouiCOM.Matrix)
                    oDB = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C5")
                    BarcodeField = "U_BCode"

                ElseIf DeleteMatrix = "Mtx7" Then
                    oMatrix = CType(objForm.Items.Item("Mtx7").Specific, SAPbouiCOM.Matrix)
                    oDB = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C6")
                    BarcodeField = "U_BCode"
                End If

                If oMatrix IsNot Nothing Then
                    objForm.Freeze(True)
                    oMatrix.FlushToDataSource()

                    For i As Integer = oDB.Size - 1 To 0 Step -1
                        If String.IsNullOrWhiteSpace(oDB.GetValue(BarcodeField, i).Trim()) Then
                            oDB.RemoveRecord(i)
                        End If
                    Next

                    For i As Integer = 0 To oDB.Size - 1
                        oDB.SetValue("LineID", i, (i + 1).ToString())
                    Next

                    oMatrix.LoadFromDataSource()

                    If DeleteMatrix = "Mtx6" Then
                        Me.SetNewLine6(objForm.UniqueID)
                    ElseIf DeleteMatrix = "Mtx7" Then
                        Me.SetNewLine7(objForm.UniqueID)
                    End If
                    objForm.Freeze(False)
                End If

                DeleteMatrix = ""







                'Dim oMatrix As SAPbouiCOM.Matrix = Nothing
                'Dim oDB As SAPbouiCOM.DBDataSource = Nothing

                'If DeleteMatrix = "Mtx6" Then
                '    oMatrix = CType(objForm.Items.Item("Mtx6").Specific, SAPbouiCOM.Matrix)
                '    oDB = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C5")
                'ElseIf DeleteMatrix = "Mtx7" Then
                '    oMatrix = CType(objForm.Items.Item("Mtx7").Specific, SAPbouiCOM.Matrix)
                '    oDB = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C6")
                'End If

                'If oMatrix IsNot Nothing Then

                '    oMatrix.FlushToDataSource()

                '    For i As Integer = 0 To oDB.Size - 1
                '        oDB.SetValue("LineID", i, (i + 1).ToString())
                '    Next

                '    oMatrix.LoadFromDataSource()

                'End If

                'DeleteMatrix = ""



            ElseIf pVal.MenuUID = "519" And pVal.BeforeAction = True Then 'Preview
                Try
                    Dim frs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'PriceChangeRequest' AND ""Type"" = 'C'")
                    If frs.RecordCount = 0 Then
                        objMain.objApplication.MessageBox("Report layout 'PriceChangeRequest' not found. Please upload with same name and Code as parameter", 0, "OK")
                    Else
                        objMain.objApplication.ActivateMenuItem(frs.Fields.Item(0).Value.ToString())
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
                        Dim CrForm As SAPbouiCOM.Form
                        Dim oedt As SAPbouiCOM.EditText
                        CrForm = objMain.objApplication.Forms.ActiveForm
                        oedt = CrForm.Items.Item("1000003").Specific
                        oedt.Value = oDBs_Head.GetValue("DocEntry", 0)
                        CrForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                Catch ex As Exception
                End Try
            ElseIf pVal.MenuUID = "520" And pVal.BeforeAction = True Then 'Print
                Try
                    Dim frs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'PriceChangeRequest' AND ""Type"" = 'C'")
                    If frs.RecordCount = 0 Then
                        objMain.objApplication.MessageBox("Report layout 'PriceChangeRequest' not found. Please upload with same name and Code as parameter", 0, "OK")
                    Else
                        objMain.objApplication.ActivateMenuItem(frs.Fields.Item(0).Value.ToString())
                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
                        Dim CrForm As SAPbouiCOM.Form
                        Dim oedt As SAPbouiCOM.EditText
                        CrForm = objMain.objApplication.Forms.ActiveForm
                        oedt = CrForm.Items.Item("1000003").Specific
                        oedt.Value = oDBs_Head.GetValue("DocEntry", 0)
                        CrForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                Catch ex As Exception
                End Try

            ElseIf pVal.MenuUID = "1281" And pVal.BeforeAction = False Then
                'objForm.Items.Item("AV").Visible = False
                'objForm.Items.Item("1").Visible = True
            ElseIf pVal.MenuUID = "Delete Row" And pVal.BeforeAction = False Then
                ' Delete selected rows from any of Mtx1..Mtx5
                objForm = objMain.objApplication.Forms.Item(objForm.UniqueID)
                ' ensure matrices and datasources are current references
                Try
                    objMatrix1 = objForm.Items.Item("Mtx1").Specific
                Catch ex As Exception
                    objMatrix1 = Nothing
                End Try
                Try
                    objMatrix2 = objForm.Items.Item("Mtx2").Specific
                Catch ex As Exception
                    objMatrix2 = Nothing
                End Try
                Try
                    objMatrix3 = objForm.Items.Item("Mtx3").Specific
                Catch ex As Exception
                    objMatrix3 = Nothing
                End Try
                Try
                    objMatrix4 = objForm.Items.Item("Mtx4").Specific
                Catch ex As Exception
                    objMatrix4 = Nothing
                End Try
                Try
                    objMatrix5 = objForm.Items.Item("Mtx5").Specific
                Catch ex As Exception
                    objMatrix5 = Nothing
                End Try
                Try
                    objMatrix6 = objForm.Items.Item("Mtx6").Specific
                Catch ex As Exception
                    objMatrix6 = Nothing
                End Try
                Try
                    objMatrix7 = objForm.Items.Item("Mtx7").Specific
                Catch ex As Exception
                    objMatrix7 = Nothing
                End Try
                Try
                    objMatrix8 = objForm.Items.Item("Mtx8").Specific
                Catch ex As Exception
                    objMatrix8 = Nothing
                End Try

                Try
                    objMatrix9 = objForm.Items.Item("Mtx9").Specific
                Catch ex As Exception
                    objMatrix9 = Nothing
                End Try

                oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C0")
                oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C1")
                oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C2")
                oDBs_Details4 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C3")
                oDBs_Details5 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C4")
                oDBs_Details6 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C5")
                oDBs_Details7 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C6")
                oDBs_Details8 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C7")
                oDBs_Details9 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C8")

                Dim targetMatrix As SAPbouiCOM.Matrix = Nothing
                Dim targetDS As SAPbouiCOM.DBDataSource = Nothing
                Dim rowdeleted As Boolean = False

                ' helper to find selected matrix
                Dim i As Integer
                If Not objMatrix1 Is Nothing Then
                    For i = objMatrix1.VisualRowCount To 1 Step -1
                        If objMatrix1.IsRowSelected(i) Then
                            targetMatrix = objMatrix1
                            targetDS = oDBs_Details1
                            Exit For
                        End If
                    Next
                End If
                If targetMatrix Is Nothing AndAlso Not objMatrix2 Is Nothing Then
                    For i = objMatrix2.VisualRowCount To 1 Step -1
                        If objMatrix2.IsRowSelected(i) Then
                            targetMatrix = objMatrix2
                            targetDS = oDBs_Details2
                            Exit For
                        End If
                    Next
                End If
                If targetMatrix Is Nothing AndAlso Not objMatrix3 Is Nothing Then
                    For i = objMatrix3.VisualRowCount To 1 Step -1
                        If objMatrix3.IsRowSelected(i) Then
                            targetMatrix = objMatrix3
                            targetDS = oDBs_Details3
                            Exit For
                        End If
                    Next
                End If
                If targetMatrix Is Nothing AndAlso Not objMatrix4 Is Nothing Then
                    For i = objMatrix4.VisualRowCount To 1 Step -1
                        If objMatrix4.IsRowSelected(i) Then
                            targetMatrix = objMatrix4
                            targetDS = oDBs_Details4
                            Exit For
                        End If
                    Next
                End If
                If targetMatrix Is Nothing AndAlso Not objMatrix5 Is Nothing Then
                    For i = objMatrix5.VisualRowCount To 1 Step -1
                        If objMatrix5.IsRowSelected(i) Then
                            targetMatrix = objMatrix5
                            targetDS = oDBs_Details5
                            Exit For
                        End If
                    Next
                End If
                If targetMatrix Is Nothing AndAlso Not objMatrix6 Is Nothing Then
                    For i = objMatrix6.VisualRowCount To 1 Step -1
                        If objMatrix6.IsRowSelected(i) Then
                            targetMatrix = objMatrix6
                            targetDS = oDBs_Details6
                            Exit For
                        End If
                    Next
                End If
                If targetMatrix Is Nothing AndAlso Not objMatrix7 Is Nothing Then
                    For i = objMatrix7.VisualRowCount To 1 Step -1
                        If objMatrix7.IsRowSelected(i) Then
                            targetMatrix = objMatrix7
                            targetDS = oDBs_Details7
                            Exit For
                        End If
                    Next
                End If
                If targetMatrix Is Nothing AndAlso Not objMatrix8 Is Nothing Then
                    For i = objMatrix8.VisualRowCount To 1 Step -1
                        If objMatrix8.IsRowSelected(i) Then
                            targetMatrix = objMatrix8
                            targetDS = oDBs_Details8
                            Exit For
                        End If
                    Next
                End If

                If targetMatrix Is Nothing AndAlso Not objMatrix9 Is Nothing Then
                    For i = objMatrix9.VisualRowCount To 1 Step -1
                        If objMatrix9.IsRowSelected(i) Then
                            targetMatrix = objMatrix9
                            targetDS = oDBs_Details9
                            Exit For
                        End If
                    Next
                End If

                If targetMatrix Is Nothing Then
                    objMain.objApplication.MessageBox("please select a row to delete.")
                    Return
                End If

                ' Perform deletion
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    For i = targetMatrix.VisualRowCount To 1 Step -1
                        If targetMatrix.IsRowSelected(i) Then
                            targetMatrix.DeleteRow(i)
                            rowdeleted = True
                        End If
                    Next
                    If rowdeleted Then
                        targetMatrix.FlushToDataSource()
                    End If
                ElseIf objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE OrElse objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    ' ensure form is in update mode
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If

                    targetMatrix.FlushToDataSource()
                    For i = targetMatrix.VisualRowCount To 1 Step -1
                        If targetMatrix.IsRowSelected(i) Then
                            Dim dsindex As Integer = i - 1
                            If Not targetDS Is Nothing AndAlso dsindex < targetDS.Size Then
                                targetDS.RemoveRecord(dsindex)
                                rowdeleted = True
                            End If
                        End If
                    Next
                    If rowdeleted Then
                        targetMatrix.LoadFromDataSource()
                    End If
                End If

                If Not rowdeleted Then
                    objMain.objApplication.MessageBox("please select a row to delete.")
                    Return
                End If

                ' Re-number LineID column
                Try
                    For i = 1 To targetMatrix.VisualRowCount
                        Try
                            targetMatrix.Columns.Item("LineID").Cells.Item(i).Specific.value = i.ToString()
                        Catch ex As Exception
                            ' ignore if column not present
                        End Try
                    Next
                Catch ex As Exception
                    ' ignore
                End Try

            ElseIf pVal.MenuUID = "1284" And pVal.BeforeAction = True Then
                Dim objForm As SAPbouiCOM.Form
                objForm = objMain.objApplication.Forms.ActiveForm

                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                    '====================================================
                    ' Check Cancellation Remark Mandatory
                    '====================================================
                    Dim Remark As String = ""

                    Try
                        Remark = Trim(oDBs_Head.GetValue("U_Remark", 0).ToString())
                    Catch
                        Remark = ""
                    End Try

                    If String.IsNullOrEmpty(Remark) Then

                        objMain.objApplication.StatusBar.SetText(
                "Cancellation remarks are mandatory. Please enter the remarks before cancelling the document.",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        BubbleEvent = False
                        Exit Sub

                    End If
                    Dim Status As String
                    Status = Trim(oDBs_Head.GetValue("U_TNXSTUS", 0))

                    If Status = "N" Or Status = "C" Then

                        objMain.objApplication.StatusBar.SetText("Document cannot be cancelled", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        BubbleEvent = False
                        Exit Sub

                    End If
                    Dim DocEntry As String
                    Dim Appication As String
                    DocEntry = oDBs_Head.GetValue("DocEntry", 0)
                    Appication = oDBs_Head.GetValue("DocEntry", 0)

                    'Update U_TNXSTUS field in database
                    Dim oRS As SAPbobsCOM.Recordset
                    oRS = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    oRS.DoQuery("UPDATE ""@TNX_PCRT"" SET ""U_TNXSTUS"" = 'N' WHERE ""DocEntry"" = " & DocEntry & "")
                    'Update OWDD Status
                    oRS.DoQuery("UPDATE ""@SBO_DD"" SET ""U_StatusN"" = 'N' WHERE ""U_docnum"" = " & Appication & " AND ""U_StatusN"" IN ('A','S')")

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS)
                End If
            End If

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)

        Try

            Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(eventInfo.FormUID)

            If eventInfo.BeforeAction Then

                DeleteMatrix = eventInfo.ItemUID

                Try
                    If eventInfo.ItemUID.StartsWith("Mtx") Then
                        objForm.EnableMenu("784", True)
                    Else
                        objForm.EnableMenu("784", False)
                    End If
                Catch ex As Exception
                End Try

                If eventInfo.ItemUID = "Mtx6" OrElse eventInfo.ItemUID = "Mtx7" Then

                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        objForm.EnableMenu("1293", True)
                    Else
                        objForm.EnableMenu("1293", False)
                    End If

                Else

                    objForm.EnableMenu("1293", False)

                End If

            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub
    Public Function Validate() As Boolean
        ' Dim CustomerCode As SAPbouiCOM.Matrix
        Try


            If oDBs_Head.GetValue("U_TNXPUTP", 0) = "MOP" Then
                If oDBs_Head.GetValue("U_PName", 0) = "" Then
                    objMain.objApplication.SetStatusBarMessage("Promotion Name is mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                    Return False
                    Exit Function
                End If
                If oDBs_Head.GetValue("U_FDATEM1", 0) = "" Then
                    objMain.objApplication.SetStatusBarMessage("From Date is mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                    Return False
                    Exit Function
                End If

                If oDBs_Head.GetValue("U_TDATEMP", 0) = "" Then
                    objMain.objApplication.SetStatusBarMessage("To Date is mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                    Return False
                    Exit Function
                End If
                If oDBs_Head.GetValue("U_FTIMEM", 0) = "" Then
                    objMain.objApplication.SetStatusBarMessage("From Time is mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                    Return False
                    Exit Function
                End If

                If oDBs_Head.GetValue("U_TTIMEM", 0) = "" Then
                    objMain.objApplication.SetStatusBarMessage("To Time is mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                    Return False
                    Exit Function
                End If
                If objMatrix3.Columns.Item("BCode").Cells.Item(1).Specific.value = "" Then
                    objMain.objApplication.SetStatusBarMessage("Please enter at least one barcode at the line level.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                    Return False
                    Exit Function
                End If

                For j As Integer = 1 To objMatrix3.VisualRowCount
                    Dim Line As String = objMatrix3.Columns.Item("LineID").Cells.Item(j).Specific.value
                    If objMatrix3.Columns.Item("ItemCode").Cells.Item(j).Specific.value <> "" Then
                        If objMatrix3.Columns.Item("CName").Cells.Item(j).Specific.value = "" Then
                            objMain.objApplication.SetStatusBarMessage("Please enter Competitor Name for line number " & Line, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                            Exit Function
                        End If

                        If objMatrix3.Columns.Item("IAttach").Cells.Item(1).Specific.value = "" Then
                            objMain.objApplication.SetStatusBarMessage("Invoice Attachment is mandatory for line number " & Line, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                            Exit Function
                        End If

                    End If


                Next

                Dim BarcodeList As New List(Of String)

                For i As Integer = 1 To objMatrix3.VisualRowCount

                    Dim Barcode As String = objMatrix3.Columns.Item("BCode").Cells.Item(i).Specific.Value.ToString().Trim()

                    'Skip empty rows
                    If Barcode <> "" Then

                        'Check duplicate barcode
                        If BarcodeList.Contains(Barcode) Then

                            objMain.objApplication.SetStatusBarMessage("Duplicate barcode found at Line " & i & ": " & Barcode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                            Return False
                        End If

                        BarcodeList.Add(Barcode)

                    End If

                Next


            End If

            If oDBs_Head.GetValue("U_TNXPUTP", 0) = "Bulk Price Update" Then
                For i As Integer = 1 To objMatrix9.VisualRowCount

                    Dim ItemCode As String = objMatrix9.Columns.Item("Col_5").Cells.Item(i).Specific.Value.ToString().Trim()

                    If ItemCode <> "" Then


                        If objMatrix9.Columns.Item("Col_3").Cells.Item(i).Specific.Value.ToString().Trim() = 0 Then
                            objMain.objApplication.SetStatusBarMessage("New Selling Price is mandatory at line " & i & ".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                            Exit Function
                        End If



                    End If

                Next
            End If

            If oDBs_Head.GetValue("U_TNXPUTP", 0) = "Bulk Cost Update" Then
                If oDBs_Head.GetValue("U_FDATEB", 0) = "" Then
                    objMain.objApplication.SetStatusBarMessage("From Date is mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                    Return False
                    Exit Function
                End If

                'If objMatrix3.Columns.Item("BCode").Cells.Item(1).Specific.value = "" Then
                '    objMain.objApplication.SetStatusBarMessage("Please enter at least one barcode at the line level.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                '    Return False
                '    Exit Function
                'End If

                For i As Integer = 1 To objMatrix7.VisualRowCount

                    Dim ItemCode As String = objMatrix7.Columns.Item("ITCD").Cells.Item(i).Specific.Value.ToString().Trim()

                    If ItemCode <> "" Then

                        Dim rsSugarTax As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        rsSugarTax.DoQuery("SELECT IFNULL(""U_SugarTax"",'N') FROM OITM WHERE ""ItemCode"" = '" & ItemCode.Replace("'", "''") & "'")

                        If rsSugarTax.RecordCount > 0 AndAlso rsSugarTax.Fields.Item(0).Value.ToString().Trim() = "Y" Then

                            If objMatrix7.Columns.Item("Base").Cells.Item(i).Specific.Value.ToString().Trim() = "" OrElse CDbl(objMatrix7.Columns.Item("Base").Cells.Item(i).Specific.Value) <= 0 Then
                                objMain.objApplication.SetStatusBarMessage("Base is mandatory for Sugar Tax item at line " & i & ".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                                Exit Function
                            End If

                            If objMatrix7.Columns.Item("Tax").Cells.Item(i).Specific.Value.ToString().Trim() = "" OrElse CDbl(objMatrix7.Columns.Item("Tax").Cells.Item(i).Specific.Value) <= 0 Then
                                objMain.objApplication.SetStatusBarMessage("Tax is mandatory for Sugar Tax item at line " & i & ".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                                Exit Function
                            End If

                        End If

                    End If

                Next
            End If

            If oDBs_Head.GetValue("U_TNXPUTP", 0) = "Flyer Price Changes" Then
                'If oDBs_Head.GetValue("U_PName", 0) = "" Then
                '    objMain.objApplication.SetStatusBarMessage("Promotion Name is mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                '    Return False
                '    Exit Function
                'End If
                Dim sUserCode As String = objMain.objCompany.UserName
                Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                Dim strSQL As String = "SELECT DISTINCT S2.""U_AUTH"" " &
                       "FROM ""@SBO_APPHDR"" T0 " &
                       "INNER JOIN ""@SBO_APPREQ"" T1 ON T1.""Code""=T0.""Code"" " &
                       "INNER JOIN ""@SBO_APPDOC"" T2 ON T2.""Code""=T0.""Code"" " &
                       "INNER JOIN ""@SBO_APPAUT"" T3 ON T3.""Code""=T0.""Code"" " &
                       "INNER JOIN ""@SBO_AST"" S1 ON T3.""U_M3_1""=S1.""Code"" " &
                       "INNER JOIN ""@SBO_AST_C0"" S2 ON S1.""Code""=S2.""Code"" " &
                       "WHERE ""U_Active"" = 'Y' " &
                       "AND T0.""U_PCR"" = 'Flyer Price Changes' " &
                       "AND S2.""U_AUTH"" = '" & sUserCode.Replace("'", "''") & "'"

                rs.DoQuery(strSQL)

                If rs.RecordCount > 0 Then

                    'Current user is an Authorizer
                    If String.IsNullOrWhiteSpace(oDBs_Head.GetValue("U_PName", 0).Trim()) Then

                        objMain.objApplication.SetStatusBarMessage("Promotion Name is mandatory.", SAPbouiCOM.BoMessageTime.bmt_Short, True)

                        Return False

                    End If

                End If

                If oDBs_Head.GetValue("U_FDATEF", 0) = "" Then
                    objMain.objApplication.SetStatusBarMessage("From Date is mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                    Return False
                    Exit Function
                End If

                If oDBs_Head.GetValue("U_TDATEF", 0) = "" Then
                    objMain.objApplication.SetStatusBarMessage("To Date is mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                    Return False
                    Exit Function
                End If
                If oDBs_Head.GetValue("U_FTIMEF", 0) = "" Then
                    objMain.objApplication.SetStatusBarMessage("From Time is mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                    Return False
                    Exit Function
                End If

                If oDBs_Head.GetValue("U_TTIMEF", 0) = "" Then
                    objMain.objApplication.SetStatusBarMessage("To Time is mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                    Return False
                    Exit Function
                End If

                If objMatrix2.Columns.Item("BrCode").Cells.Item(1).Specific.value = "" Then
                    objMain.objApplication.SetStatusBarMessage("Please enter at least one barcode at the line level.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                    Return False
                    Exit Function
                End If

                Dim BarcodeList As New List(Of String)

                For i As Integer = 1 To objMatrix2.VisualRowCount

                    Dim Barcode As String = objMatrix2.Columns.Item("BrCode").Cells.Item(i).Specific.Value.ToString().Trim()

                    'Skip empty rows
                    If Barcode <> "" Then

                        'Check duplicate barcode
                        If BarcodeList.Contains(Barcode) Then

                            objMain.objApplication.SetStatusBarMessage("Duplicate barcode found at Line " & i & ": " & Barcode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                            Return False
                        End If

                        BarcodeList.Add(Barcode)

                    End If

                Next

            End If

            If oDBs_Head.GetValue("U_TNXPUTP", 0) = "Temporary Price Reduction" Then
                If oDBs_Head.GetValue("U_TPRT", 0) = "" Then
                    objMain.objApplication.SetStatusBarMessage("TPR Type is mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                    Return False
                    Exit Function
                End If
                If oDBs_Head.GetValue("U_PName", 0) = "" Then
                    objMain.objApplication.SetStatusBarMessage("Promotion Name is mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                    Return False
                    Exit Function
                End If
                If oDBs_Head.GetValue("U_FDATET", 0) = "" Then
                    objMain.objApplication.SetStatusBarMessage("From Date is mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                    Return False
                    Exit Function
                End If

                If oDBs_Head.GetValue("U_TDATET", 0) = "" Then
                    objMain.objApplication.SetStatusBarMessage("To Date is mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                    Return False
                    Exit Function
                End If
                If oDBs_Head.GetValue("U_FTIMET", 0) = "" Then
                    objMain.objApplication.SetStatusBarMessage("From Time is mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                    Return False
                    Exit Function
                End If

                If oDBs_Head.GetValue("U_TTIMET", 0) = "" Then
                    objMain.objApplication.SetStatusBarMessage("To Time is mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                    Return False
                    Exit Function
                End If
                If objMatrix4.Columns.Item("BrCode").Cells.Item(1).Specific.value = "" Then
                    objMain.objApplication.SetStatusBarMessage("Please enter at least one barcode at the line level.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                    Return False
                    Exit Function
                End If

                Dim BarcodeList As New List(Of String)

                For i As Integer = 1 To objMatrix4.VisualRowCount

                    Dim Barcode As String = objMatrix4.Columns.Item("BrCode").Cells.Item(i).Specific.Value.ToString().Trim()

                    'Skip empty rows
                    If Barcode <> "" Then

                        'Check duplicate barcode
                        If BarcodeList.Contains(Barcode) Then

                            objMain.objApplication.SetStatusBarMessage("Duplicate barcode found at Line " & i & ": " & Barcode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                            Return False
                        End If

                        BarcodeList.Add(Barcode)

                    End If

                Next

            End If

            If oDBs_Head.GetValue("U_TNXPUTP", 0) = "Every day Special Price" Then
                If objMatrix6.Columns.Item("Col_0").Cells.Item(1).Specific.value = "" Then
                    objMain.objApplication.SetStatusBarMessage("Please enter at least one barcode at the line level.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                    Return False
                    Exit Function
                End If

                For j As Integer = 1 To objMatrix6.VisualRowCount
                    Dim Line As String = objMatrix6.Columns.Item("LineID").Cells.Item(j).Specific.value
                    If objMatrix6.Columns.Item("Col_1").Cells.Item(j).Specific.value <> "" Then
                        If objMatrix6.Columns.Item("Col_7").Cells.Item(j).Specific.value = 0 Then
                            objMain.objApplication.SetStatusBarMessage("Please enter New Cost for line number " & Line, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                            Exit Function
                        End If

                        If objMatrix6.Columns.Item("Col_8").Cells.Item(j).Specific.value = 0 Then
                            objMain.objApplication.SetStatusBarMessage("Please enter New selling for line number " & Line, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                            Exit Function
                        End If

                        If objMatrix6.Columns.Item("IAttach").Cells.Item(1).Specific.value = "" Then
                            objMain.objApplication.SetStatusBarMessage("Invoice Attachment is Mandatory for line number " & Line, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                            Exit Function
                        End If

                    End If

                Next
            End If

            If oDBs_Head.GetValue("U_TNXPUTP", 0) = "Temporary Price Reduction" Then

                Dim sup As String = oDBs_Head.GetValue("U_TPRT", 0)
                For j As Integer = 1 To objMatrix4.VisualRowCount
                    '---Madhu 12/07/2026 
                    ' If sup = "Support" Or sup = "Sell out" Then
                    Dim Line As String = objMatrix4.Columns.Item("LineID").Cells.Item(j).Specific.value
                    If sup = "Sell out" Then
                        If objMatrix4.Columns.Item("ItemCode").Cells.Item(j).Specific.value <> "" And objMatrix4.Columns.Item("Sell").Cells.Item(j).Specific.value = 0 Then
                            objMain.objApplication.SetStatusBarMessage("Sell Out Amount is Mandatory" & Line, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                            Exit Function
                        End If
                    End If
                    If sup = "Support" Then
                        If objMatrix4.Columns.Item("ItemCode").Cells.Item(1).Specific.value <> "" And objMatrix4.Columns.Item("Col_0").Cells.Item(1).Specific.value = "" Then
                            objMain.objApplication.SetStatusBarMessage("Please select CreditNote no", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                            Exit Function
                        End If
                    End If

                    If objMatrix4.Columns.Item("Qty").Cells.Item(j).Specific.value = 0 Then
                        objMain.objApplication.SetStatusBarMessage("TPR Quantity is mandatory" & Line, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                        Exit Function
                    End If

                    If objMatrix4.Columns.Item("SPrice").Cells.Item(j).Specific.value = 0 Then
                        objMain.objApplication.SetStatusBarMessage("Special Price is mandatory" & Line, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                        Exit Function
                    End If

                Next
            End If

            '---12/07/2026  Madhu Removed validation comment as per client now 
            'If oDBs_Head.GetValue("U_TNXPUTP", 0) = "Flyer Price Changes" Then


            '    For j As Integer = 1 To objMatrix2.VisualRowCount

            '        If objMatrix2.Columns.Item("ItemCode").Cells.Item(j).Specific.value <> "" And objMatrix2.Columns.Item("Col_13").Cells.Item(j).Specific.value = "" Then
            '            objMain.objApplication.SetStatusBarMessage("Please select CreditNote no", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '            Return False
            '            Exit Function
            '        End If


            '    Next
            'End If

            If oDBs_Head.GetValue("U_TNXPUTP", 0) = "Flyer Price Changes" Then


                For j As Integer = 1 To objMatrix2.VisualRowCount

                    If objMatrix2.Columns.Item("ItemCode").Cells.Item(j).Specific.value <> "" And objMatrix2.Columns.Item("Col_1").Cells.Item(j).Specific.value = 0 Then
                        objMain.objApplication.SetStatusBarMessage("Flyer Selling price is mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                        Exit Function
                    End If


                Next
            End If

            If oDBs_Head.GetValue("U_TNXPUTP", 0) = "Cost Fluctuation" Then


                For j As Integer = 1 To objMatrix1.VisualRowCount
                    Dim Item As String = objMatrix1.Columns.Item("TNXICD").Cells.Item(j).Specific.value
                    Dim Line As String = objMatrix1.Columns.Item("LineID").Cells.Item(j).Specific.value
                    If Item <> "" Then
                        If objMatrix1.Columns.Item("Col_2").Cells.Item(j).Specific.value = "" Then
                            objMain.objApplication.SetStatusBarMessage("Please select Special Price Type" & Line, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                            Exit Function
                        End If

                        If objMatrix1.Columns.Item("Col_2").Cells.Item(j).Specific.value = "Scheduled Update" Then
                            If objMatrix1.Columns.Item("Col_3").Cells.Item(j).Specific.value = "" Then
                                objMain.objApplication.SetStatusBarMessage("From Date is manditory" & Line, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                                Exit Function
                            End If
                            If objMatrix1.Columns.Item("Col_4").Cells.Item(j).Specific.value = "" Then
                                objMain.objApplication.SetStatusBarMessage("From time is manditory" & Line, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                                Exit Function
                            End If
                        End If
                        If objMatrix1.Columns.Item("NRP").Cells.Item(j).Specific.value = 0 Then
                            objMain.objApplication.SetStatusBarMessage("New Special Price is manditory for Line No. " & Line, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                            Exit Function
                        End If

                    End If

                Next

                If objMatrix1.Columns.Item("Col_5").Cells.Item(1).Specific.value = "" Then
                    objMain.objApplication.SetStatusBarMessage("Please enter at least one barcode at the line level.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                    Return False
                    Exit Function
                End If

            End If

            'If oDBs_Head.GetValue("U_PriceLNum", 0) = "" Then
            '    objMain.objApplication.SetStatusBarMessage("Price List Number Is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
            '    Return False
            '    Exit Function
            'End If

            'Dim itemCode As String = objMatrix1.Columns.Item("TNXICD").Cells.Item(1).Specific.Value
            'If itemCode = "" Then
            '    objMain.objApplication.SetStatusBarMessage("ItemCode Is Missing in Cost Fluctuation in Tab-1 ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If

            'Dim BarCode As String = objMatrix2.Columns.Item("BrCode").Cells.Item(1).Specific.Value
            'If BarCode = "" Then
            '    objMain.objApplication.SetStatusBarMessage("BarCode Is Missing ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If

            'Dim BarCode2 As String = objMatrix3.Columns.Item("BCode").Cells.Item(1).Specific.Value
            'If BarCode2 = "" Then
            '    objMain.objApplication.SetStatusBarMessage("MOP BarCode Is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If

            'Dim BarCode3 As String = objMatrix4.Columns.Item("BrCode").Cells.Item(1).Specific.Value
            'If BarCode3 = "" Then
            '    objMain.objApplication.SetStatusBarMessage("TPR BarCode Is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message & "Errors in Validation Function", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function
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
    Private Function GetNewDataFromMatrixFlyer() As DataTable

        Dim dt As New DataTable()

        dt.Columns.Add("ItemCode")
        dt.Columns.Add("Bcode")

        Try

            objMatrix2 = objForm.Items.Item("Mtx2").Specific

            'Get all current rows from Matrix
            Dim visualRows As Integer = objMatrix2.VisualRowCount

            For r As Integer = 1 To visualRows

                Dim itemCode As String = ""
                Dim barcode As String = ""

                'Get Item Code
                Try
                    itemCode = objMatrix2.Columns.Item("ItemCode").
                    Cells.Item(r).Specific.Value.ToString().Trim()
                Catch
                    itemCode = ""
                End Try

                'Get Barcode
                Try
                    barcode = objMatrix2.Columns.Item("BrCode").
                    Cells.Item(r).Specific.Value.ToString().Trim()
                Catch
                    barcode = ""
                End Try

                'Skip completely empty rows
                If itemCode = "" AndAlso barcode = "" Then
                    Continue For
                End If

                'Add ALL current matrix rows
                dt.Rows.Add(itemCode, barcode)

            Next

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
            "GetNewDataFromMatrixFlyer error: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

        Return dt

    End Function
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Try
                objForm = objMain.objApplication.Forms.Item(FormUID)
                If pVal.BeforeAction = False AndAlso objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    ApplyTabsForCurrentType(objForm)
                End If
            Catch
                ' ignore
            End Try

            Select Case pVal.EventType

                '

                '
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    'objForm.Freeze(True)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C0")
                    objMatrix1 = objForm.Items.Item("Mtx1").Specific
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal        'CHOOSE FORM LIST'
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                    'If oCFL.UniqueID = "CFL_1" And pVal.BeforeAction = True Then 'CHOOSE FORM TABLE UNIQUNE ID
                    '    CFLFilter(FormUID, "CFL_1")
                    'End If
                    'submit
                    If pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = True And pVal.ActionSuccess = False Then


                        If oCFL.UniqueID = "CFL_2" Then

                            m_FormUID = FormUID
                            m_Row = pVal.Row
                            Dim Branch As String = ""
                            Dim BrandID As String = "Select ""BPLId"" from OBPL where ""BPLName"" = '" + oDBs_Head.GetValue("U_TNXBRC", 0).Trim() + "'"
                            Dim oRsB As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsB.DoQuery(BrandID)
                            Branch = oRsB.Fields.Item("BPLId").Value.ToString().Trim()
                            If Branch = 0 Then
                                Branch = oDBs_Head.GetValue("U_TNXBRC", 0).Trim()
                            End If

                            Dim objBarcodeFilter As New ClsBarcodeFilter

                            objBarcodeFilter.MainFormUID = objForm.UniqueID


                            objMain.objClsBarcodeFilter.BarCodeFilter(FormUID, Branch)

                            BubbleEvent = False

                        End If
                        If oCFL.UniqueID = "CFL_3" Then

                            m_FormUID = FormUID
                            m_Row = pVal.Row
                            Dim Branch As String = ""
                            Dim BrandID As String = "Select ""BPLId"" from OBPL where ""BPLName"" = '" + oDBs_Head.GetValue("U_TNXBRC", 0).Trim() + "'"
                            Dim oRsB As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsB.DoQuery(BrandID)
                            Branch = oRsB.Fields.Item("BPLId").Value.ToString().Trim()
                            If Branch = 0 Then
                                Branch = oDBs_Head.GetValue("U_TNXBRC", 0).Trim()
                            End If

                            Dim objBarcodeFilter As New ClsBarcodeFilter

                            objBarcodeFilter.MainFormUID = objForm.UniqueID


                            objMain.objClsBarcodeFilter.BarCodeFilter(FormUID, Branch)

                            BubbleEvent = False

                        End If
                        If oCFL.UniqueID = "CFL_4" Then

                            m_FormUID = FormUID
                            m_Row = pVal.Row
                            Dim Branch As String = ""
                            Dim BrandID As String = "Select ""BPLId"" from OBPL where ""BPLName"" = '" + oDBs_Head.GetValue("U_TNXBRC", 0).Trim() + "'"
                            Dim oRsB As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsB.DoQuery(BrandID)
                            Branch = oRsB.Fields.Item("BPLId").Value.ToString().Trim()
                            If Branch = 0 Then
                                Branch = oDBs_Head.GetValue("U_TNXBRC", 0).Trim()
                            End If

                            Dim objBarcodeFilter As New ClsBarcodeFilter

                            objBarcodeFilter.MainFormUID = objForm.UniqueID


                            objMain.objClsBarcodeFilter.BarCodeFilter(FormUID, Branch)

                            BubbleEvent = False

                        End If
                        If oCFL.UniqueID = "CFL_6" Then

                            m_FormUID = FormUID
                            m_Row = pVal.Row
                            Dim Branch As String = ""
                            Dim BrandID As String = "Select ""BPLId"" from OBPL where ""BPLName"" = '" + oDBs_Head.GetValue("U_TNXBRC", 0).Trim() + "'"
                            Dim oRsB As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsB.DoQuery(BrandID)
                            Branch = oRsB.Fields.Item("BPLId").Value.ToString().Trim()
                            If Branch = 0 Then
                                Branch = oDBs_Head.GetValue("U_TNXBRC", 0).Trim()
                            End If

                            Dim objBarcodeFilter As New ClsBarcodeFilter

                            objBarcodeFilter.MainFormUID = objForm.UniqueID


                            objMain.objClsBarcodeFilter.BarCodeFilter(FormUID, Branch)

                            BubbleEvent = False

                        End If
                        'If oCFL.UniqueID = "CFL_9" Then

                        '    m_FormUID = FormUID
                        '    m_Row = pVal.Row
                        '    Dim Branch As String = ""
                        '    Dim BrandID As String = "Select ""BPLId"" from OBPL where ""BPLName"" = '" + oDBs_Head.GetValue("U_TNXBRC", 0).Trim() + "'"
                        '    Dim oRsB As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '    oRsB.DoQuery(BrandID)
                        '    Branch = oRsB.Fields.Item("BPLId").Value.ToString().Trim()
                        '    If Branch = 0 Then
                        '        Branch = oDBs_Head.GetValue("U_TNXBRC", 0).Trim()
                        '    End If

                        '    Dim objBarcodeFilter As New ClsBarcodeFilter

                        '    objBarcodeFilter.MainFormUID = objForm.UniqueID


                        '    objMain.objClsBarcodeFilter.BarCodeFilter(FormUID, Branch)

                        '    BubbleEvent = False

                        'End If
                        'If oCFL.UniqueID = "CFL_11" Then

                        '    m_FormUID = FormUID
                        '    m_Row = pVal.Row
                        '    Dim Branch As String = ""
                        '    Dim BrandID As String = "Select ""BPLId"" from OBPL where ""BPLName"" = '" + oDBs_Head.GetValue("U_TNXBRC", 0).Trim() + "'"
                        '    Dim oRsB As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '    oRsB.DoQuery(BrandID)
                        '    Branch = oRsB.Fields.Item("BPLId").Value.ToString().Trim()
                        '    If Branch = 0 Then
                        '        Branch = oDBs_Head.GetValue("U_TNXBRC", 0).Trim()
                        '    End If

                        '    Dim objBarcodeFilter As New ClsBarcodeFilter

                        '    objBarcodeFilter.MainFormUID = objForm.UniqueID


                        '    objMain.objClsBarcodeFilter.BarCodeFilter(FormUID, Branch)

                        '    BubbleEvent = False

                        'End If

                    End If


                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        If oCFL.UniqueID = "CFL_5" Then
                            oDBs_Head.SetValue("U_TNXUSR", oDBs_Head.Offset, oDT.GetValue("USER_CODE", 0))
                        End If

                        If oCFL.UniqueID = "CFL_1" Then
                            oDBs_Details1.SetValue("LineID", oDBs_Details1.Offset, pVal.Row)
                            oDBs_Details1.SetValue("U_TNXICD", oDBs_Details1.Offset, oDT.GetValue("ItemCode", 0))
                            oDBs_Details1.SetValue("U_TNXNM", oDBs_Details1.Offset, oDT.GetValue("ItemName", 0))
                            oDBs_Details1.SetValue("U_TNXGRPO", oDBs_Details1.Offset, objMatrix1.Columns.Item("TNXGRPO").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_ULPRC", oDBs_Details1.Offset, objMatrix1.Columns.Item("ULPRC").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_PLPRC", oDBs_Details1.Offset, objMatrix1.Columns.Item("PLPRC").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_SCOST", oDBs_Details1.Offset, objMatrix1.Columns.Item("SCOST").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_OMRG", oDBs_Details1.Offset, objMatrix1.Columns.Item("OMRG").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_NMRG", oDBs_Details1.Offset, objMatrix1.Columns.Item("NMRG").Cells.Item(pVal.Row).Specific.Value)
                            objMatrix1.SetLineData(pVal.Row)
                            'Me.SetNewLine1(objForm.UniqueID)
                            If pVal.Row = objMatrix1.VisualRowCount Then
                                Me.SetNewLine1(objForm.UniqueID)
                            End If
                        End If

                        If oCFL.UniqueID = "CFL_2" Then
                            oDBs_Details2.SetValue("LineID", oDBs_Details2.Offset, pVal.Row)
                            oDBs_Details2.SetValue("U_BrCode", oDBs_Details2.Offset, oDT.GetValue("BcdCode", 0))
                            'oDBs_Details2.SetValue("U_UOM", oDBs_Details2.Offset, oDT.GetValue("UomEntry", 0))
                            Dim avgCost As String = FetchBranchAndAvgPriceByBarcode(oDT.GetValue("BcdCode", 0))

                            Try
                                oDBs_Details3.SetValue("U_CPrice", oDBs_Details3.Offset, avgCost)
                            Catch

                            End Try
                            Try
                                objMatrix3.Columns.Item("CPrice").Cells.Item(pVal.Row).Specific.Value = avgCost
                            Catch
                            End Try

                            oDBs_Details2.SetValue("U_ItemCode", oDBs_Details2.Offset, oDT.GetValue("ItemCode", 0))
                            oDBs_Details2.SetValue("U_ITNM", oDBs_Details2.Offset, oDT.GetValue("BcdName", 0))


                            Dim rsGrp As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            Dim QRYS As String = "SELECT T2.""ItmsGrpNam"" FROM OITM T1 LEFT JOIN OITB T2 ON T1.""ItmsGrpCod"" = T2.""ItmsGrpCod"" 
                            Where T1.""ItemCode"" ='" + oDT.GetValue("ItemCode", 0) + "'"
                            rsGrp.DoQuery(QRYS)


                            Dim rsUOM As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            Dim UOM As String = "SELECT (SELECT ""UomCode"" FROM ""OUOM"" WHERE ""UomEntry"" = B.""UomEntry"") AS ""UomCode"",B.""BcdCode"",T1.""ItemCode""
FROM ""OBCD"" B LEFT JOIN ""OITM"" T1 ON T1.""ItemCode"" = B.""ItemCode""
WHERE B.""BcdCode"" = '" & oDT.GetValue("BcdCode", 0) & "' AND T1.""ItemCode"" = '" & oDT.GetValue("ItemCode", 0) & "'"
                            rsUOM.DoQuery(UOM)
                            oDBs_Details2.SetValue("U_UOM", oDBs_Details2.Offset, rsUOM.Fields.Item("UomCode").Value.ToString())
                            oDBs_Details2.SetValue("U_ITGP", oDBs_Details2.Offset, rsGrp.Fields.Item("ItmsGrpNam").Value.ToString())

                            oDBs_Details2.SetValue("U_CPrice", oDBs_Details2.Offset, objMatrix2.Columns.Item("CPrice").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details2.SetValue("U_NwPrice", oDBs_Details2.Offset, objMatrix2.Columns.Item("NwPrice").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details2.SetValue("U_Cost", oDBs_Details2.Offset, objMatrix2.Columns.Item("Cost").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details2.SetValue("U_LPrice", oDBs_Details2.Offset, objMatrix2.Columns.Item("LPrice").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details2.SetValue("U_Marg", oDBs_Details2.Offset, objMatrix2.Columns.Item("Marg").Cells.Item(pVal.Row).Specific.Value)
                            objMatrix2.SetLineData(pVal.Row)
                            If pVal.Row = objMatrix2.VisualRowCount Then
                                Me.SetNewLine2(objForm.UniqueID)
                            End If
                            'Me.SetNewLine2(objForm.UniqueID)
                        End If

                        If oCFL.UniqueID = "CFL_3" Then
                            oDBs_Details3.SetValue("LineID", oDBs_Details3.Offset, pVal.Row)
                            oDBs_Details3.SetValue("U_BCode", oDBs_Details3.Offset, oDT.GetValue("BcdCode", 0))


                            Dim avgCost As String = FetchBranchAndAvgPriceByBarcode(oDT.GetValue("BcdCode", 0))

                            oDBs_Details3.SetValue("U_ItemCode", oDBs_Details3.Offset, oDT.GetValue("ItemCode", 0))


                            Dim rsUOM As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            Dim UOM As String = "SELECT (SELECT ""UomCode"" FROM ""OUOM"" WHERE ""UomEntry"" = B.""UomEntry"") AS ""UomCode"",B.""BcdCode"",T1.""ItemCode""
FROM ""OBCD"" B LEFT JOIN ""OITM"" T1 ON T1.""ItemCode"" = B.""ItemCode""
WHERE B.""BcdCode"" = '" & oDT.GetValue("BcdCode", 0) & "' AND T1.""ItemCode"" = '" & oDT.GetValue("ItemCode", 0) & "'"
                            rsUOM.DoQuery(UOM)
                            oDBs_Details3.SetValue("U_UOM", oDBs_Details3.Offset, rsUOM.Fields.Item("UomCode").Value.ToString())




                            oDBs_Details3.SetValue("U_IName", oDBs_Details3.Offset, oDT.GetValue("BcdName", 0))
                            'oDBs_Details3.SetValue("U_UOM", oDBs_Details3.Offset, oDT.GetValue("UomEntry", 0))
                            objMatrix3.Columns.Item("CCost").Cells.Item(pVal.Row).Specific.Value = avgCost
                            Try
                                oDBs_Details3.SetValue("U_CCost", oDBs_Details3.Offset, avgCost)
                            Catch

                            End Try
                            Try
                                objMatrix3.Columns.Item("CCost").Cells.Item(pVal.Row).Specific.Value = avgCost
                            Catch
                            End Try

                            'oDBs_Details3.SetValue("U_ItemCode", oDBs_Details3.Offset, objMatrix3.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details3.SetValue("U_CSPrice", oDBs_Details3.Offset, objMatrix3.Columns.Item("CSPrice").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details3.SetValue("U_CMarg", oDBs_Details3.Offset, objMatrix3.Columns.Item("CMarg").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details3.SetValue("U_NSPrice", oDBs_Details3.Offset, objMatrix3.Columns.Item("NSPrice").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details3.SetValue("U_NMarg", oDBs_Details3.Offset, objMatrix3.Columns.Item("NMarg").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details3.SetValue("U_NCost", oDBs_Details3.Offset, objMatrix3.Columns.Item("NCost").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details3.SetValue("U_CName", oDBs_Details3.Offset, objMatrix3.Columns.Item("CName").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details3.SetValue("U_IAttach", oDBs_Details3.Offset, objMatrix3.Columns.Item("IAttach").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details3.SetValue("U_LPhPrice", oDBs_Details3.Offset, objMatrix3.Columns.Item("LPhPrice").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details3.SetValue("U_LGRPPrice", oDBs_Details3.Offset, objMatrix3.Columns.Item("LGRPPrice").Cells.Item(pVal.Row).Specific.Value)
                            objMatrix3.SetLineData(pVal.Row)
                            If pVal.Row = objMatrix3.VisualRowCount Then
                                Me.SetNewLine3(objForm.UniqueID)
                            End If
                            'Me.SetNewLine3(objForm.UniqueID)
                        End If

                        If oCFL.UniqueID = "CFL_6" Then
                            oDBs_Details6.SetValue("LineID", oDBs_Details6.Offset, pVal.Row)
                            oDBs_Details6.SetValue("U_BCode", oDBs_Details6.Offset, oDT.GetValue("BcdCode", 0))


                            Dim avgCost As String = FetchBranchAndAvgPriceByBarcode(oDT.GetValue("BcdCode", 0))

                            oDBs_Details6.SetValue("U_ICode", oDBs_Details6.Offset, oDT.GetValue("ItemCode", 0))
                            oDBs_Details6.SetValue("U_IName", oDBs_Details6.Offset, oDT.GetValue("BcdName", 0))
                            'oDBs_Details3.SetValue("U_UOM", oDBs_Details3.Offset, oDT.GetValue("UomEntry", 0))
                            Dim rsUOM As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            Dim UOM As String = "SELECT (SELECT ""UomCode"" FROM ""OUOM"" WHERE ""UomEntry"" = B.""UomEntry"") AS ""UomCode"",B.""BcdCode"",T1.""ItemCode""
FROM ""OBCD"" B LEFT JOIN ""OITM"" T1 ON T1.""ItemCode"" = B.""ItemCode""
WHERE B.""BcdCode"" = '" & oDT.GetValue("BcdCode", 0) & "' AND T1.""ItemCode"" = '" & oDT.GetValue("ItemCode", 0) & "'"
                            rsUOM.DoQuery(UOM)
                            oDBs_Details6.SetValue("U_Unit", oDBs_Details6.Offset, rsUOM.Fields.Item("UomCode").Value.ToString())

                            'objMatrix6.Columns.Item("Col_4").Cells.Item(pVal.Row).Specific.Value = avgCost
                            Try
                                oDBs_Details6.SetValue("U_CCost", oDBs_Details6.Offset, avgCost)
                            Catch

                            End Try
                            Try
                                objMatrix3.Columns.Item("CCost").Cells.Item(pVal.Row).Specific.Value = avgCost
                            Catch
                            End Try

                            oDBs_Details6.SetValue("U_NSlng", oDBs_Details6.Offset, objMatrix6.Columns.Item("Col_8").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details6.SetValue("U_CMarg", oDBs_Details6.Offset, objMatrix6.Columns.Item("Col_6").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details6.SetValue("U_CSlng", oDBs_Details6.Offset, objMatrix6.Columns.Item("Col_5").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details6.SetValue("U_NMrgn", oDBs_Details6.Offset, objMatrix6.Columns.Item("Col_9").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details6.SetValue("U_NCost", oDBs_Details6.Offset, objMatrix6.Columns.Item("Col_7").Cells.Item(pVal.Row).Specific.Value)

                            objMatrix6.SetLineData(pVal.Row)
                            If pVal.Row = objMatrix6.VisualRowCount Then
                                Me.SetNewLine6(objForm.UniqueID)
                            End If

                        End If

                        If oCFL.UniqueID = "CFL_9" Then
                            oDBs_Details7.SetValue("LineID", oDBs_Details7.Offset, pVal.Row)
                            oDBs_Details7.SetValue("U_BCode", oDBs_Details7.Offset, oDT.GetValue("BcdCode", 0))


                            Dim avgCost As String = FetchBranchAndAvgPriceByBarcode(oDT.GetValue("BcdCode", 0))

                            oDBs_Details7.SetValue("U_ITCD", oDBs_Details7.Offset, oDT.GetValue("ItemCode", 0))
                            oDBs_Details7.SetValue("U_ITNM", oDBs_Details7.Offset, oDT.GetValue("BcdName", 0))


                            Dim rsUOM As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            Dim UOM As String = "SELECT (SELECT ""UomCode"" FROM ""OUOM"" WHERE ""UomEntry"" = B.""UomEntry"") AS ""UomCode"",B.""BcdCode"",T1.""ItemCode""
FROM ""OBCD"" B LEFT JOIN ""OITM"" T1 ON T1.""ItemCode"" = B.""ItemCode""
WHERE B.""BcdCode"" = '" & oDT.GetValue("BcdCode", 0) & "' AND T1.""ItemCode"" = '" & oDT.GetValue("ItemCode", 0) & "'"
                            rsUOM.DoQuery(UOM)
                            oDBs_Details7.SetValue("U_UOM", oDBs_Details7.Offset, rsUOM.Fields.Item("UomCode").Value.ToString())




                            'oDBs_Details3.SetValue("U_UOM", oDBs_Details3.Offset, oDT.GetValue("UomEntry", 0))


                            'oDBs_Details7.SetValue("U_NSlng", oDBs_Details7.Offset, objMatrix7.Columns.Item("Col_8").Cells.Item(pVal.Row).Specific.Value)

                            objMatrix7.SetLineData(pVal.Row)
                            If pVal.Row = objMatrix7.VisualRowCount Then
                                Me.SetNewLine7(objForm.UniqueID)
                            End If

                        End If

                        If oCFL.UniqueID = "CFL_11" Then
                            oDBs_Details9.SetValue("LineID", oDBs_Details9.Offset, pVal.Row)
                            oDBs_Details9.SetValue("U_BCode", oDBs_Details9.Offset, oDT.GetValue("BcdCode", 0))


                            Dim avgCost As String = FetchBranchAndAvgPriceByBarcode(oDT.GetValue("BcdCode", 0))

                            oDBs_Details9.SetValue("U_ITCD", oDBs_Details9.Offset, oDT.GetValue("ItemCode", 0))
                            oDBs_Details9.SetValue("U_BName", oDBs_Details9.Offset, oDT.GetValue("BcdName", 0))


                            Dim rsUOM As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            Dim UOM As String = "SELECT (SELECT ""UomCode"" FROM ""OUOM"" WHERE ""UomEntry"" = B.""UomEntry"") AS ""UomCode"",B.""BcdCode"",T1.""ItemCode""
FROM ""OBCD"" B LEFT JOIN ""OITM"" T1 ON T1.""ItemCode"" = B.""ItemCode""
WHERE B.""BcdCode"" = '" & oDT.GetValue("BcdCode", 0) & "' AND T1.""ItemCode"" = '" & oDT.GetValue("ItemCode", 0) & "'"
                            rsUOM.DoQuery(UOM)
                            oDBs_Details9.SetValue("U_UOM", oDBs_Details9.Offset, rsUOM.Fields.Item("UomCode").Value.ToString())

                            objMatrix9.SetLineData(pVal.Row)
                            If pVal.Row = objMatrix9.VisualRowCount Then
                                Me.SetNewLine9(objForm.UniqueID)
                            End If

                        End If

                        If oCFL.UniqueID = "CFL_4" Then
                            oDBs_Details4.SetValue("LineID", oDBs_Details4.Offset, pVal.Row)
                            oDBs_Details4.SetValue("U_BrCode", oDBs_Details4.Offset, oDT.GetValue("BcdCode", 0))
                            'oDBs_Details4.SetValue("U_UOM1", oDBs_Details4.Offset, oDT.GetValue("UomEntry", 0))
                            Dim avgCost As String = FetchBranchAndAvgPriceByBarcode(oDT.GetValue("BcdCode", 0))

                            Try
                                oDBs_Details4.SetValue("U_ICost", oDBs_Details4.Offset, avgCost)
                            Catch

                            End Try
                            Try
                                objMatrix4.Columns.Item("ICost").Cells.Item(pVal.Row).Specific.Value = avgCost
                            Catch
                            End Try

                            oDBs_Details4.SetValue("U_ItemCode", oDBs_Details4.Offset, oDT.GetValue("ItemCode", 0))

                            Dim rsUOM As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            Dim UOM As String = "SELECT (SELECT ""UomCode"" FROM ""OUOM"" WHERE ""UomEntry"" = B.""UomEntry"") AS ""UomCode"",B.""BcdCode"",T1.""ItemCode""
FROM ""OBCD"" B LEFT JOIN ""OITM"" T1 ON T1.""ItemCode"" = B.""ItemCode""
WHERE B.""BcdCode"" = '" & oDT.GetValue("BcdCode", 0) & "' AND T1.""ItemCode"" = '" & oDT.GetValue("ItemCode", 0) & "'"
                            rsUOM.DoQuery(UOM)
                            oDBs_Details4.SetValue("U_UOM1", oDBs_Details4.Offset, rsUOM.Fields.Item("UomCode").Value.ToString())

                            oDBs_Details4.SetValue("U_ItemN", oDBs_Details4.Offset, oDT.GetValue("BcdName", 0))
                            ' oDBs_Details4.SetValue("U_UOM", oDBs_Details4.Offset, oDT.GetValue("UomEntry", 0))
                            oDBs_Details4.SetValue("U_Qty", oDBs_Details4.Offset, objMatrix4.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value)
                            'oDBs_Details4.SetValue("U_ICost", oDBs_Details4.Offset, objMatrix4.Columns.Item("ICost").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details4.SetValue("U_NSell", oDBs_Details4.Offset, objMatrix4.Columns.Item("NSell").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details4.SetValue("U_SPrice", oDBs_Details4.Offset, objMatrix4.Columns.Item("SPrice").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details4.SetValue("U_DiffSP", oDBs_Details4.Offset, objMatrix4.Columns.Item("DiffSP").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details4.SetValue("U_DSPNS", oDBs_Details4.Offset, objMatrix4.Columns.Item("DSPNS").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details4.SetValue("U_TPRC", oDBs_Details4.Offset, objMatrix4.Columns.Item("TPRC").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details4.SetValue("U_TPRS", oDBs_Details4.Offset, objMatrix4.Columns.Item("TPRS").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details4.SetValue("U_FundA", oDBs_Details4.Offset, objMatrix4.Columns.Item("FundA").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details4.SetValue("U_NGP", oDBs_Details4.Offset, objMatrix4.Columns.Item("NGP").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details4.SetValue("U_SGP", oDBs_Details4.Offset, objMatrix4.Columns.Item("SGP").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details4.SetValue("U_Supp", oDBs_Details4.Offset, objMatrix4.Columns.Item("Supp").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details4.SetValue("U_Remark", oDBs_Details4.Offset, objMatrix4.Columns.Item("Remark").Cells.Item(pVal.Row).Specific.Value)
                            objMatrix4.SetLineData(pVal.Row)
                            If pVal.Row = objMatrix4.VisualRowCount Then
                                Me.SetNewLine4(objForm.UniqueID)
                            End If
                            'Me.SetNewLine4(objForm.UniqueID)
                        End If

                        If oCFL.UniqueID = "CFL_7" Then

                            If Not oDT Is Nothing Then

                                objMatrix4.FlushToDataSource()

                                oDBs_Details4.Offset = pVal.Row - 1

                                oDBs_Details4.SetValue("U_CRNT", oDBs_Details4.Offset, oDT.GetValue("DocEntry", 0).ToString())

                                objMatrix4.LoadFromDataSource()

                            End If

                        End If

                        If oCFL.UniqueID = "CFL_8" Then

                            If Not oDT Is Nothing Then
                                oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C1")

                                objMatrix2.FlushToDataSource()

                                oDBs_Details2.Offset = pVal.Row - 1

                                oDBs_Details2.SetValue("U_CRNT", oDBs_Details2.Offset, oDT.GetValue("DocEntry", 0).ToString())

                                objMatrix2.LoadFromDataSource()

                            End If

                        End If


                    End If

                    If (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = True Then

                        If oCFL.UniqueID = "CFL_7" Then

                            Dim Item As String = objMatrix4.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value

                            Me.FilterAP(FormUID, "CFL_7", Item)

                        End If

                    End If
                    If (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = True Then

                        If oCFL.UniqueID = "CFL_8" Then

                            Dim Item As String = objMatrix2.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value

                            Me.FilterAP2(FormUID, "CFL_8", Item)

                        End If

                    End If

                    'oDBs_Details4.SetValue("U_ItemN", oDBs_Details4.Offset, objMatrix4.Columns.Item("ItemN").Cells.Item(pVal.Row).Specific.Value)
                    '' oDBs_Details4.SetValue("U_UOM", oDBs_Details4.Offset, oDT.GetValue("UomEntry", 0))
                    'oDBs_Details4.SetValue("U_Qty", oDBs_Details4.Offset, objMatrix4.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value)
                    ''oDBs_Details4.SetValue("U_ICost", oDBs_Details4.Offset, objMatrix4.Columns.Item("ICost").Cells.Item(pVal.Row).Specific.Value)
                    'oDBs_Details4.SetValue("U_NSell", oDBs_Details4.Offset, objMatrix4.Columns.Item("NSell").Cells.Item(pVal.Row).Specific.Value)
                    'oDBs_Details4.SetValue("U_SPrice", oDBs_Details4.Offset, objMatrix4.Columns.Item("SPrice").Cells.Item(pVal.Row).Specific.Value)
                    'oDBs_Details4.SetValue("U_DiffSP", oDBs_Details4.Offset, objMatrix4.Columns.Item("DiffSP").Cells.Item(pVal.Row).Specific.Value)
                    'oDBs_Details4.SetValue("U_DSPNS", oDBs_Details4.Offset, objMatrix4.Columns.Item("DSPNS").Cells.Item(pVal.Row).Specific.Value)
                    'oDBs_Details4.SetValue("U_TPRC", oDBs_Details4.Offset, objMatrix4.Columns.Item("TPRC").Cells.Item(pVal.Row).Specific.Value)
                    'oDBs_Details4.SetValue("U_TPRS", oDBs_Details4.Offset, objMatrix4.Columns.Item("TPRS").Cells.Item(pVal.Row).Specific.Value)
                    'oDBs_Details4.SetValue("U_FundA", oDBs_Details4.Offset, objMatrix4.Columns.Item("FundA").Cells.Item(pVal.Row).Specific.Value)
                    'oDBs_Details4.SetValue("U_NGP", oDBs_Details4.Offset, objMatrix4.Columns.Item("NGP").Cells.Item(pVal.Row).Specific.Value)
                    'oDBs_Details4.SetValue("U_SGP", oDBs_Details4.Offset, objMatrix4.Columns.Item("SGP").Cells.Item(pVal.Row).Specific.Value)
                    'oDBs_Details4.SetValue("U_Supp", oDBs_Details4.Offset, objMatrix4.Columns.Item("Supp").Cells.Item(pVal.Row).Specific.Value)
                    'oDBs_Details4.SetValue("U_Remark", oDBs_Details4.Offset, objMatrix4.Columns.Item("Remark").Cells.Item(pVal.Row).Specific.Value)
                    'objMatrix4.SetLineData(pVal.Row)
                    '        'If pVal.Row = objMatrix4.VisualRowCount Then
                    '        '    Me.SetNewLine4(objForm.UniqueID)
                    '        'End If
                    '        'Me.SetNewLine4(objForm.UniqueID)


                Case SAPbouiCOM.BoEventTypes.et_CLICK


                    'Mop tab Based item loading uom list
                    If pVal.ItemUID = "Mtx3" And pVal.ColUID = "UOM" And pVal.BeforeAction = True Then
                        Try
                            objForm.Freeze(True)

                            Dim row As Integer = pVal.Row

                            Dim itemCode As String = oDBs_Details3.GetValue("U_ItemCode", row - 1).Trim()
                            If itemCode = "" Then Exit Sub

                            Dim oColumn As SAPbouiCOM.Column = objMatrix3.Columns.Item("UOM")

                            ' Clear previous values
                            While oColumn.ValidValues.Count > 0
                                oColumn.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
                            End While

                            ' Fetch UOM for that item only
                            Dim rsUOM As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            rsUOM.DoQuery("SELECT DISTINCT U.""UomEntry"", U.""UomCode"" " & "FROM ""OITM"" I " &
                             "INNER JOIN ""UGP1"" G ON I.""UgpEntry"" = G.""UgpEntry"" " & "INNER JOIN ""OUOM"" U ON G.""UomEntry"" = U.""UomEntry"" " &
                                "WHERE I.""ItemCode"" = '" & itemCode & "'")

                            While Not rsUOM.EoF
                                oColumn.ValidValues.Add(rsUOM.Fields.Item("UomCode").Value.ToString(), rsUOM.Fields.Item("UomCode").Value.ToString())
                                rsUOM.MoveNext()
                            End While

                            ' Reselect stored value
                            Dim oCombo As SAPbouiCOM.ComboBox = CType(oColumn.Cells.Item(row).Specific, SAPbouiCOM.ComboBox)

                            Dim storedUOM As String = oDBs_Details3.GetValue("U_UOM", row - 1).Trim()

                            If storedUOM <> "" Then
                                Try
                                    oCombo.Select(storedUOM, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Catch
                                End Try
                            End If
                        Catch ex As Exception
                        Finally
                            objForm.Freeze(False)
                        End Try

                    End If

                    'Flyer Tab based on item loading uom list
                    If pVal.ItemUID = "Mtx2" And pVal.ColUID = "UOM" And pVal.BeforeAction = True Then
                        Try
                            objForm.Freeze(True)

                            Dim row As Integer = pVal.Row

                            Dim itemCode As String = oDBs_Details2.GetValue("U_ItemCode", row - 1).Trim()
                            If itemCode = "" Then Exit Sub

                            Dim oColumn As SAPbouiCOM.Column = objMatrix2.Columns.Item("UOM")

                            ' Clear previous values
                            While oColumn.ValidValues.Count > 0
                                oColumn.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
                            End While

                            ' Fetch UOM for that item only
                            Dim rsUOM As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            rsUOM.DoQuery("SELECT DISTINCT U.""UomEntry"", U.""UomCode"" " & "FROM ""OITM"" I " &
                             "INNER JOIN ""UGP1"" G ON I.""UgpEntry"" = G.""UgpEntry"" " & "INNER JOIN ""OUOM"" U ON G.""UomEntry"" = U.""UomEntry"" " &
                                "WHERE I.""ItemCode"" = '" & itemCode & "'")

                            While Not rsUOM.EoF
                                oColumn.ValidValues.Add(rsUOM.Fields.Item("UomCode").Value.ToString(), rsUOM.Fields.Item("UomCode").Value.ToString())
                                rsUOM.MoveNext()
                            End While

                            ' Reselect stored value
                            Dim oCombo As SAPbouiCOM.ComboBox = CType(oColumn.Cells.Item(row).Specific, SAPbouiCOM.ComboBox)

                            Dim storedUOM As String = oDBs_Details2.GetValue("U_UOM", row - 1).Trim()

                            If storedUOM <> "" Then
                                Try
                                    oCombo.Select(storedUOM, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Catch
                                End Try
                            End If
                        Catch ex As Exception
                        Finally
                            objForm.Freeze(False)
                        End Try

                    End If
                    'Tempoary price Reduction based on item loading uom list
                    If pVal.ItemUID = "Mtx4" And pVal.ColUID = "UOM1" And pVal.BeforeAction = True Then
                        Try
                            objForm.Freeze(True)

                            Dim row As Integer = pVal.Row

                            Dim itemCode As String = oDBs_Details4.GetValue("U_ItemCode", row - 1).Trim()
                            If itemCode = "" Then Exit Sub

                            Dim oColumn As SAPbouiCOM.Column = objMatrix4.Columns.Item("UOM1")

                            ' Clear previous values
                            While oColumn.ValidValues.Count > 0
                                oColumn.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
                            End While

                            ' Fetch UOM for that item only
                            Dim rsUOM As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            rsUOM.DoQuery("SELECT DISTINCT U.""UomEntry"", U.""UomCode"" " & "FROM ""OITM"" I " &
                             "INNER JOIN ""UGP1"" G ON I.""UgpEntry"" = G.""UgpEntry"" " & "INNER JOIN ""OUOM"" U ON G.""UomEntry"" = U.""UomEntry"" " &
                                "WHERE I.""ItemCode"" = '" & itemCode & "'")

                            While Not rsUOM.EoF
                                oColumn.ValidValues.Add(rsUOM.Fields.Item("UomCode").Value.ToString(), rsUOM.Fields.Item("UomCode").Value.ToString())
                                rsUOM.MoveNext()
                            End While

                            ' Reselect stored value
                            Dim oCombo As SAPbouiCOM.ComboBox = CType(oColumn.Cells.Item(row).Specific, SAPbouiCOM.ComboBox)

                            Dim storedUOM As String = oDBs_Details4.GetValue("U_UOM1", row - 1).Trim()

                            If storedUOM <> "" Then
                                Try
                                    oCombo.Select(storedUOM, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Catch
                                End Try
                            End If
                        Catch ex As Exception
                        Finally
                            objForm.Freeze(False)
                        End Try

                    End If



                '    'In Find mode after enterning Doc Num - Enter button not working 

                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    If pVal.BeforeAction = True AndAlso pVal.CharPressed = 13 Then ' Enter key
                        objForm = objMain.objApplication.Forms.Item(FormUID)
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            If pVal.ItemUID = "8" OrElse pVal.ItemUID = "DocNum" Then
                                Try
                                    BubbleEvent = False
                                    objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                Catch ex As Exception
                                    objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End Try
                            End If

                        End If

                    End If

                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    ' Handle combo selection to switch tabs (PaneLevel) and show/hide Flyer field
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    ' NOTE: control UID in XML is "TNXPUTP" (data alias is U_TNXPUTP). Use the item UID.
                    If pVal.ItemUID = "TNXPUTP" AndAlso pVal.BeforeAction = False Then
                        Try
                            Dim oCombo As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("TNXPUTP").Specific, SAPbouiCOM.ComboBox)
                            If oCombo Is Nothing OrElse oCombo.Selected Is Nothing Then
                                Return
                            End If
                            Dim selected As String = If(oCombo.Selected.Value, "").ToString().Trim()
                            Dim DocEntry As String = oDBs_Head.GetValue("DocEntry", 0).Trim()
                            If DocEntry <> "" Then


                                Dim rsTAB As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                rsTAB.DoQuery("SELECT IFNULL(""U_TNXPUTP"",'') AS ""SelectedTab"" FROM ""@TNX_PCRT"" WHERE ""DocEntry"" = '" & DocEntry & "'")

                                Dim SelectedTab As String = ""

                                If rsTAB.RecordCount > 0 Then
                                    SelectedTab = rsTAB.Fields.Item("SelectedTab").Value.ToString().Trim()
                                End If

                                If Not String.IsNullOrEmpty(SelectedTab) AndAlso SelectedTab <> selected Then

                                    objMain.objApplication.StatusBar.SetText("You have already selected '" & SelectedTab & "'. Please clear the current selection before selecting another tab.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                                    Exit Sub
                                End If


                                SelectedTab = selected
                            End If
                            ' set pane level according to selection
                            Select Case selected
                                Case "Cost Fluctuation"
                                    objForm.PaneLevel = 1
                                    objForm.Settings.MatrixUID = "Mtx1"
                                    objForm.EnableMenu("5890", True)
                                    objForm.Items.Item("Item_8").Visible = False
                                Case "Flyer Price Changes"
                                    objForm.PaneLevel = 2
                                    objForm.Settings.MatrixUID = "Mtx2"
                                    objForm.EnableMenu("5890", True)
                                    objForm.Items.Item("Item_8").Visible = True

                                Case "MOP"
                                    objForm.PaneLevel = 3
                                    objForm.Settings.MatrixUID = "Mtx3"
                                    objForm.EnableMenu("5890", True)
                                    objForm.Items.Item("Item_8").Visible = False
                                Case "Temporary Price Reduction"
                                    objForm.PaneLevel = 4
                                    objForm.Settings.MatrixUID = "Mtx4"
                                    objForm.EnableMenu("5890", True)
                                    objForm.Items.Item("Item_8").Visible = False
                                Case "Attachment Tab"
                                    objForm.PaneLevel = 8
                                    objForm.Settings.MatrixUID = "Mtx5"
                                    objForm.EnableMenu("5890", True)
                                    objForm.Items.Item("Item_8").Visible = False
                                Case "Every day Special Price"
                                    objForm.PaneLevel = 5
                                    objForm.Settings.MatrixUID = "Mtx6"
                                    objForm.EnableMenu("5890", True)
                                    objForm.Items.Item("Item_8").Visible = False
                                Case "Bulk Cost Update"
                                    objForm.PaneLevel = 6
                                    objForm.Settings.MatrixUID = "Mtx7"
                                    objForm.EnableMenu("5890", True)
                                    objForm.Items.Item("Item_8").Visible = False
                                Case "Cost Fluctuation Stock Transfer"
                                    objForm.PaneLevel = 7
                                    objForm.Settings.MatrixUID = "Mtx8"
                                    objForm.EnableMenu("5890", True)
                                    objForm.Items.Item("Item_8").Visible = False
                                Case "Bulk Price Update"
                                    objForm.PaneLevel = 8
                                    objForm.Settings.MatrixUID = "Mtx9"
                                    objForm.EnableMenu("5890", True)
                                    objForm.Items.Item("Item_8").Visible = False
                                Case Else
                                    objForm.PaneLevel = 1
                            End Select



                            ' Show Flyer control only when Flyer Price Changes selected
                            objForm.Freeze(True)
                            Dim showFlyer As Boolean = (selected = "Flyer Price Changes")
                            Try
                                objForm.Items.Item("Item_57").Visible = showFlyer
                                objForm.Items.Item("Item_10").Visible = showFlyer

                            Catch ex As Exception
                            End Try
                            Try
                                objForm.Items.Item("Flyer").Visible = showFlyer
                                objForm.Items.Item("Item_13").Visible = showFlyer
                            Catch ex As Exception
                            End Try

                            ' Show TPR control only when TPR Price Changes selected

                            Dim showTPR As Boolean = (selected = "Temporary Price Reduction")
                            Try

                                objForm.Items.Item("Item_10").Visible = showTPR

                            Catch ex As Exception
                            End Try
                            Try

                                objForm.Items.Item("Item_13").Visible = showTPR
                            Catch ex As Exception
                            End Try

                            ' when hiding TPR, clear its value from datasource to avoid stale selection
                            If Not showTPR Then
                                Try
                                    Dim headDs As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")

                                    headDs.SetValue("U_TPRT", headDs.Offset, "")
                                Catch ex As Exception
                                    ' ignore clearing errors
                                End Try
                            End If

                            If pVal.ItemUID = "TNXPUTP" AndAlso pVal.BeforeAction = False Then
                                ''Test
                                Dim oDBs_Head As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")

                                Dim updateType As String = oDBs_Head.GetValue("U_TNXPUTP", 0).Trim()

                                SetTabsByUpdateType(updateType)

                            End If

                            objForm.Freeze(False)
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText("Error switching pane / toggling Flyer:  " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If

                    If pVal.ItemUID = "Mtx1" AndAlso pVal.ColUID = "Col_2" AndAlso pVal.BeforeAction = False Then

                        Dim row As Integer = pVal.Row

                        Dim sptpValue As String = objMatrix1.Columns.Item("Col_2").Cells.Item(row).Specific.Value

                        If sptpValue = "Scheduled Update" Then

                            'Current Date
                            Dim currDate As String = DateTime.Now.ToString("yyyyMMdd")

                            'Current Time (SAP format HHmm)
                            Dim currTime As String = DateTime.Now.ToString("HHmm")

                            objMatrix1.Columns.Item("Col_3").Cells.Item(row).Specific.Value = currDate
                            objMatrix1.Columns.Item("Col_4").Cells.Item(row).Specific.Value = currTime

                        End If

                    End If

                    ' Handle Selling Price Type (TNXTYP) show/hide of scheduled fields
                    'If pVal.ItemUID = "TNXTYP" AndAlso pVal.BeforeAction = False Then
                    '    Try
                    '        Dim oTypeCombo As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("TNXTYP").Specific, SAPbouiCOM.ComboBox)
                    '        If oTypeCombo Is Nothing OrElse oTypeCombo.Selected Is Nothing Then
                    '            Return
                    '        End If
                    '        Dim selType As String = If(oTypeCombo.Selected.Value, "").ToString().Trim()

                    '        ' If Scheduled Update selected -> show date/time fields; otherwise hide them.
                    '        Dim showScheduled As Boolean = (selType = "Scheduled Update")

                    '        objForm.Freeze(True)
                    '        Try
                    '            objForm.Items.Item("Item_38").Visible = showScheduled ' label "From Date"
                    '        Catch ex As Exception
                    '        End Try
                    '        Try
                    '            objForm.Items.Item("FDate").Visible = showScheduled ' date field
                    '        Catch ex As Exception
                    '        End Try
                    '        Try
                    '            objForm.Items.Item("Item_39").Visible = showScheduled ' label "From Time"
                    '        Catch ex As Exception
                    '        End Try
                    '        Try
                    '            objForm.Items.Item("FTime").Visible = showScheduled ' time field
                    '        Catch ex As Exception
                    '        End Try

                    '        ' If hiding scheduled fields, clear their datasource values
                    '        If Not showScheduled Then
                    '            Try
                    '                Dim headDs As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
                    '                headDs.SetValue("U_FDate", headDs.Offset, "")
                    '                headDs.SetValue("U_FTime", headDs.Offset, "")
                    '            Catch ex As Exception
                    '            End Try
                    '        End If
                    '        objForm.Freeze(False)
                    '    Catch ex As Exception
                    '        objMain.objApplication.StatusBar.SetText("Error toggling Scheduled fields: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    End Try
                    'End If



                    'Based on branch selection automactially load price list num
                    If pVal.ItemUID = "TNXBRC" And pVal.BeforeAction = False Then
                        Try
                            Dim oBranchCombo As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("TNXBRC").Specific, SAPbouiCOM.ComboBox)

                            Dim oPriceListCombo As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("PriceLNum").Specific, SAPbouiCOM.ComboBox)

                            Dim branchId As String = oBranchCombo.Selected.Value


                            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            oRs.DoQuery("SELECT ""U_PricelistNo"" FROM ""OBPL"" WHERE ""BPLId"" = '" & branchId & "'")

                            If Not oRs.EoF Then
                                Dim priceListNo As String = oRs.Fields.Item("U_PricelistNo").Value.ToString()

                                If priceListNo <> "" Then
                                    oPriceListCombo.Select(priceListNo, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Else

                                    oPriceListCombo.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                End If
                            End If

                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText("Branch → PriceList load error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If

                    'disbled  submit button code here...!

                    If pVal.ItemUID = "FOUT" Then

                        Try
                            objForm.Freeze(True)

                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")

                            Dim updateType As String = ""
                            Dim foutValue As String = ""

                            Try
                                updateType = oDBs_Head.GetValue("U_TNXPUTP", 0).Trim()
                            Catch
                            End Try

                            Try
                                foutValue = oDBs_Head.GetValue("U_FOUT", 0).Trim().ToUpper()
                            Catch
                            End Try

                            'If updateType = "Flyer Price Changes" Then

                            '    If foutValue = "YES" OrElse foutValue = "Y" Then
                            '        objForm.Items.Item("Submit").Enabled = True
                            '    ElseIf foutValue = "NO" OrElse foutValue = "N" Then
                            '        objForm.Items.Item("Submit").Enabled = False
                            '    End If
                            '    Me.CheckSubmitButtonVisibility(FormUID)
                            'End If

                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(
                                ex.Message,
                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        Finally
                            objForm.Freeze(False)
                        End Try

                    End If




                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    If pVal.ItemUID = "CalGP" AndAlso pVal.Before_Action = False Then
                        CalculateTotalGP(pVal.FormUID)
                        IsGPCalculated = True
                    End If



                    If pVal.ItemUID = "1" AndAlso pVal.Before_Action = True AndAlso objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE AndAlso CType(objForm.Items.Item("TNXPUTP").Specific, SAPbouiCOM.ComboBox).Selected.Value = "Flyer Price Changes" Then

                        If IsGPCalculated = False Then

                            objMain.objApplication.StatusBar.SetText("Please click CalculateGP button before adding the document.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                            BubbleEvent = False
                            Exit Sub

                        End If

                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.Validate() = False Then

                            BubbleEvent = False
                        End If
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        If Me.Validate1() = False Then

                            BubbleEvent = False
                        End If
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        'If Me.Validate() = False Then
                        '    BubbleEvent = False
                        'End If
                        Me.SetDefault(objForm.UniqueID)
                        'If Not ValidateBarcodeByTab(objForm) Then
                        '    BubbleEvent = False
                        '    Exit Sub


                    End If

                    'If pVal.ItemUID = "1" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) And pVal.BeforeAction = True Then

                    '    If Me.CheckDocNumAndGenerateAutoID() = False Then

                    '        BubbleEvent = False

                    '        Exit Sub

                    '    End If

                    'End If



                    If pVal.ItemUID = "Item_8" And pVal.BeforeAction = False And pVal.ActionSuccess = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then

                        Dim DocEntry As String = oDBs_Head.GetValue("DocEntry", 0).Trim()
                        Me.CreateFormNEW(DocEntry)

                    End If

                    If pVal.ItemUID = "Item_23" And pVal.BeforeAction = False And pVal.ActionSuccess = True And (pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                        If oDBs_Head.GetValue("U_EXCE", 0).Trim() = "" Then
                            Dim DocEntry1 As String = oDBs_Head.GetValue("DocEntry", 0).Trim()
                            Dim ApplicationIDPOS As String = oDBs_Head.GetValue("U_AppId", 0).Trim()
                            objMain.objClsPOSBill.CreateForm(DocEntry1, ApplicationIDPOS)
                        Else
                            objMain.objApplication.StatusBar.SetText("Excel is already posted for this document " & SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If



                    End If


                    If pVal.ItemUID = "Item_28" And pVal.BeforeAction = False And pVal.ActionSuccess = True And (pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        objForm = objMain.objApplication.Forms.Item(FormUID)
                        Dim DocEntry As String = oDBs_Head.GetValue("DocEntry", 0).Trim()
                        Dim Branch As String = ""
                        ''Test
                        Dim BrandID As String = "Select ""BPLId"" from OBPL where ""BPLName"" = '" + oDBs_Head.GetValue("U_TNXBRC", 0).Trim() + "'"
                        Dim oRsB As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsB.DoQuery(BrandID)
                        Branch = oRsB.Fields.Item("BPLId").Value.ToString().Trim()
                        If Branch = 0 Then
                            Branch = oDBs_Head.GetValue("U_TNXBRC", 0).Trim()
                        End If
                        Dim EmailList As String = ""

                        Dim qry As String = "
SELECT DISTINCT EMAIL FROM (
    SELECT IFNULL(T2.""E_Mail"",'') AS EMAIL
    FROM ""@MGRCONFIG"" T0
    INNER JOIN ""@MGRCONFIG_C2"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
    LEFT JOIN ""OUSR"" T2 ON T1.""U_UserIDE"" = T2.""USERID""
    WHERE T0.""U_Doc"" = 'PCR' AND T0.""U_Status"" = 'Y' and T0.""U_Branch"" = '" & Branch & "'

    UNION

    SELECT IFNULL(T3.""email"",'') AS EMAIL
    FROM ""@MGRCONFIG"" T0
    INNER JOIN ""@MGRCONFIG_C2"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
    LEFT JOIN ""OHEM"" T3 ON T1.""U_EmpIDE"" = T3.""empID""
    WHERE T0.""U_Doc"" = 'PCR' AND T0.""U_Status"" = 'Y' and T0.""U_Branch"" = '" & Branch & "'
) A
WHERE EMAIL <> ''"

                        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRs.DoQuery(qry)

                        If oRs.RecordCount > 0 Then
                            While Not oRs.EoF

                                Dim mailId As String = oRs.Fields.Item("EMAIL").Value.ToString().Trim()

                                If mailId <> "" Then
                                    EmailList &= mailId & ","
                                End If

                                oRs.MoveNext()
                            End While
                        End If

                        Dim Email As String = EmailList.TrimEnd(","c)


                        Dim QUERYBODY As String = "Select T0.""U_MsgEML"",T0.""U_Sub"" from ""@MGRCONFIG"" T0 where T0.""U_Doc""='PCR' and T0.""U_Status""='Y' and T0.""U_Branch"" = '" & Branch & "'"

                        Dim oRsBody As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsBody.DoQuery(QUERYBODY)

                        ''Dim Email As String = "ramya.yarlagadda@10xsf.com"
                        Dim Subject As String = oRsBody.Fields.Item("U_Sub").Value.ToString().Trim()
                        Dim Body As String = oRsBody.Fields.Item("U_MsgEML").Value.ToString().Trim()
                        Dim DocType As String = "PCR"
                        Me.SendEmail(Email, Subject, Body, FormUID, DocEntry, DocType)

                    End If

                    'If pVal.ItemUID = "Item_24" And pVal.BeforeAction = False Then
                    '    Dim DocEntry1 As String = oDBs_Head.GetValue("U_DOCE", 0).Trim()
                    '    objMain.objClsPOSBill.CreateFormNEW(DocEntry1)
                    'End If



                    If pVal.ItemUID = "Check" AndAlso pVal.BeforeAction = False Then

                        objForm.Freeze(True)

                        Try
                            Dim oChk As SAPbouiCOM.CheckBox = CType(objForm.Items.Item("Check").Specific, SAPbouiCOM.CheckBox)

                            Dim oDBs As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")

                            If objForm.Items.Item("Check").Specific.Checked = True Then

                                objForm.Items.Item("TDATEMP").Enabled = True
                                objForm.Items.Item("TTIMEM").Enabled = True

                                'objForm.Items.Item("FDTT").Enabled = True
                                'objForm.Items.Item("TDTT").Enabled = True
                            Else

                                oDBs.SetValue("U_FDTT", 0, "")
                                oDBs.SetValue("U_TDTT", 0, "")

                                objForm.Items.Item("TDATEMP").Enabled = False
                                objForm.Items.Item("TTIMEM").Enabled = False
                                'objForm.Items.Item("FDTT").Enabled = False
                                'objForm.Items.Item("TDTT").Enabled = False
                            End If

                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Finally
                            objForm.Freeze(False)
                        End Try

                    End If



                    'submit
                    If pVal.ItemUID = "Submit" And pVal.BeforeAction = False Then
                        Try
                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Dim DocNum As String = objForm.Items.Item("AppID").Specific.value
                                Dim PriceChangeRequestType As String = CType(objForm.Items.Item("TNXPUTP").Specific, SAPbouiCOM.ComboBox).Selected.Value
                                Dim DOCSTATUS As String = objForm.Items.Item("TNXSTUS").Specific.value
                                If DOCSTATUS = "O" Or DOCSTATUS = "Open" Then
                                    'Validate New Selling Price for Cost Fluctuation
                                    If oDBs_Head.GetValue("U_TNXPUTP", 0) = "Cost Fluctuation" Then

                                        For j As Integer = 1 To objMatrix1.VisualRowCount

                                            Dim Item As String = objMatrix1.Columns.Item("TNXICD").Cells.Item(j).Specific.Value
                                            Dim Line As String = objMatrix1.Columns.Item("LineID").Cells.Item(j).Specific.Value

                                            If Item <> "" Then

                                                Dim NewSellingPrice As String = objMatrix1.Columns.Item("NRP").Cells.Item(j).Specific.Value

                                                '                    If String.IsNullOrWhiteSpace(NewSellingPrice) Then
                                                '                        objMain.objApplication.SetStatusBarMessage(
                                                '"New Selling Price is mandatory for Line No. " & Line,
                                                'SAPbouiCOM.BoMessageTime.bmt_Short,
                                                'SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                                                '                        Return False
                                                '                    End If

                                                Dim PriceValue As Double

                                                If Not Double.TryParse(NewSellingPrice, PriceValue) OrElse PriceValue <= 0 Then
                                                    objMain.objApplication.SetStatusBarMessage("Please enter a valid New Selling Price for Line No. " & Line, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                                                    BubbleEvent = False
                                                    Exit Sub
                                                End If

                                            End If

                                        Next

                                    End If

                                    Me.ApprovalTrigger("Price", DocNum, "@TNX_PCRT", "U_AppID", "U_TNXSTUS", "PriceChangeRequest", PriceChangeRequestType)
                                ElseIf DOCSTATUS = "R" Then
                                    Dim rs As SAPbobsCOM.Recordset
                                    rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    Dim NewAppID As String = ""
                                    Dim Cnt As Integer = 1

                                    rs.DoQuery("SELECT COUNT(*) + 1 AS ""CNT"" FROM ""@SBO_DD"" WHERE ""U_docnum"" LIKE '" & DocNum & "_%'")

                                    If rs.RecordCount > 0 Then
                                        Cnt = rs.Fields.Item("CNT").Value
                                    End If

                                    NewAppID = DocNum & "_" & Cnt
                                    Me.ApprovalTrigger("Price", NewAppID, "@TNX_PCRT", "U_AppID", "U_TNXSTUS", "PriceChangeRequest", PriceChangeRequestType)
                                End If
                                'objForm.Items.Item("2").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            ElseIf objForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                objMain.objApplication.SetStatusBarMessage("Please update the screen", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                            End If


                        Catch ex As Exception
                        End Try

                    End If

                    'working code for excel upload and copy and paste madhu
                    'If pVal.ItemUID = "Load" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or
                    '  pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then

                    '    Dim str As String = ""
                    '    Try


                    '        str = objForm.Items.Item("EXLNK").Specific.value
                    '    Catch ex As Exception
                    '        str = ""
                    '    End Try

                    '    If Not String.IsNullOrWhiteSpace(str) Then
                    '        Dim FilePathCSV As String = str.ToString().Trim()

                    '        ' Ensure header DB datasource contains the path so other code can read it if needed
                    '        Try
                    '            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
                    '            oDBs_Head.SetValue("U_EXLNK", oDBs_Head.Offset, FilePathCSV)
                    '        Catch ex As Exception
                    '            ' ignore but log if you want
                    '        End Try

                    '        ' Start background worker and pass the path as argument (so worker doesn't depend on datasource timing)
                    '        Try

                    '            If BackgroundWorker_CSVLoad.IsBusy <> True Then
                    '                BackgroundWorker_CSVLoad.RunWorkerAsync(FilePathCSV)
                    '                objMain.objApplication.StatusBar.SetText("Started loading Excel: " & FilePathCSV, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    '            Else
                    '                objMain.objApplication.StatusBar.SetText("Loader already running. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                    '            End If
                    '        Catch ex As Exception
                    '            objMain.objApplication.StatusBar.SetText("Unable to start loader: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        End Try
                    '    Else

                    '        'copy and paste mop tab data from excel to datatable
                    '        Dim dtFromMatrix As DataTable = GetDataFromMatrix()

                    '        'copy and paste Flyer tab data from excel to datatable
                    '        Dim FlyerTabMatrix As DataTable = GetDataFromMatrixFlyer()


                    '        'copy and paste Temporay tab data from excel to datatable
                    '        Dim TemporayTabMatrix As DataTable = GetDataFromMatrixTempoary()

                    '        If dtFromMatrix.Rows.Count > 0 Then
                    '            UploadFromDataTable(dtFromMatrix)

                    '        ElseIf FlyerTabMatrix.Rows.Count > 0 Then
                    '            UploadFromDataTableFlyerTab(FlyerTabMatrix)

                    '        ElseIf TemporayTabMatrix.Rows.Count > 0 Then
                    '            UploadFromDataTableTemporayTab(TemporayTabMatrix)

                    '        Else
                    '            'objMain.objApplication.StatusBar.SetText("Please select an Excel file or paste data into the matrix",
                    '            '        SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        End If

                    '        ' objMain.objApplication.StatusBar.SetText("Please select an Excel file path before clicking Load.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    End If
                    'End If




                    ' LOAD BUTTON – EXCEL UPLOAD
                    If pVal.ItemUID = "Load" AndAlso pVal.BeforeAction = True AndAlso (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then

                        Dim str As String = ""

                        Try
                            str = objForm.Items.Item("EXLNK").Specific.value
                        Catch ex As Exception
                            str = ""
                        End Try

                        If Not String.IsNullOrWhiteSpace(str) Then

                            Dim FilePathCSV As String = str.Trim()
                            Try
                                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
                                oDBs_Head.SetValue("U_EXLNK", oDBs_Head.Offset, FilePathCSV)
                            Catch ex As Exception

                            End Try

                            Try
                                If BackgroundWorker_CSVLoad.IsBusy = False Then
                                    BackgroundWorker_CSVLoad.RunWorkerAsync(FilePathCSV)

                                    objMain.objApplication.StatusBar.SetText("Started loading Excel: " & FilePathCSV, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                                Else
                                    objMain.objApplication.StatusBar.SetText("Loader already running. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                End If

                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText("Unable to start loader: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Try

                        Else
                            objMain.objApplication.StatusBar.SetText("Please select an Excel file before clicking Load.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                    End If

                    ' LOADMTX BUTTON – COPY & PASTE DATA 
                    If pVal.ItemUID = "LoadMtx" AndAlso pVal.BeforeAction = True AndAlso (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then

                        ' MOP Tab
                        Dim dtFromMatrix As DataTable = GetDataFromMatrix()

                        'Every day Special Price Tab
                        Dim dEDSPFromMatrix As DataTable = GetDataFromMatrixEDSP()

                        'Bulk Cost Upload Tab
                        Dim dtBCUFromMatrix As DataTable = GetDataFromMatrixBCU()

                        ' Flyer Tab
                        Dim FlyerTabMatrix As DataTable = GetDataFromMatrixFlyer()

                        ' Temporary Tab
                        Dim TemporayTabMatrix As DataTable = GetDataFromMatrixTempoary()

                        'Bulk Price Upload Tab
                        Dim dtBPUFromMatrix As DataTable = GetDataFromMatrixBPU()
                        'If dtFromMatrix.Rows.Count > 0 Then
                        '    UploadFromDataTable(dtFromMatrix)
                        'End If
                        'If FlyerTabMatrix.Rows.Count > 0 Then
                        '    UploadFromDataTableFlyerTab(FlyerTabMatrix)
                        'End If
                        'If TemporayTabMatrix.Rows.Count > 0 Then
                        '    UploadFromDataTableTemporayTab(TemporayTabMatrix)
                        'End If

                        'If dEDSPFromMatrix.Rows.Count > 0 Then
                        '    UploadFromDataTableEDSP(dEDSPFromMatrix)
                        'End If

                        'If dtBCUFromMatrix.Rows.Count > 0 Then
                        '    UploadFromDataTableBCU(dtBCUFromMatrix)
                        'End If

                        If dtFromMatrix.Rows.Count > 0 Then

                            UploadFromDataTable(dtFromMatrix)

                        ElseIf FlyerTabMatrix.Rows.Count > 0 Then

                            UploadFromDataTableFlyerTab(FlyerTabMatrix)

                        ElseIf TemporayTabMatrix.Rows.Count > 0 Then

                            UploadFromDataTableTemporayTab(TemporayTabMatrix)

                        ElseIf dEDSPFromMatrix.Rows.Count > 0 Then

                            UploadFromDataTableEDSP(dEDSPFromMatrix)
                        ElseIf dtBCUFromMatrix.Rows.Count > 0 Then
                            UploadFromDataTableBCU(dtBCUFromMatrix)
                        ElseIf dtBPUFromMatrix.Rows.Count > 0 Then
                            UploadFromDataTableBPU(dtBPUFromMatrix)

                        Else

                            objMain.objApplication.StatusBar.SetText("Please Copy and paste data into the matrix before clicking Load Button.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        End If

                    End If

                    If pVal.ItemUID = "LoadMtx" AndAlso pVal.BeforeAction = True AndAlso (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then

                        ' MOP Tab
                        Dim dtFromMatrix As DataTable = GetDataFromMatrix()

                        'Every day Special Price Tab
                        Dim dEDSPFromMatrix As DataTable = GetDataFromMatrixEDSP()

                        'Bulk Cost Upload Tab
                        Dim dtBCUFromMatrix As DataTable = GetDataFromMatrixBCU()

                        ' Flyer Tab
                        Dim FlyerTabMatrix As DataTable = GetDataFromMatrixFlyer()

                        ' Temporary Tab
                        Dim TemporayTabMatrix As DataTable = GetDataFromMatrixTempoary()

                        'Bulk Price Upload Tab
                        Dim dtBPUFromMatrix As DataTable = GetDataFromMatrixBPU()


                        If dtFromMatrix.Rows.Count > 0 Then

                            UploadFromDataTable1(dtFromMatrix)

                        ElseIf FlyerTabMatrix.Rows.Count > 0 Then

                            UploadFromDataTableFlyerTab1(FlyerTabMatrix)

                        ElseIf TemporayTabMatrix.Rows.Count > 0 Then

                            UploadFromDataTableTemporayTab1(TemporayTabMatrix)

                        ElseIf dEDSPFromMatrix.Rows.Count > 0 Then

                            UploadFromDataTableEDSP1(dEDSPFromMatrix)
                        ElseIf dtBCUFromMatrix.Rows.Count > 0 Then
                            UploadFromDataTableBCU1(dtBCUFromMatrix)
                        ElseIf dtBPUFromMatrix.Rows.Count > 0 Then
                            UploadFromDataTableBPU1(dtBPUFromMatrix)

                        Else

                            objMain.objApplication.StatusBar.SetText("Please Copy and paste data into the matrix before clicking Load Button.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        End If

                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        End If

                    End If

   '                 If pVal.ItemUID = "LoadMtx" AndAlso
   'pVal.BeforeAction = True AndAlso
   'pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then

   '                     Try

   '                         ' Get ONLY newly added Flyer rows
   '                         Dim FlyerTabMatrix As DataTable = GetNewDataFromMatrixFlyer()

   '                         If FlyerTabMatrix.Rows.Count > 0 Then

   '                             ' Process only new rows
   '                             UploadFromDataTableFlyerTab(FlyerTabMatrix)

   '                         Else

   '                             objMain.objApplication.StatusBar.SetText(
   '                                 "No new Flyer rows found to load.",
   '                                 SAPbouiCOM.BoMessageTime.bmt_Short,
   '                                 SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

   '                         End If

   '                     Catch ex As Exception

   '                         objMain.objApplication.StatusBar.SetText(
   '                             "Update Load Error: " & ex.Message,
   '                             SAPbouiCOM.BoMessageTime.bmt_Short,
   '                             SAPbouiCOM.BoStatusBarMessageType.smt_Error)

   '                     End Try

   '                 End If

                    'If pVal.ItemUID = "Submit" And pVal.BeforeAction = False And objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    '    Dim result As Integer = objMain.objApplication.MessageBox("You are going to update item prices. Do you want to continue?", 1, "Continue", "Cancel")
                    '    Dim oRsPriceList As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '    Dim oRsSubmit As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    '    If (result = 1) Then

                    '        objMatrix3 = objForm.Items.Item("Mtx3").Specific
                    '        Me.StartBackgroundForPriceUpdateService()

                    '    End If
                    'End If


                Case SAPbouiCOM.BoEventTypes.et_CLICK

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        'If Me.Validate() = False Then

                        '    BubbleEvent = False
                        'End If
                        If Not ValidateBarcodeByTab(objForm) Then
                            BubbleEvent = False
                            Exit Sub

                        End If

                    End If


                    'Link button in martix 4 Temporay price reduction
                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                    If pVal.ItemUID = "Mtx4" And pVal.ColUID = "PDN" And pVal.Before_Action = False Then
                        objMatrix4 = objForm.Items.Item("Mtx4").Specific

                        Dim DocumentNumber As String = oDBs_Head.GetValue("DocNum", 0).Trim()
                        objMain.objApplication.ActivateMenuItem("ME_ITEMA")
                        objNewForm = objMain.objApplication.Forms.GetForm("ITEM", objMain.objApplication.Forms.ActiveForm.TypeCount)
                        objNewForm.Freeze(True)
                        objNewForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objNewForm.Items.Item("PCRDoc").Specific.Value = DocumentNumber
                        objNewForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        objNewForm.Freeze(False)

                        'Link button in martix 3 MOP
                    ElseIf pVal.ItemUID = "Mtx3" And pVal.ColUID = "PDN" And pVal.Before_Action = False Then
                        objMatrix3 = objForm.Items.Item("Mtx3").Specific

                        Dim DocumentNumber As String = oDBs_Head.GetValue("DocNum", 0).Trim()
                        objMain.objApplication.ActivateMenuItem("ME_ITEMA")
                        objNewForm = objMain.objApplication.Forms.GetForm("ITEM", objMain.objApplication.Forms.ActiveForm.TypeCount)
                        objNewForm.Freeze(True)
                        objNewForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objNewForm.Items.Item("PCRDoc").Specific.Value = DocumentNumber
                        objNewForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        objNewForm.Freeze(False)

                        'Link button in martix 2 Flyer price change

                    ElseIf pVal.ItemUID = "Mtx2" And pVal.ColUID = "PDN" And pVal.Before_Action = False Then
                        objMatrix2 = objForm.Items.Item("Mtx2").Specific

                        Dim DocumentNumber As String = oDBs_Head.GetValue("DocNum", 0).Trim()
                        objMain.objApplication.ActivateMenuItem("ME_ITEMA")
                        objNewForm = objMain.objApplication.Forms.GetForm("ITEM", objMain.objApplication.Forms.ActiveForm.TypeCount)
                        objNewForm.Freeze(True)
                        objNewForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objNewForm.Items.Item("PCRDoc").Specific.Value = DocumentNumber
                        objNewForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        objNewForm.Freeze(False)
                    ElseIf pVal.ItemUID = "Mtx6" And pVal.ColUID = "PDN" And pVal.Before_Action = False Then
                        objMatrix2 = objForm.Items.Item("Mtx6").Specific

                        Dim DocumentNumber As String = oDBs_Head.GetValue("DocNum", 0).Trim()
                        objMain.objApplication.ActivateMenuItem("ME_ITEMA")
                        objNewForm = objMain.objApplication.Forms.GetForm("ITEM", objMain.objApplication.Forms.ActiveForm.TypeCount)
                        objNewForm.Freeze(True)
                        objNewForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        objNewForm.Items.Item("PCRDoc").Specific.Value = DocumentNumber
                        objNewForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        objNewForm.Freeze(False)

                    End If

                                      'madhu end

                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    If pVal.ItemUID = "Mtx5" And pVal.ColUID = "TPath" And pVal.BeforeAction = False Then
                        objForm = objMain.objApplication.Forms.Item(FormUID)
                        objMatrix5 = objForm.Items.Item("Mtx5").Specific
                        Dim fullPath As String = objMatrix5.Columns.Item("TPath").Cells.Item(pVal.Row).Specific.Value
                        Dim indexLoc As Integer = fullPath.LastIndexOf("\")  ''-1
                        Dim path As String = fullPath.Substring(0, indexLoc)
                        Dim filename As String = fullPath.Substring(indexLoc + 1, fullPath.Length - indexLoc - 1)
                        objMatrix5.Columns.Item("FName").Cells.Item(pVal.Row).Specific.Value = filename
                        objMatrix5.Columns.Item("ADate").Cells.Item(pVal.Row).Specific.Value() = DateTime.Now.ToString("yyyyMMdd")
                        If pVal.Row = objMatrix5.VisualRowCount Then
                            Me.SetNewLine5(objForm.UniqueID)
                            objMatrix5.AutoResizeColumns()
                        End If
                    End If





                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS


                    'Cost Fluctuation Stock Transfer calculation
                    If pVal.ItemUID = "Mtx8" And pVal.ColUID = "Col_12" And pVal.BeforeAction = False Then
                        Try

                            objForm = objMain.objApplication.Forms.Item(FormUID)


                            objMatrix8 = CType(objForm.Items.Item("Mtx8").Specific, SAPbouiCOM.Matrix)
                            Dim NEWP As Double = 0
                            Dim NEWMAC As Double = 0
                            Dim newVALUE As Double = 0
                            Dim NMNS As Double = 0


                            Double.TryParse(objMatrix8.Columns.Item("Col_12").Cells.Item(pVal.Row).Specific.Value, NEWP)
                            Double.TryParse(objMatrix8.Columns.Item("Col_6").Cells.Item(pVal.Row).Specific.Value, NEWMAC)

                            newVALUE = (NEWP - NEWMAC)
                            NMNS = (newVALUE / NEWP) * 100

                            objMatrix8.Columns.Item("Col_13").Cells.Item(pVal.Row).Specific.Value = Math.Round(newVALUE, 3).ToString()
                            objMatrix8.Columns.Item("Col_14").Cells.Item(pVal.Row).Specific.Value = Math.Round(NMNS, 3).ToString()



                        Catch ex As Exception

                        End Try
                    End If

                    'Cost Fluctuction calculation
                    If pVal.ItemUID = "Mtx1" AndAlso pVal.ColUID = "NRP" AndAlso pVal.BeforeAction = False Then

                        Dim oMatrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx1").Specific, SAPbouiCOM.Matrix)

                        Dim row As Integer = pVal.Row

                        Dim NPS As Double = 0
                        Dim SupplierCost As Double = 0
                        Dim OldLPP As Double = 0

                        Double.TryParse(CType(oMatrix.Columns.Item("NRP").Cells.Item(row).Specific, SAPbouiCOM.EditText).Value, NPS)

                        Double.TryParse(CType(oMatrix.Columns.Item("SCOST").Cells.Item(row).Specific, SAPbouiCOM.EditText).Value, SupplierCost)

                        Double.TryParse(CType(oMatrix.Columns.Item("ORP").Cells.Item(row).Specific, SAPbouiCOM.EditText).Value, OldLPP)

                        If NPS <= 0 Then Exit Sub   ' avoid divide by zero

                        '---------------- CALCULATIONS ----------------
                        Dim NewMargin As Double = ((NPS - SupplierCost) / NPS) * 100

                        Dim OldMargin As Double = ((NPS - OldLPP) / NPS) * 100

                        '---------------- SET VALUES ----------------
                        '              CType(oMatrix.Columns.Item("NewLPP").Cells.Item(row).Specific,
                        'SAPbouiCOM.EditText).Value = SupplierCost.ToString("F2")

                        '              CType(oMatrix.Columns.Item("OldMar").Cells.Item(row).Specific,
                        'SAPbouiCOM.EditText).Value = OldMargin.ToString("F2")

                        '              CType(oMatrix.Columns.Item("NewMar").Cells.Item(row).Specific,
                        'SAPbouiCOM.EditText).Value = NewMargin.ToString("F2")

                    End If

                    'calculations for mop tab new margin%
                    If pVal.ItemUID = "Mtx3" AndAlso pVal.ColUID = "NSPrice" AndAlso pVal.BeforeAction = False Then
                        Try

                            objForm = objMain.objApplication.Forms.Item(FormUID)
                            objForm.Freeze(True)
                            objMatrix3 = CType(objForm.Items.Item("Mtx3").Specific, SAPbouiCOM.Matrix)

                            Dim newSellPrice As Double = 0
                            Dim lastGRPOPrice As Double = 0
                            Dim newMargin As Double = 0


                            Double.TryParse(objMatrix3.Columns.Item("NSPrice").Cells.Item(pVal.Row).Specific.Value, newSellPrice)
                            Dim itemCode As String = objMatrix3.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.Trim()


                            If newSellPrice <= 0 OrElse itemCode = "" Then
                                Dim oDBS0 As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C2")
                                oDBS0.SetValue("U_NMarg", pVal.Row - 1, "0")
                                Exit Try
                            End If


                            Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'Dim sql As String = "SELECT ""PricePerBaseQty"" FROM (" & " SELECT T1.""ItemCode"", T1.""Price"", " & " ROW_NUMBER() OVER (PARTITION BY T1.""ItemCode"" " &
                            '              " ORDER BY T0.""DocDate"" DESC, T0.""DocEntry"" DESC, T1.""LineNum"" DESC) RN " & " FROM ""OPDN"" T0 " & " INNER JOIN ""PDN1"" T1 ON T1.""DocEntry"" = T0.""DocEntry"" " &
                            '                        " WHERE T0.""CANCELED"" = 'N' " & ") X WHERE X.""ItemCode"" = '" & itemCode.Replace("'", "''") & "' AND RN = 1"
                            Dim sql As String = "SELECT (X.""Price"" / NULLIF(X.""NumPerMsr"", 0)) AS ""PricePerBaseQty"" " & "FROM ( " &
                                               " SELECT " & " T1.""ItemCode"", " & " T1.""Price"", " & " T1.""NumPerMsr"", " & " ROW_NUMBER() OVER ( " & " PARTITION BY T1.""ItemCode"" " &
                                               " ORDER BY T0.""DocDate"" DESC, " & "  T0.""DocEntry"" DESC, " & "  T1.""LineNum"" DESC " & " ) AS RN " & " FROM ""OPDN"" T0 " &
                                               " INNER JOIN ""PDN1"" T1 ON T1.""DocEntry"" = T0.""DocEntry"" " &
                                               " WHERE T0.""CANCELED"" = 'N' " & ") X " & "WHERE X.""ItemCode"" = '" & itemCode.Replace("'", "''") & "' " & "AND X.RN = 1"

                            oRS.DoQuery(sql)

                            If Not oRS.EoF Then
                                Double.TryParse(oRS.Fields.Item("PricePerBaseQty").Value.ToString(), lastGRPOPrice)
                            End If

                            Dim NSPrice As Double = objMatrix3.Columns.Item("NSPrice").Cells.Item(pVal.Row).Specific.Value.Trim()
                            Dim CCost As Double = objMatrix3.Columns.Item("CCost").Cells.Item(pVal.Row).Specific.Value.Trim()

                            newMargin = ((NSPrice - CCost) / NSPrice) * 100
                            'newMargin = (newSellPrice - lastGRPOPrice) / newSellPrice

                            Dim oDBS As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C2")
                            oDBS.SetValue("U_NMarg", pVal.Row - 1, Math.Round(newMargin, 2).ToString())
                            objMatrix3.Columns.Item("NMarg").Cells.Item(pVal.Row).Specific.Value = Math.Round(newMargin, 2).ToString()

                            ' Optional: trigger click on last row CName
                            ' objMatrix3.Columns.Item("CName").Cells.Item(objMatrix3.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            objMatrix3.Columns.Item("CName").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)


                        Catch ex As Exception

                            objMain.objApplication.StatusBar.SetText("Error in margin calculation: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        Finally

                            If Not objForm Is Nothing Then objForm.Freeze(False)
                        End Try
                    End If

                    'Calculations for every day special price tab new margin%
                    If pVal.ItemUID = "Mtx6" AndAlso pVal.ColUID = "Col_8" AndAlso pVal.BeforeAction = False Then
                        Try

                            objForm = objMain.objApplication.Forms.Item(FormUID)
                            objForm.Freeze(True)
                            objMatrix6 = CType(objForm.Items.Item("Mtx6").Specific, SAPbouiCOM.Matrix)


                            Dim newMargin As Double = 0


                            Dim NSPrice As Double = objMatrix6.Columns.Item("Col_8").Cells.Item(pVal.Row).Specific.Value.Trim()
                            Dim CCost As Double = objMatrix6.Columns.Item("Col_7").Cells.Item(pVal.Row).Specific.Value.Trim()

                            newMargin = ((NSPrice - CCost) / NSPrice) * 100

                            objMatrix6.Columns.Item("Col_9").Cells.Item(pVal.Row).Specific.Value = Math.Round(newMargin, 2).ToString()
                            objMatrix6.Columns.Item("Col_0").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                        Catch ex As Exception

                            objMain.objApplication.StatusBar.SetText("Error in margin calculation: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        Finally

                            If Not objForm Is Nothing Then objForm.Freeze(False)
                        End Try
                    End If

                    'Flyer price changes calculation

                    If pVal.ItemUID = "Mtx2" AndAlso pVal.ColUID = "NwPrice" AndAlso pVal.BeforeAction = False Then
                        Try
                            objForm = objMain.objApplication.Forms.Item(FormUID)
                            objForm.Freeze(True)

                            Dim objMatrix2 As SAPbouiCOM.Matrix
                            objMatrix2 = CType(objForm.Items.Item("Mtx2").Specific, SAPbouiCOM.Matrix)

                            Dim newSellPrice As Double = 0
                            Dim lastpurchaseprice As Double = 0
                            Dim newMargin As Double = 0


                            ' Read values from matrix
                            Double.TryParse(objMatrix2.Columns.Item("NwPrice").Cells.Item(pVal.Row).Specific.Value, newSellPrice)
                            Double.TryParse(objMatrix2.Columns.Item("LPrice").Cells.Item(pVal.Row).Specific.Value, lastpurchaseprice)


                            ' Validation
                            If newSellPrice <= 0 OrElse lastpurchaseprice <= 0 Then
                                Dim oDBS0 As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C1")
                                oDBS0.SetValue("U_Marg", pVal.Row - 1, "0")
                                Exit Try
                            End If

                            ' Margin calculation (NO QUERY)
                            newMargin = ((newSellPrice - lastpurchaseprice) / newSellPrice) * 100


                            ' Update DB datasource
                            Dim oDBS As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C1")
                            oDBS.SetValue("U_Marg", pVal.Row - 1, Math.Round(newMargin, 2).ToString())

                            ' Update matrix cell
                            objMatrix2.Columns.Item("Marg").Cells.Item(pVal.Row).Specific.Value = Math.Round(newMargin, 2).ToString()


                            ' Optional focus trigger
                            objMatrix2.Columns.Item("PUS").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText("Error in Mtx2 margin calculation: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Finally
                            If objForm IsNot Nothing Then objForm.Freeze(False)
                        End Try
                    End If

                    'If pVal.ItemUID = "Mtx2" AndAlso pVal.ColUID = "Col_1" AndAlso pVal.BeforeAction = False Then

                    '    Try
                    '        objForm = objMain.objApplication.Forms.Item(FormUID)
                    '        objForm.Freeze(True)

                    '        Dim objMatrix2 As SAPbouiCOM.Matrix
                    '        objMatrix2 = CType(objForm.Items.Item("Mtx2").Specific, SAPbouiCOM.Matrix)

                    '        Dim Cost As Double = 0
                    '        Dim Normal As Double = 0
                    '        Dim Flyer As Double = 0
                    '        Dim NormalGP As Double = 0
                    '        Dim FlyerGP As Double = 0



                    '        Double.TryParse(objMatrix2.Columns.Item("Cost").Cells.Item(pVal.Row).Specific.Value, Cost)
                    '        Double.TryParse(objMatrix2.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific.Value, Normal)
                    '        Double.TryParse(objMatrix2.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific.Value, Flyer)



                    '        ''NormalGP = ((Normal - Cost) / Normal) * 100
                    '        FlyerGP = ((Flyer - Cost) / Flyer) * 100


                    '        ' Update DB datasource
                    '        Dim oDBS As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C1")


                    '        ''objMatrix2.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific.Value = Math.Round(NormalGP, 2).ToString()
                    '        objMatrix2.Columns.Item("Col_3").Cells.Item(pVal.Row).Specific.Value = Math.Round(FlyerGP, 2).ToString()

                    '        objMatrix2.Columns.Item("Col_4").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)


                    '    Catch ex As Exception
                    '        objMain.objApplication.StatusBar.SetText("Error in Mtx2 margin calculation: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    Finally
                    '        If objForm IsNot Nothing Then objForm.Freeze(False)
                    '    End Try
                    'End If

                    If pVal.ItemUID = "Mtx2" AndAlso pVal.ColUID = "Col_1" AndAlso pVal.BeforeAction = False Then

                        Try
                            objForm = objMain.objApplication.Forms.Item(FormUID)
                            objForm.Freeze(True)

                            Dim objMatrix2 As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx2").Specific, SAPbouiCOM.Matrix)

                            Dim Cost As Double = 0
                            Dim Flyer As Double = 0
                            Dim FlyerGP As Double = 0

                            Double.TryParse(CType(objMatrix2.Columns.Item("Cost").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value, Cost)
                            Double.TryParse(CType(objMatrix2.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value, Flyer)

                            If Flyer <> 0 Then
                                FlyerGP = ((Flyer - Cost) / Flyer) * 100
                            Else
                                FlyerGP = 0
                            End If

                            CType(objMatrix2.Columns.Item("Col_3").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = Math.Round(FlyerGP, 2).ToString("0.00")
                            'CalculateTotalGP(objForm.UniqueID)
                            objMatrix2.Columns.Item("Col_4").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText("Error in Flyer GP calculation: " & ex.Message,
                                                 SAPbouiCOM.BoMessageTime.bmt_Short,
                                                 SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Finally
                            objForm.Freeze(False)
                        End Try

                    End If

                    If pVal.ItemUID = "Mtx2" AndAlso pVal.ColUID = "Col_4" AndAlso pVal.BeforeAction = False Then

                        Try
                            objForm = objMain.objApplication.Forms.Item(FormUID)
                            objForm.Freeze(True)

                            Dim objMatrix2 As SAPbouiCOM.Matrix
                            objMatrix2 = CType(objForm.Items.Item("Mtx2").Specific, SAPbouiCOM.Matrix)


                            Dim Cost As Double = 0
                            Dim Normal As Double = 0
                            Dim Flyer As Double = 0
                            Dim NormalGP As Double = 0
                            Dim FlyerGP As Double = 0
                            Dim Exp As Double = 0
                            Dim Actual As Double = 0
                            Dim ActualSelling As Double = 0
                            Dim ActualFSelling As Double = 0
                            Dim NGPV As Double = 0
                            Dim FLYERGPV As Double = 0
                            Dim ProfitLoss As Double = 0


                            Double.TryParse(objMatrix2.Columns.Item("Cost").Cells.Item(pVal.Row).Specific.Value, Cost)
                            Double.TryParse(objMatrix2.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific.Value, Normal)
                            Double.TryParse(objMatrix2.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific.Value, Flyer)
                            Double.TryParse(objMatrix2.Columns.Item("Col_4").Cells.Item(pVal.Row).Specific.Value, Exp)


                            NormalGP = ((Normal - Cost) / Normal) * 100
                            FlyerGP = ((Flyer - Cost) / Flyer) * 100
                            Actual = Cost * Exp
                            ActualSelling = Exp * Normal
                            ActualFSelling = Exp * Flyer
                            NGPV = ActualSelling - Actual
                            FLYERGPV = ActualFSelling - Actual
                            ProfitLoss = ActualFSelling - ActualSelling

                            ' Update DB datasource
                            Dim oDBS As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C1")


                            'objMatrix2.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific.Value = Math.Round(NormalGP, 2).ToString()
                            'objMatrix2.Columns.Item("Col_3").Cells.Item(pVal.Row).Specific.Value = Math.Round(FlyerGP, 2).ToString()
                            objMatrix2.Columns.Item("Col_5").Cells.Item(pVal.Row).Specific.Value = Math.Round(Actual, 2).ToString()
                            objMatrix2.Columns.Item("Col_6").Cells.Item(pVal.Row).Specific.Value = Math.Round(ActualSelling, 2).ToString()
                            objMatrix2.Columns.Item("Col_7").Cells.Item(pVal.Row).Specific.Value = Math.Round(ActualFSelling, 2).ToString()
                            objMatrix2.Columns.Item("Col_8").Cells.Item(pVal.Row).Specific.Value = Math.Round(NGPV, 2).ToString()
                            objMatrix2.Columns.Item("Col_9").Cells.Item(pVal.Row).Specific.Value = Math.Round(FLYERGPV, 2).ToString()
                            objMatrix2.Columns.Item("Col_10").Cells.Item(pVal.Row).Specific.Value = Math.Round(ProfitLoss, 2).ToString()

                            objMatrix2.Columns.Item("PUS").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText("Error in Mtx2 margin calculation: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Finally
                            If objForm IsNot Nothing Then objForm.Freeze(False)
                        End Try
                    End If

                    'Temporay price Reduction TPR calculation

                    If pVal.ItemUID = "Mtx4" AndAlso pVal.ColUID = "SPrice" AndAlso pVal.BeforeAction = False Then

                        Try
                            objForm = objMain.objApplication.Forms.Item(FormUID)
                            objForm.Freeze(True)

                            objMatrix4 = CType(objForm.Items.Item("Mtx4").Specific, SAPbouiCOM.Matrix)

                            Dim sp As Double = 0        ' Special Price
                            Dim rsp As Double = 0       ' Normal Selling Price
                            Dim cost As Double = 0      ' Branch Avg Cost
                            Dim ItemCost As Double = 0
                            Dim ICost As Double = 0
                            Dim NSell As Double = 0
                            Dim TPRC As Double = 0
                            Dim TPRS As Double = 0
                            Dim QTY As Double = 0

                            Dim diffSPCost As Double = 0
                            Dim diffSPNS As Double = 0
                            Dim normalGP As Double = 0
                            Dim specialGP As Double = 0

                            Dim itemCode As String = objMatrix4.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.Trim()

                            Double.TryParse(objMatrix4.Columns.Item("SPrice").Cells.Item(pVal.Row).Specific.Value, sp)

                            Double.TryParse(objMatrix4.Columns.Item("NSell").Cells.Item(pVal.Row).Specific.Value, rsp)

                            If itemCode = "" OrElse sp <= 0 Then Exit Try

                            Dim BPLID As String = CType(objForm.Items.Item("TNXBRC").Specific, SAPbouiCOM.ComboBox).Value.Trim()

                            Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            Dim sqlCost As String = "SELECT MAX(T1.""AvgPrice"") AS ""COST"" " & "FROM ""OWHS"" T0 JOIN ""OITW"" T1 ON T0.""WhsCode"" = T1.""WhsCode"" " &
                                                     "WHERE T1.""ItemCode"" = '" & itemCode & "' " & "AND T0.""BPLid"" = '" & BPLID & "'"

                            oRS.DoQuery(sqlCost)
                            If Not oRS.EoF AndAlso oRS.Fields.Item("COST").Value IsNot Nothing Then
                                Double.TryParse(oRS.Fields.Item("COST").Value.ToString(), cost)
                            End If


                            Double.TryParse(objMatrix4.Columns.Item("SPrice").Cells.Item(pVal.Row).Specific.Value, ItemCost)
                            Double.TryParse(objMatrix4.Columns.Item("ICost").Cells.Item(pVal.Row).Specific.Value, ICost)
                            Double.TryParse(objMatrix4.Columns.Item("NSell").Cells.Item(pVal.Row).Specific.Value, NSell)
                            Double.TryParse(objMatrix4.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value, QTY)


                            diffSPCost = ItemCost - ICost
                            diffSPNS = ItemCost - NSell
                            TPRC = (ItemCost - ICost) * QTY
                            TPRS = (ItemCost - NSell) * QTY

                            ' ✅ NORMAL GP (FIXED & SAFE)
                            If rsp > 0 AndAlso cost > 0 Then
                                normalGP = ((rsp - cost) / rsp) * 100
                            Else
                                normalGP = 0
                            End If

                            ' SPECIAL GP
                            If sp > 0 AndAlso cost > 0 Then
                                specialGP = ((sp - cost) / sp) * 100
                            Else
                                specialGP = 0
                            End If

                            diffSPCost = Math.Round(diffSPCost, 2)
                            diffSPNS = Math.Round(diffSPNS, 2)
                            normalGP = Math.Round(normalGP, 2)
                            specialGP = Math.Round(specialGP, 2)

                            ' -------------------------------
                            ' UPDATE DATASOURCE + MATRIX
                            ' -------------------------------
                            Dim oDBS As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C3")

                            oDBS.SetValue("U_DiffSP", pVal.Row - 1, diffSPCost.ToString())
                            oDBS.SetValue("U_DSPNS", pVal.Row - 1, diffSPNS.ToString())
                            oDBS.SetValue("U_NGP", pVal.Row - 1, normalGP.ToString())
                            oDBS.SetValue("U_SGP", pVal.Row - 1, specialGP.ToString())
                            oDBS.SetValue("U_TPRC", pVal.Row - 1, TPRC.ToString())
                            oDBS.SetValue("U_TPRS", pVal.Row - 1, TPRS.ToString())

                            objMatrix4.Columns.Item("DiffSP").Cells.Item(pVal.Row).Specific.Value = diffSPCost
                            objMatrix4.Columns.Item("DSPNS").Cells.Item(pVal.Row).Specific.Value = diffSPNS
                            objMatrix4.Columns.Item("NGP").Cells.Item(pVal.Row).Specific.Value = normalGP
                            objMatrix4.Columns.Item("SGP").Cells.Item(pVal.Row).Specific.Value = specialGP
                            objMatrix4.Columns.Item("TPRC").Cells.Item(pVal.Row).Specific.Value = TPRC
                            objMatrix4.Columns.Item("TPRS").Cells.Item(pVal.Row).Specific.Value = TPRS

                            ' 🔁 KEEP EXISTING FUNCTIONALITY
                            objMatrix4.Columns.Item("Supp").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText("Temporary Price Reduction Calculation error : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Finally
                            objForm.Freeze(False)
                        End Try
                    End If



                    'MOP tab copy and paste

                    If pVal.ItemUID = "Mtx3" And pVal.ColUID = "BCode" Or pVal.ColUID = "ItemCode" And pVal.BeforeAction = False Then
                        Try
                            objForm.Freeze(True)
                            Dim objMatrix3 As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx3").Specific
                            If pVal.ItemUID = "Mtx3" Then
                                If pVal.Row = objMatrix3.VisualRowCount Then
                                    'Me.SetNewLine3(objForm.UniqueID)
                                End If
                            End If
                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                        End Try
                    End If


                    'Every day special price tab copy and paste
                    If pVal.ItemUID = "Mtx6" And pVal.ColUID = "Col_0" Or pVal.ColUID = "Col_1" And pVal.BeforeAction = False Then
                        Try
                            objForm.Freeze(True)
                            Dim objMatrix6 As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx6").Specific
                            If pVal.ItemUID = "Mtx6" Then
                                If pVal.Row = objMatrix6.VisualRowCount Then
                                    'Me.SetNewLine3(objForm.UniqueID)
                                End If
                            End If
                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                        End Try
                    End If


                    'Flyer tab copy and paste
                    If pVal.ItemUID = "Mtx2" And pVal.ColUID = "BrCode" Or pVal.ColUID = "ItemCode" And pVal.BeforeAction = False Then
                        Try
                            objForm.Freeze(True)
                            Dim objMatrix3 As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx2").Specific
                            If pVal.ItemUID = "Mtx2" Then
                                If pVal.Row = objMatrix3.VisualRowCount Then
                                    'Me.SetNewLine3(objForm.UniqueID)
                                End If
                            End If
                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                        End Try
                    End If

                    'Temporary Price Reduction Tab
                    If pVal.ItemUID = "Mtx4" And pVal.ColUID = "BrCode" Or pVal.ColUID = "ItemCode" And pVal.BeforeAction = False Then
                        Try
                            objForm.Freeze(True)
                            Dim objMatrix3 As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx2").Specific
                            If pVal.ItemUID = "Mtx4" Then
                                If pVal.Row = objMatrix3.VisualRowCount Then
                                    'Me.SetNewLine3(objForm.UniqueID)
                                End If
                            End If
                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                        End Try
                    End If

                    'Cost Fluctuation tab in that new retail price with vat calculation 
                    'If pVal.ItemUID = "Mtx1" And pVal.ColUID = "NRP" And pVal.BeforeAction = False Then
                    '    Dim objMatrix1 As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx1").Specific
                    '    Dim rs As SAPbobsCOM.Recordset
                    '    rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '    objMatrix1.FlushToDataSource()
                    '    For i As Integer = 1 To objMatrix1.RowCount
                    '        Dim priceWithVAT As Double = 0
                    '        Dim priceWithoutVAT As Double = 0
                    '        Dim vatPercent As Double = 0
                    '        Dim val As String = objMatrix1.Columns.Item("NRP").Cells.Item(i).Specific.Value
                    '        If val <> "" Then
                    '            priceWithVAT = CDbl(val)
                    '        End If

                    '        Dim itemCode As String = objMatrix1.Columns.Item("TNXICD").Cells.Item(i).Specific.Value
                    '        If itemCode <> "" Then

                    '            Dim query As String = "SELECT T1.""Rate"" FROM ""OITM"" T0 " &
                    '             "INNER JOIN ""OVTG"" T1 ON T0.""VatGourpSa"" = T1.""Code"" " &
                    '             "WHERE T0.""ItemCode"" = '" & itemCode & "'"

                    '            rs.DoQuery(query)

                    '            If rs.RecordCount > 0 Then
                    '                vatPercent = CDbl(rs.Fields.Item("Rate").Value)
                    '            End If
                    '        End If

                    '        If vatPercent > 0 Then
                    '            priceWithoutVAT = priceWithVAT / (1 + vatPercent / 100)
                    '        Else
                    '            priceWithoutVAT = priceWithVAT
                    '        End If

                    '        'disabled field code here..!
                    '        objForm = objMain.objApplication.Forms.Item(FormUID)

                    '        Try
                    '            objForm.Freeze(True)

                    '            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C0")
                    '            objMatrix1 = objForm.Items.Item("Mtx1").Specific
                    '            Dim value As String = Math.Round(priceWithoutVAT, 2).ToString()
                    '            oDBs_Details1.SetValue("U_NRPVAT", i, value)
                    '            objMatrix1.SetLineData(i)

                    '        Finally
                    '            objForm.Freeze(False)
                    '        End Try
                    '    Next
                    'End If



                    ''NEW
                    If pVal.ItemUID = "Mtx1" And pVal.ColUID = "NRP" And pVal.BeforeAction = False Then
                        Dim objMatrix1 As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx1").Specific
                        Dim rs As SAPbobsCOM.Recordset
                        rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objMatrix1.FlushToDataSource()
                        Dim i As Integer = pVal.Row
                        Dim priceWithVAT As Double = 0
                        Dim priceWithoutVAT As Double = 0
                        Dim vatPercent As Double = 0
                        Dim val As String = objMatrix1.Columns.Item("NRP").Cells.Item(i).Specific.Value
                        If val <> "" Then
                            priceWithVAT = CDbl(val)
                        End If

                        Dim itemCode As String = objMatrix1.Columns.Item("TNXICD").Cells.Item(i).Specific.Value
                        If itemCode <> "" Then

                            Dim query As String = "SELECT T1.""Rate"" FROM ""OITM"" T0 " &
                                 "INNER JOIN ""OVTG"" T1 ON T0.""VatGourpSa"" = T1.""Code"" " &
                                 "WHERE T0.""ItemCode"" = '" & itemCode & "'"

                            rs.DoQuery(query)

                            If rs.RecordCount > 0 Then
                                vatPercent = CDbl(rs.Fields.Item("Rate").Value)
                            End If
                        End If

                        If vatPercent > 0 Then
                            priceWithoutVAT = priceWithVAT / (1 + vatPercent / 100)
                        Else
                            priceWithoutVAT = priceWithVAT
                        End If

                        'disabled field code here..!
                        objForm = objMain.objApplication.Forms.Item(FormUID)

                        Try
                            objForm.Freeze(True)

                            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C0")
                            objMatrix1 = objForm.Items.Item("Mtx1").Specific
                            Dim value As String = Math.Round(priceWithoutVAT, 2).ToString()
                            objMatrix1.Columns.Item("NRPVAT").Cells.Item(i).Specific.Value = value
                            ''oDBs_Details1.SetValue("U_NRPVAT", i, value)
                            ''objMatrix1.SetLineData(i)
                            ''objMatrix1.Columns.Item("PUS").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                        Finally
                            objForm.Freeze(False)
                        End Try

                    End If

                    'Cost Fluctuation tab in that new margin percentage calculation
                    'If pVal.ItemUID = "Mtx1" And pVal.ColUID = "NRP" And pVal.BeforeAction = False Then

                    '    Dim objMatrix1 As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx1").Specific
                    '    objMatrix1.FlushToDataSource()

                    '    objForm = objMain.objApplication.Forms.Item(FormUID)

                    '    Try
                    '        objForm.Freeze(True)

                    '        oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C0")
                    '        objMatrix1 = objForm.Items.Item("Mtx1").Specific

                    '        For i As Integer = 1 To objMatrix1.RowCount

                    '            Dim RetailpricewithoutVAT_ORPVAT As Double = 0   'Retail price without VAT
                    '            Dim LatestLastpurchaseprice_ULPRC As Double = 0    'LatestLast purchase price
                    '            Dim marginPercent As Double = 0

                    '            '--- Get New Retail Price without VAT
                    '            Dim nrpvatStr As String = oDBs_Details1.GetValue("U_NRPVAT", i - 1).Trim()
                    '            If nrpvatStr <> "" Then
                    '                RetailpricewithoutVAT_ORPVAT = CDbl(nrpvatStr)
                    '            End If

                    '            '--- Get Lastest Last Purchase Price 
                    '            Dim ulprcStr As String = objMatrix1.Columns.Item("ULPRC").Cells.Item(i).Specific.Value
                    '            If ulprcStr <> "" Then
                    '                LatestLastpurchaseprice_ULPRC = CDbl(ulprcStr)
                    '            End If

                    '            '--- Calculate Margin %
                    '            If RetailpricewithoutVAT_ORPVAT > 0 Then
                    '                marginPercent = ((RetailpricewithoutVAT_ORPVAT - LatestLastpurchaseprice_ULPRC) / RetailpricewithoutVAT_ORPVAT) * 100
                    '            Else
                    '                marginPercent = 0
                    '            End If

                    '            '--- Set Value to U_Margin (create UDF if not exists)
                    '            oDBs_Details1.SetValue("U_NMRG", i - 1, Math.Round(marginPercent, 2).ToString())

                    '            objMatrix1.SetLineData(i)

                    '        Next

                    '    Finally
                    '        objForm.Freeze(False)
                    '    End Try

                    'End If

                    ''NEW 
                    If pVal.ItemUID = "Mtx1" And pVal.ColUID = "NRP" And pVal.BeforeAction = False Then

                        Dim objMatrix1 As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx1").Specific
                        objMatrix1.FlushToDataSource()

                        objForm = objMain.objApplication.Forms.Item(FormUID)

                        Try
                            objForm.Freeze(True)

                            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C0")
                            objMatrix1 = objForm.Items.Item("Mtx1").Specific

                            Dim i As Integer = pVal.Row

                            Dim RetailpricewithoutVAT_ORPVAT As Double = 0   'Retail price without VAT
                            Dim LatestLastpurchaseprice_ULPRC As Double = 0    'LatestLast purchase price
                            Dim marginPercent As Double = 0

                            '--- Get New Retail Price without VAT
                            ''Dim nrpvatStr As String = objMatrix1.Columns.Item("NRPVAT").Cells.Item(i).Specific.Value ''oDBs_Details1.GetValue("U_NRPVAT", i).Trim()
                            Dim nrpvatStr As String = objMatrix1.Columns.Item("NRP").Cells.Item(i).Specific.Value ''oDBs_Details1.GetValue("U_NRPVAT", i).Trim()

                            If nrpvatStr <> "" Then
                                RetailpricewithoutVAT_ORPVAT = CDbl(nrpvatStr)
                            End If

                            '--- Get Lastest Last Purchase Price 
                            ''Dim ulprcStr As String = objMatrix1.Columns.Item("ULPRC").Cells.Item(i).Specific.Value
                            Dim ulprcStr As String = objMatrix1.Columns.Item("ULPRCVAT").Cells.Item(i).Specific.Value
                            If ulprcStr <> "" Then
                                LatestLastpurchaseprice_ULPRC = CDbl(ulprcStr)
                            End If

                            '--- Calculate Margin %
                            If RetailpricewithoutVAT_ORPVAT > 0 Then
                                marginPercent = ((RetailpricewithoutVAT_ORPVAT - LatestLastpurchaseprice_ULPRC) / RetailpricewithoutVAT_ORPVAT) * 100
                            Else
                                marginPercent = 0
                            End If

                            '--- Set Value to U_Margin (create UDF if not exists)
                            'oDBs_Details1.SetValue("U_NMRG", i, Math.Round(marginPercent, 2).ToString())
                            objMatrix1.Columns.Item("NMRG").Cells.Item(i).Specific.Value = Math.Round(marginPercent, 2).ToString()
                            ''objMatrix1.SetLineData(i)

                            objMatrix1.Columns.Item("Col_7").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                        Finally
                            objForm.Freeze(False)
                        End Try

                    End If

                    If pVal.ItemUID = "Mtx7" AndAlso pVal.ColUID = "Base" AndAlso pVal.BeforeAction = False Then

                        Dim objMatrix7 As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx7").Specific, SAPbouiCOM.Matrix)

                        Dim BaseValue As Double = 0
                        Dim TaxValue As Double = 0

                        Double.TryParse(objMatrix7.Columns.Item("Base").Cells.Item(pVal.Row).Specific.Value.ToString(), BaseValue)
                        Double.TryParse(objMatrix7.Columns.Item("Tax").Cells.Item(pVal.Row).Specific.Value.ToString(), TaxValue)

                        If BaseValue <> 0 OrElse TaxValue <> 0 Then
                            objMatrix7.Columns.Item("NPRC").Cells.Item(pVal.Row).Specific.Value = (BaseValue + TaxValue).ToString("0.000")
                        End If

                    End If


                    If pVal.ItemUID = "Mtx7" AndAlso pVal.ColUID = "Tax" AndAlso pVal.BeforeAction = False Then

                        Dim objMatrix7 As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx7").Specific, SAPbouiCOM.Matrix)

                        Dim BaseValue As Double = 0
                        Dim TaxValue As Double = 0

                        Double.TryParse(objMatrix7.Columns.Item("Base").Cells.Item(pVal.Row).Specific.Value.ToString(), BaseValue)
                        Double.TryParse(objMatrix7.Columns.Item("Tax").Cells.Item(pVal.Row).Specific.Value.ToString(), TaxValue)

                        If BaseValue <> 0 OrElse TaxValue <> 0 Then
                            objMatrix7.Columns.Item("NPRC").Cells.Item(pVal.Row).Specific.Value = (BaseValue + TaxValue).ToString("0.000")
                        End If

                    End If


            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    'Public Sub FillBarcodeData(ByVal FormUID As String, ByVal Row As Integer, ByVal Barcode As String)
    '    Try
    '        Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUID)
    '        Dim objMatrix2 As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx2").Specific, SAPbouiCOM.Matrix)
    '        Dim oDBs_Details2 As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C1")

    '        Dim rs As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        Dim sQry As String = "SELECT T0.""BcdCode"", T0.""BcdName"", T0.""ItemCode"", T3.""UomCode"" " &
    '                         "FROM OBCD T0 " &
    '                         "INNER JOIN OUOM T3 ON T3.""UomEntry"" = T0.""UomEntry"" " &
    '                         "WHERE T0.""BcdCode"" = '" & Barcode & "'"
    '        rs.DoQuery(sQry)

    '        If rs.RecordCount > 0 Then
    '            objForm.Freeze(True)


    '            objMatrix2.FlushToDataSource()

    '            Dim targetIndex As Integer = oDBs_Details2.Size - 1

    '            If targetIndex >= 0 AndAlso Not String.IsNullOrEmpty(oDBs_Details2.GetValue("U_ItemCode", targetIndex).Trim()) Then

    '                oDBs_Details2.InsertRecord(oDBs_Details2.Size)
    '                targetIndex = oDBs_Details2.Size - 1
    '            ElseIf targetIndex < 0 Then

    '                oDBs_Details2.InsertRecord(0)
    '                targetIndex = 0
    '            End If


    '            Dim nextLineId As Integer = targetIndex + 1
    '            oDBs_Details2.SetValue("LineId", targetIndex, nextLineId.ToString())
    '            oDBs_Details2.SetValue("VisOrder", targetIndex, targetIndex.ToString())


    '            oDBs_Details2.SetValue("U_BrCode", targetIndex, rs.Fields.Item("BcdCode").Value.ToString())
    '            oDBs_Details2.SetValue("U_ItemCode", targetIndex, rs.Fields.Item("ItemCode").Value.ToString())
    '            oDBs_Details2.SetValue("U_ITNM", targetIndex, rs.Fields.Item("BcdName").Value.ToString())
    '            oDBs_Details2.SetValue("U_UOM", targetIndex, rs.Fields.Item("UomCode").Value.ToString())


    '            objMatrix2.LoadFromDataSource()
    '            Me.SetNewLine2(objForm.UniqueID)

    '            objForm.Freeze(False)
    '        End If

    '    Catch ex As Exception
    '        objForm.Freeze(False)
    '        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '    End Try
    'End Sub
    Public Sub FillBarcodeData(ByVal FormUID As String, ByVal Row As Integer, ByVal Barcode As String, ByVal Itemcode As String)

        Try

            Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUID)

            Dim objMatrix As SAPbouiCOM.Matrix = Nothing
            Dim oDBs_Details As SAPbouiCOM.DBDataSource = Nothing

            Dim oDBs_Head As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")

            Dim priceUpdateType As String = oDBs_Head.GetValue("U_TNXPUTP", 0).Trim()

            Select Case priceUpdateType

                Case "Flyer Price Changes"
                    objMatrix = CType(objForm.Items.Item("Mtx2").Specific, SAPbouiCOM.Matrix)
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C1")
                    ''CalculateGPTotals(objForm)
                    SetMatrixTotals(objForm)
                    Dim rs As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    Dim sQry As String =
                "SELECT T0.""BcdCode"",T0.""BcdName"",T0.""ItemCode"",T3.""UomCode"" " &
                "FROM OBCD T0 " &
                "INNER JOIN OUOM T3 ON T3.""UomEntry"" = T0.""UomEntry"" " &
                "WHERE T0.""BcdCode"" = '" & Barcode & "' and T0.""ItemCode""='" & Itemcode & "'"

                    rs.DoQuery(sQry)

                    If rs.RecordCount > 0 Then

                        objForm.Freeze(True)

                        objMatrix.FlushToDataSource()

                        Dim targetIndex As Integer = Row - 1
                        If targetIndex < 0 Then targetIndex = 0
                        While targetIndex >= oDBs_Details.Size
                            oDBs_Details.InsertRecord(oDBs_Details.Size)
                        End While

                        Dim avgCost As String = FetchBranchAndAvgPriceByBarcode(Barcode)


                        Dim itemGroup As String = ""

                        Dim rsGrp As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        rsGrp.DoQuery(
                    "SELECT T2.""ItmsGrpNam"" " &
                    "FROM OITM T1 " &
                    "LEFT JOIN OITB T2 ON T1.""ItmsGrpCod"" = T2.""ItmsGrpCod"" " &
                    "WHERE T1.""ItemCode"" = '" & Itemcode & "'")

                        If rsGrp.RecordCount > 0 Then
                            itemGroup = rsGrp.Fields.Item("ItmsGrpNam").Value.ToString()
                        End If

                        Dim nextLineId As Integer = targetIndex + 1

                        oDBs_Details.SetValue("LineId", targetIndex, nextLineId.ToString())
                        oDBs_Details.SetValue("VisOrder", targetIndex, targetIndex.ToString())

                        oDBs_Details.SetValue("U_BrCode", targetIndex, rs.Fields.Item("BcdCode").Value.ToString())
                        oDBs_Details.SetValue("U_ItemCode", targetIndex, Itemcode)
                        oDBs_Details.SetValue("U_ITNM", targetIndex, rs.Fields.Item("BcdName").Value.ToString())
                        oDBs_Details.SetValue("U_UOM", targetIndex, rs.Fields.Item("UomCode").Value.ToString())

                        'Additional fields same as CFL
                        oDBs_Details.SetValue("U_ITGP", targetIndex, itemGroup)
                        oDBs_Details.SetValue("U_CPrice", targetIndex, avgCost)

                        'Copy values from matrix if required
                        Try
                            oDBs_Details.SetValue("U_NwPrice", targetIndex,
                        objMatrix.Columns.Item("NwPrice").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_Cost", targetIndex,
                        objMatrix.Columns.Item("Cost").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_LPrice", targetIndex,
                        objMatrix.Columns.Item("LPrice").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_Marg", targetIndex,
                        objMatrix.Columns.Item("Marg").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        objMatrix.LoadFromDataSource()

                        If targetIndex + 1 = objMatrix.VisualRowCount Then
                            Me.SetNewLine2(objForm.UniqueID)

                        End If

                        objForm.Freeze(False)

                    End If

                Case "MOP"
                    objMatrix = CType(objForm.Items.Item("Mtx3").Specific, SAPbouiCOM.Matrix)
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C2")
                    Dim rs As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    Dim sQry As String =
                "SELECT T0.""BcdCode"",T0.""BcdName"",T0.""ItemCode"",T3.""UomCode"" " &
                "FROM OBCD T0 " &
                "INNER JOIN OUOM T3 ON T3.""UomEntry"" = T0.""UomEntry"" " &
                "WHERE T0.""BcdCode"" = '" & Barcode & "' and T0.""ItemCode""='" & Itemcode & "'"

                    rs.DoQuery(sQry)

                    If rs.RecordCount > 0 Then

                        objForm.Freeze(True)

                        objMatrix.FlushToDataSource()

                        Dim targetIndex As Integer = Row - 1
                        If targetIndex < 0 Then targetIndex = 0
                        While targetIndex >= oDBs_Details.Size
                            oDBs_Details.InsertRecord(oDBs_Details.Size)
                        End While

                        Dim avgCost As String = FetchBranchAndAvgPriceByBarcode(Barcode)

                        oDBs_Details.SetValue("LineId", targetIndex, (targetIndex + 1).ToString())

                        oDBs_Details.SetValue("U_BCode", targetIndex, rs.Fields.Item("BcdCode").Value.ToString())
                        oDBs_Details.SetValue("U_ItemCode", targetIndex, Itemcode)
                        oDBs_Details.SetValue("U_IName", targetIndex, rs.Fields.Item("BcdName").Value.ToString())
                        oDBs_Details.SetValue("U_UOM", targetIndex, rs.Fields.Item("UomCode").Value.ToString())
                        oDBs_Details.SetValue("U_CCost", targetIndex, avgCost)

                        Try
                            oDBs_Details.SetValue("U_CSPrice", targetIndex,
                                objMatrix.Columns.Item("CSPrice").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_CMarg", targetIndex,
                                objMatrix.Columns.Item("CMarg").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_NSPrice", targetIndex,
                                objMatrix.Columns.Item("NSPrice").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_NMarg", targetIndex,
                                objMatrix.Columns.Item("NMarg").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_NCost", targetIndex,
                                objMatrix.Columns.Item("NCost").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_CName", targetIndex,
                                objMatrix.Columns.Item("CName").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_IAttach", targetIndex,
                                objMatrix.Columns.Item("IAttach").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_LPhPrice", targetIndex,
                                objMatrix.Columns.Item("LPhPrice").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_LGRPPrice", targetIndex,
                                objMatrix.Columns.Item("LGRPPrice").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        objMatrix.LoadFromDataSource()

                        If targetIndex + 1 = objMatrix.VisualRowCount Then
                            Me.SetNewLine3(objForm.UniqueID)
                        End If

                        objForm.Freeze(False)

                    End If



                Case "Temporary Price Reduction"

                    objMatrix = CType(objForm.Items.Item("Mtx4").Specific, SAPbouiCOM.Matrix)
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C3")

                    Dim rs As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    Dim sQry As String =
    "SELECT T0.""BcdCode"",T0.""BcdName"",T0.""ItemCode"",T3.""UomCode"" " &
    "FROM OBCD T0 " &
    "INNER JOIN OUOM T3 ON T3.""UomEntry"" = T0.""UomEntry"" " &
    "WHERE T0.""BcdCode"" = '" & Barcode & "' AND T0.""ItemCode"" = '" & Itemcode & "'"

                    rs.DoQuery(sQry)

                    If rs.RecordCount > 0 Then

                        objForm.Freeze(True)

                        objMatrix.FlushToDataSource()

                        Dim targetIndex As Integer = Row - 1
                        If targetIndex < 0 Then targetIndex = 0
                        While targetIndex >= oDBs_Details.Size
                            oDBs_Details.InsertRecord(oDBs_Details.Size)
                        End While

                        Dim avgCost As String = FetchBranchAndAvgPriceByBarcode(Barcode)

                        oDBs_Details.SetValue("LineId", targetIndex, (targetIndex + 1).ToString())
                        oDBs_Details.SetValue("VisOrder", targetIndex, targetIndex.ToString())

                        oDBs_Details.SetValue("U_BrCode", targetIndex, rs.Fields.Item("BcdCode").Value.ToString())
                        oDBs_Details.SetValue("U_ItemCode", targetIndex, Itemcode)
                        oDBs_Details.SetValue("U_ItemN", targetIndex, rs.Fields.Item("BcdName").Value.ToString())
                        oDBs_Details.SetValue("U_UOM1", targetIndex, rs.Fields.Item("UomCode").Value.ToString())
                        oDBs_Details.SetValue("U_ICost", targetIndex, avgCost)

                        Try
                            oDBs_Details.SetValue("U_Qty", targetIndex,
                objMatrix.Columns.Item("Qty").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_NSell", targetIndex,
                objMatrix.Columns.Item("NSell").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_SPrice", targetIndex,
                objMatrix.Columns.Item("SPrice").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_DiffSP", targetIndex,
                objMatrix.Columns.Item("DiffSP").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_DSPNS", targetIndex,
                objMatrix.Columns.Item("DSPNS").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_TPRC", targetIndex,
                objMatrix.Columns.Item("TPRC").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_TPRS", targetIndex,
                objMatrix.Columns.Item("TPRS").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_FundA", targetIndex,
                objMatrix.Columns.Item("FundA").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_NGP", targetIndex,
                objMatrix.Columns.Item("NGP").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_SGP", targetIndex,
                objMatrix.Columns.Item("SGP").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_Supp", targetIndex,
                objMatrix.Columns.Item("Supp").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_Remark", targetIndex,
                objMatrix.Columns.Item("Remark").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        objMatrix.LoadFromDataSource()

                        If targetIndex + 1 = objMatrix.VisualRowCount Then
                            Me.SetNewLine4(objForm.UniqueID)
                        End If

                        objForm.Freeze(False)

                    End If


                Case "Every day Special Price"
                    objMatrix = CType(objForm.Items.Item("Mtx6").Specific, SAPbouiCOM.Matrix)
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C5")
                    SetMatrixTotals(objForm)
                    Dim rs As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    Dim sQry As String =
                "SELECT T0.""BcdCode"",T0.""BcdName"",T0.""ItemCode"",T3.""UomCode"" " &
                "FROM OBCD T0 " &
                "INNER JOIN OUOM T3 ON T3.""UomEntry"" = T0.""UomEntry"" " &
                "WHERE T0.""BcdCode"" = '" & Barcode & "' and T0.""ItemCode""='" & Itemcode & "'"

                    rs.DoQuery(sQry)

                    If rs.RecordCount > 0 Then

                        objForm.Freeze(True)

                        objMatrix.FlushToDataSource()

                        Dim targetIndex As Integer = Row - 1
                        If targetIndex < 0 Then targetIndex = 0
                        While targetIndex >= oDBs_Details.Size
                            oDBs_Details.InsertRecord(oDBs_Details.Size)
                        End While

                        Dim avgCost As String = FetchBranchAndAvgPriceByBarcode(Barcode)

                        Dim nextLineId As Integer = targetIndex + 1

                        oDBs_Details.SetValue("LineID", targetIndex, nextLineId.ToString())
                        oDBs_Details.SetValue("VisOrder", targetIndex, targetIndex.ToString())

                        oDBs_Details.SetValue("U_BCode", targetIndex, rs.Fields.Item("BcdCode").Value.ToString())
                        oDBs_Details.SetValue("U_ICode", targetIndex, Itemcode)
                        oDBs_Details.SetValue("U_IName", targetIndex, rs.Fields.Item("BcdName").Value.ToString())
                        oDBs_Details.SetValue("U_Unit", targetIndex, rs.Fields.Item("UomCode").Value.ToString())

                        oDBs_Details.SetValue("U_CCost", targetIndex, avgCost)

                        Try
                            oDBs_Details.SetValue("U_NSlng", targetIndex,
                        objMatrix.Columns.Item("Col_8").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_CMarg", targetIndex,
                        objMatrix.Columns.Item("Col_6").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_CSlng", targetIndex,
                        objMatrix.Columns.Item("Col_5").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_NMrgn", targetIndex,
                        objMatrix.Columns.Item("Col_9").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        Try
                            oDBs_Details.SetValue("U_NCost", targetIndex,
                        objMatrix.Columns.Item("Col_7").Cells.Item(Row).Specific.Value)
                        Catch
                        End Try

                        objMatrix.LoadFromDataSource()

                        If targetIndex + 1 = objMatrix.VisualRowCount Then
                            Me.SetNewLine6(objForm.UniqueID)
                        End If

                        objForm.Freeze(False)

                    End If

                Case "Bulk Cost Update"
                    objMatrix = CType(objForm.Items.Item("Mtx7").Specific, SAPbouiCOM.Matrix)
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C6")
                    SetMatrixTotals(objForm)
                    Dim rs As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    Dim sQry As String =
                "SELECT T0.""BcdCode"",T0.""BcdName"",T0.""ItemCode"",T3.""UomCode"" " &
                "FROM OBCD T0 " &
                "INNER JOIN OUOM T3 ON T3.""UomEntry"" = T0.""UomEntry"" " &
                "WHERE T0.""BcdCode"" = '" & Barcode & "' and T0.""ItemCode""='" & Itemcode & "'"

                    rs.DoQuery(sQry)

                    If rs.RecordCount > 0 Then

                        objForm.Freeze(True)

                        objMatrix.FlushToDataSource()

                        Dim targetIndex As Integer = Row - 1
                        If targetIndex < 0 Then targetIndex = 0
                        While targetIndex >= oDBs_Details.Size
                            oDBs_Details.InsertRecord(oDBs_Details.Size)
                        End While

                        Dim avgCost As String = FetchBranchAndAvgPriceByBarcode(Barcode)

                        Dim nextLineId As Integer = targetIndex + 1

                        oDBs_Details.SetValue("LineID", targetIndex, nextLineId.ToString())
                        oDBs_Details.SetValue("VisOrder", targetIndex, targetIndex.ToString())

                        oDBs_Details.SetValue("U_BCode", targetIndex, rs.Fields.Item("BcdCode").Value.ToString())
                        oDBs_Details.SetValue("U_ITCD", targetIndex, Itemcode)
                        oDBs_Details.SetValue("U_ITNM", targetIndex, rs.Fields.Item("BcdName").Value.ToString())
                        oDBs_Details.SetValue("U_UOM", targetIndex, rs.Fields.Item("UomCode").Value.ToString())

                        objMatrix.LoadFromDataSource()

                        If targetIndex + 1 = objMatrix.VisualRowCount Then
                            Me.SetNewLine7(objForm.UniqueID)
                        End If

                        objForm.Freeze(False)

                    End If

                Case "Bulk Price Update"
                    objMatrix = CType(objForm.Items.Item("Mtx9").Specific, SAPbouiCOM.Matrix)
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C8")
                    SetMatrixTotals(objForm)
                    Dim rs As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    Dim sQry As String =
                "SELECT T0.""BcdCode"",T0.""BcdName"",T0.""ItemCode"",T3.""UomCode"" " &
                "FROM OBCD T0 " &
                "INNER JOIN OUOM T3 ON T3.""UomEntry"" = T0.""UomEntry"" " &
                "WHERE T0.""BcdCode"" = '" & Barcode & "' and T0.""ItemCode""='" & Itemcode & "'"

                    rs.DoQuery(sQry)

                    If rs.RecordCount > 0 Then

                        objForm.Freeze(True)

                        objMatrix.FlushToDataSource()

                        Dim targetIndex As Integer = Row - 1
                        If targetIndex < 0 Then targetIndex = 0
                        While targetIndex >= oDBs_Details.Size
                            oDBs_Details.InsertRecord(oDBs_Details.Size)
                        End While

                        Dim avgCost As String = FetchBranchAndAvgPriceByBarcode(Barcode)

                        Dim nextLineId As Integer = targetIndex + 1

                        oDBs_Details.SetValue("LineID", targetIndex, nextLineId.ToString())
                        oDBs_Details.SetValue("VisOrder", targetIndex, targetIndex.ToString())

                        oDBs_Details.SetValue("U_BCode", targetIndex, rs.Fields.Item("BcdCode").Value.ToString())
                        oDBs_Details.SetValue("U_ITCD", targetIndex, Itemcode)
                        oDBs_Details.SetValue("U_BName", targetIndex, rs.Fields.Item("BcdName").Value.ToString())
                        oDBs_Details.SetValue("U_UOM", targetIndex, rs.Fields.Item("UomCode").Value.ToString())

                        objMatrix.LoadFromDataSource()

                        If targetIndex + 1 = objMatrix.VisualRowCount Then
                            Me.SetNewLine9(objForm.UniqueID)
                        End If

                        objForm.Freeze(False)

                    End If

            End Select




        Catch ex As Exception

            Try
                objForm.Freeze(False)
            Catch
            End Try

            objMain.objApplication.StatusBar.SetText(ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
    Public Sub SendEmail(ByVal toAddress As String, ByVal subject As String, ByVal body As String, ByVal FormUID As String, ByVal DocEntry As String, ByVal DocType As String)

        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12

        Using client As New SmtpClient()
            Using message As New MailMessage()

                message.To.Add(toAddress)
                message.IsBodyHtml = True
                message.Subject = subject
                message.Body = body

                client.EnableSsl = True

                ServicePointManager.ServerCertificateValidationCallback =
                    Function(s As Object, certificate As X509Certificate, chain As X509Chain, sslPolicyErrors As SslPolicyErrors) True


                Dim QUE As String = "Select T0.""U_SMTPsevr"", T0.""U_SMPTPort"", T0.""U_EmailFro"", T0.""U_EmailPwd"" from ""@SBO_SETCONFIG"" T0"

                Dim oRsQUE As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsQUE.DoQuery(QUE)



                Dim userName As String = oRsQUE.Fields.Item("U_EmailFro").Value.ToString().Trim()
                Dim password As String = oRsQUE.Fields.Item("U_EmailPwd").Value.ToString().Trim()
                Dim domain As String = oRsQUE.Fields.Item("U_SMTPsevr").Value.ToString().Trim()

                client.Host = oRsQUE.Fields.Item("U_SMTPsevr").Value.ToString().Trim()
                client.Port = oRsQUE.Fields.Item("U_SMPTPort").Value.ToString().Trim()

                'Dim userName As String = ConfigurationManager.AppSettings("SMTPUSER")
                'Dim password As String = ConfigurationManager.AppSettings("SMTPPWD")
                'Dim domain As String = ConfigurationManager.AppSettings("SMTPSERVER")

                'client.Host = ConfigurationManager.AppSettings("SMTPSERVER")
                'client.Port = Convert.ToInt32(ConfigurationManager.AppSettings("SMTPPORT"))

                Dim Attachment As String = Nothing
                Dim Signature As String = Nothing
                Dim Base64 As String = Nothing
                Dim AttachmentNew As Attachment
                Dim AtttM As String = objMain.objUtilities.AttachmnetSelection(DocType, "Email")
                If AtttM = 1 Then

                    ''Me.GetDocument(DocEntry)
                    Dim AttPath As String = objMain.objUtilities.GetAttachmnetPath()
                    Dim CrystalID As String = objMain.objUtilities.GetCrystalID(DocType, "Email")
                    Signature = objMain.objUtilities.GetSignaturePCR(DocType, "Email")
                    objMain.objUtilities.GetManualAttachment(DocEntry, CrystalID, DocType)
                    Dim Attachment1 As String = AttPath & DocType & "_" & DocEntry & "_" & System.DateTime.Now.ToString("yyyyMMdd") & ".pdf"
                    ''Dim Attachment1 As String = AttPath
                    message.Attachments.Add(New System.Net.Mail.Attachment(Attachment1))
                    ''Attachment = "Please find the attached document. "

                    ''Base64 = objMain.ConvertFileToBase64(AttPath + Signature)
                    'AttachmentNew = New Attachment(AttPath + Signature)
                    'AttachmentNew.ContentId = "NewImage"
                Else
                    Try
                        If objMain.objUtilities.CheckAttachment_EMAIL(DocType, "Email") = True Then
                            Dim temp As String = objMain.objUtilities.getAttachment(DocType, DocEntry)
                            message.Attachments.Add(New System.Net.Mail.Attachment(temp))
                            ''Attachment = "Please find the attached document. "
                        End If
                    Catch ex As Exception
                        objMain.objUtilities.Write_Log(ex.Message)
                    End Try
                End If

                '---------------- HTML EMAIL BODY ----------------

                Dim htmlBody As String =
            "<table>" &
            "<tr><td style='font-size:14px;font-family:Times New Roman;'><b>" & body & "</b></td></tr>" &
            "<tr><td style='font-size:14px;font-family:Times New Roman;'>" & Attachment & "</td></tr>" &
            "</table>" &
            "<br/><br/>" &
"<table>" &
"<tr><td style='font-size:14px;font-family:Times New Roman;'>" & Signature & "</td></tr>" &
"</table>" &
"<br/>" &
            "<table>" &
            "<tr><td style='font-size:14px;font-family:Times New Roman;'>This is system auto generated mail. Please do not reply.</td></tr>" &
            "</table>"

                Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(htmlBody, Nothing, "text/html")

                If AttachmentNew IsNot Nothing Then
                    Dim LR As New LinkedResource(AttachmentNew.ContentStream)
                    LR.ContentId = "NewImage"
                    htmlView.LinkedResources.Add(LR)
                End If

                message.AlternateViews.Add(htmlView)

                '--------------------------------------------------



                If userName IsNot Nothing AndAlso password IsNot Nothing AndAlso domain IsNot Nothing Then

                    client.Credentials = New NetworkCredential(userName, password)

                    Try
                        client.Send(message)
                        MessageBox.Show("Email Sent Successfully to : " & toAddress)
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try

                End If

            End Using
        End Using

    End Sub

    Sub FilterAP(ByVal FormUID As String, ByVal CFL_Id As String, ByVal ItemCode As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim oBINF As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item(CFL_Id)
            Dim emptyConds As New SAPbouiCOM.Conditions

            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
            oDBs_Details4 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C3")



            Dim BINCODE As String = "Select T0.""DocEntry"" as ""DocEntry"" FROM ORPC T0 inner Join RPC1 T1 On T0.""DocEntry""=T1.""DocEntry"" Where T1.""ItemCode"" ='" + ItemCode + "' and T0.""DocStatus""='O' and  T0.""DocEntry"" not in (SELECT TO_INTEGER(NULLIF(""U_CRNT"", '')) FROM ""@TNX_PCRT_C3"" Where ""U_CRNT"" <> '')"
            oBINF.DoQuery(BINCODE)


            For i As Integer = 1 To oBINF.RecordCount
                If i > 1 Then oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                oCon = oCons.Add()
                oCon.Alias = "DocEntry"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = oBINF.Fields.Item("DocEntry").Value
                oBINF.MoveNext()
            Next
            If oBINF.RecordCount = 0 Then
                oCon = oCons.Add()
                oCon.Alias = "DocEntry"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = "-1"
            End If
            oCFL.SetConditions(oCons)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub FilterAP2(ByVal FormUID As String, ByVal CFL_Id As String, ByVal ItemCode As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim oBINF As SAPbobsCOM.Recordset = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item(CFL_Id)
            Dim emptyConds As New SAPbouiCOM.Conditions

            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C2")



            Dim BINCODE As String = "Select T0.""DocEntry"" as ""DocEntry"" FROM ORPC T0 inner Join RPC1 T1 On T0.""DocEntry""=T1.""DocEntry"" Where T1.""ItemCode"" ='" + ItemCode + "' and T0.""DocStatus""='O' and  T0.""DocEntry"" not in (SELECT TO_INTEGER(NULLIF(""U_CRNT"", '')) FROM ""@TNX_PCRT_C1"" Where ""U_CRNT"" <> '')"
            oBINF.DoQuery(BINCODE)


            For i As Integer = 1 To oBINF.RecordCount
                If i > 1 Then oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                oCon = oCons.Add()
                oCon.Alias = "DocEntry"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = oBINF.Fields.Item("DocEntry").Value
                oBINF.MoveNext()
            Next
            If oBINF.RecordCount = 0 Then
                oCon = oCons.Add()
                oCon.Alias = "DocEntry"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = "-1"
            End If
            oCFL.SetConditions(oCons)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Private Sub SetPaneBasedOnSelection(ByVal objForm As SAPbouiCOM.Form)

        Dim selected As String = objForm.Items.Item("YourComboID").Specific.Value

        Select Case selected
            Case "Cost Fluctuation"
                objForm.PaneLevel = 1
                objForm.EnableMenu("5890", False)

            Case "Flyer Price Changes"
                objForm.PaneLevel = 2
                objForm.EnableMenu("5890", False)

            Case "MOP"
                objForm.PaneLevel = 3
                objForm.Settings.MatrixUID = "Mtx3"
                objForm.EnableMenu("5890", True)

            Case "Temporary Price Reduction"
                objForm.PaneLevel = 4
                objForm.EnableMenu("5890", False)

            Case "Attachment Tab"
                objForm.PaneLevel = 5
                objForm.EnableMenu("5890", False)

            Case Else
                objForm.PaneLevel = 1
        End Select

    End Sub
    Private Sub CalculateMargin(ByVal row As Integer)

        Dim newSellPrice As Double = 0
        Dim lastGRPOPrice As Double = 0
        Dim newMargin As Double = 0

        newSellPrice = CDbl(objMatrix3.Columns.Item("NSPrice").Cells.Item(row).Specific.Value)
        lastGRPOPrice = CDbl(objMatrix3.Columns.Item("LGPrice").Cells.Item(row).Specific.Value)

        If newSellPrice <= 0 Then Exit Sub

        newMargin = ((newSellPrice - lastGRPOPrice) / newSellPrice) * 100

        Dim oDBS As SAPbouiCOM.DBDataSource =
        objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C2")

        oDBS.SetValue("U_NMarg", row - 1, Math.Round(newMargin, 2).ToString)

        objMatrix3.Columns.Item("NMarg").Cells.Item(row).Specific.Value =
        Math.Round(newMargin, 2).ToString

    End Sub

    Private Sub CalculateTotalGP(ByVal FormUID As String)

        objForm = objMain.objApplication.Forms.Item(FormUID)

        Dim objMatrix As SAPbouiCOM.Matrix =
        CType(objForm.Items.Item("Mtx2").Specific, SAPbouiCOM.Matrix)

        Dim TotalCost As Double = 0
        Dim TotalNormalSelling As Double = 0
        Dim TotalFlyerSelling As Double = 0

        For i As Integer = 1 To objMatrix.VisualRowCount

            TotalCost += Val(CType(
            objMatrix.Columns.Item("Cost").Cells.Item(i).Specific,
            SAPbouiCOM.EditText).Value)

            TotalNormalSelling += Val(CType(
            objMatrix.Columns.Item("Col_0").Cells.Item(i).Specific,
            SAPbouiCOM.EditText).Value)

            TotalFlyerSelling += Val(CType(
            objMatrix.Columns.Item("Col_1").Cells.Item(i).Specific,
            SAPbouiCOM.EditText).Value)

        Next

        Dim NormalGP As Double = 0
        Dim FlyerGP As Double = 0

        If TotalNormalSelling <> 0 Then
            NormalGP = ((TotalNormalSelling - TotalCost) / TotalNormalSelling) * 100
        End If

        If TotalFlyerSelling <> 0 Then
            FlyerGP = ((TotalFlyerSelling - TotalCost) / TotalFlyerSelling) * 100
        End If

        CType(objForm.Items.Item("Item_31").Specific,
          SAPbouiCOM.EditText).Value = Math.Round(NormalGP, 3).ToString()

        CType(objForm.Items.Item("Item_32").Specific,
          SAPbouiCOM.EditText).Value = Math.Round(FlyerGP, 3).ToString()

    End Sub

    'Approval

    Sub ApprovalTrigger(ByVal DocType As String, ByVal AppID As String, ByVal Table As String, ByVal AppIDField As String, ByVal AppStatField As String, ByVal Document As String, ByVal PriceChangeRequestType As String)
        Dim ff As Integer
        Dim draftquery As String
        Dim approverid As String
        Dim docdate As String = objForm.Items.Item("TNXDDT").Specific.Value

        ff = objMain.objApplication.MessageBox("Do You want Submit the Application?", 1, "Ok", "Cancel")
        If ff = 1 Then
            Dim result As Integer
            result = objMain.objApplication.MessageBox("You can not change the Application after you Sending for Approval,continue", 1, "Yes", "No")
            If result = 1 Then
                Dim rsapp As SAPbobsCOM.Recordset
                Dim posrs As SAPbobsCOM.Recordset
                rsapp = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim str As String = "SELECT T0.""Code"" As ""TemplateID"",T0.""U_PCR"" As ""Price Change Request Report"",T1.""U_Name"" As ""Originator"",S1.""Code"" As ""Stage"",S2.""U_UKey"" As ""Authorizers"" " &
                " ,S1.""U_NAP"" As ""NoOfAuth"",S1.""U_NRJ"" As ""NoOfRej"",T0.""CreateDate"",(Case When T3.""LineId""='1' Then 'S' Else 'O' End) As ""Status"" FROM ""@SBO_APPHDR"" T0 " &
                " INNER JOIN ""@SBO_APPREQ"" T1 ON T1.""Code""=T0.""Code"" " &
                " INNER JOIN ""@SBO_APPDOC"" T2 ON T2.""Code"" = T0.""Code"" " &
                " INNER JOIN ""@SBO_APPAUT"" T3 ON T3.""Code"" = T0.""Code"" " &
                " INNER JOIN ""@SBO_AST"" S1 ON T3.""U_M3_1""=S1.""Code"" " &
                " INNER JOIN ""@SBO_AST_C0"" S2 ON S1.""Code""=S2.""Code"" " &
                " WHERE ""U_Active"" = 'Y' AND T1.""U_Name"" = '" & objMain.objCompany.UserName & "' AND T2.""U_" & DocType & """  = 'Y' AND T0.""U_PCR"" = '" & PriceChangeRequestType & "' AND IFNULL(S2.""U_UKey"",'')<>''  order by T3.""LineId""  ;"
                rsapp.DoQuery(str)

                If rsapp.RecordCount > 0 Then
                    posrs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    draftquery = " SELECT * FROM ""@SBO_OWDD"" WHERE ""U_statusN"" IN ('S','A') AND ""U_Isdraft"" IN ('Y','N') AND ""U_docnum"" = '" & AppID & "';"
                    'draftquery = " SELECT * FROM ""@SBO_DD"" WHERE ""U_statusN"" IN ('S','A') AND ""U_Isdraft"" IN ('Y','N') AND ""U_AppID"" = '" & AppID & "';"

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

    'Application Id
    Public Function AutoDocentryNumber()

        'Dim oRsDocNum As SAPbobsCOM.Recordset = objForm.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objForm = objMain.objApplication.Forms.Item(objForm.UniqueID)
        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
        Dim oRsDocNum As SAPbobsCOM.Recordset

        oRsDocNum = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim Query1 As String = "SELECT ifnull(Max(""DocNum""),0)+1 AS ""DocNum"" FROM ""@TNX_PCRT"" "

        oRsDocNum.DoQuery(Query1)

        oDBs_Head.SetValue("DocNum", 0, oRsDocNum.Fields.Item(0).Value)

        Dim rsAppId As SAPbobsCOM.Recordset

        rsAppId = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim str As String = "SELECT 'PCR' || LPAD('',6 - LENGTH(LTRIM(RTRIM(REPLACE(IFNULL(MAX(""U_AppID""), '0'), 'PCR', '') + 1))),'0') || LTRIM(RTRIM(REPLACE(IFNULL(MAX(""U_AppID""), '0'), 'PCR', '') + 1)) FROM ""@TNX_PCRT"";"

        rsAppId.DoQuery(str)
        Dim appp As String = rsAppId.Fields.Item(0).Value.ToString().Trim()

        oDBs_Head.SetValue("U_AppID", oDBs_Head.Offset, appp)


    End Function

    Public Sub CheckDocNumAndGenerateAutoID()

        Try
            Dim docnumcc As String = Me.objForm.Items.Item("AppID").Specific.Value

            Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            '--- Check if DocNum already exists
            Dim checkQry As String =
            "SELECT COUNT(*) FROM ""@TNX_PCRT"" WHERE ""U_AppID"" = '" & docnumcc & "'"

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

                Dim Str As String =
            "SELECT 'PCR' || LPAD('',6 - LENGTH(LTRIM(RTRIM(REPLACE(IFNULL(MAX(""U_AppID""), '0'), 'PCR', '') + 1))),'0') || LTRIM(RTRIM(REPLACE(IFNULL(MAX(""U_AppID""), '0'), 'PCR', '') + 1)) FROM ""@TNX_PCRT"";"


                rs.DoQuery(Str)

                oDBs_Head.SetValue("U_AppID", 0, rs.Fields.Item(0).Value)
                Dim newReqNo As String = rs.Fields.Item(0).Value.ToString()
                objMain.objApplication.StatusBar.SetText("Duplicate detected. New Price Change Request Number generated: " & newReqNo, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "Problem while processing Auto ID : " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub


    Sub StartBackgroundForPriceUpdateService(Optional ByVal BackGroundWrkr As Boolean = True)
        Try
            If BackGroundWrkr = True Then
                If BackgroundWorker_PriceUpdateService.IsBusy <> True Then
                    BackgroundWorker_PriceUpdateService.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Calling Unselecting worker Error : " & ex.Message)
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
        Dim Pathxl As String = ""

        Try
            If e IsNot Nothing AndAlso e.Argument IsNot Nothing Then
                Pathxl = e.Argument.ToString().Trim()
            Else
                Try
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
                    Pathxl = oDBs_Head.GetValue("U_EXLNK", 0).ToString().Trim()
                Catch
                    Pathxl = ""
                End Try
            End If
        Catch ex As Exception
            Pathxl = ""
        End Try

        If Pathxl <> "" Then
            Try
                Me.Read13(Pathxl)
            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText("Background loader error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        Else
            objMain.objApplication.StatusBar.SetText("Please select a file path in EXLNK before loading.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If
    End Sub

    Private Sub BackgroundWorker_Loading_DoWork2(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_PriceUpdateService.DoWork
        Try
            Me.UpdatePriceNew_ServiceLayer_Itemwise()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Price Update Service Error " & ex.Message)
        End Try
    End Sub

    Private Sub Read13(ByVal PackingList As String)
        Dim barcode, Pricelist, nPrice, StartDate As String
        Dim oRecordset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oRS As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim dtNew As New DataTable()

        ' === Create table if not exists ===
        Dim createTableSQL As String = "CREATE COLUMN TABLE ""TNX_PRICEUPDATE""(""BARCODE"" NVARCHAR(100),
                                       ""NEWSELLINGPRICE"" DECIMAL(11,5) CS_FIXED, ""CREATEDBY"" NVARCHAR(100), ""LINENUM"" INTEGER,""UPDT"" NVARCHAR(100)) 
                                        UNLOAD PRIORITY 5 AUTO MERGE;"

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

        If dtNew.Columns.Count = 0 Then
            objMain.objApplication.StatusBar.SetText("Excel sheet contains no columns.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If

        ' Normalize column names for lookup
        Dim colNames As New Dictionary(Of Integer, String)
        For ci As Integer = 0 To dtNew.Columns.Count - 1
            Dim cn As String = dtNew.Columns(ci).ColumnName
            cn = cn.Trim().Trim(ChrW(65279))
            Dim norm As String = cn.ToLower().Replace(" ", "").Replace("_", "")
            colNames(ci) = norm
        Next

        Dim FunctionFindCol As Func(Of String(), Integer) = Function(keys As String()) As Integer
                                                                For Each kvp In colNames
                                                                    For Each k As String In keys
                                                                        If kvp.Value.Contains(k.ToLower().Replace(" ", "").Replace("_", "")) Then
                                                                            Return kvp.Key
                                                                        End If
                                                                    Next
                                                                Next
                                                                Return -1
                                                            End Function

        Dim idxBarcode As Integer = FunctionFindCol(New String() {"barcode", "bcode", "barcod", "barcode"})
        Dim idxPricelist As Integer = FunctionFindCol(New String() {"pricelist", "plist", "pricelistname"})
        Dim idxPrice As Integer = FunctionFindCol(New String() {"newsellingprice", "nsprice", "newprice", "price", "currentsellingprice", "currentselling"})
        Dim idxStartDate As Integer = FunctionFindCol(New String() {"startdate", "updt", "date"})

        If idxBarcode = -1 AndAlso dtNew.Columns.Count > 0 Then
            idxBarcode = 0
        End If

        '' Insert into TNX_PRICEUPDATE row-by-row (avoid UI operations here)
        'csvInserted = 0
        'csvSkipped = 0
        'For i As Integer = 0 To dtNew.Rows.Count - 1
        '    Try
        '        Dim bc As String = ""
        '        If idxBarcode >= 0 AndAlso idxBarcode < dtNew.Columns.Count Then
        '            If Not IsDBNull(dtNew.Rows(i).Item(idxBarcode)) Then bc = dtNew.Rows(i).Item(idxBarcode).ToString().Trim()
        '        End If
        '        If String.IsNullOrWhiteSpace(bc) Then
        '            csvSkipped += 1
        '            Continue For
        '        End If

        '        Dim pl As String = ""
        '        If idxPricelist >= 0 AndAlso idxPricelist < dtNew.Columns.Count Then
        '            If Not IsDBNull(dtNew.Rows(i).Item(idxPricelist)) Then pl = dtNew.Rows(i).Item(idxPricelist).ToString().Trim()
        '        End If

        '        Dim pr As String = ""
        '        If idxPrice >= 0 AndAlso idxPrice < dtNew.Columns.Count Then
        '            If Not IsDBNull(dtNew.Rows(i).Item(idxPrice)) Then pr = dtNew.Rows(i).Item(idxPrice).ToString().Trim()
        '        End If

        '        Dim sd As String = ""
        '        If idxStartDate >= 0 AndAlso idxStartDate < dtNew.Columns.Count Then
        '            If Not IsDBNull(dtNew.Rows(i).Item(idxStartDate)) Then sd = dtNew.Rows(i).Item(idxStartDate).ToString().Trim()
        '        End If

        '        Dim lineNumber As Integer = i + 1

        '        bc = bc.Replace("'", "''")
        '        pl = pl.Replace("'", "''")
        '        sd = sd.Replace("'", "''")

        '        Dim priceSql As String = "NULL"
        '        If Not String.IsNullOrWhiteSpace(pr) Then
        '            Dim prClean As String = pr.Replace(",", ".")
        '            Dim d As Decimal
        '            If Decimal.TryParse(prClean, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, d) Then
        '                priceSql = d.ToString(Globalization.CultureInfo.InvariantCulture)
        '            Else
        '                priceSql = "NULL"
        '            End If
        '        End If

        '        Dim insertSql As String = String.Format("INSERT INTO ""TNX_PRICEUPDATE"" (""BARCODE"",""PRICELIST"",""PRICE"",""CREATEDBY"",""LINENUM"",""UPDT"") VALUES ('{0}','{1}', {2}, '{3}', {4}, '{5}')",
        '                                           bc, pl, priceSql, CreatedBy.Replace("'", "''"), lineNumber, sd)
        '        Try
        '            oRS.DoQuery(insertSql)
        '        Catch ex As Exception
        '            ' fallback: try insert without UPDT if DB table missing that column
        '            Dim insertSqlNoUpdt As String = String.Format("INSERT INTO ""TNX_PRICEUPDATE"" (""BARCODE"",""PRICELIST"",""PRICE"",""CREATEDBY"",""LINENUM"") VALUES ('{0}','{1}', {2}, '{3}', {4})",
        '                                                      bc, pl, priceSql, CreatedBy.Replace("'", "''"), lineNumber)
        '            oRS.DoQuery(insertSqlNoUpdt)
        '        End Try

        '        csvInserted += 1
        '    Catch ex As Exception
        '        csvSkipped += 1
        '        Try
        '            AddToFileERRIN("Insert failed (row " & (i + 1).ToString() & "): " & ex.Message)
        '        Catch
        '        End Try
        '        Continue For
        '    End Try
        'Next
        ' Insert into TNX_PRICEUPDATE row-by-row (WITH DUPLICATE PROTECTION)
        csvInserted = 0
        csvSkipped = 0

        For i As Integer = 0 To dtNew.Rows.Count - 1
            Try
                Dim bc As String = ""
                If idxBarcode >= 0 AndAlso idxBarcode < dtNew.Columns.Count Then
                    If Not IsDBNull(dtNew.Rows(i).Item(idxBarcode)) Then
                        bc = dtNew.Rows(i).Item(idxBarcode).ToString().Trim()
                    End If
                End If

                If String.IsNullOrWhiteSpace(bc) Then
                    csvSkipped += 1
                    Continue For
                End If

                Dim pl As String = ""
                If idxPricelist >= 0 AndAlso idxPricelist < dtNew.Columns.Count Then
                    If Not IsDBNull(dtNew.Rows(i).Item(idxPricelist)) Then
                        pl = dtNew.Rows(i).Item(idxPricelist).ToString().Trim()
                    End If
                End If
                'new selling price
                Dim pr As String = ""
                If idxPrice >= 0 AndAlso idxPrice < dtNew.Columns.Count Then
                    If Not IsDBNull(dtNew.Rows(i).Item(idxPrice)) Then
                        pr = dtNew.Rows(i).Item(idxPrice).ToString().Trim()
                    End If
                End If

                Dim sd As String = ""
                If idxStartDate >= 0 AndAlso idxStartDate < dtNew.Columns.Count Then
                    If Not IsDBNull(dtNew.Rows(i).Item(idxStartDate)) Then
                        sd = dtNew.Rows(i).Item(idxStartDate).ToString().Trim()
                    End If
                End If

                Dim lineNumber As Integer = i + 1

                bc = bc.Replace("'", "''")
                pl = pl.Replace("'", "''")
                sd = sd.Replace("'", "''")

                ' ✅ PRICE SAFE PARSE
                Dim priceSql As String = "NULL"
                If Not String.IsNullOrWhiteSpace(pr) Then
                    Dim prClean As String = pr.Replace(",", ".")
                    Dim d As Decimal
                    If Decimal.TryParse(prClean, Globalization.NumberStyles.Any,
                                Globalization.CultureInfo.InvariantCulture, d) Then
                        priceSql = d.ToString(Globalization.CultureInfo.InvariantCulture)
                    End If
                End If

                ' ✅ ✅ DUPLICATE CHECK (BARCODE + PRICELIST + USER)
                '    Dim chkSql As String =
                '"SELECT COUNT(*) FROM ""TNX_PRICEUPDATE"" " &
                '"WHERE ""BARCODE"" = '" & bc & "' " &
                '"AND ""PRICELIST"" = '" & pl & "' " &
                '"AND ""CREATEDBY"" = '" & CreatedBy.Replace("'", "''") & "'"
                Dim chkSql As String =
    "SELECT COUNT(*) FROM ""TNX_PRICEUPDATE"" " &
    "WHERE ""BARCODE"" = '" & bc & "' " &
    "AND ""CREATEDBY"" = '" & CreatedBy.Replace("'", "''") & "'"


                Dim chkRS As SAPbobsCOM.Recordset =
            objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                chkRS.DoQuery(chkSql)

                If chkRS.Fields.Item(0).Value > 0 Then
                    csvSkipped += 1
                    Continue For
                End If

                ' ✅ INSERT ONLY IF NOT DUPLICATE
                '    Dim insertSql As String =
                '"INSERT INTO ""TNX_PRICEUPDATE"" " &
                '"(""BARCODE"",""PRICELIST"",""PRICE"",""CREATEDBY"",""LINENUM"",""UPDT"") " &
                '"VALUES ('" & bc & "','" & pl & "'," & priceSql & ",'" &
                'CreatedBy.Replace("'", "''") & "'," & lineNumber & ",'" & sd & "')"

                Dim insertSql As String =
    "INSERT INTO ""TNX_PRICEUPDATE"" " &
    "(""BARCODE"",""NEWSELLINGPRICE"",""CREATEDBY"",""LINENUM"",""UPDT"") " &
    "VALUES ('" & bc.Replace("'", "''") & "'," & priceSql & ",'" &
    CreatedBy.Replace("'", "''") & "'," & lineNumber & ",'" & sd & "')"


                Try
                    oRS.DoQuery(insertSql)
                Catch ex As Exception
                    ' Fallback if UPDT column not present

                    '    Dim insertSqlNoUpdt As String =
                    '"INSERT INTO ""TNX_PRICEUPDATE"" " &
                    '"(""BARCODE"",""PRICELIST"",""PRICE"",""CREATEDBY"",""LINENUM"") " &
                    '"VALUES ('" & bc & "','" & pl & "'," & priceSql & ",'" &
                    'CreatedBy.Replace("'", "''") & "'," & lineNumber & ")"

                    Dim insertSqlNoUpdt As String =
    "INSERT INTO ""TNX_PRICEUPDATE"" " &
    "(""BARCODE"",""NEWSELLINGPRICE"",""CREATEDBY"",""LINENUM"") " &
    "VALUES ('" & bc.Replace("'", "''") & "'," & priceSql & ",'" &
    CreatedBy.Replace("'", "''") & "'," & lineNumber & ")"


                    oRS.DoQuery(insertSqlNoUpdt)
                End Try

                csvInserted += 1

            Catch ex As Exception
                csvSkipped += 1
                Try
                    AddToFileERRIN("Insert failed (row " & (i + 1).ToString() & "): " & ex.Message)
                Catch
                End Try
                Continue For
            End Try
        Next

        objMain.objApplication.StatusBar.SetText(String.Format("Upload completed. Inserted: {0}, Skipped: {1}", csvInserted, csvSkipped), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        ' A stored proc to prepare item details (resulting recordset is read into a DataTable for UI)
        Try
            ' Dim query As String = "CALL ""TNX_ITEMDETAILS_NEW""('" & CreatedBy & "');"
            Dim BranchID As String = objForm.Items.Item("TNXBRC").Specific.value
            Dim query As String = "CALL ""TNX_ITEMDETAILS_NEW""('" & CreatedBy & "','" & BranchID & "')"
            oRecordset.DoQuery(query)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error loading data call TNX_ITEMDETAILS_NEW(" & CreatedBy & ") : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End Try

        ' Convert oRecordset -> DataTable for UI population (no SAP UI calls here)
        Try
            Dim dtResult As New DataTable()
            If oRecordset.RecordCount > 0 Then
                ' create columns
                For fIndex As Integer = 0 To oRecordset.Fields.Count - 1
                    Dim colName As String = oRecordset.Fields.Item(fIndex).Name
                    If Not dtResult.Columns.Contains(colName) Then
                        dtResult.Columns.Add(colName, GetType(String))
                    End If
                Next

                oRecordset.MoveFirst()
                While Not oRecordset.EoF
                    Dim dr As DataRow = dtResult.NewRow()
                    For fIndex As Integer = 0 To oRecordset.Fields.Count - 1
                        Try
                            Dim val = oRecordset.Fields.Item(fIndex).Value
                            dr(fIndex) = If(val Is Nothing, String.Empty, val.ToString())
                        Catch
                            dr(fIndex) = String.Empty
                        End Try
                    Next
                    dtResult.Rows.Add(dr)
                    oRecordset.MoveNext()
                End While
            End If

            ' store for UI thread to pick up
            dtMatrix3Result = dtResult
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error converting proc result to table: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

        GC.Collect()
    End Sub

    ' --- RunWorkerCompleted event: runs on UI thread, safe to update SAP UI objects and matrix ---
    Private Sub BackgroundWorker_CSVLoad_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker_CSVLoad.RunWorkerCompleted

        Try

            If dtMatrix3Result Is Nothing OrElse dtMatrix3Result.Rows.Count = 0 Then
                objMain.objApplication.StatusBar.SetText("No data returned to populate matrix.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return
            End If

            ' Find open PCRT form
            Dim targetForm As SAPbouiCOM.Form = Nothing
            For Each f As SAPbouiCOM.Form In objMain.objApplication.Forms
                Try
                    If f.TypeEx = "PCRT" Then
                        targetForm = f
                        Exit For
                    End If
                Catch
                End Try
            Next

            If targetForm Is Nothing Then
                objMain.objApplication.StatusBar.SetText("Price Change Request form is not open.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return
            End If

            objForm = targetForm

            ' 🔹 READ HEADER COMBO
            Dim oDBs_Head As SAPbouiCOM.DBDataSource =
        objForm.DataSources.DBDataSources.Item("@TNX_PCRT")

            Dim priceUpdateType As String =
        oDBs_Head.GetValue("U_TNXPUTP", 0).Trim()

            ' 🔹 TARGET MATRIX HOLDERS
            Dim objMatrix As SAPbouiCOM.Matrix = Nothing
            Dim oDBs_Details As SAPbouiCOM.DBDataSource = Nothing

            ' 🔹 DECIDE TARGET MATRIX
            Select Case priceUpdateType

                Case "MOP"
                    objForm.Items.Item("Tab3").Click()
                    objMatrix = objForm.Items.Item("Mtx3").Specific
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C2")
                Case "Flyer Price Changes"
                    objForm.Items.Item("Tab2").Click()
                    objMatrix = objForm.Items.Item("Mtx2").Specific
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C1")

                Case "Temporary Price Reduction"
                    objForm.Items.Item("Tab4").Click()
                    objMatrix = objForm.Items.Item("Mtx4").Specific
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C3")

                Case "Every day Special Price"
                    objForm.Items.Item("Tab6").Click()
                    objMatrix = objForm.Items.Item("Mtx6").Specific
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C5")

                Case "Bulk Cost Update"
                    objForm.Items.Item("Tab7").Click()
                    objMatrix = objForm.Items.Item("Mtx7").Specific
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C6")
                Case "Cost Fluctuation Stock Transfer"
                    objForm.Items.Item("Tab8").Click()
                    objMatrix = objForm.Items.Item("Mtx8").Specific
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C7")
                Case "Bulk Price Update"
                    objForm.Items.Item("Tab9").Click()
                    objMatrix = objForm.Items.Item("Mtx9").Specific
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C8")

                Case Else
                    objMain.objApplication.StatusBar.SetText("Please select MOP or Temporary Price Reduction in header.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return
            End Select

            ' 🔹 CLEAR CORRECT MATRIX
            oDBs_Details.Clear()
            objMatrix.Clear()

            ' 🔹 POPULATE MATRIX
            For i As Integer = 0 To dtMatrix3Result.Rows.Count - 1

                oDBs_Details.InsertRecord(oDBs_Details.Size)
                oDBs_Details.Offset = oDBs_Details.Size - 1



                Dim GetVal As Func(Of String, String) =
            Function(colName As String) As String
                Try
                    If dtMatrix3Result.Columns.Contains(colName) Then
                        Return dtMatrix3Result.Rows(i)(colName).ToString()
                    End If
                Catch
                End Try
                Return ""
            End Function

                oDBs_Details.SetValue("LineID", oDBs_Details.Offset, (i + 1).ToString())

                ' ===== COMMON FIELDS =====
                If priceUpdateType = "MOP" Then

                    oDBs_Details.SetValue("U_BCode", oDBs_Details.Offset, GetVal("BARCODE"))
                    oDBs_Details.SetValue("U_ItemCode", oDBs_Details.Offset, GetVal("ItemCode"))
                    oDBs_Details.SetValue("U_IName", oDBs_Details.Offset, GetVal("ItemName"))
                    oDBs_Details.SetValue("U_UOM", oDBs_Details.Offset, GetVal("UomName"))
                    oDBs_Details.SetValue("U_CCost", oDBs_Details.Offset, GetVal("COST"))
                    oDBs_Details.SetValue("U_CSPrice", oDBs_Details.Offset, GetVal("RSP"))
                    oDBs_Details.SetValue("U_CMarg", oDBs_Details.Offset, GetVal("margin"))
                    oDBs_Details.SetValue("U_NSPrice", oDBs_Details.Offset, GetVal("NEWSELLINGPRICE"))
                    oDBs_Details.SetValue("U_NMarg", oDBs_Details.Offset, GetVal("NMARG"))
                    oDBs_Details.SetValue("U_NCost", oDBs_Details.Offset, GetVal(""))
                    oDBs_Details.SetValue("U_LPhPrice", oDBs_Details.Offset, GetVal("LPHPRICE"))
                    oDBs_Details.SetValue("U_LGRPPrice", oDBs_Details.Offset, GetVal("LGRPPRICE"))

                ElseIf priceUpdateType = "Temporary Price Reduction" Then

                    oDBs_Details.SetValue("U_BrCode", oDBs_Details.Offset, GetVal("BARCODE"))
                    oDBs_Details.SetValue("U_ItemCode", oDBs_Details.Offset, GetVal("ItemCode"))
                    oDBs_Details.SetValue("U_ItemN", oDBs_Details.Offset, GetVal("ItemName"))
                    oDBs_Details.SetValue("U_Qty", oDBs_Details.Offset, "1")
                    oDBs_Details.SetValue("U_ICost", oDBs_Details.Offset, GetVal("COST"))
                    oDBs_Details.SetValue("U_UOM1", oDBs_Details.Offset, GetVal("UomName"))
                    oDBs_Details.SetValue("U_SPrice", oDBs_Details.Offset, GetVal("NEWSELLINGPRICE"))
                    oDBs_Details.SetValue("U_NSell", oDBs_Details.Offset, GetVal("PRICE"))
                    oDBs_Details.SetValue("U_TPRC", oDBs_Details.Offset, GetVal("NCOST"))
                    oDBs_Details.SetValue("U_TPRS", oDBs_Details.Offset, GetVal("PRICE"))
                    oDBs_Details.SetValue("U_NGP", oDBs_Details.Offset, GetVal("NMARG"))
                    oDBs_Details.SetValue("U_SGP", oDBs_Details.Offset, GetVal("margin"))
                    oDBs_Details.SetValue("U_Remark", oDBs_Details.Offset, GetVal("REMARK"))

                    ' New fields from updated procedure
                    'oDBs_Details.SetValue("U_Sl", oDBs_Details.Offset, GetVal("Sl"))               ' Line number
                    oDBs_Details.SetValue("U_Qty", oDBs_Details.Offset, GetVal("Qty"))             ' Requested quantity
                    oDBs_Details.SetValue("U_DiffSP", oDBs_Details.Offset, GetVal("DIFF_SP_COST"))   ' SP - Cost
                    oDBs_Details.SetValue("U_DSPNS", oDBs_Details.Offset, GetVal("Diff_SP_NS")) ' SP - Normal Selling
                    'oDBs_Details.SetValue("U_FundAlloc", oDBs_Details.Offset, GetVal("FundAllocation")) ' Fund Allocation
                    'oDBs_Details.SetValue("U_Supplier", oDBs_Details.Offset, GetVal("Supplier"))   ' Supplier


                ElseIf priceUpdateType = "Flyer Price Changes" Then
                    oDBs_Details.SetValue("U_BrCode", oDBs_Details.Offset, GetVal("BARCODE"))
                    oDBs_Details.SetValue("U_ItemCode", oDBs_Details.Offset, GetVal("ItemCode"))
                    oDBs_Details.SetValue("U_CPrice", oDBs_Details.Offset, GetVal("RSP"))
                    oDBs_Details.SetValue("U_NwPrice", oDBs_Details.Offset, GetVal("NEWSELLINGPRICE"))
                    oDBs_Details.SetValue("U_Cost", oDBs_Details.Offset, GetVal("COST"))
                    oDBs_Details.SetValue("U_LPrice", oDBs_Details.Offset, GetVal("LPHPRICE"))
                    oDBs_Details.SetValue("U_Marg", oDBs_Details.Offset, GetVal("NEWMARG"))
                    oDBs_Details.SetValue("U_UOM", oDBs_Details.Offset, GetVal("UomName"))

                End If

            Next

            ' 🔹 LOAD INTO MATRIX
            objMatrix.LoadFromDataSource()
            objMatrix.AutoResizeColumns()

            objMain.objApplication.StatusBar.SetText("Matrix filled successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            dtMatrix3Result = Nothing

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error populating matrix: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub


    'Private Sub Read13(ByVal PackingList As String)
    '    Dim barcode, Pricelist, nPrice, StartDate As String
    '    objMatrix3 = objForm.Items.Item("Mtx3").Specific
    '    objMatrix3.Clear()
    '    Dim oRecordset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    Dim oRS As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
    '    Dim dtNew As New DataTable()

    '    ' === Create table if not exists ===
    '    Dim createTableSQL As String = "CREATE COLUMN TABLE ""TNX_PRICEUPDATE""(""BARCODE"" NVARCHAR(100), ""PRICELIST"" NVARCHAR(100), ""PRICE"" DECIMAL(11,5) CS_FIXED, ""CREATEDBY"" NVARCHAR(100), ""LINENUM"" INTEGER,""UPDT"" NVARCHAR(100)) UNLOAD PRIORITY 5 AUTO MERGE;"
    '    Try
    '        oRS.DoQuery(createTableSQL)
    '    Catch ex As Exception
    '        If Not ex.Message.ToLower().Contains("already exists") AndAlso Not ex.Message.ToLower().Contains("cannot use duplicate table name") Then
    '            objMain.objApplication.StatusBar.SetText("Error creating table: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Exit Sub
    '        End If
    '    End Try

    '    Dim CreatedBy As String = objMain.objCompany.UserSignature
    '    objMain.objApplication.StatusBar.SetText("Reading Excel & preparing upload...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

    '    ' === Read Excel ===
    '    Try
    '        Dim connStr As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & PackingList & ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';")
    '        connStr.Open()

    '        Dim dtSheets As DataTable = connStr.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
    '        If dtSheets Is Nothing OrElse dtSheets.Rows.Count = 0 Then
    '            objMain.objApplication.StatusBar.SetText("No sheets found in Excel.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            connStr.Close()
    '            Exit Sub
    '        End If

    '        Dim firstSheetName As String = dtSheets.Rows(0)("TABLE_NAME").ToString()
    '        Dim adap As New OleDbDataAdapter("SELECT * FROM [" & firstSheetName & "]", connStr)
    '        adap.Fill(dtNew)
    '        connStr.Close()
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText("Error reading Excel: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        Exit Sub
    '    End Try

    '    ' === Delete old ===
    '    Try
    '        oRS.DoQuery("DELETE FROM ""TNX_PRICEUPDATE"" WHERE ""CREATEDBY"" = '" & CreatedBy & "'")
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText("Error deleting old records: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        Exit Sub
    '    End Try

    '    If dtNew.Columns.Count = 0 Then
    '        objMain.objApplication.StatusBar.SetText("Excel sheet contains no columns.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        Exit Sub
    '    End If

    '    ' Normalize column names for lookup
    '    Dim colNames As New Dictionary(Of Integer, String)
    '    For ci As Integer = 0 To dtNew.Columns.Count - 1
    '        Dim cn As String = dtNew.Columns(ci).ColumnName
    '        cn = cn.Trim().Trim(ChrW(65279))
    '        Dim norm As String = cn.ToLower().Replace(" ", "").Replace("_", "")
    '        colNames(ci) = norm
    '    Next

    '    ' Diagnostic
    '    Dim detectedList As New StringBuilder()
    '    For Each dc As DataColumn In dtNew.Columns
    '        detectedList.Append(dc.ColumnName & ",")
    '    Next
    '    objMain.objApplication.StatusBar.SetText("Detected Excel columns: " & detectedList.ToString().TrimEnd(","c), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

    '    Dim FunctionFindCol As Func(Of String(), Integer) = Function(keys As String()) As Integer
    '                                                            For Each kvp In colNames
    '                                                                For Each k As String In keys
    '                                                                    If kvp.Value.Contains(k.ToLower().Replace(" ", "").Replace("_", "")) Then
    '                                                                        Return kvp.Key
    '                                                                    End If
    '                                                                Next
    '                                                            Next
    '                                                            Return -1
    '                                                        End Function

    '    Dim idxBarcode As Integer = FunctionFindCol(New String() {"barcode", "bcode", "barcod", "barcode"})
    '    Dim idxPricelist As Integer = FunctionFindCol(New String() {"pricelist", "plist", "pricelistname"})
    '    Dim idxPrice As Integer = FunctionFindCol(New String() {"newsellingprice", "nsprice", "newprice", "price", "currentsellingprice", "currentselling"})
    '    Dim idxCost As Integer = FunctionFindCol(New String() {"currentcost", "ccost", "cost"})
    '    Dim idxStartDate As Integer = FunctionFindCol(New String() {"startdate", "updt", "date"})

    '    If idxBarcode = -1 AndAlso dtNew.Columns.Count > 0 Then
    '        idxBarcode = 0
    '    End If

    '    ' --- Replace the INSERT loop in Read13 with this version (adds runtime check for UPDT column) ---
    '    ' Check whether UPDT column exists in TNX_PRICEUPDATE (HANA SYS view)
    '    Dim hasUpdt As Boolean = False
    '    Try
    '        Dim chkRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        chkRs.DoQuery("SELECT COUNT(*) FROM SYS.TABLE_COLUMNS WHERE TABLE_NAME = 'TNX_PRICEUPDATE' AND COLUMN_NAME = 'UPDT'")
    '        If chkRs.Fields.Item(0).Value IsNot Nothing AndAlso CInt(chkRs.Fields.Item(0).Value) > 0 Then
    '            hasUpdt = True
    '        End If
    '    Catch ex As Exception
    '        ' if check fails, assume column missing (safe fallback)
    '        hasUpdt = False
    '    End Try

    '    ' Insert row-by-row (explicit columns)
    '    Dim inserted As Integer = 0
    '    Dim skipped As New List(Of Integer)
    '    ' --- inside Read13, replace the INSERT loop with this (declare insertSql before Try so Catch can reference it) ---
    '    For i As Integer = 0 To dtNew.Rows.Count - 1
    '        Dim insertSql As String = Nothing
    '        Try
    '            Dim bc As String = ""
    '            If idxBarcode >= 0 AndAlso idxBarcode < dtNew.Columns.Count Then
    '                If Not IsDBNull(dtNew.Rows(i).Item(idxBarcode)) Then bc = dtNew.Rows(i).Item(idxBarcode).ToString().Trim()
    '            End If
    '            If String.IsNullOrWhiteSpace(bc) Then
    '                skipped.Add(i + 1)
    '                Continue For
    '            End If

    '            Dim pl As String = ""
    '            If idxPricelist >= 0 AndAlso idxPricelist < dtNew.Columns.Count Then
    '                If Not IsDBNull(dtNew.Rows(i).Item(idxPricelist)) Then pl = dtNew.Rows(i).Item(idxPricelist).ToString().Trim()
    '            End If

    '            Dim pr As String = ""
    '            If idxPrice >= 0 AndAlso idxPrice < dtNew.Columns.Count Then
    '                If Not IsDBNull(dtNew.Rows(i).Item(idxPrice)) Then pr = dtNew.Rows(i).Item(idxPrice).ToString().Trim()
    '            End If

    '            Dim sd As String = ""
    '            If idxStartDate >= 0 AndAlso idxStartDate < dtNew.Columns.Count Then
    '                If Not IsDBNull(dtNew.Rows(i).Item(idxStartDate)) Then sd = dtNew.Rows(i).Item(idxStartDate).ToString().Trim()
    '            End If

    '            Dim lineNumber As Integer = i + 1

    '            ' Escape single quotes in strings
    '            bc = bc.Replace("'", "''")
    '            pl = pl.Replace("'", "''")
    '            sd = sd.Replace("'", "''")

    '            ' Format PRICE: attempt numeric, otherwise NULL
    '            Dim priceSql As String = "NULL"
    '            If Not String.IsNullOrWhiteSpace(pr) Then
    '                Dim prClean As String = pr.Replace(",", ".")
    '                Dim d As Decimal
    '                If Decimal.TryParse(prClean, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, d) Then
    '                    priceSql = d.ToString(Globalization.CultureInfo.InvariantCulture)
    '                Else
    '                    priceSql = "NULL"
    '                End If
    '            End If

    '            If hasUpdt Then
    '                insertSql = String.Format("INSERT INTO ""TNX_PRICEUPDATE"" (""BARCODE"",""PRICELIST"",""PRICE"",""CREATEDBY"",""LINENUM"",""UPDT"") VALUES ('{0}','{1}', {2}, '{3}', {4}, '{5}')",
    '                                  bc, pl, priceSql, CreatedBy.Replace("'", "''"), lineNumber, sd)
    '            Else
    '                insertSql = String.Format("INSERT INTO ""TNX_PRICEUPDATE"" (""BARCODE"",""PRICELIST"",""PRICE"",""CREATEDBY"",""LINENUM"") VALUES ('{0}','{1}', {2}, '{3}', {4})",
    '                                  bc, pl, priceSql, CreatedBy.Replace("'", "''"), lineNumber)
    '            End If

    '            oRS.DoQuery(insertSql)
    '            inserted += 1
    '        Catch ex As Exception
    '            ' insertSql is now in scope here
    '            Try
    '                AddToFileERRIN("Insert failed (row " & (i + 1).ToString() & "): " & ex.Message & " | SQL: " & If(insertSql IsNot Nothing, insertSql, "n/a"))
    '            Catch
    '            End Try
    '            skipped.Add(i + 1)
    '            ' continue with next row
    '        End Try
    '    Next
    '    '' Insert row-by-row (explicit columns) to avoid huge UNION SQL which HANA/ODBC may reject
    '    'Dim inserted As Integer = 0
    '    'Dim skipped As New List(Of Integer)
    '    'For i As Integer = 0 To dtNew.Rows.Count - 1
    '    '    Try
    '    '        Dim bc As String = ""
    '    '        If idxBarcode >= 0 AndAlso idxBarcode < dtNew.Columns.Count Then
    '    '            If Not IsDBNull(dtNew.Rows(i).Item(idxBarcode)) Then bc = dtNew.Rows(i).Item(idxBarcode).ToString().Trim()
    '    '        End If
    '    '        If String.IsNullOrWhiteSpace(bc) Then
    '    '            skipped.Add(i + 1)
    '    '            Continue For
    '    '        End If

    '    '        Dim pl As String = ""
    '    '        If idxPricelist >= 0 AndAlso idxPricelist < dtNew.Columns.Count Then
    '    '            If Not IsDBNull(dtNew.Rows(i).Item(idxPricelist)) Then pl = dtNew.Rows(i).Item(idxPricelist).ToString().Trim()
    '    '        End If

    '    '        Dim pr As String = ""
    '    '        If idxPrice >= 0 AndAlso idxPrice < dtNew.Columns.Count Then
    '    '            If Not IsDBNull(dtNew.Rows(i).Item(idxPrice)) Then pr = dtNew.Rows(i).Item(idxPrice).ToString().Trim()
    '    '        End If

    '    '        Dim sd As String = ""
    '    '        If idxStartDate >= 0 AndAlso idxStartDate < dtNew.Columns.Count Then
    '    '            If Not IsDBNull(dtNew.Rows(i).Item(idxStartDate)) Then sd = dtNew.Rows(i).Item(idxStartDate).ToString().Trim()
    '    '        End If

    '    '        Dim lineNumber As Integer = i + 1

    '    '        ' Escape single quotes in strings
    '    '        bc = bc.Replace("'", "''")
    '    '        pl = pl.Replace("'", "''")
    '    '        sd = sd.Replace("'", "''")

    '    '        ' Format PRICE: attempt numeric, otherwise NULL
    '    '        Dim priceSql As String = "NULL"
    '    '        If Not String.IsNullOrWhiteSpace(pr) Then
    '    '            Dim prClean As String = pr.Replace(",", ".")
    '    '            Dim d As Decimal
    '    '            If Decimal.TryParse(prClean, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, d) Then
    '    '                priceSql = d.ToString(Globalization.CultureInfo.InvariantCulture)
    '    '            Else
    '    '                ' if not parseable, skip price and set NULL
    '    '                priceSql = "NULL"
    '    '            End If
    '    '        End If

    '    '        Dim insertSql As String = String.Format("INSERT INTO ""TNX_PRICEUPDATE"" (""BARCODE"",""PRICELIST"",""PRICE"",""CREATEDBY"",""LINENUM"",""UPDT"") VALUES ('{0}','{1}', {2}, '{3}', {4}, '{5}');", bc, pl, priceSql, CreatedBy.Replace("'", "''"), lineNumber, sd)
    '    '        oRS.DoQuery(insertSql)
    '    '        inserted += 1
    '    '    Catch ex As Exception
    '    '        skipped.Add(i + 1)
    '    '        ' continue with next row
    '    '    End Try
    '    'Next

    '    objMain.objApplication.StatusBar.SetText(String.Format("Upload completed. Inserted: {0}, Skipped: {1}", inserted, skipped.Count), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

    '    ' Call stored proc to prepare item details
    '    Try
    '        Dim query As String = "CALL ""TNX_ITEMDETAILS_NEW""('" & CreatedBy & "');"
    '        oRecordset.DoQuery(query)
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText("Error loading data call TNX_ITEMDETAILS_NEW(" & CreatedBy & ") : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        Exit Sub
    '    End Try

    '    ' Fill matrix from proc output (existing mapping)
    '    Try
    '        ' Replace the "Fill matrix from proc output" inner loop with this corrected mapping.
    '        ' Uses the DB datasource field unique IDs defined in SetNewLine3 and the procedure column names.

    '        If oRecordset.RecordCount > 0 Then
    '            objMatrix3 = objForm.Items.Item("Mtx3").Specific
    '            ' Clear existing datasource and use InsertRecord + Offset to safely populate
    '            Try
    '                oDBs_Details3.Clear()
    '            Catch
    '            End Try

    '            Dim lineNum As Integer = 1
    '            While Not oRecordset.EoF
    '                ' ensure a new datasource record and correct offset
    '                oDBs_Details3.InsertRecord(oDBs_Details3.Size)
    '                oDBs_Details3.Offset = oDBs_Details3.Size - 1

    '                ' helper to read proc fields safely
    '                Dim GetF As Func(Of String, String) = Function(fname As String) As String
    '                                                          Try
    '                                                              Dim f = oRecordset.Fields.Item(fname)
    '                                                              If f Is Nothing OrElse f.Value Is Nothing Then Return String.Empty
    '                                                              Return f.Value.ToString()
    '                                                          Catch
    '                                                              Return String.Empty
    '                                                          End Try
    '                                                      End Function

    '                ' Map stored-proc columns to datasource unique IDs (as declared in SetNewLine3)
    '                oDBs_Details3.SetValue("LineID", oDBs_Details3.Offset, lineNum.ToString())

    '                Dim itemCodeVal As String = If(GetF("ItemCode") <> String.Empty, GetF("ItemCode"), GetF("BCode"))
    '                oDBs_Details3.SetValue("U_BCode", oDBs_Details3.Offset, itemCodeVal)          ' ItemCode / Barcode key
    '                oDBs_Details3.SetValue("U_IName", oDBs_Details3.Offset, GetF("ItemName"))     ' Item name
    '                oDBs_Details3.SetValue("U_UOM", oDBs_Details3.Offset, GetF("UomEntry"))       ' UoM entry (numeric stored as string)
    '                oDBs_Details3.SetValue("U_CCost", oDBs_Details3.Offset, GetF("COST"))         ' Current cost from proc
    '                oDBs_Details3.SetValue("U_CSPrice", oDBs_Details3.Offset, GetF("RSP"))        ' Current selling price (RSP)
    '                oDBs_Details3.SetValue("U_CMarg", oDBs_Details3.Offset, GetF("margin"))       ' Computed margin
    '                oDBs_Details3.SetValue("U_NSPrice", oDBs_Details3.Offset, GetF("PRICE"))      ' Price (from Excel / TNX_PRICEUPDATE)
    '                ' New margin / cost may not be returned by proc — keep empty or set same margin if preferred
    '                oDBs_Details3.SetValue("U_NMarg", oDBs_Details3.Offset, String.Empty)
    '                oDBs_Details3.SetValue("U_NCost", oDBs_Details3.Offset, String.Empty)

    '                ' UoM name, conversion rate and pricelist display fields
    '                oDBs_Details3.SetValue("U_CName", oDBs_Details3.Offset, GetF("UomName"))       ' uom code/name
    '                oDBs_Details3.SetValue("U_IAttach", oDBs_Details3.Offset, GetF("PriceListName")) ' pricelist name / attachment column

    '                ' Optional: preserve additional columns returned by proc if present
    '                If GetF("REMARK") <> String.Empty Then
    '                    Try
    '                        oDBs_Details3.SetValue("U_Remarks", oDBs_Details3.Offset, GetF("REMARK"))
    '                    Catch
    '                    End Try
    '                End If

    '                ' Push values to matrix row
    '                Try
    '                    objMatrix3.LoadFromDataSource()
    '                    objMatrix3.SetLineData(lineNum)
    '                Catch
    '                    ' fallback if UI update fails for current row
    '                End Try

    '                oRecordset.MoveNext()
    '                lineNum += 1
    '            End While

    '            ' Ensure UI shows the loaded rows and a clean new line
    '            Try
    '                objMatrix3.AutoResizeColumns()
    '                Me.SetNewLine3(objForm.UniqueID)
    '            Catch
    '            End Try
    '        End If
    '        objMain.objApplication.StatusBar.SetText("Matrix filled successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText("Matrix Import Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '    End Try

    '    GC.Collect()
    'End Sub

    Sub UpdatePriceNew_ServiceLayer_Itemwise()
        Dim PDocEntry As String = ""
        Dim DocNum As String = objForm.Items.Item("DocNum").Specific.Value.ToString()
        ' Use dataHelper to get the documents to process
        Dim Qry As String = "Select (""DocEntry"")as ""DocEntry"" ,(""DocNum"")as ""DocNum"" from ""@SBO_PL"" where ""U_DocSts""<>'Closed' and ""DocNum""='" & DocNum & "'  Order By ""DocEntry"" Desc "
        Dim dsDocs As DataSet = objDataHelper.getDataSet(Qry)
        Dim dtDocs As DataTable = dsDocs.Tables(0)
        ServiceLayerCon = Me.CompanySession()
        Dim RecordCounter As Integer = 1
        Dim JSONLog As String
        Dim URLLog As String
        For Each docRow As DataRow In dtDocs.Rows
            'PDocEntry = "320"
            PDocEntry = docRow("DocEntry").ToString().Trim()
            Dim docnumm As String = docRow("DocNum").ToString().Trim()
            objMain.objApplication.StatusBar.SetText("Price update in progress, please wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            ''Me.AddToFileIN("Updating Prices for DocNum:" + docnumm & "--------------------------------------------------------------------------" & DateTime.Now)

            'Me.MESSAGE_LOG("Updating Prices for DocNum : " + docnumm, PDocEntry)
            'Dim COUNT = ConfigurationSettings.AppSettings.Get("COUNT").ToString()
            Try
                'PDocEntry = "320"
                Dim query As String = "SELECT T1.""U_ICode"" as ""ItemCode"",T1.""U_UOM"" as ""UomEntry"",T1.""U_NPrice"" as ""Price"" ,T1.""U_PLIST"" as ""Pricelist"",T1.""U_IUOM"" as ""IUOM"",T1.""U_BarCode"" as ""BarCode"",""LineId"",(Case When T1.""U_IUOM""=T1.""U_UOM"" Then T1.""U_NPrice"" Else -1 End) As ""BasePrice"" FROM ""@SBO_PL"" T0 inner join ""@SBO_PL_C0"" T1  " &
                    " on T0.""DocEntry""=T1.""DocEntry"" WHERE T0.""DocEntry"" ='" & PDocEntry & "' and  ifnull(""U_Status"",'No')<>'Yes' and IFNULL(T1.""U_ICode"",'')<>'' "

                query = "SELECT * FROM (" &
            " SELECT T1.""U_ICode"" AS ""ItemCode"", T1.""U_UOM"" AS ""UomEntry"", T1.""U_NPrice"" AS ""Price"", T1.""U_PLIST"" AS ""Pricelist"", " &
            " T1.""U_IUOM"" AS ""IUOM"", T1.""U_BarCode"" AS ""BarCode"", T1.""LineId"", " &
            " CASE WHEN T1.""U_IUOM"" = T1.""U_UOM"" THEN T1.""U_NPrice"" ELSE -1 END AS ""BasePrice"", " &
            " ROW_NUMBER() OVER (PARTITION BY T1.""U_ICode"", T1.""U_PLIST"", T1.""U_UOM"" ORDER BY T1.""LineId"" DESC) AS ""rn"", T1.""U_UPDT"" AS ""StartDate"" " &
            " FROM ""@SBO_PL"" T0 INNER JOIN ""@SBO_PL_C0"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
            " WHERE T0.""DocEntry"" = '" & PDocEntry & "' AND IFNULL(T1.""U_ICode"", '') <> '' AND IFNULL(T1.""U_Status"", 'No') <> 'Yes') sub " &
            "WHERE sub.""rn"" = 1"
                'query = "SELECT * FROM (SELECT T1.""U_ICode"" AS ""ItemCode"", T1.""U_UOM"" AS ""UomEntry"", T1.""U_NPrice"" AS ""Price"", T1.""U_PLIST"" AS ""Pricelist"" " &
                '    " , T1.""U_IUOM"" AS ""IUOM"", T1.""U_BarCode"" AS ""BarCode"", T1.""LineId"", CASE WHEN T1.""U_IUOM"" = T1.""U_UOM"" THEN T1.""U_NPrice"" ELSE -1 END AS ""BasePrice"" " &
                '    " , ROW_NUMBER() OVER (PARTITION BY T1.""U_ICode"", T1.""U_PLIST"", T1.""U_UOM"" ORDER BY T1.""LineId"" DESC) AS ""rn"",T1.""U_UPDT"" As ""StartDate"" FROM ""@SBO_PL"" T0 " &
                '    " INNER JOIN ""@SBO_PL_C0"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" WHERE T0.""DocEntry"" = '" & PDocEntry & "' AND IFNULL(""U_Status"", 'No') <> 'Yes' AND IFNULL(T1.""U_ICode"", '') <> '') sub "
                'query &= "WHERE sub.""rn"" = 1"

                Dim dsPriceList As DataSet = objDataHelper.getDataSet(query)
                Dim dtPriceList As DataTable = dsPriceList.Tables(0)
                If dtPriceList.Rows.Count = 0 Then
                    Dim UpdateQuery1 As String = "UPDATE ""@SBO_PL"" SET ""U_DocSts""='Closed' Where ""DocNum""='" & docnumm & "'"
                    objDataHelper.ExecuteNonQuery(UpdateQuery1)
                End If

                ' Group all price/UOM/price list data by item
                Dim itemData As New Dictionary(Of String, List(Of (PriceList As String, UomEntry As Integer, Price As Double, BarCode As String, LineId As Integer, IUOM As Integer, BasePrice As Double, StartDate As String)))
                For Each row As DataRow In dtPriceList.Rows
                    Dim itemCode As String = row("ItemCode").ToString()
                    Dim priceList As String = row("Pricelist").ToString()
                    Dim price As Double = Convert.ToDouble(row("Price"))
                    Dim barCode As String = row("BarCode").ToString()
                    Dim uomEntry As Integer = CInt(row("UomEntry"))
                    Dim lineid As Integer = CInt(row("LineId"))
                    Dim IUOM As Integer = CInt(row("IUOM"))
                    Dim BasePrice As Double = Convert.ToDouble(row("BasePrice"))
                    Dim StartDate As String = row("StartDate").ToString()

                    If Not itemData.ContainsKey(itemCode) Then
                        itemData(itemCode) = New List(Of (String, Integer, Double, String, Integer, Integer, Double, String))()
                    End If
                    itemData(itemCode).Add((priceList, uomEntry, price, barCode, lineid, IUOM, BasePrice, StartDate))
                Next

                ' For each item, build a single PATCH payload with all price lists and UOMs
                For Each kvp In itemData
                    Dim itemCode = kvp.Key
                    Dim priceListGroups = kvp.Value.GroupBy(Function(x) x.PriceList)

                    ' Build ItemPrices array
                    Dim itemPricesJson As New StringBuilder()
                    itemPricesJson.Append("[")
                    Dim firstPriceList = True
                    For Each plGroup In priceListGroups
                        If Not firstPriceList Then itemPricesJson.Append(",")
                        firstPriceList = False
                        Dim pl = plGroup.Key

                        ' Example loop to read filtered UomPrices values  
                        Dim filteredUomPrices = plGroup.Where(Function(x) x.UomEntry <> x.IUOM).Select(Function(x) New With {x.UomEntry, x.Price, x.StartDate}).ToList()

                        itemPricesJson.Append("{")

                        itemPricesJson.AppendFormat("""PriceList"":{0},", pl)
                        Dim BasePrice = kvp.Value.AsQueryable().Where(Function(x) x.UomEntry = x.IUOM And x.PriceList = pl).Select(Function(x) x.BasePrice).FirstOrDefault()
                        Dim strBaseStartDate As String = ""
                        Try
                            Dim BaseStartDate As Date = kvp.Value.AsQueryable().Where(Function(x) x.UomEntry = x.IUOM And x.PriceList = pl).Select(Function(x) x.StartDate).FirstOrDefault()
                            strBaseStartDate = BaseStartDate.ToString("yyyyMMdd")
                        Catch ex As Exception
                            strBaseStartDate = DateTime.Now.ToString("yyyyMMdd")
                            ''objMain.objApplication.StatusBar.SetText("Start date missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End Try

                        If BasePrice <> 0 Then
                            If strBaseStartDate = "" Then
                                itemPricesJson.AppendFormat("""Price"":{0},", BasePrice)
                            Else
                                itemPricesJson.AppendFormat("""Price"":{0},""U_UPDT"":""{1}"",""U_PSYNC"": ""N"",", BasePrice, strBaseStartDate)
                            End If
                        End If

                        Dim firstUom = True
                        itemPricesJson.AppendFormat("""UoMPrices"":[")
                        For Each uomPrice In filteredUomPrices
                            Dim uomEntry As Integer = uomPrice.UomEntry
                            Dim price As Double = uomPrice.Price
                            Dim strStartDate As String = ""
                            Try
                                Dim StartDate As Date = uomPrice.StartDate
                                strStartDate = StartDate.ToString("yyyyMMdd")
                            Catch ex As Exception
                                strStartDate = DateTime.Now.ToString("yyyyMMdd")
                                ''objMain.objApplication.StatusBar.SetText("Start date missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End Try
                            If Not firstUom Then itemPricesJson.Append(",")
                            firstUom = False
                            ' itemPricesJson.AppendFormat("{{""UoMEntry"":{0},""Price"":{1}}}", uomEntry, price)

                            itemPricesJson.AppendFormat("{{""UoMEntry"":{0},""Price"":{1},""U_UPDT"":{2},""Auto"":""tNO""}}", uomEntry, price, strStartDate)


                        Next
                        itemPricesJson.Append("]")
                        itemPricesJson.Append("}")
                    Next
                    itemPricesJson.Append("]")

                    Dim patchJson As String = $"{{""ItemPrices"":{itemPricesJson.ToString()}}}"
                    Dim patchUrl = BaseURL & "Items('" & itemCode & "')"

                    JSONLog = patchJson
                    URLLog = patchUrl
                    RecordCounter = RecordCounter + 1
                    Try
                        If RecordCounter Mod 50 = 0 Then
                            ServiceLayerCon = Me.CompanySession()
                        End If
                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText("Service Layer session issue" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    End Try

                    Dim patchResult = Me.UpdatePrices_PATCH_Response(ServiceLayerCon, patchUrl, patchJson)
                    objMain.objApplication.StatusBar.SetText("Úpdating Price for Itemcode: " & itemCode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                    ' Update all lines for this item
                    For Each up In kvp.Value
                        If patchResult.StatusCode = HttpStatusCode.NoContent Or patchResult.StatusCode = HttpStatusCode.OK Then
                            '  Me.WriteLog(PDocEntry, itemCode, up.UomEntry, up.PriceList, up.Price, 1, "Price updated successfully", "Success", docnumm)
                            ' Update status to Yes and remarks using HANA connection
                            Dim updateLineQuery As String = "UPDATE ""@SBO_PL_C0"" SET ""U_Status""='Yes',""U_Remarks""='Price Updated Successfully'  Where ""DocEntry""='" & PDocEntry & "' and ""LineId""='" & up.LineId & "'"
                            objDataHelper.ExecuteNonQuery(updateLineQuery)
                        Else
                            ' Me.WriteLog(up.BarCode, itemCode, up.UomEntry, up.PriceList, up.Price, patchResult.StatusCode, patchResult.ResponseBody, "ERROR", docnumm)
                            ' Update status to No and remarks using HANA connection
                            Dim updateLineQuery As String = "UPDATE ""@SBO_PL_C0"" SET ""U_Status""='No',""U_Remarks""='" & patchResult.ResponseBody.Replace("'", " ") & "'  Where ""DocEntry""='" & PDocEntry & "' and ""LineId""='" & up.LineId & "'"
                            objDataHelper.ExecuteNonQuery(updateLineQuery)

                            ''Me.AddToFileIN("URL:" & URLLog & vbCrLf & " JSON: " & JSONLog & vbCrLf & " Response:" & patchResult.ResponseBody & vbCrLf & "" & System.DateTime.Now & vbCrLf)
                        End If
                    Next
                Next

                ''Me.AddToFileIN("Updating Prices Completed for DocNum:" + docnumm & "--------------------------------------------------------------------------" & DateTime.Now)

                objMain.objApplication.StatusBar.SetText("Price update completed successfully.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                Dim UpdateQuery As String = "UPDATE ""@SBO_PL"" SET ""U_DocSts""='Closed' Where ""DocNum""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "'"
                Dim oRsSubmit As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsSubmit.DoQuery(UpdateQuery)

                Try
                    objMain.objApplication.ActivateMenuItem("1304")

                Catch ex As Exception

                End Try
                If Sucess <> "" Then
                    Dim UpdateQuery1 As String = "UPDATE ""@SBO_PL"" SET ""U_Success""='" & Sucess & "' Where ""DocNum""='" & docnumm & "'"
                    objDataHelper.ExecuteNonQuery(UpdateQuery1)
                End If
                If failed <> "" Then
                    Dim UpdateQuery1 As String = "UPDATE ""@SBO_PL"" SET ""U_Failed""='" & failed & "' Where ""DocNum""='" & docnumm & "'"
                    objDataHelper.ExecuteNonQuery(UpdateQuery1)
                End If

            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText("Error while updating prices. Check logs for details.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                ''Me.AddToFileIN(ex.Message)
                ''Me.AddToFileERRIN(ex.Message)
                ''Me.MESSAGE_LOG(ex.Message, PDocEntry)
            End Try
        Next
    End Sub

    Private Sub BackgroundWorker_Loading_DoWork1(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_SalaryJV.DoWork
        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
        oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C0")

        objMatrix3 = objForm.Items.Item("Mtx3").Specific
        For i As Integer = 1 To objMatrix3.VisualRowCount
            DisplayProgressBar(objMatrix3.VisualRowCount, i, "Updating Items")
            Dim Barcode As String = objMatrix3.Columns.Item("BCode").Cells.Item(i).Specific.Value.ToString().Trim()
            Dim Item As String = objMatrix3.Columns.Item("IName").Cells.Item(i).Specific.Value.ToString().Trim()
            Dim UOM As Double = CDbl(objMatrix3.Columns.Item("UOM").Cells.Item(i).Specific.Value)
            Dim CurrectCost As String = objMatrix3.Columns.Item("CCost").Cells.Item(i).Specific.Value.ToString
            Dim CurrentSellingPrice As String = objMatrix3.Columns.Item("CSPrice").Cells.Item(i).Specific.Value.ToString
            Dim CurrentMargain As String = objMatrix3.Columns.Item("CMarg").Cells.Item(i).Specific.Value.ToString
            Dim NewSellingPrice As String = objMatrix3.Columns.Item("NSPrice").Cells.Item(i).Specific.Value.ToString

            Dim NewMargin As String = objMatrix3.Columns.Item("NMarg").Cells.Item(i).Specific.Value.ToString
            Dim NewCost As String = objMatrix3.Columns.Item("NCost").Cells.Item(i).Specific.Value.ToString
            Dim CompetitorName As String = objMatrix3.Columns.Item("CName").Cells.Item(i).Specific.Value.ToString
            Dim InvoiceAttachment As String = objMatrix3.Columns.Item("IAttach").Cells.Item(i).Specific.Value.ToString
            '  If item <> "" And status <> "Yes" And Remarks <> "Price Updated Successfully" Then
            Try
                Me.UpdatePrice(Item, UOM, CurrectCost, CurrentSellingPrice, Barcode, objForm.Items.Item("DocNum").Specific.Value.ToString().Trim(), i, oDBs_Head.GetValue("DocEntry", oDBs_Head.Offset))
            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objForm.Items.Item("Submit").Enabled = True
            End Try
            ' End If
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
        Dim UpdateQuery As String = "UPDATE ""@SBO_PL"" SET ""U_DocSts""='Closed' Where ""DocNum""='" & objForm.Items.Item("DocNum").Specific.Value.ToString().Trim() & "'"
        Dim oRsSubmit As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRsSubmit.DoQuery(UpdateQuery)
        objMain.objApplication.ActivateMenuItem("1304")
    End Sub


    Sub UpdatePrice(ByVal ItemCode As String, ByVal UOM As String, ByVal PriceList As String, ByVal SalesPrice As Double, ByVal barcode As String, ByVal DocNum As String, ByVal pvalrow As String, ByVal DocEntry As String)
        Try
            Dim oItemMaster As SAPbobsCOM.Items = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
            Dim oRsPriceList As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_PL")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@SBO_PL_C0")
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

                        End If
                        Exit For
                    End If
                End If

            Next

            Dim RetValue As Integer = oItemMaster.Update()

            '-----------------------------------------

            If RetValue <> 0 Then
                Dim ErrCode As Long
                Dim ErrMsg As String = String.Empty
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


    Function CompanySession() As String
        Try
            Dim securedwebserviceurl = BaseURL & "Login"
            'Dim securedwebserviceurl = "https://192.168.100.18:50000/b1s/v1//Login"

            ServicePointManager.Expect100Continue = True
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11
            ServicePointManager.ServerCertificateValidationCallback = Function() True
            Dim httpWebRequest = CType(WebRequest.Create(securedwebserviceurl), HttpWebRequest)
            httpWebRequest.ContentType = "application/json"
            httpWebRequest.Method = "POST"

            Dim dicLoginUser As Dictionary(Of String, String) = New Dictionary(Of String, String)()
            dicLoginUser.Add("CompanyDB", CompanyDB)
            dicLoginUser.Add("UserName", UserName)
            dicLoginUser.Add("Password", Password)

            Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(dicLoginUser)

            Using streamWriter = New StreamWriter(httpWebRequest.GetRequestStream())
                streamWriter.Write(json)
                streamWriter.Flush()
                streamWriter.Close()
            End Using

            httpWebRequest.ProtocolVersion = HttpVersion.Version10
            Dim httpResponse = CType(httpWebRequest.GetResponse(), HttpWebResponse)

            Using streamReader = New StreamReader(httpResponse.GetResponseStream())
                Dim responseFromServer As String = streamReader.ReadToEnd()
            End Using

            For i As Integer = 0 To httpResponse.Headers.Count - 1
                Dim n As String = httpResponse.Headers.GetKey(i)
                If n <> "Set-Cookie" Then Continue For
                Dim value As String = httpResponse.Headers.[Get](i)
                Return value
                'Me.UpdateDocuemnt(value, AttCode, oTM, Method)
                'Dim Test As String = ""
            Next
        Catch ex As Exception
            ''Me.AddToFileIN("Service Layer Login error : " & ex.Message)
            objMain.objApplication.StatusBar.SetText("Service Layer Login error : " & ex.Message)
        End Try
    End Function

    'Sub DisplayProgressBar(totalRecords As Integer, currentRecord As Integer, Optional taskName As String = "")
    '    If totalRecords = 0 Then Exit Sub

    '    Dim progressBarLength As Integer = 50
    '    Dim progressPercentage As Double = (currentRecord / totalRecords) * 100
    '    Dim filledLength As Integer = CInt((currentRecord / totalRecords) * progressBarLength)

    '    Dim progressBar As String = New String("¦"c, filledLength).PadRight(progressBarLength, " "c)
    '    Dim message As String = If(taskName <> "", "[" & taskName & "] ", "") &
    '                        vbCr & progressBar & " " & Math.Round(progressPercentage, 2) & "%"


    '    objMain.objApplication.StatusBar.SetText(message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

    'End Sub
    Sub DisplayProgressBar(totalRecords As Integer, currentRecord As Integer, Optional taskName As String = "")
        If totalRecords = 0 Then Exit Sub

        Dim progressBarLength As Integer = 50
        Dim progressPercentage As Double = (currentRecord / totalRecords) * 100
        Dim filledLength As Integer = CInt((currentRecord / totalRecords) * progressBarLength)

        Dim progressBar As String = New String("#"c, filledLength).PadRight(progressBarLength, " "c)
        Dim message As String = If(taskName <> "", "[" & taskName & "] ", "") &
                            vbCr & progressBar & " " & Math.Round(progressPercentage, 2) & "%"


        objMain.objApplication.StatusBar.SetText(message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

    End Sub


    Private Function UpdatePrices_PATCH_Response(ByVal ccv As String, ByVal PriceUpdateURL As String, ByVal JsonString As String) As (StatusCode As HttpStatusCode, ResponseBody As String)

        Dim sep1 As Char() = {"="c}
        Dim lstChar As String() = ccv.Split(sep1, StringSplitOptions.RemoveEmptyEntries)
        Dim sep As Char() = {";"c}
        Dim cookv As String() = lstChar(1).Split(sep, StringSplitOptions.RemoveEmptyEntries)
        Dim cc As String = cookv(0)
        'Dim sUrl = ConfigurationManager.AppSettings("CurrentServiceURL")

        Try
            Dim securedwebserviceurl As String
            securedwebserviceurl = PriceUpdateURL

            ServicePointManager.Expect100Continue = True
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11
            ServicePointManager.ServerCertificateValidationCallback = Function() True
            Dim container As CookieContainer = New CookieContainer()

            container.Add(New Uri(PriceUpdateURL), New Cookie("B1SESSION", cc))

            Dim httpWebRequest = CType(WebRequest.Create(securedwebserviceurl), HttpWebRequest)
            httpWebRequest.CookieContainer = container
            httpWebRequest.ContentType = "application/json"
            httpWebRequest.Method = "PATCH"

            'Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(oTM)

            Using streamWriter = New StreamWriter(httpWebRequest.GetRequestStream())
                streamWriter.Write(JsonString)
                streamWriter.Flush()
                streamWriter.Close()
            End Using

            httpWebRequest.ProtocolVersion = HttpVersion.Version10

            Try
                Using response = CType(httpWebRequest.GetResponse(), Net.HttpWebResponse)
                    Return (response.StatusCode, "")
                End Using
            Catch ex As Net.WebException

                ''Me.AddToFileIN("URL : " & PriceUpdateURL & "JSON : " & JsonString)

                Using resp = CType(ex.Response, Net.HttpWebResponse)
                    Dim respBody As String
                    Using sr = New StreamReader(resp.GetResponseStream())
                        respBody = sr.ReadToEnd()
                    End Using
                    Return (resp.StatusCode, respBody)
                End Using
            End Try

            'Dim httpResponse = CType(httpWebRequest.GetResponse(), HttpWebResponse)
            'Dim k = httpResponse.GetResponseStream()
            'Dim responseFromServer As String
            'Using streamReader = New StreamReader(httpResponse.GetResponseStream())
            '    responseFromServer = streamReader.ReadToEnd()
            'End Using
            'Return responseFromServer

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Service layer posting error : " & ex.Message & "" & System.DateTime.Now & vbCrLf)
            ''Me.AddToFileIN("Service layer posting error : " & ex.Message & "" & System.DateTime.Now & vbCrLf)
        End Try
        Return cnfMsg
    End Function

    Private Sub AddToFileERRIN(ByVal contents As String)
        Try
            Dim LogfilePath As String = ConfigurationManager.AppSettings("ERRLogFile") & "LogFile_ERR" & Today.ToString("dd-MM-yyyy") & ".txt"
            Using TextWriter As StreamWriter = New StreamWriter(LogfilePath, True)
                TextWriter.WriteLine(contents)
            End Using
        Catch ex As Exception
        End Try
    End Sub

    Sub SetNewLine1(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C0")
            objMatrix1 = objForm.Items.Item("Mtx1").Specific
            objMatrix1.AddRow()
            oDBs_Details1.SetValue("LineID", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
            oDBs_Details1.SetValue("U_TNXICD", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_TNXNM", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_TNXGRPO", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_ULPRC", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_PLPRC", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_SCOST", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_OMRG", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_NMRG", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_UOM", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_FTYP", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_OBJT", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_OBJT", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_SPTP", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_FDate", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_FTime", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_ULPRCVAT", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_PLPRCVAT", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_ORPVAT", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_NRPVAT", oDBs_Details1.Offset, "")
            objMatrix1.AutoResizeColumns()
            objMatrix1.SetLineData(objMatrix1.VisualRowCount)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine2(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C1")
            objMatrix2 = objForm.Items.Item("Mtx2").Specific
            objMatrix2.AddRow()
            oDBs_Details2.SetValue("LineID", oDBs_Details2.Offset, objMatrix2.VisualRowCount)
            oDBs_Details2.SetValue("U_BrCode", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_ItemCode", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_ITNM", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_ITGP", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_CPrice", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_NwPrice", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_Cost", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_LPrice", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_NSNG", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_FSNG", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_NMGP", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_FLGP", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_EXPQ", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_ACost", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_ACSL", oDBs_Details2.Offset, "")

            oDBs_Details2.SetValue("U_AFSL", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_FGPV", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_NGPV", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_PRLS", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_MUpdated", oDBs_Details2.Offset, "No")
            oDBs_Details2.SetValue("U_UOM", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_PUpdate", oDBs_Details2.Offset, "No")
            oDBs_Details2.SetValue("U_PUS", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_PDN", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_PS", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_SLOT", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_CRNT", oDBs_Details2.Offset, "")

            objMatrix2.Columns.Item("BrCode").Editable = True
            objMatrix2.Columns.Item("Col_1").Editable = True
            objMatrix2.Columns.Item("Col_4").Editable = True
            objMatrix2.Columns.Item("MUpdated").Editable = True
            objMatrix2.Columns.Item("UOM").Editable = True
            objMatrix2.Columns.Item("PUpdate").Editable = True
            objMatrix2.Columns.Item("PUS").Editable = True
            objMatrix2.Columns.Item("PDN").Editable = True
            objMatrix2.Columns.Item("PS").Editable = True
            objMatrix2.Columns.Item("Col_14").Editable = True
            objMatrix2.Columns.Item("Col_13").Editable = True



            objMatrix2.AutoResizeColumns()
            objMatrix2.SetLineData(objMatrix2.VisualRowCount)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine3(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C2")
            objMatrix3 = objForm.Items.Item("Mtx3").Specific
            objMatrix3.AddRow()
            oDBs_Details3.SetValue("LineID", oDBs_Details3.Offset, objMatrix3.VisualRowCount)
            oDBs_Details3.SetValue("U_BCode", oDBs_Details3.Offset, "")
            oDBs_Details3.SetValue("U_ItemCode", oDBs_Details3.Offset, "")
            oDBs_Details3.SetValue("U_IName", oDBs_Details3.Offset, "")
            oDBs_Details3.SetValue("U_UOM", oDBs_Details3.Offset, "")
            oDBs_Details3.SetValue("U_CCost", oDBs_Details3.Offset, "")
            oDBs_Details3.SetValue("U_CSPrice", oDBs_Details3.Offset, "")
            oDBs_Details3.SetValue("U_CMarg", oDBs_Details3.Offset, "")
            oDBs_Details3.SetValue("U_NSPrice", oDBs_Details3.Offset, "")
            oDBs_Details3.SetValue("U_NMarg", oDBs_Details3.Offset, "")
            oDBs_Details3.SetValue("U_NCost", oDBs_Details3.Offset, "")
            oDBs_Details3.SetValue("U_CName", oDBs_Details3.Offset, "")
            oDBs_Details3.SetValue("U_IAttach", oDBs_Details3.Offset, "")
            oDBs_Details3.SetValue("U_LPhPrice", oDBs_Details3.Offset, "")
            oDBs_Details3.SetValue("U_LGRPPrice", oDBs_Details3.Offset, "")
            ''oDBs_Details3.SetValue("U_TPRT", oDBs_Details3.Offset, "")
            objMatrix3.AutoResizeColumns()
            objMatrix3.SetLineData(objMatrix3.VisualRowCount)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine4(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
            oDBs_Details4 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C3")
            objMatrix4 = objForm.Items.Item("Mtx4").Specific
            objMatrix4.AddRow()
            oDBs_Details4.SetValue("LineID", oDBs_Details4.Offset, objMatrix4.VisualRowCount)
            oDBs_Details4.SetValue("U_BrCode", oDBs_Details4.Offset, "")
            oDBs_Details4.SetValue("U_ItemCode", oDBs_Details4.Offset, "")
            oDBs_Details4.SetValue("U_ItemN", oDBs_Details4.Offset, "")
            oDBs_Details4.SetValue("U_Qty", oDBs_Details4.Offset, "")
            oDBs_Details4.SetValue("U_ICost", oDBs_Details4.Offset, "")
            oDBs_Details4.SetValue("U_NSell", oDBs_Details4.Offset, "")
            oDBs_Details4.SetValue("U_SPrice", oDBs_Details4.Offset, "")
            oDBs_Details4.SetValue("U_DiffSP", oDBs_Details4.Offset, "")
            oDBs_Details4.SetValue("U_DSPNS", oDBs_Details4.Offset, "")
            oDBs_Details4.SetValue("U_TPRC", oDBs_Details4.Offset, "")
            oDBs_Details4.SetValue("U_TPRS", oDBs_Details4.Offset, "")
            oDBs_Details4.SetValue("U_FundA", oDBs_Details4.Offset, "")
            oDBs_Details4.SetValue("U_NGP", oDBs_Details4.Offset, "")
            oDBs_Details4.SetValue("U_SGP", oDBs_Details4.Offset, "")
            oDBs_Details4.SetValue("U_Supp", oDBs_Details4.Offset, "")
            oDBs_Details4.SetValue("U_Remark", oDBs_Details4.Offset, "")
            oDBs_Details4.SetValue("U_UOM1", oDBs_Details4.Offset, "")
            oDBs_Details4.SetValue("U_CRNT", oDBs_Details4.Offset, "")

            objMatrix4.Columns.Item("BrCode").Editable = True
            objMatrix4.Columns.Item("Qty").Editable = True
            objMatrix4.Columns.Item("SPrice").Editable = True
            objMatrix4.Columns.Item("FundA").Editable = True
            objMatrix4.Columns.Item("Supp").Editable = True
            objMatrix4.Columns.Item("UOM1").Editable = True
            objMatrix4.Columns.Item("Remark").Editable = True
            objMatrix4.Columns.Item("MUpdated").Editable = True
            objMatrix4.Columns.Item("PUpdate").Editable = True
            objMatrix4.Columns.Item("PUS").Editable = True
            objMatrix4.Columns.Item("PriceR").Editable = True
            objMatrix4.Columns.Item("PRS").Editable = True
            objMatrix4.Columns.Item("PDN").Editable = True
            objMatrix4.Columns.Item("PS").Editable = True
            objMatrix4.Columns.Item("Col_0").Editable = True

            objMatrix4.AutoResizeColumns()
            objMatrix4.SetLineData(objMatrix4.VisualRowCount)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine5(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
            oDBs_Details5 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C4")
            objMatrix5 = objForm.Items.Item("Mtx5").Specific
            objMatrix5.AddRow()
            oDBs_Details5.SetValue("LineID", oDBs_Details5.Offset, objMatrix5.VisualRowCount)
            oDBs_Details5.SetValue("U_TPath", oDBs_Details5.Offset, "")
            oDBs_Details5.SetValue("U_FName", oDBs_Details5.Offset, "")
            oDBs_Details5.SetValue("U_ADate", oDBs_Details5.Offset, "")
            oDBs_Details5.SetValue("U_FText", oDBs_Details5.Offset, "")
            objMatrix5.AutoResizeColumns()
            objMatrix5.SetLineData(objMatrix5.VisualRowCount)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub SetNewLine6(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
            oDBs_Details6 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C5")
            objMatrix6 = objForm.Items.Item("Mtx6").Specific
            objMatrix6.AddRow()
            oDBs_Details6.SetValue("LineID", oDBs_Details6.Offset, objMatrix6.VisualRowCount)
            oDBs_Details6.SetValue("U_BCode", oDBs_Details6.Offset, "")
            oDBs_Details6.SetValue("U_ICode", oDBs_Details6.Offset, "")
            oDBs_Details6.SetValue("U_IName", oDBs_Details6.Offset, "")
            oDBs_Details6.SetValue("U_CSlng", oDBs_Details6.Offset, "")
            oDBs_Details6.SetValue("U_NSlng", oDBs_Details6.Offset, "")
            oDBs_Details6.SetValue("U_CCost", oDBs_Details6.Offset, "")
            oDBs_Details6.SetValue("U_NCost", oDBs_Details6.Offset, "")
            oDBs_Details6.SetValue("U_CMarg", oDBs_Details6.Offset, "")
            oDBs_Details6.SetValue("U_NMrgn", oDBs_Details6.Offset, "")
            oDBs_Details6.SetValue("U_CQTT", oDBs_Details6.Offset, "")
            oDBs_Details6.SetValue("U_Unit", oDBs_Details6.Offset, "")
            oDBs_Details6.SetValue("U_PUpdate", oDBs_Details6.Offset, "No")
            oDBs_Details6.SetValue("U_PUS", oDBs_Details6.Offset, "")
            oDBs_Details6.SetValue("U_IAttach", oDBs_Details6.Offset, "")
            oDBs_Details6.SetValue("U_MUpdated", oDBs_Details6.Offset, "")

            objMatrix6.AutoResizeColumns()
            objMatrix6.SetLineData(objMatrix6.VisualRowCount)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub SetNewLine7(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
            oDBs_Details7 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C6")
            objMatrix7 = objForm.Items.Item("Mtx7").Specific
            objMatrix7.AddRow()
            oDBs_Details7.SetValue("LineID", oDBs_Details7.Offset, objMatrix7.VisualRowCount)
            oDBs_Details7.SetValue("U_BCode", oDBs_Details7.Offset, "")
            oDBs_Details7.SetValue("U_ITCD", oDBs_Details7.Offset, "")
            oDBs_Details7.SetValue("U_ITNM", oDBs_Details7.Offset, "")
            oDBs_Details7.SetValue("U_NPRC", oDBs_Details7.Offset, "")
            oDBs_Details7.SetValue("U_UOM", oDBs_Details7.Offset, "")
            oDBs_Details7.SetValue("U_PUpdate", oDBs_Details7.Offset, "No")
            oDBs_Details7.SetValue("U_PUS", oDBs_Details7.Offset, "")

            objMatrix7.AutoResizeColumns()
            objMatrix7.SetLineData(objMatrix7.VisualRowCount)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub SetNewLine8(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
            oDBs_Details8 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C7")
            objMatrix8 = objForm.Items.Item("Mtx8").Specific
            objMatrix8.AddRow()
            oDBs_Details8.SetValue("LineID", oDBs_Details8.Offset, objMatrix8.VisualRowCount)
            oDBs_Details8.SetValue("U_BCode", oDBs_Details8.Offset, "")
            oDBs_Details8.SetValue("U_ITCD", oDBs_Details8.Offset, "")
            oDBs_Details8.SetValue("U_ITNM", oDBs_Details8.Offset, "")
            oDBs_Details8.SetValue("U_UOM", oDBs_Details8.Offset, "")
            oDBs_Details8.SetValue("U_QTY", oDBs_Details8.Offset, "")
            oDBs_Details8.SetValue("U_OMAC", oDBs_Details8.Offset, "")
            oDBs_Details8.SetValue("U_NMAC", oDBs_Details8.Offset, "")
            oDBs_Details8.SetValue("U_ODS", oDBs_Details8.Offset, "")
            oDBs_Details8.SetValue("U_OMA", oDBs_Details8.Offset, "")
            oDBs_Details8.SetValue("U_OMG", oDBs_Details8.Offset, "")
            oDBs_Details8.SetValue("U_NMAR", oDBs_Details8.Offset, "")
            oDBs_Details8.SetValue("U_DIFF", oDBs_Details8.Offset, "")
            oDBs_Details8.SetValue("U_NSP", oDBs_Details8.Offset, "")
            oDBs_Details8.SetValue("U_NMNS", oDBs_Details8.Offset, "")
            oDBs_Details8.SetValue("U_NMP", oDBs_Details8.Offset, "")
            oDBs_Details8.SetValue("U_PUpdate", oDBs_Details8.Offset, "No")
            oDBs_Details8.SetValue("U_PUS", oDBs_Details8.Offset, "")

            objMatrix8.AutoResizeColumns()
            objMatrix8.SetLineData(objMatrix8.VisualRowCount)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine9(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
            oDBs_Details9 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C8")
            objMatrix9 = objForm.Items.Item("Mtx9").Specific
            objMatrix9.AddRow()
            oDBs_Details9.SetValue("LineID", oDBs_Details9.Offset, objMatrix9.VisualRowCount)
            oDBs_Details9.SetValue("U_BCode", oDBs_Details9.Offset, "")
            oDBs_Details9.SetValue("U_BName", oDBs_Details9.Offset, "")
            oDBs_Details9.SetValue("U_ITCD", oDBs_Details9.Offset, "")
            oDBs_Details9.SetValue("U_UOM", oDBs_Details9.Offset, "")
            oDBs_Details9.SetValue("U_OSLP", oDBs_Details9.Offset, "")
            oDBs_Details9.SetValue("U_NSLP", oDBs_Details9.Offset, "")
            oDBs_Details9.SetValue("U_PUpdate", oDBs_Details9.Offset, "No")
            oDBs_Details9.SetValue("U_PUS", oDBs_Details9.Offset, "")

            objMatrix9.Columns.Item("Col_0").Editable = True
            objMatrix9.Columns.Item("Col_2").Editable = True
            objMatrix9.Columns.Item("Col_3").Editable = True
            objMatrix9.Columns.Item("Col_4").Editable = True
            objMatrix9.Columns.Item("PUS").Editable = True
            objMatrix9.Columns.Item("PUpdate").Editable = True

            objMatrix9.AutoResizeColumns()
            objMatrix9.SetLineData(objMatrix9.VisualRowCount)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    ' --- Add this helper near LockTabs / SetTabsByUpdateType methods ---
    Private Sub ApplyTabsForCurrentType(ByVal f As SAPbouiCOM.Form)
        Try
            f.Freeze(True)
            Dim oDBsHead As SAPbouiCOM.DBDataSource = Nothing
            Try
                oDBsHead = f.DataSources.DBDataSources.Item("@TNX_PCRT")
            Catch
            End Try

            Dim updateType As String = ""
            Try : updateType = If(oDBsHead IsNot Nothing, oDBsHead.GetValue("U_TNXPUTP", 0).Trim(), "") : Catch : End Try

            ' Disable all tabs first
            Try : f.Items.Item("Tab1").Enabled = False : Catch : End Try
            Try : f.Items.Item("Tab2").Enabled = False : Catch : End Try
            Try : f.Items.Item("Tab3").Enabled = False : Catch : End Try
            Try : f.Items.Item("Tab4").Enabled = False : Catch : End Try
            Try : f.Items.Item("Tab6").Enabled = False : Catch : End Try
            Try : f.Items.Item("Tab7").Enabled = False : Catch : End Try
            Try : f.Items.Item("Tab8").Enabled = False : Catch : End Try
            Try : f.Items.Item("Tab5").Enabled = True  ' Attachment left enabled
            Catch : End Try

            ' Enable only the tab relevant to the current header type (no Click() to avoid losing focus)
            Select Case updateType
                Case "Cost Fluctuation"
                    Try : f.Items.Item("Tab1").Enabled = True : Catch : End Try
                Case "Flyer Price Changes"
                    Try : f.Items.Item("Tab2").Enabled = True : Catch : End Try
                Case "MOP"
                    Try : f.Items.Item("Tab3").Enabled = True : Catch : End Try
                Case "Temporary Price Reduction"
                    Try : f.Items.Item("Tab4").Enabled = True : Catch : End Try
                Case "Attachment Tab"
                    Try : f.Items.Item("Tab5").Enabled = True : Catch : End Try
                Case "Every day Special Price"
                    Try : f.Items.Item("Tab6").Enabled = True : Catch : End Try
                Case "Bulk Cost upload"
                    Try : f.Items.Item("Tab7").Enabled = True : Catch : End Try
                Case "Bulk Price upload"
                    Try : f.Items.Item("Tab9").Enabled = True : Catch : End Try
                Case "Cost Fluctuation Stock Transfer"
                    Try : f.Items.Item("Tab8").Enabled = True : Catch : End Try
                Case Else
                    Try : f.Items.Item("Tab1").Enabled = True : Catch : End Try
            End Select
        Catch
            ' ignore
        Finally
            Try : f.Freeze(False) : Catch : End Try
        End Try
    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo,
                         ByRef BubbleEvent As Boolean)
        Try
            ' Get the active PCRT form
            objForm = objMain.objApplication.Forms.GetForm("PCRT",
                    objMain.objApplication.Forms.ActiveForm.TypeCount)

            Select Case BusinessObjectInfo.EventType

            '========================================================
            ' AFTER ADD : AUTO APPROVE ONLY WHEN NO TEMPLATE EXISTS
            '========================================================
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD



                    If BusinessObjectInfo.BeforeAction = False AndAlso BusinessObjectInfo.ActionSuccess = True Then
                        Try
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
                            Dim priceType As String = ""
                            Try : priceType = oDBs_Head.GetValue("U_TNXPUTP", 0).Trim() : Catch : End Try
                            Dim appId As String = ""
                            Try : appId = oDBs_Head.GetValue("U_AppID", 0).Trim() : Catch : End Try

                            ' Check approval template existence
                            Dim rsTemp As SAPbobsCOM.Recordset =
                            objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            Dim sql As String =
                            "SELECT COUNT(*) FROM ""@SBO_APPHDR"" T0 " &
                            "INNER JOIN ""@SBO_APPDOC"" T2 ON T2.""Code"" = T0.""Code"" " &
                            "WHERE IFNULL(T0.""U_Active"",'N')='Y' " &
                            "AND IFNULL(T2.""U_Price"",'N')='Y' " &
                            "AND T0.""U_PCR"" = '" & priceType.Replace("'", "''") & "'"

                            rsTemp.DoQuery(sql)

                            Dim templateCount As Integer = 0
                            Try : templateCount = CInt(rsTemp.Fields.Item(0).Value) : Catch : End Try
                            'Commented by Ramya on 11-03-2026
                            '==============================
                            ' NO TEMPLATE → AUTO APPROVE
                            '==============================
                            'If templateCount = 0 Then

                            '    ' Set status in datasource
                            '    Try
                            '        oDBs_Head.SetValue("U_TNXSTUS", oDBs_Head.Offset, "Approved")
                            '    Catch
                            '    End Try

                            '    ' Persist status in DB
                            '    Try
                            '        Dim docEntry As String = ""
                            '        Try : docEntry = oDBs_Head.GetValue("DocEntry", 0).Trim() : Catch : End Try

                            '        If Not String.IsNullOrWhiteSpace(docEntry) Then
                            '            Dim rsUpd As SAPbobsCOM.Recordset =
                            '            objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            '            rsUpd.DoQuery(
                            '            "UPDATE ""@TNX_PCRT"" SET ""U_TNXSTUS""='Approved' WHERE ""DocEntry""=" & docEntry)
                            '        End If
                            '    Catch
                            '    End Try

                            ' SEND MARKETING EMAIL  ✅ ADDED
                            ' ==============================
                            'Try
                            '    Dim DocNum As String = ""
                            '    Try : DocNum = oDBs_Head.GetValue("DocNum", 0).Trim() : Catch : End Try

                            '    If Not String.IsNullOrWhiteSpace(DocNum) Then
                            '        Dim oUtil As New Utilities()
                            '        oUtil.MarketingEmail(DocNum)
                            '    End If
                            'Catch
                            'End Try

                            '    ' Delay UI lock to avoid SAP timing issues
                            '    Dim oneShot As New System.Windows.Forms.Timer()
                            '    oneShot.Interval = 350

                            '    AddHandler oneShot.Tick,
                            '    Sub()
                            '        Try
                            '            oneShot.Stop()

                            '            Dim f As SAPbouiCOM.Form =
                            '                objMain.objApplication.Forms.GetForm("PCRT",
                            '                objMain.objApplication.Forms.ActiveForm.TypeCount)

                            '            ' Force view mode
                            '            Try : f.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE : Catch : End Try

                            '            ' Disable all items
                            '            For i As Integer = 0 To f.Items.Count - 1
                            '                Try : f.Items.Item(i).Enabled = False : Catch : End Try
                            '            Next

                            '            ' Disable Add/Update menus
                            '            Try : f.EnableMenu("1281", False) : Catch : End Try
                            '            Try : f.EnableMenu("1282", False) : Catch : End Try

                            '            ' Disable submit button
                            '            Try : f.Items.Item("Submit").Enabled = False : Catch : End Try

                            '            objMain.objApplication.StatusBar.SetText(
                            '                "Auto-approved (no template). Document locked.",
                            '                SAPbouiCOM.BoMessageTime.bmt_Short,
                            '                SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                            '        Finally
                            '            Try : oneShot.Dispose() : Catch : End Try
                            '        End Try
                            '    End Sub

                            '    oneShot.Start()
                            'End If





                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(
                            "Auto-approval error : " & ex.Message,
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If


                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                    If BusinessObjectInfo.BeforeAction = False AndAlso BusinessObjectInfo.ActionSuccess = True Then


                        objForm.Freeze(True)

                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")

                        objMatrix1 = objForm.Items.Item("Mtx1").Specific
                        objMatrix2 = objForm.Items.Item("Mtx2").Specific
                        objMatrix3 = objForm.Items.Item("Mtx3").Specific
                        objMatrix4 = objForm.Items.Item("Mtx4").Specific
                        objMatrix5 = objForm.Items.Item("Mtx5").Specific
                        objMatrix6 = objForm.Items.Item("Mtx6").Specific
                        objMatrix7 = objForm.Items.Item("Mtx7").Specific
                        objMatrix8 = objForm.Items.Item("Mtx8").Specific
                        objMatrix9 = objForm.Items.Item("Mtx9").Specific
                        ''SetMatrixTotals(objForm)
                        Dim sType As String = oDBs_Head.GetValue("U_TNXPUTP", 0).Trim()
                        objForm.EnableMenu("5890", True)
                        objForm.Items.Item("Item_28").Enabled = True
                        'Enable button only for Flyer
                        objForm.Items.Item("Item_8").Enabled = (sType = "Flyer Price Changes")

                        Select Case sType

                            Case "Cost Fluctuation"
                                objForm.PaneLevel = 1
                                objForm.EnableMenu("5890", True)
                                objForm.Settings.MatrixUID = "Mtx1"
                                objForm.Items.Item("Item_8").Visible = False

                            Case "Flyer Price Changes"
                                objForm.PaneLevel = 2
                                objForm.EnableMenu("5890", True)
                                objForm.Settings.MatrixUID = "Mtx2"
                                objForm.Items.Item("Item_8").Visible = True

                                Try
                                    objMatrix2 = objForm.Items.Item("Mtx2").Specific

                                    objMatrix2.Columns.Item("Cost").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                                    objMatrix2.Columns.Item("Col_0").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                                    objMatrix2.Columns.Item("Col_1").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                                    objMatrix2.Columns.Item("Col_1").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                                    objMatrix2.Columns.Item("Col_2").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                                    objMatrix2.Columns.Item("Col_3").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                                    objMatrix2.Columns.Item("Col_4").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                                    objMatrix2.Columns.Item("Col_5").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                                    objMatrix2.Columns.Item("Col_6").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                                    objMatrix2.Columns.Item("Col_7").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                                    objMatrix2.Columns.Item("Col_8").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                                    objMatrix2.Columns.Item("Col_9").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                                    objMatrix2.Columns.Item("Col_10").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                                    objMatrix2.Columns.Item("Col_14").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

                                Catch ex As Exception
                                    objMain.objApplication.StatusBar.SetText(ex.Message)
                                End Try

                            Case "MOP"
                                objForm.PaneLevel = 3
                                objForm.EnableMenu("5890", True)
                                objForm.Settings.MatrixUID = "Mtx3"
                                objForm.Items.Item("Item_8").Visible = False

                            Case "Temporary Price Reduction"
                                objForm.PaneLevel = 4
                                objForm.EnableMenu("5890", True)
                                objForm.Settings.MatrixUID = "Mtx4"
                                objForm.Items.Item("Item_8").Visible = False
                            Case "Every day Special Price"
                                objForm.PaneLevel = 5
                                objForm.EnableMenu("5890", True)
                                objForm.Settings.MatrixUID = "Mtx5"
                                objForm.Items.Item("Item_8").Visible = False
                            Case "Bulk Cost Update"
                                objForm.PaneLevel = 6
                                objForm.EnableMenu("5890", True)
                                objForm.Settings.MatrixUID = "Mtx7"
                                objForm.Items.Item("Item_8").Visible = False
                            Case "Cost Fluctuation Stock Transfer"
                                objForm.PaneLevel = 7
                                objForm.EnableMenu("5890", True)
                                objForm.Settings.MatrixUID = "Mtx8"
                                objForm.Items.Item("Item_8").Visible = False
                            Case "Bulk Price Update"
                                objForm.PaneLevel = 8
                                objForm.EnableMenu("5890", True)
                                objForm.Settings.MatrixUID = "Mtx9"
                                objForm.Items.Item("Item_8").Visible = False
                            Case "Attachment Tab"
                                objForm.PaneLevel = 9
                                objForm.EnableMenu("5890", True)
                                objForm.Settings.MatrixUID = "Mtx6"
                                objForm.Items.Item("Item_8").Visible = False

                        End Select

                        objForm.Freeze(False)



                    End If

                    If BusinessObjectInfo.BeforeAction = False AndAlso BusinessObjectInfo.ActionSuccess = True Then
                        objForm.Freeze(True)
                        ' Enable Form Settings
                        Try : objForm.EnableMenu("1282", True) : Catch : End Try
                        Try : objForm.EnableMenu("5890", True) : Catch : End Try

                        Try
                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")

                            objMatrix1 = objForm.Items.Item("Mtx1").Specific
                            objMatrix2 = objForm.Items.Item("Mtx2").Specific
                            objMatrix3 = objForm.Items.Item("Mtx3").Specific
                            objMatrix4 = objForm.Items.Item("Mtx4").Specific
                            objMatrix5 = objForm.Items.Item("Mtx5").Specific
                            objMatrix6 = objForm.Items.Item("Mtx6").Specific
                            objMatrix7 = objForm.Items.Item("Mtx7").Specific
                            objMatrix8 = objForm.Items.Item("Mtx8").Specific
                            objMatrix9 = objForm.Items.Item("Mtx9").Specific
                            objForm.Items.Item("Item_28").Enabled = True

                            Try
                                'Dim oChk As SAPbouiCOM.CheckBox = CType(objForm.Items.Item("Check").Specific, SAPbouiCOM.CheckBox)
                                'Dim strCheck As String = oDBs_Head.GetValue("U_Check", 0).Trim()
                                ''If oChk.Checked = True Then
                                'If strCheck = "Y" Then
                                '    objForm.Items.Item("FDTT").Enabled = True
                                '    objForm.Items.Item("TDTT").Enabled = True
                                'Else

                                '    objForm.Items.Item("FDTT").Enabled = False
                                '    objForm.Items.Item("TDTT").Enabled = False
                                'End If
                            Catch ex As Exception

                            End Try

                            Dim status As String = ""
                            Try
                                status = oDBs_Head.GetValue("U_TNXSTUS", 0).Trim().ToUpper()
                            Catch
                            End Try
                            status = status.Replace(" ", "")

                            If status = "REJECTED" Or status = "R" Then
                                Try : objForm.Items.Item("Submit").Enabled = True : Catch : End Try
                            End If
                            'If status = "APPROVED" OrElse status = "A" OrElse status = "Submitted" OrElse status = "S" Then


                            '    Try : objForm.Items.Item("Submit").Enabled = False : Catch : End Try


                            '    Try
                            '        If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                            '            objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                            '        End If
                            '    Catch
                            '    End Try

                            '    ' Lock all fields
                            '    For i As Integer = 0 To objForm.Items.Count - 1
                            '        Try : objForm.Items.Item(i).Enabled = False : Catch : End Try
                            '    Next

                            '    Try : objForm.EnableMenu("1281", False) : Catch : End Try

                            'Else

                            '    Try : objForm.EnableMenu("1281", True) : Catch : End Try
                            '    Try : objForm.Items.Item("Submit").Enabled = True : Catch : End Try

                            '    ' Keep form in OK_MODE for editing
                            '    Try
                            '        If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            '            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                            '        End If
                            '    Catch
                            '    End Try
                            'End If

                            If status = "APPROVED" OrElse status = "A" OrElse status = "Submitted" OrElse status = "S" OrElse status = "N" OrElse status = "Closed" OrElse status = "C" OrElse status = "CLOSED" OrElse status = "C" Then
                                objMatrix1 = objForm.Items.Item("Mtx1").Specific
                                ' Disable Submit button
                                Try : objForm.Items.Item("Submit").Enabled = False : Catch : End Try
                                objForm.Items.Item("Item_28").Enabled = True
                                ' Keep form in VIEW MODE

                                If oDBs_Head.GetValue("U_TNXPUTP", 0).Trim() = "Cost Fluctuation" Then
                                    Try
                                        objMatrix1.Columns.Item("Col_6").Editable = True
                                    Catch
                                    End Try
                                ElseIf oDBs_Head.GetValue("U_TNXPUTP", 0).Trim() = "Temporary Price Reduction" Then
                                    Try
                                        objMatrix1.Columns.Item("Col_0").Editable = True
                                    Catch
                                    End Try
                                ElseIf oDBs_Head.GetValue("U_TNXPUTP", 0).Trim() = "Flyer Price Changes" Then
                                    Try
                                        objForm.Items.Item("FOUT").Enabled = True
                                    Catch
                                    End Try
                                Else

                                    Try
                                        If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        End If
                                    Catch
                                    End Try

                                End If


                                ' Disable all items (Optimized: Removed redundant For loop wrapper)

                                Dim status1 As String = ""
                                status1 = oDBs_Head.GetValue("U_TNXPUTP", 0).Trim().ToUpper()
                                Dim loginUser As String = objMain.objApplication.Company.UserName.Trim().ToUpper()

                                'Dim oRS2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                'Dim sSQL2 As String = "SELECT ""U_NAME"" from OUSR Where ""USER_CODE"" = '" & loginUser.Replace("'", "''").ToUpper() & "'"
                                'oRS2.DoQuery(sSQL2)
                                'Dim loginUSERNAME As String = ""
                                'If Not oRS2.EoF Then
                                '    loginUSERNAME = oRS2.Fields.Item("U_NAME").Value.ToString().Trim().ToUpper()
                                'End If

                                If (status1 = "Cost Fluctuation" Or status1 = "COST FLUCTUATION") And (status = "A" OrElse status = "C" OrElse status = "N") Then

                                    objMatrix1.Columns.Item("UOM").Editable = False

                                    objMatrix1.Columns.Item("NRP").Editable = False
                                    objMatrix1.Columns.Item("PUpdate").Editable = False
                                    objMatrix1.Columns.Item("PUS").Editable = False
                                    objMatrix1.Columns.Item("Col_0").Editable = False
                                    objMatrix1.Columns.Item("Col_1").Editable = False
                                    objMatrix1.Columns.Item("Col_2").Editable = True
                                    objMatrix1.Columns.Item("Col_3").Editable = True
                                    objMatrix1.Columns.Item("Col_4").Editable = True
                                    objMatrix1.Columns.Item("Col_6").Editable = True
                                    objMatrix1.Columns.Item("Col_7").Editable = True
                                    objForm.Items.Item("TNXBRC").Enabled = False
                                    objForm.Items.Item("TNXPUTP").Enabled = False
                                    objForm.Items.Item("TNXUSR").Enabled = False
                                    objForm.Items.Item("PName").Enabled = False

                                ElseIf (status1 = "Temporary Price Reduction" Or status1 = "TEMPORARY PRICE REDUCTION") And (status = "A" OrElse status = "C" OrElse status = "N") Then
                                    objMatrix4.Columns.Item("BrCode").Editable = False
                                    objMatrix4.Columns.Item("Qty").Editable = False

                                    objMatrix4.Columns.Item("SPrice").Editable = False
                                    objMatrix4.Columns.Item("FundA").Editable = False
                                    objMatrix4.Columns.Item("Supp").Editable = False
                                    objMatrix4.Columns.Item("UOM1").Editable = False
                                    objMatrix4.Columns.Item("Remark").Editable = False
                                    objMatrix4.Columns.Item("MUpdated").Editable = False
                                    objMatrix4.Columns.Item("PUpdate").Editable = False
                                    objMatrix4.Columns.Item("PUS").Editable = False
                                    objMatrix4.Columns.Item("PriceR").Editable = False
                                    objMatrix4.Columns.Item("PRS").Editable = False
                                    objMatrix4.Columns.Item("PDN").Editable = False
                                    objMatrix4.Columns.Item("PS").Editable = False

                                    objMatrix4.Columns.Item("Col_0").Editable = True
                                    objForm.Items.Item("TNXBRC").Enabled = False
                                    objForm.Items.Item("TNXPUTP").Enabled = False
                                    objForm.Items.Item("TNXUSR").Enabled = False
                                    objForm.Items.Item("PName").Enabled = False

                                ElseIf (status1 = "Flyer Price Changes" Or status1 = "FLYER PRICE CHANGES") And (status = "A" OrElse status = "N" OrElse status = "C") Then
                                    objMatrix2.Columns.Item("BrCode").Editable = False

                                    objMatrix2.Columns.Item("Col_1").Editable = False
                                    objMatrix2.Columns.Item("Col_4").Editable = False
                                    objMatrix2.Columns.Item("MUpdated").Editable = False
                                    objMatrix2.Columns.Item("UOM").Editable = False
                                    objMatrix2.Columns.Item("PUpdate").Editable = False
                                    objMatrix2.Columns.Item("PUS").Editable = False
                                    objMatrix2.Columns.Item("PDN").Editable = False
                                    objMatrix2.Columns.Item("PS").Editable = False
                                    objMatrix2.Columns.Item("Col_13").Editable = False

                                    objForm.Items.Item("FOUT").Enabled = True
                                    objForm.Items.Item("FDATEF").Enabled = False
                                    objForm.Items.Item("TDATEF").Enabled = False
                                    objForm.Items.Item("FTIMEF").Enabled = False
                                    objForm.Items.Item("TTIMEF").Enabled = False
                                    objForm.Items.Item("Item_22").Enabled = False

                                ElseIf (status1 = "MOP") And (status = "A" OrElse status = "C" OrElse status = "N") Then

                                    Try
                                        objMatrix3.Columns.Item("BCode").Editable = False

                                        objMatrix3.Columns.Item("NSPrice").Editable = False
                                        objMatrix3.Columns.Item("CName").Editable = False
                                        objMatrix3.Columns.Item("IAttach").Editable = False
                                        objMatrix3.Columns.Item("MUpdated").Editable = False
                                        objMatrix3.Columns.Item("UOM").Editable = False
                                        objMatrix3.Columns.Item("PUpdate").Editable = False
                                        objMatrix3.Columns.Item("PUS").Editable = False
                                        objMatrix3.Columns.Item("PriceR").Editable = False
                                        objMatrix3.Columns.Item("PDN").Editable = False
                                        objMatrix3.Columns.Item("PS").Editable = False

                                        objForm.Items.Item("FDATEM1").Enabled = False
                                        objForm.Items.Item("TDATEMP").Enabled = False
                                        objForm.Items.Item("FTIMEM").Enabled = False
                                        objForm.Items.Item("TTIMEM").Enabled = False

                                        objForm.Items.Item("TNXBRC").Enabled = False
                                        objForm.Items.Item("TNXPUTP").Enabled = False
                                        objForm.Items.Item("TNXUSR").Enabled = False
                                        objForm.Items.Item("PName").Enabled = False
                                    Catch
                                    End Try

                                ElseIf (status1 = "Every day Special Price") And (status = "A" OrElse status = "C" OrElse status = "N") Then
                                    objMatrix6.Columns.Item("BCode").Editable = False
                                    objMatrix6.Columns.Item("Col_3").Editable = False

                                    objMatrix6.Columns.Item("Col_7").Editable = False

                                    objMatrix6.Columns.Item("Col_8").Editable = False
                                    objMatrix6.Columns.Item("Col_11").Editable = False
                                    objMatrix6.Columns.Item("Col_12").Editable = False

                                    objMatrix6.Columns.Item("IAttach").Editable = False
                                    objMatrix6.Columns.Item("Col_13").Editable = False

                                    objForm.Items.Item("TNXBRC").Enabled = False
                                    objForm.Items.Item("TNXPUTP").Enabled = False
                                    objForm.Items.Item("TNXUSR").Enabled = False
                                    objForm.Items.Item("PName").Enabled = False
                                ElseIf (status1 = "Bulk Cost Update" Or status1 = "BULK COST UPDATE") And (status = "A" OrElse status = "C" OrElse status = "N") Then
                                    objMatrix7 = objForm.Items.Item("Mtx7").Specific

                                    objMatrix7.Columns.Item("BCode").Editable = False
                                    objMatrix7.Columns.Item("UOM").Editable = False
                                    objMatrix7.Columns.Item("Base").Editable = False
                                    objMatrix7.Columns.Item("Tax").Editable = False
                                    objMatrix7.Columns.Item("PUpdate").Editable = False
                                    objMatrix7.Columns.Item("PUS").Editable = False
                                    objMatrix7.Columns.Item("NPRC").Editable = False
                                    objForm.Items.Item("FDATEB").Enabled = False
                                    objForm.Items.Item("TDATEB").Enabled = False
                                    objForm.Items.Item("TNXBRC").Enabled = False
                                    objForm.Items.Item("TNXPUTP").Enabled = False
                                    objForm.Items.Item("TNXUSR").Enabled = False
                                    objForm.Items.Item("PName").Enabled = False
                                ElseIf (status1 = "Bulk Price Update" Or status1 = "BULK PRICE UPDATE") And (status = "A" OrElse status = "C" OrElse status = "N") Then
                                    objMatrix9.Columns.Item("Col_0").Editable = False
                                    objMatrix9.Columns.Item("Col_4").Editable = False
                                    objMatrix9.Columns.Item("Col_2").Editable = False
                                    objMatrix9.Columns.Item("Col_3").Editable = False
                                    objMatrix9.Columns.Item("PUpdate").Editable = False
                                    objMatrix9.Columns.Item("PUS").Editable = False

                                    objForm.Items.Item("TNXBRC").Enabled = False
                                    objForm.Items.Item("TNXPUTP").Enabled = False
                                    objForm.Items.Item("TNXUSR").Enabled = False
                                    objForm.Items.Item("PName").Enabled = False
                                ElseIf (status1 = "Cost Fluctuation Stock Transfer") And (status = "A" OrElse status = "C" OrElse status = "N") Then
                                    objMatrix8.Columns.Item("Col_3").Editable = False
                                    objMatrix8.Columns.Item("Col_12").Editable = False
                                    objMatrix8.Columns.Item("Col_13").Editable = False
                                    objMatrix8.Columns.Item("Col_14").Editable = False
                                    objMatrix8.Columns.Item("PUpdate").Editable = False
                                    objMatrix8.Columns.Item("PUS").Editable = False
                                    objForm.Items.Item("TNXBRC").Enabled = False
                                    objForm.Items.Item("TNXPUTP").Enabled = False
                                    objForm.Items.Item("TNXUSR").Enabled = False
                                    objForm.Items.Item("PName").Enabled = False
                                End If


                                If (status1 = "Cost Fluctuation" Or status1 = "COST FLUCTUATION") And (status = "S") Then

                                    objMatrix1.Columns.Item("UOM").Editable = False

                                    Dim oRS1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    Dim sSQL1 As String = "SELECT DISTINCT  T4.""U_UKey"" FROM ""@SBO_APPHDR"" T0 " &
    "INNER JOIN ""@SBO_APPAUT"" T2 ON T0.""Code"" = T2.""Code"" INNER JOIN ""@SBO_AST"" T3 ON T3.""Code"" = T2.""U_M3_1"" " &
    "INNER JOIN ""@SBO_AST_C0"" T4 ON T4.""Code"" = T3.""Code"" WHERE T0.""U_Active"" = 'Y' AND UPPER(T0.""U_PCR"") = '" & status1.Replace("'", "''").ToUpper() & "' " &
    "AND UPPER(T4.""U_UKey"") = '" & loginUser.Replace("'", "''").ToUpper() & "'"
                                    oRS1.DoQuery(sSQL1)
                                    If oRS1.RecordCount > 0 Then
                                        'Connected user is an authorizer
                                        objMatrix1.Columns.Item("NRP").Editable = True
                                    Else
                                        'Connected user is NOT an authorizer
                                        objMatrix1.Columns.Item("NRP").Editable = False
                                    End If
                                    'objMatrix1.Columns.Item("NRP").Editable = False
                                    objMatrix1.Columns.Item("PUpdate").Editable = False
                                    objMatrix1.Columns.Item("PUS").Editable = False
                                    objMatrix1.Columns.Item("Col_0").Editable = False
                                    objMatrix1.Columns.Item("Col_1").Editable = False
                                    objMatrix1.Columns.Item("Col_2").Editable = True
                                    objMatrix1.Columns.Item("Col_3").Editable = True
                                    objMatrix1.Columns.Item("Col_4").Editable = True
                                    objMatrix1.Columns.Item("Col_6").Editable = True
                                    objMatrix1.Columns.Item("Col_7").Editable = True
                                    objForm.Items.Item("TNXBRC").Enabled = False
                                    objForm.Items.Item("TNXPUTP").Enabled = False
                                    objForm.Items.Item("TNXUSR").Enabled = False
                                    objForm.Items.Item("PName").Enabled = False

                                ElseIf (status1 = "Temporary Price Reduction" Or status1 = "TEMPORARY PRICE REDUCTION") And (status = "S") Then
                                    objMatrix4.Columns.Item("BrCode").Editable = False
                                    objMatrix4.Columns.Item("Qty").Editable = False
                                    Dim oRS1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    Dim sSQL1 As String = "SELECT DISTINCT  T4.""U_UKey"" FROM ""@SBO_APPHDR"" T0 " &
    "INNER JOIN ""@SBO_APPAUT"" T2 ON T0.""Code"" = T2.""Code"" INNER JOIN ""@SBO_AST"" T3 ON T3.""Code"" = T2.""U_M3_1"" " &
    "INNER JOIN ""@SBO_AST_C0"" T4 ON T4.""Code"" = T3.""Code"" WHERE T0.""U_Active"" = 'Y' AND UPPER(T0.""U_PCR"") = '" & status1.Replace("'", "''").ToUpper() & "' " &
    "AND UPPER(T4.""U_UKey"") = '" & loginUser.Replace("'", "''").ToUpper() & "'"
                                    oRS1.DoQuery(sSQL1)
                                    If oRS1.RecordCount > 0 Then
                                        'Connected user is an authorizer
                                        objMatrix4.Columns.Item("SPrice").Editable = True
                                    Else
                                        'Connected user is NOT an authorizer
                                        objMatrix4.Columns.Item("SPrice").Editable = False
                                    End If

                                    objMatrix4.Columns.Item("FundA").Editable = False
                                    objMatrix4.Columns.Item("Supp").Editable = False
                                    objMatrix4.Columns.Item("UOM1").Editable = False
                                    objMatrix4.Columns.Item("Remark").Editable = False
                                    objMatrix4.Columns.Item("MUpdated").Editable = False
                                    objMatrix4.Columns.Item("PUpdate").Editable = False
                                    objMatrix4.Columns.Item("PUS").Editable = False
                                    objMatrix4.Columns.Item("PriceR").Editable = False
                                    objMatrix4.Columns.Item("PRS").Editable = False
                                    objMatrix4.Columns.Item("PDN").Editable = False
                                    objMatrix4.Columns.Item("PS").Editable = False

                                    objMatrix4.Columns.Item("Col_0").Editable = True
                                    objForm.Items.Item("TNXBRC").Enabled = False
                                    objForm.Items.Item("TNXPUTP").Enabled = False
                                    objForm.Items.Item("TNXUSR").Enabled = False
                                    objForm.Items.Item("PName").Enabled = False

                                ElseIf (status1 = "Flyer Price Changes" Or status1 = "FLYER PRICE CHANGES") And (status = "S") Then
                                    objMatrix2.Columns.Item("BrCode").Editable = False
                                    Dim oRS1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    Dim sSQL1 As String = "SELECT DISTINCT  T4.""U_UKey"" FROM ""@SBO_APPHDR"" T0 " &
    "INNER JOIN ""@SBO_APPAUT"" T2 ON T0.""Code"" = T2.""Code"" INNER JOIN ""@SBO_AST"" T3 ON T3.""Code"" = T2.""U_M3_1"" " &
    "INNER JOIN ""@SBO_AST_C0"" T4 ON T4.""Code"" = T3.""Code"" WHERE T0.""U_Active"" = 'Y' AND UPPER(T0.""U_PCR"") = '" & status1.Replace("'", "''").ToUpper() & "' " &
    "AND UPPER(T4.""U_UKey"") = '" & loginUser.Replace("'", "''").ToUpper() & "'"
                                    oRS1.DoQuery(sSQL1)
                                    If oRS1.RecordCount > 0 Then
                                        'Connected user is an authorizer
                                        objMatrix2.Columns.Item("Col_1").Editable = True
                                    Else
                                        'Connected user is NOT an authorizer
                                        objMatrix2.Columns.Item("Col_1").Editable = False
                                    End If

                                    objMatrix2.Columns.Item("Col_4").Editable = False
                                    objMatrix2.Columns.Item("MUpdated").Editable = False
                                    objMatrix2.Columns.Item("UOM").Editable = False
                                    objMatrix2.Columns.Item("PUpdate").Editable = False
                                    objMatrix2.Columns.Item("PUS").Editable = False
                                    objMatrix2.Columns.Item("PDN").Editable = False
                                    objMatrix2.Columns.Item("PS").Editable = False
                                    objMatrix2.Columns.Item("Col_13").Editable = False

                                    objForm.Items.Item("FOUT").Enabled = True
                                    objForm.Items.Item("FDATEF").Enabled = False
                                    objForm.Items.Item("TDATEF").Enabled = False
                                    objForm.Items.Item("FTIMEF").Enabled = False
                                    objForm.Items.Item("TTIMEF").Enabled = False
                                    objForm.Items.Item("Item_22").Enabled = False

                                ElseIf (status1 = "MOP") And (status = "S") Then

                                    Try
                                        objMatrix3.Columns.Item("BCode").Editable = False
                                        Dim oRS1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                        Dim sSQL1 As String = "SELECT DISTINCT  T4.""U_UKey"" FROM ""@SBO_APPHDR"" T0 " &
    "INNER JOIN ""@SBO_APPAUT"" T2 ON T0.""Code"" = T2.""Code"" INNER JOIN ""@SBO_AST"" T3 ON T3.""Code"" = T2.""U_M3_1"" " &
    "INNER JOIN ""@SBO_AST_C0"" T4 ON T4.""Code"" = T3.""Code"" WHERE T0.""U_Active"" = 'Y' AND UPPER(T0.""U_PCR"") = '" & status1.Replace("'", "''").ToUpper() & "' " &
    "AND UPPER(T4.""U_UKey"") = '" & loginUser.Replace("'", "''").ToUpper() & "'"
                                        oRS1.DoQuery(sSQL1)
                                        If oRS1.RecordCount > 0 Then
                                            'Connected user is an authorizer
                                            objMatrix3.Columns.Item("NSPrice").Editable = True
                                        Else
                                            'Connected user is NOT an authorizer
                                            objMatrix3.Columns.Item("NSPrice").Editable = False
                                        End If

                                        objMatrix3.Columns.Item("CName").Editable = False
                                        objMatrix3.Columns.Item("IAttach").Editable = False
                                        objMatrix3.Columns.Item("MUpdated").Editable = False
                                        objMatrix3.Columns.Item("UOM").Editable = False
                                        objMatrix3.Columns.Item("PUpdate").Editable = False
                                        objMatrix3.Columns.Item("PUS").Editable = False
                                        objMatrix3.Columns.Item("PriceR").Editable = False
                                        objMatrix3.Columns.Item("PDN").Editable = False
                                        objMatrix3.Columns.Item("PS").Editable = False

                                        objForm.Items.Item("FDATEM1").Enabled = False
                                        objForm.Items.Item("TDATEMP").Enabled = False
                                        objForm.Items.Item("FTIMEM").Enabled = False
                                        objForm.Items.Item("TTIMEM").Enabled = False

                                        objForm.Items.Item("TNXBRC").Enabled = False
                                        objForm.Items.Item("TNXPUTP").Enabled = False
                                        objForm.Items.Item("TNXUSR").Enabled = False
                                        objForm.Items.Item("PName").Enabled = False
                                    Catch
                                    End Try

                                ElseIf (status1 = "Every day Special Price") And (status = "S") Then
                                    objMatrix6.Columns.Item("BCode").Editable = False
                                    objMatrix6.Columns.Item("Col_3").Editable = False

                                    objMatrix6.Columns.Item("Col_7").Editable = False

                                    Dim oRS1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    Dim sSQL1 As String = "SELECT DISTINCT  T4.""U_UKey"" FROM ""@SBO_APPHDR"" T0 " &
    "INNER JOIN ""@SBO_APPAUT"" T2 ON T0.""Code"" = T2.""Code"" INNER JOIN ""@SBO_AST"" T3 ON T3.""Code"" = T2.""U_M3_1"" " &
    "INNER JOIN ""@SBO_AST_C0"" T4 ON T4.""Code"" = T3.""Code"" WHERE T0.""U_Active"" = 'Y' AND UPPER(T0.""U_PCR"") = '" & status1.Replace("'", "''").ToUpper() & "' " &
    "AND UPPER(T4.""U_UKey"") = '" & loginUser.Replace("'", "''").ToUpper() & "'"
                                    oRS1.DoQuery(sSQL1)
                                    If oRS1.RecordCount > 0 Then
                                        'Connected user is an authorizer
                                        objMatrix6.Columns.Item("Col_8").Editable = True
                                    Else
                                        'Connected user is NOT an authorizer
                                        objMatrix6.Columns.Item("Col_8").Editable = False
                                    End If

                                    objMatrix6.Columns.Item("Col_11").Editable = False
                                    objMatrix6.Columns.Item("Col_12").Editable = False

                                    objMatrix6.Columns.Item("IAttach").Editable = False
                                    objMatrix6.Columns.Item("Col_13").Editable = False

                                    objForm.Items.Item("TNXBRC").Enabled = False
                                    objForm.Items.Item("TNXPUTP").Enabled = False
                                    objForm.Items.Item("TNXUSR").Enabled = False
                                    objForm.Items.Item("PName").Enabled = False
                                ElseIf (status1 = "Bulk Cost Update" Or status1 = "BULK COST UPDATE") And (status = "S") Then
                                    objMatrix7 = objForm.Items.Item("Mtx7").Specific
                                    objMatrix7.Columns.Item("PUpdate").Editable = False
                                    objMatrix7.Columns.Item("PUS").Editable = False
                                    objMatrix7.Columns.Item("BCode").Editable = False
                                    objMatrix7.Columns.Item("UOM").Editable = False
                                    objMatrix7.Columns.Item("Base").Editable = False
                                    objMatrix7.Columns.Item("Tax").Editable = False
                                    Dim oRS1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    Dim sSQL1 As String = "SELECT DISTINCT  T4.""U_UKey"" FROM ""@SBO_APPHDR"" T0 " &
    "INNER JOIN ""@SBO_APPAUT"" T2 ON T0.""Code"" = T2.""Code"" INNER JOIN ""@SBO_AST"" T3 ON T3.""Code"" = T2.""U_M3_1"" " &
    "INNER JOIN ""@SBO_AST_C0"" T4 ON T4.""Code"" = T3.""Code"" WHERE T0.""U_Active"" = 'Y' AND UPPER(T0.""U_PCR"") = '" & status1.Replace("'", "''").ToUpper() & "' " &
    "AND UPPER(T4.""U_UKey"") = '" & loginUser.Replace("'", "''").ToUpper() & "'"
                                    oRS1.DoQuery(sSQL1)
                                    If oRS1.RecordCount > 0 Then
                                        'Connected user is an authorizer
                                        objMatrix7.Columns.Item("NPRC").Editable = True
                                    Else
                                        'Connected user is NOT an authorizer
                                        objMatrix7.Columns.Item("NPRC").Editable = False
                                    End If


                                    objForm.Items.Item("FDATEB").Enabled = False
                                    objForm.Items.Item("TDATEB").Enabled = False
                                    objForm.Items.Item("TNXBRC").Enabled = False
                                    objForm.Items.Item("TNXPUTP").Enabled = False
                                    objForm.Items.Item("TNXUSR").Enabled = False
                                    objForm.Items.Item("PName").Enabled = False

                                ElseIf (status1 = "Bulk Price Update" Or status1 = "BULK PRICE UPDATE") And (status = "S") Then
                                    objMatrix9.Columns.Item("Col_0").Editable = False
                                    objMatrix9.Columns.Item("Col_4").Editable = False
                                    objMatrix9.Columns.Item("Col_2").Editable = False
                                    objMatrix9.Columns.Item("PUpdate").Editable = False
                                    objMatrix9.Columns.Item("PUS").Editable = False
                                    Dim oRS1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    Dim sSQL1 As String = "SELECT DISTINCT  T4.""U_UKey"" FROM ""@SBO_APPHDR"" T0 " &
    "INNER JOIN ""@SBO_APPAUT"" T2 ON T0.""Code"" = T2.""Code"" INNER JOIN ""@SBO_AST"" T3 ON T3.""Code"" = T2.""U_M3_1"" " &
    "INNER JOIN ""@SBO_AST_C0"" T4 ON T4.""Code"" = T3.""Code"" WHERE T0.""U_Active"" = 'Y' AND UPPER(T0.""U_PCR"") = '" & status1.Replace("'", "''").ToUpper() & "' " &
    "AND UPPER(T4.""U_UKey"") = '" & loginUser.Replace("'", "''").ToUpper() & "'"
                                    oRS1.DoQuery(sSQL1)
                                    If oRS1.RecordCount > 0 Then
                                        'Connected user is an authorizer
                                        objMatrix9.Columns.Item("Col_3").Editable = True
                                    Else
                                        'Connected user is NOT an authorizer
                                        objMatrix9.Columns.Item("Col_3").Editable = False
                                    End If



                                    objForm.Items.Item("TNXBRC").Enabled = False
                                    objForm.Items.Item("TNXPUTP").Enabled = False
                                    objForm.Items.Item("TNXUSR").Enabled = False
                                    objForm.Items.Item("PName").Enabled = False
                                ElseIf (status1 = "Cost Fluctuation Stock Transfer") And (status = "S") Then
                                    objMatrix8.Columns.Item("Col_3").Editable = False
                                    objMatrix8.Columns.Item("Col_12").Editable = False
                                    objMatrix8.Columns.Item("Col_13").Editable = False
                                    objMatrix8.Columns.Item("Col_14").Editable = False
                                    objMatrix8.Columns.Item("PUpdate").Editable = False
                                    objMatrix8.Columns.Item("PUS").Editable = False
                                    objForm.Items.Item("TNXBRC").Enabled = False
                                    objForm.Items.Item("TNXPUTP").Enabled = False
                                    objForm.Items.Item("TNXUSR").Enabled = False
                                    objForm.Items.Item("PName").Enabled = False
                                End If

                                ' Enable OK and Cancel buttons
                                Try : objForm.Items.Item("1").Enabled = True : Catch : End Try   ' OK
                                Try : objForm.Items.Item("2").Enabled = True : Catch : End Try   ' Cancel
                                Try : objForm.Items.Item("Item_28").Enabled = True : Catch : End Try
                                Try : objForm.Items.Item("Item_8").Enabled = True : Catch : End Try
                                Try : objForm.Items.Item("Item_23").Enabled = True : Catch : End Try
                                Try : objForm.Items.Item("Item_13").Enabled = True : Catch : End Try

                                ' Enable Find menu
                                Try : objForm.EnableMenu("1281", True) : Catch : End Try
                                'Enable Add menu
                                Try : objForm.EnableMenu("1282", True) : Catch : End Try


                            Else

                                'Enable all items first
                                For i As Integer = 0 To objForm.Items.Count - 1
                                    Try
                                        objForm.Items.Item(i).Enabled = True
                                    Catch
                                    End Try
                                Next

                                'Disable header fields
                                Try : objForm.Items.Item("PriceLNum").Enabled = False : Catch : End Try

                                Try : objForm.Items.Item("TNXDDT").Enabled = False : Catch : End Try
                                Try : objForm.Items.Item("DocNum").Enabled = False : Catch : End Try
                                Try : objForm.Items.Item("TNXSTUS").Enabled = False : Catch : End Try
                                Try : objForm.Items.Item("AppID").Enabled = False : Catch : End Try


                                ' Enable Add / Find menu
                                Try : objForm.EnableMenu("1281", True) : Catch : End Try
                                Try : objForm.Items.Item("Item_28").Enabled = True : Catch : End Try
                                'Try : objForm.Items.Item("Item_8").Enabled = True : Catch : End Try
                                'Try : objForm.Items.Item("Item_23").Enabled = True : Catch : End Try
                                ' Enable Submit button
                                If oDBs_Head.GetValue("U_TNXPUTP", 0).Trim() = "Flyer Price Changes" And oDBs_Head.GetValue("U_FOUT", 0).Trim() = "No" Then

                                    ''Try : objForm.Items.Item("Submit").Enabled = False : Catch : End Try
                                    Try : objForm.Items.Item("Item_28").Enabled = True : Catch : End Try
                                Else
                                    ''Try : objForm.Items.Item("Submit").Enabled = True : Catch : End Try
                                    Try : objForm.Items.Item("Item_28").Enabled = True : Catch : End Try
                                End If

                                If oDBs_Head.GetValue("U_TNXPUTP", 0).Trim() = "MOP" And oDBs_Head.GetValue("U_TNXSTUS", 0).Trim() = "C" Then

                                    If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                    End If
                                Else
                                    Try
                                        If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        End If
                                    Catch
                                    End Try

                                End If
                                ' Switch to OK MODE
                                If status = "OPEN" OrElse status = "O" Then

                                    Try : objMatrix2 = objForm.Items.Item("Mtx2").Specific : objMatrix2.Columns.Item("BrCode").Editable = True : Catch : End Try
                                    Try : objMatrix2 = objForm.Items.Item("Mtx2").Specific : objMatrix2.Columns.Item("Col_1").Editable = True : Catch : End Try
                                    Try : objMatrix2 = objForm.Items.Item("Mtx2").Specific : objMatrix2.Columns.Item("Col_4").Editable = True : Catch : End Try
                                    Try : objMatrix2 = objForm.Items.Item("Mtx2").Specific : objMatrix2.Columns.Item("MUpdated").Editable = True : Catch : End Try
                                    Try : objMatrix2 = objForm.Items.Item("Mtx2").Specific : objMatrix2.Columns.Item("UOM").Editable = True : Catch : End Try
                                    Try : objMatrix2 = objForm.Items.Item("Mtx2").Specific : objMatrix2.Columns.Item("PUS").Editable = True : Catch : End Try
                                    Try : objMatrix2 = objForm.Items.Item("Mtx2").Specific : objMatrix2.Columns.Item("PDN").Editable = True : Catch : End Try
                                    Try : objMatrix2 = objForm.Items.Item("Mtx2").Specific : objMatrix2.Columns.Item("PS").Editable = True : Catch : End Try
                                    Try : objMatrix2 = objForm.Items.Item("Mtx2").Specific : objMatrix2.Columns.Item("Col_14").Editable = True : Catch : End Try
                                    Try : objMatrix2 = objForm.Items.Item("Mtx2").Specific : objMatrix2.Columns.Item("Col_13").Editable = True : Catch : End Try

                                    Try : objMatrix3 = objForm.Items.Item("Mtx3").Specific : objMatrix3.Columns.Item("BCode").Editable = True : Catch : End Try
                                    Try : objMatrix3 = objForm.Items.Item("Mtx3").Specific : objMatrix3.Columns.Item("NSPrice").Editable = True : Catch : End Try
                                    Try : objMatrix3 = objForm.Items.Item("Mtx3").Specific : objMatrix3.Columns.Item("CName").Editable = True : Catch : End Try
                                    Try : objMatrix3 = objForm.Items.Item("Mtx3").Specific : objMatrix3.Columns.Item("IAttach").Editable = True : Catch : End Try
                                    Try : objMatrix3 = objForm.Items.Item("Mtx3").Specific : objMatrix3.Columns.Item("MUpdated").Editable = True : Catch : End Try
                                    Try : objMatrix3 = objForm.Items.Item("Mtx3").Specific : objMatrix3.Columns.Item("UOM").Editable = True : Catch : End Try
                                    Try : objMatrix3 = objForm.Items.Item("Mtx3").Specific : objMatrix3.Columns.Item("PUpdate").Editable = True : Catch : End Try
                                    Try : objMatrix3 = objForm.Items.Item("Mtx3").Specific : objMatrix3.Columns.Item("PUS").Editable = True : Catch : End Try
                                    Try : objMatrix3 = objForm.Items.Item("Mtx3").Specific : objMatrix3.Columns.Item("PriceR").Editable = True : Catch : End Try
                                    Try : objMatrix3 = objForm.Items.Item("Mtx3").Specific : objMatrix3.Columns.Item("PDN").Editable = True : Catch : End Try
                                    Try : objMatrix3 = objForm.Items.Item("Mtx3").Specific : objMatrix3.Columns.Item("PS").Editable = True : Catch : End Try

                                    Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("BrCode").Editable = True : Catch : End Try
                                    Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("Qty").Editable = True : Catch : End Try
                                    Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("Sprice").Editable = True : Catch : End Try
                                    Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("FundA").Editable = True : Catch : End Try
                                    Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("Supp").Editable = True : Catch : End Try
                                    Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("UOM1").Editable = True : Catch : End Try
                                    Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("Remark").Editable = True : Catch : End Try
                                    Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("MUpdated").Editable = True : Catch : End Try
                                    Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("PUpdate").Editable = True : Catch : End Try
                                    Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("PUS").Editable = True : Catch : End Try
                                    Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("PriceR").Editable = True : Catch : End Try
                                    Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("PRS").Editable = True : Catch : End Try
                                    Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("PDN").Editable = True : Catch : End Try
                                    Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("PS").Editable = True : Catch : End Try
                                    Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("Sell").Editable = True : Catch : End Try
                                    Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("Col_0").Editable = True : Catch : End Try


                                    Try : objMatrix6 = objForm.Items.Item("Mtx6").Specific : objMatrix6.Columns.Item("Col_0").Editable = True : Catch : End Try
                                    Try : objMatrix6 = objForm.Items.Item("Mtx6").Specific : objMatrix6.Columns.Item("Col_3").Editable = True : Catch : End Try
                                    Try : objMatrix6 = objForm.Items.Item("Mtx6").Specific : objMatrix6.Columns.Item("Col_7").Editable = True : Catch : End Try
                                    Try : objMatrix6 = objForm.Items.Item("Mtx6").Specific : objMatrix6.Columns.Item("Col_8").Editable = True : Catch : End Try
                                    Try : objMatrix6 = objForm.Items.Item("Mtx6").Specific : objMatrix6.Columns.Item("Col_11").Editable = True : Catch : End Try
                                    Try : objMatrix6 = objForm.Items.Item("Mtx6").Specific : objMatrix6.Columns.Item("Col_12").Editable = True : Catch : End Try
                                    Try : objMatrix6 = objForm.Items.Item("Mtx6").Specific : objMatrix6.Columns.Item("IAttach").Editable = True : Catch : End Try
                                    Try : objMatrix6 = objForm.Items.Item("Mtx6").Specific : objMatrix6.Columns.Item("Col_13").Editable = True : Catch : End Try
                                    Try : objMatrix6 = objForm.Items.Item("Mtx6").Specific : objMatrix6.Columns.Item("PUpdate").Editable = True : Catch : End Try
                                    Try : objMatrix6 = objForm.Items.Item("Mtx6").Specific : objMatrix6.Columns.Item("PUS").Editable = True : Catch : End Try

                                    Try : objMatrix7 = objForm.Items.Item("Mtx7").Specific : objMatrix7.Columns.Item("BCode").Editable = True : Catch : End Try
                                    Try : objMatrix7 = objForm.Items.Item("Mtx7").Specific : objMatrix7.Columns.Item("UOM").Editable = True : Catch : End Try
                                    Try : objMatrix7 = objForm.Items.Item("Mtx7").Specific : objMatrix7.Columns.Item("Base").Editable = True : Catch : End Try
                                    Try : objMatrix7 = objForm.Items.Item("Mtx7").Specific : objMatrix7.Columns.Item("Tax").Editable = True : Catch : End Try
                                    Try : objMatrix7 = objForm.Items.Item("Mtx7").Specific : objMatrix7.Columns.Item("NPRC").Editable = True : Catch : End Try
                                    Try : objMatrix7 = objForm.Items.Item("Mtx7").Specific : objMatrix7.Columns.Item("PUpdate").Editable = True : Catch : End Try
                                    Try : objMatrix7 = objForm.Items.Item("Mtx7").Specific : objMatrix7.Columns.Item("PUS").Editable = True : Catch : End Try

                                End If


                            End If


                            Try
                                Dim updateType As String = ""
                                Try : updateType = oDBs_Head.GetValue("U_TNXPUTP", 0).Trim() : Catch : End Try
                                If String.IsNullOrWhiteSpace(updateType) Then
                                    Try
                                        Dim cb As SAPbouiCOM.ComboBox =
                                        CType(objForm.Items.Item("TNXPUTP").Specific, SAPbouiCOM.ComboBox)
                                        If cb IsNot Nothing AndAlso cb.Selected IsNot Nothing Then
                                            updateType = If(cb.Selected.Value, "").ToString().Trim()
                                        End If
                                    Catch
                                    End Try
                                End If
                                If Not String.IsNullOrWhiteSpace(updateType) Then
                                    SetTabsByUpdateType(updateType)
                                Else
                                    SetTabsByUpdateType("Cost Fluctuation")
                                End If

                                ' Show flyer fields only for Flyer Price Changes
                                Dim showFlyer As Boolean = (updateType = "Flyer Price Changes")
                                objForm.Items.Item("Item_57").Visible = showFlyer
                                objForm.Items.Item("Flyer").Visible = showFlyer
                                objForm.Items.Item("Item_10").Visible = showFlyer
                                objForm.Items.Item("Item_13").Visible = showFlyer
                                If Not showFlyer Then
                                    Try : oDBs_Head.SetValue("U_Flyer", oDBs_Head.Offset, "") : Catch : End Try
                                End If


                                ' Show TPR fields only for TPR Price Changes
                                Dim showTPR As Boolean = (updateType = "Temporary Price Reduction")

                                objForm.Items.Item("Item_10").Visible = showTPR
                                objForm.Items.Item("Item_13").Visible = showTPR
                                If Not showTPR Then
                                    Try : oDBs_Head.SetValue("U_TPRT", oDBs_Head.Offset, "") : Catch : End Try
                                End If



                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(
                                "SetTabs error: " & ex.Message,
                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Try

                            ' Scheduled fields visibility
                            'Try
                            '    Dim currentType As String = ""
                            '    Try : currentType = oDBs_Head.GetValue("U_TNXTYP", 0).Trim() : Catch : End Try
                            '    Dim showScheduled As Boolean = (currentType = "Scheduled Update")
                            '    Try : objForm.Items.Item("Item_38").Visible = showScheduled : Catch : End Try
                            '    Try : objForm.Items.Item("FDate").Visible = showScheduled : Catch : End Try
                            '    Try : objForm.Items.Item("Item_39").Visible = showScheduled : Catch : End Try
                            '    Try : objForm.Items.Item("FTime").Visible = showScheduled : Catch : End Try
                            '    If Not showScheduled Then
                            '        Try : oDBs_Head.SetValue("U_FDate", oDBs_Head.Offset, "") : Catch : End Try
                            '        Try : oDBs_Head.SetValue("U_FTime", oDBs_Head.Offset, "") : Catch : End Try
                            '    End If
                            'Catch
                            'End Try

                        Finally
                            objForm.Freeze(False)
                        End Try
                    End If

    '                If BusinessObjectInfo.BeforeAction = False AndAlso BusinessObjectInfo.ActionSuccess = True Then


    '                    objForm.Freeze(True)

    '                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
    '                    ''SetMatrixTotals(objForm)
    '                    Dim sType As String = oDBs_Head.GetValue("U_TNXPUTP", 0).Trim()
    '                    objForm.EnableMenu("5890", True)
    '                    objForm.Items.Item("Item_28").Enabled = True
    '                    'Enable button only for Flyer
    '                    objForm.Items.Item("Item_8").Enabled = (sType = "Flyer Price Changes")

    '                    Select Case sType

    '                        Case "Cost Fluctuation"
    '                            objForm.PaneLevel = 1
    '                            objForm.EnableMenu("5890", True)
    '                            objForm.Settings.MatrixUID = "Mtx1"
    '                            objForm.Items.Item("Item_8").Visible = False

    '                        Case "Flyer Price Changes"
    '                            objForm.PaneLevel = 2
    '                            objForm.EnableMenu("5890", True)
    '                            objForm.Settings.MatrixUID = "Mtx2"
    '                            objForm.Items.Item("Item_8").Visible = True

    '                            Try
    '                                objMatrix2 = objForm.Items.Item("Mtx2").Specific

    '                                objMatrix2.Columns.Item("Cost").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
    '                                objMatrix2.Columns.Item("Col_0").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
    '                                objMatrix2.Columns.Item("Col_1").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
    '                                objMatrix2.Columns.Item("Col_1").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
    '                                objMatrix2.Columns.Item("Col_2").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
    '                                objMatrix2.Columns.Item("Col_3").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
    '                                objMatrix2.Columns.Item("Col_4").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
    '                                objMatrix2.Columns.Item("Col_5").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
    '                                objMatrix2.Columns.Item("Col_6").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
    '                                objMatrix2.Columns.Item("Col_7").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
    '                                objMatrix2.Columns.Item("Col_8").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
    '                                objMatrix2.Columns.Item("Col_9").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
    '                                objMatrix2.Columns.Item("Col_10").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
    '                                objMatrix2.Columns.Item("Col_14").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

    '                            Catch ex As Exception
    '                                objMain.objApplication.StatusBar.SetText(ex.Message)
    '                            End Try

    '                        Case "MOP"
    '                            objForm.PaneLevel = 3
    '                            objForm.EnableMenu("5890", True)
    '                            objForm.Settings.MatrixUID = "Mtx3"
    '                            objForm.Items.Item("Item_8").Visible = False

    '                        Case "Temporary Price Reduction"
    '                            objForm.PaneLevel = 4
    '                            objForm.EnableMenu("5890", True)
    '                            objForm.Settings.MatrixUID = "Mtx4"
    '                            objForm.Items.Item("Item_8").Visible = False
    '                        Case "Every day Special Price"
    '                            objForm.PaneLevel = 5
    '                            objForm.EnableMenu("5890", True)
    '                            objForm.Settings.MatrixUID = "Mtx5"
    '                            objForm.Items.Item("Item_8").Visible = False
    '                        Case "Attachment Tab"
    '                            objForm.PaneLevel = 6
    '                            objForm.EnableMenu("5890", True)
    '                            objForm.Settings.MatrixUID = "Mtx6"
    '                            objForm.Items.Item("Item_8").Visible = False

    '                    End Select

    '                    objForm.Freeze(False)



    '                End If

    '                If BusinessObjectInfo.BeforeAction = False AndAlso BusinessObjectInfo.ActionSuccess = True Then
    '                    objForm.Freeze(True)
    '                    ' Enable Form Settings
    '                    Try : objForm.EnableMenu("1282", True) : Catch : End Try
    '                    Try : objForm.EnableMenu("5890", True) : Catch : End Try

    '                    Try
    '                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
    '                        objForm.Items.Item("Item_28").Enabled = True

    '                        Try
    '                            'Dim oChk As SAPbouiCOM.CheckBox = CType(objForm.Items.Item("Check").Specific, SAPbouiCOM.CheckBox)
    '                            'Dim strCheck As String = oDBs_Head.GetValue("U_Check", 0).Trim()
    '                            ''If oChk.Checked = True Then
    '                            'If strCheck = "Y" Then
    '                            '    objForm.Items.Item("FDTT").Enabled = True
    '                            '    objForm.Items.Item("TDTT").Enabled = True
    '                            'Else

    '                            '    objForm.Items.Item("FDTT").Enabled = False
    '                            '    objForm.Items.Item("TDTT").Enabled = False
    '                            'End If
    '                        Catch ex As Exception

    '                        End Try

    '                        Dim status As String = ""
    '                        Try
    '                            status = oDBs_Head.GetValue("U_TNXSTUS", 0).Trim().ToUpper()
    '                        Catch
    '                        End Try
    '                        status = status.Replace(" ", "")

    '                        If status = "REJECTED" Or status = "R" Then
    '                            Try : objForm.Items.Item("Submit").Enabled = True : Catch : End Try
    '                        End If
    '                        'If status = "APPROVED" OrElse status = "A" OrElse status = "Submitted" OrElse status = "S" Then


    '                        '    Try : objForm.Items.Item("Submit").Enabled = False : Catch : End Try


    '                        '    Try
    '                        '        If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
    '                        '            objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
    '                        '        End If
    '                        '    Catch
    '                        '    End Try

    '                        '    ' Lock all fields
    '                        '    For i As Integer = 0 To objForm.Items.Count - 1
    '                        '        Try : objForm.Items.Item(i).Enabled = False : Catch : End Try
    '                        '    Next

    '                        '    Try : objForm.EnableMenu("1281", False) : Catch : End Try

    '                        'Else

    '                        '    Try : objForm.EnableMenu("1281", True) : Catch : End Try
    '                        '    Try : objForm.Items.Item("Submit").Enabled = True : Catch : End Try

    '                        '    ' Keep form in OK_MODE for editing
    '                        '    Try
    '                        '        If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
    '                        '            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
    '                        '        End If
    '                        '    Catch
    '                        '    End Try
    '                        'End If

    '                        If status = "APPROVED" OrElse status = "A" OrElse status = "Submitted" OrElse status = "S" OrElse status = "N" OrElse status = "Closed" OrElse status = "C" OrElse status = "CLOSED" OrElse status = "C" Then
    '                            objMatrix1 = objForm.Items.Item("Mtx1").Specific
    '                            ' Disable Submit button
    '                            Try : objForm.Items.Item("Submit").Enabled = False : Catch : End Try
    '                            objForm.Items.Item("Item_28").Enabled = True
    '                            ' Keep form in VIEW MODE

    '                            If oDBs_Head.GetValue("U_TNXPUTP", 0).Trim() = "Cost Fluctuation" Then
    '                                Try
    '                                    objMatrix1.Columns.Item("Col_6").Editable = True
    '                                Catch
    '                                End Try
    '                            ElseIf oDBs_Head.GetValue("U_TNXPUTP", 0).Trim() = "Temporary Price Reduction" Then
    '                                Try
    '                                    objMatrix1.Columns.Item("Col_0").Editable = True
    '                                Catch
    '                                End Try
    '                            ElseIf oDBs_Head.GetValue("U_TNXPUTP", 0).Trim() = "Flyer Price Changes" Then
    '                                Try
    '                                    objForm.Items.Item("FOUT").Enabled = True
    '                                Catch
    '                                End Try
    '                            Else

    '                                Try
    '                                    If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
    '                                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
    '                                    End If
    '                                Catch
    '                                End Try

    '                            End If


    '                            ' Disable all items
    '                            For i As Integer = 0 To objForm.Items.Count - 1

    '                                Dim status1 As String = ""
    '                                Dim loginUser As String = objMain.objApplication.Company.UserName.Trim().ToUpper()
    '                                status1 = oDBs_Head.GetValue("U_TNXPUTP", 0).Trim().ToUpper()

    '                                If (status1 = "Cost Fluctuation" Or status1 = "COST FLUCTUATION") And (status = "A" OrElse status = "C" OrElse status = "N") Then

    '                                    objMatrix1.Columns.Item("UOM").Editable = False

    '                                    objMatrix1.Columns.Item("NRP").Editable = False
    '                                    objMatrix1.Columns.Item("PUpdate").Editable = False
    '                                    objMatrix1.Columns.Item("PUS").Editable = False
    '                                    objMatrix1.Columns.Item("Col_0").Editable = False
    '                                    objMatrix1.Columns.Item("Col_1").Editable = False
    '                                    objMatrix1.Columns.Item("Col_2").Editable = True
    '                                    objMatrix1.Columns.Item("Col_3").Editable = True
    '                                    objMatrix1.Columns.Item("Col_4").Editable = True
    '                                    objMatrix1.Columns.Item("Col_6").Editable = True
    '                                    objMatrix1.Columns.Item("Col_7").Editable = True
    '                                    objForm.Items.Item("TNXBRC").Enabled = False
    '                                    objForm.Items.Item("TNXPUTP").Enabled = False
    '                                    objForm.Items.Item("TNXUSR").Enabled = False
    '                                    objForm.Items.Item("PName").Enabled = False

    '                                ElseIf (status1 = "Temporary Price Reduction" Or status1 = "TEMPORARY PRICE REDUCTION") And (status = "A" OrElse status = "C" OrElse status = "N") Then
    '                                    objMatrix4.Columns.Item("BrCode").Editable = False
    '                                    objMatrix4.Columns.Item("Qty").Editable = False

    '                                    objMatrix4.Columns.Item("SPrice").Editable = False
    '                                    objMatrix4.Columns.Item("FundA").Editable = False
    '                                    objMatrix4.Columns.Item("Supp").Editable = False
    '                                    objMatrix4.Columns.Item("UOM1").Editable = False
    '                                    objMatrix4.Columns.Item("Remark").Editable = False
    '                                    objMatrix4.Columns.Item("MUpdated").Editable = False
    '                                    objMatrix4.Columns.Item("PUpdate").Editable = False
    '                                    objMatrix4.Columns.Item("PUS").Editable = False
    '                                    objMatrix4.Columns.Item("PriceR").Editable = False
    '                                    objMatrix4.Columns.Item("PRS").Editable = False
    '                                    objMatrix4.Columns.Item("PDN").Editable = False
    '                                    objMatrix4.Columns.Item("PS").Editable = False

    '                                    objMatrix4.Columns.Item("Col_0").Editable = True
    '                                    objForm.Items.Item("TNXBRC").Enabled = False
    '                                    objForm.Items.Item("TNXPUTP").Enabled = False
    '                                    objForm.Items.Item("TNXUSR").Enabled = False
    '                                    objForm.Items.Item("PName").Enabled = False

    '                                ElseIf (status1 = "Flyer Price Changes" Or status1 = "FLYER PRICE CHANGES") And (status = "A" OrElse status = "N" OrElse status = "C") Then
    '                                    objMatrix2.Columns.Item("BrCode").Editable = False

    '                                    objMatrix2.Columns.Item("Col_1").Editable = False
    '                                    objMatrix2.Columns.Item("Col_4").Editable = False
    '                                    objMatrix2.Columns.Item("MUpdated").Editable = False
    '                                    objMatrix2.Columns.Item("UOM").Editable = False
    '                                    objMatrix2.Columns.Item("PUpdate").Editable = False
    '                                    objMatrix2.Columns.Item("PUS").Editable = False
    '                                    objMatrix2.Columns.Item("PDN").Editable = False
    '                                    objMatrix2.Columns.Item("PS").Editable = False
    '                                    objMatrix2.Columns.Item("Col_13").Editable = False

    '                                    objForm.Items.Item("FOUT").Enabled = True
    '                                    objForm.Items.Item("FDATEF").Enabled = False
    '                                    objForm.Items.Item("TDATEF").Enabled = False
    '                                    objForm.Items.Item("FTIMEF").Enabled = False
    '                                    objForm.Items.Item("TTIMEF").Enabled = False
    '                                    objForm.Items.Item("Item_22").Enabled = False

    '                                ElseIf (status1 = "MOP") And (status = "A" OrElse status = "C" OrElse status = "N") Then

    '                                    Try
    '                                        objMatrix3.Columns.Item("BCode").Editable = False

    '                                        objMatrix3.Columns.Item("NSPrice").Editable = False
    '                                        objMatrix3.Columns.Item("CName").Editable = False
    '                                        objMatrix3.Columns.Item("IAttach").Editable = False
    '                                        objMatrix3.Columns.Item("MUpdated").Editable = False
    '                                        objMatrix3.Columns.Item("UOM").Editable = False
    '                                        objMatrix3.Columns.Item("PUpdate").Editable = False
    '                                        objMatrix3.Columns.Item("PUS").Editable = False
    '                                        objMatrix3.Columns.Item("PriceR").Editable = False
    '                                        objMatrix3.Columns.Item("PDN").Editable = False
    '                                        objMatrix3.Columns.Item("PS").Editable = False

    '                                        objForm.Items.Item("FDATEM1").Enabled = False
    '                                        objForm.Items.Item("TDATEMP").Enabled = False
    '                                        objForm.Items.Item("FTIMEM").Enabled = False
    '                                        objForm.Items.Item("TTIMEM").Enabled = False

    '                                        objForm.Items.Item("TNXBRC").Enabled = False
    '                                        objForm.Items.Item("TNXPUTP").Enabled = False
    '                                        objForm.Items.Item("TNXUSR").Enabled = False
    '                                        objForm.Items.Item("PName").Enabled = False
    '                                    Catch
    '                                    End Try

    '                                ElseIf (status1 = "Every day Special Price") And (status = "A" OrElse status = "C" OrElse status = "N") Then
    '                                    objMatrix6.Columns.Item("BCode").Editable = False
    '                                    objMatrix6.Columns.Item("Col_3").Editable = False

    '                                    objMatrix6.Columns.Item("Col_7").Editable = False

    '                                    objMatrix6.Columns.Item("Col_8").Editable = False
    '                                    objMatrix6.Columns.Item("Col_11").Editable = False
    '                                    objMatrix6.Columns.Item("Col_12").Editable = False

    '                                    objMatrix6.Columns.Item("IAttach").Editable = False
    '                                    objMatrix6.Columns.Item("Col_13").Editable = False

    '                                    objForm.Items.Item("TNXBRC").Enabled = False
    '                                    objForm.Items.Item("TNXPUTP").Enabled = False
    '                                    objForm.Items.Item("TNXUSR").Enabled = False
    '                                    objForm.Items.Item("PName").Enabled = False
    '                                ElseIf (status1 = "Bulk Cost Update") And (status = "A" OrElse status = "C" OrElse status = "N") Then
    '                                    objMatrix6.Columns.Item("BCode").Editable = False
    '                                    objMatrix6.Columns.Item("UOM").Editable = False
    '                                    objMatrix6.Columns.Item("Base").Editable = False
    '                                    objMatrix6.Columns.Item("Tax").Editable = False

    '                                    objMatrix6.Columns.Item("NPRC").Editable = False
    '                                    objForm.Items.Item("FDATEB").Enabled = False
    '                                    objForm.Items.Item("TDATEB").Enabled = False
    '                                    objForm.Items.Item("TNXBRC").Enabled = False
    '                                    objForm.Items.Item("TNXPUTP").Enabled = False
    '                                    objForm.Items.Item("TNXUSR").Enabled = False
    '                                    objForm.Items.Item("PName").Enabled = False
    '                                ElseIf (status1 = "Cost Fluctuation Stock Transfer") And (status = "A" OrElse status = "C" OrElse status = "N") Then
    '                                    objMatrix6.Columns.Item("Col_3").Editable = False
    '                                    objMatrix6.Columns.Item("Col_12").Editable = False
    '                                    objMatrix6.Columns.Item("Col_13").Editable = False
    '                                    objMatrix6.Columns.Item("Col_14").Editable = False

    '                                    objForm.Items.Item("TNXBRC").Enabled = False
    '                                    objForm.Items.Item("TNXPUTP").Enabled = False
    '                                    objForm.Items.Item("TNXUSR").Enabled = False
    '                                    objForm.Items.Item("PName").Enabled = False
    '                                End If


    '                                If (status1 = "Cost Fluctuation" Or status1 = "COST FLUCTUATION") And (status = "S") Then

    '                                    objMatrix1.Columns.Item("UOM").Editable = False

    '                                    Dim oRS1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '                                    Dim sSQL1 As String = "SELECT DISTINCT  T4.""U_AUTH"" FROM ""@SBO_APPHDR"" T0 " &
    '"INNER JOIN ""@SBO_APPAUT"" T2 ON T0.""Code"" = T2.""Code"" INNER JOIN ""@SBO_AST"" T3 ON T3.""Code"" = T2.""U_M3_1"" " &
    '"INNER JOIN ""@SBO_AST_C0"" T4 ON T4.""Code"" = T3.""Code"" WHERE T0.""U_Active"" = 'Y' AND UPPER(T0.""U_PCR"") = '" & status1.Replace("'", "''").ToUpper() & "' " &
    '"AND UPPER(T4.""U_AUTH"") = '" & loginUser.Replace("'", "''").ToUpper() & "'"
    '                                    oRS1.DoQuery(sSQL1)
    '                                    If oRS1.RecordCount > 0 Then
    '                                        'Connected user is an authorizer
    '                                        objMatrix1.Columns.Item("NRP").Editable = True
    '                                    Else
    '                                        'Connected user is NOT an authorizer
    '                                        objMatrix1.Columns.Item("NRP").Editable = False
    '                                    End If
    '                                    'objMatrix1.Columns.Item("NRP").Editable = False
    '                                    objMatrix1.Columns.Item("PUpdate").Editable = False
    '                                    objMatrix1.Columns.Item("PUS").Editable = False
    '                                    objMatrix1.Columns.Item("Col_0").Editable = False
    '                                    objMatrix1.Columns.Item("Col_1").Editable = False
    '                                    objMatrix1.Columns.Item("Col_2").Editable = True
    '                                    objMatrix1.Columns.Item("Col_3").Editable = True
    '                                    objMatrix1.Columns.Item("Col_4").Editable = True
    '                                    objMatrix1.Columns.Item("Col_6").Editable = True
    '                                    objMatrix1.Columns.Item("Col_7").Editable = True
    '                                    objForm.Items.Item("TNXBRC").Enabled = False
    '                                    objForm.Items.Item("TNXPUTP").Enabled = False
    '                                    objForm.Items.Item("TNXUSR").Enabled = False
    '                                    objForm.Items.Item("PName").Enabled = False

    '                                ElseIf (status1 = "Temporary Price Reduction" Or status1 = "TEMPORARY PRICE REDUCTION") And (status = "S") Then
    '                                    objMatrix4.Columns.Item("BrCode").Editable = False
    '                                    objMatrix4.Columns.Item("Qty").Editable = False
    '                                    Dim oRS1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '                                    Dim sSQL1 As String = "SELECT DISTINCT  T4.""U_AUTH"" FROM ""@SBO_APPHDR"" T0 " &
    '"INNER JOIN ""@SBO_APPAUT"" T2 ON T0.""Code"" = T2.""Code"" INNER JOIN ""@SBO_AST"" T3 ON T3.""Code"" = T2.""U_M3_1"" " &
    '"INNER JOIN ""@SBO_AST_C0"" T4 ON T4.""Code"" = T3.""Code"" WHERE T0.""U_Active"" = 'Y' AND UPPER(T0.""U_PCR"") = '" & status1.Replace("'", "''").ToUpper() & "' " &
    '"AND UPPER(T4.""U_AUTH"") = '" & loginUser.Replace("'", "''").ToUpper() & "'"
    '                                    oRS1.DoQuery(sSQL1)
    '                                    If oRS1.RecordCount > 0 Then
    '                                        'Connected user is an authorizer
    '                                        objMatrix4.Columns.Item("SPrice").Editable = True
    '                                    Else
    '                                        'Connected user is NOT an authorizer
    '                                        objMatrix4.Columns.Item("SPrice").Editable = False
    '                                    End If

    '                                    objMatrix4.Columns.Item("FundA").Editable = False
    '                                    objMatrix4.Columns.Item("Supp").Editable = False
    '                                    objMatrix4.Columns.Item("UOM1").Editable = False
    '                                    objMatrix4.Columns.Item("Remark").Editable = False
    '                                    objMatrix4.Columns.Item("MUpdated").Editable = False
    '                                    objMatrix4.Columns.Item("PUpdate").Editable = False
    '                                    objMatrix4.Columns.Item("PUS").Editable = False
    '                                    objMatrix4.Columns.Item("PriceR").Editable = False
    '                                    objMatrix4.Columns.Item("PRS").Editable = False
    '                                    objMatrix4.Columns.Item("PDN").Editable = False
    '                                    objMatrix4.Columns.Item("PS").Editable = False

    '                                    objMatrix4.Columns.Item("Col_0").Editable = True
    '                                    objForm.Items.Item("TNXBRC").Enabled = False
    '                                    objForm.Items.Item("TNXPUTP").Enabled = False
    '                                    objForm.Items.Item("TNXUSR").Enabled = False
    '                                    objForm.Items.Item("PName").Enabled = False

    '                                ElseIf (status1 = "Flyer Price Changes" Or status1 = "FLYER PRICE CHANGES") And (status = "S") Then
    '                                    objMatrix2.Columns.Item("BrCode").Editable = False
    '                                    Dim oRS1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '                                    Dim sSQL1 As String = "SELECT DISTINCT  T4.""U_AUTH"" FROM ""@SBO_APPHDR"" T0 " &
    '"INNER JOIN ""@SBO_APPAUT"" T2 ON T0.""Code"" = T2.""Code"" INNER JOIN ""@SBO_AST"" T3 ON T3.""Code"" = T2.""U_M3_1"" " &
    '"INNER JOIN ""@SBO_AST_C0"" T4 ON T4.""Code"" = T3.""Code"" WHERE T0.""U_Active"" = 'Y' AND UPPER(T0.""U_PCR"") = '" & status1.Replace("'", "''").ToUpper() & "' " &
    '"AND UPPER(T4.""U_AUTH"") = '" & loginUser.Replace("'", "''").ToUpper() & "'"
    '                                    oRS1.DoQuery(sSQL1)
    '                                    If oRS1.RecordCount > 0 Then
    '                                        'Connected user is an authorizer
    '                                        objMatrix2.Columns.Item("Col_1").Editable = True
    '                                    Else
    '                                        'Connected user is NOT an authorizer
    '                                        objMatrix2.Columns.Item("Col_1").Editable = False
    '                                    End If

    '                                    objMatrix2.Columns.Item("Col_4").Editable = False
    '                                    objMatrix2.Columns.Item("MUpdated").Editable = False
    '                                    objMatrix2.Columns.Item("UOM").Editable = False
    '                                    objMatrix2.Columns.Item("PUpdate").Editable = False
    '                                    objMatrix2.Columns.Item("PUS").Editable = False
    '                                    objMatrix2.Columns.Item("PDN").Editable = False
    '                                    objMatrix2.Columns.Item("PS").Editable = False
    '                                    objMatrix2.Columns.Item("Col_13").Editable = False

    '                                    objForm.Items.Item("FOUT").Enabled = True
    '                                    objForm.Items.Item("FDATEF").Enabled = False
    '                                    objForm.Items.Item("TDATEF").Enabled = False
    '                                    objForm.Items.Item("FTIMEF").Enabled = False
    '                                    objForm.Items.Item("TTIMEF").Enabled = False
    '                                    objForm.Items.Item("Item_22").Enabled = False

    '                                ElseIf (status1 = "MOP") And (status = "S") Then

    '                                    Try
    '                                        objMatrix3.Columns.Item("BCode").Editable = False
    '                                        Dim oRS1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '                                        Dim sSQL1 As String = "SELECT DISTINCT  T4.""U_AUTH"" FROM ""@SBO_APPHDR"" T0 " &
    '"INNER JOIN ""@SBO_APPAUT"" T2 ON T0.""Code"" = T2.""Code"" INNER JOIN ""@SBO_AST"" T3 ON T3.""Code"" = T2.""U_M3_1"" " &
    '"INNER JOIN ""@SBO_AST_C0"" T4 ON T4.""Code"" = T3.""Code"" WHERE T0.""U_Active"" = 'Y' AND UPPER(T0.""U_PCR"") = '" & status1.Replace("'", "''").ToUpper() & "' " &
    '"AND UPPER(T4.""U_AUTH"") = '" & loginUser.Replace("'", "''").ToUpper() & "'"
    '                                        oRS1.DoQuery(sSQL1)
    '                                        If oRS1.RecordCount > 0 Then
    '                                            'Connected user is an authorizer
    '                                            objMatrix3.Columns.Item("NSPrice").Editable = True
    '                                        Else
    '                                            'Connected user is NOT an authorizer
    '                                            objMatrix3.Columns.Item("NSPrice").Editable = False
    '                                        End If

    '                                        objMatrix3.Columns.Item("CName").Editable = False
    '                                        objMatrix3.Columns.Item("IAttach").Editable = False
    '                                        objMatrix3.Columns.Item("MUpdated").Editable = False
    '                                        objMatrix3.Columns.Item("UOM").Editable = False
    '                                        objMatrix3.Columns.Item("PUpdate").Editable = False
    '                                        objMatrix3.Columns.Item("PUS").Editable = False
    '                                        objMatrix3.Columns.Item("PriceR").Editable = False
    '                                        objMatrix3.Columns.Item("PDN").Editable = False
    '                                        objMatrix3.Columns.Item("PS").Editable = False

    '                                        objForm.Items.Item("FDATEM1").Enabled = False
    '                                        objForm.Items.Item("TDATEMP").Enabled = False
    '                                        objForm.Items.Item("FTIMEM").Enabled = False
    '                                        objForm.Items.Item("TTIMEM").Enabled = False

    '                                        objForm.Items.Item("TNXBRC").Enabled = False
    '                                        objForm.Items.Item("TNXPUTP").Enabled = False
    '                                        objForm.Items.Item("TNXUSR").Enabled = False
    '                                        objForm.Items.Item("PName").Enabled = False
    '                                    Catch
    '                                    End Try

    '                                ElseIf (status1 = "Every day Special Price") And (status = "S") Then
    '                                    objMatrix6.Columns.Item("BCode").Editable = False
    '                                    objMatrix6.Columns.Item("Col_3").Editable = False

    '                                    objMatrix6.Columns.Item("Col_7").Editable = False

    '                                    Dim oRS1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '                                    Dim sSQL1 As String = "SELECT DISTINCT  T4.""U_AUTH"" FROM ""@SBO_APPHDR"" T0 " &
    '"INNER JOIN ""@SBO_APPAUT"" T2 ON T0.""Code"" = T2.""Code"" INNER JOIN ""@SBO_AST"" T3 ON T3.""Code"" = T2.""U_M3_1"" " &
    '"INNER JOIN ""@SBO_AST_C0"" T4 ON T4.""Code"" = T3.""Code"" WHERE T0.""U_Active"" = 'Y' AND UPPER(T0.""U_PCR"") = '" & status1.Replace("'", "''").ToUpper() & "' " &
    '"AND UPPER(T4.""U_AUTH"") = '" & loginUser.Replace("'", "''").ToUpper() & "'"
    '                                    oRS1.DoQuery(sSQL1)
    '                                    If oRS1.RecordCount > 0 Then
    '                                        'Connected user is an authorizer
    '                                        objMatrix6.Columns.Item("Col_8").Editable = True
    '                                    Else
    '                                        'Connected user is NOT an authorizer
    '                                        objMatrix6.Columns.Item("Col_8").Editable = False
    '                                    End If

    '                                    objMatrix6.Columns.Item("Col_11").Editable = False
    '                                    objMatrix6.Columns.Item("Col_12").Editable = False

    '                                    objMatrix6.Columns.Item("IAttach").Editable = False
    '                                    objMatrix6.Columns.Item("Col_13").Editable = False

    '                                    objForm.Items.Item("TNXBRC").Enabled = False
    '                                    objForm.Items.Item("TNXPUTP").Enabled = False
    '                                    objForm.Items.Item("TNXUSR").Enabled = False
    '                                    objForm.Items.Item("PName").Enabled = False
    '                                ElseIf (status1 = "Bulk Cost Update") And (status = "S") Then
    '                                    objMatrix6.Columns.Item("BCode").Editable = False
    '                                    objMatrix6.Columns.Item("UOM").Editable = False
    '                                    objMatrix6.Columns.Item("Base").Editable = False
    '                                    objMatrix6.Columns.Item("Tax").Editable = False
    '                                    Dim oRS1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '                                    Dim sSQL1 As String = "SELECT DISTINCT  T4.""U_AUTH"" FROM ""@SBO_APPHDR"" T0 " &
    '"INNER JOIN ""@SBO_APPAUT"" T2 ON T0.""Code"" = T2.""Code"" INNER JOIN ""@SBO_AST"" T3 ON T3.""Code"" = T2.""U_M3_1"" " &
    '"INNER JOIN ""@SBO_AST_C0"" T4 ON T4.""Code"" = T3.""Code"" WHERE T0.""U_Active"" = 'Y' AND UPPER(T0.""U_PCR"") = '" & status1.Replace("'", "''").ToUpper() & "' " &
    '"AND UPPER(T4.""U_AUTH"") = '" & loginUser.Replace("'", "''").ToUpper() & "'"
    '                                    oRS1.DoQuery(sSQL1)
    '                                    If oRS1.RecordCount > 0 Then
    '                                        'Connected user is an authorizer
    '                                        objMatrix6.Columns.Item("NPRC").Editable = True
    '                                    Else
    '                                        'Connected user is NOT an authorizer
    '                                        objMatrix6.Columns.Item("NPRC").Editable = False
    '                                    End If


    '                                    objForm.Items.Item("FDATEB").Enabled = False
    '                                    objForm.Items.Item("TDATEB").Enabled = False
    '                                    objForm.Items.Item("TNXBRC").Enabled = False
    '                                    objForm.Items.Item("TNXPUTP").Enabled = False
    '                                    objForm.Items.Item("TNXUSR").Enabled = False
    '                                    objForm.Items.Item("PName").Enabled = False
    '                                ElseIf (status1 = "Cost Fluctuation Stock Transfer") And (status = "S") Then
    '                                    objMatrix6.Columns.Item("Col_3").Editable = False
    '                                    objMatrix6.Columns.Item("Col_12").Editable = False
    '                                    objMatrix6.Columns.Item("Col_13").Editable = False
    '                                    objMatrix6.Columns.Item("Col_14").Editable = False

    '                                    objForm.Items.Item("TNXBRC").Enabled = False
    '                                    objForm.Items.Item("TNXPUTP").Enabled = False
    '                                    objForm.Items.Item("TNXUSR").Enabled = False
    '                                    objForm.Items.Item("PName").Enabled = False
    '                                End If

    '                            Next

    '                            ' Enable OK and Cancel buttons
    '                            Try : objForm.Items.Item("1").Enabled = True : Catch : End Try   ' OK
    '                            Try : objForm.Items.Item("2").Enabled = True : Catch : End Try   ' Cancel
    '                            Try : objForm.Items.Item("Item_28").Enabled = True : Catch : End Try
    '                            Try : objForm.Items.Item("Item_8").Enabled = True : Catch : End Try
    '                            Try : objForm.Items.Item("Item_23").Enabled = True : Catch : End Try
    '                            Try : objForm.Items.Item("Item_13").Enabled = True : Catch : End Try

    '                            ' Enable Find menu
    '                            Try : objForm.EnableMenu("1281", True) : Catch : End Try
    '                            'Enable Add menu
    '                            Try : objForm.EnableMenu("1282", True) : Catch : End Try


    '                        Else

    '                            'Enable all items first
    '                            For i As Integer = 0 To objForm.Items.Count - 1
    '                                Try
    '                                    objForm.Items.Item(i).Enabled = True
    '                                Catch
    '                                End Try
    '                            Next

    '                            'Disable header fields
    '                            Try : objForm.Items.Item("PriceLNum").Enabled = False : Catch : End Try

    '                            Try : objForm.Items.Item("TNXDDT").Enabled = False : Catch : End Try
    '                            Try : objForm.Items.Item("DocNum").Enabled = False : Catch : End Try
    '                            Try : objForm.Items.Item("TNXSTUS").Enabled = False : Catch : End Try
    '                            Try : objForm.Items.Item("AppID").Enabled = False : Catch : End Try


    '                            ' Enable Add / Find menu
    '                            Try : objForm.EnableMenu("1281", True) : Catch : End Try
    '                            Try : objForm.Items.Item("Item_28").Enabled = True : Catch : End Try
    '                            'Try : objForm.Items.Item("Item_8").Enabled = True : Catch : End Try
    '                            'Try : objForm.Items.Item("Item_23").Enabled = True : Catch : End Try
    '                            ' Enable Submit button
    '                            If oDBs_Head.GetValue("U_TNXPUTP", 0).Trim() = "Flyer Price Changes" And oDBs_Head.GetValue("U_FOUT", 0).Trim() = "No" Then

    '                                ''Try : objForm.Items.Item("Submit").Enabled = False : Catch : End Try
    '                                Try : objForm.Items.Item("Item_28").Enabled = True : Catch : End Try
    '                            Else
    '                                ''Try : objForm.Items.Item("Submit").Enabled = True : Catch : End Try
    '                                Try : objForm.Items.Item("Item_28").Enabled = True : Catch : End Try
    '                            End If

    '                            If oDBs_Head.GetValue("U_TNXPUTP", 0).Trim() = "MOP" And oDBs_Head.GetValue("U_TNXSTUS", 0).Trim() = "C" Then

    '                                If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
    '                                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
    '                                End If
    '                            Else
    '                                Try
    '                                    If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
    '                                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
    '                                    End If
    '                                Catch
    '                                End Try

    '                            End If
    '                            ' Switch to OK MODE
    '                            If status = "OPEN" OrElse status = "O" Then

    '                                Try : objMatrix2 = objForm.Items.Item("Mtx2").Specific : objMatrix2.Columns.Item("BrCode").Editable = True : Catch : End Try
    '                                Try : objMatrix2 = objForm.Items.Item("Mtx2").Specific : objMatrix2.Columns.Item("Col_1").Editable = True : Catch : End Try
    '                                Try : objMatrix2 = objForm.Items.Item("Mtx2").Specific : objMatrix2.Columns.Item("Col_4").Editable = True : Catch : End Try
    '                                Try : objMatrix2 = objForm.Items.Item("Mtx2").Specific : objMatrix2.Columns.Item("MUpdated").Editable = True : Catch : End Try
    '                                Try : objMatrix2 = objForm.Items.Item("Mtx2").Specific : objMatrix2.Columns.Item("UOM").Editable = True : Catch : End Try
    '                                Try : objMatrix2 = objForm.Items.Item("Mtx2").Specific : objMatrix2.Columns.Item("PUS").Editable = True : Catch : End Try
    '                                Try : objMatrix2 = objForm.Items.Item("Mtx2").Specific : objMatrix2.Columns.Item("PDN").Editable = True : Catch : End Try
    '                                Try : objMatrix2 = objForm.Items.Item("Mtx2").Specific : objMatrix2.Columns.Item("PS").Editable = True : Catch : End Try
    '                                Try : objMatrix2 = objForm.Items.Item("Mtx2").Specific : objMatrix2.Columns.Item("Col_14").Editable = True : Catch : End Try
    '                                Try : objMatrix2 = objForm.Items.Item("Mtx2").Specific : objMatrix2.Columns.Item("Col_13").Editable = True : Catch : End Try

    '                                Try : objMatrix3 = objForm.Items.Item("Mtx3").Specific : objMatrix3.Columns.Item("BCode").Editable = True : Catch : End Try
    '                                Try : objMatrix3 = objForm.Items.Item("Mtx3").Specific : objMatrix3.Columns.Item("NSPrice").Editable = True : Catch : End Try
    '                                Try : objMatrix3 = objForm.Items.Item("Mtx3").Specific : objMatrix3.Columns.Item("CName").Editable = True : Catch : End Try
    '                                Try : objMatrix3 = objForm.Items.Item("Mtx3").Specific : objMatrix3.Columns.Item("IAttach").Editable = True : Catch : End Try
    '                                Try : objMatrix3 = objForm.Items.Item("Mtx3").Specific : objMatrix3.Columns.Item("MUpdated").Editable = True : Catch : End Try
    '                                Try : objMatrix3 = objForm.Items.Item("Mtx3").Specific : objMatrix3.Columns.Item("UOM").Editable = True : Catch : End Try
    '                                Try : objMatrix3 = objForm.Items.Item("Mtx3").Specific : objMatrix3.Columns.Item("PUpdate").Editable = True : Catch : End Try
    '                                Try : objMatrix3 = objForm.Items.Item("Mtx3").Specific : objMatrix3.Columns.Item("PUS").Editable = True : Catch : End Try
    '                                Try : objMatrix3 = objForm.Items.Item("Mtx3").Specific : objMatrix3.Columns.Item("PriceR").Editable = True : Catch : End Try
    '                                Try : objMatrix3 = objForm.Items.Item("Mtx3").Specific : objMatrix3.Columns.Item("PDN").Editable = True : Catch : End Try
    '                                Try : objMatrix3 = objForm.Items.Item("Mtx3").Specific : objMatrix3.Columns.Item("PS").Editable = True : Catch : End Try

    '                                Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("BrCode").Editable = True : Catch : End Try
    '                                Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("Qty").Editable = True : Catch : End Try
    '                                Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("Sprice").Editable = True : Catch : End Try
    '                                Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("FundA").Editable = True : Catch : End Try
    '                                Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("Supp").Editable = True : Catch : End Try
    '                                Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("UOM1").Editable = True : Catch : End Try
    '                                Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("Remark").Editable = True : Catch : End Try
    '                                Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("MUpdated").Editable = True : Catch : End Try
    '                                Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("PUpdate").Editable = True : Catch : End Try
    '                                Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("PUS").Editable = True : Catch : End Try
    '                                Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("PriceR").Editable = True : Catch : End Try
    '                                Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("PRS").Editable = True : Catch : End Try
    '                                Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("PDN").Editable = True : Catch : End Try
    '                                Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("PS").Editable = True : Catch : End Try
    '                                Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("Sell").Editable = True : Catch : End Try
    '                                Try : objMatrix4 = objForm.Items.Item("Mtx4").Specific : objMatrix4.Columns.Item("Col_0").Editable = True : Catch : End Try


    '                                Try : objMatrix6 = objForm.Items.Item("Mtx6").Specific : objMatrix6.Columns.Item("Col_0").Editable = True : Catch : End Try
    '                                Try : objMatrix6 = objForm.Items.Item("Mtx6").Specific : objMatrix6.Columns.Item("Col_3").Editable = True : Catch : End Try
    '                                Try : objMatrix6 = objForm.Items.Item("Mtx6").Specific : objMatrix6.Columns.Item("Col_7").Editable = True : Catch : End Try
    '                                Try : objMatrix6 = objForm.Items.Item("Mtx6").Specific : objMatrix6.Columns.Item("Col_8").Editable = True : Catch : End Try
    '                                Try : objMatrix6 = objForm.Items.Item("Mtx6").Specific : objMatrix6.Columns.Item("Col_11").Editable = True : Catch : End Try
    '                                Try : objMatrix6 = objForm.Items.Item("Mtx6").Specific : objMatrix6.Columns.Item("Col_12").Editable = True : Catch : End Try
    '                                Try : objMatrix6 = objForm.Items.Item("Mtx6").Specific : objMatrix6.Columns.Item("IAttach").Editable = True : Catch : End Try
    '                                Try : objMatrix6 = objForm.Items.Item("Mtx6").Specific : objMatrix6.Columns.Item("Col_13").Editable = True : Catch : End Try

    '                                Try : objMatrix7 = objForm.Items.Item("Mtx7").Specific : objMatrix7.Columns.Item("BCode").Editable = True : Catch : End Try
    '                                Try : objMatrix7 = objForm.Items.Item("Mtx7").Specific : objMatrix7.Columns.Item("UOM").Editable = True : Catch : End Try
    '                                Try : objMatrix7 = objForm.Items.Item("Mtx7").Specific : objMatrix7.Columns.Item("Base").Editable = True : Catch : End Try
    '                                Try : objMatrix7 = objForm.Items.Item("Mtx7").Specific : objMatrix7.Columns.Item("Tax").Editable = True : Catch : End Try
    '                                Try : objMatrix7 = objForm.Items.Item("Mtx7").Specific : objMatrix7.Columns.Item("NPRC").Editable = True : Catch : End Try

    '                            End If


    '                        End If


    '                        Try
    '                            Dim updateType As String = ""
    '                            Try : updateType = oDBs_Head.GetValue("U_TNXPUTP", 0).Trim() : Catch : End Try
    '                            If String.IsNullOrWhiteSpace(updateType) Then
    '                                Try
    '                                    Dim cb As SAPbouiCOM.ComboBox =
    '                                    CType(objForm.Items.Item("TNXPUTP").Specific, SAPbouiCOM.ComboBox)
    '                                    If cb IsNot Nothing AndAlso cb.Selected IsNot Nothing Then
    '                                        updateType = If(cb.Selected.Value, "").ToString().Trim()
    '                                    End If
    '                                Catch
    '                                End Try
    '                            End If
    '                            If Not String.IsNullOrWhiteSpace(updateType) Then
    '                                SetTabsByUpdateType(updateType)
    '                            Else
    '                                SetTabsByUpdateType("Cost Fluctuation")
    '                            End If

    '                            ' Show flyer fields only for Flyer Price Changes
    '                            Dim showFlyer As Boolean = (updateType = "Flyer Price Changes")
    '                            objForm.Items.Item("Item_57").Visible = showFlyer
    '                            objForm.Items.Item("Flyer").Visible = showFlyer
    '                            objForm.Items.Item("Item_10").Visible = showFlyer
    '                            objForm.Items.Item("Item_13").Visible = showFlyer
    '                            If Not showFlyer Then
    '                                Try : oDBs_Head.SetValue("U_Flyer", oDBs_Head.Offset, "") : Catch : End Try
    '                            End If


    '                            ' Show TPR fields only for TPR Price Changes
    '                            Dim showTPR As Boolean = (updateType = "Temporary Price Reduction")

    '                            objForm.Items.Item("Item_10").Visible = showTPR
    '                            objForm.Items.Item("Item_13").Visible = showTPR
    '                            If Not showTPR Then
    '                                Try : oDBs_Head.SetValue("U_TPRT", oDBs_Head.Offset, "") : Catch : End Try
    '                            End If



    '                        Catch ex As Exception
    '                            objMain.objApplication.StatusBar.SetText(
    '                            "SetTabs error: " & ex.Message,
    '                            SAPbouiCOM.BoMessageTime.bmt_Short,
    '                            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                        End Try

    '                        ' Scheduled fields visibility
    '                        'Try
    '                        '    Dim currentType As String = ""
    '                        '    Try : currentType = oDBs_Head.GetValue("U_TNXTYP", 0).Trim() : Catch : End Try
    '                        '    Dim showScheduled As Boolean = (currentType = "Scheduled Update")
    '                        '    Try : objForm.Items.Item("Item_38").Visible = showScheduled : Catch : End Try
    '                        '    Try : objForm.Items.Item("FDate").Visible = showScheduled : Catch : End Try
    '                        '    Try : objForm.Items.Item("Item_39").Visible = showScheduled : Catch : End Try
    '                        '    Try : objForm.Items.Item("FTime").Visible = showScheduled : Catch : End Try
    '                        '    If Not showScheduled Then
    '                        '        Try : oDBs_Head.SetValue("U_FDate", oDBs_Head.Offset, "") : Catch : End Try
    '                        '        Try : oDBs_Head.SetValue("U_FTime", oDBs_Head.Offset, "") : Catch : End Try
    '                        '    End If
    '                        'Catch
    '                        'End Try

    '                    Finally
    '                        objForm.Freeze(False)
    '                    End Try
    '                End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                    If BusinessObjectInfo.BeforeAction = False AndAlso BusinessObjectInfo.ActionSuccess = True Then
                        objForm.Freeze(True)


                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")

                        Try
                            ''Dim oChk As SAPbouiCOM.CheckBox = CType(objForm.Items.Item("Check").Specific, SAPbouiCOM.CheckBox)
                            'Dim strCheck As String = oDBs_Head.GetValue("U_Check", 0).Trim()
                            ''If oChk.Checked = True Then
                            'If strCheck = "Y" Then
                            '    objForm.Items.Item("FDTT").Enabled = True
                            '    objForm.Items.Item("TDTT").Enabled = True
                            'Else

                            '    objForm.Items.Item("FDTT").Enabled = False
                            '    objForm.Items.Item("TDTT").Enabled = False
                            'End If
                        Catch ex As Exception

                        End Try
                        objForm.Freeze(False)
                    End If


                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    If BusinessObjectInfo.BeforeAction = False Then

                        Dim objForm As SAPbouiCOM.Form
                        objForm = objMain.objApplication.Forms.Item(BusinessObjectInfo.FormUID)

                        'Enable Form Settings (1282)
                        Try
                            objForm.EnableMenu("1282", True)
                        Catch
                        End Try




                    End If






            End Select

        Catch ex As Exception
            Try : objForm.Freeze(False) : Catch : End Try
            objMain.objApplication.StatusBar.SetText(
            ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub CalculateGPTotals(ByVal objForm As SAPbouiCOM.Form)

        Dim objMatrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx2").Specific, SAPbouiCOM.Matrix)

        Dim TotalCost As Double = 0
        Dim TotalNormal As Double = 0
        Dim TotalFlyer As Double = 0

        For i As Integer = 1 To objMatrix.VisualRowCount

            Dim Cost As Double = 0
            Dim Normal As Double = 0
            Dim Flyer As Double = 0

            Double.TryParse(CType(objMatrix.Columns.Item("Cost").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value, Cost)
            Double.TryParse(CType(objMatrix.Columns.Item("Col_0").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value, Normal)
            Double.TryParse(CType(objMatrix.Columns.Item("Col_1").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value, Flyer)

            TotalCost += Cost
            TotalNormal += Normal
            TotalFlyer += Flyer
        Next

        Dim TotalNormalGP As Double = 0
        Dim TotalFlyerGP As Double = 0

        If TotalNormal <> 0 Then
            TotalNormalGP = ((TotalNormal - TotalCost) / TotalNormal) * 100
        End If

        If TotalFlyer <> 0 Then
            TotalFlyerGP = ((TotalFlyer - TotalCost) / TotalFlyer) * 100
        End If

        objMain.objApplication.StatusBar.SetText("Normal GP: " &
        TotalNormalGP.ToString("0.00") & "%  Flyer GP: " &
        TotalFlyerGP.ToString("0.00") & "%")

    End Sub
    Private Sub SetMatrixTotals(ByVal objForm As SAPbouiCOM.Form)

        Try
            Dim objMatrix As SAPbouiCOM.Matrix
            objMatrix = objForm.Items.Item("Mtx2").Specific

            objMatrix.Columns.Item("Cost").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix.Columns.Item("Col_0").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix.Columns.Item("Col_1").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            'CalculateTotalGP(objForm.UniqueID)
            'objMatrix.Columns.Item("Col_2").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            'objMatrix.Columns.Item("Col_3").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix.Columns.Item("Col_4").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix.Columns.Item("Col_5").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix.Columns.Item("Col_6").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix.Columns.Item("Col_7").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix.Columns.Item("Col_8").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix.Columns.Item("Col_9").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix.Columns.Item("Col_10").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix.Columns.Item("Col_14").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto


            ''objMatrix2.LoadFromDataSource()
            objMatrix.FlushToDataSource()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub


    Sub CFLFilter(ByVal FormUID As String, ByVal CFL_ID As String)
        Try

            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            oCondition = oConditions.Add()
            oCondition.Alias = "validFor"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "Y"

            oCondition.Alias = "CardType"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "C"
            oChooseFromList.SetConditions(oConditions)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "", Optional ByVal Series As Integer = 0)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C1")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C2")
            oDBs_Details4 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C3")
            oDBs_Details5 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C4")
            oDBs_Details6 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C5")
            oDBs_Details7 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C6")
            oDBs_Details8 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C7")
            oDBs_Details9 = objForm.DataSources.DBDataSources.Item("@TNX_PCRT_C8")
            objMatrix1 = objForm.Items.Item("Mtx1").Specific
            objMatrix2 = objForm.Items.Item("Mtx2").Specific
            objMatrix3 = objForm.Items.Item("Mtx3").Specific
            objMatrix4 = objForm.Items.Item("Mtx4").Specific
            objMatrix5 = objForm.Items.Item("Mtx5").Specific
            objMatrix6 = objForm.Items.Item("Mtx6").Specific
            objMatrix7 = objForm.Items.Item("Mtx7").Specific
            objMatrix8 = objForm.Items.Item("Mtx8").Specific
            objMatrix9 = objForm.Items.Item("Mtx9").Specific

            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TNXPCRUDO", "Primary"))
            oDBs_Head.SetValue("U_TNXDDT", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))

            ' Set default values: Status = "Open", Type = "Cost Fluctuation"
            Try
                ' Set DB datasource default values (header)
                objForm.Items.Item("Item_28").Enabled = True
                oDBs_Head.SetValue("U_TNXSTUS", oDBs_Head.Offset, "O")
                oDBs_Head.SetValue("U_TNXPUTP", oDBs_Head.Offset, "Cost Fluctuation")
            Catch ex As Exception
                ' ignore if fields not present
            End Try

            Try
                If objForm.PaneLevel = 2 Then
                    objForm.Items.Item("Item_8").Visible = True
                Else
                    objForm.Items.Item("Item_8").Visible = False
                End If



            Catch ex As Exception
            End Try

            'Current login user should load automatically

            ' Auto-fill current login user into header (U_TNXUSR) if field exists
            Try
                Dim currentUser As String = ""
                Try
                    currentUser = If(String.IsNullOrWhiteSpace(objMain.objCompany.UserName), objMain.objCompany.UserSignature.ToString(), objMain.objCompany.UserName)
                Catch
                    Try
                        currentUser = objMain.objCompany.UserSignature.ToString()
                    Catch
                        currentUser = ""
                    End Try
                End Try

                If Not String.IsNullOrWhiteSpace(currentUser) Then
                    Try
                        oDBs_Head.SetValue("U_TNXUSR", oDBs_Head.Offset, currentUser)
                    Catch
                        ' field might not exist in this DB schema - ignore
                    End Try
                End If
            Catch ex As Exception
                ' ignore failures setting user
            End Try

            ' Also set UI combo selection if controls exist
            Try
                Dim cbStatus As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("TNXSTUS").Specific, SAPbouiCOM.ComboBox)
                If cbStatus IsNot Nothing Then
                    Try
                        cbStatus.Select("O", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    Catch ex As Exception
                        ' maybe select by index or text if value differs
                        Try
                            cbStatus.Select("Open", SAPbouiCOM.BoSearchKey.psk_ByDescription)
                        Catch ex2 As Exception
                        End Try
                    End Try
                End If
            Catch ex As Exception
                ' ignore if control missing
            End Try

            Try
                Dim cbType As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("TNXPUTP").Specific, SAPbouiCOM.ComboBox)
                If cbType IsNot Nothing Then
                    Try
                        cbType.Select("Cost Fluctuation", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    Catch ex As Exception
                        Try
                            cbType.Select("Cost Fluctuation", SAPbouiCOM.BoSearchKey.psk_ByDescription)
                        Catch ex2 As Exception
                        End Try
                    End Try
                    ' ensure pane level matches the default type
                    Try
                        objForm.PaneLevel = 1 ' Cost Fluctuation -> pane 1
                    Catch ex As Exception
                    End Try
                End If
            Catch ex As Exception
                ' ignore if control missing
            End Try
            ' Populate Branch combobox (TNXBRC) with OUBR (Branches) valid values
            Try
                Dim oRsBr As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Try
                    oRsBr.DoQuery("SELECT ""Code"", ""Name"" FROM OUBR ORDER BY ""Name""")
                    Try
                        Dim cbBranch As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("TNXBRC").Specific, SAPbouiCOM.ComboBox)
                        If cbBranch IsNot Nothing Then
                            ' Clear existing valid values if any
                            Try
                                cbBranch.ValidValues.Clear()
                            Catch ex As Exception
                            End Try
                            While Not oRsBr.EoF
                                Dim codeVal As String = If(oRsBr.Fields.Item("Code").Value Is Nothing, "", oRsBr.Fields.Item("Code").Value.ToString())
                                Dim descVal As String = If(oRsBr.Fields.Item("Name").Value Is Nothing, "", oRsBr.Fields.Item("Name").Value.ToString())
                                Try
                                    cbBranch.ValidValues.Add(codeVal, descVal)
                                Catch ex As Exception
                                    ' ignore duplicates / invalid entries
                                End Try
                                oRsBr.MoveNext()
                            End While
                        End If
                    Catch ex As Exception
                        ' ignore if control missing or cast fails
                    End Try
                Catch ex As Exception
                    ' query error - ignore but log
                    objMain.objUtilities.Write_Log("Error loading branches: " & ex.Message)
                Finally
                    Try
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRsBr)
                    Catch
                    End Try
                    oRsBr = Nothing
                End Try
            Catch ex As Exception
                ' ignore if recordset creation fails
            End Try

            objMatrix1.Clear()
            objMatrix2.Clear()
            objMatrix3.Clear()
            objMatrix4.Clear()
            objMatrix5.Clear()
            objMatrix6.Clear()
            objMatrix7.Clear()
            objMatrix8.Clear()
            objMatrix9.Clear()
            oDBs_Details1.Clear()
            oDBs_Details2.Clear()
            oDBs_Details3.Clear()
            oDBs_Details4.Clear()
            oDBs_Details5.Clear()
            oDBs_Details6.Clear()
            oDBs_Details7.Clear()
            oDBs_Details8.Clear()
            oDBs_Details9.Clear()
            objMatrix1.FlushToDataSource()
            objMatrix2.FlushToDataSource()
            objMatrix3.FlushToDataSource()
            objMatrix4.FlushToDataSource()
            objMatrix5.FlushToDataSource()
            objMatrix6.FlushToDataSource()
            objMatrix7.FlushToDataSource()
            objMatrix8.FlushToDataSource()
            objMatrix9.FlushToDataSource()
            Me.SetNewLine1(objForm.UniqueID)
            Me.SetNewLine2(objForm.UniqueID)
            Me.SetNewLine3(objForm.UniqueID)
            Me.SetNewLine4(objForm.UniqueID)
            Me.SetNewLine5(objForm.UniqueID)
            Me.SetNewLine6(objForm.UniqueID)
            Me.SetNewLine7(objForm.UniqueID)
            Me.SetNewLine8(objForm.UniqueID)
            Me.SetNewLine9(objForm.UniqueID)
            LoadBranchCombo()
            ''LoadStatusCombo()
            ' LoadUOMMOPMatxix()
            Me.AutoDocentryNumber()


            ' Hide Flyer by default
            Try
                objForm.Items.Item("Item_57").Visible = False
                objForm.Items.Item("Flyer").Visible = False
            Catch ex As Exception
                ' ignore if control not found
            End Try

            ' Hide Scheduled Update fields by default unless TNXTYP already set to Scheduled Update
            'Try
            '    Dim currentType As String = ""
            '    Try
            '        currentType = oDBs_Head.GetValue("U_TNXTYP", 0).Trim()
            '    Catch ex As Exception
            '        currentType = ""
            '    End Try

            '    Dim showScheduled As Boolean = (currentType = "Scheduled Update")
            '    Try
            '        objForm.Items.Item("Item_38").Visible = showScheduled
            '    Catch ex As Exception
            '    End Try
            '    Try
            '        objForm.Items.Item("FDate").Visible = showScheduled
            '    Catch ex As Exception
            '    End Try
            '    Try
            '        objForm.Items.Item("Item_39").Visible = showScheduled
            '    Catch ex As Exception
            '    End Try
            '    Try
            '        objForm.Items.Item("FTime").Visible = showScheduled
            '    Catch ex As Exception
            '    End Try

            '    ' If hidden, clear any existing values
            '    If Not showScheduled Then
            '        Try
            '            oDBs_Head.SetValue("U_FDate", oDBs_Head.Offset, "")
            '            oDBs_Head.SetValue("U_FTime", oDBs_Head.Offset, "")
            '        Catch ex As Exception
            '        End Try
            '    End If
            'Catch ex As Exception
            '    ' ignore
            'End Try

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub LoadStatusCombo()
        Try
            Dim oCombo As SAPbouiCOM.ComboBox
            oCombo = CType(objForm.Items.Item("TNXSTUS").Specific, SAPbouiCOM.ComboBox)

            'Add only Open status
            If Not ComboValueExists(oCombo, "O") Then
                oCombo.ValidValues.Add("O", "Open")
            End If

            oCombo.Select("O", SAPbouiCOM.BoSearchKey.psk_ByValue)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "LoadStatusCombo ERROR: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub LoadBranchCombo()
        Try
            Dim oCombo As SAPbouiCOM.ComboBox
            oCombo = CType(objForm.Items.Item("TNXBRC").Specific, SAPbouiCOM.ComboBox)

            'oCombo.ValidValues.Clear()
            Try
                oCombo.ValidValues.Remove("-2", SAPbouiCOM.BoSearchKey.psk_ByValue)
            Catch
                'Main does not exist
            End Try

            Dim oRs As SAPbobsCOM.Recordset =
            objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'oRs.DoQuery("SELECT ""BPLId"", ""BPLName"", ""U_PricelistNo"" FROM ""OBPL"" ORDER BY ""BPLId""")

            oRs.DoQuery("SELECT T2.""BPLId"",T2.""BPLName"" FROM USR6 T0  INNER JOIN OUSR T1 ON T0.""UserCode"" = T1.""USER_CODE"" " &
                    " INNER JOIN OBPL T2 ON T0.""BPLId"" = T2.""BPLId"" WHERE T1.""USER_CODE"" = '" & objMain.objCompany.UserName.ToString().Trim & "' and ""Disabled"" = 'N'")


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
            objMain.objApplication.StatusBar.SetText(
            "LoadBranchCombo ERROR: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub LoadBranchCombofind(ByVal DocNumberPCR As String)
        Try
            Dim oCombo As SAPbouiCOM.ComboBox
            oCombo = CType(objForm.Items.Item("TNXBRC").Specific, SAPbouiCOM.ComboBox)


            'oCombo.ValidValues.Clear()
            Try
                oCombo.ValidValues.Remove("-2", SAPbouiCOM.BoSearchKey.psk_ByValue)
            Catch
                'Main does not exist
            End Try

            Dim oRs As SAPbobsCOM.Recordset =
            objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRsF As SAPbobsCOM.Recordset =
            objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oRsF.DoQuery("SELECT T1.""U_TNXBRC"" as ""BPLId"" FROM ""@TNX_PCRT"" T1 WHERE T1.""U_AppID"" = '" & DocNumberPCR & "'")

            Dim BPLValue As String = oRsF.Fields.Item("BPLId").Value.ToString().Trim()
            Dim qry As String = ""

            If IsNumeric(BPLValue) Then
                qry = "SELECT T2.""BPLId"", T2.""BPLName"" FROM USR6 T0 " &
          "INNER JOIN OUSR T1 ON T0.""UserCode"" = T1.""USER_CODE"" " &
          "INNER JOIN OBPL T2 ON T0.""BPLId"" = T2.""BPLId"" " &
          "WHERE T1.""USER_CODE"" = '" & objMain.objCompany.UserName.ToString().Trim() & "' " &
          "AND T2.""BPLId"" = '" & BPLValue & "'"
            Else
                qry = "SELECT T2.""BPLId"", T2.""BPLName"" FROM USR6 T0 " &
          "INNER JOIN OUSR T1 ON T0.""UserCode"" = T1.""USER_CODE"" " &
          "INNER JOIN OBPL T2 ON T0.""BPLId"" = T2.""BPLId"" " &
          "WHERE T1.""USER_CODE"" = '" & objMain.objCompany.UserName.ToString().Trim() & "' " &
          "AND T2.""BPLName"" = '" & BPLValue & "'"
            End If

            oRs.DoQuery(qry)

            ''oRs.DoQuery("SELECT ""BPLId"", ""BPLName"", ""U_PricelistNo"" FROM ""OBPL"" ORDER BY ""BPLId""")



            'oRs.DoQuery("SELECT T2.""BPLId"",T2.""BPLName"" FROM USR6 T0  INNER JOIN OUSR T1 ON T0.""UserCode"" = T1.""USER_CODE"" " &
            '        " INNER JOIN OBPL T2 ON T0.""BPLId"" = T2.""BPLId"" WHERE T1.""USER_CODE"" = '" & objMain.objCompany.UserName.ToString().Trim & "' and T0.""BPLId""= '" & oRsF.Fields.Item("BPLId").Value.ToString() & "'")

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
            objMain.objApplication.StatusBar.SetText(
            "LoadBranchCombo ERROR: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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


    'Tabs method
    Private Sub SetTabsByUpdateType(ByVal updateType As String)

        ' Disable all tabs first
        objForm.Items.Item("Tab1").Enabled = False   ' Cost Fluctuation
        objForm.Items.Item("Tab2").Enabled = False   ' Flyer
        objForm.Items.Item("Tab3").Enabled = False  ' MOP
        objForm.Items.Item("Tab4").Enabled = False  ' Temporary Price Reduction
        objForm.Items.Item("Tab6").Enabled = False
        objForm.Items.Item("Tab7").Enabled = False
        objForm.Items.Item("Tab8").Enabled = False
        objForm.Items.Item("Tab9").Enabled = False
        objForm.Items.Item("Tab5").Enabled = True  ' Attachment

        ' Enable only selected tab
        Select Case updateType

            Case "Cost Fluctuation"
                objForm.Items.Item("Tab1").Enabled = True
                objForm.Items.Item("Tab1").Click()

            Case "Flyer Price Changes"
                objForm.Items.Item("Tab2").Enabled = True
                objForm.Items.Item("Tab2").Click()

            Case "MOP"
                objForm.Items.Item("Tab3").Enabled = True
                objForm.Items.Item("Tab3").Click()

            Case "Temporary Price Reduction"
                objForm.Items.Item("Tab4").Enabled = True
                objForm.Items.Item("Tab4").Click()
            Case "Every day Special Price"
                objForm.Items.Item("Tab6").Enabled = True
                objForm.Items.Item("Tab6").Click()

            Case "Bulk Cost Update"
                objForm.Items.Item("Tab7").Enabled = True
                objForm.Items.Item("Tab7").Click()
            Case "Cost Fluctuation Stock Transfer"
                objForm.Items.Item("Tab8").Enabled = True
                objForm.Items.Item("Tab8").Click()
            Case "Attachment Tab"
                objForm.Items.Item("Tab5").Enabled = True
                objForm.Items.Item("Tab5").Click()
            Case "Bulk Price Update"
                objForm.Items.Item("Tab9").Enabled = True
                objForm.Items.Item("Tab9").Click()

        End Select

    End Sub

    'Validation for  Barcode should validate
    Private Function ValidateBarcodeByTab(ByVal objForm As SAPbouiCOM.Form) As Boolean
        Try

            Dim objTabCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("TNXPUTP").Specific

            Dim TabID As String
            TabID = objTabCombo.Selected.Value


            Select Case TabID

                Case "Flyer Price Changes"
                    Return ValidateBarCodeMatrix(objForm, "Mtx2", "BrCode")

                Case "MOP"
                    Return ValidateBarCodeMatrix(objForm, "Mtx3", "BCode")

                Case "Temporary Price Reduction"
                    Return ValidateBarCodeMatrix(objForm, "Mtx4", "BrCode")

                Case Else
                    Return True

            End Select

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Private Function ValidateBarCodeMatrix(ByVal objForm As SAPbouiCOM.Form,
                                       ByVal matrixId As String,
                                       ByVal columnId As String) As Boolean
        Try
            Dim oMatrix As SAPbouiCOM.Matrix =
            CType(objForm.Items.Item(matrixId).Specific, SAPbouiCOM.Matrix)

            Dim oRS As SAPbobsCOM.Recordset =
            objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            For i As Integer = 1 To oMatrix.RowCount

                Dim txt As SAPbouiCOM.EditText =
                CType(oMatrix.Columns.Item(columnId).Cells.Item(i).Specific, SAPbouiCOM.EditText)

                Dim selValue As String = txt.Value.Trim()

                '-----------------------------
                ' Skip empty rows
                '-----------------------------
                If selValue = "" Then
                    Continue For
                End If

                '-----------------------------
                ' Validate barcode existence
                '-----------------------------
                Dim qry As String =
                "SELECT COUNT(*) FROM ""OBCD"" WHERE ""BcdCode"" = '" &
                selValue.Replace("'", "''") & "'"

                oRS.DoQuery(qry)

                If Convert.ToInt32(oRS.Fields.Item(0).Value) = 0 Then
                    objMain.objApplication.StatusBar.SetText(
                    "Invalid Barcode at row " & i,
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

    'Current cost is not loading  - Should load automatically 
    'Current cost - returns avg price as string
    Private Function FetchBranchAndAvgPriceByBarcode(ByVal barcode As String) As String

        Dim avgPrice As String = "0"

        Try
            If String.IsNullOrWhiteSpace(barcode) Then Return avgPrice

            Dim bc As String = barcode.Trim().Replace("'", "''")

            Dim rs As SAPbobsCOM.Recordset =
                objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


            Dim qItem As String =
                "SELECT ""ItemCode"" FROM ""OBCD"" WHERE ""BcdCode"" = '" & bc & "'"

            rs.DoQuery(qItem)
            If rs.RecordCount = 0 Then GoTo CleanExit

            Dim itemCode As String = rs.Fields.Item(0).Value.ToString()


            Dim branchId As String = ""

            Try
                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCRT")
                branchId = oDBs_Head.GetValue("U_TNXBRC", oDBs_Head.Offset).Trim()
            Catch
            End Try

            If String.IsNullOrEmpty(branchId) Then

                Dim qBranch As String = "SELECT T0.""BPLid"" " & "FROM ""OWHS"" T0 " & "JOIN ""OITW"" T1 ON T0.""WhsCode"" = T1.""WhsCode"" " &
                                         "WHERE T1.""ItemCode"" = '" & itemCode.Replace("'", "''") & "' " & "ORDER BY T1.""OnHand"" DESC LIMIT 1"

                rs.DoQuery(qBranch)

                If rs.RecordCount > 0 Then
                    branchId = rs.Fields.Item(0).Value.ToString()

                    Try
                        oDBs_Head.SetValue("U_TNXBRC", oDBs_Head.Offset, branchId)
                    Catch
                    End Try
                End If
            End If


            Dim qAvg As String = "SELECT IFNULL(T1.""AvgPrice"",0) " & "FROM ""OWHS"" T0 " & "JOIN ""OITW"" T1 ON T0.""WhsCode"" = T1.""WhsCode"" " &
                                 "WHERE T1.""ItemCode"" = '" & itemCode.Replace("'", "''") & "' " & "AND T0.""BPLid"" = '" & branchId & "'"

            rs.DoQuery(qAvg)

            If rs.RecordCount > 0 Then
                avgPrice = rs.Fields.Item(0).Value.ToString()
            End If

            Try
                Dim fields() As String = {"U_AVHPrice", "U_AvgPrice", "U_UPRICE", "U_AVGPRICE"}
                For Each f In fields
                    Try
                        oDBs_Head.SetValue(f, oDBs_Head.Offset, avgPrice)
                        Exit For
                    Catch
                    End Try
                Next
            Catch
            End Try
CleanExit:
            Try
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
            Catch
            End Try

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("FetchBranchAndAvgPrice error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

        Return avgPrice
    End Function

    'copy and  paste in mop tab code here
    Private Function GetDataFromMatrix() As DataTable
        Dim dt As New DataTable()
        dt.Columns.Add("ItemCode")
        dt.Columns.Add("Bcode")

        Try
            objMatrix3 = objForm.Items.Item("Mtx3").Specific
            ' Ensure any pending edits are committed to the matrix
            Try
                objMatrix3.FlushToDataSource()
            Catch
                ' ignore flush errors, we'll read UI directly
            End Try

            Dim visualRows As Integer = objMatrix3.VisualRowCount
            For r As Integer = 1 To visualRows
                Dim itemCode As String = ""
                Dim barcode As String = ""

                Try
                    itemCode = If(objMatrix3.Columns.Item("ItemCode").Cells.Item(r).Specific.Value, "").ToString().Trim()
                Catch
                    itemCode = ""
                End Try

                Try
                    barcode = If(objMatrix3.Columns.Item("BCode").Cells.Item(r).Specific.Value, "").ToString().Trim()
                Catch
                    barcode = ""
                End Try

                If itemCode <> "" OrElse barcode <> "" Then
                    dt.Rows.Add(itemCode, barcode)
                End If
            Next
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("GetDataFromMatrix error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

        Return dt
    End Function


    'Copy and paste in EDSP tab code here
    Private Function GetDataFromMatrixEDSP() As DataTable
        Dim dt As New DataTable()
        dt.Columns.Add("ItemCode")
        dt.Columns.Add("Bcode")

        Try
            objMatrix6 = objForm.Items.Item("Mtx6").Specific
            ' Ensure any pending edits are committed to the matrix
            Try
                objMatrix6.FlushToDataSource()
            Catch
                ' ignore flush errors, we'll read UI directly
            End Try

            Dim visualRows As Integer = objMatrix6.VisualRowCount
            For r As Integer = 1 To visualRows
                Dim itemCode As String = ""
                Dim barcode As String = ""

                Try
                    itemCode = If(objMatrix6.Columns.Item("Col_1").Cells.Item(r).Specific.Value, "").ToString().Trim()
                Catch
                    itemCode = ""
                End Try

                Try
                    barcode = If(objMatrix6.Columns.Item("Col_0").Cells.Item(r).Specific.Value, "").ToString().Trim()
                Catch
                    barcode = ""
                End Try

                If itemCode <> "" OrElse barcode <> "" Then
                    dt.Rows.Add(itemCode, barcode)
                End If
            Next
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("GetDataFromMatrix error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

        Return dt
    End Function

    'Copy and paste in BCU tab code here
    Private Function GetDataFromMatrixBCU() As DataTable
        Dim dt As New DataTable()
        dt.Columns.Add("ItemCode")
        dt.Columns.Add("Bcode")

        Try
            objMatrix6 = objForm.Items.Item("Mtx7").Specific
            ' Ensure any pending edits are committed to the matrix
            Try
                objMatrix6.FlushToDataSource()
            Catch
                ' ignore flush errors, we'll read UI directly
            End Try

            Dim visualRows As Integer = objMatrix7.VisualRowCount
            For r As Integer = 1 To visualRows
                Dim itemCode As String = ""
                Dim barcode As String = ""

                Try
                    itemCode = If(objMatrix7.Columns.Item("ITCD").Cells.Item(r).Specific.Value, "").ToString().Trim()
                Catch
                    itemCode = ""
                End Try

                Try
                    barcode = If(objMatrix7.Columns.Item("BCode").Cells.Item(r).Specific.Value, "").ToString().Trim()
                Catch
                    barcode = ""
                End Try

                If itemCode <> "" OrElse barcode <> "" Then
                    dt.Rows.Add(itemCode, barcode)
                End If
            Next
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("GetDataFromMatrix error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

        Return dt
    End Function

    'copy and paste Ain Flyer tab code here
    Private Function GetDataFromMatrixFlyer() As DataTable
        Dim dt As New DataTable()
        dt.Columns.Add("ItemCode")
        dt.Columns.Add("Bcode")

        Try
            objMatrix2 = objForm.Items.Item("Mtx2").Specific
            ' Ensure any pending edits are committed to the matrix
            Try
                objMatrix2.FlushToDataSource()
            Catch
                ' ignore flush errors, we'll read UI directly
            End Try

            Dim visualRows As Integer = objMatrix2.VisualRowCount
            For r As Integer = 1 To visualRows
                Dim itemCode As String = ""
                Dim barcode As String = ""

                Try
                    itemCode = If(objMatrix2.Columns.Item("ItemCode").Cells.Item(r).Specific.Value, "").ToString().Trim()
                Catch
                    itemCode = ""
                End Try

                Try
                    barcode = If(objMatrix2.Columns.Item("BrCode").Cells.Item(r).Specific.Value, "").ToString().Trim()
                Catch
                    barcode = ""
                End Try

                If itemCode <> "" OrElse barcode <> "" Then
                    dt.Rows.Add(itemCode, barcode)
                End If
            Next
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("GetDataFromMatrixFlyer error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

        Return dt
    End Function

    'copy and paste in Tempoary tab code here
    Private Function GetDataFromMatrixTempoary() As DataTable
        Dim dt As New DataTable()
        dt.Columns.Add("ItemCode")
        dt.Columns.Add("Bcode")

        Try
            objMatrix4 = objForm.Items.Item("Mtx4").Specific
            ' Ensure any pending edits are committed to the matrix
            Try
                objMatrix4.FlushToDataSource()
            Catch
                ' ignore flush errors, we'll read UI directly
            End Try

            Dim visualRows As Integer = objMatrix4.VisualRowCount
            For r As Integer = 1 To visualRows
                Dim itemCode As String = ""
                Dim barcode As String = ""

                Try
                    itemCode = If(objMatrix4.Columns.Item("ItemCode").Cells.Item(r).Specific.Value, "").ToString().Trim()
                Catch
                    itemCode = ""
                End Try

                Try
                    barcode = If(objMatrix4.Columns.Item("BrCode").Cells.Item(r).Specific.Value, "").ToString().Trim()
                Catch
                    barcode = ""
                End Try

                If itemCode <> "" OrElse barcode <> "" Then
                    dt.Rows.Add(itemCode, barcode)
                End If
            Next
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("GetDataFromMatrixTemporay error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

        Return dt
    End Function


    'Copy and paste in BPU tab code here
    Private Function GetDataFromMatrixBPU() As DataTable
        Dim dt As New DataTable()
        dt.Columns.Add("ItemCode")
        dt.Columns.Add("BCode")

        Try
            objMatrix9 = objForm.Items.Item("Mtx9").Specific
            ' Ensure any pending edits are committed to the matrix
            Try
                objMatrix9.FlushToDataSource()
            Catch
                ' ignore flush errors, we'll read UI directly
            End Try

            Dim visualRows As Integer = objMatrix9.VisualRowCount
            For r As Integer = 1 To visualRows
                Dim itemCode As String = ""
                Dim barcode As String = ""

                Try
                    itemCode = If(objMatrix9.Columns.Item("Col_5").Cells.Item(r).Specific.Value, "").ToString().Trim()
                Catch
                    itemCode = ""
                End Try

                Try
                    barcode = If(objMatrix9.Columns.Item("Col_0").Cells.Item(r).Specific.Value, "").ToString().Trim()
                Catch
                    barcode = ""
                End Try

                If itemCode <> "" OrElse barcode <> "" Then
                    dt.Rows.Add(itemCode, barcode)
                End If
            Next
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("GetDataFromMatrix error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

        Return dt
    End Function

    ' 'copy and  paste in mop tab code here
    ' 'copy and  paste in mop tab code here
    Private Sub UploadFromDataTablebackup(dtNew As DataTable)
        Dim ItemCode As String = ""
        Dim Barcode As String = ""
        Dim Pricelist, nPrice, StartDate As String
        objMatrix3 = objForm.Items.Item("Mtx3").Specific
        objMatrix3.Clear()
        Dim oRS As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRecordset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        ' Dim CreatedBy As String = objMain.objCompany.UserName
        Dim record As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim userCode As String = objMain.objCompany.UserName.Replace("'", "''")

        oRS.DoQuery("SELECT ""USERID"" FROM OUSR WHERE ""USER_CODE"" = '" & userCode & "'")

        Dim CreatedBy As Integer = 0
        If Not oRS.EoF Then
            CreatedBy = CInt(oRS.Fields.Item("USERID").Value)
        End If



        If dtNew.Rows.Count = 0 OrElse dtNew.Columns.Count = 0 Then
            objMain.objApplication.StatusBar.SetText("No data found in Excel sheet.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Exit Sub
        End If

        ' determine which column is Barcode / ItemCode using header sniffing + content sampling
        Dim idxItemCode As Integer = -1
        Dim idxBarcode As Integer = -1
        ' header name sniff
        For c As Integer = 0 To dtNew.Columns.Count - 1
            Dim nm As String = If(dtNew.Columns(c).ColumnName, "").ToString().Trim().ToLower()
            If nm.Contains("barcode") Or nm.Contains("bar code") Or nm.Contains("bcd") Then idxBarcode = c
            If nm.Contains("item") Or nm.Contains("itemcode") Or nm.Contains("item code") Then idxItemCode = c
        Next

        ' content sampling if header sniff failed or ambiguous
        If idxBarcode = -1 Or idxItemCode = -1 Then
            Dim likelyBarcodeCols As New List(Of Integer)
            For c As Integer = 0 To dtNew.Columns.Count - 1
                If ColumnLikelyBarcode(dtNew, c) Then likelyBarcodeCols.Add(c)
            Next
            ' prefer first likely barcode column
            If likelyBarcodeCols.Count > 0 And idxBarcode = -1 Then
                idxBarcode = likelyBarcodeCols(0)
            End If
            ' set itemcode to first non-barcode column
            If idxItemCode = -1 Then
                For c As Integer = 0 To dtNew.Columns.Count - 1
                    If c <> idxBarcode Then
                        idxItemCode = c
                        Exit For
                    End If
                Next
            End If
        End If

        ' fallback heuristics (preserve previous behavior)
        If idxBarcode = -1 Then
            If dtNew.Columns.Count = 1 Then
                idxBarcode = 0
            ElseIf dtNew.Columns.Count >= 2 Then
                idxBarcode = 1
            End If
        End If
        If idxItemCode = -1 AndAlso dtNew.Columns.Count >= 1 Then
            If dtNew.Columns.Count >= 2 Then
                idxItemCode = 0
            Else
                idxItemCode = -1
            End If
        End If

        ' === Create table if not exists ===
        Dim createTableSQL As String = "CREATE COLUMN TABLE ""TNX_PRICEUPDATE""(""BARCODE"" NVARCHAR(100),  
                                       ""NEWSELLINGPRICE"" DECIMAL(11,5) CS_FIXED, ""CREATEDBY"" NVARCHAR(100), ""LINENUM"" INTEGER,""UPDT"" NVARCHAR(100)) 
                                        UNLOAD PRIORITY 5 AUTO MERGE;"

        Try
            oRS.DoQuery(createTableSQL)
        Catch ex As Exception
            If Not ex.Message.ToLower().Contains("already exists") AndAlso Not ex.Message.ToLower().Contains("cannot use duplicate table name") Then
                objMain.objApplication.StatusBar.SetText("Error creating table: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        End Try

        ' delete old staged rows for this createdBy
        Try
            oRS.DoQuery("DELETE FROM ""TNX_PRICEUPDATE"" WHERE ""CREATEDBY"" = '" & CreatedBy & "'")
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error deleting old records: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End Try

        objMain.objApplication.StatusBar.SetText("Inserting records into SAP HANA...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        ' Bulk insert using UNION ALL; batch to avoid overly long SQL
        Try
            Dim sb As New StringBuilder()
            Dim rowsAdded As Boolean = False
            Const MAX_BATCH_LENGTH As Integer = 800000 ' adjust if needed

            For i As Integer = 0 To dtNew.Rows.Count - 1
                If dtNew.Rows(i).ItemArray.All(Function(x) x Is DBNull.Value OrElse String.IsNullOrWhiteSpace(x.ToString())) Then Continue For

                Dim rawItem As String = ""
                Dim rawBarcode As String = ""
                ' Replace the loop body inside Read13 — after resolving rawItem/rawBarcode add a barcode lookup when only ItemCode is provided.
                If idxItemCode >= 0 AndAlso idxItemCode < dtNew.Columns.Count Then
                    If Not IsDBNull(dtNew.Rows(i).Item(idxItemCode)) Then rawItem = dtNew.Rows(i).Item(idxItemCode).ToString().Trim()
                End If
                If idxBarcode >= 0 AndAlso idxBarcode < dtNew.Columns.Count Then
                    If Not IsDBNull(dtNew.Rows(i).Item(idxBarcode)) Then rawBarcode = dtNew.Rows(i).Item(idxBarcode).ToString().Trim()
                End If

                ' single column case: decide by content whether it's barcode or itemcode
                If dtNew.Columns.Count = 1 Then
                    Dim val As String = If(Not IsDBNull(dtNew.Rows(i).Item(0)), dtNew.Rows(i).Item(0).ToString().Trim(), "")
                    If IsNumericOnly(val) AndAlso val.Length >= 6 Then
                        rawBarcode = val
                        rawItem = ""
                    Else
                        rawItem = val
                        rawBarcode = ""
                    End If
                End If

                ' NEW: if user supplied ItemCode but no Barcode, try to resolve barcode from OBCD
                If String.IsNullOrEmpty(rawBarcode) AndAlso Not String.IsNullOrEmpty(rawItem) Then
                    Try
                        Dim oRSBarcode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim q As String = "SELECT ""BcdCode"" FROM OBCD WHERE ""ItemCode"" = '" & rawItem.Replace("'", "''") & "' LIMIT 1"
                        oRSBarcode.DoQuery(q)
                        If oRSBarcode.RecordCount > 0 Then
                            rawBarcode = If(oRSBarcode.Fields.Item(0).Value Is Nothing, "", oRSBarcode.Fields.Item(0).Value.ToString().Trim())
                        End If
                    Catch ex As Exception
                        ' ignore lookup failures — proceed without barcode
                    End Try
                End If

                ' Continue with existing insert logic
                If String.IsNullOrEmpty(rawItem) AndAlso String.IsNullOrEmpty(rawBarcode) Then Continue For

                rawItem = rawItem.Replace("'", "''")
                rawBarcode = rawBarcode.Replace("'", "''")

                Dim lineNumber As Integer = i + 1


                '-------------------------------------------------
                If Not rowsAdded Then
                    sb.AppendLine("INSERT INTO ""TNX_PRICEUPDATE"" (" &
                                  """BARCODE"", ""CREATEDBY"", ""LINENUM"", ""UPDT"")")
                    sb.Append("SELECT ")
                    rowsAdded = True
                Else
                    sb.AppendLine()
                    sb.Append("UNION ALL SELECT ")
                End If


                sb.AppendFormat(
                    "'{0}','{1}',{2},CURRENT_DATE FROM DUMMY",
                    rawBarcode.Replace("'", "''"),
                    CreatedBy,
                    lineNumber
                )

                If sb.Length >= MAX_BATCH_LENGTH Then
                    oRS.DoQuery(sb.ToString())
                    sb.Clear()
                    rowsAdded = False
                End If


            Next

            If sb.Length > 0 Then
                oRS.DoQuery(sb.ToString())
            End If
            '=============================
            ' Validate staged data exists
            '=============================
            Dim chkRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'chkRS.DoQuery("SELECT COUNT(*) FROM ""TNX_PRICEUPDATE"" WHERE ""CREATEDBY"" = '" & CreatedBy.Replace("'", "''") & "'")
            chkRS.DoQuery("SELECT COUNT(*) FROM ""TNX_PRICEUPDATE"" WHERE ""CREATEDBY"" = '" & CreatedBy & "'")



            If chkRS.Fields.Item(0).Value = 0 Then
                objMain.objApplication.StatusBar.SetText("No staged data found for user " & CreatedBy, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If


            Try

                Dim BranchID As String = objForm.Items.Item("TNXBRC").Specific.value
                Dim query As String = "CALL ""TNX_ITEMDETAILS_NEW""('" & CreatedBy & "','" & BranchID & "')"

                oRecordset.DoQuery(query)
            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText("Error loading data call TNX_ITEMDETAIL_SREPORT(" & CreatedBy & ") : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End Try

            ' Try
            If oRecordset.RecordCount > 0 Then
                objMatrix3.Clear()
                objMatrix3.AddRow(oRecordset.RecordCount)
                Dim lineNum As Integer = 1
                While Not oRecordset.EoF
                    oDBs_Details3.SetValue("LineID", oDBs_Details3.Offset, lineNum)
                    oDBs_Details3.SetValue("U_BCode", oDBs_Details3.Offset, oRecordset.Fields.Item("BARCODE").Value)
                    oDBs_Details3.SetValue("U_ItemCode", oDBs_Details3.Offset, oRecordset.Fields.Item("ItemCode").Value)
                    oDBs_Details3.SetValue("U_IName", oDBs_Details3.Offset, oRecordset.Fields.Item("ItemName").Value)
                    oDBs_Details3.SetValue("U_UOM", oDBs_Details3.Offset, oRecordset.Fields.Item("UomEntry").Value)
                    oDBs_Details3.SetValue("U_CQty", oDBs_Details3.Offset, oRecordset.Fields.Item("Quantity").Value)
                    oDBs_Details3.SetValue("U_CCost", oDBs_Details3.Offset, oRecordset.Fields.Item("COST").Value)
                    oDBs_Details3.SetValue("U_CSPrice", oDBs_Details3.Offset, oRecordset.Fields.Item("RSP").Value)
                    oDBs_Details3.SetValue("U_CMarg", oDBs_Details3.Offset, oRecordset.Fields.Item("margin").Value)
                    oDBs_Details3.SetValue("U_NSPrice", oDBs_Details3.Offset, oRecordset.Fields.Item("NEWSELLINGPRICE").Value)
                    oDBs_Details3.SetValue("U_NMarg", oDBs_Details3.Offset, oRecordset.Fields.Item("NMARG").Value)
                    oDBs_Details3.SetValue("U_NCost", oDBs_Details3.Offset, oRecordset.Fields.Item("NCOST").Value)
                    oDBs_Details3.SetValue("U_LPhPrice", oDBs_Details3.Offset, oRecordset.Fields.Item("LPHPRICE").Value)
                    oDBs_Details3.SetValue("U_LGRPPrice", oDBs_Details3.Offset, oRecordset.Fields.Item("LGRPPRICE").Value)

                    objMatrix3.SetLineData(lineNum)
                    oRecordset.MoveNext()
                    lineNum += 1
                End While
                objForm.Freeze(False)
            End If
            objMain.objApplication.StatusBar.SetText("MOP Tab Matrix filled successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Matrix Import Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            objForm.Freeze(False)
        End Try

        GC.Collect()
    End Sub

    Private Sub UploadFromDataTable(dtNew As DataTable)

        '====================================================
        ' Variable declarations
        '====================================================
        Dim ItemCode As String = ""
        Dim Barcode As String = ""
        Dim Pricelist, nPrice, StartDate As String

        objMatrix3 = objForm.Items.Item("Mtx3").Specific
        objMatrix3.Clear()

        Dim oRS As SAPbobsCOM.Recordset =
        CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset),
              SAPbobsCOM.Recordset)

        Dim oRecordset As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim record As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        '====================================================
        ' Get Logged-in User
        '====================================================
        Dim userCode As String = objMain.objCompany.UserName.Replace("'", "''")
        oRS.DoQuery("SELECT ""USERID"" FROM OUSR WHERE ""USER_CODE"" = '" & userCode & "'")

        Dim CreatedBy As Integer = 0
        If Not oRS.EoF Then
            CreatedBy = CInt(oRS.Fields.Item("USERID").Value)
        End If

        '====================================================
        ' Validate Excel Data
        '====================================================
        If dtNew.Rows.Count = 0 OrElse dtNew.Columns.Count = 0 Then
            objMain.objApplication.StatusBar.SetText(
            "No data found in Excel sheet.",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Exit Sub
        End If

        '====================================================
        ' Detect column indexes
        '====================================================
        Dim idxItemCode As Integer = -1
        Dim idxBarcode As Integer = -1
        Dim idxNewSellingPrice As Integer = -1

        ' Header sniffing
        For c As Integer = 0 To dtNew.Columns.Count - 1
            Dim nm As String = dtNew.Columns(c).ColumnName.ToLower().Trim()

            If nm.Contains("barcode") OrElse nm.Contains("bar code") OrElse nm.Contains("bcd") OrElse nm.Contains("bcode") Then idxBarcode = c
            If nm.Contains("item") OrElse nm.Contains("itemcode") OrElse nm.Contains("item code") Then idxItemCode = c
            If nm.Contains("newsellingprice") OrElse nm.Contains("new selling") _
           OrElse nm.Contains("selling price") Then idxNewSellingPrice = c
        Next

        '====================================================
        ' Create staging table if not exists
        '====================================================
        Dim createTableSQL As String =
        "CREATE COLUMN TABLE ""TNX_PRICEUPDATE""(" &
        """BARCODE"" NVARCHAR(100)," &
        """NEWSELLINGPRICE"" DECIMAL(11,5) CS_FIXED," &
        """CREATEDBY"" NVARCHAR(100)," &
        """LINENUM"" INTEGER," &
        """UPDT"" NVARCHAR(100)) " &
        "UNLOAD PRIORITY 5 AUTO MERGE"

        Try
            oRS.DoQuery(createTableSQL)
        Catch ex As Exception
            If Not ex.Message.ToLower().Contains("already exists") AndAlso
           Not ex.Message.ToLower().Contains("duplicate") Then
                objMain.objApplication.StatusBar.SetText(
                "Error creating table: " & ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        End Try

        '====================================================
        ' Delete old staged data
        '====================================================
        oRS.DoQuery("DELETE FROM ""TNX_PRICEUPDATE"" WHERE ""CREATEDBY"" = '" & CreatedBy & "'")

        objMain.objApplication.StatusBar.SetText(
        "Inserting records into SAP HANA...",
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        '====================================================
        ' Bulk Insert Logic
        '====================================================
        Try
            Dim sb As New StringBuilder()
            Dim rowsAdded As Boolean = False
            Const MAX_BATCH_LENGTH As Integer = 800000

            For i As Integer = 0 To dtNew.Rows.Count - 1

                If dtNew.Rows(i).ItemArray.All(Function(x) x Is DBNull.Value OrElse String.IsNullOrWhiteSpace(x.ToString())) Then Continue For

                Dim rawItem As String = ""
                Dim rawBarcode As String = ""
                Dim rawNewSellingPrice As Decimal = 0

                ' ItemCode
                If idxItemCode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxItemCode)) Then
                    rawItem = dtNew.Rows(i)(idxItemCode).ToString().Trim()
                End If

                ' Barcode
                If idxBarcode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxBarcode)) Then
                    rawBarcode = dtNew.Rows(i)(idxBarcode).ToString().Trim()
                End If

                ' NEW SELLING PRICE from Excel
                If idxNewSellingPrice >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxNewSellingPrice)) Then
                    Decimal.TryParse(dtNew.Rows(i)(idxNewSellingPrice).ToString(), rawNewSellingPrice)
                End If

                ' Resolve Barcode using ItemCode
                If rawBarcode = "" AndAlso rawItem <> "" Then
                    Dim rsB As SAPbobsCOM.Recordset =
                    objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    rsB.DoQuery("SELECT ""BcdCode"" FROM OBCD WHERE ""ItemCode"" = '" &
                            rawItem.Replace("'", "''") & "' LIMIT 1")

                    If rsB.RecordCount > 0 Then
                        rawBarcode = rsB.Fields.Item(0).Value.ToString()
                    End If
                Else
                    rawBarcode = rawBarcode
                End If

                If rawBarcode = "" AndAlso rawItem = "" Then Continue For

                Dim lineNumber As Integer = i + 1

                If Not rowsAdded Then
                    sb.AppendLine(
                    "INSERT INTO ""TNX_PRICEUPDATE"" (" &
                    """BARCODE"",""NEWSELLINGPRICE"",""CREATEDBY"",""LINENUM"",""UPDT"")")
                    sb.Append("SELECT ")
                    rowsAdded = True
                Else
                    sb.AppendLine()
                    sb.Append("UNION ALL SELECT ")
                End If

                sb.AppendFormat(
                "'{0}',{1},'{2}',{3},CURRENT_DATE FROM DUMMY",
                rawBarcode.Replace("'", "''"),
                rawNewSellingPrice,
                CreatedBy,
                lineNumber)

                If sb.Length >= MAX_BATCH_LENGTH Then
                    oRS.DoQuery(sb.ToString())
                    sb.Clear()
                    rowsAdded = False
                End If

            Next

            If sb.Length > 0 Then
                oRS.DoQuery(sb.ToString())
            End If

            '====================================================
            ' Call stored procedure
            '====================================================
            Dim BranchID As String = objForm.Items.Item("TNXBRC").Specific.value
            oRecordset.DoQuery("CALL ""TNX_ITEMDETAILS_NEW""('" & CreatedBy & "','" & BranchID & "')")

            '====================================================
            ' Fill Temporary Matrix
            '====================================================
            If oRecordset.RecordCount > 0 Then

                objMatrix3.Clear()
                objMatrix3.AddRow(oRecordset.RecordCount)

                Dim lineNum As Integer = 1
                While Not oRecordset.EoF

                    oDBs_Details3.SetValue("LineID", oDBs_Details3.Offset, lineNum)
                    oDBs_Details3.SetValue("U_BCode", oDBs_Details3.Offset, oRecordset.Fields.Item("BARCODE").Value)
                    oDBs_Details3.SetValue("U_ItemCode", oDBs_Details3.Offset, oRecordset.Fields.Item("ItemCode").Value)
                    oDBs_Details3.SetValue("U_IName", oDBs_Details3.Offset, oRecordset.Fields.Item("ItemName").Value)
                    oDBs_Details3.SetValue("U_UOM", oDBs_Details3.Offset, oRecordset.Fields.Item("UomName").Value)
                    oDBs_Details3.SetValue("U_CQty", oDBs_Details3.Offset, oRecordset.Fields.Item("Quantity").Value)
                    oDBs_Details3.SetValue("U_CCost", oDBs_Details3.Offset, oRecordset.Fields.Item("LPCOSTNEW").Value)
                    oDBs_Details3.SetValue("U_CSPrice", oDBs_Details3.Offset, oRecordset.Fields.Item("RSP").Value)
                    oDBs_Details3.SetValue("U_CMarg", oDBs_Details3.Offset, oRecordset.Fields.Item("margin").Value)
                    oDBs_Details3.SetValue("U_NSPrice", oDBs_Details3.Offset, oRecordset.Fields.Item("NEWSELLINGPRICE").Value)
                    oDBs_Details3.SetValue("U_NMarg", oDBs_Details3.Offset, oRecordset.Fields.Item("NMARG").Value)
                    oDBs_Details3.SetValue("U_NCost", oDBs_Details3.Offset, oRecordset.Fields.Item("NCOST").Value)
                    oDBs_Details3.SetValue("U_LPhPrice", oDBs_Details3.Offset, oRecordset.Fields.Item("LPHPRICE").Value)
                    oDBs_Details3.SetValue("U_LGRPPrice", oDBs_Details3.Offset, oRecordset.Fields.Item("LGRPPRICE").Value)


                    objMatrix3.SetLineData(lineNum)
                    lineNum += 1
                    oRecordset.MoveNext()
                End While
                objForm.Freeze(False)
            End If

            'new uom logic based on item load uom list

            'If oRecordset.RecordCount > 0 Then

            '    objForm.Freeze(True)

            '    Try

            '        objMatrix3.Clear()
            '        oDBs_Details3.Clear()

            '        Dim lineNum As Integer = 0

            '        ' =====================================
            '        ' LOAD MATRIX DATA FROM RECORDSET
            '        ' =====================================
            '        While Not oRecordset.EoF

            '            oDBs_Details3.InsertRecord(lineNum)

            '            oDBs_Details3.SetValue("LineID", lineNum, lineNum + 1)
            '            oDBs_Details3.SetValue("U_BCode", lineNum, oRecordset.Fields.Item("BARCODE").Value.ToString())
            '            oDBs_Details3.SetValue("U_ItemCode", lineNum, oRecordset.Fields.Item("ItemCode").Value.ToString())
            '            oDBs_Details3.SetValue("U_IName", lineNum, oRecordset.Fields.Item("ItemName").Value.ToString())
            '            oDBs_Details3.SetValue("U_UOM", lineNum, oRecordset.Fields.Item("UomEntry").Value.ToString())

            '            oDBs_Details3.SetValue("U_CCost", lineNum, oRecordset.Fields.Item("COST").Value.ToString())
            '            oDBs_Details3.SetValue("U_CSPrice", lineNum, oRecordset.Fields.Item("RSP").Value.ToString())
            '            oDBs_Details3.SetValue("U_CMarg", lineNum, oRecordset.Fields.Item("margin").Value.ToString())
            '            oDBs_Details3.SetValue("U_NSPrice", lineNum, oRecordset.Fields.Item("NEWSELLINGPRICE").Value.ToString())
            '            oDBs_Details3.SetValue("U_NMarg", lineNum, oRecordset.Fields.Item("NMARG").Value.ToString())
            '            oDBs_Details3.SetValue("U_NCost", lineNum, oRecordset.Fields.Item("NCOST").Value.ToString())
            '            oDBs_Details3.SetValue("U_LPhPrice", lineNum, oRecordset.Fields.Item("LPHPRICE").Value.ToString())
            '            oDBs_Details3.SetValue("U_LGRPPrice", lineNum, oRecordset.Fields.Item("LGRPPRICE").Value.ToString())

            '            lineNum += 1
            '            oRecordset.MoveNext()

            '        End While

            '        objMatrix3.LoadFromDataSource()

            '        ' =====================================
            '        ' LOAD UOM PER ROW (SAFE WAY)
            '        ' =====================================

            '        Dim oColumn As SAPbouiCOM.Column = objMatrix3.Columns.Item("UOM")

            '        For rowIndex As Integer = 1 To objMatrix3.RowCount

            '            Dim currentItemCode As String =
            '    oDBs_Details3.GetValue("U_ItemCode", rowIndex - 1).Trim()

            '            If currentItemCode = "" Then Continue For

            '            ' 🔴 CLEAR COLUMN VALIDVALUES (NOT CELL)
            '            While oColumn.ValidValues.Count > 0
            '                oColumn.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
            '            End While

            '            Dim rsUOM As SAPbobsCOM.Recordset =
            '    objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            '            rsUOM.DoQuery("
            '    SELECT DISTINCT
            '        U.""UomEntry"",
            '        U.""UomCode""
            '    FROM ""OITM"" I
            '    INNER JOIN ""UGP1"" G ON I.""UgpEntry"" = G.""UgpEntry""
            '    INNER JOIN ""OUOM"" U ON G.""UomEntry"" = U.""UomEntry""
            '    WHERE I.""ItemCode"" = '" & currentItemCode & "'
            '    ORDER BY U.""UomCode""")

            '            While Not rsUOM.EoF

            '                oColumn.ValidValues.Add(
            '        rsUOM.Fields.Item("UomEntry").Value.ToString(),
            '        rsUOM.Fields.Item("UomCode").Value.ToString())

            '                rsUOM.MoveNext()
            '            End While

            '            ' Select stored UOM immediately
            '            Dim storedUOM As String =
            '    oDBs_Details3.GetValue("U_UOM", rowIndex - 1).Trim()

            '            If storedUOM <> "" Then

            '                Dim oCombo As SAPbouiCOM.ComboBox =
            '        CType(oColumn.Cells.Item(rowIndex).Specific, SAPbouiCOM.ComboBox)

            '                Try
            '                    oCombo.Select(storedUOM, SAPbouiCOM.BoSearchKey.psk_ByValue)
            '                Catch
            '                End Try

            '            End If

            '        Next

            '    Catch ex As Exception
            '        ' SBO_Application.MessageBox("Matrix Load Error: " & ex.Message)
            '    Finally
            '        objForm.Freeze(False)
            '    End Try

            'End If



            objMain.objApplication.StatusBar.SetText(
            "MOP Tab Matrix filled successfully!",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "Matrix Import Error: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            objForm.Freeze(False)
        End Try

        GC.Collect()

    End Sub

    Private Sub UploadFromDataTable1(dtNew As DataTable)

        '====================================================
        ' Variable declarations
        '====================================================
        Dim ItemCode As String = ""
        Dim Barcode As String = ""
        Dim Pricelist, nPrice, StartDate As String

        objMatrix3 = objForm.Items.Item("Mtx3").Specific
        'objMatrix3.Clear()

        Dim oRS As SAPbobsCOM.Recordset =
        CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset),
              SAPbobsCOM.Recordset)

        Dim oRecordset As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim record As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        '====================================================
        ' Get Logged-in User
        '====================================================
        Dim userCode As String = objMain.objCompany.UserName.Replace("'", "''")
        oRS.DoQuery("SELECT ""USERID"" FROM OUSR WHERE ""USER_CODE"" = '" & userCode & "'")

        Dim CreatedBy As Integer = 0
        If Not oRS.EoF Then
            CreatedBy = CInt(oRS.Fields.Item("USERID").Value)
        End If

        '====================================================
        ' Validate Excel Data
        '====================================================
        If dtNew.Rows.Count = 0 OrElse dtNew.Columns.Count = 0 Then
            objMain.objApplication.StatusBar.SetText(
            "No data found in Excel sheet.",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Exit Sub
        End If

        '====================================================
        ' Detect column indexes
        '====================================================
        Dim idxItemCode As Integer = -1
        Dim idxBarcode As Integer = -1
        Dim idxNewSellingPrice As Integer = -1

        ' Header sniffing
        For c As Integer = 0 To dtNew.Columns.Count - 1
            Dim nm As String = dtNew.Columns(c).ColumnName.ToLower().Trim()

            If nm.Contains("barcode") OrElse nm.Contains("bar code") OrElse nm.Contains("bcd") OrElse nm.Contains("bcode") Then idxBarcode = c
            If nm.Contains("item") OrElse nm.Contains("itemcode") OrElse nm.Contains("item code") Then idxItemCode = c
            If nm.Contains("newsellingprice") OrElse nm.Contains("new selling") _
           OrElse nm.Contains("selling price") Then idxNewSellingPrice = c
        Next

        '====================================================
        ' Create staging table if not exists
        '====================================================
        Dim createTableSQL As String =
        "CREATE COLUMN TABLE ""TNX_PRICEUPDATE""(" &
        """BARCODE"" NVARCHAR(100)," &
        """NEWSELLINGPRICE"" DECIMAL(11,5) CS_FIXED," &
        """CREATEDBY"" NVARCHAR(100)," &
        """LINENUM"" INTEGER," &
        """UPDT"" NVARCHAR(100)) " &
        "UNLOAD PRIORITY 5 AUTO MERGE"

        Try
            oRS.DoQuery(createTableSQL)
        Catch ex As Exception
            If Not ex.Message.ToLower().Contains("already exists") AndAlso
           Not ex.Message.ToLower().Contains("duplicate") Then
                objMain.objApplication.StatusBar.SetText(
                "Error creating table: " & ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        End Try

        '====================================================
        ' Delete old staged data
        '====================================================
        oRS.DoQuery("DELETE FROM ""TNX_PRICEUPDATE"" WHERE ""CREATEDBY"" = '" & CreatedBy & "'")

        objMain.objApplication.StatusBar.SetText(
        "Inserting records into SAP HANA...",
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        '====================================================
        ' Bulk Insert Logic
        '====================================================
        Try
            Dim sb As New StringBuilder()
            Dim rowsAdded As Boolean = False
            Const MAX_BATCH_LENGTH As Integer = 800000

            For i As Integer = 0 To dtNew.Rows.Count - 1

                If dtNew.Rows(i).ItemArray.All(Function(x) x Is DBNull.Value OrElse String.IsNullOrWhiteSpace(x.ToString())) Then Continue For

                Dim rawItem As String = ""
                Dim rawBarcode As String = ""
                Dim rawNewSellingPrice As Decimal = 0

                ' ItemCode
                If idxItemCode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxItemCode)) Then
                    rawItem = dtNew.Rows(i)(idxItemCode).ToString().Trim()
                End If

                ' Barcode
                If idxBarcode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxBarcode)) Then
                    rawBarcode = dtNew.Rows(i)(idxBarcode).ToString().Trim()
                End If

                ' NEW SELLING PRICE from Excel
                If idxNewSellingPrice >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxNewSellingPrice)) Then
                    Decimal.TryParse(dtNew.Rows(i)(idxNewSellingPrice).ToString(), rawNewSellingPrice)
                End If

                ' Resolve Barcode using ItemCode
                If rawBarcode = "" AndAlso rawItem <> "" Then
                    Dim rsB As SAPbobsCOM.Recordset =
                    objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    rsB.DoQuery("SELECT ""BcdCode"" FROM OBCD WHERE ""ItemCode"" = '" &
                            rawItem.Replace("'", "''") & "' LIMIT 1")

                    If rsB.RecordCount > 0 Then
                        rawBarcode = rsB.Fields.Item(0).Value.ToString()
                    End If
                Else
                    rawBarcode = rawBarcode
                End If

                If rawBarcode = "" AndAlso rawItem = "" Then Continue For

                Dim lineNumber As Integer = i + 1

                If Not rowsAdded Then
                    sb.AppendLine(
                    "INSERT INTO ""TNX_PRICEUPDATE"" (" &
                    """BARCODE"",""NEWSELLINGPRICE"",""CREATEDBY"",""LINENUM"",""UPDT"")")
                    sb.Append("SELECT ")
                    rowsAdded = True
                Else
                    sb.AppendLine()
                    sb.Append("UNION ALL SELECT ")
                End If

                sb.AppendFormat(
                "'{0}',{1},'{2}',{3},CURRENT_DATE FROM DUMMY",
                rawBarcode.Replace("'", "''"),
                rawNewSellingPrice,
                CreatedBy,
                lineNumber)

                If sb.Length >= MAX_BATCH_LENGTH Then
                    oRS.DoQuery(sb.ToString())
                    sb.Clear()
                    rowsAdded = False
                End If

            Next

            If sb.Length > 0 Then
                oRS.DoQuery(sb.ToString())
            End If

            Dim BranchID As String = objForm.Items.Item("TNXBRC").Specific.Value
            Dim oRecordsetTemp As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


            Dim ExistingLineIDs As New List(Of Integer)
            Dim DocEntry As String = oDBs_Head.GetValue("DocEntry", 0).Trim()
            oRecordsetTemp.DoQuery("SELECT ""LineId"" FROM ""@TNX_PCRT_C2"" WHERE ""DocEntry"" = '" & DocEntry & "' and IFNULL(""U_BCode"",'') <> '' ORDER BY ""LineId""")

            While Not oRecordsetTemp.EoF
                ExistingLineIDs.Add(CInt(oRecordsetTemp.Fields.Item("LineId").Value))
                oRecordsetTemp.MoveNext()
            End While


            oRecordset.DoQuery("CALL ""TNX_ITEMDETAILS_NEW""('" & CreatedBy & "','" & BranchID & "')")

            If oRecordset.RecordCount > 0 Then
                oRecordset.MoveFirst()

                While Not oRecordset.EoF

                    Dim ProcedureLineID As Integer = CInt(oRecordset.Fields.Item("LINENUM").Value)


                    If Not ExistingLineIDs.Contains(ProcedureLineID) Then


                        Dim TargetRow As Integer = ProcedureLineID - 1

                        oDBs_Details3.Offset = TargetRow

                        oDBs_Details3.SetValue("LineId", TargetRow, ProcedureLineID)
                        oDBs_Details3.SetValue("U_BCode", TargetRow, oRecordset.Fields.Item("BARCODE").Value)
                        oDBs_Details3.SetValue("U_ItemCode", TargetRow, oRecordset.Fields.Item("ItemCode").Value)
                        oDBs_Details3.SetValue("U_IName", TargetRow, oRecordset.Fields.Item("ItemName").Value)
                        oDBs_Details3.SetValue("U_UOM", TargetRow, oRecordset.Fields.Item("UomName").Value)
                        oDBs_Details3.SetValue("U_CQty", TargetRow, oRecordset.Fields.Item("Quantity").Value)
                        oDBs_Details3.SetValue("U_CCost", TargetRow, oRecordset.Fields.Item("LPCOSTNEW").Value)
                        oDBs_Details3.SetValue("U_CSPrice", TargetRow, oRecordset.Fields.Item("RSP").Value)
                        oDBs_Details3.SetValue("U_CMarg", TargetRow, oRecordset.Fields.Item("margin").Value)
                        oDBs_Details3.SetValue("U_NSPrice", TargetRow, oRecordset.Fields.Item("NEWSELLINGPRICE").Value)
                        oDBs_Details3.SetValue("U_NMarg", TargetRow, oRecordset.Fields.Item("NMARG").Value)
                        oDBs_Details3.SetValue("U_NCost", TargetRow, oRecordset.Fields.Item("NCOST").Value)
                        oDBs_Details3.SetValue("U_LPhPrice", TargetRow, oRecordset.Fields.Item("LPHPRICE").Value)
                        oDBs_Details3.SetValue("U_LGRPPrice", TargetRow, oRecordset.Fields.Item("LGRPPRICE").Value)


                        objMatrix3.SetLineData(ProcedureLineID)

                    End If

                    oRecordset.MoveNext()
                End While
            End If

            For i As Integer = objMatrix3.RowCount To 1 Step -1
                If String.IsNullOrWhiteSpace(objMatrix3.Columns.Item("BCode").Cells.Item(i).Specific.Value.ToString().Trim()) Then
                    objMatrix3.DeleteRow(i)
                End If
            Next

            objForm.Freeze(False)




            objMain.objApplication.StatusBar.SetText(
            "MOP Tab Matrix filled successfully!",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "Matrix Import Error: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            objForm.Freeze(False)
        End Try

        GC.Collect()

    End Sub

    'copy and  paste in EDSP tab code here

    Private Sub UploadFromDataTableEDSP(dtNew As DataTable)

        '====================================================
        ' Variable declarations
        '====================================================
        Dim ItemCode As String = ""
        Dim Barcode As String = ""
        Dim Pricelist, nPrice, StartDate As String

        objMatrix6 = objForm.Items.Item("Mtx6").Specific
        objMatrix6.Clear()

        Dim oRS As SAPbobsCOM.Recordset =
        CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset),
              SAPbobsCOM.Recordset)

        Dim oRecordset As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim record As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        '====================================================
        ' Get Logged-in User
        '====================================================
        Dim userCode As String = objMain.objCompany.UserName.Replace("'", "''")
        oRS.DoQuery("SELECT ""USERID"" FROM OUSR WHERE ""USER_CODE"" = '" & userCode & "'")

        Dim CreatedBy As Integer = 0
        If Not oRS.EoF Then
            CreatedBy = CInt(oRS.Fields.Item("USERID").Value)
        End If

        '====================================================
        ' Validate Excel Data
        '====================================================
        If dtNew.Rows.Count = 0 OrElse dtNew.Columns.Count = 0 Then
            objMain.objApplication.StatusBar.SetText(
            "No data found in Excel sheet.",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Exit Sub
        End If

        '====================================================
        ' Detect column indexes
        '====================================================
        Dim idxItemCode As Integer = -1
        Dim idxBarcode As Integer = -1
        Dim idxNewSellingPrice As Integer = -1

        ' Header sniffing
        For c As Integer = 0 To dtNew.Columns.Count - 1
            Dim nm As String = dtNew.Columns(c).ColumnName.ToLower().Trim()

            If nm.Contains("barcode") OrElse nm.Contains("bar code") OrElse nm.Contains("bcd") OrElse nm.Contains("bcode") Then idxBarcode = c
            If nm.Contains("item") OrElse nm.Contains("itemcode") OrElse nm.Contains("item code") Then idxItemCode = c
            If nm.Contains("newsellingprice") OrElse nm.Contains("new selling") _
           OrElse nm.Contains("selling price") Then idxNewSellingPrice = c
        Next

        '====================================================
        ' Create staging table if not exists
        '====================================================
        Dim createTableSQL As String =
        "CREATE COLUMN TABLE ""TNX_PRICEUPDATE""(" &
        """BARCODE"" NVARCHAR(100)," &
        """NEWSELLINGPRICE"" DECIMAL(11,5) CS_FIXED," &
        """CREATEDBY"" NVARCHAR(100)," &
        """LINENUM"" INTEGER," &
        """UPDT"" NVARCHAR(100)) " &
        "UNLOAD PRIORITY 5 AUTO MERGE"

        Try
            oRS.DoQuery(createTableSQL)
        Catch ex As Exception
            If Not ex.Message.ToLower().Contains("already exists") AndAlso
           Not ex.Message.ToLower().Contains("duplicate") Then
                objMain.objApplication.StatusBar.SetText(
                "Error creating table: " & ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        End Try

        '====================================================
        ' Delete old staged data
        '====================================================
        oRS.DoQuery("DELETE FROM ""TNX_PRICEUPDATE"" WHERE ""CREATEDBY"" = '" & CreatedBy & "'")

        objMain.objApplication.StatusBar.SetText(
        "Inserting records into SAP HANA...",
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        '====================================================
        ' Bulk Insert Logic
        '====================================================
        Try
            Dim sb As New StringBuilder()
            Dim rowsAdded As Boolean = False
            Const MAX_BATCH_LENGTH As Integer = 800000

            For i As Integer = 0 To dtNew.Rows.Count - 1

                If dtNew.Rows(i).ItemArray.All(Function(x) x Is DBNull.Value OrElse String.IsNullOrWhiteSpace(x.ToString())) Then Continue For

                Dim rawItem As String = ""
                Dim rawBarcode As String = ""
                Dim rawNewSellingPrice As Decimal = 0

                ' ItemCode
                If idxItemCode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxItemCode)) Then
                    rawItem = dtNew.Rows(i)(idxItemCode).ToString().Trim()
                End If

                ' Barcode
                If idxBarcode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxBarcode)) Then
                    rawBarcode = dtNew.Rows(i)(idxBarcode).ToString().Trim()
                End If

                ' NEW SELLING PRICE from Excel
                If idxNewSellingPrice >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxNewSellingPrice)) Then
                    Decimal.TryParse(dtNew.Rows(i)(idxNewSellingPrice).ToString(), rawNewSellingPrice)
                End If

                ' Resolve Barcode using ItemCode
                If rawBarcode = "" AndAlso rawItem <> "" Then
                    Dim rsB As SAPbobsCOM.Recordset =
                    objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    rsB.DoQuery("SELECT ""BcdCode"" FROM OBCD WHERE ""ItemCode"" = '" &
                            rawItem.Replace("'", "''") & "' LIMIT 1")

                    If rsB.RecordCount > 0 Then
                        rawBarcode = rsB.Fields.Item(0).Value.ToString()
                    End If
                Else
                    rawBarcode = rawBarcode
                End If

                If rawBarcode = "" AndAlso rawItem = "" Then Continue For

                Dim lineNumber As Integer = i + 1

                If Not rowsAdded Then
                    sb.AppendLine(
                    "INSERT INTO ""TNX_PRICEUPDATE"" (" &
                    """BARCODE"",""NEWSELLINGPRICE"",""CREATEDBY"",""LINENUM"",""UPDT"")")
                    sb.Append("SELECT ")
                    rowsAdded = True
                Else
                    sb.AppendLine()
                    sb.Append("UNION ALL SELECT ")
                End If

                sb.AppendFormat(
                "'{0}',{1},'{2}',{3},CURRENT_DATE FROM DUMMY",
                rawBarcode.Replace("'", "''"),
                rawNewSellingPrice,
                CreatedBy,
                lineNumber)

                If sb.Length >= MAX_BATCH_LENGTH Then
                    oRS.DoQuery(sb.ToString())
                    sb.Clear()
                    rowsAdded = False
                End If

            Next

            If sb.Length > 0 Then
                oRS.DoQuery(sb.ToString())
            End If

            '====================================================
            ' Call stored procedure
            '====================================================
            Dim BranchID As String = objForm.Items.Item("TNXBRC").Specific.value
            oRecordset.DoQuery("CALL ""TNX_ITEMDETAILS_NEW""('" & CreatedBy & "','" & BranchID & "')")

            '====================================================
            ' Fill Temporary Matrix
            '====================================================
            If oRecordset.RecordCount > 0 Then

                objMatrix6.Clear()
                objMatrix6.AddRow(oRecordset.RecordCount)

                Dim lineNum As Integer = 1
                While Not oRecordset.EoF

                    oDBs_Details6.SetValue("LineID", oDBs_Details6.Offset, lineNum)
                    oDBs_Details6.SetValue("U_BCode", oDBs_Details6.Offset, oRecordset.Fields.Item("BARCODE").Value)
                    oDBs_Details6.SetValue("U_ICode", oDBs_Details6.Offset, oRecordset.Fields.Item("ItemCode").Value)
                    oDBs_Details6.SetValue("U_IName", oDBs_Details6.Offset, oRecordset.Fields.Item("ItemName").Value)
                    oDBs_Details6.SetValue("U_CQTT", oDBs_Details6.Offset, oRecordset.Fields.Item("Quantity").Value)
                    oDBs_Details6.SetValue("U_Unit", oDBs_Details6.Offset, oRecordset.Fields.Item("UomName").Value)
                    oDBs_Details6.SetValue("U_CCost", oDBs_Details6.Offset, oRecordset.Fields.Item("LPCOSTNEW").Value)
                    oDBs_Details6.SetValue("U_CSlng", oDBs_Details6.Offset, oRecordset.Fields.Item("RSP").Value)
                    oDBs_Details6.SetValue("U_CMarg", oDBs_Details6.Offset, oRecordset.Fields.Item("margin").Value)
                    oDBs_Details6.SetValue("U_NSlng", oDBs_Details6.Offset, oRecordset.Fields.Item("NEWSELLINGPRICE").Value)
                    oDBs_Details6.SetValue("U_NMrgn", oDBs_Details6.Offset, oRecordset.Fields.Item("NMARG").Value)
                    ''oDBs_Details6.SetValue("U_NCost", oDBs_Details6.Offset, oRecordset.Fields.Item("LGRPPRICE").Value)



                    objMatrix6.SetLineData(lineNum)
                    lineNum += 1
                    oRecordset.MoveNext()
                End While
                objForm.Freeze(False)
            End If





            objMain.objApplication.StatusBar.SetText(
            "Every day Special Price Tab Matrix filled successfully!",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "Matrix Import Error: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            objForm.Freeze(False)
        End Try

        GC.Collect()

    End Sub

    Private Sub UploadFromDataTableEDSP1(dtNew As DataTable)

        '====================================================
        ' Variable declarations
        '====================================================
        Dim ItemCode As String = ""
        Dim Barcode As String = ""
        Dim Pricelist, nPrice, StartDate As String

        objMatrix6 = objForm.Items.Item("Mtx6").Specific
        ''objMatrix6.Clear()

        Dim oRS As SAPbobsCOM.Recordset =
        CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset),
              SAPbobsCOM.Recordset)

        Dim oRecordset As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim record As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        '====================================================
        ' Get Logged-in User
        '====================================================
        Dim userCode As String = objMain.objCompany.UserName.Replace("'", "''")
        oRS.DoQuery("SELECT ""USERID"" FROM OUSR WHERE ""USER_CODE"" = '" & userCode & "'")

        Dim CreatedBy As Integer = 0
        If Not oRS.EoF Then
            CreatedBy = CInt(oRS.Fields.Item("USERID").Value)
        End If

        '====================================================
        ' Validate Excel Data
        '====================================================
        If dtNew.Rows.Count = 0 OrElse dtNew.Columns.Count = 0 Then
            objMain.objApplication.StatusBar.SetText(
            "No data found in Excel sheet.",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Exit Sub
        End If

        '====================================================
        ' Detect column indexes
        '====================================================
        Dim idxItemCode As Integer = -1
        Dim idxBarcode As Integer = -1
        Dim idxNewSellingPrice As Integer = -1

        ' Header sniffing
        For c As Integer = 0 To dtNew.Columns.Count - 1
            Dim nm As String = dtNew.Columns(c).ColumnName.ToLower().Trim()

            If nm.Contains("barcode") OrElse nm.Contains("bar code") OrElse nm.Contains("bcd") OrElse nm.Contains("bcode") Then idxBarcode = c
            If nm.Contains("item") OrElse nm.Contains("itemcode") OrElse nm.Contains("item code") Then idxItemCode = c
            If nm.Contains("newsellingprice") OrElse nm.Contains("new selling") _
           OrElse nm.Contains("selling price") Then idxNewSellingPrice = c
        Next

        '====================================================
        ' Create staging table if not exists
        '====================================================
        Dim createTableSQL As String =
        "CREATE COLUMN TABLE ""TNX_PRICEUPDATE""(" &
        """BARCODE"" NVARCHAR(100)," &
        """NEWSELLINGPRICE"" DECIMAL(11,5) CS_FIXED," &
        """CREATEDBY"" NVARCHAR(100)," &
        """LINENUM"" INTEGER," &
        """UPDT"" NVARCHAR(100)) " &
        "UNLOAD PRIORITY 5 AUTO MERGE"

        Try
            oRS.DoQuery(createTableSQL)
        Catch ex As Exception
            If Not ex.Message.ToLower().Contains("already exists") AndAlso
           Not ex.Message.ToLower().Contains("duplicate") Then
                objMain.objApplication.StatusBar.SetText(
                "Error creating table: " & ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        End Try

        '====================================================
        ' Delete old staged data
        '====================================================
        oRS.DoQuery("DELETE FROM ""TNX_PRICEUPDATE"" WHERE ""CREATEDBY"" = '" & CreatedBy & "'")

        objMain.objApplication.StatusBar.SetText(
        "Inserting records into SAP HANA...",
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        '====================================================
        ' Bulk Insert Logic
        '====================================================
        Try
            Dim sb As New StringBuilder()
            Dim rowsAdded As Boolean = False
            Const MAX_BATCH_LENGTH As Integer = 800000

            For i As Integer = 0 To dtNew.Rows.Count - 1

                If dtNew.Rows(i).ItemArray.All(Function(x) x Is DBNull.Value OrElse String.IsNullOrWhiteSpace(x.ToString())) Then Continue For

                Dim rawItem As String = ""
                Dim rawBarcode As String = ""
                Dim rawNewSellingPrice As Decimal = 0

                ' ItemCode
                If idxItemCode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxItemCode)) Then
                    rawItem = dtNew.Rows(i)(idxItemCode).ToString().Trim()
                End If

                ' Barcode
                If idxBarcode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxBarcode)) Then
                    rawBarcode = dtNew.Rows(i)(idxBarcode).ToString().Trim()
                End If

                ' NEW SELLING PRICE from Excel
                If idxNewSellingPrice >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxNewSellingPrice)) Then
                    Decimal.TryParse(dtNew.Rows(i)(idxNewSellingPrice).ToString(), rawNewSellingPrice)
                End If

                ' Resolve Barcode using ItemCode
                If rawBarcode = "" AndAlso rawItem <> "" Then
                    Dim rsB As SAPbobsCOM.Recordset =
                    objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    rsB.DoQuery("SELECT ""BcdCode"" FROM OBCD WHERE ""ItemCode"" = '" &
                            rawItem.Replace("'", "''") & "' LIMIT 1")

                    If rsB.RecordCount > 0 Then
                        rawBarcode = rsB.Fields.Item(0).Value.ToString()
                    End If
                Else
                    rawBarcode = rawBarcode
                End If

                If rawBarcode = "" AndAlso rawItem = "" Then Continue For

                Dim lineNumber As Integer = i + 1

                If Not rowsAdded Then
                    sb.AppendLine(
                    "INSERT INTO ""TNX_PRICEUPDATE"" (" &
                    """BARCODE"",""NEWSELLINGPRICE"",""CREATEDBY"",""LINENUM"",""UPDT"")")
                    sb.Append("SELECT ")
                    rowsAdded = True
                Else
                    sb.AppendLine()
                    sb.Append("UNION ALL SELECT ")
                End If

                sb.AppendFormat(
                "'{0}',{1},'{2}',{3},CURRENT_DATE FROM DUMMY",
                rawBarcode.Replace("'", "''"),
                rawNewSellingPrice,
                CreatedBy,
                lineNumber)

                If sb.Length >= MAX_BATCH_LENGTH Then
                    oRS.DoQuery(sb.ToString())
                    sb.Clear()
                    rowsAdded = False
                End If

            Next

            If sb.Length > 0 Then
                oRS.DoQuery(sb.ToString())
            End If



            Dim BranchID As String = objForm.Items.Item("TNXBRC").Specific.Value
            Dim oRecordsetTemp As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


            Dim ExistingLineIDs As New List(Of Integer)
            Dim DocEntry As String = oDBs_Head.GetValue("DocEntry", 0).Trim()
            oRecordsetTemp.DoQuery("SELECT ""LineId"" FROM ""@TNX_PCRT_C5"" WHERE ""DocEntry"" = '" & DocEntry & "' and IFNULL(""U_BCode"",'') <> '' ORDER BY ""LineId""")

            While Not oRecordsetTemp.EoF
                ExistingLineIDs.Add(CInt(oRecordsetTemp.Fields.Item("LineId").Value))
                oRecordsetTemp.MoveNext()
            End While


            oRecordset.DoQuery("CALL ""TNX_ITEMDETAILS_NEW""('" & CreatedBy & "','" & BranchID & "')")

            If oRecordset.RecordCount > 0 Then
                oRecordset.MoveFirst()

                While Not oRecordset.EoF

                    Dim ProcedureLineID As Integer = CInt(oRecordset.Fields.Item("LINENUM").Value)


                    If Not ExistingLineIDs.Contains(ProcedureLineID) Then


                        Dim TargetRow As Integer = ProcedureLineID - 1

                        oDBs_Details6.Offset = TargetRow

                        oDBs_Details6.SetValue("LineID", TargetRow, ProcedureLineID)
                        oDBs_Details6.SetValue("U_BCode", TargetRow, oRecordset.Fields.Item("BARCODE").Value)
                        oDBs_Details6.SetValue("U_ICode", TargetRow, oRecordset.Fields.Item("ItemCode").Value)
                        oDBs_Details6.SetValue("U_IName", TargetRow, oRecordset.Fields.Item("ItemName").Value)
                        oDBs_Details6.SetValue("U_CQTT", TargetRow, oRecordset.Fields.Item("Quantity").Value)
                        oDBs_Details6.SetValue("U_Unit", TargetRow, oRecordset.Fields.Item("UomName").Value)
                        oDBs_Details6.SetValue("U_CCost", TargetRow, oRecordset.Fields.Item("LPCOSTNEW").Value)
                        oDBs_Details6.SetValue("U_CSlng", TargetRow, oRecordset.Fields.Item("RSP").Value)
                        oDBs_Details6.SetValue("U_CMarg", TargetRow, oRecordset.Fields.Item("margin").Value)
                        oDBs_Details6.SetValue("U_NSlng", TargetRow, oRecordset.Fields.Item("NEWSELLINGPRICE").Value)
                        oDBs_Details6.SetValue("U_NMrgn", TargetRow, oRecordset.Fields.Item("NMARG").Value)


                        objMatrix6.SetLineData(ProcedureLineID)

                    End If

                    oRecordset.MoveNext()
                End While
            End If
            For i As Integer = objMatrix6.RowCount To 1 Step -1
                If String.IsNullOrWhiteSpace(objMatrix6.Columns.Item("BCode").Cells.Item(i).Specific.Value.ToString().Trim()) Then
                    objMatrix6.DeleteRow(i)
                End If
            Next
            objForm.Freeze(False)





            objMain.objApplication.StatusBar.SetText(
            "Every day Special Price Tab Matrix filled successfully!",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "Matrix Import Error: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            objForm.Freeze(False)
        End Try

        GC.Collect()

    End Sub

    'copy and  paste in BCU tab code here

    Private Sub UploadFromDataTableBCU(dtNew As DataTable)

        '====================================================
        ' Variable declarations
        '====================================================
        Dim ItemCode As String = ""
        Dim Barcode As String = ""
        Dim Pricelist, nPrice, StartDate As String

        objMatrix7 = objForm.Items.Item("Mtx7").Specific
        objMatrix7.Clear()

        Dim oRS As SAPbobsCOM.Recordset =
        CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset),
              SAPbobsCOM.Recordset)

        Dim oRecordset As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim record As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        '====================================================
        ' Get Logged-in User
        '====================================================
        Dim userCode As String = objMain.objCompany.UserName.Replace("'", "''")
        oRS.DoQuery("SELECT ""USERID"" FROM OUSR WHERE ""USER_CODE"" = '" & userCode & "'")

        Dim CreatedBy As Integer = 0
        If Not oRS.EoF Then
            CreatedBy = CInt(oRS.Fields.Item("USERID").Value)
        End If

        '====================================================
        ' Validate Excel Data
        '====================================================
        If dtNew.Rows.Count = 0 OrElse dtNew.Columns.Count = 0 Then
            objMain.objApplication.StatusBar.SetText(
            "No data found in Excel sheet.",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Exit Sub
        End If

        '====================================================
        ' Detect column indexes
        '====================================================
        Dim idxItemCode As Integer = -1
        Dim idxBarcode As Integer = -1
        Dim idxNewSellingPrice As Integer = -1

        ' Header sniffing
        For c As Integer = 0 To dtNew.Columns.Count - 1
            Dim nm As String = dtNew.Columns(c).ColumnName.ToLower().Trim()

            If nm.Contains("barcode") OrElse nm.Contains("bar code") OrElse nm.Contains("bcd") OrElse nm.Contains("bcode") Then idxBarcode = c
            If nm.Contains("item") OrElse nm.Contains("itemcode") OrElse nm.Contains("item code") Then idxItemCode = c
            If nm.Contains("newsellingprice") OrElse nm.Contains("new selling") _
           OrElse nm.Contains("selling price") Then idxNewSellingPrice = c
        Next

        '====================================================
        ' Create staging table if not exists
        '====================================================
        Dim createTableSQL As String =
        "CREATE COLUMN TABLE ""TNX_PRICEUPDATE""(" &
        """BARCODE"" NVARCHAR(100)," &
        """NEWSELLINGPRICE"" DECIMAL(11,5) CS_FIXED," &
        """CREATEDBY"" NVARCHAR(100)," &
        """LINENUM"" INTEGER," &
        """UPDT"" NVARCHAR(100)) " &
        "UNLOAD PRIORITY 5 AUTO MERGE"

        Try
            oRS.DoQuery(createTableSQL)
        Catch ex As Exception
            If Not ex.Message.ToLower().Contains("already exists") AndAlso
           Not ex.Message.ToLower().Contains("duplicate") Then
                objMain.objApplication.StatusBar.SetText(
                "Error creating table: " & ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        End Try

        '====================================================
        ' Delete old staged data
        '====================================================
        oRS.DoQuery("DELETE FROM ""TNX_PRICEUPDATE"" WHERE ""CREATEDBY"" = '" & CreatedBy & "'")

        objMain.objApplication.StatusBar.SetText(
        "Inserting records into SAP HANA...",
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        '====================================================
        ' Bulk Insert Logic
        '====================================================
        Try
            Dim sb As New StringBuilder()
            Dim rowsAdded As Boolean = False
            Const MAX_BATCH_LENGTH As Integer = 800000

            For i As Integer = 0 To dtNew.Rows.Count - 1

                If dtNew.Rows(i).ItemArray.All(Function(x) x Is DBNull.Value OrElse String.IsNullOrWhiteSpace(x.ToString())) Then Continue For

                Dim rawItem As String = ""
                Dim rawBarcode As String = ""
                Dim rawNewSellingPrice As Decimal = 0

                ' ItemCode
                If idxItemCode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxItemCode)) Then
                    rawItem = dtNew.Rows(i)(idxItemCode).ToString().Trim()
                End If

                ' Barcode
                If idxBarcode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxBarcode)) Then
                    rawBarcode = dtNew.Rows(i)(idxBarcode).ToString().Trim()
                End If

                ' NEW SELLING PRICE from Excel
                If idxNewSellingPrice >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxNewSellingPrice)) Then
                    Decimal.TryParse(dtNew.Rows(i)(idxNewSellingPrice).ToString(), rawNewSellingPrice)
                End If

                ' Resolve Barcode using ItemCode
                If rawBarcode = "" AndAlso rawItem <> "" Then
                    Dim rsB As SAPbobsCOM.Recordset =
                    objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    rsB.DoQuery("SELECT ""BcdCode"" FROM OBCD WHERE ""ItemCode"" = '" &
                            rawItem.Replace("'", "''") & "' LIMIT 1")

                    If rsB.RecordCount > 0 Then
                        rawBarcode = rsB.Fields.Item(0).Value.ToString()
                    End If
                Else
                    rawBarcode = rawBarcode
                End If

                If rawBarcode = "" AndAlso rawItem = "" Then Continue For

                Dim lineNumber As Integer = i + 1

                If Not rowsAdded Then
                    sb.AppendLine(
                    "INSERT INTO ""TNX_PRICEUPDATE"" (" &
                    """BARCODE"",""NEWSELLINGPRICE"",""CREATEDBY"",""LINENUM"",""UPDT"")")
                    sb.Append("SELECT ")
                    rowsAdded = True
                Else
                    sb.AppendLine()
                    sb.Append("UNION ALL SELECT ")
                End If

                sb.AppendFormat(
                "'{0}',{1},'{2}',{3},CURRENT_DATE FROM DUMMY",
                rawBarcode.Replace("'", "''"),
                rawNewSellingPrice,
                CreatedBy,
                lineNumber)

                If sb.Length >= MAX_BATCH_LENGTH Then
                    oRS.DoQuery(sb.ToString())
                    sb.Clear()
                    rowsAdded = False
                End If

            Next

            If sb.Length > 0 Then
                oRS.DoQuery(sb.ToString())
            End If

            '====================================================
            ' Call stored procedure
            '====================================================
            Dim BranchID As String = objForm.Items.Item("TNXBRC").Specific.value
            oRecordset.DoQuery("CALL ""TNX_ITEMDETAILS_NEW""('" & CreatedBy & "','" & BranchID & "')")

            '====================================================
            ' Fill Temporary Matrix
            '====================================================
            If oRecordset.RecordCount > 0 Then

                objMatrix7.Clear()
                objMatrix7.AddRow(oRecordset.RecordCount)

                Dim lineNum As Integer = 1
                While Not oRecordset.EoF

                    oDBs_Details7.SetValue("LineID", oDBs_Details7.Offset, lineNum)
                    oDBs_Details7.SetValue("U_BCode", oDBs_Details7.Offset, oRecordset.Fields.Item("BARCODE").Value)
                    oDBs_Details7.SetValue("U_ITCD", oDBs_Details7.Offset, oRecordset.Fields.Item("ItemCode").Value)
                    oDBs_Details7.SetValue("U_ITNM", oDBs_Details7.Offset, oRecordset.Fields.Item("ItemName").Value)
                    oDBs_Details7.SetValue("U_UOM", oDBs_Details7.Offset, oRecordset.Fields.Item("UomName").Value)
                    oDBs_Details7.SetValue("U_NPRC", oDBs_Details7.Offset, oRecordset.Fields.Item("NEWSELLINGPRICE").Value)




                    objMatrix7.SetLineData(lineNum)
                    lineNum += 1
                    oRecordset.MoveNext()
                End While
                objForm.Freeze(False)
            End If





            objMain.objApplication.StatusBar.SetText(
            "Bulk Cost upload Tab Matrix filled successfully!",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "Matrix Import Error: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            objForm.Freeze(False)
        End Try

        GC.Collect()

    End Sub

    'copy and  paste in BPU tab code here

    Private Sub UploadFromDataTableBPU(dtNew As DataTable)

        '====================================================
        ' Variable declarations
        '====================================================
        Dim ItemCode As String = ""
        Dim Barcode As String = ""
        Dim Pricelist, nPrice, StartDate As String

        objMatrix9 = objForm.Items.Item("Mtx9").Specific
        objMatrix9.Clear()

        Dim oRS As SAPbobsCOM.Recordset =
        CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset),
              SAPbobsCOM.Recordset)

        Dim oRecordset As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim record As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        '====================================================
        ' Get Logged-in User
        '====================================================
        Dim userCode As String = objMain.objCompany.UserName.Replace("'", "''")
        oRS.DoQuery("SELECT ""USERID"" FROM OUSR WHERE ""USER_CODE"" = '" & userCode & "'")

        Dim CreatedBy As Integer = 0
        If Not oRS.EoF Then
            CreatedBy = CInt(oRS.Fields.Item("USERID").Value)
        End If

        '====================================================
        ' Validate Excel Data
        '====================================================
        If dtNew.Rows.Count = 0 OrElse dtNew.Columns.Count = 0 Then
            objMain.objApplication.StatusBar.SetText(
            "No data found in Excel sheet.",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Exit Sub
        End If

        '====================================================
        ' Detect column indexes
        '====================================================
        Dim idxItemCode As Integer = -1
        Dim idxBarcode As Integer = -1
        Dim idxNewSellingPrice As Integer = -1

        ' Header sniffing
        For c As Integer = 0 To dtNew.Columns.Count - 1
            Dim nm As String = dtNew.Columns(c).ColumnName.ToLower().Trim()

            If nm.Contains("barcode") OrElse nm.Contains("bar code") OrElse nm.Contains("bcd") OrElse nm.Contains("bcode") Then idxBarcode = c
            If nm.Contains("item") OrElse nm.Contains("itemcode") OrElse nm.Contains("item code") Then idxItemCode = c
            If nm.Contains("newsellingprice") OrElse nm.Contains("new selling") _
           OrElse nm.Contains("selling price") Then idxNewSellingPrice = c
        Next

        '====================================================
        ' Create staging table if not exists
        '====================================================
        Dim createTableSQL As String =
        "CREATE COLUMN TABLE ""TNX_PRICEUPDATE""(" &
        """BARCODE"" NVARCHAR(100)," &
        """NEWSELLINGPRICE"" DECIMAL(11,5) CS_FIXED," &
        """CREATEDBY"" NVARCHAR(100)," &
        """LINENUM"" INTEGER," &
        """UPDT"" NVARCHAR(100)) " &
        "UNLOAD PRIORITY 5 AUTO MERGE"

        Try
            oRS.DoQuery(createTableSQL)
        Catch ex As Exception
            If Not ex.Message.ToLower().Contains("already exists") AndAlso
           Not ex.Message.ToLower().Contains("duplicate") Then
                objMain.objApplication.StatusBar.SetText(
                "Error creating table: " & ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        End Try

        '====================================================
        ' Delete old staged data
        '====================================================
        oRS.DoQuery("DELETE FROM ""TNX_PRICEUPDATE"" WHERE ""CREATEDBY"" = '" & CreatedBy & "'")

        objMain.objApplication.StatusBar.SetText(
        "Inserting records into SAP HANA...",
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        '====================================================
        ' Bulk Insert Logic
        '====================================================
        Try
            Dim sb As New StringBuilder()
            Dim rowsAdded As Boolean = False
            Const MAX_BATCH_LENGTH As Integer = 800000

            For i As Integer = 0 To dtNew.Rows.Count - 1

                If dtNew.Rows(i).ItemArray.All(Function(x) x Is DBNull.Value OrElse String.IsNullOrWhiteSpace(x.ToString())) Then Continue For

                Dim rawItem As String = ""
                Dim rawBarcode As String = ""
                Dim rawNewSellingPrice As Decimal = 0

                ' Barcode
                If idxBarcode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxBarcode)) Then
                    rawBarcode = dtNew.Rows(i)(idxBarcode).ToString().Trim()
                End If
                ' ItemCode
                If idxItemCode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxItemCode)) Then
                    rawItem = dtNew.Rows(i)(idxItemCode).ToString().Trim()
                End If




                ' NEW SELLING PRICE from Excel
                If idxNewSellingPrice >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxNewSellingPrice)) Then
                    Decimal.TryParse(dtNew.Rows(i)(idxNewSellingPrice).ToString(), rawNewSellingPrice)
                End If

                ' Resolve Barcode using ItemCode
                If rawBarcode = "" AndAlso rawItem <> "" Then
                    Dim rsB As SAPbobsCOM.Recordset =
                    objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    rsB.DoQuery("SELECT ""BcdCode"" FROM OBCD WHERE ""ItemCode"" = '" &
                            rawItem.Replace("'", "''") & "' LIMIT 1")

                    If rsB.RecordCount > 0 Then
                        rawBarcode = rsB.Fields.Item(0).Value.ToString()
                    End If
                Else
                    rawBarcode = rawBarcode
                End If

                If rawBarcode = "" AndAlso rawItem = "" Then Continue For

                Dim lineNumber As Integer = i + 1

                If Not rowsAdded Then
                    sb.AppendLine(
                    "INSERT INTO ""TNX_PRICEUPDATE"" (" &
                    """BARCODE"",""NEWSELLINGPRICE"",""CREATEDBY"",""LINENUM"",""UPDT"")")
                    sb.Append("SELECT ")
                    rowsAdded = True
                Else
                    sb.AppendLine()
                    sb.Append("UNION ALL SELECT ")
                End If

                sb.AppendFormat(
                "'{0}',{1},'{2}',{3},CURRENT_DATE FROM DUMMY",
                rawBarcode.Replace("'", "''"),
                rawNewSellingPrice,
                CreatedBy,
                lineNumber)

                If sb.Length >= MAX_BATCH_LENGTH Then
                    oRS.DoQuery(sb.ToString())
                    sb.Clear()
                    rowsAdded = False
                End If

            Next

            If sb.Length > 0 Then
                oRS.DoQuery(sb.ToString())
            End If

            '====================================================
            ' Call stored procedure
            '====================================================
            Dim BranchID As String = objForm.Items.Item("TNXBRC").Specific.value
            oRecordset.DoQuery("CALL ""TNX_ITEMDETAILS_NEW""('" & CreatedBy & "','" & BranchID & "')")

            '====================================================
            ' Fill Temporary Matrix
            '====================================================
            If oRecordset.RecordCount > 0 Then

                objMatrix9.Clear()
                objMatrix9.AddRow(oRecordset.RecordCount)

                Dim lineNum As Integer = 1
                While Not oRecordset.EoF

                    oDBs_Details9.SetValue("LineID", oDBs_Details9.Offset, lineNum)
                    oDBs_Details9.SetValue("U_BCode", oDBs_Details9.Offset, oRecordset.Fields.Item("BARCODE").Value)
                    oDBs_Details9.SetValue("U_ITCD", oDBs_Details9.Offset, oRecordset.Fields.Item("ItemCode").Value)
                    oDBs_Details9.SetValue("U_BName", oDBs_Details9.Offset, oRecordset.Fields.Item("ItemName").Value)
                    oDBs_Details9.SetValue("U_UOM", oDBs_Details9.Offset, oRecordset.Fields.Item("UomName").Value)
                    oDBs_Details9.SetValue("U_OSLP", oDBs_Details9.Offset, oRecordset.Fields.Item("RSP").Value)




                    objMatrix9.SetLineData(lineNum)
                    lineNum += 1
                    oRecordset.MoveNext()
                End While
                objForm.Freeze(False)
            End If





            objMain.objApplication.StatusBar.SetText(
            "Bulk Price upload Tab Matrix filled successfully!",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "Matrix Import Error: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            objForm.Freeze(False)
        End Try

        GC.Collect()

    End Sub

    Private Sub UploadFromDataTableBPU1(dtNew As DataTable)

        '====================================================
        ' Variable declarations
        '====================================================
        Dim ItemCode As String = ""
        Dim Barcode As String = ""
        Dim Pricelist, nPrice, StartDate As String

        objMatrix9 = objForm.Items.Item("Mtx9").Specific
        'objMatrix7.Clear()

        Dim oRS As SAPbobsCOM.Recordset =
        CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset),
              SAPbobsCOM.Recordset)

        Dim oRecordset As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim record As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        '====================================================
        ' Get Logged-in User
        '====================================================
        Dim userCode As String = objMain.objCompany.UserName.Replace("'", "''")
        oRS.DoQuery("SELECT ""USERID"" FROM OUSR WHERE ""USER_CODE"" = '" & userCode & "'")

        Dim CreatedBy As Integer = 0
        If Not oRS.EoF Then
            CreatedBy = CInt(oRS.Fields.Item("USERID").Value)
        End If

        '====================================================
        ' Validate Excel Data
        '====================================================
        If dtNew.Rows.Count = 0 OrElse dtNew.Columns.Count = 0 Then
            objMain.objApplication.StatusBar.SetText(
            "No data found in Excel sheet.",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Exit Sub
        End If

        '====================================================
        ' Detect column indexes
        '====================================================
        Dim idxItemCode As Integer = -1
        Dim idxBarcode As Integer = -1
        Dim idxNewSellingPrice As Integer = -1

        ' Header sniffing
        For c As Integer = 0 To dtNew.Columns.Count - 1
            Dim nm As String = dtNew.Columns(c).ColumnName.ToLower().Trim()

            If nm.Contains("barcode") OrElse nm.Contains("bar code") OrElse nm.Contains("bcd") OrElse nm.Contains("bcode") Then idxBarcode = c
            If nm.Contains("item") OrElse nm.Contains("itemcode") OrElse nm.Contains("item code") Then idxItemCode = c
            If nm.Contains("newsellingprice") OrElse nm.Contains("new selling") _
           OrElse nm.Contains("selling price") Then idxNewSellingPrice = c
        Next

        '====================================================
        ' Create staging table if not exists
        '====================================================
        Dim createTableSQL As String =
        "CREATE COLUMN TABLE ""TNX_PRICEUPDATE""(" &
        """BARCODE"" NVARCHAR(100)," &
        """NEWSELLINGPRICE"" DECIMAL(11,5) CS_FIXED," &
        """CREATEDBY"" NVARCHAR(100)," &
        """LINENUM"" INTEGER," &
        """UPDT"" NVARCHAR(100)) " &
        "UNLOAD PRIORITY 5 AUTO MERGE"

        Try
            oRS.DoQuery(createTableSQL)
        Catch ex As Exception
            If Not ex.Message.ToLower().Contains("already exists") AndAlso
           Not ex.Message.ToLower().Contains("duplicate") Then
                objMain.objApplication.StatusBar.SetText(
                "Error creating table: " & ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        End Try

        '====================================================
        ' Delete old staged data
        '====================================================
        oRS.DoQuery("DELETE FROM ""TNX_PRICEUPDATE"" WHERE ""CREATEDBY"" = '" & CreatedBy & "'")

        objMain.objApplication.StatusBar.SetText(
        "Inserting records into SAP HANA...",
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        '====================================================
        ' Bulk Insert Logic
        '====================================================
        Try
            Dim sb As New StringBuilder()
            Dim rowsAdded As Boolean = False
            Const MAX_BATCH_LENGTH As Integer = 800000

            For i As Integer = 0 To dtNew.Rows.Count - 1

                If dtNew.Rows(i).ItemArray.All(Function(x) x Is DBNull.Value OrElse String.IsNullOrWhiteSpace(x.ToString())) Then Continue For

                Dim rawItem As String = ""
                Dim rawBarcode As String = ""
                Dim rawNewSellingPrice As Decimal = 0

                ' ItemCode
                If idxItemCode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxItemCode)) Then
                    rawItem = dtNew.Rows(i)(idxItemCode).ToString().Trim()
                End If

                ' Barcode
                If idxBarcode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxBarcode)) Then
                    rawBarcode = dtNew.Rows(i)(idxBarcode).ToString().Trim()
                End If

                ' NEW SELLING PRICE from Excel
                If idxNewSellingPrice >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxNewSellingPrice)) Then
                    Decimal.TryParse(dtNew.Rows(i)(idxNewSellingPrice).ToString(), rawNewSellingPrice)
                End If

                ' Resolve Barcode using ItemCode
                If rawBarcode = "" AndAlso rawItem <> "" Then
                    Dim rsB As SAPbobsCOM.Recordset =
                    objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    rsB.DoQuery("SELECT ""BcdCode"" FROM OBCD WHERE ""ItemCode"" = '" &
                            rawItem.Replace("'", "''") & "' LIMIT 1")

                    If rsB.RecordCount > 0 Then
                        rawBarcode = rsB.Fields.Item(0).Value.ToString()
                    End If
                Else
                    rawBarcode = rawBarcode
                End If

                If rawBarcode = "" AndAlso rawItem = "" Then Continue For

                Dim lineNumber As Integer = i + 1

                If Not rowsAdded Then
                    sb.AppendLine(
                    "INSERT INTO ""TNX_PRICEUPDATE"" (" &
                    """BARCODE"",""NEWSELLINGPRICE"",""CREATEDBY"",""LINENUM"",""UPDT"")")
                    sb.Append("SELECT ")
                    rowsAdded = True
                Else
                    sb.AppendLine()
                    sb.Append("UNION ALL SELECT ")
                End If

                sb.AppendFormat(
                "'{0}',{1},'{2}',{3},CURRENT_DATE FROM DUMMY",
                rawBarcode.Replace("'", "''"),
                rawNewSellingPrice,
                CreatedBy,
                lineNumber)

                If sb.Length >= MAX_BATCH_LENGTH Then
                    oRS.DoQuery(sb.ToString())
                    sb.Clear()
                    rowsAdded = False
                End If

            Next

            If sb.Length > 0 Then
                oRS.DoQuery(sb.ToString())
            End If


            Dim BranchID As String = objForm.Items.Item("TNXBRC").Specific.Value
            Dim oRecordsetTemp As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


            Dim ExistingLineIDs As New List(Of Integer)
            Dim DocEntry As String = oDBs_Head.GetValue("DocEntry", 0).Trim()
            oRecordsetTemp.DoQuery("SELECT ""LineId"" FROM ""@TNX_PCRT_C8"" WHERE ""DocEntry"" = '" & DocEntry & "' and IFNULL(""U_BCode"",'') <> '' ORDER BY ""LineId""")

            While Not oRecordsetTemp.EoF
                ExistingLineIDs.Add(CInt(oRecordsetTemp.Fields.Item("LineId").Value))
                oRecordsetTemp.MoveNext()
            End While


            oRecordset.DoQuery("CALL ""TNX_ITEMDETAILS_NEW""('" & CreatedBy & "','" & BranchID & "')")

            If oRecordset.RecordCount > 0 Then
                oRecordset.MoveFirst()

                While Not oRecordset.EoF

                    Dim ProcedureLineID As Integer = CInt(oRecordset.Fields.Item("LINENUM").Value)


                    If Not ExistingLineIDs.Contains(ProcedureLineID) Then


                        Dim TargetRow As Integer = ProcedureLineID - 1

                        oDBs_Details9.Offset = TargetRow

                        oDBs_Details9.SetValue("LineID", TargetRow, ProcedureLineID)
                        oDBs_Details9.SetValue("U_BCode", TargetRow, oRecordset.Fields.Item("BARCODE").Value)
                        oDBs_Details9.SetValue("U_ITCD", TargetRow, oRecordset.Fields.Item("ItemCode").Value)
                        oDBs_Details9.SetValue("U_BName", TargetRow, oRecordset.Fields.Item("ItemName").Value)
                        oDBs_Details9.SetValue("U_UOM", TargetRow, oRecordset.Fields.Item("UomName").Value)
                        oDBs_Details9.SetValue("U_OSLP", TargetRow, oRecordset.Fields.Item("RSP").Value)


                        objMatrix9.SetLineData(ProcedureLineID)

                    End If
                    For i As Integer = objMatrix9.RowCount To 1 Step -1
                        If String.IsNullOrWhiteSpace(objMatrix9.Columns.Item("Col_0").Cells.Item(i).Specific.Value.ToString().Trim()) Then
                            objMatrix9.DeleteRow(i)
                        End If
                    Next
                    oRecordset.MoveNext()
                End While
            End If

            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(
            "Bulk Cost upload Tab Matrix filled successfully!",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "Matrix Import Error: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            objForm.Freeze(False)
        End Try

        GC.Collect()

    End Sub

    Private Sub UploadFromDataTableBCU1(dtNew As DataTable)

        '====================================================
        ' Variable declarations
        '====================================================
        Dim ItemCode As String = ""
        Dim Barcode As String = ""
        Dim Pricelist, nPrice, StartDate As String

        objMatrix7 = objForm.Items.Item("Mtx7").Specific
        'objMatrix7.Clear()

        Dim oRS As SAPbobsCOM.Recordset =
        CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset),
              SAPbobsCOM.Recordset)

        Dim oRecordset As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim record As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        '====================================================
        ' Get Logged-in User
        '====================================================
        Dim userCode As String = objMain.objCompany.UserName.Replace("'", "''")
        oRS.DoQuery("SELECT ""USERID"" FROM OUSR WHERE ""USER_CODE"" = '" & userCode & "'")

        Dim CreatedBy As Integer = 0
        If Not oRS.EoF Then
            CreatedBy = CInt(oRS.Fields.Item("USERID").Value)
        End If

        '====================================================
        ' Validate Excel Data
        '====================================================
        If dtNew.Rows.Count = 0 OrElse dtNew.Columns.Count = 0 Then
            objMain.objApplication.StatusBar.SetText(
            "No data found in Excel sheet.",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Exit Sub
        End If

        '====================================================
        ' Detect column indexes
        '====================================================
        Dim idxItemCode As Integer = -1
        Dim idxBarcode As Integer = -1
        Dim idxNewSellingPrice As Integer = -1

        ' Header sniffing
        For c As Integer = 0 To dtNew.Columns.Count - 1
            Dim nm As String = dtNew.Columns(c).ColumnName.ToLower().Trim()

            If nm.Contains("barcode") OrElse nm.Contains("bar code") OrElse nm.Contains("bcd") OrElse nm.Contains("bcode") Then idxBarcode = c
            If nm.Contains("item") OrElse nm.Contains("itemcode") OrElse nm.Contains("item code") Then idxItemCode = c
            If nm.Contains("newsellingprice") OrElse nm.Contains("new selling") _
           OrElse nm.Contains("selling price") Then idxNewSellingPrice = c
        Next

        '====================================================
        ' Create staging table if not exists
        '====================================================
        Dim createTableSQL As String =
        "CREATE COLUMN TABLE ""TNX_PRICEUPDATE""(" &
        """BARCODE"" NVARCHAR(100)," &
        """NEWSELLINGPRICE"" DECIMAL(11,5) CS_FIXED," &
        """CREATEDBY"" NVARCHAR(100)," &
        """LINENUM"" INTEGER," &
        """UPDT"" NVARCHAR(100)) " &
        "UNLOAD PRIORITY 5 AUTO MERGE"

        Try
            oRS.DoQuery(createTableSQL)
        Catch ex As Exception
            If Not ex.Message.ToLower().Contains("already exists") AndAlso
           Not ex.Message.ToLower().Contains("duplicate") Then
                objMain.objApplication.StatusBar.SetText(
                "Error creating table: " & ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        End Try

        '====================================================
        ' Delete old staged data
        '====================================================
        oRS.DoQuery("DELETE FROM ""TNX_PRICEUPDATE"" WHERE ""CREATEDBY"" = '" & CreatedBy & "'")

        objMain.objApplication.StatusBar.SetText(
        "Inserting records into SAP HANA...",
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        '====================================================
        ' Bulk Insert Logic
        '====================================================
        Try
            Dim sb As New StringBuilder()
            Dim rowsAdded As Boolean = False
            Const MAX_BATCH_LENGTH As Integer = 800000

            For i As Integer = 0 To dtNew.Rows.Count - 1

                If dtNew.Rows(i).ItemArray.All(Function(x) x Is DBNull.Value OrElse String.IsNullOrWhiteSpace(x.ToString())) Then Continue For

                Dim rawItem As String = ""
                Dim rawBarcode As String = ""
                Dim rawNewSellingPrice As Decimal = 0

                ' ItemCode
                If idxItemCode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxItemCode)) Then
                    rawItem = dtNew.Rows(i)(idxItemCode).ToString().Trim()
                End If

                ' Barcode
                If idxBarcode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxBarcode)) Then
                    rawBarcode = dtNew.Rows(i)(idxBarcode).ToString().Trim()
                End If

                ' NEW SELLING PRICE from Excel
                If idxNewSellingPrice >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxNewSellingPrice)) Then
                    Decimal.TryParse(dtNew.Rows(i)(idxNewSellingPrice).ToString(), rawNewSellingPrice)
                End If

                ' Resolve Barcode using ItemCode
                If rawBarcode = "" AndAlso rawItem <> "" Then
                    Dim rsB As SAPbobsCOM.Recordset =
                    objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    rsB.DoQuery("SELECT ""BcdCode"" FROM OBCD WHERE ""ItemCode"" = '" &
                            rawItem.Replace("'", "''") & "' LIMIT 1")

                    If rsB.RecordCount > 0 Then
                        rawBarcode = rsB.Fields.Item(0).Value.ToString()
                    End If
                Else
                    rawBarcode = rawBarcode
                End If

                If rawBarcode = "" AndAlso rawItem = "" Then Continue For

                Dim lineNumber As Integer = i + 1

                If Not rowsAdded Then
                    sb.AppendLine(
                    "INSERT INTO ""TNX_PRICEUPDATE"" (" &
                    """BARCODE"",""NEWSELLINGPRICE"",""CREATEDBY"",""LINENUM"",""UPDT"")")
                    sb.Append("SELECT ")
                    rowsAdded = True
                Else
                    sb.AppendLine()
                    sb.Append("UNION ALL SELECT ")
                End If

                sb.AppendFormat(
                "'{0}',{1},'{2}',{3},CURRENT_DATE FROM DUMMY",
                rawBarcode.Replace("'", "''"),
                rawNewSellingPrice,
                CreatedBy,
                lineNumber)

                If sb.Length >= MAX_BATCH_LENGTH Then
                    oRS.DoQuery(sb.ToString())
                    sb.Clear()
                    rowsAdded = False
                End If

            Next

            If sb.Length > 0 Then
                oRS.DoQuery(sb.ToString())
            End If


            Dim BranchID As String = objForm.Items.Item("TNXBRC").Specific.Value
            Dim oRecordsetTemp As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


            Dim ExistingLineIDs As New List(Of Integer)
            Dim DocEntry As String = oDBs_Head.GetValue("DocEntry", 0).Trim()
            oRecordsetTemp.DoQuery("SELECT ""LineId"" FROM ""@TNX_PCRT_C6"" WHERE ""DocEntry"" = '" & DocEntry & "' and IFNULL(""U_BCode"",'') <> '' ORDER BY ""LineId""")

            While Not oRecordsetTemp.EoF
                ExistingLineIDs.Add(CInt(oRecordsetTemp.Fields.Item("LineId").Value))
                oRecordsetTemp.MoveNext()
            End While


            oRecordset.DoQuery("CALL ""TNX_ITEMDETAILS_NEW""('" & CreatedBy & "','" & BranchID & "')")

            If oRecordset.RecordCount > 0 Then
                oRecordset.MoveFirst()

                While Not oRecordset.EoF

                    Dim ProcedureLineID As Integer = CInt(oRecordset.Fields.Item("LINENUM").Value)


                    If Not ExistingLineIDs.Contains(ProcedureLineID) Then


                        Dim TargetRow As Integer = ProcedureLineID - 1

                        oDBs_Details7.Offset = TargetRow

                        oDBs_Details7.SetValue("LineID", TargetRow, ProcedureLineID)
                        oDBs_Details7.SetValue("U_BCode", TargetRow, oRecordset.Fields.Item("BARCODE").Value)
                        oDBs_Details7.SetValue("U_ITCD", TargetRow, oRecordset.Fields.Item("ItemCode").Value)
                        oDBs_Details7.SetValue("U_ITNM", TargetRow, oRecordset.Fields.Item("ItemName").Value)
                        oDBs_Details7.SetValue("U_UOM", TargetRow, oRecordset.Fields.Item("UomName").Value)
                        oDBs_Details7.SetValue("U_NPRC", TargetRow, oRecordset.Fields.Item("NEWSELLINGPRICE").Value)


                        objMatrix7.SetLineData(ProcedureLineID)

                    End If
                    For i As Integer = objMatrix7.RowCount To 1 Step -1
                        If String.IsNullOrWhiteSpace(objMatrix7.Columns.Item("BCode").Cells.Item(i).Specific.Value.ToString().Trim()) Then
                            objMatrix7.DeleteRow(i)
                        End If
                    Next
                    oRecordset.MoveNext()
                End While
            End If

            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(
            "Bulk Cost upload Tab Matrix filled successfully!",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "Matrix Import Error: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            objForm.Freeze(False)
        End Try

        GC.Collect()

    End Sub

    'copy and  paste in Flyer tab code here
    Private Sub UploadFromDataTableFlyerTab(dtNew As DataTable)

        '====================================================
        ' Variable declarations
        '====================================================
        Dim ItemCode As String = ""
        Dim Barcode As String = ""
        Dim Pricelist, nPrice, StartDate As String

        objMatrix2 = objForm.Items.Item("Mtx2").Specific
        objMatrix2.Clear()

        Dim oRS As SAPbobsCOM.Recordset =
        CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset),
              SAPbobsCOM.Recordset)

        Dim oRecordset As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim record As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        '====================================================
        ' Get Logged-in SAP User
        '====================================================
        Dim userCode As String = objMain.objCompany.UserName.Replace("'", "''")
        oRS.DoQuery("SELECT ""USERID"" FROM OUSR WHERE ""USER_CODE"" = '" & userCode & "'")

        Dim CreatedBy As Integer = 0
        If Not oRS.EoF Then
            CreatedBy = CInt(oRS.Fields.Item("USERID").Value)
        End If

        '====================================================
        ' Validate Excel data
        '====================================================
        If dtNew.Rows.Count = 0 OrElse dtNew.Columns.Count = 0 Then
            objMain.objApplication.StatusBar.SetText(
            "No data found in Excel sheet.",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Exit Sub
        End If

        '====================================================
        ' Detect column indexes (ItemCode / Barcode / NewSellingPrice)
        '====================================================
        Dim idxItemCode As Integer = -1
        Dim idxBarcode As Integer = -1
        Dim idxNewSellingPrice As Integer = -1

        ' Header sniffing
        For c As Integer = 0 To dtNew.Columns.Count - 1
            Dim nm As String = dtNew.Columns(c).ColumnName.ToLower().Trim()

            If nm.Contains("barcode") OrElse nm.Contains("bar code") OrElse nm.Contains("bcd") OrElse nm.Contains("bcode") Then idxBarcode = c
            If nm.Contains("item") OrElse nm.Contains("itemcode") OrElse nm.Contains("item code") Then idxItemCode = c
            If nm.Contains("newsellingprice") OrElse nm.Contains("new selling") _
           OrElse nm.Contains("selling price") Then idxNewSellingPrice = c
        Next

        '====================================================
        ' Create staging table if not exists
        '====================================================
        Dim createTableSQL As String =
        "CREATE COLUMN TABLE ""TNX_PRICEUPDATE""(" &
        """BARCODE"" NVARCHAR(100)," &
        """NEWSELLINGPRICE"" DECIMAL(11,5) CS_FIXED," &
        """CREATEDBY"" NVARCHAR(100)," &
        """LINENUM"" INTEGER," &
        """UPDT"" NVARCHAR(100)) " &
        "UNLOAD PRIORITY 5 AUTO MERGE"

        Try
            oRS.DoQuery(createTableSQL)
        Catch ex As Exception
            If Not ex.Message.ToLower().Contains("already exists") AndAlso
           Not ex.Message.ToLower().Contains("duplicate") Then
                objMain.objApplication.StatusBar.SetText(
                "Error creating table: " & ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        End Try

        '====================================================
        ' Delete previous user data
        '====================================================
        oRS.DoQuery("DELETE FROM ""TNX_PRICEUPDATE"" WHERE ""CREATEDBY"" = '" & CreatedBy & "'")

        objMain.objApplication.StatusBar.SetText(
        "Inserting records into SAP HANA...",
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        '====================================================
        ' Bulk Insert Logic
        '====================================================
        Try
            Dim sb As New StringBuilder()
            Dim rowsAdded As Boolean = False
            Const MAX_BATCH_LENGTH As Integer = 800000

            For i As Integer = 0 To dtNew.Rows.Count - 1

                If dtNew.Rows(i).ItemArray.All(Function(x) x Is DBNull.Value OrElse String.IsNullOrWhiteSpace(x.ToString())) Then Continue For

                Dim rawItem As String = ""
                Dim rawBarcode As String = ""
                Dim rawNewSellingPrice As Decimal = 0

                ' ItemCode
                If idxItemCode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxItemCode)) Then
                    rawItem = dtNew.Rows(i)(idxItemCode).ToString().Trim()
                End If

                ' Barcode
                If idxBarcode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxBarcode)) Then
                    rawBarcode = dtNew.Rows(i)(idxBarcode).ToString().Trim()
                End If

                ' NEW SELLING PRICE from Excel
                If idxNewSellingPrice >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxNewSellingPrice)) Then
                    Decimal.TryParse(dtNew.Rows(i)(idxNewSellingPrice).ToString(), rawNewSellingPrice)
                End If

                ' Resolve barcode from ItemCode
                If rawBarcode = "" AndAlso rawItem <> "" Then
                    Dim rsB As SAPbobsCOM.Recordset =
                    objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    rsB.DoQuery("SELECT ""BcdCode"" FROM OBCD WHERE ""ItemCode"" = '" &
                            rawItem.Replace("'", "''") & "' LIMIT 1")

                    If rsB.RecordCount > 0 Then
                        rawBarcode = rsB.Fields.Item(0).Value.ToString()
                    End If
                End If

                If rawBarcode = "" AndAlso rawItem = "" Then Continue For

                Dim lineNumber As Integer = i + 1

                If Not rowsAdded Then
                    sb.AppendLine(
                    "INSERT INTO ""TNX_PRICEUPDATE"" (" &
                    """BARCODE"",""NEWSELLINGPRICE"",""CREATEDBY"",""LINENUM"",""UPDT"")")
                    sb.Append("SELECT ")
                    rowsAdded = True
                Else
                    sb.AppendLine()
                    sb.Append("UNION ALL SELECT ")
                End If

                sb.AppendFormat(
                "'{0}',{1},'{2}',{3},CURRENT_DATE FROM DUMMY",
                rawBarcode.Replace("'", "''"),
                rawNewSellingPrice,
                CreatedBy,
                lineNumber)

                If sb.Length >= MAX_BATCH_LENGTH Then
                    oRS.DoQuery(sb.ToString())
                    sb.Clear()
                    rowsAdded = False
                End If

            Next

            If sb.Length > 0 Then
                oRS.DoQuery(sb.ToString())
            End If

            '====================================================
            ' Call stored procedure
            '====================================================
            Dim BranchID As String = objForm.Items.Item("TNXBRC").Specific.value
            oRecordset.DoQuery("CALL ""TNX_ITEMDETAILS_NEW""('" & CreatedBy & "','" & BranchID & "')")

            '====================================================
            ' Load Flyer Matrix
            '====================================================
            If oRecordset.RecordCount > 0 Then
                objMatrix2.Clear()
                objMatrix2.AddRow(oRecordset.RecordCount)

                Dim lineNum As Integer = 1
                While Not oRecordset.EoF

                    oDBs_Details2.SetValue("LineID", oDBs_Details2.Offset, lineNum)
                    oDBs_Details2.SetValue("U_BrCode", oDBs_Details2.Offset, oRecordset.Fields.Item("BARCODE").Value)
                    oDBs_Details2.SetValue("U_ItemCode", oDBs_Details2.Offset, oRecordset.Fields.Item("ItemCode").Value)
                    oDBs_Details2.SetValue("U_ITNM", oDBs_Details2.Offset, oRecordset.Fields.Item("ItemName").Value)
                    oDBs_Details2.SetValue("U_ITGP", oDBs_Details2.Offset, oRecordset.Fields.Item("ItemGroup").Value)
                    oDBs_Details2.SetValue("U_CPrice", oDBs_Details2.Offset, oRecordset.Fields.Item("RSP").Value)
                    oDBs_Details2.SetValue("U_NSNG", oDBs_Details2.Offset, oRecordset.Fields.Item("RSP").Value)
                    oDBs_Details2.SetValue("U_UOM", oDBs_Details2.Offset, oRecordset.Fields.Item("UomName").Value)
                    oDBs_Details2.SetValue("U_Cost", oDBs_Details2.Offset, oRecordset.Fields.Item("LPCOSTNEW").Value) ''oRecordset.Fields.Item("COST").Value)
                    oDBs_Details2.SetValue("U_LPrice", oDBs_Details2.Offset, oRecordset.Fields.Item("LPHPRICE").Value)
                    oDBs_Details2.SetValue("U_NwPrice", oDBs_Details2.Offset, oRecordset.Fields.Item("NEWSELLINGPRICE").Value)
                    Dim Cost As Double = CDbl(oRecordset.Fields.Item("LPCOSTNEW").Value)
                    Dim NSNG As Double = CDbl(oRecordset.Fields.Item("RSP").Value)

                    Dim NMGP As Double = 0
                    If NSNG <> 0 Then
                        NMGP = ((NSNG - Cost) / NSNG) * 100
                    End If

                    oDBs_Details2.SetValue("U_NMGP", oDBs_Details2.Offset, NMGP.ToString("0.00"))
                    ' existing commented code – NOT removed
                    'oDBs_Details2.SetValue("U_Marg", oDBs_Details2.Offset, oRecordset.Fields.Item("margin").Value)

                    objMatrix2.SetLineData(lineNum)
                    lineNum += 1
                    oRecordset.MoveNext()
                End While
                objForm.Freeze(False)

            End If
            'CalculateGPTotals(objForm)
            SetMatrixTotals(objForm)
            objMain.objApplication.StatusBar.SetText(
            "Flyer Tab Matrix filled successfully!",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "Matrix Import Error: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            objForm.Freeze(False)
        End Try

        GC.Collect()

    End Sub

    Private Sub UploadFromDataTableFlyerTab1(dtNew As DataTable)

        '====================================================
        ' Variable declarations
        '====================================================
        Dim ItemCode As String = ""
        Dim Barcode As String = ""
        Dim Pricelist, nPrice, StartDate As String

        objMatrix2 = objForm.Items.Item("Mtx2").Specific
        ''objMatrix2.Clear()

        Dim oRS As SAPbobsCOM.Recordset =
        CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset),
              SAPbobsCOM.Recordset)

        Dim oRecordset As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim record As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        '====================================================
        ' Get Logged-in SAP User
        '====================================================
        Dim userCode As String = objMain.objCompany.UserName.Replace("'", "''")
        oRS.DoQuery("SELECT ""USERID"" FROM OUSR WHERE ""USER_CODE"" = '" & userCode & "'")

        Dim CreatedBy As Integer = 0
        If Not oRS.EoF Then
            CreatedBy = CInt(oRS.Fields.Item("USERID").Value)
        End If

        '====================================================
        ' Validate Excel data
        '====================================================
        If dtNew.Rows.Count = 0 OrElse dtNew.Columns.Count = 0 Then
            objMain.objApplication.StatusBar.SetText(
            "No data found in Excel sheet.",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Exit Sub
        End If

        '====================================================
        ' Detect column indexes (ItemCode / Barcode / NewSellingPrice)
        '====================================================
        Dim idxItemCode As Integer = -1
        Dim idxBarcode As Integer = -1
        Dim idxNewSellingPrice As Integer = -1

        ' Header sniffing
        For c As Integer = 0 To dtNew.Columns.Count - 1
            Dim nm As String = dtNew.Columns(c).ColumnName.ToLower().Trim()

            If nm.Contains("barcode") OrElse nm.Contains("bar code") OrElse nm.Contains("bcd") OrElse nm.Contains("bcode") Then idxBarcode = c
            If nm.Contains("item") OrElse nm.Contains("itemcode") OrElse nm.Contains("item code") Then idxItemCode = c
            If nm.Contains("newsellingprice") OrElse nm.Contains("new selling") _
           OrElse nm.Contains("selling price") Then idxNewSellingPrice = c
        Next

        '====================================================
        ' Create staging table if not exists
        '====================================================
        Dim createTableSQL As String =
        "CREATE COLUMN TABLE ""TNX_PRICEUPDATE""(" &
        """BARCODE"" NVARCHAR(100)," &
        """NEWSELLINGPRICE"" DECIMAL(11,5) CS_FIXED," &
        """CREATEDBY"" NVARCHAR(100)," &
        """LINENUM"" INTEGER," &
        """UPDT"" NVARCHAR(100)) " &
        "UNLOAD PRIORITY 5 AUTO MERGE"

        Try
            oRS.DoQuery(createTableSQL)
        Catch ex As Exception
            If Not ex.Message.ToLower().Contains("already exists") AndAlso
           Not ex.Message.ToLower().Contains("duplicate") Then
                objMain.objApplication.StatusBar.SetText(
                "Error creating table: " & ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        End Try

        '====================================================
        ' Delete previous user data
        '====================================================
        oRS.DoQuery("DELETE FROM ""TNX_PRICEUPDATE"" WHERE ""CREATEDBY"" = '" & CreatedBy & "'")

        objMain.objApplication.StatusBar.SetText(
        "Inserting records into SAP HANA...",
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        '====================================================
        ' Bulk Insert Logic
        '====================================================
        Try
            Dim sb As New StringBuilder()
            Dim rowsAdded As Boolean = False
            Const MAX_BATCH_LENGTH As Integer = 800000

            For i As Integer = 0 To dtNew.Rows.Count - 1

                If dtNew.Rows(i).ItemArray.All(Function(x) x Is DBNull.Value OrElse String.IsNullOrWhiteSpace(x.ToString())) Then Continue For

                Dim rawItem As String = ""
                Dim rawBarcode As String = ""
                Dim rawNewSellingPrice As Decimal = 0

                ' ItemCode
                If idxItemCode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxItemCode)) Then
                    rawItem = dtNew.Rows(i)(idxItemCode).ToString().Trim()
                End If

                ' Barcode
                If idxBarcode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxBarcode)) Then
                    rawBarcode = dtNew.Rows(i)(idxBarcode).ToString().Trim()
                End If

                ' NEW SELLING PRICE from Excel
                If idxNewSellingPrice >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxNewSellingPrice)) Then
                    Decimal.TryParse(dtNew.Rows(i)(idxNewSellingPrice).ToString(), rawNewSellingPrice)
                End If

                ' Resolve barcode from ItemCode
                If rawBarcode = "" AndAlso rawItem <> "" Then
                    Dim rsB As SAPbobsCOM.Recordset =
                    objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    rsB.DoQuery("SELECT ""BcdCode"" FROM OBCD WHERE ""ItemCode"" = '" &
                            rawItem.Replace("'", "''") & "' LIMIT 1")

                    If rsB.RecordCount > 0 Then
                        rawBarcode = rsB.Fields.Item(0).Value.ToString()
                    End If
                End If

                If rawBarcode = "" AndAlso rawItem = "" Then Continue For

                Dim lineNumber As Integer = i + 1

                If Not rowsAdded Then
                    sb.AppendLine(
                    "INSERT INTO ""TNX_PRICEUPDATE"" (" &
                    """BARCODE"",""NEWSELLINGPRICE"",""CREATEDBY"",""LINENUM"",""UPDT"")")
                    sb.Append("SELECT ")
                    rowsAdded = True
                Else
                    sb.AppendLine()
                    sb.Append("UNION ALL SELECT ")
                End If

                sb.AppendFormat(
                "'{0}',{1},'{2}',{3},CURRENT_DATE FROM DUMMY",
                rawBarcode.Replace("'", "''"),
                rawNewSellingPrice,
                CreatedBy,
                lineNumber)

                If sb.Length >= MAX_BATCH_LENGTH Then
                    oRS.DoQuery(sb.ToString())
                    sb.Clear()
                    rowsAdded = False
                End If

            Next

            If sb.Length > 0 Then
                oRS.DoQuery(sb.ToString())
            End If


            'CalculateGPTotals(objForm)
            Dim BranchID As String = objForm.Items.Item("TNXBRC").Specific.Value
            Dim oRecordsetTemp As SAPbobsCOM.Recordset =
    objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim ExistingLineIDs As New List(Of Integer)

            Dim DocEntry As String = oDBs_Head.GetValue("DocEntry", 0).Trim()

            oRecordsetTemp.DoQuery("SELECT ""LineId"" FROM ""@TNX_PCRT_C1"" WHERE ""DocEntry"" = '" & DocEntry & "' AND IFNULL(""U_BrCode"", '') <> '' ORDER BY ""LineId""")

            While Not oRecordsetTemp.EoF

                ExistingLineIDs.Add(CInt(oRecordsetTemp.Fields.Item("LineId").Value))

                oRecordsetTemp.MoveNext()

            End While

            oRecordset.DoQuery("CALL ""TNX_ITEMDETAILS_NEW""('" & CreatedBy & "','" & BranchID & "')")

            If oRecordset.RecordCount > 0 Then

                oRecordset.MoveFirst()

                While Not oRecordset.EoF

                    Dim ProcedureLineID As Integer = CInt(oRecordset.Fields.Item("LINENUM").Value)

                    If Not ExistingLineIDs.Contains(ProcedureLineID) Then

                        Dim TargetRow As Integer = ProcedureLineID - 1

                        oDBs_Details2.Offset = TargetRow

                        oDBs_Details2.SetValue("LineId", TargetRow, ProcedureLineID)
                        oDBs_Details2.SetValue("U_BrCode", TargetRow, oRecordset.Fields.Item("BARCODE").Value)
                        oDBs_Details2.SetValue("U_ItemCode", TargetRow, oRecordset.Fields.Item("ItemCode").Value)
                        oDBs_Details2.SetValue("U_ITNM", TargetRow, oRecordset.Fields.Item("ItemName").Value)
                        oDBs_Details2.SetValue("U_ITGP", TargetRow, oRecordset.Fields.Item("ItemGroup").Value)
                        oDBs_Details2.SetValue("U_CPrice", TargetRow, oRecordset.Fields.Item("RSP").Value)
                        oDBs_Details2.SetValue("U_NSNG", TargetRow, oRecordset.Fields.Item("RSP").Value)
                        oDBs_Details2.SetValue("U_UOM", TargetRow, oRecordset.Fields.Item("UomName").Value)
                        oDBs_Details2.SetValue("U_Cost", TargetRow, oRecordset.Fields.Item("LPCOSTNEW").Value)
                        oDBs_Details2.SetValue("U_LPrice", TargetRow, oRecordset.Fields.Item("LPHPRICE").Value)
                        oDBs_Details2.SetValue("U_NwPrice", TargetRow, oRecordset.Fields.Item("NEWSELLINGPRICE").Value)

                        Dim Cost As Double =
                CDbl(oRecordset.Fields.Item("LPCOSTNEW").Value)

                        Dim NSNG As Double =
                CDbl(oRecordset.Fields.Item("RSP").Value)

                        Dim NMGP As Double = 0

                        If NSNG <> 0 Then
                            NMGP = ((NSNG - Cost) / NSNG) * 100
                        End If

                        oDBs_Details2.SetValue("U_NMGP", TargetRow, NMGP.ToString("0.00"))

                        objMatrix2.SetLineData(ProcedureLineID)

                    End If

                    oRecordset.MoveNext()

                End While

            End If

            For i As Integer = objMatrix2.RowCount To 1 Step -1
                If String.IsNullOrWhiteSpace(objMatrix2.Columns.Item("BrCode").Cells.Item(i).Specific.Value.ToString().Trim()) Then
                    objMatrix2.DeleteRow(i)
                End If
            Next
            objForm.Freeze(False)


            SetMatrixTotals(objForm)
            objMain.objApplication.StatusBar.SetText(
            "Flyer Tab Matrix filled successfully!",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "Matrix Import Error: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            objForm.Freeze(False)
        End Try

        GC.Collect()

    End Sub

    'copy and  paste in Temporay tab code here
    Private Sub UploadFromDataTableTemporayTab(dtNew As DataTable)

        '====================================================
        ' Variable declarations
        '====================================================
        Dim ItemCode As String = ""
        Dim Barcode As String = ""
        Dim Pricelist, nPrice, StartDate As String

        objMatrix4 = objForm.Items.Item("Mtx4").Specific
        objMatrix4.Clear()

        Dim oRS As SAPbobsCOM.Recordset =
        CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset),
              SAPbobsCOM.Recordset)

        Dim oRecordset As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim record As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        '====================================================
        ' Get Logged-in User
        '====================================================
        Dim userCode As String = objMain.objCompany.UserName.Replace("'", "''")
        oRS.DoQuery("SELECT ""USERID"" FROM OUSR WHERE ""USER_CODE"" = '" & userCode & "'")

        Dim CreatedBy As Integer = 0
        If Not oRS.EoF Then
            CreatedBy = CInt(oRS.Fields.Item("USERID").Value)
        End If

        '====================================================
        ' Validate Excel Data
        '====================================================
        If dtNew.Rows.Count = 0 OrElse dtNew.Columns.Count = 0 Then
            objMain.objApplication.StatusBar.SetText(
            "No data found in Excel sheet.",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Exit Sub
        End If

        '====================================================
        ' Detect column indexes
        '====================================================
        Dim idxItemCode As Integer = -1
        Dim idxBarcode As Integer = -1
        Dim idxNewSellingPrice As Integer = -1

        ' Header sniffing
        For c As Integer = 0 To dtNew.Columns.Count - 1
            Dim nm As String = dtNew.Columns(c).ColumnName.ToLower().Trim()

            If nm.Contains("barcode") OrElse nm.Contains("bar code") OrElse nm.Contains("bcd") OrElse nm.Contains("bcode") Then idxBarcode = c
            If nm.Contains("item") OrElse nm.Contains("itemcode") OrElse nm.Contains("item code") Then idxItemCode = c
            If nm.Contains("newsellingprice") OrElse nm.Contains("new selling") _
           OrElse nm.Contains("selling price") Then idxNewSellingPrice = c
        Next

        '====================================================
        ' Create staging table if not exists
        '====================================================
        Dim createTableSQL As String =
        "CREATE COLUMN TABLE ""TNX_PRICEUPDATE""(" &
        """BARCODE"" NVARCHAR(100)," &
        """NEWSELLINGPRICE"" DECIMAL(11,5) CS_FIXED," &
        """CREATEDBY"" NVARCHAR(100)," &
        """LINENUM"" INTEGER," &
        """UPDT"" NVARCHAR(100)) " &
        "UNLOAD PRIORITY 5 AUTO MERGE"

        Try
            oRS.DoQuery(createTableSQL)
        Catch ex As Exception
            If Not ex.Message.ToLower().Contains("already exists") AndAlso
           Not ex.Message.ToLower().Contains("duplicate") Then
                objMain.objApplication.StatusBar.SetText(
                "Error creating table: " & ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        End Try

        '====================================================
        ' Delete old staged data
        '====================================================
        oRS.DoQuery("DELETE FROM ""TNX_PRICEUPDATE"" WHERE ""CREATEDBY"" = '" & CreatedBy & "'")

        objMain.objApplication.StatusBar.SetText(
        "Inserting records into SAP HANA...",
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        '====================================================
        ' Bulk Insert Logic
        '====================================================
        Try
            Dim sb As New StringBuilder()
            Dim rowsAdded As Boolean = False
            Const MAX_BATCH_LENGTH As Integer = 800000

            For i As Integer = 0 To dtNew.Rows.Count - 1

                If dtNew.Rows(i).ItemArray.All(Function(x) x Is DBNull.Value OrElse String.IsNullOrWhiteSpace(x.ToString())) Then Continue For

                Dim rawItem As String = ""
                Dim rawBarcode As String = ""
                Dim rawNewSellingPrice As Decimal = 0

                ' ItemCode
                If idxItemCode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxItemCode)) Then
                    rawItem = dtNew.Rows(i)(idxItemCode).ToString().Trim()
                End If

                ' Barcode
                If idxBarcode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxBarcode)) Then
                    rawBarcode = dtNew.Rows(i)(idxBarcode).ToString().Trim()
                End If

                ' NEW SELLING PRICE from Excel
                If idxNewSellingPrice >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxNewSellingPrice)) Then
                    Decimal.TryParse(dtNew.Rows(i)(idxNewSellingPrice).ToString(), rawNewSellingPrice)
                End If

                ' Resolve Barcode using ItemCode
                If rawBarcode = "" AndAlso rawItem <> "" Then
                    Dim rsB As SAPbobsCOM.Recordset =
                    objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    rsB.DoQuery("SELECT ""BcdCode"" FROM OBCD WHERE ""ItemCode"" = '" &
                            rawItem.Replace("'", "''") & "' LIMIT 1")

                    If rsB.RecordCount > 0 Then
                        rawBarcode = rsB.Fields.Item(0).Value.ToString()
                    End If
                End If

                If rawBarcode = "" AndAlso rawItem = "" Then Continue For

                Dim lineNumber As Integer = i + 1

                If Not rowsAdded Then
                    sb.AppendLine(
                    "INSERT INTO ""TNX_PRICEUPDATE"" (" &
                    """BARCODE"",""NEWSELLINGPRICE"",""CREATEDBY"",""LINENUM"",""UPDT"")")
                    sb.Append("SELECT ")
                    rowsAdded = True
                Else
                    sb.AppendLine()
                    sb.Append("UNION ALL SELECT ")
                End If

                sb.AppendFormat(
                "'{0}',{1},'{2}',{3},CURRENT_DATE FROM DUMMY",
                rawBarcode.Replace("'", "''"),
                rawNewSellingPrice,
                CreatedBy,
                lineNumber)

                If sb.Length >= MAX_BATCH_LENGTH Then
                    oRS.DoQuery(sb.ToString())
                    sb.Clear()
                    rowsAdded = False
                End If

            Next

            If sb.Length > 0 Then
                oRS.DoQuery(sb.ToString())
            End If

            '====================================================
            ' Call stored procedure
            '====================================================
            Dim BranchID As String = objForm.Items.Item("TNXBRC").Specific.value
            oRecordset.DoQuery("CALL ""TNX_ITEMDETAILS_NEW""('" & CreatedBy & "','" & BranchID & "')")

            '====================================================
            ' Fill Temporary Matrix
            '====================================================
            If oRecordset.RecordCount > 0 Then
                objMatrix4.Clear()
                objMatrix4.AddRow(oRecordset.RecordCount)

                Dim lineNum As Integer = 1
                While Not oRecordset.EoF


                    oDBs_Details4.SetValue("LineID", oDBs_Details4.Offset, lineNum)
                    oDBs_Details4.SetValue("U_BrCode", oDBs_Details4.Offset, oRecordset.Fields.Item("BARCODE").Value)
                    oDBs_Details4.SetValue("U_ItemCode", oDBs_Details4.Offset, oRecordset.Fields.Item("ItemCode").Value)
                    oDBs_Details4.SetValue("U_ItemN", oDBs_Details4.Offset, oRecordset.Fields.Item("ItemName").Value)
                    oDBs_Details4.SetValue("U_UOM1", oDBs_Details4.Offset, oRecordset.Fields.Item("UomName").Value)
                    oDBs_Details4.SetValue("U_ICost", oDBs_Details4.Offset, oRecordset.Fields.Item("LPCOSTNEW").Value)
                    oDBs_Details4.SetValue("U_NSell", oDBs_Details4.Offset, oRecordset.Fields.Item("RSP").Value)
                    'oDBs_Details4.SetValue("U_CMarg", oDBs_Details4.Offset, oRecordset.Fields.Item("NEWSELLINGPRICE").Value)
                    oDBs_Details4.SetValue("U_SPrice", oDBs_Details4.Offset, oRecordset.Fields.Item("NEWSELLINGPRICE").Value)
                    oDBs_Details4.SetValue("U_DiffSP", oDBs_Details4.Offset, oRecordset.Fields.Item("NMARG").Value)
                    oDBs_Details4.SetValue("U_DSPNS", oDBs_Details4.Offset, oRecordset.Fields.Item("NCOST").Value)
                    oDBs_Details4.SetValue("U_TPRC", oDBs_Details4.Offset, oRecordset.Fields.Item("LPHPRICE").Value)
                    oDBs_Details4.SetValue("U_TPRS", oDBs_Details4.Offset, oRecordset.Fields.Item("LGRPPRICE").Value)
                    objMatrix4.SetLineData(lineNum)
                    lineNum += 1
                    oRecordset.MoveNext()
                End While
                objForm.Freeze(False)
            End If

            objMain.objApplication.StatusBar.SetText(
            "Temporary Tab Matrix filled successfully!",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "Matrix Import Error: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            objForm.Freeze(False)
        End Try

        GC.Collect()

    End Sub

    Private Sub UploadFromDataTableTemporayTab1(dtNew As DataTable)

        '====================================================
        ' Variable declarations
        '====================================================
        Dim ItemCode As String = ""
        Dim Barcode As String = ""
        Dim Pricelist, nPrice, StartDate As String

        objMatrix4 = objForm.Items.Item("Mtx4").Specific
        'objMatrix4.Clear()

        Dim oRS As SAPbobsCOM.Recordset =
        CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset),
              SAPbobsCOM.Recordset)

        Dim oRecordset As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim record As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        '====================================================
        ' Get Logged-in User
        '====================================================
        Dim userCode As String = objMain.objCompany.UserName.Replace("'", "''")
        oRS.DoQuery("SELECT ""USERID"" FROM OUSR WHERE ""USER_CODE"" = '" & userCode & "'")

        Dim CreatedBy As Integer = 0
        If Not oRS.EoF Then
            CreatedBy = CInt(oRS.Fields.Item("USERID").Value)
        End If

        '====================================================
        ' Validate Excel Data
        '====================================================
        If dtNew.Rows.Count = 0 OrElse dtNew.Columns.Count = 0 Then
            objMain.objApplication.StatusBar.SetText(
            "No data found in Excel sheet.",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Exit Sub
        End If

        '====================================================
        ' Detect column indexes
        '====================================================
        Dim idxItemCode As Integer = -1
        Dim idxBarcode As Integer = -1
        Dim idxNewSellingPrice As Integer = -1

        ' Header sniffing
        For c As Integer = 0 To dtNew.Columns.Count - 1
            Dim nm As String = dtNew.Columns(c).ColumnName.ToLower().Trim()

            If nm.Contains("barcode") OrElse nm.Contains("bar code") OrElse nm.Contains("bcd") OrElse nm.Contains("bcode") Then idxBarcode = c
            If nm.Contains("item") OrElse nm.Contains("itemcode") OrElse nm.Contains("item code") Then idxItemCode = c
            If nm.Contains("newsellingprice") OrElse nm.Contains("new selling") _
           OrElse nm.Contains("selling price") Then idxNewSellingPrice = c
        Next

        '====================================================
        ' Create staging table if not exists
        '====================================================
        Dim createTableSQL As String =
        "CREATE COLUMN TABLE ""TNX_PRICEUPDATE""(" &
        """BARCODE"" NVARCHAR(100)," &
        """NEWSELLINGPRICE"" DECIMAL(11,5) CS_FIXED," &
        """CREATEDBY"" NVARCHAR(100)," &
        """LINENUM"" INTEGER," &
        """UPDT"" NVARCHAR(100)) " &
        "UNLOAD PRIORITY 5 AUTO MERGE"

        Try
            oRS.DoQuery(createTableSQL)
        Catch ex As Exception
            If Not ex.Message.ToLower().Contains("already exists") AndAlso
           Not ex.Message.ToLower().Contains("duplicate") Then
                objMain.objApplication.StatusBar.SetText(
                "Error creating table: " & ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        End Try

        '====================================================
        ' Delete old staged data
        '====================================================
        oRS.DoQuery("DELETE FROM ""TNX_PRICEUPDATE"" WHERE ""CREATEDBY"" = '" & CreatedBy & "'")

        objMain.objApplication.StatusBar.SetText(
        "Inserting records into SAP HANA...",
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        '====================================================
        ' Bulk Insert Logic
        '====================================================
        Try
            Dim sb As New StringBuilder()
            Dim rowsAdded As Boolean = False
            Const MAX_BATCH_LENGTH As Integer = 800000

            For i As Integer = 0 To dtNew.Rows.Count - 1

                If dtNew.Rows(i).ItemArray.All(Function(x) x Is DBNull.Value OrElse String.IsNullOrWhiteSpace(x.ToString())) Then Continue For

                Dim rawItem As String = ""
                Dim rawBarcode As String = ""
                Dim rawNewSellingPrice As Decimal = 0

                ' ItemCode
                If idxItemCode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxItemCode)) Then
                    rawItem = dtNew.Rows(i)(idxItemCode).ToString().Trim()
                End If

                ' Barcode
                If idxBarcode >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxBarcode)) Then
                    rawBarcode = dtNew.Rows(i)(idxBarcode).ToString().Trim()
                End If

                ' NEW SELLING PRICE from Excel
                If idxNewSellingPrice >= 0 AndAlso Not IsDBNull(dtNew.Rows(i)(idxNewSellingPrice)) Then
                    Decimal.TryParse(dtNew.Rows(i)(idxNewSellingPrice).ToString(), rawNewSellingPrice)
                End If

                ' Resolve Barcode using ItemCode
                If rawBarcode = "" AndAlso rawItem <> "" Then
                    Dim rsB As SAPbobsCOM.Recordset =
                    objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    rsB.DoQuery("SELECT ""BcdCode"" FROM OBCD WHERE ""ItemCode"" = '" &
                            rawItem.Replace("'", "''") & "' LIMIT 1")

                    If rsB.RecordCount > 0 Then
                        rawBarcode = rsB.Fields.Item(0).Value.ToString()
                    End If
                End If

                If rawBarcode = "" AndAlso rawItem = "" Then Continue For

                Dim lineNumber As Integer = i + 1

                If Not rowsAdded Then
                    sb.AppendLine(
                    "INSERT INTO ""TNX_PRICEUPDATE"" (" &
                    """BARCODE"",""NEWSELLINGPRICE"",""CREATEDBY"",""LINENUM"",""UPDT"")")
                    sb.Append("SELECT ")
                    rowsAdded = True
                Else
                    sb.AppendLine()
                    sb.Append("UNION ALL SELECT ")
                End If

                sb.AppendFormat(
                "'{0}',{1},'{2}',{3},CURRENT_DATE FROM DUMMY",
                rawBarcode.Replace("'", "''"),
                rawNewSellingPrice,
                CreatedBy,
                lineNumber)

                If sb.Length >= MAX_BATCH_LENGTH Then
                    oRS.DoQuery(sb.ToString())
                    sb.Clear()
                    rowsAdded = False
                End If

            Next

            If sb.Length > 0 Then
                oRS.DoQuery(sb.ToString())
            End If

            Dim BranchID As String = objForm.Items.Item("TNXBRC").Specific.Value
            Dim oRecordsetTemp As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


            Dim ExistingLineIDs As New List(Of Integer)
            Dim DocEntry As String = oDBs_Head.GetValue("DocEntry", 0).Trim()
            oRecordsetTemp.DoQuery("SELECT ""LineId"" FROM ""@TNX_PCRT_C3"" WHERE ""DocEntry"" = '" & DocEntry & "' and IFNULL(""U_BrCode"",'') <> '' ORDER BY ""LineId"" ")

            While Not oRecordsetTemp.EoF
                ExistingLineIDs.Add(CInt(oRecordsetTemp.Fields.Item("LineId").Value))
                oRecordsetTemp.MoveNext()
            End While


            oRecordset.DoQuery("CALL ""TNX_ITEMDETAILS_NEW""('" & CreatedBy & "','" & BranchID & "')")

            If oRecordset.RecordCount > 0 Then
                oRecordset.MoveFirst()

                While Not oRecordset.EoF

                    Dim ProcedureLineID As Integer = CInt(oRecordset.Fields.Item("LINENUM").Value)


                    If Not ExistingLineIDs.Contains(ProcedureLineID) Then


                        Dim TargetRow As Integer = ProcedureLineID - 1

                        oDBs_Details4.Offset = TargetRow

                        oDBs_Details4.SetValue("LineID", TargetRow, ProcedureLineID)
                        oDBs_Details4.SetValue("U_BrCode", TargetRow, oRecordset.Fields.Item("BARCODE").Value)
                        oDBs_Details4.SetValue("U_ItemCode", TargetRow, oRecordset.Fields.Item("ItemCode").Value)
                        oDBs_Details4.SetValue("U_ItemN", TargetRow, oRecordset.Fields.Item("ItemName").Value)
                        oDBs_Details4.SetValue("U_UOM1", TargetRow, oRecordset.Fields.Item("UomName").Value)
                        oDBs_Details4.SetValue("U_ICost", TargetRow, oRecordset.Fields.Item("LPCOSTNEW").Value)
                        oDBs_Details4.SetValue("U_NSell", TargetRow, oRecordset.Fields.Item("RSP").Value)
                        oDBs_Details4.SetValue("U_SPrice", TargetRow, oRecordset.Fields.Item("NEWSELLINGPRICE").Value)
                        oDBs_Details4.SetValue("U_DiffSP", TargetRow, oRecordset.Fields.Item("NMARG").Value)
                        oDBs_Details4.SetValue("U_DSPNS", TargetRow, oRecordset.Fields.Item("NCOST").Value)
                        oDBs_Details4.SetValue("U_TPRC", TargetRow, oRecordset.Fields.Item("LPHPRICE").Value)
                        oDBs_Details4.SetValue("U_TPRS", TargetRow, oRecordset.Fields.Item("LGRPPRICE").Value)

                        objMatrix4.SetLineData(ProcedureLineID)

                    End If

                    oRecordset.MoveNext()
                End While
            End If


            For i As Integer = objMatrix4.RowCount To 1 Step -1
                If String.IsNullOrWhiteSpace(objMatrix4.Columns.Item("BrCode").Cells.Item(i).Specific.Value.ToString().Trim()) Then
                    objMatrix4.DeleteRow(i)
                End If
            Next

            objForm.Freeze(False)




            objMain.objApplication.StatusBar.SetText(
            "Temporary Tab Matrix filled successfully!",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "Matrix Import Error: " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            objForm.Freeze(False)
        End Try

        GC.Collect()

    End Sub



    Private Function IsNumericOnly(value As String) As Boolean
        If String.IsNullOrWhiteSpace(value) Then Return False
        value = value.Trim()
        ' numeric only (digits)
        Return System.Text.RegularExpressions.Regex.IsMatch(value, "^\d+$")
    End Function
    Private Function ColumnLikelyBarcode(dt As DataTable, colIndex As Integer, Optional sampleSize As Integer = 20) As Boolean
        If dt.Rows.Count = 0 Then Return False
        Dim maxSample = Math.Min(sampleSize, dt.Rows.Count)
        Dim numericCount As Integer = 0
        Dim totalCount As Integer = 0
        Dim totalDigitsLength As Integer = 0
        For i As Integer = 0 To maxSample - 1
            Dim obj = dt.Rows(i).Item(colIndex)
            If obj Is Nothing OrElse IsDBNull(obj) Then Continue For
            Dim s As String = obj.ToString().Trim()
            If s = "" Then Continue For
            totalCount += 1
            If IsNumericOnly(s) Then
                numericCount += 1
                totalDigitsLength += s.Length
            End If
        Next
        If totalCount = 0 Then Return False
        ' criteria:
        ' - majority numeric in sample AND average numeric length >= 6
        Dim avgLen As Double = If(numericCount = 0, 0, totalDigitsLength / numericCount)
        Return numericCount >= Math.Ceiling(totalCount / 2.0) AndAlso avgLen >= 6
    End Function

    'Directly Approved without define tEmplate


    'layout printing PCR code

    Public Sub HandleLayoutKeyEvent(ByRef eventInfo As SAPbouiCOM.LayoutKeyInfo, ByRef BubbleEvent As Boolean)

        Try

            If eventInfo.BeforeAction = False Then
                Dim f As SAPbouiCOM.Form = Nothing
                Try
                    f = objMain.objApplication.Forms.Item(eventInfo.FormUID)
                Catch
                    f = Nothing
                End Try

                If f Is Nothing Then Return

                If f.TypeEx = "PCRT" OrElse f.TypeEx = "TNX_PCRT" Then
                    Try
                        'Dim txtObj As SAPbouiCOM.EditText = CType(f.Items.Item("DocEntry").Specific, SAPbouiCOM.EditText)
                        'Dim docVal As String = If(txtObj IsNot Nothing, txtObj.Value.Trim(), "")
                        Dim docVal As String
                        If String.IsNullOrWhiteSpace(docVal) Then
                            ' fallback: try header datasource DocNum
                            Try
                                Dim ds As SAPbouiCOM.DBDataSource = f.DataSources.DBDataSources.Item("@TNX_PCRT")
                                docVal = ds.GetValue("DocEntry", 0).ToString().Trim()
                            Catch
                                docVal = ""
                            End Try
                        End If

                        If Not String.IsNullOrWhiteSpace(docVal) Then
                            eventInfo.LayoutKey = docVal
                        End If
                    Catch ex As Exception
                        ' ignore - cannot supply layout key
                    End Try
                End If
            End If
        Catch ex As Exception
            Try
                objMain.objApplication.StatusBar.SetText("LayoutKey handler error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Catch
            End Try
        End Try

    End Sub

    Private Sub DisableColumns()

        Try
            objForm.Freeze(True)

            Me.objMatrix3.Columns.Item("LPhPrice").Visible = False
            Me.objMatrix3.Columns.Item("LGRPPrice").Visible = False

            Me.objMatrix2.Columns.Item("CPrice").Visible = False
            Me.objMatrix2.Columns.Item("NwPrice").Visible = False

            Me.objMatrix2.Columns.Item("Marg").Visible = False
            Me.objMatrix2.Columns.Item("LPrice").Visible = False

            Me.objMatrix4.Columns.Item("NGP").Visible = False
            Me.objMatrix4.Columns.Item("SGP").Visible = False
            Me.objMatrix3.Columns.Item("NCost").Visible = False

        Catch ex As Exception
            objForm.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        Finally
            objForm.Freeze(False)
        End Try

    End Sub

    'layout printing PCR code enabling test case
    ' 'layout printing PCR code enabling
    Public Sub CreateLayoutFolder_TNX_PCRT()
        Try
            ' Use the connected company object from MainCls (objMain.objCompany). If not available, abort with a clear message.
            Dim companyServiceSource As SAPbobsCOM.Company = Nothing
            Try
                If objMain IsNot Nothing AndAlso objMain.objCompany IsNot Nothing Then
                    companyServiceSource = objMain.objCompany
                End If
            Catch
                companyServiceSource = Nothing
            End Try

            If companyServiceSource Is Nothing Then
                objMain.objApplication.StatusBar.SetText("CreateLayoutFolder_TNX_PCRT aborted: not connected to a company.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return
            End If

            Dim oReportTypeService As SAPbobsCOM.ReportTypesService
            oReportTypeService = companyServiceSource.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService)

            Dim oList = oReportTypeService.GetReportTypeList()

            For i As Integer = 0 To oList.Count - 1
                If oList.Item(i).AddonFormType = "TNX_PCRT" Then
                    Exit Sub   ' Already exists
                End If
            Next

            Dim oReportType As SAPbobsCOM.ReportType
            oReportType = oReportTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportType)

            oReportType.TypeName = "TNX_PCRT Print"
            oReportType.AddonName = "Payroll Addon"
            oReportType.AddonFormType = "TNX_PCRT"
            oReportType.MenuID = "TNX_PCRT_MENU"

            oReportTypeService.AddReportType(oReportType)

        Catch ex As Exception
            Try
                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Catch
            End Try
        End Try
    End Sub

    Private Sub CheckSubmitButtonVisibility(ByVal FormUID As String)

        Try
            Dim objForm As SAPbouiCOM.Form =
      objMain.objApplication.Forms.Item(FormUID)

            Dim isFlyer As Boolean = False
            Dim isFoutYes As Boolean = False

            '---- Get Header Data ----
            Dim oDBs_Head As SAPbouiCOM.DBDataSource =
      objForm.DataSources.DBDataSources.Item("@TNX_PCRT")

            Dim updateType As String =
      oDBs_Head.GetValue("U_TNXPUTP", 0).Trim()

            Dim foutValue As String =
      oDBs_Head.GetValue("U_FOUT", 0).Trim().ToUpper()

            If updateType = "Flyer Price Changes" Then
                isFlyer = True
            End If

            If foutValue = "YES" Then
                isFoutYes = True
            End If

            '---- Final Condition ----
            objForm.Items.Item("Submit").Visible =
      (isFlyer AndAlso isFoutYes)

        Catch ex As Exception

        End Try

    End Sub

End Class

