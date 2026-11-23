Option Strict Off
Option Explicit On
Imports System.Configuration
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Mail
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
''--Naresh Enduri Test push''
''----second push

Public Class MainCls
    ' --- CartSuperMarket_v1
#Region "Declaration"
    Public WithEvents objApplication As SAPbouiCOM.Application
    Public objCompany As SAPbobsCOM.Company
    Public objUtilities As Utilities
    Public objDatabaseCreation As DatabaseCreation
    Public IsFormClosed As Boolean = False
    Public InteNum As Integer
    Public PODocnumentNumber As String = ""
    Public Itemmasterdataupload As Cls_ItemMasterDataUpload
    Public Validate As String = 0

    'Approval Logic
    Public WithEvents Timer As New Timer()
    Public GlobalFormUID As String = ""
    Public GlobalFormUID1 As String = ""
    Public Approver As String = ""
    Public GlobalDocnum As String = ""
    Public mst_FormUIDModal As String
    Dim Responce1 As String = ""


    'GeneralService
    Public oGeneralService As SAPbobsCOM.GeneralService
    Public oGeneralData As SAPbobsCOM.GeneralData
    Public oSons As SAPbobsCOM.GeneralDataCollection
    Public oSon As SAPbobsCOM.GeneralData
    Public oChildren As SAPbobsCOM.GeneralDataCollection
    Public oChild As SAPbobsCOM.GeneralData
    Public sCmp As SAPbobsCOM.CompanyService
    Public oGeneralParams As SAPbobsCOM.GeneralDataParams
    Public oGeneralService1 As SAPbobsCOM.GeneralService
    Public oGeneralData1 As SAPbobsCOM.GeneralData
    Public oChildren1 As SAPbobsCOM.GeneralDataCollection
    Public oChild1 As SAPbobsCOM.GeneralData
    Public sCmp1 As SAPbobsCOM.CompanyService
    Public oGeneralParams1 As SAPbobsCOM.GeneralDataParams
    Public Shared ohtLookUpForm As Hashtable = New Hashtable
    Public oitem As SAPbouiCOM.Item
    Dim SOSeries As String = ""
    Dim SODocNum As String = ""
    Dim PaymentType As String = ""


    'Addon Files
    'Public objAPInvoiceGeneration As clsAPInvoiceGeneration
    Public objAddonConfiguration As clsAddonConfiguration
    Public objPostMultipleAPInvoices As clsPostMultipleAPInvoices
    Public objPostMultipleAPCreditMemos As clsPostMultipleAPCreditMemos
    Public objPurchaseOrder As clsPurchaseOrder
    Public objSalesOrder As clsSalesOrder
    Public objPurchaserequest As clsPurchaseRequest
    Public objPurchaseInvoice As PurchaseInvoice
    Public objPurchaseRequestDetails As clsPurchaseRequestDetails
    Public objStockTransferRequest As clsStockTransferRequest
    Public objGeneratePOfromPurchaseRequest As clsGeneratePOfromPurchaseRequest
    Public objBulkPOGenerationWizard As clsBulkPOProcessingWizard
    Public objItemMaster As clsItemMaster
    Public objInventoryTransfer As clsInventoryTransfer
    'akhila
    Public objGet As cls_GetDetails
    Public objClsPOSBill As ClsPOSBill
    Public objClsBarcodeFilter As ClsBarcodeFilter
    Public objBusinessPartners As clsBusinessPartners

    Public objSupplierAgreement As clsSupplierAgreement
    Public objLoadAgreementsList As clsLoadAgreementsList
    Public objMaterialInward As clsMaterialInward
    Public objListofBinDetails As clsListofBinDetails
    Public objSerialBatchDetails As clsSerialBatchDetails
    Public objPurchaseHistoryDetails As clsPurchaseHistoryDetails
    Public objIncomingpayments As clsIncomingPayment
    Public objGRPO As clsGRPO
    Public objGRPODRAFT As clsGRPODRAFT
    Public objAutomaticProduction As clsAutomaticProduction
    Public objItemList As clsItemList
    Public objSubItemList As clsSubItemList
    Public objOutlet As clsOutlets
    Public objSupplierRebate As clsSupplierrebate1
    Public objPurchaseBlanketAgreement As clsPurchaseBlanketAgreement
    Public objPriceUpdate As clsPriceUpdate
    Public objPriceApprovalScreen As clsPriceApprovalScreen
    Public objPDCPosting As clsPDCPosting
    Public objBarcodeIM As clsBarcodeIM
    Public objBarcodePrinting As clsBarcodePrinting
    Public objBarcode As clsBarcode
    Public objBarcodeMaterailInwards As BarcodeMaterialInward
    Public objBarcodeMaster As clsBarcodeMaster
    Public objWeighingScale As clsWeighingScale
    Public objClearanceSales As clsClearanceSales
    Public objStockTransferDetails As clsStockTransferDetails
    Public objSupplAgreememntPosting As clsSupplierAgreementPosting
    Public objGoodsIssue As clsGoodsIssue
    Public objGoodsReceipt As clsGoodsReceipt
    Public objPDTStockCounting As clsPDTStockCounting
    Public objInventoryCounting As clsInventoryCounting
    Public objOutgoingPayments As ClsOutgoingPayments
    Public objMIGenerationWizard As clsMIGenerationWizard
    Public objPaymentDetails As clsPaymentDetails
    Public objRentalAgreement As clsRentalAgreement
    Public objAccounts As clsAccounts
    Public objStockCountGenerationWizard As clsStockCountGenerationWizard
    Public objStockwizard As clsStockcountwizardNew
    Public objLSIntegrationStatusScreen As clsLSIntegrationStatusScreen
    Public objProductionGenerationWizard As clsProductionGenerationWizard
    Public objUser As clsUser
    Public objUnitsofMeasure As clsUnitsofMeasure
    Public objPriceUpdateItemGroupAuthorization As clsPriceUpdateItemGroupAuthorization
    Public objPriceUpdateBranchAuthorization As clsPriceUpdateBranchAuthorization
    Public objGetfreeitems As clsBuyGetFreeItems
    Public objItemswise As clsItemWiseDiscount
    Public objGetLeastWise As clsBuyAndGetLeastItem
    Public objGetN As clsBuyNGetOneFreeItem
    Public objDiscount As clsDiscountonMaximumAmountPurchase
    Public objPriceConfiguration As clsPriceConfiguration
    Public objPromotionBTotal As clsPromotiononBillTotal
    Public objListItemPR As ClsListingItemsPR
    Public objSupDeposit As ClsSupDeposit
    Public objPettyCashMaster As ClsPettyCashMaster
    Public objExpiry As ClsExpiryDate
    Public objDelivery As ClsDelivery
    Public objGoodsReturn As ClsGoodsReturn
    Public objARInvoice As ClsARInvoice
    Public objARCreditnote As ClsARCreditNot
    Public objAPCreditmemo As ClsAPCreditMemo
    Public objSaleReturn As ClsSaleReturn
    Public objUserCreditNote As ClsUserCreditNote
    Public objDocumentDraw As Cls_DocumentDrawWizard
    Public objDocumentDraw_ITM As Cls_DocumentDraw

    Public ObjPriceUpdateService As PriceUpdateService
    Public objPromocostReduction As clsPromoCostReduction
    Public objPromoReversion As ClsPromoReversion
    Public objPromotionVoucherPurchase As clsPromotionOnVoucherPurchase

    Public objGRPODocRep As ClsGRPODocWiseReport
    Public objAllItem As Cls_AllItemsList
    Public objBarCheck As Cls_BarcodeCheck
    Public objGRPOItem As Cls_GRPOItemWise
    Public objSTCKCNT As Cls_StockCntSummary
    Public objMainWHSE As Cls_MainWHSE
    Public objOpendisp As Cls_OpenDispatch
    Public objTransferRep As Cls_TransferReceiptRpt
    Public objStochWHSE As Cls_StockWHSE
    Public objARDoc As Cls_ARInvoiceDoc
    Public objARItem As Cls_ARInvoiceItem
    Public objAPItem As Cls_APInvoiceItem
    Public objCatIN As Cls_CategoryInOB
    Public objExciseItems As Cls_ExciseItemsRp
    Public ObjBranchMaster As ClsBranchMaster

    Public objBrandMaster As ClsBrandMaster
    Public objSubGroupMaster As ClsSubGroup
    Public objSubSubGroupMaster As ClsSubSubGroup
    Public objSubSubSubGroupMaster As ClsSubSubSubGroup
    Public objDepartmentMaster As ClsDepartment
    Public Shared Error_Message As String = Nothing

    Public MessengerConfiguration As MessengerConfiguration
    Public MessageConfiguration As MessageConfiguration
    Public EmailHistory As EmailHistory
    Public Configuration As Configuration

    Public objBatches As ClsBatch
    Public objsalesQuotation As SalesQuotation
    Public objTNX_PurchaseQuotation As TNX_PurchaseQuotation
    Public objSearchAndFilter As ClsSearchAndFilter
    Public objSalesReturn As ClsSalesReturn
    Public objBatchSelection As clsBatchSelection

    Public objListReport As clsListReport
    Public objCategorySplit As clsCategorySplit
    Public objCentralizedPriceChangeRequest As clsCentralizedPriceChangeRequest

    'Approval
    Public ObjApproval As clsApproval
    Public ObjGRID As clsGRID
    Public ObjADR As clsApprovalDecision
    Public ObjAStg As ClsAPPROVALSTAGES
    Public ObjAPTEMP As clsAPPROVALTEMP
    Public ObjGRIDES As ClsGRIDES
    Public ObjPriceChangeRequest As ClsPriceChangeRequest
    ''Public ObjFLYERGRID As FLYERGRID
    Public ObjDraftProcedure As ClsDraftProcedureStages
    Public ObjVIEW As ClsVIEWFORAPP
    Public objSAPAlertWindow As clsSAPAlertWindow
    Public objBarcodeRequest As clsBarcodeRequest
    Public objSupRebateWizard As ClsSupplierRebateWizard
    'Stock Details Form
    Public objClsStockDetails As ClsStockDetails

    Public objClsBarcodeReport As clsBarcodereport
    Public objMultiplestr As ClsMultipleStockTransfer
    Public objwarehouse As ClsWarehouse
    Public objProductEnquiry As Cls_ProductEnquiry
    Public objItemMasterUploadPreview As clsItemMasterUploadPreview
    Public objPACPPriceChange As clsPACPPriceChange
    Public objSalesBlanketAgreementMaster As ClsSalesBlanketAgreementMaster


    'Public objAddonLog As ClsAddonLog


#End Region

    Public Sub New()
        objUtilities = New Utilities
        objDatabaseCreation = New DatabaseCreation
    End Sub

#Region "Initialilse"
    Public Function Initialise() As Boolean
        'AddHandler Microsoft.Win32.SystemEvents.PowerModeChanged, AddressOf PowerModeChanged
        'AddHandler Microsoft.Win32.SystemEvents.SessionEnding, AddressOf SessionEnding
        objApplication = objUtilities.GetApplication()
        objApplication.StatusBar.SetText("10X Retail Add-On is connecting , Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        If objApplication Is Nothing Then Return False
        objCompany = objUtilities.GetCompany(objApplication)
        If objCompany Is Nothing Then : Return False : Exit Function : End If

        'If Me.Addonlicence("10X_RETAIL", objCompany.CompanyDB, objCompany.UserName) = False Then
        '    MessageBox.Show("Invalid 10X Add-on licence, please contact 10X Team")
        '    objMain.objCompany.Disconnect()
        '    Exit Function
        'End If

        If Not objDatabaseCreation.CreateTables() Then Return False
        Me.LoadFromXML("Menu.xml")
        Me.CustomAuthorization()
        If objMain.objCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            IsNull = "IFNULL"
            GetDate = "NOW"
            HanaLen = "Length"
            HanaInt = "Integer"
            Hana = "Hana"
            Concate = "||"
            SelectCase = ""
        Else
            IsNull = "IsNull"
            GetDate = "GetDate"
            HanaLen = "Len"
            HanaInt = "Int"
            Sql = "Sql"
            Concate = "+"
            SelectCase = "Select"
        End If
        CreateObjects()
        Try
            If objMain.objApplication.Menus.Exists("ME_RETAIL") = True Then
                objMain.objApplication.Menus.Item("ME_RETAIL").Image = System.Windows.Forms.Application.StartupPath & "/Retail.PNG"
            End If
        Catch ex As Exception
        End Try
        objApplication.StatusBar.SetText("10X Retail Add-On is connected Successfully !!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Return True
    End Function
#End Region

#Region "Create Object"
    Private Sub CreateObjects()
        objGet = New cls_GetDetails
        objUserCreditNote = New ClsUserCreditNote
        objSaleReturn = New ClsSaleReturn
        objAPCreditmemo = New ClsAPCreditMemo
        objARCreditnote = New ClsARCreditNot
        objARInvoice = New ClsARInvoice
        objGoodsReturn = New ClsGoodsReturn
        objDelivery = New ClsDelivery
        objExpiry = New ClsExpiryDate
        objAddonConfiguration = New clsAddonConfiguration
        objPostMultipleAPInvoices = New clsPostMultipleAPInvoices
        objPostMultipleAPCreditMemos = New clsPostMultipleAPCreditMemos
        objPurchaseOrder = New clsPurchaseOrder
        objSalesOrder = New clsSalesOrder
        objPurchaserequest = New clsPurchaseRequest
        objPurchaseInvoice = New PurchaseInvoice
        objPurchaseRequestDetails = New clsPurchaseRequestDetails
        objStockTransferRequest = New clsStockTransferRequest
        objGeneratePOfromPurchaseRequest = New clsGeneratePOfromPurchaseRequest
        objBulkPOGenerationWizard = New clsBulkPOProcessingWizard
        objItemMaster = New clsItemMaster
        objInventoryTransfer = New clsInventoryTransfer
        ObjBranchMaster = New ClsBranchMaster
        objBusinessPartners = New clsBusinessPartners
        objClsPOSBill = New ClsPOSBill
        objClsBarcodeFilter = New ClsBarcodeFilter
        objSupplierAgreement = New clsSupplierAgreement
        objLoadAgreementsList = New clsLoadAgreementsList
        objMaterialInward = New clsMaterialInward
        objListofBinDetails = New clsListofBinDetails
        objSerialBatchDetails = New clsSerialBatchDetails
        objPurchaseHistoryDetails = New clsPurchaseHistoryDetails
        objGRPO = New clsGRPO
        objGRPODRAFT = New ClsGRPODRAFT
        objIncomingpayments = New clsIncomingPayment
        objAutomaticProduction = New clsAutomaticProduction
        objItemList = New clsItemList
        objSubItemList = New clsSubItemList
        objOutlet = New clsOutlets
        objSupplierRebate = New clsSupplierrebate1
        objPurchaseBlanketAgreement = New clsPurchaseBlanketAgreement
        objPriceUpdate = New clsPriceUpdate
        objPriceApprovalScreen = New clsPriceApprovalScreen
        objPDCPosting = New clsPDCPosting
        objBarcodeIM = New clsBarcodeIM
        objBarcodePrinting = New clsBarcodePrinting
        objBarcode = New clsBarcode
        objBarcodeMaterailInwards = New BarcodeMaterialInward
        objBarcodeMaster = New clsBarcodeMaster
        objWeighingScale = New clsWeighingScale
        objClearanceSales = New clsClearanceSales
        objDocumentDraw = New Cls_DocumentDrawWizard
        objDocumentDraw_ITM = New Cls_DocumentDraw

        objStockTransferDetails = New clsStockTransferDetails
        objSupplAgreememntPosting = New clsSupplierAgreementPosting
        objGoodsIssue = New clsGoodsIssue
        objGoodsReceipt = New clsGoodsReceipt
        objPDTStockCounting = New clsPDTStockCounting
        objInventoryCounting = New clsInventoryCounting
        objOutgoingPayments = New ClsOutgoingPayments
        objMIGenerationWizard = New clsMIGenerationWizard
        objPaymentDetails = New clsPaymentDetails
        objRentalAgreement = New clsRentalAgreement
        objAccounts = New clsAccounts
        objStockCountGenerationWizard = New clsStockCountGenerationWizard
        objLSIntegrationStatusScreen = New clsLSIntegrationStatusScreen
        objProductionGenerationWizard = New clsProductionGenerationWizard
        objUser = New clsUser
        objUnitsofMeasure = New clsUnitsofMeasure
        objPriceUpdateItemGroupAuthorization = New clsPriceUpdateItemGroupAuthorization
        objPriceUpdateBranchAuthorization = New clsPriceUpdateBranchAuthorization
        objGetfreeitems = New clsBuyGetFreeItems
        objItemswise = New clsItemWiseDiscount
        objGetLeastWise = New clsBuyAndGetLeastItem
        objGetN = New clsBuyNGetOneFreeItem
        objDiscount = New clsDiscountonMaximumAmountPurchase
        objPriceConfiguration = New clsPriceConfiguration
        objPromotionBTotal = New clsPromotiononBillTotal
        objListItemPR = New ClsListingItemsPR
        objStockwizard = New clsStockcountwizardNew
        objSupDeposit = New ClsSupDeposit
        objPettyCashMaster = New ClsPettyCashMaster
        ObjPriceUpdateService = New PriceUpdateService
        Itemmasterdataupload = New Cls_ItemMasterDataUpload
        objPromotionVoucherPurchase = New clsPromotionOnVoucherPurchase

        objPromoReversion = New ClsPromoReversion
        objPromocostReduction = New clsPromoCostReduction
        objGRPODocRep = New ClsGRPODocWiseReport
        objAllItem = New Cls_AllItemsList
        objBarCheck = New Cls_BarcodeCheck
        objGRPOItem = New Cls_GRPOItemWise
        objSTCKCNT = New Cls_StockCntSummary
        objMainWHSE = New Cls_MainWHSE
        objOpendisp = New Cls_OpenDispatch
        objTransferRep = New Cls_TransferReceiptRpt
        objStochWHSE = New Cls_StockWHSE
        objARDoc = New Cls_ARInvoiceDoc
        objARItem = New Cls_ARInvoiceItem
        objAPItem = New Cls_APInvoiceItem
        objCatIN = New Cls_CategoryInOB
        '' ObjFLYERGRID = New FLYERGRID
        objExciseItems = New Cls_ExciseItemsRp

        objBrandMaster = New ClsBrandMaster
        objSubGroupMaster = New ClsSubGroup
        objSubSubGroupMaster = New ClsSubSubGroup
        objSubSubSubGroupMaster = New ClsSubSubSubGroup
        objDepartmentMaster = New ClsDepartment

        MessengerConfiguration = New MessengerConfiguration
        MessageConfiguration = New MessageConfiguration
        EmailHistory = New EmailHistory
        Configuration = New Configuration

        objBatches = New ClsBatch
        objBatchSelection = New clsBatchSelection
        objsalesQuotation = New SalesQuotation
        objTNX_PurchaseQuotation = New TNX_PurchaseQuotation
        objSearchAndFilter = New ClsSearchAndFilter
        objSalesReturn = New ClsSalesReturn

        objListReport = New clsListReport
        objCategorySplit = New clsCategorySplit
        objCentralizedPriceChangeRequest = New clsCentralizedPriceChangeRequest

        'Approval
        ObjApproval = New clsApproval
        ObjGRID = New clsGRID
        ObjADR = New clsApprovalDecision
        ObjAStg = New ClsAPPROVALSTAGES
        ObjAPTEMP = New clsAPPROVALTEMP
        ObjGRIDES = New ClsGRIDES
        ObjPriceChangeRequest = New ClsPriceChangeRequest
        objSupRebateWizard = New ClsSupplierRebateWizard
        objClsStockDetails = New ClsStockDetails
        objClsBarcodeReport = New clsBarcodereport

        objMultiplestr = New ClsMultipleStockTransfer
        objwarehouse = New ClsWarehouse
        objProductEnquiry = New Cls_ProductEnquiry
        objItemMasterUploadPreview = New clsItemMasterUploadPreview
        objPACPPriceChange = New clsPACPPriceChange
        objSalesBlanketAgreementMaster = New ClsSalesBlanketAgreementMaster
        ' Enabling print layout icon
        'price change Requesy form
        Try
            ObjPriceChangeRequest.CreateLayoutFolder_TNX_PCRT()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("CreateLayoutFolder_TNX_PCRT error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

        'Item Listing screen
        Try
            objItemList.CreateLayoutFolder_ItemList()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("CreateLayoutFolder_ItemList error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

        'Resgistration  charges(Supplier Charges Button) screen
        Try
            objSupplAgreememntPosting.CreateLayoutFolder_ResgistrationCharges()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("CreateLayoutFolder_ResgistrationCharges error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

        'Stock Transfer Request screen
        Try
            objStockTransferRequest.CreateLayoutFolder_StockTransferRequest()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("CreateLayoutFolder_StockTransferRequest error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

        ObjDraftProcedure = New ClsDraftProcedureStages
        ObjVIEW = New ClsVIEWFORAPP
        objSAPAlertWindow = New clsSAPAlertWindow
        objBarcodeRequest = New clsBarcodeRequest
        'objAddonLog = New ClsAddonLog
    End Sub
#End Region

#Region "Create UDO"
    Public Sub PromoCostReductionUDO()
        If Not Me.UDOExists("TNX_RDCTNUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_RDCTNUDO", "PrmnCost ReductionUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_RDCTN", "TNX_RDCTN_C0")
            findAliasNDescription = Nothing
        End If

    End Sub

    Public Sub SalesBlanketAgreementMasterUDO()
        If Not Me.UDOExists("SBAMSTUDO") Then
            Dim findaliasndescription = New String(,) {{"Code", "Code"}}
            Me.registerUDONoLog("SBAMSTUDO", "SalesBlanketMaster1 UDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasndescription, "TNX_SBAMRM")
            findaliasndescription = Nothing
        End If
    End Sub

    Public Sub BranchUDO()
        If Not Me.UDOExists("BranchUDO") Then
            Dim findaliasndescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("BranchUDO", "Branch UDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasndescription, "TNXBRANCH", "TNXBRANCH_C0")
            findaliasndescription = Nothing
        End If
    End Sub
    Public Sub RENTPROMOUDO()
        If Not Me.UDOExists("RENTPROMOUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("RENTPROMOUDO", "RENTPROMOUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "RENTPROMO", "RENTPROMO_C0")
            findAliasNDescription = Nothing
        End If

    End Sub
    Public Sub POSBUDO()
        If Not Me.UDOExists("POSBILLUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("POSBILLUDO", "POSBILLUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "POSBIL", "POSB_C0")
            findAliasNDescription = Nothing
        End If

    End Sub
    Public Sub BarcodeReportUDO()
        If Not Me.UDOExists("TNX_BCodeUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_BCodeUDO", "Barcode ReportUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_BCRRTIM")
            findAliasNDescription = Nothing
        End If

    End Sub
    Public Sub BarcodeFilterUDO()
        If Not Me.UDOExists("TNX_BFilterUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_BFilterUDO", "Barcode FilterUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_FIND")
            findAliasNDescription = Nothing
        End If

    End Sub
    Public Sub ItemMasterUploadPreviewUDO()
        If Not Me.UDOExists("ITMPVIEWUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("ITMPVIEWUDO", "Item Master Upload PreviewUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_ITMPEVIEW")
            findAliasNDescription = Nothing
        End If

    End Sub

    Public Sub PromoCostReversionUDO()
        If Not Me.UDOExists("ME_REVERSIONUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("ME_REVERSIONUDO", "REVERSION UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_REVERSION", "TNX_REVERSION_C0", "TNX_REVERSION_C1")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub CreateLicenceNewUDO()
        If Not Me.UDOExists("TNX_LICENSE_UDO") Then
            Dim findaliasndescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_LICENSE_UDO", "Licence  Administration New1 UDO ", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasndescription, "TNX_LICENCE", "TNX_LICENCE_C0")
            findaliasndescription = Nothing
        End If
    End Sub
    Public Sub CreateCreditnoteUDO()
        If Not Me.UDOExists("OTNXCRDNT") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("OTNXCRDNT", "CreditNote UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNXCRDNT", "TNXCRDNT_C0")
            findAliasNDescription = Nothing
        End If

    End Sub
    'Approval
    Public Sub CreateAPPROVALSTAGESUDO()
        If Not Me.UDOExists("SBO_ASTGUDO") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("SBO_ASTGUDO", "SBO_ASTGUDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "SBO_AST", "SBO_AST_C0")
            findAliasNDescription = Nothing
        End If
    End Sub


    Public Sub CreateApprovalTemplatesUDO()
        If Not Me.UDOExists("SBO_APP") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("SBO_APP", "SBO_APP", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "SBO_APPHDR", "SBO_APPREQ", "SBO_APPDOC", "SBO_APPAUT")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub CreateAddonLogUDO()
        If Not Me.UDOExists("SBO_ADONLOG") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("SBO_ADONLOG", "SBO_ADONLOG", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "TNX_LOGST")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub BarcodeRequestUDO()
        If Not Me.UDOExists("TNX_BCRUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_BCRUDO", "TNX_BCRUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_BCR", "TNX_BCR_C0")
            findAliasNDescription = Nothing
        End If

    End Sub
    Public Sub PreciChangeRequestUDO()
        If Not Me.UDOExists("TNXPCRUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNXPCRUDO", "TNXPCRUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_PCRT", "TNX_PCRT_C0", "TNX_PCRT_C1", "TNX_PCRT_C2", "TNX_PCRT_C3", "TNX_PCRT_C4", "TNX_PCRT_C5", "TNX_PCRT_C6", "TNX_PCRT_C7", "TNX_PCRT_C8")
            findAliasNDescription = Nothing
        End If

    End Sub
    Public Sub DDUDO()
        If Not Me.UDOExists("SBO_DDUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("SBO_DDUDO", "SBO_DDUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "SBO_DD", "SBO_DD1")
            findAliasNDescription = Nothing
        End If

    End Sub
    Public Sub DraftUDO()
        If Not Me.UDOExists("SBO_DraftUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("SBO_DraftUDO", "SBO_DraftUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "SBO_OWDD", "SBO_WDD1")
            findAliasNDescription = Nothing
        End If

    End Sub
    Public Sub CreateExpiryDateUDO()
        If Not Me.UDOExists("OTNXEDS") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("OTNXEDS", "Expiry Date UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNXEDS", "TNXEDSC0")
            findAliasNDescription = Nothing
        End If

    End Sub
    Public Sub CreateListReportUDO()
        If Not Me.UDOExists("TNX_LSRPTUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_LSRPTUDO", "List Report UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_LSRPT")
            findAliasNDescription = Nothing
        End If

    End Sub

    Public Sub CreateCategoryUDO()
        If Not Me.UDOExists("TNX_CATUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_CATUDO", "Category Split UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_CATERY", "TNX_CATERY_C0")
            findAliasNDescription = Nothing
        End If

    End Sub

    Public Sub CreateStockTransferRequestUDO()
        If Not Me.UDOExists("ME_OSTR") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("ME_OSTR", "StockTransferRequest UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "ME_STR", "ME_STR_C0", "ME_STR_C1")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub CreateSupRebateWizardUDO()
        If Not Me.UDOExists("TNX_SUPP") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_SUPP", "TNX_SUPP", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_SUPP")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub Itemmasterdataupload1()
        If Not Me.UDOExists("TNX_ITEMUP_UDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_ITEMUP_UDO", "Item MasterData Upload UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_ITEMUP", "TNX_ITEMUP_C0", "TNX_ITEMUP_C1", "TNX_ITEMUP_C2", "TNX_ITEMUP_C3", "TNX_ITEMUP_C4", "TNX_ITEMUP_C5", "TNX_ITEMUP_C6")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CentralizedPriceChangeRequestUDO()
        If Not Me.UDOExists("TNXCPUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNXCPUDO", "TNXCPUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_CP")
            findAliasNDescription = Nothing
        End If

    End Sub
    Public Sub CreateSupplierAgreementUDO()
        If Not Me.UDOExists("ME_OSUPAG") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("ME_OSUPAG", "Supplier Agreement UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "ME_SUPAG", "ME_SUPAG_C0")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub CreateMaterialInwardUDO()
        If Not Me.UDOExists("ME_OMI") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("ME_OMI", "MaterialInward UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "ME_MI", "ME_MI_C0")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub CreateAddonConfigUDO()
        If Not Me.UDOExists("ME_OACFG") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("ME_OACFG", "Addon Config Screen UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "ME_ACFG")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub CreateAutomaticProductionUDO()
        If Not Me.UDOExists("SBO_PRUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("SBO_PRUDO", "Automatic Production UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "SBO_PR", "SBO_PR_C0")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreateItemListUDO()
        If Not Me.UDOExists("SBO_LISTUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("SBO_LISTUDO", "ItemList UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "SBO_LIST", "SBO_LIST_C0", "SBO_LIST_C1", "SBO_LIST_C2")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub SubItemListUDO()
        If Not Me.UDOExists("SBO_LSUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("SBO_LSUDO", "Sub ItemList UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "SBO_LS", "SBO_LS_C0")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub PLUDO()
        If Not Me.UDOExists("SBO_PLUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("SBO_PLUDO", "Price Update UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "SBO_PL", "SBO_PL_C0")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub OutletUDO()
        If Not Me.UDOExists("SBO_OUTLETUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("SBO_OUTLETUDO", "Outlet UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "SBO_OUTLET", "SBO_OUTLET_C0")
            findAliasNDescription = Nothing
        End If
    End Sub

    'Public Sub CreateSupplierRebateUDO()
    '    If Not Me.UDOExists("ME_OSPRB") Then
    '        Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
    '        Me.registerUDO("ME_OSPRB", "SupplierRebate UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "ME_SPRB", "ME_SPRB_C0")
    '        findAliasNDescription = Nothing
    '    End If
    'End Sub
    Public Sub CreateSupplierRebateUDO()
        If Not Me.UDOExists("SBO_SUPREBATEUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("SBO_SUPREBATEUDO", "Supplier Rebate UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "SBO_SUPREBATE", "SBO_SUPREBATE_C0", "SBO_SUPREBATE_C1", "SBO_SUPREBATE_C2")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub CreatePriceUpdateUDO()
        If Not Me.UDOExists("ME_OPRICEUPD
") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("ME_OPRICEUPD", "Price Update UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "ME_PRICEUPD", "ME_PRICEUPD_C0")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub CreatePriceConfigurationUDO()
        If Not Me.UDOExists("PRICECONFIGUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("PRICECONFIGUDO", "Price Configuration UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "SBO_PRICECONFIG", "SBO_PRICECONFIG_C0")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreateFuturePriceUpdateUDO()
        If Not Me.UDOExists("ME_OFPRICEUPD") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("ME_OFPRICEUPD", "Future Price Update UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "ME_FPRICEUPD", "ME_FPRICEUPD_C0")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreatePDCPostingUDO()
        If Not Me.UDOExists("ME_OPDC") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("ME_OPDC", "PDC Posting UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "ME_PDC", "ME_PDC_C0")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreateBarcodeUDO()
        If Not Me.UDOExists("UDO_BARCODE") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("UDO_BARCODE", "Barcode UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "ALE_BARCODE", "ALE_BARCODE1")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreateClearanceSalesUDO()
        If Not Me.UDOExists("SBO_CLEARUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("SBO_CLEARUDO", "Clearance Sales UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "SBO_CLEAR", "SBO_CLEAR_C0")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreateSuplAgreeUDO()
        If Not Me.UDOExists("ME_SUPUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("ME_SUPUDO", "Supplier Agremnt UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "ME_SUP", "ME_SUP_C0", "ME_SUP_C1", "ME_SUP_C2")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreateVehicleMasterUDO()
        If Not Me.UDOExists("ME_OVHMSTR") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("ME_OVHMSTR", "Vehicle MasterUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "ME_VHMSTR")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreatePDTStockCountUDO()
        If Not Me.UDOExists("ME_OPSC") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("ME_OPSC", "PDTStockCountUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "ME_PSC", "ME_PSC_C0")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub CreateBuyGetFreeItemsUDO()
        If Not Me.UDOExists("ME_FREEITMUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("ME_FREEITMUDO", "BuyGetFreeItemsUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_BGFREEITMS", "TNX_BGFREEITMS_C0", "TNX_BGFREEITMS_C1")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub CreateBuyGetLeastItemsUDO()
        If Not Me.UDOExists("ME_LEASTITMUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("ME_LEASTITMUDO", "BuyGetLeastItemsUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_BGLEASTITMS", "TNX_BGLEASTITMS_C0", "TNX_BGLEASTITMS_C1", "TNX_BGLEASTITMS_C2", "TNX_BGLEASTITMS_C3")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub CreateItemsWiseDiscountUDO()
        If Not Me.UDOExists("ME_ITEMUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("ME_ITEMUDO", "ItemWiseDiscountUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_ITEMDISC", "TNX_ITEMDISC_C0", "TNX_ITEMDISC_C1", "TNX_ITEMDISC_C2", "TNX_ITEMDISC_C3")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreatePaymentDetailsScreenUDO()
        If Not Me.UDOExists("ME_OPDS") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("ME_OPDS", "PaymentDetailsUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "ME_PDS", "ME_PDS_C0")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreatePUIGAUDO()
        If Not Me.UDOExists("SBO_OPUIGA") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("SBO_OPUIGA", "PUIGAUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "SBO_PUIGA", "SBO_PUIGA_C0")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreatePUBAUDO()
        If Not Me.UDOExists("SBO_OPUBA") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("SBO_OPUBA", "PUBAUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "SBO_PUBA", "SBO_PUBA_C0")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub STOCKDET()
        If Not Me.UDOExists("STOCKUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("STOCKUDO", "StockDetail_UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "STOCKD", "C_STOCKD")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub PACPPriceChangeUDO()
        If Not Me.UDOExists("PACPRMUDO") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            Me.registerUDONoLog("PACPRMUDO", "PACPPriceChangeUDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "TNX_PACP")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreatePromotiononBillTotalUDO()
        If Not Me.UDOExists("ME_PRMONTOTLUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("ME_PRMONTOTLUDO", "PromotiononBillTotalUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_PONBTOTL", "TNX_PONBTOTL_C0", "TNX_PONBTOTL_C1")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub CreatePromotionOnVoucherPurchaseUDO()
        If Not Me.UDOExists("ME_PRMONVPUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDONoLog("ME_PRMONVPUDO", "PromotionOnVoucherPurchaseUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_PONVP", "TNX_PONVP_C0", "TNX_PONVP_C1", "TNX_PONVP_C2")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreateBuyNGetOneFreeItemUDO()
        If Not Me.UDOExists("ME_BUYITEMUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("ME_BUYITEMUDO", "BuyNGetOneFreeItemUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_BNGOFITM", "TNX_BNGOFITM_C0", "TNX_BNGOFITM_C1")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub CreateStockcountWizardUDO()
        If Not Me.UDOExists("SBO_STOCKWZDUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("SBO_STOCKWZDUDO", "StockcountwizardUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "SBO_STOCKWZD", "SBO_STOCKWZD_C0")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub GRPODOCRpUDO()
        If Not Me.UDOExists("TNX_GRPODRUDO") Then
            Dim findaliasndescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("TNX_GRPODRUDO", "TNX_GRPODRUDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasndescription, "TNX_GRPODR")
            findaliasndescription = Nothing
        End If
    End Sub
    Public Sub SupDepositUDO()
        If Not Me.UDOExists("TNX_SDPTUDO") Then
            Dim findAliasNDescription = New String(,) {{"code", "code"}}
            Me.registerUDONoLog("TNX_SDPTUDO", "TNX_SDPTUDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "TNX_SDPT")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub PettyCashUDO()
        If Not Me.UDOExists("TNX_PCTMUDO") Then
            Dim findAliasNDescription = New String(,) {{"code", "code"}}
            Me.registerUDONoLog("TNX_PCTMUDO", "TNX_PCTMUDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "TNX_PCTM")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub BrandMasterUDO()
        If Not Me.UDOExists("TNXBRDUDO") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            registerUDONoLog("TNXBRDUDO", "TNXBRDUDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "TNX_BRAND")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub SubGroupUDO()
        If Not Me.UDOExists("SUBGROUPUDO") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            registerUDONoLog("SUBGROUPUDO", "SUBGROUPUDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "TNX_SUBGROUP")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub SubSubGroupUDO()
        If Not Me.UDOExists("SSGROUPUDO") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            registerUDONoLog("SSGROUPUDO", "SSGROUPUDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "TNX_SSGROUP")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub SubSubSubGroupUDO()
        If Not Me.UDOExists("SSSGROUPUDO") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            registerUDONoLog("SSSGROUPUDO", "SSSGROUPUDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "TNX_SSSGROUP")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub DepartmentMasterUDO()
        If Not Me.UDOExists("DMSUDO") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            registerUDONoLog("DMSUDO", "DMSUDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "TNX_DMASTER")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub MessageConfigUDO()

        If Not UDOExists("MGRCONFIGUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            registerUDO("MGRCONFIGUDO", "Message Configuation", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "MGRCONFIG", "MGRCONFIG_C0", "MGRCONFIG_C1", "MGRCONFIG_C2")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub SMS_EmailUDO()

        If Not UDOExists("SBO_SETCONFIGUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            registerUDO("SBO_SETCONFIGUDO", "Config UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "SBO_SETCONFIG")
            findAliasNDescription = Nothing
        End If

    End Sub

    Public Sub ConfigurationUDO()

        If Not UDOExists("CONFIGUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            registerUDO("CONFIGUDO", "Configuation", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "CONFIG")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub CreateMultipleStockTransferRequestUDO()
        If Not Me.UDOExists("TNX_MPSTRUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_MPSTRUDO", "Multiple STR UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_MPSTR", "TNX_MPSTR_C0", "TNX_MPSTR_C1")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub CreateWarehouseUDO()
        If Not Me.UDOExists("TNX_WarehouseUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_WarehouseUDO", "Warehouse UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_Warehouse", "TNX_Warehouse_C0")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub ProductEnquiryUDO()
        If Not Me.UDOExists("UDO_PRODUCTENQ") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("UDO_PRODUCTENQ", "Product Enqiry UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_PRDENQUIRY", "TNX_PRDENQUIRY_C0")
            findAliasNDescription = Nothing
        End If
    End Sub



#End Region

#Region "UDO Exists"
    Public Function UDOExists(ByVal code As String) As Boolean
        GC.Collect()
        Dim v_UDOMD As SAPbobsCOM.UserObjectsMD
        Dim v_ReturnCode As Boolean
        v_UDOMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
        v_ReturnCode = v_UDOMD.GetByKey(code)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UDOMD)
        v_UDOMD = Nothing
        Return v_ReturnCode
    End Function
#End Region

#Region "Register UDO"

    Function registerUDO(ByVal UDOCode As String, ByVal UDOName As String, ByVal UDOType As SAPbobsCOM.BoUDOObjType, ByVal findAliasNDescription As String(,), ByVal parentTableName As String, Optional ByVal childTable1 As String = "", Optional ByVal childTable2 As String = "", Optional ByVal childTable3 As String = "", Optional ByVal childTable4 As String = "", Optional ByVal childTable5 As String = "", Optional ByVal childTable6 As String = "", Optional ByVal childTable7 As String = "", Optional ByVal childTable8 As String = "", Optional ByVal childTable9 As String = "", Optional ByVal childTable10 As String = "", Optional ByVal LogOption As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO) As Boolean

        Dim actionSuccess As Boolean = False

        Try

            registerUDO = False

            Dim v_udoMD As SAPbobsCOM.UserObjectsMD

            v_udoMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

            v_udoMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO

            v_udoMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO

            v_udoMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO

            v_udoMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO

            v_udoMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES

            v_udoMD.CanLog = LogOption

            v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES

            v_udoMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES

            v_udoMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES

            v_udoMD.Code = UDOCode

            v_udoMD.Name = UDOName

            v_udoMD.TableName = parentTableName

            If LogOption = SAPbobsCOM.BoYesNoEnum.tYES Then

                v_udoMD.LogTableName = "L" & parentTableName

            End If

            v_udoMD.ObjectType = UDOType

            For i As Int16 = 0 To findAliasNDescription.GetLength(0) - 1

                If i > 0 Then v_udoMD.FindColumns.Add()

                v_udoMD.FindColumns.ColumnAlias = findAliasNDescription(i, 0)

                v_udoMD.FindColumns.ColumnDescription = findAliasNDescription(i, 1)

            Next

            If childTable1 <> "" Then

                v_udoMD.ChildTables.TableName = childTable1

                v_udoMD.ChildTables.Add()

            End If

            If childTable2 <> "" Then

                v_udoMD.ChildTables.TableName = childTable2

                v_udoMD.ChildTables.Add()

            End If

            If childTable3 <> "" Then

                v_udoMD.ChildTables.TableName = childTable3

                v_udoMD.ChildTables.Add()

            End If

            If childTable4 <> "" Then

                v_udoMD.ChildTables.TableName = childTable4

                v_udoMD.ChildTables.Add()

            End If

            If childTable5 <> "" Then

                v_udoMD.ChildTables.TableName = childTable5

                v_udoMD.ChildTables.Add()

            End If

            If childTable6 <> "" Then

                v_udoMD.ChildTables.TableName = childTable6

                v_udoMD.ChildTables.Add()

            End If

            If childTable7 <> "" Then

                v_udoMD.ChildTables.TableName = childTable7

                v_udoMD.ChildTables.Add()

            End If

            If childTable8 <> "" Then

                v_udoMD.ChildTables.TableName = childTable8

                v_udoMD.ChildTables.Add()

            End If

            If childTable9 <> "" Then

                v_udoMD.ChildTables.TableName = childTable9

                v_udoMD.ChildTables.Add()

            End If

            If childTable10 <> "" Then

                v_udoMD.ChildTables.TableName = childTable10

                v_udoMD.ChildTables.Add()

            End If

            If v_udoMD.Add() = 0 Then

                registerUDO = True

                objMain.objApplication.StatusBar.SetText("Successfully Registered UDO >" & UDOCode & ">" & UDOName & " >" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            Else

                objMain.objApplication.StatusBar.SetText("Failed to Register UDO >" & UDOCode & ">" & UDOName & " >" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                registerUDO = False

            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_udoMD)

            v_udoMD = Nothing

            GC.Collect()

        Catch ex As Exception

            objMain.objApplication.SetStatusBarMessage(ex.Message)

        End Try

    End Function



    Function registerUDONoLog(ByVal UDOCode As String, ByVal UDOName As String, ByVal UDOType As SAPbobsCOM.BoUDOObjType, ByVal findAliasNDescription As String(,), ByVal parentTableName As String, Optional ByVal childTable1 As String = "", Optional ByVal childTable2 As String = "", Optional ByVal childTable3 As String = "", Optional ByVal childTable4 As String = "", Optional ByVal childTable5 As String = "", Optional ByVal LogOption As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO) As Boolean

        Dim actionSuccess As Boolean = False

        Try
            registerUDONoLog = False
            Dim v_udoMD As SAPbobsCOM.UserObjectsMD
            v_udoMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            v_udoMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanLog = LogOption
            v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.Code = UDOCode
            v_udoMD.Name = UDOName
            v_udoMD.TableName = parentTableName
            If LogOption = SAPbobsCOM.BoYesNoEnum.tYES Then
                v_udoMD.LogTableName = "A" & parentTableName
            End If
            v_udoMD.ObjectType = UDOType
            For i As Int16 = 0 To findAliasNDescription.GetLength(0) - 1
                If i > 0 Then v_udoMD.FindColumns.Add()
                v_udoMD.FindColumns.ColumnAlias = findAliasNDescription(i, 0)
                v_udoMD.FindColumns.ColumnDescription = findAliasNDescription(i, 1)
            Next

            If childTable1 <> "" Then
                v_udoMD.ChildTables.TableName = childTable1
                v_udoMD.ChildTables.Add()
            End If

            If childTable2 <> "" Then
                v_udoMD.ChildTables.TableName = childTable2
                v_udoMD.ChildTables.Add()
            End If

            If childTable3 <> "" Then
                v_udoMD.ChildTables.TableName = childTable3
                v_udoMD.ChildTables.Add()
            End If

            If childTable4 <> "" Then

                v_udoMD.ChildTables.TableName = childTable4

                v_udoMD.ChildTables.Add()

            End If

            If childTable5 <> "" Then

                v_udoMD.ChildTables.TableName = childTable5

                v_udoMD.ChildTables.Add()

            End If

            If v_udoMD.Add() = 0 Then

                registerUDONoLog = True

                objMain.objApplication.StatusBar.SetText("Successfully Registered UDO >" & UDOCode & ">" & UDOName & " >" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            Else

                objMain.objApplication.StatusBar.SetText("Failed to Register UDO >" & UDOCode & ">" & UDOName & " >" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                registerUDONoLog = False

            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_udoMD)

            v_udoMD = Nothing

            GC.Collect()

        Catch ex As Exception

            objMain.objApplication.SetStatusBarMessage(ex.Message)

        End Try

    End Function


#End Region

#Region "Add Menu's With XML"

    Private Sub LoadFromXML(ByRef FileName As String)
        Dim oXmlDoc As Xml.XmlDocument
        oXmlDoc = New Xml.XmlDocument
        '// load the content of the XML File
        Dim sPath As String
        sPath = IO.Directory.GetParent(Application.ExecutablePath).ToString
        oXmlDoc.Load(sPath & "\" & FileName)
        '// load the form to the SBO application in one batch
        objApplication.LoadBatchActions(oXmlDoc.InnerXml)
        sPath = objApplication.GetLastBatchResults()
    End Sub
#End Region

    Private Sub CustomAuthorization()
        Dim li_RetVal As Integer
        Dim pUserPer As SAPbobsCOM.UserPermissionTree
        Try
            pUserPer = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree)

            If pUserPer.GetByKey("ME_MIGNWZRD") = False Then
                pUserPer.PermissionID = "ME_MIGNWZRD"
                pUserPer.Name = "MI Generation Wizard"
                pUserPer.Options = SAPbobsCOM.BoUPTOptions.bou_FullReadNone
                pUserPer.UserPermissionForms.FormType = "ME_MIGNWZRD_Form"
                li_RetVal = pUserPer.Add()
            End If
            If pUserPer.GetByKey("CPCR") = False Then
                pUserPer.PermissionID = "CPCR"
                pUserPer.Name = "CentralizedPriceChangeRequestReport"
                pUserPer.Options = SAPbobsCOM.BoUPTOptions.bou_FullReadNone
                pUserPer.UserPermissionForms.FormType = "CPRT"
                li_RetVal = pUserPer.Add()
            End If
            If pUserPer.GetByKey("ITEMUP") = False Then
                pUserPer.PermissionID = "ITEMUP"
                pUserPer.Name = "ItemMasterUpload"
                pUserPer.Options = SAPbobsCOM.BoUPTOptions.bou_FullReadNone
                pUserPer.UserPermissionForms.FormType = "ITEMUP"
                li_RetVal = pUserPer.Add()
            End If
        Catch ex As Exception
            Dim ErrCode As Long
            Dim ErrMsg As String
            objMain.objCompany.GetLastError(ErrCode, ErrMsg)
            objMain.objApplication.StatusBar.SetText("Authorization Problem" & ErrCode & li_RetVal, SAPbouiCOM.BoMessageTime.bmt_Short)
        End Try
    End Sub

#Region "Item Event"
    Private Sub objApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles objApplication.ItemEvent
        Try
            Try
                If Retail.MainCls.ohtLookUpForm.ContainsValue(FormUID) = True Then
                    Dim keys As ICollection = Retail.MainCls.ohtLookUpForm.Keys
                    Dim keysArray(Retail.MainCls.ohtLookUpForm.Count - 1) As String
                    keys.CopyTo(keysArray, 0)
                    For Each key As String In keysArray
                        If FormUID = Retail.MainCls.ohtLookUpForm(key) Then
                            While Retail.MainCls.ohtLookUpForm.ContainsValue(key) = True
                                For Each dKey As String In keysArray
                                    If key = Retail.MainCls.ohtLookUpForm(dKey) Then
                                        key = dKey
                                        Exit For
                                    End If
                                Next
                            End While
                            objMain.objApplication.Forms.Item(key).Select()
                            BubbleEvent = False
                            Exit Sub
                        End If
                    Next
                End If
            Catch ex As Exception
            End Try


            Select Case pVal.FormTypeEx

                Case "GETDE"
                    objGet.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "TNXCRDNT_Form"
                    objUserCreditNote.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "SBAMASTER"

                    objSalesBlanketAgreementMaster.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "ListReport_Form"
                    objListReport.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "POSBILL"
                    objClsPOSBill.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "BCode"
                    objClsBarcodeReport.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "Branch"
                    ObjBranchMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                     'Stock Details Form
                Case "STOCKDE"
                    objClsStockDetails.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "PACP"
                    objPACPPriceChange.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "180"
                    objSaleReturn.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "PriceList"
                    ObjPriceUpdateService.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "181"
                    objAPCreditmemo.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "179"
                    objARCreditnote.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "PRMONVP"
                    objPromotionVoucherPurchase.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "B_SUPP"
                    objSupRebateWizard.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "133"
                    objARInvoice.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "Category"
                    objCategorySplit.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "CPRT"
                    objCentralizedPriceChangeRequest.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "PRODEN"
                    objProductEnquiry.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "SBO_Draft"
                    ObjDraftProcedure.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "198"
                    If objSAPAlertWindow IsNot Nothing Then
                        objSAPAlertWindow.ItemEvent(FormUID, pVal, BubbleEvent)
                    End If
                Case "frm_Approve"
                    ObjVIEW.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "182"
                    objGoodsReturn.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "140"
                    objDelivery.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "REDCTN"
                    objPromocostReduction.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "ME_REVERSION"
                    objPromoReversion.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "TNXEDS_Form"
                    objExpiry.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "SBO_PUIGA_Form"
                    objPriceUpdateItemGroupAuthorization.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "SBO_PUBA_Form"
                    objPriceUpdateBranchAuthorization.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "20700"
                    objUser.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "13000001"
                    'objUnitsofMeasure.ItemEvent(FormUID, pVal, BubbleEvent)
                    'Addon Files
                Case "SBO_PGW_Form"
                    objProductionGenerationWizard.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_PDS_Form"
                    objPaymentDetails.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "804"
                    objAccounts.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_PRICEUPD_Form"
                    objPriceUpdate.ItemEvent(FormUID, pVal, BubbleEvent) 'See i kept e
                Case "1250000102"
                    objPurchaseBlanketAgreement.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "1250000100"
                    objRentalAgreement.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_AP_Form"
                    objAutomaticProduction.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_ACFG_Form"
                    objAddonConfiguration.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_LBD_Form"
                    objListofBinDetails.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_SBD_Form"
                    objSerialBatchDetails.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_PMAPINV_Form"
                    objPostMultipleAPInvoices.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "134"
                    objBusinessPartners.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "ME_PMAPCM_Form"
                    objPostMultipleAPCreditMemos.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "142"
                    objPurchaseOrder.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "139"
                    objSalesOrder.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "1470000200"
                    objPurchaserequest.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "141"
                    objPurchaseInvoice.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "143"
                    objGRPO.ItemEvent(FormUID, pVal, BubbleEvent, GlobalDocnum)
                Case "1470000020"
                    objBarcodeMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_PRD_Form"
                    objPurchaseRequestDetails.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_STR_Form"
                    objStockTransferRequest.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_GPOFPRQ_Form"
                    objGeneratePOfromPurchaseRequest.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_BPO_Form"
                    objBulkPOGenerationWizard.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "150"
                    objItemMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "170"
                    objIncomingpayments.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "940"
                    objInventoryTransfer.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_SUPAG_Form"
                    objSupplierAgreement.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_LAL_Form"
                    objLoadAgreementsList.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_MI_Form"
                    objMaterialInward.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_PHD_Form"
                    objPurchaseHistoryDetails.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_ITMLIST"
                    objItemList.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "PRICECONFIG"
                    objPriceConfiguration.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_SUBITMLST_Form"
                    objSubItemList.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "Me_Outlet"
                    objOutlet.ItemEvent(FormUID, pVal, BubbleEvent)
                    'Case "ME_SPRB_Form"
                    '    objSupplierRebate.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "S_Rebate"
                    objSupplierRebate.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_PRAPS_Form"
                    objPriceApprovalScreen.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_PDC_Form"
                    objPDCPosting.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "frmBarCodeIM"
                    objBarcodeIM.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "frmBarCd"
                    objBarcodePrinting.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "frmBarCode"
                    objBarcode.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "frmBC"
                    objBarcodeMaterailInwards.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "frmScale"
                    objWeighingScale.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "Clear"
                    objClearanceSales.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_DDW"
                    objDocumentDraw.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_ILL"
                    objDocumentDraw_ITM.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_STD_Form"
                    objStockTransferDetails.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_SUP_Form"
                    objSupplAgreememntPosting.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "720"
                    objGoodsIssue.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "721"
                    objGoodsReceipt.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_PSC_Form"
                    objPDTStockCounting.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "1474000001"
                    objInventoryCounting.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "426"
                    objOutgoingPayments.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_MIGNWZRD_Form"
                    objMIGenerationWizard.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_SCGW_Form"
                    objStockCountGenerationWizard.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_PSC_Form"
                    objStockCountGenerationWizard.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "TNX_LSINTSTS_Form"
                    objLSIntegrationStatusScreen.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "FREEITM"
                    objGetfreeitems.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ITEM"
                    objItemswise.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "LEASTITM"
                    objGetLeastWise.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "BUYITEM"
                    objGetN.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "DISCOUNT1"
                    objPromotionBTotal.ItemEvent(FormUID, pVal, BubbleEvent)
                'Case "PRMONTOTL"
                '    objPromotionBTotal.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_DISCOUNT"
                    objDiscount.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ITEMLIST"
                    objListItemPR.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_SCGW"
                    objStockwizard.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "SDPTR"
                    objSupDeposit.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "PCTMR"
                    objPettyCashMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ITEMUP"
                    Itemmasterdataupload.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "GRPODWR"
                    objGRPODocRep.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ALLITMS"
                    objAllItem.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "BARCHK"
                    objBarCheck.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GRPOITM"
                    objGRPOItem.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "STCKCNT"
                    objSTCKCNT.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "MWNEDSR"
                    objMainWHSE.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "OPDSPR"
                    objOpendisp.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "TRNSRP"
                    objTransferRep.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "STKWHSE"
                    objStochWHSE.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ARINR"
                    objARDoc.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ARITMR"
                    objARItem.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "APINR"
                    objAPItem.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "CATWIN"
                    objCatIN.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "EXITM"
                    objExciseItems.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "BRANDMAS"
                    objBrandMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "SGMASTER"
                    objSubGroupMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "SSGMASTER"
                    objSubSubGroupMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "SSSGMASTER"
                    objSubSubSubGroupMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "DMASTER"
                    objDepartmentMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "SBO_SETCONFIG"
                    MessengerConfiguration.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "MGRCONFIG"
                    MessageConfiguration.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "EMLLog"
                    EmailHistory.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "CONFIG"
                    Configuration.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "149"
                    objsalesQuotation.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "540000988"
                    objTNX_PurchaseQuotation.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "SearchAndFilter"
                    objSearchAndFilter.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "42"
                    objBatches.ItemEvent(FormUID, pVal, BubbleEvent)
                    'Approval
                Case "ME_ASR"
                    ObjApproval.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GRID"
                    ObjGRID.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_ADR"
                    ObjADR.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "SBO_AST"
                    ObjAStg.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GRIDES"
                    ObjGRIDES.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "Temp"
                    ObjAPTEMP.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "PCRT"
                    ObjPriceChangeRequest.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_FILTER"
                    objClsBarcodeFilter.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "BarcodeRequest"
                    objBarcodeRequest.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "IMPPerview"
                    objItemMasterUploadPreview.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "TNX_POPost"
                    objMultiplestr.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "W_Master"
                    objwarehouse.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ADDON"
                    'objAddonLog.ItemEvent(FormUID, pVal, BubbleEvent)
            End Select

            If Not IsNothing(mst_FormUIDModal) Then
                If FormUID <> mst_FormUIDModal Then
                    Dim dbo_FormFound As Boolean = False
                    Dim dsa_Form As SAPbouiCOM.Form
                    For Each dsa_Form In objApplication.Forms
                        If dsa_Form.UniqueID = mst_FormUIDModal Then
                            If FormUID <> mst_FormUIDModal Then
                                If Not objApplication.Forms.Item(mst_FormUIDModal).Selected() Then
                                    objApplication.Forms.Item(mst_FormUIDModal).Select()
                                End If
                                BubbleEvent = False
                                Exit Sub
                            End If
                            dsa_Form = Nothing
                            dbo_FormFound = True
                        End If
                    Next
                    dsa_Form = Nothing
                    If Not dbo_FormFound Then
                        Me.mst_FormUIDModal = Nothing
                    End If
                End If
            End If
            If pVal.FormType = "139" Or pVal.FormType = "60091" Or pVal.FormType = "142" Or pVal.FormType = "149" Or pVal.FormType = "140" Or pVal.FormType = "133" Or pVal.FormType = "180" Or
                pVal.FormType = "179" Or pVal.FormType = "540000988" Or pVal.FormType = "143" Or pVal.FormType = "141" Or pVal.FormType = "181" Or pVal.FormType = "182" Or pVal.FormType = "1250000100" Or
                pVal.FormType = "1250000102" Or pVal.FormType = "170" Or pVal.FormType = "426" Or pVal.FormType = "392" Or pVal.FormType = "721" Or pVal.FormType = "720" Or pVal.FormType = "1250000940" Or pVal.FormType = "940" Then
                Dim DocType As String = ""
                Dim DocTypeField As String = ""
                Dim DocTypeFieldSMS As String = ""
                Dim DocTypeFieldEmail As String = ""

                If pVal.FormType = "1250000100" Then
                    DocType = "SBA"
                    DocTypeField = "U_SBA"
                    DocTypeFieldEmail = "U_SBAEML"
                    DocTypeFieldSMS = "U_SBASMS"
                ElseIf pVal.FormType = "1250000102" Then
                    DocType = "PBA"
                    DocTypeField = "U_PBA"
                    DocTypeFieldEmail = "U_PBAEML"
                    DocTypeFieldSMS = "U_PBASMS"
                ElseIf pVal.FormType = "139" Then
                    DocType = "SO"
                    DocTypeField = "U_SO"
                    DocTypeFieldEmail = "U_SOEML"
                    ''DocTypeFieldEmail = "Y"
                    DocTypeFieldSMS = "U_SOSMS"
                ElseIf pVal.FormType = "142" Then
                    DocType = "PO"
                    DocTypeField = "U_PO"
                    DocTypeFieldEmail = "U_POEML"
                    DocTypeFieldSMS = "U_POSMS"
                ElseIf pVal.FormType = "149" Then
                    DocType = "SQ"
                    DocTypeField = "U_SQ"
                    DocTypeFieldEmail = "U_SQEML"
                    DocTypeFieldSMS = "U_SQSMS"
                ElseIf pVal.FormType = "140" Then
                    DocType = "DL"
                    DocTypeField = "U_DL"
                    DocTypeFieldEmail = "U_DLEML"
                    DocTypeFieldSMS = "U_DLSMS"
                ElseIf pVal.FormType = "133" Then
                    DocType = "ARINV"
                    DocTypeField = "U_ARINV"
                    DocTypeFieldEmail = "U_ARINVEML"
                    DocTypeFieldSMS = "U_ARINVSMS"
                ElseIf pVal.FormType = "60091" Then
                    DocType = "ARRINV"
                    DocTypeField = "U_ARRINV"
                    DocTypeFieldEmail = "U_ARINVEML"
                    DocTypeFieldSMS = "U_ARINVSMS"
                ElseIf pVal.FormType = "180" Then
                    DocType = "ARRET"
                    DocTypeField = "U_ARRET"
                    DocTypeFieldEmail = "U_ARRETEML"
                    DocTypeFieldSMS = "U_ARRETSMS"
                ElseIf pVal.FormType = "179" Then
                    DocType = "ARC"
                    DocTypeField = "U_ARC"
                    DocTypeFieldEmail = "U_ARCEML"
                    DocTypeFieldSMS = "U_ARCSMS"
                ElseIf pVal.FormType = "540000988" Then
                    DocType = "PQ"
                    DocTypeField = "U_PQ"
                    DocTypeFieldEmail = "U_PQEML"
                    DocTypeFieldSMS = "U_PQSMS"
                ElseIf pVal.FormType = "143" Then
                    DocType = "GRPO"
                    DocTypeField = "U_GRPO"
                    DocTypeFieldEmail = "U_GRPOEML"
                    DocTypeFieldSMS = "U_GRPOSMS"
                ElseIf pVal.FormType = "141" Then
                    DocType = "APINV"
                    DocTypeField = "U_APINV"
                    DocTypeFieldEmail = "U_APINVEML"
                    DocTypeFieldSMS = "U_APINVSMS"
                ElseIf pVal.FormType = "182" Then
                    DocType = "GR"
                    DocTypeField = "U_GR"
                    DocTypeFieldEmail = "U_GREML"
                    DocTypeFieldSMS = "U_GRSMS"
                ElseIf pVal.FormType = "181" Then
                    DocType = "APC"
                    DocTypeField = "U_APC"
                    DocTypeFieldEmail = "U_APINVEML"
                    DocTypeFieldSMS = "U_APCSMS"
                ElseIf pVal.FormType = "1250000100" Then
                    DocType = "SBA"
                    DocTypeField = "U_SBA"
                    DocTypeFieldEmail = "U_SBAML"
                    DocTypeFieldSMS = "U_SBASMS"
                ElseIf pVal.FormType = "1250000102" Then
                    DocType = "PBA"
                    DocTypeField = "U_PBA"
                    DocTypeFieldEmail = "U_PBAML"
                    DocTypeFieldSMS = "U_PBASMS"
                ElseIf pVal.FormType = "170" Then
                    DocType = "IP"
                    DocTypeField = "U_IP"
                    DocTypeFieldEmail = "U_IPEML"
                    DocTypeFieldSMS = "U_IPSMS"
                ElseIf pVal.FormType = "426" Then
                    DocType = "OP"
                    DocTypeField = "U_OP"
                    DocTypeFieldEmail = "U_OPEML"
                    DocTypeFieldSMS = "U_OPSMS"
                ElseIf pVal.FormType = "392" Then
                    DocType = "JE"
                    DocTypeField = "U_JE"
                    DocTypeFieldEmail = "U_JEEML"
                    DocTypeFieldSMS = "U_JESMS"
                ElseIf pVal.FormType = "721" Then
                    DocType = "GRP"
                    DocTypeField = "U_GRP"
                    DocTypeFieldEmail = "U_GRPEML"
                    DocTypeFieldSMS = "U_GRPSMS"
                ElseIf pVal.FormType = "720" Then
                    DocType = "GI"
                    DocTypeField = "U_GI"
                    DocTypeFieldEmail = "U_GIEML"
                    DocTypeFieldSMS = "U_GISMS"
                ElseIf pVal.FormType = "1250000940" Then
                    DocType = "ITR"
                    DocTypeField = "U_ITR"
                    DocTypeFieldEmail = "U_ITREML"
                    DocTypeFieldSMS = "U_ITRSMS"
                ElseIf pVal.FormType = "940" Then
                    DocType = "IT"
                    DocTypeField = "U_IT"
                    DocTypeFieldEmail = "U_ITEML"
                    DocTypeFieldSMS = "U_ITSMS"
                End If
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        If pVal.Before_Action = False Then
                            Dim AppTemForm As SAPbouiCOM.Form
                            AppTemForm = objApplication.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                            Dim Manual As String = objUtilities.MessageTrigger(DocType, "WhatsApp")
                            Dim ManualS As String = objUtilities.MessageTrigger(DocType, "SMS")
                            Dim ManualE As String = objUtilities.MessageTrigger(DocType, "Email")

                            If Manual = "1" Or ManualS = "1" Or ManualE = "1" Then
                                If pVal.FormType = "1250000100" Or pVal.FormType = "1250000102" Then
                                    oitem = AppTemForm.Items.Add("Send", SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO)
                                    oitem.Top = AppTemForm.Items.Item("1250000002").Top
                                    oitem.Left = AppTemForm.Items.Item("1250000002").Left + AppTemForm.Items.Item("1250000002").Width + 5
                                    oitem.Width = AppTemForm.Items.Item("1250000002").Width
                                    oitem.Height = AppTemForm.Items.Item("1250000002").Height
                                    oitem.FromPane = AppTemForm.Items.Item("1250000002").FromPane
                                    oitem.ToPane = AppTemForm.Items.Item("1250000002").ToPane
                                    oitem.AffectsFormMode = False
                                    Dim btn As SAPbouiCOM.ButtonCombo = AppTemForm.Items.Item("Send").Specific
                                    If Manual = "1" Then
                                        btn.ValidValues.Add("WhatsApp", "WhatsApp")
                                    End If
                                    If ManualS = "1" Then
                                        btn.ValidValues.Add("SMS", "SMS")
                                    End If
                                    If ManualE = "1" Then
                                        btn.ValidValues.Add("E-Mail", "E-Mail")
                                    End If
                                    btn.Caption = "Send"
                                Else
                                    oitem = AppTemForm.Items.Add("Send", SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO)
                                    oitem.Top = AppTemForm.Items.Item("2").Top
                                    oitem.Left = AppTemForm.Items.Item("2").Left + AppTemForm.Items.Item("2").Width + 5
                                    oitem.Width = AppTemForm.Items.Item("2").Width
                                    oitem.Height = AppTemForm.Items.Item("2").Height
                                    oitem.FromPane = AppTemForm.Items.Item("2").FromPane
                                    oitem.ToPane = AppTemForm.Items.Item("2").ToPane
                                    oitem.AffectsFormMode = False
                                    Dim btn As SAPbouiCOM.ButtonCombo = AppTemForm.Items.Item("Send").Specific
                                    If Manual = "1" Then
                                        btn.ValidValues.Add("WhatsApp", "WhatsApp")
                                    End If
                                    If ManualS = "1" Then
                                        btn.ValidValues.Add("SMS", "SMS")
                                    End If
                                    If ManualE = "1" Then
                                        btn.ValidValues.Add("E-Mail", "E-Mail")
                                    End If
                                    btn.Caption = "Send"
                                End If
                            End If
                        End If


                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        Dim AppTemForm1 As SAPbouiCOM.Form
                        AppTemForm1 = objApplication.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)

                        If pVal.ItemUID = "1" And (AppTemForm1.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or AppTemForm1.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.BeforeAction = True Then
                            If DocTypeField = "U_IP" Or DocTypeField = "U_OP" Then
                                SOSeries = AppTemForm1.Items.Item("87").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                SODocNum = AppTemForm1.Items.Item("3").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                Dim CardA As SAPbouiCOM.OptionBtn = AppTemForm1.Items.Item("58").Specific
                                If CardA.Selected = True Then
                                    PaymentType = "Account"
                                Else
                                    PaymentType = ""
                                End If
                            ElseIf DocTypeField = "U_SO" Then
                                SOSeries = AppTemForm1.Items.Item("88").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                SODocNum = AppTemForm1.Items.Item("8").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                            ElseIf DocTypeField = "U_GRP" Or DocTypeField = "U_GI" Then
                                SOSeries = AppTemForm1.Items.Item("30").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                SODocNum = AppTemForm1.Items.Item("7").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                            ElseIf DocTypeField = "U_ITR" Then
                                SOSeries = AppTemForm1.Items.Item("1250000068").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                SODocNum = AppTemForm1.Items.Item("11").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                            ElseIf DocTypeField = "U_IT" Then
                                SOSeries = AppTemForm1.Items.Item("40").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                SODocNum = AppTemForm1.Items.Item("11").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                            ElseIf DocTypeField = "U_JE" Then
                                SOSeries = AppTemForm1.Items.Item("137").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                SODocNum = AppTemForm1.Items.Item("141").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                            Else
                                SOSeries = AppTemForm1.Items.Item("88").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                SODocNum = AppTemForm1.Items.Item("8").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                            End If
                        End If







                        If pVal.ItemUID = "1" And (AppTemForm1.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or AppTemForm1.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE) And pVal.BeforeAction = False And pVal.Action_Success = True Then
                            Try
                                If objUtilities.DocTrigger(DocTypeField) = "Y" Then
                                    Dim Manual As String = objUtilities.MessageTrigger(DocType, "WhatsApp")
                                    If Manual = "2" Then
                                        ' objUtilities.Document_Update(DocType, SODocNum)
                                        Me.MarketingWhatsAppMessage(SODocNum, SOSeries, DocType)

                                    End If
                                ElseIf objUtilities.DocTrigger(DocTypeField) = "N" Then
                                    'objMain.objUtilities.Write_Log("Document Configuration Disable for " + DocType + " Document !!! ")
                                    'objApplication.StatusBar.SetText("Document Configuration Disable for " + DocType + " Document !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If
                                If objUtilities.DocTrigger(DocTypeFieldSMS) = "Y" Then
                                    Dim ManualS As String = objUtilities.MessageTrigger(DocType, "SMS")
                                    If ManualS = "2" Then
                                        Me.MarketingSMSMessage(SODocNum, SOSeries, DocType)
                                    End If
                                ElseIf objUtilities.DocTrigger(DocTypeFieldSMS) = "N" Then
                                    'objMain.objUtilities.Write_Log("Document Configuration Disable for " + DocType + " Document !!! ")
                                    'objApplication.StatusBar.SetText("Document Configuration Disable for " + DocType + " Document !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If
                                If objUtilities.DocTrigger(DocTypeFieldEmail) = "Y" Then
                                    Dim ManualE As String = objUtilities.MessageTrigger(DocType, "Email")
                                    If ManualE = "2" Then
                                        Me.MarketingEmail(SODocNum, SOSeries, DocType)
                                    End If
                                ElseIf objUtilities.DocTrigger(DocTypeFieldEmail) = "N" Then
                                    'objMain.objUtilities.Write_Log("Document Configuration Disable for " + DocType + " Document !!! ")
                                    'objApplication.StatusBar.SetText("Document Configuration Disable for " + DocType + " Document !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If

                                'End If
                                'AppTemForm1.Items.Item("Send").Click()

                            Catch ex As Exception
                                objMain.objUtilities.Write_Log(ex.Message)
                                objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                            End Try
                        End If
                        Try
                            If pVal.ItemUID = "Send" And AppTemForm1.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE And pVal.BeforeAction = False Then
                                Dim btnCombo As SAPbouiCOM.ButtonCombo = AppTemForm1.Items.Item("Send").Specific
                                If btnCombo.Selected.Value = "WhatsApp" Then
                                    If objUtilities.DocTrigger(DocTypeField) = "Y" Then
                                        If DocTypeField = "U_IP" Or DocTypeField = "U_OP" Then
                                            SOSeries = AppTemForm1.Items.Item("87").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                            SODocNum = AppTemForm1.Items.Item("3").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                            Dim CardA As SAPbouiCOM.OptionBtn = AppTemForm1.Items.Item("58").Specific
                                            If CardA.Selected = True Then
                                                PaymentType = "Account"
                                            Else
                                                PaymentType = ""
                                            End If
                                        ElseIf DocTypeField = "U_GRP" Or DocTypeField = "U_GI" Then
                                            SOSeries = AppTemForm1.Items.Item("30").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                            SODocNum = AppTemForm1.Items.Item("7").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                        ElseIf DocTypeField = "U_ITR" Then
                                            SOSeries = AppTemForm1.Items.Item("1250000068").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                            SODocNum = AppTemForm1.Items.Item("11").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                        ElseIf DocTypeField = "U_IT" Then
                                            SOSeries = AppTemForm1.Items.Item("40").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                            SODocNum = AppTemForm1.Items.Item("11").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                        ElseIf DocTypeField = "U_JE" Then
                                            SOSeries = AppTemForm1.Items.Item("137").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                            SODocNum = AppTemForm1.Items.Item("141").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                        Else
                                            SOSeries = AppTemForm1.Items.Item("88").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                            SODocNum = AppTemForm1.Items.Item("8").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                        End If
                                        Dim formName As String = AppTemForm1.Title

                                        MarketingWhatsAppMessage(SODocNum, SOSeries, DocType)
                                    Else
                                        objMain.objUtilities.Write_Log("Document Configuration Disable for " + DocType + " Document !!! ")
                                        objApplication.StatusBar.SetText("Document Configuration Disable for " + DocType + " Document !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If
                                End If
                                If btnCombo.Selected.Value = "SMS" Then
                                    If objUtilities.DocTrigger(DocTypeFieldSMS) = "Y" Then
                                        If DocTypeField = "U_IP" Or DocTypeField = "U_OP" Then
                                            SOSeries = AppTemForm1.Items.Item("87").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                            SODocNum = AppTemForm1.Items.Item("3").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                            Dim CardA As SAPbouiCOM.OptionBtn = AppTemForm1.Items.Item("58").Specific
                                            If CardA.Selected = True Then
                                                PaymentType = "Account"
                                            Else
                                                PaymentType = ""
                                            End If
                                        ElseIf DocTypeField = "U_GRP" Or DocTypeField = "U_GI" Then
                                            SOSeries = AppTemForm1.Items.Item("30").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                            SODocNum = AppTemForm1.Items.Item("7").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                        ElseIf DocTypeField = "U_ITR" Then
                                            SOSeries = AppTemForm1.Items.Item("1250000068").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                            SODocNum = AppTemForm1.Items.Item("11").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                        ElseIf DocTypeField = "U_IT" Then
                                            SOSeries = AppTemForm1.Items.Item("40").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                            SODocNum = AppTemForm1.Items.Item("11").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                        ElseIf DocTypeField = "U_JE" Then
                                            SOSeries = AppTemForm1.Items.Item("137").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                            SODocNum = AppTemForm1.Items.Item("141").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                        Else
                                            SOSeries = AppTemForm1.Items.Item("88").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                            SODocNum = AppTemForm1.Items.Item("8").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                        End If
                                        Me.MarketingSMSMessage(SODocNum, SOSeries, DocType)
                                    Else
                                        objMain.objUtilities.Write_Log("Document Configuration Disable for " + DocType + " Document !!! ")
                                        objApplication.StatusBar.SetText("Document Configuration Disable for " + DocType + " Document !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If
                                End If
                                If btnCombo.Selected.Value = "E-Mail" Then
                                    ' If objUtilities.DocTrigger(DocTypeFieldEmail) = "Y" Then
                                    If DocTypeField = "U_IP" Or DocTypeField = "U_OP" Then
                                        SOSeries = AppTemForm1.Items.Item("87").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                        SODocNum = AppTemForm1.Items.Item("3").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                        Dim CardA As SAPbouiCOM.OptionBtn = AppTemForm1.Items.Item("58").Specific
                                        If CardA.Selected = True Then
                                            PaymentType = "Account"
                                        Else
                                            PaymentType = ""
                                        End If
                                    ElseIf DocTypeField = "U_GRP" Or DocTypeField = "U_GI" Then
                                        SOSeries = AppTemForm1.Items.Item("30").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                        SODocNum = AppTemForm1.Items.Item("7").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                    ElseIf DocTypeField = "U_ITR" Then
                                        SOSeries = AppTemForm1.Items.Item("1250000068").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                        SODocNum = AppTemForm1.Items.Item("11").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                    ElseIf DocTypeField = "U_IT" Then
                                        SOSeries = AppTemForm1.Items.Item("40").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                        SODocNum = AppTemForm1.Items.Item("11").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                    ElseIf DocTypeField = "U_JE" Then
                                        SOSeries = AppTemForm1.Items.Item("137").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                        SODocNum = AppTemForm1.Items.Item("141").Specific.Value.ToString().TrimEnd().Replace(" ", "")
                                    Else
                                        SOSeries = AppTemForm1.Items.Item("88").Specific.Value.ToString().Trim().Replace(" ", "")
                                        SODocNum = AppTemForm1.Items.Item("8").Specific.Value.ToString().Trim().Replace(" ", "")
                                    End If
                                    Me.MarketingEmail(SODocNum, SOSeries, DocType)
                                    ' Else
                                    ' objMain.objUtilities.Write_Log("Document Configuration Disable for " + DocType + " Document !!! ")
                                    'objApplication.StatusBar.SetText("Document Configuration Disable for " + DocType + " Document !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    ' End If
                                End If
                            End If
                        Catch ex As Exception
                            objMain.objUtilities.Write_Log(ex.Message)
                            objApplication.StatusBar.SetText("Please select the message type", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        End Try

                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        Dim AppTemForm1 As SAPbouiCOM.Form
                        AppTemForm1 = objApplication.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount)
                        If pVal.ItemUID = "Send" And AppTemForm1.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE And pVal.BeforeAction = False Then
                            Dim btnCombo As SAPbouiCOM.ButtonCombo = AppTemForm1.Items.Item("Send").Specific

                        End If
                End Select
            End If
        Catch ex As Exception
            objApplication.MessageBox(ex.Message)
        End Try
    End Sub
#End Region

#Region "Menu Events"
    Private Sub objApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles objApplication.MenuEvent
        Dim objform As SAPbouiCOM.Form
        Try
            'objform = objMain.objApplication.Forms.ActiveForm
            Select Case pVal.MenuUID
                Case "ITEMUP"
                    Itemmasterdataupload.MenuEvent(pVal, BubbleEvent)
                Case "CPCR"
                    objCentralizedPriceChangeRequest.MenuEvent(pVal, BubbleEvent)
                Case "POSBILL"
                    objClsPOSBill.MenuEvent(pVal, BubbleEvent)
                Case "ITEMLIST"
                    objSubItemList.MenuEvent(pVal, BubbleEvent)
                Case "PACP"
                    objPACPPriceChange.MenuEvent(pVal, BubbleEvent)
                    'Stock Details Form
                Case "YD1"
                    objClsStockDetails.MenuEvent(pVal, BubbleEvent)
                Case "ME_SUPP"
                    objSupRebateWizard.MenuEvent(pVal, BubbleEvent)
                Case "Branch"
                    ObjBranchMaster.MenuEvent(pVal, BubbleEvent)
                Case "PRODEN"
                    objProductEnquiry.MenuEvent(pVal, BubbleEvent)
                Case "ITMUP"
                    objItemMasterUploadPreview.MenuEvent(pVal, BubbleEvent)
                Case "SBAMASTER"
                    objSalesBlanketAgreementMaster.MenuEvent(pVal, BubbleEvent)

                Case "ME_REDCTN"
                    objPromocostReduction.MenuEvent(pVal, BubbleEvent)
                Case "3073"
                    objItemMaster.MenuEvent(pVal, BubbleEvent)

                Case "ME_REVERSION"
                    objPromoReversion.MenuEvent(pVal, BubbleEvent)

                Case "BCReport"
                    objClsBarcodeReport.MenuEvent(pVal, BubbleEvent)

                Case "TNXCRDNT"
                    objUserCreditNote.MenuEvent(pVal, BubbleEvent)
                Case "PriceList"
                    ObjPriceUpdateService.MenuEvent(pVal, BubbleEvent)

                Case "ME_PRICEUPD"
                    objPriceUpdate.MenuEvent(pVal, BubbleEvent)
                Case "ME_GPOFPRQ"
                    objGeneratePOfromPurchaseRequest.MenuEvent(pVal, BubbleEvent)
                Case "ME_BPO"
                    objBulkPOGenerationWizard.MenuEvent(pVal, BubbleEvent)
                Case "ME_PMAPINV"
                    objPostMultipleAPInvoices.MenuEvent(pVal, BubbleEvent)
                Case "ME_PMAPCM"
                    objPostMultipleAPCreditMemos.MenuEvent(pVal, BubbleEvent)
                Case "ME_STR"
                    objStockTransferRequest.MenuEvent(pVal, BubbleEvent)
                Case "ME_MI"
                    objMaterialInward.MenuEvent(pVal, BubbleEvent)
                Case "ME_ACFG"
                    objAddonConfiguration.MenuEvent(pVal, BubbleEvent)
                Case "ME_AP"
                    objAutomaticProduction.MenuEvent(pVal, BubbleEvent)
                Case "ME_ITMLIST"
                    objItemList.MenuEvent(pVal, BubbleEvent)
                Case "ME_PRAPS"
                    objPriceApprovalScreen.MenuEvent(pVal, BubbleEvent)
                Case "ME_PDC"
                    objPDCPosting.MenuEvent(pVal, BubbleEvent)
                Case "ME_BR"
                    objBarcodePrinting.MenuEvent(pVal, BubbleEvent)
                Case "ME_WG"
                    objWeighingScale.MenuEvent(pVal, BubbleEvent)
                Case "ME_CLEAR"
                    objClearanceSales.MenuEvent(pVal, BubbleEvent)
                Case "ME_PSC"
                    objPDTStockCounting.MenuEvent(pVal, BubbleEvent)
                Case "ME_SCGW"
                    ' objStockCountGenerationWizard.MenuEvent(pVal, BubbleEvent)
                    objStockwizard.MenuEvent(pVal, BubbleEvent)
                Case "TNX_LSINTSTS"
                    objLSIntegrationStatusScreen.MenuEvent(pVal, BubbleEvent)
                Case "SBO_PGW"
                    objProductionGenerationWizard.MenuEvent(pVal, BubbleEvent)
                Case "ME_FREEITM"
                    objGetfreeitems.MenuEvent(pVal, BubbleEvent)
                Case "ME_ITEMA"
                    objItemswise.MenuEvent(pVal, BubbleEvent)
                Case "ME_PRMONVPA"
                    objPromotionVoucherPurchase.MenuEvent(pVal, BubbleEvent)
                Case "ME_LEASTITMA"
                    objGetLeastWise.MenuEvent(pVal, BubbleEvent)
                Case "ME_BUYITEM"
                    objGetN.MenuEvent(pVal, BubbleEvent)

                Case "ME_PRMONTOTLA"
                    objPromotionBTotal.MenuEvent(pVal, BubbleEvent)
                Case "ME_DISCOUNT"
                    objDiscount.MenuEvent(pVal, BubbleEvent)
                Case "ME_PRICECONFIG"
                    objPriceConfiguration.MenuEvent(pVal, BubbleEvent)
                Case "SDPTR"
                    objSupDeposit.MenuEvent(pVal, BubbleEvent)
                Case "PCTMR"
                    objPettyCashMaster.MenuEvent(pVal, BubbleEvent)
                Case "GRPODWR"
                    objGRPODocRep.MenuEvent(pVal, BubbleEvent)
                Case "ALLITMS"
                    objAllItem.MenuEvent(pVal, BubbleEvent)
                Case "BARCHK"
                    objBarCheck.MenuEvent(pVal, BubbleEvent)
                Case "GRPOITM"
                    objGRPOItem.MenuEvent(pVal, BubbleEvent)
                Case "STCKCNT"
                    objSTCKCNT.MenuEvent(pVal, BubbleEvent)
                Case "MWNEDSR"
                    objMainWHSE.MenuEvent(pVal, BubbleEvent)
                Case "OPDSPR"
                    objOpendisp.MenuEvent(pVal, BubbleEvent)
                Case "TRNSRP"
                    objTransferRep.MenuEvent(pVal, BubbleEvent)
                Case "STKWHSE"
                    objStochWHSE.MenuEvent(pVal, BubbleEvent)
                Case "ARINR"
                    objARDoc.MenuEvent(pVal, BubbleEvent)
                Case "ARITMR"
                    objARItem.MenuEvent(pVal, BubbleEvent)
                Case "APINR"
                    objAPItem.MenuEvent(pVal, BubbleEvent)
                Case "CATWIN"
                    objCatIN.MenuEvent(pVal, BubbleEvent)
                Case "EXITM"
                    objExciseItems.MenuEvent(pVal, BubbleEvent)

                Case "BRANDMAS"
                    objBrandMaster.MenuEvent(pVal, BubbleEvent)
                Case "SGMASTER"
                    objSubGroupMaster.MenuEvent(pVal, BubbleEvent)
                Case "SSGMASTER"
                    objSubSubGroupMaster.MenuEvent(pVal, BubbleEvent)
                Case "SSSGMASTER"
                    objSubSubSubGroupMaster.MenuEvent(pVal, BubbleEvent)
                Case "DMASTER"
                    objDepartmentMaster.MenuEvent(pVal, BubbleEvent)
                Case "SBO_SETCONFIG"
                    MessengerConfiguration.MenuEvent(pVal, BubbleEvent)



                Case "MGRCONFIG"
                    MessageConfiguration.MenuEvent(pVal, BubbleEvent)

                Case "Report"
                    objListReport.MenuEvent(pVal, BubbleEvent)
                Case "Category"
                    objCategorySplit.MenuEvent(pVal, BubbleEvent)
                Case "BCR"
                    objBarcodeRequest.MenuEvent(pVal, BubbleEvent)
                    'Approval
                Case "ME_ASR"
                    ObjApproval.MenuEvent(pVal, BubbleEvent)
                Case "GRID"
                    'ObjGRID.MenuEvent(pVal, BubbleEvent)
                Case "ME_ADR"
                    ObjADR.MenuEvent(pVal, BubbleEvent)
                Case "SBO_AST"
                    ObjAStg.MenuEvent(pVal, BubbleEvent)
                Case "Temp"
                    ObjAPTEMP.MenuEvent(pVal, BubbleEvent)
                Case "GRIDES"
                    ObjGRIDES.MenuEvent(pVal, BubbleEvent)
                Case "PCRT"
                    ObjPriceChangeRequest.MenuEvent(pVal, BubbleEvent)

                Case "TNX_MULSTR"
                    objMultiplestr.MenuEvent(pVal, BubbleEvent)
                Case "ADDONLOG"
                    'objAddonLog.MenuEvent(pVal, BubbleEvent)

                    'madhu

                'Case "1282"
                '    objform = objMain.objApplication.Forms.ActiveForm
                '    If objform.TypeEx = "134" Then
                '        objBusinessPartners.MenuEvent(pVal, BubbleEvent)
                '    ElseIf objform.TypeEx = "PCRT" Then
                '        ObjPriceChangeRequest.MenuEvent(pVal, BubbleEvent)
                '    End If

                Case "EMLLog"
                    EmailHistory.MenuEvent(pVal, BubbleEvent)
                Case "CONFIG"
                    Configuration.MenuEvent(pVal, BubbleEvent)
                Case "ME_MIGNWZRD"
                    If pVal.BeforeAction = False Then
                        Dim GetMIUsers As String = "CALL ""SBO_GetMIUser""('" & objMain.objCompany.UserName & "') "
                        Dim oRsGetMIUsers As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetMIUsers.DoQuery(GetMIUsers)
                        If oRsGetMIUsers.RecordCount > 0 Then
                            objMIGenerationWizard.MenuEvent(pVal, BubbleEvent)
                        Else
                            objMain.objApplication.StatusBar.SetText("User not authorized....")
                        End If
                    End If
                Case "1281"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "PCRT" Then
                        ObjPriceChangeRequest.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "Branch" Then
                        ObjBranchMaster.MenuEvent(pVal, BubbleEvent)
                    End If

                Case "1282"
                    objform = objMain.objApplication.Forms.ActiveForm

                    If objform.TypeEx = "ME_MI_Form" Then
                        objMaterialInward.MenuEvent(pVal, BubbleEvent)

                        'Akhila
                    ElseIf objform.TypeEx = "STOCKDE" Then
                        objClsStockDetails.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "PACP" Then
                        objPACPPriceChange.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "SBAMASTER" Then
                        objSalesBlanketAgreementMaster.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "PRODEN" Then
                        objProductEnquiry.MenuEvent(pVal, BubbleEvent)


                    ElseIf objform.TypeEx = "ME_PDC_Form" Then
                        objPDCPosting.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "PRMONVP" Then
                        objPromotionVoucherPurchase.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "TNXCRDNT_Form" Then
                        objUserCreditNote.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "PriceList" Then
                        ObjPriceUpdateService.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "PCRT" Then
                        ObjPriceChangeRequest.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "Branch" Then
                        ObjBranchMaster.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "ME_STR_Form" Then
                        objStockTransferRequest.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "ITEM" Then
                        objItemswise.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "ME_AP_Form" Then
                        objAutomaticProduction.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "PRICECONFIG" Then
                        objPriceConfiguration.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "ME_PSC_Form" Then
                        objPDTStockCounting.MenuEvent(pVal, BubbleEvent)
                        'ElseIf objform.TypeEx = "6" Then
                        '    objform.Close()
                    ElseIf objform.TypeEx = "142" Then
                        objPurchaseOrder.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "139" Then
                        'objSalesOrder.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "143" Then
                        objGRPO.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "1470000200" Then
                        objPurchaserequest.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "PRMONTOTL" Then
                        objPromotionBTotal.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "SDPTR" Then
                        objSupDeposit.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "PCTMR" Then
                        objPettyCashMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "ME_SCGW" Then
                        objStockwizard.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "ITEMUP" Then
                        Itemmasterdataupload.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "150" Then
                        objItemMaster.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "BarcodeRequest" Then
                        objBarcodeRequest.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "134" Then
                        objBusinessPartners.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "TNX_POPost" Then
                        objMultiplestr.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "1250000100" Then
                        objRentalAgreement.MenuEvent(pVal, BubbleEvent)
                    End If

                Case "1283"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "3002" Then
                        objGRPODRAFT.MenuEvent(pVal, BubbleEvent)

                    End If
                Case "Close"
                    'Try
                    objform = objMain.objApplication.Forms.ActiveForm
                    'Catch ex As Exception
                    'End Try
                    If objform.TypeEx = "ME_STR_Form" Then
                        objStockTransferRequest.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "ME_PSC_Form" Then
                        objPDTStockCounting.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "Clear" Then
                        objClearanceSales.MenuEvent(pVal, BubbleEvent)
                    End If
                Case "1292"
                    'Try
                    objform = objMain.objApplication.Forms.ActiveForm
                    'Catch ex As Exception
                    'End Try

                    If objform.TypeEx = "PRMONTOTL" Then
                        objPromotionBTotal.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "ME_PSC_Form" Then
                        objPDTStockCounting.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "PCRT" Then
                        ObjPriceChangeRequest.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "Branch" Then
                        ObjBranchMaster.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "SBO_AST" Then
                        ObjAStg.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "SBAMASTER" Then
                        objSalesBlanketAgreementMaster.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "Temp" Then
                        ObjAPTEMP.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "POSBILL" Then
                        objClsPOSBill.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "PACP" Then
                        objPACPPriceChange.MenuEvent(pVal, BubbleEvent)



                    ElseIf objform.TypeEx = "Category" Then
                        objCategorySplit.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "ME_SUBITMLST_Form" Then
                        objSubItemList.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "BarcodeRequest" Then
                        objBarcodeRequest.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "STOCKDE" Then
                        objClsStockDetails.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "S_Rebate" Then
                        objSupplierRebate.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "PRODEN" Then
                        objProductEnquiry.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "BUYITEM" Then
                        objGetN.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "LEASTITM" Then
                        objGetLeastWise.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "FREEITM" Then
                        objGetfreeitems.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "ITEM" Then
                        objItemswise.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "ME_ITMLIST" Then
                        objItemList.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "ME_SUP_Form" Then
                        objSupplAgreememntPosting.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "ME_ITMLIST" Then
                        objItemList.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "SDPTR" Then
                        objSupDeposit.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "PCTMR" Then
                        objPettyCashMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "PriceList" Then
                        ObjPriceUpdateService.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "ITEMUP" Then
                        Itemmasterdataupload.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "DMASTER" Then
                        objDepartmentMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "SSSGMASTER" Then
                        objSubSubSubGroupMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "SSGMASTER" Then
                        objSubSubGroupMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "SGMASTER" Then
                        objSubGroupMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "BRANDMAS" Then
                        objBrandMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "ME_STR_Form" Then
                        objStockTransferRequest.MenuEvent(pVal, BubbleEvent)

                    End If
                Case "774"
                    objform = objMain.objApplication.Forms.ActiveForm

                    If objform.TypeEx = "Category" Then
                        objCategorySplit.MenuEvent(pVal, BubbleEvent)
                    End If

                    If objform.TypeEx = "PACP" Then
                        objPACPPriceChange.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "SBAMASTER" Then
                        objSalesBlanketAgreementMaster.MenuEvent(pVal, BubbleEvent)
                    End If

                    If objform.TypeEx = "PRODEN" Then
                        objProductEnquiry.MenuEvent(pVal, BubbleEvent)
                    End If

                    If objform.TypeEx = "ME_SUBITMLST_Form" Then
                        objSubItemList.MenuEvent(pVal, BubbleEvent)
                    End If

                    If objform.TypeEx = "SBO_AST" Then
                        ObjAStg.MenuEvent(pVal, BubbleEvent)
                    End If

                    If objform.TypeEx = "BarcodeRequest" Then
                        objBarcodeRequest.MenuEvent(pVal, BubbleEvent)
                    End If

                    If objform.TypeEx = "S_Rebate" Then
                        objSupplierRebate.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "STOCKDE" Then
                        objClsStockDetails.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "POSBILL" Then
                        objClsPOSBill.MenuEvent(pVal, BubbleEvent)
                    End If

                    'Approval
                    If objform.TypeEx = "Temp" Then
                        ObjAPTEMP.MenuEvent(pVal, BubbleEvent)


                    End If
                Case "Delete Row"
                    'Try
                    objform = objMain.objApplication.Forms.ActiveForm
                    'Catch ex As Exception
                    'End Try
                    If objform.TypeEx = "ME_STR_Form" Then
                        objStockTransferRequest.MenuEvent(pVal, BubbleEvent)
                    End If

                    If objform.TypeEx = "TNXCRDNT_Form" Then
                        objUserCreditNote.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "ME_ITMLIST" Then
                        objItemList.MenuEvent(pVal, BubbleEvent)
                    End If


                    If objform.TypeEx = "ME_AP_Form" Then
                        objAutomaticProduction.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "ME_SUP_Form" Then
                        objSupplAgreememntPosting.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "Clear" Then
                        objClearanceSales.MenuEvent(pVal, BubbleEvent)
                    End If

                    If objform.TypeEx = "ITEM" Then
                        objItemswise.MenuEvent(pVal, BubbleEvent)
                    End If

                    If objform.TypeEx = "ME_PDS_Form" Then
                        objPaymentDetails.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "ME_PSC_Form" Then
                        objPDTStockCounting.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "SDPTR" Then
                        objSupDeposit.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "PCTMR" Then
                        objPettyCashMaster.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "PCRT" Then
                        ObjPriceChangeRequest.MenuEvent(pVal, BubbleEvent)
                    End If

                    If objform.TypeEx = "BarcodeRequest" Then
                        objBarcodeRequest.MenuEvent(pVal, BubbleEvent)
                    End If

                    If objform.TypeEx = "ITEMUP" Then
                        Itemmasterdataupload.MenuEvent(pVal, BubbleEvent)
                    End If

                    If objform.TypeEx = "1250000100" Then
                        objRentalAgreement.MenuEvent(pVal, BubbleEvent)
                    End If

                Case "Add Row"
                    Try
                        objform = objMain.objApplication.Forms.ActiveForm

                        If objform IsNot Nothing AndAlso objform.TypeEx = "1250000100" Then
                            objRentalAgreement.MenuEvent(pVal, BubbleEvent)
                        End If

                        Dim handled As Boolean = False

                        Try
                            objform = objApplication.Forms.ActiveForm

                            If objform IsNot Nothing AndAlso objform.TypeEx = "ME_PSC_Form" Then
                                objPDTStockCounting.MenuEvent(pVal, BubbleEvent)
                                handled = True
                            End If

                        Catch
                        End Try

                        If Not handled AndAlso objPDTStockCounting.objForm IsNot Nothing Then
                            objPDTStockCounting.MenuEvent(pVal, BubbleEvent)
                        End If

                    Catch ex As Exception
                    End Try


                'Case "Add Row"
                '    objform = objMain.objApplication.Forms.ActiveForm
                '    If objform.TypeEx = "1250000100" Then
                '        objRentalAgreement.MenuEvent(pVal, BubbleEvent)
                '    End If

                '    ' Case "Add Row"
                '    Try
                '        Dim handled As Boolean = False
                '        Try
                '            objform = objApplication.Forms.ActiveForm
                '            If objform IsNot Nothing AndAlso objform.TypeEx = "ME_PSC_Form" Then
                '                objPDTStockCounting.MenuEvent(pVal, BubbleEvent)
                '                handled = True
                '            End If
                '        Catch
                '        End Try

                '        If Not handled AndAlso objPDTStockCounting.objForm IsNot Nothing Then
                '            objPDTStockCounting.MenuEvent(pVal, BubbleEvent)
                '        End If
                '    Catch ex As Exception
                '    End Try

                Case "519"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "ME_MI_Form" Then
                        objMaterialInward.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "ME_STR_Form" Then
                        objStockTransferRequest.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "ME_SUP_Form" Then
                        objSupplAgreememntPosting.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "Clear" Then
                        objClearanceSales.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "ME_AP_Form" Then
                        objAutomaticProduction.MenuEvent(pVal, BubbleEvent)
                    End If

                    If objform.TypeEx = "TNXCRDNT_Form" Then
                        objUserCreditNote.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "ME_ITMLIST" Then
                        objItemList.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "PCRT" Then
                        ObjPriceChangeRequest.MenuEvent(pVal, BubbleEvent)
                    End If

                    If objform.TypeEx = "POSBILL" Then
                        objClsPOSBill.MenuEvent(pVal, BubbleEvent)
                    End If
                Case "1293"
                    objform = objMain.objApplication.Forms.ActiveForm

                    If objform.TypeEx = "Branch" Then
                        ObjBranchMaster.MenuEvent(pVal, BubbleEvent)
                    End If

                    If objform.TypeEx = "PCRT" Then
                        ObjPriceChangeRequest.MenuEvent(pVal, BubbleEvent)
                    End If

                    If objform.TypeEx = "1250000100" Then
                        objRentalAgreement.MenuEvent(pVal, BubbleEvent)
                    End If

                Case "1284"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "PCRT" Then
                        ObjPriceChangeRequest.MenuEvent(pVal, BubbleEvent)
                    End If
                Case "520"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "ME_MI_Form" Then
                        objMaterialInward.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "ME_SUP_Form" Then
                        objSupplAgreememntPosting.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "ME_STR_Form" Then
                        objStockTransferRequest.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "Clear" Then
                        objClearanceSales.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "ITEMLIST" Then
                        objListItemPR.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "ME_ITMLIST" Then
                        objItemList.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "PCRT" Then
                        ObjPriceChangeRequest.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "POSBILL" Then
                        objClsPOSBill.MenuEvent(pVal, BubbleEvent)
                    End If

                Case "1293"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "ME_SUBITMLST_Form" Then
                        objSubItemList.MenuEvent(pVal, BubbleEvent)
                    End If
            End Select


        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

#Region "Form Data Event"
    Private Sub objApplication_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles objApplication.FormDataEvent

        'Select Case BusinessObjectInfo.FormTypeEx
        '    Case "ME_MI_Form"
        '        objMaterialInward.FormDataEvent(BusinessObjectInfo, BubbleEvent)
        'End Select

        If BusinessObjectInfo.BeforeAction = False Then
            Select Case BusinessObjectInfo.FormTypeEx


                Case "TNXCRDNT_Form"
                    objUserCreditNote.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                Case "MGRCONFIG"
                    MessageConfiguration.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                Case "ME_SUPAG_Form"
                    objSupplierAgreement.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "ME_SUP_Form"
                    objSupplAgreememntPosting.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "940"
                    objInventoryTransfer.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                    'Case "ME_MI_Form"
                    '    objMaterialInward.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "ME_STR_Form"
                    objStockTransferRequest.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "ME_ITMLIST"
                    objItemList.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "720"
                    objGoodsIssue.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "134"
                    objBusinessPartners.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "142"
                    objPurchaseOrder.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "139"
                    objSalesOrder.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "170"
                    objIncomingpayments.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "1470000200"
                    objPurchaserequest.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "143"
                    objGRPO.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "141"
                    objPurchaseInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "721"
                    objGoodsReceipt.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                    'Case "1470000020"
                    '    objBarcodeMaster.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "150"
                    objItemMaster.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "1250000102"
                    objPurchaseBlanketAgreement.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "1250000100"
                    objRentalAgreement.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "ME_PRD_Form"
                    objPurchaseRequestDetails.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "S_Rebate"
                    objSupplierRebate.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                Case "PACP"
                    objPACPPriceChange.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                Case "PCRT"
                    ObjPriceChangeRequest.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                Case "ITEMUP"
                    Itemmasterdataupload.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                Case "Me_Outlet"
                    objOutlet.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                Case "SBAMASTER"
                    objSalesBlanketAgreementMaster.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                Case "IMPPerview"
                    objItemMasterUploadPreview.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                Case "BarcodeRequest"
                    objBarcodeRequest.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "426"
                    objOutgoingPayments.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "ME_PRICEUPD_Form"


                    objPriceUpdate.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "ME_PDC_Form"
                    objPDCPosting.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "frmBarCodeIM"
                    objBarcodeIM.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "Clear"
                    objClearanceSales.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "ME_PDS_Form"
                    objPaymentDetails.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "ME_SCGW"
                    objStockwizard.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "ME_FREEITM"
                    objGetfreeitems.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "ITEM"
                    objItemswise.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "ME_LEASTITM"
                    objGetLeastWise.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "ME_BUYITEM"
                    objGetN.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "ME_PRMONTOTL"
                    objPromotionBTotal.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "ME_DISCOUNT"
                    objDiscount.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "ME_AP_Form"
                    objAutomaticProduction.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "PRICECONFIG"
                    objPriceConfiguration.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                Case "140"
                    objDelivery.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                Case "181"
                    objAPCreditmemo.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                Case "179"
                    objARCreditnote.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                Case "133"
                    objARInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                Case "182"
                    objGoodsReturn.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                Case "180"
                    objSaleReturn.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "PriceList"
                    ObjPriceUpdateService.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                Case "60091"
                    objARInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)



            End Select
        End If
    End Sub
#End Region

#Region "Application Event"
    Private Sub oApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles objApplication.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged, SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_ShutDown
                objCompany.Disconnect()
                End
        End Select
    End Sub
#End Region

#Region "Right Click Event"
    Private Sub objApplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles objApplication.RightClickEvent
        Dim objForm As SAPbouiCOM.Form
        'price chnage request
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)


        If eventInfo.BeforeAction Then
            objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)

            If objForm.TypeEx = "PCRT" Then
                objForm.EnableMenu("1284", True)
            End If
        End If

        ' Create / remove Delete Row menu for Price Change Request matrices (Mtx1..Mtx5)
        Try
            oMenuItem = objMain.objApplication.Menus.Item("1280") ' Data menu
            oMenus = oMenuItem.SubMenus

            ' target matrix UIDs to enable Delete Row on
            Dim matrices As String() = New String() {"Mtx1", "Mtx2", "Mtx3", "Mtx4", "Mtx5"}

            Dim isTargetMatrix As Boolean = False
            If Not String.IsNullOrEmpty(eventInfo.ItemUID) Then
                For Each m As String In matrices
                    If String.Compare(eventInfo.ItemUID, m, True) = 0 Then
                        isTargetMatrix = True
                        Exit For
                    End If
                Next
            End If

            ' If user right-clicked on any of the matrices' LineID column -> add menu
            If isTargetMatrix AndAlso String.Compare(eventInfo.ColUID, "LineID", True) = 0 Then
                If Not oMenus.Exists("Delete Row") Then
                    oCreationPackage.UniqueID = "Delete Row"
                    oCreationPackage.String = "Delete Row"
                    oCreationPackage.Enabled = True
                    oMenus.AddEx(oCreationPackage)
                End If
            Else
                ' otherwise remove the menu if present
                If oMenus.Exists("Delete Row") Then
                    objMain.objApplication.Menus.RemoveEx("Delete Row")
                End If
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

        objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)
        If objForm.TypeEx = "ME_STR_Form" Then
            objStockTransferRequest.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "TNXCRDNT_Form" Then
            objUserCreditNote.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "S_Rebate" Then
            objSupplierRebate.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "ME_AP_Form" Then
            objAutomaticProduction.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "Clear" Then
            objClearanceSales.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "ME_PDS_Form" Then
            objPaymentDetails.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "ME_PSC_Form" Then
            objPDTStockCounting.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "PRMONTOTL" Then
            objPromotionBTotal.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "BUYITEM" Then
            objGetN.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "LEASTITM" Then
            objGetLeastWise.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "FREEITM" Then
            objGetLeastWise.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "ITEM" Then
            objItemswise.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "ME_ITMLIST" Then
            objItemList.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "ME_SUP_Form" Then
            objSupplAgreememntPosting.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "ITEMUP" Then
            Itemmasterdataupload.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "PCRT" Then
            ObjPriceChangeRequest.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "1250000100" Then
            objRentalAgreement.RightClickEvent(eventInfo, BubbleEvent)

            'ElseIf objForm.TypeEx = "PCTMR" Then
            '    objPettyCashMaster.RightClickEvent(eventInfo, BubbleEvent)
            'ElseIf objForm.TypeEx = "SDPTR" Then
            '    objSupDeposit.RightClickEvent(eventInfo, BubbleEvent)
        End If
    End Sub
#End Region

#Region "Hibernate Event"
    Private Sub PowerModeChanged(ByVal sender As System.Object,
        ByVal e As Microsoft.Win32.PowerModeChangedEventArgs)
        Select Case e.Mode
            Case Microsoft.Win32.PowerModes.Resume
                For Each prog As Process In Process.GetProcesses
                    If prog.ProcessName = "ProcessName" Then
                        prog.Kill()
                    End If
                Next
            Case Microsoft.Win32.PowerModes.Suspend
                For Each prog As Process In Process.GetProcesses
                    If prog.ProcessName = "ProcessName" Then
                        prog.Kill()
                    End If
                Next
        End Select
    End Sub
    Private Sub SessionEnding(ByVal sender As System.Object,
ByVal e As Microsoft.Win32.SessionEndingEventArgs)
        Select Case e.Reason
            Case Microsoft.Win32.SessionEndReasons.Logoff
                For Each prog As Process In Process.GetProcesses
                    If prog.ProcessName = "ProcessName" Then
                        prog.Kill()
                    End If
                Next
            Case Microsoft.Win32.SessionEndReasons.SystemShutdown
                For Each prog As Process In Process.GetProcesses
                    If prog.ProcessName = "ProcessName" Then
                        prog.Kill()
                    End If
                Next
        End Select
    End Sub
#End Region


    'Public Function licence1(ByVal Addonname As String, ByVal DBName As String, ByVal User As String) As Boolean
    '    If objMain.objCompany.Connected Then
    '        Dim rsConfig1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        Dim query As String = "Select T0.""U_License"" as ""licence"", ""U_HKey"" as ""Hardwarekey"",T0.""U_Company"", T0.""U_DB"" As ""Dbname"",T0.""U_Addon"" as ""Addon"", T0.""U_Total"" as ""Total""," &
    '                          "T1.""U_Code"" as ""User"", T1.""U_Sts"" as ""Status"", " &
    '                          "T0.""U_SDate"" as ""fromdate"", T0.""U_EDate"" as ""Todate"" " &
    '                          "from ""@TNX_LICENCE"" T0 " &
    '                          "INNER JOIN ""@TNX_LICENCE_C0"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
    '                          "WHERE T1.""U_Sts"" = 'Y' AND T1.""U_Code"" ='" & User & "' AND T0.""U_Addon"" ='" & Addonname & "' " &
    '                          "AND T0.""U_DB"" ='" & DBName & "' and CURRENT_DATE BETWEEN T0.""U_SDate"" AND T0.""U_EDate"""
    '        rsConfig1.DoQuery(query)
    '        Dim licence As String = rsConfig1.Fields.Item("licence").Value
    '        Dim FDate As DateTime = rsConfig1.Fields.Item("fromdate").Value
    '        Dim toDate As DateTime = rsConfig1.Fields.Item("Todate").Value
    '        Dim currentDate As DateTime = DateTime.Now
    '        Dim user2 As String = rsConfig1.Fields.Item("User").Value
    '        Dim Addonname1 As String = rsConfig1.Fields.Item("Addon").Value
    '        Dim dbname1 As String = rsConfig1.Fields.Item("Dbname").Value
    '        If rsConfig1.RecordCount > 0 And licence <> "" Then
    '            Dim licenseStatus As String = rsConfig1.Fields.Item("Status").Value
    '            If licenseStatus = "Y" Then
    '                Dim License, Database, Company, HKey, TotalLicense As String
    '                Dim StartDate, ExpDate As Date
    '                License = rsConfig1.Fields.Item("licence").Value
    '                Company = rsConfig1.Fields.Item("U_Company").Value
    '                Database = rsConfig1.Fields.Item("Dbname").Value
    '                HKey = rsConfig1.Fields.Item("Hardwarekey").Value
    '                TotalLicense = rsConfig1.Fields.Item("Total").Value
    '                Dim SDate As String = CDate(rsConfig1.Fields.Item("fromdate").Value).ToString("yyyyMMdd")
    '                Dim EDate As String = CDate(rsConfig1.Fields.Item("Todate").Value).ToString("yyyyMMdd")
    '                SDate = SDate.Insert(4, "/")
    '                SDate = SDate.Insert(7, "/")
    '                EDate = EDate.Insert(4, "/")
    '                EDate = EDate.Insert(7, "/")
    '                'StartDate = CDate(SDate)
    '                'ExpDate = CDate(EDate)
    '                'Dim SDate As String = rsConfig1.Fields.Item("fromdate").Value
    '                'Dim EDate As String = rsConfig1.Fields.Item("Todate").Value
    '                ' Correct date formats
    '                StartDate = SDate
    '                ExpDate = EDate

    '                Dim uniKey As String = "10X_SAP_B!_"
    '                Dim strKey As String = StartDate.ToString("dd-MMM-yyyy") & "|" & ExpDate.ToString("dd-MMM-yyyy") & "|" & Company & "|" & TotalLicense & "|" & Database & "|" & HKey & "|" & Addonname
    '                Dim TestLicense As String = GetSHA1HashData(strKey + GetSHA1HashData(uniKey))
    '                Dim GenLicense As String = GetSHA1HashData(strKey + GetSHA1HashData(uniKey))
    '                ' Check if the license is valid
    '                If License <> GetSHA1HashData(strKey + GetSHA1HashData(uniKey)) Then
    '                    objMain.objApplication.StatusBar.SetText("Invalid license file", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                    Return False
    '                Else
    '                    '  Me.SBO_Application.StatusBar.SetText("Licence Tenx is valid", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    '                    Return True
    '                End If
    '            End If
    '        Else
    '            objMain.objApplication.StatusBar.SetText("License Tenx is invalid", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Return False
    '        End If
    '    End If
    'End Function

    Public Function Addonlicence(ByVal Addonname As String, ByVal DBName As String, ByVal User As String) As Boolean
        Dim rsConfig1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            If objMain.objCompany.Connected Then

                Dim query As String = "Select T0.""U_License"" as ""licence"", ""U_HKey"" as ""Hardwarekey"",T0.""U_Company"", T0.""U_DB"" As ""Dbname"",T0.""U_Addon"" as ""Addon"", T0.""U_Total"" as ""Total""," &
                                      "T1.""U_Code"" as ""User"", T1.""U_Sts"" as ""Status"", " &
                                      "T0.""U_SDate"" as ""fromdate"", T0.""U_EDate"" as ""Todate"" " &
                                      "from ""@TNX_LICENCE"" T0 " &
                                      "INNER JOIN ""@TNX_LICENCE_C0"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" " &
                                      "WHERE T1.""U_Sts"" = 'Y' AND T1.""U_Code"" ='" & User & "' AND T0.""U_Addon"" ='" & Addonname & "' " &
                                      "AND T0.""U_DB"" ='" & DBName & "' and CURRENT_DATE BETWEEN T0.""U_SDate"" AND T0.""U_EDate"""
                rsConfig1.DoQuery(query)

                If rsConfig1.RecordCount = 0 Then
                    objMain.objApplication.StatusBar.SetText("Invalid 10X license file, Please Contact 10X Team", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                Dim licence As String = rsConfig1.Fields.Item("licence").Value
                If String.IsNullOrEmpty(licence) Then
                    objMain.objApplication.StatusBar.SetText("Invalid 10X license file, Please Contact 10X Team", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                If rsConfig1.Fields.Item("Status").Value <> "Y" Then
                    objMain.objApplication.StatusBar.SetText("Invalid 10X license file, Please Contact 10X Team", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                Dim Company As String = rsConfig1.Fields.Item("U_Company").Value
                Dim Database As String = rsConfig1.Fields.Item("Dbname").Value
                Dim HKey As String = rsConfig1.Fields.Item("Hardwarekey").Value
                Dim TotalLicense As String = rsConfig1.Fields.Item("Total").Value
                Dim Addon As String = rsConfig1.Fields.Item("Addon").Value

                Dim StartDate As Date = CDate(rsConfig1.Fields.Item("fromdate").Value)
                Dim ExpDate As Date = CDate(rsConfig1.Fields.Item("Todate").Value)

                Dim SDate As String = StartDate.ToString("yyyy/MM/dd")
                Dim EDate As String = ExpDate.ToString("yyyy/MM/dd")

                Dim uniKey As String = "10X_SAP_B!_"
                Dim strKey As String = StartDate.ToString("dd-MMM-yyyy") & "|" & ExpDate.ToString("dd-MMM-yyyy") & "|" & Company & "|" & TotalLicense & "|" & Database & "|" & HKey & "|" & Addonname
                Dim generatedLicense As String = GetSHA1HashData(strKey + GetSHA1HashData(uniKey))

                If licence <> generatedLicense Then
                    objMain.objApplication.StatusBar.SetText("Invalid license file, Please Contact 10X Team", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                Return True
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("10X License check error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            If rsConfig1 IsNot Nothing Then
                Marshal.ReleaseComObject(rsConfig1)
                rsConfig1 = Nothing
            End If
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try

    End Function

    Private Sub LoadItemsToSupplierRebate(ByVal FormUID As String,
                                      ByVal BpCode As String,
                                      ByVal ParentDocEntry As String)

        Dim objForm As SAPbouiCOM.Form =
        objMain.objApplication.Forms.Item(FormUID)

        Dim oMatrix As SAPbouiCOM.Matrix =
        CType(objForm.Items.Item("Item_10").Specific, SAPbouiCOM.Matrix)

        Dim oDBs As SAPbouiCOM.DBDataSource =
        objForm.DataSources.DBDataSources.Item("@SBO_SUPREBATE_C2")

        Dim oRs As SAPbobsCOM.Recordset =
        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim sQuery As String =
        "SELECT T1.""ItemCode"", T1.""ItemName"", T1.""PlanQty"", T1.""UomCode"", T2.""CodeBars"" " &
        "FROM ""OOAT"" T0 " &
        "INNER JOIN ""OAT1"" T1 ON T0.""AbsID"" = T1.""AgrNo"" " &
        "INNER JOIN ""OITM"" T2 ON T1.""ItemCode"" = T2.""ItemCode"" " &
        "WHERE T0.""BpCode"" = '" & BpCode & "' " &
        "AND IFNULL(T2.""U_REBEGBLE"", 'No') = 'Yes'"

        oRs.DoQuery(sQuery)

        If oRs.RecordCount = 0 Then Exit Sub

        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
            objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
        End If

        oMatrix.Clear()
        oDBs.Clear()

        Dim i As Integer = 0
        oRs.MoveFirst()

        While Not oRs.EoF
            oDBs.InsertRecord(i)

            oDBs.SetValue("U_ITMCD", i, oRs.Fields.Item("ItemCode").Value.ToString())
            oDBs.SetValue("U_ITMNM", i, oRs.Fields.Item("ItemName").Value.ToString())
            oDBs.SetValue("U_UOM", i, oRs.Fields.Item("UomCode").Value.ToString())
            oDBs.SetValue("U_BRCD", i, oRs.Fields.Item("CodeBars").Value.ToString())
            oDBs.SetValue("U_QTY", i, oRs.Fields.Item("PlanQty").Value.ToString())

            oRs.MoveNext()
            i += 1
        End While

        oMatrix.LoadFromDataSource()
        oMatrix.AutoResizeColumns()

    End Sub


    Private Function GetSHA1HashData(ByVal data As String) As String
        Dim sha2 As SHA1 = sha2.Create()
        Dim hashData As Byte() = sha2.ComputeHash(Encoding.[Default].GetBytes(data))
        Dim returnValue As System.Text.StringBuilder = New StringBuilder()
        For i As Integer = 0 To hashData.Length - 1
            returnValue.Append(hashData(i).ToString())
        Next
        Return returnValue.ToString()
    End Function
    Public Function ConvertFileToBase64(ByVal fileName As String) As String
        Return Convert.ToBase64String(System.IO.File.ReadAllBytes(fileName))
    End Function



    Sub MarketingWhatsAppMessage(ByVal DocNum As String, ByVal Series As String, ByVal DocType As String)
        Try
            Dim Template As String = ""
            Dim Msg As String = ""
            Dim DocEntry As String = ""
            Dim WMSGTXT As String = ""
            Dim MNTXT As String = ""
            Dim DetQry As String = ""

            Template = objMain.objUtilities.GetTemplateName(DocType, "WhatsApp")
            Msg = objMain.objUtilities.getMessage(DocType, "WhatsApp")
            If Msg <> "" Then
                Dim Rset As SAPbobsCOM.Recordset
                Try
                    If DocType = "SO" Then
                        DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingWhatsAppMessage','SO','" + DocNum + "','" + Series + "')"
                    ElseIf DocType = "PO" Then
                        DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingWhatsAppMessage','PO','" + DocNum + "','" + Series + "')"
                    ElseIf DocType = "SQ" Then
                        DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingWhatsAppMessage','SQ','" + DocNum + "','" + Series + "')"
                    ElseIf DocType = "DL" Then
                        DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingWhatsAppMessage','DL','" + DocNum + "','" + Series + "')"
                    ElseIf DocType = "ARINV" Then
                        DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingWhatsAppMessage','ARINV','" + DocNum + "','" + Series + "')"
                    ElseIf DocType = "ARRINV" Then
                        DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingWhatsAppMessage','ARINV','" + DocNum + "','" + Series + "')"
                    ElseIf DocType = "ARC" Then
                        DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingWhatsAppMessage','ARC','" + DocNum + "','" + Series + "')"
                    ElseIf DocType = "ARRET" Then
                        DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingWhatsAppMessage','ARRET','" + DocNum + "','" + Series + "')"
                    ElseIf DocType = "PQ" Then
                        DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingWhatsAppMessage','PQ','" + DocNum + "','" + Series + "')"
                    ElseIf DocType = "APINV" Then
                        DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingWhatsAppMessage','APINV','" + DocNum + "','" + Series + "')"
                    ElseIf DocType = "GRPO" Then
                        DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingWhatsAppMessage','GRPO','" + DocNum + "','" + Series + "')"
                    ElseIf DocType = "GR" Then
                        DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingWhatsAppMessage','GR','" + DocNum + "','" + Series + "')"
                    ElseIf DocType = "APC" Then
                        DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingWhatsAppMessage','APC','" + DocNum + "','" + Series + "')"
                    ElseIf DocType = "IP" Then
                        DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingWhatsAppMessage','IP','" + DocNum + "','" + Series + "')"
                    ElseIf DocType = "OP" Then
                        DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingWhatsAppMessage','OP','" + DocNum + "','" + Series + "')"
                    ElseIf DocType = "GRP" Then
                        DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingWhatsAppMessage','GRP','" + DocNum + "','" + Series + "')"
                    ElseIf DocType = "GI" Then
                        DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingWhatsAppMessage','GI','" + DocNum + "','" + Series + "')"
                    ElseIf DocType = "ITR" Then
                        DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingWhatsAppMessage','ITR','" + DocNum + "','" + Series + "')"
                    ElseIf DocType = "IT" Then
                        DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingWhatsAppMessage','IT','" + DocNum + "','" + Series + "')"
                    ElseIf DocType = "JE" Then
                        DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingWhatsAppMessage','JE','" + DocNum + "','" + Series + "')"
                    End If

                    Rset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Rset.DoQuery(DetQry)

                    DocEntry = Rset.Fields.Item("DocEntry").Value.ToString().Trim()

                    Msg = Msg.Replace("{Series}", Series)
                    Msg = Msg.Replace("{DocNum}", DocNum)

                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText("Error:  " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    objMain.objUtilities.Write_Log(ex.Message)
                End Try

                Try

                    Dim oRsDet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsDet.DoQuery(Msg)
                    If oRsDet.RecordCount > 0 Then
                        Dim Result As String = ""
                        objMain.objUtilities.ShowWarningMessage("")
                        Dim Mobile As String = oRsDet.Fields.Item("Mobile").Value

                        Mobile = Mobile.Replace(" ", "")
                        Mobile = Mobile.Replace("-", "")

                        Dim TotalBeforDis As Double = oRsDet.Fields.Item("DocTotal").Value + oRsDet.Fields.Item("DiscSum").Value - oRsDet.Fields.Item("VatSum").Value
                        Dim CardName As String = oRsDet.Fields.Item("CardName").Value.ToString()
                        Dim DocNum1 As String = oRsDet.Fields.Item("DocNum").Value.ToString()
                        Dim TaxDate As String = oRsDet.Fields.Item("TaxDate").Value.ToString()
                        Dim ItemCount As String = oRsDet.Fields.Item("ItemCount").Value.ToString()
                        Dim TotalBeforDis1 As String = TotalBeforDis.ToString()
                        Dim VatSum As String = oRsDet.Fields.Item("VatSum").Value.ToString()
                        Dim DocTotal As String = oRsDet.Fields.Item("DocTotal").Value.ToString()
                        objMain.objUtilities.ShowWarningMessage("WhatsApp Attachment generating in progress for Customer '" + CardName + "' , Please wait...")
                        Dim ObjTYpe As String
                        If DocType = "SQ" Then
                            ObjTYpe = "23"
                        ElseIf DocType = "ARRINV" Then
                            ObjTYpe = "13"
                        Else
                            ObjTYpe = ""
                        End If

                        Dim Attachment As String = Nothing
                        Dim Att As String = objUtilities.AttachmnetSelection(DocType, "WhatsApp")
                        If Att = 1 Then
                            Dim AttPath As String = objUtilities.GetAttachmnetPath()
                            Dim CrystalID As String = objUtilities.GetCrystalID(DocType, "WhatsApp")
                            objUtilities.GetManualAttachment(DocEntry, CrystalID, DocType, ObjTYpe)
                            Attachment = AttPath & DocType & "_" & DocEntry & "_" & System.DateTime.Now.ToString("yyyyMMdd") & ".pdf"
                        Else
                            If objMain.objUtilities.CheckAttachment(DocType, "WhatsApp") = True Then
                                Attachment = objMain.objUtilities.getAttachment(DocType, DocEntry)
                            End If
                        End If

                        If Mobile <> "" Then
                            Result = SendWhatsApp(Template, Mobile, DocType, WMSGTXT, Attachment, CardName, DocNum1, TaxDate, ItemCount, TotalBeforDis1, VatSum, DocTotal)
                            If Result = "Success" Then
                                Dim oDBDetail1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim TEXT As String = "WhatsApp message sent Successfull to " + Mobile + " for Customer " + CardName + "  !!!  "
                                TEXT = TEXT.Replace("'", "").Replace("""", "")
                                If TEXT.Length > 240 Then
                                    TEXT = TEXT.Substring(0, 240)
                                Else
                                    TEXT = TEXT
                                End If

                                Dim DetailsQry As String = "CALL ""TENX_MESSENGER_WHATSAPP_UPDATE"" ('Update','" + DocType + "','" + DocEntry + "','" + TEXT + "')"
                                oDBDetail1.DoQuery(DetailsQry)
                                objMain.objUtilities.Write_Log("WhatsApp message sent Successfull to " + Mobile + " for Customer " + CardName + "  !!! ")
                                objApplication.StatusBar.SetText("WhatsApp message sent Successfull to " + Mobile + " for Customer " + CardName + "  !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                objMain.objUtilities.AddDataToLogTable(DocType, TaxDate, CardName, "Success", Msg, "", Mobile, DocNum, DocEntry, Result, "WhatsApp")
                            Else
                                objMain.objUtilities.Write_Log("WhatsApp message sending Failed to " + Mobile + " for Customer " + CardName + "  !!!")
                                objApplication.StatusBar.SetText("WhatsApp message sending Failed to " + Mobile + " for Customer " + CardName + "  !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objMain.objUtilities.AddDataToLogTable(DocType, TaxDate, CardName, "Failed", Msg, "", Mobile, DocNum, DocEntry, Error_Message, "WhatsApp")
                            End If
                            'Else
                            '    objMain.objUtilities.Write_Log("WhatsApp message sending Failed for " + DocType + " Document, DocNumber " + DocNum + " !!! ")
                            '    objApplication.StatusBar.SetText("WhatsApp message sending Failed for " + DocType + " Document, DocNumber " + DocNum + " !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            '    objMain.objUtilities.AddDataToLogTable(DocType, TaxDate, CardName, "Failed", Msg, "", Mobile, DocNum, DocEntry, Result, "WhatsApp")
                        End If

                        Dim AddContactsQry As String = "CALL ""TENX_MESSENGER_MAIN"" ('MainCls','WhatsAppUserEmp','" + DocType + "')"
                        Dim oRsAddContacts As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsAddContacts.DoQuery(AddContactsQry)
                        If oRsAddContacts.RecordCount > 0 Then
                            For i As Integer = 1 To oRsAddContacts.RecordCount
                                Mobile = oRsAddContacts.Fields.Item(0).Value.ToString()
                                Mobile = Mobile.Replace(" ", "")
                                Mobile = Mobile.Replace("-", "")

                                If Mobile <> "" Then
                                    Result = Me.SendWhatsApp(Template, Mobile, DocType, WMSGTXT, Attachment, CardName, DocNum1, TaxDate, ItemCount, TotalBeforDis1, VatSum, DocTotal)
                                    If Result = "Success" Then
                                        Dim oDBDetail1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        Dim TEXT As String = "WhatsApp message sent Successfull to " + Mobile + " for Customer " + CardName + "  !!!  "
                                        Dim DetailsQry As String = "CALL ""TENX_MESSENGER_WHATSAPP_HISTORY"" ('Update','" + DocType + "','" + DocEntry + "','" + DocType + "')"
                                        oDBDetail1.ExecuteQuery(DetailsQry)
                                        objMain.objUtilities.Write_Log("WhatsApp message sent Successfull to " + Mobile + " for Customer " + CardName + "  !!!  ")
                                        objApplication.StatusBar.SetText("WhatsApp message sent Successfull to " + Mobile + " for Customer " + CardName + "  !!!  ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                        objMain.objUtilities.AddDataToLogTable(DocType, TaxDate, CardName, "Success", Msg, "", Mobile, DocNum, DocEntry, Result, "WhatsApp")
                                    Else
                                        objMain.objUtilities.Write_Log("WhatsApp message sending Failed to " + Mobile + " for Customer " + CardName + "  !!!  ")
                                        objApplication.StatusBar.SetText("WhatsApp message sending Failed to " + Mobile + " for Customer " + CardName + "  !!!  ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        objMain.objUtilities.AddDataToLogTable(DocType, TaxDate, CardName, "Failed", Msg, "", Mobile, DocNum, DocEntry, Error_Message, "WhatsApp")
                                    End If
                                    'Else
                                    '    objMain.objUtilities.Write_Log("WhatsApp message sending Failed for " + DocType + " Document, DocNumber " + DocNum + " !!! ")
                                    '    objApplication.StatusBar.SetText("WhatsApp message sending Failed for " + DocType + " Document, DocNumber " + DocNum + " !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    '    objMain.objUtilities.AddDataToLogTable(DocType, TaxDate, CardName, "Failed", Msg, "", Mobile, DocNum, DocEntry, Result, "WhatsApp")
                                End If
                                oRsAddContacts.MoveNext()
                            Next
                        End If
                    Else
                        Dim DraftQuery As String = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingWhatsAppMessage_Draft','" + DocNum + "','" + Series + "')"
                        Dim DraftRecord As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        DraftRecord.DoQuery(DraftQuery)
                        If DraftRecord.RecordCount > 0 Then
                            Dim MobNumber As String = DraftRecord.Fields.Item("Cellular").Value.ToString()
                            SendFormData_Initial(MobNumber).Wait()
                        Else
                            objMain.objUtilities.Write_Log("Data Not Found !!! ")
                            objMain.objApplication.StatusBar.SetText("Data Not Found !!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End If
                    End If
                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText("Error:  " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    objMain.objUtilities.Write_Log(ex.Message)
                End Try
            Else
                objMain.objUtilities.Write_Log("Message Configuration Not Set in Message Template !!! ")
                objMain.objApplication.StatusBar.SetText("Message Configuration Not Set in Message Template", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error:  " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            objMain.objUtilities.Write_Log(ex.Message)
        End Try
    End Sub
    Sub MarketingSMSMessage(ByVal DocNum As String, ByVal Series As String, ByVal DocType As String)
        Dim DetQry As String = ""
        If DocType = "SO" Then
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingSMSMessage','SO','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "PO" Then
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingSMSMessage','PO','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "SQ" Then
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingSMSMessage','SQ','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "DL" Then
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingSMSMessage','DL','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "ARINV" Then
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingSMSMessage','ARINV','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "ARC" Then
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingSMSMessage','ARC','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "ARRET" Then
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingSMSMessage','ARRET','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "PQ" Then
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingSMSMessage','PQ','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "APINV" Then
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingSMSMessage','APINV','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "GRPO" Then
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingSMSMessage','GRPO','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "GR" Then
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingSMSMessage','GR','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "APC" Then
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingSMSMessage','APC','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "IP" Then
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingSMSMessage','IP','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "OP" Then
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingSMSMessage','OP','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "GRP" Then
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingSMSMessage','GRP','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "GI" Then
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingSMSMessage','GI','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "ITR" Then
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingSMSMessage','ITR','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "IT" Then
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingSMSMessage','IT','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "JE" Then
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingSMSMessage','JE','" + DocNum + "','" + Series + "')"
        End If

        Dim oRsDet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRsDet.DoQuery(DetQry)
        Dim Msg As String = objMain.objUtilities.getMessage(DocType, "SMS")
        If Msg <> "" Then
            If oRsDet.RecordCount > 0 Then
                Dim Result As String = ""
                Dim Mobile As String = oRsDet.Fields.Item("Mobile").Value

                Mobile = Mobile.Replace(" ", "")
                Mobile = Mobile.Replace("+", "")

                Dim TotalBeforDis As Double = oRsDet.Fields.Item("DocTotal").Value + oRsDet.Fields.Item("DiscSum").Value - oRsDet.Fields.Item("VatSum").Value

                Msg = Msg.Replace("{CardName}", oRsDet.Fields.Item("CardName").Value.ToString())
                Msg = Msg.Replace("{DocNum}", oRsDet.Fields.Item("DocNum").Value.ToString())
                Msg = Msg.Replace("{TaxDate}", oRsDet.Fields.Item("TaxDate").Value.ToString())
                Msg = Msg.Replace("{ItemCount}", oRsDet.Fields.Item("ItemCount").Value.ToString())
                Msg = Msg.Replace("{Total}", TotalBeforDis.ToString())
                Msg = Msg.Replace("{VatSum}", oRsDet.Fields.Item("VatSum").Value.ToString())
                Msg = Msg.Replace("{DocTotal}", oRsDet.Fields.Item("DocTotal").Value.ToString())
                If Mobile <> "" Then
                    Result = Me.SendSMS(Mobile, Msg)
                    If Result = "Success" Then
                        objMain.objUtilities.Write_Log("SMS sent Successfull for " + DocType + " Document, DocNumber " + DocNum + " !!! ")
                        objApplication.StatusBar.SetText("SMS sent Successfull for " + DocType + " Document, DocNumber " + DocNum + " !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objMain.objUtilities.AddDataToLogTable(DocType, oRsDet.Fields.Item("TaxDate").Value.ToString(), oRsDet.Fields.Item("CardCode").Value.ToString(), "Success", Msg, "", Mobile, DocNum, oRsDet.Fields.Item("DocEntry").Value.ToString(), Result, "SMS")


                    Else
                        objMain.objUtilities.Write_Log("SMS sending Failed for " + DocType + " Document, DocNumber " + DocNum + " !!! ")
                        objApplication.StatusBar.SetText("SMS sending Failed for " + DocType + " Document, DocNumber " + DocNum + " !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        objMain.objUtilities.AddDataToLogTable(DocType, oRsDet.Fields.Item("TaxDate").Value.ToString(), oRsDet.Fields.Item("CardCode").Value.ToString(), "Failed", Msg, "", Mobile, DocNum, oRsDet.Fields.Item("DocEntry").Value.ToString(), Error_Message, "SMS")
                    End If
                    'Else
                    '    objMain.objUtilities.Write_Log("SMS sending Failed for " + DocType + " Document, DocNumber " + DocNum + " !!! ")
                    '    objApplication.StatusBar.SetText("SMS sending Failed for " + DocType + " Document, DocNumber " + DocNum + " !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    objMain.objUtilities.AddDataToLogTable(DocType, oRsDet.Fields.Item("TaxDate").Value.ToString(), oRsDet.Fields.Item("CardCode").Value.ToString(), "Failed", Msg, "", Mobile, DocNum, oRsDet.Fields.Item("DocEntry").Value.ToString(), Result, "SMS")
                End If

                Dim AddContactsQry As String = "CALL ""TENX_MESSENGER_MAIN"" ('SMSUserEmp','" + DocType + "')"
                Dim oRsAddContacts As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsAddContacts.DoQuery(AddContactsQry)
                If oRsAddContacts.RecordCount > 0 Then
                    For i As Integer = 1 To oRsAddContacts.RecordCount
                        Mobile = oRsAddContacts.Fields.Item(0).Value.ToString()
                        Mobile = Mobile.Replace(" ", "")
                        Mobile = Mobile.Replace("+", "")
                        If Mobile <> "" Then
                            Result = Me.SendSMS(Mobile, Msg)
                            If Result = "Success" Then
                                objMain.objUtilities.Write_Log("SMS sent Successfull for " + DocType + " Document, DocNumber " + DocNum + " !!! ")
                                objApplication.StatusBar.SetText("SMS sent Successfull for " + DocType + " Document, DocNumber " + DocNum + " !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                objMain.objUtilities.AddDataToLogTable(DocType, oRsDet.Fields.Item("TaxDate").Value.ToString(), oRsDet.Fields.Item("CardCode").Value.ToString(), "Success", Msg, "", Mobile, DocNum, oRsDet.Fields.Item("DocEntry").Value.ToString(), Result, "SMS")
                            Else
                                objMain.objUtilities.Write_Log("SMS sending Failed for " + DocType + " Document, DocNumber " + DocNum + " !!! ")
                                objApplication.StatusBar.SetText("SMS sending Failed for " + DocType + " Document, DocNumber " + DocNum + " !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objMain.objUtilities.AddDataToLogTable(DocType, oRsDet.Fields.Item("TaxDate").Value.ToString(), oRsDet.Fields.Item("CardCode").Value.ToString(), "Failed", Msg, "", Mobile, DocNum, oRsDet.Fields.Item("DocEntry").Value.ToString(), Error_Message, "SMS")
                            End If
                            'Else
                            '    objMain.objUtilities.Write_Log("SMS sending Failed for " + DocType + " Document, DocNumber " + DocNum + " !!! ")
                            '    objApplication.StatusBar.SetText("SMS sending Failed for " + DocType + " Document, DocNumber " + DocNum + " !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            '    objMain.objUtilities.AddDataToLogTable(DocType, oRsDet.Fields.Item("TaxDate").Value.ToString(), oRsDet.Fields.Item("CardCode").Value.ToString(), "Failed", Msg, "", Mobile, DocNum, oRsDet.Fields.Item("DocEntry").Value.ToString(), Result, "SMS")
                        End If
                        oRsAddContacts.MoveNext()
                    Next
                End If
            Else
                objMain.objUtilities.Write_Log("Data Not Found !!! ")
                objMain.objApplication.StatusBar.SetText("Data Not Found !!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If
        Else
            objMain.objUtilities.Write_Log("Message Configuration Not Set in Message Template !!! ")
            objMain.objApplication.StatusBar.SetText("Message Configuration Not Set in Message Template", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If

    End Sub
    Sub MarketingEmail(ByVal DocNum As String, ByVal Series As String, ByVal DocType As String)
        Dim DetQry As String = ""
        Dim WMSGTXT As String = ""
        Dim MNTXT As String = ""
        If DocType = "SO" Then
            MNTXT = "ORDER CONFIRMATION"
            WMSGTXT = "Thank You... Your Sales order has been placed."
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingEmail','SO','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "PO" Then
            MNTXT = "PO CONFIRMATION"
            WMSGTXT = "Thank You... Your Purchase order has been placed."
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingEmail','PO','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "SQ" Then
            MNTXT = "QUOTATION CONFIRMATION"
            WMSGTXT = "Thank You... Your Sales Quotation has been placed."
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingEmail','SQ','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "DL" Then
            MNTXT = "DELIVERY CONFIRMATION"
            WMSGTXT = "Thank You... Your Delivery Note has been placed."
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingEmail','DL','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "ARINV" Then
            MNTXT = "INVOICE CONFIRMATION"
            WMSGTXT = "Thank You... Your Sales Invoice has been placed."
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingEmail','ARINV','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "ARC" Then
            MNTXT = "CREDIT MEMO CONFIRMATION "
            WMSGTXT = "Thank You... Your A/R Credit Note has been placed."
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingEmail','ARC','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "ARRET" Then
            MNTXT = "RESERVE INVOICE CONFIRMATION"
            WMSGTXT = "Thank You... Your Sales Invoice has been placed."
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingEmail','ARRET','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "PQ" Then
            MNTXT = "PQ CONFIRMATION "
            WMSGTXT = "Thank You... Your Purchase Quotaion has been placed."
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingEmail','PQ','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "APINV" Then
            MNTXT = "A/P INVOICE CONFIRMATION "
            WMSGTXT = "Thank You... Your Purchase Invoice has been placed."
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingEmail','APINV','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "GRPO" Then
            MNTXT = "GRPO CONFIRMATION "
            WMSGTXT = "Thank You... Your Goods Return Against PO has been placed."
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingEmail','GRPO','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "GR" Then
            MNTXT = "GOODS RETURN CONFIRMATION "
            WMSGTXT = "Thank You... Your Goods Return has been placed."
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingEmail','GR','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "APC" Then
            MNTXT = "A/P CREDIT MEMO CONFIRMATION "
            WMSGTXT = "Thank You... Your A/P Credit Note has been placed."
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingEmail','APC','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "IP" Then
            MNTXT = "INCOMING PAYMENT CONFIRMATION "
            WMSGTXT = "Thank You... Your Incomming Payment has been placed."
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingEmail','IP','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "OP" Then
            MNTXT = "OUTGOING PAYMENT CONFIRMATION "
            WMSGTXT = "Thank You... Your Outgoing payment  been placed."
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingEmail','OP','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "GRP" Then
            MNTXT = "GOODS RECEIPT CONFIRMATION "
            WMSGTXT = "Thank You... Your Good Receipt has been placed."
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingEmail','GRP','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "GI" Then
            MNTXT = "GOODS ISSUE CONFIRMATION "
            WMSGTXT = "Thank You... Your Good Issue has been placed."
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingEmail','GI','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "ITR" Then
            MNTXT = "INVENTORY TRANSFER REQUEST CONFIRMATION"
            WMSGTXT = "Inventory Transfer Request is created."
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingEmail','ITR','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "IT" Then
            MNTXT = "INVENTORY TRANSFER CONFIRMATION"
            WMSGTXT = "Thank You... Your Inventory Tranfer has been placed."
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingEmail','IT','" + DocNum + "','" + Series + "')"
        ElseIf DocType = "JE" Then
            MNTXT = "JOURNAL ENTRY CONFIRMATION"
            WMSGTXT = "Thank You... Your JE has been placed."
            DetQry = "CALL ""TENX_MESSENGER_MAIN"" ('MarketingEmail','JE','" + DocNum + "','" + Series + "')"
        End If

        Dim oRsDet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRsDet.DoQuery(DetQry)
        Dim Msg As String = objMain.objUtilities.getMessage(DocType, "Email")
        If Msg <> "" Then
            If oRsDet.RecordCount > 0 Then
                Dim Result As String = ""
                Dim ToEmail As String = oRsDet.Fields.Item("Email").Value

                Dim TotalBeforDis As Double = oRsDet.Fields.Item("DocTotal").Value + oRsDet.Fields.Item("DiscSum").Value - oRsDet.Fields.Item("VatSum").Value
                Dim CardName As String = oRsDet.Fields.Item("CardName").Value.ToString()
                Dim DocNum1 As String = oRsDet.Fields.Item("DocNum").Value.ToString()
                Dim TaxDate As String = oRsDet.Fields.Item("TaxDate").Value.ToString()
                Dim ItemCount As String = oRsDet.Fields.Item("ItemCount").Value.ToString()
                Dim TotalBeforDis1 As String = TotalBeforDis.ToString()
                Dim VatSum As String = oRsDet.Fields.Item("VatSum").Value.ToString()
                Dim DocTotal As String = oRsDet.Fields.Item("DocTotal").Value.ToString()


                Msg = Msg.Replace("{CardName}", oRsDet.Fields.Item("CardName").Value.ToString())
                Msg = Msg.Replace("{DocNum}", oRsDet.Fields.Item("DocNum").Value.ToString())
                Msg = Msg.Replace("{TaxDate}", oRsDet.Fields.Item("TaxDate").Value.ToString())
                Msg = Msg.Replace("{ItemCount}", oRsDet.Fields.Item("ItemCount").Value.ToString())
                Msg = Msg.Replace("{Total}", TotalBeforDis.ToString())
                Msg = Msg.Replace("{VatSum}", oRsDet.Fields.Item("VatSum").Value.ToString())
                Msg = Msg.Replace("{DocTotal}", oRsDet.Fields.Item("DocTotal").Value.ToString())
                If ToEmail <> "" Then
                    Result = Me.SendEmail_TEST(ToEmail, MNTXT, WMSGTXT, oRsDet.Fields.Item("DocEntry").Value.ToString(), DocType, CardName, DocNum1, TaxDate, ItemCount, TotalBeforDis1, VatSum, DocTotal)
                    If Result = "Success" Then
                        objMain.objUtilities.Write_Log("Email sent Successfull for " + DocType + " Document, DocNumber " + DocNum + " !!! ")
                        objApplication.StatusBar.SetText("Email sent Successfull for " + DocType + " Document, DocNumber " + DocNum + " !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objMain.objUtilities.AddDataToLogTable(DocType, oRsDet.Fields.Item("TaxDate").Value.ToString(), oRsDet.Fields.Item("CardCode").Value.ToString(), "Success", Msg, ToEmail, "", DocNum, oRsDet.Fields.Item("DocEntry").Value.ToString(), Result, "Email")
                    Else
                        objMain.objUtilities.Write_Log("Email sending Failed for " + DocType + " Document, DocNumber " + DocNum + " !!! ")
                        objApplication.StatusBar.SetText("Email sending Failed for " + DocType + " Document, DocNumber " + DocNum + " !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        objMain.objUtilities.AddDataToLogTable(DocType, oRsDet.Fields.Item("TaxDate").Value.ToString(), oRsDet.Fields.Item("CardCode").Value.ToString(), "Failed", Msg, ToEmail, "", DocNum, oRsDet.Fields.Item("DocEntry").Value.ToString(), Error_Message, "Email")
                    End If
                    'Else
                    '    objMain.objUtilities.Write_Log("Email sending Failed for " + DocType + " Document, DocNumber " + DocNum + " !!! ")
                    '    objApplication.StatusBar.SetText("Email sending Failed for " + DocType + " Document, DocNumber " + DocNum + " !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '    objMain.objUtilities.AddDataToLogTable(DocType, oRsDet.Fields.Item("TaxDate").Value.ToString(), oRsDet.Fields.Item("CardCode").Value.ToString(), "Failed", Msg, ToEmail, "", DocNum, oRsDet.Fields.Item("DocEntry").Value.ToString(), Result, "Email")
                End If

                Dim AddContactsQry As String = "CALL ""TENX_MESSENGER_MAIN"" ('EmailUserEmp','" + DocType + "')"
                Dim oRsAddContacts As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsAddContacts.DoQuery(AddContactsQry)
                If oRsAddContacts.RecordCount > 0 Then
                    For i As Integer = 1 To oRsAddContacts.RecordCount
                        If oRsAddContacts.Fields.Item(0).Value.ToString() <> "" Then
                            Result = Me.SendEmail_TEST(oRsAddContacts.Fields.Item(0).Value.ToString(), MNTXT, WMSGTXT, oRsDet.Fields.Item("DocEntry").Value.ToString(), DocType, CardName, DocNum1, TaxDate, ItemCount, TotalBeforDis1, VatSum, DocTotal)
                            If Result = "Success" Then
                                objMain.objUtilities.Write_Log("Email sent Successfull for " + DocType + " Document, DocNumber " + DocNum + " !!! ")
                                objApplication.StatusBar.SetText("Email sent Successfull for " + DocType + " Document, DocNumber " + DocNum + " !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                objMain.objUtilities.AddDataToLogTable(DocType, oRsDet.Fields.Item("TaxDate").Value.ToString(), oRsDet.Fields.Item("CardCode").Value.ToString(), "Success", Msg, ToEmail, "", DocNum, oRsDet.Fields.Item("DocEntry").Value.ToString(), Result, "Email")
                            Else
                                objMain.objUtilities.Write_Log("Email sending Failed for " + DocType + " Document, DocNumber " + DocNum + " !!! ")
                                objApplication.StatusBar.SetText("Email sending Failed for " + DocType + " Document, DocNumber " + DocNum + " !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objMain.objUtilities.AddDataToLogTable(DocType, oRsDet.Fields.Item("TaxDate").Value.ToString(), oRsDet.Fields.Item("CardCode").Value.ToString(), "Failed", Msg, ToEmail, "", DocNum, oRsDet.Fields.Item("DocEntry").Value.ToString(), Error_Message, "Email")
                            End If
                            'Else
                            '    objMain.objUtilities.Write_Log("Email sending Failed for " + DocType + " Document, DocNumber " + DocNum + " !!! ")
                            '    objApplication.StatusBar.SetText("Email sending Failed for " + DocType + " Document, DocNumber " + DocNum + " !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            '    objMain.objUtilities.AddDataToLogTable(DocType, oRsDet.Fields.Item("TaxDate").Value.ToString(), oRsDet.Fields.Item("CardCode").Value.ToString(), "Failed", Msg, ToEmail, "", DocNum, oRsDet.Fields.Item("DocEntry").Value.ToString(), Result, "Email")
                        End If
                        oRsAddContacts.MoveNext()
                    Next
                End If
            Else
                objMain.objUtilities.Write_Log("Data Not Found !!! ")
                objMain.objApplication.StatusBar.SetText("Data Not Found !!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If
        Else
            objMain.objUtilities.Write_Log("Message Configuration Not Set in Message Template !!! ")
            objMain.objApplication.StatusBar.SetText("Message Configuration Not Set in Message Template", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If
    End Sub


    Public Function SendWhatsApp(ByVal Template As String, ByVal mobile As String, ByVal mainTxt As String, ByVal WMSGTXT As String, ByVal Attachment As String, ByVal CardName As String, ByVal DocNum1 As String, ByVal TaxDate As String, ItemCount As String, ByVal TotalBeforDis1 As String, ByVal VatSum As String, ByVal DocTotal As String) As String
        Dim Responce As String = ""
        Try

            If Attachment <> "" Then
                SendFormData_WithAttachment(Template, mobile, mainTxt, WMSGTXT, Attachment, CardName, DocNum1, TaxDate, ItemCount, TotalBeforDis1, VatSum, DocTotal).Wait()
                Responce = Responce1
            Else
                SendFormData(Template, mobile, mainTxt, WMSGTXT, CardName, DocNum1, TaxDate, ItemCount, TotalBeforDis1, VatSum, DocTotal).Wait()
                Responce = Responce1
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error:  " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            objMain.objUtilities.Write_Log(ex.Message)
        End Try
        Return Responce
    End Function
    Function SendSMS(ByVal toMobileNo As String, ByVal Message As String) As String
        Dim SMSResult As String
        Dim request As HttpWebRequest
        Dim response As HttpWebResponse = Nothing
        Dim url As String
        Try
            Dim CredQry As String = "Select ""U_SMSUrl"",""U_SMSUId"",""U_SMSPwd"" From ""@SBO_SETCONFIG"" "
            Dim oRsCred As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsCred.DoQuery(CredQry)

            Dim url1 As String = oRsCred.Fields.Item("U_SMSUrl").Value
            Dim UserName As String = oRsCred.Fields.Item("U_SMSUid").Value
            Dim Password As String = oRsCred.Fields.Item("U_SMSPwd").Value

            ''url = "https://api.brio.ae/send?username=kenz&password=kEn@z1Bc&from=KENZ HYPR&to=" + toMobileNo + "&content=" + Message + ""

            url = "" + url1 + "send?username=" + UserName + "&password=" + Password + "&from=KENZ HYPR&to=" + toMobileNo + "&content=" + Message + ""
            ''
            request = DirectCast(WebRequest.Create(url), HttpWebRequest)

            response = DirectCast(request.GetResponse(), HttpWebResponse)

            Dim enc As Encoding = System.Text.Encoding.GetEncoding(1252)

            Dim ResponseStr As String
            Using loResponseStream As StreamReader = New StreamReader(response.GetResponseStream(), enc)
                ResponseStr = loResponseStream.ReadToEnd()
            End Using
            response.Close()
            If ResponseStr.Contains("Success") Then
                objMain.objApplication.StatusBar.SetText("SMS Sent To:" & toMobileNo & "Message:" & Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                objMain.objUtilities.Write_Log("SMS Sent To:" & toMobileNo & "Message:" & Message)
                SMSResult = "Success"
            End If
            Return SMSResult
        Catch ex As Exception
            Error_Message = "Error In SMS Sending:" & ex.Message
            objMain.objUtilities.Write_Log("Error In SMS Sending:" & ex.Message)
            objMain.objApplication.StatusBar.SetText("Error:  " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function
    Function SendEmail_TEST(ByVal ToEmail As String, ByVal mainTxt As String, ByVal WMSGTXT As String, ByVal DocEntry As String, ByVal DocType As String, ByVal CardName As String, ByVal DocNum1 As String, ByVal TaxDate As String, ItemCount As String, ByVal TotalBeforDis1 As String, ByVal VatSum As String, ByVal DocTotal As String) As String
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
                Dim Message As New MailMessage(FromMail, ToEmail)
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
                Dim GetMessageTempDtls As String = "Select ""U_Sub"" From ""@MGRCONFIG"" Where ""U_Doc"" = '" & DocType & "' And IFNULL(""U_Status"",'N') = 'Y' And Ifnull(""U_Sub"",'N') <> 'N' "
                Dim oRsGetMessageTempDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetMessageTempDtls.DoQuery(GetMessageTempDtls)
                If oRsGetMessageTempDtls.RecordCount > 0 Then
                    Message.Subject = oRsGetMessageTempDtls.Fields.Item(0).Value
                Else
                    objMain.objApplication.StatusBar.SetText("Please provide subject in message template ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return "Fail"
                    Exit Function
                End If
                Dim Attachment As String = Nothing
                Dim Base64 As String = Nothing
                Dim AttachmentNew As Attachment
                'Dim Sign As String = Nothing
                'Dim sigReader As StreamReader
                Dim AtttM As String = objUtilities.AttachmnetSelection(DocType, "Email")
                If AtttM = 1 Then
                    Dim AttPath As String = objUtilities.GetAttachmnetPath()
                    Dim CrystalID As String = objUtilities.GetCrystalID(DocType, "Email")
                    Dim Signature As String = objMain.objUtilities.GetSignature(DocType, "Email")
                    ' objUtilities.GetManualAttachment(DocEntry, CrystalID, DocType, "")
                    'Dim Attachment1 As String = AttPath & DocType & "_" & DocEntry & "_" & System.DateTime.Now.ToString("yyyyMMdd") & ".pdf"
                    ' Message.Attachments.Add(New System.Net.Mail.Attachment(Attachment1))
                    ' Attachment = "Please find the attached document. "
                    'Base64 = ConvertFileToBase64(AttPath + Signature)

                    ' AttachmentNew = New Attachment(AttPath + Signature)
                    ' AttachmentNew.ContentId = "NewImage"

                    'Sign = ReadSignature(AttPath)
                    'sigReader = New StreamReader(Sign)
                Else
                    Try
                        If objMain.objUtilities.CheckAttachment_EMAIL(DocType, "Email") = True Then
                            Dim temp As String = objMain.objUtilities.getAttachment(DocType, DocEntry)
                            Message.Attachments.Add(New System.Net.Mail.Attachment(temp))
                            Attachment = "Please find the attached document. "

                        End If
                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText("Error:  " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                        objMain.objUtilities.Write_Log(ex.Message)
                    End Try
                End If
                Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString("<table> <tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'><b>" + mainTxt + "<b> <br></td></tr>" _
                                                                                          + "<tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'><b>" + CardName + "<b><br></td></tr>" _
                                                                                          + "<tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'>Order Summary </td></tr>" _
                                                                                          + "<tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'>Order Number : <b>" + DocNum1 + "<b></td></tr>" _
                                                                                          + "<tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'>Order Date : <b>" + TaxDate + " <b></td></tr>" _
                                                                                          + "<tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'>Product/SKU's Count : <b>" + ItemCount + "<b></td></tr>" _
                                                                                          + "<tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'>Total before Discount : <b>" + TotalBeforDis1 + "<b></td></tr>" _
                                                                                          + "<tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'>Vat Amount : <b>" + VatSum + "<b> </td></tr>" _
                                                                                          + "<tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'>Total Amount : <b>" + DocTotal + "<b></td></tr>" _
                                                                                          + "<tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'>" + WMSGTXT + "<br><br></td></tr>" _
                                                                                          + "<tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'>" + Attachment + "<br><br></td></tr></table>
                                                                                          <div><img src='cid:NewImage' alt='Email Image' /></div>
                                                                                          <table><tr><td style='font-size: 14; font-weight: Regular;font-family:Times New Roman ;'>This is system auto generated mail. Please do not reply. </td></tr></table>", Nothing, "text/html")

                Message.AlternateViews.Add(htmlView)
                ' Message.Attachments.Add(AttachmentNew)
                SMTPServer.Send(Message)

                Message.Dispose()
                objMain.objUtilities.Write_Log("E-Mail Sent successfully to -  " & ToEmail)
                objMain.objApplication.StatusBar.SetText("E-Mail Sent successfully to -  " & ToEmail, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Return "Success"
            End If
        Catch ex As Exception
            Error_Message = "Error In Email Sending:" & ex.Message
            objMain.objUtilities.Write_Log("Error In Email Sending:" & ex.Message)
            objMain.objApplication.StatusBar.SetText("Error In Email Sending:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
    End Function
    Async Function SendFormData_WithAttachment(ByVal Template As String, ByVal mobile As String, ByVal mainTxt As String, ByVal WMSGTXT As String, ByVal Attachment As String, ByVal CardName As String, ByVal DocNum1 As String, ByVal TaxDate As String, ItemCount As String, ByVal TotalBeforDis1 As String, ByVal VatSum As String, ByVal DocTotal As String) As Task
        Try
            Dim CredQry As String = "SELECT ""U_WUId"",""U_Name"" FROM ""@SBO_SETCONFIG"" "
            Dim oRsCred As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsCred.DoQuery(CredQry)

            Dim url1 As String = oRsCred.Fields.Item("U_WUId").Value
            Dim Authorization As String = oRsCred.Fields.Item("U_Name").Value

            Dim url As String = "" + url1 + ""

            Using httpClient1 As New HttpClient()
                httpClient1.DefaultRequestHeaders.Add("Authorization", "" + Authorization + "")
                Using formData1 As New MultipartFormDataContent()
                    Using file1 As New FileStream(Attachment, FileMode.Open)
                        formData1.Add(New StringContent("" + mobile + ""), "to")
                        formData1.Add(New StringContent("template"), "type")
                        formData1.Add(New StringContent("" + Template + ""), "template[name]")
                        formData1.Add(New StringContent("en"), "template[language][code]")
                        formData1.Add(New StringContent("header"), "template[components][0][type]")
                        formData1.Add(New StringContent("DOCUMENT"), "template[components][0][parameters][0][type]")
                        formData1.Add(New StringContent("" + mainTxt + "_" + DocNum1 + "_Attachment"), "template[components][0][parameters][0][document][filename]")
                        formData1.Add(New StringContent("{{file}}"), "template[components][0][parameters][0][document][link]")
                        formData1.Add(New StringContent("body"), "template[components][1][type]")
                        formData1.Add(New StringContent("text"), "template[components][1][parameters][0][type]")
                        formData1.Add(New StringContent("" + CardName + ""), "template[components][1][parameters][0][text]")
                        formData1.Add(New StringContent("text"), "template[components][1][parameters][1][type]")
                        formData1.Add(New StringContent("" + DocNum1 + ""), "template[components][1][parameters][1][text]")
                        formData1.Add(New StringContent("text"), "template[components][1][parameters][2][type]")
                        formData1.Add(New StringContent("" + TaxDate + ""), "template[components][1][parameters][2][text]")
                        formData1.Add(New StringContent("text"), "template[components][1][parameters][3][type]")
                        formData1.Add(New StringContent("" + ItemCount + ""), "template[components][1][parameters][3][text]")
                        formData1.Add(New StringContent("text"), "template[components][1][parameters][4][type]")
                        formData1.Add(New StringContent("" + TotalBeforDis1 + ""), "template[components][1][parameters][4][text]")
                        formData1.Add(New StringContent("text"), "template[components][1][parameters][5][type]")
                        formData1.Add(New StringContent("" + VatSum + ""), "template[components][1][parameters][5][text]")
                        formData1.Add(New StringContent("text"), "template[components][1][parameters][6][type]")
                        formData1.Add(New StringContent("" + DocTotal + ""), "template[components][1][parameters][6][text]")
                        formData1.Add(New StreamContent(file1), "template[file]", Path.GetFileName(Attachment))
                        objMain.objUtilities.ShowWarningMessage("WhatsApp Attachment generating in progress for Customer '" + CardName + "' , Please wait...")


                        Try
                            mainTxt = mainTxt

                            Dim response As HttpResponseMessage = httpClient1.PostAsync(url, formData1).Result
                            If response.IsSuccessStatusCode Then
                                Dim responseContent As String = Await response.Content.ReadAsStringAsync()
                                objMain.objUtilities.Write_Log("WhatsApp Message Sent To:" & mobile & "Message:" & responseContent)
                                objMain.objApplication.StatusBar.SetText("WhatsApp Message Sent To:" & mobile & "Message:" & responseContent, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                                Responce1 = "Success"

                            Else
                                objMain.objUtilities.Write_Log("WhatsApp Message Sent To:" & mobile & "Message:" & response.StatusCode)
                                objMain.objApplication.StatusBar.SetText("WhatsApp Message Sent To:" & mobile & "Message:" & response.StatusCode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Responce1 = "Fail"
                                Error_Message = response.StatusCode
                            End If
                        Catch ex As Exception
                            objMain.objUtilities.Write_Log("An error occurred: " & ex.Message)
                            objMain.objApplication.StatusBar.SetText("n error occurred: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End Using
                End Using
            End Using
        Catch ex As Exception
            objMain.objUtilities.Write_Log("An error occurred: " + ex.Message)
            Error_Message = ex.Message
            Responce1 = "Fail"
            objMain.objApplication.StatusBar.SetText("n error occurred: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)


        End Try

    End Function
    Async Function SendFormData(ByVal Template As String, ByVal mobile As String, ByVal mainTxt As String, ByVal WMSGTXT As String, ByVal CardName As String, ByVal DocNum1 As String, ByVal TaxDate As String, ItemCount As String, ByVal TotalBeforDis1 As String, ByVal VatSum As String, ByVal DocTotal As String) As Task
        Try

            Dim CredQry As String = "SELECT ""U_WUId"",""U_Name"" FROM ""@SBO_SETCONFIG"" "
            Dim oRsCred As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsCred.DoQuery(CredQry)

            Dim url1 As String = oRsCred.Fields.Item("U_WUId").Value
            Dim Authorization As String = oRsCred.Fields.Item("U_Name").Value

            Dim url As String = "" + url1 + ""
            Dim httpClient As New HttpClient()
            httpClient.DefaultRequestHeaders.Add("Authorization", "" + Authorization + "")

            Dim formData As New List(Of KeyValuePair(Of String, String))()
            formData.Add(New KeyValuePair(Of String, String)("to", "" + mobile + ""))
            formData.Add(New KeyValuePair(Of String, String)("type", "template"))
            formData.Add(New KeyValuePair(Of String, String)("template[name]", "order_confirmation_2"))
            formData.Add(New KeyValuePair(Of String, String)("template[language][code]", "en"))
            formData.Add(New KeyValuePair(Of String, String)("template[components][1][type]", "body"))
            formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][0][type]", "text"))
            formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][0][text]", "" + CardName + ""))
            formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][1][type]", "text"))
            formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][1][text]", "" + DocNum1 + ""))
            formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][2][type]", "text"))
            formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][2][text]", "" + TaxDate + ""))
            formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][3][type]", "text"))
            formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][3][text]", "" + ItemCount + ""))
            formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][4][type]", "text"))
            formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][4][text]", "" + TotalBeforDis1 + ""))
            formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][5][type]", "text"))
            formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][5][text]", "" + VatSum + ""))
            formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][6][type]", "text"))
            formData.Add(New KeyValuePair(Of String, String)("template[components][1][parameters][6][text]", "" + DocTotal + ""))

            Dim content As New FormUrlEncodedContent(formData)

            Try
                Dim response As HttpResponseMessage = Await httpClient.PostAsync(url, content)
                If response.IsSuccessStatusCode Then
                    Dim responseContent As String = Await response.Content.ReadAsStringAsync()
                    objMain.objUtilities.Write_Log("WhatsApp Message Sent To:" & mobile & "Message:" & responseContent)
                    objMain.objApplication.StatusBar.SetText("WhatsApp Message Sent To:" & mobile & "Message:" & responseContent, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Responce1 = "Success"
                Else
                    objMain.objUtilities.Write_Log("WhatsApp Message Sent To:" & mobile & "Message:" & response.StatusCode)
                    objMain.objApplication.StatusBar.SetText("WhatsApp Message Sent To:" & mobile & "Message:" & response.StatusCode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Responce1 = "Fail"
                    Error_Message = response.StatusCode
                End If
            Catch ex As Exception
                objMain.objUtilities.Write_Log("An error occurred: " & ex.Message)
                objMain.objApplication.StatusBar.SetText("n error occurred: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error:  " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            objMain.objUtilities.Write_Log(ex.Message)
        End Try
    End Function
    Async Function SendFormData_Initial(ByVal mobile As String) As Task
        Try
            Dim CredQry As String = "SELECT ""U_WUId"",""U_Name"" FROM ""@SBO_SETCONFIG"" "
            Dim oRsCred As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsCred.DoQuery(CredQry)

            Dim url1 As String = oRsCred.Fields.Item("U_WUId").Value
            Dim Authorization As String = oRsCred.Fields.Item("U_Name").Value

            Dim url As String = "" + url1 + ""
            Dim httpClient As New HttpClient()
            httpClient.DefaultRequestHeaders.Add("Authorization", "" + Authorization + "")

            Dim formData As New List(Of KeyValuePair(Of String, String))()
            formData.Add(New KeyValuePair(Of String, String)("to", "" + mobile + ""))
            formData.Add(New KeyValuePair(Of String, String)("type", "template"))
            formData.Add(New KeyValuePair(Of String, String)("template[name]", "review_doc"))
            formData.Add(New KeyValuePair(Of String, String)("template[language][code]", "en"))
            formData.Add(New KeyValuePair(Of String, String)("template[components][1][type]", "body"))

            Dim content As New FormUrlEncodedContent(formData)

            Try
                Dim response As HttpResponseMessage = Await httpClient.PostAsync(url, content)
                If response.IsSuccessStatusCode Then
                    Dim responseContent As String = Await response.Content.ReadAsStringAsync()
                    objMain.objUtilities.Write_Log("WhatsApp Message Sent To:" & mobile & "Message:" & responseContent)
                Else
                    objMain.objUtilities.Write_Log("WhatsApp Message Sent To:" & mobile & "Message:" & response.StatusCode)
                End If
            Catch ex As Exception
                objMain.objUtilities.Write_Log("An error occurred: " & ex.Message)
                objMain.objApplication.StatusBar.SetText("n error occurred: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error:  " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            objMain.objUtilities.Write_Log(ex.Message)
        End Try
    End Function

    Private Function ReadSignature(ByVal SignaturePath As String) As String
        Dim signature = String.Empty
        Try
            ''Dim appDataDir = Environment.GetFolderPath(SignaturePath)
            Dim diInfo As DirectoryInfo = New DirectoryInfo(SignaturePath)

            If diInfo.Exists Then
                Dim fiSignature = diInfo.GetFiles("*.htm")

                If fiSignature.Length > 0 Then
                    Dim sr As StreamReader = New StreamReader(fiSignature(0).FullName, Encoding.Default)
                    signature = sr.ReadToEnd()
                    If Not String.IsNullOrEmpty(signature) Then
                        Dim fileName = fiSignature(0).Name.Replace(fiSignature(0).Extension, String.Empty)
                        signature = signature.Replace(fileName & "_files/", SignaturePath & "/" & fileName & "_files/")
                    End If
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error:  " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
        Return signature
    End Function

End Class