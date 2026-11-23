'new
Imports System.Configuration
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Public Class clsPromoCostReduction
#Region "       Declaration             "
    Public objForm, oForm, objSTRForm As SAPbouiCOM.Form
    Dim objMatrix, oMatrix, objMatrix1, objMatrix2 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Public oDS, oDS1, oDS2, oDS3 As SAPbouiCOM.DBDataSource
    Public PrgBar As SAPbouiCOM.ProgressBar
    Dim LostFocusFlag As Boolean
    Dim MATRIXS As String
#End Region
    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("Reduction3.xml", "REDCTN", ResourceType.Embeded)
            Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.ActiveForm
            objForm = objMain.objApplication.Forms.GetForm("REDCTN", objMain.objApplication.Forms.ActiveForm.TypeCount)
            ' objForm.Freeze(True)
            objForm.Freeze(False)
            Me.EnableCoulmns(objForm.UniqueID)
            objForm.EnableMenu("1292", True)
            ' objForm.EnableMenu("1293", True)
            'objForm.EnableMenu("1282", True)
            objForm.EnableMenu("774", True)
            Me.SetDefault(objForm.UniqueID)
            objComboBox = objForm.Items.Item("BRANCH").Specific
            Dim oRsUpdate2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim BPLId As String = " Select T0.""BPLId"",T1.""BPLName"",T1.""DflWhs"" From USR6 T0 Inner Join " &
           " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "'Order By T0.""BPLId""  "
            oRsUpdate2.DoQuery(BPLId)
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, BPLId)
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("BRANCH").DisplayDesc = True
            'objForm.Items.Item("ITMCODE").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            objForm.Freeze(False)
            Me.DisableCoulmns(objForm.UniqueID)
            Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Me.objForm.Items.Item("DOCDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Me.objForm.Items.Item("DOCDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub EnableCoulmns(ByVal FormUID As String)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("RDMtx").Specific
            Try
                objMatrix.Columns.Item("NEWCST").Visible = True
            Catch ex As Exception
            End Try
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub



    'Sub SetDefault1(ByVal FormUID As String)
    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)
    '        oDS = objForm.DataSources.DBDataSources.Item("@TNX_RDCTN")
    '        oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_RDCTN_C0")
    '        objForm.Freeze(True)
    '        oDS.SetValue("DocNum", oDS.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TNX_RDCTNUDO", "Primary"))
    '        objForm.Freeze(False)
    '    Catch ex As Exception
    '        objForm.Freeze(False 
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub

    Private Sub DisableCoulmns(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("RDMtx").Specific
            Try
                objMatrix.Columns.Item("BRANCH").Visible = False

            Catch ex As Exception
            End Try
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ME_REDCTN" And pVal.BeforeAction = False Then
                Me.CreateForm()
                'Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
                ' objMatrix = objForm.Items.Item("RDMtx").Specific

            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                Me.SetNewLine(objForm.UniqueID, "RDMtx")
                'Me.SetDefault(objForm.UniqueID)
                objMatrix = objForm.Items.Item("RDMtx").Specific
                'ElseIf pVal.MenuUID = "1293" And pVal.BeforeAction = False Then
                '    objForm = objMain.objApplication.Forms.GetForm("REDCTN", objMain.objApplication.Forms.ActiveForm.TypeCount)
                '    Me.SetDefault(objForm.UniqueID)
                '    objMatrix = objForm.Items.Item("RDMtx").Specific
                '    For i As Integer = 1 To objMatrix.VisualRowCount - 1
                '        If objMatrix.IsRowSelected(i) = True Then
                '            objMatrix.DeleteRow(i)
                '        End If
                '    Next
            ElseIf pVal.MenuUID = "774" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("REDCTN", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix = objForm.Items.Item("RDMtx").Specific

                Dim row As Integer = objMatrix.VisualRowCount
                If MATRIXS.Equals("RDMtx") = True Then
                    If objMatrix.IsRowSelected(1) <> True And objMatrix.VisualRowCount < 1 Then
                        objMatrix.AddRow()
                        oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
                        objMatrix.SetLineData(objMatrix.VisualRowCount)
                    End If
                    If objMatrix.IsRowSelected(row) = True Then
                        objMatrix.DeleteRow(row)
                    Else
                        For i As Integer = 1 To objMatrix.VisualRowCount - 1
                            If objMatrix.IsRowSelected(i) = True Then
                                If i <> 1 Then
                                    objMatrix.DeleteRow(i)
                                End If
                            End If
                        Next
                    End If
                    'ElseIf pVal.MenuUID = "Delete Row" And pVal.BeforeAction = False Then
                    '    objForm = objMain.objApplication.Forms.GetForm("REDCTN", objMain.objApplication.Forms.ActiveForm.TypeCount)
                    '    objMatrix = objForm.Items.Item("RDMtx").Specific
                    '    For i As Integer = 1 To objMatrix.VisualRowCount - 1
                    '        If objMatrix.IsRowSelected(i) = True Then
                    '            objMatrix.DeleteRow(i)
                    '        End If
                    '    Next
                    '    For i As Integer = 1 To objMatrix.VisualRowCount
                    '        objMatrix.Columns.Item("LineId").Cells.Item(i).Specific.string = i
                    '    Next
                    '    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    '        objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    '       ' objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    '    End If
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If (pVal.ItemUID = "RDMtx") And pVal.BeforeAction = False Then
                        MATRIXS = pVal.ItemUID
                    End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "GLINK" And pVal.BeforeAction = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or
                            pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                        Dim AppTemForm_EMD As SAPbouiCOM.Form
                        objMain.objApplication.ActivateMenuItem("2561")
                        AppTemForm_EMD = objMain.objApplication.Forms.GetFormByTypeAndCount(134, 1)
                        AppTemForm_EMD = objMain.objApplication.Forms.ActiveForm
                        AppTemForm_EMD.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        Dim sercall As String = objForm.Items.Item("GLACNT").Specific.value
                        AppTemForm_EMD.Items.Item("5").Specific.Value = sercall
                        ' AppTemForm_EMD.Items.Item("SERDC").Specific.Value = sercall
                        AppTemForm_EMD.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        AppTemForm_EMD = objMain.objApplication.Forms.ActiveForm
                    End If

                    If pVal.ItemUID = "SLINK" And pVal.BeforeAction = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or
                            pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                        Dim AppTemForm_EMD As SAPbouiCOM.Form
                        objMain.objApplication.ActivateMenuItem("2561")
                        AppTemForm_EMD = objMain.objApplication.Forms.GetFormByTypeAndCount(134, 1)
                        AppTemForm_EMD = objMain.objApplication.Forms.ActiveForm

                        AppTemForm_EMD.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        Dim sercall As String = objForm.Items.Item("SPCODE").Specific.value
                        AppTemForm_EMD.Items.Item("5").Specific.Value = sercall
                        ' AppTemForm_EMD.Items.Item("SERDC").Specific.Value = sercall
                        AppTemForm_EMD.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        AppTemForm_EMD = objMain.objApplication.Forms.ActiveForm
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.ValidateBuuyerName(objForm.UniqueID) = False Then
                            BubbleEvent = False
                        End If
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.ValidateSupplierName(objForm.UniqueID) = False Then
                            BubbleEvent = False
                        End If
                    End If


                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.ValidateDates(objForm.UniqueID) = False Then
                            BubbleEvent = False
                        End If
                        ' Me.GoodsReceipt(objForm.UniqueID)
                    End If






                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                    '    If Me.ValidateDates(objForm.UniqueID) = False Then
                    '        BubbleEvent = False
                    '    End If
                    'End If
                    'Akhila
                    If pVal.ItemUID = "INVTR" And pVal.BeforeAction = False Then
                        Me.PostInventoryRevaluation(objForm.UniqueID)
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And
(pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        Me.SetDefault(objForm.UniqueID)
                    End If

                    'Akhila
                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                    '    If Me.ValidateRevaluationDocumentReference(objForm.UniqueID, "RDMtx") = False Then
                    '        BubbleEvent = False
                    '    End If
                    'End If
            End Select





            'new

            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    oDS = objForm.DataSources.DBDataSources.Item("@TNX_RDCTN")
                    oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_RDCTN_C0")
                    objMatrix = objForm.Items.Item("RDMtx").Specific
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                    'If oCFL.UniqueID = "CFL_SCODE" And pVal.BeforeAction = True Then 'CHOOSE FORM TABLE UNIQUNE ID
                    'Dim Suppliercode11 As String = oDT.GetValue("CardCode", 0)

                    If oCFL.UniqueID = "CFL_SCODE" And pVal.BeforeAction = True Then 'CHOOSE FORM TABLE UNIQUNE ID  And pVal.BeforeAction = False
                        Me.CFLCardCodeFilter(objForm.UniqueID, "CFL_SCODE")

                    End If

                    If oCFL.UniqueID = "CFL_ICODE" And pVal.BeforeAction = True Then
                        Me.CFLItem(objForm.UniqueID, "CFL_ICODE", objForm.Items.Item("SPCODE").Specific.Value)
                    End If

                    'If oCFL.UniqueID = "CFL_WHOUSE" And pVal.BeforeAction = True Then
                    '    Me.CFLFilter(FormUID, "CFL_WHOUSE")
                    'End If
                    'If oCFL.UniqueID = "CFL_BRANCH" And pVal.BeforeAction = True Then
                    '    Me.CFLFilter(FormUID, "CFL_WHOUSE")
                    'End If


                    'If oCFL.UniqueID = "CFL_WHSCD" And pVal.BeforeAction = True Then
                    '    Me.CFLFilter(FormUID, objForm.Items.Item("BRANCH").Specific.Value)
                    'End If
                    'If oCFL.UniqueID = "CFL_BAR" And pVal.BeforeAction = True Then

                    '    Me.CFLFilterBarcode(objForm.UniqueID, "CFL_BAR", objForm.Items.Item("SPCODE").Specific.Value)

                    'End If
                    'If pVal.ItemUID = "RDMtx" And pVal.ColUID = "WRHSE" And pVal.BeforeAction = True Then

                    '    ' If pVal.ItemUID = "" And pVal.BeforeAction = False Then
                    '    Dim branch As String = objForm.Items.Item("BRANCH").Specific.Value
                    '    Me.CFLFilter(FormUID, "CFL_WRHSE")
                    'End If



                    If pVal.ItemUID = "RDMtx" AndAlso pVal.ColUID = "WRHSE" AndAlso pVal.BeforeAction = True Then
                        Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(pVal.FormUID)
                        Dim branch As String = objForm.Items.Item("BRANCH").Specific.Value
                        Me.CFLFilter(FormUID, "CFL_WHOUSE", branch)
                    End If






                    If oCFL.UniqueID = "CFL_BAR" And pVal.BeforeAction = True Then
                        ''Me.CFLFilter("CFL_OCRD")
                        Dim StrCFLQry As String = ""
                        If objForm.Items.Item("SPCODE").Specific.Value <> "" Then
                            StrCFLQry = "SELECT B.""BcdCode"",B.""ItemCode"",B.""BcdName"" " &
                                      "FROM ""OBCD"" B " &
                                                 "INNER JOIN ""OITM"" I ON B.""ItemCode"" = I.""ItemCode"" " &
                                                  "WHERE I.""CardCode"" = '" & objForm.Items.Item("SPCODE").Specific.Value & "'"
                        End If
                        FILTER_CFL(pVal, objMain.objApplication, objMain.objCompany, StrCFLQry, "BcdCode", CFL_Id)
                    End If
                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)


                        If oCFL.UniqueID = "CFL_BAR" And pVal.BeforeAction = False Then
                            oDS1.SetValue("LineId", oDS1.Offset, pVal.Row)
                            oDS1.SetValue("U_BARCD", oDS1.Offset, oDT.GetValue("BcdCode", 0))
                            oDS1.SetValue("U_ITMCODE", oDS1.Offset, oDT.GetValue("ItemCode", 0))
                            oDS1.SetValue("U_ITMDESC", oDS1.Offset, oDT.GetValue("BcdName", 0))

                            objMatrix.SetLineData(pVal.Row)
                            If objMatrix.VisualRowCount = pVal.Row Then
                                Me.SetNewLine(objForm.UniqueID, "RDMtx")
                            End If
                            ''Me.SetNewLine(objForm.UniqueID, "RDMtx")

                        End If

                    End If











                    'If oCFL.UniqueID = "CFL_BAR" And pVal.BeforeAction = True Then

                    '    Dim Str As String = ""
                    '    Dim StrCFLQry1 As SAPbobsCOM.Recordset
                    '    StrCFLQry1 = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '    If objForm.Items.Item("SPCODE").Specific.Value <> "" Then
                    '        Str = "SELECT B.""BcdCode"",B.""ItemCode"",B.""BcdName"" " &
                    '                   "FROM ""OBCD"" B " &
                    '                               "INNER JOIN ""OITM"" I ON B.""ItemCode"" = I.""ItemCode"" " &
                    '                               "WHERE I.""CardCode"" = '" & objForm.Items.Item("SPCODE").Specific.Value & "'"
                    '        StrCFLQry1.DoQuery(Str)
                    '    End If
                    '    FILTER_CFL(pVal, objMain.objApplication, objMain.objCompany, Str, "BcdCode", CFL_Id)
                    '    oDS1.SetValue("LineId", oDS1.Offset, pVal.Row)
                    '    oDS1.SetValue("U_BARCD", oDS1.Offset, StrCFLQry1.Fields.Item("BcdCode").Value)
                    '    oDS1.SetValue("U_ITMCODE", oDS1.Offset, StrCFLQry1.Fields.Item("ItemCode").Value)
                    '    oDS1.SetValue("U_ITMDESC", oDS1.Offset, StrCFLQry1.Fields.Item("BcdName").Value)

                    '    'oDS1.SetValue("U_BARCD", oDS1.Offset, oDT.GetValue("BcdCode", 0))
                    '    'oDS1.SetValue("U_ITMCODE", oDS1.Offset, oDT.GetValue("ItemCode", 0))
                    '    'oDS1.SetValue("U_ITMDESC", oDS1.Offset, oDT.GetValue("BcdName", 0))
                    '    'objMatrix.SetLineData(pVal.Row)
                    'End If



                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        Try










                            If oCFL.UniqueID = "CFL_BYNAME" Then
                                oDS.SetValue("U_BYNAME", oDS.Offset, oDT.GetValue("SlpName", 0))
                                Dim SalesEmpCode As String = oDT.GetValue("SlpCode", 0)
                                Dim rs As SAPbobsCOM.Recordset
                                rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim query As String = "SELECT ""U_GLAccount"" FROM ""OSLP"" WHERE ""SlpCode"" = '" & SalesEmpCode & "'"
                                rs.DoQuery(query)

                                If rs.RecordCount > 0 Then
                                    Dim glAccount As String = rs.Fields.Item("U_GLAccount").Value
                                    oDS.SetValue("U_GLACNT", oDS.Offset, glAccount)

                                Else
                                    'Me.SetNewLine(objForm.UniqueID, "RDMtx")
                                End If

                                'objMatrix.AutoResizeColumns()
                                'objMatrix.FlushToDataSource()
                                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                End If
                            End If





                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message & "Page Load  Events", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try

                        If oCFL.UniqueID = "CFL_SCODE" Then 'CHOOSE FORM TABLE UNIQUNE ID
                            oDS.SetValue("U_SPCODE", oDS.Offset, oDT.GetValue("CardCode", 0))
                            oDS.SetValue("U_SPNAME", oDS.Offset, oDT.GetValue("CardName", 0))
                        End If

                        If oCFL.UniqueID = "CFL_ICODE" Then
                            'Commented
                            oDS1.SetValue("LineId", oDS1.Offset, pVal.Row)
                            oDS1.SetValue("U_ITMCODE", oDS1.Offset, oDT.GetValue("ItemCode", 0))
                            oDS1.SetValue("U_ITMDESC", oDS1.Offset, oDT.GetValue("ItemName", 0))

                            Dim SalesEmpCode As String = objForm.Items.Item("BYNAME").Specific.Value
                            Dim rs As SAPbobsCOM.Recordset
                            rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim query As String = "SELECT ""U_GLAccount"" FROM ""OSLP"" WHERE ""SlpName"" = '" & SalesEmpCode & "'"
                            rs.DoQuery(query)
                            'Commented
                            oDS1.SetValue("U_BARCD", oDS1.Offset, objMatrix.Columns.Item("BARCD").Cells.Item(pVal.Row).Specific.Value)

                            oDS1.SetValue("U_WRHSE", oDS1.Offset, objMatrix.Columns.Item("WRHSE").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_ISTQTY", oDS1.Offset, objMatrix.Columns.Item("ISTQTY").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_OLDCST", oDS1.Offset, objMatrix.Columns.Item("OLDCST").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_NEWCST", oDS1.Offset, objMatrix.Columns.Item("NEWCST").Cells.Item(pVal.Row).Specific.Value)
                            'oDS1.SetValue("U_GLACNT", oDS1.Offset, rs.Fields.Item(0).Value)
                            oDS1.SetValue("U_JRMRK", oDS1.Offset, objMatrix.Columns.Item("JRMRK").Cells.Item(pVal.Row).Specific.Value)
                            'oDS1.SetValue("U_BRANCH", oDS1.Offset, objMatrix.Columns.Item("BRANCH").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_REF1", oDS1.Offset, objMatrix.Columns.Item("REF1").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_REF2", oDS1.Offset, objMatrix.Columns.Item("REF2").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_IREVNO", oDS1.Offset, objMatrix.Columns.Item("IREVNO").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_PSTS", oDS1.Offset, objMatrix.Columns.Item("PSTS").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_REVDRF", oDS1.Offset, objMatrix.Columns.Item("REVDRF").Cells.Item(pVal.Row).Specific.Value)
                            objMatrix.SetLineData(pVal.Row)
                            ' Me.SetNewLine(objForm.UniqueID, "RDMtx")

                            If objMatrix.VisualRowCount = pVal.Row Then
                                Me.SetNewLine(objForm.UniqueID, "RDMtx")
                            End If
                        End If

                        If oCFL.UniqueID = "CFL_WHOUSE" Then
                            oDS1.SetValue("LineId", oDS1.Offset, pVal.Row)
                            oDS1.SetValue("U_BARCD", oDS1.Offset, objMatrix.Columns.Item("BARCD").Cells.Item(pVal.Row).Specific.Value)

                            oDS1.SetValue("U_ITMCODE", oDS1.Offset, objMatrix.Columns.Item("ITMCODE").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_ITMDESC", oDS1.Offset, objMatrix.Columns.Item("ITMDESC").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_WRHSE", oDS1.Offset, oDT.GetValue("WhsCode", 0))
                            oDS1.SetValue("U_ISTQTY", oDS1.Offset, objMatrix.Columns.Item("ISTQTY").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_OLDCST", oDS1.Offset, objMatrix.Columns.Item("OLDCST").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_NEWCST", oDS1.Offset, objMatrix.Columns.Item("NEWCST").Cells.Item(pVal.Row).Specific.Value)
                            'oDS1.SetValue("U_GLACNT", oDS1.Offset, objMatrix.Columns.Item("GLACNT").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_JRMRK", oDS1.Offset, objMatrix.Columns.Item("JRMRK").Cells.Item(pVal.Row).Specific.Value)
                            'oDS1.SetValue("U_BRANCH", oDS1.Offset, objMatrix.Columns.Item("BRANCH").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_REF1", oDS1.Offset, objMatrix.Columns.Item("REF1").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_REF2", oDS1.Offset, objMatrix.Columns.Item("REF2").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_IREVNO", oDS1.Offset, objMatrix.Columns.Item("IREVNO").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_PSTS", oDS1.Offset, objMatrix.Columns.Item("PSTS").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_REVDRF", oDS1.Offset, objMatrix.Columns.Item("REVDRF").Cells.Item(pVal.Row).Specific.Value)
                            objMatrix.SetLineData(pVal.Row)
                            If objMatrix.VisualRowCount = pVal.Row Then
                                Me.SetNewLine(objForm.UniqueID, "RDMtx")
                            End If
                            'Me.SetNewL
                        End If

                        'If oCFL.UniqueID = "CFL_BRANCH" Then
                        '    oDS.SetValue("U_BRANCH", oDS.Offset, oDT.GetValue("BPLId", 0))
                        '    'Commented

                        'End If

                        'If oCFL.UniqueID = "CFL_BRANCH" Then
                        '    oDS1.SetValue("LineId", oDS1.Offset, pVal.Row)
                        '    oDS1.SetValue("U_ITMCODE", oDS1.Offset, objMatrix.Columns.Item("ITMCODE").Cells.Item(pVal.Row).Specific.Value)
                        '    oDS1.SetValue("U_ITMDESC", oDS1.Offset, objMatrix.Columns.Item("ITMDESC").Cells.Item(pVal.Row).Specific.Value)
                        '    'Commented
                        '    oDS1.SetValue("U_WRHSE", oDS1.Offset, objMatrix.Columns.Item("WRHSE").Cells.Item(pVal.Row).Specific.Value)
                        '    oDS1.SetValue("U_ISTQTY", oDS1.Offset, objMatrix.Columns.Item("ISTQTY").Cells.Item(pVal.Row).Specific.Value)
                        '    oDS1.SetValue("U_OLDCST", oDS1.Offset, objMatrix.Columns.Item("OLDCST").Cells.Item(pVal.Row).Specific.Value)
                        '    oDS1.SetValue("U_NEWCST", oDS1.Offset, objMatrix.Columns.Item("NEWCST").Cells.Item(pVal.Row).Specific.Value)
                        '    'oDS1.SetValue("U_GLACNT", oDS1.Offset, objMatrix.Columns.Item("GLACNT").Cells.Item(pVal.Row).Specific.Value)
                        '    oDS1.SetValue("U_JRMRK", oDS1.Offset, objMatrix.Columns.Item("JRMRK").Cells.Item(pVal.Row).Specific.Value)
                        '    oDS1.SetValue("U_BRANCH", oDS1.Offset, oDT.GetValue("BPLName", 0))
                        '    oDS1.SetValue("U_REF1", oDS1.Offset, objMatrix.Columns.Item("REF1").Cells.Item(pVal.Row).Specific.Value)
                        '    oDS1.SetValue("U_REF2", oDS1.Offset, objMatrix.Columns.Item("REF2").Cells.Item(pVal.Row).Specific.Value)
                        '    oDS1.SetValue("U_IREVNO", oDS1.Offset, objMatrix.Columns.Item("IREVNO").Cells.Item(pVal.Row).Specific.Value)
                        '    oDS1.SetValue("U_PSTS", oDS1.Offset, objMatrix.Columns.Item("PSTS").Cells.Item(pVal.Row).Specific.Value)
                        '    oDS1.SetValue("U_REVDRF", oDS1.Offset, objMatrix.Columns.Item("REVDRF").Cells.Item(pVal.Row).Specific.Value)
                        '    objMatrix.SetLineData(pVal.Row)
                        'End If

                        If oCFL.UniqueID = "CFL_IREVNO" Then
                            oDS1.SetValue("LineId", oDS1.Offset, pVal.Row)
                            oDS1.SetValue("U_ITMCODE", oDS1.Offset, objMatrix.Columns.Item("ITMCODE").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_ITMDESC", oDS1.Offset, objMatrix.Columns.Item("ITMDESC").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_WRHSE", oDS1.Offset, objMatrix.Columns.Item("WRHSE").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_ISTQTY", oDS1.Offset, objMatrix.Columns.Item("ISTQTY").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_OLDCST", oDS1.Offset, objMatrix.Columns.Item("OLDCST").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_NEWCST", oDS1.Offset, objMatrix.Columns.Item("NEWCST").Cells.Item(pVal.Row).Specific.Value)
                            'oDS1.SetValue("U_GLACNT", oDS1.Offset, objMatrix.Columns.Item("GLACNT").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_JRMRK", oDS1.Offset, objMatrix.Columns.Item("JRMRK").Cells.Item(pVal.Row).Specific.Value)
                            ' oDS1.SetValue("U_BRANCH", oDS1.Offset, objMatrix.Columns.Item("BRANCH").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_REF1", oDS1.Offset, objMatrix.Columns.Item("REF1").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_REF2", oDS1.Offset, objMatrix.Columns.Item("REF2").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_IREVNO", oDS1.Offset, oDT.GetValue("DocEntry", 0))
                            oDS1.SetValue("U_PSTS", oDS1.Offset, objMatrix.Columns.Item("PSTS").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_REVDRF", oDS1.Offset, objMatrix.Columns.Item("REVDRF").Cells.Item(pVal.Row).Specific.Value)
                            objMatrix.SetLineData(pVal.Row)
                        End If


                        'Commented
                        If oCFL.UniqueID = "CFL_WHOUSE" Then
                            ' Get the selected values from the Matrix
                            Dim itemCode As String = objMatrix.Columns.Item("ITMCODE").Cells.Item(pVal.Row).Specific.Value
                            Dim warehouseCode As String = objMatrix.Columns.Item("WRHSE").Cells.Item(pVal.Row).Specific.Value
                            If Not String.IsNullOrEmpty(itemCode) And Not String.IsNullOrEmpty(warehouseCode) Then

                                ' Initialize Recordset
                                Dim oRecordSet As SAPbobsCOM.Recordset
                                oRecordSet = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                ' SQL Query to get the details
                                Dim query As String = "SELECT T0.""OnHand"", T0.""AvgPrice"" FROM OITW T0 WHERE T0.""ItemCode"" = '" & itemCode & "' AND T0.""WhsCode"" = '" & warehouseCode & "'"
                                oRecordSet.DoQuery(query)
                                ' Check if data exists
                                If oRecordSet.RecordCount > 0 Then
                                    oRecordSet.MoveFirst()
                                    ' Get values from the query
                                    oDS1.SetValue("LineId", oDS1.Offset, pVal.Row)

                                    Dim inStockQty As String = oRecordSet.Fields.Item("OnHand").Value.ToString()
                                    Dim avgPrice As String = oRecordSet.Fields.Item("AvgPrice").Value.ToString()
                                    ' Set the matrix values
                                    objMatrix.Columns.Item("ISTQTY").Cells.Item(pVal.Row).Specific.Value = inStockQty
                                    objMatrix.Columns.Item("OLDCST").Cells.Item(pVal.Row).Specific.Value = avgPrice
                                    'objMatrix.Columns.Item("NEWCST").Cells.Item(pVal.Row).Specific.Value = avgPrice
                                Else
                                    objMain.objApplication.MessageBox("No data found for the selected Item and Warehouse.")
                                End If
                            Else
                                objMain.objApplication.MessageBox("Item Code or Warehouse is empty.")
                            End If
                            ' Set other values as before
                            oDS1.SetValue("LineId", oDS1.Offset, pVal.Row)
                            oDS1.SetValue("U_ITMCODE", oDS1.Offset, itemCode)
                            oDS1.SetValue("U_WRHSE", oDS1.Offset, warehouseCode)
                            oDS1.SetValue("U_ISTQTY", oDS1.Offset, objMatrix.Columns.Item("ISTQTY").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_OLDCST", oDS1.Offset, objMatrix.Columns.Item("OLDCST").Cells.Item(pVal.Row).Specific.Value)
                            oDS1.SetValue("U_NEWCST", oDS1.Offset, objMatrix.Columns.Item("NEWCST").Cells.Item(pVal.Row).Specific.Value)
                            objMatrix.SetLineData(pVal.Row)
                            If objMatrix.VisualRowCount = pVal.Row Then
                                Me.SetNewLine(objForm.UniqueID, "RDMtx")
                            End If
                            'Me.SetNewL
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "RDMtx" And pVal.ColUID = "ITMCODE" And pVal.BeforeAction = False Then
                        Try
                            objForm.Freeze(True)
                            Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("RDMtx").Specific
                            If pVal.ItemUID = "RDMtx" Then
                                If pVal.Row = objMatrix.VisualRowCount Then
                                    Me.SetNewLine(objForm.UniqueID, "RDMtx")
                                End If
                            End If
                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                        End Try
                    End If






                    'working code
                    'If (pVal.ColUID = "OLDCST" Or pVal.ColUID = "NEWCST") And pVal.BeforeAction = False Then
                    '    Dim TotalCostDifference As Double = 0
                    '    For i As Integer = 1 To objMatrix.RowCount

                    '        Dim oldCostStr As String = objMatrix.Columns.Item("OLDCST").Cells.Item(i).Specific.Value
                    '        Dim newCostStr As String = objMatrix.Columns.Item("NEWCST").Cells.Item(i).Specific.Value
                    '        If oldCostStr <> "" And newCostStr <> "" Then
                    '            Dim oldCost As Double = CDbl(oldCostStr)
                    '            Dim newCost As Double = CDbl(newCostStr)
                    '            ' Calculate cost difference
                    '            Dim costDiff As Double = Math.Abs(oldCost - newCost)
                    '            objMatrix.Columns.Item("COSTD").Cells.Item(i).Specific.Value = costDiff.ToString("F2")
                    '            TotalCostDifference += costDiff
                    '        End If

                    '    Next
                    '    oDS.SetValue("U_TOTAL", oDS.Offset, TotalCostDifference.ToString("F2"))
                    'End If




                    'If (pVal.ColUID = "OLDCST" Or pVal.ColUID = "NEWCST" Or pVal.ColUID = "ISTQTY" Or pVal.ColUID = "CSTDTL") And pVal.BeforeAction = False Then
                    '    Dim TotalCostDifference As Double = 0
                    '    For i As Integer = 1 To objMatrix.RowCount
                    '        Dim oldCostStr As String = objMatrix.Columns.Item("OLDCST").Cells.Item(i).Specific.Value
                    '        Dim newCostStr As String = objMatrix.Columns.Item("NEWCST").Cells.Item(i).Specific.Value
                    '        Dim IEMSTOCKStr As String = objMatrix.Columns.Item("ISTQTY").Cells.Item(i).Specific.Value

                    '        If oldCostStr <> "" And newCostStr <> "" And IEMSTOCKStr <> "" Then
                    '            Dim oldCost As Double = CDbl(oldCostStr)
                    '            Dim newCost As Double = CDbl(newCostStr)
                    '            Dim IEMSTOCK As Double = CDbl(IEMSTOCKStr)

                    '            ' Calculate cost difference
                    '            Dim costDiff As Double = Math.Abs(oldCost - newCost)

                    '            ' Set individual cost difference
                    '            objMatrix.Columns.Item("COSTD").Cells.Item(i).Specific.Value = costDiff.ToString("F2")

                    '            ' Multiply cost difference by stock quantity for total
                    '            Dim lineCostDiff As Double = costDiff * IEMSTOCK
                    '            TotalCostDifference += lineCostDiff
                    '            objMatrix.Columns.Item("CSTDTL").Cells.Item(i).Specific.Value = TotalCostDifference

                    '        End If
                    '    Next

                    '    oDS.SetValue("U_TOTAL", oDS.Offset, TotalCostDifference.ToString("F2"))
                    'End If

                    If (pVal.ColUID = "OLDCST" Or pVal.ColUID = "NEWCST" Or pVal.ColUID = "ISTQTY") And pVal.BeforeAction = False Then
                        Dim TotalCostDifference As Double = 0

                        For i As Integer = 1 To objMatrix.RowCount
                            Dim oldCostStr As String = objMatrix.Columns.Item("OLDCST").Cells.Item(i).Specific.Value
                            Dim newCostStr As String = objMatrix.Columns.Item("NEWCST").Cells.Item(i).Specific.Value
                            Dim IEMSTOCKStr As String = objMatrix.Columns.Item("ISTQTY").Cells.Item(i).Specific.Value

                            If oldCostStr <> "" And newCostStr <> "" And IEMSTOCKStr <> "" Then
                                Dim oldCost As Double = 0
                                Dim newCost As Double = 0
                                Dim IEMSTOCK As Double = 0

                                Double.TryParse(oldCostStr, oldCost)
                                Double.TryParse(newCostStr, newCost)
                                Double.TryParse(IEMSTOCKStr, IEMSTOCK)

                                ' Calculate cost difference
                                Dim costDiff As Double = Math.Abs(oldCost - newCost)
                                objMatrix.Columns.Item("COSTD").Cells.Item(i).Specific.Value = costDiff.ToString("F2")

                                ' cost difference * Item stock quantity
                                Dim lineCostDiff As Double = costDiff * IEMSTOCK
                                objMatrix.Columns.Item("CSTDTL").Cells.Item(i).Specific.Value = lineCostDiff.ToString("F2")

                                ' Accumulate for total
                                TotalCostDifference += lineCostDiff
                            Else
                                objMatrix.Columns.Item("COSTD").Cells.Item(i).Specific.Value = ""
                                objMatrix.Columns.Item("CSTDTL").Cells.Item(i).Specific.Value = ""
                            End If
                        Next

                        ' final total cost difference
                        oDS.SetValue("U_TOTAL", oDS.Offset, TotalCostDifference.ToString("F2"))
                    End If







            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub



    'Akhila
    Public Sub PostInventoryRevaluation(ByVal FormUID As String)
        ' errorMessage = ""
        Try
            Dim oMaterialRevaluation As SAPbobsCOM.MaterialRevaluation
            ' Header information 
            Dim objRecordSet11 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim query As String = "SELECT T0.""U_TOTAL"", T1.""U_CSTDTL"", T0.""U_BRANCH"",T0.""DocNum"", T1.""U_PSTS"",T1.""U_REF1"", T1.""U_REF2"", T1.""U_IREVNO"" " &
                ", T1.""U_JRMRK"", T1.""U_ITMCODE"", T1.""U_OLDCST"", T1.""U_NEWCST"", T0.""U_GLACNT"", T1.""U_WRHSE"",T1.""LineId"" " &
                " FROM ""@TNX_RDCTN"" T0 INNER JOIN ""@TNX_RDCTN_C0"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" WHERE T1.""U_ITMCODE"" <> ''  and 
              IFNULL(T1.""U_IREVNO"",'') = '' and IFNULL(T1.""U_PSTS"",'')='Open'"
            objRecordSet11.DoQuery(query)

            If objRecordSet11.RecordCount > 0 Then
                For j As Integer = 1 To objRecordSet11.RecordCount
                    Dim oMaterialRevaluationSNBLines As SAPbobsCOM.SNBLines
                    oMaterialRevaluation = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
                    oMaterialRevaluation.DocDate = System.DateTime.Now
                    oMaterialRevaluation.RevalType = "P"

                    Dim Item As String = objRecordSet11.Fields.Item("U_ITMCODE").Value
                    Dim DocNum As String = objRecordSet11.Fields.Item("DocNum").Value

                    'Line information
                    Dim InvtyNo As String = objRecordSet11.Fields.Item("U_IREVNO").Value
                    Dim Comments As String = objRecordSet11.Fields.Item("U_REF1").Value
                    Dim ActualCost As Double = objRecordSet11.Fields.Item("U_OLDCST").Value
                    Dim Newcost As Double = objRecordSet11.Fields.Item("U_NEWCST").Value
                    Dim Warehouse As String = objRecordSet11.Fields.Item("U_WRHSE").Value
                    Dim GLAccnt As String = objRecordSet11.Fields.Item("U_GLACNT").Value
                    Dim Ref1 As String = objRecordSet11.Fields.Item("U_REF1").Value
                    Dim Ref2 As String = objRecordSet11.Fields.Item("U_REF2").Value
                    Dim Jonrlrmrk As String = objRecordSet11.Fields.Item("U_JRMRK").Value
                    Dim Cline As String = objRecordSet11.Fields.Item("LineId").Value
                    Dim Branch As String = objRecordSet11.Fields.Item("U_BRANCH").Value
                    Dim Docnum2 As String = objRecordSet11.Fields.Item("DocNum").Value

                    Dim GLInAct As String = objRecordSet11.Fields.Item("U_CSTDTL").Value
                    Dim TotalAmt As String = objRecordSet11.Fields.Item("U_TOTAL").Value

                    ' Step 2: Get Revaluation Accounts from OGAR table
                    'Dim glAccRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    ''Dim glQuery As String = "SELECT ""StockAct"" FROM OGAR WHERE ""ItemCode"" = '" & Item & "' AND ""WhsCode"" = '" & Warehouse & "'"
                    'Dim glQuery As String = "Select T0.""StockAct"", T1.""ItmsGrpCod"" FROM OGAR T0  INNER JOIN OITM T1 ON T0.""ItemCode"" = T1.""ItemCode"" Where ""ItemCode"" = '" & Item & "'"
                    'glAccRs.DoQuery(glQuery)

                    'If glAccRs.RecordCount > 0 Then
                    '    Dim RevalIncAcct As String = glAccRs.Fields.Item("StockAct").Value.ToString().Trim()
                    '    oMaterialRevaluation.Lines.RevaluationIncrementAccount = RevalIncAcct
                    'Else

                    'End If

                    oMaterialRevaluation.Reference2 = Docnum2
                    oMaterialRevaluation.Lines.ItemCode = Item
                    oMaterialRevaluation.Lines.WarehouseCode = Warehouse
                    oMaterialRevaluation.Comments = Comments
                    oMaterialRevaluation.JournalMemo = Jonrlrmrk
                    oMaterialRevaluation.Lines.RevaluationDecrementAccount = GLAccnt
                    oMaterialRevaluation.Lines.RevaluationIncrementAccount = GLAccnt
                    ' oMaterialRevaluation.Lines.RevaluationDecrementAccount = GLAccnt
                    'oMaterialRevaluation.Lines.RevaluationIncrementAccount = GLAccnt
                    oMaterialRevaluation.Lines.Price = Newcost
                    oMaterialRevaluation.Lines.Add()

                    Dim result As Integer = oMaterialRevaluation.Add()
                    If result <> 0 Then
                        Dim errCode As Integer
                        Dim errMsg As String
                        objMain.objCompany.GetLastError(errCode, errMsg)
                        objMain.objApplication.MessageBox("Error " & errCode & ": " & errMsg)
                    Else
                        Dim DocEntry As Integer
                        DocEntry = objMain.objCompany.GetNewObjectKey()

                        Dim updateQuery As String = "UPDATE ""@TNX_RDCTN_C0"" Set ""U_IREVNO"" = '" & DocEntry.ToString() & "' , ""U_PSTS""='Posted'  WHERE ifnull(""U_IREVNO"",'') = '' AND ""LineId"" = '" & Cline.ToString() & "'   and ""DocEntry"" = (SELECT ""DocEntry"" FROM ""@TNX_RDCTN""  WHERE ""DocNum"" = '" & DocNum & "')"
                        Dim objRecord As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objRecord.DoQuery(updateQuery)

                        ' ✅ Success Message
                        objMain.objApplication.StatusBar.SetText("Inventory Revaluation is successfully added for item " & Item & " with Revaluation No: " & DocEntry, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    End If

                    objRecordSet11.MoveNext()
                Next
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    'Public Sub PostInventoryRevaluation(ByVal FormUID As String)
    '    ' errorMessage = ""
    '    Try
    '        Dim oMaterialRevaluation As SAPbobsCOM.MaterialRevaluation
    '        ' Header information

    '        Dim objRecordSet11 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        Dim query As String = "SELECT T1.""U_BRANCH"",T0.""DocNum"", T1.""U_PSTS"",T1.""U_REF1"", T1.""U_REF2"", T1.""U_IREVNO"", T1.""U_JRMRK"", T1.""U_ITMCODE"", T1.""U_OLDCST"", T1.""U_NEWCST"", T0.""U_GLACNT"", T1.""U_WRHSE"",T1.""LineId"" FROM ""@TNX_RDCTN"" T0 INNER JOIN ""@TNX_RDCTN_C0"" T1 ON T0.""DocEntry"" = T1.""DocEntry"" WHERE T1.""U_ITMCODE"" <> ''  "
    '        objRecordSet11.DoQuery(query)
    '        If objRecordSet11.RecordCount > 0 Then
    '            For j As Integer = 1 To objRecordSet11.RecordCount
    '                ' Dim oMaterialRevaluation As SAPbobsCOM.MaterialRevaluation
    '                Dim oMaterialRevaluationSNBLines As SAPbobsCOM.SNBLines
    '                oMaterialRevaluation = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation)
    '                oMaterialRevaluation.DocDate = System.DateTime.Now
    '                oMaterialRevaluation.RevalType = "P"
    '                Dim Item As String = objRecordSet11.Fields.Item("U_ITMCODE").Value
    '                Dim DocNum As String = objRecordSet11.Fields.Item("DocNum").Value
    '                ' Dim rec1 As String = "SELECT ""U_ITMCODE"" FROM ""@TNX_RDCTN_C0""  WHERE ""U_ITMCODE"" ='" & Item & "'  "
    '                Dim rec1 As String = "SELECT ""ItemCode"" FROM ""MRV1""  WHERE ""ItemCode"" ='" & Item & "'  "
    '                Dim rec As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                rec.DoQuery(rec1)
    '                Dim itemcode As String = rec.Fields.Item("ItemCode").Value
    '                If itemcode = "" Then
    '                    'Line information
    '                    Dim InvtyNo As String = objRecordSet11.Fields.Item("U_IREVNO").Value
    '                    Dim Comments As String = objRecordSet11.Fields.Item("U_REF1").Value
    '                    Dim ActualCost As Double = objRecordSet11.Fields.Item("U_OLDCST").Value
    '                    Dim Newcost As Double = objRecordSet11.Fields.Item("U_NEWCST").Value
    '                    Dim Warehouse As String = objRecordSet11.Fields.Item("U_WRHSE").Value
    '                    Dim GLAccnt As String = objRecordSet11.Fields.Item("U_GLACNT").Value
    '                    Dim Ref1 As String = objRecordSet11.Fields.Item("U_REF1").Value
    '                    Dim Ref2 As String = objRecordSet11.Fields.Item("U_REF2").Value
    '                    Dim Jonrlrmrk As String = objRecordSet11.Fields.Item("U_JRMRK").Value
    '                    Dim Cline As String = objRecordSet11.Fields.Item("LineId").Value
    '                    Dim Branch As String = objRecordSet11.Fields.Item("U_BRANCH").Value

    '                    Dim Docnum2 As String = objRecordSet11.Fields.Item("DocNum").Value
    '                    oMaterialRevaluation.Reference2 = Docnum2


    '                    oMaterialRevaluation.Lines.ItemCode = Item

    '                    oMaterialRevaluation.Lines.WarehouseCode = Warehouse
    '                    oMaterialRevaluation.Comments = Comments
    '                    oMaterialRevaluation.JournalMemo = Jonrlrmrk
    '                    oMaterialRevaluation.Lines.RevaluationDecrementAccount = GLAccnt
    '                    oMaterialRevaluation.Lines.RevaluationIncrementAccount = GLAccnt
    '                    'oMaterialRevaluation.Lines.ActualPrice = Newcost
    '                    oMaterialRevaluation.Lines.Price = Newcost
    '                    ' oMaterialRevaluation.Reference2 = Ref2
    '                    'oMaterialRevaluation.Reference1 = Ref1
    '                    oMaterialRevaluation.Lines.Add()
    '                    Dim result As Integer = oMaterialRevaluation.Add()
    '                    If result <> 0 Then
    '                        'errorMessage = objMain.objCompany.GetLastErrorDescription()
    '                        'Return False
    '                        Dim errCode As Integer
    '                        Dim errMsg As String
    '                        objMain.objCompany.GetLastError(errCode, errMsg)
    '                        objMain.objApplication.MessageBox("Error " & errCode & ": " & errMsg)

    '                    Else
    '                    End If

    '                    Dim DocEntry As Integer
    '                    DocEntry = objMain.objCompany.GetNewObjectKey()
    '                    Dim updateQuery As String = "UPDATE ""@TNX_RDCTN_C0"" Set ""U_IREVNO"" = '" & DocEntry.ToString() & "' , ""U_PSTS""='Posted'  WHERE ifnull(""U_IREVNO"",'') = '' AND ""LineId"" = '" & Cline.ToString() & "'   and ""DocEntry"" = (SELECT ""DocEntry"" FROM ""@TNX_RDCTN""  WHERE ""DocNum"" = '" & DocNum & "')"
    '                    Dim objRecord As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                    objRecord.DoQuery(updateQuery)
    '                End If
    '                objRecordSet11.MoveNext()
    '            Next
    '        End If
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '        ' Return False
    '    End Try
    'End Sub





    'Akhila


    Public Shared Sub FILTER_CFL(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef oApplication As SAPbouiCOM.Application, ByRef oCompany As SAPbobsCOM.Company, ByVal Query As String, ByVal Code As String, ByVal CFL_Id As String)
        Try
            Dim oCLFEvent As SAPbouiCOM.ChooseFromListEvent = Nothing
            Dim oCell As SAPbouiCOM.Column
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFL_collection As SAPbouiCOM.ChooseFromListCollection
            Dim oCon As SAPbouiCOM.Condition = Nothing
            Dim emptyConds As SAPbouiCOM.Conditions = Nothing
            Dim oCons As SAPbouiCOM.Conditions = Nothing
            Dim oRec As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRec.DoQuery(Query)
            Dim strCFLID As String = CFL_Id
            oCLFEvent = CType((pVal), SAPbouiCOM.ChooseFromListEvent)
            strCFLID = oCLFEvent.ChooseFromListUID
            oCFL = oApplication.Forms.ActiveForm.ChooseFromLists.Item(strCFLID)
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            Dim rr As Integer = oRec.RecordCount


            If rr = 0 Then
                oCon = oCons.Add()
                oCon.[Alias] = Code
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = oRec.Fields.Item(Code).Value.ToString()
            Else


                For i As Integer = 0 To oRec.RecordCount - 1
                    oCon = oCons.Add()
                    oCon.[Alias] = Code
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRec.Fields.Item(Code).Value.ToString()

                    If i <> oRec.RecordCount - 1 Then
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    End If
                    oRec.MoveNext()
                Next
            End If


            oCFL.SetConditions(oCons)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)


        End Try


    End Sub
    'Akhila
    'Sub CFLFilter(ByVal FormUID As String, ByVal Warehouse As String)
    'Sub CFLFilter(ByVal FormUID As String, ByVal CFL_UID As String, ByVal branch As String)

    '    Try
    '        objMatrix = objForm.Items.Item("RDMtx").Specific

    '        objForm = objMain.objApplication.Forms.Item(FormUID)
    '        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        Dim Selection As String = "select ""WhsCode"",""WhsName"",""BPLid"" from OWHS where ""BPLid""='" & branch & "'"
    '        oRs.DoQuery(Selection)

    '        Dim oConditions As SAPbouiCOM.Conditions
    '        Dim oCondition As SAPbouiCOM.Condition
    '        Dim oChooseFromList As SAPbouiCOM.ChooseFromList
    '        Dim emptyCon As New SAPbouiCOM.Conditions
    '        oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_UID)
    '        oChooseFromList.SetConditions(emptyCon)
    '        oConditions = oChooseFromList.GetConditions()
    '        If oRs.RecordCount > 0 Then
    '            For i As Integer = 1 To oRs.RecordCount

    '                oCondition = oConditions.Add()
    '                oCondition.Alias = "WhsCode"
    '                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '                oCondition.CondVal = oRs.Fields.Item("WhsCode").Value
    '                oRs.MoveNext()

    '            Next
    '            oChooseFromList.SetConditions(oConditions)

    '        End If
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)

    '    End Try
    'End Sub




    Sub CFLFilter(ByVal FormUID As String, ByVal CFL_UID As String, ByVal branch As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim GetSelection As String = "select ""WhsCode"",""WhsName"",""BPLid"" from OWHS where ""BPLid""='" & branch & "'"

            Dim oRsGetSelectionItems As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oRsGetSelectionItems.DoQuery(GetSelection)

            Dim oConditions As SAPbouiCOM.Conditions

            Dim oCondition As SAPbouiCOM.Condition

            Dim oChooseFromList As SAPbouiCOM.ChooseFromList

            Dim emptyCon As New SAPbouiCOM.Conditions

            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_UID)

            oChooseFromList.SetConditions(emptyCon)

            oConditions = oChooseFromList.GetConditions()

            If oRsGetSelectionItems.RecordCount > 0 Then

                oCondition = oConditions.Add()

                oCondition.Alias = "WhsCode"

                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL

                oCondition.CondVal = oRsGetSelectionItems.Fields.Item(0).Value

                oChooseFromList.SetConditions(oConditions)

                oRsGetSelectionItems.MoveNext()

                For i As Integer = 1 To oRsGetSelectionItems.RecordCount - 1

                    oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR

                    oCondition = oConditions.Add()

                    oCondition.Alias = "WhsCode"

                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL

                    oCondition.CondVal = oRsGetSelectionItems.Fields.Item(0).Value

                    oChooseFromList.SetConditions(oConditions)

                    oRsGetSelectionItems.MoveNext()

                Next

            Else

                oCondition = oConditions.Add()

                oCondition.Alias = "WhsCode"

                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL

                oCondition.CondVal = ""

                oChooseFromList.SetConditions(oConditions)

            End If

        Catch ex As Exception

        End Try

    End Sub


















    Sub CFLFilterBarcode(ByVal FormUID As String, ByVal CFL_ID As String, ByVal Suppliefr As String)


        Try
            'objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim GetSelection As String = "SELECT B.""BcdCode"" " &
                                   "FROM ""OBCD"" B " &
                                   "INNER JOIN ""OITM"" I ON B.""ItemCode"" = I.""ItemCode"" " &
            "WHERE I.""CardCode"" = '" & objForm.Items.Item("SPCODE").Specific.Value & "'"

            Dim oRsGetSelectionItems As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetSelectionItems.DoQuery(GetSelection)

            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            If oRsGetSelectionItems.RecordCount > 0 Then
                oCondition = oConditions.Add()
                oCondition.Alias = "ItemCode"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = oRsGetSelectionItems.Fields.Item(0).Value
                oChooseFromList.SetConditions(oConditions)
                oRsGetSelectionItems.MoveNext()
                For i As Integer = 1 To oRsGetSelectionItems.RecordCount - 1
                    oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCondition = oConditions.Add()
                    oCondition.Alias = "ItemCode"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = oRsGetSelectionItems.Fields.Item(0).Value
                    oChooseFromList.SetConditions(oConditions)
                    oRsGetSelectionItems.MoveNext()
                Next
            Else
                oCondition = oConditions.Add()
                oCondition.Alias = "ItemCode"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = ""
                oChooseFromList.SetConditions(oConditions)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try
    End Sub
    'Akila 
    Sub CFLItem(ByVal FormUID As String, ByVal CFL_ID As String, ByVal Vendorcode As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            'Dim GetBarCode As String = " Select T0.""ItemCode"" ,T0.""ItemName"" From  OITM T0 " &
            '                           "  Where T0.""ItmsGrpCod"" " &
            '                           " =  '" & ItemGroup & "'"
            Dim GetItem As String = " SELECT T0.""ItemCode"", T0.""ItemName"" FROM OITM T0  INNER JOIN ITM2 T1 ON T0.""ItemCode"" = T1.""ItemCode"" WHERE T1.""VendorCode""='" & Vendorcode & "'"

            Dim oRsGetBarCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetBarCode.DoQuery(GetItem)
            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            If oRsGetBarCode.RecordCount > 0 Then
                oCondition = oConditions.Add()
                oCondition.Alias = "ItemCode"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = oRsGetBarCode.Fields.Item(0).Value
                oChooseFromList.SetConditions(oConditions)
                oRsGetBarCode.MoveNext()
                For i As Integer = 1 To oRsGetBarCode.RecordCount - 1
                    oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCondition = oConditions.Add()
                    oCondition.Alias = "ItemCode"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = oRsGetBarCode.Fields.Item(0).Value
                    oChooseFromList.SetConditions(oConditions)
                    oRsGetBarCode.MoveNext()
                Next
            Else
                oCondition = oConditions.Add()
                oCondition.Alias = "ItemCode"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = ""
                oChooseFromList.SetConditions(oConditions)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    'Warehouse branch load 















    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    'Public Function GetFieldValue(ByVal query As String) As Object
    '    Try
    '        Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oRs.DoQuery(query)
    '        Dim str As String = oRs.Fields.Item(0).Value.ToString()
    '        Return str
    '    Catch ex As Exception
    '        Return ""
    '    End Try
    'End Function

    Sub SetDefault(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oDS = Me.objForm.DataSources.DBDataSources.Add("@TNX_RDCTN")
            oDS.SetValue("DocNum", oDS.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TNX_RDCTNUDO", "Primary"))
            objForm.DataBrowser.BrowseBy = "DocNum"
            oDS.SetValue("U_STS", oDS.Offset, "Active")
            oDS.SetValue("U_DOCDATE", oDS.Offset, DateTime.Now.ToString("yyyyMMdd"))

            Dim objMatrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("RDMtx").Specific, SAPbouiCOM.Matrix)
            Dim oDS1 As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Item("@TNX_RDCTN_C0")

            oDS1 = Me.objForm.DataSources.DBDataSources.Add("@TNX_RDCTN_C0")
            'objMatrix = objForm.Items.Item("RDMtx").Specific
            'objMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

            'oDS1.SetValue("U_PSTS", oDS1.Offset, "Open")

            'objMatrix.Clear()
            oDS1.Clear()
            objMatrix.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "RDMtx")
            objMatrix.AutoResizeColumns()

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub SetNewLine(ByVal FormUID As String, ByVal MatrixUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDS = objForm.DataSources.DBDataSources.Item("@TNX_RDCTN")
            oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_RDCTN_C0")
            objMatrix = objForm.Items.Item("RDMtx").Specific

            Select Case MatrixUID
                Case "RDMtx"
                    objMatrix.AddRow()
                    oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
                    oDS1.SetValue("U_ITMCODE", oDS1.Offset, "")
                    oDS1.SetValue("U_ITMDESC", oDS1.Offset, "")
                    oDS1.SetValue("U_BARCD", oDS1.Offset, "")

                    oDS1.SetValue("U_WRHSE", oDS1.Offset, "")
                    oDS1.SetValue("U_ISTQTY", oDS1.Offset, "")
                    oDS1.SetValue("U_OLDCST", oDS1.Offset, "")
                    oDS1.SetValue("U_NEWCST", oDS1.Offset, "")
                    'oDS1.SetValue("U_GLACNT", oDS1.Offset, "")
                    oDS1.SetValue("U_JRMRK", oDS1.Offset, "")
                    ' oDS1.SetValue("U_BRANCH", oDS1.Offset, "")
                    oDS1.SetValue("U_REF1", oDS1.Offset, "")
                    oDS1.SetValue("U_REF2", oDS1.Offset, "")
                    oDS1.SetValue("U_IREVNO", oDS1.Offset, "")
                    oDS1.SetValue("U_PSTS", oDS1.Offset, "Open")
                    oDS1.SetValue("U_REVDRF", oDS1.Offset, "")
                    objMatrix.AutoResizeColumns()
                    objMatrix.SetLineData(objMatrix.VisualRowCount)
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Function Validate(ByVal FormUID As String, ByVal MatrixUID As String) As Boolean
        Try
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

#Region " Right Click Event"



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

    '                    If eventInfo.ItemUID = "RDMtx" And eventInfo.ColUID = "LineId" Then
    '                        Try
    '                            oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
    '                            oMenus = oMenuItem.SubMenus
    '                            If oMenus.Exists("Delete Row") = False Then
    '                                oCreationPackage.UniqueID = "Delete Row"
    '                                oCreationPackage.String = "Delete Row"
    '                                oCreationPackage.Enabled = True
    '                                oMenus.AddEx(oCreationPackage)
    '                            End If
    '                        Catch ex As Exception
    '                            objMain.objApplication.StatusBar.SetText(ex.Message)
    '                        End Try

    '                    Else
    '                        Try
    '                            oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
    '                            oMenus = oMenuItem.SubMenus
    '                            If oMenus.Exists("Delete Row") = True Then
    '                                objMain.objApplication.Menus.RemoveEx("Delete Row")
    '                            End If
    '                            If oMenus.Exists("Close") = True Then
    '                                objMain.objApplication.Menus.RemoveEx("Close")
    '                            End If
    '                        Catch ex As Exception
    '                            objMain.objApplication.StatusBar.SetText(ex.Message)
    '                        End Try
    '                    End If
    '                End If
    '            End If
    '        End If


    '        ' System.Diagnostics.Process.Start()
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub



    Sub CFLCardCodeFilter(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
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
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    'Sub CFLCardCodeFilt(ByVal FormUID As String, ByVal CFL_ID As String)
    '    Try



    '        objForm = objMain.objApplication.Forms.Item(FormUID)
    '        Dim oConditions As SAPbouiCOM.Conditions
    '        Dim oCondition As SAPbouiCOM.Condition
    '        Dim oChooseFromList As SAPbouiCOM.ChooseFromList
    '        Dim emptyCon As New SAPbouiCOM.Conditions
    '        oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
    '        oChooseFromList.SetConditions(emptyCon)
    '        oConditions = oChooseFromList.GetConditions()
    '        oCondition = oConditions.Add()
    '        oCondition.Alias = "ItemCode"
    '        oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '        oCondition.CondVal = "S"
    '        oChooseFromList.SetConditions(oConditions)
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub


    'Sub CFLCardCodeFilter1(ByVal FormUID As String, ByVal CFL_ID As String)
    '    Try



    '        objForm = objMain.objApplication.Forms.Item(FormUID)
    '        Dim oConditions As SAPbouiCOM.Conditions
    '        Dim oCondition As SAPbouiCOM.Condition
    '        Dim oChooseFromList As SAPbouiCOM.ChooseFromList
    '        Dim emptyCon As New SAPbouiCOM.Conditions
    '        oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
    '        oChooseFromList.SetConditions(emptyCon)
    '        oConditions = oChooseFromList.GetConditions()
    '        oCondition = oConditions.Add()
    '        oCondition.Alias = "CardType"
    '        oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '        oCondition.CondVal = "S"
    '        oChooseFromList.SetConditions(oConditions)
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub

    Private Function ValidateBuuyerName(ByVal formUID As String) As Boolean

        Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(formUID)
        Dim oEditText As SAPbouiCOM.EditText = objForm.Items.Item("BYNAME").Specific
        Dim promoName As String = oEditText.Value.Trim()
        If String.IsNullOrWhiteSpace(promoName) Then
            objMain.objApplication.StatusBar.SetText("Buyer Name cannot be empty.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If

        Return True
    End Function
    Private Function ValidateSupplierName(ByVal formUID As String) As Boolean
        Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(formUID)
        Dim oEditText As SAPbouiCOM.EditText = objForm.Items.Item("SPNAME").Specific
        Dim promoName As String = oEditText.Value.Trim()

        If String.IsNullOrWhiteSpace(promoName) Then
            objMain.objApplication.StatusBar.SetText("Supplier Name cannot be empty.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        Return True
    End Function
    'Private Function ValidateDates(ByVal formUID As String) As Boolean

    '    If oDS.GetValue("U_PSTART", 0) = "" Then
    '        objMain.objApplication.StatusBar.SetText("Please Choose Start Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        Return False
    '    End If
    '    If oDS.GetValue("U_PEND", 0) = "" Then
    '        objMain.objApplication.StatusBar.SetText("Please Choose End Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        Return False
    '    End If
    '    Return True

    'End Function


    Function ValidateDates(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDS = objForm.DataSources.DBDataSources.Item("@TNX_RDCTN")
            'If objForm.Items.Item("5").Specific.Value = "" Then
            '    objMain.objApplication.StatusBar.SetText("MaterialRequestNo Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False

            If objForm.Items.Item("PSTART").Specific.Value.Trim = "" Then
                objMain.objApplication.StatusBar.SetText("Please Choose Start Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If objForm.Items.Item("PEND").Specific.Value.Trim = "" Then
                objMain.objApplication.StatusBar.SetText("Please Choose End Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            If objForm.Items.Item("BRANCH").Specific.Value.Trim = "" Then
                objMain.objApplication.StatusBar.SetText("Please Select the Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function













    'Private Function ValidateRevaluationDocumentReference(ByVal formUID As String, ByVal MatrixUID As String) As Boolean
    '    Try
    '        objMatrix = objForm.Items.Item("RDMtx").Specific
    '        '  objMatrix1 = objForm.Items.Item("IGMtx1").Specific
    '        'objMatrix2 = objForm.Items.Item("SHMtx").Specific
    '        Select Case MatrixUID
    '            Case "RDMtx"
    '                'Dim type As String = objForm.Items.Item("PRTYP").Specific.Value
    '                'If type = "ItemWise" Then
    '                For i As Integer = 1 To objMatrix.VisualRowCount
    '                    If objMatrix.Columns.Item("REVDRF").Cells.Item(1).Specific.Value.ToString().Trim() = "" Then
    '                        'objMain.objApplication.StatusBar.SetText("Please Choose Different selection And press Load Button", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                        objMain.objApplication.StatusBar.SetText("Revaluation Document Reference cannot be null", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                        Return False
    '                    End If
    '                Next
    '        End Select
    '        Return True
    '    Catch ex As Exception
    '    End Try
    'End Function



#End Region
End Class



