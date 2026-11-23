Imports System.Net.Http
Imports SAPbobsCOM
Public Class clsBarcodeRequest

#Region "    Declaration   "
    Dim objForm, objFormX As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix

    Dim oEditText As SAPbouiCOM.EditText
    Dim oDt, oDt1, oDataTable As SAPbouiCOM.DataTable
    Dim objGrid, oGrid As SAPbouiCOM.Grid
    Dim sSupplierCode, sSupplierName As String
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim MATRIXS As String
    Dim LostFocusFlag As Boolean = False


#End Region

#Region "    Public Methods   "
    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("BarCodereq.xml", "BarcodeRequest", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("BarcodeRequest", objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.EnableMenu("774", True)
            objForm.EnableMenu("1292", True)
            'objForm.EnableMenu("1293", True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_BCR")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_BCR_C0")
            objForm.Freeze(True)

            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TNX_BCRUDO"))
            oDBs_Head.SetValue("U_BRD", 0, DateTime.Now.ToString("yyyyMMdd"))

            objForm.DataBrowser.BrowseBy = "DocNum"

            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            Me.SetNewLine(objForm.UniqueID)
            'Me.SetDefault(objForm.UniqueID)
            objForm.Freeze(False)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("clsCategorySplit.CreateForm: " & ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region


#Region "    Events   "

    Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) And pVal.ActionSuccess = True Then

                        Me.SetDefault(objForm.UniqueID)
                    End If


                    If pVal.ItemUID = "BarcodeBtn" AndAlso pVal.BeforeAction = False Then

                        objMatrix = objForm.Items.Item("Mtx").Specific
                        Dim barcodeValue As String = objMatrix.Columns.Item("Barcode").Cells.Item(1).Specific.Value
                        If barcodeValue = "Yes" Then
                            objMain.objApplication.StatusBar.SetText("Barcode will be generated for this product to PDT.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        Else
                            objMain.objApplication.StatusBar.SetText("Please select 'Yes' in the Barcode dropdown to generate a barcode.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUID)
                    Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx").Specific

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    'objForm.Freeze(True)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_BCR")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_BCR_C0")
                    objMatrix = objForm.Items.Item("Mtx").Specific
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal        'CHOOSE FORM LIST'
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                    'If oCFL.UniqueID = "CFL_0" And pVal.BeforeAction = True Then 'CHOOSE FORM TABLE UNIQUNE ID
                    '    CFLFilter(FormUID, "CFL_0")
                    'End If
                    'submit

                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        If oCFL.UniqueID = "CFL_0" Then
                            oDBs_Details.SetValue("U_ItemCode", oDBs_Details.Offset, oDT.GetValue("ItemCode", 0))
                            oDBs_Details.SetValue("U_ItemDes", oDBs_Details.Offset, oDT.GetValue("ItemName", 0))
                            objMatrix.SetLineData(pVal.Row)
                            Me.SetNewLine(objForm.UniqueID)
                        End If

                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) And pVal.BeforeAction = False Then
                        Retail.MainCls.ohtLookUpForm.Remove(objForm.UniqueID)
                    End If

            End Select

        Catch ex As Exception
            Try : objForm.Freeze(False) : Catch : End Try
            objMain.objApplication.StatusBar.SetText("clsBarcodeRequest.ItemEvent: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "BCR" And pVal.BeforeAction = False Then
                Me.CreateForm()
            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                Me.SetNewLine(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1293" And pVal.BeforeAction = False Then

            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)


            ElseIf pVal.MenuUID = "774" And pVal.BeforeAction = False Then

                objMatrix = objForm.Items.Item("Mtx").Specific

                Dim row As Integer = objMatrix.VisualRowCount

                If objMatrix.IsRowSelected(1) <> True And objMatrix.VisualRowCount < 1 Then

                    objMatrix.AddRow()

                    oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)

                    objMatrix.SetLineData(objMatrix.VisualRowCount)

                End If

                If objMatrix.IsRowSelected(row) = True Then

                    objMatrix.DeleteRow(row)

                Else

                    For i As Integer = 1 To objMatrix.VisualRowCount - 1

                        If objMatrix.IsRowSelected(i) = True Then

                            objMatrix.DeleteRow(i)

                        End If

                    Next

                End If

                For i As Integer = 1 To objMatrix.VisualRowCount

                    objMatrix.Columns.Item("LineId").Cells.Item(i).Specific.Value = i

                Next
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If
            End If


        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub


    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_BCR")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_BCR_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific

            objMatrix.AddRow()

            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
            oDBs_Details.SetValue("U_PONum", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_ItemCode", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_ItemDes", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_New", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Packing", oDBs_Details.Offset, "")

            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_BCR")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_BCR_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific

            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TNX_BCRUDO", "Primary"))
            oDBs_Head.SetValue("U_BRD", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))


            'objMatrix.Clear()
            'oDBs_Head.Clear()
            'oDBs_Details.Clear()
            'objMatrix.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID)
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try

            objForm = objMain.objApplication.Forms.GetForm("BarcodeRequest", objMain.objApplication.Forms.ActiveForm.TypeCount)

            Select Case BusinessObjectInfo.EventType

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                    If BusinessObjectInfo.BeforeAction = False AndAlso BusinessObjectInfo.ActionSuccess = True Then
                        Try
                            'oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_BCR")


                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText("Auto-approval error : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If



                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                    If BusinessObjectInfo.BeforeAction = False AndAlso BusinessObjectInfo.ActionSuccess = True Then
                        objForm.Freeze(True)

                        Try


                        Finally
                            objForm.Freeze(False)
                        End Try
                    End If
            End Select

        Catch ex As Exception
            Try : objForm.Freeze(False) : Catch : End Try
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region


End Class

