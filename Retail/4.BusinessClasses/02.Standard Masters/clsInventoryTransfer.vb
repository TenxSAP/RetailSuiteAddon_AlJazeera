Public Class clsInventoryTransfer

#Region "        Declaration        "
    Dim objForm, AltForm As SAPbouiCOM.Form
    Dim objMatrix As SAPbouiCOM.Matrix
#End Region

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.BeforeAction = False Then
                        Me.AddItems(objForm.UniqueID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.BeforeAction = False And pVal.ItemUID = "btn_LStr" Then
                        objMain.objStockTransferDetails.CreateForm(objForm.UniqueID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix = objForm.Items.Item("23").Specific
                    If pVal.BeforeAction = False And pVal.ItemUID = "23" And pVal.ColUID = "10" Then
                        Dim Qty As Double = CDbl(objMatrix.Columns.Item("10").Cells.Item(pVal.Row).Specific.Value)
                        Dim STRQTY As Double = CDbl(objMatrix.Columns.Item("U_METSTQTY").Cells.Item(pVal.Row).Specific.Value)
                        If Qty > 0 And STRQTY > 0 Then
                            If Qty > STRQTY Then
                                objMatrix.Columns.Item("10").Cells.Item(pVal.Row).Specific.Value = STRQTY
                                objMain.objApplication.StatusBar.SetText("Qty should not be greater than request qty  !!.... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                        End If
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub AddItems(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objMain.objUtilities.AddLabel(objForm.UniqueID, "TRTYPE", objForm.Items.Item("1470000099").Top + 30, objForm.Items.Item("1470000099").Left, objForm.Items.Item("1470000099").Width, "Transaction Type", "1470000099")
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "U_TRTYPE", objForm.Items.Item("1470000101").Top + 30, objForm.Items.Item("1470000101").Left, objForm.Items.Item("1470000101").Width, "OWTR", "U_TRTYPE", "1470000101")

            objMain.objUtilities.AddLabel(objForm.UniqueID, "INVTRREQ", objForm.Items.Item("1470000099").Top + 45, objForm.Items.Item("1470000099").Left, objForm.Items.Item("1470000099").Width, "StockRequestNo", "1470000099")
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "U_INVTRREQ", objForm.Items.Item("1470000101").Top + 45, objForm.Items.Item("1470000101").Left, (objForm.Items.Item("1470000101").Width / 2) - 1, "OWTR", "U_INVTRREQ", "1470000101")
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "U_INVTRREN", objForm.Items.Item("1470000101").Top + 45, objForm.Items.Item("1470000101").Left + objForm.Items.Item("U_INVTRREQ").Width + 1, objForm.Items.Item("1470000101").Width / 2, "OWTR", "U_INVTRREN", "1470000101")

            objForm.Items.Item("U_INVTRREQ").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("U_INVTRREQ").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("U_INVTRREN").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("U_INVTRREN").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("U_TRTYPE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("U_TRTYPE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            'objMain.objUtilities.AddButton(objForm.UniqueID, "btn_LStr", objForm.Items.Item("2").Top, _
            '                              objForm.Items.Item("2").Left + objForm.Items.Item("2").Width + 5, 120, objForm.Items.Item("2").Height, "2", "&Load STR Details")
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("940", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Try
                            Dim GetInventoryDtls As String = " Select ""DocEntry"" As ""ITEntry"" , ""U_INVTRREN"" As ""STREntery"" From WTR1 " &
                                                             " Where ""DocEntry"" = (Select Max(""DocEntry"")  From OWTR) And " & IsNull & "(""U_INVTRREN"",'') <> '' " &
                                                             " Group by  ""DocEntry"",  ""U_INVTRREN"" "
                            Dim oRsGetInventoryDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetInventoryDtls.DoQuery(GetInventoryDtls)
                            If oRsGetInventoryDtls.RecordCount > 0 Then
                                oRsGetInventoryDtls.MoveFirst()
                                For i As Integer = 1 To oRsGetInventoryDtls.RecordCount
                                    Me.UpdateSTR(objForm.UniqueID, oRsGetInventoryDtls.Fields.Item("ITEntry").Value, oRsGetInventoryDtls.Fields.Item("STREntery").Value)
                                    oRsGetInventoryDtls.MoveNext()
                                Next
                            End If
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Try
                            Dim GetInventoryDtls As String = " Select ""DocEntry"" As ""ITEntry"" , ""U_INVTRREN"" As ""STREntery"" From WTR1 " &
                                                             " Where ""DocEntry"" = (Select Max(""DocEntry"")  From OWTR) And " & IsNull & "(""U_INVTRREN"",'') <> '' " &
                                                             " Group by  ""DocEntry"",  ""U_INVTRREN"" """
                            Dim oRsGetInventoryDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetInventoryDtls.DoQuery(GetInventoryDtls)
                            If oRsGetInventoryDtls.RecordCount > 0 Then
                                oRsGetInventoryDtls.MoveFirst()
                                For i As Integer = 1 To oRsGetInventoryDtls.RecordCount
                                    Me.UpdateSTR(objForm.UniqueID, oRsGetInventoryDtls.Fields.Item("ITEntry").Value, oRsGetInventoryDtls.Fields.Item("STREntery").Value)
                                    oRsGetInventoryDtls.MoveNext()
                                Next
                            End If
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub UpdateSTR(ByVal FormUID As String, ByVal DocEntry As String, ByVal STREntry As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim GetDetails As String = " Select ""ItemCode""  , (""U_METSTQTY"" - (""U_MEREQQTY"" - ""Quantity"")) ""DispatchQty"" , ""U_MEREQQTY"" ""STRQty"", " & _
                                       " (""U_MEREQQTY"" - ""Quantity"") ""OpenQty"" , ""U_INVTRRLN"" As ""LineId"" , (""U_METSTQTY""-(""U_METSTQTY"" - (""U_MEREQQTY"" - ""Quantity""))) ""LineStatus""  From  WTR1 T1  Where ""DocEntry""  = '" & DocEntry & "' And ""Quantity"" <= ""U_MEREQQTY"""
            Dim oRsGetDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDetails.DoQuery(GetDetails)

            If oRsGetDetails.RecordCount > 0 Then
                oRsGetDetails.MoveFirst()
                objMain.sCmp = objMain.objCompany.GetCompanyService
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("ME_OSTR")
                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                objMain.oGeneralParams.SetProperty("DocEntry", STREntry)
                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                objMain.oChildren = objMain.oGeneralData.Child("ME_STR_C0")
                For i As Integer = 1 To oRsGetDetails.RecordCount
                    Dim LineId As Integer = CInt(oRsGetDetails.Fields.Item("LineId").Value)
                    objMain.oChildren.Item(LineId - 1).SetProperty("U_MERELQTY", oRsGetDetails.Fields.Item("DispatchQty").Value)
                    objMain.oChildren.Item(LineId - 1).SetProperty("U_MEOPNQTY", oRsGetDetails.Fields.Item("OpenQty").Value)
                    'If oRsGetDetails.Fields.Item("OpenQty").Value = 0 Then
                    objMain.oChildren.Item(LineId - 1).SetProperty("U_MELNSTS", "Close")
                    'Else
                    'objMain.oChildren.Item(LineId - 1).SetProperty("U_MELNSTS", "Open")
                    'End If
                    objMain.oGeneralService.Update(objMain.oGeneralData)
                    oRsGetDetails.MoveNext()
                Next
                'Dim GetOpenQty As String = "Select ""U_MEOPNQTY"" From ""@ME_STR_C0"" Where ""DocEntry"" = '" & STREntry & "' And ""U_MEOPNQTY"" > 0 And ""U_MEITMCD"" <> '' "
                'Dim oRsGetOpenQty As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oRsGetOpenQty.DoQuery(GetOpenQty)
                'If oRsGetOpenQty.RecordCount = 0 Then
                objMain.oGeneralData.SetProperty("U_MESTS", "Close")
                objMain.oGeneralService.Update(objMain.oGeneralData)
                'End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
