Public Class clsIncomingPayment
 
#Region "  Declaration "
        Public objform As SAPbouiCOM.Form
        Dim oDBs_Head As SAPbouiCOM.DBDataSource
        Dim oItem As SAPbouiCOM.Item
        Dim objComboBox As SAPbouiCOM.ComboBox
        Dim objMatrix As SAPbouiCOM.Matrix
        Dim strValQry As String = ""
    Dim oRecvQRY, rs As SAPbobsCOM.Recordset
#End Region

        Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
            Try
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        objform = objMain.objApplication.Forms.Item(FormUID)
                        If pVal.BeforeAction = True Then
                        Me.AddItems(objform.UniqueID)
                            'Me.AddChooseFromListItms(objform.UniqueID, "CFL_ACCT")
                            'Me.CFLFilter("CFL_ACCT")
                        End If

                    'Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    '    objform = objMain.objApplication.Forms.Item(FormUID)
                    '    If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                    '                  (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                    '        'If Me.Validation(objform.UniqueID) = False Then
                    '        '    BubbleEvent = False
                    '        'End If
                    '        ' For i As Integer = 1 To oRsReceipt1.RecordCount
                    '        Dim UpdateSplitItems As String = " Update ""@SBO_OUTLET_C0"" Set ""U_Payment"" =(Select Max(""DocEntry"") FROM ORCT)'" &
                    '                                             " Where ""DocEntry""='" & oRsReceipt1.Fields.Item("DocEntry").Value & "' And " &
                    '                                             " ""LineNum"" = '" & oRsReceipt1.Fields.Item("LineNum").Value & "'  "
                    '        Dim oRsUpdateSplitItems As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '        oRsUpdateSplitItems.DoQuery(UpdateSplitItems)
                    '        oRsReceipt1.MoveNext()
                    '    Next
                    '    End If




                    'End If
            End Select
            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private Sub AddItems(ByVal FormUID As String)
            Try
                objform = objMain.objApplication.Forms.Item(FormUID)

            objMain.objUtilities.AddLabel(objform.UniqueID, "lblOT", objform.Items.Item("53").Top + 22, objform.Items.Item("53").Left, objform.Items.Item("53").Width, "Outlet Number", "53", objform.Items.Item("53").Height)
            objMain.objUtilities.AddEditBox(objform.UniqueID, "OutNum", objform.Items.Item("52").Top + 22, objform.Items.Item("52").Left, objform.Items.Item("52").Width, "ORCT", "U_OutNum", "lblOT", objform.Items.Item("52").Height)


            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText(ex.Message)
            End Try
        End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objform = objMain.objApplication.Forms.GetForm("170", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        'Dim Query2 As String = "Select  ""DocNum"",""OutNum"" FROM ORCT where ""DocEntry""= (Select Max(""DocEntry"") FROM ORCT) "
                        ''rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ''rs.DoQuery(Query2)
                        Dim DocNum As String = objform.Items.Item("OutNum").Specific.Value
                        Dim GetJEDocEntry As String = "SELECT *  FROM ""@SBO_OUTLET""  T0 inner join  ""@SBO_OUTLET_C0""  T1 on  T1.""DocEntry""  =T0.""DocEntry"" WHERE T0.""DocNum"" ='" & DocNum & "'"
                        Dim oRsGetDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetDocEntry.DoQuery(GetJEDocEntry)


                        Dim Query1 As String = "Update ""@SBO_OUTLET_C0"" Set ""U_Payment""=(Select  ""DocNum"" FROM ORCT where ""DocEntry""= (Select Max(""DocEntry"") FROM ORCT)) Where ""DocEntry""='" & oRsGetDocEntry.Fields.Item("DocEntry").Value & "'and ""LineId""='1'"
                        rs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        rs.DoQuery(Query1)

                    End If

            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    End Class
