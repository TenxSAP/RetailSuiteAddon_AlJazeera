Public Class Configuration
    Private oForm As SAPbouiCOM.Form
    Dim oDBHeader, oDBDetail1 As SAPbouiCOM.DBDataSource
    Dim oItem, oItemUP, oItemUP1 As SAPbobsCOM.Items
    Dim Error_Code As Integer = 1
    Dim Error_Msg As String = ""

    Private Sub DataSourcesLoad(ByRef oForm As SAPbouiCOM.Form)
        Try
            oDBHeader = oForm.DataSources.DBDataSources.Item("@CONFIG")
        Catch ex As Exception
            objMain.objUtilities.Write_Log(ex.Message)
        End Try
    End Sub
    Sub FormLoad(ByRef oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            Dim Qry As String = "Select * From ""@CONFIG"""
            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery(Qry)
            If oRs.RecordCount > 0 Then
                objMain.objApplication.ActivateMenuItem("1290")
            Else
                SetDefault()
            End If
            oForm.Items.Item("Email").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case "CONFIG"
                    If pVal.BeforeAction = True Then
                        objMain.objUtilities.LoadForm("Configurations.xml", "CONFIG", ResourceType.Embeded)
                    Else
                        oForm = objMain.objApplication.Forms.ActiveForm
                        DataSourcesLoad(oForm)
                        FormLoad(oForm)

                    End If
                Case "1292" ''''''''''''''''''''''''''''''''''''''''''  Add Row
                    oForm = objMain.objApplication.Forms.ActiveForm
                    If pVal.BeforeAction = False Then
                        DataSourcesLoad(oForm)
                        oForm.Freeze(True)

                        oForm.Freeze(False)
                    End If
                Case "1293" ''''''''''''''''''''''''''''''''''''''''''  Delete Row
                    oForm = objMain.objApplication.Forms.ActiveForm
                    If pVal.BeforeAction = False Then
                        oForm.Freeze(True)
                        DataSourcesLoad(oForm)

                        oForm.Freeze(False)
                    End If
                Case "1281" '''''''''''''''''''''''''''''''''''''''''''  Find Mode
                    oForm = objMain.objApplication.Forms.ActiveForm
                    If pVal.BeforeAction = False Then
                    End If
                Case "1282" '''''''''''''''''''''''''''''''''''''''''''  Add Mode
                    oForm = objMain.objApplication.Forms.ActiveForm
                    If pVal.BeforeAction = False Then
                        DataSourcesLoad(oForm)
                        FormLoad(oForm)
                    End If
                Case "1290", "1288", "1289", "1291"  ''''''''''''''''''' First Next Prev Last Record
                    oForm = objMain.objApplication.Forms.ActiveForm
                    If pVal.BeforeAction = False Then
                        DataSourcesLoad(oForm)
                    End If

            End Select

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            oForm = objMain.objApplication.Forms.Item(pVal.FormUID)
            DataSourcesLoad(oForm)
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Select Case pVal.ItemUID
                        Case "1"
                            Try
                                If pVal.BeforeAction = True And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    oForm.Freeze(True)
                                    'Dim Query As String = "SELECT " + IsNull + "(MAX(""DocNum""),0)+1 FROM ""@CONFIG"""
                                    'Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    'rs.DoQuery(Query)
                                    'oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                    'oForm.Items.Item("DocNum").Specific.Value = rs.Fields.Item(0).Value
                                    'oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    oForm.Freeze(False)
                                End If
                            Catch ex As Exception
                                oForm.Freeze(False)
                            End Try
                    End Select

            End Select
        Catch ex As Exception
            oForm.Freeze(False)
            If (ex.Message = "Form - Invalid Form") Then
            Else
                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        End Try
    End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = objMain.objApplication.Forms.ActiveForm
            DataSourcesLoad(oForm)
            oForm.Freeze(True)

            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Private Function Validation() As Boolean
        Try


        Catch ex As Exception
        End Try

        Return True
    End Function
    Sub SetDefault()
        Try
            Dim Query As String
            If objMain.objCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                Query = "SELECT " + IsNull + "(MAX(TO_INTEGER (""DocNum"")),0)+1 FROM ""@CONFIG"""
            Else
                Query = "SELECT ISNULL(MAX(CAST(""DocNum"" AS int)),0)+1 FROM ""@CONFIG"""
            End If
            Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rs.DoQuery(Query)
            oDBHeader.SetValue("DocNum", 0, rs.Fields.Item(0).Value)
        Catch ex As Exception
        End Try
    End Sub

End Class
