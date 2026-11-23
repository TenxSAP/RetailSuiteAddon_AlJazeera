Imports SAPbobsCOM

Public Class Cls_DocumentDraw

#Region "Declarations"
    Dim objForm, objBaseForm, objITRForm As SAPbouiCOM.Form
    Dim objMatrix, objBaseMatrix, objITRMatrix As SAPbouiCOM.Matrix
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objGrid As SAPbouiCOM.Grid
    Dim objGridColumn As SAPbouiCOM.GridColumn
    Dim objGridColumns As SAPbouiCOM.GridColumns
    Dim objComboGrid As SAPbouiCOM.ComboBoxColumn
    Dim objCheckBox As SAPbouiCOM.CheckBox
    Dim oDt As SAPbouiCOM.DataTable
    Dim sCardCode As String
    Dim sCardName As String
#End Region

    Sub CreateForm(ByVal FormUID As String, ByVal CardCode As String, ByVal CardName As String)
        Try
            objMain.objUtilities.LoadForm("DOCUMENTDRAW_ITM.xml", "ME_ILL", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_ILL", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objBaseForm = objMain.objApplication.Forms.Item(FormUID)
            sCardCode = CardCode
            sCardName = CardName
            If Retail.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = False Then
                Retail.MainCls.ohtLookUpForm.Add(objForm.UniqueID, FormUID)
            End If
            oDt = objForm.DataSources.DataTables.Add("dt")
            oDt = objForm.DataSources.DataTables.Item("dt")

            Me.LoadGrid(objForm.UniqueID, "N", CardCode)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub LoadGrid(ByVal FormUID As String, ByVal CheckValue As String, ByVal CARDCODE As String)
        Try
            objForm.Freeze(True)
            objMain.objApplication.StatusBar.SetText("Loading item list data, please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objGrid = objForm.Items.Item("Item_3").Specific
            objGrid.DataTable = oDt
            objGrid.DataTable.Clear()

            Dim Query As String = " CALL  ""TNX_ITEMLLIST""('" & CARDCODE & "')"
            objGrid.DataTable.ExecuteQuery(Query)
            If objGrid.Rows.Count = 0 Then
                objMain.objApplication.MessageBox("No records found")
                Exit Sub
            End If
            For i As Integer = 0 To objGrid.Columns.Count - 1
                objGrid.Columns.Item(i).Editable = False
            Next
            'objGrid.Columns.Item("CheckBox").Editable = False
            '    objGrid.Columns.Item("CheckBox").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            '   objGrid.Columns.Item("CheckBox").Editable = True
            Dim oEditTextCol As SAPbouiCOM.EditTextColumn
            oEditTextCol = objGrid.Columns.Item("CardCode")
            oEditTextCol.LinkedObjectType = 2
            Dim oEditTextCo As SAPbouiCOM.EditTextColumn
            oEditTextCo = objGrid.Columns.Item("AppID")
            oEditTextCo.LinkedObjectType = 4



            objGrid.AutoResizeColumns()
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Item list loaded successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "4" And pVal.BeforeAction = False Then
                        'objCheckBox = objForm.Items.Item("4").Specific
                        objForm.Freeze(True)
                        objMain.objItemList.CreateForm(objForm.UniqueID, sCardCode, sCardName, "N")
                        objForm.Freeze(False)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                    If pVal.ItemUID = "Item_3" And pVal.ColUID = "AppID" And pVal.BeforeAction = True Then
                        objForm = objMain.objApplication.Forms.Item(FormUID)
                        objGrid = objForm.Items.Item("Item_3").Specific
                        BubbleEvent = False
                        Dim row As Integer = pVal.Row
                        Dim AppID As String = objGrid.DataTable.GetValue("AppID", row)
                        Dim CardCode As String = objGrid.DataTable.GetValue("CardCode", row).ToString()
                        Dim CardName As String = objGrid.DataTable.GetValue("CardName", row).ToString()
                        Dim Series As String = objGrid.DataTable.GetValue("Series", row).ToString()


                        If AppID <> "" Then

                            objMain.objApplication.StatusBar.SetText("Opening Item list data , please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                            objMain.objItemList.CreateForm(objForm.UniqueID, CardCode, CardName, "Y", AppID)


                        End If
                    End If





                    '================ SUPPLIER REBATE =================


                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                    Try
                        If pVal.BeforeAction = False Then
                            objForm = objMain.objApplication.Forms.Item(FormUID)
                            objGrid = objForm.Items.Item("Item_3").Specific
                            If pVal.ItemUID = "Item_3" And pVal.ColUID = "RowsHeader" Then
                                Dim rowIndex As Integer = pVal.Row
                                If rowIndex >= 0 Then
                                    ' Get values from DataTable
                                    Dim CardCode As String = objGrid.DataTable.GetValue("CardCode", rowIndex).ToString()
                                    Dim CardName As String = objGrid.DataTable.GetValue("CardName", rowIndex).ToString()
                                    Dim AppID As String = objGrid.DataTable.GetValue("AppID", rowIndex).ToString()
                                    Dim Series As String = objGrid.DataTable.GetValue("Series", rowIndex).ToString()

                                    objMain.objApplication.StatusBar.SetText("Opening Item list data , please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                                    objMain.objItemList.CreateForm(objForm.UniqueID, CardCode, CardName, "Y", AppID)
                                    '       objForm.Close()
                                End If

                            End If

                        End If

                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                Case SAPbouiCOM.BoEventTypes.et_CLICK

            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub



End Class
