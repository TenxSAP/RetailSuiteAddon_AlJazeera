Imports System.Collections.Generic
Imports System.Linq

Public Class clsRentalAgreement

#Region "  Declaration "
    Public objform As SAPbouiCOM.Form
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oItem As SAPbouiCOM.Item
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim strValQry As String = ""
    Dim oRecvQRY As SAPbobsCOM.Recordset
    Dim prevTypeDisValue As String = ""
    Private SelectedPromoRow As Integer = 0
    Private LoadedAbsID As String = ""
#End Region

#Region " Helper Functions "
    Private Function ItemExists(ByVal oForm As SAPbouiCOM.Form, ByVal ItemUID As String) As Boolean
        Try
            Dim itm As SAPbouiCOM.Item = oForm.Items.Item(ItemUID)
            Return (itm IsNot Nothing)
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function CFLExists(ByVal oForm As SAPbouiCOM.Form, ByVal CFL_ID As String) As Boolean
        Try
            Dim cfl As SAPbouiCOM.ChooseFromList = oForm.ChooseFromLists.Item(CFL_ID)
            Return (cfl IsNot Nothing)
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function UserDataSourceExists(ByVal oForm As SAPbouiCOM.Form, ByVal udsUID As String) As Boolean
        Try
            Dim uds As SAPbouiCOM.UserDataSource = oForm.DataSources.UserDataSources.Item(udsUID)
            Return (uds IsNot Nothing)
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub RenumberMatrixRows(ByVal oMatrix As SAPbouiCOM.Matrix)
        Try
            For r As Integer = 1 To oMatrix.RowCount
                Try
                    oMatrix.Columns.Item("colRowNo").Cells.Item(r).Specific.Value = CStr(r)
                Catch
                End Try
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Function GetAbsID(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByVal oForm As SAPbouiCOM.Form) As String
        Dim sAbsID As String = ""
        Try
            If BusinessObjectInfo IsNot Nothing AndAlso Not String.IsNullOrEmpty(BusinessObjectInfo.ObjectKey) Then
                Dim xmlDoc As New System.Xml.XmlDocument()
                xmlDoc.LoadXml(BusinessObjectInfo.ObjectKey)
                Dim node As System.Xml.XmlNode = xmlDoc.SelectSingleNode("//*[translate(local-name(), 'ABSID', 'absid')='absid' or translate(local-name(), 'AGREEMENTNO', 'agreementno')='agreementno']")
                If node IsNot Nothing Then
                    sAbsID = node.InnerText.Trim()
                End If
                If String.IsNullOrEmpty(sAbsID) Then
                    Dim m As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(BusinessObjectInfo.ObjectKey, "\d+")
                    If m.Success Then
                        sAbsID = m.Value
                    End If
                End If
            End If
        Catch ex As Exception
        End Try

        If String.IsNullOrEmpty(sAbsID) AndAlso oForm IsNot Nothing Then
            Try
                Dim oDBs As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("OOAT")
                If oDBs IsNot Nothing Then
                    sAbsID = oDBs.GetValue("AbsID", 0).Trim()
                End If
            Catch ex As Exception
            End Try
        End If

        If String.IsNullOrEmpty(sAbsID) OrElse sAbsID = "0" Then
            Try
                Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRS.DoQuery("SELECT MAX(""AbsID"") AS ""AbsID"" FROM ""OOAT""")
                If oRS.RecordCount > 0 Then
                    sAbsID = oRS.Fields.Item("AbsID").Value.ToString().Trim()
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS)
            Catch ex As Exception
            End Try
        End If

        Return sAbsID
    End Function
#End Region

#Region " Promotion Items DB Persistence (@RENTPROMO / @RENTPROMO_C0) "
    ''' <summary>
    ''' Saves promotion item details from the mtxPromo matrix into the custom UDTs.
    ''' Header:  @RENTPROMO   – one row per agreement, keyed by U_AgrNo.
    ''' Lines:   @RENTPROMO_C0 – child rows holding U_ItemCode and U_ItemName.
    ''' </summary>
    Public Sub SavePromotionDetails(ByVal sAbsID As String, ByVal oForm As SAPbouiCOM.Form)
        If String.IsNullOrEmpty(sAbsID) OrElse sAbsID = "0" OrElse oForm Is Nothing Then Exit Sub
        Try
            If Not ItemExists(oForm, "mtxPromo") Then Exit Sub
            Dim oPromoMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("mtxPromo").Specific
            Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            ' ── Collect non-empty promotion items from matrix ──
            Dim promoList As New List(Of Tuple(Of String, String))()
            For r As Integer = 1 To oPromoMatrix.RowCount
                Dim sCode As String = ""
                Dim sName As String = ""
                Try
                    sCode = oPromoMatrix.Columns.Item("U_ItemCode").Cells.Item(r).Specific.Value.ToString().Trim()
                Catch
                End Try
                Try
                    sName = oPromoMatrix.Columns.Item("U_ItemName").Cells.Item(r).Specific.Value.ToString().Trim()
                Catch
                End Try

                If Not String.IsNullOrEmpty(sCode) Then
                    promoList.Add(New Tuple(Of String, String)(sCode, sName))
                End If
            Next

            ' ── Upsert header row in @RENTPROMO ──
            Dim sHeaderDocEntry As String = ""
            oRS.DoQuery("SELECT ""DocEntry"" FROM ""@RENTPROMO"" WHERE ""U_AgrNo"" = '" & sAbsID.Replace("'", "''") & "'")

            If oRS.RecordCount > 0 Then
                ' Header already exists – remember DocEntry for child operations
                sHeaderDocEntry = oRS.Fields.Item("DocEntry").Value.ToString().Trim()
            Else
                ' Calculate next DocEntry
                oRS.DoQuery("SELECT " & IsNull & "(MAX(""DocEntry""), 0) + 1 AS ""NextDoc"" FROM ""@RENTPROMO""")
                Dim sNextDoc As String = oRS.Fields.Item("NextDoc").Value.ToString().Trim()

                ' Insert a new header row
                Dim sInsertHdr As String = "INSERT INTO ""@RENTPROMO"" (""DocEntry"", ""U_AgrNo"") VALUES (" &
                    sNextDoc & ", '" & sAbsID.Replace("'", "''") & "')"
                oRS.DoQuery(sInsertHdr)

                sHeaderDocEntry = sNextDoc
            End If

            If String.IsNullOrEmpty(sHeaderDocEntry) Then
                objMain.objApplication.StatusBar.SetText("Save Promotion: Could not resolve @RENTPROMO header.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            ' ── Delete existing child rows for this header ──
            oRS.DoQuery("DELETE FROM ""@RENTPROMO_C0"" WHERE ""DocEntry"" = " & sHeaderDocEntry)

            ' ── Insert fresh child rows from the matrix ──
            For i As Integer = 0 To promoList.Count - 1
                Dim itemCode As String = promoList(i).Item1.Replace("'", "''")
                Dim itemName As String = promoList(i).Item2.Replace("'", "''")

                Dim sInsertLine As String = "INSERT INTO ""@RENTPROMO_C0"" (""DocEntry"", ""LineId"", ""U_ItemCode"", ""U_ItemName"") VALUES (" &
                    sHeaderDocEntry & ", " & i & ", '" & itemCode & "', '" & itemName & "')"
                oRS.DoQuery(sInsertLine)
            Next

            LoadedAbsID = sAbsID
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Save Promotion Details Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    ''' <summary>
    ''' Loads promotion item details from @RENTPROMO / @RENTPROMO_C0 back into the mtxPromo matrix.
    ''' </summary>
    Public Sub LoadPromotionDetails(ByVal sAbsID As String, ByVal oForm As SAPbouiCOM.Form)
        If String.IsNullOrEmpty(sAbsID) OrElse sAbsID = "0" OrElse oForm Is Nothing Then Exit Sub
        Try
            If Not ItemExists(oForm, "mtxPromo") Then Exit Sub
            Dim oPromoMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("mtxPromo").Specific
            Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            ' Query child rows joined through the header to match on AgrNo
            Dim sQuery As String = "SELECT C.""LineId"", " & IsNull & "(C.""U_ItemCode"", '') AS ""U_ItemCode"", " &
                IsNull & "(C.""U_ItemName"", '') AS ""U_ItemName"" " &
                "FROM ""@RENTPROMO"" H INNER JOIN ""@RENTPROMO_C0"" C ON H.""DocEntry"" = C.""DocEntry"" " &
                "WHERE H.""U_AgrNo"" = '" & sAbsID.Replace("'", "''") & "' " &
                "AND " & IsNull & "(C.""U_ItemCode"", '') <> '' " &
                "ORDER BY C.""LineId"""
            oRS.DoQuery(sQuery)

            ' Save current form mode so loading doesn't change it
            Dim savedMode As SAPbouiCOM.BoFormMode = oForm.Mode

            oForm.Freeze(True)
            oPromoMatrix.Clear()

            Dim iRow As Integer = 1
            While Not oRS.EoF
                Dim sCode As String = oRS.Fields.Item("U_ItemCode").Value.ToString().Trim()
                Dim sName As String = oRS.Fields.Item("U_ItemName").Value.ToString().Trim()

                If Not String.IsNullOrEmpty(sCode) Then
                    oPromoMatrix.AddRow()
                    Try
                        oPromoMatrix.Columns.Item("colRowNo").Cells.Item(iRow).Specific.Value = CStr(iRow)
                    Catch
                    End Try
                    Try
                        oPromoMatrix.Columns.Item("U_ItemCode").Cells.Item(iRow).Specific.Value = sCode
                    Catch
                    End Try
                    Try
                        oPromoMatrix.Columns.Item("U_ItemName").Cells.Item(iRow).Specific.Value = sName
                    Catch
                    End Try
                    iRow += 1
                End If

                oRS.MoveNext()
            End While

            ' Ensure at least one empty row for user input
            Dim bNeedBlankRow As Boolean = True
            If oPromoMatrix.RowCount > 0 Then
                Dim sLastCode As String = ""
                Try
                    sLastCode = oPromoMatrix.Columns.Item("U_ItemCode").Cells.Item(oPromoMatrix.RowCount).Specific.Value.ToString().Trim()
                Catch
                End Try
                If String.IsNullOrEmpty(sLastCode) Then
                    bNeedBlankRow = False
                End If
            End If

            If bNeedBlankRow Then
                oPromoMatrix.AddRow()
                Dim newRow As Integer = oPromoMatrix.RowCount
                Try
                    oPromoMatrix.Columns.Item("U_ItemCode").Cells.Item(newRow).Specific.Value = ""
                Catch
                End Try
                Try
                    oPromoMatrix.Columns.Item("U_ItemName").Cells.Item(newRow).Specific.Value = ""
                Catch
                End Try
            End If

            RenumberMatrixRows(oPromoMatrix)
            LoadedAbsID = sAbsID

            ' Restore original form mode (loading should not dirty the form)
            If oForm.Mode <> savedMode Then
                oForm.Mode = savedMode
            End If

            oForm.Freeze(False)

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS)
        Catch ex As Exception
            oForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Load Promotion Details Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.BeforeAction = True Then
                        Me.AddItems(objform.UniqueID)
                        Me.AddChooseFromListItms(objform.UniqueID, "CFL_ACCT")
                        Me.CFLFilter("CFL_ACCT")
                    End If
                    If pVal.BeforeAction = False Then
                        LoadBranchCombo(FormUID)
                        LoadPaymentTerms(FormUID)
                        If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Me.PopulateRentalAccount(FormUID)
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUID)

                    If pVal.ItemUID = "TypeDis" AndAlso pVal.BeforeAction = False Then
                        Dim oCombo As SAPbouiCOM.ComboBox = oForm.Items.Item("TypeDis").Specific
                        If oCombo.Selected IsNot Nothing Then
                            Dim selectedValue As String = oCombo.Selected.Value.Trim()
                            If selectedValue = "Define New" Then
                                Dim AbsID As String = oForm.DataSources.DBDataSources.Item("OOAT").GetValue("AbsID", 0).Trim()
                                Dim objSalesBlanketAgreementMaster As New ClsSalesBlanketAgreementMaster
                                objSalesBlanketAgreementMaster.CreateForm(AbsID)
                            End If
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    objform = objMain.objApplication.Forms.Item(FormUID)

                    If pVal.ItemUID = "TypeDis" AndAlso pVal.BeforeAction = True Then
                        Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUID)
                        Dim oCombo As SAPbouiCOM.ComboBox = CType(oForm.Items.Item("TypeDis").Specific, SAPbouiCOM.ComboBox)
                        LoadTypeDisCombo_Union(oForm)
                    End If

                    If pVal.ItemUID = "U_PAYMENT" AndAlso pVal.BeforeAction = True Then
                        If objform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            objMain.objApplication.StatusBar.SetText("User Not Autorized to modify", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Select
                        End If
                    End If

                    If pVal.ItemUID = "1250000001" AndAlso pVal.Before_Action = True Then
                        objform = objMain.objApplication.Forms.Item(FormUID)
                        objMatrix = objform.Items.Item("1250000045").Specific
                        If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            For i As Integer = 1 To objMatrix.RowCount
                                If objMatrix.Columns.Item("1250000001").Cells.Item(i).Specific.Value <> "" Then
                                    strValQry = "SELECT * FROM OOAT T0 INNER JOIN OAT1 T1 ON T0.""AbsID"" = T1.""AgrNo"" WHERE T0.""BpCode"" = '" & objform.Items.Item("1250000006").Specific.Value & "' AND T1.""ItemCode"" = '" & objMatrix.Columns.Item("1250000001").Cells.Item(i).Specific.Value & "' AND T0.""StartDate"" >= '" & objform.Items.Item("1250000014").Specific.Value & "' AND T0.""EndDate"" <= '" & objform.Items.Item("1250000016").Specific.Value & "'"
                                    oRecvQRY = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRecvQRY.DoQuery(strValQry)
                                    If oRecvQRY.RecordCount > 0 Then
                                        objMain.objApplication.StatusBar.SetText("Same Supplier, Same ItemCode And similar From_Date and To_Date already available", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                    End If
                                End If
                            Next
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                        Dim CFL_Id As String
                        CFL_Id = CFLEvent.ChooseFromListUID
                        oCFL = objform.ChooseFromLists.Item(CFL_Id)
                        Dim oDT As SAPbouiCOM.DataTable
                        oDT = CFLEvent.SelectedObjects
                        objform = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        If Not (oDT Is Nothing) AndAlso pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE AndAlso pVal.BeforeAction = False Then
                            If oCFL.UniqueID = "CFL_ACCT" Then
                                If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                Dim AccountCode As String = oDT.GetValue("AcctCode", 0)
                                Dim AccountName As String = oDT.GetValue("AcctName", 0)
                                objform = objMain.objApplication.Forms.GetForm("1250000100", objMain.objApplication.Forms.ActiveForm.TypeCount)
                                Try
                                    objform.Items.Item("edtGlCd").Specific.Value = AccountCode
                                Catch ex As Exception
                                End Try
                                Try
                                    objform.Items.Item("edtGlNm").Specific.Value = AccountName
                                Catch ex As Exception
                                End Try
                                If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                End If
                                If objform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                                    objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                End If
                            End If

                            '-- Promotion Item Details: populate Item Code and Item Name (U_ItemCode, U_ItemName)
                            If oCFL.UniqueID = "CFL_PROMO" Then
                                Try
                                    Dim sItemCode As String = ""
                                    Dim sItemName As String = ""
                                    Try
                                        sItemCode = oDT.GetValue("ItemCode", 0).ToString().Trim()
                                    Catch
                                    End Try
                                    Try
                                        sItemName = oDT.GetValue("ItemName", 0).ToString().Trim()
                                    Catch
                                    End Try

                                    If Not String.IsNullOrEmpty(sItemCode) Then
                                        Dim oPromoMatrix As SAPbouiCOM.Matrix = objform.Items.Item("mtxPromo").Specific
                                        Dim iRow As Integer = pVal.Row

                                        If iRow >= 1 AndAlso iRow <= oPromoMatrix.RowCount Then
                                            objform.Freeze(True)

                                            'Set Item Code
                                            Try
                                                oPromoMatrix.Columns.Item("U_ItemCode").Cells.Item(iRow).Specific.Value = sItemCode
                                            Catch
                                            End Try

                                            'Set Item Name (auto-populated)
                                            Try
                                                oPromoMatrix.Columns.Item("U_ItemName").Cells.Item(iRow).Specific.Value = sItemName
                                            Catch
                                            End Try

                                            'Row addition is handled by et_LOST_FOCUS to avoid duplicate rows
                                            'Renumber all rows
                                            RenumberMatrixRows(oPromoMatrix)

                                            If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            End If

                                            objform.Freeze(False)
                                        End If
                                    End If
                                Catch ex As Exception
                                    objform.Freeze(False)
                                    objMain.objApplication.StatusBar.SetText(
                                        "Promotion CFL Error: " & ex.Message,
                                        SAPbouiCOM.BoMessageTime.bmt_Short,
                                        SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End Try
                            End If
                        End If

                    Catch ex As Exception
                    End Try

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objform = objMain.objApplication.Forms.Item(FormUID)

                    ' Promotion Matrix Item Code lost focus (manual entry fallback)
                    If pVal.ItemUID = "mtxPromo" AndAlso pVal.ColUID = "U_ItemCode" AndAlso pVal.BeforeAction = False Then
                        Try
                            Dim oPromoMatrix As SAPbouiCOM.Matrix = objform.Items.Item("mtxPromo").Specific
                            Dim iRow As Integer = pVal.Row
                            If iRow >= 1 AndAlso iRow <= oPromoMatrix.RowCount Then
                                Dim sCode As String = ""
                                Try
                                    sCode = oPromoMatrix.Columns.Item("U_ItemCode").Cells.Item(iRow).Specific.Value.ToString().Trim()
                                Catch
                                End Try

                                If Not String.IsNullOrEmpty(sCode) Then
                                    Dim sName As String = ""
                                    Try
                                        sName = oPromoMatrix.Columns.Item("U_ItemName").Cells.Item(iRow).Specific.Value.ToString().Trim()
                                    Catch
                                    End Try

                                    If String.IsNullOrEmpty(sName) Then
                                        Dim sQry As String = "SELECT ""ItemName"" FROM OITM WHERE ""ItemCode"" = '" & sCode.Replace("'", "''") & "'"
                                        Dim oRsItem As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRsItem.DoQuery(sQry)
                                        If oRsItem.RecordCount > 0 Then
                                            sName = oRsItem.Fields.Item("ItemName").Value.ToString().Trim()
                                            objform.Freeze(True)
                                            Try
                                                oPromoMatrix.Columns.Item("U_ItemName").Cells.Item(iRow).Specific.Value = sName
                                            Catch
                                            End Try
                                            objform.Freeze(False)
                                        End If
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRsItem)
                                    End If

                                    If iRow = oPromoMatrix.RowCount Then
                                        objform.Freeze(True)
                                        oPromoMatrix.AddRow()
                                        Dim newRow As Integer = oPromoMatrix.RowCount
                                        Try
                                            oPromoMatrix.Columns.Item("U_ItemCode").Cells.Item(newRow).Specific.Value = ""
                                        Catch
                                        End Try
                                        Try
                                            oPromoMatrix.Columns.Item("U_ItemName").Cells.Item(newRow).Specific.Value = ""
                                        Catch
                                        End Try
                                        RenumberMatrixRows(oPromoMatrix)
                                        objform.Freeze(False)
                                    End If

                                    If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            objform.Freeze(False)
                        End Try
                    End If

                    objMatrix = objform.Items.Item("1250000045").Specific
                    If (pVal.ItemUID = "1250000006" OrElse pVal.ItemUID = "1250000014" OrElse pVal.ItemUID = "1250000016") AndAlso pVal.Before_Action = False AndAlso objform.Items.Item("1250000006").Specific.Value <> "" AndAlso objform.Items.Item("1250000014").Specific.Value <> "" AndAlso objform.Items.Item("1250000016").Specific.Value <> "" AndAlso objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        For i As Integer = 1 To objMatrix.RowCount
                            If objMatrix.Columns.Item("1250000001").Cells.Item(i).Specific.Value <> "" Then
                                strValQry = "SELECT * FROM OOAT T0 INNER JOIN OAT1 T1 ON T0.""AbsID"" = T1.""AgrNo"" WHERE T0.""BpCode"" = '" & objform.Items.Item("1250000006").Specific.Value & "' AND T1.""ItemCode"" = '" & objMatrix.Columns.Item("1250000001").Cells.Item(i).Specific.Value & "' AND T0.""StartDate"" >= '" & objform.Items.Item("1250000014").Specific.Value & "' AND T0.""EndDate"" <= '" & objform.Items.Item("1250000016").Specific.Value & "'"
                                oRecvQRY = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecvQRY.DoQuery(strValQry)
                                If oRecvQRY.RecordCount > 0 Then
                                    objMain.objApplication.StatusBar.SetText("Same Supplier, Same ItemCode And similar From_Date and To_Date already available", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If
                            End If
                        Next
                    End If

                    If pVal.ItemUID = "1250000045" AndAlso pVal.ColUID = "1250000001" AndAlso pVal.Before_Action = False AndAlso objform.Items.Item("1250000006").Specific.Value <> "" AndAlso objform.Items.Item("1250000014").Specific.Value <> "" AndAlso objform.Items.Item("1250000016").Specific.Value <> "" AndAlso objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If objMatrix.Columns.Item("1250000001").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            strValQry = "SELECT * FROM OOAT T0 INNER JOIN OAT1 T1 ON T0.""AbsID"" = T1.""AgrNo"" WHERE T0.""BpCode"" = '" & objform.Items.Item("1250000006").Specific.Value & "' AND T1.""ItemCode"" = '" & objMatrix.Columns.Item("1250000001").Cells.Item(pVal.Row).Specific.Value & "' AND T0.""StartDate"" >= '" & objform.Items.Item("1250000014").Specific.Value & "' AND T0.""EndDate"" <= '" & objform.Items.Item("1250000016").Specific.Value & "'"
                            oRecvQRY = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecvQRY.DoQuery(strValQry)
                            If oRecvQRY.RecordCount > 0 Then
                                objMain.objApplication.StatusBar.SetText("Same Supplier, Same ItemCode And similar From_Date and To_Date already available", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objform = objMain.objApplication.Forms.Item(FormUID)

                    ' Add button pressed (BeforeAction) -> Validate only in ADD mode
                    If (pVal.ItemUID = "1250000001" OrElse pVal.ItemUID = "1") AndAlso pVal.BeforeAction = True Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            If Me.Validation(objform.UniqueID) = False Then
                                BubbleEvent = False
                            End If
                        End If
                    End If

                    ' Add / Update button pressed (AfterAction) -> Save promotion details
                    If (pVal.ItemUID = "1250000001" OrElse pVal.ItemUID = "1") AndAlso pVal.BeforeAction = False Then
                        Try
                            Dim sAbsID As String = ""
                            Try
                                oDBs_Head = objform.DataSources.DBDataSources.Item("OOAT")
                                If oDBs_Head IsNot Nothing Then
                                    sAbsID = oDBs_Head.GetValue("AbsID", 0).Trim()
                                End If
                            Catch ex As Exception
                            End Try

                            If String.IsNullOrEmpty(sAbsID) OrElse sAbsID = "0" Then
                                sAbsID = GetAbsID(Nothing, objform)
                            End If

                            If Not String.IsNullOrEmpty(sAbsID) AndAlso sAbsID <> "0" Then
                                Me.SavePromotionDetails(sAbsID, objform)
                                If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                    objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                    objMain.objApplication.StatusBar.SetText("Operation completed successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                End If
                            End If
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText("Update Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If

                    '-- Tab switching: Promotion Item Details tab
                    If pVal.BeforeAction = False Then
                        Try
                            If pVal.ItemUID = "fldPromo" Then
                                objform.PaneLevel = 99
                                oDBs_Head = objform.DataSources.DBDataSources.Item("OOAT")
                                Dim sAbsID As String = ""
                                If oDBs_Head IsNot Nothing Then
                                    sAbsID = oDBs_Head.GetValue("AbsID", 0).Trim()
                                End If

                                If Not String.IsNullOrEmpty(sAbsID) AndAlso sAbsID <> "0" Then
                                    If ItemExists(objform, "mtxPromo") Then
                                        Dim oPromoMatrix As SAPbouiCOM.Matrix = objform.Items.Item("mtxPromo").Specific
                                        If sAbsID <> LoadedAbsID OrElse oPromoMatrix.RowCount <= 1 Then
                                            Me.LoadPromotionDetails(sAbsID, objform)
                                        End If
                                    End If
                                Else
                                    If ItemExists(objform, "mtxPromo") Then
                                        Dim oPromoMatrix As SAPbouiCOM.Matrix = objform.Items.Item("mtxPromo").Specific
                                        If oPromoMatrix.RowCount = 0 Then
                                            Dim savedMode As SAPbouiCOM.BoFormMode = objform.Mode
                                            objform.Freeze(True)
                                            oPromoMatrix.AddRow()
                                            Try
                                                oPromoMatrix.Columns.Item("U_ItemCode").Cells.Item(1).Specific.Value = ""
                                            Catch
                                            End Try
                                            Try
                                                oPromoMatrix.Columns.Item("U_ItemName").Cells.Item(1).Specific.Value = ""
                                            Catch
                                            End Try
                                            RenumberMatrixRows(oPromoMatrix)
                                            ' Restore mode — adding empty row should not dirty the form
                                            If objform.Mode <> savedMode Then
                                                objform.Mode = savedMode
                                            End If
                                            objform.Freeze(False)
                                        End If
                                    End If
                                End If
                            ElseIf pVal.ItemUID = "1320000072" OrElse pVal.ItemUID = "1320000068" OrElse pVal.ItemUID = "1320000070" Then
                                Dim oFolder As SAPbouiCOM.Folder = objform.Items.Item(pVal.ItemUID).Specific
                                objform.PaneLevel = oFolder.Pane
                            End If
                        Catch ex As Exception
                        End Try
                    End If

                    If pVal.ItemUID = "Btn_PD" AndAlso pVal.BeforeAction = False Then
                        objMatrix = objform.Items.Item("1250000045").Specific
                        oDBs_Head = objform.DataSources.DBDataSources.Item("OOAT")
                        Dim CheckIfTCExist As String = "Select " & IsNull & "(""U_PAYMENT"",'') ""PaymentDetails"" From ""OOAT"" Where ""AbsID""='" & oDBs_Head.GetValue("AbsID", 0) & "' And IFNULL(""U_PAYMENT"",'') <> '' "
                        Dim oRsCheckIfTCExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsCheckIfTCExist.DoQuery(CheckIfTCExist)

                        Dim GetSuppCode As String = "Select ""ConnBP"" , (Select ""CardName"" From OCRD Where ""CardCode"" = A.""ConnBP"" ) ""BPName"" From OCRD A Where ""CardCode"" = '" & oDBs_Head.GetValue("BpCode", 0) & "'"
                        Dim oRsGetSuppCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetSuppCode.DoQuery(GetSuppCode)

                        If oRsGetSuppCode.Fields.Item(0).Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Provide Supplier Details in BP Master Record....!!!")
                            Exit Sub
                        End If
                        Dim AM As String = ""
                        Try
                            Dim AutoManual As SAPbouiCOM.ComboBox = objform.Items.Item("edtPTyp").Specific
                            AM = AutoManual.Value
                        Catch ex As Exception
                        End Try
                        Dim Dsc As String = objform.Items.Item("1250000018").Specific.Value
                        If Dsc = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Provide Description....!!!")
                            Exit Sub
                        End If

                        If oRsCheckIfTCExist.Fields.Item("PaymentDetails").Value <> "" Then
                            Dim BRANCH As String = ""

                            Try
                                Dim oItem As SAPbouiCOM.Item

                                Try
                                    oItem = objform.Items.Item("U_BranchID")
                                    BRANCH = oItem.Specific.Value
                                Catch
                                    Try
                                        oItem = objform.Items.Item("BranchID")
                                        BRANCH = oItem.Specific.Value
                                    Catch
                                        BRANCH = ""
                                    End Try
                                End Try

                            Catch ex As Exception
                                BRANCH = ""
                            End Try
                            objMain.objPaymentDetails.CreateForm(objform.UniqueID, "Y", oRsGetSuppCode.Fields.Item(0).Value, oRsGetSuppCode.Fields.Item(1).Value, oDBs_Head.GetValue("AbsID", 0), oDBs_Head.GetValue("Number", 0), objform.Items.Item("1250000014").Specific.Value, objform.Items.Item("1250000016").Specific.Value, GetDoubleValue(objMatrix.Columns.Item("1250000007").Cells.Item(1).Specific.value), GetDoubleValue(objMatrix.Columns.Item("1250000009").Cells.Item(1).Specific.value.ToString.Replace("QR", "")), oRsCheckIfTCExist.Fields.Item("PaymentDetails").Value, BRANCH, AM, objform.Items.Item("edtGlCd").Specific.Value, objform.Items.Item("edtGlNm").Specific.Value, Dsc)
                        Else
                            If objform.Items.Item("1250000014").Specific.Value = "" Then
                                objMain.objApplication.StatusBar.SetText("Please Provide Start Date....!!!")
                                Exit Sub
                            End If

                            If objform.Items.Item("1250000016").Specific.Value = "" Then
                                objMain.objApplication.StatusBar.SetText("Please Provide End Date....!!!")
                                Exit Sub
                            End If

                            If objform.Items.Item("edtGlCd").Specific.Value = "" Then
                                objMain.objApplication.StatusBar.SetText("Please Provide Account Code....!!!")
                                Exit Sub
                            End If

                            If objform.Items.Item("edtGlNm").Specific.Value = "" Then
                                objMain.objApplication.StatusBar.SetText("Please Provide Account Name....!!!")
                                Exit Sub
                            End If

                            If objform.Items.Item("1250000014").Specific.Value = objform.Items.Item("1250000016").Specific.Value Then
                                objMain.objApplication.StatusBar.SetText("Start Date And End Date Can Not Be Equal....!!!")
                                Exit Sub
                            End If

                            If GetDoubleValue(objMatrix.Columns.Item("1250000007").Cells.Item(1).Specific.value) = "0.0" Then
                                objMain.objApplication.StatusBar.SetText("Please Provide Qty in Details Tab....!!!")
                                Exit Sub
                            End If

                            If GetDoubleValue(objMatrix.Columns.Item("1250000009").Cells.Item(1).Specific.value) = "0.0" Then
                                objMain.objApplication.StatusBar.SetText("Please Provide Price in Details Tab....!!!")
                                Exit Sub
                            End If
                            Dim BRANCH As String = ""

                            Try
                                Dim oItem As SAPbouiCOM.Item

                                Try
                                    oItem = objform.Items.Item("U_BranchID")
                                    BRANCH = oItem.Specific.Value
                                Catch
                                    Try
                                        oItem = objform.Items.Item("BranchID")
                                        BRANCH = oItem.Specific.Value
                                    Catch
                                        BRANCH = ""
                                    End Try
                                End Try

                            Catch ex As Exception
                                BRANCH = ""
                            End Try
                            If BRANCH = "" Then
                                objMain.objApplication.StatusBar.SetText("Please Provide Branch....!!!")
                                Exit Sub
                            End If

                            If AM = "Auto" Then
                                Dim GetNumberOfMonths As String = "CALL ""SBO_GetRentalAggMonthBtwn""('" & objform.Items.Item("1250000014").Specific.Value & "' , '" & objform.Items.Item("1250000016").Specific.Value & "')"
                                Dim oRsGetNumberOfMonths As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetNumberOfMonths.DoQuery(GetNumberOfMonths)
                                If oRsGetNumberOfMonths.RecordCount > 0 Then
                                    If oRsGetNumberOfMonths.Fields.Item(0).Value <> GetDoubleValue(objMatrix.Columns.Item("1250000007").Cells.Item(1).Specific.value) Then
                                        objMain.objApplication.StatusBar.SetText("Planned Qty should be " & oRsGetNumberOfMonths.Fields.Item(0).Value & "...!!!")
                                        Exit Sub
                                    End If
                                End If
                            End If
                            Dim Approval As SAPbouiCOM.ComboBox = objform.Items.Item("1250000036").Specific
                            If Approval.Value <> "A" Then
                                objMain.objApplication.StatusBar.SetText("Approval is required....!!!")
                                Exit Sub
                            End If

                            Dim amount As Double = GetDoubleValue(objMatrix.Columns.Item("1250000007").Cells.Item(1).Specific.value)

                            Try
                                Dim oItem As SAPbouiCOM.Item

                                Try
                                    oItem = objform.Items.Item("U_BranchID")
                                    BRANCH = oItem.Specific.Value
                                Catch
                                    Try
                                        oItem = objform.Items.Item("BranchID")
                                        BRANCH = oItem.Specific.Value
                                    Catch
                                        BRANCH = ""
                                    End Try
                                End Try

                            Catch ex As Exception
                                BRANCH = ""
                            End Try

                            objMain.objPaymentDetails.CreateForm(objform.UniqueID, "N", oRsGetSuppCode.Fields.Item(0).Value, oRsGetSuppCode.Fields.Item(1).Value, oDBs_Head.GetValue("AbsID", 0), oDBs_Head.GetValue("Number", 0), objform.Items.Item("1250000014").Specific.Value, objform.Items.Item("1250000016").Specific.Value, amount, GetDoubleValue(objMatrix.Columns.Item("1250000009").Cells.Item(1).Specific.value.ToString.Replace("QR", "")), oRsCheckIfTCExist.Fields.Item("PaymentDetails").Value, BRANCH, AM, objform.Items.Item("edtGlCd").Specific.Value, objform.Items.Item("edtGlNm").Specific.Value, Dsc)
                        End If
                    End If

                    'If pVal.BeforeAction = False Then

                    '    'Standard Add / Update button
                    '    If pVal.ItemUID = "1250000001" Then

                    '        'Check only after Update
                    '        If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                    '            Me.SetStatusFieldEditable(objform)

                    '        End If

                    '    End If

                    'End If


            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub SetStatusFieldEditable(objForm As SAPbouiCOM.Form, Optional sAbsID As String = "")

        Try
            Dim oStatusItem As SAPbouiCOM.Item =
            objForm.Items.Item("1250000036")

            Dim sStatus As String = ""

            'Get latest status from database when AbsID is available
            If Not String.IsNullOrEmpty(sAbsID) AndAlso sAbsID <> "0" Then

                Dim oRecordset As SAPbobsCOM.Recordset =
                objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                Dim sQuery As String =
                "SELECT ""Status"" FROM ""OOAT"" WHERE ""AbsID"" = " & sAbsID

                oRecordset.DoQuery(sQuery)

                If Not oRecordset.EoF Then
                    sStatus = oRecordset.Fields.Item("Status").Value.ToString().Trim()
                End If

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset)

            Else
                'Fallback to current form value
                Dim oStatusCombo As SAPbouiCOM.ComboBox =
                CType(oStatusItem.Specific, SAPbouiCOM.ComboBox)

                If oStatusCombo.Selected IsNot Nothing Then
                    sStatus = oStatusCombo.Selected.Value.Trim()
                End If

            End If

            'A = Approved
            If sStatus = "A" Then
                oStatusItem.Enabled = False
            Else
                oStatusItem.Enabled = True
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "SetStatusFieldEditable Error: " & ex.Message)
        End Try

    End Sub

    Private Sub SetStatusFieldEditable(objForm As SAPbouiCOM.Form)

        Try
            Dim oStatusItem As SAPbouiCOM.Item = objForm.Items.Item("1250000036")

            Dim oStatusCombo As SAPbouiCOM.ComboBox =
            CType(oStatusItem.Specific, SAPbouiCOM.ComboBox)

            Dim sStatus As String = ""

            If oStatusCombo.Selected IsNot Nothing Then
                sStatus = oStatusCombo.Selected.Value.Trim()
            End If

            If sStatus = "A" Then
                oStatusItem.Enabled = False
            Else
                oStatusItem.Enabled = True
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "SetStatusFieldEditable Error: " & ex.Message)
        End Try

    End Sub

    Private Sub SetStatusFieldEditable1(objForm As SAPbouiCOM.Form)

        Try
            Dim oStatusItem As SAPbouiCOM.Item =
            objForm.Items.Item("1250000036")

            Dim oStatusCombo As SAPbouiCOM.ComboBox =
            CType(oStatusItem.Specific, SAPbouiCOM.ComboBox)

            Dim sStatus As String = ""

            If oStatusCombo.Selected IsNot Nothing Then
                sStatus = oStatusCombo.Selected.Value.Trim()
            End If

            If sStatus = "A" Then

                ' Move focus before disabling the Status field
                objForm.Items.Item("edtGlCd").Click()

                ' Disable Status field
                oStatusItem.Enabled = False

            Else
                oStatusItem.Enabled = True
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "SetStatusFieldEditable Error: " & ex.Message)
        End Try

    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False AndAlso BusinessObjectInfo.ActionSuccess = True Then
                Try
                    objform = objMain.objApplication.Forms.Item(BusinessObjectInfo.FormUID)
                Catch ex As Exception
                    objform = objMain.objApplication.Forms.GetForm("1250000100", 1)
                End Try

                Dim sAbsID As String = GetAbsID(BusinessObjectInfo, objform)

                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                        If Not String.IsNullOrEmpty(sAbsID) AndAlso sAbsID <> "0" Then
                            Me.LoadPromotionDetails(sAbsID, objform)
                        End If

                        'Check status after updating
                        Me.SetStatusFieldEditable(objform)

                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                        If Not String.IsNullOrEmpty(sAbsID) AndAlso sAbsID <> "0" Then
                            Me.SavePromotionDetails(sAbsID, objform)
                        End If
                        Me.PopulateRentalAccount(objform.UniqueID)

                        'Check status after updating
                        Me.SetStatusFieldEditable(objform)

                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                        If Not String.IsNullOrEmpty(sAbsID) AndAlso sAbsID <> "0" Then
                            Me.SavePromotionDetails(sAbsID, objform)
                        End If

                        'Check status after updating
                        Me.SetStatusFieldEditable1(objform)


                End Select
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("FormDataEvent Error: " & ex.Message)
        End Try
    End Sub

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        Try
            objform = objMain.objApplication.Forms.Item(eventInfo.FormUID)
            If eventInfo.BeforeAction = True Then
                If objform.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE AndAlso objform.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                    If eventInfo.ItemUID = "mtxPromo" Then
                        SelectedPromoRow = eventInfo.Row
                        Try
                            oMenuItem = objMain.objApplication.Menus.Item("1280") ' Data
                            oMenus = oMenuItem.SubMenus
                            If Not oMenus.Exists("Add Row") Then
                                oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                                oCreationPackage.UniqueID = "Add Row"
                                oCreationPackage.String = "Add Row"
                                oCreationPackage.Enabled = True
                                oMenus.AddEx(oCreationPackage)
                            End If
                            If Not oMenus.Exists("Delete Row") Then
                                oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                                oCreationPackage.UniqueID = "Delete Row"
                                oCreationPackage.String = "Delete Row"
                                oCreationPackage.Enabled = True
                                oMenus.AddEx(oCreationPackage)
                            End If
                        Catch ex As Exception
                        End Try
                    Else
                        Try
                            oMenuItem = objMain.objApplication.Menus.Item("1280")
                            oMenus = oMenuItem.SubMenus
                            If oMenus.Exists("Add Row") Then
                                objMain.objApplication.Menus.RemoveEx("Add Row")
                            End If
                            If oMenus.Exists("Delete Row") Then
                                objMain.objApplication.Menus.RemoveEx("Delete Row")
                            End If
                        Catch ex As Exception
                        End Try
                    End If
                End If
            Else
                ' AfterAction
                Try
                    oMenuItem = objMain.objApplication.Menus.Item("1280")
                    oMenus = oMenuItem.SubMenus
                    If oMenus.Exists("Add Row") Then
                        objMain.objApplication.Menus.RemoveEx("Add Row")
                    End If
                    If oMenus.Exists("Delete Row") Then
                        objMain.objApplication.Menus.RemoveEx("Delete Row")
                    End If
                Catch ex As Exception
                End Try
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("RightClick Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction = False Then
                objform = objMain.objApplication.Forms.ActiveForm
                If objform Is Nothing OrElse objform.TypeEx <> "1250000100" Then Exit Sub

                ' ---- ADD MODE (1282) ----
                If pVal.MenuUID = "1282" Then
                    Me.PopulateRentalAccount(objform.UniqueID)
                    Exit Sub
                End If

                If Not ItemExists(objform, "mtxPromo") Then Exit Sub
                Dim oPromoMatrix As SAPbouiCOM.Matrix = objform.Items.Item("mtxPromo").Specific

                ' ---- ADD ROW ----
                If pVal.MenuUID = "Add Row" Then
                    objform.Freeze(True)

                    ' Determine insert position: after right-clicked row, or at the end
                    Dim insertAfter As Integer = SelectedPromoRow
                    If insertAfter < 1 OrElse insertAfter > oPromoMatrix.RowCount Then
                        insertAfter = oPromoMatrix.RowCount
                    End If

                    ' Insert a row after the selected position
                    oPromoMatrix.AddRow(1, insertAfter)
                    Dim newRow As Integer = insertAfter + 1

                    ' Clear cells of the new row
                    Try
                        oPromoMatrix.Columns.Item("U_ItemCode").Cells.Item(newRow).Specific.Value = ""
                    Catch
                    End Try
                    Try
                        oPromoMatrix.Columns.Item("U_ItemName").Cells.Item(newRow).Specific.Value = ""
                    Catch
                    End Try

                    RenumberMatrixRows(oPromoMatrix)
                    SelectedPromoRow = 0

                    If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If

                    objform.Freeze(False)
                End If

                ' ---- DELETE ROW ----
                If pVal.MenuUID = "Delete Row" OrElse pVal.MenuUID = "1293" Then
                    objform.Freeze(True)

                    Dim rowDeleted As Boolean = False

                    ' Check visual row selection
                    For r As Integer = oPromoMatrix.VisualRowCount To 1 Step -1
                        If oPromoMatrix.IsRowSelected(r) Then
                            oPromoMatrix.DeleteRow(r)
                            rowDeleted = True
                        End If
                    Next

                    ' If no row was selected with matrix selector, delete the right-clicked row
                    If Not rowDeleted AndAlso SelectedPromoRow >= 1 AndAlso SelectedPromoRow <= oPromoMatrix.RowCount Then
                        oPromoMatrix.DeleteRow(SelectedPromoRow)
                        rowDeleted = True
                    End If

                    ' If all rows were deleted, keep at least one empty row
                    If oPromoMatrix.RowCount = 0 Then
                        oPromoMatrix.AddRow()
                        Try
                            oPromoMatrix.Columns.Item("U_ItemCode").Cells.Item(1).Specific.Value = ""
                        Catch
                        End Try
                        Try
                            oPromoMatrix.Columns.Item("U_ItemName").Cells.Item(1).Specific.Value = ""
                        Catch
                        End Try
                    End If

                    RenumberMatrixRows(oPromoMatrix)
                    SelectedPromoRow = 0

                    If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If

                    objform.Freeze(False)
                End If
            End If
        Catch ex As Exception
            If objform IsNot Nothing Then objform.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Menu Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Function GetDoubleValue(ByVal value As String) As Double
        Dim result As Double = 0
        If String.IsNullOrWhiteSpace(value) Then
            Return 0
        End If
        Dim cleaned As String = System.Text.RegularExpressions.Regex.Replace(value, "[^0-9\.\-]", "")
        If String.IsNullOrWhiteSpace(cleaned) OrElse cleaned = "-" OrElse cleaned = "." Then
            Return 0
        End If
        If Double.TryParse(cleaned, result) Then
            Return result
        End If
        Return 0
    End Function

    Private Function AddChooseFromListItms(ByVal FormUID As String, ByVal CFL_ID As String) As Boolean
        Dim bOk As Boolean = True
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)

            If Not CFLExists(objform, CFL_ID) Then
                Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
                Dim oCons As SAPbouiCOM.Conditions
                Dim oCon As SAPbouiCOM.Condition
                oCFLs = objform.ChooseFromLists
                Dim oCFL As SAPbouiCOM.ChooseFromList
                Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
                oCFLCreationParams = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
                oCFLCreationParams.MultiSelection = False
                oCFLCreationParams.ObjectType = "1"
                oCFLCreationParams.UniqueID = CFL_ID
                oCFL = oCFLs.Add(oCFLCreationParams)
                oCons = oCFL.GetConditions()
                oCon = oCons.Add()
                Dim oItem As SAPbouiCOM.Item = objform.Items.Item("edtGlCd")
                Dim oEdit As SAPbouiCOM.EditText = oItem.Specific
                oEdit.ChooseFromListUID = CFL_ID
                oEdit.ChooseFromListAlias = "AcctCode"
            End If

        Catch ex As Exception
            bOk = False
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
        Return bOk
    End Function

    Public Sub PopulateRentalAccount(ByVal FormUID As String)
        Dim oRS As SAPbobsCOM.Recordset = Nothing
        Try
            Dim targetForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUID)

            If Not ItemExists(targetForm, "edtGlCd") Then Exit Sub

            Dim sQuery As String = "SELECT T0.""U_RNTLACT"", " & IsNull & "(T1.""AcctName"", '') AS ""AcctName"" " &
                                   "FROM ""@ME_ACFG"" T0 " &
                                   "LEFT JOIN ""OACT"" T1 ON T1.""AcctCode"" = T0.""U_RNTLACT"" " &
                                   "WHERE " & IsNull & "(T0.""U_RNTLACT"", '') <> ''"

            oRS = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRS.DoQuery(sQuery)

            If oRS.RecordCount > 0 Then
                Dim sAcctCode As String = oRS.Fields.Item("U_RNTLACT").Value.ToString().Trim()
                Dim sAcctName As String = oRS.Fields.Item("AcctName").Value.ToString().Trim()

                If Not String.IsNullOrEmpty(sAcctCode) Then
                    Try
                        targetForm.Items.Item("edtGlCd").Specific.Value = sAcctCode
                    Catch ex As Exception
                        Try
                            Dim oDBs As SAPbouiCOM.DBDataSource = targetForm.DataSources.DBDataSources.Item("OOAT")
                            If oDBs IsNot Nothing Then
                                oDBs.SetValue("U_GLCODE", 0, sAcctCode)
                            End If
                        Catch
                        End Try
                    End Try

                    If ItemExists(targetForm, "edtGlNm") Then
                        Try
                            targetForm.Items.Item("edtGlNm").Specific.Value = sAcctName
                        Catch ex As Exception
                            Try
                                targetForm.Items.Item("edtGlNm").Enabled = True
                                targetForm.Items.Item("edtGlNm").Specific.Value = sAcctName
                                targetForm.Items.Item("edtGlNm").Enabled = False
                            Catch
                            End Try
                            Try
                                Dim oDBs As SAPbouiCOM.DBDataSource = targetForm.DataSources.DBDataSources.Item("OOAT")
                                If oDBs IsNot Nothing Then
                                    oDBs.SetValue("U_GLNAME", 0, sAcctName)
                                End If
                            Catch
                            End Try
                        End Try
                    End If
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Populate Rental Account Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            If oRS IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS)
                oRS = Nothing
            End If
        End Try
    End Sub

    Private Sub AutoPaymentValidation(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub AddItems(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)

            If Not ItemExists(objform, "Btn_PD") Then
                objMain.objUtilities.AddButtonH(objform.UniqueID, "Btn_PD", objform.Items.Item("1250000002").Top, objform.Items.Item("1250000002").Left + objform.Items.Item("1250000002").Width + 5, 150, "1250000002", "Payment &Details", 0, 0, objform.Items.Item("1250000002").Height)
                objform.Items.Item("Btn_PD").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                objform.Items.Item("Btn_PD").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

            If Not ItemExists(objform, "lblGlCd") Then
                objMain.objUtilities.AddLabel(objform.UniqueID, "lblGlCd", objform.Items.Item("1250000018").Top, objform.Items.Item("1250000018").Left + objform.Items.Item("1250000018").Width + 10, 90, "Rental Account", "1250000018")
            End If

            If Not ItemExists(objform, "edtGlCd") Then
                objMain.objUtilities.AddEditBox(objform.UniqueID, "edtGlCd", objform.Items.Item("1250000018").Top, objform.Items.Item("lblGlCd").Left + objform.Items.Item("lblGlCd").Width + 10, 140, "OOAT", "U_GLCODE", "lblGlCd")
                objform.Items.Item("edtGlCd").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            End If

            If Not ItemExists(objform, "lblGlNm") Then
                objMain.objUtilities.AddLabel(objform.UniqueID, "lblGlNm", objform.Items.Item("lblGlCd").Top + 20, objform.Items.Item("lblGlCd").Left, objform.Items.Item("lblGlCd").Width, "Account Name", "lblGlCd")
            End If

            If Not ItemExists(objform, "edtGlNm") Then
                objMain.objUtilities.AddEditBox(objform.UniqueID, "edtGlNm", objform.Items.Item("edtGlCd").Top + 20, objform.Items.Item("edtGlCd").Left, objform.Items.Item("edtGlCd").Width, "OOAT", "U_GLNAME", "edtGlCd")
                objform.Items.Item("edtGlNm").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "-1", SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objform.Items.Item("edtGlNm").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, "4", SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            End If

            If Not ItemExists(objform, "BranchID") Then
                objMain.objUtilities.AddComboBox(objform.UniqueID, "BranchID", objform.Items.Item("edtGlCd").Top, objform.Items.Item("edtGlCd").Left - 320, objform.Items.Item("edtGlCd").Width, objform.Items.Item("edtGlCd").Height, "OOAT", "U_BranchID")
            End If

            If Not ItemExists(objform, "BrNm") Then
                objMain.objUtilities.AddLabel(objform.UniqueID, "BrNm", objform.Items.Item("BranchID").Top, objform.Items.Item("BranchID").Left - 110, 100, "Branch Selection", "BranchID")
            End If

            If Not ItemExists(objform, "PMNTDTLS") Then
                objMain.objUtilities.AddComboBox(objform.UniqueID, "PMNTDTLS", objform.Items.Item("BranchID").Top + objform.Items.Item("BranchID").Height + 5, objform.Items.Item("BranchID").Left, objform.Items.Item("BranchID").Width, objform.Items.Item("BranchID").Height, "OOAT", "U_PMNTDTLS")
            End If

            If Not ItemExists(objform, "PayNm") Then
                objMain.objUtilities.AddLabel(objform.UniqueID, "PayNm", objform.Items.Item("PMNTDTLS").Top, objform.Items.Item("PMNTDTLS").Left - 100, objform.Items.Item("PMNTDTLS").Width, "Payment Terms", "PMNTDTLS")
            End If

            If Not ItemExists(objform, "TypeDis") Then
                objMain.objUtilities.AddComboBox(objform.UniqueID, "TypeDis", objform.Items.Item("PMNTDTLS").Top + objform.Items.Item("PMNTDTLS").Height + 5, objform.Items.Item("PMNTDTLS").Left, objform.Items.Item("PMNTDTLS").Width, objform.Items.Item("PMNTDTLS").Height, "OOAT", "U_TYPEDIS")
            End If

            If Not ItemExists(objform, "TypeDisLbl") Then
                objMain.objUtilities.AddLabel(objform.UniqueID, "TypeDisLbl", objform.Items.Item("TypeDis").Top - 3, objform.Items.Item("TypeDis").Left - 100, objform.Items.Item("TypeDis").Width, "Type of Display", "TypeDis")
            End If

            '--------------------------------------------------
            ' Promotion Item Details Tab (Folder)
            '--------------------------------------------------
            If Not ItemExists(objform, "fldPromo") Then
                Try
                    Dim oFolderItem As SAPbouiCOM.Item = objform.Items.Add("fldPromo", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
                    oFolderItem.Top = objform.Items.Item("1320000072").Top
                    oFolderItem.Left = objform.Items.Item("1320000072").Left + objform.Items.Item("1320000072").Width + 20
                    oFolderItem.Width = objform.Items.Item("1320000072").Width + 60
                    oFolderItem.Height = objform.Items.Item("1320000072").Height
                    oFolderItem.AffectsFormMode = False

                    Dim oFolder As SAPbouiCOM.Folder = oFolderItem.Specific
                    oFolder.Caption = "Promotion Item Details"
                    oFolder.GroupWith("1320000072")
                    oFolder.Pane = 99
                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText("Folder Error: " & ex.Message)
                End Try
            End If

            '-- Register ChooseFromList for Item Master (ObjectType 4)
            If Not CFLExists(objform, "CFL_PROMO") Then
                Try
                    Dim oCFLs As SAPbouiCOM.ChooseFromListCollection = objform.ChooseFromLists
                    Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
                    oCFLCreationParams = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
                    oCFLCreationParams.MultiSelection = False
                    oCFLCreationParams.ObjectType = "4"
                    oCFLCreationParams.UniqueID = "CFL_PROMO"
                    oCFLs.Add(oCFLCreationParams)
                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText("CFL Registration Error: " & ex.Message)
                End Try
            End If

            '-- Add UserDataSources for Promotion Matrix columns
            Try
                If Not UserDataSourceExists(objform, "ud_RowNo") Then
                    objform.DataSources.UserDataSources.Add("ud_RowNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
                End If
                If Not UserDataSourceExists(objform, "ud_ICode") Then
                    objform.DataSources.UserDataSources.Add("ud_ICode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)
                End If
                If Not UserDataSourceExists(objform, "ud_IName") Then
                    objform.DataSources.UserDataSources.Add("ud_IName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200)
                End If
            Catch ex As Exception
            End Try

            '-- Create Promotion Matrix
            If Not ItemExists(objform, "mtxPromo") Then
                Try
                    Dim oMatrixItem As SAPbouiCOM.Item = objform.Items.Add("mtxPromo", SAPbouiCOM.BoFormItemTypes.it_MATRIX)
                    oMatrixItem.Top = objform.Items.Item("1250000045").Top
                    oMatrixItem.Left = objform.Items.Item("1250000045").Left
                    oMatrixItem.Width = objform.Items.Item("1250000045").Width
                    oMatrixItem.Height = objform.Items.Item("1250000045").Height
                    oMatrixItem.FromPane = 99
                    oMatrixItem.ToPane = 99

                    Dim oMatrix As SAPbouiCOM.Matrix = oMatrixItem.Specific
                    Dim oColumn As SAPbouiCOM.Column

                    '-- Row # column
                    oColumn = oMatrix.Columns.Add("colRowNo", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                    oColumn.TitleObject.Caption = "#"
                    oColumn.Width = 30
                    oColumn.Editable = False
                    oColumn.DataBind.SetBound(True, "", "ud_RowNo")

                    '-- Item Code column (with ChooseFromList)
                    oColumn = oMatrix.Columns.Add("U_ItemCode", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                    oColumn.TitleObject.Caption = "Item Code"
                    oColumn.Width = 150
                    oColumn.Editable = True
                    oColumn.DataBind.SetBound(True, "", "ud_ICode")
                    oColumn.ChooseFromListUID = "CFL_PROMO"
                    oColumn.ChooseFromListAlias = "ItemCode"

                    '-- Item Name column (read-only)
                    oColumn = oMatrix.Columns.Add("U_ItemName", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                    oColumn.TitleObject.Caption = "Item Name"
                    oColumn.Width = 250
                    oColumn.Editable = False
                    oColumn.DataBind.SetBound(True, "", "ud_IName")

                    '-- Add initial empty row
                    oMatrix.AddRow()
                    Try
                        oMatrix.Columns.Item("U_ItemCode").Cells.Item(1).Specific.Value = ""
                    Catch
                    End Try
                    Try
                        oMatrix.Columns.Item("U_ItemName").Cells.Item(1).Specific.Value = ""
                    Catch
                    End Try
                    RenumberMatrixRows(oMatrix)

                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText("Matrix Error: " & ex.Message)
                End Try
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CFLFilter(ByVal CFL_ID As String)
        Try
            If CFLExists(objform, CFL_ID) Then
                Dim oConditions As SAPbouiCOM.Conditions
                Dim oCondition As SAPbouiCOM.Condition
                Dim oChooseFromList As SAPbouiCOM.ChooseFromList
                Dim emptyCon As New SAPbouiCOM.Conditions
                oChooseFromList = objform.ChooseFromLists.Item(CFL_ID)
                oChooseFromList.SetConditions(emptyCon)
                oConditions = oChooseFromList.GetConditions()
                oCondition = oConditions.Add()
                oCondition.Alias = "Postable"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = "Y"
                oChooseFromList.SetConditions(oConditions)
                oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                oCondition = oConditions.Add()
                oCondition.Alias = "U_RENTLAG"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = "Y"
                oChooseFromList.SetConditions(oConditions)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Function Validation(ByVal FormUID As String) As Boolean
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)

            If objform.Items.Item("edtGlCd").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Account Code Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            If objform.Items.Item("BranchID").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Branch Cannot be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            If objform.Items.Item("PMNTDTLS").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Payment Details Number should be empty...!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
            Return False
        End Try
    End Function

    Private Sub LoadBranchCombo(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)
            objform.Freeze(True)

            Dim oCombo As SAPbouiCOM.ComboBox
            oCombo = objform.Items.Item("BranchID").Specific

            'Clear existing values
            For i As Integer = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
            Next

            Dim oRs As SAPbobsCOM.Recordset =
            objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'Query based on logged-in user
            Dim sQuery As String =
            "SELECT T2.""BPLName"" FROM USR6 T0 " &
            "INNER JOIN OUSR T1 ON T0.""UserCode"" = T1.""USER_CODE"" " &
            "INNER JOIN OBPL T2 ON T0.""BPLId"" = T2.""BPLId"" " &
            "WHERE T1.""USER_CODE"" = '" &
            objMain.objCompany.UserName.ToString().Trim() & "'"

            oRs.DoQuery(sQuery)

            If oRs.RecordCount > 0 Then
                While Not oRs.EoF
                    Dim sBranchName As String =
                    oRs.Fields.Item("BPLName").Value.ToString().Trim()

                    'Store and Show Branch Name only
                    oCombo.ValidValues.Add(sBranchName, sBranchName)

                    oRs.MoveNext()
                End While

                'Select first branch
                oCombo.Select(oCombo.ValidValues.Item(0).Value,
                          SAPbouiCOM.BoSearchKey.psk_ByValue)
            End If

            objform.Freeze(False)

        Catch ex As Exception
            objform.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub LoadPaymentTerms(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)
            objform.Freeze(True)

            Dim oCombo As SAPbouiCOM.ComboBox
            oCombo = objform.Items.Item("PMNTDTLS").Specific

            'Clear existing values
            For i As Integer = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
            Next

            'Get Recordset
            Dim oRs As SAPbobsCOM.Recordset
            oRs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim sQuery As String =
            "SELECT ""PymntGroup"" FROM ""OCTG"" ORDER BY ""PymntGroup"""

            oRs.DoQuery(sQuery)

            'Load ONLY Description
            While Not oRs.EoF
                Dim sName As String = oRs.Fields.Item("PymntGroup").Value.ToString().Trim()

                'Value = Description
                oCombo.ValidValues.Add(sName, sName)

                oRs.MoveNext()
            End While

            'Select first value
            If oCombo.ValidValues.Count > 0 Then
                oCombo.Select(oCombo.ValidValues.Item(0).Value,
                          SAPbouiCOM.BoSearchKey.psk_ByValue)
            End If

            objform.Freeze(False)

        Catch ex As Exception
            objform.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub LoadTypeDisCombo_Union(ByVal oForm As SAPbouiCOM.Form)
        Dim oCombo As SAPbouiCOM.ComboBox = oForm.Items.Item("TypeDis").Specific
        Dim oRS As SAPbobsCOM.Recordset
        oRS = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim strQuery As String
        strQuery = "SELECT ""Code"", ""Name"" FROM ""@TNX_SBAMRM""UNION ALL SELECT 'Define New', 'Define New' FROM DUMMY"
        objMain.objUtilities.ComboBoxLoadValuesNEW(oCombo, strQuery)
    End Sub

End Class
