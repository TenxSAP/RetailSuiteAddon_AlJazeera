Imports System.Collections.Specialized
Imports System.Configuration
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports SAPbobsCOM

Public Class Cls_ProductEnquiry

#Region "Declaration"
    Public objForm As SAPbouiCOM.Form
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Public oDS, oDS1 As SAPbouiCOM.DBDataSource
    Public PrgBar As SAPbouiCOM.ProgressBar
    Dim LostFocusFlag As Boolean = True
#End Region


    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("ProductEnquiry1.xml", "PRODEN", ResourceType.Embeded)

            objForm = objMain.objApplication.Forms.GetForm("PRODEN", objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.Freeze(True)

            oDS = objForm.DataSources.DBDataSources.Item("@TNX_PRDENQUIRY")
            oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_PRDENQUIRY_C0")

            objMatrix = objForm.Items.Item("Mtx").Specific

            objForm.EnableMenu("774", True)
            objForm.EnableMenu("1292", True)

            ' DocNum settings
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            ' Date settings
            'objForm.Items.Item("ENQDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("ENQDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("ENQDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("ENQDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("DDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            Me.SetDefault(objForm.UniqueID)
            Me.LoadBranchCombo()
            objForm.Freeze(False)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "PRODEN" And pVal.BeforeAction = False Then

                Me.CreateForm()

            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then

                Me.SetDefault(objForm.UniqueID)

            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then

                Me.SetNewLine(objForm.UniqueID)

            ElseIf pVal.MenuUID = "774" And pVal.BeforeAction = False Then

                objMatrix = objForm.Items.Item("Mtx").Specific

                Dim row As Integer = objMatrix.VisualRowCount

                If objMatrix.VisualRowCount = 0 Then
                    Me.SetNewLine(objForm.UniqueID)
                ElseIf objMatrix.IsRowSelected(row) = True Then
                    objMatrix.DeleteRow(row)
                Else
                    For i As Integer = 1 To objMatrix.VisualRowCount - 1
                        If objMatrix.IsRowSelected(i) Then
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
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" AndAlso pVal.BeforeAction = False AndAlso pVal.ActionSuccess = True AndAlso pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Me.SetDefault(objForm.UniqueID)
                    End If

                    'Madhu we required we can use
                    'Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                    '    objForm = objMain.objApplication.Forms.Item(FormUID)
                    '    oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_PRDENQUIRY_C0")
                    '    objMatrix = objForm.Items.Item("Mtx").Specific

                    '    Dim oCFL As SAPbouiCOM.ChooseFromList
                    '    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    '    oCFL = objForm.ChooseFromLists.Item(CFLEvent.ChooseFromListUID)

                    '    Dim oDT As SAPbouiCOM.DataTable = CFLEvent.SelectedObjects

                    '    If (Not oDT Is Nothing) AndAlso pVal.BeforeAction = False Then

                    '        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    '            objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    '        End If

                    '        If oCFL.UniqueID = "CFL_0" Then

                    '            Dim row As Integer = pVal.Row - 1

                    '            oDS1.SetValue("U_ITMNM", row, oDT.GetValue("ItemName", 0))
                    '            objMatrix.SetLineData(pVal.Row)

                    '            ' ✅ Add new line ONLY if last row
                    '            If pVal.Row = objMatrix.VisualRowCount Then
                    '                Me.SetNewLine(objForm.UniqueID)
                    '            End If

                    '        End If

                    '    End If

            End Select

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oDS = objForm.DataSources.DBDataSources.Item("@TNX_PRDENQUIRY")
            oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_PRDENQUIRY_C0")

            objMatrix = objForm.Items.Item("Mtx").Specific


            oDS1.Clear()
            objMatrix.Clear()


            Dim Query As String = "SELECT IFNULL(MAX(""DocNum""),0) + 1 AS ""NextDocNum"" FROM ""@TNX_PRDENQUIRY"""

            Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rs.DoQuery(Query)

            If Not rs.EoF Then
                oDS.SetValue("DocNum", oDS.Offset, rs.Fields.Item("NextDocNum").Value)
            End If

            oDS.SetValue("U_ENQDATE", oDS.Offset, DateTime.Now.ToString("yyyyMMdd"))
            oDS.SetValue("U_DDate", oDS.Offset, DateTime.Now.ToString("yyyyMMdd"))


            Me.SetNewLine(objForm.UniqueID)

            objForm.Freeze(False)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_PRDENQUIRY_C0")
            objMatrix = objForm.Items.Item("Mtx").Specific

            objMatrix.FlushToDataSource()

            objMatrix.AddRow()

            Dim row As Integer = objMatrix.VisualRowCount

            oDS1.SetValue("LineId", oDS1.Offset, row)
            oDS1.SetValue("U_ITMNM", oDS1.Offset, "")
            oDS1.SetValue("U_Desc", oDS1.Offset, "")
            oDS1.SetValue("U_Qty", oDS1.Offset, "")
            oDS1.SetValue("U_BUYER", oDS1.Offset, "")
            oDS1.SetValue("U_Remarks", oDS1.Offset, "")
            oDS1.SetValue("U_ATTCH", oDS1.Offset, "")

            ' 🔥 Push datasource to matrix row
            objMatrix.SetLineData(row)

            objMatrix.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub LoadBranchCombo()
        Try
            Dim oCombo As SAPbouiCOM.ComboBox
            oCombo = CType(objForm.Items.Item("Branch").Specific, SAPbouiCOM.ComboBox)

            'oCombo.ValidValues.Clear()

            Dim oRs As SAPbobsCOM.Recordset =
            objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oRs.DoQuery("SELECT T2.""BPLId"",T2.""BPLName"" FROM USR6 T0  INNER JOIN OUSR T1 ON T0.""UserCode"" = T1.""USER_CODE"" " &
                    " INNER JOIN OBPL T2 ON T0.""BPLId"" = T2.""BPLId"" WHERE T1.""USER_CODE"" = '" & objMain.objCompany.UserName.ToString().Trim & "'")


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

End Class
