Imports System.Drawing.Drawing2D
Imports System.Data
Imports System.Data.OleDb
Imports SAPbobsCOM
Public Class ClsBarcodeFilter

#Region "Declaration"

    Public objForm As SAPbouiCOM.Form
    Dim objContentMx, objMatrix As SAPbouiCOM.Matrix

    Private WithEvents BackgroundWorker_CSVLoad As New System.ComponentModel.BackgroundWorker
    Private WithEvents BackgroundWorker_PriceUpdateService As New System.ComponentModel.BackgroundWorker

    Dim objComboBox As SAPbouiCOM.ComboBox

    Public oDBs_Head, oDBs_Details, oDBs_HeaderF, oDBs_Content, oDBs_Summary, oDBs_Attacht, oDS, oDS1 As SAPbouiCOM.DBDataSource

    Public Prgbar As SAPbouiCOM.ProgressBar
    Dim objutilities As Utilities
    Public MainFormUID As String = ""
    Public Branch As String = ""
    Private m_Row As Integer = 0
    Private m_ColUID As String = ""

#End Region

    Sub BarCodeFilter(ByVal DocEntry As String, ByVal Branch As String)
        Me.Branch = Branch
        Try
            objMain.objUtilities.LoadForm("BarCodeGrid.xml", "ME_FILTER", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_FILTER", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Select()
            Dim oGrid As SAPbouiCOM.Grid
            Dim oDT As SAPbouiCOM.DataTable
            oGrid = objForm.Items.Item("Item_0").Specific
            oDT = objForm.DataSources.DataTables.Item("DT_0")
            Dim sQuery As String = "SELECT T0.""BcdCode"" AS ""Bar Code"",T0.""BcdName"" AS ""Bar Code Name"",
T0.""UomEntry"",T3.""UomCode"",T1.""ItemCode"" AS ""Item No"",T1.""ItemName"",SUM(T2.""OnHand"") AS ""Available Stock""
FROM OBCD T0
INNER JOIN OITM T1 ON T0.""ItemCode"" = T1.""ItemCode""
INNER JOIN OUOM T3 ON T3.""UomEntry"" = T0.""UomEntry""
INNER JOIN OITW T2 ON T2.""ItemCode"" = T1.""ItemCode""
INNER JOIN OWHS T4 ON T4.""WhsCode"" = T2.""WhsCode""
WHERE T4.""BPLid"" = '" & Branch & "' GROUP BY T0.""BcdCode"",T0.""BcdName"",T0.""UomEntry"",T3.""UomCode"",T1.""ItemCode"",T1.""ItemName""
ORDER BY T0.""BcdCode"",T1.""ItemCode"""
            oGrid.DataTable.ExecuteQuery(sQuery)
            For i As Integer = 0 To oGrid.Columns.Count - 1
                oGrid.Columns.Item(i).Editable = False
            Next

            CType(oGrid.Columns.Item("Bar Code"), SAPbouiCOM.EditTextColumn).TitleObject.Sortable = True
            CType(oGrid.Columns.Item("Item No"), SAPbouiCOM.EditTextColumn).TitleObject.Sortable = True
            CType(oGrid.Columns.Item("Available Stock"), SAPbouiCOM.EditTextColumn).TitleObject.Sortable = True
            'Golden Arrow for Item Master
            Dim oCol As SAPbouiCOM.EditTextColumn
            oCol = CType(oGrid.Columns.Item("Item No"), SAPbouiCOM.EditTextColumn)
            oCol.LinkedObjectType = "4"
            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK



                    ''If pVal.FormTypeEx = "ME_FILTER" AndAlso pVal.ItemUID = "Item_0" AndAlso pVal.BeforeAction = False Then

                    ''    Try

                    ''        Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(pVal.FormUID)

                    ''        Dim oGrid As SAPbouiCOM.Grid = CType(oForm.Items.Item("Item_0").Specific, SAPbouiCOM.Grid)
                    ''        If pVal.Row < 0 Then Exit Sub
                    ''        Dim sBarCode As String = ""

                    ''        Dim val As Object = oGrid.DataTable.GetValue("Bar Code", pVal.Row)

                    ''        If val IsNot Nothing AndAlso val IsNot DBNull.Value Then
                    ''            sBarCode = val.ToString().Trim()
                    ''        End If

                    ''        Dim sItemCode As String = ""

                    ''        Dim val1 As Object = oGrid.DataTable.GetValue("Item No", pVal.Row)

                    ''        If val1 IsNot Nothing AndAlso val1 IsNot DBNull.Value Then
                    ''            sItemCode = val1.ToString().Trim()
                    ''        End If

                    ''        'Store row safely
                    ''        Dim selectedRow As Integer = pVal.Row

                    ''        'Try to get PCRT form BEFORE closing ME_FILTER
                    ''        Dim oPCRTForm As SAPbouiCOM.Form = Nothing

                    ''        Try
                    ''            'Better approach: loop and find PCRT form instead of ActiveForm
                    ''            For i As Integer = 0 To objMain.objApplication.Forms.Count - 1
                    ''                Dim f As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(i)
                    ''                If f.TypeEx = "PCRT" Then
                    ''                    oPCRTForm = f
                    ''                    Exit For
                    ''                End If
                    ''            Next
                    ''        Catch
                    ''            oPCRTForm = Nothing
                    ''        End Try

                    ''        'Close ME_FILTER
                    ''        'oForm.Freeze(True)

                    ''        'oForm.Close()

                    ''        'oForm.Freeze(False)

                    ''        'If PCRT found, call method
                    ''        If oPCRTForm IsNot Nothing Then
                    ''            objMain.ObjPriceChangeRequest.FillBarcodeData(oPCRTForm.UniqueID, selectedRow, sBarCode, sItemCode)

                    ''            For i As Integer = objMain.objApplication.Forms.Count - 1 To 0 Step -1

                    ''                Try
                    ''                    Dim oFrm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(i)

                    ''                    If oFrm.Title.Contains("Row Details") Then
                    ''                        oFrm.Close()
                    ''                    End If

                    ''                Catch
                    ''                End Try

                    ''            Next

                    ''            For i As Integer = objMain.objApplication.Forms.Count - 1 To 0 Step -1

                    ''                Try

                    ''                    Dim oFrm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(i)

                    ''                    If oFrm.TypeEx = "ME_FILTER" Then
                    ''                        oFrm.Close()
                    ''                    End If

                    ''                Catch
                    ''                End Try

                    ''            Next

                    ''        End If

                    ''    Catch ex As Exception

                    ''        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    ''    End Try

                    ''End If
                    '---Madhu 12/07/2026 barcode is issue is fixed
                    If pVal.FormTypeEx = "ME_FILTER" AndAlso pVal.ItemUID = "Item_0" AndAlso pVal.BeforeAction = False Then
                        Try
                            Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(pVal.FormUID)
                            Dim oGrid As SAPbouiCOM.Grid = CType(oForm.Items.Item("Item_0").Specific, SAPbouiCOM.Grid)

                            ' Ensure the row is valid
                            If pVal.Row < 0 Then Exit Sub

                            ' FIX: Get the actual data row index from the Grid to handle filtering correctly
                            Dim nActualRow As Integer = oGrid.GetDataTableRowIndex(pVal.Row)

                            Dim sBarCode As String = ""
                            Dim val As Object = oGrid.DataTable.GetValue("Bar Code", nActualRow)
                            If val IsNot Nothing AndAlso val IsNot DBNull.Value Then
                                sBarCode = val.ToString().Trim()
                            End If

                            Dim sItemCode As String = ""
                            Dim val1 As Object = oGrid.DataTable.GetValue("Item No", nActualRow)
                            If val1 IsNot Nothing AndAlso val1 IsNot DBNull.Value Then
                                sItemCode = val1.ToString().Trim()
                            End If

                            ' Store row safely
                            Dim selectedRow As Integer = objMain.ObjPriceChangeRequest.m_Row

                            ' Try to get PCRT form
                            Dim oPCRTForm As SAPbouiCOM.Form = Nothing
                            Try
                                For i As Integer = 0 To objMain.objApplication.Forms.Count - 1
                                    Dim f As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(i)
                                    If f.TypeEx = "PCRT" Then
                                        oPCRTForm = f
                                        Exit For
                                    End If
                                Next
                            Catch
                                oPCRTForm = Nothing
                            End Try

                            ' If PCRT found, call method
                            If oPCRTForm IsNot Nothing Then
                                objMain.ObjPriceChangeRequest.FillBarcodeData(oPCRTForm.UniqueID, selectedRow, sBarCode, sItemCode)

                                ' Close Row Details and ME_FILTER
                                For i As Integer = objMain.objApplication.Forms.Count - 1 To 0 Step -1
                                    Try
                                        Dim oFrm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(i)
                                        If oFrm.Title.Contains("Row Details") Then oFrm.Close()
                                    Catch
                                    End Try
                                Next

                                For i As Integer = objMain.objApplication.Forms.Count - 1 To 0 Step -1
                                    Try
                                        Dim oFrm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(i)
                                        If oFrm.TypeEx = "ME_FILTER" Then oFrm.Close()
                                    Catch
                                    End Try
                                Next
                            End If

                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If pVal.FormTypeEx = "ME_FILTER" AndAlso pVal.ItemUID = "Item_3" AndAlso pVal.BeforeAction = False Then

                        SearchBarcode(pVal.FormUID, Me.Branch)

                    End If

            End Select
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SearchBarcode(ByVal FormUID As String, ByVal Branch As String)

        Try

            Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUID)

            Dim oGrid As SAPbouiCOM.Grid = CType(oForm.Items.Item("Item_0").Specific, SAPbouiCOM.Grid)

            Dim oEdit As SAPbouiCOM.EditText = CType(oForm.Items.Item("Item_2").Specific, SAPbouiCOM.EditText)

            Dim sFind As String = oEdit.Value.Trim()

            Dim sQuery As String =
        "SELECT T0.""BcdCode"" AS ""Bar Code"",T0.""BcdName"" AS ""Bar Code Name""," &
        "T0.""UomEntry"",T3.""UomCode"",T1.""ItemCode"" AS ""Item No"",T1.""ItemName"",SUM(T2.""OnHand"") AS ""Available Stock"" " &
        "FROM OBCD T0 " &
        "INNER JOIN OITM T1 ON T0.""ItemCode"" = T1.""ItemCode"" " &
        "INNER JOIN OUOM T3 ON T3.""UomEntry"" = T0.""UomEntry"" " &
        "INNER JOIN OITW T2 ON T2.""ItemCode"" = T1.""ItemCode"" " &
        "INNER JOIN OWHS T4 ON T4.""WhsCode"" = T2.""WhsCode"" " &
        "WHERE T4.""BPLid"" = '" & Branch & "' "

            If sFind <> "" Then
                Dim searchText As String = sFind.Replace("'", "''")
                '  sQuery &= " AND (UPPER(T0.""BcdCode"") LIKE UPPER('%" & sFind.Replace("'", "''") & "%') " &
                '" OR UPPER(T1.""ItemCode"") LIKE UPPER('%" & sFind.Replace("'", "''") & "%')) "
                sQuery &= " AND (UPPER(T0.""BcdCode"") LIKE UPPER('%" & searchText & "%') " &
              " OR UPPER(T0.""BcdName"") LIKE UPPER('%" & searchText & "%') " &
              " OR UPPER(T1.""ItemCode"") LIKE UPPER('%" & searchText & "%') " &
              " OR UPPER(T1.""ItemName"") LIKE UPPER('%" & searchText & "%')) "
            End If

            sQuery &=
        "GROUP BY T0.""BcdCode"",T0.""BcdName"",T0.""UomEntry"",T3.""UomCode"",T1.""ItemCode"",T1.""ItemName"" " &
        "ORDER BY T0.""BcdCode"",T1.""ItemCode"""

            oGrid.DataTable.ExecuteQuery(sQuery)

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub



End Class
