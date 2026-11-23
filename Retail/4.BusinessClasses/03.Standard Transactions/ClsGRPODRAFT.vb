Public Class ClsGRPODRAFT

#Region "Declaration"
    Public objForm As SAPbouiCOM.Form
    Dim objMatrix, objMatrix1 As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Details As SAPbouiCOM.DBDataSource
    Dim oitem, oItem1, oItem2, oItem3, oItem4 As SAPbouiCOM.Item
    Dim result As Integer
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim BaseUom As String
    Dim CalcualtionBtnFlag As Boolean = False
    Dim DirectGRPOFlag As Boolean = True
    Dim SplitCombo As SAPbouiCOM.ComboBox
    Dim SplitFlgCombo As SAPbouiCOM.ComboBox
    Dim DiscPrcnt As Double = 0
    Dim DiscUnitPrice As Double = 0
    Dim GlobalDocnum As String = 0
    Dim NST As String = ""
#End Region
    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)

        Try

            If pVal.MenuUID = "1283" AndAlso pVal.BeforeAction Then

                Dim Ret As Integer

                Ret = objMain.objApplication.MessageBox("Are you sure you want to remove this draft?", 2, "Yes", "No")

                If Ret = 2 Then
                    BubbleEvent = False
                    Exit Sub
                End If

            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub


End Class
