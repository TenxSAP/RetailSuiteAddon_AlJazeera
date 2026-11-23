'NEW
Imports SAPbobsCOM

Public Class clsPromotiononBillTotal
#Region "       Declaration             "
    Public objForm, objSTRForm As SAPbouiCOM.Form
    'Dim oDS, oDS1 As SAPbouiCOM.DBDataSource
    Dim objMatrix, objMatrix1, objMatrix2, objMatrix3 As SAPbouiCOM.Matrix
    Dim objComboBox, objComboBox1 As SAPbouiCOM.ComboBox
    Dim oDt As SAPbouiCOM.DataTable
    Dim ItemGroup, subGroup, ItemWise, ItemSubGroup, sQury, sQury2, sQury3 As String
    Public oDS, oDS1, oDS2, oDS3, oDS4 As SAPbouiCOM.DBDataSource
    Public PrgBar As SAPbouiCOM.ProgressBar
    Dim LostFocusFlag As Boolean = True
    Dim MATRIXS As String
    Dim cdiscountpercentage, cofferprice As String

#End Region
    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("BillTotalNew.xml", "DISCOUNT1", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("DISCOUNT1", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            oDS = objForm.DataSources.DBDataSources.Item("@TNX_PBILL")
            objForm.EnableMenu("774", True)

            'Me.GetNextDocNum()
            'objForm.DataSources.UserDataSources.Add("U_TYPE", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            objForm.DataSources.UserDataSources.Add("U_ITGRP", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            objForm.DataSources.UserDataSources.Add("U_SBGRP", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            objForm.DataSources.UserDataSources.Add("U_SBSBGRP", SAPbouiCOM.BoDataType.dt_LONG_TEXT)


            objComboBox = Me.objForm.Items.Item("STATUS").Specific
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            'objForm.Items.Item("0_U_FD").Width = objForm.Items.Item("0_U_FD").Width + 140
            'objForm.Items.Item("Item_10").Left = objForm.Items.Item("Item_10").Left + 140
            ' objForm.Items.Item("Branch").Left = objForm.Items.Item("Branch").Left + 140
            objForm.Items.Item("0_U_FD").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            objComboBox = objForm.Items.Item("PRTYP").Specific
            'objComboBox.DataBind.SetBound(True, "", "U_TYPE")
            objForm.Items.Item("PRTYP").AffectsFormMode = False
            objComboBox.ValidValues.Add("All", "All")
            objComboBox.ValidValues.Add("ItemWise", "Item Wise")


            objComboBox = objForm.Items.Item("ITMGRP1").Specific
            'objComboBox.DataBind.SetBound(True, "", "U_ITGRP")
            objForm.Items.Item("ITMGRP1").AffectsFormMode = False
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, "Select ""ItmsGrpCod""  , ""ItmsGrpNam""  From OITB Order By  ""ItmsGrpNam""")
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("ITMGRP1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("ITMGRP1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("ITMGRP1").DisplayDesc = True

            Dim oComb As SAPbouiCOM.ComboBox = objForm.Items.Item("BPLID").Specific

            'objMain.objUtilities.ComboBoxLoadValues(oComb, "Select ""BPLId"",""BPLName"" FROM OBPL Where ""BPLId""<>'1'")
            objMain.objUtilities.ComboBoxLoadValues(oComb, "Select T0.""BPLId"" ,  T1.""BPLName"" From USR6 T0 Inner Join " & " OBPL T1 On T1.""BPLId"" = T0.""BPLId"" Where ""UserCode"" = '" & objMain.objCompany.UserName & "' Order By T0.""BPLId""   ")

            oComb.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            objForm.Items.Item("BPLID").DisplayDesc = True

            ''objComboBox = objForm.Items.Item("ITMGRP1").Specific
            ''objComboBox.DataBind.SetBound(True, "", "U_ITGRP")
            ''objForm.Items.Item("ITMGRP1").AffectsFormMode = False
            'Dim sQury As String = "Select ""ItmsGrpCod"" , ""ItmsGrpNam""   From OITB  "
            'objComboBox = objForm.Items.Item("ITMGRP1").Specific
            'objMain.objUtilities.ComboBoxLoadValues(objComboBox, sQury)
            'objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            'objForm.Items.Item("ITMGRP1").DisplayDesc = True
            'Dim ItemGroup As String
            ''objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            'objForm.Items.Item("ITMGRP1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("ITMGRP1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            'objForm.Items.Item("ITMGRP1").DisplayDesc = True

            Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Me.objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            Me.objForm.Items.Item("DOCDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Me.objForm.Items.Item("DOCDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Me.objForm.Items.Item("DOCDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Me.objForm.Items.Item("DOCDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            oDt = objForm.DataSources.DataTables.Add("dt1")
            oDt = objForm.DataSources.DataTables.Item("dt1")

            objMain.objApplication.StatusBar.SetText("Form loaded successfully.....continue.!!.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized

            '' objComboBox = objForm.Items.Item("ITMGRP1").Specific

            'objMain.objUtilities.ComboBoxLoadValues(objComboBox, sQury)
            'objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            ''objForm.Items.Item("ITMGRP1").DisplayDesc = True
            ''Dim ItemGroup As String
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Shared Sub CFL_Filter(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef oApplication As SAPbouiCOM.Application, ByRef oCompany As SAPbobsCOM.Company, ByVal Query As String, ByVal Code As String, ByVal ItemCode As String, ByVal Subgrp As String, ByVal Subsubgrp As String)
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
            Dim strCFLID As String = Nothing
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


        End Try


    End Sub


    Public Shared Sub CFL_Filter1(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef oApplication As SAPbouiCOM.Application, ByRef oCompany As SAPbobsCOM.Company, ByVal Query As String, ByVal Code As String, ByVal ItemCode As String, ByVal Subgrp As String, ByVal Subsubgrp As String)
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
            Dim strCFLID As String = Nothing
            oCLFEvent = CType((pVal), SAPbouiCOM.ChooseFromListEvent)
            strCFLID = oCLFEvent.ChooseFromListUID
            oCFL = oApplication.Forms.ActiveForm.ChooseFromLists.Item(strCFLID)
            oCFL.SetConditions(emptyConds)
            oCons = oCFL.GetConditions()
            Dim rr As Integer = oRec.RecordCount


            If rr = 0 Then
                oCon = oCons.Add()
                oCon.[Alias] = ItemCode
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = oRec.Fields.Item(ItemCode).Value.ToString()
            Else


                For i As Integer = 0 To oRec.RecordCount - 1
                    oCon = oCons.Add()
                    oCon.[Alias] = ItemCode
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = oRec.Fields.Item(ItemCode).Value.ToString()


                    If i <> oRec.RecordCount - 1 Then
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    End If


                    oRec.MoveNext()
                Next
            End If


            oCFL.SetConditions(oCons)
        Catch ex As Exception


        End Try


    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ME_PRMONTOTLA" And pVal.BeforeAction = False Then

                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)
                Me.BranchLoad(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                'Me.CreateForm()
                'Me.GetNextDocNum()
                Me.SetDefault(objForm.UniqueID)
                Me.BranchLoad(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                Me.SetNewLine(objForm.UniqueID, MATRIXS)
                Me.SetDefault(objForm.UniqueID)
                ' Me.BranchLoad(objForm.UniqueID)

            ElseIf pVal.MenuUID = "774" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("DISCOUNT1", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix = objForm.Items.Item("Mtx").Specific
                objMatrix3 = objForm.Items.Item("EMtx1").Specific
                Dim row As Integer = objMatrix.VisualRowCount
                If MATRIXS.Equals("Mtx") = True Then
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
                    For i As Integer = 1 To objMatrix.VisualRowCount
                        objMatrix.Columns.Item("LineId").Cells.Item(i).Specific.Value = i
                        objMatrix.FlushToDataSource()

                    Next
                ElseIf MATRIXS.Equals("EMtx1") = True Then
                    If objMatrix3.IsRowSelected(1) <> True And objMatrix3.VisualRowCount < 1 Then
                        objMatrix3.AddRow()
                        oDS4.SetValue("LineId", oDS1.Offset, objMatrix3.VisualRowCount)
                        objMatrix3.SetLineData(objMatrix3.VisualRowCount)
                    End If
                    If objMatrix3.IsRowSelected(row) = True Then
                        objMatrix3.DeleteRow(row)
                    Else
                        For i As Integer = 1 To objMatrix3.VisualRowCount - 1

                            If objMatrix3.IsRowSelected(i) = True Then
                                '  If i <> 1 Then
                                objMatrix3.DeleteRow(i)
                                '   End If
                            End If
                        Next
                    End If
                    For i As Integer = 1 To objMatrix3.VisualRowCount
                        objMatrix3.Columns.Item("LineId").Cells.Item(i).Specific.Value = i
                    Next
                End If
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "Mtx" And pVal.ColUID = "OFFP" And pVal.BeforeAction = True Then
                        Try
                            objForm.Freeze(True)
                            Dim RetPrice As Double = 0

                            Dim ObjMtx1 As SAPbouiCOM.Matrix = objForm.Items.Item(pVal.ItemUID).Specific
                            Dim ODbs_child As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Item("@TNX_PBILL_C0")

                            If CStr(ObjMtx1.Columns.Item("PRICE").Cells.Item(pVal.Row).Specific.Value).Trim <> "" Then
                                RetPrice = CDbl(CStr(ObjMtx1.Columns.Item("PRICE").Cells.Item(pVal.Row).Specific.Value).Trim)
                            End If

                            Dim OfferPrice As Double = 0

                            If CStr(ObjMtx1.Columns.Item("OFFP").Cells.Item(pVal.Row).Specific.Value).Trim <> "" Then
                                OfferPrice = CDbl(CStr(ObjMtx1.Columns.Item("OFFP").Cells.Item(pVal.Row).Specific.Value).Trim)
                            End If

                            If OfferPrice > RetPrice Then
                                objMain.objApplication.StatusBar.SetText("Offer Price should be less then or equal to retail price")
                                objForm.Freeze(False)
                                BubbleEvent = False
                                Exit Try
                            End If
                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If

                    If pVal.ItemUID = "Mtx" And (pVal.ColUID = "OFFP") And pVal.BeforeAction = False And pVal.ItemChanged = True Then
                        Try
                            objForm.Freeze(True)

                            Dim RetPrice As Double = 0

                            Dim ObjMtx1 As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx").Specific
                            Dim ODbs_child As SAPbouiCOM.DBDataSource = objForm.DataSources.DBDataSources.Item("@TNX_PBILL_C0")

                            If CStr(ObjMtx1.Columns.Item("PRICE").Cells.Item(pVal.Row).Specific.Value).Trim <> "" Then
                                RetPrice = CDbl(CStr(ObjMtx1.Columns.Item("PRICE").Cells.Item(pVal.Row).Specific.Value).Trim)
                            End If

                            Dim OfferPrice As Double = 0

                            If CStr(ObjMtx1.Columns.Item("OFFP").Cells.Item(pVal.Row).Specific.Value).Trim <> "" Then
                                OfferPrice = CDbl(CStr(ObjMtx1.Columns.Item("OFFP").Cells.Item(pVal.Row).Specific.Value).Trim)
                            End If

                            If OfferPrice > RetPrice Then
                                objMain.objApplication.StatusBar.SetText("Offer Price should be less then or equal to retail price")
                                objForm.Freeze(False)
                                BubbleEvent = False
                                Exit Try
                            Else
                                Dim DiscountAmount As Double = RetPrice - OfferPrice
                                Dim DiscPer As Double = (DiscountAmount / RetPrice) * 100
                                ODbs_child.SetValue("LineId", ODbs_child.Offset, pVal.Row)
                                ODbs_child.SetValue("Code", ODbs_child.Offset, ObjMtx1.Columns.Item("BCODE").Cells.Item(pVal.Row).Specific.Value)
                                ODbs_child.SetValue("U_ITMCODE", ODbs_child.Offset, ObjMtx1.Columns.Item("ITMCODE").Cells.Item(pVal.Row).Specific.Value)
                                ODbs_child.SetValue("U_ITMNM", ODbs_child.Offset, ObjMtx1.Columns.Item("ITMNM").Cells.Item(pVal.Row).Specific.Value)
                                ODbs_child.SetValue("U_UOM", ODbs_child.Offset, ObjMtx1.Columns.Item("UOM").Cells.Item(pVal.Row).Specific.Value)

                                ODbs_child.SetValue("U_PRICE", ODbs_child.Offset, ObjMtx1.Columns.Item("PRICE").Cells.Item(pVal.Row).Specific.Value)
                                ODbs_child.SetValue("U_OFFP", ODbs_child.Offset, ObjMtx1.Columns.Item("OFFP").Cells.Item(pVal.Row).Specific.Value)
                                ODbs_child.SetValue("U_REMK", ODbs_child.Offset, ObjMtx1.Columns.Item("REMK").Cells.Item(pVal.Row).Specific.Value)

                                ODbs_child.SetValue("U_DAMT", ODbs_child.Offset, DiscountAmount)
                                ODbs_child.SetValue("U_DPERCENT", ODbs_child.Offset, DiscPer)
                                ObjMtx1.SetLineData(pVal.Row)
                            End If
                            objForm.Freeze(False)
                        Catch ex As Exception
                            objForm.Freeze(False)
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If


                    '''lost focus -----------------------

                    'working before
                    '             Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    '                 objForm = objMain.objApplication.Forms.Item(FormUID)



                    '                 If pVal.ItemUID = "Mtx" AndAlso
                    '(pVal.ColUID = "DAMT" OrElse
                    ' pVal.ColUID = "DPERCENT" OrElse
                    ' pVal.ColUID = "REMK") AndAlso
                    'pVal.BeforeAction = False Then

                    '                     Dim ObjMtx As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx").Specific

                    '                     Dim row As Integer = pVal.Row

                    '                     ' Read values from the row

                    '                     Dim retailPrice As Double = Val(ObjMtx.Columns.Item("PRICE").Cells.Item(row).Specific.Value)

                    '                     Dim discountAmount As Double = Val(ObjMtx.Columns.Item("DAMT").Cells.Item(row).Specific.Value)

                    '                     Dim discountPercent As Double = Val(ObjMtx.Columns.Item("DPERCENT").Cells.Item(row).Specific.Value)

                    '                     If retailPrice > 0 Then

                    '                         ' If user edited DISCOUNT AMOUNT → calculate DISCOUNT %

                    '                         If pVal.ColUID = "DAMT" Then

                    '                             discountPercent = (discountAmount / retailPrice) * 100

                    '                             ObjMtx.Columns.Item("DPERCENT").Cells.Item(row).Specific.Value =
                    '             discountPercent.ToString("0.00")
                    '                         End If

                    '                         ' If user edited DISCOUNT % → calculate DISCOUNT AMOUNT

                    '                         If pVal.ColUID = "DPERCENT" Then

                    '                             discountAmount = (retailPrice * discountPercent) / 100

                    '                             ObjMtx.Columns.Item("DAMT").Cells.Item(row).Specific.Value =
                    '             discountAmount.ToString("0.00")
                    '                         End If

                    '                     End If

                    '                     ' Move cursor to REMK

                    '                     ObjMtx.Columns.Item("REMK").Cells.Item(row).Click(
                    '     SAPbouiCOM.BoCellClickType.ct_Regular)
                    '                 End If



                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "Mtx" AndAlso (pVal.ColUID = "DAMT") AndAlso pVal.BeforeAction = False Then

                        Dim ObjMtx1 As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx").Specific

                        Dim retailPrice As Double = Convert.ToDouble(ObjMtx1.Columns.Item("PRICE").Cells.Item(pVal.Row).Specific.Value)

                        Dim discountAmount As Double = Convert.ToDouble(ObjMtx1.Columns.Item("DAMT").Cells.Item(pVal.Row).Specific.Value)

                        Dim offerPrice As Double = retailPrice - discountAmount

                        Dim discountPercentage As Double = (discountAmount / retailPrice) * 100

                        ObjMtx1.Columns.Item("OFFP").Cells.Item(pVal.Row).Specific.Value = offerPrice.ToString("0.00")

                        ObjMtx1.Columns.Item("DPERCENT").Cells.Item(pVal.Row).Specific.Value = discountPercentage.ToString("0.00")

                        cofferprice = offerPrice.ToString("0.00")

                        cdiscountpercentage = discountPercentage.ToString("0.00")

                        ' Optional: remove this click if it causes continuous triggering

                        ' ObjMtx1.Columns.Item("REMK").Cells.Item(ObjMtx1.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                    End If

                    If pVal.ItemUID = "Mtx" AndAlso (pVal.ColUID = "DPERCENT") AndAlso pVal.BeforeAction = False Then

                        Dim discountper As String = objMatrix.Columns.Item("DPERCENT").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()

                        Dim offerpri As String = objMatrix.Columns.Item("OFFP").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()

                        If discountper <> cdiscountpercentage Then

                            Dim ObjMtx1 As SAPbouiCOM.Matrix = objForm.Items.Item("Mtx").Specific

                            Dim retailprice As Double = Convert.ToDouble(ObjMtx1.Columns.Item("PRICE").Cells.Item(pVal.Row).Specific.Value)

                            Dim discountpercentage As Double = Convert.ToDouble(ObjMtx1.Columns.Item("DPERCENT").Cells.Item(pVal.Row).Specific.Value)

                            ObjMtx1.Columns.Item("DAMT").Cells.Item(pVal.Row).Specific.Value = ((retailprice * discountpercentage) / 100).ToString("0.00")

                            Dim discountprice As Double = Convert.ToDouble(ObjMtx1.Columns.Item("DAMT").Cells.Item(pVal.Row).Specific.Value)

                            ObjMtx1.Columns.Item("OFFP").Cells.Item(pVal.Row).Specific.Value = (retailprice - discountprice).ToString("0.00")

                            cofferprice = ObjMtx1.Columns.Item("OFFP").Cells.Item(pVal.Row).Specific.Value

                            cdiscountpercentage = ObjMtx1.Columns.Item("DPERCENT").Cells.Item(pVal.Row).Specific.Value
                            ObjMtx1.Columns.Item("REMK").Cells.Item(ObjMtx1.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                        End If


                    End If







                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If (pVal.ItemUID = "Mtx" Or pVal.ItemUID = "IGMtx1" Or pVal.ItemUID = "BMtx" Or pVal.ItemUID = "EMtx1") And pVal.BeforeAction = False Then ' Or pVal.ItemUID = "Matrx5" Or pVal.ItemUID = "Matrix" Then
                        MATRIXS = pVal.ItemUID
                    End If
                    If (pVal.ItemUID = "Mtx" Or pVal.ItemUID = "IGMtx1" Or pVal.ItemUID = "BMtx" Or pVal.ItemUID = "EMtx1") And pVal.BeforeAction = True Then ' Or pVal.ItemUID = "Matrx5" Or pVal.ItemUID = "Matrix" Then
                        MATRIXS = pVal.ItemUID
                    End If
                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                    If pVal.ItemUID = "IGMtx1" And pVal.ColUID = "CHK" And pVal.BeforeAction = False Then
                        'If pVal.BeforeAction = False Then
                        objMatrix = Me.objForm.Items.Item("IGMtx1").Specific
                        Dim visualRowCount As Integer = Me.objMatrix.VisualRowCount
                        Dim checkBox As SAPbouiCOM.CheckBox = DirectCast(Me.objMatrix.Columns.Item("CHK").Cells.Item(1).Specific, SAPbouiCOM.CheckBox)
                        Dim Check As Boolean
                        If (Not checkBox.Checked) Then
                            Check = True
                        Else
                            Check = False
                        End If
                        For j As Integer = 1 To visualRowCount Step 1
                            checkBox = DirectCast(Me.objMatrix.Columns.Item("CHK").Cells.Item(j).Specific, SAPbouiCOM.CheckBox)
                            checkBox.Checked = Check
                        Next
                        'End If
                    End If
                    If pVal.ItemUID = "BMtx" And pVal.ColUID = "CHK" And pVal.BeforeAction = False Then
                        'If pVal.BeforeAction = False Then
                        objMatrix = Me.objForm.Items.Item("BMtx").Specific
                        Dim visualRowCount As Integer = Me.objMatrix.VisualRowCount
                        Dim checkBox As SAPbouiCOM.CheckBox = DirectCast(Me.objMatrix.Columns.Item("CHK").Cells.Item(1).Specific, SAPbouiCOM.CheckBox)
                        Dim Check As Boolean
                        If (Not checkBox.Checked) Then
                            Check = True
                        Else
                            Check = False
                        End If
                        For j As Integer = 1 To visualRowCount Step 1
                            checkBox = DirectCast(Me.objMatrix.Columns.Item("CHK").Cells.Item(j).Specific, SAPbouiCOM.CheckBox)
                            checkBox.Checked = Check
                        Next
                        'End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And
             (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        Me.SetDefault(objForm.UniqueID)
                        Me.BranchLoad(objForm.UniqueID)
                        'Me.GetNextDocNum()
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                                  (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                        Dim projtype As String = objForm.Items.Item("PRTYP").Specific.Value
                        If Me.Validate(objForm.UniqueID, projtype) = False Then
                            LostFocusFlag = False
                            BubbleEvent = False
                        End If

                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                                  (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                        Dim projtype As String = objForm.Items.Item("PRTYP").Specific.Value
                        If Me.Duplicatebarcode(objForm.UniqueID, projtype) = False Then
                            LostFocusFlag = False
                            BubbleEvent = False
                        End If
                    End If


                    'coment
                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                    '    If Me.validateDiscountType(objForm.UniqueID) = False Then
                    '        BubbleEvent = False
                    '    End If
                    'End If

                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                    '    If Me.ValidatePromotionName(objForm.UniqueID) = False Then
                    '        BubbleEvent = False
                    '    End If
                    'End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                        'If pVal.ItemUID = "1" AndAlso pVal.BeforeAction = True AndAlso (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Dim FromTime As String = CType(objForm.Items.Item("FRMTIME").Specific, SAPbouiCOM.EditText).Value
                        Dim ToTime As String = CType(objForm.Items.Item("TOTIME").Specific, SAPbouiCOM.EditText).Value
                        Dim fromDate As Date
                        Dim toDate As Date
                        If Not DateTime.TryParseExact(CType(objForm.Items.Item("FRMDATE").Specific, SAPbouiCOM.EditText).Value.Trim(), "yyyyMMdd", Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, fromDate) Then
                            objMain.objApplication.StatusBar.SetText("Please Choose 'From Date'. Please use format yyyyMMdd.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If Not DateTime.TryParseExact(CType(objForm.Items.Item("TODATE").Specific, SAPbouiCOM.EditText).Value.Trim(), "yyyyMMdd", Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, toDate) Then
                            objMain.objApplication.StatusBar.SetText("please Choose 'To Date'. Please use format yyyyMMdd.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If oDS.GetValue("U_FRMTIME", 0) = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter From Time", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If oDS.GetValue("U_TOTIME", 0) = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter To Time", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If oDS.GetValue("U_BPLID", 0) = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Select Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If


                        For i As Integer = 1 To objMatrix.VisualRowCount
                            objMatrix.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                            If objMatrix.Columns.Item("ITMCODE").Cells.Item(1).Specific.Value.ToString().Trim() = "" Then
                                'objMain.objApplication.StatusBar.SetText("Please Choose Different selection And press Load Button", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objMain.objApplication.StatusBar.SetText("Please Choose  Item Code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False

                            End If
                        Next

                    End If
                    '----
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.validateDiscountType(objForm.UniqueID) = False Then
                            BubbleEvent = False
                        End If
                    End If


                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.validateDiscountAmount(objForm.UniqueID) = False Then
                            BubbleEvent = False
                        End If
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.validateDiscountpercentage(objForm.UniqueID) = False Then
                            BubbleEvent = False
                        End If
                    End If




                    'Akhila Overlapping code

                    '                 If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                    '(pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                    '                     Dim FromTime As String = CType(objForm.Items.Item("FRMTIME").Specific, SAPbouiCOM.EditText).Value.Trim()
                    '                     Dim ToTime As String = CType(objForm.Items.Item("TOTIME").Specific, SAPbouiCOM.EditText).Value.Trim()

                    '                     Dim fromDate As String = objForm.Items.Item("FRMDATE").Specific.Value.Trim()
                    '                     Dim toDate As String = objForm.Items.Item("TODATE").Specific.Value.Trim()
                    '                     Dim docnum As String = objForm.Items.Item("DocNum").Specific.Value.Trim()
                    '                     Dim currentStatus As String = CType(objForm.Items.Item("STATUS").Specific, SAPbouiCOM.ComboBox).Selected.Value.Trim()

                    '                     Dim statustest As Boolean = False

                    '                     If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    '                         statustest = True
                    '                     ElseIf pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    '                         ' old status
                    '                         Dim oldStatusQuery As String = "SELECT ""U_STATUS"" FROM ""@TNX_PBILL"" WHERE ""DocNum"" = '" & docnum & "'"
                    '                         Dim rsOldStatus As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                         rsOldStatus.DoQuery(oldStatusQuery)
                    '                         If rsOldStatus.RecordCount > 0 Then
                    '                             Dim oldStatus As String = rsOldStatus.Fields.Item("U_STATUS").Value.ToString().Trim()
                    '                             '  if status is being changed from Inactive to Active
                    '                             If oldStatus = "Inactive" And currentStatus = "Active" Then
                    '                                 statustest = True
                    '                             End If
                    '                         End If
                    '                     End If

                    '                     If statustest Then
                    '                         Dim matrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx").Specific, SAPbouiCOM.Matrix)

                    '                         For i As Integer = 1 To matrix.VisualRowCount
                    '                             matrix.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    '                             Dim Barcode As String = CType(matrix.Columns.Item("BCODE").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value.Trim()
                    '                             Dim UOM As String = CType(matrix.Columns.Item("UOM").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value.Trim()

                    '                             Dim itemCode As String = CType(matrix.Columns.Item("ITMCODE").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value.Trim()
                    '                             If itemCode <> "" Then
                    '                                 Dim query As String = "SELECT H1.""DocNum"" FROM ""@TNX_PBILL"" H1 " &
                    '                 "INNER JOIN ""@TNX_PBILL_C0"" C1 ON H1.""DocEntry"" = C1.""DocEntry"" " &
                    '                 "WHERE C1.""U_ITMCODE"" = '" & itemCode & "' " &
                    '                 "AND H1.""U_FRMDATE"" <= '" & fromDate & "' " &
                    '                 "AND H1.""U_TODATE"" >= '" & toDate & "' " &
                    '                 "AND H1.""U_FRMTIME"" <= '" & FromTime & "' " &
                    '                 "AND H1.""U_TOTIME"" >= '" & ToTime & "' " &
                    '                 "AND H1.""U_STATUS"" = 'Active'" &
                    '                                     "AND C1.""U_UOM""='" & UOM & "'"

                    '                                 ' "AND C1.""Code""='" & Barcode & "' " &
                    '                                 Dim recordSet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                                 recordSet.DoQuery(query)
                    '                                 Dim doc As String = recordSet.Fields.Item("DocNum").Value

                    '                                 If recordSet.Fields.Item(0).Value > 0 Then
                    '                                     objMain.objApplication.StatusBar.SetText("DocNum " & doc & " - Row " & i.ToString() & ": Item " & itemCode & " is already in another active promotion with overlapping dates.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '                                     BubbleEvent = False
                    '                                     'Exit Sub
                    '                                 End If
                    '                             End If
                    '                         Next
                    '                     End If
                    '                 End If


                    'QUERY
                    '                 If pVal.ItemUID = "1" And pVal.BeforeAction = True And
                    '(pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                    '                     Dim matrix11 As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx").Specific, SAPbouiCOM.Matrix)

                    '                     For i As Integer = 1 To matrix11.VisualRowCount
                    '                         'matrix11.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                    '                         Dim itemCode1 As String = CType(matrix11.Columns.Item("ITMCODE").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value.Trim()
                    '                         If String.IsNullOrEmpty(itemCode1) Then Continue For

                    '                         Dim fromTime1 As String = CType(objForm.Items.Item("FRMTIME").Specific, SAPbouiCOM.EditText).Value.Trim()
                    '                         Dim toTime1 As String = CType(objForm.Items.Item("TOTIME").Specific, SAPbouiCOM.EditText).Value.Trim()
                    '                         Dim fromDate1 As String = objForm.Items.Item("FRMDATE").Specific.Value.Trim()
                    '                         Dim toDate1 As String = objForm.Items.Item("TODATE").Specific.Value.Trim()
                    '                         Dim currentDocNum As String = objForm.Items.Item("DocNum").Specific.Value.Trim()

                    '                         Dim query As String =
                    ' "WITH ItemExistence AS ( " &
                    ' "     SELECT 'Buy and get least Item' AS TableName,a.""DocNum"" as DocNum " &
                    ' "    FROM ""@TNX_BGLEASTITMS"" a " &
                    ' "    JOIN ""@TNX_BGLEASTITMS_C0"" b ON a.""DocEntry"" = b.""DocEntry"" " &
                    ' "    WHERE b.""U_ITMCODE"" = '" & itemCode1 & "' " &
                    ' "      AND a.""U_STATUS"" = 'ACTIVE' " &
                    ' "      AND a.""U_FRMDATE"" <= '" & fromDate1 & "' " &
                    ' "      AND a.""U_TODATE"" >= '" & toDate1 & "' " &
                    ' "      AND '" & toTime1 & "' <= a.""U_TOTIME"" " &
                    ' "      AND '" & fromTime1 & "' >= a.""U_FRMTIME"" " &
                    ' "    UNION " &
                    ' "     Select 'Item wise Discount' AS TableName,a.""DocNum"" as DocNum " &
                    ' "    FROM ""@TNX_ITEMDISC"" a " &
                    ' "    JOIN ""@TNX_ITEMDISC_C0"" b ON a.""DocEntry"" = b.""DocEntry"" " &
                    ' "    WHERE b.""U_ITMCODE"" = '" & itemCode1 & "' " &
                    ' "      AND a.""U_STATUS"" = 'Active' " &
                    ' "      AND a.""U_FRMDATE"" <= '" & fromDate1 & "' " &
                    ' "      AND a.""U_TODATE"" >= '" & toDate1 & "' " &
                    ' "      AND '" & toTime1 & "' <= a.""U_TOTIME"" " &
                    ' "      AND '" & fromTime1 & "' >= a.""U_FRMTIME"" " &
                    ' "    UNION " &
                    ' "    SELECT 'Promotion on voucher purchase' AS TableName ,a.""DocNum"" as DocNum " &
                    ' "    FROM ""@TNX_PONVP"" a " &
                    ' "    JOIN ""@TNX_PONVP_C0"" b ON a.""DocEntry"" = b.""DocEntry"" " &
                    ' "    WHERE b.""U_VITMCDE"" ='" & itemCode1 & "' " &
                    ' "      AND a.""U_STATUS"" = 'Active' " &
                    ' "      AND a.""U_FRMDATE"" <= '" & fromDate1 & "' " &
                    ' "      AND a.""U_TODATE"" >= '" & toDate1 & "' " &
                    ' "      AND'" & toTime1 & "' <= a.""U_TOTIME"" " &
                    ' "      AND '" & fromTime1 & "' >= a.""U_FRMTIME"" " &
                    ' ") " &
                    ' "SELECT TableName, DocNum FROM ItemExistence " &
                    ' "UNION ALL " &
                    ' "SELECT 'No' AS TableName,'0' As DocNum FROM DUMMY " &
                    ' "WHERE NOT EXISTS (SELECT 1 FROM ItemExistence);"

                    '                         Dim oRecordset As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                         oRecordset.DoQuery(query)
                    '                         If oRecordset.RecordCount > 0 Then
                    '                             Dim conflictTable As String = oRecordset.Fields.Item("TableName").Value.ToString()
                    '                             Dim conflictDocNum As String = oRecordset.Fields.Item("DocNum").Value.ToString()
                    '                             If conflictTable <> "No" Then
                    '                                 'objMain.objApplication.StatusBar.SetText("Item Code '" & itemCode1 & " is already linked to an active promotion with overlapping period in the screen " & conflictTable & " Duplicate promotion entries are not allowed.'",
                    '                                 '                                                     SAPbouiCOM.BoMessageTime.bmt_Short,
                    '                                 '                                                     SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    '                                 objMain.objApplication.StatusBar.SetText(
                    ' "Item Code '" & itemCode1 & "' is already linked to an active promotion with overlapping period in the screen " & conflictTable &
                    ' " (DocNum: " & conflictDocNum & "). Duplicate promotion entries are not allowed.",
                    ' SAPbouiCOM.BoMessageTime.bmt_Short,
                    ' SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    '                                 BubbleEvent = False
                    '                                 Exit Sub
                    '                             End If

                    '                         End If
                    '                     Next

                    '                 End If




                    'If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        If Me.ValidatePromotionName(objForm.UniqueID) = False Then
                            BubbleEvent = False
                        End If
                    End If



                    If pVal.ItemUID = "Upload" And pVal.BeforeAction = False Then
                        ' 🔽 Get file path from the edit text field
                        Dim filePath As String = CType(objForm.Items.Item("PATH").Specific, SAPbouiCOM.EditText).Value
                        If System.IO.File.Exists(filePath) Then
                            Me.LoadExcelToMatrix(objForm, filePath)
                        Else
                            objMain.objApplication.MessageBox("File not found or path is invalid.")
                        End If
                    End If

                    If pVal.ItemUID = "btnGO" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("SBCODE").Specific.Value <> "" Then
                            Dim bcode As String = objForm.Items.Item("SBCODE").Specific.Value
                            Dim Afterchk As String = " "
                            Dim TYPE As String = objForm.Items.Item("PRTYP").Specific.Value
                            Select Case TYPE
                                Case "ItemWise"
                                    For i As Integer = 1 To objMatrix.VisualRowCount
                                        If bcode = objMatrix.Columns.Item("BCODE").Cells.Item(i).Specific.Value Then
                                            objMatrix.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            ' objForm.Items.Item("SBCDE").Specific.Value = Afterchk
                                            Exit For
                                        End If
                                    Next
                                    'Case "ItemGroupWise"
                                    '    For i As Integer = 1 To objMatrix3.VisualRowCount
                                    '        If bcode = objMatrix3.Columns.Item("BCODE").Cells.Item(i).Specific.Value Then
                                    '            objMatrix3.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    '            ' objForm.Items.Item("SBCDE").Specific.Value = Afterchk
                                    '            Exit For
                                    '        End If
                                    '    Next
                            End Select
                        Else
                            objMain.objApplication.StatusBar.SetText("Please Enter Barcode", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "PRTYP" And pVal.BeforeAction = False Then
                        Select Case objForm.Items.Item("PRTYP").Specific.Selected.Value
                            Case "ItemWise"
                                objForm.Items.Item("ITMGRP1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("ITMGRP1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objComboBox = objForm.Items.Item("ITMGRP1").Specific
                                ' objComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

                                'objForm.Items.Item("SubGrp1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("SubGrp1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objComboBox = objForm.Items.Item("SubGrp1").Specific
                                ' objComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

                            'Case "ItemGroup"
                            '    objForm.Items.Item("ITMGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            '    objForm.Items.Item("SubGrp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            '    objForm.Items.Item("SubGrp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            '    objForm.Items.Item("SubSubGrp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            '    objForm.Items.Item("SubSubGrp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            Case "SubGroup"
                                objForm.Items.Item("ITMGRP1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("SubGrp1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("SubSubGrp1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("SubSubGrp1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            Case "SubSubGroup"
                                objForm.Items.Item("ITMGRP1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("SubGrp1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("SubSubGrp1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        End Select
                    End If

                    'new series add-------
                    If pVal.ItemUID = "19" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("19").Specific.Selected.Value <> "" Then
                            Dim series As Integer = objForm.Items.Item("19").Specific.Selected.Value
                            SetDefault(objForm.UniqueID, "CobmoSelect", series)
                        End If
                    End If



                    'end new series-------


                    If pVal.BeforeAction = False AndAlso pVal.ItemUID = "PRTYP" Then

                        Dim objCombo As SAPbouiCOM.ComboBox = objForm.Items.Item("PRTYP").Specific

                        Dim promoType As String = objCombo.Value.Trim()

                        ' Get combo fields

                        Dim cmbItemGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("ITMGRP1").Specific

                        Dim cmbSubGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("SubGrp1").Specific

                        Dim cmbSubSubGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("SubSubGrp1").Specific

                        If promoType = "All" Then

                            objForm.Items.Item("ITMGRP1").Enabled = False

                            objForm.Items.Item("SubGrp1").Enabled = False

                            objForm.Items.Item("SubSubGrp1").Enabled = False

                            cmbItemGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)

                            cmbSubGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)

                            cmbSubSubGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)

                        Else

                            objForm.Items.Item("ITMGRP1").Enabled = True

                            objForm.Items.Item("SubGrp1").Enabled = True

                            objForm.Items.Item("SubSubGrp1").Enabled = True

                            cmbItemGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)

                            cmbSubGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)

                            cmbSubSubGroup.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)

                        End If

                    End If





                    'If pVal.ItemUID = "ITMGRP1" Then
                    '    Dim cmb As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("ITMGRP1").Specific, SAPbouiCOM.ComboBox)
                    '    If cmb.Selected IsNot Nothing Then
                    '        Dim itemGroup As String = cmb.Selected.Value
                    '        LoadSubGroups1(objForm, itemGroup)
                    '        objForm.Items.Item("SubGrp1").Enabled = True
                    '    End If
                    'End If
                    'If pVal.ItemUID = "SubGrp1" Then
                    '    Dim cmbGrp As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("ITMGRP1").Specific, SAPbouiCOM.ComboBox)
                    '    Dim cmbSub As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("SubGrp1").Specific, SAPbouiCOM.ComboBox)
                    '    If cmbGrp.Selected IsNot Nothing AndAlso cmbSub.Selected IsNot Nothing Then
                    '        Dim itemGroup As String = cmbGrp.Selected.Value
                    '        Dim subGroup As String = cmbSub.Selected.Value
                    '        LoadSubSubGroups(objForm, itemGroup, subGroup)
                    '        objForm.Items.Item("SubSubGrp1").Enabled = True
                    '    End If
                    'End If



                    If objForm.Items.Item("ITMGRP1").Specific.Value <> "" Then
                        Dim objComboBoxItemGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("ITMGRP1").Specific
                        If objComboBoxItemGroup.Selected IsNot Nothing Then
                            Dim itemGroupCode As String = objComboBoxItemGroup.Selected.Value

                            Dim sql As String = "SELECT DISTINCT T0.""U_VSubGroup"", T0.""U_VSubGroup"" " &
                            "FROM OITM T0 " &
                            "WHERE T0.""ItmsGrpCod"" = '" & itemGroupCode & "' " &
                            "AND T0.""U_VSubGroup"" IS NOT NULL " &
                            "ORDER BY T0.""U_VSubGroup"""

                            Dim objComboBoxSubGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("SubGrp1").Specific

                            objMain.objUtilities.ComboBoxLoadValues(objComboBoxSubGroup, sql)
                            objComboBoxSubGroup.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                            objForm.Items.Item("SubGrp1").DisplayDesc = True
                        End If
                    End If





                    If objForm.Items.Item("SubGrp1").Specific.Value <> "" Then
                        Dim objComboBoxSubGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("SubGrp1").Specific

                        If objComboBoxSubGroup.Selected IsNot Nothing Then
                            Dim Subgroup As String = objComboBoxSubGroup.Selected.Value

                            ' Ensure ItemGroup is defined somewhere earlier, or retrieve it here again
                            Dim objComboBoxItemGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("ITMGRP1").Specific
                            Dim ItemGroup As String = ""
                            If objComboBoxItemGroup.Selected IsNot Nothing Then
                                ItemGroup = objComboBoxItemGroup.Selected.Value
                            Else
                                Exit Sub ' or handle as needed
                            End If


                            Dim sql As String = "Select DISTINCT
    ""U_VSubSubGroup"", 
    ""U_VSubSubGroup"" AS ""Description"" 
FROM OITM 
WHERE ""ItmsGrpCod"" = ' " & ItemGroup & "'
  And ""U_VSubSubGroup"" Is Not NULL 
  And ""U_VSubGroup"" = '" & Subgroup & "'  "

                            Dim objComboBoxSubSubGroup As SAPbouiCOM.ComboBox = objForm.Items.Item("SubSubGrp1").Specific

                            objMain.objUtilities.ComboBoxLoadValues(objComboBoxSubSubGroup, sql)
                            objComboBoxSubSubGroup.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                            objForm.Items.Item("SubSubGrp1").DisplayDesc = True
                        End If
                    End If







                    If pVal.ItemUID = "DISCTYP1" And pVal.BeforeAction = False Then
                        Dim oComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("DISCTYP1").Specific
                        Dim objMatrix1 As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx").Specific, SAPbouiCOM.Matrix)

                        If oComboBox.Selected.Value = "A" Then
                            objMatrix1.Columns.Item("DPERCENT").Editable = False
                            objMatrix1.Columns.Item("DAMT").Editable = True

                        ElseIf oComboBox.Selected.Value = "P" Then
                            objMatrix1.Columns.Item("DPERCENT").Editable = True
                            objMatrix1.Columns.Item("DAMT").Editable = False
                        Else
                            objMatrix1.Columns.Item("DPERCTG").Editable = False
                            objMatrix1.Columns.Item("DAMT").Editable = False
                        End If

                    End If
                    If pVal.ItemUID = "PRTYP" Then
                        'Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(pVal.FormUID)
                        Dim objComboBox As SAPbouiCOM.ComboBox = objForm.Items.Item("PRTYP").Specific

                        If objComboBox.Value = "ItemWise" Then
                            objForm.Items.Item("ITMGRP1").Enabled = True
                        End If
                    End If


                    If pVal.ItemUID = "ITMGRP1" Then
                        Dim cmb As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("ITMGRP1").Specific, SAPbouiCOM.ComboBox)
                        If cmb.Selected IsNot Nothing Then
                            Dim itemGroup As String = cmb.Selected.Value
                            LoadSubGroups1(objForm, itemGroup)
                        End If
                    End If

                    If pVal.ItemUID = "SubGrp1" Then
                        Dim cmbGrp As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("ITMGRP1").Specific, SAPbouiCOM.ComboBox)
                        Dim cmbSub As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("SubGrp1").Specific, SAPbouiCOM.ComboBox)
                        If cmbGrp.Selected IsNot Nothing AndAlso cmbSub.Selected IsNot Nothing Then
                            Dim itemGroup As String = cmbGrp.Selected.Value
                            Dim subGroup As String = cmbSub.Selected.Value
                            LoadSubSubGroups(objForm, itemGroup, subGroup)
                        End If
                    End If

                    If pVal.ItemUID = "STATUS" And pVal.BeforeAction = False Then
                        Dim status As String = objForm.Items.Item("STATUS").Specific.Value
                        If status = "Active" Then
                            objForm.Items.Item("FRMDATE").Enabled = True
                            objForm.Items.Item("TODATE").Enabled = True
                            objForm.Items.Item("MAXAP").Enabled = True
                            objForm.Items.Item("SBCODE").Enabled = True
                            objForm.Items.Item("FRMTIME").Enabled = True
                            objForm.Items.Item("TOTIME").Enabled = True
                            objForm.Items.Item("ITMGRP1").Enabled = True
                            objForm.Items.Item("PRTYP").Enabled = True
                            objForm.Items.Item("PNAME").Enabled = True
                            objForm.Items.Item("PATH").Enabled = True
                            objForm.Items.Item("Mtx").Enabled = True
                            objForm.Items.Item("IGMtx1").Enabled = True
                            objForm.Items.Item("EMtx1").Enabled = True
                            objForm.Items.Item("BMtx").Enabled = True
                        ElseIf status = "Inactive" Then
                            objForm.Items.Item("FRMDATE").Enabled = False
                            objForm.Items.Item("TODATE").Enabled = False
                            objForm.Items.Item("MAXAP").Enabled = False
                            objForm.Items.Item("SBCODE").Enabled = False
                            objForm.Items.Item("FRMTIME").Enabled = False
                            objForm.Items.Item("TOTIME").Enabled = False
                            objForm.Items.Item("ITMGRP1").Enabled = False
                            objForm.Items.Item("PRTYP").Enabled = False
                            objForm.Items.Item("PNAME").Enabled = False
                            objForm.Items.Item("PATH").Enabled = False
                            objForm.Items.Item("Mtx").Enabled = False
                            objForm.Items.Item("IGMtx1").Enabled = False
                            objForm.Items.Item("EMtx1").Enabled = False
                            objForm.Items.Item("BMtx").Enabled = False
                        End If
                    End If

                'Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                '    If pVal.ItemUID = "TODATE" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                '        If objForm.Items.Item("FRMDATE").Specific.Value > objForm.Items.Item("TODATE").Specific.Value Then
                '            objMain.objApplication.StatusBar.SetText("To Date must be after date to From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '            BubbleEvent = False
                '        End If
                '    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDS = objForm.DataSources.DBDataSources.Item("@TNX_PBILL")
                    oDS1 = objForm.DataSources.DBDataSources.Item("@TNX_PBILL_C0")
                    objMatrix = objForm.Items.Item("Mtx").Specific
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects




                    If (oCFL.UniqueID = "CFL_2" Or oCFL.UniqueID = "CFL_ITM") And pVal.BeforeAction = True Then
                        'oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)

                        ''Me.CFLFilter("CFL_OCRD")
                        ' objMatrix.AddRow()

                        Dim StrCFLQry As String = ""
                        If objForm.Items.Item("ITMGRP1").Specific.Value <> "" Then


                            StrCFLQry = "SELECT B.""ItemCode"",  B.""BcdCode"" ,I.""U_VSubGroup"",I.""U_VSubSubGroup""
                 FROM ""OBCD"" B " &
                 "INNER JOIN ""OITM"" I ON B.""ItemCode"" = I.""ItemCode"" " &
                 "WHERE   B.""ItemCode"" IS NOT NULL "


                            If CStr(objForm.Items.Item("ITMGRP1").Specific.Value).Trim <> "" Then
                                StrCFLQry = StrCFLQry & "  and  I.""ItmsGrpCod"" = '" & objForm.Items.Item("ITMGRP1").Specific.Value & "'"
                            End If

                            If CStr(objForm.Items.Item("SubGrp1").Specific.Value).Trim <> "" Then
                                StrCFLQry = StrCFLQry & "  And  I.""U_VSubGroup"" = '" & objForm.Items.Item("SubGrp1").Specific.Value & "'"
                            End If

                            If CStr(objForm.Items.Item("SubSubGrp1").Specific.Value).Trim <> "" Then
                                StrCFLQry = StrCFLQry & "  And  I.""U_VSubSubGroup"" = '" & objForm.Items.Item("SubSubGrp1").Specific.Value & "'"
                            End If



                        End If

                        If oCFL.UniqueID = "CFL_ITM" Then
                            CFL_Filter1(pVal, objMain.objApplication, objMain.objCompany, StrCFLQry, "BcdCode", "U_VSubSubGroup", "U_VSubSubGroup", "ItemCode")
                        End If


                        If oCFL.UniqueID = "CFL_2" Then

                            CFL_Filter(pVal, objMain.objApplication, objMain.objCompany, StrCFLQry, "BcdCode", "ItemCode", "U_VSubGroup", "U_VSubSubGroup")
                        End If
                    End If

                    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE



                        If oCFL.UniqueID = "CFL_ITM" Then

                            ' Prevent duplicate item in the same row
                            'Dim currentBarcode As String = objMatrix.Columns.Item("BCODE").Cells.Item(pVal.Row).Specific.Value

                            Dim Barcode As String = oDT.GetValue("CodeBars", 0)
                            Dim FromDate As String = oDS.GetValue("U_FRMDATE", 0)
                            Dim ToDate As String = oDS.GetValue("U_TODATE", 0)
                            Dim Brch As String = objForm.Items.Item("BPLID").Specific.Value

                            oDS1.SetValue("LineId", oDS1.Offset, pVal.Row)

                            ' Get Barcode from OBCD table
                            Dim GetBarCode As String = " SELECT Top 1 ""BcdCode"" FROM OBCD WHERE ""ItemCode"" = '" & oDT.GetValue("ItemCode", 0) & "' ORDER BY ""BcdEntry"" ASC "
                            Dim oRsGetBarCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetBarCode.DoQuery(GetBarCode)
                            Dim Barcode1 As String = oRsGetBarCode.Fields.Item("BcdCode").Value.ToString()

                            ' Set data to data source
                            oDS1.SetValue("Code", oDS1.Offset, Barcode1)
                            oDS1.SetValue("U_ITMCODE", oDS1.Offset, oDT.GetValue("ItemCode", 0))
                            oDS1.SetValue("U_ITMNM", oDS1.Offset, oDT.GetValue("ItemName", 0))
                            'oDS1.SetValue("U_UOM", oDS1.Offset, oDT.GetValue("InvntryUom", 0))


                            'ADDON-

                            Dim Itemcodenew As String = oDT.GetValue("ItemCode", 0)
                            Dim Itemuom As String = "SELECT DISTINCT A.""UomEntry"", " &
"(SELECT ""UomCode"" FROM ""OUOM"" WHERE ""UomEntry"" = A.""UomEntry"") AS ""UomName"" " &
"FROM ""OBCD"" A " &
"WHERE A.""ItemCode"" = '" & Itemcodenew & "'"

                            Dim itmcode As SAPbobsCOM.Recordset
                            itmcode = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            itmcode.DoQuery(Itemuom)
                            If itmcode.RecordCount > 0 Then
                                Dim itm As String = itmcode.Fields.Item("UomName").Value.ToString()
                                oDS1.SetValue("U_UOM", oDS1.Offset, itm)
                            End If







                            ' Set barcode to matrix
                            Barcode = objMatrix.Columns.Item("BCODE").Cells.Item(pVal.Row).Specific.Value
                            objMatrix.SetLineData(pVal.Row)

                            ' Get price using stored procedure
                            Dim query2 As String = "CALL ""GetRetailPrice""('" & Barcode1 & "','" & Brch.Trim & "')"
                            Dim orset2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            orset2.DoQuery(query2)

                            If orset2.RecordCount > 0 Then
                                Dim Retailprice As Double = orset2.Fields.Item("Price").Value
                                oDS1.SetValue("U_PRICE", oDS1.Offset, Retailprice)
                                objMatrix.SetLineData(pVal.Row)
                            End If

                            ' Add new line ONLY ONCE after all processing
                            If pVal.Row = objMatrix.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "Mtx")
                            End If

                            ' Add new line ONLY ONCE after all processing
                            If pVal.Row = objMatrix.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "Mtx")
                            End If
                        End If
                        If oCFL.UniqueID = "CFL_Ex" Then
                            oDS4.SetValue("LineId", oDS4.Offset, pVal.Row)
                            oDS4.SetValue("U_BCCODE", oDS4.Offset, oDT.GetValue("CodeBars", 0))
                            oDS4.SetValue("U_ITTMCODE", oDS4.Offset, oDT.GetValue("ItemCode", 0))
                            oDS4.SetValue("U_ITMNM", oDS4.Offset, oDT.GetValue("ItemName", 0))
                            oDS4.SetValue("U_UOM", oDS4.Offset, oDT.GetValue("InvntryUom", 0))
                            objMatrix3.SetLineData(pVal.Row)

                            If pVal.Row = objMatrix3.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "EMtx1")
                            End If
                            ''  Me.SetNewLine(objForm.UniqueID, "EMtx1")
                        End If

                        If oCFL.UniqueID = "CFL_3" And pVal.BeforeAction = False Then
                            Dim bcode As String = oDT.GetValue("BcdCode", 0)
                            Dim FromDate As String = oDS.GetValue("U_FRMDATE", 0)
                            Dim ToDate As String = oDS.GetValue("U_TODATE", 0)



                            Dim Itemcode As String = ""
                            oDS4.SetValue("U_BCCODE", oDS4.Offset, oDT.GetValue("BcdCode", 0))
                            Dim itmqury As String = "SELECT T1.""ItemCode"", T1.""ItemName"", T0.""BcdCode"", T2.""UomCode"" " &
                                                    "FROM ""OBCD"" T0 INNER JOIN ""OITM"" T1 ON T0.""ItemCode"" = T1.""ItemCode"" INNER JOIN OUOM T2 ON " &
                                                        " T0.""UomEntry"" = T2.""UomEntry"" Where T0.""BcdCode"" = '" & bcode & "'"
                            Dim oRitmqury As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRitmqury.DoQuery(itmqury)
                            oDS4.SetValue("LineId", oDS4.Offset, pVal.Row)

                            oDS4.SetValue("U_ITTMCODE", oDS4.Offset, oRitmqury.Fields.Item("ItemCode").Value)
                            oDS4.SetValue("U_ITMNM", oDS4.Offset, oRitmqury.Fields.Item("ItemName").Value)
                            oDS4.SetValue("U_UOM", oDS4.Offset, oRitmqury.Fields.Item("UomCode").Value)
                            oDS4.SetValue("U_BCCODE", oDS4.Offset, oRitmqury.Fields.Item("BcdCode").Value)
                            'oDS1.SetValue("U_FRMDATE", oDS1.Offset, oDS.GetValue("U_FRMDATE", 0))
                            'oDS1.SetValue("U_TODATE", oDS1.Offset, oDS.GetValue("U_TODATE", 0))

                            objMatrix3.SetLineData(pVal.Row)
                            'Me.SetNewLine(objForm.UniqueID, "EMtx1")


                            If pVal.Row = objMatrix3.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "EMtx1")
                            End If

                        End If
                    End If




                    If oCFL.UniqueID = "CFL_2" And pVal.BeforeAction = False Then

                        If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then

                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            End If

                            Dim Brch As String = objForm.Items.Item("BPLID").Specific.Value
                            Dim bcode As String = oDT.GetValue("BcdCode", 0).ToString().Trim()
                            Dim exists As Boolean = False

                            For i As Integer = 1 To objMatrix.VisualRowCount
                                Dim existing As String = objMatrix.Columns.Item("BCODE").Cells.Item(i).Specific.Value.ToString().Trim()
                                If existing = bcode And i <> pVal.Row Then
                                    exists = True
                                    Exit For
                                End If
                            Next
                            If exists = True Then
                                Exit Sub
                            End If

                            Dim itmqry As String =
"SELECT T1.""ItemCode"", T1.""ItemName"", T0.""BcdCode"", T2.""UomCode"" " &
"FROM ""OBCD"" T0 
  INNER JOIN ""OITM"" T1 ON T0.""ItemCode"" = T1.""ItemCode"" 
  INNER JOIN ""OUOM"" T2 ON T0.""UomEntry"" = T2.""UomEntry"" 
 WHERE T0.""BcdCode"" = '" & bcode & "'"

                            Dim rs As SAPbobsCOM.Recordset =
            objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            rs.DoQuery(itmqry)
                            If rs.RecordCount > 0 Then
                                oDS1.SetValue("LineId", oDS1.Offset, pVal.Row)
                                oDS1.SetValue("Code", oDS1.Offset, rs.Fields.Item("BcdCode").Value)
                                oDS1.SetValue("U_ITMCODE", oDS1.Offset, rs.Fields.Item("ItemCode").Value)
                                oDS1.SetValue("U_ITMNM", oDS1.Offset, rs.Fields.Item("ItemName").Value)
                                oDS1.SetValue("U_UOM", oDS1.Offset, rs.Fields.Item("UomCode").Value)
                                objMatrix.SetLineData(pVal.Row)
                            End If
                            Dim query2 As String = "CALL ""GetRetailPrice""('" & bcode & "', '" & Brch.Trim & "')"
                            Dim rsPrice As SAPbobsCOM.Recordset =
            objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            rsPrice.DoQuery(query2)
                            If rsPrice.RecordCount > 0 Then
                                Dim RetailPrice As Double = rsPrice.Fields.Item("Price").Value
                                oDS1.SetValue("U_PRICE", oDS1.Offset, RetailPrice)
                                objMatrix.SetLineData(pVal.Row)
                            End If
                            '----------------------------------------------------
                            '  CLEAR OFFER PRICE & DISCOUNT FOR ROW 
                            '----------------------------------------------------
                            If pVal.Row > 1 Then
                                oDS1.SetValue("U_OFFP", oDS1.Offset, "")
                                oDS1.SetValue("U_DAMT", oDS1.Offset, "")
                                oDS1.SetValue("U_DPERCENT", oDS1.Offset, "")
                                objMatrix.SetLineData(pVal.Row)
                            End If
                            '----------------------------------------------------
                            If pVal.Row = objMatrix.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "Mtx")
                            End If
                        End If
                    End If







            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("DISCOUNT1", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            If objForm.Items.Item("PRTYP").Specific.Value = "ItemWise" Then
                                objForm.Items.Item("0_U_FD").Enabled = True
                                objForm.Items.Item("IGroup").Enabled = False
                                objForm.Items.Item("ExItem").Enabled = False
                                objForm.Items.Item("Item_10").Enabled = True
                                objForm.Items.Item("0_U_FD").Visible = True
                                objForm.Items.Item("IGroup").Visible = False
                                objForm.Items.Item("ExItem").Visible = False
                                objForm.Items.Item("Item_10").Visible = True
                                objForm.Items.Item("0_U_FD").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            ElseIf objForm.Items.Item("PRTYP").Specific.Value = "ITMGRP1" Then
                                objForm.Items.Item("0_U_FD").Enabled = False
                                objForm.Items.Item("IGroup").Enabled = True
                                objForm.Items.Item("ExItem").Enabled = True
                                objForm.Items.Item("Item_10").Enabled = True
                                objForm.Items.Item("0_U_FD").Visible = False
                                objForm.Items.Item("IGroup").Visible = True
                                objForm.Items.Item("IGroup").Left = objForm.Items.Item("IGroup").Left - 100
                                objForm.Items.Item("ExItem").Visible = True
                                objForm.Items.Item("Item_10").Visible = True
                                objForm.Items.Item("IGroup").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                '    'Dim grp As String = "SELECT T0.""ItmsGrpCod"", T0.""ItmsGrpNam"" FROM OITB T0"
                                'Dim oRgrp As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                'oRgrp.DoQuery(grp)
                                'If oRgrp.RecordCount > 0 Then
                                '    objMatrix1.Clear()
                                '    For i As Integer = 0 To oRgrp.RecordCount - 1
                                '        objMatrix1.AddRow()
                                '        oDS2.SetValue("LineId", oDS2.Offset, objMatrix1.VisualRowCount)
                                '        oDS2.SetValue("U_ITMCDE", oDS2.Offset, oRgrp.Fields.Item("ItmsGrpCod").Value)
                                '        oDS2.SetValue("U_ITMGRP1", oDS2.Offset, oRgrp.Fields.Item("ItmsGrpNam").Value)
                                '        objMatrix1.SetLineData(objMatrix1.VisualRowCount)
                                '        oRgrp.MoveNext()
                                '    Next
                                'End If
                            End If
                        End If
                        Dim status As String = objForm.Items.Item("STATUS").Specific.Value
                        If status = "Active" Then
                            objForm.Items.Item("FRMDATE").Enabled = True
                            objForm.Items.Item("TODATE").Enabled = True
                            objForm.Items.Item("MAXAP").Enabled = True
                            objForm.Items.Item("SBCODE").Enabled = True
                            objForm.Items.Item("FRMTIME").Enabled = True
                            objForm.Items.Item("TOTIME").Enabled = True
                            objForm.Items.Item("ITMGRP1").Enabled = True
                            objForm.Items.Item("PRTYP").Enabled = True
                            objForm.Items.Item("PNAME").Enabled = True
                            objForm.Items.Item("PATH").Enabled = True
                            objForm.Items.Item("Mtx").Enabled = True
                            objForm.Items.Item("IGMtx1").Enabled = True
                            objForm.Items.Item("EMtx1").Enabled = True
                            objForm.Items.Item("BMtx").Enabled = True
                        ElseIf status = "Inactive" Then
                            objForm.Items.Item("FRMDATE").Enabled = False
                            objForm.Items.Item("TODATE").Enabled = False
                            objForm.Items.Item("MAXAP").Enabled = False
                            objForm.Items.Item("SBCODE").Enabled = False
                            objForm.Items.Item("FRMTIME").Enabled = False
                            objForm.Items.Item("TOTIME").Enabled = False
                            objForm.Items.Item("ITMGRP1").Enabled = False
                            objForm.Items.Item("PRTYP").Enabled = False
                            objForm.Items.Item("PNAME").Enabled = False
                            objForm.Items.Item("PATH").Enabled = False
                            objForm.Items.Item("Mtx").Enabled = False
                            objForm.Items.Item("IGMtx1").Enabled = False
                            objForm.Items.Item("EMtx1").Enabled = False
                            objForm.Items.Item("BMtx").Enabled = False
                        End If

                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Function GetFieldValue(ByVal query As String) As Object
        Try
            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery(query)
            Dim str As String = oRs.Fields.Item(0).Value.ToString()
            Return str
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "", Optional ByVal Series As Integer = 0)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            oDS = Me.objForm.DataSources.DBDataSources.Add("@TNX_PBILL")
            oDS1 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PBILL_C0")
            oDS2 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PBILL_C1")
            oDS3 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PBILL_C2")
            oDS4 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PBILL_C3")

            objMatrix = objForm.Items.Item("Mtx").Specific
            objMatrix2 = objForm.Items.Item("BMtx").Specific
            objMatrix1 = objForm.Items.Item("IGMtx1").Specific
            objMatrix3 = objForm.Items.Item("EMtx1").Specific

            oDS.SetValue("U_STATUS", oDS.Offset, "Active")



            ' Load series combo box
            Dim cmbSeries As SAPbouiCOM.ComboBox = objForm.Items.Item("19").Specific
            Dim seriesQuery As String = "SELECT ""Series"", ""SeriesName"" FROM NNM1 WHERE ""ObjectCode"" = 'ME_PBILLUDO' AND ""Locked"" = 'N'"
            objMain.objUtilities.ComboBoxLoadValuesSeries(cmbSeries, seriesQuery)
            objForm.Items.Item("19").DisplayDesc = True
            cmbSeries.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            ' Get default series
            '            Dim defaultSeriesQuery As String = "
            'SELECT T1.""Series"", T1.""SeriesName""
            'FROM ONNM T0 
            'JOIN NNM1 T1 ON T1.""ObjectCode"" = T0.""ObjectCode"" AND T1.""Series"" = T0.""DfltSeries"" 
            'WHERE T0.""ObjectCode"" IN ('ME_PBILLUDO','@TNX_PBILL')"




            Dim defaultSeriesQuery As String =
    "SELECT ""Series"", ""SeriesName"", ""ObjectCode"", ""Indicator"" " &
    "FROM NNM1 " &
    "WHERE ""ObjectCode"" IN ('ME_PBILLUDO', '@TNX_PBILL') " &
    "AND ""Indicator"" <> 'Default'"



            Dim oRsDflt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsDflt.DoQuery(defaultSeriesQuery)

            If Flag = "" Then
                Dim nextDocNum = objMain.objUtilities.GetNextDocNumSeries(objForm, "ME_PBILLUDO", oRsDflt.Fields.Item("Series").Value)
                oDS.SetValue("DocNum", oDS.Offset, nextDocNum)
                oDS.SetValue("U_SERIE", oDS.Offset, oRsDflt.Fields.Item("Series").Value)
            Else
                Dim nextDocNum = objMain.objUtilities.GetNextDocNumSeries(objForm, "ME_PBILLUDO", Series)
                oDS.SetValue("DocNum", oDS.Offset, nextDocNum)
                oDS.SetValue("U_SERIE", oDS.Offset, Series)
            End If

            oDS.SetValue("U_DOCDATE", oDS.Offset, DateTime.Now.ToString("yyyyMMdd"))

            'end series






            'oDS.SetValue("U_PRTYP", oDS.Offset, "ItemWise")
            'If objForm.Items.Item("PRTYP").Specific.Value = "ItemWise" Then
            '    'objForm.Items.Item("0_U_FD").Enabled = True
            '    'objForm.Items.Item("IGroup").Enabled = False
            '    'objForm.Items.Item("ExItem").Enabled = False
            '    'objForm.Items.Item("Item_10").Enabled = True
            '    objForm.Items.Item("0_U_FD").Visible = True
            '    objForm.Items.Item("IGroup").Visible = False
            '    objForm.Items.Item("ExItem").Visible = False
            '    objForm.Items.Item("Item_10").Visible = True
            '    objForm.Items.Item("0_U_FD").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            'End If
            objMatrix.Clear()
            oDS1.Clear()
            objMatrix.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "Mtx")

            objMatrix.AutoResizeColumns()
            objMatrix1.Clear()
            oDS2.Clear()
            objMatrix1.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "IGMtx1")
            objMatrix1.AutoResizeColumns()

            objMatrix2.Clear()
            oDS3.Clear()
            objMatrix2.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "BMtx")
            objMatrix2.AutoResizeColumns()

            objMatrix3.Clear()
            oDS4.Clear()
            objMatrix3.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "EMtx1")
            objMatrix3.AutoResizeColumns()
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Public Function GetNextDocNum()
        Dim Str As String
        Dim oRs As SAPbobsCOM.Recordset
        Dim DocNum As Integer
        oRs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Str = "SELECT IFNULL(MAX(""DocNum""),0)+1 as ""DocNum"" FROM ""@TNX_PBILL"""
        Try
            oRs.DoQuery(Str)
            oRs.MoveFirst()
            If oRs.RecordCount > 0 Then
                DocNum = oRs.Fields.Item(0).Value
                oDS.SetValue("DocNum", 0, oRs.Fields.Item(0).Value)
            End If
            If DocNum = 0 Then DocNum = 1
            Return DocNum
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Sub SetNewLine(ByVal FormUID As String, ByVal MatrixUID As String)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oDS = Me.objForm.DataSources.DBDataSources.Add("@TNX_PBILL")
        oDS1 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PBILL_C0")
        oDS2 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PBILL_C1")
        oDS3 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PBILL_C2")
        oDS4 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PBILL_C3")

        objMatrix = objForm.Items.Item("Mtx").Specific
        objMatrix1 = objForm.Items.Item("IGMtx1").Specific
        objMatrix2 = objForm.Items.Item("BMtx").Specific
        objMatrix3 = objForm.Items.Item("EMtx1").Specific

        Select Case MatrixUID
            Case "Mtx"
                objMatrix.AddRow()
                oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
                oDS1.SetValue("U_BCODE", oDS1.Offset, "")
                oDS1.SetValue("U_ITMCODE", oDS1.Offset, "")
                oDS1.SetValue("U_ITMNM", oDS1.Offset, "")
                oDS1.SetValue("U_UOM", oDS1.Offset, "")
                oDS1.SetValue("U_OFFP", oDS1.Offset, "")

                oDS1.SetValue("U_PRICE", oDS1.Offset, "")
                oDS1.SetValue("U_DAMT", oDS1.Offset, "")
                oDS1.SetValue("U_DPERCENT", oDS1.Offset, "")

                objMatrix.AutoResizeColumns()
                objMatrix.SetLineData(objMatrix.VisualRowCount)
            Case "IGMtx1"
                objMatrix1.AddRow()
                oDS2.SetValue("LineId", oDS2.Offset, objMatrix1.VisualRowCount)
                oDS2.SetValue("U_CHCK", oDS2.Offset, "")
                oDS2.SetValue("U_ITMMCDE", oDS2.Offset, "")
                oDS2.SetValue("U_ITMMGRP", oDS2.Offset, "")
                ' oDS2.SetValue("U_DPERC", oDS2.Offset, "")
                objMatrix1.SetLineData(objMatrix1.VisualRowCount)
                objMatrix1.AutoResizeColumns()
            Case "BMtx"
                objMatrix2.AddRow()
                oDS3.SetValue("Lineid", oDS3.Offset, objMatrix2.VisualRowCount)
                oDS3.SetValue("U_CHK", oDS3.Offset, "")
                oDS3.SetValue("U_BRCODE", oDS3.Offset, "")
                oDS3.SetValue("U_BCHNME", oDS3.Offset, "")
                objMatrix2.SetLineData(objMatrix2.VisualRowCount)
                objMatrix2.AutoResizeColumns()
            Case "EMtx1"
                objMatrix3.AddRow()
                oDS4.SetValue("LineId", oDS4.Offset, objMatrix3.VisualRowCount)
                oDS4.SetValue("U_BCCODE", oDS4.Offset, "")
                ' oDS4.SetValue("U_CHK", oDS4.Offset, "")
                oDS4.SetValue("U_ITTMCODE", oDS4.Offset, "")
                oDS4.SetValue("U_ITMNM", oDS4.Offset, "")
                oDS4.SetValue("U_UOM", oDS4.Offset, "")
                objMatrix3.AutoResizeColumns()
                objMatrix3.SetLineData(objMatrix3.VisualRowCount)

        End Select
    End Sub
    Sub BranchLoad(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix2 = objForm.Items.Item("BMtx").Specific
            oDS3 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PBILL_C2")
            Dim branch As String = "SELECT T0.""BPLId"", T0.""BPLName"" FROM OBPL T0"
            Dim oRgrp As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRgrp.DoQuery(branch)
            objMatrix2.Clear()
            objMatrix2.FlushToDataSource()

            If oRgrp.RecordCount > 0 Then
                For i As Integer = 1 To oRgrp.RecordCount
                    objMatrix2.AddRow()
                    oDS3.SetValue("LineId", oDS3.Offset, i)

                    oDS3.SetValue("U_BRCODE", oDS3.Offset, oRgrp.Fields.Item("BPLId").Value)
                    oDS3.SetValue("U_BCHNME", oDS3.Offset, oRgrp.Fields.Item("BPLName").Value)
                    objMatrix2.SetLineData(i)

                    oRgrp.MoveNext()
                Next
            End If


        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
    End Sub


    Public Function Duplicatebarcode(ByVal FormUID As String, ByVal MatrixUID As String) As Boolean
        Try
            For ij As Integer = 1 To objMatrix.VisualRowCount



                Dim Barcode1 As String = objMatrix.Columns.Item("BCODE").Cells.Item(ij).Specific.Value
                For ik As Integer = 1 To objMatrix.VisualRowCount
                    If ij <> ik Then
                        If Barcode1 = objMatrix.Columns.Item("BCODE").Cells.Item(ik).Specific.Value Then
                            objMain.objApplication.StatusBar.SetText("Duplicate Barcode " & Barcode1 & " Exists in row number " & ik & "...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                Next



            Next
            Dim FromDate As String = oDS.GetValue("U_FRMDATE", 0)
            Dim ToDate As String = oDS.GetValue("U_TODATE", 0)
            Dim Barcode As String = ""
            For ijk As Integer = 1 To objMatrix.VisualRowCount



                Barcode = objMatrix.Columns.Item("BCODE").Cells.Item(ijk).Specific.Value
                '              Dim Qry As String = "select distinct T0.""Code"",T1.""U_FRMDATE"", T1.""U_TODATE"", T1.""DocNum"" " &
                '"from ""@TNX_PBILL_C0"" T0 inner join ""@TNX_PBILL"" T1 on T0.""DocEntry""=T1.""DocEntry"" " &
                '" where   (T1.""U_FRMDATE"" between '" & FromDate & "' and '" & ToDate & "' " &
                '" and T0.""Code""<> '' And T1.""U_STATUS""='Active' and T1.""DocNum"" <> '" & objForm.Items.Item("DocNum").Specific.Value & "' ) or " &
                '"(T1.""U_TODATE"" between '" & FromDate & "' and '" & ToDate & "' and T0.""Code""<> '' And T1.""U_STATUS""='Active' and T1.""DocNum"" <> '" & objForm.Items.Item("DocNum").Specific.Value & "' )"


                Dim Qry As String = "select distinct T0.""Code"", T1.""U_FRMDATE"", T1.""U_TODATE"", T1.""DocNum"" " &
"from ""@TNX_PBILL_C0"" T0 inner join ""@TNX_PBILL"" T1 on T0.""DocEntry"" = T1.""DocEntry"" " &
"where (T1.""U_FRMDATE"" between '" & FromDate & "' and '" & ToDate & "' " &
"and T0.""Code"" <> '' And T1.""U_STATUS"" = 'Active' and T1.""DocNum"" <> '" & objForm.Items.Item("DocNum").Specific.Value & "') or " &
"(T1.""U_TODATE"" between '" & FromDate & "' and '" & ToDate & "' " &
"and T0.""Code"" <> '' And T1.""U_STATUS"" = 'Active' and T1.""DocNum"" <> '" & objForm.Items.Item("DocNum").Specific.Value & "')"



                Dim rs2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim chkBarCode As String
                Dim chkDocNum As String
                Dim chkFromDate As String
                Dim chkToDate As String



                rs2.DoQuery(Qry)
                If rs2.RecordCount > 0 Then
                    For i As Integer = 1 To rs2.RecordCount
                        chkBarCode = rs2.Fields.Item("Code").Value
                        chkDocNum = rs2.Fields.Item("DocNum").Value
                        chkFromDate = Format(CType(rs2.Fields.Item("U_FRMDATE").Value, Date), "yyyyMMdd")
                        chkToDate = Format(CType(rs2.Fields.Item("U_TODATE").Value, Date), "yyyyMMdd")
                        If chkBarCode = Barcode And ((chkFromDate >= FromDate Or chkFromDate <= ToDate) Or (chkToDate >= FromDate Or chkToDate <= ToDate)) Then
                            objMain.objApplication.StatusBar.SetText("This Barcode  " & Barcode & "  Already Exist in DocNum " & chkDocNum & "  between " & rs2.Fields.Item("U_FRMDATE").Value & "  to  " & rs2.Fields.Item("U_TODATE").Value & "  ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                        rs2.MoveNext()



                    Next
                End If
            Next
            Return True

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function












    Public Function Validate(ByVal FormUID As String, ByVal MatrixUID As String) As Boolean
        Try
            oDS1 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PBILL_C0")
            oDS2 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PBILL_C1")
            oDS3 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PBILL_C2")
            oDS4 = Me.objForm.DataSources.DBDataSources.Add("@TNX_PBILL_C3")

            objMatrix = objForm.Items.Item("Mtx").Specific
            objMatrix1 = objForm.Items.Item("IGMtx1").Specific
            objMatrix2 = objForm.Items.Item("BMtx").Specific
            objMatrix3 = objForm.Items.Item("EMtx1").Specific
            Select Case MatrixUID


                Case "ItemWise"
                    Dim type As String = objForm.Items.Item("PRTYP").Specific.Value
                    If type = "ItemWise" Then
                        For i As Integer = 1 To objMatrix.VisualRowCount
                            If objMatrix.Columns.Item("BCODE").Cells.Item(1).Specific.Value.ToString().Trim() = "" Then
                                'objMain.objApplication.StatusBar.SetText("Please Choose Different selection And press Load Button", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objMain.objApplication.StatusBar.SetText("Please Choose BarCode", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If

                        Next
                        For ij As Integer = 1 To objMatrix.VisualRowCount

                            Dim Barcode1 As String = objMatrix.Columns.Item("BCODE").Cells.Item(ij).Specific.Value
                            For ik As Integer = 1 To objMatrix.VisualRowCount
                                If ij <> ik Then
                                    If Barcode1 = objMatrix.Columns.Item("BCODE").Cells.Item(ik).Specific.Value Then
                                        objMain.objApplication.StatusBar.SetText("Duplicate Barcode " & Barcode1 & " Exists in row number " & ik & "...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Return False
                                    End If
                                End If
                            Next

                        Next
                        Dim FromDate As String = oDS.GetValue("U_FRMDATE", 0)
                        Dim ToDate As String = oDS.GetValue("U_TODATE", 0)
                        Dim Barcode As String = ""
                        For ijk As Integer = 1 To objMatrix.VisualRowCount

                            Barcode = objMatrix.Columns.Item("BCODE").Cells.Item(ijk).Specific.Value
                            Dim Qry As String = "select distinct T0.""Code"",T1.""U_FRMDATE"", T1.""U_TODATE"", T1.""DocNum"" " &
                                             "from ""@TNX_PBILL_C0"" T0 inner join ""@TNX_PBILL"" T1 on T0.""DocEntry""=T1.""DocEntry"" " &
                                          " where   (T1.""U_FRMDATE"" between '" & FromDate & "' and '" & ToDate & "' " &
                                         " and T0.""Code""<> '' And T1.""U_STATUS""='Active' and T1.""DocNum"" <> '" & objForm.Items.Item("DocNum").Specific.Value & "' ) or " &
                                         "(T1.""U_TODATE"" between '" & FromDate & "' and '" & ToDate & "' and T0.""Code""<> '' And T1.""U_STATUS""='Active' and T1.""DocNum"" <> '" & objForm.Items.Item("DocNum").Specific.Value & "' )"

                            Dim rs2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim chkBarCode As String
                            Dim chkDocNum As String
                            Dim chkFromDate As String
                            Dim chkToDate As String

                            rs2.DoQuery(Qry)
                            If rs2.RecordCount > 0 Then
                                For i As Integer = 1 To rs2.RecordCount
                                    chkBarCode = rs2.Fields.Item("Code").Value
                                    chkDocNum = rs2.Fields.Item("DocNum").Value
                                    chkFromDate = Format(CType(rs2.Fields.Item("U_FRMDATE").Value, Date), "yyyyMMdd")
                                    chkToDate = Format(CType(rs2.Fields.Item("U_TODATE").Value, Date), "yyyyMMdd")
                                    If chkBarCode = Barcode And ((chkFromDate >= FromDate Or chkFromDate <= ToDate) Or (chkToDate >= FromDate Or chkToDate <= ToDate)) Then
                                        objMain.objApplication.StatusBar.SetText("This Barcode  " & Barcode & "  Already Exist in DocNum " & chkDocNum & "  between " & rs2.Fields.Item("U_FRMDATE").Value & "  to  " & rs2.Fields.Item("U_TODATE").Value & "  ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Return False
                                    End If
                                    rs2.MoveNext()

                                Next
                            End If
                        Next
                    End If
                    For j As Integer = 1 To objMatrix2.VisualRowCount
                        Dim oCheckBox As SAPbouiCOM.CheckBox
                        oCheckBox = objMatrix2.Columns.Item("CHK").Cells.Item(j).Specific
                        For k As Integer = 1 To objMatrix2.VisualRowCount
                            oCheckBox = objMatrix2.Columns.Item("CHK").Cells.Item(k).Specific
                            If oCheckBox.Checked = True Then
                                Exit For
                            End If
                        Next
                        If oCheckBox.Checked = False Then
                            objMain.objApplication.StatusBar.SetText("Please Select Atleast one Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    Next

                Case "ItemGroup"

                    Dim type As String = objForm.Items.Item("PRTYP").Specific.Value
                    If type = "ItemGroup" Then
                        objMatrix1 = objForm.Items.Item("Mtx").Specific
                        Dim oCheckBox As SAPbouiCOM.CheckBox
                        Dim flag As Integer = 0
                        For k As Integer = 1 To objMatrix1.VisualRowCount
                            oCheckBox = objMatrix1.Columns.Item("CHK").Cells.Item(k).Specific
                            If (oCheckBox.Checked = True) Then 'Or (objMatrix1.Columns.Item("DPERC").Cells.Item(k).Specific.Value > 0)
                                If oCheckBox.Checked = True Then
                                    If objMatrix1.Columns.Item("DPERCENT").Cells.Item(k).Specific.Value <= 0 Then
                                        objMain.objApplication.StatusBar.SetText("Please Enter Discount % greater than zero in Line No " & k, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Return False
                                    End If
                                ElseIf objMatrix1.Columns.Item("DPERCENT").Cells.Item(k).Specific.Value > 0 Then
                                    If oCheckBox.Checked = False Then
                                        objMain.objApplication.StatusBar.SetText("Please Check Item Group in Line No " & k, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Return False
                                    End If
                                End If
                            Else
                                flag = flag + 1
                                If flag = objMatrix1.VisualRowCount Then
                                    objMain.objApplication.StatusBar.SetText("Please Check Any Group Item", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Return False
                                End If
                            End If
                        Next
                        For i As Integer = 1 To objMatrix3.VisualRowCount
                            If objMatrix3.Columns.Item("BCODE").Cells.Item(1).Specific.Value.ToString().Trim() = "" Then
                                'objMain.objApplication.StatusBar.SetText("Please Choose Different selection And press Load Button", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objMain.objApplication.StatusBar.SetText("Please Choose BarCode in Exclusion Tab ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                        Next
                        'Akhila 
                        'duplicate itm
                        For ij As Integer = 1 To objMatrix3.VisualRowCount
                            Dim Barcode1 As String = objMatrix3.Columns.Item("BCODE").Cells.Item(ij).Specific.Value
                            For ik As Integer = 1 To objMatrix3.VisualRowCount
                                If ij <> ik Then
                                    If Barcode1 = objMatrix3.Columns.Item("BCODE").Cells.Item(ik).Specific.Value Then
                                        objMain.objApplication.StatusBar.SetText("Duplicate Barcode " & Barcode1 & " Exists in row number " & ik & "...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Return False
                                    End If
                                End If
                            Next
                        Next
                    End If

                    Dim FromDate As String = oDS.GetValue("U_FRMDATE", 0)
                    Dim ToDate As String = oDS.GetValue("U_TODATE", 0)
                    Dim Barcode As String = ""
                    For ijk As Integer = 1 To objMatrix3.VisualRowCount
                        Barcode = objMatrix3.Columns.Item("BCODE").Cells.Item(ijk).Specific.Value

                        Dim Qry As String = "select distinct T0.""Code"",T1.""U_FRMDATE"", T1.""U_TODATE"", T1.""DocNum""  from ""@TNX_PBILL"" T0 inner join ""@TNX_PBILL"" T1 on T0.""DocEntry""=T1.""DocEntry"" " &
                                                                " where  T1.""U_STATUS""='Active' and  (T1.""U_FRMDATE"" between '" & FromDate & "' and '" & ToDate & "' and  T0.""Code""<> '')  or (T1.""U_TODATE"" between '" & FromDate & "' and '" & ToDate & "' and  T0.""Code""<> '')"
                        Dim rs2 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim chkBarCode As String
                        Dim chkDocNum As String
                        Dim chkFromDate As String
                        Dim chkToDate As String

                        rs2.DoQuery(Qry)
                        If rs2.RecordCount > 0 Then
                            For i As Integer = 1 To rs2.RecordCount
                                chkBarCode = rs2.Fields.Item("Code").Value
                                chkDocNum = rs2.Fields.Item("DocNum").Value
                                chkFromDate = Format(CType(rs2.Fields.Item("U_FRMDATE").Value, Date), "yyyyMMdd")
                                chkToDate = Format(CType(rs2.Fields.Item("U_TODATE").Value, Date), "yyyyMMdd")
                                If chkBarCode = Barcode And ((chkFromDate >= FromDate Or chkFromDate <= ToDate) Or (chkToDate >= FromDate Or chkToDate <= ToDate)) Then
                                    objMain.objApplication.StatusBar.SetText("This Barcode " & Barcode & " Already Exist in Document Number  " & chkDocNum & "  between " & rs2.Fields.Item("U_FRMDATE").Value & "  to  " & rs2.Fields.Item("U_TODATE").Value & "  ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Return False
                                End If
                                rs2.MoveNext()
                            Next
                        End If
                    Next

                    For j As Integer = 1 To objMatrix2.VisualRowCount
                        Dim oCheckBox As SAPbouiCOM.CheckBox
                        oCheckBox = objMatrix2.Columns.Item("CHK").Cells.Item(j).Specific
                        For k As Integer = 1 To objMatrix2.VisualRowCount
                            oCheckBox = objMatrix2.Columns.Item("CHK").Cells.Item(k).Specific
                            If oCheckBox.Checked = True Then
                                Exit For
                            End If
                        Next
                        If oCheckBox.Checked = False Then
                            objMain.objApplication.StatusBar.SetText("Please Select Atleast one Branch", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    Next
            End Select
            If objForm.Items.Item("SBCODE").Specific.Value <> "" Then
                Dim bcode As String = objForm.Items.Item("SBCODE").Specific.Value
                Dim Afterchk As String = " "
                objForm.Items.Item("SBCODE").Specific.Value = Afterchk
            End If
            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Dim objForm As SAPbouiCOM.Form
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)
        Try
            If eventInfo.FormUID = objForm.UniqueID Then
                If (eventInfo.BeforeAction = True) Then
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        objForm.EnableMenu("1292", True)
                        objForm.EnableMenu("774", True)
                    End If
                End If
            End If
            ' System.Diagnostics.Process.Start()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub LoadSubGroups1(objForm As SAPbouiCOM.Form, itemGroup As String)

        If String.IsNullOrEmpty(itemGroup) OrElse Not IsNumeric(itemGroup.Trim()) Then

            'objMain.objApplication.MessageBox("Invalid or empty Item Group. Exiting.")
            Exit Sub
        End If

        Dim objCombo1 As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("SubGrp1").Specific, SAPbouiCOM.ComboBox)

        ' === CLEAR EXISTING COMBO VALUES ===
        If objCombo1.ValidValues.Count > 0 Then
            For i As Integer = objCombo1.ValidValues.Count - 1 To 0 Step -1
                objCombo1.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
        End If
        ' === QUERY SUBGROUPS ===
        Dim oRS As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

        ' Build query with numeric-safe filter
        'Dim sql As String = "SELECT DISTINCT T0.""U_VSubGroup"" " &
        '                "FROM OITM T0 " &
        '                "INNER JOIN OITB T1 ON T0.""ItmsGrpCod"" = T1.""ItmsGrpCod"" " &
        '                "WHERE T0.""ItmsGrpCod"" = " & itemGroup & " " & ' No quotes for numeric value
        '                "AND T0.""U_VSubGroup"" IS NOT NULL " &
        '                "ORDER BY T0.""U_VSubGroup"""
        Try
            objCombo1.ValidValues.Add("", "")
        Catch ex As Exception
        End Try
        Dim sql As String = "SELECT DISTINCT T0.""U_VSubGroup"" " &
                        "FROM OITM T0 " &
                        "WHERE T0.""ItmsGrpCod"" = " & itemGroup & " " &
                        "AND T0.""U_VSubGroup"" IS NOT NULL " &
                        "ORDER BY T0.""U_VSubGroup"""

        oRS.DoQuery(sql)







        ' === ADD TO COMBO AND SELECT FIRST ===
        Dim firstValue As String = ""
        While Not oRS.EoF
            Dim subGroup As String = oRS.Fields.Item(0).Value.ToString()
            If Not String.IsNullOrEmpty(subGroup) Then
                objCombo1.ValidValues.Add(subGroup, subGroup)
                If firstValue = "" Then
                    firstValue = subGroup
                End If
            End If
            oRS.MoveNext()
        End While
        'If objCombo1.ValidValues.Count > 0 AndAlso Not String.IsNullOrEmpty(firstValue) Then
        '    objCombo1.Select(firstValue, SAPbouiCOM.BoSearchKey.psk_ByValue)
        'End If
    End Sub

    Private Sub LoadSubSubGroups(objForm As SAPbouiCOM.Form, itemGroup As String, subGroup As String)
        If String.IsNullOrEmpty(itemGroup) OrElse Not IsNumeric(itemGroup.Trim()) Then
            Exit Sub
        End If
        Dim objCombo As SAPbouiCOM.ComboBox = CType(objForm.Items.Item("SubSubGrp1").Specific, SAPbouiCOM.ComboBox)
        If objCombo.ValidValues.Count > 0 Then
            For i As Integer = objCombo.ValidValues.Count - 1 To 0 Step -1
                objCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
        End If
        Dim oRS As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        'Dim sql As String = "SELECT DISTINCT T0.""U_VSubSubGroup"" " &
        '                    "FROM OITM T0 " &
        '                    "INNER JOIN OITB T1 ON T0.""ItmsGrpCod"" = T1.""ItmsGrpCod"" " &
        '                    "WHERE T0.""ItmsGrpCod"" = " & itemGroup & " " & ' No quotes for numeric
        '                    "AND T0.""U_VSubSubGroup"" IS NOT NULL " &
        '                    "ORDER BY T0.""U_VSubSubGroup"""

        Try
            objCombo.ValidValues.Add("", "")
        Catch ex As Exception
        End Try
        Dim sql As String = "SELECT DISTINCT T0.""U_VSubSubGroup"" " &
                            "FROM OITM T0 " &
                            "WHERE T0.""ItmsGrpCod"" = " & itemGroup & " " &
                            "AND T0.""U_VSubSubGroup"" IS NOT NULL " &
                            "AND T0.""U_VSubGroup""='" & subGroup & "' " &
                            "ORDER BY T0.""U_VSubSubGroup"""

        oRS.DoQuery(sql)
        Dim firstValue As String = ""
        While Not oRS.EoF
            Dim subsubGroup As String = oRS.Fields.Item(0).Value.ToString()
            If Not String.IsNullOrEmpty(subsubGroup) Then
                objCombo.ValidValues.Add(subsubGroup, subsubGroup)
                If firstValue = "" Then
                    firstValue = subsubGroup
                End If
            End If
            oRS.MoveNext()
        End While
        'If objCombo.ValidValues.Count > 0 AndAlso Not String.IsNullOrEmpty(firstValue) Then
        '    objCombo.Select(firstValue, SAPbouiCOM.BoSearchKey.psk_ByValue)
        'End If
    End Sub



    Private Sub LoadExcelToMatrix(objForm As SAPbouiCOM.Form, filePath As String)
        Try
            Dim oDataTable As SAPbouiCOM.DataTable
            ' Check and create DT_MATRIX if not exists
            If objForm.DataSources.DataTables.Count = 0 OrElse objForm.DataSources.DataTables.Item("DT_BTMTX") Is Nothing Then
                objForm.DataSources.DataTables.Add("DT_BTMTX")
            End If
            oDataTable = objForm.DataSources.DataTables.Item("DT_BTMTX")
            ' If columns are not already added, add them
            If oDataTable.Columns.Count = 0 Then
                oDataTable.Columns.Add("Barcode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                oDataTable.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100)
                oDataTable.Columns.Add("ItemDescription", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 254)
                oDataTable.Columns.Add("UOM", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50)
                'oDataTable.Columns.Add("FromDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
                'oDataTable.Columns.Add("ToDate", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20)
                oDataTable.Columns.Add("PRICE", SAPbouiCOM.BoFieldsType.ft_Price, 20)
                'oDataTable.Columns.Add("RetailPriceWithVAT", SAPbouiCOM.BoFieldsType.ft_Price, 20)
                'oDataTable.Columns.Add("OfferPriceWithVAT", SAPbouiCOM.BoFieldsType.ft_Price, 20)
                'oDataTable.Columns.Add("DiscountAmount", SAPbouiCOM.BoFieldsType.ft_Price, 20)
                'oDataTable.Columns.Add("DiscountPercentage", SAPbouiCOM.BoFieldsType.ft_Percent, 20)
            End If
            ' Ensure DataTable exists
            If objForm.DataSources.DataTables.Count = 0 OrElse objForm.DataSources.DataTables.Item("DT_BTMTX") Is Nothing Then
                objForm.DataSources.DataTables.Add("DT_BTMTX")
            End If
            Dim objMatrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx").Specific, SAPbouiCOM.Matrix)
            Dim oDataSource As SAPbouiCOM.DataTable = objForm.DataSources.DataTables.Item("DT_BTMTX")
            ' Clear rows safely
            If oDataSource.Rows.Count > 0 Then
                oDataSource.Rows.Clear()
            End If
            ' Load Excel file using OLEDB
            Dim connStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties='Excel 12.0 Xml;HDR=YES';"
            Dim dt As New DataTable()
            Using conn As New OleDb.OleDbConnection(connStr)
                conn.Open()
                Dim cmd As New OleDb.OleDbCommand("SELECT * FROM [Sheet1$]", conn)
                Dim adapter As New OleDb.OleDbDataAdapter(cmd)
                adapter.Fill(dt)
                conn.Close()
            End Using
            ' Fill SAP DataTable
            For i As Integer = 0 To dt.Rows.Count - 1
                oDataTable.Rows.Add()
                oDataTable.SetValue("Barcode", i, dt.Rows(i)("Barcode").ToString())
                oDataTable.SetValue("ItemCode", i, dt.Rows(i)("Item Code").ToString())
                oDataTable.SetValue("ItemDescription", i, dt.Rows(i)("Item Description").ToString())
                oDataTable.SetValue("UOM", i, dt.Rows(i)("UOM").ToString())
                oDataTable.SetValue("PRICE", i, dt.Rows(i)("PRICE").ToString())
                'oDataTable.SetValue("DAMNT", i, dt.Rows(i)("DiscountAmount").ToString())
                'oDataTable.SetValue("DPERCENT", i, dt.Rows(i)("DiscountPercentage").ToString())

            Next
            ' Load into matrix

            objMatrix.Clear()
            oDS1.Clear()
            objMatrix.FlushToDataSource()

            'objMatrix.LoadFromDataSource()
            'Dim objMatrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Mtx").Specific, SAPbouiCOM.Matrix)
            ' Example: binding each column of matrix to datatable
            ' Dim col As SAPbouiCOM.Column
            ' Clear existing rows in matrix (optional)
            'If objMatrix.RowCount > 0 Then
            '    For i As Integer = objMatrix.RowCount To 1 Step -1
            '        objMatrix.DeleteRow(i)
            '    Next
            'End If
            ' Add rows and set values manually

            Dim FromDateStr As String = objForm.Items.Item("FRMDATE").Specific.Value
            Dim ToDateStr As String = objForm.Items.Item("TODATE").Specific.Value
            For i As Integer = 0 To oDataSource.Rows.Count - 1
                objMatrix.AddRow()
                oDS1.SetValue("LineId", oDS1.Offset, objMatrix.VisualRowCount)
                oDS1.SetValue("Code", oDS1.Offset, oDataSource.GetValue("Barcode", i).ToString())
                oDS1.SetValue("U_ITMCODE", oDS1.Offset, oDataSource.GetValue("ItemCode", i).ToString())
                oDS1.SetValue("U_ITMNM", oDS1.Offset, oDataSource.GetValue("ItemDescription", i).ToString())
                oDS1.SetValue("U_UOM", oDS1.Offset, oDataSource.GetValue("UOM", i).ToString())

                oDS1.SetValue("U_PRICE", oDS1.Offset, oDataSource.GetValue("PRICE", i).ToString())

                oDS1.SetValue("U_DAMT", oDS1.Offset, 0)
                oDS1.SetValue("U_DPERCENT", oDS1.Offset, 0)
                objMatrix.SetLineData(objMatrix.VisualRowCount)

                'For i As Integer = 0 To oDataSource.Rows.Count - 1
                'objMatrix.AddRow()
                'objMatrix.Columns.Item("LineId").Cells.Item(i + 1).Specific.Value = (i + 1).ToString()
                'objMatrix.Columns.Item("BCODE").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("Barcode", i).ToString()
                'objMatrix.Columns.Item("ITMCODE").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("ItemCode", i).ToString()
                'objMatrix.Columns.Item("ITMNM").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("ItemDescription", i).ToString()
                'objMatrix.Columns.Item("UOM").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("UOM", i).ToString()
                'objMatrix.Columns.Item("PRICE").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("PRICE", i).ToString()
                ''objMatrix.Columns.Item("DAMNT").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("DiscountAmount", i).ToString()
                ''objMatrix.Columns.Item("DPERCENT").Cells.Item(i + 1).Specific.Value = oDataSource.GetValue("DiscountPercentage", i).ToString()
            Next
        Catch ex As Exception
            objMain.objApplication.MessageBox("Error loading Excel to matrix: " & ex.Message)
        End Try
    End Sub
    Private Function ValidatePromotionName(ByVal formUID As String) As Boolean
        Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(formUID)
        Dim oEditText As SAPbouiCOM.EditText = objForm.Items.Item("PNAME").Specific
        Dim promoName As String = oEditText.Value.Trim()
        If String.IsNullOrWhiteSpace(promoName) Then
            objMain.objApplication.StatusBar.SetText("Promotion Name cannot be empty.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        '  Max 10 characters
        If promoName.Length > 10 Then
            'objMain.objApplication.MessageBox("Promotion Name must be a maximum of 10 characters.")
            objMain.objApplication.StatusBar.SetText("Promotion Name must be a maximum of 10 characters.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        ' No spaces allowed
        If promoName.Contains(" ") Then
            'objMain.objApplication.MessageBox("Promotion Name Should not contain spaces.")
            objMain.objApplication.StatusBar.SetText("Promotion Name Should not contain spaces.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If

        '  Must be unique — check in the database
        Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs.DoQuery("SELECT COUNT(*) AS Cnt FROM ""@TNX_PBILL"" WHERE ""U_PNAME""= '" & promoName.Replace("'", "''") & "'")
        If rs.Fields.Item("Cnt").Value > 0 Then
            ' objMain.objApplication.MessageBox("Promotion Name already exists. Please enter a unique name.")
            objMain.objApplication.StatusBar.SetText("Promotion Name already exists. Please enter a unique name", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        Return True
    End Function
    Function validateDiscountType(ByVal FormUiD As String)
        Try

            objForm = objMain.objApplication.Forms.Item(FormUiD)
            If objForm.Items.Item("DISCTYP1").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText(" Discount Type Cannot be Null", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True
        Catch ex As Exception

        End Try
    End Function


    'Function validateDiscountpercentage(ByVal FormUiD As String)
    '    Try
    '        Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUiD)

    '        Dim oCombo As SAPbouiCOM.ComboBox = CType(oForm.Items.Item("DISCTYP1").Specific, SAPbouiCOM.ComboBox)
    '        If oCombo.Selected IsNot Nothing AndAlso oCombo.Selected.Value = "P" Then
    '            For i As Integer = 1 To objMatrix.VisualRowCount
    '                Dim valueStr As String = objMatrix.Columns.Item("DPERCENT").Cells.Item(i).Specific.Value.ToString().Trim()
    '                If valueStr = "" OrElse CDbl(valueStr) = 0 Then
    '                    objMain.objApplication.StatusBar.SetText("Please enter a valid Discount Percentage in row " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                    Return False
    '                End If
    '            Next
    '        End If

    '        Return True
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        Return False
    '    End Try
    'End Function

    'Function validateDiscountAmount(ByVal FormUiD As String)
    '    Try
    '        Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUiD)

    '        Dim oCombo As SAPbouiCOM.ComboBox = CType(oForm.Items.Item("DISCTYP1").Specific, SAPbouiCOM.ComboBox)
    '        If oCombo.Selected IsNot Nothing AndAlso oCombo.Selected.Value = "A" Then
    '            For i As Integer = 1 To objMatrix.VisualRowCount
    '                Dim valueStr As String = objMatrix.Columns.Item("DAMNT").Cells.Item(i).Specific.Value.ToString().Trim()
    '                If valueStr = "" OrElse CDbl(valueStr) = 0 Then
    '                    objMain.objApplication.StatusBar.SetText("Please enter a valid Discount Amount in row " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                    Return False
    '                End If
    '            Next
    '        End If

    '        Return True
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        Return False
    '    End Try
    'End Function
    Function validateDiscountAmount(ByVal FormUiD As String)
        Try
            Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUiD)
            Dim oCombo As SAPbouiCOM.ComboBox = CType(oForm.Items.Item("DISCTYP1").Specific, SAPbouiCOM.ComboBox)
            If oCombo.Selected IsNot Nothing AndAlso oCombo.Selected.Value = "A" Then
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    Dim itemCode As String = objMatrix.Columns.Item("ITMCODE").Cells.Item(i).Specific.Value.ToString().Trim()
                    If itemCode <> "" Then
                        Dim valueStr As String = objMatrix.Columns.Item("DAMT").Cells.Item(i).Specific.Value.ToString().Trim()
                        If valueStr = "" OrElse CDbl(valueStr) = 0 Then
                            objMain.objApplication.StatusBar.SetText("Please enter a valid Discount Amount in row " & i & " for item " & itemCode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                Next
            End If
            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Function validateDiscountpercentage(ByVal FormUiD As String)
        Try
            Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUiD)

            Dim oCombo As SAPbouiCOM.ComboBox = CType(oForm.Items.Item("DISCTYP1").Specific, SAPbouiCOM.ComboBox)

            If oCombo.Selected IsNot Nothing AndAlso oCombo.Selected.Value = "P" Then
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.Columns.Item("LineId").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                    Dim itemCode As String = objMatrix.Columns.Item("ITMCODE").Cells.Item(i).Specific.Value.ToString().Trim()
                    If itemCode <> "" Then
                        Dim valueStr As String = objMatrix.Columns.Item("DPERCENT").Cells.Item(i).Specific.Value.ToString().Trim()
                        If valueStr = "" OrElse CDbl(valueStr) = 0 Then
                            objMain.objApplication.StatusBar.SetText("Please enter a valid Discount Percentage in row " & i & " for item " & itemCode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If

                Next
            End If

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function






End Class
