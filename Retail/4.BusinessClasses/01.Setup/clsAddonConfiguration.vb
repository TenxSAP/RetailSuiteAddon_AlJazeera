Public Class clsAddonConfiguration

#Region "        Declaration        "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim objComboBox, objCombobox1, objCombobox2 As SAPbouiCOM.ComboBox
    Dim objCheckBox As SAPbouiCOM.CheckBox
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("AddonConfiguration1.xml", "ME_ACFG_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("ME_ACFG_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_ACFG")

            Dim Check As String = "Select ""DocNum"" From ""@ME_ACFG"" "
            Dim oRsChec As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsChec.DoQuery(Check)

            objForm.Items.Item("29").TextStyle = 1
            objForm.Items.Item("30").TextStyle = 1
            objForm.Items.Item("27").TextStyle = 1
            objForm.Items.Item("31").TextStyle = 1
            objForm.Items.Item("58").TextStyle = 1
            ' objForm.Items.Item("34").TextStyle = 1

            objForm.Items.Item("7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("11").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("21").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("21").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("31").TextStyle = SAPbouiCOM.BoTextStyle.ts_UNDERLINE
            'Dim objComboBox1 As SAPbouiCOM.ComboBox = objForm.Items.Item("33").Specific
            'objMain.objUtilities.ComboBoxLoadValues(objForm.Items.Item("33").Specific, "Select ""posID"" , ""name""  From OHPS")
            'objForm.Items.Item("33").DisplayDesc = True
            'objComboBox1.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly


            'objComboBox1 = objForm.Items.Item("35").Specific
            'objMain.objUtilities.ComboBoxLoadValues(objForm.Items.Item("35").Specific, "Select ""posID"" , ""name""  From OHPS")
            'objForm.Items.Item("35").DisplayDesc = True
            'objComboBox1.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            objComboBox = objForm.Items.Item("60").Specific
            objMain.objUtilities.ComboBoxLoadValues(objComboBox, "Select ""Indicator"" , ""Indicator"" From OPID ")
            objForm.Items.Item("60").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            objCombobox1 = objForm.Items.Item("VatCode").Specific
            objMain.objUtilities.ComboBoxLoadValues(objCombobox1, "Select ""Code"" , ""Name"" From OVTG where ""Account""<>'' ")
            objForm.Items.Item("VatCode").DisplayDesc = True
            objCombobox1.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            objCombobox2 = objForm.Items.Item("PriceLst").Specific
            'objMain.objUtilities.ComboBoxLoadValues(objCombobox2, "Select ""ListNum"" , ""ListName"" From OPLN  ")
            objMain.objUtilities.ComboBoxLoadValues(objCombobox2, "SELECT ""ListNum"", ""ListNum"" || ' - ' || ""ListName"" FROM OPLN")
            objForm.Items.Item("PriceLst").DisplayDesc = True
            objCombobox2.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            If oRsChec.RecordCount = 0 Then
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, 1)
                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                objForm.Items.Item("5").Specific.Value = 1
                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Else
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                objForm.Items.Item("5").Specific.Value = 1
                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            End If

            objForm.Items.Item("Item_0").Width = 250
            objForm.Items.Item("Item_12").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@ME_ACFG")


                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        If oCFL.UniqueID = "CFL_MNWH" Then
                            oDBs_Head.SetValue("U_MEMWHSCD", oDBs_Head.Offset, oDT.GetValue("WhsCode", 0))
                            oDBs_Head.SetValue("U_MEMWHSNM", oDBs_Head.Offset, oDT.GetValue("WhsName", 0))
                        End If
                        If oCFL.UniqueID = "CFL_TRSIT" Then
                            oDBs_Head.SetValue("U_METWHSCD", oDBs_Head.Offset, oDT.GetValue("WhsCode", 0))
                            oDBs_Head.SetValue("U_METWHSNM", oDBs_Head.Offset, oDT.GetValue("WhsName", 0))
                        End If
                        If oCFL.UniqueID = "CFL_DWHS" Then
                            oDBs_Head.SetValue("U_MEDWHSCD", oDBs_Head.Offset, oDT.GetValue("WhsCode", 0))
                            oDBs_Head.SetValue("U_MEDWHSNM", oDBs_Head.Offset, oDT.GetValue("WhsName", 0))
                        End If
                        If oCFL.UniqueID = "CFL_EXACT" Then
                            oDBs_Head.SetValue("U_MEEXACCT", oDBs_Head.Offset, oDT.GetValue("AcctCode", 0))
                        End If
                        If oCFL.UniqueID = "CFL_ENACT" Then
                            oDBs_Head.SetValue("U_MEENTACT", oDBs_Head.Offset, oDT.GetValue("AcctCode", 0))
                        End If
                        If oCFL.UniqueID = "CFL_SACACT" Then
                            oDBs_Head.SetValue("U_SAGCRACT", oDBs_Head.Offset, oDT.GetValue("AcctCode", 0))
                        End If
                        If oCFL.UniqueID = "CFL_RNTAC" Then
                            oDBs_Head.SetValue("U_RNTLACT", oDBs_Head.Offset, oDT.GetValue("AcctCode", 0))
                        End If
                        If oCFL.UniqueID = "CFL_REG" Then
                            oDBs_Head.SetValue("U_REG", oDBs_Head.Offset, oDT.GetValue("AcctCode", 0))
                        End If
                        If oCFL.UniqueID = "CFL_REGCH" Then
                            oDBs_Head.SetValue("U_RChar", oDBs_Head.Offset, oDT.GetValue("AcctCode", 0))
                        End If
                        If oCFL.UniqueID = "CFL_RDA" Then
                            oDBs_Head.SetValue("U_Round", oDBs_Head.Offset, oDT.GetValue("AcctCode", 0))
                        End If
                        If oCFL.UniqueID = "CFL_GL" Then
                            oDBs_Head.SetValue("U_SuppReb", oDBs_Head.Offset, oDT.GetValue("AcctCode", 0))
                        End If
                        'If oCFL.UniqueID = "CFL_EDSAC" Then
                        '    oDBs_Head.SetValue("U_MEDMEXAC", oDBs_Head.Offset, oDT.GetValue("AcctCode", 0))
                        'End If
                        'If oCFL.UniqueID = "CFL_ESTAC" Then
                        '    oDBs_Head.SetValue("U_METSEXAC", oDBs_Head.Offset, oDT.GetValue("AcctCode", 0))
                        'End If
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "ME_ACFG" And pVal.BeforeAction = False Then
                Me.CreateForm()
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

End Class
