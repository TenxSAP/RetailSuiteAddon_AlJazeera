Imports TaskScheduler

Public Class clsLSIntegrationStatusScreen

#Region "        Declaration        "
    Dim objForm, objICForm As SAPbouiCOM.Form
    Dim oGrid As SAPbouiCOM.Grid
    Dim oDt As SAPbouiCOM.DataTable
    Dim oEditText As SAPbouiCOM.EditText
    Dim objComboBox As SAPbouiCOM.ComboBox
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("LSIntegrationStatusScreen.xml", "TNX_LSINTSTS_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("TNX_LSINTSTS_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objMain.objApplication.StatusBar.SetText("Please wait... Form is Initialilsing....", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            objForm.DataSources.UserDataSources.Add("U_FROMDT", SAPbouiCOM.BoDataType.dt_DATE)
            objForm.DataSources.UserDataSources.Add("U_TODT", SAPbouiCOM.BoDataType.dt_DATE)
            objForm.DataSources.UserDataSources.Add("U_DOCTYP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_STATUS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            objForm.DataSources.UserDataSources.Add("U_NOR", SAPbouiCOM.BoDataType.dt_LONG_NUMBER)
            objForm.DataSources.UserDataSources.Add("U_NOF", SAPbouiCOM.BoDataType.dt_LONG_NUMBER)
            objForm.DataSources.UserDataSources.Add("U_NOS", SAPbouiCOM.BoDataType.dt_LONG_NUMBER)
            objForm.DataSources.UserDataSources.Add("U_PATH", SAPbouiCOM.BoDataType.dt_LONG_TEXT)
            objForm.DataSources.UserDataSources.Add("U_LRT", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_NRT", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_SSTS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            oEditText = objForm.Items.Item("21").Specific
            oEditText.DataBind.SetBound(True, "", "U_SSTS")
            objForm.Items.Item("21").AffectsFormMode = False

            oEditText = objForm.Items.Item("19").Specific
            oEditText.DataBind.SetBound(True, "", "U_LRT")
            objForm.Items.Item("19").AffectsFormMode = False

            oEditText = objForm.Items.Item("22").Specific
            oEditText.DataBind.SetBound(True, "", "U_NRT")
            objForm.Items.Item("22").AffectsFormMode = False

            oEditText = objForm.Items.Item("24").Specific
            oEditText.DataBind.SetBound(True, "", "U_PATH")
            objForm.Items.Item("24").AffectsFormMode = False
            objForm.DataSources.UserDataSources.Item("U_PATH").ValueEx = "C:\Logs\Logs_SAPDataConnector"

            oEditText = objForm.Items.Item("10").Specific
            oEditText.DataBind.SetBound(True, "", "U_NOR")
            objForm.Items.Item("10").AffectsFormMode = False

            oEditText = objForm.Items.Item("12").Specific
            oEditText.DataBind.SetBound(True, "", "U_NOF")
            objForm.Items.Item("12").AffectsFormMode = False

            oEditText = objForm.Items.Item("14").Specific
            oEditText.DataBind.SetBound(True, "", "U_NOS")
            objForm.Items.Item("14").AffectsFormMode = False

            oEditText = objForm.Items.Item("3").Specific
            oEditText.DataBind.SetBound(True, "", "U_FROMDT")
            objForm.Items.Item("3").AffectsFormMode = False
            objForm.DataSources.UserDataSources.Item("U_FROMDT").Value = DateTime.Now.ToString("yyyyMMdd")

            oEditText = objForm.Items.Item("5").Specific
            oEditText.DataBind.SetBound(True, "", "U_TODT")
            objForm.Items.Item("5").AffectsFormMode = False
            objForm.DataSources.UserDataSources.Item("U_TODT").Value = DateTime.Now.ToString("yyyyMMdd")

            objComboBox = objForm.Items.Item("17").Specific
            objComboBox.DataBind.SetBound(True, "", "U_STATUS")
            objForm.Items.Item("17").AffectsFormMode = False           
            objComboBox.ValidValues.Add("All", "All")
            objComboBox.ValidValues.Add("Success", "Success")
            objComboBox.ValidValues.Add("Fail", "Fail")
            objForm.Items.Item("17").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            objComboBox = objForm.Items.Item("7").Specific
            objComboBox.DataBind.SetBound(True, "", "U_DOCTYP")
            objForm.Items.Item("7").AffectsFormMode = False
            objComboBox.ValidValues.Add("ART", "Article")
            objComboBox.ValidValues.Add("UOM", "UOM")
            objComboBox.ValidValues.Add("DIV", "Division")
            objComboBox.ValidValues.Add("CAT", "Item Category")
            objComboBox.ValidValues.Add("RPG", "Retail Product Group")
            objComboBox.ValidValues.Add("ISG", "Item Special Group")
            objComboBox.ValidValues.Add("BRD", "Brands")
            objComboBox.ValidValues.Add("BSP", "BaseUom Prices")
            objComboBox.ValidValues.Add("OTP", "OtherUom Prices")
            objComboBox.ValidValues.Add("BAR", "Barcode")
            objComboBox.ValidValues.Add("UMG", "UOM Group")
            objComboBox.ValidValues.Add("CST", "Customers")
            objComboBox.ValidValues.Add("SEP", "Sale Employess")
            objForm.Items.Item("7").DisplayDesc = True
            objComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            oDt = objForm.DataSources.DataTables.Add("dt1")
            oDt = objForm.DataSources.DataTables.Item("dt1")

            Me.GetTaksDetails()

            objMain.objApplication.StatusBar.SetText("Form loaded successfully.....countinue.!!.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        Finally
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "TNX_LSINTSTS" And pVal.BeforeAction = False Then
                Me.CreateForm()
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "15" And pVal.BeforeAction = False Then
                        oGrid = objForm.Items.Item("8").Specific
                        If objForm.Items.Item("3").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        ElseIf objForm.Items.Item("5").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            Me.LoadGrid(objForm.UniqueID)
                        End If
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub LoadGrid(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            Dim sSuccess As Integer = 0
            Dim sFail As Integer = 0

            objMain.objApplication.StatusBar.SetText("Loading initiated...please wait a while...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            objForm.Items.Item("15").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oGrid = objForm.Items.Item("8").Specific
            oGrid.DataTable = oDt
            Try
                oGrid.DataTable.Clear()
            Catch ex As Exception
            End Try

            objComboBox = objForm.Items.Item("17").Specific
            Dim Status As String = objComboBox.Selected.Value.ToString

            Dim Query As String = ""
            Dim oEditTextCol As SAPbouiCOM.EditTextColumn
            objComboBox = objForm.Items.Item("7").Specific
            Select Case objComboBox.Selected.Value.ToString
                Case "ART"
                    Query = " Select AA.*  From (" & _
                            " Select ""ITEMCODE"" As ""Code"" , 'Success' As ""Status"" , ""ERRORMSG"" As  ""Response""   From ""10X_ITEM"" Where IFNULL(""RESPONSE"",'') <> '' " & _
                            " And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "'  Union ALL " & _
                            " Select ""ITEMCODE"" As ""Code"" , 'Fail' As ""Status"" , ""ERRORMSG"" As  ""Response""  From ""10X_ITEM"" Where IFNULL(""RESPONSE"",'') = '' " & _
                            "  And ""ERRORMSG""  <>'' And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "') AA  "
                    If Status = "All" Then
                        Query = Query & " Group By AA.""Code"" , AA.""Status"" , AA.""Response"" "
                    Else
                        Query = Query & " Where  AA.""Status"" = '" & Status & "' Group By AA.""Code"" , AA.""Status"" , AA.""Response"" "
                    End If
                    oGrid.DataTable.ExecuteQuery(Query)
                    oEditTextCol = oGrid.Columns.Item("Code")
                    oEditTextCol.LinkedObjectType = 4
                Case "UOM"
                    Query = " Select AA.*  From (" & _
                            " Select ""UOMCODE"" As ""Code"" , 'Success' As ""Status"" , ""ERRORMSG"" As  ""Response""   From ""10X_UOM"" Where IFNULL(""RESPONSE"",'') <> '' " & _
                            " And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "'  Union ALL " & _
                            " Select ""UOMCODE"" As ""Code"" , 'Fail' As ""Status"" , ""ERRORMSG"" As  ""Response""  From ""10X_UOM"" Where IFNULL(""RESPONSE"",'') = '' " & _
                            "  And ""ERRORMSG""  <>'' And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "') AA    "
                    If Status = "All" Then
                        Query = Query & " Group By AA.""Code"" , AA.""Status"" , AA.""Response"" "
                    Else
                        Query = Query & " Where  AA.""Status"" = '" & Status & "' Group By AA.""Code"" , AA.""Status"" , AA.""Response"" "
                    End If
                    oGrid.DataTable.ExecuteQuery(Query)
                    oEditTextCol = oGrid.Columns.Item("Code")
                    oEditTextCol.LinkedObjectType = 10000199
                Case "DIV"
                    Query = " Select AA.*  From (" & _
                          " Select ""ITMSGRPCOD"" As ""Code"" , 'Success' As ""Status"" , ""ERRORMSG"" As  ""Response""   From ""10X_ITEMGROUPS"" Where IFNULL(""RESPONSE"",'') <> '' " & _
                          " And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "'  Union ALL " & _
                          " Select ""ITMSGRPCOD"" As ""Code"" , 'Fail' As ""Status"" , ""ERRORMSG"" As  ""Response""  From ""10X_ITEMGROUPS"" Where IFNULL(""RESPONSE"",'') = '' " & _
                          "  And ""ERRORMSG""  <>'' And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "') AA  "
                    If Status = "All" Then
                        Query = Query & "  Group By AA.""Code"" , AA.""Status"" , AA.""Response"" "
                    Else
                        Query = Query & " Where AA.""Status"" = '" & Status & "'  Group By AA.""Code"" , AA.""Status"" , AA.""Response"" "
                    End If
                    oGrid.DataTable.ExecuteQuery(Query)
                    oEditTextCol = oGrid.Columns.Item("Code")
                    oEditTextCol.LinkedObjectType = 52
                Case "CAT"
                    Query = " Select AA.*  From (" & _
                           " Select  ""Code"" , 'Success' As ""Status"" , ""ERRORMSG"" As  ""Response""   From ""10X_ITEMCATEGORY"" Where IFNULL(""RESPONSE"",'') <> '' " & _
                           " And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "'  Union ALL " & _
                           " Select  ""Code"" , 'Fail' As ""Status"" , ""ERRORMSG"" As  ""Response""  From ""10X_ITEMCATEGORY"" Where IFNULL(""RESPONSE"",'') = '' " & _
                           "  And ""ERRORMSG""  <>'' And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "') AA  "
                    If Status = "All" Then
                        Query = Query & "  Group By AA.""Code"" , AA.""Status"" , AA.""Response"" "
                    Else
                        Query = Query & " Where AA.""Status"" = '" & Status & "'  Group By AA.""Code"" , AA.""Status"" , AA.""Response"" "
                    End If
                    oGrid.DataTable.ExecuteQuery(Query)
                Case "RPG"
                    Query = " Select AA.*  From (" & _
                            " Select ""CODE"" As ""Code"" , 'Success' As ""Status"" , ""ERRORMSG"" As  ""Response""   From ""10X_RETAILITMGRP"" Where IFNULL(""RESPONSE"",'') <> '' " & _
                            " And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "'  Union ALL " & _
                            " Select ""CODE"" As ""Code"" , 'Fail' As ""Status"" , ""ERRORMSG"" As  ""Response""  From ""10X_RETAILITMGRP"" Where IFNULL(""RESPONSE"",'') = '' " & _
                            "  And ""ERRORMSG""  <>'' And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "') AA  "
                    If Status = "All" Then
                        Query = Query & "  Group By AA.""Code"" , AA.""Status"" , AA.""Response""  "
                    Else
                        Query = Query & " Where  AA.""Status"" = '" & Status & "'  Group By AA.""Code"" , AA.""Status"" , AA.""Response""  "
                    End If
                    oGrid.DataTable.ExecuteQuery(Query)
                Case "ISG"
                    Query = " Select AA.*  From (" & _
                           " Select ""CODE"" As ""Code"" , 'Success' As ""Status"" , ""ERRORMSG"" As  ""Response""   From ""10X_ITEMSPECIALGROUP"" Where IFNULL(""RESPONSE"",'') <> '' " & _
                           " And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "'  Union ALL " & _
                           " Select ""CODE"" As ""Code"" , 'Fail' As ""Status"" , ""ERRORMSG"" As  ""Response""  From ""10X_ITEMSPECIALGROUP"" Where IFNULL(""RESPONSE"",'') = '' " & _
                           "  And ""ERRORMSG""  <>'' And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "') AA  "
                    If Status = "All" Then
                        Query = Query & "  Group By AA.""Code"" , AA.""Status"" , AA.""Response"" "
                    Else
                        Query = Query & " Where AA.""Status"" = '" & Status & "'  Group By AA.""Code"" , AA.""Status"" , AA.""Response""  "
                    End If
                    oGrid.DataTable.ExecuteQuery(Query)
                Case "BRD"
                    Query = " Select AA.*  From (" & _
                           " Select ""CODE"" As ""Code"" , 'Success' As ""Status"" , ""ERRORMSG"" As  ""Response""   From ""10X_BRAND"" Where IFNULL(""RESPONSE"",'') <> '' " & _
                           " And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "'  Union ALL " & _
                           " Select ""CODE"" As ""Code"" , 'Fail' As ""Status"" , ""ERRORMSG"" As  ""Response""  From ""10X_BRAND"" Where IFNULL(""RESPONSE"",'') = '' " & _
                           "  And ""ERRORMSG""  <>'' And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "') AA  "
                    If Status = "All" Then
                        Query = Query & "  Group By AA.""Code"" , AA.""Status"" , AA.""Response"" "
                    Else
                        Query = Query & " Where  AA.""Status"" = '" & Status & "'  Group By AA.""Code"" , AA.""Status"" , AA.""Response""  "
                    End If
                    oGrid.DataTable.ExecuteQuery(Query)
                    'Case "SLP"
                    '    Query = " Select AA.*  From (Select ""ITEMCODE"" As ""Code"" , ""UOMCODE"" As ""UomCode"" , ""PRICE"" As ""Price"" ,'BaseUOMPrice' As ""Type"" , 'Success' As ""Status"" , ""ERRORMSG"" As  ""Response""   From ""10X_BASEUOMPRICE"" Where IFNULL(""RESPONSE"",'') <> '' " & _
                    '            " And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "' Union ALL " & _
                    '            " Select ""ITEMCODE"" As ""Code"" , ""UOMCODE"" As ""UomCode"" , ""PRICE"" As ""Price"" , 'BaseUOMPrice' As ""Type"" ,  'Fail' As ""Status"" , ""ERRORMSG"" As  ""Response""  From ""10X_BASEUOMPRICE"" Where IFNULL(""RESPONSE"",'') = '' " & _
                    '            " And ""ERRORMSG""  <>'' And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "'  And '" & objForm.Items.Item("3").Specific.Value & "' Union ALL " & _
                    '            " Select ""ITEMCODE"" As ""Code"" , ""UOMCODE"" As ""UomCode"" , ""PRICE"" As ""Price"" ,'OtherUOMPrice' As ""Type"" , 'Success' As ""Status"" , ""ERRORMSG"" As  ""Response""   From ""10X_OTHERUOMPRICE"" Where IFNULL(""RESPONSE"",'') <> '' " & _
                    '            " And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "' Union ALL " & _
                    '            " Select ""ITEMCODE"" As ""Code"" , ""UOMCODE"" As ""UomCode"" , ""PRICE"" As ""Price"" , 'OtherUOMPrice' As ""Type"" ,  'Fail' As ""Status"" , ""ERRORMSG"" As  ""Response""  From ""10X_OTHERUOMPRICE"" Where IFNULL(""RESPONSE"",'') = '' " & _
                    '            " And ""ERRORMSG""  <>'' And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("3").Specific.Value & "') AA  "
                    '    If Status = "All" Then
                    '        Query = Query & "  Group By AA.""Code"" , AA.""UomCode"" , AA.""Price"" , AA.""Status"" , AA.""Response"" , AA.""Type"" "
                    '    Else
                    '        Query = Query & " Where AA.""Status"" = '" & Status & "'   Group By AA.""Code"" , AA.""UomCode"" , AA.""Price"" , AA.""Status"" , AA.""Response"" , AA.""Type"" "
                    '    End If
                    '    oGrid.DataTable.ExecuteQuery(Query)

                    '    oGrid.Columns.Item("UomCode").Editable = False
                    '    oGrid.Columns.Item("Price").Editable = False
                    '    oGrid.Columns.Item("Type").Editable = False

                    '    oEditTextCol = oGrid.Columns.Item("Code")
                    '    oEditTextCol.LinkedObjectType = 2
                Case "BSP"
                    Query = " Select AA.*  From (Select ""ITEMCODE"" As ""Code"" , ""UOMCODE"" As ""UomCode"" , ""PRICE"" As ""Price"" ,'BaseUOMPrice' As ""Type"" , 'Success' As ""Status"" , ""ERRORMSG"" As  ""Response""   From ""10X_BASEUOMPRICE"" Where IFNULL(""RESPONSE"",'') <> '' " & _
                            " And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "' Union ALL " & _
                            " Select ""ITEMCODE"" As ""Code"" , ""UOMCODE"" As ""UomCode"" , ""PRICE"" As ""Price"" , 'BaseUOMPrice' As ""Type"" ,  'Fail' As ""Status"" , ""ERRORMSG"" As  ""Response""  From ""10X_BASEUOMPRICE"" Where IFNULL(""RESPONSE"",'') = '' " & _
                            " And ""ERRORMSG""  <>'' And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "'  And '" & objForm.Items.Item("3").Specific.Value & "') AA  "
                    If Status = "All" Then
                        Query = Query & "  Group By AA.""Code"" , AA.""UomCode"" , AA.""Price"" , AA.""Status"" , AA.""Response"" , AA.""Type"" "
                    Else
                        Query = Query & " Where AA.""Status"" = '" & Status & "'   Group By AA.""Code"" , AA.""UomCode"" , AA.""Price"" , AA.""Status"" , AA.""Response"" , AA.""Type"" "
                    End If
                    oGrid.DataTable.ExecuteQuery(Query)

                    oGrid.Columns.Item("UomCode").Editable = False
                    oGrid.Columns.Item("Price").Editable = False
                    oGrid.Columns.Item("Type").Editable = False

                    oEditTextCol = oGrid.Columns.Item("Code")
                    oEditTextCol.LinkedObjectType = 2
                Case "OTP"
                    Query = " Select AA.*  From ( Select ""ITEMCODE"" As ""Code"" , " & _
                            " ""UOMCODE"" As ""UomCode"" , ""PRICE"" As ""Price"" ,'OtherUOMPrice' As ""Type"" , 'Success' As ""Status"" , ""ERRORMSG"" As  ""Response""   From ""10X_OTHERUOMPRICE"" Where IFNULL(""RESPONSE"",'') <> '' " & _
                            " And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "' Union ALL " & _
                            " Select ""ITEMCODE"" As ""Code"" , ""UOMCODE"" As ""UomCode"" , ""PRICE"" As ""Price"" , 'OtherUOMPrice' As ""Type"" ,  'Fail' As ""Status"" , ""ERRORMSG"" As  ""Response""  From ""10X_OTHERUOMPRICE"" Where IFNULL(""RESPONSE"",'') = '' " & _
                            " And ""ERRORMSG""  <>'' And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("3").Specific.Value & "') AA  "
                    If Status = "All" Then
                        Query = Query & "  Group By AA.""Code"" , AA.""UomCode"" , AA.""Price"" , AA.""Status"" , AA.""Response"" , AA.""Type"" "
                    Else
                        Query = Query & " Where AA.""Status"" = '" & Status & "'   Group By AA.""Code"" , AA.""UomCode"" , AA.""Price"" , AA.""Status"" , AA.""Response"" , AA.""Type"" "
                    End If
                    oGrid.DataTable.ExecuteQuery(Query)

                    oGrid.Columns.Item("UomCode").Editable = False
                    oGrid.Columns.Item("Price").Editable = False
                    oGrid.Columns.Item("Type").Editable = False

                    oEditTextCol = oGrid.Columns.Item("Code")
                    oEditTextCol.LinkedObjectType = 2
                Case "BAR"
                    Query = " Select AA.*  From (" & _
                            " Select ""BARCODE"" As ""Code"" , ""ITEMCODE"" As ""ItemCode"" , 'Success' As ""Status"" , ""ERRORMSG"" As  ""Response""   From ""10X_BARCODE"" Where IFNULL(""RESPONSE"",'') <> '' " & _
                            " And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "'  Union ALL " & _
                            " Select ""BARCODE"" As ""Code"" , ""ITEMCODE"" As ""ItemCode""  , 'Fail' As ""Status"" , ""ERRORMSG"" As  ""Response""  From ""10X_BARCODE"" Where IFNULL(""RESPONSE"",'') = '' " & _
                            "  And ""ERRORMSG""  <>'' And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("3").Specific.Value & "')AA   "
                    If Status = "All" Then
                        Query = Query & "  Group By AA.""Code"" , AA.""ItemCode"" , AA.""Status"" , AA.""Response"" "
                    Else
                        Query = Query & " Where AA.""Status"" = '" & Status & "'  Group By AA.""Code"" , AA.""ItemCode"" , AA.""Status"" , AA.""Response""  "
                    End If
                    oGrid.DataTable.ExecuteQuery(Query)
                    oEditTextCol = oGrid.Columns.Item("ItemCode")
                    oEditTextCol.LinkedObjectType = 4
                    oGrid.Columns.Item("ItemCode").Editable = False
                Case "UMG"
                    Query = " Select AA.*  From (" & _
                             " Select ""UGPCODE"" As ""Code"" , ""UOMCODE"" As ""UomCode"" , 'Success' As ""Status"" , ""ERRORMSG"" As  ""Response""   From ""10X_ITEMUOM"" Where IFNULL(""RESPONSE"",'') <> '' " & _
                             " And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "'  Union ALL " & _
                             " Select ""UGPCODE"" As ""Code"" , ""UOMCODE"" As ""UomCode""  , 'Fail' As ""Status"" , ""ERRORMSG"" As  ""Response""  From ""10X_ITEMUOM"" Where IFNULL(""RESPONSE"",'') = '' " & _
                             "  And ""ERRORMSG""  <>'' And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("3").Specific.Value & "')AA  "
                    If Status = "All" Then
                        Query = Query & " Group By AA.""Code"" , AA.""UomCode"" , AA.""Status"" , AA.""Response"" "
                    Else
                        Query = Query & " Where  AA.""Status"" = '" & Status & "'   Group By AA.""Code"" , AA.""UomCode"" , AA.""Status"" , AA.""Response""  "
                    End If
                    oGrid.DataTable.ExecuteQuery(Query)
                    oEditTextCol = oGrid.Columns.Item("Code")
                    oEditTextCol.LinkedObjectType = 10000197
                    oGrid.Columns.Item("UomCode").Editable = False
                Case "CST"
                    Query = " Select AA.*  From (" & _
                           " Select ""CARDCODE"" As ""Code"" , 'Success' As ""Status"" , ""ERRORMSG"" As  ""Response""   From ""10X_CUSTOMER"" Where IFNULL(""RESPONSE"",'') <> '' " & _
                           " And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "'  Union ALL " & _
                           " Select ""CARDCODE"" As ""Code"" , 'Fail' As ""Status"" , ""ERRORMSG"" As  ""Response""  From ""10X_CUSTOMER"" Where IFNULL(""RESPONSE"",'') = '' " & _
                           "  And ""ERRORMSG""  <>'' And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "') AA  "
                    If Status = "All" Then
                        Query = Query & "  Group By AA.""Code"" , AA.""Status"" , AA.""Response"" "
                    Else
                        Query = Query & " Where  AA.""Status"" = '" & Status & "'    Group By AA.""Code"" , AA.""Status"" , AA.""Response""  "
                    End If
                    oGrid.DataTable.ExecuteQuery(Query)
                    oEditTextCol = oGrid.Columns.Item("Code")
                    oEditTextCol.LinkedObjectType = 2
                Case "SEP"
                    Query = " Select AA.*  From (" & _
                          " Select ""EMPID"" As ""Code"" , 'Success' As ""Status"" , ""ERRORMSG"" As  ""Response""   From ""10X_SALESEMPLOYEE"" Where IFNULL(""RESPONSE"",'') <> '' " & _
                          " And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "'  Union ALL " & _
                          " Select ""EMPID"" As ""Code"" , 'Fail' As ""Status"" , ""ERRORMSG"" As  ""Response""  From ""10X_SALESEMPLOYEE"" Where IFNULL(""RESPONSE"",'') = '' " & _
                          "  And ""ERRORMSG""  <>'' And ""DATE"" Between '" & objForm.Items.Item("3").Specific.Value & "' And '" & objForm.Items.Item("5").Specific.Value & "') AA  "
                    If Status = "All" Then
                        Query = Query & "  Group By AA.""Code"" , AA.""Status"" , AA.""Response"" "
                    Else
                        Query = Query & " Where AA.""Status"" = '" & Status & "'  Group By AA.""Code"" , AA.""Status"" , AA.""Response"" "
                    End If
                    oGrid.DataTable.ExecuteQuery(Query)
                    oEditTextCol = oGrid.Columns.Item("Code")
                    oEditTextCol.LinkedObjectType = 171
            End Select

            If oGrid.DataTable.Columns.Item(0).Cells.Item(0).Value = "" Then
                objForm.Freeze(False)
                objForm.DataSources.UserDataSources.Item("U_NOR").Value = 0
                objForm.DataSources.UserDataSources.Item("U_NOS").Value = 0
                objForm.DataSources.UserDataSources.Item("U_NOF").Value = 0
                objForm.Items.Item("15").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                objMain.objApplication.MessageBox("No records found")
                Exit Sub
            End If

            oGrid.Columns.Item("Code").Editable = False
            oGrid.Columns.Item("Status").Editable = False
            oGrid.Columns.Item("Response").Editable = False

            If Query <> "" Then
                Dim GetSucessCount As String = "Select Count(*) From (" & Query & ") Where ""Status"" = 'Success' "
                Dim oRsGetSucessCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetSucessCount.DoQuery(GetSucessCount)
                objForm.DataSources.UserDataSources.Item("U_NOS").Value = oRsGetSucessCount.Fields.Item(0).Value

                Dim GetFailCount As String = "Select Count(*) From (" & Query & ") Where ""Status"" = 'Fail' "
                Dim oRsGetFailCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetFailCount.DoQuery(GetFailCount)
                objForm.DataSources.UserDataSources.Item("U_NOF").Value = oRsGetFailCount.Fields.Item(0).Value
            End If

            oGrid.AutoResizeColumns()
            objForm.Items.Item("15").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            If oGrid.DataTable.Columns.Item(0).Cells.Item(0).Value = "" Then
                objForm.DataSources.UserDataSources.Item("U_NOR").Value = 0
            Else
                objForm.DataSources.UserDataSources.Item("U_NOR").Value = oGrid.Rows.Count
            End If

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objForm.Items.Item("15").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        Finally
            objMain.objApplication.StatusBar.SetText("Operation completed Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End Try
    End Sub

    Public Sub GetTaksDetails()
        Dim taskService As ITaskService = New TaskScheduler.TaskScheduler()
        ' Connect to the TaskService
        'taskService.Connect("192.168.100.120", "Administrator", "DESKTOP-AJGNQ2S", "Net@12345")
        taskService.Connect()
        ' Get the task folder that contains the tasks. 
        Dim rootFolder
        rootFolder = taskService.GetFolder("\")
        Dim taskCollection
        taskCollection = rootFolder.GetTasks(0)
        Dim numberOfTasks
        numberOfTasks = taskCollection.Count
        If numberOfTasks = 0 Then          
        Else           
            Dim registeredTask
            For Each registeredTask In taskCollection
                ' WScript.Echo "Task Name: " & registeredTask.Name
                If registeredTask.Name = "10XSAPDataConnector" Then
                    Dim taskState = ""
                    Select Case registeredTask.State
                        Case "0"
                            taskState = "Unknown"
                        Case "1"
                            taskState = "Disabled"
                        Case "2"
                            taskState = "Queued"
                        Case "3"
                            taskState = "Ready"
                        Case "4"
                            taskState = "Running"
                    End Select
                    objForm.Items.Item("21").Specific.Value = taskState
                    objForm.Items.Item("19").Specific.Value = registeredTask.LastRunTime
                    objForm.Items.Item("22").Specific.Value = registeredTask.NextRunTime
                    Exit For
                End If
            Next
        End If
        taskService = Nothing
        GC.Collect() ' Start .NET CLR Garbage Collection
        GC.WaitForPendingFinalizers() ' Wait for Garbage Collection to finish
    End Sub

End Class
