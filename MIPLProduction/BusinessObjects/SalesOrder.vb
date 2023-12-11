Public Class SalesOrder

    Public frmSalesOrder, frmACC As SAPbouiCOM.Form
    Dim oDBDSHeader, oDBDSDetail, oDBDSDetail_ACC As SAPbouiCOM.DBDataSource
    Public oMatrix As SAPbouiCOM.Matrix
    Dim DataTable_ACC As New DataTable
    Public SalesOrderId As String = ""
    Dim INVNo As String = ""
    Dim EroVal As String = ""
    Dim ForEx As String = ""
    Dim BankRef As String = ""
    Dim SODOCENTRY As String = ""
    Dim CurrentRow As Integer = 0
    Dim ItemCode As String = ""
    Sub LoadForm()
        Try
            frmSalesOrder = oApplication.Forms.ActiveForm
            SalesOrderId = frmSalesOrder.UniqueID
            'Assign Data Source
            oDBDSHeader = frmSalesOrder.DataSources.DBDataSources.Item(0)
            oDBDSDetail = frmSalesOrder.DataSources.DBDataSources.Item(1)
            oMatrix = frmSalesOrder.Items.Item("38").Specific
            Me.InitForm()
            Me.DefineModesForFields()
        Catch ex As Exception
            oGFun.Msg("Load Form Method Failed:" & ex.Message)
        End Try
    End Sub

    Sub InitForm()
        Try
            frmSalesOrder.Freeze(True)

            'Incoming Payments
            DataTable_ACC.Columns.Clear()
            DataTable_ACC.Columns.Add("ParentCode")
            DataTable_ACC.Columns.Add("PLineId")
            DataTable_ACC.Columns.Add("AccCode")
            DataTable_ACC.Columns.Add("AccName")
            DataTable_ACC.Columns.Add("CrAmt")
            DataTable_ACC.Columns.Add("DrAmt")

            DataTable_ACC.Columns.Add("INVNo")
            DataTable_ACC.Columns.Add("EroVal")
            DataTable_ACC.Columns.Add("ForEx")
            DataTable_ACC.Columns.Add("BankRef")

        Catch ex As Exception
            oApplication.StatusBar.SetText("InitForm Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
            frmSalesOrder.Freeze(False)

        End Try
    End Sub

    Sub DefineModesForFields()
        Try
        Catch ex As Exception
            oApplication.StatusBar.SetText("DefineModesForFields Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Function ValidateAll() As Boolean
        Try
            ValidateAll = True
        Catch ex As Exception
            oApplication.StatusBar.SetText("Validate Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            ValidateAll = False
        Finally
        End Try

    End Function

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    Try
                        'Assign Selected Rows
                        Dim oDataTable As SAPbouiCOM.DataTable
                        Dim oCFLE As SAPbouiCOM.ChooseFromListEvent = pVal
                        oDataTable = oCFLE.SelectedObjects
                        Dim rset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'Filter before open the CFL
                        If pVal.BeforeAction Then
                            'Select Case oCFLE.ChooseFromListUID
                            '    Case "1"
                            'End Select
                        End If
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Choose From List Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Finally
                    End Try
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    Try
                        Select Case pVal.ItemUID
                            Case "1"
                                'Add,Update Event
                                If pVal.BeforeAction = True And (frmSalesOrder.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or frmSalesOrder.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If Me.ValidateAll() = False Then
                                        System.Media.SystemSounds.Asterisk.Play()
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                    End If
                                End If
                        End Select
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Click Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        BubbleEvent = False
                    Finally
                    End Try
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText("Item Event Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID

                Case "1282"
                    Me.InitForm()
                Case "1293"
                    '  oGFun.DeleteRow(oMatrix, oDBDSDetail)
                Case "1284"
                    If pVal.BeforeAction = False Then

                        Dim strsql1 As String = oGFun.getSingleValue("select owko.docentry from  [@mipl_OWKO] owko join [@mipl_WKO1] wko1 on owko.docentry=wko1.docentry join ordr on ordr.docentry=wko1.U_baseentry and wko1.U_baseentry='" & SODOCENTRY.Trim & "'")
                        strsql1 = IIf(strsql1 = "", "", strsql1)
                        Dim strsql As String = "update [@mipl_OWKO]  set status='N',canceled='Y' where  docentry='" & strsql1.Trim & "'"
                        Dim rs_Loc As SAPbobsCOM.Recordset = oGFun.DoQuery(strsql)
                    End If
            End Select

            If pVal.MenuUID.Contains(WorkOrderSalesFormId) And pVal.BeforeAction = True Then
                Dim oMenu As SAPbouiCOM.MenuItem = oApplication.Menus.Item(pVal.MenuUID)
                Dim Location() As String = oMenu.String.Split("-")
                If Location.Length >= 2 Then
                    oWorkOrder.LocationCode = oGFun.getSingleValue(" Select Code from [@PROLO] where Name ='" & Location(1) & "'")
                    oWorkOrder.LocationDesc = Location(1)

                    oWorkOrder.ParentCode = frmSalesOrder.Items.Item("8").Specific.Value
                    oWorkOrder.CardCode = frmSalesOrder.Items.Item("4").Specific.Value
                    oWorkOrder.CardName = frmSalesOrder.Items.Item("54").Specific.Value
                    oWorkOrder.ItemCode = ItemCode

                    oWorkOrder.LoadForm(CurrentRow)
                End If
            End If

        Catch ex As Exception
            oApplication.StatusBar.SetText("Menu Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
             Select eventInfo.ItemUID
                Case "38"
                    If frmSalesOrder.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or frmSalesOrder.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or frmSalesOrder.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then


                        If eventInfo.Row <> oMatrix.VisualRowCount And eventInfo.BeforeAction And oMatrix.GetCellSpecific("U_PROLO", eventInfo.Row).Value.ToString.Trim <> "" Then

                            Dim oCmb As SAPbouiCOM.ComboBox = oMatrix.GetCellSpecific("U_PROLO", eventInfo.Row)
                            oGFun.setRightMenu(WorkOrderSalesFormId, "Work Order-" + oCmb.Selected.Description.ToString.Trim)
                            CurrentRow = oMatrix.GetCellSpecific("0", eventInfo.Row).Value
                            ItemCode = oMatrix.GetCellSpecific("1", eventInfo.Row).Value
                        End If
                    End If
            End Select
         
        Catch ex As Exception
            oApplication.StatusBar.SetText("Right Click Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                    Try
                        If BusinessObjectInfo.BeforeAction Then
                            If Me.ValidateAll() = False Then
                                System.Media.SystemSounds.Asterisk.Play()
                                BubbleEvent = False
                                Exit Sub
                            Else

                            End If
                        End If
                        If BusinessObjectInfo.ActionSuccess Then
                            DataTable_ACC.Clear()
                        End If
                    Catch ex As Exception
                        BubbleEvent = False
                    Finally
                    End Try

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.ActionSuccess = True Then
                    End If
            End Select
        Catch ex As Exception
            If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            oApplication.StatusBar.SetText("Form Data Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

End Class
