Public Class WorkOrderList
    Public frmWorkOrderList As SAPbouiCOM.Form
    Dim ogrid As SAPbouiCOM.Grid

    Sub LoadForm()
        Try
            oGFun.LoadXML(frmWorkOrderList, WorkOrderListFormId, WorkOrderListXML)
            frmWorkOrderList = oApplication.Forms.Item(WorkOrderListFormId)
            frmWorkOrderList.DataSources.DataTables.Add("Datatable")
            Me.InitForm()
        Catch ex As Exception
            oApplication.StatusBar.SetText("LoadForm Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub

    Sub InitForm()
        Try
            frmWorkOrderList.Freeze(True)
            Dim status As SAPbouiCOM.ComboBox
            Dim Orders As SAPbouiCOM.ComboBox
            Dim Location As SAPbouiCOM.ComboBox
            status = frmWorkOrderList.Items.Item("c_Status").Specific
            status.ValidValues.Add("-", "")
            status.ValidValues.Add("O", "Open")
            status.ValidValues.Add("C", "Close")
            Orders = frmWorkOrderList.Items.Item("c_Orders").Specific
            Orders.ValidValues.Add("-", "")
            Orders.ValidValues.Add("1", "Stock Order")
            Orders.ValidValues.Add("2", "Sales Order")
            Orders.ValidValues.Add("3", "Sales Return")
            Orders.ValidValues.Add("4", "WIP Update")
            'For Location
            Location = frmWorkOrderList.Items.Item("c_Location").Specific
            oGFun.setComboBoxValue(Location, "Select Code,Name from [@PROLO]")
            Location.ValidValues.Add("-", "")

            Orders.Select("-", SAPbouiCOM.BoSearchKey.psk_ByValue)
            status.Select("-", SAPbouiCOM.BoSearchKey.psk_ByValue)
            Location.Select("-", SAPbouiCOM.BoSearchKey.psk_ByValue)
            'Get WoNo
            Dim WoNo = frmWorkOrderList.Items.Item("t_WoNo").Specific.Value
            Dim CardCode = frmWorkOrderList.Items.Item("t_CardCode").Specific.Value

            'Load the Work Orders
            LoadGrid(status.Selected.Value, Orders.Selected.Value, Location.Selected.Value, WoNo, CardCode)

        Catch ex As Exception
            oApplication.StatusBar.SetText("InitForm Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
            frmWorkOrderList.Freeze(False)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    Try
                         Dim oDataTable As SAPbouiCOM.DataTable
                        Dim oCFLE As SAPbouiCOM.ChooseFromListEvent = pVal
                        oDataTable = oCFLE.SelectedObjects
                        Select Case oCFLE.ChooseFromListUID
                            Case "CFL_2A"
                                Try
                                    If pVal.BeforeAction Then
                                        oGFun.ChooseFromListFilteration(frmWorkOrderList, "CFL_2A", "CardType", "Select 'C'")
                                    End If
                                    If pVal.BeforeAction = False Then
                                        frmWorkOrderList.Items.Item("t_CardCode").Specific.Value = oDataTable.GetValue("CardCode", 0)
                                    End If
                                Catch ex As Exception
                                End Try
                        End Select
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("ChooseFrom List Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try

                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    Try
                        'If (pVal.ItemUID = "c_Status" Or pVal.ItemUID = "c_Orders") And pVal.BeforeAction = False Then
                        '    Dim oCmb As SAPbouiCOM.ComboBox = frmWorkOrderList.Items.Item("c_Status").Specific
                        '    Dim ocmb_Orders As SAPbouiCOM.ComboBox = frmWorkOrderList.Items.Item("c_Orders").Specific
                        '    ' LoadGrid(oCmb.Selected.Value, ocmb_Orders.Selected.Value)
                        'End If
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Combo Select Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    Try
                        Select Case pVal.ItemUID
                            Case "Grid"
                                If pVal.BeforeAction Then
                                    BubbleEvent = False
                                    ogrid = frmWorkOrderList.Items.Item("Grid").Specific
                                    For i As Integer = 1 To ogrid.Rows.Count
                                        If ogrid.Rows.IsSelected(i - 1) Then
                                            Dim WorkType = oGFun.getSingleValue("select isnull(U_WorkType,2) from [@MIPL_OWKO] where DocNum = '" & ogrid.DataTable.GetValue("DocNum", i - 1) & "'")
                                            If WorkType.Trim = "1" Then
                                                oWorkOrder.LoadForm(ogrid.DataTable.GetValue("DocNum", i - 1))
                                            Else
                                                oWorkOrder.LoadForm(ogrid.DataTable.GetValue("DocNum", i - 1))
                                            End If
                                            Return
                                        End If
                                    Next
                                End If
                                If pVal.BeforeAction Then
                                    ogrid.Columns.Item(pVal.ColUID).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
                                End If

                                'Case "b_Find"
                                '    If pVal.BeforeAction = False Then
                                '        Dim oCmb As SAPbouiCOM.ComboBox = frmWorkOrderList.Items.Item("c_Status").Specific
                                '        Dim ocmb_Orders As SAPbouiCOM.ComboBox = frmWorkOrderList.Items.Item("c_Orders").Specific
                                '        Dim oCmb_Loc As SAPbouiCOM.ComboBox = frmWorkOrderList.Items.Item("c_Location").Specific
                                '        Dim WoNo As String = frmWorkOrderList.Items.Item("t_WoNo").Specific.Value
                                '        Dim CardCode As String = frmWorkOrderList.Items.Item("t_CardCode").Specific.Value
                                '        'Load the Values
                                '        Me.LoadGrid(oCmb.Selected.Value, ocmb_Orders.Selected.Value, oCmb_Loc.Selected.Value, WoNo, CardCode)
                                '    End If
                        End Select
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Double Click Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Finally
                    End Try

                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    Try
                        Select Case pVal.ItemUID
                            Case "b_Ok"
                                If pVal.ActionSuccess Then
                                    frmWorkOrderList.Close()
                                End If
                            Case "b_Find"
                                If pVal.BeforeAction = False Then
                                    Dim oCmb As SAPbouiCOM.ComboBox = frmWorkOrderList.Items.Item("c_Status").Specific
                                    Dim ocmb_Orders As SAPbouiCOM.ComboBox = frmWorkOrderList.Items.Item("c_Orders").Specific
                                    Dim oCmb_Loc As SAPbouiCOM.ComboBox = frmWorkOrderList.Items.Item("c_Location").Specific
                                    Dim WoNo As String = frmWorkOrderList.Items.Item("t_WoNo").Specific.Value
                                    Dim CardCode As String = frmWorkOrderList.Items.Item("t_CardCode").Specific.Value
                                    'Load the Values
                                    Me.LoadGrid(oCmb.Selected.Value, ocmb_Orders.Selected.Value, oCmb_Loc.Selected.Value, WoNo, CardCode)
                                End If
                        End Select
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Click Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText("Item Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub LoadGrid(ByVal Status As String, ByVal Orders As String, ByVal Location As String, ByVal WoNo As String, ByVal CardCode As String)
        Try
            Dim condition As String = ""

            If Status.Trim <> "-" Then
                condition = condition & " and Status ='" & Status.Trim & "'"
            End If

            If Orders.Trim <> "-" Then
                condition = condition & " and isnull(U_WorkType,'') = '" & Orders.Trim & "' "
            End If

            If Location.Trim <> "-" Then
                condition = condition & " and isnull(U_Location,'') ='" & Location.Trim & "'"
            End If

            If WoNo.Trim <> "" Then
                condition = condition & " and DocNum ='" & WoNo.Trim & "'"
            End If

            If CardCode.Trim <> "" Then
                condition = condition & " and U_CardCode ='" & CardCode.Trim & "'"
            End If

            Dim str As String
            str = "Select DocNum,U_SONo SONo,U_CardCode CustomerCode,U_CardName CustomerName,U_Location Location, " & _
            "U_WORegNo WORegNo,U_IWODate IWDate,U_CustPoNo CustPoNo,U_RevNo RevNo,U_RegDate RegDate,U_IWORev IWORev," & _
            "U_IWORevDt IWORevDt,U_PODate PODate,case when Status ='O' then 'Open' when Status ='C' then 'Close' End Status from [@MIPL_OWKO] where 1=1 and status<>'N' and canceled<>'Y'" & _
            condition

            ogrid = frmWorkOrderList.Items.Item("Grid").Specific
            ogrid.DataTable = frmWorkOrderList.DataSources.DataTables.Item("Datatable")
            ogrid.DataTable.ExecuteQuery(str)
            For i As Integer = 0 To ogrid.Columns.Count - 1
                ogrid.Columns.Item(i).TitleObject.Sortable = True
            Next
        Catch ex As Exception
            oApplication.StatusBar.SetText("Load Grid Function failed" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case "1282"
                    Me.InitForm()
            End Select
        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText("Menu Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

End Class
