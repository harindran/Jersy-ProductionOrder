Public Class ProductionOrder

    Public frmProductionOrder, frmPRM, frmPMC, frmPTL, frmPSC, frmPCO, frmPLR, frmWO, FrmPID As SAPbouiCOM.Form
    Dim oDBDSHeader, oDBDSDetail As SAPbouiCOM.DBDataSource
    Dim oMatrix As SAPbouiCOM.Matrix

    Dim ParentCode As String = ""
    Dim CurrentRow As Integer = 0
    Dim FormExist = False

    Dim Instruction As String = ""
    Dim ParCode
    Dim docnum = ""
    Dim FGItemCode = ""
    Dim InQty As Double = 0
    Dim PreWODocEntry = ""
    Sub LoadForm()
        Try
            frmProductionOrder = objaddon.objapplication.Forms.ActiveForm
            'Assign Data Source
            oDBDSHeader = frmProductionOrder.DataSources.DBDataSources.Item(0)
            oDBDSDetail = frmProductionOrder.DataSources.DBDataSources.Item(1)
            oMatrix = frmProductionOrder.Items.Item("37").Specific
            Me.InitForm()
            Me.DefineModesForFields()
        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage("Load Form Method Failed:" & ex.Message)
         ()
        End Try
    End Sub

    Sub InitForm()
        Try
            frmProductionOrder.Freeze(True)




            Dim oItem As SAPbouiCOM.Item
            Dim oLabel As SAPbouiCOM.StaticText
            Dim oEditText As SAPbouiCOM.EditText
            Dim oButton As SAPbouiCOM.Button
            Dim oComboBox As SAPbouiCOM.ComboBox

            'Production No.

            oItem = frmProductionOrder.Items.Add("t_WONo", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = frmProductionOrder.Items.Item("78").Left
            oItem.Width = frmProductionOrder.Items.Item("78").Width
            oItem.Height = frmProductionOrder.Items.Item("78").Height
            oItem.Top = frmProductionOrder.Items.Item("78").Top + 15
            oItem.Enabled = False
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "OWOR", "U_WONo")

            oItem = frmProductionOrder.Items.Add("l_WONo", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = frmProductionOrder.Items.Item("77").Left
            oItem.Top = frmProductionOrder.Items.Item("t_WONo").Top
            oItem.Height = frmProductionOrder.Items.Item("77").Height
            oItem.Width = frmProductionOrder.Items.Item("77").Width - 15
            oItem.LinkTo = "t_WONo"
            oLabel = oItem.Specific
            oLabel.Caption = "Work Order No."

            oItem = frmProductionOrder.Items.Add("b_WONo", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = frmProductionOrder.Items.Item("t_WONo").Left + frmProductionOrder.Items.Item("t_WONo").Width + 2
            oItem.Width = 20
            oItem.Height = frmProductionOrder.Items.Item("t_WONo").Height
            oItem.Top = frmProductionOrder.Items.Item("t_WONo").Top
            oButton = oItem.Specific
            oButton.Caption = "||"


            'Status
            oItem = frmProductionOrder.Items.Add("b_Status", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = frmProductionOrder.Items.Item("t_WONo").Left + frmProductionOrder.Items.Item("t_WONo").Width + 2
            oItem.Width = 20
            oItem.Height = frmProductionOrder.Items.Item("t_WONo").Height
            oItem.Top = frmProductionOrder.Items.Item("10000141").Top
            oButton = oItem.Specific
            oButton.Caption = "||"

            'Planned Qty
            oItem = frmProductionOrder.Items.Add("t_ParPONo", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = frmProductionOrder.Items.Item("78").Left
            oItem.Width = frmProductionOrder.Items.Item("t_WONo").Width
            oItem.Height = frmProductionOrder.Items.Item("t_WONo").Height
            oItem.Top = frmProductionOrder.Items.Item("10000141").Top
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "OWOR", "U_ParentPONo")

            oItem = frmProductionOrder.Items.Add("l_ParPONo", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = frmProductionOrder.Items.Item("77").Left
            oItem.Top = frmProductionOrder.Items.Item("10000141").Top
            oItem.Height = frmProductionOrder.Items.Item("t_WONo").Height
            oItem.Width = frmProductionOrder.Items.Item("t_WONo").Width - 15
            oItem.LinkTo = "t_ParPONo"
            oLabel = oItem.Specific
            oLabel.Caption = "Parent PO No"

            'Finished Qty
            oItem = frmProductionOrder.Items.Add("t_FinQty", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = frmProductionOrder.Items.Item("76").Left
            oItem.Width = frmProductionOrder.Items.Item("76").Width
            oItem.Height = frmProductionOrder.Items.Item("76").Height
            oItem.Top = frmProductionOrder.Items.Item("76").Top + 15
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "OWOR", "U_FinQty")

            oItem = frmProductionOrder.Items.Add("l_FinQty", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = frmProductionOrder.Items.Item("75").Left
            oItem.Top = frmProductionOrder.Items.Item("75").Top + 15
            oItem.Height = frmProductionOrder.Items.Item("75").Height
            oItem.Width = frmProductionOrder.Items.Item("75").Width
            oItem.LinkTo = "t_FinQty"
            oLabel = oItem.Specific
            oLabel.Caption = "Fin Qty"


            'Replanning
            'TextBox
            oItem = frmProductionOrder.Items.Add("C_Replan", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oItem.Left = frmProductionOrder.Items.Item("10").Left + frmProductionOrder.Items.Item("10").Width + 10
            oItem.Width = frmProductionOrder.Items.Item("10").Width
            oItem.Height = frmProductionOrder.Items.Item("10").Height
            oItem.Top = frmProductionOrder.Items.Item("10").Top
            oItem.Enabled = True
            oComboBox = oItem.Specific
            oComboBox.DataBind.SetBound(True, "OWOR", "U_Replan")



            'Child PO
            'TextBox
            oItem = frmProductionOrder.Items.Add("C_ChildPO", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oItem.Left = frmProductionOrder.Items.Item("76").Left + 60
            oItem.Width = frmProductionOrder.Items.Item("78").Width
            oItem.Height = frmProductionOrder.Items.Item("78").Height
            oItem.Top = frmProductionOrder.Items.Item("78").Top + 30
            oItem.Enabled = True
            oComboBox = oItem.Specific
            oComboBox.DataBind.SetBound(True, "OWOR", "U_ChildPO")

            'Static Box
            oItem = frmProductionOrder.Items.Add("s_ChildPO", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = frmProductionOrder.Items.Item("75").Left + 20
            oItem.Top = frmProductionOrder.Items.Item("78").Top + 30
            oItem.Height = frmProductionOrder.Items.Item("78").Height
            oItem.Width = frmProductionOrder.Items.Item("78").Width + 50
            oItem.LinkTo = "C_ChildPO"
            oLabel = oItem.Specific
            oLabel.Caption = "Child PO Creation"

            'Part No.
            oItem = frmProductionOrder.Items.Add("t_PartNo", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = frmProductionOrder.Items.Item("6").Left + frmProductionOrder.Items.Item("6").Width + 2
            oItem.Width = frmProductionOrder.Items.Item("6").Width - 5
            oItem.Height = frmProductionOrder.Items.Item("6").Height
            oItem.Top = frmProductionOrder.Items.Item("6").Top
            oItem.Enabled = False
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "OWOR", "U_PartNo")

            'Base Num    WObaseLN
            oItem = frmProductionOrder.Items.Add("t_WObaseLN", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = frmProductionOrder.Items.Item("6").Left + frmProductionOrder.Items.Item("6").Width + 2
            oItem.Width = 0
            oItem.Height = 0
            oItem.Top = -5
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "OWOR", "U_WObaseLN")

            'Base Num    WObaseLN
            oItem = frmProductionOrder.Items.Add("t_BaseNum", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = frmProductionOrder.Items.Item("6").Left + frmProductionOrder.Items.Item("6").Width + 2
            oItem.Width = 0
            oItem.Height = 0
            oItem.Top = -5
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "OWOR", "U_BaseNum")

            'Base Entry
            oItem = frmProductionOrder.Items.Add("t_BaseEntr", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = frmProductionOrder.Items.Item("6").Left + frmProductionOrder.Items.Item("6").Width + 2
            oItem.Width = 0
            oItem.Height = 0
            oItem.Top = -5
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "OWOR", "U_BaseEntry")

            'Base LineId
            oItem = frmProductionOrder.Items.Add("t_BaseLine", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = frmProductionOrder.Items.Item("t_BaseEntr").Left + frmProductionOrder.Items.Item("t_BaseEntr").Width + 2
            oItem.Width = 0
            oItem.Height = 0
            oItem.Top = -5
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "OWOR", "U_BaseLineId")



            'oMatrix.Columns.Item("U_ReqQty").Editable = False

            'oMatrix.Columns.Item("U_AccQty").Editable = False
            'oMatrix.Columns.Item("U_RewQty").Editable = False
            'oMatrix.Columns.Item("U_RejQty").Editable = False
            'oMatrix.Columns.Item("U_InQty").Editable = False
            'oMatrix.Columns.Item("U_OutQty").Editable = False


        Catch ex As Exception

            objaddon.objapplication.StatusBar.SetText("InitForm Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
            frmProductionOrder.Freeze(False)

        End Try
    End Sub

    Sub DefineModesForFields()
        Try

            frmProductionOrder.Items.Item("t_WONo").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            frmProductionOrder.Items.Item("t_WONo").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionOrder.Items.Item("t_WONo").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

        Catch ex As Exception
            objaddon.objapplication.StatusBar.SetText("DefineModesForFields Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Function ValidateAll() As Boolean
        Try



            ValidateAll = True
        Catch ex As Exception
            objaddon.objapplication.StatusBar.SetText("Validate Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
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
                        Dim rset As SAPbobsCOM.Recordset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'Filter before open the CFL
                        If pVal.BeforeAction Then
                            Select Case oCFLE.ChooseFromListUID
                                'Case "1"
                                '    oGFun.ChooseFromListFilteration(frmProductionOrder, "1", "ItemCode", "Select Itemcode from OITM where  ItmsGrpCod = '" & Operations & "'")
                            End Select
                        End If
                        If pVal.BeforeAction = False Then
                        End If
                    Catch ex As Exception
                        objaddon.objapplication.StatusBar.SetText("Choose From List Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Finally
                    End Try
                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                    Try
                        Try

                        Catch ex As Exception

                        End Try
                    Catch ex As Exception
                        objaddon.objapplication.StatusBar.SetText("FORM_CLOSE Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try

                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    Try
                        Select Case pVal.CharPressed
                            Case "15"
                                If pVal.BeforeAction Then
                                    oMatrix.Columns.Item("4").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Double)
                                End If
                        End Select
                    Catch ex As Exception
                        objaddon.objapplication.StatusBar.SetText("ITEM_PRESSED Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    Try
                        Select Case pVal.ItemUID
                            Case "6"
                                Try
                                    Dim n = frmProductionOrder.Items.Item("6").Specific.Value()
                                Catch ex As Exception
                                    Return
                                End Try

                                If pVal.BeforeAction = False And frmProductionOrder.Items.Item("6").Specific.Value.ToString.Trim <> "" Then

                                    frmProductionOrder.Items.Item("t_PartNo").Specific.Value = oGFun.getSingleValue(" Select FrgnName from OITM where ItemCode = '" & frmProductionOrder.Items.Item("6").Specific.Value() & "' ")



                                    Dim rs As SAPbobsCOM.Recordset = oGFun.DoQuery("Select * from ITT1 where father = '" & frmProductionOrder.Items.Item("6").Specific.Value & "' ")

                                    If rs.RecordCount > 0 Then

                                        Dim Qty = frmProductionOrder.Items.Item("12").Specific.Value
                                        Dim WhsCode = frmProductionOrder.Items.Item("78").Specific.Value


                                        For i As Integer = 1 To oMatrix.VisualRowCount - 1

                                            oMatrix.GetCellSpecific("U_SeqNo", i).Value = rs.Fields.Item("U_SeqNo").Value
                                            'oMatrix.GetCellSpecific("U_PType", i).Select(rs.Fields.Item("U_PType").Value.ToString.Trim)

                                            oMatrix.GetCellSpecific("10", i).Value = rs.Fields.Item("WareHouse").Value

                                            rs.MoveNext()

                                        Next
                                    End If


                                End If


                        End Select
                    Catch ex As Exception
                        objaddon.objapplication.StatusBar.SetText("Last Focus Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Finally
                        'oMatrix.Columns.Item("U_ReqQty").Editable = False
                    End Try
                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    Try
                        Select Case pVal.ColUID
                            Case "U_SeqNo"
                                If pVal.BeforeAction Then
                                    Dim strSeqList = ""
                                    For i As Integer = 1 To oMatrix.VisualRowCount - 1

                                        Dim strSeq = oMatrix.Columns.Item("U_SeqNo").Cells.Item(i).Specific.value.ToString.Trim

                                        If strSeq.Trim = "" Then Continue For

                                        If strSeqList.Contains("/\" + strSeq + "/\") = True Then

                                            objAddOn.objApplication.SetStatusBarMessage("Line No." & i & " Sequance No. Should Not Be Duplicate")
                                            BubbleEvent = False
                                            Return

                                        End If
                                        strSeqList = strSeqList + "/\" + strSeq + "/\"
                                    Next
                                End If
                        End Select
                    Catch ex As Exception
                        objaddon.objapplication.StatusBar.SetText("Validate Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Try
                        Select Case pVal.ItemUID
                            Case "b_Status"
                                If pVal.BeforeAction = False And frmProductionOrder.Items.Item("6").Specific.value <> "" Then
                                    Try

                                        objaddon.objapplication.Menus.Item("PAR").Activate()

                                        Dim frm As SAPbouiCOM.Form
                                        frm = objaddon.objapplication.Forms.ActiveForm
                                        frm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                                        frm.Items.Item("7").Specific.Value = oGFun.GetWONO(oDBDSHeader.GetValue("DocNum", 0))
                                        frm.Items.Item("lt_FGPO").Specific.Value = oGFun.GetParentPONO(oDBDSHeader.GetValue("DocNum", 0))
                                        frm.Items.Item("8").Click()


                                    Catch ex As Exception
                                        objaddon.objapplication.StatusBar.SetText(ex.Message)
                                    End Try
                                End If
                            Case "1"
                                If pVal.BeforeAction = False And frmProductionOrder.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    Dim WKONO = oGFun.getSingleValue("select U_BaseNum from OWOR where docentry=(select max(docentry) from OWOR where U_BaseNum<>'')")
                                    If WKONO <> "" Then
                                        Try

                                            objaddon.objapplication.Menus.Item("PAR").Activate()

                                            Dim frm As SAPbouiCOM.Form
                                            frm = objaddon.objapplication.Forms.ActiveForm
                                            frm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                                            ' frm.Items.Item("7").Specific.Value = oDBDSHeader.GetValue("U_BaseNum", 0)
                                            frm.Items.Item("7").Specific.Value = oGFun.GetWONO(CDbl(oDBDSHeader.GetValue("DocNum", 0)) - 1)
                                            frm.Items.Item("lt_FGPO").Specific.Value = oGFun.GetParentPONO(CDbl(oDBDSHeader.GetValue("DocNum", 0)) - 1)
                                            frm.Items.Item("8").Click()


                                        Catch ex As Exception
                                            objaddon.objapplication.StatusBar.SetText(ex.Message)
                                        End Try
                                    End If
                                End If
                            Case "b_WONo"
                                If pVal.BeforeAction = False Then
                                    CreateMySimpleForm()
                                End If
                        End Select
                    Catch ex As Exception
                        objaddon.objapplication.StatusBar.SetText("Item Press Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    Try
                        Select Case pVal.ItemUID
                            Case "1"
                                'Add,Update Event
                                'If pVal.BeforeAction = True And frmProductionOrder.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                '    docnum = oDBDSHeader.GetValue("Docnum", 0)
                                '    FGItemCode = frmProductionOrder.Items.Item("6").ToString.Trim

                                '    Dim ChildPoCreationFlag = frmProductionOrder.Items.Item("C_ChildPO").Specific.value.ToString.Trim
                                '    If pVal.BeforeAction = True And frmProductionOrder.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And ChildPoCreationFlag = "Y" Then
                                '        If objaddon.objcompany.InTransaction = False Then objaddon.objcompany.StartTransaction()
                                '        For i As Integer = 1 To oMatrix.VisualRowCount - 1
                                '            If Me.AutoProduction(oMatrix.Columns.Item("4").Cells.Item(i).Specific.Value, oMatrix.Columns.Item("14").Cells.Item(i).Specific.Value) = False Then
                                '                If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                '                BubbleEvent = False
                                '            End If
                                '        Next
                                '        If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                '    End If


                                'End If



                                'Changed by kannan
                                If pVal.BeforeAction = True And frmProductionOrder.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    docnum = oDBDSHeader.GetValue("Docnum", 0)
                                    FGItemCode = frmProductionOrder.Items.Item("6").ToString.Trim

                                    Dim ChildPoCreationFlag = frmProductionOrder.Items.Item("C_ChildPO").Specific.value.ToString.Trim
                                    Dim Replan = frmProductionOrder.Items.Item("C_Replan").Specific.value.ToString.Trim
                                    If pVal.BeforeAction = True And frmProductionOrder.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And ChildPoCreationFlag = "Y" And Replan <> "Replan" Then
                                        If objaddon.objcompany.InTransaction = False Then objaddon.objcompany.StartTransaction()
                                        For i As Integer = 1 To oMatrix.VisualRowCount - 1
                                            If Me.AutoProduction(oMatrix.Columns.Item("4").Cells.Item(i).Specific.Value, oMatrix.Columns.Item("14").Cells.Item(i).Specific.Value) = False Then
                                                If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                BubbleEvent = False
                                            End If
                                        Next
                                        If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                    End If

                                    'Replan
                                    If pVal.BeforeAction = True And frmProductionOrder.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And ChildPoCreationFlag = "Y" And Replan = "Replan" Then
                                        InQty = 0
                                        PreWODocEntry = ""
                                        If objaddon.objcompany.InTransaction = False Then objaddon.objcompany.StartTransaction()
                                        For i As Integer = 1 To oMatrix.VisualRowCount - 1
                                            Dim WONO = frmProductionOrder.Items.Item("t_WONo").Specific.value.ToString.Trim
                                            Dim strArr() As String
                                            strArr = WONO.Split("/")
                                            Dim ArrEndDate As String = strArr(1) + "/" + strArr(0) + "/" + strArr(2)
                                            Dim PlannedQty = oGFun.getSingleValue("select U_PlantQty from [@MIPL_WKO1] where docentry=(select max(docentry) from [@MIPL_OWKO] where docnum='" & strArr(0) & "')")
                                            PreWODocEntry = oGFun.getSingleValue("select max(docentry) from OWOR where U_WONO='" & WONO & "'")
                                            'Dim rs As SAPbobsCOM.Recordset
                                            'rs = oGFun.DoQuery(strsql)
                                            'For k As Integer = 1 To rs.RecordCount
                                            Dim strsql1 = "select * from WOR1 where Docentry='" & PreWODocEntry & "' and ItemCode='" & oMatrix.Columns.Item("4").Cells.Item(i).Specific.Value.ToString.Trim & "'"
                                            Dim rs1 As SAPbobsCOM.Recordset
                                            rs1 = oGFun.DoQuery(strsql1)
                                            For j As Integer = 1 To rs1.RecordCount
                                                Dim ProcessItemgroupcode = oGFun.getSingleValue("select itmsGrpCod from OITM where ItemCode='" & rs1.Fields.Item("ItemCode").Value & "'")
                                                ProcessItemgroupcode = IIf(ProcessItemgroupcode = "", 0, ProcessItemgroupcode)
                                                If ProcessItemgroupcode = 108 Then

                                                    If rs1.Fields.Item("ItemCode").Value.ToString.Trim = oMatrix.Columns.Item("4").Cells.Item(i).Specific.Value.ToString.Trim Then
                                                        'InQty = CDbl(InQty) + CDbl(rs1.Fields.Item("U_OutQty").Value)
                                                        InQty = CDbl(rs1.Fields.Item("U_InQty").Value)
                                                    End If

                                                End If
                                                rs1.MoveNext()
                                            Next
                                            'rs.MoveNext()

                                            'Next

                                            Dim OutQty = CDbl(frmProductionOrder.Items.Item("12").Specific.value) - CDbl(InQty)
                                            InQty = 0
                                            If CDbl(OutQty) > 0 Then
                                                If Me.AutoProduction(oMatrix.Columns.Item("4").Cells.Item(i).Specific.Value, OutQty) = False Then
                                                    If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                    BubbleEvent = False
                                                End If
                                            End If

                                        Next
                                        If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                    End If
                                End If



                            Case "33"
                                'Link Button on SAP BOM
                                Try
                                    If pVal.BeforeAction = False Then
                                        objaddon.objapplication.Menus.Item("4353").Activate()
                                        'Else
                                        Dim frm As SAPbouiCOM.Form
                                        frm = objaddon.objapplication.Forms.ActiveForm
                                        frm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        frm.Items.Item("4").Specific.Value = oDBDSHeader.GetValue("ItemCode", 0)
                                        frm.Items.Item("1").Click()
                                    End If
                                Catch ex As Exception
                                    objaddon.objapplication.StatusBar.SetText(ex.Message)
                                End Try

                        End Select
                    Catch ex As Exception
                        objaddon.objapplication.StatusBar.SetText("Click Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        BubbleEvent = False
                    Finally
                    End Try

            End Select
        Catch ex As Exception
            objaddon.objapplication.StatusBar.SetText("Item Event Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Function AutoProduction(ByVal Itemcode As String, ByVal Qty As Double) As Boolean
        Try

            Dim isBOM As String = oGFun.getSingleValue("select Case When Count(*) >0 then 'True' else 'False' end From OITT  where code='" & Itemcode & "' ")

            If CBool(isBOM) Then

                If Not AddProduction(Itemcode, Qty) Then Return False
                Dim rs As SAPbobsCOM.Recordset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim StrSql As String = "select Code,Quantity * " & Qty & " [Qty] from ITT1  where Father='" & Itemcode & "' "
                ' Dim StrSql As String = "select Code,Quantity [Qty] from ITT1  where Father='" & Itemcode & "' "
                rs.DoQuery(StrSql)
                For j = 0 To rs.RecordCount - 1
                    Dim isChilBOM As String = oGFun.getSingleValue("select Case When Count(*) >0 then 'True' else 'False' end From OITT  where code='" & rs.Fields.Item("Code").Value & "' ")
                    If CBool(isChilBOM) Then
                        Dim qty1 = rs.Fields.Item("Qty").Value
                        If AddProduction(rs.Fields.Item("Code").Value, rs.Fields.Item("Qty").Value) Then
                            AutoProduction(rs.Fields.Item("Code").Value, rs.Fields.Item("Qty").Value)
                        Else
                            Return False
                        End If

                    End If
                    rs.MoveNext()
                Next
            End If
            Return True
        Catch ex As Exception
            objaddon.objapplication.StatusBar.SetText("Production Order Stock Posting Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function
    Function AddProduction(ByVal ItemCode As String, ByVal PlannedQty As Double) As Boolean
        Try
            Dim WhsCode As String = oGFun.getSingleValue("Select ToWH from OITT where code='" & ItemCode & "'")
            Dim ChildPOCreation As String = oGFun.getSingleValue("Select U_ChildPOCreation from OITT where code='" & ItemCode & "'")
            Dim docEntry As String = ""
            docEntry = oGFun.getSingleValue(" Select Max(DocEntry) from OWOR")
            Dim ProItemcode As String = ""
            ProItemcode = oGFun.getSingleValue(" Select Itemcode from OWOR where docentry='" & docEntry & "'")

            If ItemCode.Trim <> ProItemcode.Trim And ChildPOCreation = "Y" Then
                Dim ErrCode
                Dim PostingDate = oGFun.getSingleValue(" Select Convert(DateTime,'" & frmProductionOrder.Items.Item("26").Specific.Value & "') Dt ")
                Dim OrderDate = oGFun.getSingleValue(" Select Convert(DateTime,'" & frmProductionOrder.Items.Item("24").Specific.Value & "') Dt ")
                Dim oProductionorder As SAPbobsCOM.ProductionOrders = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                oProductionorder.PostingDate = CDate(OrderDate)
                oProductionorder.DueDate = CDate(PostingDate)
                oProductionorder.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotStandard
                oProductionorder.ItemNo = ItemCode


                oProductionorder.PlannedQuantity = CDbl(PlannedQty)

                ' oProductionorder.Warehouse = frmProductionOrder.Items.Item("78").Specific.Value 'WhsCode
                oProductionorder.Warehouse = WhsCode

                ErrCode = oProductionorder.Add()

                If ErrCode <> 0 Then
                    objAddOn.objApplication.SetStatusBarMessage("Child Production Posting Error : " & objaddon.objcompany.GetLastErrorDescription)
                    Return False
                Else
                    objAddOn.objApplication.SetStatusBarMessage("Child Production OrderCreated Successfully, Item Code [" & ItemCode & "]")
                    If FGItemCode.ToString.Trim <> ItemCode.ToString.Trim Then
                        ParentPoNo(ItemCode)
                    End If
                End If
            End If




            Return True
        Catch ex As Exception
            objaddon.objapplication.StatusBar.SetText("Production Order Stock Posting Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function

    Function ParentPoNo(ByVal ItemCode As String)
        Dim FGCode = oGFun.getSingleValue("select Father from ITT1 where code='" & ItemCode & "' ")
        Dim ParentPOdocEntry = oGFun.getSingleValue(" Select Max(DocEntry) from OWOR where Itemcode='" & FGCode & "'")
        Dim ParentPODocnum = oGFun.getSingleValue(" Select DocNum from OWOR where DocEntry='" & ParentPOdocEntry & "'")
        Dim POdocEntry = oGFun.getSingleValue(" Select Max(DocEntry) from OWOR")


        Dim strsql = "Update OWOR set U_ParentPONo='" & ParentPODocnum & "' where docentry='" & POdocEntry & "'"
        oGFun.DoExQuery(strsql)


    End Function
    Sub Sub_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    'Assign Selected Rows
                    Dim oDataTable As SAPbouiCOM.DataTable
                    Dim oCFLE As SAPbouiCOM.ChooseFromListEvent = pVal
                    oDataTable = oCFLE.SelectedObjects
                    Dim rset As SAPbobsCOM.Recordset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'Filter before open the CFL
                    If pVal.BeforeAction Then
                        Select Case oCFLE.ChooseFromListUID
                            Case "CFL_2A"
                                oGFun.ChooseFromListFilteration(frmWO, oCFLE.ChooseFromListUID, "CardType", "Select 'C'")
                            Case "CFL_2B"
                                oGFun.ChooseFromListFilteration(frmWO, oCFLE.ChooseFromListUID, "CardType", "Select 'C'")
                        End Select
                    Else
                        Select Case oCFLE.ChooseFromListUID
                            Case "CFL_2A"
                                Try
                                    frmWO.Items.Item("t_CardCode").Specific.Value = oDataTable.GetValue("CardCode", 0)
                                Catch ex As Exception
                                End Try
                            Case "CFL_2B"
                                Try
                                    frmWO.Items.Item("t_CardName").Specific.Value = oDataTable.GetValue("CardName", 0)
                                Catch ex As Exception
                                End Try
                        End Select
                    End If

                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    'Case SAPbouiCOM.BoEventTypes.et_CLICK
                    Try
                        If pVal.BeforeAction = False Then
                            If pVal.ItemUID = "Grid" Then

                                Dim oGrid As SAPbouiCOM.Grid = frmWO.Items.Item("Grid").Specific
                                If pVal.Row < 0 Then Exit Sub

                                Dim DocNum = oGrid.DataTable.GetValue("DocNum", pVal.Row)
                                Dim DocEntry = oGrid.DataTable.GetValue("DocEntry", pVal.Row)
                                Dim lineid
                                Dim LineId1 = oGrid.DataTable.GetValue("lineid", pVal.Row)
                                Dim Worktype = oGrid.DataTable.GetValue("Orders Type", pVal.Row)
                                Dim WOBLN = oGrid.DataTable.GetValue("WOBLN", pVal.Row)
                                ' If Worktype = "1" Then
                                'lineid = oGrid.DataTable.GetValue("lineid", pVal.Row) t_WObaseLN  WOBLN
                                'Else
                                lineid = oGrid.DataTable.GetValue("SO LineNum", pVal.Row)
                                ' End If

                                frmProductionOrder.Items.Item("t_BaseNum").Specific.Value = DocNum
                                frmProductionOrder.Items.Item("t_BaseEntr").Specific.Value = DocEntry
                                frmProductionOrder.Items.Item("t_BaseLine").Specific.Value = lineid
                                frmProductionOrder.Items.Item("t_WObaseLN").Specific.Value = WOBLN

                                Dim Count = oGFun.getSingleValue("Select Count(*) +1  from OWOR where U_BaseNum = '" & DocNum & "'  and U_BaseLineId = '" & lineid & "' ")
                                frmProductionOrder.Items.Item("t_WONo").Specific.Value = DocNum.ToString.Trim + "/" + LineId1.ToString.Trim + "/" + Count.Trim

                                'frmProductionOrder.Items.Item("26").Specific.Value = oGrid.DataTable.GetValue("ProductionDate", i - 1)
                                frmProductionOrder.Items.Item("6").Specific.Value = oGrid.DataTable.GetValue("ItemCode", pVal.Row)
                                frmProductionOrder.Items.Item("12").Specific.Value = oGrid.DataTable.GetValue("ReqQty", pVal.Row)
                                Dim format As String
                                Dim result As DateTime
                                Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture
                                format = "yyyymmdd"

                                Dim docdate1 As String = oGrid.DataTable.GetValue("plantdate", pVal.Row)
                                result = Convert.ToDateTime(docdate1,
System.Globalization.CultureInfo.GetCultureInfo("hi-in").DateTimeFormat)
                                frmProductionOrder.Items.Item("26").Specific.string = result
                                ' Dim PostingDate = oGFun.getSingleValue(" Select Convert(DateTime,'" & oGrid.DataTable.GetValue("plantDate", pVal.Row) & "') Dt ")
                                ' frmProductionOrder.Items.Item("26").Specific.string = CDate(oGrid.DataTable.GetValue("plantDate", pVal.Row))
                                ' frmProductionOrder.Items.Item("26").Specific.string = CDate(PostingDate)
                                frmProductionOrder.Items.Item("78").Click()

                                Try
                                    frmProductionOrder.Items.Item("32").Specific.Value = oGrid.DataTable.GetValue("SONO", pVal.Row)
                                Catch ex As Exception
                                End Try
                                frmWO.Close()
                                Return


                            End If
                        End If
                    Catch ex As Exception
                        objaddon.objapplication.StatusBar.SetText(" Click Event Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    Try
                        Select Case pVal.ItemUID
                            Case "Grid"
                                Dim oGrid As SAPbouiCOM.Grid = frmWO.Items.Item("Grid").Specific
                                If pVal.BeforeAction Then
                                    oGrid.Columns.Item(pVal.ColUID).TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending)
                                End If
                        End Select
                    Catch ex As Exception

                    End Try
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED, SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    If pVal.ItemUID = "b_OK" Then
                        If pVal.BeforeAction = False Then
                            Dim oGrid As SAPbouiCOM.Grid = frmWO.Items.Item("Grid").Specific
                            For i As Integer = 0 To oGrid.Rows.Count - 1
                                If oGrid.Rows.IsSelected(i) Then
                                    Dim DocNum = oGrid.DataTable.GetValue("DocNum", i)
                                    Dim DocEntry = oGrid.DataTable.GetValue("DocEntry", i)
                                    Dim LineId1 = oGrid.DataTable.GetValue("lineid", i)
                                    Dim WorkType = oGrid.DataTable.GetValue("Orders Type", i)
                                    Dim WOBLN = oGrid.DataTable.GetValue("WOBLN", i)
                                    Dim lineid = oGrid.DataTable.GetValue("SO LineNum", i)

                                    frmProductionOrder.Items.Item("t_BaseNum").Specific.Value = DocNum
                                    frmProductionOrder.Items.Item("t_BaseEntr").Specific.Value = DocEntry
                                    frmProductionOrder.Items.Item("t_BaseLine").Specific.Value = lineid
                                    frmProductionOrder.Items.Item("t_WObaseLN").Specific.Value = WOBLN

                                    Dim Count = oGFun.getSingleValue("Select Count(*) +1  from OWOR where U_BaseNum = '" & DocNum & "'  and U_BaseLineId = '" & lineid & "' ")
                                    frmProductionOrder.Items.Item("t_WONo").Specific.Value = DocNum.ToString.Trim + "/" + LineId1.ToString.Trim + "/" + Count.Trim

                                    'frmProductionOrder.Items.Item("26").Specific.Value = oGrid.DataTable.GetValue("ProductionDate", i - 1)
                                    frmProductionOrder.Items.Item("6").Specific.Value = oGrid.DataTable.GetValue("ItemCode", i)
                                    frmProductionOrder.Items.Item("12").Specific.Value = oGrid.DataTable.GetValue("ReqQty", i)
                                    Dim format As String
                                    Dim result As DateTime
                                    Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture
                                    format = "yyyyMMdd"
                                    '                                Dim docdate1 As String = oGrid.DataTable.GetValue("plantDate", i)
                                    '                                result = Convert.ToDateTime(docdate1,
                                    'System.Globalization.CultureInfo.GetCultureInfo("hi-IN").DateTimeFormat)
                                    '                                frmProductionOrder.Items.Item("26").Specific.string = result
                                    'frmProductionOrder.Items.Item("26").Specific.string = CDate(oGFun.getSingleValue("select U_PlantDate from [@MIPL_wko1] WHERE docentry=(select max(DocEntry) from [@MIPL_OWKO] where docnum='" & DocNum & "')"))
                                    'frmProductionOrder.Items.Item("26").Specific.string = CDate(oGrid.DataTable.GetValue("plantDate", i))

                                    frmProductionOrder.Items.Item("78").Click()

                                    Try
                                        frmProductionOrder.Items.Item("32").Specific.Value = oGrid.DataTable.GetValue("SONO", i)
                                    Catch ex As Exception
                                    End Try

                                    frmWO.Close()
                                    Return
                                End If
                            Next

                        End If
                    End If
                    'Find btn 
                    If pVal.ItemUID = "b_Find" Then
                        If pVal.BeforeAction = False Then
                            Dim c_Orders As SAPbouiCOM.ComboBox = frmWO.Items.Item("c_Orders").Specific
                            Dim Orders As String
                            Orders = c_Orders.Selected.Value.Trim
                            Dim PartNo = frmWO.Items.Item("t_PartNo").Specific.Value
                            Dim CardCode = frmWO.Items.Item("t_CardCode").Specific.Value
                            Dim CardName = frmWO.Items.Item("t_CardName").Specific.Value
                            loadGrid(frmWO.Items.Item("t_WONO").Specific.Value.ToString.Trim, Orders, PartNo, CardCode, CardName)
                            Dim oGrid As SAPbouiCOM.Grid = frmWO.Items.Item("Grid").Specific
                            For i As Integer = 0 To oGrid.Columns.Count - 1
                                oGrid.Columns.Item(i).TitleObject.Sortable = True
                            Next
                        End If
                    End If
            End Select

        Catch ex As Exception
            objaddon.objapplication.StatusBar.SetText("Item Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub

    Private Sub CreateMySimpleForm()

        Dim oCreationParams As SAPbouiCOM.FormCreationParams

        If oGFun.FormExist("WORKorder") Then
            objaddon.objapplication.Forms.Item("WORKorder").Visible = True
        Else
            oCreationParams = objaddon.objapplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            oCreationParams.UniqueID = "WORKorder"
            frmWO = objaddon.objapplication.Forms.AddEx(oCreationParams)
            frmWO.Title = "WORK Order List"
            frmWO.Left = 400
            frmWO.Top = 100
            frmWO.ClientHeight = 345 '335
            frmWO.ClientWidth = 800
            frmWO = objaddon.objapplication.Forms.Item("WORKorder")
            Dim oitm As SAPbouiCOM.Item

            Dim stext As SAPbouiCOM.StaticText
            Dim etext As SAPbouiCOM.EditText
            Dim ocmbo As SAPbouiCOM.ComboBox
            Dim obtn As SAPbouiCOM.Button

            'Add button for Find
            oitm = frmWO.Items.Add("b_Find", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oitm.Top = 25
            oitm.Left = 650
            obtn = frmWO.Items.Item("b_Find").Specific
            obtn.Caption = "Find"
            frmWO.DefButton = "b_Find"

            'Add combobox for Orders Type
            oitm = frmWO.Items.Add("c_Orders", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oitm.Top = 10
            oitm.Left = 115
            oitm.Height = 14
            oitm.Width = 90
            oitm.DisplayDesc = True
            ocmbo = frmWO.Items.Item("c_Orders").Specific
            ocmbo.ValidValues.Add("-", "")
            ocmbo.ValidValues.Add("1", "Stock Order")
            ocmbo.ValidValues.Add("2", "Sales Order")
            ocmbo.ValidValues.Add("3", "Sales Return")
            ocmbo.ValidValues.Add("4", "WIP Update")
            ocmbo.Select("-", SAPbouiCOM.BoSearchKey.psk_ByValue)

            'Add Static Text for Orders
            oitm = frmWO.Items.Add("l_Orders", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oitm.Top = 10
            oitm.Left = 10
            oitm.Height = 14
            oitm.Width = 80
            oitm.LinkTo = "c_Orders"
            stext = frmWO.Items.Item("l_Orders").Specific
            stext.Caption = "Orders Type"

            'Add Edittext for PartNo 
            oitm = frmWO.Items.Add("t_PartNo", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oitm.Top = 25
            oitm.Left = 115
            oitm.Height = 14
            oitm.Width = 90
            etext = frmWO.Items.Item("t_PartNo").Specific

            'Add Static text for PartNo
            oitm = frmWO.Items.Add("l_PartNo", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oitm.Top = 25
            oitm.Left = 10
            oitm.Height = 14
            oitm.Width = 80
            oitm.LinkTo = "t_PartNo"
            stext = frmWO.Items.Item("l_PartNo").Specific
            stext.Caption = "PartNo"

            'Add Edit text for Work Order
            oitm = frmWO.Items.Add("t_WONO", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oitm.Top = 10
            oitm.Left = 350
            oitm.Height = 14
            oitm.Width = 80
            etext = frmWO.Items.Item("t_WONO").Specific

            'Add Static text for Work Order No
            oitm = frmWO.Items.Add("l_WONO", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oitm.Top = 10
            oitm.Left = 225
            oitm.Height = 14
            oitm.Width = 80
            oitm.LinkTo = "t_WONO"
            stext = frmWO.Items.Item("l_WONO").Specific
            stext.Caption = "Work Order No"

            'Add Edittext for Customer Name
            oitm = frmWO.Items.Add("t_CardCode", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oitm.Top = 25
            oitm.Left = 350
            oitm.Height = 14
            oitm.Width = 80
            etext = oitm.Specific

            'DataBind
            frmWO.DataSources.UserDataSources.Add("CardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            etext.DataBind.SetBound(True, "", "CardCode")

            'For CFL
            Dim oCFL As SAPbouiCOM.ChooseFromListCreationParams
            oCFL = objaddon.objapplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFL.UniqueID = "CFL_2A"
            oCFL.ObjectType = "2"
            frmWO.ChooseFromLists.Add(oCFL)

            etext.ChooseFromListUID = "CFL_2A"
            etext.ChooseFromListAlias = "CardCode"



            'Add Static text for Customer Name
            oitm = frmWO.Items.Add("l_CardCode", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oitm.Top = 25
            oitm.Left = 225
            oitm.Height = 14
            oitm.Width = 80
            oitm.LinkTo = "t_CardCode"
            stext = frmWO.Items.Item("l_CardCode").Specific
            stext.Caption = "Customer Code"


            'Add Edittext for Customer Name
            oitm = frmWO.Items.Add("t_CardName", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oitm.Top = 9
            oitm.Left = 520
            oitm.Height = 14
            oitm.Width = 80
            etext = oitm.Specific

            'DataBind
            frmWO.DataSources.UserDataSources.Add("CardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            etext.DataBind.SetBound(True, "", "CardName")

            'For CFL
            Dim oCFL1 As SAPbouiCOM.ChooseFromListCreationParams
            oCFL1 = objaddon.objapplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFL1.UniqueID = "CFL_2B"
            oCFL1.ObjectType = "2"
            frmWO.ChooseFromLists.Add(oCFL1)

            etext.ChooseFromListUID = "CFL_2B"
            etext.ChooseFromListAlias = "CardName"



            'Add Static text for Customer Name
            oitm = frmWO.Items.Add("l_CardName", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oitm.Top = 9
            oitm.Left = 225 + 210
            oitm.Height = 14
            oitm.Width = 80
            oitm.LinkTo = "t_CardName"
            stext = frmWO.Items.Item("l_CardName").Specific
            stext.Caption = "Customer Name"




            ''Add combobox for Sorting
            'oitm = frmWO.Items.Add("c_Sort", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            'oitm.Top = 25
            'oitm.Left = 650
            'oitm.Height = 14
            'oitm.Width = 90
            'oitm.DisplayDesc = True
            'ocmbo = frmWO.Items.Item("c_Sort").Specific
            'ocmbo.ValidValues.Add("_", "")
            'ocmbo.ValidValues.Add("1", "Order Type")
            'ocmbo.ValidValues.Add("2", "PartNo")
            'ocmbo.ValidValues.Add("3", "Workorderno")
            'ocmbo.ValidValues.Add("4", "CardCode")
            'ocmbo.ValidValues.Add("5", "CardName")
            'ocmbo.Select("_", SAPbouiCOM.BoSearchKey.psk_ByValue)

            ''Add Static Text for Sorting
            'oitm = frmWO.Items.Add("l_Sort", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            'oitm.Top = 25
            'oitm.Left = 600
            'oitm.Height = 14
            'oitm.Width = 80
            'oitm.LinkTo = "c_Sort"
            'stext = frmWO.Items.Item("l_Sort").Specific
            'stext.Caption = "Sort By"

            Dim oGrid As SAPbouiCOM.Grid
            oitm = frmWO.Items.Add("Grid", SAPbouiCOM.BoFormItemTypes.it_GRID)
            oitm.Top = 45
            oitm.Left = 2
            oitm.Width = 780
            oitm.Height = 270
            oGrid = frmWO.Items.Item("Grid").Specific
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto
            frmWO.DataSources.DataTables.Add("DataTable")

            oitm = frmWO.Items.Add("b_OK", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oitm.Top = frmWO.Items.Item("Grid").Top + frmWO.Items.Item("Grid").Height + 5
            oitm.Left = 10
            Dim btn As SAPbouiCOM.Button = frmWO.Items.Item("b_OK").Specific
            btn.Caption = "OK"
            Dim location = oGFun.getSingleValue(" Select U_Location from OUDG where Code=(select DfltsGroup from OUSR where USER_CODE ='" & objaddon.objcompany.UserName & "')")

            'Dim str_Sql = " Select OWKO.DocNum ,OWKO.DocEntry,lineid,U_BaseLineId [SO LineNum],(select docnum from ORDR where docentry=U_baseentry) [SONO],isnull(OWKO.U_WorkType,'') [Orders Type], OWKO.U_CardCode, OWKO.U_CardName,OITM.FrgnName PartCode, U_ParCode ItemCode," & _
            '    " U_ParName ItemName,U_SchQty ScheduledQty,U_SchDate, WKO1.U_ProQty ProducedQty, WKO1.U_ProDate ProducedDate, U_PlantDate plantDate,wko1.LineId WOBLN,isnull(U_SchQty,0) - isnull(U_StkQty,0) - isnull(U_ProQty,0) ReqQty" & _
            '    " from [@MIPL_OWKO] OWKO, [@MIPL_WKO1] WKO1,OITM where OWKO.U_Location = '" & location & "' and  OWKO.DocEntry = WKO1.DocEntry and OITM.ItemCode = WKO1.U_ParCode and WKO1.U_Status = 'O' " & _
            '" and isnull(U_SchQty,0) - isnull(U_StkQty,0) - isnull(U_ProQty,0) > 0 and WKO1.U_ParCode <> 'Cancelled' and owko.status<>'N' and owko.canceled<>'Y' order by OWKO.DocNum desc"

            '  Dim str_Sql = " Select OWKO.DocNum ,OWKO.DocEntry,lineid,U_BaseLineId [SO LineNum],(select docnum from ORDR where docentry=U_baseentry) [SONO],isnull(OWKO.U_WorkType,'') [Orders Type], OWKO.U_CardCode, OWKO.U_CardName,OITM.FrgnName PartCode, U_ParCode ItemCode," & _
            '    " U_ParName ItemName,U_SchQty ScheduledQty,U_SchDate, WKO1.U_ProQty ProducedQty, WKO1.U_ProDate ProducedDate, U_PlantDate plantDate,wko1.LineId WOBLN,isnull(U_PlantQty,0) - isnull(U_ComQty,0) ReqQty" & _
            '    " from [@MIPL_OWKO] OWKO, [@MIPL_WKO1] WKO1,OITM where OWKO.U_Location = '" & location & "' and  OWKO.DocEntry = WKO1.DocEntry and OITM.ItemCode = WKO1.U_ParCode and WKO1.U_Status = 'O' " & _
            '" and isnull(U_PlantQty,0) - isnull(U_ComQty,0) > 0 and WKO1.U_ParCode <> 'Cancelled' and owko.status<>'N' and owko.canceled<>'Y' order by OWKO.DocNum desc"

            Dim str_Sql = " Select OWKO.DocNum ,OWKO.DocEntry,lineid,U_BaseLineId [SO LineNum],ORDR.DocNum [SONO],isnull(OWKO.U_WorkType,'') [Orders Type], OWKO.U_CardCode, OWKO.U_CardName,OITM.FrgnName PartCode, U_ParCode ItemCode," & _
           " U_ParName ItemName,U_SchQty ScheduledQty,U_SchDate, WKO1.U_ProQty ProducedQty, WKO1.U_ProDate ProducedDate, U_PlantDate plantDate,wko1.LineId WOBLN,isnull(U_PlantQty,0) - isnull(U_ComQty,0) ReqQty" & _
           " from [@MIPL_OWKO] OWKO, [@MIPL_WKO1] WKO1,OITM,ORDR where OWKO.U_Location = '" & location & "' and  WKO1.U_BaseEntry=ORDR.DocEntry and OWKO.DocEntry = WKO1.DocEntry and OITM.ItemCode = WKO1.U_ParCode and WKO1.U_Status = 'O' " & _
       " and isnull(U_PlantQty,0) - isnull(U_ComQty,0) > 0 and WKO1.U_ParCode <> 'Cancelled' and owko.status<>'N' and owko.canceled<>'Y' order by OWKO.DocNum desc"


            frmWO.DataSources.DataTables.Item("DataTable").ExecuteQuery(str_Sql)
            oGrid.DataTable = frmWO.DataSources.DataTables.Item("DataTable")

            For i As Integer = 0 To oGrid.Columns.Count - 1

                If oGrid.Columns.Item(i).TitleObject.Caption.Trim <> "ReqQty" Then
                    oGrid.Columns.Item(i).Editable = False
                End If
            Next
            For i As Integer = 0 To oGrid.Columns.Count - 1
                oGrid.Columns.Item(i).TitleObject.Sortable = True
            Next
            frmWO.Visible = True
        End If
    End Sub

    Sub loadGrid(ByVal WONo As String, ByVal Orders As String, ByVal PartNo As String, ByVal CardCode As String, ByVal cardname As String)
        Try
            Dim condition As String = ""
            'If WONo.Trim = "" And Orders.Trim = "-" Then
            '    condition = ""
            'ElseIf WONo.Trim <> "" And Orders.Trim = "-" Then
            '    condition = " and OWKO.DocNum = '" & WONo & "' "
            'ElseIf WONo.Trim = "" And Orders.Trim <> "-" Then
            '    condition = " and isnull(U_WorkType,'')= '" & Orders.Trim & "' "
            'ElseIf WONo.Trim <> "" And Orders.Trim <> "-" Then
            '    condition = " and OWKO.DocNum = '" & WONo & "' and isnull(U_WorkType,'')= '" & Orders.Trim & "' "
            'End If

            If Orders.Trim <> "-" Then
                condition = condition & " and isnull(U_WorkType,'') = '" & Orders.Trim & "' "
            End If

            If WONo.Trim <> "" Then
                condition = condition & " and DocNum ='" & WONo.Trim & "'"
            End If

            If PartNo.Trim <> "" Then
                condition = condition & " and U_PartNo ='" & PartNo.Trim & "'"
            End If

            If CardCode.Trim <> "" Then
                condition = condition & " and OWKO.U_CardCode ='" & CardCode.Trim & "'"
            End If

            If CardName.Trim <> "" Then
                condition = condition & " and OWKO.U_CardName ='" & cardname.Trim & "'"
            End If

            Dim location = oGFun.getSingleValue(" Select U_Location from OUDG where Code=(select DfltsGroup from OUSR where USER_CODE ='" & objaddon.objcompany.UserName & "')")

            'Dim str_Sql = " Select OWKO.DocNum, OWKO.DocEntry,lineid,U_BaseLineId [SO LineNum],(select docnum from ORDR where docentry=U_baseentry) [SONO],OWKO.U_WorkType [Orders Type], OWKO.U_CardCode, OWKO.U_CardName,OITM.FrgnName PartCode, U_ParCode ItemCode," & _
            '    " U_ParName ItemName,U_SchQty ScheduledQty,U_SchDate, WKO1.U_ProQty ProducedQty, WKO1.U_ProDate ProducedDate,U_PlantDate plantDate,wko1.LineId WOBLN, isnull(U_SchQty,0) - isnull(U_StkQty,0) - isnull(U_ProQty,0) ReqQty" & _
            '    " from [@MIPL_OWKO] OWKO, [@MIPL_WKO1] WKO1,OITM where OWKO.U_Location = '" & location & "' and  OWKO.DocEntry = WKO1.DocEntry and OITM.ItemCode = WKO1.U_ParCode and WKO1.U_Status = 'O' " & _
            '    " and isnull(U_SchQty,0) - isnull(U_StkQty,0) - isnull(U_ProQty,0) > 0 and WKO1.U_ParCode <> 'Cancelled' " & _
            'condition

            Dim str_Sql = " Select OWKO.DocNum, OWKO.DocEntry,lineid,U_BaseLineId [SO LineNum],(select docnum from ORDR where docentry=U_baseentry) [SONO],OWKO.U_WorkType [Orders Type], OWKO.U_CardCode, OWKO.U_CardName,OITM.FrgnName PartCode, U_ParCode ItemCode," & _
              " U_ParName ItemName,U_SchQty ScheduledQty,U_SchDate, WKO1.U_ProQty ProducedQty, WKO1.U_ProDate ProducedDate,U_PlantDate plantDate,wko1.LineId WOBLN, isnull(U_PlantQty,0) - isnull(U_ComQty,0) ReqQty" & _
              " from [@MIPL_OWKO] OWKO, [@MIPL_WKO1] WKO1,OITM where OWKO.U_Location = '" & location & "' and  OWKO.DocEntry = WKO1.DocEntry and OITM.ItemCode = WKO1.U_ParCode and WKO1.U_Status = 'O' " & _
              " and isnull(U_PlantQty,0) - isnull(U_ComQty,0)  > 0 and WKO1.U_ParCode <> 'Cancelled' " & _
          condition

            frmWO.DataSources.DataTables.Item("DataTable").ExecuteQuery(str_Sql)

        Catch ex As Exception
        End Try

    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID

                Case "1282"
                    Me.InitForm()
                Case "1293"
                    oGFun.DeleteRow(oMatrix, oDBDSDetail)
                Case "1287"


            End Select
            ParentCode = frmProductionOrder.Items.Item("6").Specific.Value

        Catch ex As Exception
            objaddon.objapplication.StatusBar.SetText("Menu Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub



    Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)

    End Sub













    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                    Try
                        'Dim ChildPoCreationFlag = frmProductionOrder.Items.Item("C_ChildPO").Specific.value.ToString.Trim
                        'If BusinessObjectInfo.ActionSuccess And frmProductionOrder.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And ChildPoCreationFlag = "Y" Then
                        '    If objaddon.objcompany.InTransaction = False Then objaddon.objcompany.StartTransaction()
                        '    For i As Integer = 1 To oMatrix.VisualRowCount - 1
                        '        If Me.AutoProduction(oMatrix.Columns.Item("4").Cells.Item(i).Specific.Value, oMatrix.Columns.Item("14").Cells.Item(i).Specific.Value) = False Then
                        '            If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        '            BubbleEvent = False
                        '        End If
                        '    Next
                        '    If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        'End If
                        'Changed by kannan
                        If BusinessObjectInfo.ActionSuccess And frmProductionOrder.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            docnum = oDBDSHeader.GetValue("Docnum", 0)
                            FGItemCode = frmProductionOrder.Items.Item("6").ToString.Trim

                            Dim ChildPoCreationFlag = frmProductionOrder.Items.Item("C_ChildPO").Specific.value.ToString.Trim
                            Dim Replan = frmProductionOrder.Items.Item("C_Replan").Specific.value.ToString.Trim
                            If BusinessObjectInfo.ActionSuccess And frmProductionOrder.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And ChildPoCreationFlag = "Y" And Replan <> "Replan" Then
                                If objaddon.objcompany.InTransaction = False Then objaddon.objcompany.StartTransaction()
                                For i As Integer = 1 To oMatrix.VisualRowCount - 1
                                    If Me.AutoProduction(oMatrix.Columns.Item("4").Cells.Item(i).Specific.Value, oMatrix.Columns.Item("14").Cells.Item(i).Specific.Value) = False Then
                                        If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                        BubbleEvent = False
                                    End If
                                Next
                                If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                            End If

                            'Replan
                            If BusinessObjectInfo.ActionSuccess And frmProductionOrder.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And ChildPoCreationFlag = "Y" And Replan = "Replan" Then
                                InQty = 0
                                If objaddon.objcompany.InTransaction = False Then objaddon.objcompany.StartTransaction()
                                For i As Integer = 1 To oMatrix.VisualRowCount - 1
                                    Dim WONO = frmProductionOrder.Items.Item("t_WONo").Specific.value.ToString.Trim
                                    Dim strArr() As String
                                    strArr = WONO.Split("/")
                                    Dim ArrEndDate As String = strArr(1) + "/" + strArr(0) + "/" + strArr(2)
                                    Dim PlannedQty = oGFun.getSingleValue("select U_PlantQty from [@MIPL_WKO1] where docentry=(select max(docentry) from [@MIPL_OWKO] where docnum='" & strArr(0) & "')")

                                    '  Dim Docentry = oGFun.getSingleValue("select max(docentry) from OWOR where U_WONO='" & WONO & "'")
                                    'Dim rs As SAPbobsCOM.Recordset
                                    'rs = oGFun.DoQuery(strsql)
                                    'For k As Integer = 1 To rs.RecordCount
                                    ' Dim strsql1 = "select * from WOR1 where Docentry=(select max(docentry) from OWOR where docnum='" & rs.Fields.Item("Docnum").Value & "') and ItemCode='" & oMatrix.Columns.Item("4").Cells.Item(i).Specific.Value.ToString.Trim & "'"
                                    Dim strsql1 = "select * from WOR1 where Docentry='" & PreWODocEntry & "' and ItemCode='" & oMatrix.Columns.Item("4").Cells.Item(i).Specific.Value.ToString.Trim & "'"
                                    Dim rs1 As SAPbobsCOM.Recordset
                                    rs1 = oGFun.DoQuery(strsql1)
                                    For j As Integer = 1 To rs1.RecordCount
                                        Dim ProcessItemgroupcode = oGFun.getSingleValue("select itmsGrpCod from OITM where ItemCode='" & rs1.Fields.Item("ItemCode").Value & "'")
                                        ProcessItemgroupcode = IIf(ProcessItemgroupcode = "", 0, ProcessItemgroupcode)
                                        If ProcessItemgroupcode = 108 Then

                                            If rs1.Fields.Item("ItemCode").Value.ToString.Trim = oMatrix.Columns.Item("4").Cells.Item(i).Specific.Value.ToString.Trim Then
                                                ' InQty = CDbl(InQty) + CDbl(rs1.Fields.Item("U_OutQty").Value)
                                                InQty = CDbl(rs1.Fields.Item("U_InQty").Value)
                                            End If

                                        End If
                                        rs1.MoveNext()
                                    Next
                                    ' rs.MoveNext()

                                    ' Next

                                    ' Dim OutQty = CDbl(PlannedQty) - CDbl(InQty)
                                    Dim OutQty = CDbl(frmProductionOrder.Items.Item("12").Specific.value) - CDbl(InQty)

                                    If CDbl(OutQty) > 0 Then
                                        If Me.AutoProduction(oMatrix.Columns.Item("4").Cells.Item(i).Specific.Value, OutQty) = False Then
                                            If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                            BubbleEvent = False
                                        End If
                                    End If
                                    'changed By Kannan
                                    'Production order in qty Update
                                    Dim ProDocNum = oDBDSHeader.GetValue("Docnum", 0)
                                    Dim docentry = oGFun.getSingleValue("select max(docentry) from OWOR where docnum='" & ProDocNum & "'")
                                    Dim strSql = "update WOR1 set U_InQty='" & InQty & "' where docentry='" & docentry & "' and ItemCode='" & oMatrix.Columns.Item("4").Cells.Item(i).Specific.Value.ToString.Trim & "'"
                                    oGFun.DoQuery(strSql)
                                    InQty = 0
                                Next
                                If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                            End If
                        End If






                    Catch ex As Exception
                        objaddon.objapplication.StatusBar.SetText("Form Data ADD Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        BubbleEvent = False
                    Finally
                    End Try

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.ActionSuccess = True Then
                        Try

                        Catch ex As Exception
                            objaddon.objapplication.StatusBar.SetText("Data Load Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End Try
                    End If
            End Select
        Catch ex As Exception
            objaddon.objapplication.StatusBar.SetText("Form Data Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    
  

    

    

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
