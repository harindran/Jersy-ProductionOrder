Imports System.Linq
Public Class clsProductionOrder
    Public Const frmType As String = "65211"
    Dim objForm As SAPbouiCOM.Form
    Dim frmWO As SAPbouiCOM.Form
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim docnum As String
    Dim FGItemCode As String
    Dim SuccessFlag As Boolean = True
    Dim ItemCodes As String()
    Dim ItemCount As Integer = 0
    Dim POStat As Boolean = True

    Public Sub ItemEvent(FormUID As String, pval As SAPbouiCOM.ItemEvent, BubbleEvent As System.Boolean)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        If pval.BeforeAction = True Then
            Select Case pval.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    CreateObjects(FormUID)
            End Select
        Else
            Select Case pval.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pval.ItemUID = "b_WONo" Then
                        If CStr(objForm.DataSources.DBDataSources.Item("OWOR").GetValue("U_FGPOEntry", 0)) = "" Then
                            CreateMySimpleForm(CStr(objForm.DataSources.DBDataSources.Item("OWOR").GetValue("DocEntry", 0)))
                        Else
                            CreateMySimpleForm(CStr(objForm.DataSources.DBDataSources.Item("OWOR").GetValue("U_FGPOEntry", 0)))
                        End If
                    End If
            End Select


        End If
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As System.Boolean)
        objForm = objAddOn.objApplication.Forms.Item(BusinessObjectInfo.FormUID)
        Try
            If BusinessObjectInfo.BeforeAction Then
            Else
                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                        Try
                            If BusinessObjectInfo.ActionSuccess And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                SuccessFlag = True : POStat = True
                                docnum = objForm.DataSources.DBDataSources.Item("OWOR").GetValue("DocNum", 0)
                                FGItemCode = objForm.Items.Item("6").Specific.string
                                ReDim ItemCodes(30)
                                ItemCodes(ItemCount) = FGItemCode
                                objMatrix = objForm.Items.Item("37").Specific
                                Dim ChildPoCreationFlag = objForm.Items.Item("C_ChildPO").Specific.value.ToString.Trim
                                '  Dim Replan = objform.Items.Item("C_Replan").Specific.value.ToString.Trim
                                Dim FinalDT As New DataTable
                                Dim Code As String, ConsolidatedPO As String = ""
                                Dim Quant As Double
                                If ChildPoCreationFlag = "Y" Then
                                    If objAddOn.HANA Then
                                        ConsolidatedPO = objAddOn.objGenFunc.getSingleValue("select ""U_POCon"" from OADM where ifnull(""U_POCon"",'')='Y'")
                                    Else
                                        ConsolidatedPO = objAddOn.objGenFunc.getSingleValue("select U_POCon from OADM where isnull(U_POCon,'')='Y'")
                                    End If

                                    If ConsolidatedPO = "Y" Then
                                        For i As Integer = 1 To objMatrix.RowCount - 1
                                            FinalDT = GettingBOM(objMatrix.Columns.Item("4").Cells.Item(i).Specific.Value, objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value)
                                        Next
                                        Dim sums = From dr In FinalDT.AsEnumerable()
                                                   Group dr By Ph = dr.Field(Of String)("Code") Into drg = Group
                                                   Select New With {
                                                   .Ph = Ph,
                                                   .LengthSum = drg.Sum(Function(dr) dr.Field(Of Double)("Qty"))
                                                   }
                                        If objAddOn.objCompany.InTransaction = False Then objAddOn.objCompany.StartTransaction()
                                        For Each RowID In sums
                                            Code = RowID.Ph.ToString()
                                            Quant = RowID.LengthSum
                                            If AddProduction(Code, Quant) = False Then
                                                BubbleEvent = False
                                                SuccessFlag = False
                                                POStat = False
                                            End If
                                        Next
                                    Else
                                        If objAddOn.objCompany.InTransaction = False Then objAddOn.objCompany.StartTransaction()
                                        For i As Integer = 1 To objMatrix.RowCount - 1
                                            If Me.AutoProduction(objMatrix.Columns.Item("4").Cells.Item(i).Specific.Value, objMatrix.Columns.Item("14").Cells.Item(i).Specific.Value) = False Then
                                                'If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                SuccessFlag = False
                                                BubbleEvent = False
                                                POStat = False
                                            End If
                                        Next
                                    End If
                                    If SuccessFlag = False Or POStat = False Or BubbleEvent = False Then
                                        If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                        objAddOn.objApplication.MessageBox("Production Orders Rolled Back...", , "OK")
                                        objAddOn.objApplication.StatusBar.SetText("Production Orders Rolled Back...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False : Exit Sub
                                    Else
                                        If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                        objAddOn.objApplication.MessageBox("Production Orders Created Successfully...", , "OK")
                                        objAddOn.objApplication.StatusBar.SetText("Production Orders Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    End If
                                    ItemCount = 0
                                    Array.Clear(ItemCodes, 0, ItemCodes.Length)
                                End If
                            End If

                            GC.Collect()
                        Catch ex As Exception
                            GC.Collect()
                            objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End Try
                End Select
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub CreateObjects(ByVal FormUID As String)

        Dim oItem As SAPbouiCOM.Item
        Dim oLabel As SAPbouiCOM.StaticText
        Dim oEditText As SAPbouiCOM.EditText
        Dim oButton As SAPbouiCOM.Button
        Dim oComboBox As SAPbouiCOM.ComboBox
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        'Production No.
        Try
            'Child PO
            'ComboBox
            oItem = objForm.Items.Add("C_ChildPO", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oItem.Left = objForm.Items.Item("78").Left
            oItem.Width = objForm.Items.Item("78").Width
            oItem.Height = objForm.Items.Item("78").Height
            oItem.Top = objForm.Items.Item("234000018").Top + objForm.Items.Item("234000018").Height + 2 'objForm.Items.Item("78").Top + 30
            oItem.Enabled = True
            oComboBox = oItem.Specific
            oComboBox.DataBind.SetBound(True, "OWOR", "U_ChildPO")

            'Static Box
            oItem = objForm.Items.Add("s_ChildPO", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = objForm.Items.Item("77").Left
            oItem.Top = objForm.Items.Item("234000018").Top + objForm.Items.Item("234000018").Height + 2
            oItem.Height = objForm.Items.Item("77").Height
            oItem.Width = objForm.Items.Item("77").Width
            oItem.LinkTo = "C_ChildPO"
            oLabel = oItem.Specific
            oLabel.Caption = "Child PO Creation"

            oItem = objForm.Items.Add("t_WONo", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = objForm.Items.Item("78").Left
            oItem.Width = objForm.Items.Item("78").Width
            oItem.Height = objForm.Items.Item("78").Height
            oItem.Top = objForm.Items.Item("C_ChildPO").Top + objForm.Items.Item("C_ChildPO").Height + 2
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "OWOR", "U_WONo") '25Jun2019

            oItem = objForm.Items.Add("l_WONo", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = objForm.Items.Item("77").Left
            oItem.Top = objForm.Items.Item("s_ChildPO").Top + objForm.Items.Item("s_ChildPO").Height + 2
            oItem.Height = objForm.Items.Item("77").Height
            oItem.Width = objForm.Items.Item("77").Width
            oItem.LinkTo = "t_WONo"
            oLabel = oItem.Specific
            oLabel.Caption = "Prod Order Entry"

            oItem = objForm.Items.Add("b_WONo", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = objForm.Items.Item("t_WONo").Left + objForm.Items.Item("t_WONo").Width + 2
            oItem.Width = 20
            oItem.Height = objForm.Items.Item("t_WONo").Height
            oItem.Top = objForm.Items.Item("t_WONo").Top
            oButton = oItem.Specific
            oButton.Caption = "||"

        Catch ex As Exception
            GC.Collect()
            objAddOn.objApplication.StatusBar.SetText("InitForm Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
            objForm.Freeze(False)

        End Try
    End Sub

    Dim objDT As New DataTable
    Dim StrDT As String
    Dim FinalD As DataTable

    Function AutoProduction(ByVal Itemcode As String, ByVal Qty As Double) As Boolean
        Try
            Dim StrSql As String

            Dim isBOM As String
            If objAddOn.HANA Then
                isBOM = objAddOn.objGenFunc.getSingleValue("select Case When Count(*) >0 then 'True' else 'False' end From OITT  where ""Code""='" & Itemcode & "' ")
            Else
                isBOM = objAddOn.objGenFunc.getSingleValue("select Case When Count(*) >0 then 'True' else 'False' end From OITT  where Code='" & Itemcode & "' ")
            End If
            If CBool(isBOM) Then
                If Not AddProduction(Itemcode, Qty) Then Return False
                Dim rs As SAPbobsCOM.Recordset = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If objAddOn.HANA Then
                    StrSql = "select ""Code"",""Quantity"" * " & Qty & " ""Qty""  from ITT1  where ""Father""='" & Itemcode & "' "
                Else
                    StrSql = "select Code,Quantity * " & Qty & " [Qty] from ITT1  where Father='" & Itemcode & "' "
                End If

                ' Dim StrSql As String = "select Code,Quantity [Qty] from ITT1  where Father='" & Itemcode & "' "
                rs.DoQuery(StrSql)
                For j = 0 To rs.RecordCount - 1
                    Dim isChilBOM As String
                    If objAddOn.HANA Then
                        isChilBOM = objAddOn.objGenFunc.getSingleValue("SELECT CASE WHEN COUNT(*) > 0 THEN 'True' ELSE 'False' END FROM OITT WHERE ""Code"" = '" & rs.Fields.Item("Code").Value & "';")
                    Else
                        isChilBOM = objAddOn.objGenFunc.getSingleValue("select Case When Count(*) >0 then 'True' else 'False' end From OITT  where code='" & rs.Fields.Item("Code").Value & "' ")
                    End If
                    ' objAddOn.WriteSMSLog("recordset index" & j & Itemcode)
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
            GC.Collect()
            objAddOn.objApplication.StatusBar.SetText("Production Order Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function

    Function AddProduction(ByVal ItemCode As String, ByVal PlannedQty As Double) As Boolean
        Dim oProductionorder As SAPbobsCOM.ProductionOrders = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
        Try
            Dim DBDataSource As SAPbouiCOM.DBDataSource
            DBDataSource = objForm.DataSources.DBDataSources.Item("OWOR")
            Dim WhsCode As String
            Dim docEntry As String = ""
            Dim ChildPOCreation As String
            Dim ProItemcode As String = ""
            If objAddOn.HANA Then
                WhsCode = objAddOn.objGenFunc.getSingleValue("Select ""ToWH"" from OITT where ""Code""='" & ItemCode & "'")
                docEntry = objAddOn.objGenFunc.getSingleValue(" Select Max(""DocEntry"") from OWOR")
                ChildPOCreation = objAddOn.objGenFunc.getSingleValue("Select ""U_ChildPOCreation"" from OITT where ""Code""='" & ItemCode & "'")
                ProItemcode = objAddOn.objGenFunc.getSingleValue(" Select ""ItemCode"" from OWOR where ""DocEntry""='" & docEntry & "'")
            Else
                WhsCode = objAddOn.objGenFunc.getSingleValue("Select ToWH from OITT where Code='" & ItemCode & "'")
                ChildPOCreation = objAddOn.objGenFunc.getSingleValue("Select U_ChildPOCreation from OITT where code='" & ItemCode & "'")
                docEntry = objAddOn.objGenFunc.getSingleValue(" Select Max(DocEntry) from OWOR")
                ProItemcode = objAddOn.objGenFunc.getSingleValue(" Select Itemcode from OWOR where docentry='" & docEntry & "'")
            End If

            If ItemCode.Trim <> ProItemcode.Trim And ChildPOCreation = "Y" Then
                Dim ErrCode
                Dim PostingDate
                Dim OrderDate

                If objAddOn.HANA Then
                    PostingDate = objAddOn.objGenFunc.getSingleValue("SELECT CAST('" & objForm.Items.Item("26").Specific.Value & "' AS timestamp) AS ""Dt"" FROM DUMMY;")
                    OrderDate = objAddOn.objGenFunc.getSingleValue(" Select CAST('" & objForm.Items.Item("24").Specific.Value & "' AS timestamp) AS ""Dt"" FROM DUMMY;")
                Else
                    PostingDate = objAddOn.objGenFunc.getSingleValue(" Select Convert(DateTime,'" & objForm.Items.Item("26").Specific.Value & "') Dt ")
                    OrderDate = objAddOn.objGenFunc.getSingleValue(" Select Convert(DateTime,'" & objForm.Items.Item("24").Specific.Value & "') Dt ")
                End If

                oProductionorder.PostingDate = CDate(OrderDate)
                oProductionorder.DueDate = CDate(PostingDate)
                oProductionorder.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotStandard
                oProductionorder.ItemNo = ItemCode

                oProductionorder.PlannedQuantity = CDbl(PlannedQty)

                ' oProductionorder.Warehouse = objform.Items.Item("78").Specific.Value 'WhsCode
                oProductionorder.Warehouse = WhsCode
                oProductionorder.Remarks = "Through Add-on " & Now.ToString
                ErrCode = oProductionorder.Add()

                If ErrCode <> 0 Then
                    objAddOn.objApplication.MessageBox("Child Production Posting Error : " & objAddOn.objCompany.GetLastErrorDescription, , "OK")
                    objAddOn.objApplication.SetStatusBarMessage("Child Production Posting Error : " & objAddOn.objCompany.GetLastErrorDescription)
                    POStat = False
                    Return False
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oProductionorder)
                    GC.Collect()
                Else
                    'objAddOn.objApplication.SetStatusBarMessage("Child Production OrderCreated Successfully, Item Code [" & ItemCode & "]", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    objAddOn.objApplication.StatusBar.SetText("Child Production OrderCreated Successfully, Item Code [" & ItemCode & "]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    ItemCount = ItemCount + 1
                    ItemCodes(ItemCount) = ItemCode
                    If FGItemCode.ToString.Trim <> ItemCode.ToString.Trim Then
                        'ParentPoNo(ItemCode, objForm.DataSources.DBDataSources.Item("OWOR").GetValue("DocEntry", 0))
                        ParentPoNo(ItemCode, DBDataSource.GetValue("DocEntry", 0))
                    End If
                End If
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oProductionorder)
            GC.Collect()
            Return True
        Catch ex As Exception
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oProductionorder)
            GC.Collect()
            objAddOn.objApplication.StatusBar.SetText("PO Stock Posting Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function

    Function ParentPoNo(ByVal ItemCode As String, ByVal FGPOEntry As String)
        Dim FGCode
        Dim ParentPOdocEntry
        Dim ParentPODocnum
        Dim POdocEntry
        Dim strsql
        Try
            If objAddOn.HANA Then
                Dim StrItemCodes As String = ""
                For i As Integer = 0 To ItemCodes.Length - 1
                    If ItemCodes(i) <> "" Then
                        StrItemCodes = StrItemCodes + ItemCodes(i) + "','"
                    End If
                Next
                StrItemCodes = StrItemCodes.Remove(StrItemCodes.Length - 2)
                StrItemCodes = "'" + StrItemCodes
                strsql = "SELECT ""Father"" FROM ITT1 WHERE ""Code"" = '" & ItemCode & "' AND ""Father"" in (" & StrItemCodes & ")"

                FGCode = objAddOn.objGenFunc.getSingleValue(strsql)
                ParentPOdocEntry = objAddOn.objGenFunc.getSingleValue(" SELECT MAX(""DocEntry"") FROM OWOR WHERE ""ItemCode"" ='" & FGCode & "'")

                ParentPODocnum = objAddOn.objGenFunc.getSingleValue(" SELECT ""DocNum"" FROM OWOR WHERE ""DocEntry""='" & ParentPOdocEntry & "'")

                POdocEntry = objAddOn.objGenFunc.getSingleValue(" SELECT MAX(""DocEntry"") FROM OWOR")
                ' objAddOn.WriteSMSLog("Update Query" & FGCode & "ParentPOdocentry" & ParentPOdocEntry & "PODocnum" & ParentPODocnum & "POdocentry" & POdocEntry)
                strsql = "Update OWOR set ""U_ParentPONo""='" & ParentPODocnum & "', ""U_ParentPOEntry""='" & ParentPOdocEntry & "', ""U_FGPOEntry""='" & FGPOEntry & "'  where ""DocEntry""='" & POdocEntry & "'"
            Else
                FGCode = objAddOn.objGenFunc.getSingleValue("select Father from ITT1 where code='" & ItemCode & "' ")
                ParentPOdocEntry = objAddOn.objGenFunc.getSingleValue(" Select Max(DocEntry) from OWOR where Itemcode='" & FGCode & "'")
                ParentPODocnum = objAddOn.objGenFunc.getSingleValue(" Select DocNum from OWOR where DocEntry='" & ParentPOdocEntry & "'")
                POdocEntry = objAddOn.objGenFunc.getSingleValue(" Select Max(DocEntry) from OWOR")
                strsql = "Update OWOR set U_ParentPONo='" & ParentPODocnum & "', U_ParentPOEntry='" & ParentPOdocEntry & "' , U_FGPOEntry='" & FGPOEntry & "' where docentry='" & POdocEntry & "'" ' 25Jun2019
            End If

            ' objAddOn.WriteSMSLog("ParentPoNo" & ItemCount)
            Dim objRS As SAPbobsCOM.Recordset
            objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRS.DoQuery(strsql)
            objRS = Nothing
            GC.Collect()
            'Array.Clear(ItemCodes, 0, ItemCodes.Length)
        Catch ex As Exception
            GC.Collect()
            objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Function

    Private Sub CreateMySimpleForm(ByVal ParentPOEntry As String)

        Dim oCreationParams As SAPbouiCOM.FormCreationParams
        Try
            objAddOn.objApplication.Forms.Item("ProdOrderStatus").Visible = True
        Catch ex As Exception
            oCreationParams = objAddOn.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            oCreationParams.UniqueID = "ProdOrderStatus"
            frmWO = objAddOn.objApplication.Forms.AddEx(oCreationParams)
            frmWO.Title = "Multi-Production Order List"
            frmWO.Left = 400
            frmWO.Top = 100
            frmWO.ClientHeight = 360 '335
            frmWO.ClientWidth = 800
            frmWO = objAddOn.objApplication.Forms.Item("ProdOrderStatus")
            Dim oitm As SAPbouiCOM.Item

            Dim stext As SAPbouiCOM.StaticText
            Dim etext As SAPbouiCOM.EditText
            Dim ocmbo As SAPbouiCOM.ComboBox
            Dim obtn As SAPbouiCOM.Button

            'Add button for Find
            'oitm = frmWO.Items.Add("b_Find", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            'oitm.Top = 25
            'oitm.Left = 650
            'obtn = frmWO.Items.Item("b_Find").Specific
            'obtn.Caption = "Find"
            'frmWO.DefButton = "b_Find"

            ''Add combobox for Orders Type
            'oitm = frmWO.Items.Add("c_Orders", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            'oitm.Top = 10
            'oitm.Left = 115
            'oitm.Height = 14
            'oitm.Width = 90
            'oitm.DisplayDesc = True
            'ocmbo = frmWO.Items.Item("c_Orders").Specific
            'ocmbo.ValidValues.Add("-", "")
            'ocmbo.ValidValues.Add("1", "Stock Order")
            'ocmbo.ValidValues.Add("2", "Sales Order")
            'ocmbo.ValidValues.Add("3", "Sales Return")
            'ocmbo.ValidValues.Add("4", "WIP Update")
            'ocmbo.Select("-", SAPbouiCOM.BoSearchKey.psk_ByValue)

            ''Add Static Text for Orders
            'oitm = frmWO.Items.Add("l_Orders", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            'oitm.Top = 10
            'oitm.Left = 10
            'oitm.Height = 14
            'oitm.Width = 80
            'oitm.LinkTo = "c_Orders"
            'stext = frmWO.Items.Item("l_Orders").Specific
            'stext.Caption = "Orders Type"

            ''Add Edittext for PartNo 
            'oitm = frmWO.Items.Add("t_PartNo", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'oitm.Top = 25
            'oitm.Left = 115
            'oitm.Height = 14
            'oitm.Width = 90
            'etext = frmWO.Items.Item("t_PartNo").Specific

            ''Add Static text for PartNo
            'oitm = frmWO.Items.Add("l_PartNo", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            'oitm.Top = 25
            'oitm.Left = 10
            'oitm.Height = 14
            'oitm.Width = 80
            'oitm.LinkTo = "t_PartNo"
            'stext = frmWO.Items.Item("l_PartNo").Specific
            'stext.Caption = "PartNo"

            Dim ousrdtsource As SAPbouiCOM.UserDataSource
            'Add Edit text for Work Order
            oitm = frmWO.Items.Add("t_WONO", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oitm.Top = 10
            oitm.Left = 200
            oitm.Height = 19
            oitm.Width = 100
            etext = frmWO.Items.Item("t_WONO").Specific
            ousrdtsource = frmWO.DataSources.UserDataSources.Add("DS1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)
            etext.DataBind.SetBound(True, "", "DS1")
            etext.Item.Enabled = False

            'Add Static text for Work Order No
            oitm = frmWO.Items.Add("l_WONO", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oitm.Top = 10
            oitm.Left = 25
            oitm.Height = 14
            oitm.Width = 160
            oitm.LinkTo = "t_WONO"
            stext = frmWO.Items.Item("l_WONO").Specific
            stext.Caption = "Parent Production No."

            Dim oGrid As SAPbouiCOM.Grid
            oitm = frmWO.Items.Add("Grid", SAPbouiCOM.BoFormItemTypes.it_GRID)
            oitm.Top = 45
            oitm.Left = 2
            oitm.Width = 780
            oitm.Height = 270
            oGrid = frmWO.Items.Item("Grid").Specific


            ' oGrid.SelectionMode = SAPbouiCOM.BobjMatrixSelect.ms_Auto
            frmWO.DataSources.DataTables.Add("DataTable")

            oitm = frmWO.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oitm.Top = frmWO.Items.Item("Grid").Top + frmWO.Items.Item("Grid").Height + 5
            oitm.Left = 10
            'Dim btn As SAPbouiCOM.Button = frmWO.Items.Item("2").Specific
            'btn.Caption = "Close"
            '  Dim location = objAddOn.objGenFunc.getSingleValue(" Select U_Location from OUDG where Code=(select DfltsGroup from OUSR where USER_CODE ='" & objAddOn.objCompany.UserName & "')")
            objAddOn.objApplication.SetStatusBarMessage("PO List Loading Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            Dim str_sql As String = ""
            If objAddOn.HANA Then
                str_sql = "WITH ""BOM"" AS (SELECT OWOR.""DocEntry"", ""DocNum"", OWOR.""ItemCode"", ""ItemName"", ""PlannedQty"", ""CmpltQty"", ""Status"" FROM ""OWOR"" " &
                    "INNER JOIN OITM ON OITM.""ItemCode"" = OWOR.""ItemCode"" WHERE OWOR.""DocEntry"" = '" & ParentPOEntry & "'  UNION ALL SELECT T0.""DocEntry"", T0.""DocNum"", T0.""ItemCode"", " &
                    " OITM.""ItemName"", T0.""PlannedQty"", T0.""CmpltQty"", T0.""Status"" FROM OWOR T0 INNER JOIN BOM B ON B.""DocEntry"" = T0.""U_ParentPOEntry"" " &
                    " INNER JOIN OITM ON OITM.""ItemCode"" = B.""ItemCode"") SELECT A.""DocEntry"", A.""DocNum"", A.""ItemCode"", A.""ItemName"", A.""PlannedQty"", A.""CmpltQty"", " &
                    " A.""Status"" FROM BOM A;" ' Recursion is not possible in HANA 
                str_sql = "SELECT OWOR.""DocEntry"", ""DocNum"", OWOR.""ItemCode"", ""ItemName"", ""PlannedQty"", ""CmpltQty"", ""Status"" FROM ""OWOR"" " &
                 "INNER JOIN OITM ON OITM.""ItemCode"" = OWOR.""ItemCode"" WHERE OWOR.""DocEntry"" = '" & ParentPOEntry & "'"
                Dim objDT As SAPbouiCOM.DataTable
                objDT = frmWO.DataSources.DataTables.Item("DataTable")
                objDT.ExecuteQuery(str_sql)
                Dim objRS As SAPbobsCOM.Recordset
                objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim i As Integer = 1
                'While 1 = 1
                '    Dim DocEntry As String
                '    DocEntry = CStr(objDT.GetValue("DocEntry", objDT.Rows.Count - 1))
                '    str_sql = "SELECT OWOR.""DocEntry"", ""DocNum"", OWOR.""ItemCode"", ""ItemName"", ""PlannedQty"", ""CmpltQty"", ""Status"" FROM ""OWOR"" " & _
                ' "INNER JOIN OITM ON OITM.""ItemCode"" = OWOR.""ItemCode"" WHERE OWOR.""U_ParentPOEntry"" = '" & DocEntry & "'"
                '    objRS.DoQuery(str_sql)
                '    If Not objRS.EoF Then
                '        objDT.Rows.Add()
                '        objDT.Columns.Item("DocEntry").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("DocEntry").Value
                '        objDT.Columns.Item("DocNum").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("DocNum").Value
                '        objDT.Columns.Item("ItemCode").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("ItemCode").Value
                '        objDT.Columns.Item("ItemName").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("ItemName").Value
                '        objDT.Columns.Item("PlannedQty").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("PlannedQty").Value
                '        objDT.Columns.Item("CmpltQty").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("CmpltQty").Value
                '        objDT.Columns.Item("Status").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("Status").Value
                '    Else
                '        Exit While
                '    End If
                'End While

                'While 1 = 1
                '    Dim DocEntry As String
                '    DocEntry = CStr(objDT.GetValue("DocEntry", objDT.Rows.Count - 1))
                '    str_sql = "SELECT OWOR.""DocEntry"", ""DocNum"", OWOR.""ItemCode"", ""ItemName"", ""PlannedQty"", ""CmpltQty"", ""Status"" FROM ""OWOR"" " & _
                ' "INNER JOIN OITM ON OITM.""ItemCode"" = OWOR.""ItemCode"" WHERE OWOR.""U_ParentPOEntry"" = '" & DocEntry & "' order by OWOR.""DocNum"" "
                '    objRS.DoQuery(str_sql)
                '    objAddOn.WriteSMSLog("Formload" & str_sql & "Doc" & DocEntry)
                '    If objRS.RecordCount > 1 Then
                '        While Not objRS.EoF
                '            objDT.Rows.Add()
                '            objDT.Columns.Item("DocEntry").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("DocEntry").Value
                '            objDT.Columns.Item("DocNum").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("DocNum").Value
                '            objDT.Columns.Item("ItemCode").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("ItemCode").Value
                '            objDT.Columns.Item("ItemName").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("ItemName").Value
                '            objDT.Columns.Item("PlannedQty").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("PlannedQty").Value
                '            objDT.Columns.Item("CmpltQty").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("CmpltQty").Value
                '            objDT.Columns.Item("Status").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("Status").Value
                '            objRS.MoveNext()
                '        End While
                '    Else
                '        If Not objRS.EoF Then
                '            objDT.Rows.Add()
                '            objDT.Columns.Item("DocEntry").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("DocEntry").Value
                '            objDT.Columns.Item("DocNum").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("DocNum").Value
                '            objDT.Columns.Item("ItemCode").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("ItemCode").Value
                '            objDT.Columns.Item("ItemName").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("ItemName").Value
                '            objDT.Columns.Item("PlannedQty").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("PlannedQty").Value
                '            objDT.Columns.Item("CmpltQty").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("CmpltQty").Value
                '            objDT.Columns.Item("Status").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("Status").Value

                '        Else
                '            Exit While
                '        End If
                '    End If

                'End While


                Dim DocEntry As String
                DocEntry = CStr(objDT.GetValue("DocEntry", objDT.Rows.Count - 1))
                str_sql = "SELECT OWOR.""DocEntry"", ""DocNum"", OWOR.""ItemCode"", ""ItemName"", ""PlannedQty"", ""CmpltQty"", ""Status"" FROM ""OWOR"" " &
             "INNER JOIN OITM ON OITM.""ItemCode"" = OWOR.""ItemCode"" WHERE OWOR.""U_FGPOEntry"" = '" & DocEntry & "' order by OWOR.""DocNum"" "
                objRS.DoQuery(str_sql)
                'objAddOn.WriteSMSLog("Formload" & str_sql & "Doc" & DocEntry)

                While Not objRS.EoF
                    objDT.Rows.Add()
                    objDT.Columns.Item("DocEntry").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("DocEntry").Value
                    objDT.Columns.Item("DocNum").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("DocNum").Value
                    objDT.Columns.Item("ItemCode").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("ItemCode").Value
                    objDT.Columns.Item("ItemName").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("ItemName").Value
                    objDT.Columns.Item("PlannedQty").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("PlannedQty").Value
                    objDT.Columns.Item("CmpltQty").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("CmpltQty").Value
                    objDT.Columns.Item("Status").Cells.Item(objDT.Rows.Count - 1).Value = objRS.Fields.Item("Status").Value
                    objRS.MoveNext()
                End While

                objRS = Nothing
            Else
                str_sql = "with BOM(DocEntry, DocNum, ItemCode,ItemName,PlannedQty,Cmpltqty ,Status) As " &
                      " ( " &
                      " select OWOR.DocEntry, DocNum, OWOR.ItemCode, ItemName ,PlannedQty,Cmpltqty ,Status from owor join OITM on OITM.ItemCode = OWOR.ItemCode " &
                      " where OWOR.DocEntry = '" & ParentPOEntry & "' " &
                      " union all " &
                      " select T0.DocEntry, T0.DocNum, T0.ItemCode, OITM.ItemName,T0.PlannedQty,T0.Cmpltqty ,T0.Status  " &
                      " from OWOR T0 join BOM B  on B.DocEntry = T0.U_ParentPOEntry  join OITM on OITM.ItemCode = B.ItemCode " &
                      " )" &
                      "select A.DocEntry, A.DocNum, A.ItemCode,A.ItemName,A.PlannedQty,A.Cmpltqty ,A.Status from BOM A "

                frmWO.DataSources.DataTables.Item("DataTable").ExecuteQuery(str_sql)

            End If

            oGrid.DataTable = frmWO.DataSources.DataTables.Item("DataTable")

            For i As Integer = 0 To oGrid.Columns.Count - 1
                oGrid.Columns.Item(i).TitleObject.Sortable = True
                oGrid.Columns.Item(i).Editable = False
            Next
            oGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oGrid.RowHeaders.TitleObject.Caption = "#"
            For i As Integer = 0 To oGrid.Rows.Count - 1
                oGrid.RowHeaders.SetText(i, i + 1)
            Next
            Dim col As SAPbouiCOM.EditTextColumn

            col = oGrid.Columns.Item(0)
            col.LinkedObjectType = "202"

            frmWO.Visible = True
            frmWO.Items.Item("t_WONO").Specific.string = ParentPOEntry
            frmWO.Update()
            oGrid.AutoResizeColumns()
            objAddOn.objApplication.SetStatusBarMessage("PO List Loaded...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        End Try
    End Sub

    Sub Sub_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    'Assign Selected Rows
                    Dim oDataTable As SAPbouiCOM.DataTable
                    Dim oCFLE As SAPbouiCOM.ChooseFromListEvent = pVal
                    oDataTable = oCFLE.SelectedObjects
                    Dim rset As SAPbobsCOM.Recordset = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'Filter before open the CFL
                    If pVal.BeforeAction Then
                        'Select Case oCFLE.ChooseFromListUID
                        '    Case "CFL_2A"
                        '        objAddOn.objGenFunc.ChooseFromListFilteration(frmWO, oCFLE.ChooseFromListUID, "CardType", "Select 'C'")
                        '    Case "CFL_2B"
                        '        objAddOn.objGenFunc.ChooseFromListFilteration(frmWO, oCFLE.ChooseFromListUID, "CardType", "Select 'C'")
                        'End Select
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

                                objForm.Items.Item("t_BaseNum").Specific.Value = DocNum
                                objForm.Items.Item("t_BaseEntr").Specific.Value = DocEntry
                                objForm.Items.Item("t_BaseLine").Specific.Value = lineid
                                objForm.Items.Item("t_WObaseLN").Specific.Value = WOBLN

                                Dim Count = objAddOn.objGenFunc.getSingleValue("Select Count(*) +1  from OWOR where U_BaseNum = '" & DocNum & "'  and U_BaseLineId = '" & lineid & "' ")
                                objForm.Items.Item("t_WONo").Specific.Value = DocNum.ToString.Trim + "/" + LineId1.ToString.Trim + "/" + Count.Trim

                                'objform.Items.Item("26").Specific.Value = oGrid.DataTable.GetValue("ProductionDate", i - 1)
                                objForm.Items.Item("6").Specific.Value = oGrid.DataTable.GetValue("ItemCode", pVal.Row)
                                objForm.Items.Item("12").Specific.Value = oGrid.DataTable.GetValue("ReqQty", pVal.Row)
                                Dim format As String
                                Dim result As DateTime
                                Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture
                                format = "yyyymmdd"

                                Dim docdate1 As String = oGrid.DataTable.GetValue("plantdate", pVal.Row)
                                result = Convert.ToDateTime(docdate1,
                                System.Globalization.CultureInfo.GetCultureInfo("hi-in").DateTimeFormat)
                                objForm.Items.Item("26").Specific.string = result
                                ' Dim PostingDate = objAddOn.objGenFunc.getSingleValue(" Select Convert(DateTime,'" & oGrid.DataTable.GetValue("plantDate", pVal.Row) & "') Dt ")
                                ' objform.Items.Item("26").Specific.string = CDate(oGrid.DataTable.GetValue("plantDate", pVal.Row))
                                ' objform.Items.Item("26").Specific.string = CDate(PostingDate)
                                objForm.Items.Item("78").Click()

                                Try
                                    objForm.Items.Item("32").Specific.Value = oGrid.DataTable.GetValue("SONO", pVal.Row)
                                Catch ex As Exception
                                End Try
                                frmWO.Close()
                                Return


                            End If
                        End If
                    Catch ex As Exception
                        objAddOn.objApplication.StatusBar.SetText(" Click Event Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
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

                                    objForm.Items.Item("t_BaseNum").Specific.Value = DocNum
                                    objForm.Items.Item("t_BaseEntr").Specific.Value = DocEntry
                                    objForm.Items.Item("t_BaseLine").Specific.Value = lineid
                                    objForm.Items.Item("t_WObaseLN").Specific.Value = WOBLN

                                    Dim Count = objAddOn.objGenFunc.getSingleValue("Select Count(*) +1  from OWOR where U_BaseNum = '" & DocNum & "'  and U_BaseLineId = '" & lineid & "' ")
                                    objForm.Items.Item("t_WONo").Specific.Value = DocNum.ToString.Trim + "/" + LineId1.ToString.Trim + "/" + Count.Trim

                                    'objform.Items.Item("26").Specific.Value = oGrid.DataTable.GetValue("ProductionDate", i - 1)
                                    objForm.Items.Item("6").Specific.Value = oGrid.DataTable.GetValue("ItemCode", i)
                                    objForm.Items.Item("12").Specific.Value = oGrid.DataTable.GetValue("ReqQty", i)
                                    Dim format As String
                                    Dim result As DateTime
                                    Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture
                                    format = "yyyyMMdd"
                                    '                                Dim docdate1 As String = oGrid.DataTable.GetValue("plantDate", i)
                                    '                                result = Convert.ToDateTime(docdate1,
                                    'System.Globalization.CultureInfo.GetCultureInfo("hi-IN").DateTimeFormat)
                                    '                                objform.Items.Item("26").Specific.string = result
                                    'objform.Items.Item("26").Specific.string = CDate(objAddOn.objGenFunc.getSingleValue("select U_PlantDate from [@MIPL_wko1] WHERE docentry=(select max(DocEntry) from [@MIPL_OWKO] where docnum='" & DocNum & "')"))
                                    'objform.Items.Item("26").Specific.string = CDate(oGrid.DataTable.GetValue("plantDate", i))

                                    objForm.Items.Item("78").Click()

                                    Try
                                        objForm.Items.Item("32").Specific.Value = oGrid.DataTable.GetValue("SONO", i)
                                    Catch ex As Exception
                                    End Try

                                    frmWO.Close()
                                    Return
                                End If
                            Next

                        End If
                    End If
                    'Find btn 
                    'If pVal.ItemUID = "b_Find" Then
                    '    If pVal.BeforeAction = False Then
                    '        Dim c_Orders As SAPbouiCOM.ComboBox = frmWO.Items.Item("c_Orders").Specific
                    '        Dim Orders As String
                    '        Orders = c_Orders.Selected.Value.Trim
                    '        Dim PartNo = frmWO.Items.Item("t_PartNo").Specific.Value
                    '        Dim CardCode = frmWO.Items.Item("t_CardCode").Specific.Value
                    '        Dim CardName = frmWO.Items.Item("t_CardName").Specific.Value
                    '        loadGrid(frmWO.Items.Item("t_WONO").Specific.Value.ToString.Trim, Orders, PartNo, CardCode, CardName)
                    '        Dim oGrid As SAPbouiCOM.Grid = frmWO.Items.Item("Grid").Specific
                    '        For i As Integer = 0 To oGrid.Columns.Count - 1
                    '            oGrid.Columns.Item(i).TitleObject.Sortable = True
                    '        Next
                    '    End If
                    'End If
            End Select

        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText("Item Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub

    Private Function GettingBOM(ByVal ItemCode As String, ByVal Qty As Double) As DataTable
        Try
            Dim StrSql As String
            Dim isBOM As String
            If objDT.Columns.Count = 0 Then
                objDT.Columns.Add("Code", GetType(String))
                objDT.Columns.Add("Qty", GetType(Double))
            End If
            If objAddOn.HANA Then
                isBOM = objAddOn.objGenFunc.getSingleValue("select Case When Count(*) >0 then 'True' else 'False' end From OITT  where ""Code""='" & ItemCode & "' ")
            Else
                isBOM = objAddOn.objGenFunc.getSingleValue("select Case When Count(*) >0 then 'True' else 'False' end From OITT  where Code='" & ItemCode & "' ")
            End If

            If CBool(isBOM) Then
                If objDT.Rows.Count = 0 Then
                    objDT.Rows.Add(ItemCode, Qty)
                End If
                Dim rs As SAPbobsCOM.Recordset = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If objAddOn.HANA Then
                    StrSql = "select ""Code"",""Quantity"" * " & Qty & " ""Qty"" from ITT1  where ""Father""='" & ItemCode & "' "
                Else
                    StrSql = "select Code,Quantity * " & Qty & " [Qty] from ITT1  where Father='" & ItemCode & "' "
                End If
                rs.DoQuery(StrSql)

                For j = 0 To rs.RecordCount - 1
                    Dim isChilBOM As String
                    If objAddOn.HANA Then
                        isChilBOM = objAddOn.objGenFunc.getSingleValue("SELECT CASE WHEN COUNT(*) > 0 THEN 'True' ELSE 'False' END FROM OITT WHERE ""Code"" = '" & rs.Fields.Item("Code").Value & "';")
                    Else
                        isChilBOM = objAddOn.objGenFunc.getSingleValue("select Case When Count(*) >0 then 'True' else 'False' end From OITT  where code='" & rs.Fields.Item("Code").Value & "' ")
                    End If
                    If CBool(isChilBOM) Then
                        Dim qty1 = rs.Fields.Item("Qty").Value
                        objDT.Rows.Add(rs.Fields.Item("Code").Value, rs.Fields.Item("Qty").Value)
                        GettingBOM(rs.Fields.Item("Code").Value, rs.Fields.Item("Qty").Value)
                    End If
                    rs.MoveNext()
                Next
            End If
            Return objDT
        Catch ex As Exception
            GC.Collect()
            objAddOn.objApplication.StatusBar.SetText("Production Order Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Return Nothing
            ' Finally
        End Try
    End Function

End Class
