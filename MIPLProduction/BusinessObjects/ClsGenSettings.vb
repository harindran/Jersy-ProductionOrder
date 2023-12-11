Public Class ClsGenSettings
    Public Const frmType As String = "138"
    Dim objForm As SAPbouiCOM.Form

    Public Sub ItemEvent(FormUID As String, pval As SAPbouiCOM.ItemEvent, BubbleEvent As System.Boolean)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        If pval.BeforeAction = True Then
            Select Case pval.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    FieldCreationUI(FormUID)
            End Select
        Else
            Select Case pval.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

            End Select


        End If
    End Sub

    Private Sub FieldCreationUI(FormUID As String)
        Dim oChkBox As SAPbouiCOM.CheckBox
        Dim oItem As SAPbouiCOM.Item
        Try
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            oItem = objForm.Items.Add("ChkPO", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
            oItem.Left = objForm.Items.Item("234000069").Left ' 594 234000037
            oItem.Width = 355
            oItem.Height = 18
            oItem.Top = objForm.Items.Item("234000069").Top + objForm.Items.Item("234000069").Height + 2 '321
            oItem.LinkTo = "234000069"
            oChkBox = oItem.Specific
            oChkBox.Caption = "Enable Consolidated Production Order Creation"
            oChkBox.DataBind.SetBound(True, "OADM", "U_POCon")
            oChkBox.Item.FromPane = 4
            oChkBox.Item.ToPane = 4
        Catch ex As Exception

        End Try
    End Sub
End Class
