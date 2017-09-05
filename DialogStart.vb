Imports System.Windows.Forms

Public Class DialogStart

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        If txtOperator.Text.Length = 0 Then
            MsgBox("Please Enter operator name.")
            Exit Sub
        ElseIf txtWO.Text.Length = 0 Then
            MsgBox("Please Scan Traveller")
            Exit Sub
        End If

        Call BarDecode(txtWO.Text)
        frmMain.txtOperator.Text = txtOperator.Text
        frmMain.txtLot.Text = id_wo
        'BaseFN = "CLC-PLF-" & id_fn

        'dBoxFactory += "\" & BaseFN
        Me.Close()
    End Sub

    Public Sub BarDecode(ByVal barscan As String)

        id_wo = barscan
        BaseFN = barscan
        dBoxFactory += "\" & BaseFN
        'id_ptype = barscan.Substring(10, 1)
        id_ptype = CInt(barscan.Substring(9, 1))

        If id_ptype = 0 Then
            id_cri = 95
        Else
            id_cri = id_ptype * 10
        End If

        id_cct = CInt(barscan.Substring(10, 2)) * 100
        'BinType = id_cct & "K"
        id_fn = (id_ptype) & (id_cct / 100)
    End Sub

End Class

