Public Class PEDIDO
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim fecini, fecfin As String
        fecini = Format(DateTimePicker1.Value, "yyyy-MM-dd")
        fecfin = Format(DateTimePicker2.Value, "yyyy-MM-dd")
        RPT_PEDIDOS.TextBox1.Text = (Replace(fecini, "-", ""))
        RPT_PEDIDOS.TextBox2.Text = (Replace(fecfin, "-", ""))

        RPT_PEDIDOS.Show()
    End Sub
End Class