Public Class ProductosLinea
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        pproductos.CCIA.Text = Label2.Text
        pproductos.ALMACEN.Text = "10"
        pproductos.ANO.Text = Label1.Text
        pproductos.FECHA.Text = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
        pproductos.TextBox1.Text = "10"
        pproductos.Label3.Text = 6
        pproductos.Show()
    End Sub
End Class