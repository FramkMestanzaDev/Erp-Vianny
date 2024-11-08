Public Class Reporte_ventasn
    Dim DT As New DataTable

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.F1 Then
            Clientes.TextBox3.Text = "88"
            Clientes.Show()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim KJ As New fventasn
        Dim KJ1 As New vventas_n

        DataGridView1.DataSource = ""
        KJ1.gcliente = Trim(TextBox1.Text)
        DT = KJ.buscar_ventasNN2(KJ1)
        DataGridView1.DataSource = DT
        DataGridView1.Columns(6).Width = 350
    End Sub

    Private Sub Reporte_ventasn_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Rpts_Cobranzas.TextBox1.Text = TextBox1.Text
        Rpts_Cobranzas.Show()
    End Sub
End Class