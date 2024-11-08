Public Class observaciones_vn
    Private Sub Observaciones_vn_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Form3.DataGridView1.Rows(TextBox2.Text).Cells(18).Value = TextBox1.Text
    End Sub

    Private Sub Observaciones_vn_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class