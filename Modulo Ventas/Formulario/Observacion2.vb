Public Class Observacion2
    Private Sub Observacion2_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        If Label1.Text = "1" Then

        Else
            Cancelar_Factura.DataGridView1.Rows(TextBox2.Text).Cells(13).Value = TextBox1.Text
        End If

    End Sub

    Private Sub Observacion2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class