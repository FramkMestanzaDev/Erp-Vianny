Public Class ObsMatriz
    Private Sub ObsMatriz_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        If Label2.Text = "1" Then
            Matriz_Avios.DataGridView1.Rows(Label1.Text).Cells(11).Value = TextBox1.Text.ToString().Trim()
        Else
            If Label2.Text = "3" Then
                MatrizAviosOd.DataGridView1.Rows(Label1.Text).Cells(11).Value = TextBox1.Text.ToString().Trim()
            Else

                DetalleMatrizAvios.DataGridView1.Rows(Label1.Text).Cells(6).Value = TextBox1.Text.ToString().Trim()
            End If
        End If
    End Sub

    Private Sub ObsMatriz_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class