Public Class FormLpn
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Function validar() As Boolean
        If TextBox1.Text.ToString().Trim.Length = 0 Then
            MsgBox("Falta Ingrsar la Orden de Compra")
            Return False
        End If
        If TextBox2.Text.ToString().Trim.Length = 0 Then
            MsgBox("Falta Ingrsar la Cita")
            Return False
        End If
        Return True
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If validar() = True Then
            RptLpnSaga.TextBox1.Text = TextBox1.Text.ToString().Trim()
            RptLpnSaga.TextBox2.Text = TextBox2.Text.ToString().Trim()
            RptLpnSaga.Show(Me)
        End If
    End Sub
End Class