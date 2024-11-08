Public Class FormLpnRipley
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
            If Label3.Text = 1 Then
                RptLpnRojo.TextBox1.Text = TextBox1.Text.ToString().Trim()
                RptLpnRojo.TextBox2.Text = TextBox2.Text.ToString().Trim()
                RptLpnRojo.Show(Me)
            Else
                RptLpnBlanco.TextBox1.Text = TextBox1.Text.ToString().Trim()
                RptLpnBlanco.TextBox2.Text = TextBox2.Text.ToString().Trim()
                RptLpnBlanco.Show(Me)
            End If

        End If
    End Sub
End Class