Public Class VENTA_GRAUS
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If RadioButton3.Checked = True Then
            Registro_Ventas.TextBox16.Text = "0005"
            Registro_Ventas.TextBox17.Text = "NOTEX S"
            Registro_Ventas.TextBox13.Text = "NOTEX S"
            Registro_Ventas.TextBox24.Text = RadioButton3.Text
            Registro_Ventas.Label26.Text = Label2.Text
            Registro_Ventas.Show()
        Else
            If RadioButton11.Checked = True Then
                Registro_Ventas.TextBox16.Text = "0038"
                Registro_Ventas.TextBox17.Text = "HILADO"
                Registro_Ventas.TextBox13.Text = "HILADO"
                Registro_Ventas.TextBox24.Text = RadioButton3.Text
                Registro_Ventas.Label26.Text = Label2.Text
                Registro_Ventas.Show()
            End If
        End If
        Registro_Ventas.Label17.Text = Label1.Text
    End Sub
End Class