Public Class RUBRO_GRAUS
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If RadioButton1.Checked = True Then
            Registro_Facturas.Label4.Text = "0038"
            Registro_Facturas.TextBox13.Text = RadioButton1.Text
            Registro_Facturas.Label27.Text = Label3.Text
            Registro_Facturas.TextBox13.Text = RadioButton1.Text
        Else
            If RadioButton2.Checked = True Then
                Registro_Facturas.Label4.Text = "0038"
                Registro_Facturas.TextBox13.Text = RadioButton2.Text
                Registro_Facturas.Label27.Text = Label3.Text
                Registro_Facturas.TextBox13.Text = RadioButton2.Text
            End If
        End If
        Registro_Facturas.TextBox8.Text = "ALMACEN" & "" & Label1.Text & "-  REGISTRO DE VENTA  "
        Registro_Facturas.Label5.Text = Label1.Text
        Registro_Facturas.Label29.Text = Label4.Text
        Registro_Facturas.TextBox10.Text = Label1.Text
        Registro_Facturas.Show()
    End Sub
End Class