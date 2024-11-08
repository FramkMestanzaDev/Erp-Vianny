Public Class Tipo_Venta
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If RadioButton2.Checked = True Then
            Orden_Produccion.TextBox9.Text = "20508740361"
            Orden_Produccion.TextBox2.Text = "CONSORCIO TEXTIL VIANNY"
            Orden_Produccion.TextBox20.Text = "TELA"
            Orden_Produccion.TextBox11.Text = MDIParent1.Label3.Text
            Orden_Produccion.Label29.Text = 1
            Orden_Produccion.Label30.Text = Label1.Text
            Orden_Produccion.Label33.Text = Label3.Text
            Orden_Produccion.Show()
        Else
            If RadioButton1.Checked = True Then
                Orden_Produccion.TextBox9.Text = "20508740361"
                Orden_Produccion.TextBox2.Text = "CONSORCIO TEXTIL VIANNY"

                Orden_Produccion.TextBox11.Text = MDIParent1.Label3.Text
                Orden_Produccion.TextBox20.Text = "SERVICIO"
                Orden_Produccion.Label29.Text = 2
                Orden_Produccion.Label30.Text = Label1.Text
                Orden_Produccion.Label33.Text = Label3.Text
                Orden_Produccion.Show()
            Else
                If RadioButton3.Checked = True Then
                    'Op_Manufactura.TextBox22.Text = Productos_Vianny.TextBox5.Text
                    'Op_Manufactura.TextBox2.Text = Productos_Vianny.TextBox6.Text
                    Op_Manufactura.TextBox10.Text = Label2.Text
                    Op_Manufactura.Label33.Text = Label3.Text
                    Op_Manufactura.Show()
                Else
                    Orden_Produccion.TextBox9.Text = "20508740361"
                    Orden_Produccion.TextBox2.Text = "CONSORCIO TEXTIL VIANNY"

                    Orden_Produccion.TextBox11.Text = MDIParent1.Label3.Text
                    Orden_Produccion.TextBox20.Text = "MUESTRA"
                    Orden_Produccion.Label29.Text = 3
                    Orden_Produccion.Label30.Text = Label1.Text
                    Orden_Produccion.Label33.Text = Label3.Text
                    Orden_Produccion.Show()
                End If
            End If
        End If
    End Sub

    Private Sub Tipo_Venta_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class