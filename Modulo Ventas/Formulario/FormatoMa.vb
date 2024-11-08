Public Class FormatoMa
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Formato_solic_avios.Label7.Text = Label1.Text
        Formato_solic_avios.Label15.Text = Label2.Text
        If RadioButton1.Checked = True Then
            Formato_solic_avios.Label16.Text = "CONFECCION"
            Formato_solic_avios.CheckBox1.Checked = True
            Formato_solic_avios.TextBox1.Text = "AC"
        Else

            If RadioButton2.Checked = True Then
                Formato_solic_avios.Label16.Text = "ACABADOS"
                Formato_solic_avios.CheckBox2.Checked = True
                Formato_solic_avios.TextBox1.Text = "AA"

            End If
        End If

        Formato_solic_avios.Label17.Text = Label3.Text

        Formato_solic_avios.Show()
    End Sub

    Private Sub FormatoMa_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button1.Select()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Button1.Select()
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        Button1.Select()
    End Sub
End Class