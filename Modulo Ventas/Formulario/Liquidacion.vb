Public Class Liquidacion
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Select Case Trim(ComboBox1.Text)
            Case "CORTE" : Rpt_liquidaciones.Label7.Text = "0301"
            Case "APLICACIONES Y PIEZAS" : Rpt_liquidaciones.Label7.Text = "0701"
            Case "CONFECCION" : Rpt_liquidaciones.Label7.Text = "0421"
            Case "LAVANDERIA" : Rpt_liquidaciones.Label7.Text = "0601"
            Case "ACABADOS" : Rpt_liquidaciones.Label7.Text = "0522"
        End Select
        Select Case Trim(ComboBox1.Text)
            Case "CORTE" : Rpt_liquidaciones.Label6.Text = "0300"
            Case "APLICACIONES Y PIEZAS" : Rpt_liquidaciones.Label6.Text = "0700"
            Case "CONFECCION" : Rpt_liquidaciones.Label6.Text = "0400"
            Case "LAVANDERIA" : Rpt_liquidaciones.Label6.Text = "0600"
            Case "ACABADOS" : Rpt_liquidaciones.Label6.Text = "0500"
        End Select
        Rpt_liquidaciones.TextBox2.Text = Trim(TextBox2.Text)
        Rpt_liquidaciones.TextBox3.Text = Trim(Label3.Text)
        Rpt_liquidaciones.TextBox4.Text = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
        Rpt_liquidaciones.Label5.Text = Trim(Label4.Text)
        Rpt_liquidaciones.ShowDialog()
    End Sub
End Class