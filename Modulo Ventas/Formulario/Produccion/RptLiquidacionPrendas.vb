Public Class RptLiquidacionPrendas
    Private Sub button1_Click(sender As Object, e As EventArgs) Handles button1.Click
        Dim lp As New RptLiquidacionppp

        Select Case comboBox1.Text.ToString().Trim()
            Case "ENERO" : lp.TextBox1.Text = "01"
            Case "FEBRERO" : lp.TextBox1.Text = "02"
            Case "MARZO" : lp.TextBox1.Text = "03"
            Case "ABRIL" : lp.TextBox1.Text = "04"
            Case "MAYO" : lp.TextBox1.Text = "05"
            Case "JUNIO" : lp.TextBox1.Text = "06"
            Case "JULIO" : lp.TextBox1.Text = "07"
            Case "AGOSTO" : lp.TextBox1.Text = "08"
            Case "SETIEMBRE" : lp.TextBox1.Text = "09"
            Case "OCTUBRE" : lp.TextBox1.Text = "10"
            Case "NOVIEMBRE" : lp.TextBox1.Text = "11"
            Case "DICIEMBRE" : lp.TextBox1.Text = "12"
        End Select

        lp.textBox2.Text = textBox3.Text.ToString().Trim()
        lp.textBox3.Text = textBox2.Text.ToString().Trim()
        lp.ShowDialog()
    End Sub
End Class