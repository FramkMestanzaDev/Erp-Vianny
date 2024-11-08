Public Class Reporte_calidad_aca
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Select Case Trim(TextBox1.Text).Length
            Case "1" : TextBox1.Text = "000000000" & TextBox1.Text
            Case "2" : TextBox1.Text = "00000000" & TextBox1.Text
            Case "3" : TextBox1.Text = "0000000" & TextBox1.Text
            Case "4" : TextBox1.Text = "000000" & TextBox1.Text
            Case "5" : TextBox1.Text = "00000" & TextBox1.Text
            Case "6" : TextBox1.Text = "0000" & TextBox1.Text
            Case "7" : TextBox1.Text = "000" & TextBox1.Text
            Case "8" : TextBox1.Text = "00" & TextBox1.Text
            Case "9" : TextBox1.Text = "0" & TextBox1.Text
            Case "10" : TextBox1.Text = TextBox1.Text
        End Select
        Rpt_Calidad_acabado.TextBox1.Text = Trim(TextBox1.Text)

        Rpt_Calidad_acabado.Show()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case Trim(TextBox1.Text).Length
                Case "1" : TextBox1.Text = "000000000" & TextBox1.Text
                Case "2" : TextBox1.Text = "00000000" & TextBox1.Text
                Case "3" : TextBox1.Text = "0000000" & TextBox1.Text
                Case "4" : TextBox1.Text = "000000" & TextBox1.Text
                Case "5" : TextBox1.Text = "00000" & TextBox1.Text
                Case "6" : TextBox1.Text = "0000" & TextBox1.Text
                Case "7" : TextBox1.Text = "000" & TextBox1.Text
                Case "8" : TextBox1.Text = "00" & TextBox1.Text
                Case "9" : TextBox1.Text = "0" & TextBox1.Text
                Case "10" : TextBox1.Text = TextBox1.Text
            End Select
        End If

    End Sub
End Class