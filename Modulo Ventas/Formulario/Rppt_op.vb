Public Class Rppt_op
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Select Case TextBox1.Text.Length

            Case "1" : TextBox1.Text = "010000000" & "" & TextBox1.Text
            Case "2" : TextBox1.Text = "01000000" & "" & TextBox1.Text
            Case "3" : TextBox1.Text = "0100000" & "" & TextBox1.Text
            Case "4" : TextBox1.Text = "010000" & "" & TextBox1.Text
            Case "5" : TextBox1.Text = "01000" & "" & TextBox1.Text
            Case "6" : TextBox1.Text = "0100" & "" & TextBox1.Text
            Case "7" : TextBox1.Text = "010" & "" & TextBox1.Text
            Case "8" : TextBox1.Text = "01" & "" & TextBox1.Text
            Case "9" : TextBox1.Text = "0" & TextBox1.Text
        End Select
        If Label3.Text = "01" Then
            Rpt_Op.TextBox1.Text = Label3.Text
            Rpt_Op.TextBox2.Text = TextBox1.Text
            Rpt_Op.TextBox3.Text = TextBox1.Text
            Rpt_Op.TextBox4.Text = 0
            Rpt_Op.TextBox5.Text = 1
            Rpt_Op.Show()
        Else
            RpOpGraus.TextBox1.Text = Label3.Text
            RpOpGraus.TextBox2.Text = TextBox1.Text
            RpOpGraus.TextBox3.Text = TextBox1.Text
            RpOpGraus.TextBox4.Text = 0
            RpOpGraus.TextBox5.Text = 0
            RpOpGraus.Show()
        End If

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = Keys.Enter Then
            Select Case TextBox1.Text.Length

                Case "1" : TextBox1.Text = "010000000" & "" & TextBox1.Text
                Case "2" : TextBox1.Text = "01000000" & "" & TextBox1.Text
                Case "3" : TextBox1.Text = "0100000" & "" & TextBox1.Text
                Case "4" : TextBox1.Text = "010000" & "" & TextBox1.Text
                Case "5" : TextBox1.Text = "01000" & "" & TextBox1.Text
                Case "6" : TextBox1.Text = "0100" & "" & TextBox1.Text
                Case "7" : TextBox1.Text = "010" & "" & TextBox1.Text
                Case "8" : TextBox1.Text = "01" & "" & TextBox1.Text
                Case "9" : TextBox1.Text = "0" & TextBox1.Text
            End Select
            Button1.Select()
        End If
    End Sub
    Public Sub NumConFrac(ByVal CajaTexto As Windows.Forms.TextBox, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Char.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False

        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Try
            NumConFrac(TextBox1, e)
        Catch ex As Exception
            MsgBox("FALTA INGRESAR UN NUMERO")
        End Try
    End Sub
End Class