Public Class Form1
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim va As New fusuario
        Dim va1 As New vusuario

        va1.gid_usuario = TextBox1.Text
        va1.gnombre = TextBox2.Text
        va1.gusuario = TextBox3.Text
        va1.gclave = TextBox4.Text
        va1.ggrupo = ComboBox1.Text
        va1.gcorreo = TextBox5.Text
        va.insertar_usuario(va1)
        MsgBox("SE REGISTRO CORRECTAMENTE LA INFORMACION")
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case TextBox1.Text.Length


                Case "1" : TextBox1.Text = "000" & "" & TextBox1.Text
                Case "2" : TextBox1.Text = "00" & "" & TextBox1.Text
                Case "3" : TextBox1.Text = "0" & "" & TextBox1.Text
                Case "4" : TextBox1.Text = TextBox1.Text
            End Select
        End If

    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click

    End Sub
End Class