Public Class Guia_Factura
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim ft As New vfacturasistema
        Dim gy As New ffacturavianny
        ft.gccia = Label5.Text
        ft.gsguia = TextBox1.Text
        ft.gcguia = TextBox2.Text
        ft.gsfactura = TextBox3.Text
        ft.gcfactura = TextBox4.Text
        gy.actualizar_guiafactura(ft)
        MsgBox("LA ACTUALIZACION SE REALIZO CON EXITO")
        TextBox2.Text = ""
        TextBox4.Text = ""
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case TextBox2.Text.Length

                Case "1" : TextBox2.Text = "0000000" & "" & TextBox2.Text
                Case "2" : TextBox2.Text = "000000" & "" & TextBox2.Text
                Case "3" : TextBox2.Text = "00000" & "" & TextBox2.Text
                Case "4" : TextBox2.Text = "0000" & "" & TextBox2.Text
                Case "5" : TextBox2.Text = "000" & "" & TextBox2.Text
                Case "6" : TextBox2.Text = "00" & "" & TextBox2.Text
                Case "7" : TextBox2.Text = "0" & "" & TextBox2.Text
                Case "8" : TextBox2.Text = TextBox2.Text

            End Select
        End If
    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case TextBox4.Text.Length

                Case "1" : TextBox4.Text = "0000000" & "" & TextBox4.Text
                Case "2" : TextBox4.Text = "000000" & "" & TextBox4.Text
                Case "3" : TextBox4.Text = "00000" & "" & TextBox4.Text
                Case "4" : TextBox4.Text = "0000" & "" & TextBox4.Text
                Case "5" : TextBox4.Text = "000" & "" & TextBox4.Text
                Case "6" : TextBox4.Text = "00" & "" & TextBox4.Text
                Case "7" : TextBox4.Text = "0" & "" & TextBox4.Text
                Case "8" : TextBox4.Text = TextBox4.Text

            End Select
        End If
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
End Class