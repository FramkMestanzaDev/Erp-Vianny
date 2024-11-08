Public Class Form7
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox2.Enabled = True
            CheckBox2.Checked = False
            TextBox2.Text = ""
        Else
            TextBox2.Enabled = False

        End If
        TextBox2.Select()
    End Sub



    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case Trim(TextBox3.Text).Length
                Case "1" : TextBox3.Text = "01" & "0000000" & TextBox3.Text
                Case "2" : TextBox3.Text = "01" & "000000" & TextBox3.Text
                Case "3" : TextBox3.Text = "01" & "00000" & TextBox3.Text
                Case "4" : TextBox3.Text = "01" & "0000" & TextBox3.Text
                Case "5" : TextBox3.Text = "01" & "000" & TextBox3.Text
                Case "6" : TextBox3.Text = "01" & "00" & TextBox3.Text
                Case "7" : TextBox3.Text = "01" & "0" & TextBox3.Text
                Case "8" : TextBox3.Text = "01" & TextBox3.Text
                Case "9" : TextBox3.Text = "0" & TextBox3.Text
            End Select
            Button1.Select()
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If CheckBox1.Checked = True And CheckBox3.Checked = True Then
            Status_talla.TextBox1.Text = Trim(TextBox2.Text)
            Status_talla.TextBox2.Text = Trim(TextBox4.Text)
            Status_talla.Show()
        Else
            If CheckBox1.Checked = True And CheckBox3.Checked = False Then
                status.TextBox1.Text = Trim(TextBox2.Text)
                status.TextBox2.Text = Trim(TextBox4.Text)
                status.Show()
            Else
                If CheckBox2.Checked = True And CheckBox3.Checked = True Then
                    Status_talla_op.TextBox1.Text = Trim(TextBox3.Text)
                    Status_talla_op.TextBox2.Text = Trim(TextBox4.Text)
                    Status_talla_op.Show()
                Else
                    If CheckBox2.Checked = True And CheckBox3.Checked = False Then
                        Status_op_simple.TextBox1.Text = Trim(TextBox3.Text)
                        Status_op_simple.TextBox2.Text = Trim(TextBox4.Text)
                        Status_op_simple.Show()
                    End If
                End If
            End If

        End If

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            TextBox3.Enabled = True
            CheckBox1.Checked = False
        Else
            TextBox3.Text = ""
            TextBox3.Enabled = False
        End If

        TextBox3.Select()

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown


    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub
End Class