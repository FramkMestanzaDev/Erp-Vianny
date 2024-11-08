Public Class Analisis_movimientos
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Select Case ComboBox1.Text

            Case "CORTE" : TextBox2.Text = "0301"
            Case "APLICACIONES Y PIEZAS" : TextBox2.Text = "0701"
            Case "CONFECCION" : TextBox2.Text = "0421"
            Case "LAVANDERIA" : TextBox2.Text = "0601"
            Case "ACABADO" : TextBox2.Text = "0522"
        End Select


    End Sub

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

        Movimiento_Produccion.TextBox1.Text = TextBox1.Text
        Movimiento_Produccion.TextBox2.Text = TextBox2.Text
        Movimiento_Produccion.TextBox3.Text = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
        Movimiento_Produccion.TextBox4.Text = Trim(TextBox3.Text)
        If CheckBox2.Checked = True Then
            Movimiento_Produccion.TextBox5.Text = "1"
        Else
            Movimiento_Produccion.TextBox5.Text = "2"
        End If
        Movimiento_Produccion.TextBox6.Text = Label4.Text
        Movimiento_Produccion.Show()

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
        End If




    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox3.Enabled = True
        Else
            TextBox3.Enabled = False
        End If
    End Sub

    Private Sub Analisis_movimientos_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class