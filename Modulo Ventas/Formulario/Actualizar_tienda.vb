Public Class Actualizar_tienda
    Private Sub Actualizar_tienda_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = DateTimePicker1.Value.Month - 1
        If ComboBox1.Text = "VIANNY" Then
            TextBox2.Text = "01"
            TextBox3.Text = "B001"
            TextBox4.Text = "F007 "
            TextBox5.Text = "32"
        Else
            If ComboBox1.Text = "GRAUS" Then
                TextBox2.Text = "02"
                TextBox3.Text = "B001"
                TextBox4.Text = "F001"
                TextBox5.Text = "32"
            End If
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "VIANNY" Then
            TextBox2.Text = "01"
            TextBox3.Text = "B001"
            TextBox4.Text = "F007"
            TextBox5.Text = "32"
        Else
            If ComboBox1.Text = "GRAUS" Then
                TextBox2.Text = "02"
                TextBox3.Text = "B012"
                TextBox4.Text = "F012"
                TextBox5.Text = "32"
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim fg As New fcontailidad
        Dim gh As New vcontablidad

        Select Case ComboBox2.Text
            Case "ENERO" : gh.gMES = "01"
            Case "FEBRERO" : gh.gMES = "02"
            Case "MARZO" : gh.gMES = "03"
            Case "ABRIL" : gh.gMES = "04"
            Case "MAYO" : gh.gMES = "05"
            Case "JUNIO" : gh.gMES = "06"
            Case "JULIO" : gh.gMES = "07"
            Case "AGOSTO" : gh.gMES = "08"
            Case "SETIEMBRE" : gh.gMES = "09"
            Case "OCTUBRE" : gh.gMES = "10"
            Case "NOVIEMBRE" : gh.gMES = "11"
            Case "DICIEMBRE" : gh.gMES = "12"
        End Select

        gh.gcia = TextBox2.Text
        gh.gtienda = TextBox5.Text
        gh.gserieb = TextBox3.Text
        gh.gserief = TextBox4.Text
        gh.gperiodo = TextBox1.Text
        If Trim(TextBox2.Text) = "01" Then
            fg.actualizar_tienda(gh)
        Else
            fg.actualizar_tienda_graus(gh)
        End If

        MsgBox("SE ACTUALIZO TODOS LOS REGISTROS DEL MES")
    End Sub
End Class