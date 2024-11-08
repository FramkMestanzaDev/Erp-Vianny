Public Class Registro_Celulares
    Dim DT As New DataTable
    Private Sub Registro_Celulares_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ComboBox3.SelectedIndex = 0
        ComboBox4.SelectedIndex = 0
        ComboBox5.SelectedIndex = 0
        ComboBox6.SelectedIndex = 0
        TextBox7.Select()
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim fg As New vpartida
            Dim fg1 As New fpartida
            Select Case TextBox1.Text.Length
                Case "1" : TextBox1.Text = "0000" & "" & TextBox1.Text
                Case "2" : TextBox1.Text = "000" & "" & TextBox1.Text
                Case "3" : TextBox1.Text = "00" & "" & TextBox1.Text
                Case "4" : TextBox1.Text = "0" & "" & TextBox1.Text
                Case "5" : TextBox1.Text = TextBox1.Text
            End Select
            Select Case TextBox1.Text.Length
                Case "1" : fg.gcodigo = "0000" & "" & TextBox1.Text
                Case "2" : fg.gcodigo = "000" & "" & TextBox1.Text
                Case "3" : fg.gcodigo = "00" & "" & TextBox1.Text
                Case "4" : fg.gcodigo = "0" & "" & TextBox1.Text
                Case "5" : fg.gcodigo = TextBox1.Text
            End Select

            TextBox2.Text = fg1.buscar_persona(fg)
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim jk As New fmoviles
        Dim jkl As New vmoviles
        Dim kk As New vmoviles
        kk.glinea = TextBox7.Text
        jk.eliminar_movil(kk)
        jkl.gproveedor = ComboBox6.Text
        jkl.glinea = TextBox7.Text
        jkl.gcodigo = TextBox1.Text
        jkl.gtrabajador = TextBox2.Text
        jkl.garea = ComboBox1.Text
        jkl.gmequipo = TextBox3.Text
        jkl.gmoequipo = TextBox4.Text
        jkl.gpreequipo = TextBox5.Text
        jkl.gcuotas = ComboBox2.Text
        jkl.gserie = TextBox6.Text
        jkl.gplam = ComboBox3.Text
        jkl.gmodalidad = ComboBox4.Text
        jkl.gempresa = ComboBox5.Text
        jkl.gfechas = DateTimePicker1.Value
        jk.insertar_moviles(jkl)
        MsgBox("SE REGISTRO LA INFORMACION SOLICITADA")
        Reporte_Linea.TextBox1.Text = TextBox7.Text
        Reporte_Linea.Show()
        blaquear()
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim jas As New fmoviles
            Dim ll As New vmoviles
            ComboBox6.Select()
            ll.glinea = TextBox7.Text
            DT = jas.mostrar_lineas(ll)
            If DT.Rows.Count <> 0 Then
                DataGridView1.DataSource = DT

                ComboBox6.Text = DataGridView1.Rows(0).Cells(0).Value
                TextBox7.Text = DataGridView1.Rows(0).Cells(1).Value
                TextBox1.Text = DataGridView1.Rows(0).Cells(2).Value
                TextBox2.Text = DataGridView1.Rows(0).Cells(3).Value
                ComboBox1.Text = DataGridView1.Rows(0).Cells(4).Value
                TextBox3.Text = DataGridView1.Rows(0).Cells(5).Value
                TextBox4.Text = DataGridView1.Rows(0).Cells(6).Value
                TextBox5.Text = DataGridView1.Rows(0).Cells(7).Value
                ComboBox2.Text = DataGridView1.Rows(0).Cells(8).Value
                TextBox6.Text = DataGridView1.Rows(0).Cells(9).Value
                ComboBox3.Text = DataGridView1.Rows(0).Cells(10).Value
                ComboBox3.Text = DataGridView1.Rows(0).Cells(11).Value
                ComboBox4.Text = DataGridView1.Rows(0).Cells(12).Value
                ComboBox5.Text = DataGridView1.Rows(0).Cells(13).Value
                Button2.Enabled = True
                Button1.Enabled = False
                Button4.Visible = True
                Button4.Enabled = True
                bloquear()
            Else
                habilitar()
                Button2.Enabled = False
                Button1.Enabled = True
            End If
        End If
    End Sub
    Sub blaquear()

        TextBox7.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""

        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
    End Sub
    Sub bloquear()

        TextBox7.Enabled = False
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False
        ComboBox5.Enabled = False
        ComboBox6.Enabled = False
    End Sub
    Sub habilitar()

        TextBox7.Enabled = True
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        ComboBox3.Enabled = True
        ComboBox4.Enabled = True
        ComboBox5.Enabled = True
        ComboBox6.Enabled = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim respuesta As DialogResult


        respuesta = MessageBox.Show("QUIERES ELIMINAR EL REGISTRO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            Dim jk As New fmoviles

            Dim kk As New vmoviles
            kk.glinea = TextBox7.Text
            jk.eliminar_movil(kk)
            MsgBox("SE ELIMINO LO SOLICITADO")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Button4.Visible = True
        Button1.Enabled = True
        habilitar()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Reporte_Linea.TextBox1.Text = TextBox7.Text
        Reporte_Linea.Show()
    End Sub
End Class