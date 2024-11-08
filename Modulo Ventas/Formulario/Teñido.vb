Public Class Teñido
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox3.Text = "" And TextBox4.Text = "" And TextBox5.Text = "" And TextBox6.Text = "" And TextBox7.Text = "" Then
            MsgBox("EXISTEN DATOS EN BLANCO OBLIGATORIOS")
        Else
            DataGridView1.Rows.Add(1)
            Dim J As Integer
            J = DataGridView1.Rows.Count
            DataGridView1.Rows(J - 1).Cells(0).Value = DateTime.Now.ToString("hh:mm")
            If J > 1 Then
                Dim k As Integer
                Dim jh As New vtenido
                Dim jh1, jh2 As New vtenido
                Dim bb As New ftenido
                k = DataGridView1.Rows.Count
                jh1.gntenido = Trim(TextBox7.Text)
                jh1.gfecha = DateTimePicker1.Value
                jh1.gmetraje = TextBox6.Text
                jh1.garticulo = TextBox3.Text
                jh1.glote = TextBox4.Text
                jh1.gnhoja = TextBox5.Text
                jh1.gestado = 1
                jh2.gntenido = TextBox7.Text
                bb.eliminar_cabecer_tenido(jh2)
                bb.insertar_tenido_cab(jh1)
                For i = k - 2 To k - 2
                    'DataGridView1.Rows(i).Frozen = True
                    jh.gntenido = Trim(TextBox7.Text)
                    jh.ghora = DataGridView1.Rows(i).Cells(0).Value
                    If DataGridView1.Rows(i).Cells(1).Value = "0.00" Then
                        jh.gph = 0
                    Else
                        jh.gph = DataGridView1.Rows(i).Cells(1).Value
                    End If
                    If DataGridView1.Rows(i).Cells(2).Value = "0.00" Then
                        jh.gmv = 0
                    Else
                        jh.gmv = DataGridView1.Rows(i).Cells(2).Value
                    End If
                    If DataGridView1.Rows(i).Cells(3).Value = "0.00" Then
                        jh.gindigo = 0
                    Else
                        jh.gindigo = DataGridView1.Rows(i).Cells(3).Value
                    End If
                    If DataGridView1.Rows(i).Cells(4).Value = "0.00" Then
                        jh.ghidro = 0
                    Else
                        jh.ghidro = DataGridView1.Rows(i).Cells(4).Value
                    End If
                    If DataGridView1.Rows(i).Cells(5).Value = "0.00" Then
                        jh.gbe = 0
                    Else
                        jh.gbe = DataGridView1.Rows(i).Cells(5).Value
                    End If
                    If DataGridView1.Rows(i).Cells(6).Value = "0.00" Then
                        jh.gveloc = 0
                    Else
                        jh.gveloc = DataGridView1.Rows(i).Cells(6).Value
                    End If
                    If DataGridView1.Rows(i).Cells(7).Value = "0.00" Then
                        jh.gtinas = 0
                    Else
                        jh.gtinas = DataGridView1.Rows(i).Cells(7).Value
                    End If
                    If DataGridView1.Rows(i).Cells(8).Value = "0.00" Then
                        jh.gds_colorant = 0
                    Else
                        jh.gds_colorant = DataGridView1.Rows(i).Cells(8).Value
                    End If
                    If DataGridView1.Rows(i).Cells(9).Value = "0.00" Then
                        jh.gds_hidrosuf = 0
                    Else
                        jh.gds_hidrosuf = DataGridView1.Rows(i).Cells(9).Value
                    End If
                    If DataGridView1.Rows(i).Cells(10).Value = "0.00" Then
                        jh.gds_soda = 0
                    Else
                        jh.gds_soda = DataGridView1.Rows(i).Cells(10).Value
                    End If
                    If DataGridView1.Rows(i).Cells(11).Value = "0.00" Then
                        jh.gds_pquimic = 0
                    Else
                        jh.gds_pquimic = DataGridView1.Rows(i).Cells(11).Value
                    End If
                    If DataGridView1.Rows(i).Cells(12).Value = "0.00" Then
                        jh.gds_neutral = 0
                    Else
                        jh.gds_neutral = DataGridView1.Rows(i).Cells(12).Value
                    End If
                    If DataGridView1.Rows(i).Cells(13).Value = "0.00" Then
                        jh.gds_fijad = 0
                    Else
                        jh.gds_fijad = DataGridView1.Rows(i).Cells(13).Value
                    End If
                    If DataGridView1.Rows(i).Cells(14).Value = "0.00" Then
                        jh.gph_neutral = 0
                    Else
                        jh.gph_neutral = DataGridView1.Rows(i).Cells(14).Value
                    End If
                    If DataGridView1.Rows(i).Cells(15).Value = "0.00" Then
                        jh.gph_fijad = 0
                    Else
                        jh.gph_fijad = DataGridView1.Rows(i).Cells(15).Value
                    End If
                    If DataGridView1.Rows(i).Cells(16).Value = "0.00" Then
                        jh.gph_suaviz = 0
                    Else
                        jh.gph_suaviz = DataGridView1.Rows(i).Cells(16).Value
                    End If
                    If DataGridView1.Rows(i).Cells(17).Value = "0.00" Then
                        jh.gsol = 0
                    Else
                        jh.gsol = DataGridView1.Rows(i).Cells(17).Value
                    End If
                    If DataGridView1.Rows(i).Cells(18).Value = "" Then
                        jh.gobservaciones = ""
                    Else
                        jh.gobservaciones = DataGridView1.Rows(i).Cells(18).Value
                    End If
                    jh.gestado1 = 1
                    bb.insertar_tenido_deta(jh)
                    PictureBox1.Enabled = True
                    PictureBox2.Enabled = True
                    PictureBox3.Enabled = True
                    PictureBox4.Enabled = True
                Next
            End If
        End If
    End Sub
    Dim dt, dt1 As New DataTable

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim k As Integer
        Dim jh As New vtenido
        Dim jh1, jh2 As New vtenido
        Dim bb As New ftenido
        k = DataGridView1.Rows.Count

        For i = k - 1 To k - 1
            jh.gntenido = Trim(TextBox7.Text)
            jh.ghora = DataGridView1.Rows(i).Cells(0).Value
            If DataGridView1.Rows(i).Cells(1).Value = "0.00" Then
                jh.gph = 0
            Else
                jh.gph = DataGridView1.Rows(i).Cells(1).Value
            End If
            If DataGridView1.Rows(i).Cells(2).Value = "0.00" Then
                jh.gmv = 0
            Else
                jh.gmv = DataGridView1.Rows(i).Cells(2).Value
            End If
            If DataGridView1.Rows(i).Cells(3).Value = "0.00" Then
                jh.gindigo = 0
            Else
                jh.gindigo = DataGridView1.Rows(i).Cells(3).Value
            End If
            If DataGridView1.Rows(i).Cells(4).Value = "0.00" Then
                jh.ghidro = 0
            Else
                jh.ghidro = DataGridView1.Rows(i).Cells(4).Value
            End If
            If DataGridView1.Rows(i).Cells(5).Value = "0.00" Then
                jh.gbe = 0
            Else
                jh.gbe = DataGridView1.Rows(i).Cells(5).Value
            End If
            If DataGridView1.Rows(i).Cells(6).Value = "0.00" Then
                jh.gveloc = 0
            Else
                jh.gveloc = DataGridView1.Rows(i).Cells(6).Value
            End If
            If DataGridView1.Rows(i).Cells(7).Value = "0.00" Then
                jh.gtinas = 0
            Else
                jh.gtinas = DataGridView1.Rows(i).Cells(7).Value
            End If
            If DataGridView1.Rows(i).Cells(8).Value = "0.00" Then
                jh.gds_colorant = 0
            Else
                jh.gds_colorant = DataGridView1.Rows(i).Cells(8).Value
            End If
            If DataGridView1.Rows(i).Cells(9).Value = "0.00" Then
                jh.gds_hidrosuf = 0
            Else
                jh.gds_hidrosuf = DataGridView1.Rows(i).Cells(9).Value
            End If
            If DataGridView1.Rows(i).Cells(10).Value = "0.00" Then
                jh.gds_soda = 0
            Else
                jh.gds_soda = DataGridView1.Rows(i).Cells(10).Value
            End If
            If DataGridView1.Rows(i).Cells(11).Value = "0.00" Then
                jh.gds_pquimic = 0
            Else
                jh.gds_pquimic = DataGridView1.Rows(i).Cells(11).Value
            End If
            If DataGridView1.Rows(i).Cells(12).Value = "0.00" Then
                jh.gds_neutral = 0
            Else
                jh.gds_neutral = DataGridView1.Rows(i).Cells(12).Value
            End If
            If DataGridView1.Rows(i).Cells(13).Value = "0.00" Then
                jh.gds_fijad = 0
            Else
                jh.gds_fijad = DataGridView1.Rows(i).Cells(13).Value
            End If
            If DataGridView1.Rows(i).Cells(14).Value = "0.00" Then
                jh.gph_neutral = 0
            Else
                jh.gph_neutral = DataGridView1.Rows(i).Cells(14).Value
            End If
            If DataGridView1.Rows(i).Cells(15).Value = "0.00" Then
                jh.gph_fijad = 0
            Else
                jh.gph_fijad = DataGridView1.Rows(i).Cells(15).Value
            End If
            If DataGridView1.Rows(i).Cells(16).Value = "0.00" Then
                jh.gph_suaviz = 0
            Else
                jh.gph_suaviz = DataGridView1.Rows(i).Cells(16).Value
            End If
            If DataGridView1.Rows(i).Cells(17).Value = "0.00" Then
                jh.gsol = 0
            Else
                jh.gsol = DataGridView1.Rows(i).Cells(17).Value
            End If
            If DataGridView1.Rows(i).Cells(18).Value = "" Then
                jh.gobservaciones = ""
            Else
                jh.gobservaciones = DataGridView1.Rows(i).Cells(18).Value
            End If
            jh.gestado1 = 1
            bb.insertar_tenido_deta(jh)


        Next
        jh2.gntenido = Trim(TextBox7.Text)
        bb.actualizar_estado(jh2)
        Dim respuesta As DialogResult

        respuesta = MessageBox.Show("DESEA IMPRIMIR EL REPORTE DE TEÑIDO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then

            Reporte_teñid.TextBox1.Text = Trim(TextBox7.Text)
            Reporte_teñid.Show()
        End If
        Label8.Visible = True
        DataGridView1.Enabled = False
        Button1.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox7.Enabled = False
        DateTimePicker1.Enabled = False
        PictureBox1.Enabled = False
    End Sub

    Private Sub Teñido_Load(sender As Object, e As EventArgs) Handles MyBase.Load




        DataGridView1.Enabled = True
        Button1.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = True
        TextBox7.Enabled = True
        DateTimePicker1.Enabled = True
        PictureBox1.Enabled = True
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Dim respuesta As DialogResult

        respuesta = MessageBox.Show("DESEA IMPRIMIR EL REPORTE DE TEÑIDO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then

            Reporte_teñid.TextBox1.Text = Trim(TextBox7.Text)
            Reporte_teñid.Show()

        End If
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Dim respuesta As DialogResult

        respuesta = MessageBox.Show("DESEA IMPRIMIR LA GRAFICA DE TEÑIDO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then

            Report_Tenido.TextBox1.Text = Trim(TextBox7.Text)
            Report_Tenido.Show()

        End If
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        Dim respuesta As DialogResult

        respuesta = MessageBox.Show("DESEA IMPRIMIR LA GRAFICA DE TEÑIDO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then

            Rpt_teñido3.TextBox1.Text = Trim(TextBox7.Text)
            Rpt_teñido3.Show()

        End If
    End Sub

    Private Sub PictureBox4_MouseHover(sender As Object, e As EventArgs) Handles PictureBox4.MouseHover
        ToolTip1.SetToolTip(PictureBox4, "GRAFICA TEÑIDO DETALLADA")
        ToolTip1.ToolTipTitle = "GRAFICA TEÑIDO"
        ToolTip1.ToolTipIcon = ToolTipIcon.Info

        ToolTip2.SetToolTip(PictureBox3, "GRAFICA TEÑIDO GENERAL")
        ToolTip2.ToolTipTitle = "GRAFICA TEÑIDO"
        ToolTip2.ToolTipIcon = ToolTipIcon.Info

        ToolTip3.SetToolTip(PictureBox2, "REPORTE TEÑIDO")
        ToolTip3.ToolTipTitle = "HOJA DE TEÑIDO"
        ToolTip3.ToolTipIcon = ToolTipIcon.Info

        ToolTip4.SetToolTip(PictureBox1, "PROCESO DE TEÑIDO")
        ToolTip4.ToolTipTitle = "LACRAR HOJA DE TEÑIDO"
        ToolTip4.ToolTipIcon = ToolTipIcon.Info

        ToolTip5.SetToolTip(Button1, "INGRESAR")
        ToolTip5.ToolTipTitle = "AGREGAR INFORMACION"
        ToolTip5.ToolTipIcon = ToolTipIcon.Info
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim respuesta As DialogResult
        Dim TAB As Integer
        TAB = DataGridView1.Rows.Count
        If TAB <> 0 Then
            respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then

                DataGridView1.Rows.Remove(DataGridView1.CurrentRow)





            End If
        Else
            MsgBox("NO HAY PRODUCTOS A ELIMINAR")
        End If


    End Sub

    Private Sub TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim jk As New ftenido
            Dim hk As New vtenido
            Dim l As Integer
            hk.gntenido = TextBox7.Text
            dt = jk.buscar_cabecera(hk)
            If dt.Rows.Count <> 0 Then
                DataGridView2.DataSource = dt
                PictureBox2.Enabled = True
                PictureBox3.Enabled = True
                PictureBox4.Enabled = True
                TextBox3.Text = DataGridView2.Rows(0).Cells(3).Value
                TextBox4.Text = DataGridView2.Rows(0).Cells(4).Value
                TextBox5.Text = DataGridView2.Rows(0).Cells(5).Value
                TextBox6.Text = DataGridView2.Rows(0).Cells(2).Value
                TextBox7.Text = DataGridView2.Rows(0).Cells(0).Value
                DateTimePicker1.Value = DataGridView2.Rows(0).Cells(1).Value
                dt1 = jk.buscar_detalle(hk)
                DataGridView3.DataSource = dt1
                l = DataGridView3.Rows.Count
                DataGridView1.Rows.Add(l)
                For i = 0 To l - 1
                    DataGridView1.Rows(i).Cells(0).Value = DataGridView3.Rows(i).Cells(2).Value
                    DataGridView1.Rows(i).Cells(1).Value = DataGridView3.Rows(i).Cells(3).Value
                    DataGridView1.Rows(i).Cells(2).Value = DataGridView3.Rows(i).Cells(4).Value
                    DataGridView1.Rows(i).Cells(3).Value = DataGridView3.Rows(i).Cells(5).Value
                    DataGridView1.Rows(i).Cells(4).Value = DataGridView3.Rows(i).Cells(6).Value
                    DataGridView1.Rows(i).Cells(5).Value = DataGridView3.Rows(i).Cells(7).Value
                    DataGridView1.Rows(i).Cells(6).Value = DataGridView3.Rows(i).Cells(8).Value
                    DataGridView1.Rows(i).Cells(7).Value = DataGridView3.Rows(i).Cells(9).Value
                    DataGridView1.Rows(i).Cells(8).Value = DataGridView3.Rows(i).Cells(10).Value
                    DataGridView1.Rows(i).Cells(9).Value = DataGridView3.Rows(i).Cells(11).Value
                    DataGridView1.Rows(i).Cells(10).Value = DataGridView3.Rows(i).Cells(12).Value
                    DataGridView1.Rows(i).Cells(11).Value = DataGridView3.Rows(i).Cells(13).Value
                    DataGridView1.Rows(i).Cells(12).Value = DataGridView3.Rows(i).Cells(14).Value
                    DataGridView1.Rows(i).Cells(13).Value = DataGridView3.Rows(i).Cells(15).Value
                    DataGridView1.Rows(i).Cells(14).Value = DataGridView3.Rows(i).Cells(16).Value
                    DataGridView1.Rows(i).Cells(15).Value = DataGridView3.Rows(i).Cells(17).Value
                    DataGridView1.Rows(i).Cells(16).Value = DataGridView3.Rows(i).Cells(18).Value
                    DataGridView1.Rows(i).Cells(17).Value = DataGridView3.Rows(i).Cells(19).Value
                    DataGridView1.Rows(i).Cells(18).Value = DataGridView3.Rows(i).Cells(20).Value
                Next
            End If
        End If
    End Sub
End Class