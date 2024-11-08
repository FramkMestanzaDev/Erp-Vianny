Public Class Copiar_Desde
    Dim DT1, DT2, DT3, DT4, DT5, DT6, DT7 As New DataTable

    Private Sub Copiar_Desde_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim I, i10, i11 As Integer
        Dim mr As New FCOTIZACION
        Dim mr1 As New VCOTIZACION
        Dim FG As New HojaCotizacion

        'TextBox11.Select()
        I = Len(TextBox1.Text)
        Select Case I
            Case "1" : TextBox1.Text = "0000000" & "" & TextBox1.Text
            Case "2" : TextBox1.Text = "000000" & "" & TextBox1.Text
            Case "3" : TextBox1.Text = "00000" & "" & TextBox1.Text
            Case "4" : TextBox1.Text = "0000" & "" & TextBox1.Text
            Case "5" : TextBox1.Text = "000" & "" & TextBox1.Text
            Case "6" : TextBox1.Text = "00" & "" & TextBox1.Text
            Case "7" : TextBox1.Text = "0" & "" & TextBox1.Text
            Case "8" : TextBox1.Text = TextBox1.Text
        End Select
        mr1.gid_cotizacion = TextBox1.Text
        If mr.mostrar_COTIZACION(mr1) = True Then

            Dim mv As New ftela
            Dim mv1 As New vTela

            mv1.gid_cotizacion = TextBox1.Text
            DT1 = mv.mostrar_tela(mv1)
            If DT1.Rows.Count <> 0 Then

                DataGridView1.DataSource = DT1
                i10 = DataGridView1.Rows.Count
                HojaCotizacion.DataGridView1.Rows.Add(i10)
                For i2 = 0 To i10 - 1
                    HojaCotizacion.DataGridView1.Rows(i2).Cells(0).Value = DataGridView1.Rows(i2).Cells(1).Value
                    HojaCotizacion.DataGridView1.Rows(i2).Cells(1).Value = DataGridView1.Rows(i2).Cells(2).Value
                    HojaCotizacion.DataGridView1.Rows(i2).Cells(2).Value = DataGridView1.Rows(i2).Cells(3).Value
                    HojaCotizacion.DataGridView1.Rows(i2).Cells(3).Value = DataGridView1.Rows(i2).Cells(4).Value
                    HojaCotizacion.DataGridView1.Rows(i2).Cells(4).Value = DataGridView1.Rows(i2).Cells(5).Value
                    HojaCotizacion.DataGridView1.Rows(i2).Cells(5).Value = DataGridView1.Rows(i2).Cells(6).Value
                    HojaCotizacion.DataGridView1.Rows(i2).Cells(6).Value = DataGridView1.Rows(i2).Cells(7).Value
                    HojaCotizacion.DataGridView1.Rows(i2).Cells(7).Value = DataGridView1.Rows(i2).Cells(8).Value
                    HojaCotizacion.DataGridView1.Rows(i2).Cells(8).Value = DataGridView1.Rows(i2).Cells(9).Value

                Next
            End If

            Dim mc As New favios_costura
            Dim mc1 As New vAvios_Costura

            mc1.gid_cotizacion = TextBox1.Text
            DT2 = mc.mostrar_avios_costura(mc1)
            If DT2.Rows.Count <> 0 Then

                DataGridView2.DataSource = DT2
                i11 = DataGridView2.Rows.Count
                HojaCotizacion.DataGridView2.Rows.Add(i11)
                For i3 = 0 To i11 - 1
                    HojaCotizacion.DataGridView2.Rows(i3).Cells(0).Value = DataGridView2.Rows(i3).Cells(1).Value
                    HojaCotizacion.DataGridView2.Rows(i3).Cells(1).Value = DataGridView2.Rows(i3).Cells(2).Value
                    HojaCotizacion.DataGridView2.Rows(i3).Cells(2).Value = DataGridView2.Rows(i3).Cells(3).Value
                    HojaCotizacion.DataGridView2.Rows(i3).Cells(3).Value = DataGridView2.Rows(i3).Cells(4).Value
                    HojaCotizacion.DataGridView2.Rows(i3).Cells(4).Value = DataGridView2.Rows(i3).Cells(5).Value
                    HojaCotizacion.DataGridView2.Rows(i3).Cells(5).Value = DataGridView2.Rows(i3).Cells(6).Value
                    HojaCotizacion.DataGridView2.Rows(i3).Cells(6).Value = DataGridView2.Rows(i3).Cells(7).Value
                    HojaCotizacion.DataGridView2.Rows(i3).Cells(7).Value = DataGridView2.Rows(i3).Cells(8).Value
                    HojaCotizacion.DataGridView2.Rows(i3).Cells(8).Value = DataGridView2.Rows(i3).Cells(9).Value
                Next
            End If

            Dim mk As New favios_acabados
            Dim mk1 As New vAvios_Acabados

            mk1.gid_cotizacion = TextBox1.Text
            DT3 = mk.mostrar_avios_acabados(mk1)
            If DT3.Rows.Count <> 0 Then

                DataGridView3.DataSource = DT3
                Dim i12 As Integer
                i12 = DataGridView3.Rows.Count
                HojaCotizacion.DataGridView3.Rows.Add(i12)
                For i4 = 0 To i12 - 1
                    HojaCotizacion.DataGridView3.Rows(i4).Cells(0).Value = DataGridView3.Rows(i4).Cells(1).Value
                    HojaCotizacion.DataGridView3.Rows(i4).Cells(1).Value = DataGridView3.Rows(i4).Cells(2).Value
                    HojaCotizacion.DataGridView3.Rows(i4).Cells(2).Value = DataGridView3.Rows(i4).Cells(3).Value
                    HojaCotizacion.DataGridView3.Rows(i4).Cells(3).Value = DataGridView3.Rows(i4).Cells(4).Value
                    HojaCotizacion.DataGridView3.Rows(i4).Cells(4).Value = DataGridView3.Rows(i4).Cells(5).Value
                    HojaCotizacion.DataGridView3.Rows(i4).Cells(5).Value = DataGridView3.Rows(i4).Cells(6).Value
                    HojaCotizacion.DataGridView3.Rows(i4).Cells(6).Value = DataGridView3.Rows(i4).Cells(7).Value
                    HojaCotizacion.DataGridView3.Rows(i4).Cells(7).Value = DataGridView3.Rows(i4).Cells(8).Value
                    HojaCotizacion.DataGridView3.Rows(i4).Cells(8).Value = DataGridView3.Rows(i4).Cells(9).Value
                Next
            End If


            Dim mh As New flavados
            Dim mh1 As New vLavados

            mh1.gid_cotizacion = TextBox1.Text
            DT4 = mh.mostrar_lavados(mh1)
            If DT4.Rows.Count <> 0 Then

                DataGridView4.DataSource = DT4
                Dim i22 As Integer
                i22 = DataGridView4.Rows.Count
                HojaCotizacion.DataGridView4.Rows.Add(i22)
                For i44 = 0 To i22 - 1
                    HojaCotizacion.DataGridView4.Rows(i44).Cells(0).Value = DataGridView4.Rows(i44).Cells(1).Value
                    HojaCotizacion.DataGridView4.Rows(i44).Cells(1).Value = DataGridView4.Rows(i44).Cells(2).Value
                    HojaCotizacion.DataGridView4.Rows(i44).Cells(2).Value = DataGridView4.Rows(i44).Cells(3).Value
                    HojaCotizacion.DataGridView4.Rows(i44).Cells(3).Value = DataGridView4.Rows(i44).Cells(4).Value
                    HojaCotizacion.DataGridView4.Rows(i44).Cells(4).Value = DataGridView4.Rows(i44).Cells(5).Value
                    HojaCotizacion.DataGridView4.Rows(i44).Cells(5).Value = DataGridView4.Rows(i44).Cells(6).Value
                    HojaCotizacion.DataGridView4.Rows(i44).Cells(6).Value = DataGridView4.Rows(i44).Cells(7).Value
                    HojaCotizacion.DataGridView4.Rows(i44).Cells(7).Value = DataGridView4.Rows(i44).Cells(8).Value
                Next
            End If

            Dim mf As New faplicaciones
            Dim mf1 As New Aplicaciones

            mf1.gid_cotizacion = TextBox1.Text
            DT5 = mf.mostrar_aplicaciones(mf1)
            If DT5.Rows.Count <> 0 Then

                DataGridView5.DataSource = DT5
                Dim i13 As Integer
                i13 = DataGridView5.Rows.Count
                HojaCotizacion.DataGridView5.Rows.Add(i13)
                For i5 = 0 To i13 - 1
                    HojaCotizacion.DataGridView5.Rows(i5).Cells(0).Value = DataGridView5.Rows(i5).Cells(1).Value
                    HojaCotizacion.DataGridView5.Rows(i5).Cells(1).Value = DataGridView5.Rows(i5).Cells(2).Value
                    HojaCotizacion.DataGridView5.Rows(i5).Cells(2).Value = DataGridView5.Rows(i5).Cells(3).Value
                    HojaCotizacion.DataGridView5.Rows(i5).Cells(3).Value = DataGridView5.Rows(i5).Cells(4).Value
                    HojaCotizacion.DataGridView5.Rows(i5).Cells(4).Value = DataGridView5.Rows(i5).Cells(5).Value
                    HojaCotizacion.DataGridView5.Rows(i5).Cells(5).Value = DataGridView5.Rows(i5).Cells(6).Value
                    HojaCotizacion.DataGridView5.Rows(i5).Cells(6).Value = DataGridView5.Rows(i5).Cells(7).Value
                    HojaCotizacion.DataGridView5.Rows(i5).Cells(7).Value = DataGridView5.Rows(i5).Cells(8).Value
                Next
            End If

            Dim m45 As New fmano_obra
            Dim my1 As New vManoObra

            my1.gid_cotizacion = TextBox1.Text
            DT6 = m45.mostrar_manoobra(my1)
            If DT6.Rows.Count <> 0 Then

                DataGridView6.DataSource = DT6
                Dim i14 As Integer
                i14 = DataGridView6.Rows.Count
                HojaCotizacion.DataGridView6.Rows.Add(i14)
                For i6 = 0 To i14 - 1

                    HojaCotizacion.DataGridView6.Rows(i6).Cells(0).Value = DataGridView6.Rows(i6).Cells(1).Value
                    HojaCotizacion.DataGridView6.Rows(i6).Cells(1).Value = DataGridView6.Rows(i6).Cells(2).Value
                    HojaCotizacion.DataGridView6.Rows(i6).Cells(2).Value = DataGridView6.Rows(i6).Cells(3).Value
                    HojaCotizacion.DataGridView6.Rows(i6).Cells(3).Value = DataGridView6.Rows(i6).Cells(4).Value
                    HojaCotizacion.DataGridView6.Rows(i6).Cells(4).Value = DataGridView6.Rows(i6).Cells(5).Value
                    HojaCotizacion.DataGridView6.Rows(i6).Cells(5).Value = DataGridView6.Rows(i6).Cells(6).Value
                    HojaCotizacion.DataGridView6.Rows(i6).Cells(6).Value = DataGridView6.Rows(i6).Cells(7).Value
                Next
            End If



            mr1.gid_cotizacion = TextBox1.Text
            DT7 = mr.mostrar_COTIZACION2(mr1)
            If DT7.Rows.Count <> 0 Then

                DataGridView7.DataSource = DT7
                If DataGridView7.Rows.Count <> 0 Then
                    MsgBox("SE COPIO CON EXITOS")
                    HojaCotizacion.TextBox2.Text = DataGridView7.Rows(0).Cells(1).Value
                    HojaCotizacion.TextBox3.Text = DataGridView7.Rows(0).Cells(2).Value
                    HojaCotizacion.TextBox4.Text = DataGridView7.Rows(0).Cells(3).Value
                    HojaCotizacion.TextBox5.Text = DataGridView7.Rows(0).Cells(4).Value
                    HojaCotizacion.TextBox8.Text = DataGridView7.Rows(0).Cells(5).Value
                    If DataGridView7.Rows(0).Cells(6).Value = 1 Then
                        HojaCotizacion.ComboBox1.Text = "SOLES"
                    Else
                        HojaCotizacion.ComboBox1.Text = "DOLARES"
                    End If
                    HojaCotizacion.TextBox9.Text = DataGridView7.Rows(0).Cells(8).Value
                    'HojaCotizacion.TextBox10.Text = DataGridView7.Rows(0).Cells(9).Value
                    HojaCotizacion.TextBox11.Text = DataGridView7.Rows(0).Cells(10).Value
                    HojaCotizacion.TextBox6.Text = DataGridView7.Rows(0).Cells(11).Value
                    HojaCotizacion.TextBox12.Text = DataGridView7.Rows(0).Cells(12).Value
                    HojaCotizacion.TextBox13.Text = DataGridView7.Rows(0).Cells(14).Value
                    HojaCotizacion.TextBox14.Text = DataGridView7.Rows(0).Cells(15).Value
                    HojaCotizacion.TextBox15.Text = DataGridView7.Rows(0).Cells(16).Value
                    HojaCotizacion.TextBox16.Text = DataGridView7.Rows(0).Cells(17).Value
                    HojaCotizacion.TextBox7.Text = DataGridView7.Rows(0).Cells(20).Value
                    HojaCotizacion.TextBox17.Text = DataGridView7.Rows(0).Cells(22).Value
                    HojaCotizacion.TextBox18.Text = DataGridView7.Rows(0).Cells(24).Value
                    HojaCotizacion.TextBox19.Text = DataGridView7.Rows(0).Cells(21).Value
                    HojaCotizacion.TextBox20.Text = DataGridView7.Rows(0).Cells(23).Value
                    HojaCotizacion.TextBox21.Text = DataGridView7.Rows(0).Cells(25).Value
                    HojaCotizacion.TextBox23.Text = DataGridView7.Rows(0).Cells(13).Value

                End If

                HojaCotizacion.Button13.Enabled = True

            End If
            Close()
        Else
            MsgBox("LA COTIZACION INGRESADA NO EXISTE")
        End If
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case TextBox1.Text.Length

                Case "1" : TextBox1.Text = "0000000" & "" & TextBox1.Text
                Case "2" : TextBox1.Text = "000000" & "" & TextBox1.Text
                Case "3" : TextBox1.Text = "00000" & "" & TextBox1.Text
                Case "4" : TextBox1.Text = "0000" & "" & TextBox1.Text
                Case "5" : TextBox1.Text = "000" & "" & TextBox1.Text
                Case "6" : TextBox1.Text = "00" & "" & TextBox1.Text
                Case "7" : TextBox1.Text = "0" & "" & TextBox1.Text
                Case "8" : TextBox1.Text = TextBox1.Text

            End Select
            Button1.Enabled = True
        End If
    End Sub
End Class
