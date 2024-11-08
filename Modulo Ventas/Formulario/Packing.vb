Public Class Packing
    Dim DT As New DataTable


    Private Sub Packing_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim func As New fpack_tela
        Me.TextBox1.Text = func.correlativo_pack_tela + 1
        Select Case TextBox1.Text.Length
            Case "1" : TextBox1.Text = "000000000" & "" & TextBox1.Text
            Case "2" : TextBox1.Text = "00000000" & "" & TextBox1.Text
            Case "3" : TextBox1.Text = "0000000" & "" & TextBox1.Text
            Case "4" : TextBox1.Text = "000000" & "" & TextBox1.Text
            Case "5" : TextBox1.Text = "00000" & "" & TextBox1.Text
            Case "6" : TextBox1.Text = "0000" & "" & TextBox1.Text
            Case "7" : TextBox1.Text = "000" & "" & TextBox1.Text
            Case "8" : TextBox1.Text = "00" & "" & TextBox1.Text
            Case "9" : TextBox1.Text = "0" & "" & TextBox1.Text
            Case "10" : TextBox1.Text = TextBox1.Text
        End Select
        TextBox13.Enabled = False
        Button1.Enabled = False

    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click

    End Sub
    Dim dt4 As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Dim u As Integer

        u = DataGridView1.Rows.Count


        'If u = 0 Then
        '    
        'Else
        '    Dim oj, kj As Integer
        '    oj = DataGridView1.Rows.Count

        '    For i = 0 To oj - 1
        '        If ComboBox1.Text = DataGridView1.Rows(i).Cells(5).Value Then
        '            kj = 1
        '        End If
        '    Next
        '    MsgBox(kj)
        '    If kj = 1 Then
        '        MsgBox("LA PARTIDA YA ESTA INGRESADA")
        '    Else
        '        Dim jh As New vpartida
        '        Dim jh1 As New fpartida
        '        Dim i As Integer
        '        jh.gpartida = ComboBox1.Text
        '        jh.galmacen = TextBox12.Text
        '        dt4 = jh1.buscar_partida2(jh)

        '        If dt4.Rows.Count <> 0 Then
        '            DataGridView2.DataSource = dt4
        '            i = DataGridView2.Rows.Count
        '            DataGridView1.Rows.Add(i)
        '            For a = 0 To i - 1
        '                DataGridView1.Rows(a).Cells(1).Value = DataGridView2.Rows(a).Cells(0).Value
        '                DataGridView1.Rows(a).Cells(2).Value = DataGridView2.Rows(a).Cells(1).Value
        '                DataGridView1.Rows(a).Cells(3).Value = DataGridView2.Rows(a).Cells(2).Value
        '                DataGridView1.Rows(a).Cells(4).Value = DataGridView2.Rows(a).Cells(3).Value
        '                DataGridView1.Rows(a).Cells(5).Value = ComboBox1.Text
        '                DataGridView1.Rows(a).Cells(0).Value = True
        '                DataGridView1.Rows(a).Cells(6).Value = DataGridView2.Rows(a).Cells(4).Value
        '            Next
        '        End If
        '    End If
        'End If
    End Sub




    Private Sub TextBox9_Click(sender As Object, e As EventArgs) Handles TextBox9.Click
        Dim j As Integer
        Dim cant9, sum9 As Double
        j = DataGridView1.Rows.Count


        For i = 0 To j - 1
            If DataGridView1.Rows(i).Cells(0).Value = True Then
                cant9 = Val(DataGridView1.Rows(i).Cells(2).Value)

                sum9 = cant9 + Val(sum9)
            End If


        Next


        TextBox9.Text = sum9.ToString("N2")

    End Sub
    Dim dt5 As New DataTable
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim k1, k2 As Integer
        k1 = DataGridView1.Rows.Count()

        For i = 0 To k1 - 1
            If Me.DataGridView1.Rows(i).Cells(0).Value = True Then
                k2 = k2 + 1
            End If
        Next
        If TextBox9.Text = "" Then
            MsgBox("EL CAMPO KG SELECCIONADOS ESTA VACIO")
        Else
            If k1 = 0 Or k2 = 0 Then
                MsgBox("NO SE HA SELECCIONADO NINGUN ITEMSS")
            Else
                Dim uj As New fpack_tela
                Dim uj1 As New vpackin_tela_cab
                Dim jh As New vpackin_tela_det
                Dim gh As New fingresopac
                Dim vg5 As New vpackingtela
                Dim jk As New fnopedido
                Dim jk1 As New vnapedido
                Dim k As Integer
                uj1.gid_packing = TextBox1.Text
                uj1.gcod_pedido = TextBox6.Text
                uj1.gnota_pedido = TextBox7.Text
                uj1.gcodigo_pro = TextBox8.Text
                uj1.gdescrip_pro = TextBox2.Text
                uj1.gcantidad = TextBox3.Text
                uj1.gf_despacho = DateTimePicker1.Value
                uj1.gf_actual = DateTimePicker2.Value
                uj1.gcliente = TextBox4.Text
                uj1.galmacen = TextBox12.Text
                Select Case TextBox5.Text
                    Case "GBEDON" : uj1.gVendedor = "0013"
                    Case "VINCIO" : uj1.gVendedor = "0021"
                    Case "DBRAVO" : uj1.gVendedor = "0023"
                    Case "JSALINAS" : uj1.gVendedor = "0026"
                    Case "GCUEVA" : uj1.gVendedor = "0027"
                    Case "AMENDO" : uj1.gVendedor = "0024"
                    Case "VGRAUS" : uj1.gVendedor = "0025"
                    Case "VPIZARRO" : uj1.gVendedor = "0011"
                    Case "JBALVIN" : uj1.gVendedor = "0003"
                    Case "VSILVERIO" : uj1.gVendedor = "0001"
                End Select

                uj1.gestado = 0
                uj1.gkg_seleccionados = TextBox9.Text
                uj.insertar_packin_tela(uj1)

                k = DataGridView1.Rows.Count()

                For i = 0 To k - 1
                    If Me.DataGridView1.Rows(i).Cells(0).Value = True Then

                        jh.grollo = DataGridView1.Rows(i).Cells(1).Value
                        jh.gcantidad = DataGridView1.Rows(i).Cells(2).Value
                        jh.gubicacion = DataGridView1.Rows(i).Cells(3).Value
                        jh.galmacen = DataGridView1.Rows(i).Cells(4).Value
                        jh.gid_packing = TextBox1.Text
                        jh.gcodigo_pro = TextBox8.Text
                        jh.gpartida = DataGridView1.Rows(i).Cells(5).Value
                        jh.gestado = 0
                        uj.insertar_packin_tela_DETALLE(jh)
                        Dim ch As New vpackingtela
                        Dim bh1 As New fingresopac
                        ch.gcodigo_det = DataGridView1.Rows(i).Cells(6).Value
                        ch.gcodigo2 = DataGridView1.Rows(i).Cells(7).Value
                        bh1.actualizar_packing(ch)
                    End If
                Next

                jk1.gnumero_pedido = TextBox7.Text

                jk.actualizar_estado_almacen(jk1)
                'Dim jp, jp1, jp2 As Integer
                'vg5.gpartida = ComboBox1.Text
                'vg5.galmacen = TextBox12.Text
                'jp = gh.buscar_cant_packin(vg5)
                'jp1 = gh.buscar_cant_packin_almacen(vg5)

                'jp2 = jp1 + k
                'If jp2 = jp Then
                '    MsgBox("estan registrado todos")
                'End If

                Dim a As Integer
                a = DataGridView1.Rows.Count

                MsgBox("LA INFORMACION SE REGISTRO CON EXITO")
                DataGridView1.DataSource = ""
                DataGridView2.DataSource = ""
                DataGridView3.DataSource = ""
                DataGridView4.DataSource = ""
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox5.Text = ""
                TextBox6.Text = ""
                TextBox7.Text = ""
                TextBox8.Text = ""
                TextBox9.Text = ""
                TextBox10.Text = ""
                TextBox11.Text = ""
                TextBox12.Text = ""
                TextBox13.Text = ""
                Me.Close()
            End If
        End If
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim gh As New fproductos
        Dim gh1 As New vproducot

        gh1.gcodigo = TextBox8.Text

        DT = gh.buscar_PARTIDA(gh1)
        Partidas.DataGridView1.DataSource = DT

        Partidas.Show()
        DataGridView2.DataSource = ""
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox13.Enabled = True
            Button1.Enabled = True
        End If
    End Sub
    Dim dt8 As New DataTable
    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Dim gy, ki As New vpackingtela
        Dim gy1 As New fingresopac
        Dim gh As New fproductos
        Dim gh1 As New vproducot
        Dim h, kj As Integer
        Dim jh As String
        jh = ""
        h = DataGridView1.Rows.Count
        For i = 0 To h - 1

            If DataGridView1.Rows(i).Cells(5).Value.ToString.Trim = TextBox13.Text Then
                ki.gcodigo_det = DataGridView1.Rows(i).Cells(6).Value.ToString.Trim
                ki.ges = 0
                jh = DataGridView1.Rows(i).Cells(7).Value
              
                gy1.actualizar_seleccionado_detalle(ki)
            End If
        Next
        gy.gcodigo_det = jh
        gy.ges = 0

        gy1.actualizar_seleccionado(gy)


        gh1.gcodigo = TextBox8.Text

        dt8 = gh.buscar_PARTIDA_estado1(gh1)
        DataGridView1.Rows.Clear()
        DataGridView4.DataSource = dt8

        kj = DataGridView4.Rows.Count
        DataGridView1.Rows.Add(kj)
        For t = 0 To kj - 1
            DataGridView1.Rows(t).Cells(1).Value = DataGridView4.Rows(t).Cells(0).Value
            DataGridView1.Rows(t).Cells(2).Value = DataGridView4.Rows(t).Cells(1).Value
            DataGridView1.Rows(t).Cells(3).Value = DataGridView4.Rows(t).Cells(2).Value
            DataGridView1.Rows(t).Cells(4).Value = DataGridView4.Rows(t).Cells(3).Value
            DataGridView1.Rows(t).Cells(5).Value = DataGridView4.Rows(t).Cells(4).Value
            DataGridView1.Rows(t).Cells(6).Value = DataGridView4.Rows(t).Cells(5).Value
            DataGridView1.Rows(t).Cells(7).Value = DataGridView4.Rows(t).Cells(6).Value
        Next
        TextBox13.Text = ""
        TextBox13.Enabled = False
        Button1.Enabled = False

    End Sub

    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

        If e.ColumnIndex = 0 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            DataGridView1.Rows(num1).Cells(8).Value = False


        Else
            If e.ColumnIndex = 8 Then

                ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
                Dim num1, num2 As Integer
                num1 = e.RowIndex.ToString
                num2 = e.ColumnIndex.ToString
                DataGridView1.Rows(num1).Cells(0).Value = False
                DataGridView1.Rows(num1).Cells(9).ReadOnly = False
            Else

            End If
        End If

    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim fila As Integer

        fila = e.RowIndex
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "CANT." Then
            If DataGridView1.Rows(fila).Cells(9).Value > DataGridView1.Rows(fila).Cells(2).Value Then
                MsgBox("LA CANTIDAD ES MAYOR AL STOCK")
                DataGridView1.Rows(fila).Cells(9).Value = 0.00
            End If

        End If
    End Sub

    Private Sub Packing_Closed(sender As Object, e As EventArgs) Handles Me.Closed

        Dim j As Integer
        Dim gy, gy2 As New vpackingtela
        Dim gy1 As New fingresopac
        j = DataGridView1.Rows.Count

        For i = 0 To j - 1
            gy.gcodigo_det = DataGridView1.Rows(i).Cells(7).Value
            gy.ges = 0
            gy1.actualizar_seleccionado(gy)
            gy2.gcodigo_det = DataGridView1.Rows(i).Cells(6).Value
            gy2.ges = 0
            gy1.actualizar_seleccionado_detalle(gy2)
        Next

    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs)

    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged

    End Sub
End Class