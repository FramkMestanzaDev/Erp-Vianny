Public Class HojaCotizacion
    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form_Cotizacion.TextBox3.Text = 1
        Dim va As Integer
        va = DataGridView1.Rows.Count
        If va > 0 Then
            Form_Cotizacion.DataGridView1.Rows.Add(va)
            For i = 0 To va - 1
                Form_Cotizacion.DataGridView1.Rows(i).Cells(0).Value = DataGridView1.Rows(i).Cells(0).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(1).Value = DataGridView1.Rows(i).Cells(1).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(2).Value = DataGridView1.Rows(i).Cells(2).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(3).Value = DataGridView1.Rows(i).Cells(3).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(4).Value = DataGridView1.Rows(i).Cells(4).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(5).Value = DataGridView1.Rows(i).Cells(5).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(6).Value = DataGridView1.Rows(i).Cells(6).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(7).Value = DataGridView1.Rows(i).Cells(7).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(8).Value = DataGridView1.Rows(i).Cells(8).Value
            Next
        End If

        Form_Cotizacion.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Form_Cotizacion.TextBox3.Text = 2
        Dim va As Integer
        va = DataGridView2.Rows.Count
        If va > 0 Then
            Form_Cotizacion.DataGridView1.Rows.Add(va)
            For i = 0 To va - 1
                Form_Cotizacion.DataGridView1.Rows(i).Cells(0).Value = DataGridView2.Rows(i).Cells(0).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(1).Value = DataGridView2.Rows(i).Cells(1).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(2).Value = DataGridView2.Rows(i).Cells(2).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(3).Value = DataGridView2.Rows(i).Cells(3).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(4).Value = DataGridView2.Rows(i).Cells(4).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(5).Value = DataGridView2.Rows(i).Cells(5).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(6).Value = DataGridView2.Rows(i).Cells(6).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(7).Value = DataGridView2.Rows(i).Cells(7).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(8).Value = DataGridView2.Rows(i).Cells(8).Value
            Next
        End If
        Form_Cotizacion.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Form_Cotizacion.TextBox3.Text = 3
        Dim va As Integer
        va = DataGridView3.Rows.Count
        If va > 0 Then
            Form_Cotizacion.DataGridView1.Rows.Add(va)
            For i = 0 To va - 1
                Form_Cotizacion.DataGridView1.Rows(i).Cells(0).Value = DataGridView3.Rows(i).Cells(0).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(1).Value = DataGridView3.Rows(i).Cells(1).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(2).Value = DataGridView3.Rows(i).Cells(2).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(3).Value = DataGridView3.Rows(i).Cells(3).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(4).Value = DataGridView3.Rows(i).Cells(4).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(5).Value = DataGridView3.Rows(i).Cells(5).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(6).Value = DataGridView3.Rows(i).Cells(6).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(7).Value = DataGridView3.Rows(i).Cells(7).Value
                Form_Cotizacion.DataGridView1.Rows(i).Cells(8).Value = DataGridView3.Rows(i).Cells(8).Value
            Next
        End If
        Form_Cotizacion.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form_Cotizacion2.TextBox2.Text = 1
        Dim va As Integer
        va = DataGridView4.Rows.Count
        If va > 0 Then
            Form_Cotizacion2.DataGridView1.Rows.Add(va)
            For i = 0 To va - 1
                Form_Cotizacion2.DataGridView1.Rows(i).Cells(0).Value = DataGridView4.Rows(i).Cells(0).Value
                Form_Cotizacion2.DataGridView1.Rows(i).Cells(1).Value = DataGridView4.Rows(i).Cells(1).Value
                Form_Cotizacion2.DataGridView1.Rows(i).Cells(2).Value = DataGridView4.Rows(i).Cells(2).Value
                Form_Cotizacion2.DataGridView1.Rows(i).Cells(3).Value = DataGridView4.Rows(i).Cells(3).Value
                Form_Cotizacion2.DataGridView1.Rows(i).Cells(4).Value = DataGridView4.Rows(i).Cells(4).Value
                Form_Cotizacion2.DataGridView1.Rows(i).Cells(5).Value = DataGridView4.Rows(i).Cells(5).Value
                Form_Cotizacion2.DataGridView1.Rows(i).Cells(6).Value = DataGridView4.Rows(i).Cells(6).Value
                Form_Cotizacion2.DataGridView1.Rows(i).Cells(7).Value = DataGridView4.Rows(i).Cells(7).Value
            Next
        End If
        Form_Cotizacion2.Show()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Form_Cotizacion2.TextBox2.Text = 2
        Dim va As Integer
        va = DataGridView5.Rows.Count
        If va > 0 Then
            Form_Cotizacion2.DataGridView1.Rows.Add(va)
            For i = 0 To va - 1
                Form_Cotizacion2.DataGridView1.Rows(i).Cells(0).Value = DataGridView5.Rows(i).Cells(0).Value
                Form_Cotizacion2.DataGridView1.Rows(i).Cells(1).Value = DataGridView5.Rows(i).Cells(1).Value
                Form_Cotizacion2.DataGridView1.Rows(i).Cells(2).Value = DataGridView5.Rows(i).Cells(2).Value
                Form_Cotizacion2.DataGridView1.Rows(i).Cells(3).Value = DataGridView5.Rows(i).Cells(3).Value
                Form_Cotizacion2.DataGridView1.Rows(i).Cells(4).Value = DataGridView5.Rows(i).Cells(4).Value
                Form_Cotizacion2.DataGridView1.Rows(i).Cells(5).Value = DataGridView5.Rows(i).Cells(5).Value
                Form_Cotizacion2.DataGridView1.Rows(i).Cells(6).Value = DataGridView5.Rows(i).Cells(6).Value
                Form_Cotizacion2.DataGridView1.Rows(i).Cells(7).Value = DataGridView5.Rows(i).Cells(7).Value
            Next
        End If
        Form_Cotizacion2.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim va As Integer
        va = DataGridView6.Rows.Count

        If va > 0 Then
            Form_Cotizacion3.DataGridView1.Rows.Add(va)
            For i = 0 To va - 1
                Form_Cotizacion3.DataGridView1.Rows(i).Cells(0).Value = DataGridView6.Rows(i).Cells(0).Value
                Form_Cotizacion3.DataGridView1.Rows(i).Cells(1).Value = DataGridView6.Rows(i).Cells(1).Value
                Form_Cotizacion3.DataGridView1.Rows(i).Cells(2).Value = DataGridView6.Rows(i).Cells(2).Value
                Form_Cotizacion3.DataGridView1.Rows(i).Cells(3).Value = DataGridView6.Rows(i).Cells(3).Value
                Form_Cotizacion3.DataGridView1.Rows(i).Cells(4).Value = DataGridView6.Rows(i).Cells(4).Value
                Form_Cotizacion3.DataGridView1.Rows(i).Cells(5).Value = DataGridView6.Rows(i).Cells(5).Value
                Form_Cotizacion3.DataGridView1.Rows(i).Cells(6).Value = DataGridView6.Rows(i).Cells(6).Value
            Next
        End If
        'Form_Cotizacion3.TextBox1.Text = DataGridView6.Rows(0).Cells(2).Value
        Form_Cotizacion3.Show()
    End Sub
    Dim dt As DataTable
    Private Sub HojaCotizacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox22.Text = 0


        TextBox8.Text = 0
        ComboBox1.SelectedIndex = 0
        ComboBox3.SelectedIndex = 0

        TextBox7.Text = 0.00
        TextBox17.Text = 0.00
        TextBox18.Text = 0.00
        TextBox19.Text = 0.00
        TextBox20.Text = 0.00
        TextBox21.Text = 0.00
        ComboBox1.SelectedIndex = 0
        Dim func As New FCOTIZACION
        Me.TextBox1.Text = func.CORRELATIVO_COTIZACION() + 1
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
        TextBox1.Select()
        BLOQUEAR_registros()
        Dim func12 As New ftcambio
        Dim dts34 As New vtcambio
        dts34.gfecha = DateTimePicker1.Text
        TextBox8.Text = func12.mostrar_tipo_cambio(dts34)
    End Sub
    Sub BLOQUEAR_registros()

        TextBox2.Enabled = False
        TextBox3.Enabled = False
        ComboBox1.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox8.Enabled = False

        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        Button7.Enabled = False
        Button8.Enabled = False
        PictureBox1.Enabled = False
        Button11.Enabled = False
        Button12.Enabled = False
        Button13.Enabled = False
        Button14.Enabled = False
    End Sub
    Sub habilitar_registros()
        TextBox1.Enabled = False
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        ComboBox1.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox8.Enabled = True
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        Button5.Enabled = True
        Button6.Enabled = True
        Button7.Enabled = True
        Button8.Enabled = True
        PictureBox1.Enabled = False
        'PictureBox2.Enabled = True
        Button12.Enabled = False
        Button13.Enabled = False
        Button14.Enabled = True
    End Sub


    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click

    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged

    End Sub
    Dim dt1, dt2, dt3, dt4, dt5, dt6, dt7 As DataTable
    Sub limiar()
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox7.Text = 0
        TextBox17.Text = 0
        TextBox18.Text = 0
        TextBox19.Text = 0
        TextBox20.Text = 0
        TextBox21.Text = 0
        TextBox1.Select()
        BLOQUEAR_registros()
    End Sub
    Sub HABILITAR()
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox8.Enabled = True

        ComboBox1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        Button5.Enabled = True
        Button6.Enabled = True
        Button7.Enabled = True
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Copiar_Desde.Show()
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Clientes.TextBox3.Text = 18
        Clientes.Show()
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Adicionar_op.Show()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        '****guardar informacion cabecera cotizacion

        If TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox8.Text = 0 Then
            MsgBox("ALGUNOS CAMPOS OBLIGATORIOS ESTAN VACIOS", vbExclamation)
        Else
            Dim ui As New VCOTIZACION
            Dim fz As New FCOTIZACION
            Dim cotiza As String
            Dim accion As String
            If TextBox22.Text = 1 Then
                cotiza = TextBox1.Text
                ui.gid_cotizacion = TextBox1.Text
                fz.eliminar_cotizacion(ui)
                accion = 2
            Else
                ui.gid_cotizacion = TextBox1.Text
                If fz.mostrar_COTIZACION(ui) = True Then
                    cotiza = fz.CORRELATIVO_COTIZACION() + 1
                    Select Case cotiza.Length

                        Case "1" : cotiza = "0000000" & "" & cotiza
                        Case "2" : cotiza = "000000" & "" & cotiza
                        Case "3" : cotiza = "00000" & "" & cotiza
                        Case "4" : cotiza = "0000" & "" & cotiza
                        Case "5" : cotiza = "000" & "" & cotiza
                        Case "6" : cotiza = "00" & "" & cotiza
                        Case "7" : cotiza = "0" & "" & cotiza
                        Case "8" : cotiza = cotiza
                    End Select
                    MsgBox("LA COTIZACION Nº" & "  " & TextBox1.Text.ToString & "  " & "YA ESTA REGISTRADA", vbExclamation)

                    MsgBox("SE ACTUALIZARA A LA COTIZACION Nº" & "  " & cotiza.ToString, vbInformation)
                    accion = 1
                Else
                    cotiza = TextBox1.Text
                    accion = 1
                End If
            End If

            ui.gid_cotizacion = cotiza
            ui.gcliente = TextBox2.Text
            ui.gdescripcion = TextBox3.Text
            ui.gestilo = TextBox4.Text
            ui.grango_tallas = TextBox5.Text
            ui.gti_cambio = TextBox8.Text
            If ComboBox1.Text = "SOLES" Then
                ui.gmoneda = 1
            Else
                ui.gmoneda = 2
            End If
            Select Case ComboBox3.Text

                Case "PRENDAS" : ui.glinea = 1
                Case "TELA NOTEX" : ui.glinea = 2
                Case "INDUMENTARIA MEDICA" : ui.glinea = 3
                Case "TELA PUNTO" : ui.glinea = 4
            End Select
            ui.ggasto_logistico = TextBox9.Text
            ui.ggasto_operativo = TextBox24.Text
            ui.ggasto_administrativo = TextBox11.Text
            ui.ggasto_financiero = TextBox6.Text
            ui.ggasto_venta = TextBox12.Text
            ui.gcosto_producto = TextBox23.Text
            ui.gmargen_utilidad = TextBox13.Text
            ui.gmargen_utilidad_moneda = TextBox10.Text
            ui.gvalor_venta = TextBox14.Text
            ui.gigv = TextBox15.Text
            ui.gprecio_venta = TextBox16.Text
            ui.gestado = 0
            ui.gimagen = "POR DENIFIR"
            ui.gcostot_tela = TextBox7.Text
            ui.gcostot_AviosC = TextBox19.Text
            ui.gcostot_AviosA = TextBox17.Text
            ui.gcostot_Lavado = TextBox20.Text
            ui.gcostot_Aplicaciones = TextBox18.Text
            ui.gcostot_MObra = TextBox21.Text
            ui.gvendedor = ComboBox2.Text
            ui.gfecha = DateTimePicker1.Value

            If fz.insertar_cotizacion(ui) = True Then

                '--- REGISTRO TELA
                Dim UH As New vTela
                Dim FZ1 As New ftela
                Dim J As New Integer
                J = DataGridView1.Rows.Count
                For I = 0 To J - 1
                    UH.gcodigo = DataGridView1.Rows(I).Cells(0).Value
                    UH.gdescripcion = DataGridView1.Rows(I).Cells(1).Value
                    UH.gmerma = DataGridView1.Rows(I).Cells(2).Value
                    UH.gconsumo = DataGridView1.Rows(I).Cells(3).Value
                    UH.gunidad = DataGridView1.Rows(I).Cells(4).Value
                    UH.gconsumo_total = DataGridView1.Rows(I).Cells(5).Value
                    UH.gccosto_unitario = DataGridView1.Rows(I).Cells(6).Value
                    UH.gcosto_total = DataGridView1.Rows(I).Cells(7).Value
                    UH.gitems = DataGridView1.Rows(I).Cells(8).Value
                    UH.gid_cotizacion = cotiza
                    FZ1.insertar_tela(UH)
                Next
                '----- REGISTRO AVIOS COSTURA
                Dim GH As New vAvios_Costura
                Dim FZ2 As New favios_costura
                Dim J1 As New Integer
                J1 = DataGridView2.Rows.Count
                For I1 = 0 To J1 - 1
                    GH.gcodigo = DataGridView2.Rows(I1).Cells(0).Value
                    GH.gdescripcion = DataGridView2.Rows(I1).Cells(1).Value
                    GH.gmerma = DataGridView2.Rows(I1).Cells(2).Value
                    GH.gconsumo = DataGridView2.Rows(I1).Cells(3).Value
                    GH.gunidad = DataGridView2.Rows(I1).Cells(4).Value
                    GH.gconsumo_total = DataGridView2.Rows(I1).Cells(5).Value
                    GH.gccosto_unitario = DataGridView2.Rows(I1).Cells(6).Value
                    GH.gcosto_total = DataGridView2.Rows(I1).Cells(7).Value
                    GH.gitems = DataGridView2.Rows(I1).Cells(8).Value
                    GH.gid_cotizacion = cotiza
                    FZ2.insertar_avios_costura(GH)
                Next


                '---REGISTRO AVIOS ACABADOS
                Dim HJ As New vAvios_Acabados
                Dim FZ3 As New favios_acabados
                Dim J2 As New Integer
                J2 = DataGridView3.Rows.Count
                For I2 = 0 To J2 - 1
                    HJ.gcodigo = DataGridView3.Rows(I2).Cells(0).Value
                    HJ.gdescripcion = DataGridView3.Rows(I2).Cells(1).Value
                    HJ.gmerma = DataGridView3.Rows(I2).Cells(2).Value
                    HJ.gconsumo = DataGridView3.Rows(I2).Cells(3).Value
                    HJ.gunidad = DataGridView3.Rows(I2).Cells(4).Value
                    HJ.gconsumo_total = DataGridView3.Rows(I2).Cells(5).Value
                    HJ.gccosto_unitario = DataGridView3.Rows(I2).Cells(6).Value
                    HJ.gcosto_total = DataGridView3.Rows(I2).Cells(7).Value
                    HJ.gitems = DataGridView3.Rows(I2).Cells(8).Value
                    HJ.gid_cotizacion = cotiza
                    FZ3.insertar_avios_acabados(HJ)
                Next

                '-----REGISTRO LAVADOS
                Dim FT As New vLavados
                Dim FZ4 As New flavados
                Dim J3 As New Integer
                J3 = DataGridView4.Rows.Count
                For I3 = 0 To J3 - 1
                    FT.gdescripcion = DataGridView4.Rows(I3).Cells(0).Value
                    FT.gmerma = DataGridView4.Rows(I3).Cells(1).Value
                    FT.gconsumo = DataGridView4.Rows(I3).Cells(2).Value
                    FT.gunidad = DataGridView4.Rows(I3).Cells(3).Value
                    FT.gconsumo_tot = DataGridView4.Rows(I3).Cells(4).Value
                    FT.gcosto_unitario = DataGridView4.Rows(I3).Cells(5).Value
                    FT.gcosto_total = DataGridView4.Rows(I3).Cells(6).Value
                    FT.gitems = DataGridView4.Rows(I3).Cells(7).Value
                    FT.gid_cotizacion = cotiza

                    FZ4.insertar_lavados(FT)
                Next

                '-----REGISTRO APLICACIONES
                Dim FT1 As New Aplicaciones
                Dim FZ5 As New faplicaciones
                Dim J4 As New Integer
                J4 = DataGridView5.Rows.Count
                For I4 = 0 To J4 - 1

                    FT1.gdescripcion = DataGridView5.Rows(I4).Cells(0).Value
                    FT1.gmerma = DataGridView5.Rows(I4).Cells(1).Value
                    FT1.gconsumo = DataGridView5.Rows(I4).Cells(2).Value
                    FT1.gunidad = DataGridView5.Rows(I4).Cells(3).Value
                    FT1.gconsumo_tot = DataGridView5.Rows(I4).Cells(4).Value
                    FT1.gcosto_unitario = DataGridView5.Rows(I4).Cells(5).Value
                    FT1.gcosto_total = DataGridView5.Rows(I4).Cells(6).Value
                    FT1.gitems = DataGridView5.Rows(I4).Cells(7).Value
                    FT1.gid_cotizacion = cotiza

                    FZ5.insertar_aplicaciones(FT1)
                Next

                '---- INSERTAR MANO OBRA
                Dim FT2 As New vManoObra
                Dim FZ6 As New fmano_obra
                Dim J5 As New Integer
                J5 = DataGridView6.Rows.Count
                For I5 = 0 To J5 - 1

                    FT2.gdescripcion = DataGridView6.Rows(I5).Cells(0).Value
                    FT2.gmerma = DataGridView6.Rows(I5).Cells(1).Value
                    FT2.gconsumo = DataGridView6.Rows(I5).Cells(2).Value
                    FT2.gunidad = DataGridView6.Rows(I5).Cells(3).Value
                    FT2.gconsumo_total = DataGridView6.Rows(I5).Cells(4).Value
                    FT2.gcosto_unitario = DataGridView6.Rows(I5).Cells(5).Value
                    FT2.gcosto_total = DataGridView6.Rows(I5).Cells(6).Value
                    FT2.gid_cotizacion = cotiza


                    FZ6.insertar_mano_obra(FT2)
                Next
                Dim resultado As String
                resultado = MsgBox("¿Quieres imprimir la Cotizacion?", vbYesNo, "Colorear A1")

                If resultado = vbYes Then
                    MsgBox("SE REGISTRO LA COTIZACION Nº" & "  " & cotiza.ToString & "  " & "CORRECTAMENTE")
                    Imprimir_hco.TextBox1.Text = cotiza
                    Imprimir_hco.Show()
                End If
                DataGridView1.Rows.Clear()
                DataGridView2.Rows.Clear()
                DataGridView3.Rows.Clear()
                DataGridView4.Rows.Clear()
                DataGridView5.Rows.Clear()
                DataGridView6.Rows.Clear()
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox5.Text = ""

                TextBox7.Text = ""
                TextBox8.Text = ""
                TextBox9.Text = ""
                'TextBox10.Text = ""
                TextBox11.Text = ""
                TextBox12.Text = ""

                TextBox14.Text = ""
                TextBox15.Text = ""
                TextBox16.Text = ""
                TextBox17.Text = ""
                TextBox18.Text = ""
                TextBox19.Text = ""
                TextBox20.Text = ""
                TextBox21.Text = ""
                Dim func As New FCOTIZACION
                Me.TextBox1.Text = func.CORRELATIVO_COTIZACION() + 1
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
                TextBox1.Select()
                BLOQUEAR_registros()
                limiar()

                Dim func10 As New FCOTIZACION
                Me.TextBox1.Text = func10.CORRELATIVO_COTIZACION() + 1
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
                Dim hj2 As New flog
                Dim hj1 As New vlog

                hj1.gmodulo = "HCOTIZACION-MENU"
                hj1.gcuser = MDIParent1.Label3.Text
                hj1.gaccion = accion
                hj1.gpc = My.Computer.Name
                hj1.gfecha = DateTimePicker1.Value
                hj1.gdetalle = TextBox1.Text
                hj2.insertar_log(hj1)

            End If
            TextBox1.Enabled = True
            TextBox22.Text = 0


        End If

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        HABILITAR()
        Button8.Enabled = True
        TextBox22.Text = 1
        Button11.Enabled = True
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Dim func As New FCOTIZACION
        Dim func2 As New VCOTIZACION
        Dim respuesta As DialogResult

        respuesta = MessageBox.Show("DESEA CANCELAR EL INGRESO DE LA COTIZACION?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            'func2.gid_cotizacion = TextBox30.Text
            'func.eliminar_cotizacion(func2)
            TextBox1.Select()
            TextBox1.ReadOnly = False
            BLOQUEAR_registros()
            limiar()
            Dim ai As New FCOTIZACION
            Dim jh As New VCOTIZACION
            jh.gtipo = ComboBox3.Text
            dt = ai.mostrar_costos(jh)
            If dt.Rows.Count <> 0 Then
                DataGridView7.DataSource = dt
                'TextBox10.Text = DataGridView7.Rows(0).Cells(1).Value
                TextBox6.Text = DataGridView7.Rows(0).Cells(2).Value
                TextBox13.Text = DataGridView7.Rows(0).Cells(3).Value
            End If
            Dim func10 As New FCOTIZACION
            Me.TextBox1.Text = func10.CORRELATIVO_COTIZACION() + 1
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
        Else

        End If
    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged

    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If TextBox7.Text <> 0 Or TextBox17.Text <> 0 Or TextBox19.Text <> 0 Or Trim(ComboBox3.Text) = "" Then


            Dim HU4 As Double

            'PictureBox2.Enabled = True
            Dim ai As New FCOTIZACION
            If ComboBox3.Text = "PRENDAS" Then

                TextBox24.Text = CDbl(TextBox7.Text) + CDbl(TextBox19.Text) + CDbl(TextBox17.Text) + CDbl(TextBox20.Text) + CDbl(TextBox18.Text) + CDbl(TextBox21.Text) + CDbl(TextBox9.Text)
                Dim hj As New VCOTIZACION
                hj.gtipo = 1

                dt = ai.mostrar_costos(hj)
                If dt.Rows.Count <> 0 Then
                    DataGridView7.DataSource = dt

                    TextBox9.Text = (1 * (1 + (DataGridView1.Rows(0).Cells(2).Value / 100))) * DataGridView7.Rows(0).Cells(2).Value
                    TextBox11.Text = Math.Round((1 * (1 + (DataGridView1.Rows(0).Cells(2).Value / 100))) * DataGridView7.Rows(0).Cells(3).Value, 3)
                    TextBox12.Text = Math.Round((1 * (1 + (DataGridView1.Rows(0).Cells(2).Value / 100))) * DataGridView7.Rows(0).Cells(4).Value, 3)
                    TextBox6.Text = Math.Round((1 * (1 + (DataGridView1.Rows(0).Cells(2).Value / 100))) * DataGridView7.Rows(0).Cells(5).Value, 3)
                    TextBox13.Text = DataGridView7.Rows(0).Cells(6).Value

                End If
            End If
            If ComboBox3.Text = "TELA NOTEX" Then
                Dim hj As New VCOTIZACION
                hj.gtipo = 2
                dt = ai.mostrar_costos(hj)
                If dt.Rows.Count <> 0 Then
                    DataGridView7.DataSource = dt
                    'TextBox10.Text = (1 * (1 + (DataGridView1.Rows(0).Cells(2).Value / 100))) * DataGridView7.Rows(0).Cells(1).Value
                    TextBox9.Text = (1 * (1 + (DataGridView1.Rows(0).Cells(2).Value / 100))) * DataGridView7.Rows(0).Cells(2).Value
                    TextBox11.Text = (1 * (1 + (DataGridView1.Rows(0).Cells(2).Value / 100))) * DataGridView7.Rows(0).Cells(3).Value
                    TextBox12.Text = (1 * (1 + (DataGridView1.Rows(0).Cells(2).Value / 100))) * DataGridView7.Rows(0).Cells(4).Value
                    TextBox6.Text = (1 * (1 + (DataGridView1.Rows(0).Cells(2).Value / 100))) * DataGridView7.Rows(0).Cells(5).Value
                    TextBox13.Text = DataGridView7.Rows(0).Cells(6).Value

                End If
            End If
            If ComboBox3.Text = "INDUMENTARIA MEDICA" Then
                Dim hj As New VCOTIZACION
                hj.gtipo = 3
                dt = ai.mostrar_costos(hj)
                If dt.Rows.Count <> 0 Then
                    DataGridView7.DataSource = dt
                    'TextBox10.Text = (1 * (1 + (DataGridView1.Rows(0).Cells(2).Value / 100))) * DataGridView7.Rows(0).Cells(1).Value
                    TextBox9.Text = (1 * (1 + (DataGridView1.Rows(0).Cells(2).Value / 100))) * DataGridView7.Rows(0).Cells(2).Value
                    TextBox11.Text = (1 * (1 + (DataGridView1.Rows(0).Cells(2).Value / 100))) * DataGridView7.Rows(0).Cells(3).Value
                    TextBox12.Text = (1 * (1 + (DataGridView1.Rows(0).Cells(2).Value / 100))) * DataGridView7.Rows(0).Cells(4).Value
                    TextBox6.Text = (1 * (1 + (DataGridView1.Rows(0).Cells(2).Value / 100))) * DataGridView7.Rows(0).Cells(5).Value
                    TextBox13.Text = DataGridView7.Rows(0).Cells(6).Value

                End If
            End If
            If ComboBox3.Text = "TELA PUNTO" Then
                Dim hj As New VCOTIZACION
                hj.gtipo = 4
                dt = ai.mostrar_costos(hj)
                If dt.Rows.Count <> 0 Then
                    DataGridView7.DataSource = dt
                    'TextBox10.Text = (1 * (1 + (DataGridView1.Rows(0).Cells(2).Value / 100))) * DataGridView7.Rows(0).Cells(1).Value
                    TextBox9.Text = (1 * (1 + (DataGridView1.Rows(0).Cells(2).Value / 100))) * DataGridView7.Rows(0).Cells(2).Value
                    TextBox11.Text = (1 * (1 + (DataGridView1.Rows(0).Cells(2).Value / 100))) * DataGridView7.Rows(0).Cells(3).Value
                    TextBox12.Text = (1 * (1 + (DataGridView1.Rows(0).Cells(2).Value / 100))) * DataGridView7.Rows(0).Cells(4).Value
                    TextBox6.Text = (1 * (1 + (DataGridView1.Rows(0).Cells(2).Value / 100))) * DataGridView7.Rows(0).Cells(5).Value
                    TextBox13.Text = DataGridView7.Rows(0).Cells(6).Value

                End If
            End If

            TextBox23.Text = Convert.ToDouble(TextBox9.Text) + 0 + Convert.ToDouble(TextBox11.Text) + Convert.ToDouble(TextBox6.Text) + Convert.ToDouble(TextBox12.Text) + Convert.ToDouble(TextBox7.Text).ToString("N3") + Convert.ToDouble(TextBox17.Text) +
            Convert.ToDouble(TextBox18.Text) + Convert.ToDouble(TextBox19.Text) + Convert.ToDouble(TextBox20.Text) + Convert.ToDouble(TextBox21.Text)
            HU4 = Convert.ToDouble(TextBox23.Text) + (Convert.ToDouble(TextBox23.Text) * (Convert.ToDouble(TextBox13.Text)) / 100)
            TextBox14.Text = Math.Round(HU4, 3)
            TextBox15.Text = Math.Round(Convert.ToDouble(TextBox14.Text) * 0.18, 3)
            TextBox16.Text = Math.Round(Convert.ToDouble(TextBox14.Text) + Convert.ToDouble(TextBox15.Text), 3)
            Button12.Enabled = True
        Else
            MsgBox("FALTA ALGUNOS DATOS OBLIGATORIOS *")
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
    Dim dt8, dt9, dt10, dt11, dt12, dt13, dt14 As New DataTable
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim I, i10, i11 As Integer
            Dim ml As New FCOTIZACION
            Dim ml1 As New VCOTIZACION
            'habilitar_registros()

            'TextBox11.Select()
            I = Len(TextBox1.Text.ToString().Trim())
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
            ml1.gid_cotizacion = TextBox1.Text
            MsgBox(ml1.gid_cotizacion)
            MsgBox(ml.mostrar_COTIZACION(ml1))
            'If ml.mostrar_COTIZACION(ml1) > 0 Then

            '    Dim mv As New ftela
            '    Dim mv1 As New vTela

            '    mv1.gid_cotizacion = TextBox1.Text
            '    dt8 = mv.mostrar_tela(mv1)
            '    If dt8.Rows.Count <> 0 Then

            '        DataGridView8.DataSource = dt8
            '        i10 = DataGridView8.Rows.Count
            '        DataGridView1.Rows.Add(i10)
            '        For i2 = 0 To i10 - 1
            '            DataGridView1.Rows(i2).Cells(0).Value = DataGridView8.Rows(i2).Cells(1).Value
            '            DataGridView1.Rows(i2).Cells(1).Value = DataGridView8.Rows(i2).Cells(2).Value
            '            DataGridView1.Rows(i2).Cells(2).Value = DataGridView8.Rows(i2).Cells(3).Value
            '            DataGridView1.Rows(i2).Cells(3).Value = DataGridView8.Rows(i2).Cells(4).Value
            '            DataGridView1.Rows(i2).Cells(4).Value = DataGridView8.Rows(i2).Cells(5).Value
            '            DataGridView1.Rows(i2).Cells(6).Value = DataGridView8.Rows(i2).Cells(6).Value
            '            DataGridView1.Rows(i2).Cells(7).Value = DataGridView8.Rows(i2).Cells(7).Value
            '            DataGridView1.Rows(i2).Cells(5).Value = DataGridView8.Rows(i2).Cells(8).Value
            '        Next
            '    End If

            '    Dim mc As New favios_costura
            '    Dim mc1 As New vAvios_Costura

            '    mc1.gid_cotizacion = TextBox1.Text
            '    dt9 = mc.mostrar_avios_costura(mc1)
            '    If dt9.Rows.Count <> 0 Then

            '        DataGridView9.DataSource = dt9
            '        i11 = DataGridView9.Rows.Count
            '        DataGridView2.Rows.Add(i11)
            '        For i3 = 0 To i11 - 1
            '            DataGridView2.Rows(i3).Cells(0).Value = DataGridView9.Rows(i3).Cells(1).Value
            '            DataGridView2.Rows(i3).Cells(1).Value = DataGridView9.Rows(i3).Cells(2).Value
            '            DataGridView2.Rows(i3).Cells(2).Value = DataGridView9.Rows(i3).Cells(3).Value
            '            DataGridView2.Rows(i3).Cells(3).Value = DataGridView9.Rows(i3).Cells(4).Value
            '            DataGridView2.Rows(i3).Cells(4).Value = DataGridView9.Rows(i3).Cells(5).Value
            '            DataGridView2.Rows(i3).Cells(6).Value = DataGridView9.Rows(i3).Cells(6).Value
            '            DataGridView2.Rows(i3).Cells(7).Value = DataGridView9.Rows(i3).Cells(7).Value
            '            DataGridView2.Rows(i3).Cells(5).Value = DataGridView9.Rows(i3).Cells(8).Value
            '        Next
            '    End If

            '    Dim mk As New favios_acabados
            '    Dim mk1 As New vAvios_Acabados

            '    mk1.gid_cotizacion = TextBox1.Text
            '    dt10 = mk.mostrar_avios_acabados(mk1)
            '    If dt10.Rows.Count <> 0 Then

            '        DataGridView10.DataSource = dt10
            '        Dim i12 As Integer
            '        i12 = DataGridView10.Rows.Count
            '        DataGridView3.Rows.Add(i12)
            '        For i4 = 0 To i12 - 1
            '            DataGridView3.Rows(i4).Cells(0).Value = DataGridView10.Rows(i4).Cells(1).Value
            '            DataGridView3.Rows(i4).Cells(1).Value = DataGridView10.Rows(i4).Cells(2).Value
            '            DataGridView3.Rows(i4).Cells(2).Value = DataGridView10.Rows(i4).Cells(3).Value
            '            DataGridView3.Rows(i4).Cells(3).Value = DataGridView10.Rows(i4).Cells(4).Value
            '            DataGridView3.Rows(i4).Cells(4).Value = DataGridView10.Rows(i4).Cells(5).Value
            '            DataGridView3.Rows(i4).Cells(6).Value = DataGridView10.Rows(i4).Cells(6).Value
            '            DataGridView3.Rows(i4).Cells(7).Value = DataGridView10.Rows(i4).Cells(7).Value
            '            DataGridView3.Rows(i4).Cells(5).Value = DataGridView10.Rows(i4).Cells(8).Value
            '        Next
            '    End If


            '    Dim mh As New flavados
            '    Dim mh1 As New vLavados

            '    mh1.gid_cotizacion = TextBox1.Text
            '    dt11 = mh.mostrar_lavados(mh1)
            '    If dt11.Rows.Count <> 0 Then

            '        DataGridView11.DataSource = dt11
            '        Dim i22 As Integer
            '        i22 = DataGridView11.Rows.Count
            '        DataGridView4.Rows.Add(i22)
            '        For i44 = 0 To i22 - 1
            '            DataGridView4.Rows(i44).Cells(0).Value = DataGridView11.Rows(i44).Cells(1).Value
            '            DataGridView4.Rows(i44).Cells(1).Value = DataGridView11.Rows(i44).Cells(2).Value
            '            DataGridView4.Rows(i44).Cells(2).Value = DataGridView11.Rows(i44).Cells(4).Value
            '            DataGridView4.Rows(i44).Cells(3).Value = DataGridView11.Rows(i44).Cells(5).Value
            '            DataGridView4.Rows(i44).Cells(4).Value = DataGridView11.Rows(i44).Cells(7).Value
            '            DataGridView4.Rows(i44).Cells(5).Value = DataGridView11.Rows(i44).Cells(3).Value
            '            DataGridView4.Rows(i44).Cells(6).Value = DataGridView11.Rows(i44).Cells(6).Value
            '        Next
            '    End If

            '    Dim mf As New faplicaciones
            '    Dim mf1 As New Aplicaciones

            '    mf1.gid_cotizacion = TextBox1.Text
            '    dt12 = mf.mostrar_aplicaciones(mf1)
            '    If dt12.Rows.Count <> 0 Then

            '        DataGridView12.DataSource = dt12
            '        Dim i13 As Integer
            '        i13 = DataGridView12.Rows.Count
            '        DataGridView5.Rows.Add(i13)
            '        For i5 = 0 To i13 - 1
            '            DataGridView5.Rows(i5).Cells(0).Value = DataGridView12.Rows(i5).Cells(1).Value
            '            DataGridView5.Rows(i5).Cells(1).Value = DataGridView12.Rows(i5).Cells(2).Value
            '            DataGridView5.Rows(i5).Cells(2).Value = DataGridView12.Rows(i5).Cells(3).Value
            '            DataGridView5.Rows(i5).Cells(3).Value = DataGridView12.Rows(i5).Cells(4).Value
            '            DataGridView5.Rows(i5).Cells(4).Value = DataGridView12.Rows(i5).Cells(7).Value
            '            DataGridView5.Rows(i5).Cells(5).Value = DataGridView12.Rows(i5).Cells(5).Value
            '            DataGridView5.Rows(i5).Cells(6).Value = DataGridView12.Rows(i5).Cells(6).Value
            '        Next
            '    End If

            '    Dim my As New fmano_obra
            '    Dim my1 As New vManoObra

            '    my1.gid_cotizacion = TextBox1.Text
            '    dt13 = my.mostrar_manoobra(my1)
            '    If dt13.Rows.Count <> 0 Then

            '        DataGridView13.DataSource = dt13
            '        Dim i14 As Integer
            '        i14 = DataGridView13.Rows.Count
            '        DataGridView6.Rows.Add(i14)
            '        For i6 = 0 To i14 - 1

            '            DataGridView6.Rows(i6).Cells(0).Value = DataGridView13.Rows(i6).Cells(1).Value
            '            DataGridView6.Rows(i6).Cells(1).Value = DataGridView13.Rows(i6).Cells(5).Value
            '            DataGridView6.Rows(i6).Cells(2).Value = DataGridView13.Rows(i6).Cells(2).Value
            '            DataGridView6.Rows(i6).Cells(3).Value = DataGridView13.Rows(i6).Cells(3).Value
            '        Next
            '    End If


            '    Dim mr As New FCOTIZACION
            '    Dim mr1 As New VCOTIZACION

            '    mr1.gid_cotizacion = TextBox1.Text
            '    dt14 = mr.mostrar_COTIZACION2(mr1)
            '    If dt14.Rows.Count <> 0 Then

            '        cotizacion.DataSource = dt14
            '        If cotizacion.Rows.Count <> 0 Then
            '            TextBox2.Text = cotizacion.Rows(0).Cells(1).Value
            '            TextBox3.Text = cotizacion.Rows(0).Cells(2).Value
            '            TextBox4.Text = cotizacion.Rows(0).Cells(3).Value
            '            TextBox5.Text = cotizacion.Rows(0).Cells(4).Value
            '            TextBox8.Text = cotizacion.Rows(0).Cells(5).Value
            '            If cotizacion.Rows(0).Cells(6).Value = 1 Then
            '                ComboBox1.Text = "SOLES"
            '            Else
            '                ComboBox1.Text = "DOLARES"
            '            End If
            '            TextBox9.Text = cotizacion.Rows(0).Cells(8).Value
            '            'TextBox10.Text = cotizacion.Rows(0).Cells(8).Value
            '            TextBox11.Text = cotizacion.Rows(0).Cells(10).Value
            '            TextBox6.Text = cotizacion.Rows(0).Cells(11).Value
            '            TextBox12.Text = cotizacion.Rows(0).Cells(12).Value
            '            TextBox13.Text = cotizacion.Rows(0).Cells(14).Value
            '            TextBox14.Text = cotizacion.Rows(0).Cells(15).Value
            '            TextBox15.Text = cotizacion.Rows(0).Cells(16).Value
            '            TextBox16.Text = cotizacion.Rows(0).Cells(17).Value
            '            TextBox7.Text = cotizacion.Rows(0).Cells(20).Value
            '            TextBox19.Text = cotizacion.Rows(0).Cells(21).Value
            '            TextBox17.Text = cotizacion.Rows(0).Cells(22).Value
            '            TextBox20.Text = cotizacion.Rows(0).Cells(23).Value

            '            TextBox18.Text = cotizacion.Rows(0).Cells(24).Value
            '            TextBox21.Text = cotizacion.Rows(0).Cells(25).Value
            '            TextBox10.Text = cotizacion.Rows(0).Cells(28).Value
            '            TextBox24.Text = cotizacion.Rows(0).Cells(9).Value
            '            TextBox23.Text = cotizacion.Rows(0).Cells(13).Value
            '            'DateTimePicker1.Value = cotizacion.Rows(0).Cells(25).Value
            '        End If



            '        'Select Case cotizacion.Rows(0).Cells(10).Value

            '        '    Case "1" : Label6.Text = "DESAPROBADO"
            '        '    Case "2" : Label6.Text = "APROBADO"

            '        'End Select
            '        'If Label6.Text = "APROBADO" Then
            '        '    MsgBox("LA COTIZACION ESTA APROBADO POR GERENCIA")
            '        '    BLOQUEAR_registros()
            '        '    TextBox11.Enabled = False
            '        '    PictureBox5.Enabled = True
            '        'End If
            '    End If
            '    'BLOQUEAR_registros()
            '    TextBox1.ReadOnly = True
            '    MsgBox("paso1")
            'Else

            '    MsgBox("paso2")
            '    TextBox1.ReadOnly = True
            '    Button11.Enabled = True
            '    'limpiar_registros()
            '    'TextBox11.Text = ""
            '    'TextBox11.Enabled = True
            'End If

        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        'e.Handled = Not IsNumeric(e.KeyChar) And Not Char.IsControl(e.KeyChar)
    End Sub

    Private Sub TextBox14_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox14.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox10.Text = Math.Round(CDbl(TextBox14.Text) - CDbl(TextBox23.Text), 3)
            TextBox13.Text = Math.Round((CDbl(TextBox10.Text) * 100) / CDbl(TextBox14.Text), 3)
            TextBox16.Text = Math.Round(CDbl(TextBox14.Text) * 1.18, 1)
            TextBox15.Text = Math.Round(CDbl(TextBox16.Text) - CDbl(TextBox14.Text), 3)
        End If
    End Sub
End Class