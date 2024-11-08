Imports System.Data.SqlClient
Imports System.Net.Mail

Public Class CotizacionUdp
    Public Property Padre As CotizacionUdp
    Public conx As SqlConnection
    Dim Rsr3 As SqlDataReader
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim fc As New Form_Cotizacion
        fc.Padre = Me
        fc.TextBox3.Text = 1
        Dim va As Integer
        va = DataGridView1.Rows.Count
        If va > 0 Then
            fc.DataGridView1.Rows.Clear()
            fc.DataGridView1.Rows.Add(va)
            For i = 0 To va - 1
                fc.DataGridView1.Rows(i).Cells(0).Value = DataGridView1.Rows(i).Cells(0).Value
                fc.DataGridView1.Rows(i).Cells(1).Value = DataGridView1.Rows(i).Cells(1).Value
                fc.DataGridView1.Rows(i).Cells(2).Value = DataGridView1.Rows(i).Cells(2).Value
                fc.DataGridView1.Rows(i).Cells(3).Value = DataGridView1.Rows(i).Cells(3).Value
                fc.DataGridView1.Rows(i).Cells(4).Value = DataGridView1.Rows(i).Cells(4).Value
                fc.DataGridView1.Rows(i).Cells(5).Value = DataGridView1.Rows(i).Cells(5).Value
                fc.DataGridView1.Rows(i).Cells(6).Value = DataGridView1.Rows(i).Cells(6).Value
                fc.DataGridView1.Rows(i).Cells(7).Value = DataGridView1.Rows(i).Cells(7).Value
                fc.DataGridView1.Rows(i).Cells(8).Value = DataGridView1.Rows(i).Cells(8).Value
            Next
            If Label12.Text = "2" Then
                fc.DataGridView1.Columns(0).ReadOnly = True
                fc.DataGridView1.Columns(1).ReadOnly = True
                fc.DataGridView1.Columns(2).ReadOnly = True
                fc.DataGridView1.Columns(3).ReadOnly = True
                fc.DataGridView1.Columns(4).ReadOnly = True
                fc.DataGridView1.Columns(5).ReadOnly = True
                fc.DataGridView1.Columns(7).ReadOnly = True
                fc.DataGridView1.Columns(8).ReadOnly = True
            Else
                If Label12.Text = "1" Then
                    fc.DataGridView1.Columns(0).ReadOnly = False
                    fc.DataGridView1.Columns(1).ReadOnly = False
                    fc.DataGridView1.Columns(2).ReadOnly = False
                    fc.DataGridView1.Columns(3).ReadOnly = False
                    fc.DataGridView1.Columns(4).ReadOnly = False
                    fc.DataGridView1.Columns(5).ReadOnly = False
                    fc.DataGridView1.Columns(7).ReadOnly = False
                    fc.DataGridView1.Columns(8).ReadOnly = False
                End If
            End If
        End If
        fc.ShowDialog()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim fc1 As New Form_Cotizacion()
        fc1.Padre = Me
        fc1.DataGridView1.Rows.Clear()
        fc1.TextBox3.Text = 2
        Dim va As Integer
        va = DataGridView2.Rows.Count
        If va > 0 Then

            fc1.DataGridView1.Rows.Add(va)
            For i = 0 To va - 1
                fc1.DataGridView1.Rows(i).Cells(0).Value = DataGridView2.Rows(i).Cells(0).Value
                fc1.DataGridView1.Rows(i).Cells(1).Value = DataGridView2.Rows(i).Cells(1).Value
                fc1.DataGridView1.Rows(i).Cells(2).Value = DataGridView2.Rows(i).Cells(2).Value
                fc1.DataGridView1.Rows(i).Cells(3).Value = DataGridView2.Rows(i).Cells(3).Value
                fc1.DataGridView1.Rows(i).Cells(4).Value = DataGridView2.Rows(i).Cells(4).Value
                fc1.DataGridView1.Rows(i).Cells(5).Value = DataGridView2.Rows(i).Cells(5).Value
                fc1.DataGridView1.Rows(i).Cells(6).Value = DataGridView2.Rows(i).Cells(6).Value
                fc1.DataGridView1.Rows(i).Cells(7).Value = DataGridView2.Rows(i).Cells(7).Value
                fc1.DataGridView1.Rows(i).Cells(8).Value = DataGridView2.Rows(i).Cells(8).Value
            Next
            If Label12.Text = "2" Then
                fc1.DataGridView1.Columns(0).ReadOnly = True
                fc1.DataGridView1.Columns(1).ReadOnly = True
                fc1.DataGridView1.Columns(2).ReadOnly = True
                fc1.DataGridView1.Columns(3).ReadOnly = True
                fc1.DataGridView1.Columns(4).ReadOnly = True
                fc1.DataGridView1.Columns(5).ReadOnly = True
                fc1.DataGridView1.Columns(7).ReadOnly = True
                fc1.DataGridView1.Columns(8).ReadOnly = True
            Else
                If Label12.Text = "1" Then
                    fc1.DataGridView1.Columns(0).ReadOnly = False
                    fc1.DataGridView1.Columns(1).ReadOnly = False
                    fc1.DataGridView1.Columns(2).ReadOnly = False
                    fc1.DataGridView1.Columns(3).ReadOnly = False
                    fc1.DataGridView1.Columns(4).ReadOnly = False
                    fc1.DataGridView1.Columns(5).ReadOnly = False
                    fc1.DataGridView1.Columns(7).ReadOnly = False
                    fc1.DataGridView1.Columns(8).ReadOnly = False
                End If
            End If

        End If
        fc1.ShowDialog()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim fc2 As New Form_Cotizacion()
        fc2.Padre = Me
        fc2.DataGridView1.Rows.Clear()
        fc2.TextBox3.Text = 3
        Dim va As Integer
        va = DataGridView3.Rows.Count
        If va > 0 Then

            fc2.DataGridView1.Rows.Add(va)
            For i = 0 To va - 1
                fc2.DataGridView1.Rows(i).Cells(0).Value = DataGridView3.Rows(i).Cells(0).Value
                fc2.DataGridView1.Rows(i).Cells(1).Value = DataGridView3.Rows(i).Cells(1).Value
                fc2.DataGridView1.Rows(i).Cells(2).Value = DataGridView3.Rows(i).Cells(2).Value
                fc2.DataGridView1.Rows(i).Cells(3).Value = DataGridView3.Rows(i).Cells(3).Value
                fc2.DataGridView1.Rows(i).Cells(4).Value = DataGridView3.Rows(i).Cells(4).Value
                fc2.DataGridView1.Rows(i).Cells(5).Value = DataGridView3.Rows(i).Cells(5).Value
                fc2.DataGridView1.Rows(i).Cells(6).Value = DataGridView3.Rows(i).Cells(6).Value
                fc2.DataGridView1.Rows(i).Cells(7).Value = DataGridView3.Rows(i).Cells(7).Value
                fc2.DataGridView1.Rows(i).Cells(8).Value = DataGridView3.Rows(i).Cells(8).Value
            Next
            If Label12.Text = "2" Then
                fc2.DataGridView1.Columns(0).ReadOnly = True
                fc2.DataGridView1.Columns(1).ReadOnly = True
                fc2.DataGridView1.Columns(2).ReadOnly = True
                fc2.DataGridView1.Columns(3).ReadOnly = True
                fc2.DataGridView1.Columns(4).ReadOnly = True
                fc2.DataGridView1.Columns(5).ReadOnly = True
                fc2.DataGridView1.Columns(7).ReadOnly = True
                fc2.DataGridView1.Columns(8).ReadOnly = True
            Else
                If Label12.Text = "1" Then
                    fc2.DataGridView1.Columns(0).ReadOnly = False
                    fc2.DataGridView1.Columns(1).ReadOnly = False
                    fc2.DataGridView1.Columns(2).ReadOnly = False
                    fc2.DataGridView1.Columns(3).ReadOnly = False
                    fc2.DataGridView1.Columns(4).ReadOnly = False
                    fc2.DataGridView1.Columns(5).ReadOnly = False
                    fc2.DataGridView1.Columns(7).ReadOnly = False
                    fc2.DataGridView1.Columns(8).ReadOnly = False
                End If
            End If
        End If
        fc2.ShowDialog()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim ct2 As New Form_Cotizacion2()
        ct2.Padre2 = Me
        ct2.TextBox2.Text = 1
        Dim va As Integer
        va = DataGridView4.Rows.Count
        If va > 0 Then
            ct2.DataGridView1.Rows.Clear()
            ct2.DataGridView1.Rows.Add(va)


            For i = 0 To va - 1
                ct2.DataGridView1.Rows(i).Cells(0).Value = DataGridView4.Rows(i).Cells(0).Value
                ct2.DataGridView1.Rows(i).Cells(1).Value = DataGridView4.Rows(i).Cells(1).Value
                ct2.DataGridView1.Rows(i).Cells(2).Value = DataGridView4.Rows(i).Cells(2).Value
                ct2.DataGridView1.Rows(i).Cells(3).Value = DataGridView4.Rows(i).Cells(3).Value
                ct2.DataGridView1.Rows(i).Cells(4).Value = DataGridView4.Rows(i).Cells(4).Value
                ct2.DataGridView1.Rows(i).Cells(5).Value = DataGridView4.Rows(i).Cells(5).Value
                ct2.DataGridView1.Rows(i).Cells(6).Value = DataGridView4.Rows(i).Cells(6).Value
                ct2.DataGridView1.Rows(i).Cells(7).Value = DataGridView4.Rows(i).Cells(7).Value
            Next
        End If
        ct2.ShowDialog()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim ct3 As New Form_Cotizacion2()
        ct3.Padre2 = Me
        ct3.TextBox2.Text = 2
        Dim va As Integer
        va = DataGridView5.Rows.Count
        If va > 0 Then
            ct3.DataGridView1.Rows.Clear()
            ct3.DataGridView1.Rows.Add(va)
            For i = 0 To va - 1
                ct3.DataGridView1.Rows(i).Cells(0).Value = DataGridView5.Rows(i).Cells(0).Value
                ct3.DataGridView1.Rows(i).Cells(1).Value = DataGridView5.Rows(i).Cells(1).Value
                ct3.DataGridView1.Rows(i).Cells(2).Value = DataGridView5.Rows(i).Cells(2).Value
                ct3.DataGridView1.Rows(i).Cells(3).Value = DataGridView5.Rows(i).Cells(3).Value
                ct3.DataGridView1.Rows(i).Cells(4).Value = DataGridView5.Rows(i).Cells(4).Value
                ct3.DataGridView1.Rows(i).Cells(5).Value = DataGridView5.Rows(i).Cells(5).Value
                ct3.DataGridView1.Rows(i).Cells(6).Value = DataGridView5.Rows(i).Cells(6).Value
                ct3.DataGridView1.Rows(i).Cells(7).Value = DataGridView5.Rows(i).Cells(7).Value
            Next
        End If
        ct3.ShowDialog()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim dt As New Form_Cotizacion3
        dt.padre = Me
        Dim va As Integer
        va = DataGridView6.Rows.Count

        If va > 0 Then
            dt.DataGridView1.Rows.Clear()
            dt.DataGridView1.Rows.Add(va)


            For i = 0 To va - 1
                dt.DataGridView1.Rows(i).Cells(0).Value = DataGridView6.Rows(i).Cells(0).Value
                dt.DataGridView1.Rows(i).Cells(1).Value = DataGridView6.Rows(i).Cells(1).Value
                dt.DataGridView1.Rows(i).Cells(2).Value = DataGridView6.Rows(i).Cells(2).Value
                dt.DataGridView1.Rows(i).Cells(3).Value = DataGridView6.Rows(i).Cells(3).Value
                dt.DataGridView1.Rows(i).Cells(4).Value = DataGridView6.Rows(i).Cells(4).Value
                dt.DataGridView1.Rows(i).Cells(5).Value = DataGridView6.Rows(i).Cells(5).Value
            Next
        End If
        'Form_Cotizacion3.TextBox1.Text = DataGridView6.Rows(0).Cells(2).Value
        dt.ShowDialog()
    End Sub
    Sub limiar()
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""

        TextBox7.Text = 0
        TextBox17.Text = 0
        TextBox18.Text = 0
        TextBox19.Text = 0
        TextBox20.Text = 0
        TextBox21.Text = 0
        TextBox1.Select()
        BLOQUEAR_registros()
    End Sub
    Sub limpiarData()
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
        DataGridView3.Rows.Clear()
        DataGridView4.Rows.Clear()
        DataGridView5.Rows.Clear()
        DataGridView13.Rows.Clear()
        DataGridView8.DataSource = ""
        DataGridView9.DataSource = ""
        DataGridView10.DataSource = ""
        DataGridView11.DataSource = ""
        DataGridView12.DataSource = ""


    End Sub
    Sub BLOQUEAR_registros()

        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox8.Enabled = False

        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        Button7.Enabled = False

        PictureBox1.Enabled = False
        Button11.Enabled = False
        Button12.Enabled = False
        Button13.Enabled = False
        Button14.Enabled = False
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
        Else

        End If
    End Sub
    Sub HABILITAR()
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox8.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        Button5.Enabled = True
        Button6.Enabled = True
        Button7.Enabled = True
        Button13.Enabled = True
    End Sub
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        HABILITAR()

        TextBox22.Text = 1
        Button11.Enabled = True
    End Sub
    Dim Rsr2125 As SqlDataReader
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

        'Dim primerCaracter As String = TextBox99.Text.Substring(0, 1)
        'If primerCaracter = "Z" Then
        '    MsgBox("Es z")
        'End If

        Dim ui As New VCOTIZACION
        Dim fz As New FCOTIZACION
        Dim cotiza As String
        Dim accion, Area As String
        If TextBox22.Text = 1 Then
            cotiza = TextBox10.Text
            ui.gid_cotizacion = TextBox10.Text
            fz.eliminar_cotizacion(ui)
            accion = 2
        Else
            ui.gid_cotizacion = TextBox10.Text
            If fz.mostrar_COTIZACION(ui) >= 1 Then
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
                MsgBox("LA COTIZACION Nº" & "  " & TextBox10.Text.ToString & "  " & "YA ESTA REGISTRADA", vbExclamation)

                MsgBox("SE ACTUALIZARA A LA COTIZACION Nº" & "  " & cotiza.ToString, vbInformation)
                accion = 1
            Else
                cotiza = TextBox10.Text
                accion = 1
            End If
        End If
        abrir()

        Dim sql1 As String = "SELECT count(id_cotizacion) FROM HojaCotizacion WHERE od ='" + TextBox1.Text.ToString().Trim() + "' and od_version ='" + TextBox6.Text.ToString().Trim() + "'"
        Dim cmd1 As New SqlCommand(sql1, conx)
        Rsr2125 = cmd1.ExecuteReader()
        If Rsr2125.Read() = True And Convert.ToInt32(Rsr2125(0)) > 0 And TextBox22.Text = 0 Then
            MsgBox("LA Od: " + TextBox1.Text.ToString().Trim() & "Vs : " & TextBox6.Text.ToString().Trim() & " ya esta Registrada")
        Else
            ui.gid_cotizacion = cotiza
            ui.gcliente = TextBox2.Text
            ui.gdescripcion = TextBox3.Text
            ui.gestilo = TextBox4.Text
            ui.grango_tallas = TextBox5.Text
            ui.gti_cambio = TextBox13.Text
            If TextBox11.Text = "SOLES" Then
                ui.gmoneda = 1
            Else
                ui.gmoneda = 2
            End If
            ui.glinea = 1
            ui.ggasto_logistico = 0
            ui.ggasto_operativo = 0
            ui.ggasto_administrativo = 0
            ui.ggasto_financiero = 0
            ui.ggasto_venta = 0
            ui.gcosto_producto = 0
            ui.gmargen_utilidad = 0
            ui.gmargen_utilidad_moneda = 0
            ui.gvalor_venta = 0
            ui.gigv = 0
            ui.gprecio_venta = 0
            If CheckBox1.Checked = True Then
                If Label12.Text = 1 Then
                    ui.gestado = 2
                Else
                    If Label12.Text = 3 Then
                        ui.gestado = 4
                    Else
                        If Label12.Text = 5 Then
                            ui.gestado = 6
                        End If
                    End If
                End If

            Else
                ui.gestado = Label12.Text
            End If

            ui.gimagen = "POR DENIFIR"
            ui.gcostot_tela = TextBox7.Text
            ui.gcostot_AviosC = TextBox19.Text
            ui.gcostot_AviosA = TextBox17.Text
            ui.gcostot_Lavado = TextBox20.Text
            ui.gcostot_Aplicaciones = TextBox18.Text
            ui.gcostot_MObra = TextBox21.Text
            ui.gvendedor = TextBox12.Text
            ui.gfecha = DateTimePicker1.Value
            ui.god = Trim(TextBox1.Text)
            ui.god_version = Trim(TextBox6.Text)
            ui.gPm = Trim(TextBox9.Text)
            ui.gTela = Trim(TextBox8.Text)
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
                    FT2.gconsumo = DataGridView6.Rows(I5).Cells(1).Value
                    FT2.gunidad = DataGridView6.Rows(I5).Cells(2).Value
                    FT2.gconsumo_total = DataGridView6.Rows(I5).Cells(3).Value
                    FT2.gcosto_unitario = DataGridView6.Rows(I5).Cells(4).Value
                    FT2.gcosto_total = DataGridView6.Rows(I5).Cells(5).Value
                    FT2.gid_cotizacion = cotiza


                    FZ6.insertar_mano_obra(FT2)
                Next
                'Dim resultado As String
                'resultado = MsgBox("¿Quieres imprimir la Cotizacion?", vbYesNo, "Colorear A1")

                'If resultado = vbYes Then
                MsgBox("SE REGISTRO LA COTIZACION Nº" & "  " & cotiza.ToString & "  " & "CORRECTAMENTE")
                If CheckBox1.Checked = True Then
                    ENVIAR_CORREO_UDP()
                End If

                HojaContizacionAlm.validar()
                '    Imprimir_hco.TextBox1.Text = cotiza
                '    Imprimir_hco.Show()
                'End If
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
                TextBox17.Text = ""
                TextBox18.Text = ""
                TextBox19.Text = ""
                TextBox20.Text = ""
                TextBox21.Text = ""
                Label11.Text = "0"
                'Dim func As New FCOTIZACION
                'Me.TextBox10.Text = func.CORRELATIVO_COTIZACION() + 1
                'Select Case TextBox10.Text.Length

                '    Case "1" : TextBox10.Text = "0000000" & "" & TextBox10.Text
                '    Case "2" : TextBox10.Text = "000000" & "" & TextBox10.Text
                '    Case "3" : TextBox10.Text = "00000" & "" & TextBox10.Text
                '    Case "4" : TextBox10.Text = "0000" & "" & TextBox10.Text
                '    Case "5" : TextBox10.Text = "000" & "" & TextBox10.Text
                '    Case "6" : TextBox10.Text = "00" & "" & TextBox10.Text
                '    Case "7" : TextBox10.Text = "0" & "" & TextBox10.Text
                '    Case "8" : TextBox10.Text = TextBox10.Text
                'End Select
                AlmacenCotizacion.Button1.PerformClick()
                TextBox10.Select()
                SendKeys.Send("{ENTER}")
                'BLOQUEAR_registros()
                'limiar()

                'Dim func10 As New FCOTIZACION
                'Me.TextBox10.Text = func10.CORRELATIVO_COTIZACION() + 1
                'Select Case TextBox10.Text.Length

                '    Case "1" : TextBox10.Text = "0000000" & "" & TextBox10.Text
                '    Case "2" : TextBox10.Text = "000000" & "" & TextBox10.Text
                '    Case "3" : TextBox10.Text = "00000" & "" & TextBox10.Text
                '    Case "4" : TextBox10.Text = "0000" & "" & TextBox10.Text
                '    Case "5" : TextBox10.Text = "000" & "" & TextBox10.Text
                '    Case "6" : TextBox10.Text = "00" & "" & TextBox10.Text
                '    Case "7" : TextBox10.Text = "0" & "" & TextBox10.Text
                '    Case "8" : TextBox10.Text = TextBox10.Text
                'End Select
                Dim hj2 As New flog
                Dim hj1 As New vlog

                hj1.gmodulo = "HCOTIZACION-MENU"
                hj1.gcuser = MDIParent1.Label3.Text
                hj1.gaccion = accion
                hj1.gpc = My.Computer.Name
                hj1.gfecha = DateTimePicker1.Value
                hj1.gdetalle = TextBox10.Text
                hj1.gccia = "01"
                hj2.insertar_log(hj1)
            End If
            TextBox10.Enabled = True
            TextBox22.Text = 0
        End If
        Rsr2125.Close()


    End Sub
    Sub habilitar_registros()
        TextBox1.Enabled = False
        TextBox2.Enabled = True
        TextBox3.Enabled = True

        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox8.Enabled = True

        TextBox6.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        Button5.Enabled = True
        Button6.Enabled = True
        Button7.Enabled = True

        PictureBox1.Enabled = False

        Button12.Enabled = False
        Button13.Enabled = False
        Button14.Enabled = True
    End Sub
    Dim dt1, dt2, dt3, dt4, dt5, dt6, dt7, dt8, dt9, dt10, dt11, dt12, dt13, dt14 As DataTable

    Private Sub Button39_Click(sender As Object, e As EventArgs) Handles Button39.Click
        If TextBox99.Text.ToString().Trim() <> "" Then
            Process.Start(Trim(TextBox99.Text))
        End If

    End Sub

    Private Sub Button38_Click(sender As Object, e As EventArgs) Handles Button38.Click
        OpenFileDialog1.ShowDialog()
        TextBox99.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Process.Start(Trim(TextBox5.Text))
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Select Case Label12.Text
            Case "1" : RptFormCotizacion.TextBox2.Text = "1"
            Case "2" : RptFormCotizacion.TextBox2.Text = "1"
            Case "3" : RptFormCotizacion.TextBox2.Text = "1"
            Case "4" : RptFormCotizacion.TextBox2.Text = "2"
        End Select
        RptFormCotizacion.TextBox1.Text = TextBox10.Text
        RptFormCotizacion.ShowDialog()
    End Sub

    Private Sub CotizacionUdp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox22.Text = 0
        TextBox7.Text = 0.00
        TextBox19.Text = 0.00
        TextBox17.Text = 0.00
        TextBox18.Text = 0.00
        TextBox20.Text = 0.00
        TextBox21.Text = 0.00
        CheckBox1.Checked = False
        BLOQUEAR_registros()
        Dim func12 As New ftcambio
        Dim dts34 As New vtcambio
        dts34.gfecha = DateTimePicker1.Text
        TextBox13.Text = func12.mostrar_tipo_cambio(dts34)
        'DataGridView3.Rows.Clear()
        'cargarInformacionAviosAcabados()
        TextBox10.Select()
    End Sub
    Sub cargartela()
        If TextBox14.Text.ToString().Trim().Length > 0 Then

            If TextBox22.Text = 0 Then
                DataGridView1.Rows.Add()
                Dim va As Integer = DataGridView1.Rows.Count
                DataGridView1.Rows(va - 1).Cells(0).Value = TextBox14.Text.ToString().Trim()
                DataGridView1.Rows(va - 1).Cells(1).Value = TextBox8.Text.ToString().Trim()
                DataGridView1.Rows(va - 1).Cells(2).Value = 1
                DataGridView1.Rows(va - 1).Cells(8).Value = va
                DataGridView1.Rows(va - 1).Cells(3).Value = 0
                DataGridView1.Rows(va - 1).Cells(5).Value = 0

                DataGridView1.Rows(va - 1).Cells(7).Value = 0
                DataGridView1.Rows(va - 1).Cells(4).Value = "MTS" & "/PRD"
                abrir()
                Dim sql102213 As String = "EXEC Sp_ListadoUltimoPrecios '01',NULL,'" + TextBox14.Text.ToString().Trim() + "','" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "'"
                Dim cmd102213 As New SqlCommand(sql102213, conx)
                Rsr3 = cmd102213.ExecuteReader()
                If Rsr3.Read() = True Then
                    DataGridView1.Rows(va - 1).Cells(6).Value = Trim(Rsr3(0).ToString())
                End If
                Rsr3.Close()
            End If
        End If
    End Sub
    Sub cargarSErvicios()
        DataGridView6.Rows.Add(7)
        DataGridView6.Rows(0).Cells(0).Value = "Corte"
        DataGridView6.Rows(1).Cells(0).Value = "Costura"
        DataGridView6.Rows(2).Cells(0).Value = "Acabados"
        DataGridView6.Rows(3).Cells(0).Value = "Perchado"
        DataGridView6.Rows(4).Cells(0).Value = "Rama"
        DataGridView6.Rows(5).Cells(0).Value = "Teñido"
        DataGridView6.Rows(6).Cells(0).Value = "Flete Transporte"

        DataGridView6.Rows(0).Cells(1).Value = 1
        DataGridView6.Rows(1).Cells(1).Value = 1
        DataGridView6.Rows(2).Cells(1).Value = 1
        DataGridView6.Rows(3).Cells(1).Value = 1
        DataGridView6.Rows(4).Cells(1).Value = 1
        DataGridView6.Rows(5).Cells(1).Value = 1
        DataGridView6.Rows(6).Cells(1).Value = 1

        DataGridView6.Rows(0).Cells(3).Value = 1
        DataGridView6.Rows(1).Cells(3).Value = 1
        DataGridView6.Rows(2).Cells(3).Value = 1
        DataGridView6.Rows(3).Cells(3).Value = 1
        DataGridView6.Rows(4).Cells(3).Value = 1
        DataGridView6.Rows(5).Cells(3).Value = 1
        DataGridView6.Rows(6).Cells(3).Value = 1

        DataGridView6.Rows(0).Cells(4).Value = 0
        DataGridView6.Rows(1).Cells(4).Value = 0
        DataGridView6.Rows(2).Cells(4).Value = 0
        DataGridView6.Rows(3).Cells(4).Value = 0
        DataGridView6.Rows(4).Cells(4).Value = 0
        DataGridView6.Rows(5).Cells(4).Value = 0
        DataGridView6.Rows(6).Cells(4).Value = 0

        DataGridView6.Rows(0).Cells(5).Value = 0
        DataGridView6.Rows(1).Cells(5).Value = 0
        DataGridView6.Rows(2).Cells(5).Value = 0
        DataGridView6.Rows(3).Cells(5).Value = 0
        DataGridView6.Rows(4).Cells(5).Value = 0
        DataGridView6.Rows(5).Cells(5).Value = 0
        DataGridView6.Rows(6).Cells(5).Value = 0

        DataGridView6.Rows(0).Cells(2).Value = "UND/PRD"
        DataGridView6.Rows(1).Cells(2).Value = "UND/PRD"
        DataGridView6.Rows(2).Cells(2).Value = "UND/PRD"
        DataGridView6.Rows(3).Cells(2).Value = "UND/PRD"
        DataGridView6.Rows(4).Cells(2).Value = "UND/PRD"
        DataGridView6.Rows(5).Cells(2).Value = "UND/PRD"
        DataGridView6.Rows(6).Cells(2).Value = "UND/PRD"
    End Sub
    Sub cargarInformacionAviosAcabados()
        DataGridView3.Rows.Add(12)

        DataGridView3.Rows(0).Cells(0).Value = "ETIMARNC00000000625"
        DataGridView3.Rows(0).Cells(1).Value = "ETIQUETA TEJIDA MARCA TALLA*RASS VINTAGE*F/BEIGE L/ COLORES*SMALL"
        DataGridView3.Rows(0).Cells(2).Value = "1"
        DataGridView3.Rows(0).Cells(3).Value = "1"
        DataGridView3.Rows(0).Cells(4).Value = "UND"
        DataGridView3.Rows(0).Cells(5).Value = Convert.ToDouble(DataGridView3.Rows(0).Cells(2).Value) * Convert.ToDouble(DataGridView3.Rows(0).Cells(3).Value)
        DataGridView3.Rows(0).Cells(6).Value = "0"
        DataGridView3.Rows(0).Cells(7).Value = "0"
        DataGridView3.Rows(0).Cells(8).Value = "1"
        '------------------------------------------------
        DataGridView3.Rows(1).Cells(0).Value = "ETIMARNC00000002254"
        DataGridView3.Rows(1).Cells(1).Value = "ETIQUETA BADANA BLUES - X DEFINIR - 2022"
        DataGridView3.Rows(1).Cells(2).Value = "1"
        DataGridView3.Rows(1).Cells(3).Value = "1"
        DataGridView3.Rows(1).Cells(4).Value = "UND"
        DataGridView3.Rows(1).Cells(5).Value = Convert.ToDouble(DataGridView3.Rows(0).Cells(2).Value) * Convert.ToDouble(DataGridView3.Rows(0).Cells(3).Value)
        DataGridView3.Rows(1).Cells(6).Value = "0"
        DataGridView3.Rows(1).Cells(7).Value = "0"
        DataGridView3.Rows(1).Cells(8).Value = "1"
        '------------------------------------------------
        DataGridView3.Rows(2).Cells(0).Value = "BOTMETNC00000000726"
        DataGridView3.Rows(2).Cells(1).Value = "BOTON*X  DEFINIR  POR  EL  CLIENTE"
        DataGridView3.Rows(2).Cells(2).Value = "1"
        DataGridView3.Rows(2).Cells(3).Value = "1"
        DataGridView3.Rows(2).Cells(4).Value = "UND"
        DataGridView3.Rows(2).Cells(5).Value = Convert.ToDouble(DataGridView3.Rows(0).Cells(2).Value) * Convert.ToDouble(DataGridView3.Rows(0).Cells(3).Value)
        DataGridView3.Rows(2).Cells(6).Value = "0"
        DataGridView3.Rows(2).Cells(7).Value = "0"
        DataGridView3.Rows(2).Cells(8).Value = "1"
        '-----------------------------------------------
        DataGridView3.Rows(3).Cells(0).Value = "REMMETNC00000000355"
        DataGridView3.Rows(3).Cells(1).Value = "REMACHE*X DEFINIR"
        DataGridView3.Rows(3).Cells(2).Value = "1"
        DataGridView3.Rows(3).Cells(3).Value = "1"
        DataGridView3.Rows(3).Cells(4).Value = "UND"
        DataGridView3.Rows(3).Cells(5).Value = Convert.ToDouble(DataGridView3.Rows(0).Cells(2).Value) * Convert.ToDouble(DataGridView3.Rows(0).Cells(3).Value)
        DataGridView3.Rows(3).Cells(6).Value = "0"
        DataGridView3.Rows(3).Cells(7).Value = "0"
        DataGridView3.Rows(3).Cells(8).Value = "1"
        '-----------------------------------------------
        DataGridView3.Rows(4).Cells(0).Value = "HATCARNC00000001789"
        DataGridView3.Rows(4).Cells(1).Value = "HANG TAG  POR  DEFINIR"
        DataGridView3.Rows(4).Cells(2).Value = "1"
        DataGridView3.Rows(4).Cells(3).Value = "1"
        DataGridView3.Rows(4).Cells(4).Value = "UND"
        DataGridView3.Rows(4).Cells(5).Value = Convert.ToDouble(DataGridView3.Rows(0).Cells(2).Value) * Convert.ToDouble(DataGridView3.Rows(0).Cells(3).Value)
        DataGridView3.Rows(4).Cells(6).Value = "0"
        DataGridView3.Rows(4).Cells(7).Value = "0"
        DataGridView3.Rows(4).Cells(8).Value = "1"
        '-----------------------------------------------
        DataGridView3.Rows(5).Cells(0).Value = "HATCARNC00000000951"
        DataGridView3.Rows(5).Cells(1).Value = "HANG TAG HECHO  EN PERU*F/NEGRO L/BLANCO*OE"
        DataGridView3.Rows(5).Cells(2).Value = "1"
        DataGridView3.Rows(5).Cells(3).Value = "1"
        DataGridView3.Rows(5).Cells(4).Value = "UND"
        DataGridView3.Rows(5).Cells(5).Value = Convert.ToDouble(DataGridView3.Rows(0).Cells(2).Value) * Convert.ToDouble(DataGridView3.Rows(0).Cells(3).Value)
        DataGridView3.Rows(5).Cells(6).Value = "0"
        DataGridView3.Rows(5).Cells(7).Value = "0"
        DataGridView3.Rows(5).Cells(8).Value = "1"
        '-----------------------------------------------
        DataGridView3.Rows(6).Cells(0).Value = "STKTALNC00000000013"
        DataGridView3.Rows(6).Cells(1).Value = "STICKER TALLERO*X DEFINIR "
        DataGridView3.Rows(6).Cells(2).Value = "1"
        DataGridView3.Rows(6).Cells(3).Value = "1"
        DataGridView3.Rows(6).Cells(4).Value = "UND"
        DataGridView3.Rows(6).Cells(5).Value = Convert.ToDouble(DataGridView3.Rows(0).Cells(2).Value) * Convert.ToDouble(DataGridView3.Rows(0).Cells(3).Value)
        DataGridView3.Rows(6).Cells(6).Value = "0"
        DataGridView3.Rows(6).Cells(7).Value = "0"
        DataGridView3.Rows(6).Cells(8).Value = "1"
        '-----------------------------------------------
        DataGridView3.Rows(7).Cells(0).Value = "STICODNC00000000652"
        DataGridView3.Rows(7).Cells(1).Value = "STICKER DE PRECIO  X DEFINIR"
        DataGridView3.Rows(7).Cells(2).Value = "1"
        DataGridView3.Rows(7).Cells(3).Value = "1"
        DataGridView3.Rows(7).Cells(4).Value = "UND"
        DataGridView3.Rows(7).Cells(5).Value = Convert.ToDouble(DataGridView3.Rows(0).Cells(2).Value) * Convert.ToDouble(DataGridView3.Rows(0).Cells(3).Value)
        DataGridView3.Rows(7).Cells(6).Value = "0"
        DataGridView3.Rows(7).Cells(7).Value = "0"
        DataGridView3.Rows(7).Cells(8).Value = "1"
        '-----------------------------------------------
        DataGridView3.Rows(8).Cells(0).Value = "SUEBOLNC00000000046"
        DataGridView3.Rows(8).Cells(1).Value = "BOLSA*POR DEFINIR"
        DataGridView3.Rows(8).Cells(2).Value = "1"
        DataGridView3.Rows(8).Cells(3).Value = "1"
        DataGridView3.Rows(8).Cells(4).Value = "UND"
        DataGridView3.Rows(8).Cells(5).Value = Convert.ToDouble(DataGridView3.Rows(0).Cells(2).Value) * Convert.ToDouble(DataGridView3.Rows(0).Cells(3).Value)
        DataGridView3.Rows(8).Cells(6).Value = "0"
        DataGridView3.Rows(8).Cells(7).Value = "0"
        DataGridView3.Rows(8).Cells(8).Value = "1"
        '-----------------------------------------------
        DataGridView3.Rows(9).Cells(0).Value = "SUECAJNC00000000043"
        DataGridView3.Rows(9).Cells(1).Value = "CAJA*X DEFINIR"
        DataGridView3.Rows(9).Cells(2).Value = "1"
        DataGridView3.Rows(9).Cells(3).Value = "1"
        DataGridView3.Rows(9).Cells(4).Value = "UND"
        DataGridView3.Rows(9).Cells(5).Value = Convert.ToDouble(DataGridView3.Rows(0).Cells(2).Value) * Convert.ToDouble(DataGridView3.Rows(0).Cells(3).Value)
        DataGridView3.Rows(9).Cells(6).Value = "0"
        DataGridView3.Rows(9).Cells(7).Value = "0"
        DataGridView3.Rows(9).Cells(8).Value = "1"
        '-----------------------------------------------
        DataGridView3.Rows(10).Cells(0).Value = "APLCINNC00000000015"
        DataGridView3.Rows(10).Cells(1).Value = "CINTA   POR  DEFINIR                                                                                                                                  "
        DataGridView3.Rows(10).Cells(2).Value = "1"
        DataGridView3.Rows(10).Cells(3).Value = "1"
        DataGridView3.Rows(10).Cells(4).Value = "UND"
        DataGridView3.Rows(10).Cells(5).Value = Convert.ToDouble(DataGridView3.Rows(0).Cells(2).Value) * Convert.ToDouble(DataGridView3.Rows(0).Cells(3).Value)
        DataGridView3.Rows(10).Cells(6).Value = "0"
        DataGridView3.Rows(10).Cells(7).Value = "0"
        DataGridView3.Rows(10).Cells(8).Value = "1"
        '-----------------------------------------------
        DataGridView3.Rows(11).Cells(0).Value = "ACMIMPNC00000000002"
        DataGridView3.Rows(11).Cells(1).Value = "IMPERDIBLE DORADO*LAR-1.9CM ANC 5MM                                                                                                                   "
        DataGridView3.Rows(11).Cells(2).Value = "1"
        DataGridView3.Rows(11).Cells(3).Value = "1"
        DataGridView3.Rows(11).Cells(4).Value = "UND"
        DataGridView3.Rows(11).Cells(5).Value = Convert.ToDouble(DataGridView3.Rows(0).Cells(2).Value) * Convert.ToDouble(DataGridView3.Rows(0).Cells(3).Value)
        DataGridView3.Rows(11).Cells(6).Value = "0"
        DataGridView3.Rows(11).Cells(7).Value = "0"
        DataGridView3.Rows(11).Cells(8).Value = "1"
    End Sub
    Private Sub TextBox10_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox10.KeyPress
        e.Handled = Not IsNumeric(e.KeyChar) And Not Char.IsControl(e.KeyChar)
    End Sub

    Private Sub TextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            DataGridView3.Rows.Clear()
            DataGridView6.Rows.Clear()
            DataGridView1.Rows.Clear()
            Dim I, i10, i11 As Integer
            Dim ml As New FCOTIZACION
            Dim ml1 As New VCOTIZACION
            habilitar_registros()


            I = Len(TextBox10.Text)
            Select Case I
                Case "1" : TextBox10.Text = "0000000" & "" & TextBox10.Text
                Case "2" : TextBox10.Text = "000000" & "" & TextBox10.Text
                Case "3" : TextBox10.Text = "00000" & "" & TextBox10.Text
                Case "4" : TextBox10.Text = "0000" & "" & TextBox10.Text
                Case "5" : TextBox10.Text = "000" & "" & TextBox10.Text
                Case "6" : TextBox10.Text = "00" & "" & TextBox10.Text
                Case "7" : TextBox10.Text = "0" & "" & TextBox10.Text
                Case "8" : TextBox10.Text = TextBox10.Text
            End Select
            ml1.gid_cotizacion = TextBox10.Text
            If ml.mostrar_COTIZACION(ml1) > 0 Then

                Dim mv As New ftela
                Dim mv1 As New vTela

                mv1.gid_cotizacion = TextBox10.Text
                dt8 = mv.mostrar_tela(mv1)
                If dt8.Rows.Count <> 0 Then
                    Dim sum As Double
                    DataGridView8.DataSource = dt8
                    i10 = DataGridView8.Rows.Count
                    DataGridView1.Rows.Add(i10)
                    For i2 = 0 To i10 - 1
                        DataGridView1.Rows(i2).Cells(0).Value = DataGridView8.Rows(i2).Cells(1).Value
                        DataGridView1.Rows(i2).Cells(1).Value = DataGridView8.Rows(i2).Cells(2).Value
                        DataGridView1.Rows(i2).Cells(2).Value = DataGridView8.Rows(i2).Cells(3).Value
                        DataGridView1.Rows(i2).Cells(3).Value = DataGridView8.Rows(i2).Cells(4).Value
                        DataGridView1.Rows(i2).Cells(4).Value = DataGridView8.Rows(i2).Cells(5).Value
                        DataGridView1.Rows(i2).Cells(5).Value = DataGridView8.Rows(i2).Cells(6).Value
                        DataGridView1.Rows(i2).Cells(6).Value = DataGridView8.Rows(i2).Cells(7).Value
                        DataGridView1.Rows(i2).Cells(7).Value = DataGridView8.Rows(i2).Cells(8).Value
                        DataGridView1.Rows(i2).Cells(8).Value = DataGridView8.Rows(i2).Cells(9).Value
                        sum = sum + Convert.ToDouble(DataGridView8.Rows(i2).Cells(8).Value)
                    Next
                    TextBox7.Text = sum.ToString("N2")
                End If

                Dim mc As New favios_costura
                Dim mc1 As New vAvios_Costura

                mc1.gid_cotizacion = TextBox10.Text
                dt9 = mc.mostrar_avios_costura(mc1)
                If dt9.Rows.Count <> 0 Then
                    Dim sum1 As Double
                    DataGridView9.DataSource = dt9
                    i11 = DataGridView9.Rows.Count
                    DataGridView2.Rows.Add(i11)
                    For i3 = 0 To i11 - 1
                        DataGridView2.Rows(i3).Cells(0).Value = DataGridView9.Rows(i3).Cells(1).Value
                        DataGridView2.Rows(i3).Cells(1).Value = DataGridView9.Rows(i3).Cells(2).Value
                        DataGridView2.Rows(i3).Cells(2).Value = DataGridView9.Rows(i3).Cells(3).Value
                        DataGridView2.Rows(i3).Cells(3).Value = DataGridView9.Rows(i3).Cells(4).Value
                        DataGridView2.Rows(i3).Cells(4).Value = DataGridView9.Rows(i3).Cells(5).Value
                        DataGridView2.Rows(i3).Cells(5).Value = DataGridView9.Rows(i3).Cells(6).Value
                        DataGridView2.Rows(i3).Cells(6).Value = DataGridView9.Rows(i3).Cells(7).Value
                        DataGridView2.Rows(i3).Cells(7).Value = DataGridView9.Rows(i3).Cells(8).Value
                        DataGridView2.Rows(i3).Cells(8).Value = DataGridView9.Rows(i3).Cells(9).Value
                        sum1 = sum1 + Convert.ToDouble(DataGridView9.Rows(i3).Cells(8).Value)
                    Next
                    TextBox19.Text = sum1.ToString("N2")
                End If

                Dim mk As New favios_acabados
                Dim mk1 As New vAvios_Acabados

                mk1.gid_cotizacion = TextBox10.Text
                dt10 = mk.mostrar_avios_acabados(mk1)
                If dt10.Rows.Count <> 0 Then
                    Dim sum2 As Double
                    DataGridView10.DataSource = dt10
                    Dim i12 As Integer
                    i12 = DataGridView10.Rows.Count
                    DataGridView3.Rows.Add(i12)
                    For i4 = 0 To i12 - 1
                        DataGridView3.Rows(i4).Cells(0).Value = DataGridView10.Rows(i4).Cells(1).Value
                        DataGridView3.Rows(i4).Cells(1).Value = DataGridView10.Rows(i4).Cells(2).Value
                        DataGridView3.Rows(i4).Cells(2).Value = DataGridView10.Rows(i4).Cells(3).Value
                        DataGridView3.Rows(i4).Cells(3).Value = DataGridView10.Rows(i4).Cells(4).Value
                        DataGridView3.Rows(i4).Cells(4).Value = DataGridView10.Rows(i4).Cells(5).Value
                        DataGridView3.Rows(i4).Cells(5).Value = DataGridView10.Rows(i4).Cells(6).Value
                        DataGridView3.Rows(i4).Cells(6).Value = DataGridView10.Rows(i4).Cells(7).Value
                        DataGridView3.Rows(i4).Cells(7).Value = DataGridView10.Rows(i4).Cells(8).Value
                        DataGridView3.Rows(i4).Cells(8).Value = DataGridView10.Rows(i4).Cells(9).Value
                        sum2 = sum2 + Convert.ToDouble(DataGridView10.Rows(i4).Cells(8).Value)
                    Next
                    TextBox17.Text = sum2.ToString("N2")
                End If


                Dim mh As New flavados
                Dim mh1 As New vLavados

                mh1.gid_cotizacion = TextBox10.Text
                dt11 = mh.mostrar_lavados(mh1)
                If dt11.Rows.Count <> 0 Then
                    Dim sum3 As Double
                    DataGridView11.DataSource = dt11
                    Dim i22 As Integer
                    i22 = DataGridView11.Rows.Count
                    DataGridView4.Rows.Add(i22)
                    For i44 = 0 To i22 - 1
                        DataGridView4.Rows(i44).Cells(0).Value = DataGridView11.Rows(i44).Cells(1).Value
                        DataGridView4.Rows(i44).Cells(1).Value = DataGridView11.Rows(i44).Cells(2).Value
                        DataGridView4.Rows(i44).Cells(2).Value = DataGridView11.Rows(i44).Cells(3).Value
                        DataGridView4.Rows(i44).Cells(3).Value = DataGridView11.Rows(i44).Cells(4).Value
                        DataGridView4.Rows(i44).Cells(4).Value = DataGridView11.Rows(i44).Cells(5).Value
                        DataGridView4.Rows(i44).Cells(5).Value = DataGridView11.Rows(i44).Cells(6).Value
                        DataGridView4.Rows(i44).Cells(6).Value = DataGridView11.Rows(i44).Cells(7).Value
                        DataGridView4.Rows(i44).Cells(7).Value = DataGridView11.Rows(i44).Cells(8).Value
                        sum3 = sum3 + Convert.ToDouble(DataGridView11.Rows(i44).Cells(7).Value)
                    Next
                    TextBox20.Text = sum3.ToString("N2")
                End If

                Dim mf As New faplicaciones
                Dim mf1 As New Aplicaciones

                mf1.gid_cotizacion = TextBox10.Text
                dt12 = mf.mostrar_aplicaciones(mf1)
                If dt12.Rows.Count <> 0 Then
                    Dim sum4 As Double
                    DataGridView12.DataSource = dt12
                    Dim i13 As Integer
                    i13 = DataGridView12.Rows.Count
                    DataGridView5.Rows.Add(i13)
                    For i5 = 0 To i13 - 1
                        DataGridView5.Rows(i5).Cells(0).Value = DataGridView12.Rows(i5).Cells(1).Value
                        DataGridView5.Rows(i5).Cells(1).Value = DataGridView12.Rows(i5).Cells(2).Value
                        DataGridView5.Rows(i5).Cells(2).Value = DataGridView12.Rows(i5).Cells(3).Value
                        DataGridView5.Rows(i5).Cells(3).Value = DataGridView12.Rows(i5).Cells(4).Value
                        DataGridView5.Rows(i5).Cells(4).Value = DataGridView12.Rows(i5).Cells(5).Value
                        DataGridView5.Rows(i5).Cells(5).Value = DataGridView12.Rows(i5).Cells(6).Value
                        DataGridView5.Rows(i5).Cells(6).Value = DataGridView12.Rows(i5).Cells(7).Value
                        DataGridView5.Rows(i5).Cells(7).Value = DataGridView12.Rows(i5).Cells(8).Value
                        sum4 = sum4 + Convert.ToDouble(DataGridView12.Rows(i5).Cells(7).Value)
                    Next
                    TextBox18.Text = sum4.ToString("N2")
                End If
                Dim my As New fmano_obra
                Dim my1 As New vManoObra
                my1.gid_cotizacion = TextBox10.Text
                dt13 = my.mostrar_manoobra(my1)
                If dt13.Rows.Count <> 0 Then
                    Dim sum5 As Double
                    DataGridView13.DataSource = dt13
                    Dim i14 As Integer
                    i14 = DataGridView13.Rows.Count
                    DataGridView6.Rows.Add(i14)
                    For i6 = 0 To i14 - 1
                        DataGridView6.Rows(i6).Cells(0).Value = DataGridView13.Rows(i6).Cells(1).Value
                        DataGridView6.Rows(i6).Cells(1).Value = DataGridView13.Rows(i6).Cells(3).Value
                        DataGridView6.Rows(i6).Cells(2).Value = DataGridView13.Rows(i6).Cells(4).Value
                        DataGridView6.Rows(i6).Cells(3).Value = DataGridView13.Rows(i6).Cells(5).Value
                        DataGridView6.Rows(i6).Cells(4).Value = DataGridView13.Rows(i6).Cells(6).Value
                        DataGridView6.Rows(i6).Cells(5).Value = DataGridView13.Rows(i6).Cells(7).Value
                        sum5 = sum5 + Convert.ToDouble(DataGridView13.Rows(i6).Cells(7).Value)
                    Next
                    TextBox21.Text = sum5.ToString("N2")
                End If
                If Label11.Text = 0 Then
                    Dim mr As New FCOTIZACION
                    Dim mr1 As New VCOTIZACION
                    mr1.gid_cotizacion = TextBox10.Text
                    dt14 = mr.mostrar_COTIZACION2(mr1)
                    If dt14.Rows.Count <> 0 Then
                        cotizacion.DataSource = dt14
                        If cotizacion.Rows.Count <> 0 Then
                            TextBox2.Text = cotizacion.Rows(0).Cells(1).Value
                            TextBox3.Text = cotizacion.Rows(0).Cells(2).Value
                            TextBox4.Text = cotizacion.Rows(0).Cells(3).Value
                            TextBox5.Text = cotizacion.Rows(0).Cells(4).Value
                            'TextBox8.Text = cotizacion.Rows(0).Cells(5).Value
                            If cotizacion.Rows(0).Cells(6).Value = 1 Then
                                TextBox11.Text = "SOLES"
                            Else
                                TextBox11.Text = "DOLARES"
                            End If
                            'TextBox9.Text = cotizacion.Rows(0).Cells(8).Value
                            'TextBox11.Text = cotizacion.Rows(0).Cells(10).Value
                            'TextBox6.Text = cotizacion.Rows(0).Cells(11).Value
                            TextBox12.Text = cotizacion.Rows(0).Cells(12).Value
                            TextBox7.Text = cotizacion.Rows(0).Cells(20).Value
                            TextBox19.Text = cotizacion.Rows(0).Cells(21).Value
                            TextBox17.Text = cotizacion.Rows(0).Cells(22).Value
                            TextBox20.Text = cotizacion.Rows(0).Cells(23).Value
                            TextBox18.Text = cotizacion.Rows(0).Cells(24).Value
                            TextBox21.Text = cotizacion.Rows(0).Cells(25).Value
                            'TextBox10.Text = cotizacion.Rows(0).Cells(28).Value
                            dt14.Clear()
                        End If
                    End If
                End If

                BLOQUEAR_registros()
                Button12.Enabled = True
                Button13.Enabled = True
                TextBox7.Select()
                dt8.Clear()
                dt9.Clear()
                dt10.Clear()
                dt11.Clear()
                dt12.Clear()
                dt13.Clear()

            Else
                If TextBox22.Text = 0 Then
                    cargarInformacionAviosAcabados()
                    cargartela()
                    cargarSErvicios()
                End If
                Button11.Enabled = True
                TextBox7.Select()
            End If
        End If
    End Sub
    Sub ENVIAR_CORREO_UDP()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim fk As New fnopedido
        Dim jj As New vnapedido
        Dim corre As String
        Dim contra, cabecera, DETALLE As String
        Dim area As String
        Select Case Label12.Text
            Case "1" : area = "Udp"
            Case "2" : area = "Logistica"
            Case "3" : area = "Produccion"
        End Select

        contra = "20Via$&it2"
        corre = "admin@viannysac.com"
        If TextBox22.Text = 0 Then
            cabecera = "SE CREO LA HOJA DE COTIZACION"

        Else
            cabecera = "SE EDITO LA HOJA DE COTIZACION"

        End If


        Dim Mailz As String = "" &
      "<!DOCTYPE html>
<html xmlns='http://www.w3.org/1999/xhtml'>
    <head>
        <title>" + cabecera + "</title>
    </head>
    <body>
        <font color='green'>" + cabecera + "</font><br/><br/>
        <table border='1' cellspacing='0' cellpadding='5'>
            <thead>
                <tr style='background-color: #e0f7fa;'>
                    <th align='center' ><font color='black'>HOJA DE COTIZACION N°</font></th>
                    <th align='center'><font color='black'>OD</font></th>
                    <th align='center'><font color='black'>VERSION</font></th>
                    <th align='center' ><font color='black'>PM</font></th>
                    <th align='center' ><font color='black'>ESTILO</font></th>
                    <th align='center' ><font color='black'>PRODUCTO</font></th>
                    <th align='center' ><font color='black'>CLIENTE</font></th> 
                    <th align='center' ><font color='black'>AREA</font></th> 
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td align='center'> " + TextBox10.Text.ToString().Trim() + "</td>
                    <td align='center'>" + TextBox1.Text.ToString().Trim() + " </td>
                    <td align='center'>" + TextBox6.Text.ToString().Trim() + "</td>
                    <td align='center'>" + TextBox9.Text.ToString().Trim() + " </td>
                    <td align='center'>" + TextBox4.Text.ToString().Trim() + " </td>
                    <td align='center'>" + TextBox3.Text.ToString().Trim() + " </td>
                    <td align='center'>" + TextBox2.Text.ToString().Trim() + " </td> 
                    <td align='center'>" + area + " </td>   
                </tr>
            </tbody>
        
      </table>
     <br/><br/>
     <font color='green'> INFORMACION DE LA TELA</font><br/><br/>
        <table border='1' cellspacing='0' cellpadding='5'>
            <thead>
                <tr style='background-color: #e0f7fa;'>
                    <th align='center'><font color='black'>NUEVO CAMPO 1</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 2</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 3</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 4</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 5</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 6</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 7</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 8</font></th>
                </tr>
            </thead>
            <tbody>"


        ' Revisar la cantidad de filas en el DataGridView
        For Each row As DataGridViewRow In DataGridView1.Rows
            If Not row.IsNewRow Then
                ' Capturar la información de cada fila y agregarla al cuerpo HTML
                Mailz &=
                    "<tr>" &
                 "<td align='center'>" + row.Cells(0).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(1).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(2).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(3).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(4).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(5).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(6).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(7).Value.ToString() + "</td>" &
                 "</tr>"
            End If
        Next
        ' Finalizar el cuerpo del HTML
        Mailz &= " </tbody></table>"
        Mailz &= "<br/><br/>
     <font color='green'> INFORMACION DE LOS AVIOS DE COSTURA</font><br/><br/>
        <table border='1' cellspacing='0' cellpadding='5'>
            <thead>
                <tr style='background-color: #e0f7fa;'>
                    <th align='center'><font color='black'>NUEVO CAMPO 1</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 2</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 3</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 4</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 5</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 6</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 7</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 8</font></th>
                </tr>
            </thead>
            <tbody>"
        ' Revisar la cantidad de filas en el DataGridView
        For Each row As DataGridViewRow In DataGridView2.Rows
            If Not row.IsNewRow Then
                ' Capturar la información de cada fila y agregarla al cuerpo HTML
                Mailz &=
                    "<tr>" &
                 "<td align='center'>" + row.Cells(0).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(1).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(2).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(3).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(4).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(5).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(6).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(7).Value.ToString() + "</td>" &
                 "</tr>"
            End If
        Next
        Mailz &= " </tbody></table>"
        Mailz &= "<br/><br/>
     <font color='green'> INFORMACION DE LOS AVIOS DE ACABADOS</font><br/><br/>
        <table border='1' cellspacing='0' cellpadding='5'>
            <thead>
                <tr style='background-color: #e0f7fa;'>
                    <th align='center'><font color='black'>NUEVO CAMPO 1</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 2</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 3</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 4</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 5</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 6</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 7</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 8</font></th>
                </tr>
            </thead>
            <tbody>"
        ' Revisar la cantidad de filas en el DataGridView
        For Each row As DataGridViewRow In DataGridView3.Rows
            If Not row.IsNewRow Then
                ' Capturar la información de cada fila y agregarla al cuerpo HTML
                Mailz &=
                    "<tr>" &
                 "<td align='center'>" + row.Cells(0).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(1).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(2).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(3).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(4).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(5).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(6).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(7).Value.ToString() + "</td>" &
                 "</tr>"
            End If
        Next
        Mailz &= " </tbody></table>"
        Mailz &= "<br/><br/>
     <font color='green'> INFORMACION DE LAVADOS</font><br/><br/>
        <table border='1' cellspacing='0' cellpadding='5'>
            <thead>
                <tr style='background-color: #e0f7fa;'>
                    <th align='center'><font color='black'>NUEVO CAMPO 1</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 2</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 3</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 4</font></th>                    
                </tr>
            </thead>
            <tbody>"
        ' Revisar la cantidad de filas en el DataGridView
        For Each row As DataGridViewRow In DataGridView4.Rows
            If Not row.IsNewRow Then
                ' Capturar la información de cada fila y agregarla al cuerpo HTML
                Mailz &=
                    "<tr>" &
                 "<td align='center'>" + row.Cells(0).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(1).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(2).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(3).Value.ToString() + "</td>" &
                 "</tr>"
            End If
        Next
        Mailz &= " </tbody></table>"
        Mailz &= "<br/><br/>
     <font color='green'> INFORMACION DE APLICACIONES</font><br/><br/>
        <table border='1' cellspacing='0' cellpadding='5'>
            <thead>
                <tr style='background-color: #e0f7fa;'>
                    <th align='center'><font color='black'>NUEVO CAMPO 1</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 2</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 3</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 4</font></th>                    
                </tr>
            </thead>
            <tbody>"
        ' Revisar la cantidad de filas en el DataGridView
        For Each row As DataGridViewRow In DataGridView5.Rows
            If Not row.IsNewRow Then
                ' Capturar la información de cada fila y agregarla al cuerpo HTML
                Mailz &=
                    "<tr>" &
                 "<td align='center'>" + row.Cells(0).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(1).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(2).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(3).Value.ToString() + "</td>" &
                 "</tr>"
            End If
        Next
        Mailz &= " </tbody></table>"
        Mailz &= "<br/><br/>
     <font color='green'> INFORMACION DE LOS SERVICIOS</font><br/><br/>
        <table border='1' cellspacing='0' cellpadding='5'>
            <thead>
                <tr style='background-color: #e0f7fa;'>
                    <th align='center'><font color='black'>NUEVO CAMPO 1</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 2</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 3</font></th>
                    <th align='center'><font color='black'>NUEVO CAMPO 4</font></th>                    
                </tr>
            </thead>
            <tbody>"
        ' Revisar la cantidad de filas en el DataGridView
        For Each row As DataGridViewRow In DataGridView6.Rows
            If Not row.IsNewRow Then
                ' Capturar la información de cada fila y agregarla al cuerpo HTML
                Mailz &=
                    "<tr>" &
                 "<td align='center'>" + row.Cells(0).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(1).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(2).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(3).Value.ToString() + "</td>" &
                 "</tr>"
            End If
        Next
        Mailz &= " </tbody></table></body></html>"
        With message
            .From = New System.Net.Mail.MailAddress(corre)
            .To.Add("coordinadorit@viannysac.com")
            .IsBodyHtml = True

            .Body = Mailz

            If Label12.Text = "1" Then
                .Subject = "El area de Udp Registro la informacion de  la Cotizacion N°" & TextBox10.Text & " -- OD" & TextBox1.Text
            Else
                If Label12.Text = "2" Then
                    .Subject = "El area de Compras Registro la informacion de  la Cotizacion N°" & TextBox10.Text & " --  OD" & TextBox1.Text
                Else
                    .Subject = "El area de Produccion Registro la informacion de  la Cotizacion N°" & TextBox10.Text & " --  OD" & TextBox1.Text
                End If

            End If

            .Priority = System.Net.Mail.MailPriority.Normal
        End With

        With smtp
            .EnableSsl = True
            .Port = "587"
            .Host = "smtppro.zoho.com"
            .Credentials = New Net.NetworkCredential(corre, contra)

            .Send(message)
        End With

        Try
            MessageBox.Show("EL mensaje de correo ha sido enviado.", "Correo enviado",
                             MessageBoxButtons.OK)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error al enviar correo",
                             MessageBoxButtons.OK)
        End Try
    End Sub
End Class