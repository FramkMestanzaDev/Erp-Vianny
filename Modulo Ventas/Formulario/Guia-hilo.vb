Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System.Net.Mail
Public Class Guia_hilo
    Public enunciado As SqlCommand
    Public respuesta As SqlDataReader
    Public enunciado1 As SqlCommand
    Public respuesta1 As SqlDataReader
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Public enunciado3 As SqlCommand
    Public respuesta3 As SqlDataReader
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Public comando As SqlCommand
    Dim ty, ty2, ty3, ty5 As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Sub llenar_combo_box(ByVal cb As ComboBox)
        Try

            conn = New SqlDataAdapter("select rtrim(ltrim(motiv_17f)) AS motiv_17f,rtrim(ltrim(nomb_17f)) AS nomb_17f from custom_vianny.dbo.Fag1700  where ccia_17F ='" + Trim(Label30.Text) + "'", conx)
            conn.Fill(ty)
            ComboBox1.DataSource = ty
            ComboBox1.DisplayMember = "nomb_17f"
            ComboBox1.ValueMember = "motiv_17f"
            'respuesta = enunciado.ExecuteReader
            'While respuesta.Read
            '    cb.Items.Add(respuesta.Item("nomb_17f"))
            'End While
            'respuesta.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub llenar_combo_box2(ByVal cb As ComboBox, ByVal alm As String)
        Try

            conn = New SqlDataAdapter("select dest_21s,rtrim(ltrim(nomb_21s)) as nomb_21s from custom_vianny.dbo.cag210s where  Empr_21s =" + Trim(Label30.Text) + " and almac_21s=" + alm, conx)
            conn.Fill(ty2)
            ComboBox2.DataSource = ty2
            ComboBox2.DisplayMember = "nomb_21s"
            ComboBox2.ValueMember = "dest_21s"
            'respuesta = enunciado.ExecuteReader
            'While respuesta.Read
            '    cb.Items.Add(respuesta.Item("nomb_17f"))
            'End While
            'respuesta.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Dim Rsr30012 As SqlDataReader
    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        ToolTip1.SetToolTip(PictureBox5, "Anular Informacion")
        Dim respuesta As DialogResult
        Dim ml As New fguiasistema
        Dim mk As New vguiacabecera
        Dim ll As Integer
        Dim sql102213 As String = "SELECT cierre_3 FROM custom_vianny.dbo.Cie0300 A  Where a.ccia ='" + Label30.Text + "' AND A.talm_3 ='" + Label23.Text + "' AND A.fcie_3='" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "'"
        Dim cmd102213 As New SqlCommand(sql102213, conx)
        Rsr30012 = cmd102213.ExecuteReader()
        Dim jo As Integer
        If Rsr30012.Read() Then
            jo = Rsr30012(0)
        Else
            jo = 0
        End If
        Rsr30012.Close()
        respuesta = MessageBox.Show("QUIERES ANULAR LA GUIA REMISION?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            If jo = 0 Then
                ll = DataGridView1.Rows.Count
                For i = 0 To ll - 1
                    mk.gsfactu = TextBox1.Text
                    mk.gnfactu = TextBox2.Text
                    mk.gano = Year(DateTimePicker1.Value)
                    mk.galmacen = Label23.Text
                    mk.gcodigo = DataGridView1.Rows(i).Cells(2).Value
                    mk.gpartida = DataGridView1.Rows(i).Cells(6).Value
                    mk.gcantidad = DataGridView1.Rows(i).Cells(7).Value
                    mk.gccia = Label30.Text
                    ml.anular_guia(mk)
                Next
            Else
                MsgBox("DIA " & DateTimePicker1.Value & " CERRADO,NO SE PUEDE EDITAR LA INFORMACION - CONSULTE CON EL AREA CONTABLE ")
            End If

        End If



        MsgBox("SE ANULO LA GUIA SOLICITADA")
        Me.Close()

    End Sub
    Public Sub sieditable()


        TextBox8.Enabled = True
        TextBox10.Enabled = True
        TextBox11.Enabled = True
        TextBox12.Enabled = True
        TextBox13.Enabled = True
        TextBox14.Enabled = True
        TextBox15.Enabled = True
        TextBox16.Enabled = True
        TextBox4.Enabled = True
        DateTimePicker1.Enabled = True
        DateTimePicker2.Enabled = True
        PictureBox1.Enabled = True
        Button5.Enabled = True
        PictureBox4.Enabled = True
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        ComboBox3.Enabled = True
    End Sub
    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        ToolTip1.SetToolTip(PictureBox4, "Agregar Informacion")
        Dim camposFaltantes As New List(Of String) ' Lista para almacenar los nombres de los campos faltantes
        If String.IsNullOrEmpty(TextBox8.Text) Then
            camposFaltantes.Add(" Cliente ")
        End If


        If camposFaltantes.Count > 0 Then
            ' Mostrar mensaje de error indicando los campos faltantes
            Dim camposFaltantesStr As String = String.Join(", ", camposFaltantes)
            MessageBox.Show("Falta ingresar información en los siguientes campos: " & camposFaltantesStr, "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim productos As New Productos_Hilo
            productos.Label3.Text = Label23.Text
            productos.Label4.Text = 1
            productos.Label5.Text = Label29.Text
            productos.Label6.Text = Label30.Text
            Select Case ComboBox3.Text
                Case "13" : productos.Label7.Text = "1"
                Case "05" : productos.Label7.Text = "1"
                Case "14" : productos.Label7.Text = "1"
                Case "11" : productos.Label7.Text = "1"

                Case Else
                    productos.Label7.Text = "2"
            End Select
            productos.Show()
        End If
    End Sub
    Private Sub PictureBox1_MouseHover(sender As Object, e As EventArgs) Handles PictureBox1.MouseHover
        ToolTip1.SetToolTip(PictureBox1, "IMPRIMIR INFORMACION")
        ToolTip1.ToolTipTitle = "GUIA REMISION"
        ToolTip1.ToolTipIcon = ToolTipIcon.Info
    End Sub
    Private Sub PictureBox2_MouseHover(sender As Object, e As EventArgs) Handles PictureBox2.MouseHover
        ToolTip2.SetToolTip(PictureBox2, "EDITAR INFORMACION")
        ToolTip2.ToolTipTitle = "GUIA REMISION"
        ToolTip2.ToolTipIcon = ToolTipIcon.Info
    End Sub
    Private Sub PictureBox3_MouseHover(sender As Object, e As EventArgs)
        ToolTip3.SetToolTip(Button5, "GUARDAR INFORMACION")
        ToolTip3.ToolTipTitle = "GUIA REMISION"
        ToolTip3.ToolTipIcon = ToolTipIcon.Info
    End Sub
    Private Sub PictureBox4_MouseHover(sender As Object, e As EventArgs) Handles PictureBox4.MouseHover
        ToolTip4.SetToolTip(PictureBox4, "AGREGAR INFORMACION")
        ToolTip4.ToolTipTitle = "GUIA REMISION"
        ToolTip4.ToolTipIcon = ToolTipIcon.Info
    End Sub
    Private Sub PictureBox5_MouseHover(sender As Object, e As EventArgs) Handles PictureBox5.MouseHover
        ToolTip5.SetToolTip(PictureBox5, "ANULAR INFORMACION")
        ToolTip5.ToolTipTitle = "GUIA REMISION"
        ToolTip5.ToolTipIcon = ToolTipIcon.Info
    End Sub
    Dim Rsr3001 As SqlDataReader
    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        ToolTip1.SetToolTip(PictureBox2, "Editar Informacion")
        Dim sql102213 As String = "SELECT cierre_3 FROM custom_vianny.dbo.Cie0300 A  Where a.ccia ='" + Label30.Text + "' AND A.talm_3 ='" + Label23.Text + "' AND A.fcie_3='" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "'"
        Dim cmd102213 As New SqlCommand(sql102213, conx)
        Rsr3001 = cmd102213.ExecuteReader()
        Dim jo As Integer
        If Rsr3001.Read() Then
            jo = Rsr3001(0)
        Else
            jo = 0
        End If
        Rsr3001.Close()
        Dim respuesta As DialogResult
        respuesta = MessageBox.Show("QUIERES EDITAR ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            If jo = 0 Then
                sieditable()
                DataGridView1.Enabled = True
                Button1.Enabled = True
            Else
                MsgBox("DIA " & DateTimePicker1.Value & " CERRADO,NO SE PUEDE EDITAR LA INFORMACION - CONSULTE CON EL AREA CONTABLE ")
            End If
        End If
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Reporte_Guia.TextBox1.Text = TextBox1.Text
        Reporte_Guia.TextBox2.Text = TextBox2.Text
        Reporte_Guia.TextBox3.Text = Label30.Text
        Reporte_Guia.TextBox4.Text = Label37.Text
        Reporte_Guia.Show()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            TextBox10.Text = "Av. Las Torres Mz. F Lote 8 - Huachipa -Chosica - Lima"
        Else
            If CheckBox2.Checked = False Then
                TextBox10.Text = "JR. LOS OLMOS NRO. 358 URB. CANTO BELLO (ALT. PARDERO 1 AV. "
            End If
        End If
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub
    Sub llenar_combo_box12()
        Try

            conn = New SqlDataAdapter("SELECT cele,dele FROM custom_vianny.dbo.TAB0100 A  Where A.ccia='" + Label30.Text + "' AND A.CTAB='AREAMC'", conx)
            conn.Fill(ty5)
            ComboBox4.DataSource = ty5
            ComboBox4.DisplayMember = "dele"
            ComboBox4.ValueMember = "cele"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Guia_hilo_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Dim func As New fguiasistema
        Dim fu As New vguiacabecera
        Dim fu1 As New vguiacabecera
        'If Label30.Text = "01" Then
        '    TextBox1.ReadOnly = True
        'Else
        '    TextBox1.ReadOnly = False
        'End If

        If Label30.Text = "02" Then
            TextBox3.Text = "JR. LOS OLMOS NRO. 358 (PARADEROA AV. 1 CANTO GRANDE)"
            TextBox11.Text = "20459785834"
            TextBox1.Text = "0010"
            TextBox12.Text = "GRAUS INDUSTRIAS TEXTIL S.A."
            TextBox13.Text = "JR. LOS OLMOS NRO. 358 (PARADEROA AV. 1 CANTO GRANDE)"
        End If
        fu.gserie = TextBox1.Text
        fu.gccia = Label30.Text
        TextBox2.Text = func.mostrar_guia_correlativo(fu) + 1
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

        fu1.gmes = DateTime.Now.ToString("MM")
        fu1.galmacen = Label23.Text
        fu1.gano = Label29.Text
        fu1.gccia = Label30.Text
        Label3.Text = func.mostrar_correlativo_nsa(fu1) + 1
        Select Case Label3.Text.Length

            Case "1" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "000" & "" & Label3.Text
            Case "2" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "00" & "" & Label3.Text
            Case "3" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "0" & "" & Label3.Text
            Case "4" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "" & Label3.Text
        End Select
        abrir()
        llenar_combo_box(ComboBox1)
        llenar_combo_box2(ComboBox2, Label23.Text)
        llenar_combo_box12()
        TextBox2.Select()

        RadioButton1.Checked = True
        RadioButton3.Checked = True
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox5.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        TextBox10.Enabled = False
        TextBox11.Enabled = False
        TextBox13.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        PictureBox1.Enabled = False
        PictureBox2.Enabled = False
        Button5.Enabled = False
        PictureBox4.Enabled = False
        PictureBox5.Enabled = False
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ComboBox3.SelectedIndex = 1
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.F1 Then
            Dim trans As New Transportistas
            trans.TextBox1.Text = 2
            trans.ShowDialog()

        Else
            If e.KeyCode = Keys.Enter Then
                abrir()
                enunciado3 = New SqlCommand("select nomb_18 , direcc_18,ruc_18,chofer_18,placa_18 + marca_18 AS PLA,brevete_18 from custom_vianny.dbo.Fag1800 where empr_18 =" + Trim(Label30.Text) + " and trans_18 =" + TextBox4.Text, conx)
                respuesta3 = enunciado3.ExecuteReader
                While respuesta3.Read
                    TextBox5.Text = respuesta3.Item("nomb_18")
                    TextBox12.Text = respuesta3.Item("nomb_18")
                    TextBox13.Text = respuesta3.Item("direcc_18")
                    TextBox11.Text = respuesta3.Item("ruc_18")
                    TextBox16.Text = respuesta3.Item("chofer_18")
                    TextBox14.Text = respuesta3.Item("PLA")
                    TextBox15.Text = respuesta3.Item("brevete_18")

                End While
                respuesta3.Close()
            End If
        End If
    End Sub



    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.F1 Then
            If Trim(TextBox4.Text).Length = 0 Then
                MessageBox.Show("FALTA QUE AGREGAR EL TRANSPORTE", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                Dim cli As New Clientes
                cli.TextBox3.Text = "300"
                cli.Show()
            End If

        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
    'Sub ENVIAR_CORREO()
    '    Dim message As New MailMessage
    '    Dim smtp As New SmtpClient
    '    Dim fk As New fnopedido
    '    Dim jj As New vnapedido
    '    'Dim corre As String
    '    'jj.gvendedor = TextBox8.Text
    '    'corre = fk.buscar_correo(jj)
    '    With message
    '        .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
    '        .To.Add("hrivera@viannysac.com,asistentecobranzas@viannysac.com")
    '        .Body = "Nota Genero la Guia N°" & TextBox1.Text & "-" & TextBox2.Text
    '        .Subject = "Cliente" & TextBox9.Text & "Almacen" & Label23.Text
    '        .Priority = System.Net.Mail.MailPriority.Normal
    '    End With

    '    With smtp
    '        .EnableSsl = True
    '        .Port = "587"
    '        .Host = "smtp.gmail.com"
    '        .Credentials = New Net.NetworkCredential("admin@viannysac.com", "$i$tema$$i$tem@$")
    '        .Send(message)
    '    End With

    '    Try
    '        MessageBox.Show("Su mensaje de correo ha sido enviado.", "Correo enviado",
    '                         MessageBoxButtons.OK)
    '    Catch ex As Exception
    '        MessageBox.Show("Error: " & ex.Message, "Error al enviar correo",
    '                         MessageBoxButtons.OK)
    '    End Try
    'End Sub
    Dim dt1, dt2, hg As DataTable

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim respuesta As DialogResult
        Dim I18, VAL, TAB As Integer
        TAB = DataGridView1.Rows.Count
        If TAB <> 0 Then
            respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then

                If DataGridView1.SelectedRows.Count > 0 Then
                    Dim filaSeleccionada As DataGridViewRow = DataGridView1.SelectedRows(0)
                    Dim valorCelda As String = filaSeleccionada.Cells(21).Value.ToString()
                    'MsgBox(valorCelda)

                    abrir()
                    Dim agregar1 As String = "delete  Almacen_Ficticio  WHERE  ItmAfi ='" + valorCelda + "' and AlmAfi ='" + Trim(Label23.Text) + "' AND EmpAfi = '" + Trim(Label30.Text) + "' AND GuiAfi ='" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + "' AND PerAfi='" + Trim(Label29.Text) + "'"
                    ELIMINAR(agregar1)
                End If
                DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
                I18 = DataGridView1.Rows.Count
                For i1 = 0 To I18 - 1

                    VAL = i1 + 1
                    Select Case VAL.ToString.Length
                        Case "1" : DataGridView1.Rows(i1).Cells(0).Value = "00" & "" & VAL
                        Case "2" : DataGridView1.Rows(i1).Cells(0).Value = "0" & "" & VAL
                        Case "3" : DataGridView1.Rows(i1).Cells(0).Value = VAL

                    End Select
                Next

            End If
        Else
            MsgBox("NO HAY PRODUCTOS A ELIMINAR")
        End If

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If Label30.Text = "01" Then
            If ComboBox3.Text = "MISMA EMPRESA" Then
                TextBox8.Text = "20508740361"
                TextBox9.Text = "CONSORCIO TEXTIL VIANNY"
                TextBox10.Text = "JR. LOS OLMOS 358 CANTO BELLO SJL"
            End If
        Else
            If ComboBox3.Text = "MISMA EMPRESA" Then
                TextBox8.Text = "20459785834"
                TextBox9.Text = "GRAUS INDUSTRIAS TEXTIL S.A."
                TextBox10.Text = "JR. LOS OLMOS 358 CANTO BELLO SJL"
            End If
        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ToolTip1.SetToolTip(Button5, "Guardar Informacion")
        Dim valo, valo1 As String
        Dim sum, vg, us As Integer
        sum = 0
        us = 0
        Dim i1, a1, HJ As Integer


        i1 = DataGridView1.Rows.Count
        a1 = 0
        For a1 = 0 To i1 - 1
            valo = DataGridView1.Rows(a1).Cells(4).Value
            valo1 = DataGridView1.Rows(a1).Cells(10).Value
            If valo = "" Then
                sum = sum + 1

            End If
            vg = vg + DataGridView1.Rows(a1).Cells(13).Value

            If Convert.ToDecimal(DataGridView1.Rows(a1).Cells(7).Value) > Convert.ToDecimal(DataGridView1.Rows(a1).Cells(11).Value) Or Convert.ToDecimal(DataGridView1.Rows(a1).Cells(7).Value) < 0 Then
                DataGridView1.Rows(a1).DefaultCellStyle.BackColor = Color.Red
                DataGridView1.Rows(a1).DefaultCellStyle.ForeColor = Color.White
                us = us + 1
            Else
                DataGridView1.Rows(a1).DefaultCellStyle.BackColor = Color.White
                DataGridView1.Rows(a1).DefaultCellStyle.ForeColor = Color.Black
            End If

        Next
        Dim cant, Cop, Cup, Ccon, Cpartida, Cum As Integer

        cant = DataGridView1.Rows.Count
        Cop = 0
        Cup = 0
        Ccon = 0
        Cpartida = 0
        Cum = 0
        For i1 = 0 To cant - 1
            If Trim(DataGridView1.Rows(i1).Cells(1).Value).Length = 0 Then
                Cop = Cop + 1
            End If
            If Trim(DataGridView1.Rows(i1).Cells(4).Value).Length = 0 Then
                Cup = Cup + 1
            End If
            If Trim(DataGridView1.Rows(i1).Cells(7).Value).Length = 0 Then
                Ccon = Ccon + 1
            End If
            If Trim(DataGridView1.Rows(i1).Cells(6).Value).Length = 0 Then
                Cpartida = Cpartida + 1
            End If
            If Trim(DataGridView1.Rows(i1).Cells(5).Value).Length = 0 Then
                Cum = Cum + 1
            End If
        Next
        Dim VacioCliente As Boolean
        If String.IsNullOrEmpty(TextBox8.Text) Then
            VacioCliente = True
        End If
        If Cop > 0 And Trim(ComboBox1.Text) <> "VENTA" Then
            MessageBox.Show("FALTA INGRESAR LA OP EN UNA CELDA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)


        Else
            If Cup > 0 Then
                MessageBox.Show("FALTA INGRESAR LA CANTIDAD DE PRESENTACION EN ALGUNA CELDA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)

            Else
                If Ccon > 0 Then
                    MessageBox.Show("FALTA INGRESAR LA CANTIDAD EN KG/UND EN ALGUNA CELDA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)

                Else
                    If Cpartida > 0 And (Trim(Label23.Text) = "03" Or Trim(Label23.Text) = "59") Then
                        MessageBox.Show("FALTA INGRESAR EL LOTE  EN ALGUNA CELDA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)

                    Else
                        If Cum > 0 Then
                            MessageBox.Show("FALTA INGRESAR LA UNIDAD DE PRESENTACION EN ALGUNA CELDA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)

                        Else
                            If VacioCliente = True Then
                                MessageBox.Show("FALTA INGRESAR EL CLIENTE", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Else
                                If us > 0 Then
                                    MsgBox("EXISTE PRODUCTOS QUE ESTAN EN NEGATIVO O TIENEN EXCEDENTES")
                                Else

                                    '--------------------NOTA SALIDA
                                    Dim ns1 As New vnsacabece
                                    Dim ns2 As New fnsa


                                    ns1.gncom = Trim(Label3.Text)
                                    ns1.gfecha = DateTimePicker1.Value
                                    ns1.gglosa = Trim(TextBox6.Text)
                                    ns1.gdoc = "009"
                                    ns1.gguia = TextBox1.Text & "-" & TextBox2.Text
                                    ns1.gsfactu = Trim(TextBox7.Text)
                                    ns1.gnfactu = Trim(TextBox17.Text)
                                    ns1.gruc = Trim(TextBox8.Text)
                                    ns1.galmacen = Trim(Label23.Text)
                                    If ComboBox3.Text = "INTERNA" Then
                                        ns1.gtipointexte = "INT"
                                    Else
                                        ns1.gtipointexte = "EXT"
                                    End If
                                    ns1.gorige_sali = ComboBox2.SelectedValue.ToString
                                    ns1.gtdocento = "009"
                                    ns1.gusuario = TextBox19.Text
                                    ns1.gano = Label29.Text
                                    ns1.gfase = Trim(TextBox24.Text)
                                    ns1.gccia = Label30.Text
                                    ns1.gpedorig_3 = ""
                                    If RadioButton3.Checked = True Then
                                        ns1.gadevol = "0"
                                    Else
                                        If RadioButton4.Checked = True Then
                                            ns1.gadevol = "1"
                                        End If

                                    End If
                                    '--------------------------------------- GUIA REMISION
                                    Dim gr1 As New vguiacabecera
                                    Dim gr2 As New fguiasistema
                                    Dim gr7 As New vguiacabecera
                                    gr7.gsfactu = TextBox1.Text
                                    gr7.gnfactu = TextBox2.Text
                                    gr7.gncom = Label3.Text
                                    gr7.galmacen = Label23.Text
                                    gr7.gano = Label29.Text
                                    gr7.gccia = Label30.Text
                                    gr2.eliminar_guia(gr7)


                                    '-----eliminar nsa
                                    Dim hn As New vnia
                                    Dim fg As New fnia
                                    Dim con As String
                                    con = Label3.Text
                                    hn.gcomprobante = con
                                    hn.galmacen = Label23.Text
                                    hn.gncom = 2
                                    hn.gano = Label29.Text
                                    hn.gccia = Label30.Text
                                    fg.eliminar_nia(hn)

                                    '-----------------------------------------
                                    gr1.gsfactu = TextBox1.Text
                                    gr1.gnfactu = TextBox2.Text
                                    gr1.gruc = TextBox8.Text
                                    gr1.gruc3 = TextBox11.Text
                                    gr1.gtip_documento = "001"
                                    gr1.gruc_transpo = TextBox12.Text
                                    gr1.gdirec_transport = Trim(TextBox13.Text.ToString)
                                    gr1.grsocial = Trim(TextBox9.Text)
                                    gr1.gfcom = DateTimePicker1.Value
                                    gr1.gfcom1 = DateTimePicker2.Value
                                    gr1.gdireccion = Trim(TextBox10.Text.ToString)
                                    gr1.gplaca = TextBox14.Text
                                    gr1.gdireccion_partida = Trim(TextBox3.Text.ToString)
                                    gr1.gNOTASALIDA = DateTime.Now.ToString("yyyy") & Label23.Text & "2" & Label3.Text
                                    gr1.gchofer = TextBox16.Text
                                    gr1.gserie = TextBox7.Text
                                    gr1.gcorrelativo = TextBox17.Text
                                    gr1.galmacen = Label23.Text
                                    gr1.glicencia = TextBox15.Text
                                    gr1.gusuario = "almace"
                                    gr1.gfase = ""
                                    gr1.gdestino = ComboBox2.SelectedValue.ToString
                                    gr1.gtrpo = TextBox4.Text
                                    gr1.gccia = Label30.Text
                                    gr1.gglosa = Trim(TextBox6.Text)
                                    gr1.gordens_3 = TextBox21.Text
                                    gr1.gsubfase_3 = TextBox24.Text

                                    If RadioButton3.Checked = True Then
                                        gr1.gsituacion = "1"
                                    Else
                                        If RadioButton4.Checked = True Then
                                            gr1.gsituacion = "2"
                                        End If

                                    End If


                                    gr1.gmotivo = ComboBox1.SelectedValue.ToString




                                    'Dim i5, num2, I17 As Integer
                                    '
                                    If ns2.insertar_cabecera_nsa(ns1) = True And gr2.insertar_cabecera_guia_sistema(gr1) = True Then
                                        MessageBox.Show("SE REGISTRO LA GUIA REMISION CORRECTAMENTE ", "Guardando registros", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                        'For i5 = 1 To i1 - 1
                                        '    If DataGridView3.Rows(i5).Cells(4).Value < DataGridView3.Rows(i5 - 1).Cells(4).Value Then
                                        '        num2 = DataGridView3.Rows(i5 - 1).Cells(4).Value + 1
                                        '        I17 = i5

                                        '    End If
                                        'Next

                                        Dim i10, a As Integer
                                        i10 = DataGridView1.Rows.Count


                                        For a = 0 To i10 - 1



                                            Dim ns3 As New vnsadetalle
                                            Dim ns4 As New fnsa
                                            Dim nu2 As String


                                            nu2 = DataGridView1.Rows(a).Cells(0).Value
                                            Select Case nu2.Length
                                                Case "1" : nu2 = "00" & "" & nu2
                                                Case "2" : nu2 = "0" & "" & nu2
                                                Case "3" : nu2 = nu2
                                            End Select

                                            ns3.gncom = Label3.Text
                                            ns3.gitem = nu2
                                            ns3.glinea = Trim(DataGridView1.Rows(a).Cells(2).Value)
                                            ns3.gcantidad = DataGridView1.Rows(a).Cells(7).Value
                                            ns3.gund = Trim(DataGridView1.Rows(a).Cells(8).Value)
                                            If Trim(DataGridView1.Rows(a).Cells(1).Value) = "" Then
                                                ns3.gop = ""
                                            Else
                                                ns3.gop = Trim(DataGridView1.Rows(a).Cells(1).Value)
                                            End If
                                            ns3.gunidad_medidad = DataGridView1.Rows(a).Cells(5).Value
                                            If Trim(Convert.ToString(DataGridView1.Rows(a).Cells(10).Value)) = "" Then
                                                ns3.gpartida = ""
                                            Else
                                                ns3.gpartida = Trim(DataGridView1.Rows(a).Cells(10).Value)
                                            End If
                                            ns3.gcantenvio = DataGridView1.Rows(a).Cells(14).Value
                                            ns3.grollo1 = DataGridView1.Rows(a).Cells(4).Value
                                            ns3.galmacen = Label23.Text
                                            If Trim(Convert.ToString(DataGridView1.Rows(a).Cells(9).Value)) = "" Then
                                                ns3.gordtejeduria = ""
                                            Else
                                                ns3.gordtejeduria = Trim(DataGridView1.Rows(a).Cells(9).Value)
                                            End If

                                            ns3.gano = Label29.Text
                                            If Trim(Convert.ToString(DataGridView1.Rows(a).Cells(6).Value)) = "" Then
                                                ns3.glote = ""
                                            Else
                                                ns3.glote = Trim(DataGridView1.Rows(a).Cells(6).Value)
                                            End If
                                            ns3.gclasif = ""
                                            ns3.gccia = Label30.Text
                                            ns3.gOcompra = ""
                                            ns4.insertar_detalle_nsa(ns3)

                                            '-------------------------------- etalle guia
                                            Dim gr3 As New vguiadetalle
                                            Dim gr4 As New fguiasistema
                                            Dim nu1 As String

                                            gr3.gsfactu = TextBox1.Text
                                            gr3.gnfactu = TextBox2.Text
                                            nu1 = DataGridView1.Rows(a).Cells(0).Value
                                            Select Case nu1.Length
                                                Case "1" : nu1 = "00" & "" & nu1
                                                Case "2" : nu1 = "0" & "" & nu1
                                                Case "3" : nu1 = nu1
                                            End Select
                                            gr3.gitems = nu1
                                            gr3.glinea = Trim(DataGridView1.Rows(a).Cells(2).Value)
                                            gr3.gcantidad = DataGridView1.Rows(a).Cells(7).Value
                                            gr3.gcanencio = DataGridView1.Rows(a).Cells(14).Value
                                            If Trim(Convert.ToString(DataGridView1.Rows(a).Cells(1).Value)) = "" Then
                                                gr3.gordens = ""
                                            Else
                                                gr3.gordens = Trim(DataGridView1.Rows(a).Cells(1).Value)
                                            End If

                                            gr3.grollo = DataGridView1.Rows(a).Cells(4).Value
                                            If Trim(Convert.ToString(DataGridView1.Rows(a).Cells(10).Value)) = "" Then
                                                gr3.gparti = ""
                                            Else
                                                gr3.gparti = Trim(DataGridView1.Rows(a).Cells(10).Value)
                                            End If

                                            gr3.gunidad_medida = Trim(DataGridView1.Rows(a).Cells(5).Value)
                                            gr3.gunidad_medida2 = Trim(DataGridView1.Rows(a).Cells(8).Value)
                                            If Trim(Convert.ToString(DataGridView1.Rows(a).Cells(9).Value)) = "" Then
                                                gr3.gordtejeduria = ""
                                            Else
                                                gr3.gordtejeduria = Trim(DataGridView1.Rows(a).Cells(9).Value)
                                            End If
                                            If Trim(Convert.ToString(DataGridView1.Rows(a).Cells(6).Value)) = "" Then
                                                gr3.glote = ""
                                            Else
                                                gr3.glote = Trim(DataGridView1.Rows(a).Cells(6).Value)
                                            End If
                                            gr3.gccia = Label30.Text

                                            If Trim(Convert.ToString(DataGridView1.Rows(a).Cells(16).Value)) = "" Then
                                                gr3.gobser1_3a = ""
                                            Else
                                                gr3.gobser1_3a = Trim(DataGridView1.Rows(a).Cells(16).Value)
                                            End If

                                            If Trim(Convert.ToString(DataGridView1.Rows(a).Cells(17).Value)) = "" Then
                                                gr3.gobser2_3a = ""
                                            Else
                                                gr3.gobser2_3a = Trim(DataGridView1.Rows(a).Cells(17).Value)
                                            End If

                                            If Trim(Convert.ToString(DataGridView1.Rows(a).Cells(18).Value)) = "" Then
                                                gr3.gobser3_3a = ""
                                            Else
                                                gr3.gobser3_3a = Trim(DataGridView1.Rows(a).Cells(18).Value)
                                            End If

                                            If Trim(Convert.ToString(DataGridView1.Rows(a).Cells(19).Value)) = "" Then
                                                gr3.gobser4_3a = ""
                                            Else
                                                gr3.gobser4_3a = Trim(DataGridView1.Rows(a).Cells(19).Value)
                                            End If

                                            If Trim(Convert.ToString(DataGridView1.Rows(a).Cells(20).Value)) = "" Then
                                                gr3.gobser5_3a = ""
                                            Else
                                                gr3.gobser5_3a = Trim(DataGridView1.Rows(a).Cells(20).Value)
                                            End If
                                            If Trim(Label23.Text) = "59" Then
                                                gr3.gclasif = "1"
                                            Else
                                                gr3.gclasif = ""
                                            End If

                                            gr4.insertar_detalle_guia_sistema(gr3)


                                        Next

                                        'If Trim(ComboBox1.Text) = "VENTA" Then
                                        '    ENVIAR_CORREO()
                                        'End If
                                        Dim respuesta As DialogResult

                                        respuesta = MessageBox.Show("DESEA IMPRIMIR LA NOTA DE SALIDA?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                                        If (respuesta = Windows.Forms.DialogResult.Yes) Then

                                            Reporte_Guia.TextBox1.Text = TextBox1.Text
                                            Reporte_Guia.TextBox2.Text = TextBox2.Text
                                            Reporte_Guia.TextBox3.Text = Label30.Text
                                            Reporte_Guia.TextBox4.Text = Label37.Text
                                            Reporte_Guia.Show()
                                        End If
                                        abrir()
                                        Dim agregar1 As String = "delete  Almacen_Ficticio  WHERE  AlmAfi ='" + Trim(Label23.Text) + "' AND EmpAfi = '" + Trim(Label30.Text) + "' AND GuiAfi ='" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + "' AND PerAfi='" + Trim(Label29.Text) + "'"
                                        ELIMINAR(agregar1)

                                        Dim func As New fguiasistema
                                        Dim func1 As New vguiacabecera

                                        func1.gserie = TextBox1.Text
                                        func1.galmacen = Label23.Text
                                        func1.gccia = Label30.Text
                                        TextBox2.Text = func.mostrar_guia_correlativo(func1) + 1

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
                                        func1.gmes = DateTime.Now.ToString("MM")
                                        func1.galmacen = Label23.Text
                                        func1.gano = Label29.Text
                                        func1.gccia = Label30.Text
                                        Label3.Text = func.mostrar_correlativo_nsa(func1) + 1
                                        Select Case Label3.Text.Length

                                            Case "1" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "000" & "" & Label3.Text
                                            Case "2" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "00" & "" & Label3.Text
                                            Case "3" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "0" & "" & Label3.Text
                                            Case "4" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "" & Label3.Text
                                        End Select

                                        Label39.Text = "0"
                                        RadioButton1.Checked = True
                                        Button1.Enabled = False
                                        TextBox4.Text = ""
                                        TextBox5.Text = ""
                                        TextBox8.Text = ""
                                        TextBox9.Text = ""
                                        TextBox10.Text = ""
                                        TextBox11.Text = ""
                                        TextBox12.Text = ""
                                        TextBox13.Text = ""
                                        TextBox14.Text = ""
                                        TextBox15.Text = ""
                                        TextBox16.Text = ""
                                        TextBox6.Text = ""
                                        DataGridView1.Rows.Clear()
                                        RadioButton1.Checked = True
                                        TextBox3.Enabled = False
                                        TextBox4.Enabled = False
                                        TextBox5.Enabled = False
                                        TextBox5.Enabled = False
                                        TextBox8.Enabled = False
                                        TextBox9.Enabled = False
                                        TextBox10.Enabled = False
                                        TextBox11.Enabled = False
                                        TextBox13.Enabled = False
                                        ComboBox1.Enabled = False
                                        ComboBox2.Enabled = False
                                        ComboBox3.Enabled = False
                                        PictureBox1.Enabled = False
                                        PictureBox2.Enabled = False
                                        Button5.Enabled = False
                                        PictureBox4.Enabled = False
                                        PictureBox5.Enabled = False
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub TextBox20_TextChanged(sender As Object, e As EventArgs) Handles TextBox20.TextChanged

    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
    Dim Rsr215, Rsr211 As SqlDataReader
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        ty2.Clear()
        ty.Clear()
        abrir()
        llenar_combo_box(ComboBox1)
        llenar_combo_box2(ComboBox2, Label23.Text)

        Dim i, i2 As Integer
        Dim ml, m As New vguiacabecera
        Dim ml1 As New vguiacabecera
        Dim mk As New fguiasistema
        i = Len(TextBox2.Text)
        i2 = Len(TextBox1.Text)
        TextBox8.Enabled = True
        TextBox9.Enabled = True
        TextBox20.Enabled = True
        If e.KeyCode = Keys.Enter Then
            DataGridView1.Rows.Clear()
            Select Case i
                Case "1" : TextBox2.Text = "0000000" & "" & TextBox2.Text
                Case "2" : TextBox2.Text = "000000" & "" & TextBox2.Text
                Case "3" : TextBox2.Text = "00000" & "" & TextBox2.Text
                Case "4" : TextBox2.Text = "0000" & "" & TextBox2.Text
                Case "5" : TextBox2.Text = "000" & "" & TextBox2.Text
                Case "6" : TextBox2.Text = "00" & "" & TextBox2.Text
                Case "7" : TextBox2.Text = "0" & "" & TextBox2.Text
                Case "8" : TextBox2.Text = TextBox2.Text
            End Select
            Select Case i2
                Case "1" : TextBox1.Text = "000" & "" & TextBox1.Text
                Case "2" : TextBox1.Text = "00" & "" & TextBox1.Text
                Case "3" : TextBox1.Text = "0" & "" & TextBox1.Text
                Case "4" : TextBox1.Text = TextBox1.Text

            End Select
            m.gcorrelativo = Trim(TextBox2.Text)
            m.gserie = Trim(TextBox1.Text)
            m.gccia = Trim(Label30.Text)

            If mk.buscar_ano_guia(m) = Trim(Label29.Text) Or (mk.buscar_ano_guia(m) = False) Then
                If Label35.Text = "0" Then
                    'TextBox22.Text = 0
                    'TextBox21.Text = 0

                    ml.gcorrelativo = TextBox2.Text
                    ml.gserie = TextBox1.Text
                    ml.galmacen = Label23.Text
                    ml.gccia = Label30.Text
                    ml1.gcorrelativo = TextBox2.Text
                    ml1.gserie = TextBox1.Text
                    ml1.gccia = Label30.Text
                    If mk.mostrar_correlativo_guia1(ml) = True Then
                        If mk.mostrar_almacen_guia(ml) = Label23.Text Then
                            Dim jk As New fguiasistema

                            dt1 = jk.mostrar_cabecera_guia(ml)
                            dt2 = jk.mostrar_detalle_guia_hilo(ml1)

                            If dt1.Rows.Count <> 0 Then
                                DataGridView2.DataSource = dt1
                                TextBox8.Text = DataGridView2.Rows(0).Cells(0).Value
                                TextBox9.Text = DataGridView2.Rows(0).Cells(1).Value
                                TextBox10.Text = DataGridView2.Rows(0).Cells(2).Value
                                TextBox16.Text = DataGridView2.Rows(0).Cells(3).Value
                                TextBox14.Text = DataGridView2.Rows(0).Cells(4).Value
                                TextBox15.Text = DataGridView2.Rows(0).Cells(5).Value
                                TextBox12.Text = DataGridView2.Rows(0).Cells(9).Value
                                TextBox13.Text = DataGridView2.Rows(0).Cells(10).Value
                                TextBox11.Text = DataGridView2.Rows(0).Cells(13).Value
                                TextBox4.Text = DataGridView2.Rows(0).Cells(14).Value
                                TextBox18.Text = "FAC"
                                TextBox7.Text = DataGridView2.Rows(0).Cells(6).Value
                                TextBox17.Text = DataGridView2.Rows(0).Cells(7).Value
                                TextBox19.Text = DataGridView2.Rows(0).Cells(19).Value
                                TextBox6.Text = Trim(DataGridView2.Rows(0).Cells(18).Value)
                                DateTimePicker1.Value = DataGridView2.Rows(0).Cells(20).Value
                                DateTimePicker2.Value = DataGridView2.Rows(0).Cells(21).Value
                                TextBox24.Text = Trim(DataGridView2.Rows(0).Cells(22).Value)
                                If Trim(DataGridView2.Rows(0).Cells(23).Value) = "" Then
                                    TextBox23.Text = ""
                                Else
                                    TextBox23.Text = Trim(DataGridView2.Rows(0).Cells(23).Value)
                                End If
                                If Trim(DataGridView2.Rows(0).Cells(24).Value) = "1" Then
                                    RadioButton3.Checked = True
                                    RadioButton4.Checked = False
                                Else
                                    RadioButton3.Checked = False
                                    RadioButton4.Checked = True
                                End If
                                Label3.Text = DataGridView2.Rows(0).Cells(11).Value & "" & DataGridView2.Rows(0).Cells(12).Value
                                abrir()
                                enunciado = New SqlCommand("select nomb_18 as EMPRESA from custom_vianny.dbo.Fag1800 where empr_18 ='" + Trim(Label30.Text) + "' and trans_18 ='" + Trim(TextBox4.Text) + "'", conx)
                                respuesta = enunciado.ExecuteReader
                                While respuesta.Read
                                    TextBox5.Text = respuesta.Item("EMPRESA")
                                End While
                                respuesta.Close()
                                enunciado1 = New SqlCommand("select rtrim(ltrim(nomb_17f)) AS nomb_17f from custom_vianny.dbo.Fag1700 where ccia_17F ='" + Trim(Label30.Text) + "' and motiv_17f = '" + Trim(DataGridView2.Rows(0).Cells(15).Value) + "'", conx)
                                respuesta1 = enunciado1.ExecuteReader
                                While respuesta1.Read
                                    ComboBox1.Text = respuesta1.Item("nomb_17f")
                                End While
                                respuesta1.Close()
                                enunciado2 = New SqlCommand("select rtrim(ltrim(nomb_21s)) as nomb_21s from custom_vianny.dbo.cag210s where Empr_21s='" + Trim(Label30.Text) + "' and almac_21s='" + Trim(Label23.Text) + "' and dest_21s='" + Trim(DataGridView2.Rows(0).Cells(16).Value) + "'", conx)
                                respuesta2 = enunciado2.ExecuteReader
                                While respuesta2.Read
                                    ComboBox2.Text = respuesta2.Item("nomb_21s")
                                End While
                                respuesta2.Close()
                                If DataGridView2.Rows(0).Cells(15).Value = "EXT" Then
                                    ComboBox3.Text = "EXTERNO"
                                Else
                                    If DataGridView2.Rows(0).Cells(15).Value = "INT" Then
                                        ComboBox3.Text = "MISMA EMPRESA"
                                    End If
                                End If
                                If DataGridView2.Rows(0).Cells(8).Value = 0 Then
                                    RadioButton2.Checked = True
                                    RadioButton1.Checked = False
                                    RadioButton1.Enabled = False
                                    Label28.Visible = True
                                Else
                                    RadioButton2.Checked = False
                                    RadioButton1.Checked = True
                                    RadioButton1.Enabled = False
                                    Label28.Visible = False
                                End If


                            End If
                            If dt2.Rows.Count <> 0 Then
                                DataGridView3.DataSource = dt2
                                Dim j As Integer
                                j = DataGridView3.Rows.Count
                                DataGridView1.Rows.Add(j)
                                For p = 0 To j - 1
                                    DataGridView1.Rows(p).Cells(0).Value = DataGridView3.Rows(p).Cells(4).Value
                                    DataGridView1.Rows(p).Cells(1).Value = DataGridView3.Rows(p).Cells(6).Value
                                    DataGridView1.Rows(p).Cells(2).Value = DataGridView3.Rows(p).Cells(5).Value
                                    DataGridView1.Rows(p).Cells(3).Value = DataGridView3.Rows(p).Cells(3).Value
                                    DataGridView1.Rows(p).Cells(4).Value = DataGridView3.Rows(p).Cells(0).Value
                                    DataGridView1.Rows(p).Cells(5).Value = DataGridView3.Rows(p).Cells(10).Value
                                    DataGridView1.Rows(p).Cells(6).Value = DataGridView3.Rows(p).Cells(7).Value
                                    DataGridView1.Rows(p).Cells(7).Value = DataGridView3.Rows(p).Cells(1).Value
                                    DataGridView1.Rows(p).Cells(8).Value = DataGridView3.Rows(p).Cells(2).Value
                                    DataGridView1.Rows(p).Cells(10).Value = DataGridView3.Rows(p).Cells(8).Value
                                    DataGridView1.Rows(p).Cells(9).Value = DataGridView3.Rows(p).Cells(9).Value
                                    DataGridView1.Rows(p).Cells(14).Value = DataGridView3.Rows(p).Cells(11).Value
                                    Dim fh As New fguiasistema
                                    Dim gb As New vguiacabecera
                                    gb.gccia = Label30.Text
                                    gb.gCodArtIni = DataGridView3.Rows(p).Cells(5).Value
                                    gb.galmacen = Label23.Text
                                    If Trim(DataGridView1.Rows(p).Cells(6).Value) = "" Then
                                        gb.gFiltroDescrip = "1"
                                        gb.gModal = 1
                                    Else
                                        gb.gFiltroDescrip = DataGridView1.Rows(p).Cells(6).Value
                                        gb.gModal = 2
                                    End If
                                    gb.gano = Label29.Text
                                    hg = fh.stock_guia(gb)


                                    If hg.Rows.Count <> 0 Then
                                        DataGridView4.DataSource = hg
                                        DataGridView1.Rows(p).Cells(11).Value = DataGridView4.Rows(0).Cells(10).Value + DataGridView1.Rows(p).Cells(7).Value
                                    Else
                                        DataGridView1.Rows(p).Cells(11).Value = DataGridView1.Rows(p).Cells(7).Value
                                    End If
                                    DataGridView4.DataSource = Nothing

                                Next
                                DataGridView1.Columns(0).Frozen = True
                                DataGridView1.Columns(1).Frozen = True
                                DataGridView1.Columns(2).Frozen = True

                            End If
                            'noeditable()
                            DataGridView1.Enabled = False

                            PictureBox2.Enabled = True
                            PictureBox5.Enabled = True
                            Button1.Enabled = False
                            If RadioButton2.Checked = True Then
                                Label15.Visible = True
                            Else
                                Label15.Visible = False
                            End If

                            'RadioButton2.Checked = True
                            'RadioButton1.Checked = True
                            PictureBox1.Enabled = True
                        Else
                            MsgBox("LA GUIA PERTENECE A OTRO ALMACEN")
                        End If
                    Else


                        Select Case TextBox11.Text.Length
                            Case "1" : TextBox11.Text = "0000000" & "" & TextBox11.Text
                            Case "2" : TextBox11.Text = "000000" & "" & TextBox11.Text
                            Case "3" : TextBox11.Text = "00000" & "" & TextBox11.Text
                            Case "4" : TextBox11.Text = "0000" & "" & TextBox11.Text
                            Case "5" : TextBox11.Text = "000" & "" & TextBox11.Text
                            Case "6" : TextBox11.Text = "00" & "" & TextBox11.Text
                            Case "7" : TextBox11.Text = "0" & "" & TextBox11.Text
                            Case "8" : TextBox11.Text = TextBox11.Text
                        End Select
                        Dim func As New fguiasistema

                        Dim fu1 As New vguiacabecera
                        fu1.gmes = DateTime.Now.ToString("MM")
                        fu1.galmacen = Label23.Text
                        fu1.gano = Label29.Text
                        fu1.gccia = Label30.Text
                        Label3.Text = func.mostrar_correlativo_nsa(fu1) + 1
                        Select Case Label3.Text.Length

                            Case "1" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "000" & "" & Label3.Text
                            Case "2" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "00" & "" & Label3.Text
                            Case "3" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "0" & "" & Label3.Text
                            Case "4" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "" & Label3.Text
                        End Select
                        TextBox3.Enabled = True
                        TextBox4.Enabled = True
                        TextBox8.Enabled = True
                        TextBox10.Enabled = True
                        TextBox11.Enabled = True
                        TextBox13.Enabled = True
                        ComboBox1.Enabled = True
                        ComboBox2.Enabled = True
                        ComboBox3.Enabled = True
                        PictureBox1.Enabled = True

                        Button5.Enabled = True
                        PictureBox4.Enabled = True


                        TextBox4.Text = ""


                        ComboBox1.Text = ""
                        ComboBox2.Text = ""
                        ComboBox3.Text = ""
                    End If
                    TextBox4.Select()
                Else
                    Dim sql10210 As String = "SELECT ac.po,ac.op, linea,detalle,sum(total),um,sum(total) FROM requerimiento_avios_detalle ad INNER JOIN requerimiento_avios_cabecera ac on ad.id_cabecera = ac.id_requeminieto where id_cabecera ='" + Trim(Label35.Text) + "'  group by linea,detalle,ac.po,ac.op,um order by linea,detalle"
                    Dim cmd10210 As New SqlCommand(sql10210, conx)
                    Rsr215 = cmd10210.ExecuteReader()
                    Dim i5 As Integer
                    i5 = 0
                    While Rsr215.Read()
                        DataGridView1.Rows.Add()
                        DataGridView1.Rows(i5).Cells(0).Value = i5 + 1
                        DataGridView1.Rows(i5).Cells(1).Value = Rsr215(1)
                        DataGridView1.Rows(i5).Cells(2).Value = Rsr215(2)
                        DataGridView1.Rows(i5).Cells(3).Value = Rsr215(3)
                        DataGridView1.Rows(i5).Cells(4).Value = Rsr215(4)
                        DataGridView1.Rows(i5).Cells(5).Value = Rsr215(5)
                        DataGridView1.Rows(i5).Cells(7).Value = Rsr215(6)
                        DataGridView1.Rows(i5).Cells(8).Value = Rsr215(5)
                        i5 = i5 + 1
                    End While
                    Rsr215.Close()
                    Dim J As Integer
                    J = DataGridView1.Rows.Count
                    For i = 0 To J - 1
                        Dim sql10211 As String = "EXEC custom_vianny.dbo.CaeSoft_ReporteStockFisico '" + Trim(Label30.Text) + "','" + Trim(Label29.Text) + "','13','" + Trim(Label29.Text) + "0101','" + Trim(Label29.Text) + "1231','" + DataGridView1.Rows(i).Cells(2).Value + "','" + DataGridView1.Rows(i).Cells(2).Value + "',NULL, 1"
                        Dim cmd10211 As New SqlCommand(sql10211, conx)
                        Rsr211 = cmd10211.ExecuteReader()
                        If Rsr211.Read() Then
                            DataGridView1.Rows(i).Cells(11).Value = Rsr211(10)
                        Else
                            DataGridView1.Rows(i).Cells(11).Value = 0
                        End If

                        If DataGridView1.Rows(i).Cells(11).Value <= 0 Then
                            DataGridView1.Rows(i).Cells(13).Value = 1
                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                            DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                        Else
                            DataGridView1.Rows(i).Cells(13).Value = 0
                        End If
                        Rsr211.Close()
                    Next

                    TextBox3.Enabled = True
                    TextBox4.Enabled = True
                    TextBox8.Enabled = True
                    TextBox10.Enabled = True
                    TextBox11.Enabled = True
                    TextBox13.Enabled = True
                    ComboBox1.Enabled = True
                    ComboBox2.Enabled = True
                    ComboBox3.Enabled = True
                    PictureBox1.Enabled = True

                    Button5.Enabled = True
                    PictureBox4.Enabled = True
                End If
            Else
                MsgBox("EL PERIODO INGRESADO NO COINCIDE CON EL DE LA GUIA - INGRESE AL AÑO CORRECTO Y INGRESE A LA GUIA")
            End If




        Else
            If e.KeyCode = Keys.F1 Then
                Dim det As New Detalle_guia
                det.TextBox2.Text = TextBox1.Text
                det.Label3.Text = Label23.Text
                det.Label4.Text = Trim(Label30.Text)
                det.ShowDialog()
            End If

        End If
    End Sub

    Private Sub TextBox2_DoubleClick(sender As Object, e As EventArgs) Handles TextBox2.DoubleClick
        Dim func As New fguiasistema
        Dim fu As New vguiacabecera
        Dim fu1 As New vguiacabecera
        fu.gserie = TextBox1.Text
        fu.gccia = Label30.Text
        TextBox2.Text = func.mostrar_guia_correlativo(fu) + 1
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
    End Sub

    Private Sub TextBox24_TextChanged(sender As Object, e As EventArgs) Handles TextBox24.TextChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim fu As New vguiacabecera
            Dim func As New fguiasistema
            Dim fu1 As New vguiacabecera


            Select Case TextBox1.Text.Length

                Case "1" : TextBox1.Text = "000" & "" & TextBox1.Text
                Case "2" : TextBox1.Text = "00" & "" & TextBox1.Text
                Case "3" : TextBox1.Text = "0" & "" & TextBox1.Text
                Case "4" : TextBox1.Text = TextBox1.Text

            End Select

            fu.gserie = TextBox1.Text
            fu.gccia = Label30.Text

            TextBox2.Text = func.mostrar_guia_correlativo(fu) + 1
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

            TextBox8.Text = ""
            TextBox9.Text = ""
            TextBox10.Text = ""
            TextBox4.Text = ""
            TextBox7.Text = ""
            TextBox17.Text = ""
            TextBox5.Text = ""
            TextBox14.Text = ""
            TextBox15.Text = ""
            TextBox16.Text = ""
            TextBox6.Text = ""
            If Label30.Text = "01" Then
                TextBox11.Text = "20508740361"
                TextBox12.Text = "CONSORCIO TEXTIL VIANNY SAC"
                TextBox13.Text = "JR. LOS OLMOS NRO. 358 URB. CANTO BELLO (ALT. PARDERO 1 AV."
            Else
                TextBox11.Text = "20459785834"
                TextBox12.Text = "GRAUS INDUSTRIAS TEXTIL S.A."
                TextBox13.Text = "JR. LOS OLMOS NRO. 358 URB. CANTO BELLO (ALT. PARDERO 1 AV."
            End If

            DataGridView1.Rows.Clear()
            TextBox2.Select()
        End If
    End Sub
    Dim Rsr21, Rsr300129, Rsr212, Rsr3001296 As SqlDataReader
    Private Sub TextBox20_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox20.KeyDown
        If e.KeyCode = Keys.Enter Then

            DataGridView1.Rows.Clear()
            Select Case TextBox20.Text.Length

                Case "1" : TextBox20.Text = "010000000" & "" & TextBox20.Text
                Case "2" : TextBox20.Text = "01000000" & "" & TextBox20.Text
                Case "3" : TextBox20.Text = "0100000" & "" & TextBox20.Text
                Case "4" : TextBox20.Text = "010000" & "" & TextBox20.Text
                Case "5" : TextBox20.Text = "01000" & "" & TextBox20.Text
                Case "6" : TextBox20.Text = "0100" & "" & TextBox20.Text
                Case "7" : TextBox20.Text = "010" & "" & TextBox20.Text
                Case "8" : TextBox20.Text = "01" & "" & TextBox20.Text
            End Select

            Label37.Text = TextBox20.Text & "," & TextBox20.Text
            abrir()
            Dim cmd As New SqlDataAdapter("SELECT q.fich_3,q.descri_3,q.cants_3,cantp_3,porcadi_3,program_3,item_8,area_8,t.dele,nomb_8,a.linea_17,a.nomb_17 ,total,udm_8,ncom_3,g.nomb_10   FROM custom_vianny.dbo.qag0300 q 
inner join custom_vianny.dbo.qamc800 c on q.ccia = c.ccia_8 and q.ncom_3 =c.ncom_8  
left join custom_vianny.dbo.cag1700 a on c.linea_8 = a.linea_17 and c.ccia_8 =a.ccia 
left join custom_vianny.dbo.TAB0100 t on c.area_8 = t.cele
left join custom_vianny.dbo.cag1000 g on q.fich_3 = g.ruc_10  and g.ccia ='" + Trim(Label30.Text) + "'
where ncom_3='" + Trim(TextBox20.Text) + "' and q.ccia ='" + Trim(Label30.Text) + "' and  t.ccia='" + Trim(Label30.Text) + "' AND t.CTAB='AREAMC' and area_8 =ISNULL(" + ComboBox4.SelectedValue.ToString + ",area_8) order by cast(item_8 as int)", conx)

            Dim da12 As New DataTable
            cmd.Fill(da12)
            DataGridView5.DataSource = da12
            Dim K As Integer
            K = DataGridView5.Rows.Count
            If K > 0 Then
                DataGridView1.Rows.Add(K)
                DataGridView1.Columns(11).Visible = True
                Dim vg As Integer
                Dim suma As Integer
                Dim suma2 As String
                suma = 1

                For i = 0 To K - 1
                    Select Case i.ToString.Length
                        Case "1" : suma2 = "00" & "" & suma + i
                        Case "2" : suma2 = "0" & "" & suma + i
                        Case "3" : suma2 = suma + i
                    End Select
                    DataGridView1.Rows(i).Cells(0).Value = suma2
                    DataGridView1.Rows(i).Cells(1).Value = Trim(DataGridView5.Rows(i).Cells(5).Value)
                    DataGridView1.Rows(i).Cells(2).Value = Trim(DataGridView5.Rows(i).Cells(10).Value)
                    DataGridView1.Rows(i).Cells(3).Value = Trim(DataGridView5.Rows(i).Cells(11).Value)
                    DataGridView1.Rows(i).Cells(8).Value = Trim(DataGridView5.Rows(i).Cells(13).Value)
                    DataGridView1.Rows(i).Cells(5).Value = Trim(DataGridView5.Rows(i).Cells(13).Value)
                    DataGridView1.Rows(i).Cells(7).Value = Trim(DataGridView5.Rows(i).Cells(12).Value)
                    DataGridView1.Rows(i).Cells(4).Value = Trim(DataGridView5.Rows(i).Cells(12).Value)
                    DataGridView1.Rows(i).Cells(21).Value = suma2



                    Dim stock As Double
                    stock = 0
                    Dim sql1021 As String = "EXEC custom_vianny.dbo.CaeSoft_ReporteStockFisico '" + Trim(Label30.Text) + "','" + Trim(Label29.Text) + "','13','" + Trim(Label29.Text) + "0101','" + Trim(Label29.Text) + "1231','" + DataGridView1.Rows(i).Cells(2).Value + "','" + DataGridView1.Rows(i).Cells(2).Value + "',NULL, 1"
                    Dim cmd1021 As New SqlCommand(sql1021, conx)
                    Rsr21 = cmd1021.ExecuteReader()

                    If Rsr21.Read() Then

                        stock = Convert.ToDouble(Rsr21(10))

                        Rsr21.Close()
                        Dim sql1022139 As String = "select SUM(CanAfi) from Almacen_Ficticio where CodAfi='" + Trim(DataGridView5.Rows(i).Cells(10).Value) + "' and AlmAfi='" + Trim(Label23.Text) + "' and EmpAfi='" + Trim(Label30.Text) + "' group by CodAfi,AlmAfi,EmpAfi"
                        Dim cmd1022139 As New SqlCommand(sql1022139, conx)
                        Rsr300129 = cmd1022139.ExecuteReader()
                        Dim jo As Double
                        If Rsr300129.Read() Then
                            jo = Convert.ToDouble(Rsr300129(0))
                        Else
                            jo = 0
                        End If

                        Rsr300129.Close()
                        'MsgBox(stock.ToString())
                        'MsgBox(jo.ToString())
                        DataGridView1.Rows(i).Cells(11).Value = (stock - jo).ToString()
                        'MsgBox(DataGridView1.Rows(i).Cells(11).Value)
                    Else
                        DataGridView1.Rows(i).Cells(11).Value = 0
                        Rsr21.Close()
                    End If
                    Dim cmd16 As New SqlCommand("insert into Almacen_Ficticio (ItmAfi,CodAfi,AlmAfi,EmpAfi,GuiAfi,PerAfi,CanAfi) values (@ItmAfi,@CodAfi,@AlmAfi,@EmpAfi,@GuiAfi,@PerAfi,@CanAfi)", conx)
                    cmd16.Parameters.AddWithValue("@ItmAfi", Trim(DataGridView1.Rows(i).Cells(0).Value))
                    cmd16.Parameters.AddWithValue("@CodAfi", Trim(DataGridView5.Rows(i).Cells(10).Value))
                    cmd16.Parameters.AddWithValue("@AlmAfi", Trim(Label23.Text))
                    cmd16.Parameters.AddWithValue("@EmpAfi", Trim(Label30.Text))
                    cmd16.Parameters.AddWithValue("@GuiAfi", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                    cmd16.Parameters.AddWithValue("@PerAfi", Trim(Label29.Text))
                    cmd16.Parameters.AddWithValue("@CanAfi", Convert.ToDouble(DataGridView5.Rows(i).Cells(12).Value))
                    cmd16.ExecuteNonQuery()



                    If Convert.ToDouble(DataGridView1.Rows(i).Cells(11).Value) <= Convert.ToDouble(DataGridView1.Rows(i).Cells(7).Value) Then
                        DataGridView1.Rows(i).Cells(13).Value = 1
                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                        DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                    Else
                        DataGridView1.Rows(i).Cells(13).Value = 0
                    End If

                Next
                Label39.Text = "1"
            Else
                MsgBox("NO HAY INFORMACION EN OP PARA AGREGAR")
            End If
        End If


    End Sub

    Dim Rsr2222 As SqlDataReader

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

    End Sub
    Function ELIMINAR1(ByVal sql)
        abrir()
        comando = New SqlCommand(sql, conx)
        Dim i As Integer = comando.ExecuteNonQuery
        If i > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Function ELIMINAR(ByVal sql)
        abrir()
        comando = New SqlCommand(sql, conx)
        Dim i As Integer = comando.ExecuteNonQuery
        If i > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim i, a9, fila, i10, sum As Integer

        sum = DataGridView1.Rows.Count
        fila = e.RowIndex
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Partida" Then
            abrir()
            Dim sql102 As String = "select ncom_4, SUBSTRING(ncom_4,1,10)  from custom_vianny.dbo.matg040f where partidA_4 ='" + DataGridView1.Rows(e.RowIndex).Cells(10).Value + "'"
            Dim cmd102 As New SqlCommand(sql102, conx)
            Rsr2222 = cmd102.ExecuteReader()
            If Rsr2222.Read = True Then
                DataGridView1.Rows(e.RowIndex).Cells(1).Value = Rsr2222(1)
                DataGridView1.Rows(e.RowIndex).Cells(9).Value = Rsr2222(0)
            End If



        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Cant. Consumo" Then

            If Label39.Text = "1" Then
                abrir()
                Dim agregar As String = "delete  Almacen_Ficticio  WHERE  CodAfi = '" + Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value) + "' AND AlmAfi ='" + Trim(Label23.Text) + "' AND EmpAfi = '" + Trim(Label30.Text) + "' AND GuiAfi ='" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + "' AND PerAfi='" + Trim(Label29.Text) + "'"
                ELIMINAR1(agregar)
                For i10 = 0 To sum - 1

                    If Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value) = Trim(DataGridView1.Rows(i10).Cells(2).Value) Then

                        Dim stock As Double
                        stock = 0


                        Dim sql1021 As String = "EXEC custom_vianny.dbo.CaeSoft_ReporteStockFisico '" + Trim(Label30.Text) + "','" + Trim(Label29.Text) + "','13','" + Trim(Label29.Text) + "0101','" + Trim(Label29.Text) + "1231','" + DataGridView1.Rows(i10).Cells(2).Value + "','" + DataGridView1.Rows(i10).Cells(2).Value + "',NULL, 1"
                        Dim cmd1021 As New SqlCommand(sql1021, conx)
                        Rsr212 = cmd1021.ExecuteReader()

                        If Rsr212.Read() Then

                            stock = Convert.ToDouble(Rsr212(10))

                            Rsr212.Close()

                            Dim sql1022139 As String = "select SUM(CanAfi) from Almacen_Ficticio where CodAfi='" + Trim(DataGridView1.Rows(i10).Cells(2).Value) + "' and AlmAfi='" + Trim(Label23.Text) + "' and EmpAfi='" + Trim(Label30.Text) + "' group by CodAfi,AlmAfi,EmpAfi"
                            Dim cmd1022139 As New SqlCommand(sql1022139, conx)
                            Rsr3001296 = cmd1022139.ExecuteReader()
                            Dim jo As Double
                            If Rsr3001296.Read() Then
                                jo = Convert.ToDouble(Rsr3001296(0))
                            Else
                                jo = 0
                            End If

                            Rsr3001296.Close()

                            'MsgBox(stock.ToString())
                            'MsgBox(jo.ToString())
                            DataGridView1.Rows(i10).Cells(11).Value = (stock - jo).ToString()
                            'MsgBox(DataGridView1.Rows(i).Cells(11).Value)

                            'Registo de informacion
                            Dim cmd168 As New SqlCommand("insert into Almacen_Ficticio (ItmAfi,CodAfi,AlmAfi,EmpAfi,GuiAfi,PerAfi,CanAfi) values (@ItmAfi,@CodAfi,@AlmAfi,@EmpAfi,@GuiAfi,@PerAfi,@CanAfi)", conx)
                            cmd168.Parameters.AddWithValue("@ItmAfi", Trim(DataGridView1.Rows(i10).Cells(0).Value))
                            cmd168.Parameters.AddWithValue("@CodAfi", Trim(DataGridView1.Rows(i10).Cells(2).Value))
                            cmd168.Parameters.AddWithValue("@AlmAfi", Trim(Label23.Text))
                            cmd168.Parameters.AddWithValue("@EmpAfi", Trim(Label30.Text))
                            cmd168.Parameters.AddWithValue("@GuiAfi", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                            cmd168.Parameters.AddWithValue("@PerAfi", Trim(Label29.Text))
                            cmd168.Parameters.AddWithValue("@CanAfi", Convert.ToDouble(DataGridView1.Rows(i10).Cells(7).Value))
                            cmd168.ExecuteNonQuery()

                            'fin del registro
                        Else
                            DataGridView1.Rows(i10).Cells(11).Value = 0
                            Rsr212.Close()
                        End If
                    End If
                Next


            End If

            'If DataGridView1.Rows(fila).Cells(7).Value > DataGridView1.Rows(fila).Cells(11).Value Then
            '    MsgBox("LA CANTIDAD SOLICITADA ES MAYOR AL STOCK")
            '    DataGridView1.Rows(fila).Cells(7).Value = 0

        End If


    End Sub

    Private Sub Guia_hilo_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If Label35.Text <> "0" Then
            Dim p As Integer
            p = DataGridView1.Rows.Count
            For i = 0 To p - 1
                Dim cmd157 As New SqlCommand("update id_cabecera set estado =0 where id_requeminieto_detalle=@id_requeminieto_detalle", conx)
                cmd157.Parameters.AddWithValue("@id_requeminieto_detalle", Trim(Label21.Text))
                cmd157.ExecuteNonQuery()
            Next


        End If
        If Label39.Text = "1" Then
            abrir()
            Dim agregar As String = "delete  Almacen_Ficticio  WHERE  AlmAfi ='" + Trim(Label23.Text) + "' AND EmpAfi = '" + Trim(Label30.Text) + "' AND GuiAfi ='" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + "' AND PerAfi='" + Trim(Label29.Text) + "'"
            ELIMINAR1(agregar)
        End If

        Label21.Text = 0
    End Sub

    Private Sub TextBox25_TextChanged(sender As Object, e As EventArgs) Handles TextBox25.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Agregar_requerimiento.TextBox1.Text = Trim(TextBox10.Text)
        Agregar_requerimiento.Label2.Text = TextBox8.Text
        Agregar_requerimiento.Label5.Text = ComboBox3.Text
        Agregar_requerimiento.Label3.Text = Label13.Text
        Agregar_requerimiento.Label4.Text = Label10.Text
        Agregar_requerimiento.Label6.Text = 3
        Agregar_requerimiento.Show()
    End Sub

    Private Sub TextBox24_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox24.KeyDown
        If e.KeyCode = Keys.F1 Then
            proceso.Label2.Text = 1
            proceso.ShowDialog()
        End If
    End Sub

    Private Sub TextBox20_QueryAccessibilityHelp(sender As Object, e As QueryAccessibilityHelpEventArgs) Handles TextBox20.QueryAccessibilityHelp

    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress

    End Sub

    Private Sub Button5_SizeChanged(sender As Object, e As EventArgs) Handles Button5.SizeChanged

    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "OBS" Then
            'OBSERVACION.Label2.Text = 5
            'OBSERVACION.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(17).Value)
            'OBSERVACION.Label3.Text = e.RowIndex
            'OBSERVACION.ShowDialog()

            OB_SC.Label1.Text = e.RowIndex
            OB_SC.Label2.Text = 3
            OB_SC.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(16).Value)
            OB_SC.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(17).Value)
            OB_SC.TextBox3.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(18).Value)
            OB_SC.TextBox4.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(19).Value)
            OB_SC.TextBox5.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(20).Value)
            OB_SC.ShowDialog()
        End If
    End Sub
    Dim Rsr3465 As SqlDataReader
    Private Sub TextBox25_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox25.KeyDown
        If e.KeyCode = Keys.Enter Then
            abrir()
            Dim sql102213465 As String = "select linea_17,nomb_17,unid_17 from custom_vianny.dbo.cag1700 where linea_17 ='" + TextBox25.Text + "' and ccia ='" + Label30.Text + "'"
            Dim cmd102213465 As New SqlCommand(sql102213465, conx)
            Rsr3465 = cmd102213465.ExecuteReader()
            Dim i51 As Integer
            i51 = DataGridView1.Rows.Count
            If Rsr3465.Read() Then
                DataGridView1.Rows.Add()
                DataGridView1.Rows(i51).Cells(0).Value = i51 + 1
                DataGridView1.Rows(i51).Cells(2).Value = Rsr3465(0)
                DataGridView1.Rows(i51).Cells(3).Value = Rsr3465(1)
                DataGridView1.Rows(i51).Cells(5).Value = Rsr3465(2)
                DataGridView1.Rows(i51).Cells(8).Value = Rsr3465(2)
                Rsr3465.Close()
                Dim sql1021 As String = "EXEC custom_vianny.dbo.CaeSoft_ReporteStockFisico '" + Trim(Label30.Text) + "','" + Trim(Label29.Text) + "','13','" + Trim(Label29.Text) + "0101','" + Trim(Label29.Text) + "1231','" + DataGridView1.Rows(i51).Cells(2).Value + "','" + DataGridView1.Rows(i51).Cells(2).Value + "',NULL, 1"
                Dim cmd1021 As New SqlCommand(sql1021, conx)
                Rsr21 = cmd1021.ExecuteReader()

                If Rsr21.Read() Then
                    DataGridView1.Rows(i51).Cells(11).Value = Rsr21(10)
                    DataGridView1.Rows(i51).Cells(4).Value = Rsr21(10)
                    DataGridView1.Rows(i51).Cells(7).Value = Rsr21(10)
                Else
                    DataGridView1.Rows(i51).Cells(11).Value = 0
                    DataGridView1.Rows(i51).Cells(4).Value = 0
                    DataGridView1.Rows(i51).Cells(7).Value = 0
                End If
                Rsr21.Close()
                TextBox3.Text = ""
                i51 = i51 + 1
            End If
        End If
    End Sub
End Class