Imports System.Data.SqlClient
Imports System.Data.Sql

Public Class NsaHilo
    Public enunciado As SqlCommand
    Public respuesta As SqlDataReader
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Public enunciado3 As SqlCommand
    Public respuesta3 As SqlDataReader
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty, ty2, ty3, ty4, ty30 As New DataTable
    Dim dt1, dt2, HG As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Dim Rsr215 As SqlDataReader
    Sub llenar_combo_alamcenes(ByVal ser As String)

        Try

            conn = New SqlDataAdapter("select talm_15m + '  |  '+ nomb_15m as nombre, talm_15m as almacen from custom_vianny.dbo.Mag1500  where ccia ='" + Trim(Label13.Text) + "' and  flag_15m ='1' and  (seriencr_15m ='" + ser + "' or seriencr_15m ='1_2' )", conx)
            conn.Fill(ty30)
            ComboBox3.DataSource = ty30
            ComboBox3.DisplayMember = "nombre"
            ComboBox3.ValueMember = "almacen"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Dim Rsr24, Rsr25 As SqlDataReader
    Private Sub NsaHilo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        llenar_combo_alamcenes(Trim(TextBox18.Text))
        Dim sql1023 As String = "SELECT A.nomb_15m FROM custom_vianny.dbo.Mag1500 A  Where a.ccia ='" + Trim(Label13.Text) + "' and a.talm_15m ='" + Trim(ComboBox3.SelectedValue.ToString) + "'"
        Dim cmd1023 As New SqlCommand(sql1023, conx)
        Rsr24 = cmd1023.ExecuteReader()
        If Rsr24.Read() = True Then
            TextBox2.Text = Rsr24(0)
        End If
        Rsr24.Close()
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        Button2.Enabled = False
        DataGridView1.Enabled = False
        PictureBox1.Enabled = False
        Button5.Enabled = False
        Button7.Enabled = False
        Button8.Enabled = False
        Button6.Enabled = False
        RadioButton1.Checked = True
        ComboBox2.SelectedIndex = 0
        ComboBox6.SelectedIndex = 0
        TextBox4.Focus()
        TextBox4.Select()
    End Sub


    Sub llenar_combo_box2(ByVal cb As ComboBox, ByVal alm As String)
        Try

            conn = New SqlDataAdapter("select dest_21s,rtrim(ltrim(nomb_21s)) as nomb_21s from custom_vianny.dbo.cag210s where almac_21s =" + alm + " AND Empr_21s =" + Trim(Label13.Text) + "  order by nomb_21s", conx)
            conn.Fill(ty2)
            ComboBox1.DataSource = ty2
            ComboBox1.DisplayMember = "nomb_21s"
            ComboBox1.ValueMember = "dest_21s"
            'respuesta = enunciado.ExecuteReader
            'While respuesta.Read
            '    cb.Items.Add(respuesta.Item("nomb_17f"))
            'End While
            'respuesta.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs)
        ty2.Clear()
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        Button2.Enabled = False
        DataGridView1.Enabled = False
        DataGridView1.Rows.Clear()
        PictureBox1.Enabled = False
        Button5.Enabled = False
        Button7.Enabled = False
        Button8.Enabled = False
        Button6.Enabled = False
        RadioButton1.Checked = True
        limpiar()
        TextBox8.Enabled = False
        TextBox5.ReadOnly = False
        TextBox4.ReadOnly = False
        TextBox1.ReadOnly = False
        TextBox5.Enabled = True
        TextBox4.Enabled = True
        ComboBox3.Enabled = True
        TextBox4.Select()
    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs)

    End Sub
    Dim Rsr30, Rsr211, Rsr9 As SqlDataReader
    Private Sub PictureBox3_Click(sender As Object, e As EventArgs)

    End Sub
    Public Sub limpiar()
        TextBox9.Text = ""
        TextBox5.Text = ""
        TextBox8.Text = ""
        TextBox10.Text = ""
        TextBox1.Text = ""
        'TextBox17.Text = ""
        'TextBox19.Text = ""
    End Sub
    Private Sub PictureBox4_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        If Trim(TextBox8.Text).Length = 0 Then
            MessageBox.Show("FALTA INGRESAR EL CLIENTE", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If Trim(ComboBox3.Text) = "59  |  SERVICIOS A TERCEROS" Then

                ToolTip1.SetToolTip(PictureBox1, "AGREGAR PRODUCTOS")
                Dim art As New Articulos_stock
                art.Label3.Text = ComboBox3.SelectedValue.ToString
                art.Label4.Text = 400
                art.Label5.Text = Label10.Text
                art.Label6.Text = Label13.Text
                art.Label7.Text = 2
                art.ShowDialog()
            Else
                Dim prod As New Productos_Hilo
                prod.Label3.Text = ComboBox3.SelectedValue.ToString
                prod.Label4.Text = 2
                prod.Label5.Text = Label10.Text
                prod.Label6.Text = Label13.Text
                Select Case ComboBox3.Text
                    Case "13" : prod.Label7.Text = "1"
                    Case "05" : prod.Label7.Text = "1"
                    Case "14" : prod.Label7.Text = "1"
                    Case "11" : prod.Label7.Text = "1"

                    Case Else
                        prod.Label7.Text = "2"
                End Select
                prod.Show()
            End If
        End If

    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim respuesta As DialogResult
        Dim I18, VAL As Integer
        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
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
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub



    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        ty2.Clear()
        'abrir()
        'llenar_combo_box2(ComboBox1, ComboBox3.Text)


        Dim sql1023 As String = " SELECT A.nomb_15m FROM custom_vianny.dbo.Mag1500 A  Where a.ccia ='" + Trim(Label13.Text) + "' and a.talm_15m ='" + Trim(ComboBox3.SelectedValue.ToString) + "'"
        Dim cmd1023 As New SqlCommand(sql1023, conx)
        Rsr25 = cmd1023.ExecuteReader()
        If Rsr25.Read() = True Then
            TextBox2.Text = Rsr25(0)

        End If
        Rsr25.Close()
        TextBox4.Text = DateTimePicker1.Value.Month
        Select Case TextBox4.Text.Length
            Case "1" : TextBox4.Text = "0" & DateTimePicker1.Value.Month
            Case "2" : TextBox4.Text = TextBox4.Text

        End Select
        TextBox4.Enabled = True
        TextBox4.Select()


    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If Label13.Text = "01" Then
            If ComboBox2.Text = "INTERNA" Then
                TextBox8.Text = "20508740361"
                TextBox10.Text = "CONSORCIO TEXTIL VIANNY"
            Else
                TextBox8.ReadOnly = False
            End If
        Else
            If ComboBox2.Text = "INTERNA" Then
                TextBox8.Text = "20459785834"
                TextBox10.Text = "GRAUS INDUSTRIAS TEXTIL"

            End If
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            Try
                Dim JU As Integer
                Select Case TextBox4.Text.Length

                    Case "1" : TextBox4.Text = "0" & "" & TextBox4.Text
                    Case "2" : TextBox4.Text = TextBox4.Text
                End Select
                TextBox5.Enabled = True
                abrir()
                'MsgBox("select top 1 ncom_3, substring(ncom_3,3,7) as ncom from custom_vianny.dbo.mag030f  where ccia_3 = '01' and cperiodo_3 =  YEAR(GETDATE()) and talm_3 = " + ComboBox3.Text + " " + "and ncom_3 like" + " " + TextBox4.Text + "% and ccom_3 = '2' order by ncom_3 desc")
                enunciado = New SqlCommand("select top 1 ncom_3, substring(ncom_3,3,7) as ncom from custom_vianny.dbo.mag030f  where ccia_3 = " + Trim(Label13.Text) + " and cperiodo_3 = '" + Label10.Text + "' and talm_3 = '" + ComboBox3.SelectedValue.ToString + "' " + "and ncom_3 like" + " '" + TextBox4.Text + "%' and ccom_3 = '2' order by ncom_3 desc", conx)
                respuesta = enunciado.ExecuteReader
                While respuesta.Read
                    JU = respuesta.Item("ncom")
                End While
                respuesta.Close()
                TextBox5.Text = JU + 1
                Select Case TextBox5.Text.Length

                    Case "1" : TextBox5.Text = "00000" & "" & TextBox5.Text
                    Case "2" : TextBox5.Text = "0000" & "" & TextBox5.Text
                    Case "3" : TextBox5.Text = "000" & "" & TextBox5.Text
                    Case "4" : TextBox5.Text = "00" & "" & TextBox5.Text
                    Case "5" : TextBox5.Text = "0" & "" & TextBox5.Text
                    Case "6" : TextBox5.Text = TextBox5.Text
                End Select

                TextBox5.Select()
                TextBox5.ReadOnly = False
            Catch ex As Exception
            End Try
        Else
            If e.KeyCode = Keys.F1 Then
                Ingresos_Salidas_almacen.Label3.Text = Label13.Text
                Ingresos_Salidas_almacen.Label4.Text = Label10.Text
                Ingresos_Salidas_almacen.Label5.Text = ComboBox3.SelectedValue.ToString
                Ingresos_Salidas_almacen.Label6.Text = Trim(TextBox4.Text)
                Ingresos_Salidas_almacen.Label7.Text = "3"
                Ingresos_Salidas_almacen.Show()
            End If
        End If
    End Sub

    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            ty2.Clear()
            abrir()
            DataGridView1.Rows.Clear()
            TextBox4.Enabled = False
            TextBox5.Enabled = False
            If Label21.Text = "0" Then
                llenar_combo_box2(ComboBox1, ComboBox3.SelectedValue.ToString)

                If Trim(TextBox4.Text) = "14" Then
                    Dim me12 As String
                    me12 = Mid(TextBox5.Text, 1, 2) - Month(DateTimePicker1.Value)
                    DateTimePicker1.Value = DateTimePicker1.Value.AddMonths(me12)
                Else
                    Dim me12 As String
                    me12 = TextBox4.Text - Month(DateTimePicker1.Value)
                    DateTimePicker1.Value = DateTimePicker1.Value.AddMonths(me12)
                End If




                Dim ml As New vnia
                Dim mk As New fnsa
                Dim i As Integer
                i = Len(TextBox5.Text)

                Select Case i
                    Case "1" : TextBox5.Text = "00000" & "" & TextBox5.Text
                    Case "2" : TextBox5.Text = "0000" & "" & TextBox5.Text
                    Case "3" : TextBox5.Text = "000" & "" & TextBox5.Text
                    Case "4" : TextBox5.Text = "00" & "" & TextBox5.Text
                    Case "5" : TextBox5.Text = "0" & "" & TextBox5.Text
                    Case "6" : TextBox5.Text = TextBox5.Text

                End Select
                ml.gcomprobante = TextBox4.Text & TextBox5.Text

                ml.galmacen1 = ComboBox3.SelectedValue.ToString
                ml.gano = Label10.Text
                ml.gccia = Label13.Text
                If mk.mostrar_correlativo_nsa1(ml) = True Then
                    Dim jk As New fnsa

                    Button6.Enabled = True
                    Button8.Enabled = True
                    DataGridView1.Rows.Clear()
                    dt1 = jk.mostrar_cabecera_nsa(ml)
                    dt2 = jk.mostrar_detalle_nsa(ml)
                    If dt1.Rows.Count <> 0 Then
                        DataGridView3.DataSource = dt1

                        TextBox9.Text = DataGridView3.Rows(0).Cells(0).Value
                        TextBox8.Text = DataGridView3.Rows(0).Cells(4).Value
                        If Convert.ToString(DataGridView3.Rows(0).Cells(6).Value) = "" Then
                            TextBox10.Text = ""
                        Else
                            TextBox10.Text = DataGridView3.Rows(0).Cells(6).Value
                        End If

                        DateTimePicker1.Value = DataGridView3.Rows(0).Cells(1).Value
                        TextBox1.Text = DataGridView3.Rows(0).Cells(2).Value
                        If Trim(DataGridView3.Rows(0).Cells(9).Value) = "0" Then
                            CheckBox3.Checked = False
                        Else
                            CheckBox3.Checked = True
                        End If
                        abrir()
                        enunciado2 = New SqlCommand("select rtrim(ltrim(nomb_21s)) as nomb_21s from custom_vianny.dbo.cag210s where Empr_21s ='" + Trim(Label13.Text) + "' and almac_21s='" + ComboBox3.SelectedValue.ToString + "'  and dest_21s='" + Trim(DataGridView3.Rows(0).Cells(5).Value) + "'", conx)
                        respuesta2 = enunciado2.ExecuteReader
                        While respuesta2.Read

                            ComboBox1.Text = respuesta2.Item("nomb_21s")
                        End While
                        respuesta2.Close()
                        If Trim(DataGridView3.Rows(0).Cells(8).Value) = 1 Then
                            TextBox13.Text = ""
                            TextBox11.Text = ""
                        Else
                            TextBox11.Text = DataGridView3.Rows(0).Cells(8).Value
                            Select Case TextBox11.Text.Length
                                Case "1" : TextBox11.Text = "000" & "" & TextBox11.Text
                                Case "2" : TextBox11.Text = "00" & "" & TextBox11.Text
                                Case "3" : TextBox11.Text = "0" & "" & TextBox11.Text
                                Case "4" : TextBox11.Text = TextBox11.Text


                            End Select
                            enunciado3 = New SqlCommand(" SELECT  nomb_4 as Nombre FROM custom_vianny.dbo.Qag0400 A  Where CCIA = '" + Trim(Label13.Text) + "' and flag_4 ='1' and fase_4 ='" + Trim(TextBox11.Text) + "'", conx)
                            respuesta3 = enunciado3.ExecuteReader
                            While respuesta3.Read
                                TextBox13.Text = respuesta3.Item("Nombre")
                            End While
                            respuesta3.Close()
                        End If

                        ComboBox4.Text = "GUIA"
                        If Trim(DataGridView3.Rows(0).Cells(7).Value) = "EXT" Then
                            ComboBox2.SelectedIndex = 0
                        Else
                            If Trim(DataGridView3.Rows(0).Cells(7).Value) = "INT" Then
                                ComboBox2.SelectedIndex = 1
                            End If

                        End If
                        If DataGridView3.Rows(0).Cells(3).Value = 1 Then
                            RadioButton2.Checked = False
                            RadioButton1.Checked = True
                            RadioButton1.Enabled = True
                            Label12.Visible = False
                            Button7.Enabled = False
                        Else
                            RadioButton2.Checked = True
                            RadioButton1.Checked = False
                            RadioButton1.Enabled = False
                            Label12.Visible = True
                        End If

                    End If
                    If dt2.Rows.Count <> 0 Then
                        Dim nu1 As Integer
                        DataGridView4.DataSource = dt2
                        nu1 = DataGridView4.Rows.Count
                        DataGridView1.Rows.Add(nu1)
                        For i1 = 0 To nu1 - 1

                            DataGridView1.Rows(i1).Cells(0).Value = DataGridView4.Rows(i1).Cells(0).Value
                            DataGridView1.Rows(i1).Cells(1).Value = DataGridView4.Rows(i1).Cells(1).Value
                            DataGridView1.Rows(i1).Cells(2).Value = DataGridView4.Rows(i1).Cells(2).Value
                            DataGridView1.Rows(i1).Cells(3).Value = DataGridView4.Rows(i1).Cells(3).Value
                            DataGridView1.Rows(i1).Cells(4).Value = DataGridView4.Rows(i1).Cells(4).Value
                            DataGridView1.Rows(i1).Cells(6).Value = DataGridView4.Rows(i1).Cells(5).Value
                            DataGridView1.Rows(i1).Cells(7).Value = DataGridView4.Rows(i1).Cells(6).Value
                            DataGridView1.Rows(i1).Cells(11).Value = DataGridView4.Rows(i1).Cells(7).Value
                            DataGridView1.Rows(i1).Cells(8).Value = DataGridView4.Rows(i1).Cells(8).Value
                            DataGridView1.Rows(i1).Cells(5).Value = DataGridView4.Rows(i1).Cells(9).Value
                            DataGridView1.Rows(i1).Cells(9).Value = DataGridView4.Rows(i1).Cells(10).Value
                            DataGridView1.Rows(i1).Cells(12).Value = DataGridView4.Rows(i1).Cells(11).Value
                            Dim fh As New fguiasistema
                            Dim gb As New vguiacabecera
                            gb.gccia = Label13.Text
                            gb.gCodArtIni = DataGridView4.Rows(i1).Cells(2).Value
                            gb.galmacen = ComboBox3.SelectedValue.ToString
                            If DataGridView1.Rows(i1).Cells(8).Value = "" Then
                                gb.gFiltroDescrip = "null"
                                gb.gModal = 1
                            Else
                                gb.gFiltroDescrip = DataGridView4.Rows(i1).Cells(8).Value
                                gb.gModal = 2
                            End If
                            gb.gano = Label10.Text

                            HG = fh.stock_guia(gb)
                            DataGridView5.DataSource = HG
                            If HG.Rows.Count <> 0 Then
                                DataGridView1.Rows(i1).Cells(10).Value = DataGridView5.Rows(0).Cells(10).Value + DataGridView1.Rows(i1).Cells(6).Value
                            Else
                                DataGridView1.Rows(i1).Cells(10).Value = DataGridView1.Rows(i1).Cells(6).Value
                            End If

                        Next
                        DataGridView1.Columns(0).Frozen = True
                        DataGridView1.Columns(1).Frozen = True
                        DataGridView1.Columns(2).Frozen = True
                    End If
                Else
                    ComboBox1.Enabled = True
                    ComboBox2.Enabled = True
                    ComboBox3.Enabled = False
                    ComboBox4.Enabled = True
                    TextBox8.Enabled = True
                    TextBox9.Enabled = True
                    TextBox1.Enabled = True
                    TextBox1.ReadOnly = False
                    Button2.Enabled = True
                    DataGridView1.Enabled = True
                    PictureBox1.Enabled = True
                    Button5.Enabled = True
                    Button7.Enabled = True
                    DataGridView1.Rows.Clear()
                    TextBox9.Text = ""
                    TextBox1.Text = ""
                    TextBox8.Text = ""
                    TextBox10.Text = ""
                End If
                ComboBox3.Enabled = False
                ComboBox4.SelectedIndex = -1
                'TextBox1.Text = 1
                DateTimePicker1.Select()
                Button1.Enabled = True
                TextBox15.Enabled = True

            Else
                llenar_combo_box2(ComboBox1, ComboBox3.SelectedValue.ToString)
                Dim sql10210 As String = "SELECT ac.po,ac.op, linea,detalle,sum(total),um,sum(total) FROM requerimiento_avios_detalle ad INNER JOIN requerimiento_avios_cabecera ac on ad.id_cabecera = ac.id_requeminieto where id_cabecera ='" + Trim(Label21.Text) + "'   and  estado1 = 1 group by linea,detalle,ac.po,ac.op,um order by linea,detalle"
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
                    DataGridView1.Rows(i5).Cells(6).Value = Rsr215(6)
                    DataGridView1.Rows(i5).Cells(11).Value = Rsr215(5)
                    i5 = i5 + 1
                End While
                Rsr215.Close()
                Dim J As Integer
                J = DataGridView1.Rows.Count
                For i = 0 To J - 1
                    Dim sql10211 As String = "EXEC custom_vianny.dbo.CaeSoft_ReporteStockFisico '" + Trim(Label13.Text) + "','" + Trim(Label10.Text) + "','13','" + Trim(Label10.Text) + "0101','" + Trim(Label10.Text) + "1231','" + DataGridView1.Rows(i).Cells(2).Value + "','" + DataGridView1.Rows(i).Cells(2).Value + "',NULL, 1"
                    Dim cmd10211 As New SqlCommand(sql10211, conx)
                    Rsr211 = cmd10211.ExecuteReader()
                    If Rsr211.Read() Then
                        DataGridView1.Rows(i).Cells(10).Value = Rsr211(10)
                    Else
                        DataGridView1.Rows(i).Cells(10).Value = 0
                    End If

                    If DataGridView1.Rows(i).Cells(10).Value <= 0 Then
                        DataGridView1.Rows(i).Cells(16).Value = 1
                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                        DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                    Else
                        DataGridView1.Rows(i).Cells(16).Value = 0
                    End If
                    Rsr211.Close()
                Next


                TextBox8.Select()
                Button6.Enabled = True
                Button8.Enabled = True
                ComboBox1.Enabled = True
                ComboBox2.Enabled = True
                ComboBox3.Enabled = False
                ComboBox4.Enabled = True
                TextBox8.Enabled = True
                TextBox9.Enabled = True
                TextBox1.Enabled = True
                TextBox1.ReadOnly = False
                Button2.Enabled = True
                DataGridView1.Enabled = True
                PictureBox1.Enabled = True
                Button5.Enabled = True
                Button7.Enabled = True
                'modulo_solicitud.DataGridView1.Rows.Clear()
                modulo_solicitud.Button1_Click(Me, EventArgs.Empty)
            End If
        End If

    End Sub
    Public Sub NumConFrac(ByVal CajaTexto As Windows.Forms.TextBox, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Char.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False

        Else
            e.Handled = True
        End If
    End Sub
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        'If Trim(TextBox4.Text) <> Month(DateTimePicker1.Value) Then
        '    MsgBox("LA FECHA DE EM0ISION NO CORRESPONDE AL MES DE REGISTRO")
        '    Dim me12 As String
        '    me12 = TextBox4.Text - Month(DateTimePicker1.Value)
        '    DateTimePicker1.Value = DateTimePicker1.Value.AddMonths(me12)
        'End If
    End Sub

    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.F1 Then
            If Trim(TextBox5.Text).Length > 0 Then
                Dim cli As New Clientes
                cli.TextBox3.Text = "400"
                cli.ShowDialog()

            Else
                MessageBox.Show("FALTA INGRESAR EL CORRELATIVO DEL COMPROBANTE", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub TextBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox11.KeyDown
        If e.KeyCode = Keys.F1 Then
            Procesos.Label1.Text = "4"
            Procesos.Label2.Text = Label13.Text
            Procesos.Show()
            TextBox8.Select()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim fun As New vinsertar_nia
        Dim func As New fnia
        Dim func1 As New insertardetallenia
        Dim func2 As New vnia
        Dim con As String
        Dim hn As New vnia
        Dim fg As New fnia
        Dim x, mj, valo, mp, mop, final, correlativo As String
        Dim pp As Integer
        Dim i1, a1, vg As Integer
        i1 = DataGridView1.Rows.Count
        a1 = 0
        For a1 = 0 To i1 - 1
            vg = vg + DataGridView1.Rows(a1).Cells(16).Value
        Next

        Dim cant, Cop, Cup, Ccon, Cpartida, Cum, va1, valor As Integer
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
            If Trim(DataGridView1.Rows(i1).Cells(6).Value).Length = 0 Then
                Ccon = Ccon + 1
            End If
            If Trim(DataGridView1.Rows(i1).Cells(8).Value).Length = 0 Then
                Cpartida = Cpartida + 1
            End If
            If Trim(DataGridView1.Rows(i1).Cells(5).Value).Length = 0 Then
                Cum = Cum + 1
            End If
        Next


        If vg > 0 Then
            MsgBox("HAY PRODUCTOS QUE NO TIENEN STOCK")
        Else
            If Cop > 0 And Trim(ComboBox1.Text) <> "VENTA" Then
                MessageBox.Show("FALTA INGRESAR LA OP EN UNA CELDA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If Cup > 0 Then
                    MessageBox.Show("FALTA INGRESAR LA CANTIDAD DE PRESENTACION EN ALGUNA CELDA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    If Ccon > 0 Then
                        MessageBox.Show("FALTA INGRESAR LA CANTIDAD EN KG/UND EN ALGUNA CELDA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Else
                        If Cpartida > 0 And (Trim(ComboBox3.SelectedValue.ToString) = "03" Or Trim(ComboBox3.SelectedValue.ToString) = "59") Then
                            MessageBox.Show("FALTA INGRESAR EL LOTE O PARTIDA EN ALGUNA CELDA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Else
                            If Cum > 0 Then
                                MessageBox.Show("FALTA INGRESAR LA UM DE PRESENTACION EN ALGUNA CELDA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Else
                                x = Convert.ToString(TextBox1.Text.IndexOf("-", 0) + 1)
                                If x > 0 Then
                                    valo = Trim(TextBox1.Text)
                                    Select Case x

                                        Case "1" : mj = "0000"
                                        Case "2" : mj = "000" & TextBox1.Text.Substring(0, x - 1)
                                        Case "3" : mj = "00" & TextBox1.Text.Substring(0, x - 1)
                                        Case "4" : mj = "0" & TextBox1.Text.Substring(0, x - 1)
                                        Case "5" : mj = TextBox1.Text.Substring(0, x - 1)
                                    End Select
                                    pp = Convert.ToInt32(TextBox1.Text.Length) - Convert.ToInt32(x)
                                    mp = TextBox1.Text.Substring(x, pp)
                                    Select Case mp.Length
                                        Case "1" : mop = "0000000" & mp
                                        Case "2" : mop = "000000" & mp
                                        Case "3" : mop = "00000" & mp
                                        Case "4" : mop = "0000" & mp
                                        Case "5" : mop = "000" & mp
                                        Case "6" : mop = "00" & mp
                                        Case "7" : mop = "0" & mp
                                        Case "8" : mop = mp
                                    End Select
                                    final = mj & "-" & mop
                                Else
                                    final = Trim(TextBox1.Text) & Trim(Label23.Text)
                                End If
                                Select Case TextBox4.Text.Length
                                    Case "1" : TextBox4.Text = "0" & TextBox4.Text
                                    Case "2" : TextBox4.Text = TextBox4.Text
                                End Select
                                Select Case TextBox5.Text.Length
                                    Case "1" : TextBox5.Text = "00000" & TextBox5.Text
                                    Case "2" : TextBox5.Text = "0000" & TextBox5.Text
                                    Case "3" : TextBox5.Text = "000" & TextBox5.Text
                                    Case "4" : TextBox5.Text = "00" & TextBox5.Text
                                    Case "5" : TextBox5.Text = "0" & TextBox5.Text
                                    Case "6" : TextBox5.Text = TextBox5.Text
                                End Select
                                Dim sql10233 As String = "select COUNT(ncom_3) from custom_vianny.dbo.mag030f where CCIA_3 ='" + Label13.Text + "' and CPERIODO_3 ='" + Label10.Text + "' and talm_3 ='" + ComboBox3.SelectedValue.ToString + "' and ncom_3 ='" + TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() + "' and ccom_3 ='2'"
                                Dim cmd10233 As New SqlCommand(sql10233, conx)
                                Rsr9 = cmd10233.ExecuteReader()
                                If Rsr9.Read() = True Then
                                    va1 = Convert.ToInt32(Rsr9(0))
                                End If
                                Rsr9.Close()
                                If va1 > 0 Then
                                    Dim func125 As New fnsa
                                    Dim dts4 As New vnia
                                    TextBox4.Text = DateTime.Now.ToString("MM")

                                    dts4.gmes = Me.TextBox4.Text
                                    dts4.galmacen = ComboBox3.SelectedValue.ToString
                                    dts4.gano = Label10.Text
                                    dts4.gccia = Label13.Text
                                    correlativo = func125.buscar_nsa(dts4) + 1
                                    Select Case correlativo.Length

                                        Case "1" : correlativo = "00000" & "" & correlativo
                                        Case "2" : correlativo = "0000" & "" & correlativo
                                        Case "3" : correlativo = "000" & "" & correlativo
                                        Case "4" : correlativo = "00" & "" & correlativo
                                        Case "5" : correlativo = "0" & "" & correlativo
                                        Case "6" : correlativo = correlativo
                                    End Select
                                    con = TextBox4.Text & correlativo
                                    MsgBox("La Nota de Salida Ingresada Ya existe, Se esta Actualizando al nuevo numero de Salida:  " + con)
                                Else
                                    con = TextBox4.Text & TextBox5.Text
                                End If

                                If Label24.Text = "1" Then
                                    hn.gcomprobante = con
                                    hn.galmacen = ComboBox3.SelectedValue.ToString
                                    hn.gncom = 2
                                    hn.gano = Label10.Text
                                    hn.gccia = Label13.Text
                                    fg.eliminar_nia(hn)
                                End If

                                Dim ns1 As New vnsacabece
                                Dim ns2 As New fnsa
                                ns1.gncom = con
                                ns1.gfecha = DateTimePicker1.Value
                                ns1.gorige_sali = ComboBox1.SelectedValue.ToString
                                ns1.gglosa = TextBox9.Text
                                Select Case ComboBox4.Text
                                    Case "GUIA" : ns1.gdoc = "009"
                                    Case "FACT" : ns1.gdoc = "001"
                                    Case "OTRO" : ns1.gdoc = "002"
                                    Case "" : ns1.gdoc = ""
                                End Select
                                ns1.gguia = final
                                ns1.gsfactu = Mid(final, 1, 4)
                                ns1.gnfactu = Mid(final, 6, 8)
                                ns1.gruc = TextBox8.Text
                                ns1.galmacen = ComboBox3.SelectedValue.ToString
                                ns1.gusuario = TextBox7.Text
                                ns1.gano = Label10.Text
                                ns1.gccia = Label13.Text
                                If ComboBox2.Text = "INTERNA" Then
                                    ns1.gtipointexte = "INT"
                                Else
                                    ns1.gtipointexte = "EXT"
                                End If
                                Select Case ComboBox4.Text
                                    Case "GUIA" : ns1.gtdocento = "009"
                                    Case "FACT" : ns1.gtdocento = "001"
                                    Case "OTRO" : ns1.gtdocento = "002"
                                    Case "" : ns1.gtdocento = ""
                                End Select
                                ns1.gfase = TextBox11.Text
                                ns1.gadevol = "0"
                                ns1.gpedorig_3 = Trim(Label21.Text)
                                Dim i, num2 As Integer
                                num2 = DataGridView1.Rows.Count
                                For i = 0 To num2 - 1
                                    Dim nu As String
                                    nu = DataGridView1.Rows(i).Cells(0).Value
                                    Select Case nu.Length
                                        Case "1" : nu = "00" & "" & nu
                                        Case "2" : nu = "0" & "" & nu
                                        Case "3" : nu = nu
                                    End Select
                                    Dim ns3 As New vnsadetalle
                                    Dim ns4 As New fnsa
                                    ns3.gncom = con
                                    ns3.gitem = nu
                                    ns3.glinea = DataGridView1.Rows(i).Cells(2).Value
                                    ns3.gcantidad = DataGridView1.Rows(i).Cells(6).Value
                                    ns3.gund = DataGridView1.Rows(i).Cells(11).Value
                                    If DataGridView1.Rows(i).Cells(1).Value = "" Then
                                        ns3.gop = ""
                                    Else
                                        ns3.gop = DataGridView1.Rows(i).Cells(1).Value
                                    End If
                                    If Convert.ToString(DataGridView1.Rows(i).Cells(5).Value) = "" Then
                                        ns3.gunidad_medidad = ""
                                    Else
                                        ns3.gunidad_medidad = DataGridView1.Rows(i).Cells(5).Value
                                    End If

                                    If Convert.ToString(DataGridView1.Rows(i).Cells(7).Value) = "" Then
                                        ns3.gpartida = ""
                                    Else
                                        ns3.gpartida = DataGridView1.Rows(i).Cells(7).Value
                                    End If
                                    ns3.grollo1 = DataGridView1.Rows(i).Cells(4).Value
                                    ns3.galmacen = ComboBox3.SelectedValue.ToString
                                    If Convert.ToString(DataGridView1.Rows(i).Cells(9).Value) = "" Then
                                        ns3.gordtejeduria = ""
                                    Else
                                        ns3.gordtejeduria = DataGridView1.Rows(i).Cells(9).Value
                                    End If
                                    ns3.gano = Label10.Text
                                    If Convert.ToString(DataGridView1.Rows(i).Cells(8).Value) = "" Then
                                        ns3.glote = ""
                                    Else
                                        ns3.glote = DataGridView1.Rows(i).Cells(8).Value
                                    End If
                                    ns3.gcantenvio = DataGridView1.Rows(i).Cells(12).Value
                                    ns3.gccia = Label13.Text
                                    If Trim(ComboBox3.SelectedValue.ToString) = "59" Then

                                        ns3.gclasif = "1"
                                    Else
                                        ns3.gclasif = ""
                                    End If
                                    ns3.gOcompra = ""
                                    ns4.insertar_detalle_nsa(ns3)

                                    If Trim(DataGridView1.Rows(i).Cells(14).Value) <> "" Then
                                        abrir()
                                        Dim cmd157 As New SqlCommand("update requerimiento_avios_detalle set estado =2 where id_requeminieto_detalle=@id_requeminieto_detalle", conx)
                                        cmd157.Parameters.AddWithValue("@id_requeminieto_detalle", Trim(DataGridView1.Rows(i).Cells(14).Value))
                                        cmd157.ExecuteNonQuery()
                                        Dim a, b As Integer

                                        Dim sql102213 As String = "select COUNT(id_requeminieto_detalle) from requerimiento_avios_detalle where id_cabecera ='" + Trim(DataGridView1.Rows(i).Cells(15).Value) + "' and estado ='2'"
                                        Dim cmd102213 As New SqlCommand(sql102213, conx)
                                        Rsr3 = cmd102213.ExecuteReader()
                                        If Rsr3.Read() = True Then
                                            a = Rsr3(0)
                                        End If
                                        Rsr3.Close()
                                        Dim sql1022134 As String = "select COUNT(id_requeminieto_detalle) from requerimiento_avios_detalle where id_cabecera ='" + Trim(DataGridView1.Rows(i).Cells(15).Value) + "'"
                                        Dim cmd1022134 As New SqlCommand(sql1022134, conx)
                                        Rsr21 = cmd1022134.ExecuteReader()
                                        If Rsr21.Read() = True Then
                                            b = Rsr21(0)
                                        End If
                                        Rsr21.Close()
                                        'MsgBox(a)
                                        'MsgBox(b)
                                        If a = b Then
                                            Dim cmd1577 As New SqlCommand("update requerimiento_avios_cabecera set estado = 4 where id_requeminieto='" + Trim(DataGridView1.Rows(i).Cells(15).Value) + "'", conx)
                                            cmd1577.Parameters.AddWithValue("@id_requeminieto_detalle", Trim(DataGridView1.Rows(i).Cells(14).Value))
                                            cmd1577.ExecuteNonQuery()
                                        End If
                                    End If

                                Next
                                If ns2.insertar_cabecera_nsa(ns1) Then

                                    MessageBox.Show("Datos registrados correctamente ", "Guardando registros", MessageBoxButtons.OK, MessageBoxIcon.Information)

                                    Dim hj2 As New flog
                                    Dim hj1 As New vlog
                                    hj1.gmodulo = "NSA-ALMACEN"
                                    hj1.gcuser = MDIParent1.Label3.Text
                                    If Label20.Text = "0" Then
                                        hj1.gaccion = 1
                                    Else
                                        hj1.gaccion = 2
                                    End If

                                    hj1.gpc = My.Computer.Name
                                    hj1.gfecha = DateTimePicker2.Value
                                    hj1.gdetalle = Trim(ComboBox3.SelectedValue.ToString) & Trim(TextBox4.Text) & Trim(TextBox5.Text)
                                    hj1.gccia = Label13.Text
                                    hj2.insertar_log(hj1)
                                    Dim respuesta As DialogResult
                                    respuesta = MessageBox.Show("DESEA IMPRIMIR LA NOTA DE SALIDA?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                                    If (respuesta = Windows.Forms.DialogResult.Yes) Then

                                        Reporte_Nia_Nsa.TextBox1.Text = ComboBox3.SelectedValue.ToString
                                        Reporte_Nia_Nsa.TextBox2.Text = 2
                                        Reporte_Nia_Nsa.TextBox3.Text = TextBox4.Text & "" & TextBox5.Text
                                        Reporte_Nia_Nsa.TextBox4.Text = Label10.Text
                                        Reporte_Nia_Nsa.TextBox5.Text = Label13.Text
                                        Reporte_Nia_Nsa.Show()
                                    End If
                                    limpiar()

                                    RadioButton1.Checked = True

                                End If

                                TextBox4.Enabled = False
                                TextBox5.Enabled = True
                                TextBox9.Enabled = False
                                TextBox8.Enabled = False
                                DateTimePicker1.Enabled = False
                                DataGridView1.Enabled = False
                                ComboBox1.Enabled = False
                                ComboBox2.Enabled = False
                                ComboBox4.Enabled = False
                                DataGridView1.Rows.Clear()


                                TextBox5.Select()
                                Dim func12 As New fnsa
                                Dim dts As New vnia
                                TextBox4.Text = DateTime.Now.ToString("MM")

                                dts.gmes = Me.TextBox4.Text
                                dts.galmacen = ComboBox3.SelectedValue.ToString
                                dts.gano = Label10.Text
                                dts.gccia = Label13.Text
                                Me.TextBox5.Text = func12.buscar_nsa(dts) + 1
                                Select Case TextBox5.Text.Length

                                    Case "1" : TextBox5.Text = "00000" & "" & TextBox5.Text
                                    Case "2" : TextBox5.Text = "0000" & "" & TextBox5.Text
                                    Case "3" : TextBox5.Text = "000" & "" & TextBox5.Text
                                    Case "4" : TextBox5.Text = "00" & "" & TextBox5.Text
                                    Case "5" : TextBox5.Text = "0" & "" & TextBox5.Text
                                    Case "6" : TextBox5.Text = TextBox5.Text
                                End Select
                                TextBox5.Select()
                                Label21.Text = 0
                                Label24.Text = "0"
                            End If
                                End If
                            End If
                        End If
                        End If
                            End If


    End Sub

    Dim ok As New DataTable
    Dim Rsr3, Rsr21, Rsr300 As SqlDataReader
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.Rows.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("SELECT q.fich_3,q.descri_3,q.cants_3,cantp_3,porcadi_3,program_3,item_8,area_8,t.dele,nomb_8,linea_3,a.nomb_17 ,total,udm_8,ncom_3,g.nomb_10   FROM custom_vianny.dbo.qag0300 q 
inner join custom_vianny.dbo.qamc800 c on q.ccia = c.ccia_8 and q.ncom_3 =c.ncom_8  
left join custom_vianny.dbo.cag1700 a on q.linea_3 = a.linea_17 and q.ccia =a.ccia 
left join custom_vianny.dbo.TAB0100 t on c.area_8 = t.cele
left join custom_vianny.dbo.cag1000 g on q.fich_3 = g.ruc_10  and g.ccia ='02'
where ncom_3='" + Trim(TextBox15.Text) + "' and q.ccia ='" + Trim(Label13.Text) + "' and  t.ccia='01' AND t.CTAB='AREAMC' order by item_8", conx)
        Dim da As New DataTable
        cmd.Fill(da)
        DataGridView7.DataSource = da
        Dim K As Integer
        K = DataGridView7.Rows.Count
        DataGridView1.Rows.Add(K)
        For i = 0 To K - 1
            DataGridView1.Rows(i).Cells(0).Value = DataGridView7.Rows(i).Cells(6).Value
            DataGridView1.Rows(i).Cells(1).Value = DataGridView7.Rows(i).Cells(5).Value
            DataGridView1.Rows(i).Cells(2).Value = DataGridView7.Rows(i).Cells(10).Value
            DataGridView1.Rows(i).Cells(3).Value = DataGridView7.Rows(i).Cells(11).Value
            DataGridView1.Rows(i).Cells(11).Value = DataGridView7.Rows(i).Cells(13).Value
            DataGridView1.Rows(i).Cells(6).Value = DataGridView7.Rows(i).Cells(12).Value
        Next

    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim ll As New fnsa
        Dim kk As New vnsadetalle
        Dim i, fila As Integer
        i = DataGridView1.Rows.Count
        fila = e.RowIndex
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Partida" Then

            kk.gpartida = DataGridView1.Rows(fila).Cells(7).Value
            kk.gccia = Label13.Text
            ok = ll.buscar_ot_po(kk)
            DataGridView6.DataSource = ok
            If ok.Rows.Count > 0 Then
                DataGridView1.Rows(fila).Cells(1).Value = DataGridView6.Rows(0).Cells(1).Value
                DataGridView1.Rows(fila).Cells(9).Value = DataGridView6.Rows(0).Cells(0).Value
            Else
                DataGridView6.DataSource = ""
                DataGridView1.Rows(fila).Cells(1).Value = ""
                DataGridView1.Rows(fila).Cells(9).Value = ""
            End If

        Else
            Dim i1, a9, fila1 As Integer

            i1 = DataGridView1.Rows.Count
            fila1 = e.RowIndex
            If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Cant. Consumo" Then

                If DataGridView1.Rows(fila1).Cells(6).Value > DataGridView1.Rows(fila1).Cells(10).Value Then
                    MsgBox("LA CANTIDAD SOLICITADA ES MAYOR AL STOCK")
                    DataGridView1.Rows(fila1).Cells(6).Value = 0

                End If
            End If

        End If
    End Sub
    Sub llenar_combo_box3()
        Try

            conn = New SqlDataAdapter("SELECT ncom_3 FROM custom_vianny.DBO.qag0300 where program_3 ='" + Trim(TextBox14.Text) + "' order by ncom_3", conx)
            conn.Fill(ty4)
            ComboBox5.DataSource = ty4
            ComboBox5.DisplayMember = "ncom_3"
            ComboBox5.ValueMember = "ncom_3"
            'respuesta = enunciado.ExecuteReader
            'While respuesta.Read
            '    cb.Items.Add(respuesta.Item("nomb_17f"))
            'End While
            'respuesta.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Dim Rsr20 As SqlDataReader
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        '  abrir()
        '  Dim sql1020 As String = "SELECT A.ccia_8,a.ncom_8,a.linea_8,a.factor_8,b.nomb_17,b.talm_17,b.unid_17 FROM custom_vianny.dbo.Qamc800 A
        ' LEFT JOIN custom_vianny.dbo.CAG1700 B ON '01' = B.ccia AND A.LINEA_8 = B.linea_17
        'Where A.ccia_8='" + Label13.Text + "' AND A.ncom_8='" + ComboBox5.SelectedValue.ToString + "' and talcol_8 =0"
        '  Dim cmd1020 As New SqlCommand(sql1020, conx)
        '  Rsr20 = cmd1020.ExecuteReader()
        '  Dim i5 As Integer
        '  i5 = 0
        '  While Rsr20.Read()
        '      DataGridView1.Rows.Add()
        '      DataGridView1.Rows(i5).Cells(0).Value = i5 + 1
        '      DataGridView1.Rows(i5).Cells(1).Value = Rsr20(1)
        '      DataGridView1.Rows(i5).Cells(2).Value = Rsr20(2)
        '      DataGridView1.Rows(i5).Cells(3).Value = Rsr20(4)
        '      DataGridView1.Rows(i5).Cells(11).Value = Rsr20(6)
        '      i5 = i5 + 1

        '  End While
        '  Rsr20.Close()
    End Sub

    Private Sub TextBox14_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox14.KeyDown
        If e.KeyCode = Keys.Enter Then
            ty4.Clear()
            abrir()
            llenar_combo_box3()
        End If
    End Sub

    Private Sub DateTimePicker1_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker1.KeyDown
        If e.KeyCode = Keys.Enter Then
            ComboBox1.Select()
        End If
    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged

    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox11.Select()
        End If
    End Sub

    Private Sub TextBox5_LostFocus(sender As Object, e As EventArgs) Handles TextBox5.LostFocus

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'If Trim(TextBox8.Text) = "" Then
        '    MsgBox("NO EXISTE NINGUN RUC, INGRESELO PARA UTILIZAR ESTA OPCION")
        'Else
        Agregar_requerimiento.TextBox1.Text = Trim(TextBox10.Text)
            Agregar_requerimiento.Label2.Text = TextBox8.Text
        Agregar_requerimiento.Label5.Text = ComboBox3.SelectedValue.ToString
        Agregar_requerimiento.Label3.Text = Label13.Text
            Agregar_requerimiento.Label4.Text = Label10.Text
            Agregar_requerimiento.Label6.Text = 1
            Agregar_requerimiento.Show()
        'End If
    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged

    End Sub

    Private Sub GroupBox4_Enter(sender As Object, e As EventArgs) Handles GroupBox4.Enter

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Label15.Text = "OP"
            ComboBox5.Enabled = False
        Else
            Label15.Text = "PO"
            ComboBox5.Enabled = True
        End If
    End Sub
    Dim loDataTable As New DataTable
    Sub llenar_combo_box4()
        Try
            loDataTable.Rows.Clear()
            Dim lsQuery As String = "select ocorte_3cg from custom_vianny.dbo.Qag300Cc where pedido_3cg ='" + Trim(ComboBox5.Text) + "'"

            Dim loDataAdapter As New SqlDataAdapter(lsQuery, conx)
            loDataAdapter.Fill(loDataTable)
            ComboBox7.DataSource = loDataTable

            ComboBox7.DisplayMember = "ocorte_3cg"
            ComboBox7.ValueMember = "ocorte_3cg"
            ComboBox7.DropDownWidth = 90
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        abrir()
        llenar_combo_box4()

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If RadioButton2.Checked = True Then
            MsgBox("NO SE PUEDE EDITAR INFORMACION EN UNA NOTA DE SALIDA ANULADA")
        Else

            Dim sql102213 As String = "SELECT cierre_3 FROM custom_vianny.dbo.Cie0300 A  Where a.ccia ='" + Label13.Text + "' AND A.talm_3 ='" + ComboBox3.SelectedValue.ToString + "' AND A.fcie_3='" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "'"
            Dim cmd102213 As New SqlCommand(sql102213, conx)
            Rsr300 = cmd102213.ExecuteReader()
            Dim jo As Integer
            If Rsr300.Read() Then
                jo = Rsr300(0)
            Else
                jo = 0
            End If
            Rsr300.Close()
            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("QUIERES EDITAR ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                If jo = 0 Then
                    DataGridView1.Enabled = True
                    TextBox9.Enabled = True
                    DateTimePicker1.Enabled = True
                    ComboBox1.Enabled = True
                    ComboBox2.Enabled = True
                    ComboBox4.Enabled = True
                    TextBox4.Enabled = False
                    TextBox1.ReadOnly = False
                    TextBox8.ReadOnly = False
                    TextBox1.Enabled = True
                    TextBox5.Enabled = False
                    TextBox8.Enabled = True
                    Button2.Enabled = True
                    PictureBox1.Enabled = True
                    Button5.Enabled = True
                    Button7.Enabled = True
                    Label20.Text = "1"
                    Label24.Text = "1"
                Else
                    MsgBox("DIA " & DateTimePicker1.Value & " CERRADO,NO SE PUEDE EDITAR LA INFORMACION - CONSULTE CON EL AREA CONTABLE ")
                End If

            End If
        End If

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim sql102 As String = "SELECT cierre_3 FROM custom_vianny.dbo.Cie0300 A  Where a.ccia ='" + Label13.Text + "' AND A.talm_3 ='" + ComboBox3.SelectedValue.ToString + "' AND A.fcie_3='" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "'"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr30 = cmd102.ExecuteReader()
        Dim jo As Integer
        If Rsr30.Read() Then
            jo = Rsr30(0)
        Else
            jo = 0
        End If
        Rsr30.Close()
        Dim respuesta As DialogResult
        Dim ml As New vnsacabece
        Dim mk As New fnsa
        ml.gncom = TextBox4.Text & TextBox5.Text
        ml.galmacen = ComboBox3.SelectedValue.ToString
        ml.gano = Label10.Text
        ml.gccia = Label13.Text
        respuesta = MessageBox.Show("QUIERES ANULAR NOTA DE SALIDA?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            If jo = 0 Then
                Label12.Visible = True
                RadioButton2.Checked = True
                RadioButton1.Checked = False

                mk.anular_nSA(ml)
                TextBox4.Enabled = True
                TextBox4.Select()
                DateTimePicker1.Enabled = False
                TextBox9.Enabled = False

                TextBox5.Enabled = True

                DataGridView1.Enabled = False
                Button5.Enabled = False
                Button8.Enabled = False
                Button7.Enabled = False
                Button6.Enabled = False

                Label12.Visible = False
                Dim hj2 As New flog
                Dim hj1 As New vlog
                hj1.gmodulo = "NSA-ALMACEN"
                hj1.gcuser = MDIParent1.Label3.Text

                hj1.gaccion = 3


                hj1.gpc = My.Computer.Name
                hj1.gfecha = DateTimePicker2.Value
                hj1.gdetalle = Trim(ComboBox3.SelectedValue.ToString) & Trim(TextBox4.Text) & Trim(TextBox5.Text)
                hj1.gccia = Label13.Text
                hj2.insertar_log(hj1)
                MsgBox("SE ANULO LA INFORMACION CORRECTAMENTE")
                TextBox4.Enabled = False
                TextBox5.Enabled = True
                TextBox9.Enabled = False
                TextBox8.Enabled = False
                DateTimePicker1.Enabled = False
                DataGridView1.Enabled = False
                ComboBox1.Enabled = False
                ComboBox2.Enabled = False
                ComboBox4.Enabled = False
                DataGridView1.Rows.Clear()


                TextBox5.Select()
                Dim func12 As New fnsa
                Dim dts As New vnia
                TextBox4.Text = DateTime.Now.ToString("MM")

                dts.gmes = Me.TextBox4.Text
                dts.galmacen = ComboBox3.SelectedValue.ToString
                dts.gano = Label10.Text
                dts.gccia = Label13.Text
                Me.TextBox5.Text = func12.buscar_nsa(dts) + 1
                Select Case TextBox5.Text.Length

                    Case "1" : TextBox5.Text = "00000" & "" & TextBox5.Text
                    Case "2" : TextBox5.Text = "0000" & "" & TextBox5.Text
                    Case "3" : TextBox5.Text = "000" & "" & TextBox5.Text
                    Case "4" : TextBox5.Text = "00" & "" & TextBox5.Text
                    Case "5" : TextBox5.Text = "0" & "" & TextBox5.Text
                    Case "6" : TextBox5.Text = TextBox5.Text
                End Select
            Else
                MsgBox("DIA " & DateTimePicker1.Value & " CERRADO,NO SE PUEDE EDITAR LA INFORMACION - CONSULTE CON EL AREA CONTABLE ")
            End If

        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        ToolTip1.SetToolTip(Button8, "IMPRIMIR COMPROANTE")
        Reporte_Nia_Nsa.TextBox1.Text = ComboBox3.Text
        Reporte_Nia_Nsa.TextBox2.Text = 2
        Reporte_Nia_Nsa.TextBox3.Text = TextBox4.Text & "" & TextBox5.Text
        Reporte_Nia_Nsa.TextBox4.Text = Label10.Text
        Reporte_Nia_Nsa.TextBox5.Text = Label13.Text
        Reporte_Nia_Nsa.Show()
    End Sub

    Private Sub DataGridView1_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles DataGridView1.CellValidating
        If e.ColumnIndex = 4 Or e.ColumnIndex = 6 Then
            Dim cellValue As String = e.FormattedValue.ToString()
            If Not IsValidNumber(cellValue) Then
                MessageBox.Show("El valor ingresado no es válido. Debe ser un número entero o decimal con al menos un número antes del punto decimal.", "Error de Validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Cancel = True
            End If
        End If
    End Sub
    Private Function IsValidNumber(value As String) As Boolean
        Dim pattern As String = "^\d+(\.\d+)?$"
        Return System.Text.RegularExpressions.Regex.IsMatch(value, pattern)
    End Function
    Private Sub NsaHilo_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Me.Activate()  ' Activa el formulario
        Me.Focus()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        ty2.Clear()
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        Button2.Enabled = False
        DataGridView1.Enabled = False
        DataGridView1.Rows.Clear()
        PictureBox1.Enabled = False
        Button5.Enabled = False
        Button7.Enabled = False
        Button8.Enabled = False
        Button6.Enabled = False
        RadioButton1.Checked = True
        limpiar()
        TextBox8.Enabled = False
        TextBox5.ReadOnly = False
        TextBox4.ReadOnly = False
        TextBox1.ReadOnly = False
        TextBox5.Enabled = True
        TextBox4.Enabled = True
        ComboBox3.Enabled = True
        TextBox4.Select()
    End Sub

    Private Sub NsaHilo_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If Label21.Text <> "0" Then
            Dim p As Integer
            p = DataGridView1.Rows.Count
            For i = 0 To p - 1
                Dim cmd157 As New SqlCommand("update requerimiento_avios_detalle set estado =0 where id_cabecera=@id_requeminieto_detalle", conx)
                cmd157.Parameters.AddWithValue("@id_requeminieto_detalle", Label21.Text)
                cmd157.ExecuteNonQuery()
            Next


        End If
        Label21.Text = 0

    End Sub



    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        NumConFrac(TextBox5, e)
    End Sub
    Dim dt As New DataTable
    Private Sub TextBox16_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox16.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button2.Enabled = True
            Dim func As New fventas
            Dim dts As New vventas

            DataGridView1.Enabled = True
            Button5.Enabled = True


            If TextBox16.Text = "" Then


                MsgBox("EL CAMPO CODIGO DE BARRA ESTA VACIO")
            Else

                dts.gcodigo = Trim(TextBox16.Text)
                dt = func.buscar_codigo_barra(dts)
                If dt.Rows.Count <> 0 Then
                    DataGridView1.Rows.Add(1)
                    DataGridView8.DataSource = dt

                    Dim num, num2, i17 As Integer
                    Dim num12, num10, re As Integer
                    num = 0
                    re = 0
                    If DataGridView1.Rows(0).Cells(0).Value = 0 Then

                        DataGridView1.Rows(0).Cells(0).Value = num + 1
                        Select Case DataGridView1.Rows(0).Cells(0).Value.ToString.Length
                            Case "1" : DataGridView1.Rows(0).Cells(0).Value = "00" & "" & DataGridView1.Rows(0).Cells(0).Value
                            Case "2" : DataGridView1.Rows(0).Cells(0).Value = "0" & "" & DataGridView1.Rows(0).Cells(0).Value
                            Case "3" : DataGridView1.Rows(0).Cells(0).Value = DataGridView1.Rows(0).Cells(0).Value

                        End Select
                        DataGridView1.Rows(0).Cells(2).Value = DataGridView8.Rows(0).Cells(0).Value
                        DataGridView1.Rows(0).Cells(3).Value = DataGridView8.Rows(0).Cells(1).Value
                        DataGridView1.Rows(0).Cells(5).Value = DataGridView8.Rows(0).Cells(2).Value
                        DataGridView1.Rows(0).Cells(11).Value = DataGridView8.Rows(0).Cells(2).Value
                        DataGridView1.Rows(0).Cells(4).Value = 1
                        DataGridView1.Rows(0).Cells(6).Value = 1
                        DataGridView1.Rows(0).Cells(1).ReadOnly = False
                        'DataGridView1.Rows(0).Cells(7).ReadOnly = False
                        'DataGridView1.Columns(7).MaxInputLength = 6
                        CType(DataGridView1.Columns(7), DataGridViewTextBoxColumn).MaxInputLength = 8
                        DataGridView8.DataSource = ""
                        TextBox16.Text = ""

                    Else
                        num10 = DataGridView1.Rows.Count
                        For i = 1 To num10 - 1
                            If Convert.ToInt16(DataGridView1.Rows(i).Cells(0).Value) < Convert.ToInt16(DataGridView1.Rows(i - 1).Cells(0).Value) Then
                                num2 = Convert.ToInt16(DataGridView1.Rows(i - 1).Cells(0).Value) + 1
                                i17 = i

                            End If


                        Next
                        Dim j As Integer
                        j = 0

                        'If i17 = "200" Then
                        '    MsgBox("SOLO SE PERMITE 50 REGISTROS")
                        'Else
                        If num12 <> 1 Then
                            For i = 0 To i17 - 1
                                DataGridView1.Rows(i).Selected = False
                                If DataGridView8.Rows(0).Cells(0).Value = DataGridView1.Rows(i).Cells(2).Value Then
                                    re = re + 1
                                    j = i
                                End If
                            Next
                            If re >= 1 Then
                                    DataGridView1.Rows(j).Cells(4).Value = DataGridView1.Rows(j).Cells(4).Value + 1
                                    DataGridView1.Rows(j).Cells(6).Value = DataGridView1.Rows(j).Cells(6).Value + 1
                                    TextBox16.Text = ""
                                    DataGridView8.DataSource = ""
                                DataGridView1.Rows.RemoveAt(DataGridView1.Rows.Count - 1)

                            Else

                                DataGridView1.Rows(i17).Cells(0).Value = num2
                                    Select Case DataGridView1.Rows(i17).Cells(0).Value.ToString.Length
                                        Case "1" : DataGridView1.Rows(i17).Cells(0).Value = "00" & "" & DataGridView1.Rows(i17).Cells(0).Value
                                        Case "2" : DataGridView1.Rows(i17).Cells(0).Value = "0" & "" & DataGridView1.Rows(i17).Cells(0).Value
                                        Case "3" : DataGridView1.Rows(i17).Cells(0).Value = DataGridView1.Rows(i17).Cells(0).Value

                                    End Select
                                    DataGridView1.Rows(i17).Cells(2).Value = DataGridView8.Rows(0).Cells(0).Value
                                    DataGridView1.Rows(i17).Cells(3).Value = DataGridView8.Rows(0).Cells(1).Value
                                    DataGridView1.Rows(i17).Cells(5).Value = DataGridView8.Rows(0).Cells(2).Value
                                    DataGridView1.Rows(i17).Cells(11).Value = DataGridView8.Rows(0).Cells(2).Value
                                    DataGridView1.Rows(i17).Cells(4).Value = 1
                                    DataGridView1.Rows(i17).Cells(6).Value = 1
                                    DataGridView1.Rows(i17).Selected = True

                                    DataGridView1.Rows(i17).Cells(1).ReadOnly = False
                                    DataGridView1.Rows(i17).Cells(5).ReadOnly = False
                                    DataGridView1.Rows(i17).Cells(4).ReadOnly = False
                                    'DataGridView1.Rows(i17).Cells(7).ReadOnly = True
                                    TextBox16.Text = ""
                                    DataGridView8.DataSource = ""
                                End If





                        End If
                        'End If


                    End If

                Else
                    MsgBox("EL CODIGO DE BARRA NO EXISTE")
                    TextBox16.Text = ""
                    DataGridView7.DataSource = ""
                End If
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        TextBox17.Text = DataGridView1.Rows(e.RowIndex).Cells(10).Value
        validar_ingresos(e.ColumnIndex, e.RowIndex)
    End Sub
    Sub validar_ingresos(columna As Integer, fila As Integer)
        Dim columnIndex As Integer = 5 ' Puedes cambiar este valor al índice de la columna que necesitas
        Dim columnIndex1 As Integer = 4
        Dim columnIndex2 As Integer = 6
        Dim columnIndex3 As Integer = 10
        If columna = columnIndex AndAlso fila >= 0 Then
            ' Activa el modo de edición de la celda
            DataGridView1.BeginEdit(True)
        End If
        If columna = columnIndex1 AndAlso fila >= 0 Then
            ' Activa el modo de edición de la celda
            DataGridView1.BeginEdit(True)
        End If
        If columna = columnIndex2 AndAlso fila >= 0 Then
            ' Activa el modo de edición de la celda
            DataGridView1.BeginEdit(True)
        End If
        If columna = columnIndex3 AndAlso fila >= 0 Then
            ' Activa el modo de edición de la celda
            DataGridView1.BeginEdit(True)
        End If
    End Sub

    Private Sub DataGridView1_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEnter
        validar_ingresos(e.ColumnIndex, e.RowIndex)
    End Sub

    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        Dim columnIndex2 As Integer = 1 ' Cambia este valor al índice de la columna que quieres validar
        Dim columnIndex As Integer = 6 ' Cambia este valor al índice de la columna que quieres validar
        Dim columnIndex1 As Integer = 4 ' Cambia este valor al índice de la columna que quieres validar
        If DataGridView1.CurrentCell.ColumnIndex = columnIndex AndAlso TypeOf e.Control Is TextBox Then
            ' Agrega el controlador de eventos KeyPress al cuadro de texto de edición
            AddHandler DirectCast(e.Control, TextBox).KeyPress, AddressOf DataGridView1_KeyPress
        End If
        If DataGridView1.CurrentCell.ColumnIndex = columnIndex1 AndAlso TypeOf e.Control Is TextBox Then
            ' Agrega el controlador de eventos KeyPress al cuadro de texto de edición
            AddHandler DirectCast(e.Control, TextBox).KeyPress, AddressOf DataGridView1_KeyPress
        End If
        If DataGridView1.CurrentCell.ColumnIndex = columnIndex2 AndAlso TypeOf e.Control Is TextBox Then
            ' Agrega el controlador de eventos KeyPress al cuadro de texto de edición
            AddHandler DirectCast(e.Control, TextBox).KeyPress, AddressOf DataGridView1_KeyPress
        End If
    End Sub

    Private Sub DataGridView1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles DataGridView1.KeyPress
        Dim currentCell As DataGridViewCell = DataGridView1.CurrentCell
        If currentCell.ColumnIndex = 6 Then
            If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ControlChars.Back AndAlso e.KeyChar <> "." Then
                ' Si el carácter no es numérico, de retroceso ni un punto, cancela el evento KeyPress
                e.Handled = True
            End If
            If e.KeyChar = "." AndAlso DirectCast(sender, TextBox).Text.IndexOf(".") > -1 Then
                e.Handled = True
            End If
        End If
        If currentCell.ColumnIndex = 4 Then
            If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ControlChars.Back Then
                ' Si el carácter no es numérico ni de retroceso, cancela el evento KeyPress
                e.Handled = True
            End If
        End If
        If currentCell.ColumnIndex = 1 Then
            If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ControlChars.Back Then
                ' Si el carácter no es numérico ni de retroceso, cancela el evento KeyPress
                e.Handled = True
            End If
        End If
    End Sub
End Class