Imports System.Data.SqlClient
Imports System.Data.Sql
Public Class NiaHilo
    Public enunciado As SqlCommand
    Public respuesta As SqlDataReader
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Public enunciado3 As SqlCommand
    Public respuesta3 As SqlDataReader
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty, ty2, ty3, ty30, ty32, ty33 As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Sub llenar_combo_box2(ByVal alm As String)
        Try

            conn = New SqlDataAdapter("select dest_21e,rtrim(ltrim(nomb_21e)) as nomb_21e from custom_vianny.dbo.cag210e where Empr_21e =" + Trim(Label13.Text) + " AND almac_21e =" + alm, conx)
            conn.Fill(ty2)
            ComboBox1.DataSource = ty2
            ComboBox1.DisplayMember = "nomb_21e"
            ComboBox1.ValueMember = "dest_21e"
            'respuesta = enunciado.ExecuteReader
            'While respuesta.Read
            '    cb.Items.Add(respuesta.Item("nomb_17f"))
            'End While
            'respuesta.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Dim Rsr24, Rsr25 As SqlDataReader
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        ty2.Clear()
        Label15.Text = 1

        Dim sql1023 As String = " SELECT A.nomb_15m FROM custom_vianny.dbo.Mag1500 A  Where a.ccia ='" + Trim(Label13.Text) + "' and a.talm_15m ='" + Trim(ComboBox3.SelectedValue.ToString) + "'"
        Dim cmd1023 As New SqlCommand(sql1023, conx)
        Rsr25 = cmd1023.ExecuteReader()
        If Rsr25.Read() = True Then
            TextBox2.Text = Rsr25(0)

        End If
        Rsr25.Close()
        abrir()



        TextBox4.Text = DateTimePicker1.Value.Month
        Select Case TextBox4.Text.Length
            Case "1" : TextBox4.Text = "0" & DateTimePicker1.Value.Month
            Case "2" : TextBox4.Text = TextBox4.Text

        End Select
        TextBox4.Select()
        ComboBox4.SelectedIndex = 0

        TextBox4.ReadOnly = False
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
    Private Sub NiaHilo_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        DateTimePicker1.Enabled = True
        TextBox9.Enabled = False
        ComboBox2.SelectedIndex = 0

        TextBox12.Enabled = False
        DataGridView1.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox4.Enabled = False

        DateTimePicker1.Enabled = False
        Button3.Enabled = False
        Button5.Enabled = False
        Button7.Enabled = False
        TextBox1.Enabled = False
        Button6.Enabled = False
        Button1.Enabled = False
        RadioButton1.Enabled = False
        RadioButton2.Enabled = False

        abrir()
        llenar_combo_alamcenes(Trim(TextBox16.Text))
        'Dim sql1023 As String = "SELECT A.nomb_15m FROM custom_vianny.dbo.Mag1500 A  Where a.ccia ='" + Trim(Label13.Text) + "' and a.talm_15m ='" + Trim(ComboBox3.SelectedValue.ToString) + "'"
        'Dim cmd1023 As New SqlCommand(sql1023, conx)
        'Rsr24 = cmd1023.ExecuteReader()
        'If Rsr24.Read() = True Then
        '    TextBox2.Text = Rsr24(0)

        'End If
        'Rsr24.Close()


    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub
    Sub correlativo()
        Select Case TextBox4.Text.Length
            Case "1" : TextBox4.Text = "0" & TextBox4.Text
            Case "2" : TextBox4.Text = TextBox4.Text

        End Select
        TextBox5.Select()
        TextBox5.Enabled = True
        TextBox5.ReadOnly = False


        Dim func As New fnia
        Dim dts As New vnia
        Select Case TextBox4.Text.Length
            Case "1" : dts.gmes = "0" & TextBox4.Text
            Case "2" : dts.gmes = TextBox4.Text

        End Select
        dts.galmacen = ComboBox3.SelectedValue.ToString
        dts.gano = Label11.Text
        dts.gccia = Label13.Text
        Me.TextBox5.Text = func.buscar_nia(dts) + 1
        Select Case TextBox5.Text.Length

            Case "1" : TextBox5.Text = "00000" & "" & TextBox5.Text
            Case "2" : TextBox5.Text = "0000" & "" & TextBox5.Text
            Case "3" : TextBox5.Text = "000" & "" & TextBox5.Text
            Case "4" : TextBox5.Text = "00" & "" & TextBox5.Text
            Case "5" : TextBox5.Text = "0" & "" & TextBox5.Text
            Case "6" : TextBox5.Text = TextBox5.Text
        End Select

    End Sub
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            correlativo()

        Else
            If e.KeyCode = Keys.F1 Then
                Ingresos_Salidas_almacen.Label3.Text = Label13.Text
                Ingresos_Salidas_almacen.Label4.Text = Label11.Text
                Ingresos_Salidas_almacen.Label5.Text = ComboBox3.SelectedValue.ToString
                Ingresos_Salidas_almacen.Label6.Text = Trim(TextBox4.Text)
                Ingresos_Salidas_almacen.Label7.Text = "2"
                Ingresos_Salidas_almacen.Show()
            End If
        End If
    End Sub
    Private Sub bloqueado()
        DateTimePicker1.Enabled = True
        TextBox9.Enabled = False

        TextBox12.Enabled = False
        DataGridView1.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox4.Enabled = False

        DateTimePicker1.Enabled = False
        Button3.Enabled = False
        Button5.Enabled = False
        Button7.Enabled = False
        TextBox1.Enabled = False
        Button6.Enabled = False
        Button1.Enabled = False
        RadioButton1.Enabled = False
        RadioButton2.Enabled = False
    End Sub
    Public Sub limpiar()
        TextBox9.Text = ""
        TextBox12.Text = ""
        TextBox1.Text = ""

        TextBox5.Text = ""

    End Sub
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)


    End Sub
    Dim Rsr3002 As SqlDataReader

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs)

    End Sub
    Dim Rsr300 As SqlDataReader
    Private Sub PictureBox3_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.Text = "INTERNA" Then
            'TextBox8.Text = "20508740361"
            'TextBox10.Text = "CONSORCIO TEXTIL VIANNY"
        Else
            If ComboBox2.Text = "EXTERNA" Then
                'TextBox8.Text = ""
                'TextBox10.Text = ""
                TextBox8.ReadOnly = False
            End If
        End If
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Lista_ARti.Label2.Text = ComboBox3.Text
        'Lista_ARti.Label3.Text = Label13.Text
        'Lista_ARti.Show()
        If Trim(TextBox8.Text).Length = 0 Then
            MessageBox.Show("FALTA INGRESAR EL CLIENTE", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            pproductos.CCIA.Text = Label13.Text
            pproductos.ALMACEN.Text = ComboBox3.SelectedValue.ToString
            pproductos.ANO.Text = Label11.Text
            pproductos.FECHA.Text = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
            pproductos.TextBox1.Text = ComboBox3.SelectedValue.ToString
            pproductos.Label3.Text = 2
            pproductos.Show()
        End If

    End Sub

    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click
        Productos_Hilo.Label3.Text = ComboBox3.SelectedValue.ToString
        'Productos_Hilo.Label3.Text = "11"
        Productos_Hilo.Label4.Text = 200
        Productos_Hilo.Label5.Text = Label11.Text - 1
        'Productos_Hilo.Label5.Text = Label11.Text
        Productos_Hilo.Label6.Text = Label13.Text

        Select Case ComboBox3.SelectedValue.ToString
            Case "13" : Productos_Hilo.Label7.Text = "1"
            Case "05" : Productos_Hilo.Label7.Text = "1"
            Case "14" : Productos_Hilo.Label7.Text = "1"
            Case "11" : Productos_Hilo.Label7.Text = "1"

            Case Else
                Productos_Hilo.Label7.Text = "2"
        End Select
        Productos_Hilo.Show()
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub

    Dim hj2, hj1 As DataTable

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

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

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged

    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        'If Label15.Text = 1 Then


        'Else
        '    If Label15.Text = 2 Then
        '        abrir()
        '        Dim Rsr19912 As SqlDataReader
        '        Dim cco2 As String
        '        Dim sql10112 As String = "select ocomp_21e from custom_vianny.dbo.cag210e where Empr_21e ='" + Trim(Label13.Text) + "' AND almac_21e ='" + ComboBox3.Text + "' and  dest_21e ='" + ComboBox1.SelectedValue.ToString + "'"
        '        Dim cmd10112 As New SqlCommand(sql10112, conx)
        '        Rsr19912 = cmd10112.ExecuteReader()
        '        Rsr19912.Read()
        '        cco2 = Rsr19912(0)
        '        If cco2 = 1 Then
        '            TextBox15.Enabled = True
        '        Else
        '            TextBox15.Enabled = False
        '        End If
        '        Rsr19912.Close()
        '    End If

        'End If


    End Sub

    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            If TextBox5.Text = "" Then
                correlativo()
            End If
            ty2.Clear()
            abrir()
            llenar_combo_box2(ComboBox3.SelectedValue.ToString)
            Dim me12 As String
            me12 = TextBox4.Text - Month(DateTimePicker1.Value)
            DateTimePicker1.Value = DateTimePicker1.Value.AddMonths(me12)

            Dim i As Integer
            Dim ml As New vnia
            Dim mk As New fnia

            i = Len(TextBox5.Text)
            RadioButton1.Checked = True

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
            ml.gano = Label11.Text
            ml.gccia = Label13.Text
            If mk.mostrar_correlativo_nia1(ml) = True Then
                TextBox12.Enabled = False
                DataGridView1.Rows.Clear()
                Dim jk As New fnia
                Dim gu As String
                hj2 = jk.mostrar_cabecera_nia(ml)
                hj1 = jk.mostrar_detalle_nia(ml)
                If hj2.Rows.Count <> 0 Then
                    DataGridView9.DataSource = hj2

                    TextBox12.Text = Trim(DataGridView9.Rows(0).Cells(2).Value)
                    TextBox9.Text = Trim(DataGridView9.Rows(0).Cells(0).Value)
                    DateTimePicker1.Value = DataGridView9.Rows(0).Cells(1).Value
                    TextBox8.Text = Trim(DataGridView9.Rows(0).Cells(4).Value)
                    gu = DataGridView9.Rows(0).Cells(5).Value
                    TextBox10.Text = Trim(DataGridView9.Rows(0).Cells(6).Value)
                    If Trim(DataGridView9.Rows(0).Cells(8).Value) = "INT" Then
                        ComboBox2.SelectedIndex = 0
                    Else
                        ComboBox2.SelectedIndex = 1
                    End If

                    Select Case Trim(DataGridView9.Rows(0).Cells(9).Value)
                        Case "009" : ComboBox4.Text = "GUIA"
                        Case "001" : ComboBox4.Text = "FACT"
                        Case "002" : ComboBox4.Text = "OTRO"
                        Case "" : ComboBox4.Text = ""
                    End Select

                    abrir()
                    enunciado2 = New SqlCommand("select rtrim(ltrim(nomb_21e)) as nomb_21e from custom_vianny.dbo.cag210e where  Empr_21e ='" + Trim(Label13.Text) + "' AND almac_21e='" + ComboBox3.SelectedValue.ToString + "' and dest_21e='" + DataGridView9.Rows(0).Cells(5).Value + "'", conx)
                    respuesta2 = enunciado2.ExecuteReader
                    While respuesta2.Read
                        ComboBox1.Text = respuesta2.Item("nomb_21e")
                    End While
                    respuesta2.Close()
                    If Trim(DataGridView9.Rows(0).Cells(7).Value) = 1 Then
                        TextBox13.Text = ""
                        TextBox11.Text = ""
                    Else
                        TextBox11.Text = DataGridView9.Rows(0).Cells(7).Value
                        Select Case TextBox11.Text.Length
                            Case "1" : TextBox11.Text = "000" & "" & TextBox11.Text
                            Case "2" : TextBox11.Text = "00" & "" & TextBox11.Text
                            Case "3" : TextBox11.Text = "0" & "" & TextBox11.Text
                            Case "4" : TextBox11.Text = TextBox11.Text


                        End Select
                        enunciado3 = New SqlCommand(" SELECT  nomb_4 as Nombre FROM custom_vianny.dbo.Qag0400 A  Where CCIA = '" + Trim(Label13.Text) + "' and flag_4 ='1' and fase_4 ='" + TextBox11.Text + "'", conx)
                        respuesta3 = enunciado3.ExecuteReader
                        While respuesta3.Read
                            TextBox13.Text = respuesta3.Item("Nombre")
                        End While
                        respuesta3.Close()
                    End If

                    If DataGridView9.Rows(0).Cells(3).Value = 1 Then
                        RadioButton2.Checked = False
                        RadioButton1.Checked = True
                        RadioButton1.Enabled = True
                        Label12.Visible = False
                        Button5.Enabled = False
                    Else
                        RadioButton2.Checked = True
                        RadioButton1.Checked = False
                        RadioButton1.Enabled = False
                        Label12.Visible = True
                    End If


                End If

                If hj1.Rows.Count <> 0 Then
                    Dim nu1 As Integer
                    nu1 = hj1.Rows.Count
                    DataGridView1.Rows.Add(nu1)
                    For i1 = 0 To nu1 - 1
                        DataGridView10.DataSource = hj1
                        DataGridView1.Rows(i1).Cells(0).Value = DataGridView10.Rows(i1).Cells(0).Value
                        DataGridView1.Rows(i1).Cells(1).Value = Trim(DataGridView10.Rows(i1).Cells(1).Value)
                        DataGridView1.Rows(i1).Cells(2).Value = Trim(DataGridView10.Rows(i1).Cells(2).Value)
                        DataGridView1.Rows(i1).Cells(3).Value = Trim(DataGridView10.Rows(i1).Cells(3).Value)
                        DataGridView1.Rows(i1).Cells(4).Value = DataGridView10.Rows(i1).Cells(4).Value
                        DataGridView1.Rows(i1).Cells(6).Value = DataGridView10.Rows(i1).Cells(5).Value
                        DataGridView1.Rows(i1).Cells(10).Value = Trim(DataGridView10.Rows(i1).Cells(6).Value)
                        DataGridView1.Rows(i1).Cells(7).Value = Trim(DataGridView10.Rows(i1).Cells(7).Value)
                        DataGridView1.Rows(i1).Cells(8).Value = Trim(DataGridView10.Rows(i1).Cells(8).Value)
                        DataGridView1.Rows(i1).Cells(5).Value = Trim(DataGridView10.Rows(i1).Cells(9).Value)
                        DataGridView1.Rows(i1).Cells(9).Value = Trim(DataGridView10.Rows(i1).Cells(10).Value)
                        DataGridView1.Rows(i1).Cells(12).Value = DataGridView10.Rows(i1).Cells(11).Value
                        DataGridView1.Rows(i1).Cells(11).Value = Trim(DataGridView10.Rows(i1).Cells(12).Value)
                    Next
                End If

                Button3.Enabled = True
                Button5.Enabled = True
                Button7.Enabled = True
                Button6.Enabled = True
                Button6.Enabled = True
                TextBox4.Enabled = False
                TextBox5.Enabled = False
                Button2.Enabled = False
                TextBox8.ReadOnly = False

            Else

                Select Case TextBox5.Text.Length
                    Case "1" : TextBox5.Text = "00000" & "" & TextBox5.Text
                    Case "2" : TextBox5.Text = "0000" & "" & TextBox5.Text
                    Case "3" : TextBox5.Text = "000" & "" & TextBox5.Text
                    Case "4" : TextBox5.Text = "00" & "" & TextBox5.Text
                    Case "5" : TextBox5.Text = "0" & "" & TextBox5.Text
                    Case "6" : TextBox5.Text = TextBox5.Text
                End Select
                'habilitar()
                TextBox4.Enabled = False
                TextBox5.Enabled = False
                TextBox8.ReadOnly = False
                TextBox8.Enabled = True
                Button2.Enabled = False
                ComboBox2.Enabled = True
                ComboBox1.Enabled = True
                ComboBox4.Enabled = True
                ComboBox4.SelectedIndex = -1
                DateTimePicker1.Enabled = True
                TextBox9.ReadOnly = False
                TextBox9.Enabled = True
                DataGridView1.Enabled = True
                TextBox12.Enabled = False

                Button3.Enabled = True
                Button5.Enabled = True
                Button6.Enabled = True
                Button7.Enabled = True
                DataGridView1.Rows.Clear()
                TextBox9.Text = ""
                'If Label16.Text = "1" Then
                '    AlmacenOrdenCompra.Listar_Informacion()
                'End If

            End If
            Button2.Enabled = True
            DateTimePicker1.Select()
            abrir()
            Dim Rsr1991 As SqlDataReader
            Dim cco As String
            Dim sql1011 As String = "select ocomp_21e from custom_vianny.dbo.cag210e where Empr_21e ='" + Trim(Label13.Text) + "' AND almac_21e ='" + ComboBox3.SelectedValue.ToString + "' and  dest_21e ='" + ComboBox1.SelectedValue.ToString + "'"
            Dim cmd1011 As New SqlCommand(sql1011, conx)
            Rsr1991 = cmd1011.ExecuteReader()
            Rsr1991.Read()
            cco = Rsr1991(0)
            If cco = 1 Then
                TextBox15.Enabled = True
            Else
                TextBox15.Enabled = False
            End If
            Rsr1991.Close()
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.F1 Then
            If Trim(TextBox5.Text).Length > 0 Then
                Dim cli As New Clientes
                cli.TextBox3.Text = "350"
                cli.ShowDialog()
            Else
                MessageBox.Show("FALTA INGRESAR EL CORRELATIVO DEL COMPROBANTE", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        MsgBox(DataGridView1.Rows(0).Cells(10).Value)
    End Sub
    Dim Rsr3001, Rsr3 As SqlDataReader
    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        ToolTip1.SetToolTip(Button5, "Guardar Nota Ingreso")
        Dim fun As New vinsertar_nia
        Dim func As New fnia
        Dim func1 As New insertardetallenia
        Dim func2 As New fnia
        Dim con As String
        Dim hn As New vnia
        Dim fg As New fnia

        Dim Cup, Ccon, Cpartida, Cum, ui, ct1, valor As Integer
        ui = DataGridView1.Rows.Count
        For i13 = 0 To ui - 1

            If Trim(DataGridView1.Rows(i13).Cells(4).Value).Length = 0 Then
                Cup = Cup + 1
            End If
            If Trim(DataGridView1.Rows(i13).Cells(5).Value).Length = 0 Then
                Cum = Cum + 1
            End If
            If Trim(DataGridView1.Rows(i13).Cells(6).Value).Length = 0 Then
                Ccon = Ccon + 1
            End If
            If Trim(DataGridView1.Rows(i13).Cells(8).Value).Length = 0 Then
                Cpartida = Cpartida + 1
            End If
            If Trim(DataGridView1.Rows(i13).Cells(11).Value).Length = 0 Then
                ct1 = ct1 + 1
            End If
        Next

        Dim sql102213 As String = "select ocomp_21e from custom_vianny.dbo.cag210e where dest_21e ='" + ComboBox1.SelectedValue.ToString + "' and Empr_21e ='" + Label13.Text + "' and almac_21e ='" + Trim(ComboBox3.SelectedValue.ToString) + "'"
        Dim cmd102213 As New SqlCommand(sql102213, conx)
        Rsr3001 = cmd102213.ExecuteReader()
        Dim jo As Integer
        If Rsr3001.Read() Then
            jo = Rsr3001(0)
        End If
        Rsr3001.Close()

        Dim sql1 As String = "select COUNT(ncom_3a) from custom_vianny.dbo.map030f where CCIA_3A ='" + Label13.Text + "' and CPERIODO_3A ='" + Label11.Text + "' and talm_3a ='" + Trim(ComboBox3.SelectedValue.ToString) + "' and ccom_3a ='1' and ncom_3a ='" + TextBox4.Text & TextBox5.Text + "'"
        Dim cmd1 As New SqlCommand(sql1, conx)
        Rsr3 = cmd1.ExecuteReader()

        If Rsr3.Read() Then
            valor = Convert.ToInt32(Rsr3(0))
        End If
        Rsr3.Close()

        If jo = 1 And (Trim(TextBox12.Text).Length = 0) Then
            MessageBox.Show("EL MOTIVO COMPRA ES OBLIGATORIO INGRESAR  REGISTRO DE GUIA REMISION O FACTURA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else

            If jo = 1 And ct1 > 0 Then
                MessageBox.Show("EL MOTIVO COMPRA ES OBLIGATORIO INGRESAR LA OC EN CADA ITEMS", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else

                If Cum > 0 Then
                    MessageBox.Show("FALTA AGREGAR LA UM DE PRESENTACION EN ALGUNA CELDA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else

                    If Cup > 0 Then
                        MessageBox.Show("FALTA INGRESAR LA CANTIDAD DE PRESENTACION EN ALGUNA CELDA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Else
                        If Ccon > 0 Then
                            MessageBox.Show("FALTA INGRESAR LA CANTIDAD EN KG/MTS EN ALGUNA CELDA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)

                        Else
                            If Cpartida > 0 And (Trim(ComboBox3.SelectedValue.ToString) = "03" Or Trim(ComboBox3.SelectedValue.ToString) = "59") Then
                                MessageBox.Show("FALTA INGRESAR EL LOTE EN ALGUNA CELDA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Else
                                Dim cor1, correlativo As String
                                If valor > 0 And Label15.Text = "0" Then
                                    Dim func14 As New fnia
                                    Dim dt4 As New vnia
                                    Dim cor As Integer

                                    TextBox4.Text = DateTime.Now.ToString("MM")
                                    dt4.gccia = Label13.Text
                                    dt4.gmes = Me.TextBox4.Text
                                    dt4.galmacen = ComboBox3.SelectedValue.ToString
                                    dt4.gano = Label11.Text

                                    cor = func14.buscar_nia(dt4) + 1
                                    Select Case cor.ToString().Length

                                        Case "1" : cor1 = "00000" & "" & cor
                                        Case "2" : cor1 = "0000" & "" & cor
                                        Case "3" : cor1 = "000" & "" & cor
                                        Case "4" : cor1 = "00" & "" & cor
                                        Case "5" : cor1 = "0" & "" & cor
                                        Case "6" : cor1 = cor
                                    End Select
                                    correlativo = TextBox4.Text & cor1
                                    MessageBox.Show("EL NUMERO DE REGISTRO YA ESTA EN LA BASE DE DATOS, SE ACTUALIZARA AL SIGUIENTE  -- " + correlativo, "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                Else
                                    correlativo = TextBox4.Text & TextBox5.Text
                                End If

                                Dim x, mj, valo, mp, mop, final As String
                                    Dim pp As Integer

                                    x = Convert.ToString(TextBox12.Text.IndexOf("-", 0) + 1)
                                    If x > 0 Then
                                        valo = Trim(TextBox12.Text)
                                        Select Case x

                                            Case "1" : mj = "0000"
                                            Case "2" : mj = "000" & TextBox12.Text.Substring(0, x - 1)
                                            Case "3" : mj = "00" & TextBox12.Text.Substring(0, x - 1)
                                            Case "4" : mj = "0" & TextBox12.Text.Substring(0, x - 1)
                                            Case "5" : mj = TextBox12.Text.Substring(0, x - 1)

                                        End Select


                                        pp = Convert.ToInt32(TextBox12.Text.Length) - Convert.ToInt32(x)

                                        mp = TextBox12.Text.Substring(x, pp)
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
                                        final = TextBox12.Text
                                    End If


                                    con = TextBox4.Text & TextBox5.Text
                                    hn.gcomprobante = con
                                    hn.galmacen = ComboBox3.SelectedValue.ToString
                                    hn.gncom = 1
                                    hn.gano = Label11.Text
                                    hn.gccia = Label13.Text
                                    fg.eliminar_nia(hn)


                                    fun.gncom = con
                                    fun.gglosa = Trim(TextBox9.Text)
                                    fun.gfecha = DateTimePicker1.Value
                                    fun.gguia = final
                                    fun.gano = Label11.Text
                                    Select Case ComboBox4.Text
                                        Case "GUIA" : fun.gdoc = "009"
                                        Case "FACT" : fun.gdoc = "001"
                                        Case "OTRO" : fun.gdoc = "002"
                                        Case "" : fun.gdoc = ""
                                    End Select
                                    fun.galmacen = ComboBox3.SelectedValue.ToString
                                    fun.gempresa = Trim(TextBox8.Text)
                                    If ComboBox2.Text = "INTERNA" Then
                                        fun.gint = "INT"
                                    Else
                                        fun.gint = "EXT"
                                    End If

                                    fun.gorigen = ComboBox1.SelectedValue.ToString
                                    Select Case ComboBox4.Text
                                        Case "GUIA" : fun.gtidoc = "009"
                                        Case "FACT" : fun.gtidoc = "001"
                                        Case "OTRO" : fun.gtidoc = "002"
                                        Case "" : fun.gtidoc = ""
                                    End Select
                                    fun.gusuario = Trim(TextBox7.Text)
                                    fun.gfase = Trim(TextBox11.Text)
                                    fun.gccia = Label13.Text
                                    Dim i, num2 As Integer
                                    'For i = 1 To 29
                                    '    If DataGridView1.Rows(i).Cells(0).Value < DataGridView1.Rows(i - 1).Cells(0).Value Then
                                    '        num2 = DataGridView1.Rows(i - 1).Cells(0).Value + 1
                                    '        i17 = i

                                    '    End If
                                    'Next

                                    If func.insertar_cabecera_nia(fun) Then
                                        num2 = DataGridView1.Rows.Count
                                        DataGridView1.Rows(num2 - 1).Cells(11).Selected = True

                                        For i = 0 To num2 - 1
                                            Dim nu As String
                                            func1.gncom = con

                                            nu = DataGridView1.Rows(i).Cells(0).Value

                                            Select Case nu.Length
                                                Case "1" : nu = "00" & "" & nu
                                                Case "2" : nu = "0" & "" & nu
                                                Case "3" : nu = nu
                                            End Select
                                            func1.gitem = nu

                                            func1.glinea = Trim(DataGridView1.Rows(i).Cells(2).Value)
                                            If Trim(DataGridView1.Rows(i).Cells(1).Value) = "" Then
                                                func1.gop = ""
                                            Else
                                                func1.gop = Trim(DataGridView1.Rows(i).Cells(1).Value)
                                            End If

                                            func1.gund = Trim(DataGridView1.Rows(i).Cells(7).Value)
                                            func1.gcantidad = DataGridView1.Rows(i).Cells(4).Value
                                            func1.gcantidad1 = DataGridView1.Rows(i).Cells(6).Value
                                            If Trim(Convert.ToString(DataGridView1.Rows(i).Cells("Partida").Value)) = "" Then
                                                func1.gpartida = ""
                                            Else
                                                func1.gpartida = Trim(Convert.ToString(DataGridView1.Rows(i).Cells("Partida").Value))
                                            End If

                                            func1.galmacen = ComboBox3.SelectedValue.ToString

                                            If Trim(Convert.ToString(DataGridView1.Rows(i).Cells(5).Value)) = "" Then
                                                func1.gumpresentacion = ""
                                            Else
                                                func1.gumpresentacion = Trim(DataGridView1.Rows(i).Cells(5).Value)
                                            End If
                                            If Trim(Convert.ToString(DataGridView1.Rows(i).Cells(9).Value)) = "" Then
                                                func1.gotejeduria = ""
                                            Else
                                                func1.gotejeduria = Trim(DataGridView1.Rows(i).Cells(9).Value)
                                            End If


                                            If Trim(DataGridView1.Rows(i).Cells(11).Value) = "" Then
                                                func1.goc = ""
                                            Else
                                                func1.goc = Trim(DataGridView1.Rows(i).Cells(11).Value)
                                            End If
                                            func1.gano = Label11.Text
                                            If Trim(DataGridView1.Rows(i).Cells(8).Value) = "" Then
                                                func1.glote = ""
                                            Else
                                                func1.glote = Trim(DataGridView1.Rows(i).Cells(8).Value)
                                            End If
                                            If DataGridView1.Rows(i).Cells(12).Value = "0.00" Or Trim(DataGridView1.Rows(i).Cells(12).Value) = "" Then
                                                func1.gcenvio = "0.00"
                                            Else
                                                func1.gcenvio = DataGridView1.Rows(i).Cells(12).Value
                                            End If
                                            func1.gccia = Label13.Text
                                            If Trim(ComboBox3.SelectedValue.ToString) = "59" Then
                                                func1.gclasif = "1"
                                            Else
                                                func1.gclasif = ""
                                            End If

                                            func2.insertar_detalle_nia(func1)
                                            'actualizar informacion de la oc
                                            Dim cmd20 As New SqlCommand("update custom_vianny.dbo.lap0300 set percep_3a='1' where ccia_3a =@ccia_3a and cperiodo_3a =@cperiodo_3a and ncom_3a =@ncom_3a and item_3a =@item_3a", conx)
                                            cmd20.Parameters.AddWithValue("@ccia_3a", (Label13.Text.ToString.Trim))
                                            cmd20.Parameters.AddWithValue("@cperiodo_3a", (Label11.Text.ToString.Trim))
                                            cmd20.Parameters.AddWithValue("@ncom_3a", (DataGridView1.Rows(i).Cells(11).Value.ToString.Trim))
                                            cmd20.Parameters.AddWithValue("@item_3a", (DataGridView1.Rows(i).Cells(0).Value.ToString.Trim))
                                            cmd20.ExecuteNonQuery()
                                        Next
                                        abrir()
                                    Dim sql1022 As String = "select COUNT(item_3a),ISNULL( SUM(CAST(ISNULL(percep_3a, 0) AS INT)),0) from custom_vianny.dbo.lap0300 where  ncom_3a ='" + DataGridView1.Rows(0).Cells(11).Value.ToString.Trim + "' and ccia_3a='" + Label13.Text.ToString.Trim + "'"
                                    Dim cmd1022 As New SqlCommand(sql1022, conx)
                                        Rsr25 = cmd1022.ExecuteReader()
                                        Dim i5 As Integer
                                        i5 = 0
                                        If Rsr25.Read() Then
                                        If Convert.ToInt32(Rsr25(0)) = Convert.ToInt32(Rsr25(1)) Then
                                            Rsr25.Close()
                                            Dim cmd20 As New SqlCommand("UPDATE custom_vianny.dbo.lag0300 set termina_3 ='1' where  ncom_3 =@ncom_3a and ccia_3=@ccia_3a", conx)
                                            cmd20.Parameters.AddWithValue("@ccia_3a", (Label13.Text.ToString.Trim))
                                            cmd20.Parameters.AddWithValue("@ncom_3a", (DataGridView1.Rows(0).Cells(11).Value.ToString.Trim))
                                            cmd20.ExecuteNonQuery()
                                        End If
                                    End If
                                        Rsr25.Close()
                                        If Label16.Text = "1" Then
                                            AlmacenOrdenCompra.Listar_Informacion()
                                        End If
                                        MessageBox.Show("Datos registrados correctamente ", "Guardando registros", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                        Dim respuesta As DialogResult
                                        Dim hj2 As New flog
                                        Dim hj1 As New vlog
                                        hj1.gmodulo = "NIA-ALMACEN"
                                        hj1.gcuser = MDIParent1.Label3.Text
                                        If Label14.Text = "0" Then
                                            hj1.gaccion = 1
                                        Else
                                            hj1.gaccion = 2
                                        End If

                                        hj1.gpc = My.Computer.Name
                                        hj1.gfecha = DateTimePicker2.Value
                                        hj1.gdetalle = Trim(ComboBox3.SelectedValue.ToString) & Trim(TextBox4.Text) & Trim(TextBox5.Text)
                                        hj1.gccia = Label13.Text
                                        hj2.insertar_log(hj1)
                                        respuesta = MessageBox.Show("DESEA IMPRIMIR LA NOTA DE INGRESO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                                        If (respuesta = Windows.Forms.DialogResult.Yes) Then

                                            Reporte_Nia_Nsa.TextBox1.Text = ComboBox3.SelectedValue.ToString
                                            Reporte_Nia_Nsa.TextBox2.Text = 1
                                            Reporte_Nia_Nsa.TextBox3.Text = TextBox4.Text & "" & TextBox5.Text
                                            Reporte_Nia_Nsa.TextBox4.Text = Label11.Text
                                            Reporte_Nia_Nsa.TextBox5.Text = Label13.Text
                                            Reporte_Nia_Nsa.Show()
                                        End If
                                        limpiar()


                                    End If
                                    limpiar()
                                    TextBox4.Enabled = True
                                    TextBox5.Enabled = False
                                    TextBox9.Enabled = False
                                    TextBox12.Enabled = False

                                    DateTimePicker1.Enabled = False
                                    DataGridView1.Enabled = False
                                    ComboBox4.SelectedIndex = -1
                                    DataGridView1.Rows.Clear()
                                    TextBox5.Select()
                                    Dim func12 As New fnia
                                    Dim dts As New vnia
                                    TextBox4.Text = DateTime.Now.ToString("MM")

                                    dts.gmes = Me.TextBox4.Text
                                    dts.galmacen = ComboBox3.SelectedValue.ToString
                                    dts.gano = Label11.Text
                                    dts.gccia = Label13.Text
                                    Me.TextBox5.Text = func12.buscar_nia(dts) + 1
                                    Select Case TextBox5.Text.Length

                                        Case "1" : TextBox5.Text = "00000" & "" & TextBox5.Text
                                        Case "2" : TextBox5.Text = "0000" & "" & TextBox5.Text
                                        Case "3" : TextBox5.Text = "000" & "" & TextBox5.Text
                                        Case "4" : TextBox5.Text = "00" & "" & TextBox5.Text
                                        Case "5" : TextBox5.Text = "0" & "" & TextBox5.Text
                                        Case "6" : TextBox5.Text = TextBox5.Text
                                    End Select
                                    TextBox5.Enabled = True
                                    TextBox5.ReadOnly = False
                                    RadioButton1.Checked = True
                                    Label16.Text = "0"
                                    TextBox5.Select()
                                End If
                            End If
                        End If
                    End If
                End If
            End If

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        TextBox12.Enabled = True
        TextBox12.ReadOnly = False
    End Sub

    Private Sub TextBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox11.KeyDown
        If e.KeyCode = Keys.F1 Then
            Procesos.Label1.Text = "1"
            Procesos.Label2.Text = Label13.Text
            Procesos.Show()
        Else
            If e.KeyCode = Keys.Enter Then
                TextBox8.Select()
            End If
        End If
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

        If Trim(TextBox4.Text) <> Month(DateTimePicker1.Value) Then
            MsgBox("LA FECHA DE EM0ISION NO CORRESPONDE AL MES DE REGISTRO")
            Dim me12 As String
            me12 = TextBox4.Text - Month(DateTimePicker1.Value)
            DateTimePicker1.Value = DateTimePicker1.Value.AddMonths(me12)
        End If
    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox11.Select()
        End If
    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged

    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        'DataGridView1.BeginEdit(True)
    End Sub

    Private Sub TextBox5_MouseDown(sender As Object, e As MouseEventArgs) Handles TextBox5.MouseDown

    End Sub

    Private Sub ComboBox1_Click(sender As Object, e As EventArgs) Handles ComboBox1.Click
        Label15.Text = 2
    End Sub

    Private Sub TextBox15_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox15.KeyDown
        If e.KeyCode = Keys.Enter Then
            If TextBox15.Text.Length < 8 Then
                MsgBox("LA ORDEN DE COMPRA NO TIENE 8 DIGITOS")
            Else
                DataGridView1.Rows.Clear()
                abrir()
                Dim Rsr1991 As SqlDataReader
                Dim cco1, crq As String
                If CheckBox1.Checked = True Then
                    Dim sql1011 As String = "select aprobado_3,fich_3 from custom_vianny.dbo.Lag0300  WHERE CCia_3 = '" + Label13.Text + "' AND NCom_3 = '" + TextBox15.Text + "'"
                    Dim cmd1011 As New SqlCommand(sql1011, conx)
                    Rsr1991 = cmd1011.ExecuteReader()
                    Rsr1991.Read()
                    cco1 = Rsr1991(0)
                    crq = Rsr1991(1)
                Else
                    If CheckBox2.Checked = True Then
                        Dim sql1011 As String = " select aprobado_4,fich_4 from custom_vianny.dbo.lag0400  WHERE ccia_4 = '" + Label13.Text + "' AND ncom_4 = '" + TextBox15.Text + "'"
                        Dim cmd1011 As New SqlCommand(sql1011, conx)
                        Rsr1991 = cmd1011.ExecuteReader()
                        Rsr1991.Read()
                        cco1 = Rsr1991(0)
                        crq = Rsr1991(1)
                    End If
                End If





                If cco1 = True Then
                    If Trim(crq) = Trim(TextBox8.Text) Then
                        Rsr1991.Close()
                        abrir()
                        If CheckBox1.Checked = True Then
                            Dim Rsr19918 As SqlDataReader
                            Dim sql10118 As String = "select ncom_3a,item_3a,pedido_3a, linea_3a,c.nomb_17, cante_3a,unide2_3a,cant_3a,c.unid_17 from custom_vianny.dbo.LAp0300 lp left join custom_vianny.dbo.cag1700 c on lp.linea_3a=c.linea_17 and lp.ccia_3a=c.ccia where ncom_3a ='" + TextBox15.Text + "' and ccia_3a ='" + Label13.Text + "'"
                            Dim cmd10118 As New SqlCommand(sql10118, conx)
                            Rsr19918 = cmd10118.ExecuteReader()
                            Dim i5 As Integer
                            i5 = DataGridView1.Rows.Count

                            While Rsr19918.Read()
                                DataGridView1.Rows.Add()
                                DataGridView1.Rows(i5).Cells(0).Value = i5 + 1
                                DataGridView1.Rows(i5).Cells(1).Value = Rsr19918(2)
                                DataGridView1.Rows(i5).Cells(2).Value = Rsr19918(3)
                                DataGridView1.Rows(i5).Cells(3).Value = Rsr19918(4)
                                DataGridView1.Rows(i5).Cells(4).Value = Rsr19918(5)
                                DataGridView1.Rows(i5).Cells(5).Value = Rsr19918(6)
                                DataGridView1.Rows(i5).Cells(6).Value = Rsr19918(7)
                                DataGridView1.Rows(i5).Cells(7).Value = Rsr19918(8)
                                DataGridView1.Rows(i5).Cells(11).Value = Rsr19918(0)
                                i5 = i5 + 1
                                'DataGridView2.Rows(i5).Cells(6).Value = (Rsr2(12) * DataGridView1.Rows(0).Cells(10).Value) / 100
                            End While
                            Rsr19918.Close()
                            TextBox15.Text = ""
                        Else
                            Dim Rsr19918 As SqlDataReader
                            Dim sql10118 As String = "select ncom_4a,item_4a,pedido_4a, linea_4a,c.nomb_17, cant_4a,unid_17,cant_4a,c.unid_17 from custom_vianny.dbo.lap0400 lp left join custom_vianny.dbo.cag1700 c on lp.linea_4a=c.linea_17 and lp.ccia_4a=c.ccia where ncom_4a ='" + TextBox15.Text + "' and ccia_4a ='" + Label13.Text + "'"
                            Dim cmd10118 As New SqlCommand(sql10118, conx)
                            Rsr19918 = cmd10118.ExecuteReader()
                            Dim i5 As Integer
                            i5 = DataGridView1.Rows.Count

                            While Rsr19918.Read()
                                DataGridView1.Rows.Add()
                                DataGridView1.Rows(i5).Cells(0).Value = i5 + 1
                                DataGridView1.Rows(i5).Cells(1).Value = Rsr19918(2)
                                DataGridView1.Rows(i5).Cells(2).Value = Rsr19918(3)
                                DataGridView1.Rows(i5).Cells(3).Value = Rsr19918(4)
                                DataGridView1.Rows(i5).Cells(4).Value = Rsr19918(5)
                                DataGridView1.Rows(i5).Cells(5).Value = Rsr19918(6)
                                DataGridView1.Rows(i5).Cells(6).Value = Rsr19918(7)
                                DataGridView1.Rows(i5).Cells(7).Value = Rsr19918(8)
                                DataGridView1.Rows(i5).Cells(11).Value = Rsr19918(0)
                                i5 = i5 + 1
                                'DataGridView2.Rows(i5).Cells(6).Value = (Rsr2(12) * DataGridView1.Rows(0).Cells(10).Value) / 100
                            End While
                            Rsr19918.Close()
                            TextBox15.Text = ""
                        End If

                    Else
                        MsgBox("LA ORDEN DE COMPRA O SERVICIO NO PERTENECE AL PROVEEDOR")
                    End If

                Else
                    MsgBox("LA ORDEN DE COMPRA O SERVICIO NO SE PUEDE AGREGAR PORQUE NO ESTA APROBADA")
                End If
            End If

        End If
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        NumConFrac(TextBox4, e)
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            CheckBox2.Checked = False
            TextBox15.Enabled = True
        Else
            CheckBox2.Checked = True
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ToolTip1.SetToolTip(Button5, "Editar Nota Ingreso")
        If RadioButton2.Checked = True Then
            MsgBox("NO SE PUEDE EDITAR INFORMACION EN UNA NOTA DE INGRESO ANULADA")
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

                    TextBox4.Enabled = False

                    TextBox5.Enabled = False
                    Button2.Enabled = True
                    Button3.Enabled = True
                    ComboBox1.Enabled = True
                    ComboBox2.Enabled = True
                    Label14.Text = "1"
                    ComboBox4.Enabled = True
                Else
                    MsgBox("DIA " & DateTimePicker1.Value & " CERRADO,NO SE PUEDE EDITAR LA INFORMACION - CONSULTE CON EL AREA CONTABLE ")
                End If


            End If
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ToolTip1.SetToolTip(Button5, "Anular Nota Ingreso")
        'System.Diagnostics.Process.Start("www.google.com")
        Dim sql102213 As String = "SELECT cierre_3 FROM custom_vianny.dbo.Cie0300 A  Where a.ccia ='" + Label13.Text + "' AND A.talm_3 ='" + ComboBox3.SelectedValue.ToString + "' AND A.fcie_3='" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "'"
        Dim cmd102213 As New SqlCommand(sql102213, conx)
        Rsr3002 = cmd102213.ExecuteReader()
        Dim jo As Integer
        If Rsr3002.Read() Then
            jo = Rsr3002(0)
        Else
            jo = 0
        End If
        Rsr3002.Close()
        Dim respuesta As DialogResult
        Dim ml As New vnia
        Dim mk As New fnia
        ml.gcomprobante = TextBox4.Text & TextBox5.Text
        ml.galmacen = ComboBox3.SelectedValue.ToString
        ml.gano = Label11.Text
        ml.gccia = Label13.Text
        respuesta = MessageBox.Show("QUIERES ANULAR NOTA DE INGRESO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            If jo = 0 Then
                Label12.Visible = True
                RadioButton2.Checked = True
                RadioButton1.Checked = False
                mk.anular_nia(ml)
                TextBox4.Enabled = True
                TextBox4.Select()
                DateTimePicker1.Enabled = False
                TextBox9.Enabled = False
                TextBox12.Enabled = False
                TextBox5.Enabled = True

                DataGridView1.Enabled = False
                Button3.Enabled = False
                Button6.Enabled = False
                Button5.Enabled = False

                Dim hj2 As New flog
                Dim hj1 As New vlog
                hj1.gmodulo = "NIA-ALMACEN"
                hj1.gcuser = MDIParent1.Label3.Text

                hj1.gaccion = 3

                hj1.gpc = My.Computer.Name
                hj1.gfecha = DateTimePicker2.Value
                hj1.gdetalle = Trim(ComboBox3.SelectedValue.ToString) & Trim(TextBox4.Text) & Trim(TextBox5.Text)
                hj1.gccia = Label13.Text
                hj2.insertar_log(hj1)
                Label12.Visible = False
                MsgBox("SE ANULO LA INFORMACION CORRECTAMENTE")
                limpiar()
                TextBox4.Enabled = True
                TextBox5.Enabled = False
                TextBox9.Enabled = False
                TextBox12.Enabled = False

                DateTimePicker1.Enabled = False
                DataGridView1.Enabled = False
                ComboBox4.SelectedIndex = -1
                DataGridView1.Rows.Clear()
                TextBox5.Select()
                Dim func12 As New fnia
                Dim dts As New vnia
                TextBox4.Text = DateTime.Now.ToString("MM")

                dts.gmes = Me.TextBox4.Text
                dts.galmacen = ComboBox3.SelectedValue.ToString
                dts.gano = Label11.Text
                dts.gccia = Label13.Text
                Me.TextBox5.Text = func12.buscar_nia(dts) + 1
                Select Case TextBox5.Text.Length

                    Case "1" : TextBox5.Text = "00000" & "" & TextBox5.Text
                    Case "2" : TextBox5.Text = "0000" & "" & TextBox5.Text
                    Case "3" : TextBox5.Text = "000" & "" & TextBox5.Text
                    Case "4" : TextBox5.Text = "00" & "" & TextBox5.Text
                    Case "5" : TextBox5.Text = "0" & "" & TextBox5.Text
                    Case "6" : TextBox5.Text = TextBox5.Text
                End Select
                TextBox5.Enabled = True
                TextBox5.ReadOnly = False
                RadioButton1.Checked = True
                TextBox5.Select()
            Else
                MsgBox("DIA " & DateTimePicker1.Value & " CERRADO,NO SE PUEDE EDITAR LA INFORMACION - CONSULTE CON EL AREA CONTABLE ")
            End If

        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ToolTip1.SetToolTip(Button5, "Imprimir Nota Ingreso")
        Reporte_Nia_Nsa.TextBox1.Text = ComboBox3.SelectedValue.ToString
        Reporte_Nia_Nsa.TextBox2.Text = 1
        Reporte_Nia_Nsa.TextBox3.Text = TextBox4.Text & "" & TextBox5.Text
        Reporte_Nia_Nsa.TextBox4.Text = Label11.Text
        Reporte_Nia_Nsa.TextBox5.Text = Label13.Text
        Reporte_Nia_Nsa.Show()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        ToolTip1.SetToolTip(Button5, "Retroceder Nota Ingreso")
        Label15.Text = 1
        bloqueado()
        limpiar()
        TextBox8.Enabled = False
        TextBox5.ReadOnly = False
        TextBox4.ReadOnly = False
        TextBox5.Enabled = True
        TextBox4.Enabled = True
        TextBox4.Select()
        ComboBox3.Enabled = True
        DataGridView1.Rows.Clear()
        ty2.Clear()
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            CheckBox1.Checked = False
            TextBox15.Enabled = True
        Else
            CheckBox1.Checked = True
        End If
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        NumConFrac(TextBox5, e)
    End Sub
    Dim Rsr34, Rsr35 As SqlDataReader
    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit

        Dim valor, valor1, total As Integer
        valor = 0
        valor1 = 0
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "OC/OS" Then
            Dim sql1022134 As String = "select COUNT(ncom_3a) from custom_vianny.DBO.lap0300 where ncom_3a ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(11).Value) + "' and ccia_3a ='" + Label13.Text + "' and linea_3a ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value) + "'"
            Dim cmd1022134 As New SqlCommand(sql1022134, conx)
            Rsr34 = cmd1022134.ExecuteReader()
            If Rsr34.Read() = True Then
                valor = Rsr34(0)
            Else
                valor = 0
                'If Rsr34(0) > 0 Then
                '    DataGridView1.Rows(e.RowIndex).Cells(11).Value = ""
                'Else
                '    MsgBox("EL ARTICULO NO PERTENECE A LA ORDEN INGRESADA, REVISE SU ORDEN DE COMPRA O SERVICIO")
                'End If
            End If
            Rsr34.Close()
            Dim sql10221348 As String = "select COUNT(ncom_4a) from custom_vianny.DBO.lap0400 where ncom_4a ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(11).Value) + "' and ccia_4a ='" + Label13.Text + "' and linea_4a ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value) + "'"
            Dim cmd10221348 As New SqlCommand(sql10221348, conx)
            Rsr35 = cmd10221348.ExecuteReader()
            If Rsr35.Read = True Then
                valor1 = Rsr35(0)
            Else
                valor1 = 0
            End If
            Rsr35.Close()

            total = valor + valor1

            If total > 0 Then

            Else
                DataGridView1.Rows(e.RowIndex).Cells(11).Value = ""
                MsgBox("EL ARTICULO NO PERTENECE A LA ORDEN INGRESADA, REVISE SU ORDEN DE COMPRA O SERVICIO")
            End If
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Lote" Then
            If Trim(ComboBox3.SelectedValue.ToString) = "59" Then
                DataGridView1.Rows(e.RowIndex).Cells(10).Value = DataGridView1.Rows(e.RowIndex).Cells(8).Value
            End If
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Partida" Then
            If Trim(ComboBox3.SelectedValue.ToString) = "59" Then
                DataGridView1.Rows(e.RowIndex).Cells(8).Value = DataGridView1.Rows(e.RowIndex).Cells(10).Value
            End If
        End If
    End Sub

    Private Sub TextBox15_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox15.KeyPress

    End Sub
    Public Sub RecibirDatos(ByVal filas As List(Of DataGridViewRow))
        ' Iteramos sobre las filas recibidas
        For Each fila In filas
            ' Creamos una nueva fila en el DataGridView de NotaIngreso
            Dim index As Integer = DataGridView1.Rows.Add()

            ' Transferimos los valores de las columnas
            DataGridView1.Rows(index).Cells(0).Value = fila.Cells("ITEMS").Value
            DataGridView1.Rows(index).Cells(1).Value = fila.Cells("PEDIDO").Value
            DataGridView1.Rows(index).Cells(2).Value = fila.Cells("CODIGO").Value
            DataGridView1.Rows(index).Cells(3).Value = fila.Cells("PRODUCTO").Value
            DataGridView1.Rows(index).Cells(4).Value = fila.Cells("CANT PRES").Value
            DataGridView1.Rows(index).Cells(6).Value = fila.Cells("CANT").Value
            DataGridView1.Rows(index).Cells(7).Value = fila.Cells("UND").Value
            DataGridView1.Rows(index).Cells(11).Value = fila.Cells("OC").Value
            DataGridView1.Rows(index).Cells(13).Value = fila.Cells("valor").Value
            ' Continúa para todas las columnas que desees transferir
        Next
    End Sub
    Sub alamcenes(ByVal almacen As String)

        Try

            conn = New SqlDataAdapter("select talm_15m + '  |  '+ nomb_15m as nombre, talm_15m as almacen from custom_vianny.dbo.Mag1500  where ccia ='" + Trim(Label13.Text) + "' and  flag_15m ='1'  and talm_15m ='" + almacen + "'", conx)
            conn.Fill(ty32)
            ComboBox3.DataSource = ty32
            ComboBox3.DisplayMember = "nombre"
            ComboBox3.ValueMember = "almacen"
            'respuesta = enunciado.ExecuteReader
            'While respuesta.Read
            '    cb.Items.Add(respuesta.Item("nomb_17f"))
            'End While
            'respuesta.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub LLenarAlmacenAviosSuministros(ByVal ser As String,ByVal almacen As String)

        Try

            conn = New SqlDataAdapter("select talm_15m + '  |  '+ nomb_15m as nombre, talm_15m as almacen from custom_vianny.dbo.Mag1500  where ccia ='" + Trim(Label13.Text) + "' and  flag_15m ='1' and  (seriencr_15m ='" + ser + "' or seriencr_15m ='1_2' ) and talm_15m ='" + almacen + "'", conx)
            conn.Fill(ty33)
            ComboBox3.DataSource = ty33
            ComboBox3.DisplayMember = "nombre"
            ComboBox3.ValueMember = "almacen"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub compras()
        DataGridView1.Enabled = True
        Button3.Enabled = True
        Button7.Enabled = True
        Button8.Enabled = True
        Button2.Enabled = True
        Button4.Enabled = True
        Button5.Enabled = False
        Button6.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        ComboBox4.Enabled = True
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = False
        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
    End Sub
    Public Sub EstablecerAlmacen(almacen As String)
        ' Verificar si el valor existe en el ComboBox antes de asignarlo
        alamcenes(almacen)
    End Sub
End Class