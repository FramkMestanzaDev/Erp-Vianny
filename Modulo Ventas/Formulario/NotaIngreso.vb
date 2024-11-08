Imports System.Data.SqlClient
Imports System.Data.Sql
Public Class NotaIngreso
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox16.Enabled = False
        TextBox17.Enabled = False
        TextBox4.Enabled = True

        TextBox4.ReadOnly = False
        TextBox5.ReadOnly = False
        TextBox4.Select()
    End Sub
    Private Sub TextBox16_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox16.KeyDown
        If e.KeyCode = Keys.Enter Then
            'Select Case TextBox16.Text.Length
            '    Case "1" : TextBox16.Text = "000" & "" & TextBox16.Text
            '    Case "2" : TextBox16.Text = "00" & "" & TextBox16.Text
            '    Case "3" : TextBox16.Text = "0" & "" & TextBox16.Text
            '    Case "4" : TextBox16.Text = TextBox16.Text
            'End Select
            TextBox17.Select()
        End If
    End Sub
    Dim dt2 As DataTable
    Private Sub TextBox17_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox17.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim fn As New fnia
            Dim valor As String
            Dim fn1 As New vnia

            Dim a9, w1
            Dim cant9, sum9 As Double
            w1 = DataGridView1.Rows.Count
            Select Case TextBox17.Text.Length
                Case "1" : TextBox17.Text = "0000000" & "" & TextBox17.Text
                Case "2" : TextBox17.Text = "000000" & "" & TextBox17.Text
                Case "3" : TextBox17.Text = "00000" & "" & TextBox17.Text
                Case "4" : TextBox17.Text = "0000" & "" & TextBox17.Text
                Case "5" : TextBox17.Text = "000" & "" & TextBox17.Text
                Case "6" : TextBox17.Text = "00" & "" & TextBox17.Text
                Case "6" : TextBox17.Text = "0" & "" & TextBox17.Text
                Case "7" : TextBox17.Text = TextBox17.Text
            End Select

            'ComboBox4.SelectedIndex = 0
            ComboBox4.Enabled = True

            'fn1.gano = DateTime.Now.ToString("yyyy")
            'fn1.gccia = Label14.Text
            'fn1.gguia = TextBox16.Text & "-" & TextBox17.Text
            'fn1.galmacen = ComboBox1.Text
            'valor = fn.buscar_guia(fn1)

            'If valor = True Then
            '    MsgBox("LA GUIA YA ESTA REGISTRADA")
            '    TextBox17.Text = ""
            '    TextBox12.Text = ""
            '    TextBox1.Text = ""
            '    DataGridView1.Rows.Clear()
            '    DataGridView1.Rows.Add(15)
            '    DataGridView1.Columns(0).Width = 37
            '    DataGridView1.Columns(1).Width = 85
            '    DataGridView1.Columns(2).Width = 140
            '    DataGridView1.Columns(3).Width = 300
            '    DataGridView1.Columns(4).Width = 60
            '    DataGridView1.Columns(5).Width = 85

            '    DataGridView1.Columns(0).DefaultCellStyle.BackColor = Color.Silver
            '    DataGridView1.Columns(2).DefaultCellStyle.BackColor = Color.Silver
            'Else



            'TextBox4.Text = DateTime.Now.ToString("MM")




            'mostrar()
            'mostrar1()
            'mostrar2()
            'habilitar()
            Dim ml As New vnia
            Dim mk As New fnsa
            Dim jk As New fnsa
            ml.gcomprobante = Trim(TextBox17.Text)
            ml.galmacen1 = Trim(TextBox16.Text)
            ml.gano = Label11.Text
            ml.gccia = Label14.Text
            dt2 = jk.mostrar_detalle_nsa(ml)
            If dt2.Rows.Count <> 0 Then
                Dim nu1 As Integer
                DataGridView12.DataSource = dt2
                nu1 = DataGridView12.Rows.Count
                DataGridView1.Rows.Add(nu1)
                For i1 = 0 To nu1 - 1

                    DataGridView1.Rows(i1).Cells(0).Value = DataGridView12.Rows(i1).Cells(0).Value
                    DataGridView1.Rows(i1).Cells(1).Value = DataGridView12.Rows(i1).Cells(1).Value
                    DataGridView1.Rows(i1).Cells(2).Value = DataGridView12.Rows(i1).Cells(2).Value
                    DataGridView1.Rows(i1).Cells(3).Value = DataGridView12.Rows(i1).Cells(3).Value
                    DataGridView1.Rows(i1).Cells(4).Value = DataGridView12.Rows(i1).Cells(4).Value
                    DataGridView1.Rows(i1).Cells(5).Value = DataGridView12.Rows(i1).Cells(5).Value
                    DataGridView1.Rows(i1).Cells(6).Value = DataGridView12.Rows(i1).Cells(7).Value
                    DataGridView1.Rows(i1).Cells(7).Value = Microsoft.VisualBasic.Mid(DataGridView12.Rows(i1).Cells(6).Value, 1, 6)
                    DataGridView1.Rows(i1).Cells(8).Value = Microsoft.VisualBasic.Mid(DataGridView12.Rows(i1).Cells(6).Value, 7, 1)
                    DataGridView1.Rows(i1).Cells(9).Value = DataGridView12.Rows(i1).Cells(11).Value
                    Select Case Trim(DataGridView12.Rows(i1).Cells(12).Value)
                        Case "1" : DataGridView1.Rows(i1).Cells(11).Value = "VERDE"
                        Case "2" : DataGridView1.Rows(i1).Cells(11).Value = "AMARILLO"
                        Case "3" : DataGridView1.Rows(i1).Cells(11).Value = "CELESTE"
                        Case "4" : DataGridView1.Rows(i1).Cells(11).Value = "ROJO"
                        Case "" : DataGridView1.Rows(i1).Cells(11).Value = ""
                    End Select


                Next
                DataGridView1.Columns(0).Frozen = True
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
            End If

            TextBox8.ReadOnly = False
                TextBox4.Enabled = False
                TextBox5.Enabled = False
                Button2.Enabled = True
                Button5.Enabled = True

                For a9 = 0 To w1 - 1
                    cant9 = Val(DataGridView1.Rows(a9).Cells(5).Value)
                    sum9 = cant9 + Val(sum9)
                Next
                TextBox1.Text = sum9.ToString("N2")

            'End If
        End If
    End Sub
    Dim dt, dt1, hj As New DataTable
    Private Sub mostrar1()
        Try
            Dim func As New fguia2
            Dim dts As New vguia2
            func.gsfactu = TextBox16.Text
            Select Case TextBox17.Text.Length

                Case "1" : func.gnfactu = "0000000" & "" & TextBox17.Text
                Case "2" : func.gnfactu = "000000" & "" & TextBox17.Text
                Case "3" : func.gnfactu = "00000" & "" & TextBox17.Text
                Case "4" : func.gnfactu = "0000" & "" & TextBox17.Text
                Case "5" : func.gnfactu = "000" & "" & TextBox17.Text
                Case "6" : func.gnfactu = "00" & "" & TextBox17.Text
                Case "7" : func.gnfactu = "0" & "" & TextBox17.Text
                Case "8" : func.gnfactu = TextBox17.Text
            End Select
            func.gccia = Label14.Text
            dt = dts.consultar_cabecera_guia(func)
            If dt.Rows.Count <> 0 Then
                DataGridView5.DataSource = dt

                TextBox12.Text = DataGridView5.Rows(0).Cells(1).Value
                DateTimePicker1.Value = DataGridView5.Rows(0).Cells(0).Value
            End If
        Catch ex As Exception
        End Try
    End Sub
    Private Sub mostrar()
        Try
            Dim func As New fnia
            Dim dts As New vnia


            dts.gmes = TextBox4.Text
            dts.galmacen = ComboBox3.SelectedValue.ToString
            dts.gano = Label11.Text
            dts.gccia = Label14.Text
            Me.TextBox5.Text = func.buscar_nia(dts) + 1
            Select Case TextBox5.Text.Length

                Case "1" : TextBox5.Text = "00000" & "" & TextBox5.Text
                Case "2" : TextBox5.Text = "0000" & "" & TextBox5.Text
                Case "3" : TextBox5.Text = "000" & "" & TextBox5.Text
                Case "4" : TextBox5.Text = "00" & "" & TextBox5.Text
                Case "5" : TextBox5.Text = "0" & "" & TextBox5.Text
                Case "6" : TextBox5.Text = TextBox5.Text
            End Select

        Catch ex As Exception
            DataGridView1.DataSource = ""


        End Try
    End Sub
    Private Sub mostrar2()
        Try
            'Dim a9
            'Dim cant9, sum9 As Double
            Dim func1 As New fguia2
            Dim dts1 As New vguia2

            func1.gccia = Label14.Text
            func1.gsfactu = TextBox16.Text
            Select Case TextBox17.Text.Length

                Case "1" : func1.gnfactu = "0000000" & "" & TextBox17.Text
                Case "2" : func1.gnfactu = "000000" & "" & TextBox17.Text
                Case "3" : func1.gnfactu = "00000" & "" & TextBox17.Text
                Case "4" : func1.gnfactu = "0000" & "" & TextBox17.Text
                Case "5" : func1.gnfactu = "000" & "" & TextBox17.Text
                Case "6" : func1.gnfactu = "00" & "" & TextBox17.Text
                Case "7" : func1.gnfactu = "0" & "" & TextBox17.Text
                Case "8" : func1.gnfactu = TextBox17.Text
            End Select

            hj = dts1.consultar_detalle_guia(func1)
            If hj.Rows.Count <> 0 Then

                DataGridView6.DataSource = hj

                Dim w1 As Integer
                w1 = DataGridView6.Rows.Count
                DataGridView1.Rows.Add(w1)
                For i = 0 To w1 - 1
                    DataGridView1.Rows(i).Cells(6).ReadOnly = True
                    DataGridView1.Rows(i).Cells(0).Value = DataGridView6.Rows(i).Cells(0).Value
                    DataGridView1.Rows(i).Cells(1).Value = DataGridView6.Rows(i).Cells(1).Value
                    DataGridView1.Rows(i).Cells(2).Value = DataGridView6.Rows(i).Cells(2).Value
                    DataGridView1.Rows(i).Cells(3).Value = DataGridView6.Rows(i).Cells(3).Value
                    DataGridView1.Rows(i).Cells(4).Value = DataGridView6.Rows(i).Cells(5).Value
                    DataGridView1.Rows(i).Cells(5).Value = DataGridView6.Rows(i).Cells(6).Value
                    DataGridView1.Rows(i).Cells(7).Value = Mid(DataGridView6.Rows(i).Cells(8).Value, 1, 6)
                    DataGridView1.Rows(i).Cells(8).Value = Mid(DataGridView6.Rows(i).Cells(8).Value, 7, 1)
                    DataGridView1.Rows(i).Cells(6).Value = DataGridView6.Rows(i).Cells(7).Value
                    TextBox1.Text = DataGridView6.Rows(i).Cells(6).Value
                Next
            End If
        Catch ex As Exception

            MsgBox("NO EXISTE INFORMACION PARA MOSTRAR DETALLE")

        End Try
    End Sub
    Public enunciado As SqlCommand
    Public respuesta As SqlDataReader
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Public enunciado3 As SqlCommand
    Public respuesta3 As SqlDataReader
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty, ty2, ty3, ty30, ty32 As New DataTable
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

    Private Sub TextBox20_TextChanged(sender As Object, e As EventArgs) Handles TextBox20.TextChanged

    End Sub
    Dim Rsr24, Rsr As SqlDataReader
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        ty2.Clear()
        'TextBox18.Text = ComboBox3.SelectedValue.ToString
        Dim sql1023 As String = " SELECT A.nomb_15m FROM custom_vianny.dbo.Mag1500 A  Where a.ccia ='" + Trim(Label14.Text) + "' and a.talm_15m ='" + Trim(ComboBox3.SelectedValue.ToString) + "'"
        Dim cmd1023 As New SqlCommand(sql1023, conx)
        Rsr24 = cmd1023.ExecuteReader()
        If Rsr24.Read() = True Then
            TextBox2.Text = Rsr24(0)

        End If
        Rsr24.Close()

        llenar_combo_box2(ComboBox1, ComboBox3.SelectedValue.ToString)
        'ComboBox3.Enabled = False
        TextBox16.Enabled = True
        TextBox17.Enabled = True
        TextBox16.Select()
        Button3.Enabled = True
        TextBox4.Text = DateTimePicker1.Value.Month
        Select Case TextBox4.Text.Length
            Case "1" : TextBox4.Text = "0" & DateTimePicker1.Value.Month
            Case "2" : TextBox4.Text = TextBox4.Text

        End Select
        TextBox4.Select()


        TextBox4.Select()
        Button3.Enabled = False
        TextBox4.ReadOnly = False
        TextBox5.ReadOnly = False

    End Sub

    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged

    End Sub

    Sub llenar_combo_box2(ByVal cb As ComboBox, ByVal alm As String)

        Try

            conn = New SqlDataAdapter("select dest_21e,rtrim(ltrim(nomb_21e)) as nomb_21e from custom_vianny.dbo.cag210e where  Empr_21e ='" + Trim(Label14.Text) + "' AND almac_21e ='" + alm + "'", conx)
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
    Sub llenar_combo_alamcenes(ByVal ser As String)

        Try

            conn = New SqlDataAdapter("select talm_15m + '  |  '+ nomb_15m as nombre, talm_15m as almacen from custom_vianny.dbo.Mag1500  where ccia ='" + Trim(Label14.Text) + "' and  flag_15m ='1' and (seriencr_15m ='" + ser + "' or seriencr_15m ='1_2' )", conx)
            conn.Fill(ty30)
            ComboBox3.DataSource = ty30
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
    Sub alamcenes(ByVal almacen As String)

        Try

            conn = New SqlDataAdapter("select talm_15m + '  |  '+ nomb_15m as nombre, talm_15m as almacen from custom_vianny.dbo.Mag1500  where ccia ='" + Trim(Label14.Text) + "' and  flag_15m ='1'  and talm_15m ='" + almacen + "'", conx)
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
    Public Sub NumConFrac(ByVal CajaTexto As Windows.Forms.TextBox, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Char.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False

        Else
            e.Handled = True
        End If
    End Sub
    Private Sub NotaIngreso_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Label14.Text = "02" Then
            'ComboBox3.Items.Clear()
            'ComboBox3.Items.Add("08")
            'ComboBox3.Items.Add("10")
            'ComboBox3.Items.Add("24")
            'ComboBox3.Items.Add("67")
            TextBox8.Text = "20459785834"
            TextBox10.Text = "GRAUS INDUSTRIAS TEXTIL"
        Else
            TextBox8.Text = "20508740361"
            TextBox10.Text = "CONSORCIO TEXTIL VIANNY"
        End If
        bloqueado()
        ComboBox3.Select()
        ComboBox4.SelectedIndex = -1
        ComboBox2.SelectedIndex = 0

        abrir()
        llenar_combo_alamcenes("1")
        Dim sql1023 As String = " SELECT A.nomb_15m FROM custom_vianny.dbo.Mag1500 A  Where a.ccia ='" + Trim(Label14.Text) + "' and a.talm_15m ='" + Trim(ComboBox3.SelectedValue.ToString) + "'"
        Dim cmd1023 As New SqlCommand(sql1023, conx)
        Rsr = cmd1023.ExecuteReader()
        If Rsr.Read() = True Then
            TextBox2.Text = Rsr(0)

        End If
        Rsr.Close()
        'TextBox12.Text = 1
    End Sub
    Private Sub bloqueado()
        DateTimePicker1.Enabled = True
        TextBox9.Enabled = False
        TextBox20.Enabled = False
        TextBox16.Enabled = False
        TextBox17.Enabled = False
        TextBox12.Enabled = False
        DataGridView1.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox4.Enabled = False
        TextBox13.Enabled = False
        DateTimePicker1.Enabled = False
        Button5.Enabled = False
        Button7.Enabled = False
        Button9.Enabled = False
        TextBox1.Enabled = False
        Button8.Enabled = False
        Button1.Enabled = False
        RadioButton1.Enabled = False
        RadioButton2.Enabled = False
    End Sub
    Public Sub limpiar()
        TextBox9.Text = ""
        TextBox12.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        TextBox1.Text = ""
        'TextBox4.Text = ""
        TextBox5.Text = ""

    End Sub
    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox17_Click(sender As Object, e As EventArgs) Handles TextBox17.Click
        Select Case TextBox16.Text.Length
            Case "1" : TextBox16.Text = "000" & "" & TextBox16.Text
            Case "2" : TextBox16.Text = "00" & "" & TextBox16.Text
            Case "3" : TextBox16.Text = "0" & "" & TextBox16.Text
            Case "4" : TextBox16.Text = TextBox16.Text
        End Select
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged


    End Sub

    Private Sub TextBox20_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox20.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button2.Enabled = True
            Dim func As New fventas
            Dim dts As New vventas
            Dim text As String
            text = TextBox20.Text
            DataGridView1.Enabled = True
            Button5.Enabled = True
            dts.gcodigo = text

            If TextBox20.Text = "" Then


                MsgBox("EL CAMPO CODIGO ESTA VACIO")
            Else


                dt = func.buscar_codigo(dts)
                If dt.Rows.Count <> 0 Then
                    DataGridView1.Rows.Add(1)
                    DataGridView7.DataSource = dt

                    Dim num, num2, i17 As Integer
                    Dim num10, num12 As Integer
                    num = 0

                    If DataGridView1.Rows(0).Cells(0).Value = 0 Then

                        DataGridView1.Rows(0).Cells(0).Value = num + 1
                        Select Case DataGridView1.Rows(0).Cells(0).Value.ToString.Length
                            Case "1" : DataGridView1.Rows(0).Cells(0).Value = "00" & "" & DataGridView1.Rows(0).Cells(0).Value
                            Case "2" : DataGridView1.Rows(0).Cells(0).Value = "0" & "" & DataGridView1.Rows(0).Cells(0).Value
                            Case "3" : DataGridView1.Rows(0).Cells(0).Value = DataGridView1.Rows(0).Cells(0).Value

                        End Select
                        DataGridView1.Rows(0).Cells(2).Value = DataGridView7.Rows(0).Cells(0).Value
                        DataGridView1.Rows(0).Cells(3).Value = DataGridView7.Rows(0).Cells(1).Value
                        DataGridView1.Rows(0).Cells(6).Value = DataGridView7.Rows(0).Cells(2).Value
                        DataGridView1.Rows(0).Cells(1).ReadOnly = False
                        'DataGridView1.Rows(0).Cells(7).ReadOnly = False
                        'DataGridView1.Columns(7).MaxInputLength = 6
                        CType(DataGridView1.Columns(7), DataGridViewTextBoxColumn).MaxInputLength = 8
                        DataGridView7.DataSource = ""
                        TextBox20.Text = ""

                    Else
                        num10 = DataGridView1.Rows.Count
                        For i = 1 To num10 - 1
                            If Convert.ToInt16(DataGridView1.Rows(i).Cells(0).Value) < Convert.ToInt16(DataGridView1.Rows(i - 1).Cells(0).Value) Then
                                num2 = Convert.ToInt16(DataGridView1.Rows(i - 1).Cells(0).Value) + 1
                                i17 = i

                            End If


                        Next

                        If i17 = "50" Then
                            MsgBox("SOLO SE PERMITE 50 REGISTROS")
                        Else
                            If num12 <> 1 Then
                                DataGridView1.Rows(i17).Cells(0).Value = num2
                                Select Case DataGridView1.Rows(i17).Cells(0).Value.ToString.Length
                                    Case "1" : DataGridView1.Rows(i17).Cells(0).Value = "00" & "" & DataGridView1.Rows(i17).Cells(0).Value
                                    Case "2" : DataGridView1.Rows(i17).Cells(0).Value = "0" & "" & DataGridView1.Rows(i17).Cells(0).Value
                                    Case "3" : DataGridView1.Rows(i17).Cells(0).Value = DataGridView1.Rows(i17).Cells(0).Value

                                End Select
                                DataGridView1.Rows(i17).Cells(2).Value = DataGridView7.Rows(0).Cells(0).Value
                                DataGridView1.Rows(i17).Cells(3).Value = DataGridView7.Rows(0).Cells(1).Value
                                DataGridView1.Rows(i17).Cells(6).Value = DataGridView7.Rows(0).Cells(2).Value
                                DataGridView1.Rows(i17).Selected = True
                                DataGridView1.Rows(i17).Cells(1).ReadOnly = False
                                DataGridView1.Rows(i17).Cells(5).ReadOnly = False
                                DataGridView1.Rows(i17).Cells(4).ReadOnly = False
                                'DataGridView1.Rows(i17).Cells(7).ReadOnly = True
                                TextBox20.Text = ""
                                DataGridView2.DataSource = ""
                                For i = 0 To i17 - 1
                                    DataGridView1.Rows(i).Selected = False
                                Next
                            End If
                        End If


                    End If

                Else
                    MsgBox("EL CODIGO NO EXISTE")
                    TextBox20.Text = ""
                    DataGridView7.DataSource = ""
                End If
            End If
        End If
    End Sub
    Dim DT30 As DataTable

    Private Sub TextBox13_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox13.KeyDown
        If e.KeyCode = Keys.Enter Then

            Button2.Enabled = True
            Dim func As New fventas
            Dim dts As New vventas
            Dim text As String
            text = Trim(TextBox13.Text)
            DataGridView1.Enabled = True
            Button5.Enabled = True
            dts.gop = text
            dts.gccia = Label14.Text
            dts.galmacen = ComboBox3.Text
            If Trim(TextBox13.Text) = "" Then
                MsgBox("EL CAMPO PARTIDA ESTA VACIO")
            Else
                If Trim(TextBox8.Text).Length = 0 Then
                    MessageBox.Show("FALTA INGRESAR EL CLIENTE", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    DT30 = func.buscar_op(dts)
                    If DT30.Rows.Count <> 0 Then
                        DataGridView1.Rows.Add(1)
                        DataGridView8.DataSource = DT30
                        Dim num, num2, i17 As Integer
                        Dim num10, num11, num12, num14 As Integer
                        num = 0

                        If DataGridView1.Rows(0).Cells(0).Value = 0 Then

                            DataGridView1.Rows(0).Cells(0).Value = num + 1
                            Select Case DataGridView1.Rows(0).Cells(0).Value.ToString.Length
                                Case "1" : DataGridView1.Rows(0).Cells(0).Value = "00" & "" & DataGridView1.Rows(0).Cells(0).Value
                                Case "2" : DataGridView1.Rows(0).Cells(0).Value = "0" & "" & DataGridView1.Rows(0).Cells(0).Value
                                Case "3" : DataGridView1.Rows(0).Cells(0).Value = DataGridView1.Rows(0).Cells(0).Value

                            End Select
                            DataGridView1.Rows(0).Cells(1).Value = DataGridView8.Rows(0).Cells(0).Value
                            DataGridView1.Rows(0).Cells(2).Value = DataGridView8.Rows(0).Cells(1).Value
                            DataGridView1.Rows(0).Cells(3).Value = DataGridView8.Rows(0).Cells(2).Value
                            DataGridView1.Rows(0).Cells(7).Value = TextBox13.Text
                            DataGridView1.Rows(0).Cells(6).Value = "KGS"
                            'TextBox14.Text = DataGridView2.Rows(0).Cells(2).Value
                            'TextBox15.Text = DataGridView2.Rows(0).Cells(5).Value
                            CType(DataGridView1.Columns(7), DataGridViewTextBoxColumn).MaxInputLength = 6
                            DataGridView1.Rows(0).Cells(1).ReadOnly = False
                            DataGridView8.DataSource = ""
                            TextBox13.Text = ""
                        Else
                            num10 = DataGridView1.Rows.Count
                            For i = 1 To num10 - 1
                                If Convert.ToInt16(DataGridView1.Rows(i).Cells(0).Value) < Convert.ToInt16(DataGridView1.Rows(i - 1).Cells(0).Value) Then
                                    num2 = Convert.ToInt16(DataGridView1.Rows(i - 1).Cells(0).Value) + 1
                                    i17 = i
                                End If
                            Next
                            If i17 = "50" Then
                                MsgBox("SOLO SE PERMITE 50 REGISTROS")
                            Else
                                For i1 = 0 To i17 - 1
                                    num10 = Trim(Mid(TextBox13.Text, 1, 6))
                                    num11 = Trim(DataGridView1.Rows(i1).Cells(7).Value)
                                    num14 = Mid(num11, 1, 6)
                                Next
                                If num12 <> 1 Then
                                    DataGridView1.Rows(i17).Cells(0).Value = num2
                                    Select Case DataGridView1.Rows(i17).Cells(0).Value.ToString.Length
                                        Case "1" : DataGridView1.Rows(i17).Cells(0).Value = "00" & "" & DataGridView1.Rows(i17).Cells(0).Value
                                        Case "2" : DataGridView1.Rows(i17).Cells(0).Value = "0" & "" & DataGridView1.Rows(i17).Cells(0).Value
                                        Case "3" : DataGridView1.Rows(i17).Cells(0).Value = DataGridView1.Rows(i17).Cells(0).Value
                                    End Select
                                    DataGridView1.Rows(i17).Cells(1).Value = DataGridView8.Rows(0).Cells(0).Value
                                    DataGridView1.Rows(i17).Cells(2).Value = DataGridView8.Rows(0).Cells(1).Value
                                    DataGridView1.Rows(i17).Cells(3).Value = DataGridView8.Rows(0).Cells(2).Value
                                    DataGridView1.Rows(i17).Cells(7).Value = TextBox13.Text
                                    DataGridView1.Rows(i17).Cells(6).Value = "KGS"
                                    DataGridView1.Rows(i17).Selected = True
                                    DataGridView1.Rows(i17).Cells(1).ReadOnly = False
                                    DataGridView1.Rows(i17).Cells(4).ReadOnly = False
                                    DataGridView1.Rows(i17).Cells(6).ReadOnly = True
                                    TextBox13.Text = ""
                                    DataGridView8.DataSource = ""
                                    For i = 0 To i17 - 1
                                        DataGridView1.Rows(i).Selected = False
                                    Next
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    Dim Rsr300 As SqlDataReader
    Private Sub PictureBox3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs)

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

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub

    Friend Sub correlativo()
        Select Case TextBox4.Text.Length
            Case "1" : TextBox4.Text = "0" & "" & TextBox4.Text
            Case "2" : TextBox4.Text = TextBox4.Text

        End Select
        TextBox5.Enabled = True

        Try
            Dim func As New fnia
            Dim dts As New vnia


            Select Case TextBox4.Text.Length
                Case "1" : dts.gmes = "0" & "" & TextBox4.Text
                Case "2" : dts.gmes = TextBox4.Text

            End Select
            dts.galmacen = ComboBox3.SelectedValue.ToString
            dts.gano = Label11.Text
            dts.gccia = Label14.Text
            Me.TextBox5.Text = func.buscar_nia(dts) + 1
            Select Case TextBox5.Text.Length

                Case "1" : TextBox5.Text = "00000" & "" & TextBox5.Text
                Case "2" : TextBox5.Text = "0000" & "" & TextBox5.Text
                Case "3" : TextBox5.Text = "000" & "" & TextBox5.Text
                Case "4" : TextBox5.Text = "00" & "" & TextBox5.Text
                Case "5" : TextBox5.Text = "0" & "" & TextBox5.Text
                Case "6" : TextBox5.Text = TextBox5.Text
            End Select
            TextBox5.Select()

        Catch ex As Exception
        End Try
    End Sub
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            correlativo()
        Else
            If e.KeyCode = Keys.F1 Then
                Ingresos_Salidas_almacen.Label3.Text = Label14.Text
                Ingresos_Salidas_almacen.Label4.Text = Label11.Text
                Ingresos_Salidas_almacen.Label5.Text = ComboBox3.SelectedValue.ToString
                Ingresos_Salidas_almacen.Label6.Text = Trim(TextBox4.Text)
                Ingresos_Salidas_almacen.Label7.Text = "1"
                Ingresos_Salidas_almacen.Show()
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If ComboBox3.SelectedValue.ToString = "06" Then
            Lista_ARti.Label2.Text = ComboBox3.SelectedValue.ToString
            Lista_ARti.Label3.Text = Label14.Text
            Lista_ARti.Show()
        Else
            Lista_Articulos.Label2.Text = Label14.Text
            Lista_Articulos.Show()
        End If

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles TextBox17.TextChanged

    End Sub

    Private Sub ToolTip1_Popup(sender As Object, e As PopupEventArgs) Handles ToolTip1.Popup

    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        'Dim kk As New Agregar_Packing
        Select Case ComboBox3.SelectedValue.ToString
            Case "10" : Agregar_Packing.TextBox1.Text = "PRIMERA"
            Case "10" : Agregar_Packing.TextBox1.Text = "SEGUNDA"
            Case "10" : Agregar_Packing.TextBox1.Text = "TERCERA"
        End Select
        Agregar_Packing.Label3.Text = ComboBox3.SelectedValue.ToString
        Agregar_Packing.Label2.Text = Label11.Text
        Agregar_Packing.Label4.Text = Trim(Label14.Text)
        Agregar_Packing.Show()
    End Sub

    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click
        Articulos_stock.Label3.Text = ComboBox3.SelectedValue.ToString
        'Articulos_stock.Label3.Text = "24"
        Articulos_stock.Label4.Text = 20
        Articulos_stock.Label5.Text = Label11.Text - 1
        'Articulos_stock.Label5.Text = Label11.Text
        Articulos_stock.Label6.Text = Label14.Text
        Articulos_stock.Label7.Text = "2"
        Articulos_stock.Show()
    End Sub

    Private Sub habilitar()

        TextBox9.Enabled = True

        TextBox12.Enabled = True
        DataGridView1.Enabled = True

        TextBox13.Enabled = True

        Button1.Enabled = True
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        RadioButton1.Checked = True
    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged

    End Sub

    Dim hj2, hj1 As DataTable

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Trim(ComboBox1.Text) = "COMPRA PROVEDOR NACIONAL" Then
            TextBox15.Enabled = True
        End If
    End Sub
    Sub compras()
        DataGridView1.Enabled = True
        Button5.Enabled = True
        Button7.Enabled = False
        Button8.Enabled = False
        Button9.Enabled = False
        Button10.Enabled = True

        TextBox4.Enabled = False
        TextBox5.Enabled = False
        ComboBox4.Enabled = True
        Button1.Enabled = False
        Button2.Enabled = False
        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
    End Sub
    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            If TextBox5.Text = "" Then
                correlativo()
            End If
            ty2.Clear()
            abrir()
            llenar_combo_box2(ComboBox1, ComboBox3.SelectedValue.ToString)
            Dim me12 As String
            me12 = TextBox4.Text - Month(DateTimePicker1.Value)

            DateTimePicker1.Value = DateTimePicker1.Value.AddMonths(me12)
            Dim i As Integer
            Dim ml As New vnia
            Dim mk As New fnia
            TextBox13.Enabled = True
            'TextBox12.Enabled = True
            'TextBox12.ReadOnly = False

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
            ml.gccia = Label14.Text
            If mk.mostrar_correlativo_nia1(ml) = True Then
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
                    TextBox8.Text = DataGridView9.Rows(0).Cells(4).Value
                    gu = DataGridView9.Rows(0).Cells(5).Value
                    TextBox10.Text = DataGridView9.Rows(0).Cells(6).Value
                    ComboBox2.SelectedItem = 0

                    Select Case Trim(DataGridView9.Rows(0).Cells(9).Value)
                        Case "009" : ComboBox4.Text = "GUIA"
                        Case "001" : ComboBox4.Text = "FACT"
                        Case "002" : ComboBox4.Text = "OTRO"
                        Case "" : ComboBox4.Text = ""
                    End Select

                    abrir()
                    enunciado2 = New SqlCommand("select rtrim(ltrim(nomb_21e)) as nomb_21e from custom_vianny.dbo.cag210e where  Empr_21e ='" + Trim(Label14.Text) + "' AND almac_21e='" + ComboBox3.SelectedValue.ToString + "' and dest_21e='" + Trim(DataGridView9.Rows(0).Cells(5).Value) + "'", conx)
                    respuesta2 = enunciado2.ExecuteReader
                    While respuesta2.Read
                        ComboBox1.Text = respuesta2.Item("nomb_21e")
                    End While
                    respuesta2.Close()

                    If Convert.ToString(DataGridView9.Rows(0).Cells(7).Value) = 1 Then
                        TextBox11.Text = ""
                        TextBox14.Text = ""
                    Else
                        TextBox14.Text = DataGridView9.Rows(0).Cells(7).Value
                        Select Case TextBox14.Text.Length
                            Case "1" : TextBox14.Text = "000" & "" & TextBox14.Text
                            Case "2" : TextBox14.Text = "00" & "" & TextBox14.Text
                            Case "3" : TextBox14.Text = "0" & "" & TextBox14.Text
                            Case "4" : TextBox14.Text = TextBox14.Text


                        End Select
                        enunciado3 = New SqlCommand("SELECT  nomb_4 as Nombre FROM custom_vianny.dbo.Qag0400 A  Where CCIA = '" + Trim(Label14.Text) + "'  and flag_4 ='1' and fase_4 ='" + Trim(TextBox14.Text) + "'", conx)
                        respuesta3 = enunciado3.ExecuteReader
                        While respuesta3.Read
                            TextBox11.Text = respuesta3.Item("Nombre")
                        End While
                        respuesta3.Close()
                    End If

                    If DataGridView9.Rows(0).Cells(3).Value = 1 Then
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

                If hj1.Rows.Count <> 0 Then
                    Dim nu1 As Integer
                    nu1 = hj1.Rows.Count
                    DataGridView1.Rows.Add(nu1)
                    For i1 = 0 To nu1 - 1
                        DataGridView10.DataSource = hj1
                        DataGridView1.Rows(i1).Cells(0).Value = DataGridView10.Rows(i1).Cells(0).Value
                        DataGridView1.Rows(i1).Cells(1).Value = DataGridView10.Rows(i1).Cells(1).Value
                        DataGridView1.Rows(i1).Cells(2).Value = DataGridView10.Rows(i1).Cells(2).Value
                        DataGridView1.Rows(i1).Cells(3).Value = DataGridView10.Rows(i1).Cells(3).Value
                        DataGridView1.Rows(i1).Cells(4).Value = DataGridView10.Rows(i1).Cells(4).Value
                        DataGridView1.Rows(i1).Cells(5).Value = DataGridView10.Rows(i1).Cells(5).Value
                        If Mid(DataGridView10.Rows(i1).Cells(2).Value, 4, 3) <> "NOT" Then
                            DataGridView1.Rows(i1).Cells(7).Value = Mid(DataGridView10.Rows(i1).Cells(8).Value, 1, 6)
                            DataGridView1.Rows(i1).Cells(8).Value = Mid(DataGridView10.Rows(i1).Cells(8).Value, 7, 1)
                        Else
                            DataGridView1.Rows(i1).Cells(7).Value = DataGridView10.Rows(i1).Cells(8).Value
                            DataGridView1.Rows(i1).Cells(8).Value = ""
                        End If

                        DataGridView1.Rows(i1).Cells(6).Value = DataGridView10.Rows(i1).Cells(7).Value
                        DataGridView1.Rows(i1).Cells(10).Value = DataGridView10.Rows(i1).Cells(11).Value
                        DataGridView1.Rows(i1).Cells(9).Value = DataGridView10.Rows(i1).Cells(12).Value

                        Select Case Trim(DataGridView10.Rows(i1).Cells(13).Value)
                            Case "1" : DataGridView1.Rows(i1).Cells(11).Value = "VERDE"
                            Case "2" : DataGridView1.Rows(i1).Cells(11).Value = "AMARILLO"
                            Case "3" : DataGridView1.Rows(i1).Cells(11).Value = "CELESTE"
                            Case "4" : DataGridView1.Rows(i1).Cells(11).Value = "ROJO"
                            Case "" : DataGridView1.Rows(i1).Cells(11).Value = ""

                        End Select
                    Next
                End If

                Button5.Enabled = False
                Button7.Enabled = True
                Button9.Enabled = True
                Button8.Enabled = True
                Button8.Enabled = True
                TextBox4.Enabled = False
                TextBox5.Enabled = False
                ComboBox3.Enabled = False
                Button2.Enabled = False
                TextBox8.ReadOnly = False
                CheckBox1.Enabled = False
                CheckBox2.Enabled = False
            Else
                TextBox8.Enabled = True
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
                ComboBox3.Enabled = False
                TextBox8.ReadOnly = False
                Button2.Enabled = False
                ComboBox2.Enabled = True
                ComboBox1.Enabled = True
                ComboBox4.Enabled = True
                'DateTimePicker1.Enabled = True
                TextBox9.ReadOnly = False
                TextBox9.Enabled = True
                DataGridView1.Enabled = True
                'TextBox12.Enabled = True
                TextBox20.Enabled = True
                Button5.Enabled = True
                Button7.Enabled = True
                Button8.Enabled = True
                Button9.Enabled = True
                DataGridView1.Rows.Clear()
                TextBox9.Text = ""
                'TextBox12.Text = 1
                CheckBox1.Enabled = True
                CheckBox2.Enabled = True
            End If


            PictureBox5.Enabled = True
            DateTimePicker1.Select()
        End If
    End Sub
    Dim Rsr3001, Rsr25, Rsr3 As SqlDataReader

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ToolTip1.SetToolTip(Button5, "Guardar Informacion")
        Dim fun As New vinsertar_nia
        Dim func As New fnia
        Dim func1 As New insertardetallenia
        Dim func2 As New fnia
        Dim con As String
        Dim hn As New vnia
        Dim fg As New fnia


        Dim ui, ct, ct1 As Integer
        Dim Cup, Ccon, Cpartida, Coc As Integer
        Cup = 0
        Ccon = 0
        Cpartida = 0
        Coc = 0

        ct = 0
        ct1 = 0
        ui = DataGridView1.Rows.Count
        If ComboBox3.Text <> "06" Then
            For i11 = 0 To ui - 1
                If DataGridView1.Rows(i11).Cells(11).Value Is Nothing Then
                    ct = ct + 1
                End If
            Next
        End If
        For i13 = 0 To ui - 1
            If Trim(DataGridView1.Rows(i13).Cells(9).Value) = "" Then
                ct1 = ct1 + 1
            End If
            If Trim(DataGridView1.Rows(i13).Cells(4).Value).Length = 0 Then
                Cup = Cup + 1
            End If
            If Trim(DataGridView1.Rows(i13).Cells(5).Value).Length = 0 Then
                Ccon = Ccon + 1
            End If
            If Trim(DataGridView1.Rows(i13).Cells(7).Value).Length = 0 Then
                Cpartida = Cpartida + 1
            End If

        Next

        Dim sql102213 As String = "select ocomp_21e from custom_vianny.dbo.cag210e where dest_21e ='" + ComboBox1.SelectedValue.ToString + "' and Empr_21e ='" + Label14.Text + "' and almac_21e ='" + Trim(ComboBox3.SelectedValue.ToString) + "'"
        Dim cmd102213 As New SqlCommand(sql102213, conx)
        Rsr3001 = cmd102213.ExecuteReader()
        Dim jo, valor As Integer
        If Rsr3001.Read() Then
            jo = Rsr3001(0)
        End If
        Rsr3001.Close()

        Dim sql1 As String = "select COUNT(ncom_3a) from custom_vianny.dbo.map030f where CCIA_3A ='" + Label14.Text + "' and CPERIODO_3A ='" + Label11.Text + "' and talm_3a ='" + Trim(ComboBox3.SelectedValue.ToString) + "' and ccom_3a ='1' and ncom_3a ='" + TextBox4.Text & TextBox5.Text + "'"
        Dim cmd1 As New SqlCommand(sql1, conx)
        Rsr3 = cmd1.ExecuteReader()

        If Rsr3.Read() Then
            valor = Convert.ToInt32(Rsr3(0))
        End If
        Rsr3.Close()


        If jo = 1 And Trim(TextBox12.Text).Length = 0 Then
            MessageBox.Show("EL MOTIVO COMPRA ES OBLIGATORIO INGRESAR  REGISTRO DE GUIA REMISION O FACTURA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else

            If jo = 1 And ct1 > 0 Then
                MessageBox.Show("EL MOTIVO COMPRA ES OBLIGATORIO INGRESAR LA OC EN CADA ITEMS", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else

                If ct > 0 Then
                    MessageBox.Show("FALTA AGREGAR LA CLASIFICACION DEL UN ARTICULO EN UNA FILA DE LA TABLA NSA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else

                    If Cup > 0 Then
                        MessageBox.Show("FALTA INGRESAR LA CANTIDAD DE PRESENTACION EN ALGUNA CELDA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Else
                        If Ccon > 0 Then
                            MessageBox.Show("FALTA INGRESAR LA CANTIDAD EN KG/MTS EN ALGUNA CELDA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)

                        Else
                            If Cpartida > 0 Then
                                MessageBox.Show("FALTA INGRESAR LA PARTIDA O LOTE EN ALGUNA CELDA", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Else
                                Dim cor1, correlativo As String

                                If valor > 0 And Label15.Text = "0" Then

                                    Dim func14 As New fnia
                                    Dim dt4 As New vnia
                                    Dim cor As Integer

                                    TextBox4.Text = DateTime.Now.ToString("MM")
                                    dt4.gccia = Label14.Text
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

                                    con = correlativo
                                    hn.gccia = Label14.Text
                                    hn.gcomprobante = con
                                    hn.galmacen = ComboBox3.SelectedValue.ToString
                                    hn.gncom = 1
                                    hn.gano = Label11.Text
                                    fg.eliminar_nia(hn)

                                    fun.gccia = Label14.Text
                                    fun.gncom = con
                                    fun.gglosa = TextBox9.Text
                                    fun.gfecha = DateTimePicker1.Value
                                    fun.gguia = final
                                    fun.gano = Label11.Text
                                    Select Case Trim(ComboBox4.Text)
                                        Case "GUIA" : fun.gdoc = "009"
                                        Case "FACT" : fun.gdoc = "001"
                                        Case "OTRO" : fun.gdoc = "002"
                                        Case "" : fun.gdoc = ""
                                    End Select
                                    fun.galmacen = ComboBox3.SelectedValue.ToString
                                    fun.gempresa = TextBox8.Text
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
                                    fun.gusuario = TextBox7.Text
                                    fun.gfase = TextBox14.Text
                                    Dim i, num2 As Integer

                                    num2 = DataGridView1.Rows.Count
                                    For i = 0 To num2 - 1
                                        Dim nu As String
                                        Dim jj As New fingresopac
                                        Dim aa As New vpackingtela
                                        func1.gccia = Label14.Text
                                        func1.gncom = con

                                        nu = DataGridView1.Rows(i).Cells(0).Value

                                        Select Case nu.Length
                                            Case "1" : nu = "00" & "" & nu
                                            Case "2" : nu = "0" & "" & nu
                                            Case "3" : nu = nu
                                        End Select
                                        func1.gitem = nu

                                        func1.glinea = DataGridView1.Rows(i).Cells(2).Value
                                        If Trim(DataGridView1.Rows(i).Cells(1).Value) = "" Then
                                            func1.gop = ""
                                        Else
                                            func1.gop = DataGridView1.Rows(i).Cells(1).Value
                                        End If

                                        func1.gund = DataGridView1.Rows(i).Cells(6).Value
                                        func1.gcantidad = DataGridView1.Rows(i).Cells(4).Value
                                        func1.gcantidad1 = DataGridView1.Rows(i).Cells(5).Value
                                        If DataGridView1.Rows(i).Cells(7).Value = "" Then
                                            func1.gpartida = ""
                                        Else
                                            func1.gpartida = DataGridView1.Rows(i).Cells(7).Value & DataGridView1.Rows(i).Cells(8).Value
                                        End If

                                        func1.galmacen = ComboBox3.SelectedValue.ToString
                                        func1.gumpresentacion = "RLL"
                                        func1.gotejeduria = ""

                                        If Trim(DataGridView1.Rows(i).Cells(9).Value) = "" Then
                                            func1.goc = ""
                                        Else
                                            func1.goc = DataGridView1.Rows(i).Cells(9).Value
                                        End If
                                        func1.gano = Label11.Text
                                        If Trim(DataGridView1.Rows(i).Cells(7).Value) = "" Then
                                            func1.glote = ""
                                        Else
                                            func1.glote = DataGridView1.Rows(i).Cells(7).Value & DataGridView1.Rows(i).Cells(8).Value
                                        End If
                                        If DataGridView1.Rows(i).Cells(10).Value = "0.00" Then
                                            func1.gcenvio = "0.00"
                                        Else
                                            func1.gcenvio = DataGridView1.Rows(i).Cells(10).Value
                                        End If
                                        Select Case Trim(DataGridView1.Rows(i).Cells(11).Value)
                                            Case "VERDE" : func1.gclasif = "1"
                                            Case "AMARILLO" : func1.gclasif = "2"
                                            Case "CELESTE" : func1.gclasif = "3"
                                            Case "ROJO" : func1.gclasif = "4"
                                            Case "" : func1.gclasif = ""
                                        End Select

                                        If Trim(DataGridView1.Rows(i).Cells(7).Value) = "" Then
                                            aa.gpartida = ""
                                        Else
                                            aa.gpartida = DataGridView1.Rows(i).Cells(7).Value & DataGridView1.Rows(i).Cells(8).Value
                                        End If
                                        aa.galmacen = ComboBox3.SelectedValue.ToString
                                        aa.gseleccionado = 1
                                        aa.gguia = ""
                                        jj.actualizar_packing(aa)
                                        func2.insertar_detalle_nia(func1)
                                        'actualizar informacion de la oc
                                        Dim cmd20 As New SqlCommand("update custom_vianny.dbo.lap0300 set percep_3a='1' where ccia_3a =@ccia_3a  and ncom_3a =@ncom_3a and item_3a =@item_3a", conx)
                                        cmd20.Parameters.AddWithValue("@ccia_3a", (Label14.Text.ToString.Trim))
                                        If DataGridView1.Rows(i).Cells(9).Value IsNot Nothing AndAlso DataGridView1.Rows(i).Cells(9).Value.ToString.Trim.Length > 0 Then
                                            cmd20.Parameters.AddWithValue("@ncom_3a", (DataGridView1.Rows(i).Cells(9).Value.ToString.Trim))
                                        Else
                                            cmd20.Parameters.AddWithValue("@ncom_3a", "")
                                        End If
                                        cmd20.Parameters.AddWithValue("@item_3a", (DataGridView1.Rows(i).Cells(0).Value.ToString.Trim))
                                        cmd20.ExecuteNonQuery()
                                    Next
                                    If func.insertar_cabecera_nia(fun) Then
                                        Dim oc As String
                                        If DataGridView1.Rows(0).Cells(9).Value IsNot Nothing AndAlso DataGridView1.Rows(0).Cells(9).Value.ToString.Trim.Length > 0 Then
                                            oc = DataGridView1.Rows(0).Cells(9).Value.ToString.Trim
                                        Else
                                            oc = ""
                                        End If
                                        abrir()
                                        Dim sql1022 As String = "select COUNT(item_3a),isnull(SUM(cast (percep_3a  as int)),0) from custom_vianny.dbo.lap0300 where  ncom_3a ='" + oc + "' and ccia_3a='" + Label14.Text.ToString.Trim + "'"
                                        Dim cmd1022 As New SqlCommand(sql1022, conx)
                                        Rsr25 = cmd1022.ExecuteReader()
                                        Dim i5 As Integer
                                        i5 = 0
                                        If Rsr25.Read() Then
                                            If Convert.ToInt32(Rsr25(0)) = Convert.ToInt32(Rsr25(1)) Then
                                                Rsr25.Close()
                                                Dim cmd20 As New SqlCommand("UPDATE custom_vianny.dbo.lag0300 set termina_3 ='1' where  ncom_3 =@ncom_3a and ccia_3=@ccia_3a", conx)
                                            cmd20.Parameters.AddWithValue("@ccia_3a", (Label14.Text.ToString.Trim))
                                            If DataGridView1.Rows(0).Cells(9).Value Is Nothing Then
                                                cmd20.Parameters.AddWithValue("@ncom_3a", "")
                                            Else
                                                cmd20.Parameters.AddWithValue("@ncom_3a", (DataGridView1.Rows(0).Cells(9).Value.ToString.Trim))
                                            End If

                                            cmd20.ExecuteNonQuery()
                                            End If
                                        End If
                                        Rsr25.Close()
                                        If Label16.Text = "1" Then
                                            AlmacenOrdenCompra.Listar_Informacion()
                                        End If
                                        MessageBox.Show("Datos registrados correctamente ", "Guardando registros", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                        Dim hj2 As New flog
                                        Dim hj1 As New vlog
                                        hj1.gmodulo = "NIA-ALMACEN"
                                        hj1.gcuser = MDIParent1.Label3.Text
                                        If Label15.Text = "0" Then
                                            hj1.gaccion = 1
                                        Else
                                            hj1.gaccion = 2
                                        End If

                                        hj1.gpc = My.Computer.Name
                                        hj1.gfecha = DateTimePicker2.Value
                                        hj1.gdetalle = Trim(ComboBox3.SelectedValue.ToString) & Trim(TextBox4.Text) & Trim(TextBox5.Text)
                                        hj1.gccia = Label14.Text
                                        hj2.insertar_log(hj1)
                                        Dim respuesta As DialogResult

                                        respuesta = MessageBox.Show("DESEA IMPRIMIR LA NOTA DE INGRESO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                                        If (respuesta = Windows.Forms.DialogResult.Yes) Then

                                            Reporte_Nia_Nsa.TextBox1.Text = ComboBox3.SelectedValue.ToString
                                            Reporte_Nia_Nsa.TextBox2.Text = 1
                                            Reporte_Nia_Nsa.TextBox3.Text = TextBox4.Text & "" & TextBox5.Text
                                            Reporte_Nia_Nsa.TextBox4.Text = Label11.Text
                                            Reporte_Nia_Nsa.TextBox5.Text = Label14.Text
                                            Reporte_Nia_Nsa.Show()
                                        End If
                                        limpiar()
                                    End If
                                    limpiar()
                                    TextBox8.Enabled = False
                                    TextBox8.Text = ""
                                    TextBox15.Text = ""
                                    TextBox10.Text = ""
                                    CheckBox1.Checked = False
                                    TextBox4.Enabled = True
                                    TextBox5.Enabled = False
                                    TextBox9.Enabled = False
                                    TextBox12.Enabled = False
                                    ComboBox4.SelectedIndex = -1
                                    DateTimePicker1.Enabled = False
                                    DataGridView1.Enabled = False
                                    TextBox13.Enabled = False
                                    DataGridView1.Rows.Clear()
                                    TextBox5.Select()
                                    Dim func12 As New fnia
                                    Dim dts As New vnia
                                    TextBox4.Text = DateTime.Now.ToString("MM")
                                    dts.gccia = Label14.Text
                                    dts.gmes = Me.TextBox4.Text
                                    dts.galmacen = ComboBox3.SelectedValue.ToString
                                    dts.gano = Label11.Text

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
                                    TextBox4.Select()
                                    Label15.Text = "0"
                                    TextBox16.Text = "1"
                                End If
                            End If
                        End If
                    End If
                End If
            End If


    End Sub

    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.F1 Then
            If Trim(TextBox5.Text).Length > 0 Then
                Dim cli As New Clientes
                cli.TextBox3.Text = "35"
                cli.ShowDialog()
            Else
                MessageBox.Show("FALTA INGRESAR EL CORRELATIVO DEL COMPROBANTE", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged

    End Sub

    Private Sub PictureBox7_MouseHover(sender As Object, e As EventArgs) Handles PictureBox7.MouseHover
        ToolTip1.SetToolTip(PictureBox7, "STOCK INCIAL")
        ToolTip1.ToolTipTitle = "NOTA INGRESO"
        ToolTip1.ToolTipIcon = ToolTipIcon.Info
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        TextBox12.Enabled = True
        TextBox12.ReadOnly = False
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If Trim(TextBox8.Text).Length = 0 Then
            MessageBox.Show("FALTA INGRESAR EL CLIENTE", "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim pro As New pproductos
            pro.CCIA.Text = Label14.Text
            pro.ALMACEN.Text = ComboBox3.SelectedValue.ToString
            pro.ANO.Text = Label11.Text
            pro.FECHA.Text = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
            pro.TextBox1.Text = ComboBox3.SelectedValue.ToString
            pro.Label3.Text = 1
            pro.Show()

        End If




    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        If Trim(TextBox4.Text) <> Month(DateTimePicker1.Value) Then
            MsgBox("LA FECHA DE EMISION NO CORRESPONDE AL MES DE REGISTRO")
            Dim me12 As String
            me12 = TextBox4.Text - Month(DateTimePicker1.Value)
            DateTimePicker1.Value = DateTimePicker1.Value.AddMonths(me12)
        End If
    End Sub

    Private Sub TextBox14_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox14.KeyDown
        If e.KeyCode = Keys.F1 Then
            Procesos.Label1.Text = "2"
            Procesos.Label2.Text = Label14.Text
            Procesos.Show()
        Else
            If e.KeyCode = Keys.Enter Then
                TextBox8.Select()
            End If
        End If
    End Sub

    Private Sub PictureBox2_KeyDown(sender As Object, e As KeyEventArgs)

    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox14.Select()
        End If
    End Sub

    Private Sub TextBox13_HideSelectionChanged(sender As Object, e As EventArgs) Handles TextBox13.HideSelectionChanged

    End Sub

    Private Sub TextBox15_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox15.KeyDown
        If e.KeyCode = Keys.Enter Then
            'Dim k As Integer
            'k = DataGridView1.Rows.Count


            If CheckBox1.Checked = True Then
                'DataGridView1.Rows.Clear()
                abrir()
                Dim cmd As New SqlDataAdapter("select item_3a,linea_3a,nomb_17 ,cante_3a,cant_3a,unide2_3a,c.unid_17,ncom_3a,ISNULL(op_3a,'') from custom_vianny.dbO.laP0300 l
left join custom_vianny.dbo.cag1700 c on l.ccia_3a = c.ccia and l.linea_3a =c.linea_17
WHERE ncom_3a ='" + Trim(TextBox15.Text) + "' and ccia ='" + Trim(Label14.Text) + "'  and flag_3a =1", conx)
                Dim da As New DataTable
                cmd.Fill(da)
                If da.Rows.Count > 0 Then
                    DataGridView11.DataSource = da

                    Dim f As Integer
                    f = DataGridView11.Rows.Count
                    DataGridView1.Rows.Add(f)
                    For i = 0 To f - 1
                        DataGridView1.Rows(i).Cells(0).Value = DataGridView11.Rows(i).Cells(0).Value
                        DataGridView1.Rows(i).Cells(1).Value = DataGridView11.Rows(i).Cells(8).Value
                        DataGridView1.Rows(i).Cells(2).Value = DataGridView11.Rows(i).Cells(1).Value
                        DataGridView1.Rows(i).Cells(3).Value = DataGridView11.Rows(i).Cells(2).Value
                        DataGridView1.Rows(i).Cells(6).Value = DataGridView11.Rows(i).Cells(6).Value
                        DataGridView1.Rows(i).Cells(5).Value = DataGridView11.Rows(i).Cells(4).Value
                        DataGridView1.Rows(i).Cells(9).Value = DataGridView11.Rows(i).Cells(7).Value
                    Next

                    Dim Rs1 As SqlDataReader
                    Dim sql1 As String = "select fich_3,nomb_10,gloa_3 from custom_vianny.dbO.lag0300 l
left join custom_vianny.dbO.cag1000 c on '01' = c.ccia and l.fich_3 = c.fich_10
WHERE ncom_3 ='" + Trim(TextBox15.Text) + "' and ccia_3 ='" + Trim(Label14.Text) + "'  and flag_3 =1"
                    Dim cmd1 As New SqlCommand(sql1, conx)
                    Rs1 = cmd1.ExecuteReader()
                    Rs1.Read()
                    ComboBox2.SelectedIndex = 1
                    TextBox8.Text = Rs1(0)
                    TextBox10.Text = Rs1(1)
                    TextBox9.Text = Rs1(2)

                    Rs1.Close()
                Else
                    MsgBox("LA ORDEN DE COMRPA NO EXISTE")
                End If

            Else
                If CheckBox2.Checked = True Then
                    'DataGridView1.Rows.Clear()
                    abrir()
                    Dim cmd As New SqlDataAdapter("Select item_4a,linea_4a,nomb_17 ,cant_4a,unid_17,c.unid_17,ncom_4a, op_4a from custom_vianny.dbO.lap0400 l
left join custom_vianny.dbo.cag1700 c on l.ccia_4a = c.ccia and l.linea_4a =c.linea_17
WHERE ncom_4a ='" + Trim(TextBox15.Text) + "' and ccia ='" + Trim(Label14.Text) + "'  and flag_4a =1", conx)
                    Dim da As New DataTable
                    cmd.Fill(da)
                    DataGridView11.DataSource = da

                    Dim f As Integer
                    f = DataGridView11.Rows.Count
                    DataGridView1.Rows.Add(f)
                    For i = 0 To f - 1
                        DataGridView1.Rows(i).Cells(0).Value = DataGridView11.Rows(i).Cells(0).Value
                        DataGridView1.Rows(i).Cells(1).Value = DataGridView11.Rows(i).Cells(7).Value
                        DataGridView1.Rows(i).Cells(2).Value = DataGridView11.Rows(i).Cells(1).Value
                        DataGridView1.Rows(i).Cells(3).Value = DataGridView11.Rows(i).Cells(2).Value
                        DataGridView1.Rows(i).Cells(6).Value = DataGridView11.Rows(i).Cells(5).Value
                        DataGridView1.Rows(i).Cells(5).Value = DataGridView11.Rows(i).Cells(3).Value
                        DataGridView1.Rows(i).Cells(9).Value = DataGridView11.Rows(i).Cells(6).Value
                    Next

                    Dim Rs1 As SqlDataReader
                    Dim sql1 As String = "select fich_4,nomb_10,gloa_4 from custom_vianny.dbO.lag0400 l
left join custom_vianny.dbO.cag1000 c on l.ccia_4 = c.ccia and l.fich_4 = c.fich_10
WHERE ncom_4 ='" + Trim(TextBox15.Text) + "' and ccia_4 ='" + Trim(Label14.Text) + "'  and flag_4 =1"
                    Dim cmd1 As New SqlCommand(sql1, conx)
                    Rs1 = cmd1.ExecuteReader()
                    Rs1.Read()
                    ComboBox2.SelectedIndex = 1
                    TextBox8.Text = Rs1(0)
                    TextBox10.Text = Rs1(1)
                    TextBox9.Text = Rs1(2)

                    Rs1.Close()
                End If
            End If

        End If
    End Sub

    Private Sub DateTimePicker1_Click(sender As Object, e As EventArgs) Handles DateTimePicker1.Click
        'MsgBox("BIEN")
        'MsgBox(Month(DateTimePicker1.Value))
        'If Trim(TextBox4.Text) <> Month(DateTimePicker1.Value) Then
        '    MsgBox("LA FECHA DE EMISION NO CORRESPONDE AL MES DE REGISTRO")
        'End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            CheckBox2.Checked = False
            TextBox15.Enabled = True
            TextBox15.Select()
        Else
            CheckBox2.Checked = True
            TextBox15.Select()
        End If
    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            CheckBox1.Checked = False
            TextBox15.Enabled = True
        Else
            CheckBox1.Checked = True
        End If
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        NumConFrac(TextBox4, e)
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        NumConFrac(TextBox5, e)
    End Sub

    Private Sub TextBox20_HideSelectionChanged(sender As Object, e As EventArgs) Handles TextBox20.HideSelectionChanged

    End Sub
    Dim Rsr34, Rsr35, Rs1 As SqlDataReader

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        ToolTip1.SetToolTip(Button8, "Anular Nota Ingreso")
        'System.Diagnostics.Process.Start("www.google.com")
        Dim sql102213 As String = "SELECT cierre_3 FROM custom_vianny.dbo.Cie0300 A  Where a.ccia ='" + Label14.Text + "' AND A.talm_3 ='" + ComboBox3.SelectedValue.ToString + "' AND A.fcie_3='" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "'"
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
        Dim ml As New vnia
        Dim mk As New fnia
        ml.gcomprobante = TextBox4.Text & TextBox5.Text
        ml.galmacen = ComboBox3.SelectedValue.ToString
        ml.gano = Label11.Text
        ml.gccia = Label14.Text
        respuesta = MessageBox.Show("QUIERES ANULAR NOTA DE INGRESO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            If jo = 0 Then
                Label12.Visible = True
                RadioButton2.Checked = True
                RadioButton1.Checked = False
                mk.anular_nia(ml)

                Dim hj2 As New flog
                Dim hj1 As New vlog
                hj1.gmodulo = "NIA-ALMACEN"
                hj1.gcuser = MDIParent1.Label3.Text
                hj1.gaccion = 3
                hj1.gpc = My.Computer.Name
                hj1.gfecha = DateTimePicker2.Value
                hj1.gdetalle = Trim(ComboBox3.SelectedValue.ToString) & Trim(TextBox4.Text) & Trim(TextBox5.Text)
                hj1.gccia = Label14.Text
                hj2.insertar_log(hj1)
                TextBox4.Enabled = True
                TextBox4.Select()
                DateTimePicker1.Enabled = False
                TextBox9.Enabled = False
                TextBox12.Enabled = False
                TextBox5.Enabled = True
                TextBox13.Enabled = False
                DataGridView1.Enabled = False
                Button5.Enabled = False
                Button8.Enabled = False
                Button7.Enabled = False
                Button9.Enabled = False

                Label12.Visible = False
                limpiar()
                TextBox4.Enabled = True
                TextBox5.Enabled = False
                TextBox9.Enabled = False
                TextBox12.Enabled = False
                ComboBox4.SelectedIndex = -1
                DateTimePicker1.Enabled = False
                DataGridView1.Enabled = False
                TextBox13.Enabled = False
                DataGridView1.Rows.Clear()
                TextBox5.Select()
                Dim func12 As New fnia
                Dim dts As New vnia
                TextBox4.Text = DateTime.Now.ToString("MM")
                dts.gccia = Label14.Text
                dts.gmes = Me.TextBox4.Text
                dts.galmacen = ComboBox3.SelectedValue.ToString
                dts.gano = Label11.Text

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
                TextBox4.Select()
                TextBox8.Text = ""
                TextBox10.Text = ""
            Else
                MsgBox("DIA " & DateTimePicker1.Value & " CERRADO,NO SE PUEDE EDITAR LA INFORMACION - CONSULTE CON EL AREA CONTABLE ")
            End If

        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        ToolTip1.SetToolTip(Button9, "Imprimir Informacion")
        Reporte_Nia_Nsa.TextBox1.Text = ComboBox3.SelectedValue.ToString
        Reporte_Nia_Nsa.TextBox2.Text = 1
        Reporte_Nia_Nsa.TextBox3.Text = TextBox4.Text & "" & TextBox5.Text
        Reporte_Nia_Nsa.TextBox4.Text = Label11.Text
        Reporte_Nia_Nsa.TextBox5.Text = Label14.Text
        Reporte_Nia_Nsa.Show()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        ToolTip1.SetToolTip(Button10, "Retroceder")
        bloqueado()
        limpiar()
        TextBox8.Enabled = False
        TextBox5.ReadOnly = False
        TextBox4.ReadOnly = False
        TextBox5.Enabled = True
        TextBox4.Enabled = True
        Label12.Visible = False
        TextBox4.Select()
        ComboBox3.Enabled = True
        DataGridView1.Rows.Clear()
        ComboBox3.Enabled = True
        ty2.Clear()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ToolTip1.SetToolTip(Button7, "Editar Informacion")
        If RadioButton2.Checked = True Then
            MsgBox("LA NOTA DE INGRESO ESTA ANULADA NO SE PUEDE EDITAR")
        Else
            Dim sql102213 As String = "SELECT cierre_3 FROM custom_vianny.dbo.Cie0300 A  Where a.ccia ='" + Label14.Text + "' AND A.talm_3 ='" + ComboBox3.SelectedValue.ToString + "' AND A.fcie_3='" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "'"
            Dim cmd102213 As New SqlCommand(sql102213, conx)
            Rs1 = cmd102213.ExecuteReader()
            Dim jo As Integer
            If Rs1.Read() Then
                jo = Rs1(0)
            Else
                jo = 0
            End If
            Rs1.Close()
            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("QUIERES EDITAR ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                If jo = 0 Then
                    TextBox13.Enabled = True
                    DataGridView1.Enabled = True
                    DataGridView1.Columns(7).ReadOnly = False

                    TextBox9.Enabled = True
                    DateTimePicker1.Enabled = True
                    'TextBox12.Enabled = True
                    TextBox12.ReadOnly = False
                    TextBox4.Enabled = False
                    TextBox20.Enabled = True
                    TextBox5.Enabled = False
                    Button2.Enabled = True
                    Button5.Enabled = True
                    ComboBox1.Enabled = True
                    ComboBox2.Enabled = True
                    ComboBox4.Enabled = True
                    TextBox16.Enabled = True
                    TextBox17.Enabled = True
                    TextBox8.Enabled = True
                    Label15.Text = "1"

                Else
                    MsgBox("DIA " & DateTimePicker1.Value & " CERRADO,NO SE PUEDE EDITAR LA INFORMACION - CONSULTE CON EL AREA CONTABLE ")
                End If
                'If DT30.Rows.Count <> 0 Then
                '    DataGridView1.Rows.Add(1)
                '    DataGridView8.DataSource = DT30
                '    Dim num, num2, i17 As Integer
                '    Dim num10, num11, num12, num14 As Integer
                '    num = 0

                '    If DataGridView1.Rows(0).Cells(0).Value = 0 Then

                '        DataGridView1.Rows(0).Cells(0).Value = num + 1
                '        Select Case DataGridView1.Rows(0).Cells(0).Value.ToString.Length
                '            Case "1" : DataGridView1.Rows(0).Cells(0).Value = "00" & "" & DataGridView1.Rows(0).Cells(0).Value
                '            Case "2" : DataGridView1.Rows(0).Cells(0).Value = "0" & "" & DataGridView1.Rows(0).Cells(0).Value
                '            Case "3" : DataGridView1.Rows(0).Cells(0).Value = DataGridView1.Rows(0).Cells(0).Value

                '        End Select
                '        DataGridView1.Rows(0).Cells(1).Value = DataGridView8.Rows(0).Cells(0).Value
                '        DataGridView1.Rows(0).Cells(2).Value = DataGridView8.Rows(0).Cells(1).Value
                '        DataGridView1.Rows(0).Cells(3).Value = DataGridView8.Rows(0).Cells(2).Value

                '        DataGridView1.Rows(0).Cells(7).Value = TextBox13.Text
                '        DataGridView1.Rows(0).Cells(6).Value = "KGS"
                '        'TextBox14.Text = DataGridView2.Rows(0).Cells(2).Value
                '        'TextBox15.Text = DataGridView2.Rows(0).Cells(5).Value
                '        CType(DataGridView1.Columns(7), DataGridViewTextBoxColumn).MaxInputLength = 6

                '        DataGridView1.Rows(0).Cells(1).ReadOnly = False



                '        DataGridView8.DataSource = ""
                '        TextBox13.Text = ""

                '    Else
                '        num10 = DataGridView1.Rows.Count
                '        For i = 1 To num10 - 1
                '            If Convert.ToInt16(DataGridView1.Rows(i).Cells(0).Value) < Convert.ToInt16(DataGridView1.Rows(i - 1).Cells(0).Value) Then
                '                num2 = Convert.ToInt16(DataGridView1.Rows(i - 1).Cells(0).Value) + 1
                '                i17 = i

                '            End If


                '        Next

                '        If i17 = "50" Then
                '            MsgBox("SOLO SE PERMITE 50 REGISTROS")
                '        Else
                '            For i1 = 0 To i17 - 1


                '                num10 = Trim(Mid(TextBox13.Text, 1, 6))
                '                num11 = Trim(DataGridView1.Rows(i1).Cells(7).Value)
                '                num14 = Mid(num11, 1, 6)

                '            Next
                '            If num12 <> 1 Then
                '                DataGridView1.Rows(i17).Cells(0).Value = num2
                '                Select Case DataGridView1.Rows(i17).Cells(0).Value.ToString.Length
                '                    Case "1" : DataGridView1.Rows(i17).Cells(0).Value = "00" & "" & DataGridView1.Rows(i17).Cells(0).Value
                '                    Case "2" : DataGridView1.Rows(i17).Cells(0).Value = "0" & "" & DataGridView1.Rows(i17).Cells(0).Value
                '                    Case "3" : DataGridView1.Rows(i17).Cells(0).Value = DataGridView1.Rows(i17).Cells(0).Value

                '                End Select
                '                DataGridView1.Rows(i17).Cells(1).Value = DataGridView8.Rows(0).Cells(0).Value
                '                DataGridView1.Rows(i17).Cells(2).Value = DataGridView8.Rows(0).Cells(1).Value
                '                DataGridView1.Rows(i17).Cells(3).Value = DataGridView8.Rows(0).Cells(2).Value
                '                DataGridView1.Rows(i17).Cells(7).Value = TextBox13.Text
                '                DataGridView1.Rows(i17).Cells(6).Value = "KGS"
                '                DataGridView1.Rows(i17).Selected = True
                '                DataGridView1.Rows(i17).Cells(1).ReadOnly = False

                '                DataGridView1.Rows(i17).Cells(4).ReadOnly = False
                '                DataGridView1.Rows(i17).Cells(6).ReadOnly = True
                '                TextBox13.Text = ""
                '                DataGridView8.DataSource = ""
                '                For i = 0 To i17 - 1
                '                    DataGridView1.Rows(i).Selected = False
                '                Next
                '            End If
                '        End If


                '    End If

                'Else
                '    MsgBox("LA PARTIDA INGRESADA NO EXISTE")
                '    TextBox13.Text = ""
                '    'DataGridView8.DataSource = ""
                'End If
            End If
        End If

    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim valor, valor1, total As Integer
        valor = 0
        valor1 = 0
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "OC/OS" Then
            Dim sql1022134 As String = "select COUNT(ncom_3a) from custom_vianny.DBO.lap0300 where ncom_3a ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(9).Value) + "' and ccia_3a ='" + Label14.Text + "' and linea_3a ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value) + "'"
            Dim cmd1022134 As New SqlCommand(sql1022134, conx)
            Rsr34 = cmd1022134.ExecuteReader()
            If Rsr34.Read() = True Then
                valor = Rsr34(0)
            Else
                valor = 0
                'If Rsr34(0) > 0 Then
                '    DataGridView1.Rows(e.RowIndex).Cells(9).Value = ""
                'Else
                '    MsgBox("EL ARTICULO NO PERTENECE A LA ORDEN INGRESADA, REVISE SU ORDEN DE COMPRA O SERVICIO")
                'End If
            End If
            Rsr34.Close()
            Dim sql10221348 As String = "select COUNT(ncom_4a) from custom_vianny.DBO.lap0400 where ncom_4a ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(9).Value) + "' and ccia_4a ='" + Label14.Text + "' and linea_4a ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value) + "'"
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
                DataGridView1.Rows(e.RowIndex).Cells(9).Value = ""
                MsgBox("EL ARTICULO NO PERTENECE A LA ORDEN INGRESADA, REVISE SU ORDEN DE COMPRA O SERVICIO")
            End If
        End If


    End Sub
    Sub validar_ingresos(columna As Integer, fila As Integer)
        Dim columnIndex As Integer = 5 ' Puedes cambiar este valor al índice de la columna que necesitas
        Dim columnIndex1 As Integer = 4
        Dim columnIndex2 As Integer = 7
        Dim columnIndex3 As Integer = 11
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
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        validar_ingresos(e.ColumnIndex, e.RowIndex)
    End Sub

    Private Sub DataGridView1_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEnter
        validar_ingresos(e.ColumnIndex, e.RowIndex)
    End Sub

    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        ' Verifica si la celda actual pertenece a la columna que deseas validar
        Dim columnIndex As Integer = 5 ' Cambia este valor al índice de la columna que quieres validar
        Dim columnIndex1 As Integer = 4 ' Cambia este valor al índice de la columna que quieres validar
        If DataGridView1.CurrentCell.ColumnIndex = columnIndex AndAlso TypeOf e.Control Is TextBox Then
            ' Agrega el controlador de eventos KeyPress al cuadro de texto de edición
            AddHandler DirectCast(e.Control, TextBox).KeyPress, AddressOf DataGridView1_KeyPress
        End If
        If DataGridView1.CurrentCell.ColumnIndex = columnIndex1 AndAlso TypeOf e.Control Is TextBox Then
            ' Agrega el controlador de eventos KeyPress al cuadro de texto de edición
            AddHandler DirectCast(e.Control, TextBox).KeyPress, AddressOf DataGridView1_KeyPress
        End If
    End Sub

    Private Sub DataGridView1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles DataGridView1.KeyPress
        Dim currentCell As DataGridViewCell = DataGridView1.CurrentCell
        If currentCell.ColumnIndex = 5 Then
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
            DataGridView1.Rows(index).Cells(5).Value = fila.Cells("CANT").Value
            DataGridView1.Rows(index).Cells(6).Value = fila.Cells("UND").Value
            DataGridView1.Rows(index).Cells(9).Value = fila.Cells("OC").Value
            DataGridView1.Rows(index).Cells(12).Value = fila.Cells("valor").Value
            ' Continúa para todas las columnas que desees transferir
        Next
    End Sub
    Public Sub EstablecerAlmacen(almacen As String)
        ' Verificar si el valor existe en el ComboBox antes de asignarlo
        alamcenes(almacen)
    End Sub
End Class