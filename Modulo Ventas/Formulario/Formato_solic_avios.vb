Imports System.Data.SqlClient
Imports System.Net.Mail
Public Class Formato_solic_avios
    Public comando As SqlCommand
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim Rsr2, Rsr22, Rsr3, Rsr34, Rsr341, Rsr346, Rsr29, Rsr244, Rsr223, Rsr227, Rsr3465, Rsr27 As SqlDataReader
    Dim f As String
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        abrir()
        DataGridView1.Rows.Clear()

        TextBox3.Enabled = True




        '-------------------------
        If CheckBox5.Checked = True Then

            Dim sql10227 As String = "select count(ncom_3) from custom_vianny.dbo.qag0300 where program_3 ='" + Trim(TextBox15.Text) + "' and ccia ='" + Label7.Text + "'"
            Dim cmd10227 As New SqlCommand(sql10227, conx)
            Rsr227 = cmd10227.ExecuteReader()
            Rsr227.Read()

            If Rsr227(0) > 0 Then
                Rsr227.Close()
                Dim f As String
                If CheckBox1.Checked = True Then
                    f = "1"
                Else
                    f = "2"
                End If
                Dim sql1022 As String = "EXEC CaeSoft_ReporteMatrizConsumoPrueba  '" + Label7.Text + "','" + Trim(TextBox15.Text) + "','" + Trim(TextBox15.Text) + "',NULL,'01','" + f + "' "
                Dim cmd1022 As New SqlCommand(sql1022, conx)
                Rsr22 = cmd1022.ExecuteReader()
                Dim i51 As Integer
                i51 = 0
                While Rsr22.Read()

                    DataGridView1.Rows.Add()

                    DataGridView1.Rows(i51).Cells(1).Value = i51 + 1
                    DataGridView1.Rows(i51).Cells(2).Value = Rsr22(0)
                    DataGridView1.Rows(i51).Cells(3).Value = Rsr22(1)
                    DataGridView1.Rows(i51).Cells(5).Value = 0
                    DataGridView1.Rows(i51).Cells(6).Value = Rsr22(2)
                    DataGridView1.Rows(i51).Cells(7).Value = 0
                    DataGridView1.Rows(i51).Cells(9).Value = Rsr22(3)

                    i51 = i51 + 1
                End While
                DataGridView1.Columns(4).Visible = False
                DataGridView1.Columns(6).Visible = False
                Rsr22.Close()

                Dim sql10223 As String = "select  program_3 ,c.nomb_10,descri_3,sum(cants_3) as solicitado,sum(cantp_3) as programado,estilo_3 from custom_vianny.dbo.qag0300 q inner join custom_vianny.dbo.cag1000 c on q.fich_3 = c.fich_10 where program_3 ='" + Trim(TextBox15.Text) + "' and q.ccia ='" + Label7.Text + "' group by  program_3 ,c.nomb_10,descri_3,estilo_3"
                Dim cmd10223 As New SqlCommand(sql10223, conx)
                Rsr223 = cmd10223.ExecuteReader()
                If Rsr223.Read() Then
                    TextBox13.Text = Rsr223(0)
                    TextBox9.Text = Rsr223(0)
                    TextBox7.Text = Rsr223(1)
                    TextBox8.Text = Rsr223(2)
                    TextBox11.Text = Rsr223(3)
                    TextBox12.Text = Rsr223(4)
                    TextBox10.Text = Rsr223(5)
                End If
                Rsr223.Close()
                DataGridView1.EditMode = DataGridViewEditMode.EditOnEnter
                DataGridView1.CurrentCell = DataGridView1.Rows(0).Cells(8)
                DataGridView1.CurrentCell.Selected = True
                TextBox5.Text = ComboBox1.SelectedValue.ToString
            Else
                MsgBox("LA PO NO EXISTE")
                TextBox15.Text = ""
                TextBox15.Select()
                Rsr227.Close()

            End If
            Rsr227.Close()


        Else
            'Dim respuesta As DialogResult
            'respuesta = MessageBox.Show("NO ESTA SELECCIONANDO NINGUN CORTE ES CORRECTO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            'If (respuesta = Windows.Forms.DialogResult.Yes) Then
            Dim sql10227 As String = "select count(ncom_3) from custom_vianny.dbo.qag0300 where ncom_3 ='" + Trim(TextBox6.Text) + "' and ccia ='" + Label7.Text + "'"
                Dim cmd10227 As New SqlCommand(sql10227, conx)
                Rsr227 = cmd10227.ExecuteReader()
                Rsr227.Read()
                If Rsr227(0) > 0 Then
                    Rsr227.Close()


                    Dim sql102217 As String = "SELECT isnull( SUM(Canti_3q),0) FROM custom_vianny.dbo.Qaq300cc A WHERE Empr_3q = '" + Trim(Label7.Text) + "' AND Pedido_3q = '" + Trim(TextBox6.Text) + "'    and corte_3q ='" + ComboBox1.SelectedValue.ToString + "'"
                    Dim cmd102217 As New SqlCommand(sql102217, conx)
                Rsr27 = cmd102217.ExecuteReader()
                If Rsr27.Read() = True Then
                    TextBox14.Text = Rsr27(0)
                End If
                Rsr27.Close()

                Dim fg As Char
                If CheckBox1.Checked = True Then
                        fg = "1"
                    Else
                        fg = "2"
                    End If
                    Dim sql1022 As String = "EXEC CaeSoft_ReporteMatrizConsumo '" + Label7.Text + "','" + TextBox6.Text + "','" + TextBox6.Text + "',NULL,'01','" + fg + "'"
                    Dim cmd1022 As New SqlCommand(sql1022, conx)
                    Rsr22 = cmd1022.ExecuteReader()
                    Dim i51 As Integer
                    i51 = 0
                    While Rsr22.Read()
                        TextBox2.Enabled = False
                        TextBox7.Text = Rsr22(1)
                        TextBox8.Text = Rsr22(5)
                        TextBox9.Text = Rsr22(19)
                        TextBox10.Text = Rsr22(6)
                        TextBox11.Text = Rsr22(3)
                        TextBox12.Text = Rsr22(4)
                        TextBox13.Text = TextBox6.Text
                        DataGridView1.Rows.Add()

                        DataGridView1.Rows(i51).Cells(1).Value = i51 + 1
                        DataGridView1.Rows(i51).Cells(2).Value = Rsr22(11)
                        DataGridView1.Rows(i51).Cells(3).Value = Rsr22(12)
                        DataGridView1.Rows(i51).Cells(5).Value = Rsr22(14)
                        DataGridView1.Rows(i51).Cells(6).Value = Rsr22(15)
                        DataGridView1.Rows(i51).Cells(4).Value = Rsr22(21)

                        i51 = i51 + 1
                    End While
                    Rsr22.Close()


                    If CheckBox4.Checked = True Then

                        Dim sql10221 As String = "SELECT isnull( SUM(Canti_3q),0) FROM custom_vianny.dbo.Qaq300cc A WHERE Empr_3q = '" + Trim(Label7.Text) + "' AND Pedido_3q = '" + Trim(TextBox6.Text) + "'    and corte_3q ='" + ComboBox1.SelectedValue.ToString + "'"
                        Dim cmd10221 As New SqlCommand(sql10221, conx)
                        Rsr2 = cmd10221.ExecuteReader()


                        If Rsr2.Read() = True Then
                            For i = 0 To DataGridView1.Rows.Count - 1
                            If Trim(DataGridView1.Rows(i).Cells(4).Value) = "" Then
                                TextBox14.Text = Rsr2(0)
                                DataGridView1.Rows(i).Cells(7).Value = Rsr2(0)
                                DataGridView1.Rows(i).Cells(9).Value = Convert.ToString(Convert.ToDouble(DataGridView1.Rows(i).Cells(5).Value) * Convert.ToDouble(DataGridView1.Rows(i).Cells(7).Value))
                            End If
                        Next
                        End If

                        Rsr2.Close()



                        For i = 0 To DataGridView1.Rows.Count - 1
                            If Trim(DataGridView1.Rows(i).Cells(4).Value) <> "" Then
                                Dim sql102244 As String = " SELECT isnull(SUM(Canti_3q),0) FROM custom_vianny.dbo.Qaq300cc A WHERE Empr_3q = '" + Trim(Label7.Text) + "' AND Pedido_3q = '" + Trim(TextBox6.Text) + "' AND Corte_3q = ISNULL('" + ComboBox1.SelectedValue.ToString + "',Corte_3q) AND Item_3q = ISNULL('01',Item_3q)  and talla_3q ='" + Trim(DataGridView1.Rows(i).Cells(4).Value) + "'"
                                Dim cmd102244 As New SqlCommand(sql102244, conx)
                                Rsr244 = cmd102244.ExecuteReader()
                                If Rsr244.Read() = True Then
                                    DataGridView1.Rows(i).Cells(7).Value = Rsr244(0)
                                    DataGridView1.Rows(i).Cells(9).Value = Convert.ToString(Convert.ToDouble(DataGridView1.Rows(i).Cells(5).Value) * Convert.ToDouble(DataGridView1.Rows(i).Cells(7).Value))
                                End If
                                Rsr244.Close()
                            End If
                        Next



                    Else
                        Dim sql102219 As String = "SELECT isnull( SUM(Canti_3q),0) FROM custom_vianny.dbo.Qaq300cc A WHERE Empr_3q = '" + Trim(Label7.Text) + "' AND Pedido_3q = '" + Trim(TextBox6.Text) + "'"
                        Dim cmd102219 As New SqlCommand(sql102219, conx)
                        Rsr29 = cmd102219.ExecuteReader()


                        If Rsr29.Read() = True Then
                            For i = 0 To DataGridView1.Rows.Count - 1
                                If Trim(DataGridView1.Rows(i).Cells(4).Value) = "" Then
                                TextBox14.Text = Rsr29(0)
                                DataGridView1.Rows(i).Cells(7).Value = Rsr29(0)
                                    DataGridView1.Rows(i).Cells(9).Value = Convert.ToString(Convert.ToDouble(DataGridView1.Rows(i).Cells(5).Value) * Convert.ToDouble(DataGridView1.Rows(i).Cells(7).Value))
                                End If

                            Next
                        End If

                        Rsr29.Close()

                        For i = 0 To DataGridView1.Rows.Count - 1
                            If Trim(DataGridView1.Rows(i).Cells(4).Value) <> "" Then
                                Dim sql102244 As String = " SELECT isnull( SUM(Canti_3q),0) FROM custom_vianny.dbo.Qaq300cc A WHERE Empr_3q = '" + Trim(Label7.Text) + "' AND Pedido_3q = '" + Trim(TextBox6.Text) + "'    and talla_3q ='" + Trim(DataGridView1.Rows(i).Cells(4).Value) + "'"
                                Dim cmd102244 As New SqlCommand(sql102244, conx)
                                Rsr244 = cmd102244.ExecuteReader()
                            If Rsr244.Read() = True Then

                                DataGridView1.Rows(i).Cells(7).Value = Rsr244(0)
                                DataGridView1.Rows(i).Cells(9).Value = Convert.ToString(Convert.ToDouble(DataGridView1.Rows(i).Cells(5).Value) * Convert.ToDouble(DataGridView1.Rows(i).Cells(7).Value))
                            End If
                            Rsr244.Close()
                            Else

                            End If
                        Next
                    End If
                    DataGridView1.EditMode = DataGridViewEditMode.EditOnEnter
                'DataGridView1.CurrentCell = DataGridView1.Rows(0).Cells(8)
                'DataGridView1.CurrentCell = DataGridView1.Rows(0).Cells(9)
                'DataGridView1.CurrentCell.Selected = True
                TextBox5.Text = ComboBox1.SelectedValue.ToString
                Else
                    MsgBox("LA OP NO EXISTE")
                    TextBox6.Text = ""
                    TextBox6.Select()
                    Rsr227.Close()
                End If

                'End If
            End If



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
                'Select Case VAL.ToString.Length
                '    Case "1" : DataGridView1.Rows(i1).Cells(1).Value = "00" & "" & VAL
                '    Case "2" : DataGridView1.Rows(i1).Cells(1).Value = "0" & "" & VAL
                DataGridView1.Rows(i1).Cells(1).Value = VAL
                'End Select
            Next
        End If
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs)

    End Sub
    Dim loDataTable As New DataTable
    Sub llenar_combo_box3()
        Try
            loDataTable.Rows.Clear()
            Dim lsQuery As String = "select ocorte_3cg from custom_vianny.dbo.Qag300Cc where pedido_3cg ='" + Trim(TextBox6.Text) + "'"

            Dim loDataAdapter As New SqlDataAdapter(lsQuery, conx)
            loDataAdapter.Fill(loDataTable)
            ComboBox1.DataSource = loDataTable

            ComboBox1.DisplayMember = "ocorte_3cg"
            ComboBox1.ValueMember = "ocorte_3cg"
            ComboBox1.DropDownWidth = 90
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Formato_solic_avios_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim sql102213 As String = "exec soiget_ultimo_requerimiento_avios '" + Label7.Text + "','" + TextBox1.Text + "'"
        Dim cmd102213 As New SqlCommand(sql102213, conx)
        Rsr3 = cmd102213.ExecuteReader()
        If Rsr3.Read() = True Then
            TextBox2.Text = Trim(Rsr3(0))
        End If

        Rsr3.Close()

        TextBox6.Enabled = False

        TextBox3.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        CheckBox6.Checked = False
    End Sub
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
    Sub enviar_correO()
        abrir()
        Dim Rs1 As SqlDataReader
        Dim sql1 As String = "select correo from lista_correos where formulario ='SOLICITUD_REQUERIMI '"
        Dim cmd1 As New SqlCommand(sql1, conx)
        Rs1 = cmd1.ExecuteReader()
        Rs1.Read()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim fk As New fnopedido
        Dim jj As New vnapedido
        Dim OPP As String
        If Label20.Text = "0" Then
            OPP = "CREADO"
        Else
            OPP = "MODIFICADO"
        End If
        Dim corre As String
        'jj.gvendedor = Label1.Text
        corre = fk.buscar_correo(jj)
        Dim kj As String
        If CheckBox4.Checked = True Then
            kj = Trim(ComboBox1.SelectedValue.ToString)
        Else
            kj = "TODOS"
        End If


        Rs1.Close()

    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        abrir()

        Dim agregar As String = "delete from requerimiento_avios_cabecera where id_requeminieto ='" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + "' and empresa ='" + Label7.Text + "' and periodo ='" + Label15.Text + "'"
        Dim agregar1 As String = "delete from requerimiento_avios_detalle where id_cabecera ='" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + "' and empresa ='" + Label7.Text + "' and periodo ='" + Label15.Text + "'"
        ELIMINAR(agregar)
        ELIMINAR(agregar1)
        Dim cmd16 As New SqlCommand("insert into requerimiento_avios_cabecera(id_requeminieto,area,op,po,cliente,producto,estilo,c_sol,c_prog,corte,empresa,periodo,fcom,estado,cant_corta)
								                                      values (@id_requeminieto,@area,@op,@po,@cliente,@producto,@estilo,@c_sol,@c_prog,@corte,@empresa,@periodo,@fcom,@estado,@cant_corta)", conx)
        cmd16.Parameters.AddWithValue("@id_requeminieto", Trim(TextBox1.Text) & Trim(TextBox2.Text))
        If CheckBox1.Checked = True Then
            cmd16.Parameters.AddWithValue("@area", "1")
        Else
            If CheckBox2.Checked = True Then
                cmd16.Parameters.AddWithValue("@area", "2")


            End If
        End If

        If CheckBox5.Checked = True Then
            cmd16.Parameters.AddWithValue("@po", Trim(TextBox13.Text))
        Else
            cmd16.Parameters.AddWithValue("@po", Trim(TextBox9.Text))
        End If
        cmd16.Parameters.AddWithValue("@op", Trim(TextBox13.Text))
        'cmd16.Parameters.AddWithValue("@po", Trim(TextBox9.Text))
        cmd16.Parameters.AddWithValue("@cliente", Trim(TextBox7.Text))
        cmd16.Parameters.AddWithValue("@producto", Trim(TextBox8.Text))
        cmd16.Parameters.AddWithValue("@estilo", Trim(TextBox10.Text))
        cmd16.Parameters.AddWithValue("@c_sol", Convert.ToDouble(TextBox11.Text))
        cmd16.Parameters.AddWithValue("@c_prog", Convert.ToDouble(TextBox12.Text))
        If CheckBox4.Checked = True Then
            cmd16.Parameters.AddWithValue("@corte", TextBox5.Text)
        Else
            cmd16.Parameters.AddWithValue("@corte", "")
        End If
        cmd16.Parameters.AddWithValue("@empresa", Label7.Text)
        cmd16.Parameters.AddWithValue("@periodo", Label15.Text)
        cmd16.Parameters.AddWithValue("@fcom", DateTimePicker1.Value)
        cmd16.Parameters.AddWithValue("@estado", "0")
        cmd16.Parameters.AddWithValue("@cant_corta", TextBox14.Text)
        cmd16.ExecuteNonQuery()
        Dim p As Integer
        p = DataGridView1.Rows.Count
        For i = 0 To p - 1
            'If DataGridView1.Rows(i).Visible = True Then
            ' If DataGridView1.Rows(i).Cells(0).Value = True Then
            abrir()
            Dim cmd15 As New SqlCommand("INSERT INTO requerimiento_avios_detalle (id_cabecera,items,linea,detalle,factor,um,cantidad,total,empresa,periodo,estado,estado1,talla,correlativo) 
								VALUES (@id_cabecera,@items,@linea,@detalle,@factor,@um,@cantidad,@total,@empresa,@periodo,@estado,@estado1,@talla,@correlativo)", conx)
            cmd15.Parameters.AddWithValue("@id_cabecera", Trim(TextBox1.Text) + Trim(TextBox2.Text))
                    cmd15.Parameters.AddWithValue("@items", Trim(DataGridView1.Rows(i).Cells(1).Value))
                    cmd15.Parameters.AddWithValue("@linea", Trim(DataGridView1.Rows(i).Cells(2).Value))
                    cmd15.Parameters.AddWithValue("@detalle", Trim(DataGridView1.Rows(i).Cells(3).Value))
            If Trim(DataGridView1.Rows(i).Cells(5).Value) = 0 Then
                cmd15.Parameters.AddWithValue("@factor", 0)
            Else
                cmd15.Parameters.AddWithValue("@factor", Convert.ToDouble(DataGridView1.Rows(i).Cells(5).Value))
                End If
                'cmd15.Parameters.AddWithValue("@factor", Convert.ToDouble(DataGridView1.Rows(i).Cells(3).Value))
                cmd15.Parameters.AddWithValue("@um", Trim(DataGridView1.Rows(i).Cells(6).Value))
            If Trim(DataGridView1.Rows(i).Cells(7).Value) = 0 Then
                cmd15.Parameters.AddWithValue("@cantidad", 0)
            Else
                cmd15.Parameters.AddWithValue("@cantidad", Convert.ToDouble(DataGridView1.Rows(i).Cells(7).Value))
                End If

            cmd15.Parameters.AddWithValue("@total", Convert.ToDouble(DataGridView1.Rows(i).Cells(9).Value))
            cmd15.Parameters.AddWithValue("@empresa", Trim(Label7.Text))
                    cmd15.Parameters.AddWithValue("@periodo", Trim(Label15.Text))
                    cmd15.Parameters.AddWithValue("@estado", "0")
            cmd15.Parameters.AddWithValue("@estado1", "0")
            cmd15.Parameters.AddWithValue("@talla", Trim(DataGridView1.Rows(i).Cells(4).Value))
            cmd15.Parameters.AddWithValue("@correlativo", "0")
            cmd15.ExecuteNonQuery()
            '   End If

            '  End If

        Next
        MsgBox("LA INFORMACION SE GUARDO CORRECTAMENTE")
        MessageBox.Show("Datos registrados correctamente ", "Guardando registros", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Dim respuesta As DialogResult

        respuesta = MessageBox.Show("DESEA IMPRIMIR EL REQUERIMIENTO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            rPT_SOLICITUD.TextBox1.Text = TextBox1.Text & TextBox2.Text
            rPT_SOLICITUD.TextBox2.Text = Label7.Text
            rPT_SOLICITUD.TextBox3.Text = Label15.Text
            rPT_SOLICITUD.ShowDialog()
        End If
        'enviar_correO()
        limpiar()
        TextBox2.Enabled = True
        TextBox2.Select()
    End Sub

    Sub limpiar()

        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""

        DataGridView1.Rows.Clear()
        CheckBox3.Checked = False
        ComboBox1.DataSource = Nothing
        Dim sql102213 As String = "exec soiget_ultimo_requerimiento_avios '" + Label7.Text + "','" + TextBox1.Text + "'"
        Dim cmd102213 As New SqlCommand(sql102213, conx)
        Rsr3 = cmd102213.ExecuteReader()
        If Rsr3.Read() = True Then
            TextBox2.Text = Trim(Rsr3(0))
        End If

        Rsr3.Close()


        TextBox6.Enabled = False
        ComboBox1.Enabled = False

        Button2.Enabled = False
        Button3.Enabled = False
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case TextBox6.Text.Length

                Case "1" : TextBox6.Text = "010000000" & "" & TextBox6.Text
                Case "2" : TextBox6.Text = "01000000" & "" & TextBox6.Text
                Case "3" : TextBox6.Text = "0100000" & "" & TextBox6.Text
                Case "4" : TextBox6.Text = "010000" & "" & TextBox6.Text
                Case "5" : TextBox6.Text = "01000" & "" & TextBox6.Text
                Case "6" : TextBox6.Text = "0100" & "" & TextBox6.Text
                Case "7" : TextBox6.Text = "010" & "" & TextBox6.Text
                Case "8" : TextBox6.Text = "01" & "" & TextBox6.Text
            End Select
            abrir()
            llenar_combo_box3()
            Button1.Enabled = True
            CheckBox4.Enabled = True
            CheckBox4.Checked = True
            Button1.Select()
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If Label14.Text = "ANULADO" Then
            MsgBox("EL REQUERIMIENTO ESTA ANULADO NO PUEDE MODIFICARSE")
        Else
            TextBox2.Enabled = False
            TextBox3.Enabled = True
            DateTimePicker1.Enabled = True

            Button2.Enabled = True
            DataGridView1.Enabled = True
            Button3.Enabled = True
            Label20.Text = "1"
            DataGridView1.Select()
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Sub enviar_corre_ANULADO()
        abrir()
        Dim Rs1 As SqlDataReader
        Dim sql1 As String = "select correo from lista_correos where formulario ='SOLICITUD_REQUERIMI '"
        Dim cmd1 As New SqlCommand(sql1, conx)
        Rs1 = cmd1.ExecuteReader()
        Rs1.Read()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim fk As New fnopedido
        Dim jj As New vnapedido
        Dim KJ As String
        If CheckBox4.Checked = True Then
            KJ = Trim(ComboBox1.SelectedValue.ToString)
        Else
            KJ = "TODOS"
        End If
        Dim corre As String
        'jj.gvendedor = Label1.Text
        corre = fk.buscar_correo(jj)

        Dim Mailz As String = "" &
        " <!DOCTYPE html>" &
        "<html xmlns='http://www.w3.org/1999/xhtml'>" &
        "<head>" &
        "    <title></title>" &
        "</head>" &
        "<body>
     <FONT COLOR='green'>INFORMACION DEL REQUERIMIENTO DE AVIOS - PRODUCCION</FONT><br/><br/>

<FONT COLOR='blue'>* ESTADO : </FONT> REQUERIMIENTO ANULADO <br/>
<FONT COLOR='blue'>* N° REQUERIMIENTO : </FONT>" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + " <br/>
<FONT COLOR='blue'>* AREA : </FONT> " + Trim(Label16.Text) + " <br/>
<FONT COLOR='blue'>* OP : </FONT>" + Trim(TextBox13.Text) + " <br/>
<FONT COLOR='blue'>* PO : </FONT>" + Trim(TextBox9.Text) + " <br/>
<FONT COLOR='blue'>* CORTE : </FONT>" + KJ + " <br/>
<FONT COLOR='blue'>* CLIENTE : </FONT>" + Trim(TextBox7.Text) + "<br/>
<FONT COLOR='blue'>* PRODUCTO : </FONT>" + Trim(TextBox8.Text) + "<br/>

<FONT COLOR='blue'>* FECHA ANULACION : </FONT> " + DateTimePicker1.Value + " <br/>
<FONT COLOR='blue'>* USUARIO QUE APROBO : </FONT>" + Trim(Label17.Text) + "<br/>

</body>

</html>"
        With message
            .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
            .To.Add(Rs1(0))
            .IsBodyHtml = True
            .Body = Mailz
            .Subject = "REQUERIMIENTO DE AVIOS N°" & Trim(TextBox1.Text) & Trim(TextBox2.Text) & " - OP  " & Trim(TextBox13.Text)
            .Priority = System.Net.Mail.MailPriority.Normal
        End With

        With smtp
            .EnableSsl = True
            .Port = "25"
            .Host = "mail.onehostingperu.com"
            .Credentials = New Net.NetworkCredential("admin@viannysac.com", "I9!?@ni2;+go")
            .Send(message)
        End With

        Try
            MessageBox.Show("Su mensaje de correo ha sido enviado.", "Correo enviado",
                             MessageBoxButtons.OK)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error al enviar correo",
                             MessageBoxButtons.OK)
        End Try

        Rs1.Close()
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            ComboBox1.Enabled = True
        Else
            ComboBox1.Enabled = False
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            TextBox6.Enabled = False
            TextBox15.Enabled = True
            TextBox15.Select()
            Button1.Enabled = True
        Else
            TextBox6.Enabled = True
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TextBox15.Text = TextBox13.Text
        TextBox6.Text = TextBox13.Text
        mostrar_resultado()
    End Sub

    Private Sub TextBox3_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = True Then
            MsgBox("paso")
            Label23.Text = "1"
        Else
            Label23.Text = "2"
        End If
    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        abrir()
        Dim cmd157 As New SqlCommand("delete from requerimiento_avios_cabecera where id_requeminieto =@id_requeminieto ", conx)
        cmd157.Parameters.AddWithValue("@id_requeminieto", Trim(TextBox1.Text) + Trim(TextBox2.Text))
        cmd157.ExecuteNonQuery()

        Dim cmd158 As New SqlCommand("delete from requerimiento_avios_detalle where id_cabecera =@id_requeminieto", conx)
        cmd158.Parameters.AddWithValue("@id_requeminieto", Trim(TextBox1.Text) + Trim(TextBox2.Text))
        cmd158.ExecuteNonQuery()
        MsgBox("SE ELIMINO EL REQUERIMIENTO SOLICITADO")

        limpiar()
        TextBox2.Enabled = True
        TextBox2.Select()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        limpiar()
        abrir()
        Dim sql102213 As String = "exec soiget_ultimo_requerimiento_avios '" + Label7.Text + "','" + TextBox1.Text + "'"
        Dim cmd102213 As New SqlCommand(sql102213, conx)
        Rsr3 = cmd102213.ExecuteReader()
        If Rsr3.Read() = True Then
            TextBox2.Text = Trim(Rsr3(0))
        End If

        Rsr3.Close()


        TextBox6.Enabled = False
        ComboBox1.Enabled = False

        Button2.Enabled = False
        Button3.Enabled = False
        ComboBox1.DataSource = Nothing
        TextBox2.Enabled = True
        TextBox2.Select()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        rPT_SOLICITUD.TextBox1.Text = TextBox1.Text & TextBox2.Text
        rPT_SOLICITUD.TextBox2.Text = Label7.Text
        rPT_SOLICITUD.TextBox3.Text = Label15.Text
        rPT_SOLICITUD.Show()
    End Sub


    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            abrir()
            DataGridView1.Rows.Clear()
            ComboBox1.Items.Clear()

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

            Dim ui As String
            If Trim(Label16.Text) = "CONFECCION" Then
                ui = "1"
            Else
                ui = "2"
            End If

            Dim vas As Integer
            Dim sql10221342 As String = "select COUNT(id_requeminieto) from requerimiento_avios_cabecera where id_requeminieto ='" + TextBox1.Text & TextBox2.Text + "' and area ='" + ui + "'"
            Dim cmd10221342 As New SqlCommand(sql10221342, conx)
            Rsr341 = cmd10221342.ExecuteReader()
            Rsr341.Read()
            vas = Rsr341(0)

            Rsr341.Close()

            If vas = 0 Then

                TextBox6.Enabled = True
                'ComboBox1.Enabled = True
                DataGridView1.Enabled = True
                Button2.Enabled = True
                Button3.Enabled = True

                TextBox6.Select()


            Else
                Dim sql1022134 As String = "select * from requerimiento_avios_cabecera where id_requeminieto ='" + TextBox1.Text & TextBox2.Text + "' and area ='" + ui + "'"
                Dim cmd1022134 As New SqlCommand(sql1022134, conx)
                Rsr34 = cmd1022134.ExecuteReader()
                If Rsr34.Read() = True Then
                    'TextBox3.Text = Rsr34(2)
                    'TextBox4.Text = Rsr34(3)
                    'TextBox5.Text = Rsr34(4)
                    If Rsr34(13) = 7 Then
                        Label14.Text = "ANULADO"
                    Else
                        Label14.Text = "ACTIVO"
                    End If
                    TextBox7.Text = Rsr34(4)
                    TextBox8.Text = Rsr34(5)
                    TextBox9.Text = Rsr34(3)
                    TextBox10.Text = Rsr34(6)
                    TextBox11.Text = Rsr34(7)
                    TextBox12.Text = Rsr34(8)
                    TextBox5.Text = Rsr34(9)
                    TextBox13.Text = Rsr34(2)
                    TextBox14.Text = Rsr34(15)
                    'TextBox14.Text = Rsr34(20)
                    DateTimePicker1.Value = Rsr34(12)
                    'DateTimePicker2.Value = Rsr34(16)
                    ComboBox1.Items.Add(Rsr34(9))
                    ComboBox1.SelectedIndex = 0
                    If Rsr34(9).ToString().Trim().Length = 0 Then
                        CheckBox4.Checked = False
                    Else
                        CheckBox4.Checked = True
                    End If
                    'If Rsr34(12).ToString.Trim.Length > 0 Then
                    '    CheckBox4.Checked = True
                    '    CheckBox4.Enabled = False
                    'End If
                    If Rsr34(13) > 1 Then
                        CheckBox3.Checked = True
                    Else
                        CheckBox3.Checked = False
                    End If

                End If

                Rsr34.Close()

                Dim sql10221346 As String = "select items,linea,detalle,factor,um,cantidad,total,TALLA from requerimiento_avios_detalle where id_cabecera ='" + TextBox1.Text & TextBox2.Text + "' order by convert(int,items)"
                Dim cmd10221346 As New SqlCommand(sql10221346, conx)
                Rsr346 = cmd10221346.ExecuteReader()
                Dim i51 As Integer
                i51 = 0
                Dim val As Double
                val = 0
                While Rsr346.Read()
                    DataGridView1.Rows.Add()
                    DataGridView1.Rows(i51).Cells(1).Value = Rsr346(0)
                    DataGridView1.Rows(i51).Cells(2).Value = Trim(Rsr346(1))
                    DataGridView1.Rows(i51).Cells(3).Value = Trim(Rsr346(2))
                    DataGridView1.Rows(i51).Cells(4).Value = Rsr346(7)
                    DataGridView1.Rows(i51).Cells(5).Value = Rsr346(3)
                    DataGridView1.Rows(i51).Cells(6).Value = Rsr346(4)
                    DataGridView1.Rows(i51).Cells(7).Value = Rsr346(5)
                    DataGridView1.Rows(i51).Cells(9).Value = Rsr346(6)
                    'If DataGridView1.Rows(i51).Cells(7).Value <> 0 Then

                    '    val = ((Rsr346(3) * Rsr346(6) * 100) / Rsr346(5)) - 100
                    '    DataGridView1.Rows(i51).Cells(8).Value = val
                    'End If


                    If Label14.Text = "ANULADO" Then
                        DataGridView1.Rows(i51).DefaultCellStyle.BackColor = Color.Red
                        DataGridView1.Enabled = False
                    Else
                        DataGridView1.Enabled = True
                    End If

                    i51 = i51 + 1
                End While

                Rsr346.Close()
                DataGridView1.Enabled = False
            End If

        Else
            If e.KeyCode = Keys.F1 Then
                Lista_Solicitud.Label2.Text = Trim(Label15.Text)
                Lista_Solicitud.Label3.Text = Trim(TextBox1.Text)
                Lista_Solicitud.Show()
            End If
        End If






    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "CORTADO" Then
            DataGridView1.Rows(e.RowIndex).Cells(9).Value = DataGridView1.Rows(e.RowIndex).Cells(5).Value * DataGridView1.Rows(e.RowIndex).Cells(7).Value
        Else
            If DataGridView1.Columns(e.ColumnIndex).HeaderText = "%" Then
                If (Label20.Text = "1" And Label23.Text = "2") Then

                Else
                    If (Label20.Text = "1" And Label23.Text = "1") Then


                        If DataGridView1.Rows(e.RowIndex).Cells(7).Value = 0 Then
                            DataGridView1.Rows(e.RowIndex).Cells(8).Value = ""
                            MsgBox("NO SE PUEDE AGRAGRA % A CANTIDADES CORTADAS 0")
                        Else
                            If Trim(DataGridView1.Rows(e.RowIndex).Cells(8).Value) = "" Then
                                'If DataGridView1.Rows(e.RowIndex).Cells(7).Value <> 0 Then
                                '    DataGridView1.Rows(e.RowIndex).Cells(9).Value = DataGridView1.Rows(e.RowIndex).Cells(5).Value * DataGridView1.Rows(e.RowIndex).Cells(7).Value
                                'End If

                            Else
                                DataGridView1.Rows(e.RowIndex).Cells(9).Value = (DataGridView1.Rows(e.RowIndex).Cells(5).Value * DataGridView1.Rows(e.RowIndex).Cells(7).Value) + (((DataGridView1.Rows(e.RowIndex).Cells(5).Value * DataGridView1.Rows(e.RowIndex).Cells(7).Value) * DataGridView1.Rows(e.RowIndex).Cells(8).Value) / 100)

                            End If
                        End If
                    Else

                        If DataGridView1.Rows(e.RowIndex).Cells(7).Value = 0 Then
                            DataGridView1.Rows(e.RowIndex).Cells(8).Value = ""
                            MsgBox("NO SE PUEDE AGRAGRA % A CANTIDADES CORTADAS 0")
                        Else
                            If Trim(DataGridView1.Rows(e.RowIndex).Cells(8).Value) = "" Then
                                'If DataGridView1.Rows(e.RowIndex).Cells(7).Value <> 0 Then
                                '    DataGridView1.Rows(e.RowIndex).Cells(9).Value = DataGridView1.Rows(e.RowIndex).Cells(5).Value * DataGridView1.Rows(e.RowIndex).Cells(7).Value
                                'End If

                            Else
                                DataGridView1.Rows(e.RowIndex).Cells(9).Value = (DataGridView1.Rows(e.RowIndex).Cells(5).Value * DataGridView1.Rows(e.RowIndex).Cells(7).Value) + (((DataGridView1.Rows(e.RowIndex).Cells(5).Value * DataGridView1.Rows(e.RowIndex).Cells(7).Value) * DataGridView1.Rows(e.RowIndex).Cells(8).Value) / 100)

                            End If
                        End If
                    End If


                End If




            End If

        End If
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.F1 Then
            pproductos.CCIA.Text = Label7.Text
            pproductos.ALMACEN.Text = 13
            pproductos.ANO.Text = Label15.Text
            pproductos.FECHA.Text = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
            pproductos.TextBox1.Text = "13"
            pproductos.Label3.Text = 5
            pproductos.ShowDialog()
        Else
            If e.KeyCode = Keys.Enter Then
                abrir()
                Dim sql102213465 As String = "select linea_17,nomb_17,unid_17 from custom_vianny.dbo.cag1700 where linea_17 ='" + TextBox3.Text + "' and ccia ='" + Label7.Text + "'"
                Dim cmd102213465 As New SqlCommand(sql102213465, conx)
                Rsr3465 = cmd102213465.ExecuteReader()
                Dim i51 As Integer
                i51 = DataGridView1.Rows.Count
                While Rsr3465.Read()

                    DataGridView1.Rows.Add()
                    DataGridView1.Rows(i51).Cells(1).Value = i51 + 1
                    DataGridView1.Rows(i51).Cells(2).Value = Rsr3465(0)
                    DataGridView1.Rows(i51).Cells(3).Value = Rsr3465(1)
                    DataGridView1.Rows(i51).Cells(5).Value = 0
                    DataGridView1.Rows(i51).Cells(6).Value = Rsr3465(2)
                    DataGridView1.Rows(i51).Cells(7).Value = 0
                    DataGridView1.Rows(i51).Cells(8).Value = 0
                    i51 = i51 + 1
                End While

                Rsr3465.Close()
                TextBox3.Text = ""
            End If
        End If

    End Sub

    Private Sub DataGridView1_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles DataGridView1.RowsAdded
        'DataGridView1.CurrentCell = DataGridView1.Rows(e.RowIndex).Cells(8).Value
        'DataGridView1.BeginEdit(True)
    End Sub

    Private Sub Button1_KeyDown(sender As Object, e As KeyEventArgs) Handles Button1.KeyDown

        If e.KeyCode = Keys.Enter Then

            mostrar_resultado()
        End If
    End Sub
    Sub mostrar_resultado()
        abrir()
        DataGridView1.Rows.Clear()

        TextBox3.Enabled = True
        '-------------------------
        If CheckBox5.Checked = True Then

            Dim sql10227 As String = "select count(ncom_3) from custom_vianny.dbo.qag0300 where program_3 ='" + Trim(TextBox15.Text) + "' and ccia ='" + Label7.Text + "'"
            Dim cmd10227 As New SqlCommand(sql10227, conx)
            Rsr227 = cmd10227.ExecuteReader()
            Rsr227.Read()

            If Rsr227(0) > 0 Then
                Rsr227.Close()
                Dim f As String
                If CheckBox1.Checked = True Then
                    f = "1"
                Else
                    f = "2"
                End If
                Dim sql1022 As String = "EXEC CaeSoft_ReporteMatrizConsumoPrueba  '" + Label7.Text + "','" + Trim(TextBox15.Text) + "','" + Trim(TextBox15.Text) + "',NULL,'01','" + f + "' "
                Dim cmd1022 As New SqlCommand(sql1022, conx)
                Rsr22 = cmd1022.ExecuteReader()
                Dim i51 As Integer
                i51 = 0
                While Rsr22.Read()

                    DataGridView1.Rows.Add()

                    DataGridView1.Rows(i51).Cells(1).Value = i51 + 1
                    DataGridView1.Rows(i51).Cells(2).Value = Rsr22(0)
                    DataGridView1.Rows(i51).Cells(3).Value = Rsr22(1)
                    DataGridView1.Rows(i51).Cells(5).Value = 0
                    DataGridView1.Rows(i51).Cells(6).Value = Rsr22(2)
                    DataGridView1.Rows(i51).Cells(7).Value = 0
                    DataGridView1.Rows(i51).Cells(9).Value = Rsr22(3)

                    i51 = i51 + 1
                End While
                DataGridView1.Columns(4).Visible = False
                DataGridView1.Columns(6).Visible = False
                Rsr22.Close()

                Dim sql10223 As String = "select  program_3 ,c.nomb_10,descri_3,sum(cants_3) as solicitado,sum(cantp_3) as programado,estilo_3 from custom_vianny.dbo.qag0300 q inner join custom_vianny.dbo.cag1000 c on q.fich_3 = c.fich_10 where program_3 ='" + Trim(TextBox15.Text) + "' and q.ccia ='" + Label7.Text + "' group by  program_3 ,c.nomb_10,descri_3,estilo_3"
                Dim cmd10223 As New SqlCommand(sql10223, conx)
                Rsr223 = cmd10223.ExecuteReader()
                If Rsr223.Read() Then
                    TextBox13.Text = Rsr223(0)
                    TextBox9.Text = Rsr223(0)
                    TextBox7.Text = Rsr223(1)
                    TextBox8.Text = Rsr223(2)
                    TextBox11.Text = Rsr223(3)
                    TextBox12.Text = Rsr223(4)
                    TextBox10.Text = Rsr223(5)
                End If
                Rsr223.Close()
            Else
                MsgBox("LA PO NO EXISTE")
                TextBox15.Text = ""
                TextBox15.Select()
                Rsr227.Close()

            End If
            Rsr227.Close()


        Else
            Dim sql10227 As String = "select count(ncom_3) from custom_vianny.dbo.qag0300 where ncom_3 ='" + Trim(TextBox13.Text) + "' and ccia ='" + Label7.Text + "'"
            Dim cmd10227 As New SqlCommand(sql10227, conx)
            Rsr227 = cmd10227.ExecuteReader()
            Rsr227.Read()
            If Rsr227(0) > 0 Then
                Rsr227.Close()
                Dim fg As Char
                If CheckBox1.Checked = True Then
                    fg = "1"
                Else
                    fg = "2"
                End If
                Dim sql1022 As String = "EXEC CaeSoft_ReporteMatrizConsumo '" + Label7.Text + "','" + TextBox13.Text + "','" + TextBox13.Text + "',NULL,'01','" + fg + "'"
                Dim cmd1022 As New SqlCommand(sql1022, conx)
                Rsr22 = cmd1022.ExecuteReader()
                Dim i51 As Integer
                i51 = 0
                While Rsr22.Read()
                    TextBox2.Enabled = False
                    TextBox7.Text = Rsr22(1)
                    TextBox8.Text = Rsr22(5)
                    TextBox9.Text = Rsr22(19)
                    TextBox10.Text = Rsr22(6)
                    TextBox11.Text = Rsr22(3)
                    TextBox12.Text = Rsr22(4)
                    TextBox13.Text = TextBox6.Text
                    DataGridView1.Rows.Add()

                    DataGridView1.Rows(i51).Cells(1).Value = i51 + 1
                    DataGridView1.Rows(i51).Cells(2).Value = Rsr22(11)
                    DataGridView1.Rows(i51).Cells(3).Value = Rsr22(12)
                    DataGridView1.Rows(i51).Cells(5).Value = Rsr22(14)
                    DataGridView1.Rows(i51).Cells(6).Value = Rsr22(15)
                    DataGridView1.Rows(i51).Cells(4).Value = Rsr22(21)



                    i51 = i51 + 1
                End While
                Rsr22.Close()


                If CheckBox4.Checked = True Then
                    'MsgBox(ComboBox1.Text)
                    Dim sql10221 As String = "SELECT isnull( SUM(cant_3c),0) as CANT FROM custom_vianny.dbo.Qap300Cc A  Where a.empr_3c + a.pedido_3c + a.color_3c='" + Trim(Label7.Text) & Trim(TextBox13.Text) & "01" + "' AND ocorte_3c ='" + ComboBox1.Text + "'  and item_3c ='01'"
                    Dim cmd10221 As New SqlCommand(sql10221, conx)
                    Rsr2 = cmd10221.ExecuteReader()


                    If Rsr2.Read() = True Then
                        For i = 0 To DataGridView1.Rows.Count - 1
                            If Trim(DataGridView1.Rows(i).Cells(4).Value) = "" Then
                                DataGridView1.Rows(i).Cells(7).Value = Rsr2(0)
                                DataGridView1.Rows(i).Cells(9).Value = Convert.ToString(Convert.ToDouble(DataGridView1.Rows(i).Cells(5).Value) * Convert.ToDouble(DataGridView1.Rows(i).Cells(7).Value))
                            End If
                        Next
                    End If

                    Rsr2.Close()



                    For i11 = 0 To DataGridView1.Rows.Count - 1
                        If Trim(DataGridView1.Rows(i11).Cells(4).Value) <> "" Then

                            Dim sql102244 As String = " SELECT isnull(SUM(Canti_3q),0) FROM custom_vianny.dbo.Qaq300cc A WHERE Empr_3q = '" + Trim(Label7.Text) + "' AND Pedido_3q = '" + Trim(TextBox13.Text) + "' AND Corte_3q = ISNULL('" + ComboBox1.Text + "',Corte_3q) AND Item_3q = ISNULL('01',Item_3q)  and talla_3q ='" + Trim(DataGridView1.Rows(i11).Cells(4).Value) + "'"
                            Dim cmd102244 As New SqlCommand(sql102244, conx)
                            Rsr244 = cmd102244.ExecuteReader()
                            If Rsr244.Read() = True Then
                                DataGridView1.Rows(i11).Cells(7).Value = Rsr244(0)
                                DataGridView1.Rows(i11).Cells(9).Value = Convert.ToString(Convert.ToDouble(DataGridView1.Rows(i11).Cells(5).Value) * Convert.ToDouble(DataGridView1.Rows(i11).Cells(7).Value))
                            End If
                            Rsr244.Close()
                        End If
                    Next



                Else
                    Dim sql102219 As String = "SELECT isnull( SUM(cant_3c),0) as CANT FROM custom_vianny.dbo.Qap300Cc A  Where a.empr_3c + a.pedido_3c + a.color_3c='" + Trim(Label7.Text) & Trim(TextBox13.Text) & "01" + "' and item_3c ='01' "
                    Dim cmd102219 As New SqlCommand(sql102219, conx)
                    Rsr29 = cmd102219.ExecuteReader()


                    If Rsr29.Read() = True Then
                        For i = 0 To DataGridView1.Rows.Count - 1
                            If Trim(DataGridView1.Rows(i).Cells(4).Value) = "" Then
                                DataGridView1.Rows(i).Cells(7).Value = Rsr29(0)
                                DataGridView1.Rows(i).Cells(9).Value = Convert.ToString(Convert.ToDouble(DataGridView1.Rows(i).Cells(5).Value) * Convert.ToDouble(DataGridView1.Rows(i).Cells(7).Value))
                            End If

                        Next
                    End If

                    Rsr29.Close()

                    For i = 0 To DataGridView1.Rows.Count - 1
                        If Trim(DataGridView1.Rows(i).Cells(4).Value) <> "" Then
                            Dim sql102244 As String = " SELECT isnull( SUM(Canti_3q),0) FROM custom_vianny.dbo.Qaq300cc A WHERE Empr_3q = '" + Trim(Label7.Text) + "' AND Pedido_3q = '" + Trim(TextBox13.Text) + "'    and talla_3q ='" + Trim(DataGridView1.Rows(i).Cells(4).Value) + "'"
                            Dim cmd102244 As New SqlCommand(sql102244, conx)
                            Rsr244 = cmd102244.ExecuteReader()
                            If Rsr244.Read() = True Then
                                DataGridView1.Rows(i).Cells(7).Value = Rsr244(0)
                                DataGridView1.Rows(i).Cells(9).Value = Convert.ToString(Convert.ToDouble(DataGridView1.Rows(i).Cells(5).Value) * Convert.ToDouble(DataGridView1.Rows(i).Cells(7).Value))
                            End If
                            Rsr244.Close()
                        End If
                    Next
                End If

            Else
                MsgBox("LA OP NO EXISTE")
                TextBox6.Text = ""
                TextBox6.Select()
                Rsr227.Close()
            End If



        End If
        DataGridView1.EditMode = DataGridViewEditMode.EditOnEnter
        DataGridView1.CurrentCell = DataGridView1.Rows(0).Cells(9)
        DataGridView1.CurrentCell.Selected = True
    End Sub

    Private Sub TextBox15_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox15.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button1.Select()
        End If
    End Sub
    Dim Rsr21 As SqlDataReader
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        abrir()
        Dim sql1021 As String = "EXEC custom_vianny.dbo.CaeSoft_ReporteStockFisico '" + Trim(Label7.Text) + "','" + Trim(Label15.Text) + "','13','" + Trim(Label15.Text) + "0101','" + Trim(Label15.Text) + "1231','" + DataGridView1.Rows(e.RowIndex).Cells(2).Value + "','" + DataGridView1.Rows(e.RowIndex).Cells(2).Value + "',NULL, 1"
        Dim cmd1021 As New SqlCommand(sql1021, conx)
        Rsr21 = cmd1021.ExecuteReader()

        If Rsr21.Read() Then
            TextBox4.Text = Rsr21(10)
        Else
            TextBox4.Text = 0
        End If

        Rsr21.Close()
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        DataGridView1.BeginEdit(True)
        Dim J As Integer

        J = DataGridView1.Rows.Count
        For I = 0 To J - 1
            If Me.DataGridView1.CurrentRow.Index.ToString() = I Then

                DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.DarkCyan
                DataGridView1.Rows(I).DefaultCellStyle.ForeColor = Color.White
            Else
                DataGridView1.Rows(I).DefaultCellStyle.ForeColor = Color.Black
                DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.White
            End If
        Next
    End Sub

    Private Sub TextBox2_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox2.MouseClick

    End Sub
End Class