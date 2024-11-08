Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.Net.Mail
Public Class Hoja_Reclamos
    Dim DT, DT2 As New DataTable
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim ty, ty2, ty3 As New DataTable
    Public conn As SqlDataAdapter

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub

    'Sub llenar_combo_box2()
    '    Try

    '        conn = New SqlDataAdapter("SELECT codigo_ven,alias_ven FROM  custom_vianny.dbo.Vendedores WHERE  rpm_ven='2' and ccia_ven ='01'", conx)
    '        conn.Fill(ty2)
    '        ComboBox1.DataSource = ty2
    '        ComboBox1.DisplayMember = "alias_ven"
    '        ComboBox1.ValueMember = "codigo_ven"
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    'End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged

    End Sub

    Private Sub Hoja_Reclamos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        CORRELATIVO()
        '    llenar_combo_box2()

        abrir()
        Dim sql102 As String = "SELECT alias_ven FROM  custom_vianny.dbo.Vendedores WHERE codigo_ven ='" + Trim(TextBox14.Text) + "'"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        If Rsr2.Read() = True Then
            TextBox9.Text = Rsr2(0)
        End If

        Rsr2.Close()
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.F1 Then
            'Clientes.TextBox3.Text = "800"
            'Clientes.Show()

            Listar_Clientes.TextBox2.Text = "800"
            Listar_Clientes.TextBox3.Text = TextBox9.Text
            Listar_Clientes.Show()
        End If
    End Sub

    Private Sub Label19_Click(sender As Object, e As EventArgs) Handles Label19.Click

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'If TextBox12.Text = "" Then
        '    MsgBox("SE DEBE AGREGAR NOTA DE PEDIDO")
        'Else
        DataGridView1.Rows.Add()
        'End If


    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown

    End Sub

    Private Sub TextBox12_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox12.KeyDown
        If e.KeyCode = Keys.F1 Then
            detalle_reclamo.Label1.Text = "GUIA"
            detalle_reclamo.Label2.Text = Label27.Text
            detalle_reclamo.Label3.Text = Label26.Text
            detalle_reclamo.Label4.Text = TextBox14.Text
            detalle_reclamo.Label5.Text = TextBox2.Text
            detalle_reclamo.Show()
        End If
    End Sub

    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs)

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Dim Rsr2, Rsr22 As SqlDataReader
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
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Me.ValidateChildren() And TextBox1.Text = String.Empty And TextBox2.Text = String.Empty And TextBox12.Text = String.Empty Then
            MsgBox("FALTAN INGREAR ALGUNO CAMPOS OBLIGATORIOS")
        Else
            abrir()
            Dim va As String
            Dim Rsr2 As SqlDataReader
            Dim sql102 As String = "SELECT TOP 1  numero   AS VAL FROM Hoja_Reclamo_tela   order by numero desc"
            Dim cmd102 As New SqlCommand(sql102, conx)
            Rsr2 = cmd102.ExecuteReader()
            If Rsr2.Read() = True Then
                va = Rsr2(0) + 1
            Else
                va = 1
            End If
            Rsr2.Close()

            Select Case Trim(va).Length
                Case "1" : va = "000000000" & "" & va
                Case "2" : va = "00000000" & "" & va
                Case "3" : va = "0000000" & "" & va
                Case "4" : va = "000000" & "" & va
                Case "5" : va = "00000" & "" & va
                Case "6" : va = "0000" & "" & va
                Case "7" : va = "000" & "" & va
                Case "8" : va = "00" & "" & va
                Case "9" : va = "0" & "" & va
                Case "10" : va = va
            End Select
            If Label14.Text = "1" Then
                Dim cmd17 As New SqlCommand("update  Hoja_Reclamo_tela set fecha=@fecha ,ruc=@ruc ,cliente=@cliente ,vendedor=@vendedor ,ap_am=@ap_am,telefono=@telefono ,email=@email,lugar_visita =@lugar_visita,nota_pedido=@nota_pedido,factura=@factura,f_despacho=@f_despacho ,scma=@scma ,srp=@srp,sncd=@sncd,sncm=@sncm ,sca=@sca ,valorizado=@valorizado ,observacion_comercial=@observacion_comercial,@imagen=imagen where  numero= @numero
    ", conx)
                cmd17.Parameters.AddWithValue("@numero", Trim(TextBox1.Text))
                cmd17.Parameters.AddWithValue("@fecha", DateTimePicker1.Value.Date)
                cmd17.Parameters.AddWithValue("@ruc", Trim(TextBox2.Text))
                cmd17.Parameters.AddWithValue("@cliente", Trim(TextBox3.Text))
                cmd17.Parameters.AddWithValue("@vendedor", Trim(TextBox14.Text))
                cmd17.Parameters.AddWithValue("@ap_am", Trim(TextBox4.Text))
                cmd17.Parameters.AddWithValue("@telefono", Trim(TextBox5.Text))
                cmd17.Parameters.AddWithValue("@email", Trim(TextBox7.Text))
                cmd17.Parameters.AddWithValue("@lugar_visita", Trim(TextBox6.Text))
                cmd17.Parameters.AddWithValue("@nota_pedido", Trim(TextBox12.Text))
                cmd17.Parameters.AddWithValue("@factura", Trim(TextBox8.Text))
                cmd17.Parameters.AddWithValue("@f_despacho", DateTimePicker2.Value.Date)

                If CheckBox1.Checked = True Then
                    cmd17.Parameters.AddWithValue("@scma", "1")
                Else
                    cmd17.Parameters.AddWithValue("@scma", "0")
                End If
                If CheckBox3.Checked = True Then
                    cmd17.Parameters.AddWithValue("@srp", "1")
                Else
                    cmd17.Parameters.AddWithValue("@srp", "0")
                End If
                If CheckBox4.Checked = True Then
                    cmd17.Parameters.AddWithValue("@sncd", "1")
                Else
                    cmd17.Parameters.AddWithValue("@sncd", "0")
                End If
                If CheckBox5.Checked = True Then
                    cmd17.Parameters.AddWithValue("@sncm", "1")
                Else
                    cmd17.Parameters.AddWithValue("@sncm", "0")
                End If
                If CheckBox2.Checked = True Then
                    cmd17.Parameters.AddWithValue("@sca", "1")
                Else
                    cmd17.Parameters.AddWithValue("@sca", "0")
                End If

                cmd17.Parameters.AddWithValue("@valorizado", Convert.ToDouble(TextBox11.Text))
                cmd17.Parameters.AddWithValue("@observacion_comercial", Trim(TextBox15.Text))
                cmd17.Parameters.AddWithValue("@imagen", Trim(TextBox16.Text))
                cmd17.ExecuteNonQuery()

                Dim agregar As String = "delete from detalle_reclamo_tela where numero_reclamo = '" + Trim(TextBox1.Text) + "'"
                ELIMINAR(agregar)
                Dim p As Integer
                p = DataGridView1.Rows.Count

                For i = 0 To p - 1
                    Dim cmd18 As New SqlCommand("insert into detalle_reclamo_tela (codigo,articulo,partida,cantidad_reclamo,motivo,numero_reclamo,estado) values (@codigo,@articulo,@partida,@cantidad_reclamo,@motivo,@numero_reclamo,@estado)", conx)
                    cmd18.Parameters.AddWithValue("@codigo", Trim(DataGridView1.Rows(i).Cells(1).Value))
                    cmd18.Parameters.AddWithValue("@articulo", Trim(DataGridView1.Rows(i).Cells(2).Value))
                    cmd18.Parameters.AddWithValue("@partida", Trim(DataGridView1.Rows(i).Cells(3).Value))

                    cmd18.Parameters.AddWithValue("@cantidad_reclamo", DataGridView1.Rows(i).Cells(4).Value)
                    cmd18.Parameters.AddWithValue("@motivo", Trim(DataGridView1.Rows(i).Cells(5).Value))
                    cmd18.Parameters.AddWithValue("@numero_reclamo", Trim(TextBox1.Text))
                    cmd18.Parameters.AddWithValue("@estado", "1")
                    cmd18.ExecuteNonQuery()
                Next
                MsgBox("LA INFORMACION SE ACTUALIZO CORRECTAMENTE")
                Dim hj2 As New flog
                Dim hj11 As New vlog
                hj11.gmodulo = "HOJA_RECLAMO"
                hj11.gcuser = MDIParent1.Label3.Text
                hj11.gaccion = "EDITAR"
                hj11.gpc = My.Computer.Name
                hj11.gfecha = String.Format("{0:G}", DateTime.Now)
                hj11.gdetalle = Trim(va)
                hj11.gccia = "01"
                hj2.insertar_log(hj11)
                ENVIAR_CORREO2()
            Else
                Dim cmd15 As New SqlCommand("insert into Hoja_Reclamo_tela (numero ,fecha,ruc ,cliente ,vendedor ,ap_am,telefono ,email,lugar_visita ,nota_pedido,factura,f_despacho,scma ,srp,sncd,sncm ,sca ,valorizado ,observacion_comercial,estado,imagen,f_cierre) 
                                                                values (@numero ,@fecha,@ruc ,@cliente ,@vendedor ,@ap_am,@telefono ,@email,@lugar_visita ,@nota_pedido,@factura,@f_despacho,@scma ,@srp,@sncd,@sncm ,@sca ,@valorizado ,@observacion_comercial,@estado,@imagen,@f_cierre)", conx)
                cmd15.Parameters.AddWithValue("@numero", Trim(va))
                cmd15.Parameters.AddWithValue("@fecha", DateTimePicker1.Value.Date)
                cmd15.Parameters.AddWithValue("@ruc", Trim(TextBox2.Text))
                cmd15.Parameters.AddWithValue("@cliente", Trim(TextBox3.Text))
                cmd15.Parameters.AddWithValue("@vendedor", Trim(TextBox14.Text))
                cmd15.Parameters.AddWithValue("@ap_am", Trim(TextBox4.Text))
                cmd15.Parameters.AddWithValue("@telefono", Trim(TextBox5.Text))
                cmd15.Parameters.AddWithValue("@email", Trim(TextBox7.Text))
                cmd15.Parameters.AddWithValue("@lugar_visita", Trim(TextBox6.Text))
                cmd15.Parameters.AddWithValue("@nota_pedido", Trim(TextBox12.Text))
                cmd15.Parameters.AddWithValue("@factura", Trim(TextBox8.Text))
                cmd15.Parameters.AddWithValue("@f_despacho", DateTimePicker2.Value.Date)

                If CheckBox1.Checked = True Then
                    cmd15.Parameters.AddWithValue("@scma", "1")
                Else
                    cmd15.Parameters.AddWithValue("@scma", "0")
                End If
                If CheckBox3.Checked = True Then
                    cmd15.Parameters.AddWithValue("@srp", "1")
                Else
                    cmd15.Parameters.AddWithValue("@srp", "0")
                End If
                If CheckBox4.Checked = True Then
                    cmd15.Parameters.AddWithValue("@sncd", "1")
                Else
                    cmd15.Parameters.AddWithValue("@sncd", "0")
                End If
                If CheckBox5.Checked = True Then
                    cmd15.Parameters.AddWithValue("@sncm", "1")
                Else
                    cmd15.Parameters.AddWithValue("@sncm", "0")
                End If
                If CheckBox2.Checked = True Then
                    cmd15.Parameters.AddWithValue("@sca", "1")
                Else
                    cmd15.Parameters.AddWithValue("@sca", "0")
                End If

                cmd15.Parameters.AddWithValue("@valorizado", Convert.ToDouble(TextBox11.Text))
                cmd15.Parameters.AddWithValue("@observacion_comercial", Trim(TextBox15.Text))
                cmd15.Parameters.AddWithValue("@estado", "1")
                cmd15.Parameters.AddWithValue("@imagen", Trim(TextBox16.Text))
                cmd15.Parameters.AddWithValue("@f_cierre", DBNull.Value)
                cmd15.ExecuteNonQuery()

                Dim p As Integer
                p = DataGridView1.Rows.Count

                For i = 0 To p - 1
                    Dim cmd16 As New SqlCommand("insert into detalle_reclamo_tela (codigo,articulo,partida,cantidad_reclamo,motivo,numero_reclamo,estado) values (@codigo,@articulo,@partida,@cantidad_reclamo,@motivo,@numero_reclamo,@estado)", conx)
                    cmd16.Parameters.AddWithValue("@codigo", Trim(DataGridView1.Rows(i).Cells(1).Value))
                    cmd16.Parameters.AddWithValue("@articulo", Trim(DataGridView1.Rows(i).Cells(2).Value))
                    cmd16.Parameters.AddWithValue("@partida", Trim(DataGridView1.Rows(i).Cells(3).Value))

                    cmd16.Parameters.AddWithValue("@cantidad_reclamo", DataGridView1.Rows(i).Cells(4).Value)
                    cmd16.Parameters.AddWithValue("@motivo", Trim(DataGridView1.Rows(i).Cells(5).Value))
                    cmd16.Parameters.AddWithValue("@numero_reclamo", Trim(va))
                    cmd16.Parameters.AddWithValue("@estado", "1")
                    cmd16.ExecuteNonQuery()
                Next
                MsgBox("LA INFORMACION SE REGISTRO CORRECTAMENTE")
                Dim hj2 As New flog
                Dim hj11 As New vlog
                hj11.gmodulo = "HOJA_RECLAMO"
                hj11.gcuser = MDIParent1.Label3.Text
                hj11.gaccion = "CREAR"
                hj11.gpc = My.Computer.Name
                hj11.gfecha = String.Format("{0:G}", DateTime.Now)
                hj11.gdetalle = Trim(va)
                hj11.gccia = "01"
                hj2.insertar_log(hj11)
                ENVIAR_CORREO1()
            End If
            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("QUIERES IMPRIMIR LA HOJA DE RECLAMO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                If Label14.Text = "1" Then
                    Rpt_Reclamo.TextBox1.Text = Trim(TextBox1.Text)
                Else
                    Rpt_Reclamo.TextBox1.Text = Trim(va)
                End If

                Rpt_Reclamo.Show()

            End If
            TextBox1.Select()
            Button1.Enabled = False
            LIMPIAR()
            CORRELATIVO()
        End If
    End Sub

    Sub ENVIAR_CORREO1()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim fk As New fnopedido
        Dim jj As New vnapedido
        'Dim corre As String
        'jj.gvendedor = TextBox8.Text
        'corre = fk.buscar_correo(jj)

        Dim corre As String
        jj.gvendedor = TextBox14.Text
        corre = fk.buscar_correo(jj)
        Dim Rs As SqlDataReader
        Dim sql As String = "select correo from lista_correos where formulario ='HOJA_RECLAMO'"
        Dim cmd As New SqlCommand(sql, conx)
        Rs = cmd.ExecuteReader()
        Rs.Read()
        Dim Mailz As String = "" &
      " <!DOCTYPE html>" &
"<html xmlns='http://www.w3.org/1999/xhtml'>" &
"<head>" &
"    <title></title>" &
"</head>" &
"<body>
     <FONT COLOR='green'>INFORMACION DE LA HOJA RECLAMO</FONT><br/><br/>

<FONT COLOR='blue'>* ESTADO : </FONT> CREADO <br/>
<FONT COLOR='blue'>* Nª RECLAMO : </FONT>" + Trim(TextBox1.Text) + " <br/>
<FONT COLOR='blue'>* VENDEDOR : </FONT>" + Trim(TextBox9.Text) + " <br/>
<FONT COLOR='blue'>* EMPRESA : </FONT>" + Trim(MDIParent1.Label7.Text) + " <br/>
<FONT COLOR='blue'>* CLIENTE : </FONT>" + Trim(TextBox3.Text) + "<br/>
<FONT COLOR='blue'>* NOTA PEDIDO: </FONT>" + Trim(TextBox12.Text) + "<br/>
<FONT COLOR='blue'>* FACTURA : </FONT>" + Trim(TextBox8.Text) + "<br/>
</body>

</html>"
        With message
            .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
            .To.Add(Rs(0) & "," & corre)
            Rs.Close()
            .IsBodyHtml = True
            .Body = Mailz
            .Subject = "HOJA RECLAMO N°" & Trim(TextBox1.Text)
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
    End Sub
    Sub ENVIAR_CORREO2()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim fk As New fnopedido
        Dim jj As New vnapedido
        'Dim corre As String
        'jj.gvendedor = TextBox8.Text
        'corre = fk.buscar_correo(jj)
        Dim corre As String
        jj.gvendedor = TextBox14.Text
        corre = fk.buscar_correo(jj)
        Dim Rs As SqlDataReader
        Dim sql As String = "select correo from lista_correos where formulario ='HOJA_RECLAMO'"
        Dim cmd As New SqlCommand(sql, conx)
        Rs = cmd.ExecuteReader()
        Rs.Read()
        Dim Mailz As String = "" &
      " <!DOCTYPE html>" &
"<html xmlns='http://www.w3.org/1999/xhtml'>" &
"<head>" &
"    <title></title>" &
"</head>" &
"<body>
     <FONT COLOR='green'>INFORMACION DE LA HOJA RECLAMO</FONT><br/><br/>

<FONT COLOR='blue'>* ESTADO : </FONT> EDITADO <br/>
<FONT COLOR='blue'>* Nª RECLAMO : </FONT>" + Trim(TextBox1.Text) + " <br/>
<FONT COLOR='blue'>* VENDEDOR : </FONT>" + Trim(TextBox9.Text) + " <br/>
<FONT COLOR='blue'>* EMPRESA : </FONT>" + Trim(MDIParent1.Label7.Text) + " <br/>
<FONT COLOR='blue'>* CLIENTE : </FONT>" + Trim(TextBox3.Text) + "<br/>
<FONT COLOR='blue'>* NOTA PEDIDO: </FONT>" + Trim(TextBox12.Text) + "<br/>
<FONT COLOR='blue'>* FACTURA : </FONT>" + Trim(TextBox8.Text) + "<br/>
</body>

</html>"
        With message
            .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
            .To.Add(Rs(0) & "," & corre)
            Rs.Close()
            .IsBodyHtml = True
            .Body = Mailz
            .Subject = "HOJA RECLAMO N°" & Trim(TextBox1.Text)
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
    End Sub
    Sub LIMPIAR()
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""

        TextBox15.Text = ""
        TextBox11.Text = 0
        'TextBox13.Text = ""
        'TextBox10.Text = ""
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        CheckBox4.Checked = False
        CheckBox5.Checked = False
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        DataGridView1.Rows.Clear()
        TextBox12.Text = ""

        DateTimePicker1.Refresh()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim respuesta As DialogResult
        Dim I18, VAL As Integer
        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)

        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Label14.Text = "1"
        DateTimePicker1.Enabled = True
        DateTimePicker2.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = True
        TextBox7.Enabled = True
        TextBox8.Enabled = True
        TextBox9.Enabled = True
        TextBox12.Enabled = True
        TextBox14.Enabled = True
        Button6.Enabled = True
        DataGridView1.Enabled = True
        CheckBox1.Enabled = True
        CheckBox2.Enabled = True
        CheckBox3.Enabled = True
        CheckBox4.Enabled = True
        CheckBox5.Enabled = True
        TextBox11.Enabled = True
        TextBox15.Enabled = True
        Button2.Enabled = True
        Button1.Enabled = True
        Button7.Enabled = True
        TextBox2.Select()
    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "PARTIDA" Then

            abrir()
            Dim sql102 As String = "SELECT substring( linea_3p,1,7) +''+q1.color_3n+''+substring( linea_3p,14,19), isnull(nomb_sin,nomb_17) , unid_17,regis_3n FROM custom_vianny.dbo.qanp300 qa 
inner join custom_vianny.dbo.qan0300 q1 on qa.regis_3p = q1.regis_3n 
inner join 
custom_vianny.dbo.cag1700 cg on   cg.linea_17 = (substring( linea_3p,1,7) +''+q1.color_3n+''+substring( linea_3p,14,19))
left join custom_vianny.dbo.qag0300 q on Op_3p = ncom_3
left join custom_vianny.dbo.Cag1700_Sinonimos cs on cg.linea_17 = cs.codigo_sin and ccia_sin = '01' and item_sin = '002'
WHERE  regis_3P ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(3).Value) + "'"
            Dim cmd102 As New SqlCommand(sql102, conx)
            Rsr2 = cmd102.ExecuteReader()
            If Rsr2.Read() = True Then
                DataGridView1.Rows(e.RowIndex).Cells(1).Value = Trim(Rsr2(0))
                DataGridView1.Rows(e.RowIndex).Cells(2).Value = Trim(Rsr2(1))
            End If

            Rsr2.Close()


        End If

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "A" Then

            Detalle_pedido.Label1.Text = TextBox12.Text
            Detalle_pedido.Label2.Text = e.RowIndex
            Detalle_pedido.ShowDialog()

        End If
    End Sub
    Sub CORRELATIVO()
        abrir()
        Dim Rsr2 As SqlDataReader
        Dim sql102 As String = "SELECT TOP 1  numero   AS VAL FROM Hoja_Reclamo_tela   order by numero desc"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        If Rsr2.Read() = True Then
            TextBox1.Text = Rsr2(0) + 1
        Else
            TextBox1.Text = 1
        End If

        Select Case Trim(TextBox1.Text).Length
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
        Rsr2.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        abrir()
        Dim cmd151 As New SqlCommand("UPDATE Hoja_Reclamo_tela SET estado ='0' WHERE numero = @numero", conx)
        cmd151.Parameters.AddWithValue("@numero", Trim(TextBox1.Text))
        cmd151.ExecuteNonQuery()

        Dim cmd152 As New SqlCommand("update detalle_reclamo_tela set estado ='0' where numero_reclamo =@numero_reclamo", conx)
        cmd152.Parameters.AddWithValue("@numero_reclamo", Trim(TextBox1.Text))

        cmd152.ExecuteNonQuery()
        MsgBox("SE ANULO LO SOLICITADO")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        LIMPIAR()
        CORRELATIVO()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Try
            Dim archivo As New OpenFileDialog
            archivo.Filter = "Todos los archivos de imágenes|*.gif;*.jpg;*.jpe;*.bmp;*.png|Imágenes GIF (*.gif)|*.gif|Imágenes JPG (*.jpg, *.jpe)|*.jpg;*.jpe|Imágenes de mapas de bits (*.bmp)|*.bmp|Imágenes PNG (*.png)|*.png|Todos los archivos (*.*)|*.*"
            If archivo.ShowDialog = DialogResult.OK Then
                TextBox16.Text = archivo.FileName

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then





            Dim Rsr21, Rsr22, Rsr23, Rsr235 As SqlDataReader

            Dim sql1021 As String = "SELECT COUNT(numero),vendedor FROM Hoja_Reclamo_tela where numero='" + Trim(TextBox1.Text) + "' group by vendedor"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr21 = cmd1021.ExecuteReader()
            If Rsr21.Read() = True Then
                If Trim(Rsr21(1)) = Trim(TextBox14.Text) Then

                    If Rsr21(0) > 0 Then
                        Rsr21.Close()

                        Dim sql1023 As String = "SELECT * FROM Hoja_Reclamo_tela  hr inner join detalle_reclamo_tela dr on hr.numero = dr.numero_reclamo where numero  ='" + Trim(TextBox1.Text) + "'"
                        Dim cmd1023 As New SqlCommand(sql1023, conx)
                        Rsr23 = cmd1023.ExecuteReader()
                        Dim i51 As Integer
                        i51 = 0
                        While Rsr23.Read()
                            DateTimePicker1.Value = Rsr23(1)
                            TextBox2.Text = Trim(Rsr23(2))
                            TextBox3.Text = Trim(Rsr23(3))

                            TextBox14.Text = Trim(Rsr23(4))
                            TextBox4.Text = Trim(Rsr23(5))
                            TextBox5.Text = Trim(Rsr23(6))
                            TextBox7.Text = Trim(Rsr23(7))
                            TextBox6.Text = Trim(Rsr23(8))
                            TextBox12.Text = Trim(Rsr23(9))
                            TextBox8.Text = Trim(Rsr23(10))
                            TextBox16.Text = Trim(Rsr23(20))
                            DateTimePicker2.Value = Rsr23(11)


                            If Rsr23(12) = "1" Then
                                CheckBox1.Checked = True

                            Else
                                CheckBox1.Checked = False
                            End If
                            If Rsr23(13) = "1" Then
                                CheckBox3.Checked = True

                            Else
                                CheckBox3.Checked = False
                            End If
                            If Rsr23(14) = "1" Then
                                CheckBox4.Checked = True

                            Else
                                CheckBox4.Checked = False
                            End If
                            If Rsr23(15) = "1" Then
                                CheckBox5.Checked = True

                            Else
                                CheckBox5.Checked = False
                            End If
                            If Rsr23(16) = "1" Then
                                CheckBox2.Checked = True

                            Else
                                CheckBox2.Checked = False
                            End If
                            TextBox15.Text = Trim(Rsr23(18))
                            TextBox11.Text = Trim(Rsr23(17))
                            DataGridView1.Rows.Add()

                            DataGridView1.Rows(i51).Cells(1).Value = Trim(Rsr23(23))
                            DataGridView1.Rows(i51).Cells(2).Value = Trim(Rsr23(24))
                            DataGridView1.Rows(i51).Cells(3).Value = Trim(Rsr23(25))
                            DataGridView1.Rows(i51).Cells(4).Value = Trim(Rsr23(26))
                            DataGridView1.Rows(i51).Cells(5).Value = Trim(Rsr23(27))

                            i51 = i51 + 1
                        End While
                        Rsr23.Close()

                        Dim sql10234 As String = "SELECT * FROM calidad_reclamo_tela where numero ='" + Trim(TextBox1.Text) + "'"
                        Dim cmd10234 As New SqlCommand(sql10234, conx)
                        Rsr235 = cmd10234.ExecuteReader()
                        Dim i511 As Integer
                        i511 = 0
                        While Rsr235.Read()
                            DataGridView2.Rows.Add()
                            DataGridView2.Rows(i511).Cells(0).Value = Trim(Rsr235(2))
                            DataGridView2.Rows(i511).Cells(1).Value = Trim(Rsr235(3))
                            DataGridView2.Rows(i511).Cells(2).Value = Trim(Rsr235(4))
                            DataGridView2.Rows(i511).Cells(3).Value = Trim(Rsr235(5))

                            If Rsr235(1) = "2" Then
                                RadioButton2.Checked = True
                                RadioButton1.Checked = False
                            Else
                                If Rsr235(1) = "1" Then
                                    RadioButton2.Checked = False
                                    RadioButton1.Checked = True
                                Else
                                    RadioButton2.Checked = False
                                    RadioButton1.Checked = False
                                End If
                                i511 = i511 + 1
                            End If
                        End While
                        Rsr235.Close()
                        DateTimePicker1.Enabled = False
                        DateTimePicker2.Enabled = False
                        TextBox2.Enabled = False
                        TextBox3.Enabled = False
                        TextBox4.Enabled = False
                        TextBox5.Enabled = False
                        TextBox6.Enabled = False
                        TextBox7.Enabled = False
                        TextBox8.Enabled = False
                        TextBox9.Enabled = False
                        TextBox12.Enabled = False
                        TextBox14.Enabled = False
                        Button6.Enabled = False
                        Button7.Enabled = False
                        DataGridView1.Enabled = False
                        CheckBox1.Enabled = False
                        CheckBox2.Enabled = False
                        CheckBox3.Enabled = False
                        CheckBox4.Enabled = False
                        CheckBox5.Enabled = False
                        TextBox11.Enabled = False
                        TextBox15.Enabled = False
                        Button2.Enabled = False
                        Button1.Enabled = False
                        Button3.Enabled = True
                    Else

                    End If

                Else
                    Rsr21.Close()
                    MsgBox("EL NUMERO DE RECLAMO LE PERTENECE A OTRO VENDEDOR")
                End If
                'fin
            Else

                Rsr21.Close()
                DateTimePicker1.Enabled = True
                DateTimePicker2.Enabled = True
                TextBox2.Enabled = True
                TextBox3.Enabled = True
                TextBox4.Enabled = True
                TextBox5.Enabled = True
                TextBox6.Enabled = True
                TextBox7.Enabled = True
                TextBox8.Enabled = True
                TextBox9.Enabled = True
                TextBox12.Enabled = True
                TextBox14.Enabled = True
                Button6.Enabled = True
                DataGridView1.Enabled = True
                CheckBox1.Enabled = True
                CheckBox2.Enabled = True
                CheckBox3.Enabled = True
                CheckBox4.Enabled = True
                CheckBox5.Enabled = True
                TextBox11.Enabled = True
                TextBox15.Enabled = True
                Button2.Enabled = True
                Button7.Enabled = True
                TextBox2.Select()
            End If

        Else
            If e.KeyCode = Keys.F1 Then
                Lista_reclamo.Label3.Text = TextBox14.Text
                Lista_reclamo.Show()
            End If
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs)
        DataGridView2.Rows.Add()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs)
        Dim respuesta As DialogResult
        Dim I18, VAL As Integer
        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            DataGridView2.Rows.Remove(DataGridView2.CurrentRow)

        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Rpt_Reclamo.TextBox1.Text = TextBox1.Text
        Rpt_Reclamo.Show()
    End Sub

    Private Sub Button9_Click_1(sender As Object, e As EventArgs) Handles Button9.Click
        Dim cmd17 As New SqlCommand("update Hoja_Reclamo_tela set f_cierre = @f_cierre where numero =@numero and vendedor =@vendedor", conx)
        cmd17.Parameters.AddWithValue("@numero", Trim(TextBox1.Text))
        cmd17.Parameters.AddWithValue("@vendedor", Trim(TextBox14.Text))
        cmd17.Parameters.AddWithValue("@f_cierre", DateTimePicker3.Value.Date)
        cmd17.ExecuteNonQuery()
    End Sub

    Private Sub TextBox1_Validating(sender As Object, e As CancelEventArgs) Handles TextBox1.Validating
        If DirectCast(sender, TextBox).Text.Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "EL Nº RECLAMO ESTA VACIO")
        End If
    End Sub

    Private Sub TextBox2_Validating(sender As Object, e As CancelEventArgs) Handles TextBox2.Validating
        If DirectCast(sender, TextBox).Text.Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR EL RUC")
        End If
    End Sub

    Private Sub TextBox12_Validating(sender As Object, e As CancelEventArgs) Handles TextBox12.Validating
        If DirectCast(sender, TextBox).Text.Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR NUMERO DE PEDIDO")
        End If
    End Sub

    Private Sub TextBox16_DoubleClick(sender As Object, e As EventArgs) Handles TextBox16.DoubleClick
        If TextBox16.Text = "" Then
            MsgBox("FALTA ADJUNTAR UNA IMAGEN")
        Else

            Documento_Pdf.WebBrowser1.Navigate(TextBox16.Text.ToString, False)
            Documento_Pdf.Show()
        End If
    End Sub
End Class