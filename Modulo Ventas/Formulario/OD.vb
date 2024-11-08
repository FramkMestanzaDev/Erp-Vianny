Imports System.Data.SqlClient
Imports System.Net.Mail
Public Class OD
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim Rsr2, Rsr3, Rsr212, Rsr2126, Rsr2125 As SqlDataReader
    Dim familia As String
    Dim pm As String
    Dim estilo As String
    Dim producto As String
    Dim tela As String
    Dim lavado As String
    Dim observacion As String
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Sub limpiar()
        TextBox81.Text = ""
        TextBox88.Text = ""
        TextBox58.Text = ""
        TextBox1.Text = ""
        TextBox3.Text = ""
        TextBox72.Text = ""
        TextBox4.Text = ""
        TextBox73.Text = ""
        TextBox5.Text = ""
        TextBox74.Text = ""
        TextBox9.Text = ""
        TextBox11.Text = ""
        TextBox16.Text = ""
        TextBox10.Text = ""
        TextBox17.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox20.Text = ""
        TextBox14.Text = ""
        TextBox8.Text = ""
        TextBox18.Text = ""
        TextBox6.Text = ""
        TextBox19.Text = ""
        TextBox59.Text = ""
        TextBox75.Text = ""
        TextBox76.Text = ""
        TextBox21.Text = ""
        TextBox69.Text = ""
        TextBox70.Text = ""
        TextBox71.Text = ""
        TextBox79.Text = ""
        TextBox77.Text = ""
        TextBox96.Text = ""
        TextBox97.Text = ""
        TextBox24.Text = ""
        TextBox25.Text = ""
        TextBox35.Text = ""
        TextBox36.Text = ""
        TextBox37.Text = ""
    End Sub
    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim lc As New Listar_Clientes
        lc.TextBox2.Text = 6
        lc.TextBox3.Text = Trim(TextBox15.Text)
        lc.ShowDialog()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'FormFamilia.Label2.Text = 2
        'FormFamilia.ShowDialog()
        Dim comb As New combinaciones
        comb.Label2.Text = Label23.Text
        comb.Label3.Text = 2
        comb.Label4.Text = 2
        comb.ShowDialog()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim talla As New Form_Padron
        talla.Label2.Text = Label23.Text
        talla.Label3.Text = 2
        talla.ShowDialog()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        'Colores.TextBox2.Text = 5
        Colores.Label2.Text = Label23.Text
        Colores.ShowDialog()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        CPrenda.Label2.Text = Label23.Text
        CPrenda.Label3.Text = 2
        CPrenda.Label4.Text = 1
        CPrenda.ShowDialog()
    End Sub

    Private Sub GroupBox6_Enter(sender As Object, e As EventArgs) Handles GroupBox6.Enter

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Form_Cliente.Label1.Text = "DIVISIONOD"
        Form_Cliente.TextBox2.Text = TextBox22.Text
        Form_Cliente.ShowDialog()
    End Sub
    Function correlativo()
        abrir()
        Dim bas, num, cor As Integer
        Dim ode, version As String
        Dim sql102 As String = "select top 1 cast( SUBSTRING( ncom_3,1,2) as int) as Bas, cast( SUBSTRING( ncom_3,3,7) as int) as Num,cast( SUBSTRING( ncom_3,9,1) as int) as Cor  from custom_vianny.DBO.qag0300 where ncom_3 like '03%'  order by Num desc"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        If Rsr2.Read() = True Then
            bas = Rsr2(0)
            num = Rsr2(1)
            cor = Rsr2(2)
        End If
        Rsr2.Close()
        ode = num + 1
        version = 1
        Select Case ode.Length
            Case "1" : ode = "03000000" & ode
            Case "2" : ode = "0300000" & ode
            Case "3" : ode = "030000" & ode
            Case "4" : ode = "03000" & ode
            Case "5" : ode = "0300" & ode
            Case "6" : ode = "030" & ode
            Case "7" : ode = "03" & ode
        End Select
        Return ode & version
    End Function
    Sub BLOQUEAR_TElas()
        TextBox38.Enabled = False
        TextBox24.Enabled = False
        TextBox26.Enabled = False
        TextBox28.Enabled = False
        TextBox30.Enabled = False
        Button16.Enabled = False
        ComboBox5.Enabled = False

    End Sub

    Dim CodVen As String
    Private Sub OD_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        DateTimePicker2.Value = DateTimePicker1.Value.AddDays(7)
        ComboBox4.SelectedIndex = 0
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ComboBox6.SelectedIndex = 1

        If Label48.Text = 0 Then

            TextBox22.Text = Microsoft.VisualBasic.Mid(correlativo(), 1, 9)
            TextBox23.Text = Microsoft.VisualBasic.Mid(correlativo(), 10, 1)
        End If

        TabControl1.Enabled = False

        CodVen = ""
        abrir()
        Dim sql1021 As String = "Select cele FROM custom_vianny.DBO.TAB0100 A  Where ccia='" + Label23.Text + "' AND CTAB='MERCHA' and codigo ='" + Trim(TextBox15.Text) + "'"
        Dim cmd1021 As New SqlCommand(sql1021, conx)
        Rsr212 = cmd1021.ExecuteReader()
        If Rsr212.Read() = True Then
            CodVen = Rsr212(0)
        End If
        Rsr212.Close()

        TextBox78.Text = CodVen
        TextBox23.Select()
    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox58.Text.ToString().Trim().Length = 0 Then
            MsgBox("Primero Tiene que agregar el cliente")
        Else
            Dim fc As New Form_Cliente
            fc.Label1.Text = "MARCA"
            fc.TextBox2.Text = TextBox22.Text
            fc.ShowDialog()
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim pais As New Form_Cliente
        pais.Label1.Text = "PAIS_OD"
        pais.TextBox2.Text = TextBox22.Text
        pais.ShowDialog()
    End Sub
    Sub IngresarNotificacion()
        Dim info As String = ""
        If Label43.Text = "1" Then
            info = "Creacion"
        Else
            info = "Edicion"
        End If
        Dim cmd16 As New SqlCommand("INSERT INTO Notificaciones(FecNot,DetNot,JCoNot,SisNot,Co1Not,Co2Not,JPrNot,APrNot,Pr1Not,Pr2Not,Pr3Not,An1Not,An2Not,JudNot,ComNot,EmpNot,ModNot) 
                                                       VALUES (@FecNot,@DetNot,@JCoNot,@SisNot,@Co1Not,@Co2Not,@JPrNot,@APrNot,@Pr1Not,@Pr2Not,@Pr3Not,@An1Not,@An2Not,@JudNot,@ComNot,@EmpNot,@ModNot)", conx)
        cmd16.Parameters.AddWithValue("@FecNot", DateTimePicker1.Value)
        cmd16.Parameters.AddWithValue("@DetNot", "" + info + " de la Pm " & TextBox82.Text.ToString().Trim() & TextBox81.Text.ToString().Trim() & TextBox8.Text.ToString().Trim() & " Od :" & TextBox22.Text.ToString().Trim() & " Version :" & TextBox23.Text.ToString().Trim())
        cmd16.Parameters.AddWithValue("@JCoNot", "0")
        cmd16.Parameters.AddWithValue("@SisNot", "0")
        cmd16.Parameters.AddWithValue("@Co1Not", "0")
        cmd16.Parameters.AddWithValue("@Co2Not", "0")
        cmd16.Parameters.AddWithValue("@JPrNot", "0")
        cmd16.Parameters.AddWithValue("@APrNot", "0")
        cmd16.Parameters.AddWithValue("@Pr1Not", "0")
        cmd16.Parameters.AddWithValue("@Pr2Not", "0")
        cmd16.Parameters.AddWithValue("@Pr3Not", "0")
        cmd16.Parameters.AddWithValue("@An1Not", "0")
        cmd16.Parameters.AddWithValue("@An2Not", "0")
        cmd16.Parameters.AddWithValue("@JudNot", "0")
        cmd16.Parameters.AddWithValue("@ComNot", "0")
        cmd16.Parameters.AddWithValue("@EmpNot", "01")
        cmd16.Parameters.AddWithValue("@ModNot", "OD")
        cmd16.ExecuteNonQuery()
    End Sub
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Dim camposFaltantes As New List(Of String) ' Lista para almacenar los nombres de los campos faltantes

        Dim mensaje As DialogResult
        mensaje = MessageBox.Show("ESTA SEGURO DE GUARDAR LA INFORMACION?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (mensaje = Windows.Forms.DialogResult.Yes) Then
            If String.IsNullOrEmpty(TextBox23.Text) Then
                camposFaltantes.Add(" Correlativo OD ")
            End If
            If String.IsNullOrEmpty(TextBox81.Text) Then
                camposFaltantes.Add(" PM ")
            End If
            If String.IsNullOrEmpty(TextBox12.Text) Then
                camposFaltantes.Add(" Descripcion del Producto ")
            End If
            If String.IsNullOrEmpty(TextBox69.Text) Then
                camposFaltantes.Add(" Cantidad Solicitada ")
            End If
            If String.IsNullOrEmpty(TextBox16.Text) Then
                camposFaltantes.Add(" FAMILIA ")
            End If
            If String.IsNullOrEmpty(TextBox78.Text) Then
                camposFaltantes.Add(" VENDEDOR ")
            End If

            ' Agregar más validaciones para otros campos obligatorios según sea necesario

            ' Verificar si hay campos faltantes
            If camposFaltantes.Count > 0 Then
                ' Mostrar mensaje de error indicando los campos faltantes
                Dim camposFaltantesStr As String = String.Join(", ", camposFaltantes)
                MessageBox.Show("Falta ingresar información en los siguientes campos: " & camposFaltantesStr, "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                ' Todos los campos obligatorios están llenos, proceder con el proceso de guardado
                Guardar_Informacion()
                IngresarNotificacion()
                ENVIAR_CORREO(odf)
                TextBox22.Text = Microsoft.VisualBasic.Mid(correlativo(), 1, 9)
                TextBox23.Text = Microsoft.VisualBasic.Mid(correlativo(), 10, 1)

                MsgBox("SE GUARDO LA INFORMACION CORRECTAMENTE")
                limpiar()
                TabControl1.Enabled = False
                Label48.Text = 0
                TextBox23.Select()
                TextBox23.Enabled = True
                'TextBox23.ReadOnly = True
            End If
        End If

    End Sub

    Private Sub TextBox23_TextChanged(sender As Object, e As EventArgs) Handles TextBox23.TextChanged

    End Sub
    Sub ENVIAR_CORREO(odf)
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim fk As New fnopedido
        Dim jj As New vnapedido
        Dim colorCelda, colorCelda1, colorCelda2, colorCelda3, colorCelda4, colorCelda5, colorCelda6 As String
        Dim cabecera, DETALLE, usuario, contras As String
        usuario = ""
        contras = ""
        Dim Rsr1991 As SqlDataReader
        Dim sql1011 As String = "SELECT correo,contrasena FROM usuario_modulo where merchan_3 ='" + Trim(TextBox78.Text) + "'"
        Dim cmd1011 As New SqlCommand(sql1011, conx)
        Rsr1991 = cmd1011.ExecuteReader()
        If Rsr1991.Read() Then
            usuario = Rsr1991(0).ToString().Trim()
            contras = Rsr1991(1).ToString().Trim()
        End If
        Rsr1991.Close()
        If Label48.Text = 1 Then
            cabecera = "EDICION DE LA OD"
            DETALLE = "SE EDITO  LA OD"
        Else
            cabecera = "CREACION DE LA OD"
            DETALLE = "SE REGISTRO  LA OD"
        End If
        If (familia.ToString().Trim() <> TextBox11.Text.ToString().Trim()) Then

            colorCelda = "style='background-color:yellow;'"  ' Si coincide con el valor, cambia el fondo a amarillo
        Else
            colorCelda = "style='background-color:#e0f7fa;'"  ' De lo contrario, deja el fondo blanco
        End If
        If (pm.ToString().Trim() <> TextBox82.Text.ToString().Trim() & TextBox81.Text.ToString().Trim() & TextBox8.Text.ToString().Trim()) Then

            colorCelda1 = "style='background-color:yellow;'"  ' Si coincide con el valor, cambia el fondo a amarillo
        Else
            colorCelda1 = "style='background-color:#e0f7fa;'"  ' De lo contrario, deja el fondo blanco
        End If
        If (estilo.ToString().Trim() <> TextBox9.Text.ToString().Trim()) Then

            colorCelda2 = "style='background-color:yellow;'"  ' Si coincide con el valor, cambia el fondo a amarillo
        Else
            colorCelda2 = "style='background-color:#e0f7fa;'"  ' De lo contrario, deja el fondo blanco
        End If

        If (producto.ToString().Trim() <> TextBox12.Text.ToString().Trim()) Then

            colorCelda3 = "style='background-color:yellow;'"  ' Si coincide con el valor, cambia el fondo a amarillo
        Else
            colorCelda3 = "style='background-color:#e0f7fa;'"  ' De lo contrario, deja el fondo blanco
        End If
        If (tela.ToString().Trim() <> TextBox75.Text.ToString().Trim()) Then

            colorCelda4 = "style='background-color:yellow;'"  ' Si coincide con el valor, cambia el fondo a amarillo
        Else
            colorCelda4 = "style='background-color:#e0f7fa;'"  ' De lo contrario, deja el fondo blanco
        End If
        If (lavado.ToString().Trim() <> TextBox76.Text.ToString().Trim()) Then

            colorCelda5 = "style='background-color:yellow;'"  ' Si coincide con el valor, cambia el fondo a amarillo
        Else
            colorCelda5 = "style='background-color:#e0f7fa;'"  ' De lo contrario, deja el fondo blanco
        End If
        If (observacion.ToString().Trim() <> TextBox21.Text.ToString().Trim()) Then

            colorCelda6 = "style='background-color:yellow;'"  ' Si coincide con el valor, cambia el fondo a amarillo
        Else
            colorCelda6 = "style='background-color:#e0f7fa;'"  ' De lo contrario, deja el fondo blanco
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
                    <th align='center' ><font color='black'>OD</font></th>
                    <th align='center'><font color='black'>Versión</font></th>
                    <th align='center' " + colorCelda1 + "><font color='black'>PM</font></th>
                    <th align='center' " + colorCelda + "><font color='black'>Familia</font></th>
                    <th align='center' ><font color='black'>Cliente</font></th>
                    <th align='center' " + colorCelda2 + "><font color='black'>Estilo</font></th>
                    <th align='center' " + colorCelda3 + "><font color='black'>Producto</font></th>
                    <th align='center'><font color='black'>Fecha de Entrega</font></th>
                    <th align='center'><font color='black'>Cantidad Solicitada</font></th>
                    <th align='center' ><font color='black'>Cantidad a Producir</font></th>
                    <th align='center' ><font color='black'>Codigo Tela</font></th>
                    <th align='center' " + colorCelda4 + "><font color='black'>Tela</font></th>
                    <th align='center' " + colorCelda5 + "><font color='black'>Lavado</font></th>
                    <th align='center' " + colorCelda6 + "><font color='black'>Observación</font></th>
                    <th align='center'><font color='black'>Ficha</font></th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td align='center'> " + Mid(odf, 1, 9) + "</td>
                    <td align='center'>" + Mid(odf, 10, 1) + " </td>
                    <td align='center'>" + Trim(TextBox82.Text) & Trim(TextBox81.Text) & Trim(TextBox8.Text) + "</td>
                    <td align='center'>" + Trim(TextBox11.Text) + " </td>
                    <td align='center'>" + Trim(TextBox1.Text) + " </td>
                    <td align='center'>" + Trim(TextBox9.Text) + " </td>
                    <td align='center'>" + Trim(TextBox12.Text) + " </td>
                    <td align='center'> " + DateTimePicker2.Value.ToString() + " </td>
                    <td align='center'>" + Trim(TextBox69.Text) + " </td>
                    <td align='center'>" + Trim(TextBox71.Text) + " </td>
                    <td align='center'>" + Trim(TextBox95.Text) + " </td>
                    <td align='center'> " + Trim(TextBox75.Text) + " </td>
                    <td align='center'> " + Trim(TextBox76.Text & TextBox96.Text & TextBox97.Text) + " </td>
                    <td align='center'>" + Trim(TextBox21.Text) + " </td>
                    <td align='center'><a href=" + Trim(TextBox88.Text) + ">" + Trim(TextBox88.Text) + "</a></td>
                </tr>
            </tbody>
        </table>
    </body>
                         </html>"
        Dim Rs As SqlDataReader
        Dim sql As String = "select correo from lista_correos where formulario ='OD'"
        Dim cmd As New SqlCommand(sql, conx)
        Rs = cmd.ExecuteReader()
        Rs.Read()

        With message

            .From = New System.Net.Mail.MailAddress(usuario)
            .To.Add(Rs(0).ToString().Trim())
            '.To.Add("fmestanza@viannysac.com")
            Rs.Close()
            .IsBodyHtml = True
            .Body = Mailz
            .Subject = DETALLE & "-" & Mid(odf, 1, 9) & " - VERSION - " & Mid(odf, 10, 1)
            .Priority = System.Net.Mail.MailPriority.Normal
        End With


        With smtp

                .EnableSsl = True
                .Port = "587"
                .Host = "smtppro.zoho.com"
            '.Credentials = New Net.NetworkCredential("admin@viannysac.com", "20Via$&it2")
            .Credentials = New Net.NetworkCredential(usuario, contras)
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
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim camposFaltantes As New List(Of String) ' Lista para almacenar los nombres de los campos faltantes

        ' Validar campos obligatorios
        If String.IsNullOrEmpty(TextBox13.Text) Then
            camposFaltantes.Add(" Tallas ")
        End If

        If String.IsNullOrEmpty(TextBox16.Text) Then
            camposFaltantes.Add(" Familia ")
        End If

        'If String.IsNullOrEmpty(TextBox18.Text) Then
        '    camposFaltantes.Add(" Color Tela ")
        'End If
        If String.IsNullOrEmpty(TextBox10.Text) Then
            camposFaltantes.Add(" Estilo Interno ")
        End If
        If String.IsNullOrEmpty(TextBox15.Text) Then
            camposFaltantes.Add(" Ejecutivo  ")
        End If
        ' Agregar más validaciones para otros campos obligatorios según sea necesario

        ' Verificar si hay campos faltantes
        If camposFaltantes.Count > 0 Then
            ' Mostrar mensaje de error indicando los campos faltantes
            Dim camposFaltantesStr As String = String.Join(", ", camposFaltantes)
            MessageBox.Show("Falta ingresar información en los siguientes campos: " & camposFaltantesStr, "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            ' Todos los campos obligatorios están llenos, proceder con el proceso de guardado
            Dim tallas As New Tallas_od
            tallas.DataGridView1.Rows.Clear()
            tallas.DataGridView1.Rows.Add(11)
            'MsgBox(DataGridView2.Rows(0).Cells(12).Value)
            For i = 0 To 9
                tallas.DataGridView1.Rows(i).Cells(0).Value = DataGridView5.Rows(0).Cells(i + 13).Value

            Next
            For i2 = 0 To 9
                tallas.DataGridView1.Rows(i2).Cells(1).Value = DataGridView4.Rows(1).Cells(i2 + 1).Value
                tallas.DataGridView1.Rows(i2).Cells(2).Value = DataGridView5.Rows(1).Cells(i2 + 1).Value
            Next
            'For i1 = 0 To 9
            '    If Trim(Tallas_od.DataGridView1.Rows(i1).Cells(0).Value).Length = 0 Then
            '        Tallas_od.DataGridView1.Rows(i1).ReadOnly = True
            '    End If
            'Next
            tallas.TextBox1.Text = TextBox77.Text

            tallas.ShowDialog()

        End If
    End Sub
    Dim da As New DataTable

    Sub mostrar_rutas()
        Dim cmd As New SqlDataAdapter("SELECT * FROM  Ruta_Udp WHERE OdRut ='" + Trim(TextBox22.Text) & Trim(TextBox23.Text) + "' order by EtaRut", conx)
        cmd.Fill(da)
        DataGridView6.DataSource = da
        Dim O As Integer
        O = DataGridView6.Rows.Count

        If O > 0 Then
            GroupBox5.Enabled = False
            Button20.Enabled = False
            TextBox60.Text = DataGridView6.Rows(0).Cells(5).Value
            TextBox61.Text = DataGridView6.Rows(1).Cells(5).Value
            TextBox62.Text = DataGridView6.Rows(2).Cells(5).Value
            TextBox63.Text = DataGridView6.Rows(3).Cells(5).Value
            TextBox64.Text = DataGridView6.Rows(4).Cells(5).Value
            TextBox65.Text = DataGridView6.Rows(5).Cells(5).Value
            TextBox66.Text = DataGridView6.Rows(6).Cells(5).Value
            TextBox56.Text = DataGridView6.Rows(7).Cells(5).Value
            TextBox80.Text = DataGridView6.Rows(1).Cells(4).Value
            TextBox83.Text = DataGridView6.Rows(3).Cells(4).Value
            TextBox84.Text = DataGridView6.Rows(4).Cells(4).Value
            TextBox85.Text = DataGridView6.Rows(5).Cells(4).Value
            DateTimePicker4.Value = DataGridView6.Rows(0).Cells(3).Value
            DateTimePicker5.Value = DataGridView6.Rows(1).Cells(3).Value
            DateTimePicker6.Value = DataGridView6.Rows(2).Cells(3).Value
            DateTimePicker7.Value = DataGridView6.Rows(3).Cells(3).Value
            DateTimePicker8.Value = DataGridView6.Rows(4).Cells(3).Value
            DateTimePicker9.Value = DataGridView6.Rows(5).Cells(3).Value
            DateTimePicker10.Value = DataGridView6.Rows(6).Cells(3).Value
            DateTimePicker11.Value = DataGridView6.Rows(7).Cells(3).Value
            DateTimePicker12.Value = DataGridView6.Rows(5).Cells(6).Value

            If Trim(DataGridView6.Rows(0).Cells(7).Value) = "1" Then
                Button26.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox57.Text = 1
                CheckBox1.Enabled = False
            Else
                Button26.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox1.Enabled = True
            End If
            If Trim(DataGridView6.Rows(1).Cells(7).Value) = "1" Then
                Button27.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox67.Text = 1
                CheckBox2.Enabled = False
            Else
                Button27.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox2.Enabled = True
            End If
            If Trim(DataGridView6.Rows(2).Cells(7).Value) = "1" Then
                Button28.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox68.Text = 1
                CheckBox3.Enabled = False
            Else
                Button28.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox3.Enabled = True
            End If
            If Trim(DataGridView6.Rows(3).Cells(7).Value) = "1" Then
                Button29.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox89.Text = 1
                CheckBox4.Enabled = False
            Else
                Button29.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox4.Enabled = True
            End If
            If Trim(DataGridView6.Rows(4).Cells(7).Value) = "1" Then
                Button30.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox90.Text = 1
                CheckBox5.Enabled = False
            Else
                Button30.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox5.Enabled = True
            End If
            If Trim(DataGridView6.Rows(5).Cells(7).Value) = "1" Then
                Button31.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox91.Text = 1
                CheckBox6.Enabled = False
            Else
                Button31.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox6.Enabled = True
            End If
            If Trim(DataGridView6.Rows(6).Cells(7).Value) = "1" Then
                Button32.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox92.Text = 1
                CheckBox7.Enabled = False
            Else
                Button32.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox7.Enabled = True
            End If
            If Trim(DataGridView6.Rows(7).Cells(7).Value) = "1" Then
                Button33.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox93.Text = 1
                CheckBox8.Enabled = False
            Else
                Button33.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox8.Enabled = True
            End If

        End If



    End Sub
    Sub mostrar_telas()
        Dim Rsr1991 As SqlDataReader
        Dim sql1011 As String = "select isnull(tela2_3,''),isnull(c.nomb_17,''),ISNULL( tela3_3,''),isnull(c1.nomb_17,''), isnull(tela4_3,''),isnull(c2.nomb_17,'') as nombre4, isnull(lavadoten_3,''),isnull(otros1_3,''),isnull(otros2_3,''),isnull(estimg_3,''),isnull(otros3_3,'') from  
                                custom_vianny.DBO.qag0300  q 
                                left join custom_vianny.dbo.cag1700 c on q.ccia = c.ccia and q.tela2_3 = c.linea_17
                                left join custom_vianny.dbo.cag1700 c1 on q.ccia = c1.ccia and q.tela3_3 = c1.linea_17
                                left join custom_vianny.dbo.cag1700 c2 on q.ccia = c2.ccia and q.tela4_3 = c2.linea_17
                                where ncom_3 = '" + Trim(TextBox22.Text) & Trim(TextBox23.Text) + "' and q.ccia ='01'"
        Dim cmd1011 As New SqlCommand(sql1011, conx)
        Rsr1991 = cmd1011.ExecuteReader()
        If Rsr1991.Read() Then
            TextBox26.Text = Rsr1991(0)
            TextBox27.Text = Rsr1991(1)
            TextBox28.Text = Rsr1991(2)
            TextBox29.Text = Rsr1991(3)
            TextBox30.Text = Rsr1991(4)
            TextBox31.Text = Rsr1991(5)
            TextBox35.Text = Rsr1991(6)
            TextBox36.Text = Rsr1991(7)
            TextBox39.Text = Rsr1991(8)
            TextBox94.Text = Rsr1991(9)
            If Rsr1991(10) = "BASICO" Then
                ComboBox5.SelectedIndex = 0
            Else
                If Rsr1991(10) = "SEMI MODA" Then
                    ComboBox5.SelectedIndex = 1
                Else
                    ComboBox5.SelectedIndex = 2
                End If
            End If
            GroupBox2.Enabled = False
            GroupBox3.Enabled = False
            GroupBox4.Enabled = False
            Button19.Enabled = False
        End If
        Rsr1991.Close()
    End Sub
    Dim ll, FTY As DataTable
    Dim Rsr21, Rsr22 As SqlDataReader
    Private Sub TextBox23_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox23.KeyDown
        If e.KeyCode = Keys.Enter Then


            Dim gh As New fop
            Dim kk As New vop
            Dim lp As Integer
            TextBox1.Enabled = False


            kk.gncom_3 = Trim(TextBox22.Text) & Trim(TextBox23.Text)
            kk.gcia = Label23.Text
            ll = gh.buscar_op(kk)

            DataGridView3.DataSource = ll
            lp = DataGridView3.Rows.Count
            If lp <> 0 Then
                For i = 0 To lp - 1

                    If Trim(TextBox78.Text) = Trim(DataGridView3.Rows(0).Cells(44).Value) Or Trim(MDIParent1.Label4.Text) = "ADMINISTRADOR" Then

                        TextBox58.Text = DataGridView3.Rows(0).Cells(7).Value
                        TextBox1.Text = DataGridView3.Rows(0).Cells(8).Value
                        TextBox8.Text = Microsoft.VisualBasic.Mid(Trim(DataGridView3.Rows(0).Cells(4).Value), 7, 4)
                        TextBox81.Text = Microsoft.VisualBasic.Mid(Trim(DataGridView3.Rows(0).Cells(4).Value), 2, 5)
                        TextBox4.Text = DataGridView3.Rows(0).Cells(10).Value
                        TextBox5.Text = DataGridView3.Rows(0).Cells(11).Value
                        TextBox9.Text = Trim(DataGridView3.Rows(0).Cells(81).Value)
                        TextBox88.Text = Trim(DataGridView3.Rows(0).Cells(86).Value)
                        TextBox24.Text = DataGridView3.Rows(0).Cells(18).Value
                        TextBox25.Text = DataGridView3.Rows(0).Cells(19).Value
                        DateTimePicker1.Value = DataGridView3.Rows(0).Cells(12).Value
                        DateTimePicker2.Value = DataGridView3.Rows(0).Cells(13).Value
                        TextBox78.Text = DataGridView3.Rows(0).Cells(44).Value
                        Select Case TextBox78.Text
                            Case "0001" : TextBox15.Text = "VSILVERIO"
                            Case "0003" : TextBox15.Text = "GBALVIN"
                            Case "0025" : TextBox15.Text = "VGRAUS"
                        End Select
                        TextBox10.Text = Trim(DataGridView3.Rows(0).Cells(9).Value)
                        TextBox70.Text = DataGridView3.Rows(0).Cells(45).Value
                        TextBox71.Text = DataGridView3.Rows(0).Cells(84).Value
                        TextBox73.Text = DataGridView3.Rows(0).Cells(79).Value
                        TextBox74.Text = DataGridView3.Rows(0).Cells(80).Value
                        TextBox75.Text = Trim(DataGridView3.Rows(0).Cells(82).Value)

                        TextBox76.Text = Trim(DataGridView3.Rows(0).Cells(89).Value)
                        TextBox96.Text = Trim(DataGridView3.Rows(0).Cells(90).Value)
                        TextBox97.Text = Trim(DataGridView3.Rows(0).Cells(91).Value)

                        TextBox19.Text = DataGridView3.Rows(0).Cells(24).Value
                        TextBox12.Text = Trim(DataGridView3.Rows(0).Cells(17).Value)
                        TextBox39.Text = Trim(DataGridView3.Rows(0).Cells(17).Value)
                        TextBox69.Text = DataGridView3.Rows(0).Cells(27).Value
                        TextBox11.Text = DataGridView3.Rows(0).Cells(53).Value
                        TextBox79.Text = DataGridView3.Rows(0).Cells(43).Value
                        TextBox95.Text = DataGridView3.Rows(0).Cells(88).Value
                        TextBox16.Text = DataGridView3.Rows(0).Cells(52).Value
                        TextBox17.Text = DataGridView3.Rows(0).Cells(15).Value
                        TextBox13.Text = Trim(DataGridView3.Rows(0).Cells(51).Value)

                        TextBox3.Text = Trim(DataGridView3.Rows(0).Cells(26).Value)
                        TextBox72.Text = Trim(DataGridView3.Rows(0).Cells(25).Value)
                        TextBox21.Text = Trim(DataGridView3.Rows(0).Cells(42).Value)
                        TextBox18.Text = DataGridView3.Rows(0).Cells(85).Value
                        TextBox24.Text = DataGridView3.Rows(0).Cells(88).Value
                        TextBox25.Text = DataGridView3.Rows(0).Cells(82).Value
                        TextBox94.Text = "M" & Microsoft.VisualBasic.Mid(Trim(DataGridView3.Rows(0).Cells(4).Value), 2, 5) & Microsoft.VisualBasic.Mid(Trim(DataGridView3.Rows(0).Cells(4).Value), 7, 4)
                        RUTAMOLDE.Text = DataGridView3.Rows(0).Cells(92).Value.ToString().Trim()
                        ESTILOM.Text = DataGridView3.Rows(0).Cells(93).Value.ToString().Trim()
                        LAVADOM.Text = DataGridView3.Rows(0).Cells(94).Value.ToString().Trim()
                        TELAM.Text = DataGridView3.Rows(0).Cells(95).Value.ToString().Trim()
                        DESCRIPM.Text = DataGridView3.Rows(0).Cells(96).Value.ToString().Trim()
                        MODELISTA.Text = DataGridView3.Rows(0).Cells(97).Value.ToString().Trim()
                        ANALISTA.Text = DataGridView3.Rows(0).Cells(98).Value.ToString().Trim()
                        TextBox98.Text = DataGridView3.Rows(0).Cells(99).Value.ToString().Trim()
                        TextBox99.Text = DataGridView3.Rows(0).Cells(100).Value.ToString().Trim()
                        If Trim(DataGridView3.Rows(0).Cells(86).Value) = "D" Then
                            ComboBox6.SelectedIndex = 0
                        Else
                            ComboBox6.SelectedIndex = 1
                        End If
                        TextBox59.Text = Trim(DataGridView3.Rows(0).Cells(49).Value)
                        TextBox6.Text = Trim(DataGridView3.Rows(0).Cells(48).Value)
                        If Trim(DataGridView3.Rows(0).Cells(54).Value) = "N" Then
                            ComboBox1.SelectedIndex = 0
                        Else
                            ComboBox1.SelectedIndex = 1
                        End If
                        If Trim(DataGridView3.Rows(0).Cells(56).Value) = "01" Then
                            ComboBox2.SelectedIndex = 0
                        Else
                            ComboBox2.SelectedIndex = 1
                        End If

                        mostrar_telas()
                        mostrar_rutas()
                        Dim ab As New vtallas
                        Dim ab1 As New Padron_tallas
                        ab.gcodigo = Trim(TextBox13.Text)
                        ab.gccia = Label23.Text
                        FTY = ab1.mostrar_padrom_tallas_SELECCIONADO(ab)
                        DataGridView5.DataSource = FTY
                        DataGridView4.DataSource = FTY
                        TextBox20.Text = Trim(DataGridView5.Rows(0).Cells(13).Value) & "/" & Trim(DataGridView5.Rows(0).Cells(14).Value) & "/" & Trim(DataGridView5.Rows(0).Cells(15).Value) & "/" & Trim(DataGridView5.Rows(0).Cells(16).Value) & "/" & Trim(DataGridView5.Rows(0).Cells(17).Value) & "/" & Trim(DataGridView5.Rows(0).Cells(18).Value) & "/" & Trim(DataGridView5.Rows(0).Cells(19).Value) & "/" & Trim(DataGridView5.Rows(0).Cells(20).Value) & "/" & Trim(DataGridView5.Rows(0).Cells(21).Value) & "/" & Trim(DataGridView5.Rows(0).Cells(22).Value) & "/" & Trim(DataGridView5.Rows(0).Cells(23).Value)
                        abrir()
                        Dim sql1021 As String = "select * from custom_vianny.dbo.qag160d where ncom_16d ='" + (TextBox22.Text) & (TextBox23.Text) + "' and ccia ='" + Label23.Text + "' order by correl_16d"
                        Dim cmd1021 As New SqlCommand(sql1021, conx)
                        Rsr21 = cmd1021.ExecuteReader()
                        Dim i51 As Integer
                        i51 = 1

                        While Rsr21.Read()

                            DataGridView4.Rows(1).Cells(i51).Value = Rsr21(5)

                            i51 = i51 + 1
                        End While
                        Rsr21.Close()
                        familia = TextBox11.Text.ToString().Trim()
                        pm = TextBox82.Text.ToString().Trim() & TextBox81.Text.ToString().Trim() & TextBox8.Text.ToString().Trim()
                        estilo = TextBox9.Text.ToString().Trim()
                        producto = TextBox12.Text.ToString().Trim()
                        tela = TextBox75.Text.ToString().Trim()
                        lavado = TextBox76.Text.ToString().Trim()
                        observacion = TextBox21.Text.ToString().Trim()
                        Dim sql10217 As String = "select * from custom_vianny.dbo.qag160c where ncom_16c ='" + (TextBox22.Text) & (TextBox23.Text) + "' and ccia ='" + Label23.Text + "' order by correl_16c"
                        Dim cmd10217 As New SqlCommand(sql10217, conx)
                        Rsr22 = cmd10217.ExecuteReader()
                        Dim i517 As Integer
                        i517 = 1

                        While Rsr22.Read()

                            DataGridView5.Rows(1).Cells(i517).Value = Rsr22(6)

                            i517 = i517 + 1
                        End While
                        Rsr22.Close()


                        TabControl1.Enabled = True

                        TextBox2.Enabled = False
                        TextBox5.Enabled = False
                        TextBox6.Enabled = False
                        TextBox8.Enabled = False
                        TextBox17.Enabled = False
                        TextBox18.Enabled = False
                        TextBox21.Enabled = False
                        TextBox19.Enabled = False
                        TextBox25.Enabled = False
                        TextBox13.Enabled = False

                        TextBox16.Enabled = False
                        TextBox20.Enabled = False

                        TextBox29.Enabled = False
                        Button10.Enabled = False
                        Button3.Enabled = False
                        Button1.Enabled = False
                        Button4.Enabled = False
                        Button5.Enabled = False
                        Button2.Enabled = False
                        Button6.Enabled = False


                        bloqueado()
                        Button13.Enabled = True
                        Button15.Enabled = True
                        If Trim(MDIParent1.Label4.Text) = "ADMINISTRADOR" Then
                            Button14.Enabled = True
                            Button15.Enabled = True
                            Button19.Enabled = True
                            Button21.Enabled = True
                            Button22.Enabled = True
                            Button20.Enabled = True
                        Else
                            If Trim(MDIParent1.Label4.Text) = "COMERCIAL MANUFACTURA" Then
                                'Button14.Enabled = True
                                Button15.Enabled = True
                            Button19.Enabled = False
                            Button21.Enabled = False
                            Button22.Enabled = False
                            Button20.Enabled = False
                        Else
                            Button14.Enabled = False
                                Button15.Enabled = True
                                Button19.Enabled = True
                            Button21.Enabled = True
                            Button22.Enabled = True
                            Button20.Enabled = True
                        End If
                    End If
                    Else
                        MsgBox("LA OD A COSULTAR NO ES DE SERVICIO O NO PERTENECE AL VENDEDOR")

                        TextBox2.Enabled = False
                        TextBox5.Enabled = False
                        TextBox6.Enabled = False
                        TextBox8.Enabled = False
                        TextBox17.Enabled = False
                        TextBox18.Enabled = False
                        TextBox21.Enabled = False
                        TextBox19.Enabled = False
                        TextBox25.Enabled = False
                        TextBox13.Enabled = False

                        TextBox16.Enabled = False
                        TextBox20.Enabled = False

                        TextBox29.Enabled = False
                        Button10.Enabled = False
                        Button3.Enabled = False
                        Button1.Enabled = False
                        Button4.Enabled = False
                        Button5.Enabled = False
                        Button2.Enabled = False
                        Button6.Enabled = False
                        ComboBox1.Enabled = False
                        ComboBox2.Enabled = False
                        ComboBox3.Enabled = False
                        ComboBox4.Enabled = False


                        TextBox2.Enabled = False
                        TextBox3.Enabled = False
                        TextBox4.Enabled = False
                        TextBox5.Enabled = False
                        TextBox6.Enabled = False
                        TextBox7.Enabled = False
                    End If
                Next
            Else
                Button23.Enabled = True
                Button25.Enabled = True
                DateTimePicker1.Value = Now.Date
                DateTimePicker2.Value = DateTimePicker1.Value.AddDays(7)
                If Trim(MDIParent1.Label4.Text) = "ADMINISTRADOR" Then
                    APERTURAR()
                Else
                    If Trim(TextBox78.Text).Length = 0 Then
                        MsgBox("NO TIENE PERMISO PARA CREAR UNA ORDEN DE DESARROLLO")
                    Else
                        APERTURAR()
                    End If
                End If
            End If
        End If
    End Sub
    Sub APERTURAR()
        TabControl1.Enabled = True

        abrir()
        Dim cantod As Integer
        Dim sql1021 As String = "select COUNT( ncom_3) as Num from custom_vianny.DBO.qag0300 where ncom_3 = '" + Trim(TextBox22.Text) & Trim(TextBox23.Enabled) + "'"
        Dim cmd1021 As New SqlCommand(sql1021, conx)
        Rsr3 = cmd1021.ExecuteReader()
        If Rsr3.Read() = True Then
            cantod = Rsr3(0)
        End If
        Rsr3.Close()
        If cantod > 0 Then

            bloqueado()
            Button14.Enabled = False
            Button15.Enabled = True
        Else

            habilitado()
            Button14.Enabled = True
            Button15.Enabled = False
        End If

        TextBox58.Select()
    End Sub
    Public Sub bloqueado()
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        Button7.Enabled = False
        Button8.Enabled = False
        Button9.Enabled = False
        Button14.Enabled = False
        Button15.Enabled = False
        TextBox9.Enabled = False
        TextBox12.Enabled = False
        TextBox10.Enabled = False
        TextBox17.Enabled = False
        TextBox14.Enabled = False
        TextBox75.Enabled = False
        TextBox76.Enabled = False
        TextBox21.Enabled = False
        TextBox70.Enabled = False
        TextBox71.Enabled = False
        TextBox8.Enabled = False

    End Sub

    Private Sub TextBox70_TextChanged(sender As Object, e As EventArgs) Handles TextBox70.TextChanged

    End Sub

    Public Sub habilitado()
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        Button5.Enabled = True
        Button6.Enabled = True
        Button7.Enabled = True
        Button8.Enabled = True
        Button9.Enabled = True
        TextBox9.Enabled = True
        TextBox12.Enabled = True
        TextBox10.Enabled = True
        TextBox17.Enabled = True
        TextBox14.Enabled = True
        TextBox75.Enabled = True
        TextBox76.Enabled = True
        TextBox21.Enabled = True
        TextBox70.Enabled = True
        TextBox71.Enabled = True
        TextBox8.Enabled = True

    End Sub

    Private Sub TextBox71_TextChanged(sender As Object, e As EventArgs) Handles TextBox71.TextChanged

    End Sub

    Private Sub TextBox70_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox70.KeyPress
        ' Verificar si la tecla presionada es un dígito numérico, una tecla de control o el punto decimal
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "." Then
            ' Si no es un dígito numérico, una tecla de control ni el punto decimal, no permitir la entrada
            e.Handled = True
        End If

        ' Verificar que solo se permita un punto decimal y que no haya más de uno
        If e.KeyChar = "." AndAlso TextBox70.Text.Contains(".") Then
            e.Handled = True
        End If

    End Sub

    Private Sub TextBox22_TextChanged(sender As Object, e As EventArgs) Handles TextBox22.TextChanged

    End Sub

    Private Sub TextBox71_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox71.KeyPress
        ' Verificar si la tecla presionada es un dígito numérico o la tecla de retroceso (Backspace)
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            ' Si no es un dígito numérico ni una tecla de control, no permitir la entrada
            e.Handled = True
        End If
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Lavados.Label7.Text = "2"
        Lavados.ShowDialog()
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        TextBox22.Text = Microsoft.VisualBasic.Mid(correlativo(), 1, 9)
        TextBox23.Text = Microsoft.VisualBasic.Mid(correlativo(), 10, 1)
        DateTimePicker1.Value = DateTime.Now
        DateTimePicker2.Value = DateTimePicker1.Value.AddDays(7)
        limpiar()
        TextBox23.Select()
        TextBox22.Enabled = True
        TabControl1.Enabled = False
        DataGridView5.DataSource = ""
        DataGridView4.DataSource = ""
        correlativo()
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Dim ven As New Lista_Vendedores
        ven.ShowDialog()
    End Sub

    Private Sub TextBox22_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox22.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox23.Select()
            Button15.Enabled = True
        End If
        If e.KeyCode = Keys.F1 Then
            limpiar()
            TextBox23.Select()
            TextBox22.Enabled = True
            TabControl1.Enabled = False
            ListaOD.Label4.Text = 2
            ListaOD.ShowDialog()
        End If
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        habilitado()
        Label48.Text = 1
        Button14.Enabled = True
    End Sub
    Dim od, odf As String
    Private Sub Guardar_Informacion()

        Dim ki As New vopmanufac
        Dim ki1 As New FOPMANUF
        Dim ju As New vopmanufac
        Dim hj As New insertar_codigo
        Dim ghj, codigo, op As String
        Dim uo, ko, mn, lñ As New vresto_op
        Dim j, SUMA, NA As Integer

        'verificar si existe este od
        abrir()

        OD = TextBox22.Text & TextBox23.Text

        Dim sql1 As String = "select count(ncom_3) from custom_vianny.dbo.qag0300 where ncom_3 ='" + od + "' and ccia ='" + Label23.Text + "'"
        Dim cmd1 As New SqlCommand(sql1, conx)
        Rsr2125 = cmd1.ExecuteReader()
        If Rsr2125.Read() = True And Convert.ToInt32(Rsr2125(0)) > 0 Then
            If Label48.Text = 1 Then
                odf = od
            Else
                odf = correlativo()
                MsgBox("LA OD INGRESADA YA EXISTE, LA NUEVA OD ES : " + Microsoft.VisualBasic.Mid(correlativo(), 1, 9) + "  VERSION :" + Microsoft.VisualBasic.Mid(correlativo(), 1, 1))
            End If
        Else
            odf = od
        End If
        Rsr2125.Close()

        'fin

        ki.gncom_3 = odf
        ki.gfich_3 = Trim(TextBox58.Text)
        ki.gfcom_3 = DateTimePicker1.Value
        ki.gmone_3 = 0

        ki.gmruta1_3 = Trim(TextBox88.Text.ToString())
        ki.gmruta5_3 = Trim(TextBox98.Text.ToString())
        ki.gmruta6_3 = Trim(TextBox99.Text.ToString())
        ki.gfamilia = Trim(TextBox79.Text)
        ki.gccia = Label23.Text
        ki.gtallador = Trim(TextBox13.Text)
        ki.gFCome1_3 = DateTimePicker2.Value
        ki.gfrequerida_3 = DateTimePicker2.Value
        ki.gtcam_3 = 0.0000
        ki.gprogram_3 = Trim(TextBox82.Text) & Trim(TextBox81.Text) & Trim(TextBox8.Text)
        ki.gflag_3 = 1
        ki.gdescri_3 = TextBox12.Text
        ki.galterno_3 = TextBox76.Text
        ki.gzona = TextBox74.Text
        ki.gestilo_3 = Trim(TextBox10.Text)
        ki.gtela1_3 = Trim(TextBox95.Text)
        ki.glavadoten_3 = Trim(TextBox76.Text)
        ki.gotros1_3 = Trim(TextBox96.Text)
        ki.gotros2_3 = Trim(TextBox97.Text)


        If Trim(TextBox70.Text) = "" Then
            ki.gpfob_3 = 0
        Else
            ki.gpfob_3 = TextBox70.Text
        End If
        If Trim(TextBox71.Text) = "" Then
            ki.gcantp_3 = 0
        Else
            ki.gcantp_3 = TextBox71.Text
        End If

        ki.gcants_3 = TextBox69.Text
        ki.gmdvent_3 = ""
        Select Case ComboBox1.Text
            Case "EXTRANJERO" : ki.gtipo_mercado = "E"
            Case "NACIONAL" : ki.gtipo_mercado = "N"
        End Select
        Dim va11, va21 As String
        va11 = ""
        va21 = ""
        ki.gmerchan_3 = Trim(TextBox78.Text)
        ki.gfpago = ""
        ki.gbroker_3 = "0007"
        ki.gviatra = "01"
        ki.gdivision = TextBox73.Text
        ki.gestilo_empresa = Trim(TextBox9.Text)
        ki.gmarcacli = Trim(TextBox72.Text)
        Select Case ComboBox2.Text
            Case "PLANO" : ki.gtipo_tejido = "01"
            Case "PUNTO" : ki.gtipo_tejido = "02"
        End Select
        If Trim(ComboBox6.Text) = "MUESTRA" Then
            ki.gtipped_3 = "M"
        Else
            ki.gtipped_3 = "D"
        End If

        ki.gobserv_3 = Trim(TextBox21.Text)
        ki.gcod_color = TextBox17.Text
        ki.gcolor = TextBox6.Text
        ki.gtela = Trim(TextBox75.Text.ToString())
        'ELIMINAR OP
        Dim JJ As New vop

        Dim vg As New fop
        If Label48.Text = 1 Then
            JJ.gncom_3 = odf
            JJ.glinea_3 = Trim(TextBox17.Text)
            JJ.gcia = Label23.Text
            vg.ELIMINAR_OP2(JJ)
        End If



        ju.gfamilia = Trim(TextBox79.Text)
        ju.gSubFamilia = Trim(TextBox16.Text)
        ju.gOrigen = ki.gtipo_mercado
        ju.gcolor = Trim(TextBox6.Text)
        ju.gccia = Trim(Label23.Text)

        ghj = ki1.CORRELTIVO_PRODUCTO(ju)
        If Trim(TextBox17.Text) = "" Then
            codigo = TextBox78.Text.ToString.Trim & TextBox16.Text.ToString.Trim & ki.gtipo_mercado & "00000922" & ghj
        Else
            codigo = Trim(TextBox17.Text)
        End If


        ki.glinea_3 = codigo
        ki.gOd = ""
        ki.gOdV = ""
        ki.gmruta7_3 = RUTAMOLDE.Text.ToString().Trim()
        ki.gmruta8_3 = ESTILOM.Text.ToString().Trim()
        ki.gscobs1_3 = LAVADOM.Text.ToString().Trim()
        ki.gscobs2_3 = TELAM.Text.ToString().Trim()
        ki.gscobs3_3 = DESCRIPM.Text.ToString().Trim()
        ki.gmodelista_3 = MODELISTA.Text.ToString().Trim()
        ki.ganalista_3 = ANALISTA.Text.ToString().Trim()
        If Trim(ComboBox6.Text) = "MUESTRA" Then
            ki.gttela = "01"
        Else
            ki.gttela = "01"
        End If
        Try
            Button14.Enabled = False
            ki1.insertar_op(ki)
        Catch ex As Exception
            MessageBox.Show("Ocurrió un error al guardar los datos: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Button14.Enabled = False
        End Try



        j = Convert.ToInt32(DataGridView5.ColumnCount.ToString())

        For iP = 12 To j - 1

            If DataGridView5.Rows(0).Cells(iP).Value.ToString.Trim <> "" Then
                SUMA = SUMA + 1
            End If
        Next

        For t = 1 To SUMA
            Select Case t.ToString.Count
                Case "1" : uo.gcorel = "010" & t
                Case "2" : uo.gcorel = "01" & t

            End Select

            uo.gncom_3 = odf
            uo.glinea_3 = ki.glinea_3

                uo.gcolor = TextBox6.Text.ToString.Trim & " " & DataGridView5.Rows(0).Cells(11 + t).Value
                uo.gccia = Label23.Text
            If Not IsDBNull(DataGridView5.Rows(0).Cells(11 + t).Value) AndAlso Trim(DataGridView5.Rows(0).Cells(11 + t).Value.ToString()).Length > 0 Then
                ki1.insertar_cag170l(uo)
            End If

        Next
        For k = 1 To SUMA
            ko.gncom = odf
            Select Case k.ToString.Count
                Case "1" : ko.gcorrel = "0" & k
                Case "2" : ko.gcorrel = k

            End Select

            ko.gcantidad = DataGridView4.Rows(1).Cells(k).Value
            ko.gcodcom = Trim(TextBox59.Text)
            ko.gcodtol = DataGridView4.Rows(0).Cells(k + 1).Value
            ko.gccia = Label23.Text
            ki1.insertar_qag160d(ko)
        Next
        For kh = 1 To SUMA
            mn.gncom4 = odf
            Select Case kh.ToString.Count
                Case "1" : mn.gcorrelq = "0" & kh
                Case "2" : mn.gcorrelq = kh

            End Select

            mn.gtalla = DataGridView5.Rows(0).Cells(kh + 12).Value
            mn.gcodtal = DataGridView5.Rows(0).Cells(kh + 1).Value
            mn.gccia = Label23.Text
            If Trim(DataGridView5.Rows(1).Cells(kh).Value) = "" Then
                mn.gratio = "0"
            Else
                mn.gratio = DataGridView5.Rows(1).Cells(kh).Value
            End If

            ki1.insertar_qag160c(mn)
        Next

        lñ.gncom4 = odf
        lñ.gcorrelq = "01"
        lñ.gcolor = TextBox19.Text
        lñ.gcantidad = TextBox69.Text
        lñ.gcodcom = TextBox59.Text
        lñ.gcolorrt = TextBox6.Text
        lñ.gccia = Label23.Text
        ki1.insertar_qag160b(lñ)

        hj.glinea_17 = codigo
        hj.gnomb_17 = TextBox12.Text
        hj.gunid_17 = "UND"
        hj.gfamilia_17 = TextBox79.Text
        hj.gsubfam_17 = TextBox16.Text
        hj.gtallaje_17 = TextBox13.Text
        hj.gorigen_17 = ki.gtipo_mercado
        hj.gidcolor_17 = TextBox6.Text
        hj.garticulo_17 = TextBox12.Text
        hj.gdmarca_17 = "0114"
        hj.gcodprod_17 = TextBox9.Text
        hj.gncolor_17 = TextBox19.Text
        hj.gccia = Label23.Text
        ki1.insertar_cag1700(hj)
        MsgBox("SE REGISTRO LA ORDEN DE DESARROLLO CON EXITO")
    End Sub

    Private Sub TextBox39_TextChanged(sender As Object, e As EventArgs) Handles TextBox39.TextChanged

    End Sub

    Private Sub TabPage3_Click(sender As Object, e As EventArgs) Handles TabPage3.Click

    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) 

    End Sub

    Private Sub TextBox24_TextChanged(sender As Object, e As EventArgs) Handles TextBox24.TextChanged

    End Sub

    Private Sub TabControl2_Click(sender As Object, e As EventArgs) Handles TabControl2.Click

    End Sub

    Private Sub TextBox26_TextChanged(sender As Object, e As EventArgs) Handles TextBox26.TextChanged

    End Sub

    Private Sub TextBox22_MouseDown(sender As Object, e As MouseEventArgs) Handles TextBox22.MouseDown

    End Sub

    Private Sub TextBox28_TextChanged(sender As Object, e As EventArgs) Handles TextBox28.TextChanged

    End Sub

    Private Sub TextBox23_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox23.KeyPress

    End Sub

    Private Sub TextBox30_TextChanged(sender As Object, e As EventArgs) Handles TextBox30.TextChanged

    End Sub

    Private Sub TextBox24_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox24.KeyDown
        If e.KeyCode = Keys.F1 Then
            pproductos.CCIA.Text = Label23.Text
            pproductos.ALMACEN.Text = "08"
            pproductos.ANO.Text = MDIParent1.Label5.Text
            pproductos.FECHA.Text = Replace(DateTimePicker3.Value.ToString("yyyy-MM-dd"), "-", "")
            pproductos.TextBox1.Text = "08"
            pproductos.Label3.Text = 7

            pproductos.ShowDialog()
        End If

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            DateTimePicker4.Enabled = True
            TextBox60.Enabled = True
            TextBox80.Enabled = True
            TextBox60.Select()
        Else
            DateTimePicker4.Enabled = False
            TextBox80.Enabled = True
            TextBox60.Enabled = False
        End If

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            DateTimePicker5.Enabled = True
            TextBox80.Enabled = True
            TextBox61.Enabled = True
            TextBox61.Select()
        Else
            DateTimePicker5.Enabled = False
            TextBox80.Enabled = False
            TextBox61.Enabled = False
        End If

    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            DateTimePicker6.Enabled = True
            TextBox62.Enabled = True
            TextBox62.Select()
        Else
            DateTimePicker6.Enabled = False
            TextBox62.Enabled = False
        End If

    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            DateTimePicker7.Enabled = True
            TextBox63.Enabled = True
            TextBox83.Enabled = True
            TextBox83.Select()
        Else
            DateTimePicker7.Enabled = False
            TextBox83.Enabled = False
            TextBox63.Enabled = False
        End If

    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            DateTimePicker8.Enabled = True
            TextBox64.Enabled = True
            TextBox84.Enabled = True
            TextBox84.Select()
        Else
            DateTimePicker8.Enabled = False
            TextBox85.Enabled = True
            TextBox64.Enabled = False
        End If

    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = True Then
            DateTimePicker9.Enabled = True
            DateTimePicker12.Enabled = True
            TextBox65.Enabled = True
            TextBox85.Enabled = True
            TextBox85.Select()
        Else
            DateTimePicker9.Enabled = False
            DateTimePicker12.Enabled = False
            TextBox85.Enabled = True
            TextBox65.Enabled = False
        End If

    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked = True Then
            DateTimePicker10.Enabled = True
            TextBox66.Enabled = True
            TextBox66.Select()
        Else
            DateTimePicker10.Enabled = False
            TextBox66.Enabled = False
        End If

    End Sub



    Private Sub TextBox83_TextChanged(sender As Object, e As EventArgs) Handles TextBox83.TextChanged

    End Sub

    Private Sub TextBox26_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox26.KeyDown
        If e.KeyCode = Keys.F1 Then
            pproductos.CCIA.Text = Label23.Text
            pproductos.ALMACEN.Text = "08"
            pproductos.ANO.Text = MDIParent1.Label5.Text
            pproductos.FECHA.Text = Replace(DateTimePicker3.Value.ToString("yyyy-MM-dd"), "-", "")
            pproductos.TextBox1.Text = "08"
            pproductos.Label3.Text = 8

            pproductos.ShowDialog()
        End If
    End Sub

    Private Sub TextBox84_TextChanged(sender As Object, e As EventArgs) Handles TextBox84.TextChanged

    End Sub

    Private Sub TextBox28_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox28.KeyDown
        If e.KeyCode = Keys.F1 Then
            pproductos.CCIA.Text = Label23.Text
            pproductos.ALMACEN.Text = "08"
            pproductos.ANO.Text = MDIParent1.Label5.Text
            pproductos.FECHA.Text = Replace(DateTimePicker3.Value.ToString("yyyy-MM-dd"), "-", "")
            pproductos.TextBox1.Text = "08"
            pproductos.Label3.Text = 9

            pproductos.ShowDialog()
        End If
    End Sub

    Private Sub TextBox85_TextChanged(sender As Object, e As EventArgs) Handles TextBox85.TextChanged

    End Sub

    Private Sub TextBox30_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox30.KeyDown
        If e.KeyCode = Keys.F1 Then
            pproductos.CCIA.Text = Label23.Text
            pproductos.ALMACEN.Text = "08"
            pproductos.ANO.Text = MDIParent1.Label5.Text
            pproductos.FECHA.Text = Replace(DateTimePicker3.Value.ToString("yyyy-MM-dd"), "-", "")
            pproductos.TextBox1.Text = "08"
            pproductos.Label3.Text = 10

            pproductos.ShowDialog()
        End If
    End Sub
    Sub bloquear()
        TextBox24.Enabled = False
        TextBox26.Enabled = False
        TextBox28.Enabled = False
        TextBox30.Enabled = False
        Button16.Enabled = False
        ComboBox5.Enabled = False
        TextBox38.Enabled = False
    End Sub
    Sub desbloquear()
        TextBox24.Enabled = True
        TextBox26.Enabled = True
        TextBox28.Enabled = True
        TextBox30.Enabled = True
        Button16.Enabled = True
        ComboBox5.Enabled = True
        TextBox38.Enabled = True
    End Sub
    Sub bloquear_proceso()
        DateTimePicker4.Enabled = False
        DateTimePicker5.Enabled = False
        DateTimePicker6.Enabled = False
        DateTimePicker7.Enabled = False
        DateTimePicker8.Enabled = False
        DateTimePicker9.Enabled = False
        DateTimePicker10.Enabled = False
        DateTimePicker12.Enabled = False
        TextBox60.Enabled = False
        TextBox61.Enabled = False
        TextBox62.Enabled = False
        TextBox63.Enabled = False
        TextBox64.Enabled = False
        TextBox65.Enabled = False
        TextBox66.Enabled = False

        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
        CheckBox3.Enabled = False
        CheckBox4.Enabled = False
        CheckBox5.Enabled = False
        CheckBox6.Enabled = False
        CheckBox7.Enabled = False

    End Sub
    Sub habilitar_proceso()

        If Trim(TextBox57.Text) = "0" Then
            CheckBox1.Enabled = True
        Else
            CheckBox1.Enabled = False
        End If
        If Trim(TextBox67.Text) = "0" Then
            CheckBox2.Enabled = True
        Else
            CheckBox2.Enabled = False
        End If
        If Trim(TextBox68.Text) = "0" Then
            CheckBox3.Enabled = True
        Else
            CheckBox3.Enabled = False
        End If
        If Trim(TextBox89.Text) = "0" Then
            CheckBox4.Enabled = True
        Else
            CheckBox4.Enabled = False
        End If
        If Trim(TextBox90.Text) = "0" Then
            CheckBox5.Enabled = True
        Else
            CheckBox5.Enabled = True
        End If
        If Trim(TextBox91.Text) = "0" Then
            CheckBox6.Enabled = True
        Else
            CheckBox6.Enabled = False
        End If
        If Trim(TextBox92.Text) = "0" Then
            CheckBox7.Enabled = True
        Else
            CheckBox7.Enabled = False
        End If
        If Trim(TextBox93.Text) = "0" Then
            CheckBox8.Enabled = True
        Else
            CheckBox8.Enabled = False
        End If


    End Sub
    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        abrir()
        Dim cmd20 As New SqlCommand("update custom_vianny.dbo.qag0300 set tela1_3=@tela1_3,telaprinc_3=@telaprinc_3, tela2_3=@tela2_3, tela3_3=@tela3_3,tela4_3=@tela4_3,lavadoten_3=@lavadoten_3, otros1_3 =@otros1_3,otros2_3=@otros2_3, estimg_3=@estimg_3,otros3_3=@otros3_3  where ccia =@ccia  and ncom_3 =@ncom_3 ", conx)
        cmd20.Parameters.AddWithValue("@tela1_3", Trim(TextBox24.Text))
        cmd20.Parameters.AddWithValue("@telaprinc_3", Trim(TextBox25.Text))
        cmd20.Parameters.AddWithValue("@tela2_3", Trim(TextBox26.Text))
        cmd20.Parameters.AddWithValue("@tela3_3", Trim(TextBox28.Text))
        cmd20.Parameters.AddWithValue("@tela4_3", Trim(TextBox30.Text))
        cmd20.Parameters.AddWithValue("@lavadoten_3", Trim(TextBox35.Text))
        cmd20.Parameters.AddWithValue("@otros1_3", Trim(TextBox36.Text))
        cmd20.Parameters.AddWithValue("@otros2_3", Trim(TextBox37.Text))
        cmd20.Parameters.AddWithValue("@otros3_3", Trim(ComboBox5.Text))
        cmd20.Parameters.AddWithValue("@estimg_3", Trim(TextBox38.Text))
        cmd20.Parameters.AddWithValue("@ccia", Trim(Label23.Text))
        cmd20.Parameters.AddWithValue("@ncom_3", Trim(TextBox22.Text) & Trim(TextBox23.Text))
        cmd20.ExecuteNonQuery()
        MsgBox("SE ACTUALIZO LA INFORMACION CORRECTAMENTE")
        bloquear()

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
    Sub eliminar_Rutas()
        Dim agregar As String = "delete Ruta_Udp where OdRut ='" + Trim(TextBox22.Text) & Trim(TextBox23.Text) + "'"
        ELIMINAR(agregar)
    End Sub
    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        abrir()
        eliminar_Rutas()
        If Trim(TextBox57.Text) = "0" Then
            Dim cmd20 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut) ", conx)
            cmd20.Parameters.AddWithValue("@OdRut", Trim(TextBox22.Text) & Trim(TextBox23.Text))
            cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox40.Text))
            cmd20.Parameters.AddWithValue("@NomRut", Trim(TextBox41.Text))
            cmd20.Parameters.AddWithValue("@FechRut", DateTimePicker4.Value)
            cmd20.Parameters.AddWithValue("@AsiRut", "")
            cmd20.Parameters.AddWithValue("@ObsRut", Trim(TextBox60.Text))
            cmd20.Parameters.AddWithValue("@CiRut", "0")
            cmd20.ExecuteNonQuery()
        Else
            Dim cmd20 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut) ", conx)
            cmd20.Parameters.AddWithValue("@OdRut", Trim(TextBox22.Text) & Trim(TextBox23.Text))
            cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox40.Text))
            cmd20.Parameters.AddWithValue("@NomRut", Trim(TextBox41.Text))
            cmd20.Parameters.AddWithValue("@FechRut", DateTimePicker4.Value)
            cmd20.Parameters.AddWithValue("@AsiRut", "")
            cmd20.Parameters.AddWithValue("@ObsRut", Trim(TextBox60.Text))
            cmd20.Parameters.AddWithValue("@CiRut", "1")
            cmd20.ExecuteNonQuery()
        End If

        If Trim(TextBox67.Text) = "0" Then
            Dim cmd21 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut) ", conx)
            cmd21.Parameters.AddWithValue("@OdRut", Trim(TextBox22.Text) & Trim(TextBox23.Text))
            cmd21.Parameters.AddWithValue("@EtaRut", Trim(TextBox43.Text))
            cmd21.Parameters.AddWithValue("@NomRut", Trim(TextBox42.Text))
            cmd21.Parameters.AddWithValue("@FechRut", DateTimePicker5.Value)
            cmd21.Parameters.AddWithValue("@AsiRut", Trim(TextBox80.Text))
            cmd21.Parameters.AddWithValue("@ObsRut", Trim(TextBox61.Text))
            cmd21.Parameters.AddWithValue("@CiRut", "0")
            cmd21.ExecuteNonQuery()
        Else
            Dim cmd21 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut) ", conx)
            cmd21.Parameters.AddWithValue("@OdRut", Trim(TextBox22.Text) & Trim(TextBox23.Text))
            cmd21.Parameters.AddWithValue("@EtaRut", Trim(TextBox43.Text))
            cmd21.Parameters.AddWithValue("@NomRut", Trim(TextBox42.Text))
            cmd21.Parameters.AddWithValue("@FechRut", DateTimePicker5.Value)
            cmd21.Parameters.AddWithValue("@AsiRut", Trim(TextBox80.Text))
            cmd21.Parameters.AddWithValue("@ObsRut", Trim(TextBox61.Text))
            cmd21.Parameters.AddWithValue("@CiRut", "1")
            cmd21.ExecuteNonQuery()
        End If

        If Trim(TextBox68.Text) = "0" Then
            Dim cmd22 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut) ", conx)
            cmd22.Parameters.AddWithValue("@OdRut", Trim(TextBox22.Text) & Trim(TextBox23.Text))
            cmd22.Parameters.AddWithValue("@EtaRut", Trim(TextBox45.Text))
            cmd22.Parameters.AddWithValue("@NomRut", Trim(TextBox44.Text))
            cmd22.Parameters.AddWithValue("@FechRut", DateTimePicker6.Value)
            cmd22.Parameters.AddWithValue("@AsiRut", "")
            cmd22.Parameters.AddWithValue("@ObsRut", Trim(TextBox62.Text))
            cmd22.Parameters.AddWithValue("@CiRut", "0")
            cmd22.ExecuteNonQuery()
        Else
            Dim cmd22 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut) ", conx)
            cmd22.Parameters.AddWithValue("@OdRut", Trim(TextBox22.Text) & Trim(TextBox23.Text))
            cmd22.Parameters.AddWithValue("@EtaRut", Trim(TextBox45.Text))
            cmd22.Parameters.AddWithValue("@NomRut", Trim(TextBox44.Text))
            cmd22.Parameters.AddWithValue("@FechRut", DateTimePicker6.Value)
            cmd22.Parameters.AddWithValue("@AsiRut", "")
            cmd22.Parameters.AddWithValue("@ObsRut", Trim(TextBox62.Text))
            cmd22.Parameters.AddWithValue("@CiRut", "1")
            cmd22.ExecuteNonQuery()
        End If


        If Trim(TextBox89.Text) = "0" Then
            Dim cmd24 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut) ", conx)
            cmd24.Parameters.AddWithValue("@OdRut", Trim(TextBox22.Text) & Trim(TextBox23.Text))
            cmd24.Parameters.AddWithValue("@EtaRut", Trim(TextBox47.Text))
            cmd24.Parameters.AddWithValue("@NomRut", Trim(TextBox46.Text))
            cmd24.Parameters.AddWithValue("@FechRut", DateTimePicker7.Value)
            cmd24.Parameters.AddWithValue("@AsiRut", Trim(TextBox83.Text))
            cmd24.Parameters.AddWithValue("@ObsRut", Trim(TextBox63.Text))
            cmd24.Parameters.AddWithValue("@CiRut", "0")
            cmd24.ExecuteNonQuery()
        Else
            Dim cmd24 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut) ", conx)
            cmd24.Parameters.AddWithValue("@OdRut", Trim(TextBox22.Text) & Trim(TextBox23.Text))
            cmd24.Parameters.AddWithValue("@EtaRut", Trim(TextBox47.Text))
            cmd24.Parameters.AddWithValue("@NomRut", Trim(TextBox46.Text))
            cmd24.Parameters.AddWithValue("@FechRut", DateTimePicker7.Value)
            cmd24.Parameters.AddWithValue("@AsiRut", Trim(TextBox83.Text))
            cmd24.Parameters.AddWithValue("@ObsRut", Trim(TextBox63.Text))
            cmd24.Parameters.AddWithValue("@CiRut", "1")
            cmd24.ExecuteNonQuery()
        End If

        If Trim(TextBox90.Text) = "0" Then
            Dim cmd25 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut) ", conx)
            cmd25.Parameters.AddWithValue("@OdRut", Trim(TextBox22.Text) & Trim(TextBox23.Text))
            cmd25.Parameters.AddWithValue("@EtaRut", Trim(TextBox49.Text))
            cmd25.Parameters.AddWithValue("@NomRut", Trim(TextBox48.Text))
            cmd25.Parameters.AddWithValue("@FechRut", DateTimePicker8.Value)
            cmd25.Parameters.AddWithValue("@AsiRut", Trim(TextBox84.Text))
            cmd25.Parameters.AddWithValue("@ObsRut", Trim(TextBox64.Text))
            cmd25.Parameters.AddWithValue("@CiRut", "0")
            cmd25.ExecuteNonQuery()
        Else
            Dim cmd25 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut) ", conx)
            cmd25.Parameters.AddWithValue("@OdRut", Trim(TextBox22.Text) & Trim(TextBox23.Text))
            cmd25.Parameters.AddWithValue("@EtaRut", Trim(TextBox49.Text))
            cmd25.Parameters.AddWithValue("@NomRut", Trim(TextBox48.Text))
            cmd25.Parameters.AddWithValue("@FechRut", DateTimePicker8.Value)
            cmd25.Parameters.AddWithValue("@AsiRut", Trim(TextBox84.Text))
            cmd25.Parameters.AddWithValue("@ObsRut", Trim(TextBox64.Text))
            cmd25.Parameters.AddWithValue("@CiRut", "1")
            cmd25.ExecuteNonQuery()
        End If



        If Trim(TextBox91.Text) = "0" Then
            Dim cmd27 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut,FecIde) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut,@FecIde) ", conx)
            cmd27.Parameters.AddWithValue("@OdRut", Trim(TextBox22.Text) & Trim(TextBox23.Text))
            cmd27.Parameters.AddWithValue("@EtaRut", Trim(TextBox51.Text))
            cmd27.Parameters.AddWithValue("@NomRut", Trim(TextBox50.Text))
            cmd27.Parameters.AddWithValue("@FechRut", DateTimePicker9.Value)
            cmd27.Parameters.AddWithValue("@AsiRut", Trim(TextBox85.Text))
            cmd27.Parameters.AddWithValue("@ObsRut", Trim(TextBox65.Text))
            cmd27.Parameters.AddWithValue("@FecIde", DateTimePicker12.Value)
            cmd27.Parameters.AddWithValue("@CiRut", "0")
            cmd27.ExecuteNonQuery()
        Else
            Dim cmd27 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut,FecIde) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut,@FecIde) ", conx)
            cmd27.Parameters.AddWithValue("@OdRut", Trim(TextBox22.Text) & Trim(TextBox23.Text))
            cmd27.Parameters.AddWithValue("@EtaRut", Trim(TextBox51.Text))
            cmd27.Parameters.AddWithValue("@NomRut", Trim(TextBox50.Text))
            cmd27.Parameters.AddWithValue("@FechRut", DateTimePicker9.Value)
            cmd27.Parameters.AddWithValue("@AsiRut", Trim(TextBox85.Text))
            cmd27.Parameters.AddWithValue("@ObsRut", Trim(TextBox65.Text))
            cmd27.Parameters.AddWithValue("@FecIde", DateTimePicker12.Value)
            cmd27.Parameters.AddWithValue("@CiRut", "1")
            cmd27.ExecuteNonQuery()
        End If

        If Trim(TextBox92.Text) = "0" Then
            Dim cmd28 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut) ", conx)
            cmd28.Parameters.AddWithValue("@OdRut", Trim(TextBox22.Text) & Trim(TextBox23.Text))
            cmd28.Parameters.AddWithValue("@EtaRut", Trim(TextBox53.Text))
            cmd28.Parameters.AddWithValue("@NomRut", Trim(TextBox52.Text))
            cmd28.Parameters.AddWithValue("@FechRut", DateTimePicker10.Value)
            cmd28.Parameters.AddWithValue("@AsiRut", "")
            cmd28.Parameters.AddWithValue("@ObsRut", Trim(TextBox66.Text))
            cmd28.Parameters.AddWithValue("@CiRut", "0")
            cmd28.ExecuteNonQuery()
        Else
            Dim cmd28 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut) ", conx)
            cmd28.Parameters.AddWithValue("@OdRut", Trim(TextBox22.Text) & Trim(TextBox23.Text))
            cmd28.Parameters.AddWithValue("@EtaRut", Trim(TextBox53.Text))
            cmd28.Parameters.AddWithValue("@NomRut", Trim(TextBox52.Text))
            cmd28.Parameters.AddWithValue("@FechRut", DateTimePicker10.Value)
            cmd28.Parameters.AddWithValue("@AsiRut", "")
            cmd28.Parameters.AddWithValue("@ObsRut", Trim(TextBox66.Text))
            cmd28.Parameters.AddWithValue("@CiRut", "1")
            cmd28.ExecuteNonQuery()
        End If

        If Trim(TextBox93.Text) = "0" Then
            Dim cmd29 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut) ", conx)
            cmd29.Parameters.AddWithValue("@OdRut", Trim(TextBox22.Text) & Trim(TextBox23.Text))
            cmd29.Parameters.AddWithValue("@EtaRut", Trim(TextBox55.Text))
            cmd29.Parameters.AddWithValue("@NomRut", Trim(TextBox54.Text))
            cmd29.Parameters.AddWithValue("@FechRut", DateTimePicker11.Value)
            cmd29.Parameters.AddWithValue("@AsiRut", "")
            cmd29.Parameters.AddWithValue("@ObsRut", Trim(TextBox56.Text))
            cmd29.Parameters.AddWithValue("@CiRut", "0")
            cmd29.ExecuteNonQuery()
        Else
            Dim cmd29 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut) ", conx)
            cmd29.Parameters.AddWithValue("@OdRut", Trim(TextBox22.Text) & Trim(TextBox23.Text))
            cmd29.Parameters.AddWithValue("@EtaRut", Trim(TextBox55.Text))
            cmd29.Parameters.AddWithValue("@NomRut", Trim(TextBox54.Text))
            cmd29.Parameters.AddWithValue("@FechRut", DateTimePicker11.Value)
            cmd29.Parameters.AddWithValue("@AsiRut", "")
            cmd29.Parameters.AddWithValue("@ObsRut", Trim(TextBox56.Text))
            cmd29.Parameters.AddWithValue("@CiRut", "1")
            cmd29.ExecuteNonQuery()
        End If



        MsgBox("SE ACTUALIZO LA INFORMACION CORRECTAMENTE")
        bloquear_proceso()
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        desbloquear()
        GroupBox2.Enabled = True
        GroupBox3.Enabled = True
        GroupBox4.Enabled = True
        Button19.Enabled = True
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        habilitar_proceso()
        GroupBox5.Enabled = True
        Button20.Enabled = True
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        Dim gh As New fop
        Dim kk As New vop
        Dim lp As Integer
        TextBox1.Enabled = False


        kk.gncom_3 = TextBox22.Text & "1"
        kk.gcia = Label23.Text
        ll = gh.buscar_op(kk)

        DataGridView3.DataSource = ll
        lp = DataGridView3.Rows.Count
        If lp <> 0 Then
            For i = 0 To lp - 1

                If Trim(TextBox78.Text) = Trim(DataGridView3.Rows(0).Cells(44).Value) Or Trim(MDIParent1.Label4.Text) = "ADMINISTRADOR" Then
                    TextBox58.Text = DataGridView3.Rows(0).Cells(7).Value
                    TextBox1.Text = DataGridView3.Rows(0).Cells(8).Value
                    TextBox8.Text = Microsoft.VisualBasic.Mid(Trim(DataGridView3.Rows(0).Cells(4).Value), 7, 4)
                    TextBox81.Text = Microsoft.VisualBasic.Mid(Trim(DataGridView3.Rows(0).Cells(4).Value), 2, 5)
                    TextBox4.Text = DataGridView3.Rows(0).Cells(10).Value
                    TextBox5.Text = DataGridView3.Rows(0).Cells(11).Value
                    TextBox9.Text = DataGridView3.Rows(0).Cells(81).Value

                    TextBox78.Text = DataGridView3.Rows(0).Cells(44).Value
                    Select Case TextBox78.Text
                        Case "0001" : TextBox15.Text = "VSILVERIO"
                        Case "0003" : TextBox15.Text = "GBALVIN"
                        Case "0025" : TextBox15.Text = "VGRAUS"
                    End Select
                    TextBox10.Text = DataGridView3.Rows(0).Cells(9).Value
                    TextBox70.Text = DataGridView3.Rows(0).Cells(45).Value
                    TextBox71.Text = DataGridView3.Rows(0).Cells(84).Value
                    TextBox73.Text = DataGridView3.Rows(0).Cells(79).Value
                    TextBox74.Text = DataGridView3.Rows(0).Cells(80).Value
                    TextBox75.Text = DataGridView3.Rows(0).Cells(82).Value
                    TextBox76.Text = DataGridView3.Rows(0).Cells(83).Value
                    TextBox19.Text = DataGridView3.Rows(0).Cells(24).Value
                    TextBox12.Text = Trim(DataGridView3.Rows(0).Cells(17).Value)
                    TextBox39.Text = Trim(DataGridView3.Rows(0).Cells(17).Value)
                    TextBox69.Text = DataGridView3.Rows(0).Cells(27).Value
                    TextBox11.Text = DataGridView3.Rows(0).Cells(53).Value
                    TextBox79.Text = DataGridView3.Rows(0).Cells(43).Value
                    TextBox16.Text = DataGridView3.Rows(0).Cells(52).Value

                    TextBox13.Text = Trim(DataGridView3.Rows(0).Cells(51).Value)
                    'TextBox28.Text = Trim(DataGridView3.Rows(0).Cells(42).Value)
                    TextBox3.Text = Trim(DataGridView3.Rows(0).Cells(26).Value)
                    TextBox72.Text = Trim(DataGridView3.Rows(0).Cells(25).Value)
                    TextBox21.Text = Trim(DataGridView3.Rows(0).Cells(42).Value)
                    TextBox18.Text = DataGridView3.Rows(0).Cells(85).Value
                    TextBox59.Text = Trim(DataGridView3.Rows(0).Cells(49).Value)
                    TextBox6.Text = Trim(DataGridView3.Rows(0).Cells(48).Value)
                    If Trim(DataGridView3.Rows(0).Cells(54).Value) = "N" Then
                        ComboBox1.SelectedIndex = 0
                    Else
                        ComboBox1.SelectedIndex = 1
                    End If



                    Dim ab As New vtallas
                    Dim ab1 As New Padron_tallas
                    ab.gcodigo = Trim(TextBox13.Text)
                    ab.gccia = Label23.Text
                    FTY = ab1.mostrar_padrom_tallas_SELECCIONADO(ab)
                    DataGridView5.DataSource = FTY
                    DataGridView4.DataSource = FTY
                    TextBox20.Text = Trim(DataGridView5.Rows(0).Cells(13).Value) & "/" & Trim(DataGridView5.Rows(0).Cells(14).Value) & "/" & Trim(DataGridView5.Rows(0).Cells(15).Value) & "/" & Trim(DataGridView5.Rows(0).Cells(16).Value) & "/" & Trim(DataGridView5.Rows(0).Cells(17).Value) & "/" & Trim(DataGridView5.Rows(0).Cells(18).Value) & "/" & Trim(DataGridView5.Rows(0).Cells(19).Value) & "/" & Trim(DataGridView5.Rows(0).Cells(20).Value) & "/" & Trim(DataGridView5.Rows(0).Cells(21).Value) & "/" & Trim(DataGridView5.Rows(0).Cells(22).Value) & "/" & Trim(DataGridView5.Rows(0).Cells(23).Value)
                End If
            Next
        End If

    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        OpenFileDialog1.ShowDialog()
        TextBox88.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub TextBox75_TextChanged(sender As Object, e As EventArgs) Handles TextBox75.TextChanged

    End Sub

    Private Sub TextBox83_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox83.KeyDown

    End Sub

    Private Sub TextBox84_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox84.KeyDown
        If e.KeyCode = Keys.F1 Then
            Detalle_ruta.Label2.Text = 1
            Detalle_ruta.ShowDialog()
        End If
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        If CheckBox8.Checked = True Then
            DateTimePicker11.Enabled = True
            TextBox56.Enabled = True
            TextBox56.Select()
        Else
            DateTimePicker11.Enabled = False
            TextBox56.Enabled = False
        End If
    End Sub

    Private Sub Label17_Click(sender As Object, e As EventArgs) Handles Label17.Click

    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        'Dim detalle As New Detalle_clonar
        Detalle_clonar.Label2.Text = Label23.Text
        Detalle_clonar.ShowDialog()

    End Sub

    Private Sub TextBox80_TextChanged(sender As Object, e As EventArgs) Handles TextBox80.TextChanged

    End Sub

    Private Sub TextBox85_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox85.KeyDown
        If e.KeyCode = Keys.F1 Then
            Detalle_ruta.Label2.Text = 2
            Detalle_ruta.ShowDialog()
        End If


    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        abrir()
        Dim cantod As String
        Dim sql1021 As String = "Select CiRut from Ruta_Udp where OdRut ='" + (TextBox22.Text) & (TextBox23.Text) + "' and EtaRut ='" + Trim(TextBox40.Text) + "'"
        Dim cmd1021 As New SqlCommand(sql1021, conx)
        Rsr3 = cmd1021.ExecuteReader()
        If Rsr3.Read() = True Then
            cantod = Rsr3(0)
        End If
        Rsr3.Close()

        If cantod = 0 Then
            Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =1 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
            cmd20.Parameters.AddWithValue("@OdRut", (TextBox22.Text) & (TextBox23.Text))
            cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox40.Text))
            cmd20.ExecuteNonQuery()
            TextBox57.Text = "1"
            CheckBox1.Enabled = False
            Button26.Image = Global.Modulo_Ventas.Resources.can_abi
            MsgBox("SE CERRO EL PROCESO ANALISIS DE PRENDA")
        Else
            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("EL PREOCESO ESTA CERRADO, DESEA APERTURARLO", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =0 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                cmd20.Parameters.AddWithValue("@OdRut", (TextBox22.Text) & (TextBox23.Text))
                cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox40.Text))
                cmd20.ExecuteNonQuery()
                TextBox57.Text = "0"
                CheckBox1.Enabled = True
                Button26.Image = Global.Modulo_Ventas.Resources.can_cer
                MsgBox("SE APERTURO EL PROCESO ANALISIS DE PRENDA")
            End If
        End If

    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        If Trim(TextBox57.Text) = "0" Then
            MsgBox("NO SE PUEDE CERRAR ESTE PROCESO, CIERRE PRIMERO ANALISIS DE PRENDA")
        Else
            abrir()
            Dim cantod As String
            Dim sql1021 As String = "select CiRut from Ruta_Udp where OdRut ='" + (TextBox22.Text) & (TextBox23.Text) + "' and EtaRut ='" + Trim(TextBox43.Text) + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr3 = cmd1021.ExecuteReader()
            If Rsr3.Read() = True Then
                cantod = Rsr3(0)
            End If
            Rsr3.Close()

            If cantod = 0 Then
                Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =1 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                cmd20.Parameters.AddWithValue("@OdRut", (TextBox22.Text) & (TextBox23.Text))
                cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox43.Text))
                cmd20.ExecuteNonQuery()
                Button27.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox67.Text = "1"
                CheckBox2.Enabled = False
                MsgBox("SE CERRO EL PROCESO CREACION DE MOLDE")
            Else
                Dim respuesta As DialogResult
                respuesta = MessageBox.Show("EL PREOCESO ESTA CERRADO, DESEA APERTURARLO", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If (respuesta = Windows.Forms.DialogResult.Yes) Then
                    Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =0 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                    cmd20.Parameters.AddWithValue("@OdRut", (TextBox22.Text) & (TextBox23.Text))
                    cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox43.Text))
                    cmd20.ExecuteNonQuery()
                    TextBox67.Text = "0"
                    CheckBox2.Enabled = True
                    Button27.Image = Global.Modulo_Ventas.Resources.can_cer
                    MsgBox("SE APERTURO EL PROCESO CREACION DE MOLDE")
                End If
            End If

        End If


    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click

        If Trim(TextBox67.Text) = "0" Then
            MsgBox("NO SE PUEDE CERRAR ESTE PROCESO, CIERRE PRIMERO CREACION DE MOLDE")
        Else
            Dim cantod As String
            Dim sql1021 As String = "select CiRut from Ruta_Udp where OdRut ='" + (TextBox22.Text) & (TextBox23.Text) + "' and EtaRut ='" + Trim(TextBox45.Text) + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr3 = cmd1021.ExecuteReader()
            If Rsr3.Read() = True Then
                cantod = Rsr3(0)
            End If
            Rsr3.Close()

            If cantod = 0 Then
                Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =1 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                cmd20.Parameters.AddWithValue("@OdRut", (TextBox22.Text) & (TextBox23.Text))
                cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox45.Text))
                cmd20.ExecuteNonQuery()
                TextBox68.Text = "1"
                CheckBox3.Enabled = False
                Button28.Image = Global.Modulo_Ventas.Resources.can_abi

                MsgBox("SE CERRO EL PROCESO DE CORTE")
            Else
                Dim respuesta As DialogResult
                respuesta = MessageBox.Show("EL PREOCESO ESTA CERRADO, DESEA APERTURARLO", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If (respuesta = Windows.Forms.DialogResult.Yes) Then
                    Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =0 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                    cmd20.Parameters.AddWithValue("@OdRut", (TextBox22.Text) & (TextBox23.Text))
                    cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox45.Text))
                    cmd20.ExecuteNonQuery()
                    TextBox68.Text = "0"
                    CheckBox3.Enabled = True
                    Button28.Image = Global.Modulo_Ventas.Resources.can_cer
                    MsgBox("SE APERTURO EL PROCESO DE CORTE")
                End If
            End If
        End If


    End Sub

    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        If Trim(TextBox68.Text) = "0" Then
            MsgBox("NO SE PUEDE CERRAR ESTE PROCESO, CIERRE PRIMERO CORTE")
        Else
            Dim cantod As String
            Dim sql1021 As String = "select CiRut from Ruta_Udp where OdRut ='" + (TextBox22.Text) & (TextBox23.Text) + "' and EtaRut ='" + Trim(TextBox47.Text) + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr3 = cmd1021.ExecuteReader()
            If Rsr3.Read() = True Then
                cantod = Rsr3(0)
            End If
            Rsr3.Close()

            If cantod = 0 Then
                Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =1 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                cmd20.Parameters.AddWithValue("@OdRut", (TextBox22.Text) & (TextBox23.Text))
                cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox47.Text))
                cmd20.ExecuteNonQuery()
                Button29.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox89.Text = "1"
                CheckBox4.Enabled = False
                MsgBox("SE CERRO EL PROCESO DE APLICACIONES")
            Else
                Dim respuesta As DialogResult
                respuesta = MessageBox.Show("EL PREOCESO ESTA CERRADO, DESEA APERTURARLO", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If (respuesta = Windows.Forms.DialogResult.Yes) Then
                    Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =0 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                    cmd20.Parameters.AddWithValue("@OdRut", (TextBox22.Text) & (TextBox23.Text))
                    cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox47.Text))
                    cmd20.ExecuteNonQuery()
                    TextBox89.Text = "0"
                    CheckBox4.Enabled = True
                    Button29.Image = Global.Modulo_Ventas.Resources.can_cer
                    MsgBox("SE APERTURO EL PROCESO DE APLICACIONES")
                End If
            End If
        End If


    End Sub

    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        If Trim(TextBox89.Text) = "0" Then
            MsgBox("NO SE PUEDE CERRAR ESTE PROCESO, CIERRE PRIMERO APLICACIONES")
        Else
            Dim cantod As String
            Dim sql1021 As String = "select CiRut from Ruta_Udp where OdRut ='" + (TextBox22.Text) & (TextBox23.Text) + "' and EtaRut ='" + Trim(TextBox49.Text) + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr3 = cmd1021.ExecuteReader()
            If Rsr3.Read() = True Then
                cantod = Rsr3(0)
            End If
            Rsr3.Close()

            If cantod = 0 Then
                Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =1 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                cmd20.Parameters.AddWithValue("@OdRut", (TextBox22.Text) & (TextBox23.Text))
                cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox49.Text))
                cmd20.ExecuteNonQuery()
                TextBox90.Text = "1"
                CheckBox5.Enabled = False
                Button30.Image = Global.Modulo_Ventas.Resources.can_abi
                MsgBox("SE CERRO EL PROCESO DE CONFECCION")
            Else
                Dim respuesta As DialogResult
                respuesta = MessageBox.Show("EL PREOCESO ESTA CERRADO, DESEA APERTURARLO", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If (respuesta = Windows.Forms.DialogResult.Yes) Then
                    Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =0 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                    cmd20.Parameters.AddWithValue("@OdRut", (TextBox22.Text) & (TextBox23.Text))
                    cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox49.Text))
                    cmd20.ExecuteNonQuery()
                    TextBox90.Text = "0"
                    CheckBox5.Enabled = True
                    Button30.Image = Global.Modulo_Ventas.Resources.can_cer
                    MsgBox("SE APERTURO EL PROCESO DE CONFECCION")
                End If
            End If
        End If

    End Sub

    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click
        If Trim(TextBox90.Text) = "0" Then
            MsgBox("NO SE PUEDE CERRAR ESTE PROCESO, CIERRE PRIMERO CONFECCIONES")
        Else
            Dim cantod As String
            Dim sql1021 As String = "select CiRut from Ruta_Udp where OdRut ='" + (TextBox22.Text) & (TextBox23.Text) + "' and EtaRut ='" + Trim(TextBox51.Text) + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr3 = cmd1021.ExecuteReader()
            If Rsr3.Read() = True Then
                cantod = Rsr3(0)
            End If
            Rsr3.Close()

            If cantod = 0 Then
                Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =1 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                cmd20.Parameters.AddWithValue("@OdRut", (TextBox22.Text) & (TextBox23.Text))
                cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox51.Text))
                cmd20.ExecuteNonQuery()
                TextBox91.Text = "1"
                CheckBox6.Enabled = False
                Button31.Image = Global.Modulo_Ventas.Resources.can_abi
                MsgBox("SE CERRO EL PROCESO DE LAVADO")
            Else
                Dim respuesta As DialogResult
                respuesta = MessageBox.Show("EL PREOCESO ESTA CERRADO, DESEA APERTURARLO", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If (respuesta = Windows.Forms.DialogResult.Yes) Then
                    Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =0 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                    cmd20.Parameters.AddWithValue("@OdRut", (TextBox22.Text) & (TextBox23.Text))
                    cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox51.Text))
                    cmd20.ExecuteNonQuery()
                    CheckBox6.Enabled = True
                    TextBox91.Text = "0"
                    Button31.Image = Global.Modulo_Ventas.Resources.can_cer
                    MsgBox("SE APERTURO EL PROCESO DE LAVADO")
                End If
            End If
        End If




    End Sub

    Private Sub Button32_Click(sender As Object, e As EventArgs) Handles Button32.Click
        If Trim(TextBox91.Text) = "0" Then
            MsgBox("NO SE PUEDE CERRAR ESTE PROCESO, CIERRE PRIMERO LAVADO")
        Else
            Dim cantod As String
            Dim sql1021 As String = "select CiRut from Ruta_Udp where OdRut ='" + (TextBox22.Text) & (TextBox23.Text) + "' and EtaRut ='" + Trim(TextBox53.Text) + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr3 = cmd1021.ExecuteReader()
            If Rsr3.Read() = True Then
                cantod = Rsr3(0)
            End If
            Rsr3.Close()

            If cantod = 0 Then
                Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =1 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                cmd20.Parameters.AddWithValue("@OdRut", (TextBox22.Text) & (TextBox23.Text))
                cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox53.Text))
                cmd20.ExecuteNonQuery()
                Button32.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox92.Text = "1"
                CheckBox7.Enabled = False
                MsgBox("SE CERRO EL PROCESO DE ACABADOS")
            Else
                Dim respuesta As DialogResult
                respuesta = MessageBox.Show("EL PREOCESO ESTA CERRADO, DESEA APERTURARLO", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If (respuesta = Windows.Forms.DialogResult.Yes) Then
                    Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =0 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                    cmd20.Parameters.AddWithValue("@OdRut", (TextBox22.Text) & (TextBox23.Text))
                    cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox53.Text))
                    cmd20.ExecuteNonQuery()
                    TextBox92.Text = "0"
                    CheckBox7.Enabled = True
                    Button32.Image = Global.Modulo_Ventas.Resources.can_cer
                    MsgBox("SE APERTURO EL PROCESO DE ACABADOS")
                End If
            End If
        End If




    End Sub

    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Button33.Click
        If Trim(TextBox92.Text) = "0" Then
            MsgBox("NO SE PUEDE CERRAR ESTE PROCESO, CIERRE PRIMERO ACABADOS")
        Else
            Dim cantod As String
            Dim sql1021 As String = "select CiRut from Ruta_Udp where OdRut ='" + (TextBox22.Text) & (TextBox23.Text) + "' and EtaRut ='" + Trim(TextBox55.Text) + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr3 = cmd1021.ExecuteReader()
            If Rsr3.Read() = True Then
                cantod = Rsr3(0)
            End If
            Rsr3.Close()

            If cantod = 0 Then
                Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =1 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                cmd20.Parameters.AddWithValue("@OdRut", (TextBox22.Text) & (TextBox23.Text))
                cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox55.Text))
                cmd20.ExecuteNonQuery()
                TextBox93.Text = "1"
                CheckBox8.Enabled = False
                Button33.Image = Global.Modulo_Ventas.Resources.can_abi
                MsgBox("SE CERRO EL PROCESO ENTREGA COMERCIAL")
            Else
                Dim respuesta As DialogResult
                respuesta = MessageBox.Show("EL PREOCESO ESTA CERRADO, DESEA APERTURARLO", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If (respuesta = Windows.Forms.DialogResult.Yes) Then
                    Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =0 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                    cmd20.Parameters.AddWithValue("@OdRut", (TextBox22.Text) & (TextBox23.Text))
                    cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox55.Text))
                    cmd20.ExecuteNonQuery()
                    TextBox93.Text = "0"
                    CheckBox8.Enabled = True
                    Button33.Image = Global.Modulo_Ventas.Resources.can_cer
                    MsgBox("SE APERTURO EL PROCESO ENTREGA COMERCIAL")
                End If
            End If

        End If


    End Sub

    Private Sub Button34_Click(sender As Object, e As EventArgs) Handles Button34.Click
        Dim lava As New Lavados
        lava.Label7.Text = "1"
        lava.TextBox4.Text = Trim(TextBox76.Text)
        lava.TextBox5.Text = Trim(TextBox96.Text)
        lava.TextBox6.Text = Trim(TextBox97.Text)
        lava.ShowDialog()
    End Sub

    Private Sub Button35_Click(sender As Object, e As EventArgs) Handles Button35.Click
        abrir()
        Dim respuesta As DialogResult
        respuesta = MessageBox.Show("DESEA ANULAO LA ORDEN DE DESARROLLO ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            Dim agregar As String = "UPDATE custom_vianny.dbo.qag0300 SET flag_3 ='0' WHERE ncom_3 ='" + Trim(TextBox22.Text) & Trim(TextBox23.Text) + "' AND ccia ='01'"
            ELIMINAR(agregar)
            correlativo()
            'ENVIAR_CORREO()

            limpiar()
            TabControl1.Enabled = False

            TextBox23.Select()
            TextBox23.Enabled = True
            'TextBox23.ReadOnly = True
            MsgBox("SE ANULO LA OD CORRECTAMENTE")
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox75_HideSelectionChanged(sender As Object, e As EventArgs) Handles TextBox75.HideSelectionChanged

    End Sub

    Private Sub TextBox8_FontChanged(sender As Object, e As EventArgs) Handles TextBox8.FontChanged

    End Sub

    Private Sub TextBox8_Leave(sender As Object, e As EventArgs) Handles TextBox8.Leave
        Select Case TextBox8.Text.ToString().Trim().Length
            Case 1 : TextBox8.Text = "000" & TextBox8.Text
            Case 2 : TextBox8.Text = "00" & TextBox8.Text
            Case 3 : TextBox8.Text = "0" & TextBox8.Text
            Case 4 : TextBox8.Text = TextBox8.Text
        End Select
    End Sub

    Private Sub Button37_Click(sender As Object, e As EventArgs) Handles Button37.Click
        OpenFileDialog1.ShowDialog()
        TextBox98.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button38_Click(sender As Object, e As EventArgs) Handles Button38.Click
        OpenFileDialog1.ShowDialog()
        TextBox89.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button41_Click(sender As Object, e As EventArgs) Handles Button41.Click
        Process.Start(Trim(TextBox88.Text))
    End Sub

    Private Sub Button40_Click(sender As Object, e As EventArgs) Handles Button40.Click
        Process.Start(Trim(TextBox98.Text))
    End Sub

    Private Sub Button39_Click(sender As Object, e As EventArgs) Handles Button39.Click
        Process.Start(Trim(TextBox99.Text))
    End Sub

    Private Sub TextBox58_TextChanged(sender As Object, e As EventArgs) Handles TextBox58.TextChanged

    End Sub

    Private Sub TextBox75_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox75.KeyDown
        If e.KeyCode = Keys.F1 Then
            pproductos.CCIA.Text = Label23.Text
            pproductos.ALMACEN.Text = "08"
            pproductos.ANO.Text = MDIParent1.Label5.Text
            pproductos.FECHA.Text = Replace(DateTimePicker3.Value.ToString("yyyy-MM-dd"), "-", "")
            pproductos.TextBox1.Text = "08"
            pproductos.Label3.Text = 11

            pproductos.ShowDialog()
        End If
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        DateTimePicker2.Value = DateTimePicker1.Value.AddDays(7)
    End Sub

    Private Sub Button36_Click(sender As Object, e As EventArgs) Handles Button36.Click
        Select Case TextBox1.Text.Length

            Case "1" : TextBox1.Text = "010000000" & "" & TextBox1.Text
            Case "2" : TextBox1.Text = "01000000" & "" & TextBox1.Text
            Case "3" : TextBox1.Text = "0100000" & "" & TextBox1.Text
            Case "4" : TextBox1.Text = "010000" & "" & TextBox1.Text
            Case "5" : TextBox1.Text = "01000" & "" & TextBox1.Text
            Case "6" : TextBox1.Text = "0100" & "" & TextBox1.Text
            Case "7" : TextBox1.Text = "010" & "" & TextBox1.Text
            Case "8" : TextBox1.Text = "01" & "" & TextBox1.Text
            Case "9" : TextBox1.Text = "0" & TextBox1.Text
        End Select
        Rpt_Op.TextBox1.Text = "01"
        Rpt_Op.TextBox2.Text = Trim(TextBox22.Text) & Trim(TextBox23.Text)
        Rpt_Op.TextBox3.Text = Trim(TextBox22.Text) & Trim(TextBox23.Text)
        Rpt_Op.TextBox4.Text = 0
        Rpt_Op.TextBox5.Text = 1
        Rpt_Op.Show()
    End Sub

    Private Sub TextBox80_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox80.KeyDown
        If e.KeyCode = Keys.F1 Then
            Detalle_ruta.Label2.Text = 3
            Detalle_ruta.ShowDialog()
        End If
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.F1 Then
            Detalle_ruta.Label2.Text = 3
            Detalle_ruta.ShowDialog()
        End If
    End Sub

    Private Sub TextBox58_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox58.KeyDown
        If e.KeyCode = Keys.F1 Then
            Listar_Clientes.TextBox2.Text = 6
            Listar_Clientes.TextBox3.Text = Trim(TextBox15.Text)
            Listar_Clientes.ShowDialog()
        End If

    End Sub
End Class