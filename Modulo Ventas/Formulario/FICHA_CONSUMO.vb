Imports System.Data.SqlClient
Imports System.Net.Mail

Public Class FICHA_CONSUMO
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public conn As SqlDataAdapter
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Dim Rsr21, Rsr211, Rsr213 As SqlDataReader

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.F1 Then
            DET_FICHA_CONSUMO.ShowDialog()
        Else
            If e.KeyCode = Keys.Enter Then
                Select Case TextBox1.Text.Length
                    Case "1" : TextBox1.Text = "010000000" & "" & TextBox1.Text
                    Case "2" : TextBox1.Text = "01000000" & "" & TextBox1.Text
                    Case "3" : TextBox1.Text = "0100000" & "" & TextBox1.Text
                    Case "4" : TextBox1.Text = "010000" & "" & TextBox1.Text
                    Case "5" : TextBox1.Text = "01000" & "" & TextBox1.Text
                    Case "6" : TextBox1.Text = "0100" & "" & TextBox1.Text
                    Case "7" : TextBox1.Text = "010" & "" & TextBox1.Text
                    Case "8" : TextBox1.Text = "01" & "" & TextBox1.Text
                    Case "9" : TextBox1.Text = "0" & "" & TextBox1.Text
                    Case "10" : TextBox1.Text = TextBox1.Text
                End Select
                abrir()
                Dim sql1021 As String = "EXEC custom_vianny.dbo.GetallFichaConsumoTextilCab '01','" + Trim(TextBox1.Text) + "',NULL"
                Dim cmd1021 As New SqlCommand(sql1021, conx)
                Rsr21 = cmd1021.ExecuteReader()
                Dim i5 As Integer
                i5 = 0
                If Rsr21.Read() = True Then
                    TextBox2.Text = Trim(Rsr21(4))
                    TextBox3.Text = Trim(Rsr21(5))
                    TextBox4.Text = Trim(Rsr21(6))
                    TextBox5.Text = Trim(Rsr21(7))
                    TextBox6.Text = Trim(Rsr21(8))
                    TextBox7.Text = Trim(Rsr21(9))
                    TextBox10.Text = Trim(Rsr21(12))
                    TextBox11.Text = Trim(Rsr21(16))
                    TextBox12.Text = Trim(Rsr21(15))
                    TextBox1.Enabled = False
                    Rsr21.Close()
                    Dim sql10213 As String = "EXEC custom_vianny.dbo.GetallFichaConsumoTextilDet '01','" + Trim(TextBox1.Text) + "'"
                    Dim cmd10213 As New SqlCommand(sql10213, conx)
                    Rsr213 = cmd10213.ExecuteReader()
                    While Rsr213.Read()
                        DataGridView1.Rows.Add()
                        DataGridView1.Rows(i5).Cells(0).Value = i5 + 1
                        DataGridView1.Rows(i5).Cells(1).Value = Rsr213(5)
                        DataGridView1.Rows(i5).Cells(2).Value = Rsr213(3)
                        DataGridView1.Rows(i5).Cells(3).Value = Rsr213(2)
                        DataGridView1.Rows(i5).Cells(4).Value = Rsr213(7)
                        DataGridView1.Rows(i5).Cells(5).Value = Rsr213(19)
                        DataGridView1.Rows(i5).Cells(6).Value = Rsr213(8)
                        DataGridView1.Rows(i5).Cells(7).Value = Rsr213(16)
                        DataGridView1.Rows(i5).Cells(8).Value = Rsr213(9)
                        DataGridView1.Rows(i5).Cells(9).Value = Rsr213(10)
                        DataGridView1.Rows(i5).Cells(10).Value = Rsr213(13)
                        DataGridView1.Rows(i5).Cells(11).Value = Rsr213(14)
                        DataGridView1.Rows(i5).Cells(12).Value = Rsr213(15)
                        DataGridView1.Rows(i5).Cells(13).Value = Rsr213(4)
                        i5 = i5 + 1

                    End While
                    Rsr213.Close()
                    DataGridView1.ReadOnly = True
                Else
                    MsgBox("LA OP NO EXISTE")
                End If
                Rsr21.Close()

            End If
        End If


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DataGridView1.Rows.Clear()
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox1.Text = ""
        TextBox1.Enabled = True
        TextBox1.Select()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'If Convert.ToDouble(TextBox10.Text) <= 0 Then
        '    MessageBox.Show("La Cantidad a Producir iene que ser Mayor a 0", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        'Else
        TELAS_OP.Label2.Text = TextBox1.Text
        TELAS_OP.ShowDialog()
        'End If

    End Sub

    Private Sub FICHA_CONSUMO_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Select()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        abrir()
        'If e.ColumnIndex = 4 Then
        Dim sql10211 As String = "select nomb_17 from custom_vianny.dbo.cag1700 where linea_17 ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value) + "' and ccia ='01'"
        Dim cmd10211 As New SqlCommand(sql10211, conx)
        Rsr211 = cmd10211.ExecuteReader()
        If Rsr211.Read = True Then
            TextBox8.Text = Rsr211(0)
        Else
            TextBox8.Text = ""
        End If
        Rsr211.Close()
        'End If

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        DataGridView1.ReadOnly = False
        Button3.Enabled = True
        Button1.Enabled = True
        Button2.Enabled = True
        Button6.Enabled = True
        Label12.Text = "1"
        If MDIParent2.Label6.Text.ToString().Trim() = "WROSALES" Then
            DataGridView1.Columns(11).ReadOnly = False
            DataGridView1.Columns(12).ReadOnly = False
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox9.Enabled = True
            TextBox9.Select()
        Else
            TextBox9.Enabled = False
        End If
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged

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
    Private Function Validar_Textobox() As Boolean
        Dim metros As Integer
        Dim kilos As Integer

        For i = 0 To DataGridView1.Rows.Count - 1
            If Convert.ToDouble(DataGridView1.Rows(i).Cells(11).Value) = 0 Then
                metros = metros + 1
            End If
            If Convert.ToDouble(DataGridView1.Rows(i).Cells(12).Value) = 0 Then
                kilos = kilos + 1
            End If
        Next
        If DataGridView1.Rows.Count = 0 Then
            MessageBox.Show("Falta Ingresar informacion en el Detalle del Consumo", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If
        'If Convert.ToDouble(TextBox10.Text) <= 0 Then
        '    MessageBox.Show("Falta Ingresar informacion en la Cantidad a Producir", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        '    Return False
        'End If
        If metros > 0 Then
            MessageBox.Show("El Campo Consumo Lineal tiene que ser Mayor a 0", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If
        If kilos > 0 Then
            MessageBox.Show("El Campo Consumo Kgs tiene que ser Mayor a 0", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If
        If Convert.ToInt16(TextBox1.Text.Length) < 10 Then
            MessageBox.Show("Varificar la Op, porque no tiene la estructura correcta", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If

        Return True
    End Function


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Dim ca As Integer
            ca = DataGridView1.Rows.Count
            If ca = 0 Then
                MsgBox("NO SE PUEDE GUARDAR SI NO HAY INFORMACION DE LAS TELAS A UTILIZAR")
            Else
                If Validar_Textobox() Then


                    Dim agregar As String = "delete from custom_vianny.dbo.Consumo_Op_cab where ccia ='" + Label8.Text + "' and op ='" + Trim(TextBox1.Text) + " '"
                    Dim agregar1 As String = "delete from custom_vianny.dbo.Consumo_Op_Det where ccia ='" + Label8.Text + "' and op ='" + Trim(TextBox1.Text) + "'"
                    ELIMINAR(agregar)
                    ELIMINAR(agregar1)
                    Dim cmd200 As New SqlCommand("EXEC FichaConsumoTextilCabInsertUpdated @ccia, @op , @fecha ,1", conx)
                    cmd200.Parameters.AddWithValue("@ccia", Label8.Text)
                    cmd200.Parameters.AddWithValue("@op", Trim(TextBox1.Text))
                    cmd200.Parameters.AddWithValue("@fecha", DateTimePicker1.Value)

                    cmd200.ExecuteNonQuery()
                    Dim p As Integer
                    p = DataGridView1.Rows.Count

                    For i = 0 To p - 1
                        Dim cmd2001 As New SqlCommand("EXEC FichaConsumoTextilDetInsertUpdatedvs2 @ccia , @op,0.00 , @anchotubular , '' , @detalle,@tela, @anchoabierto , @densidad ,@largotizaso,@prenda , 0.000 , 0.000 , @merma,@consumolineal ,@consumokg , @eficiencia , @items,@anchoutilizado,''", conx)
                        cmd2001.Parameters.AddWithValue("@ccia", Label8.Text)
                        cmd2001.Parameters.AddWithValue("@op", Trim(TextBox1.Text))
                        cmd2001.Parameters.AddWithValue("@anchotubular", (DataGridView1.Rows(i).Cells(3).Value))
                        cmd2001.Parameters.AddWithValue("@detalle", Trim(DataGridView1.Rows(i).Cells(13).Value))
                        cmd2001.Parameters.AddWithValue("@tela", Trim(DataGridView1.Rows(i).Cells(1).Value))
                        cmd2001.Parameters.AddWithValue("@anchoabierto", (DataGridView1.Rows(i).Cells(4).Value))
                        cmd2001.Parameters.AddWithValue("@densidad", (DataGridView1.Rows(i).Cells(6).Value))
                        cmd2001.Parameters.AddWithValue("@largotizaso", (DataGridView1.Rows(i).Cells(8).Value))
                        cmd2001.Parameters.AddWithValue("@prenda", Trim(DataGridView1.Rows(i).Cells(9).Value))
                        cmd2001.Parameters.AddWithValue("@merma", (DataGridView1.Rows(i).Cells(10).Value))
                        cmd2001.Parameters.AddWithValue("@consumolineal", (DataGridView1.Rows(i).Cells(11).Value))
                        cmd2001.Parameters.AddWithValue("@consumokg", (DataGridView1.Rows(i).Cells(12).Value))
                        cmd2001.Parameters.AddWithValue("@eficiencia", (DataGridView1.Rows(i).Cells(7).Value))
                        cmd2001.Parameters.AddWithValue("@items", Trim(DataGridView1.Rows(i).Cells(13).Value))
                        cmd2001.Parameters.AddWithValue("@anchoutilizado", (DataGridView1.Rows(i).Cells(5).Value))
                        cmd2001.ExecuteNonQuery()
                    Next
                    actualizar_estado()
                    If Label12.Text = "1" Then
                        InsertarLog(2)
                    Else
                        InsertarLog(1)
                    End If
                    MsgBox("SE GUARDO LA INFORMACION CORRECTAMENTE")
                    ENVIAR_CORREO()


                    DataGridView1.Rows.Clear()
                        TextBox2.Text = ""
                        TextBox3.Text = ""
                        TextBox4.Text = ""
                        TextBox5.Text = ""
                        TextBox6.Text = ""
                        TextBox7.Text = ""
                        TextBox1.Text = ""
                        TextBox1.Enabled = True
                        Label12.Text = "0"
                        TextBox1.Select()
                        Me.Close()
                    End If
                End If
        Catch ex As Exception
            MsgBox((ex.ToString()))
        End Try


    End Sub
    Sub InsertarLog(accion)
        Dim hj2 As New flog
        Dim hj1 As New vlog
        hj1.gmodulo = "CONSUMO-TELA"
        hj1.gcuser = MDIParent2.Label6.Text
        hj1.gaccion = accion
        hj1.gpc = My.Computer.Name
        hj1.gfecha = DateTimePicker1.Value
        hj1.gdetalle = Trim(TextBox1.Text)
        hj1.gccia = Label8.Text
        hj2.insertar_log(hj1)
    End Sub
    Sub ENVIAR_CORREO()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim fk As New fnopedido
        Dim jj As New vnapedido
        Dim cabecera, DETALLE, usuario, contras As String
        Dim colorCelda, colorCelda1, colorCelda2, colorCelda3, colorCelda4, colorCelda5, colorCelda6, colorCelda7, colorCelda8, colorCelda9, colorCelda10 As String
        Dim suma1 = 0, suma2 = 0, suma3 = 0, suma4 = 0, suma5 = 0, suma6 = 0, suma7 = 0, suma8 = 0, suma9 As Integer
        usuario = ""
        suma9 = 0
        contras = ""
        Dim Rsr1991 As SqlDataReader
        Dim sql1011 As String = "SELECT correo,contrasena FROM usuario_modulo where usuario ='" + Trim(Label13.Text) + "'"
        Dim cmd1011 As New SqlCommand(sql1011, conx)
        Rsr1991 = cmd1011.ExecuteReader()
        If Rsr1991.Read() Then
            usuario = Rsr1991(0).ToString().Trim()
            contras = Rsr1991(1).ToString().Trim()
        End If
        Rsr1991.Close()

        If Label12.Text = "1" Then
            cabecera = "EDICION DEL CONSUMO DE LA TELA"
            DETALLE = "SE EDITO EL CONSUMO DE LA TELA: OP -"
        Else
            cabecera = "CREACION DEL CONSUMO DE LA TELA"
            DETALLE = "SE CREO EL CONSUMO DE LA TELA:  OP -"
        End If


        If celdasModificadas.Count > 0 Then
            For Each celda In celdasModificadas

                If (celda.ToString().Trim() = "TELA") Then
                    suma1 = suma1 + 1
                End If
                If (celda.ToString().Trim() = "ANCH") Then
                    suma2 = suma2 + 1
                End If
                If (celda.ToString().Trim() = "DENSIDAD") Then
                    suma3 = suma3 + 1
                End If
                If (celda.ToString().Trim() = "EFICIENCIA") Then
                    suma4 = suma4 + 1
                End If
                If (celda.ToString().Trim() = "LARGO") Then
                    suma5 = suma5 + 1
                End If
                If (celda.ToString().Trim() = "PRENDAS") Then
                    suma6 = suma6 + 1
                End If
                If (celda.ToString().Trim() = "MERMA") Then
                    suma7 = suma7 + 1
                End If
                If (celda.ToString().Trim() = "C") Then
                    suma8 = suma8 + 1
                End If
                If (celda.ToString().Trim() = "CONSUMO") Then
                    suma9 = suma9 + 1
                End If
            Next
        End If


        If (suma1 >0) Then                 
              colorCelda2 = "style='background-color:yellow;'"  ' Si coincide con el valor, cambia el fondo a amarillo
                     Else
              colorCelda2 = "style='background-color:#e0f7fa;'"  ' De lo contrario, deja el fondo blanco           
                End If
        If (suma2 > 0) Then
            colorCelda3 = "style='background-color:yellow;'"  ' Si coincide con el valor, cambia el fondo a amarillo
        Else
            colorCelda3 = "style='background-color:#e0f7fa;'"  ' De lo contrario, deja el fondo blanco       
                End If
        If (suma3 > 0) Then
            colorCelda4 = "style='background-color:yellow;'"  ' Si coincide con el valor, cambia el fondo a amarillo
        Else
            colorCelda4 = "style='background-color:#e0f7fa;'"  ' De lo contrario, deja el fondo blanco       
                End If
              If (suma4 >0) Then
              colorCelda5 = "style='background-color:yellow;'"  ' Si coincide con el valor, cambia el fondo a amarillo
                     Else
              colorCelda5 = "style='background-color:#e0f7fa;'"  ' De lo contrario, deja el fondo blanco       
                End If
                If (suma5 >0) Then
               colorCelda6 = "style='background-color:yellow;'"  ' Si coincide con el valor, cambia el fondo a amarillo
                     Else
              colorCelda6 = "style='background-color:#e0f7fa;'"  ' De lo contrario, deja el fondo blanco       
                End If
               If (suma6 >0) Then
               colorCelda7 = "style='background-color:yellow;'"  ' Si coincide con el valor, cambia el fondo a amarillo
                     Else
              colorCelda7 = "style='background-color:#e0f7fa;'"  ' De lo contrario, deja el fondo blanco       
                End If
                If (suma7 >0) Then
               colorCelda8 = "style='background-color:yellow;'"  ' Si coincide con el valor, cambia el fondo a amarillo
                     Else
              colorCelda8 = "style='background-color:#e0f7fa;'"  ' De lo contrario, deja el fondo blanco       
                End If
                If (suma8 >0) Then
              colorCelda9 = "style='background-color:yellow;'"  ' Si coincide con el valor, cambia el fondo a amarillo
                     Else
              colorCelda9 = "style='background-color:#e0f7fa;'"  ' De lo contrario, deja el fondo blanco       
                End If
               If (suma9 >0) Then
               colorCelda10 = "style='background-color:yellow;'"  ' Si coincide con el valor, cambia el fondo a amarillo
                     Else
              colorCelda10 = "style='background-color:#e0f7fa;'"  ' De lo contrario, deja el fondo blanco       
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
                <th align='center' " + colorCelda + ">PO</font></th>
                <th align='center' " + colorCelda1 + ">OP</font></th>
                <th align='center' " + colorCelda2 + ">CODIGO</font></th>
                <th align='center' " + colorCelda3 + ">ANCHO UTILIZADO</font></th>
                <th align='center' " + colorCelda4 + ">DENSIDAD</font></th>
                <th align='center' " + colorCelda5 + "><EFICIENCIA</font></th>
                <th align='center' " + colorCelda6 + ">LARGO TIZADO</font></th>
                <th align='center' " + colorCelda7 + ">PRENDAS</font></th>
                <th align='center' " + colorCelda8 + ">MERMA TENDIDO</font></th>
                <th align='center' " + colorCelda9 + ">C. LINEAL</font></th>
                <th align='center' " + colorCelda10 + ">CONSUMO KGS</font></th>
            </tr>
        </thead>
        <tbody>"
        ' Revisar la cantidad de filas en el DataGridView
        For Each row As DataGridViewRow In DataGridView1.Rows
            If Not row.IsNewRow Then
                ' Capturar la información de cada fila y agregarla al cuerpo HTML
                Mailz &=
                    "<tr>" &
                 "<td align='center'>" + TextBox12.Text.ToString().Trim() + "</td>" &
                 "<td align='center'>" + TextBox1.Text.ToString().Trim() + "</td>" &
                 "<td align='center'>" + row.Cells(1).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(5).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(6).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(7).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(8).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(9).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(10).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(11).Value.ToString() + "</td>" &
                 "<td align='center'>" + row.Cells(12).Value.ToString() + "</td>" &
                 "</tr>"
            End If
        Next
        ' Finalizar el cuerpo del HTML
        Mailz &= " </tbody></table>"
        Dim Rs As SqlDataReader
        Dim sql As String = "select correo from lista_correos where formulario ='PRODUCCION          '"
        Dim cmd As New SqlCommand(sql, conx)
        Rs = cmd.ExecuteReader()
        Rs.Read()

        With message

            .From = New System.Net.Mail.MailAddress(usuario)
            '.From = New System.Net.Mail.MailAddress("admin@viannysac.com")
            '.To.Add(Rs(0).ToString().Trim())
            .To.Add("coordinadorit@viannysac.com")

            Rs.Close()
            .IsBodyHtml = True
            .Body = Mailz
            .Subject = DETALLE & "-" & TextBox1.Text
            .Priority = System.Net.Mail.MailPriority.Normal
        End With


        With smtp

            .EnableSsl = True
            .Port = "587"
            .Host = "smtp.gmail.com"
            '.Host = "smtppro.zoho.com"
            '.Credentials = New Net.NetworkCredential("admin@viannysac.com", "20Via$&it2")

            .Credentials = New Net.NetworkCredential(usuario, contras)
            .DeliveryMethod = SmtpDeliveryMethod.Network
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
    Sub actualizar_estado()
        Dim cmd16 As New SqlCommand("UPDATE custom_vianny.DBO.qag0300 Set ncorte_3='2' WHERE ncom_3 =@OP And ccia =@CCIA", conx)
        cmd16.Parameters.AddWithValue("@OP", Trim(TextBox1.Text))
        cmd16.Parameters.AddWithValue("@CCIA", Trim(Label8.Text))
        cmd16.ExecuteNonQuery()
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
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
    Dim COLUMNA As String
    Dim celdasModificadas As New List(Of String)
    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If Label12.Text = "1" Then
            If e.RowIndex >= 0 And e.ColumnIndex >= 0 Then      '
                Dim columnName As String = DataGridView1.Columns(e.ColumnIndex).HeaderText
                celdasModificadas.Add(columnName)

            End If
        End If

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)
        Dim numero As Integer
        DataGridView1.Rows.Add(1)
        numero = DataGridView1.Rows.Count
        DataGridView1.Rows(numero - 1).Cells(0).Value = numero
    End Sub
    Private WithEvents editingControl As Control
    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        'If editingControl IsNot Nothing Then
        '    RemoveHandler editingControl.KeyDown, AddressOf EditingControl_KeyDown
        'End If
    End Sub
    'Private Sub EditingControl_KeyDown(sender As Object, e As KeyEventArgs)
    '    ' Verificar si la tecla presionada es F1
    '    If e.KeyCode = Keys.F1 Then
    '        ' Verificar si la celda actual está en la columna "Factor de Consumo"

    '        If DataGridView1.CurrentCell.ColumnIndex = 1 Then
    '            Dim fechaActual As Date = Today

    '            pproductos.CCIA.Text = Label8.Text
    '            pproductos.ALMACEN.Text = "08"
    '            pproductos.ANO.Text = Year(fechaActual)
    '            pproductos.FECHA.Text = Replace(fechaActual.ToString("yyyy-MM-dd"), "-", "")
    '            pproductos.TextBox1.Text = "13"
    '            pproductos.Label3.Text = 12
    '            pproductos.Label5.Text = DataGridView1.CurrentCell.RowIndex.ToString
    '            pproductos.TextBox2.Text = ""
    '            pproductos.TextBox3.Text = ""
    '            pproductos.Show()
    '        End If
    '    End If
    'End Sub
    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim multi As Double
        Dim porcet As Double
        Dim suma, sumi As Double
        Dim Division, inverso As Double
        Dim resultado As String
        multi = (Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(5).Value) * Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(8).Value) * Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(6).Value))
        porcet = multi * (Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(10).Value) / 100)
        suma = multi + porcet
        Division = (suma / Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(9).Value)) / 1000
        resultado = Division.ToString("N4")

        sumi = 1000 * Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(11).Value) * Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(9).Value)

        'MsgBox(DataGridView1.Rows(e.RowIndex).Cells(5).Value)
        'MsgBox(DataGridView1.Rows(e.RowIndex).Cells(4).Value)
        'MsgBox(DataGridView1.Rows(e.RowIndex).Cells(6).Value)
        'MsgBox(multi.ToString)
        'MsgBox(DataGridView1.Rows(e.RowIndex).Cells(10).Value)
        'MsgBox(porcet.ToString)
        'MsgBox(suma.ToString)
        'MsgBox(Division.ToString)
        'MsgBox(resultado)
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "PRENDAS" Then
            DataGridView1.Rows(e.RowIndex).Cells(11).Value = (DataGridView1.Rows(e.RowIndex).Cells(8).Value / DataGridView1.Rows(e.RowIndex).Cells(9).Value) * (1 + (DataGridView1.Rows(e.RowIndex).Cells(10).Value) / 100)

            DataGridView1.Rows(e.RowIndex).Cells(12).Value = resultado
            Me.DataGridView1.Columns(11).DefaultCellStyle.Format = "n5"
            Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n4"
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "MERMA TENDIDO" Then
            DataGridView1.Rows(e.RowIndex).Cells(11).Value = (DataGridView1.Rows(e.RowIndex).Cells(8).Value / DataGridView1.Rows(e.RowIndex).Cells(9).Value) * (1 + (DataGridView1.Rows(e.RowIndex).Cells(10).Value) / 100)
            DataGridView1.Rows(e.RowIndex).Cells(12).Value = resultado
            Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n4"
            Me.DataGridView1.Columns(11).DefaultCellStyle.Format = "n5"
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "LARGO TIZADO" Then
            DataGridView1.Rows(e.RowIndex).Cells(11).Value = (DataGridView1.Rows(e.RowIndex).Cells(8).Value / DataGridView1.Rows(e.RowIndex).Cells(9).Value) * (1 + (DataGridView1.Rows(e.RowIndex).Cells(10).Value) / 100)
            DataGridView1.Rows(e.RowIndex).Cells(12).Value = resultado
            Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n4"
            Me.DataGridView1.Columns(11).DefaultCellStyle.Format = "n5"
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "ANCHO UTILIZADO" Then
            DataGridView1.Rows(e.RowIndex).Cells(11).Value = (DataGridView1.Rows(e.RowIndex).Cells(8).Value / DataGridView1.Rows(e.RowIndex).Cells(9).Value) * (1 + (DataGridView1.Rows(e.RowIndex).Cells(10).Value) / 100)
            DataGridView1.Rows(e.RowIndex).Cells(12).Value = resultado
            Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n4"
            Me.DataGridView1.Columns(11).DefaultCellStyle.Format = "n5"
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "DENSIDAD" Then
            DataGridView1.Rows(e.RowIndex).Cells(11).Value = (DataGridView1.Rows(e.RowIndex).Cells(8).Value / DataGridView1.Rows(e.RowIndex).Cells(9).Value) * (1 + (DataGridView1.Rows(e.RowIndex).Cells(10).Value) / 100)
            DataGridView1.Rows(e.RowIndex).Cells(12).Value = resultado
            Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n4"
            Me.DataGridView1.Columns(11).DefaultCellStyle.Format = "n5"
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "C.LINEAL" Then
            DataGridView1.Rows(e.RowIndex).Cells(8).Value = (DataGridView1.Rows(e.RowIndex).Cells(11).Value / (1 + (DataGridView1.Rows(e.RowIndex).Cells(10).Value) / 100)) * DataGridView1.Rows(e.RowIndex).Cells(9).Value
            Me.DataGridView1.Columns(8).DefaultCellStyle.Format = "n2"
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "CONSUMO KGS" Then
            DataGridView1.Rows(e.RowIndex).Cells(8).Value = ((sumi / (Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(10).Value) / 100))) / ((Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(5).Value) * Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(6).Value)))
            Me.DataGridView1.Columns(8).DefaultCellStyle.Format = "n2"
        End If
    End Sub

    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case TextBox9.Text.Length
                Case "1" : TextBox9.Text = "010000000" & "" & TextBox9.Text
                Case "2" : TextBox9.Text = "01000000" & "" & TextBox9.Text
                Case "3" : TextBox9.Text = "0100000" & "" & TextBox9.Text
                Case "4" : TextBox9.Text = "010000" & "" & TextBox9.Text
                Case "5" : TextBox9.Text = "01000" & "" & TextBox9.Text
                Case "6" : TextBox9.Text = "0100" & "" & TextBox9.Text
                Case "7" : TextBox9.Text = "010" & "" & TextBox9.Text
                Case "8" : TextBox9.Text = "01" & "" & TextBox9.Text
                Case "9" : TextBox9.Text = "0" & "" & TextBox9.Text
                Case "10" : TextBox9.Text = TextBox9.Text
            End Select
            abrir()
            Dim sql1021 As String = "EXEC custom_vianny.dbo.GetallFichaConsumoTextilCab '01','" + Trim(TextBox9.Text) + "',NULL"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr21 = cmd1021.ExecuteReader()
            Dim i5 As Integer
            i5 = 0
            If Rsr21.Read() = True Then
                TextBox2.Text = Trim(Rsr21(4))
                TextBox3.Text = Trim(Rsr21(5))
                TextBox4.Text = Trim(Rsr21(6))
                TextBox5.Text = Trim(Rsr21(7))
                TextBox6.Text = Trim(Rsr21(8))
                TextBox7.Text = Trim(Rsr21(9))

                TextBox1.Enabled = False
                Rsr21.Close()
                Dim sql10213 As String = "EXEC custom_vianny.dbo.GetallFichaConsumoTextilDet '01','" + Trim(TextBox9.Text) + "'"
                Dim cmd10213 As New SqlCommand(sql10213, conx)
                Rsr213 = cmd10213.ExecuteReader()
                While Rsr213.Read()
                    DataGridView1.Rows.Add()
                    DataGridView1.Rows(i5).Cells(0).Value = i5 + 1
                    DataGridView1.Rows(i5).Cells(1).Value = Rsr213(5)
                    DataGridView1.Rows(i5).Cells(2).Value = Rsr213(3)
                    DataGridView1.Rows(i5).Cells(3).Value = Rsr213(2)
                    DataGridView1.Rows(i5).Cells(4).Value = Rsr213(7)
                    DataGridView1.Rows(i5).Cells(5).Value = Rsr213(19)
                    DataGridView1.Rows(i5).Cells(6).Value = Rsr213(8)
                    DataGridView1.Rows(i5).Cells(7).Value = Rsr213(16)
                    DataGridView1.Rows(i5).Cells(8).Value = Rsr213(9)
                    DataGridView1.Rows(i5).Cells(9).Value = Rsr213(10)
                    DataGridView1.Rows(i5).Cells(10).Value = Rsr213(13)
                    DataGridView1.Rows(i5).Cells(11).Value = Rsr213(14)
                    DataGridView1.Rows(i5).Cells(12).Value = Rsr213(15)
                    DataGridView1.Rows(i5).Cells(13).Value = Rsr213(4)
                    i5 = i5 + 1

                End While
                Rsr213.Close()
                DataGridView1.ReadOnly = False
            End If
            TextBox9.Text = ""
        End If
    End Sub
End Class