Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.Net.Mail
Public Class Requerimiento__comercial
    Public conx As SqlConnection
    Public comando As SqlCommand
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Sub ENVIAR_CORREO2()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim fk As New fnopedido
        Dim jj As New vnapedido

        With message
            abrir()
            Dim Rsr1 As SqlDataReader
            Dim sql1 As String = "SELECT correo FROM lista_correos WHERE formulario ='REQUERIMIENTO_COM   '"
            Dim cmd10 As New SqlCommand(sql1, conx)
            Rsr1 = cmd10.ExecuteReader()
            Rsr1.Read()

            .From = New System.Net.Mail.MailAddress("admin@viannysac.com")

            .To.Add(Rsr1(0))
            .Body = "CLIENTE :" & Trim(TextBox1.Text) & " --- " & "PO :" & "  " & Trim(TextBox2.Text) & " --- " & " Nº Requerimiento :" & "   " & Trim(TextBox3.Text)
            If Label6.Text = "1" Then
                .Subject = "Se Actualizo el Requerimiento Textil" & "---" & TextBox3.Text
            Else
                .Subject = "Se Creo el Requerimiento Textil" & "---" & TextBox3.Text
            End If

            .Priority = System.Net.Mail.MailPriority.Normal
        End With

        With smtp
            .EnableSsl = True
            .Port = "587"
            .Host = "smtppro.zoho.com"
            .Credentials = New Net.NetworkCredential("admin@viannysac.com", "20Via$&it2")
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.Rows.Add(1)
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
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        abrir()
        If Me.ValidateChildren() And TextBox1.Text = String.Empty And TextBox3.Text = String.Empty Then
            MsgBox("FALTAN INGREAR ALGUNO CAMPOS OBLIGATORIOS")
        Else
            Dim io, cant1, cant2 As Integer

            io = DataGridView1.Rows.Count
            cant1 = 0
            cant2 = 0

            For i = 0 To io - 1
                If Trim(DataGridView1.Rows(i).Cells(0).Value) = "" Then
                    cant1 = cant1 + 1
                End If
                If Trim(DataGridView1.Rows(i).Cells(1).Value) = "" Then
                    cant2 = cant2 + 1
                End If
            Next

            If cant1 > 0 Or cant2 > 0 Then
                MsgBox("FALTA INGRESAR EL NOMBRE DEL ARTICULO O LA CANTIDAD EN UNA FILA DEL DETALLE REQUERIMIENTO")
            Else
                abrir()
                'Dim agregar As String = "DELETE from det_req_comercial WHERE nu_requerimientio ='" + Trim(TextBox3.Text) + "'"
                'Dim agregar1 As String = "DELETE FROM cab_req_comercial  WHERE nu_requerimientio ='" + Trim(TextBox3.Text) + "'"
                'ELIMINAR(agregar)
                'ELIMINAR(agregar1)
                Try
                    If Label6.Text = 1 Then

                        Dim cmd15 As New SqlCommand("update cab_req_comercial set cliente=@cliente,po=@po,tipo=@tipo,fecha=@fecha,VENDEDOR=@VENDEDOR where nu_requerimientio = @nu_requerimientio", conx)
                        cmd15.Parameters.AddWithValue("@nu_requerimientio", Trim(TextBox3.Text))
                        cmd15.Parameters.AddWithValue("@cliente", Trim(TextBox1.Text))
                        cmd15.Parameters.AddWithValue("@po", Trim(TextBox2.Text))
                        cmd15.Parameters.AddWithValue("@tipo", ComboBox1.Text)
                        cmd15.Parameters.AddWithValue("@fecha", DateTimePicker1.Value)
                        cmd15.Parameters.AddWithValue("@VENDEDOR", Trim(ComboBox2.Text))
                        cmd15.ExecuteNonQuery()

                    Else
                        Dim cmd15 As New SqlCommand("insert into cab_req_comercial (nu_requerimientio,cliente,po,tipo,fecha,VENDEDOR,estado) values (@nu_requerimientio,@cliente,@po,@tipo,@fecha,@VENDEDOR,@estado)", conx)
                        cmd15.Parameters.AddWithValue("@nu_requerimientio", Trim(TextBox3.Text))
                        cmd15.Parameters.AddWithValue("@cliente", Trim(TextBox1.Text))
                        cmd15.Parameters.AddWithValue("@po", Trim(TextBox2.Text))
                        cmd15.Parameters.AddWithValue("@tipo", ComboBox1.Text)
                        cmd15.Parameters.AddWithValue("@fecha", DateTimePicker1.Value)
                        cmd15.Parameters.AddWithValue("@VENDEDOR", Trim(ComboBox2.Text))
                        cmd15.Parameters.AddWithValue("@estado", "0")
                        cmd15.ExecuteNonQuery()
                    End If




                    Dim p As Integer
                    p = DataGridView1.Rows.Count
                    If Label6.Text = 1 Then
                        'Dim agregar As String = "delete from det_req_comercial where nu_requerimientio = " + Trim(TextBox3.Text) + ""
                        'ELIMINAR(agregar)
                        For i = 0 To p - 1
                            If Trim(DataGridView1.Rows(i).Cells(2).Value) = "" Then
                                Dim cmd16 As New SqlCommand("insert into det_req_comercial (nu_requerimientio,prducto,cantidad,estado) values (@nu_requerimientio,@prducto,@cantidad,@estado)", conx)
                                cmd16.Parameters.AddWithValue("@nu_requerimientio", Trim(TextBox3.Text))
                                cmd16.Parameters.AddWithValue("@prducto", Trim(DataGridView1.Rows(i).Cells(0).Value))
                                cmd16.Parameters.AddWithValue("@cantidad", Convert.ToDouble(DataGridView1.Rows(i).Cells(1).Value))
                                cmd16.Parameters.AddWithValue("@estado", "0")

                                cmd16.ExecuteNonQuery()
                            Else
                                Dim cmd16 As New SqlCommand("update  det_req_comercial set nu_requerimientio=@nu_requerimientio, prducto =@prducto,cantidad =@cantidad where det_req =@det", conx)
                                cmd16.Parameters.AddWithValue("@nu_requerimientio", Trim(TextBox3.Text))
                                cmd16.Parameters.AddWithValue("@prducto", Trim(DataGridView1.Rows(i).Cells(0).Value))
                                cmd16.Parameters.AddWithValue("@cantidad", Convert.ToDouble(DataGridView1.Rows(i).Cells(1).Value))
                                cmd16.Parameters.AddWithValue("@det", DataGridView1.Rows(i).Cells(2).Value)

                                cmd16.ExecuteNonQuery()
                            End If

                        Next
                        'Dim cmd168 As New SqlCommand("insert into det_req_comercial (nu_requerimientio,prducto,cantidad,estado) values (@nu_requerimientio,@prducto,@cantidad,@estado)", conx)
                        'cmd168.Parameters.AddWithValue("@nu_requerimientio", Trim(TextBox3.Text))
                        'cmd168.Parameters.AddWithValue("@prducto", Trim(DataGridView1.Rows(i).Cells(0).Value))
                        'cmd168.Parameters.AddWithValue("@cantidad", DataGridView1.Rows(i).Cells(1).Value)
                        'cmd168.Parameters.AddWithValue("@estado", "0")
                        'cmd168.ExecuteNonQuery()

                    Else

                        For i = 0 To p - 1
                            Dim cmd16 As New SqlCommand("insert into det_req_comercial (nu_requerimientio,prducto,cantidad,estado) values (@nu_requerimientio,@prducto,@cantidad,@estado)", conx)
                            cmd16.Parameters.AddWithValue("@nu_requerimientio", Trim(TextBox3.Text))
                            cmd16.Parameters.AddWithValue("@prducto", Trim(DataGridView1.Rows(i).Cells(0).Value))
                            cmd16.Parameters.AddWithValue("@cantidad", Convert.ToDouble(DataGridView1.Rows(i).Cells(1).Value))
                            cmd16.Parameters.AddWithValue("@estado", "0")

                            cmd16.ExecuteNonQuery()


                        Next
                    End If


                    MsgBox("SE INGRESO LA INFORMACION CORRECTAMENTE")
                    ENVIAR_CORREO2()
                    TextBox1.Text = ""
                    TextBox2.Text = ""
                    DataGridView1.Rows.Clear()
                    CORRELATIVO()


                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If


        End If


    End Sub
    Sub CORRELATIVO()
        abrir()
        Dim Rsr2 As SqlDataReader
        Dim sql102 As String = "SELECT TOP 1  nu_requerimientio   AS VAL FROM cab_req_comercial   order by nu_requerimientio desc"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        If Rsr2.Read() = True Then
            TextBox3.Text = Rsr2(0) + 1
        Else
            TextBox3.Text = 1
        End If

        Select Case Trim(TextBox3.Text).Length
            Case "1" : TextBox3.Text = "000000000" & "" & TextBox3.Text
            Case "2" : TextBox3.Text = "00000000" & "" & TextBox3.Text
            Case "3" : TextBox3.Text = "0000000" & "" & TextBox3.Text
            Case "4" : TextBox3.Text = "000000" & "" & TextBox3.Text
            Case "5" : TextBox3.Text = "00000" & "" & TextBox3.Text
            Case "6" : TextBox3.Text = "0000" & "" & TextBox3.Text
            Case "7" : TextBox3.Text = "000" & "" & TextBox3.Text
            Case "8" : TextBox3.Text = "00" & "" & TextBox3.Text
            Case "9" : TextBox3.Text = "0" & "" & TextBox3.Text
            Case "10" : TextBox3.Text = TextBox3.Text
        End Select

    End Sub
    Private Sub Requerimiento__comercial_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CORRELATIVO()
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_Validating(sender As Object, e As CancelEventArgs) Handles TextBox1.Validating
        If DirectCast(sender, TextBox).Text.Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR CLIENTE")
        End If
    End Sub


    Private Sub TextBox3_Validating(sender As Object, e As CancelEventArgs) Handles TextBox3.Validating
        If DirectCast(sender, TextBox).Text.Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR Nº REQUERIMIENTO")
        End If
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox3_Resize(sender As Object, e As EventArgs) Handles TextBox3.Resize

    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            DataGridView1.Rows.Clear()

            Select Case Trim(TextBox3.Text).Length
                Case "1" : TextBox3.Text = "000000000" & "" & TextBox3.Text
                Case "2" : TextBox3.Text = "00000000" & "" & TextBox3.Text
                Case "3" : TextBox3.Text = "0000000" & "" & TextBox3.Text
                Case "4" : TextBox3.Text = "000000" & "" & TextBox3.Text
                Case "5" : TextBox3.Text = "00000" & "" & TextBox3.Text
                Case "6" : TextBox3.Text = "0000" & "" & TextBox3.Text
                Case "7" : TextBox3.Text = "000" & "" & TextBox3.Text
                Case "8" : TextBox3.Text = "00" & "" & TextBox3.Text
                Case "9" : TextBox3.Text = "0" & "" & TextBox3.Text
                Case "10" : TextBox3.Text = TextBox3.Text
            End Select
            TextBox1.Select()
            abrir()
            Dim Rsr21, Rsr22, Rsr23 As SqlDataReader

            Dim sql1021 As String = "SELECT COUNT(nu_requerimientio) FROM cab_req_comercial  WHERE nu_requerimientio ='" + TextBox3.Text + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr21 = cmd1021.ExecuteReader()
            If Rsr21.Read() = True Then
                If Rsr21(0) > 0 Then

                    Rsr21.Close()
                    Button5.Visible = True
                    Dim sql1022 As String = "SELECT * FROM cab_req_comercial  WHERE nu_requerimientio ='" + TextBox3.Text + "'"
                    Dim cmd1022 As New SqlCommand(sql1022, conx)
                    Rsr22 = cmd1022.ExecuteReader()
                    Rsr22.Read()
                    TextBox1.Text = Trim(Rsr22(1))
                    TextBox2.Text = Rsr22(2)
                    ComboBox1.Text = Rsr22(3)
                    DateTimePicker1.Text = Rsr22(4)
                    If Rsr22(6) = 2 Then
                        Label8.Visible = True
                    Else
                        Label8.Visible = False
                    End If
                    Rsr22.Close()
                    Dim sql1023 As String = "select * from det_req_comercial WHERE nu_requerimientio ='" + TextBox3.Text + "'"
                    Dim cmd1023 As New SqlCommand(sql1023, conx)
                    Rsr23 = cmd1023.ExecuteReader()
                    Dim i51 As Integer
                    i51 = 0
                    While Rsr23.Read()
                        DataGridView1.Rows.Add()
                        DataGridView1.Rows(i51).Cells(0).Value = Trim(Rsr23(2))
                        DataGridView1.Rows(i51).Cells(1).Value = Trim(Rsr23(3))
                        DataGridView1.Rows(i51).Cells(2).Value = Rsr23(0)
                        i51 = i51 + 1
                    End While
                    Rsr23.Close()
                Else
                    TextBox1.Select()
                    Button5.Visible = False
                End If
                TextBox1.Enabled = False
                TextBox2.Enabled = False
                ComboBox1.Enabled = False
                DataGridView1.Enabled = False
                Button1.Enabled = False
                Button2.Enabled = False
                Button4.Enabled = False
            Else
                TextBox1.Enabled = True
                TextBox2.Enabled = True
                ComboBox1.Enabled = True
                DataGridView1.Enabled = True
                Button1.Enabled = True
                Button2.Enabled = True
                Button4.Enabled = True
            End If

        End If
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.F1 Then

            Clientes.TextBox3.Text = "3363"
            Clientes.Show()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim respuesta As DialogResult
        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
            Dim fila As String
            fila = DataGridView1.CurrentRow.Index

            Dim agregar As String = "delete from det_req_comercial where det_req = '" + Trim(DataGridView1.Rows(fila).Cells(2).Value) + "'"
            ELIMINAR(agregar)
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Label6.Text = 1
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        ComboBox1.Enabled = True
        DataGridView1.Enabled = True
        Button1.Enabled = True
        Button2.Enabled = True
        Button4.Enabled = True
    End Sub

    Private Sub TextBox3_MultilineChanged(sender As Object, e As EventArgs) Handles TextBox3.MultilineChanged

    End Sub

    Private Sub TextBox3_DoubleClick(sender As Object, e As EventArgs) Handles TextBox3.DoubleClick
        TextBox1.Text = ""
        TextBox2.Text = ""
        DataGridView1.Rows.Clear()
        CORRELATIVO()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        abrir()
        Dim cmd151 As New SqlCommand("update cab_req_comercial set estado=@ESTADO where nu_requerimientio = @nu_requerimientio", conx)
        cmd151.Parameters.AddWithValue("@nu_requerimientio", Trim(TextBox3.Text))
        cmd151.Parameters.AddWithValue("@ESTADO", 2)
        cmd151.ExecuteNonQuery()

        Dim cmd152 As New SqlCommand("update det_req_comercial set estado=@ESTADO where nu_requerimientio = @nu_requerimientio", conx)
        cmd152.Parameters.AddWithValue("@nu_requerimientio", Trim(TextBox3.Text))
        cmd152.Parameters.AddWithValue("@ESTADO", 2)
        cmd152.ExecuteNonQuery()
        MsgBox("SE ANULO LO SOLICITADO")
        TextBox1.Text = ""
        TextBox2.Text = ""
        DataGridView1.Rows.Clear()
        CORRELATIVO()
    End Sub
End Class