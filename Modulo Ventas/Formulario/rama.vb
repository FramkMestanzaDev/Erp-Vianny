Imports System.ComponentModel
Imports System.Net.Mail
Imports System.Data.SqlClient
Public Class rama
    Dim dt As New DataTable
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
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then

            Try
                abrir()
                Dim Rs, Rs1 As SqlDataReader
                Dim sql As String = "SELECT COUNT(partida) FROM validar_partida WHERE PARTIDA ='" + TextBox1.Text + "' AND estado ='1' and empresa='" + Label8.Text + "' "
                Dim cmd As New SqlCommand(sql, conx)
                Rs = cmd.ExecuteReader()
                Rs.Read()
                If Rs(0) > 0 Then
                    Dim gy As New frama
                    Dim ko As New vrama
                    Dim L As Integer
                    Dim num, num2, num3, num10, ji As Integer
                    Dim sum As Double
                    Dim sql1 As String = "SELECT ,rollospesados,kgsobtenidos FROM validar_partida WHERE PARTIDA ='" + TextBox1.Text + "' AND estado ='1' and empresa='" + Label8.Text + "' "
                    Dim cmd1 As New SqlCommand(sql1, conx)
                    Rs1 = cmd1.ExecuteReader()
                    Rs1.Read()


                    ko.gpartida = TextBox1.Text
                    ko.gccia = Label8.Text
                    num = 0
                    num3 = 0
                    dt = gy.buscar_partida(ko)
                    If dt.Rows.Count > 0 Then

                        DataGridView2.DataSource = dt
                        L = DataGridView2.Rows.Count
                        DataGridView1.Rows.Add(1)
                        If DataGridView1.Rows(0).Cells(0).Value = 0 Then
                            DataGridView1.Rows(0).Cells(0).Value = num + 1
                            DataGridView1.Rows(0).Cells(1).Value = DataGridView2.Rows(0).Cells(0).Value
                            DataGridView1.Rows(0).Cells(2).Value = DataGridView2.Rows(0).Cells(1).Value
                            DataGridView1.Rows(0).Cells(3).Value = DataGridView2.Rows(0).Cells(2).Value
                            DataGridView1.Rows(0).Cells(4).Value = DataGridView2.Rows(0).Cells(3).Value
                            DataGridView1.Rows(0).Cells(5).Value = DataGridView2.Rows(0).Cells(4).Value
                            DataGridView1.Rows(0).Cells(6).Value = Rs1(0)
                            DataGridView1.Rows(0).Cells(7).Value = Rs1(1)
                        Else
                            num10 = DataGridView1.Rows.Count
                            For i1 = 1 To num10 - 1
                                If Convert.ToInt16(DataGridView1.Rows(i1).Cells(0).Value) < Convert.ToInt16(DataGridView1.Rows(i1 - 1).Cells(0).Value) Then
                                    num2 = Convert.ToInt16(DataGridView1.Rows(i1 - 1).Cells(0).Value) + 1
                                    num3 = i1
                                End If

                            Next
                            For i2 = 0 To num10 - 1
                                If Trim(TextBox1.Text) = Trim(DataGridView1.Rows(i2).Cells(2).Value) Then
                                    ji = ji + 1

                                End If
                            Next

                            If ji >= 1 Then
                                MsgBox("LA PARTIDA YA ESTA REGISTRADA")
                                DataGridView1.Rows.RemoveAt(num3)

                            Else
                                DataGridView1.Rows(num3).Cells(0).Value = num2
                                DataGridView1.Rows(num3).Cells(1).Value = DataGridView2.Rows(0).Cells(0).Value
                                DataGridView1.Rows(num3).Cells(2).Value = DataGridView2.Rows(0).Cells(1).Value
                                DataGridView1.Rows(num3).Cells(3).Value = DataGridView2.Rows(0).Cells(2).Value
                                DataGridView1.Rows(num3).Cells(4).Value = DataGridView2.Rows(0).Cells(3).Value
                                DataGridView1.Rows(num3).Cells(5).Value = DataGridView2.Rows(0).Cells(4).Value
                                DataGridView1.Rows(num3).Cells(6).Value = DataGridView2.Rows(0).Cells(5).Value
                                DataGridView1.Rows(num3).Cells(7).Value = DataGridView2.Rows(0).Cells(6).Value
                            End If

                        End If
                        SUMAR()
                        Button2.Enabled = False
                    Else
                        MsgBox("LA PARTIDA NO EXISTE")
                    End If
                    TextBox1.Text = ""
                    DataGridView2.DataSource = ""
                Else
                    MsgBox("LA PARTIDA NO ESTA VALIDADA PARA PASAR POR RAMA")

                End If


            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If

    End Sub
    Sub SUMAR()
        Dim jh As Integer
        Dim SUM As Double
        jh = DataGridView1.Rows.Count
        For i5 = 0 To jh - 1
            sum = sum + DataGridView1.Rows(i5).Cells(6).Value
        Next
        TextBox2.Text = sum
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "PESO" Then
            Dim SUM As Double
            Dim J As Integer
            J = DataGridView1.Rows.Count
            For I = 0 To J - 1
                SUM = SUM + DataGridView1.Rows(I).Cells(6).Value
            Next
            TextBox2.Text = SUM
        End If
    End Sub
    Sub ENVIAR_CORREO()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient

        Dim Mailz As String = "" &
      " <!DOCTYPE html>" &
"<html xmlns='http://www.w3.org/1999/xhtml'>" &
"<head>" &
"    <title></title>" &
"</head>" &
"<body>
     <FONT COLOR='green'>INFORMACION DEL PROGRAMA RAMA</FONT><br/><br/>

<FONT COLOR='blue'>* ESTADO : </FONT> CREADO <br/>
<FONT COLOR='blue'>* Nº PROGRAMA : </FONT>" + Trim(TextBox1.Text) + " <br/>
<FONT COLOR='blue'>* KILOS : </FONT>" + Trim(TextBox2.Text) + " <br/>
<FONT COLOR='blue'>* FECHA RAMEAR : </FONT>" + DateTimePicker1.Value + " <br/>
<FONT COLOR='blue'>* GENERADO POR : </FONT>" + Trim(Label5.Text) + "<br/>

</body>

</html>"
        With message
            .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
            .To.Add("fmestanza@viannysac.com,asilva@viannysac.com,hfarje@viannysac.com,irodriguez@viannysac.com")
            .IsBodyHtml = True
            .Body = Mailz
            .Subject = "Programa de Rama N°" & TextBox3.Text
            .Priority = System.Net.Mail.MailPriority.Normal
        End With
        With smtp
            .EnableSsl = False
            .Port = "25"
            .Host = "mail.onehostingperu.com"
            .Credentials = New Net.NetworkCredential("admin@viannysac.com", "$i$tema$$i$tem@$")
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
        Dim Mailz As String = "" &
      " <!DOCTYPE html>" &
"<html xmlns='http://www.w3.org/1999/xhtml'>" &
"<head>" &
"    <title></title>" &
"</head>" &
"<body>
     <FONT COLOR='green'>INFORMACION DEL PROGRAMA RAMA</FONT><br/><br/>

<FONT COLOR='blue'>* ESTADO : </FONT> ACTUALIZADO <br/>
<FONT COLOR='blue'>* Nº PROGRAMA : </FONT>" + Trim(TextBox1.Text) + " <br/>
<FONT COLOR='blue'>* KILOS : </FONT>" + Trim(TextBox2.Text) + " <br/>
<FONT COLOR='blue'>* FECHA RAMEAR : </FONT>" + DateTimePicker1.Value + " <br/>
<FONT COLOR='blue'>* GENERADO POR : </FONT>" + Trim(Label5.Text) + "<br/>

</body>

</html>"
        With message
            .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
            .To.Add("fmestanza@viannysac.com,asilva@viannysac.com,hfarje@viannysac.com,irodriguez@viannysac.com")
            .IsBodyHtml = True
            .Body = Mailz
            .Subject = "Actualizacion del Programa de Rama N°" & TextBox3.Text
            .Priority = System.Net.Mail.MailPriority.Normal
        End With
        With smtp
            .EnableSsl = False
            .Port = "25"
            .Host = "mail.onehostingperu.com"
            .Credentials = New Net.NetworkCredential("admin@viannysac.com", "$i$tema$$i$tem@$")
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
    Sub ENVIAR_CORREO3()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        With message
            .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
            .To.Add("fmestanza@viannysac.com,asilva@viannysac.com,hfarje@viannysac.com,irodriguez@viannysac.com")
            .Body = "Se CERRO El Programa de Rama"
            .Subject = "Programa de Rama N°" & TextBox3.Text
            .Priority = System.Net.Mail.MailPriority.Normal
        End With
        With smtp
            .EnableSsl = False
            .Port = "25"
            .Host = "mail.onehostingperu.com"
            .Credentials = New Net.NetworkCredential("admin@viannysac.com", "$i$tema$$i$tem@$")
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
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        MsgBox("SE EDITARA LA INFORACION DE PROGRAMA DE RAMA   " & TextBox3.Text)
        TextBox4.Text = "2"
        DataGridView1.Enabled = True
        Button2.Enabled = True
        Button1.Enabled = True
        Button3.Enabled = True
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
    Private Sub rama_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Dim jj1 As New frama
        Dim ooo As New vrama

        ooo.gccia = Label8.Text

        Me.TextBox3.Text = jj1.correlativo_rama(ooo) + 1
        Select Case TextBox3.Text.Length
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
    End Sub

    Dim dy As New DataTable

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim jh As New frama
            Dim ll As New vrama
            Dim gg As Integer
            Dim sum As Double
            Select Case TextBox3.Text.Length
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
            ll.gcodigo = TextBox3.Text
            ll.gccia = Label8.Text
            dy = jh.buscar_codigo_rama(ll)

            If dy.Rows.Count > 0 Then
                DataGridView1.Rows.Clear()
                DataGridView3.DataSource = dy
                gg = DataGridView3.Rows.Count

                DataGridView1.Rows.Add(gg)

                For i = 0 To gg - 1

                    DataGridView1.Rows(i).Cells(0).Value = DataGridView3.Rows(i).Cells(2).Value
                    DataGridView1.Rows(i).Cells(1).Value = DataGridView3.Rows(i).Cells(3).Value
                    DataGridView1.Rows(i).Cells(2).Value = DataGridView3.Rows(i).Cells(4).Value
                    DataGridView1.Rows(i).Cells(3).Value = DataGridView3.Rows(i).Cells(5).Value
                    DataGridView1.Rows(i).Cells(4).Value = DataGridView3.Rows(i).Cells(6).Value
                    DataGridView1.Rows(i).Cells(5).Value = DataGridView3.Rows(i).Cells(7).Value
                    DataGridView1.Rows(i).Cells(6).Value = DataGridView3.Rows(i).Cells(8).Value
                    DataGridView1.Rows(i).Cells(7).Value = DataGridView3.Rows(i).Cells(9).Value
                    DataGridView1.Rows(i).Cells(8).Value = DataGridView3.Rows(i).Cells(10).Value
                    DataGridView1.Rows(i).Cells(9).Value = DataGridView3.Rows(i).Cells(11).Value
                    DataGridView1.Rows(i).Cells(10).Value = DataGridView3.Rows(i).Cells(17).Value
                    DataGridView1.Rows(i).Cells(11).Value = DataGridView3.Rows(i).Cells(18).Value
                    DataGridView1.Rows(i).Cells(12).Value = DataGridView3.Rows(i).Cells(19).Value
                    DataGridView1.Rows(i).Cells(13).Value = DataGridView3.Rows(i).Cells(20).Value
                    DataGridView1.Rows(i).Cells(14).Value = DataGridView3.Rows(i).Cells(21).Value
                    DataGridView1.Rows(i).Cells(15).Value = DataGridView3.Rows(i).Cells(22).Value
                    DataGridView1.Rows(i).Cells(16).Value = DataGridView3.Rows(i).Cells(23).Value
                    DataGridView1.Rows(i).Cells(17).Value = DataGridView3.Rows(i).Cells(25).Value
                    DataGridView1.Rows(i).Cells(18).Value = DataGridView3.Rows(i).Cells(26).Value
                    DataGridView1.Rows(i).Cells(19).Value = DataGridView3.Rows(i).Cells(27).Value
                    DataGridView1.Columns(17).Visible = False
                    DataGridView1.Columns(18).Visible = False
                    sum = sum + DataGridView1.Rows(i).Cells(6).Value
                Next
                TextBox2.Text = sum
                DateTimePicker1.Value = DataGridView3.Rows(0).Cells(16).Value
                DataGridView1.Enabled = False
                Button2.Enabled = False
                Button1.Enabled = False
                Button3.Enabled = False
                If DataGridView3.Rows(0).Cells(12).Value = "2" Then
                    Label7.Text = "CERRADO"
                    PictureBox2.Enabled = False
                    Button5.Enabled = False
                Else
                    Label7.Text = "ACTIVO"
                    PictureBox2.Enabled = True
                    Button5.Enabled = True
                End If
            Else
                DataGridView1.Rows.Clear()
                PictureBox2.Enabled = True
            End If

        End If
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Dim respuesta As DialogResult
        Dim vv As New vrama
        Dim jj As New frama

        respuesta = MessageBox.Show("QUIERES ANULAR NOTA DE PEDIDO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            vv.gcodigo = TextBox3.Text
            vv.gccia = Label8.Text
            jj.ANULAR_rama(vv)
            MsgBox("SE REALIZO LO SOLICITADO")
            Me.Close()

        End If
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click

        Rpt_Rama.TextBox1.Text = TextBox3.Text
        Rpt_Rama.Show()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        SUMAR()
        'Me.DataGridView1.Sort(colum, System.ComponentModel.ListSortDirection.Ascending)
        'Me.Column1.SortMode = DataGridViewColumnSortMode.Programmatic
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        PARTIDA_RAMA.Show()
    End Sub
    Public Sub NumConFrac1(ByVal CajaTexto As Windows.Forms.TextBox, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Char.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False

        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        NumConFrac1(TextBox1, e)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim hj As New vrama
        Dim kk As New vrama
        Dim jj As New frama
        Dim ml As Integer
        If TextBox4.Text = 2 Then
            If TextBox3.Text = "" Then
                MsgBox("FALTA INGRESAR EL NUMERO DE PROGRAMA DE RAMA")
            Else

                Try
                    Dim nombre As String
                    'Entrada de datos mediante un inputbox
                    nombre = InputBox("Ingrese Las Observaciones",
                                 "Registro De Informacion De Rama",
                                 " ", 200, 0)
                    kk.gcodigo = TextBox3.Text
                    kk.gccia = Label8.Text
                    jj.eliminar_rama(kk)
                    ml = DataGridView1.Rows.Count
                    For i = 0 To ml - 1

                        hj.gcodigo = TextBox3.Text

                        If Convert.ToString(DataGridView1.Rows(i).Cells(0).Value) = "" Then
                            hj.gorden = ""
                        Else
                            hj.gorden = DataGridView1.Rows(i).Cells(0).Value
                        End If
                        If Convert.ToString(DataGridView1.Rows(i).Cells(1).Value) = "" Then
                            hj.gpo = ""
                        Else
                            hj.gpo = DataGridView1.Rows(i).Cells(1).Value
                        End If
                        If Convert.ToString(DataGridView1.Rows(i).Cells(2).Value) = "" Then
                            hj.gpartida = ""
                        Else
                            hj.gpartida = DataGridView1.Rows(i).Cells(2).Value
                        End If
                        If Convert.ToString(DataGridView1.Rows(i).Cells(3).Value) = "" Then
                            hj.gcolor = ""
                        Else
                            hj.gcolor = DataGridView1.Rows(i).Cells(3).Value
                        End If
                        If Convert.ToString(DataGridView1.Rows(i).Cells(4).Value) = "" Then
                            hj.garticulo = ""
                        Else
                            hj.garticulo = DataGridView1.Rows(i).Cells(4).Value
                        End If
                        If Convert.ToString(DataGridView1.Rows(i).Cells(5).Value) = "" Then
                            hj.grollos = ""
                        Else
                            hj.grollos = DataGridView1.Rows(i).Cells(5).Value
                        End If
                        If Convert.ToString(DataGridView1.Rows(i).Cells(6).Value) = "" Then
                            hj.gpeso = "0.00"
                        Else
                            hj.gpeso = DataGridView1.Rows(i).Cells(6).Value
                        End If
                        If Convert.ToString(DataGridView1.Rows(i).Cells(7).Value) = "" Then
                            hj.glote = ""
                        Else
                            hj.glote = DataGridView1.Rows(i).Cells(7).Value
                        End If
                        If Convert.ToString(DataGridView1.Rows(i).Cells(8).Value) = "" Then
                            hj.gancho = ""
                        Else
                            hj.gancho = DataGridView1.Rows(i).Cells(8).Value
                        End If
                        If Convert.ToString(DataGridView1.Rows(i).Cells(9).Value) = "" Then
                            hj.gdensidad = ""
                        Else
                            hj.gdensidad = DataGridView1.Rows(i).Cells(9).Value
                        End If
                        hj.gflag = 1
                        hj.gpeso_total = TextBox2.Text
                        hj.gobservacion = nombre
                        hj.gfecha = DateTimePicker3.Value.ToString("dd/MM/yyyy")
                        hj.gfecha_rama = DateTimePicker1.Value.ToString("dd/MM/yyyy")
                        If Convert.ToString(DataGridView1.Rows(i).Cells(10).Value) = "-" Then
                            hj.ghora = "-"
                        Else
                            hj.ghora = DataGridView1.Rows(i).Cells(10).Value
                        End If
                        If Convert.ToString(DataGridView1.Rows(i).Cells(11).Value) = 0 Then
                            hj.gpasado = 0
                        Else
                            hj.gpasado = DataGridView1.Rows(i).Cells(11).Value
                        End If
                        If Convert.ToString(DataGridView1.Rows(i).Cells(12).Value) = "-" Then
                            hj.ghora2 = "-"
                        Else
                            hj.ghora2 = DataGridView1.Rows(i).Cells(12).Value
                        End If
                        If Convert.ToString(DataGridView1.Rows(i).Cells(13).Value) = "0.00" Then
                            hj.ganchor = "0.00"
                        Else
                            hj.ganchor = DataGridView1.Rows(i).Cells(13).Value
                        End If
                        If Convert.ToString(DataGridView1.Rows(i).Cells(14).Value) = "0" Then
                            hj.gdensidadr = "0.00"
                        Else
                            hj.gdensidadr = DataGridView1.Rows(i).Cells(14).Value
                        End If
                        If Convert.ToString(DataGridView1.Rows(i).Cells(15).Value) = "0.00" Then
                            hj.gvelocidad = "0.00"
                        Else
                            hj.gvelocidad = DataGridView1.Rows(i).Cells(15).Value
                        End If
                        If Convert.ToString(DataGridView1.Rows(i).Cells(16).Value) = "0.00" Then
                            hj.gtemperatura = "0.00"
                        Else
                            hj.gtemperatura = DataGridView1.Rows(i).Cells(16).Value
                        End If

                        If Convert.ToString(DataGridView1.Rows(i).Cells(17).Value) = "0.00" Then
                            hj.gsalimentacion = "0.00"
                        Else
                            hj.gsalimentacion = DataGridView1.Rows(i).Cells(17).Value
                        End If
                        If Trim(Convert.ToString(DataGridView1.Rows(i).Cells(18).Value)) = "" Then
                            hj.gdpesado = ""
                        Else
                            hj.gdpesado = DataGridView1.Rows(i).Cells(18).Value
                        End If
                        If Trim(Convert.ToString(DataGridView1.Rows(i).Cells(19).Value)) = "" Then
                            hj.gmaquina = ""
                        Else
                            hj.gmaquina = DataGridView1.Rows(i).Cells(19).Value
                        End If
                        hj.gigsa = "0"
                        hj.gccia = Label8.Text
                        If jj.insertar_rama(hj) Then
                            abrir()
                            Dim cmd15 As New SqlCommand("UPDATE validar_partida SET estado ='2' WHERE partida=@partida and empresa=@empresa", conx)
                            cmd15.Parameters.AddWithValue("@partida", hj.gpartida)
                            cmd15.Parameters.AddWithValue("@empresa", hj.gccia)
                            cmd15.ExecuteNonQuery()
                        End If

                    Next
                    Button2.Enabled = True

                    MsgBox("SE ACTUALIZO LA INFORMACION CON EXITO")
                    'ENVIAR_CORREO2()
                    Dim jh As New Rpt_Rama
                    jh.TextBox1.Text = TextBox3.Text
                    jh.Show()
                    DataGridView1.Rows.Clear()
                    TextBox1.Select()
                    Dim jj1 As New frama
                    Dim kkk As New vrama

                    kkk.gccia = Label8.Text
                    Me.TextBox3.Text = jj1.correlativo_rama(kkk) + 1
                    Select Case TextBox3.Text.Length
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
                Catch ex As Exception
                    MsgBox("ERROR AL GUARDAR LA INFORMACION")
                End Try
            End If
        Else
            If TextBox4.Text = 1 Then
                If TextBox3.Text = "" Then
                    MsgBox("FALTA INGRESAR EL NUMERO DE PROGRAMA DE RAMA")
                Else
                    Try
                        ml = DataGridView1.Rows.Count
                        kk.gcodigo = TextBox3.Text
                        kk.gccia = Label8.Text
                        jj.eliminar_rama(kk)
                        For i = 0 To ml - 1

                            hj.gcodigo = TextBox3.Text

                            If Convert.ToString(DataGridView1.Rows(i).Cells(0).Value) = "" Then
                                hj.gorden = ""
                            Else
                                hj.gorden = DataGridView1.Rows(i).Cells(0).Value
                            End If
                            If Convert.ToString(DataGridView1.Rows(i).Cells(1).Value) = "" Then
                                hj.gpo = ""
                            Else
                                hj.gpo = DataGridView1.Rows(i).Cells(1).Value
                            End If
                            If Convert.ToString(DataGridView1.Rows(i).Cells(2).Value) = "" Then
                                hj.gpartida = ""
                            Else
                                hj.gpartida = DataGridView1.Rows(i).Cells(2).Value
                            End If
                            If Convert.ToString(DataGridView1.Rows(i).Cells(3).Value) = "" Then
                                hj.gcolor = ""
                            Else
                                hj.gcolor = DataGridView1.Rows(i).Cells(3).Value
                            End If
                            If Convert.ToString(DataGridView1.Rows(i).Cells(4).Value) = "" Then
                                hj.garticulo = ""
                            Else
                                hj.garticulo = DataGridView1.Rows(i).Cells(4).Value
                            End If
                            If Convert.ToString(DataGridView1.Rows(i).Cells(5).Value) = "" Then
                                hj.grollos = ""
                            Else
                                hj.grollos = DataGridView1.Rows(i).Cells(5).Value
                            End If
                            If Convert.ToString(DataGridView1.Rows(i).Cells(6).Value) = "" Then
                                hj.gpeso = ""
                            Else
                                hj.gpeso = DataGridView1.Rows(i).Cells(6).Value
                            End If
                            If Convert.ToString(DataGridView1.Rows(i).Cells(7).Value) = "" Then
                                hj.glote = ""
                            Else
                                hj.glote = DataGridView1.Rows(i).Cells(7).Value
                            End If
                            If Convert.ToString(DataGridView1.Rows(i).Cells(8).Value) = "" Then
                                hj.gancho = ""
                            Else
                                hj.gancho = DataGridView1.Rows(i).Cells(8).Value
                            End If
                            If Convert.ToString(DataGridView1.Rows(i).Cells(9).Value) = "" Then
                                hj.gdensidad = ""
                            Else
                                hj.gdensidad = DataGridView1.Rows(i).Cells(9).Value
                            End If
                            If Trim(Convert.ToString(DataGridView1.Rows(i).Cells(19).Value)) = "" Then
                                hj.gmaquina = ""
                            Else
                                hj.gmaquina = DataGridView1.Rows(i).Cells(19).Value
                            End If
                            hj.gflag = 1
                            hj.gpeso_total = TextBox2.Text
                            hj.gobservacion = ""
                            hj.gfecha = DateTimePicker3.Value.ToString("dd/MM/yyyy")
                            hj.gfecha_rama = DateTimePicker1.Value.ToString("dd/MM/yyyy")
                            hj.ghora = "-"
                            hj.gpasado = 0
                            hj.ghora2 = "-"
                            hj.gccia = Label8.Text
                            If jj.insertar_rama(hj) Then
                                abrir()
                                Dim cmd15 As New SqlCommand("UPDATE validar_partida SET estado ='2' WHERE partida=@partida and empresa =@empresa", conx)
                                cmd15.Parameters.AddWithValue("@partida", hj.gpartida)
                                cmd15.Parameters.AddWithValue("@empresa", hj.gccia)
                                cmd15.ExecuteNonQuery()
                            End If

                        Next
                        MsgBox("SE REGISTRO LA INFORMACION CON EXITO")
                        ENVIAR_CORREO()
                        Dim jh As New Rpt_Rama
                        jh.TextBox1.Text = TextBox3.Text
                        jh.Show()
                        TextBox2.Text = ""
                        DataGridView1.Rows.Clear()
                        Dim jj1 As New frama
                        Dim jju As New vrama

                        jju.gccia = Label8.Text

                        Me.TextBox3.Text = jj1.correlativo_rama(jju) + 1
                        Select Case TextBox3.Text.Length
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
                    Catch ex As Exception
                        MsgBox("ERROR AL GUARDAR LA INFORMACION")
                    End Try

                End If
            End If

        End If

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim respuesta As DialogResult
        Dim I18, VAL, TAB As Integer
        TAB = DataGridView1.Rows.Count
        If TAB <> 0 Then
            respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then

                DataGridView1.Rows.Remove(DataGridView1.CurrentRow)

                I18 = DataGridView1.Rows.Count

                For i1 = 0 To I18 - 1
                    DataGridView1.Rows(i1).Cells(0).Value = i1 + 1
                Next

            End If
        Else
            MsgBox("NO HAY PRODUCTOS A ELIMINAR")
        End If

    End Sub



    Private Sub TextBox3_DoubleClick(sender As Object, e As EventArgs) Handles TextBox3.DoubleClick
        'Dim jj1 As New frama

        'Me.TextBox3.Text = jj1.correlativo_rama + 1
        'Select Case TextBox3.Text.Length
        '    Case "1" : TextBox3.Text = "000000000" & "" & TextBox3.Text
        '    Case "2" : TextBox3.Text = "00000000" & "" & TextBox3.Text
        '    Case "3" : TextBox3.Text = "0000000" & "" & TextBox3.Text
        '    Case "4" : TextBox3.Text = "000000" & "" & TextBox3.Text
        '    Case "5" : TextBox3.Text = "00000" & "" & TextBox3.Text
        '    Case "6" : TextBox3.Text = "0000" & "" & TextBox3.Text
        '    Case "7" : TextBox3.Text = "000" & "" & TextBox3.Text
        '    Case "8" : TextBox3.Text = "00" & "" & TextBox3.Text
        '    Case "9" : TextBox3.Text = "0" & "" & TextBox3.Text
        '    Case "10" : TextBox3.Text = TextBox3.Text
        'End Select
        'TextBox1.Select()
        'DataGridView1.DataSource = ""
        'TextBox2.Text = ""
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub DataGridView1_SortCompare(sender As Object, e As DataGridViewSortCompareEventArgs) Handles DataGridView1.SortCompare
        If Double.Parse(e.CellValue1) > Double.Parse(e.CellValue2) Then
            e.SortResult = 1
        ElseIf Double.Parse(e.CellValue1) < Double.Parse(e.CellValue2) Then
            e.SortResult = -1
        Else
            e.SortResult = 0
        End If
        e.Handled = True
    End Sub
    Private Sub Grid_SortCompare(
ByVal sender As Object, ByVal e As DataGridViewSortCompareEventArgs) _
Handles DataGridView1.SortCompare

        e.SortResult = CompareEx(e.CellValue1.ToString, e.CellValue2.ToString)
        e.Handled = True
        Exit Sub
    End Sub
    Private Function CompareEx(ByVal s1 As Object, ByVal s2 As Object) As Integer

        Try
            ' convert the objects to string if possible.
            Dim a As String = CType(s1, String)
            Dim b As String = CType(s2, String)

            ' If the values are the same, then return 0 to indicate as much
            If s1 = s2 Then Return 0

            ' Look to see if either of the values are numbers
            Dim IsNum1 As Boolean = IsNumeric(a)
            Dim IsNum2 As Boolean = IsNumeric(b)

            ' If both values are numeric, then do a numeric compare
            If IsNum1 And IsNum2 Then
                If Double.Parse(s1) > Double.Parse(s2) Then
                    Return 1
                ElseIf Double.Parse(s1) < Double.Parse(s2) Then
                    Return -1
                Else
                    Return 0
                End If
                ' If the first value is a number, but the second is not, then assume the number is "higher in the list"
            ElseIf IsNum1 And IsNum2 = False Then
                Return -1
                ' if the first values is not a number, but the second is, then assume the number is higher
            ElseIf IsNum1 = False And IsNum2 Then
                Return 1
            Else
                ' If both values are non nuermic, then do a normal string compare
                Return String.Compare(s1, s2)
            End If
        Catch ex As Exception
            'Console.WriteLine(ex.ToString)
        End Try

        ' If we got here, then we failed to compare, so return 0
        Return 0
    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        DataGridView1.Sort(DataGridView1.Columns(0), ListSortDirection.Ascending)

        Dim k As Integer
        k = DataGridView1.Rows.Count

        For i = 0 To k - 1
            DataGridView1.Rows(i).Cells(0).Value = i + 1
        Next
    End Sub

    Private Sub PictureBox1_Click_1(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim JO As New Exportar
        JO.llenarExcel(DataGridView1)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim j As Integer
        j = DataGridView1.Rows.Count

        abrir()
        Dim cmd As New SqlCommand("update rama set flag='2' where codigo =@numero and ccia = @ccia", conx)
        cmd.Parameters.AddWithValue("@numero", Trim(TextBox3.Text))
        cmd.Parameters.AddWithValue("@ccia", Trim(Label8.Text))
        cmd.ExecuteNonQuery()
        ENVIAR_CORREO3()

        TextBox2.Text = ""
        DataGridView1.Rows.Clear()
        Dim jj1 As New frama
        Dim pp As New vrama

        pp.gccia = Label8.Text
        Me.TextBox3.Text = jj1.correlativo_rama(pp) + 1
        Select Case TextBox3.Text.Length
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
        MsgBox("SE CERRO EL PROGRAMA DE RAMA Nº" + Trim(TextBox3.Text))
    End Sub
End Class