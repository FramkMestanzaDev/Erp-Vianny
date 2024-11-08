Imports System.Net.Mail
Imports System.Data.SqlClient
Imports System.Data.Sql
Public Class Nota_Pedido
    Public enunciado3 As SqlCommand
    Public respuesta3 As SqlDataReader
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public enunciado As SqlCommand
    Public respuesta As SqlDataReader
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Public enunciado4 As SqlCommand
    Public respuesta4 As SqlDataReader
    Public conn As SqlDataAdapter
    Dim ty, ty2, ty3 As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub TextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox10.KeyDown
        If Trim(TextBox2.Text) <> "" Then
            If e.KeyCode = Keys.F1 Then
                Productos_Vianny.ShowDialog()
                Productos_Vianny.TextBox2.Text = TextBox18.Text
                Productos_Vianny.TextBox5.Text = TextBox2.Text
                Productos_Vianny.TextBox6.Text = TextBox3.Text
                Productos_Vianny.TextBox7.Text = TextBox8.Text
                Productos_Vianny.TextBox8.Text = TextBox9.Text
                Productos_Vianny.Label4.Text = Label25.Text
                Productos_Vianny.Label5.Text = Label26.Text
            End If
        Else
            MsgBox("FALTA INGRESAR EL CLIENTE")
        End If

    End Sub
    Dim dt3 As DataTable
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Direccion_Despacho.Show()

        Dim r As New vdetdireccion
        Dim r1 As New fdireccion

        r.gcodigo = TextBox11.Text
        dt3 = r1.buscar_direc_despacc(r)
        If dt3.Rows.Count <> 0 Then
            Direccion_Despacho.DataGridView1.DataSource = dt3
            Direccion_Despacho.DataGridView1.Columns(1).Width = 440
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Listar_Clientes.TextBox2.Text = 1
        Listar_Clientes.TextBox3.Text = Trim(TextBox9.Text)
        Listar_Clientes.Label2.Text = Label26.Text
        Listar_Clientes.ShowDialog()
        'Clientes.TextBox3.Text = 19
        'Clientes.Show()
    End Sub
    Private Sub LOAD12()
        CheckBox4.Checked = True
        RadioButton2.Checked = True
        Dim func As New fnopedido
        Dim hi As New vnapedido
        hi.gempresa = Label26.Text
        Me.TextBox1.Text = func.correlativo_npedido(hi) + 1
        Select Case TextBox1.Text.Length
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
        TextBox1.Select()
        If Label28.Text = "ADMINISTRADOR" Or Label28.Text = "JEFATURA COMERCIAL" Then
            ComboBox1.Visible = True
            'ComboBox1.SelectedIndex = 6
        Else
            ComboBox1.Visible = False
        End If
        Button2.Enabled = False
        TextBox20.ReadOnly = True
        ComboBox2.SelectedIndex = 0
        TextBox2.ReadOnly = True
        TextBox13.ReadOnly = True
        ComboBox2.Enabled = False
        TextBox7.ReadOnly = True
        Button3.Enabled = False
        TextBox12.ReadOnly = True
        TextBox14.ReadOnly = True
        DataGridView1.Enabled = False
        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
        CheckBox3.Enabled = False
        TextBox10.Enabled = False

        'Button9.Enabled = True
        Button10.Enabled = False
        Button11.Enabled = False
        Button12.Enabled = False
        RadioButton1.Enabled = False
        RadioButton2.Enabled = False
        CheckBox3.Checked = True
        TextBox2.ReadOnly = True
    End Sub
    Private Sub Nota_Pedido_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LOAD12()
        abrir()
        llenar_combo_box2()
        If Label28.Text = "ADMINISTRADOR" Or Label28.Text = "COBRANZAS" Then

            ComboBox1.SelectedIndex = 0
        End If
        TextBox1.Select()

    End Sub
    Sub llenar_combo_box2()
        Try

            conn = New SqlDataAdapter("SELECT codigo_ven,alias_ven FROM  custom_vianny.dbo.Vendedores WHERE  rpm_ven='2'  and admin_ven in (0,2) ", conx)
            conn.Fill(ty2)
            ComboBox1.DataSource = ty2
            ComboBox1.DisplayMember = "alias_ven"
            ComboBox1.ValueMember = "codigo_ven"
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        Label5.Text = "Ruc :"
        TextBox2.MaxLength = 11
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Label5.Text = "Cod :"
        TextBox2.MaxLength = 4
    End Sub

    Dim dt As New DataTable
    Dim dt1 As New DataTable
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            If RadioButton1.Checked = True Then
                Dim g As New vcliente
                Dim g1 As New fcliente
                g.gcodigo = TextBox2.Text

                dt = g1.buscar_cliente(g)

                If dt.Rows.Count <> 0 Then
                    DataGridView3.DataSource = dt
                    TextBox2.Text = DataGridView3.Rows(0).Cells(3).Value
                    TextBox3.Text = DataGridView3.Rows(0).Cells(2).Value
                    TextBox4.Text = DataGridView3.Rows(0).Cells(4).Value
                    TextBox5.Text = DataGridView3.Rows(0).Cells(6).Value
                    TextBox6.Text = DataGridView3.Rows(0).Cells(5).Value
                    TextBox8.Text = DataGridView3.Rows(0).Cells(8).Value

                    'Select Case TextBox8.Text
                    '    Case "0010" : TextBox9.Text = "GUILIANA BEDON"
                    '    Case "0007" : TextBox9.Text = "VICTO GRAUS"
                    '    Case "0016" : TextBox9.Text = "HECTOR GONZALEZ"
                    '    Case "0017" : TextBox9.Text = "NORMA CERRALTA"
                    '    Case "0022" : TextBox9.Text = "VERONICA INSIO"
                    '    Case "0023" : TextBox9.Text = "DIEGO BRAVO"
                    'End Select
                    TextBox11.Text = DataGridView3.Rows(0).Cells(0).Value
                Else
                    MsgBox("NO EXISTE EL CLIENTE SOLICITADO")
                End If

            End If

            If RadioButton2.Checked = True Then
                Dim g As New vcliente
                Dim g1 As New fcliente
                g.gruc = TextBox2.Text

                dt1 = g1.buscar_cliente_ruc(g)

                If dt1.Rows.Count <> 0 Then
                    DataGridView2.DataSource = dt1
                    TextBox3.Text = DataGridView2.Rows(0).Cells(5).Value
                    TextBox4.Text = DataGridView2.Rows(0).Cells(8).Value
                    TextBox5.Text = DataGridView2.Rows(0).Cells(10).Value
                    TextBox6.Text = DataGridView2.Rows(0).Cells(11).Value


                    'Select Case TextBox8.Text
                    '    Case "0010" : TextBox9.Text = "GUILIANA BEDON"
                    '    Case "0007" : TextBox9.Text = "VICTO GRAUS"
                    '    Case "0016" : TextBox9.Text = "HECTOR GONZALEZ"
                    '    Case "0017" : TextBox9.Text = "NORMA CERRALTA"
                    '    Case "0022" : TextBox9.Text = "VERONICA INSIO"
                    '    Case "0023" : TextBox9.Text = "DIEGO BRAVO"
                    'End Select

                    TextBox11.Text = DataGridView2.Rows(0).Cells(0).Value
                Else
                    MsgBox("NO EXISTE EL CLIENTE SOLICITADO")
                End If
            End If
        End If
    End Sub
    Dim dt10 As New DataTable
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
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Sub ENVIAR_CORREO2()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim fk As New fnopedido
        Dim jj As New vnapedido
        Dim corre As String
        jj.gvendedor = TextBox8.Text
        corre = fk.buscar_correo(jj)


        Dim Mailz As String = "" &
      " <!DOCTYPE html>" &
"<html xmlns='http://www.w3.org/1999/xhtml'>" &
"<head>" &
"    <title></title>" &
"</head>" &
"<body>
     <FONT COLOR='green'>INFORMACION DE LA NOTA DE PEDIDO</FONT><br/><br/>

<FONT COLOR='blue'>* ESTADO : </FONT> ACTUALIZACION <br/>
<FONT COLOR='blue'>* NOTA PEDIDO : </FONT>" + Trim(TextBox1.Text) + " <br/>
<FONT COLOR='blue'>* VENDEDOR : </FONT>" + Trim(TextBox9.Text) + " <br/>
<FONT COLOR='blue'>* EMPRESA : </FONT>" + Trim(Label27.Text) + " <br/>
<FONT COLOR='blue'>* CLIENTE : </FONT>" + Trim(TextBox3.Text) + "<br/>
<FONT COLOR='blue'>* FORMA DE PAGO : </FONT>" + Trim(TextBox20.Text) + "<br/>
<FONT COLOR='blue'>* TOTAL VENTA : </FONT>" + Trim(TextBox17.Text) + "<br/>

</body>

</html>"
        With message
            .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
            .To.Add("coordinadorit@viannysac.com,hrivera@viannysac.com" & corre)
            .IsBodyHtml = True
            .Body = Mailz
            .Subject = "Nota Pedido N°" & TextBox1.Text
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

        End Try
    End Sub
    Sub ENVIAR_CORREO3()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim fk As New fnopedido
        Dim jj As New vnapedido
        Dim corre As String
        jj.gvendedor = TextBox8.Text
        corre = fk.buscar_correo(jj)

        Dim Mailz As String = "" &
      " <!DOCTYPE html>" &
"<html xmlns='http://www.w3.org/1999/xhtml'>" &
"<head>" &
"    <title></title>" &
"</head>" &
"<body>
     <FONT COLOR='green'>INFORMACION DE LA NOTA DE PEDIDO</FONT><br/><br/>

<FONT COLOR='blue'>* ESTADO : </FONT> ANULADO <br/>
<FONT COLOR='blue'>* NOTA PEDIDO : </FONT>" + Trim(TextBox1.Text) + " <br/>
<FONT COLOR='blue'>* VENDEDOR : </FONT>" + Trim(TextBox9.Text) + " <br/>
<FONT COLOR='blue'>* EMPRESA : </FONT>" + Trim(Label27.Text) + " <br/>
<FONT COLOR='blue'>* CLIENTE : </FONT>" + Trim(TextBox3.Text) + "<br/>
<FONT COLOR='blue'>* FORMA DE PAGO : </FONT>" + Trim(TextBox20.Text) + "<br/>
<FONT COLOR='blue'>* TOTAL VENTA : </FONT>" + Trim(TextBox17.Text) + "<br/>

</body>

</html>"
        With message
            .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
            .To.Add("fmestanza@viannysac.com,dbravo@viannysac.com," & corre)
            .IsBodyHtml = True
            .Body = Mailz
            .Subject = "Nota Pedido N°" & TextBox1.Text
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

        End Try
    End Sub
    Sub LIMPIAR()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox13.Text = ""
        TextBox18.Text = ""

        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox11.Text = ""
        TextBox20.Text = ""
        TextBox21.Text = ""
        TextBox14.Text = ""
        TextBox12.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
        DataGridView3.Rows.Clear()
        'DataGridView4.Rows.Clear()
        'DataGridView5.Rows.Clear()
        'DataGridView6.Rows.Clear()
    End Sub
    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit, DataGridView1.CellContentClick
        Dim i, a9, fila, d1 As Integer
        Dim cant9, sum9, cant8, sum8, cant7, sum7, canti As Double
        i = DataGridView1.Rows.Count
        d1 = DataGridView7.Rows.Count
        For i2 = 0 To d1 - 1
            canti = canti + DataGridView7.Rows(i2).Cells(1).Value
        Next
        fila = e.RowIndex
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Cant." Then
            If Convert.ToDouble(DataGridView1.Rows(fila).Cells(8).Value) <> canti And Trim(DataGridView1.Rows(fila).Cells(5).Value).Length > 0 Then
                MsgBox("LA CANTIDAD AGERGADA DE : " + Convert.ToString(canti) + "  DEL LOS ROLLOS SELCCIONADOS, NO COINCIDE CON CON EL AGREGADO EN LA COLUMNA CANT.. SE ACTUALIZARA A LO CORRECTO")
                DataGridView1.Rows(fila).Cells(8).Value = canti
            Else

                If Trim(TextBox19.Text = "1") Then
                    If DataGridView1.Rows(fila).Cells(8).Value <= DataGridView1.Rows(fila).Cells(13).Value Then

                        If CheckBox3.Checked = True And CheckBox4.Checked = True Then
                            DataGridView1.Rows(fila).Cells(12).Value = DataGridView1.Rows(fila).Cells(8).Value * DataGridView1.Rows(fila).Cells(9).Value
                            DataGridView1.Rows(fila).Cells(10).Value = DataGridView1.Rows(fila).Cells(12).Value / 1.18
                            DataGridView1.Rows(fila).Cells(11).Value = DataGridView1.Rows(fila).Cells(12).Value - DataGridView1.Rows(fila).Cells(10).Value

                            Me.DataGridView1.Columns(10).DefaultCellStyle.Format = "n2"
                            Me.DataGridView1.Columns(11).DefaultCellStyle.Format = "n2"
                            Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                        Else
                            If CheckBox3.Checked = False And CheckBox4.Checked = True Then

                                DataGridView1.Rows(fila).Cells(10).Value = DataGridView1.Rows(fila).Cells(8).Value * DataGridView1.Rows(fila).Cells(9).Value
                                DataGridView1.Rows(fila).Cells(11).Value = DataGridView1.Rows(fila).Cells(10).Value * 0.18
                                DataGridView1.Rows(fila).Cells(12).Value = DataGridView1.Rows(fila).Cells(10).Value + DataGridView1.Rows(fila).Cells(11).Value
                                Me.DataGridView1.Columns(10).DefaultCellStyle.Format = "n2"
                                Me.DataGridView1.Columns(11).DefaultCellStyle.Format = "n2"
                                Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                            Else
                                If CheckBox3.Checked = False And CheckBox4.Checked = False Then
                                    DataGridView1.Rows(fila).Cells(10).Value = DataGridView1.Rows(fila).Cells(8).Value * DataGridView1.Rows(fila).Cells(9).Value
                                    DataGridView1.Rows(fila).Cells(11).Value = 0
                                    DataGridView1.Rows(fila).Cells(12).Value = DataGridView1.Rows(fila).Cells(8).Value * DataGridView1.Rows(fila).Cells(9).Value
                                    Me.DataGridView1.Columns(10).DefaultCellStyle.Format = "n2"
                                    Me.DataGridView1.Columns(11).DefaultCellStyle.Format = "n2"
                                    Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                                End If

                            End If

                        End If
                    Else
                        MsgBox("LA CANTIDAD SOLICITADA ES MAYOR AL STOCK")
                        DataGridView1.Rows(fila).Cells(8).Value = 0

                    End If
                Else
                    If CheckBox3.Checked = True And CheckBox4.Checked = True Then
                        DataGridView1.Rows(fila).Cells(12).Value = DataGridView1.Rows(fila).Cells(8).Value * DataGridView1.Rows(fila).Cells(9).Value
                        DataGridView1.Rows(fila).Cells(10).Value = DataGridView1.Rows(fila).Cells(12).Value / 1.18
                        DataGridView1.Rows(fila).Cells(11).Value = DataGridView1.Rows(fila).Cells(12).Value - DataGridView1.Rows(fila).Cells(10).Value

                        Me.DataGridView1.Columns(10).DefaultCellStyle.Format = "n2"
                        Me.DataGridView1.Columns(11).DefaultCellStyle.Format = "n2"
                        Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                    Else
                        If CheckBox3.Checked = False And CheckBox4.Checked = True Then

                            DataGridView1.Rows(fila).Cells(10).Value = DataGridView1.Rows(fila).Cells(8).Value * DataGridView1.Rows(fila).Cells(9).Value
                            DataGridView1.Rows(fila).Cells(11).Value = DataGridView1.Rows(fila).Cells(10).Value * 0.18
                            DataGridView1.Rows(fila).Cells(12).Value = DataGridView1.Rows(fila).Cells(10).Value + DataGridView1.Rows(fila).Cells(11).Value
                            Me.DataGridView1.Columns(10).DefaultCellStyle.Format = "n2"
                            Me.DataGridView1.Columns(11).DefaultCellStyle.Format = "n2"
                            Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                        Else
                            If CheckBox3.Checked = False And CheckBox4.Checked = False Then
                                DataGridView1.Rows(fila).Cells(10).Value = DataGridView1.Rows(fila).Cells(8).Value * DataGridView1.Rows(fila).Cells(9).Value
                                DataGridView1.Rows(fila).Cells(11).Value = 0
                                DataGridView1.Rows(fila).Cells(12).Value = DataGridView1.Rows(fila).Cells(8).Value * DataGridView1.Rows(fila).Cells(9).Value
                                Me.DataGridView1.Columns(10).DefaultCellStyle.Format = "n2"
                                Me.DataGridView1.Columns(11).DefaultCellStyle.Format = "n2"
                                Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                            End If

                        End If

                    End If
                End If


            End If

        Else
            If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Precio Uni." Then

                If CheckBox3.Checked = True And CheckBox4.Checked = True Then
                    DataGridView1.Rows(fila).Cells(12).Value = DataGridView1.Rows(fila).Cells(8).Value * DataGridView1.Rows(fila).Cells(9).Value
                    DataGridView1.Rows(fila).Cells(10).Value = DataGridView1.Rows(fila).Cells(12).Value / 1.18
                    DataGridView1.Rows(fila).Cells(11).Value = DataGridView1.Rows(fila).Cells(12).Value - DataGridView1.Rows(fila).Cells(10).Value

                    Me.DataGridView1.Columns(10).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(11).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                Else
                    If CheckBox3.Checked = False And CheckBox4.Checked = True Then

                        DataGridView1.Rows(fila).Cells(10).Value = DataGridView1.Rows(fila).Cells(8).Value * DataGridView1.Rows(fila).Cells(9).Value
                        DataGridView1.Rows(fila).Cells(11).Value = DataGridView1.Rows(fila).Cells(10).Value * 0.18
                        DataGridView1.Rows(fila).Cells(12).Value = DataGridView1.Rows(fila).Cells(10).Value + DataGridView1.Rows(fila).Cells(11).Value
                        Me.DataGridView1.Columns(10).DefaultCellStyle.Format = "n2"
                        Me.DataGridView1.Columns(11).DefaultCellStyle.Format = "n2"
                        Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                    Else
                        If CheckBox3.Checked = False And CheckBox4.Checked = False Then
                            DataGridView1.Rows(fila).Cells(10).Value = DataGridView1.Rows(fila).Cells(8).Value * DataGridView1.Rows(fila).Cells(9).Value
                            DataGridView1.Rows(fila).Cells(11).Value = 0
                            DataGridView1.Rows(fila).Cells(12).Value = DataGridView1.Rows(fila).Cells(8).Value * DataGridView1.Rows(fila).Cells(9).Value
                            Me.DataGridView1.Columns(10).DefaultCellStyle.Format = "n2"
                            Me.DataGridView1.Columns(11).DefaultCellStyle.Format = "n2"
                            Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                        End If

                    End If

                End If

            End If
        End If
        For a9 = 0 To i - 1
            cant9 = Val(DataGridView1.Rows(a9).Cells(10).Value)
            sum9 = cant9 + Val(sum9)
            cant8 = Val(DataGridView1.Rows(a9).Cells(11).Value)
            sum8 = cant8 + Val(sum8)
            cant7 = Val(DataGridView1.Rows(a9).Cells(12).Value)
            sum7 = cant7 + Val(sum7)

        Next

        TextBox15.Text = sum9.ToString("N2")
        TextBox16.Text = sum8.ToString("N2")
        TextBox17.Text = sum7.ToString("N2")




    End Sub
    Dim dr, dr1, dr2 As New DataTable

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub CheckBox3_Click(sender As Object, e As EventArgs) Handles CheckBox3.Click
        Dim i, fila As Integer
        Dim a9 As Integer
        Dim cant9, sum9, cant8, sum8, cant7, sum7 As Double

        fila = DataGridView1.Rows.Count

        For i = 0 To fila - 1
            If CheckBox3.Checked = True Then
                DataGridView1.Rows(i).Cells(12).Value = DataGridView1.Rows(i).Cells(8).Value * DataGridView1.Rows(i).Cells(9).Value
                DataGridView1.Rows(i).Cells(10).Value = DataGridView1.Rows(i).Cells(12).Value / 1.18
                DataGridView1.Rows(i).Cells(11).Value = DataGridView1.Rows(i).Cells(12).Value - DataGridView1.Rows(i).Cells(10).Value

                Me.DataGridView1.Columns(10).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(11).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"

                cant9 = Val(DataGridView1.Rows(i).Cells(10).Value)
                sum9 = cant9 + Val(sum9)
                cant8 = Val(DataGridView1.Rows(i).Cells(11).Value)
                sum8 = cant8 + Val(sum8)
                cant7 = Val(DataGridView1.Rows(i).Cells(12).Value)
                sum7 = cant7 + Val(sum7)


            Else
                DataGridView1.Rows(i).Cells(10).Value = DataGridView1.Rows(i).Cells(8).Value * DataGridView1.Rows(i).Cells(9).Value
                DataGridView1.Rows(i).Cells(11).Value = DataGridView1.Rows(i).Cells(10).Value * 0.18
                DataGridView1.Rows(i).Cells(12).Value = DataGridView1.Rows(i).Cells(10).Value + DataGridView1.Rows(i).Cells(11).Value
                Me.DataGridView1.Columns(10).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(11).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"

                cant9 = Val(DataGridView1.Rows(i).Cells(10).Value)
                sum9 = cant9 + Val(sum9)
                cant8 = Val(DataGridView1.Rows(i).Cells(11).Value)
                sum8 = cant8 + Val(sum8)
                cant7 = Val(DataGridView1.Rows(i).Cells(12).Value)
                sum7 = cant7 + Val(sum7)


            End If
        Next

        TextBox15.Text = sum9.ToString("N2")
            TextBox16.Text = sum8.ToString("N2")
            TextBox17.Text = sum7.ToString("N2")


    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged




    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Eliminar.Show()
        Dim respuesta As DialogResult
        Dim I18, VAL As Integer
        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)


        End If

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form_Cliente.Label1.Text = "F. PAGO"
        Form_Cliente.TextBox3.Text = 3

        Form_Cliente.ShowDialog()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Form2.TextBox3.Text = 3
        Form2.Show()
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged

    End Sub



    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        DataGridView1.Rows.Add(1)
        Dim ui As Integer
        ui = DataGridView1.Rows.Count

        If ui = 1 Then
            DataGridView1.Rows(0).Cells(0).Value = 1
            DataGridView1.Rows(0).Cells(1).Value = "60"
            DataGridView1.Rows(0).Cells(7).Value = "KGS"
            DataGridView1.Rows(0).Cells(3).ReadOnly = False
            DataGridView1.Rows(0).Cells(13).Value = "999"
            DataGridView1.Rows(0).Cells(2).Value = "SERVICIO"
        Else
            DataGridView1.Rows(ui - 1).Cells(0).Value = DataGridView1.Rows(ui - 2).Cells(0).Value + 1
            DataGridView1.Rows(ui - 1).Cells(1).Value = "60"
            DataGridView1.Rows(ui - 1).Cells(7).Value = "KGS"
            DataGridView1.Rows(ui - 1).Cells(13).Value = "999"
            DataGridView1.Rows(ui - 1).Cells(3).ReadOnly = False
            DataGridView1.Rows(ui - 1).Cells(2).Value = "SERVICIO"
        End If

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        DataGridView1.Rows.Add(1)
        Dim ui As Integer
        ui = DataGridView1.Rows.Count

        If ui = 1 Then
            DataGridView1.Rows(0).Cells(0).Value = 1
            DataGridView1.Rows(0).Cells(1).Value = "PRIMERA"
            DataGridView1.Rows(0).Cells(7).Value = "KGS"
            DataGridView1.Rows(0).Cells(3).ReadOnly = False
            DataGridView1.Rows(0).Cells(13).Value = "9999"
            DataGridView1.Rows(0).Cells(2).Value = "X PRODUCIR"
            DataGridView1.Rows(0).Cells(14).Value = "68"
        Else
            DataGridView1.Rows(ui - 1).Cells(0).Value = DataGridView1.Rows(ui - 2).Cells(0).Value + 1
            DataGridView1.Rows(ui - 1).Cells(1).Value = "PRIMERA"
            DataGridView1.Rows(ui - 1).Cells(7).Value = "KGS"
            DataGridView1.Rows(ui - 1).Cells(13).Value = "9999"
            DataGridView1.Rows(ui - 1).Cells(3).ReadOnly = False
            DataGridView1.Rows(ui - 1).Cells(2).Value = "X PRODUCIR"
            DataGridView1.Rows(ui - 1).Cells(1).Value = "10"
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub
    Dim Rsr2 As SqlDataReader
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If Label26.Text = "01" Then
            DataGridView1.Rows.Add(1)
            Dim ui As Integer
            ui = DataGridView1.Rows.Count

            If ui = 1 Then
                DataGridView1.Rows(0).Cells(0).Value = 1
                DataGridView1.Rows(0).Cells(1).Value = "PRIMERA"
                DataGridView1.Rows(0).Cells(7).Value = "KGS"
                DataGridView1.Rows(0).Cells(3).ReadOnly = False
                DataGridView1.Rows(0).Cells(13).Value = "999"
                DataGridView1.Rows(0).Cells(2).Value = "REG PEDIDO"
                DataGridView1.Rows(0).Cells(4).ReadOnly = False
                DataGridView1.Rows(0).Cells(14).Value = "68"
            Else
                DataGridView1.Rows(ui - 1).Cells(0).Value = DataGridView1.Rows(ui - 2).Cells(0).Value + 1
                DataGridView1.Rows(ui - 1).Cells(1).Value = "PRIMERA"
                DataGridView1.Rows(ui - 1).Cells(7).Value = "KGS"
                DataGridView1.Rows(ui - 1).Cells(13).Value = "999"
                DataGridView1.Rows(ui - 1).Cells(3).ReadOnly = False
                DataGridView1.Rows(ui - 1).Cells(2).Value = "REG PEDIDO"
                DataGridView1.Rows(ui - 1).Cells(4).ReadOnly = False
                DataGridView1.Rows(ui - 1).Cells(14).Value = "68"
            End If
        Else
            DataGridView1.Rows.Add(1)
            Dim ui As Integer
            ui = DataGridView1.Rows.Count

            If ui = 1 Then
                DataGridView1.Rows(0).Cells(0).Value = 1
                DataGridView1.Rows(0).Cells(1).Value = "PRIMERA"
                DataGridView1.Rows(0).Cells(7).Value = "KGS"
                DataGridView1.Rows(0).Cells(3).ReadOnly = False
                DataGridView1.Rows(0).Cells(13).Value = "999"
                DataGridView1.Rows(0).Cells(2).Value = "REG PEDIDO"
                DataGridView1.Rows(0).Cells(4).ReadOnly = False
                DataGridView1.Rows(0).Cells(14).Value = "67"
            Else
                DataGridView1.Rows(ui - 1).Cells(0).Value = DataGridView1.Rows(ui - 2).Cells(0).Value + 1
                DataGridView1.Rows(ui - 1).Cells(1).Value = "PRIMERA"
                DataGridView1.Rows(ui - 1).Cells(7).Value = "KGS"
                DataGridView1.Rows(ui - 1).Cells(13).Value = "999"
                DataGridView1.Rows(ui - 1).Cells(3).ReadOnly = False
                DataGridView1.Rows(ui - 1).Cells(2).Value = "REG PEDIDO"
                DataGridView1.Rows(ui - 1).Cells(4).ReadOnly = False
                DataGridView1.Rows(ui - 1).Cells(14).Value = "67"
            End If
        End If


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        abrir()
        Dim sql102 As String = "SELECT codigo_ven,alias_ven  FROM  custom_vianny.dbo.Vendedores WHERE alias_ven ='" + ComboBox1.Text + "'  AND ccia_ven ='" + Label26.Text + "'"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        If Rsr2.Read() = True Then
            TextBox8.Text = Rsr2(0)
            TextBox9.Text = Rsr2(1)
            TextBox22.Text = Rsr2(0)
        End If

        Rsr2.Close()

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            DataGridView1.Rows.Clear()
            Button2.Enabled = True
            Dim by As New vnapedido
            Dim by1 As New fnopedido
            Dim bg As String
            bg = TextBox1.Text
            ComboBox1.Enabled = True
            Select Case bg.Length
                Case "1" : TextBox1.Text = "000000000" & "" & bg
                Case "2" : TextBox1.Text = "00000000" & "" & bg
                Case "3" : TextBox1.Text = "0000000" & "" & bg
                Case "4" : TextBox1.Text = "000000" & "" & bg
                Case "5" : TextBox1.Text = "00000" & "" & bg
                Case "6" : TextBox1.Text = "0000" & "" & bg
                Case "7" : TextBox1.Text = "000" & "" & bg
                Case "8" : TextBox1.Text = "00" & "" & bg
                Case "9" : TextBox1.Text = "0" & "" & bg
                Case "10" : TextBox1.Text = bg
            End Select
            by.gnumero_pedido = TextBox1.Text
            by.gempresa = Label26.Text
            dr = by1.buscar_pedido_cabecera(by)
            dr1 = by1.buscar_pedido_detalle(by)

            If dr.Rows.Count <> 0 Then
                DataGridView4.DataSource = dr

                Button12.Enabled = True
                Button10.Enabled = True
                Button9.Enabled = False
                If Trim(DataGridView4.Rows(0).Cells(8).Value) = Trim(TextBox8.Text) Or Trim(Label28.Text) = "ADMINISTRADOR" Or Trim(Label28.Text) = "JEFATURA COMERCIAL" Then
                    Label30.Text = "2"
                    If DataGridView4.Rows(0).Cells(3).Value = "S" Then
                        ComboBox2.SelectedIndex = 0
                    Else
                        ComboBox2.SelectedIndex = 1
                    End If
                    TextBox22.Text = Trim(DataGridView4.Rows(0).Cells(8).Value)
                    TextBox13.Text = DataGridView4.Rows(0).Cells(4).Value
                    If DataGridView4.Rows(0).Cells(5).Value = 1 Then
                        CheckBox2.Checked = True
                    Else
                        CheckBox2.Checked = False
                    End If
                    TextBox7.Text = DataGridView4.Rows(0).Cells(6).Value
                    DateTimePicker1.Value = DataGridView4.Rows(0).Cells(1).Value
                    TextBox12.Text = DataGridView4.Rows(0).Cells(9).Value
                    DateTimePicker2.Value = DataGridView4.Rows(0).Cells(10).Value
                    If DataGridView4.Rows(0).Cells(11).Value = 1 Then
                        CheckBox1.Checked = True
                    Else
                        CheckBox1.Checked = False
                    End If
                    TextBox14.Text = DataGridView4.Rows(0).Cells(12).Value
                    TextBox15.Text = Convert.ToDouble(DataGridView4.Rows(0).Cells(13).Value).ToString("N2")
                    TextBox16.Text = Convert.ToDouble(DataGridView4.Rows(0).Cells(14).Value).ToString("N2")
                    TextBox17.Text = Convert.ToDouble(DataGridView4.Rows(0).Cells(15).Value).ToString("N2")
                    TextBox11.Text = DataGridView4.Rows(0).Cells(2).Value
                    TextBox21.Text = DataGridView4.Rows(0).Cells(7).Value
                    If Trim(DataGridView4.Rows(0).Cells(17).Value) = "1" Then
                        CheckBox3.Checked = True
                    Else
                        CheckBox3.Checked = False
                    End If
                    If Trim(DataGridView4.Rows(0).Cells(22).Value) = "1" Then
                        CheckBox4.Checked = True
                    Else
                        CheckBox4.Checked = False
                    End If
                    abrir()
                    enunciado4 = New SqlCommand("SELECT alias_ven FROM  custom_vianny.dbo.Vendedores WHERE  rpm_ven='2'  and admin_ven ='0' AND codigo_ven ='" + Trim(TextBox22.Text) + "' ", conx)
                    respuesta4 = enunciado4.ExecuteReader
                    While respuesta4.Read
                        ComboBox1.Text = respuesta4.Item("alias_ven")
                    End While
                    respuesta4.Close()
                    If Convert.ToString(Trim(DataGridView4.Rows(0).Cells(7).Value)).Length = 0 Then
                        TextBox20.Text = ""
                        TextBox21.Text = ""

                    Else

                        enunciado3 = New SqlCommand("SELECT  nomb_15k AS DESCRIPCION FROM custom_vianny.dbo.kag1500 A  Where A.ccia_15k='01' and flag_15k='1' and cond_15k=" + TextBox21.Text, conx)
                        respuesta3 = enunciado3.ExecuteReader
                        While respuesta3.Read
                            TextBox20.Text = respuesta3.Item("DESCRIPCION")
                        End While
                        respuesta3.Close()
                    End If
                    If DataGridView4.Rows(0).Cells(18).Value = 1 Then
                        Label19.Text = "APROBADO"
                    Else
                        If DataGridView4.Rows(0).Cells(18).Value = 0 Then
                            Label19.Text = "DESAPROBADO"
                        End If
                    End If

                    Dim uy As New fcliente
                    Dim uy1 As New vcliente

                    uy1.gcodigo = DataGridView4.Rows(0).Cells(2).Value

                    dr2 = uy.buscar_cliente(uy1)

                    If dr2.Rows.Count <> 0 Then
                        DataGridView6.DataSource = dr2
                        TextBox2.Text = DataGridView6.Rows(0).Cells(19).Value
                        TextBox3.Text = DataGridView6.Rows(0).Cells(25).Value
                        TextBox4.Text = DataGridView6.Rows(0).Cells(8).Value
                        TextBox6.Text = DataGridView6.Rows(0).Cells(9).Value
                        TextBox5.Text = DataGridView6.Rows(0).Cells(10).Value


                    End If

                    If dr1.Rows.Count <> 0 Then
                        DataGridView5.DataSource = dr1
                        Dim ip As Integer
                        ip = DataGridView5.Rows.Count

                        DataGridView1.Rows.Add(ip)
                        For o = 0 To ip - 1
                            DataGridView1.Rows(o).Cells(0).Value = DataGridView5.Rows(o).Cells(1).Value
                            DataGridView1.Rows(o).Cells(1).Value = DataGridView5.Rows(o).Cells(9).Value
                            DataGridView1.Rows(o).Cells(2).Value = DataGridView5.Rows(o).Cells(2).Value
                            DataGridView1.Rows(o).Cells(3).Value = DataGridView5.Rows(o).Cells(3).Value
                            DataGridView1.Rows(o).Cells(4).Value = DataGridView5.Rows(o).Cells(8).Value
                            DataGridView1.Rows(o).Cells(5).Value = DataGridView5.Rows(o).Cells(20).Value
                            DataGridView1.Rows(o).Cells(6).Value = DataGridView5.Rows(o).Cells(6).Value
                            DataGridView1.Rows(o).Cells(7).Value = DataGridView5.Rows(o).Cells(7).Value
                            DataGridView1.Rows(o).Cells(8).Value = DataGridView5.Rows(o).Cells(10).Value
                            DataGridView1.Rows(o).Cells(9).Value = DataGridView5.Rows(o).Cells(11).Value
                            DataGridView1.Rows(o).Cells(10).Value = DataGridView5.Rows(o).Cells(15).Value
                            DataGridView1.Rows(o).Cells(11).Value = DataGridView5.Rows(o).Cells(12).Value
                            DataGridView1.Rows(o).Cells(12).Value = DataGridView5.Rows(o).Cells(13).Value
                            DataGridView1.Rows(o).Cells(13).Value = DataGridView5.Rows(o).Cells(10).Value
                            DataGridView1.Rows(o).Cells(14).Value = DataGridView5.Rows(o).Cells(19).Value
                        Next


                    End If
                Else
                    TextBox2.Text = ""
                    TextBox3.Text = ""
                    TextBox4.Text = ""
                    TextBox5.Text = ""
                    TextBox6.Text = ""
                    TextBox13.Text = ""
                    TextBox15.Text = ""
                    TextBox16.Text = ""
                    TextBox17.Text = ""
                    MsgBox("ESTA NOTA DE PEDIDO PERTENECE A OTRO VENDEDOR")
                End If


            End If



            If dr1.Rows.Count = 0 And dr2.Rows.Count = 0 Then
                TextBox2.ReadOnly = False
                TextBox13.ReadOnly = False
                ComboBox2.Enabled = True
                TextBox7.ReadOnly = False
                Button3.Enabled = True
                TextBox12.ReadOnly = False
                TextBox14.ReadOnly = False
                DataGridView1.Enabled = True

                CheckBox1.Enabled = True
                CheckBox2.Enabled = True
                CheckBox3.Enabled = True
                TextBox10.Enabled = True
                TextBox20.Enabled = True
                RadioButton1.Enabled = True
                RadioButton2.Enabled = True
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox5.Text = ""
                TextBox13.Text = ""
                Button9.Enabled = True
            Else
                If DataGridView4.Rows(0).Cells(18).Value = 1 Then
                    Label19.Text = "APROBADO"
                    MsgBox("LA NOTA DE PEDIDO ESTA APROBADO NO PUEDE REALIZAR CAMBIOS")

                    Button9.Enabled = False
                    Button10.Enabled = False
                    Button11.Enabled = False
                    Button12.Enabled = True
                End If
                If DataGridView4.Rows(0).Cells(16).Value = 1 Then
                    Label24.Text = "ANULADO"
                    Label24.Visible = True
                    Button9.Enabled = False
                    Button10.Enabled = False
                    Button11.Enabled = False
                    Button12.Enabled = True
                Else

                    Label24.Visible = False
                    Button11.Enabled = True
                    Button10.Enabled = True
                End If


            End If

            TextBox2.Select()
        End If

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim df As New fnopedido
        Dim df1 As New vnapedido
        Dim df2 As New vnapedido
        Dim df3 As New vnapedido
        Dim pedido As String

        If Trim(TextBox22.Text) <> Trim(TextBox8.Text) Then
            MsgBox("EL CLIENTE NO PERTENECE A EL VENDEDOR AGREGADO")
        Else

            pedido = ""

            Dim op As Integer
            Dim sumw, te17, REDO As Double
            sumw = 0
            te17 = 0
            op = DataGridView1.Rows.Count

            For l = 0 To op - 1
                sumw = sumw + Convert.ToDouble(DataGridView1.Rows(l).Cells(12).Value)
            Next
            te17 = Convert.ToDouble(TextBox17.Text)
            REDO = Convert.ToString(Math.Round(sumw, 3))
            'MsgBox(Format(REDO, "#.#0"))
            'MsgBox(Convert.ToString(Math.Round(sumw, 3)))
            'MsgBox(Convert.ToString(te17))
            'MsgBox(Format(te17, "#.#0"))
            If Format(REDO, "#.#0") = Format(te17, "#.#0") Then
                If String.IsNullOrEmpty(TextBox20.Text) Or String.IsNullOrEmpty(TextBox3.Text) Or String.IsNullOrEmpty(TextBox2.Text) Then
                    MsgBox("FALTA INGRESAR EL Nº DE RUC, LA FORMA DE PAGO O LA RAZON SOCIAL")

                Else
                    df3.gnumero_pedido = TextBox1.Text
                    df3.gempresa = Label26.Text
                    If df.validar_pedido(df3) = True Then
                        If TextBox19.Text = 2 Then
                            df2.gnumero_pedido = TextBox1.Text
                            df2.gempresa = Label26.Text
                            df.eliminar_pedido(df2)
                            pedido = TextBox1.Text

                        Else
                            If TextBox19.Text = 1 Then

                                Dim func As New fnopedido
                                Dim hi As New vnapedido
                                hi.gempresa = Label26.Text
                                pedido = func.correlativo_npedido(hi) + 1
                                'pedido = TextBox1.Text + 1
                                Select Case pedido.Length
                                    Case "1" : pedido = "000000000" & "" & pedido
                                    Case "2" : pedido = "00000000" & "" & pedido
                                    Case "3" : pedido = "0000000" & "" & pedido
                                    Case "4" : pedido = "000000" & "" & pedido
                                    Case "5" : pedido = "00000" & "" & pedido
                                    Case "6" : pedido = "0000" & "" & pedido
                                    Case "7" : pedido = "000" & "" & pedido
                                    Case "8" : pedido = "00" & "" & pedido
                                    Case "9" : pedido = "0" & "" & pedido
                                    Case "10" : pedido = pedido
                                End Select
                                MsgBox("El Numero de Pedido Ingresado ya existe, su nuevo pedido es----" + pedido)

                            End If

                        End If
                    Else
                        pedido = TextBox1.Text
                    End If

                    df3.gnumero_pedido = pedido

                    df1.gnumero_pedido = pedido

                    df1.gfecha = DateTimePicker1.Text
                    df1.gcodigo_cli = TextBox2.Text
                    df1.gmoneda = ComboBox2.Text
                    df1.gresponsable = TextBox13.Text
                    If CheckBox2.Checked = True Then
                        df1.gvalor_comercial = 1 ' si es con valor omercial
                    Else
                        df1.gvalor_comercial = 2
                    End If

                    df1.gdir_despacho = TextBox7.Text
                    df1.gforma_pago = TextBox21.Text
                    df1.gvendedor = TextBox8.Text
                    df1.gor_compra = TextBox12.Text
                    df1.gf_enrtrega = DateTimePicker2.Text
                    If CheckBox1.Checked = True Then
                        df1.ginmediato = 1 ' si es inmediato
                    Else
                        df1.ginmediato = 2
                    End If

                    df1.gobservaciones = TextBox14.Text
                    df1.gsub_total = TextBox15.Text
                    df1.gigv = TextBox16.Text
                    df1.gtotal = TextBox17.Text
                    df1.gestado = 0 'estado sin aprobar
                    If CheckBox3.Checked = True Then
                        df1.gpreinc = 1 ' con factura
                    Else
                        df1.gpreinc = 0 ' sin factura
                    End If

                    df1.gescom = 0
                    df1.gesalm = 0
                    df1.gesttrans = 0
                    df1.gcotizacion = ""
                    df1.gempresa = Label26.Text
                    If CheckBox4.Checked = True Then
                        df1.gfactura = 1 ' con factura
                    Else
                        df1.gfactura = 0 ' sin factura
                    End If

                    df.insertar_nota_pedido(df1)
                    Dim g As New Integer
                    g = DataGridView1.RowCount
                    Dim agregar As String = "delete from pedido_separado where numero_pedido ='" + Trim(TextBox1.Text) + "' and EMPRESA ='" + Trim(Label26.Text) + "'"
                    ELIMINAR1(agregar)
                    For i = 0 To g - 1
                        Dim hj As New cdetalle_pedido
                        Dim hj1 As New fnopedido
                        hj.gitems = DataGridView1.Rows(i).Cells(0).Value


                        hj.gcodigo_tela = DataGridView1.Rows(i).Cells(2).Value



                        hj.gdescripcion = DataGridView1.Rows(i).Cells(3).Value
                        hj.gcolor = ""
                        'If Convert.ToString(DataGridView1.Rows(i).Cells(5).Value) = "" Then
                        hj.gdensidad = ""
                        'Else
                        '    hj.gdensidad = DataGridView1.Rows(i).Cells(5).Value
                        'End If
                        If Convert.ToString(DataGridView1.Rows(i).Cells(6).Value) = "" Then
                            hj.gancho = ""
                        Else
                            hj.gancho = DataGridView1.Rows(i).Cells(6).Value
                        End If

                        hj.gum = DataGridView1.Rows(i).Cells(7).Value
                        If Convert.ToString(DataGridView1.Rows(i).Cells(4).Value) = "" Then
                            hj.gpartida = ""
                        Else
                            hj.gpartida = DataGridView1.Rows(i).Cells(4).Value
                        End If
                        hj.galmacen = Trim(DataGridView1.Rows(i).Cells(1).Value)

                        hj.gcantidad = DataGridView1.Rows(i).Cells(8).Value
                        hj.gprecio_unitario = DataGridView1.Rows(i).Cells(9).Value
                        hj.gigv = DataGridView1.Rows(i).Cells(11).Value
                        hj.gTotal = DataGridView1.Rows(i).Cells(12).Value
                        hj.gsubtotal = DataGridView1.Rows(i).Cells(10).Value
                        hj.gnumero_pedido = pedido
                        hj.gestado = 0
                        hj.gestado2 = 0
                        hj.gempresa = Label26.Text
                        hj.galm_num = Trim(DataGridView1.Rows(i).Cells(14).Value)
                        hj.grollos = Trim(DataGridView1.Rows(i).Cells(5).Value)
                        hj1.insertar_detalle_pedido(hj)


                        'If DataGridView1.Rows(i).Cells(4).Value <> "" Then
                        '    abrir()
                        '    Dim cmd200 As New SqlCommand("insert into pedido_separado(part_lote,estado,codigo,cantidad,numero_pedido,EMPRESA,ALMACEN) values (@part_lote,@estado,@codigo,@cantidad,@numero_pedido,@EMPRESA,@ALMACEN)", conx)
                        '    cmd200.Parameters.AddWithValue("@part_lote", Trim(DataGridView1.Rows(i).Cells(4).Value))
                        '    cmd200.Parameters.AddWithValue("@codigo", Trim(DataGridView1.Rows(i).Cells(2).Value))
                        '    cmd200.Parameters.AddWithValue("@cantidad", DataGridView1.Rows(i).Cells(8).Value)
                        '    cmd200.Parameters.AddWithValue("@estado", 1)
                        '    cmd200.Parameters.AddWithValue("@numero_pedido", Trim(TextBox1.Text))
                        '    cmd200.Parameters.AddWithValue("@EMPRESA", Trim(Label26.Text))
                        '    cmd200.Parameters.AddWithValue("@ALMACEN", Trim(DataGridView1.Rows(i).Cells(14).Value))
                        '    cmd200.ExecuteNonQuery()
                        'Else

                        'End If


                    Next
                    Label30.Text = "1"
                    Button9.Enabled = False
                    MsgBox("SE REGISTRO LA NOTA DE PEDIDO SOLICITADA")

                    Dim respuesta As DialogResult
                    respuesta = MessageBox.Show("QUIERES IMPRIMIR LA NOTA DE PEDIDO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If (respuesta = Windows.Forms.DialogResult.Yes) Then
                        REPORTE_NP.TextBox1.Text = pedido
                        REPORTE_NP.TextBox2.Text = Label26.Text
                        REPORTE_NP.Show()
                        LIMPIAR()
                        LOAD12()
                    End If

                    Dim hj2 As New flog
                    Dim hj11 As New vlog
                    hj11.gmodulo = "PEDIDO-COMERCIAL"
                    hj11.gcuser = MDIParent1.Label3.Text
                    hj11.gaccion = TextBox19.Text
                    hj11.gpc = My.Computer.Name
                    hj11.gfecha = String.Format("{0:G}", DateTime.Now)
                    hj11.gdetalle = pedido
                    hj11.gccia = Label26.Text
                    hj2.insertar_log(hj11)
                    '                If TextBox19.Text = 1 Then
                    '                    Dim message As New MailMessage
                    '                    Dim smtp As New SmtpClient
                    '                    Dim fk As New fnopedido
                    '                    Dim jj As New vnapedido
                    '                    Dim corre As String
                    '                    jj.gvendedor = TextBox8.Text
                    '                    corre = fk.buscar_correo(jj)

                    '                    Dim Mailz As String = "" &
                    '      " <!DOCTYPE html>" &
                    '"<html xmlns='http://www.w3.org/1999/xhtml'>" &
                    '"<head>" &
                    '"    <title></title>" &
                    '"</head>" &
                    '"<body>
                    '                         <FONT COLOR='green'>INFORMACION DE LA NOTA DE PEDIDO</FONT><br/><br/>

                    '                    <FONT COLOR='blue'>* ESTADO : </FONT> CREADO <br/>
                    '                    <FONT COLOR='blue'>* NOTA DE PEDIDO : </FONT>" + Trim(TextBox1.Text) + " <br/>
                    '                    <FONT COLOR='blue'>* VENDEDOR : </FONT>" + Trim(TextBox9.Text) + " <br/>
                    '                    <FONT COLOR='blue'>* EMPRESA : </FONT>" + Trim(Label27.Text) + " <br/>
                    '                    <FONT COLOR='blue'>* CLIENTE : </FONT>" + Trim(TextBox3.Text) + "<br/>
                    '                    <FONT COLOR='blue'>* FORMA DE PAGO : </FONT>" + Trim(TextBox20.Text) + "<br/>
                    '                    <FONT COLOR='blue'>* TOTAL VENTA : </FONT>" + Trim(TextBox17.Text) + "<br/>

                    '                    </body>

                    '                    </html>"

                    '                    With message
                    '                        .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
                    '                        .To.Add("fmestanza@viannysac.com,lgonzalez@viannysac.com," & corre)
                    '                        .IsBodyHtml = True
                    '                        .Body = Mailz
                    '                        .Subject = "Nota Pedido N°" & pedido
                    '                        .Priority = System.Net.Mail.MailPriority.Normal
                    '                    End With

                    '                    With smtp

                    '                        .EnableSsl = True
                    '                        .Port = "587"
                    '                        .Host = "smtppro.zoho.com"
                    '                        .Credentials = New Net.NetworkCredential("admin@viannysac.com", "20Via$&it2")
                    '                        .Send(message)
                    '                    End With

                    '                    Try
                    '                        MessageBox.Show("Su mensaje de correo ha sido enviado.", "Correo enviado",
                    '                                 MessageBoxButtons.OK)
                    '                    Catch ex As Exception

                    '                    End Try
                    '                Else
                    '                    If TextBox19.Text = 2 Then
                    '                        ENVIAR_CORREO2()
                    '                    Else
                    '                        ENVIAR_CORREO3()
                    '                    End If

                    '                End If



                    LIMPIAR()
                    LOAD12()
                    Label24.Visible = False
                End If
            Else
                MsgBox("EL MONTO TOTAL DEL DETALLE NO COINCIDE CON EL TOTAL GENERAL")
            End If

        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        e.Handled = Not IsNumeric(e.KeyChar) And Not Char.IsControl(e.KeyChar)
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged


    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If Label19.Text = "APROBADO" Then
            MsgBox("NECESITA QUE LO DESAPRUEBEN PARA MODIFICARLO")
        Else
            Button9.Enabled = True
            Button10.Enabled = False
            TextBox2.ReadOnly = False
            TextBox13.ReadOnly = False
            ComboBox2.Enabled = True
            TextBox7.ReadOnly = False
            Button3.Enabled = True
            TextBox12.ReadOnly = False
            TextBox14.ReadOnly = False
            DataGridView1.Enabled = True
            CheckBox1.Enabled = True
            CheckBox2.Enabled = True
            CheckBox3.Enabled = True
            TextBox10.ReadOnly = False
            TextBox20.Enabled = True
            RadioButton1.Enabled = True
            RadioButton2.Enabled = True
            TextBox19.Text = 2
            TextBox10.Enabled = True
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If Label19.Text = "APROBADO" Then
            MsgBox("NECESITA QUE LO DESAPRUEBEN PARA ANULARLO")
        Else
            Dim FF As New fnopedido
        Dim JJ As New vnapedido


            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("QUIERES ANULAR NOTA DE PEDIDO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                JJ.gnumero_pedido = TextBox1.Text
                JJ.gempresa = Label26.Text
                FF.ANULAR_pedido(JJ)
                TextBox19.Text = 3
                MsgBox("SE REALIZO LO SOLICITADO")
                Dim hj2 As New flog
                Dim hj1 As New vlog
                hj1.gmodulo = "PEDIDO-COMERCIAL"
                hj1.gcuser = MDIParent1.Label3.Text
                hj1.gaccion = 3
                hj1.gpc = My.Computer.Name
                hj1.gfecha = DateTimePicker1.Value
                hj1.gdetalle = TextBox1.Text
                hj1.gccia = Label26.Text
                hj2.insertar_log(hj1)
                LIMPIAR()
                LOAD12()
            End If
        End If
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        REPORTE_NP.TextBox1.Text = TextBox1.Text
        REPORTE_NP.TextBox2.Text = Label26.Text
        REPORTE_NP.Show()
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs)

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        log.Label1.Text = Trim(TextBox1.Text)
        log.Label2.Text = "PEDIDO-COMERCIAL"
        log.Label3.Text = Label26.Text
        log.Show()
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        If Trim(TextBox2.Text) <> "" Then

            Productos_Vianny.TextBox2.Text = TextBox18.Text
            Productos_Vianny.TextBox5.Text = TextBox2.Text
            Productos_Vianny.TextBox6.Text = TextBox3.Text
            Productos_Vianny.TextBox7.Text = TextBox8.Text
            Productos_Vianny.TextBox8.Text = TextBox9.Text
            Productos_Vianny.Label4.Text = Label25.Text
            Productos_Vianny.Label5.Text = Label26.Text

            Productos_Vianny.ShowDialog()
        Else
            MsgBox("FALTA INGRESAR EL CLIENTE")
        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            If e.ColumnIndex = 5 Then
                DataGridView7.Rows.Clear()
                DataGridView1.Rows(e.RowIndex).Cells(5).Value = ""
                Dim num1, num2 As Integer
                num1 = e.RowIndex.ToString
                Orden_Packing.Label3.Text = "10"
                Orden_Packing.Label9.Text = num1
                Orden_Packing.TextBox6.Text = (TextBox1.Text)
                Orden_Packing.Label10.Text = DataGridView1.Rows(num1).Cells(1).Value
                Orden_Packing.Label4.Text = DataGridView1.Rows(num1).Cells(4).Value
                Orden_Packing.TextBox1.Text = DataGridView1.Rows(num1).Cells(4).Value
                Orden_Packing.TextBox4.Text = DataGridView1.Rows(num1).Cells(8).Value
                Orden_Packing.Show()
            End If
        End If
    End Sub

    Private Sub PictureBox4_KeyDown(sender As Object, e As KeyEventArgs)

    End Sub


    Private Sub TextBox1_DoubleClick(sender As Object, e As EventArgs) Handles TextBox1.DoubleClick
        Label30.Text = "1"
        LIMPIAR()
        LOAD12()
    End Sub

    Private Sub CheckBox4_Click(sender As Object, e As EventArgs) Handles CheckBox4.Click
        Dim cant10, sum10, cant9, sum9, cant8, sum8 As Double
        If CheckBox4.Checked = False Then
            CheckBox3.Checked = False
            CheckBox3.Enabled = False
            Dim i8 As Integer
            i8 = DataGridView1.RowCount
            For i = 0 To i8 - 1

                DataGridView1.Rows(i).Cells(12).Value = DataGridView1.Rows(i).Cells(8).Value * DataGridView1.Rows(i).Cells(9).Value
                DataGridView1.Rows(i).Cells(10).Value = DataGridView1.Rows(i).Cells(8).Value * DataGridView1.Rows(i).Cells(9).Value
                Me.DataGridView1.Columns(10).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(11).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                DataGridView1.Rows(i).Cells(11).Value = "0.00"
            Next

            For a9 = 0 To i8 - 1
                cant10 = Val(DataGridView1.Rows(a9).Cells(10).Value)
                sum10 = cant10 + Val(sum10)
                cant9 = Val(DataGridView1.Rows(a9).Cells(11).Value)
                sum9 = cant9 + Val(sum9)
                cant8 = Val(DataGridView1.Rows(a9).Cells(12).Value)
                sum8 = cant8 + Val(sum8)

            Next
            TextBox15.Text = sum8.ToString("N2")
            TextBox16.Text = sum9.ToString("N2")

            TextBox17.Text = sum10.ToString("N2")
        Else
            If CheckBox4.Checked = True Then
                CheckBox3.Checked = False
                CheckBox3.Enabled = True
                Dim i8 As Integer
                i8 = DataGridView1.RowCount
                For i = 0 To i8 - 1

                    DataGridView1.Rows(i).Cells(10).Value = DataGridView1.Rows(i).Cells(8).Value * DataGridView1.Rows(i).Cells(9).Value
                    DataGridView1.Rows(i).Cells(11).Value = DataGridView1.Rows(i).Cells(10).Value * 0.18
                    DataGridView1.Rows(i).Cells(12).Value = DataGridView1.Rows(i).Cells(10).Value + DataGridView1.Rows(i).Cells(11).Value
                    Me.DataGridView1.Columns(10).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(11).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                Next

                For a9 = 0 To i8 - 1
                    cant10 = Val(DataGridView1.Rows(a9).Cells(10).Value)
                    sum10 = cant10 + Val(sum10)
                    cant9 = Val(DataGridView1.Rows(a9).Cells(11).Value)
                    sum9 = cant9 + Val(sum9)
                    cant8 = Val(DataGridView1.Rows(a9).Cells(12).Value)
                    sum8 = cant8 + Val(sum8)

                Next

                TextBox15.Text = sum10.ToString("N2")
                TextBox16.Text = sum9.ToString("N2")

                TextBox17.Text = sum8.ToString("N2")


            End If

        End If
    End Sub

    Private Sub ComboBox1_Click(sender As Object, e As EventArgs) Handles ComboBox1.Click

    End Sub
End Class